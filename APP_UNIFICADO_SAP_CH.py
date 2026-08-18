# -*- coding: utf-8 -*-
"""
Automação Integrada — SAP & Cargo Heroes
Criação massiva de Requisições de Compra (ME51N) e atualização no Cargo Heroes.

Refatoração Profissional UI (CustomTkinter) + Correções de Segurança/Threading
"""

import pandas as pd
import win32com.client
import sys
import gspread
from datetime import datetime, timedelta
import subprocess
import time
import re
import os
import configparser
import pywintypes
import pythoncom
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import threading
from PIL import Image, ImageTk
import io
import base64
import logging
from logging.handlers import RotatingFileHandler
import ssl
import certifi
from typing import Optional, Any

# --- Imports Selenium ---
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    WebDriverException,
    NoSuchWindowException,
    TimeoutException,
)

# --- Imports Keyring (armazenamento seguro de senhas) ---
try:
    import keyring
    KEYRING_DISPONIVEL = True
except ImportError:
    KEYRING_DISPONIVEL = False

# --- Retry para chamadas de API ---
try:
    from tenacity import retry, wait_exponential, stop_after_attempt, retry_if_exception_type
    TENACITY_DISPONIVEL = True
except ImportError:
    TENACITY_DISPONIVEL = False

# Constantes de Segurança
KEYRING_SERVICE_SAP = "sap_automation_req_massivo"
KEYRING_SERVICE_CH = "cargo_heroes_automation"

# Configuração SSL segura
ssl._create_default_https_context = lambda: ssl.create_default_context(cafile=certifi.where())

# Configuração CustomTkinter
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


class LogRedirector:
    def __init__(self, text_widget: ctk.CTkTextbox) -> None:
        self.text_widget = text_widget
        self.tag_map: dict[str, str] = {
            "<<RESET>>": "RESET",
            "<<VERDE>>": "VERDE",
            "<<AMARELO>>": "AMARELO",
            "<<VERMELHO>>": "VERMELHO",
            "<<AZUL>>": "AZUL",
            "<<CIANO>>": "CIANO",
        }
        self.default_tag = "RESET"

    def write(self, string: str) -> None:
        try:
            self.text_widget.configure(state="normal")
            segments = re.split(
                f"({'|'.join(re.escape(k) for k in self.tag_map.keys())})", string
            )
            current_tag = self.default_tag
            for segment in segments:
                if segment in self.tag_map:
                    current_tag = self.tag_map[segment]
                elif segment:
                    self.text_widget.insert(tk.END, segment, tags=(current_tag,))
            self.text_widget.see(tk.END)
            self.text_widget.configure(state="disabled")
        except Exception as e:
            logging.debug(f"LogRedirector falhou: {e}")

    def flush(self) -> None:
        pass


class SAPAutomationGUI:
    DEPOSITO_MAPPING: dict[str, str] = {
        'BR0G': 'AE01', 'BR0Q': 'AE01', 'BR0D': 'AE01', 'BR0H': 'AE01', 'BR0O': 'AE01',
        'BR0P': 'AE01', 'BR0E': 'AE01', 'BR0R': 'AE01', 'BR0S': 'AE01', 'BR0Y': 'AE01',
        'BR0Z': 'AE01', 'BR1A': 'AE01', 'BR1C': 'AE01', 'BR1D': 'AE01', 'BR1G': 'AE01',
        'BR1I': 'AE01', 'BR1J': 'AE01', 'BR1K': 'AE01', 'BR1L': 'AE01', 'BR1T': 'AE01',
        'BR2A': 'AE01', 'BR2B': 'AE01', 'BR2C': 'AE01', 'BR2D': 'AE01', 'BR2E': 'AE01',
        'BR2Q': 'AE01', 'BR2U': 'AE01', 'BR2V': 'AE01', 'BR3A': 'AE01', 'BR3E': 'AE01',
        'BR3F': 'AE01', 'BR3K': 'AE01', 'BR3N': 'AE01', 'BRDN': 'AE01', 'BR8A': 'AE13',
        'BR2I': 'AE01', 'BR0I': 'AE13', 'BR0U': 'AE01', 'BR0K': 'AE13', 'BR0X': 'AE13',
        'BR0J': 'AE01', 'BR1E': 'AE01', 'BR1F': 'AE01', 'BR0V': 'AE01', 'BR8E': 'AE13',
        'BR1B': 'AE01', 'BR0F': 'AE01', 'BR8I': 'AE01', 'BRIJ': 'AE01', 'BR8G': 'AE01',
    }

    def __init__(self, root: ctk.CTk) -> None:
        self.root = root
        self.root.title("Automação Integrada - SAP & Cargo Heroes")
        self.root.geometry("1000x700")
        self.root.minsize(900, 600)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        self._session_lock = threading.Lock()
        self._session: Optional[Any] = None
        self.running: bool = False

        self.config = configparser.ConfigParser()
        self.data_path: str = self.get_data_path()
        self.resource_path: str = self.get_resource_path()
        
        self.config_path: str = os.path.join(self.data_path, 'config.ini')
        self.logs_print_path: str = os.path.join(self.data_path, 'logs', 'prints')

        if not os.path.exists(self.logs_print_path):
            os.makedirs(self.logs_print_path, exist_ok=True)

        self._setup_file_logger()

        try:
            self.config.read(self.config_path, encoding='utf-8')
        except Exception:
            self.create_default_config()

        self.load_icon()

        icon_start_b64 = self.fix_base64_padding("iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsTAAALEwEAmpwYAAABWWlUWHRYTUw6Y29tLmFkb2JlLnhtcAAAAAAAPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iWE1QIENvcmUgNS40LjAiPgogICA8cmRmOlJERiB4bWxuczpyZGY9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkvMDIvMjItcmRmLXN5bnRheC1ucyMiPgogICAgICA8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIgogICAgICAgICAgICB4bWxuczp0aWZmPSJodHRwOi8vbnMuYWRvYmUuY29tL3RpZmYvMS4wLyI+CiAgICAgICAgIDx0aWZmOk9yaWVudGF0aW9uPjE8L3RpZmY6T3JpZW50YXRpb24+CiAgICAgIDwvcmRmOkRlc2NyaXB0aW9uPgogICA8L3JkZjpSREY+CjwveDp4bXBtZXRhPgpMTE82AAAByklEQVQ4EaVTTUhUURQ+V3d1JzUzDBf9B0sLw0lDItocoQd9QUG3LhpEL7pw7aJdKyIi6FYQ2oZ1UXQRhFAb60PQg1ZCi2kStZpoGjcz9768eW/G6Mwbw72X+37n3HPvPQCB8Q+QnL8xQCeAF/D/o/kOWG01P4BHgCfz30KMR8s+AaYAPND9GkAy8BpwGSgLHsA74JESi2kUeC2APdC6BN8B+u0mYwTo/2QjE58D2pXfAYwA215W25gBFoD1/R9IZzGgYv03gC2gC1gDvgIHADgACsAecD6hL8eA/4A/AT8BfwL6n2WnAY+A+cBj4DHvEY+BTYBT4B/gY+An4G9gOvgZfAReA/4FvgK/AqcB14DwwB+H/t9w9OAD+An/d8D/4L/EXkX2gN8A1XkLwNngGfA58Ar4DNwFvgE/A28Bv4Cfi/sXv0C/A78DPgIeBR4DvgE+A54AvAE+A9YAhYABYAzYC/gH2APWApWAcvAdqAfeAesAyuBvA74Anjkl+sN4EPgG+A3wBvgS5Y4ADwGjgIfAm8BD4GngCPANeA0cAW4BnwG3AcuArdBB/gMvAasA/8CPga+Ad8AvwB+A/YD/gD+AfwB+D/AP8C/gb8F/gX+A/wH/A78CvwNfA1sApsBGgHNgNngCXgGfAYeBIMAP+A18B6sAmsA58AW8AC8A2YAxaBLWAb2A3uAx8AZ4BnwEvAm8BTYB/4BOwBDoCTQBDwBrgD3AE+Ax4AZoCNb+V2d+AocB14B/gJ+BG4C1wBfgE+A74DPgX+D/gb+B/wL/A/4DvgF+BGYAhZ/f+AasAm8BWwA8+B14BTwGnAEOAX8B7wG3gV8H4L/f7bB/yXgG/A18CXwDvgQ2AH+AF8CJ4DngHPAK8Ar4PXgGfA78DPwN/Ad8A3wL+Bn5B/gP8F/rX7F/gV+B/wI/AJ8K/AnwA/8AMxS88wI6H8lAAAAAElFTSuQmCC")
        icon_stop_b64 = self.fix_base64_padding("iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAA7SURBVDhPY/wPBAxUACZA1gBTgA2K/1EGBgaG/2A8HIBoMRhIMeAgYADEDwFGBgYGDjA50uAzAAgwAK0/AwMT5urRAAAAAElFTSuQmCC")

        try:
            self.icons = {
                "start": ctk.CTkImage(light_image=Image.open(io.BytesIO(base64.b64decode(icon_start_b64))), size=(16, 16)),
                "stop": ctk.CTkImage(light_image=Image.open(io.BytesIO(base64.b64decode(icon_stop_b64))), size=(16, 16))
            }
        except Exception:
            self.icons = {"start": None, "stop": None}

        self.create_widgets()
        self._configure_log_tags()

    @property
    def session(self) -> Optional[Any]:
        with self._session_lock:
            return self._session

    @session.setter
    def session(self, value: Optional[Any]) -> None:
        with self._session_lock:
            self._session = value

    def fix_base64_padding(self, b64_string: str) -> str:
        return b64_string + "=" * ((4 - len(b64_string) % 4) % 4)

    def _setup_file_logger(self) -> None:
        log_file_path = os.path.join(self.data_path, 'app_log.txt')
        file_handler = RotatingFileHandler(
            log_file_path, maxBytes=5 * 1024 * 1024, backupCount=3, encoding='utf-8'
        )
        file_handler.setFormatter(logging.Formatter(
            '[%(asctime)s] [%(levelname)s] %(message)s',
            datefmt='%d/%m/%Y %H:%M:%S'
        ))
        file_handler.setLevel(logging.INFO)

        logger = logging.getLogger()
        logger.setLevel(logging.DEBUG)
        for handler in logger.handlers[:]:
            logger.removeHandler(handler)
        logger.addHandler(file_handler)

    def get_data_path(self) -> str:
        if getattr(sys, 'frozen', False):
            return os.path.dirname(sys.executable)
        return os.path.dirname(os.path.abspath(__file__))

    def get_resource_path(self) -> str:
        if getattr(sys, 'frozen', False):
            return sys._MEIPASS
        return os.path.dirname(os.path.abspath(__file__))

    def is_executable(self) -> bool:
        return getattr(sys, 'frozen', False)

    def load_icon(self) -> None:
        try:
            icon_path = os.path.join(self.resource_path, 'icone.ico')
            if not os.path.exists(icon_path):
                icon_path = os.path.join(self.data_path, 'icone.ico')
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception: pass

    def create_default_config(self) -> None:
        self.config['SAP'] = {'caminho_logon': '', 'sistema': '', 'usuario': ''}
        self.config['GOOGLE'] = {'credenciais': 'credentials.json', 'planilha': '', 'aba': ''}
        self.config['CARGO_HEROES'] = {'email': ''}
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                self.config.write(f)
        except OSError: pass

    def _obter_senha(self, servico: str, chave: str, fallback_section: str = '', fallback_key: str = 'senha') -> str:
        if KEYRING_DISPONIVEL:
            try:
                senha = keyring.get_password(servico, chave)
                if senha: return senha
            except Exception: pass
        if fallback_section:
            return self.config.get(fallback_section, fallback_key, fallback='')
        return ''

    def _salvar_senha(self, servico: str, chave: str, senha: str) -> bool:
        if KEYRING_DISPONIVEL and senha:
            try:
                keyring.set_password(servico, chave, senha)
                return True
            except Exception: pass
        return False

    def create_widgets(self) -> None:
        self.header_label = ctk.CTkLabel(self.root, text="Automação Integrada - SAP & Cargo Heroes", font=ctk.CTkFont(size=24, weight="bold"))
        self.header_label.pack(pady=(20, 0))

        self.tabview = ctk.CTkTabview(self.root, width=950, height=550)
        self.tabview.pack(padx=20, pady=10, fill="both", expand=True)
        self.tabview.add("Automação")
        self.tabview.add("Configurações")

        self.setup_main_tab(self.tabview.tab("Automação"))
        self.setup_config_tab(self.tabview.tab("Configurações"))

        self.status_frame = ctk.CTkFrame(self.root, height=35, corner_radius=0, fg_color="transparent")
        self.status_frame.pack(side="bottom", fill="x", padx=20, pady=(0, 10))

        self.sap_status_var = tk.StringVar(value="SAP: Desconectado")
        self.sap_status_label = ctk.CTkLabel(self.status_frame, textvariable=self.sap_status_var, text_color="#ef5350", font=ctk.CTkFont(weight="bold"))
        self.sap_status_label.pack(side="left")

        self.status_var = tk.StringVar(value="Pronto")
        self.status_label = ctk.CTkLabel(self.status_frame, textvariable=self.status_var, text_color="gray")
        self.status_label.pack(side="right")

    def setup_main_tab(self, parent: ctk.CTkFrame) -> None:
        control_frame = ctk.CTkFrame(parent, fg_color="transparent")
        control_frame.pack(pady=(10, 15))

        self.start_button = ctk.CTkButton(
            control_frame, text="Iniciar SAP", image=self.icons["start"],
            fg_color="#4CAF50", hover_color="#45a049",
            command=self.start_automation, width=180, height=45, font=ctk.CTkFont(size=14, weight="bold")
        )
        self.start_button.pack(side="left", padx=15)

        self.ch_button = ctk.CTkButton(
            control_frame, text="Atualizar CH", image=self.icons["start"],
            fg_color="#2196F3", hover_color="#1E88E5",
            command=self.start_ch_automation, width=180, height=45, font=ctk.CTkFont(size=14, weight="bold")
        )
        self.ch_button.pack(side="left", padx=15)

        self.stop_button = ctk.CTkButton(
            control_frame, text="Parar Automação", image=self.icons["stop"],
            fg_color="#f44336", hover_color="#d32f2f", state="disabled",
            command=self.stop_automation, width=180, height=45, font=ctk.CTkFont(size=14, weight="bold")
        )
        self.stop_button.pack(side="left", padx=15)

        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ctk.CTkProgressBar(parent, variable=self.progress_var)
        self.progress_bar.pack(fill="x", padx=15, pady=(0, 10))

        log_label = ctk.CTkLabel(parent, text="Log de Execução", font=ctk.CTkFont(weight="bold"))
        log_label.pack(anchor="w", padx=15)

        self.log_area = ctk.CTkTextbox(parent, font=ctk.CTkFont(family="Consolas", size=12), fg_color="#1e1e1e", text_color="#d4d4d4")
        self.log_area.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        self.log_area.configure(state="disabled")

    def _configure_log_tags(self):
        self.log_area.tag_config("RESET", foreground="#D0D0D0")
        self.log_area.tag_config("VERDE", foreground="#66bb6a")
        self.log_area.tag_config("AMARELO", foreground="#ffa726")
        self.log_area.tag_config("VERMELHO", foreground="#ef5350")
        self.log_area.tag_config("AZUL", foreground="#42a5f5")
        self.log_area.tag_config("CIANO", foreground="#26c6da")

    def setup_config_tab(self, parent: ctk.CTkFrame) -> None:
        scroll_frame = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll_frame.pack(fill="both", expand=True)

        # SAP
        sap_frame = ctk.CTkFrame(scroll_frame)
        sap_frame.pack(fill="x", pady=10, padx=10)
        ctk.CTkLabel(sap_frame, text="Credenciais SAP", font=ctk.CTkFont(weight="bold", size=16)).pack(anchor="w", padx=10, pady=(10, 5))

        self.sap_path_var = tk.StringVar(value=self.config.get('SAP', 'caminho_logon', fallback=''))
        self.create_config_row(sap_frame, "Caminho Logon.exe:", self.sap_path_var, show_browse=True)

        self.sap_system_var = tk.StringVar(value=self.config.get('SAP', 'sistema', fallback=''))
        self.create_config_row(sap_frame, "Sistema / Conexão:", self.sap_system_var)

        self.sap_user_var = tk.StringVar(value=self.config.get('SAP', 'usuario', fallback=''))
        self.create_config_row(sap_frame, "Usuário:", self.sap_user_var)

        self.sap_password_var = tk.StringVar(value=self._obter_senha(KEYRING_SERVICE_SAP, 'senha', 'SAP', 'senha'))
        self.create_config_row(sap_frame, "Senha:", self.sap_password_var, is_password=True)

        # GOOGLE
        google_frame = ctk.CTkFrame(scroll_frame)
        google_frame.pack(fill="x", pady=10, padx=10)
        ctk.CTkLabel(google_frame, text="Configurações Google Sheets", font=ctk.CTkFont(weight="bold", size=16)).pack(anchor="w", padx=10, pady=(10, 5))

        self.google_creds_var = tk.StringVar(value=self.config.get('GOOGLE', 'credenciais', fallback=''))
        self.create_config_row(google_frame, "JSON Credenciais:", self.google_creds_var, show_browse=True)

        self.google_sheet_var = tk.StringVar(value=self.config.get('GOOGLE', 'planilha', fallback=''))
        self.create_config_row(google_frame, "Nome da Planilha:", self.google_sheet_var)

        self.google_tab_var = tk.StringVar(value=self.config.get('GOOGLE', 'aba', fallback=''))
        self.create_config_row(google_frame, "Nome da Aba Principal:", self.google_tab_var)

        # CARGO HEROES
        ch_frame = ctk.CTkFrame(scroll_frame)
        ch_frame.pack(fill="x", pady=10, padx=10)
        ctk.CTkLabel(ch_frame, text="Credenciais Cargo Heroes", font=ctk.CTkFont(weight="bold", size=16)).pack(anchor="w", padx=10, pady=(10, 5))

        self.ch_email_var = tk.StringVar(value=self.config.get('CARGO_HEROES', 'email', fallback=''))
        self.create_config_row(ch_frame, "Email:", self.ch_email_var)

        self.ch_pass_var = tk.StringVar(value=self._obter_senha(KEYRING_SERVICE_CH, 'senha', 'CARGO_HEROES', 'senha'))
        self.create_config_row(ch_frame, "Senha:", self.ch_pass_var, is_password=True)

        # Botão Salvar
        ctk.CTkButton(scroll_frame, text="Salvar Configurações", command=self.save_config, width=250, height=40, font=ctk.CTkFont(weight="bold")).pack(pady=20)

    def create_config_row(self, parent, label_text, variable, is_password=False, show_browse=False):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", padx=10, pady=5)
        
        lbl = ctk.CTkLabel(frame, text=label_text, width=150, anchor="e", font=ctk.CTkFont(size=13))
        lbl.pack(side="left", padx=(0, 10))
        
        show_char = "*" if is_password else ""
        entry = ctk.CTkEntry(frame, textvariable=variable, show=show_char, height=35)
        entry.pack(side="left", fill="x", expand=True)
        
        if show_browse:
            btn = ctk.CTkButton(frame, text="Selecionar", width=80, height=35, command=lambda: self.browse_file(variable))
            btn.pack(side="left", padx=(10, 0))

    def browse_file(self, variable: tk.StringVar) -> None:
        filename = filedialog.askopenfilename()
        if filename: variable.set(filename)

    def save_config(self) -> None:
        if 'SAP' not in self.config: self.config['SAP'] = {}
        self.config['SAP']['caminho_logon'] = self.sap_path_var.get()
        self.config['SAP']['sistema'] = self.sap_system_var.get()
        self.config['SAP']['usuario'] = self.sap_user_var.get()

        if 'GOOGLE' not in self.config: self.config['GOOGLE'] = {}
        self.config['GOOGLE']['credenciais'] = self.google_creds_var.get()
        self.config['GOOGLE']['planilha'] = self.google_sheet_var.get()
        self.config['GOOGLE']['aba'] = self.google_tab_var.get()

        if 'CARGO_HEROES' not in self.config: self.config['CARGO_HEROES'] = {}
        self.config['CARGO_HEROES']['email'] = self.ch_email_var.get()

        try:
            with open(self.config_path, 'w', encoding='utf-8') as configfile:
                self.config.write(configfile)

            senha_sap = self.sap_password_var.get()
            senha_ch = self.ch_pass_var.get()
            sap_salvo = self._salvar_senha(KEYRING_SERVICE_SAP, 'senha', senha_sap)
            ch_salvo = self._salvar_senha(KEYRING_SERVICE_CH, 'senha', senha_ch)

            if KEYRING_DISPONIVEL and sap_salvo and ch_salvo:
                if self.config.has_option('SAP', 'senha'): self.config.remove_option('SAP', 'senha')
                if self.config.has_option('CARGO_HEROES', 'senha'): self.config.remove_option('CARGO_HEROES', 'senha')
                with open(self.config_path, 'w', encoding='utf-8') as configfile:
                    self.config.write(configfile)
                messagebox.showinfo("Sucesso", "Configurações salvas de forma segura no Windows Credential Manager.")
            else:
                self.config['SAP']['senha'] = senha_sap
                self.config['CARGO_HEROES']['senha'] = senha_ch
                with open(self.config_path, 'w', encoding='utf-8') as configfile:
                    self.config.write(configfile)
                messagebox.showwarning("Aviso", "Keyring ausente. Salvo no config.ini.")
        except OSError as e:
            messagebox.showerror("Erro", f"Erro: {e}")

    def on_closing(self) -> None:
        if self.running:
            if not messagebox.askyesno("Sair", "A automação está em andamento. Deseja sair?"):
                return
            self.stop_automation()
        self.restore_stdout()
        self.root.destroy()

    def toggle_buttons(self, habilitado: bool) -> None:
        estado_principal = "normal" if habilitado else "disabled"
        estado_parar = "disabled" if habilitado else "normal"
        self.start_button.configure(state=estado_principal)
        self.ch_button.configure(state=estado_principal)
        self.stop_button.configure(state=estado_parar)

    def atualizar_status_sap(self, conectado: bool = False, mensagem: Optional[str] = None) -> None:
        if conectado:
            self.sap_status_var.set("SAP: Conectado")
            self.sap_status_label.configure(text_color="#66bb6a")
        else:
            self.sap_status_var.set("SAP: Desconectado")
            self.sap_status_label.configure(text_color="#ef5350")
        if mensagem:
            self.status_var.set(mensagem)

    def validate_config(self) -> bool:
        return all([self.sap_path_var.get(), self.sap_system_var.get(), self.google_creds_var.get(), self.google_sheet_var.get()])

    def update_progress(self, value: float) -> None:
        self.root.after(0, lambda: self.progress_var.set(value / 100.0))

    def setup_log_redirector(self) -> None:
        sys.stdout = LogRedirector(self.log_area)

    def restore_stdout(self) -> None:
        sys.stdout = sys.__stdout__

    def _get_timestamp(self) -> str:
        return datetime.now().strftime('%H:%M:%S')

    def print_header(self, texto: str) -> None:
        log_text = f"\n[{self._get_timestamp()}] 🚀 {texto.upper()}\n"
        print(f"<<AZUL>>{log_text}<<RESET>>")
        logging.info(texto)

    def print_sucesso(self, texto: str) -> None:
        log_text = f"[{self._get_timestamp()}] ✔  SUCESSO:  {texto}\n"
        print(f"<<VERDE>>{log_text}<<RESET>>")
        logging.info(f"[SUCESSO] {texto}")

    def print_info(self, texto: str) -> None:
        log_text = f"[{self._get_timestamp()}] ℹ  INFO:     {texto}\n"
        print(f"<<CIANO>>{log_text}<<RESET>>")
        logging.info(texto)

    def print_aviso(self, texto: str) -> None:
        log_text = f"[{self._get_timestamp()}] ⚠  AVISO:    {texto}\n"
        print(f"<<AMARELO>>{log_text}<<RESET>>")
        logging.warning(texto)

    def print_erro(self, texto: str) -> None:
        log_text = f"[{self._get_timestamp()}] ✖  ERRO:     {texto}\n"
        print(f"<<VERMELHO>>{log_text}<<RESET>>")
        logging.error(texto)

    # =========================================================================
    #  AUTOMAÇÃO SAP
    # =========================================================================
    def start_automation(self) -> None:
        if self.running: return
        if not self.validate_config():
            messagebox.showerror("Erro", "Preencha todas as configurações.")
            return
        self.running = True
        self.toggle_buttons(False)
        self.status_var.set("Executando SAP...")
        self.log_area.configure(state="normal")
        self.log_area.delete("1.0", "end")
        self.log_area.configure(state="disabled")
        self.setup_log_redirector()
        threading.Thread(target=self.run_sap_automation, daemon=True).start()

    def stop_automation(self) -> None:
        if not self.running: return
        self.running = False
        self.status_var.set("Parando...")
        self.print_aviso("Solicitação de parada. Aguarde o fim do ciclo atual...")

    def run_sap_automation(self) -> None:
        try:
            pythoncom.CoInitialize()
        except pywintypes.com_error as e:
            self.print_erro(f"Erro COM: {e}")
            return
        try:
            self.print_header("Iniciando SAP Robot")
            if not self.is_session_valid():
                self.root.after(0, lambda: self.atualizar_status_sap(False, "Conectando..."))
                self.session = self.sap_login_handler()
            if not self.session:
                self.print_erro("Falha na conexão SAP.")
                self.root.after(0, lambda: self.atualizar_status_sap(False, "Falha conexão"))
                return
            self.print_sucesso("Sessão SAP OK!")
            self.root.after(0, lambda: self.atualizar_status_sap(True, "Conectado ao SAP"))
            
            try:
                self.print_header("CONECTANDO PLANILHA")
                gc = gspread.service_account(filename=self.google_creds_var.get())
                sh = gc.open(self.google_sheet_var.get())
                ws = sh.worksheet(self.google_tab_var.get())
                self.print_sucesso("Conexão Google Sheets estabelecida.")
                
                df = pd.DataFrame(ws.get_all_records())
                if df.empty:
                    self.print_aviso("Planilha vazia.")
                    return
                df['linha_planilha'] = df.index + 2
                df_pendente = df[df['Status'] == ''].copy()
                if df_pendente.empty:
                    self.print_aviso("Nenhum registro pendente para criar.")
                else:
                    self.processar_lotes(df_pendente, ws, ws.row_values(1).index("Status")+1, ws.row_values(1).index("REQUISIÇÃO")+1)
            except Exception as e:
                self.print_erro(f"Erro de API Planilha: {e}")
        finally:
            self.print_header("FIM DO CICLO")
            self.root.after(0, self.finalize_automation)
            try: pythoncom.CoUninitialize()
            except: pass

    def finalize_automation(self) -> None:
        self.running = False
        self.toggle_buttons(True)
        self.status_var.set("Pronto")
        self.update_progress(0)
        self.restore_stdout()

    def aguardar_sap(self, timeout: int = 30) -> bool:
        if not self.session: return False
        t0 = time.time()
        while self.running:
            try:
                if not self.session.busy: return True
            except: return False
            if time.time() - t0 > timeout: return False
            time.sleep(0.2)
        return False

    def is_session_valid(self) -> bool:
        if not self.session: return False
        try:
            self.session.findById("wnd[0]")
            return True
        except: return False

    def sap_login_handler(self) -> Optional[Any]:
        try:
            sap_gui = win32com.client.GetObject("SAPGUI")
            app = sap_gui.GetScriptingEngine
            for i in range(app.Connections.Count):
                conn = app.Connections(i)
                for j in range(conn.Sessions.Count):
                    sess = conn.Sessions(j)
                    try:
                        sess.findById("wnd[0]")
                        return sess
                    except: pass
        except: pass
        return self.open_and_login_sap()

    def open_and_login_sap(self) -> Optional[Any]:
        try:
            subprocess.Popen(self.config['SAP']['caminho_logon'])
            time.sleep(5)
            sap_gui = win32com.client.GetObject("SAPGUI")
            app = sap_gui.GetScriptingEngine
            conn = app.OpenConnection(self.config['SAP']['sistema'], True)
            time.sleep(3)
            sess = conn.Children(0)
            t0 = time.time()
            while sess.busy:
                if time.time() - t0 > 30: return None
                time.sleep(0.5)
            try:
                sess.findById("wnd[0]/usr/txtRSYST-BNAME").text = self.config['SAP']['usuario']
                sess.findById("wnd[0]/usr/pwdRSYST-BCODE").text = self._obter_senha(KEYRING_SERVICE_SAP, 'senha', 'SAP', 'senha')
                sess.findById("wnd[0]").sendVKey(0)
                t0 = time.time()
                while sess.busy:
                    if time.time() - t0 > 30: return None
                    time.sleep(0.5)
                try: sess.findById("wnd[1]").sendVKey(0)
                except: pass
                txt = sess.findById("wnd[0]").text.lower()
                if "easy access" in txt or "menú" in txt: return sess
            except: return sess
        except: return None

    def _extrair_dados_item(self, item: pd.Series) -> dict[str, str]:
        return {
            'material_id': str(item.get('PN') or item.get('Material ID', '')),
            'origem': str(item.get('ORIGEM') or item.get('Origem Sigla', '')),
            'destino': str(item.get('DESTINO') or item.get('Destino Sigla', '')),
            'quantidade': str(item.get('QTD') or item.get('Quantidade', '1')).replace(',', '.'),
            'texto': str(item.get('TEXTO') or item.get('Logística', '')),
        }

    def _navegar_me51n(self) -> Optional[Any]:
        try:
            self.session.findById("wnd[0]").maximize()
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/NME51N"
            self.session.findById("wnd[0]").sendVKey(0)
            self.aguardar_sap()
            time.sleep(1)
            self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:3327/cmbMEREQ_TOPLINE-BSART").key = "ZRT"
            self.session.findById("wnd[0]").sendVKey(0)
            self.aguardar_sap()
            return self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell")
        except: return None

    def _preencher_grid_item(self, grid: Any, indice: int, dados: dict[str, str], data_hoje: str) -> None:
        grid.modifyCell(indice, "MATNR", dados['material_id'])
        grid.modifyCell(indice, "MENGE", dados['quantidade'])
        grid.modifyCell(indice, "RESWK", dados['origem'])
        grid.modifyCell(indice, "EEIND", data_hoje)
        grid.modifyCell(indice, "EPSTP", "U")
        grid.modifyCell(indice, "NAME1", dados['destino'])
        grid.modifyCell(indice, "EKGRP", "P04")
        grid.modifyCell(indice, "TXZ01", dados['texto'])

    def _sap_voltar_tela_inicial(self) -> None:
        try:
            if self.is_session_valid():
                self.session.findById("wnd[0]/tbar[0]/okcd").text = "/N"
                self.session.findById("wnd[0]").sendVKey(0)
        except: pass

    def _batch_update_planilha(self, worksheet: Any, updates: list[dict]) -> None:
        if not updates: return
        if TENACITY_DISPONIVEL:
            @retry(wait=wait_exponential(min=2, max=30), stop=stop_after_attempt(3))
            def _upd(): worksheet.batch_update(updates)
            try: _upd()
            except Exception as e: self.print_erro(f"Falha API: {e}")
        else:
            try: worksheet.batch_update(updates)
            except Exception as e: self.print_erro(f"Erro API: {e}")

    def processar_lotes(self, df: pd.DataFrame, ws: Any, col_st: int, col_req: int) -> None:
        hoje = datetime.now().strftime('%d.%m.%Y')
        origem_col = 'ORIGEM' if 'ORIGEM' in df.columns else 'Origem Sigla'
        dest_col = 'DESTINO' if 'DESTINO' in df.columns else 'Destino Sigla'
        lotes = []
        for _, g in df.groupby([origem_col, dest_col]):
            for i in range(0, len(g), 10):
                c = g.iloc[i:i+10].copy()
                c['grid_index'] = range(len(c))
                lotes.append(c)
        tot = len(lotes)
        for i, lote in enumerate(lotes):
            if not self.running: break
            if not self.is_session_valid(): self.session = self.sap_login_handler()
            if not self.session: break
            self.update_progress(((i+1)/tot)*100)
            self.print_header(f"Processando Lote {i+1}/{tot}")
            
            res = self.validar_lote_na_rc(lote, hoje)
            upds = []
            ok_lines = []
            for r in res:
                upds.append({'range': gspread.utils.rowcol_to_a1(r["linha_planilha"], col_st), 'values': [[str(r['status'])]]})
                upds.append({'range': gspread.utils.rowcol_to_a1(r["linha_planilha"], col_req), 'values': [[str(r['numero_rc'])]]})
                if r['status'] == 'OK': ok_lines.append(r['linha_planilha'])
            self._batch_update_planilha(ws, upds)
            
            if not ok_lines: continue
            lok = lote[lote['linha_planilha'].isin(ok_lines)].copy()
            lok['grid_index'] = range(len(lok))
            if not self.is_session_valid(): break
            
            rc, m = self.criar_rc_para_lote_ok(lok, hoje)
            fin_upds = []
            for lin in lok['linha_planilha']:
                fin_upds.append({'range': gspread.utils.rowcol_to_a1(lin, col_st), 'values': [[m]]})
                if rc: fin_upds.append({'range': gspread.utils.rowcol_to_a1(lin, col_req), 'values': [[rc]]})
            if fin_upds:
                self._batch_update_planilha(ws, fin_upds)
                if rc: self.print_sucesso(f"RC {rc} Criada com Sucesso")

    def validar_lote_na_rc(self, lote: pd.DataFrame, data: str) -> list[dict]:
        res = []
        try:
            grid = self._navegar_me51n()
            if not grid:
                for _, r in lote.iterrows(): res.append({'linha_planilha': r['linha_planilha'], 'status': 'Falha ME51N', 'numero_rc': 'ERRO'})
                return res
            for _, item in lote.iterrows():
                if not self.running: break
                idx = int(item['grid_index'])
                d = self._extrair_dados_item(item)
                st = "OK"
                try:
                    self._preencher_grid_item(grid, idx, d, data)
                    self.session.findById("wnd[0]").sendVKey(0)
                    self.aguardar_sap()
                    time.sleep(1)
                    try: self.session.findById("wnd[1]").sendVKey(0)
                    except: pass
                    sb = self.session.findById("wnd[0]/sbar")
                    if sb.messageType in ('E', 'A') or "não está atualizado" in sb.text: st = sb.text
                except Exception as e: st = f"Erro item: {e}"
                res.append({'linha_planilha': item['linha_planilha'], 'status': st, 'numero_rc': '' if st == 'OK' else 'ERRO'})
            return res
        finally: self._sap_voltar_tela_inicial()

    def criar_rc_para_lote_ok(self, lote: pd.DataFrame, data: str) -> tuple[Optional[str], str]:
        try:
            grid = self._navegar_me51n()
            if not grid: return None, "Falha ME51N"
            for i, item in lote.reset_index(drop=True).iterrows():
                if not self.running: return None, "Cancelado"
                self._preencher_grid_item(grid, i, self._extrair_dados_item(item), data)
            self.session.findById("wnd[0]").sendVKey(0)
            self.aguardar_sap()
            for i, item in lote.reset_index(drop=True).iterrows():
                if not self.running: return None, "Cancelado"
                dep = self.DEPOSITO_MAPPING.get(str(item.get('ORIGEM') or item.get('Origem Sigla')).strip().upper(), 'AE01')
                grid.setCurrentCell(i, "MATNR")
                self.aguardar_sap()
                self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT16").select()
                self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT16/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsTABREITER1/tabpTRANS").select()
                self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT16/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsTABREITER1/tabpTRANS/ssubSUBBILD1:SAPLXM02:0114/ctxtEBAN-ZZDEP_FORNEC").text = str(dep)
                if i < len(lote) - 1:
                    self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/btn%#AUTOTEXT002").press()
                    self.aguardar_sap()
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            self.aguardar_sap()
            try: self.session.findById("wnd[1]").sendVKey(0)
            except: pass
            m = self.session.findById("wnd[0]/sbar").text
            ma = re.search(r'(\d{10,})', m)
            if ma: return ma.group(0), m
            return None, m
        except Exception as e: return None, f"Erro criar RC: {e}"

    # =========================================================================
    #  AUTOMAÇÃO CARGO HEROES (SELENIUM RESTAURADO)
    # =========================================================================

    def ch_extrair_horarios(self, texto_logistica: str) -> tuple[Optional[str], Optional[str]]:
        texto = str(texto_logistica).strip()
        padrao_hora = r'\b(?:[01]?\d|2[0-3]):[0-5]\d\b'
        horarios = re.findall(padrao_hora, texto)
        if len(horarios) >= 2: return horarios[0], horarios[1]
        return None, None

    def ch_calcular_data_hora(self, horario_str: str) -> Optional[str]:
        try:
            agora = datetime.now()
            parts = horario_str.split(':')
            hora, minuto = int(parts[0]), int(parts[1])
            data_alvo = agora.replace(hour=hora, minute=minuto, second=0, microsecond=0)
            if data_alvo < agora - timedelta(hours=6):
                data_alvo += timedelta(days=1)
            return data_alvo.strftime("%d%m%Y%H%M")
        except: return None

    def ch_preencher_data_js(self, driver: webdriver.Chrome, wait: WebDriverWait, xpath_id: str, texto_num: str, descricao: str = "Data") -> bool:
        try:
            if not texto_num or len(texto_num) < 12: return False
            dia, mes, ano = texto_num[:2], texto_num[2:4], texto_num[4:8]
            hora, minuto = texto_num[8:10], texto_num[10:12]
            data_iso = f"{ano}-{mes}-{dia}T{hora}:{minuto}"
            el = wait.until(EC.presence_of_element_located((By.XPATH, xpath_id)))
            driver.execute_script("""
                arguments[0].value = arguments[1];
                arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
            """, el, data_iso)
            return True
        except: return False

    def ch_acao(self, driver: webdriver.Chrome, wait: WebDriverWait, xpath: str, acao: str = "clicar", texto: Optional[str] = None, desc: str = "") -> bool:
        if not self.running: return False
        try:
            el = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            if acao == "clicar":
                try: el.click()
                except: driver.execute_script("arguments[0].click();", el)
            elif acao == "escrever":
                el.clear()
                el.send_keys(texto)
            return True
        except: return False

    def ch_busca_material(self, driver: webdriver.Chrome, wait: WebDriverWait, material: str) -> bool:
        try:
            el = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@formcontrolname='equipmentCode'] | //input[contains(@data-placeholder, 'Materia')]")))
            el.click(); time.sleep(0.2)
            el.send_keys(Keys.CONTROL + "a"); time.sleep(0.1)
            el.send_keys(Keys.BACKSPACE); time.sleep(0.1)
            el.clear(); el.send_keys(material); time.sleep(1)
            self.ch_acao(driver, wait, "//button[contains(., 'Procurar')]", "clicar")
            time.sleep(2)
            return True
        except: return False

    def start_ch_automation(self) -> None:
        if self.running: return
        email = self.ch_email_var.get()
        senha = self.ch_pass_var.get()
        if not email or not senha:
            messagebox.showerror("Erro", "Configure Email e Senha do CH.")
            return
        self.running = True
        self.toggle_buttons(False)
        self.status_var.set("Executando Cargo Heroes...")
        self.log_area.configure(state="normal")
        self.log_area.delete("1.0", "end")
        self.log_area.configure(state="disabled")
        self.setup_log_redirector()
        threading.Thread(target=self.run_ch_automation, daemon=True).start()

    def run_ch_automation(self) -> None:
        driver = None
        try:
            self.print_header("Iniciando Cargo Heroes Updater")
            opts = Options()
            opts.add_argument("--start-maximized")
            driver = webdriver.Chrome(options=opts)
            wait = WebDriverWait(driver, 20)
            
            self.print_info("Conectando Planilha Google...")
            c = ServiceAccountCredentials.from_json_keyfile_name(self.google_creds_var.get(), ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"])
            gc = gspread.authorize(c)
            sh = gc.open(self.google_sheet_var.get())
            aba_n = sh.worksheet(self.google_tab_var.get())
            try: aba_m = sh.worksheet('MAPEAMENTO')
            except: aba_m = None
            
            driver.get("https://cargo-heroes.appslatam.com/#/login")
            if self.ch_realizar_login(driver, wait):
                if self.ch_navegar_detalhe(driver, wait):
                    if self.running: self.ch_processar_normal(driver, wait, aba_n)
                    if self.running and aba_m: self.ch_processar_mapeamento(driver, wait, aba_m)
                    self.print_sucesso("Processo CH finalizado!")
            else:
                self.print_erro("Falha no Login CH.")
        except Exception as e: self.print_erro(f"Erro Fatal CH: {e}")
        finally:
            if driver:
                try: driver.quit()
                except: pass
            self.root.after(0, self.finalize_automation)

    def ch_realizar_login(self, driver: webdriver.Chrome, wait: WebDriverWait) -> bool:
        self.print_info("🔑 Realizando Login...")
        try:
            email = self.ch_email_var.get()
            senha = self.ch_pass_var.get()
            try:
                iframe = wait.until(EC.presence_of_element_located((By.XPATH, "//iframe[contains(@src, 'gsi/button')]")))
                driver.switch_to.frame(iframe)
                driver.execute_script("arguments[0].click();", wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@role='button']"))))
                driver.switch_to.default_content()
            except: driver.switch_to.default_content()
            try:
                WebDriverWait(driver, 5).until(EC.number_of_windows_to_be(2))
                jp = driver.current_window_handle
                jv = [j for j in driver.window_handles if j != jp][0]
                driver.switch_to.window(jv)
                try:
                    c = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.ID, "i0116")))
                    c.send_keys(email); driver.find_element(By.ID, "idSIButton9").click()
                    s = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "i0118")))
                    s.send_keys(senha); time.sleep(0.5); driver.find_element(By.ID, "idSIButton9").click()
                    try: WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
                    except: pass
                except: pass
                driver.switch_to.window(jp)
            except: pass
            
            t0 = time.time()
            while time.time() - t0 < 60 and self.running:
                u = driver.current_url
                if "/login" not in u and "app" in u: return True
                time.sleep(1)
        except Exception as e: self.print_erro(f"Erro login CH: {e}")
        return False

    def ch_navegar_detalhe(self, driver: webdriver.Chrome, wait: WebDriverWait) -> bool:
        self.print_info("Navegando ao menu Detalhe por linha...")
        try: wait.until(EC.invisibility_of_element_located((By.ID, "loading-bar")))
        except: pass
        time.sleep(2)
        try:
            wait.until(EC.element_to_be_clickable((By.XPATH, "//mat-icon[contains(text(), 'menu')]"))).click()
            time.sleep(1)
            wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(., 'Logística')]"))).click()
            time.sleep(1)
            wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(., 'Detalhe por linha')]"))).click()
            time.sleep(3)
            return True
        except: return False

    def ch_processar_normal(self, driver: webdriver.Chrome, wait: WebDriverWait, aba: Any) -> None:
        self.print_header("Processando Aba Normal - Cargo Heroes")
        d = aba.get_all_records()
        try: col = aba.row_values(1).index("CH OK") + 1
        except: return
        for i, l in enumerate(d, start=2):
            if not self.running: break
            mat = str(l.get('Material ID', '')).strip()
            if not mat or str(l.get('CH OK', '')).upper() == "OK": continue
            self.print_info(f"Linha {i}: Editando {mat}...")
            try:
                if not self.ch_busca_material(driver, wait, mat): raise Exception("Falha busca")
                self.ch_acao(driver, wait, "//*[@id='dataTable']/tbody/tr/td[1]/a/i", "clicar")
                time.sleep(3)
                
                req = str(l.get('REQUISIÇÃO', ''))
                self.ch_acao(driver, wait, "//input[@formcontrolname='requirement']", "escrever", texto=req)
                self.ch_acao(driver, wait, "//*[@formcontrolname='typeAtd']", "clicar")
                time.sleep(0.5)
                self.ch_acao(driver, wait, "//mat-option//span[contains(text(), 'Material')]", "clicar")
                
                tm = "Aéreo" if "Aéreo" in str(l.get('Tipo de Transporte', '')) else "Terrestre"
                self.ch_acao(driver, wait, "//*[@formcontrolname='modal']", "clicar")
                time.sleep(0.5)
                self.ch_acao(driver, wait, f"//mat-option//span[contains(text(), '{tm}')]", "clicar")
                
                self.ch_acao(driver, wait, "//input[@formcontrolname='origin']", "escrever", texto=str(l.get('Origem Sigla')))
                self.ch_acao(driver, wait, "//input[@formcontrolname='destination']", "escrever", texto=str(l.get('Destino Sigla')))
                
                tl = str(l.get('Logística', ''))
                self.ch_acao(driver, wait, "//input[@formcontrolname='desc']", "escrever", texto=tl)
                
                hs, hc = self.ch_extrair_horarios(tl)
                if hs and hc:
                    ds, dc = self.ch_calcular_data_hora(hs), self.ch_calcular_data_hora(hc)
                    if ds and dc:
                        self.ch_preencher_data_js(driver, wait, "//*[@id='dateBoarding']", ds)
                        self.ch_preencher_data_js(driver, wait, "//*[@id='dateLanding']", dc)
                
                self.ch_acao(driver, wait, "//button[contains(., 'Salve')] | //button[contains(., 'Salvar')]", "clicar")
                try:
                    wait.until(EC.visibility_of_element_located((By.XPATH, "//*[contains(text(), 'Solicitação Atendida')]")))
                    aba.update_cell(i, col, "OK")
                    self.print_sucesso(f"Linha {i} OK")
                except: raise Exception("Confirmação não apareceu")
                
                driver.refresh(); time.sleep(3); self.ch_navegar_detalhe(driver, wait)
            except Exception as e:
                self.print_erro(f"Erro linha {i}: {e}")
                aba.update_cell(i, col, "ERRO")
                driver.refresh(); time.sleep(3); self.ch_navegar_detalhe(driver, wait)

    def ch_processar_mapeamento(self, driver: webdriver.Chrome, wait: WebDriverWait, aba: Any) -> None:
        self.print_header("Processando Aba Mapeamento - Cargo Heroes")
        d = aba.get_all_records()
        try: col = aba.row_values(1).index("CH OK") + 1
        except: return
        for i, l in enumerate(d, start=2):
            if not self.running: break
            mat = str(l.get('Material ID', '')).strip()
            st = str(l.get('CH OK', '')).strip().upper()
            orig = str(l.get('ORIGEM', '')).strip().upper()
            if not mat or st == "OK" or ("NA BASE" not in orig and "ZERO" not in orig): continue
            
            self.print_info(f"Mapeamento Linha {i}: {mat}")
            try:
                if not self.ch_busca_material(driver, wait, mat): raise Exception("Falha busca")
                self.ch_acao(driver, wait, "//*[@id='dataTable']/tbody/tr/td[1]/a/i", "clicar")
                time.sleep(3)
                
                if "NA BASE" in orig: self.ch_acao(driver, wait, "//button[contains(., 'Mtl na Base')]", "clicar")
                elif "ZERO" in orig: self.ch_acao(driver, wait, "//button[contains(., 'Mtl Stk Zero')]", "clicar")
                
                self.ch_acao(driver, wait, "//button[contains(., 'Salve')] | //button[contains(., 'Salvar')]", "clicar")
                try:
                    wait.until(EC.visibility_of_element_located((By.XPATH, "//*[contains(text(), 'Solicitação Atendida')]")))
                    aba.update_cell(i, col, "OK")
                    self.print_sucesso(f"Linha {i} Mapeamento OK")
                except: raise Exception("Sem confirmação")
                driver.refresh(); time.sleep(3); self.ch_navegar_detalhe(driver, wait)
            except Exception as e:
                self.print_erro(f"Erro mapeamento {i}: {e}")
                driver.refresh(); time.sleep(3); self.ch_navegar_detalhe(driver, wait)


def main() -> None:
    root = ctk.CTk()
    app = SAPAutomationGUI(root)
    root.attributes('-topmost', True)
    root.update()
    root.attributes('-topmost', False)
    root.mainloop()

if __name__ == "__main__":
    main()