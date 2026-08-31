# -*- coding: utf-8 -*-
"""
Automação Integrada — SAP & Cargo Heroes
Criação massiva de Requisições de Compra (ME51N) e atualização no Cargo Heroes via API.

Refatoração Profissional UI (CustomTkinter LATAM Theme) + Correções de Segurança/Threading + API Bypass
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
import ssl
import certifi
from typing import Optional, Any
import json

# --- Imports Selenium ---
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

# --- Imports Keyring (armazenamento seguro de senhas) ---
try:
    import keyring
    KEYRING_DISPONIVEL = True
except ImportError:
    KEYRING_DISPONIVEL = False

# --- Retry para chamadas de API ---
try:
    from tenacity import retry, wait_exponential, stop_after_attempt
    TENACITY_DISPONIVEL = True
except ImportError:
    TENACITY_DISPONIVEL = False

# Constantes de Segurança
KEYRING_SERVICE_SAP = "sap_automation_req_massivo"
KEYRING_SERVICE_CH = "cargo_heroes_automation"

# Configuração SSL segura
ssl._create_default_https_context = lambda: ssl.create_default_context(cafile=certifi.where())

# Configuração CustomTkinter (Modo Claro/Light para fundo branco)
ctk.set_appearance_mode("light")
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
        except Exception:
            pass

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
        
        # JANELA MAIS COMPACTA (650x450)
        self.root.geometry("650x450")
        self.root.minsize(650, 450)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Paleta de Cores LATAM Airlines
        self.COR_AZUL_LATAM = "#2A0087"
        self.COR_VERMELHO_LATAM = "#EE1750"
        self.COR_AZUL_HOVER = "#1A0050"
        self.COR_VERMELHO_HOVER = "#B3103A"

        self._session_lock = threading.Lock()
        self._session: Optional[Any] = None
        self.running: bool = False

        self.config = configparser.ConfigParser()
        self.data_path: str = self.get_data_path()
        self.resource_path: str = self.get_resource_path()
        
        self.config_path: str = os.path.join(self.data_path, 'config.ini')

        try:
            self.config.read(self.config_path, encoding='utf-8')
        except Exception:
            self.create_default_config()

        self.load_icon()
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

    def get_data_path(self) -> str:
        if getattr(sys, 'frozen', False):
            return os.path.dirname(sys.executable)
        return os.path.dirname(os.path.abspath(__file__))

    def get_resource_path(self) -> str:
        if getattr(sys, 'frozen', False):
            return sys._MEIPASS
        return os.path.dirname(os.path.abspath(__file__))

    def load_icon(self) -> None:
        icon_start_b64 = self.fix_base64_padding("iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsTAAALEwEAmpwYAAABWWlUWHRYTUw6Y29tLmFkb2JlLnhtcAAAAAAAPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iWE1QIENvcmUgNS40LjAiPgogICA8cmRmOlJERiB4bWxuczpyZGY9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkvMDIvMjItcmRmLXN5bnRheC1ucyMiPgogICAgICA8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIgogICAgICAgICAgICB4bWxuczp0aWZmPSJodHRwOi8vbnMuYWRvYmUuY29tL3RpZmYvMS4wLyI+CiAgICAgICAgIDx0aWZmOk9yaWVudGF0aW9uPjE8L3RpZmY6T3JpZW50YXRpb24+CiAgICAgIDwvcmRmOkRlc2NyaXB0aW9uPgogICA8L3JkZjpSREY+CjwveDp4bXBtZXRhPgpMTE82AAAByklEQVQ4EaVTTUhUURQ+V3d1JzUzDBf9B0sLw0lDItocoQd9QUG3LhpEL7pw7aJdKyIi6FYQ2oZ1UXQRhFAb60PQg1ZCi2kStZpoGjcz9768eW/G6Mwbw72X+37n3HPvPQCB8Q+QnL8xQCeAF/D/o/kOWG01P4BHgCfz30KMR8s+AaYAPND9GkAy8BpwGSgLHsA74JESi2kUeC2APdC6BN8B+u0mYwTo/2QjE58D2pXfAYwA215W25gBFoD1/R9IZzGgYv03gC2gC1gDvgIHADgACsAecD6hL8eA/4A/AT8BfwL6n2WnAY+A+cBj4DHvEY+BTYBT4B/gY+An4G9gOvgZfAReA/4FvgK/AqcB14DwwB+H/t9w9OAD+An/d8D/4L/EXkX2gN8A1XkLwNngGfA58Ar4DNwFvgE/A28Bv4Cfi/sXv0C/A78DPgIeBR4DvgE+A54AvAE+A9YAhYABYAzYC/gH2APWApWAcvAdqAfeAesAyuBvA74Anjkl+sN4EPgG+A3wBvgS5Y4ADwGjgIfAm8BD4GngCPANeA0cAW4BnwG3AcuArdBB/gMvAasA/8CPga+Ad8AvwB+A/YD/gD+AfwB+D/AP8C/gb8F/gX+A/wH/A78CvwNfA1sApsBGgHNgNngCXgGfAYeBIMAP+A18B6sAmsA58AW8AC8A2YAxaBLWAb2A3uAx8AZ4BnwEvAm8BTYB/4BOwBDoCTQBDwBrgD3AE+Ax4AZoCNb+V2d+AocB14B/gJ+BG4C1wBfgE+A74DPgX+D/gb+B/wL/A/4DvgF+BGYAhZ/f+AasAm8BWwA8+B14BTwGnAEOAX8B7wG3gV8H4L/f7bB/yXgG/A18CXwDvgQ2AH+AF8CJ4DngHPAK8Ar4PXgGfA78DPwN/Ad8A3wL+Bn5B/gP8F/rX7F/gV+B/wI/AJ8K/AnwA/8AMxS88wI6H8lAAAAAElFTSuQmCC")
        icon_stop_b64 = self.fix_base64_padding("iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAA7SURBVDhPY/wPBAxUACZA1gBTgA2K/1EGBgaG/2A8HIBoMRhIMeAgYADEDwFGBgYGDjA50uAzAAgwAK0/AwMT5urRAAAAAElFTSuQmCC")

        try:
            self.icons = {
                "start": ctk.CTkImage(light_image=Image.open(io.BytesIO(base64.b64decode(icon_start_b64))), size=(16, 16)),
                "stop": ctk.CTkImage(light_image=Image.open(io.BytesIO(base64.b64decode(icon_stop_b64))), size=(16, 16))
            }
        except Exception:
            self.icons = {"start": None, "stop": None}

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
        # Título menor e com menos margem
        self.header_label = ctk.CTkLabel(self.root, text="Automação Integrada - SAP & Cargo Heroes", font=ctk.CTkFont(size=18, weight="bold"), text_color=self.COR_AZUL_LATAM)
        self.header_label.pack(pady=(10, 5))

        self.tabview = ctk.CTkTabview(self.root, width=620, height=330,
                                      segmented_button_selected_color=self.COR_AZUL_LATAM,
                                      segmented_button_selected_hover_color=self.COR_AZUL_HOVER)
        self.tabview.pack(padx=15, pady=0, fill="both", expand=True)
        self.tabview.add("Automação")
        self.tabview.add("Configurações")

        self.setup_main_tab(self.tabview.tab("Automação"))
        self.setup_config_tab(self.tabview.tab("Configurações"))

        self.status_frame = ctk.CTkFrame(self.root, height=25, corner_radius=0, fg_color="transparent")
        self.status_frame.pack(side="bottom", fill="x", padx=15, pady=(0, 5))

        self.sap_status_var = tk.StringVar(value="SAP: Desconectado")
        self.sap_status_label = ctk.CTkLabel(self.status_frame, textvariable=self.sap_status_var, text_color=self.COR_VERMELHO_LATAM, font=ctk.CTkFont(weight="bold", size=11))
        self.sap_status_label.pack(side="left")

        self.status_var = tk.StringVar(value="Pronto")
        self.status_label = ctk.CTkLabel(self.status_frame, textvariable=self.status_var, text_color="#666666", font=ctk.CTkFont(size=11))
        self.status_label.pack(side="right")

    def setup_main_tab(self, parent: ctk.CTkFrame) -> None:
        control_frame = ctk.CTkFrame(parent, fg_color="transparent")
        control_frame.pack(pady=(5, 10))

        # Botões reduzidos para height=32
        self.start_button = ctk.CTkButton(
            control_frame, text="Iniciar SAP", image=self.icons["start"],
            fg_color=self.COR_AZUL_LATAM, hover_color=self.COR_AZUL_HOVER,
            command=self.start_automation, width=140, height=32, font=ctk.CTkFont(size=12, weight="bold")
        )
        self.start_button.pack(side="left", padx=8)

        self.ch_button = ctk.CTkButton(
            control_frame, text="Atualizar CH", image=self.icons["start"],
            fg_color=self.COR_AZUL_LATAM, hover_color=self.COR_AZUL_HOVER,
            command=self.start_ch_automation, width=140, height=32, font=ctk.CTkFont(size=12, weight="bold")
        )
        self.ch_button.pack(side="left", padx=8)

        self.stop_button = ctk.CTkButton(
            control_frame, text="Parar Automação", image=self.icons["stop"],
            fg_color=self.COR_VERMELHO_LATAM, hover_color=self.COR_VERMELHO_HOVER, state="disabled",
            command=self.stop_automation, width=140, height=32, font=ctk.CTkFont(size=12, weight="bold")
        )
        self.stop_button.pack(side="left", padx=8)

        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ctk.CTkProgressBar(parent, variable=self.progress_var, progress_color=self.COR_AZUL_LATAM, height=8)
        self.progress_bar.pack(fill="x", padx=10, pady=(0, 5))

        self.log_area = ctk.CTkTextbox(parent, font=ctk.CTkFont(family="Consolas", size=11), 
                                       fg_color="#F2F2F2", text_color="#333333", 
                                       border_width=1, border_color="#DDDDDD")
        self.log_area.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.log_area.configure(state="disabled")

    def _configure_log_tags(self):
        self.log_area.tag_config("RESET", foreground="#333333")
        self.log_area.tag_config("VERDE", foreground="#008000")
        self.log_area.tag_config("AMARELO", foreground="#CC7700")
        self.log_area.tag_config("VERMELHO", foreground=self.COR_VERMELHO_LATAM)
        self.log_area.tag_config("AZUL", foreground=self.COR_AZUL_LATAM)
        self.log_area.tag_config("CIANO", foreground="#00838F")

    def setup_config_tab(self, parent: ctk.CTkFrame) -> None:
        scroll_frame = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll_frame.pack(fill="both", expand=True)

        # Margens (pady) reduzidas drasticamente
        sap_frame = ctk.CTkFrame(scroll_frame, fg_color="#F9F9F9", border_color="#DDDDDD", border_width=1)
        sap_frame.pack(fill="x", pady=4, padx=5)
        ctk.CTkLabel(sap_frame, text="Credenciais SAP", font=ctk.CTkFont(weight="bold", size=13), text_color=self.COR_AZUL_LATAM).pack(anchor="w", padx=10, pady=(5, 0))

        self.sap_path_var = tk.StringVar(value=self.config.get('SAP', 'caminho_logon', fallback=''))
        self.create_config_row(sap_frame, "Caminho Logon.exe:", self.sap_path_var, show_browse=True)
        self.sap_system_var = tk.StringVar(value=self.config.get('SAP', 'sistema', fallback=''))
        self.create_config_row(sap_frame, "Sistema:", self.sap_system_var)
        self.sap_user_var = tk.StringVar(value=self.config.get('SAP', 'usuario', fallback=''))
        self.create_config_row(sap_frame, "Usuário:", self.sap_user_var)
        self.sap_password_var = tk.StringVar(value=self._obter_senha(KEYRING_SERVICE_SAP, 'senha', 'SAP', 'senha'))
        self.create_config_row(sap_frame, "Senha:", self.sap_password_var, is_password=True)

        google_frame = ctk.CTkFrame(scroll_frame, fg_color="#F9F9F9", border_color="#DDDDDD", border_width=1)
        google_frame.pack(fill="x", pady=4, padx=5)
        ctk.CTkLabel(google_frame, text="Google Sheets", font=ctk.CTkFont(weight="bold", size=13), text_color=self.COR_AZUL_LATAM).pack(anchor="w", padx=10, pady=(5, 0))

        self.google_creds_var = tk.StringVar(value=self.config.get('GOOGLE', 'credenciais', fallback=''))
        self.create_config_row(google_frame, "Credenciais:", self.google_creds_var, show_browse=True)
        self.google_sheet_var = tk.StringVar(value=self.config.get('GOOGLE', 'planilha', fallback=''))
        self.create_config_row(google_frame, "Planilha:", self.google_sheet_var)
        self.google_tab_var = tk.StringVar(value=self.config.get('GOOGLE', 'aba', fallback=''))
        self.create_config_row(google_frame, "Aba:", self.google_tab_var)

        ch_frame = ctk.CTkFrame(scroll_frame, fg_color="#F9F9F9", border_color="#DDDDDD", border_width=1)
        ch_frame.pack(fill="x", pady=4, padx=5)
        ctk.CTkLabel(ch_frame, text="Cargo Heroes", font=ctk.CTkFont(weight="bold", size=13), text_color=self.COR_AZUL_LATAM).pack(anchor="w", padx=10, pady=(5, 0))

        self.ch_email_var = tk.StringVar(value=self.config.get('CARGO_HEROES', 'email', fallback=''))
        self.create_config_row(ch_frame, "Email:", self.ch_email_var)
        self.ch_pass_var = tk.StringVar(value=self._obter_senha(KEYRING_SERVICE_CH, 'senha', 'CARGO_HEROES', 'senha'))
        self.create_config_row(ch_frame, "Senha:", self.ch_pass_var, is_password=True)

        ctk.CTkButton(scroll_frame, text="Salvar Configurações", command=self.save_config, 
                      width=200, height=32, font=ctk.CTkFont(weight="bold", size=12), 
                      fg_color=self.COR_AZUL_LATAM, hover_color=self.COR_AZUL_HOVER).pack(pady=15)

    def create_config_row(self, parent, label_text, variable, is_password=False, show_browse=False):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", padx=10, pady=2) # Pady bem reduzido
        
        lbl = ctk.CTkLabel(frame, text=label_text, width=110, anchor="e", font=ctk.CTkFont(size=11), text_color="#333333")
        lbl.pack(side="left", padx=(0, 5))
        
        show_char = "*" if is_password else ""
        # Altura do input reduzida para 28
        entry = ctk.CTkEntry(frame, textvariable=variable, show=show_char, height=28, fg_color="white", text_color="black", border_color="#CCCCCC", font=ctk.CTkFont(size=11))
        entry.pack(side="left", fill="x", expand=True)
        
        if show_browse:
            # Altura do botão selecionar reduzida para 28
            btn = ctk.CTkButton(frame, text="Selecionar", width=70, height=28, command=lambda: self.browse_file(variable), fg_color=self.COR_AZUL_LATAM, hover_color=self.COR_AZUL_HOVER, font=ctk.CTkFont(size=11))
            btn.pack(side="left", padx=(5, 0))

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
            self.sap_status_label.configure(text_color="#008000") # Verde
        else:
            self.sap_status_var.set("SAP: Desconectado")
            self.sap_status_label.configure(text_color=self.COR_VERMELHO_LATAM)
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
        print(f"<<AZUL>>\n[{self._get_timestamp()}] 🚀 {texto.upper()}\n<<RESET>>")

    def print_sucesso(self, texto: str) -> None:
        print(f"<<VERDE>>[{self._get_timestamp()}] ✔  SUCESSO:  {texto}\n<<RESET>>")

    def print_info(self, texto: str) -> None:
        print(f"<<CIANO>>[{self._get_timestamp()}] ℹ  INFO:     {texto}\n<<RESET>>")

    def print_aviso(self, texto: str) -> None:
        print(f"<<AMARELO>>[{self._get_timestamp()}] ⚠  AVISO:    {texto}\n<<RESET>>")

    def print_erro(self, texto: str) -> None:
        print(f"<<VERMELHO>>[{self._get_timestamp()}] ✖  ERRO:     {texto}\n<<RESET>>")

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
    #  AUTOMAÇÃO CARGO HEROES VIA API (RÁPIDA)
    # =========================================================================

    def ch_extrair_horarios(self, texto_logistica: str) -> tuple[Optional[str], Optional[str]]:
        texto = str(texto_logistica).strip()
        padrao_hora = r'\b(?:[01]?\d|2[0-3]):[0-5]\d\b'
        horarios = re.findall(padrao_hora, texto)
        if len(horarios) >= 2: return horarios[0], horarios[1]
        return None, None

    def ch_calcular_data_hora_iso(self, horario_str: str) -> Optional[str]:
        try:
            agora = datetime.now()
            parts = horario_str.split(':')
            hora, minuto = int(parts[0]), int(parts[1])
            data_alvo = agora.replace(hour=hora, minute=minuto, second=0, microsecond=0)
            if data_alvo < agora - timedelta(hours=6):
                data_alvo += timedelta(days=1)
            return data_alvo.strftime("%Y-%m-%dT%H:%M:%S.000Z")
        except: return None

    # -------------------------------------------------------------------------
    # FUNÇÕES INJETORAS DE API (JAVASCRIPT -> BROWSER -> SERVIDOR)
    # -------------------------------------------------------------------------
    def ch_atualizar_normal_api(self, driver: webdriver.Chrome, material: str, dados: dict) -> dict:
        email_usuario = self.ch_email_var.get()
        js_script = """
        var done = arguments[arguments.length - 1];
        var material = arguments[0];
        var dados = arguments[1];
        var userEmail = arguments[2];

        var token = sessionStorage.getItem('acme-user-token');
        if (!token) { done({ok: false, error: 'TOKEN_NOT_FOUND'}); return; }

        // --- CORREÇÃO DE FUSO HORÁRIO (Soma 3h para anular o desconto automático do Cargo Heroes) ---
        function fixTimezone(val) {
            if (!val) return null;
            var strVal = String(val);
            
            // Se for string ISO (ex: "2026-08-31T19:00:00")
            if (strVal.indexOf('T') !== -1) {
                var cleanDate = strVal.replace('Z', '').replace('-03:00', '').replace('+00:00', '');
                var parts = cleanDate.split('T');
                var dParts = parts[0].split('-');
                var tParts = parts[1].split(':');
                
                var year = parseInt(dParts[0], 10);
                var month = parseInt(dParts[1], 10) - 1;
                var day = parseInt(dParts[2], 10);
                var hour = parseInt(tParts[0], 10);
                var min = parseInt(tParts[1], 10);
                var sec = tParts.length > 2 ? parseInt(tParts[2].split('.')[0], 10) : 0;
                
                // Cria o timestamp UTC adicionando 3 horas (+3) para compensar o fuso de Brasília
                return String(Date.UTC(year, month, day, hour + 3, min, sec));
            }
            
            // Se já for timestamp direto do Python, adiciona 3 horas em milissegundos (10800000 ms)
            if (!isNaN(strVal) && strVal.length > 10) {
                return String(parseInt(strVal, 10) + 10800000);
            }
            return val;
        }

        var payloadSearch = {
            extended: {
                states: [],
                statesRq: ["1", "2", "3", "4", "5", "6", "7", "8"],
                page: { page: 0, size: 20 },
                orderBy: [{ attribute: "equipmentCode", direction: "desc" }],
                statusExclude: false,
                expired: "",
                lang: "PT",
                gmt: "GMT-0300"
            },
            requestCode: "", requestDate: "", baseId: "", aircraftId: "", criticalId: "", 
            barcode: "", userId: "", description: "", requirement: "", 
            equipmentCode: parseInt(material, 10), 
            eqTypeId: "", partNumber: "", origin: "", destination: ""
        };

        var headersCargos = { 
            'Authorization': 'Bearer ' + token, 
            'Content-Type': 'application/json',
            'Accept': 'application/json, text/plain, */*',
            'X-Application-Name': 'LATAM',
            'x-user-logon': String(userEmail).toUpperCase(),
            'x-user-module': 'Logistic',
            'x-user-screen': 'DetailLineForm'
        };

        fetch('/api/bff/requests/equipments/logistics/search', {
            method: 'POST', headers: headersCargos, credentials: 'include', mode: 'cors',
            body: JSON.stringify(payloadSearch)
        })
        .then(res => res.json())
        .then(list => {
            if (!list || !list.equipments || list.equipments.length === 0) {
                throw new Error("Material não encontrado no banco de dados.");
            }

            var itemCompleto = list.equipments[0];
            var logisticData = itemCompleto.logistic;
            var equipmentData = itemCompleto.equipment;
            
            var requestId = itemCompleto.request.requestCode;
            var equipmentCode = equipmentData.equipmentCode;

            var modalCode = dados.modal === 'Aéreo' ? 1 : 2;
            var modalName = dados.modal === 'Aéreo' ? 'Aéreo' : 'Terrestre';

            // =========================================================
            // PACOTE 1: Atualização de Logística
            // =========================================================
            var updateLogistics = Object.assign({}, logisticData);
            updateLogistics.requirement = String(dados.req);
            updateLogistics.description = dados.texto;
            updateLogistics.modal = { modalCode: modalCode, modalName: modalName, languages: null, langs: logisticData.modal ? logisticData.modal.langs : null };
            updateLogistics.type = { typeCode: 3, equipmentType: "Logistic", name: "Material", languages: null }; 
            
            var agoraMs = String(Date.now()); 
            updateLogistics.requisitionDate = agoraMs;
            updateLogistics.updatedDate = agoraMs;
            
            var oldFlight = logisticData.flight || {};
            updateLogistics.flight = {
                origin: { baseCode: String(dados.origem).toUpperCase(), baseState: "1" },
                destination: { baseCode: String(dados.destino).toUpperCase(), baseState: "1" },
                // Função fixTimezone aplicada aos horários
                dateBoarding: fixTimezone(dados.dateBoarding) || oldFlight.dateBoarding || null,
                dateLanding: fixTimezone(dados.dateLanding) || oldFlight.dateLanding || null,
                flightConnection: oldFlight.flightConnection || null,
                finalDestination: oldFlight.finalDestination || null
            };
            updateLogistics.userId = userEmail;

            // =========================================================
            // PACOTE 2: Atualização de Equipamento (Status e Tipo)
            // =========================================================
            var STATE_AGUARDANDO_SEPARACAO = 6; 
            var tipoMaterialOuFerramenta = equipmentData.type ? equipmentData.type.typeCode : 1;

            var updateEquipment = {
                extended: equipmentData.extended,
                type: { typeCode: tipoMaterialOuFerramenta }, 
                state: { stateCode: STATE_AGUARDANDO_SEPARACAO },
                equipmentCode: equipmentData.equipmentCode,
                partNumber: equipmentData.partNumber,
                codOnu: equipmentData.codOnu,
                dgrCode: equipmentData.dgrCode,
                iw: equipmentData.iw,
                description: equipmentData.description,
                quantity: equipmentData.quantity,
                observations: equipmentData.observations,
                userId: userEmail
            };

            var urlLogistics = '/api/bff/requests/' + requestId + '/equipments/' + equipmentCode + '/logistics';
            var urlEquipment = '/api/bff/requests/' + requestId + '/equipments/' + equipmentCode + '/upd';

            return Promise.all([
                fetch(urlLogistics, { method: 'POST', headers: headersCargos, credentials: 'include', mode: 'cors', body: JSON.stringify(updateLogistics) }),
                fetch(urlEquipment, { method: 'POST', headers: headersCargos, credentials: 'include', mode: 'cors', body: JSON.stringify(updateEquipment) })
            ]);
        })
        .then(responses => {
            for(let res of responses) {
                if(!res.ok) throw new Error("HTTP " + res.status + " em uma das rotas de salvamento.");
            }
            done({ok: true});
        })
        .catch(err => done({ok: false, error: err.toString()}));
        """
        try:
            driver.set_script_timeout(15)
            return driver.execute_async_script(js_script, material, dados, email_usuario)
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def ch_atualizar_mapeamento_api(self, driver: webdriver.Chrome, material: str, acao_str: str) -> dict:
        email_usuario = self.ch_email_var.get()
        js_script = """
        var done = arguments[arguments.length - 1];
        var material = arguments[0];
        var acao = arguments[1]; 
        var userEmail = arguments[2];

        var token = sessionStorage.getItem('acme-user-token');
        if (!token) { done({ok: false, error: 'TOKEN_NOT_FOUND'}); return; }

        var payloadSearch = {
            extended: {
                states: [],
                statesRq: ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "23"], // Expandido para cobrir os novos status
                page: { page: 0, size: 20 },
                orderBy: [{ attribute: "equipmentCode", direction: "desc" }],
                statusExclude: false,
                expired: "",
                lang: "PT",
                gmt: "GMT-0300"
            },
            requestCode: "", requestDate: "", baseId: "", aircraftId: "", criticalId: "", 
            barcode: "", userId: "", description: "", requirement: "", 
            equipmentCode: parseInt(material, 10), 
            eqTypeId: "", partNumber: "", origin: "", destination: ""
        };

        var headersCargos = { 
            'Authorization': 'Bearer ' + token, 
            'Content-Type': 'application/json',
            'Accept': 'application/json, text/plain, */*',
            'X-Application-Name': 'LATAM',
            'x-user-logon': String(userEmail).toUpperCase(),
            'x-user-module': 'Logistic',
            'x-user-screen': 'DetailLineForm'
        };

        fetch('/api/bff/requests/equipments/logistics/search', {
            method: 'POST', headers: headersCargos, credentials: 'include', mode: 'cors',
            body: JSON.stringify(payloadSearch)
        })
        .then(res => res.json())
        .then(list => {
            if (!list || !list.equipments || list.equipments.length === 0) {
                throw new Error("Material não encontrado no banco de dados.");
            }

            var itemCompleto = list.equipments[0];
            var equipmentData = itemCompleto.equipment; 
            
            var requestId = itemCompleto.request.requestCode;
            var equipmentCode = equipmentData.equipmentCode;

            // Define os State Codes exatos capturados nos seus prints
            var novoStateCode = acao === 'BASE' ? 10 : 23; 
            
            // Verifica o tipo do equipamento (se não existir, assume Material = 1)
            var tipoMaterialOuFerramenta = equipmentData.type ? equipmentData.type.typeCode : 1;

            var updatePayload = {
                extended: equipmentData.extended,
                type: { typeCode: tipoMaterialOuFerramenta },
                state: { stateCode: novoStateCode },
                equipmentCode: equipmentData.equipmentCode,
                partNumber: equipmentData.partNumber,
                codOnu: equipmentData.codOnu,
                dgrCode: equipmentData.dgrCode,
                iw: equipmentData.iw,
                description: equipmentData.description,
                quantity: equipmentData.quantity,
                observations: equipmentData.observations,
                userId: userEmail
            };

            return fetch('/api/bff/requests/' + requestId + '/equipments/' + equipmentCode + '/upd', { 
                method: 'POST', headers: headersCargos, credentials: 'include', mode: 'cors',
                body: JSON.stringify(updatePayload)
            });
        })
        .then(res => {
            if(!res.ok) return res.text().then(errText => { throw new Error("HTTP " + res.status + " - " + errText); });
            done({ok: true});
        })
        .catch(err => done({ok: false, error: err.toString()}));
        """
        try:
            driver.set_script_timeout(15)
            return driver.execute_async_script(js_script, material, acao_str, email_usuario)
        except Exception as e:
            return {"ok": False, "error": str(e)}
        
    # -------------------------------------------------------------------------
    # FLUXO PRINCIPAL - CARGO HEROES
    # -------------------------------------------------------------------------
    def start_ch_automation(self) -> None:
        if self.running: return
        email = self.ch_email_var.get()
        senha = self.ch_pass_var.get()
        if not email or not senha:
            messagebox.showerror("Erro", "Configure Email e Senha do CH.")
            return
        self.running = True
        self.toggle_buttons(False)
        self.status_var.set("Executando Cargo Heroes via API...")
        self.log_area.configure(state="normal")
        self.log_area.delete("1.0", "end")
        self.log_area.configure(state="disabled")
        self.setup_log_redirector()
        threading.Thread(target=self.run_ch_automation, daemon=True).start()

    def run_ch_automation(self) -> None:
        driver = None
        try:
            self.print_header("Iniciando Cargo Heroes Updater (Modo API Rápida)")
            opts = Options()
            opts.add_argument("--start-maximized")
            opts.add_argument("--disable-blink-features=AutomationControlled")
            opts.add_experimental_option("excludeSwitches", ["enable-automation"])
            opts.add_experimental_option('useAutomationExtension', False)
            
            driver = webdriver.Chrome(options=opts)
            driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {
                'source': "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
            })
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
                self.print_info("Login Realizado. Carregando sessão interna do sistema...")
                driver.get("https://cargo-heroes.appslatam.com/#/app/requests")
                time.sleep(4) 
                
                # --- MELHORIA 1: Minimiza o navegador para rodar em "Segundo Plano" ---
                driver.minimize_window()
                self.print_info("Navegador minimizado. O processamento via API continuará em segundo plano.")
                
                if self.running: self.ch_processar_normal(driver, aba_n)
                if self.running and aba_m: self.ch_processar_mapeamento(driver, aba_m)
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
        self.print_info("🔑 Realizando Login Automático (SSO/SAML LATAM)...")
        try:
            email = self.ch_email_var.get()
            senha = self.ch_pass_var.get()
            
            janela_principal = driver.current_window_handle
            
            # 1. Aciona o Botão do Google
            try:
                iframe = wait.until(EC.presence_of_element_located((By.XPATH, "//iframe[contains(@src, 'gsi/button')]")))
                driver.switch_to.frame(iframe)
                btn_google = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@role='button']")))
                driver.execute_script("arguments[0].click();", btn_google)
                driver.switch_to.default_content()
            except: 
                driver.switch_to.default_content()
            
            # 2. Máquina de Estados para lidar com o Popup Dinâmico
            try:
                WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
                janela_popup = [j for j in driver.window_handles if j != janela_principal][0]
                driver.switch_to.window(janela_popup)
                
                t_end = time.time() + 60
                while time.time() < t_end:
                    if len(driver.window_handles) == 1:
                        break # O popup fechou sozinho, login concluído!
                        
                    try:
                        url_popup = driver.current_url.lower()
                        
                        # FASE A: Inserir Email no Google
                        if "accounts.google.com" in url_popup and "identifier" in url_popup:
                            try:
                                inp = driver.find_element(By.XPATH, "//input[@type='email' or @id='identifierId']")
                                if inp.is_displayed() and inp.get_attribute('value') == "":
                                    inp.clear()
                                    inp.send_keys(email)
                                    inp.send_keys(Keys.ENTER)
                                    time.sleep(2)
                            except: pass

                        # FASE B: Tela "Confirme que é você"
                        elif "speedbump/samlconfirmaccount" in url_popup:
                            try:
                                btn = driver.find_element(By.XPATH, "//span[text()='Continuar' or text()='Continue']/ancestor::button")
                                if btn.is_displayed():
                                    driver.execute_script("arguments[0].click();", btn)
                                    time.sleep(2)
                            except: pass

                        # FASE C: Microsoft SSO (Email e Senha)
                        elif "microsoftonline.com" in url_popup or "live.com" in url_popup:
                            try:
                                # Insere Email (se pedir)
                                inp_email = driver.find_element(By.ID, "i0116")
                                if inp_email.is_displayed() and inp_email.get_attribute('value') == "":
                                    inp_email.clear()
                                    inp_email.send_keys(email)
                                    inp_email.send_keys(Keys.ENTER)
                                    time.sleep(2)
                            except: pass
                            
                            try:
                                # Insere Senha
                                inp_senha = driver.find_element(By.ID, "i0118")
                                if inp_senha.is_displayed() and inp_senha.get_attribute('value') == "":
                                    inp_senha.clear()
                                    inp_senha.send_keys(senha)
                                    inp_senha.send_keys(Keys.ENTER)
                                    time.sleep(2)
                            except: pass
                            
                            try:
                                # Botão "Continuar conectado?"
                                btn_yes = driver.find_element(By.ID, "idSIButton9")
                                if btn_yes.is_displayed():
                                    btn_yes.click()
                                    time.sleep(2)
                            except: pass
                        
                        time.sleep(1)
                    except:
                        time.sleep(1)
                
                # Foco de volta à janela principal
                try: driver.switch_to.window(janela_principal)
                except: pass

            except Exception as popup_err:
                self.print_aviso(f"Tentativa de bypass manual de popup. Aguardando. Detalhe: {popup_err}")
                try: driver.switch_to.window(janela_principal)
                except: pass
            
            # 3. Confirmação do sucesso
            t0 = time.time()
            while time.time() - t0 < 60 and self.running:
                u = driver.current_url
                if "/login" not in u and "app" in u: 
                    return True
                time.sleep(1)
                
        except Exception as e: 
            self.print_erro(f"Erro fluxo login CH: {e}")
            
        return False
    
    def ch_processar_normal(self, driver: webdriver.Chrome, aba: Any) -> None:
        self.print_header("Processando Aba Normal (Via API) - Cargo Heroes")
        d = aba.get_all_records()
        try: col = aba.row_values(1).index("CH OK") + 1
        except: return
        
        updates_planilha = [] 
        
        for i, l in enumerate(d, start=2):
            if not self.running: break
            
            mat = str(l.get('Material ID', '')).strip()
            if not mat or str(l.get('CH OK', '')).upper() == "OK": continue
            
            self.print_info(f"Linha {i}: Editando ID de {mat} via API...")
            
            req = str(l.get('REQUISIÇÃO', ''))
            tm = "Aéreo" if "Aéreo" in str(l.get('Tipo de Transporte', '')) else "Terrestre"
            origem = str(l.get('Origem Sigla', ''))
            destino = str(l.get('Destino Sigla', ''))
            tl = str(l.get('Logística', ''))
            
            hs, hc = self.ch_extrair_horarios(tl)
            ds_iso, dc_iso = None, None
            if hs and hc:
                ds_iso = self.ch_calcular_data_hora_iso(hs)
                dc_iso = self.ch_calcular_data_hora_iso(hc)

            dados_api = {
                "req": req, "modal": tm, "origem": origem, "destino": destino,
                "texto": tl, "dateBoarding": ds_iso, "dateLanding": dc_iso
            }

            resultado = self.ch_atualizar_normal_api(driver, mat, dados_api)
            
            celula_a1 = gspread.utils.rowcol_to_a1(i, col)
            if resultado.get("ok"):
                updates_planilha.append({'range': celula_a1, 'values': [["OK"]]})
                self.print_sucesso(f"Linha {i} ({mat}) Atualizada Instantaneamente!")
            else:
                self.print_erro(f"Erro na linha {i} ({mat}): {resultado.get('error')}")
                updates_planilha.append({'range': celula_a1, 'values': [["ERRO"]]})
            
            time.sleep(0.3)
            
        if updates_planilha:
            self.print_info("Sincronizando resultados com o Google Sheets...")
            self._batch_update_planilha(aba, updates_planilha)

    def ch_processar_mapeamento(self, driver: webdriver.Chrome, aba: Any) -> None:
        self.print_header("Processando Aba Mapeamento (Via API) - Cargo Heroes")
        d = aba.get_all_records()
        try: col = aba.row_values(1).index("CH OK") + 1
        except: return
        
        updates_planilha = []
        
        for i, l in enumerate(d, start=2):
            if not self.running: break
            
            mat = str(l.get('Material ID', '')).strip()
            st = str(l.get('CH OK', '')).strip().upper()
            orig = str(l.get('ORIGEM', '')).strip().upper()
            
            if not mat or st == "OK" or ("NA BASE" not in orig and "ZERO" not in orig): continue
            
            acao = 'BASE' if 'NA BASE' in orig else 'ZERO'
            self.print_info(f"Mapeamento Linha {i}: Ajustando {mat} ({acao}) via API...")
            
            resultado = self.ch_atualizar_mapeamento_api(driver, mat, acao)
            
            celula_a1 = gspread.utils.rowcol_to_a1(i, col)
            if resultado.get("ok"):
                updates_planilha.append({'range': celula_a1, 'values': [["OK"]]})
                self.print_sucesso(f"Linha {i} ({mat}) Mapeada Instantaneamente!")
            else:
                self.print_erro(f"Erro na linha {i} ({mat}): {resultado.get('error')}")
                updates_planilha.append({'range': celula_a1, 'values': [["ERRO"]]})
            
            time.sleep(0.3)
            
        if updates_planilha:
            self.print_info("Sincronizando resultados do mapeamento com o Google Sheets...")
            self._batch_update_planilha(aba, updates_planilha)

if __name__ == "__main__":
    root = ctk.CTk()
    app = SAPAutomationGUI(root)
    root.attributes('-topmost', True)
    root.update()
    root.attributes('-topmost', False)
    root.mainloop()