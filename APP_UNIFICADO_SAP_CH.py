# -*- coding: utf-8 -*-
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
from tkinter import ttk, scrolledtext, filedialog, messagebox
import threading
from PIL import Image, ImageTk
import io
import base64
import logging
import ssl
import pyautogui

# --- Imports Selenium ---
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException, NoSuchWindowException

# Ajuste SSL
ssl._create_default_https_context = ssl._create_unverified_context

# --- Classes e Funções de Interface Gráfica ---

class LogRedirector:
    """Redireciona a saída para o widget de texto e aplica cores com base em tags."""
    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.tag_map = {
            "<<RESET>>": "RESET",
            "<<VERDE>>": "VERDE",
            "<<AMARELO>>": "AMARELO",
            "<<VERMELHO>>": "VERMELHO",
            "<<AZUL>>": "AZUL",
            "<<CIANO>>": "CIANO"
        }
        self.default_tag = "RESET"

    def write(self, string):
        try:
            self.text_widget.config(state=tk.NORMAL)
            segments = re.split(f"({'|'.join(re.escape(k) for k in self.tag_map.keys())})", string)
            current_tag = self.default_tag
            for segment in segments:
                if segment in self.tag_map:
                    current_tag = self.tag_map[segment]
                elif segment:
                    self.text_widget.insert(tk.END, segment, current_tag)
            self.text_widget.see(tk.END)
            self.text_widget.config(state=tk.DISABLED)
        except Exception:
            pass

    def flush(self):
        pass

class SAPAutomationGUI:
    # Mapeamento de Depósitos por Origem (Hardcoded conforme solicitação)
    DEPOSITO_MAPPING = {
        'BR0G': 'AE01', 'BR0Q': 'AE01', 'BR0D': 'AE01', 'BR0H': 'AE01', 'BR0O': 'AE01',
        'BR0P': 'AE01', 'BR0E': 'AE01', 'BR0R': 'AE01', 'BR0S': 'AE01', 'BR0Y': 'AE01',
        'BR0Z': 'AE01', 'BR1A': 'AE01', 'BR1C': 'AE01', 'BR1D': 'AE01', 'BR1G': 'AE01',
        'BR1I': 'AE01', 'BR1J': 'AE01', 'BR1K': 'AE01', 'BR1L': 'AE01', 'BR1T': 'AE01',
        'BR2A': 'AE01', 'BR2B': 'AE01', 'BR2C': 'AE01', 'BR2D': 'AE01', 'BR2E': 'AE01',
        'BR2Q': 'AE01', 'BR2U': 'AE01', 'BR2V': 'AE01', 'BR3A': 'AE01', 'BR3E': 'AE01',
        'BR3F': 'AE01', 'BR3K': 'AE01', 'BR3N': 'AE01', 'BRDN': 'AE01', 'BR8A': 'AE13',
        'BR2I': 'AE01', 'BR0I': 'AE13', 'BR0U': 'AE01', 'BR0K': 'AE13', 'BR0X': 'AE13',
        'BR0J': 'AE01', 'BR1E': 'AE01', 'BR1F': 'AE01', 'BR0V': 'AE01', 'BR8E': 'AE13',
        'BR1B': 'AE01', 'BR0F': 'AE01', 'BR8I': 'AE01', 'BRIJ': 'AE01', 'BR8G': 'AE01'
    }

    def __init__(self, root):
        self.root = root
        self.root.title("Automação Integrada - SAP & Cargo Heroes")
        self.root.geometry("1000x700")
        self.root.minsize(900, 600)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.configure(bg='#2E2E2E')

        # --- Estilo Dark Mode ---
        style = ttk.Style()
        style.theme_use('clam')
        BG_COLOR = "#2E2E2E"
        FG_COLOR = "#FFFFFF"
        LIGHT_BG = "#3C3C3C"
        BORDER_COLOR = "#505050"
        SELECT_BG = "#555555"

        style.configure('.', background=BG_COLOR, foreground=FG_COLOR, fieldbackground=LIGHT_BG, bordercolor=BORDER_COLOR)
        style.map('.', background=[('active', SELECT_BG)])
        style.configure('TNotebook', background=BG_COLOR, borderwidth=0)
        style.configure('TNotebook.Tab', background=BG_COLOR, foreground=FG_COLOR, padding=[10, 5], font=('Segoe UI', 10))
        style.map('TNotebook.Tab', background=[('selected', LIGHT_BG)], foreground=[('selected', 'white')])
        style.configure('TFrame', background=BG_COLOR)
        style.configure('TLabelframe', background=BG_COLOR, bordercolor=BORDER_COLOR)
        style.configure('TLabelframe.Label', background=BG_COLOR, foreground=FG_COLOR, font=('Segoe UI', 10, 'bold'))
        style.configure('TButton', padding=6, relief="flat", font=('Segoe UI', 10), background=LIGHT_BG, foreground=FG_COLOR)
        style.map('TButton', background=[('active', SELECT_BG)])
        style.configure('Start.TButton', background='#4CAF50', foreground='white', font=('Segoe UI', 11, 'bold'))
        style.map('Start.TButton', background=[('active', '#45a049')])
        style.configure('CH.TButton', background='#2196F3', foreground='white', font=('Segoe UI', 11, 'bold'))
        style.map('CH.TButton', background=[('active', '#0b7dda')])
        style.configure('Stop.TButton', background='#f44336', foreground='white', font=('Segoe UI', 11, 'bold'))
        style.map('Stop.TButton', background=[('active', '#da190b')])
        style.configure('TEntry', fieldbackground=LIGHT_BG, foreground=FG_COLOR, bordercolor=BORDER_COLOR, insertcolor=FG_COLOR)

        # Ícones Base64
        self.icons = {
            "start": self.create_icon_from_base64("iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsTAAALEwEAmpwYAAABWWlUWHRYTUw6Y29tLmFkb2JlLnhtcAAAAAAAPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iWE1QIENvcmUgNS40LjAiPgogICA8cmRmOlJERiB4bWxuczpyZGY9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkvMDIvMjItcmRmLXN5bnRheC1ucyMiPgogICAgICA8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIgogICAgICAgICAgICB4bWxuczp0aWZmPSJodHRwOi8vbnMuYWRvYmUuY29tL3RpZmYvMS4wLyI+CiAgICAgICAgIDx0aWZmOk9yaWVudGF0aW9uPjE8L3RpZmY6T3JpZW50YXRpb24+CiAgICAgIDwvcmRmOkRlc2NyaXB0aW9uPgogICA8L3JkZjpSREY+CjwveDp4bXBtZXRhPgpMTE82AAAByklEQVQ4EaVTTUhUURQ+V3d1JzUzDBf9B0sLw0lDItocoQd9QUG3LhpEL7pw7aJdKyIi6FYQ2oZ1UXQRhFAb60PQg1ZCi2kStZpoGjcz9768eW/G6Mwbw72X+37n3HPvPQCB8Q+QnL8xQCeAF/D/o/kOWG01P4BHgCfz30KMR8s+AaYAPND9GkAy8BpwGSgLHsA74JESi2kUeC2APdC6BN8B+u0mYwTo/2QjE58D2pXfAYwA215W25gBFoD1/R9IZzGgYv03gC2gC1gDvgIHADgACsAecD6hL8eA/4A/AT8BfwL6n2WnAY+A+cBj4DHvEY+BTYBT4B/gY+An4G9gOvgZfAReA/4FvgK/AqcB14DwwB+H/t9w9OAD+An/d8D/4L/EXkX2gN8A1XkLwNngGfA58Ar4DNwFvgE/A28Bv4Cfi/sXv0C/A78DPgIeBR4DvgE+A54AvAE+A9YAhYABYAzYC/gH2APWApWAcvAdqAfeAesAyuBvA74Anjkl+sN4EPgG+A3wBvgS5Y4ADwGjgIfAm8BD4GngCPANeA0cAW4BnwG3AcuArdBB/gMvAasA/8CPga+Ad8AvwB+A/YD/gD+AfwB+D/AP8C/gb8F/gX+A/wH/A78CvwNfA1sApsBGgHNgNngCXgGfAYeBIMAP+A18B6sAmsA58AW8AC8A2YAxaBLWAb2A3uAx8AZ4BnwEvAm8BTYB/4BOwBDoCTQBDwBrgD3AE+Ax4AZoCNb+V2d+AocB14B/gJ+BG4C1wBfgE+A74DPgX+D/gb+B/wL/A/4DvgF+BGYAhZ/f+AasAm8BWwA8+B14BTwGnAEOAX8B7wG3gV8H4L/f7bB/yXgG/A18CXwDvgQ2AH+AF8CJ4DngHPAK8Ar4PXgGfA78DPwN/Ad8A3wL+Bn5B/gP8F/rX7F/gV+B/wI/AJ8K/AnwA/8AMxS88wI6H8lAAAAAElFTSuQmCC"),
            "stop": self.create_icon_from_base64("iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAA7SURBVDhPY/wPBAxUACZA1gBTgA2K/1EGBgaG/2A8HIBoMRhIMeAgYADEDwFGBgYGDjA50uAzAAgwAK0/AwMT5urRAAAAAElFTSuQmCC"),
        }
        
        self.running = False
        self.session = None
        self.config = configparser.ConfigParser()
        self.base_path = self.get_base_path()
        self.config_path = os.path.join(self.base_path, 'config.ini')
        
        # Configuração de logs
        self.log_to_file = True
        self.log_file_path = os.path.join(self.base_path, 'app_log.txt')
        self.logs_print_path = os.path.join(self.base_path, 'logs/prints')
        
        if not os.path.exists(self.logs_print_path):
            os.makedirs(self.logs_print_path, exist_ok=True)

        if self.log_to_file:
            with open(self.log_file_path, 'a', encoding='utf-8') as log_file:
                log_file.write(f"\n{'='*50}\nSessão iniciada em {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
                log_file.write(f"Modo: {'Executável' if self.is_executable() else 'Script Python'}\n")
                log_file.write(f"Caminho base: {self.base_path}\n{'='*50}\n")
        
        try:
            self.config.read(self.config_path, encoding='utf-8')
        except Exception:
            self.create_default_config()
        
        self.load_icon()
        self.create_widgets()
        
        # Tags de Cor Log (FONTE REDUZIDA PARA MELHOR VISUALIZAÇÃO)
        self.log_area.tag_configure("RESET", foreground="#D0D0D0", font=('Consolas', 8))
        self.log_area.tag_configure("VERDE", foreground="#66bb6a", font=('Consolas', 8, 'bold'))
        self.log_area.tag_configure("AMARELO", foreground="#ffa726", font=('Consolas', 8, 'bold'))
        self.log_area.tag_configure("VERMELHO", foreground="#ef5350", font=('Consolas', 8, 'bold'))
        self.log_area.tag_configure("AZUL", foreground="#42a5f5", font=('Consolas', 8, 'bold'))
        self.log_area.tag_configure("CIANO", foreground="#26c6da", font=('Consolas', 8, 'bold'))
        
    def create_icon_from_base64(self, base64_string):
        try:
            img_data = base64.b64decode(base64_string)
            img = Image.open(io.BytesIO(img_data))
            return ImageTk.PhotoImage(img)
        except Exception:
            return None
            
    def get_base_path(self):
        if getattr(sys, 'frozen', False):
            return os.path.dirname(sys.executable)
        return os.path.dirname(os.path.abspath(__file__))
    
    def is_executable(self):
        return getattr(sys, 'frozen', False)

    def load_icon(self):
        try:
            icon_path = os.path.join(self.base_path, 'icone.ico')
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception: pass
    
    def create_default_config(self):
        self.config['SAP'] = {'caminho_logon': r'C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe', 'sistema': 'ECC PRODUÇÃO', 'usuario': '', 'senha': ''}
        self.config['GOOGLE'] = {'credenciais': 'credentials.json', 'planilha': 'Nome da Planilha', 'aba': 'Aba1'}
        self.config['CARGO_HEROES'] = {'email': '', 'senha': ''}
        with open(self.config_path, 'w', encoding='utf-8') as configfile:
            self.config.write(configfile)
    
    def create_widgets(self):
        self.menu_bar = tk.Menu(self.root, bg="#2E2E2E", fg="white", activebackground="#555555", relief=tk.FLAT)
        self.root.config(menu=self.menu_bar)
        file_menu = tk.Menu(self.menu_bar, tearoff=0, bg="#3C3C3C", fg="white")
        self.menu_bar.add_cascade(label="Arquivo", menu=file_menu)
        file_menu.add_command(label="Sair", command=self.on_closing)
        
        self.status_frame = ttk.Frame(self.root, style='TFrame')
        self.status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Status SAP
        self.sap_status_var = tk.StringVar(value="SAP: Desconectado")
        self.sap_status_label = ttk.Label(self.status_frame, textvariable=self.sap_status_var, foreground="red", padding=(5, 2))
        self.sap_status_label.pack(side=tk.LEFT)
        ttk.Separator(self.status_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)

        self.status_var = tk.StringVar(value="Pronto")
        self.status_label = ttk.Label(self.status_frame, textvariable=self.status_var, padding=(5, 2))
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.notebook = ttk.Notebook(self.root, style='TNotebook')
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.main_tab = ttk.Frame(self.notebook, style='TFrame')
        self.notebook.add(self.main_tab, text="Automação")
        self.config_tab = ttk.Frame(self.notebook, style='TFrame')
        self.notebook.add(self.config_tab, text="Configurações")
        
        self.setup_main_tab()
        self.setup_config_tab()
    
    def setup_main_tab(self):
        control_frame = ttk.Frame(self.main_tab, style='TFrame')
        control_frame.pack(pady=15)
        
        self.start_button = ttk.Button(control_frame, text="Iniciar SAP", image=self.icons["start"], compound=tk.LEFT, command=self.start_automation, style='Start.TButton', width=18)
        self.start_button.pack(side=tk.LEFT, padx=10, ipady=5)
        
        self.ch_button = ttk.Button(control_frame, text="Atualizar CH", image=self.icons["start"], compound=tk.LEFT, command=self.start_ch_automation, style='CH.TButton', width=18)
        self.ch_button.pack(side=tk.LEFT, padx=10, ipady=5)
        
        self.stop_button = ttk.Button(control_frame, text="Parar Automação", image=self.icons["stop"], compound=tk.LEFT, command=self.stop_automation, state=tk.DISABLED, style='Stop.TButton', width=18)
        self.stop_button.pack(side=tk.LEFT, padx=10, ipady=5)
        
        log_frame = ttk.LabelFrame(self.main_tab, text="Log de Execução", style='TLabelframe')
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        # FONTE REDUZIDA AQUI TAMBÉM (size 8)
        self.log_area = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, state=tk.DISABLED, font=("Consolas", 8), bg="#252526", fg="#D4D4D4", relief=tk.FLAT, borderwidth=0, insertbackground="white")
        self.log_area.pack(fill=tk.BOTH, expand=True, padx=2, pady=2) # Reduzi o padding interno também

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.main_tab, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, padx=10, pady=(5, 0))
    
    def setup_config_tab(self):
        config_main_frame = ttk.Frame(self.config_tab, padding="10", style='TFrame')
        config_main_frame.pack(fill=tk.BOTH, expand=True)

        # SAP
        sap_frame = ttk.LabelFrame(config_main_frame, text="Configurações SAP", padding="10", style='TLabelframe')
        sap_frame.pack(fill=tk.X, expand=False, pady=5)
        
        ttk.Label(sap_frame, text="Caminho SAP:").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.sap_path_var = tk.StringVar(value=self.config.get('SAP', 'caminho_logon', fallback=''))
        ttk.Entry(sap_frame, textvariable=self.sap_path_var, width=50).grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Button(sap_frame, text="...", command=lambda: self.browse_file(self.sap_path_var)).grid(row=0, column=2, padx=5)

        ttk.Label(sap_frame, text="Sistema:").grid(row=1, column=0, sticky=tk.W, padx=5)
        self.sap_system_var = tk.StringVar(value=self.config.get('SAP', 'sistema', fallback=''))
        ttk.Entry(sap_frame, textvariable=self.sap_system_var, width=30).grid(row=1, column=1, sticky=tk.W, padx=5)

        ttk.Label(sap_frame, text="Usuário:").grid(row=2, column=0, sticky=tk.W, padx=5)
        self.sap_user_var = tk.StringVar(value=self.config.get('SAP', 'usuario', fallback=''))
        ttk.Entry(sap_frame, textvariable=self.sap_user_var, width=30).grid(row=2, column=1, sticky=tk.W, padx=5)

        ttk.Label(sap_frame, text="Senha:").grid(row=3, column=0, sticky=tk.W, padx=5)
        self.sap_password_var = tk.StringVar(value=self.config.get('SAP', 'senha', fallback=''))
        ttk.Entry(sap_frame, textvariable=self.sap_password_var, show="*", width=30).grid(row=3, column=1, sticky=tk.W, padx=5)

        # GOOGLE
        google_frame = ttk.LabelFrame(config_main_frame, text="Configurações Google Sheets", padding="10", style='TLabelframe')
        google_frame.pack(fill=tk.X, expand=False, pady=5)
        
        ttk.Label(google_frame, text="JSON Credenciais:").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.google_creds_var = tk.StringVar(value=self.config.get('GOOGLE', 'credenciais', fallback=''))
        ttk.Entry(google_frame, textvariable=self.google_creds_var, width=50).grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Button(google_frame, text="...", command=lambda: self.browse_file(self.google_creds_var)).grid(row=0, column=2, padx=5)
        
        ttk.Label(google_frame, text="Planilha:").grid(row=1, column=0, sticky=tk.W, padx=5)
        self.google_sheet_var = tk.StringVar(value=self.config.get('GOOGLE', 'planilha', fallback=''))
        ttk.Entry(google_frame, textvariable=self.google_sheet_var, width=30).grid(row=1, column=1, sticky=tk.W, padx=5)
        
        ttk.Label(google_frame, text="Aba:").grid(row=2, column=0, sticky=tk.W, padx=5)
        self.google_tab_var = tk.StringVar(value=self.config.get('GOOGLE', 'aba', fallback=''))
        ttk.Entry(google_frame, textvariable=self.google_tab_var, width=30).grid(row=2, column=1, sticky=tk.W, padx=5)

        # CARGO HEROES
        ch_frame = ttk.LabelFrame(config_main_frame, text="Configurações Cargo Heroes", padding="10", style='TLabelframe')
        ch_frame.pack(fill=tk.X, expand=False, pady=5)

        ttk.Label(ch_frame, text="Email CH:").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.ch_email_var = tk.StringVar(value=self.config.get('CARGO_HEROES', 'email', fallback=''))
        ttk.Entry(ch_frame, textvariable=self.ch_email_var, width=40).grid(row=0, column=1, sticky=tk.W, padx=5)

        ttk.Label(ch_frame, text="Senha CH:").grid(row=1, column=0, sticky=tk.W, padx=5)
        self.ch_pass_var = tk.StringVar(value=self.config.get('CARGO_HEROES', 'senha', fallback=''))
        ttk.Entry(ch_frame, textvariable=self.ch_pass_var, show="*", width=30).grid(row=1, column=1, sticky=tk.W, padx=5)

        ttk.Button(config_main_frame, text="Salvar Configurações", command=self.save_config).pack(pady=20)
    
    def browse_file(self, var):
        filename = filedialog.askopenfilename()
        if filename: var.set(filename)
    
    def save_config(self):
        self.config['SAP']['caminho_logon'] = self.sap_path_var.get()
        self.config['SAP']['sistema'] = self.sap_system_var.get()
        self.config['SAP']['usuario'] = self.sap_user_var.get()
        self.config['SAP']['senha'] = self.sap_password_var.get()
        self.config['GOOGLE']['credenciais'] = self.google_creds_var.get()
        self.config['GOOGLE']['planilha'] = self.google_sheet_var.get()
        self.config['GOOGLE']['aba'] = self.google_tab_var.get()
        self.config['CARGO_HEROES']['email'] = self.ch_email_var.get()
        self.config['CARGO_HEROES']['senha'] = self.ch_pass_var.get()
        try:
            with open(self.config_path, 'w', encoding='utf-8') as configfile:
                self.config.write(configfile)
            messagebox.showinfo("Sucesso", "Configurações salvas!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar: {str(e)}")

    def on_closing(self):
        if self.running:
            if not messagebox.askyesno("Confirmação", "A automação está em execução. Deseja sair?"): return
            self.stop_automation()
        self.restore_stdout()
        self.root.destroy()

    def toggle_buttons(self, state):
        s = tk.NORMAL if state else tk.DISABLED
        ns = tk.DISABLED if state else tk.NORMAL
        self.start_button.config(state=s)
        self.ch_button.config(state=s)
        self.stop_button.config(state=ns)
    
    def atualizar_status_sap(self, conectado=False, mensagem=None):
        if conectado:
            self.sap_status_var.set("SAP: Conectado")
            self.sap_status_label.configure(foreground="#66bb6a")
        else:
            self.sap_status_var.set("SAP: Desconectado")
            self.sap_status_label.configure(foreground="#ef5350")
        if mensagem:
            self.status_var.set(mensagem)

    def validate_config(self):
        return all([
            self.sap_path_var.get(), self.sap_system_var.get(),
            self.google_creds_var.get(), self.google_sheet_var.get(), self.google_tab_var.get()
        ])
    
    def update_progress(self, value):
        self.root.after(0, lambda: self.progress_var.set(value))

    # --- Lógica de Log ---
    def setup_log_redirector(self):
        sys.stdout = LogRedirector(self.log_area)

    def restore_stdout(self):
        sys.stdout = sys.__stdout__

    def print_header(self, texto):
        log_text = f"\n{'='*60}\n {texto.center(58)}\n {'='*60}\n"
        print(f"<<AZUL>>{log_text}<<RESET>>")
        self._write_to_log_file(log_text)

    def print_sucesso(self, texto):
        log_text = f"[SUCESSO] {texto}\n"
        print(f"<<VERDE>>{log_text}<<RESET>>")
        self._write_to_log_file(log_text)

    def print_info(self, texto):
        log_text = f"[INFO]    {texto}\n"
        print(f"<<CIANO>>{log_text}<<RESET>>")
        self._write_to_log_file(log_text)

    def print_aviso(self, texto):
        log_text = f"[AVISO]   {texto}\n"
        print(f"<<AMARELO>>{log_text}<<RESET>>")
        self._write_to_log_file(log_text)

    def print_erro(self, texto):
        log_text = f"[ERRO]    {texto}\n"
        print(f"<<VERMELHO>>{log_text}<<RESET>>")
        self._write_to_log_file(log_text)

    def _write_to_log_file(self, text_to_log):
        if self.log_to_file:
            try:
                timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                cleaned_text = re.sub(r'<<.*?>>', '', text_to_log)
                with open(self.log_file_path, 'a', encoding='utf-8') as log_file:
                    log_file.write(f"[{timestamp}] {cleaned_text.strip()}\n")
            except Exception: pass

    # =========================================================================
    #  AUTOMAÇÃO SAP
    # =========================================================================
    def start_automation(self):
        if self.running: return
        
        if not self.validate_config():
            messagebox.showerror("Erro", "Preencha todas as configurações SAP/Google.")
            return

        self.running = True
        self.toggle_buttons(False)
        self.status_var.set("Executando Automação SAP...")
        
        self.log_area.config(state=tk.NORMAL)
        self.log_area.delete(1.0, tk.END)
        self.log_area.config(state=tk.DISABLED)
        
        self.setup_log_redirector()
        
        # Iniciar thread do SAP
        self.automation_thread = threading.Thread(target=self.run_sap_automation, daemon=True)
        self.automation_thread.start()

    def stop_automation(self):
        if not self.running: return
        self.running = False
        self.status_var.set("Parando automação...")
        self.print_aviso("Solicitação de parada recebida. Aguarde o fim do ciclo atual...")

    def run_sap_automation(self):
        try:
            pythoncom.CoInitialize()
            self.print_info("COM inicializado para threading")
        except Exception as e:
            self.print_erro(f"Erro ao inicializar COM: {str(e)}")
            
        try:
            self.print_header("Iniciando Robô de Requisição de Compra no SAP")
            
            # Conexão SAP
            if not self.is_session_valid():
                self.print_aviso("Sessão SAP inválida ou inexistente. Tentando conectar...")
                self.root.after(0, lambda: self.atualizar_status_sap(False, "Conectando ao SAP..."))
                self.session = self.sap_login_handler()

            if not self.session:
                self.print_erro("Falha na conexão com o SAP. Verifique as configurações e se o SAP está acessível.")
                self.root.after(0, lambda: self.atualizar_status_sap(False, "Falha na conexão"))
                return
            
            self.print_sucesso("Sessão SAP estabelecida com sucesso!")
            self.root.after(0, lambda: self.atualizar_status_sap(True, "Conectado ao SAP"))
            
            # Processamento Planilha
            try:
                self.print_header("CONECTANDO À PLANILHA")
                credentials_file = self.google_creds_var.get()
                gc = gspread.service_account(filename=credentials_file)
                spreadsheet = gc.open(self.google_sheet_var.get())
                worksheet = spreadsheet.worksheet(self.google_tab_var.get())
                self.print_sucesso("Conexão com a planilha estabelecida.")
                
                headers = worksheet.row_values(1)
                status_col_index = headers.index("Status") + 1
                req_col_index = headers.index("REQUISIÇÃO") + 1
                
                df = pd.DataFrame(worksheet.get_all_records())
                df['linha_planilha'] = df.index + 2
                # Considera apenas linhas sem status
                df_para_processar = df[df['Status'] == ''].copy()

                if df_para_processar.empty:
                    self.print_aviso("Nenhuma linha nova para validar e criar.")
                else:
                    self.processar_lotes(df_para_processar, worksheet, status_col_index, req_col_index)

            except Exception as e:
                self.print_erro(f"Erro crítico no ciclo principal: {e}")
                self.session = None 
                self.root.after(0, lambda: self.atualizar_status_sap(False, f"Erro SAP"))
            
        except Exception as e:
            self.print_erro(f"Erro fatal na automação: {str(e)}")
        finally:
            self.print_header("FIM DO CICLO")
            self.root.after(0, self.finalize_automation)
            pythoncom.CoUninitialize()

    # --- Helpers SAP ---
    
    def finalize_automation(self):
        if self.running:
            self.print_info("Automação concluída.")
            self.running = False
        else:
            self.print_aviso("Automação interrompida pelo usuário.")

        self.toggle_buttons(True)
        self.status_var.set("Pronto para iniciar")
        self.update_progress(0)
        self.restore_stdout()
        
    def aguardar_sap(self, timeout=30):
        if not self.session: return False
        start_time = time.time()
        while self.running:
            try:
                if not self.session.busy: return True
            except: return False
            if time.time() - start_time > timeout:
                self.print_aviso(f"Timeout ao aguardar SAP após {timeout} segundos")
                return False
            time.sleep(0.2)
        return False

    def is_session_valid(self):
        if self.session is None: return False
        try:
            self.session.findById("wnd[0]")
            return True
        except (pywintypes.com_error, Exception):
            return False

    def sap_login_handler(self):
        try:
            self.print_info("Procurando por uma sessão SAP GUI...")
            sap_gui_auto = win32com.client.GetObject("SAPGUI")
            application = sap_gui_auto.GetScriptingEngine
            
            if application.Connections.Count > 0:
                for conn_idx in range(application.Connections.Count):
                    connection = application.Connections(conn_idx)
                    if connection.Sessions.Count > 0:
                        for session_idx in range(connection.Sessions.Count):
                            session = connection.Sessions(session_idx)
                            try:
                                session.findById("wnd[0]")
                                self.print_sucesso(f"Sessão SAP ativa encontrada.")
                                return session
                            except: continue
            
            self.print_aviso("Nenhuma sessão SAP válida encontrada. Iniciando nova conexão...")
            return self.open_and_login_sap()
        except (pywintypes.com_error, Exception):
            self.print_aviso("Iniciando processo de login...")
            return self.open_and_login_sap()

    def open_and_login_sap(self):
        try:
            sap_path = self.config['SAP']['caminho_logon']
            sap_system = self.config['SAP']['sistema'].strip()
            
            if not os.path.exists(sap_path):
                self.print_erro(f"Caminho do SAP Logon não encontrado: '{sap_path}'")
                return None
                
            self.print_info(f"Abrindo SAP Logon...")
            subprocess.Popen(sap_path)
            time.sleep(5)
            
            sap_gui_auto = win32com.client.GetObject("SAPGUI")
            application = sap_gui_auto.GetScriptingEngine
            
            self.print_info(f"Conectando ao sistema '{sap_system}'...")
            connection = application.OpenConnection(sap_system, True)
            time.sleep(3)
            session = connection.Children(0)
            
            start_time = time.time()
            while session.busy:
                time.sleep(0.5)
                if time.time() - start_time > 30: return None
            
            main_window = session.findById("wnd[0]")
            
            try:
                main_window.findById("usr/txtRSYST-BNAME")
                self.print_info("Preenchendo credenciais...")
                user = self.config.get('SAP', 'usuario')
                password = self.config.get('SAP', 'senha')
                
                main_window.findById("usr/txtRSYST-BNAME").text = user
                main_window.findById("usr/pwdRSYST-BCODE").text = password
                main_window.sendVKey(0)

                start_time = time.time()
                while session.busy:
                    time.sleep(0.5)
                    if time.time() - start_time > 30: return None
                
                try: session.findById("wnd[1]").sendVKey(0) 
                except: pass
                
                if "easy access" in session.findById("wnd[0]").text.lower() or "menú" in session.findById("wnd[0]").text.lower():
                    self.print_sucesso("Login no SAP realizado com sucesso!")
                    return session
                else:
                    self.print_erro(f"Falha no login: {session.findById('sbar').text}")
                    return None
            except:
                self.print_sucesso("Sessão existente detectada.")
                return session

        except Exception as e:
            self.print_erro(f"Erro crítico login: {str(e)}")
            return None

    def processar_lotes(self, df_para_processar, worksheet, status_col_index, req_col_index):
        self.print_info(f"Encontradas {len(df_para_processar)} linhas pendentes.")
        
        # 1. Definir data do dia (Variável solicitada)
        data_hoje = datetime.now().strftime('%d.%m.%Y')
        self.print_info(f"Data de Remessa definida para hoje: {data_hoje}")

        # 2. Agrupamento Inteligente (Substituindo 'Validador')
        # PRIORIZAR ORIGEM/DESTINO (Correção solicitada)
        col_origem = 'ORIGEM' if 'ORIGEM' in df_para_processar.columns else 'Origem Sigla'
        col_destino = 'DESTINO' if 'DESTINO' in df_para_processar.columns else 'Destino Sigla'
        
        # Agrupa por Origem e Destino
        grupos = df_para_processar.groupby([col_origem, col_destino])
        
        lotes_para_processar = []
        
        for (origem, destino), grupo in grupos:
            # Divide cada grupo em pedaços de 10 linhas
            for i in range(0, len(grupo), 10):
                chunk = grupo.iloc[i : i + 10].copy()
                # Adiciona índice do grid (0 a 9) necessário para o SAP preencher a tabela
                chunk['grid_index'] = range(len(chunk))
                lotes_para_processar.append(chunk)
        
        total_lotes = len(lotes_para_processar)
        self.print_info(f"Total de RCs a serem criadas (Lotes): {total_lotes}")
        
        for idx, lote_df in enumerate(lotes_para_processar):
            if not self.running: break
            if not self.is_session_valid():
                self.print_erro("Sessão SAP perdida. Tentando reconectar...")
                self.session = self.sap_login_handler()
                if not self.session: break
            
            self.update_progress(((idx + 1) / total_lotes) * 100)
            origem_val = lote_df.iloc[0][col_origem]
            destino_val = lote_df.iloc[0][col_destino]
            self.print_header(f"Processando Lote {idx + 1}/{total_lotes} | {origem_val} -> {destino_val}")
            
            # --- Validação (com data automática) ---
            resultados = self.validar_lote_na_rc(lote_df, data_hoje)
            
            # Atualização da Planilha (Validação)
            validation_updates = []
            linhas_ok = []
            for res in resultados:
                validation_updates.append({'range': f'{gspread.utils.rowcol_to_a1(res["linha_planilha"], status_col_index)}', 'values': [[str(res['status'])]]})
                validation_updates.append({'range': f'{gspread.utils.rowcol_to_a1(res["linha_planilha"], req_col_index)}', 'values': [[str(res['numero_rc'])]]})
                if res['status'] == 'OK': linhas_ok.append(res['linha_planilha'])

            if validation_updates:
                try: worksheet.batch_update(validation_updates)
                except Exception as e: self.print_erro(f"Erro update planilha: {e}")

            # --- Criação (apenas itens OK) ---
            if not linhas_ok:
                self.print_aviso("Nenhum item válido neste lote. Pulando criação.")
                continue
                
            lote_df_ok = lote_df[lote_df['linha_planilha'].isin(linhas_ok)].copy()
            # Reindexa grid_index sequencialmente para o lote de criação
            lote_df_ok['grid_index'] = range(len(lote_df_ok))
            
            if not self.is_session_valid():
                self.print_erro("Sessão SAP perdida.")
                break
            
            numero_rc, msg_status = self.criar_rc_para_lote_ok(lote_df_ok, data_hoje)
            
            # Atualização Final
            creation_updates = []
            for linha in lote_df_ok['linha_planilha']:
                creation_updates.append({'range': f'{gspread.utils.rowcol_to_a1(linha, status_col_index)}', 'values': [[str(msg_status)]]})
                if numero_rc:
                    creation_updates.append({'range': f'{gspread.utils.rowcol_to_a1(linha, req_col_index)}', 'values': [[str(numero_rc)]]})
            
            if creation_updates:
                try:
                    worksheet.batch_update(creation_updates)
                    self.print_sucesso("RC Criada e Planilha atualizada.")
                except Exception as e: self.print_erro(f"Erro update final: {e}")

    def validar_lote_na_rc(self, lote_de_itens, data_hoje):
        if lote_de_itens.empty: return []
        resultados_finais = []
        try:
            self.print_header(f"Validando Lote ({len(lote_de_itens)} itens)")
            self.session.findById("wnd[0]").maximize()
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/NME51N"
            self.session.findById("wnd[0]").sendVKey(0)
            self.aguardar_sap()
            time.sleep(1)
            self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:3327/cmbMEREQ_TOPLINE-BSART").key = "ZRT"
            self.session.findById("wnd[0]").sendVKey(0)
            self.aguardar_sap()
            grid = self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell")
            
            for _, item in lote_de_itens.iterrows():
                if not self.running: break
                
                # Usa grid_index gerado no processar_lotes
                grid_index = int(item['grid_index'])
                
                # Mapeamento de colunas (Prioridade para ORIGEM/DESTINO do SAP)
                mat_id = item.get('PN') or item.get('Material ID')
                origem = item.get('ORIGEM') or item.get('Origem Sigla')
                destino = item.get('DESTINO') or item.get('Destino Sigla')
                qtd = str(item.get('QTD') or item.get('Quantidade', '1')).replace(',', '.')
                # Texto/Logística
                texto = item.get('TEXTO') or item.get('Logística')
                
                status_item = "OK"
                self.print_header(f"Processando Item {grid_index + 1} (Mat: {mat_id})")
                try:
                    grid.modifyCell(grid_index, "MATNR", str(mat_id))
                    grid.modifyCell(grid_index, "MENGE", qtd)
                    grid.modifyCell(grid_index, "RESWK", str(origem))
                    grid.modifyCell(grid_index, "EEIND", data_hoje) # Usa data automática
                    grid.modifyCell(grid_index, "EPSTP", "U")
                    grid.modifyCell(grid_index, "NAME1", str(destino))
                    grid.modifyCell(grid_index, "EKGRP", "P04")
                    grid.modifyCell(grid_index, "TXZ01", str(texto))
                    self.session.findById("wnd[0]").sendVKey(0)
                    self.aguardar_sap()
                    time.sleep(1.5)
                    try: self.session.findById("wnd[1]").sendVKey(0)
                    except: pass
                    
                    status_bar = self.session.findById("wnd[0]/sbar")
                    if status_bar.messageType in ('E', 'A') or "não está atualizado no centro" in status_bar.text:
                        status_item = status_bar.text
                        self.print_erro(f"Erro: {status_item}")
                    else:
                        status_item = "OK"
                        self.print_sucesso("Item OK")
                except Exception as e:
                    status_item = f"Erro crítico: {str(e)}"
                    self.print_erro(status_item)
                resultados_finais.append({'linha_planilha': item['linha_planilha'], 'status': status_item, 'numero_rc': '' if status_item == 'OK' else 'ERRO'})
            return resultados_finais
        finally:
            try:
                if self.is_session_valid():
                    self.session.findById("wnd[0]/tbar[0]/okcd").text = "/N"
                    self.session.findById("wnd[0]").sendVKey(0)
            except: pass

    def criar_rc_para_lote_ok(self, lote_de_itens_ok, data_hoje):
        if lote_de_itens_ok.empty: return None, "Lote vazio."
        try:
            self.print_header(f"Criando RC para {len(lote_de_itens_ok)} itens")
            self.session.findById("wnd[0]").maximize()
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/NME51N"
            self.session.findById("wnd[0]").sendVKey(0)
            self.aguardar_sap()
            self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:3327/cmbMEREQ_TOPLINE-BSART").key = "ZRT"
            self.session.findById("wnd[0]").sendVKey(0)
            self.aguardar_sap()
            grid = self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell")
            
            lote = lote_de_itens_ok.reset_index(drop=True)
            for i, item in lote.iterrows():
                if not self.running: return None, "Cancelado."
                
                # Mapeamento (Prioridade para ORIGEM/DESTINO do SAP)
                mat_id = item.get('PN') or item.get('Material ID')
                origem = item.get('ORIGEM') or item.get('Origem Sigla')
                destino = item.get('DESTINO') or item.get('Destino Sigla')
                qtd = str(item.get('QTD') or item.get('Quantidade', '1')).replace(',', '.')
                texto = item.get('TEXTO') or item.get('Logística')

                grid.modifyCell(i, "MATNR", str(mat_id))
                grid.modifyCell(i, "MENGE", qtd)
                grid.modifyCell(i, "RESWK", str(origem))
                grid.modifyCell(i, "EEIND", data_hoje) # Data automática
                grid.modifyCell(i, "EPSTP", "U")
                grid.modifyCell(i, "NAME1", str(destino))
                grid.modifyCell(i, "EKGRP", "P04")
                grid.modifyCell(i, "TXZ01", str(texto))
            
            self.session.findById("wnd[0]").sendVKey(0)
            self.aguardar_sap()
            
            self.print_info("Inserindo Depósitos...")
            for i, item in lote.iterrows():
                if not self.running: return None, "Cancelado."
                
                # Lógica de Mapeamento de Depósito
                origem_key = str(item.get('ORIGEM') or item.get('Origem Sigla')).strip().upper()
                deposito = self.DEPOSITO_MAPPING.get(origem_key, 'AE01') # Default 'AE01' conforme solicitação
                
                grid.setCurrentCell(i, "MATNR")
                self.aguardar_sap()
                self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT16").select()
                self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT16/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsTABREITER1/tabpTRANS").select()
                depot = self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT16/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsTABREITER1/tabpTRANS/ssubSUBBILD1:SAPLXM02:0114/ctxtEBAN-ZZDEP_FORNEC")
                depot.text = str(deposito)
                if i < len(lote) - 1:
                    self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/btn%#AUTOTEXT002").press()
                    self.aguardar_sap()

            self.print_info("Salvando RC...")
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            self.aguardar_sap()

            try:
                self.session.findById("wnd[1]").sendVKey(0) 
                self.aguardar_sap()
            except: pass
            
            msg = self.session.findById("wnd[0]/sbar").text
            match = re.search(r'(\d{10,})', msg)
            if match:
                rc = match.group(0)
                self.print_sucesso(f"RC Criada: {rc}")
                return rc, msg
            else:
                self.print_erro(f"Falha: {msg}")
                return None, msg
        except Exception as e:
            return None, f"Erro criação: {e}"

    # =========================================================================
    #  AUTOMAÇÃO CARGO HEROES
    # =========================================================================
    
    # --- Métodos Auxiliares Restaurados do Script Original ---
    
    def ch_extrair_horarios(self, texto_logistica):
        """Extrai horários (HH:MM) de forma robusta."""
        texto = str(texto_logistica).strip()
        padrao_hora = r'\b(?:[01]?\d|2[0-3]):[0-5]\d\b'
        horarios_encontrados = re.findall(padrao_hora, texto)
        
        if len(horarios_encontrados) >= 2:
            return horarios_encontrados[0], horarios_encontrados[1]
        return None, None

    def ch_calcular_data_hora(self, horario_str):
        try:
            agora = datetime.now()
            parts = horario_str.split(':')
            h, m = int(parts[0]), int(parts[1])
            data_alvo = agora.replace(hour=h, minute=m, second=0, microsecond=0)
            
            if data_alvo < agora - timedelta(hours=6):
                data_alvo += timedelta(days=1)
            return data_alvo.strftime("%d%m%Y%H%M")
        except: return None

    def ch_preencher_data_js(self, driver, wait, xpath_id, texto_numerico, descricao="Data/Hora"):
        """Usa injeção de JS para garantir que a data entre no campo do Angular/React."""
        try:
            self.print_info(f"📅 [DATA] {descricao}: {texto_numerico}")
            if not texto_numerico or len(texto_numerico) < 12: return False
            
            dia, mes, ano = texto_numerico[:2], texto_numerico[2:4], texto_numerico[4:8]
            hora, minuto = texto_numerico[8:10], texto_numerico[10:12]
            data_iso = f"{ano}-{mes}-{dia}T{hora}:{minuto}"
            
            elemento = wait.until(EC.presence_of_element_located((By.XPATH, xpath_id)))
            driver.execute_script("""
                arguments[0].value = arguments[1];
                arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
            """, elemento, data_iso)
            return True
        except Exception as e:
            self.print_erro(f"❌ Erro data '{descricao}': {e}")
            return False

    def start_ch_automation(self):
        if self.running: return
        
        email = self.ch_email_var.get()
        senha = self.ch_pass_var.get()
        if not email or not senha:
            messagebox.showerror("Erro", "Configure Email e Senha do Cargo Heroes.")
            return

        self.running = True
        self.toggle_buttons(False)
        self.status_var.set("Executando Automação Cargo Heroes...")
        
        self.log_area.config(state=tk.NORMAL)
        self.log_area.delete(1.0, tk.END)
        self.log_area.config(state=tk.DISABLED)
        
        self.setup_log_redirector()
        
        thread = threading.Thread(target=self.run_ch_automation, daemon=True)
        thread.start()

    def run_ch_automation(self):
        driver = None
        try:
            self.print_header("Iniciando Cargo Heroes Updater")
            
            opts = Options()
            opts.add_argument("--start-maximized")
            opts.add_argument("--no-sandbox")
            opts.add_argument("--disable-dev-shm-usage")
            opts.add_experimental_option("excludeSwitches", ["enable-automation"])
            opts.add_experimental_option('useAutomationExtension', False)
            
            self.print_info("Abrindo navegador Chrome...")
            driver = webdriver.Chrome(options=opts)
            driver.set_page_load_timeout(60)
            wait = WebDriverWait(driver, 20)

            # Conectar Planilha
            self.print_info("Conectando planilhas...")
            try:
                scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
                creds = ServiceAccountCredentials.from_json_keyfile_name(self.google_creds_var.get(), scope)
                client = gspread.authorize(creds)
                planilha = client.open(self.google_sheet_var.get())
                aba_normal = planilha.worksheet(self.google_tab_var.get())
                try:
                    aba_map = planilha.worksheet('MAPEAMENTO')
                except:
                    self.print_aviso("Aba 'MAPEAMENTO' não encontrada.")
                    aba_map = None
            except Exception as e:
                self.print_erro(f"Erro ao conectar Google Sheets: {e}")
                return

            # Login CH
            self.print_info("Acessando Cargo Heroes...")
            driver.get("https://cargo-heroes.appslatam.com/#/login")
            
            if self.ch_realizar_login(driver, wait, self.ch_email_var.get(), self.ch_pass_var.get()):
                if self.ch_navegar_detalhe(driver, wait):
                    # Processo Normal
                    if self.running:
                        self.ch_processar_normal(driver, wait, aba_normal)
                    
                    # Processo Mapeamento
                    if self.running and aba_map:
                        self.ch_processar_mapeamento(driver, wait, aba_map)
                        
                    self.print_sucesso("Processo CH Concluído!")
            else:
                self.print_erro("Falha no Login CH.")
                
        except Exception as e:
            self.print_erro(f"Erro Fatal CH: {e}")
            self.ch_tirar_print(driver, "CRASH_GERAL")
        finally:
            if driver:
                try: driver.quit()
                except: pass
            self.finalize_automation()

    def ch_tirar_print(self, driver, nome):
        try:
            if driver:
                path = os.path.join(self.logs_print_path, f"{datetime.now().strftime('%H%M%S')}_{nome}.png")
                driver.save_screenshot(path)
                self.print_info(f"Print salvo: {os.path.basename(path)}")
        except: pass

    def ch_acao(self, driver, wait, xpath, acao="clicar", texto=None, desc=""):
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
        except Exception as e:
            self.print_erro(f"Falha ação '{desc}': {e}")
            return False

    def ch_realizar_login(self, driver, wait, email, senha):
        self.print_info("🔑 --- Iniciando Login ---")
        janela_principal = driver.current_window_handle
        try:
            if "login" not in driver.current_url and ("app" in driver.current_url or "logistics" in driver.current_url):
                self.print_sucesso("✔ Já logado.")
                return True

            self.print_info("1. Tentando clicar no botão Google...")
            wait_fast = WebDriverWait(driver, 5) 
            try: wait_fast.until(EC.invisibility_of_element_located((By.XPATH, "//img[contains(@src, 'loading')]")))
            except: pass

            try:
                iframe = wait.until(EC.presence_of_element_located((By.XPATH, "//iframe[contains(@src, 'gsi/button')]")))
                driver.switch_to.frame(iframe)
                botao_google = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@role='button'] | //span[contains(text(), 'Sign in')]")))
                driver.execute_script("arguments[0].click();", botao_google)
                driver.switch_to.default_content()
            except Exception as e:
                self.print_info(f"Botão Google (iframe) não encontrado ou erro: {e}")
        except Exception: pass

        try:
            try: WebDriverWait(driver, 5).until(EC.number_of_windows_to_be(2))
            except: 
                if "login" not in driver.current_url: return True
            
            janelas = driver.window_handles
            if len(janelas) > 1:
                janela_popup = [j for j in janelas if j != janela_principal][0]
                driver.switch_to.window(janela_popup)
                try:
                    inp_email = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//input[@id='identifierId']")))
                    inp_email.clear(); inp_email.send_keys(email)
                    driver.find_element(By.ID, "identifierNext").click()
                except: pass
                try:
                    time.sleep(1)
                    btns = driver.find_elements(By.XPATH, "//*[contains(text(), 'Continuar')]")
                    if btns and btns[0].is_displayed(): btns[0].click()
                except: pass
                try:
                    inp_email_ms = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "i0116")))
                    inp_email_ms.clear(); inp_email_ms.send_keys(email)
                    driver.find_element(By.ID, "idSIButton9").click()
                    inp_senha = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "i0118")))
                    inp_senha.clear(); inp_senha.send_keys(senha)
                    time.sleep(0.5)
                    driver.find_element(By.ID, "idSIButton9").click()
                    try: WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
                    except: pass
                except Exception: pass
        except Exception as e: 
            self.print_erro(f"Erro no fluxo de popup: {e}")

        self.print_info("⏳ Validando acesso...")
        driver.switch_to.window(janela_principal)
        inicio = time.time()
        while time.time() - inicio < 60 and self.running:
            try:
                url_atual = driver.current_url
                if "/login" not in url_atual and ("app" in url_atual or "logistics" in url_atual or "#/home" in url_atual):
                    return True
                if len(driver.find_elements(By.XPATH, "//mat-icon")) > 0:
                    return True
                time.sleep(1)
            except: time.sleep(1)
        return False

    def ch_navegar_detalhe(self, driver, wait):
        self.print_info("Navegando para Detalhe por Linha...")
        try: wait.until(EC.invisibility_of_element_located((By.ID, "loading-bar")))
        except: pass
        time.sleep(2) # Pausa estratégica do script original
        if not self.ch_acao(driver, wait, "//mat-icon[contains(text(), 'menu')]", "clicar", desc="Menu"): return False
        time.sleep(1)
        if not self.ch_acao(driver, wait, "//a[contains(., 'Logística')]", "clicar", desc="Logística"): return False
        time.sleep(1)
        if not self.ch_acao(driver, wait, "//a[contains(., 'Detalhe por linha')]", "clicar", desc="Detalhe"): return False
        time.sleep(3) # Tempo de carregamento da tabela
        return True

    def ch_processar_normal(self, driver, wait, aba):
        self.print_header("Processando Aba Normal")
        dados = aba.get_all_records()
        headers = aba.row_values(1)
        try: col_ok = headers.index("CH OK") + 1
        except: 
            self.print_erro("Coluna 'CH OK' não encontrada.")
            return

        for i, linha in enumerate(dados, start=2):
            if not self.running: break
            
            mat = str(linha.get('Material ID', '')).strip()
            status = str(linha.get('CH OK', '')).strip().upper()
            
            if not mat or status == "OK": continue
            
            self.print_info(f"Processando Linha {i}: {mat}")
            try:
                # 1. Busca
                if not self.ch_busca_material(driver, wait, mat):
                    raise Exception("Material não encontrado ou erro na busca")
                
                # 2. Edição
                self.ch_acao(driver, wait, "//*[@id='dataTable']/tbody/tr/td[1]/a/i", "clicar", desc="Lápis")
                time.sleep(3) # Pausa do script original para abrir modal
                
                # --- PREENCHIMENTO COMPLETO (RESTAURADO) ---
                
                # Requisição
                req = str(linha.get('REQUISIÇÃO', ''))
                self.ch_acao(driver, wait, "//input[@formcontrolname='requirement']", "escrever", texto=req, desc="Requisição")
                
                # Tipo de Atendimento -> Material
                self.ch_acao(driver, wait, "//*[@formcontrolname='typeAtd']", "clicar", desc="Abrir Tipo")
                time.sleep(0.5)
                self.ch_acao(driver, wait, "//mat-option//span[contains(text(), 'Material')]", "clicar", desc="Selecionar Material")
                
                # Modal (Aéreo ou Terrestre)
                modal_txt = "Aéreo" if "Aéreo" in str(linha.get('Tipo de Transporte', '')) else "Terrestre"
                self.ch_acao(driver, wait, "//*[@formcontrolname='modal']", "clicar", desc="Abrir Modal")
                time.sleep(0.5)
                self.ch_acao(driver, wait, f"//mat-option//span[contains(text(), '{modal_txt}')]", "clicar", desc=f"Selecionar {modal_txt}")

                # Origem e Destino
                self.ch_acao(driver, wait, "//input[@formcontrolname='origin']", "escrever", texto=str(linha.get('Origem Sigla')), desc="Origem")
                self.ch_acao(driver, wait, "//input[@formcontrolname='destination']", "escrever", texto=str(linha.get('Destino Sigla')), desc="Destino")
                
                # Logística e Datas (CRÍTICO)
                log_txt = str(linha.get('Logística', ''))
                self.ch_acao(driver, wait, "//input[@formcontrolname='desc']", "escrever", texto=log_txt, desc="Desc Logística")
                
                h_s, h_c = self.ch_extrair_horarios(log_txt)
                if h_s and h_c:
                    ds = self.ch_calcular_data_hora(h_s)
                    dc = self.ch_calcular_data_hora(h_c)
                    if ds and dc:
                        if int(h_c.replace(':','')) < int(h_s.replace(':','')):
                             d_obj = datetime.strptime(dc, "%d%m%Y%H%M") + timedelta(days=1)
                             dc = d_obj.strftime("%d%m%Y%H%M")
                        
                        # Usa a função com injeção de JS restaurada
                        self.ch_preencher_data_js(driver, wait, "//*[@id='dateBoarding']", ds, descricao="Saída")
                        self.ch_preencher_data_js(driver, wait, "//*[@id='dateLanding']", dc, descricao="Chegada")
                
                # Salvar
                if self.ch_acao(driver, wait, "//button[contains(., 'Salve')] | //button[contains(., 'Salvar')]", "clicar", desc="Salvar"):
                    try: 
                        self.print_info("⏳ Confirmando salvamento...")
                        wait.until(EC.visibility_of_element_located((By.XPATH, "//*[contains(text(), 'Solicitação Atendida')]")))
                        aba.update_cell(i, col_ok, "OK")
                        self.print_sucesso(f"Linha {i} OK")
                    except:
                        raise Exception("Confirmação visual falhou")
                
                driver.refresh()
                time.sleep(3)
                self.ch_navegar_detalhe(driver, wait)
                
            except Exception as e:
                self.print_erro(f"Erro linha {i}: {e}")
                aba.update_cell(i, col_ok, "ERRO")
                driver.refresh()
                time.sleep(3)
                self.ch_navegar_detalhe(driver, wait)

    def ch_processar_mapeamento(self, driver, wait, aba):
        self.print_header("Processando Aba Mapeamento")
        dados = aba.get_all_records()
        headers = aba.row_values(1)
        try: col_ok = headers.index("CH OK") + 1
        except: return

        for i, linha in enumerate(dados, start=2):
            if not self.running: break
            
            mat = str(linha.get('Material ID', '')).strip()
            status = str(linha.get('CH OK', '')).strip().upper()
            origem = str(linha.get('ORIGEM', '')).strip().upper()
            
            if not mat or status == "OK": continue
            if "NA BASE" not in origem and "ZERO" not in origem: continue
            
            self.print_info(f"Mapeamento Linha {i}: {mat}")
            try:
                if not self.ch_busca_material(driver, wait, mat):
                    raise Exception("Falha busca mapeamento")

                self.ch_acao(driver, wait, "//*[@id='dataTable']/tbody/tr/td[1]/a/i", "clicar", desc="Lápis")
                time.sleep(3) # Pausa original
                
                if "NA BASE" in origem:
                    self.ch_acao(driver, wait, "//button[contains(., 'Mtl na Base')]", "clicar")
                elif "ZERO" in origem:
                    self.ch_acao(driver, wait, "//button[contains(., 'Mtl Stk Zero')]", "clicar")
                
                self.ch_acao(driver, wait, "//button[contains(., 'Salve')] | //button[contains(., 'Salvar')]", "clicar")
                
                try: 
                    wait.until(EC.visibility_of_element_located((By.XPATH, "//*[contains(text(), 'Solicitação Atendida')]")))
                    aba.update_cell(i, col_ok, "OK")
                    self.print_sucesso(f"Linha {i} OK")
                except:
                    raise Exception("Confirmação visual falhou")
                
                driver.refresh()
                time.sleep(3)
                self.ch_navegar_detalhe(driver, wait)
            except Exception as e:
                self.print_erro(f"Erro linha {i}: {e}")
                driver.refresh()
                time.sleep(3)
                self.ch_navegar_detalhe(driver, wait)

    def ch_busca_material(self, driver, wait, material):
        # Lógica de limpeza robusta do script original
        input_busca = "//input[@formcontrolname='equipmentCode'] | //input[contains(@data-placeholder, 'Materia')]"
        try:
            el = wait.until(EC.element_to_be_clickable((By.XPATH, input_busca)))
            el.click()
            time.sleep(0.2)
            el.send_keys(Keys.CONTROL + "a")
            time.sleep(0.1)
            el.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)
            el.clear()
            el.send_keys(material)
            time.sleep(1)
            self.ch_acao(driver, wait, "//button[contains(., 'Procurar')]", "clicar")
            time.sleep(2) # Pausa essencial para o grid carregar
            return True
        except: return False

# --- Função Principal ---
def main():
    root = tk.Tk()
    app = SAPAutomationGUI(root)
    root.attributes('-topmost', True)
    root.update()
    root.attributes('-topmost', False)
    root.mainloop()

if __name__ == "__main__":
    main()