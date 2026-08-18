import os
import sys
import PyInstaller.__main__

def build():
    # Caminho base do projeto
    project_dir = os.path.dirname(os.path.abspath(__file__))
    main_script = os.path.join(project_dir, 'APP_UNIFICADO_SAP_CH.py')
    
    # Ícone (se existir na mesma pasta)
    icon_path = os.path.join(project_dir, 'icone.ico')
    icon_arg = f'--icon={icon_path}' if os.path.exists(icon_path) else ''
    
    print("Iniciando compilação com PyInstaller...")
    
    args = [
        main_script,
        '--name=Automacao_SAP_CH',
        '--onefile',       # Gera apenas 1 arquivo .exe
        '--noconsole',     # Esconde a janela do terminal (prompt)
        '--clean',         # Limpa cache antes do build
        '--hidden-import=certifi',
        '--hidden-import=keyring',
        '--hidden-import=tenacity',
        '--hidden-import=customtkinter',
        '--hidden-import=pandas',
        '--hidden-import=gspread',
        '--hidden-import=selenium',
        '--hidden-import=win32com',
    ]
    
    if icon_arg:
        args.append(icon_arg)
        # Opcional: Adicionar o icone como resource também
        args.append(f'--add-data={icon_path};.')

    PyInstaller.__main__.run(args)
    print("\nCompilação concluída! O arquivo .exe está na pasta 'dist'.")

if __name__ == "__main__":
    build()
