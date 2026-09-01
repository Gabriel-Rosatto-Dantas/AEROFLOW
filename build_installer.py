import subprocess, sys, glob, os

# Localiza a pasta de runtime gerada pelo PyArmor (nome contém hash)
runtime_dirs = glob.glob("obf/pyarmor_runtime_*")
if not runtime_dirs:
    sys.exit("ERRO: nenhuma pasta pyarmor_runtime_* encontrada em obf/")

runtime_dir  = runtime_dirs[0]
runtime_name = os.path.basename(runtime_dir)

cmd = [
    sys.executable, "-m", "PyInstaller",
    "obf/APP_UNIFICADO_SAP_CH.py",
    "--name",      "AeroFlow",
    "--onedir",
    "--noconsole",
    "--clean",
    "--icon",      "icone.ico",
    "--add-data",  f"{runtime_dir};{runtime_name}",
    "--add-data",  "icone.ico;.",
]
print("Executando:", " ".join(cmd))
subprocess.run(cmd, check=True)
