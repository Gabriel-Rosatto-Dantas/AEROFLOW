import subprocess, sys

cmd = [
    sys.executable, "-m", "PyInstaller",
    "APP_UNIFICADO_SAP_CH.py",
    "--name",      "AeroFlow",
    "--onedir",
    "--noconsole",
    "--clean",
    "--icon",      "icone.ico",
    "--add-data",  "icone.ico;.",
]
print("Executando:", " ".join(cmd))
subprocess.run(cmd, check=True)
