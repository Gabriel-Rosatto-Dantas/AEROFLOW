import subprocess, sys

# PyArmor 7 gera obf/pytransform/ (pasta fixa, sem hash no nome)
cmd = [
    sys.executable, "-m", "PyInstaller",
    "obf/APP_UNIFICADO_SAP_CH.py",
    "--name",      "AeroFlow",
    "--onedir",
    "--noconsole",
    "--clean",
    "--icon",      "icone.ico",
    "--add-data",  "obf/pytransform;pytransform",
    "--add-data",  "icone.ico;.",
]
print("Executando:", " ".join(cmd))
subprocess.run(cmd, check=True)
