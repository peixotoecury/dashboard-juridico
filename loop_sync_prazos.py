"""
loop_sync_prazos.py
Roda sync_prazos.py a cada 30 minutos, sem precisar de admin.
Iniciar junto com o Windows adicionando um atalho na pasta Startup.
"""
import time, subprocess, sys, logging
from pathlib import Path
from datetime import datetime

PASTA    = Path(__file__).parent
SYNC     = PASTA / "sync_prazos.py"
LOG      = PASTA / "loop_sync_prazos.log"
INTERVALO = 20 * 60  # 20 minutos em segundos

logging.basicConfig(filename=LOG, level=logging.INFO,
    format="%(asctime)s  %(message)s", datefmt="%Y-%m-%d %H:%M:%S")

def rodar():
    logging.info("Rodando sync_prazos...")
    resultado = subprocess.run(
        [sys.executable, str(SYNC)],
        capture_output=True, text=True, timeout=120
    )
    if resultado.returncode == 0:
        logging.info(f"OK: {resultado.stdout.strip()}")
    else:
        logging.error(f"ERRO: {resultado.stderr.strip()}")

if __name__ == "__main__":
    logging.info("=== Loop sync prazos iniciado ===")
    rodar()  # roda imediatamente ao iniciar
    while True:
        time.sleep(INTERVALO)
        rodar()
