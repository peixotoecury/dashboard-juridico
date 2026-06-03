"""
Monitora Prazos_e_agendas.xlsx e dispara sync_prazos.py automaticamente.
Roda em segundo plano (configurar no Task Scheduler do Windows).
"""
import time, subprocess, sys, logging
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

PASTA    = Path(__file__).parent
EXCEL    = "v_pc_controle_prazos_trabalhista.xlsx"
PYTHON   = sys.executable
SYNC     = PASTA / "sync_prazos.py"
LOG      = PASTA / "watcher_prazos.log"
COOLDOWN = 600  # 10 minutos minimos entre syncs

logging.basicConfig(filename=LOG, level=logging.INFO,
    format="%(asctime)s  %(message)s", datefmt="%Y-%m-%d %H:%M:%S")

ultimo_sync = 0

def rodar_sync(motivo):
    global ultimo_sync
    agora = time.time()
    if agora - ultimo_sync < COOLDOWN:
        return
    ultimo_sync = agora
    logging.info(f"Sync prazos disparado: {motivo}")
    subprocess.Popen([PYTHON, str(SYNC)], creationflags=subprocess.CREATE_NO_WINDOW)

class ExcelHandler(FileSystemEventHandler):
    def on_modified(self, event):
        if Path(event.src_path).name == EXCEL:
            rodar_sync("Excel modificado")
    def on_created(self, event):
        if Path(event.src_path).name == EXCEL:
            rodar_sync("Excel salvo/criado")

if __name__ == "__main__":
    logging.info("=== Watcher prazos iniciado ===")
    observer = Observer()
    observer.schedule(ExcelHandler(), str(PASTA), recursive=False)
    observer.start()
    try:
        while True:
            time.sleep(5)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
