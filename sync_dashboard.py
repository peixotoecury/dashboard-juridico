"""
Sincronização automática: Excel de prazos trabalhistas → Prazos_e_agendas.xlsx → GitHub
Executa a cada 2h via Task Scheduler.
O dashboard lê o arquivo via URL configurada em localStorage (pc_github_prazos).
URL pública: https://raw.githubusercontent.com/peixotoecury/dashboard-juridico/main/Prazos_e_agendas.xlsx
"""

import pandas as pd
import base64
import requests
import json
import logging
import shutil
import tempfile
import time
from datetime import date, timedelta
from pathlib import Path

# ─── CONFIGURAÇÃO ────────────────────────────────────────────────────────────
EXCEL_PATH = (
    r"C:\Users\ach\OneDrive - Peixoto e Cury Advogados"
    r"\Pastas - Time B\Controladoria"
    r"\Base diária de atualização sócios"
    r"\v_pc_controle_prazos_trabalhista.xlsx"
)

try:
    from config_local import GITHUB_TOKEN
except ImportError:
    import os
    GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN", "")
GITHUB_REPO   = "peixotoecury/dashboard-juridico"
GITHUB_FILE   = "Prazos_e_agendas.xlsx"
GITHUB_BRANCH = "main"

MATRICULAS_EXCLUIR = {"CONTROL", "AUDIENCIA", "BLU", "MSS", "PRAZO", "TAC"}

PASTA_BASE = (
    r"C:\Users\ach\OneDrive - Peixoto e Cury Advogados"
    r"\Pastas - Time B\Controladoria"
    r"\Base diária de atualização sócios"
)

SUPABASE_URL   = "https://kgyeeuynerhpaqlhqwkx.supabase.co"
SUPABASE_KEY   = "sb_publishable_MdUFFFJ5xRqiQysyxlWHFA_yk_LsknZ"
SUPABASE_TABLE = "base_planilha"

XLSX_LOCAL = Path(__file__).parent / "Prazos_e_agendas.xlsx"
LOG_FILE   = Path(__file__).parent / "sync.log"
# ─────────────────────────────────────────────────────────────────────────────

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s  %(levelname)s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)


def atualizar_excel_legal_manager() -> bool:
    """Abre o Excel via COM e dispara Refresh All. Retorna True se conseguiu."""
    import threading
    resultado = [False]

    def _refresh():
        xl = None
        try:
            import win32com.client
            logging.info("Abrindo Excel para atualizar conexão com Legal Manager...")
            xl = win32com.client.Dispatch("Excel.Application")
            xl.Visible = False
            xl.DisplayAlerts = False
            wb = xl.Workbooks.Open(EXCEL_PATH, UpdateLinks=False, ReadOnly=False)
            wb.RefreshAll()
            # Aguarda consultas — máx 3 min via polling
            import pythoncom
            pythoncom.CoInitialize()
            deadline = time.time() + 180
            while time.time() < deadline:
                time.sleep(5)
                try:
                    if not any(c.Refreshing for c in wb.Connections):
                        break
                except Exception:
                    break
            wb.Save()
            wb.Close(SaveChanges=False)
            xl.Quit()
            logging.info("Excel atualizado com dados do Legal Manager.")
            resultado[0] = True
        except Exception as e:
            logging.warning(f"Refresh do Excel falhou: {e}. Usando dados existentes.")
            try:
                if xl: xl.Quit()
            except Exception:
                pass

    t = threading.Thread(target=_refresh, daemon=True)
    t.start()
    t.join(timeout=200)  # máximo 200 segundos
    if t.is_alive():
        logging.warning("Timeout no refresh do Excel (200s). Usando dados existentes.")
    return resultado[0]


def _proximo_dia_util(d: date) -> date:
    prox = d + timedelta(days=1)
    while prox.weekday() >= 5:  # 5=sábado, 6=domingo
        prox += timedelta(days=1)
    return prox


def _dia_util_anterior(d: date) -> date:
    ant = d - timedelta(days=1)
    while ant.weekday() >= 5:
        ant -= timedelta(days=1)
    return ant


def calcular_status(data_val) -> str:
    hoje        = date.today()
    fatal_data  = _dia_util_anterior(hoje)   # último dia útil antes de hoje
    d2_data     = _proximo_dia_util(hoje)    # próximo dia útil após hoje

    if pd.isna(data_val):
        return "Futuro"
    d = data_val.date() if hasattr(data_val, "date") else data_val

    if d == hoje:
        return "D-1"
    elif d == d2_data:
        return "D-2"
    elif d == fatal_data:
        return "Fatal"
    elif d < fatal_data:
        return "Vencido"
    else:
        return "Futuro"


def gerar_xlsx() -> int:
    logging.info("Lendo Excel...")
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp_path = tmp.name
    shutil.copy2(EXCEL_PATH, tmp_path)
    df = pd.read_excel(tmp_path, sheet_name="Andamento", dtype=str,
                       parse_dates=["data"])
    Path(tmp_path).unlink(missing_ok=True)

    # Filtros
    df = df[df["area"].str.strip().str.upper() == "TRABALHISTA"]
    df = df[~df["matricula"].str.strip().str.upper().isin(MATRICULAS_EXCLUIR)]

    # Calcula status com base na data
    df["status"] = df["data"].apply(calcular_status)

    # Seleciona colunas que o dashboard usa (inclui apenas as que existem na planilha)
    colunas_desejadas = ["data", "status", "situacao_pasta", "responsavel", "grupo_processo",
                         "cliente", "tipo_agenda", "descricao_tarefa",
                         "numero_processo", "parte_contraria", "importante", "matricula"]
    colunas_existentes = [c for c in colunas_desejadas if c in df.columns]
    resultado = df[colunas_existentes].copy()

    # Ordena por prioridade: Fatal → D-1 → D-2 → Vencido → Futuro, depois por data
    ordem_status = {"Fatal": 0, "D-1": 1, "D-2": 2, "Vencido": 3, "Futuro": 4}
    resultado["_ordem"] = resultado["status"].map(ordem_status).fillna(9)
    resultado = resultado.sort_values(["_ordem", "data"]).drop(columns=["_ordem"])

    resultado.to_excel(XLSX_LOCAL, index=False, sheet_name="Andamento")
    logging.info(f"XLSX gerado: {len(resultado)} registros → {XLSX_LOCAL}")
    return resultado


def gerar_audiencias() -> pd.DataFrame:
    """Lê a Pauta de Audiências mais recente e retorna df no formato do dashboard."""
    import glob, os
    padrao = os.path.join(PASTA_BASE, "Pauta de Audiências*.xlsx")
    arquivos = sorted(glob.glob(padrao), key=os.path.getmtime, reverse=True)
    if not arquivos:
        logging.warning("Pauta de Audiências não encontrada.")
        return pd.DataFrame()

    path = arquivos[0]
    logging.info(f"Lendo pauta: {os.path.basename(path)}")
    df = pd.read_excel(path, sheet_name=0, header=1, parse_dates=["Data"])

    df = df.rename(columns={
        "Tipo Agendamento":                        "tipo_agenda",
        "Data":                                    "data",
        "Advogado Responsável Audiência/SO":       "responsavel",
        "Equipe responsável":                      "grupo_processo",
        "Processo":                                "numero_processo",
        "Parte Principal":                         "parte_principal",
        "Parte Contrária":                         "parte_contraria",
        "Código P&C":                              "codigo_externo",
        "Horário":                                 "horario",
        "Fase":                                    "fase",
        "Ação":                                    "acao",
        "Comarca":                                 "comarca",
    })

    df["cliente"]         = df["parte_principal"]
    df["descricao_tarefa"] = df["tipo_agenda"]
    df["matricula"]       = ""
    df["status"]          = df["data"].apply(calcular_status)

    colunas = ["data", "status", "responsavel", "grupo_processo", "cliente",
               "tipo_agenda", "descricao_tarefa", "numero_processo",
               "parte_contraria", "parte_principal", "horario"]
    colunas_existentes = [c for c in colunas if c in df.columns]
    return df[colunas_existentes].copy()


def atualizar_supabase(df: pd.DataFrame):
    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json",
        "Prefer": "return=minimal",
    }
    base = f"{SUPABASE_URL}/rest/v1/{SUPABASE_TABLE}"

    # Sem delete — usamos upsert (IDs fixos para prazos/agendas, preserva audiências)
    logging.info("Preparando upsert no Supabase...")

    # Prepara registros — serializa datas e NaN para JSON puro
    def limpar(v):
        if pd.isna(v) if not isinstance(v, (list, dict)) else False:
            return None
        if hasattr(v, "strftime"):
            return v.strftime("%Y-%m-%dT12:00:00")  # noon evita deslocamento UTC→Brasil
        return v

    # Prazos/Agendas: IDs 100000+
    registros = []
    for i, row in enumerate(df.to_dict(orient="records"), start=100000):
        registros.append({"id": i, "dados": {k: limpar(v) for k, v in row.items()}})

    # Insere em lotes de 500
    LOTE = 500
    total = len(registros)
    for i in range(0, total, LOTE):
        lote = registros[i:i + LOTE]
        h = {**headers, "Prefer": "resolution=merge-duplicates,return=minimal"}
        r = requests.post(base, headers=h, data=json.dumps(lote))
        if r.status_code not in (200, 201):
            raise RuntimeError(f"Falha ao inserir lote {i//LOTE+1}: {r.status_code} {r.text[:200]}")
        logging.info(f"Supabase: inseridos {min(i+LOTE, total)}/{total}")

    logging.info(f"Supabase: {total} prazos/agendas inseridos.")

    # Audiências: IDs 200000+
    df_aud = gerar_audiencias()
    if not df_aud.empty:
        # Apaga audiências anteriores
        requests.delete(base, headers=headers, params={"id": "gte.200000"})
        aud_registros = []
        for i, row in enumerate(df_aud.to_dict(orient="records"), start=200000):
            aud_registros.append({"id": i, "dados": {k: limpar(v) for k, v in row.items()}})
        for i in range(0, len(aud_registros), LOTE):
            lote = aud_registros[i:i + LOTE]
            h = {**headers, "Prefer": "resolution=merge-duplicates,return=minimal"}
            r = requests.post(base, headers=h, data=json.dumps(lote))
            if r.status_code not in (200, 201):
                logging.warning(f"Audiências lote {i//LOTE+1}: {r.status_code} {r.text[:100]}")
        logging.info(f"Supabase: {len(df_aud)} audiências inseridas.")
    total_geral = total + len(df_aud)
    logging.info(f"Supabase atualizado: {total_geral} registros totais.")


def _push_arquivo_github(nome_arquivo: str, conteudo_bytes: bytes, mensagem: str):
    """Envia um arquivo para o GitHub via API (cria ou atualiza)."""
    api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{nome_arquivo}"
    headers = {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3+json",
    }
    r = requests.get(api_url, headers=headers, params={"ref": GITHUB_BRANCH})
    sha = r.json().get("sha") if r.status_code == 200 else None

    payload = {
        "message": mensagem,
        "content": base64.b64encode(conteudo_bytes).decode(),
        "branch": GITHUB_BRANCH,
    }
    if sha:
        payload["sha"] = sha

    r = requests.put(api_url, headers=headers, data=json.dumps(payload))
    r.raise_for_status()
    logging.info(f"Push GitHub: {nome_arquivo} publicado.")


def push_github():
    _push_arquivo_github(
        GITHUB_FILE,
        XLSX_LOCAL.read_bytes(),
        f"Atualização automática prazos — {date.today()}",
    )


def push_dados_json(df_prazos: "pd.DataFrame", df_aud: "pd.DataFrame"):
    """Gera dados.json com prazos + audiências e faz push para o GitHub.
    O dashboard lê este JSON diretamente, sem passar pelo Supabase."""

    def limpar(v):
        try:
            if pd.isna(v):
                return None
        except (TypeError, ValueError):
            pass
        if hasattr(v, "strftime"):
            return v.strftime("%Y-%m-%dT12:00:00")
        return v

    def df_para_registros(df: "pd.DataFrame") -> list:
        df = df.copy()
        df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]
        return [{k: limpar(v) for k, v in row.items()} for row in df.to_dict(orient="records")]

    registros = df_para_registros(df_prazos)
    if not df_aud.empty:
        registros += df_para_registros(df_aud)

    # Inclui timestamp para o dashboard detectar mudança
    payload = {
        "atualizado_em": date.today().isoformat(),
        "total": len(registros),
        "dados": registros,
    }
    json_bytes = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    _push_arquivo_github(
        "dados.json",
        json_bytes,
        f"dados: sync automático {date.today()} — {len(registros)} registros",
    )
    logging.info(f"dados.json publicado: {len(registros)} registros.")


def main():
    logging.info("=== Início da sincronização ===")
    try:
        atualizar_excel_legal_manager()
        df = gerar_xlsx()
        df_aud = gerar_audiencias()
        push_github()
        push_dados_json(df, df_aud)
        total = len(df)
        logging.info(f"=== Concluído. {total} prazos + {len(df_aud)} audiências publicados via GitHub. ===")
        print(f"OK — {total} prazos e {len(df_aud)} audiências publicados no dashboard.")
    except Exception as e:
        logging.error(f"ERRO: {e}", exc_info=True)
        print(f"ERRO: {e}")
        raise


if __name__ == "__main__":
    main()
