"""
sync_audiencias.py
Lê a Pauta de Audiências*.xlsx da pasta e sobe para o Supabase prazos_agendas.
Usa tipo_agenda = 'Audiência' e IDs determinísticos (não conflita com prazos do LM).
Pode rodar manualmente ou via Task Scheduler junto com sync_prazos.py.
"""
import hashlib, logging, re
from datetime import date, timedelta, datetime, timezone
from pathlib import Path
import openpyxl
import requests

PASTA    = Path(__file__).parent
LOG      = PASTA / "sync_audiencias.log"

SUPA_URL    = "https://rpibvjcnrseuugpkfmdj.supabase.co"
SUPA_KEY    = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJwaWJ2amNucnNldXVncGtmbWRqIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODE1NTc3MTcsImV4cCI6MjA5NzEzMzcxN30.ecihol8JESMH7cgFSvWKIzp-OwoPRFqdK3aCDwpCeg8"
SUPA_TABELA = "prazos_agendas"

logging.basicConfig(
    filename=LOG, level=logging.INFO,
    format="%(asctime)s  %(levelname)s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)


def dias_uteis_diff(data_dt: date, hoje: date) -> int:
    cal_diff = (data_dt - hoje).days
    if abs(cal_diff) > 14:
        return cal_diff
    if cal_diff == 0:
        return 0
    step = 1 if cal_diff > 0 else -1
    count = 0
    cur = date(hoje.year, hoje.month, hoje.day)
    while cur != data_dt:
        cur += timedelta(days=step)
        if cur.weekday() < 5:
            count += step
    return count


def calc_status(diff) -> str:
    if diff is None:  return "Futuro"
    if diff < -1:     return "Vencido"
    if diff == -1:    return "Fatal"
    if diff == 0:     return "D-1"
    if diff == 1:     return "D-2"
    return "Futuro"


def _gerar_id(numero_processo: str, data_iso: str, responsavel: str) -> int:
    """ID determinístico 52-bit. Prefixo alto para não colidir com IDs do LM."""
    chave = f"AUD|{numero_processo}|{data_iso}|{responsavel}"
    base = int(hashlib.md5(chave.encode()).hexdigest()[:13], 16)
    return base | (1 << 51)


def _buscar_mapa_matricula_nome() -> dict:
    """Busca mapa matricula -> nome completo a partir do Supabase."""
    headers = {"apikey": SUPA_KEY, "Authorization": f"Bearer {SUPA_KEY}",
               "Range-Unit": "items", "Range": "0-4999"}
    r = requests.get(f"{SUPA_URL}/rest/v1/{SUPA_TABELA}", headers=headers,
                     params={"select": "matricula,responsavel", "matricula": "neq."})
    mapa = {}
    for row in r.json():
        mat = (row.get("matricula") or "").strip().upper()
        nome = (row.get("responsavel") or "").strip()
        if mat and nome and mat not in mapa:
            mapa[mat] = nome
    return mapa


def ler_pauta() -> list:
    """Lê o arquivo de pauta mais recente da pasta."""
    arquivos = sorted(PASTA.glob("Pauta de Audiências*.xlsx"),
                      key=lambda p: p.stat().st_mtime, reverse=True)
    arquivos = [a for a in arquivos if not a.name.startswith("~$")]
    if not arquivos:
        raise FileNotFoundError("Nenhum arquivo 'Pauta de Audiências*.xlsx' encontrado na pasta.")

    path = arquivos[0]
    logging.info(f"Lendo pauta: {path.name}")

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    # Detecta se a linha 1 é cabeçalho ou contador:
    # se a linha 0 contiver "Data" ou "Código", são os cabeçalhos direto.
    row0_strs = [str(h).strip().lower() if h else "" for h in rows[0]]
    if any(x in row0_strs for x in ("data", "código p&c", "codigo p&c", "processo")):
        header_row = rows[0]
        data_rows  = rows[1:]
    else:
        header_row = rows[1]
        data_rows  = rows[2:]

    headers = [str(h).strip() if h else "" for h in header_row]

    def col(row, nome):
        for h_idx, h in enumerate(headers):
            if h.lower() == nome.lower():
                try:
                    return row[h_idx]
                except IndexError:
                    return None
        return None

    mapa_matricula = _buscar_mapa_matricula_nome()
    hoje = date.today()
    registros = []

    for row in data_rows:
        if not any(row):
            continue

        # Data
        data_val = col(row, "Data")
        if not data_val:
            continue
        if isinstance(data_val, datetime):
            data_dt = data_val.date()
        elif isinstance(data_val, date):
            data_dt = data_val
        else:
            try:
                data_dt = date.fromisoformat(str(data_val)[:10])
            except Exception:
                continue

        data_iso = data_dt.isoformat()
        diff     = dias_uteis_diff(data_dt, hoje)
        status   = calc_status(diff)

        # Responsável — pode ser matrícula ou nome
        resp_raw = str(col(row, "Advogado Responsável Audiência/SO") or "").strip()
        if not resp_raw or resp_raw.upper() in ("N/A", ""):
            resp_raw = str(col(row, "Advogado Responsável Preparo/Análise") or "").strip()

        # Resolve nome se for matrícula (até 5 chars, sem espaço)
        mat_upper = resp_raw.upper()
        if len(resp_raw) <= 5 and " " not in resp_raw:
            responsavel = mapa_matricula.get(mat_upper, resp_raw)
            matricula   = mat_upper
        else:
            responsavel = resp_raw
            matricula   = ""

        numero_processo = str(col(row, "Processo") or "").strip()
        tipo_aud        = str(col(row, "Tipo Agendamento") or "Audiência").strip()
        parte_principal = str(col(row, "Parte Principal") or "").strip()
        parte_contraria = str(col(row, "Parte Contrária") or "").strip()
        grupo_processo  = str(col(row, "Equipe responsável") or "").strip()
        codigo_externo  = str(col(row, "Código P&C") or "").strip()
        comarca         = str(col(row, "Comarca") or "").strip()
        estado          = str(col(row, "Estado") or "").strip()

        horario_val = col(row, "Horário")
        if horario_val:
            horario_str = str(horario_val)[:5] if len(str(horario_val)) >= 5 else str(horario_val)
        else:
            horario_str = ""

        descricao = tipo_aud
        if horario_str and horario_str not in ("00:00", "0:00:00"):
            descricao = f"{tipo_aud} {horario_str}"
        if comarca:
            descricao += f" — {comarca}{'/'+estado if estado else ''}"

        rid = _gerar_id(numero_processo, data_iso, responsavel)

        if rid in {r["id"] for r in registros}:
            continue  # duplicata dentro da própria pauta

        registros.append({
            "id":               rid,
            "data":             data_iso,
            "status":           status,
            "dias_restantes":   diff,
            "situacao_pasta":   "Andamento",
            "responsavel":      responsavel,
            "grupo_processo":   grupo_processo,
            "cliente":          parte_principal,
            "tipo_agenda":      "Audiência",
            "descricao_tarefa": descricao,
            "numero_processo":  numero_processo,
            "parte_contraria":  parte_contraria,
            "importante":       "",
            "matricula":        matricula,
            "parte_principal":  parte_principal,
            "area":             "TRABALHISTA",
            "codigo_processo":  codigo_externo,
            "codigo_agenda":    "",
            "atualizado_em":    datetime.now(timezone.utc).isoformat(),
        })

    return registros


def upload_supabase(registros: list, run_utc: str):
    headers_up = {
        "apikey":        SUPA_KEY,
        "Authorization": f"Bearer {SUPA_KEY}",
        "Content-Type":  "application/json",
        "Prefer":        "resolution=merge-duplicates,return=minimal",
    }
    headers_del = {
        "apikey":        SUPA_KEY,
        "Authorization": f"Bearer {SUPA_KEY}",
        "Content-Type":  "application/json",
    }

    # Remove audiências antigas (IDs com bit 51 setado) para limpar pauta anterior
    min_aud_id = 1 << 51
    r = requests.delete(
        f"{SUPA_URL}/rest/v1/{SUPA_TABELA}",
        headers=headers_del,
        params={"id": f"gte.{min_aud_id}"},
    )
    logging.info(f"DELETE audiências antigas: {r.status_code}")

    # Upsert em lotes
    LOTE = 200
    erros = 0
    for i in range(0, len(registros), LOTE):
        lote = registros[i:i + LOTE]
        r = requests.post(f"{SUPA_URL}/rest/v1/{SUPA_TABELA}",
                          headers=headers_up, json=lote, timeout=30)
        if r.status_code in (200, 201):
            logging.info(f"Lote {i//LOTE+1}: {len(lote)} audiências ok")
        else:
            erros += 1
            logging.error(f"Lote {i//LOTE+1}: {r.status_code} {r.text[:200]}")

    return erros


def main():
    logging.info("=== Inicio sync audiencias ===")
    try:
        run_utc   = datetime.now(timezone.utc).isoformat()
        registros = ler_pauta()
        logging.info(f"{len(registros)} audiências lidas da pauta.")
        erros = upload_supabase(registros, run_utc)
        hoje_count = sum(1 for r in registros if r["status"] == "D-1")
        logging.info(f"=== Concluido: {len(registros)} audiencias | {hoje_count} hoje | {erros} erros ===")
        print(f"OK — {len(registros)} audiências | {hoje_count} hoje | {erros} erros")
    except Exception as ex:
        logging.error(f"ERRO: {ex}", exc_info=True)
        print(f"ERRO: {ex}")
        raise


if __name__ == "__main__":
    main()
