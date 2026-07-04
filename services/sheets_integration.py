"""
sheets_integration.py
NIT – Integração Google Sheets (Cliente B e futuros clientes)
"""

import os
import json
import logging
import time
from datetime import datetime, date
from typing import Optional

import firebase_admin
from firebase_admin import credentials, db as rtdb
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ---------------------------------------------
# Logging estruturado
# ---------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    datefmt="%Y-%m-%dT%H:%M:%S",
)
log = logging.getLogger("nit.sheets")

# ---------------------------------------------
# Google Sheets – autenticação
# ---------------------------------------------
_sheets_service = None

def get_sheets_service():
    """Retorna o serviço autenticado do Google Sheets (singleton)."""
    global _sheets_service
    if _sheets_service:
        return _sheets_service

    creds_raw = os.environ.get("GOOGLE_SHEETS_CREDENTIALS")
    if not creds_raw:
        raise EnvironmentError("GOOGLE_SHEETS_CREDENTIALS não configurado.")

    creds_dict = json.loads(creds_raw)
    creds = Credentials.from_service_account_info(
        creds_dict,
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    _sheets_service = build("sheets", "v4", credentials=creds, cache_discovery=False)
    return _sheets_service


# ---------------------------------------------
# Helpers de data/hora
# ---------------------------------------------

def _fmt_date(val: Optional[str]) -> str:
    """Normaliza datas para DD/MM/AAAA (aceita AAAA-MM-DD ou DD/MM/AAAA)."""
    if not val:
        return ""
    val = str(val).strip()
    if "-" in val:
        try:
            d = datetime.strptime(val[:10], "%Y-%m-%d")
            return d.strftime("%d/%m/%Y")
        except ValueError:
            pass
    return val


def _fmt_time(val: Optional[str]) -> str:
    """Retorna hora no formato HH:MM."""
    if not val:
        return ""
    return str(val).strip()[:5]


def calcular_tempo_atendimento(
    data_inicio: str,
    hora_inicio: str,
    data_fim: str,
    hora_fim: str,
) -> str:
    """
    Calcula duração entre início e fim. Retorna HH:MM:SS ou ''.
    Aceita datas nos formatos YYYY-MM-DD e DD/MM/YYYY.
    """
    def _parse(data: str, hora: str) -> datetime:
        data = data.strip()[:10]
        hora = hora.strip()[:5]
        for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
            try:
                return datetime.strptime(f"{data} {hora}", f"{fmt} %H:%M")
            except ValueError:
                continue
        raise ValueError(f"Formato de data não reconhecido: {data!r}")

    try:
        dt_in  = _parse(data_inicio, hora_inicio)
        dt_fim = _parse(data_fim,    hora_fim)
        delta  = dt_fim - dt_in
        if delta.total_seconds() < 0:
            return ""
        total_sec = int(delta.total_seconds())
        hh = total_sec // 3600
        mm = (total_sec % 3600) // 60
        ss = total_sec % 60
        return f"{hh:02d}:{mm:02d}:{ss:02d}"
    except Exception as exc:
        log.warning("calcular_tempo_atendimento falhou: %s", exc)
        return ""


def _concat_data_hora(data: str, hora: str) -> str:
    """Monta concatenação DATA_HORA no padrão DD/MM/AAAA HH:MM."""
    d = _fmt_date(data)
    h = _fmt_time(hora)
    if d and h:
        return f"{d} {h}"
    return d or h


# ---------------------------------------------
# Mapeamento Firebase → linha da planilha
# ---------------------------------------------

def _montar_linha_nova(dados: dict, cfg: dict, id_ocorrencia: str) -> list:
    """
    Converte um dict do Firebase em lista de valores ordenada por coluna.
    Ordem: A→R (incluindo coluna oculta R = ID_OCORRENCIA).
    """
    mp = cfg["mapeamento_firebase"]

    def fb(campo_planilha: str) -> str:
        chave = mp.get(campo_planilha, "")
        return str(dados.get(chave, "")).strip() if chave else ""

    inicio_raw = fb("data_inicio")
    if " " in inicio_raw:
        data_inicio_raw, hora_inicio_raw = inicio_raw.split(" ", 1)
    else:
        data_inicio_raw = inicio_raw
        hora_inicio_raw = fb("hora_inicio")

    pl_map = {"norm": "NORMALIZADO", "atend": "EM ATENDIMENTO", "aguard": "AGUARDANDO", "sn": "SEM NECESSIDADE"}
    status_atual_raw = fb("status_atual")
    status_atual = pl_map.get(status_atual_raw, status_atual_raw.upper() if status_atual_raw else "")

    row = [
        fb("scn"),                                              # A – SCN
        fb("localizacao"),                                      # B – LOCALIZAÇÃO
        _fmt_date(data_inicio_raw),                             # C – DATA_PLANTAO
        fb("plantonista"),                                      # D – PLANTONISTA
        fb("status_falha").upper() if fb("status_falha") else "",  # E – STATUS_FALHA
        fb("causa").upper() if fb("causa") else "",             # F – CAUSA
        _fmt_date(data_inicio_raw),                             # G – DATA_INICIO
        _fmt_time(hora_inicio_raw),                             # H – HORA_INICIO
        _concat_data_hora(data_inicio_raw, hora_inicio_raw),    # I – DATA_HORA_IN
        "",                                                     # J – DATA_FIM
        "",                                                     # K – HORA_FIM
        "",                                                     # L – DATA_HORA_FIM
        status_atual or "AGUARDANDO",                           # M – STATUS_ATUAL
        dados.get("colunaN") or {"vl": "VIA LIVRE", "amc": "AMC"}.get(
            str(dados.get("sub", "")).strip().lower(), fb("operando_cruzamento")
        ),                                                      # N – OPERANDO_CRUZAMENTO
        "",                                                     # O – TEMPO DE ATENDIMENTO
        fb("bairro"),                                           # P – BAIRRO
        fb("observacoes"),                                      # Q – OBSERVAÇÕES
        id_ocorrencia,                                          # R – ID_OCORRENCIA (oculta)
    ]
    return row


# ---------------------------------------------
# Retry exponencial
# ---------------------------------------------

def _retry(fn, max_tries=3, base_delay=2):
    """Executa fn com retry exponencial."""
    for attempt in range(1, max_tries + 1):
        try:
            return fn()
        except HttpError as exc:
            if exc.resp.status in (429, 500, 503) and attempt < max_tries:
                delay = base_delay ** attempt
                log.warning("HTTP %s – tentativa %d/%d, aguardando %ds",
                            exc.resp.status, attempt, max_tries, delay)
                time.sleep(delay)
            else:
                raise


# ---------------------------------------------
# Cache em memória: coluna R por planilha
# ---------------------------------------------
_cache_coluna_r: dict = {}

def _invalidar_cache_linha(spreadsheet_id: str, sheet_name: str):
    _cache_coluna_r.pop((spreadsheet_id, sheet_name), None)


# ---------------------------------------------
# Localizar linha pelo ID_OCORRENCIA (coluna R)
# ---------------------------------------------

def _encontrar_linha(service, spreadsheet_id: str, sheet_name: str, id_ocorrencia: str) -> Optional[int]:
    """
    Percorre a coluna R para encontrar o índice (1-based) da linha com o ID.
    Se houver duplicatas (herança diária gera múltiplas linhas com mesmo ID),
    retorna a ÚLTIMA ocorrência — que é sempre a mais recente.
    """
    cache_key = (spreadsheet_id, sheet_name)

    if cache_key not in _cache_coluna_r:
        range_r = f"'{sheet_name}'!R:R"

        def _read():
            return (
                service.spreadsheets()
                .values()
                .get(spreadsheetId=spreadsheet_id, range=range_r)
                .execute()
            )

        result = _retry(_read)
        values = result.get("values", [])
        # Dict comprehension: última ocorrência de cada ID vence (herança diária)
        _cache_coluna_r[cache_key] = {
            cell[0]: i + 1
            for i, cell in enumerate(values)
            if cell and cell[0]
        }
        log.debug("cache coluna R carregado | planilha=%s | entradas=%d",
                  sheet_name, len(_cache_coluna_r[cache_key]))

    return _cache_coluna_r[cache_key].get(str(id_ocorrencia))


# ---------------------------------------------
# API pública
# ---------------------------------------------

def append_nova_ocorrencia(
    cliente_config: dict,
    dados_ocorrencia: dict,
    id_ocorrencia: str,
    dry_run: bool = False,
) -> bool:
    """Adiciona uma nova linha ao final da planilha."""
    spreadsheet_id = cliente_config["spreadsheet_id"]
    sheet_name     = cliente_config["sheet_name"]
    row            = _montar_linha_nova(dados_ocorrencia, cliente_config, id_ocorrencia)

    log.info("APPEND | id=%s | scn=%s | dry_run=%s", id_ocorrencia, row[0], dry_run)

    if dry_run:
        log.info("DRY_RUN – linha que seria inserida: %s", row)
        return True

    service = get_sheets_service()
    range_a1 = f"'{sheet_name}'!A:R"

    def _write():
        return (
            service.spreadsheets()
            .values()
            .append(
                spreadsheetId=spreadsheet_id,
                range=range_a1,
                valueInputOption="RAW",
                insertDataOption="INSERT_ROWS",
                body={"values": [row]},
            )
            .execute()
        )

    try:
        result = _retry(_write)
        log.info("APPEND OK | updates=%s", result.get("updates", {}).get("updatedRows"))
        _invalidar_cache_linha(spreadsheet_id, sheet_name)
        return True
    except Exception as exc:
        log.error("APPEND FAIL | id=%s | erro=%s", id_ocorrencia, exc)
        return False


def append_heranca_diaria(
    cliente_config: dict,
    dados_ocorrencia: dict,
    id_ocorrencia: str,
    data_plantao: str = "",
    dry_run: bool = False,
) -> bool:
    """
    Registra herança diária de uma ocorrência pendente na planilha.

    Quando uma ocorrência atravessa a virada do dia sem ser normalizada,
    o próximo plantão a herda. Isso gera uma nova linha na planilha com:
      - DATA_PLANTAO (col C) = data do novo plantão
      - Início original preservado (cols G, H, I)
      - Status = EM ATENDIMENTO
      - Fim/tempo zerados (cols J, K, L, O)
      - Mesmo ID_OCORRENCIA (col R) → _encontrar_linha retorna a última linha
        (mais recente), então update_ocorrencia_normalizada atualiza esta.

    Args:
        cliente_config: bloco do clientes.json
        dados_ocorrencia: dict do Firebase (nó completo)
        id_ocorrencia: chave do nó Firebase
        data_plantao: data do novo plantão de referência (DD/MM/AAAA ou AAAA-MM-DD)
        dry_run: se True, apenas loga sem escrever

    Returns:
        True em sucesso, False em falha
    """
    spreadsheet_id = cliente_config["spreadsheet_id"]
    sheet_name     = cliente_config["sheet_name"]

    # Monta linha base e ajusta campos de herança
    row = _montar_linha_nova(dados_ocorrencia, cliente_config, id_ocorrencia)

    # C (índice 2) — DATA_PLANTAO → novo dia de referência
    if data_plantao:
        row[2] = _fmt_date(data_plantao)

    # M (índice 12) — STATUS_ATUAL → sempre EM ATENDIMENTO na herança
    row[12] = "EM ATENDIMENTO"

    # J, K, L (índices 9–11) — fim → vazio (ainda não normalizado)
    row[9]  = ""
    row[10] = ""
    row[11] = ""

    # O (índice 14) — TEMPO ATENDIMENTO → vazio
    row[14] = ""

    log.info("HERANÇA | id=%s | data_plantao=%s | scn=%s | dry_run=%s",
             id_ocorrencia, data_plantao, row[0], dry_run)

    if dry_run:
        log.info("DRY_RUN HERANÇA – linha que seria inserida: %s", row)
        return True

    service  = get_sheets_service()
    range_a1 = f"'{sheet_name}'!A:R"

    def _write():
        return (
            service.spreadsheets()
            .values()
            .append(
                spreadsheetId=spreadsheet_id,
                range=range_a1,
                valueInputOption="RAW",
                insertDataOption="INSERT_ROWS",
                body={"values": [row]},
            )
            .execute()
        )

    try:
        result = _retry(_write)
        log.info("HERANÇA OK | id=%s | updates=%s",
                 id_ocorrencia, result.get("updates", {}).get("updatedRows"))
        # Invalida cache: nova linha altera índices de busca futuros
        _invalidar_cache_linha(spreadsheet_id, sheet_name)
        return True
    except Exception as exc:
        log.error("HERANÇA FAIL | id=%s | erro=%s", id_ocorrencia, exc)
        return False


def update_ocorrencia_normalizada(
    cliente_config: dict,
    id_ocorrencia: str,
    dados_fim: dict,
    dry_run: bool = False,
) -> bool:
    """
    Atualiza as colunas de finalização de uma ocorrência já existente na planilha.
    Se houver múltiplas linhas com o mesmo ID (herança diária), atualiza a última.
    """
    spreadsheet_id = cliente_config["spreadsheet_id"]
    sheet_name     = cliente_config["sheet_name"]

    data_inicio = str(dados_fim.get("data_inicio", "")).strip()
    hora_inicio = str(dados_fim.get("hora_inicio", "")).strip()

    # Firebase grava "inicio" como campo composto "DD/MM/YYYY HH:MM"
    # data_inicio e hora_inicio separados não existem — extrair do campo composto
    if not data_inicio:
        inicio_raw = str(dados_fim.get("inicio", "")).strip()
        if " " in inicio_raw:
            data_inicio, hora_inicio = inicio_raw.split(" ", 1)
        elif inicio_raw:
            data_inicio = inicio_raw
    data_fim    = str(dados_fim.get("data_fim",    "")).strip()
    hora_fim    = str(dados_fim.get("hora_fim",    "")).strip()

    pl_map = {"norm": "NORMALIZADO", "atend": "EM ATENDIMENTO", "aguard": "AGUARDANDO", "sn": "SEM NECESSIDADE"}
    pl_raw = str(dados_fim.get("pl", "")).strip()
    status_atual = pl_map.get(pl_raw, pl_raw.upper() if pl_raw else "NORMALIZADO")

    tempo = calcular_tempo_atendimento(data_inicio, hora_inicio, data_fim, hora_fim)

    log.info("UPDATE | id=%s | data_fim=%s | hora_fim=%s | tempo=%s | dry_run=%s",
             id_ocorrencia, data_fim, hora_fim, tempo, dry_run)

    if dry_run:
        log.info("DRY_RUN – dados que seriam atualizados: data_fim=%s, hora_fim=%s, tempo=%s, status=%s",
                 data_fim, hora_fim, tempo, status_atual)
        return True

    service = get_sheets_service()

    row_num = _encontrar_linha(service, spreadsheet_id, sheet_name, id_ocorrencia)
    if row_num is None:
        log.error("UPDATE FAIL | id=%s | linha não encontrada na planilha", id_ocorrencia)
        return False

    updates = [
        (f"'{sheet_name}'!J{row_num}", _fmt_date(data_fim)),
        (f"'{sheet_name}'!K{row_num}", _fmt_time(hora_fim)),
        (f"'{sheet_name}'!L{row_num}", _concat_data_hora(data_fim, hora_fim)),
        (f"'{sheet_name}'!M{row_num}", status_atual),
        (f"'{sheet_name}'!O{row_num}", tempo),
    ]

    # Coluna N — atualiza com colunaN do histórico da Central da Ocorrência
    coluna_n_valor = str(dados_fim.get("colunaN", "")).strip()
    if coluna_n_valor:
        updates.append((f"'{sheet_name}'!N{row_num}", coluna_n_valor))
        log.info("UPDATE | id=%s | coluna N → %s", id_ocorrencia, coluna_n_valor)

    data_body = {
        "valueInputOption": "USER_ENTERED",
        "data": [
            {"range": rng, "values": [[val]]}
            for rng, val in updates
        ],
    }

    def _write():
        return (
            service.spreadsheets()
            .values()
            .batchUpdate(spreadsheetId=spreadsheet_id, body=data_body)
            .execute()
        )

    try:
        result = _retry(_write)
        log.info("UPDATE OK | id=%s | row=%d | totalUpdated=%s",
                 id_ocorrencia, row_num,
                 result.get("totalUpdatedCells"))
        return True
    except Exception as exc:
        log.error("UPDATE FAIL | id=%s | erro=%s", id_ocorrencia, exc)
        return False


# ---------------------------------------------
# Cursor Firebase
# ---------------------------------------------

def get_cursor(cursor_node: str) -> dict:
    """Lê o cursor de exportação do Firebase."""
    try:
        ref = rtdb.reference(cursor_node)
        val = ref.get() or {}
        return {
            "ultimo_ts":             val.get("ultimo_ts",             0),
            "ultima_atualizacao":    val.get("ultima_atualizacao",    0),
            "ultimo_dataReferencia": val.get("ultimo_dataReferencia", 0),
        }
    except Exception as exc:
        log.error("get_cursor FAIL | node=%s | erro=%s", cursor_node, exc)
        return {"ultimo_ts": 0, "ultima_atualizacao": 0, "ultimo_dataReferencia": 0}


def update_cursor(cursor_node: str, ultimo_ts: Optional[int] = None,
                  ultima_atualizacao: Optional[int] = None,
                  ultimo_dataReferencia: Optional[int] = None):
    """Atualiza o cursor de exportação no Firebase."""
    try:
        ref = rtdb.reference(cursor_node)
        patch = {}
        if ultimo_ts is not None:
            patch["ultimo_ts"] = ultimo_ts
        if ultima_atualizacao is not None:
            patch["ultima_atualizacao"] = ultima_atualizacao
        if ultimo_dataReferencia is not None:
            patch["ultimo_dataReferencia"] = ultimo_dataReferencia
        if patch:
            ref.update(patch)
            log.info("cursor atualizado | node=%s | patch=%s", cursor_node, patch)
    except Exception as exc:
        log.error("update_cursor FAIL | node=%s | erro=%s", cursor_node, exc)
