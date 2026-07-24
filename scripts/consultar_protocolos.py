"""
Consulta automática de protocolos no site do Registro de Imóveis de Indaial.

Fluxo:
1. Lê protocolos.json (lista de protocolo/senha/referencia).
2. Para cada protocolo, abre o site oficial (riindaial.com.br) com um navegador
   headless (Playwright), preenche código/senha e extrai o resultado.
3. Baixa os PDFs de exigência (quando existirem).
4. Gera uma planilha .xlsx com o relatório (mesmo layout do painel manual).
5. Atualiza uma aba do Google Sheets com os mesmos dados.
6. Envia um e-mail com o relatório e os PDFs anexados.

Variáveis de ambiente esperadas (configuradas como GitHub Secrets):
- GOOGLE_SERVICE_ACCOUNT_JSON : conteúdo do JSON da conta de serviço do Google
- GOOGLE_SHEET_ID             : ID da planilha do Google Sheets
- EMAIL_HOST, EMAIL_PORT      : servidor SMTP (ex: smtp.gmail.com, 587)
- EMAIL_USER, EMAIL_PASS      : credenciais do e-mail remetente (senha de app)
- EMAIL_TO                    : destinatário(s), separados por vírgula
"""

import json
import os
import re
import smtplib
import ssl
from datetime import datetime, date
from email.message import EmailMessage
from pathlib import Path

from playwright.sync_api import sync_playwright
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent.parent
PROTOCOLOS_FILE = BASE_DIR / "protocolos.json"
OUTPUT_DIR = BASE_DIR / "output"
PDFS_DIR = OUTPUT_DIR / "pdfs"
REPORT_PATH = OUTPUT_DIR / "relatorio_protocolos_ri.xlsx"

SITE_URL = "https://riindaial.com.br/acompanhar-solicitacao/"

NAVY = "1A2B45"
RED_BG = "FBEAE8"
AMBER_BG = "FAF1DE"
GREEN_BG = "EAF2EC"
WHITE = "FFFFFF"


def parse_br_date(s):
    if not s:
        return None
    m = re.search(r"(\d{2})/(\d{2})/(\d{4})", s)
    if not m:
        return None
    d, mth, y = m.groups()
    return date(int(y), int(mth), int(d))


def consultar_um_protocolo(page, protocolo, senha):
    """Preenche o formulário do site e extrai os dados do resultado."""
    page.goto(SITE_URL, wait_until="networkidle")

    # Aceitar cookies se aparecer
    try:
        page.click("text=Aceitar", timeout=3000)
    except Exception:
        pass

    # Selecionar Protocolo / Normal (garante estado limpo)
    page.check("input[value='PROTOCOLO']")
    page.check("input[value='NORMAL']")

    codigo_input = page.locator("#ca-codigo")
    senha_input = page.locator("#ca-senha")
    codigo_input.fill("")
    codigo_input.fill(protocolo)
    senha_input.fill("")
    senha_input.fill(senha)

    page.click("button[type='submit']")
    page.wait_for_timeout(2500)

    resultado_texto = page.locator("text=Resultado").locator("..").inner_text()

    def campo(rotulo):
        m = re.search(rf"{rotulo}\s*\n?\s*([^\n]+)", resultado_texto, re.IGNORECASE)
        return m.group(1).strip() if m else None

    if "não encontrad" in resultado_texto.lower() or "não encontrada" in resultado_texto.lower():
        return {"protocolo": protocolo, "erro": "Protocolo não encontrado. Confira código/senha."}

    dados = {
        "protocolo": protocolo,
        "situacao": campo("SITUAÇÃO \\(ETAPA ATUAL\\)"),
        "cadastro": campo("CADASTRO"),
        "prazoQualificacao": campo("PRAZO QUALIFICAÇÃO"),
        "vencimentoPrenotacao": campo("VENCIMENTO PRENOTAÇÃO"),
        "interessado": campo("INTERESSADO"),
        "solicitante": campo("SOLICITANTE"),
        "natureza": campo("NATUREZA"),
        "total": campo("TOTAL"),
        "exigencias": [],
        "exigencia_links": [],
    }

    # Links das exigências (abrem o Asgard); cada <a> na lista de exigências
    exigencia_links = page.locator("a", has_text="Exigência em")
    count = exigencia_links.count()
    for i in range(count):
        el = exigencia_links.nth(i)
        texto = el.inner_text().strip()
        href = el.get_attribute("href")
        dados["exigencias"].append(texto)
        if href:
            dados["exigencia_links"].append(href)

    return dados


def baixar_pdfs_exigencia(page, protocolo, links):
    """Baixa cada PDF de exigência para output/pdfs/<protocolo>/"""
    pasta = PDFS_DIR / protocolo
    pasta.mkdir(parents=True, exist_ok=True)
    caminhos = []
    for i, link in enumerate(links, start=1):
        try:
            with page.expect_download(timeout=15000) as download_info:
                page.goto(link)
            download = download_info.value
            destino = pasta / f"exigencia_{i}.pdf"
            download.save_as(str(destino))
            caminhos.append(str(destino))
        except Exception as e:
            print(f"  [aviso] Não foi possível baixar PDF {i} do protocolo {protocolo}: {e}")
    return caminhos


def rodar_consultas():
    protocolos_cfg = json.loads(PROTOCOLOS_FILE.read_text(encoding="utf-8"))
    resultados = []

    with sync_playwright() as p:
        browser = p.chromium.launch()
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        for item in protocolos_cfg:
            protocolo = item["protocolo"]
            senha = item["senha"]
            referencia = item.get("referencia", "")
            print(f"Consultando protocolo {protocolo}...")
            dados = consultar_um_protocolo(page, protocolo, senha)
            dados["referencia"] = referencia

            if dados.get("exigencia_links"):
                dados["pdfs"] = baixar_pdfs_exigencia(page, protocolo, dados["exigencia_links"])
            else:
                dados["pdfs"] = []

            resultados.append(dados)

        browser.close()

    return resultados


def gerar_planilha(resultados):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Protocolos RI"

    ws.merge_cells("A1:L1")
    ws["A1"] = "Painel de Protocolos — Registro de Imóveis"
    ws["A1"].font = Font(name="Arial", size=16, bold=True, color=WHITE)
    ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:L2")
    ws["A2"] = f"Atualizado automaticamente em {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws["A2"].font = Font(name="Arial", size=10, italic=True, color=NAVY)

    headers = ["Protocolo", "Referência", "Interessado", "Solicitante", "Natureza",
               "Situação Atual", "Prazo Qualificação", "Vencimento Prenotação",
               "Total (R$)", "Qtde Exigências", "Links de Exigência", "Última Consulta"]
    header_row = 4
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=header_row, column=i, value=h)
        c.font = Font(name="Arial", size=10, bold=True, color=WHITE)
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[header_row].height = 32

    today = date.today()
    row = header_row + 1
    for d in resultados:
        if d.get("erro"):
            ws.cell(row=row, column=1, value=d["protocolo"])
            ws.cell(row=row, column=2, value=d.get("referencia", ""))
            ws.cell(row=row, column=6, value=d["erro"]).font = Font(color="A83C34", bold=True)
            row += 1
            continue

        prazo_dt = parse_br_date(d.get("prazoQualificacao"))
        vencido = prazo_dt is not None and prazo_dt < today
        tem_exigencia = len(d.get("exigencias", [])) > 0
        fill = RED_BG if tem_exigencia else (AMBER_BG if vencido else GREEN_BG)

        valores = [
            d["protocolo"], d.get("referencia", ""), d.get("interessado", ""),
            d.get("solicitante", ""), d.get("natureza", ""), d.get("situacao", ""),
            d.get("prazoQualificacao", ""), d.get("vencimentoPrenotacao", ""),
            d.get("total", ""), len(d.get("exigencias", [])),
            "; ".join(d.get("exigencia_links", [])),
            datetime.now().strftime("%d/%m/%Y %H:%M"),
        ]
        for i, v in enumerate(valores, start=1):
            c = ws.cell(row=row, column=i, value=v)
            c.font = Font(name="Arial", size=10)
            c.fill = PatternFill("solid", fgColor=fill)
            c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        row += 1

    widths = [12, 22, 26, 20, 16, 26, 16, 16, 12, 10, 40, 16]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A5"

    wb.save(REPORT_PATH)
    return REPORT_PATH


def atualizar_google_sheets(resultados):
    creds_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    sheet_id = os.environ.get("GOOGLE_SHEET_ID")
    if not creds_json or not sheet_id:
        print("[aviso] Credenciais do Google Sheets não configuradas — pulando essa etapa.")
        return

    import gspread
    from google.oauth2.service_account import Credentials

    creds_dict = json.loads(creds_json)
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(sheet_id)
    ws = sh.sheet1
    ws.clear()

    headers = ["Protocolo", "Referência", "Interessado", "Solicitante", "Natureza",
               "Situação Atual", "Prazo Qualificação", "Vencimento Prenotação",
               "Total (R$)", "Qtde Exigências", "Links de Exigência", "Última Consulta"]
    linhas = [headers]
    for d in resultados:
        if d.get("erro"):
            linhas.append([d["protocolo"], d.get("referencia", ""), "", "", "", d["erro"], "", "", "", "", "", ""])
            continue
        linhas.append([
            d["protocolo"], d.get("referencia", ""), d.get("interessado", ""),
            d.get("solicitante", ""), d.get("natureza", ""), d.get("situacao", ""),
            d.get("prazoQualificacao", ""), d.get("vencimentoPrenotacao", ""),
            d.get("total", ""), len(d.get("exigencias", [])),
            "; ".join(d.get("exigencia_links", [])),
            datetime.now().strftime("%d/%m/%Y %H:%M"),
        ])
    ws.update(linhas)
    print("Google Sheets atualizado.")


def enviar_email(resultados, anexo_planilha):
    host = os.environ.get("EMAIL_HOST")
    port = int(os.environ.get("EMAIL_PORT", "587"))
    user = os.environ.get("EMAIL_USER")
    senha = os.environ.get("EMAIL_PASS")
    destinatarios = os.environ.get("EMAIL_TO", "")
    if not all([host, user, senha, destinatarios]):
        print("[aviso] Credenciais de e-mail não configuradas — pulando essa etapa.")
        return

    total_exigencias = sum(len(d.get("exigencias", [])) for d in resultados if not d.get("erro"))
    vencidos = sum(
        1 for d in resultados
        if not d.get("erro") and parse_br_date(d.get("prazoQualificacao"))
        and parse_br_date(d.get("prazoQualificacao")) < date.today()
    )

    msg = EmailMessage()
    msg["Subject"] = f"Relatório de Protocolos RI — {datetime.now().strftime('%d/%m/%Y')}"
    msg["From"] = user
    msg["To"] = destinatarios
    corpo = (
        f"Relatório automático de acompanhamento de protocolos no RI Indaial.\n\n"
        f"Total de protocolos monitorados: {len(resultados)}\n"
        f"Com exigência pendente: {total_exigencias}\n"
        f"Com prazo de qualificação vencido: {vencidos}\n\n"
        f"Planilha completa em anexo. PDFs de exigência (quando houver) também anexados.\n"
    )
    msg.set_content(corpo)

    with open(anexo_planilha, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=anexo_planilha.name,
        )

    for d in resultados:
        for caminho in d.get("pdfs", []):
            p = Path(caminho)
            if p.exists():
                with open(p, "rb") as f:
                    msg.add_attachment(
                        f.read(), maintype="application", subtype="pdf",
                        filename=f"{d['protocolo']}_{p.name}",
                    )

    context = ssl.create_default_context()
    with smtplib.SMTP(host, port) as server:
        server.starttls(context=context)
        server.login(user, senha)
        server.send_message(msg)
    print("E-mail enviado.")


if __name__ == "__main__":
    resultados = rodar_consultas()
    planilha = gerar_planilha(resultados)
    atualizar_google_sheets(resultados)
    enviar_email(resultados, planilha)
    print("Concluído.")
