#!/usr/bin/env python3
"""
Robô RPA — Prefeitura Inteligente: Triagem e Encaminhamento de Solicitações
Disciplina: Robotic Process Automation (Avaliação Substitutiva)
Autor: Lucas de Freitas Pagung — RM 553242

Regras implementadas (conforme Prova Substitutiva RPA):
1. Consulta solicitações não processadas na aba 'Solicitacoes'.
2. Identificação de Risco Imediato (SEMPRE ANTES das demais validações):
   - Registra na aba 'Situações Críticas' mesmo se houver campos incompletos,
     duplicidades ou inconsistências.
3. Validação de campos obrigatórios -> 'Pendente de regularização'.
4. Detecção de duplicidades (mesmo CEP + Endereço + Número + Bairro + Categoria)
   -> 'Duplicada' (mantendo rastreabilidade do protocolo original).
5. Verificação de região/bairro ativo na aba 'Regioes'
   -> 'Fora da área de atendimento' (registrando motivo).
6. Mapeamento do departamento responsável na aba 'Departamentos'
   -> 'Análise manual' se categoria não cadastrada.
7. Classificação de prioridade (prevalece a maior aplicável):
   - Risco imediato = Sim -> Crítica
   - Alagamento / Árvore com risco de queda -> Alta
   - Vazamento de água / Iluminação pública -> Média
   - Demais categorias válidas -> Normal
8. Cálculo de prazo em dias úteis (ignorando sábados e domingos) a partir da data
   de processamento:
   - Crítica -> Encaminhamento imediato
   - Alta -> 1 dia útil
   - Média -> min(3 dias úteis, prazo padrão do departamento)
   - Normal -> prazo padrão do departamento
9. Salvamento da planilha consolidada, geração de planilhas individuais por
   departamento e encaminhamento por e-mail (SMTP real opcional ou geração de .eml/.html).
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import re
import smtplib
import time
import unicodedata
from email.message import EmailMessage
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# E-mails fictícios institucionais por departamento (para encaminhamento)
EMAILS_DEPARTAMENTOS = {
    "Secretaria de Obras": "obras@prefeiturainteligente.gov.br",
    "Departamento de Iluminação": "iluminacao@prefeiturainteligente.gov.br",
    "Secretaria de Serviços Urbanos": "servicos.urbanos@prefeiturainteligente.gov.br",
    "Companhia de Saneamento": "saneamento@prefeiturainteligente.gov.br",
    "Secretaria do Meio Ambiente": "meio.ambiente@prefeiturainteligente.gov.br",
    "Defesa Civil": "defesa.civil@prefeiturainteligente.gov.br",
    "Secretaria de Mobilidade Urbana": "mobilidade@prefeiturainteligente.gov.br",
}

ORDEM_PRIORIDADE = {"Crítica": 1, "Alta": 2, "Média": 3, "Normal": 4}


def normalizar(texto: Any) -> str:
    if texto is None:
        return ""
    s = str(texto).strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", s).upper().strip()


def esta_vazio(valor: Any) -> bool:
    if valor is None:
        return True
    s = str(valor).strip()
    return s == "" or s.lower() in {"nan", "none", "null"}


def adicionar_dias_uteis(data_base: dt.date, dias_uteis: int) -> dt.date:
    """Soma dias úteis a partir de data_base, desconsiderando sábados (5) e domingos (6)."""
    atual = data_base
    restantes = dias_uteis
    while restantes > 0:
        atual += dt.timedelta(days=1)
        if atual.weekday() < 5:  # 0=Seg .. 4=Sex
            restantes -= 1
    return atual


def validar_campos_obrigatorios(reg: Dict[str, Any]) -> List[str]:
    obrigatorios = [
        "Protocolo",
        "Nome do cidadão",
        "CEP",
        "Endereço",
        "Número",
        "Bairro",
        "Categoria",
        "Descrição",
        "Data da solicitação",
        "Risco imediato",
    ]
    faltantes = [c for c in obrigatorios if esta_vazio(reg.get(c))]

    # Validação de formato do CEP (quando preenchido)
    cep = str(reg.get("CEP") or "").strip()
    if cep and not re.fullmatch(r"\d{5}-?\d{3}", cep):
        faltantes.append("CEP (formato inválido)")

    # Validação de data da solicitação
    data_val = reg.get("Data da solicitação")
    if not esta_vazio(data_val):
        if isinstance(data_val, (dt.date, dt.datetime)):
            pass
        else:
            s_data = str(data_val).strip()[:10]
            try:
                dt.datetime.strptime(s_data, "%Y-%m-%d")
            except ValueError:
                try:
                    dt.datetime.strptime(s_data, "%d/%m/%Y")
                except ValueError:
                    faltantes.append("Data da solicitação (inválida)")

    return faltantes


def classificar_prioridade(categoria: str, risco_imediato: bool) -> str:
    """
    Regras de prioridade (prevalece a maior aplicável):
    1) Risco imediato == Sim -> Crítica
    2) Alagamento ou Árvore com risco de queda -> Alta
    3) Vazamento de água ou Iluminação pública -> Média
    4) Demais categorias válidas -> Normal
    """
    if risco_imediato:
        return "Crítica"
    cat_norm = normalizar(categoria)
    if cat_norm in {"ALAGAMENTO", "ARVORE COM RISCO DE QUEDA"}:
        return "Alta"
    if cat_norm in {"VAZAMENTO DE AGUA", "ILUMINACAO PUBLICA"}:
        return "Média"
    return "Normal"


def calcular_prazo(
    prioridade: str,
    prazo_padrao_depto: int,
    data_proc: dt.date,
) -> Tuple[str, Optional[dt.date], int]:
    """
    Retorna: (texto_prazo, data_limite, qtd_dias_uteis)
    """
    if prioridade == "Crítica":
        return "Encaminhamento imediato", data_proc, 0
    if prioridade == "Alta":
        dias = 1
        data_lim = adicionar_dias_uteis(data_proc, dias)
        return f"{data_lim.strftime('%d/%m/%Y')} (1 dia útil)", data_lim, dias
    if prioridade == "Média":
        dias = min(3, int(prazo_padrao_depto))
        data_lim = adicionar_dias_uteis(data_proc, dias)
        return f"{data_lim.strftime('%d/%m/%Y')} ({dias} dias úteis)", data_lim, dias
    # Normal
    dias = int(prazo_padrao_depto)
    data_lim = adicionar_dias_uteis(data_proc, dias)
    rotulo_dia = "dia útil" if dias == 1 else "dias úteis"
    return f"{data_lim.strftime('%d/%m/%Y')} ({dias} {rotulo_dia})", data_lim, dias


def aplicar_estilo_planilha(ws, cor_header: str = "1E3A5F") -> None:
    thin = Side(style="thin", color="D1D5DB")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor=cor_header)
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    body_font = Font(name="Calibri", size=10, color="111827")
    zebra_fill = PatternFill("solid", fgColor="F8FAFC")

    cores_status = {
        "Encaminhada": ("DCFCE7", "166534"),
        "Pendente de regularização": ("FEF3C7", "92400E"),
        "Fora da área de atendimento": ("FEE2E2", "991B1B"),
        "Análise manual": ("E0E7FF", "3730A3"),
        "Duplicada": ("F3F4F6", "4B5563"),
        "Crítica - Avaliação Humana Prioritária": ("FEE2E2", "991B1B"),
    }
    cores_prio = {
        "Crítica": ("DC2626", "FFFFFF"),
        "Alta": ("F97316", "FFFFFF"),
        "Média": ("F59E0B", "111827"),
        "Normal": ("E2E8F0", "1E293B"),
    }

    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    col_status = headers.index("Status") + 1 if "Status" in headers else None
    col_prio = headers.index("Prioridade") + 1 if "Prioridade" in headers else None

    for c in range(1, ws.max_column + 1):
        cell = ws.cell(1, c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    ws.row_dimensions[1].height = 26

    for r in range(2, ws.max_row + 1):
        ws.row_dimensions[r].height = 20
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            cell.font = body_font
            cell.border = border
            cell.alignment = Alignment(vertical="center", horizontal="left")
            if r % 2 == 1:
                cell.fill = zebra_fill

            if c == col_status and cell.value in cores_status:
                bg, fg = cores_status[cell.value]
                cell.fill = PatternFill("solid", fgColor=bg)
                cell.font = Font(name="Calibri", size=10, bold=True, color=fg)
                cell.alignment = Alignment(horizontal="center", vertical="center")

            if c == col_prio and cell.value in cores_prio:
                bg, fg = cores_prio[cell.value]
                cell.fill = PatternFill("solid", fgColor=bg)
                cell.font = Font(name="Calibri", size=10, bold=True, color=fg)
                cell.alignment = Alignment(horizontal="center", vertical="center")

    # Ajusta larguras automaticamente
    for c in range(1, ws.max_column + 1):
        max_len = max(
            len(str(ws.cell(r, c).value or "")) for r in range(1, ws.max_row + 1)
        )
        ws.column_dimensions[get_column_letter(c)].width = min(max(max_len + 4, 12), 42)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"


def gerar_planilha_departamento(
    depto: str,
    registros: List[Dict[str, Any]],
    pasta_saida: Path,
) -> Path:
    wb_dep = openpyxl.Workbook()
    ws = wb_dep.active
    ws.title = "Encaminhamentos"

    colunas = [
        "Protocolo",
        "Prioridade",
        "Prazo",
        "Região administrativa",
        "Categoria",
        "Nome do cidadão",
        "CEP",
        "Endereço",
        "Número",
        "Bairro",
        "Descrição",
        "Data da solicitação",
        "Risco imediato",
        "Status",
    ]
    ws.append(colunas)

    # Ordena por prioridade (Crítica -> Alta -> Média -> Normal) e Protocolo
    ordenados = sorted(
        registros,
        key=lambda x: (ORDEM_PRIORIDADE.get(x.get("Prioridade", "Normal"), 9), x.get("Protocolo", "")),
    )
    for reg in ordenados:
        ws.append([reg.get(col, "") for col in colunas])

    aplicar_estilo_planilha(ws, cor_header="0F766E")
    slug = re.sub(r"[^a-z0-9]+", "_", normalizar(depto).lower()).strip("_")
    caminho = pasta_saida / f"Encaminhamento_{slug}.xlsx"
    wb_dep.save(caminho)
    return caminho


def encaminhar_email_departamento(
    depto: str,
    registros: List[Dict[str, Any]],
    anexo_xlsx: Path,
    pasta_emails: Path,
    data_proc: dt.date,
    enviar_smtp: bool = False,
) -> Dict[str, Any]:
    destinatario = EMAILS_DEPARTAMENTOS.get(depto, "triagem@prefeiturainteligente.gov.br")
    remetente = os.getenv("SMTP_USER", "rpa.central@prefeiturainteligente.gov.br")

    qtd_crit = sum(1 for r in registros if r["Prioridade"] == "Crítica")
    qtd_alta = sum(1 for r in registros if r["Prioridade"] == "Alta")
    qtd_media = sum(1 for r in registros if r["Prioridade"] == "Média")
    qtd_norm = sum(1 for r in registros if r["Prioridade"] == "Normal")

    assunto = (
        f"[RPA Prefeitura] {len(registros)} solicitação(ões) encaminhada(s) — {depto} "
        f"({data_proc.strftime('%d/%m/%Y')})"
    )
    if qtd_crit > 0:
        assunto = f"[URGENTE - {qtd_crit} CRÍTICA(S)] " + assunto

    linhas_tabela = ""
    ordenados = sorted(
        registros,
        key=lambda x: (ORDEM_PRIORIDADE.get(x["Prioridade"], 9), x["Protocolo"]),
    )
    for r in ordenados:
        linhas_tabela += (
            f"<tr>"
            f"<td style='padding:6px 8px;border:1px solid #ddd;'><b>{r['Protocolo']}</b></td>"
            f"<td style='padding:6px 8px;border:1px solid #ddd;'>{r['Prioridade']}</td>"
            f"<td style='padding:6px 8px;border:1px solid #ddd;'>{r['Categoria']}</td>"
            f"<td style='padding:6px 8px;border:1px solid #ddd;'>{r['Bairro']} ({r['Região administrativa']})</td>"
            f"<td style='padding:6px 8px;border:1px solid #ddd;'>{r['Endereço']}, {r['Número']}</td>"
            f"<td style='padding:6px 8px;border:1px solid #ddd;'><b>{r['Prazo']}</b></td>"
            f"</tr>"
        )

    corpo_html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<body style="font-family: Arial, sans-serif; color: #1f2937;">
  <h2 style="color:#1e3a5f;">Prefeitura Inteligente — Encaminhamento Automático RPA</h2>
  <p>Prezada equipe da <b>{depto}</b>,</p>
  <p>Segue o lote diário de solicitações triadas automaticamente em <b>{data_proc.strftime('%d/%m/%Y')}</b>:</p>
  <ul>
    <li><b>Total encaminhado:</b> {len(registros)} solicitação(ões)</li>
    <li><b>Crítica (encaminhamento imediato):</b> {qtd_crit} | <b>Alta:</b> {qtd_alta} | <b>Média:</b> {qtd_media} | <b>Normal:</b> {qtd_norm}</li>
  </ul>
  <table style="border-collapse:collapse; width:100%; font-size:13px;">
    <thead>
      <tr style="background:#1e3a5f; color:#ffffff;">
        <th style="padding:8px;border:1px solid #1e3a5f;">Protocolo</th>
        <th style="padding:8px;border:1px solid #1e3a5f;">Prioridade</th>
        <th style="padding:8px;border:1px solid #1e3a5f;">Categoria</th>
        <th style="padding:8px;border:1px solid #1e3a5f;">Bairro / Região</th>
        <th style="padding:8px;border:1px solid #1e3a5f;">Endereço</th>
        <th style="padding:8px;border:1px solid #1e3a5f;">Prazo SLA</th>
      </tr>
    </thead>
    <tbody>{linhas_tabela}</tbody>
  </table>
  <p style="margin-top:16px; font-size:12px; color:#6b7280;">
    Anexo: <code>{anexo_xlsx.name}</code> · Gerado automaticamente pelo Robô RPA de Triagem Municipal.
  </p>
</body>
</html>"""

    msg = EmailMessage()
    msg["Subject"] = assunto
    msg["From"] = remetente
    msg["To"] = destinatario
    msg.set_content(
        f"Prefeitura Inteligente - {depto}\n"
        f"Data: {data_proc.strftime('%d/%m/%Y')}\n"
        f"Total encaminhado: {len(registros)} (Críticas: {qtd_crit}, Altas: {qtd_alta}, Médias: {qtd_media}, Normais: {qtd_norm})\n"
        f"Planilha em anexo: {anexo_xlsx.name}"
    )
    msg.add_alternative(corpo_html, subtype="html")
    with open(anexo_xlsx, "rb") as fa:
        msg.add_attachment(
            fa.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=anexo_xlsx.name,
        )

    slug = re.sub(r"[^a-z0-9]+", "_", normalizar(depto).lower()).strip("_")
    caminho_eml = pasta_emails / f"Email_{slug}.eml"
    caminho_html = pasta_emails / f"Preview_{slug}.html"
    with open(caminho_eml, "wb") as fe:
        fe.write(bytes(msg))
    with open(caminho_html, "w", encoding="utf-8") as fh:
        fh.write(corpo_html)

    status_envio = "Simulado (.eml + .html gerados com anexo)"
    if enviar_smtp:
        host = os.getenv("SMTP_HOST")
        port = int(os.getenv("SMTP_PORT", "587"))
        user = os.getenv("SMTP_USER")
        pwd = os.getenv("SMTP_PASS")
        if host and user and pwd:
            with smtplib.SMTP(host, port, timeout=15) as server:
                server.starttls()
                server.login(user, pwd)
                server.send_message(msg)
            status_envio = f"Enviado via SMTP ({host})"
        else:
            status_envio = "SMTP não configurado — salvo em .eml/.html"

    return {
        "departamento": depto,
        "destinatario": destinatario,
        "assunto": assunto,
        "quantidade": len(registros),
        "planilha_anexa": str(anexo_xlsx),
        "arquivo_eml": str(caminho_eml),
        "status_envio": status_envio,
    }


def executar_robo(
    caminho_entrada: Path,
    caminho_saida: Path,
    pasta_deptos: Path,
    pasta_emails: Path,
    data_processamento: Optional[dt.date] = None,
    modo_demo: bool = False,
    enviar_smtp: bool = False,
) -> Dict[str, Any]:
    inicio = time.perf_counter()
    data_proc = data_processamento or dt.date.today()

    pasta_deptos.mkdir(parents=True, exist_ok=True)
    pasta_emails.mkdir(parents=True, exist_ok=True)

    print("=" * 78)
    print("  ROBÔ RPA — PREFEITURA INTELIGENTE: TRIAGEM DE SOLICITAÇÕES AO CIDADÃO")
    print("=" * 78)
    print(f"  Arquivo de entrada : {caminho_entrada}")
    print(f"  Data referência    : {data_proc.strftime('%d/%m/%Y')} (dias úteis ignoram sáb/dom)")
    print("-" * 78)

    wb = openpyxl.load_workbook(caminho_entrada)
    ws_sol = wb["Solicitacoes"]
    ws_dep = wb["Departamentos"]
    ws_reg = wb["Regioes"]
    ws_crit = wb["Situações Críticas"]

    # 1) Carrega tabela de Departamentos
    departamentos: Dict[str, Tuple[str, int]] = {}
    for row in ws_dep.iter_rows(min_row=2, values_only=True):
        if row and not esta_vazio(row[0]):
            departamentos[normalizar(row[0])] = (str(row[1]).strip(), int(row[2]))

    # 2) Carrega tabela de Regiões
    regioes: Dict[str, Tuple[str, str]] = {}
    for row in ws_reg.iter_rows(min_row=2, values_only=True):
        if row and not esta_vazio(row[0]):
            regioes[normalizar(row[0])] = (str(row[1]).strip(), str(row[2]).strip())

    # 3) Garante colunas extras de auditoria na aba Solicitacoes (mantendo as 13 originais intactas!)
    headers_sol = [ws_sol.cell(1, c).value for c in range(1, ws_sol.max_column + 1)]
    col_map = {h: idx + 1 for idx, h in enumerate(headers_sol) if h}
    colunas_extras = ["Região administrativa", "Prioridade", "Motivo / Observação", "Data processamento"]
    for extra in colunas_extras:
        if extra not in col_map:
            nova_col = ws_sol.max_column + 1
            ws_sol.cell(1, nova_col, extra)
            col_map[extra] = nova_col

    # Limpa eventuais registros antigos em Situações Críticas (mantendo cabeçalho)
    if ws_crit.max_row > 1:
        ws_crit.delete_rows(2, ws_crit.max_row - 1)

    chaves_vistas: Dict[Tuple[str, str, str, str, str], str] = {}
    lotes_departamento: Dict[str, List[Dict[str, Any]]] = {}

    contadores = {
        "total_lidas": 0,
        "situacoes_criticas_registradas": 0,
        "encaminhadas": 0,
        "pendente_regularizacao": 0,
        "duplicadas": 0,
        "fora_area_atendimento": 0,
        "analise_manual": 0,
        "por_prioridade": {"Crítica": 0, "Alta": 0, "Média": 0, "Normal": 0},
        "por_departamento": {},
    }

    for r in range(2, ws_sol.max_row + 1):
        protocolo = ws_sol.cell(r, col_map["Protocolo"]).value
        if esta_vazio(protocolo):
            continue

        status_atual = ws_sol.cell(r, col_map["Status"]).value
        if not esta_vazio(status_atual) and normalizar(status_atual) != "RECEBIDA":
            continue  # Já analisada anteriormente

        contadores["total_lidas"] += 1
        reg = {
            h: ws_sol.cell(r, col_map[h]).value
            for h in [
                "Protocolo", "Nome do cidadão", "CEP", "Endereço", "Número",
                "Bairro", "Categoria", "Descrição", "Data da solicitação", "Risco imediato"
            ]
        }

        risco_sim = normalizar(reg["Risco imediato"]) == "SIM"

        # ------------------------------------------------------------------
        # ETAPA 3: Risco imediato SEMPRE ANTES de qualquer validação
        # ------------------------------------------------------------------
        if risco_sim:
            contadores["situacoes_criticas_registradas"] += 1
            ws_crit.append([
                reg["Protocolo"],
                reg["Nome do cidadão"],
                reg["CEP"],
                reg["Endereço"],
                reg["Número"],
                reg["Bairro"],
                reg["Categoria"],
                reg["Descrição"],
                reg["Data da solicitação"],
                reg["Risco imediato"],
                "Crítica - Avaliação Humana Prioritária",
                "Encaminhamento imediato",
            ])

        ws_sol.cell(r, col_map["Data processamento"], data_proc.strftime("%d/%m/%Y"))

        # ------------------------------------------------------------------
        # ETAPA 4: Validação de campos obrigatórios
        # ------------------------------------------------------------------
        faltantes = validar_campos_obrigatorios(reg)
        if faltantes:
            status_final = "Pendente de regularização"
            motivo = f"Campos ausentes/inválidos: {', '.join(faltantes)}"
            if risco_sim:
                motivo += " | Registrada em Situações Críticas"
            ws_sol.cell(r, col_map["Status"], status_final)
            ws_sol.cell(r, col_map["Motivo / Observação"], motivo)
            contadores["pendente_regularizacao"] += 1
            print(f"  [PENDENTE]   {protocolo} | {motivo}")
            if modo_demo:
                time.sleep(0.08)
            continue

        # ------------------------------------------------------------------
        # VERIFICAÇÃO DE DUPLICIDADE (mesmo endereço + bairro + categoria)
        # ------------------------------------------------------------------
        chave_dup = (
            normalizar(reg["CEP"]),
            normalizar(reg["Endereço"]),
            normalizar(reg["Número"]),
            normalizar(reg["Bairro"]),
            normalizar(reg["Categoria"]),
        )
        if chave_dup in chaves_vistas:
            prot_orig = chaves_vistas[chave_dup]
            status_final = "Duplicada"
            motivo = f"Duplicidade detectada (ref. {prot_orig})"
            if risco_sim:
                motivo += " | Risco registrado em Situações Críticas"
            ws_sol.cell(r, col_map["Status"], status_final)
            ws_sol.cell(r, col_map["Motivo / Observação"], motivo)
            contadores["duplicadas"] += 1
            print(f"  [DUPLICADA]  {protocolo} | {motivo}")
            if modo_demo:
                time.sleep(0.08)
            continue
        chaves_vistas[chave_dup] = str(protocolo)

        # ------------------------------------------------------------------
        # ETAPA 5: Verificação da região de atendimento (aba Regioes)
        # ------------------------------------------------------------------
        bairro_norm = normalizar(reg["Bairro"])
        if bairro_norm not in regioes:
            status_final = "Fora da área de atendimento"
            motivo = f"Bairro não cadastrado: '{reg['Bairro']}'"
            if risco_sim:
                motivo += " | Risco registrado em Situações Críticas"
            ws_sol.cell(r, col_map["Status"], status_final)
            ws_sol.cell(r, col_map["Motivo / Observação"], motivo)
            contadores["fora_area_atendimento"] += 1
            print(f"  [FORA ÁREA]  {protocolo} | {motivo}")
            if modo_demo:
                time.sleep(0.08)
            continue

        regiao_adm, status_regiao = regioes[bairro_norm]
        if normalizar(status_regiao) != "ATIVO":
            status_final = "Fora da área de atendimento"
            motivo = f"Atendimento inativo no bairro '{reg['Bairro']}' (Região {regiao_adm})"
            if risco_sim:
                motivo += " | Risco registrado em Situações Críticas"
            ws_sol.cell(r, col_map["Região administrativa"], regiao_adm)
            ws_sol.cell(r, col_map["Status"], status_final)
            ws_sol.cell(r, col_map["Motivo / Observação"], motivo)
            contadores["fora_area_atendimento"] += 1
            print(f"  [FORA ÁREA]  {protocolo} | {motivo}")
            if modo_demo:
                time.sleep(0.08)
            continue

        ws_sol.cell(r, col_map["Região administrativa"], regiao_adm)

        # ------------------------------------------------------------------
        # ETAPA 6: Identificação do departamento responsável (aba Departamentos)
        # ------------------------------------------------------------------
        cat_norm = normalizar(reg["Categoria"])
        if cat_norm not in departamentos:
            status_final = "Análise manual"
            motivo = f"Categoria não cadastrada: '{reg['Categoria']}'"
            ws_sol.cell(r, col_map["Status"], status_final)
            ws_sol.cell(r, col_map["Motivo / Observação"], motivo)
            contadores["analise_manual"] += 1
            print(f"  [MANUAL]     {protocolo} | {motivo}")
            if modo_demo:
                time.sleep(0.08)
            continue

        depto, prazo_padrao = departamentos[cat_norm]

        # ------------------------------------------------------------------
        # ETAPAS 7 & 8: Classificação de prioridade e definição de prazo SLA
        # ------------------------------------------------------------------
        prioridade = classificar_prioridade(str(reg["Categoria"]), risco_sim)
        prazo_txt, _, _ = calcular_prazo(prioridade, prazo_padrao, data_proc)

        status_final = "Encaminhada"
        motivo = f"Encaminhada à {depto} (Região {regiao_adm})"
        if risco_sim:
            motivo = "Risco imediato registrado em Situações Críticas | " + motivo

        ws_sol.cell(r, col_map["Status"], status_final)
        ws_sol.cell(r, col_map["Responsável"], depto)
        ws_sol.cell(r, col_map["Prazo"], prazo_txt)
        ws_sol.cell(r, col_map["Prioridade"], prioridade)
        ws_sol.cell(r, col_map["Motivo / Observação"], motivo)

        contadores["encaminhadas"] += 1
        contadores["por_prioridade"][prioridade] += 1
        contadores["por_departamento"][depto] = contadores["por_departamento"].get(depto, 0) + 1

        reg_enriquecido = dict(reg)
        reg_enriquecido.update({
            "Região administrativa": regiao_adm,
            "Responsável": depto,
            "Prioridade": prioridade,
            "Prazo": prazo_txt,
            "Status": status_final,
        })
        lotes_departamento.setdefault(depto, []).append(reg_enriquecido)

        tag_prio = f"[{prioridade.upper():7s}]"
        print(f"  {tag_prio}  {protocolo} -> {depto} | Prazo: {prazo_txt}")
        if modo_demo:
            time.sleep(0.08)

    # Estiliza abas e salva a planilha consolidada
    aplicar_estilo_planilha(ws_sol, cor_header="1E3A5F")
    aplicar_estilo_planilha(ws_crit, cor_header="991B1B")
    aplicar_estilo_planilha(ws_dep, cor_header="334155")
    aplicar_estilo_planilha(ws_reg, cor_header="334155")
    wb.save(caminho_saida)

    # ------------------------------------------------------------------
    # ETAPA 10: Gera planilhas por setor e encaminha e-mails
    # ------------------------------------------------------------------
    print("-" * 78)
    print("  ETAPA 10 — GERAÇÃO DE LOTES E ENCAMINHAMENTO POR E-MAIL")
    print("-" * 78)
    despachos = []
    for depto, regs in sorted(lotes_departamento.items()):
        arq_dep = gerar_planilha_departamento(depto, regs, pasta_deptos)
        info_email = encaminhar_email_departamento(
            depto, regs, arq_dep, pasta_emails, data_proc, enviar_smtp=enviar_smtp
        )
        despachos.append(info_email)
        print(f"  [E-MAIL] {depto:32s} -> {info_email['destinatario']} ({len(regs)} itens)")

    duracao = round(time.perf_counter() - inicio, 2)
    relatorio = {
        "data_processamento": data_proc.strftime("%Y-%m-%d"),
        "duracao_segundos": duracao,
        "arquivo_processado": str(caminho_saida),
        "resumo": contadores,
        "encaminhamentos_email": despachos,
    }
    caminho_json = caminho_saida.parent / "relatorio_execucao_rpa.json"
    with open(caminho_json, "w", encoding="utf-8") as fj:
        json.dump(relatorio, fj, indent=2, ensure_ascii=False)

    print("=" * 78)
    print(f"  CONCLUÍDO EM {duracao:.2f}s | {contadores['total_lidas']} solicitações processadas!")
    print(f"  • Encaminhadas: {contadores['encaminhadas']}  |  Situações Críticas (aba): {contadores['situacoes_criticas_registradas']}")
    print(f"  • Pendentes: {contadores['pendente_regularizacao']} | Duplicadas: {contadores['duplicadas']} | Fora da área: {contadores['fora_area_atendimento']} | Análise manual: {contadores['analise_manual']}")
    print(f"  • Planilha salva em: {caminho_saida}")
    print("=" * 78)
    return relatorio


def main() -> None:
    parser = argparse.ArgumentParser(description="Robô RPA de Triagem — Prefeitura Inteligente")
    parser.add_argument("--entrada", default="entrada/Base_Prefeitura.xlsx", help="Caminho da planilha de entrada")
    parser.add_argument("--saida", default="Base_Prefeitura_Processada.xlsx", help="Caminho da planilha processada")
    parser.add_argument("--data", default=None, help="Data de referência YYYY-MM-DD (padrão: hoje)")
    parser.add_argument("--demo", action="store_true", help="Ativa pausas suaves para gravação de vídeo")
    parser.add_argument("--smtp", action="store_true", help="Ativa envio real via SMTP (requer variáveis SMTP_*)")
    args = parser.parse_args()

    data_ref = dt.datetime.strptime(args.data, "%Y-%m-%d").date() if args.data else dt.date.today()
    executar_robo(
        caminho_entrada=Path(args.entrada),
        caminho_saida=Path(args.saida),
        pasta_deptos=Path("saida_departamentos"),
        pasta_emails=Path("saida_emails"),
        data_processamento=data_ref,
        modo_demo=args.demo,
        enviar_smtp=args.smtp,
    )


if __name__ == "__main__":
    main()
