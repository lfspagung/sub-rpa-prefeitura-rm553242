"""
App Streamlit — Robô RPA Prefeitura Inteligente
Disciplina: Robotic Process Automation (Avaliação Substitutiva)
Autor: Lucas de Freitas Pagung — RM 553242
"""

from __future__ import annotations

import contextlib
import datetime as dt
import io
import json
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from robo_triagem_prefeitura import executar_robo

RAIZ = Path(__file__).resolve().parent
ENTRADA_PADRAO = RAIZ / "entrada" / "Base_Prefeitura.xlsx"
BPMN_IMG = RAIZ / "bizagi" / "Diagrama_BPMN_TO_BE_Prefeitura.png"
ROI_XLSX = RAIZ / "Process_Assessment_Business_Case_ROI_RM553242.xlsx"

WORK_DIR = Path(tempfile.gettempdir()) / "rpa_prefeitura_rm553242"
WORK_DIR.mkdir(parents=True, exist_ok=True)

st.set_page_config(
    page_title="RPA Prefeitura Inteligente — RM 553242",
    page_icon="🤖",
    layout="wide",
)

st.markdown(
    """
    <style>
        .rpa-banner {
            background: linear-gradient(90deg, #1B365D 0%, #0F766E 100%);
            color: #FFFFFF;
            padding: 18px 24px;
            border-radius: 12px;
            margin-bottom: 16px;
        }
        .rpa-banner h1 {
            margin: 0;
            font-size: 1.65rem;
            color: #FFFFFF;
        }
        .rpa-banner p {
            margin: 4px 0 0 0;
            font-size: 0.92rem;
            color: #E2E8F0;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="rpa-banner">
        <h1>🤖 Robô RPA — Prefeitura Inteligente: Triagem ao Cidadão</h1>
        <p>Robotic Process Automation (Avaliação Substitutiva) · FIAP · Lucas de Freitas Pagung (RM 553242)</p>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("Configuração do Robô")
    arquivo_upload = st.file_uploader(
        "Planilha de entrada (.xlsx) — opcional",
        type=["xlsx"],
        help="Se não enviar nenhum arquivo, o robô utiliza entrada/Base_Prefeitura.xlsx.",
    )
    data_ref = st.date_input(
        "Data de referência (cálculo de dias úteis)",
        value=dt.date.today(),
        format="DD/MM/YYYY",
    )
    modo_demo = st.checkbox("Modo demonstração (animação suave)", value=True)
    enviar_smtp = st.checkbox("Tentar envio SMTP real (se configurado)", value=False)

    st.divider()
    executar_btn = st.button("▶ Executar Robô RPA", type="primary", width="stretch")

    st.divider()
    st.caption("Arquivos oficiais da entrega:")
    if ROI_XLSX.exists():
        st.download_button(
            "📊 Baixar Planilha Assessment & ROI",
            data=ROI_XLSX.read_bytes(),
            file_name=ROI_XLSX.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )

if executar_btn:
    if arquivo_upload is not None:
        caminho_entrada = WORK_DIR / "Base_Upload_Temp.xlsx"
        caminho_entrada.write_bytes(arquivo_upload.getvalue())
    else:
        caminho_entrada = ENTRADA_PADRAO

    if not caminho_entrada.exists():
        st.error(f"Planilha de entrada não encontrada em: {caminho_entrada}")
    else:
        saida_exec = WORK_DIR / "Base_Prefeitura_Processada.xlsx"
        pasta_deptos_exec = WORK_DIR / "saida_departamentos"
        pasta_emails_exec = WORK_DIR / "saida_emails"

        barra = st.progress(0, text="Iniciando Robô RPA...")
        buffer_log = io.StringIO()
        barra.progress(20, text="Lendo abas Solicitacoes, Departamentos e Regioes...")
        with contextlib.redirect_stdout(buffer_log):
            relatorio = executar_robo(
                caminho_entrada=caminho_entrada,
                caminho_saida=saida_exec,
                pasta_deptos=pasta_deptos_exec,
                pasta_emails=pasta_emails_exec,
                data_processamento=data_ref,
                modo_demo=modo_demo,
                enviar_smtp=enviar_smtp,
            )
        barra.progress(100, text="Triagem RPA concluída com sucesso!")
        st.session_state["rpa_relatorio"] = relatorio
        st.session_state["rpa_log"] = buffer_log.getvalue()
        st.session_state["rpa_saida_xlsx"] = str(saida_exec)
        st.session_state["rpa_base_dir"] = str(WORK_DIR)
        st.success(
            f"Robô executado em **{relatorio['duracao_segundos']:.2f}s** — "
            f"**{relatorio['resumo']['total_lidas']}** solicitações processadas, "
            f"**{relatorio['resumo']['situacoes_criticas_registradas']}** situações críticas registradas e "
            f"**{len(relatorio['encaminhamentos_email'])}** lotes enviados por e-mail."
        )

relatorio_atual = st.session_state.get("rpa_relatorio")
if relatorio_atual is not None:
    saida_ativa = Path(st.session_state["rpa_saida_xlsx"])
    base_ativa = Path(st.session_state["rpa_base_dir"])
else:
    saida_ativa = RAIZ / "Base_Prefeitura_Processada.xlsx"
    base_ativa = RAIZ
    rel_json = RAIZ / "relatorio_execucao_rpa.json"
    if rel_json.exists():
        relatorio_atual = json.loads(rel_json.read_text(encoding="utf-8"))

aba_exec, aba_planilha, aba_emails, aba_roi = st.tabs(
    [
        "1. Execução do Robô & Resumo",
        "2. Planilha Processada & Críticas",
        "3. E-mails & Lotes por Setor",
        "4. Fluxo BPMN & Business Case (ROI)",
    ]
)

with aba_exec:
    if relatorio_atual is None:
        st.info("Clique em **▶ Executar Robô RPA** na barra lateral para iniciar a triagem automática.")
    else:
        res = relatorio_atual["resumo"]
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c1.metric("Total Lidas", res["total_lidas"])
        c2.metric("Encaminhadas", res["encaminhadas"])
        c3.metric("Situações Críticas", res["situacoes_criticas_registradas"])
        c4.metric("Duplicadas", res["duplicadas"])
        c5.metric("Pendentes / Fora Área", res["pendente_regularizacao"] + res["fora_area_atendimento"])
        c6.metric("Análise Manual", res["analise_manual"])

        col_a, col_b = st.columns(2)
        with col_a:
            st.subheader("Distribuição por Prioridade (Encaminhadas)")
            df_prio = pd.DataFrame(
                [{"Prioridade": k, "Quantidade": v} for k, v in res["por_prioridade"].items()]
            )
            st.dataframe(df_prio, width="stretch", hide_index=True)

        with col_b:
            st.subheader("Lotes por Departamento")
            df_dep = pd.DataFrame(
                [{"Departamento": k, "Solicitações": v} for k, v in sorted(res["por_departamento"].items())]
            )
            st.dataframe(df_dep, width="stretch", hide_index=True)

        if saida_ativa.exists():
            st.download_button(
                "📥 Baixar Base_Prefeitura_Processada.xlsx",
                data=saida_ativa.read_bytes(),
                file_name="Base_Prefeitura_Processada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        log_txt = st.session_state.get("rpa_log")
        if log_txt:
            st.subheader("Log da Execução em Tempo Real")
            st.code(log_txt, language="text")

with aba_planilha:
    if not saida_ativa.exists():
        st.info("Execute o robô para visualizar as abas processadas.")
    else:
        st.subheader("Aba 'Situações Críticas' (Registradas antes das demais validações)")
        df_crit = pd.read_excel(saida_ativa, sheet_name="Situações Críticas")
        st.dataframe(df_crit, width="stretch", hide_index=True)

        st.divider()
        st.subheader("Aba 'Solicitacoes' (39 registros triados)")
        df_sol = pd.read_excel(saida_ativa, sheet_name="Solicitacoes")
        f1, f2 = st.columns(2)
        status_opts = ["Todos"] + sorted(df_sol["Status"].dropna().astype(str).unique().tolist())
        prio_opts = (
            ["Todas"] + sorted(df_sol["Prioridade"].dropna().astype(str).unique().tolist())
            if "Prioridade" in df_sol.columns
            else ["Todas"]
        )
        sel_status = f1.selectbox("Filtrar por Status:", status_opts)
        sel_prio = f2.selectbox("Filtrar por Prioridade:", prio_opts)

        df_view = df_sol.copy()
        if sel_status != "Todos":
            df_view = df_view[df_view["Status"].astype(str) == sel_status]
        if sel_prio != "Todas" and "Prioridade" in df_view.columns:
            df_view = df_view[df_view["Prioridade"].astype(str) == sel_prio]

        st.dataframe(df_view, width="stretch", hide_index=True)

with aba_emails:
    if relatorio_atual is None or not relatorio_atual.get("encaminhamentos_email"):
        st.info("Execute o robô para gerar os lotes setoriais e pré-visualizar os e-mails.")
    else:
        despachos = relatorio_atual["encaminhamentos_email"]
        nomes_dep = [d["departamento"] for d in despachos]
        depto_sel = st.selectbox("Selecione o Departamento para inspecionar o e-mail e anexo:", nomes_dep)
        item = next(d for d in despachos if d["departamento"] == depto_sel)

        st.markdown(
            f"**Destinatário:** `{item['destinatario']}`  \n"
            f"**Assunto:** `{item['assunto']}`  \n"
            f"**Status do envio:** {item['status_envio']}"
        )

        arq_xlsx = Path(item["planilha_anexa"])
        if not arq_xlsx.is_absolute():
            arq_xlsx = base_ativa / arq_xlsx
        arq_eml = Path(item["arquivo_eml"])
        if not arq_eml.is_absolute():
            arq_eml = base_ativa / arq_eml
        arq_html = arq_eml.parent / arq_eml.name.replace("Email_", "Preview_").replace(".eml", ".html")

        d1, d2 = st.columns(2)
        if arq_xlsx.exists():
            d1.download_button(
                f"📥 Baixar planilha anexa ({arq_xlsx.name})",
                data=arq_xlsx.read_bytes(),
                file_name=arq_xlsx.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch",
            )
        if arq_eml.exists():
            d2.download_button(
                f"✉️ Baixar arquivo de e-mail ({arq_eml.name})",
                data=arq_eml.read_bytes(),
                file_name=arq_eml.name,
                mime="message/rfc822",
                width="stretch",
            )

        if arq_html.exists():
            st.markdown("### Pré-visualização do E-mail Encaminhado")
            components.html(arq_html.read_text(encoding="utf-8"), height=420, scrolling=True)

with aba_roi:
    st.subheader("Process Assessment & Business Case (ROI)")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Score Process Assessment", "4,80 / 5,00", "96% — Altamente Viável")
    k2.metric("Custo Anual AS IS", "R$ 379.764,00", "703,27 h/mês (4,00 FTEs)")
    k3.metric("Ganho Líquido Ano 1", "R$ 274.457,60", "ROI: 454,4%")
    k4.metric("Payback Estimado", "1,11 mês", "3,52 FTEs liberados")

    df_resumo_roi = pd.DataFrame(
        [
            {"Indicador": "Solicitações/dia (AS IS)", "Valor": "350/dia (7.700/mês | 92.400/ano)"},
            {"Indicador": "Tempo e retrabalho AS IS", "Valor": "5 min/sol. + 12% retrabalho (+4 min) = 703,27 h/mês"},
            {"Indicador": "Custo mensal AS IS (R$ 45/h)", "Valor": "R$ 31.647,00/mês (R$ 379.764,00/ano)"},
            {"Indicador": "Investimento RPA (CAPEX + OPEX)", "Valor": "CAPEX R$ 28.000,00 | OPEX R$ 2.700,00/mês"},
            {"Indicador": "Economia líquida mensal (pós-OPEX)", "Valor": "R$ 25.204,80/mês (-620,11 h/mês)"},
            {"Indicador": "Impacto Social", "Valor": "Atendimento imediato a situações críticas e realocação de ~3,5 servidores"},
        ]
    )
    st.dataframe(df_resumo_roi, width="stretch", hide_index=True)

    st.divider()
    st.subheader("Diagrama BPMN 2.0 (TO BE — Padrão Bizagi)")
    if BPMN_IMG.exists():
        st.image(str(BPMN_IMG), caption="Fluxo BPMN TO BE — Prefeitura Inteligente com RPA", width="stretch")
    else:
        st.warning("Imagem do diagrama BPMN não encontrada em bizagi/Diagrama_BPMN_TO_BE_Prefeitura.png.")
