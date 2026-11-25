# Avaliação Substitutiva — Robotic Process Automation (RPA)
## Prefeitura Inteligente: Atendimento ao Cidadão

- **Aluno:** Lucas de Freitas Pagung — **RM 553242**
- **Turma:** 4ESPR - AV. Paulista - Noturno.
- **Disciplina:** Robotic Process Automation (FIAP)
- **Link GIT:** https://github.com/lfspagung/sub-rpa-prefeitura-rm553242
- **Link Streamlit:** https://sub-rpa-prefeitura-rm553242-488byuqehzviy4x3apprh5p.streamlit.app/
- **Link Picth Youtube:**  
 
## Conteúdo da Entrega
1. **`Process_Assessment_Business_Case_ROI_RM553242.xlsx`** — Planilha completa com 5 abas (`Resumo Executivo`, `Process Assessment`, `Business Case & ROI`, `Fluxo de Caixa 12M`, `Regras e Evidência RPA`), fórmulas dinâmicas, formatação condicional e gráficos.
2. **`bizagi/Prefeitura_Inteligente_TO_BE_Bizagi.bpmn`** — Arquivo BPMN 2.0 XML pronto para importar no **Bizagi Modeler** (`Arquivo > Importar > BPMN`), acompanhado da visualização em alta resolução (`bizagi/Diagrama_BPMN_TO_BE_Prefeitura.png`).
3. **`robo_triagem_prefeitura.py`** — Robô RPA funcional em Python (`openpyxl` + `smtplib`) que processa `entrada/Base_Prefeitura.xlsx`, gera `Base_Prefeitura_Processada.xlsx`, cria 7 planilhas setoriais em `saida_departamentos/` e 7 e-mails prontos (`.eml` + `.html`) em `saida_emails/`.

## Como Executar o Robô RPA no macOS / Windows / Linux
```bash
python3 -m venv .venv_rpa
source .venv_rpa/bin/activate
pip install openpyxl

# Execução rápida (< 1 segundo):
python3 robo_triagem_prefeitura.py

# Execução em modo demonstração (com pausas suaves para gravar o vídeo):
python3 robo_triagem_prefeitura.py --demo
```
