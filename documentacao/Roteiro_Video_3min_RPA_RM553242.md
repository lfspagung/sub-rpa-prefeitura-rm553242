# Roteiro do Vídeo Demonstrativo (Até 3 Minutos) — Substitutiva RPA

**Aluno:** Lucas de Freitas Pagung — **RM 553242**  
**Tema:** Prefeitura Inteligente: Triagem e Encaminhamento Automático ao Cidadão  
**Duração alvo:** **2 min 45 s** (margem segura abaixo do limite de 3 minutos)

---

## Preparação Rápida antes de Gravar (30 segundos)
Deixe abertos lado a lado no Mac:
1. O **Diagrama BPMN** (`bizagi/Diagrama_BPMN_TO_BE_Prefeitura.png` ou aberto no **Bizagi Modeler**).
2. A planilha **`Process_Assessment_Business_Case_ROI_RM553242.xlsx`** na aba `Resumo Executivo`.
3. O **Terminal** na pasta `~/Desktop/sub_rpa_prefeitura_rm553242` com o comando pronto:
   `python3 robo_triagem_prefeitura.py --demo`
4. A pasta `saida_emails/` e a planilha `Base_Prefeitura_Processada.xlsx` para abrir assim que o robô concluir.

---

## Roteiro Cronometrado (Fala Natural)

### 0:00 – 0:25 | Abertura e Problema (25s)
> "Olá! Meu nome é Lucas de Freitas Pagung, RM 553242, e esta é a entrega da Substitutiva de Robotic Process Automation: Prefeitura Inteligente no atendimento ao cidadão.
> Hoje, a prefeitura recebe 350 solicitações por dia, levando 5 minutos por triagem manual e sofrendo com 12% de retrabalho — o que consome mais de 700 horas por mês, quase 4 servidores inteiros só conferindo planilha, atrasando ocorrências de risco."

### 0:25 – 0:55 | Estrutura do Fluxo BPMN e Regras de Negócio (30s)
*(Mostre o fluxo desenhado no Bizagi / BPMN)*
> "Desenhei o fluxo TO BE em três raias: o cidadão registra no portal, o robô RPA assume toda a triagem determinística e a equipe humana foca apenas em exceções e casos críticos.
> Seguindo rigorosamente as regras: primeiro o robô identifica risco imediato e grava na aba Situações Críticas — mesmo se a solicitação tiver dados faltando ou for duplicada. Depois valida campos obrigatórios, barra duplicidades, checa se o bairro está ativo na aba Regiões, mapeia o departamento, classifica a prioridade sempre pela maior aplicável e calcula o prazo em dias úteis pulando sábados e domingos."

### 0:55 – 1:45 | Demonstração do Robô Funcionando ao Vivo (50s)
*(Mude para o Terminal e pressione Enter em `python3 robo_triagem_prefeitura.py --demo`)*
> "Vou executar o robô agora sobre os 39 registros reais da `Base_Prefeitura.xlsx`.
> Veja no terminal ele aplicando cada regra em tempo real: classificando prioridades Crítica, Alta, Média e Normal, calculando prazos como 1 dia útil para Alta e o menor valor entre 3 dias e o prazo padrão para Média.
> No final do lote, ele detectou as 4 duplicidades (PROT-0031 a 0034), marcou PROT-0035 e 0040 como pendentes de regularização por falta de endereço e CEP, barrou os bairros Jardim Novo inativo e Bairro Desconhecido como fora da área, e enviou Ruído Urbano para análise manual. Em menos de 1 segundo, gerou a planilha atualizada, 7 planilhas setoriais e os 7 e-mails prontos com anexo."

### 1:45 – 2:15 | Evidência nos Resultados Gerados (30s)
*(Abra rapidamente `Base_Prefeitura_Processada.xlsx` mostrando `Solicitacoes` e `Situações Críticas`, e abra um `Preview_*.html` de e-mail)*
> "Aqui na planilha processada vemos os 39 registros preenchidos com status, responsável, prioridade e prazo. Na aba Situações Críticas estão todas as 7 ocorrências com risco imediato — inclusive a 0034 duplicada, a 0035 sem endereço e a 0037 de bairro desconhecido, garantindo que nenhum risco à população passe despercebido. E aqui temos o e-mail gerado automaticamente para o departamento com a planilha filtrada do setor."

### 2:15 – 2:45 | Process Assessment, ROI e Impacto Social (30s)
*(Mostre a aba `Resumo Executivo` e `Business Case & ROI` no Excel)*
> "No Process Assessment, o processo atingiu nota 4,80 de 5 (96%) — altamente viável para RPA. No Business Case, o custo atual AS IS é de R$ 31.647 por mês, ou R$ 379.764 por ano. Com a automação, considerando R$ 28 mil de implantação e R$ 2.700 mensais de sustentação, temos uma economia líquida de R$ 25.204 por mês, ganho líquido de R$ 274.457 no primeiro ano, ROI de 454,4% e payback em apenas 1,1 mês.
> Mais importante que o financeiro é o impacto social: riscos imediatos vão na hora para a equipe humana, e liberamos o equivalente a 3,5 servidores para atuar na resolução real dos problemas da cidade. Muito obrigado!"
