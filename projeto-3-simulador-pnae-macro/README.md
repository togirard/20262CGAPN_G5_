# Projeto 3: Simulador PNAE com Registro de Simulações – Grupo 5

Portfólio do Grupo 5 na disciplina Análise de Dados para Pesquisas em Políticas Públicas. Este repositório reúne três projetos, cada um na sua pasta.

## Portfólio

| Projeto | O que é | Pasta |
|---|---|---|
| Projeto 1: Simulador PNAE | Simulador em HTML do repasse do PNAE, montado a partir de uma planilha Excel | [/projeto-1-simulador-pnae](./projeto-1-simulador-pnae) |
| Projeto 2: Painel Censo | Dashboard em Excel com dados do Censo Escolar 2024 | [/projeto-2-painel-censo](./projeto-2-painel-censo) |
| Projeto 3: Simulador PNAE com macro | Simulador em Excel com macro VBA que registra cada simulação | [/projeto-3-simulador-pnae-macro](./projeto-3-simulador-pnae-macro) |

### Estrutura do Projeto 3
- /projeto-3-simulador-pnae-macro/Projeto 1 - Simulador PNAE.xlsm: a planilha com a macro em VBA.
- /projeto-3-simulador-pnae-macro/README.md: este documento.

## Projeto 3: Simulador PNAE com Registro de Simulações

### Objetivo
Este projeto dá continuidade ao simulador de repasse do PNAE (Projeto 1) e acrescenta uma macro em VBA. Ao clicar no botão "Salvar Simulação", a planilha confere se os campos foram preenchidos e grava o cenário na aba Banco_de_Dados. O registro traz o fator de ajuste, a justificativa, as matrículas ajustadas, o repasse ajustado e o usuário. Assim, o grupo mantém um histórico das simulações realizadas.

### Como usar
1. Baixe a planilha "Projeto 1 - Simulador PNAE.xlsm" desta pasta e abra no Excel instalado no computador.
2. Clique em "Habilitar Conteúdo" para ativar as macros.
3. Na aba Simulador_Escola, preencha Usuário, Racional da Taxa e Fator de Ajuste (ex.: 5% ou -10%).
4. Clique em "Salvar Simulação".
5. Confira o registro na aba Banco_de_Dados.

### Capturas de tela
<img width="835" height="787" alt="Aba Simulador_Escola" src="https://github.com/user-attachments/assets/44f21058-bc9e-47aa-8968-750fd821e1fb" />

<img width="1051" height="720" alt="Aba Banco_de_Dados" src="https://github.com/user-attachments/assets/5854f2d6-e128-4f72-822f-2823da3a65d4" />

## Uso de Inteligência Artificial
Ferramenta utilizada: Claude (Anthropic)
Para que foi usada: apoio na redação deste README. A planilha e a macro em VBA não foram feitas com IA. O grupo implementou o campo Usuário seguindo o passo a passo comentado no próprio código.
Exemplo de prompt utilizado: "Nosso grupo está realizando o seguinte trabalho no Github. Estou fazendo o Readme e gostaria de ajuda para estruturar o relatório e o objetivo"
O que foi ajustado manualmente: todo o conteúdo técnico do README, como o passo a passo de uso, a estrutura das abas e os papéis de cada integrante, foi conferido e corrigido pelo grupo.

## Fonte de Dados
Fonte oficial: Resolução CD/FNDE nº 1, de 18 de fevereiro de 2026
Link oficial: https://www.gov.br/fnde/pt-br/acesso-a-informacao/legislacao/resolucoes/2026/resolucao-cd_fnde-no-1-de-18-de-fevereiro-de-2026-dou-imprensa-nacional.pdf/view
O que os dados representam: os valores per capita do PNAE por modalidade de ensino, em reais por dia, vigentes em 2026, e os 200 dias letivos usados no cálculo. A resolução altera a Resolução CD/FNDE nº 6/2020 e traz reajuste médio de 14,35% em relação a 2025. A escola (EMEB Vila Quitaúna, em Osasco/SP) e suas matrículas são fictícias e foram criadas para fins didáticos. As faixas de porte também são didáticas.
Estrutura:
- Parametros_PNAE: valor per capita (R$/dia) por modalidade, dias letivos e faixas de porte.
- Simulador_Escola: matrículas por modalidade, fator de ajuste, matrículas ajustadas e repasse anual estimado.
- Banco_de_Dados: ID, Data/Hora, Fator de Ajuste, Racional da Taxa, Total de Matrículas Ajustadas, Repasse Ajustado (R$) e Usuário.

## Participação do Grupo
O que aprendemos com este projeto: aprendemos a ler e adaptar um código em VBA, entendendo como declarar variáveis, capturar valores digitados nas células e gravá-los em outra aba. Ao implementar o campo Usuário, vimos na prática como uma macro valida o preenchimento antes de salvar. Isso transforma uma planilha de cálculo em um registro organizado das simulações. Também percebemos a importância de guardar a justificativa de cada cenário, e não apenas o resultado, para que a análise possa ser entendida e revisada depois.

Papel de cada integrante:
- Arthur Condão: responsável pelo upload no GitHub.
- João Pedro Bersi: responsável pelo README.
- Kauã Olimar: responsável pelo VBA.
- Roger Alencar: responsável pelo upload no GitHub.
- Thiago Ramos Giacomo: responsável pelo README.
- Tomás Lenci Girard: responsável pelo VBA.
