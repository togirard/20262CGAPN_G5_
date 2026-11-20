# Projeto 4: Painel do Censo Escolar 2024 com Power Query – Grupo 5

Portfólio do Grupo 5 na disciplina Análise de Dados para Pesquisas em Políticas Públicas. Este repositório reúne quatro projetos, cada um na sua pasta.

## Portfólio

| Projeto | O que é | Pasta |
|---|---|---|
| Projeto 1: Simulador PNAE | Simulador em HTML do repasse do PNAE, montado a partir de uma planilha Excel | [Acessar Projeto 1](../projeto-1-simulador-pnae/) |
| Projeto 2: Painel Censo | Dashboard em Excel com dados do Censo Escolar 2024 | [Acessar Projeto 2](../projeto-2-painel-censo/) |
| Projeto 3: Simulador PNAE com macro | Simulador em Excel com macro VBA que registra cada simulação | [Acessar Projeto 3](../projeto-3-simulador-pnae-macro/) |
| Projeto 4: Automação com Power Query | Painel do Censo Escolar 2024 de Osasco, tratado com Power Query | [Acessar Projeto 4](../projeto-4-automacao-powerquery/) |

### Estrutura do Projeto 4
- /projeto-4-automacao-powerquery/Projeto 2- 2nda entrega github.xlsx: a planilha com o painel, a base e as consultas do Power Query.
- /projeto-4-automacao-powerquery/README.md: este documento.

## Projeto 4: Painel do Censo Escolar 2024

### Objetivo
Este projeto trata, com Power Query, a base do Censo Escolar 2024 (INEP) e monta no Excel um painel dinâmico das 467 escolas de Osasco (SP). O painel mostra, por porte da escola, a quantidade de escolas e as matrículas femininas e masculinas. Mostra também quantas escolas há em cada dependência administrativa (estadual, municipal e privada). O intuito é permitir que essas diferenças sejam vistas de forma clara e interativa, por meio de tabelas dinâmicas, gráfico dinâmico e segmentações.

### Como usar
Abra o arquivo "Projeto 2- 2nda entrega github.xlsx" no Excel e vá para a aba Controle, que é o painel. À esquerda, em "Filtrar tabela Microdados", ficam as três segmentações (filtros clicáveis): Dependência, Porte da escola e Situação de funcionamento. À direita, em "Tabela Dinâmica", ficam as duas tabelas dinâmicas. O gráfico dinâmico de colunas fica abaixo das segmentações, ligado à primeira tabela. Ao clicar nos botões da segmentação, tabelas e gráfico se atualizam.

Os dados já estão carregados na planilha. Atualizar as consultas exigiria o arquivo nacional do Censo (`Censo_2024_Excel.xlsx`), que não faz parte deste repositório.

| Aba | O que contém |
|---|---|
| Controle | Painel: 3 segmentações, 2 tabelas dinâmicas e 1 gráfico dinâmico. A primeira tabela (K16:O23) agrupa por porte da escola e mostra a contagem de Lixo e de Esgoto (ou seja, de escolas) e a soma de matrículas femininas e masculinas. A segunda (K27:L31) conta escolas por dependência. |
| Microdados | Resultado da consulta Microdados: 467 escolas e 69 colunas. É a fonte das tabelas dinâmicas. |
| Dependencia, Localizacao, LocDiferenciada, Situacao | Tabelas de apoio, uma por consulta, que traduzem códigos em texto (por exemplo, 2 = Estadual, 1 = Urbana, 1 = Ativa). |
| Filtro de Municípios | Resultado da consulta de mesmo nome, com uma linha: SP / Osasco. |
| Planilha1 | Tabela Tabela6 (SP / Osasco), de onde a consulta Filtro de Municípios lê o município escolhido. |

### Prints do resultado
Painel na aba Controle (segmentações, tabelas dinâmicas e gráfico dinâmico)

<img width="1133" height="515" alt="Painel na aba Controle" src="https://github.com/user-attachments/assets/ea6ad3e7-880d-4928-b959-11d269da3143" />
<img width="896" height="534" alt="Tabelas dinâmicas" src="https://github.com/user-attachments/assets/32e57843-e512-4e0f-95cf-f5abc6c699cc" />
<img width="1146" height="498" alt="Gráfico dinâmico" src="https://github.com/user-attachments/assets/1a0afbb0-4e5c-41b2-937a-3580b3f844dd" />

## Fórmulas e Tratamento dos Dados (Power Query)

A planilha não tem fórmulas escritas nas células. Todo o tratamento foi feito em seis consultas do Power Query: Dependencia, Localizacao, LocDiferenciada, Situacao, Filtro de Municípios e Microdados. As quatro primeiras apenas importam as tabelas de apoio do arquivo nacional e definem os tipos de dados. A consulta Filtro de Municípios lê a tabela Tabela6. A consulta Microdados faz o trabalho principal, em cinco etapas.


Etapa 1: importa a aba Microdados do arquivo nacional e define o tipo de cada coluna (texto, número inteiro).


Etapa 2: inner join com a tabela Filtro de Municípios, que reduz a base nacional às escolas de Osasco:

Table.NestedJoin(#"Tipo Alterado", {"SG_UF", "NO_MUNICIPIO"}, #"Filtro de Municípios", {"Estado (Sigla)", "Município"}, "Filtro de Municípios", JoinKind.Inner)


Etapa 3: quatro left joins, um por tabela de apoio, que trazem o nome no lugar do código. Fazem o papel do PROCV. Exemplo com a dependência administrativa:
Table.NestedJoin(#"Filtro Expandido", {"TP_DEPENDENCIA"}, Dependencia, {"TP_DEPENDENCIA"}, "Dependencia", JoinKind.LeftOuter)

As outras três seguem o mesmo modelo, com TP_LOCALIZACAO, TP_LOCALIZACAO_DIFERENCIADA e TP_SITUACAO_FUNCIONAMENTO.


Etapa 4: coluna condicional TAM_ESCOLA, que classifica o porte pelo total de matrículas da educação básica. Faz o papel do SE:

```powerquery
if [QT_MAT_BAS] = null then "Sem Dados"
else if [QT_MAT_BAS] <= 50 then "Micro"
else if [QT_MAT_BAS] <= 200 then "Pequena"
else if [QT_MAT_BAS] <= 500 then "Media"
else if [QT_MAT_BAS] <= 1000 then "Grande"
else if [QT_MAT_BAS] <= 5000 then "Muito Grande"
else "Mega Grande"
```


Etapa 5: quatro colunas condicionais de infraestrutura (Agua, Energia, Esgoto e Lixo). Cada uma lê os indicadores IN_* na ordem de prioridade e devolve o primeiro marcado com 1. Se nenhum estiver marcado, o resultado é "Sem Dados". Exemplo do esgoto:

if [IN_ESGOTO_REDE_PUBLICA] = 1 then "Rede Publica"
else if [IN_ESGOTO_FOSSA_SEPTICA] = 1 then "Fossa Septica"
else if [IN_ESGOTO_FOSSA_COMUM] = 1 then "Fossa Comum"
else if [IN_ESGOTO_FOSSA] = 1 then "Fossa"
else if [IN_ESGOTO_INEXISTENTE] = 1 then "Inexistente"
else "Sem Dados"



As ordens de prioridade das outras três colunas são:
- Agua: rede pública, poço artesiano, cacimba, fonte ou rio, carro-pipa e inexistente.
- Energia: rede pública, gerador fóssil, renovável e inexistente.
- Lixo: serviço de coleta, destino final público, queima, enterra e descarta em outra área.

## Uso de Inteligência Artificial

Ferramenta utilizada: Claude

Para que foi usada: buscar informações e pensar em formas diferentes e adequadas de ilustrar os gráficos do painel.

Exemplo de prompt utilizado: "[Como fazer os dados de distribuição de raça por modalidade escolar ser mais visualmente agradável]"

O que foi ajustado manualmente: títulos dos gráficos e outros ajustes pequenos de formatação.

## Fonte de Dados

Fonte oficial: Censo Escolar 2024 (INEP)

Link oficial: [https://censobasico.inep.gov.br/censobasico_2024/](https://censobasico.inep.gov.br/censobasico_2024/)

O que os dados representam: O Censo Escolar é a principal pesquisa estatística sobre a educação básica no Brasil, realizada pelo INEP com informações declaradas pelas escolas. Cada linha da base é uma escola de Osasco e descreve onde ela fica, quem a administra, se está ativa, sua infraestrutura (água, energia, esgoto e lixo) e suas matrículas por etapa de ensino, sexo, raça/cor e faixa etária. Com todos os filtros liberados, o painel soma 467 escolas (68 estaduais, 150 municipais e 249 privadas), com 86.811 matrículas femininas e 85.780 masculinas.

Estrutura: A aba Microdados tem 467 linhas e 69 colunas, sendo 59 colunas originais do INEP e 10 criadas no Power Query. As principais são:
- NO_MUNICIPIO, SG_UF, NO_ENTIDADE e CO_ENTIDADE: município, estado, nome e código da escola.
- TP_DEPENDENCIA e a coluna criada Dependencia.DEPENDENCIA: código e nome da dependência administrativa (Federal, Estadual, Municipal ou Privada).
- TP_LOCALIZACAO e Localizacao.LOCALIZACAO: código e nome da localização (urbana ou rural).
- TP_LOCALIZACAO_DIFERENCIADA e LocDiferenciada.LOC_DIFERENCIADA: localização diferenciada (assentamento, terra indígena, comunidade quilombola etc.).
- TP_SITUACAO_FUNCIONAMENTO e Situacao.SITUACAO: código e nome da situação (Ativa ou Inativa).
- TAM_ESCOLA, Agua, Energia, Esgoto e Lixo: colunas calculadas descritas acima.
- QT_MAT_BAS_FEM e QT_MAT_BAS_MASC: matrículas por sexo, usadas no painel.
- QT_MAT_BAS, QT_MAT_INF, QT_MAT_FUND, QT_MAT_MED, QT_MAT_PROF, QT_MAT_EJA e QT_MAT_ESP: matrículas totais e por modalidade de ensino.
- QT_MAT_BAS_BRANCA, _PRETA, _PARDA, _AMARELA e _INDIGENA: matrículas por raça/cor.
- QT_MAT_BAS_0_3 até QT_MAT_BAS_18_MAIS: matrículas por faixa etária.

## Participação do Grupo

O que aprendemos com este projeto: na disciplina Análise de Dados para Pesquisas em Políticas Públicas, trabalhamos diretamente com os dados nacionais do Censo Escolar. Aprendemos a tratar uma base grande no Power Query (importar, filtrar por município, mesclar tabelas de apoio e criar colunas condicionais). Também aprendemos a montar tabelas dinâmicas, um gráfico dinâmico e segmentações, e a escolher a melhor forma de comunicar os resultados usando o Excel como ferramenta.

Papel de cada integrante:
- Arthur: responsável pelo upload do projeto no GitHub e pelos gráficos no dashboard.
- Kauã: auxiliou no desenvolvimento da planilha.
- Roger: auxiliou no desenvolvimento das fórmulas e dos gráficos no Excel e desenvolveu as expressões condicionais no Power Query.
- João Pedro: auxiliou no desenvolvimento da comunicação visual das informações no Excel.
- Thiago: auxiliou no desenvolvimento da planilha e do projeto, incluindo a representação visual.
- Tomás: auxiliou no desenvolvimento das fórmulas e dos gráficos no Excel.
