# Automação Auxiliar no Lançamento de Notas no Diário de Classe SIEPE

Essa automação visa auxiliar professores e/ou coordenadores da EAD da Secretaria de Educação e Esportes de Pernambuco(SEE-PE), lotados na Escola Técnica Estadual Professor Antônio Carlos Gomes da Costa (ETEPAC), na comparação dos estudantes presentes no SIEPE (https://siepe.educacao.pe.gov.br) com as planilhas de notas obtidas no AVA (https://ead.educacao.pe.gov.br), no processamento das notas conforme a estratégia de quantidade de competências do componente curricular e, por fim, no lançamento automático das notas dos estudantes localizados e que não possuem homônimos no Diário de Classe SIEPE.

Nota: cabe ao professor(a) e/ou ao coordenador(a), a conferência e o salvamento das notas no SIEPE; a automação não salvará automaticamente no SIEPE; é importante que o Diário de Classe a ser manipulado esteja completamente em branco, sem marcações anteriores das caixas de NC (não compareceu); as planilhas da NOA (Novas Oportunidades de Aprendizagem) devem iniciar com o termo "NOA".

## 1. Preparação do Ambiente

##### 1.1. Baixar e instalar o python 3.11 ou versão estável superior (https://www.python.org); 
Nota: antes de instalar o Python, desmarcar a opção "Install launcher for all users (recommended)", marcar a opção "Add Python 3.x to PATH", clicar na opção "Install Now";
##### 1.2. Baixar e instalar o Git (https://git-scm.com/download/windows); 
Nota: não precisa ajustar nenhuma configuração, clicar sempre em "Next", por fim em "Install";
##### 1.3. Instalar as bibliotecas Pandas, Openpyxl e PlayWright, complementos do PlayWright, entrar na unidade C: e realizar o download da Automação, preferencialmente através do PowerShell: "pip install pandas, openpyxl, playwright ; playwright install ; cd \ ; git clone https://github.com/joelsonalves/BOT_SIEPE.git";

## 2. Preparação dos Dados de Entrada

##### 2.1. Acessar o AVA da EAD (https://ead.educacao.pe.gov.br);
##### 2.2. Baixar as planilhas de notas no formato XLSX (Microsoft Excel);
##### 2.3. Mover as planilhas baixadas para a pasta "C:\BOT_SIEPE\PLANILHAS_DE_NOTAS";

## 3. Execução da Automação

##### 3.1. Abrir o PowerShell (console da Microsoft);
##### 3.2. Executar o comando: "cd c:\BOT_SIEPE ; git pull ; python bot.py";
##### 3.3. Escolher a planilha de notas no console;
##### 3.4. Abrir a turma desejada no navegador;
##### 3.5. Expandir o espaço de apontamentos no navegador;
##### 3.6. Escolher o componente curricular desejado no navegador;
##### 3.7. Teclar [Enter] para iniciar o processamento, digitar "TROCAR" seguido de [Enter] para trocar a planilha, digitar "LIMPAR" seguido de [ENTER] para limpar o diário de classe, digitar "NC" seguido de [ENTER] para marcar os não localizados como não compareceram e/ou digitar "SAIR" seguido de [Enter] para encerrar a automação.

Sucesso.

# Ferramenta para Juntar Planilhas XLSX

Essa ferramenta escrita em Python serve para juntar as planilhas de um mesmo componente curricular colocadas na pasta "JUNTAR_PLANILHAS", desde que elas contem com a mesma estrutura de colunas.

## 4. Preparação dos Dados de Entrada

##### 4.1. Acessar o AVA da EAD (https://ead.educacao.pe.gov.br);
##### 4.2. Baixar as planilhas de notas no formato XLSX (Microsoft Excel);
##### 4.3. Mover as planilhas baixadas para a pasta "C:\BOT_SIEPE\JUNTAR_PLANILHAS";

## 5. Execução da Ferramenta

##### 5.1. Abrir o PowerShell (console da Microsoft);
##### 5.2. Executar o comando: "cd c:\BOT_SIEPE ; git pull ; python juntar.py";


