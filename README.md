# Automação Auxiliar no Lançamento de Notas no Diário de Classe SIEPE

Essa automação visa auxiliar professores e/ou coordenadores da EAD da SEE-PE, lotados na ETEPAC, na comparação dos estudantes presentes no SIEPE (https://siepe.educacao.pe.gov.br) com as planilhas de notas obtidas no AVA (https://ead.educacao.pe.gov.br), no processamento das notas conforme a estratégia de quantidade de competências do componente curricular e, por fim, no lançamento automático das notas dos estudantes localizados e que não possuem homônimos no Diário de Classe SIEPE.

Nota: cabe ao professor(a) e/ou ao coordenador(a), a conferência e o salvamento das notas no SIEPE; a automação não salvará automaticamente no SIEPE; é importante que o Diário de Classe a ser manipulado esteja completamente em branco, sem marcações anteriores das caixas de NC (não compareceu).

## 1. Preparação do Ambiente

##### 1.1. Baixar e instalar o python 3.11 ou versão estável superior (https://www.python.org);
##### 1.2. Baixar e instalar o Git (https://git-scm.com/download/windows);
##### 1.2. Instalar as bibliotecas Pandas, Openpyxl e PlayWright, complementos do PlayWright, entrar na unidade C: e realizar o download da Automação, preferencialmente através do PowerShell: "pip install pandas, openpyxl, playwriht ; playwright install ; cd \ ; ";

## 2. Implantação da Automação

##### 2.1. Criar a pasta "BOT_SIEPE" na unidade "C:";
##### 2.2. Copiar o arquivo "bot.py" dentro da pasta "BOT_SIEPE";
##### 2.3. Criar a pasta "PLANILHAS_DE_NOTAS" dentro da pasta "BOT_SIEPE";

## 3. Preparação dos Dados de Entrada

##### 3.1. Acessar o AVA da EAD (https://ead.educacao.pe.gov.br);
##### 3.2. Baixar as planilhas de notas no formato XLSX (Microsoft Excel);
##### 3.3. Mover as planilhas baixadas para a pasta "C:\BOT_SIEPE\PLANILHAS_DE_NOTAS";

## 4. Execução da Automação

##### 4.1. Abrir o PowerShell (console da Microsoft);
##### 4.2. Executar o comando: "cd c:\BOT_SIEPE ; python bot.py"
##### 4.3. Escolher a planilha de notas no console;
##### 4.4. Abrir a turma desejada no navegador;
##### 4.5. Expandir o espaço de apontamentos no navegador;
##### 4.6. Escolher o componente curricular desejado no navegador;
##### 4.7. Teclar [Enter] para iniciar o processamento, digitar "TROCAR" seguido de [Enter] para trocar a planilha e/ou digitar "SAIR" seguido de [Enter] para encerrar a automação.

Sucesso.
