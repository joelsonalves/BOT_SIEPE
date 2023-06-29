from playwright.sync_api import sync_playwright
import pandas as pd
import os

SECRETARIA = 'Secretaria de Educação e Esportes'
SIGLA_SECRETARIA = 'SEE-PE'
GRE = 'Gerência Regional de Educação Recife Norte'
SIGLA_GRE = 'GRE Recife Norte'
ESCOLA = 'Escola Técnica Professor Antônio Carlos Gomes da Costa'
SIGLA_ESCOLA = 'ETEPAC'

EQUIPE_GESTORA = [
    'Arnaldo Luiz da Silva Junior (Diretor Escolar)',
    'Paulo Euzebio Bispo (Assistente de Gestão)',
    'Maria do Rosário Costa Cordouro de Vasconcelos (Chefe de Secretaria)',
    'Gustavo Henrique Tavares Ribeiro (Educador de Apoio)',
]

NOME_DA_AUTOMACAO = 'Automação Auxiliar no Lançamento de Notas no Diário de Classe SIEPE'
VERSAO = '2023.06.29'
DESENVOLVEDOR = 'Joelson Alves de Melo Junior'
CONTATO = 'joelsonjunior@educacao.pe.gov.br'
TESTADORES = {
    '2' : [
        'Dayvson Ricardo Rufino da Silva (ADM 2023.1)',
        'Cristiane Maria de Oliveira (ADM 2023.1)',
    ],
    '3' : [
        'Crisleide Maria da Silva Nascimento Acioly (SEG 2023.1)',
        'Eraldo de Jesus Argolo (SEG 2023.1)',
    ],
    '4' : [
        'Dayvson Ricardo Rufino da Silva (ADM 2023.1)',
        'Cristiane Maria de Oliveira (ADM 2023.1)',
    ],
    '5' : [
        'Aguardando validação',
    ],
    '6' : [
        'Aguardando validação',
    ],
    '7' : [
        'Aguardando validação',
    ]
}

MATRICULA   = 1
NOME        = 0

AT1     = 0
AT2     = 1
RAT1    = 2
RAT2    = 3
NOTA2   = 4

ATIVIDADE_SEMANAL       = 'ATIVIDADE SEMANAL'
AULA_ATIVIDADE          = 'AULA ATIVIDADE'
AVALIACAO_DA_DISCIPLINA = 'AVALIAÇÃO DA DISCIPLINA'

LISTA_PARA_TROCA_DE_CARACTERES = [
    ['A', ['Á', 'À', 'Ã', 'Â', 'Ä']],
    ['E', ['É', 'È', 'Ẽ', 'Ê', 'Ë']],
    ['I', ['Í', 'Ì', 'Ĩ', 'Î', 'Ï']],
    ['O', ['Ó', 'Ò', 'Õ', 'Ô', 'Ö']],
    ['U', ['Ú', 'Ù', 'Ũ', 'Û', 'Ü']],
    ['N', ['Ñ']],
    ['C', ['Ç']]
]

class Bot():

    def __init__(self):
        super().__init__()
        self.__pagina_inicial = 'https://siepe.educacao.pe.gov.br'
        self.__pagina_apos_login = 'https://siepe.educacao.pe.gov.br/Site/'
        self.__pagina_diario_de_classe = 'https://siepe.educacao.pe.gov.br/diarioclasse/DiarioClasse.do'
        self.__lista_de_estudantes_siepe = []
        self.__lista_de_notas_ava = []
        self.__pasta_planilhas_de_notas = 'PLANILHAS_DE_NOTAS'
        self.__arquivo_xlsx = 'planilha_ausente.xlsx'
        self.__numero_de_competencias = 7

    def __selecionar_planilha_de_notas(self):
        arquivos = []
        for diretorio, subpastas, arquivos in os.walk(self.__pasta_planilhas_de_notas):
            pass
        if len(arquivos) > 0:
            print('\n####### LISTA DE PLANILHAS DE NOTAS #######\n')
            arquivos.sort()
            for i in range(len(arquivos)):
                print(str(i + 1).zfill(3) + '\t' + arquivos[i])
            while (True):
                num_arquivo = input('\nDigite o número do arquivo seguido de [ENTER]: ')
                try:
                    int(num_arquivo)
                except:
                    continue
                if (int(num_arquivo) - 1) >= 0 and (int(num_arquivo) - 1) < len(arquivos):
                    self.__arquivo_xlsx = arquivos[int(num_arquivo) - 1]
                    print()
                    break
        else:
            print(f'!!! A PASTA {self.__pasta_planilhas_de_notas.upper()} ESTÁ VAZIA !!!')

    def __ajustar_texto(self, texto):
        texto = texto.strip().replace('  ', ' ').upper()
        for linha in LISTA_PARA_TROCA_DE_CARACTERES:
            for c in linha[1]:
                texto = texto.replace(c, linha[0])
        for c in texto:
            if not ((c >= 'A' and c <= 'Z') or c == ' '):
                texto = texto.replace(c, '') 
        return texto
    
    def __ajustar_nome_turma_siepe(self, texto):
        texto = texto.strip().replace('  ', ' ').upper()
        for linha in LISTA_PARA_TROCA_DE_CARACTERES:
            for c in linha[1]:
                texto = texto.replace(c, linha[0])
        for c in texto:
            if not ((c >= 'A' and c <= 'Z') or (c >= '0' and c <= '9') or c == ' '):
                texto = texto.replace(c, '') 
        texto = texto.lower().replace(' ', '_')
        return texto
    
    def __verificar_se_o_navegador_ainda_esta_funcional(self, page):
        try:
            page.title()
        except BaseException:
            return False
        return True

    def __fazer_login(self, page):
        page.goto(self.__pagina_inicial)   
        print('Aguardando login no SIEPE...')
        
        while (True):
            page.wait_for_timeout(1000)
            if (page.url.find(self.__pagina_apos_login) == 0):
                break
        
        page.goto(self.__pagina_diario_de_classe)

    def __extrair_lista_de_estudantes(self, page):
        
        self.__lista_de_estudantes_siepe = page.evaluate(''' () => { 
        
            var lista = []; 
            document.querySelectorAll('table.TabelaDiarioClasse a').forEach((linha) => { 
                if (linha.innerText !== '') lista.push([linha.innerText]); 
            }); 
            var i = 0; 
            var arr_input = document.querySelectorAll('table.TabelaDiarioClasse input'); 
            for (var j = 0; j < arr_input.length; j++) { 
                if (arr_input[j].id.includes('nota_1_')) { 
                    let aux = arr_input[j].id.replace('nota_1_',''); 
                    lista[i].push(aux);
                    i++; 
                } 
            }
            return lista;

        } ''')

        for i in range(len(self.__lista_de_estudantes_siepe)):
            self.__lista_de_estudantes_siepe[i][0] = self.__ajustar_texto(self.__lista_de_estudantes_siepe[i][0])

    def __extrair_lista_de_estudantes_noa(self, page):
        
        self.__lista_de_estudantes_siepe = page.evaluate(''' () => { 
        
            var lista = []; 
            document.querySelectorAll('table.TabelaDiarioClasse a').forEach((linha) => { 
                if (linha.innerText !== '') lista.push([linha.innerText]); 
            }); 
            var i = 0; 
            var arr_input = document.querySelectorAll('table.TabelaDiarioClasse input'); 
            for (var j = 0; j < arr_input.length; j++) { 
                if (arr_input[j].id.includes('nota_')) { 
                    let aux = arr_input[j].id.replace('nota_',''); 
                    lista[i].push(aux);
                    i++; 
                } 
            }
            return lista;

        } ''')

        for i in range(len(self.__lista_de_estudantes_siepe)):
            self.__lista_de_estudantes_siepe[i][0] = self.__ajustar_texto(self.__lista_de_estudantes_siepe[i][0])

    def __verificar_numero_de_competencias(self):
        
        df_resultados_ava = pd.read_excel(os.path.join(self.__pasta_planilhas_de_notas, self.__arquivo_xlsx), dtype=str)
        # Colocar colunas em maiúsculo e remover excesso de espaço
        for nome_coluna in df_resultados_ava.columns:
            df_resultados_ava.rename(columns={nome_coluna: nome_coluna.upper().strip()}, inplace=True)
        localizado = False
        for competencia in range (7, 0, -1):

            for nome_coluna in df_resultados_ava:
                if nome_coluna.find(f'{ATIVIDADE_SEMANAL} {competencia}') > -1:
                    self.__numero_de_competencias = competencia
                    localizado = True
                    break
            if (localizado):
                break

    def __comparar_lista_de_estudantes_e_processar_notas(self):

        if len(self.__lista_de_estudantes_siepe) > 0:

            self.__lista_de_notas_ava = []
            for i in range(len(self.__lista_de_estudantes_siepe)):
                # [AT1, AT2, RAT1, RAT2, NOTA2] = ['', '', '', '', '']
                self.__lista_de_notas_ava.append([''] * 5)

            df_resultados_ava = pd.read_excel(os.path.join(self.__pasta_planilhas_de_notas, self.__arquivo_xlsx), dtype=str)

            # Colocar colunas em maiúsculo e remover excesso de espaço
            for nome_coluna in df_resultados_ava.columns:
                df_resultados_ava.rename(columns={nome_coluna: nome_coluna.upper().strip()}, inplace=True)

            # Ajustar nomes das colunas
            for competencia in range(1,8):
                for nome_coluna in df_resultados_ava.columns:
                    texto_aux = f'{ATIVIDADE_SEMANAL} {competencia}'
                    if nome_coluna.find(texto_aux) >= 0: 
                        df_resultados_ava.rename(columns={nome_coluna: texto_aux}, inplace=True)
                        continue
                    texto_aux = f'{AULA_ATIVIDADE} {competencia}'
                    if nome_coluna.find(texto_aux) >= 0:  
                        df_resultados_ava.rename(columns={nome_coluna: texto_aux}, inplace=True)
                        continue
                    texto_aux = f'{AVALIACAO_DA_DISCIPLINA}'
                    if nome_coluna.find(texto_aux) >= 0:  
                        df_resultados_ava.rename(columns={nome_coluna: texto_aux}, inplace=True)

            # Ajustar nome dos estudantes no DataFrame
            for i in df_resultados_ava.index:
                df_resultados_ava.loc[i, 'NOME'] = self.__ajustar_texto(df_resultados_ava.loc[i, 'NOME'])

                if df_resultados_ava.loc[i, 'NOME'].find('ALUNO') == 0:
                    df_resultados_ava.loc[i, 'NOME'] = self.__ajustar_texto(df_resultados_ava.loc[i, 'SOBRENOME'])

            # Localizar estudantes e processar notas
            for i in range(len(self.__lista_de_estudantes_siepe)):
                linha = df_resultados_ava.loc[df_resultados_ava['NOME']==self.__lista_de_estudantes_siepe[i][NOME]].reset_index(drop=True)
                quant_linhas = linha.shape[0]
                if quant_linhas > 1:
                    self.__lista_de_notas_ava[i][AT1] = 'HOMONIMO'
                elif linha.shape[0] == 1:
                    
                    # Primeira Nota
                    # Componentes curriculares com mais de 3 competências
                    if self.__numero_de_competencias >= 4:
                        # Verificar se os campos da Atividade 1 estão vazios ou não
                        semana = 1
                        if linha.loc[0, f'{ATIVIDADE_SEMANAL} {semana}'] == '-' and linha.loc[0, f'{AULA_ATIVIDADE} {semana}'] == '-':
                            self.__lista_de_notas_ava[i][AT1] = 'NC'
                        elif linha.loc[0, f'{ATIVIDADE_SEMANAL} {semana}'] == '-':
                            self.__lista_de_notas_ava[i][AT1] = float(linha.loc[0, f'{AULA_ATIVIDADE} {semana}']) / 2
                        elif linha.loc[0, f'{AULA_ATIVIDADE} {semana}'] == '-':
                            self.__lista_de_notas_ava[i][AT1] = float(linha.loc[0, f'{ATIVIDADE_SEMANAL} {semana}']) / 2
                        else:
                            self.__lista_de_notas_ava[i][AT1] = (float(linha.loc[0, f'{ATIVIDADE_SEMANAL} {semana}']) + float(linha.loc[0, f'{AULA_ATIVIDADE} {semana}'])) / 2
                        semana = 2
                        if linha.loc[0, f'{ATIVIDADE_SEMANAL} {semana}'] == '-' and linha.loc[0, f'{AULA_ATIVIDADE} {semana}'] == '-':
                            self.__lista_de_notas_ava[i][AT2] = 'NC'
                        elif linha.loc[0, f'{ATIVIDADE_SEMANAL} {semana}'] == '-':
                            self.__lista_de_notas_ava[i][AT2] = float(linha.loc[0, f'{AULA_ATIVIDADE} {semana}']) / 2
                        elif linha.loc[0, f'{AULA_ATIVIDADE} {semana}'] == '-':
                            self.__lista_de_notas_ava[i][AT2] = float(linha.loc[0, f'{ATIVIDADE_SEMANAL} {semana}']) / 2
                        else:
                            self.__lista_de_notas_ava[i][AT2] = (float(linha.loc[0, f'{ATIVIDADE_SEMANAL} {semana}']) + float(linha.loc[0, f'{AULA_ATIVIDADE} {semana}'])) / 2

                    # Componentes curriculares com menos de 4 competências
                    else:
                        semana = 1
                        if linha.loc[0, f'{ATIVIDADE_SEMANAL} {semana}'] == '-' and linha.loc[0, f'{AULA_ATIVIDADE} {semana}'] == '-':
                            self.__lista_de_notas_ava[i][AT1] = 'NC'
                            self.__lista_de_notas_ava[i][AT2] = 'NC'
                        elif linha.loc[0, f'{AULA_ATIVIDADE} {semana}'] == '-':
                            self.__lista_de_notas_ava[i][AT1] = 'NC'
                            self.__lista_de_notas_ava[i][AT2] = float(linha.loc[0, f'{ATIVIDADE_SEMANAL} {semana}'])
                        elif linha.loc[0, f'{ATIVIDADE_SEMANAL} {semana}'] == '-':
                            self.__lista_de_notas_ava[i][AT1] = float(linha.loc[0, f'{AULA_ATIVIDADE} {semana}'])
                            self.__lista_de_notas_ava[i][AT2] = 'NC'
                        else:
                            self.__lista_de_notas_ava[i][AT1] = float(linha.loc[0, f'{AULA_ATIVIDADE} {semana}'])
                            self.__lista_de_notas_ava[i][AT2] = float(linha.loc[0, f'{ATIVIDADE_SEMANAL} {semana}'])
                    
                    # Segunda Nota
                    # Componente curricular com 2 competências
                    if self.__numero_de_competencias == 2:
                        semana = 2
                        if linha.loc[0, f'{ATIVIDADE_SEMANAL} {semana}'] == '-' and linha.loc[0, f'{AULA_ATIVIDADE} {semana}'] == '-': 
                            self.__lista_de_notas_ava[i][NOTA2] = 'NC'
                        elif linha.loc[0, f'{AULA_ATIVIDADE} {semana}'] == '-':
                            self.__lista_de_notas_ava[i][NOTA2] = float(linha.loc[0, f'{ATIVIDADE_SEMANAL} {semana}'])
                        elif linha.loc[0, f'{ATIVIDADE_SEMANAL} {semana}'] == '-':
                            self.__lista_de_notas_ava[i][NOTA2] = float(linha.loc[0, f'{AULA_ATIVIDADE} {semana}'])
                        else:
                            self.__lista_de_notas_ava[i][NOTA2] = float(linha.loc[0, f'{AULA_ATIVIDADE} {semana}']) + float(linha.loc[0, f'{ATIVIDADE_SEMANAL} {semana}'])
                    
                    elif self.__numero_de_competencias == 3:
                        semanaInicial = 2
                        semanaFinal = 3
                    elif self.__numero_de_competencias == 4:
                        semanaInicial = 3
                        semanaFinal = 4
                    elif self.__numero_de_competencias == 5:
                        semanaInicial = 3
                        semanaFinal = 5
                    elif self.__numero_de_competencias == 6:
                        semanaInicial = 3
                        semanaFinal = 6
                    elif self.__numero_de_competencias == 7:
                        semanaInicial = 3
                        semanaFinal = 7
                    
                    # Cálculo da Segunda Nota para componentes curriculares com mais de 2 competências
                    if self.__numero_de_competencias > 2:
                        soma = 0.0
                        nao_compareceu = True
                        for semana in range(semanaInicial, semanaFinal + 1):
                            if linha.loc[0, f'{ATIVIDADE_SEMANAL} {semana}'] == '-' and linha.loc[0, f'{AULA_ATIVIDADE} {semana}'] == '-': 
                                soma += 0.0
                            else:
                                if nao_compareceu:
                                    nao_compareceu = False

                                if linha.loc[0, f'{AULA_ATIVIDADE} {semana}'] == '-':
                                    soma += float(linha.loc[0, f'{ATIVIDADE_SEMANAL} {semana}'])
                                elif linha.loc[0, f'{ATIVIDADE_SEMANAL} {semana}'] == '-':
                                    soma += float(linha.loc[0, f'{AULA_ATIVIDADE} {semana}'])
                                else:
                                    soma += (float(linha.loc[0, f'{AULA_ATIVIDADE} {semana}']) + float(linha.loc[0, f'{ATIVIDADE_SEMANAL} {semana}']))

                        if nao_compareceu:
                            self.__lista_de_notas_ava[i][NOTA2] = 'NC'
                        else:
                            self.__lista_de_notas_ava[i][NOTA2] = soma / (semanaFinal - semanaInicial + 1)

                    if linha.loc[0, f'{AVALIACAO_DA_DISCIPLINA}'] == '-':
                        self.__lista_de_notas_ava[i][RAT1] = 'NC'
                        self.__lista_de_notas_ava[i][RAT2] = 'NC'

                    else:
                        self.__lista_de_notas_ava[i][RAT1] = float(linha.loc[0, f'{AVALIACAO_DA_DISCIPLINA}']) / 2
                        self.__lista_de_notas_ava[i][RAT2] = float(linha.loc[0, f'{AVALIACAO_DA_DISCIPLINA}']) / 2
                        if self.__lista_de_notas_ava[i][NOTA2] == 'NC' or (float(linha.loc[0, f'{AVALIACAO_DA_DISCIPLINA}']) > self.__lista_de_notas_ava[i][4]):
                            self.__lista_de_notas_ava[i][NOTA2] = float(linha.loc[0, f'{AVALIACAO_DA_DISCIPLINA}'])
                        
                else:
                    self.__lista_de_notas_ava[i][AT1] = 'NAO_LOCALIZADO'
                
            df_resultados_ava = None

    def __comparar_lista_de_estudantes_e_processar_notas_da_noa(self):

        if len(self.__lista_de_estudantes_siepe) > 0:

            self.__lista_de_notas_ava = []
            for i in range(len(self.__lista_de_estudantes_siepe)):
                self.__lista_de_notas_ava.append('')

            df_resultados_ava = pd.read_excel(os.path.join(self.__pasta_planilhas_de_notas, self.__arquivo_xlsx), dtype=str)

            # Colocar colunas em maiúsculo e remover excesso de espaço
            for nome_coluna in df_resultados_ava.columns:
                df_resultados_ava.rename(columns={nome_coluna: nome_coluna.upper().strip()}, inplace=True)

            # Ajustar nomes das colunas
            for competencia in range(1,8):
                for nome_coluna in df_resultados_ava.columns:
                    texto_aux = 'NOA'
                    if nome_coluna.find(f'({texto_aux})') >= 0: 
                        df_resultados_ava.rename(columns={nome_coluna: texto_aux}, inplace=True)
                        break

            # Ajustar nome dos estudantes no DataFrame
            for i in df_resultados_ava.index:
                df_resultados_ava.loc[i, 'NOME'] = self.__ajustar_texto(df_resultados_ava.loc[i, 'NOME'])

                if df_resultados_ava.loc[i, 'NOME'].find('ALUNO') == 0:
                    df_resultados_ava.loc[i, 'NOME'] = self.__ajustar_texto(df_resultados_ava.loc[i, 'SOBRENOME'])

            # Localizar estudantes e processar notas
            for i in range(len(self.__lista_de_estudantes_siepe)):
                linha = df_resultados_ava.loc[df_resultados_ava['NOME']==self.__lista_de_estudantes_siepe[i][NOME]].reset_index(drop=True)
                quant_linhas = linha.shape[0]
                if quant_linhas > 1:
                    self.__lista_de_notas_ava[i] = 'HOMONIMO'
                elif linha.shape[0] == 1:
                    
                    # Verificar se o campo da NOA está vazio ou não
                    if linha.loc[0, 'NOA'] == '-':
                        self.__lista_de_notas_ava[i] = 'NC'
                    else:
                        self.__lista_de_notas_ava[i] = float(linha.loc[0, 'NOA'])  
                        
                else:
                    self.__lista_de_notas_ava[i] = 'NAO_LOCALIZADO'
                
            df_resultados_ava = None
    
    def __salvar_arquivo_csv(self, nome_arquivo_turma_siepe, lista_de_estudantes_com_matriculas):
        # Sobrescrever o arquivo com o nome da turma no SIEPE com a lista de homônimos ou não localizados
        with open(os.path.join(self.__pasta_planilhas_de_notas, nome_arquivo_turma_siepe), 'w') as arquivo:
            arquivo.write('"MATRÍCULA","NOME"\n')
            for linha in lista_de_estudantes_com_matriculas:
                arquivo.write(f'"{linha[1]}","{linha[0]}"\n')

    def __capturar_nome_da_turma_no_siepe(self, page):
        nome_turma_siepe = page.evaluate(''' () => { 
        
            return document.querySelectorAll('.fieldsetBase .TabelaDiarioClasse td')[4].innerText

        } ''')
        return self.__ajustar_nome_turma_siepe(nome_turma_siepe)

    def __preencher_notas_no_siepe(self, page):

        homonimos = []
        nao_localizados = []

        for i in range(len(self.__lista_de_notas_ava)):

            matricula   = self.__lista_de_estudantes_siepe[i][MATRICULA]

            nota_at1    = self.__lista_de_notas_ava[i][AT1]
            nota_at2    = self.__lista_de_notas_ava[i][AT2]
            nota_rat1   = self.__lista_de_notas_ava[i][RAT1]
            nota_rat2   = self.__lista_de_notas_ava[i][RAT2]
            nota_2      = self.__lista_de_notas_ava[i][NOTA2]

            if not nota_at1 in ['HOMONIMO','NAO_LOCALIZADO']:
                # AT1
                if nota_at1 == 'NC':
                    page.locator(f'input#chkNaoCompareceu_1_{matricula}').click()
                else:
                    page.locator(f'input#nota_1_{matricula}').fill(str(nota_at1).replace('.',','))

                # AT2
                if nota_at2 == 'NC':
                    page.locator(f'input#chkNaoCompareceu_2_{matricula}').click()
                else:
                    page.locator(f'input#nota_2_{matricula}').fill(str(nota_at2).replace('.',','))

                # RAT1
                if nota_rat1 == 'NC':
                    page.locator(f'input#chkNaoCompareceuRP_1_{matricula}').click()
                else:
                    page.locator(f'input#notaRec_1_{matricula}').fill(str(nota_rat1).replace('.',','))

                # RAT2
                if nota_rat1 == 'NC':
                    page.locator(f'input#chkNaoCompareceuRP_2_{matricula}').click()
                else:
                    page.locator(f'input#notaRec_2_{matricula}').fill(str(nota_rat2).replace('.',','))

                # NOTA2
                if nota_2 == 'NC':
                    page.locator(f'input#chkNaoCompareceu_7_{matricula}').click()
                else:
                    page.locator(f'input#nota_7_{matricula}').fill(str(nota_2).replace('.',','))               
                
            elif nota_at1 == 'HOMONIMO':
                homonimos.append(i)
            else:
                nao_localizados.append(i)

        # Pressiona Tab
        page.keyboard.press('Tab')

        # Volta para a parte superior da página
        page.keyboard.press('Home') 

        # Capturar o nome da turma
        nome_turma_siepe = self.__capturar_nome_da_turma_no_siepe(page)

        print('\n+++++++ LISTA DE HOMÔNIMOS +++++++')
        for i in homonimos:

            matricula   = self.__lista_de_estudantes_siepe[i][MATRICULA]
            nome        = self.__lista_de_estudantes_siepe[i][NOME]

            print(f'Estudante nº {str(i + 1).zfill(3)} | {nome} ({matricula})')
        print(f'TOTAL: {len(homonimos)} estudante(s)')

        # Salvar lista de homônimos, se houver
        lista_de_estudantes_com_matriculas = []
        for pos in homonimos:
            lista_de_estudantes_com_matriculas.append(self.__lista_de_estudantes_siepe[pos])
        self.__salvar_arquivo_csv(f'{nome_turma_siepe}_homonimos.csv', lista_de_estudantes_com_matriculas)

        print('\n------- LISTA DE NÃO LOCALIZADOS -------')
        for i in nao_localizados:

            matricula   = self.__lista_de_estudantes_siepe[i][MATRICULA]
            nome        = self.__lista_de_estudantes_siepe[i][NOME]

            print(f'Estudante nº {str(i + 1).zfill(3)} | {nome} ({matricula})')
        print(f'TOTAL: {len(nao_localizados)} estudante(s)')

        # Salvar lista de não localizados, se houver
        lista_de_estudantes_com_matriculas = []
        for pos in nao_localizados:
            lista_de_estudantes_com_matriculas.append(self.__lista_de_estudantes_siepe[pos])
        self.__salvar_arquivo_csv(f'{nome_turma_siepe}_nao_localizados.csv', lista_de_estudantes_com_matriculas)

    def __preencher_noa_no_siepe(self, page):

        homonimos = []
        nao_localizados = []

        for i in range(len(self.__lista_de_notas_ava)):

            matricula   = self.__lista_de_estudantes_siepe[i][MATRICULA]
            nota_noa    = self.__lista_de_notas_ava[i]

            if not nota_noa in ['HOMONIMO','NAO_LOCALIZADO']:
                
                if nota_noa == 'NC':
                    page.locator(f'input#chkNaoCompareceuBimestre_{matricula}').click()
                else:
                    page.locator(f'input#nota_{matricula}').fill(str(nota_noa).replace('.',','))          
                
            elif nota_noa == 'HOMONIMO':
                homonimos.append(i)
            else:
                nao_localizados.append(i)

        # Pressiona Tab
        page.keyboard.press('Tab')

        # Volta para a parte superior da página
        page.keyboard.press('Home') 

        # Capturar o nome da turma
        nome_turma_siepe = self.__capturar_nome_da_turma_no_siepe(page)

        print('\n+++++++ LISTA DE HOMÔNIMOS +++++++')
        for i in homonimos:

            matricula   = self.__lista_de_estudantes_siepe[i][MATRICULA]
            nome        = self.__lista_de_estudantes_siepe[i][NOME]

            print(f'Estudante nº {str(i + 1).zfill(3)} | {nome} ({matricula})')
        print(f'TOTAL: {len(homonimos)} estudante(s)')

        # Salvar lista de homônimos, se houver
        lista_de_estudantes_com_matriculas = []
        for pos in homonimos:
            lista_de_estudantes_com_matriculas.append(self.__lista_de_estudantes_siepe[pos])
        self.__salvar_arquivo_csv(f'{nome_turma_siepe}_homonimos.csv', lista_de_estudantes_com_matriculas)

        print('\n------- LISTA DE NÃO LOCALIZADOS -------')
        for i in nao_localizados:

            matricula   = self.__lista_de_estudantes_siepe[i][MATRICULA]
            nome        = self.__lista_de_estudantes_siepe[i][NOME]

            print(f'Estudante nº {str(i + 1).zfill(3)} | {nome} ({matricula})')
        print(f'TOTAL: {len(nao_localizados)} estudante(s)')

        # Salvar lista de não localizados, se houver
        lista_de_estudantes_com_matriculas = []
        for pos in nao_localizados:
            lista_de_estudantes_com_matriculas.append(self.__lista_de_estudantes_siepe[pos])
        self.__salvar_arquivo_csv(f'{nome_turma_siepe}_nao_localizados.csv', lista_de_estudantes_com_matriculas)

    def __limpar_diario_de_classe(self, page):

        print('\nAguarde o processamento...')

        # Injetar um JS para verificar se o checkbox está marcado,
        # caso positivo, clicar nele uma vez,
        # caso negativo, clivar nele duas vezes.
        page.evaluate(''' () => { 
        
            document.querySelectorAll('input[type="checkbox"]').forEach((checkbox) => {
                if (checkbox.checked) checkbox.click();
                else for (let i = 0; i < 2; i++) checkbox.click();
            });

        } ''')

    def __preencher_nc_nos_nao_localizados(self, page):

        for i in range(len(self.__lista_de_notas_ava)):

            matricula   = self.__lista_de_estudantes_siepe[i][MATRICULA]
            nome        = self.__lista_de_estudantes_siepe[i][NOME]
            nota        = self.__lista_de_notas_ava[i][AT1]

            if nota == 'NAO_LOCALIZADO':

                # AT1
                page.locator(f'input#chkNaoCompareceu_1_{matricula}').click()

                # AT2
                page.locator(f'input#chkNaoCompareceu_2_{matricula}').click()

                # RAT1
                page.locator(f'input#chkNaoCompareceuRP_1_{matricula}').click()

                # RAT2
                page.locator(f'input#chkNaoCompareceuRP_2_{matricula}').click()

                # NOTA2
                page.locator(f'input#chkNaoCompareceu_7_{matricula}').click() 

                print(f'Estudante nº {str(i + 1).zfill(3)} {nome} ({matricula}) marcado(a) como não compareceu.')

        # Pressiona Tab
        page.keyboard.press('Tab')

        # Volta para a parte superior da página
        page.keyboard.press('Home')

    def __preencher_nc_nos_nao_localizados_na_noa(self, page):
        
        for i in range(len(self.__lista_de_notas_ava)):

            matricula   = self.__lista_de_estudantes_siepe[i][MATRICULA]
            nome        = self.__lista_de_estudantes_siepe[i][NOME]
            nota        = self.__lista_de_notas_ava[i]

            if nota == 'NAO_LOCALIZADO':

                # NOA
                page.locator(f'input#chkNaoCompareceuBimestre_{matricula}').click() 

                print(f'Estudante nº {str(i + 1).zfill(3)} {nome} ({matricula}) marcado(a) como não compareceu.')

        # Pressiona Tab
        page.keyboard.press('Tab')

        # Volta para a parte superior da página
        page.keyboard.press('Home')

    def run():

        falha_critica = False

        with sync_playwright() as p:

            separador = '#' * 2
            print(f'''\n{separador} {SECRETARIA.upper()} | {SIGLA_SECRETARIA}''')
            print(f'''{separador} {GRE.upper()} | {SIGLA_GRE}''')
            print(f'''{separador} {ESCOLA.upper()} | {SIGLA_ESCOLA}''')
            print(f'''\n{'#' * 73}\n{separador} {NOME_DA_AUTOMACAO.upper()} {separador}''')
            print(f'''{'#' * 73}\n\n{separador} Versão: {VERSAO}''')
            print(f'''{separador} Desenvolvedor: {DESENVOLVEDOR}''')
            print(f'''{separador} E-mail: {CONTATO}''')
            separador = None

            bot = Bot()

            bot.__selecionar_planilha_de_notas()

            browser = p.chromium.launch(headless=False)
            context = browser.new_context()
            context.clear_cookies()
            page = context.new_page()

            try:

                bot.__fazer_login(page)

            except Exception:

                if not bot.__verificar_se_o_navegador_ainda_esta_funcional(page):

                    print('!!! HOUVE UMA FALHA CRÍTICA NO LOGIN !!!\n')
                    falha_critica = True

                else:

                    print('!!! HOUVE UMA FALHA NO LOGIN !!!\n')

            except BaseException:

                print('!!! HOUVE UMA FALHA CRÍTICA NO LOGIN !!!\n')
                falha_critica = True

            sequencia_de_processamento = 1

            while not falha_critica:

                if not bot.__verificar_se_o_navegador_ainda_esta_funcional(page):

                    print('!!! HOUVE UMA FALHA CRÍTICA NO PROCESSAMENTO !!!\n')
                    falha_critica = True
                    break

                print(f'\nSequência de Processamento: {str(sequencia_de_processamento).zfill(4)}\n')
                print(f'Planilha selecionada: {bot.__arquivo_xlsx}\n')
                print('(a) Pressione [ENTER] para processar uma turma.')
                print('(b) Digite "LIMPAR" seguido de [ENTER] para limpar o diário de classe.')
                print('(c) Digite "TROCAR" seguido de [ENTER] para alterar a planilha.')
                print('(d) Digite "NC" seguido de [ENTER] para marcar que os não localizados não compareceram.')
                print('(e) Digite "SAIR" seguido de [ENTER] para encerrar.\n')
                
                entrada = input('O que você deseja agora? ')

                if entrada.upper() == 'TROCAR':

                    bot.__selecionar_planilha_de_notas()

                elif entrada.upper() == 'LIMPAR':

                    try:

                        bot.__limpar_diario_de_classe(page)

                    except Exception:

                        if not bot.__verificar_se_o_navegador_ainda_esta_funcional(page):

                            print('!!! HOUVE UMA FALHA CRÍTICA NO LOGIN !!!\n')
                            falha_critica = True
                            break

                        else:

                            print('!!! HOUVE UMA FALHA NO PROCESSAMENTO !!!\n')

                    except BaseException:

                        print('!!! HOUVE UMA FALHA CRÍTICA NO PROCESSAMENTO !!!\n')
                        falha_critica = True
                        break

                elif entrada.upper() == 'NC':

                    eh_noa = bot.__arquivo_xlsx.strip().lower().find('noa') == 0

                    try:

                        if not eh_noa:

                            bot.__verificar_numero_de_competencias()
                            bot.__extrair_lista_de_estudantes(page)
                            bot.__comparar_lista_de_estudantes_e_processar_notas()
                            bot.__preencher_nc_nos_nao_localizados(page)

                        else:

                            bot.__extrair_lista_de_estudantes_noa(page)
                            bot.__comparar_lista_de_estudantes_e_processar_notas_da_noa()
                            bot.__preencher_nc_nos_nao_localizados_na_noa(page)

                    except Exception:

                        if not bot.__verificar_se_o_navegador_ainda_esta_funcional(page):

                            print('!!! HOUVE UMA FALHA CRÍTICA NO LOGIN !!!\n')
                            falha_critica = True
                            break

                        else:

                            print('!!! HOUVE UMA FALHA NO PROCESSAMENTO !!!\n')

                    except BaseException:

                        print('!!! HOUVE UMA FALHA CRÍTICA NO PROCESSAMENTO !!!\n')
                        falha_critica = True
                        break

                    sequencia_de_processamento += 1


                elif entrada == '':

                    eh_noa = bot.__arquivo_xlsx.strip().lower().find('noa') == 0

                    try:

                        if not eh_noa:

                            bot.__verificar_numero_de_competencias()
                            bot.__extrair_lista_de_estudantes(page)
                            bot.__comparar_lista_de_estudantes_e_processar_notas()
                            bot.__preencher_notas_no_siepe(page)

                        else:

                            bot.__extrair_lista_de_estudantes_noa(page)
                            bot.__comparar_lista_de_estudantes_e_processar_notas_da_noa()
                            bot.__preencher_noa_no_siepe(page)

                    except Exception:

                        if not bot.__verificar_se_o_navegador_ainda_esta_funcional(page):

                            print('!!! HOUVE UMA FALHA CRÍTICA NO LOGIN !!!\n')
                            falha_critica = True
                            break

                        else:

                            print('!!! HOUVE UMA FALHA NO PROCESSAMENTO !!!\n')

                    except BaseException:

                        print('!!! HOUVE UMA FALHA CRÍTICA NO PROCESSAMENTO !!!\n')
                        falha_critica = True
                        break

                    sequencia_de_processamento += 1

                elif entrada.upper() == 'SAIR':

                    break

                else:

                    continue
                    
            if not falha_critica:
                print('\nFinalizando Bot...')

                try:

                    page.set_default_timeout(5000)
                    page.close()
                    browser.close()

                except Exception:

                    print('!!! HOUVE UMA FALHA NA FINALIZAÇÃO !!!\n')

                except BaseException:

                    print('!!! HOUVE UMA FALHA CRÍTICA NA FINALIZAÇÃO !!!\n')

            page = None
            context = None
            browser = None
            bot = None

            separador = '#' * 2
            print(f'''\n{'#' * 73}\n{separador} {NOME_DA_AUTOMACAO.upper()} {separador}''')
            print(f'''{'#' * 73}\n\n{separador} Versão: {VERSAO}\n{separador}''')
            print(f'{separador} CRÉDITOS:')
            print(f'''{separador}\n{separador} {SECRETARIA.upper()} | {SIGLA_SECRETARIA}''')
            print(f'''{separador} {GRE.upper()} | {SIGLA_GRE}''')
            print(f'''{separador} {ESCOLA.upper()} | {SIGLA_ESCOLA}''')
            print(f'''{separador}\n{separador} Equipe Gestora:''')
            for nome in EQUIPE_GESTORA:
                print(f'''{separador} \t {nome}''')
            print(f'''{separador}\n{separador} Desenvolvedor: \n{separador} \t {DESENVOLVEDOR}''')
            for num_competencias in range(2,8):
                print(f'''{separador}\n{separador} Equipe de Testadores e Validadores de Notas de {num_competencias} Competências:''')
                for nome in TESTADORES[f'''{str(num_competencias)}''']:
                    print(f'''{separador} \t {nome}''')
            print('\nAutomação encerrada.')  

if __name__ == '__main__':

    Bot.run()
