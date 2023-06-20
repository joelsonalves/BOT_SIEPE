# Desenvolvido por Joelson Alves de Melo Junior
# https://github.com/joelsonalves

import pandas as pd
import os

INSTRUCOES = [
    'Colocar as planilhas no formato XLSX dentro da pasta "JUNTAR_PLANILHAS".',
    'Executar a aplicação teclando [ENTER].'
]

OBSERVACOES = [
    'A planilha resultante receberá o nome do arquivo com menos caracteres.',
    'Ela será salva dentro da pasta "PLANILHAS_DE_NOTAS".'
]

PASTA_ENTRADA = 'JUNTAR_PLANILHAS'
PASTA_SAIDA = 'PLANILHAS_DE_NOTAS'
EXTENSAO_DO_ARQUIVO = '.xlsx'

def imprimir_na_tela(texto, lista = []):
    if type(lista) == list and len(lista) == 0:
        print('\n' + '#' * len(texto))
        print(texto)
        print('#' * len(texto))
    else:
        imprimir_na_tela(texto)
        print()
        for i in range(len(lista)):
            print(f'{i + 1}. {lista[i]}')

def capturar_nomes_dos_arquivos():
    arquivos = []
    for diretorio, subpastas, arquivos in os.walk(PASTA_ENTRADA):
        pass
    return arquivos

def selecionar_menor_nome_de_arquivo(arquivos):
    if type(arquivos) == list and len(arquivos) > 0:
        menor = arquivos[0]
        for arquivo in arquivos:
            if len(menor) > len(arquivo):
                menor = arquivo
        return menor
    return None
    
def juntar_planilhas(lista_de_arquivos, nome_arquivo_final):
    if type(lista_de_arquivos) == list and len(lista_de_arquivos) > 1 and type(nome_arquivo_final) == str and len(nome_arquivo_final) > 0:
        df_principal = pd.read_excel(os.path.join(PASTA_ENTRADA, selecionar_menor_nome_de_arquivo(lista_de_arquivos)), dtype=str)
        
        print(f'\n{selecionar_menor_nome_de_arquivo(lista_de_arquivos)}: {df_principal.shape[0]} X {df_principal.shape[1]}')

        lista_aux = lista_de_arquivos.copy()
        lista_aux.remove(selecionar_menor_nome_de_arquivo(lista_de_arquivos))
        for nome_arquivo in lista_aux:
            df_aux = pd.read_excel(os.path.join(PASTA_ENTRADA, nome_arquivo), dtype=str)

            print(f'\n{nome_arquivo}: {df_aux.shape[0]} X {df_aux.shape[1]}')

            df_principal = pd.concat([df_principal, df_aux])

        df_principal = df_principal.reset_index(drop=True)

        print(f'\n{nome_arquivo_final}: {df_principal.shape[0]} X {df_principal.shape[1]}')

        df_principal.to_excel(os.path.join(PASTA_SAIDA, nome_arquivo_final), index=False)

        print('\nNova planilha salva com sucesso.')

if __name__ == '__main__':

    imprimir_na_tela('## FERRAMENTA PARA JUNTAR PLANILHAS XLSX ##')
    imprimir_na_tela('## INSTRUÇÕES ##', INSTRUCOES)
    imprimir_na_tela('## OBSERVAÇÕES ##', OBSERVACOES)

    input('\nTecle [ENTER] para juntar as planilhas...')

    try:

        lista_de_arquivos = capturar_nomes_dos_arquivos()

        if (len(lista_de_arquivos) > 1):

            nome_arquivo_final = selecionar_menor_nome_de_arquivo(lista_de_arquivos).lower().replace(EXTENSAO_DO_ARQUIVO, '')
            nome_arquivo_final += f' [JOIN {len(lista_de_arquivos)} planilhas]{EXTENSAO_DO_ARQUIVO}'

            juntar_planilhas(lista_de_arquivos, nome_arquivo_final)

        else:

            print(f'\nPasta "{PASTA_ENTRADA}" sem arquivos suficientes ou vazia.')

    except Exception:

        print('\nHouve uma falha durante o processamento.')

    except BaseException:

        print('\nHouve uma falha crítica durante o processamento.')