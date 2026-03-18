""" 

1° -> Pegar os dados da planilha para preencher campos mutáveis de um certificado padrão
2° -> Transferir para a imagem do certificado os dados da planilha

"""

import openpyxl # type: ignore
from PIL import Image, ImageDraw, ImageFont # type: ignore

def get_dados_planilha():
    #Carregar qualquer planilha Excel dentro de uma variável
    workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx');
    #Pegar página específica da planilha
    pagAlunos = workbook_alunos['Sheet1'];
    #Acessar cada linha da planilha e pegar os dados de cada coluna (Com uma iteração)
    for linha in pagAlunos.iter_rows(min_row=2):
        #Acessar cada célula da linha e pegar o valor
        nomeCurso = linha[0].value
        nomeAluno = linha[1].value
        tipoCertificado = linha[2].value
        dataInicio = linha[3].value
        dataFim = linha[4].value
        cargaHoraria = linha[5].value
        dataEmissao = linha[6].value

    
    
    
get_dados_planilha()
