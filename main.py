""" 

1° -> Pegar os dados da planilha para preencher campos mutáveis de um certificado padrão
2° -> Transferir para a imagem do certificado os dados da planilha

"""

import openpyxl # type: ignore
from PIL import Image, ImageDraw, ImageFont # type: ignore

def main():
    try:
        # Carregar qualquer planilha Excel dentro de uma variável
        workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
        #Pegar página específica da planilha
        pagAlunos = workbook_alunos['Sheet1']
        #Acessar cada linha da planilha e pegar os dados de cada coluna (Com uma iteração)
        for index, linha in enumerate(pagAlunos.iter_rows(min_row=2), start=1):
            #Acessar cada célula da linha e pegar o valor
            nomeCurso = linha[0].value
            nomeAluno = linha[1].value
            tipoCertificado = linha[2].value
            dataInicio = linha[3].value
            dataFim = linha[4].value
            cargaHoraria = linha[5].value
            dataEmissao = linha[6].value
    
            #Definir fontes de cada atributo mutávelfontNome
            fonteNome = ImageFont.truetype('./fonts/tahomabd.ttf', 90)
            fonteGeral = ImageFont.truetype('./fonts/tahoma.ttf', 46)
    
    
            if not nomeAluno:
                continue
        
            image = Image.open('./certificado_padrao.jpg')
            draw = ImageDraw.Draw(image)
        
            draw.text((1048,887), nomeAluno, fill="black", font=fonteNome)
            image.save('./certificado_{}_{}.jpg'.format(index, nomeAluno))
        
    except FileNotFoundError as e:
        print(f"Erro: Certifique-se de que os arquivos de imagem, fonte e planilha existem. {e}")
    except Exception as e:
        print(f"Ocorreu um erro inesperado: {e}")    

main()

