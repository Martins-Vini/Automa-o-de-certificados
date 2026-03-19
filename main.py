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
            nomeCurso = str(linha[0].value)
            nomeAluno = str(linha[1].value)
            tipoCertificado = str(linha[2].value)
            
            # Tratando as datas (converte para string e pega só a parte da data antes do espaço)
            dataInicio = str(linha[3].value).split(' ')[0]
            dataFim = str(linha[4].value).split(' ')[0]
            dataEmissao = str(linha[6].value).split(' ')[0]

            # Tratando a carga horária: remove o .0 com int() e volta para str()
            cargaHoraria = str(int(linha[5].value))
    
            #Definir fontes de cada atributo mutávelfontNome
            fonteNome = ImageFont.truetype('./fonts/tahomabd.ttf', 90)
            fonteGeral = ImageFont.truetype('./fonts/tahoma.ttf', 46)
    
            # Se nome não for encontrado, pula para o próximo loop
            if not nomeAluno:
                continue
        
            image = Image.open('./certificado_padrao.jpg')
            draw = ImageDraw.Draw(image)
            
            draw.text((1048, 803), nomeAluno, fill="black", font=fonteNome)
            draw.text((1094, 990), nomeCurso, fill="black", font=fonteGeral)
            draw.text((1450, 1093), tipoCertificado, fill="black", font=fonteGeral)
            draw.text((744, 1805), dataInicio, fill="black", font=fonteGeral)
            draw.text((735, 1955), dataFim, fill="black", font=fonteGeral)
            draw.text((1501, 1225), cargaHoraria, fill="black", font=fonteGeral)
            draw.text((2200, 1926), dataEmissao, fill="black", font=fonteGeral)

            image.save('./certificadosGerados/certificado_{}_{}.jpg'.format(index, nomeAluno))
        
    except FileNotFoundError as e:
        print(f"Erro: Certifique-se de que os arquivos de imagem, fonte e planilha existem. {e}")
    except Exception as e:
        print(f"Ocorreu um erro inesperado: {e}")    

main()

