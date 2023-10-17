from docx2pdf import convert
import os;
import sys


def main(argumento):
    PastaBase = argumento
    Pasta_word = "\\word\\"
    Pasta_PDF ="\\pdf\\"

    PastaBaseWord = PastaBase + Pasta_word
    PastaBasePDF = PastaBase + Pasta_PDF
    lendo_arquivos = os.listdir(PastaBaseWord)
    for item in lendo_arquivos:
        caminhoConvertido = PastaBaseWord + item
        caminhoConvertido2 = PastaBasePDF + item
        caminhoConvertido2 = caminhoConvertido2.replace(".docx",".pdf")
        convert(caminhoConvertido, caminhoConvertido2)

    print("Conversão concluída com sucesso.")
    

if __name__ == "__main__":
    try:
        print("Informar onde o arquivo Base se encontra")
        caminho =input("Infomr o caminho do arquivo base")
        main(caminho)
    except:
        print("Erro no programa")