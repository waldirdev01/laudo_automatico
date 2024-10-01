import PyPDF2

def ler_ocorrencia(pdf_file):
    with open(pdf_file, 'rb') as file:
        leitor = PyPDF2.PdfReader(file)
        num_paginas = len(leitor.pages)
        texto = ""
        for pagina in range(num_paginas):
            texto += leitor.pages[pagina].extract_text()
        return texto
