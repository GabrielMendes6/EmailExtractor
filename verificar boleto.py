import pdfplumber

def extract_text_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = ''
        for page in pdf.pages:
            text += page.extract_text()
    return text

pdf_text = extract_text_from_pdf('boleto_com_codigo_de_barras.pdf')
print(pdf_text)
