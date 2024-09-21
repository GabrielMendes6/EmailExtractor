from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.graphics.barcode import code128

def generate_boleto(filename, data):
    # Criar um novo arquivo PDF
    c = canvas.Canvas(filename, pagesize=letter)

    # Definir o título do boleto
    c.setFont("Helvetica", 16)
    c.drawString(100, 750, "Boleto Bancário")

    # Adicionar informações do boleto
    c.setFont("Helvetica", 12)
    y = 700
    for key, value in data.items():
        text = f"{key}: {value}"
        c.drawString(100, y, text)
        y -= 20

    # Adicionar código de barras
    barcode = code128.Code128(data["Codigo de Barras"])
    barcode.drawOn(c, 100, 580)

    # Salvar o arquivo PDF
    c.save()

# Dados do boleto
boleto_data = {
    "Nome": "Gabriel",
    "CPF/CNPJ": "123.456.789-00",
    "Valor": "R$ 100,00",
    "Vencimento": "10/10/2024",
    "Codigo de Barras": "1234567890123456789012345678901234567890"  # Exemplo de código de barras
    # Adicione mais informações do boleto conforme necessário
}

# Gerar o boleto
generate_boleto("boleto`%d`.pdf", boleto_data)
