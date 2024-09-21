import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import imaplib
import email
import pdfplumber
import re
import openpyxl

def check_imap_enabled(emailAddress, password):
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(emailAddress, password)
        mail.logout()
        return True
    except imaplib.IMAP4.error as e:
        if "not enabled for IMAP use" in str(e):
            messagebox.showerror("Erro IMAP", "Sua conta não está habilitada para uso do IMAP. Por favor, ative o IMAP nas configurações do Gmail.")
        else:
            messagebox.showerror("Erro de Autenticação", "Erro de autenticação. Por favor, verifique suas credenciais.")
        return False

def extractPdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = ''
        for page in pdf.pages:
            text += page.extract_text()
        

    # Expressões regulares para encontrar os valores
    namePattern = r'Nome: ([A-Za-z ]+)'  # Exemplo: Nome: John Doe
    cpfPattern = r'CPF: (\d{3}\.\d{3}\.\d{3}-\d{2})'  # Exemplo: CPF: 123.456.789-00
    valorPattern = r'Valor: R\$ (\d+,\d{2})'  # Exemplo: Valor: R$ 100,00
    vencimentoPattern = r'Vencimento: (\d{2}/\d{2}/\d{4})'  # Exemplo: Vencimento: 10/06/2024
    codigoBarrasPattern = r'Código de Barras: (\d{40})'  # Exemplo: Código de Barras: 1234567890123456789012345678901234567890

    # Procurar por padrões nos dados extraídos
    nameMatch = re.search(namePattern, text)
    cpfMatch = re.search(cpfPattern, text)
    valorMatch = re.search(valorPattern, text)
    vencimentoMatch = re.search(vencimentoPattern, text)
    codigoBarrasMatch = re.search(codigoBarrasPattern, text)
    

    # Extrair os valores correspondentes, se encontrados
    name = nameMatch.group(1) if nameMatch else None
    cpf = cpfMatch.group(1) if cpfMatch else None
    valor = valorMatch.group(1) if valorMatch else None
    vencimento = vencimentoMatch.group(1) if vencimentoMatch else None
    codigo_barras = codigoBarrasMatch.group(1) if codigoBarrasMatch else None

    return name, cpf, valor, vencimento, codigo_barras

def ProccesEmail():
    emailAddress = emailEntry.get()
    password = passEntry.get()
    FolderAnexos = EntryAnexo.get()
    FolderPlan = EntryPlan.get()

    if not check_imap_enabled(emailAddress, password):
        return

    messagebox.showinfo("Login", "Entrando na conta de e-mail...")

    # Conectar à conta de e-mail
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(emailAddress, password)
    mail.select("inbox")
    messagebox.showinfo("Login", "Login efetuado com sucesso!")

    # Obter todos os e-mails não lidos
    status, messages = mail.search(None, "UNSEEN")
    if not messages[0]:
        messagebox.showinfo("Busca", "Nenhum e-mail não lido encontrado.")
    else:
        messagebox.showinfo("Busca", "E-mails não lidos encontrados.")
        for num in messages[0].split():
            status, data = mail.fetch(num, "(RFC822)")
            email_message = email.message_from_bytes(data[0][1])

            # Verificar se o e-mail possui anexos
            if email_message.get_content_maintype() == "multipart":
                for part in email_message.walk():
                    if part.get_content_maintype() == "multipart" or part.get("Content-Disposition") is None:
                        continue
                    filename = part.get_filename()
                    if filename:
                        # Processar o anexo
                        filepath = os.path.join(FolderAnexos, filename)
                        with open(filepath, "wb") as f:
                            f.write(part.get_payload(decode=True))
                        messagebox.showinfo("Download", f"Arquivo {filename} baixado com sucesso.")

                        # Lógica para extrair dados
                        name, cpf, valor, vencimento, codigo_barras = extractPdf(filepath)
                        
                        # Criando a planilha
                        add_to_spreadsheet(name, cpf, valor, vencimento, codigo_barras, filename, FolderPlan)
            else:
                messagebox.showinfo("Nenhum Anexo encontrado!", "Verificamos e nenhum anexo foi encontrado nos emails recebidos!")

    mail.logout()
    messagebox.showinfo("Concluído", "Processo de e-mail concluído.")

def add_to_spreadsheet(name, cpf, valor, vencimento, codigo_barras, filename, FolderPlan):
    # Abrir ou criar planilha
    wbPath = os.path.join(FolderPlan, "Anexos_Email.xlsx")
    if os.path.exists(wbPath):
        wb = openpyxl.load_workbook(wbPath)
        messagebox.showinfo("Planilha", "Planilha existente carregada.")
    else:
        wb = openpyxl.Workbook()
        messagebox.showinfo("Planilha", "Nova planilha criada.")
    ws = wb.active
    
    # Adicionar os dados à planilha
    ws.append([name, cpf, valor, vencimento, codigo_barras, filename])
    
    # Salvar a planilha
    wb.save(wbPath)
    messagebox.showinfo("Planilha", "Dados adicionados à planilha com sucesso.")

def directoryAnexos():
    directory = filedialog.askdirectory()
    if directory:
        EntryAnexo.delete(0, tk.END)
        EntryAnexo.insert(0, directory)
        EntryAnexo.config(state="readonly")

def directoryPlan():
    directory = filedialog.askdirectory()
    if directory:
        EntryPlan.delete(0, tk.END)
        EntryPlan.insert(0, directory)
        EntryPlan.config(state="readonly")

# Criar a janela principal
root = tk.Tk()
root.title("Bot")
root.geometry("600x250")

# Configurar o layout de grade para a janela principal
root.grid_rowconfigure(0, weight=1)
root.grid_rowconfigure(1, weight=1)
root.grid_rowconfigure(2, weight=1)
root.grid_rowconfigure(3, weight=1)
root.grid_rowconfigure(4, weight=0)
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=4)

#widgets

#Email
emailLabel = ttk.Label(root, text="E-mail", justify="left")
emailLabel.grid(column=0, row=0, padx=(0, 20))
emailEntry = ttk.Entry(root, text="Seu Email...",)
emailEntry.grid(column=1, row=0, sticky="ew", padx=(0, 20))

#senha
passLabel = ttk.Label(root, text="Senha",  justify="left")
passLabel.grid(column=0, row=1, padx=(0, 20))
passEntry = ttk.Entry(root, text="sua senha...", show="*",)
passEntry.grid(column=1, row=1, sticky="ew", padx=(0, 20))

#Selecionar pasta para anexo
btnAnexo = ttk.Button(root, text="Pasta dos Anexos", command=directoryAnexos)
btnAnexo.grid(column=0, row=2, padx=(0, 20))
EntryAnexo = ttk.Entry(root,)
EntryAnexo.grid(column=1, row=2, sticky="ew", padx=(0, 20))

#selecionar Pasta para Planilha
btnPlan = ttk.Button(root, text="Pasta das Planilhas", command=directoryPlan)
btnPlan.grid(column=0, row=3, padx=(0, 20))
EntryPlan = ttk.Entry(root,)
EntryPlan.grid(column=1, row=3, sticky="ew", padx=(0, 20))

#botão para Iniciar Processo!
frame = ttk.Frame(root, padding=20)
frame.grid(row=5, column=0, columnspan=2,)
btnFinal = ttk.Button(frame, width=40, text="Finalizar", command=ProccesEmail)
btnFinal.grid(row=0, column=0,sticky="nsew",)

# Iniciar o loop principal
root.mainloop()