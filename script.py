import fitz  # PyMuPDF
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

def replace_marker_in_pdf(input_pdf_path, output_pdf_path, placeholder, replacement_text):
    # Abre o documento PDF
    document = fitz.open(input_pdf_path)

    # Itera por todas as páginas do PDF
    for page_num in range(len(document)):
        page = document.load_page(page_num)  # Carrega a página
        text_instances = page.search_for(placeholder)  # Procura o marcador

        # Verifica se encontrou o marcador
        if not text_instances:
            print(f"Marcador '{placeholder}' não encontrado na página {page_num + 1}")

        # Itera por todas as instâncias encontradas do marcador
        for inst in text_instances:
            # Extrai o retângulo onde o texto do marcador está
            rect = fitz.Rect(inst)

            # Desenha um retângulo com a cor #FFED00 para cobrir o texto existente
            page.draw_rect(rect, color=(1, 0.93, 0), fill=(1, 0.93, 0))

            # Calcula a posição do novo texto um pouco mais abaixo
            new_text_position = fitz.Point(rect.x0, rect.y1 - 2.5)

            # Adiciona o texto de substituição na nova posição
            page.insert_text(new_text_position, replacement_text, fontsize=12, color=(0, 0, 0))

    # Salva o PDF com as substituições feitas
    document.save(output_pdf_path)
    print(f"PDF salvo como {output_pdf_path}")

def send_email_with_pdf(smtp_server, smtp_port, sender_email, sender_password, recipient_email, subject, body, attachment_path):
    # Cria a mensagem de email
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject

    # Adiciona o corpo do email
    msg.attach(MIMEText(body, 'plain'))

    # Adiciona o anexo
    with open(attachment_path, "rb") as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header(
            'Content-Disposition',
            f'attachment; filename= {os.path.basename(attachment_path)}',
        )
        msg.attach(part)

    # Conecta ao servidor SMTP e envia o email
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        text = msg.as_string()
        server.sendmail(sender_email, recipient_email, text)
        server.quit()
        print(f"Email enviado para {recipient_email}")
    except Exception as e:
        print(f"Falha ao enviar email para {recipient_email}: {str(e)}")

# Exemplo de uso:
input_pdf_path = 'template.pdf'
placeholder = 'Marcador'
people = [
    {
        "name": "Yago Gomes",
        "email": "yago.fgomes@gmail.com"
    },
    {
        "name": "Thais Nunes",
        "email": "thais.dnunes@hotmail.com"
    }
]

sender_email = "thais.testes@outlook.pt"
sender_password = "Rayka101010*"
smtp_server = "smtp-mail.outlook.com"
smtp_port = 587

for p in people:
    name = p["name"]
    email = p["email"]
    output_pdf_path = f"{name}.pdf"

    replace_marker_in_pdf(input_pdf_path, output_pdf_path, placeholder, name)
    
    # Envia o email com o PDF gerado
    subject = "Seu Certificado"
    body = f"Olá {name},\n\nSegue em anexo o seu certificado.\n\nAtenciosamente,\nEquipe"
    send_email_with_pdf(smtp_server, smtp_port, sender_email, sender_password, email, subject, body, output_pdf_path)
