import fitz  # PyMuPDF

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

for p in people:
    name = p["name"]
    email = p["email"]
    output_pdf_path = f"{name}.pdf"

    replace_marker_in_pdf(input_pdf_path, output_pdf_path, placeholder, name)
