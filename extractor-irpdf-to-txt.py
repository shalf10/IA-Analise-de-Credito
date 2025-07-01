import pdfplumber

def extrair_texto_pdf(caminho_pdf, salvar_em_txt=False):
    texto_final = ""

    with pdfplumber.open(caminho_pdf) as pdf:
        for i, pagina in enumerate(pdf.pages):
            texto = pagina.extract_text()
            if texto:
                # Remove linhas inúteis (ex: página X de Y, headers etc.)
                linhas = texto.split('\n')
                linhas_filtradas = [linha for linha in linhas if not linha.lower().startswith("página") and linha.strip() != ""]
                texto_limpo = '\n'.join(linhas_filtradas)
                
                texto_final += f"\n--- Página {i+1} ---\n"
                texto_final += texto_limpo + "\n"

    if salvar_em_txt:
        with open("saida_pdf_limpo.txt", "w", encoding="utf-8") as f:
            f.write(texto_final)
        print("Texto extraído e salvo em 'relatório de analise.txt'.")

    return texto_final

# Use assim:
caminho_arquivo = "documento.pdf"
texto_extraido = extrair_texto_pdf(caminho_arquivo, salvar_em_txt=True)
