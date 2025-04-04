import os  # Importa o módulo 'os' para interagir com o sistema operacional (lidar com arquivos, pastas, etc.)
import subprocess  # Importa o módulo 'subprocess' para executar comandos do sistema operacional

caminho_ghostscript = r"C:\Program Files\gs\gs10.05.0\bin\gswin64c.exe"  # Define o caminho para o executável do Ghostscript

def comprimir_pdf_ghostscript(caminho_entrada, caminho_saida):
    """
    Comprime um arquivo PDF usando o Ghostscript.
    """
    try:
        subprocess.run([  # Executa o Ghostscript como um subprocesso
            caminho_ghostscript,  # Caminho para o executável do Ghostscript
            "-sDEVICE=pdfwrite",  # Especifica que a saída é um PDF
            "-dCompatibilityLevel=1.4",  # Define a compatibilidade do PDF (1.4 é uma boa opção)
            "-dPDFSETTINGS=/ebook",  # Configurações de compressão para ebooks
            "-dNOPAUSE",  # Desativa a pausa entre páginas
            "-dQUIET",  # Desativa a saída detalhada
            "-dBATCH",  # Desativa o modo interativo
            f"-sOutputFile={caminho_saida}",  # Define o caminho do arquivo de saída
            caminho_entrada,  # Caminho do arquivo de entrada
        ], check=True)  # Verifica se o comando foi executado com sucesso
        print(f"PDF comprimido e salvo em {caminho_saida}")  # Imprime mensagem de sucesso
    except subprocess.CalledProcessError as e:  # Captura erros do Ghostscript
        print(f"Erro ao comprimir PDF: {e}")  # Imprime mensagem de erro

def comprimir_pdfs_pasta(pasta_entrada, pasta_saida, renomear=False):
    """
    Comprime todos os arquivos PDF em uma pasta e suas subpastas.

    Args:
        pasta_entrada (str): Caminho da pasta de entrada.
        pasta_saida (str): Caminho da pasta de saída.
        renomear (bool, opcional): Se True, renomeia e substitui os originais. Padrão é False.
    """
    for raiz, subpastas, arquivos in os.walk(pasta_entrada):  # Percorre a pasta e subpastas
        for nome_arquivo in arquivos:  # Itera sobre os arquivos na pasta atual
            caminho_entrada = os.path.join(raiz, nome_arquivo)  # Caminho completo do arquivo de entrada
            if nome_arquivo.lower().endswith('.pdf'):  # Verifica se é um arquivo PDF
                if renomear:
                    # Renomeia para "comprimido_<nome_original>"
                    nome_arquivo_comprimido = "comprimido_" + nome_arquivo
                    caminho_saida = os.path.join(raiz, nome_arquivo_comprimido)
                    comprimir_pdf_ghostscript(caminho_entrada, caminho_saida)

                    # Exclui o arquivo original
                    os.remove(caminho_entrada)

                    # Renomeia para o nome original
                    caminho_saida_final = os.path.join(raiz, nome_arquivo)
                    os.rename(caminho_saida, caminho_saida_final)
                else:
                    # Comprime e salva na pasta de saída
                    caminho_saida = os.path.join(pasta_saida, nome_arquivo)
                    comprimir_pdf_ghostscript(caminho_entrada, caminho_saida)

# Exemplo de uso
pasta_entrada = r"\\simp20019\dados\SIMPRESS\RH\Adm. Pessoal\IRON\Automação\2025"  # Pasta de entrada
pasta_saida = r"\\simp20019\dados\SIMPRESS\RH\Adm. Pessoal\IRON\Automação\2025"  # Pasta de saída

comprimir_pdfs_pasta(pasta_entrada, pasta_saida, renomear=True)  # Chama a função para comprimir os PDFs