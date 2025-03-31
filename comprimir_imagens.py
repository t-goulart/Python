from PIL import Image  # Importa a biblioteca PIL (Pillow) para manipulação de imagens
import os  # Importa o módulo os para interagir com o sistema operacional (arquivos e pastas)

def comprimir_imagem(caminho_entrada, caminho_saida, qualidade=85):
    """
    Comprime uma imagem.

    Args:
        caminho_entrada: Caminho da imagem de entrada.
        caminho_saida: Caminho da imagem de saída.
        qualidade: Qualidade da imagem (0-100).
    """
    try:
        with Image.open(caminho_entrada) as img:  # Abre a imagem no caminho de entrada
            # Comprimir a imagem
            if img.mode in ("RGBA", "P"):  # Verifica se a imagem tem canal alfa (transparência) ou é uma imagem paletizada
                img = img.convert("RGB")  # Converte a imagem para o modo RGB (remove o canal alfa)
            img.save(caminho_saida, quality=qualidade, optimize=True)  # Salva a imagem comprimida no caminho de saída, com a qualidade especificada e otimização
        print(f"Imagem comprimida e salva em {caminho_saida}")  # Imprime uma mensagem de sucesso
    except FileNotFoundError:
        print(f"Erro: Arquivo não encontrado em {caminho_entrada}")  # Imprime uma mensagem de erro se o arquivo de entrada não for encontrado
    except Exception as e:
        print(f"Erro ao comprimir imagem: {e}")  # Imprime uma mensagem de erro para outras exceções

def comprimir_imagens_pasta(pasta_entrada, pasta_saida, qualidade=85):
    """
    Comprime todas as imagens em uma pasta.

    Args:
        pasta_entrada: Caminho da pasta de entrada.
        pasta_saida: Caminho da pasta de saída.
        qualidade: Qualidade da imagem (0-100).
    """
    # Cria a pasta de saída se ela não existir
    if not os.path.exists(pasta_saida):  # Verifica se a pasta de saída existe
        os.makedirs(pasta_saida)  # Cria a pasta de saída se ela não existir

    for nome_arquivo in os.listdir(pasta_entrada):  # Itera sobre todos os arquivos na pasta de entrada
        caminho_entrada = os.path.join(pasta_entrada, nome_arquivo)  # Cria o caminho completo do arquivo de entrada
        if os.path.isfile(caminho_entrada) and nome_arquivo.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp')):  # Verifica se o arquivo é uma imagem (com base na extensão)
            caminho_saida = os.path.join(pasta_saida, nome_arquivo)  # Cria o caminho completo do arquivo de saída
            comprimir_imagem(caminho_entrada, caminho_saida, qualidade)  # Chama a função comprimir_imagem para comprimir a imagem

# Exemplo de uso
caminho_origem = r"D:\03. Cursos\Python\Teste\Entrada"  # Define o caminho raiz a ser varrido
caminho_destino = r"D:\03. Cursos\Python\Teste\Saida"  # Define o caminho raiz a ser varrido
comprimir_imagem(caminho_origem, caminho_destino, qualidade=30)  # Chama a função varrer_pastas() para iniciar a compressão'
