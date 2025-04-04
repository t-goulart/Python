'''
Move arquivos para uma pasta superior de acordo com o nível informado
Sobre os níveis consideramos como nível 1 a pasta raiz e assim sucessivamente
Exemplo: 
Vamos usar um link como exemplo

D:\Empresa\Departamento Pessoal\Funcionarios\Prontuario\

A pasta do link sera considerada como nivel_superior 1, consequentemente cada pasta sera o proximo nivel.
Vamos supor que dentro de Prontuario temos a pasta '100000' (D:\Empresa\Departamento Pessoal\Funcionarios\Prontuario\100000), dentro desta pasta temos 7 subpastas, cada uma com diversos arquivos e documentos.
Se quiser transferir tudo para a pasta '100000', será necessário informar o nivel da pasta que precisa receber todos os arquivos dentro dela

caminho_raiz = r"D:\Empresa\Departamento Pessoal\Funcionarios\Prontuario\" # informe a pasta raiz
nivel_superior = 2 # informe o nivel | Prontuario seria nivel 1, então 100000 será nivel 2 | Logo tudo que está nas subpastas de 100000 será movido para essa pasta
mover_arquivos_para_nivel_superior(caminho_raiz, nivel_superior)
'''

import os  # Importa o módulo 'os' para interagir com o sistema operacional (lidar com arquivos, pastas, etc.)
import shutil  # Importa o módulo 'shutil' para operações de movimentação de arquivos

def mover_arquivos_para_nivel_superior(caminho_raiz, nivel_superior):
    """
    Move arquivos de subpastas para um nível superior especificado.

    Args:
        caminho_raiz (str): Caminho da pasta raiz onde a busca será realizada.
        nivel_superior (int): Nível superior para o qual os arquivos serão movidos.
    """
    if not os.path.exists(caminho_raiz):  # Verifica se o caminho raiz existe
        print(f"Erro: Caminho não encontrado: {caminho_raiz}")  # Imprime mensagem de erro se o caminho não existir
        return  # Sai da função se o caminho não existir

    raiz_nivel = caminho_raiz.count(os.sep)  # Calcula o nível da pasta raiz contando o número de separadores de caminho
    # Por exemplo, em "C:\pasta1\pasta2", o nível seria 2

    for raiz, diretorios, arquivos in os.walk(caminho_raiz):
        # Percorre recursivamente a pasta raiz e suas subpastas
        # 'raiz' é o caminho da pasta atual, 'diretorios' são as subpastas, 'arquivos' são os arquivos na pasta

        nivel_atual = raiz.count(os.sep) - raiz_nivel
        # Calcula o nível da pasta atual em relação à pasta raiz
        # Se 'raiz' for "C:\pasta1\pasta2\subpasta", 'nivel_atual' seria 1

        if nivel_atual > nivel_superior:
            # Verifica se o nível atual é maior que o nível superior desejado

            caminho_destino = os.path.dirname(raiz)
            # Obtém o caminho da pasta pai da pasta atual
            # Se 'raiz' for "C:\pasta1\pasta2\subpasta", 'caminho_destino' seria "C:\pasta1\pasta2"

            while nivel_atual > nivel_superior:  # Adicionado para mover até o nível correto.
                # Loop para subir até o nível superior desejado
                caminho_destino = os.path.dirname(caminho_destino)
                # Sobe um nível no caminho de destino
                nivel_atual -= 1
                # Decrementa o nível atual

            for arquivo in arquivos:
                # Loop para iterar sobre os arquivos na pasta atual

                caminho_origem = os.path.join(raiz, arquivo)
                # Constrói o caminho completo do arquivo de origem

                caminho_destino_arquivo = os.path.join(caminho_destino, arquivo)
                # Constrói o caminho completo do arquivo de destino

                try:
                    shutil.move(caminho_origem, caminho_destino_arquivo)
                    # Move o arquivo da origem para o destino
                    print(f"Arquivo movido: {caminho_origem} -> {caminho_destino_arquivo}")
                except Exception as e:
                    # Captura exceções se a movimentação falhar
                    print(f"Erro ao mover arquivo {caminho_origem}: {e}")

# Exemplo de uso
caminho_raiz = r""
nivel_superior = 2
mover_arquivos_para_nivel_superior(caminho_raiz, nivel_superior)
