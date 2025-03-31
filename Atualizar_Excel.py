# Esse código tem como objetivo abrir um arquivo Excel, conectar com a fonte¹, atualizar², salvar e fechar
# 1. Verifique se existe uma conexão previamente cadastrada no gerenciador de credenciais
# 2. Atualiza qualquer conexão, seja com sharepoint, Excel, Access e etc.

import os  # Importa o módulo 'os' para interagir com o sistema operacional (caminhos de arquivos, pastas, etc.)
import win32com.client  # Importa o módulo 'win32com.client' para interagir com aplicativos do Windows (Excel, etc.)
import time  # Importa o módulo 'time' para usar funções relacionadas a tempo (pausar a execução, etc.)

def atualizar_arquivos_excel(caminhos_arquivos, tempo_espera_minutos=1):
    """
    Abre arquivos Excel, atualiza, salva e fecha cada um deles.

    Args:
        caminhos_arquivos (list): Lista com os caminhos completos dos arquivos Excel.
        tempo_espera_minutos (int): Tempo de espera em minutos após a atualização.
    """
    try:
        excel = win32com.client.Dispatch("Excel.Application")  # Cria uma instância do aplicativo Excel
        excel.Visible = True  # Torna o Excel visível
        excel.DisplayAlerts = False  # Desativa a exibição de alertas
        excel.ScreenUpdating = False # Desativa a exibição de alertas

        for caminho_arquivo in caminhos_arquivos:  # Percorre a lista de caminhos de arquivos Excel
            if os.path.exists(caminho_arquivo):  # Verifica se o arquivo existe
                print(f"Abrindo e atualizando: {os.path.basename(caminho_arquivo)}")  # Imprime mensagem informando qual arquivo está sendo aberto e atualizado
                workbook = excel.Workbooks.Open(caminho_arquivo)  # Abre o arquivo Excel
                workbook.RefreshAll()  # Atualiza todas as conexões de dados e tabelas dinâmicas no arquivo Excel

                tempo_espera_segundos = tempo_espera_minutos * 60  # Converte o tempo de espera de minutos para segundos
                print(f"Aguardando {tempo_espera_segundos} segundos para conclusão das atualizações...")  # Imprime mensagem informando o tempo de espera
                time.sleep(tempo_espera_segundos)  # Pausa a execução do código pelo tempo de espera

                workbook.Save()  # Salva as alterações no arquivo Excel
                workbook.Close()  # Fecha o arquivo Excel
                print(f"{os.path.basename(caminho_arquivo)} atualizado com sucesso.")  # Imprime mensagem informando que o arquivo foi atualizado com sucesso
            else:
                print(f"Arquivo não encontrado: {caminho_arquivo}")  # Imprime mensagem se o arquivo não for encontrado

        excel.DisplayAlerts = True  # Reativa a exibição de alertas
        excel.ScreenUpdating = True # Reativa a exibição de alertas
        excel.Quit()  # Sai do aplicativo Excel
        print("Processo concluído.")  # Imprime mensagem informando que o processo foi concluído

    except Exception as e:  # Captura exceções que possam ocorrer durante a execução do código
        print(f"Ocorreu um erro: {e}")  # Imprime mensagem de erro

# Lista com os caminhos dos arquivos Excel
arquivos_excel = [
    r"D:\03. Cursos\Python\Teste\Base de Vendas.xlsx"  # Caminho do arquivo Excel "Cargos.xlsx"
    # Adicione outros caminhos de arquivos aqui, se necessário
]

# Chama a função 'atualizar_arquivos_excel' com a lista de arquivos e o tempo de espera de 1 minuto
atualizar_arquivos_excel(arquivos_excel, tempo_espera_minutos=1)  