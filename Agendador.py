# Esse codigo faz a leitura do XML que contém as informações de uma pasta específica do Agendador de Tarefas do Windows
# Depois criar um arquivo Excel em uma pasta específica e organiza os dados em formato de tabela com o cabeçalho: ["Nome", "Status", "Tipo Disparador", "Data Início", "Hora Início", "Repetição Disparador", "Comandos"] 
# Será necessário anterar a pasta onde as suas automações estão sendo executadas e também o local onde o arquivo será criado como o seu nome
# Macros no Agendador: Linha 16 do código
# Local que vai salvar o arquivo Excel: Linha 117 do código | Exemplo: \\\\Rede\\Pasta\\Agendamento.xlsx

import win32com.client # Importa a biblioteca pywin32 para interagir com componentes COM do Windows.
import openpyxl # Importa a biblioteca openpyxl para trabalhar com arquivos Excel (ler e escrever).
import xml.etree.ElementTree as ET # Importa a biblioteca ElementTree para analisar arquivos XML.
import os # Importa o módulo os para funcionalidades relacionadas ao sistema operacional (não está sendo usado diretamente neste código).

def capturar_agendamentos_macros_xml():
    try:
        scheduler = win32com.client.Dispatch("Schedule.Service") # Cria um objeto para interagir com o serviço do Agendador de Tarefas do Windows.
        scheduler.Connect() # Conecta-se ao serviço do Agendador de Tarefas local.
        root_folder = scheduler.GetFolder("\\Macros") # Obtém a pasta raiz chamada "Macros" dentro do Agendador de Tarefas.
        tasks = root_folder.GetTasks(0) # Obtém uma coleção de todas as tarefas dentro da pasta "Macros". O argumento 0 indica que não há filtros específicos.

        dados_agendamentos = [] # Inicializa uma lista vazia para armazenar os dados de cada agendamento.
        for task in tasks: # Itera sobre cada tarefa encontrada na pasta "Macros".
            nome_tarefa = task.Name # Obtém o nome da tarefa.
            status_tarefa = task.State # Obtém o estado atual da tarefa (ex: Pronto, Em Execução, Desabilitado).
            disparadores_info = [] # Inicializa uma lista vazia para armazenar informações sobre os disparadores (triggers) da tarefa.
            comandos = [] # Inicializa uma lista vazia para armazenar os comandos que a tarefa executa.

            try:
                xml_content = task.Xml # Obtém o conteúdo XML que define a configuração da tarefa.
                if xml_content: # Verifica se o conteúdo XML não está vazio.
                    root = ET.fromstring(xml_content) # Analisa a string XML e cria um elemento raiz para facilitar a navegação.
                    ns = {'task': 'http://schemas.microsoft.com/windows/2004/02/mit/task'} # Define um namespace para simplificar a busca de elementos XML específicos do Agendador de Tarefas.
                    triggers_element = root.find('.//task:Triggers', namespaces=ns) # Busca o elemento "Triggers" dentro do XML, utilizando o namespace definido.
                    if triggers_element is not None: # Verifica se o elemento "Triggers" foi encontrado.
                        for trigger in triggers_element: # Itera sobre cada disparador (trigger) dentro do elemento "Triggers".
                            trigger_type = trigger.tag.split('}')[-1] # Obtém o tipo do disparador (ex: TimeTrigger, CalendarTrigger) removendo o namespace.
                            start_boundary_element = trigger.find('./task:StartBoundary', namespaces=ns) # Busca o elemento "StartBoundary" que define a data e hora de início do disparador.
                            start_time_str = start_boundary_element.text if start_boundary_element is not None else "" # Obtém o texto do elemento "StartBoundary" se existir, caso contrário, define como string vazia.
                            start_date = "" # Inicializa a variável para armazenar a data de início.
                            start_hour = "" # Inicializa a variável para armazenar a hora de início.
                            if "T" in start_time_str: # Verifica se a string de data e hora contém o separador "T".
                                start_date, start_hour = start_time_str.split("T") # Separa a data e a hora da string.

                            repetition_interval_element = trigger.find('./task:Repetition/task:Interval', namespaces=ns) # Busca o elemento "Interval" dentro de "Repetition", que define o intervalo de repetição.
                            interval = repetition_interval_element.text if repetition_interval_element is not None else "Uma vez" # Obtém o texto do intervalo de repetição ou define como "Uma vez" se não houver.
                            interval_description = interval # Inicializa a descrição da repetição com o valor do intervalo.

                            if trigger_type == "CalendarTrigger": # Verifica se o tipo do disparador é um agendamento baseado em calendário.
                                schedule_by_month = trigger.find('./task:ScheduleByMonth', namespaces=ns) # Busca o elemento "ScheduleByMonth" para agendamentos mensais.
                                if schedule_by_month is not None: # Verifica se há um agendamento mensal.
                                    months_element = schedule_by_month.find('./task:Months', namespaces=ns) # Busca o elemento "Months" que especifica os meses de execução.
                                    days_element = schedule_by_month.find('./task:DaysOfMonth', namespaces=ns) # Busca o elemento "DaysOfMonth" que especifica os dias do mês de execução.

                                    months_list = [month.tag.split('}')[-1] for month in months_element] if months_element is not None else [] # Cria uma lista dos meses de execução, removendo o namespace.
                                    months_str = f" nos meses de {', '.join(months_list)}" if months_list else "" # Formata a string dos meses.

                                    days_of_month_elements = days_element.findall('./task:Day', namespaces=ns) if days_element is not None else [] # Busca todos os elementos "Day" dentro de "DaysOfMonth".
                                    days_of_month = [day.text for day in days_of_month_elements] # Cria uma lista dos dias do mês de execução.
                                    days_str = f" nos dias {', '.join(days_of_month)}" if days_of_month else "" # Formata a string dos dias do mês.

                                    if months_str or days_str: # Verifica se há informações de meses ou dias do mês.
                                        interval_description = f"Mensalmente{months_str}{days_str}" # Formata a descrição completa para agendamentos mensais com meses e dias.
                                    else:
                                        interval_description = "Mensalmente" # Caso haja ScheduleByMonth, mas sem detalhes de dias ou meses.
                                elif interval_description == "Uma vez" and repetition_interval_element is not None:
                                    interval_description = interval # Se for um agendamento único dentro de um CalendarTrigger.

                                # Capturar dias da semana (SE NÃO FOR MENSAL ESPECÍFICO)
                                schedule_by_week = trigger.find('./task:ScheduleByWeek', namespaces=ns) # Busca o elemento "ScheduleByWeek" para agendamentos semanais.
                                days_of_week = [] # Inicializa a lista para armazenar os dias da semana.
                                if schedule_by_week is not None: # Verifica se há um agendamento semanal.
                                    weeks_interval = schedule_by_week.find('./task:WeeksInterval', namespaces=ns) # Busca o elemento "WeeksInterval" que define a frequência semanal.
                                    weeks_interval_str = f" a cada {weeks_interval.text} semana(s)" if weeks_interval is not None and weeks_interval.text != "1" else " semanalmente" # Formata a string do intervalo semanal.
                                    days_element_week = schedule_by_week.find('./task:DaysOfWeek', namespaces=ns) # Busca o elemento "DaysOfWeek" que especifica os dias da semana.
                                    if days_element_week is not None: # Verifica se há dias da semana especificados.
                                        for day in days_element_week: # Itera sobre cada dia da semana.
                                            days_of_week.append(day.tag.split('}')[-1]) # Adiciona o dia da semana à lista, removendo o namespace.
                                        if days_of_week: # Verifica se há dias da semana na lista.
                                            interval_description = f"Semanalmente ({', '.join(days_of_week)}){weeks_interval_str}" # Formata a descrição completa para agendamentos semanais.
                                elif interval_description == "Uma vez" and repetition_interval_element is not None and trigger_type != "CalendarTrigger":
                                    interval_description = interval # Para outros tipos de triggers com repetição única.

                            disparadores_info.append({ # Adiciona as informações do disparador à lista.
                                "Tipo": trigger_type,
                                "Data_Inicio": start_date,
                                "Hora_Inicio": start_hour,
                                "Repeticao": interval_description
                            })

                    actions_element = root.find('.//task:Actions', namespaces=ns) # Busca o elemento "Actions" que contém as ações que a tarefa executa.
                    if actions_element is not None: # Verifica se o elemento "Actions" foi encontrado.
                        for action in actions_element: # Itera sobre cada ação dentro do elemento "Actions".
                            if action.tag.split('}')[-1] == "Exec": # Verifica se o tipo da ação é "Exec" (executar um comando).
                                command_element = action.find('./task:Command', namespaces=ns) # Busca o elemento "Command" que contém o comando a ser executado.
                                if command_element is not None: # Verifica se o elemento "Command" foi encontrado.
                                    comandos.append(command_element.text) # Adiciona o comando à lista de comandos.

                else:
                    disparadores_info.append({"Tipo": "Erro", "Data_Inicio": "Erro", "Hora_Inicio": "Erro", "Repeticao": "Erro ao ler XML"}) # Adiciona informações de erro caso não consiga ler o XML.
                    comandos.append("Erro ao ler XML") # Adiciona uma mensagem de erro à lista de comandos.

            except Exception as e_xml:
                disparadores_info.append({"Tipo": "Erro", "Data_Inicio": "Erro", "Hora_Inicio": "Erro", "Repeticao": f"Erro ao analisar XML: {e_xml}"}) # Adiciona informações de erro caso ocorra um erro ao analisar o XML.
                comandos.append(f"Erro ao analisar XML: {e_xml}") # Adiciona uma mensagem de erro à lista de comandos.

            dados_agendamentos.append({ # Adiciona os dados do agendamento (nome, status, disparadores e comandos) à lista principal.
                "Nome": nome_tarefa,
                "Status": status_tarefa,
                "Disparadores": disparadores_info,
                "Comandos": "; ".join(comandos) # Junta os comandos em uma única string separada por "; ".
            })

        return dados_agendamentos # Retorna a lista com os dados de todos os agendamentos encontrados.

    except Exception as e_scheduler:
        print(f"Erro ao acessar o Agendador de Tarefas: {e_scheduler}") # Imprime uma mensagem de erro se houver falha ao acessar o Agendador de Tarefas.
        return [] # Retorna uma lista vazia em caso de erro ao acessar o Agendador de Tarefas.

def salvar_agendamentos_excel(dados, nome_arquivo="Caminho e nome do arquivo Excel que os dados devem ser salvos"):
    workbook = openpyxl.Workbook() # Cria uma nova planilha Excel na memória.
    sheet = workbook.active # Obtém a folha de trabalho ativa (a primeira por padrão).
    cabecalhos = ["Nome", "Status", "Tipo Disparador", "Data Início", "Hora Início", "Repetição Disparador", "Comandos"] # Define os cabeçalhos das colunas.
    sheet.append(cabecalhos) # Adiciona os cabeçalhos à primeira linha da planilha.
    for agendamento in dados: # Itera sobre cada agendamento capturado.
        if agendamento.get("Disparadores"): # Verifica se o agendamento possui informações de disparadores.
            for disparador in agendamento["Disparadores"]: # Itera sobre cada disparador do agendamento.
                sheet.append([ # Adiciona uma nova linha na planilha com os detalhes do agendamento e do disparador.
                    agendamento.get("Nome", ""),
                    agendamento.get("Status", ""),
                    disparador.get("Tipo", ""),
                    disparador.get("Data_Inicio", ""),
                    disparador.get("Hora_Inicio", ""),
                    disparador.get("Repeticao", ""),
                    agendamento.get("Comandos", "")
                ])
        else: # Caso o agendamento não tenha disparadores.
            sheet.append([ # Adiciona uma nova linha na planilha com informações de que não há disparadores.
                agendamento.get("Nome", ""),
                agendamento.get("Status", ""),
                "Nenhum",
                "Nenhum",
                "Nenhum",
                "Nenhum",
                agendamento.get("Comandos", "")
            ])
    try:
        workbook.save(nome_arquivo) # Salva a planilha Excel no arquivo especificado.
        print(f"Dados salvos com sucesso em '{nome_arquivo}'") # Imprime uma mensagem de sucesso.
    except Exception as e_save:
        print(f"Erro ao salvar o arquivo Excel: {e_save}") # Imprime uma mensagem de erro se houver falha ao salvar o arquivo.

if __name__ == "__main__":
    dados = capturar_agendamentos_macros_xml() # Chama a função para capturar os dados dos agendamentos.
    if dados: # Verifica se a lista de dados não está vazia (ou seja, se agendamentos foram encontrados).
        salvar_agendamentos_excel(dados) # Chama a função para salvar os dados em um arquivo Excel.
