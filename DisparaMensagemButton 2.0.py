# Importar bibliotecas necessárias
import tkinter as tk
from tkinter import ttk
import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
from ttkbootstrap.style import Style    
from urllib.parse import quote
from time import sleep
from datetime import datetime
import pandas as pd
import webbrowser
import threading
import pyautogui
import json
import os
import logging
import openpyxl
import tkinter.messagebox





#================================================================================================================================



logging.basicConfig(filename='mesagens_enviadas.log', level=logging.INFO, format='%(asctime)s - %(message)s')



#================================================================================================================================



# Função para carregar o histórico de linha da última mensagem enviada
def load_last_line():
    if os.path.exists("last_line.json"):
        with open("last_line.json", "r") as f:
            data = json.load(f)
            return data.get("Ultima_linha_enviada", 0) # Retorna 0 se não encontrar a chave "ultima_linha"
    return 0

# Função para salvar a última mensagem enviada no histórico
def save_last_line(last_line):
    with open("last_line.json", "w") as f:
        json.dump({"Ultima_linha_enviada": last_line}, f)



# carregar as configurações do arquivo JSON
def load_settings():
    if os.path.exists("settings.json"):
        with open("settings.json", "r") as f:
            return json.load(f)
    else:
        return {"theme": "journal"} 
    


# salvar as configurações no arquivo JSON
def save_settings(settings):
    with open("settings.json", "w") as f:
        json.dump(settings, f)


# Função para carregar os números enviados de um arquivo JSON
def carregar_numeros_enviados():
    if os.path.exists("numeros_enviados.json"):
        with open("numeros_enviados.json", "r") as f:
         return set(json.load(f)) # Carrega como lista e converte para conjunto
    return set()

def salvar_numeros_enviados(numeros):
    with open("numeros_enviados.json", "w") as f:
        json.dump(list(numeros), f) # Converte conjunto para lista e salva


# ================================================================================================================================


# Criar uma classe para a GUI
class CourseOfferGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Disparador de Mensagem Cursos")
        self.root.geometry("600x600")


        # Carregar configurações
        self.settings = load_settings()

        # Aplicar o tema carregado
        self.style = Style(theme=self.settings["theme"])

        # Carregar a última linha enviada
        self.last_line = load_last_line()
        
        # Carregar números enviados do arquivo
        self.numeros_enviados = carregar_numeros_enviados();

        # Criação de um frame principal para organizar os widgets
        main_frame = ttk.Frame(self.root)
        main_frame.pack(expand=True, fill="both", padx=0, pady=0)

    
        # ================================================================================================================================


    
        # Lista os temas disponíveis
        available_themes = ttk.Style().theme_names()
        print("Temas disponíveis:", available_themes)

        # ADICIONA UMA OPÇÃO DE MENU PARA ESCOLHER O TEMA
        self.menu_theme = ttk.Menubutton(root, text="Escolher Tema")
        self.theme = tk.Menu(self.menu_theme, tearoff=1)

        themes = ["superhero", "flatly", "darkly", "journal", "cyborg", "lumen", "minty", "pulse", "sandstone", "solar", "united", "yeti", "cerulean", "cosmo", "litera", "morph", "simplex", "vapor"]

        for theme in themes:
            self.theme.add_command(label=theme.capitalize(), command=lambda t=theme: self.change_theme(t))

        self.menu_theme.configure(menu=self.theme)
        self.menu_theme.pack(pady=1)


    
        # ================================================================================================================================
        

        self.escolher_aba_button = ttkb.Button(self.root, text="Selecionar Aba", command=lambda: self.escolher_aba("PRE_MATRICULA_AMTECH.xlsx"))
        self.escolher_aba_button.pack(padx=0, pady=1)

        
        # Adicionar funcionalidade para listar e selecionar abas
    def escolher_aba(self, workbook_path):
        try:
            # Remover widgets antigos, se existirem
            if hasattr(self, 'dynamic_frame'):
                self.dynamic_frame.destroy()  # Remove o frame antigo

            # Criar um novo frame dinâmico
            self.dynamic_frame = ttk.Frame(self.root)
            self.dynamic_frame.pack(pady=10)

            # Carregar a planilha
            workbook = openpyxl.load_workbook(workbook_path, data_only=True)
            self.abas_disponiveis = workbook.sheetnames

            # Exibir as abas disponíveis
            self.aba = tk.StringVar(value=self.abas_disponiveis[0])
            self.aba_label = ttk.Label(self.dynamic_frame, text="Escolha umas das abas:", font=('Arial', 10), justify='center')
            self.aba_label.pack(pady=5, padx=0)

                
            # Exibir as abas disponíveis
            self.aba_menu = ttk.Combobox(self.dynamic_frame, textvariable=self.aba, values=self.abas_disponiveis)
            self.aba_menu.pack(pady=10)
            self.aba_menu.bind("<<ComboboxSelected>>", lambda e: print(f"Aba selecionada: {self.aba.get()}"))


            # Botão para confirmar seleção
            self.confirmar_aba_botao = ttk.Button(self.dynamic_frame, text="Confirmar Aba", command=lambda: self.processar_aba(workbook))
            self.confirmar_aba_botao.pack(pady=10)

            
        except Exception as e:
            print(f"Erro ao carregar a planilha: {e}")
            tkinter.messagebox.showerror("Erro", f"Erro ao carregar a planilha: {e}")
        
        
        # Processar a aba selecionada
    def processar_aba(self, workbook):
        try:

            # Selecionar a aba
            aba_selecionada = self.aba.get()
            self.aba_selecionada = workbook[aba_selecionada]
            print(f"Processando aba: {aba_selecionada}")


            # Atualizar o DataFrame com os dados da aba selecionada
            self.alunos_df = pd.DataFrame(self.aba_selecionada.values)
            self.alunos_df.columns = self.alunos_df.iloc[0]  # Define a primeira linha como cabeçalhos
            self.alunos_df = self.alunos_df[1:]  # Excluir a primeira linha
            print("Dados carregados com sucesso.")


        except Exception as e:
            print(f"Erro ao processar a aba: {e}")
            tkinter.messagebox.showerror("Erro", "Erro ao processar a aba: {e}")
        


        # ================================================================================================================================


        #  título da aplicação
        self.title_label = ttk.Label(self.root, text='Disparador de Mensagem Cursos', font=('Arial', 20))
        self.title_label.pack(pady=10)


        # ===============================================================================================================================



        # Recebe os dados de entrada.
        self.menu_course = ttk.Menubutton(root, text="Curso")
        self.course_selected = tk.StringVar()
        self.course = tk.Menu(self.menu_course, tearoff=0)



        # Adicione os cursos aqui
        self.course.add_command(label="ANIMAÇÃO DE PERSONAGENS PARA GAMES", command=lambda: self.Chosing_Course("ANIMAÇÃO DE PERSONAGENS PARA GAMES"))
        self.course.add_command(label="CRIACAO E MANIPULACAO DE IMAGENS-GIMP", command=lambda: self.Chosing_Course("CRIACAO E MANIPULACAO DE IMAGENS-GIMP"))
        self.course.add_command(label="CURRICULO,LINKEDIN E IMAGEM PESSOAL", command=lambda: self.Chosing_Course("CURRICULO,LINKEDIN E IMAGEM PESSOAL"))
        self.course.add_command(label="ESTRUTURA DE DADOS PYTHON", command=lambda: self.Chosing_Course("ESTRUTURA DE DADOS PYTHON"))
        self.course.add_command(label="EXCELÊNCIA EM ATENDIMENTO AO CLIENTE E-COMMERCE", command=lambda: self.Chosing_Course("EXCELÊNCIA EM ATENDIMENTO AO CLIENTE E-COMMERCE"))
        self.course.add_command(label="HTML E CSS PARA INICIANTES", command=lambda: self.Chosing_Course("HTML E CSS PARA INICIANTES"))
        self.course.add_command(label="HTML E CSSS - CRIAÇÃO DE SITES", command=lambda: self.Chosing_Course("HTML E CSSS - CRIAÇÃO DE SITES"))
        self.course.add_command(label="IMPRESSORA 3D – BÁSICO", command=lambda: self.Chosing_Course("IMPRESSORA 3D – BÁSICO"))
        self.course.add_command(label="INTRODUÇÃO A INOVAÇÃO E DESIGN", command=lambda: self.Chosing_Course("INTRODUÇÃO A INOVAÇÃO E DESIGN"))
        self.course.add_command(label="INTRODUÇÃO A INOVAÇÃO E DESIGN", command=lambda: self.Chosing_Course("INTRODUÇÃO A INOVAÇÃO E DESIGN"))
        self.course.add_command(label="LÓGICA DE PROGRAMAÇÃO", command=lambda: self.Chosing_Course("LÓGICA DE PROGRAMAÇÃO"))
        self.course.add_command(label="MARKETING DIGITAL", command=lambda: self.Chosing_Course("MARKETING DIGITAL"))
        self.course.add_command(label="OFICINA - FUNÇÕES BÁSICAS DO TELEFONE CELULAR", command=lambda: self.Chosing_Course("OFICINA - FUNÇÕES BÁSICAS DO TELEFONE CELULAR"))
        self.course.add_command(label="PROGRAMAÇÃO PARA ROBÓTICA", command=lambda: self.Chosing_Course("PROGRAMAÇÃO PARA ROBÓTICA"))
        self.course.add_command(label="PROGRAMAÇÃO EM FLUTTER PARA INICIANTES", command=lambda: self.Chosing_Course("PROGRAMAÇÃO EM FLUTTER PARA INICIANTES"))
        self.course.add_command(label="PHOTOSHOP - EDITOR GRÁFICO", command=lambda: self.Chosing_Course("PHOTOSHOP - EDITOR GRÁFICO"))



        self.menu_course.configure(menu=self.course)
        self.menu_course.pack(pady=5)



        # ================================================================================================================================



        # criar o menu de seleção de parceiros
        self.menu_button = ttk.Menubutton(root, text="Instituição Parceira")
        self.partner_selected = tk.StringVar()
        self.parceiro = tk.Menu(self.menu_button, tearoff=0)

        # Adiciona os parceiros aqui
        self.parceiro.add_command(label="SENAC", command=lambda: self.Chosing_Partner("SENAC"))
        self.parceiro.add_command(label="SENAI", command=lambda: self.Chosing_Partner("SENAI"))

        self.menu_button.configure(menu=self.parceiro)
        self.menu_button.pack(pady=5)



        # ================================================================================================================================



        # Função para criar o menu de seleção de grupos
        self.menu_group = ttk.Menubutton(root, text="Deseja enviar mensagem por grupos?")
        self.group_selected = tk.StringVar()
        self.group = tk.Menu(self.menu_group, tearoff=0)
        self.group.add_command(label="SIM", command=lambda: self.Chosing_Group("SIM"))
        self.group.add_command(label="NÃO", command=lambda: self.Chosing_Group("NÃO"))
        self.menu_group.configure(menu=self.group)
        self.menu_group.pack(pady=5)



        # ================================================================================================================================



        self.schedule_label = ttk.Label(root, text="Horário:")
        self.schedule_label.pack()
        self.schedule_entry = ttk.Entry(self.root, width=30, bootstyle='info')
        self.schedule_entry.pack()

        self.minage_label = ttk.Label(root, text="Idade mínima:")
        self.minage_label.pack()
        self.minage_entry = ttk.Entry(root, width=30, bootstyle='info')
        self.minage_entry.pack()

        self.duration_label = ttk.Label(root, text="Duração:")
        self.duration_label.pack()
        self.duration_entry = ttk.Entry(root, width=30, bootstyle='info')
        self.duration_entry.pack()

        self.minrange_label = ttk.Label(root, text="De qual linha devo começar:")
        self.minrange_label.pack()
        self.minrange_entry = ttk.Entry(root, width=30, bootstyle='info')
        self.minrange_entry.insert(0, str(self.last_line)) # Inserir o valor da última linha enviada
        self.minrange_entry.pack()

        self.maxrange_label = ttk.Label(root, text="Até qual linha devo enviar:")
        self.maxrange_label.pack()
        self.maxrange_entry = ttk.Entry(root, width=30, bootstyle='info')
        self.maxrange_entry.pack()

        # Criar um botão para enviar mensagens
        self.send_button = ttkb.Button(self.root, text="Enviar Mensagens", command=self.start_sending, bootstyle='success')
        self.send_button.pack(pady=10)


        # Criar um botão para cancelar o código
        self.cancel_button = ttkb.Button(self.root, text="Interromper código", command=self.interromper_codigo, bootstyle='danger')
        self.cancel_button.pack(pady=10)


        self.running = False


        
        # Adicionar o botão de reset na inicialização da GUI
        self.reset_button = ttkb.Button(self.root, text="Resetar Programa", command=self.reset_program, bootstyle='danger')
        self.reset_button.pack(pady=10)
    

    ################################
    def reset_program(self):
        # Limpa variáveis relacionadas ao envio
        self.alunos_df = None
        self.aba_selecionada = None
        self.numeros_enviados = set()
        self.course_selected.set("")
        self.partner_selected.set("")
        self.group_selected.set("")


        # Atualiza a interface para o estado inicial
        self.menu_course.config(text="Curso")
        self.menu_button.config(text="Instituição Parceira")
        self.menu_group.config(text="Deseja enviar mensagem por grupos?")
        self.schedule_entry.delete(0, tk.END)
        self.minage_entry.delete(0, tk.END)
        self.duration_entry.delete(0, tk.END)
        self.minrange_entry.delete(0, tk.END)
        self.maxrange_entry.delete(0, tk.END)
        self.minrange_entry.insert(0, str(self.last_line)) # Reinicia com a última linha salva

         # Remover widgets antigos, se existirem
        if hasattr(self, 'dynamic_frame'):
            self.dynamic_frame.destroy()  # Remove o frame antigo

        # Recriar o frame principal para organizar os widgets
        self.dynamic_frame = ttk.Frame(self.root)
        self.dynamic_frame.pack(pady=10)

        print("Programa resetado!")

         # CRÉDITOS
        self.credits_label = ttkb.Label(self.root, text='developed by: Eduardo Zanin', font=('Arial', 7, 'bold'))
        self.credits_label.pack(pady=5)

        self.credits_label = ttkb.Label(self.root, text='developed by: Lucas Ferrari', font=('Arial', 7, 'bold'))
        self.credits_label.pack(pady=5)


    # ================================================================================================================================


    # Funções de suporte.
    def interromper_codigo(self):
        print("Envio encerrado")
        self.running = False

    def start_sending(self):
        if self.running:
            print("O envio de mensagens já está em andamento.")
            return
        self.running = True
        print('Iniciado o envio de mensagens!')
        t = threading.Thread(target=self.send_messages)
        t.start()

    def Chosing_Course(self, curso: str):
        self.menu_course.config(text=curso)
        self.course_selected.set(curso)

    def Chosing_Partner(self, parceiro: str):
        self.menu_button.config(text=parceiro)
        self.partner_selected.set(parceiro)

    def Chosing_Group(self, grupo: str):
        self.menu_group.config(text=grupo)
        self.group_selected.set(grupo)

    def change_theme(self, theme):
        # Atualiza o tema na interface
        self.style.theme_use(theme)
        # Salva o tema escolhido nas configurações
        self.settings["theme"] = theme
        save_settings(self.settings)
    


    # ================================================================================================================================



    def encontra_categoria(self, curso_de_envio: str) -> list:
        # DEFINE AS CATEGORIAS
        Dev = ["ESTRUTURA DE DADOS PYTHON", "HTML E CSSS - CRIAÇÃO DE SITES, INTRODUÇÃO A LÓGICA DE PROGRAMAÇÃO", "INTRODUÇÃO A LÓGICA DE PROGRAMAÇÃO COM ALGORÍTIMOS", "INTRODUÇÃO A UTILIZAÇÃO DE IAS E CHATBOTS DE FORMA PRODUTIVA", "ANIMAÇÃO DE PERSONAGENS PARA GAMES", "MONTAGEM E MANUTENÇÃO DE COMPUTADORES", "IMPRESSORA 3D - BÁSICO", "INFORMÁTICA BÁSICA"]

        Marketing = ["MARKETING DIGITAL - WHATSAPP BUSINESS", "MARKETING DIGITAL", "MARKETPLACES - ESTRATÉGIAS PARA VENDAS E OPORTUNIDADES ONLINE", "EXCELÊNCIA EM ATENDIMENTO AO CLIENTE E-COMMERCE", "EMPREENDEDORISMO E ESTRATÉGIA DE MERCADO PARA NOVOS NEGÓCIOS", "EMPREENDEDORISMO E ESTRATÉGIA DE MERCADO PARA NOVOS NEGÓCIOS",  "ACELERA IDEAÇÃO - PROGRAMA DE IDEAÇÃO DE STARTUPS", "COMUNICAÇÃO E IMAGEM NO ATENDIMENTO AO CLIENTE", "INTELIGÊNCIA ARTIFICIAL (IA) PARA GESTÃO DE NEGÓCIOS"]

        Design = ["PHOTOSHOP - EDITOR GRÁFICO", "INTRODUÇÃO A INOVAÇÃO E DESIGN", "CRIAÇÃO E MANIPULAÇÃO DE IMAGENS - GIMP", "ACELERA IDEAÇÃO - PROGRAMA DE IDEAÇÃO DE STARTUPS"]

        Basicos = ["OFICINA - FUNÇÕES BÁSICAS DO TELEFONE CELULAR", "INFORMÁTICA BÁSICA"]

        lista_categoria = [Dev, Marketing, Design, Basicos]

        for Categorias in lista_categoria:
            if curso_de_envio in Categorias:
                return Categorias




    # ================================================================================================================================
    def send_messages(self):
      try:
        # Salva as entradas obtidas na interface.
        curso_de_envio = self.course_selected.get()
        parceiro = self.partner_selected.get()
        horario_do_curso = self.schedule_entry.get()
        data_de_duracao = self.duration_entry.get()
        linhamin = int(self.minrange_entry.get().strip())
        linhamax = int(self.maxrange_entry.get().strip())
        
        idademin = int(self.minage_entry.get().strip())
        por_grupo = self.group_selected.get()
        

        # ================================================================================================================================



        # Verifique se o DataFrame está carregado corretamente
        if self.alunos_df is None or self.alunos_df.empty:
            print("Erro: Nenhum dado disponível no DataFrame.")
            tkinter.messagebox.showerror("Erro", "Nenhum dado disponível para envio.")
            return

        # Verifique se os limites estão corretos
        self.runing = True
        total_alunos = linhamax - linhamin
        print(f"Enviando mensagens para {total_alunos} alunos...")

        # Ler o arquivo Excel
        alunos = self.alunos_df

        # Conjunto para armazenar números de telefone já processados
        numeros_enviados = self.numeros_enviados

        ultima_linha_enviada = None



        # ================================================================================================================================



        for x in range(linhamin, linhamax):
            if not self.running:
                print('Código interrompido na linha: {0}'.format(x))
                break


            try:
                linha_correta = x + 2


                cursos = alunos.loc[x, "Dentre as opções qual curso gostaria de fazer?"]
                lista_cursos = cursos.split(sep=', ')


                if por_grupo == "SIM":
                    categoria = self.encontra_categoria(curso_de_envio)
                    for curso in lista_cursos:
                        if curso in categoria:
                            nome = alunos.loc[x, 'Nome Completo']
                            telefone = int(alunos.loc[x, "Whatsapp com DDD (somente números - sem espaço)"])


                            if telefone in self.numeros_enviados:
                                print(f"Número {telefone} já recebeu mensagem. Pulando...")
                                continue


                            mensagem = (
                                f"Olá *{nome}.* Nós somos da AMTECH - Agência Maringá de Tecnologia e Inovação. "
                                f"entramos em contato porque você demonstrou interesse em cursos de tecnologia preenchendo um formulário.📋\n\n "
                                f" Nós iremos iniciar em parceria com o *{parceiro}*, o curso:\n\n 🌟*{curso_de_envio}*.🌟 \n\n "
                                f"Todos podem participar desde que sejam maior de *{idademin}* anos e tenham a escolaridade mínima 5º ano do Ensino Fundamental.🎓\n\n"
                                f"🎯 Duração do curso: *{data_de_duracao}*\n\n 🕒 Horário: *{horario_do_curso}* \n\n"
                                f"⚠️ Atenção: As vagas são limitadas! Responda o mais rápido possível! 🏃‍♂️💨 📢*\n\n"
                                f"*📍Local: Acesso 1 | Piso Superior Terminal Urbano - Av. Tamandaré, 600 - Zona 01, Maringá🗺️ -*\n\n"
                                f"*🏫 MODALIDADE: curso é PRESENCIAL E 100% GRATUITO! 🎉* \n\n Qualquer dúvida, estamos à disposição! Esperamos você! 😉"
                            )
                            print(f"Enviando mensagem para: {nome}, Telefone: {telefone}")

                            link_mensagem_whatsapp = f'https://web.whatsapp.com/send/?phone={telefone}&text={quote(mensagem)}'
                            webbrowser.open(link_mensagem_whatsapp)
                            sleep(2)
                            sleep(6)
                            pyautogui.press('enter')
                            sleep(6)
                            pyautogui.hotkey('ctrl', 'w')


                            self.numeros_enviados.add(telefone)  # Adiciona o número ao conjunto
                            salvar_numeros_enviados(self.numeros_enviados)


                            # Log completo para nome, telefone, curso e linha (corrigindo a linha)
                            logging.info(f'Messagem enviada para: {nome}, Telefone: {telefone}, Curso: {curso_de_envio}, Linha: {linha_correta}')
                            logging.info(f'Última linha enviada: {linha_correta}')
                            save_last_line(linha_correta)


                            if ultima_linha_enviada is not None:
                                print(f'Ultima linha enviada: {ultima_linha_enviada}')
                                logging.info(f'Ultima linha enviada: {ultima_linha_enviada}')



                else:
                    for curso in lista_cursos:
                        if curso.upper() == curso_de_envio.upper():
                            nome = alunos.loc[x, 'Nome Completo']
                            telefone = int(alunos.loc[x, 'Whatsapp com DDD (somente números - sem espaço)'])


                            if telefone in numeros_enviados:
                                print(f"Número {telefone} já recebeu mensagem. Pulando...")
                                continue


                            mensagem = (
                                f"Olá *{nome}.* Nós somos da AMTECH - Agência Maringá de Tecnologia e Inovação. "
                                f"entramos em contato porque você demonstrou interesse em cursos de tecnologia preenchendo um formulário.📋\n\n "
                                f" Nós iremos iniciar em parceria com o *{parceiro}*, o curso:\n\n 🌟*{curso_de_envio}*.🌟 \n\n "
                                f"Todos podem participar desde que sejam maior de *{idademin}* anos e tenham a escolaridade mínima 5º ano do Ensino Fundamental.🎓\n\n"
                                f"🎯 Duração do curso: *{data_de_duracao}*\n\n 🕒 Horário: *{horario_do_curso}* \n\n"
                                f"⚠️ Atenção: As vagas são limitadas! Responda o mais rápido possível! 🏃‍♂️💨 📢*\n\n"
                                f"*📍Local: Acesso 1 | Piso Superior Terminal Urbano - Av. Tamandaré, 600 - Zona 01, Maringá🗺️ -*\n\n"
                                f"*🏫 MODALIDADE: curso é PRESENCIAL E 100% GRATUITO! 🎉* \n\n Qualquer dúvida, estamos à disposição! Esperamos você! 😉"
                            )
                            print(f"Enviando mensagem para: {nome}, Telefone: {telefone}")


                            link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
                            webbrowser.open(link_mensagem_whatsapp)
                            sleep(2)
                            sleep(6)
                            pyautogui.press('enter')
                            sleep(6)
                            pyautogui.hotkey('ctrl', 'w')


                            self.numeros_enviados.add(telefone)
                            salvar_numeros_enviados(self.numeros_enviados)


                            logging.info(f'Messagem enviada para: {nome}, Telefone: {telefone}, Curso: {curso_de_envio}')
                            logging.info(f'Última linha enviada: {ultima_linha_enviada}')


                            if ultima_linha_enviada is not None:
                                print(f'Ultima linha enviada: {ultima_linha_enviada}')
                                logging.info(f'Ultima linha enviada: {ultima_linha_enviada}')
                                save_last_line(ultima_linha_enviada)
            except Exception as e:
                print(f"Erro ao processar linha {x}: {e}")
                continue

        print("Todas as linhas foram lidas!")
        self.running = False

        ttk.Label(self.root, text="Envio de mensagens concluído!", foreground="green").pack(pady=10)

      except Exception as e:
        print(f"Erro no envio de mensagens: {e}")
        tkinter.messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

# Criar a GUI
root = tk.Tk()
gui = CourseOfferGUI(root)
root.mainloop()













