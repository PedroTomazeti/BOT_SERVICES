import customtkinter as ctk
from tkinter import messagebox
import queue
import threading
import os
import sqlite3
from PIL import Image
from processos.pesquisa_xml import main_xml
from web.web_app import iniciar_driver

# Configuração inicial do CustomTkinter
ctk.set_appearance_mode("Dark")  # Modo: "System", "Dark" ou "Light"
ctk.set_default_color_theme("dark-blue")  # Tema: "blue", "dark-blue", "green"

log_queue = queue.Queue()
running_event = threading.Event()

class QueueOutput:
    def __init__(self, log_queue):
        self.log_queue = log_queue

    def write(self, message):
        if message.strip():  # Adiciona apenas mensagens não vazias
            self.log_queue.put(message)

    def flush(self):
        pass

# Atualiza a interface com as mensagens do log
def update_text_widget(text_widget, log_queue):
    try:
        while True:
            message = log_queue.get_nowait()
            text_widget.configure(state='normal')
            text_widget.insert(ctk.END, message + '\n')
            text_widget.configure(state='disabled')
            text_widget.see(ctk.END)
    except queue.Empty:
        pass
    text_widget.after(500, update_text_widget, text_widget, log_queue)

# Mude de acordo com o número e localização da sua unidade
def escolha_unidade(unidade):
    '''
    Mude de acordo com o número e localização da sua unidade cadastrado no sistema.
    '''
    match unidade:
        case "0102-SLZ":
            return 1
        
        case "0103-PRP":
            return 2
        
        case "0104-SJC":
            return 3

# Função para iniciar a análise de serviços
def iniciar_analise():
    unidade_selecionada = unidade_box.get()
    mes_selecionado = mes_box.get()
    ano_selecionado = ano_box.get()

    if not unidade_selecionada or not mes_selecionado or not ano_selecionado:
        messagebox.showerror("Erro", "Por favor, selecione a unidade, o mês e o ano corretamente.")
        return

    unidades_map = {
        "0102-SLZ": "03 - Notas Filial I São Luís",
        "0103-PRP": "04 - Notas Filial II Parauapebas",
        "0104-SJC": "05 - Notas Filial III São José dos Campos",
    }

    filial_map = {
        "0102-SLZ": "Filial I",
        "0103-PRP": "Filial II",
        "0104-SJC": "Filial III",        
    }

    pasta_base = r'caminho_base'
    pasta_xml = r'caminho_xml'
    pasta_unidade = unidades_map[unidade_selecionada]
    filial = filial_map[unidade_selecionada]
    caminho_pasta = os.path.join(pasta_base, pasta_unidade, f"Notas {ano_selecionado}", mes_selecionado, "02 - Serviços")
    caminho_xml = os.path.join(pasta_xml, pasta_unidade, f"Notas {ano_selecionado}", mes_selecionado, "02 - Serviços")

    if os.path.exists(caminho_xml):
        messagebox.showinfo("Informação", f"Análise iniciada para: {caminho_xml}")
        janela_secundaria.destroy()
        threading.Thread(target=main_xml, args=(caminho_xml, caminho_pasta, log_queue, filial), daemon=True).start()
    else:
        messagebox.showerror("Erro", f"A pasta '{caminho_xml}' não foi encontrada.")
        janela_secundaria.destroy()

# Função para inserir no sistema
def inserir_no_sistema():
    unidade_selecionada = unidade_box.get()
    mes_selecionado = mes_box.get()
    ano_selecionado = ano_box.get()

    # Mude de acordo com o número e localização da sua unidade cadastrado no sistema.
    unidades_map = {
        "0102-SLZ": "são_luís",
        "0103-PRP": "parauapebas",
        "0104-SJC": "são_josé",
    }

    num_unidade = escolha_unidade(unidade_selecionada)

    db_unidade = unidades_map[unidade_selecionada]
    db = f"/caminho/para/dist/notas_{db_unidade}.db"

    if not mes_selecionado or not ano_selecionado:
        messagebox.showerror("Erro", "Por favor, selecione o mês e o ano.")
        return

    mes = mes_selecionado.split(" - ")[1].strip().lower()
    nome_tabela = f"{mes}_{ano_selecionado}"  # Nome da tabela baseado no mês e ano

    if os.path.isfile(db):
        # Conectar ao banco de dados
        conn = sqlite3.connect(db)
        cursor = conn.cursor()

        # Verificar se a tabela existe
        cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{nome_tabela}';")
        tabela_existe = cursor.fetchone()

        if tabela_existe:
            messagebox.showinfo("Informação", f"A tabela '{nome_tabela}' foi encontrada no banco de dados.\n"
                                              f"Inserindo no sistema para: {mes} de {ano_selecionado}.")
            threading.Thread(target=iniciar_driver, args=(num_unidade, db_unidade, nome_tabela, log_queue, mes_selecionado), daemon=True).start()
        else:
            messagebox.showerror("Erro", f"A tabela '{nome_tabela}' não foi encontrada no banco de dados '{db}'.\n"
                                         f"Verifique se a análise foi feita corretamente.")

        # Fechar conexão
        cursor.close()
        conn.close()
        
        janela_secundaria.destroy()
    
    else:
        messagebox.showerror("Erro", f"O banco de dados '{db}' não foi encontrado.\n"
                                     f"Verifique se o arquivo está na pasta correta.")
        janela_secundaria.destroy()
    
# Criar a janela principal
janela = ctk.CTk()
janela.geometry("700x500")
janela.title("BOT-SERVICO")
janela.resizable(False, False)

# Carrega a imagem original
bg_image = Image.open("path/para/o/background.jpg")

# Redimensiona a imagem conforme necessário
bg_image = bg_image.resize((730, 530))

# Cria um objeto CTkImag
bg_ctk_image = ctk.CTkImage(light_image=bg_image, size=(730, 530))

# Passa o CTkImage para o CTkLabel
background_label = ctk.CTkLabel(janela, image=bg_ctk_image, text="")
background_label.place(x=0, y=0, relwidth=1, relheight=1)

# Centralizar a janela principal
largura_tela = janela.winfo_screenwidth()
altura_tela = janela.winfo_screenheight()
largura_janela = 700
altura_janela = 500
pos_x = (largura_tela // 2) - (largura_janela // 2)
pos_y = (altura_tela // 2) - (altura_janela // 2)
janela.geometry(f"{largura_janela}x{altura_janela}+{pos_x}+{pos_y}")

# Label e selection box para unidade
unidade_label = ctk.CTkLabel(janela, text="Selecione a Unidade:")
unidade_label.pack(pady=15)

unidade_box = ctk.CTkComboBox(janela, values=["0102-SLZ", "0103-PRP", "0104-SJC"])
unidade_box.pack(pady=10)

# Botão para abrir a janela secundária
def abrir_janela_secundaria(opcao):
    global janela_secundaria
    janela_secundaria = ctk.CTkToplevel(janela)
    janela_secundaria.title("Opções")
    janela_secundaria.geometry("400x350")

    largura_tela_sec = janela_secundaria.winfo_screenwidth()
    altura_tela_sec = janela_secundaria.winfo_screenheight()
    largura_janela_sec = 400
    altura_janela_sec = 350
    pos_x_sec = (largura_tela_sec // 2) - (largura_janela_sec // 2)
    pos_y_sec = (altura_tela_sec // 2) - (altura_janela_sec // 2)
    janela_secundaria.geometry(f"{largura_janela_sec}x{altura_janela_sec}+{pos_x_sec}+{pos_y_sec}")

    janela_secundaria.grab_set()

    # Elementos comuns (Mês e Ano)
    global mes_box, ano_box
    mes_label = ctk.CTkLabel(janela_secundaria, text="Selecione o Mês:")
    mes_label.pack(pady=5)
    mes_box = ctk.CTkComboBox(janela_secundaria, values=[
        "01 - Janeiro", "02 - Fevereiro", "03 - Março", "04 - Abril", "05 - Maio", "06 - Junho",
        "07 - Julho", "08 - Agosto", "09 - Setembro", "10 - Outubro", "11 - Novembro", "12 - Dezembro"
    ])
    mes_box.pack(pady=5)

    ano_label = ctk.CTkLabel(janela_secundaria, text="Selecione o Ano:")
    ano_label.pack(pady=5)
    ano_box = ctk.CTkComboBox(janela_secundaria, values=["2024", "2025"])
    ano_box.pack(pady=5)

    # Elementos específicos para cada opção
    if opcao == "Inserindo no Sistema":
        botao_confirmar = ctk.CTkButton(janela_secundaria, text="Inserir", command=inserir_no_sistema)
    else:
        botao_confirmar = ctk.CTkButton(janela_secundaria, text="Iniciar Análise", command=iniciar_analise)

    botao_confirmar.pack(pady=15)

# Botões principais
botao_analise = ctk.CTkButton(janela, text="Análise de Serviços", command=lambda: abrir_janela_secundaria("Análise"))
botao_analise.pack(pady=10)

botao_inserir = ctk.CTkButton(janela, text="Inserindo no Sistema", command=lambda: abrir_janela_secundaria("Inserindo no Sistema"))
botao_inserir.pack(pady=10)

# Caixa de log para exibir os prints
log_box = ctk.CTkTextbox(janela, width=500, height=300, state="disabled")
log_box.pack(pady=20)

# Inicia a atualização dos logs na interface
update_text_widget(log_box, log_queue)

janela.mainloop()