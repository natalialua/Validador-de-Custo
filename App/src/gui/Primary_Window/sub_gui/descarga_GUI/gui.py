import os
import re
import sys
import threading
import tkinter as tk
import tkinter.messagebox as tk1
import tkinter.filedialog as filedialog
import webbrowser
import win10toast
from pathlib import Path
from tkinter import Canvas, Entry, Button, PhotoImage, ttk
from PIL import Image, ImageTk
import sqlite3
import win32api
import win32con
import win10toast
import shutil
from backend.Descarga import CalcDes


OUTPUT_PATH = Path(__file__).parent
ASSETS_PATH = OUTPUT_PATH / Path("./assets")

def relative_to_assets(path: str) -> Path:
    return ASSETS_PATH / Path(path)

# Função para obter dados do banco de dados
def get_data_from_db(query):
    db_path = r"G:"
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute(query)
        data = cursor.fetchall()
        conn.close()
        return [item for item in data]
    except sqlite3.Error as e:
        print(f"Erro ao acessar o banco de dados: {e}")
        return []

# Dropdowns
transp_options = get_data_from_db("SELECT DISTINCT TRANSPORTADORA FROM TABELAS_FRACIONADO")
transp_options = get_data_from_db("SELECT DISTINCT TRANSPORTADORA FROM TABELAS_LOTACAO")

def create_dropdown(parent, options, label_text, x, y):
    font = ("Fira Sans", 10, "bold")
    font_color = "#10264E"
    label = tk.Label(parent, text=label_text, bg="#FFFFFF", font=font, fg=font_color)
    label.place(x=x, y=y)
    dropdown = ttk.Combobox(parent, values=options, font=font)
    dropdown.place(x=x, y=y+20)
    return dropdown

# Definição das variáveis globais para os dropdowns
transp_dropdown = None
transp_options = []
tx_descarga_label = None



def update_transp_options(event):
    global transp_options, transp_dropdown
    selected_value = tp_exped_dropdown.get()

    if selected_value:
        transp_dropdown.config(state="normal")
    else:
        transp_dropdown.config(state="disabled")

    if selected_value in ["Z1", "E1"]:
        transp_options = get_data_from_db("""
            SELECT DISTINCT TRANSPORTADORA 
            FROM tabelas_lotacao 
            WHERE UPPER(TRIM(STATUS)) = 'ATIVA'
        """)
    elif selected_value == "Z2":
        transp_options = get_data_from_db("""
            SELECT DISTINCT TRANSPORTADORA 
            FROM tabelas_fracionado 
            WHERE UPPER(TRIM(STATUS)) = 'ATIVA'
        """)
    else:
        transp_options = []

    transp_dropdown["values"] = transp_options


##################### valor de TX Descarga   ##################### 
    
def create_label(parent, text, x, y):
        font = ("Fira Sans", 10, "bold")
        font_color = "#10264E"
        label = tk.Label(parent, text=text, bg="#FFFFFF", font=font, fg=font_color)
        label.place(x=x, y=y)
        return label

# def update_tx_descarga(event):
#     selected_transp = transp_dropdown.get()
#     selected_type = tp_exped_dropdown.get()  # Verifica o tipo selecionado (Z1 ou Z2)

#     # Redefinir o valor do rótulo para "-"
#     tx_descarga_label.config(text="Tx Descarga:\n-")

#     if selected_transp:
#         # Define a tabela com base no tipo selecionado
#         table_name = "tabelas_lotacao" if selected_type == "Z1" else "tabelas_fracionado"
        
#         # Buscar o valor da coluna "TX Descarga"
#         tx_descarga_result = get_data_from_db(f"""
#             SELECT "FRANQUIA" 
#             FROM {table_name} 
#             WHERE TRANSPORTADORA = '{selected_transp}'
#         """)


def update_tx_descarga(event):
    selected_transp = transp_dropdown.get()
    selected_type = tp_exped_dropdown.get()  # Verifica o tipo selecionado (Z1, Z2 ou E1)

    # Redefinir o valor do rótulo para "-"
    tx_descarga_label.config(text="Tx Descarga:\n-")

    if selected_transp:
        # Define a tabela com base no tipo selecionado
        if selected_type in ["Z1", "E1"]:
            table_name = "tabelas_lotacao"
        else:
            table_name = "tabelas_fracionado"
        
        # Buscar o valor da coluna "FRANQUIA"
        tx_descarga_result = get_data_from_db(f"""
            SELECT "FRANQUIA" 
            FROM {table_name} 
            WHERE TRANSPORTADORA = '{selected_transp}'
        """)


        if tx_descarga_result:
            tx_descarga_value = tx_descarga_result[0][0] if isinstance(tx_descarga_result[0], (list, tuple)) else tx_descarga_result[0]
            formatted_value = f"{tx_descarga_value:.2f}"

            tx_descarga_label.config(text=f"Franquia:\n{formatted_value}")

################################ || ########################################
            
def Descarga(pai):
    global tp_exped_dropdown, transp_dropdown, entrada_caminho, tx_descarga_label

    canvas = tk.Canvas(pai, bg="#FFFFFF", height=405, width=675, bd=0, highlightthickness=0, relief="ridge")
    canvas.place(x=230, y=72)

    tp_exped_dropdown = create_dropdown(canvas, ["Z1", "Z2","E1"], "Tp Expedição", 60, 50)
    transp_dropdown = create_dropdown(canvas, transp_options, "Transportadora", 260, 50)
    tx_descarga_label = create_label(canvas, "Franquia:", 480, 50)

    tp_exped_dropdown.bind("<<ComboboxSelected>>", update_transp_options)
    transp_dropdown.bind("<<ComboboxSelected>>", update_tx_descarga)   

    ################################ || ########################################
    
    # Anexo
    entrada_caminho = Entry(canvas, font=("Fira Sans", 10), bd=0, highlightthickness=0)  # Fonte menor e sem bordas
    canvas_width = 700  # Largura do canvas
    entry_width = 200   

    # Cálculo para centralizar horizontalmente
    x_position = (canvas_width - entry_width) // 2

    # Posição mais abaixo
    y_position = 250

    entrada_caminho.place(x=x_position, y=y_position, width=entry_width)

    ############################### Anexo Email ########################################
    path_image = Image.open(ASSETS_PATH / "Path1.png")
    path_image = path_image.resize((160, 50), Image.LANCZOS)  # Ajuste de tamanho para menor
    path_image_tk = ImageTk.PhotoImage(path_image)

    # Exibir Path.png no centro do canvas
    path_label = tk.Label(canvas, image=path_image_tk, bg="#FFFFFF", bd=0)
    path_label.place(relx=0.2, rely=0.80, anchor="center")  # Centralizado

    referencias_imagens["path_image"] = path_image_tk

    # Criar um rótulo para "Inserir Anexo" um pouco mais abaixo
    inserir_anexo_label = tk.Label(canvas, text="Anexo Email", bg="#DADADA", fg="#222222", font=("Dm Sans", 7))
    inserir_anexo_label.place(relx=0.2, rely=0.80, anchor="center")  # Ajustado para mais abaixo

    # file2 (Correção: Criar variável separada)
    file2_img = PhotoImage(file=ASSETS_PATH / "file2.png")
    
    path_picker_button2 = Button(
    canvas,
    image=file2_img,
    borderwidth=0,
    highlightthickness=0,
    command=lambda: selecionar_caminho("email", canvas),  # Passando canvas
    relief='flat',
    bg="#DADADA",
    activebackground="#DADADA"
)

    path_picker_button2.place(relx=0.28, rely=0.80, anchor="center")  
    path_picker_button2.lift()

    referencias_imagens["file2_img"] = file2_img  # Adiciona referência para não ser deletado


    ################################ Anexo SAP ########################################

    # Exibir Path.png no centro do canvas
    path_label1 = tk.Label(canvas, image=path_image_tk, bg="#FFFFFF", bd=0)
    path_label1.place(relx=0.5, rely=0.80, anchor="center")  # Mais para o lado direito

    referencias_imagens["path_image"] = path_image_tk

    # Criar um rótulo para "Inserir Anexo" um pouco mais abaixo
    inserir_anexo_label1 = tk.Label(canvas, text="Anexo SAP", bg="#DADADA", fg="#222222", font=("Dm Sans", 7))
    inserir_anexo_label1.place(relx=0.5, rely=0.80, anchor="center")  # Ajustado para mais abaixo

    # Carregar e posicionar o ícone de pasta (file1.png)
    path_picker_img = PhotoImage(file=ASSETS_PATH / "file1.png")

    path_picker_button = Button(
    canvas,
    image=path_picker_img,
    borderwidth=0,
    highlightthickness=0,
    command=lambda: selecionar_caminho("sap", canvas),  # Passando canvas
    relief='flat',
    bg="#DADADA",
    activebackground="#DADADA"
)

    path_picker_button.place(relx=0.58, rely=0.80, anchor="center")
    path_picker_button.lift()

    referencias_imagens["path_picker_img"] = path_picker_img
   
    global Descarga_button_image_1
    Descarga_button_image_1 = tk.PhotoImage(file=relative_to_assets("button_2.png"))

    tk.Button(
            image=Descarga_button_image_1,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: threading.Thread(target=Download_button_clicked, args=(pai,), daemon=True).start(),
            relief="flat",
            bg='#FFFFFF',
            activebackground='#FFFFFF'
        ).place(x=715.0, y=380.0)  
    
caminho_saida_email = ""
caminho_saida_sap = ""

PASTA_DESTINO = r"G:"


label_email = None
label_sap = None

def selecionar_caminho(tipo_anexo, canvas):
    global caminho_saida_email, caminho_saida_sap, label_email, label_sap

    if tipo_anexo == "email":
        filetypes = [("Excel files", "*.xlsx;*.xls")]
        title = "Selecione um arquivo Excel para anexar ao Email"
    elif tipo_anexo == "sap":
        filetypes = [("Text files", "*.txt")]
        title = "Selecione um arquivo TXT para anexar ao SAP"
    else:
        filetypes = [("Todos os arquivos suportados", "*.xlsx;*.xls;*.txt")]

    caminho_selecionado = filedialog.askopenfilename(filetypes=filetypes, title=title)

    if caminho_selecionado:
        print(f"Caminho selecionado ({tipo_anexo}): {caminho_selecionado}")  # Depuração
        
        if tipo_anexo == "email":
            caminho_saida_email = caminho_selecionado
            nome_arquivo_email = os.path.basename(caminho_saida_email)

            if label_email:
                label_email.config(text=nome_arquivo_email)
            else:
                label_email = tk.Label(canvas, text=nome_arquivo_email, bg="#FFFFFF", fg="#222222", font=("Dm Sans", 7))
                label_email.place(relx=0.2, rely=0.90, anchor="center")

        elif tipo_anexo == "sap":
            caminho_saida_sap = caminho_selecionado
            nome_arquivo_sap = os.path.basename(caminho_saida_sap)

            if label_sap:
                label_sap.config(text=nome_arquivo_sap)
            else:
                label_sap = tk.Label(canvas, text=nome_arquivo_sap, bg="#FFFFFF", fg="#222222", font=("Dm Sans", 7))
                label_sap.place(relx=0.5, rely=0.90, anchor="center")

def salvar_anexos():
    """ Copia os arquivos selecionados para a pasta de destino. """
    if caminho_saida_email:
        destino_email = Path(PASTA_DESTINO) / Path(caminho_saida_email).name
        shutil.copy(caminho_saida_email, destino_email)
        print(f"Arquivo do Email salvo em: {destino_email}")

    if caminho_saida_sap:
        destino_sap = Path(PASTA_DESTINO) / Path(caminho_saida_sap).name
        shutil.copy(caminho_saida_sap, destino_sap)
        print(f"Arquivo do SAP salvo em: {destino_sap}")

def Download_button_clicked(root):
    """ Função chamada ao clicar no botão de download. """
    global caminho_saida_email, caminho_saida_sap

    if not caminho_saida_email and not caminho_saida_sap:
        root.after(100, lambda: toast.show_toast("Erro", "Insira um anexo para cálculo", duration=3, icon_path=relative_to_assets("")))
        return

    try:
        caminho_destino = PASTA_DESTINO
     
        if caminho_saida_email:
            file_name_email = os.path.basename(caminho_saida_email)
            shutil.copy(caminho_saida_email, os.path.join(caminho_destino, file_name_email))
            print(f"Arquivo de Email salvo em: {caminho_destino}/{file_name_email}")
   
        if caminho_saida_sap:
            file_name_sap = os.path.basename(caminho_saida_sap)
            shutil.copy(caminho_saida_sap, os.path.join(caminho_destino, file_name_sap))
            print(f"Arquivo do SAP salvo em: {caminho_destino}/{file_name_sap}")
        
        win32api.MessageBox(0, "Aguarde enquanto o arquivo é gerado", "Carregando arquivo", win32con.MB_ICONINFORMATION)

        transportadora = transp_dropdown.get()

        CalcDes.main(transportadora)

        win32api.MessageBox(0, "O arquivo foi criado com sucesso!", "Arquivo gerado", win32con.MB_ICONINFORMATION)

    except Exception as e:
        print("Erro ao baixar anexo:", e)
        win32api.MessageBox(0, f"Ocorreu um erro: {e}", "Erro", win32con.MB_ICONERROR)
       

referencias_imagens = {}
