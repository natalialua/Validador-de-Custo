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
from backend.Paletização import calc
import win32api
import win32con

# Configuração do toast notifier
toast = win10toast.ToastNotifier()

# Configuração dos caminhos
OUTPUT_PATH = Path(__file__).parent
ASSETS_PATH = OUTPUT_PATH / Path("./assets")

def relative_to_assets(path: str) -> Path:
    return ASSETS_PATH / Path(path)

################################ || ########################################

# SQLITE 
def get_data_from_db(query):
    db_path = r""
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute(query)
        data = cursor.fetchall()
        conn.close()
        return [item[0] for item in data]
    except sqlite3.Error as e:
        print(f"Erro ao acessar o banco de dados: {e}")
        return []
    
################################ || ########################################
    
# Dropdowns
transp_options = get_data_from_db("SELECT DISTINCT TRANSPORTADORA FROM TABELAS_FRACIONADO")
# origem_options = get_data_from_db("SELECT DISTINCT REGIÃO FROM TABELAS_FRACIONADO")

transp_options = get_data_from_db("SELECT DISTINCT TRANSPORTADORA FROM TABELAS_LOTACAO")
# origem_options = get_data_from_db("SELECT DISTINCT REGIÃO FROM TABELAS_LOTACAO")

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
tx_paletizacao_label = None

def update_transp_options(event):
    global transp_options, transp_dropdown
    selected_value = tp_exped_dropdown.get()

    if selected_value:
        transp_dropdown.config(state="normal")
    else:
        transp_dropdown.config(state="disabled")

    if selected_value == "Z1" or selected_value == "E1":
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


##################### valor de TX PALETIZAÇÃO   ##################### 
    
def create_label(parent, text, x, y):
        font = ("Fira Sans", 10, "bold")
        font_color = "#10264E"
        label = tk.Label(parent, text=text, bg="#FFFFFF", font=font, fg=font_color)
        label.place(x=x, y=y)
        return label

def update_tx_paletizacao(event):
    selected_transp = transp_dropdown.get()
    selected_type = tp_exped_dropdown.get()  # Verifica o tipo selecionado (Z1, Z2 ou E1)

    # Redefinir o valor do rótulo para "-"
    tx_paletizacao_label.config(text="Tx Paletização:\n-")

    if selected_transp:
        # Define a tabela com base no tipo selecionado
        if selected_type in ["Z1", "E1"]:
            table_name = "tabelas_lotacao"
        elif selected_type == "Z2":
            table_name = "tabelas_fracionado"
        else:
            table_name = None

        if table_name:
            # Buscar o valor da coluna "TX PALETIZAÇÃO"
            tx_paletizacao_result = get_data_from_db(f"""
                SELECT "TX PALETIZAÇÃO" 
                FROM {table_name} 
                WHERE TRANSPORTADORA = '{selected_transp}'
            """)

        if tx_paletizacao_result:
            # Supondo que o resultado seja uma lista com uma única linha
            tx_paletizacao_value = tx_paletizacao_result[0][0] if isinstance(tx_paletizacao_result[0], (list, tuple)) else tx_paletizacao_result[0]
            
            # Formatando o valor para duas casas decimais
            formatted_value = f"{tx_paletizacao_value:.2f}"

            # Exibir o valor formatado
            tx_paletizacao_label.config(text=f"Tx Paletização:\n{formatted_value}")

################################ || ########################################
            
def Paletizacao(pai):
    global tp_exped_dropdown, transp_dropdown, entrada_caminho, tx_paletizacao_label

    canvas = tk.Canvas(pai, bg="#FFFFFF", height=405, width=675, bd=0, highlightthickness=0, relief="ridge")
    canvas.place(x=230, y=72)

    tp_exped_dropdown = create_dropdown(canvas, ["Z1", "Z2", "E1"], "Tp Expedição", 80, 50)
    transp_dropdown = create_dropdown(canvas, transp_options, "Transportadora", 280, 50)
    tx_paletizacao_label = create_label(canvas, "Tx Paletização", 480, 50)

    tp_exped_dropdown.bind("<<ComboboxSelected>>", update_transp_options)
    transp_dropdown.bind("<<ComboboxSelected>>", update_tx_paletizacao)           

    
################################ || ########################################
    
    #Anexo

    entrada_caminho = Entry(canvas, font=("Fira Sans", 10), bd=0, highlightthickness=0)  # Fonte menor e sem bordas
    canvas_width = 700  # Largura do canvas
    entry_width = 200   

    # Cálculo para centralizar horizontalmente
    x_position = (canvas_width - entry_width) // 2

    # Posição mais abaixo
    y_position = 250

    entrada_caminho.place(x=x_position, y=y_position, width=entry_width)

################################ || ########################################

    # Carregar e redimensionar a imagem de fundo (Path.png)
    path_image = Image.open(ASSETS_PATH / "Path.png")
    path_image = path_image.resize((300, 60), Image.LANCZOS)  # Ajuste de tamanho
    path_image_tk = ImageTk.PhotoImage(path_image)
    
    # Exibir Path.png no centro do canvas
    path_label = tk.Label(canvas, image=path_image_tk, bg="#FFFFFF", bd=0)
    path_label.place(relx=0.5, rely=0.5, anchor="center")  # Centralizado
    
    referencias_imagens["path_image"] = path_image_tk

    # Criar um rótulo para "Inserir Anexo"
    inserir_anexo_label = tk.Label(canvas, text="Inserir Anexo", bg="#DADADA", fg="#222222", font=("Dm Sans", 8))
    inserir_anexo_label.place(relx=0.5, rely=0.5, anchor="center")  
    
################################ || ########################################

    # Carregar e posicionar o ícone de pasta (folder.png)
    path_picker_img = PhotoImage(file=ASSETS_PATH / "file.png")
    path_picker_button = Button(
        canvas,
        image=path_picker_img,
        borderwidth=0,
        highlightthickness=0,
        command=selecionar_caminho,
        relief='flat',
        bg="#DADADA",
        activebackground="#DADADA"
    )
    
    def ajustar_posicoes():
        path_width = path_image_tk.width()
        path_height = path_image_tk.height()

    # Define a posição do folder.png no canto direito de Path.png, ajustando mais para a esquerda
        path_picker_button.place(
        relx=0.5 + (path_width / 2 - 17) / canvas.winfo_width(),  # Move um pouco para a esquerda (-25)
        rely=0.5,  # Mantém alinhado verticalmente
        anchor="e",  # Ancorado na borda direita
        width=40, height=38  # Mantém o botão visível
    )

    # Adiciona um delay para garantir que os widgets foram renderizados
    canvas.after_idle(ajustar_posicoes)

    referencias_imagens["path_picker_img"] = path_picker_img

    global Paletizacao_button_image_1
    Paletizacao_button_image_1 = tk.PhotoImage(file=relative_to_assets("button_1.png"))
    tk.Button(
        image=Paletizacao_button_image_1,
        borderwidth=0,
        highlightthickness=0,
        command=lambda: threading.Thread(target=Download_button_clicked, args=(pai,), daemon=True).start(),
        relief="flat",
        bg='#FFFFFF',
        activebackground='#FFFFFF'
    ).place(x=460.0, y=390.0, width=190.0, height=100.0)
    
################################ || ########################################

def selecionar_caminho():
    global caminho_saida
    caminho_saida = filedialog.askopenfilename(
        filetypes=[
            ("Todos os arquivos suportados", "*.xlsx;*.xls;*.txt"),
            ("Excel files", "*.xlsx;*.xls"),
            ("Text files", "*.txt")
        ],
        title="Selecione um arquivo"
    )
    if caminho_saida:
        print(f"Caminho selecionado: {caminho_saida}")  # Depuração
        nome_arquivo = os.path.basename(caminho_saida)
        entrada_caminho.config(state=tk.NORMAL, bg="white"),
        entrada_caminho.delete(0, tk.END)
        entrada_caminho.insert(0, nome_arquivo)
        entrada_caminho.config(state="readonly", readonlybackground="white")


################################ || ########################################
        
    
def Download_button_clicked(root):
        global tp_exped_dropdown, transp_dropdown

        if not caminho_saida:
            root.after(100, lambda: toast.show_toast("Erro", "Insira um anexo para cálculo", duration=3, icon_path=relative_to_assets("")))
            return

        try:
            # Caminho de destino para salvar o anexo
            caminho_destino = r"G:"
            
            file_name = os.path.basename(caminho_saida)  # Pega o nome do arquivo sem o caminho completo
            
            with open(caminho_saida, 'rb') as f:
                file_data = f.read()
            
            # Salvar o anexo no caminho de destino
            with open(os.path.join(caminho_destino, file_name), 'wb') as f:
                f.write(file_data)

            win32api.MessageBox(0, "Aguarde enquanto o arquivo é gerado", "Carregando arquivo", win32con.MB_ICONINFORMATION)



            tp_exped = tp_exped_dropdown.get()
            transportadora = transp_dropdown.get()

            # Chama a função corretamente passando o nome do arquivo
            calc.main(tp_exped, transportadora, file_name)

            win32api.MessageBox(0, "O arquivo foi criado com sucesso!", "Arquivo gerado", win32con.MB_ICONINFORMATION)


        except Exception as e:
            print("Erro ao baixar anexo:", e)
       

referencias_imagens = {}