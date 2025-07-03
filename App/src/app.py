
##Inicio##
import sys
import os
from collections import defaultdict
from pathlib import Path
from tkinter import *
from gui.Primary_Window.sub_gui.estadia_GUI.gui import Estadia
from gui.Primary_Window.sub_gui.paletizacao_GUI.gui import Paletizacao
from gui.Primary_Window.sub_gui.armazenagem_GUI.gui import armazenagem
from gui.Primary_Window.sub_gui.descarga_GUI.gui import Descarga
from gui.Primary_Window.sub_gui.dedicado_GUI.gui import Dedicado
import webbrowser
import threading
from PIL import Image, ImageTk

##############Buscar Imagens################

# OUTPUT_PATH = Path(__file__).parent
# ASSETS_PATH = OUTPUT_PATH / Path("./gui/Primary_Window/assets")
# def relative_to_assets(path: str) -> Path:
#     return ASSETS_PATH / Path(path)

def relative_to_assets(path: str) -> str:
    """
    Retorna o caminho absoluto para o arquivo de assets.
    Funciona com PyInstaller (sys._MEIPASS) e em desenvolvimento.
    """
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, "gui", "Primary_Window", "assets", path)

############ Menu #################

def handle_button_press(btn_name):
    global current_window
    if btn_name == "Paletização":
        home_button_clicked()
        current_window = Paletizacao(window)
    elif btn_name == "armazenagem":
        armazenagem_button_clicked()
        current_window = armazenagem(window)
    elif btn_name=="Descarga":
        descarga_button_clicked()
        current_window = Descarga(window)
    elif btn_name == "Estadia":
        Estadia_button_clicked()
        current_window=Estadia(window)
    elif btn_name == "Dedicado":
        Dedicado_button_clicked()
        current_window=Dedicado(window)   
  

def home_button_clicked(): # (coordinates : x= 0 , y= 133)
    print("Paletização button clicked")
    canvas.itemconfig(page_navigator, text="Paletização")
    sidebar_navigator.place(x=0, y=133)    

def armazenagem_button_clicked(): # (coordinates : x= 0 , y= 184)
    print("armazenagem button clicked")
    canvas.itemconfig(page_navigator, text="Armazenagem")
    sidebar_navigator.place(x=0, y=184)

def descarga_button_clicked():
    print("descarga button clicked")
    canvas.itemconfig(page_navigator, text="Descarga")
    sidebar_navigator.place(x=0, y=232)

def Estadia_button_clicked(): # (coordinates : x= 0 , y= 232)
    print("Estadia button clicked")
    canvas.itemconfig(page_navigator, text="Estadia")
    sidebar_navigator.place(x=0, y=280)

def Dedicado_button_clicked(): # (coordinates : x= 0 , y= 232)
    print("Dedicado button clicked")
    canvas.itemconfig(page_navigator, text="Dedicado")
    sidebar_navigator.place(x=0, y=328)    
    
############Título################
    
window = Tk()
window.title("Validador de Custo RCA")
window.geometry("930x506")
window.configure(bg = "#171435")


'''For Icon'''
window.iconbitmap(relative_to_assets("icon2.ico"))

img = PhotoImage(file=relative_to_assets("image_1.png"))
img = PhotoImage(file=relative_to_assets("bi.png"))
img = PhotoImage(file=relative_to_assets(""))


canvas = Canvas(
    window,
    bg = '#171435',
    height = 506,
    width = 930,
    bd = 0,
    highlightthickness = 0,
    relief = "ridge"
)
canvas.place(x = 0, y = 0)
background_image = PhotoImage(
    file=relative_to_assets("image_1.png"))
image_1 = canvas.create_image(
    566.0,
    253.0,
    image=background_image
)

#############Janelas#################

current_window=Paletizacao(window)


home_button_text = Button(
    text="Paletização",
    font=("Roboto", 16),  # Define a fonte e o tamanho
    bg="#171435",
    fg="white",  # Cor do texto
    borderwidth=0,
    highlightthickness=0,
    relief="sunken",
    activebackground="#171435",
    activeforeground="#171435",
    anchor="w"  # Alinha o texto à esquerda
)
home_button_text.place(
    x=30,
    y=133.0,
    width=170.0,
    height=47.0
)

armazenagem_button_text = Button(
    text="Armazenagem",
    font=("Roboto", 16),  # Define a fonte e o tamanho
    bg="#171435",
    fg="white",  # Cor do texto
    borderwidth=0,
    highlightthickness=0,
    relief="sunken",
    activebackground="#171435",
    activeforeground="#171435",
    anchor="w"  # Alinha o texto à esquerda
)
armazenagem_button_text.place(
    x=30,
    y=184.0,
    width=170,
    height=47.0
)

Descarga_button_text = Button(
    text="Descarga",
    font=("Roboto", 16),  # Define a fonte e o tamanho
    bg="#171435",
    fg="white",  # Cor do texto
    borderwidth=0,
    highlightthickness=0,
    relief="sunken",
    activebackground="#171435",
    activeforeground="#171435",
    anchor="w"  # Alinha o texto à esquerda
)
Descarga_button_text.place(
    x=30,
    y=232.0,
    width=170.146240234375,
    height=47.0
)

Estadia_button_text = Button(
    text="Estadia",
    font=("Roboto", 16),  # Define a fonte e o tamanho
    bg="#171435",
    fg="white",  # Cor do texto
    borderwidth=0,
    highlightthickness=0,
    relief="sunken",
    activebackground="#171435",
    activeforeground="#171435",
    anchor="w"  # Alinha o texto à esquerda
)
Estadia_button_text.place(
    x=30,
    y=280.0,
    width=170.146240234375,
    height=47.0
)
Dedicado_button_text = Button(
    text="Dedicado",
    font=("Roboto", 16),  # Define a fonte e o tamanho
    bg="#171435",
    fg="white",  # Cor do texto
    borderwidth=0,
    highlightthickness=0,
    relief="sunken",
    activebackground="#171435",
    activeforeground="#171435",
    anchor="w"  # Alinha o texto à esquerda
)
Dedicado_button_text.place(
    x=30,
    y=328.0,
    width=170.146240234375,
    height=47.0
)


####### Power BI ################

def open_power_bi():
    webbrowser.open("https://app.powerbi.com/links/1Xl-qzgcZW?ctid=3f7a3df4-f85b-4ca8-98d0-08b1034e6567&pbi_source=linkShare")

bi_button_image = PhotoImage(
    file=relative_to_assets("bi.png")
)

bi_button = Button(
    image=bi_button_image,
    borderwidth=0,
    bg="#171435",
    highlightthickness=0,
    command=open_power_bi,  # Atualizado para usar a nova função
    relief="sunken",
    activebackground="#171435",
    activeforeground="#171435"
)

bi_button.place(
    x=7.351776123046875,
    y=450.0,
    width=191.146240234375,
    height=47.0
)
#####################################


##############Eventos############
home_button_text.bind("<Button-1>", lambda event: handle_button_press("Paletização"))
armazenagem_button_text.bind("<Button-1>", lambda event: handle_button_press("armazenagem"))
Descarga_button_text.bind("<Button-1>", lambda event: handle_button_press("Descarga"))
Estadia_button_text.bind("<Button-1>", lambda event: handle_button_press("Estadia"))
Dedicado_button_text.bind("<Button-1>", lambda event: handle_button_press("Dedicado"))



##################### Navigators ###############################

####### (i)  SIDEBAR NAVIGATOR #########
sidebar_navigator = Frame(background="#FFFFFF")
sidebar_navigator.place(x=0, y=133, height=47, width=7)


####### (ii)  PAGE NAVIGATOR ###########
page_navigator = canvas.create_text(
    251.0,
    37.0,
    anchor="nw",
    text="Paletização",
    fill="#171435",
    font=("Montserrat Bold", 26 * -1))



#Logo

# Carregar e redimensionar a imagem com Pillow
original_image = Image.open(relative_to_assets(""))
resized_image = original_image.resize((int(original_image.width * 0.5), int(original_image.height * 0.5)), Image.LANCZOS)

# Converter para PhotoImage

logo_image = ImageTk.PhotoImage(resized_image)

# Criar o label com a imagem
logo_label = Label(image=logo_image, bg="#171435", borderwidth=0)

# Posicionar mais à esquerda (reduzir o valor de x)
logo_label.place(x=35, y=25)  

window.resizable(False, False)
window.mainloop()


##Fim##


