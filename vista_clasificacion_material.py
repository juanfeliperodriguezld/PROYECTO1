from tkinter import Tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import NORMAL, DISABLED
from tkinter import Label, Button
import os
from abastecimiento import main
class Ap:
    def __init__(self):
        self.path = ""
ap = Ap()
window = Tk()
window.geometry("450x250")
window.title("vista de clasificación material")
def selectroute():
    global path
    global route
    window.route = filedialog.askopenfilename(title='Selecciona el Input')
    ap.path = window.route
    extension = os.path.splitext(ap.path)[1]
    if extension.lower() in [".xlsx", ".xlsm", ".xls", ".csv"]:
        print(window.route)
        btn1.config(state=DISABLED)
        start = Button(
            window,
            text="Ejecutar",
            height="2",
            width="30",
            command=lambda: [window.iconify(), main(window.route)]
        )
        start.pack()
    elif len(ap.path) == 0:
        messagebox.showwarning(
            'Invalided ',
            "Seleccione un archivo"
        )
    else:
        messagebox.showerror(
            'Error Invalided ',
            "Debes elegir un archivo de Excel compatible"
        )
Label(
    text="Elija el archivo que desea utilizar : ",
    fg="#F4FEFF",
    bg="blue",
    width="300",
    height="2",
    font=("Corbel 17")
).pack()
Label(text="").pack()
btn1 = Button(
    window,
    text="Selecciona la ruta",
    height="2",
    width="30",
    state=NORMAL,
    command=selectroute
)
btn1.pack()
Label(text="").pack()
window.mainloop()
