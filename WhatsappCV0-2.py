from random import randint
import pywhatkit, time, pyautogui, openpyxl
from tkinter import *
from tkinter import filedialog

def abrir_excel():
    seleccionMSJ = 1
    archivo = filedialog.askopenfilename(initialdir='/', title='Seleccione el Excel pertinente', filetypes=(('Archivo Excel','*.xlsx'),('Cualquiera','*.*')))
    
    if archivo[-4:] == 'xlsx':
        Label(text=f'El Excel seleccionado es: {archivo}').place(x=20, y=340)
        excel = openpyxl.load_workbook(archivo)

        excelActivo = excel.active
        celdas = excelActivo[desdeVar.get() : hastaVar.get()]
        msg1 = msg.get(1.0, "end")
        
        for fila in celdas:
            for celda in fila:
                seleccionMSJ = randint(1,6)
                if seleccionMSJ == 7: seleccionMSJ = 1
                aviso = Label(text=f'[+] ENVIANDO MENSAJE A: {celda.value}', foreground='green').place(x=20, y=370)
                print(f"[+] ENVIANDO MENSAJE A: {celda.value}")
                print(f"num= {seleccionMSJ}")
                time.sleep(3)
                pywhatkit.sendwhatmsg_instantly(celda.value, msg1, 15)
                time.sleep(int(TiempoEsperaVar.get()))
                pyautogui.click(1326, 698)
                time.sleep(int(TiempoEsperaVar.get()))
                pyautogui.hotkey('ctrl', 'w')

        aviso = Label(text=f'[+] TERMINADO CON EXITO!!!', foreground='green').place(x=20, y=370)

    else:
        verif1=Label(text='IMPORTANTE: SI NO SELECCIONA UN ARCHIVO EXCEL EL PROCESO NO FUNCIONARA', foreground='red').place(x=20, y=40)
        verif2=Label(text='VUELVA A SELECCION UN ARCHIVO', foreground='red').place(x=20, y=60)


ventana = Tk()

ventana.title('Envio Whatsapp Automatico')
ventana.config(width=450, height=500)

Label(text='Escriba el Texto a Enviar a sus Destinatarios: ').place(x=20, y=15)
msg=Text(width=50, height=10, yscrollcommand=True)
msg.place(x=20, y=40)

Label(text='Coloque desde y hasta donde estan los datos a entrar en el Excel').place(x=20, y=210)
Label(text='DESDE').place(x=20, y=230)
Label(text='HASTA').place(x=75, y=230)

desdeVar=StringVar()
desde = Entry(ventana, width=6, textvariable=desdeVar)
desde.place(x=25, y=255)

hastaVar = StringVar()
hasta = Entry(ventana, width=6, textvariable=hastaVar)
hasta.place(x=80, y=255)

Label(text='Elija el tiempo de espera entre tarea y tarea: ').place(x=20, y=280)

TiempoEsperaVar=StringVar()
TiempoEspera = Entry(ventana, width=4, textvariable=TiempoEsperaVar)
TiempoEspera.place(x=260, y=280)

textoInsExcel = Label(text='Seleccione el Excel con los numeros: ').place(x=20, y=315)
botonBuscarExcel = Button(text='Buscar Excel', command=abrir_excel).place(x=225, y=315)

ventana.mainloop()