import threading
from queue import Queue
from random import random
from tkinter import filedialog
from tkinter.filedialog import asksaveasfile
from tkinter.ttk import Progressbar
from shutil import copyfile
import os
import xlsxwriter
import json
from datetime import datetime
import requests
import time
from tkinter import messagebox
from tkinter import *
import PIL
from PIL import Image, ImageTk


try:
    import Tkinter as tk
except:
    import tkinter as tk

global cont
global n_2
global m_e
global lista_err
global lista_det
global n_thre
global data
global num_analisi

num_analisi=1

data = []

cont = 0
lista_err = []
lista_det = []
n_e = 2
m_e = 0


class StartApp(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.title("SSL SCAN GUI")
        self._frame = None
        directory_path = os.path.dirname(__file__)
        file_path = os.path.join(directory_path, 'icona_pri.png')
        photo_exit = PhotoImage(file=file_path)
        self.iconphoto(False, photo_exit)

        self.switch_frame(StartPage)
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def switch_frame(self, frame_class):
        new_frame = frame_class(self)
        if self._frame is not None:
            self._frame.destroy()
        self._frame = new_frame
        self._frame.pack()


    def on_closing(self):
        directory_path = os.path.dirname(__file__)
        file_path = os.path.join(directory_path, 'icona_pri.ico')
        self.iconbitmap(file_path)
        directory_path = os.path.dirname(__file__)
        file_path = os.path.join(directory_path, 'icona_pri.png')
        photo_exit = PhotoImage(file=file_path)
        self.iconphoto(False, photo_exit)
        if messagebox.askokcancel("Exit", "Vuoi uscire dal programma ?"):
            try:
                os.remove("temp_json.json")
            except IOError:
                print("")
            try:
                os.remove("temp_excel.xlsx")
            except IOError:
                print("")
            try:
                os.remove("errore.xlsx")
            except IOError:
                print("")
            self.destroy()


class StartPage(tk.Frame):
    def __init__(self, master, bg="blank"):
        tk.Frame.__init__(self, master, bg='snow')
        self.master.geometry("400x400")
        self.master.resizable(width=FALSE, height=FALSE)
        directory_path = os.path.dirname(__file__)
        file_path = os.path.join(directory_path, 'icona_pri.png')
        load = PIL.Image.open("icona_pri.png")
        render = ImageTk.PhotoImage(load)
        img = Label(self, image=render, bg='snow')
        img.image = render


        tk.Label(self, bg='snow', text="HOME", font=('Helvetica', 18, "bold")).pack(side="top", fill="x", pady=6)

        img.pack()
        tk.Button(self, text="Scansione Host Singolo",
                  command=lambda: master.switch_frame(ScansioneHost)).pack(fill="both",pady=6)
        tk.Button(self, text="Scansione Lista Host",
                  command=lambda: master.switch_frame(ScansioneListaHost)).pack(fill="both",pady=6)
        tk.Button(self, text="Analisi",
                  command=lambda: master.switch_frame(AnalisiJs)).pack(fill="both", pady=6)


class ScansioneHost(tk.Frame):

    def __init__(self, master):
        tk.Frame.__init__(self, master, bg='snow')
        tk.Frame.configure(self)
        tk.Label(self, bg='snow', text="Scansione Host Singolo", font=('Helvetica', 18, "bold")).pack(side="top", fill="x", pady=6)
        directory_path = os.path.dirname(__file__)
        file_path = os.path.join(directory_path, 'icona_pri.png')
        load = PIL.Image.open(file_path)
        render = ImageTk.PhotoImage(load)
        img = Label(self, image=render, bg='snow')
        img.image = render
        img.pack()

        #Label(self, text="Inserisci nome file excel da creare").pack()
        #box_nome_excel = Entry(self)
        #box_nome_excel.pack()
        #Label(self, text="Inserisci il nome del file json da creare").pack()
        #box_nome_json = Entry(self)
        #box_nome_json.pack()
        Label(self, text="Inserisci l'host da analizzare",bg='snow').pack(pady=6)
        box_nome_host = Entry(self,bg='snow')
        box_nome_host.pack(pady=6)

        send_button_start = Button(self, text="Conferma",
                                   command=lambda: analisi_host(self, box_nome_host.get()))
        send_button_start.pack(fill="both", pady=6)

        tk.Button(self, text="Home",
                  command=lambda: master.switch_frame(StartPage)).pack(fill="both", pady=6)


class ThreadScanner(threading.Thread):
    global data

    def __init__(self, nome, linea):
        threading.Thread.__init__(self)
        self.nome = nome
        self.linea = linea

    def run(self):
        global n_thre
        global lista_err
        global lista_det

        params = (
            ('host', self.linea),
            ('all', 'done'),
        )
        response = requests.get('https://api.ssllabs.com/api/v3/analyze', params=params)
        print("\nAnalisi in corso di : " + self.linea)
        time.sleep(15)
        time_to_fail = 0

        while True:
            if 'errors' not in response.json():
                break
            elif 'errors' in response.json():
                if response.json()['errors'][0]['message'] == 'Running at full capacity. Please try again later.':
                    print(
                        "\nCapacità di scansione elevata. La scansione dell'host " + self.linea + " verra' riprovata in seguito")
                    print("\nIl processo potrebbe rallentare")
                    if time_to_fail > 30:
                        break
                    if n_thre < 5:
                        time_to_fail = time_to_fail + 1
                        ran = random.randint(1, 40)
                        time_to_sleep = 60 + ran
                        time.sleep(time_to_sleep)
                        time_to_fail = time_to_fail + 1
                        response = requests.get('https://api.ssllabs.com/api/v3/analyze', params=params)

                    else:
                        time_to_fail = time_to_fail + 1
                        ran = random.randint(1, 40)
                        time_to_sleep = 15 + ran
                        time.sleep(time_to_sleep)
                        time_to_fail = time_to_fail + 1
                        response = requests.get('https://api.ssllabs.com/api/v3/analyze', params=params)
                else:
                    break

        if 'errors' not in response.json():
            while response.json()['status'] != 'READY':
                if response.json()['status'] == "ERROR":
                    print("\nL'host " + self.linea + " non e' stato scansionato ")
                    lista_det.append(response.json()['statusMessage'])
                    lista_err.append(self.linea)
                    break
                else:
                    time.sleep(20)
                    response = requests.get('https://api.ssllabs.com/api/v3/analyze', params=params)
            if response.json()['status'] == 'READY':
                if response.json()['endpoints'][0]['statusMessage'] != "Ready":
                    lista_det.append(response.json()['endpoints'][0]['statusMessage'])
                    lista_err.append(self.linea)

        else:
            print("\nHost " + self.linea + " non scansionabile ")
            lista_det.append(response.json()['errors'][0]['message'])
            lista_err.append(self.linea)

        print("\nAnalisi di " + self.linea + " terminata !!")

        global cont
        cont = cont - 1
        print("response : " + str(response.json()))
        data.append(response.json())


class ScansioneListaHost(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master, bg='snow')
        tk.Frame.configure(self)
        tk.Label(self, bg='snow', text="Scansione Lista di Host", font=('Helvetica', 18, "bold")).pack(side="top", fill="x",pady=6)
        directory_path = os.path.dirname(__file__)
        file_path = os.path.join(directory_path, 'icona_pri.png')
        load = PIL.Image.open(file_path)
        render = ImageTk.PhotoImage(load)
        img = Label(self, image=render, bg='snow')
        img.image = render
        img.pack()

        #Label(self, text="Inserisci nome file excel da creare").pack()
        #box_nome_excel = Entry(self)
        #box_nome_excel.pack()
        #Label(self, text="Inserisci il nome del file json da creare").pack()
        #box_nome_json = Entry(self)
        #box_nome_json.pack()
        Label(self, bg='snow',text="Inserisci la lista di elementi (.txt)").pack(pady=6)
        global folder_path
        folder_path = StringVar()
        lbl1 = Label(self, bg='snow', textvariable=folder_path).pack(pady=6)
        button_sfoglia = Button(self, text="Sfoglia", command=lambda: browse_button()).pack(fill="both", pady=6)

        send_button_start = Button(self, text="Conferma",
                                   command=lambda: analisi_lista_host(self, folder_path))
        send_button_start.pack(fill="both", pady=6)

        tk.Button(self, text="Home",
                  command=lambda: master.switch_frame(StartPage)).pack(fill="both", pady=6)


def browse_button():
    global folder_path
    filename = filedialog.askopenfilename()
    folder_path.set(filename)


class AnalisiJs(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master, bg='snow')
        tk.Frame.configure(self)
        tk.Label(self,bg='snow', text="Analisi", font=('Helvetica', 18, "bold")).pack(side="top", fill="x", pady=6)
        directory_path = os.path.dirname(__file__)
        file_path = os.path.join(directory_path, 'icona_pri.png')
        load = PIL.Image.open(file_path)
        render = ImageTk.PhotoImage(load)
        img = Label(self, image=render, bg='snow')
        img.image = render
        img.pack()
        #Label(self, text="Inserisci nome file excel da creare").pack()
        #box_nome_excel = Entry(self)
        #box_nome_excel.pack()
        Label(self, bg='snow', text="Inserisci il nome del file da analizzare").pack(pady=6)
        #box_nome_json = Entry(self)
        #box_nome_json.pack(pady=6)
        global folder_path
        folder_path = StringVar()
        lbl2 = Label(self, bg='snow', textvariable=folder_path).pack(pady=6)
        button_sfoglia = Button(self, text="Sfoglia", command=lambda: browse_button()).pack(fill="both", pady=6)
        send_button_start = Button(self, text="Conferma",
                                   command=lambda: esecuzione_excel(self, folder_path.get()))
        send_button_start.pack(fill="both", pady=6)

        tk.Button(self, text="Home",
                  command=lambda: master.switch_frame(StartPage)).pack(fill="both", pady=6)


def analisi_lista_host(self, nome_doc):

    try:
        os.remove("temp_json.json")
    except IOError:
        print("")

    try:
        os.remove("temp_excel.xlsx")
    except IOError:
        print("")

    try:
        os.remove("errore.xlsx")
    except IOError:
        print("")

    window = tk.Toplevel(self, bg="snow")
    window.title(nome_doc.get())
    window.geometry("450x250")
    window.resizable(width=FALSE, height=FALSE)
    directory_path = os.path.dirname(__file__)
    file_path = os.path.join(directory_path, 'icona_pri.ico')
    window.iconbitmap(file_path)
    directory_path = os.path.dirname(__file__)
    file_path = os.path.join(directory_path, 'icona_pri.png')
    load = PIL.Image.open(file_path)
    render = ImageTk.PhotoImage(load)
    img = Label(window, image=render, bg='snow')
    img.image = render

    global data
    global cont
    data = []
    nome_doc = nome_doc.get()
    f = open(nome_doc, "r")
    righe_file = f.readlines()
    o = open("temp_json.json", "a")
    q = Queue()
    x_t = 0
    lista_thread = []
    maxm=len(righe_file)

    progress_b = Progressbar(window, orient=HORIZONTAL,
                           length=100, mode='determinate')
    progress_b.pack(fill="both",padx=6, pady=8)

    for e in righe_file:
        q.put(e)

    avanzamento = (50 / len(righe_file))
    val=0

    while x_t < len(righe_file):
        response_info = requests.get('https://api.ssllabs.com/api/v3/info')
        time.sleep(3)
        global n_thr
        n_thre = response_info.json()['maxAssessments'] - response_info.json()['currentAssessments']

        if n_thre < 3:
            n_thre = 1
        if cont < n_thre:
            time.sleep(6)
            linea = q.get()
            cont = cont + 1
            x_t = x_t + 1
            val = val+avanzamento
            progress_b['value'] = val
            progress_b.update()
            thread1 = ThreadScanner("Thread#1", linea)
            lista_thread.append(thread1)
            thread1.start()
        else:
            time.sleep(5)

    val = val+15
    progress_b['value'] = val
    progress_b.update()

    for x in lista_thread:
        x.join()

    val = val + 15
    progress_b['value'] = val
    progress_b.update()

    global lista_err
    global lista_det

    if len(lista_err) > 0:

        global n_e
        global m_e
        n_e = 2
        m_e = 0
        workbook_err = xlsxwriter.Workbook("errore.xlsx")
        worksheet_err = workbook_err.add_worksheet('Host non scansionati')
        worksheet_err.write('A1', 'Host')
        worksheet_err.write('B1', 'Errore')
        for l in lista_err:
            worksheet_err.write('A' + str(n_e), l)
            worksheet_err.write('B' + str(n_e), lista_det[m_e])
            n_e = n_e + 1
            m_e = m_e + 1
        workbook_err.close()


    progress_b['value'] = 90
    progress_b.update()
    json.dump(data, o)
    o.close()
    f.close()
    time.sleep(2)
    progress_b['value'] = 100
    progress_b.update()
    time.sleep(2)
    progress_b.destroy()
    time.sleep(1)
    Label(window, bg="snow", text="Analisi completata").pack(pady=6)
    time.sleep(0.5)
    esecuzione_excel(window, "temp_json.json")


def analisi_host(self, host_singolo):

    try:
        os.remove("temp_json.json")
    except IOError:
        print("")
    try:
        os.remove("temp_excel.xlsx")
    except IOError:
        print("")

    window = tk.Toplevel(self, bg="snow")
    window.title(host_singolo)
    window.geometry("450x250")
    window.resizable(width=600, height=False)
    directory_path = os.path.dirname(__file__)
    file_path = os.path.join(directory_path, 'icona_pri.ico')
    window.iconbitmap(file_path)
    directory_path = os.path.dirname(__file__)
    file_path = os.path.join(directory_path, 'icona_pri.png')
    load = PIL.Image.open(file_path)
    render = ImageTk.PhotoImage(load)
    img = Label(window, image=render, bg='snow')
    img.image = render
    progress = Progressbar(window, orient=HORIZONTAL,
                           length=100, mode='determinate')
    progress.pack(fill="both",padx=6,pady=15)

    #if ".json" not in file_da_analizzare:
        #file_da_analizzare = file_da_analizzare+".json"


    global data
    data = []

    o = open("temp_json.json", "w")
    params = (
        ('host', host_singolo),
        ('all', 'done'),
    )
    response = requests.get('https://api.ssllabs.com/api/v3/analyze', params=params)
    print("\nAnalisi in corso di : " + host_singolo)
    progress['value'] = 20
    progress.update()
    time.sleep(15)
    progress['value'] = 50
    progress.update()
    if 'errors' not in response.json():
        while response.json()['status'] != 'READY':
            if 'errors' in response.json():
                break
            elif response.json()['status'] == "ERROR":
                break
            else:
                time.sleep(20)
                progress['value'] = 70
                progress.update()
                response = requests.get('https://api.ssllabs.com/api/v3/analyze', params=params)
    else:
        Label(window, bg='snow', text="Host non scansionabile").pack(pady=15)
        print("\nHost non scansionabile ")

    data.append(response.json())
    json.dump(data, o)
    o.close()
    progress['value'] = 95
    progress.update()
    time.sleep(0.5)
    progress['value'] = 100
    progress.update()
    time.sleep(2)
    progress.destroy()
    time.sleep(1.5)

    if 'errors' in data[0]:
        err = data[0]['errors'][0]['message']
        print(err)
        messagebox.showerror("Error", err)
        window.destroy()
    elif data[0]['status'] == "ERROR":
        err = data[0]['statusMessage']
        print(err)
        messagebox.showerror("Error", err)
        window.destroy()
    elif data[0]['status'] == "READY":
        complete = Label(window, bg="snow", text="Analisi completata").pack(pady=15)
        esecuzione_excel(window, "temp_json.json")


def down_err(err):
    file = filedialog.asksaveasfile(mode="w", defaultextension=".xlsx")
    copyfile(err, file.name)



def save_ex(file_ex):

    file = filedialog.asksaveasfile(mode="w", defaultextension=".xlsx")
    copyfile(file_ex,file.name)


def save_js(file_js):

    file = filedialog.asksaveasfile(mode="w", defaultextension=".json")
    copyfile(file_js, file.name)

def esecuzione_excel(self, file_da_analizzare):

    global lista_err
    global lista_det
    global num_analisi


    if str(self) == ".!analisijs" or str(self) == ".!analisijs"+str(num_analisi):
        window = tk.Toplevel(self,bg="snow")
        window.title(file_da_analizzare)
        num_analisi = num_analisi+1
    else:
        window = self

    window.geometry("450x250")
    window.resizable(width=FALSE, height=FALSE)
    directory_path = os.path.dirname(__file__)
    file_path = os.path.join(directory_path, 'icona_pri.ico')
    window.iconbitmap(file_path)
    directory_path = os.path.dirname(__file__)
    file_path = os.path.join(directory_path, 'icona_pri.png')
    load = PIL.Image.open(file_path)
    render = ImageTk.PhotoImage(load)
    img = Label(window, image=render, bg='snow')
    img.image = render

    Label(window, bg="snow", text="Preparazione Documenti").pack(pady=6)
    progress_c = Progressbar(window, orient=HORIZONTAL,
                             length=100, mode='determinate')

    progress_c.pack(fill="both", padx=6, pady=6)

    progress_c['value'] = 10
    progress_c.update()
    controllo = ("temp_excel.xlsx")
    workbook = xlsxwriter.Workbook("temp_excel.xlsx")
    worksheet = workbook.add_worksheet('Generale')
    worksheet.write('A1', 'IP')
    worksheet.write('B1', 'HOSTNAME')
    worksheet.write('C1', 'NOME COMUNE')
    worksheet.write('D1', 'SCADENZA CERTIFICATO PRINCIPALE')
    worksheet.write('E1', 'GRADO')
    worksheet.write('F1', 'ALGORITMO CHIAVE')
    worksheet.write('G1', 'LUNGHEZZA CHIAVE')
    worksheet.write('H1', 'TLS 1.0')
    worksheet.write('I1', 'TLS 1.1')
    worksheet.write('J1', 'TLS 1.2')
    worksheet.write('K1', 'TLS 1.3')

    # altri certificati

    worksheet2 = workbook.add_worksheet('Certificati Aggiuntivi')
    worksheet2.write('A1', 'CERTIFICATO')
    worksheet2.write('B1', 'NOME')
    worksheet2.write('C1', 'DATA SCADENZA')
    worksheet2.write('D1', 'ALGORITMO CHIAVE')
    worksheet2.write('E1', 'LUNGHEZZA CHIAVE')
    worksheet2.write('F1', 'HOST')
    worksheet2.write('G1', 'IP')

    # weak key
    worksheet3 = workbook.add_worksheet('Weak Key')
    worksheet3.write('A1', 'NUMERO PROTOCOLLO')
    worksheet3.write('B1', 'ID')
    worksheet3.write('C1', 'TLS')
    worksheet3.write('D1', 'NOME')
    worksheet3.write('E1', 'CIPHER STRENGHT')
    worksheet3.write('F1', 'HOST')
    worksheet3.write('G1', 'IP')
    worksheet3.write('H1', 'Weak/Insecure')

    progress_c['value'] = 20
    progress_c.update()
    # host non scansionati

    with open(file_da_analizzare, 'r') as json_file:
        data = json.load(json_file)
        ip_analizzati = 0
        conta_excel = 2
        riga_exc = 1
        enumcer = 2
        conta_we = 2

        while ip_analizzati < len(data):
            oth_ip = 0
            if 'errors' in data[ip_analizzati]:
                print("\nErrore : " + data[ip_analizzati]['errors'][oth_ip]['message'] + " \n")
                lista_det.append(data[ip_analizzati]['errors'][0]['message'])
                lista_err.append(data[ip_analizzati]['host'])
            elif data[ip_analizzati]['status'] != "READY":
                print("\nErrore - Host non analizzato \n")
                lista_det.append(data[ip_analizzati]['statusMessage'])
                lista_err.append(data[ip_analizzati]['host'])
            elif data[ip_analizzati]['endpoints'][oth_ip]['statusMessage'] != "Ready":
                print("Errore : " + data[ip_analizzati]['endpoints'][oth_ip]['statusMessage'])
                lista_det.append(data[ip_analizzati]['endpoints'][oth_ip]['statusMessage'])
                lista_err.append(data[ip_analizzati]['host'])
            elif data[ip_analizzati]['status'] == 'READY':
                while oth_ip < len(data[ip_analizzati]['endpoints']):
                    if 'errors' in data[ip_analizzati]:
                        print("\nErrore : " + data[ip_analizzati]['errors'][oth_ip]['message'] + " \n")
                    elif data[ip_analizzati]['status'] != "READY":
                        print("\nErrore - Host non analizzato \n")
                    elif data[ip_analizzati]['endpoints'][oth_ip]['statusMessage'] != "Ready":
                        print("Errore : " + data[ip_analizzati]['endpoints'][oth_ip]['statusMessage'])
                    elif data[ip_analizzati]['status'] == 'READY':
                        print("---------------------------------------------------------------------")
                        print("-----------------------------" + data[ip_analizzati]['endpoints'][oth_ip][
                            'ipAddress'] + "-----------------------------")
                        print("---------------------------------------------------------------------")
                        print("Informazioni in dettaglio per l'ip " + data[ip_analizzati]['endpoints'][oth_ip][
                            'ipAddress'] + " con hostname : " + data[ip_analizzati]['host'])
                        timestamp = str(data[ip_analizzati]['startTime'])
                        time_ti = int(timestamp[:10])
                        dt_object = datetime.fromtimestamp(time_ti)
                        print("data inizio scansione (+GMT 2h) =", dt_object)
                        if 'subject' in data[ip_analizzati]['certs'][0]:
                            print("certificato primario: " + data[ip_analizzati]['certs'][0]['subject'])
                        else:
                            print("")
                        print("nome comune : " + data[ip_analizzati]['certs'][0]['commonNames'][0])
                        print("host: " + data[ip_analizzati]['host'])
                        print("port: " + str(data[ip_analizzati]['port']))
                        print("protocollo: " + data[ip_analizzati]['protocol'])
                        print("grado: " + data[ip_analizzati]['endpoints'][oth_ip]['grade'])
                        if 'serverSignature' in data[ip_analizzati]['endpoints'][oth_ip]['details']:
                            print("server signature: " + data[ip_analizzati]['endpoints'][oth_ip]['details'][
                                'serverSignature'])
                        if 'httpForwarding' in data[ip_analizzati]['endpoints'][oth_ip]['details']:
                            print("server signature: " + data[ip_analizzati]['endpoints'][oth_ip]['details'][
                                'httpForwarding'])
                        print("algoritmo chiave: " + data[ip_analizzati]['certs'][0]['keyAlg'])
                        lun_key = data[ip_analizzati]['certs'][0]['keySize']
                        print("lunghezza chiave: " + str(lun_key))
                        # timestamp
                        timestamp = str(data[ip_analizzati]['certs'][0]['notAfter'])
                        time_ti = int(timestamp[:10])
                        dt_object = datetime.fromtimestamp(time_ti)
                        print("data scadenza certificato (+GMT 2h) =", dt_object)

                        # EXCEL
                        worksheet.write('A' + str(conta_excel),
                                        data[ip_analizzati]['endpoints'][oth_ip]['ipAddress'])
                        worksheet.write('B' + str(conta_excel), data[ip_analizzati]['host'])
                        worksheet.write('C' + str(conta_excel), data[ip_analizzati]['certs'][0]['commonNames'][0])
                        worksheet.write('D' + str(conta_excel), str(dt_object))
                        worksheet.write('E' + str(conta_excel), data[ip_analizzati]['endpoints'][oth_ip]['grade'])
                        worksheet.write('F' + str(conta_excel), data[ip_analizzati]['certs'][0]['keyAlg'])
                        worksheet.write_number('G' + str(conta_excel), lun_key)
                        conta_excel = conta_excel + 1

                        progress_c['value'] = 30
                        progress_c.update()
                        # protocolli TLS
                        print("\n")
                        worksheet.write(riga_exc, 7, "No")
                        worksheet.write(riga_exc, 8, "No")
                        worksheet.write(riga_exc, 9, "No")
                        worksheet.write(riga_exc, 10, "No")
                        print("----PROTOCOLLI TLS----")
                        version_tls = 0
                        while version_tls < len(data[ip_analizzati]['endpoints'][oth_ip]['details']['protocols']):
                            if data[ip_analizzati]['endpoints'][oth_ip]['details']['protocols'][version_tls][
                                'version'] == "1.0":
                                print("TLS 1.0")
                                worksheet.write(riga_exc, 7, "Si")

                            if data[ip_analizzati]['endpoints'][oth_ip]['details']['protocols'][version_tls][
                                'version'] == "1.1":
                                print("TLS 1.1")
                                worksheet.write(riga_exc, 8, "Si")

                            if data[ip_analizzati]['endpoints'][oth_ip]['details']['protocols'][version_tls][
                                'version'] == "1.2":
                                print("TLS 1.2")
                                worksheet.write(riga_exc, 9, "Si")

                            if data[ip_analizzati]['endpoints'][oth_ip]['details']['protocols'][version_tls][
                                'version'] == "1.3":
                                print("TLS 1.3")
                                worksheet.write(riga_exc, 10, "Si")

                            version_tls = version_tls + 1

                        print("\n")
                        riga_exc = riga_exc + 1

                        progress_c['value'] = 35
                        progress_c.update()
                        # altri certificati

                        print("----ALTRI CERTIFICATI----")
                        certificati = 1
                        while certificati < len(data[ip_analizzati]['certs']):
                            print("certificato aggiuntivo: " + data[ip_analizzati]['certs'][certificati]['subject'])
                            worksheet2.write('A' + str(enumcer),
                                             data[ip_analizzati]['certs'][certificati]['subject'])
                            print("nome comune : " + data[ip_analizzati]['certs'][certificati]['commonNames'][0])
                            worksheet2.write('B' + str(enumcer),
                                             data[ip_analizzati]['certs'][certificati]['commonNames'][0])
                            timestamp = str(data[ip_analizzati]['certs'][certificati]['notAfter'])
                            time_ti = int(timestamp[:10])
                            dt_object = datetime.fromtimestamp(time_ti)
                            print("data scadenza certificato (+GMT 2h) =", dt_object)
                            worksheet2.write('C' + str(enumcer), str(dt_object))
                            print("algoritmo chiave: " + data[ip_analizzati]['certs'][certificati]['keyAlg'])
                            worksheet2.write('D' + str(enumcer),
                                             data[ip_analizzati]['certs'][certificati]['keyAlg'])
                            lun_key = data[ip_analizzati]['certs'][certificati]['keySize']
                            print("lunghezza chiave: " + str(lun_key))
                            worksheet2.write_number('E' + str(enumcer), lun_key)
                            worksheet2.write("F" + str(enumcer), data[ip_analizzati]['host'])
                            worksheet2.write("G" + str(enumcer),
                                             data[ip_analizzati]['endpoints'][oth_ip]['ipAddress'])
                            certificati = certificati + 1
                            enumcer = enumcer + 1
                            print("\n")

                        print("\n")

                        progress_c['value'] = 60
                        progress_c.update()
                        # weak key
                        print("----WEAK KEY----\n")
                        s = 0
                        while s < len(data[ip_analizzati]['endpoints'][oth_ip]['details']['suites']):
                            v = 0
                            while v < len(data[ip_analizzati]['endpoints'][oth_ip]['details']['suites'][s]['list']):
                                if 'q' in data[ip_analizzati]['endpoints'][oth_ip]['details']['suites'][s]['list'][
                                    v]:
                                    if data[ip_analizzati]['endpoints'][oth_ip]['details']['suites'][s][
                                        'protocol'] == 769:
                                        print("-----------")
                                        print("| TLS 1.0 |")
                                        print("-----------\n")
                                        protocollo = \
                                            data[ip_analizzati]['endpoints'][oth_ip]['details']['suites'][s]['protocol']
                                        print("protocol : " + str(protocollo))
                                    elif data[ip_analizzati]['endpoints'][oth_ip]['details']['suites'][s][
                                        'protocol'] == 770:
                                        print("-----------")
                                        print("| TLS 1.1 |")
                                        print("-----------\n")
                                        protocollo = \
                                            data[ip_analizzati]['endpoints'][oth_ip]['details']['suites'][s]['protocol']
                                        print("protocol : " + str(protocollo))
                                    elif data[ip_analizzati]['endpoints'][oth_ip]['details']['suites'][s][
                                        'protocol'] == 771:
                                        print("-----------")
                                        print("| TLS 1.2 |")
                                        print("-----------\n")
                                        protocollo = \
                                            data[ip_analizzati]['endpoints'][oth_ip]['details']['suites'][s]['protocol']
                                        print("protocol : " + str(protocollo))
                                    elif data[ip_analizzati]['endpoints'][oth_ip]['details']['suites'][s][
                                        'protocol'] == 772:
                                        print("-----------")
                                        print("| TLS 1.3 |")
                                        print("-----------\n")
                                        protocollo = \
                                            data[ip_analizzati]['endpoints'][oth_ip]['details']['suites'][s]['protocol']
                                        print("protocol : " + str(protocollo))

                                    # preference
                                    if 'preference' in \
                                            data[ip_analizzati]['endpoints'][oth_ip]['details']['suites'][s]:
                                        preferenza = \
                                            data[ip_analizzati]['endpoints'][oth_ip]['details']['suites'][s][
                                                'preference']
                                        print("preferenza : " + str(preferenza))
                                        print("\n")

                                    print("protocollo numero : " + str(protocollo) + " - " + str(v))
                                    worksheet3.write('A' + str(conta_we), str(protocollo) + " - " + str(v))
                                    print(
                                        "id : " + str(
                                            data[ip_analizzati]['endpoints'][oth_ip]['details']['suites'][s][
                                                'list'][v]['id']))
                                    worksheet3.write('B' + str(conta_we), str(
                                        data[ip_analizzati]['endpoints'][oth_ip]['details']['suites'][s]['list'][v][
                                            'id']))
                                    if data[ip_analizzati]['endpoints'][oth_ip]['details']['suites'][s][
                                        'protocol'] == 769:
                                        print("TLS 1.0")
                                        worksheet3.write('C' + str(conta_we), "TLS 1.0")
                                    elif data[ip_analizzati]['endpoints'][oth_ip]['details']['suites'][s][
                                        'protocol'] == 770:
                                        print("TLS 1.1")
                                        worksheet3.write('C' + str(conta_we), "TLS 1.1")
                                    elif data[ip_analizzati]['endpoints'][oth_ip]['details']['suites'][s][
                                        'protocol'] == 771:
                                        print("TLS 1.2")
                                        worksheet3.write('C' + str(conta_we), "TLS 1.2")
                                    elif data[ip_analizzati]['endpoints'][oth_ip]['details']['suites'][s][
                                        'protocol'] == 772:
                                        print("TLS 1.3")
                                        worksheet3.write('C' + str(conta_we), "TLS 1.3")
                                    print(
                                        "nome : " +
                                        data[ip_analizzati]['endpoints'][oth_ip]['details']['suites'][s]['list'][v][
                                            'name'])
                                    worksheet3.write('D' + str(conta_we),
                                                     data[ip_analizzati]['endpoints'][oth_ip]['details']['suites'][
                                                         s]['list'][v]['name'])
                                    print("cipher strenght : " + str(
                                        data[ip_analizzati]['endpoints'][oth_ip]['details']['suites'][s]['list'][v][
                                            'cipherStrength']))
                                    print("\n")
                                    worksheet3.write_number('E' + str(conta_we),
                                                            data[ip_analizzati]['endpoints'][oth_ip]['details'][
                                                                'suites'][s]['list'][v][
                                                                'cipherStrength'])
                                    worksheet3.write("F" + str(conta_we), data[ip_analizzati]['host'])
                                    worksheet3.write("G" + str(conta_we),
                                                     data[ip_analizzati]['endpoints'][oth_ip]['ipAddress'])
                                    if data[ip_analizzati]['endpoints'][oth_ip]['details']['suites'][s]['list'][v][
                                        'q'] == 1:
                                        worksheet3.write('H' + str(conta_we), "Weak")
                                    else:
                                        worksheet3.write('H' + str(conta_we), "Insecure")
                                    conta_we = conta_we + 1

                                v = v + 1
                            s = s + 1
                    oth_ip = oth_ip + 1
            else:
                print("Analisi non completa")

            progress_c['value'] = 80
            progress_c.update()
            ip_analizzati = ip_analizzati + 1

    workbook.close()
    progress_c['value'] = 100
    progress_c.update()
    time.sleep(1)
    progress_c.destroy()
    time.sleep(2)
    Label(window, bg="snow", text="Documenti Pronti").pack(pady=6)
    time.sleep(2)
    if str(self) == ".!analisijs":
        btn_ex = Button(window, text='Save Excel', command=lambda: save_ex(controllo))
        btn_ex.pack(side=TOP, pady=10)
    else:
        btn_ex = Button(window, text='Save Excel', command=lambda: save_ex(controllo))
        btn_ex.pack(side=TOP, pady=10)
        btn_js = Button(window, text='Save Json', command=lambda: save_js(file_da_analizzare))
        btn_js.pack(side=TOP, pady=10)


    if len(lista_err)>0:
        btn_error = Button(window, text='Save Error', command=lambda: down_err("errore.xlsx"))
        btn_error.pack(side=TOP, pady=10)
        lista_err = []
        lista_det = []
    else:
        print("")





if __name__ == "__main__":

    app = StartApp()
    app.configure(bg='snow')
    directory_path = os.path.dirname(__file__)
    file_path = os.path.join(directory_path, 'icona_pri.ico')
    app.iconbitmap(file_path)
    app.mainloop()

