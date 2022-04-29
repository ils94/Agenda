# coding=utf8

import os
import pathlib
import threading
from datetime import date
from datetime import datetime
from tkinter import *
from tkinter import messagebox
from tkinter import ttk
import psycopg2
import json

inserir_banco = "INSERT INTO AGENDA (ATENDENTE, SOLICITANTE, ASSUNTO, DATA, HORA, URGENCIA, MENSAGEM) VALUES (%s, %s, %s, %s, %s, %s, %s)"
carregar_banco = "SELECT * FROM AGENDA"
editar_banco = "UPDATE AGENDA SET ATENDENTE = %s, SOLICITANTE = %s, ASSUNTO = %s, DATA = %s, HORA = %s, URGENCIA = %s, MENSAGEM = %s WHERE ID = %s"
deletar_banco = "DELETE FROM AGENDA WHERE ID = %s"
pesquisar_banco = "SELECT * FROM AGENDA WHERE ATENDENTE ILIKE %s OR SOLICITANTE ILIKE %s OR ASSUNTO ILIKE %s OR DATA ILIKE %s OR HORA ILIKE %s OR URGENCIA ILIKE %s OR MENSAGEM ILIKE %s"

user_home = "Z:/" + str(os.getlogin())

json_arquivo = pathlib.Path(user_home + "/agenda/cfg.json")

banco = None


def fixed_map(option):
    return [elm for elm in style.map("Treeview", query_opt=option)
            if elm[:2] != ("!disabled", "!selected")]


root = Tk()

janela_width = 750
janela_height = 760

screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

x = (screen_width / 2) - (janela_width / 2)
y = (screen_height / 2) - (janela_height / 2)

root.geometry("750x760+" + str(int(x)) + "+" + str(int(y)))
root.title("Agenda")
root.resizable(False, False)
root.iconbitmap("icones/agenda.ico")


def criar_json(usuario, dbName, dbUser, dbPass, dbHost):

    data = {}

    data['atendente'] = usuario
    data['dbName'] = dbName
    data['dbUser'] = dbUser
    data['dbPass'] = dbPass
    data['dbHost'] = dbHost

    json_data = json.dumps(data)

    salvar_json = open(user_home + "/agenda/cfg.json", "w")
    salvar_json.write(str(json_data))
    salvar_json.close()


def salvar_credenciais():

    credenciais = Toplevel(root)
    credenciais.geometry("300x160+" + str(int(x)) + "+" + str(int(y)))
    credenciais.resizable(False, False)
    credenciais.iconbitmap("icones/cadeado.ico")
    credenciais.title("Salvar Credenciais")
    credenciais.attributes("-topmost", True)

    def salvar():
        criar_json(entry_usuario_nome.get(), entry_dbname.get(), entry_dbuser.get(), entry_dbpass.get(),
                   entry_dbhost.get())

    labelframe_credenciais = LabelFrame(credenciais, text="Inserir dados")
    labelframe_credenciais.pack(fill=X, side=TOP, padx=5)

    label_usuario_nome = Label(labelframe_credenciais, text="Usuário:", width=10, height=1, anchor=W)

    entry_usuario_nome = Entry(labelframe_credenciais, width=30)

    label_dbname = Label(labelframe_credenciais, text="DB NAME:", width=10, height=1, anchor=W)

    entry_dbname = Entry(labelframe_credenciais, width=30)

    label_dbuser = Label(labelframe_credenciais, text="DB USER:", width=10, height=1, anchor=W)

    entry_dbuser = Entry(labelframe_credenciais, width=30)

    label_dbpass = Label(labelframe_credenciais, text="DB PASS:", width=10, height=1, anchor=W)

    entry_dbpass = Entry(labelframe_credenciais, width=30)

    label_dbhost = Label(labelframe_credenciais, text="DB HOST:", width=10, height=1, anchor=W)

    entry_dbhost = Entry(labelframe_credenciais, width=30)

    button_cre_salvar = Button(credenciais, text="Salvar", width=10, height=1, command=salvar)

    label_usuario_nome.grid(row=0, column=0)
    label_dbname.grid(row=1, column=0)
    label_dbuser.grid(row=2, column=0)
    label_dbpass.grid(row=3, column=0)
    label_dbhost.grid(row=4, column=0)

    entry_usuario_nome.grid(row=0, column=1)
    entry_dbname.grid(row=1, column=1)
    entry_dbuser.grid(row=2, column=1)
    entry_dbpass.grid(row=3, column=1)
    entry_dbhost.grid(row=4, column=1)

    button_cre_salvar.pack(side=LEFT, padx=5)

    credenciais.mainloop()


def multithreading(funcao):
    threading.Thread(target=funcao).start()


def auto_completar(event):
    try:

        if json_arquivo.exists():
            with open(json_arquivo) as js:
                dados = json.load(js)

        entry_atendete.delete(0, END)
        entry_atendete.insert(0, dados["atendente"])

        js.close()

    except Exception as e:
        mensagens_de_erro(e)


def limpar_campos():
    entry_solicitante.delete(0, END)
    entry_data.delete(0, END)
    entry_hora.delete(0, END)
    entry_assunto.delete(0, END)
    text_mensagem.delete('1.0', 'end')


def data_hora():
    dia = date.today()
    dia_string = str(dia).split("-")
    dia_format = dia_string[2] + '/' + dia_string[1] + '/' + dia_string[0]
    dia_format_ext = dia_format.replace("/01/", "/jan/").replace("/02/", "/fev/").replace("/03/",
                                                                                          "/mar/").replace(
        "/04/", "/abri/").replace("/05/", "/mai/").replace("/06/", "/jun/").replace("/07/", "/jul/").replace("/08/",
                                                                                                             "/ago/").replace(
        "/09/", "/set/").replace("/10/", "/out/").replace("/11/", "/nov/").replace("/12/", "/dez/")

    hora = datetime.now()
    hora_atual = hora.strftime("%H:%M:%S")

    entry_hora.delete(0, END)
    entry_data.delete(0, END)
    entry_hora.insert(0, hora_atual)
    entry_data.insert(0, dia_format_ext)


def mensagens_de_erro(e):
    messagebox.showerror("Erro", e)


def conectar():
    global banco

    try:

        if json_arquivo.exists():
            with open(json_arquivo) as js:
                dados = json.load(js)

                DB_NAME = dados["dbName"]
                DB_USER = dados["dbUser"]
                DB_PASS = dados["dbPass"]
                DB_HOST = dados["dbHost"]

            banco = psycopg2.connect(dbname=DB_NAME, user=DB_USER, password=DB_PASS, host=DB_HOST)
            multithreading(carregar_agenda)

            js.close()

    except Exception as e:
        mensagens_de_erro(e)


def banco_queries(**kwargs):
    global banco

    variaveis = kwargs.get("variaveis")

    modificar = kwargs.get("modificar")
    carregar = kwargs.get("carregar")
    pesquisar = kwargs.get("pesquisar")

    try:
        if modificar:
            cursor = banco.cursor()
            cursor.execute(modificar, variaveis)
            banco.commit()
            carregar_agenda()
        if carregar:
            cursor = banco.cursor()
            cursor.execute(carregar)
            return cursor
        if pesquisar:
            cursor = banco.cursor()
            cursor.execute(pesquisar, variaveis)
            return cursor
    except Exception as e:
        mensagens_de_erro(e)


def inserir_agenda():
    if entry_atendete.get() == "" or entry_solicitante.get() == "" or entry_data.get() == "" or entry_hora.get() == "":
        messagebox.showerror('Erro', 'Todos os campos precisam está preenchidos!')
    else:
        pergunta = messagebox.askyesno('Inserir na agenda', 'Inserir as informações na agenda?')
        if pergunta:
            variaveis = (entry_atendete.get().upper(),
                         entry_solicitante.get().upper(),
                         entry_assunto.get().upper(),
                         entry_data.get().upper(),
                         entry_hora.get().upper(),
                         variavel.get().upper(),
                         text_mensagem.get("1.0", END).upper())
            banco_queries(modificar=inserir_banco, variaveis=variaveis)


def carregar_agenda():
    cursor = banco_queries(carregar=carregar_banco)
    inserir_visualizador(cursor)


def alterar_agenda():
    global id

    pergunta = messagebox.askyesno("Alteração",
                                   'Essa ação irá alterar o Processo com o ID: "' + str(
                                       id) + '" com os valores dos campos acima, deseja continuar?')
    if pergunta:
        variaveis = (entry_atendete.get().upper(), entry_solicitante.get().upper(), entry_assunto.get().upper(),
                     entry_data.get().upper(), entry_hora.get().upper(), variavel.get().upper(),
                     text_mensagem.get("1.0", END).upper(), id)

        banco_queries(modificar=editar_banco, variaveis=variaveis)


def deletar_agenda():
    global id

    pergunta = messagebox.askyesno("Deletar",
                                   "Deletar a entrada com o id: '" + str(id) + "'")
    if pergunta:
        variaveis = (id,)
        banco_queries(modificar=deletar_banco, variaveis=variaveis)


def pesquisar_agenda(event):
    variaveis = (
        "%" + entry_pesquisar.get() + "%",
        "%" + entry_pesquisar.get() + "%",
        "%" + entry_pesquisar.get() + "%",
        "%" + entry_pesquisar.get() + "%",
        "%" + entry_pesquisar.get() + "%",
        "%" + entry_pesquisar.get() + "%",
        "%" + entry_pesquisar.get() + "%")

    cursor = banco_queries(pesquisar=pesquisar_banco, variaveis=variaveis)
    inserir_visualizador(cursor)


def items(event):
    global id

    entry_atendete.delete(0, END)
    entry_solicitante.delete(0, END)
    entry_assunto.delete(0, END)
    entry_data.delete(0, END)
    entry_hora.delete(0, END)
    text_mensagem.delete('1.0', 'end')

    id = tv.item(tv.selection())["values"][0]

    entry_atendete.insert(0, tv.item(tv.selection())["values"][1])
    entry_solicitante.insert(0, tv.item(tv.selection())["values"][2])
    entry_assunto.insert(0, tv.item(tv.selection())["values"][3])
    entry_data.insert(0, tv.item(tv.selection())["values"][4])
    entry_hora.insert(0, tv.item(tv.selection())["values"][5])
    variavel.set(tv.item(tv.selection())["values"][6])
    text_mensagem.insert('1.0', tv.item(tv.selection())["values"][7])


def destacar_rows():
    if button_destacar['text'] == 'Com Cores':

        button_destacar.configure(text='Sem Cores')

        tv.tag_configure('NORMAL', background='yellow')
        tv.tag_configure('URGENTE', background='red')
        tv.tag_configure('RESOLVIDO', background='green')

    else:
        button_destacar.configure(text='Com Cores')
        tv.tag_configure('NORMAL', background='white')
        tv.tag_configure('URGENTE', background='white')
        tv.tag_configure('RESOLVIDO', background='white')


def inserir_visualizador(cursor):
    try:
        tv.delete(*tv.get_children())
        for row in cursor:
            tv.insert("", index="end", values=(
                row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7]), tags=(row[6],))
    except Exception as e:
        mensagens_de_erro(e)


menu_bar = Menu(root)

salvar_cfg = Menu(menu_bar, tearoff=0)

salvar_cfg.add_command(label="Salvar Credenciais", command=salvar_credenciais)

menu_bar.add_cascade(label="Menu", menu=salvar_cfg)

root.config(menu=menu_bar)

labelframe_dados = LabelFrame(root, text='Inserir dados')
labelframe_dados.pack(side=TOP, padx=5, pady=5)

label_atendente = Label(labelframe_dados, text='Atendente', width=20, height=1)

entry_atendete = Entry(labelframe_dados, width=20)
entry_atendete.bind('<FocusIn>', auto_completar)

label_solicitante = Label(labelframe_dados, text='Solicitante', width=20, height=1)

entry_solicitante = Entry(labelframe_dados, width=20)

label_data = Label(labelframe_dados, text='Data', width=20, height=1)

entry_data = Entry(labelframe_dados, width=20)

label_hora = Label(labelframe_dados, text='Hora', width=20, height=1)

entry_hora = Entry(labelframe_dados, width=20)

variavel = StringVar(root)
variavel.set("NORMAL")

label_status = Label(labelframe_dados, text='Status', width=20, height=1)

op_status = OptionMenu(labelframe_dados, variavel, 'NORMAL', 'URGENTE', 'RESOLVIDO')

button_adicionar = Button(root, text='Adicionar', width=10, height=1, command=lambda: multithreading(inserir_agenda))

button_editar = Button(root, text='Editar', width=10, height=1, command=lambda: multithreading(alterar_agenda))

button_tempo = Button(root, text='Hora e Dia', width=10, height=1, command=data_hora)

button_limpar = Button(root, text='Limpar Campos', width=15, height=1, command=limpar_campos)

label_atendente.grid(row=0, column=0)
label_solicitante.grid(row=0, column=1)
label_data.grid(row=0, column=2)
label_hora.grid(row=0, column=3)
label_status.grid(row=0, column=4)

entry_atendete.grid(row=1, column=0)
entry_solicitante.grid(row=1, column=1)
entry_data.grid(row=1, column=2)
entry_hora.grid(row=1, column=3)
op_status.grid(row=1, column=4)

button_adicionar.place(x=10, y=80)
button_editar.place(x=95, y=80)
button_tempo.place(x=180, y=80)
button_limpar.place(x=625, y=340)

labelframe_mensagem = LabelFrame(root, text='Mensagem', width=730, height=220)
labelframe_mensagem.pack(pady=30)

label_assunto = Label(labelframe_mensagem, text='Assunto:', width=20, height=1, anchor=W).place(x=5, y=0)
entry_assunto = Entry(labelframe_mensagem, width=110)
entry_assunto.place(x=56, y=2)

text_mensagem = Text(labelframe_mensagem, width=89, height=10)
text_mensagem.place(x=5, y=30)

labelframe_pesquisa = LabelFrame(root, text='Pesquisar', width=265, height=50)
labelframe_pesquisa.place(x=10, y=330)

entry_pesquisar = Entry(labelframe_pesquisa, width=40)
entry_pesquisar.bind("<Return>", pesquisar_agenda)

button_carregar = Button(root, text='Carregar', width=10, height=1, command=lambda: multithreading(carregar_agenda))

button_deletar = Button(root, text='Deletar', width=10, height=1, command=lambda: multithreading(deletar_agenda))

button_destacar = Button(root, text='Com Cores', width=10, height=1, command=destacar_rows)

labelframe_agenda = LabelFrame(root, text='Agenda', width=730, height=300)

entry_pesquisar.pack(padx=5, pady=5)

button_carregar.place(x=10, y=385)
button_destacar.place(x=95, y=385)
button_deletar.place(x=660, y=385)

style = ttk.Style()

style.map("Treeview", foreground=fixed_map("foreground"), background=fixed_map("background"))

style.configure('Treeview', rowheight=25)

xsb = ttk.Scrollbar(labelframe_agenda, orient=HORIZONTAL)
xsb.pack(side=BOTTOM, fill=X)

ysb = ttk.Scrollbar(labelframe_agenda, orient=VERTICAL)
ysb.pack(side=RIGHT, fill=Y)

tv = ttk.Treeview(labelframe_agenda, selectmode='browse', show='headings', xscrollcommand=xsb.set,
                  yscrollcommand=ysb.set)

tv.bind("<<TreeviewSelect>>", items)

xsb.config(command=tv.xview)
ysb.config(command=tv.yview)

tv['columns'] = ("ID", "Atendente", "Solicitante", "Assunto", "Data", "Hora", "Urgencia", "Mensagem")

tv.column("#0", width=2, minwidth=1)
tv.column("ID", width=25, minwidth=10)
tv.column("Atendente", width=100, minwidth=10)
tv.column("Solicitante", width=100, minwidth=10)
tv.column("Assunto", width=300, minwidth=10)
tv.column("Data", width=100, minwidth=10)
tv.column("Hora", width=100, minwidth=10)
tv.column("Urgencia", width=100, minwidth=10)
tv.column("Mensagem", width=1, minwidth=1)

tv.heading("#0", text="", anchor=W)
tv.heading("ID", text="ID", anchor=W)
tv.heading("Atendente", text="Atendente", anchor=W)
tv.heading("Solicitante", text="Solicitante", anchor=W)
tv.heading("Assunto", text="Assunto", anchor=W)
tv.heading("Data", text="Data", anchor=W)
tv.heading("Hora", text="Hora", anchor=W)
tv.heading("Urgencia", text="Urgencia", anchor=W)
tv.heading("Mensagem", text="Mensagem", anchor=W)

tv.pack(padx=5, pady=5)
labelframe_agenda.pack(side=BOTTOM, padx=5, pady=5)

atendente = pathlib.Path(user_home + "/agenda")

if not atendente.exists():
    os.makedirs(user_home + "/agenda")

conectar()

root.mainloop()
