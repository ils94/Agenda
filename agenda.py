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

inserir_query = "INSERT INTO AGENDA (ATENDENTE, SOLICITANTE, ASSUNTO, DATA, HORA, STATUS, CONCLUIDO, REABERTO, DETALHES) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)"

carregar_query = "SELECT * FROM AGENDA"

alterar_query = "UPDATE AGENDA SET ATENDENTE = %s, SOLICITANTE = %s, ASSUNTO = %s, DATA = %s, HORA = %s, STATUS = %s, DETALHES = %s WHERE ID = %s"

concluido_query = "UPDATE AGENDA SET STATUS = %s, CONCLUIDO = %s WHERE ID = %s"

reabrir_query = "UPDATE AGENDA SET STATUS = %s, CONCLUIDO = %s, REABERTO = %s WHERE ID = %s"

deletar_query = "DELETE FROM AGENDA WHERE ID = %s"

pesquisar_query = "SELECT * FROM AGENDA WHERE ATENDENTE ILIKE %s OR SOLICITANTE ILIKE %s OR ASSUNTO ILIKE %s OR DATA ILIKE %s OR HORA ILIKE %s OR STATUS ILIKE %s OR CONCLUIDO ILIKE %s OR REABERTO ILIKE %s OR DETALHES ILIKE %s"

user_home = "Z:/" + str(os.getlogin())

json_arquivo = pathlib.Path(user_home + "/agenda/cfg.json")

banco = None
timer = None
credenciais = None

atendente = ""
status = ""
id = ""


def fixed_map(option):
    return [elm for elm in style.map("Treeview", query_opt=option)
            if elm[:2] != ("!disabled", "!selected")]


root = Tk()

janela_width = 750
janela_height = 730

screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

x = (screen_width / 2) - (janela_width / 2)
y = (screen_height / 2) - (janela_height / 2)

root.geometry("750x730+" + str(int(x)) + "+" + str(int(y)))
root.title("Agenda")
root.resizable(False, False)
root.iconbitmap("icones/agenda.ico")


def usuario_inativo():
    global banco

    if banco is not None:
        banco.close()

        pergunta = messagebox.askyesno("Desconectado",
                                       "Você foi desconectado por inatividade.\n\nDeseja se reconectar com o banco?")

        if pergunta:
            multithreading(conectar)
            reset_timer()
        else:
            sys.exit()
    else:
        reset_timer()


def reset_timer(event=None):
    global timer

    if timer is not None:
        root.after_cancel(timer)

    timer = root.after(60000, usuario_inativo)


def criar_json(usuario, dbName, dbUser, dbPass, dbHost, dbPort):
    global credenciais

    try:
        data = {}

        data['atendente'] = usuario
        data['dbName'] = dbName
        data['dbUser'] = dbUser
        data['dbPass'] = dbPass
        data['dbHost'] = dbHost
        data['dbPort'] = dbPort

        json_data = json.dumps(data)

        salvar_json = open(user_home + "/agenda/cfg.json", "w")
        salvar_json.write(str(json_data))
        salvar_json.close()

        messagebox.showinfo("Salvo", "As credenciais foram salvas com sucesso.")

        credenciais.destroy()

        conectar()

    except Exception as e:
        mensagens_de_erro(e)


def salvar_credenciais():
    global credenciais

    credenciais = Toplevel(root)
    credenciais.geometry("300x180+" + str(int(x)) + "+" + str(int(y)))
    credenciais.resizable(False, False)
    credenciais.iconbitmap("icones/cadeado.ico")
    credenciais.title("Salvar Credenciais")
    credenciais.attributes("-topmost", True)

    def salvar():
        if entry_usuario_nome.get() == "" or entry_dbname.get() == "" or entry_dbuser.get() == "" or entry_dbpass.get() == "" or entry_dbhost.get() == "" or entry_dbport.get() == "":
            mensagens_de_erro("É necessário preencher todos os campos.")
        elif len(entry_usuario_nome.get()) > 10:
            mensagens_de_erro("Usuário não pode ter mais que 10 caracteres.")
        else:
            criar_json(entry_usuario_nome.get().upper(), entry_dbname.get(), entry_dbuser.get(), entry_dbpass.get(),
                       entry_dbhost.get(), entry_dbport.get())

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

    label_dbport = Label(labelframe_credenciais, text="DB PORT:", width=10, height=1, anchor=W)

    entry_dbport = Entry(labelframe_credenciais, width=30)

    button_cre_salvar = Button(credenciais, text="Salvar", width=10, height=1, command=salvar)

    label_usuario_nome.grid(row=0, column=0)
    label_dbname.grid(row=1, column=0)
    label_dbuser.grid(row=2, column=0)
    label_dbpass.grid(row=3, column=0)
    label_dbhost.grid(row=4, column=0)
    label_dbport.grid(row=5, column=0)

    entry_usuario_nome.grid(row=0, column=1)
    entry_dbname.grid(row=1, column=1)
    entry_dbuser.grid(row=2, column=1)
    entry_dbpass.grid(row=3, column=1)
    entry_dbhost.grid(row=4, column=1)
    entry_dbport.grid(row=5, column=1)

    button_cre_salvar.pack(side=LEFT, padx=5)

    try:
        if json_arquivo.exists():
            with open(json_arquivo) as js:
                dados = json.load(js)

                entry_usuario_nome.insert(0, dados["atendente"])
                entry_dbname.insert(0, dados["dbName"])
                entry_dbuser.insert(0, dados["dbUser"])
                entry_dbpass.insert(0, dados["dbPass"])
                entry_dbhost.insert(0, dados["dbHost"])
                entry_dbport.insert(0, dados["dbPort"])
        else:
            entry_dbport.insert(0, "5432")
    except Exception as e:
        mensagens_de_erro(e)

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
    dia_format_ext = dia_format.replace("/01/", "/jan/") \
        .replace("/02/", "/fev/") \
        .replace("/03/", "/mar/") \
        .replace("/04/", "/abr/") \
        .replace("/05/", "/mai/") \
        .replace("/06/", "/jun/") \
        .replace("/07/", "/jul/") \
        .replace("/08/", "/ago/") \
        .replace("/09/", "/set/") \
        .replace("/10/", "/out/") \
        .replace("/11/", "/nov/") \
        .replace("/12/", "/dez/")

    hora = datetime.now()
    hora_atual = hora.strftime("%H:%M")

    return hora_atual, dia_format_ext.upper()


def inserir_data_hora():
    dataHora = data_hora()

    entry_hora.delete(0, END)
    entry_data.delete(0, END)
    entry_hora.insert(0, dataHora[0])
    entry_data.insert(0, dataHora[1])


def mensagens_de_erro(e):
    messagebox.showerror("Erro", e)


def conectar():
    global banco
    global atendente

    try:
        if json_arquivo.exists():
            with open(json_arquivo) as js:
                dados = json.load(js)

                atendente = dados["atendente"]

                entry_atendete.delete(0, END)

                entry_atendete.insert(0, atendente)

                DB_NAME = dados["dbName"]
                DB_USER = dados["dbUser"]
                DB_PASS = dados["dbPass"]
                DB_HOST = dados["dbHost"]
                DB_PORT = dados["dbPort"]

            banco = psycopg2.connect(dbname=DB_NAME, user=DB_USER, password=DB_PASS, host=DB_HOST, port=DB_PORT)
            multithreading(carregar_agenda)

            js.close()
    except Exception as e:
        mensagens_de_erro(e)


def reconectar_banco():
    global banco

    if banco is not None:
        banco.close()

    multithreading(conectar)


def banco_queries(**kwargs):
    global banco
    global id

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
            id = ""
            return cursor
        if pesquisar:
            cursor = banco.cursor()
            cursor.execute(pesquisar, variaveis)
            id = ""
            return cursor
    except Exception as e:
        mensagens_de_erro(e)


def inserir_agenda():
    global id

    if entry_atendete.get() == "" or entry_solicitante.get() == "" or entry_data.get() == "" or entry_hora.get() == "":
        messagebox.showerror('Erro', "Todos os campos precisam está preenchidos!")
    elif len(entry_atendete.get()) > 10:
        mensagens_de_erro("Atendente não pode ter mais que 10 caracteres.")
    elif len(entry_solicitante.get()) > 50:
        mensagens_de_erro("Solicitante não pode ter mais que 50 caracteres.")
    elif len(entry_data.get()) > 11:
        mensagens_de_erro("Data não pode ter mais que 11 caracteres.")
    elif len(entry_hora.get()) > 5:
        mensagens_de_erro("Hora não pode ter mais que 5 caracteres.")
    elif len(entry_assunto.get()) > 50:
        mensagens_de_erro("Assunto não pode ter mais que 50 caracteres.")
    else:
        pergunta = messagebox.askyesno("Inserir na agenda", "Inserir as informações na agenda?")
        if pergunta:
            variaveis = (entry_atendete.get().upper(),
                         entry_solicitante.get().upper(),
                         entry_assunto.get().upper(),
                         entry_data.get().upper(),
                         entry_hora.get().upper(),
                         variavel.get().upper(),
                         "AINDA EM ABERTO",
                         "NINGUEM",
                         text_mensagem.get("1.0", END).upper())
            banco_queries(modificar=inserir_query, variaveis=variaveis)

            limpar_campos()


def carregar_agenda():
    cursor = banco_queries(carregar=carregar_query)
    inserir_visualizador(cursor)


def alterar_agenda():
    global id

    if status == "RESOLVIDO":
        mensagens_de_erro("Não é possível editar uma tarefa já concluída.")
    elif id == "":
        mensagens_de_erro("Selecione um item na Agenda primeiro.")
    elif len(entry_atendete.get()) > 10:
        mensagens_de_erro("Atendente não pode ter mais que 10 caracteres.")
    elif len(entry_solicitante.get()) > 50:
        mensagens_de_erro("Solicitante não pode ter mais que 50 caracteres.")
    elif len(entry_data.get()) > 11:
        mensagens_de_erro("Data não pode ter mais que 11 caracteres.")
    elif len(entry_hora.get()) > 5:
        mensagens_de_erro("Hora não pode ter mais que 5 caracteres.")
    elif len(entry_assunto.get()) > 50:
        mensagens_de_erro("Assunto não pode ter mais que 50 caracteres.")
    else:
        pergunta = messagebox.askyesno("Alteração", "Alterar a tarefa na Agenda com o id: '" + str(id) + "' ?")

        if pergunta:
            variaveis = (entry_atendete.get().upper(), entry_solicitante.get().upper(), entry_assunto.get().upper(),
                         entry_data.get().upper(), entry_hora.get().upper(), variavel.get().upper(),
                         text_mensagem.get("1.0", END).upper(), id)

            banco_queries(modificar=alterar_query, variaveis=variaveis)


def concluido_agenda():
    global id
    global atendente
    global status

    dataHora = data_hora()

    if id == "":
        mensagens_de_erro("Selecione um item na Agenda primeiro.")
    elif status == "RESOLVIDO":
        mensagens_de_erro("A tarefa já está marcada como resolvida.")
    else:
        pergunta = messagebox.askyesno("Concluir Tarefa", "Marcar a tarefa com o ID: '" + str(id) + "' como concluída?")

        if pergunta:
            variaveis = ("RESOLVIDO", atendente.upper() + " - " + dataHora[1] + " - " + dataHora[0], id)

            banco_queries(modificar=concluido_query, variaveis=variaveis)


def reabrir_tarefa():
    global id
    global status
    global atendente

    dataHora = data_hora()

    if id == "":
        mensagens_de_erro("Selecione um item na Agenda primeiro.")
    elif status == "URGENTE" or status == "NORMAL":
        mensagens_de_erro("A tarefa já está em aberto.")
    else:
        pergunta = messagebox.askyesno("Reabrir Tarefa", "Reabrir a tarefa com o ID: '" + str(id) + "' ?")

        if pergunta:
            variaveis = (variavel.get(), "AINDA EM ABERTO",
                         atendente.upper() + " - " + dataHora[1] + " - " + dataHora[0], id)

            banco_queries(modificar=reabrir_query, variaveis=variaveis)


def deletar_agenda():
    global id

    if id == "":
        mensagens_de_erro("Selecione um item na Agenda primeiro.")
    else:
        pergunta = messagebox.askyesno("Deletar",
                                       "Deletar a entrada com o id: '" + str(id) + "'")
        if pergunta:
            variaveis = (id,)
            banco_queries(modificar=deletar_query, variaveis=variaveis)


def pesquisar_agenda(event):
    variaveis = (
        "%" + entry_pesquisar.get() + "%",
        "%" + entry_pesquisar.get() + "%",
        "%" + entry_pesquisar.get() + "%",
        "%" + entry_pesquisar.get() + "%",
        "%" + entry_pesquisar.get() + "%",
        "%" + entry_pesquisar.get() + "%",
        "%" + entry_pesquisar.get() + "%",
        "%" + entry_pesquisar.get() + "%",
        "%" + entry_pesquisar.get() + "%")

    cursor = banco_queries(pesquisar=pesquisar_query, variaveis=variaveis)
    inserir_visualizador(cursor)


def items(event):
    global id
    global status

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
    variavel.set(tv.item(tv.selection())["values"][6].replace("RESOLVIDO", "NORMAL"))
    status = tv.item(tv.selection())["values"][6]
    text_mensagem.insert('1.0', tv.item(tv.selection())["values"][9])


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
                row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9]), tags=(row[6],))
    except Exception as e:
        mensagens_de_erro(e)


menu_bar = Menu(root)

menu = Menu(menu_bar, tearoff=0)

menu.add_command(label="Salvar Credenciais", command=salvar_credenciais)
menu.add_command(label="Reconectar", command=reconectar_banco)

menu_bar.add_cascade(label="Menu", menu=menu)

root.config(menu=menu_bar)

labelframe_1 = LabelFrame(root)
labelframe_1.pack(fill=X, padx=5)

label_atendente = Label(labelframe_1, text='Atendente', width=20, height=1)
entry_atendete = Entry(labelframe_1, width=20)

label_solicitante = Label(labelframe_1, text='Solicitante', width=20, height=1)
entry_solicitante = Entry(labelframe_1, width=20)

label_data = Label(labelframe_1, text='Data', width=20, height=1)
entry_data = Entry(labelframe_1, width=20)

label_hora = Label(labelframe_1, text='Hora', width=20, height=1)
entry_hora = Entry(labelframe_1, width=20)

variavel = StringVar(root)
variavel.set("NORMAL")

label_status = Label(labelframe_1, text="Status", width=20, height=1)
op_status = OptionMenu(labelframe_1, variavel, "NORMAL", "URGENTE")

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

frame_1 = Frame(root)
frame_1.pack(fill=X, pady=5)

button_adicionar = Button(frame_1, text='Adicionar', width=10, height=1, command=lambda: multithreading(inserir_agenda))
button_editar = Button(frame_1, text='Editar', width=10, height=1, command=lambda: multithreading(alterar_agenda))
button_tempo = Button(frame_1, text='Hora e Dia', width=10, height=1, command=inserir_data_hora)

button_adicionar.pack(side=LEFT, padx=5)
button_editar.pack(side=LEFT, padx=5)
button_tempo.pack(side=LEFT, padx=5)

labelframe_2 = LabelFrame(root)
labelframe_2.pack(fill=X, padx=5)

frame_2 = Frame(labelframe_2)
frame_2.pack(fill=X, pady=5)

label_assunto = Label(frame_2, text="Assunto:", width=10, height=1)
label_assunto.pack(side=LEFT)

entry_assunto = Entry(frame_2, width=107)
entry_assunto.pack(side=LEFT, padx=5)

text_mensagem = Text(labelframe_2, width=89, height=10)
text_mensagem.pack(padx=5, pady=5)

frame_3 = Frame(root)
frame_3.pack(fill=X)

labelframe_pesquisa = LabelFrame(frame_3, text="Pesquisar:", width=265, height=50)
labelframe_pesquisa.pack(side=LEFT, padx=5)

entry_pesquisar = Entry(labelframe_pesquisa, width=40)
entry_pesquisar.pack(side=LEFT, padx=5, pady=5)
entry_pesquisar.bind("<Return>", pesquisar_agenda)

entry_pesquisar.pack(padx=5, pady=5)

button_limpar = Button(frame_3, text='Limpar Campos', width=15, height=1, command=limpar_campos)

button_limpar.pack(side=TOP, anchor=E, padx=5, pady=5)

frame_4 = Frame(root)
frame_4.pack(fill=X, pady=5)

button_carregar = Button(frame_4, text="Carregar", width=10, height=1, command=lambda: multithreading(carregar_agenda))
button_destacar = Button(frame_4, text="Com Cores", width=10, height=1, command=destacar_rows)
button_concluido = Button(frame_4, text="Concluído", width=10, height=1,
                          command=lambda: multithreading(concluido_agenda))
button_reabrir = Button(frame_4, text="Reabrir", width=10, height=1, command=lambda: multithreading(reabrir_tarefa))
button_deletar = Button(frame_4, text="Deletar", width=10, height=1, command=lambda: multithreading(deletar_agenda))

button_carregar.pack(side=LEFT, padx=5)
button_destacar.pack(side=LEFT, padx=5)
button_concluido.pack(side=LEFT, padx=5)
button_reabrir.pack(side=LEFT, padx=5)
button_deletar.pack(side=RIGHT, padx=5)

labelframe_3 = LabelFrame(root, text="Agenda")
labelframe_3.pack(padx=5)

style = ttk.Style()

style.map("Treeview", foreground=fixed_map("foreground"), background=fixed_map("background"))

style.configure('Treeview', rowheight=25)

xsb = ttk.Scrollbar(labelframe_3, orient=HORIZONTAL)
xsb.pack(side=BOTTOM, fill=X)

ysb = ttk.Scrollbar(labelframe_3, orient=VERTICAL)
ysb.pack(side=RIGHT, fill=Y)

tv = ttk.Treeview(labelframe_3, selectmode='browse', show='headings', xscrollcommand=xsb.set,
                  yscrollcommand=ysb.set)

tv.bind("<<TreeviewSelect>>", items)

xsb.config(command=tv.xview)
ysb.config(command=tv.yview)

tv['columns'] = (
    "ID", "Atendente", "Solicitante", "Assunto", "Data", "Hora", "Status", "Concluído por", "Reaberto por", "Detalhes")

tv.column("#0", width=2, minwidth=1)
tv.column("ID", width=25, minwidth=10)
tv.column("Atendente", width=100, minwidth=10)
tv.column("Solicitante", width=100, minwidth=10)
tv.column("Assunto", width=300, minwidth=10)
tv.column("Data", width=100, minwidth=10)
tv.column("Hora", width=100, minwidth=10)
tv.column("Status", width=100, minwidth=10)
tv.column("Concluído por", width=200, minwidth=10)
tv.column("Reaberto por", width=200, minwidth=10)
tv.column("Detalhes", width=1, minwidth=1)

tv.heading("#0", text="", anchor=W)
tv.heading("ID", text="ID", anchor=W)
tv.heading("Atendente", text="Atendente", anchor=W)
tv.heading("Solicitante", text="Solicitante", anchor=W)
tv.heading("Assunto", text="Assunto", anchor=W)
tv.heading("Data", text="Data", anchor=W)
tv.heading("Hora", text="Hora", anchor=W)
tv.heading("Status", text="Status", anchor=W)
tv.heading("Concluído por", text="Concluído por", anchor=W)
tv.heading("Reaberto por", text="Reaberto por", anchor=W)
tv.heading("Detalhes", text="Detalhes", anchor=W)

tv.pack(padx=5, pady=5)
labelframe_3.pack()

atendente = pathlib.Path(user_home + "/agenda")

if not atendente.exists():
    os.makedirs(user_home + "/agenda")

multithreading(conectar)

root.bind_all("<Any-KeyPress>", reset_timer)
root.bind_all("<Any-ButtonPress>", reset_timer)

root.mainloop()
