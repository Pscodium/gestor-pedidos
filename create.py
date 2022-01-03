import tkinter as tk
from tkinter import *
from tkinter import ttk
import data
from tkinter import messagebox
import pandas as pd
from random import randint

## Cores
color1 = '#ffffff'
color2 = '#7CEBEA'
color3 = '#D6EB4D'
color4 = '#AB4DEB'
color5 = '#EB8F59'
selected ='#66d1d0' 


## Cria uma planilha para cada pedido
def create_spreadsheet():
    x = randint(1, 9999)
    vcon = data.ConnectDB()
    query = 'SELECT * FROM tb_spreadsheet order by CODE'
    lista = data.fill(vcon,query)
    dados = pd.DataFrame(list(lista), columns=["Código", "Produto", "Média", "Qtd", "Condição"])    
    dados.to_excel(f'faturamento_{x}.xlsx', sheet_name='faturamento', index=False)


## Deleta todos os itens do banco de dados tb_spreadsheet
def clear():
    condition='A'
    vcon=data.ConnectDB()
    query="DELETE FROM tb_spreadsheet"
    data.clean(vcon, query)
    banco_fill()

## Deleta itens do banco de dados 
def delete():
    try:
        itemSelection = app.selection()[0]
        valores = app.item(itemSelection, "values")
        code = valores[0]
        query = "DELETE FROM tb_spreadsheet WHERE CODE="+code
        vcon = data.ConnectDB()
        data.delete(vcon, query)
    except:
        messagebox.showinfo(title="ERRO", message="Selecione um serviço")
        return
    banco_fill()

## Atualização conexão com o banco de dados. (roda basicamente à todo comando, pois é responsável por atualizar os produtos que estão na tabela com os que foram adicionados no banco de dados)
def banco_fill():
        app.delete(*app.get_children())
        vquery="SELECT * FROM tb_spreadsheet order by CODE"
        vcon = data.ConnectDB()
        linhas = data.fill(vcon,vquery)
        for i in linhas:
            app.insert("","end",values=i) 


## Seleciona os códigos dos produtos dentro do banco de dados e retorna uma lista para a combobox
def listagem():
    listas = []
    query = "SELECT CODE FROM tb_products order by CODE"
    vcon = data.ConnectDB()
    lista = data.fill(vcon,query)
    listas = list(lista)   
    return listas
        

combolist = listagem()

## Seleciona itens requisitados dentro da tabela tb_product do banco de dados e insere essas infomações dentro da tabela create_spreadsheet
def adding():
    id = combo.get()
    quantidade = entrada_quantidade.get()
    if id == '' or quantidade == '':
        messagebox.showinfo(title="ERRO", message="Preencha todas as informações")
        return
    vcon = data.ConnectDB()
    query = "SELECT * FROM tb_products WHERE CODE="+id
    linhas = data.select(vcon, query)
    lista = list(linhas)
    codigo = lista[0]
    produto = lista[1]
    media = lista[2]
    condition = 'A'
    sql = "INSERT INTO tb_spreadsheet (CODE, PRODUCT, MEAN, VALUE, CONDITION) VALUES('"+str(codigo)+"','"+produto+"','"+str(media)+"', '"+str(quantidade)+"','"+condition+"')"
    data.insert(vcon, sql)
    banco_fill()
    combo.delete(0, END)
    entrada_quantidade.delete(0, END)
 
## Janela principal
window = tk.Tk()
window.title('Criador de Produtos')
window.configure(bg=color2)
window.geometry("800x700")
window.resizable(False, False)

## Tabela principal
app = ttk.Treeview(window, columns=('codigo','produto','media','qtd'), show='headings', height=23)
app.column('codigo', minwidth=0, width=100)
app.column('produto', minwidth=0, width=300)
app.column('media', minwidth=0, width=100)
app.column('qtd', minwidth=0, width=100)
app.heading('codigo', text='código')
app.heading('produto', text='produto')
app.heading('media', text='media')
app.heading('qtd', text='qtd')
app.place(x=100,y=70)

## Label e Combobox dos produtos
lbcombo = Label(window, text='Produto', fg='black', bg=color2)
lbcombo.place(x=100, y=612)
combo = ttk.Combobox(window, values=combolist, width=15, justify='center')
combo.place(x=100, y=632)

## Label e Entrada de quantidade
lbquantidade = Label(window, text='Quantidade', fg='black', bg=color2)
lbquantidade.place(x=250, y=612)
entrada_quantidade = Entry(window, width=10, justify='center', bg=color1)
entrada_quantidade.place(x=250, y=632)

## Botão Inserir
badding = Button(window, text='Inserir', width=10, height=0,bg=color4,fg=color1,border=0,command=adding)
badding.place(x=370,y=630)

## Botão Deletar
bdelete = Button(window, text='Deletar', width=10, height=0,bg=color4,fg=color1,border=0,command=delete)
bdelete.place(x=480,y=630)

## Botão Limpar
bdelete = Button(window, text='Limpar', width=10, height=0,bg=color5,fg=color1,border=0,command=clear)
bdelete.place(x=596,y=551)

## Botão Planilha
bspread = Button(window, text='Planilha', width=10, height=0,bg=color3,fg=color1,border=0,command=create_spreadsheet)
bspread.place(x=100,y=551)

banco_fill() ## Atualiza conexão com banco de dados
window.mainloop() ## Atualiza visualização da janela