import tkinter as tk
from tkinter import *
from tkinter import ttk
import data
from tkinter import messagebox

## Cores
color1 = '#ffffff' 
color2 = '#7CEBEA'
color3 = '#D6EB4D'
color4 = '#AB4DEB'
color5 = '#EB8F59'
selected ='#66d1d0' 

## Janela Principal
window = tk.Tk()
window.title('Criador de Produtos')
window.configure(bg=color2)
window.geometry("800x700")
window.resizable(False, False)


## Lista de mercados da combobox
mercados = ['Condor', 'Bistek', 'Hipermais Joao Costa', 'Hipermais Araquari', 'Fort Atacadista', 'Rodrigues']

## Atualiza conexão com o banco de dados do mercado selecionado mostrando assim todos os elementos dentro dele
def mercado_selected():
    app.delete(*app.get_children())
    mercado = cmercados.get()
    query = "SELECT CODE, PRODUCT, MEAN, VALUE FROM '"+mercado+"' order by CODE"
    vcon = data.ConnectDB()
    linhas = data.fill(vcon,query)
    for i in linhas:
        app.insert("","end",values=i)  
 

## Deleta item selecionado no mercado selecionado
def delete():
    try:
        mercado = cmercados.get()
        itemSelection = app.selection()[0]
        valores = app.item(itemSelection, "values")
        code = valores[0]
        query = "DELETE FROM '"+mercado+"' WHERE CODE="+code
        vcon = data.ConnectDB()
        data.delete(vcon, query)
    except:
        messagebox.showinfo(title="ERRO", message="Selecione um serviço")
        return
    mercado_selected()

## Adiciona o item ao banco de dados do mercado selecionado
def adding():
    def insert_banco():
        mercado = cmercados.get()
        code = entrada_codigo.get()
        product = entrada_produto.get()
        mean = entrada_media.get()
        vcon = data.ConnectDB()
        query="INSERT INTO '"+mercado+"' (CODE, PRODUCT, MEAN) VALUES('"+code+"','"+product+"','"+mean+"')"
        data.insert(vcon,query)
    
    ## Insere os Produtos na tabela da Treeview
    def insert():
        code = entrada_codigo.get()
        product = entrada_produto.get()
        mean = entrada_media.get()
        if code=="" or product=="" or mean=="":
            messagebox.showinfo(title='ERRO', message="Digite todos os dados")
            return
        app.insert("","end", values=(code,product,mean))
        insert_banco()
        entrada_codigo.delete(0, END)
        entrada_media.delete(0, END)
        entrada_produto.delete(0, END)
        entrada_codigo.focus()
        mercado_selected()
   

    ## Janela para adição de produtos
    janela = tk.Tk()
    janela.title('Adicione um produto')
    janela.configure(bg=color2)
    janela.geometry('400x280')
    janela.resizable(False, False)

    ## Label e Entrada do Código de Produto
    codigo = Label(janela, text='Código', fg='black', bg=color2)
    codigo.place(x=50, y=20)
    entrada_codigo = Entry(janela, width=8, justify='center', bg=color1)
    entrada_codigo.place(x=50, y=40)

    ## Label e Entrada do Nome do Produto
    produto = Label(janela, text='Nome do produto', fg='black', bg=color2)
    produto.place(x=50, y=80)
    entrada_produto = Entry(janela, width=36, justify='center', bg=color1)
    entrada_produto.place(x=50, y=100)

    ## Label e Entrada da Média de vendas diárias do Produto
    media = Label(janela, text='Média', fg='black', bg=color2)
    media.place(x=50, y=140)
    entrada_media = Entry(janela, width=10, justify='center', bg=color1)
    entrada_media.place(x=50, y=160)
    
    ## Botão responsável por adicionar produto através da função insert()
    lbadding = Button(janela, text='Adicione um produto', width=20, height=0,bg=color4,fg=color1,border=0,command=insert)
    lbadding.place(x=50,y=210)



    janela.mainloop()


## Tabela principal conectada ao banco de dados
app = ttk.Treeview(window, columns=('codigo','produto','media'), show='headings', height=23)
app.column('codigo', minwidth=0, width=100)
app.column('produto', minwidth=0, width=400)
app.column('media', minwidth=0, width=100)
app.heading('codigo', text='código')
app.heading('produto', text='produto')
app.heading('media', text='media')
app.place(x=100,y=70)

## Botão para abrir janela de adicionar produtos
badding = Button(window, text='Adicione um produto', width=20, height=0,bg=color4,fg=color1,border=0,command=adding)
badding.place(x=100,y=600)

## Botão para Deletar produtos do banco de dados
bdelete = Button(window, text='Delete', width=10, height=0,bg=color4,fg=color1,border=0,command=delete)
bdelete.place(x=596,y=551)


## Combobox com os mercados
cmercados = ttk.Combobox(window, values=mercados, width=15, justify='center')
cmercados.set("Condor")
cmercados.place(x=100, y = 30)


## Botão que seleciona mercado da combobox
bmercados = Button(window, text='Selecionar', width=10, height=0, bg=selected, fg=color1, border=0, command=mercado_selected)
bmercados.place(x=260, y= 27)


mercado_selected() ## Atualiza conexão com o banco de dados
window.mainloop()

