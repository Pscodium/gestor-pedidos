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

## lista da combobox de mercados
mercados = ['Condor', 'Bistek', 'Hipermais Joao Costa', 'Hipermais Araquari', 'Fort Atacadista', 'Rodrigues']

## Lista em uma combobox os códigos dos produtos inseridos no banco do mercado selecionado através da main.py
def listing(): ## Essa função fica responsável por atualizar a conexão com o banco de dados para retornar as informações dele
    mercado = cmercados.get()
    combo.set('')
    query = "SELECT CODE FROM '"+mercado+"' order by CODE"
    vcon = data.ConnectDB()
    linhas = data.fill(vcon, query)
    listas = list(linhas)
    combo['values'] = (listas)
    return listas 
    
## Uitilizando a combobox de mercados para conectar a tree com o banco de dados de planilha de cada mercado
def mercado_selected():
    app.delete(*app.get_children())
    mercado = cmercados.get()
    if mercado == "Condor":
        market = "Condor_pl"
    elif mercado == "Bistek":
        market = "Bistek_pl"
    elif mercado == "Hipermais Joao Costa":
        market = "Hiper_joao_pl"
    elif mercado == "Hipermais Araquari":
        market = "Hiper_ara_pl"
    elif mercado == "Fort Atacadista":
        market = "Fort_pl"
    elif mercado == "Rodrigues":
        market = "Rodrigues_pl"
    listing()
        
    query = "SELECT CODE, PRODUCT, MEAN, VALUE FROM '"+market+"' order by CODE"
    vcon = data.ConnectDB()
    linhas = data.fill(vcon,query)
    for i in linhas:
        app.insert("","end",values=i)  
    



## Cria uma planilha para cada pedido realizando uma consulta no banco de dados de cada mercado
def create_spreadsheet():
    x = randint(1, 999999)
    vcon = data.ConnectDB()
    condor_query = 'SELECT CODE, PRODUCT, MEAN, VALUE FROM Condor_pl order by CODE'
    condor = data.fill(vcon,condor_query)
    condor = list(condor)

    bistek_query = 'SELECT CODE, PRODUCT, MEAN, VALUE FROM Bistek_pl order by CODE'
    bistek = data.fill(vcon,bistek_query)
    bistek = list(bistek)

    hipjc_query = 'SELECT CODE, PRODUCT, MEAN, VALUE FROM Hiper_joao_pl order by CODE'
    hipjc = data.fill(vcon,hipjc_query)
    hipjc = list(hipjc)

    hipara_query = 'SELECT CODE, PRODUCT, MEAN, VALUE FROM Hiper_ara_pl order by CODE'
    hipara = data.fill(vcon,hipara_query)
    hipara = list(hipara)

    fort_query = 'SELECT CODE, PRODUCT, MEAN, VALUE FROM Fort_pl order by CODE'
    fort = data.fill(vcon,fort_query)
    fort = list(fort)

    rod_query = 'SELECT CODE, PRODUCT, MEAN, VALUE FROM Rodrigues_pl order by CODE'
    rod = data.fill(vcon,rod_query)
    rod = list(rod)

    writer = pd.ExcelWriter(f'faturamento_{x}.xlsx', engine='xlsxwriter')

    df1 = pd.DataFrame(condor, columns=["Código", "Produto", "Média", "Qtd"])
    df1.to_excel(writer, sheet_name='Condor', index=False)

    df2 = pd.DataFrame(bistek, columns=["Código", "Produto", "Média", "Qtd"])
    df2.to_excel(writer, sheet_name='Bistek', index=False)

    df3 = pd.DataFrame(hipjc, columns=["Código", "Produto", "Média", "Qtd"])
    df3.to_excel(writer, sheet_name='Hipermais Joao Costa', index=False)

    df4 = pd.DataFrame(hipara, columns=["Código", "Produto", "Média", "Qtd"])
    df4.to_excel(writer, sheet_name='Hipermais Araquari', index=False)

    df5 = pd.DataFrame(fort, columns=["Código", "Produto", "Média", "Qtd"])
    df5.to_excel(writer, sheet_name='Fort Atacadista', index=False)

    df6 = pd.DataFrame(rod, columns=["Código", "Produto", "Média", "Qtd"])
    df6.to_excel(writer, sheet_name='Rodrigues', index=False)

    writer.save()
      


## Deleta todos os itens do banco de dados criados dentro da database selecionada
def clear():
    mercado = cmercados.get()
    if mercado == "Condor":
        market = "Condor_pl"
    elif mercado == "Bistek":
        market = "Bistek_pl"
    elif mercado == "Hipermais Joao Costa":
        market = "Hiper_joao_pl"
    elif mercado == "Hipermais Araquari":
        market = "Hiper_ara_pl"
    elif mercado == "Fort Atacadista":
        market = "Fort_pl"
    elif mercado == "Rodrigues":
        market = "Rodrigues_pl"

    vcon=data.ConnectDB()
    query="DELETE FROM '"+market+"'"
    data.clean(vcon, query)
    mercado_selected()

## Deleta item selecionado
def delete():
    try: 
        mercado = cmercados.get()
        if mercado == "Condor":
            market = "Condor_pl"
        elif mercado == "Bistek":
            market = "Bistek_pl"
        elif mercado == "Hipermais Joao Costa":
            market = "Hiper_joao_pl"
        elif mercado == "Hipermais Araquari":
            market = "Hiper_ara_pl"
        elif mercado == "Fort Atacadista":
            market = "Fort_pl"
        elif mercado == "Rodrigues":
            market = "Rodrigues_pl"
        itemSelection = app.selection()[0]
        valores = app.item(itemSelection, "values")
        code = valores[0]
        query = "DELETE FROM '"+market+"' WHERE CODE="+code
        vcon = data.ConnectDB()
        data.delete(vcon, query)
    except:
        messagebox.showinfo(title="ERRO", message="Selecione um serviço")
        return
    mercado_selected()



## Seleciona itens produtos cadastrado dentro dos bancos de produtos para adicionar ao banco de dados de planilhas
def adding():
    mercado = cmercados.get()
    if mercado == "Condor":
        market = "Condor_pl"
    elif mercado == "Bistek":
        market = "Bistek_pl"
    elif mercado == "Hipermais Joao Costa":
        market = "Hiper_joao_pl"
    elif mercado == "Hipermais Araquari":
        market = "Hiper_ara_pl"
    elif mercado == "Fort Atacadista":
        market = "Fort_pl"
    elif mercado == "Rodrigues":
        market = "Rodrigues_pl"
    id = combo.get()
    quantidade = entrada_quantidade.get()
    if id == '' or quantidade == '':
        messagebox.showinfo(title="ERRO", message="Preencha todas as informações")
        return
    vcon = data.ConnectDB()
    query = "SELECT * FROM '"+mercado+"' WHERE CODE="+id   ## Seleciona o nome do mercado dentro da combobox e puxa todos os códigos de produtos do mercado selecionado
    linhas = data.select(vcon, query)
    lista = list(linhas)
    codigo = lista[0]
    produto = lista[1]
    media = lista[2]
    condition = 'A'
    sql = "INSERT INTO '"+market+"' (CODE, PRODUCT, MEAN, VALUE) VALUES('"+str(codigo)+"','"+produto+"','"+str(media)+"', '"+str(quantidade)+"')" ## Insere no banco de dados de planilha de cada mercado as informações do produto
    data.insert(vcon, sql)
    mercado_selected()
    combo.delete(0, END)
    entrada_quantidade.delete(0, END)



## Janela principal
window = tk.Tk()
window.title('Gerador de Relatórios')
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


## Combobox dos mercados disponíveis 
cmercados = ttk.Combobox(window, values=mercados, width=15, justify='center')
cmercados.set("Condor") ## Seta um mercado como a primeira escolha da combobox
cmercados.place(x=100, y = 30)
mercado = cmercados.get()

## Botão de seleção de mercados
bmercados = Button(window, text='Selecionar', width=10, height=0, bg=selected, fg=color1, border=0, command=mercado_selected)
bmercados.place(x=260, y= 27)


## Label e Combobox dos produtos
lbcombo = Label(window, text='Produto', fg='black', bg=color2)
lbcombo.place(x=100, y=612)
value = StringVar()
combo = ttk.Combobox(window, textvariable=value, width=15, justify='center')
combolist = listing()
combo['values'] = (combolist) ## atualiza a lista de códigos de produtos
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


mercado_selected() ## Atualiza conexão com banco de dados
window.mainloop() ## Atualiza visualização da janela