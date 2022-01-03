import sqlite3
from sqlite3 import Error


def ConnectDB(): ## Conexão com o banco
    way = 'products.db'
    con=None
    try:
        con=sqlite3.connect(way)

    except Error as ex:
        print(ex)
    return con

vcon = ConnectDB()


def insert(connection, sql):
    try:
        c= connection.cursor()
        c.execute(sql)
        connection.commit()
        print("Registro Inserido com SUCESSO")
    except Error as ex:
        print(ex)


def fill(connection, sql):
    try:
        c= connection.cursor()
        c.execute(sql)
        connection.commit()
        res=c.fetchall()
        print('Registros Atualizados com SUCESSO')
    except Error as ex:
        print(ex)
    return res

def delete(connection, sql):
    try:
        c= connection.cursor()
        c.execute(sql)
        connection.commit()
        print('Registro Removido com SUCESSO')
    except Error as ex:
        print(ex)

def select(connection, sql):
    try:
        c= connection.cursor()
        c.execute(sql)
        connection.commit()
        res=c.fetchone()
        print('Item selecionado do Banco de Dados')
    except Error as ex:
        print(ex)
    return res
    
def clean(connection, sql):
    try:
        c= connection.cursor()
        c.execute(sql)
        connection.commit()
        res=c.fetchall()
        print('Todos itens deletados com SUCESSO')
    except Error as ex:
        print(ex)
    return res