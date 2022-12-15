import pyautogui as pg
import time
import os
import pandas as pd

def options() :
    opcoes = """ 
    Opcoes de pesquisa:

    0 - Exit
    1 - Soma
    2 - Show
    3 - Numero de Itens
    4 - Quantidade

    Digite a opcao: 
    """
    return opcoes

def planilha() :
    name_ = input("Arquivo: ")
    return name_

def tools() :
    numero = int(input("Numero de arquivos: "))
    for i in range(numero) :
        i =+ 1
        print(i)
        nome = planilha() 
        lista = pd.read_excel(f'{nome}.xlsx')
        pesquisa = input("Search? ")
        if pesquisa == "y" :
            print(lista)
            menu = str(options())
            search = input("Pesquisa: ")
            resultado = lista[[f'{search}']]
            print(resultado)
            while True :
                type_search = input(menu)
                if int(type_search) == 1 :
                    a = resultado.sum()
                    print("'''''''''''''''''''")
                    print(a)
                    print("'''''''''''''''''''")

                if int(type_search) == 2 :
                    a = resultado.describe()
                    print("'''''''''''''''''''")
                    print(a)
                    print("'''''''''''''''''''")

                if int(type_search) == 3 :
                    print("'''''''''''''''''''")
                    a = resultado.count()
                    print("'''''''''''''''''''")
                    print(a)

                if int(type_search) == 4 :
                    a = resultado.value_counts()
                    print("'''''''''''''''''''")
                    print(a)
                    print("'''''''''''''''''''")

                if int(type_search) == 0 :
                    break
            
        else :
            print(lista)
        
        #print(lista)
        #mensagens = lista[['contato', 'mensagem']]
tools()