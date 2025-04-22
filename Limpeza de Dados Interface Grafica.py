from tkinter import StringVar

import pandas as pd
import tkinter as tk
from tkinter import messagebox


#-------------------------------------------------------------------------------------------------------------------
#TRANSFORMAÇÃO GERAL DE DADOS
#--------------------------------------------------------------------------------------------------------------------

def transformar_dados():

#validação
    if not y.get().endswith('.xlsx'):
        messagebox.showerror('Atenção!!!', 'Atenção, inserir (.xlsx) no final do nome do arquivo')


    if not z.get().endswith('.xlsx'):
            messagebox.showerror('Atenção!!!', 'Atenção, inserir (.xlsx) no final do nome do arquivo')


# ABRIR ARQUIVO EXCEL (read_excel)
    df = pd.read_excel(y.get())


#Deletar linhas e colunas completamente vazias
    df = df.dropna(how='all')
    df = df.dropna(how='all', axis=1)

# Remover espaços das colunas (apenas o nome das colunas) (str.strip)
    df.columns = df.columns.str.strip()

# Remover espaços das células (apenas células de str)
    for col in df.columns:
        df[col]=df[col].apply(lambda x: x.strip()
                            if isinstance (x, str)
                            else x)

# Células de strings transformadas para str.title(maiuscula+minuscula)
    for col in df.columns:
        df[col] = df[col].apply(lambda x: x.title()
                                if isinstance (x, str)
                                else x)

# Substituindo caracteres por espaços (replace)
    for col in df.columns:
        df[col]=df[col].replace('[R$,]','', regex=True)


# Converter para valor numérico (to_numeric)
    for col in df.columns:
        df[col]=df[col].apply(lambda x: float(x)
                              if isinstance (x, (int,float))
                                else x)

# Nome das colunas em maiusculo (str.upper)
    df.columns=df.columns.str.upper()

# Preencher valores nulos (fillna)
    df.fillna(0)

# Remover Duplicatas (drop_duplicates)
    df.drop_duplicates(inplace=True)

# Visualizar os dados
    #print(df)

#-----------------------------------------------------------------------------------------------------------------
#TRANSFORMAÇÃO ESPECÍFICA DE DADOS (SUBSTITUIR)
#-----------------------------------------------------------------------------------------------------------------

# Deletar linhas com dados essenciais ausentes
    df = df.dropna(subset=['NOME', 'SALÁRIO'])


# Converter tipo para data (to_datetime)
    df['DATA ADMISSÃO']= pd.to_datetime(df['DATA ADMISSÃO'], errors='coerce', dayfirst=True)

#print(df)

# -------------------------------------------------------------------------------

# SALVAR DADOS ALTERADOS (to_excel) - DAR NOVO NOME PARA ARQUIVO
    df.to_excel(z.get())


# Mensagem de que deu tudo certo!
    messagebox.showinfo('Transfomação feita com sucesso', 'Arquivo transformado com sucesso!')


#-------------------------------------------------------------------------------------

#CRIANDO INTERFACE GRÁFICA
janela = tk.Tk()
janela.title('Limpeza de Dados Excel')
janela.geometry("400x400")

texto_fixo = tk.Label(janela, text='INSIRA OS DADOS PARA LIMPAR SUA TABELA', font=('Ariel', 10))
texto_fixo.grid(column=0, row=0, padx=40, pady=10)


y= StringVar()
z= StringVar()

texto1 = tk.Label(janela, text='Arquivo Excel Sujo: ')
texto1.grid(column=0, row=3, padx=1, pady=10)
campo = tk.Entry(janela, textvariable= y)
campo.grid(column=0, row=4, padx=1, pady=20)

texto2 = tk.Label(janela, text='Nome do Novo Arquivo Limpo: ')
texto2.grid(column=0, row=9, padx=1, pady=10)
campo2 = tk.Entry(janela, textvariable= z)
campo2.grid(column=0, row=10, padx=1, pady=20)

botao = tk.Button(janela, text='Limpar Agora', command= transformar_dados)
botao.grid(column=0, row=12, padx=1, pady=20)


janela.mainloop()
