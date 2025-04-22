import pandas as pd

# ABRIR ARQUIVO EXCEL (read_excel)
df = pd.read_excel('dados_sujos.xlsx')

#-------------------------------------------------------------------------------------------------------------------
#TRANSFORMAÇÃO GERAL DE DADOS
#--------------------------------------------------------------------------------------------------------------------

# Deletar linhas e colunas completamente vazias
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
print(df)

#-----------------------------------------------------------------------------------------------------------------
#TRANSFORMAÇÃO ESPECÍFICA DE DADOS (SUBSTITUIR)
#-----------------------------------------------------------------------------------------------------------------

# Deletar linhas com dados essenciais ausentes
df = df.dropna(subset=['NOME', 'SALÁRIO'])


# Converter tipo para data (to_datetime)
df['DATA ADMISSÃO']= pd.to_datetime(df['DATA ADMISSÃO'], errors='coerce', dayfirst=True)

print(df)


# -------------------------------------------------------------------------------

# SALVAR DADOS ALTERADOS (to_excel) - DAR NOVO NOME PARA ARQUIVO
df.to_excel('Vini.xlsx')