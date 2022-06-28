import pandas as pd
import psycopg2

try:
    con = psycopg2.connect(host='host', database='banco', user='usuario', password='senha')
except:
    print('Sem conexão online')
    exit(1)

#arquivo = pd.read_excel(r"Biblioteca.xlsx")
arquivo = pd.read_excel(r"Componentes.xlsx")


dimensions = arquivo.shape
print(dimensions)
linhas = dimensions[0]
colunas = dimensions[1]

print(linhas)
print(colunas)
print('------')
for i in range(0, linhas, 1):
    print(i)
    # instruções para incluir o arquivo biblioteca.xlsx
    #titulo = arquivo['Titulo'][i]
    #print('Titulo: ', arquivo['Titulo'][i])
    #autor = arquivo['Autor'][i]
    #print('Autor: ', arquivo['Autor'][i])
    #qtde = arquivo['qtde_exe'][i]
    #print('Qtde: ', round(arquivo['qtde_exe'][i], 3))
    #pagina = arquivo['Pagina'][i]
    #print('Pagina: ', arquivo['Pagina'][i])
    #print('------')

    # instruções para incluir o arquivo componentes.xlsx
    #local = arquivo['local'][i]
    categoria = arquivo['categoria'][i]
    print('categoria: ', arquivo['categoria'][i])
    foto = arquivo['categoria'][i]
    foto += ".jpg"
    nome = arquivo['Nome'][i]
    print('Nome: ', arquivo['Nome'][i])
    qtde = arquivo['Qtde'][i]
    print('Qtde: ', round(arquivo['Qtde'][i], 3))
    info = arquivo['Info'][i]
    print('Info: ', arquivo['Info'][i])
    print('------')

    cursor = con.cursor()
    # inclui o arquivo biblioteca.xlsx
    # cursor.execute("INSERT INTO tb_livros (titulo, autor, qtde, pagina) VALUES(%s, %s, %s, %s)", (titulo, autor, qtde, pagina,))

    # inclui o arquivo componentes.xlsx
    #cursor.execute("INSERT INTO tb_almoxarife (local, categoria, foto, nome, qtde, info) VALUES(%s, %s, %s, %s, %s, %s)", (local, categoria, foto, nome, int(qtde), info,))
    cursor.execute(
        "INSERT INTO tb_almoxarife (categoria, foto, nome, qtde, info) VALUES(%s, %s, %s, %s, %s)",
        (categoria, foto, nome, int(qtde), info,))

    con.commit()