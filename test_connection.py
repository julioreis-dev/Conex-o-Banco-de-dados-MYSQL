import mysql.connector


def testar_bd(*args):
    index = 0
    try:
        connection = mysql.connector.connect(host=args[0], database=args[1], user=args[2],
                                             password=args[3], port=args[4])
    except mysql.connector.Error as error:
        if error.errno == 1049:
            index = 1
        elif error.errno == 1045:
            index = 2
        else:
            print(error)
    else:
        connection.close()
    resultados = ['ok', 'Banco de dados inexistente.', 'Nome do usuário ou password incorreto.']
    return resultados[index]
