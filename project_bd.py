import mysql.connector
from environs import Env


class CreateBD:
    def __init__(self, host_client, database_client, user_client, password_client, port_client):
        self.host = host_client
        self.database = database_client
        self.user = user_client
        self.password = password_client
        self.port = port_client

    def create_database(self, database_name):
        try:
            connection = mysql.connector.connect(host=self.host, user=self.user, password=self.password)

            mycursor = connection.cursor()
            mycursor.execute(f"CREATE DATABASE {database_name}")
        except mysql.connector.errors.DatabaseError:
            return f'Banco de dados "{database_name}" já existente e em uso!!!'
        else:
            return f'Banco de dados "{database_name}" criado com sucesso!!!'

    def create_tables(self, list_instr):
        connection = mysql.connector.connect(host=self.host, database=self.database, user=self.user,
                                             password=self.password, port=self.port)
        for query in list_instr:
            cursor = connection.cursor()
            sql = query
            cursor.execute(sql)
        connection.close()


def main():
    sql1 = """CREATE TABLE IF NOT EXISTS clientes (
               idclientes INT NOT NULL AUTO_INCREMENT,
               cpf VARCHAR(11) NOT NULL,
               nome_completo VARCHAR(100) NOT NULL,
               nascimento VARCHAR(10) NOT NULL,
               email VARCHAR(100) NOT NULL,
               telefone VARCHAR(11) NOT NULL,
               salario FLOAT NOT NULL,
               proposta_enviada DATETIME NULL,
               idproposta INT NULL,
               status_proposta VARCHAR(30) NULL,
               PRIMARY KEY (idclientes))"""

    sql2 = """CREATE TABLE IF NOT EXISTS list_ofertas (
                    idofertas INT NOT NULL AUTO_INCREMENT,
                    ofertas JSON NULL,
                    inclusao DATETIME NULL,
                    PRIMARY KEY (idofertas))"""

    list_sql = (sql1, sql2)
    env = Env()
    env.read_env()
    host = env.str('NT_BD_HOST')
    database = env.str('NT_BD_DATABASE')
    user = env.str('NT_BD_USER')
    password = env.str('NT_BD_PASSWORD')
    port = env.int('NT_BD_PORT')
    db = CreateBD(host, database, user, password, port)
    return_db = db.create_database(database)
    db.create_tables(list_sql)
    print(return_db)
