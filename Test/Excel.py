'''
Cristian Sánchez
31/05/22

cris_s98@ciencias.unam.mx
'''

import os
import glob
import pandas as pd
import openpyxl
import mariadb

globalUser = 'cristiansanchez'
globalPassword = 'Alonso321'
globalHost = 'localhost'
glopalPort = 3306

def updateToXlsx(fileName):
    """Funcion para actualizar las extensiones de los
    archivos .xls, se crean archivos auxiliares con firma
    .xlsx para poder ser utilizado con la paqueteria openpyxl"""

    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_colwidth', None)

    newfileName = fileName.replace('.xls', '')

    excel = pd.read_excel(fileName)
    df = pd.DataFrame(excel)
    df.to_excel(newfileName + '.xlsx')


def deleteAuxXlsx():
    """Funcion que elimina los archivos auxiliares .xlsx"""

    xlsx_files = glob.glob('*.xlsx')
    for xlsx_file in xlsx_files:
        try:
            os.remove(xlsx_file)
        except OSError as e:
            print(f'Error:{ e.strerror}')


def initDataBase():
    """"Funcion para inicializar una base de datos.
    Por defecto trae las credenciales de origen"""

    try:
        conn = mariadb.connect(
            user=globalUser, password=globalPassword, host=globalHost, port=glopalPort)

    except mariadb.Error as e:
        print(f'Error al conectar con MariaDB: {e}')

    cur = conn.cursor()
    try:
        cur.execute('CREATE DATABASE JobTraveler')
    except mariadb.Error as e:
        print(f'Ya existe la base de datos')

    conn.close()


def createTable():
    """" Funcion para crear una tabla en la base de datos
    JobTraveler, por defecto trae las credenciales de origen"""

    try:
        conn = mariadb.connect(user=globalUser, password=globalPassword,
                               host=globalHost, port=glopalPort, database='JobTraveler')

    except mariadb.Error as e:
        print(f'Error al conectar con MariaDB: {e}')

    cur = conn.cursor()
    cur.execute('DROP TABLE IF EXISTS info')
    query = """CREATE TABLE info (meter_no INT NOT NULL AUTO_INCREMENT,
    serial_number varchar(100),
    panel_number varchar(100),
    job_number varchar(100),
    job_name varchar(100),
    is_seal int(36),
    type char(36),
    modbus_id int(36),
    PRIMARY KEY(meter_no))"""
    cur.execute(query)
    conn.commit()
    conn.close()


def insertToTable(nameFile):
    """ Funcion para insertar los valores a la tabla info"""

    info = getCellInfo(nameFile)
    serialNumber = getSerialNumber(nameFile)
    serialNumberID = serialNumber[0]
    tamano = len(serialNumber)
    info.insert(0, -1)

    try:
        conn = mariadb.connect(user=globalUser, password=globalPassword,
                               host=globalHost, port=glopalPort, database='JobTraveler')

    except mariadb.Error as e:
        print(f'Error connecting to MariaDB Platform: {e}')

    cur = conn.cursor(len(serialNumber))
    infoTable = ', '.join('?' * len(info))

    for i in range(len(serialNumber)):
        info[0] = serialNumber[i]
        query = ("INSERT INTO info (serial_number, panel_number, job_number, job_name, is_seal, type, modbus_id) VALUES (%s);" % infoTable)
        cur.execute(query, info)

    conn.commit()
    conn.close()


def getCellInfo(nameFile):
    """Funcion para obtener los valores especificos de las celdas especificas.
    Se devuelve una lista con todos los datos obtenidos"""

    excel = openpyxl.load_workbook(nameFile, read_only=False, keep_vba=True)
    sheet = excel.active

    # Guardando Valores
    panelNumber = sheet['E3'].value  # str
    jobNumber = sheet['E4'].value  # str
    jobName = sheet['E5'].value  # str
    seal = int('X' in str(sheet['K3'].value))  # bool
    typeEx = sheet['C28'].value  # str
    modbusID = sheet['D33'].value  # int

    info = [panelNumber, jobNumber, jobName, seal,
            typeEx, modbusID]

    return info


def getSerialNumber(nameFile):
    """ Funcion para obtener los numeros de serie.
    Se devuelve una lista con todos los numeros de serie."""
    excel = openpyxl.load_workbook(nameFile, read_only=False, keep_vba=True)
    sheet = excel.active

    serialNumberNonNormalized = []
    for i in range(50, 74):

        serialNumberNonNormalized.append(sheet['D' + str(i)].value)

    serialNumber = []
    for i in serialNumberNonNormalized:
        if i != None:
            serialNumber.append(i)

    return serialNumber


if __name__ == '__main__':

    archivos = ['-3 2DPEA.xls', '-5 6DPEA.xls', '-8 2PP3BT.xls']
    for i in range(len(archivos)):
        updateToXlsx(archivos[i])

    initDataBase()
    createTable()

    archivos = ['-3 2DPEA.xlsx', '-5 6DPEA.xlsx', '-8 2PP3BT.xlsx']
    for i in range(len(archivos)):
        insertToTable(archivos[i])

    deleteAuxXlsx()

