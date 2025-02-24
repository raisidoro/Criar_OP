import pyodbc

def dbConn():
    server = '181.41.175.39,37000'
    database = 'CF260B_132575_PR_PD'
    username = 'CLT132575HRKOAC_READ'
    password = 'qkwgb18064UZBRW@!'
    strConn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};Server=' +
                            server+';DATABASE='+database+';UID='+username+';PWD='+password)
    cursor = strConn.cursor()

    return cursor






dbConn()
#Sample select query