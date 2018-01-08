# encoding=iso-8859-1
import sys
import win32com.client
import pyodbc
import os
from PyQt4 import QtGui, uic
import requests

reload(sys)
sys.setdefaultencoding('iso-8859-1')

qtCreatorFile = "talanqueraUi.ui"  # Nombre del archivo UI '.ui'

Ui_MainWindow, QtBaseClass = uic.loadUiType(qtCreatorFile)


class CallServer:
    # Class para requests del sistema.
    __Host = ''
    __Responce = None

    def __init__(self, host=''):
        self.__Host = host

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        pass

    def LogIn(self, userName='', userPass=''):
        # Funcion que valida usuario y contraseña.
        payload = {
                'cmd': 'logIn',
                'usr': userName,
                'pwd': userPass
            }

        self.__Responce = requests.post(self.__Host, params=payload)

        jsonOBJ = self.__Responce.json()

        return jsonOBJ

    def getEndDates(self, userName=''):
        # Funcion que devuelve el resultSet del sql del servlet.
        payload = {
                'cmd': 'DeadEnd',
                'usr': userName
            }

        self.__Responce = requests.post(self.__Host, params=payload)

        jsonOBJ = self.__Responce.json()

        return jsonOBJ

# Url del servlet al cual se hacen los requests.
WSURL = "https://diceros.ls-sys.com/Sistema/talanquera"


class TalanqueraUi(QtGui.QMainWindow, Ui_MainWindow):
    # Class principal. (Form)
    def __init__(self):
        # Con Frame:
        QtGui.QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)

        modoglobal = 'odbc'     # Representa la funcion que se usara para interactuar con la DB Access.
        condoaddress = "diceros.ls-sys.com"  # Se usa para verificar conexion a internet.

        # Setup del frame. Se fija el tamaño de la pantalla para deshabilitar el boton de maximizar.
        QtGui.QMainWindow.setFixedWidth(self, 684)
        QtGui.QMainWindow.setFixedHeight(self, 300)

        # setParent(None) para remover elementos que se crearon pero ya no se usaran.
        self.gtxResult.setParent(None)
        self.lblOracleDB.setParent(None)
        self.txtODB.setParent(None)
        self.pbTestODB.setParent(None)
        self.massdisable(1)

        if self.testodb(condoaddress):
            self.massdisable(1)
        else:
            self.massdisable()

        self.pbActualizar.setDisabled(True)
        self.txtADB.setDisabled(True)

        self.pbLogin.clicked.connect(lambda: self.login(str(self.txtUsuario.text()), str(self.txtPwd.text())))
        self.pbTestADB.clicked.connect(lambda: self.inter_access(modoglobal))
        self.pbActualizar.clicked.connect(lambda: self.inter_access('actualizar', modoglobal))

    def massdisable(self, paso=0):
        stat = True
        if paso == 0:
            self.pbTestODB.setDisabled(stat)
            self.pbTestADB.setDisabled(stat)
            self.pbActualizar.setDisabled(stat)
            self.txtADB.setDisabled(stat)
            self.txtODB.setDisabled(stat)
            self.lblOracleDB.setDisabled(stat)
            self.lblAccessDB.setDisabled(stat)
            self.txtUsuario.setDisabled(stat)
            self.txtPwd.setDisabled(stat)
            self.pbLogin.setDisabled(stat)
        elif paso == 1:
            self.pbTestADB.setDisabled(stat)
            self.pbActualizar.setDisabled(stat)
            self.txtADB.setDisabled(stat)
            self.lblAccessDB.setDisabled(stat)
        elif paso == 2:
            self.pbTestADB.setDisabled(not stat)
            self.pbActualizar.setDisabled(stat)
            self.lblAccessDB.setDisabled(stat)
            self.txtUsuario.setDisabled(stat)
            self.txtPwd.setDisabled(stat)
            self.pbLogin.setDisabled(stat)

    def inter_access(self, destino, mod=None):
        dato = self.txtADB.text()
        path = dato
        self.txtADB.setText(path)
        if destino == 'odbc':
            self.odbc(path)
        elif destino == 'ado':
            self.ado(path)
        elif destino == 'actualizar':
            self.actualizar(path, mod)
        else:
            pass

    def alert(self, title, text):
        QtGui.QMessageBox.about(self, title, text)

    def ado(self, dbpath):
        try:
            """
            connect with com dispatch objs
            """
            db = dbpath
            conn = win32com.client.Dispatch(r'ADODB.Connection')
            dsn = ('PROVIDER = Microsoft.Jet.OLEDB.4.0;DATA SOURCE = ' + db + ';')
            conn.Open(dsn)
            self.alert("CONNECTION SUCCESSFUL!", "¡Conexión exitosa!")
            conn.Close()

            self.pbActualizar.setDisabled(False)
            self.pbTestADB.setDisabled(True)

            return True
        except Exception:
            self.alert("CONNECTION ERROR!",
                       "¡Se ha producido un error de conexión!" +
                       "\nRevise que el nombre de la base de datos Access sea correcta.")
            return False

    def testodb(self, direct):
        direccion = str(direct)
        resp = (os.system("ping -n 1 " + direccion))
        if resp == 0:
            self.alert("Success!", "¡Conexión estable a internet!")
            return True
        else:
            self.alert("Failed!", "¡Revise su conexión a internet!")
            return False

    def odbc(self, dbruote):
        try:
            """
            connects with odbc
            """
            db = dbruote
            constr = ('Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=' + db)
            conn = pyodbc.connect(str(constr), autocommit=True)
            self.alert("Success!", "¡Conexión exitosa!")
            conn.close()

            self.pbActualizar.setDisabled(False)
            self.pbTestADB.setDisabled(True)
            self.txtADB.setDisabled(True)

            return True
        except Exception:
            self.alert("Failed!",
                       "¡Se ha producido un error de conexión!" +
                       "\nRevise que el nombre de la base de datos Access sea correcta.")
            return False

    def login(self, usr, pwd):

        with CallServer(WSURL) as ws:
            response = ws.LogIn(usr.upper(), pwd)
            
            if response['ACK'] == '1':
                self.alert("Control Condominio", "¡Usuario y Contraseña VÁLIDOS!")
                self.massdisable(2)
                self.txtADB.setText(response['PDB'])
            else:
                self.alert("Control Condominio", "¡Usuario y Contraseña NO VÁLIDOS!")

    def actualizar(self, dbruote, modo):
        try:

            SQLsUpdate = []

            with CallServer(WSURL) as ws:
                strUsr = str(self.txtUsuario.text())
                strUsr = strUsr.upper()
                response = ws.getEndDates(strUsr)
                counterI = 0
                in_clause = ""

                for codTarjeta, bloque in response.iteritems():
                    tipo = bloque[1]
                    if tipo == '2':
                        sql = "UPDATE {0} SET {1} = Format('{2}', 'yyyy-mm-dd') " \
                              "where Mid( Mid([EmployeeCode], InStr(1, [EmployeeCode], '-')+1)," \
                              "InStr(1, Mid([EmployeeCode], InStr(1, [EmployeeCode], '-')+1) , '-')+1) = '{4}'" \
                              .format("TEmployee", "EndDate", bloque[0][0:10], "EmployeeCode", codTarjeta)
                        SQLsUpdate.append(sql)
                        in_clause += "'" + codTarjeta + "', "
                    elif tipo == '1':
                        sql = "Insert into {0}" \
                              "(EmployeeID, EmployeeCode, EmployeeName, CardNo, pin, EmpEnable, Birthday, " \
                              "RegDate, EndDate, ACCESSID, Deleted, Leave, Password)" \
                              "values('{1}', '{1}', '{2}', '{3}', {4}, {5}, {6}, {7}, " \
                              "Format('{8}', 'yyyy-mm-dd'), {9}, {10}, {11}, '{12}')" \
                              .format("TEmployee", bloque[2], bloque[3].upper(), codTarjeta, 1234, 0, "Date()",
                                      "Date()", bloque[0][0:10], 0, 0, 0, 1234)
                        SQLsUpdate.append(sql)
                        counterI += 1
            if modo == 'odbc':
                """
                connects with odbc
                """
                db = dbruote
                constr = ('Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=' + db)
                conn = pyodbc.connect(str(constr), autocommit=True)
                cur = conn.cursor()

                fin = len(in_clause)
                fin -= 2
                in_clause = in_clause[0:fin]

                verifier = "select enddate from TEmployee " \
                           "where Mid( Mid([EmployeeCode], InStr(1, [EmployeeCode], '-')+1)," \
                           "InStr(1, Mid([EmployeeCode], InStr(1, [EmployeeCode], '-')+1) , '-')+1) in ({0})" \
                           .format(in_clause)

                for sentencia in SQLsUpdate:
                    cur.execute(sentencia)

                cur.execute(verifier)
                t = tuple(cur)
                counterU = len(t)

                conn.close()

                if counterI > 0:
                    self.alert("Success!", "¡Tarjetas Actualizadas! \n {0} Actualizaciones.".format(counterU) +
                               " \n {0} Ingresos Nuevos.".format(counterI))
                else:
                    self.alert("Success!", "¡Tarjetas Actualizadas! \n {0} Actualizaciones.".format(counterU))
                sys.exit(0)
            elif modo == 'ado':
                """
                connect with com dispatch objs
                """
                db = dbruote
                conn = win32com.client.Dispatch(r'ADODB.Connection')
                dsn = ('PROVIDER = Microsoft.Jet.OLEDB.4.0;DATA SOURCE = ' + db + ';')
                conn.Open(dsn)

                rs = win32com.client.Dispatch(r'ADODB.Recordset')

                fin = len(in_clause)
                fin -= 2
                in_clause = in_clause[0:fin]

                verifier = "select enddate from TEmployee " \
                           "where Mid( Mid([EmployeeCode], InStr(1, [EmployeeCode], '-')+1)," \
                           "InStr(1, Mid([EmployeeCode], InStr(1, [EmployeeCode], '-')+1) , '-')+1) in ({0})" \
                           .format(in_clause)

                for sentencia in SQLsUpdate:
                    rs.Open(sentencia, conn, 1, 3)

                t = rs.Open(verifier, conn, 1, 3)
                counterU = len(t)

                conn.Close()

                if counterI > 0:
                    self.alert("Success!", "¡Tarjetas Actualizadas! \n {0} Actualizaciones.".format(counterU) +
                               " \n {0} Ingresos Nuevos.".format(counterI))
                else:
                    self.alert("Success!", "¡Tarjetas Actualizadas! \n {0} Actualizaciones.".format(counterU))
                sys.exit(0)
            else:
                pass
        except Exception:
            self.alert("Failed!", "¡Se ha producido un error!")


if __name__ == "__main__":
    app = QtGui.QApplication(sys.argv)
    window = TalanqueraUi()
    window.show()
    sys.exit(app.exec_())
