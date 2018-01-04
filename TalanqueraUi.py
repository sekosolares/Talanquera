# encoding=iso-8859-1
import sys
import win32com.client
import pyodbc
import os
from PyQt4 import QtGui, uic
import requests

reload(sys)
sys.setdefaultencoding('iso-8859-1')

qtCreatorFile = "talanqueraUi.ui"  # Enter file here. extension '.ui'

Ui_MainWindow, QtBaseClass = uic.loadUiType(qtCreatorFile)


class CallServer:

    __Host = ''
    __Responce = None

    def __init__(self, host=''):
        self.__Host = host

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        pass

    def LogIn(self, userName='', userPass=''):

        payload = {
                'cmd': 'logIn',
                'usr': userName,
                'pwd': userPass
            }

        self.__Responce = requests.post(self.__Host, params=payload)

        jsonOBJ = self.__Responce.json()

        return jsonOBJ

    def getEndDates(self, userName=''):

        payload = {
                'cmd': 'DeadEnd',
                'usr': userName
            }

        self.__Responce = requests.post(self.__Host, params=payload)

        jsonOBJ = self.__Responce.json()

        return jsonOBJ


WSURL = "https://diceros.ls-sys.com/Sistema/talanquera"

# Clase principal. (Form)
class TalanqueraUi(QtGui.QMainWindow, Ui_MainWindow):
    def __init__(self):
        # Con Frame:
        QtGui.QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)

        modoglobal = 'odbc'
        condoaddress = "diceros.ls-sys.com"

        QtGui.QMainWindow.setFixedWidth(self, 684)
        QtGui.QMainWindow.setFixedHeight(self, 300)
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
                counterU = 0
                counterI = 0

                '''
                for codTarjeta, endDate in response.iteritems():
                    sql = "Update {0} set {1} = Format('{2}', 'yyyy-mm-dd') where Trim({3}) = Trim('{4}')"\
                        .format("TEmployee", "EndDate", endDate[0:10], "CardNo", codTarjeta)
                    SQLsUpdate.append(sql)
                '''
                for codTarjeta, bloque in response.iteritems():
                    if bloque[1] == '2':
                        sql = "Update {0} set {1} = Format('{2}', 'yyyy-mm-dd') where Trim({3}) = Trim('{4}')"\
                            .format("TEmployee", "EndDate", bloque[0][0:10], "CardNo", codTarjeta)
                        SQLsUpdate.append(sql)
                        counterU += 1
                    elif bloque[1] == '1':
                        sql = "Insert into {0}" \
                              "(EmployeeID, EmployeeCode, EmployeeName, CardNo, pin, EmpEnable, Sex, Birthday, " \
                              "RegDate, EndDate, ACCESSID, Deleted, Leave, Password)" \
                            .format("TEmployee", "EndDate", bloque[0][0:10], "CardNo", codTarjeta)
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

                for sentencia in SQLsUpdate:
                    cur.execute(sentencia)

                conn.close()

                self.alert("Success!", "¡Tarjetas Actualizadas!")
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

                for sentencia in SQLsUpdate:
                    rs.Open(sentencia, conn, 1, 3)

                conn.Close()

                self.alert("Success!", "¡Tarjetas Actualizadas!")
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
