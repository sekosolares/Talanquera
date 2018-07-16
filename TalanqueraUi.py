# encoding=iso-8859-1
import sys
import win32com.client
import pyodbc
import os
import time
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
        log.write("# [func.LogIn]:Llamado desde la función de Login de la forma. Valida usuario y contraseña."
                  " Parametros: Usuario={0} ; Contraseña={1}\n".format(userName, userPass))

        self.__Responce = requests.post(self.__Host, params=payload)
        log.write("# [func.LogIn]:Se ha realizado un post al server. Respuesta del server:"
                  " {0}\n".format(self.__Responce))

        jsonOBJ = self.__Responce.json()

        return jsonOBJ

    def getEndDates(self, userName=''):
        # Funcion que devuelve el resultSet del sql del servlet.
        payload = {
                'cmd': 'DeadEnd',
                'usr': userName
            }
        log.write("# [func.getEndDates]:Llamada desde la función actualizar. Contacto al server con cmd=DeadEnd "
                  "y usr={0}\n".format(userName))
        self.__Responce = requests.post(self.__Host, params=payload)
        log.write("# [func.getEndDates]:Se ha realizado un post al server. "
                  "Respuesta del server: {0}\n".format(self.__Responce))
        jsonOBJ = self.__Responce.json()

        return jsonOBJ


# Url del servlet al cual se hacen los requests.
WSURL = "https://diceros.ls-sys.com/Sistema/talanquera"

nombre = 'log-'  # Nombre del archivo que lleva los logs del programa.
fecha = time.strftime("%d-%m-%Y_%H-%M-%S")  # Fecha y hora para concatenar con el log.
nombre += fecha
texto_info = ''

# Abriendo el archivo log para escribir en el.
log = open(nombre+'.txt', "w")


class TalanqueraUi(QtGui.QMainWindow, Ui_MainWindow):
    # Class principal. (Form)
    def __init__(self):
        # Con Frame:
        QtGui.QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)

        modoglobal = 'ado'     # Representa la funcion que se usara para interactuar con la DB Access.
        condoaddress = "diceros.ls-sys.com"  # Se usa para verificar conexion a internet.

        # Setup del frame. Se fija el tamaño de la pantalla para deshabilitar el boton de maximizar.
        QtGui.QMainWindow.setFixedWidth(self, 684)  # Ancho del form.
        # QtGui.QMainWindow.setFixedHeight(self, 300)  # Alto del form.
        QtGui.QMainWindow.setFixedHeight(self, 490)  # Alto del form.

        log.write("##### Inicio del log #####\n"
                  "# [class.TalanqueraUi]:Se ha setteado correctamente el tama{0}o de la ventana.\n".format(chr(164)))

        # setParent(None) para remover elementos que se crearon pero ya no se usaran.
        # self.gtxResult.setParent(None)
        self.lblOracleDB.setParent(None)
        self.txtODB.setParent(None)
        self.pbTestODB.setParent(None)
        self.gtxResult.setDisabled(True)
        log.write("# [class.TalanqueraUi]:Se deshabilitan los campos, a excepción de los campos de inicio de sesión.\n")
        self.massdisable(1)

        log.write("# [class.TalanqueraUi]:A punto de probar conexión a internet.\n")
        if self.testodb(condoaddress):
            log.write("# [class.TalanqueraUi]:Conexión exitosa a internet. Procede...\n")
            self.massdisable(1)
        else:
            log.write("# [class.TalanqueraUi]:Fallo en la conexión a internet. Deshabilitar la forma.\n")
            self.massdisable()

        self.pbActualizar.setDisabled(True)
        self.txtADB.setDisabled(True)

        self.pbLogin.clicked.connect(lambda: self.login(str(self.txtUsuario.text()), str(self.txtPwd.text())))
        self.pbTestADB.clicked.connect(lambda: self.inter_access(modoglobal))
        self.pbActualizar.clicked.connect(lambda: self.inter_access('actualizar', modoglobal))

    def massdisable(self, paso=0):
        log.write("# [func.massdisable]:Llamada a función correcta. Params: paso={0}\n".format(paso))
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
            log.write("# [func.massdisable]:Todo deshabilitado. No existió conexión a internet.\n")
        elif paso == 1:
            self.pbTestADB.setDisabled(stat)
            self.pbActualizar.setDisabled(stat)
            # self.txtADB.setDisabled(stat)
            self.lblAccessDB.setDisabled(stat)
            log.write("# [func.massdisable]:La única sección habilitada es la de inicio de sesión.\n")
        elif paso == 2:
            self.pbTestADB.setDisabled(not stat)
            self.pbActualizar.setDisabled(stat)
            self.lblAccessDB.setDisabled(stat)
            self.txtUsuario.setDisabled(stat)
            self.txtPwd.setDisabled(stat)
            self.pbLogin.setDisabled(stat)
            self.txtADB.setDisabled(not stat)
            log.write("# [func.massdisable]:Inicio de sesión exitoso. "
                      "Se habilita el botón para probar conexión a base de datos Access.\n")

    def inter_access(self, destino, mod=None):
        log.write("# [func.inter_access]:Llamada a función para interacción con Access correcta."
                  "Params: destino={0} ; mod={1}\n".format(destino, mod))
        dato = self.txtADB.text()
        path = str(dato)
        # self.txtADB.setText(path)
        if destino == 'odbc':
            log.write("# [func.inter_access]:Se usa el modo '{0}' para conectar a Access.\n".format(destino))
            self.odbc(path)
        elif destino == 'ado':
            log.write("# [func.inter_access]:Se usa el modo '{0}' para conectar a Access.\n".format(destino))
            self.ado(path)
        elif destino == 'actualizar':
            log.write("# [func.inter_access]:Se actualizará la DB Access por medio de '{0}'.\n".format(mod))
            self.actualizar(path, mod)
        else:
            log.write("# [func.inter_access]:Los parámetros no fueron correctos!\n")
            pass

    def alert(self, title, text):
        QtGui.QMessageBox.about(self, title, text)

    def ado(self, dbpath):
        log.write("# [func.ado]:Llamada a función correcta. Params: bdpath={0}\n".format(dbpath))
        try:
            """
            connect with com dispatch objs
            """
            db = dbpath
            conn = win32com.client.Dispatch(r'ADODB.Connection')
            log.write("# [func.ado]:Conexión por iniciar... Variable conn={0}\n".format(conn))
            dsn = ('PROVIDER = Microsoft.Jet.OLEDB.4.0;DATA SOURCE = ' + db + ';')
            log.write("# [func.ado]:String que se abrirá para la conexión: dsn={0}\n".format(dsn))
            conn.Open(dsn)
            self.alert("CONNECTION SUCCESSFUL!", "¡Conexión exitosa!")
            log.write("# [func.ado]:Conexión a Access realizada con éxito.\n")
            conn.Close()
            log.write("# [func.ado]:Conexión a Access cerrada...\n")

            self.pbActualizar.setDisabled(False)
            self.pbTestADB.setDisabled(True)
            self.txtADB.setDisabled(True)

            return True
        except Exception as ex:
            self.alert("CONNECTION ERROR!",
                       "¡Se ha producido un error de conexión!" +
                       "\nRevise que el nombre de la base de datos Access sea correcta.")
            log.write("# [func.ado]:Se produjo un error con la siguiente excepción:  {0}\n".format(ex))
            return False

    def testodb(self, direct):  # testodb (odb = Oracle Data Base).
        # Funcion que hace ping a la direccion dada en el parametro, para establecer conexion a internet.
        log.write("# [func.testodb]:Llamada a función correcta. Params: direct={0}\n".format(direct))
        direccion = str(direct)
        resp = (os.system("ping -n 1 " + direccion))
        log.write("# [func.testodb]:Ping realizado:  ping -n 1 {0} ; Respuesta: {1}\n".format(direccion, resp))
        if resp == 0:
            log.write("# [func.testodb]:Existe conexión estable a la red.\n")
            self.alert("Success!", "¡Conexión estable a internet!")
            return True
        else:
            log.write("# [func.testodb]:Conexión no disponible...")
            self.alert("Failed!", "¡Revise su conexión a internet!")
            return False

    def odbc(self, dbruote):
        log.write("# [func.odbc]:Llamada a función correcta. Params: dbroute={0}\n".format(dbruote))
        try:
            """
            connects with odbc
            """
            db = dbruote
            constr = ('Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=' + db)
            log.write("# [func.odbc]:A punto de hacer conexión. String de Conexión={0}\n".format(constr))
            conn = pyodbc.connect(str(constr), autocommit=True)
            log.write("# [func.odbc]:Conexión: conn={0}\n".format(conn))
            self.alert("Success!", "¡Conexión exitosa!")
            log.write("# [func.odbc]:Conexión realizada con éxito...\n")
            conn.close()
            log.write("# [func.odbc]:Conexión a Access cerrada...\n")

            self.pbActualizar.setDisabled(False)
            self.pbTestADB.setDisabled(True)
            self.txtADB.setDisabled(True)

            return True
        except Exception as e:
            self.alert("Failed!",
                       "¡Se ha producido un error de conexión!" +
                       "\nRevise que el nombre de la base de datos Access sea correcta.")
            log.write("# [func.odbc]:Ha ocurrido la siguiente excepción:  {0}\n".format(e))
            return False

    def login(self, usr, pwd):
        log.write("# [fun.login]:Botón de Login presionado. Llamada a función correcta...\n")
        self.gtxResult.setText("")
        self.gtxResult.setText(str(self.gtxResult.toPlainText()) + "Enviando datos al servidor...\n")
        with CallServer(WSURL) as ws:
            self.gtxResult.setText(str(self.gtxResult.toPlainText()) + "Esperando respuesta del servidor...\n")
            response = ws.LogIn(usr.upper(), pwd)
            log.write("# [func.login]:Se evalua el json recibido: {0}\n".format(response))
            self.gtxResult.setText(str(self.gtxResult.toPlainText()) + "El servidor ha respondido...\n")
            self.gtxResult.setText(str(self.gtxResult.toPlainText()) + "Evaluando respuesta del servidor...\n")
            if response['ACK'] == '1':
                log.write("# [func.login]:Valores correctos. Inicio de Sesión exitoso...\n")
                self.gtxResult.setText(str(self.gtxResult.toPlainText()) + "Correcto... Ingreso exitoso.\n")
                self.alert("Control Condominio", "¡Usuario y Contraseña VÁLIDOS!")
                log.write("# [func.login]:Procede a deshabilitar sección de Inicio de Sesión.\n")
                self.massdisable(2)
                # self.txtADB.setText(response['PDB'])
            else:
                log.write("# [func.login]:Valores incorrectos. No Inicia Sesión...\n")
                self.gtxResult.setText(str(self.gtxResult.toPlainText()) + "Incorrecto... Ingreso fallido.\n")
                self.alert("Control Condominio", "¡Usuario y Contraseña NO VÁLIDOS!")

    def actualizar(self, dbruote, modo):
        log.write("# [func.actualizar]:Llamada a función correcta. "
                  "Params: dbroute={0} ; modo={1}\n".format(dbruote, modo))
        self.gtxResult.setText("")
        try:

            with CallServer(WSURL) as ws:
                strUsr = str(self.txtUsuario.text())
                strUsr = strUsr.upper()
                log.write("# [func.actualizar]:Variable de usuario utilizada: {0}\n".format(strUsr))
                response = ws.getEndDates(strUsr)
                log.write("# [func.actualizar]:Respuesta recibida del server: {0}\n".format(response))
                counterI = 0
                counterU = 0
                ubicador = 0

            if modo == 'odbc':
                # Conexion con odbc.

                db = dbruote
                constr = ('Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=' + db)
                log.write("# [func.actualizar]:Conexión a Access. String de conexión: {0}\n".format(constr))
                conn = pyodbc.connect(str(constr), autocommit=True)
                log.write("# [func.actualizar]:Conexión: conn={0}\n".format(conn))

                cur = conn.cursor()
                log.write("# [func.actualizar]:Conexión: cur={0}\n".format(cur))
                self.gtxResult.setText(str(self.gtxResult.toPlainText()) + "Iniciando interaccion"
                                                                           " en base de datos...\n")
                log.write("# [func.actualizar]:Inicio de bucle que ejecuta los updates/inserts...\n")

                QtGui.QApplication.processEvents()
                for codTarjeta, bloque in response.iteritems():
                    tipo = bloque[1]
                    log.write("# [func.actualizar]:Tipo a evaluar. 2=existe ; 1=nuevo... "
                              "tipo={0}\n".format(tipo))

                    if tipo == '2':
                        sql = "UPDATE {0} SET {1} = Format('{2}', 'yyyy-mm-dd') " \
                              "where Mid( Mid([EmployeeCode], InStr(1, [EmployeeCode], '-')+1)," \
                              "InStr(1, Mid([EmployeeCode], InStr(1, [EmployeeCode], '-')+1) , '-')+1) = '{4}'" \
                              .format("TEmployee", "EndDate", bloque[0][0:10], "EmployeeCode", codTarjeta)
                        log.write("# [func.actualizar]:Realiza update en Access. sql= {0}\n".format(sql))
                        cur.execute(sql)
                        counterU += 1
                        ubicador += 1
                        self.gtxResult.setText(str(self.gtxResult.toPlainText())
                                               + "Actualizado... {0}\n".format(counterU))
                        log.write("# [func.actualizar]:Update en Access realizado. Numero {0}\n".format(counterU))
                        for x in range(ubicador):
                            self.gtxResult.moveCursor(QtGui.QTextCursor.Down, QtGui.QTextCursor.MoveAnchor)
                        if ubicador % 4 == 1:
                            QtGui.QApplication.processEvents()
                    elif tipo == '1':
                        sql = "Insert into {0}" \
                              "(EmployeeID, EmployeeCode, EmployeeName, CardNo, pin, EmpEnable, Birthday, " \
                              "RegDate, EndDate, ACCESSID, Deleted, Leave, Password)" \
                              "values('{1}', '{1}', '{2}', '{3}', {4}, {5}, {6}, {7}, " \
                              "Format('{8}', 'yyyy-mm-dd'), {9}, {10}, {11}, '{12}')" \
                              .format("TEmployee", bloque[2], bloque[3].upper(), codTarjeta, 1234, 0, "Date()",
                                      "Date()", bloque[0][0:10], 0, 0, 0, 1234)
                        log.write("# [func.actualizar]:Realiza insert en Access. sql=  {0}\n".format(sql))
                        cur.execute(sql)
                        counterI += 1
                        ubicador += 1
                        self.gtxResult.setText(str(self.gtxResult.toPlainText())
                                               + "Ingresado... {0}\n".format(counterI))
                        log.write("# [func.actualizar]:Insert en Access realizado. Numero:: {0}\n".format(counterI))
                        for x in range(ubicador):
                            self.gtxResult.moveCursor(QtGui.QTextCursor.Down, QtGui.QTextCursor.MoveAnchor)
                        if ubicador % 4 == 1:
                            QtGui.QApplication.processEvents()

                log.write("\n# [func.actualizar]:Fin de bucle...\n")

                conn.close()
                log.write("# [func.actualizar]:Conexión a Access cerrada...\n")

                if counterI > 0:
                    self.alert("Success!", "¡Tarjetas Actualizadas! \n {0} Actualizaciones.".format(counterU) +
                               " \n {0} Ingresos Nuevos.".format(counterI))
                    log.write("# [func.actualizar]:Se insertaron nuevos registros. Registros nuevos: {0}  ;  "
                              "Actualizaciones: {1}\n".format(counterI, counterU))
                else:
                    self.alert("Success!", "¡Tarjetas Actualizadas! \n {0} Actualizaciones.".format(counterU))
                    log.write("# [func.actualizar]:Solo hubieron actualizaciones."
                              " Actualizaciones: {0}\n".format(counterU))

                log.write("# [func.actualizar]:Linea antes de cerrar el archivo.\n#### Fin del log ####")
                log.close()
                sys.exit(0)
            elif modo == 'ado':
                # Conexion con ado
                db = dbruote
                conn = win32com.client.Dispatch(r'ADODB.Connection')
                log.write("# [func.actualizar]:Conexión a Access en proceso... conn={0}\n".format(conn))
                dsn = ('PROVIDER = Microsoft.Jet.OLEDB.4.0;DATA SOURCE = ' + db + ';')
                log.write("# [func.actualizar]:Conexión: dsn={0}\n".format(dsn))
                conn.Open(dsn)

                rs = win32com.client.Dispatch(r'ADODB.Recordset')
                log.write("# [func.actualizar]:Conectado a Access... rs={0}\n".format(rs))
                self.gtxResult.setText(str(self.gtxResult.toPlainText()) + "Iniciando interaccion "
                                                                           "en base de datos...\n")
                log.write("# [func.actualizar]:Inicio de bucle que ejecuta los updates/inserts...\n")

                QtGui.QApplication.processEvents()
                for codTarjeta, bloque in response.iteritems():
                    tipo = bloque[1]
                    log.write("# [func.actualizar]:Tipo a evaluar. 2=existe ; 1=nuevo... "
                              "tipo={0}\n".format(tipo))

                    if tipo == '2':
                        sql = "UPDATE {0} SET {1} = Format('{2}', 'yyyy-mm-dd') " \
                              "where Mid( Mid([EmployeeCode], InStr(1, [EmployeeCode], '-')+1)," \
                              "InStr(1, Mid([EmployeeCode], InStr(1, [EmployeeCode], '-')+1) , '-')+1) = '{4}'" \
                              .format("TEmployee", "EndDate", bloque[0][0:10], "EmployeeCode", codTarjeta)
                        log.write("# [func.actualizar]:Realiza update en Access. sql= {0}\n".format(sql))
                        rs.Open(sql, conn, 1, 3)
                        counterU += 1
                        ubicador += 1
                        log.write("# [func.actualizar]:Update en Access realizado. Numero:: {0}\n".format(counterU))
                        self.gtxResult.setText(str(self.gtxResult.toPlainText())
                                               + "Actualizado... {0}\n".format(counterU))
                        for x in range(ubicador):
                            self.gtxResult.moveCursor(QtGui.QTextCursor.Down, QtGui.QTextCursor.MoveAnchor)
                        if ubicador % 4 == 1:
                            QtGui.QApplication.processEvents()
                    elif tipo == '1':
                        sql = "Insert into {0}" \
                              "(EmployeeID, EmployeeCode, EmployeeName, CardNo, pin, EmpEnable, Birthday, " \
                              "RegDate, EndDate, ACCESSID, Deleted, Leave, Password)" \
                              "values('{1}', '{1}', '{2}', '{3}', {4}, {5}, {6}, {7}, " \
                              "Format('{8}', 'yyyy-mm-dd'), {9}, {10}, {11}, '{12}')" \
                              .format("TEmployee", bloque[2], bloque[3].upper(), codTarjeta, 1234, 0, "Date()",
                                      "Date()", bloque[0][0:10], 0, 0, 0, 1234)
                        log.write("# [func.actualizar]:Realiza insert en Access. sql=  {0}\n".format(sql))
                        rs.Open(sql, conn, 1, 3)
                        counterI += 1
                        ubicador += 1
                        self.gtxResult.setText(str(self.gtxResult.toPlainText())
                                               + "Ingresado... {0}\n".format(counterI))
                        log.write("# [func.actualizar]:Insert en Access realizado. Numero:: {0}\n".format(counterI))
                        for x in range(ubicador):
                            self.gtxResult.moveCursor(QtGui.QTextCursor.Down, QtGui.QTextCursor.MoveAnchor)
                        if ubicador % 4 == 1:
                            QtGui.QApplication.processEvents()

                log.write("\n# [func.actualizar]:Fin de bucle...\n")

                conn.Close()
                log.write("# [func.actualizar]:Conexión a Access cerrada...\n")

                if counterI > 0:
                    self.alert("Success!", "¡Tarjetas Actualizadas! \n {0} Actualizaciones.".format(counterU) +
                               " \n {0} Ingresos Nuevos.".format(counterI))
                    log.write("# [func.actualizar]:Se insertaron nuevos registros. Registros nuevos: {0}  ;  "
                              "Actualizaciones: {1}\n".format(counterI, counterU))
                else:
                    self.alert("Success!", "¡Tarjetas Actualizadas! \n {0} Actualizaciones.".format(counterU))
                    log.write("# [func.actualizar]:Solo hubieron actualizaciones."
                              " Actualizaciones: {0}\n".format(counterU))

                log.write("# [func.actualizar]:Linea antes de cerrar el archivo.\n#### Fin del log ####")
                log.close()
                sys.exit(0)
            else:
                pass
        except Exception as exc:
            self.alert("Failed!", "¡Se ha producido un error al actualizar!")
            log.write("# [func.actualizar]:Se produjo un error con la siguiente excepción:  {0}\n".format(exc))
            log.write("#### Fin del log ####")
            log.close()


if __name__ == "__main__":
    app = QtGui.QApplication(sys.argv)
    window = TalanqueraUi()
    window.show()
    sys.exit(app.exec_())
