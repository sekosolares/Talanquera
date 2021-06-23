# -*- coding: iso-8859-1 -*-
import sys
import win32com.client
import pyodbc
import os
import time
# import psutil
from PyQt4 import QtGui, uic, Qt
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
        # Funcion que valida usuario y contrasenia.
        payload = {
            'cmd': 'logIn',
            'usr': userName,
            'pwd': userPass
        }
        log.write("# [func.LogIn]:Llamado desde la funcion de Login de la forma. Valida usuario y contrasenia."
                  " Parametros: Usuario={0} ; Contrasenia={1}\n".format(userName, userPass))

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
        log.write("# [func.getEndDates]:Llamada desde la funcion actualizar. Contacto al server con cmd=DeadEnd "
                  "y usr={0}\n".format(userName))
        self.__Responce = requests.post(self.__Host, params=payload)
        log.write("# [func.getEndDates]:Se ha realizado un post al server. "
                  "Respuesta del server: {0}\n".format(self.__Responce))
        jsonOBJ = self.__Responce.json()

        return jsonOBJ

    def updateEstado(self, sWhere='', estado=0):
        payload = {
            'cmd': 'cambioEstado',
            'where': sWhere,
            'dto': estado
        }
        log.write("# [func.updateEstado]:Llamada desde la funcion actualizar. Contacto al server con cmd=cambioEstado "
                  ",where={0} y dto={1}\n".format(sWhere, estado))
        self.__Responce = requests.post(self.__Host, params=payload)
        log.write("# [func.updateEstado]:Se ha realizado un post al server. "
                  "Respuesta del server: {0}\n".format(self.__Responce))

        jsonOBJ = self.__Responce.json()

        return jsonOBJ


# Url del servlet al cual se hacen los requests.
WSURL = "https://diceros.ls-sys.com/Sistema/talanquera"

os.system("md Logs")
nombre = 'Logs/log-'  # Nombre del archivo que lleva los logs del programa.
# Fecha y hora para concatenar con el log.
fecha = time.strftime("%d-%m-%Y_%H-%M-%S")
nombre += fecha
texto_info = ''
version_producto = "V 06.2021.1"

# Abriendo el archivo log para escribir en el.
log = open(nombre+'.txt', "w")


class TalanqueraUi(QtGui.QMainWindow, Ui_MainWindow):
    # Class principal. (Form)
    def __init__(self):
        # Con Frame:
        QtGui.QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)

        # Representa la funcion que se usara para interactuar con la DB Access.
        modoglobal = 'ado'
        # Se usa para verificar conexion a internet.
        condoaddress = "diceros.ls-sys.com"

        # Setup del frame. Se fija el tamanio de la pantalla para deshabilitar el boton de maximizar.
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
        log.write(
            "# [class.TalanqueraUi]:Se deshabilitan los campos, a excepcion de los campos de inicio de sesion.\n")
        self.massdisable(1)

        log.write(
            "# [class.TalanqueraUi]:A punto de probar conexion a internet.\n")
        if self.testodb(condoaddress):
            log.write(
                "# [class.TalanqueraUi]:Conexion exitosa a internet. Procede...\n")
            self.massdisable(1)
        else:
            log.write(
                "# [class.TalanqueraUi]:Fallo en la conexion a internet. Deshabilitar la forma.\n")
            self.massdisable()

        self.pbActualizar.setDisabled(True)
        self.txtADB.setDisabled(True)

        self.lblVersion.setText(version_producto)

        self.connect(self, Qt.SIGNAL('triggered()'), self.closeEvent)
        self.pbLogin.clicked.connect(lambda: self.login(
            str(self.txtUsuario.text()), str(self.txtPwd.text())))
        self.pbTestADB.clicked.connect(lambda: self.inter_access(modoglobal))
        self.pbActualizar.clicked.connect(
            lambda: self.inter_access('actualizar', modoglobal))

    def closeEvent(self, event):
        log.write(
            "# [func.closeEvent]:Boton 'X' presonado por el usuario. Salida del programa.\n")
        log.write("#### Fin del log ####")
        log.close()

    def massdisable(self, paso=0):
        log.write(
            "# [func.massdisable]:Llamada a funcion correcta. Params: paso={0}\n".format(paso))
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
            log.write(
                "# [func.massdisable]:Todo deshabilitado. No existio conexion a internet.\n")
        elif paso == 1:
            self.pbTestADB.setDisabled(stat)
            self.pbActualizar.setDisabled(stat)
            # self.txtADB.setDisabled(stat)
            self.lblAccessDB.setDisabled(stat)
            log.write(
                "# [func.massdisable]:La unica seccion habilitada es la de inicio de sesion.\n")
        elif paso == 2:
            self.pbTestADB.setDisabled(not stat)
            self.pbActualizar.setDisabled(stat)
            self.lblAccessDB.setDisabled(stat)
            self.txtUsuario.setDisabled(stat)
            self.txtPwd.setDisabled(stat)
            self.pbLogin.setDisabled(stat)
            self.txtADB.setDisabled(not stat)
            log.write("# [func.massdisable]:Inicio de sesion exitoso. "
                      "Se habilita el boton para probar conexion a base de datos Access.\n")

    def inter_access(self, destino, mod=None):
        log.write("# [func.inter_access]:Llamada a funcion para interaccion con Access correcta."
                  "Params: destino={0} ; mod={1}\n".format(destino, mod))
        dato = self.txtADB.text()
        path = str(dato)
        # self.txtADB.setText(path)
        if destino == 'odbc':
            log.write(
                "# [func.inter_access]:Se usa el modo '{0}' para conectar a Access.\n".format(destino))
            self.odbc(path)
        elif destino == 'ado':
            log.write(
                "# [func.inter_access]:Se usa el modo '{0}' para conectar a Access.\n".format(destino))
            self.ado(path)
        elif destino == 'actualizar':
            log.write(
                "# [func.inter_access]:Se actualizara la DB Access por medio de '{0}'.\n".format(mod))
            self.actualizar(path, mod)
        else:
            log.write(
                "# [func.inter_access]:Los parametros no fueron correctos!\n")
            pass

    def alert(self, title, text):
        QtGui.QMessageBox.about(self, title, text)

    def ado(self, dbpath):
        log.write(
            "# [func.ado]:Llamada a funcion correcta. Params: bdpath={0}\n".format(dbpath))
        try:
            """
            connect with com dispatch objs
            """
            db = dbpath
            conn = win32com.client.Dispatch(r'ADODB.Connection')
            log.write(
                "# [func.ado]:Conexion por iniciar... Variable conn={0}\n".format(conn))
            dsn = ('PROVIDER = Microsoft.Jet.OLEDB.4.0;DATA SOURCE = ' + db + ';')
            log.write(
                "# [func.ado]:String que se abrira para la conexion: dsn={0}\n".format(dsn))
            conn.Open(dsn)
            self.alert("CONNECTION SUCCESSFUL!", "Conexion exitosa!")
            log.write("# [func.ado]:Conexion a Access realizada con exito.\n")
            conn.Close()
            log.write("# [func.ado]:Conexion a Access cerrada...\n")

            self.pbActualizar.setDisabled(False)
            self.pbTestADB.setDisabled(True)
            self.txtADB.setDisabled(True)

            return True
        except Exception as ex:
            self.alert("CONNECTION ERROR!",
                       "Se ha producido un error de conexion!" +
                       "\nRevise que el nombre de la base de datos Access sea correcta.")
            log.write(
                "# [func.ado]:Se produjo un error con la siguiente excepcion:  {0}\n".format(ex))
            return False

    def testodb(self, direct):  # testodb (odb = Oracle Data Base).
        # Funcion que hace ping a la direccion dada en el parametro, para establecer conexion a internet.
        log.write(
            "# [func.testodb]:Llamada a funcion correcta. Params: direct={0}\n".format(direct))
        direccion = str(direct)
        process_to_find = "code.exe"
        os.system('tasklist /fi "ImageName eq {0}" /fo csv > process.txt'.format(process_to_find))
        log.write("# [func.testodb]:Archivo con informacion del tasklist creado.\n")
        process_file = open("process.txt", "r")
        text = process_file.read()
        text = text.upper()
        log.write("# [func.testodb]:Texto en el archivo: {0}.\n".format(text))
        process_found = process_to_find.upper() in text
        log.write("# [func.testodb]:Proceso corriendo? {0}.\n".format(process_found))
        process_file.close()
        log.write("# [func.testodb]:Cerrando archivo...\n")
        os.system("del process.txt")
        log.write("# [func.testodb]:Archivo eliminado.\n")
        if process_found:
            log.write("# [func.testodb]:El proceso {0} se encuentra corriendo. No se puede continuar.\n")
            self.alert("Failed!", "Debe cerrar la aplicacion de las tarjetas para poder continuar.")
            return False
        resp = (os.system("ping -n 1 " + direccion))
        log.write(
            "# [func.testodb]:Ping realizado:  ping -n 1 {0} ; Respuesta: {1}\n".format(direccion, resp))
        if resp == 0:
            log.write("# [func.testodb]:Existe conexion estable a la red.\n")
            self.alert("Success!", "Conexion estable a internet!")
            return True
        else:
            log.write("# [func.testodb]:Conexion no disponible...")
            self.alert("Failed!", "Revise su conexion a internet!")
            return False

    def odbc(self, dbruote):
        log.write(
            "# [func.odbc]:Llamada a funcion correcta. Params: dbroute={0}\n".format(dbruote))
        try:
            """
            connects with odbc
            """
            db = dbruote
            constr = (
                'Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=' + db)
            log.write(
                "# [func.odbc]:A punto de hacer conexion. String de Conexion={0}\n".format(constr))
            conn = pyodbc.connect(str(constr), autocommit=True)
            log.write("# [func.odbc]:Conexion: conn={0}\n".format(conn))
            self.alert("Success!", "Conexion exitosa!")
            log.write("# [func.odbc]:Conexion realizada con exito...\n")
            conn.close()
            log.write("# [func.odbc]:Conexion a Access cerrada...\n")

            self.pbActualizar.setDisabled(False)
            self.pbTestADB.setDisabled(True)
            self.txtADB.setDisabled(True)

            return True
        except Exception as e:
            self.alert("Failed!",
                       "Se ha producido un error de conexion!" +
                       "\nRevise que el nombre de la base de datos Access sea correcta.")
            log.write(
                "# [func.odbc]:Ha ocurrido la siguiente excepcion:  {0}\n".format(e))
            return False

    def login(self, usr, pwd):
        log.write(
            "# [fun.login]:Boton de Login presionado. Llamada a funcion correcta...\n")
        self.gtxResult.setText("")
        self.gtxResult.setText(
            str(self.gtxResult.toPlainText()) + "Enviando datos al servidor...\n")

        with CallServer(WSURL) as ws:
            self.gtxResult.setText(
                str(self.gtxResult.toPlainText()) + "Esperando respuesta del servidor...\n")
            response = ws.LogIn(usr.upper(), pwd)
            log.write(
                "# [func.login]:Se evalua el json recibido: {0}\n".format(response))
            self.gtxResult.setText(
                str(self.gtxResult.toPlainText()) + "El servidor ha respondido...\n")
            self.gtxResult.setText(
                str(self.gtxResult.toPlainText()) + "Evaluando respuesta del servidor...\n")
            if response['ACK'] == '1':
                log.write(
                    "# [func.login]:Valores correctos. Inicio de Sesion exitoso...\n")
                self.gtxResult.setText(
                    str(self.gtxResult.toPlainText()) + "Correcto... Ingreso exitoso.\n")
                self.alert("Control Condominio",
                           "Usuario y Contrasenia VALIDOS!")
                log.write(
                    "# [func.login]:Procede a deshabilitar seccion de Inicio de Sesion.\n")
                self.massdisable(2)
                self.txtADB.setText(response['PDB'])
            else:
                log.write(
                    "# [func.login]:Valores incorrectos. No Inicia Sesion...\n")
                self.gtxResult.setText(
                    str(self.gtxResult.toPlainText()) + "Incorrecto... Ingreso fallido.\n")
                self.alert("Control Condominio",
                           "Usuario y Contrasenia NO VALIDOS!")

    def actualizar(self, dbruote, modo):
        log.write("# [func.actualizar]:Llamada a funcion correcta. "
                  "Params: dbroute={0} ; modo={1}\n".format(dbruote, modo))
        self.gtxResult.setText("")
        try:

            with CallServer(WSURL) as ws:
                strUsr = str(self.txtUsuario.text())
                strUsr = strUsr.upper()
                log.write(
                    "# [func.actualizar]:Variable de usuario utilizada: {0}\n".format(strUsr))
                response = ws.getEndDates(strUsr)
                log.write(
                    "# [func.actualizar]:Respuesta recibida del server: {0}\n".format(response))
                counterI = 0
                counterU = 0
                ubicador = 0

            if modo == 'odbc':
                # Conexion con odbc.

                db = dbruote
                constr = (
                    'Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=' + db)
                log.write(
                    "# [func.actualizar]:Conexion a Access. String de conexion: {0}\n".format(constr))
                conn = pyodbc.connect(str(constr), autocommit=True)
                log.write(
                    "# [func.actualizar]:Conexion: conn={0}\n".format(conn))

                cur = conn.cursor()
                log.write(
                    "# [func.actualizar]:Conexion: cur={0}\n".format(cur))
                self.gtxResult.setText(str(self.gtxResult.toPlainText()) + "Iniciando interaccion"
                                                                           " en base de datos...\n")
                log.write(
                    "# [func.actualizar]:Inicio de bucle que ejecuta los updates/inserts...\n")

                QtGui.QApplication.processEvents()
                for codTarjeta, bloque in response.iteritems():
                    tipo = bloque[1]
                    log.write("# [func.actualizar]:Tipo a evaluar. 2=existe ; 1=nuevo... "
                              "tipo={0}\n".format(tipo))

                    if tipo == '2':
                        sql = "UPDATE {0} SET {1} = Format('{2}', 'yyyy-mm-dd') " \
                              "where ( Mid( Mid([EmployeeCode], InStr(1, [EmployeeCode], '-')+1)," \
                              "InStr(1, Mid([EmployeeCode], InStr(1, [EmployeeCode], '-')+1) , '-')+1) = '{4}' or CardNo = '{4}' )" \
                              .format("TEmployee", "EndDate", bloque[0][0:10], "EmployeeCode", codTarjeta)
                        log.write(
                            "# [func.actualizar]:Realiza update en Access. sql= {0}\n".format(sql))
                        cur.execute(sql)
                        counterU += 1
                        ubicador += 1
                        self.gtxResult.setText(str(self.gtxResult.toPlainText())
                                               + "Actualizado... {0}\n".format(counterU))
                        log.write(
                            "# [func.actualizar]:Update en Access realizado. Numero {0}\n".format(counterU))
                        for x in range(ubicador):
                            self.gtxResult.moveCursor(
                                QtGui.QTextCursor.Down, QtGui.QTextCursor.MoveAnchor)
                        if ubicador % 4 == 1:
                            QtGui.QApplication.processEvents()
                    elif tipo == '1':
                        with CallServer(WSURL) as ws:
                            respuesta = ws.updateEstado("empresa = {0} and agencia = {1} and casa = {2} "
                                                        "and linea = {3}"
                                                        .format(bloque[4], bloque[5], bloque[6], bloque[7]), 2)
                            log.write("# [func.actualizar]:Contacto con el servidor del sistema para hacer update de "
                                      "las tarjetas nuevas. Cambio de Nuevas a Existente. Respuesta del servidor: {0}\n"
                                      .format(respuesta))
                        if respuesta['ACK'] == '1':
                            sql = "INSERT INTO [{0}]" \
                                  "(EmployeeID, EmployeeCode, EmployeeName, CardNo, pin, EmpEnable, [Birthday], " \
                                  "[RegDate], [EndDate], ACCESSID, Deleted, Leave, Password)" \
                                  "values('{1}', '{1}', '{2}', '{3}', {4}, {5}, {6}, {7}, " \
                                  "Format('{8}', 'yyyy-mm-dd'), {9}, {10}, {11}, '{12}')" \
                                  .format("TEmployee", bloque[2], bloque[3].upper(), codTarjeta, 1234, 0, "Date()",
                                          "Date()", bloque[0][0:10], 0, 0, 0, 1234)
                            log.write(
                                "# [func.actualizar]:Realiza insert en Access. sql=  {0}\n".format(sql))
                            cur.execute(sql)
                            counterI += 1
                            ubicador += 1
                            self.gtxResult.setText(str(self.gtxResult.toPlainText())
                                                   + "Ingresado... {0}\n".format(counterI))
                            log.write(
                                "# [func.actualizar]:Insert en Access realizado. Numero:: {0}\n".format(counterI))
                            for x in range(ubicador):
                                self.gtxResult.moveCursor(
                                    QtGui.QTextCursor.Down, QtGui.QTextCursor.MoveAnchor)
                            if ubicador % 4 == 1:
                                QtGui.QApplication.processEvents()

                log.write("\n# [func.actualizar]:Fin de bucle...\n")

                conn.close()
                log.write("# [func.actualizar]:Conexion a Access cerrada...\n")

                if counterI > 0:
                    self.alert("Success!", "Tarjetas Actualizadas! \n {0} Actualizaciones.".format(counterU) +
                               " \n {0} Ingresos Nuevos.".format(counterI))
                    log.write("# [func.actualizar]:Se insertaron nuevos registros. Registros nuevos: {0}  ;  "
                              "Actualizaciones: {1}\n".format(counterI, counterU))
                else:
                    self.alert(
                        "Success!", "Tarjetas Actualizadas! \n {0} Actualizaciones.".format(counterU))
                    log.write("# [func.actualizar]:Solo hubieron actualizaciones."
                              " Actualizaciones: {0}\n".format(counterU))

                log.write(
                    "# [func.actualizar]:Linea antes de cerrar el archivo.\n#### Fin del log ####")
                log.close()
                sys.exit(0)

            elif modo == 'ado':
                # Conexion con ado
                db = dbruote
                conn = win32com.client.Dispatch(r'ADODB.Connection')
                log.write(
                    "# [func.actualizar]:Conexion a Access en proceso... conn={0}\n".format(conn))
                dsn = ('PROVIDER = Microsoft.Jet.OLEDB.4.0;DATA SOURCE = ' + db + ';')
                log.write(
                    "# [func.actualizar]:Conexion: dsn={0}\n".format(dsn))
                conn.Open(dsn)

                rs = win32com.client.Dispatch(r'ADODB.Recordset')
                log.write(
                    "# [func.actualizar]:Conectado a Access... rs={0}\n".format(rs))
                self.gtxResult.setText(str(self.gtxResult.toPlainText()) + "Iniciando interaccion "
                                                                           "en base de datos...\n")
                log.write(
                    "# [func.actualizar]:Inicio de bucle que ejecuta los updates/inserts...\n")

                QtGui.QApplication.processEvents()
                for codTarjeta, bloque in response.iteritems():
                    tipo = bloque[1]
                    log.write("# [func.actualizar]:Tipo a evaluar. 2=existe ; 1=nuevo... "
                              "tipo={0}\n".format(tipo))

                    if tipo == '2':
                        sql = "UPDATE {0} SET {1} = Format('{2}', 'yyyy-mm-dd') " \
                              "where ( Mid( Mid([EmployeeCode], InStr(1, [EmployeeCode], '-')+1)," \
                              "InStr(1, Mid([EmployeeCode], InStr(1, [EmployeeCode], '-')+1) , '-')+1) = '{4}' or CardNo = '{4}' )" \
                              .format("TEmployee", "EndDate", bloque[0][0:10], "EmployeeCode", codTarjeta)
                        log.write(
                            "# [func.actualizar]:Realiza update en Access. sql= {0}\n".format(sql))
                        rs.Open(sql, conn, 1, 3)
                        counterU += 1
                        ubicador += 1
                        log.write(
                            "# [func.actualizar]:Update en Access realizado. Numero:: {0}\n".format(counterU))
                        self.gtxResult.setText(str(self.gtxResult.toPlainText())
                                               + "Actualizado... {0}\n".format(counterU))
                        for x in range(ubicador):
                            self.gtxResult.moveCursor(
                                QtGui.QTextCursor.Down, QtGui.QTextCursor.MoveAnchor)
                        if ubicador % 4 == 1:
                            QtGui.QApplication.processEvents()

                    elif tipo == '1':
                        with CallServer(WSURL) as ws:
                            respuesta = ws.updateEstado("empresa = {0} and agencia = {1} and casa = {2} "
                                                        "and linea = {3}"
                                                        .format(bloque[4], bloque[5], bloque[6], bloque[7]), 2)
                            log.write("# [func.actualizar]:Contacto con el servidor del sistema para hacer update de "
                                      "las tarjetas nuevas. Cambio de Nuevas a Existente. Respuesta del servidor: {0}\n"
                                      .format(respuesta))

                        if respuesta['ACK'] == '1':
                            sql = "INSERT INTO [{0}]" \
                                  "(EmployeeCode, EmployeeName, CardNo)" \
                                  "VALUES" \
                                  "('{1}', '{2}', '{3}')" \
                                .format("TEmployee", bloque[2], bloque[3].upper(), codTarjeta)
                            log.write(
                                "# [func.actualizar]:Realiza insert en Access. sql=  {0}\n".format(sql))
                            rs.Open(sql, conn, 1, 3)
                            sql = "UPDATE {0} SET {1} = Format('{2}', 'yyyy-mm-dd')," \
                                  " RegDate = Format('{2}', 'yyyy-mm-dd'), Birthday = Format('{2}', 'yyyy-mm-dd')" \
                                  "where ( Mid( Mid([EmployeeCode], InStr(1, [EmployeeCode], '-')+1)," \
                                  "InStr(1, Mid([EmployeeCode], InStr(1, [EmployeeCode], '-')+1) , '-')+1) = '{4}' or CardNo = '{4}' )" \
                                .format("TEmployee", "EndDate", bloque[0][0:10], "EmployeeCode", codTarjeta)
                            rs.Open(sql, conn, 1, 3)
                            counterI += 1
                            ubicador += 1
                            self.gtxResult.setText(str(self.gtxResult.toPlainText())
                                                   + "Ingresado... {0}\n".format(counterI))
                            log.write(
                                "# [func.actualizar]:Insert en Access realizado. Numero:: {0}\n".format(counterI))
                            for x in range(ubicador):
                                self.gtxResult.moveCursor(
                                    QtGui.QTextCursor.Down, QtGui.QTextCursor.MoveAnchor)
                            if ubicador % 4 == 1:
                                QtGui.QApplication.processEvents()

                log.write("\n# [func.actualizar]:Fin de bucle...\n")

                conn.Close()
                log.write("# [func.actualizar]:Conexion a Access cerrada...\n")

                if counterI > 0:
                    self.alert("Success!", "Tarjetas Actualizadas! \n {0} Actualizaciones.".format(counterU) +
                               " \n {0} Ingresos Nuevos.".format(counterI))
                    log.write("# [func.actualizar]:Se insertaron nuevos registros. Registros nuevos: {0}  ;  "
                              "Actualizaciones: {1}\n".format(counterI, counterU))
                else:
                    self.alert(
                        "Success!", "Tarjetas Actualizadas! \n {0} Actualizaciones.".format(counterU))
                    log.write("# [func.actualizar]:Solo hubieron actualizaciones."
                              " Actualizaciones: {0}\n".format(counterU))

                log.write(
                    "# [func.actualizar]:Linea antes de cerrar el archivo.\n#### Fin del log ####")
                log.close()
                sys.exit(0)
            else:
                pass
        except Exception as exc:
            self.alert("Failed!", "Se ha producido un error al actualizar!")
            log.write(
                "# [func.actualizar]:Se produjo un error con la siguiente excepcion:  {0}\n".format(exc))
            log.write("#### Fin del log ####")
            log.close()


if __name__ == "__main__":
    app = QtGui.QApplication(sys.argv)
    window = TalanqueraUi()
    window.show()
    sys.exit(app.exec_())
