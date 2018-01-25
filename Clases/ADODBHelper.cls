VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ADODBHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function StandardQueryWithSQL(sSQL As String) As ADODB.Recordset
    On Error GoTo CtrErr
    Dim rsRS As New ADODB.Recordset
    Set rsRS = New ADODB.Recordset
    rsRS.Open sSQL, cnx, adOpenStatic, adLockPessimistic
    Set StandardQueryWithSQL = rsRS
    Exit Function
CtrErr:
    Set StandardQueryWithSQL = Nothing
    MsgBox Err.Description, vbCritical, "Error"
End Function
Public Function MySQLConnectionString() As String
    Dim columna() As String, lineaTexto As String, rutaDeArchivoConf As String
    Dim driver As String, database As String, server As String, user As String, password As String, port As String
    On Error GoTo CtrErr
    rutaDeArchivoConf = App.Path & "\BD\Conf.txt"
    'Lee del archivo de configuracion los parametros de conexion para MySQL
     '**************************************************************************************************
    Open rutaDeArchivoConf For Input As #1
    Line Input #1, lineaTexto
    Do
        Column = Split(lineaTexto, "!")
        driver = Column(0)
        database = Column(1)
        server = Column(2)
        user = Column(3)
        password = Column(4)
        port = Column(5)
        Line Input #1, lineaTexto
    Loop While Not EOF(1)
    Close #1
    MySQLConnectionString = "DRIVER= " & driver & ";DATABASE = " & database & ";SERVER = " & server & ";UID = " & user & ";PWD =" & password & ";PORT =" & port & ";OPTION =3"
    Exit Function
CtrErr:
    Select Case Err.Number
        Case 53
            MsgBox "No se econtró el archivo de configuración en la carpeta local, no puede conectarse al servidor.", vbExclamation, "Error"
        Case 62
            Resume Next
        Case 9
            Close #1
            MsgBox "El archivo de configuracion no presenta el formato solicitado, o tiene errores en su contenido.", vbCritical, "Error en el archivo enviado por la entidad"
        Case Else
            Close #1
            MsgBox Err.Description, vbExclamation
    End Select
End Function
