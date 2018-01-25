Attribute VB_Name = "AutoImpresionModuloPrincipal"
Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_LBUTTONDBLCLK = &H203 'Double-click
Public Const WM_LBUTTONDOWN = &H201 'Button down
Public Const WM_LBUTTONUP = &H202 'Button up
Public Const WM_RBUTTONDBLCLK = &H206 'Double-click
Public Const WM_RBUTTONDOWN = &H204 'Button down
Public Const WM_RBUTTONUP = &H205 'Button up
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public ADODBHelperMan As New ADODBHelper
Public cnx As ADODB.Connection
Public rs As ADODB.Recordset
Public objetoDBHelper As New ADODBHelper
Sub Main()
    If IniciarConexion Then
        Cola.Show 1
    End If
End Sub
Private Function IniciarConexion() As Boolean
    Set cnx = New ADODB.Connection
    Set rs = New ADODB.Recordset
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    On Error GoTo CtrErr
    'CONEXION MSACCESS
    'cnx.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BD\AutoImpresion.mdb;Persist Security Info=False"
    'CONEXION MySQL
    cnx.ConnectionString = objetoDBHelper.MySQLConnectionString
    cnx.Open
    IniciarConexion = True
    Exit Function
CtrErr:
    IniciarConexion = False
    Select Case Err.Number
        Case -2147467259
            MsgBox Err.Description, vbExclamation, "Error"
    End Select
End Function
