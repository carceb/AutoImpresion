VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCola"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private claseSQL As String
Public Function ObtenerCola(ByVal estatusCola As Integer, ByVal codigoProducto As String) As ADODB.Recordset
    Dim rsColaP As New ADODB.Recordset
    If codigoProducto = "" Then
        claseSQL = "SELECT * FROM print_queue Where status = " & estatusCola & ""
    Else
        claseSQL = "SELECT * FROM print_queue Where status = " & estatusCola & " AND partno LIKE '%" & codigoProducto & "%'"
    End If
    Set ObtenerCola = ADODBHelperMan.StandardQueryWithSQL(claseSQL)
End Function
Public Function ObtenerColaImprimir(codigoRegistro As Integer) As ADODB.Recordset
    Dim rsColaP As New ADODB.Recordset
    claseSQL = "SELECT * FROM print_queue WHERE id = " & codigoRegistro & ""
    Set ObtenerColaImprimir = ADODBHelperMan.StandardQueryWithSQL(claseSQL)
End Function
Public Sub ActualizarColaImpresa(codigoRegistro As Integer)
    claseSQL = "UPDATE print_queue SET status = 0 WHERE id = " & codigoRegistro & ""
    cnx.Execute claseSQL
End Sub
Public Sub ModificarCantidades(productoID As Integer, qtyLabel As Integer, qtyCarton As Integer)
    claseSQL = "UPDATE print_queue SET qty_label = " & qtyLabel & ", qty_carton = " & qtyCarton & " WHERE id =" & productoID & ""
    cnx.Execute claseSQL
End Sub
