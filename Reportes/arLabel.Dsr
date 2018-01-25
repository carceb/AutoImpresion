VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arLabel 
   Caption         =   "Label"
   ClientHeight    =   13380
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   28560
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   50377
   _ExtentY        =   23601
   SectionData     =   "arLabel.dsx":0000
End
Attribute VB_Name = "arLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_Initialize()
    On Error GoTo CtrErr
    DataControl1.ConnectionString = objetoDBHelper.MySQLConnectionString
    Me.Printer.PaperSize = 1
    Me.Printer.Orientation = ddOPortrait
    Exit Sub
CtrErr:
    Select Case Err.Number
        Case -2147467259
            MsgBox "La impresora no está disponible", vbExclamation, "Error"
        Case Else
            MsgBox Err.Description, vbExclamation
    End Select
End Sub
Private Sub GroupHeader1_Format()
    imgBC.Picture = LoadPicture(App.Path & "\test.bmp")
    Me.documentName = "Etiqueta " & Field3.Text
End Sub
