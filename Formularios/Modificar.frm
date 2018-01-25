VERSION 5.00
Begin VB.Form Modificar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modificar"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2820
   Icon            =   "Modificar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   2820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMatrix 
      Caption         =   "&Cerrar"
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdMatrix 
      Caption         =   "&Aceptar"
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Coloque el valor"
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2535
      Begin VB.TextBox txtMatrix 
         Height          =   285
         Index           =   1
         Left            =   1320
         MaxLength       =   5
         TabIndex        =   1
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtMatrix 
         Height          =   285
         Index           =   0
         Left            =   1320
         MaxLength       =   5
         TabIndex        =   0
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Qty Carton:"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Qty Label:"
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "Modificar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private indexItem As Integer, idProducto As Integer
Private nombreDeLvCola As String
Private objetoCola As New CCola
Public Sub MostrarModificar(lvListView As ListView)
    nombreDeLvCola = nombreLVCola
    For i = 1 To lvListView.ListItems.count
        If lvListView.ListItems(i).Checked = True Then
            txtMatrix(0).Text = lvListView.ListItems(i).ListSubItems(4).Text
            txtMatrix(1).Text = lvListView.ListItems(i).ListSubItems(5).Text
            nombreDeLvCola = lvListView.ListItems(i).ListSubItems(7).Text
            idProducto = lvListView.ListItems(i).ListSubItems(6).Text
            indexItem = i
            Exit For
        End If
    Next
    Me.Show 1
End Sub

Private Sub cmdMatrix_Click(Index As Integer)
    Select Case Index
        Case 0
            ModificarItem
        Case 1
            Unload Me
    End Select
End Sub
Private Sub ModificarItem()
    If nombreDeLvCola = "POR IMPRIMIR" Then
        Cola.lvCola.ListItems(indexItem).ListSubItems(4).Text = txtMatrix(0).Text
        Cola.lvCola.ListItems(indexItem).ListSubItems(5).Text = txtMatrix(1).Text
        objetoCola.ModificarCantidades Cola.lvCola.ListItems(indexItem).ListSubItems(6).Text, txtMatrix(0), txtMatrix(1)
    Else
        Cola.lvColaImpresa.ListItems(indexItem).ListSubItems(4).Text = txtMatrix(0).Text
        Cola.lvColaImpresa.ListItems(indexItem).ListSubItems(5).Text = txtMatrix(1).Text
        objetoCola.ModificarCantidades Cola.lvColaImpresa.ListItems(indexItem).ListSubItems(6).Text, txtMatrix(0), txtMatrix(1)
    End If
End Sub

Private Sub Form_Load()
    indexItem = 0
    idProducto = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    indexItem = 0
    idProducto = 0
End Sub
Private Sub txtMatrix_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case Else
        KeyAscii = 0
    End Select
End Sub
