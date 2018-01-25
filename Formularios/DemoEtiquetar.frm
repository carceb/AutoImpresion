VERSION 5.00
Begin VB.Form DemoEtiquetar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Demo Etiquetar"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdMatrix 
      Caption         =   "&Cancelar"
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdMatrix 
      Caption         =   "&Aceptar"
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   3
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Valores de ejemplo"
      Height          =   3855
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   4695
      Begin VB.TextBox txtMatrix 
         Height          =   375
         Index           =   6
         Left            =   1200
         TabIndex        =   15
         Tag             =   "qty_label"
         Top             =   3240
         Width           =   3375
      End
      Begin VB.TextBox txtMatrix 
         Height          =   375
         Index           =   5
         Left            =   1200
         TabIndex        =   13
         Tag             =   "qty_label"
         Top             =   2760
         Width           =   3375
      End
      Begin VB.TextBox txtMatrix 
         Height          =   375
         Index           =   4
         Left            =   1200
         TabIndex        =   11
         Tag             =   "qty_label"
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox txtMatrix 
         Height          =   375
         Index           =   3
         Left            =   1200
         TabIndex        =   9
         Tag             =   "qty_label"
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox txtMatrix 
         Height          =   375
         Index           =   2
         Left            =   1200
         TabIndex        =   2
         Tag             =   "qty_label"
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox txtMatrix 
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   1
         Tag             =   "desc_proc"
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox txtMatrix 
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   0
         Tag             =   "art_des"
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "status:"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   3360
         Width           =   465
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "qty_carton:"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   2880
         Width           =   795
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "qty_label:"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   2400
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "type_label:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1920
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "origen:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "descripcion:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "part_no:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   585
      End
   End
End
Attribute VB_Name = "DemoEtiquetar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMatrix_Click(Index As Integer)
    Select Case Index
        Case 0
            AgregarEtiqueta
        Case 1
            Unload Me
    End Select
End Sub
Private Function EsTodoCorrecto() As Boolean
    EsTodoCorrecto = True
    For Each k In txtMatrix
        If k.Text = "" Then
            MsgBox "Debe agregar el " & k.Tag, vbCritical, "Error"
            EsTodoCorrecto = False
            Exit For
        End If
    Next
End Function
Private Sub AgregarEtiqueta()
    Dim sqlSQL As String
    If EsTodoCorrecto Then
        sqlSQL = "INSERT INTO print_queue(partno, descripcion, origen, type_label,qty_label,qty_carton) " & _
        "VALUES('" & UCase(txtMatrix(0)) & "',  '" & UCase(txtMatrix(1)) & "','" & txtMatrix(2) & "', " & UCase(txtMatrix(3)) & "," & UCase(txtMatrix(4)) & "," & UCase(txtMatrix(5)) & ")"
        cnx.Execute sqlSQL
        MsgBox "Registro actualizado", vbInformation
        LimpiarTodo
    End If
End Sub
Private Sub LimpiarTodo()
    For Each k In txtMatrix
        k.Text = ""
    Next
    txtMatrix(0).SetFocus
End Sub

Private Sub txtMatrix_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 3 To 6
            If KeyAscii = 46 Then
                KeyAscii = 44
            End If
            Select Case KeyAscii
            Case 48 To 57
            Case 44
            Case 8
            Case Else
            KeyAscii = 0
        End Select
    End Select
End Sub
