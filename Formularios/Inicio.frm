VERSION 5.00
Begin VB.Form Inicio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inicio"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6885
   ControlBox      =   0   'False
   Icon            =   "Inicio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   120
      ScaleHeight     =   705
      ScaleWidth      =   4305
      TabIndex        =   6
      Top             =   1800
      Width           =   4305
   End
   Begin VB.Frame Frame1 
      Caption         =   "Servicio de Impresión"
      Height          =   1455
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   4695
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   255
         Left            =   3000
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "Demo"
         Height          =   495
         Index           =   1
         Left            =   1680
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Mostrar vista previa etiqueta"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "Iniciar Servicio"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Tag             =   "Iniciar"
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "Cerrar Servicio"
         Height          =   495
         Index           =   2
         Left            =   3120
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   2280
   End
End
Attribute VB_Name = "Inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nid As NOTIFYICONDATA
Function TrayName() As String
    TrayName = App.EXEName & " Ver." & App.Major & "." & App.Minor & vbCrLf & App.CompanyName & vbNullChar
End Function

Sub minToTray()
    nid.cbSize = Len(nid)
    nid.hWnd = Me.hWnd
    nid.uId = vbNull
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nid.uCallBackMessage = WM_MOUSEMOVE
    nid.hIcon = Inicio.Icon
    nid.szTip = TrayName
    Shell_NotifyIcon NIM_ADD, nid
End Sub
Private Sub cmdMatrix_Click(Index As Integer)
    Select Case Index
        Case 0
            If cmdMatrix(0).Tag = "Iniciar" Then
                If MsgBox("¿Desea dar inicio al servicio automático de impresión?", vbQuestion + vbYesNo, "Atención") = vbYes Then
                    Timer1.Interval = 10000
                    cmdMatrix(0).Caption = "Detener Servicio"
                    cmdMatrix(0).Tag = "Detener"
                    Me.Hide
                    minToTray
                    Timer1_Timer
                End If
            Else
                Timer1.Interval = 0
                cmdMatrix(0).Caption = "Iniciar Servicio"
                cmdMatrix(0).Tag = "Iniciar"
            End If
        Case 1
            DemoEtiquetar.Show
        Case 2
            End
    End Select
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Msg As Long
    Dim sFilter As String
    Msg = x / Screen.TwipsPerPixelX

        Select Case Msg
            Case WM_LBUTTONDOWN
            Case WM_LBUTTONUP
            Case WM_LBUTTONDBLCLK
                Me.Show
                Me.WindowState = vbNormal

                Shell_NotifyIcon NIM_DELETE, nid
            Case WM_RBUTTONDOWN
                PopupMenu mExit
            Case WM_RBUTTONUP
            Case WM_RBUTTONDBLCLK
        End Select
End Sub

Private Sub Timer1_Timer()
On Error GoTo CtrErr
    Dim rsLabel As New ADODB.Recordset
    Dim sqlSQL As String
    Dim numeroLabel As Integer
    If Timer1.Interval <> 0 Then
        sqlSQL = "SELECT * FROM print_queue"
        rsLabel.Open sqlSQL, cnx, adOpenStatic
        If rsLabel.EOF = False Then
            numeroLabel = rsLabel!ID
            printBar128 rsLabel!partno
            arLabel.DataControl1.Source = "SELECT * FROM print_queue WHERE id = " & numeroLabel & ""
            If Check1.Value = 0 Then
                arLabel.Printer.StartJob "IMPRESION DE ETIQUETA"
            Else
                arLabel.Show 1
            End If
            sqlSQL = "UPDATE print_queue SET status = 1 WHERE id = " & numeroLabel & ""
            cnx.Execute sqlSQL
        End If
        rsLabel.Close
        Set rsLabel = Nothing
    End If
    Exit Sub
CtrErr:
    Select Case Err.Number
        Case 2009
            arLabel.Printer.AbortJob
        Case Else
            MsgBox Err.Description, vbExclamation
    End Select
End Sub
Private Sub printBar128(valorBR As String)


    'Combination of bar
    'Start Character    3 character (Fixed)
    'Data
    'Check Character    3 character (Depends upon then value of the bar)
    'Stop Character     4 character (Fixed)
    
    
    '//######################################################
'    //PARAMETERS AND THIER MEANING
'    //a=LEFT
'    //b=TOP
'    //hgt=Height of the Barcode
'    //width=width of the thin Barcode in pixel
'    //r1=ratio of the thick barcode and thin barcode
'    //str=Value of the barcode
'    //align=alignment ofthe text i.e 1=left,2=center,3=right,4=justify
'    //textdisp= text position with respect to barcode i.e 2=TOP or 1=BOTTOM
'    //extra=distance of the text from the barcode
'    //ln=device context of the out put device
'//######################################################
Dim Dl As Long
Dim MinWidth  As Long
Picture1.Cls

    
Dim RT_VAL As RET_VAL
    
    With bar
        .crBack = RGB(255, 255, 255)
        .crFore = RGB(0, 0, 0)
        .lalign = 1 'Alignment of the text
        .lExtra = 4   'Distance between the barcode and the text
        .lheight = 40  'Height of the bar
        .lLeft = 20     'Left Position of the bar in the specified device (here e.g. Pictire1)
        .lR1 = 1        'Ratio between thin and thick bar (Standard all over world)
        .lR2 = 1        'Not necessary
        .lRetHeight = 0 'Returns the actual height of the bar code
        .lRetWidth = 0  'Returns the actual width of the bar code
        .lRotation = 0  'to rotate the bar code 0=0degree , 1=90 degree etc.
        .lShowCheck = 1 'Whether check digit will be displayed or not in the bar
        .lstyle = 1    'Bold, Italic, Underline or Strikethrough of text
        .lTop = 1      'Top Position of the bar in the specified device (here e.g. Picture1)
        .ltxtdisp = 2 'Whether text displayed at bottom(1) or top(2) of the bar
        .lWidth = 2    'Width of thin bar in pixel
        .nsize = 10     'Font Size of bar
        .szAdDigit = "" 'Not necessary
        .szBarCaption = valorBR
        .szDigit = ""   'Not necessary
        .szReadText = valorBR
        .szSymbology = 16
        .TextColor = RGB(255, 0, 0) 'Color of text
        .tiFaceName = "Courier New"       'Font name of text
    End With
    
       Set Target = Picture1
       Dl = Special_128b(bar, Target.hDc)
            
    
    If Dl <> 0 Then MsgBox ErrSpecial_128bMessage(Dl)
    On Error Resume Next
    Kill App.Path & "\test.bmp"
    Picture1.Picture = Picture1.Image
    SavePicture Picture1.Picture, App.Path & "\test.bmp"
End Sub
