VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Cola 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cola de etiquetas"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12975
   Icon            =   "Cola.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   12975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSeleccionImpresora 
      Caption         =   "Seleccion de impresora"
      Height          =   375
      Left            =   7680
      TabIndex        =   11
      Top             =   8640
      Width           =   2175
   End
   Begin VB.CommandButton cmdMatrix 
      Caption         =   "&Salir"
      Height          =   495
      Index           =   1
      Left            =   11400
      TabIndex        =   6
      Top             =   8400
      Width           =   1335
   End
   Begin VB.CheckBox chkVistaPreliminar 
      Caption         =   "Vista preliminar"
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      Top             =   8280
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   240
      ScaleHeight     =   705
      ScaleWidth      =   4305
      TabIndex        =   8
      Top             =   8880
      Visible         =   0   'False
      Width           =   4305
   End
   Begin VB.CommandButton cmdMatrix 
      Caption         =   "&Imprimir"
      Height          =   495
      Index           =   0
      Left            =   9960
      TabIndex        =   5
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lista"
      Height          =   8055
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   12495
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "&Modificar"
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   10560
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "&Buscar codigo"
         Height          =   375
         Index           =   2
         Left            =   8880
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtMatrix 
         Height          =   285
         Index           =   0
         Left            =   7200
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin TabDlg.SSTab sstCola 
         Height          =   7335
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   12938
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Por Imprimir"
         TabPicture(0)   =   "Cola.frx":1CFA
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "lvCola"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Impresas"
         TabPicture(1)   =   "Cola.frx":1D16
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "lvColaImpresa"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin MSComctlLib.ListView lvCola 
            Height          =   6735
            Left            =   -74880
            TabIndex        =   0
            Top             =   480
            Width           =   11535
            _ExtentX        =   20346
            _ExtentY        =   11880
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   1
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "partno"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "descripcion"
               Object.Width           =   6174
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "origen"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "type_label"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "qty_label"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "qty_carton"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "id"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Estatus"
               Object.Width           =   2646
            EndProperty
         End
         Begin MSComctlLib.ListView lvColaImpresa 
            Height          =   6735
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Width           =   11535
            _ExtentX        =   20346
            _ExtentY        =   11880
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   1
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "partno"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "descripcion"
               Object.Width           =   6174
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "origen"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "type_label"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "qty_label"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "qty_carton"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "id"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Estatus"
               Object.Width           =   2293
            EndProperty
         End
      End
   End
   Begin VB.Label lblImpresas 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total etiquetas por imprimir: 2.500"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   345
      Left            =   240
      TabIndex        =   13
      Top             =   8520
      Visible         =   0   'False
      Width           =   4365
   End
   Begin VB.Label lblPorImprimir 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total etiquetas por imprimir: 2.500"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   345
      Left            =   240
      TabIndex        =   12
      Top             =   8520
      Visible         =   0   'False
      Width           =   4365
   End
End
Attribute VB_Name = "Cola"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private objetoCola As New CCola
Private Sub CargarCola()
    Dim lisItems As ListItem
    Dim lvColas As ListView
    Dim estatusDeCola As Integer
    Dim descripcionEstatusCola As String
    Dim totalEtiquetasPorImprimir As Integer
    Dim totalEtiquetasImpresas As Integer
    Dim rsRS As ADODB.Recordset
    totalEtiquetasPorImprimir = 0
    totalEtiquetasImpresas = 0
    For i = 1 To 2
        If i = 1 Then
            Set lvColas = lvCola
            estatusDeCola = 1
            descripcionEstatusCola = "POR IMPRIMIR"
        Else
            Set lvColas = lvColaImpresa
            estatusDeCola = 0
            descripcionEstatusCola = "IMPRESA"
        End If
        lvColas.ListItems.Clear
        Set rsRS = objetoCola.ObtenerCola(estatusDeCola, Trim(UCase(txtMatrix(0).Text)))
        If rsRS.EOF = False Then
            Do
                Set lisItems = lvColas.ListItems.Add(, , rsRS!partno)
                lisItems.Bold = True
                lisItems.ListSubItems.Add 1, , rsRS!descripcion
                lisItems.ListSubItems.Add 2, , rsRS!origen
                lisItems.ListSubItems.Add 3, , rsRS!type_label
                lisItems.ListSubItems.Add 4, , rsRS!qty_label
                lisItems.ListSubItems.Add 5, , rsRS!qty_carton
                lisItems.ListSubItems.Add 6, , rsRS!ID
                lisItems.ListSubItems.Add 7, , descripcionEstatusCola
                lisItems.ListSubItems(7).ForeColor = vbRed
                If i = 1 Then
                    totalEtiquetasPorImprimir = totalEtiquetasPorImprimir + 1
                Else
                    totalEtiquetasImpresas = totalEtiquetasImpresas + 1
                End If
                   rsRS.MoveNext
            Loop While Not rsRS.EOF
        End If
        rsRS.Close
    Next
    lblPorImprimir.Caption = "Total etiquetas por imprimir: " & totalEtiquetasPorImprimir
    lblImpresas.Caption = "Total etiquetas impresas: " & totalEtiquetasImpresas
End Sub
Private Sub cmdMatrix_Click(Index As Integer)
    Select Case Index
        Case 0
            ImprimirCola
        Case 1
            If MsgBox("¿Desea salir del sistema?", vbQuestion + vbOKCancel, "Atención") = vbOK Then
                End
            End If
        Case 2
            CargarCola
        Case 3
            ModificarValores
    End Select
End Sub

Private Sub Form_Load()
    Me.Caption = App.ProductName & " Versión " & App.Major & "." & App.Minor & " r" & App.Revision
    sstCola.Tab = 0
    CargarCola
    lblPorImprimir.Visible = True
End Sub
Private Sub ImprimirCola()
    Dim rsRS As ADODB.Recordset
    Dim lvColas As ListView
    Dim EsColaImpresa As Boolean
    If sstCola.Tab = 0 Then
        Set lvColas = lvCola
        EsColaImpresa = False
    Else
        Set lvColas = lvColaImpresa
        EsColaImpresa = True
    End If
    For i = 1 To lvColas.ListItems.count
        If lvColas.ListItems(i).Checked = True Then
            Set rsRS = objetoCola.ObtenerColaImprimir(lvColas.ListItems(i).ListSubItems(6).Text)
            If rsRS.EOF = False Then
                printBar128 rsRS!partno
                arLabel.DataControl1.Source = "SELECT * FROM print_queue WHERE id = " & rsRS!ID & ""
                If chkVistaPreliminar.Value = 0 Then
                    For x = 1 To CInt(lvColas.ListItems(i).ListSubItems(4).Text)
                        arLabel.PrintReport CBool(chkSeleccionImpresora.Value)
                    Next
                    'Limpia el objeto para recomenzar un nuevo reporte
                    Set arLabel = Nothing
                Else
                    arLabel.Show 1
                End If
                If EsColaImpresa = False Then
                    objetoCola.ActualizarColaImpresa rsRS!ID
                End If
            End If
            rsRS.Close
        End If
    Next
    CargarCola
End Sub
Private Sub ModificarValores()
    Dim lvColas As ListView
    Dim EsColaImpresa As Boolean
    If sstCola.Tab = 0 Then
        Set lvColas = lvCola
    Else
        Set lvColas = lvColaImpresa
    End If
    For i = 1 To lvColas.ListItems.count
        If lvColas.ListItems(i).Checked = True Then
            Modificar.MostrarModificar lvColas
        End If
    Next
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
        .lLeft = 5     'Left Position of the bar in the specified device (here e.g. Pictire1)
        .lR1 = 1        'Ratio between thin and thick bar (Standard all over world)
        .lR2 = 1        'Not necessary
        .lRetHeight = 0 'Returns the actual height of the bar code
        .lRetWidth = 0  'Returns the actual width of the bar code
        .lRotation = 0  'to rotate the bar code 0=0degree , 1=90 degree etc.
        .lShowCheck = 1 'Whether check digit will be displayed or not in the bar
        .lstyle = 1    'Bold, Italic, Underline or Strikethrough of text
        .lTop = 1      'Top Position of the bar in the specified device (here e.g. Picture1)
        .ltxtdisp = 1 'Whether text displayed at bottom(1) or top(2) of the bar
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("¿Desea salir del sistema?", vbQuestion + vbOKCancel, "Atención") = vbOK Then
        End
    End If
End Sub

Private Sub lvCola_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim totalChecked As Integer
    For i = 1 To lvCola.ListItems.count
        If lvCola.ListItems(i).Checked = True Then
            totalChecked = totalChecked + 1
        End If
    Next
    If totalChecked = 1 Then
        cmdMatrix(3).Enabled = True
    Else
        cmdMatrix(3).Enabled = False
    End If
    totalChecked = 0
End Sub

Private Sub lvColaImpresa_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim totalChecked As Integer
    For i = 1 To lvColaImpresa.ListItems.count
        If lvColaImpresa.ListItems(i).Checked = True Then
            totalChecked = totalChecked + 1
        End If
    Next
    If totalChecked = 1 Then
        cmdMatrix(3).Enabled = True
    Else
        cmdMatrix(3).Enabled = False
    End If
    totalChecked = 0
End Sub

Private Sub sstCola_Click(PreviousTab As Integer)
    If PreviousTab = 0 Then
        lblImpresas.Visible = True
        lblPorImprimir.Visible = False
    Else
        lblImpresas.Visible = False
        lblPorImprimir.Visible = True
    End If
End Sub

