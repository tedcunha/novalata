VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRELCLIENTES 
   Caption         =   "Relatórios de Clientes"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9645
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   9645
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab stTabComerciais 
      Height          =   3015
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   5318
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Clientes"
      TabPicture(0)   =   "frmRELCLIENTES.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Zona Geográfica"
      TabPicture(1)   =   "frmRELCLIENTES.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(2)=   "Frame3"
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame6 
         Caption         =   "[ Abre por Estados ]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   735
         Left            =   -71280
         TabIndex        =   28
         Top             =   1440
         Width           =   2295
         Begin VB.OptionButton optAbreEstadosSN 
            Caption         =   "Não"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   30
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton optAbreEstadosSN 
            Caption         =   "Sim"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "[ Ordem ]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   735
         Left            =   -74880
         TabIndex        =   25
         Top             =   1440
         Width           =   3495
         Begin VB.OptionButton optZonGeoOrdem 
            Caption         =   "Descrição"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   27
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton optZonGeoOrdem 
            Caption         =   "Código"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1095
         Left            =   -74880
         TabIndex        =   16
         Top             =   360
         Width           =   9375
         Begin VB.CommandButton Command2 
            Height          =   315
            Left            =   3240
            Picture         =   "frmRELCLIENTES.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   240
            Width           =   375
         End
         Begin VB.ComboBox cboGEOINI 
            Height          =   315
            Left            =   3600
            TabIndex        =   21
            Text            =   "cboGEOINI"
            Top             =   255
            Width           =   5655
         End
         Begin VB.TextBox txtCODZGEOINI 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2280
            TabIndex        =   20
            Text            =   "txtCODZGEOINI"
            Top             =   255
            Width           =   975
         End
         Begin VB.CommandButton Command1 
            Height          =   315
            Left            =   3240
            Picture         =   "frmRELCLIENTES.frx":013A
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   600
            Width           =   375
         End
         Begin VB.ComboBox cboGEOFIN 
            Height          =   315
            Left            =   3600
            TabIndex        =   18
            Text            =   "cboGEOFIN"
            Top             =   615
            Width           =   5655
         End
         Begin VB.TextBox txtCODZGEOFIN 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2280
            TabIndex        =   17
            Text            =   "txtCODZGEOFIN"
            Top             =   615
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Zona Geografica Inicial"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label3 
            Caption         =   "Zona Geografica Final"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   615
            Width           =   2055
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   9375
         Begin VB.TextBox txtCODCLIFIN 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1800
            TabIndex        =   13
            Text            =   "txtCODCLIFIN"
            Top             =   615
            Width           =   975
         End
         Begin VB.ComboBox cboCLIEFIN 
            Height          =   315
            Left            =   3120
            TabIndex        =   12
            Text            =   "cboCLIEFIN"
            Top             =   615
            Width           =   6135
         End
         Begin VB.CommandButton cmdPesqCLIFIN 
            Height          =   315
            Left            =   2760
            Picture         =   "frmRELCLIENTES.frx":023C
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtCODCLIINI 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1800
            TabIndex        =   10
            Text            =   "txtCODCLIINI"
            Top             =   255
            Width           =   975
         End
         Begin VB.ComboBox cboCLIEINI 
            Height          =   315
            Left            =   3120
            TabIndex        =   9
            Text            =   "cboCLIEINI"
            Top             =   255
            Width           =   6135
         End
         Begin VB.CommandButton cmdPesqCLI 
            Height          =   315
            Left            =   2760
            Picture         =   "frmRELCLIENTES.frx":033E
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "Cliente Final"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   615
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Cliente Inicial"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "[ Ordem ]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   3855
         Begin VB.OptionButton optOrdemCLI 
            Caption         =   "Código"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optOrdemFIN 
            Caption         =   "Razão Social"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   5
            Top             =   240
            Width           =   1695
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      Begin VB.CommandButton cmdVoltar 
         Caption         =   "&Volta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         Picture         =   "frmRELCLIENTES.frx":0440
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdImpressao 
         Caption         =   "Im&prime"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   960
         Picture         =   "frmRELCLIENTES.frx":0542
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Exclui Empresa"
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmRELCLIENTES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho     As String
Public Linha        As Variant
Public FILIAL       As Integer
Public strAcesso    As String
Dim objBLBFunc      As Object
Dim objRELCLIENTES  As Object
Dim objPESQPADRAO   As Object
Dim objREL          As Object


Private Sub cboCLIEFIN_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboCLIEFIN, KeyAscii
End Sub

Private Sub cboCLIEFIN_Validate(Cancel As Boolean)
    If cboCLIEFIN.ListIndex > -1 Then txtCODCLIFIN.Text = cboCLIEFIN.ItemData(cboCLIEFIN.ListIndex)
End Sub

Private Sub cboCLIEINI_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboCLIEINI, KeyAscii
End Sub

Private Sub cboCLIEINI_Validate(Cancel As Boolean)
    If cboCLIEINI.ListIndex > -1 Then txtCODCLIINI.Text = cboCLIEINI.ItemData(cboCLIEINI.ListIndex)
End Sub

Private Sub cboGEOFIN_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboGEOFIN, KeyAscii
End Sub

Private Sub cboGEOFIN_Validate(Cancel As Boolean)
    If cboGEOFIN.ListIndex > -1 Then txtCODZGEOFIN.Text = cboGEOFIN.ItemData(cboGEOFIN.ListIndex)
End Sub

Private Sub cboGEOINI_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboGEOINI, KeyAscii
End Sub

Private Sub cboGEOINI_Validate(Cancel As Boolean)
    If cboGEOINI.ListIndex > -1 Then txtCODZGEOINI.Text = cboGEOINI.ItemData(cboGEOINI.ListIndex)
End Sub

Private Sub cmdImpressao_Click()
    If ConfereCampos = False Then Exit Sub
    If stTabComerciais.Tab = 0 Then Call ImprimirCliente
    If stTabComerciais.Tab = 1 Then Call ImprimirZonaGeografica
End Sub

Private Sub cmdPesqCLI_Click()

    ReDim arrCAMPOS(1 To 3, 1 To 4) As String
    ReDim arrTABELA(1 To 1) As String
    
    arrTABELA(1) = "Select * From SGI_CADCLIENTE"
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    
    arrCAMPOS(2, 1) = "SGI_CPFCNPJ"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "1500"
    
    arrCAMPOS(3, 1) = "SGI_RAZAOSOC"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "Razão Social"
    arrCAMPOS(3, 4) = "5000"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Clientes")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODCLIINI.Text = varRETORNO
    
    cboCLIEINI.ListIndex = -1
    txtCODCLIINI.SetFocus

End Sub

Private Sub cmdPesqCLIFIN_Click()

    ReDim arrCAMPOS(1 To 3, 1 To 4) As String
    ReDim arrTABELA(1 To 1) As String
    
    arrTABELA(1) = "Select * From SGI_CADCLIENTE"
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    
    arrCAMPOS(2, 1) = "SGI_CPFCNPJ"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "1500"
    
    arrCAMPOS(3, 1) = "SGI_RAZAOSOC"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "Razão Social"
    arrCAMPOS(3, 4) = "5000"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Clientes")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODCLIFIN.Text = varRETORNO
    
    cboCLIEFIN.ListIndex = -1
    txtCODCLIFIN.SetFocus

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objRELCLIENTES = Nothing
    Set objPESQPADRAO = Nothing
    Set objREL = Nothing
    Unload Me
End Sub

Private Sub Command1_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    arrTABELA(1) = "Select * From SGI_CADZONAGEO"
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRI"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "SGI_DESCRI"
    
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Zona Geográfica")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODZGEOFIN.Text = varRETORNO
    
    cboGEOFIN.ListIndex = -1
    txtCODZGEOFIN.SetFocus

End Sub

Private Sub Command2_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    arrTABELA(1) = "Select * From SGI_CADZONAGEO"
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRI"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "SGI_DESCRI"
    
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Zona Geográfica")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODZGEOINI.Text = varRETORNO
    
    cboGEOINI.ListIndex = -1
    txtCODZGEOINI.SetFocus

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objRELCLIENTES = CreateObject("RELCOMERCIAL.clsRELCLIENTES")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objBLBFunc.LimpaCampos frmRELCLIENTES
    
    objRELCLIENTES.FILIAL = FILIAL
    
    stTabComerciais.Tab = 0
    
    objRELCLIENTES.PreencheComboClientes cboCLIEINI
    objRELCLIENTES.PreencheComboClientes cboCLIEFIN
    
    objRELCLIENTES.PreencheComboZonaGeografica cboGEOINI
    objRELCLIENTES.PreencheComboZonaGeografica cboGEOFIN
    
    optOrdemCLI(0).value = True
    optZonGeoOrdem(0).value = True
    optAbreEstadosSN(1).value = True
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    
    
End Sub


Private Sub txtCODCLIFIN_GotFocus()
    objBLBFunc.SelecionaCampos txtCODCLIFIN.Name, frmRELCLIENTES
End Sub

Private Sub txtCODCLIFIN_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODCLIFIN.Text
End Sub

Private Sub txtCODCLIFIN_Validate(Cancel As Boolean)

    Dim i As Integer
    
    If Len(Trim(txtCODCLIFIN.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODCLIFIN.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODCLIFIN.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboCLIEFIN.ListIndex = -1
    For i = 0 To (cboCLIEFIN.ListCount - 1)
        If cboCLIEFIN.ItemData(i) = Str(Val(txtCODCLIFIN.Text)) Then cboCLIEFIN.ListIndex = i
    Next i
    
    If cboCLIEFIN.ListIndex = -1 Then
       MsgBox "Este Cliente não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODCLIFIN.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtCODCLIINI_GotFocus()
    objBLBFunc.SelecionaCampos txtCODCLIINI.Name, frmRELCLIENTES
End Sub

Private Sub txtCODCLIINI_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODCLIINI.Text
End Sub

Private Sub txtCODCLIINI_Validate(Cancel As Boolean)

    Dim i As Integer
    
    If Len(Trim(txtCODCLIINI.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODCLIINI.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODCLIINI.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboCLIEINI.ListIndex = -1
    For i = 0 To (cboCLIEINI.ListCount - 1)
        If cboCLIEINI.ItemData(i) = Str(Val(txtCODCLIINI.Text)) Then cboCLIEINI.ListIndex = i
    Next i
    
    If cboCLIEINI.ListIndex = -1 Then
       MsgBox "Este Cliente não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODCLIINI.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub


Private Function ConfereCampos() As Boolean
    
    ConfereCampos = False
    If stTabComerciais.Tab = 0 Then
        If Len(Trim(txtCODCLIINI.Text)) > 0 And Len(Trim(txtCODCLIFIN.Text)) > 0 Then
           If CLng(txtCODCLIINI.Text) > CLng(txtCODCLIFIN.Text) Then
              MsgBox "Cliente Inicial não pode ser maior que Cliente Final !!!", vbOKOnly + vbExclamation, "Aviso"
              txtCODCLIINI.SetFocus
              Exit Function
           End If
        End If
    ElseIf stTabComerciais.Tab = 0 Then
        If Len(Trim(txtCODZGEOINI.Text)) > 0 And Len(Trim(txtCODZGEOFIN.Text)) > 0 Then
           If CLng(txtCODZGEOINI.Text) > CLng(txtCODZGEOFIN.Text) Then
              MsgBox "Zona Geográfica Inicial não pode ser maior que Zona Geográfica Final !!!", vbOKOnly + vbExclamation, "Aviso"
              txtCODZGEOINI.SetFocus
              Exit Function
           End If
        End If
    End If
    ConfereCampos = True

End Function

Private Sub ImprimirCliente()
    
    Dim strCABEC1 As String
    Dim strCABEC2 As String
    
    sSql = "Select "
    sSql = sSql & "       SGI_CADCLIENTE.SGI_CODIGO "
    sSql = sSql & "     , SGI_CADCLIENTE.SGI_RAZAOSOC "
    sSql = sSql & "     , SGI_CADCLIENTE.SGI_CPFCNPJ "
    sSql = sSql & "     , SGI_CADCLIENTE.SGI_DTCADASTRO "
    sSql = sSql & "  From "
    sSql = sSql & "       SGI_CADCLIENTE SGI_CADCLIENTE "
    sSql = sSql & " Where "
    sSql = sSql & "       SGI_CADCLIENTE.SGI_FILIAL = " & FILIAL
    
    If Len(Trim(txtCODCLIINI.Text)) > 0 And Len(Trim(txtCODCLIFIN.Text)) = 0 Then
        sSql = sSql & "   And SGI_CADCLIENTE.SGI_CODIGO = " & txtCODCLIINI.Text
    ElseIf Len(Trim(txtCODCLIINI.Text)) > 0 And Len(Trim(txtCODCLIFIN.Text)) > 0 Then
        sSql = sSql & "   And (SGI_CADCLIENTE.SGI_CODIGO >= " & txtCODCLIINI.Text & " And SGI_CADCLIENTE.SGI_CODIGO <= " & txtCODCLIFIN.Text & ")"
    End If
    
    sSql = sSql & " Order by "
    If optOrdemCLI(0).value = True Then
       sSql = sSql & "          SGI_CADCLIENTE.SGI_CODIGO "
       strCABEC2 = "Por Ordem de Código"
    ElseIf optOrdemCLI(1).value = True Then
       sSql = sSql & "          SGI_CADCLIENTE.SGI_RAZAOSOC "
       strCABEC2 = "Por Ordem de Razão Social"
    End If
    
    strCABEC1 = "Relatório de Clientes"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF Then
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    '' Chamada do Relatório
    Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelComercial & "RECLIENTES.rpt", Linha, 1, strCABEC1, strCABEC2, False)
   
End Sub

Private Sub txtCODZGEOFIN_GotFocus()
    objBLBFunc.SelecionaCampos txtCODZGEOFIN.Name, frmRELCLIENTES
End Sub

Private Sub txtCODZGEOFIN_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODZGEOFIN.Text
End Sub

Private Sub txtCODZGEOFIN_Validate(Cancel As Boolean)

    Dim i As Integer
    
    If Len(Trim(txtCODZGEOFIN.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODZGEOINI.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODZGEOFIN.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboGEOFIN.ListIndex = -1
    For i = 0 To (cboGEOFIN.ListCount - 1)
        If cboGEOFIN.ItemData(i) = Str(Val(txtCODZGEOFIN.Text)) Then cboGEOFIN.ListIndex = i
    Next i
    
    If cboGEOFIN.ListIndex = -1 Then
       MsgBox "Esta Zona Geográfica não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODZGEOFIN.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtCODZGEOINI_GotFocus()
    objBLBFunc.SelecionaCampos txtCODZGEOINI.Name, frmRELCLIENTES
End Sub

Private Sub txtCODZGEOINI_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODZGEOINI.Text
End Sub

Private Sub txtCODZGEOINI_Validate(Cancel As Boolean)

    Dim i As Integer
    
    If Len(Trim(txtCODZGEOINI.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODZGEOINI.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODZGEOINI.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboGEOINI.ListIndex = -1
    For i = 0 To (cboGEOINI.ListCount - 1)
        If cboGEOINI.ItemData(i) = Str(Val(txtCODZGEOINI.Text)) Then cboGEOINI.ListIndex = i
    Next i
    
    If cboGEOINI.ListIndex = -1 Then
       MsgBox "Esta Zona Geográfica não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODZGEOINI.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub ImprimirZonaGeografica()
    
    Dim strCABEC1 As String
    Dim strCABEC2 As String
    Dim strArq    As String
    
    sSql = "Select "
    sSql = sSql & "       SGI_CADZONAGEO.SGI_CODIGO "
    sSql = sSql & "     , SGI_CADZONAGEO.SGI_DESCRI "
    sSql = sSql & "  From "
    sSql = sSql & "       SGI_CADZONAGEO SGI_CADZONAGEO "
    sSql = sSql & " Where "
    sSql = sSql & "       SGI_CADZONAGEO.SGI_FILIAL = " & FILIAL
    
    If Len(Trim(txtCODZGEOINI.Text)) > 0 And Len(Trim(txtCODZGEOFIN.Text)) = 0 Then
        sSql = sSql & "   And SGI_CADZONAGEO.SGI_CODIGO = " & txtCODZGEOINI.Text
    ElseIf Len(Trim(txtCODZGEOINI.Text)) > 0 And Len(Trim(txtCODZGEOFIN.Text)) > 0 Then
        sSql = sSql & "   And (SGI_CADZONAGEO.SGI_CODIGO >= " & txtCODZGEOINI.Text & " And SGI_CADZONAGEO.SGI_CODIGO <= " & txtCODZGEOFIN.Text & ")"
    End If
    
    sSql = sSql & " Order by "
    If optZonGeoOrdem(0).value = True Then
       sSql = sSql & "          SGI_CADZONAGEO.SGI_CODIGO "
       strCABEC2 = "( Por Ordem de Código )"
    ElseIf optZonGeoOrdem(1).value = True Then
       sSql = sSql & "          SGI_CADZONAGEO.SGI_DESCRI "
       strCABEC2 = "( Por Ordem de Descrição )"
    End If
    
    strCABEC1 = "Relatório de Zona Geográfica"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF Then
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    If optAbreEstadosSN(0).value = True Then strArq = "REZONAGEOEST.rpt"
    If optAbreEstadosSN(1).value = True Then strArq = "REZONAGEO.rpt"
        
    '' Chamada do Relatório
    Call objREL.REL(FILIAL, sSql, strCamArgs & cCamRelComercial & strArq, Linha, 1, strCABEC1, strCABEC2, False)
   
End Sub

