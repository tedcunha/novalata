VERSION 5.00
Begin VB.Form frmCADPRODUTO 
   Caption         =   "Relatório de Produtos"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6330
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   6330
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame9 
      Caption         =   "[ Produtos ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   0
      TabIndex        =   29
      Top             =   2160
      Width           =   6255
      Begin VB.OptionButton optProdAtivInat 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   5040
         TabIndex        =   33
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optProdAtivInat 
         Caption         =   "Aguardando Liberação"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   32
         Top             =   240
         Width           =   2295
      End
      Begin VB.OptionButton optProdAtivInat 
         Caption         =   "Ativos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optProdAtivInat 
         Caption         =   "Inativos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "[ Forma de Visualização ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   0
      TabIndex        =   26
      Top             =   4560
      Width           =   6255
      Begin VB.OptionButton optTIPOVISU 
         Caption         =   "Em Excel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   1320
         TabIndex        =   28
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optTIPOVISU 
         Caption         =   "Em Tela"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "[ TIPO ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   0
      TabIndex        =   22
      Top             =   1560
      Width           =   6255
      Begin VB.OptionButton optTIPO 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   25
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton optTIPO 
         Caption         =   "Homologada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   24
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton optTIPO 
         Caption         =   "Normal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "[ EMPRESA ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   0
      TabIndex        =   18
      Top             =   3960
      Width           =   6255
      Begin VB.OptionButton optEMPRESA 
         Caption         =   "STEEL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   21
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optEMPRESA 
         Caption         =   "NOVALATA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   20
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optEMPRESA 
         Caption         =   "TODOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "[ Quebra ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   2640
      TabIndex        =   14
      Top             =   960
      Width           =   3615
      Begin VB.OptionButton optQUEBRA 
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optQUEBRA 
         Caption         =   "Capacidade"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optQUEBRA 
         Caption         =   "Rótulo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "[ Cliente ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   0
      TabIndex        =   10
      Top             =   3360
      Width           =   6255
      Begin VB.CommandButton cmdClie 
         Height          =   315
         Left            =   1080
         Picture         =   "frmCADPRODUTO.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtCodCliente 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         MaxLength       =   10
         TabIndex        =   11
         Text            =   "txtCodClie"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblDesclie 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDesclie"
         Height          =   315
         Left            =   1440
         TabIndex        =   13
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "[ Capacidade ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   2760
      Width           =   6255
      Begin VB.CommandButton cmdLinProd 
         Height          =   315
         Left            =   1080
         Picture         =   "frmCADPRODUTO.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtLinProd 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         MaxLength       =   10
         TabIndex        =   7
         Text            =   "txtLinProd"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblLinhProd 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblLinhProd"
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame Frame2 
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
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   2655
      Begin VB.OptionButton optOrdem 
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
         Left            =   1320
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optOrdem 
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
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
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
         Picture         =   "frmCADPRODUTO.frx":0204
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
         Picture         =   "frmCADPRODUTO.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Exclui Empresa"
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADPRODUTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public FILIAL           As Integer
Public strAcesso        As String
Public lngCodUsuario    As Long

Dim objBLBFunc      As Object
Dim objRELCADPROD   As Object
Dim objPESQPADRAO   As Object
Dim objREL          As Object

Private Sub cmdClie_Click()

    ReDim arrCAMPOS(1 To 3, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "700"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_RAZAOSOC"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Razão Social"
    arrCAMPOS(2, 4) = "3000"
    arrCAMPOS(2, 5) = "SGI_RAZAOSOC"
    
    arrCAMPOS(3, 1) = "SGI_NOMFANTA"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "Nome Fantasia"
    arrCAMPOS(3, 4) = "3000"
    arrCAMPOS(3, 5) = "SGI_NOMFANTA"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Clientes", "CADCLIENTE.clsCADCLIENTE")
    
    If Len(Trim(varRETORNO)) > 0 Then
       txtCodCliente.Text = varRETORNO
       lblDesclie.Caption = PegaDescClie(CLng(txtCodCliente.Text))
    End If
    txtCodCliente.SetFocus

End Sub

Private Sub cmdImpressao_Click()
    Call Imprime
End Sub

Private Sub cmdLinProd_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODLIN"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "700"
    arrCAMPOS(1, 5) = "SGI_CODLIN"
    
    arrCAMPOS(2, 1) = "SGI_DESCRI"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "3000"
    arrCAMPOS(2, 5) = "SGI_DESCRI"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Linha de Produto", "CADLINHAPROD.clsCADLINHAPROD")
    
    If Len(Trim(varRETORNO)) > 0 Then
       txtLinProd.Text = varRETORNO
       lblLinhProd.Caption = PegaDescLinProd(CLng(txtLinProd.Text))
    End If
    
    txtLinProd.SetFocus

End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub


Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objRELCADPROD = CreateObject("RELESTOQUE.clsRELCADPROD")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objRELCADPROD.FILIAL = FILIAL
    objBLBFunc.LimpaCampos frmCADPRODUTO
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    
    optOrdem(0).value = True
    optQUEBRA(0).value = True
    optEMPRESA(0).value = True
    optTIPO(0).value = True
    optTIPOVISU(0).value = True
    optProdAtivInat(3).value = True
    
    Call LimpaLabels
    
    '' --------------------------------------

End Sub

Private Sub DestroiObjeto()
    Set objBLBFunc = Nothing
    Set objRELCADPROD = Nothing
    Set objPESQPADRAO = Nothing
    Set objREL = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DestroiObjeto
End Sub


Private Sub LimpaLabels()
    lblLinhProd.Caption = ""
    lblDesclie.Caption = ""
End Sub

Private Function PegaDescLinProd(lngCodLinProd As Long) As String

    PegaDescLinProd = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO " & vbCrLf
    sSql = sSql & "  Where " & vbCrLf
    sSql = sSql & "        SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "    And SGI_CODLIN = " & lngCodLinProd
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaDescLinProd = BREC!SGI_DESCRI
    BREC.Close
    
End Function

Private Sub txtCodCliente_GotFocus()
    objBLBFunc.SelecionaCampos txtCodCliente.Name, frmCADPRODUTO
End Sub

Private Sub txtCodCliente_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCodCliente.Text
End Sub

Private Sub txtCodCliente_Validate(Cancel As Boolean)

    If Len(Trim(txtCodCliente.Text)) = 0 Then
       lblDesclie.Caption = ""
       Exit Sub
    End If
    
    If IsNumeric(txtCodCliente.Text) = False Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       Cancel = True
       Exit Sub
    End If
    
    lblDesclie.Caption = PegaDescClie(CLng(txtCodCliente.Text))
    If Len(Trim(lblDesclie.Caption)) = 0 Then
        MsgBox "Cliente Não Cadastrado !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCodCliente.Text = ""
        Cancel = True
        Exit Sub
    End If

End Sub

Private Sub txtLinProd_GotFocus()
    objBLBFunc.SelecionaCampos txtLinProd.Name, frmCADPRODUTO
End Sub

Private Sub txtLinProd_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtLinProd.Text
End Sub

Private Sub txtLinProd_Validate(Cancel As Boolean)

    If Len(Trim(txtLinProd.Text)) = 0 Then
        lblLinhProd.Caption = ""
        Exit Sub
    End If
    
    If IsNumeric(txtLinProd.Text) = False Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtLinProd.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    lblLinhProd.Caption = PegaDescLinProd(CLng(txtLinProd.Text))
    If Len(Trim(lblLinhProd.Caption)) = 0 Then
       MsgBox "Linha de Produto não cadastrada !!!", vbOKOnly + vbExclamation, "Aviso"
       txtLinProd.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub
Private Function PegaDescClie(lngCodClie As Long) As String

    PegaDescClie = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
    sSql = sSql & "  Where " & vbCrLf
    sSql = sSql & "        SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "    And SGI_CODIGO = " & lngCodClie
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaDescClie = BREC!SGI_RAZAOSOC
    BREC.Close
    
End Function

Private Sub Imprime()

On Error GoTo err_Imp

    Dim boolExiste              As Boolean
    Dim strNomRel               As String
    Dim strCABEC1               As String
    Dim strCABEC2               As String
    Dim boolARVORE              As Boolean
    Dim arrDADOSTAB()           As String
    Dim arrDADOSTAB_STEEL()     As String
    Dim lngREGS                 As Long
    
    boolExiste = True
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO.SGI_IDPRODUTO " & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_CODLINPROD " & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_CODCLIE " & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_CODROTULO " & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_DIGVERIF " & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_CODIGO" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_PRODUTOTIPO " & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_PRODUTOESTILO " & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_CODTIPO " & vbCrLf
    
    If optTIPOVISU(0).value = True Then
        If optQUEBRA(1).value = True Then
            sSql = sSql & "      ,SGI_CADLINHAPRODUTO.SGI_DESCRI" & vbCrLf
        ElseIf optQUEBRA(2).value = True Then
            sSql = sSql & "      ,SGI_CADCLIENTE.SGI_RAZAOSOC" & vbCrLf
        End If
    Else
        sSql = sSql & "      ,SGI_CADLINHAPRODUTO.SGI_DESCRI" & vbCrLf
        sSql = sSql & "      ,SGI_CADCLIENTE.SGI_RAZAOSOC" & vbCrLf
    End If
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO      SGI_CADPRODUTO" & vbCrLf
    
    If optTIPOVISU(0).value = True Then
        If optQUEBRA(1).value = True Then
            sSql = sSql & "      ,SGI_CADLINHAPRODUTO SGI_CADLINHAPRODUTO" & vbCrLf
        ElseIf optQUEBRA(2).value = True Then
            sSql = sSql & "      ,SGI_CADCLIENTE SGI_CADCLIENTE" & vbCrLf
        End If
    Else
        sSql = sSql & "      ,SGI_CADLINHAPRODUTO SGI_CADLINHAPRODUTO" & vbCrLf
        sSql = sSql & "      ,SGI_CADCLIENTE SGI_CADCLIENTE" & vbCrLf
    End If
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO.SGI_FILIAL        = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_CODIGO Is Not Null" & vbCrLf
     
    '' Status
    If optProdAtivInat(1).value = True Then
        sSql = sSql & "   And SGI_CADPRODUTO.SGI_STATUS = 1" & vbCrLf '' Ativo
    ElseIf optProdAtivInat(0).value = True Then
        sSql = sSql & "   And SGI_CADPRODUTO.SGI_STATUS = 0" & vbCrLf '' Inativo
    ElseIf optProdAtivInat(2).value = True Then
        sSql = sSql & "   And SGI_CADPRODUTO.SGI_STATUS = 2" & vbCrLf '' Aguardando Liberação
    End If
    
    If Len(Trim(txtLinProd.Text)) > 0 Then sSql = sSql & " And SGI_CADPRODUTO.SGI_CODLINPROD = " & Trim(txtLinProd.Text) & vbCrLf
    If Len(Trim(txtCodCliente.Text)) > 0 Then sSql = sSql & " And SGI_CADPRODUTO.SGI_CODCLIE = " & Trim(txtCodCliente.Text) & vbCrLf
        
    strCABEC1 = "Empresa : Todos"
    If optEMPRESA(1).value = True Then
        sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIALPED     = 0" & vbCrLf
        strCABEC1 = "Empresa : NOVALATA"
    ElseIf optEMPRESA(2).value = True Then
        sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIALPED     = 1" & vbCrLf
        strCABEC1 = "Empresa : STEEL"
    End If
    
    If optTIPO(0).value = True Then         '' normal
        sSql = sSql & "   And SGI_CADPRODUTO.SGI_CODTIPO       = 1" & vbCrLf
    ElseIf optTIPO(1).value = True Then     '' Homologada
        sSql = sSql & "   And SGI_CADPRODUTO.SGI_CODTIPO       = 2" & vbCrLf
    End If
    
    If optTIPOVISU(0).value = True Then
        If optQUEBRA(1).value = True Then
            sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIAL        = SGI_CADLINHAPRODUTO.SGI_FILIAL" & vbCrLf
            sSql = sSql & "   And SGI_CADPRODUTO.SGI_CODLINPROD    = SGI_CADLINHAPRODUTO.SGI_CODLIN" & vbCrLf
        ElseIf optQUEBRA(2).value = True Then
            sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIAL        = SGI_CADCLIENTE.SGI_FILIAL" & vbCrLf
            sSql = sSql & "   And SGI_CADPRODUTO.SGI_CODCLIE       = SGI_CADCLIENTE.SGI_CODIGO" & vbCrLf
        End If
    Else
        sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIAL        = SGI_CADLINHAPRODUTO.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And SGI_CADPRODUTO.SGI_CODLINPROD    = SGI_CADLINHAPRODUTO.SGI_CODLIN" & vbCrLf
        sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIAL        = SGI_CADCLIENTE.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And SGI_CADPRODUTO.SGI_CODCLIE       = SGI_CADCLIENTE.SGI_CODIGO" & vbCrLf
    End If
    
    If optQUEBRA(0).value = True Then
        If optOrdem(0).value = True Then sSql = sSql & "Order By SGI_CADPRODUTO.SGI_CODIGO"
        If optOrdem(1).value = True Then sSql = sSql & "Order By SGI_CADPRODUTO.SGI_DESCRICAO"
    ElseIf optQUEBRA(1).value = True Then
        If optOrdem(0).value = True Then sSql = sSql & "Order By SGI_CADPRODUTO.SGI_CODLINPROD"
        If optOrdem(1).value = True Then sSql = sSql & "Order By SGI_CADLINHAPRODUTO.SGI_DESCRI"
    ElseIf optQUEBRA(2).value = True Then
        If optOrdem(0).value = True Then sSql = sSql & "Order By SGI_CADPRODUTO.SGI_CODCLIE"
        If optOrdem(1).value = True Then sSql = sSql & "Order By SGI_CADCLIENTE.SGI_RAZAOSOC"
    End If
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF() Then
       MsgBox "Não existe dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       boolExiste = False
    Else
        '' Em Excel
        If optTIPOVISU(1).value = True Then
            strNomRel = "RELPRODUTO.xls"
            
            If Not BREC.EOF() Then
                lngREGS = 0
                Do While Not BREC.EOF()
                    lngREGS = (lngREGS + 1)
                    BREC.MoveNext
                Loop
                
                ReDim arrDADOSTAB(1 To lngREGS, 1 To 5) As String
                BREC.MoveFirst
                lngREGS = 0
                Do While Not BREC.EOF()
                    lngREGS = (lngREGS + 1)
                    
                    arrDADOSTAB(lngREGS, 1) = Trim(BREC!SGI_CODIGO)
                    arrDADOSTAB(lngREGS, 2) = Trim(BREC!SGI_DESCRICAO)
                    arrDADOSTAB(lngREGS, 3) = Trim(BREC!SGI_RAZAOSOC)
                    arrDADOSTAB(lngREGS, 4) = Trim(BREC!SGI_DESCRI)
                    If BREC!SGI_CODTIPO = 2 Then
                        arrDADOSTAB(lngREGS, 5) = "HOMOLOGADA"
                    Else
                        arrDADOSTAB(lngREGS, 5) = "NORMAL"
                    End If
                    
                    BREC.MoveNext
                Loop
            
            End If
        End If
    End If
    BREC.Close
    
    If optTIPOVISU(0).value Then
        If optQUEBRA(0).value = True Then
            strNomRel = "RELPRODUTOROT.rpt"
            boolARVORE = False
        ElseIf optQUEBRA(1).value = True Then
            strNomRel = "RELPRODUTOLIN.rpt"
            boolARVORE = True
        ElseIf optQUEBRA(2).value = True Then
            strNomRel = "RELPRODUTOCLIE.rpt"
            boolARVORE = True
        End If
        strCABEC2 = "Relatório de Produtos"
    End If
    
    If boolExiste = True Then
        If optTIPOVISU(0).value = True Then
            '' Chamada do Relatório
            Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelEstoque & strNomRel, Linha, 1, strCABEC2, strCABEC1, boolARVORE)
        Else
            Call ExportaParaExcel(arrDADOSTAB, lngREGS)
        End If
    End If

    Exit Sub

err_Imp:

    MsgBox "Erro N: " & Err.Number & vbCrLf & _
           "Descr : " & Err.Description, vbOKOnly + vbExclamation, "Aviso"
           
           

End Sub


Private Sub ExportaParaExcel(arrDADOSTAB() As String, lngQTDREGSNOVA As Long)

On Error GoTo Handle_Error

    Dim myExcelFile             As New clsExcelFile
    Dim FileName$
    Dim boolTemDados            As Boolean
    
    Dim lngREGS                 As Long
    Dim lngLINHA                As Long
    Dim lngQTDPED               As Long
    Dim lngQTDFAT               As Long
    Dim lngSALDO                As Long

    If lngQTDREGSNOVA = 0 Then
        MsgBox "Atenção - Não há dados para gerar o arquivo !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If

    With myExcelFile
        'Create the new spreadsheet
        If optEMPRESA(0).value = True Then
           FileName$ = strCamRelNovo & "RELPREPARA\RELPRODUTO_NS.xls"
        ElseIf optEMPRESA(1).value = True Then
           FileName$ = strCamRelNovo & "RELPREPARA\RELPRODUTO_NOVALATA.xls"
        ElseIf optEMPRESA(2).value = True Then
            FileName$ = strCamRelNovo & "RELPREPARA\RELPRODUTO_STEEL.xls"
        End If
        
        .CreateFile FileName$
        
        'set a Password for the file. If set, the rest of the spreadsheet will
        'be encrypted. If a password is used it must immediately follow the
        'CreateFile method.
        'This is different then protecting the spreadsheet (see below).
        'NOTE: For some reason this function does not work. Excel will
        'recognize that the file is password protected, but entering the password
        'will not work. Also, the file is not encrypted. Therefore, do not use
        'this function until I can figure out why it doesn't work. There is not
        'much documentation on this function available.
        '.SetFilePassword "PAUL"
        
        'specify whether to print the gridlines or not
        'this should come before the setting of fonts and margins
        .PrintGridLines = False
        
        'it is a good idea to set margins, fonts and column widths
        'prior to writing any text/numerics to the spreadsheet. These
        'should come before setting the fonts.
        
        .SetMargin xlsTopMargin, 1.5   'set to 1.5 inches
        .SetMargin xlsLeftMargin, 1.5
        .SetMargin xlsRightMargin, 1.5
        .SetMargin xlsBottomMargin, 1.5
        
        'to insert a Horizontal Page Break you need to specify the row just
        'after where you want the page break to occur. You can insert as many
        'page breaks as you wish (in any order).
        .InsertHorizPageBreak 10
        .InsertHorizPageBreak 20
        
        'set a default row height for the entire spreadsheet (1/20th of a point)
        .SetDefaultRowHeight 14
        
        'Up to 4 fonts can be specified for the spreadsheet. This is a
        'limitation of the Excel 2.1 format. For each value written to the
        'spreadsheet you can specify which font to use.
        
        .SetFont "Arial", 10, xlsNoFormat              'font0
        .SetFont "Arial", 10, xlsBold                  'font1
        .SetFont "Arial", 10, xlsBold + xlsUnderline   'font2
        .SetFont "Courier", 16, xlsBold + xlsItalic    'font3
        
        'Column widths are specified in Excel as 1/256th of a character.
        
        .SetColumnWidth 1, 1, 18
        .SetColumnWidth 2, 2, 60
        .SetColumnWidth 3, 3, 30
        .SetColumnWidth 4, 4, 60
        .SetColumnWidth 5, 5, 20
        
        
        'Set special row heights for row 1 and 2
        ''.SetRowHeight 1, 30
        ''.SetRowHeight 2, 30
        
        'set any header or footer that you want to print on
        'every page. This text will be centered at the top and/or
        'bottom of each page. The font will always be the font that
        'is specified as font0, therefore you should only set the
        'header/footer after specifying the fonts through SetFont.
        ''.SetHeader "BIFF 2.1 API"
        ''.SetFooter "Paul Squires - Excel BIFF Class"
        
        'write a normal left aligned string using font3 (Courier Italic)
        ''.WriteValue xlsText, xlsFont3, xlsLeftAlign, xlsNormal, 1, 1, "Quarterly Report"
        ''.WriteValue xlsText, xlsFont1, xlsLeftAlign, xlsNormal, 2, 1, "Cool Guy Corporation"
        
        'write some data to the spreadsheet
        'Use the default format #3 "#,##0" (refer to the WriteDefaultFormats function)
        'The WriteDefaultFormats function is compliments of Dieter Hauk in Germany.
        ''.WriteValue xlsinteger, xlsFont0, xlsLeftAlign, xlsNormal, 6, 1, 2000, 3
        
        'write a cell with a shaded number with a bottom border
        ''.WriteValue xlsnumber, xlsFont1, xlsrightAlign + xlsBottomBorder + xlsShaded, xlsNormal, 7, 1, 12123.456, 4
        
        'write a normal left aligned string using font2 (bold & underline)
        ''.WriteValue xlsText, xlsFont2, xlsLeftAlign, xlsNormal, 8, 1, "This is a test string"
        
        'write a locked cell. The cell will not be able to be overwritten, BUT you
        'must set the sheet PROTECTION to on before it will take effect!!!
        ''.WriteValue xlsText, xlsFont3, xlsLeftAlign, xlsLocked, 9, 1, "This cell is locked"
        
        'fill the cell with "F"'s
        ''.WriteValue xlsText, xlsFont0, xlsFillCell, xlsNormal, 10, 1, "F"
        
        'write a hidden cell to the spreadsheet. This only works for cells
        'that contain formula. Text, Number, Integer value text can not be hidden
        'using this feature. It is included here for the sake of completeness.
        ''.WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsHidden, 11, 1, "If this were a formula it would be hidden!"
        
        'write some dates to the file. NOTE: you need to write dates as xlsNumber
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 1, "Rótulo", 12
        .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, 1, 2, "Descrição Rótulo", 12
        .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, 1, 3, "Capacidade", 12
        .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, 1, 4, "Cliente", 12
        .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, 1, 5, "Tipo", 12
        
        If lngQTDREGSNOVA > 0 Then
        
            '' Jogando os Dados na Planilha
            '' NOVALATA
            lngLINHA = 1
            
            For lngREGS = 1 To UBound(arrDADOSTAB) '' Novalata
                lngLINHA = (lngLINHA + 1)
                
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 1, arrDADOSTAB(lngREGS, 1), 12
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 2, arrDADOSTAB(lngREGS, 2), 12
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 3, arrDADOSTAB(lngREGS, 4), 12
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 4, arrDADOSTAB(lngREGS, 3), 12
                .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 5, arrDADOSTAB(lngREGS, 5), 12
        
            Next lngREGS
        
        End If
        
        
        'PROTECT the spreadsheet so any cells specified as LOCKED will not be
        'overwritten. Also, all cells with HIDDEN set will hide their formula.
        'PROTECT does not use a password.
        .ProtectSpreadsheet = False 'False | True
        
        'Finally, close the spreadsheet
        .CloseFile
        
        MsgBox "Arquivo Excel : " & " foi Criado !", vbInformation + vbOKOnly, "Aviso do Sistema"
    End With
    
    Exit Sub
    
Handle_Error:

    If BREC.State = 1 Then BREC.Close
    MsgBox "Número: " & Err.Number & vbCrLf & "Descrição: " & Err.Description, vbOKOnly + vbCritical, "Aviso"

        
End Sub

