VERSION 5.00
Begin VB.Form frmMOVESTOQUE 
   Caption         =   "Movimentação de Estoque de Rótulos"
   ClientHeight    =   3675
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   8850
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "[ Filtro ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   6720
      TabIndex        =   24
      Top             =   960
      Width           =   2055
      Begin VB.OptionButton optFiltoSN 
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
         Index           =   1
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton optFiltoSN 
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
         Index           =   0
         Left            =   1080
         TabIndex        =   25
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame fraFiltros 
      Caption         =   "[ Filtros ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1575
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   8655
      Begin VB.CommandButton cmdPesq 
         Height          =   315
         Left            =   2760
         Picture         =   "frmMOVESTOQUE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtProduto 
         Height          =   315
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   2
         Text            =   "txtProduto"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   2760
         Picture         =   "frmMOVESTOQUE.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtCIDCLIE 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "txtCIDCLIE"
         Top             =   735
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   2760
         Picture         =   "frmMOVESTOQUE.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtCodLin 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "txtCodLin"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   23
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblDescProd 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescProd"
         Height          =   315
         Left            =   3120
         TabIndex        =   22
         Top             =   1080
         Width           =   5415
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblDescCliente 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescCliente"
         Height          =   285
         Left            =   3120
         TabIndex        =   19
         Top             =   720
         Width           =   5415
      End
      Begin VB.Label Label1 
         Caption         =   "Linha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblDescLinha 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescLinha"
         Height          =   285
         Left            =   3120
         TabIndex        =   16
         Top             =   360
         Width           =   5415
      End
   End
   Begin VB.Frame fraFiltro 
      Caption         =   "[ Ordem - Quebras ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   0
      TabIndex        =   9
      Top             =   960
      Width           =   3495
      Begin VB.OptionButton optFiltros 
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
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optFiltros 
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
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton optFiltros 
         Caption         =   "Linha"
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
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame fraSaldos 
      Caption         =   "[ Saldos ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   3600
      TabIndex        =   6
      Top             =   960
      Width           =   3015
      Begin VB.OptionButton optSaldos 
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
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton optSaldos 
         Caption         =   "> 0"
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
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton optSaldos 
         Caption         =   "= 0"
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
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8775
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
         Picture         =   "frmMOVESTOQUE.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Exclui Empresa"
         Top             =   240
         Width           =   855
      End
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
         Picture         =   "frmMOVESTOQUE.frx":0408
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmMOVESTOQUE"
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
Dim objRELMOVEST    As Object
Dim objPESQPADRAO   As Object
Dim objREL          As Object

Private Sub cmdImpressao_Click()
    Call Imprime
End Sub

Private Sub cmdPesq_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select Case PRO.SGI_PRODUTOTIPO " & vbCrLf
    sSql = sSql & "            When 1 Then replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD))) + '.' + " & vbCrLf
    sSql = sSql & "                        replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE))) + '.' + " & vbCrLf
    sSql = sSql & "                        replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO))) + '.' + " & vbCrLf
    sSql = sSql & "                        (Case " & vbCrLf
    sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Null Then '0' " & vbCrLf
    sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(2),PRO.SGI_DIGVERIF))) End) " & vbCrLf
    sSql = sSql & "            When 0 Then PRO.SGI_CODIGO End As SGI_CODIGO " & vbCrLf
    sSql = sSql & ",SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "         SGI_CADPRODUTO PRO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL     = " & FILIAL & vbCrLf
    ''sSql = sSql & "   And SGI_IDPRODUTO <> " & arrPROVARV(lngINDICE).lngProdutoID & vbCrLf
    ''sSql = sSql & "   And SGI_STATUS     = 1"
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "S"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "2000"
    arrCAMPOS(1, 5) = "PRO.SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "PRO.SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Produtos")
    
    If Len(Trim(varRETORNO)) > 0 Then
       txtProduto.Text = varRETORNO
       Call PegaProduto(varRETORNO)
       lblDescProd.Caption = PegaDescProd(txtProduto.Tag)
    End If
    txtProduto.SetFocus

End Sub

Private Sub Command1_Click()

    ReDim arrCAMPOS(1 To 5, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_CPFCNPJ"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "CNPJ"
    arrCAMPOS(2, 4) = "1500"
    arrCAMPOS(2, 5) = "SGI_CPFCNPJ"
    
    arrCAMPOS(3, 1) = "SGI_RAZAOSOC"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "Razão Social"
    arrCAMPOS(3, 4) = "3000"
    arrCAMPOS(3, 5) = "SGI_RAZAOSOC"
    
    arrCAMPOS(4, 1) = "SGI_NOMFANTA"
    arrCAMPOS(4, 2) = "S"
    arrCAMPOS(4, 3) = "Nome Fantasia"
    arrCAMPOS(4, 4) = "2000"
    arrCAMPOS(4, 5) = "SGI_NOMFANTA"
    
    arrCAMPOS(5, 1) = "SGI_CIDNORM"
    arrCAMPOS(5, 2) = "S"
    arrCAMPOS(5, 3) = "Cidade"
    arrCAMPOS(5, 4) = "1500"
    arrCAMPOS(5, 5) = "SGI_CIDNORM"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Clientes", "CADCLIENTE.clsCADCLIENTE")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCIDCLIE.Text = varRETORNO
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", varRETORNO, lblDescCliente)
    If Len(Trim(lblDescCliente.Caption)) = 0 Then txtCIDCLIE.Text = ""
    txtCIDCLIE.SetFocus

End Sub

Private Sub Command2_Click()

    ReDim arrCAMPOS(1 To 3, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "ID"
    arrCAMPOS(1, 4) = "800"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_CODLIN"
    arrCAMPOS(2, 2) = "N"
    arrCAMPOS(2, 3) = "Linha"
    arrCAMPOS(2, 4) = "1000"
    arrCAMPOS(2, 5) = "SGI_CODLIN"
    
    arrCAMPOS(3, 1) = "SGI_DESCRI"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "Descrição"
    arrCAMPOS(3, 4) = "5000"
    arrCAMPOS(3, 5) = "SGI_DESCRI"
    
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Linha de Produtos")
    
    If Len(Trim(varRETORNO)) > 0 Then
        txtCodLin.Tag = varRETORNO
        lblDescLinha.Caption = PegaDescrLinha(varRETORNO, "SGI_CODIGO")
    End If
    txtCodLin.SetFocus

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objRELMOVEST = Nothing
    Set objPESQPADRAO = Nothing
    Set objREL = Nothing
End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objRELMOVEST = CreateObject("RELESTOQUE.clsMOVESTOQUE")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objRELMOVEST.FILIAL = FILIAL
    objBLBFunc.LimpaCampos frmMOVESTOQUE
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    
    '' --------------------------------------
    optFiltros(0).Value = True
    optSaldos(2).Value = True
    optFiltoSN(0).Value = True

    Call LimpaCamposLabel

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

Private Sub LimpaCamposLabel()
    lblDescLinha.Caption = ""
    lblDescCliente.Caption = ""
    lblDescProd.Caption = ""
End Sub

Private Function PegaDescrLinha(strCodigo As String, strCampo As String) As String
    
    PegaDescrLinha = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And " & strCampo & " = " & strCodigo
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
       txtCodLin.Text = BREC!SGI_CODLIN
       txtCodLin.Tag = BREC!SGI_CODIGO
       PegaDescrLinha = BREC!SGI_DESCRI
    End If
    BREC.Close
    
End Function

Private Sub optFiltoSN_Click(Index As Integer)
    If Index = 1 Then fraFiltros.Enabled = True
    If Index = 0 Then fraFiltros.Enabled = False
End Sub

Private Sub txtCIDCLIE_GotFocus()
    objBLBFunc.SelecionaCampos txtCIDCLIE.Name, frmMOVESTOQUE
End Sub

Private Sub txtCIDCLIE_KeyPress(KeyAscii As Integer)
   objBLBFunc.SoNumeroPonto KeyAscii, txtCIDCLIE.Text
End Sub

Private Sub txtCIDCLIE_Validate(Cancel As Boolean)

    If Len(Trim(txtCIDCLIE.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCIDCLIE.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCIDCLIE.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", txtCIDCLIE.Text, lblDescCliente)
    If Len(Trim(lblDescCliente.Caption)) = 0 Then
       txtCIDCLIE.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtCodLin_GotFocus()
    objBLBFunc.SelecionaCampos txtCodLin.Name, frmMOVESTOQUE
End Sub

Private Sub PegaDescTabelas(StrCampoPesq As String, StrCampoRetorno As String, strTabela As String, strCodigo As String, lblLabel As Label)

    lblLabel.Caption = ""
    
    If Len(Trim(strCodigo)) = 0 Then Exit Sub
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoRetorno) & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & Trim(strTabela) & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoPesq) & " = " & Trim(strCodigo)
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then
       lblLabel.Caption = Trim(BREC10(Trim(StrCampoRetorno)))
    Else
       MsgBox "Registro Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC10.Close
    
End Sub

Private Sub txtCodLin_KeyPress(KeyAscii As Integer)
   objBLBFunc.SoNumeroPonto KeyAscii, txtCodLin.Text
End Sub

Private Sub txtCodLin_Validate(Cancel As Boolean)

   If Len(Trim(txtCodLin.Text)) = 0 Then
      lblDescLinha.Caption = ""
      Exit Sub
   End If
   
   If Not IsNumeric(txtCodLin.Text) Then
        MsgBox "Atenção - Somente é permitido numeros !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCodLin.Text = ""
        txtCodLin.Tag = ""
        lblDescLinha.Caption = ""
        Cancel = True
        Exit Sub
   End If
   
   lblDescLinha.Caption = PegaDescrLinha(txtCodLin.Text, "SGI_CODLIN")
   If Len(Trim(lblDescLinha.Caption)) = 0 Then
        MsgBox "Esta Linha não existe !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCodLin.Text = ""
        txtCodLin.Tag = ""
        Cancel = True
    End If

End Sub

Private Function PegaDescProd(strProdutoID As String) As String
    PegaDescProd = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_IDPRODUTO = " & strProdutoID
    
    BREC5.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC5.EOF() Then PegaDescProd = Trim(BREC5!SGI_DESCRICAO)
    BREC5.Close
    
End Function

Private Sub PegaProduto(strPRODUTO As String)

    sSql = ""
    
    sSql = "Select SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & ",Case PRO.SGI_PRODUTOTIPO " & vbCrLf
    sSql = sSql & "            When 1 Then replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD))) + '.' + " & vbCrLf
    sSql = sSql & "                        replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE))) + '.' + " & vbCrLf
    sSql = sSql & "                        replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO))) + '.' + " & vbCrLf
    sSql = sSql & "                        (Case " & vbCrLf
    sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Null Then '0' " & vbCrLf
    sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(2),PRO.SGI_DIGVERIF))) End) " & vbCrLf
    sSql = sSql & "            When 0 Then PRO.SGI_CODIGO End As SGI_CODIGO " & vbCrLf
    sSql = sSql & ",SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "         SGI_CADPRODUTO PRO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_STATUS = 1" & vbCrLf
    sSql = sSql & "   And (Case PRO.SGI_PRODUTOTIPO " & vbCrLf
    sSql = sSql & "            When 1 Then replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD))) + '.' + " & vbCrLf
    sSql = sSql & "                        replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE))) + '.' + " & vbCrLf
    sSql = sSql & "                        replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO))) + '.' + " & vbCrLf
    sSql = sSql & "                        (Case " & vbCrLf
    sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Null Then '0' " & vbCrLf
    sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(2),PRO.SGI_DIGVERIF))) End) " & vbCrLf
    sSql = sSql & "            When 0 Then PRO.SGI_CODIGO End) Like '" & Trim(strPRODUTO) & "%'"

    BREC4.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC4.EOF() Then
        txtProduto.Tag = Trim(Str(BREC4!SGI_IDPRODUTO))
    Else
        MsgBox "Produto Inexistente !!!", vbOKOnly + vbExclamation, "Aviso de Sistema"
    End If
    BREC4.Close
    
End Sub

Private Sub txtProduto_GotFocus()
    objBLBFunc.SelecionaCampos txtProduto.Name, frmMOVESTOQUE
End Sub

Private Sub txtProduto_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtProduto_Validate(Cancel As Boolean)

   If Len(Trim(txtProduto.Text)) = 0 Then Exit Sub
   
   Call PegaProduto(txtProduto.Text)
   lblDescProd.Caption = PegaDescProd(txtProduto.Tag)
   If Len(Trim(lblDescProd.Caption)) = 0 Then
      MsgBox "Este Produto não existe !!!", vbOKOnly + vbExclamation, "Aviso"
      Cancel = True
   End If

End Sub

Private Sub Imprime()

On Error GoTo err_Imp

    Dim boolExiste  As Boolean
    Dim strNomRel   As String
    Dim strCABEC2   As String
    
    boolExiste = True
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PROD.SGI_IDPRODUTO " & vbCrLf
    sSql = sSql & "      ,PROD.SGI_CODLINPROD " & vbCrLf
    sSql = sSql & "      ,PROD.SGI_CODCLIE " & vbCrLf
    sSql = sSql & "      ,PROD.SGI_CODROTULO " & vbCrLf
    sSql = sSql & "      ,PROD.SGI_DIGVERIF " & vbCrLf
    sSql = sSql & "      ,PROD.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "      ,PROD.SGI_PRODUTOTIPO " & vbCrLf
    sSql = sSql & "      ,PROD.SGI_PRODUTOESTILO " & vbCrLf
    sSql = sSql & "      ,PROD.SGI_SALDO " & vbCrLf
    
    sSql = sSql & "      ,SALD.SGI_CODCLIENTE" & vbCrLf
    ''sSql = sSql & "      ,SALD.SGI_SALDO" & vbCrLf
    
    sSql = sSql & "      ,CLIE.SGI_CODIGO" & vbCrLf
    sSql = sSql & "      ,CLIE.SGI_RAZAOSOC" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "        SGI_CADCLIENTE CLIE" & vbCrLf
    sSql = sSql & "       ,SGI_CADPRODUTO PROD" & vbCrLf
    sSql = sSql & "       ,SGI_PRODSALDOS SALD" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    
    sSql = sSql & "       PROD.SGI_FILIAL        = " & FILIAL & vbCrLf
    sSql = sSql & "   And PROD.SGI_PRODUTOTIPO   = 1" & vbCrLf
    sSql = sSql & "   And PROD.SGI_PRODUTOESTILO = 0" & vbCrLf
    
    
    If optSaldos(0).Value = True Then
        sSql = sSql & "   And (PROD.SGI_SALDO Is Null or PROD.SGI_SALDO = 0"""
    ElseIf optSaldos(1).Value = True Then
        sSql = sSql & "   And PROD.SGI_SALDO > 0"
    End If
    
    If Len(Trim(txtCodLin.Text)) > 0 Then sSql = sSql & " And PROD.SGI_CODLINPROD = " & Trim(txtCodLin.Text)
    If Len(Trim(txtCIDCLIE.Text)) > 0 Then sSql = sSql & " And PROD.SGI_CODCLIE = " & Trim(txtCIDCLIE.Text)
    If Len(Trim(txtProduto.Tag)) > 0 Then sSql = sSql & " And PROD.SGI_IDPRODUTO = " & Trim(txtProduto.Tag)
    
    sSql = sSql & "   And SALD.SGI_FILIAL        = PROD.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And SALD.SGI_IDPRODUTO     = PROD.SGI_IDPRODUTO  " & vbCrLf
     
    sSql = sSql & "   And SALD.SGI_FILIAL        = CLIE.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SALD.SGI_CODCLIENTE    = CLIE.SGI_CODIGO" & vbCrLf
    
    If optFiltros(0).Value = True Then sSql = sSql & "Order By PROD.SGI_CODLINPROD "
    If optFiltros(1).Value = True Then sSql = sSql & "Order By PROD.SGI_CODCLIE "
    If optFiltros(2).Value = True Then sSql = sSql & "Order By PROD.SGI_CODLINPROD,PROD.SGI_CODCLIE,PROD.SGI_CODROTULO,PROD.SGI_DIGVERIF"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF() Then
       MsgBox "Não existe dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       boolExiste = False
    End If
    BREC.Close
    
    If optFiltros(0).Value = True Then
        strNomRel = "RELSDPRODLIN.rpt"
        strCABEC2 = "Relatório de Saldos de Estoque [Por Linha de Produto]"
    ElseIf optFiltros(1).Value = True Then
        strNomRel = "RELSDPRODCLI.rpt"
        strCABEC2 = "Relatório de Saldos de Estoque [Por Cliente]"
    ElseIf optFiltros(2).Value = True Then
        strNomRel = "RELSDPRODROT.rpt"
        strCABEC2 = "Relatório de Saldos de Estoque [Por Rótulo]"
    End If
    
    If boolExiste = True Then
        '' Chamada do Relatório
        objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelEstoque & strNomRel, Linha, 1, strCABEC2, "", True
    End If

    Exit Sub

err_Imp:

    MsgBox "Erro N: " & Err.Number & vbCrLf & _
           "Descr : " & Err.Description, vbOKOnly + vbExclamation, "Aviso"
           
           

End Sub

