VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRELMOVESTLIT 
   Caption         =   "Relatório de Movimentação de Estoque de Litografia"
   ClientHeight    =   4350
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10530
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   10530
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame8 
      Caption         =   "[ forma de Geração ]"
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
      TabIndex        =   28
      Top             =   3600
      Width           =   3615
      Begin VB.OptionButton optFormaGeracao 
         Caption         =   "Em arquivo EXCEL"
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
         Left            =   1560
         TabIndex        =   30
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton optFormaGeracao 
         Caption         =   "Impressão"
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
         TabIndex        =   29
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame7 
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
      Height          =   735
      Left            =   0
      TabIndex        =   24
      Top             =   2880
      Width           =   10455
      Begin VB.TextBox txtCodLin 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   26
         Text            =   "txtCodLin"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   1320
         Picture         =   "frmRELMOVESTLIT.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblDescLinha 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescLinha"
         Height          =   285
         Left            =   1680
         TabIndex        =   27
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "[ Setor ]"
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
      Height          =   735
      Left            =   0
      TabIndex        =   20
      Top             =   2160
      Width           =   10455
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   1320
         Picture         =   "frmRELMOVESTLIT.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtCIDCLIE 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   21
         Text            =   "txtCIDCLIE"
         Top             =   255
         Width           =   1215
      End
      Begin VB.Label lblDescCliente 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescCliente"
         Height          =   285
         Left            =   1680
         TabIndex        =   23
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "[ Empresa ]"
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
      Left            =   5520
      TabIndex        =   16
      Top             =   1560
      Width           =   4935
      Begin VB.OptionButton optEmpresa 
         Caption         =   "NOVALATA e STEEL"
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
         TabIndex        =   19
         Top             =   240
         Width           =   2175
      End
      Begin VB.OptionButton optEmpresa 
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
         Index           =   1
         Left            =   1440
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optEmpresa 
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
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "[ Ordem do Relatório ]"
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
      TabIndex        =   12
      Top             =   1560
      Width           =   5535
      Begin VB.OptionButton optOrdemRel 
         Caption         =   "Por Setor"
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
         Left            =   3360
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optOrdemRel 
         Caption         =   "Por Capacidade"
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
         TabIndex        =   14
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton optOrdemRel 
         Caption         =   "Por Data"
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
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "[ Tipo de Movimentação ]"
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
      Left            =   5520
      TabIndex        =   8
      Top             =   960
      Width           =   4935
      Begin VB.OptionButton optEntrSaid 
         Caption         =   "ENTRADAS E SAIDAS"
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
         TabIndex        =   11
         Top             =   240
         Width           =   2295
      End
      Begin VB.OptionButton optEntrSaid 
         Caption         =   "ENTRADAS"
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
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optEntrSaid 
         Caption         =   "SAIDAS"
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
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   5535
      Begin MSMask.MaskEdBox mskDTFIN 
         Height          =   285
         Left            =   3960
         TabIndex        =   3
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDTINI 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Data Inicial"
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
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Data Final"
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
         Left            =   2880
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10455
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
         Picture         =   "frmRELMOVESTLIT.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "frmRELMOVESTLIT.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Exclui Empresa"
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmRELMOVESTLIT"
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
Dim objRELMOVESTLIT As Object
Dim objPESQPADRAO   As Object
Dim objREL          As Object

Private Sub cmdImpressao_Click()
    If ConfereCampos = False Then Exit Sub
    If optEntrSaid(0).value = True Then Call ImprimeEntradas
End Sub

Private Sub Command1_Click()

    ReDim arrCAMPOS(1 To 6, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_CODIGO " & vbCrLf
    sSql = sSql & "      ,SGI_CPFCNPJ" & vbCrLf
    sSql = sSql & "      ,SGI_RAZAOSOC" & vbCrLf
    sSql = sSql & "      ,SGI_NOMFANTA" & vbCrLf
    sSql = sSql & "      ,SGI_CIDNORM" & vbCrLf
    sSql = sSql & "      ,SGI_CODREF" & vbCrLf
    
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    
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
    arrCAMPOS(3, 4) = "4500"
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
    
    arrCAMPOS(6, 1) = "SGI_CODREF"
    arrCAMPOS(6, 2) = "S"
    arrCAMPOS(6, 3) = "Cód.Antigo"
    arrCAMPOS(6, 4) = "1500"
    arrCAMPOS(6, 5) = "SGI_CODREF"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Setor")
    
    If Len(Trim(varRETORNO)) > 0 Then
        txtCIDCLIE.Text = varRETORNO
        Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", varRETORNO, lblDescCliente)
    End If
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

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objRELMOVESTLIT = CreateObject("RELESTOQUE.clsRELMOVESTLIT")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objRELMOVESTLIT.FILIAL = FILIAL
    objBLBFunc.LimpaCampos Me
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    
    optEntrSaid(0).value = True
    optOrdemRel(0).value = True
    optEmpresa(0).value = True
    optFormaGeracao(0).value = True
    
    mskDTINI.Text = Format(Now, "DD/MM/YYYY")
    mskDTFIN.Text = Format(Now, "DD/MM/YYYY")
    
    Call LimpaCamposLabel

End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objRELMOVESTLIT = Nothing
    Set objPESQPADRAO = Nothing
    Set objREL = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

Private Sub mskDTFIN_GotFocus()
    objBLBFunc.SelecionaCampos mskDTFIN.Name, Me
End Sub

Private Sub mskDTINI_GotFocus()
    objBLBFunc.SelecionaCampos mskDTINI.Name, Me
End Sub

Private Function ConfereCampos() As Boolean
    
    ConfereCampos = False
        
    If Not IsDate(mskDTINI.Text) Then
        MsgBox "Data Inicial Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
        mskDTINI.SetFocus
        Exit Function
    End If
    If Not IsDate(mskDTFIN.Text) Then
        MsgBox "Data Final Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
        mskDTFIN.SetFocus
        Exit Function
    End If
    
    If CDate(mskDTINI.Text) > CDate(mskDTFIN.Text) Then
        MsgBox "Data Inicial não pode ser maior que Data Final !!!", vbOKOnly + vbExclamation, "Aviso"
        mskDTINI.SetFocus
        Exit Function
    End If
    
    ConfereCampos = True

End Function

Private Sub ImprimeEntradas()
    
    Dim strNomRel       As String
    Dim strTipRel       As String
    Dim strCABEC1       As String
    Dim strCABEC2       As String
    Dim boolTemReg      As Boolean
    
    boolTemReg = False
    
    If optOrdemRel(0).value = True Then
        strNomRel = "RPTESTLITENTDATA.rpt"
        strTipRel = " / " & optOrdemRel(0).Caption
    ElseIf optOrdemRel(1).value = True Then
        strNomRel = "RPTESTLITENTCAPAC.rpt"
        strTipRel = " / " & optOrdemRel(1).Caption
    ElseIf optOrdemRel(2).value = True Then
        strNomRel = "RPTESTLITENTCLIE.rpt"
        strTipRel = " / " & optOrdemRel(2).Caption
    End If
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "        SGI_CADENTROTLIT_IT.SGI_CODIGO" & vbCrLf
    sSql = sSql & "      , SGI_CADENTROTLIT_IT.SGI_CODOP" & vbCrLf
    sSql = sSql & "      , SGI_CADENTROTLIT_IT.SGI_CODPED" & vbCrLf
    sSql = sSql & "      , SGI_CADENTROTLIT_IT.SGI_CODCAPAC" & vbCrLf
    sSql = sSql & "      , SGI_CADENTROTLIT_IT.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "      , SGI_CADENTROTLIT_IT.SGI_PRODUTO" & vbCrLf
    sSql = sSql & "      , SGI_CADENTROTLIT_IT.SGI_EXPESS" & vbCrLf
    sSql = sSql & "      , SGI_CADENTROTLIT_IT.SGI_LARGUR" & vbCrLf
    sSql = sSql & "      , SGI_CADENTROTLIT_IT.SGI_COMPRI" & vbCrLf
    sSql = sSql & "      , SGI_CADENTROTLIT_IT.SGI_QTDECORP" & vbCrLf
    sSql = sSql & "      , SGI_CADENTROTLIT_IT.SGI_QTDEFOLHAS" & vbCrLf
    sSql = sSql & "      , SGI_CADENTROTLIT_IT.SGI_PESO" & vbCrLf
    sSql = sSql & "      , SGI_CADENTROTLIT_IT.SGI_QTDELATAS" & vbCrLf
    sSql = sSql & "      , SGI_CADENTROTLIT_IT.SGI_QTDEFARDOS" & vbCrLf
    
    sSql = sSql & "      , SGI_CADENTROTLIT.SGI_DTENTRADA" & vbCrLf
    sSql = sSql & "      , SGI_CADENTROTLIT.SGI_EMPRESA" & vbCrLf
    sSql = sSql & "      , SGI_CADENTROTLIT.SGI_CODCLIE" & vbCrLf
    
    sSql = sSql & "      , SGI_CADCLIENTE.SGI_RAZAOSOC" & vbCrLf
    
    sSql = sSql & "      , SGI_CADLINHAPRODUTO.SGI_DESCRI" & vbCrLf
    
    sSql = sSql & "      , SGI_CADPRODUTO.SGI_DESCRICAO" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "        SGI_CADCLIENTE      SGI_CADCLIENTE" & vbCrLf
    sSql = sSql & "      , SGI_CADENTROTLIT    SGI_CADENTROTLIT" & vbCrLf
    sSql = sSql & "      , SGI_CADENTROTLIT_IT SGI_CADENTROTLIT_IT" & vbCrLf
    sSql = sSql & "      , SGI_CADLINHAPRODUTO SGI_CADLINHAPRODUTO" & vbCrLf
    sSql = sSql & "      , SGI_CADPRODUTO      SGI_CADPRODUTO" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "        SGI_CADENTROTLIT_IT.SGI_FILIAL      = " & FILIAL & vbCrLf
    
    sSql = sSql & "    And SGI_CADENTROTLIT_IT.SGI_FILIAL      = SGI_CADENTROTLIT.SGI_FILIAL" & vbCrLf
    sSql = sSql & "    And SGI_CADENTROTLIT_IT.SGI_CODIGO      = SGI_CADENTROTLIT.SGI_CODIGO" & vbCrLf
    
    sSql = sSql & "    And SGI_CADENTROTLIT_IT.SGI_FILIAL      = SGI_CADLINHAPRODUTO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "    And SGI_CADENTROTLIT_IT.SGI_CODCAPAC    = SGI_CADLINHAPRODUTO.SGI_CODLIN" & vbCrLf
    
    sSql = sSql & "    And SGI_CADENTROTLIT.SGI_FILIAL         = SGI_CADCLIENTE.SGI_FILIAL" & vbCrLf
    sSql = sSql & "    And SGI_CADENTROTLIT.SGI_CODCLIE        = SGI_CADCLIENTE.SGI_CODIGO" & vbCrLf
    
    sSql = sSql & "    And SGI_CADENTROTLIT_IT.SGI_FILIAL      = SGI_CADPRODUTO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "    And SGI_CADENTROTLIT_IT.SGI_IDPRODUTO   = SGI_CADPRODUTO.SGI_IDPRODUTO" & vbCrLf
    
    If optEmpresa(0).value = True Then sSql = sSql & "    And SGI_CADENTROTLIT.SGI_EMPRESA = 0" & vbCrLf
    If optEmpresa(1).value = True Then sSql = sSql & "    And SGI_CADENTROTLIT.SGI_EMPRESA = 1" & vbCrLf
    
    If Len(Trim(txtCIDCLIE.Text)) > 0 Then
        sSql = sSql & "    And SGI_CADENTROTLIT.SGI_CODCLIE        = " & Trim(txtCIDCLIE.Text) & vbCrLf
    End If
    If Len(Trim(txtCodLin.Text)) > 0 Then
        sSql = sSql & "    And SGI_CADENTROTLIT_IT.SGI_CODCAPAC    = " & Trim(txtCodLin.Text) & vbCrLf
    End If
    
    sSql = sSql & "Order By" & vbCrLf
    If optOrdemRel(0).value = True Then sSql = sSql & "        SGI_CADENTROTLIT.SGI_DTENTRADA"
    If optOrdemRel(1).value = True Then sSql = sSql & "        SGI_CADLINHAPRODUTO.SGI_DESCRI"
    If optOrdemRel(2).value = True Then sSql = sSql & "        SGI_CADCLIENTE.SGI_RAZAOSOC"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then boolTemReg = True
    BREC.Close
    
    If boolTemReg = False Then
        MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    If optEntrSaid(0).value = True Then strCABEC1 = "Relatório de Entrada de Litografia" & strTipRel
    If optEntrSaid(1).value = True Then strCABEC1 = "Relatório de Saidas de Litografia" & strTipRel
    
    strCABEC2 = "No Periodo de " & mskDTINI.Text & " a " & mskDTFIN.Text
    
    If optFormaGeracao(0).value = True Then
        If Len(Trim(strNomRel)) > 0 Then Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelEstoque & Trim(strNomRel), Linha, 1, strCABEC1, strCABEC2, True)
    Else
        If ExcelExists = 0 Then
            MsgBox "ATENÇÃO - O Excel não esta instalado na sua maquina !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
        End If
        Call ExportaParaExcel(sSql)
        
    End If
    
End Sub

Private Sub txtCIDCLIE_GotFocus()
    objBLBFunc.SelecionaCampos txtCIDCLIE.Name, Me
End Sub

Private Sub LimpaCamposLabel()
    lblDescCliente.Caption = ""
    lblDescLinha.Caption = ""
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
    End If

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

Private Sub txtCodLin_GotFocus()
    objBLBFunc.SelecionaCampos txtCodLin.Name, Me
End Sub

Private Sub txtCodLin_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCodLin.Text
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

Private Function ExcelExists() As Integer
    On Error Resume Next
    
    Dim objExcel As Object
    Dim intExists As Integer
    
    'try to create a new instance of MS Excel
    Set objExcel = CreateObject("Excel.Application")
    
    'if the instance of MS Excel does not exist then MS Excel is not installed
    If objExcel Is Nothing Then
        intExists = 0
        
    'else, MS Excel is installed
    Else
        intExists = 1
    End If
    
    'distroy the object
    Set objExcel = Nothing
    
    'return the status of MS Excel being installed
    ExcelExists = intExists
End Function


Private Sub ExportaParaExcel(strQuery As String)

On Error GoTo Handle_Error

    Dim myExcelFile     As New clsExcelFile
    Dim FileName$
    Dim boolTemDados    As Boolean
    Dim arrDADOSTAB()   As String
    Dim lngREGS         As Long
    Dim lngLINHA        As Long

    boolTemDados = False

    BREC.Open strQuery, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        boolTemDados = True
        lngREGS = 0
        Do While Not BREC.EOF()
            lngREGS = (lngREGS + 1)
            BREC.MoveNext
        Loop
    
        ReDim arrDADOSTAB(1 To lngREGS, 1 To 15) As String
        BREC.MoveFirst
        lngREGS = 0
    
        Do While Not BREC.EOF()
            lngREGS = (lngREGS + 1)
            arrDADOSTAB(lngREGS, 1) = Format(BREC!SGI_DTENTRADA, "DD/MM/YYYY")
            arrDADOSTAB(lngREGS, 2) = Trim(BREC!SGI_RAZAOSOC)
            arrDADOSTAB(lngREGS, 3) = BREC!SGI_CODOP
            arrDADOSTAB(lngREGS, 4) = BREC!SGI_CODPED
            arrDADOSTAB(lngREGS, 5) = Trim(BREC!SGI_PRODUTO)
            arrDADOSTAB(lngREGS, 6) = Trim(BREC!SGI_DESCRI)
            arrDADOSTAB(lngREGS, 7) = Trim(BREC!SGI_DESCRICAO)
            arrDADOSTAB(lngREGS, 8) = Format(BREC!SGI_EXPESS, "#,##0.00")
            arrDADOSTAB(lngREGS, 9) = BREC!SGI_LARGUR
            arrDADOSTAB(lngREGS, 10) = BREC!SGI_COMPRI
            arrDADOSTAB(lngREGS, 11) = BREC!SGI_QTDECORP
            arrDADOSTAB(lngREGS, 12) = BREC!SGI_QTDEFOLHAS
            arrDADOSTAB(lngREGS, 13) = BREC!SGI_PESO
            arrDADOSTAB(lngREGS, 14) = BREC!SGI_QTDELATAS
            arrDADOSTAB(lngREGS, 15) = BREC!SGI_QTDEFARDOS
            
            BREC.MoveNext
        Loop
    End If
    BREC.Close
    
    If boolTemDados = False Then
        MsgBox "Atenção - Não há dados para gerar o arquivo !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If


    With myExcelFile
        'Create the new spreadsheet
        FileName$ = strCamRelNovo & "RELPREPARA\RELMOVESTLIT.xls"
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
        .SetColumnWidth 2, 2, 25
        .SetColumnWidth 3, 5, 18
        .SetColumnWidth 6, 6, 25
        .SetColumnWidth 7, 7, 70
        .SetColumnWidth 8, 15, 12
        
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
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 1, "DATA", 12
        .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, 1, 2, "SETOR", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 3, "OP", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 4, "PEDIDO", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 5, "CÓDIGO", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 6, "CAPACIDADE", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 7, "LITOGRAFIA AGUARDANDO 2ª ORDEM", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 8, "Espessura", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 9, "Largura", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 10, "Comprimento", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 11, "Qtde.Corpos", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 12, "Qtde.Folhas", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 13, "Peso KG", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 14, "Qtde.Latas", 12
        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, 1, 15, "Qtde.Fardos", 12
        
        '' Jogando os Dados na Planilha
        lngLINHA = 2
        For lngREGS = 1 To UBound(arrDADOSTAB)
            
            .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 1, arrDADOSTAB(lngREGS, 1), 12
            .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 2, Trim(arrDADOSTAB(lngREGS, 2)), 12
            .WriteValue xlsnumber, xlsFont0, xlsrightAlign, xlsNormal, lngLINHA, 3, CLng(arrDADOSTAB(lngREGS, 3)), 1
            .WriteValue xlsnumber, xlsFont0, xlsrightAlign, xlsNormal, lngLINHA, 4, CLng(arrDADOSTAB(lngREGS, 4)), 1
            .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 5, Trim(arrDADOSTAB(lngREGS, 5)), 12
            .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 6, Trim(arrDADOSTAB(lngREGS, 6)), 12
            .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, lngLINHA, 7, Trim(arrDADOSTAB(lngREGS, 7)), 12
            .WriteValue xlsnumber, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 8, CCur(arrDADOSTAB(lngREGS, 8)), 2
            .WriteValue xlsnumber, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 9, CLng(arrDADOSTAB(lngREGS, 9)), 1
            .WriteValue xlsnumber, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 10, CLng(arrDADOSTAB(lngREGS, 10)), 1
            .WriteValue xlsnumber, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 11, CLng(arrDADOSTAB(lngREGS, 11)), 1
            .WriteValue xlsnumber, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 12, CLng(arrDADOSTAB(lngREGS, 12)), 1
            .WriteValue xlsnumber, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 13, CLng(arrDADOSTAB(lngREGS, 13)), 1
            .WriteValue xlsnumber, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 14, CLng(arrDADOSTAB(lngREGS, 14)), 1
            .WriteValue xlsnumber, xlsFont0, xlsCentreAlign, xlsNormal, lngLINHA, 15, CLng(arrDADOSTAB(lngREGS, 15)), 1
            
            lngLINHA = (lngLINHA + 1)
        Next lngREGS
        
        
        'PROTECT the spreadsheet so any cells specified as LOCKED will not be
        'overwritten. Also, all cells with HIDDEN set will hide their formula.
        'PROTECT does not use a password.
        .ProtectSpreadsheet = False 'False | True
        
        'Finally, close the spreadsheet
        .CloseFile
        
        MsgBox "Arquivo Excel : RELMOVESTLIT.xls foi Criado !", vbInformation + vbOKOnly, "Aviso do Sistema"
    End With
    
    Exit Sub
    
Handle_Error:
''    Debug.Print "Número: " & Err.Number & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Linha: " & Erl & vbCrLf

        
End Sub
