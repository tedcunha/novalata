VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmELPEDDTENTREGA 
   Caption         =   "Relatório de Pedidos por data de entrega"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7830
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   7830
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame7 
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
      TabIndex        =   23
      Top             =   2760
      Width           =   7815
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
      Begin VB.CommandButton cmdClie 
         Height          =   315
         Left            =   1080
         Picture         =   "frmELPEDDTENTREGA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblDesclie 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDesclie"
         Height          =   315
         Left            =   1440
         TabIndex        =   25
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.Frame Frame6 
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
      TabIndex        =   20
      Top             =   2160
      Width           =   7815
      Begin VB.TextBox txtLinProd 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         MaxLength       =   10
         TabIndex        =   10
         Text            =   "txtLinProd"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdLinProd 
         Height          =   315
         Left            =   1080
         Picture         =   "frmELPEDDTENTREGA.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblLinhProd 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblLinhProd"
         Height          =   315
         Left            =   1440
         TabIndex        =   22
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.Frame Frame5 
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
      Left            =   4680
      TabIndex        =   19
      Top             =   1560
      Width           =   3135
      Begin VB.OptionButton optEmpresa 
         Caption         =   "STEEL ROL"
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
         TabIndex        =   9
         Top             =   240
         Width           =   1455
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
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "[ Status do PEdido ]"
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
      Top             =   1560
      Width           =   4695
      Begin VB.OptionButton optStatus 
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
         Left            =   3600
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optStatus 
         Caption         =   "Parcial"
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
         Left            =   2640
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optStatus 
         Caption         =   "Fechadas"
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
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optStatus 
         Caption         =   "Em Aberto"
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
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
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
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   5400
      TabIndex        =   17
      Top             =   960
      Width           =   2415
      Begin VB.OptionButton optTipo 
         Caption         =   "Por Data de Entrega"
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
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   14
      Top             =   960
      Width           =   5415
      Begin MSMask.MaskEdBox mskDTFIN 
         Height          =   285
         Left            =   3840
         TabIndex        =   2
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
         Left            =   1320
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
         Left            =   2760
         TabIndex        =   16
         Top             =   240
         Width           =   1095
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
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
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
         Picture         =   "frmELPEDDTENTREGA.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdImpressao 
         Caption         =   "Im&prime"
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
         Left            =   960
         Picture         =   "frmELPEDDTENTREGA.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Exclui Empresa"
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmELPEDDTENTREGA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho      As String
Public Linha         As Variant
Public FILIAL        As Integer
Public strAcesso     As String
Public lngCodUsuario As Long
Dim objBLBFunc       As Object
Dim objRELPDDTENTR   As Object
Dim objPESQPADRAO    As Object
Dim objREL           As Object
Dim strCABEC1        As String
Dim strCABEC2        As String


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
    If ConfereCampos = False Then Exit Sub
    If optTipo(0).Value = True Then Call Imprime           '' Por data de entrega
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
    Set objRELPDDTENTR = CreateObject("RELPCP.clsRELPDDTENTR")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objBLBFunc.LimpaCampos frmELPEDDTENTREGA
    objRELPDDTENTR.FILIAL = FILIAL

    optStatus(0).Value = True
    optTipo(0).Value = True
    optEmpresa(0).Value = True
    
    mskDTINI.Text = Format(Now, "DD/MM/YYYY")
    mskDTFIN.Text = Format((Now + 7), "DD/MM/YYYY")

    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    Call LimpaLabels

End Sub

Private Sub DestroiObjeto()
    Set objBLBFunc = Nothing
    Set objRELPDDTENTR = Nothing
    Set objPESQPADRAO = Nothing
    Set objREL = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DestroiObjeto
End Sub

Private Sub Imprime()
    
    Dim strNomRel       As String
    Dim strTipRel       As String
    Dim strTabela       As String
    Dim strNOMEMPRESA   As String
    
    strTabela = ""
    If optEmpresa(1).Value = True Then strTabela = "_STEEL"
    
    
    If optEmpresa(0).Value = True Then strNOMEMPRESA = "NOVALATA"
    If optEmpresa(1).Value = True Then strNOMEMPRESA = "STEEL ROL"
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "        SGI_ORDEMPROD" & strTabela & ".SGI_DATENTREGA" & vbCrLf
    sSql = sSql & "       ,SGI_ORDEMPROD" & strTabela & ".SGI_CODIGO" & vbCrLf
    sSql = sSql & "       ,SGI_ORDEMPROD" & strTabela & ".SGI_QTDEPED" & vbCrLf
    sSql = sSql & "       ,SGI_ORDEMPROD" & strTabela & ".SGI_CODPROD" & vbCrLf
     
    sSql = sSql & "        SGI_CADPRODUTO.SGI_DESCRICAO" & vbCrLf
     
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "        SGI_CADPRODUTO SGI_CADPRODUTO"
    sSql = sSql & "       ,SGI_ORDEMPROD" & strTabela & " SGI_ORDEMPROD" & strTabela & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD" & strTabela & ".SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_ORDEMPROD" & strTabela & ".SGI_DATENTREGA Between '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'" & vbCrLf
    
    sSql = sSql & "   And SGI_ORDEMPROD" & strTabela & ".SGI_FILIAL = SGI_CADPRODUTO.SGI_FILIAL "
    
    strTipRel = "Todos"
    If optStatus(0).Value = True Then
        strTipRel = "Aberto"
        sSql = sSql & "   And SGI_ORDEMPROD" & strTabela & ".SGI_STATUS = 0" & vbCrLf
    ElseIf optStatus(1).Value = True Then
        strTipRel = "Fechado"
        sSql = sSql & "   And SGI_ORDEMPROD" & strTabela & ".SGI_STATUS = 2" & vbCrLf
    ElseIf optStatus(2).Value = True Then
        strTipRel = "Parcial"
        sSql = sSql & "   And SGI_ORDEMPROD" & strTabela & ".SGI_STATUS = 1" & vbCrLf
    End If
    
    If Len(Trim(txtLinProd.Text)) > 0 Then sSql = sSql & " And SGI_CADPRODUTO.SGI_CODLINPROD = " & Trim(txtLinProd.Text) & vbCrLf
    If Len(Trim(txtCodCliente.Text)) > 0 Then sSql = sSql & " And SGI_CADPRODUTO.SGI_CODCLIE = " & Trim(txtCodCliente.Text) & vbCrLf
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF() Then
        BREC.Close
        MsgBox "Não há dados para Imprimr !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    BREC.Close
    
    strCABEC1 = "Pedidos em Aberto " & strNOMEMPRESA & " por data de Entrega [ " & Trim(strTipRel) & " ]"
    strCABEC2 = "No Periodo de " & mskDTINI.Text & " a " & mskDTFIN.Text
    
    If optEmpresa(0).Value = True Then strNomRel = "RELPDDTENTR2.rpt"
    If optEmpresa(1).Value = True Then strNomRel = "RELPDDTENTR_STEEL02.rpt"

    If Len(Trim(strNomRel)) > 0 Then
        Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelPCP2 & Trim(strNomRel), Linha, 1, strCABEC1, strCABEC2, True)
    End If

End Sub


Private Function ConfereCampos() As Boolean
    
    ConfereCampos = False
        
        If Not IsDate(mskDTINI.Text) Then
            MsgBox "Data inicial inválida !!!", vbOKOnly + vbExclamation, "Aviso"
            mskDTINI.SetFocus
            Exit Function
        End If
        If Not IsDate(mskDTFIN.Text) Then
            MsgBox "Data final inválida !!!", vbOKOnly + vbExclamation, "Aviso"
            mskDTFIN.SetFocus
            Exit Function
        End If
        
        If CDate(mskDTINI.Text) > CDate(mskDTFIN.Text) Then
            MsgBox "Data inicial não pode ser maior que data final !!!", vbOKOnly + vbExclamation, "Aviso"
            mskDTINI.SetFocus
            Exit Function
        End If
    
    ConfereCampos = True

End Function


Private Sub txtCodCliente_GotFocus()
    objBLBFunc.SelecionaCampos txtCodCliente.Name, frmELPEDDTENTREGA
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
    objBLBFunc.SelecionaCampos txtLinProd.Name, frmELPEDDTENTREGA
End Sub

Private Sub txtLinProd_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtLinProd.Text
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

