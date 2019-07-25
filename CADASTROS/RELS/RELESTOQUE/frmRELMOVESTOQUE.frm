VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRELMOVESTOQUE 
   Caption         =   "Movimentação de Estoque"
   ClientHeight    =   3570
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9060
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   9060
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame7 
      Caption         =   "[ Código do Lote ]"
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
      Left            =   120
      TabIndex        =   23
      Top             =   2880
      Width           =   4335
      Begin VB.TextBox txtCODLOTE 
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Text            =   "txtCODLOTE"
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame Frame6 
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
      Left            =   120
      TabIndex        =   19
      Top             =   2160
      Width           =   8895
      Begin VB.TextBox txtCIDCLIE 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   45
         MaxLength       =   10
         TabIndex        =   21
         Text            =   "txtCIDCLIE"
         Top             =   255
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   1245
         Picture         =   "frmRELMOVESTOQUE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblDescCliente 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescCliente"
         Height          =   285
         Left            =   1605
         TabIndex        =   22
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.Frame Frame5 
      Height          =   615
      Left            =   6240
      TabIndex        =   16
      Top             =   1560
      Width           =   2775
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
         Left            =   1560
         TabIndex        =   18
         Top             =   240
         Width           =   1095
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
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "[ Agrupamento ]"
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
      Left            =   120
      TabIndex        =   15
      Top             =   1560
      Width           =   3495
      Begin VB.OptionButton optAgrup 
         Caption         =   "Dia"
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
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optAgrup 
         Caption         =   "Semana"
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
         Left            =   840
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optAgrup 
         Caption         =   "Mês"
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
         Left            =   1920
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optAgrup 
         Caption         =   "Ano"
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
         Left            =   2640
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "[ Tipo ]"
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
      Left            =   3720
      TabIndex        =   14
      Top             =   1560
      Width           =   2415
      Begin VB.OptionButton optTipo 
         Caption         =   "Sintético"
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
         Left            =   1200
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Análitico"
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
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   960
      Width           =   9015
      Begin MSMask.MaskEdBox mskDTFIN 
         Height          =   285
         Left            =   4560
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
      Begin MSMask.MaskEdBox mskDTINI 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
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
         Left            =   3480
         TabIndex        =   13
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
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9015
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
         Picture         =   "frmRELMOVESTOQUE.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Picture         =   "frmRELMOVESTOQUE.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Exclui Empresa"
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmRELMOVESTOQUE"
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

Dim objBLBFunc          As Object
Dim objRELMOVESTOQUE    As Object
Dim objPESQPADRAO       As Object
Dim objREL              As Object

Private Sub cmdImpressao_Click()
    If ConfereCampos = False Then Exit Sub
    Call Imprime
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Command2_Click()

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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Clientes")
    
    If Len(Trim(varRETORNO)) > 0 Then
        txtCIDCLIE.Text = varRETORNO
        Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", varRETORNO, lblDescCliente)
    End If
    txtCIDCLIE.SetFocus

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objRELMOVESTOQUE = CreateObject("RELESTOQUE.clsRELMOVESTOQUE")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objRELMOVESTOQUE.FILIAL = FILIAL
    objBLBFunc.LimpaCampos frmRELMOVESTOQUE
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    
    optAgrup(0).Value = True
    optTipo(0).Value = True
    optEntrSaid(0).Value = True
    Call LimpaCamposLabel
    
    mskDTINI.Text = Format(Now, "DD/MM/YYYY")
    mskDTFIN.Text = Format(Now, "DD/MM/YYYY")
    
End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objRELMOVESTOQUE = Nothing
    Set objPESQPADRAO = Nothing
    Set objREL = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

Private Sub mskDTFIN_GotFocus()
    objBLBFunc.SelecionaCampos mskDTFIN.Name, frmRELMOVESTOQUE
End Sub

Private Sub mskDTINI_GotFocus()
    objBLBFunc.SelecionaCampos mskDTINI.Name, frmRELMOVESTOQUE
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


Private Sub Imprime()
    
    Dim strNomRel       As String
    Dim strTipRel       As String
    Dim strCABEC1       As String
    Dim strCABEC2       As String
    
    Dim strNOMTAB       As String
    
    If optEntrSaid(0).Value = True Then strNOMTAB = "SGI_CADITREQENTRMAT"
    If optEntrSaid(1).Value = True Then strNOMTAB = "SGI_CADITREQSAIMAT"
    
    sSql = ""
    
    sSql = "Select " & vbCrLf

    sSql = sSql & "       " & strNOMTAB & ".SGI_DATREQ" & vbCrLf
    sSql = sSql & "      ," & strNOMTAB & ".SGI_PRODUTO" & vbCrLf
    sSql = sSql & "      ," & strNOMTAB & ".SGI_QTD" & vbCrLf
    sSql = sSql & "      ," & strNOMTAB & ".SGI_QTDEKG" & vbCrLf
    sSql = sSql & "      ," & strNOMTAB & ".SGI_CODLOTE" & vbCrLf
    sSql = sSql & "      ," & strNOMTAB & ".SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_DESCRICAO" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & strNOMTAB & " " & strNOMTAB & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO SGI_CADPRODUTO" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    
    sSql = sSql & "       " & strNOMTAB & ".SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And " & strNOMTAB & ".SGI_DATREQ between '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'"
    
    sSql = sSql & "   And " & strNOMTAB & ".SGI_FILIAL    = SGI_CADPRODUTO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And " & strNOMTAB & ".SGI_IDPRODUTO = SGI_CADPRODUTO.SGI_IDPRODUTO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF() Then
        BREC.Close
        MsgBox "Não há dados para Imprimr !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    BREC.Close
    
    If optTipo(0).Value = True Then strTipRel = "Analitico"
    If optTipo(1).Value = True Then strTipRel = "Sintético"
    
    If optEntrSaid(0).Value = True Then strCABEC1 = "Relatório de Entrada de Produtos [ " & Trim(strTipRel) & " ]"
    If optEntrSaid(1).Value = True Then strCABEC1 = "Relatório de Saidas de Produtos [ " & Trim(strTipRel) & " ]"
    
    strCABEC2 = "No Periodo de " & mskDTINI.Text & " a " & mskDTFIN.Text
    
    If optEntrSaid(0).Value = True Then '' Entradas
        If optTipo(0).Value = True Then
            If optAgrup(0).Value = True Then strNomRel = "REMOVEST01A.rpt"
            If optAgrup(1).Value = True Then strNomRel = "REMOVEST02A.rpt"
            If optAgrup(2).Value = True Then strNomRel = "REMOVEST03A.rpt"
            If optAgrup(3).Value = True Then strNomRel = "REMOVEST04A.rpt"
        ElseIf optTipo(1).Value = True Then
            If optAgrup(0).Value = True Then strNomRel = "REMOVEST01S.rpt"
            If optAgrup(1).Value = True Then strNomRel = "REMOVEST02S.rpt"
            If optAgrup(2).Value = True Then strNomRel = "REMOVEST03S.rpt"
            If optAgrup(3).Value = True Then strNomRel = "REMOVEST04S.rpt"
        End If
    ElseIf optEntrSaid(0).Value = True Then '' Saidas
        If optTipo(0).Value = True Then
            If optAgrup(0).Value = True Then strNomRel = "REMOVESTS01A.rpt"
            If optAgrup(1).Value = True Then strNomRel = "REMOVESTS02A.rpt"
            If optAgrup(2).Value = True Then strNomRel = "REMOVESTS03A.rpt"
            If optAgrup(3).Value = True Then strNomRel = "REMOVESTS04A.rpt"
        ElseIf optTipo(1).Value = True Then
            If optAgrup(0).Value = True Then strNomRel = "REMOVESTS01S.rpt"
            If optAgrup(1).Value = True Then strNomRel = "REMOVESTS02S.rpt"
            If optAgrup(2).Value = True Then strNomRel = "REMOVESTS03S.rpt"
            If optAgrup(3).Value = True Then strNomRel = "REMOVESTS04S.rpt"
        End If
    End If
        
    If Len(Trim(strNomRel)) > 0 Then
            Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelPCP2 & Trim(strNomRel), Linha, 1, strCABEC1, strCABEC2, True)
    End If

End Sub


Private Sub LimpaCamposLabel()
    lblDescCliente.Caption = ""
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


Private Sub txtCIDCLIE_GotFocus()
    objBLBFunc.SelecionaCampos txtCIDCLIE.Name, frmRELMOVESTOQUE
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

Private Sub txtCODLOTE_GotFocus()
    objBLBFunc.SelecionaCampos txtCODLOTE.Name, frmRELMOVESTOQUE
End Sub

Private Sub txtCODLOTE_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub
