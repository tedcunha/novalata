VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRELMAPACLIE 
   Caption         =   "Mapa de Vendas Por Clientes"
   ClientHeight    =   2205
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12495
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   12495
   StartUpPosition =   1  'CenterOwner
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
      Left            =   9360
      TabIndex        =   19
      Top             =   1560
      Width           =   3015
      Begin VB.OptionButton optEmpresa 
         Caption         =   "STELL ROL"
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
         TabIndex        =   21
         Top             =   240
         Width           =   1335
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
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command1 
      Height          =   315
      Left            =   1800
      Picture         =   "frmRELMAPACLIE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox txtCIDCLIE 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   840
      MaxLength       =   10
      TabIndex        =   3
      Text            =   "txtCIDCLIE"
      Top             =   1815
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   14
      Top             =   960
      Width           =   5415
      Begin MSMask.MaskEdBox mskDTFIN 
         Height          =   285
         Left            =   3960
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
      Left            =   5520
      TabIndex        =   9
      Top             =   960
      Width           =   3735
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   240
         Width           =   1215
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
         Left            =   2040
         TabIndex        =   11
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
         Left            =   2880
         TabIndex        =   10
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
      Left            =   9360
      TabIndex        =   6
      Top             =   960
      Width           =   3015
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
         Left            =   1680
         TabIndex        =   8
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
         Left            =   480
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   12375
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
         Picture         =   "frmRELMAPACLIE.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "frmRELMAPACLIE.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Label3 
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
      Left            =   120
      TabIndex        =   22
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label lblDescCliente 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblDescCliente"
      Height          =   285
      Left            =   2160
      TabIndex        =   18
      Top             =   1800
      Width           =   7095
   End
End
Attribute VB_Name = "frmRELMAPACLIE"
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
Dim objRELMAPACLIE  As Object
Dim objPESQPADRAO   As Object
Dim objREL          As Object
Dim strCABEC1       As String
Dim strCABEC2       As String
Dim strNomRel       As String
Dim strEMPRESADESC  As String

Private Sub cmdImpressao_Click()
    If ConfereCampos = False Then Exit Sub
    Call ImprimeMapaDiaDia
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
    Set objRELMAPACLIE = Nothing
    Set objPESQPADRAO = Nothing
    Set objREL = Nothing
End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objRELMAPACLIE = CreateObject("RELCOMERCIAL.clsRELMAPACLIE")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objBLBFunc.LimpaCampos frmRELMAPACLIE
    
    objRELMAPACLIE.FILIAL = FILIAL

    optAgrup(0).Value = True
    optTipo(0).Value = True
    
    mskDTINI.Text = Format(Now, "DD/MM/YYYY")
    mskDTFIN.Text = Format((Now + 30), "DD/MM/YYYY")
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    lblDescCliente.Caption = ""
    optEmpresa(0).Value = True

    Me.Caption = Me.Caption & " / " & Me.Name

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

Private Sub PegaDescTabelas(StrCampoPesq As String, StrCampoRetorno As String, strTABELA As String, StrCodigo As String, lblLabel As Label)

    lblLabel.Caption = ""
    
    If Len(Trim(StrCodigo)) = 0 Then Exit Sub
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoRetorno) & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & Trim(strTABELA) & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoPesq) & " = " & Trim(StrCodigo)
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then
       lblLabel.Caption = Trim(BREC10(Trim(StrCampoRetorno)))
    Else
       MsgBox "Registro Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC10.Close
    
End Sub

Private Sub mskDTFIN_GotFocus()
    objBLBFunc.SelecionaCampos mskDTFIN.Name, frmRELMAPACLIE
End Sub

Private Sub mskDTINI_GotFocus()
    objBLBFunc.SelecionaCampos mskDTINI.Name, frmRELMAPACLIE
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

Private Sub ImprimeMapaDiaDia()

    Dim strTABELA As String
    
    strTABELA = ""
    If optEmpresa(1).Value = True Then strTABELA = "_STEEL"
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "          SGI_CADPEDVENDI" & strTABELA & ".SGI_CODIGO" & vbCrLf
    sSql = sSql & "         ,SGI_CADPEDVENDI" & strTABELA & ".SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "         ,SGI_CADPEDVENDI" & strTABELA & ".SGI_CODPROD" & vbCrLf
    sSql = sSql & "         ,SGI_CADPEDVENDI" & strTABELA & ".SGI_QTDE" & vbCrLf
    sSql = sSql & "         ,SGI_CADPEDVENDI" & strTABELA & ".SGI_VLUNIT" & vbCrLf
    
    sSql = sSql & "         ,SGI_CADPEDVENDH" & strTABELA & ".SGI_DATAPED" & vbCrLf
    sSql = sSql & "         ,SGI_CADPEDVENDH" & strTABELA & ".SGI_CODCLI" & vbCrLf
    sSql = sSql & "         ,SGI_CADPRODUTO.SGI_CODLINPROD" & vbCrLf
    sSql = sSql & "         ,SGI_CADPRODUTO.SGI_DESCRICAO" & vbCrLf
    sSql = sSql & "         ,SGI_CADLINHAPRODUTO.SGI_DESCRI" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "         SGI_CADCLIENTE SGI_CADCLIENTE " & vbCrLf
    sSql = sSql & "        ,SGI_CADLINHAPRODUTO SGI_CADLINHAPRODUTO" & vbCrLf
    sSql = sSql & "        ,SGI_CADPEDVENDH" & strTABELA & " SGI_CADPEDVENDH" & strTABELA & vbCrLf
    sSql = sSql & "        ,SGI_CADPEDVENDI" & strTABELA & " SGI_CADPEDVENDI" & strTABELA & vbCrLf
    sSql = sSql & "        ,SGI_CADPRODUTO SGI_CADPRODUTO" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "        SGI_CADPEDVENDI" & strTABELA & ".SGI_FILIAL = " & FILIAL & vbCrLf
    
    If optEmpresa(0).Value = True Then
        strEMPRESADESC = "NOVALATA"
    ElseIf optEmpresa(1).Value = True Then
        strEMPRESADESC = "STEEL ROL"
    End If
    
    sSql = sSql & "   And  SGI_CADPEDVENDI" & strTABELA & ".SGI_FILIAL = SGI_CADPEDVENDH" & strTABELA & ".SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And  SGI_CADPEDVENDI" & strTABELA & ".SGI_CODIGO = SGI_CADPEDVENDH" & strTABELA & ".SGI_CODIGO " & vbCrLf
    
    sSql = sSql & "   And  SGI_CADPEDVENDI" & strTABELA & ".SGI_FILIAL = SGI_CADPRODUTO.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And  SGI_CADPEDVENDI" & strTABELA & ".SGI_IDPRODUTO = SGI_CADPRODUTO.SGI_IDPRODUTO " & vbCrLf
    
    sSql = sSql & "   And  SGI_CADPRODUTO.SGI_FILIAL = SGI_CADLINHAPRODUTO.SGI_FILIAL"
    sSql = sSql & "   And  SGI_CADPRODUTO.SGI_CODLINPROD = SGI_CADLINHAPRODUTO.SGI_CODLIN"
    
    sSql = sSql & "   And  SGI_CADPEDVENDH" & strTABELA & ".SGI_FILIAL = SGI_CADCLIENTE.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And  SGI_CADPEDVENDH" & strTABELA & ".SGI_CODCLI = SGI_CADCLIENTE.SGI_CODIGO" & vbCrLf
    
    sSql = sSql & "   And  SGI_CADPEDVENDH" & strTABELA & ".SGI_DATAPED BETWEEN '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'"
    
    If Len(Trim(txtCIDCLIE.Text)) > 0 Then
       sSql = sSql & "   And  SGI_CADPEDVENDH" & strTABELA & ".SGI_CODCLI = " & Trim(txtCIDCLIE.Text) & vbCrLf
    End If
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    strCABEC1 = "Mapa de pedidos por Clientes / " & strEMPRESADESC
    
    If optTipo(0).Value = True Then strCABEC1 = strCABEC1 & " [ Análitico/"
    If optTipo(1).Value = True Then strCABEC1 = strCABEC1 & " [ Sintético/"
    
    If CDate(mskDTINI.Text) <> CDate(mskDTFIN.Text) Then strCABEC2 = "No Periodo de " & mskDTINI.Text & " a " & mskDTFIN.Text
    If CDate(mskDTINI.Text) = CDate(mskDTFIN.Text) Then strCABEC2 = "Na Data de " & mskDTINI.Text

    strNomRel = ""
    If optTipo(0).Value = True Then '' Relatórios Análiticos
        If optEmpresa(0).Value = True Then
            If optAgrup(0).Value = True Then strNomRel = "RELMAPAVENDCLIE01.rpt"  '' Agrupamento por Dia
            If optAgrup(1).Value = True Then strNomRel = "RELMAPAVENDCLIE02.rpt"  '' Agrupamento por Semana
            If optAgrup(2).Value = True Then strNomRel = "RELMAPAVENDCLIE03.rpt"  '' Agrupamento por Mês
            If optAgrup(3).Value = True Then strNomRel = "RELMAPAVENDCLIE04.rpt"  '' Agrupamento por Ano
        ElseIf optEmpresa(1).Value = True Then
            If optAgrup(0).Value = True Then strNomRel = "RELMAPAVENDCLIE01_STEEL.rpt"  '' Agrupamento por Dia
            If optAgrup(1).Value = True Then strNomRel = "RELMAPAVENDCLIE02_STEEL.rpt"  '' Agrupamento por Semana
            If optAgrup(2).Value = True Then strNomRel = "RELMAPAVENDCLIE03_STEEL.rpt"  '' Agrupamento por Mês
            If optAgrup(3).Value = True Then strNomRel = "RELMAPAVENDCLIE04_STEEL.rpt"  '' Agrupamento por Ano
        End If
    End If

    If optAgrup(0).Value = True Then strCABEC1 = strCABEC1 & "Dia ]"        '' Agrupamento por Dia
    If optAgrup(1).Value = True Then strCABEC1 = strCABEC1 & "Semana ]"     '' Agrupamento por Semana
    If optAgrup(2).Value = True Then strCABEC1 = strCABEC1 & "Mês ]"        '' Agrupamento por Mês
    If optAgrup(3).Value = True Then strCABEC1 = strCABEC1 & "Ano ]"        '' Agrupamento por Ano

    If Len(Trim(strNomRel)) > 0 Then
        Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelComercial & Trim(strNomRel), Linha, 1, strCABEC1, strCABEC2, True)
    End If

End Sub

Private Sub txtCIDCLIE_GotFocus()
    objBLBFunc.SelecionaCampos txtCIDCLIE.Name, frmRELMAPACLIE
End Sub

Private Sub txtCIDCLIE_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCIDCLIE.Text
End Sub

Private Sub txtCIDCLIE_Validate(Cancel As Boolean)

    Dim I As Integer
    
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
