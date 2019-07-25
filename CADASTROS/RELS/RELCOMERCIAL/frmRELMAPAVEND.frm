VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRELMAPAVEND 
   Caption         =   "Mapa de Vendas por Vendedores"
   ClientHeight    =   2430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   12405
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
      TabIndex        =   20
      Top             =   1560
      Width           =   3015
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
         TabIndex        =   22
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
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command2 
      Height          =   315
      Left            =   2400
      Picture         =   "frmRELMAPAVEND.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox txtCODVEND 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   3
      Text            =   "txtCODVEND"
      Top             =   1815
      Width           =   975
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
      TabIndex        =   14
      Top             =   960
      Width           =   3015
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
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
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
         Left            =   1800
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
         Left            =   2760
         TabIndex        =   13
         Top             =   240
         Width           =   855
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
         TabIndex        =   12
         Top             =   240
         Width           =   1215
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
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
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
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   6
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
         TabIndex        =   8
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
         Picture         =   "frmRELMAPAVEND.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   0
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
         Picture         =   "frmRELMAPAVEND.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Exclui Empresa"
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Vendedor"
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
      TabIndex        =   19
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblDescVendedor 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblDescVendedor"
      Height          =   285
      Left            =   2760
      TabIndex        =   18
      Top             =   1800
      Width           =   6495
   End
End
Attribute VB_Name = "frmRELMAPAVEND"
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
Dim objRELMAPAVEND  As Object
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

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Command2_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "        SGI_CODIGO    " & vbCrLf
    sSql = sSql & "       ,SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "        SGI_CADVENDEDOR " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL     = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Venderores", "CADVENDEDOR.clsCADVENDEDOR")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODVEND.Text = varRETORNO
    
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADVENDEDOR", varRETORNO, lblDescVendedor)
    If Len(Trim(lblDescVendedor.Caption)) = 0 Then txtCODVEND.Text = ""
    txtCODVEND.SetFocus

End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objRELMAPAVEND = CreateObject("RELCOMERCIAL.clsRELMAPAVEND")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objBLBFunc.LimpaCampos frmRELMAPAVEND
    
    objRELMAPAVEND.FILIAL = FILIAL

    optAgrup(0).value = True
    optTIPO(0).value = True
    
    mskDTINI.Text = Format(Now, "DD/MM/YYYY")
    mskDTFIN.Text = Format((Now + 30), "DD/MM/YYYY")
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

    lblDescVendedor.Caption = ""
    optEmpresa(0).value = True

    Me.Caption = Me.Caption & " / " & Me.Name

End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objRELMAPAVEND = Nothing
    Set objPESQPADRAO = Nothing
    Set objREL = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
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

    
    Dim strEMPRESA As String
    
    strEMPRESA = ""
    If optEmpresa(1).value = True Then strEMPRESA = "_STEEL"
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "          SGI_CADPEDVENDI" & strEMPRESA & ".SGI_CODIGO" & vbCrLf
    sSql = sSql & "         ,SGI_CADPEDVENDI" & strEMPRESA & ".SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "         ,SGI_CADPEDVENDI" & strEMPRESA & ".SGI_CODPROD" & vbCrLf
    sSql = sSql & "         ,SGI_CADPEDVENDI" & strEMPRESA & ".SGI_QTDE" & vbCrLf
    sSql = sSql & "         ,SGI_CADPEDVENDI" & strEMPRESA & ".SGI_VLUNIT" & vbCrLf
    
    sSql = sSql & "         ,SGI_CADPEDVENDH" & strEMPRESA & ".SGI_DATAPED" & vbCrLf
    sSql = sSql & "         ,SGI_CADPEDVENDH" & strEMPRESA & ".SGI_CODVEND" & vbCrLf
    sSql = sSql & "         ,SGI_CADPRODUTO.SGI_CODLINPROD" & vbCrLf
    sSql = sSql & "         ,SGI_CADPRODUTO.SGI_DESCRICAO" & vbCrLf
    sSql = sSql & "         ,SGI_CADLINHAPRODUTO.SGI_DESCRI" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "         SGI_CADLINHAPRODUTO SGI_CADLINHAPRODUTO" & vbCrLf
    sSql = sSql & "        ,SGI_CADPEDVENDH" & strEMPRESA & " SGI_CADPEDVENDH" & strEMPRESA & vbCrLf
    sSql = sSql & "        ,SGI_CADPEDVENDI" & strEMPRESA & " SGI_CADPEDVENDI" & strEMPRESA & vbCrLf
    sSql = sSql & "        ,SGI_CADPRODUTO SGI_CADPRODUTO" & vbCrLf
    sSql = sSql & "        ,SGI_CADVENDEDOR SGI_CADVENDEDOR" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "        SGI_CADPEDVENDI" & strEMPRESA & ".SGI_FILIAL = " & FILIAL & vbCrLf
    
    sSql = sSql & "   And  SGI_CADPEDVENDI" & strEMPRESA & ".SGI_FILIAL = SGI_CADPEDVENDH" & strEMPRESA & ".SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And  SGI_CADPEDVENDI" & strEMPRESA & ".SGI_CODIGO = SGI_CADPEDVENDH" & strEMPRESA & ".SGI_CODIGO " & vbCrLf
    
    sSql = sSql & "   And  SGI_CADPEDVENDI" & strEMPRESA & ".SGI_FILIAL = SGI_CADPRODUTO.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And  SGI_CADPEDVENDI" & strEMPRESA & ".SGI_IDPRODUTO = SGI_CADPRODUTO.SGI_IDPRODUTO " & vbCrLf
    
    If optEmpresa(0).value = True Then
        strEMPRESADESC = "NOVALATA"
    ElseIf optEmpresa(1).value = True Then
        strEMPRESADESC = "STEEL ROL"
    End If
    
    sSql = sSql & "   And  SGI_CADPEDVENDH" & strEMPRESA & ".SGI_FILIAL = SGI_CADVENDEDOR.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And  SGI_CADPEDVENDH" & strEMPRESA & ".SGI_CODVEND = SGI_CADVENDEDOR.SGI_CODIGO " & vbCrLf
    
    If Len(Trim(txtCODVEND.Text)) > 0 Then
       sSql = sSql & "   And  SGI_CADPEDVENDH" & strEMPRESA & ".SGI_CODVEND = " & Trim(txtCODVEND.Text) & vbCrLf
    End If
    
    sSql = sSql & "   And  SGI_CADPRODUTO.SGI_FILIAL = SGI_CADLINHAPRODUTO.SGI_FILIAL"
    sSql = sSql & "   And  SGI_CADPRODUTO.SGI_CODLINPROD = SGI_CADLINHAPRODUTO.SGI_CODLIN"
    
    sSql = sSql & "   And  SGI_CADPEDVENDH" & strEMPRESA & ".SGI_DATAPED BETWEEN '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'"
    
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    strCABEC1 = "Mapa de pedidos por vendedores / " & strEMPRESADESC
    
    If optTIPO(0).value = True Then strCABEC1 = strCABEC1 & " [ Análitico/"
    If optTIPO(1).value = True Then strCABEC1 = strCABEC1 & " [ Sintético/"
    
    If CDate(mskDTINI.Text) <> CDate(mskDTFIN.Text) Then strCABEC2 = "No Periodo de " & mskDTINI.Text & " a " & mskDTFIN.Text
    If CDate(mskDTINI.Text) = CDate(mskDTFIN.Text) Then strCABEC2 = "Na Data de " & mskDTINI.Text

    strNomRel = ""
    If optTIPO(0).value = True Then '' Relatórios Análiticos
        If optEmpresa(0).value = True Then '' Novalata
            If optAgrup(0).value = True Then strNomRel = "RELMAPAVENDREP01.rpt"  '' Agrupamento por Dia
            If optAgrup(1).value = True Then strNomRel = "RELMAPAVENDREP02.rpt"  '' Agrupamento por Semana
            If optAgrup(2).value = True Then strNomRel = "RELMAPAVENDREP03.rpt"  '' Agrupamento por Mês
            If optAgrup(3).value = True Then strNomRel = "RELMAPAVENDREP04.rpt"  '' Agrupamento por Ano
        ElseIf optEmpresa(1).value = True Then '' Steel
            If optAgrup(0).value = True Then strNomRel = "RELMAPAVENDREP01_STEEL.rpt"  '' Agrupamento por Dia
            If optAgrup(1).value = True Then strNomRel = "RELMAPAVENDREP02_STEEL.rpt"  '' Agrupamento por Semana
            If optAgrup(2).value = True Then strNomRel = "RELMAPAVENDREP03_STEEL.rpt"  '' Agrupamento por Mês
            If optAgrup(3).value = True Then strNomRel = "RELMAPAVENDREP04_STEEL.rpt"  '' Agrupamento por Ano
        End If
    End If

    If optAgrup(0).value = True Then strCABEC1 = strCABEC1 & "Dia ]"        '' Agrupamento por Dia
    If optAgrup(1).value = True Then strCABEC1 = strCABEC1 & "Semana ]"     '' Agrupamento por Semana
    If optAgrup(2).value = True Then strCABEC1 = strCABEC1 & "Mês ]"        '' Agrupamento por Mês
    If optAgrup(3).value = True Then strCABEC1 = strCABEC1 & "Ano ]"        '' Agrupamento por Ano

    If Len(Trim(strNomRel)) > 0 Then
        Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelComercial & Trim(strNomRel), Linha, 1, strCABEC1, strCABEC2, True)
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub PegaDescTabelas(StrCampoPesq As String, StrCampoRetorno As String, strTabela As String, strCODIGO As String, lblLabel As Label)

    lblLabel.Caption = ""
    
    If Len(Trim(strCODIGO)) = 0 Then Exit Sub
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoRetorno) & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & Trim(strTabela) & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoPesq) & " = " & Trim(strCODIGO)
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then
       lblLabel.Caption = Trim(BREC10(Trim(StrCampoRetorno)))
    Else
       MsgBox "Registro Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC10.Close
    
End Sub

Private Sub mskDTFIN_GotFocus()
    objBLBFunc.SelecionaCampos mskDTFIN.Name, frmRELMAPAVEND
End Sub

Private Sub mskDTINI_GotFocus()
    objBLBFunc.SelecionaCampos mskDTINI.Name, frmRELMAPAVEND
End Sub

Private Sub txtCODVEND_GotFocus()
    objBLBFunc.SelecionaCampos txtCODVEND.Name, frmRELMAPAVEND
End Sub

Private Sub txtCODVEND_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODVEND.Text
End Sub

Private Sub txtCODVEND_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODVEND.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODVEND.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODVEND.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADVENDEDOR", txtCODVEND.Text, lblDescVendedor)
    If Len(Trim(lblDescVendedor.Caption)) = 0 Then
       txtCODVEND.Text = ""
       Cancel = True
    End If

End Sub
