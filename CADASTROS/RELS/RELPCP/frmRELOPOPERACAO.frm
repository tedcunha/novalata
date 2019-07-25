VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRELOPOPERACAO 
   Caption         =   "Relatório de OP por Setores"
   ClientHeight    =   1920
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13980
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   13980
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "[ Verniz / Esmalte ]"
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
      Left            =   8640
      TabIndex        =   14
      Top             =   960
      Width           =   5295
      Begin VB.OptionButton optVZES 
         Caption         =   "Esmalte"
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
         Left            =   3960
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optVZES 
         Caption         =   "Verniz Interno 02"
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
         Left            =   2040
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optVZES 
         Caption         =   "Verniz Interno 01"
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
         Width           =   1815
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
      Left            =   5400
      TabIndex        =   13
      Top             =   960
      Width           =   3135
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
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
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
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   13935
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
         Picture         =   "frmRELOPOPERACAO.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Picture         =   "frmRELOPOPERACAO.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Exclui Empresa"
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "[ Data Empresa ]"
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
      TabIndex        =   0
      Top             =   960
      Width           =   5295
      Begin MSMask.MaskEdBox mskDTFIN 
         Height          =   285
         Left            =   3840
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
         TabIndex        =   10
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
         Left            =   2760
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmRELOPOPERACAO"
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
Dim objRELOPOPERACAO    As Object
Dim objPESQPADRAO       As Object
Dim objREL              As Object
Dim strCABEC1           As String
Dim strCABEC2           As String

Private Sub cmdImpressao_Click()
    If ConfereCampos = False Then Exit Sub
    Call Imprime
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
    Set objRELOPOPERACAO = CreateObject("RELPCP.clsRELPDDTENTR")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)
    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    Call objBLBFunc.LimpaCampos(Me)
    objRELOPOPERACAO.FILIAL = FILIAL

    mskDTINI.Text = Format(Now, "DD/MM/YYYY")
    mskDTFIN.Text = Format((Now + 7), "DD/MM/YYYY")

    optEmpresa(1).Value = True
    optVZES(0).Value = True

    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objRELOPOPERACAO = Nothing
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

Private Sub Imprime()
    
    Dim strNomRel       As String
    Dim strTipRel       As String
    Dim strTabela       As String
    Dim strNOMEMPRESA   As String
    Dim strTABVERNIZ    As String
    
    strTabela = ""
    If optEmpresa(1).Value = True Then strTabela = "_STEEL"
    
    If optVZES(0).Value = True Then strTABVERNIZ = "SGI_VERNIZPROD"
    If optVZES(1).Value = True Then strTABVERNIZ = "SGI_VERNIZPROD2"
    
    If optEmpresa(0).Value = True Then strNOMEMPRESA = "NOVALATA"
    If optEmpresa(1).Value = True Then strNOMEMPRESA = "STEEL ROL"
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD" & strTabela & ".SGI_DATENTREGA" & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD" & strTabela & ".SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD" & strTabela & ".SGI_CODIGO" & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD" & strTabela & ".SGI_CODPROD" & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD" & strTabela & ".SGI_QTDE" & vbCrLf
    
    sSql = sSql & "      ," & strTABVERNIZ & ".SGI_PRODUTO " & vbCrLf
    
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_DESCRICAO" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_CODLINPROD" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_QTDCORPSPADRAOSN" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_QTDEPORFOLHA" & vbCrLf
    
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO.SGI_QTDECORPOS" & vbCrLf
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO.SGI_PERDPROC" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO SGI_CADLINHAPRODUTO" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO SGI_CADPRODUTO" & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD" & strTabela & " SGI_ORDEMPROD" & strTabela & vbCrLf
    sSql = sSql & "      ," & strTABVERNIZ & " " & strTABVERNIZ & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD" & strTabela & ".SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_ORDEMPROD" & strTabela & ".SGI_DATENTREGA Between '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'" & vbCrLf
    
    sSql = sSql & "   And SGI_ORDEMPROD" & strTabela & ".SGI_FILIAL    = " & strTABVERNIZ & ".SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And SGI_ORDEMPROD" & strTabela & ".SGI_IDPRODUTO = " & strTABVERNIZ & ".SGI_IDPRODUTO " & vbCrLf
    
    sSql = sSql & "   And SGI_ORDEMPROD" & strTabela & ".SGI_FILIAL    = SGI_CADPRODUTO.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And SGI_ORDEMPROD" & strTabela & ".SGI_IDPRODUTO = SGI_CADPRODUTO.SGI_IDPRODUTO " & vbCrLf
    
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIAL     = SGI_CADLINHAPRODUTO.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_CODLINPROD = SGI_CADLINHAPRODUTO.SGI_CODLIN " & vbCrLf
    
    strTipRel = "Todos"
    ''If optStatus(0).Value = True Then
        strTipRel = "Aberto"
        sSql = sSql & "   And SGI_ORDEMPROD" & strTabela & ".SGI_STATUS = 0" & vbCrLf
    ''ElseIf optStatus(1).Value = True Then
''        strTipRel = "Fechado"
''        sSql = sSql & "   And SGI_ORDEMPROD" & strTabela & ".SGI_STATUS = 2" & vbCrLf
    ''ElseIf optStatus(2).Value = True Then
''        strTipRel = "Parcial"
''        sSql = sSql & "   And SGI_ORDEMPROD" & strTabela & ".SGI_STATUS = 1" & vbCrLf
    ''End If
    
''    sSql = sSql & "   And SGI_ORDEMPROD" & strTabela & ".SGI_FILIAL = SGI_CADPRODUTO.SGI_FILIAL " & vbCrLf
''    sSql = sSql & "   And SGI_ORDEMPROD" & strTabela & ".SGI_IDPRODUTO = SGI_CADPRODUTO.SGI_IDPRODUTO " & vbCrLf
    
''    sSql = sSql & "   And SGI_ORDEMPROD" & strTabela & ".SGI_FILIAL = SGI_CADPEDVENDH" & strTabela & ".SGI_FILIAL " & vbCrLf
''    sSql = sSql & "   And SGI_ORDEMPROD" & strTabela & ".SGI_CODPED = SGI_CADPEDVENDH" & strTabela & ".SGI_CODIGO " & vbCrLf
    
''    sSql = sSql & "   And SGI_CADPEDVENDH" & strTabela & ".SGI_FILIAL = SGI_CADCLIENTE.SGI_FILIAL " & vbCrLf
''    sSql = sSql & "   And SGI_CADPEDVENDH" & strTabela & ".SGI_CODCLI = SGI_CADCLIENTE.SGI_CODIGO " & vbCrLf
    
''    sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIAL = SGI_CADLINHAPRODUTO.SGI_FILIAL " & vbCrLf
''    sSql = sSql & "   And SGI_CADPRODUTO.SGI_CODLINPROD = SGI_CADLINHAPRODUTO.SGI_CODLIN " & vbCrLf
    
''    If Len(Trim(txtLinProd.Text)) > 0 Then sSql = sSql & " And SGI_CADPRODUTO.SGI_CODLINPROD = " & Trim(txtLinProd.Text) & vbCrLf
''    If Len(Trim(txtCodCliente.Text)) > 0 Then sSql = sSql & " And SGI_CADPRODUTO.SGI_CODCLIE = " & Trim(txtCodCliente.Text) & vbCrLf
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF() Then
        BREC.Close
        MsgBox "Não há dados para Imprimr !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    BREC.Close
    
    strCABEC1 = "Ordens de Produção " & strNOMEMPRESA & " por data de Entrega [ " & Trim(strTipRel) & " ]"
    strCABEC2 = "No Periodo de " & mskDTINI.Text & " a " & mskDTFIN.Text
    
    If optEmpresa(0).Value = True Then strNomRel = ""
    If optEmpresa(1).Value = True Then strNomRel = "RELOPOPERACAO01A.rpt"

    If Len(Trim(strNomRel)) > 0 Then
        Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelPCP2 & Trim(strNomRel), Linha, 1, strCABEC1, strCABEC2, True)
    End If

End Sub

