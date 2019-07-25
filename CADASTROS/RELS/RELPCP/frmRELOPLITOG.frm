VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRELOPLITOG 
   Caption         =   "Relatório de Ordem de Produção para Litografia"
   ClientHeight    =   1830
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10770
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   10770
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   12
      Top             =   960
      Width           =   5055
      Begin MSMask.MaskEdBox mskDTFIN 
         Height          =   285
         Left            =   3720
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
         Left            =   2640
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   240
         Width           =   1095
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
      Left            =   5160
      TabIndex        =   8
      Top             =   960
      Width           =   2655
      Begin VB.OptionButton optFiltro 
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
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optFiltro 
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
         Index           =   1
         Left            =   960
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optFiltro 
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
         Index           =   2
         Left            =   1800
         TabIndex        =   9
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
      Left            =   7920
      TabIndex        =   5
      Top             =   960
      Width           =   2775
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
         TabIndex        =   7
         Top             =   240
         Width           =   1215
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
         Left            =   1440
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10695
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
         Picture         =   "frmRELOPLITOG.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "frmRELOPLITOG.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Exclui Empresa"
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmRELOPLITOG"
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
Dim objRELOPLITOG    As Object
Dim objPESQPADRAO    As Object
Dim objREL           As Object
Dim strCABEC1        As String
Dim strCABEC2        As String

Private Sub cmdImpressao_Click()
    If ConfereCampos = False Then Exit Sub
    Call Imprime
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
    Set objRELOPLITOG = CreateObject("RELPCP.clsRELOPLITOG")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objBLBFunc.LimpaCampos frmRELOPLITOG
    objRELOPLITOG.FILIAL = FILIAL

    mskDTINI.Text = Format(Now, "DD/MM/YYYY")
    mskDTFIN.Text = Format((Now + 7), "DD/MM/YYYY")

    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

    optFiltro(0).Value = True
    optTipo(0).Value = True

End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objRELOPLITOG = Nothing
    Set objPESQPADRAO = Nothing
    Set objREL = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
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
    Dim arrREGISTROS    As Variant
    Dim lngREGS         As Long
    Dim strValor        As String
    
    sSql = ""
    
    arrREGISTROS = Empty
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       ORDP.SGI_CODIGO AS SGI_CODOP " & vbCrLf
    sSql = sSql & "      ,ORDP.SGI_DATENTREGA " & vbCrLf
    sSql = sSql & "      ,ORDP.SGI_CODPED " & vbCrLf
    sSql = sSql & "      ,ORDP.SGI_CODPROD " & vbCrLf
    sSql = sSql & "      ,ORDP.SGI_QTDE " & vbCrLf
    sSql = sSql & "      ,ORDP.SGI_IDPRODUTO " & vbCrLf
    sSql = sSql & "      ,PROD.SGI_CODLINPROD " & vbCrLf
    sSql = sSql & "      ,PED.SGI_CODCLI " & vbCrLf
    sSql = sSql & "      ,CORES.SGI_CODCOR " & vbCrLf
    sSql = sSql & "      ,PROD2.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD ORDP" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO PROD" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH PED" & vbCrLf
    sSql = sSql & "      ,SGI_CORESPROD CORES" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO PROD2" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       ORDP.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And ORDP.SGI_DATENTREGA Between  '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'" & vbCrLf
    sSql = sSql & "   And ORDP.SGI_STATUS     = 0" & vbCrLf
    sSql = sSql & "   And PROD.SGI_FILIAL     = ORDP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PROD.SGI_IDPRODUTO  = ORDP.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "   And PED.SGI_FILIAL      = ORDP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PED.SGI_CODIGO      = ORDP.SGI_CODPED" & vbCrLf
    sSql = sSql & "   And CORES.SGI_FILIAL    = PROD.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And CORES.SGI_IDPRODUTO = PROD.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "   And PROD2.SGI_FILIAL    = CORES.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PROD2.SGI_IDPRODUTO = CORES.SGI_CODCOR"
    
    lngREGS = 0
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        Do While Not BREC.EOF()
            lngREGS = (lngREGS + 1)
            BREC.MoveNext
        Loop
        BREC.MoveFirst
        
        ReDim arrREGISTROS(1 To lngREGS, 1 To 10) As String
        lngREGS = 0
        Do While Not BREC.EOF()
            lngREGS = (lngREGS + 1)
            arrREGISTROS(lngREGS, 1) = Trim(Str(BREC!SGI_CODOP))
            arrREGISTROS(lngREGS, 2) = "'" & Format(BREC!SGI_DATENTREGA, "MM/DD/YYYY") & "'"
            arrREGISTROS(lngREGS, 3) = Trim(Str(BREC!SGI_CODPED))
            arrREGISTROS(lngREGS, 4) = "'" & Trim(BREC!SGI_CODPROD) & "'"
            
            strValor = Replace(Format(BREC!SGI_QTDE, "#,##0.00"), ".", "")
            strValor = Replace(strValor, ",", ".")
            arrREGISTROS(lngREGS, 5) = Trim(strValor)
            
            arrREGISTROS(lngREGS, 6) = Trim(Str(BREC!SGI_IDPRODUTO))
            arrREGISTROS(lngREGS, 7) = Trim(Str(BREC!SGI_CODLINPROD))
            arrREGISTROS(lngREGS, 8) = Trim(Str(BREC!SGI_CODCLI))
            arrREGISTROS(lngREGS, 9) = Trim(Str(BREC!SGI_CODCOR))
            arrREGISTROS(lngREGS, 10) = "'" & Trim(BREC!SGI_DESCRICAO) & "'"
            
            BREC.MoveNext
        Loop
    End If
    BREC.Close
    
    objRELOPLITOG.CODIGO = objBLBFunc.Gera_Codigo(Trim(Me.Name), FILIAL, Linha)
    objRELOPLITOG.RELOP = arrREGISTROS
    
    Call objRELOPLITOG.GRAVA("I")
    
    Exit Sub
    
    
    strTipRel = "Todos"
    
    If optTipo(0).Value = True Then strTipRel = strTipRel & "/Análitico"
    If optTipo(1).Value = True Then strTipRel = strTipRel & "/Sintético"
    
    ''BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    ''If BREC.EOF() Then
    ''    BREC.Close
    ''    MsgBox "Não há dados para Imprimr !!!", vbOKOnly + vbExclamation, "Aviso"
    ''    Exit Sub
    ''End If
    ''BREC.Close
    
    strCABEC1 = "Ordem de produção por data de Entrega para Litografia [ " & Trim(strTipRel) & " ]"
    strCABEC2 = "No Periodo de " & mskDTINI.Text & " a " & mskDTFIN.Text
    
    If optFiltro(0).Value = True Then strNomRel = "RELOPLITOG01.rpt"
    If optFiltro(1).Value = True Then strNomRel = ""
    If optFiltro(2).Value = True Then strNomRel = ""

    If Len(Trim(strNomRel)) > 0 Then
        Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelPCP2 & Trim(strNomRel), Linha, 1, strCABEC1, strCABEC2, True)
    End If

End Sub

Private Sub mskDTFIN_GotFocus()
    objBLBFunc.SelecionaCampos mskDTFIN.Name, frmRELOPLITOG
End Sub

Private Sub mskDTINI_GotFocus()
    objBLBFunc.SelecionaCampos mskDTINI.Name, frmRELOPLITOG
End Sub
