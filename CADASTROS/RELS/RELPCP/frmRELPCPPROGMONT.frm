VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRELPCPPROGMONT 
   Caption         =   "Programação de Montagem"
   ClientHeight    =   2700
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11940
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   11940
   StartUpPosition =   1  'CenterOwner
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
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   8520
      TabIndex        =   11
      Top             =   960
      Width           =   3375
      Begin VB.OptionButton optORDEM 
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
         Index           =   1
         Left            =   1680
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optORDEM 
         Caption         =   "Por Linha"
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
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "[ Filial ]"
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
      Width           =   2895
      Begin VB.OptionButton optFilial 
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
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optFilial 
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
         Left            =   1680
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Width           =   5415
      Begin MSMask.MaskEdBox mskDTFIN 
         Height          =   285
         Left            =   3960
         TabIndex        =   1
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDTINI 
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Left            =   120
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11895
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
         Picture         =   "frmRELPCPPROGMONT.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Picture         =   "frmRELPCPPROGMONT.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmRELPCPPROGMONT"
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

Dim objBLBFunc              As Object
Dim objRELPCPPROGMONT       As Object
Dim objPESQPADRAO           As Object
Dim objREL                  As Object

Dim strCABEC1               As String
Dim strCABEC2               As String

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
    Set objRELPCPPROGMONT = CreateObject("RELPCP.clsRELPCPPROGMONT")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objBLBFunc.LimpaCampos Me
    objRELPCPPROGMONT.FILIAL = FILIAL

    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    
    mskDTINI.Text = Format(Now, "DD/MM/YYYY")
    mskDTFIN.Text = Format(Now, "DD/MM/YYYY")

    optFilial(0).value = True
    optORDEM(0).value = True
    
    ''Call LimpaListBox
    ''Call PopLSTBoxFam
    ''Call LimpaCamposLabel

End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objRELPCPPROGMONT = Nothing
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

    Dim strTabela       As String
    Dim strNOMEMPRESA   As String
    Dim strNomRel       As String
    Dim boolTEMDADOS    As Boolean
    
    boolTEMDADOS = False
    
    strTabela = ""
    If optFilial(0).value = True Then
        strNOMEMPRESA = "NOVALATA"
    ElseIf optFilial(1).value = True Then
        strTabela = "_STEEL"
        strNOMEMPRESA = "STEEL ROL"
    End If
    
    sSql = ""
    
    sSql = sSql & "Select" & vbCrLf
    sSql = sSql & "       SGI_CADMOVPCP" & strTabela & ".SGI_DATAPROG" & vbCrLf
    sSql = sSql & "     , SGI_CADMOVPCP" & strTabela & ".SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "     , SGI_CADMOVPCP" & strTabela & ".SGI_DATAENTR" & vbCrLf
    sSql = sSql & "     , SGI_CADMOVPCP" & strTabela & ".SGI_CODOP" & vbCrLf
    
    sSql = sSql & "     , SGI_CADPRODUTO.SGI_CODLINPROD" & vbCrLf
    sSql = sSql & "     , SGI_CADPRODUTO.SGI_CODIGO" & vbCrLf
    sSql = sSql & "     , SGI_CADPRODUTO.SGI_DESCRICAO" & vbCrLf
    sSql = sSql & "     , SGI_CADPRODUTO.SGI_NECKIN" & vbCrLf
    
    sSql = sSql & "     , SGI_CADLINHAPRODUTO.SGI_DESCRI" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO SGI_CADLINHAPRODUTO" & vbCrLf
    sSql = sSql & "     , SGI_CADMOVPCP" & strTabela & " SGI_CADMOVPCP" & strTabela & vbCrLf
    sSql = sSql & "     , SGI_CADPRODUTO SGI_CADPRODUTO" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_CADMOVPCP" & strTabela & ".SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CADMOVPCP" & strTabela & ".SGI_DATAPROG  Between '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'" & vbCrLf
    
    sSql = sSql & "   And SGI_CADMOVPCP" & strTabela & ".SGI_FILIAL    = SGI_CADPRODUTO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADMOVPCP" & strTabela & ".SGI_IDPRODUTO = SGI_CADPRODUTO.SGI_IDPRODUTO" & vbCrLf
    
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIAL     = SGI_CADLINHAPRODUTO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADPRODUTO.SGI_CODLINPROD = SGI_CADLINHAPRODUTO.SGI_CODLIN"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then boolTEMDADOS = True
    BREC.Close
    
    If boolTEMDADOS = False Then
        MsgBox "ATENÇÂO" & vbCrLf & _
               "Não há dados para Imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    strCABEC1 = "Programação de Montagem " & strNOMEMPRESA
    strCABEC2 = "No Periodo de " & mskDTINI.Text & " a " & mskDTFIN.Text
    
    strNomRel = ""
    If optFilial(0).value = True Then
        strNomRel = ""
    ElseIf optFilial(1).value = True Then
        strNomRel = "RELPCPPROGMONT" & strTabela
    End If
    
    If optORDEM(0).value = True Then
        strNomRel = strNomRel & "_LIN.rpt"
    ElseIf optORDEM(1).value = True Then
        strNomRel = strNomRel & "_DT.rpt"
    End If

    If Len(Trim(strNomRel)) > 0 Then
        Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelPCP2 & Trim(strNomRel), Linha, 1, strCABEC1, strCABEC2, True)
    End If


End Sub

Private Sub mskDTFIN_GotFocus()
    objBLBFunc.SelecionaCampos mskDTFIN.Name, Me
End Sub

Private Sub mskDTINI_GotFocus()
    objBLBFunc.SelecionaCampos mskDTINI.Name, Me
End Sub
