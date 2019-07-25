VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmVALPAGOS 
   Caption         =   "Relatório de Vales Pagos"
   ClientHeight    =   1770
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11595
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   11595
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame6 
      Caption         =   "[ Relatório ]"
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
      Left            =   8160
      TabIndex        =   12
      Top             =   960
      Width           =   3375
      Begin VB.OptionButton optRELCOTAANSIN 
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
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optRELCOTAANSIN 
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
         Left            =   600
         TabIndex        =   13
         Top             =   240
         Width           =   1215
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
      Left            =   5400
      TabIndex        =   8
      Top             =   960
      Width           =   2655
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
         Index           =   2
         Left            =   1800
         TabIndex        =   11
         Top             =   240
         Width           =   735
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
         Index           =   1
         Left            =   960
         TabIndex        =   10
         Top             =   240
         Width           =   735
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
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Width           =   5295
      Begin MSMask.MaskEdBox mskDTFIN 
         Height          =   285
         Left            =   3960
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
         Left            =   2880
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11535
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
         Picture         =   "frmVALPAGOS.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "frmVALPAGOS.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Exclui Empresa"
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmVALPAGOS"
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
Dim objVALPAGOS         As Object
Dim objPESQPADRAO       As Object
Dim objREL              As Object
Dim strCABEC1           As String
Dim strCABEC2           As String

Private Sub cmdImpressao_Click()
    If ConfereCampos = False Then Exit Sub
    Call ImpRel
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
    Set objVALPAGOS = CreateObject("RELCOMERCIAL.clsVALPAGOS")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objBLBFunc.LimpaCampos frmVALPAGOS
    objVALPAGOS.FILIAL = FILIAL

    mskDTINI.Text = Format(Now, "DD/MM/YYYY")
    mskDTFIN.Text = Format((Now + 30), "DD/MM/YYYY")
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

    optAgrup(0).Value = True
    optRELCOTAANSIN(0).Value = True

End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objVALPAGOS = Nothing
    Set objPESQPADRAO = Nothing
    Set objREL = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

Private Sub mskDTFIN_GotFocus()
    objBLBFunc.SelecionaCampos mskDTFIN.Name, frmVALPAGOS
End Sub

Private Sub mskDTINI_GotFocus()
    objBLBFunc.SelecionaCampos mskDTINI.Name, frmVALPAGOS
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

Private Sub ImpRel()

    Dim strNomRel   As String
    Dim strTIPO     As String
    Dim strAGRUPA   As String
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       SGI_VALESHEAD.SGI_DTPGTO" & vbCrLf
    sSql = sSql & "      ,SGI_VALESHEAD.SGI_CODVALE" & vbCrLf
    sSql = sSql & "      ,SGI_VALESHEAD.SGI_DTVALE" & vbCrLf
    sSql = sSql & "      ,SGI_VALESHEAD.SGI_VALOR" & vbCrLf
    sSql = sSql & "      ,SGI_VALESHEAD.SGI_PORCDESC" & vbCrLf
    sSql = sSql & "      ,SGI_VALESHEAD.SGI_VLPAGO" & vbCrLf
    
    sSql = sSql & "      ,SGI_CADPEDVENDH.SGI_CODCLI" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_RAZAOSOC" & vbCrLf
     
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "      SGI_CADCLIENTE SGI_CADCLIENTE" & vbCrLf
    sSql = sSql & "     ,SGI_CADPEDVENDH SGI_CADPEDVENDH" & vbCrLf
    sSql = sSql & "     ,SGI_VALESHEAD SGI_VALESHEAD" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_VALESHEAD.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_VALESHEAD.SGI_DTPGTO BETWEEN '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'" & vbCrLf
    
    sSql = sSql & "   And SGI_VALESHEAD.SGI_FILIAL  = SGI_CADPEDVENDH.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_VALESHEAD.SGI_CODVALE = SGI_CADPEDVENDH.SGI_CODIGO" & vbCrLf
    
    sSql = sSql & "   And SGI_CADPEDVENDH.SGI_FILIAL = SGI_CADCLIENTE.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADPEDVENDH.SGI_CODCLI = SGI_CADCLIENTE.SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF Then
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    If optAgrup(0).Value = True Then strAGRUPA = "Dia"
    If optAgrup(1).Value = True Then strAGRUPA = "Mês"
    If optAgrup(2).Value = True Then strAGRUPA = "Ano"
    
    strCABEC1 = "Relatório de vendas realizadas agrupado por " & strAGRUPA & " "
    
    If optRELCOTAANSIN(0).Value = True Then
        If optAgrup(0).Value = True Then strNomRel = "RELVALES01.rpt"
        If optAgrup(1).Value = True Then strNomRel = "RELVALES02.rpt"
        If optAgrup(2).Value = True Then strNomRel = "RELVALES03.rpt"
        strCABEC2 = "Analitico"
    ElseIf optRELCOTAANSIN(1).Value = True Then
        If optAgrup(0).Value = True Then strNomRel = ""
        If optAgrup(1).Value = True Then strNomRel = ""
        If optAgrup(2).Value = True Then strNomRel = ""
        strCABEC2 = "Sintético"
    End If
    
    If CDate(mskDTINI.Text) = CDate(mskDTFIN.Text) Then
        strCABEC2 = strCABEC2 & " - Dia : " & mskDTINI.Text
    ElseIf CDate(mskDTINI.Text) <> CDate(mskDTFIN.Text) Then
        strCABEC2 = strCABEC2 & " - no periodo de " & mskDTINI.Text & " a " & mskDTFIN.Text
    End If
    
    If Len(Trim(strNomRel)) > 0 Then
        Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelComercial & Trim(strNomRel), Linha, 1, strCABEC1, strCABEC2, True)
    End If
    
    Exit Sub
    
End Sub

