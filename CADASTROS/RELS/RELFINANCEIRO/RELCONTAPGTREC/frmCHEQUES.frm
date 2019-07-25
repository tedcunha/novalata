VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCHEQUES 
   Caption         =   "Relatório de Cheques"
   ClientHeight    =   1905
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14490
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   14490
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      Caption         =   "[ Filtrar Por Data ]"
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
      Height          =   855
      Left            =   0
      TabIndex        =   19
      Top             =   960
      Width           =   1815
      Begin VB.OptionButton optFiltraData 
         Caption         =   "Lançamento"
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
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optFiltraData 
         Caption         =   "Cheque"
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
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "[ Status ]"
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
      Height          =   855
      Left            =   10920
      TabIndex        =   15
      Top             =   960
      Width           =   3495
      Begin VB.OptionButton optStatus 
         Caption         =   "Devolvidos"
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
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optStatus 
         Caption         =   "Pagos"
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
         Left            =   1080
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
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
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
   End
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
      Height          =   855
      Left            =   9360
      TabIndex        =   12
      Top             =   960
      Width           =   1455
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
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   1095
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
         Left            =   120
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
      Height          =   855
      Left            =   6840
      TabIndex        =   8
      Top             =   960
      Width           =   2415
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
         Left            =   840
         TabIndex        =   10
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
         Index           =   2
         Left            =   1560
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   1920
      TabIndex        =   5
      Top             =   960
      Width           =   4815
      Begin MSMask.MaskEdBox mskDTFIN 
         Height          =   285
         Left            =   3480
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
         Left            =   1200
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
         Left            =   2520
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   14415
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
         Picture         =   "frmCHEQUES.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "frmCHEQUES.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Exclui Empresa"
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCHEQUES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho     As String
Public Linha        As Variant
Public FILIAL       As Integer
Public strAcesso    As String
Dim strCABEC1       As String
Dim strCABEC2       As String
Dim strCABEC3       As String
Dim objBLBFunc      As Object
Dim objCHEQUE       As Object
Dim objPESQPADRAO   As Object
Dim objREL          As Object

Private Sub cmdImpressao_Click()
    If ConfereCampos = False Then Exit Sub
    Call ImpRel
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
    Set objCHEQUE = CreateObject("RELCONTAPGTREC.clsCHEQUE")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")

    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objCHEQUE.FILIAL = FILIAL
    
    objBLBFunc.LimpaCampos frmCHEQUES

    mskDTINI.Text = Format(Now, "DD/MM/YYYY")
    mskDTFIN.Text = Format((Now + 30), "DD/MM/YYYY")
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

    optAgrup(0).Value = True
    optRELCOTAANSIN(0).Value = True
    optStatus(0).Value = True
    optFiltraData(1).Value = True

End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objCHEQUE = Nothing
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

Private Sub ImpRel()

    Dim strNomRel   As String
    Dim strTIPO     As String
    Dim strAGRUPA   As String
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_CHEQUESITEN.SGI_CODIGO " & vbCrLf
    sSql = sSql & "      ,SGI_CHEQUESITEN.SGI_DTCHEQUE " & vbCrLf
    sSql = sSql & "      ,SGI_CHEQUESITEN.SGI_BANCO " & vbCrLf
    sSql = sSql & "      ,SGI_CHEQUESITEN.SGI_AGENCIA " & vbCrLf
    sSql = sSql & "      ,SGI_CHEQUESITEN.SGI_CONTA " & vbCrLf
    sSql = sSql & "      ,SGI_CHEQUESITEN.SGI_NCHEQUE " & vbCrLf
    sSql = sSql & "      ,SGI_CHEQUESITEN.SGI_CNPJCPF " & vbCrLf
    sSql = sSql & "      ,SGI_CHEQUESITEN.SGI_VALOR " & vbCrLf
    sSql = sSql & "      ,SGI_CHEQUESITEN.SGI_STATUS " & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CHEQUESHEAD SGI_CHEQUESHEAD " & vbCrLf
    sSql = sSql & "      ,SGI_CHEQUESITEN SGI_CHEQUESITEN " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_CHEQUESITEN.SGI_FILIAL = " & FILIAL & vbCrLf
    
    sSql = sSql & "   And SGI_CHEQUESITEN.SGI_FILIAL = SGI_CHEQUESHEAD.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CHEQUESITEN.SGI_CODIGO = SGI_CHEQUESHEAD.SGI_CODIGO" & vbCrLf
    
    If optFiltraData(1).Value = True Then
        sSql = sSql & "   And SGI_CHEQUESHEAD.SGI_DTLCTO Between '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'" & vbCrLf
    ElseIf optFiltraData(0).Value = True Then
        sSql = sSql & "   And SGI_CHEQUESITEN.SGI_DTCHEQUE Between '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'" & vbCrLf
    End If
        
    If optStatus(1).Value = True Then sSql = sSql & "   And SGI_CHEQUESITEN.SGI_STATUS = 'P' "
    If optStatus(2).Value = True Then sSql = sSql & "   And SGI_CHEQUESITEN.SGI_STATUS = 'D' "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF Then
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    If optStatus(0).Value = True Then strCABEC3 = "Todos"
    If optStatus(1).Value = True Then strCABEC3 = "Pagos"
    If optStatus(2).Value = True Then strCABEC3 = "Devolvidos"
    
    If optAgrup(0).Value = True Then strAGRUPA = "Dia"
    If optAgrup(1).Value = True Then strAGRUPA = "Mês"
    If optAgrup(2).Value = True Then strAGRUPA = "Ano"
    
    strCABEC1 = "Relatório de cheques " & strCABEC3 & "agrupado por " & strAGRUPA & " "
    
    If optRELCOTAANSIN(0).Value = True Then
        If optFiltraData(1).Value = True Then
            If optAgrup(0).Value = True Then strNomRel = "RELCHEQ01A_COP.rpt"
            If optAgrup(1).Value = True Then strNomRel = "RELCHEQ02A_COP.rpt"
            If optAgrup(2).Value = True Then strNomRel = "RELCHEQ03A_COP.rpt"
        ElseIf optFiltraData(0).Value = True Then
            If optAgrup(0).Value = True Then strNomRel = "RELCHEQ01_COP.rpt"
            If optAgrup(1).Value = True Then strNomRel = "RELCHEQ02_COP.rpt"
            If optAgrup(2).Value = True Then strNomRel = "RELCHEQ03_COP.rpt"
        End If
        strCABEC2 = "Analitico"
    ElseIf optRELCOTAANSIN(1).Value = True Then
        If optAgrup(0).Value = True Then strNomRel = ""
        If optAgrup(1).Value = True Then strNomRel = ""
        If optAgrup(2).Value = True Then strNomRel = ""
        strCABEC2 = "Sintético"
    End If
    
    If optFiltraData(1).Value = True Then
        If CDate(mskDTINI.Text) = CDate(mskDTFIN.Text) Then
            strCABEC2 = strCABEC2 & " - Dia : " & mskDTINI.Text
        ElseIf CDate(mskDTINI.Text) <> CDate(mskDTFIN.Text) Then
            strCABEC2 = strCABEC2 & " - no periodo de " & mskDTINI.Text & " a " & mskDTFIN.Text
        End If
    End If
    
    If Len(Trim(strNomRel)) > 0 Then
        Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelContasAPG & Trim(strNomRel), Linha, 1, strCABEC1, strCABEC2, True)
    End If
    
    Exit Sub
    
End Sub

