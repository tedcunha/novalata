VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRELCADEIAVALOR 
   Caption         =   "Relatório de Cadeia de Valor"
   ClientHeight    =   9825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15060
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9825
   ScaleWidth      =   15060
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   15015
      Begin MSMask.MaskEdBox mskDataIni 
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDataFin 
         Height          =   285
         Left            =   2880
         TabIndex        =   6
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
         AutoSize        =   -1  'True
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   2640
         TabIndex        =   7
         Top             =   240
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Periodo de :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1050
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15015
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
         Picture         =   "frmRELCADEIAVALOR.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Picture         =   "frmRELCADEIAVALOR.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Exclui Empresa"
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmRELCADEIAVALOR"
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
Dim objCADEIAVALOR  As Object
Dim objPESQPADRAO   As Object
Dim objREL          As Object

Dim arrProdPadrao()   As DadosProdPadrao


Private Sub cmdImpressao_Click()
    
    If ValidaCampos = False Then Exit Sub
    
    Call PegaDadosPadrao
    
End Sub

Private Sub cmdVoltar_Click()

    Set objBLBFunc = Nothing
    Set objCADEIAVALOR = Nothing
    Set objPESQPADRAO = Nothing
    Set objREL = Nothing
    
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
    Set objCADEIAVALOR = CreateObject("RELPCP.clsCADEIAVALOR")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objCADEIAVALOR.FILIAL = FILIAL
    
    objBLBFunc.LimpaCampos frmRELCADEIAVALOR
    
    mskDataIni.Text = Format(Date, "DD/MM/YYYY")
    mskDataFin.Text = Format(Date, "DD/MM/YYYY")
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    
    '' --------------------------------------
End Sub


Private Function ValidaCampos() As Boolean
    ValidaCampos = False
    
    If Not IsDate(mskDataIni.Text) Then
       MsgBox "Data Inicial Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Function
    End If
    If Not IsDate(mskDataFin.Text) Then
       MsgBox "Data Final Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Function
    End If
    If CDate(mskDataIni.Text) > CDate(mskDataFin.Text) Then
       MsgBox "Data Inicial não pode ser maior que Data Final !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Function
    End If
    
    ValidaCampos = True
End Function

Private Sub PegaDadosPadrao()

    Dim curMetro2 As Currency
    Dim curKlGrm  As Currency
    
    '' =====================================================
    '' Dados do Produto Padrão
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_PADRAO = 1"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
       
       curMetro2 = (BREC!SGI_METROS * (BREC!SGI_LARGURA / 100))
       curKlGrm = (BREC!SGI_PESOUNIT * 1000)
       
       ReDim arrProdPadrao(1 To 1) As DadosProdPadrao
       
       arrProdPadrao(1).curPesoPadrao = BREC!SGI_PESOUNIT
       arrProdPadrao(1).curMetros = BREC!SGI_METROS
       arrProdPadrao(1).curGramMetro2 = Format((curKlGrm / curMetro2), "#,##0.00")
       arrProdPadrao(1).curLargura = BREC!SGI_LARGURA
       arrProdPadrao(1).lngFamMaquinas = BREC!SGI_CADFAMMAQ
       
       arrProdPadrao(1).lngQtdPorTurno = 3
       arrProdPadrao(1).curHorasPorDia = 24
       arrProdPadrao(1).lngMinPorTurno = 480
       arrProdPadrao(1).curSegundosPorTurno = 28800
       
    End If
    BREC.Close
    '' =====================================================
End Sub


Private Function CONVHRMIN(strHORA As String) As Long
    
    CONVHRMIN = 0
    
    If Len(Trim(strHORA)) = 0 Then Exit Function
    
    Dim HORAS       As Long
    Dim MINUTOS     As Long
    Dim TOTMINUTOS  As Long
    
    HORAS = Hour(CDate(strHORA))
    MINUTOS = Minute(CDate(strHORA))
    TOTMINUTOS = ((HORAS * 60) + MINUTOS)
    
    CONVHRMIN = TOTMINUTOS

End Function

Private Function CONVMINHR(lngMINUTOS As Long) As String
    
    CONVMINHR = ""
    
    If lngMINUTOS = 0 Then Exit Function
    
    Dim TOTMINUTOS  As Double
    Dim HORA        As Long
    Dim MINUTO      As Long
    Dim strHORAS    As String
    Dim arrHRMN()   As String
    
    TOTMINUTOS = Round((lngMINUTOS / 60), 2)
    strHORAS = Format(TOTMINUTOS, "###,##000.00")
    arrHRMN = Split(strHORAS, ",")
    
    HORA = CLng(arrHRMN(0))
    MINUTO = (CLng(arrHRMN(1)) * (0.6))
    
    CONVMINHR = Trim(Format(HORA, "##00") & ":" & Format(MINUTO, "##00") & ":" & "00")

End Function

