VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmRELCORES 
   Caption         =   "Cores de Rótulos"
   ClientHeight    =   3735
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4890
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4890
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   3000
      Width           =   4815
      Begin ComctlLib.ProgressBar pgbProgresso 
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   1335
      Left            =   0
      TabIndex        =   7
      Top             =   1680
      Width           =   4815
      Begin VB.ListBox lstFamProd 
         Height          =   960
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   8
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   4815
      Begin VB.OptionButton optEMpresa 
         Caption         =   "NOVALATA/STEEL"
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
         Top             =   360
         Width           =   2055
      End
      Begin VB.OptionButton optEMpresa 
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
         Left            =   1560
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optEMpresa 
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
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.CommandButton cmdImpressao 
         Caption         =   "&Gera"
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
         Picture         =   "frmRELCORES.frx":0000
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
         Picture         =   "frmRELCORES.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmRELCORES"
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
Dim objRELCORES      As Object
Dim objPESQPADRAO    As Object
Dim objREL           As Object

Dim strCABEC1               As String
Dim strCABEC2               As String
Dim strNomRel               As String

Dim lngPORC                 As Long

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
    Set objRELCORES = CreateObject("RELPCP.clsRELCORES")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")

    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objBLBFunc.LimpaCampos Me
    objRELCORES.FILIAL = FILIAL

    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7) & "RELPREPARA\"
    
    
    optEmpresa(2).value = True
    Call LimpaListBox
    Call PopLSTBoxFam

    Frame4.Enabled = False
    pgbProgresso.Min = 0
    

End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objRELCORES = Nothing
    Set objPESQPADRAO = Nothing
    Set objREL = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

Private Sub LimpaListBox()
    lstFamProd.Clear
End Sub

Private Sub PopLSTBoxFam()

    sSql = ""

    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADGRUPROD " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC.EOF()
        lstFamProd.AddItem Trim(BREC!SGI_DESCRICAO)
        lstFamProd.ItemData(lstFamProd.NewIndex) = BREC!SGI_CODIGO
        BREC.MoveNext
    Loop
    BREC.Close

End Sub


Private Sub GeraXLS()

    strNomRel = ""
    If optEmpresa(0).value = True Then
        strNomRel = "RELCORES_NOVALATA.xls"
    ElseIf optEmpresa(1).value = True Then
        strNomRel = "RELCORES_STEEL.xls"
    ElseIf optEmpresa(2).value = True Then
        strNomRel = "RELCORES_TODOS.xls"
    End If

End Sub
