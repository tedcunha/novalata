VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPDVROTULO 
   Caption         =   "Relatório de Pedidos de Vendas Por Rótulo"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   9855
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   0
      TabIndex        =   25
      Top             =   1560
      Width           =   9855
      Begin VB.CommandButton cmdPesqCLI 
         Height          =   315
         Left            =   2640
         Picture         =   "frmPDVROTULO.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtCODLININI 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Text            =   "txtCODLININI"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdPesqCLIFIN 
         Height          =   315
         Left            =   2640
         Picture         =   "frmPDVROTULO.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtCODLINFIN 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Text            =   "txtCODLINFIN"
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblLINFIN 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblLINFIN"
         Height          =   285
         Left            =   3000
         TabIndex        =   31
         Top             =   600
         Width           =   6735
      End
      Begin VB.Label lblLININI 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblLININI"
         Height          =   285
         Left            =   3000
         TabIndex        =   30
         Top             =   240
         Width           =   6735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Linha Final"
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
         Left            =   240
         TabIndex        =   29
         Top             =   600
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Linha Inicial"
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
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   1050
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
      Height          =   615
      Left            =   5160
      TabIndex        =   24
      Top             =   3360
      Width           =   4695
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   15
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "[ Periodo ]"
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
      Left            =   0
      TabIndex        =   23
      Top             =   3360
      Width           =   5175
      Begin VB.OptionButton optDiaMesAno 
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optDiaMesAno 
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   4200
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optDiaMesAno 
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optDiaMesAno 
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "[ Pedidos ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   22
      Top             =   2520
      Width           =   9855
      Begin VB.OptionButton optTdSomSemPed 
         Caption         =   "Todas Faturadas"
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
         Height          =   255
         Index           =   8
         Left            =   5160
         TabIndex        =   34
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton optTdSomSemPed 
         Caption         =   "Todas Liberadas"
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
         Height          =   255
         Index           =   7
         Left            =   3240
         TabIndex        =   33
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton optTdSomSemPed 
         Caption         =   "Fat.Parcial"
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
         Height          =   255
         Index           =   6
         Left            =   3240
         TabIndex        =   32
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optTdSomSemPed 
         Caption         =   "Liberados"
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
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optTdSomSemPed 
         Caption         =   "Fat.Total"
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
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optTdSomSemPed 
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optTdSomSemPed 
         Caption         =   "Bloqueados"
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
         Height          =   255
         Index           =   3
         Left            =   5160
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optTdSomSemPed 
         Caption         =   "Reprovados"
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
         Height          =   255
         Index           =   4
         Left            =   6960
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optTdSomSemPed 
         Caption         =   "Lib. Financeiro"
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
         Height          =   255
         Index           =   5
         Left            =   1440
         TabIndex        =   7
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   19
      Top             =   960
      Width           =   9855
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
         Left            =   1320
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   2880
         TabIndex        =   21
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   9855
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
         Picture         =   "frmPDVROTULO.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   18
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
         Picture         =   "frmPDVROTULO.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Exclui Empresa"
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmPDVROTULO"
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
Dim objRELPDVROTULO  As Object
Dim objPESQPADRAO    As Object
Dim objREL           As Object
Dim strCABEC1        As String
Dim strCABEC2        As String

Private Sub cmdImpressao_Click()
    If ConfereCampos = False Then Exit Sub
    Call ImpRelPedRotulo
End Sub

Private Sub cmdPesqCLI_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODLIN"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODLIN"
    
    arrCAMPOS(2, 1) = "SGI_DESCRI"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "SGI_DESCRI"
    
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Linha de Produtos")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODLININI.Text = varRETORNO
    
    lblLININI.Caption = Trim(PegaDescLinha(txtCODLININI.Text))
    txtCODLININI.SetFocus

End Sub

Private Sub cmdPesqCLIFIN_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODLIN"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODLIN"
    
    arrCAMPOS(2, 1) = "SGI_DESCRI"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "SGI_DESCRI"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Linha de Produtos")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODLINFIN.Text = varRETORNO
    
    lblLINFIN.Caption = Trim(PegaDescLinha(txtCODLINFIN.Text))
    txtCODLINFIN.SetFocus

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objRELPDVROTULO = Nothing
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
    Set objRELPDVROTULO = CreateObject("RELCOMERCIAL.clsPDVROTULO")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objBLBFunc.LimpaCampos frmPDVROTULO
    objRELPDVROTULO.FILIAL = FILIAL

    mskDTINI.Text = Format(Date, "DD/MM/YYYY")
    mskDTFIN.Text = Format(Date + 30, "DD/MM/YYYY")

    optTdSomSemPed(0).Value = True
    optDiaMesAno(0).Value = True
    optRELCOTAANSIN(0).Value = True
    
    lblLININI.Caption = ""
    lblLINFIN.Caption = ""
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    

End Sub

Private Sub mskDTFIN_GotFocus()
    objBLBFunc.SelecionaCampos mskDTFIN.Name, frmPDVROTULO
End Sub

Private Sub mskDTINI_GotFocus()
    objBLBFunc.SelecionaCampos mskDTINI.Name, frmPDVROTULO
End Sub

Private Function ConfereCampos() As Boolean
    
    ConfereCampos = False
        
        If Len(Trim(txtCODLININI.Text)) > 0 And Len(Trim(txtCODLINFIN.Text)) > 0 Then
           If CLng(txtCODLININI.Text) > CLng(txtCODLINFIN.Text) Then
              MsgBox "Linha Inicial não pode ser maior que linha Final !!!", vbOKOnly + vbExclamation, "Aviso"
              txtCODLININI.SetFocus
              Exit Function
           End If
        End If
        
        
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


Private Sub ImpRelPedRotulo()

    Dim strNomRel As String
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    
    sSql = sSql & "      SGI_CADPEDVENDI.SGI_CODLINPROD" & vbCrLf
    sSql = sSql & "     ,SGI_CADPEDVENDI.SGI_DATAPED" & vbCrLf
    sSql = sSql & "     ,SGI_CADPEDVENDI.SGI_CODPROD" & vbCrLf
    sSql = sSql & "     ,SGI_CADPEDVENDI.SGI_CODIGO" & vbCrLf
    sSql = sSql & "     ,SGI_CADPEDVENDI.SGI_QTDE" & vbCrLf
    sSql = sSql & "     ,SGI_CADPEDVENDI.SGI_VLTOT" & vbCrLf
    sSql = sSql & "     ,SGI_CADPEDVENDI.SGI_VLIPI" & vbCrLf
    
    sSql = sSql & "     ,SGI_CADPEDVENDH.SGI_CODCLI" & vbCrLf
    sSql = sSql & "     ,SGI_CADPEDVENDH.SGI_CODVEND" & vbCrLf
    
    sSql = sSql & "     ,SGI_CADPRODUTO.SGI_IDPRODUTO" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    
    sSql = sSql & "      SGI_CADCLIENTE SGI_CADCLIENTE" & vbCrLf
    sSql = sSql & "     ,SGI_CADLINHAPRODUTO SGI_CADLINHAPRODUTO" & vbCrLf
    sSql = sSql & "     ,SGI_CADPEDVENDH SGI_CADPEDVENDH" & vbCrLf
    sSql = sSql & "     ,SGI_CADPEDVENDI SGI_CADPEDVENDI" & vbCrLf
    sSql = sSql & "     ,SGI_CADPRODUTO SGI_CADPRODUTO" & vbCrLf
    sSql = sSql & "     ,SGI_CADVENDEDOR SGI_CADVENDEDOR" & vbCrLf
        
    sSql = sSql & " Where " & vbCrLf
    
    sSql = sSql & "       SGI_CADPEDVENDI.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And (SGI_CADPEDVENDI.SGI_DATAPED Between '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "')" & vbCrLf
    
    If Len(Trim(txtCODLININI.Text)) > 0 And Len(Trim(txtCODLINFIN.Text)) > 0 Then
       sSql = sSql & "  And   (SGI_CADPEDVENDI.SGI_CODLINPROD Between " & Trim(txtCODLININI.Text) & " And " & Trim(txtCODLINFIN.Text) & ")" & vbCrLf
    ElseIf Len(Trim(txtCODLININI.Text)) > 0 And Len(Trim(txtCODLINFIN.Text)) = 0 Then
       sSql = sSql & "  And   SGI_CADPEDVENDI.SGI_CODLINPROD = " & Trim(txtCODLININI.Text) & vbCrLf
    End If
    
    sSql = sSql & "  And SGI_CADPEDVENDI.SGI_FILIAL = SGI_CADPEDVENDH.SGI_FILIAL " & vbCrLf
    sSql = sSql & "  And SGI_CADPEDVENDI.SGI_CODIGO = SGI_CADPEDVENDH.SGI_CODIGO " & vbCrLf
    
    sSql = sSql & "  And SGI_CADPEDVENDI.SGI_FILIAL = SGI_CADLINHAPRODUTO.SGI_FILIAL " & vbCrLf
    sSql = sSql & "  And SGI_CADPEDVENDI.SGI_CODLINPROD = SGI_CADLINHAPRODUTO.SGI_CODLIN " & vbCrLf
    
    sSql = sSql & "  And SGI_CADPEDVENDI.SGI_FILIAL = SGI_CADPRODUTO.SGI_FILIAL " & vbCrLf
    sSql = sSql & "  And SGI_CADPEDVENDI.SGI_IDPRODUTO = SGI_CADPRODUTO.SGI_IDPRODUTO " & vbCrLf
    
    If optTdSomSemPed(1).Value = True Then
       sSql = sSql & "  And   SGI_CADPEDVENDH.SGI_STATUS = 'F' " & vbCrLf
    ElseIf optTdSomSemPed(2).Value = True Then
       sSql = sSql & "  And   SGI_CADPEDVENDH.SGI_STATUS = 'L'" & vbCrLf
    ElseIf optTdSomSemPed(3).Value = True Then
       sSql = sSql & "  And   SGI_CADPEDVENDH.SGI_STATUS = 'B' " & vbCrLf
    ElseIf optTdSomSemPed(4).Value = True Then
       sSql = sSql & "  And   SGI_CADPEDVENDH.SGI_STATUS = 'R' " & vbCrLf
    ElseIf optTdSomSemPed(5).Value = True Then
       sSql = sSql & "  And   SGI_CADPEDVENDH.SGI_STATUS = 'N' " & vbCrLf
    ElseIf optTdSomSemPed(6).Value = True Then
       sSql = sSql & "  And   SGI_CADPEDVENDH.SGI_STATUS = 'P' " & vbCrLf
    ElseIf optTdSomSemPed(7).Value = True Then
       sSql = sSql & "  And   (SGI_CADPEDVENDH.SGI_STATUS = 'L' Or SGI_CADPEDVENDH.SGI_STATUS = 'N')" & vbCrLf
    ElseIf optTdSomSemPed(8).Value = True Then
       sSql = sSql & "  And   (SGI_CADPEDVENDH.SGI_STATUS = 'F' or SGI_CADPEDVENDH.SGI_STATUS = 'P')" & vbCrLf
    End If
    
    sSql = sSql & "  And SGI_CADPEDVENDH.SGI_FILIAL = SGI_CADCLIENTE.SGI_FILIAL " & vbCrLf
    sSql = sSql & "  And SGI_CADPEDVENDH.SGI_CODCLI = SGI_CADCLIENTE.SGI_CODIGO " & vbCrLf
    
    sSql = sSql & "  And SGI_CADPEDVENDH.SGI_FILIAL  = SGI_CADVENDEDOR.SGI_FILIAL " & vbCrLf
    sSql = sSql & "  And SGI_CADPEDVENDH.SGI_CODVEND = SGI_CADVENDEDOR.SGI_CODIGO " & vbCrLf
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    strCABEC1 = "Relatório de Pedidos por Rótulos "
    
    If optTdSomSemPed(0).Value = True Or _
       optTdSomSemPed(8).Value = True Or _
       optTdSomSemPed(7).Value = True Then
       strCABEC1 = strCABEC1 & " [Todos] "
    ElseIf optTdSomSemPed(1).Value = True Then
       strCABEC1 = strCABEC1 & " [Fat.Total] "
    ElseIf optTdSomSemPed(2).Value = True Then
       strCABEC1 = strCABEC1 & " [Liberados] "
    ElseIf optTdSomSemPed(3).Value = True Then
       strCABEC1 = strCABEC1 & " [Bloqueados] "
    ElseIf optTdSomSemPed(4).Value = True Then
       strCABEC1 = strCABEC1 & " [Reprovados] "
    ElseIf optTdSomSemPed(5).Value = True Then
       strCABEC1 = strCABEC1 & " [Liberado Financeiro] "
    ElseIf optTdSomSemPed(6).Value = True Then
       strCABEC1 = strCABEC1 & " [Fat.Parcial] "
    End If
    
    If optDiaMesAno(0).Value = True Then
       If CDate(mskDTINI.Text) <> CDate(mskDTFIN.Text) Then strCABEC2 = "No Periodo de " & mskDTINI.Text & " a " & mskDTFIN.Text
       If CDate(mskDTINI.Text) = CDate(mskDTFIN.Text) Then strCABEC2 = "Na Data de " & mskDTINI.Text
    ElseIf optDiaMesAno(1).Value = True Then
       If CDate(mskDTINI.Text) <> CDate(mskDTFIN.Text) Then strCABEC2 = "No Mês " & Format(Month(CDate(mskDTINI.Text)), "##00") & "/" & Year(CDate(mskDTINI.Text)) & " ao Mês " & Format(Month(CDate(mskDTFIN.Text)), "##00") & "/" & Year(CDate(mskDTFIN.Text))
       If CDate(mskDTINI.Text) = CDate(mskDTFIN.Text) Then strCABEC2 = "No Mês " & Format(Month(CDate(mskDTINI.Text)), "##00") & "/" & Year(CDate(mskDTFIN.Text))
    ElseIf optDiaMesAno(2).Value = True Then
       If CDate(mskDTINI.Text) <> CDate(mskDTFIN.Text) Then strCABEC2 = "Do Ano " & Year(CDate(mskDTINI.Text)) & " ao Ano " & Year(CDate(mskDTFIN.Text))
       If CDate(mskDTINI.Text) = CDate(mskDTFIN.Text) Then strCABEC2 = "No Ano " & Year(CDate(mskDTFIN.Text))
    ElseIf optDiaMesAno(2).Value = True Then
       If CDate(mskDTINI.Text) <> CDate(mskDTFIN.Text) Then strCABEC2 = "Da Data " & Year(CDate(mskDTINI.Text)) & " a Data " & Year(CDate(mskDTFIN.Text))
       If CDate(mskDTINI.Text) = CDate(mskDTFIN.Text) Then strCABEC2 = "Na Data " & Year(CDate(mskDTFIN.Text))
    End If
    
    strNomRel = ""
    If optDiaMesAno(0).Value = True Then
        If optRELCOTAANSIN(0).Value = True Then
            strNomRel = "RELPDVROTDIAAN2.rpt"
        ElseIf optRELCOTAANSIN(1).Value = True Then
            strNomRel = ""
            ''strNomRel = "RELPDVROTDIASI.rpt"
        End If
    ElseIf optDiaMesAno(1).Value = True Then
        If optRELCOTAANSIN(0).Value = True Then
            strNomRel = "RELPDVROTMESAN2.rpt"
        ElseIf optRELCOTAANSIN(1).Value = True Then
            strNomRel = ""
''            strNomRel = "RELPDVROTMESSI.rpt"
        End If
    ElseIf optDiaMesAno(2).Value = True Then
        If optRELCOTAANSIN(0).Value = True Then
            strNomRel = "RELPDVROTANOAN2.rpt"
        ElseIf optRELCOTAANSIN(1).Value = True Then
            strNomRel = ""
''            strNomRel = "RELPDVROTANOSI.rpt"
        End If
    ElseIf optDiaMesAno(3).Value = True Then
        If optRELCOTAANSIN(0).Value = True Then
            strNomRel = "RELPDVROTSEMAN2.rpt"
        ElseIf optRELCOTAANSIN(1).Value = True Then
            strNomRel = ""
''            strNomRel = "RELPDVROTSEMSI.rpt"
        End If
    End If
    
    If Len(Trim(strNomRel)) > 0 Then
        Call objREL.REL(FILIAL, sSql, strCamRelNovo & cCamRelComercial & Trim(strNomRel), Linha, 1, strCABEC1, strCABEC2, True)
    End If
    
End Sub

Private Sub txtCODLINFIN_GotFocus()
    objBLBFunc.SelecionaCampos txtCODLINFIN.Name, frmPDVROTULO
End Sub

Private Sub txtCODLINFIN_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODLINFIN.Text
End Sub

Private Sub txtCODLINFIN_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODLINFIN.Text)) = 0 Then
       lblLINFIN.Caption = ""
       Exit Sub
    End If
    
    If Not IsNumeric(txtCODLINFIN.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODLINFIN.Text = ""
       Cancel = True
       Exit Sub
    End If

    lblLINFIN.Caption = Trim(PegaDescLinha(txtCODLINFIN.Text))
    If Len(Trim(lblLINFIN.Caption)) = 0 Then
        MsgBox "Linha de Produto Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
        lblLINFIN.Caption = ""
        txtCODLINFIN.Text = ""
        Cancel = True
    End If

End Sub

Private Sub txtCODLININI_GotFocus()
    objBLBFunc.SelecionaCampos txtCODLININI.Name, frmPDVROTULO
End Sub

Private Sub txtCODLININI_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODLININI.Text
End Sub

Private Sub txtCODLININI_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODLININI.Text)) = 0 Then
       lblLININI.Caption = ""
       Exit Sub
    End If
    
    If Not IsNumeric(txtCODLININI.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODLININI.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    lblLININI.Caption = Trim(PegaDescLinha(txtCODLININI.Text))
    If Len(Trim(lblLININI.Caption)) = 0 Then
        MsgBox "Linha de Produto Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
        lblLININI.Caption = ""
        txtCODLININI.Text = ""
        Cancel = True
    End If
        

End Sub

Private Function PegaDescLinha(strCODIGO As String) As String

        PegaDescLinha = ""
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       * " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADLINHAPRODUTO " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_CODLIN = " & Trim(strCODIGO)
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC.EOF() Then PegaDescLinha = BREC!SGI_DESCRI
        BREC.Close
        
End Function
