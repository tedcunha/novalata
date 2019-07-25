VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmRELESTATUSCLIE 
   Caption         =   "Relatório de Estatus do Cliente"
   ClientHeight    =   2910
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12345
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   12345
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "[ Progresso....]"
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
      TabIndex        =   12
      Top             =   2160
      Width           =   12255
      Begin ComctlLib.ProgressBar prgProg 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "[ Vendedores ]"
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
      TabIndex        =   8
      Top             =   1560
      Width           =   12255
      Begin VB.TextBox txtCODVEND 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   10
         Text            =   "txtCODVEND"
         Top             =   255
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   1320
         Picture         =   "frmRELESTATUSCLIE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblDescVendedor 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescVendedor"
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Top             =   240
         Width           =   7335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "[ STATUS DO CLIENTE ]"
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
      TabIndex        =   3
      Top             =   960
      Width           =   6495
      Begin VB.OptionButton optEstatus 
         Caption         =   "TODOS OS STATUS"
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
         Left            =   4080
         MaskColor       =   &H00800000&
         TabIndex        =   7
         Top             =   240
         Width           =   2295
      End
      Begin VB.OptionButton optEstatus 
         Caption         =   "INATIVO"
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
         Left            =   2760
         MaskColor       =   &H00800000&
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optEstatus 
         Caption         =   "REGULAR"
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
         Left            =   1320
         MaskColor       =   &H00800000&
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optEstatus 
         Caption         =   "ATIVO"
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
         MaskColor       =   &H00800000&
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12255
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
         Picture         =   "frmRELESTATUSCLIE.frx":0102
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
         Picture         =   "frmRELESTATUSCLIE.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Exclui Empresa"
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmRELESTATUSCLIE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho             As String
Public Linha                As Variant
Public FILIAL               As Integer
Public strAcesso            As String

Dim arrDADOSNOVA()          As String
Dim arrDADOSSTEEL()         As String

Dim objBLBFunc              As Object
Dim objRELESTATUSCLIE       As Object
Dim objPESQPADRAO           As Object
Dim objREL                  As Object

Private Sub Command2_Click()

On Error GoTo Err_Command2_Click


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
    
    If txtCODVEND.Enabled = True Then txtCODVEND.SetFocus

    Exit Sub
    
Err_Command2_Click:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, "C", "Função : Command2_Click()", Me.Name, "Command2_Click()", strCAMARQERRO)

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
    Set objRELESTATUSCLIE = CreateObject("RELCOMERCIAL.clsRELESTATUSCLIE")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objBLBFunc.LimpaCampos Me
    
    objRELESTATUSCLIE.FILIAL = FILIAL
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    Me.Caption = Me.Caption & " / " & Trim(Me.Name)
    
    optEstatus(3).value = True
    Frame3.Visible = False
    prgProg.Min = 0
    
    Call LimpaCampos

End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objRELESTATUSCLIE = Nothing
    Set objPESQPADRAO = Nothing
    Set objREL = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

Private Sub LimpaCampos()
    lblDescVendedor.Caption = ""
End Sub

Private Sub txtCODVEND_GotFocus()

On Error GoTo Err_txtCODVEND_GotFocus
    
    objBLBFunc.SelecionaCampos txtCODVEND.Name, Me

    Exit Sub
    
Err_txtCODVEND_GotFocus:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, "C", "Função : txtCODVEND_GotFocus()", Me.Name, "txtCODVEND_GotFocus()", strCAMARQERRO)

End Sub

Private Sub txtCODVEND_Validate(Cancel As Boolean)

On Error GoTo Err_txtCODVEND_Validate

    Dim i As Integer
    
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
    
    Exit Sub
    
Err_txtCODVEND_Validate:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, "C", "Função : txtCODVEND_Validate()", Me.Name, "txtCODVEND_Validate()", strCAMARQERRO)

End Sub

Private Sub PegaDescTabelas(StrCampoPesq As String, StrCampoRetorno As String, strTabela As String, strCODIGO As String, lblLabel As Label)

On Error GoTo Err_PegaDescTabelas

    lblLabel.Caption = ""
    
    If BREC10.State = 1 Then BREC10.Close
    
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
    
    Exit Sub
    
Err_PegaDescTabelas:

    If BREC10.State = 1 Then BREC10.Close
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, "C", "Função : PegaDescTabelas()", Me.Name, "PegaDescTabelas()", strCAMARQERRO)

End Sub

