VERSION 5.00
Begin VB.Form frmCADGRUPREC 
   Caption         =   "Cadastro de Grupo de Recebimento"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   7845
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7815
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
         Picture         =   "frmCADGRUPREC.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdSalva 
         Caption         =   "&Salva"
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
         Left            =   1800
         MaskColor       =   &H8000000F&
         Picture         =   "frmCADGRUPREC.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdAltera 
         Caption         =   "&Altera"
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
         Picture         =   "frmCADGRUPREC.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   7815
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   2
         Text            =   "txtDescricao"
         Top             =   600
         Width           =   6375
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "txtCodigo"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
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
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
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
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmCADGRUPREC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho     As String
Public Linha        As Variant
Public cTipOper     As String
Public iCodigo      As Long
Public FILIAL       As Integer
Public strACESSO    As String
Public strMODPAI    As String
Public strUSUARIO   As String
Dim objBLBFunc      As Object
Dim objCADGRUPREC   As Object

Private Sub cmdAltera_Click()

    cmdAltera.Enabled = False
    CmdSalva.Enabled = True
    Frame2.Enabled = True
    
    txtDescricao.SetFocus
    
    cTipOper = "A"
    
    Me.Caption = "Cadastro de Grupo de Recebimento - [ ALTERAÇÃO ]"

End Sub

Private Sub CmdSalva_Click()

    Dim I As Integer
    
    If Verifica_Campos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADGRUPREC.CODIGO = objCADGRUPREC.Gera_Codigo(Me.Name)
    
    objCADGRUPREC.DESCRI = txtDescricao.Text
    
    
    '' Grava as informações
    If objCADGRUPREC.GRAVA(cTipOper) = False Then Exit Sub
    If objCADGRUPREC.Atualiza(cTipOper, Str(objCADGRUPREC.CODIGO), FILIAL, Me.Name) = False Then Exit Sub
    
    MsgBox "O Grupo foi " & IIf(cTipOper = "I", "incluso", IIf(cTipOper = "A", "alterado", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
       
    If cTipOper = "I" Then
       Set objBLBFunc = Nothing
       Set objCADGRUPREC = Nothing
       Unload Me
    End If

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objCADGRUPREC = Nothing
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
   Set objCADGRUPREC = CreateObject("CADGRUPREC.clsCADGRUPREC")
   
   objCADGRUPREC.FILIAL = FILIAL
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de Grupo de Recebimento - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADGRUPREC
    
End Sub


Private Function Verifica_Campos() As Boolean

    Verifica_Campos = False
    
    Dim j As Integer
    Dim I As Integer
    Dim blAchou As Boolean
    
    If Len(Trim(txtDescricao.Text)) = 0 Then
       MsgBox "Informe a descrição !!!", vbOKOnly + vbExclamation, "Aviso"
       txtDescricao.SetFocus
       Exit Function
    End If
    
    If cTipOper = "I" Then
       
       sSql = "Select " & vbCrLf
       sSql = sSql & "       * " & vbCrLf
       sSql = sSql & "  From " & vbCrLf
       sSql = sSql & "       SGI_CADGRUPREC " & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
       sSql = sSql & "   And SGI_DESCRI = '" & Trim(txtDescricao.Text) & "'"
       
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
       If Not BREC.EOF Then
          MsgBox "Esta Descrição de Grupo já existe !!!", vbOKOnly + vbExclamation, "Aviso"
          BREC.Close
          txtDescricao.SetFocus
          Exit Function
       End If
       BREC.Close
       
    End If
    If cTipOper = "A" Then
    
       If objCADGRUPREC.DESCRI <> txtDescricao.Text Then
       
          sSql = "Select " & vbCrLf
          sSql = sSql & "       * " & vbCrLf
          sSql = sSql & "  From " & vbCrLf
          sSql = sSql & "       SGI_CADGRUPREC " & vbCrLf
          sSql = sSql & " Where " & vbCrLf
          sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
          sSql = sSql & "   And SGI_DESCRI = '" & Trim(txtDescricao.Text) & "'"
       
          BREC.Open sSql, adoBanco_Dados, adOpenDynamic
          If Not BREC.EOF Then
             MsgBox "Está Descrição de Grupo já existe !!!", vbOKOnly + vbExclamation, "Aviso"
             BREC.Close
             txtDescricao.Text = objCADGRUPREC.DESCRI
             txtDescricao.SetFocus
             Exit Function
          End If
          BREC.Close
          
       End If
       
    End If
    
    Verifica_Campos = True

End Function

Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, frmCADGRUPREC
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
    
    Me.Caption = "Cadastro de Grupo de Recebimento - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADGRUPREC
    
    objCADGRUPREC.CODIGO = iCodigo
    
    If objCADGRUPREC.Carrega_campos = True Then
       txtCodigo.Text = objCADGRUPREC.CODIGO
       txtDescricao.Text = objCADGRUPREC.DESCRI
    End If

End Sub

Private Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    
    Me.Caption = "Cadastro de Grupo de Recebimento - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADGRUPREC
    
    objCADGRUPREC.CODIGO = iCodigo
    
    If objCADGRUPREC.Carrega_campos = True Then
       txtCodigo.Text = objCADGRUPREC.CODIGO
       txtDescricao.Text = objCADGRUPREC.DESCRI
    End If

End Sub

