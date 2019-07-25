VERSION 5.00
Begin VB.Form frmCADESPTEC 
   Caption         =   "Cadastro de especificação técnica"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9435
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   9435
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9375
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
         Left            =   1080
         Picture         =   "frmCADESPTEC.frx":0000
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
         Left            =   1920
         Picture         =   "frmCADESPTEC.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "frmCADESPTEC.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   9375
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   2
         Text            =   "txtCodigo"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtDescricao 
         Height          =   1965
         Left            =   1200
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Text            =   "frmCADESPTEC.frx":0736
         Top             =   600
         Width           =   7815
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Width           =   660
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmCADESPTEC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho   As String
Public Linha      As Variant
Public cTipOper   As String
Public iCodigo    As Long
Public Filial     As Long
Public strAcesso  As String
Dim objBLBFunc    As Object
Dim objCADESPTEC  As Object

Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de especificação técnica - [ ALTERACAO ]"
    cTipOper = "A"
    
    txtDescricao.SetFocus

End Sub

Private Sub CmdSalva_Click()

    If ValidaCampos = True Then
       
       If cTipOper = "I" Then objCADESPTEC.ESPTECCOD = objCADESPTEC.Gera_Codigo(Me.Name)
       
       objCADESPTEC.ESPTECDES = txtDescricao.Text
       
       If objCADESPTEC.GRAVA(cTipOper) = True Then
          
          MsgBox "A especificação técnica foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
          If cTipOper = "I" Then
             Set objBLBFunc = Nothing
             Set objCADESPTEC = Nothing
             Unload Me
          End If
          
       End If
    
    End If

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objCADESPTEC = Nothing
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
   Set objCADESPTEC = CreateObject("CADESPTEC.clsCADESPTEC")
   
   objCADESPTEC.Filial = Filial
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

End Sub


Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, frmCADESPTEC
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
   KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de especificação técnica - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADESPTEC
    
    txtCodigo.Text = ""
   
End Sub

Public Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    
    Me.Caption = "Cadastro de especificação técnica - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADESPTEC
    
    objCADESPTEC.ESPTECCOD = iCodigo
    
    If objCADESPTEC.Carrega_campos = True Then
    
       txtCodigo.Text = Str(objCADESPTEC.ESPTECCOD)
       txtDescricao.Text = objCADESPTEC.ESPTECDES
       
    End If
    
End Sub

Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
   
    Me.Caption = "Cadastro de especificação técnica - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADESPTEC
    
    objCADESPTEC.ESPTECCOD = iCodigo
    
    If objCADESPTEC.Carrega_campos = True Then
    
       txtCodigo.Text = Str(objCADESPTEC.ESPTECCOD)
       txtDescricao.Text = objCADESPTEC.ESPTECDES
       
    End If

End Sub

Private Function ValidaCampos() As Boolean

     ValidaCampos = False
     
     If Trim(Len(txtDescricao.Text)) = 0 Then
        MsgBox "Especificação técnica inválida !!!", vbOKOnly + vbCritical, "Aviso"
        txtDescricao.SetFocus
        Exit Function
     End If
     
     ValidaCampos = True
     
End Function

