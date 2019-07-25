VERSION 5.00
Begin VB.Form frmCADPARFICHA 
   Caption         =   "Cadastro de Parâmetros da Ficha Técnica"
   ClientHeight    =   2010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10125
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   10125
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   10095
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   6
         Text            =   "txtDescricao"
         Top             =   600
         Width           =   8175
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "txtCodigo"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
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
         TabIndex        =   8
         Top             =   600
         Width           =   870
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
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10095
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
         Picture         =   "frmCADPARFICHA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "frmCADPARFICHA.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Picture         =   "frmCADPARFICHA.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADPARFICHA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public cTipOper         As String
Public iCodigo          As Integer
Public FILIAL           As Integer
Public strAcesso        As String
Public strMODPAI        As String
Dim objBLBFunc          As Object
Dim objCADPARFICHA      As Object

Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de Parâmetros da Ficha Técnica - [ ALTERAÇÃO ]"
    cTipOper = "A"
    
    txtDescricao.SetFocus

End Sub

Private Sub CmdSalva_Click()

    If ValidaCampos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADPARFICHA.CODIGO = objCADPARFICHA.Gera_Codigo(Me.Name)
    
    objCADPARFICHA.DESCRI = txtDescricao.Text

    If objCADPARFICHA.GRAVA(cTipOper) = False Then Exit Sub
    If objCADPARFICHA.Atualiza(cTipOper, Str(objCADPARFICHA.CODIGO), FILIAL, Me.Name) = False Then Exit Sub
          
    MsgBox "O Parâmetro foi " & IIf(cTipOper = "I", "incluso", IIf(cTipOper = "A", "alterado", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
       
    If cTipOper = "I" Then
       Set objBLBFunc = Nothing
       Set objCADPARFICHA = Nothing
       Unload Me
    End If

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objCADPARFICHA = Nothing
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
   Set objCADPARFICHA = CreateObject("CADPARFICHA.clsCADPARFICHA")
      
   objCADPARFICHA.FILIAL = FILIAL
   
   Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

   If adoBanco_Dados.State = 0 Then
      MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
      Exit Sub
   End If
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

End Sub


Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, frmCADPARFICHA
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub


Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de Parâmetros da Ficha Técnica - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADPARFICHA
    
    txtCodigo.Text = ""
   
End Sub


Private Function ValidaCampos() As Boolean

     ValidaCampos = False
     
     If Trim(Len(txtDescricao.Text)) = 0 Then
        MsgBox "Informe a Descrição !!!", vbOKOnly + vbCritical, "Aviso"
        txtDescricao.SetFocus
        Exit Function
     End If
     
     If cTipOper = "I" Then
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       * " & vbCrLf
        sSql = sSql & "  from " & vbCrLf
        sSql = sSql & "       SGI_CADPARFICHA  " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_DESCRI = '" & txtDescricao.Text & "'" & vbCrLf
        sSql = sSql & "   And SGI_FILIAL = " & FILIAL
        
        BREC.Open sSql, adoBanco_Dados
        If Not BREC.EOF Then
           MsgBox "Este Parâmetro já existe !!!", vbOKOnly + vbCritical, "Aviso"
           txtDescricao.SetFocus
           BREC.Close
           Exit Function
        End If
        BREC.Close
     
     End If
     
     If cTipOper = "A" Then
        
        If objCADPARFICHA.DESCRI <> txtDescricao.Text Then
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  from " & vbCrLf
           sSql = sSql & "       SGI_CADPARFICHA " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_DESCRI = '" & txtDescricao.Text & "'" & vbCrLf
           sSql = sSql & "   And SGI_FILIAL    = " & FILIAL
           
           BREC.Open sSql, adoBanco_Dados
           If Not BREC.EOF Then
              MsgBox "Este Parâmetro já existe !!!", vbOKOnly + vbCritical, "Aviso"
              txtDescricao.Text = objCADPARFICHA.DESCRI
              txtDescricao.SetFocus
              BREC.Close
              Exit Function
           End If
           BREC.Close
        End If
     
     End If
     
     ValidaCampos = True
     
End Function

Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
   
    Me.Caption = "Cadastro de Parâmetros da Ficha Técnica - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADPARFICHA
    
    objCADPARFICHA.CODIGO = iCodigo
    
    If objCADPARFICHA.Carrega_campos = True Then
    
        txtCodigo.Text = objCADPARFICHA.CODIGO
        txtDescricao.Text = objCADPARFICHA.DESCRI
    
    End If

End Sub

Public Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    
    Me.Caption = "Cadastro de Parâmetros da Ficha Técnica - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADPARFICHA
        
    objCADPARFICHA.CODIGO = iCodigo
    
    If objCADPARFICHA.Carrega_campos = True Then
        txtCodigo.Text = objCADPARFICHA.CODIGO
        txtDescricao.Text = objCADPARFICHA.DESCRI
    End If
    
End Sub
