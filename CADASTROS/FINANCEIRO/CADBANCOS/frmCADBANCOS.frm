VERSION 5.00
Begin VB.Form frmCADBANCOS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de bancos"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   5985
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   5895
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
         Picture         =   "frmCADBANCOS.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmCADBANCOS.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "frmCADBANCOS.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   5895
      Begin VB.TextBox txtCC 
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   2
         Text            =   "txtCC"
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtAgencia 
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "txtAgencia"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "txtCodigo"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   3
         Text            =   "txtDescricao"
         Top             =   1320
         Width           =   4575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "C/C:"
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
         Left            =   600
         TabIndex        =   12
         Top             =   960
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Agência:"
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
         TabIndex        =   11
         Top             =   600
         Width           =   765
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
         TabIndex        =   9
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
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
         Left            =   400
         TabIndex        =   8
         Top             =   1320
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmCADBANCOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cCaminho    As String
Public Linha       As Variant
Public cTipOper    As String
Public iCodigo     As Integer
Public FILIAL      As Integer
Public strACESSO   As String
Public strMODPAI   As String
Public strUSUARIO  As String
Dim objBLBFunc     As Object
Dim objCADBANCOS   As Object

Private Sub cmdAltera_Click()

    If bjBLBFunc.ChecaAcesso2("A", strACESSO) = False Then Exit Sub
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de bancos - [ ALTERACAO ]"
    cTipOper = "A"
    
    txtDescricao.SetFocus

End Sub

Private Sub CmdSalva_Click()
    
    If ValidaCampos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADBANCOS.CADANCOCOD = objCADBANCOS.Gera_Codigo(Me.Name)
    objCADBANCOS.CADANCOAGE = txtAgencia.Text
    objCADBANCOS.CADANCOCOR = txtCC.Text
    objCADBANCOS.CADANCODES = txtDescricao.Text
    
    If objCADBANCOS.GRAVA(cTipOper) = False Then Exit Sub
          
    MsgBox "O banco foi " & IIf(cTipOper = "I", "incluso", IIf(cTipOper = "A", "alterado", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
    If objCADBANCOS.Atualiza(cTipOper, Str(objCADBANCOS.CADANCOCOD), FILIAL, Me.Name) = False Then Exit Sub
       
    If cTipOper = "I" Then
       Set objBLBFunc = Nothing
       Set objCADBANCOS = Nothing
       Unload Me
    End If
    
End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objCADBANCOS = Nothing
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
   Set objCADBANCOS = CreateObject("CADBANCOS.clsCADBANCOS")
      
   objCADBANCOS.FILIAL = FILIAL
   
   Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

   If adoBanco_Dados.State = 0 Then
      MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
      Exit Sub
   End If
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de bancos - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADBANCOS
    
    txtCodigo.Text = ""
   
End Sub

Private Sub txtAgencia_GotFocus()
    objBLBFunc.SelecionaCampos txtAgencia.Name, frmCADBANCOS
End Sub

Private Sub txtAgencia_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtCC_GotFocus()
    objBLBFunc.SelecionaCampos txtCC.Name, frmCADBANCOS
End Sub

Private Sub txtCC_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, frmCADBANCOS
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Function ValidaCampos() As Boolean

     ValidaCampos = False
     
     If Trim(Len(txtDescricao.Text)) = 0 Then
        MsgBox "Banco inválido !!!", vbOKOnly + vbCritical, "Aviso"
        txtDescricao.SetFocus
        Exit Function
     End If
     
     If cTipOper = "I" Then
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       * " & vbCrLf
        sSql = sSql & "  from " & vbCrLf
        sSql = sSql & "       SGI_CADBANCOS  " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_DESCRICAO = '" & txtDescricao.Text & "'" & vbCrLf
        sSql = sSql & "   And SGI_FILIAL = " & FILIAL
        
        BREC.Open sSql, adoBanco_Dados
        
        If Not BREC.EOF Then
           MsgBox "Este banco já existe !!!", vbOKOnly + vbCritical, "Aviso"
           txtDescricao.SetFocus
           BREC.Close
           Exit Function
        End If
        
        BREC.Close
     
     End If
     
     If cTipOper = "A" Then
        
        If objCADBANCOS.CADANCODES <> txtDescricao.Text Then
        
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  from " & vbCrLf
           sSql = sSql & "       SGI_CADBANCOS  " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_DESCRICAO = '" & txtDescricao.Text & "'" & vbCrLf
           sSql = sSql & "   And SGI_FILIAL    = " & FILIAL
           
           BREC.Open sSql, adoBanco_Dados
        
           If Not BREC.EOF Then
              MsgBox "Este banco já existe !!!", vbOKOnly + vbCritical, "Aviso"
              txtDescricao.Text = objCADBANCOS.CADANCODES
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
   
    Me.Caption = "Cadastro de bancos - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADBANCOS
    
    objCADBANCOS.CADANCOCOD = iCodigo
    
    If objCADBANCOS.Carrega_campos = True Then
       txtCodigo.Text = Str(objCADBANCOS.CADANCOCOD)
       txtAgencia.Text = objCADBANCOS.CADANCOAGE
       txtCC.Text = objCADBANCOS.CADANCOCOR
       txtDescricao.Text = objCADBANCOS.CADANCODES
    End If

End Sub

Public Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    
    Me.Caption = "Cadastro de bancos - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADBANCOS
        
    objCADBANCOS.CADANCOCOD = iCodigo
    
    If objCADBANCOS.Carrega_campos = True Then
       txtCodigo.Text = Str(objCADBANCOS.CADANCOCOD)
       txtAgencia.Text = objCADBANCOS.CADANCOAGE
       txtCC.Text = objCADBANCOS.CADANCOCOR
       txtDescricao.Text = objCADBANCOS.CADANCODES
    End If
    
End Sub

