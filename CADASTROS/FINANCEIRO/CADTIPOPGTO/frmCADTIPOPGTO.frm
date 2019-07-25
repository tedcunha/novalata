VERSION 5.00
Begin VB.Form frmCADTIPOPGTO 
   Caption         =   "Cadastro de Tipo de Pagamento"
   ClientHeight    =   2340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5970
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   5970
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Width           =   5895
      Begin VB.OptionButton optOperacao 
         Caption         =   "Crédito"
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
         TabIndex        =   11
         Top             =   970
         Width           =   1095
      End
      Begin VB.OptionButton optOperacao 
         Caption         =   "Débito"
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
         Left            =   1200
         TabIndex        =   10
         Top             =   970
         Width           =   975
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "txtDescricao"
         Top             =   600
         Width           =   4575
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Operação:"
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
         Index           =   1
         Left            =   150
         TabIndex        =   9
         Top             =   960
         Width           =   900
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
         Index           =   0
         Left            =   120
         TabIndex        =   8
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   5
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
         Picture         =   "frmCADTIPOPGTO.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   975
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
         Picture         =   "frmCADTIPOPGTO.frx":0102
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
         Left            =   1080
         Picture         =   "frmCADTIPOPGTO.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADTIPOPGTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho    As String
Public Linha       As Variant
Public cTipOper    As String
Public iCodigo     As Long
Public Filial      As Long
Public strACESSO   As String
Public strUSUARIO  As String
Dim objBLBFunc     As Object
Dim objCADTIPOPGTO As Object

Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strACESSO) = False Then Exit Sub
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de Tipo de Pagamento - [ ALTERAÇÃO ]"
    cTipOper = "A"
    
    txtDescricao.SetFocus

End Sub

Private Sub CmdSalva_Click()

    If ValidaCampos = True Then
       
       If cTipOper = "I" Then objCADTIPOPGTO.TIPPGTOCOD = objCADTIPOPGTO.Gera_Codigo(Me.Name)
      
       objCADTIPOPGTO.TIPPGTODES = txtDescricao.Text
       If optOperacao(0).Value = True Then objCADTIPOPGTO.OPERACAO = 1
       If optOperacao(1).Value = True Then objCADTIPOPGTO.OPERACAO = 2
       
       If objCADTIPOPGTO.GRAVA(cTipOper) = True Then
          
          MsgBox "A tipo de pagamento foi " & IIf(cTipOper = "I", "incluso", IIf(cTipOper = "A", "alterado", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
          If cTipOper = "I" Then
             Set objBLBFunc = Nothing
             Set objCADTIPOPGTO = Nothing
             Unload Me
          End If
          
       End If
    
    End If

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objCADTIPOPGTO = Nothing
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
   Set objCADTIPOPGTO = CreateObject("CADTIPOPGTO.clsCADTIPOPGTO")
   
   objCADTIPOPGTO.Filial = Filial
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de Tipo de Pagamento - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADTIPOPGTO
    
    optOperacao(0).Value = False
    optOperacao(1).Value = False
    
    txtCodigo.Text = ""
   
End Sub

Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
   
    Me.Caption = "Cadastro de Tipo de Pagamento - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADTIPOPGTO
    optOperacao(0).Value = False
    optOperacao(1).Value = False
    
    objCADTIPOPGTO.TIPPGTOCOD = iCodigo
    
    If objCADTIPOPGTO.Carrega_campos = True Then
    
       txtCodigo.Text = Str(objCADTIPOPGTO.TIPPGTOCOD)
       txtDescricao.Text = objCADTIPOPGTO.TIPPGTODES
       If objCADTIPOPGTO.OPERACAO = 1 Then optOperacao(0).Value = True
       If objCADTIPOPGTO.OPERACAO = 2 Then optOperacao(1).Value = True
       
    End If

End Sub

Public Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    
    Me.Caption = "Cadastro de Tipo de Pagamento - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADTIPOPGTO
    optOperacao(0).Value = False
    optOperacao(1).Value = False
    
    objCADTIPOPGTO.TIPPGTOCOD = iCodigo
    
    If objCADTIPOPGTO.Carrega_campos = True Then
    
       txtCodigo.Text = Str(objCADTIPOPGTO.TIPPGTOCOD)
       txtDescricao.Text = objCADTIPOPGTO.TIPPGTODES
       
    End If
    
End Sub

Private Function ValidaCampos() As Boolean

     ValidaCampos = False
     
     If Trim(Len(txtDescricao.Text)) = 0 Then
        MsgBox "Tipo de documento inválido !!!", vbOKOnly + vbCritical, "Aviso"
        txtDescricao.SetFocus
        Exit Function
     End If
     
     If optOperacao(0).Value = False And _
        optOperacao(1).Value = False Then
        MsgBox "Tipo de operação não pode ser nulo !!!", vbOKOnly + vbCritical, "Aviso"
        Exit Function
     End If
     
     If cTipOper = "I" Then
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       * " & vbCrLf
        sSql = sSql & "  from " & vbCrLf
        sSql = sSql & "       SGI_CADTIPOPGTO " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_DESCRICAO = '" & txtDescricao.Text & "'" & vbCrLf
        If optOperacao(0).Value = True Then sSql = sSql & "   And SGI_SINAL     = '-'"
        If optOperacao(1).Value = True Then sSql = sSql & "   And SGI_SINAL     = '+'"
        sSql = sSql & "   And SGI_FILIAL    = " & Filial
        
        BREC.Open sSql, adoBanco_Dados
        
        If Not BREC.EOF Then
           MsgBox "Tipo de pagamento já existe !!!", vbOKOnly + vbCritical, "Aviso"
           txtDescricao.SetFocus
           BREC.Close
           Exit Function
        End If
        
        BREC.Close
     
     End If
     
     If cTipOper = "A" Then
                
        If objCADTIPOPGTO.TIPPGTODES <> txtDescricao.Text Then
        
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  from " & vbCrLf
           sSql = sSql & "       SGI_CADTIPOPGTO " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_DESCRICAO = '" & txtDescricao.Text & "'" & vbCrLf
           If optOperacao(0).Value = True Then sSql = sSql & "       SGI_SINAL     = '-'"
           If optOperacao(1).Value = True Then sSql = sSql & "       SGI_SINAL     = '+'"
           sSql = sSql & "    And SGI_FILIAL   =  " & Filial
           
           BREC.Open sSql, adoBanco_Dados
          
           If Not BREC.EOF Then
              MsgBox "Tipo de pagamento já existe !!!", vbOKOnly + vbCritical, "Aviso"
              txtDescricao.Text = objCADTIPOPGTO.TIPPGTODES
              txtDescricao.SetFocus
              BREC.Close
              Exit Function
           End If
        
           BREC.Close
        
        End If
     
     End If
     
     ValidaCampos = True
     
End Function

Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, frmCADTIPOPGTO
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

