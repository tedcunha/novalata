VERSION 5.00
Begin VB.Form frmCADESPREPROV 
   Caption         =   "Cadastro de Tipos de Reprovação"
   ClientHeight    =   2130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6435
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   6435
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6375
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
         Picture         =   "frmCADESPREPROV.frx":0000
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
         Picture         =   "frmCADESPREPROV.frx":0102
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
         Picture         =   "frmCADESPREPROV.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   6375
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "txtCodigo"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   1
         Text            =   "txtDescricao"
         Top             =   600
         Width           =   4695
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
         Left            =   360
         TabIndex        =   6
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
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmCADESPREPROV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho       As String
Public Linha          As Variant
Public cTipOper       As String
Public iCodigo        As Integer
Public FILIAL         As Integer
Public strAcesso      As String
Public strMODPAI      As String
Dim objBLBFunc        As Object
Dim objCADESPREPROV   As Object

Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de Tipos de Reprovação - [ ALTERACAO ]"
    cTipOper = "A"
    
    txtDescricao.SetFocus

End Sub

Private Sub CmdSalva_Click()

    If ValidaCampos = False Then Exit Sub
       
    If cTipOper = "I" Then objCADESPREPROV.CODIGO = objCADESPREPROV.Gera_Codigo(Me.Name)
    objCADESPREPROV.DESCRI = txtDescricao.Text
       
    If objCADESPREPROV.GRAVA(cTipOper) = False Then Exit Sub
          
    MsgBox "O tipo de reprovação foi " & IIf(cTipOper = "I", "incluso", IIf(cTipOper = "A", "alterado", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
    If cTipOper = "I" Then
       Set objBLBFunc = Nothing
       Set objCADESPREPROV = Nothing
       Unload Me
    End If
          
End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objCADESPREPROV = Nothing
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
   Set objCADESPREPROV = CreateObject("CADESPREPROV.clsCADESPREPROV")
      
   objCADESPREPROV.FILIAL = FILIAL
   
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
   
    Me.Caption = "Cadastro de Tipos de Reprovação - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADESPREPROV
    
    txtCodigo.Text = ""
   
End Sub

Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, frmCADESPREPROV
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Function ValidaCampos() As Boolean

     ValidaCampos = False
     
     If Trim(Len(txtDescricao.Text)) = 0 Then
        MsgBox "Espécie do tipo de produto inválido !!!", vbOKOnly + vbCritical, "Aviso"
        txtDescricao.SetFocus
        Exit Function
     End If
     
     If cTipOper = "I" Then
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       * " & vbCrLf
        sSql = sSql & "  from " & vbCrLf
        sSql = sSql & "       SGI_CADTIPREP " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_DESCRI = '" & txtDescricao.Text & "'" & vbCrLf
        sSql = sSql & "   And SGI_FILIAL = " & FILIAL
        
        BREC.Open sSql, adoBanco_Dados
        
        If Not BREC.EOF Then
           MsgBox "Espécie do tipo de produto já existe !!!", vbOKOnly + vbCritical, "Aviso"
           txtDescricao.SetFocus
           BREC.Close
           Exit Function
        End If
        
        BREC.Close
     
     End If
     
     If cTipOper = "A" Then
        
        If objCADESPREPROV.DESCRI <> txtDescricao.Text Then
        
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  from " & vbCrLf
           sSql = sSql & "       SGI_CADTIPREP " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_DESCRI = '" & txtDescricao.Text & "'" & vbCrLf
           sSql = sSql & "   And SGI_FILIAL = " & FILIAL
           
           BREC.Open sSql, adoBanco_Dados
        
           If Not BREC.EOF Then
              MsgBox "Espécie do tipo de produto existe !!!", vbOKOnly + vbCritical, "Aviso"
              txtDescricao.Text = objCADESPREPROV.DESCRI
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
   
    Me.Caption = "Cadastro de Tipos de Reprovação - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADESPREPROV
    
    objCADESPREPROV.CODIGO = iCodigo
    
    If objCADESPREPROV.Carrega_campos = True Then
       txtCodigo.Text = Str(objCADESPREPROV.CODIGO)
       txtDescricao.Text = objCADESPREPROV.DESCRI
    End If

End Sub

Public Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    
    Me.Caption = "Cadastro de Tipos de Reprovação - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADESPREPROV
        
    objCADESPREPROV.CODIGO = iCodigo
    
    If objCADESPREPROV.Carrega_campos = True Then
       txtCodigo.Text = Str(objCADESPREPROV.CODIGO)
       txtDescricao.Text = objCADESPREPROV.DESCRI
    End If
    
End Sub

