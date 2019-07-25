VERSION 5.00
Begin VB.Form frmCADTIPPRO 
   Caption         =   "Cadastro de tipo de Produtos"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6390
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6390
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Width           =   6375
      Begin VB.Frame Frame5 
         Caption         =   "[ Homologada ]"
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
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   2655
         Begin VB.OptionButton optHOMSN 
            Caption         =   "SIM"
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
            Left            =   360
            TabIndex        =   17
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton optHOMSN 
            Caption         =   "NÃO"
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
            Left            =   1440
            TabIndex        =   16
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Ativo na Arvore de outro Produto"
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
         Left            =   3000
         TabIndex        =   12
         Top             =   1080
         Width           =   3255
         Begin VB.OptionButton optAtNao 
            Caption         =   "Não"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2280
            TabIndex        =   14
            Top             =   360
            Width           =   735
         End
         Begin VB.OptionButton optAtSim 
            Caption         =   "Sim"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tem Arvore de Produtos"
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
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   2655
         Begin VB.OptionButton optAPNao 
            Caption         =   "Não"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1560
            TabIndex        =   11
            Top             =   360
            Width           =   735
         End
         Begin VB.OptionButton optAPSim 
            Caption         =   "Sim"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "txtDescricao"
         Top             =   600
         Width           =   5055
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
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   375
         TabIndex        =   7
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6375
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
         Left            =   840
         Picture         =   "frmCADTIPPRO.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   735
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
         Left            =   1560
         Picture         =   "frmCADTIPPRO.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   735
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
         Picture         =   "frmCADTIPPRO.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmCADTIPPRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho   As String
Public Linha      As Variant
Public cTipOper   As String
Public iCodigo    As Integer
Public FILIAL     As Integer
Public strAcesso  As String
Public strUSUARIO As String
Dim objBLBFunc    As Object
Dim objCADTIPPROD As Object

Private Sub cmdAltera_Click()
    
    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de Tipos de Produtos - [ ALTERACAO ]"
    cTipOper = "A"
    
    txtDescricao.SetFocus

End Sub

Private Sub CmdSalva_Click()

    If ValidaCampos = True Then
       
       If cTipOper = "I" Then
          objCADTIPPROD.TIPPRODCOD = objCADTIPPROD.Gera_Codigo(Me.Name)
       End If
       
       objCADTIPPROD.TIPPRODESC = txtDescricao.Text
       
       If optAPSim.Value = True Then objCADTIPPROD.TEMLST = 1
       If optAPNao.Value = True Then objCADTIPPROD.TEMLST = 2
       
       If optAtSim.Value = True Then objCADTIPPROD.COMPLST = 1
       If optAtNao.Value = True Then objCADTIPPROD.COMPLST = 2
       
       If optHOMSN(0).Value = True Then objCADTIPPROD.HOMOLOG = 0
       If optHOMSN(1).Value = True Then objCADTIPPROD.HOMOLOG = 1
       
       If objCADTIPPROD.GRAVA(cTipOper) = True Then
          
          MsgBox "O tipo de produto foi " & IIf(cTipOper = "I", "incluso", IIf(cTipOper = "A", "alterado", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
          If cTipOper = "I" Then Inclui
          
       End If
    
    End If

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objCADTIPPROD = Nothing
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
   Set objCADTIPPROD = CreateObject("CADTIPPRO.clsCADTIPPRO")
   
   objCADTIPPROD.FILIAL = FILIAL
   
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
    Frame4.Enabled = False
   
    Me.Caption = "Cadastro de tipos de produtos - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADTIPPRO
    
    txtCodigo.Text = ""
    optAPNao.Value = True
    optAtNao.Value = True
   
    optHOMSN(0).Value = True
   
End Sub

Public Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    
    Me.Caption = "Cadastro de tipos de produtos - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADTIPPRO
    
    objCADTIPPROD.TIPPRODCOD = iCodigo
    
    If objCADTIPPROD.Carrega_campos = True Then
    
       txtCodigo.Text = Str(objCADTIPPROD.TIPPRODCOD)
       txtDescricao.Text = objCADTIPPROD.TIPPRODESC
       
       If objCADTIPPROD.TEMLST = 1 Then optAPSim.Value = True
       If objCADTIPPROD.TEMLST = 2 Then optAPNao.Value = True
       If objCADTIPPROD.TEMLST = 0 Then optAPNao.Value = True
       
       If objCADTIPPROD.COMPLST = 1 Then optAtSim.Value = True
       If objCADTIPPROD.COMPLST = 2 Then optAtNao.Value = True
       If objCADTIPPROD.COMPLST = 0 Then optAtNao.Value = True
       
       optHOMSN(objCADTIPPROD.HOMOLOG).Value = True
    
    End If
    
End Sub

Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
   
    Me.Caption = "Cadastro de tipos de produtos - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADTIPPRO
    
    objCADTIPPROD.TIPPRODCOD = iCodigo
    
    If objCADTIPPROD.Carrega_campos = True Then
    
       txtCodigo.Text = Str(objCADTIPPROD.TIPPRODCOD)
       txtDescricao.Text = objCADTIPPROD.TIPPRODESC
       
       If objCADTIPPROD.TEMLST = 1 Then optAPSim.Value = True
       If objCADTIPPROD.TEMLST = 2 Then optAPNao.Value = True
       If objCADTIPPROD.TEMLST = 0 Then optAPNao.Value = True
       
       If objCADTIPPROD.COMPLST = 1 Then optAtSim.Value = True
       If objCADTIPPROD.COMPLST = 2 Then optAtNao.Value = True
       If objCADTIPPROD.COMPLST = 0 Then optAtNao.Value = True
    
       optHOMSN(objCADTIPPROD.HOMOLOG).Value = True
    End If

End Sub

Private Function ValidaCampos() As Boolean

     ValidaCampos = False
     
     If Trim(Len(txtDescricao.Text)) = 0 Then
        MsgBox "Descrição do tipo de produto inválido !!!", vbOKOnly + vbCritical, "Aviso"
        txtDescricao.SetFocus
        Exit Function
     End If
     
     If cTipOper = "I" Then
        
        'sSql = "Select * from SGI_CADTIPPROD Where SGI_DESCRICAO ='" & txtDescricao.Text & "'"
        'sSql = sSql & " And SGI_FILIAL = " & FILIAL
        
        'BREC.Open sSql, adoBanco_Dados
        
        'If Not BREC.EOF Then
        '   MsgBox "Descrição do tipo de produto já existe !!!", vbOKOnly + vbCritical, "Aviso"
        '   txtDescricao.SetFocus
        '   BREC.Close
        '   Exit Function
        'End If
        
        'BREC.Close
     
     End If
     
     If cTipOper = "A" Then
        
        If objCADTIPPROD.TIPPRODESC <> txtDescricao.Text Then
        
           sSql = "Select * from SGI_CADTIPPROD Where SGI_DESCRICAO ='" & txtDescricao.Text & "'"
           sSql = sSql & " And SGI_FILIAL = " & FILIAL
           BREC.Open sSql, adoBanco_Dados
        
           If Not BREC.EOF Then
              MsgBox "Descrição do tipo de produto existe !!!", vbOKOnly + vbCritical, "Aviso"
              txtDescricao.Text = objCADTIPPROD.TIPPRODESC
              txtDescricao.SetFocus
              BREC.Close
              Exit Function
           End If
        
           BREC.Close
        
        End If
     
     End If
     
     ValidaCampos = True
     
End Function

Private Sub optAPNao_Click()
    optAtNao.Value = True
    Frame4.Enabled = False
End Sub

Private Sub optAPSim_Click()
    Frame4.Enabled = True
End Sub

Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, frmCADTIPPRO
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
   KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub
