VERSION 5.00
Begin VB.Form frmRELTABCONVERCAO 
   Caption         =   "Relatório de Tabelas de Converção"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7080
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   7080
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   6975
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
         Picture         =   "frmRELTABCONVERCAO.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Picture         =   "frmRELTABCONVERCAO.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Exclui Empresa"
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "[ Código da Tabela ]"
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
      Height          =   615
      Left            =   0
      TabIndex        =   8
      Top             =   960
      Width           =   6975
      Begin VB.TextBox txtCodInic 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         TabIndex        =   10
         Text            =   "txtCodInic"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtCodFim 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   9
         Text            =   "txtCodFim"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Inicial:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Final:"
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
         Index           =   1
         Left            =   2280
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "[ Descrição da Tabela ]"
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
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   1560
      Width           =   6975
      Begin VB.TextBox txtDescIni 
         Height          =   285
         Left            =   840
         MaxLength       =   50
         TabIndex        =   5
         Text            =   "txtDescIni"
         Top             =   240
         Width           =   6015
      End
      Begin VB.TextBox txtDescFim 
         Height          =   285
         Left            =   840
         MaxLength       =   50
         TabIndex        =   4
         Text            =   "txtDescFim"
         Top             =   600
         Width           =   6015
      End
      Begin VB.Label Label1 
         Caption         =   "Inicial:"
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
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Final:"
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
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "[ Ordem ]"
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
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   2520
      Width           =   6975
      Begin VB.OptionButton optOrcem 
         Caption         =   "Código"
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
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optOrcem 
         Caption         =   "Alfabética"
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
         Index           =   1
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmRELTABCONVERCAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public FILIAL           As Integer
Public strAcesso        As String
Dim objBLBFunc          As Object
Dim objRELTABCOMVERCAO  As Object
Dim objPESQPADRAO       As Object
Dim objREL              As Object
''Dim cCamRel             As String

Private Sub cmdImpressao_Click()
    If ValidaCampos = False Then Exit Sub
    Call ImpRel
End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objRELTABCOMVERCAO = Nothing
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
    Set objRELTABCOMVERCAO = CreateObject("RELESTOQUE.clsRELTABCONVERCAO")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objRELTABCOMVERCAO.FILIAL = FILIAL
    
    objBLBFunc.LimpaCampos frmRELTABCONVERCAO
    
    optOrcem(0).Value = True
    
    '' --------------------------------------
    ''cCamRel = "C:\RICARDO\SGI\RELATORIOS\MOSTRAREL\RPT\ESTOQUE\"
    ''cCamRel = "\\pc6\HD\RICARDO\SGI\RELATORIOS\MOSTRAREL\RPT\ESTOQUE\"
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)


End Sub

Private Sub txtCodFim_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCodFim.Text
End Sub

Private Sub txtCodFim_Validate(Cancel As Boolean)
    If Len(Trim(txtCodFim.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCodFim.Text) Then
       MsgBox "Somente é Permitido Numeros !!!", vbOKOnly + vbExclamation, "Aviso"
       txtCodInic.Text = ""
       Cancel = True
    End If
End Sub

Private Sub txtCodInic_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCodInic.Text
End Sub

Private Sub txtCodInic_Validate(Cancel As Boolean)
    If Len(Trim(txtCodInic.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCodInic.Text) Then
       MsgBox "Somente é Permitido Numeros !!!", vbOKOnly + vbExclamation, "Aviso"
       txtCodInic.Text = ""
       Cancel = True
    End If

End Sub

Private Sub txtDescFim_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtDescIni_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Function ValidaCampos() As Boolean

    ValidaCampos = False
    
    If Len(Trim(txtCodInic.Text)) = 0 And Len(Trim(txtCodFim.Text)) > 0 Then
        MsgBox "Código Inicial não pode ser nulo quando existe código Final !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCodInic.SetFocus
        Exit Function
    End If
    If Len(Trim(txtCodInic.Text)) > 0 And Len(Trim(txtCodFim.Text)) > 0 Then
       If CInt(txtCodInic.Text) > CInt(txtCodFim.Text) Then
          MsgBox "Código Inicial não pode ser maior que código Final !!!", vbOKOnly + vbExclamation, "Aviso"
          txtCodInic.Text = ""
          txtCodInic.SetFocus
          Exit Function
       End If
    End If
    If Len(Trim(txtDescIni.Text)) = 0 And Len(Trim(txtDescFim.Text)) > 0 Then
        MsgBox "Descrição Inicial não pode ser nulo quando existe descrição Final !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCodInic.SetFocus
        Exit Function
    End If
    
    ValidaCampos = True

End Function

Private Sub ImpRel()

   
   Dim strCABEC2 As String
   
   sSql = "Select "
   
   sSql = sSql & "       SGI_CADTABCONV.SGI_CODIGO "
   sSql = sSql & "      ,SGI_CADTABCONV.SGI_CODUNIDDE "
   sSql = sSql & "      ,SGI_CADTABCONV.SGI_CODUNIDPARA "
   sSql = sSql & "      ,SGI_CADTABCONV.SGI_QTDEDE "
   sSql = sSql & "      ,SGI_CADTABCONV.SGI_QTDEPARA "
   sSql = sSql & "      ,SGI_CADUNIMED.SGI_DESCRICAO "
   sSql = sSql & "      ,SGI_CADUNIMED.SGI_UNIDADE "
   
   sSql = sSql & "  From "
   sSql = sSql & "       SGI_CADTABCONV SGI_CADTABCONV "
   sSql = sSql & "      ,SGI_CADUNIMED  SGI_CADUNIMED "
   
   sSql = sSql & " Where "
   sSql = sSql & "       SGI_CADTABCONV.SGI_FILIAL = " & FILIAL
   
   '' Regra para Código
   If Len(Trim(txtCodInic.Text)) > 0 And Len(Trim(txtCodFim.Text)) = 0 Then
      sSql = sSql & "   And SGI_CADTABCONV.SGI_CODIGO = " & Trim(txtCodInic.Text)
   End If
   If Len(Trim(txtCodInic.Text)) > 0 And Len(Trim(txtCodFim.Text)) > 0 Then
      sSql = sSql & "   And (SGI_CADTABCONV.SGI_CODIGO >= " & Trim(txtCodInic.Text) & " And SGI_CADTABCONV.SGI_CODIGO <= " & Trim(txtCodFim.Text) & ")"
   End If
      
   sSql = sSql & "   And  SGI_CADUNIMED.SGI_FILIAL = SGI_CADTABCONV.SGI_FILIAL "
   sSql = sSql & "   And  SGI_CADUNIMED.SGI_CODIGO = SGI_CADTABCONV.SGI_CODUNIDDE "
      
   '' Regra para Descrição
   If Len(Trim(txtDescIni.Text)) > 0 And Len(Trim(txtDescFim.Text)) = 0 Then
      sSql = sSql & "   And SGI_CADUNIMED.SGI_DESCRICAO LIKE '" & Trim(txtDescIni.Text) & "%'"
   End If
   If Len(Trim(txtDescIni.Text)) > 0 And Len(Trim(txtDescFim.Text)) > 0 Then
      sSql = sSql & "   And (SGI_CADUNIMED.SGI_DESCRICAO  >= '" & Trim(txtDescIni.Text) & "' And SGI_CADUNIMED.SGI_DESCRICAO <= '" & Trim(txtDescFim.Text) & "')"
   End If
   
   If optOrcem(0).Value = True Then
      sSql = sSql & " Order by SGI_CADTABCONV.SGI_CODIGO"
      strCABEC2 = "Relatório de Tabela de converção por Ordem de " & optOrcem(0).Caption
   ElseIf optOrcem(1).Value = True Then
      sSql = sSql & " Order by SGI_CADUNIMED.SGI_DESCRICAO"
      strCABEC2 = "Relatório de Tabela de converção por Ordem de " & optOrcem(1).Caption
   End If
   
   BREC.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If BREC.EOF Then
      BREC.Close
      MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
      Exit Sub
   End If
   
   BREC.Close
   
   '' Chamada do Relatório
   objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelEstoque & "RELTABCONVERCAO.rpt", Linha, 1, strCABEC2, "", True
   
   txtCodInic.Text = ""
   txtCodFim.Text = ""
   txtDescIni.Text = ""
   txtDescFim.Text = ""
   optOrcem(0).Value = True
   
   txtCodInic.SetFocus

End Sub

