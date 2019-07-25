VERSION 5.00
Begin VB.Form frmRELESPECIE 
   Caption         =   "Relatório de Especie de Produtos"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7035
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   7035
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Caption         =   "[ Ordem ]"
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   0
      TabIndex        =   15
      Top             =   2520
      Width           =   6975
      Begin VB.OptionButton optOrcem 
         Caption         =   "Alfabética"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optOrcem 
         Caption         =   "Código"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "[ Descrição da Espécie ]"
      ForeColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   12
      Top             =   1560
      Width           =   6975
      Begin VB.TextBox txtDescFim 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         MaxLength       =   50
         TabIndex        =   3
         Text            =   "txtDescFim"
         Top             =   600
         Width           =   6015
      End
      Begin VB.TextBox txtDescIni 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         MaxLength       =   50
         TabIndex        =   2
         Text            =   "txtDescIni"
         Top             =   240
         Width           =   6015
      End
      Begin VB.Label Label1 
         Caption         =   "Final:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Inicial:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "[ Código da Espécie ]"
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   960
      Width           =   6975
      Begin VB.TextBox txtCodFim 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   1
         Text            =   "txtCodFim"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtCodInic 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   0
         Text            =   "txtCodInic"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Final:"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Inicial:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6975
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
         Picture         =   "frmRELESPECIE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Exclui Empresa"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdVoltar 
         Caption         =   "&Volta"
         Height          =   615
         Left            =   120
         Picture         =   "frmRELESPECIE.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmRELESPECIE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public FILIAL           As Integer
Public strAcesso        As String
Public lngCodUsuario    As Long

Dim objBLBFunc      As Object
Dim objRELESPECIE   As Object
Dim objPESQPADRAO   As Object
Dim objREL          As Object

Private Sub cmdImpressao_Click()
    If ValidaCampos = False Then Exit Sub
    Call ImpRel
End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objRELESPECIE = Nothing
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
    Set objRELESPECIE = CreateObject("RELESTOQUE.clsRELESPECIE")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objRELESPECIE.FILIAL = FILIAL
    
    objBLBFunc.LimpaCampos frmRELESPECIE
    
    optOrcem(0).Value = True
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    
    '' --------------------------------------
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
Private Sub ImpRel()
   
   Dim strCABEC2 As String
   
   sSql = "Select "
   sSql = sSql & "       SGI_CADESPPROD.SGI_CODIGO "
   sSql = sSql & "      ,SGI_CADESPPROD.SGI_DESCRICAO "
   sSql = sSql & "  From "
   sSql = sSql & "        SGI_CADESPPROD SGI_CADESPPROD "
   sSql = sSql & " Where "
   sSql = sSql & "       SGI_CADESPPROD.SGI_FILIAL = " & FILIAL
   
   '' Regra para Código
   If Len(Trim(txtCodInic.Text)) > 0 And Len(Trim(txtCodFim.Text)) = 0 Then
      sSql = sSql & "   And SGI_CADESPPROD.SGI_CODIGO = " & Trim(txtCodInic.Text)
   End If
   If Len(Trim(txtCodInic.Text)) > 0 And Len(Trim(txtCodFim.Text)) > 0 Then
      sSql = sSql & "   And (SGI_CADESPPROD.SGI_CODIGO >= " & Trim(txtCodInic.Text) & " And SGI_CADESPPROD.SGI_CODIGO <= " & Trim(txtCodFim.Text) & ")"
   End If
   
   '' Regra para Descrição
   If Len(Trim(txtDescIni.Text)) > 0 And Len(Trim(txtDescFim.Text)) = 0 Then
      sSql = sSql & "   And SGI_CADESPPROD.SGI_DESCRICAO LIKE '" & Trim(txtDescIni.Text) & "%'"
   End If
   If Len(Trim(txtDescIni.Text)) > 0 And Len(Trim(txtDescFim.Text)) > 0 Then
      sSql = sSql & "   And (SGI_CADESPPROD.SGI_DESCRICAO  >= '" & Trim(txtDescIni.Text) & "' And SGI_CADESPPROD.SGI_DESCRICAO <= '" & Trim(txtDescFim.Text) & "')"
   End If
   
   If optOrcem(0).Value = True Then
      sSql = sSql & " Order by SGI_CADESPPROD.SGI_CODIGO"
      strCABEC2 = "Relatório de Especie de Produtos por Ordem de " & optOrcem(0).Caption
   ElseIf optOrcem(1).Value = True Then
      sSql = sSql & " Order by SGI_CADESPPROD.SGI_DESCRICAO"
      strCABEC2 = "Relatório de Especie de Produtos por Ordem de " & optOrcem(1).Caption
   End If
   
   BREC.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If BREC.EOF Then
      BREC.Close
      MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
      Exit Sub
   End If
   
   BREC.Close
   
   '' Chamada do Relatório
   objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelEstoque & "RELESPPROD.rpt", Linha, 1, strCABEC2, "", True
   
   txtCodInic.Text = ""
   txtCodFim.Text = ""
   txtDescIni.Text = ""
   txtDescFim.Text = ""
   optOrcem(0).Value = True
   
   txtCodInic.SetFocus

End Sub

