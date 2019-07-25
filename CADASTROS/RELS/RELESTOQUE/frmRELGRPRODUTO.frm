VERSION 5.00
Begin VB.Form frmRELGRPRODUTO 
   Caption         =   "Relatório de Grupo de Produtos"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7320
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   7320
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   7215
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
         Picture         =   "frmRELGRPRODUTO.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Exclui Empresa"
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
         Picture         =   "frmRELGRPRODUTO.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "[ Tipo de Agrupamento ]"
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
      TabIndex        =   12
      Top             =   1080
      Width           =   7215
      Begin VB.OptionButton optAgrup 
         Caption         =   "Normal"
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
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optAgrup 
         Caption         =   "Por Sub-Grupo de Produtos"
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
         TabIndex        =   13
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "[ Grupo ]"
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
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   7215
      Begin VB.TextBox txtCodGrpIni 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Text            =   "txtCodGrpIni"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtCodGrpFim 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Text            =   "txtCodGrpFim"
         Top             =   720
         Width           =   1095
      End
      Begin VB.Frame Frame6 
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
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   6975
         Begin VB.OptionButton optOrdSubGrp 
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
            TabIndex        =   7
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optOrdSubGrp 
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
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   6
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.ComboBox cboEspGrpInic 
         Height          =   315
         Left            =   2640
         TabIndex        =   4
         Text            =   "cboEspGrpInic"
         Top             =   360
         Width           =   4455
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   2280
         Picture         =   "frmRELGRPRODUTO.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   375
      End
      Begin VB.ComboBox cboEspGrpFim 
         Height          =   315
         Left            =   2640
         TabIndex        =   2
         Text            =   "cboEspGrpInic"
         Top             =   735
         Width           =   4455
      End
      Begin VB.CommandButton Command3 
         Height          =   315
         Left            =   2280
         Picture         =   "frmRELGRPRODUTO.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Cód. Inicial"
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
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Cód. Final"
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
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmRELGRPRODUTO"
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

Dim objBLBFunc       As Object
Dim objRELGRPRODUTO  As Object
Dim objPESQPADRAO    As Object
Dim objREL           As Object
''Dim cCamRel          As String
Private Sub cmdImpressao_Click()
    If ValidaCampos = False Then Exit Sub
    Call ImpRel
End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objRELGRPRODUTO = Nothing
    Set objPESQPADRAO = Nothing
    Set objREL = Nothing
    Unload Me
End Sub

Private Sub Command2_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 4) As String
    ReDim arrTABELA(1 To 1) As String
    
    arrTABELA(1) = "Select * from SGI_CADGRUPROD Where SGI_FILIAL = " & FILIAL
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "3000"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Grupo de Produtos")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCodGrpIni.Text = varRETORNO
    
    cboEspGrpInic.ListIndex = -1
    txtCodGrpIni.SetFocus

End Sub

Private Sub Command3_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 4) As String
    ReDim arrTABELA(1 To 1) As String
    
    arrTABELA(1) = "Select * from SGI_CADGRUPROD Where SGI_FILIAL = " & FILIAL
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "3000"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Grupo de Produtos")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCodGrpFim.Text = varRETORNO
    
    cboEspGrpFim.ListIndex = -1
    txtCodGrpFim.SetFocus


End Sub

Private Sub Form_Activate()
    If optAgrup(0).Value = True Then txtCodGrpIni.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objRELGRPRODUTO = CreateObject("RELESTOQUE.clsRELGRPRODUTO")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    objRELGRPRODUTO.FILIAL = FILIAL
    
    objBLBFunc.LimpaCampos frmRELGRPRODUTO
    
    
    Call objRELGRPRODUTO.PreencheComboGrpProdI(cboEspGrpInic)
    Call objRELGRPRODUTO.PreencheComboGrpProdF(cboEspGrpFim)
    
    optAgrup(0).Value = True
    optOrdSubGrp(0).Value = True
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    
    '' --------------------------------------
    ''cCamRel = "C:\RICARDO\SGI\RELATORIOS\MOSTRAREL\RPT\ESTOQUE\"
    ''cCamRel = "\\pc6\HD\RICARDO\SGI\RELATORIOS\MOSTRAREL\RPT\ESTOQUE\"

End Sub

Private Function ValidaCampos() As Boolean

    ValidaCampos = False
    
    If Len(Trim(txtCodGrpIni.Text)) = 0 And Len(Trim(txtCodGrpFim.Text)) > 0 Then
       MsgBox "O código do Sub-Grupo de produto inicial não pode ser vázio quando o código do Sub-Grupo final está informada !!!", vbOKOnly + vbExclamation, "Aviso"
       txtCodGrpIni.SetFocus
       Exit Function
    End If
    If Len(Trim(txtCodGrpIni.Text)) > 0 And Len(Trim(txtCodGrpFim.Text)) > 0 Then
       If Int(txtCodGrpIni.Text) > Int(txtCodGrpFim.Text) Then
          MsgBox "O código do Sub-Grupo de produto inicial não pode ser maior que o código de Sub-Grupo final !!!", vbOKOnly + vbExclamation, "Aviso"
          txtCodGrpIni.SetFocus
          Exit Function
       End If
    End If
    If Len(Trim(cboEspGrpInic.Text)) = 0 And Len(Trim(cboEspGrpFim.Text)) > 0 Then
       MsgBox "O Grupo final está informada e o Grupo inicial não foi informada !!!", vbOKOnly + vbExclamation, "Aviso"
       cboEspGrpInic.SetFocus
       Exit Function
    End If
    
    ValidaCampos = True

End Function

Private Sub ImpRel()

   Dim strCABEC2 As String
   Dim strCABEC3 As String
   
   sSql = "Select "
   
   sSql = sSql & "       SGI_CADGRUPROD.SGI_CODIGO "
   sSql = sSql & "      ,SGI_CADGRUPROD.SGI_DESCRICAO "
   sSql = sSql & "  From "
   sSql = sSql & "       SGI_CADGRUPROD SGI_CADGRUPROD "
   sSql = sSql & " Where "
   sSql = sSql & "       SGI_CADGRUPROD.SGI_FILIAL = " & FILIAL
   
   '' Regra para Código
   If Len(Trim(txtCodGrpIni.Text)) > 0 And Len(Trim(txtCodGrpFim.Text)) = 0 Then
      sSql = sSql & "   And SGI_CADGRUPROD.SGI_CODIGO = " & Trim(txtCodGrpIni.Text)
   End If
   If Len(Trim(txtCodGrpIni.Text)) > 0 And Len(Trim(txtCodGrpFim.Text)) > 0 Then
      sSql = sSql & "   And (SGI_CADGRUPROD.SGI_CODIGO >= " & Trim(txtCodGrpIni.Text) & " And SGI_CADGRUPROD.SGI_CODIGO <= " & Trim(txtCodGrpFim.Text) & ")"
   End If
   
   '' Regra para Descrição
   If Len(Trim(cboEspGrpInic.Text)) > 0 And Len(Trim(cboEspGrpFim.Text)) = 0 Then
      sSql = sSql & "   And SGI_CADGRUPROD.SGI_DESCRICAO LIKE '" & Trim(cboEspGrpInic.Text) & "%'"
   End If
   If Len(Trim(cboEspGrpInic.Text)) > 0 And Len(Trim(cboEspGrpFim.Text)) > 0 Then
      sSql = sSql & "   And (SGI_CADGRUPROD.SGI_DESCRICAO  >= '" & Trim(cboEspGrpInic.Text) & "' And SGI_CADGRUPROD.SGI_DESCRICAO <= '" & Trim(cboEspGrpFim.Text) & "')"
   End If
  
   If optOrdSubGrp(0).Value = True Then
      sSql = sSql & " Order by SGI_CADGRUPROD.SGI_CODIGO"
      strCABEC2 = "Relatório de Grupo de Produtos por Ordem de " & optOrdSubGrp(0).Caption
   ElseIf optOrdSubGrp(1).Value = True Then
      sSql = sSql & " Order by SGI_CADGRUPROD.SGI_DESCRICAO"
      strCABEC2 = "Relatório de Grupo de Produtos por Ordem de " & optOrdSubGrp(1).Caption
   End If
   
   BREC.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If BREC.EOF Then
      BREC.Close
      MsgBox "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
      Exit Sub
   End If
   
   BREC.Close
   
   '' Chamada do Relatório
   If optAgrup(0).Value = True Then '' Normal
     objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelEstoque & "RELGRPRODUTO.rpt", Linha, 1, strCABEC2, "", True
   ElseIf optAgrup(1).Value = True Then '' Por Agrupamento de Grupo
     strCABEC3 = "( Por Agrupamento de Sub-Grupo de Produtos )"
     objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelEstoque & "RELGRPRODSUBGR.rpt", Linha, 1, strCABEC2, strCABEC3, True
   End If

End Sub

Private Sub txtCodGrpFim_GotFocus()
    objBLBFunc.SelecionaCampos txtCodGrpFim.Name, frmRELGRPRODUTO
End Sub

Private Sub txtCodGrpFim_Validate(Cancel As Boolean)
    Dim I As Integer
    
    If Len(Trim(txtCodGrpFim.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCodGrpFim.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCodGrpFim.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtCodGrpIni_GotFocus()
    objBLBFunc.SelecionaCampos txtCodGrpIni.Name, frmRELGRPRODUTO
End Sub

Private Sub txtCodGrpIni_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCodGrpIni.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCodGrpIni.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCodGrpIni.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub
