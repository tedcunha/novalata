VERSION 5.00
Begin VB.Form frmCADTIPOOPERACAO 
   Caption         =   "Cadastro de Tipos de Operação"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7275
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   7275
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7215
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
         Picture         =   "frmCADTIPOOPERACAO.frx":0000
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
         Picture         =   "frmCADTIPOOPERACAO.frx":0532
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
         Picture         =   "frmCADTIPOOPERACAO.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   7215
      Begin VB.CommandButton cmdFamMaq 
         Height          =   315
         Left            =   2880
         Picture         =   "frmCADTIPOOPERACAO.frx":0736
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtCodFamMaq 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         TabIndex        =   10
         Text            =   "txtCodFamMaq"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   2
         Text            =   "txtCodigo"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "txtDescricao"
         Top             =   600
         Width           =   5055
      End
      Begin VB.Label lblDescFamOpe 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescFamOpe"
         Height          =   285
         Left            =   3240
         TabIndex        =   12
         Top             =   960
         Width           =   3855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Familia de Mãquinas"
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
         TabIndex        =   9
         Top             =   960
         Width           =   1740
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label2 
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
         TabIndex        =   3
         Top             =   600
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmCADTIPOOPERACAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public cTipOper         As String
Public iCodigo          As Long
Public Filial           As Integer
Public strAcesso        As String
Dim objBLBFunc          As Object
Dim objCADTIPOPERACAO   As Object
Dim objPESQPADRAO       As Object
Dim arrDESPESAS         As Variant

Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de Tipos de Operação - [ ALTERAÇÃO ]"
    cTipOper = "A"
    
    txtDescricao.SetFocus

End Sub

Private Sub cmdFamMaq_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
        
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADFAMMAQUINAS " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL  = " & Filial
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1500"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRI"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "6000"
    arrCAMPOS(2, 5) = "SGI_DESCRI"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, Filial, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Familia de Máquinas")
    
    If Len(Trim(varRETORNO)) > 0 Then
        txtCodFamMaq.Text = varRETORNO
        Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRI", "SGI_CADFAMMAQUINAS", varRETORNO, lblDescFamOpe)
    End If
    txtCodFamMaq.SetFocus

End Sub

Private Sub CmdSalva_Click()

    Dim I As Integer
    
    If ValidaCampos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADTIPOPERACAO.CODIGO = objCADTIPOPERACAO.Gera_Codigo(Me.Name)
    objCADTIPOPERACAO.DESCRI = txtDescricao.Text
    
    
    objCADTIPOPERACAO.CODFAMMAQ = "Null"
    If Len(Trim(txtCodFamMaq.Text)) > 0 Then objCADTIPOPERACAO.CODFAMMAQ = txtCodFamMaq.Text
    
    If objCADTIPOPERACAO.GRAVA(cTipOper) = False Then Exit Sub
          
    MsgBox "O Tipo de Operação foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
    If objCADTIPOPERACAO.Atualiza(cTipOper, Str(objCADTIPOPERACAO.CODIGO), Filial, Me.Name) = False Then Exit Sub
       
    If cTipOper = "I" Then
       Set objBLBFunc = Nothing
       Set objCADTIPOPERACAO = Nothing
       Set objPESQPADRAO = Nothing
       Unload Me
    End If

End Sub

Private Sub cmdVoltar_Click()
   Set objBLBFunc = Nothing
   Set objCADTIPOPERACAO = Nothing
   Set objPESQPADRAO = Nothing
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
   Set objCADTIPOPERACAO = CreateObject("CADTIPOOPERACAO.clsCADTIPOOPERACAO")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
   objCADTIPOPERACAO.Filial = Filial
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de Tipos de Operação - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADTIPOOPERACAO
    
    txtCodigo.Text = ""
    
    Call LimpaLabel

End Sub

Public Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    
    Me.Caption = "Cadastro de Tipos de Operação - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADTIPOOPERACAO
        
    objCADTIPOPERACAO.CODIGO = iCodigo
    
    Call LimpaLabel
    
    If objCADTIPOPERACAO.Carrega_campos = True Then
        txtCodigo.Text = Str(objCADTIPOPERACAO.CODIGO)
        txtDescricao.Text = objCADTIPOPERACAO.DESCRI
        txtCodFamMaq.Text = objCADTIPOPERACAO.CODFAMMAQ
        If Len(Trim(txtCodFamMaq.Text)) > 0 Then Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRI", "SGI_CADFAMMAQUINAS", txtCodFamMaq.Text, lblDescFamOpe)
    End If
    
End Sub

Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
   
    Me.Caption = "Cadastro de Tipos de Operação - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADTIPOOPERACAO
    
    objCADTIPOPERACAO.CODIGO = iCodigo
    
    Call LimpaLabel
    
    If objCADTIPOPERACAO.Carrega_campos = True Then
        txtCodigo.Text = Str(objCADTIPOPERACAO.CODIGO)
        txtDescricao.Text = objCADTIPOPERACAO.DESCRI
        txtCodFamMaq.Text = objCADTIPOPERACAO.CODFAMMAQ
        If Len(Trim(txtCodFamMaq.Text)) > 0 Then Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRI", "SGI_CADFAMMAQUINAS", txtCodFamMaq.Text, lblDescFamOpe)
    End If


End Sub


Private Function ValidaCampos() As Boolean

     ValidaCampos = False
     
     If Trim(Len(txtDescricao.Text)) = 0 Then
        MsgBox "Descrição de Tipo de Operação !!!", vbOKOnly + vbCritical, "Aviso"
        txtDescricao.SetFocus
        Exit Function
     End If
     
     If cTipOper = "I" Then
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       * " & vbCrLf
        sSql = sSql & "  from " & vbCrLf
        sSql = sSql & "       SGI_TIPOPERACAO  " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_DESCRI = '" & txtDescricao.Text & "'" & vbCrLf
        sSql = sSql & "   And SGI_FILIAL = " & Filial
        
        BREC.Open sSql, adoBanco_Dados
        
        If Not BREC.EOF Then
           MsgBox "Este Tipo de Operação !!!", vbOKOnly + vbCritical, "Aviso"
           txtDescricao.SetFocus
           BREC.Close
           Exit Function
        End If
        
        BREC.Close
     
     End If
     
     If cTipOper = "A" Then
        
        If objCADTIPOPERACAO.DESCRI <> txtDescricao.Text Then
        
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  from " & vbCrLf
           sSql = sSql & "       SGI_TIPOPERACAO  " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_DESCRI = '" & txtDescricao.Text & "'" & vbCrLf
           sSql = sSql & "   And SGI_FILIAL    = " & Filial
           
           BREC.Open sSql, adoBanco_Dados
        
           If Not BREC.EOF Then
              MsgBox "Este Tipo de Operação !!!", vbOKOnly + vbCritical, "Aviso"
              txtDescricao.Text = objCADTIPOPERACAO.DESCRI
              txtDescricao.SetFocus
              BREC.Close
              Exit Function
           End If
        
           BREC.Close
        
        End If
     
     End If
     
     ValidaCampos = True
     
End Function

Private Sub txtCodFamMaq_GotFocus()
    objBLBFunc.SelecionaCampos txtCodFamMaq.Name, frmCADTIPOOPERACAO
End Sub

Private Sub txtCodFamMaq_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCodFamMaq.Text
End Sub

Private Sub txtCodFamMaq_Validate(Cancel As Boolean)

    Dim I As Integer

    If Len(Trim(txtCodFamMaq.Text)) = 0 Then Exit Sub
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRI", "SGI_CADFAMMAQUINAS", txtCodFamMaq.Text, lblDescFamOpe)
    If Len(Trim(lblDescFamOpe.Caption)) = 0 Then
        txtCodFamMaq.Text = ""
        Cancel = True
        Exit Sub
    End If

End Sub

Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, frmCADTIPOOPERACAO
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub


Private Sub LimpaLabel()
    lblDescFamOpe.Caption = ""
End Sub


Private Sub PegaDescTabelas(StrCampoPesq As String, StrCampoRetorno As String, strTabela As String, strCodigo As String, lblLabel As Label)

    lblLabel.Caption = ""
    
    If Len(Trim(strCodigo)) = 0 Then Exit Sub
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoRetorno) & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & Trim(strTabela) & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoPesq) & " = " & Trim(strCodigo)
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then
       lblLabel.Caption = Trim(BREC10(Trim(StrCampoRetorno)))
    Else
       MsgBox "Registro Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC10.Close
    
End Sub

