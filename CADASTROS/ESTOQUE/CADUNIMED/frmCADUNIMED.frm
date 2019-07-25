VERSION 5.00
Begin VB.Form frmCADUNIMED 
   Caption         =   "Cadastro de unidade de medidas"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8235
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   8235
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8175
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
         Picture         =   "frmCADUNIMED.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Picture         =   "frmCADUNIMED.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   735
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
         Left            =   840
         Picture         =   "frmCADUNIMED.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   8175
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   6240
         TabIndex        =   18
         Top             =   1680
         Width           =   1815
         Begin VB.OptionButton optSinNao 
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
            ForeColor       =   &H8000000D&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton optSinNao 
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
            ForeColor       =   &H8000000D&
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   19
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.TextBox txtFator 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         TabIndex        =   16
         Text            =   "txtFator"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtCODFAMUNIDADE 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         TabIndex        =   14
         Text            =   "txtCODFAMUNIDADE"
         Top             =   1320
         Width           =   855
      End
      Begin VB.ComboBox cboCODFAMUNIDADE 
         Height          =   315
         Left            =   3240
         TabIndex        =   13
         Text            =   "cboCODFAMUNIDADE"
         Top             =   1320
         Width           =   4815
      End
      Begin VB.CommandButton Command25 
         Height          =   315
         Left            =   2880
         Picture         =   "frmCADUNIMED.frx":0736
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtUnidade 
         Height          =   285
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "txtUnidade"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   0
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
         Width           =   6015
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Unidade Padrão:"
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
         Index           =   3
         Left            =   4680
         TabIndex        =   17
         Top             =   1680
         Width           =   1440
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fator de Converção:"
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
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   1755
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Familia de Unidade:"
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
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Unidade:"
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
         Left            =   135
         TabIndex        =   10
         Top             =   960
         Width           =   780
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
         Left            =   135
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmCADUNIMED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho     As String
Public Linha        As Variant
Public cTipOper     As String
Public iCodigo      As Integer
Public FILIAL       As Integer
Public strAcesso    As String
Dim objBLBFunc      As Object
Dim objCADUNIMED    As Object
Dim objPESQPADRAO   As Object

Private Sub cboCODFAMUNIDADE_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboCODFAMUNIDADE, KeyAscii
End Sub

Private Sub cboCODFAMUNIDADE_Validate(Cancel As Boolean)
    If cboCODFAMUNIDADE.ListIndex > -1 Then txtCODFAMUNIDADE.Text = Str(cboCODFAMUNIDADE.ItemData(cboCODFAMUNIDADE.ListIndex))
End Sub

Private Sub cmdAltera_Click()
    
    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de unidade de medidas - [ ALTERACAO ]"
    cTipOper = "A"
    
    txtDescricao.SetFocus

End Sub

Private Sub CmdSalva_Click()

    If ValidaCampos = True Then
       
       If cTipOper = "I" Then
          objCADUNIMED.UNIMEDCOD = objCADUNIMED.Gera_Codigo(Me.Name)
       End If
       
       objCADUNIMED.UNIMEDDES = txtDescricao.Text
       objCADUNIMED.UNIMEDUNI = txtUnidade.Text
       
       If Len(Trim(txtCODFAMUNIDADE.Text)) > 0 Then objCADUNIMED.CODFAMUNID = CLng(txtCODFAMUNIDADE.Text)
       If optSinNao(0).Value = True Then objCADUNIMED.PADRAO = 0
       If optSinNao(1).Value = True Then objCADUNIMED.PADRAO = 1
       If Len(Trim(txtFator.Text)) > 0 Then objCADUNIMED.FATOR = CCur(txtFator.Text)
       
       If objCADUNIMED.GRAVA(cTipOper) = True Then
          
          MsgBox "A unidade de medida foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
          If cTipOper = "I" Then
             Set objBLBFunc = Nothing
             Set objCADUNIMED = Nothing
             Unload Me
          End If
          
       End If
    
    End If

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objCADUNIMED = Nothing
    Set objPESQPADRAO = Nothing
    Unload Me
End Sub

Private Sub Command25_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
        
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADFAMUNIDADE " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL  = " & FILIAL
    
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Familia de Unidades")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODFAMUNIDADE.Text = varRETORNO
    
    cboCODFAMUNIDADE.ListIndex = -1
    txtCODFAMUNIDADE.SetFocus

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()
   
   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADUNIMED = CreateObject("CADUNIMED.clsCADUNIMED")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
   objCADUNIMED.FILIAL = FILIAL
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta


End Sub

Private Sub optSinNao_Click(Index As Integer)
    If Index = 1 Then txtFator.Text = 1
    If Index = 0 Then txtFator.Text = 0
End Sub

Private Sub txtCODFAMUNIDADE_GotFocus()
    objBLBFunc.SelecionaCampos txtCODFAMUNIDADE.Name, frmCADUNIMED
End Sub

Private Sub txtCODFAMUNIDADE_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODFAMUNIDADE.Text
End Sub

Private Sub txtCODFAMUNIDADE_Validate(Cancel As Boolean)

   Dim I As Integer

   If Len(Trim(txtCODFAMUNIDADE.Text)) = 0 Then Exit Sub
    
   cboCODFAMUNIDADE.ListIndex = -1
   For I = 0 To (cboCODFAMUNIDADE.ListCount - 1)
       If cboCODFAMUNIDADE.ItemData(I) = CInt(txtCODFAMUNIDADE.Text) Then cboCODFAMUNIDADE.ListIndex = I
   Next I
    
   If cboCODFAMUNIDADE.ListIndex = -1 Then
      MsgBox "Esta familia de unidade não existe !!!", vbOKOnly + vbCritical, "Aviso"
      txtCODFAMUNIDADE.Text = ""
      Cancel = True
      Exit Sub
   End If

End Sub

Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, frmCADUNIMED
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtFator_GotFocus()
    objBLBFunc.SelecionaCampos txtFator.Name, frmCADUNIMED
End Sub

Private Sub txtFator_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtFator.Text
End Sub

Private Sub txtFator_Validate(Cancel As Boolean)

    If Len(Trim(txtFator.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtFator.Text) = False Then
       MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
       txtFator.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtFator.Text = Format(txtFator.Text, "#,####0.0000")

End Sub

Private Sub txtUnidade_GotFocus()
    objBLBFunc.SelecionaCampos txtUnidade.Name, frmCADUNIMED
End Sub

Private Sub txtUnidade_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de unidade de medidas - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADUNIMED
    
    txtCodigo.Text = ""
    optSinNao(0).Value = True
    
    Call objCADUNIMED.PreenchComboFamiliaUnidade(cboCODFAMUNIDADE)
   
End Sub

Private Function ValidaCampos() As Boolean

     ValidaCampos = False
     
     If Trim(Len(txtDescricao.Text)) = 0 Then
        MsgBox "Descrição da unidade de medida inválido !!!", vbOKOnly + vbCritical, "Aviso"
        txtDescricao.SetFocus
        Exit Function
     End If
     
     If Trim(Len(txtUnidade.Text)) = 0 Then
        MsgBox "Unidade de medida inválido !!!", vbOKOnly + vbCritical, "Aviso"
        txtUnidade.SetFocus
        Exit Function
     End If
     
     If Len(Trim(txtCODFAMUNIDADE.Text)) = 0 Then
        MsgBox "Informe a Familia de Unidade !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCODFAMUNIDADE.SetFocus
        Exit Function
     End If
     
     If Len(Trim(txtFator.Text)) = 0 Then
        MsgBox "Informe o Fator de Conversão !!!", vbOKOnly + vbExclamation, "Aviso"
        txtFator.SetFocus
        Exit Function
     End If
     
     If CCur(txtFator.Text) <= 0 Then
        MsgBox "Informe o Fator de Conversão !!!", vbOKOnly + vbExclamation, "Aviso"
        txtFator.SetFocus
        Exit Function
     End If
     
     If cTipOper = "I" Then
        
        sSql = "Select * from SGI_CADUNIMED Where SGI_DESCRICAO ='" & txtDescricao.Text & "'"
        sSql = sSql & " And SGI_FILIAL = " & FILIAL
        
        BREC.Open sSql, adoBanco_Dados
        
        If Not BREC.EOF Then
           MsgBox "Descrição da unidade de medida já existe !!!", vbOKOnly + vbCritical, "Aviso"
           txtDescricao.SetFocus
           BREC.Close
           Exit Function
        End If
        
        BREC.Close
     
        '' -------------------------------
        sSql = "Select * from SGI_CADUNIMED Where SGI_UNIDADE ='" & txtUnidade.Text & "'"
        sSql = sSql & " And SGI_FILIAL = " & FILIAL
        
        BREC.Open sSql, adoBanco_Dados
        
        If Not BREC.EOF Then
           MsgBox "Unidade de medida já existe !!!", vbOKOnly + vbCritical, "Aviso"
           txtUnidade.SetFocus
           BREC.Close
           Exit Function
        End If
        
        BREC.Close
        
        '' -------------------------------
        If optSinNao(1).Value = True Then
            sSql = "Select " & vbCrLf
            sSql = sSql & "       * " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CADUNIMED " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       SGI_FILIAL     = " & FILIAL & vbCrLf
            sSql = sSql & "   And SGI_CODFAMUNID = " & txtCODFAMUNIDADE.Text & vbCrLf
            sSql = sSql & "   And SGI_PADRAO     = 1"
            BREC.Open sSql, adoBanco_Dados
            If Not BREC.EOF Then
               MsgBox "Já existe uma unidade definida como padrão para esta familia de Unidades !!!", vbOKOnly + vbExclamation, "Aviso"
               BREC.Close
               Exit Function
            End If
            BREC.Close
        End If
        
     End If
     
     If cTipOper = "A" Then
        
        If objCADUNIMED.UNIMEDDES <> txtDescricao.Text Then
        
           sSql = "Select * from SGI_CADUNIMED Where SGI_DESCRICAO ='" & txtDescricao.Text & "'"
           sSql = sSql & " And SGI_FILIAL = " & FILIAL
           BREC.Open sSql, adoBanco_Dados
        
           If Not BREC.EOF Then
              MsgBox "Descrição da unidade de medida já existe !!!", vbOKOnly + vbCritical, "Aviso"
              txtDescricao.Text = objCADUNIMED.UNIMEDDES
              txtDescricao.SetFocus
              BREC.Close
              Exit Function
           End If
        
           BREC.Close
           
        End If
        
        '' -----------------------
        If objCADUNIMED.UNIMEDUNI <> txtUnidade.Text Then
        
           sSql = "Select * from SGI_CADUNIMED Where SGI_UNIDADE ='" & txtUnidade.Text & "'"
           sSql = sSql & " And SGI_FILIAL = " & FILIAL
           BREC.Open sSql, adoBanco_Dados
        
           If Not BREC.EOF Then
              MsgBox "Unidade de medida já existe !!!", vbOKOnly + vbCritical, "Aviso"
              txtUnidade.Text = objCADUNIMED.UNIMEDUNI
              txtUnidade.SetFocus
              BREC.Close
              Exit Function
           End If
        
           BREC.Close
        
        End If
        
        '' -----------------------
        If optSinNao(1).Value = True Then
            sSql = "Select " & vbCrLf
            sSql = sSql & "       * " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CADUNIMED " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       SGI_FILIAL     = " & FILIAL & vbCrLf
            sSql = sSql & "   And SGI_CODIGO     <> " & objCADUNIMED.UNIMEDCOD & vbCrLf
            sSql = sSql & "   And SGI_CODFAMUNID = " & txtCODFAMUNIDADE.Text & vbCrLf
            sSql = sSql & "   And SGI_PADRAO     = 1"
            BREC.Open sSql, adoBanco_Dados
            If Not BREC.EOF Then
               MsgBox "Já existe uma unidade definida como padrão para esta familia de Unidades !!!", vbOKOnly + vbExclamation, "Aviso"
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
   
    Me.Caption = "Cadastro de unidade de medidas - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADUNIMED
    
    objCADUNIMED.UNIMEDCOD = iCodigo
    
    optSinNao(0).Value = True
    
    Call objCADUNIMED.PreenchComboFamiliaUnidade(cboCODFAMUNIDADE)
    
    If objCADUNIMED.Carrega_campos = True Then
    
       txtCodigo.Text = Str(objCADUNIMED.UNIMEDCOD)
       txtDescricao.Text = objCADUNIMED.UNIMEDDES
       txtUnidade.Text = objCADUNIMED.UNIMEDUNI
    
       If objCADUNIMED.CODFAMUNID > 0 Then
          txtCODFAMUNIDADE.Text = objCADUNIMED.CODFAMUNID
          For I = 0 To (cboCODFAMUNIDADE.ListCount - 1)
              If cboCODFAMUNIDADE.ItemData(I) = txtCODFAMUNIDADE.Text Then cboCODFAMUNIDADE.ListIndex = I
          Next I
       End If
       If objCADUNIMED.FATOR > 0 Then txtFator.Text = Format(objCADUNIMED.FATOR, "#,####0.0000")
       optSinNao(objCADUNIMED.PADRAO).Value = True
       
    End If

End Sub

Public Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    
    Me.Caption = "Cadastro de unidade de medidas - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADUNIMED
    
    optSinNao(0).Value = True
    
    objCADUNIMED.UNIMEDCOD = iCodigo
    
    Call objCADUNIMED.PreenchComboFamiliaUnidade(cboCODFAMUNIDADE)
    
    If objCADUNIMED.Carrega_campos = True Then
    
       txtCodigo.Text = Str(objCADUNIMED.UNIMEDCOD)
       txtDescricao.Text = objCADUNIMED.UNIMEDDES
       txtUnidade.Text = objCADUNIMED.UNIMEDUNI
    
       If objCADUNIMED.CODFAMUNID > 0 Then
          txtCODFAMUNIDADE.Text = objCADUNIMED.CODFAMUNID
          For I = 0 To (cboCODFAMUNIDADE.ListCount - 1)
              If cboCODFAMUNIDADE.ItemData(I) = txtCODFAMUNIDADE.Text Then cboCODFAMUNIDADE.ListIndex = I
          Next I
       End If
       If objCADUNIMED.FATOR > 0 Then txtFator.Text = Format(objCADUNIMED.FATOR, "#,####0.0000")
       optSinNao(objCADUNIMED.PADRAO).Value = True
    
    End If
    
End Sub

