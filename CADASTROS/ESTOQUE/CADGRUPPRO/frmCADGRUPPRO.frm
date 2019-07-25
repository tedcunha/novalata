VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCADGRUPPRO 
   Caption         =   "Cadastro de familia de produtos"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6795
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   6795
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Height          =   2295
      Left            =   0
      TabIndex        =   15
      Top             =   2760
      Width           =   6735
      Begin MSFlexGridLib.MSFlexGrid flxGRUPPROD 
         Height          =   1935
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   3413
         _Version        =   393216
         FixedCols       =   0
         HighLight       =   2
         SelectionMode   =   1
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6735
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
         Picture         =   "frmCADGRUPPRO.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
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
         Picture         =   "frmCADGRUPPRO.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "frmCADGRUPPRO.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   0
      TabIndex        =   8
      Top             =   960
      Width           =   6735
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   9
         Text            =   "txtCodigo"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   0
         Text            =   "txtDescricao"
         Top             =   600
         Width           =   5415
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
         TabIndex        =   11
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
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   930
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   2040
      Width           =   6735
      Begin VB.ComboBox cboSubGruProd 
         Height          =   315
         Left            =   3480
         TabIndex        =   2
         Text            =   "cboSubGruProd"
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtSubGruProd 
         Height          =   285
         Left            =   2340
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "txtSubGruP"
         Top             =   240
         Width           =   750
      End
      Begin VB.CommandButton cmdPesq 
         Height          =   315
         Left            =   3120
         Picture         =   "frmCADGRUPPRO.frx":0736
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmbGravEsp 
         Height          =   315
         Left            =   6240
         Picture         =   "frmCADGRUPPRO.frx":0838
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Sub-familia de Produto :"
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
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmCADGRUPPRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho    As String
Public Linha       As Variant
Public cTipOper    As String
Public iCodigo     As Integer
Public FILIAL      As Integer
Public strAcesso   As String
Public strUSUARIO  As String
Dim objBLBFunc     As Object
Dim objCADGRUPPROD As Object
Dim objPESQPADRAO  As Object
Dim arrSUBGRUP     As Variant

Private Sub cmbGravEsp_Click()
    If cTipOper = "I" Then PreenchGrid
    If cTipOper = "A" Then PreenchGrid
End Sub

Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame3.Enabled = True
       
    Me.Caption = "Cadastro de familia de produtos - [ ALTERACAO ]"
    cTipOper = "A"
    
    txtDescricao.SetFocus
    
    txtSubGruProd.Text = ""
    If cboSubGruProd.ListCount > 0 Then cboSubGruProd.ListIndex = -1

End Sub

Private Sub cmdPesq_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    arrTABELA(1) = "Select * From SGI_CADSUBGRPROD Where SGI_FILIAL = " & FILIAL
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "700"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "3000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Sub-familia de produtos")
    
    If Len(Trim(varRETORNO)) > 0 Then txtSubGruProd.Text = varRETORNO
        
    cboSubGruProd.ListIndex = -1
    txtSubGruProd.SetFocus

End Sub

Private Sub CmdSalva_Click()
    
    Dim I As Integer
    
    If ValidaCampos = True Then
       
       If cTipOper = "I" Then
          objCADGRUPPROD.GRUPRODCOD = objCADGRUPPROD.Gera_Codigo(Me.Name)
       End If
       
       If (flxGRUPPROD.Rows - 1) > 0 Then
          ReDim arrSUBGRUP(1 To (flxGRUPPROD.Rows - 1)) As Integer
          For I = 1 To UBound(arrSUBGRUP)
              arrSUBGRUP(I) = Val(flxGRUPPROD.TextMatrix(I, 0))
          Next I
       Else
          ReDim arrSUBGRUP(0) As Integer
       End If
       
       objCADGRUPPROD.GRUPRODESC = txtDescricao.Text
       objCADGRUPPROD.GRUPPRO = arrSUBGRUP
       
       If objCADGRUPPROD.GRAVA(cTipOper) = True Then
          
          MsgBox "A familia de produto foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
          If cTipOper = "I" Then
             Set objBLBFunc = Nothing
             Set objCADGRUPPROD = Nothing
             Unload Me
          End If
          
       End If
    
    End If

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objCADGRUPPROD = Nothing
    Unload Me
End Sub

Private Sub flxGRUPPROD_Click()

  Dim I As Integer
  
  If flxGRUPPROD.Rows = 1 Then Exit Sub
  
  cboSubGruProd.ListIndex = -1
  For I = 0 To (cboSubGruProd.ListCount - 1)
      If cboSubGruProd.ItemData(I) = Str(Val(flxGRUPPROD.TextMatrix(flxGRUPPROD.RowSel, 0))) Then cboSubGruProd.ListIndex = I
  Next I
  txtSubGruProd.Text = Format(flxGRUPPROD.TextMatrix(flxGRUPPROD.RowSel, 0), "00")

End Sub

Private Sub flxGRUPPROD_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDelete Then
       
       If cTipOper = "C" Then Exit Sub
      
       If flxGRUPPROD.Rows = 2 Then flxGRUPPROD.Rows = 1
       
       If flxGRUPPROD.Rows = 1 Then
          txtSubGruProd.Text = ""
          cboSubGruProd.ListIndex = -1
          Exit Sub
       End If
       
       flxGRUPPROD.RemoveItem flxGRUPPROD.RowSel
       flxGRUPPROD_RowColChange
       
    End If

End Sub


Private Sub flxGRUPPROD_RowColChange()
  
  Dim I As Integer
  
  If flxGRUPPROD.Rows = 1 Then Exit Sub
  
  cboSubGruProd.ListIndex = -1
  For I = 0 To (cboSubGruProd.ListCount - 1)
      If cboSubGruProd.ItemData(I) = Str(Val(flxGRUPPROD.TextMatrix(flxGRUPPROD.RowSel, 0))) Then cboSubGruProd.ListIndex = I
  Next I
  txtSubGruProd.Text = Format(flxGRUPPROD.TextMatrix(flxGRUPPROD.RowSel, 0), "00")

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADGRUPPROD = CreateObject("CADGRUPPRO.clsCADGRUPPRO")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
   objCADGRUPPROD.FILIAL = FILIAL
   
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
    Frame3.Enabled = True
   
    Me.Caption = "Cadastro de familia de produtos - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADGRUPPRO
    
    txtCodigo.Text = ""
    
    ConfGrid
    objCADGRUPPROD.PreenchComboSubGru cboSubGruProd
    
    If cboSubGruProd.ListCount > 0 Then cboSubGruProd.ListIndex = -1
   
End Sub

Public Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame3.Enabled = True
    
    Me.Caption = "Cadastro de familia de produtos - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADGRUPPRO
    
    objCADGRUPPROD.GRUPRODCOD = iCodigo
    
    If objCADGRUPPROD.Carrega_campos = True Then
       
       txtCodigo.Text = Str(objCADGRUPPROD.GRUPRODCOD)
       txtDescricao.Text = objCADGRUPPROD.GRUPRODESC
       
       ConfGrid
       CarrGridSubGru
       
       objCADGRUPPROD.PreenchComboSubGru cboSubGruProd
    
       If cboSubGruProd.ListCount > 0 Then cboSubGruProd.ListIndex = -1
    
    End If
    
End Sub

Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
    Frame3.Enabled = False
    
    Me.Caption = "Cadastro de familia de produtos - [ CONSULTA ]"
        
    objBLBFunc.LimpaCampos frmCADGRUPPRO

    objCADGRUPPROD.GRUPRODCOD = iCodigo
    
    If objCADGRUPPROD.Carrega_campos = True Then
       
       txtCodigo.Text = Str(objCADGRUPPROD.GRUPRODCOD)
       txtDescricao.Text = objCADGRUPPROD.GRUPRODESC
       
       ConfGrid
       CarrGridSubGru
       
       objCADGRUPPROD.PreenchComboSubGru cboSubGruProd
    
       If cboSubGruProd.ListCount > 0 Then cboSubGruProd.ListIndex = -1
    
    End If

End Sub

Private Sub CarrGridSubGru()

       Dim I As Integer
       
       arrSUBGRUP = objCADGRUPPROD.GRUPPRO
       
       If IsArray(arrSUBGRUP) = True Then
          
          BREC.ActiveConnection = adoBanco_Dados
          
          For I = 1 To UBound(arrSUBGRUP)
              
              sSql = "Select " & vbCrLf
              sSql = sSql & "       * " & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADSUBGRPROD " & vbCrLf
              sSql = sSql & " Where " & vbCrLf
              sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
              sSql = sSql & "   And SGI_CODIGO = " & Str(arrSUBGRUP(I)) & vbCrLf
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
              If Not BREC.EOF Then flxGRUPPROD.AddItem arrSUBGRUP(I) & vbTab & arrSUBGRUP(I) & vbTab & BREC!SGI_DESCRICAO
              BREC.Close
          
          Next I
       End If

End Sub

Private Sub ConfGrid()
    
    flxGRUPPROD.Rows = 1
    flxGRUPPROD.Cols = 3
    
    flxGRUPPROD.TextMatrix(0, 0) = ""
    flxGRUPPROD.TextMatrix(0, 1) = "Código"
    flxGRUPPROD.TextMatrix(0, 2) = "Espécie"
    
    flxGRUPPROD.ColWidth(0) = 0
    flxGRUPPROD.ColWidth(1) = 700
    flxGRUPPROD.ColWidth(2) = 5000
    
End Sub

Private Sub PreenchGrid()

   Dim I As Integer
      
   If Len(Trim(cboSubGruProd.Text)) = 0 Then
      MsgBox "Informe a espécie de produto !!!", vbOKOnly + vbCritical, "Aviso"
      txtSubGruProd.SetFocus
      Exit Sub
   End If
   
   For I = 1 To (flxGRUPPROD.Rows - 1)
       If cboSubGruProd.ItemData(cboSubGruProd.ListIndex) = flxGRUPPROD.TextMatrix(I, 0) Then Exit Sub
   Next I
   
   flxGRUPPROD.AddItem cboSubGruProd.ItemData(cboSubGruProd.ListIndex) & vbTab & cboSubGruProd.ItemData(cboSubGruProd.ListIndex) & vbTab & cboSubGruProd.Text
   
   txtSubGruProd.Text = ""
   cboSubGruProd.ListIndex = -1
   
   txtSubGruProd.SetFocus
   
End Sub

Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, frmCADGRUPPRO
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtSubGruProd_GotFocus()
    objBLBFunc.SelecionaCampos txtSubGruProd.Name, frmCADGRUPPRO
End Sub

Private Sub txtSubGruProd_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtSubGruProd_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtSubGruProd.Text)) = 0 Then Exit Sub
        
    If IsNumeric(txtSubGruProd.Text) = False Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtSubGruProd.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboSubGruProd.ListIndex = -1
    For I = 0 To (cboSubGruProd.ListCount - 1)
        If cboSubGruProd.ItemData(I) = Str(Val(txtSubGruProd.Text)) Then cboSubGruProd.ListIndex = I
    Next I
    
    If cboSubGruProd.ListIndex = -1 Then
       MsgBox "O Sub Grupo de produto não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtSubGruProd.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Function ValidaCampos() As Boolean

     ValidaCampos = False
     
     If Trim(Len(txtDescricao.Text)) = 0 Then
        MsgBox "Subgrupo de produto inválido !!!", vbOKOnly + vbCritical, "Aviso"
        txtDescricao.SetFocus
        Exit Function
     End If
     
     If cTipOper = "I" Then
        
        sSql = "Select * from SGI_CADGRUPROD Where SGI_DESCRICAO ='" & txtDescricao.Text & "'"
        sSql = sSql & " And SGI_FILIAL = " & FILIAL
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        
        If Not BREC.EOF Then
           MsgBox "Familia de produto já existe !!!", vbOKOnly + vbCritical, "Aviso"
           txtDescricao.SetFocus
           BREC.Close
           Exit Function
        End If
        
        BREC.Close
     
     End If
     
     If cTipOper = "A" Then
        
        If objCADGRUPPROD.GRUPRODESC <> txtDescricao.Text Then
        
           sSql = "Select * from SGI_CADGRUPROD Where SGI_DESCRICAO ='" & txtDescricao.Text & "'"
           sSql = sSql & " And SGI_FILIAL = " & FILIAL
           BREC.Open sSql, adoBanco_Dados
        
           If Not BREC.EOF Then
              MsgBox "Familia de produto existe !!!", vbOKOnly + vbCritical, "Aviso"
              txtDescricao.Text = objCADGRUPPROD.GRUPRODESC
              txtDescricao.SetFocus
              BREC.Close
              Exit Function
           End If
        
           BREC.Close
        
        End If
     
     End If
     
     ValidaCampos = True
     
End Function

