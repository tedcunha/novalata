VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmCADSUBGRPRO 
   Caption         =   "Cadastro de Sub-Familia de produtos"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6765
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6765
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Height          =   2295
      Left            =   0
      TabIndex        =   14
      Top             =   2640
      Width           =   6735
      Begin MSFlexGridLib.MSFlexGrid flxESPECIE 
         Height          =   1935
         Left            =   120
         TabIndex        =   16
         Top             =   240
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
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   0
      TabIndex        =   12
      Top             =   2040
      Width           =   6735
      Begin VB.CommandButton cmbGravEsp 
         Height          =   315
         Left            =   6240
         Picture         =   "frmCADSUBGRPRO.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdPesq 
         Height          =   315
         Left            =   3120
         Picture         =   "frmCADSUBGRPRO.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtEspProd 
         Height          =   285
         Left            =   2340
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "txtEspProd"
         Top             =   240
         Width           =   750
      End
      Begin VB.ComboBox cboEspProd 
         Height          =   315
         Left            =   3480
         TabIndex        =   2
         Text            =   "cboEspProd"
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Espécie de Produto :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1800
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   0
      TabIndex        =   8
      Top             =   960
      Width           =   6735
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   0
         Text            =   "txtDescricao"
         Top             =   600
         Width           =   4815
      End
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
         Height          =   195
         Left            =   120
         TabIndex        =   11
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
         Height          =   195
         Left            =   375
         TabIndex        =   10
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6735
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
         Picture         =   "frmCADSUBGRPRO.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "frmCADSUBGRPRO.frx":0736
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "frmCADSUBGRPRO.frx":0838
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmCADSUBGRPRO"
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
Public strUSUARIO   As String
Dim objBLBFunc      As Object
Dim objCADSUBGRPROD As Object
Dim objPESQPADRAO   As Object
Dim arrEspecie      As Variant

Private Sub cboEspProd_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboEspProd, KeyAscii
End Sub

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
       
    Me.Caption = "Cadastro de Sub-Familia de produtos - [ ALTERACAO ]"
    cTipOper = "A"
    
    txtDescricao.SetFocus
    
    txtEspProd.Text = ""
    If cboEspProd.ListCount > 0 Then cboEspProd.ListIndex = -1

End Sub

Private Sub cmdPesq_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    arrTABELA(1) = "Select * From SGI_CADESPPROD Where SGI_FILIAL = " & FILIAL
    
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "especie de produtos")
    
    If Len(Trim(varRETORNO)) > 0 Then txtEspProd.Text = varRETORNO
        
    cboEspProd.ListIndex = -1
    txtEspProd.SetFocus

End Sub

Private Sub CmdSalva_Click()

    Dim I As Integer
    
    If ValidaCampos = False Then Exit Sub
       
    If cTipOper = "I" Then objCADSUBGRPROD.SUBGRPRODCOD = objCADSUBGRPROD.Gera_Codigo(Me.Name)
    
    If (flxESPECIE.Rows - 1) > 0 Then
       ReDim arrEspecie(1 To (flxESPECIE.Rows - 1)) As Integer
       For I = 1 To UBound(arrEspecie)
           arrEspecie(I) = Val(flxESPECIE.TextMatrix(I, 0))
       Next I
    Else
       ReDim arrEspecie(0) As Integer
    End If
    
    objCADSUBGRPROD.SUBGRPRODESC = txtDescricao.Text
    objCADSUBGRPROD.PRODESPECIE = arrEspecie
    
    If objCADSUBGRPROD.GRAVA(cTipOper) = True Then
       
       MsgBox "A sub-familia de produto foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
       
       If cTipOper = "I" Then
          Set objBLBFunc = Nothing
          Set objCADSUBGRPROD = Nothing
          Unload Me
       End If
       
    End If
    

End Sub

Private Sub cmdVoltar_Click()
   Set objBLBFunc = Nothing
   Set objCADSUBGRPROD = Nothing
   Set objPESQPADRAO = Nothing
   Unload Me
End Sub


Private Sub flxESPECIE_Click()
   
  Dim I As Integer
  
  If flxESPECIE.Rows = 1 Then Exit Sub
  
  For I = 0 To (cboEspProd.ListCount - 1)
      If cboEspProd.ItemData(I) = CInt(flxESPECIE.TextMatrix(flxESPECIE.Row, 0)) Then cboEspProd.ListIndex = I
  Next I
  txtEspProd.Text = Format(flxESPECIE.TextMatrix(flxESPECIE.Row, 0), "00")
  
End Sub

Private Sub flxESPECIE_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
       
       If cTipOper = "C" Then Exit Sub
       
       If flxESPECIE.Rows = 2 Then flxESPECIE.Rows = 1
       
       If flxESPECIE.Rows = 1 Then
          txtEspProd.Text = ""
          cboEspProd.ListIndex = -1
          Exit Sub
       End If
       
       flxESPECIE.RemoveItem flxESPECIE.RowSel
       flxESPECIE_RowColChange
       
    End If
End Sub

Private Sub flxESPECIE_RowColChange()

  If flxESPECIE.Rows = 1 Then Exit Sub
  
  cboEspProd.ListIndex = Val(flxESPECIE.TextMatrix(flxESPECIE.RowSel, 0)) - 1
  txtEspProd.Text = Format(flxESPECIE.TextMatrix(flxESPECIE.RowSel, 0), "00")

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADSUBGRPROD = CreateObject("CADSUBGRPRO.clsCADSUBGRPRO")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
   objCADSUBGRPROD.FILIAL = FILIAL
   
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
   
    Me.Caption = "Cadastro de Sub-Familia de produtos - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADSUBGRPRO
    
    txtCodigo.Text = ""
    
    ConfGrid
    objCADSUBGRPROD.PreenchComboEspecie cboEspProd
    
    If cboEspProd.ListCount > 0 Then cboEspProd.ListIndex = -1
   
End Sub

Private Sub ConfGrid()

    flxESPECIE.Rows = 1
    flxESPECIE.Cols = 3
    
    flxESPECIE.TextMatrix(0, 0) = ""
    flxESPECIE.TextMatrix(0, 1) = "Código"
    flxESPECIE.TextMatrix(0, 2) = "Espécie"
    
    flxESPECIE.ColWidth(0) = 0
    flxESPECIE.ColWidth(1) = 700
    flxESPECIE.ColWidth(2) = 5000
    
End Sub

Private Sub PreenchGrid()

   Dim I As Integer
   
   If Len(Trim(cboEspProd.Text)) = 0 Then
      MsgBox "Informe a espécie de produto !!!", vbOKOnly + vbCritical, "Aviso"
      txtEspProd.SetFocus
      Exit Sub
   End If
   
   For I = 1 To (flxESPECIE.Rows - 1)
       If cboEspProd.ItemData(cboEspProd.ListIndex) = flxESPECIE.TextMatrix(I, 0) Then Exit Sub
   Next I
   
   flxESPECIE.AddItem cboEspProd.ItemData(cboEspProd.ListIndex) & vbTab & cboEspProd.ItemData(cboEspProd.ListIndex) & vbTab & cboEspProd.Text
   
   txtEspProd.Text = ""
   cboEspProd.ListIndex = -1
   
   txtEspProd.SetFocus
   
   
End Sub

Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, frmCADSUBGRPRO
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub


Private Sub txtEspProd_GotFocus()
    objBLBFunc.SelecionaCampos txtEspProd.Name, frmCADSUBGRPRO
End Sub

Private Sub txtEspProd_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtEspProd_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtEspProd.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtEspProd.Text) = False Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtEspProd.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboEspProd.ListIndex = -1
    For I = 0 To (cboEspProd.ListCount - 1)
        If cboEspProd.ItemData(I) = Str(Val(txtEspProd.Text)) Then cboEspProd.ListIndex = I
    Next I
    
    If cboEspProd.ListIndex = -1 Then
       MsgBox "Esta espécie de produto não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtEspProd.Text = ""
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
        
        sSql = "Select * from SGI_CADSUBGRPROD Where SGI_DESCRICAO ='" & txtDescricao.Text & "'"
        sSql = sSql & " And SGI_FILIAL = " & FILIAL
        
        BREC.Open sSql, adoBanco_Dados
        
        If Not BREC.EOF Then
           MsgBox "Sub-familia de produto já existe !!!", vbOKOnly + vbCritical, "Aviso"
           txtDescricao.SetFocus
           BREC.Close
           Exit Function
        End If
        
        BREC.Close
     
     End If
     
     If cTipOper = "A" Then
        
        If objCADSUBGRPROD.SUBGRPRODESC <> txtDescricao.Text Then
        
           sSql = "Select * from SGI_CADSUBGRPROD Where SGI_DESCRICAO ='" & txtDescricao.Text & "'"
           sSql = sSql & " And SGI_FILIAL = " & FILIAL
           BREC.Open sSql, adoBanco_Dados
        
           If Not BREC.EOF Then
              MsgBox "Sub-familia de produto existe !!!", vbOKOnly + vbCritical, "Aviso"
              txtDescricao.Text = objCADSUBGRPROD.SUBGRPRODESC
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
    Frame3.Enabled = False
    
    Me.Caption = "Cadastro de sub-familia de produtos - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADSUBGRPRO
    
    objCADSUBGRPROD.SUBGRPRODCOD = iCodigo
    
    If objCADSUBGRPROD.Carrega_campos = True Then
       
       txtCodigo.Text = Str(objCADSUBGRPROD.SUBGRPRODCOD)
       txtDescricao.Text = objCADSUBGRPROD.SUBGRPRODESC
       
       ConfGrid
       CarrGridEspecie
       
       objCADSUBGRPROD.PreenchComboEspecie cboEspProd
    
       If cboEspProd.ListCount > 0 Then cboEspProd.ListIndex = -1
    
    End If

End Sub

Private Sub CarrGridEspecie()

       Dim I As Integer
       
       arrEspecie = objCADSUBGRPROD.PRODESPECIE
       
       If IsArray(arrEspecie) = True Then
          
          BREC.ActiveConnection = adoBanco_Dados
          
          For I = 1 To UBound(arrEspecie)
              
              sSql = "Select " & vbCrLf
              sSql = sSql & "       * " & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADESPPROD " & vbCrLf
              sSql = sSql & " Where " & vbCrLf
              sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
              sSql = sSql & "   And SGI_CODIGO = " & Str(arrEspecie(I)) & vbCrLf
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
              If Not BREC.EOF Then flxESPECIE.AddItem arrEspecie(I) & vbTab & arrEspecie(I) & vbTab & BREC!SGI_DESCRICAO
              BREC.Close
          
          Next I
       End If

End Sub

Public Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame3.Enabled = True
    
    Me.Caption = "Cadastro de Sub-Familia de produtos - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADSUBGRPRO
    
    objCADSUBGRPROD.SUBGRPRODCOD = iCodigo
    
    If objCADSUBGRPROD.Carrega_campos = True Then
       
       txtCodigo.Text = Str(objCADSUBGRPROD.SUBGRPRODCOD)
       txtDescricao.Text = objCADSUBGRPROD.SUBGRPRODESC
       
       ConfGrid
       CarrGridEspecie
       
       objCADSUBGRPROD.PreenchComboEspecie cboEspProd
    
       If cboEspProd.ListCount > 0 Then cboEspProd.ListIndex = -1
    
    End If
    
End Sub

