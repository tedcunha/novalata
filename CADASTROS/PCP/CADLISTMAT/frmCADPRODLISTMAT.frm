VERSION 5.00
Begin VB.Form frmCADPRODLISTMAT 
   Caption         =   "Cadastro de Lista de Material"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10485
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   10485
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Caption         =   $"frmCADPRODLISTMAT.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   10455
      Begin VB.TextBox txtProduto 
         Height          =   315
         Left            =   120
         MaxLength       =   15
         TabIndex        =   0
         Text            =   "txtProduto"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtQtde 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8040
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "txtQtde"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdPesq 
         Height          =   315
         Left            =   1560
         Picture         =   "frmCADPRODLISTMAT.frx":00A1
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox cboUnidCons 
         Height          =   315
         Left            =   9480
         TabIndex        =   7
         Text            =   "cboUnidCons"
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cboUnidConvComp 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2640
         TabIndex        =   6
         Text            =   "cboUnidConvComp"
         Top             =   600
         Width           =   4575
      End
      Begin VB.TextBox txtUnidConvComp 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "txtUnidCon"
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblDescProd 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescProd"
         Height          =   315
         Left            =   1920
         TabIndex        =   10
         Top             =   240
         Width           =   6015
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Unidade Conversão:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   630
         Width           =   1740
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10455
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
         Picture         =   "frmCADPRODLISTMAT.frx":01A3
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADPRODLISTMAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public cTipOper         As String
Public iCodigo          As String
Public FILIAL           As Integer
Public strAcesso        As String
Public strMODPAI        As String
Public lngINDICE        As Long
Dim objBLBFunc          As Object
Dim objCADARVPROD       As Object
Dim objPESQPADRAO       As Object


Private Sub cboUnidCons_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboUnidCons, KeyAscii
End Sub

Private Sub cboUnidConvComp_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboUnidConvComp, KeyAscii
End Sub

Private Sub cboUnidConvComp_Validate(Cancel As Boolean)
    If cboUnidConvComp.ListIndex > -1 Then txtUnidConvComp.Text = cboUnidConvComp.ItemData(cboUnidConvComp.ListIndex)
End Sub

Private Sub cmdPesq_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select Case PRO.SGI_PRODUTOTIPO " & vbCrLf
    sSql = sSql & "            When 1 Then replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD))) + '.' + " & vbCrLf
    sSql = sSql & "                        replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE))) + '.' + " & vbCrLf
    sSql = sSql & "                        replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO))) + '.' + " & vbCrLf
    sSql = sSql & "                        (Case " & vbCrLf
    sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Null Then '0' " & vbCrLf
    sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(2),PRO.SGI_DIGVERIF))) End) " & vbCrLf
    sSql = sSql & "            When 0 Then PRO.SGI_CODIGO End As SGI_CODIGO " & vbCrLf
    sSql = sSql & ",SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "         SGI_CADPRODUTO PRO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_IDPRODUTO <> " & arrPROVARV(lngINDICE).lngProdutoID & vbCrLf
    sSql = sSql & "   And SGI_STATUS     = 1"
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "S"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "2000"
    arrCAMPOS(1, 5) = "PRO.SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "PRO.SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Produtos")
    
    If Len(Trim(varRETORNO)) > 0 Then
       txtProduto.Text = varRETORNO
       Call PegaProduto(varRETORNO)
       lblDescProd.Caption = PegaDescProd(txtProduto.Tag)
       Call PegaUniMed
       txtQtde.SetFocus
    End If

End Sub

Private Sub cmdVoltar_Click()
    
    If cTipOper <> "C" Then
       If ConsisteCampos = True Then
            Call PopArray
       End If
    End If
    
    Set objBLBFunc = Nothing
    Set objCADARVPROD = Nothing
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
   Set objCADARVPROD = CreateObject("CADLISTMAT.clsCADLISTMAT")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
   objCADARVPROD.FILIAL = FILIAL
      
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta
      

End Sub


Private Sub Inclui()
    
    Dim I As Integer
    
    Frame4.Enabled = True
    lblDescProd.Caption = ""
    
    Me.Caption = "Cadastro de Lista de Material - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADPRODLISTMAT
    
    objCADARVPROD.PreenchComboUnidade cboUnidCons
    lblDescProd.Caption = ""
    
    If Len(Trim(arrPROVARV(lngINDICE).strPRODUTO)) > 0 Then
    
       txtProduto.Text = arrPROVARV(lngINDICE).strPRODUTO
       txtProduto.Tag = Trim(Str(arrPROVARV(lngINDICE).lngProdutoID))
       lblDescProd.Caption = PegaDescProd(txtProduto.Tag)
       
       txtQtde.Text = Format(arrPROVARV(lngINDICE).curQTDCONS, " #,###0.000")
       
       cboUnidCons.ListIndex = -1
       For I = 0 To (cboUnidCons.ListCount - 1)
           If cboUnidCons.ItemData(I) = arrPROVARV(lngINDICE).lngCodUniMed Then
              cboUnidCons.ListIndex = I
              Exit For
           End If
       Next I
       
    End If
    
    
End Sub

Private Sub txtProduto_GotFocus()
    objBLBFunc.SelecionaCampos txtProduto.Name, frmCADPRODLISTMAT
End Sub

Private Sub txtProduto_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtProduto_Validate(Cancel As Boolean)

   Dim I As Integer

   If Len(Trim(txtProduto.Text)) = 0 Then Exit Sub
   
   Call PegaProduto(txtProduto.Text)
   lblDescProd.Caption = PegaDescProd(txtProduto.Tag)
   If Len(Trim(lblDescProd.Caption)) > 0 Then
        Call PegaUniMed
        txtQtde.SetFocus
   End If
End Sub

Private Sub PegaUniMed()
   
   Dim I As Integer
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "      PRO.* " & vbCrLf
   sSql = sSql & "     ,UNI.* " & vbCrLf
   sSql = sSql & " from " & vbCrLf
   sSql = sSql & "      SGI_CADPRODUTO PRO" & vbCrLf
   sSql = sSql & "     ,SGI_CADUNIMED  UNI" & vbCrLf
   sSql = sSql & "Where " & vbCrLf
   sSql = sSql & "      PRO.SGI_FILIAL     = " & FILIAL & vbCrLf
   sSql = sSql & "  And PRO.SGI_IDPRODUTO  = " & txtProduto.Tag & vbCrLf
   sSql = sSql & "  And UNI.SGI_FILIAL = PRO.SGI_FILIAL " & vbCrLf
   sSql = sSql & "  And UNI.SGI_CODIGO = PRO.SGI_UNIDMEDIDA " & vbCrLf
   
   BREC6.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC6.EOF Then
      cboUnidCons.ListIndex = -1
      For I = 0 To (cboUnidCons.ListCount - 1)
          If BREC6!SGI_UNIDMEDIDA = cboUnidCons.ItemData(I) Then
             cboUnidCons.ListIndex = I
             Exit For
          End If
      Next I
   End If
   BREC6.Close

End Sub

Private Sub txtQtde_GotFocus()
    objBLBFunc.SelecionaCampos txtQtde.Name, frmCADPRODLISTMAT
End Sub

Private Sub txtQtde_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtQtde.Text
End Sub

Private Sub txtQtde_Validate(Cancel As Boolean)

    If Len(Trim(txtQtde.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtQtde.Text) Then
       MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "aviso"
       txtQtde.Text = ""
       txtQtde.SetFocus
       Exit Sub
    End If
    
    txtQtde.Text = Format(txtQtde.Text, "#,####0.0000")

End Sub

Private Sub txtUnidConvComp_GotFocus()
    objBLBFunc.SelecionaCampos txtUnidConvComp.Name, frmCADPRODLISTMAT
End Sub

Private Sub txtUnidConvComp_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtUnidConvComp.Text
End Sub

Private Sub txtUnidConvComp_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtUnidConvComp.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtUnidConvComp.Text) = False Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtUnidConvComp.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboUnidConvComp.ListIndex = -1
    For I = 0 To (cboUnidConvComp.ListCount - 1)
        If cboUnidConvComp.ItemData(I) = CInt(txtUnidConvComp.Text) Then cboUnidConvComp.ListIndex = I
    Next I
    
    If cboUnidConvComp.ListIndex = -1 Then
       MsgBox "Esta tabela de conversão não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtUnidConvComp.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub Consulta()
    
    Dim I As Integer
    
    Me.Caption = "Cadastro de Lista de Material - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADPRODLISTMAT
    
    objCADARVPROD.PreenchComboUnidade cboUnidCons
    
    Frame4.Enabled = False
    lblDescProd.Caption = ""
    
    If Len(Trim(arrPROVARV(lngINDICE).strPRODUTO)) > 0 Then
    
       txtProduto.Text = arrPROVARV(lngINDICE).strPRODUTO
       txtProduto.Tag = arrPROVARV(lngINDICE).lngProdutoID
       txtQtde.Text = Format(arrPROVARV(lngINDICE).curQTDCONS, "#,####0.0000")
       lblDescProd.Caption = PegaDescProd(txtProduto.Tag)
       
       cboUnidCons.ListIndex = -1
       For I = 0 To (cboUnidCons.ListCount - 1)
           If cboUnidCons.ItemData(I) = arrPROVARV(lngINDICE).lngCodUniMed Then
              cboUnidCons.ListIndex = I
              Exit Sub
           End If
       Next I
       
    End If
    
End Sub

Private Sub Altera()
     
    Dim I As Integer
    
    Me.Caption = "Cadastro de Lista de Material - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADPRODLISTMAT
    
    objCADARVPROD.PreenchComboUnidade cboUnidCons
    
    Frame4.Enabled = True
    lblDescProd.Caption = ""
    
    If Len(Trim(arrPROVARV(lngINDICE).strPRODUTO)) > 0 Then
    
       txtProduto.Text = arrPROVARV(lngINDICE).strPRODUTO
       txtProduto.Tag = arrPROVARV(lngINDICE).lngProdutoID
       txtQtde.Text = Format(arrPROVARV(lngINDICE).curQTDCONS, "#,####0.0000")
       lblDescProd.Caption = PegaDescProd(txtProduto.Tag)
       
       cboUnidCons.ListIndex = -1
       For I = 0 To (cboUnidCons.ListCount - 1)
           If cboUnidCons.ItemData(I) = arrPROVARV(lngINDICE).lngCodUniMed Then
              cboUnidCons.ListIndex = I
              Exit For
           End If
       Next I
       
    End If
    
    
End Sub


Private Sub PopArray()

    If cTipOper = "A" Or cTipOper = "I" Then
        If Len(Trim(txtProduto.Text)) > 0 And Len(Trim(txtQtde.Text)) > 0 Then
           If arrPROVARV(lngINDICE).lngProdutoID <> CLng(txtProduto.Tag) Or _
              arrPROVARV(lngINDICE).lngCodUniMed <> cboUnidCons.ItemData(cboUnidCons.ListIndex) Or _
              arrPROVARV(lngINDICE).curQTDCONS <> CCur(txtQtde.Text) Then
              If cTipOper = "A" And arrPROVARV(lngINDICE).intAction2Do = dacEnumUpdateAction_Ignore Then
                 arrPROVARV(lngINDICE).intAction2Do = dacEnumUpdateAction_update
              End If
           End If
             
           arrPROVARV(lngINDICE).strPRODUTO = Trim(txtProduto.Text)
           arrPROVARV(lngINDICE).lngProdutoID = CLng(txtProduto.Tag)
           arrPROVARV(lngINDICE).lngCodUniMed = cboUnidCons.ItemData(cboUnidCons.ListIndex)
           arrPROVARV(lngINDICE).strUNIDADE = cboUnidCons.List(cboUnidCons.ListIndex)
           arrPROVARV(lngINDICE).curQTDCONS = CCur(txtQtde.Text)
           arrPROVARV(lngINDICE).lngTipo = PegaTipo
        End If
    End If

End Sub

Private Function ConsisteCampos() As Boolean
    ConsisteCampos = False
    
    If lngINDICE <> 1 Then
        If Len(Trim(txtProduto.Text)) = 0 Then
           ''MsgBox "Informe o Produto da Lista !!!", vbOKOnly + vbExclamation, "Aviso"
           ''txtProduto.SetFocus
           Exit Function
        End If
        If cboUnidCons.ListIndex < -1 Then
           MsgBox "Informe a unidade de Medida de consumo !!!", vbOKOnly + vbExclamation, "Aviso"
           cboUnidCons.SetFocus
           Exit Function
        End If
    End If
    
    ConsisteCampos = True
End Function

Private Function PegaTipo() As Long

    PegaTipo = 0
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = '" & Trim(txtProduto.Text) & "'"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaTipo = BREC!SGI_PRODUTOTIPO
    BREC.Close
    
End Function

Private Sub PegaProduto(strPRODUTO As String)

    sSql = ""
    
    sSql = "Select SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & ",Case PRO.SGI_PRODUTOTIPO " & vbCrLf
    sSql = sSql & "            When 1 Then replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD))) + '.' + " & vbCrLf
    sSql = sSql & "                        replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE))) + '.' + " & vbCrLf
    sSql = sSql & "                        replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO))) + '.' + " & vbCrLf
    sSql = sSql & "                        (Case " & vbCrLf
    sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Null Then '0' " & vbCrLf
    sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(2),PRO.SGI_DIGVERIF))) End) " & vbCrLf
    sSql = sSql & "            When 0 Then PRO.SGI_CODIGO End As SGI_CODIGO " & vbCrLf
    sSql = sSql & ",SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "         SGI_CADPRODUTO PRO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_STATUS = 1" & vbCrLf
    sSql = sSql & "   And (Case PRO.SGI_PRODUTOTIPO " & vbCrLf
    sSql = sSql & "            When 1 Then replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD))) + '.' + " & vbCrLf
    sSql = sSql & "                        replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE))) + '.' + " & vbCrLf
    sSql = sSql & "                        replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO))) + '.' + " & vbCrLf
    sSql = sSql & "                        (Case " & vbCrLf
    sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Null Then '0' " & vbCrLf
    sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(2),PRO.SGI_DIGVERIF))) End) " & vbCrLf
    sSql = sSql & "            When 0 Then PRO.SGI_CODIGO End) Like '" & Trim(strPRODUTO) & "%'"

    BREC4.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC4.EOF() Then
        txtProduto.Tag = Trim(Str(BREC4!SGI_IDPRODUTO))
    Else
        MsgBox "Produto Inexistente !!!", vbOKOnly + vbExclamation, "Aviso de Sistema"
        If Len(Trim(Str(arrPROVARV(lngINDICE).lngProdutoID))) > 0 Then
            txtProduto.Tag = Trim(Str(arrPROVARV(lngINDICE).lngProdutoID))
            txtProduto.Text = Trim(arrPROVARV(lngINDICE).strPRODUTO)
        End If
    End If
    BREC4.Close
    
End Sub

Private Function PegaDescProd(strProdutoID As String) As String
    PegaDescProd = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_IDPRODUTO = " & strProdutoID
    
    BREC5.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC5.EOF() Then PegaDescProd = Trim(BREC5!SGI_DESCRICAO)
    BREC5.Close
    
End Function
