VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmCADREQMAT 
   Caption         =   "Cadastro de Requisição de Materiais"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11940
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   11940
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Height          =   6135
      Left            =   0
      TabIndex        =   15
      Top             =   2280
      Width           =   11895
      Begin VSFlex8LCtl.VSFlexGrid grdITREQMAT 
         Height          =   5775
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   11175
         _cx             =   19711
         _cy             =   10186
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.CommandButton cmbGravPagto 
         Height          =   315
         Left            =   11400
         Picture         =   "frmCADREQMAT.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   0
      TabIndex        =   8
      Top             =   960
      Width           =   11895
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   2760
         Picture         =   "frmCADREQMAT.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton cmdFornec 
         Height          =   315
         Left            =   2760
         Picture         =   "frmCADREQMAT.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtCODUSUARIO 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   3
         Text            =   "txtCODUSUARIO"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtCODDEPTO 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   2
         Text            =   "txtCODDEPTO"
         Top             =   600
         Width           =   1215
      End
      Begin MSMask.MaskEdBox mskDTREQ 
         Height          =   285
         Left            =   10560
         TabIndex        =   1
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblDESCUSUARIO 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDESCUSUARIO"
         Height          =   285
         Left            =   3120
         TabIndex        =   19
         Top             =   960
         Width           =   6375
      End
      Begin VB.Label lblDESCDEPTO 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDESCDEPTO"
         Height          =   285
         Left            =   3120
         TabIndex        =   18
         Top             =   600
         Width           =   6375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuário.:"
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
         Left            =   650
         TabIndex        =   12
         Top             =   960
         Width           =   780
      End
      Begin VB.Label lblCODREQ 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblCODREQ"
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
         Height          =   285
         Index           =   3
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Departamento.:"
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
         TabIndex        =   11
         Top             =   600
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Req.:"
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
         Left            =   9480
         TabIndex        =   10
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código Req.:"
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
         Left            =   320
         TabIndex        =   9
         Top             =   240
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11895
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
         Picture         =   "frmCADREQMAT.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "frmCADREQMAT.frx":0838
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "frmCADREQMAT.frx":093A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmCADREQMAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho      As String
Public Linha         As Variant
Public cTipOper      As String
Public iCodigo       As Integer
Public FILIAL        As Integer
Public strAcesso     As String
Public strMODPAI     As String
Public strUSUARIO    As String
Dim objBLBFunc       As Object
Dim objCADREQMAT     As Object
Dim objPESQPADRAO    As Object
Dim arrITENSREQ      As Variant

Const conCOL_Prod_ID                              As Integer = 0
Const conCOL_Prod_Rotulo                          As Integer = 1
Const conCOL_Prod_PesqRot                         As Integer = 2
Const conCOL_Prod_DescrProd                       As Integer = 3
Const conCOL_Prod_Qtde                            As Integer = 4
Const conCOL_Prod_QtdeAtend                       As Integer = 5
Const conCOL_Prod_Saldo                           As Integer = 6
Const conCOL_Prod_Status                          As Integer = 7
Const conCOL_Prod_FormatString                    As String = "=ID|Rótulo|...|Descrição|Qtde|Qtde.Atend|Saldo|Status"
Const conColumnsIn_Prod                           As Integer = 8

Private Sub cmbGravPagto_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncGridReq
End Sub

Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    If VerifReqSai = True Then
       MsgBox "Existe requisição de saidas já emitida !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Sub
    End If
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame4.Enabled = True
   
    Me.Caption = "Cadastro de Requisição de Materiais - [ ALTERAÇÃO ]"
    cTipOper = "A"
    
    txtCODDEPTO.SetFocus

End Sub

Private Sub cmdFornec_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    arrTABELA(1) = "Select * from SGI_CADDEPTO"
    
    arrCAMPOS(1, 1) = "SGI_CODDEPTO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODDEPTO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "3000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Departamentos")
    
    If Len(Trim(varRETORNO)) = 0 Then Exit Sub
    
    txtCODDEPTO.Text = varRETORNO
    Call PegaDescTabelas("SGI_CODDEPTO", "SGI_DESCRICAO", "SGI_CADDEPTO", varRETORNO, lblDESCDEPTO)
    txtCODDEPTO.SetFocus


End Sub

Private Sub CmdSalva_Click()
    
    Dim I As Integer
    
    If Valida_Campos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADREQMAT.CADREQCOD = objCADREQMAT.Gera_Codigo(Me.Name)
    
    objCADREQMAT.CADDEPCOD = CLng(txtCODDEPTO.Text)
    objCADREQMAT.CADUSUCOD = CLng(txtCODUSUARIO.Text)
    objCADREQMAT.CADDTREQ = CDate(mskDTREQ.Text)
    
    ''If (flxITREQMAT.Rows - 1) > 0 Then
    ''   ReDim arrITENSREQ(1 To (flxITREQMAT.Rows - 1), 1 To 4) As Variant
    ''   For I = 1 To (flxITREQMAT.Rows - 1)
    ''       arrITENSREQ(I, 1) = flxITREQMAT.TextMatrix(I, 0)
    ''       arrITENSREQ(I, 2) = flxITREQMAT.TextMatrix(I, 1)
    ''       arrITENSREQ(I, 3) = flxITREQMAT.TextMatrix(I, 2)
    ''       arrITENSREQ(I, 4) = flxITREQMAT.TextMatrix(I, 3)
    ''   Next I
    ''   objCADREQMAT.ITENSREQ = arrITENSREQ
    ''End If
    
    If objCADREQMAT.GRAVA(cTipOper) = False Then Exit Sub
          
    MsgBox "A requisição foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", IIf(cTipOper = "L", "liberado", ""))) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
    Call Destroy_Objeto
    Unload Me

End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    arrTABELA(1) = "Select * from SGI_USUARIO"
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_NOME"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "3000"
    arrCAMPOS(2, 5) = "SGI_NOME"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Usuários")
    
    If Len(Trim(varRETORNO)) = 0 Then Exit Sub
    txtCODUSUARIO.Text = varRETORNO
    Call PegaDescTabelas("SGI_CODIGO", "SGI_NOME", "SGI_USUARIO", varRETORNO, lblDESCUSUARIO)
    txtCODUSUARIO.SetFocus

End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADREQMAT = CreateObject("CADREQMAT.clsCADREQMAT")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
   objCADREQMAT.FILIAL = FILIAL
   
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
    Frame4.Enabled = True
    
    mskDTREQ.Enabled = True
    
    Me.Caption = "Cadastro de Requisição de Materiais - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADREQMAT
    lblCODREQ(3).Caption = ""
    
    mskDTREQ.Text = Format(Date, "DD/MM/YYYY")
    
    Call ConfGridItReq
    Call LimpaCamposLabel
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

Private Sub mskDTREQ_GotFocus()
    objBLBFunc.SelecionaCampos mskDTREQ.Name, frmCADREQMAT
End Sub

Private Sub txtCODDEPTO_GotFocus()
    objBLBFunc.SelecionaCampos txtCODDEPTO.Name, frmCADREQMAT
End Sub

Private Sub txtCODDEPTO_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODDEPTO.Text
End Sub

Private Sub txtCODDEPTO_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODDEPTO.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODDEPTO.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODDEPTO.Text = ""
       Cancel = True
       Exit Sub
    End If
    
End Sub



Private Sub txtCODUSUARIO_GotFocus()
    objBLBFunc.SelecionaCampos txtCODUSUARIO.Name, frmCADREQMAT
End Sub

Private Sub txtCODUSUARIO_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODUSUARIO.Text
End Sub

Private Sub txtCODUSUARIO_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODUSUARIO.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODUSUARIO.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODUSUARIO.Text = ""
       Cancel = True
       Exit Sub
    End If
    
End Sub

Private Sub ConfGridItReq()

    
    
    
    
    
    
    ''If VerifReqSai = True Then
    ''End If
    

End Sub

Private Sub IncGridReq()

End Sub


Private Function Valida_Campos() As Boolean

     Valida_Campos = False
     
     If Len(Trim(txtCODDEPTO.Text)) = 0 Then
        MsgBox "O departamento não pode ser vázio !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCODDEPTO.SetFocus
        Exit Function
     End If
     
     If Len(Trim(txtCODUSUARIO.Text)) = 0 Then
        MsgBox "O usuário não pode ser vázio !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCODUSUARIO.SetFocus
        Exit Function
     End If
     
     If (grdITREQMAT.Rows - 1) = 0 Then
        MsgBox "Não Foi Informado Itens para a requisição !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Function
     End If
     
     Valida_Campos = True

End Function

Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
    Frame4.Enabled = True
   
    Me.Caption = "Cadastro de Requisição de Materiais - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADREQMAT
    objCADREQMAT.CADREQCOD = iCodigo
    
    ConfGridItReq
    Call LimpaCamposLabel
    
    If objCADREQMAT.Carrega_campos = True Then
    
      lblCODREQ(3).Caption = objCADREQMAT.CADREQCOD
      txtCODDEPTO.Text = objCADREQMAT.CADDEPCOD
      txtCODUSUARIO.Text = objCADREQMAT.CADUSUCOD
      mskDTREQ.Text = Format(objCADREQMAT.CADDTREQ, "DD/MM/YYYY")
      arrITENSREQ = objCADREQMAT.ITENSREQ
      
       
    End If

End Sub

Private Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame4.Enabled = True
    
    mskDTREQ.Enabled = False
   
    Me.Caption = "Cadastro de Requisição de Materiais - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADREQMAT
    objCADREQMAT.CADREQCOD = iCodigo
    
    ConfGridItReq
    
    Call LimpaCamposLabel
    
    If objCADREQMAT.Carrega_campos = True Then
    
      lblCODREQ(3).Caption = objCADREQMAT.CADREQCOD
      txtCODDEPTO.Text = objCADREQMAT.CADDEPCOD
      txtCODUSUARIO.Text = objCADREQMAT.CADUSUCOD
      mskDTREQ.Text = Format(objCADREQMAT.CADDTREQ, "DD/MM/YYYY")
      arrITENSREQ = objCADREQMAT.ITENSREQ
      
      Call PopGrdReg
    
    End If

End Sub

Private Function VerifReqSai() As Boolean

    VerifReqSai = False
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADITREQSAIMAT " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODREQ = " & objCADREQMAT.CADREQCOD

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then VerifReqSai = True
    BREC.Close
    
End Function

Private Sub LimpaCamposLabel()
    lblDESCDEPTO.Caption = ""
    lblDESCUSUARIO.Caption = ""
End Sub

Private Sub PegaDescTabelas(StrCampoPesq As String, StrCampoRetorno As String, strTabela As String, StrCodigo As String, lblLabel As Label)

    lblLabel.Caption = ""
    
    If Len(Trim(StrCodigo)) = 0 Then Exit Sub
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoRetorno) & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & Trim(strTabela) & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoPesq) & " = " & Trim(StrCodigo)
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then
       lblLabel.Caption = Trim(BREC10(Trim(StrCampoRetorno)))
    Else
       MsgBox "Registro Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC10.Close
    
End Sub


Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objCADREQMAT = Nothing
    Set objPESQPADRAO = Nothing
End Sub

Private Sub PopGrdReg()
      Dim I As Integer
      '' Itens da Requisição
      If IsEmpty(arrITENSREQ) = False Then
         With grdITREQMAT
            For I = 1 To UBound(arrITENSREQ)
                If arrITENSREQ(I, 5) > 0 Then
                   .AddItem arrITENSREQ(I, 1) & vbTab & _
                            arrITENSREQ(I, 2) & vbTab & _
                            arrITENSREQ(I, 3) & vbTab & _
                            Format(arrITENSREQ(I, 4), "#,###0.000") & vbTab & _
                            Format(arrITENSREQ(I, 5), "#,###0.000") & vbTab & _
                            Format(arrITENSREQ(I, 6), "#,###0.000") & vbTab & _
                            IIf(arrITENSREQ(I, 6) > 0, "Aberto", "Atendido")
             
                Else
                   .AddItem arrITENSREQ(I, 1) & vbTab & _
                            arrITENSREQ(I, 2) & vbTab & _
                            arrITENSREQ(I, 3) & vbTab & _
                            Format(arrITENSREQ(I, 4), "#,###0.000")
                End If
            Next I
         End With
      End If
End Sub
