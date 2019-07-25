VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmCADORDEMCORES 
   Caption         =   "Cadastro de Ordem de Cores"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10035
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   10035
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Height          =   300
      Left            =   9600
      Picture         =   "frmCADORDEMCORES.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1920
      Width           =   300
   End
   Begin VB.CommandButton Command3 
      Height          =   300
      Left            =   9600
      Picture         =   "frmCADORDEMCORES.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2280
      Width           =   300
   End
   Begin VSFlex8LCtl.VSFlexGrid grdCORES 
      Height          =   4575
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   9375
      _cx             =   16536
      _cy             =   8070
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
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   9855
      Begin VB.Frame Frame3 
         Caption         =   "[ Defalt ]"
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
         Left            =   3360
         TabIndex        =   7
         Top             =   120
         Width           =   1935
         Begin VB.OptionButton optDEFALT 
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
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton optDEFALT 
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
            Left            =   960
            TabIndex        =   8
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "txtCodigo"
         Top             =   240
         Width           =   1215
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9855
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
         Left            =   960
         Picture         =   "frmCADORDEMCORES.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Left            =   1800
         Picture         =   "frmCADORDEMCORES.frx":07C6
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Picture         =   "frmCADORDEMCORES.frx":08C8
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADORDEMCORES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public cTipOper         As String
Public iCodigo          As Long
Public FILIAL           As Integer
Public strAcesso        As String
Public strMODPAI        As String
Public strUsuario       As String
Public lngCodVendedor   As Long
Public lngCodUsuario    As Long
Public intFILIALPED     As Integer

Dim objBLBFunc          As Object
Dim objCADORDEMCORES    As Object
Dim objPESQPADRAO       As Object
Dim strModulo           As String
Dim strCAPTION          As String
Dim arrCORES            As Variant

'' -----------------------------------------------------------------------------------
Const conCOL_Cores_Ordem                       As Integer = 0
Const conCOL_Cores_CodCor                      As Integer = 1
Const conCOL_Cores_PesqCor                     As Integer = 2
Const conCOL_Cores_DescCor                     As Integer = 3
Const conCOL_Cores_IDProd                      As Integer = 4
Const conCOL_Cores_FormatString                As String = "=Ordem|Cód.Cor|...|Descrição da Cor|IDProduto"
Const conColumnsIn_Cores                       As Integer = 5

Private Sub CmdSalva_Click()

    Dim I As Long
    
    Call objBLBFunc.removeLinhaVazia(grdCORES, conCOL_Cores_CodCor)

    If ValidaCampos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADORDEMCORES.CODIGO = objBLBFunc.Gera_Codigo(Trim(Me.Name), FILIAL, Linha)

    If optDEFALT(0).Value = True Then objCADORDEMCORES.DEFALT = 0
    If optDEFALT(1).Value = True Then objCADORDEMCORES.DEFALT = 1
    
    arrCORES = Empty
    With grdCORES
        If (.Rows - 1) > 0 Then
            ReDim arrCORES(1 To (.Rows - 1), 1 To 2) As String
            For I = 1 To (.Rows - 1)
                arrCORES(I, 1) = .Cell(flexcpText, I, conCOL_Cores_Ordem)
                arrCORES(I, 2) = .Cell(flexcpText, I, conCOL_Cores_IDProd)
            Next I
        End If
    End With
    objCADORDEMCORES.CORES = arrCORES
    
    If objCADORDEMCORES.GRAVA(cTipOper) = False Then Exit Sub
          
    MsgBox "A ordem das Cores foram " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
    If cTipOper = "I" Then Unload Me

End Sub

Private Sub Command2_Click()
    Call IncRegGridCores
End Sub

Private Sub Command3_Click()
    If cTipOper = "C" Then Exit Sub
    Call objBLBFunc.ExclLinhaGrid(grdCORES, grdCORES.Row)
    Call Controy_Ordem
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objCADORDEMCORES = CreateObject("CADORDEMCORES.clsCADORDEMCORES")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
    
    If adoBanco_Dados.State = 0 Then Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)
    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
   
    strCAPTION = "Cadastro de Ordem de Cores - "
   
    objCADORDEMCORES.FILIAL = FILIAL
   
    Call IniciaForm

End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objCADORDEMCORES = Nothing
    Set objPESQPADRAO = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

Private Sub IniciaForm()
    
    Call objBLBFunc.MudaBotoes(CmdSalva, cmdAltera, cTipOper)
    Call objBLBFunc.MudaCaption(Me, strCAPTION, cTipOper)
    Call objBLBFunc.LimpaCampos(Me)
    
    Call DesabilitaCampos(Trim(cTipOper))
    Call ConfGrd
    
    If cTipOper = "I" Then iCodigo = 0
    objCADORDEMCORES.CODIGO = iCodigo
    optDEFALT(0).Value = True
    
    Call CarregaCampos
    
End Sub

Private Sub DesabilitaCampos(strTipOper As String)
    If strTipOper = "I" Or strTipOper = "A" Then
        Frame2.Enabled = True
    ElseIf strTipOper = "C" Then
        Frame2.Enabled = False
    End If
End Sub

Private Sub grdCORES_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With grdCORES
        
        If (.Rows - 1) = 0 Then Exit Sub
        If Row = 0 Then Exit Sub
        
        Select Case Col
               Case conCOL_Cores_Ordem
        End Select
    
    End With
End Sub

Private Sub grdCORES_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    With grdCORES
        Select Case Col
               Case conCOL_Cores_Ordem, _
                    conCOL_Cores_DescCor, _
                    conCOL_Cores_IDProd
                    Cancel = True
               Case conCOL_Cores_CodCor, _
                    conCOL_Cores_PesqCor
                    If cTipOper = "C" Then Cancel = True
               Case Else
                   .ComboList = ""
               End Select
    End With
    
    Exit Sub

End Sub

Private Sub grdCORES_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    With grdCORES
        If (.Rows - 1) = 0 Then Exit Sub
        
        Select Case Col
            Case conCOL_Cores_PesqCor
                
                Call PesqCores(Row)
                
                Exit Sub
                
        End Select
    End With

End Sub

Private Sub grdCORES_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdCORES
          Select Case Col
                    Case conCOL_Cores_CodCor
                         KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
          End Select
     End With
End Sub

Private Sub grdCORES_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

     With grdCORES
          Select Case Col
                 Case conCOL_Cores_CodCor
                        If .EditText = Empty Then Exit Sub
                        
                        If objBLBFunc.FcVerifItensRepetidos(grdCORES, Row, conCOL_Cores_CodCor, .EditText) = False Then
                           MsgBox "ATENÇÃO - A Cor ja foi relacionada na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                           grdCORES.Cell(flexcpText, Row, conCOL_Cores_CodCor) = Empty
                           grdCORES.Cell(flexcpText, Row, conCOL_Cores_DescCor) = Empty
                           grdCORES.Cell(flexcpText, Row, conCOL_Cores_IDProd) = Empty
                           Cancel = True
                           Exit Sub
                        End If
                        
                        .Cell(flexcpText, Row, conCOL_Cores_IDProd) = PegaIDProduto(.EditText)
                        If Len(Trim(.Cell(flexcpText, Row, conCOL_Cores_IDProd))) = 0 Then
                           MsgBox "ATENÇÂO - Cor não existe !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, Col) = Empty
                           Cancel = True
                           Exit Sub
                        End If
                        
                        .Cell(flexcpText, Row, conCOL_Cores_CodCor) = .EditText
                        .Cell(flexcpText, Row, conCOL_Cores_DescCor) = PegaDescrProduto(.Cell(flexcpText, Row, conCOL_Cores_IDProd))
                        
                        If Len(Trim(.Cell(flexcpText, Row, conCOL_Cores_DescCor))) = 0 Then
                           MsgBox "ATENÇÃO - Cor não existe !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, Col) = Empty
                           .Cell(flexcpText, Row, conCOL_Cores_DescCor) = Empty
                           .Cell(flexcpText, Row, conCOL_Cores_IDProd) = Empty
                           Cancel = True
                           Exit Sub
                        End If
          End Select
     End With

End Sub

Private Sub PegaDescTabelas(strCAMPOPESQ As String, StrCampoRetorno As String, strTABELA As String, strCODIGO As String, lblLabel As Label)

    lblLabel.Caption = ""
    
    If Len(Trim(strCODIGO)) = 0 Then Exit Sub
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoRetorno) & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & Trim(strTABELA) & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       " & Trim(strCAMPOPESQ) & " = " & Trim(strCODIGO)
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then
       lblLabel.Caption = Trim(BREC10(Trim(StrCampoRetorno)))
    Else
       MsgBox "Registro Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC10.Close
    
End Sub

Private Sub ConfGrd()

    With grdCORES

       .Cols = conColumnsIn_Cores
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_Cores_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeNone
       
       .Cell(flexcpData, 0, conCOL_Cores_Ordem) = ""
       .ColDataType(conCOL_Cores_Ordem) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Cores_CodCor) = ""
       .ColDataType(conCOL_Cores_CodCor) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Cores_PesqCor) = ""
       .ColDataType(conCOL_Cores_PesqCor) = flexDTString
       .ColComboList(conCOL_Cores_PesqCor) = "..."
       
       .Cell(flexcpData, 0, conCOL_Cores_DescCor) = ""
       .ColDataType(conCOL_Cores_DescCor) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_Cores_IDProd) = ""
       .ColDataType(conCOL_Cores_IDProd) = flexDTLong
       
       .ColWidth(conCOL_Cores_Ordem) = 1000
       .ColWidth(conCOL_Cores_CodCor) = 1200
       .ColWidth(conCOL_Cores_PesqCor) = 300
       .ColWidth(conCOL_Cores_DescCor) = 5000
       .ColWidth(conCOL_Cores_IDProd) = 0
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack

    End With
    
End Sub


Private Sub PesqCores(lngROW As Long)
    With grdCORES
        If (.Rows - 1) = 0 Then Exit Sub
            
        If cTipOper = "C" Then Exit Sub
        
        Dim lngLinha                    As Long
        ReDim arrCAMPOS(1 To 2, 1 To 5) As String
        ReDim arrTABELA(1 To 1) As String
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       * " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL      = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_PRODUTOTIPO = 0"
        
        arrTABELA(1) = sSql
        
        arrCAMPOS(1, 1) = "SGI_CODIGO"
        arrCAMPOS(1, 2) = "S"
        arrCAMPOS(1, 3) = "Código"
        arrCAMPOS(1, 4) = "1500"
        arrCAMPOS(1, 5) = "SGI_CODIGO"
        
        arrCAMPOS(2, 1) = "SGI_DESCRICAO"
        arrCAMPOS(2, 2) = "S"
        arrCAMPOS(2, 3) = "Nome"
        arrCAMPOS(2, 4) = "5000"
        arrCAMPOS(2, 5) = "SGI_DESCRICAO"
        
        varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Cores")
        
        If Len(Trim(varRETORNO)) = 0 Then Exit Sub
        
        If objBLBFunc.FcVerifItensRepetidos(grdCORES, .Row, conCOL_Cores_CodCor, Trim(varRETORNO)) = False Then
           MsgBox "ATENÇÃO - A Cor ja foi relacionada na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
           grdCORES.Cell(flexcpText, .Row, conCOL_Cores_CodCor) = Empty
           grdCORES.Cell(flexcpText, .Row, conCOL_Cores_DescCor) = Empty
           grdCORES.Cell(flexcpText, .Row, conCOL_Cores_IDProd) = Empty
           Exit Sub
        End If
        
        .Cell(flexcpText, lngROW, conCOL_Cores_IDProd) = PegaIDProduto(varRETORNO)
        .Cell(flexcpText, lngROW, conCOL_Cores_CodCor) = varRETORNO
        .Cell(flexcpText, lngROW, conCOL_Cores_DescCor) = PegaDescrProduto(.Cell(flexcpText, lngROW, conCOL_Cores_IDProd))
    
    End With
End Sub

Private Function PegaIDProduto(strCodProduto As String) As String

    PegaIDProduto = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_CODIGO = '" & Trim(UCase(strCodProduto)) & "'" & vbCrLf
    sSql = sSql & "   And SGI_FILIAL = " & FILIAL

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaIDProduto = BREC!SGI_IDPRODUTO
    BREC.Close
    
End Function


Private Function PegaDescrProduto(strIDProduto As String) As String
    
    PegaDescrProduto = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_IDPRODUTO = " & strIDProduto
    
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC2.EOF Then PegaDescrProduto = BREC2!SGI_DESCRICAO
    BREC2.Close
    
End Function


Private Sub IncRegGridCores()
   
    If cTipOper = "C" Then Exit Sub
    
    If objBLBFunc.FcExisteLinhaVazia(grdCORES, conCOL_Cores_CodCor) = False Then Exit Sub
    
    With grdCORES
        .AddItem "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 ""
    End With
    
    Call Controy_Ordem
   
End Sub

Private Sub Controy_Ordem()

    Dim intINDICE As Long

    With grdCORES
        If (.Rows - 1) > 0 Then
            For intINDICE = 1 To (.Rows - 1)
                .Cell(flexcpText, intINDICE, conCOL_Cores_Ordem) = intINDICE
            Next intINDICE
        End If
    End With

End Sub

Private Function ValidaCampos() As Boolean

     ValidaCampos = False
     
     Dim boolExisteDefault As Boolean
     
     If (grdCORES.Rows - 1) = 0 Then
        MsgBox "ATENÇÃO - Não foi informado nenhuma cor !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Function
     End If
     
     sSql = ""
     boolExisteDefault = True
     
     If optDEFALT(1).Value = True Then
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "        Count(SGI_DEFALT) As SGI_QTDE" & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "        SGI_CADORDEMCORES " & vbCrLf
        sSql = sSql & "  Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL     = " & FILIAL & vbCrLf
        If cTipOper = "A" Then
           sSql = sSql & "   And SGI_CODIGO     <> " & objCADORDEMCORES.CODIGO & vbCrLf
        End If
        sSql = sSql & "   And SGI_DEFALT     = 1"
            
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC.EOF() Then
           If BREC!SGI_QTDE > 0 Then
               MsgBox "ATENÇÃO - já existe uma maquina como default não é possivel cadastrar !!!", vbOKOnly + vbExclamation, "Aviso"
               boolExisteDefault = False
           End If
        End If
        BREC.Close
        
        If boolExisteDefault = False Then Exit Function
     
     End If
     
     ValidaCampos = True
     
End Function


Private Sub CarregaCampos()

On Error GoTo Err_CarregaCampos
    
    If objCADORDEMCORES.Carrega_campos = True Then
    
        txtCodigo.Text = objCADORDEMCORES.CODIGO
        optDEFALT(objCADORDEMCORES.DEFALT).Value = True
        
        Call PopGrdLanctos
    End If
    
    Exit Sub

Err_CarregaCampos:
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : CarregaCampos", Me.Name, "CarregaCampos")
    
End Sub


Private Sub PopGrdLanctos()

    Dim I As Integer
    
    arrCORES = objCADORDEMCORES.CORES
    If IsArray(arrCORES) Then
        With grdCORES
            For I = 1 To UBound(arrCORES)
                .AddItem arrCORES(I, 1) & vbTab & _
                         arrCORES(I, 2) & vbTab & _
                         "" & vbTab & _
                         arrCORES(I, 3) & vbTab & _
                         arrCORES(I, 4)
        
            Next I
        End With
    End If

End Sub
