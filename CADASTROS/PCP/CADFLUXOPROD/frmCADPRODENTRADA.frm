VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmCADPRODENTRADA 
   Caption         =   "Cadastro de Produtos de Entrada"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   11820
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "[ Produtos que serão produzidos ]"
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
      Height          =   3255
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   11775
      Begin VSFlex8LCtl.VSFlexGrid grdPRODUTOS 
         Height          =   2895
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   11535
         _cx             =   20346
         _cy             =   5106
         Appearance      =   1
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
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11775
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
         Picture         =   "frmCADPRODENTRADA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "[ Produtos de Entrada ]"
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
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   4080
      Width           =   11775
      Begin VB.CommandButton cmdExcIten 
         Height          =   315
         Left            =   11280
         Picture         =   "frmCADPRODENTRADA.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdIncIten 
         Height          =   315
         Left            =   11280
         Picture         =   "frmCADPRODENTRADA.frx":068C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
      Begin VSFlex8LCtl.VSFlexGrid grdPRODENTR 
         Height          =   4215
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   11055
         _cx             =   19500
         _cy             =   7435
         Appearance      =   1
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
   End
End
Attribute VB_Name = "frmCADPRODENTRADA"
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
Public strPRODUTO       As String
Public strIDProduto     As String
Public lngIndice        As Long

Dim objBLBFunc          As Object
Dim objCADPRODENTRADA   As Object
Dim objPESQPADRAO       As Object

Const conCOL_SonProd_CodProd           As Integer = 0
Const conCOL_SonProd_Desc_Prod         As Integer = 1
Const conCOL_SonProd_UnidMed           As Integer = 2
Const conCOL_SonProd_IDProduto         As Integer = 3
Const conCOL_SonProd_FormatString      As String = "= Cod. Produto|Descrição do Produto|Unidade|IDProduto"
Const conColumnsIn_SonProd             As Integer = 4

Const conCOL_SonProdEnt_CodProd        As Integer = 0
Const conCOL_SonProdEnt_PesqProd       As Integer = 1
Const conCOL_SonProdEnt_Desc_Prod      As Integer = 2
Const conCOL_SonProdEnt_UnidMed        As Integer = 3
Const conCOL_SonProdEnt_Estoque        As Integer = 4
Const conCOL_SonProdEnt_Cadencia       As Integer = 5
Const conCOL_SonProdEnt_Pai            As Integer = 6
Const conCOL_SonProdEnt_FormatString   As String = "= Cod. Produto|...|Descrição do Produto|Unidade|Estoque|Cadência|Pai"
Const conColumnsIn_SonProdEnt          As Integer = 7

Private Sub cmdExcIten_Click()
    If cTipOper = "I" Or cTipOper = "A" Then Call objBLBFunc.ExclLinhaGrid(grdPRODENTR, grdPRODENTR.Row)
End Sub

Private Sub cmdIncIten_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncRegGrid
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub InitGridProdEntr()

  With grdPRODENTR
       .Cols = conColumnsIn_SonProdEnt
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonProdEnt_FormatString
       .AutoSizeMouse = True
       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonProdEnt_CodProd) = ""
       .ColDataType(conCOL_SonProdEnt_CodProd) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonProdEnt_PesqProd) = ""
       .ColDataType(conCOL_SonProdEnt_PesqProd) = flexDTString
       .ColComboList(conCOL_SonProdEnt_PesqProd) = "..."
       
       .Cell(flexcpData, 0, conCOL_SonProdEnt_Desc_Prod) = ""
       .ColDataType(conCOL_SonProdEnt_Desc_Prod) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonProdEnt_UnidMed) = ""
       .ColDataType(conCOL_SonProdEnt_UnidMed) = flexDTString
       .ColComboList(conCOL_SonProdEnt_UnidMed) = objCADPRODENTRADA.PreenchComboUnidade
       
       .Cell(flexcpData, 0, conCOL_SonProdEnt_Estoque) = ""
       .ColDataType(conCOL_SonProdEnt_Estoque) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonProdEnt_Cadencia) = ""
       .ColDataType(conCOL_SonProdEnt_Cadencia) = flexDTBoolean
       .ColFormat(conCOL_SonProdEnt_Cadencia) = "Sim;Não"
       
       .Cell(flexcpData, 0, conCOL_SonProdEnt_Pai) = ""
       .ColDataType(conCOL_SonProdEnt_Pai) = flexDTString
       
       .ColWidth(conCOL_SonProdEnt_CodProd) = 1500
       .ColWidth(conCOL_SonProdEnt_PesqProd) = 300
       .ColWidth(conCOL_SonProdEnt_Desc_Prod) = 4000
       .ColWidth(conCOL_SonProdEnt_UnidMed) = 1500
       .ColWidth(conCOL_SonProdEnt_Estoque) = 1500
       .ColWidth(conCOL_SonProdEnt_Cadencia) = 1500
       
       .ColHidden(conCOL_SonProdEnt_Pai) = True
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
  End With

End Sub

Private Sub InitGridProd()

  With grdPRODUTOS
       .Cols = conColumnsIn_SonProd
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonProd_FormatString
       .AutoSizeMouse = True
       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonProd_CodProd) = ""
       .ColDataType(conCOL_SonProd_CodProd) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonProd_Desc_Prod) = ""
       .ColDataType(conCOL_SonProd_Desc_Prod) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonProd_UnidMed) = ""
       .ColDataType(conCOL_SonProd_UnidMed) = flexDTString
       .ColComboList(conCOL_SonProd_UnidMed) = objCADPRODENTRADA.PreenchComboUnidade
       
       .Cell(flexcpData, 0, conCOL_SonProd_IDProduto) = ""
       .ColDataType(conCOL_SonProd_IDProduto) = flexDTString
       
       .ColWidth(conCOL_SonProd_CodProd) = 1500
       .ColWidth(conCOL_SonProd_Desc_Prod) = 4000
       .ColWidth(conCOL_SonProd_UnidMed) = 1500
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
  End With

End Sub


Private Sub Form_Load()
   
   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADPRODENTRADA = CreateObject("CADFLUXOPROD.clsCADFLUXOPROD")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
   objCADPRODENTRADA.FILIAL = FILIAL
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

    
End Sub

Private Sub Inclui()

    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de Produtos de Entrada - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADPRODENTRADA
    
    Call InitGridProd
    Call InitGridProdEntr
    
    Call PopGrdProdutos
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DestroyObjeto
End Sub

Private Sub grdPRODENTR_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Select Case Col
    Case conCOL_SonProdEnt_CodProd
         If grdPRODENTR.Cell(flexcpText, Row, conCOL_SonProdEnt_CodProd) <> Empty Then
            grdPRODENTR.Cell(flexcpText, Row, conCOL_SonProdEnt_Desc_Prod) = PegaDescrProduto(Trim(grdPRODENTR.Cell(flexcpText, Row, conCOL_SonProdEnt_CodProd)))
            grdPRODENTR.Cell(flexcpText, Row, conCOL_SonProdEnt_UnidMed) = PegaUnidMes(Trim(grdPRODENTR.Cell(flexcpText, Row, conCOL_SonProdEnt_CodProd)))
         End If
    Case conCOL_SonProdEnt_Estoque
         If grdPRODENTR.Cell(flexcpText, Row, conCOL_SonProdEnt_Estoque) <> Empty Then
            grdPRODENTR.Cell(flexcpText, Row, conCOL_SonProdEnt_Estoque) = Format(grdPRODENTR.Cell(flexcpText, Row, conCOL_SonProdEnt_Estoque), "#,###0.000")
         End If
    Case conCOL_SonProdEnt_Cadencia
         If grdPRODENTR.Cell(flexcpTextDisplay, Row, conCOL_SonProdEnt_Cadencia) = "Sim" Then Call MudaCadencia(Row, grdPRODUTOS.Cell(flexcpText, grdPRODUTOS.Row, conCOL_SonProd_CodProd))
    End Select
    Exit Sub
End Sub

Private Sub grdPRODENTR_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonProdEnt_Desc_Prod
         Cancel = True
    Case conCOL_SonProdEnt_CodProd, _
         conCOL_SonProdEnt_PesqProd, _
         conCOL_SonProdEnt_UnidMed, _
         conCOL_SonProdEnt_Estoque, _
         conCOL_SonProdEnt_Cadencia
         If cTipOper = "C" Then Cancel = True
    Case Else
        grdPRODENTR.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub grdPRODENTR_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    If cTipOper = "C" Then Exit Sub
    
    If (grdPRODENTR.Rows - 1) = 0 Then Exit Sub
    
    ReDim arrCAMPOS(1 To 3, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    Select Case Col
        Case conCOL_SonProdEnt_PesqProd
    
            sSql = "Select " & vbCrLf
            sSql = sSql & "       PRD.* " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "        SGI_LISTAMATPROD LST " & vbCrLf
            sSql = sSql & "       ,SGI_CADPRODUTO   PRD " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       LST.SGI_FILIAL     = " & FILIAL & vbCrLf
            sSql = sSql & "   And LST.SGI_IDPRODLST  = " & grdPRODUTOS.Cell(flexcpText, grdPRODUTOS.Row, conCOL_SonProd_IDProduto) & vbCrLf
            sSql = sSql & "   And PRD.SGI_FILIAL     = LST.SGI_FILIAL    " & vbCrLf
            sSql = sSql & "   And PRD.SGI_IDPRODUTO  = LST.SGI_IDPRODUTO " & vbCrLf
            
            arrTABELA(1) = sSql
            
            arrCAMPOS(1, 1) = "SGI_IDPRODUTO"
            arrCAMPOS(1, 2) = "S"
            arrCAMPOS(1, 3) = "Id.Produto"
            arrCAMPOS(1, 4) = "1500"
            arrCAMPOS(1, 5) = "PRD.SGI_IDPRODUTO"
            
            arrCAMPOS(2, 1) = "SGI_CODIGO"
            arrCAMPOS(2, 2) = "S"
            arrCAMPOS(2, 3) = "Código"
            arrCAMPOS(2, 4) = "1500"
            arrCAMPOS(2, 5) = "PRD.SGI_CODIGO"
            
            arrCAMPOS(3, 1) = "SGI_DESCRICAO"
            arrCAMPOS(3, 2) = "S"
            arrCAMPOS(3, 3) = "Nome"
            arrCAMPOS(3, 4) = "5000"
            arrCAMPOS(3, 5) = "PRD.SGI_DESCRICAO"
            
            varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Produtos")
            
            If Len(Trim(varRETORNO)) > 0 Then
               
               If VerifItensRepetidos(Row, conCOL_SonProdEnt_CodProd, varRETORNO, conCOL_SonProdEnt_Pai, grdPRODUTOS.Cell(flexcpText, grdPRODUTOS.Row, conCOL_SonProd_IDProduto)) = False Then
                  MsgBox "Este Produto já foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                  grdPRODENTR.Cell(flexcpText, Row, conCOL_SonProdEnt_CodProd) = Empty
                  grdPRODENTR.Cell(flexcpText, Row, conCOL_SonProdEnt_Desc_Prod) = Empty
                  Exit Sub
               End If
               
               grdPRODENTR.Cell(flexcpText, Row, conCOL_SonProdEnt_CodProd) = varRETORNO
               grdPRODENTR.Cell(flexcpText, Row, conCOL_SonProdEnt_Desc_Prod) = PegaDescrProduto(Trim(grdPRODENTR.Cell(flexcpText, Row, conCOL_SonProdEnt_CodProd)))
               grdPRODENTR.Cell(flexcpText, Row, conCOL_SonProdEnt_UnidMed) = PegaUnidMes(Trim(grdPRODENTR.Cell(flexcpText, Row, conCOL_SonProdEnt_CodProd)))
               
            End If
            
    End Select

End Sub

Private Sub grdPRODENTR_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdPRODENTR
          Select Case Col
                    Case conCOL_SonProdEnt_Estoque
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 3, myvarAsCurrency)
          End Select
     End With
End Sub

Private Function VerifItensRepetidos(intROW As Long, intCol As Long, varCampo As Variant, intColPai As Long, varCampoPai As Variant) As Boolean
    VerifItensRepetidos = False
    Dim I As Integer
    
    If Not IsNumeric(varCampo) Then varCampo = UCase(Trim(varCampo))
    For I = 1 To (grdPRODENTR.Rows - 1)
        If I <> intROW And grdPRODENTR.Cell(flexcpText, I, intCol) = varCampo And grdPRODENTR.Cell(flexcpText, I, intColPai) = varCampoPai Then Exit Function
    Next I
    VerifItensRepetidos = True
End Function

Private Function ConsisteProd(strProd As String) As Boolean
    ConsisteProd = False
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PRD.* " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_LISTAMAT   LST " & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO PRD " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       LST.SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And LST.SGI_PRODUTO   = '" & Trim(grdPRODUTOS.Cell(flexcpText, grdPRODUTOS.Row, conCOL_SonProd_CodProd)) & "'"
    sSql = sSql & "   And LST.SGI_PRODLST   = '" & Trim(strProd) & "'" & vbCrLf
    sSql = sSql & "   And PRD.SGI_FILIAL    = LST.SGI_FILIAL  " & vbCrLf
    sSql = sSql & "   And PRD.SGI_CODIGO    = LST.SGI_PRODLST " & vbCrLf
    ''sSql = sSql & "   And PRD.SGI_CADFAMMAQ = " & arrPROCESSO(lngIndice).typProdutos(grdPRODUTOS.Row).lngCODFAMMAQ
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then ConsisteProd = True
    BREC.Close
    
End Function

Private Function PegaDescrProduto(strCODPROD As String) As String
    PegaDescrProduto = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = '" & strCODPROD & "'"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaDescrProduto = BREC!SGI_DESCRICAO
    BREC.Close
    
End Function

Private Function PegaUnidMes(strCODPROD As String) As Long
    PegaUnidMes = 0
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = '" & Trim(strCODPROD) & "'"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaUnidMes = BREC!SGI_UNIDMEDIDA
    BREC.Close
    
End Function

Private Sub IncRegGrid()
   
    If grdPRODUTOS.Row = 0 Then
       MsgBox "Primeiro Selecione um Produto da Grid Produtos que serão Produzidos !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Sub
    End If
    
    If ExisteLinhaVazia = False Then Exit Sub
    
    grdPRODENTR.AddItem "" & vbTab & _
                        "" & vbTab & _
                        "" & vbTab & _
                        "" & vbTab & _
                        "" & vbTab & _
                        "" & vbTab & _
                        Trim(grdPRODUTOS.Cell(flexcpText, grdPRODUTOS.Row, conCOL_SonProd_CodProd))
                            
End Sub

Private Function ExisteLinhaVazia() As Boolean
    ExisteLinhaVazia = False
    
    Dim I As Integer
    
    For I = 1 To (grdPRODENTR.Rows - 1)
        If grdPRODENTR.Cell(flexcpText, I, conCOL_SonProdEnt_CodProd) = Empty Then Exit Function
    Next I
    
    ExisteLinhaVazia = True
End Function


Private Sub PopGrdProdutos()

    Dim I As Integer
    Dim j As Integer
    
    For I = 1 To UBound(arrPROCESSO(lngIndice).typProdutos)
        grdPRODUTOS.AddItem arrPROCESSO(lngIndice).typProdutos(I).strPRODUTO & vbTab & _
                            PegaDescrProduto(Trim(arrPROCESSO(lngIndice).typProdutos(I).strPRODUTO)) & vbTab & _
                            arrPROCESSO(lngIndice).typProdutos(I).lngCodUniMed & vbTab & arrPROCESSO(lngIndice).typProdutos(I).lngIDPRODUTO
                            
        If arrPROCESSO(lngIndice).typProdutos(I).lngTOTPRODENTRADA > 0 Then
           For j = 1 To UBound(arrPROCESSO(lngIndice).typProdutos(I).typProdEntrada)
               grdPRODENTR.AddItem arrPROCESSO(lngIndice).typProdutos(I).typProdEntrada(j).strCODPROD & vbTab & _
                                   "" & vbTab & _
                                   PegaDescrProduto(arrPROCESSO(lngIndice).typProdutos(I).typProdEntrada(j).strCODPROD) & vbTab & _
                                   arrPROCESSO(lngIndice).typProdutos(I).typProdEntrada(j).lngUNIDMED & vbTab & _
                                   Format(arrPROCESSO(lngIndice).typProdutos(I).typProdEntrada(j).curQTDESTOQUE, "#,###0.000") & vbTab & _
                                   arrPROCESSO(lngIndice).typProdutos(I).typProdEntrada(j).intCADENCIA & vbTab & _
                                   arrPROCESSO(lngIndice).typProdutos(I).typProdEntrada(j).strPAI
           Next j
        End If
    Next I
    
    If (grdPRODUTOS.Rows - 1) > 0 Then
        grdPRODUTOS.Row = 1
        Call PosRegProdEntr(grdPRODUTOS.Cell(flexcpText, grdPRODUTOS.Row, conCOL_SonProd_IDProduto))
    End If

End Sub

Private Sub PosRegProdEntr(strCODPROD As String)
    Dim I As Integer
    With grdPRODENTR
         For I = 1 To (.Rows - 1)
             If .Cell(flexcpText, I, conCOL_SonProdEnt_Pai) <> strCODPROD Then
                .RowHidden(I) = True
             Else
                .RowHidden(I) = False
             End If
         Next I
    End With
End Sub

Private Sub grdPRODENTR_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonProdEnt_CodProd
         With grdPRODENTR
            If .EditText <> Empty Then
                If VerifItensRepetidos(Row, conCOL_SonProdEnt_CodProd, .EditText, conCOL_SonProdEnt_Pai, .Cell(flexcpText, Row, conCOL_SonProdEnt_Pai)) = False Then
                   MsgBox "Este Produto já foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                   Cancel = True
                   Exit Sub
                End If
                If ConsisteProd(.EditText) = False Then
                   MsgBox "Este Produto não Existe. !!!", vbOKOnly + vbExclamation, "Aviso"
                   Cancel = True
                   Exit Sub
                End If
            End If
         End With
    End Select
    Exit Sub
End Sub

Private Sub grdPRODUTOS_Click()
    If (grdPRODUTOS.Rows - 1) > 0 And grdPRODUTOS.Row > 0 Then Call PosRegProdEntr(grdPRODUTOS.Cell(flexcpText, grdPRODUTOS.Row, conCOL_SonProd_CodProd))
End Sub

Private Sub grdPRODUTOS_RowColChange()
    If (grdPRODUTOS.Rows - 1) > 0 And grdPRODUTOS.Row > 0 Then Call PosRegProdEntr(grdPRODUTOS.Cell(flexcpText, grdPRODUTOS.Row, conCOL_SonProd_CodProd))
End Sub

Private Sub PopArray()

    Dim I                 As Integer
    Dim j                 As Integer
    Dim intTotRegProdEntr As Integer
    
    Dim arrPRODENTRA()      As ProdEntrada
    Dim arrPRODENTRAV()     As ProdEntrada
    
    If (grdPRODUTOS.Rows - 1) > 0 Then
       For I = 1 To (grdPRODUTOS.Rows - 1)
           '' =======================================================
           '' Produtos de Entrada
           intTotRegProdEntr = 0
           For j = 1 To (grdPRODENTR.Rows - 1)
               If Trim(grdPRODENTR.Cell(flexcpText, j, conCOL_SonProdEnt_Pai)) = Trim(grdPRODUTOS.Cell(flexcpText, I, conCOL_SonProd_CodProd)) Then intTotRegProdEntr = (intTotRegProdEntr + 1)
           Next j
           If intTotRegProdEntr > 0 Then
              ReDim arrPRODENTRA(1 To intTotRegProdEntr) As ProdEntrada
              intTotRegProdEntr = 0
              For j = 1 To (grdPRODENTR.Rows - 1)
                  If Trim(grdPRODENTR.Cell(flexcpText, j, conCOL_SonProdEnt_Pai)) = Trim(grdPRODUTOS.Cell(flexcpText, I, conCOL_SonProd_CodProd)) Then
                     intTotRegProdEntr = intTotRegProdEntr + 1
                     arrPRODENTRA(intTotRegProdEntr).intTipo = 0
                     arrPRODENTRA(intTotRegProdEntr).strCODPROD = grdPRODENTR.Cell(flexcpText, j, conCOL_SonProdEnt_CodProd)
                     If grdPRODENTR.Cell(flexcpText, j, conCOL_SonProdEnt_UnidMed) <> Empty Then
                        arrPRODENTRA(intTotRegProdEntr).lngUNIDMED = grdPRODENTR.Cell(flexcpText, j, conCOL_SonProdEnt_UnidMed)
                     End If
                     If grdPRODENTR.Cell(flexcpText, j, conCOL_SonProdEnt_Estoque) <> Empty Then
                        arrPRODENTRA(intTotRegProdEntr).curQTDESTOQUE = grdPRODENTR.Cell(flexcpText, j, conCOL_SonProdEnt_Estoque)
                     End If
                     arrPRODENTRA(intTotRegProdEntr).intCADENCIA = IIf(grdPRODENTR.Cell(flexcpTextDisplay, j, conCOL_SonProdEnt_Cadencia) = "Sim", 1, 0)
                     arrPRODENTRA(intTotRegProdEntr).strPAI = Trim(grdPRODUTOS.Cell(flexcpText, I, conCOL_SonProd_CodProd))
                  End If
              Next j
              arrPROCESSO(lngIndice).typProdutos(I).lngTOTPRODENTRADA = intTotRegProdEntr
              arrPROCESSO(lngIndice).typProdutos(I).typProdEntrada = arrPRODENTRA
           Else
              arrPROCESSO(lngIndice).curMELHORCENARIO = 0
              arrPROCESSO(lngIndice).curPIORCENARIO = 0
              arrPROCESSO(lngIndice).typProdutos(I).curMELHORCENARIO = 0
              arrPROCESSO(lngIndice).typProdutos(I).curPIORCENARIO = 0
              arrPROCESSO(lngIndice).typProdutos(I).lngTOTPRODENTRADA = 0
              arrPROCESSO(lngIndice).typProdutos(I).typProdEntrada = arrPRODENTRAV
              For j = 1 To arrPROCESSO(lngIndice).typProdutos(I).lngTOTMAQUINAS
                  arrPROCESSO(lngIndice).typProdutos(I).typMaquinas(j).curMELHORCENARIO = 0
                  arrPROCESSO(lngIndice).typProdutos(I).typMaquinas(j).curPIORCENARIO = 0
                  arrPROCESSO(lngIndice).typProdutos(I).typMaquinas(j).curTEMPPROD = 0
              Next j
           End If
       Next I
    End If
          
End Sub

Private Sub MudaCadencia(lngRow As Long, strPRODUTO As String)
    Dim I As Long
    With grdPRODENTR
        For I = 1 To (.Rows - 1)
            If lngRow <> I And Trim(.Cell(flexcpText, I, conCOL_SonProdEnt_Pai)) = Trim(strPRODUTO) Then .Cell(flexcpText, I, conCOL_SonProdEnt_Cadencia) = 0
        Next I
    End With
End Sub


Private Sub Consulta()

    Me.Caption = "Cadastro de Produtos de Entrada - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADPRODENTRADA
    
    Call InitGridProd
    Call InitGridProdEntr
    
    Call PopGrdProdutos
    
End Sub

Private Sub Altera()

    Me.Caption = "Cadastro de Produtos de Entrada - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADPRODENTRADA
    
    Call InitGridProd
    Call InitGridProdEntr
    
    Call PopGrdProdutos
    
End Sub


Private Sub DestroyObjeto()
    Call objBLBFunc.RemoveLinhaVazia(grdPRODENTR, conCOL_SonProdEnt_CodProd)
    Call PopArray
    
    Set objBLBFunc = Nothing
    Set objCADPRODENTRADA = Nothing
    Set objPESQPADRAO = Nothing
End Sub
