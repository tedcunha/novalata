VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmCADPROCESSO 
   Caption         =   "Cadastro do Processo"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14475
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   14475
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraCenario 
      Caption         =   "[ Cenários ]"
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
      Height          =   2775
      Left            =   13080
      TabIndex        =   12
      Top             =   5400
      Width           =   1335
      Begin VSFlex8LCtl.VSFlexGrid grdCENARIOS 
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Visible         =   0   'False
         Width           =   255
         _cx             =   450
         _cy             =   450
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
      Begin VB.OptionButton optCENARIO 
         Caption         =   "Médio"
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
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.OptionButton optCENARIO 
         Caption         =   "Pior"
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
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton optCENARIO 
         Caption         =   "Melhor"
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
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame framaquinas 
      Caption         =   "[ Máquinas ]"
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
      Height          =   2775
      Left            =   0
      TabIndex        =   4
      Top             =   5400
      Width           =   12975
      Begin VSFlex8LCtl.VSFlexGrid grdMaquinas 
         Height          =   2415
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   12735
         _cx             =   22463
         _cy             =   4260
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
   End
   Begin VB.Frame fraProcesso 
      Caption         =   "[ Processo ]"
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
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   14415
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1305
         MaxLength       =   10
         TabIndex        =   9
         Text            =   "txtCodigo"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtCADPROCESSO 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3825
         MaxLength       =   10
         TabIndex        =   8
         Text            =   "txtCADPROC"
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cboProcesso 
         Height          =   315
         Left            =   5400
         TabIndex        =   7
         Text            =   "cboProcesso"
         Top             =   240
         Width           =   8775
      End
      Begin VB.CommandButton Command8 
         Height          =   315
         Left            =   5040
         Picture         =   "frmCADPROCESSO.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Processo:"
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
         Left            =   2640
         TabIndex        =   11
         Top             =   240
         Width           =   855
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
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   285
         Width           =   660
      End
   End
   Begin VB.Frame fraProdutos 
      Caption         =   "[ Operações ]"
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
      Height          =   3855
      Left            =   0
      TabIndex        =   2
      Top             =   1560
      Width           =   14415
      Begin VSFlex8LCtl.VSFlexGrid grdOperacoes 
         Height          =   3495
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   14175
         _cx             =   25003
         _cy             =   6165
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
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14415
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
         Picture         =   "frmCADPROCESSO.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADPROCESSO"
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
Public lngTipo          As Long
Public strPRODUTO       As String
Public lngIndice        As Long

Dim objBLBFunc          As Object
Dim objCADPROCESSO      As Object
Dim objPESQPADRAO       As Object

Const conCOL_SonOpera_Ordem                     As Integer = 0
Const conCOL_SonOpera_CodOper                   As Integer = 1
Const conCOL_SonOpera_Desc_Oper                 As Integer = 2
Const conCOL_SonOpera_CodFamMaq                 As Integer = 3
Const conCOL_SonOpera_DescFamMaq                As Integer = 4
Const conCOL_SonOpera_ProdPai                   As Integer = 5
Const conCOL_SonOpera_IDProd                    As Integer = 6
Const conCOL_SonOpera_FormatString              As String = "=Ordem|Cod.Oper|Desc. Operação|Cod. Fam. Máquina|Descr. Fam. Máquina|Pai|IDProduto"
Const conColumnsIn_SonOpera                     As Integer = 7

Const conCOL_SonMaq_CodMaq                      As Integer = 0
Const conCOL_SonMaq_Desc_maq                    As Integer = 1
Const conCOL_SonMaq_PecasPorMn                  As Integer = 2
Const conCOL_SonMaq_TempoProd                   As Integer = 3
Const conCOL_SonMaq_Indice                      As Integer = 4
Const conCOL_SonMaq_Pai                         As Integer = 5
Const conCOL_SonMaq_FormatString                As String = "=Cod. Maq.|Desc. Máquina|Peças por Hora|Temp.Ptod.|Indice|Pai"
Const conColumnsIn_SonMaq                       As Integer = 6
Private Sub cboProcesso_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboProcesso, KeyAscii
End Sub

Private Sub cboProcesso_Validate(Cancel As Boolean)
    Dim arrPRODUTOSV()   As PRODUTOS
    If cboProcesso.ListIndex > -1 Then
       txtCADPROCESSO.Text = cboProcesso.ItemData(cboProcesso.ListIndex)
       If arrPROCESSO(lngIndice).lngCODIGO <> CLng(txtCADPROCESSO.Text) Then
          Call InitGridMaquinas
          Call InitGridOperacoes
          arrPROCESSO(lngIndice).lngCODIGO = Empty
          arrPROCESSO(lngIndice).strDESCRI = Empty
          arrPROCESSO(lngIndice).lngQTDPRODUTOS = 0
          arrPROCESSO(lngIndice).typProdutos = arrPRODUTOSV
       End If
    End If
End Sub


Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Command8_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_CODIGO " & vbCrLf
    sSql = sSql & "      ,SGI_DESCRI " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPROCESSO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1500"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRI"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "SGI_DESCRI"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Processos Produtivo")
    
    If Len(Trim(varRETORNO)) > 0 Then
        txtCADPROCESSO.Text = varRETORNO
        Call PopGrdOperacoes(txtCADPROCESSO.Text)
    End If
    
    cboProcesso.ListIndex = -1
    txtCADPROCESSO.SetFocus

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADPROCESSO = CreateObject("CADFLUXOPROD.clsCADFLUXOPROD")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
   objCADPROCESSO.FILIAL = FILIAL
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

End Sub

Private Sub Inclui()
    
    Dim I As Integer
    
    fraProcesso.Enabled = True
       
    Me.Caption = "Cadastro de Processos - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADPROCESSO
    
    objCADPROCESSO.PreenchComboProcesso cboProcesso
    
    Call InitGridOperacoes
    Call InitGridMaquinas
    
    If arrPROCESSO(lngIndice).lngCODIGO <> Empty Then
       
       txtCADPROCESSO.Text = arrPROCESSO(lngIndice).lngCODIGO
       cboProcesso.ListIndex = -1
       For I = 0 To (cboProcesso.ListCount - 1)
           If cboProcesso.ItemData(I) = arrPROCESSO(lngIndice).lngCODIGO Then cboProcesso.ListIndex = I
       Next I
       
       Call PopGrdProdMaq
       
    End If
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DestroyObjeto
End Sub

Private Sub grdMaquinas_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    Select Case Col
    Case conCOL_SonMaq_PecasPorMn
         If grdMaquinas.Cell(flexcpText, Row, conCOL_SonMaq_PecasPorMn) <> Empty Then
            grdMaquinas.Cell(flexcpText, Row, conCOL_SonMaq_PecasPorMn) = Format(grdMaquinas.Cell(flexcpText, Row, conCOL_SonMaq_PecasPorMn), "#,###0.000")
         End If
    End Select
    Exit Sub

End Sub

Private Sub grdMaquinas_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonMaq_CodMaq, _
         conCOL_SonMaq_Desc_maq, _
         conCOL_SonMaq_TempoProd
         Cancel = True
    Case conCOL_SonMaq_PecasPorMn
         If cTipOper = "C" Then Cancel = True
    Case Else
        grdMaquinas.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub grdMaquinas_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdMaquinas
          Select Case Col
                    Case conCOL_SonMaq_PecasPorMn, conCOL_SonMaq_TempoProd
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 3, myvarAsCurrency)
          End Select
     End With
End Sub

Private Sub grdOperacoes_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonOpera_Ordem, _
         conCOL_SonOpera_CodOper, _
         conCOL_SonOpera_Desc_Oper, _
         conCOL_SonOpera_CodFamMaq, _
         conCOL_SonOpera_Desc_Oper
         Cancel = True
    Case Else
        grdOperacoes.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub grdOperacoes_Click()
    If (grdOperacoes.Rows - 1) > 0 And (grdOperacoes.Row) > 0 Then
        Call PosRegGrdMaq(grdOperacoes.Cell(flexcpText, grdOperacoes.Row, conCOL_SonOpera_IDProd))
    End If
End Sub

Private Sub grdOperacoes_RowColChange()
    If (grdOperacoes.Rows - 1) > 0 And (grdOperacoes.Row) > 0 Then
        Call PosRegGrdMaq(grdOperacoes.Cell(flexcpText, grdOperacoes.Row, conCOL_SonOpera_IDProd))
    End If
End Sub

Private Sub optCENARIO_Click(Index As Integer)
    Call PintaMAquinaCenarios(Index)
End Sub

Private Sub txtCADPROCESSO_GotFocus()
    objBLBFunc.SelecionaCampos txtCADPROCESSO.Name, frmCADPROCESSO
End Sub

Private Sub txtCADPROCESSO_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCADPROCESSO.Text
End Sub

Private Sub txtCADPROCESSO_Validate(Cancel As Boolean)

    Dim I                   As Integer
    Dim arrPRODUTOSV()      As PRODUTOS
    
    If Len(Trim(txtCADPROCESSO.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCADPROCESSO.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCADPROCESSO.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    If arrPROCESSO(lngIndice).lngCODIGO <> CLng(txtCADPROCESSO.Text) Then
       Call InitGridMaquinas
       Call InitGridOperacoes
       arrPROCESSO(lngIndice).lngCODIGO = Empty
       arrPROCESSO(lngIndice).strDESCRI = Empty
       arrPROCESSO(lngIndice).lngQTDPRODUTOS = 0
       arrPROCESSO(lngIndice).typProdutos = arrPRODUTOSV
    End If
    
    cboProcesso.ListIndex = -1
    For I = 0 To (cboProcesso.ListCount - 1)
        If cboProcesso.ItemData(I) = CInt(txtCADPROCESSO.Text) Then cboProcesso.ListIndex = I
    Next I
    
    If cboProcesso.ListIndex = -1 Then
       MsgBox "Este processo não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCADPROCESSO.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Call PopGrdOperacoes(txtCADPROCESSO.Text)

End Sub

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

Private Function PegaCodFamMaq(strCODPROD As String) As Long
    
    PegaCodFamMaq = 0
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = '" & Trim(strCODPROD) & "'"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
       If Not IsNull(BREC!SGI_CADFAMMAQ) Then PegaCodFamMaq = BREC!SGI_CADFAMMAQ
    End If
    BREC.Close
    
End Function


Private Function PegaDescFamMaq(lngCODFAMMAQ As Long) As String
    PegaDescFamMaq = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADFAMMAQUINAS " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCODFAMMAQ
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaDescFamMaq = BREC!SGI_DESCRI
    BREC.Close
    
End Function


Private Sub InitGridMaquinas()

    With grdMaquinas
    
       .Cols = conColumnsIn_SonMaq
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonMaq_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonMaq_CodMaq) = ""
       .ColDataType(conCOL_SonMaq_CodMaq) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonMaq_Desc_maq) = ""
       .ColDataType(conCOL_SonMaq_Desc_maq) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMaq_PecasPorMn) = ""
       .ColDataType(conCOL_SonMaq_PecasPorMn) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonMaq_TempoProd) = ""
       .ColDataType(conCOL_SonMaq_TempoProd) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonMaq_Indice) = ""
       .ColDataType(conCOL_SonMaq_Indice) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMaq_Pai) = ""
       .ColDataType(conCOL_SonMaq_Pai) = flexDTString
       
       .ColWidth(conCOL_SonMaq_CodMaq) = 1500
       .ColWidth(conCOL_SonMaq_Desc_maq) = 4000
       .ColWidth(conCOL_SonMaq_PecasPorMn) = 1500
       .ColWidth(conCOL_SonMaq_TempoProd) = 1500
       .ColWidth(conCOL_SonMaq_Indice) = 0
       .ColWidth(conCOL_SonMaq_Pai) = 0
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub


Private Sub PopGrdMaquina(lngFAMMAQ As Long, strCODPROD As String)

    Dim I       As Integer
    Dim intREG  As Integer
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADMAQUINA " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODFAMILIA = " & lngFAMMAQ & vbCrLf
    sSql = sSql & " Order by SGI_CODIGO "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC.EOF
        intREG = grdMaquinas.FindRow(Trim(Str(BREC!SGI_CODIGO)) & Trim(strCODPROD), , conCOL_SonMaq_Indice)
        If intREG = -1 Then
           grdMaquinas.AddItem BREC!SGI_CODIGO & vbTab & _
                               BREC!SGI_DESCRI & vbTab & _
                               "" & vbTab & _
                               "" & vbTab & _
                               Trim(Str(BREC!SGI_CODIGO)) & Trim(strCODPROD) & vbTab & _
                               Trim(strCODPROD)
        End If
        BREC.MoveNext
    Loop
    BREC.Close
    
End Sub

Private Sub PosRegGrdMaq(strCODPROD As String)
    Dim I As Integer
    With grdMaquinas
         For I = 1 To (.Rows - 1)
             If .Cell(flexcpText, I, conCOL_SonMaq_Pai) <> strCODPROD Then
                .RowHidden(I) = True
             Else
                .RowHidden(I) = False
             End If
         Next I
    End With
End Sub

Private Function ConsisteProd(strProd As String) As Boolean
    ConsisteProd = False
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PRD.* " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPROCESSO PR  " & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO  PRD " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       PR.SGI_FILIAL      = " & FILIAL & vbCrLf
    sSql = sSql & "   And PR.SGI_CODIGO      = " & Trim(txtCADPROCESSO.Text) & vbCrLf
    sSql = sSql & "   And PRD.SGI_FILIAL     = PR.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PRD.SGI_CADFAMMAQ  = PR.SGI_CODFAMILIA " & vbCrLf
    sSql = sSql & "   And PRD.SGI_CODIGO     = '" & Trim(strProd) & "'"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then ConsisteProd = True
    BREC.Close
    
End Function

Private Function PegaDescMaq(lngCODMAQ As Long) As String
    PegaDescMaq = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADMAQUINA " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCODMAQ
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaDescMaq = BREC!SGI_DESCRI
    BREC.Close
    
End Function

Private Sub PopGrdProdMaq()

       Dim I As Integer
       Dim j As Integer
       Dim curESTINTERM As Currency
       Dim curTEMPPROD  As Currency
       
       '' =====================================================
       '' Popula a Grid de Produtos
       If arrPROCESSO(lngIndice).lngQTDPRODUTOS > 0 Then
          For I = 1 To UBound(arrPROCESSO(lngIndice).typProdutos)
              grdOperacoes.AddItem I & vbTab & _
                                   arrPROCESSO(lngIndice).typProdutos(I).strPRODUTO & vbTab & _
                                   PegaDescProc(arrPROCESSO(lngIndice).typProdutos(I).strPRODUTO) & vbTab & _
                                   arrPROCESSO(lngIndice).typProdutos(I).lngCODFAMMAQ & vbTab & _
                                   PegaDescFamMaq(arrPROCESSO(lngIndice).typProdutos(I).lngCODFAMMAQ) & vbTab & _
                                   strPRODUTO & vbTab & _
                                   arrPROCESSO(lngIndice).typProdutos(I).lngIDPRODUTO
                                  
              '' Pegando os Estoques Intemediário
              curESTINTERM = 0
              If arrPROCESSO(lngIndice).typProdutos(I).lngTOTPRODENTRADA > 0 Then
                 For j = 1 To UBound(arrPROCESSO(lngIndice).typProdutos(I).typProdEntrada)
                     If arrPROCESSO(lngIndice).typProdutos(I).typProdEntrada(j).intCADENCIA = 1 Then
                        '' Valor do Estoque Intemediário
                        curESTINTERM = arrPROCESSO(lngIndice).typProdutos(I).typProdEntrada(j).curQTDESTOQUE
                     End If
                 Next j
              End If
              If arrPROCESSO(lngIndice).typProdutos(I).lngTOTMAQUINAS > 0 Then
                 For j = 1 To UBound(arrPROCESSO(lngIndice).typProdutos(I).typMaquinas)
                     If curESTINTERM > 0 Then
                        If arrPROCESSO(lngIndice).typProdutos(I).typMaquinas(j).lngQTDPCMIN > 0 Then
                           curTEMPPROD = (curESTINTERM / arrPROCESSO(lngIndice).typProdutos(I).typMaquinas(j).lngQTDPCMIN)
                           arrPROCESSO(lngIndice).typProdutos(I).typMaquinas(j).curTEMPPROD = curTEMPPROD
                        End If
                     End If
                     grdMaquinas.AddItem arrPROCESSO(lngIndice).typProdutos(I).typMaquinas(j).lngCODMAQ & vbTab & _
                                         PegaDescMaq(arrPROCESSO(lngIndice).typProdutos(I).typMaquinas(j).lngCODMAQ) & vbTab & _
                                         Format(arrPROCESSO(lngIndice).typProdutos(I).typMaquinas(j).lngQTDPCMIN, "#,###0.000") & vbTab & _
                                         Format(arrPROCESSO(lngIndice).typProdutos(I).typMaquinas(j).curTEMPPROD, "#,##0.00") & vbTab & _
                                         arrPROCESSO(lngIndice).typProdutos(I).typMaquinas(j).strINDICE & vbTab & _
                                         arrPROCESSO(lngIndice).typProdutos(I).typMaquinas(j).strPAI
                 Next j
              End If
              
          Next I
          If (grdOperacoes.Rows - 1) > 0 Then
              grdOperacoes.Row = 1
              Call PosRegGrdMaq(grdOperacoes.Cell(flexcpText, grdOperacoes.Row, conCOL_SonOpera_IDProd))
          End If
       End If
       '' =====================================================

End Sub

Private Sub PopArray()

    Dim I               As Integer
    Dim j               As Integer
    Dim intTotRegMaq    As Integer
    
    Dim arrPRODUTOS()       As PRODUTOS
    Dim arrMAQUINAS()       As Maquinas
    Dim arrMAQUINASV()      As Maquinas
    Dim arrPRODENTRADA()    As ProdEntrada
    Dim arrPRODSAIDA()      As ProdSaida
    
    If Len(Trim(txtCADPROCESSO.Text)) > 0 Then
       arrPROCESSO(lngIndice).lngCODIGO = CLng(txtCADPROCESSO.Text)
       arrPROCESSO(lngIndice).strDESCRI = cboProcesso.List(cboProcesso.ListIndex)
       
       If (grdOperacoes.Rows - 1) > 0 Then
       
          '' Pegando Produtos
          ReDim arrPRODUTOS(1 To (grdOperacoes.Rows - 1)) As PRODUTOS
          For I = 1 To (grdOperacoes.Rows - 1)
              
              '' ========================================================
              '' Produtos de Entrada
              If I <= arrPROCESSO(lngIndice).lngQTDPRODUTOS Then
              If arrPROCESSO(lngIndice).typProdutos(I).lngTOTPRODENTRADA > 0 Then
                 ReDim arrPRODENTRADA(1 To arrPROCESSO(lngIndice).typProdutos(I).lngTOTPRODENTRADA) As ProdEntrada
                 For j = 1 To UBound(arrPRODENTRADA)
                     arrPRODENTRADA(j).strCODPROD = arrPROCESSO(lngIndice).typProdutos(I).typProdEntrada(j).strCODPROD
                     arrPRODENTRADA(j).lngUNIDMED = arrPROCESSO(lngIndice).typProdutos(I).typProdEntrada(j).lngUNIDMED
                     arrPRODENTRADA(j).curQTDESTOQUE = arrPROCESSO(lngIndice).typProdutos(I).typProdEntrada(j).curQTDESTOQUE
                     arrPRODENTRADA(j).intCADENCIA = arrPROCESSO(lngIndice).typProdutos(I).typProdEntrada(j).intCADENCIA
                     arrPRODENTRADA(j).intTipo = arrPROCESSO(lngIndice).typProdutos(I).typProdEntrada(j).intTipo
                     arrPRODENTRADA(j).strPAI = arrPROCESSO(lngIndice).typProdutos(I).typProdEntrada(j).strPAI
                 Next j
              End If
              '' ========================================================
              
              '' ========================================================
              '' Produtos de Saida
              If arrPROCESSO(lngIndice).typProdutos(I).lngTOTPRODSAIDA > 0 Then
                 ReDim arrPRODSAIDA(1 To arrPROCESSO(lngIndice).typProdutos(I).lngTOTPRODSAIDA) As ProdSaida
                 For j = 1 To UBound(arrPRODSAIDA)
                     arrPRODSAIDA(j).strCODPROD = arrPROCESSO(lngIndice).typProdutos(I).typProdSaida(j).strCODPROD
                     arrPRODSAIDA(j).lngUNIDMED = arrPROCESSO(lngIndice).typProdutos(I).typProdSaida(j).lngUNIDMED
                     arrPRODSAIDA(j).curQTDESTOQUE = arrPROCESSO(lngIndice).typProdutos(I).typProdSaida(j).curQTDESTOQUE
                     arrPRODSAIDA(j).intTipo = arrPROCESSO(lngIndice).typProdutos(I).typProdSaida(j).intTipo
                     arrPRODSAIDA(j).strPAI = arrPROCESSO(lngIndice).typProdutos(I).typProdSaida(j).strPAI
                 Next j
              End If
              End If
              '' ========================================================
              
              
              arrPRODUTOS(I).intTipo = 0
              arrPRODUTOS(I).lngIDPRODUTO = grdOperacoes.Cell(flexcpText, I, conCOL_SonOpera_IDProd)
              arrPRODUTOS(I).strPRODUTO = grdOperacoes.Cell(flexcpText, I, conCOL_SonOpera_CodOper)
              arrPRODUTOS(I).lngCodUniMed = 0
              If grdOperacoes.Cell(flexcpText, I, conCOL_SonOpera_CodFamMaq) <> Empty Then
                 arrPRODUTOS(I).lngCODFAMMAQ = grdOperacoes.Cell(flexcpText, I, conCOL_SonOpera_CodFamMaq)
              End If
              arrPRODUTOS(I).lngTOTMAQUINAS = 0
              arrPRODUTOS(I).typMaquinas = arrMAQUINASV
              
              '' =======================================================
              '' Maquinas
              intTotRegMaq = 0
              For j = 1 To (grdMaquinas.Rows - 1)
                  If Trim(grdMaquinas.Cell(flexcpText, j, conCOL_SonMaq_Pai)) = Trim(grdOperacoes.Cell(flexcpText, I, conCOL_SonOpera_IDProd)) Then intTotRegMaq = (intTotRegMaq + 1)
              Next j
              If intTotRegMaq > 0 Then
                 ReDim arrMAQUINAS(1 To intTotRegMaq) As Maquinas
                 intTotRegMaq = 0
                 For j = 1 To (grdMaquinas.Rows - 1)
                     If Trim(grdMaquinas.Cell(flexcpText, j, conCOL_SonMaq_Pai)) = Trim(grdOperacoes.Cell(flexcpText, I, conCOL_SonOpera_IDProd)) Then
                        intTotRegMaq = (intTotRegMaq + 1)
                        arrMAQUINAS(intTotRegMaq).intTipo = 0
                        arrMAQUINAS(intTotRegMaq).lngCODMAQ = grdMaquinas.Cell(flexcpText, j, conCOL_SonMaq_CodMaq)
                        If grdMaquinas.Cell(flexcpText, j, conCOL_SonMaq_PecasPorMn) <> Empty Then
                           arrMAQUINAS(intTotRegMaq).lngQTDPCMIN = grdMaquinas.Cell(flexcpText, j, conCOL_SonMaq_PecasPorMn)
                        End If
                        If grdMaquinas.Cell(flexcpText, j, conCOL_SonMaq_TempoProd) <> Empty Then
                           arrMAQUINAS(intTotRegMaq).curTEMPPROD = grdMaquinas.Cell(flexcpText, j, conCOL_SonMaq_TempoProd)
                        End If
                        arrMAQUINAS(intTotRegMaq).strINDICE = grdMaquinas.Cell(flexcpText, j, conCOL_SonMaq_Indice)
                        arrMAQUINAS(intTotRegMaq).strPAI = grdMaquinas.Cell(flexcpText, j, conCOL_SonMaq_Pai)
                     End If
                 Next j
                 arrPRODUTOS(I).lngTOTMAQUINAS = intTotRegMaq
                 
                 If I <= arrPROCESSO(lngIndice).lngQTDPRODUTOS Then
                    arrPRODUTOS(I).lngTOTPRODENTRADA = arrPROCESSO(lngIndice).typProdutos(I).lngTOTPRODENTRADA
                    arrPRODUTOS(I).lngTOTPRODSAIDA = arrPROCESSO(lngIndice).typProdutos(I).lngTOTPRODSAIDA
                 End If
                 
                 arrPRODUTOS(I).typMaquinas = arrMAQUINAS
                 arrPRODUTOS(I).typProdEntrada = arrPRODENTRADA
                 arrPRODUTOS(I).typProdSaida = arrPRODSAIDA
              End If
              '' =======================================================
              
          Next I
          arrPROCESSO(lngIndice).lngQTDPRODUTOS = (grdOperacoes.Rows - 1)
          arrPROCESSO(lngIndice).typProdutos = arrPRODUTOS
       Else
          arrPROCESSO(lngIndice).lngQTDPRODUTOS = 0
          arrPROCESSO(lngIndice).typProdutos = arrPRODUTOS
       End If
    Else
       arrPROCESSO(lngIndice).lngCODIGO = Empty
       arrPROCESSO(lngIndice).strDESCRI = Empty
       arrPROCESSO(lngIndice).lngQTDPRODUTOS = 0
       arrPROCESSO(lngIndice).typProdutos = arrPRODUTOS
    End If

End Sub

Private Function MelhorCenario(strPRODUTO As String) As String
        
        MelhorCenario = ""
        Dim I               As Integer
        Dim j               As Integer
        Dim intROW          As Integer
        With grdMaquinas
             grdCENARIOS.Cols = 3
             grdCENARIOS.FixedRows = 1
             grdCENARIOS.FixedCols = 0
             grdCENARIOS.Rows = 1
             grdCENARIOS.ColWidth(0) = 1000
             grdCENARIOS.ColWidth(1) = 1000
             grdCENARIOS.ColWidth(2) = 1000
             grdCENARIOS.ColHidden(0) = True
             Const conCOL_Cenarios_FormatString  As String = "=|Máquina|Valor"
             grdCENARIOS.FormatString = conCOL_Cenarios_FormatString
             For I = 1 To (.Rows - 1)
                 If .Cell(flexcpText, I, conCOL_SonMaq_Pai) = Trim(strPRODUTO) Then
                    If .Cell(flexcpText, I, conCOL_SonMaq_TempoProd) <> Empty Then
                        grdCENARIOS.AddItem strPRODUTO & vbTab & .Cell(flexcpText, I, conCOL_SonMaq_CodMaq) & vbTab & _
                                            .Cell(flexcpText, I, conCOL_SonMaq_TempoProd)
                    End If
                 End If
             Next I
             If grdCENARIOS.Rows - 1 > 0 Then
                grdCENARIOS.Subtotal flexSTMin, 0, 2, , , , , "Melhor"
                For I = 1 To (grdCENARIOS.Rows - 1)
                    If grdCENARIOS.Cell(flexcpText, I, 0) = "Melhor" Then
                       For j = 1 To (grdCENARIOS.Rows - 1)
                           If j <> I And grdCENARIOS.Cell(flexcpText, j, 2) = grdCENARIOS.Cell(flexcpText, I, 2) Then
                              MelhorCenario = grdCENARIOS.Cell(flexcpText, j, 1) & grdCENARIOS.Cell(flexcpText, j, 0)
                              Exit For
                           End If
                       Next j
                    End If
                Next I
             End If
        End With
        
End Function

Private Function PiorCenario(strPRODUTO As String) As String
        PiorCenario = ""
        Dim I               As Integer
        Dim j               As Integer
        With grdMaquinas
             grdCENARIOS.Cols = 3
             grdCENARIOS.FixedRows = 1
             grdCENARIOS.FixedCols = 0
             grdCENARIOS.Rows = 1
             grdCENARIOS.ColWidth(0) = 1000
             grdCENARIOS.ColWidth(1) = 500
             grdCENARIOS.ColWidth(2) = 1000
             Const conCOL_Cenarios_FormatString  As String = "=|Máquina|Valor"
             grdCENARIOS.FormatString = conCOL_Cenarios_FormatString
             For I = 1 To (.Rows - 1)
                 If .Cell(flexcpText, I, conCOL_SonMaq_Pai) = Trim(strPRODUTO) Then
                    If .Cell(flexcpText, I, conCOL_SonMaq_TempoProd) <> Empty Then
                        grdCENARIOS.AddItem strPRODUTO & vbTab & .Cell(flexcpText, I, conCOL_SonMaq_CodMaq) & vbTab & _
                                            .Cell(flexcpText, I, conCOL_SonMaq_TempoProd)
                    End If
                 End If
             Next I
             If grdCENARIOS.Rows - 1 > 0 Then
                grdCENARIOS.Subtotal flexSTMax, 0, 2, , , , , "Pior"
                For I = 1 To (grdCENARIOS.Rows - 1)
                    If grdCENARIOS.Cell(flexcpText, I, 0) = "Pior" Then
                       For j = 1 To (grdCENARIOS.Rows - 1)
                           If j <> I And grdCENARIOS.Cell(flexcpText, j, 2) = grdCENARIOS.Cell(flexcpText, I, 2) Then
                              PiorCenario = grdCENARIOS.Cell(flexcpText, j, 1) & grdCENARIOS.Cell(flexcpText, j, 0)
                              Exit For
                           End If
                       Next j
                    End If
                Next I
             End If
        End With
End Function

Private Function MediaCenario(strPRODUTO As String) As String
        MediaCenario = ""
        Dim I               As Integer
        Dim j               As Integer
        With grdMaquinas
             grdCENARIOS.Cols = 3
             grdCENARIOS.FixedRows = 1
             grdCENARIOS.FixedCols = 0
             grdCENARIOS.Rows = 1
             grdCENARIOS.ColWidth(0) = 1000
             grdCENARIOS.ColWidth(1) = 500
             grdCENARIOS.ColWidth(2) = 1000
             Const conCOL_Cenarios_FormatString  As String = "=|Máquina|Valor"
             grdCENARIOS.FormatString = conCOL_Cenarios_FormatString
             For I = 1 To (.Rows - 1)
                 If .Cell(flexcpText, I, conCOL_SonMaq_Pai) = Trim(strPRODUTO) Then
                    If .Cell(flexcpText, I, conCOL_SonMaq_TempoProd) <> Empty Then
                        grdCENARIOS.AddItem strPRODUTO & vbTab & .Cell(flexcpText, I, conCOL_SonMaq_CodMaq) & vbTab & _
                                            .Cell(flexcpText, I, conCOL_SonMaq_TempoProd)
                    End If
                 End If
             Next I
             If grdCENARIOS.Rows - 1 > 0 Then
                grdCENARIOS.Subtotal flexSTAverage, 0, 2, , , , , "Médio"
                For I = 1 To (grdCENARIOS.Rows - 1)
                    If grdCENARIOS.Cell(flexcpText, I, 0) = "Médio" Then
                       For j = 1 To (grdCENARIOS.Rows - 1)
                           If j <> I And grdCENARIOS.Cell(flexcpText, j, 2) = grdCENARIOS.Cell(flexcpText, I, 2) Then
                              MediaCenario = grdCENARIOS.Cell(flexcpText, j, 1) & grdCENARIOS.Cell(flexcpText, j, 0)
                              Exit For
                           End If
                       Next j
                    End If
                Next I
             End If
        End With
End Function

Private Sub PintaMAquinaCenarios(Index As Integer)

    Dim intROW As Integer
    Dim I      As Integer
    With grdMaquinas
        For I = 1 To (grdMaquinas.Rows - 1)
            .Cell(flexcpForeColor, I, conCOL_SonMaq_CodMaq, I, .Cols - 1) = vbBlack
        Next I
        If Index = 0 Then
           intROW = .FindRow(MelhorCenario(grdOperacoes.Cell(flexcpText, grdOperacoes.Row, conCOL_SonOpera_IDProd)), , conCOL_SonMaq_Indice)
           If intROW <> -1 Then .Cell(flexcpForeColor, intROW, conCOL_SonMaq_CodMaq, intROW, .Cols - 1) = &HC000&      '' Verde
        ElseIf Index = 1 Then
           intROW = .FindRow(PiorCenario(grdOperacoes.Cell(flexcpText, grdOperacoes.Row, conCOL_SonOpera_IDProd)), , conCOL_SonMaq_Indice)
           If intROW <> -1 Then .Cell(flexcpForeColor, intROW, conCOL_SonMaq_CodMaq, intROW, .Cols - 1) = &HFF&        '' Vermelho
        ElseIf Index = 2 Then
           intROW = .FindRow(MediaCenario(grdOperacoes.Cell(flexcpText, grdOperacoes.Row, conCOL_SonOpera_IDProd)), , conCOL_SonMaq_Indice)
           If intROW <> -1 Then .Cell(flexcpForeColor, intROW, conCOL_SonMaq_CodMaq, intROW, .Cols - 1) = &HC0C0&      '' Amarelo
        End If
    End With

End Sub

Private Sub ExcluiFilho()
       Dim I As Integer
VOLTA:
       For I = 1 To (grdMaquinas.Rows - 1)
           If grdMaquinas.Cell(flexcpText, I, conCOL_SonMaq_Pai) = grdOperacoes.Cell(flexcpText, grdOperacoes.Row, conCOL_SonOpera_IDProd) Then
              If (grdMaquinas.Rows - 1) = 2 Then grdMaquinas.Rows = 1
              If (grdMaquinas.Rows - 1) > 2 Then grdMaquinas.RemoveItem I
              GoTo VOLTA
           End If
       Next I
End Sub


Private Sub Consulta()
    
    Dim I As Integer
    
    fraProcesso.Enabled = False
       
    Me.Caption = "Cadastro de Processos - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADPROCESSO
    
    objCADPROCESSO.PreenchComboProcesso cboProcesso
    
    Call InitGridOperacoes
    Call InitGridMaquinas
    
    If arrPROCESSO(lngIndice).lngCODIGO <> Empty Then
       
       txtCodigo.Text = iCodigo
       txtCADPROCESSO.Text = arrPROCESSO(lngIndice).lngCODIGO
       cboProcesso.ListIndex = -1
       For I = 0 To (cboProcesso.ListCount - 1)
           If cboProcesso.ItemData(I) = arrPROCESSO(lngIndice).lngCODIGO Then cboProcesso.ListIndex = I
       Next I
       
       Call PopGrdProdMaq
       
    End If
    
    
End Sub

Private Sub Altera()
    
    Dim I As Integer
    
    fraProcesso.Enabled = True
       
    Me.Caption = "Cadastro de Processos - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADPROCESSO
    
    objCADPROCESSO.PreenchComboProcesso cboProcesso
    
    Call InitGridOperacoes
    Call InitGridMaquinas
    
    If arrPROCESSO(lngIndice).lngCODIGO <> Empty Then
       
       txtCodigo.Text = iCodigo
       txtCADPROCESSO.Text = arrPROCESSO(lngIndice).lngCODIGO
       cboProcesso.ListIndex = -1
       For I = 0 To (cboProcesso.ListCount - 1)
           If cboProcesso.ItemData(I) = arrPROCESSO(lngIndice).lngCODIGO Then cboProcesso.ListIndex = I
       Next I
       
       Call PopGrdProdMaq
       
    End If
    
    
End Sub


Private Sub DestroyObjeto()
    Call objBLBFunc.RemoveLinhaVazia(grdOperacoes, conCOL_SonOpera_CodOper)
    Call objBLBFunc.RemoveLinhaVazia(grdOperacoes, conCOL_SonOpera_IDProd)
    Call PopArray
    
    Set objBLBFunc = Nothing
    Set objCADPROCESSO = Nothing
    Set objPESQPADRAO = Nothing
End Sub

Private Sub InitGridOperacoes()

    With grdOperacoes
    
       .Cols = conColumnsIn_SonOpera
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonOpera_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonOpera_Ordem) = ""
       .ColDataType(conCOL_SonOpera_Ordem) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonOpera_CodOper) = ""
       .ColDataType(conCOL_SonOpera_CodOper) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonOpera_Desc_Oper) = ""
       .ColDataType(conCOL_SonOpera_Desc_Oper) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonOpera_CodFamMaq) = ""
       .ColDataType(conCOL_SonOpera_CodFamMaq) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonOpera_DescFamMaq) = ""
       .ColDataType(conCOL_SonOpera_DescFamMaq) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonOpera_ProdPai) = ""
       .ColDataType(conCOL_SonOpera_ProdPai) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonOpera_IDProd) = ""
       .ColDataType(conCOL_SonOpera_IDProd) = flexDTLong
       
       .ColWidth(conCOL_SonOpera_Ordem) = 1000
       .ColWidth(conCOL_SonOpera_CodOper) = 1500
       .ColWidth(conCOL_SonOpera_Desc_Oper) = 4000
       .ColWidth(conCOL_SonOpera_CodFamMaq) = 1500
       .ColWidth(conCOL_SonOpera_DescFamMaq) = 4000
       .ColWidth(conCOL_SonOpera_ProdPai) = 0
       .ColWidth(conCOL_SonOpera_IDProd) = 0
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub


Private Sub PopGrdOperacoes(strCodOper As String)

    Call InitGridOperacoes
    Call InitGridMaquinas
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PRO.SGI_ORDEM" & vbCrLf
    sSql = sSql & "      ,PRO.SGI_CODOPER" & vbCrLf
    sSql = sSql & "      ,OPER.SGI_DESCRI" & vbCrLf
    sSql = sSql & "      ,OPER.SGI_CODFAMMAQ" & vbCrLf
    sSql = sSql & "      ,FMQ.SGI_DESCRI As SGI_DESCFAMMAQ " & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPROCOPERA PRO" & vbCrLf
    sSql = sSql & "      ,SGI_TIPOPERACAO  OPER" & vbCrLf
    sSql = sSql & "      ,SGI_CADFAMMAQUINAS FMQ" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       PRO.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And PRO.SGI_CODIGO = " & strCodOper & vbCrLf
    sSql = sSql & "   And OPER.SGI_FILIAL = PRO.SGI_FILIAL  " & vbCrLf
    sSql = sSql & "   And OPER.SGI_CODIGO = PRO.SGI_CODOPER " & vbCrLf
    sSql = sSql & "   And FMQ.SGI_FILIAL  = OPER.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And FMQ.SGI_CODIGO  = OPER.SGI_CODFAMMAQ "
    sSql = sSql & "Order By PRO.SGI_ORDEM"

    BREC12.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC12.EOF() Then
        With grdOperacoes
            Do While Not BREC12.EOF()
                
                .AddItem BREC12!SGI_ORDEM & vbTab & _
                         BREC12!SGI_CODOPER & vbTab & _
                         BREC12!SGI_DESCRI & vbTab & _
                         BREC12!SGI_CODFAMMAQ & vbTab & _
                         BREC12!SGI_DESCFAMMAQ & vbTab & _
                         Trim(strPRODUTO) & vbTab & _
                         Trim(Str(BREC12!SGI_ORDEM)) & Trim(Str(BREC12!SGI_CODOPER))
                         
                Call PopGrdMaqOperacao(Str(BREC12!SGI_CODFAMMAQ), Str(BREC12!SGI_CODOPER), Str(BREC12!SGI_ORDEM))
                
                BREC12.MoveNext
            Loop
        End With
    End If
    BREC12.Close
        
    If (grdOperacoes.Rows - 1) > 0 Then
        grdOperacoes.Row = 1
        Call PosRegGrdMaq(grdOperacoes.Cell(flexcpText, grdOperacoes.Row, conCOL_SonOpera_IDProd))
    End If


End Sub

Private Function PegaDescProc(strCodOper As String) As String

    PegaDescProc = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_DESCRI " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_TIPOPERACAO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & strCodOper

    BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC3.EOF() Then PegaDescProc = BREC3!SGI_DESCRI
    BREC3.Close
    
End Function

Private Sub PopGrdMaqOperacao(strCodFamMaq As String, strCodOper As String, strOrdem As String)

    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADMAQUINA " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODFAMILIA = " & strCodFamMaq & vbCrLf
    sSql = sSql & " Order by SGI_CODIGO "
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then
        With grdMaquinas
            Do While Not BREC10.EOF()
                .AddItem BREC10!SGI_CODIGO & vbTab & _
                         BREC10!SGI_DESCRI & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         Trim(Str(BREC10!SGI_CODIGO)) & Trim(strCodOper) & vbTab & _
                         Trim(strOrdem) & Trim(strCodOper)
                BREC10.MoveNext
            Loop
        End With
    End If
    BREC10.Close
    
End Sub
