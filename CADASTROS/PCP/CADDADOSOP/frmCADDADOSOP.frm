VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCADDADOSOP 
   Caption         =   "Cadastra Dados da OP"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11130
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   11130
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Caption         =   "[ Cores ]"
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
      Height          =   3615
      Left            =   0
      TabIndex        =   14
      Top             =   3840
      Width           =   11055
      Begin VSFlex8LCtl.VSFlexGrid grdCORES 
         Height          =   3255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   10815
         _cx             =   19076
         _cy             =   5741
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
   Begin VB.Frame Frame3 
      Caption         =   "[ Verniz ]"
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
      Height          =   2415
      Left            =   0
      TabIndex        =   12
      Top             =   1440
      Width           =   11055
      Begin VSFlex8LCtl.VSFlexGrid grdVERNIZ 
         Height          =   2055
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   10815
         _cx             =   19076
         _cy             =   3625
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
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   11055
      Begin VB.TextBox txtCodOrdFat 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6360
         TabIndex        =   10
         Text            =   "txtCodOrdFat"
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdPedido 
         Height          =   315
         Left            =   7680
         Picture         =   "frmCADDADOSOP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskDATALOTE 
         Height          =   285
         Left            =   3840
         TabIndex        =   8
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblCODPROD 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblCODPROD"
         Height          =   285
         Left            =   8880
         TabIndex        =   17
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Produto"
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
         Index           =   4
         Left            =   8160
         TabIndex        =   16
         Top             =   120
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código da OP"
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
         Index           =   2
         Left            =   5160
         TabIndex        =   11
         Top             =   120
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data do Lote"
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
         Index           =   1
         Left            =   2640
         TabIndex        =   7
         Top             =   120
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código da Lote"
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
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   150
         Width           =   1305
      End
      Begin VB.Label lblCODIGO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblCODIGO"
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11055
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
         Picture         =   "frmCADDADOSOP.frx":0102
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
         Picture         =   "frmCADDADOSOP.frx":0634
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
         Picture         =   "frmCADDADOSOP.frx":0736
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADDADOSOP"
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
Dim lngCodLog           As Long
Dim objBLBFunc          As Object
Dim objCADDADOSOPC      As Object
Dim objPESQPADRAO       As Object

Dim arrVERNIZ           As Variant
Dim arrCORES            As Variant

'' --------------------------------------------------------------------
Const conCOL_OrdVerniz_IdProduto                 As Integer = 0
Const conCOL_OrdVerniz_VernzDescr                As Integer = 1
Const conCOL_OrdVerniz_Produto                   As Integer = 2
Const conCOL_OrdVerniz_Descricao                 As Integer = 3
Const conCOL_OrdVerniz_Lote                      As Integer = 4
Const conCOL_OrdVerniz_Producao                  As Integer = 5
Const conCOL_OrdVerniz_Data                      As Integer = 6
Const conCOL_OrdVerniz_Maquina                   As Integer = 7
Const conCOL_OrdVerniz_PesqMaq                   As Integer = 8
Const conCOL_OrdVerniz_DescMaq                   As Integer = 9
Const conCOL_OrdVerniz_Turno                     As Integer = 10
Const conCOL_OrdVerniz_PesqTur                   As Integer = 11
Const conCOL_OrdVerniz_DescTur                   As Integer = 12

Const conLIN_OrdVerniz_VernInt                   As Integer = 1
Const conLIN_OrdVerniz_Esmalte                   As Integer = 2
Const conLIN_OrdVerniz_VernAcab                  As Integer = 3

Const conCOL_OrdVerniz_FormatString              As String = "=IDProd| |Produto|Descrição|Lote|Produção|Data|Máquina|...|Descr.Máquina|Turno|...|Desc.Turno"
Const conColumnsIn_OrdVerniz                     As Integer = 13

'' --------------------------------------------------------------------
Const conCOL_OrdCores_IdProduto                 As Integer = 0
Const conCOL_OrdCores_VernzDescr                As Integer = 1
Const conCOL_OrdCores_Produto                   As Integer = 2
Const conCOL_OrdCores_Descricao                 As Integer = 3
Const conCOL_OrdCores_Lote                      As Integer = 4
Const conCOL_OrdCores_Producao                  As Integer = 5
Const conCOL_OrdCores_Data                      As Integer = 6
Const conCOL_OrdCores_Maquina                   As Integer = 7
Const conCOL_OrdCores_PesqMaq                   As Integer = 8
Const conCOL_OrdCores_DescMaq                   As Integer = 9
Const conCOL_OrdCores_Turno                     As Integer = 10
Const conCOL_OrdCores_PesqTur                   As Integer = 11
Const conCOL_OrdCores_DescTur                   As Integer = 12

Const conCOL_OrdCores_FormatString              As String = "=IDProd| |Produto|Descrição|Lote|Produção|Data|Máquina|...|Descr.Máquina|Turno|...|Desc.Turno"
Const conColumnsIn_OrdCores                     As Integer = 13
'' --------------------------------------------------------------------


Private Sub cmdAltera_Click()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    
    Me.Caption = "Cadastra Dados da OP - [ ALTERAÇÃO ]"
    
    Frame2.Enabled = False
    
    cTipOper = "A"

End Sub

Private Sub cmdPedido_Click()

    ReDim arrCAMPOS(1 To 4, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       ORD.SGI_CODIGO " & vbCrLf
    sSql = sSql & "      ,ORD.SGI_CODPED " & vbCrLf
    sSql = sSql & "      ,ORD.SGI_CODPROD " & vbCrLf
    sSql = sSql & "      ,CLIE.SGI_RAZAOSOC " & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD    ORD " & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH  VENDH" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE   CLIE" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       ORD.SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "   And VENDH.SGI_FILIAL = ORD.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And VENDH.SGI_CODIGO = ORD.SGI_CODPED " & vbCrLf
    sSql = sSql & "   And CLIE.SGI_FILIAL  = VENDH.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLIE.SGI_CODIGO  = VENDH.SGI_CODCLI "
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Cód.OP"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "ORD.SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_CODPED"
    arrCAMPOS(2, 2) = "N"
    arrCAMPOS(2, 3) = "Cód.Pedido"
    arrCAMPOS(2, 4) = "1000"
    arrCAMPOS(2, 5) = "ORD.SGI_CODPED"
    
    arrCAMPOS(3, 1) = "SGI_CODPROD"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "Produto"
    arrCAMPOS(3, 4) = "1500"
    arrCAMPOS(3, 5) = "ORD.SGI_CODPROD"
    
    arrCAMPOS(4, 1) = "SGI_RAZAOSOC"
    arrCAMPOS(4, 2) = "S"
    arrCAMPOS(4, 3) = "Razão Social"
    arrCAMPOS(4, 4) = "4500"
    arrCAMPOS(4, 5) = "CLIE.SGI_RAZAOSOC"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Pesquisa Ordem de Produção", "")
    
    If Len(Trim(varRETORNO)) > 0 Then
       txtCodOrdFat.Text = varRETORNO
       Call txtCodOrdFat_Validate(False)
    End If


End Sub

Private Sub CmdSalva_Click()

On Error GoTo errGrava

    Dim I As Integer
    
    If ConsisteCampos = False Then Exit Sub

    If cTipOper = "I" Then objCADDADOSOPC.CODIGO = objBLBFunc.Gera_Codigo(Me.Name, FILIAL, Linha) & Year(Now)
    
    objCADDADOSOPC.DATALOTE = CDate(mskDATALOTE.Text)
    objCADDADOSOPC.CODOP = CLng(txtCodOrdFat.Text)
    
    '' ----------------------------------------------
    '' Gravando a Verniz
    arrVERNIZ = Empty
    With grdVERNIZ
        If (.Rows - 1) > 0 Then
            ReDim arrVERNIZ(1 To (.Rows - 1), 1 To 7) As String
            For I = 1 To (.Rows - 1)
                If Len(Trim(.Cell(flexcpText, I, conCOL_OrdVerniz_IdProduto))) > 0 Then
                    arrVERNIZ(I, 1) = .Cell(flexcpText, I, conCOL_OrdVerniz_IdProduto)
                    arrVERNIZ(I, 2) = .Cell(flexcpText, I, conCOL_OrdVerniz_Lote)
                    arrVERNIZ(I, 3) = .Cell(flexcpText, I, conCOL_OrdVerniz_Producao)
                    arrVERNIZ(I, 4) = Format(CDate(.Cell(flexcpText, I, conCOL_OrdVerniz_Data)), "MM/DD/YYYY")
                    arrVERNIZ(I, 5) = .Cell(flexcpText, I, conCOL_OrdVerniz_Maquina)
                    arrVERNIZ(I, 6) = .Cell(flexcpText, I, conCOL_OrdVerniz_Turno)
                    arrVERNIZ(I, 7) = I
                End If
            Next I
        End If
    End With
    objCADDADOSOPC.VERNIZ = arrVERNIZ
    '' ----------------------------------------------
    
    '' ----------------------------------------------
    '' Gravando Cores
    arrCORES = Empty
    With grdCORES
        If (.Rows - 1) > 0 Then
            ReDim arrCORES(1 To (.Rows - 1), 1 To 6) As String
            For I = 1 To (.Rows - 1)
                arrCORES(I, 1) = .Cell(flexcpText, I, conCOL_OrdCores_IdProduto)
                arrCORES(I, 2) = .Cell(flexcpText, I, conCOL_OrdCores_Lote)
                arrCORES(I, 3) = .Cell(flexcpText, I, conCOL_OrdCores_Producao)
                arrCORES(I, 4) = Format(CDate(.Cell(flexcpText, I, conCOL_OrdCores_Data)), "MM/DD/YYYY")
                arrCORES(I, 5) = .Cell(flexcpText, I, conCOL_OrdCores_Maquina)
                arrCORES(I, 6) = .Cell(flexcpText, I, conCOL_OrdCores_Turno)
            Next I
        End If
    End With
    objCADDADOSOPC.CORES = arrCORES
    '' ----------------------------------------------
    
    '' Gravando as Informações no banco
    If objCADDADOSOPC.GRAVA(cTipOper) = False Then Exit Sub
    
    '' Atualizando os Dados
    If objBLBFunc.Atualiza(cTipOper, Str(objCADDADOSOPC.CODIGO), FILIAL, Me.Name, Linha) = False Then Exit Sub
    
    '' Gerand Log de Sistema
    lngCodLog = objBLBFunc.Gera_Codigo("SGI_LOGMODULO", FILIAL, Linha)
    Call objBLBFunc.GravaLogModulo(FILIAL, lngCodLog, Me.Name, cTipOper, lngCodUsuario, Str(objCADDADOSOPC.CODIGO), Linha)
    
    MsgBox "Os Dados da OP ( " & Trim(Str(objCADDADOSOPC.CODIGO)) & " ) foi " & IIf(cTipOper = "I", "incluso", IIf(cTipOper = "A", "alterado", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
    
    Unload Me

    Exit Sub

errGrava:

    MsgBox "Erro Nº : " & Err.Number & vbCrLf & _
           "Descr.  : " & Err.Description, vbOKOnly + vbExclamation, "Aviso"

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

Private Sub Form_Load()

   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADDADOSOPC = CreateObject("CADDADOSOP.clsCADDADOSOP")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
   objCADDADOSOPC.FILIAL = FILIAL
   
   If adoBanco_Dados.State = 0 Then Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)
   If adoBanco_Dados.State = 0 Then
      MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
      Exit Sub
   End If
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta


End Sub

Private Sub FechaOBJ()
    Set objBLBFunc = Nothing
    Set objCADDADOSOPC = Nothing
    Set objPESQPADRAO = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call FechaOBJ
End Sub


Private Sub ConfGridVerniz()

    Dim strINC_CAMPOS As String
    
    With grdVERNIZ
    
       .Cols = conColumnsIn_OrdVerniz
       .Rows = 1
       .FixedCols = 4
       .FormatString = conCOL_OrdVerniz_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeNone
       
       .Cell(flexcpData, 0, conCOL_OrdVerniz_IdProduto) = ""
       .ColDataType(conCOL_OrdVerniz_IdProduto) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdVerniz_VernzDescr) = ""
       .ColDataType(conCOL_OrdVerniz_VernzDescr) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_OrdVerniz_Produto) = ""
       .ColDataType(conCOL_OrdVerniz_Produto) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_OrdVerniz_Descricao) = ""
       .ColDataType(conCOL_OrdVerniz_Descricao) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_OrdVerniz_Lote) = ""
       .ColDataType(conCOL_OrdVerniz_Lote) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_OrdVerniz_Producao) = ""
       .ColDataType(conCOL_OrdVerniz_Producao) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdVerniz_Data) = ""
       .ColDataType(conCOL_OrdVerniz_Data) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_OrdVerniz_Maquina) = ""
       .ColDataType(conCOL_OrdVerniz_Maquina) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdVerniz_PesqMaq) = ""
       .ColDataType(conCOL_OrdVerniz_PesqMaq) = flexDTString
       .ColComboList(conCOL_OrdVerniz_PesqMaq) = "..."
       
       .Cell(flexcpData, 0, conCOL_OrdVerniz_DescMaq) = ""
       .ColDataType(conCOL_OrdVerniz_DescMaq) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_OrdVerniz_Lote) = ""
       .ColDataType(conCOL_OrdVerniz_Lote) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdVerniz_PesqTur) = ""
       .ColDataType(conCOL_OrdVerniz_PesqTur) = flexDTString
       .ColComboList(conCOL_OrdVerniz_PesqTur) = "..."
       
       .Cell(flexcpData, 0, conCOL_OrdVerniz_DescTur) = ""
       .ColDataType(conCOL_OrdVerniz_DescTur) = flexDTString
       
       .ColWidth(conCOL_OrdVerniz_IdProduto) = 0
       .ColWidth(conCOL_OrdVerniz_VernzDescr) = 1000
       .ColWidth(conCOL_OrdVerniz_Produto) = 800
       .ColWidth(conCOL_OrdVerniz_Descricao) = 3500
       .ColWidth(conCOL_OrdVerniz_Lote) = 2000
       .ColWidth(conCOL_OrdVerniz_Producao) = 800
       .ColWidth(conCOL_OrdVerniz_Data) = 1000
       .ColWidth(conCOL_OrdVerniz_Maquina) = 800
       .ColWidth(conCOL_OrdVerniz_PesqMaq) = 300
       .ColWidth(conCOL_OrdVerniz_DescMaq) = 2000
       .ColWidth(conCOL_OrdVerniz_Turno) = 800
       .ColWidth(conCOL_OrdVerniz_PesqTur) = 300
       .ColWidth(conCOL_OrdVerniz_DescTur) = 2000
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
       strINC_CAMPOS = "" & vbTab & _
                       "" & vbTab & _
                       "" & vbTab & _
                       "" & vbTab & _
                       "" & vbTab & _
                       "" & vbTab & _
                       "" & vbTab & _
                       "" & vbTab & _
                       "" & vbTab & _
                       "" & vbTab & _
                       "" & vbTab & _
                       "" & vbTab & _
                       ""
       
       .AddItem strINC_CAMPOS
       .AddItem strINC_CAMPOS
       .AddItem strINC_CAMPOS
       
       .Cell(flexcpText, conLIN_OrdVerniz_VernInt, conCOL_OrdVerniz_VernzDescr) = "Vern.Interno"
       .Cell(flexcpText, conLIN_OrdVerniz_Esmalte, conCOL_OrdVerniz_VernzDescr) = "Esmalte"
       .Cell(flexcpText, conLIN_OrdVerniz_VernAcab, conCOL_OrdVerniz_VernzDescr) = "Vern.Acab"
    
    End With
    
End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    lblCODIGO.Caption = ""
    
    Frame2.Enabled = True
    
    Me.Caption = "Cadastra Dados da OP - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADDADOSOP
    Call LimpaCamposLabel
    
    mskDATALOTE.Text = Format(Now, "DD/MM/YYYY")
    
    Call ConfGridVerniz
    Call ConfGridCores
    
End Sub


Private Sub ConfGridCores()

    With grdCORES
    
       .Cols = conColumnsIn_OrdCores
       .Rows = 1
       .FixedCols = 4
       .FormatString = conCOL_OrdCores_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeNone
       
       .Cell(flexcpData, 0, conCOL_OrdCores_IdProduto) = ""
       .ColDataType(conCOL_OrdCores_IdProduto) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdCores_VernzDescr) = ""
       .ColDataType(conCOL_OrdCores_VernzDescr) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_OrdCores_Produto) = ""
       .ColDataType(conCOL_OrdCores_Produto) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_OrdCores_Descricao) = ""
       .ColDataType(conCOL_OrdCores_Descricao) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_OrdCores_Lote) = ""
       .ColDataType(conCOL_OrdCores_Lote) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_OrdCores_Producao) = ""
       .ColDataType(conCOL_OrdCores_Producao) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdCores_Data) = ""
       .ColDataType(conCOL_OrdCores_Data) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_OrdCores_Maquina) = ""
       .ColDataType(conCOL_OrdCores_Maquina) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdCores_PesqMaq) = ""
       .ColDataType(conCOL_OrdCores_PesqMaq) = flexDTString
       .ColComboList(conCOL_OrdCores_PesqMaq) = "..."
       
       .Cell(flexcpData, 0, conCOL_OrdCores_DescMaq) = ""
       .ColDataType(conCOL_OrdCores_DescMaq) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_OrdCores_Lote) = ""
       .ColDataType(conCOL_OrdCores_Lote) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdCores_PesqTur) = ""
       .ColDataType(conCOL_OrdCores_PesqTur) = flexDTString
       .ColComboList(conCOL_OrdCores_PesqTur) = "..."
       
       .Cell(flexcpData, 0, conCOL_OrdCores_DescTur) = ""
       .ColDataType(conCOL_OrdCores_DescTur) = flexDTString
       
       .ColWidth(conCOL_OrdCores_IdProduto) = 0
       .ColWidth(conCOL_OrdCores_VernzDescr) = 800
       .ColWidth(conCOL_OrdCores_Produto) = 800
       .ColWidth(conCOL_OrdCores_Descricao) = 3500
       .ColWidth(conCOL_OrdCores_Lote) = 2000
       .ColWidth(conCOL_OrdCores_Producao) = 800
       .ColWidth(conCOL_OrdCores_Data) = 1000
       .ColWidth(conCOL_OrdCores_Maquina) = 800
       .ColWidth(conCOL_OrdCores_PesqMaq) = 300
       .ColWidth(conCOL_OrdCores_DescMaq) = 2000
       .ColWidth(conCOL_OrdCores_Turno) = 800
       .ColWidth(conCOL_OrdCores_PesqTur) = 300
       .ColWidth(conCOL_OrdCores_DescTur) = 2000
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
    End With
    
End Sub


Private Sub grdCORES_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    Select Case Col
    Case conCOL_OrdCores_DescMaq, _
         conCOL_OrdCores_DescTur
         Cancel = True
    Case conCOL_OrdCores_Lote, _
         conCOL_OrdCores_Producao, _
         conCOL_OrdCores_Data, _
         conCOL_OrdCores_Maquina, _
         conCOL_OrdCores_PesqMaq, _
         conCOL_OrdCores_Turno, _
         conCOL_OrdCores_PesqTur
         If Len(Trim(grdCORES.Cell(flexcpText, Row, conCOL_OrdCores_Produto))) = 0 Then Cancel = True
         If cTipOper = "C" Then Cancel = True
    Case Else
        grdCORES.ComboList = ""
    End Select
    Exit Sub

End Sub

Private Sub grdCORES_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    Select Case Col
        Case conCOL_OrdCores_PesqMaq
    
            If cTipOper = "C" Then Exit Sub
    
            ReDim arrCAMPOS(1 To 2, 1 To 5) As String
            ReDim arrTABELA(1 To 1) As String
    
            sSql = ""
            
            sSql = "Select " & vbCrLf
            sSql = sSql & "       MAQ.SGI_CODIGO " & vbCrLf
            sSql = sSql & "      ,MAQ.SGI_DESCRI " & vbCrLf
            sSql = sSql & "  from " & vbCrLf
            sSql = sSql & "       SGI_CADMAQUINA MAQ" & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       MAQ.SGI_FILIAL = " & FILIAL
    
            arrTABELA(1) = sSql
            
            arrCAMPOS(1, 1) = "SGI_CODIGO"
            arrCAMPOS(1, 2) = "S"
            arrCAMPOS(1, 3) = "Código"
            arrCAMPOS(1, 4) = "1000"
            arrCAMPOS(1, 5) = "MAQ.SGI_CODIGO"
            
            arrCAMPOS(2, 1) = "SGI_DESCRI"
            arrCAMPOS(2, 2) = "S"
            arrCAMPOS(2, 3) = "Descrição"
            arrCAMPOS(2, 4) = "5000"
            arrCAMPOS(2, 5) = "MAQ.SGI_DESCRI"
            
            varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Máquinas")
    
            If Len(Trim(varRETORNO)) > 0 Then
                grdCORES.Cell(flexcpText, grdCORES.Row, conCOL_OrdCores_Maquina) = varRETORNO
                Call ConfMaquina(Trim(varRETORNO), grdCORES, conCOL_OrdCores_DescMaq)
            End If
        Case conCOL_OrdCores_PesqTur
    
            If cTipOper = "C" Then Exit Sub
    
            ReDim arrCAMPOS(1 To 2, 1 To 5) As String
            ReDim arrTABELA(1 To 1) As String
    
            sSql = ""
            
            sSql = "Select " & vbCrLf
            sSql = sSql & "       TURN.SGI_CODIGO " & vbCrLf
            sSql = sSql & "      ,TURN.SGI_DESCRI " & vbCrLf
            sSql = sSql & "  from " & vbCrLf
            sSql = sSql & "       SGI_CADQTDETURN TURN" & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       TURN.SGI_FILIAL = " & FILIAL
    
            arrTABELA(1) = sSql
            
            arrCAMPOS(1, 1) = "SGI_CODIGO"
            arrCAMPOS(1, 2) = "S"
            arrCAMPOS(1, 3) = "Código"
            arrCAMPOS(1, 4) = "1000"
            arrCAMPOS(1, 5) = "TURN.SGI_CODIGO"
            
            arrCAMPOS(2, 1) = "SGI_DESCRI"
            arrCAMPOS(2, 2) = "S"
            arrCAMPOS(2, 3) = "Descrição"
            arrCAMPOS(2, 4) = "5000"
            arrCAMPOS(2, 5) = "TURN.SGI_DESCRI"
            
            varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Turnos")
    
            If Len(Trim(varRETORNO)) > 0 Then
                grdCORES.Cell(flexcpText, grdCORES.Row, conCOL_OrdCores_Turno) = varRETORNO
                Call ConfTurno(Trim(varRETORNO), grdCORES, conCOL_OrdCores_DescTur)
            End If
    
    End Select

End Sub

Private Sub grdCORES_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdCORES
          Select Case Col
                    Case conCOL_OrdCores_Lote
                         KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
                    Case conCOL_OrdCores_Producao, _
                         conCOL_OrdCores_Maquina, _
                         conCOL_OrdCores_Turno
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
                    Case conCOL_OrdCores_Data
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsDate)
          End Select
     End With
End Sub

Private Sub grdCORES_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
     With grdCORES
          Select Case Col
                 Case conCOL_OrdCores_Maquina, _
                      conCOL_OrdCores_Turno
                    If .EditText = Empty Then Exit Sub
                    If Col = conCOL_OrdCores_Maquina Then Cancel = ConfMaquina(Trim(.EditText), grdCORES, conCOL_OrdCores_DescMaq)
                    If Col = conCOL_OrdCores_Turno Then Cancel = ConfTurno(Trim(.EditText), grdCORES, conCOL_OrdCores_DescTur)
                 Case conCOL_OrdCores_Lote
                    If Len(Trim(.EditText)) > 50 Then
                        MsgBox "Somente é permitido 50 Digitos !!!", vbOKOnly + vbExclamation, "Aviso"
                        Cancel = True
                        Exit Sub
                    End If
          End Select
     End With

End Sub

Private Sub grdVERNIZ_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_OrdVerniz_DescMaq, _
         conCOL_OrdVerniz_DescTur
         Cancel = True
    Case conCOL_OrdVerniz_Lote, _
         conCOL_OrdVerniz_Producao, _
         conCOL_OrdVerniz_Data, _
         conCOL_OrdVerniz_Maquina, _
         conCOL_OrdVerniz_PesqMaq, _
         conCOL_OrdVerniz_Turno, _
         conCOL_OrdVerniz_PesqTur
         If Len(Trim(grdVERNIZ.Cell(flexcpText, Row, conCOL_OrdVerniz_Produto))) = 0 Then Cancel = True
         If cTipOper = "C" Then Cancel = True
    Case Else
        grdVERNIZ.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub grdVERNIZ_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    Select Case Col
        Case conCOL_OrdVerniz_PesqMaq
    
            If cTipOper = "C" Then Exit Sub
    
            ReDim arrCAMPOS(1 To 2, 1 To 5) As String
            ReDim arrTABELA(1 To 1) As String
    
            sSql = ""
            
            sSql = "Select " & vbCrLf
            sSql = sSql & "       MAQ.SGI_CODIGO " & vbCrLf
            sSql = sSql & "      ,MAQ.SGI_DESCRI " & vbCrLf
            sSql = sSql & "  from " & vbCrLf
            sSql = sSql & "       SGI_CADMAQUINA MAQ" & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       MAQ.SGI_FILIAL = " & FILIAL
    
            arrTABELA(1) = sSql
            
            arrCAMPOS(1, 1) = "SGI_CODIGO"
            arrCAMPOS(1, 2) = "S"
            arrCAMPOS(1, 3) = "Código"
            arrCAMPOS(1, 4) = "1000"
            arrCAMPOS(1, 5) = "MAQ.SGI_CODIGO"
            
            arrCAMPOS(2, 1) = "SGI_DESCRI"
            arrCAMPOS(2, 2) = "S"
            arrCAMPOS(2, 3) = "Descrição"
            arrCAMPOS(2, 4) = "5000"
            arrCAMPOS(2, 5) = "MAQ.SGI_DESCRI"
            
            varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Máquinas")
    
            If Len(Trim(varRETORNO)) > 0 Then
                grdVERNIZ.Cell(flexcpText, grdVERNIZ.Row, conCOL_OrdVerniz_Maquina) = varRETORNO
                Call ConfMaquina(Trim(varRETORNO), grdVERNIZ, conCOL_OrdVerniz_DescMaq)
            End If
        Case conCOL_OrdVerniz_PesqTur
    
            If cTipOper = "C" Then Exit Sub
    
            ReDim arrCAMPOS(1 To 2, 1 To 5) As String
            ReDim arrTABELA(1 To 1) As String
    
            sSql = ""
            
            sSql = "Select " & vbCrLf
            sSql = sSql & "       TURN.SGI_CODIGO " & vbCrLf
            sSql = sSql & "      ,TURN.SGI_DESCRI " & vbCrLf
            sSql = sSql & "  from " & vbCrLf
            sSql = sSql & "       SGI_CADQTDETURN TURN" & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       TURN.SGI_FILIAL = " & FILIAL
    
            arrTABELA(1) = sSql
            
            arrCAMPOS(1, 1) = "SGI_CODIGO"
            arrCAMPOS(1, 2) = "S"
            arrCAMPOS(1, 3) = "Código"
            arrCAMPOS(1, 4) = "1000"
            arrCAMPOS(1, 5) = "TURN.SGI_CODIGO"
            
            arrCAMPOS(2, 1) = "SGI_DESCRI"
            arrCAMPOS(2, 2) = "S"
            arrCAMPOS(2, 3) = "Descrição"
            arrCAMPOS(2, 4) = "5000"
            arrCAMPOS(2, 5) = "TURN.SGI_DESCRI"
            
            varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Turnos")
    
            If Len(Trim(varRETORNO)) > 0 Then
                grdVERNIZ.Cell(flexcpText, grdVERNIZ.Row, conCOL_OrdVerniz_Turno) = varRETORNO
                Call ConfTurno(Trim(varRETORNO), grdVERNIZ, conCOL_OrdVerniz_DescTur)
            End If
    
    End Select

End Sub

Private Sub grdVERNIZ_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdVERNIZ
          Select Case Col
                    Case conCOL_OrdVerniz_Lote
                         KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
                    Case conCOL_OrdVerniz_Producao, _
                         conCOL_OrdVerniz_Maquina, _
                         conCOL_OrdVerniz_Turno
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
                    Case conCOL_OrdVerniz_Data
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsDate)
          End Select
     End With
End Sub


Private Sub grdVERNIZ_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
     With grdVERNIZ
          Select Case Col
                 Case conCOL_OrdVerniz_Maquina, _
                      conCOL_OrdVerniz_Turno
                    If .EditText = Empty Then Exit Sub
                    If Col = conCOL_OrdVerniz_Maquina Then Cancel = ConfMaquina(Trim(.EditText), grdVERNIZ, conCOL_OrdVerniz_DescMaq)
                    If Col = conCOL_OrdVerniz_Turno Then Cancel = ConfTurno(Trim(.EditText), grdVERNIZ, conCOL_OrdVerniz_DescTur)
                 Case conCOL_OrdVerniz_Lote
                    If Len(Trim(.EditText)) > 50 Then
                        MsgBox "Somente é permitido 50 Digitos !!!", vbOKOnly + vbExclamation, "Aviso"
                        Cancel = True
                        Exit Sub
                    End If
          End Select
     End With
End Sub

Private Sub mskDATALOTE_GotFocus()
    objBLBFunc.SelecionaCampos mskDATALOTE.Name, frmCADDADOSOP
End Sub

Private Function ConfMaquina(strCODIGO As String, grdGENERICO As VSFlexGrid, lngCOL As Long) As Boolean

    ConfMaquina = False
    
    If Len(Trim(strCODIGO)) = 0 Then Exit Function
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADMAQUINA " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & Trim(strCODIGO)

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF() Then
       MsgBox "Esta Maquina Não Existe !!!", vbOKOnly + vbExclamation, "Aviso"
       ConfMaquina = True
    Else
        grdGENERICO.Cell(flexcpText, grdGENERICO.Row, lngCOL) = Trim(BREC!SGI_DESCRI)
    End If
    BREC.Close
    
End Function


Private Function ConfTurno(strCODIGO As String, grdGENERICO As VSFlexGrid, lngCOL As Long) As Boolean

    ConfTurno = False
    
    If Len(Trim(strCODIGO)) = 0 Then Exit Function
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADQTDETURN " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & Trim(strCODIGO)

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF() Then
       MsgBox "Este Turno Não Existe !!!", vbOKOnly + vbExclamation, "Aviso"
       ConfTurno = True
    Else
        grdGENERICO.Cell(flexcpText, grdGENERICO.Row, lngCOL) = Trim(BREC!SGI_DESCRI)
    End If
    BREC.Close
    
End Function

Private Sub txtCodOrdFat_GotFocus()
    objBLBFunc.SelecionaCampos txtCodOrdFat.Name, frmCADDADOSOP
End Sub

Private Sub txtCodOrdFat_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCodOrdFat.Text
End Sub

Private Sub txtCodOrdFat_Validate(Cancel As Boolean)
    
    If Len(Trim(txtCodOrdFat.Text)) = 0 Then Exit Sub
    
    Cancel = PedaDadosOrdem(Trim(txtCodOrdFat.Text))

End Sub

Private Function PedaDadosOrdem(strCODORD As String) As Boolean

    PedaDadosOrdem = False
    
    If Len(Trim(strCODORD)) = 0 Then Exit Function
    
    Dim intQTDCORES As Integer
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       ORDP.* " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD ORDP" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       ORDP.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And ORDP.SGI_CODIGO = " & strCODORD & vbCrLf
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF() Then
       MsgBox "Esta Ordem de Produção não existe !!!", vbOKOnly + vbExclamation, "Aviso"
       PedaDadosOrdem = True
    Else
    
        lblCODPROD.Caption = Trim(BREC!SGI_CODPROD)
        objCADDADOSOPC.IDPRODUTO = BREC!SGI_IDPRODUTO
        
        Call ConfGridCores
        
        '' -------------------------------------------------------
        '' Pega o Veniz Interno do Produto
        sSql = "Select " & vbCrLf
        sSql = sSql & "       VERN.* " & vbCrLf
        sSql = sSql & "      ,PROD.SGI_CODIGO " & vbCrLf
        sSql = sSql & "      ,PROD.SGI_DESCRICAO " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_VERNIZPROD VERN" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO PROD" & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       VERN.SGI_FILIAL    = " & FILIAL & vbCrLf
        sSql = sSql & "   And VERN.SGI_IDPRODUTO = " & BREC!SGI_IDPRODUTO & vbCrLf
        sSql = sSql & "   And PROD.SGI_FILIAL    = VERN.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And PROD.SGI_IDPRODUTO = VERN.SGI_PRODUTO "
        
        BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC2.EOF() Then
            With grdVERNIZ
                .Cell(flexcpText, conLIN_OrdVerniz_VernInt, conCOL_OrdVerniz_IdProduto) = BREC2!SGI_PRODUTO
                .Cell(flexcpText, conLIN_OrdVerniz_VernInt, conCOL_OrdVerniz_Produto) = BREC2!SGI_CODIGO
                .Cell(flexcpText, conLIN_OrdVerniz_VernInt, conCOL_OrdVerniz_Descricao) = BREC2!SGI_DESCRICAO
            End With
        End If
        BREC2.Close
    
        '' -------------------------------------------------------
        '' Esmalte
        sSql = "Select " & vbCrLf
        sSql = sSql & "       ESM.* " & vbCrLf
        sSql = sSql & "      ,PROD.SGI_CODIGO " & vbCrLf
        sSql = sSql & "      ,PROD.SGI_DESCRICAO " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_ESMALTEPROD ESM" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO PROD" & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       ESM.SGI_FILIAL    = " & FILIAL & vbCrLf
        sSql = sSql & "   And ESM.SGI_IDPRODUTO = " & BREC!SGI_IDPRODUTO & vbCrLf
        sSql = sSql & "   And PROD.SGI_FILIAL    = ESM.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And PROD.SGI_IDPRODUTO = ESM.SGI_PRODUTO "
        
        BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC2.EOF() Then
            With grdVERNIZ
                .Cell(flexcpText, conLIN_OrdVerniz_Esmalte, conCOL_OrdVerniz_IdProduto) = BREC2!SGI_PRODUTO
                .Cell(flexcpText, conLIN_OrdVerniz_Esmalte, conCOL_OrdVerniz_Produto) = BREC2!SGI_CODIGO
                .Cell(flexcpText, conLIN_OrdVerniz_Esmalte, conCOL_OrdVerniz_Descricao) = BREC2!SGI_DESCRICAO
            End With
        End If
        BREC2.Close
    
        '' -------------------------------------------------------
        '' Verniz Acabamento
        sSql = "Select " & vbCrLf
        sSql = sSql & "       VEP.* " & vbCrLf
        sSql = sSql & "      ,PROD.SGI_CODIGO " & vbCrLf
        sSql = sSql & "      ,PROD.SGI_DESCRICAO " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_VERNIZPRODACAB VEP" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO PROD" & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       VEP.SGI_FILIAL    = " & FILIAL & vbCrLf
        sSql = sSql & "   And VEP.SGI_IDPRODUTO = " & BREC!SGI_IDPRODUTO & vbCrLf
        sSql = sSql & "   And PROD.SGI_FILIAL    = VEP.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And PROD.SGI_IDPRODUTO = VEP.SGI_PRODUTO "
        
        BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC2.EOF() Then
            With grdVERNIZ
                .Cell(flexcpText, conLIN_OrdVerniz_VernAcab, conCOL_OrdVerniz_IdProduto) = BREC2!SGI_PRODUTO
                .Cell(flexcpText, conLIN_OrdVerniz_VernAcab, conCOL_OrdVerniz_Produto) = BREC2!SGI_CODIGO
                .Cell(flexcpText, conLIN_OrdVerniz_VernAcab, conCOL_OrdVerniz_Descricao) = BREC2!SGI_DESCRICAO
            End With
        End If
        BREC2.Close
        
        
        '' -------------------------------------------------------
        '' Cores
        sSql = "Select " & vbCrLf
        sSql = sSql & "       COR.* " & vbCrLf
        sSql = sSql & "      ,PROD.SGI_CODIGO " & vbCrLf
        sSql = sSql & "      ,PROD.SGI_DESCRICAO " & vbCrLf
        
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CORESPROD  COR " & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO PROD " & vbCrLf
        
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       COR.SGI_FILIAL     = " & FILIAL & vbCrLf
        sSql = sSql & "   And COR.SGI_IDPRODUTO  = " & BREC!SGI_IDPRODUTO & vbCrLf
        sSql = sSql & "   And PROD.SGI_FILIAL    = COR.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And PROD.SGI_IDPRODUTO = COR.SGI_CODCOR "
    
        BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC2.EOF() Then
            With grdCORES
                intQTDCORES = 1
                Do While Not BREC2.EOF()
                
                    .AddItem BREC2!SGI_CODCOR & vbTab & _
                             Trim(Str(intQTDCORES)) & "º Cor" & vbTab & _
                             BREC2!SGI_CODIGO & vbTab & _
                             BREC2!SGI_DESCRICAO & vbTab & _
                             ""
                             
                    intQTDCORES = (intQTDCORES + 1)
                    BREC2.MoveNext
                Loop
            End With
        End If
        BREC2.Close
            
    End If
    BREC.Close
    

End Function

Private Sub LimpaCamposLabel()
    lblCODIGO.Caption = ""
    lblCODPROD.Caption = ""
End Sub

Private Function ConsisteCampos() As Boolean
    ConsisteCampos = False
    
    Dim I       As Integer
    Dim qtdNull As Integer
    
    If Not IsDate(mskDATALOTE.Text) Then
        MsgBox "Daya do Lote inválida !!!", vbOKOnly + vbExclamation, "Aviso"
        mskDATALOTE.SetFocus
        Exit Function
    End If
    
    '' --------------------------------
    '' Verniz
    qtdNull = 0
    With grdVERNIZ
        For I = 1 To (.Rows - 1)
            If Len(Trim(.Cell(flexcpText, I, conCOL_OrdVerniz_IdProduto))) = 0 Then qtdNull = (qtdNull + 1)
            If Len(Trim(.Cell(flexcpText, I, conCOL_OrdVerniz_IdProduto))) > 1 Then
                If Len(Trim(.Cell(flexcpText, I, conCOL_OrdVerniz_Lote))) = 0 Then
                    MsgBox "Informe o Lote !!!", vbOKOnly + vbExclamation, "Aviso"
                    Exit Function
                End If
                If Len(Trim(.Cell(flexcpText, I, conCOL_OrdVerniz_Producao))) = 0 Then
                    MsgBox "Informe a Produção !!!", vbOKOnly + vbExclamation, "Aviso"
                    Exit Function
                End If
                If Not IsDate(.Cell(flexcpText, I, conCOL_OrdVerniz_Data)) Then
                    MsgBox "Informe a data da Produção !!!", vbOKOnly + vbExclamation, "Aviso"
                    Exit Function
                End If
                If Len(Trim(.Cell(flexcpText, I, conCOL_OrdVerniz_Maquina))) = 0 Then
                    MsgBox "Informe a Máquina !!!", vbOKOnly + vbExclamation, "Aviso"
                    Exit Function
                End If
                If Len(Trim(.Cell(flexcpText, I, conCOL_OrdVerniz_Turno))) = 0 Then
                    MsgBox "Informe o Turno !!!", vbOKOnly + vbExclamation, "Aviso"
                    Exit Function
                End If
            End If
        Next I
        If qtdNull = (.Rows - 1) Then
            MsgBox "Não foi informado nenhum verniz !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Function
        End If
    End With
    
    '' --------------------------------
    '' Cores
    With grdCORES
        If (.Rows - 1) = 0 Then
           MsgBox "Não foi informado nenhuma Cor !!!", vbOKOnly + vbExclamation, "Aviso"
           Exit Function
        End If
        For I = 1 To (.Rows - 1)
            If Len(Trim(.Cell(flexcpText, I, conCOL_OrdVerniz_Lote))) = 0 Then
                MsgBox "Informe o Lote !!!", vbOKOnly + vbExclamation, "Aviso"
                Exit Function
            End If
            If Len(Trim(.Cell(flexcpText, I, conCOL_OrdVerniz_Producao))) = 0 Then
                MsgBox "Informe a Produção !!!", vbOKOnly + vbExclamation, "Aviso"
                Exit Function
            End If
            If Not IsDate(.Cell(flexcpText, I, conCOL_OrdVerniz_Data)) Then
                MsgBox "Informe a Data !!!", vbOKOnly + vbExclamation, "Aviso"
                Exit Function
            End If
            If Len(Trim(.Cell(flexcpText, I, conCOL_OrdVerniz_Maquina))) = 0 Then
                MsgBox "Informe a máquina !!!", vbOKOnly + vbExclamation, "Aviso"
                Exit Function
            End If
            If Len(Trim(.Cell(flexcpText, I, conCOL_OrdVerniz_Turno))) = 0 Then
                MsgBox "Informe o turno !!!", vbOKOnly + vbExclamation, "Aviso"
                Exit Function
            End If
        Next I
    End With
    
    ConsisteCampos = True
End Function

Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    
    Me.Caption = "Cadastra Dados da OP - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADDADOSOP

    Frame2.Enabled = False
    
    objCADDADOSOPC.CODIGO = iCodigo
    
    Call ConfGridVerniz
    Call ConfGridCores
    
    Call LimpaCamposLabel
    Call CarregaCampos
    
End Sub

Private Sub CarregaCampos()

    If objCADDADOSOPC.Carrega_Campos = True Then
        
        lblCODIGO.Caption = objCADDADOSOPC.CODIGO
        mskDATALOTE.Text = Format(objCADDADOSOPC.DATALOTE, "DD/MM/YYYY")
        txtCodOrdFat.Text = objCADDADOSOPC.CODOP
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       SGI_CODPROD " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_ORDEMPROD " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_CODIGO    = " & objCADDADOSOPC.CODOP
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC.EOF() Then lblCODPROD.Caption = Trim(BREC!SGI_CODPROD)
        BREC.Close
        
        Call PopGrdVerniz
        Call PopGrdCores
    
    End If

End Sub

 
Private Sub PopGrdVerniz()

    Dim I           As Long
    
    arrVERNIZ = objCADDADOSOPC.VERNIZ
    If IsArray(arrVERNIZ) Then
        With grdVERNIZ
            For I = 1 To UBound(arrVERNIZ)
                .Cell(flexcpText, I, conCOL_OrdVerniz_IdProduto) = arrVERNIZ(I, 1)
                .Cell(flexcpText, I, conCOL_OrdVerniz_Lote) = arrVERNIZ(I, 2)
                .Cell(flexcpText, I, conCOL_OrdVerniz_Producao) = arrVERNIZ(I, 3)
                .Cell(flexcpText, I, conCOL_OrdVerniz_Data) = arrVERNIZ(I, 4)
                .Cell(flexcpText, I, conCOL_OrdVerniz_Maquina) = arrVERNIZ(I, 5)
                .Cell(flexcpText, I, conCOL_OrdVerniz_Turno) = arrVERNIZ(I, 6)
                
                Call PegaDadosProd(Str(arrVERNIZ(I, 1)), I, grdVERNIZ, conCOL_OrdVerniz_Produto, conCOL_OrdVerniz_Descricao)
                Call PegaDadosMaq(Str(arrVERNIZ(I, 5)), I, grdVERNIZ, conCOL_OrdVerniz_DescMaq)
                Call PegaDadosTurn(Str(arrVERNIZ(I, 6)), I, grdVERNIZ, conCOL_OrdVerniz_DescTur)
            Next I
        End With
    End If

End Sub

Private Sub PegaDadosProd(strIDPRODUTO As String, lngLinha As Long, grdGENERICA As VSFlexGrid, lngCOL_CODPROD As Long, lngCOD_DESC As Long)

    If Len(Trim(strIDPRODUTO)) = 0 Then Exit Sub
    
    With grdGENERICA
    
        sSql = "Select " & vbCrLf
        sSql = sSql & "       SGI_CODIGO " & vbCrLf
        sSql = sSql & "      ,SGI_DESCRICAO " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_IDPRODUTO = " & Trim(strIDPRODUTO)
    
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC.EOF() Then
            .Cell(flexcpText, lngLinha, lngCOL_CODPROD) = BREC!SGI_CODIGO
            .Cell(flexcpText, lngLinha, lngCOD_DESC) = BREC!SGI_DESCRICAO
        End If
        BREC.Close
        
    End With

End Sub


Private Sub PegaDadosMaq(strMAQUINA As String, lngLinha As Long, grdGENERICO As VSFlexGrid, lngCOL_DESC As Long)

    If Len(Trim(strMAQUINA)) = 0 Then Exit Sub
    
    With grdGENERICO
    
        sSql = "Select " & vbCrLf
        sSql = sSql & "       SGI_DESCRI " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADMAQUINA " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_CODIGO = " & Trim(strMAQUINA)
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC.EOF() Then .Cell(flexcpText, lngLinha, lngCOL_DESC) = Trim(BREC!SGI_DESCRI)
        BREC.Close
        
    End With

End Sub

Private Sub PegaDadosTurn(strTURNO As String, lngLinha As Long, grdGENERICO As VSFlexGrid, lngCOL_DESC As Long)

    If Len(Trim(strTURNO)) = 0 Then Exit Sub
    
    With grdGENERICO
    
        sSql = "Select " & vbCrLf
        sSql = sSql & "       SGI_DESCRI " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADQTDETURN " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_CODIGO = " & Trim(strTURNO)
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC.EOF() Then .Cell(flexcpText, lngLinha, lngCOL_DESC) = Trim(BREC!SGI_DESCRI)
        BREC.Close
        
    End With

End Sub

Private Sub PopGrdCores()

    Dim I           As Long
    
    arrCORES = objCADDADOSOPC.CORES
    If IsArray(arrCORES) Then
        With grdCORES
            For I = 1 To UBound(arrCORES)
            
                .AddItem arrCORES(I, 1) & vbTab & _
                         I & "º Cor" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         arrCORES(I, 2) & vbTab & _
                         arrCORES(I, 3) & vbTab & _
                         arrCORES(I, 4) & vbTab & _
                         arrCORES(I, 5) & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         arrCORES(I, 6) & vbTab & _
                         "" & vbTab & _
                         ""
                          
                Call PegaDadosProd(Str(arrCORES(I, 1)), I, grdCORES, conCOL_OrdCores_Produto, conCOL_OrdCores_Descricao)
                Call PegaDadosMaq(Str(arrCORES(I, 5)), I, grdCORES, conCOL_OrdCores_DescMaq)
                Call PegaDadosTurn(Str(arrCORES(I, 6)), I, grdCORES, conCOL_OrdCores_DescTur)
            Next I
        End With
    End If

End Sub

Private Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    
    Me.Caption = "Cadastra Dados da OP - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADDADOSOP

    Frame2.Enabled = False
    Frame3.Enabled = True
    
    objCADDADOSOPC.CODIGO = iCodigo
    
    Call ConfGridCores
    Call ConfGridVerniz
    Call LimpaCamposLabel
    
    Call CarregaCampos
    
End Sub

