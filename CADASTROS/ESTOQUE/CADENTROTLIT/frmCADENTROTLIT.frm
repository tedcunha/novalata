VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCADENTROTLIT 
   Caption         =   "Envio de Material (Folhas Litografadas)"
   ClientHeight    =   8730
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18915
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   18915
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      Caption         =   "[ Sub-Total Folhas / Sub-Total Latas ]"
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
      Height          =   1695
      Left            =   0
      TabIndex        =   19
      Top             =   6960
      Width           =   12735
      Begin VSFlex8LCtl.VSFlexGrid grdSUBTOT 
         Height          =   1335
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   12495
         _cx             =   22040
         _cy             =   2355
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
   Begin VB.TextBox txtTotFardos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   14400
      TabIndex        =   18
      Text            =   "txtTotFardos"
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "[ Lançamentos ]"
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
      Height          =   4695
      Left            =   0
      TabIndex        =   13
      Top             =   2280
      Width           =   18855
      Begin VB.CommandButton Command3 
         Height          =   300
         Left            =   18360
         Picture         =   "frmCADENTROTLIT.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   600
         Width           =   300
      End
      Begin VB.CommandButton Command2 
         Height          =   300
         Left            =   18360
         Picture         =   "frmCADENTROTLIT.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   300
      End
      Begin VSFlex8LCtl.VSFlexGrid grdLANCTOS 
         Height          =   4335
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   18135
         _cx             =   31988
         _cy             =   7646
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
      Height          =   1335
      Left            =   0
      TabIndex        =   12
      Top             =   960
      Width           =   18855
      Begin VB.CommandButton Command4 
         Height          =   315
         Left            =   7245
         Picture         =   "frmCADENTROTLIT.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtCLIEDEST 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   23
         Text            =   "txtCLIEDES"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtCIDCLIE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "txtCIDCLIE"
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   7245
         Picture         =   "frmCADENTROTLIT.frx":0396
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   600
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskDTENTRADA 
         Height          =   285
         Left            =   3120
         TabIndex        =   3
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "txtCodigo"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblSTATUS 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblSTATUS"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6000
         TabIndex        =   27
         Top             =   240
         Width           =   8055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Status:"
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
         Index           =   5
         Left            =   4440
         TabIndex        =   26
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblDescClienteDest 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescClienteDest"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7605
         TabIndex        =   25
         Top             =   960
         Width           =   8055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente Destino"
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
         Index           =   3
         Left            =   4440
         TabIndex        =   22
         Top             =   960
         Width           =   1305
      End
      Begin VB.Label lblDescCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescCliente"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7605
         TabIndex        =   16
         Top             =   600
         Width           =   8055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente Origem"
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
         Left            =   4440
         TabIndex        =   4
         Top             =   600
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
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
         Left            =   2520
         TabIndex        =   2
         Top             =   240
         Width           =   480
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
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   18855
      Begin VB.CommandButton cmdImpressao 
         Caption         =   "Im&prime"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   18000
         Picture         =   "frmCADENTROTLIT.frx":0498
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Imprime Registro"
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
         Picture         =   "frmCADENTROTLIT.frx":059A
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Picture         =   "frmCADENTROTLIT.frx":069C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   855
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
         Left            =   960
         Picture         =   "frmCADENTROTLIT.frx":079E
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total de Fardos"
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
      Left            =   12840
      TabIndex        =   17
      Top             =   7080
      Width           =   1350
   End
End
Attribute VB_Name = "frmCADENTROTLIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public cTipOper         As String
Public iCodigo          As Long
Public strMODPAI        As String
Public FILIAL           As Integer
Public strAcesso        As String
Public strUsuario       As String
Public lngCodUsuario    As Long
Public strNOMETABELA    As String
Public strNOMEFILIAL    As String
Public intFILIALPED     As Integer

Dim objBLBFunc          As Object
Dim objCADENTROTLIT     As Object
Dim objPESQPADRAO       As Object
Dim objRel              As Object
Dim strCAPTION          As String
Dim arrENTROTLIT        As Variant

Const conCOL_ENTROTLIT_ID                      As Integer = 0
Const conCOL_ENTROTLIT_OP                      As Integer = 1
Const conCOL_ENTROTLIT_PEDIDO                  As Integer = 2
Const conCOL_ENTROTLIT_IDPRODUTO               As Integer = 3
Const conCOL_ENTROTLIT_CODIGO                  As Integer = 4
Const conCOL_ENTROTLIT_CODCAPAC                As Integer = 5
Const conCOL_ENTROTLIT_CAPAC                   As Integer = 6
Const conCOL_ENTROTLIT_LITSTEEL                As Integer = 7
Const conCOL_ENTROTLIT_PADRAO                  As Integer = 8

Const conCOL_ENTROTLIT_CODFOLHAUSADA           As Integer = 9
Const conCOL_ENTROTLIT_PESQFOLHAUSADA          As Integer = 10
Const conCOL_ENTROTLIT_FOLHAUSADA              As Integer = 11

Const conCOL_ENTROTLIT_ESPESS                  As Integer = 12
Const conCOL_ENTROTLIT_LARG                    As Integer = 13
Const conCOL_ENTROTLIT_COMP                    As Integer = 14
Const conCOL_ENTROTLIT_QTDECORP                As Integer = 15
Const conCOL_ENTROTLIT_PERDPRODC               As Integer = 16
Const conCOL_ENTROTLIT_QTDEFOLHAS              As Integer = 17
Const conCOL_ENTROTLIT_PESO                    As Integer = 18
Const conCOL_ENTROTLIT_UNID                    As Integer = 19
Const conCOL_ENTROTLIT_QTDELATAS               As Integer = 20
Const conCOL_ENTROTLIT_QTDEFARDOS              As Integer = 21
Const conCOL_ENTROTLIT_FILIALENTSAI            As Integer = 22
Const conCOL_ENTROTLIT_DESCENTSAI              As Integer = 23
Const conCOL_ENTROTLIT_FormatString            As String = "=ID|Cód.OP|Pedido|IDPRODUTO|Código|Cod.Capac|Capac.|Descrição Litografia|Padrão|Cod.Folha|...|Folha.Usada|Espessura|Largura|Comprimento|Qtde.Corpos|Perd.Proc|Qtde.Folhas|Peso|Unid.|Qtde.Latas|Qtde.Fardos|Filial Ent/Sai|Desc. Ent/Sai"
Const conColumnsIn_ENTROTLIT                   As Integer = 24

Const conCOL_SUBTOT_CODCAPC                    As Integer = 0
Const conCOL_SUBTOT_DESCAPA                    As Integer = 1
Const conCOL_SUBTOT_SUBTOTFL                   As Integer = 2
Const conCOL_SUBTOT_SUBTOTLT                   As Integer = 3
Const conCOL_SUBTOT_SUBTOTPS                   As Integer = 4
Const conCOL_SUBTOT_FormatString               As String = "=Cod.Capacidade|Descrição da Capacidade|Sub-Total Folhas|Sub-Total Latas|Sub-Total Peso"
Const conColumnsIn_SUBTOT                      As Integer = 5

Private Sub cmdAltera_Click()

    If objCADENTROTLIT.STATUS = "REC" Then
        MsgBox "ATENÇÃO" & vbCrLf & _
               "Não é permitido alterar !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    cTipOper = "A"
    If objBLBFunc.ChecaAcesso2(cTipOper, strAcesso) = False Then Exit Sub
    
    Call objBLBFunc.MudaBotoes(CmdSalva, cmdAltera, cTipOper)
    Call objBLBFunc.MudaCaption(Me, strCAPTION, cTipOper)
    Call DesabilitaCampos(Trim(cTipOper))

End Sub

Private Sub cmdImpressao_Click()
    If cTipOper = "C" Or cTipOper = "A" Then Call Imprime
End Sub

Private Sub CmdSalva_Click()

    Dim i                   As Long
    Dim sValor              As String
    
    If ValidaCampos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADENTROTLIT.CODIGO = objBLBFunc.Gera_Codigo(Trim(Me.Name), FILIAL, Linha)
    
    objCADENTROTLIT.STATUS = "'" & objCADENTROTLIT.STATUS & "'"
    If cTipOper = "I" Then objCADENTROTLIT.STATUS = "'ENV'"
    
    objCADENTROTLIT.DTENTRADA = "'" & Format(CDate(mskDTENTRADA.Text), "MM/DD/YYYY") & "'"
    objCADENTROTLIT.CODCLIE = Trim(txtCIDCLIE.Text)
    objCADENTROTLIT.CODCLIEDEST = Trim(txtCLIEDEST.Text)
    objCADENTROTLIT.ENTSAI = intFILIALPED
    
    arrENTROTLIT = Empty
    With grdLANCTOS
        ReDim arrENTROTLIT(1 To (.Rows - 1), 1 To 19) As String
        For i = 1 To (.Rows - 1)
            
            arrENTROTLIT(i, 1) = objBLBFunc.Gera_Codigo(Trim(Me.Name & "_CODITEN"), FILIAL, Linha)
            arrENTROTLIT(i, 2) = Trim(Str(.Cell(flexcpText, i, conCOL_ENTROTLIT_OP)))
            arrENTROTLIT(i, 3) = Trim(Str(.Cell(flexcpText, i, conCOL_ENTROTLIT_PEDIDO)))
            arrENTROTLIT(i, 4) = Trim(Str(.Cell(flexcpText, i, conCOL_ENTROTLIT_IDPRODUTO)))
            arrENTROTLIT(i, 5) = "'" & Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_CODIGO)) & "'"
            arrENTROTLIT(i, 6) = Trim(Str(.Cell(flexcpText, i, conCOL_ENTROTLIT_CODCAPAC)))
            arrENTROTLIT(i, 7) = Trim(Str(.Cell(flexcpText, i, conCOL_ENTROTLIT_PADRAO)))
            
            sValor = "Null"
            If Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_ESPESS))) > 0 Then
               sValor = Replace(.Cell(flexcpText, i, conCOL_ENTROTLIT_ESPESS), ".", "")
               sValor = Replace(sValor, ",", ".")
            End If
            arrENTROTLIT(i, 8) = sValor
            
            sValor = "Null"
            If Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_LARG))) > 0 Then
               sValor = Replace(.Cell(flexcpText, i, conCOL_ENTROTLIT_LARG), ".", "")
               sValor = Replace(sValor, ",", ".")
            End If
            arrENTROTLIT(i, 9) = sValor
            
            sValor = "Null"
            If Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_COMP))) > 0 Then
               sValor = Replace(.Cell(flexcpText, i, conCOL_ENTROTLIT_COMP), ".", "")
               sValor = Replace(sValor, ",", ".")
            End If
            arrENTROTLIT(i, 10) = sValor
            
            sValor = "Null"
            If Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_QTDECORP))) > 0 Then
               sValor = Replace(.Cell(flexcpText, i, conCOL_ENTROTLIT_QTDECORP), ".", "")
               sValor = Replace(sValor, ",", ".")
            End If
            arrENTROTLIT(i, 11) = sValor
            
            sValor = "Null"
            If Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_PERDPRODC))) > 0 Then
               sValor = Replace(.Cell(flexcpText, i, conCOL_ENTROTLIT_PERDPRODC), ".", "")
               sValor = Replace(sValor, ",", ".")
            End If
            arrENTROTLIT(i, 12) = sValor
            
            sValor = "Null"
            If Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_QTDEFOLHAS))) > 0 Then
               sValor = Replace(.Cell(flexcpText, i, conCOL_ENTROTLIT_QTDEFOLHAS), ".", "")
               sValor = Replace(sValor, ",", ".")
            End If
            arrENTROTLIT(i, 13) = sValor
            
            sValor = "Null"
            If Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_PESO))) > 0 Then
               sValor = Replace(.Cell(flexcpText, i, conCOL_ENTROTLIT_PESO), ".", "")
               sValor = Replace(sValor, ",", ".")
            End If
            arrENTROTLIT(i, 14) = sValor
            
            arrENTROTLIT(i, 15) = "Null"
            If Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_UNID))) > 0 Then arrENTROTLIT(i, 15) = "'" & Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_UNID)) & "'"
            
            sValor = "Null"
            If Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_QTDELATAS))) > 0 Then
               sValor = Replace(.Cell(flexcpText, i, conCOL_ENTROTLIT_QTDELATAS), ".", "")
               sValor = Replace(sValor, ",", ".")
            End If
            arrENTROTLIT(i, 16) = sValor
            
            sValor = "Null"
            If Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_QTDEFARDOS))) > 0 Then
               sValor = Replace(.Cell(flexcpText, i, conCOL_ENTROTLIT_QTDEFARDOS), ".", "")
               sValor = Replace(sValor, ",", ".")
            End If
            arrENTROTLIT(i, 17) = sValor
            
            arrENTROTLIT(i, 18) = .Cell(flexcpText, i, conCOL_ENTROTLIT_FILIALENTSAI)       '' Entrada e Saida
            arrENTROTLIT(i, 19) = .Cell(flexcpText, i, conCOL_ENTROTLIT_CODFOLHAUSADA)
            
        Next i
    End With
    objCADENTROTLIT.LANCTOS = arrENTROTLIT
    
    If objCADENTROTLIT.GRAVA(cTipOper) = False Then Exit Sub
          
    MsgBox "A Movimentação de Litografia foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
    If cTipOper = "I" Then Unload Me

End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

    ReDim arrCAMPOS(1 To 6, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_CODIGO " & vbCrLf
    sSql = sSql & "      ,SGI_CPFCNPJ" & vbCrLf
    sSql = sSql & "      ,SGI_RAZAOSOC" & vbCrLf
    sSql = sSql & "      ,SGI_NOMFANTA" & vbCrLf
    sSql = sSql & "      ,SGI_CIDNORM" & vbCrLf
    sSql = sSql & "      ,SGI_CODREF" & vbCrLf
    
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_VISTELENTEST = 1 " & vbCrLf
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_CPFCNPJ"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "CNPJ"
    arrCAMPOS(2, 4) = "1500"
    arrCAMPOS(2, 5) = "SGI_CPFCNPJ"
    
    arrCAMPOS(3, 1) = "SGI_RAZAOSOC"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "Razão Social"
    arrCAMPOS(3, 4) = "4500"
    arrCAMPOS(3, 5) = "SGI_RAZAOSOC"
    
    arrCAMPOS(4, 1) = "SGI_NOMFANTA"
    arrCAMPOS(4, 2) = "S"
    arrCAMPOS(4, 3) = "Nome Fantasia"
    arrCAMPOS(4, 4) = "2000"
    arrCAMPOS(4, 5) = "SGI_NOMFANTA"
    
    arrCAMPOS(5, 1) = "SGI_CIDNORM"
    arrCAMPOS(5, 2) = "S"
    arrCAMPOS(5, 3) = "Cidade"
    arrCAMPOS(5, 4) = "1500"
    arrCAMPOS(5, 5) = "SGI_CIDNORM"
    
    arrCAMPOS(6, 1) = "SGI_CODREF"
    arrCAMPOS(6, 2) = "S"
    arrCAMPOS(6, 3) = "Cód.Antigo"
    arrCAMPOS(6, 4) = "1500"
    arrCAMPOS(6, 5) = "SGI_CODREF"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Clientes")
    
    If Len(Trim(varRETORNO)) > 0 Then
        txtCIDCLIE.Text = varRETORNO
        Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", varRETORNO, lblDescCliente, "CLIE")
    End If
    txtCIDCLIE.SetFocus

End Sub

Private Sub Command2_Click()
    Call IncRegGrid
End Sub

Private Sub Command3_Click()
    If cTipOper = "C" Then Exit Sub
    If grdLANCTOS.Row = 0 Then Exit Sub
    Call objBLBFunc.ExclLinhaGrid(grdLANCTOS, grdLANCTOS.Row)
End Sub

Private Sub Command4_Click()

    ReDim arrCAMPOS(1 To 6, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_CODIGO " & vbCrLf
    sSql = sSql & "      ,SGI_CPFCNPJ" & vbCrLf
    sSql = sSql & "      ,SGI_RAZAOSOC" & vbCrLf
    sSql = sSql & "      ,SGI_NOMFANTA" & vbCrLf
    sSql = sSql & "      ,SGI_CIDNORM" & vbCrLf
    sSql = sSql & "      ,SGI_CODREF" & vbCrLf
    
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_VISTELENTEST = 1 " & vbCrLf
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_CPFCNPJ"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "CNPJ"
    arrCAMPOS(2, 4) = "1500"
    arrCAMPOS(2, 5) = "SGI_CPFCNPJ"
    
    arrCAMPOS(3, 1) = "SGI_RAZAOSOC"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "Razão Social"
    arrCAMPOS(3, 4) = "4500"
    arrCAMPOS(3, 5) = "SGI_RAZAOSOC"
    
    arrCAMPOS(4, 1) = "SGI_NOMFANTA"
    arrCAMPOS(4, 2) = "S"
    arrCAMPOS(4, 3) = "Nome Fantasia"
    arrCAMPOS(4, 4) = "2000"
    arrCAMPOS(4, 5) = "SGI_NOMFANTA"
    
    arrCAMPOS(5, 1) = "SGI_CIDNORM"
    arrCAMPOS(5, 2) = "S"
    arrCAMPOS(5, 3) = "Cidade"
    arrCAMPOS(5, 4) = "1500"
    arrCAMPOS(5, 5) = "SGI_CIDNORM"
    
    arrCAMPOS(6, 1) = "SGI_CODREF"
    arrCAMPOS(6, 2) = "S"
    arrCAMPOS(6, 3) = "Cód.Antigo"
    arrCAMPOS(6, 4) = "1500"
    arrCAMPOS(6, 5) = "SGI_CODREF"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Clientes")
    
    If Len(Trim(varRETORNO)) > 0 Then
        txtCLIEDEST.Text = varRETORNO
        Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", varRETORNO, lblDescClienteDest, "CLIE")
    End If
    txtCLIEDEST.SetFocus
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objCADENTROTLIT = CreateObject("CADENTROTLIT.clsCADENTROTLIT")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objRel = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    If adoBanco_Dados.State = 0 Then Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)
    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
   
    strCAPTION = "Envio de Material (Folhas Litografadas)"
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
   
    objCADENTROTLIT.FILIAL = FILIAL
   
    Call IniciaForm

End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objCADENTROTLIT = Nothing
    Set objPESQPADRAO = Nothing
    Set objRel = Nothing
End Sub

Private Sub IniciaForm()
    
    Call objBLBFunc.MudaBotoes(CmdSalva, cmdAltera, cTipOper)
    Call objBLBFunc.MudaCaption(Me, strCAPTION, cTipOper)
    Call objBLBFunc.LimpaCampos(Me)
    
    Call DesabilitaCampos(Trim(cTipOper))
    
    Call ConfGrd
    Call ConfGrdSubTot
    Call LimpaCamposLabel
    
    If cTipOper = "I" Then iCodigo = 0
    objCADENTROTLIT.CODIGO = iCodigo
    
    mskDTENTRADA.Text = Format(Now, "DD/MM/YYYY")
    
    Call CarregaCampos
    
End Sub

Private Sub DesabilitaCampos(strTipOper As String)
    If strTipOper = "I" Or strTipOper = "A" Then
        Frame2.Enabled = True
        cmdImpressao.Enabled = True
        If strTipOper = "I" Then cmdImpressao.Enabled = False
    ElseIf strTipOper = "C" Then
        Frame2.Enabled = False
        cmdImpressao.Enabled = True
    End If
End Sub


Private Sub ConfGrd()

    With grdLANCTOS

       .Cols = conColumnsIn_ENTROTLIT
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_ENTROTLIT_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeNone
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_ID) = ""
       .ColDataType(conCOL_ENTROTLIT_ID) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_OP) = ""
       .ColDataType(conCOL_ENTROTLIT_OP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_PEDIDO) = ""
       .ColDataType(conCOL_ENTROTLIT_PEDIDO) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_IDPRODUTO) = ""
       .ColDataType(conCOL_ENTROTLIT_IDPRODUTO) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_CODIGO) = ""
       .ColDataType(conCOL_ENTROTLIT_CODIGO) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_CODCAPAC) = ""
       .ColDataType(conCOL_ENTROTLIT_CODCAPAC) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_CAPAC) = ""
       .ColDataType(conCOL_ENTROTLIT_CAPAC) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_LITSTEEL) = ""
       .ColDataType(conCOL_ENTROTLIT_LITSTEEL) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_PADRAO) = ""
       .ColDataType(conCOL_ENTROTLIT_PADRAO) = flexDTString
       .ColComboList(conCOL_ENTROTLIT_PADRAO) = "|#1;Sim|#0;Não"
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_ESPESS) = ""
       .ColDataType(conCOL_ENTROTLIT_ESPESS) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_LARG) = ""
       .ColDataType(conCOL_ENTROTLIT_LARG) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_COMP) = ""
       .ColDataType(conCOL_ENTROTLIT_COMP) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_QTDECORP) = ""
       .ColDataType(conCOL_ENTROTLIT_QTDECORP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_PERDPRODC) = ""
       .ColDataType(conCOL_ENTROTLIT_PERDPRODC) = flexDTDouble
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_QTDEFOLHAS) = ""
       .ColDataType(conCOL_ENTROTLIT_QTDEFOLHAS) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_PESO) = ""
       .ColDataType(conCOL_ENTROTLIT_PESO) = flexDTDouble
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_QTDELATAS) = ""
       .ColDataType(conCOL_ENTROTLIT_QTDELATAS) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_QTDEFARDOS) = ""
       .ColDataType(conCOL_ENTROTLIT_QTDEFARDOS) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_FILIALENTSAI) = ""
       .ColDataType(conCOL_ENTROTLIT_FILIALENTSAI) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_DESCENTSAI) = ""
       .ColDataType(conCOL_ENTROTLIT_DESCENTSAI) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_DESCENTSAI) = ""
       .ColDataType(conCOL_ENTROTLIT_DESCENTSAI) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_CODFOLHAUSADA) = ""
       .ColDataType(conCOL_ENTROTLIT_CODFOLHAUSADA) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_FOLHAUSADA) = ""
       .ColDataType(conCOL_ENTROTLIT_FOLHAUSADA) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_PESQFOLHAUSADA) = ""
       .ColDataType(conCOL_ENTROTLIT_PESQFOLHAUSADA) = flexDTString
       .ColComboList(conCOL_ENTROTLIT_PESQFOLHAUSADA) = "..."
       
       .ColWidth(conCOL_ENTROTLIT_ID) = 0
       .ColWidth(conCOL_ENTROTLIT_OP) = 1100
       .ColWidth(conCOL_ENTROTLIT_PEDIDO) = 1100
       .ColWidth(conCOL_ENTROTLIT_IDPRODUTO) = 0
       .ColWidth(conCOL_ENTROTLIT_CODIGO) = 1500
       .ColWidth(conCOL_ENTROTLIT_CODCAPAC) = 0
       .ColWidth(conCOL_ENTROTLIT_CAPAC) = 0
       .ColWidth(conCOL_ENTROTLIT_LITSTEEL) = 5000
       .ColWidth(conCOL_ENTROTLIT_ESPESS) = 1000
       .ColWidth(conCOL_ENTROTLIT_LARG) = 1000
       .ColWidth(conCOL_ENTROTLIT_COMP) = 1000
       .ColWidth(conCOL_ENTROTLIT_QTDECORP) = 1000
       .ColWidth(conCOL_ENTROTLIT_PERDPRODC) = 0
       .ColWidth(conCOL_ENTROTLIT_QTDEFOLHAS) = 1000
       .ColWidth(conCOL_ENTROTLIT_PESO) = 1000
       .ColWidth(conCOL_ENTROTLIT_UNID) = 0
       .ColWidth(conCOL_ENTROTLIT_QTDELATAS) = 1000
       .ColWidth(conCOL_ENTROTLIT_QTDEFARDOS) = 1000
       .ColWidth(conCOL_ENTROTLIT_FILIALENTSAI) = 0
       .ColWidth(conCOL_ENTROTLIT_DESCENTSAI) = 0
       .ColWidth(conCOL_ENTROTLIT_PADRAO) = 0
       
       .ColWidth(conCOL_ENTROTLIT_CODFOLHAUSADA) = 1000
       .ColWidth(conCOL_ENTROTLIT_PESQFOLHAUSADA) = 300
       .ColWidth(conCOL_ENTROTLIT_FOLHAUSADA) = 1500
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
       
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

Private Sub IncRegGrid()
   
    If cTipOper = "C" Then Exit Sub
    
    If Len(Trim(txtCIDCLIE.Text)) = 0 Then
        MsgBox "Informe o Setor !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    If objBLBFunc.FcExisteLinhaVazia(grdLANCTOS, conCOL_ENTROTLIT_OP) = False Then Exit Sub
    If objBLBFunc.FcExisteLinhaVazia(grdLANCTOS, conCOL_ENTROTLIT_PESO) = False Then Exit Sub
    If objBLBFunc.FcExisteLinhaVazia(grdLANCTOS, conCOL_ENTROTLIT_QTDEFARDOS) = False Then Exit Sub
    
    With grdLANCTOS
        .AddItem "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 0 & vbTab & _
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
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & ""
                 
        Call PintaCelula((.Rows - 1))
    
    End With
   
End Sub


Private Sub grdLANCTOS_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With grdLANCTOS
        If (.Rows - 1) = 0 Then Exit Sub
        If Row = 0 Then Exit Sub
        Select Case Col
               Case conCOL_ENTROTLIT_OP
                    Call PopSubTotais
               Case conCOL_ENTROTLIT_ID
               Case conCOL_ENTROTLIT_QTDEFOLHAS
                    Call SubTotFLLT
               Case conCOL_ENTROTLIT_PESO
                    If Len(Trim(.Cell(flexcpText, Row, Col))) > 0 Then .Cell(flexcpText, Row, Col) = Format(.Cell(flexcpText, Row, Col), "#,####0.0000")
                    Call SubTotFLLT
               Case conCOL_ENTROTLIT_ESPESS, _
                    conCOL_ENTROTLIT_LARG, _
                    conCOL_ENTROTLIT_COMP
                    If Len(Trim(.Cell(flexcpText, Row, Col))) > 0 Then .Cell(flexcpText, Row, Col) = Format(.Cell(flexcpText, Row, Col), "#,##0.00")
               Case conCOL_ENTROTLIT_QTDEFARDOS
                    If TotalFardos > 0 Then txtTotFardos.Text = TotalFardos
               Case Else
                    .ComboList = ""
                    
        End Select
    End With
End Sub

Private Sub grdLANCTOS_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With grdLANCTOS
        Select Case Col
               Case conCOL_ENTROTLIT_ID, _
                    conCOL_ENTROTLIT_PEDIDO, _
                    conCOL_ENTROTLIT_IDPRODUTO, _
                    conCOL_ENTROTLIT_CODIGO, _
                    conCOL_ENTROTLIT_CODCAPAC, _
                    conCOL_ENTROTLIT_CAPAC, _
                    conCOL_ENTROTLIT_LITSTEEL, _
                    conCOL_ENTROTLIT_QTDECORP, _
                    conCOL_ENTROTLIT_PERDPRODC, _
                    conCOL_ENTROTLIT_QTDELATAS, _
                    conCOL_ENTROTLIT_FILIALENTSAI, _
                    conCOL_ENTROTLIT_DESCENTSAI, _
                    conCOL_ENTROTLIT_FOLHAUSADA
                    Cancel = True
               Case conCOL_ENTROTLIT_OP, _
                    conCOL_ENTROTLIT_PESO, _
                    conCOL_ENTROTLIT_UNID, _
                    conCOL_ENTROTLIT_QTDEFARDOS, _
                    conCOL_ENTROTLIT_QTDEFOLHAS, _
                    conCOL_ENTROTLIT_CODFOLHAUSADA, _
                    conCOL_ENTROTLIT_PESQFOLHAUSADA

                    If Col = conCOL_ENTROTLIT_PESQFOLHAUSADA Or Col = conCOL_ENTROTLIT_CODFOLHAUSADA Then
                        If cTipOper = "C" Or cTipOper = "R" Then
                           Cancel = True
                        Else
                           If Len(Trim(.Cell(flexcpText, Row, conCOL_ENTROTLIT_OP))) = 0 Then Cancel = True
                        End If
                    Else
                        If cTipOper = "C" Or _
                           cTipOper = "R" Then
                            Cancel = True
                        End If
                    End If
                    
               Case conCOL_ENTROTLIT_ESPESS, _
                    conCOL_ENTROTLIT_LARG, _
                    conCOL_ENTROTLIT_COMP
                    Cancel = True
''                    If cTipOper = "C" Then
''                       Cancel = True
''                    Else
''                        If .Cell(flexcpText, Row, conCOL_ENTROTLIT_PADRAO) = 1 Then Cancel = True
''                    End If
               Case Else
                   .ComboList = ""
               End Select
    End With
    Exit Sub
End Sub

Private Sub grdLANCTOS_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    Dim strDESCPROD As String
    Dim strINDICE   As String
    
    If cTipOper = "C" Then Exit Sub
    
    With grdLANCTOS
        If (.Rows - 1) <= 0 Then Exit Sub
        If (.Row) = 0 Then Exit Sub
    
        Select Case Col
            Case conCOL_ENTROTLIT_PESQFOLHAUSADA
                
                ReDim arrCAMPOS(1 To 2, 1 To 5) As String
                ReDim arrTABELA(1 To 1) As String
                
                sSql = ""
                
                sSql = "Select " & vbCrLf
                sSql = sSql & "        DIMC.SGI_CODIGO" & vbCrLf
                sSql = sSql & "      , DIMC.SGI_DESCORTE" & vbCrLf
                
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "        SGI_CADLINHAPRODUTO LINH" & vbCrLf
                sSql = sSql & "      , SGI_MEDCORTELINHA   MEDC" & vbCrLf
                sSql = sSql & "      , SGI_CADDIMCORTE     DIMC" & vbCrLf
                
                sSql = sSql & "  Where" & vbCrLf
                sSql = sSql & "        LINH.SGI_FILIAL     = " & FILIAL & vbCrLf
                sSql = sSql & "    And LINH.SGI_CODLIN     = " & Trim(.Cell(flexcpText, Row, conCOL_ENTROTLIT_CODCAPAC)) & vbCrLf
                
                sSql = sSql & "    And MEDC.SGI_FILIAL     = LINH.SGI_FILIAL" & vbCrLf
                sSql = sSql & "    And MEDC.SGI_CODIGO     = LINH.SGI_CODIGO" & vbCrLf
                
                sSql = sSql & "    And DIMC.SGI_FILIAL     = MEDC.SGI_FILIAL" & vbCrLf
                sSql = sSql & "    And DIMC.SGI_CODIGO     = MEDC.SGI_CODMEDCORT"
                
                arrTABELA(1) = sSql
                
                arrCAMPOS(1, 1) = "SGI_CODIGO"
                arrCAMPOS(1, 2) = "N"
                arrCAMPOS(1, 3) = "Código"
                arrCAMPOS(1, 4) = "2000"
                arrCAMPOS(1, 5) = "DIMC.SGI_CODIGO"
                
                arrCAMPOS(2, 1) = "SGI_DESCORTE"
                arrCAMPOS(2, 2) = "S"
                arrCAMPOS(2, 3) = "Descrição FF"
                arrCAMPOS(2, 4) = "5000"
                arrCAMPOS(2, 5) = "DIMC.SGI_DESCORTE"
                
                varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de FF")
                
                If Len(Trim(varRETORNO)) > 0 Then
                    .Cell(flexcpText, Row, conCOL_ENTROTLIT_CODFOLHAUSADA) = varRETORNO
                    Call PegaDadosFF(.Cell(flexcpText, Row, conCOL_ENTROTLIT_CODCAPAC), varRETORNO, Row)
                    
                    If Len(Trim(.Cell(flexcpText, Row, conCOL_ENTROTLIT_PESO))) = 0 Then Exit Sub
                    
                    .Cell(flexcpText, Row, conCOL_ENTROTLIT_QTDEFOLHAS) = QtdeFolhas(Row, CCur(.Cell(flexcpText, Row, conCOL_ENTROTLIT_PESO)))
                    .Cell(flexcpText, Row, conCOL_ENTROTLIT_QTDELATAS) = QtdeLatas(Row)
                    
                End If
                
        End Select
    End With
End Sub

Private Sub grdLANCTOS_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdLANCTOS
          Select Case Col
                    Case conCOL_ENTROTLIT_OP, _
                         conCOL_ENTROTLIT_QTDEFARDOS, _
                         conCOL_ENTROTLIT_CODFOLHAUSADA, _
                         conCOL_ENTROTLIT_QTDEFOLHAS
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
                    Case conCOL_ENTROTLIT_PESO
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 4, myvarAsDouble)
                    Case conCOL_ENTROTLIT_ESPESS, _
                         conCOL_ENTROTLIT_LARG, _
                         conCOL_ENTROTLIT_COMP
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 2, myvarAsDouble)
          End Select
     End With
End Sub

Private Sub grdLANCTOS_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

     Dim curSALDOPESO   As Currency
     Dim lngSALDOFARDO  As Long

     With grdLANCTOS
          Select Case Col
                 Case conCOL_ENTROTLIT_OP
                        If .EditText = Empty Then Exit Sub
                        
                        If Not IsNumeric(.EditText) Then
                            MsgBox "ATENÇÃO - Código da OP inválido !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                        
                        ''If Verif_se_JaExiste(.EditText) = True Then
                        ''    MsgBox "ATENÇÃO" & vbCrLf & _
                        ''           "Esta OP já foi lançada !!!", vbOKOnly + vbExclamation, "Aviso"
                        ''    .EditText = Empty
                        ''    Cancel = True
                        ''    Exit Sub
                        ''End If
                        
                        If PegaDadosOP(.EditText, Row) = False Then
                            .EditText = Empty
                            MsgBox "ATENÇÃO - OP não existe !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                 Case conCOL_ENTROTLIT_QTDEFOLHAS
                        If .EditText = Empty Then Exit Sub
                        
                        If Not IsNumeric(.EditText) Then
                            MsgBox "ATENÇÃO - O Peso está Inválido !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                        
                        .Cell(flexcpText, Row, conCOL_ENTROTLIT_PESO) = Format(CalculaPeso(Row, CLng(.EditText)), "#,####0.0000")
                        ''.Cell(flexcpText, Row, conCOL_ENTROTLIT_QTDEFOLHAS) = QtdeFolhas(Row, CCur(.Cell(flexcpText, Row, conCOL_ENTROTLIT_PESO)))
                        .Cell(flexcpText, Row, conCOL_ENTROTLIT_QTDEFOLHAS) = .EditText
                        .Cell(flexcpText, Row, conCOL_ENTROTLIT_QTDELATAS) = QtdeLatas(Row)
                        
                 Case conCOL_ENTROTLIT_PESO
                        If .EditText = Empty Then Exit Sub
                        
                        If Not IsNumeric(.EditText) Then
                            MsgBox "ATENÇÃO - O Peso está Inválido !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
           
                        .Cell(flexcpText, Row, conCOL_ENTROTLIT_QTDEFOLHAS) = QtdeFolhas(Row, CCur(.EditText))
                        .Cell(flexcpText, Row, conCOL_ENTROTLIT_QTDELATAS) = QtdeLatas(Row)
                 
                 Case conCOL_ENTROTLIT_ESPESS, _
                      conCOL_ENTROTLIT_LARG, _
                      conCOL_ENTROTLIT_COMP
                        If .EditText = Empty Then Exit Sub
                        
                        If Not IsNumeric(.EditText) Then
                            MsgBox "ATENÇÃO - Valor Inválido !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                        
                 Case conCOL_ENTROTLIT_QTDEFARDOS
                        If .EditText = Empty Then
                            Exit Sub
                        End If
                        
                        If Not IsNumeric(.EditText) Then
                            MsgBox "ATENÇÃO - Qtde. Fardo inválido !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                 
                 Case conCOL_ENTROTLIT_CODFOLHAUSADA
                        If .EditText = Empty Then Exit Sub
                        
                        If Not IsNumeric(.EditText) Then
                            MsgBox "ATENÇÃO - Código da FF inválida !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                        
                        Cancel = PegaDadosFF(.Cell(flexcpText, Row, conCOL_ENTROTLIT_CODCAPAC), .EditText, Row)
                        If Cancel = True Then Exit Sub
                        
                        If Len(Trim(.Cell(flexcpText, Row, conCOL_ENTROTLIT_PESO))) = 0 Then
                            Cancel = True
                            Exit Sub
                        End If
                        
                        .Cell(flexcpText, Row, conCOL_ENTROTLIT_QTDEFOLHAS) = QtdeFolhas(Row, CCur(.Cell(flexcpText, Row, conCOL_ENTROTLIT_PESO)))
                        .Cell(flexcpText, Row, conCOL_ENTROTLIT_QTDELATAS) = QtdeLatas(Row)
                        
          End Select
     End With

End Sub

Private Function PegaDadosOP(strCODOP As String, lngROW As Long) As Boolean

    PegaDadosOP = False

    If Len(Trim(strCODOP)) = 0 Then Exit Function
    
    Dim strEMPRESA      As String
    Dim boolNOVALATA    As Boolean
    Dim boolSTEEL       As Boolean
    Dim intFILIALOP     As Integer
    
    strEMPRESA = ""
    
    
    '' Novalata
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "      *" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_ORDEMPROD" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & strCODOP & vbCrLf
    sSql = sSql & "   And (SGI_STATUS = 0 or SGI_STATUS = 1)"
    
    boolNOVALATA = False
    BREC11.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC11.EOF() Then
       boolNOVALATA = True
       intFILIALOP = 1
    End If
    BREC11.Close
    
    If boolNOVALATA = False Then
       
        '' Steel
        sSql = ""
        
        sSql = "Select" & vbCrLf
        sSql = sSql & "      *" & vbCrLf
        
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "      SGI_ORDEMPROD_STEEL" & vbCrLf
        
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_CODIGO = " & strCODOP & vbCrLf
        sSql = sSql & "   And (SGI_STATUS = 0 or SGI_STATUS = 1)"
        
        boolSTEEL = False
        BREC11.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC11.EOF() Then
            boolSTEEL = True
            strEMPRESA = "_STEEL"
            intFILIALOP = 0
        End If
        BREC11.Close
        
    End If
    
    If (boolNOVALATA = False And boolSTEEL = False) Then
       MsgBox "Atenção esta OP não Existe ou já esta fechada !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    
   
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       OP.*" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_CODLINPROD" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_DESCRICAO          As SGI_DESCPROD" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_QTDCORPSPADRAOSN" & vbCrLf
    sSql = sSql & "      ,LINHA.SGI_DESCRI            As SGI_CAPACIDADE" & vbCrLf
    sSql = sSql & "      ,LINHA.SGI_FILIALPED" & vbCrLf
    
    ''sSql = sSql & "      ,MEDCORTE.SGI_CODMEDCORT" & vbCrLf
    ''sSql = sSql & "      ,MEDCORTE.SGI_EXPESS" & vbCrLf
    ''sSql = sSql & "      ,MEDCORTE.SGI_LARGUR" & vbCrLf
    ''sSql = sSql & "      ,MEDCORTE.SGI_COMPRI" & vbCrLf
    ''sSql = sSql & "      ,MEDCORTE.SGI_QTDECORPOS As SGI_QTDECORPOSPADRAO" & vbCrLf
    ''sSql = sSql & "      ,MEDCORTE.SGI_PERDPROC" & vbCrLf
    ''sSql = sSql & "      ,MEDCORTE.SGI_PADRAO" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD" & strEMPRESA & " OP" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO      PROD" & vbCrLf
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO LINHA" & vbCrLf
    ''sSql = sSql & "      ,SGI_MEDCORTELINHA   MEDCORTE" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       OP.SGI_FILIAL       = " & FILIAL & vbCrLf
    sSql = sSql & "   And OP.SGI_CODIGO       = " & strCODOP & vbCrLf
    sSql = sSql & "   And PROD.SGI_FILIAL     = OP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PROD.SGI_IDPRODUTO  = OP.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "   And LINHA.SGI_FILIAL    = PROD.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And LINHA.SGI_CODLIN    = PROD.SGI_CODLINPROD" & vbCrLf
    ''sSql = sSql & "   And MEDCORTE.SGI_FILIAL = LINHA.SGI_FILIAL" & vbCrLf
    ''sSql = sSql & "   And MEDCORTE.SGI_CODIGO = LINHA.SGI_CODIGO" & vbCrLf
    ''sSql = sSql & "   And MEDCORTE.SGI_PADRAO = 1"
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC10.EOF() Then
        
        ''If optEMPRESA(BREC10!SGI_FILIALPED).Value = False Then
        ''   MsgBox "Atenção esta OP não pertence a " & IIf(BREC10!SGI_FILIALPED = 0, "NOVALATA", "STEEL") & " !!!", vbOKOnly + vbExclamation, "Aviso"
        ''Else
            PegaDadosOP = True
            With grdLANCTOS
            
                .Cell(flexcpText, lngROW, conCOL_ENTROTLIT_PEDIDO) = BREC10!SGI_CODPED
                .Cell(flexcpText, lngROW, conCOL_ENTROTLIT_IDPRODUTO) = BREC10!SGI_IDPRODUTO
                .Cell(flexcpText, lngROW, conCOL_ENTROTLIT_CODIGO) = BREC10!SGI_CODPROD
                .Cell(flexcpText, lngROW, conCOL_ENTROTLIT_CODCAPAC) = BREC10!SGI_CODLINPROD
                .Cell(flexcpText, lngROW, conCOL_ENTROTLIT_CAPAC) = BREC10!SGI_CAPACIDADE
                .Cell(flexcpText, lngROW, conCOL_ENTROTLIT_LITSTEEL) = BREC10!SGI_DESCPROD
                .Cell(flexcpText, lngROW, conCOL_ENTROTLIT_PADRAO) = 0
                
                .ColComboList(conCOL_ENTROTLIT_FOLHAUSADA) = objCADENTROTLIT.PreenchComboFolhaFF(Str(BREC10!SGI_CODLINPROD))
                
                ''    .Cell(flexcpText, lngROW, conCOL_ENTROTLIT_FOLHAUSADA) = BREC10!SGI_CODMEDCORT
                ''    .Cell(flexcpText, lngROW, conCOL_ENTROTLIT_ESPESS) = Format(BREC10!SGI_EXPESS, "#,####0.0000")
                ''    .Cell(flexcpText, lngROW, conCOL_ENTROTLIT_LARG) = Format(BREC10!SGI_LARGUR, "#,####0.0000")
                ''    .Cell(flexcpText, lngROW, conCOL_ENTROTLIT_COMP) = Format(BREC10!SGI_COMPRI, "#,####0.0000")
                ''    .Cell(flexcpText, lngROW, conCOL_ENTROTLIT_QTDECORP) = BREC10!SGI_QTDECORPOSPADRAO
                ''    .Cell(flexcpText, lngROW, conCOL_ENTROTLIT_PERDPRODC) = Format(BREC10!SGI_PERDPROC, "#,####0.0000")
                
                .Cell(flexcpText, lngROW, conCOL_ENTROTLIT_FILIALENTSAI) = intFILIALOP
               
                If intFILIALOP = 1 Then .Cell(flexcpText, lngROW, conCOL_ENTROTLIT_DESCENTSAI) = "NOVALATA"
                If intFILIALOP = 0 Then .Cell(flexcpText, lngROW, conCOL_ENTROTLIT_DESCENTSAI) = "STEEL"
              
            End With
        ''End If
    End If
    BREC10.Close
    
End Function

Private Sub LimpaCamposLabel()
    lblDescCliente.Caption = ""
    lblDescClienteDest.Caption = ""
    lblSTATUS.Caption = ""
End Sub


Private Sub mskDTENTRADA_GotFocus()
    objBLBFunc.SelecionaCampos mskDTENTRADA.Name, Me
End Sub


Private Sub txtCIDCLIE_GotFocus()
    objBLBFunc.SelecionaCampos txtCIDCLIE.Name, Me
End Sub

Private Sub txtCIDCLIE_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCIDCLIE.Text
End Sub

Private Sub txtCIDCLIE_Validate(Cancel As Boolean)

    If Len(Trim(txtCIDCLIE.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCIDCLIE.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCIDCLIE.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", txtCIDCLIE.Text, lblDescCliente, "CLIE")
    If Len(Trim(lblDescCliente.Caption)) = 0 Then
       txtCIDCLIE.Text = ""
       Cancel = True
       Exit Sub
    End If
    
End Sub

Private Sub PegaDescTabelas(StrCampoPesq As String, StrCampoRetorno As String, strTabela As String, strCodigo As String, lblLabel As Label, strTIPO As String)

    lblLabel.Caption = ""
    
    If Len(Trim(strCodigo)) = 0 Then Exit Sub
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoRetorno) & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & Trim(strTabela) & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoPesq) & " = " & Trim(strCodigo)
    
    If strTIPO = "CLIE" Then sSql = sSql & "   And SGI_VISTELENTEST = 1 " & vbCrLf
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then
       lblLabel.Caption = Trim(BREC10(Trim(StrCampoRetorno)))
    Else
       MsgBox "Registro Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC10.Close
    
End Sub

Private Function QtdeFolhas(lngROW As Long, currPeso As Currency) As Long

    QtdeFolhas = 0

    Dim curPESO        As Currency
    Dim curExpessura   As Currency
    Dim curLargura     As Currency
    Dim curComprimento As Currency

    With grdLANCTOS
        curPESO = currPeso
        curExpessura = 0
        curLargura = 0
        curComprimento = 0
    
        If Len(Trim(.Cell(flexcpText, lngROW, conCOL_ENTROTLIT_ESPESS))) > 0 Then curExpessura = CCur(.Cell(flexcpText, lngROW, conCOL_ENTROTLIT_ESPESS))
        If Len(Trim(.Cell(flexcpText, lngROW, conCOL_ENTROTLIT_LARG))) > 0 Then curLargura = CCur(.Cell(flexcpText, lngROW, conCOL_ENTROTLIT_LARG))
        If Len(Trim(.Cell(flexcpText, lngROW, conCOL_ENTROTLIT_COMP))) > 0 Then curComprimento = CCur(.Cell(flexcpText, lngROW, conCOL_ENTROTLIT_COMP))
    
    End With

    If curExpessura > 0 And curLargura > 0 And curComprimento > 0 Then
        QtdeFolhas = (curPESO / (curExpessura * (curLargura / 1000) * (curComprimento / 1000) * 7.85))
    End If

End Function

Private Function QtdeLatas(lngROW As Long) As Long

    QtdeLatas = 0
    
    Dim lngQTDEFOLHAS As Long
    Dim lngQTDECORPOS As Long
    
    lngQTDEFOLHAS = 0
    lngQTDECORPOS = 0
    
    With grdLANCTOS
        If Len(Trim(.Cell(flexcpText, lngROW, conCOL_ENTROTLIT_QTDEFOLHAS))) > 0 Then lngQTDEFOLHAS = CLng(.Cell(flexcpText, lngROW, conCOL_ENTROTLIT_QTDEFOLHAS))
        If Len(Trim(.Cell(flexcpText, lngROW, conCOL_ENTROTLIT_QTDECORP))) > 0 Then lngQTDECORPOS = CLng(.Cell(flexcpText, lngROW, conCOL_ENTROTLIT_QTDECORP))
    End With
    
    QtdeLatas = (lngQTDEFOLHAS * lngQTDECORPOS)
    
End Function


Public Function TotalFardos() As Long

    Dim lngTotFardos    As Long
    Dim lngQtdeFardos   As Long
    Dim i               As Integer
    
    With grdLANCTOS
        For i = 1 To (.Rows - 1)
            lngQtdeFardos = 0
            If Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_QTDEFARDOS))) > 0 Then lngQtdeFardos = CLng(.Cell(flexcpText, i, conCOL_ENTROTLIT_QTDEFARDOS))
            lngTotFardos = (lngTotFardos + lngQtdeFardos)
        Next i
    End With

    TotalFardos = lngTotFardos

End Function

Private Sub ConfGrdSubTot()

    With grdSUBTOT

       .Cols = conColumnsIn_SUBTOT
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SUBTOT_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeNone
       
       .Cell(flexcpData, 0, conCOL_SUBTOT_CODCAPC) = ""
       .ColDataType(conCOL_SUBTOT_CODCAPC) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SUBTOT_DESCAPA) = ""
       .ColDataType(conCOL_SUBTOT_DESCAPA) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SUBTOT_SUBTOTFL) = ""
       .ColDataType(conCOL_SUBTOT_SUBTOTFL) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SUBTOT_SUBTOTLT) = ""
       .ColDataType(conCOL_SUBTOT_SUBTOTLT) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SUBTOT_SUBTOTPS) = ""
       .ColDataType(conCOL_SUBTOT_SUBTOTPS) = flexDTLong
       
       .ColWidth(conCOL_SUBTOT_CODCAPC) = 0
       .ColWidth(conCOL_SUBTOT_DESCAPA) = 4000
       .ColWidth(conCOL_SUBTOT_SUBTOTFL) = 1300
       .ColWidth(conCOL_SUBTOT_SUBTOTLT) = 1200
       .ColWidth(conCOL_SUBTOT_SUBTOTPS) = 1200
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack

    End With
    
End Sub


Private Sub PopSubTotais()

    Call ConfGrdSubTot
    
    Dim i           As Integer
    Dim lngINDICE   As Integer
    
    With grdLANCTOS
         
         For i = 1 To (.Rows - 1)
         
             If Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_CODCAPAC))) > 0 Then
         
                 lngINDICE = grdSUBTOT.FindRow(CLng(.Cell(flexcpText, i, conCOL_ENTROTLIT_CODCAPAC)), , conCOL_SUBTOT_CODCAPC)
        
                 If lngINDICE = -1 Then
                 
                    grdSUBTOT.AddItem .Cell(flexcpText, i, conCOL_ENTROTLIT_CODCAPAC) & vbTab & _
                                      .Cell(flexcpText, i, conCOL_ENTROTLIT_CAPAC) & vbTab & _
                                      "" & vbTab & _
                                      "" & vbTab & _
                                      ""
                 
                 End If
             
             End If
    
         Next i
    End With

End Sub

Private Sub SubTotFLLT()

    Dim i           As Integer
    Dim j           As Integer
    Dim lngROW      As Long
    Dim lngQTDFL    As Long
    Dim lngQTDLT    As Long
    Dim lngTOTFL    As Long
    Dim lngTOTLT    As Long
    Dim lngQTDPS    As Long
    Dim lngTOTPS    As Long
    
    With grdSUBTOT
        
        For i = 1 To (.Rows - 1)
            
            lngTOTFL = 0
            lngTOTLT = 0
            lngTOTPS = 0
            For j = 1 To (grdLANCTOS.Rows - 1)
            
                lngQTDFL = 0
                lngQTDLT = 0
                lngQTDPS = 0
                If grdLANCTOS.Cell(flexcpText, j, conCOL_ENTROTLIT_CODCAPAC) = .Cell(flexcpText, i, conCOL_SUBTOT_CODCAPC) Then
                    If Len(Trim(grdLANCTOS.Cell(flexcpText, j, conCOL_ENTROTLIT_QTDEFOLHAS))) > 0 Then lngQTDFL = CLng(grdLANCTOS.Cell(flexcpText, j, conCOL_ENTROTLIT_QTDEFOLHAS))
                    If Len(Trim(grdLANCTOS.Cell(flexcpText, j, conCOL_ENTROTLIT_QTDELATAS))) > 0 Then lngQTDLT = CLng(grdLANCTOS.Cell(flexcpText, j, conCOL_ENTROTLIT_QTDELATAS))
                    If Len(Trim(grdLANCTOS.Cell(flexcpText, j, conCOL_ENTROTLIT_PESO))) > 0 Then lngQTDPS = CLng(grdLANCTOS.Cell(flexcpText, j, conCOL_ENTROTLIT_PESO))
                End If
                
                lngTOTFL = (lngTOTFL + lngQTDFL)
                lngTOTLT = (lngTOTLT + lngQTDLT)
                lngTOTPS = (lngTOTPS + lngQTDPS)
            
            Next j
    
            If lngTOTFL > 0 Then .Cell(flexcpText, i, conCOL_SUBTOT_SUBTOTFL) = lngTOTFL
            If lngTOTLT > 0 Then .Cell(flexcpText, i, conCOL_SUBTOT_SUBTOTLT) = lngTOTLT
            If lngTOTPS > 0 Then .Cell(flexcpText, i, conCOL_SUBTOT_SUBTOTPS) = lngTOTPS
        
        Next i
        
    End With

End Sub

Private Function ValidaCampos() As Boolean

        ValidaCampos = False
     
        Dim i As Integer
                
        If Len(Trim(txtCIDCLIE.Text)) = 0 Then
            MsgBox "ATENÇÃO" & vbCrLf & _
                   "Campo de Origem Cliente não pode ser vázio !!!", vbOKOnly + vbExclamation, "Acviso"
                   txtCIDCLIE.SetFocus
                   Exit Function
        End If
     
        If Len(Trim(Replace(mskDTENTRADA.Text, "/", ""))) = 0 Then
            MsgBox "ATENÇÃO" & vbCrLf & _
                   "Campo Data de Entrada não pode ser vázio !!!", vbOKOnly + vbExclamation, "Acviso"
                   mskDTENTRADA.SetFocus
                   Exit Function
        End If
     
        If Not IsDate(mskDTENTRADA.Text) Then
            MsgBox "ATENÇÃO" & vbCrLf & _
                   "Campo Data inválido !!!", vbOKOnly + vbExclamation, "Acviso"
                   mskDTENTRADA.SetFocus
                   Exit Function
        End If
        
        If Len(Trim(txtCLIEDEST.Text)) = 0 Then
            MsgBox "ATENÇÃO" & vbCrLf & _
                   "Campo de Destino Cliente não pode ser vázio !!!", vbOKOnly + vbExclamation, "Acviso"
                   txtCLIEDEST.SetFocus
                   Exit Function
        End If
        
        ''If optEMPRESA(0).Value = True And optEMPRESA(1).Value = True Then
        ''    MsgBox "ATENÇÃO" & vbCrLf & _
        ''           "Escolha uma empresa !!!", vbOKOnly + vbExclamation, "Acviso"
        ''           Exit Function
        ''End If
     
        If (grdLANCTOS.Rows - 1) = 0 Then
            MsgBox "ATENÇÃO" & vbCrLf & _
                   "Não foram informados movimentos !!!", vbOKOnly + vbExclamation, "Acviso"
                   Exit Function
        End If
        
        With grdLANCTOS
            For i = 1 To (.Rows - 1)
                If Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_QTDEFOLHAS))) = 0 Then
                    MsgBox "ATENÇÃO" & vbCrLf & _
                           "O campo Qtde. de Folhas deve ser informado !!!", vbOKOnly + vbExclamation, "Acviso"
                           Exit Function
                End If
                If Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_PESO))) = 0 Then
                    MsgBox "ATENÇÃO" & vbCrLf & _
                           "O campo Peso deve ser informado !!!", vbOKOnly + vbExclamation, "Acviso"
                           Exit Function
                End If
                If Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_QTDEFARDOS))) = 0 Then
                    MsgBox "ATENÇÃO" & vbCrLf & _
                           "O campo Fardos deve ser informado !!!", vbOKOnly + vbExclamation, "Acviso"
                           Exit Function
                End If
                If Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_CODFOLHAUSADA))) = 0 Then
                    MsgBox "ATENÇÃO" & vbCrLf & _
                           "O campo Folha usada deve ser informado !!!", vbOKOnly + vbExclamation, "Acviso"
                           Exit Function
                End If
            Next i
        End With
        
        ValidaCampos = True
     
End Function


Private Sub CarregaCampos()

On Error GoTo Err_CarregaCampos
    
    If objCADENTROTLIT.Carrega_campos = True Then
        
        txtCodigo.Text = objCADENTROTLIT.CODIGO
        mskDTENTRADA.Text = objCADENTROTLIT.DTENTRADA
        txtCIDCLIE.Text = objCADENTROTLIT.CODCLIE
        txtCLIEDEST.Text = objCADENTROTLIT.CODCLIEDEST
        
        
        Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", txtCIDCLIE.Text, lblDescCliente, "CLIE")
        Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", txtCLIEDEST, lblDescClienteDest, "CLIE")
    
        Call PopGrdLancto
    
        Call PopSubTotais
        Call SubTotFLLT
        If TotalFardos > 0 Then txtTotFardos.Text = TotalFardos
        
        If objCADENTROTLIT.STATUS = "ENV" Then lblSTATUS.Caption = "ENVIADO AGUARDANDO CONFIRMAÇÂO"
        If objCADENTROTLIT.STATUS = "REC" Then lblSTATUS.Caption = "RECEBIDO JÁ CONFIRMADO"
            
    
    End If
    
    Exit Sub

Err_CarregaCampos:
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : CarregaCampos", Me.Name, "CarregaCampos")
    
End Sub

Private Sub PopGrdLancto()

    Dim i As Integer
    
    arrENTROTLIT = objCADENTROTLIT.LANCTOS
    If IsArray(arrENTROTLIT) Then
        With grdLANCTOS
            For i = 1 To UBound(arrENTROTLIT)
                .AddItem arrENTROTLIT(i, 1) & vbTab & _
                         arrENTROTLIT(i, 2) & vbTab & _
                         arrENTROTLIT(i, 3) & vbTab & _
                         arrENTROTLIT(i, 4) & vbTab & _
                         arrENTROTLIT(i, 5) & vbTab & _
                         arrENTROTLIT(i, 6) & vbTab & _
                         PegaDescCapac(Str(arrENTROTLIT(i, 6))) & vbTab & _
                         PegaDescProd(Str(arrENTROTLIT(i, 4))) & vbTab & _
                         arrENTROTLIT(i, 7) & vbTab & _
                         arrENTROTLIT(i, 19) & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         arrENTROTLIT(i, 8) & vbTab & _
                         arrENTROTLIT(i, 9) & vbTab & _
                         arrENTROTLIT(i, 10) & vbTab & _
                         arrENTROTLIT(i, 11) & vbTab & _
                         arrENTROTLIT(i, 12) & vbTab & _
                         arrENTROTLIT(i, 13) & vbTab & _
                         arrENTROTLIT(i, 14) & vbTab & _
                         arrENTROTLIT(i, 15) & vbTab & _
                         arrENTROTLIT(i, 16) & vbTab & _
                         arrENTROTLIT(i, 17) & vbTab & _
                         arrENTROTLIT(i, 18) & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & ""
                
                         If arrENTROTLIT(i, 18) = 1 Then .Cell(flexcpText, (.Rows - 1), conCOL_ENTROTLIT_DESCENTSAI) = "NOVALATA"
                         If arrENTROTLIT(i, 18) = 0 Then .Cell(flexcpText, (.Rows - 1), conCOL_ENTROTLIT_DESCENTSAI) = "STEEL"
                         
                         
                Call PegaDadosFF(Str(arrENTROTLIT(i, 6)), Str(arrENTROTLIT(i, 19)), (.Rows - 1))
                Call PintaCelula(i)
            
            Next i
        End With
    End If

End Sub

Private Sub PintaCelula(intROW As Integer)
    With grdLANCTOS
        .Cell(flexcpBackColor, intROW, conCOL_ENTROTLIT_OP) = &H80C0FF
        .Cell(flexcpBackColor, intROW, conCOL_ENTROTLIT_QTDEFOLHAS) = &H80C0FF
        .Cell(flexcpBackColor, intROW, conCOL_ENTROTLIT_PESO) = &H80C0FF
        .Cell(flexcpBackColor, intROW, conCOL_ENTROTLIT_QTDEFARDOS) = &H80C0FF
    End With
End Sub

Private Function PegaDescCapac(strCODCAPAC As String) As String

    PegaDescCapac = ""
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       SGI_DESCRI" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODLIN = " & strCODCAPAC
    
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC2.EOF() Then PegaDescCapac = BREC2!SGI_DESCRI
    BREC2.Close

End Function


Private Function PegaDescProd(strIDPROD As String) As String

    PegaDescProd = ""
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       SGI_DESCRICAO" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_IDPRODUTO = " & strIDPROD
    
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC2.EOF() Then PegaDescProd = BREC2!SGI_DESCRICAO
    BREC2.Close

End Function

Private Sub Imprime()

    Dim boolTemRegistro As Boolean
    Dim strNomRel       As String
    Dim strCABEC2       As String
    
    strNomRel = "RPTCADENTROTLIT.RPT"
    
    boolTemRegistro = False
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       SGI_CADENTROTLIT_IT.SGI_CODIGO" & vbCrLf
    sSql = sSql & "      ,SGI_CADENTROTLIT_IT.SGI_CODOP" & vbCrLf
    sSql = sSql & "      ,SGI_CADENTROTLIT_IT.SGI_CODPED" & vbCrLf
    sSql = sSql & "      ,SGI_CADENTROTLIT_IT.SGI_PRODUTO" & vbCrLf
    sSql = sSql & "      ,SGI_CADENTROTLIT_IT.SGI_CODCAPAC" & vbCrLf
    sSql = sSql & "      ,SGI_CADENTROTLIT_IT.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "      ,SGI_CADENTROTLIT_IT.SGI_EXPESS" & vbCrLf
    sSql = sSql & "      ,SGI_CADENTROTLIT_IT.SGI_LARGUR" & vbCrLf
    sSql = sSql & "      ,SGI_CADENTROTLIT_IT.SGI_COMPRI" & vbCrLf
    sSql = sSql & "      ,SGI_CADENTROTLIT_IT.SGI_QTDECORP" & vbCrLf
    sSql = sSql & "      ,SGI_CADENTROTLIT_IT.SGI_QTDEFOLHAS" & vbCrLf
    sSql = sSql & "      ,SGI_CADENTROTLIT_IT.SGI_PESO" & vbCrLf
    sSql = sSql & "      ,SGI_CADENTROTLIT_IT.SGI_QTDELATAS" & vbCrLf
    sSql = sSql & "      ,SGI_CADENTROTLIT_IT.SGI_QTDEFARDOS" & vbCrLf
    
    sSql = sSql & "      ,SGI_CADENTROTLIT.SGI_DTENTRADA" & vbCrLf
    sSql = sSql & "      ,SGI_CADENTROTLIT.SGI_EMPRESA" & vbCrLf
    sSql = sSql & "      ,SGI_CADENTROTLIT.SGI_CODCLIE" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE.SGI_RAZAOSOC" & vbCrLf
    
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO.SGI_DESCRI" & vbCrLf
    
    sSql = sSql & "      ,SGI_CADPRODUTO.SGI_DESCRICAO" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE" & vbCrLf
    sSql = sSql & "      ,SGI_CADENTROTLIT" & vbCrLf
    sSql = sSql & "      ,SGI_CADENTROTLIT_IT" & vbCrLf
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_CADENTROTLIT_IT.SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CADENTROTLIT_IT.SGI_CODIGO    = " & objCADENTROTLIT.CODIGO & vbCrLf
    
    sSql = sSql & "   And SGI_CADENTROTLIT_IT.SGI_FILIAL    = SGI_CADENTROTLIT.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADENTROTLIT_IT.SGI_CODIGO    = SGI_CADENTROTLIT.SGI_CODIGO" & vbCrLf
    
    sSql = sSql & "   And SGI_CADENTROTLIT_IT.SGI_FILIAL    = SGI_CADLINHAPRODUTO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADENTROTLIT_IT.SGI_CODCAPAC  = SGI_CADLINHAPRODUTO.SGI_CODLIN" & vbCrLf
    
    sSql = sSql & "   And SGI_CADENTROTLIT_IT.SGI_FILIAL    = SGI_CADPRODUTO.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADENTROTLIT_IT.SGI_IDPRODUTO = SGI_CADPRODUTO.SGI_IDPRODUTO" & vbCrLf
    
    sSql = sSql & "   And SGI_CADENTROTLIT.SGI_FILIAL       = SGI_CADCLIENTE.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And SGI_CADENTROTLIT.SGI_CODCLIE      = SGI_CADCLIENTE.SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then boolTemRegistro = True
    BREC.Close
    
    If boolTemRegistro = False Then
        MsgBox "ATENÇÃO" & vbCrLf & _
               "Não há dados para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    strCABEC2 = "ENVIO DE MATERIAIS"
    
    Call objRel.REL(FILIAL, sSql, strCamRelNovo & cCamRelEstoque & strNomRel, Linha, 1, strCABEC2, "", False)
    
End Sub

Private Function Verif_se_JaExiste(strCODOP As String) As Boolean

    Verif_se_JaExiste = False


    If Len(Trim(strCODOP)) = 0 Then Exit Function

    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADENTROTLIT_IT" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODOP  = " & strCODOP & vbCrLf
    sSql = sSql & "   And SGI_ENTSAI = " & intFILIALPED
    
    BREC8.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC8.EOF() Then Verif_se_JaExiste = True
    BREC8.Close

End Function

Private Function PegaDadosFF(strCODLINHA As String, strCODFF As String, lngLINHA As Long) As Boolean

    PegaDadosFF = True
    
    If Len(Trim(strCODLINHA)) = 0 Then Exit Function
    If Len(Trim(strCODFF)) = 0 Then Exit Function
    If lngLINHA <= 0 Then Exit Function
    
    With grdLANCTOS
    
        sSql = ""
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "        MEDC.*" & vbCrLf
        sSql = sSql & "      , DIMC.SGI_DESCORTE" & vbCrLf
        
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "        SGI_CADLINHAPRODUTO LINH" & vbCrLf
        sSql = sSql & "      , SGI_MEDCORTELINHA   MEDC" & vbCrLf
        sSql = sSql & "      , SGI_CADDIMCORTE     DIMC" & vbCrLf
         
        sSql = sSql & "  Where" & vbCrLf
        sSql = sSql & "        LINH.SGI_FILIAL     = " & FILIAL & vbCrLf
        sSql = sSql & "    And LINH.SGI_CODLIN     = " & strCODLINHA & vbCrLf
        
        sSql = sSql & "    And MEDC.SGI_FILIAL     = LINH.SGI_FILIAL" & vbCrLf
        sSql = sSql & "    And MEDC.SGI_CODIGO     = LINH.SGI_CODIGO" & vbCrLf
        sSql = sSql & "    And MEDC.SGI_CODMEDCORT = " & strCODFF & vbCrLf
        sSql = sSql & "    And DIMC.SGI_FILIAL     = MEDC.SGI_FILIAL" & vbCrLf
        sSql = sSql & "    And DIMC.SGI_CODIGO     = MEDC.SGI_CODMEDCORT"
        
        
        BREC11.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC11.EOF() Then
        
            .Cell(flexcpText, lngLINHA, conCOL_ENTROTLIT_FOLHAUSADA) = Trim(BREC11!SGI_DESCORTE)
            .Cell(flexcpText, lngLINHA, conCOL_ENTROTLIT_ESPESS) = Format(BREC11!SGI_EXPESS, "#,####0.0000")
            .Cell(flexcpText, lngLINHA, conCOL_ENTROTLIT_LARG) = Format(BREC11!SGI_LARGUR, "#,####0.0000")
            .Cell(flexcpText, lngLINHA, conCOL_ENTROTLIT_COMP) = Format(BREC11!SGI_COMPRI, "#,####0.0000")
            .Cell(flexcpText, lngLINHA, conCOL_ENTROTLIT_QTDECORP) = BREC11!SGI_QTDECORPOS
            .Cell(flexcpText, lngLINHA, conCOL_ENTROTLIT_PERDPRODC) = Format(BREC11!SGI_PERDPROC, "#,####0.0000")
            PegaDadosFF = False
        
        Else
            MsgBox "ATENÇÂO" & vbCrLf & "Esta FF não existe para esta linha !!!", vbOKOnly + vbExclamation, "Aviso"
        End If
        BREC11.Close

    End With
End Function

Private Sub txtCLIEDEST_GotFocus()
    objBLBFunc.SelecionaCampos txtCIDCLIE.Name, Me
End Sub

Private Sub txtCLIEDEST_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCLIEDEST.Text
End Sub

Private Sub txtCLIEDEST_Validate(Cancel As Boolean)

    If Len(Trim(txtCLIEDEST.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCLIEDEST.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCLIEDEST.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", txtCLIEDEST.Text, lblDescClienteDest, "CLIE")
    If Len(Trim(lblDescClienteDest.Caption)) = 0 Then
       txtCLIEDEST.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub


Private Function CalcSaldoPeso(strPESO As String, strCONFPEDO As String) As Currency
    
    Dim lngFARDO      As Currency
    Dim lngCONFFARDO  As Currency
    Dim lngSALDOFARDO As Currency

    With grdLANCTOS
        If Len(Trim(strPESO)) > 0 And _
           Len(Trim(strCONFPEDO)) > 0 Then
           
           lngFARDO = CCur(strPESO)
           lngCONFFARDO = CCur(strCONFPEDO)
           lngSALDOFARDO = (lngFARDO - lngCONFFARDO)
           
        End If
    End With
    
    CalcSaldoPeso = lngSALDOFARDO
End Function


Private Function CalcSaldoFardo(strPESO As String, strCONFPEDO As String) As Long
    
    Dim lngFARDO      As Long
    Dim lngCONFFARDO  As Long
    Dim lngSALDOFARDO As Long

    With grdLANCTOS
        If Len(Trim(strPESO)) > 0 And _
           Len(Trim(strCONFPEDO)) > 0 Then
           
           lngFARDO = CLng(strPESO)
           lngCONFFARDO = CLng(strCONFPEDO)
           lngSALDOFARDO = (lngFARDO - lngCONFFARDO)
           
        End If
    End With
    
    CalcSaldoFardo = lngSALDOFARDO
End Function


Private Function CalculaPeso(lngROW As Long, lngFOLHAS As Long) As Double

    CalculaPeso = 0

    Dim curPESO         As Currency
    Dim lngQTDFOLHAS    As Long
    Dim curExpessura    As Currency
    Dim curLargura      As Currency
    Dim curComprimento  As Currency

    With grdLANCTOS
        lngQTDFOLHAS = lngFOLHAS
        curExpessura = 0
        curLargura = 0
        curComprimento = 0
    
        If Len(Trim(.Cell(flexcpText, lngROW, conCOL_ENTROTLIT_ESPESS))) > 0 Then curExpessura = CCur(.Cell(flexcpText, lngROW, conCOL_ENTROTLIT_ESPESS))
        If Len(Trim(.Cell(flexcpText, lngROW, conCOL_ENTROTLIT_LARG))) > 0 Then curLargura = CCur(.Cell(flexcpText, lngROW, conCOL_ENTROTLIT_LARG))
        If Len(Trim(.Cell(flexcpText, lngROW, conCOL_ENTROTLIT_COMP))) > 0 Then curComprimento = CCur(.Cell(flexcpText, lngROW, conCOL_ENTROTLIT_COMP))
    
    End With

    If curExpessura > 0 And curLargura > 0 And curComprimento > 0 Then
        CalculaPeso = (lngFOLHAS * ((curExpessura * (curLargura / 1000) * (curComprimento / 1000) * 7.85)))
    End If

End Function

