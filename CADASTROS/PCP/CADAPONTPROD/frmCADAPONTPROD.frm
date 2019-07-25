VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCADAPONTPROD 
   Caption         =   "Apontamento de Produção"
   ClientHeight    =   8865
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18165
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   18165
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "[ Paradas ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   2175
      Left            =   120
      TabIndex        =   24
      Top             =   6600
      Width           =   11775
      Begin VB.CommandButton Command7 
         Height          =   300
         Left            =   11400
         Picture         =   "frmCADAPONTPROD.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   240
         Width           =   300
      End
      Begin VB.CommandButton Command6 
         Height          =   300
         Left            =   11400
         Picture         =   "frmCADAPONTPROD.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   600
         Width           =   300
      End
      Begin VSFlex8LCtl.VSFlexGrid grdPARADAS 
         Height          =   1815
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   11175
         _cx             =   19711
         _cy             =   3201
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
   Begin VB.CommandButton Command3 
      Height          =   300
      Left            =   17760
      Picture         =   "frmCADAPONTPROD.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2760
      Width           =   300
   End
   Begin VB.CommandButton Command2 
      Height          =   300
      Left            =   17760
      Picture         =   "frmCADAPONTPROD.frx":03DE
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2400
      Width           =   300
   End
   Begin VSFlex8LCtl.VSFlexGrid grdAPONT 
      Height          =   4095
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   17535
      _cx             =   30930
      _cy             =   7223
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
      Height          =   1335
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   18015
      Begin VB.CommandButton Command5 
         Height          =   315
         Left            =   11400
         Picture         =   "frmCADAPONTPROD.frx":0528
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtCODOPERADOR 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   10320
         TabIndex        =   3
         Text            =   "txtCODOPERADOR"
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   2880
         Picture         =   "frmCADAPONTPROD.frx":062A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtCODTURNO 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Text            =   "txtCODTURNO"
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Height          =   315
         Left            =   2880
         Picture         =   "frmCADAPONTPROD.frx":072C
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtCODMAQ 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Text            =   "txtCODMAQ"
         Top             =   600
         Width           =   1215
      End
      Begin MSMask.MaskEdBox mskDTLCTO 
         Height          =   285
         Left            =   5640
         TabIndex        =   1
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblDescOperador 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescOperador"
         Height          =   285
         Left            =   11760
         TabIndex        =   23
         Top             =   600
         Width           =   5775
      End
      Begin VB.Label lblDescTurno 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescTurno"
         Height          =   285
         Left            =   3240
         TabIndex        =   22
         Top             =   960
         Width           =   5895
      End
      Begin VB.Label lblDescMaq 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescMaq"
         Height          =   285
         Left            =   3240
         TabIndex        =   21
         Top             =   600
         Width           =   5895
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Operador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   5
         Left            =   9240
         TabIndex        =   19
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data do Lançamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   4
         Left            =   3240
         TabIndex        =   13
         Top             =   240
         Width           =   2145
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Turno"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Máquina"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   270
         Width           =   750
      End
      Begin VB.Label lblCODIGO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblCODIGO"
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   18015
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
         Left            =   1560
         Picture         =   "frmCADAPONTPROD.frx":082E
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdSalva 
         Caption         =   "&Salva <F2>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2400
         Picture         =   "frmCADAPONTPROD.frx":0D60
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdVoltar 
         Caption         =   "&Volta <ESC>"
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
         Picture         =   "frmCADAPONTPROD.frx":0E62
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmCADAPONTPROD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public cTipOper         As String
Public iCodigo          As Integer
Public iParcela         As Integer
Public FILIAL           As Integer
Public strACESSO        As String
Public strMODPAI        As String
Public strUsuario       As String
Public lngCODUSUARIO    As Long
Public intFILIALPED     As Integer
Public strFILIAL        As String


Dim lngCodLog           As Long
Dim strValor            As String
Dim strCAPTION          As String
Dim strNOMFILIAL        As String

Dim objBLBFunc          As New clsFuncoes
Dim objCADAPONTPROD     As New clsCADAPONTPROD
Dim objPESQPADRAO       As Object

Dim arrAPONT            As Variant
Dim arrPARADAS          As Variant

Const conCOL_SonMov_CodOP                                   As Integer = 0
Const conCOL_SonMov_Capacidade                              As Integer = 1
Const conCOL_SonMov_CodProd                                 As Integer = 2
Const conCOL_SonMov_DescRot                                 As Integer = 3
Const conCOL_SonMov_CodTipo                                 As Integer = 4
Const conCOL_SonMov_PesqTipo                                As Integer = 5
Const conCOL_SonMov_DescTipo                                As Integer = 6
Const conCOL_SonMov_AplicLote                               As Integer = 7
Const conCOL_SonMov_Lote                                    As Integer = 8
Const conCOL_SonMov_QtdeEntr                                As Integer = 9
Const conCOL_SonMov_QtdeSaida                               As Integer = 10
Const conCOL_SonMov_Retrabalho                              As Integer = 11
Const conCOL_SonMov_IDPRODUTO                               As Integer = 12
Const conCOL_SonMov_INDICE                                  As Integer = 13
Const conCOL_SonMov_Action2Do                               As Integer = 14
Const conCOL_SonMov_HORINI                                  As Integer = 15
Const conCOL_SonMov_HORFIN                                  As Integer = 16
Const conCOL_SonMov_TotalLiq                                As Integer = 17
Const conCOL_SonMov_IDINTERNO                               As Integer = 18
Const conCOL_SonMov_INDICEBKP                               As Integer = 19
Const conCOL_SonMov_FILIALPED                               As Integer = 20
Const conCOL_SonMov_DESCFILIALPED                           As Integer = 21
Const conCOL_SonMov_FormatString                            As String = "=Cod.OP|Capacidade|Rótulo|Descr.Rótulo|Cód.TIpo|...|Desc.Tipo|Aplicação/Lote|Lote|Qtde.Entrada|Qtde.Saida|Retrabalho|IDPRODUTO|INDICE|Action2Do|Hor.Ini|Hor.Fin|Tot.Hor|IDINTERNO|INDICEBKP|FILIALPED|FILIAL"
Const conColumnsIn_SonMov                                   As Integer = 22

Const conCOL_SonMovParada_Indice                            As Integer = 0
Const conCOL_SonMovParada_CodParada                         As Integer = 1
Const conCOL_SonMovParada_CodIntParada                      As Integer = 2
Const conCOL_SonMovParada_PesqParada                        As Integer = 3
Const conCOL_SonMovParada_DescParada                        As Integer = 4
Const conCOL_SonMovParada_Action2Do                         As Integer = 5
Const conCOL_SonMovParada_IDINTERNO                         As Integer = 6
Const conCOL_SonMovParada_HORINI                            As Integer = 7
Const conCOL_SonMovParada_HORFIN                            As Integer = 8
Const conCOL_SonMovParada_TotalLiq                          As Integer = 9
Const conCOL_SonMovParada_IDPAI                             As Integer = 10
Const conCOL_SonMovParada_FILIALPED                         As Integer = 11
Const conCOL_SonMovParada_FormatString                      As String = "=Indice|ID|Código|...|Decrição|Action2Do|IDINTERNO|Hor.Ini|Hor.Fin|Tot.Hor|IDPAI|FILIALPED"
Const conColumnsIn_SonMovParada                             As Integer = 12

Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strACESSO) = False Then Exit Sub
    cTipOper = "A"
    
    Call objBLBFunc.MudaBotoes(CmdSalva, cmdAltera, cTipOper)
    Call objBLBFunc.MudaCaption(Me, strCAPTION, cTipOper)
    Call DesabilitaCampos

End Sub

Private Sub CmdSalva_Click()

    Dim I           As Integer
    Dim j           As Integer
    Dim intRESP     As Integer
    Dim lngITENS    As Integer
    Dim sValor      As String
    
    Call objBLBFunc.RemoveLinhaVazia(grdAPONT, conCOL_SonMov_CodOP)
    Call objBLBFunc.RemoveLinhaVazia(grdPARADAS, conCOL_SonMovParada_CodParada)
    
    If Valida_Campos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADAPONTPROD.Codigo = objBLBFunc.Gera_Codigo(Trim(Me.Name) & strNOMFILIAL, FILIAL, Linha)

    objCADAPONTPROD.CODMAQ = Trim(txtCODMAQ.Text)
    objCADAPONTPROD.CODOPER = Trim(txtCODOPERADOR.Text)
    objCADAPONTPROD.CODTURN = Trim(txtCODTURNO.Text)
    objCADAPONTPROD.DTLANCT = "'" & Format(CDate(mskDTLCTO.Text), "MM/DD/YYYY") & "'"

    '' Apontamento
    arrAPONT = Empty
    With grdAPONT
        If (.Rows - 1) > 0 Then
            ReDim arrAPONT(1 To (.Rows - 1), 1 To 14) As String
            For I = 1 To (.Rows - 1)
            
                arrAPONT(I, 1) = .Cell(flexcpText, I, conCOL_SonMov_CodOP)
                arrAPONT(I, 2) = .Cell(flexcpText, I, conCOL_SonMov_CodTipo)
                
                arrAPONT(I, 3) = "Null"
                If Len(Trim(.Cell(flexcpText, I, conCOL_SonMov_AplicLote))) > 0 Then arrAPONT(I, 3) = "'" & Trim(.Cell(flexcpText, I, conCOL_SonMov_AplicLote)) & "'"
                
                arrAPONT(I, 4) = "Null"
                If Len(Trim(.Cell(flexcpText, I, conCOL_SonMov_Lote))) > 0 Then arrAPONT(I, 4) = "'" & Trim(.Cell(flexcpText, I, conCOL_SonMov_Lote)) & "'"
                 
                arrAPONT(I, 5) = .Cell(flexcpText, I, conCOL_SonMov_QtdeEntr)
                arrAPONT(I, 6) = .Cell(flexcpText, I, conCOL_SonMov_QtdeSaida)
                
                arrAPONT(I, 7) = "Null"
                If Len(Trim(.Cell(flexcpText, I, conCOL_SonMov_Retrabalho))) > 0 Then arrAPONT(I, 7) = .Cell(flexcpText, I, conCOL_SonMov_Retrabalho)
                
                arrAPONT(I, 8) = "'" & Trim(.Cell(flexcpText, I, conCOL_SonMov_HORINI)) & "'"
                arrAPONT(I, 9) = "'" & Trim(.Cell(flexcpText, I, conCOL_SonMov_HORFIN)) & "'"
                arrAPONT(I, 10) = "'" & Trim(.Cell(flexcpText, I, conCOL_SonMov_TotalLiq)) & "'"
                arrAPONT(I, 11) = "'" & Trim(.Cell(flexcpText, I, conCOL_SonMov_INDICE)) & "'"
                
                If .Cell(flexcpText, I, conCOL_SonMov_Action2Do) = dacEnumUpdateAction_Insert Then
                    arrAPONT(I, 12) = objBLBFunc.Gera_Codigo(Trim(Me.Name) & strNOMFILIAL & "_APID", FILIAL, Linha)
                Else
                    arrAPONT(I, 12) = .Cell(flexcpText, I, conCOL_SonMov_IDINTERNO)
                End If
            
                arrAPONT(I, 13) = .Cell(flexcpText, I, conCOL_SonMov_Action2Do)
                arrAPONT(I, 14) = .Cell(flexcpText, I, conCOL_SonMov_FILIALPED)
            
            Next I
        End If
    End With
    objCADAPONTPROD.APONT = arrAPONT

    '' Paradas
    arrPARADAS = Empty
    With grdPARADAS
        If (.Rows - 1) > 0 Then
            ReDim arrPARADAS(1 To (.Rows - 1), 1 To 9) As String
            For I = 1 To (.Rows - 1)
            
                arrPARADAS(I, 1) = .Cell(flexcpText, I, conCOL_SonMovParada_CodParada)
                arrPARADAS(I, 2) = "'" & Trim(.Cell(flexcpText, I, conCOL_SonMovParada_HORINI)) & "'"
                arrPARADAS(I, 3) = "'" & Trim(.Cell(flexcpText, I, conCOL_SonMovParada_HORFIN)) & "'"
                arrPARADAS(I, 4) = "'" & Trim(.Cell(flexcpText, I, conCOL_SonMovParada_TotalLiq)) & "'"
                arrPARADAS(I, 5) = .Cell(flexcpText, I, conCOL_SonMovParada_Action2Do)
                
                If .Cell(flexcpText, I, conCOL_SonMovParada_Action2Do) = dacEnumUpdateAction_Insert Then
                    arrPARADAS(I, 6) = objBLBFunc.Gera_Codigo(Trim(Me.Name) & strNOMFILIAL & "_PAID", FILIAL, Linha)
                Else
                    arrPARADAS(I, 6) = .Cell(flexcpText, I, conCOL_SonMovParada_IDINTERNO)
                End If
            
                arrPARADAS(I, 7) = .Cell(flexcpText, I, conCOL_SonMovParada_IDPAI)
                arrPARADAS(I, 8) = "'" & Trim(.Cell(flexcpText, I, conCOL_SonMovParada_Indice)) & "'"
                arrPARADAS(I, 9) = .Cell(flexcpText, I, conCOL_SonMovParada_FILIALPED)
                
            Next I
        End If
    End With
    objCADAPONTPROD.PARADAS = arrPARADAS

    If objCADAPONTPROD.GRAVA(cTipOper, strNOMFILIAL) = False Then Exit Sub

    MsgBox "O Apontamento foi " & IIf(cTipOper = "I", "incluso", IIf(cTipOper = "A", "alterado", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
    If cTipOper = "I" Then Unload Me
    If cTipOper = "A" Then
        Call LimpaCamposlabel
        Call InitGridMov
        Call InitGridParada
        Call CarregaCampos
    End If
    

End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

    If Len(Trim(txtCODMAQ.Text)) = 0 Then
        MsgBox "Informe Primeiro o Código da Máquina que será realizado o Apontamento !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If


    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       MAQTUR.SGI_CODIGO" & vbCrLf
    sSql = sSql & "      ,MAQTUR.SGI_DESCRI" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADMAQTURN  CADTUR" & vbCrLf
    sSql = sSql & "      ,SGI_CADQTDETURN MAQTUR" & vbCrLf

    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       CADTUR.SGI_FILIAL  = " & FILIAL & vbCrLf
    sSql = sSql & "   And CADTUR.SGI_CODIGO  = " & Trim(txtCODMAQ.Text) & vbCrLf
    sSql = sSql & "   And CADTUR.SGI_FILIAL  = MAQTUR.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And CADTUR.SGI_CODTURN = MAQTUR.SGI_CODIGO" & vbCrLf
    sSql = sSql & "   And MAQTUR.SGI_ATIVO = 1"
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Cód.Turno"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "MAQTUR.SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRI"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "MAQTUR.SGI_DESCRI"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strACESSO, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Turnos")
    
    If Len(Trim(varRETORNO)) > 0 Then
        txtCODTURNO.Text = varRETORNO
        Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRI", "SGI_CADQTDETURN", varRETORNO, lblDescTurno, "ATENÇÃO - Turno Inexistente !!!")
        txtCODOPERADOR.SetFocus
    End If

End Sub

Private Sub Command2_Click()
    Call IncRegGrid
End Sub

Private Sub Command3_Click()

On Error GoTo Err_Command3_Click
    
    If cTipOper = "C" Then Exit Sub
    If cTipOper = "I" Or cTipOper = "A" Then
        With grdAPONT
            If Len(Trim(.Cell(flexcpText, .Row, conCOL_SonMov_CodOP))) = 0 Then Exit Sub
            If .Cell(flexcpText, .Row, conCOL_SonMov_Action2Do) = dacEnumUpdateAction_Insert Then
                If (.Rows - 1) = 1 Then .Rows = 1
                If (.Rows - 1) > 1 Then
                   Call objBLBFunc.ExcLinhaGrdFilhoAct2Do(grdPARADAS, conCOL_SonMovParada_Indice, grdAPONT.Cell(flexcpText, .Row, conCOL_SonMov_INDICE), conCOL_SonMovParada_Action2Do)
                   Call objBLBFunc.ExclLinhaGrid(grdAPONT, grdAPONT.Row)
                End If
            Else
                .Cell(flexcpText, .Row, conCOL_SonMov_Action2Do) = dacEnumUpdateAction_delete
                Call objBLBFunc.ExcLinhaGrdFilhoAct2Do(grdPARADAS, conCOL_SonMovParada_Indice, grdAPONT.Cell(flexcpText, .Row, conCOL_SonMov_INDICE), conCOL_SonMovParada_Action2Do)
                Call objBLBFunc.ExclLinhaGridAction2Do(grdAPONT, .Row, conCOL_SonMov_Action2Do)
            End If
        End With
    End If
    
    Exit Sub
    
Err_Command3_Click:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Command3_Click()", Me.Name, "Command3_Click()", strCAMARQERRO)

End Sub

Private Sub Command4_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADMAQUINA " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_ATIVA  = 0" '' Maquinas Ativas
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRI"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "3000"
    arrCAMPOS(2, 5) = "SGI_DESCRI"
    
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strACESSO, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Máquinas")
    
    If Len(Trim(varRETORNO)) > 0 Then
        txtCODMAQ.Text = varRETORNO
        Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRI", "SGI_CADMAQUINA", varRETORNO, lblDescMaq, "ATENÇÃO - Maquina Inexistente !!!")
        txtCODTURNO.SetFocus
    End If

End Sub

Private Sub Command5_Click()

    If Len(Trim(txtCODMAQ.Text)) = 0 Then
        MsgBox "ATENÇÃO - Primeiro informe a máquina !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       OPERA.SGI_CODIGO " & vbCrLf
    sSql = sSql & "      ,OPERA.SGI_DESCRI " & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADMAQOPER  MAQOPE" & vbCrLf
    sSql = sSql & "      ,SGI_CADOPERADOR OPERA" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       MAQOPE.SGI_FILIAL  = " & FILIAL & vbCrLf
    sSql = sSql & "   And MAQOPE.SGI_CODIGO  = " & Trim(txtCODMAQ.Text) & vbCrLf
    sSql = sSql & "   And MAQOPE.SGI_FILIAL  = OPERA.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And MAQOPE.SGI_CODOPER = OPERA.SGI_CODIGO" & vbCrLf
    sSql = sSql & "   And OPERA.SGI_ATIVO = 1"
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "OPERA.SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRI"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "OPERA.SGI_DESCRI"
    
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strACESSO, V_Usuario, arrCAMPOS, arrTABELA, "Operadores")
    
    If Len(Trim(varRETORNO)) > 0 Then
        txtCODOPERADOR.Text = varRETORNO
        Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRI", "SGI_CADOPERADOR", varRETORNO, lblDescOperador, "ATENÇÃO - Operador Inexistente !!!")
    End If

End Sub

Private Sub Command6_Click()

On Error GoTo Err_Command6_Click
    
    If cTipOper = "C" Then Exit Sub
    If cTipOper = "I" Or cTipOper = "A" Then
        With grdPARADAS
            If Len(Trim(.Cell(flexcpText, .Row, conCOL_SonMovParada_CodParada))) = 0 Then Exit Sub
            If .Cell(flexcpText, .Row, conCOL_SonMovParada_Action2Do) = dacEnumUpdateAction_Insert Then
                If (.Rows - 1) = 1 Then .Rows = 1
                If (.Rows - 1) > 1 Then Call objBLBFunc.ExclLinhaGrid(grdPARADAS, grdPARADAS.Row)
            Else
                .Cell(flexcpText, .Row, conCOL_SonMovParada_Action2Do) = dacEnumUpdateAction_delete
                Call objBLBFunc.ExclLinhaGridAction2Do(grdPARADAS, .Row, conCOL_SonMovParada_Action2Do)
            End If
        End With
    End If
    
    Exit Sub
    
Err_Command6_Click:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Command6_Click()", Me.Name, "Command6_Click()", strCAMARQERRO)

End Sub

Private Sub Command7_Click()
    Call IncRegGridParada
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

''    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
''    Set objCADAPONTPROD = CreateObject("CADAPONTPROD.clsCADAPONTPROD")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
    objCADAPONTPROD.FILIAL = FILIAL
   
    If intFILIALPED = 0 Then strFILIAL = "NOVALATA"
    If intFILIALPED = 1 Then strFILIAL = "STEEL"
    
    strNOMFILIAL = ""
    If intFILIALPED = 1 Then strNOMFILIAL = "_STEEL"
    
    strCAPTION = "Apontamento de Produção "
    
    Call IniciaForm

End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objCADAPONTPROD = Nothing
    Set objPESQPADRAO = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

Private Sub IniciaForm()

    Call objBLBFunc.MudaBotoes(CmdSalva, cmdAltera, cTipOper)
    Call objBLBFunc.MudaCaption(frmCADAPONTPROD, strCAPTION, cTipOper)
    Call objBLBFunc.LimpaCampos(frmCADAPONTPROD)
    
    Call LimpaCamposlabel
    Call DesabilitaCampos
    Call InitGridMov
    Call InitGridParada
    
    objCADAPONTPROD.Codigo = iCodigo
    mskDTLCTO.Text = Format(Now, "DD/MM/YYYY")
    
    Call CarregaCampos
    
End Sub


Private Sub LimpaCamposlabel()
    lblCODIGO.Caption = ""
    lblDescMaq.Caption = ""
    lblDescTurno.Caption = ""
    lblDescOperador.Caption = ""
End Sub

Private Sub DesabilitaCampos()
    If cTipOper = "I" Then Frame2.Enabled = True
    If cTipOper = "C" Or cTipOper = "A" Then Frame2.Enabled = False
End Sub

Private Sub grdAPONT_AfterEdit(ByVal Row As Long, ByVal Col As Long)

On Error GoTo Err_grdAPONT_AfterEdit

    Dim strTOTALPERIODO As String
    Dim dtTotalLiquido  As Date

     With grdAPONT
          Select Case Col
                Case conCOL_SonMov_HORINI, _
                     conCOL_SonMov_HORFIN
                     
                        If Len(Trim(Replace(.Cell(flexcpText, Row, conCOL_SonMov_HORINI), ":", ""))) > 0 And _
                           Len(Trim(Replace(.Cell(flexcpText, Row, conCOL_SonMov_HORFIN), ":", ""))) > 0 Then
                           
                           strTOTALPERIODO = objBLBFunc.CalcTempo(.Cell(flexcpText, Row, conCOL_SonMov_HORINI), .Cell(flexcpText, Row, conCOL_SonMov_HORFIN))
                           
                           dtTotalLiquido = CDate(strTOTALPERIODO)
                           .Cell(flexcpText, Row, conCOL_SonMov_TotalLiq) = Format(dtTotalLiquido, "HH:MM")
                           
                        End If
          End Select
     End With
     Exit Sub

Err_grdAPONT_AfterEdit:
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdAPONT_AfterEdit", Me.Name, "AfterEdit")

End Sub

Private Sub grdAPONT_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

On Error GoTo Err_grdAPONT_BeforeEdit
    
    With grdAPONT
        Select Case Col
            Case conCOL_SonMov_Capacidade, _
                 conCOL_SonMov_CodProd, _
                 conCOL_SonMov_DescRot, _
                 conCOL_SonMov_DescTipo, _
                 conCOL_SonMov_IDPRODUTO, _
                 conCOL_SonMov_INDICE, _
                 conCOL_SonMov_Action2Do, _
                 conCOL_SonMov_TotalLiq, _
                 conCOL_SonMov_IDINTERNO, _
                 conCOL_SonMov_INDICEBKP, _
                 conCOL_SonMov_FILIALPED, _
                 conCOL_SonMov_DESCFILIALPED
                 Cancel = True
            Case conCOL_SonMov_CodOP, _
                 conCOL_SonMov_CodTipo, _
                 conCOL_SonMov_PesqTipo, _
                 conCOL_SonMov_QtdeEntr, _
                 conCOL_SonMov_QtdeSaida, _
                 conCOL_SonMov_Retrabalho, _
                 conCOL_SonMov_AplicLote, _
                 conCOL_SonMov_Lote, _
                 conCOL_SonMov_HORINI, _
                 conCOL_SonMov_HORFIN
                 If cTipOper = "C" Then Cancel = True
            Case Else
                .ComboList = ""
        End Select
    End With
    
    Exit Sub

Err_grdAPONT_BeforeEdit:
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdAPONT_BeforeEdit", Me.Name, "BeforeEdit")

End Sub

Private Sub grdAPONT_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

On Error GoTo Err_grdAPONT_CellButtonClick
    
    With grdAPONT
        
        If (.Rows - 1) = 0 Then
            MsgBox "Insira 1 registro !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
        End If
        If (.Row = 0) Then
            MsgBox "Selecione 1 registro !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
        End If
        If Len(Trim(.Cell(flexcpText, .Row, conCOL_SonMov_CodOP))) = 0 Then
            MsgBox "Informe o Código da OP !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
        End If
        
        Select Case Col
            Case conCOL_SonMov_PesqTipo
                
                ReDim arrCAMPOS(1 To 2, 1 To 5) As String
                ReDim arrTABELA(1 To 1) As String
                
                sSql = ""
                
                sSql = "Select " & vbCrLf
                sSql = sSql & "       *" & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_CADTIPAPONT" & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       SGI_FILIAL      = " & FILIAL & vbCrLf
                sSql = sSql & "   And SGI_ATIVO       = 1"
                
                arrTABELA(1) = sSql
                
                arrCAMPOS(1, 1) = "SGI_CODIGO"
                arrCAMPOS(1, 2) = "N"
                arrCAMPOS(1, 3) = "Código"
                arrCAMPOS(1, 4) = "1500"
                arrCAMPOS(1, 5) = "SGI_CODIGO"
                
                arrCAMPOS(2, 1) = "SGI_DESCRI"
                arrCAMPOS(2, 2) = "S"
                arrCAMPOS(2, 3) = "Descrição"
                arrCAMPOS(2, 4) = "4000"
                arrCAMPOS(2, 5) = "SGI_DESCRI"
                
                varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strACESSO, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Tipo de Apontamento")
                
                If Len(Trim(varRETORNO)) > 0 Then
                
                    If .Cell(flexcpText, Row, conCOL_SonMov_Action2Do) = dacEnumUpdateAction_Ignore Then
                        If Trim(varRETORNO) <> Trim(.Cell(flexcpText, Row, Col)) Then .Cell(flexcpText, Row, conCOL_SonMov_Action2Do) = dacEnumUpdateAction_update
                    End If
                
                    Call IqualaTipoApont(varRETORNO, Row)
                    
                    Call CriaChave(.Cell(flexcpText, Row, conCOL_SonMov_CodOP), .Cell(flexcpText, Row, conCOL_SonMov_IDPRODUTO), varRETORNO, Row)
                    If objBLBFunc.FcVerifItensRepetidos(grdAPONT, Row, conCOL_SonMov_INDICE, .Cell(flexcpText, Row, conCOL_SonMov_INDICE)) = False Then
                       MsgBox "Este produto ja foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                       .Cell(flexcpText, Row, Col) = Empty
                       .Cell(flexcpText, Row, conCOL_SonMov_DescTipo) = Empty
                       .Cell(flexcpText, Row, conCOL_SonMov_INDICE) = Empty
                       Exit Sub
                    End If
                    
                    Call TrocaIndice(.Cell(flexcpText, Row, conCOL_SonMov_INDICEBKP), .Cell(flexcpText, Row, conCOL_SonMov_INDICE), Row)
                    
                    Exit Sub
                End If
            
        End Select
    
    End With
    
    Exit Sub

Err_grdAPONT_CellButtonClick:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdAPONT_CellButtonClick", Me.Name, "CellButtonClick")

End Sub

Private Sub grdAPONT_Click()
    Call MostraDados
End Sub

Private Sub grdAPONT_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)

On Error GoTo Err_grdAPONT_KeyPressEdit
     
     With grdAPONT
          Select Case Col
                    Case conCOL_SonMov_CodOP, _
                         conCOL_SonMov_CodTipo, _
                         conCOL_SonMov_QtdeEntr, _
                         conCOL_SonMov_QtdeSaida, _
                         conCOL_SonMov_Retrabalho
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
          End Select
     End With

     Exit Sub

Err_grdAPONT_KeyPressEdit:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdAPONT_KeyPressEdit", Me.Name, "KeyPressEdit")

End Sub

Private Sub grdAPONT_RowColChange()
    Call MostraDados
End Sub

Private Sub grdAPONT_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

On Error GoTo Err_grdAPONT_ValidateEdit
     
    Dim intHoras        As Integer
    Dim intMinutos      As Integer
    Dim boolOPNOVALATA  As Boolean
    Dim boolOPSTEEL     As Boolean
     
     With grdAPONT
          Select Case Col
                 Case conCOL_SonMov_CodOP
                        If .EditText = Empty Then Exit Sub
                        If Not IsNumeric(.EditText) Then
                            MsgBox "Códgo da OP inválido !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                        
                        If .Cell(flexcpText, Row, conCOL_SonMov_Action2Do) = dacEnumUpdateAction_Ignore Then
                            If Trim(.EditText) <> Trim(.Cell(flexcpText, Row, Col)) Then .Cell(flexcpText, Row, conCOL_SonMov_Action2Do) = dacEnumUpdateAction_update
                        End If
                        
                        boolOPNOVALATA = PegaOPNOVALATA(Trim(.EditText), Row)
                        boolOPSTEEL = PegaOPSTEEL(Trim(.EditText), Row)
                        
                        If boolOPNOVALATA = False And boolOPSTEEL = False Then
                            MsgBox "Esta OP - Não Existe !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                        
                        
                        Call CriaChave(.EditText, .Cell(flexcpText, Row, conCOL_SonMov_IDPRODUTO), .Cell(flexcpText, Row, conCOL_SonMov_CodTipo), Row)
                        If objBLBFunc.FcVerifItensRepetidos(grdAPONT, Row, conCOL_SonMov_INDICE, .Cell(flexcpText, Row, conCOL_SonMov_INDICE)) = False Then
                           MsgBox "Esta OP ja foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, Col) = Empty
                           .Cell(flexcpText, Row, conCOL_SonMov_Capacidade) = Empty
                           .Cell(flexcpText, Row, conCOL_SonMov_CodProd) = Empty
                           .Cell(flexcpText, Row, conCOL_SonMov_DescRot) = Empty
                           .Cell(flexcpText, Row, conCOL_SonMov_INDICE) = Empty
                           .Cell(flexcpText, Row, conCOL_SonMov_IDPRODUTO) = Empty
                           Cancel = True
                           Exit Sub
                        End If
                 
                        Call TrocaIndice(.Cell(flexcpText, Row, conCOL_SonMov_INDICEBKP), .Cell(flexcpText, Row, conCOL_SonMov_INDICE), Row)
                 
                 Case conCOL_SonMov_CodTipo
                        
                        If Len(Trim(.Cell(flexcpText, Row, conCOL_SonMov_CodOP))) = 0 Then
                            MsgBox "Primeiro Informe o Código da OP !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                        
                        If .EditText = Empty Then Exit Sub
                        If Not IsNumeric(.EditText) Then
                            MsgBox "Códgo do Tipo de Atividade inválido !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, Col) = Empty
                           .Cell(flexcpText, Row, conCOL_SonMov_DescTipo) = Empty
                           .Cell(flexcpText, Row, conCOL_SonMov_INDICE) = Empty
                            Cancel = True
                            Exit Sub
                        End If
                        
                        If .Cell(flexcpText, Row, conCOL_SonMov_Action2Do) = dacEnumUpdateAction_Ignore Then
                            If Trim(.EditText) <> Trim(.Cell(flexcpText, Row, Col)) Then .Cell(flexcpText, Row, conCOL_SonMov_Action2Do) = dacEnumUpdateAction_update
                        End If
                        
                        Call IqualaTipoApont(.EditText, Row)
                        
                        Call CriaChave(.Cell(flexcpText, Row, conCOL_SonMov_CodOP), .Cell(flexcpText, Row, conCOL_SonMov_IDPRODUTO), .EditText, Row)
                        If objBLBFunc.FcVerifItensRepetidos(grdAPONT, Row, conCOL_SonMov_INDICE, .Cell(flexcpText, Row, conCOL_SonMov_INDICE)) = False Then
                           MsgBox "Este produto ja foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, Col) = Empty
                           .Cell(flexcpText, Row, conCOL_SonMov_DescTipo) = Empty
                           .Cell(flexcpText, Row, conCOL_SonMov_INDICE) = Empty
                           Cancel = True
                           Exit Sub
                        End If
                 
                 Case conCOL_SonMov_AplicLote
                        If .EditText = Empty Then Exit Sub
                        If .Cell(flexcpText, Row, conCOL_SonMov_Action2Do) = dacEnumUpdateAction_Ignore Then
                            If Trim(.EditText) <> Trim(.Cell(flexcpText, Row, Col)) Then .Cell(flexcpText, Row, conCOL_SonMov_Action2Do) = dacEnumUpdateAction_update
                        End If
                 Case conCOL_SonMov_Lote
                        If .EditText = Empty Then Exit Sub
                        If .Cell(flexcpText, Row, conCOL_SonMov_Action2Do) = dacEnumUpdateAction_Ignore Then
                            If Trim(.EditText) <> Trim(.Cell(flexcpText, Row, Col)) Then .Cell(flexcpText, Row, conCOL_SonMov_Action2Do) = dacEnumUpdateAction_update
                        End If
                 Case conCOL_SonMov_QtdeEntr
                        If .EditText = Empty Then Exit Sub
                        If .Cell(flexcpText, Row, conCOL_SonMov_Action2Do) = dacEnumUpdateAction_Ignore Then
                            If Trim(.EditText) <> Trim(.Cell(flexcpText, Row, Col)) Then .Cell(flexcpText, Row, conCOL_SonMov_Action2Do) = dacEnumUpdateAction_update
                        End If
                 Case conCOL_SonMov_QtdeSaida
                        If .EditText = Empty Then Exit Sub
                        If .Cell(flexcpText, Row, conCOL_SonMov_Action2Do) = dacEnumUpdateAction_Ignore Then
                            If Trim(.EditText) <> Trim(.Cell(flexcpText, Row, Col)) Then .Cell(flexcpText, Row, conCOL_SonMov_Action2Do) = dacEnumUpdateAction_update
                        End If
                 Case conCOL_SonMov_Retrabalho
                        If .EditText = Empty Then Exit Sub
                        If .Cell(flexcpText, Row, conCOL_SonMov_Action2Do) = dacEnumUpdateAction_Ignore Then
                            If Trim(.EditText) <> Trim(.Cell(flexcpText, Row, Col)) Then .Cell(flexcpText, Row, conCOL_SonMov_Action2Do) = dacEnumUpdateAction_update
                        End If
                 Case conCOL_SonMov_HORINI, _
                      conCOL_SonMov_HORFIN
                      If .EditText = "  :  " Or _
                         Len(Trim(.EditText)) = 0 Then
                         Cancel = False
                         Exit Sub
                      End If
                     
                      '' ==================================
                      '' Validando Campo Horas
                      If Len(Trim(Mid(.EditText, 1, 2))) = 0 And Len(Trim(Mid(.EditText, 4, 2))) > 0 Then
                         MsgBox "Hora inválida !!!", vbOKOnly + vbExclamation, "Aviso"
                         Cancel = True
                         Exit Sub
                      End If
                      If Len(Trim(Mid(.EditText, 1, 2))) > 0 And Len(Trim(Mid(.EditText, 4, 2))) = 0 Then
                         MsgBox "Hora inválida !!!", vbOKOnly + vbExclamation, "Aviso"
                         Cancel = True
                         Exit Sub
                      End If
                     
                     
                      '' ==================================
                      '' Validando Horas
                      intHoras = CInt(Mid(.EditText, 1, 2))
                      intMinutos = CInt(Mid(.EditText, 4, 2))
                      If intHoras >= 24 Or intHoras < 0 Then
                         MsgBox "Hora Inválida o Dia vai somente até 24:00 !!!", vbOKOnly + vbExclamation, "Aviso"
                         Cancel = True
                         Exit Sub
                      End If
                      If intMinutos >= 60 Then
                         MsgBox "Minutos Inválido os minutos somente devem ser informados de 00 a 59 !!!", vbOKOnly + vbExclamation, "Aviso"
                         Cancel = True
                         Exit Sub
                      End If
                    
                    If .Cell(flexcpText, Row, conCOL_SonMov_Action2Do) = dacEnumUpdateAction_Ignore Then
                        If Trim(.EditText) <> Trim(.Cell(flexcpText, Row, Col)) Then .Cell(flexcpText, Row, conCOL_SonMov_Action2Do) = dacEnumUpdateAction_update
                    End If
                 
                 
          End Select
     End With
    
    Exit Sub

Err_grdAPONT_ValidateEdit:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdAPONT_ValidateEdit", Me.Name, "ValidateEdit")

End Sub


Private Sub grdPARADAS_AfterEdit(ByVal Row As Long, ByVal Col As Long)

On Error GoTo Err_grdPARADAS_AfterEdit

    Dim strTOTALPERIODO     As String
    Dim dtTotalLiquido      As Date

     With grdPARADAS
          Select Case Col
                Case conCOL_SonMovParada_HORINI, _
                     conCOL_SonMovParada_HORFIN
                     
                        If Len(Trim(Replace(.Cell(flexcpText, Row, conCOL_SonMovParada_HORINI), ":", ""))) > 0 And _
                           Len(Trim(Replace(.Cell(flexcpText, Row, conCOL_SonMovParada_HORFIN), ":", ""))) > 0 Then
                           
                           strTOTALPERIODO = objBLBFunc.CalcTempo(.Cell(flexcpText, Row, conCOL_SonMovParada_HORINI), .Cell(flexcpText, Row, conCOL_SonMovParada_HORFIN))
                           
                           dtTotalLiquido = CDate(strTOTALPERIODO)
                           .Cell(flexcpText, Row, conCOL_SonMovParada_TotalLiq) = Format(dtTotalLiquido, "HH:MM")
                           
                        End If
          
          End Select
     End With
     Exit Sub

Err_grdPARADAS_AfterEdit:
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdPARADAS_AfterEdit", Me.Name, "AfterEdit")

End Sub

Private Sub grdPARADAS_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

On Error GoTo Err_grdPARADAS_BeforeEdit
    
    With grdAPONT
        Select Case Col
            Case conCOL_SonMovParada_Indice, _
                 conCOL_SonMovParada_CodParada, _
                 conCOL_SonMovParada_DescParada, _
                 conCOL_SonMovParada_Action2Do, _
                 conCOL_SonMovParada_IDINTERNO, _
                 conCOL_SonMovParada_TotalLiq, _
                 conCOL_SonMovParada_IDINTERNO, _
                 conCOL_SonMovParada_IDPAI, _
                 conCOL_SonMovParada_FILIALPED
                 Cancel = True
            Case conCOL_SonMovParada_CodIntParada, _
                 conCOL_SonMovParada_PesqParada, _
                 conCOL_SonMovParada_HORINI, _
                 conCOL_SonMovParada_HORFIN
                 If cTipOper = "C" Then Cancel = True
            Case Else
                .ComboList = ""
        End Select
    End With
    
    Exit Sub

Err_grdPARADAS_BeforeEdit:
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdPARADAS_BeforeEdit", Me.Name, "BeforeEdit")

End Sub

Private Sub grdPARADAS_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

On Error GoTo Err_grdPARADAS_CellButtonClick
    
    With grdPARADAS
        
        Select Case Col
            Case conCOL_SonMovParada_PesqParada
                
                ReDim arrCAMPOS(1 To 3, 1 To 5) As String
                ReDim arrTABELA(1 To 1) As String
                
                sSql = ""
                
                sSql = "Select " & vbCrLf
                sSql = sSql & "       *" & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_CADPARADAS" & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       SGI_FILIAL      = " & FILIAL & vbCrLf
                sSql = sSql & "   And SGI_ATIVO       = 1"
                
                arrTABELA(1) = sSql
                
                arrCAMPOS(1, 1) = "SGI_CODIGO"
                arrCAMPOS(1, 2) = "N"
                arrCAMPOS(1, 3) = "ID"
                arrCAMPOS(1, 4) = "1500"
                arrCAMPOS(1, 5) = "SGI_CODIGO"
                
                arrCAMPOS(2, 1) = "SGI_CODINT"
                arrCAMPOS(2, 2) = "N"
                arrCAMPOS(2, 3) = "Código"
                arrCAMPOS(2, 4) = "1500"
                arrCAMPOS(2, 5) = "SGI_CODINT"
                
                arrCAMPOS(3, 1) = "SGI_DESCRI"
                arrCAMPOS(3, 2) = "S"
                arrCAMPOS(3, 3) = "Descrição"
                arrCAMPOS(3, 4) = "4000"
                arrCAMPOS(3, 5) = "SGI_DESCRI"
                
                varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strACESSO, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de paradas")
                
                If Len(Trim(varRETORNO)) > 0 Then
                    
                    If .Cell(flexcpText, Row, conCOL_SonMovParada_Action2Do) = dacEnumUpdateAction_Ignore Then
                        If Trim(varRETORNO) <> Trim(.Cell(flexcpText, Row, conCOL_SonMovParada_CodParada)) Then .Cell(flexcpText, Row, conCOL_SonMovParada_Action2Do) = dacEnumUpdateAction_update
                    End If
                    
                    .Cell(flexcpText, Row, conCOL_SonMovParada_CodParada) = Trim(varRETORNO)
                    If PegaCodParada(varRETORNO, Row, 1) = True Then
                        .Cell(flexcpText, Row, conCOL_SonMovParada_CodParada) = Empty
                        .Cell(flexcpText, Row, conCOL_SonMovParada_CodIntParada) = Empty
                        .Cell(flexcpText, Row, conCOL_SonMovParada_DescParada) = Empty
                    End If
                    Exit Sub
                End If
            
        End Select
    
    End With
    
    Exit Sub

Err_grdPARADAS_CellButtonClick:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdPARADAS_CellButtonClick", Me.Name, "CellButtonClick")

End Sub

Private Sub grdPARADAS_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)

On Error GoTo Err_grdPARADAS_KeyPressEdit
     
     With grdPARADAS
          Select Case Col
                    Case conCOL_SonMovParada_CodIntParada
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
          End Select
     End With

     Exit Sub

Err_grdPARADAS_KeyPressEdit:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdPARADAS_KeyPressEdit", Me.Name, "KeyPressEdit")

End Sub

Private Sub grdPARADAS_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

On Error GoTo Err_grdPARADAS_ValidateEdit
     
     Dim intHoras   As Integer
     Dim intMinutos As Integer
     
     With grdPARADAS
          Select Case Col
                 Case conCOL_SonMovParada_CodIntParada
                        If .EditText = Empty Then Exit Sub
                        If Not IsNumeric(.EditText) Then
                            MsgBox "Códgo da Parada inválido !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                        
                        If .Cell(flexcpText, Row, conCOL_SonMovParada_Action2Do) = dacEnumUpdateAction_Ignore Then
                            If Trim(.EditText) <> Trim(.Cell(flexcpText, Row, Col)) Then .Cell(flexcpText, Row, conCOL_SonMovParada_Action2Do) = dacEnumUpdateAction_update
                        End If
                        
                        Cancel = PegaCodParada(Trim(.EditText), Row, 2)
                        If Cancel = True Then Exit Sub
            
            Case conCOL_SonMovParada_HORINI, _
                 conCOL_SonMovParada_HORFIN
                 If .EditText = "  :  " Or _
                    Len(Trim(.EditText)) = 0 Then
                    Cancel = False
                    Exit Sub
                 End If
                 
                 '' ==================================
                 '' Validando Campo Horas
                 If Len(Trim(Mid(.EditText, 1, 2))) = 0 And Len(Trim(Mid(.EditText, 4, 2))) > 0 Then
                    MsgBox "Hora inválida !!!", vbOKOnly + vbExclamation, "Aviso"
                    Cancel = True
                    Exit Sub
                 End If
                 If Len(Trim(Mid(.EditText, 1, 2))) > 0 And Len(Trim(Mid(.EditText, 4, 2))) = 0 Then
                    MsgBox "Hora inválida !!!", vbOKOnly + vbExclamation, "Aviso"
                    Cancel = True
                    Exit Sub
                 End If
                 
                 
                 '' ==================================
                 '' Validando Horas
                 intHoras = CInt(Mid(.EditText, 1, 2))
                 intMinutos = CInt(Mid(.EditText, 4, 2))
                 If intHoras >= 24 Or intHoras < 0 Then
                    MsgBox "Hora Inválida o Dia vai somente até 24:00 !!!", vbOKOnly + vbExclamation, "Aviso"
                    Cancel = True
                    Exit Sub
                 End If
                 If intMinutos >= 60 Then
                    MsgBox "Minutos Inválido os minutos somente devem ser informados de 00 a 59 !!!", vbOKOnly + vbExclamation, "Aviso"
                    Cancel = True
                    Exit Sub
                 End If
          
                If .Cell(flexcpText, Row, conCOL_SonMovParada_Action2Do) = dacEnumUpdateAction_Ignore Then
                    If Trim(.EditText) <> Trim(.Cell(flexcpText, Row, Col)) Then .Cell(flexcpText, Row, conCOL_SonMovParada_Action2Do) = dacEnumUpdateAction_update
                End If
          
          End Select
     End With
    
    Exit Sub

Err_grdPARADAS_ValidateEdit:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdPARADAS_ValidateEdit", Me.Name, "ValidateEdit")

End Sub

Private Sub mskDTLCTO_GotFocus()
    Call objBLBFunc.SelecionaCampos(mskDTLCTO.Name, frmCADAPONTPROD)
End Sub

Private Sub txtCODMAQ_GotFocus()
    Call objBLBFunc.SelecionaCampos(txtCODMAQ.Name, frmCADAPONTPROD)
End Sub

Private Sub txtCODMAQ_KeyPress(KeyAscii As Integer)
    Call objBLBFunc.SoNumeroPonto(KeyAscii, txtCODMAQ.Text)
End Sub

Private Sub txtCODMAQ_Validate(Cancel As Boolean)

    If Len(Trim(txtCODMAQ.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODMAQ.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODMAQ.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRI", "SGI_CADMAQUINA", txtCODMAQ.Text, lblDescMaq, "ATENÇÃO - Máquina Inexistente !!!")
    If Len(Trim(lblDescMaq.Caption)) = 0 Then
       txtCODMAQ.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    If ConfereStatusMaq(txtCODMAQ.Text) = False Then
        lblDescMaq.Caption = ""
        txtCODMAQ.Text = ""
        Cancel = True
        Exit Sub
    End If

End Sub

Private Sub txtCODOPERADOR_GotFocus()
    Call objBLBFunc.SelecionaCampos(txtCODOPERADOR.Name, frmCADAPONTPROD)
End Sub

Private Sub txtCODOPERADOR_KeyPress(KeyAscii As Integer)
    Call objBLBFunc.SoNumeroPonto(KeyAscii, txtCODOPERADOR.Text)
End Sub

Private Sub txtCODOPERADOR_Validate(Cancel As Boolean)

    If Len(Trim(txtCODMAQ.Text)) = 0 Then
        MsgBox "ATENÇÃO - Primeiro informe a máquina !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If

    If Len(Trim(txtCODOPERADOR.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODOPERADOR.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODOPERADOR.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRI", "SGI_CADOPERADOR", Trim(txtCODOPERADOR.Text), lblDescOperador, "ATENÇÃO - Operador Inexistente !!!")
    If Len(Trim(lblDescOperador.Caption)) = 0 Then
       txtCODOPERADOR.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    If ConfereStatusOper(Trim(txtCODOPERADOR.Text), Trim(txtCODMAQ.Text)) = False Then
        lblDescOperador.Caption = ""
        txtCODOPERADOR.Text = ""
        Cancel = True
        Exit Sub
    End If

End Sub

Private Sub txtCODTURNO_GotFocus()
    Call objBLBFunc.SelecionaCampos(txtCODTURNO.Name, frmCADAPONTPROD)
End Sub

Private Sub txtCODTURNO_KeyPress(KeyAscii As Integer)
    Call objBLBFunc.SoNumeroPonto(KeyAscii, txtCODTURNO.Text)
End Sub

Private Sub PegaDescTabelas(strCAMPOPESQ As String, StrCampoRetorno As String, strTABELA As String, strCODIGO As String, lblLabel As Label, strDESCERRO As String)

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
       If Len(Trim(strDESCERRO)) = 0 Then
            MsgBox "Registro Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
       Else
            MsgBox Trim(strDESCERRO), vbOKOnly + vbExclamation, "Aviso"
       End If
    End If
    BREC10.Close
    
End Sub


Private Function ConfereStatusMaq(strCODMAQ As String) As Boolean

    ConfereStatusMaq = True

    sSql = ""

    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADMAQUINA " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & strCODMAQ & vbCrLf
    sSql = sSql & "   And SGI_ATIVA  = 0" '' Ativo
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF() Then
        ConfereStatusMaq = False
        MsgBox "Máquina Desativada impossivel Utilizar !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC.Close
    
End Function

Private Sub txtCODTURNO_Validate(Cancel As Boolean)

    If Len(Trim(txtCODMAQ.Text)) = 0 Then
        MsgBox "ATENÇÃO - Primeiro informe a máquina !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    If Len(Trim(txtCODTURNO.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODTURNO.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODTURNO.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRI", "SGI_CADQTDETURN", txtCODTURNO.Text, lblDescTurno, "ATENÇÃO - Turno Inexistente !!!")
    If Len(Trim(lblDescTurno.Caption)) = 0 Then
       txtCODTURNO.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    If ConfereStatusTurn(Trim(txtCODTURNO.Text), Trim(txtCODMAQ.Text)) = False Then
        lblDescTurno.Caption = ""
        txtCODTURNO.Text = ""
        Cancel = True
        Exit Sub
    End If
    
End Sub

Private Function ConfereStatusOper(strCODOPER As String, strCODMAQ As String) As Boolean

    ConfereStatusOper = True

    sSql = ""

    sSql = "Select" & vbCrLf
    sSql = sSql & "      OPERA.*" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_CADMAQOPER MAQOPE" & vbCrLf
    sSql = sSql & "     ,SGI_CADOPERADOR OPERA" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      MAQOPE.SGI_FILIAL  = " & FILIAL & vbCrLf
    sSql = sSql & "  And MAQOPE.SGI_CODIGO  = " & Trim(strCODMAQ) & vbCrLf
    sSql = sSql & "  And MAQOPE.SGI_CODOPER = " & Trim(strCODOPER) & vbCrLf
    sSql = sSql & "  And MAQOPE.SGI_FILIAL  = OPERA.SGI_FILIAL" & vbCrLf
    sSql = sSql & "  And MAQOPE.SGI_CODOPER = OPERA.SGI_CODIGO" & vbCrLf
    sSql = sSql & "  And OPERA.SGI_ATIVO = 1"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF() Then
        ConfereStatusOper = False
        MsgBox "ATENÇÂO = Operador Desativado ou inexistente impossivel Utilizar !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC.Close
    
End Function

Private Function ConfereStatusTurn(strCODTURN As String, strCODMAQ As String) As Boolean

    ConfereStatusTurn = True

    sSql = ""

    sSql = "Select " & vbCrLf
    sSql = sSql & "       CADTUR.SGI_CODTURN" & vbCrLf
    sSql = sSql & "      ,MAQTUR.SGI_DESCRI" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADMAQTURN  CADTUR" & vbCrLf
    sSql = sSql & "      ,SGI_CADQTDETURN MAQTUR" & vbCrLf

    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       CADTUR.SGI_FILIAL  = " & FILIAL & vbCrLf
    sSql = sSql & "   And CADTUR.SGI_CODIGO  = " & Trim(strCODMAQ) & vbCrLf
    sSql = sSql & "   And CADTUR.SGI_CODTURN = " & Trim(strCODTURN) & vbCrLf
    
    sSql = sSql & "   And CADTUR.SGI_FILIAL  = MAQTUR.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And CADTUR.SGI_CODTURN = MAQTUR.SGI_CODIGO" & vbCrLf
    sSql = sSql & "   And MAQTUR.SGI_ATIVO = 1"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF() Then
        ConfereStatusTurn = False
        MsgBox "ATENÇÂO = Turno Desativado ou inexistente impossivel Utilizar !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC.Close
    
End Function


Private Sub InitGridMov()

    With grdAPONT
    
       .Cols = conColumnsIn_SonMov
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonMov_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonMov_CodOP) = ""
       .ColDataType(conCOL_SonMov_CodOP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonMov_Capacidade) = ""
       .ColDataType(conCOL_SonMov_Capacidade) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMov_CodProd) = ""
       .ColDataType(conCOL_SonMov_CodProd) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMov_DescRot) = ""
       .ColDataType(conCOL_SonMov_DescRot) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMov_CodTipo) = ""
       .ColDataType(conCOL_SonMov_CodTipo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonMov_PesqTipo) = ""
       .ColDataType(conCOL_SonMov_PesqTipo) = flexDTString
       .ColComboList(conCOL_SonMov_PesqTipo) = "..."
       
       .Cell(flexcpData, 0, conCOL_SonMov_DescTipo) = ""
       .ColDataType(conCOL_SonMov_DescTipo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMov_AplicLote) = ""
       .ColDataType(conCOL_SonMov_AplicLote) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMov_Lote) = ""
       .ColDataType(conCOL_SonMov_Lote) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMov_QtdeEntr) = ""
       .ColDataType(conCOL_SonMov_QtdeEntr) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonMov_QtdeSaida) = ""
       .ColDataType(conCOL_SonMov_QtdeSaida) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonMov_Retrabalho) = ""
       .ColDataType(conCOL_SonMov_Retrabalho) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonMov_IDPRODUTO) = ""
       .ColDataType(conCOL_SonMov_IDPRODUTO) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonMov_INDICE) = ""
       .ColDataType(conCOL_SonMov_INDICE) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonMov_Action2Do) = ""
       .ColDataType(conCOL_SonMov_Action2Do) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonMov_HORINI) = ""
       .ColDataType(conCOL_SonMov_HORINI) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMov_HORFIN) = ""
       .ColDataType(conCOL_SonMov_HORFIN) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMov_TotalLiq) = ""
       .ColDataType(conCOL_SonMov_TotalLiq) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMov_IDINTERNO) = ""
       .ColDataType(conCOL_SonMov_IDINTERNO) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonMov_INDICEBKP) = ""
       .ColDataType(conCOL_SonMov_INDICEBKP) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMov_FILIALPED) = ""
       .ColDataType(conCOL_SonMov_FILIALPED) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonMov_DESCFILIALPED) = ""
       .ColDataType(conCOL_SonMov_DESCFILIALPED) = flexDTString
       
       .ColWidth(conCOL_SonMov_CodOP) = 1200
       .ColWidth(conCOL_SonMov_Capacidade) = 2000
       .ColWidth(conCOL_SonMov_CodProd) = 1200
       .ColWidth(conCOL_SonMov_DescRot) = 5000
       .ColWidth(conCOL_SonMov_CodTipo) = 1000
       .ColWidth(conCOL_SonMov_PesqTipo) = 300
       .ColWidth(conCOL_SonMov_DescTipo) = 2000
       .ColWidth(conCOL_SonMov_AplicLote) = 1200
       .ColWidth(conCOL_SonMov_Lote) = 900
       .ColWidth(conCOL_SonMov_QtdeEntr) = 1100
       .ColWidth(conCOL_SonMov_QtdeSaida) = 900
       .ColWidth(conCOL_SonMov_Retrabalho) = 900
       
       .ColWidth(conCOL_SonMov_IDPRODUTO) = 0
       .ColWidth(conCOL_SonMov_INDICE) = 0
       .ColWidth(conCOL_SonMov_Action2Do) = 0
       
       .ColWidth(conCOL_SonMov_HORINI) = 800
       .ColWidth(conCOL_SonMov_HORFIN) = 800
       .ColWidth(conCOL_SonMov_TotalLiq) = 800
       .ColWidth(conCOL_SonMov_IDINTERNO) = 0
       .ColWidth(conCOL_SonMov_INDICEBKP) = 0
       
       .ColWidth(conCOL_SonMov_FILIALPED) = 0
       .ColWidth(conCOL_SonMov_DESCFILIALPED) = 1000
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
    End With
    
End Sub


Private Sub IncRegGrid()
   
    If cTipOper = "C" Then Exit Sub
    
    If Len(Trim(txtCODMAQ.Text)) = 0 Then
        MsgBox "ATENÇÃO - A Máquina deve ser informada !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCODMAQ.SetFocus
        Exit Sub
    End If
    If Len(Trim(txtCODTURNO.Text)) = 0 Then
        MsgBox "ATENÇÃO - O Turno deve ser informado !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCODTURNO.SetFocus
        Exit Sub
    End If
    If Len(Trim(txtCODOPERADOR.Text)) = 0 Then
        MsgBox "ATENÇÃO - O Operador deve ser informado !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCODOPERADOR.SetFocus
        Exit Sub
    End If
    
    If objBLBFunc.FcExisteLinhaVazia(grdAPONT, conCOL_SonMov_CodOP) = False Then Exit Sub
    
    With grdAPONT
    
        .AddItem "" & vbTab & _
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
                 dacEnumUpdateAction_Insert & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 ""
    
    End With
    
End Sub

Private Sub InitGridParada()

    With grdPARADAS
    
       .Cols = conColumnsIn_SonMovParada
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonMovParada_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonMovParada_Indice) = ""
       .ColDataType(conCOL_SonMovParada_Indice) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMovParada_CodParada) = ""
       .ColDataType(conCOL_SonMovParada_CodParada) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonMovParada_CodIntParada) = ""
       .ColDataType(conCOL_SonMovParada_CodIntParada) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMovParada_PesqParada) = ""
       .ColDataType(conCOL_SonMovParada_PesqParada) = flexDTString
       .ColComboList(conCOL_SonMovParada_PesqParada) = "..."
       
       .Cell(flexcpData, 0, conCOL_SonMovParada_DescParada) = ""
       .ColDataType(conCOL_SonMovParada_DescParada) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMovParada_Action2Do) = ""
       .ColDataType(conCOL_SonMovParada_Action2Do) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonMovParada_IDINTERNO) = ""
       .ColDataType(conCOL_SonMovParada_IDINTERNO) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonMovParada_HORINI) = ""
       .ColDataType(conCOL_SonMovParada_HORINI) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMovParada_HORFIN) = ""
       .ColDataType(conCOL_SonMovParada_HORFIN) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMovParada_TotalLiq) = ""
       .ColDataType(conCOL_SonMovParada_TotalLiq) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonMovParada_IDPAI) = ""
       .ColDataType(conCOL_SonMovParada_IDPAI) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonMovParada_FILIALPED) = ""
       .ColDataType(conCOL_SonMovParada_FILIALPED) = flexDTLong
       
       .ColWidth(conCOL_SonMovParada_CodIntParada) = 1000
       .ColWidth(conCOL_SonMovParada_PesqParada) = 300
       .ColWidth(conCOL_SonMovParada_DescParada) = 4000
       .ColWidth(conCOL_SonMovParada_HORINI) = 1200
       .ColWidth(conCOL_SonMovParada_HORFIN) = 1200
       .ColWidth(conCOL_SonMovParada_TotalLiq) = 1200
       
       .ColWidth(conCOL_SonMovParada_Indice) = 0
       .ColWidth(conCOL_SonMovParada_CodParada) = 0
       .ColWidth(conCOL_SonMovParada_Action2Do) = 0
       .ColWidth(conCOL_SonMovParada_IDINTERNO) = 0
       .ColWidth(conCOL_SonMovParada_IDPAI) = 0
       .ColWidth(conCOL_SonMovParada_FILIALPED) = 0
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
    End With
    
End Sub

Private Sub IncRegGridParada()
   
    If cTipOper = "C" Then Exit Sub
    
    With grdAPONT
        If (.Rows - 1) = 0 Then
            MsgBox "ATENÇÃO - Não foi informado nenhum apontamento !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
        End If
        If (.Row) = 0 Then
            MsgBox "ATENÇÃO - Não foi selecionado nenhum registro no apontamento !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
        End If
        If Len(Trim(.Cell(flexcpText, .Row, conCOL_SonMov_CodOP))) = 0 Then
            MsgBox "Informe o Código da OP !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
        End If
        If Len(Trim(.Cell(flexcpText, .Row, conCOL_SonMov_CodTipo))) = 0 Then
            MsgBox "Informe o Tipo da Atividade !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
        End If
    End With
    
    If objBLBFunc.FcExisteLinhaVazia(grdPARADAS, conCOL_SonMovParada_CodParada) = False Then Exit Sub
    
    With grdPARADAS
        .AddItem grdAPONT.Cell(flexcpText, grdAPONT.Row, conCOL_SonMov_INDICE) & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 dacEnumUpdateAction_Insert & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 grdAPONT.Cell(flexcpText, grdAPONT.Row, conCOL_SonMov_IDINTERNO) & vbTab & _
                 grdAPONT.Cell(flexcpText, grdAPONT.Row, conCOL_SonMov_FILIALPED)
    End With
    
End Sub

Private Function PegaOPNOVALATA(strCODOP As String, lngLINHA As Long) As Boolean

    PegaOPNOVALATA = False
    
    If Len(Trim(strCODOP)) = 0 Then Exit Function
    
    Dim strBCOFILIAL As String
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       OP.* " & vbCrLf
    sSql = sSql & "      ,PROD.*" & vbCrLf
    sSql = sSql & "      ,LPROD.*" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD OP" & vbCrLf
    sSql = sSql & "     , SGI_CADPRODUTO      PROD" & vbCrLf
    sSql = sSql & "     , SGI_CADLINHAPRODUTO LPROD" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       OP.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And OP.SGI_CODIGO = " & strCODOP & vbCrLf
    
    sSql = sSql & "   And PROD.SGI_FILIAL    = OP.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PROD.SGI_IDPRODUTO = OP.SGI_IDPRODUTO " & vbCrLf
     
    sSql = sSql & "   And LPROD.SGI_FILIAL   = PROD.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And LPROD.SGI_CODLIN   = PROD.SGI_CODLINPROD"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC.EOF() Then
        If BREC!SGI_STATUS = 2 Then
            MsgBox "ATENÇÂO - Estam OP Já foi finalizada, Faturada !!!", vbOKOnly + vbExclamation, "Aviso"
        Else
            With grdAPONT
                  .Cell(flexcpText, lngLINHA, conCOL_SonMov_Capacidade) = BREC!SGI_DESCRI
                  .Cell(flexcpText, lngLINHA, conCOL_SonMov_CodProd) = BREC!SGI_CODPROD
                  .Cell(flexcpText, lngLINHA, conCOL_SonMov_DescRot) = Trim(BREC!SGI_DESCRICAO)
                  .Cell(flexcpText, lngLINHA, conCOL_SonMov_IDPRODUTO) = BREC!SGI_IDPRODUTO
                  .Cell(flexcpText, lngLINHA, conCOL_SonMov_FILIALPED) = BREC!SGI_FILIALPED
                  .Cell(flexcpText, lngLINHA, conCOL_SonMov_DESCFILIALPED) = NomeFilial(BREC!SGI_FILIALPED)
            End With
            PegaOPNOVALATA = True
        End If
    End If
    BREC.Close

End Function

Private Sub IqualaTipoApont(strCODTIPO As String, lngLINHA As Long)

    If Len(Trim(strCODTIPO)) = 0 Then Exit Sub
    
    Dim strINDICE As String
    
    sSql = ""

    sSql = "Select" & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADTIPAPONT " & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & strCODTIPO & vbCrLf

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC.EOF() Then
    
        If BREC!SGI_ATIVO = 0 Then
            MsgBox "O Tipo de Apontamento não Esta Ativo !!!", vbOKOnly + vbExclamation, "Aviso"
        Else
        
            With grdAPONT
        
                .Cell(flexcpText, lngLINHA, conCOL_SonMov_CodTipo) = BREC!SGI_CODIGO
                .Cell(flexcpText, lngLINHA, conCOL_SonMov_DescTipo) = BREC!SGI_DESCRI
        
            End With
        End If
    Else
        MsgBox "ATENÇÂO - Tipo de Apontamento Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC.Close

End Sub

Private Sub CriaChave(strCODOP As String, strCODIDPROD As String, strCODTIPO As String, lngROW As Long)

    Dim strINDICE As String
    
    strINDICE = Trim(strCODOP & strCODIDPROD & strCODTIPO)

    With grdAPONT
            .Cell(flexcpText, lngROW, conCOL_SonMov_INDICE) = strINDICE
    End With
    
End Sub

Private Sub MostraDados()
    
On Error GoTo Err_MostraDados
    
    With grdAPONT
        If (.Rows - 1) > 0 And .Row > 0 Then
            Dim strINDICE As String
            strINDICE = ""
            If Len(Trim(.Cell(flexcpText, .Row, conCOL_SonMov_INDICE))) > 0 Then strINDICE = Trim(.Cell(flexcpText, .Row, conCOL_SonMov_INDICE))
            Call objBLBFunc.CarregaDadosGrdFilho(grdPARADAS, conCOL_SonMovParada_Action2Do, conCOL_SonMovParada_Indice, strINDICE)
        End If
    End With
    
    Exit Sub
    
Err_MostraDados:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : MostraDados()", Me.Name, "MostraDados()", strCAMARQERRO)
    
End Sub

Private Function PegaCodParada(strCODPARADA As String, lngROW As Long, intTipo As Long) As Boolean

    PegaCodParada = True
    
    If Len(Trim(strCODPARADA)) = 0 Then Exit Function
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPARADAS" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    
    If intTipo = 1 Then sSql = sSql & "   And SGI_CODIGO = " & strCODPARADA & vbCrLf
    If intTipo = 2 Then sSql = sSql & "   And SGI_CODINT = " & strCODPARADA & vbCrLf
    
    sSql = sSql & "   And SGI_ATIVO  = 1"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        
        With grdPARADAS
            .Cell(flexcpText, lngROW, conCOL_SonMovParada_CodParada) = BREC!SGI_CODIGO
            .Cell(flexcpText, lngROW, conCOL_SonMovParada_CodIntParada) = BREC!SGI_CODINT
            .Cell(flexcpText, lngROW, conCOL_SonMovParada_DescParada) = BREC!SGI_DESCRI
        End With
    
        PegaCodParada = False
    Else
        MsgBox "Este Código de Parada não existe !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC.Close

End Function
 

Private Function Valida_Campos() As Boolean

On Error GoTo Err_Valida_Campos
     
     Dim lngLINHA   As Long
     Dim I          As Long
     
     Valida_Campos = False
     
     If Len(Trim(txtCODMAQ.Text)) = 0 Then
        MsgBox "ATENÇÂO - Informe o código da máquina !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCODMAQ.SetFocus
        Exit Function
     End If
     If Len(Trim(txtCODOPERADOR.Text)) = 0 Then
        MsgBox "ATENÇÂO - Informe o código do Operador !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCODOPERADOR.SetFocus
        Exit Function
     End If
     If Len(Trim(txtCODTURNO.Text)) = 0 Then
        MsgBox "ATENÇÂO - Informe o código do turno !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCODTURNO.SetFocus
        Exit Function
     End If
     
     lngLINHA = 0
     For I = 1 To (grdAPONT.Rows - 1)
        If grdAPONT.Cell(flexcpText, I, conCOL_SonMov_Action2Do) <> dacEnumUpdateAction_delete Then
            lngLINHA = lngLINHA + 1
            If Len(Trim(grdAPONT.Cell(flexcpText, I, conCOL_SonMov_AplicLote))) > 20 Or _
               Len(Trim(grdAPONT.Cell(flexcpText, I, conCOL_SonMov_Lote))) > 20 Then
                MsgBox "ATENÇÃO - Somente é permitido 20 digitos !!!", vbOKOnly + vbExclamation, "Aviso"
                Exit Function
            End If
        End If
     Next I
     If lngLINHA = 0 Then
        MsgBox "ATENÇÃO - Informe pelo menos 1 Apontamento !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Function
     End If
     Valida_Campos = True

    Exit Function
    
Err_Valida_Campos:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Valida_Campos()", Me.Name, "Valida_Campos()", strCAMARQERRO)

End Function


Private Sub CarregaCampos()
    If objCADAPONTPROD.Carrega_Campos(strNOMFILIAL) = True Then
    
        lblCODIGO.Caption = objCADAPONTPROD.Codigo
        txtCODMAQ.Text = Trim(objCADAPONTPROD.CODMAQ)
        txtCODTURNO.Text = Trim(objCADAPONTPROD.CODTURN)
        txtCODOPERADOR.Text = objCADAPONTPROD.CODOPER
        mskDTLCTO.Text = objCADAPONTPROD.DTLANCT
    
        If Len(Trim(txtCODMAQ.Text)) > 0 Then Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRI", "SGI_CADMAQUINA", txtCODMAQ.Text, lblDescMaq, "ATENÇÃO - Máquina Inexistente !!!")
        If Len(Trim(txtCODTURNO.Text)) > 0 Then Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRI", "SGI_CADQTDETURN", txtCODTURNO.Text, lblDescTurno, "ATENÇÃO - Turno Inexistente !!!")
        If Len(Trim(txtCODOPERADOR.Text)) > 0 Then Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRI", "SGI_CADOPERADOR", Trim(txtCODOPERADOR.Text), lblDescOperador, "ATENÇÃO - Operador Inexistente !!!")
    
        arrAPONT = objCADAPONTPROD.APONT
        arrPARADAS = objCADAPONTPROD.PARADAS
        
        Call PopGrdApont
        Call PopGrdParadas
    
        If (grdAPONT.Rows - 1) > 0 Then grdAPONT.Row = 1
        
    End If
End Sub


Private Sub PopGrdApont()
    Dim I As Long
    If IsArray(arrAPONT) Then
        With grdAPONT
            For I = 1 To UBound(arrAPONT)
                .AddItem arrAPONT(I, 1) & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         arrAPONT(I, 2) & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         arrAPONT(I, 3) & vbTab & _
                         arrAPONT(I, 4) & vbTab & _
                         arrAPONT(I, 5) & vbTab & _
                         arrAPONT(I, 6) & vbTab & _
                         arrAPONT(I, 7) & vbTab & _
                         "" & vbTab & _
                         arrAPONT(I, 12) & vbTab & _
                         dacEnumUpdateAction_Ignore & vbTab & _
                         arrAPONT(I, 8) & vbTab & _
                         arrAPONT(I, 9) & vbTab & _
                         arrAPONT(I, 10) & vbTab & _
                         arrAPONT(I, 11) & vbTab & _
                         arrAPONT(I, 12) & vbTab & _
                         arrAPONT(I, 13) & vbTab & _
                         ""
                         
                Call PegaOPNOVALATA(Str(arrAPONT(I, 1)), (.Rows - 1))
                Call PegaOPSTEEL(Str(arrAPONT(I, 1)), (.Rows - 1))
                Call IqualaTipoApont(Str(arrAPONT(I, 2)), (.Rows - 1))
            
            Next I
        End With
    End If
End Sub

Private Sub PopGrdParadas()
    Dim I As Long
    If IsArray(arrPARADAS) Then
        With grdPARADAS
            For I = 1 To UBound(arrPARADAS)
                .AddItem arrPARADAS(I, 1) & vbTab & _
                         arrPARADAS(I, 2) & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         dacEnumUpdateAction_Ignore & vbTab & _
                         arrPARADAS(I, 3) & vbTab & _
                         arrPARADAS(I, 4) & vbTab & _
                         arrPARADAS(I, 5) & vbTab & _
                         arrPARADAS(I, 6) & vbTab & _
                         arrPARADAS(I, 7) & vbTab & _
                         arrPARADAS(I, 8)
            
                Call PegaCodParada(Trim(Str(arrPARADAS(I, 2))), (.Rows - 1), 1)
            Next I
        End With
    End If
End Sub

Private Sub TrocaIndice(strINDICEBKP As String, strINDICENOVO As String, lngROW As Long)

    Dim I As Integer
    With grdPARADAS
        For I = 1 To (.Rows - 1)
            If .Cell(flexcpText, I, conCOL_SonMovParada_Indice) = strINDICEBKP Then
               .Cell(flexcpText, I, conCOL_SonMovParada_Indice) = strINDICENOVO
               .Cell(flexcpText, I, conCOL_SonMovParada_Action2Do) = dacEnumUpdateAction_update
            End If
        Next I
    End With
    
    grdAPONT.Cell(flexcpText, lngROW, conCOL_SonMov_INDICEBKP) = strINDICENOVO
    
End Sub

Private Function PegaOPSTEEL(strCODOP As String, lngLINHA As Long) As Boolean

    PegaOPSTEEL = False
    
    If Len(Trim(strCODOP)) = 0 Then Exit Function
    
    Dim strBCOFILIAL As String
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       OP.* " & vbCrLf
    sSql = sSql & "      ,PROD.*" & vbCrLf
    sSql = sSql & "      ,LPROD.*" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD_STEEL OP" & vbCrLf
    sSql = sSql & "     , SGI_CADPRODUTO      PROD" & vbCrLf
    sSql = sSql & "     , SGI_CADLINHAPRODUTO LPROD" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       OP.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And OP.SGI_CODIGO = " & strCODOP & vbCrLf
    
    sSql = sSql & "   And PROD.SGI_FILIAL    = OP.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PROD.SGI_IDPRODUTO = OP.SGI_IDPRODUTO " & vbCrLf
     
    sSql = sSql & "   And LPROD.SGI_FILIAL   = PROD.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And LPROD.SGI_CODLIN   = PROD.SGI_CODLINPROD"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC.EOF() Then
        If BREC!SGI_STATUS = 2 Then
            MsgBox "ATENÇÂO - Estam OP Já foi finalizada, Faturada !!!", vbOKOnly + vbExclamation, "Aviso"
        Else
            With grdAPONT
                  .Cell(flexcpText, lngLINHA, conCOL_SonMov_Capacidade) = BREC!SGI_DESCRI
                  .Cell(flexcpText, lngLINHA, conCOL_SonMov_CodProd) = BREC!SGI_CODPROD
                  .Cell(flexcpText, lngLINHA, conCOL_SonMov_DescRot) = Trim(BREC!SGI_DESCRICAO)
                  .Cell(flexcpText, lngLINHA, conCOL_SonMov_IDPRODUTO) = BREC!SGI_IDPRODUTO
                  .Cell(flexcpText, lngLINHA, conCOL_SonMov_FILIALPED) = BREC!SGI_FILIALPED
                  .Cell(flexcpText, lngLINHA, conCOL_SonMov_DESCFILIALPED) = NomeFilial(BREC!SGI_FILIALPED)
            End With
            PegaOPSTEEL = True
        End If
    End If
    BREC.Close

End Function


Private Function NomeFilial(intCODFILIALPED As Integer) As String
    NomeFilial = ""
    If intCODFILIALPED = 0 Then NomeFilial = "NOVALATA"
    If intCODFILIALPED = 1 Then NomeFilial = "STEEL ROL"
End Function
