VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCADPLAMESTRE 
   Caption         =   "Cadastro de plano mestre de produção"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7665
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   7665
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab StPlano 
      Height          =   6735
      Left            =   0
      TabIndex        =   12
      Top             =   960
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   11880
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Plano Mestre"
      TabPicture(0)   =   "frmCADPLAMESTRE.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "mtvSemana"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.Frame Frame6 
         Caption         =   "[ DIas da Semana ]"
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
         Height          =   2535
         Left            =   120
         TabIndex        =   29
         Top             =   4080
         Width           =   6855
         Begin VSFlex8LCtl.VSFlexGrid grdDIAS 
            Height          =   2175
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   6615
            _cx             =   11668
            _cy             =   3836
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
      Begin VB.Frame Frame4 
         Caption         =   "[ Semamas ]"
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
         Height          =   2295
         Left            =   120
         TabIndex        =   27
         Top             =   1800
         Width           =   6855
         Begin VSFlex8LCtl.VSFlexGrid grdDIASSEM 
            Height          =   1935
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   6615
            _cx             =   11668
            _cy             =   3413
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
      Begin VB.CommandButton Command4 
         Height          =   300
         Left            =   7080
         Picture         =   "frmCADPLAMESTRE.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Exclui a linha da Gride Selecionada"
         Top             =   2280
         Width           =   300
      End
      Begin VB.CommandButton Command3 
         Height          =   300
         Left            =   7080
         Picture         =   "frmCADPLAMESTRE.frx":0166
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Inclui uma nova linha na Gride"
         Top             =   1920
         Width           =   300
      End
      Begin MSComCtl2.MonthView mtvSemana 
         Height          =   2370
         Left            =   4920
         TabIndex        =   21
         Top             =   3120
         Visible         =   0   'False
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         Enabled         =   0   'False
         StartOfWeek     =   151519233
         CurrentDate     =   41307
      End
      Begin VB.Frame Frame2 
         Height          =   1455
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   7335
         Begin VB.Frame Frame3 
            Caption         =   "[ Ativo ]"
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
            Height          =   495
            Left            =   5400
            TabIndex        =   24
            Top             =   840
            Width           =   1815
            Begin VB.OptionButton optATIVOSN 
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
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   26
               Top             =   240
               Width           =   735
            End
            Begin VB.OptionButton optATIVOSN 
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
               Height          =   195
               Index           =   0
               Left            =   960
               TabIndex        =   25
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame5 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   5040
            TabIndex        =   19
            Top             =   195
            Visible         =   0   'False
            Width           =   1815
            Begin VB.OptionButton optSimNao 
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
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   5
               Top             =   0
               Width           =   735
            End
            Begin VB.OptionButton optSimNao 
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
               Height          =   195
               Index           =   1
               Left            =   1080
               TabIndex        =   6
               Top             =   0
               Width           =   735
            End
         End
         Begin VB.TextBox txtQTDE 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4200
            TabIndex        =   7
            Text            =   "txtQTDE"
            Top             =   960
            Width           =   1095
         End
         Begin VB.ComboBox cboAno 
            Height          =   315
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   960
            Width           =   855
         End
         Begin VB.ComboBox cboMes 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtCodProd 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   960
            MaxLength       =   10
            TabIndex        =   1
            Text            =   "txtCodProd"
            Top             =   600
            Width           =   1095
         End
         Begin VB.CommandButton Command2 
            Height          =   315
            Left            =   2040
            Picture         =   "frmCADPLAMESTRE.frx":02B0
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtCODIGO 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   960
            TabIndex        =   0
            Text            =   "txtCODIGO"
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblDescLinha 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDescLinha"
            Height          =   285
            Left            =   2400
            TabIndex        =   20
            Top             =   600
            Width           =   4575
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Considera Pedidos"
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
            Left            =   3360
            TabIndex        =   18
            Top             =   240
            Visible         =   0   'False
            Width           =   1590
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade:"
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
            Left            =   3120
            TabIndex        =   17
            Top             =   960
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Mês/Ano:"
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
            Left            =   120
            TabIndex        =   16
            Top             =   960
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Linha"
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
            Left            =   120
            TabIndex        =   15
            Top             =   600
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
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   660
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7575
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
         Picture         =   "frmCADPLAMESTRE.frx":03B2
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
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
         Picture         =   "frmCADPLAMESTRE.frx":04B4
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
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
         Picture         =   "frmCADPLAMESTRE.frx":05B6
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmCADPLAMESTRE"
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
Public lngCodUsuario    As Long
Public intFILIALTAB     As Integer

Dim objBLBFunc          As Object
Dim objCADPLAMESTRE     As Object
Dim objPESQPADRAO       As Object
Dim arrITENSPM          As Variant
Dim arrITENSDIAS        As Variant
Dim dtDIAMESANO         As Date
Dim strTABELAFILIAL     As String
Dim strNOMEFILIAL       As String

Const conCOL_PmS_Semana                             As Integer = 0
Const conCOL_PmS_Qtde                               As Integer = 1
Const conCOL_PmS_IDINTERNO                          As Integer = 2
Const conCOL_PmS_Ativo                              As Integer = 3
Const conCOL_PmS_FormatString                       As String = "=Semana|Qtde.Latas|IDINTERNO|Ativo"
Const conColumnsIn_PmS                              As Integer = 4

Const conCOL_PmSDIAS_SemanaIndice                   As Integer = 0
Const conCOL_PmSDIAS_DATA                           As Integer = 1
Const conCOL_PmSDIAS_DIASEMANA                      As Integer = 2
Const conCOL_PmSDIAS_DIA                            As Integer = 3
Const conCOL_PmSDIAS_QTDE                           As Integer = 4
Const conCOL_PmSDIAS_IDINTERNO                      As Integer = 5
Const conCOL_PmSDIAS_Ativo                          As Integer = 6
Const conCOL_PmSDIAS_FormatString                   As String = "=SemIndice|Data|Dia/Sem|Dia|Qtde.|IDINTERNO|Ativo"
Const conColumnsIn_PmSDIAS                          As Integer = 7

Private Sub cmdAltera_Click()
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    
    Me.Caption = "Cadastro de plano mestre de produção - [ ALTERAÇÃO ]"
    
    cTipOper = "A"
    
    txtCodProd.Enabled = False
    Command2.Enabled = False
    cboMes.Enabled = False
    cboAno.Enabled = False

End Sub

Private Sub CmdSalva_Click()

    Dim I As Integer
    
    If ValidaCampos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADPLAMESTRE.CODIGO = objCADPLAMESTRE.Gera_Codigo(Me.Name & strTABELAFILIAL)
    
    objCADPLAMESTRE.CODLINHA = CLng(txtCodProd.Text)
    objCADPLAMESTRE.MES = cboMes.ItemData(cboMes.ListIndex)
    objCADPLAMESTRE.ANO = cboAno.ItemData(cboAno.ListIndex)
    If optSimNao.Item(0).Value = True Then objCADPLAMESTRE.optSimNao = 0
    If optSimNao.Item(1).Value = True Then objCADPLAMESTRE.optSimNao = 1
    objCADPLAMESTRE.QTDE = CCur(txtQTDE.Text)
    
    If optATIVOSN(0).Value = True Then objCADPLAMESTRE.ATIVO = 0
    If optATIVOSN(1).Value = True Then objCADPLAMESTRE.ATIVO = 1
    
    '' Semana
    arrITENSPM = Empty
    With grdDIASSEM
        If (.Rows - 1) > 0 Then
            ReDim arrITENSPM(1 To (.Rows - 1), 1 To 4) As String
            For I = 1 To (.Rows - 1)
                arrITENSPM(I, 1) = .Cell(flexcpText, I, conCOL_PmS_Semana)
                
                arrITENSPM(I, 2) = "Null"
                If Len(Trim(.Cell(flexcpText, I, conCOL_PmS_Qtde))) > 0 Then arrITENSPM(I, 2) = .Cell(flexcpText, I, conCOL_PmS_Qtde)
                
                If cTipOper = "I" Then
                    arrITENSPM(I, 3) = objCADPLAMESTRE.Gera_Codigo(Me.Name & "_PLMESTRE" & strTABELAFILIAL)
                ElseIf cTipOper = "A" Then
                    If Len(Trim(.Cell(flexcpText, I, conCOL_PmS_IDINTERNO))) = 0 Then
                        arrITENSPM(I, 3) = objCADPLAMESTRE.Gera_Codigo(Me.Name & "_PLMESTRE" & strTABELAFILIAL)
                    Else
                        arrITENSPM(I, 3) = .Cell(flexcpText, I, conCOL_PmS_IDINTERNO)
                    End If
                End If
            
                arrITENSPM(I, 4) = .Cell(flexcpText, I, conCOL_PmS_Ativo)
            Next I
        End If
    End With
    objCADPLAMESTRE.ITENSDIAS = arrITENSPM
    
    '' Dia da Semana
    arrITENSDIAS = Empty
    With grdDIAS
        If (.Rows - 1) Then
            ReDim arrITENSDIAS(1 To (.Rows - 1), 1 To 6) As String
            For I = 1 To (.Rows - 1)
                arrITENSDIAS(I, 1) = .Cell(flexcpText, I, conCOL_PmSDIAS_SemanaIndice)
                arrITENSDIAS(I, 2) = "'" & Format(CDate(.Cell(flexcpText, I, conCOL_PmSDIAS_DATA)), "MM/DD/YYYY") & "'"
                arrITENSDIAS(I, 3) = .Cell(flexcpText, I, conCOL_PmSDIAS_DIASEMANA)
                
                arrITENSDIAS(I, 4) = "Null"
                If Len(Trim(.Cell(flexcpText, I, conCOL_PmSDIAS_QTDE))) > 0 Then arrITENSDIAS(I, 4) = .Cell(flexcpText, I, conCOL_PmSDIAS_QTDE)
            
                If cTipOper = "I" Then
                    arrITENSDIAS(I, 5) = objCADPLAMESTRE.Gera_Codigo(Me.Name & "_PLMESTRE_DS" & strTABELAFILIAL)
                ElseIf cTipOper = "A" Then
                    If Len(Trim(.Cell(flexcpText, I, conCOL_PmSDIAS_IDINTERNO))) = 0 Then
                        arrITENSDIAS(I, 5) = objCADPLAMESTRE.Gera_Codigo(Me.Name & "_PLMESTRE_DS" & strTABELAFILIAL)
                    Else
                        arrITENSDIAS(I, 5) = .Cell(flexcpText, I, conCOL_PmSDIAS_IDINTERNO)
                    End If
                End If
            
                arrITENSDIAS(I, 6) = .Cell(flexcpText, I, conCOL_PmSDIAS_Ativo)
            Next I
        End If
    End With
    objCADPLAMESTRE.ITENSDIASSEM = arrITENSDIAS
    
    If objCADPLAMESTRE.GRAVA(cTipOper, strTABELAFILIAL) = False Then Exit Sub
          
    MsgBox "O Plano mestre foi " & IIf(cTipOper = "I", "incluso", IIf(cTipOper = "A", "alterado", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
    If cTipOper = "I" Then Unload Me
    
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()


End Sub

Private Sub Command2_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADGRUPLINHA" & strTABELAFILIAL & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_ATIVO  = 1" & vbCrLf
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "ID"
    arrCAMPOS(1, 4) = "800"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRI"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição da Linha"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "SGI_DESCRI"
    
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Linha de Produtos")
    
    If Len(Trim(varRETORNO)) > 0 Then
        txtCodProd.Text = varRETORNO
        lblDescLinha.Caption = PegaDescrLinha(varRETORNO, "SGI_CODIGO")
    End If
    txtCodProd.SetFocus

End Sub


Private Sub Command3_Click()
    If cTipOper = "I" Or cTipOper = "A" Then Call IncRegGridCotas
End Sub

Private Sub Command4_Click()
    If cTipOper = "I" Or cTipOper = "A" Then Call objBLBFunc.ExclLinhaGrid(grdDIASSEM, grdDIASSEM.Row)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objCADPLAMESTRE = CreateObject("CADPLAMESTRE.clsCADPLAMESTRE")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    
    objCADPLAMESTRE.FILIAL = FILIAL
   
    strTABELAFILIAL = ""
    strNOMEFILIAL = "NOVALATA"
    If intFILIALTAB = 1 Then
        strTABELAFILIAL = "_STEEL"
        strNOMEFILIAL = "STEEL"
    End If
    
    If cTipOper = "I" Then Inclui
    If cTipOper = "A" Then Altera
    If cTipOper = "C" Then Consulta

    Me.Caption = Me.Caption & " / " & strNOMEFILIAL

End Sub

Private Sub Inclui()

    Dim I As Integer
    
    StPlano.Tab = 0
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de plano mestre de produção - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADPLAMESTRE
    
    objCADPLAMESTRE.PreenchComboMes cboMes
    objCADPLAMESTRE.PreenchComboAno cboAno
    
    '' Mês Vigente
    For I = 0 To (cboMes.ListCount - 1)
        If cboMes.ItemData(I) = Month(Date) Then cboMes.ListIndex = I
    Next I
    
    '' Ano Vigente
    For I = 0 To (cboAno.ListCount - 1)
        If cboAno.ItemData(I) = Year(Date) Then cboAno.ListIndex = I
    Next I
    
    optSimNao.Item(1).Value = True
    
    txtCodProd.Enabled = True
    Command2.Enabled = True
    cboMes.Enabled = True
    cboAno.Enabled = True
    
    StPlano.Tab = 0
    LimpaCamposLabel
    optATIVOSN(1).Value = True
    
    Call ConfGridPMS
    Call ConfGridPMSDIAS

    mtvSemana.Value = Now

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub


Private Sub grdDIAS_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     With grdDIAS
          Select Case Col
                 Case conCOL_PmSDIAS_SemanaIndice
                 Case conCOL_PmSDIAS_Ativo
                     Call RecalcCotas(CLng(.Cell(flexcpText, Row, conCOL_PmSDIAS_SemanaIndice)))
          End Select
     End With
End Sub

Private Sub grdDIAS_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_PmSDIAS_SemanaIndice, _
         conCOL_PmSDIAS_DATA, _
         conCOL_PmSDIAS_DIASEMANA, _
         conCOL_PmSDIAS_DIA, _
         conCOL_PmSDIAS_IDINTERNO
         Cancel = True
    Case conCOL_PmSDIAS_QTDE, _
         conCOL_PmSDIAS_Ativo
         If cTipOper = "C" Then
            Cancel = True
            Exit Sub
         End If
         If grdDIASSEM.Cell(flexcpText, grdDIASSEM.Row, conCOL_PmS_Ativo) = 0 Then Cancel = True
    Case Else
        grdDIAS.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub grdDIAS_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdDIAS
          Select Case Col
                    Case conCOL_PmSDIAS_QTDE
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
          End Select
     End With
End Sub

Private Sub grdDIAS_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
     With grdDIAS
          Select Case Col
                 Case conCOL_PmSDIAS_QTDE
                        If .EditText = Empty Then Exit Sub
          End Select
     End With
End Sub

Private Sub grdDIASSEM_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     With grdDIASSEM
          Select Case Col
                 Case conCOL_PmS_Qtde
                        Call RecalcCotas(CLng(.Cell(flexcpText, Row, conCOL_PmS_Semana)))
                 Case conCOL_PmS_Ativo
                        Call DesabFilho(Row)
          End Select
     End With
End Sub

Private Sub grdDIASSEM_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_PmS_Semana, _
         conCOL_PmS_IDINTERNO
        Cancel = True
    Case conCOL_PmS_Qtde, _
         conCOL_PmS_Ativo
        If cTipOper = "C" Then Cancel = True
    Case Else
        grdDIASSEM.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub grdDIASSEM_Click()
    Call MostraDados
End Sub

Private Sub grdDIASSEM_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdDIASSEM
          Select Case Col
                    Case conCOL_PmS_Qtde
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
          End Select
     End With
End Sub

Private Sub grdDIASSEM_RowColChange()
    Call MostraDados
End Sub

Private Sub grdDIASSEM_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
     With grdDIASSEM
          Select Case Col
                 Case conCOL_PmS_Qtde
                        If .EditText = Empty Then Exit Sub
          End Select
     End With
End Sub

Private Sub optSimNao_Click(Index As Integer)
     
     If cTipOper = "C" Then Exit Sub
     
     txtQTDE.Text = ""
     If Index = 0 Then
     
        If cboMes.ListIndex = -1 Then
           MsgBox "Informe o Mês dos pedidos !!!", vbOKOnly + vbExclamation, "Aviso"
           optSimNao.Item(1).Value = True
           Exit Sub
        End If
        If cboAno.ListIndex = -1 Then
           MsgBox "Informe o Ano dos pedidos !!!", vbOKOnly + vbExclamation, "Aviso"
           optSimNao.Item(1).Value = True
           Exit Sub
        End If
        If Len(Trim(txtCodProd.Text)) = 0 Then
            MsgBox "A Linha esta vázia !!!", vbOKOnly + vbExclamation, "Aviso"
            optSimNao(0).Value = False
            txtCodProd.SetFocus
            Exit Sub
        End If
        
        objCADPLAMESTRE.CODLINHA = CLng(txtCodProd.Text)
        objCADPLAMESTRE.QTDPEDIDOS = objCADPLAMESTRE.PegaPedidos(cboMes.ItemData(cboMes.ListIndex), cboAno.ItemData(cboAno.ListIndex))
        If objCADPLAMESTRE.QTDPEDIDOS > 0 Then txtQTDE.Text = Format(objCADPLAMESTRE.QTDPEDIDOS, "#0")
        
     End If
     
End Sub

Private Sub txtCodProd_GotFocus()
    objBLBFunc.SelecionaCampos txtCodProd.Name, frmCADPLAMESTRE
End Sub

Private Sub txtCodProd_KeyPress(KeyAscii As Integer)
   objBLBFunc.SoNumeroPonto KeyAscii, txtCodProd.Text
End Sub

Private Sub txtCodProd_Validate(Cancel As Boolean)

   Dim I As Integer

   If Len(Trim(txtCodProd.Text)) = 0 Then
      lblDescLinha.Caption = ""
      Exit Sub
   End If
   
   If Not IsNumeric(txtCodProd.Text) Then
        MsgBox "Atenção - Somente é permitido numeros !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCodProd.Text = ""
        txtCodProd.Tag = ""
        lblDescLinha.Caption = ""
        Cancel = True
        Exit Sub
   End If
   
   lblDescLinha.Caption = PegaDescrLinha(txtCodProd.Text, "SGI_CODIGO")
   If Len(Trim(lblDescLinha.Caption)) = 0 Then
        MsgBox "Esta Linha não existe !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCodProd.Text = ""
        Cancel = True
    End If
    
End Sub



Private Sub txtQTDE_GotFocus()
    objBLBFunc.SelecionaCampos txtQTDE.Name, frmCADPLAMESTRE
End Sub

Private Sub txtQTDE_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtQTDE.Text
End Sub

Private Sub txtQTDE_Validate(Cancel As Boolean)

    If Len(Trim(txtQTDE.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtQTDE.Text) Then
       MsgBox "Somente é permitido numeros e pontos !!!", vbOKOnly + vbExclamation, "Aviso"
       txtQTDE.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtQTDE.Text = Format(txtQTDE.Text, "#0")

End Sub

Private Function ValidaCampos() As Boolean

     ValidaCampos = False
     
     Dim curOF          As Currency
     Dim I              As Long
     Dim j              As Long
     Dim lngQTDE        As Long
     Dim lngQTDEINF     As Long
     Dim lngQTDDIAS     As Long
     Dim lngQTDEPASS    As Long
     Dim lngDIFDIAS     As Long
     
     If Trim(Len(txtCodProd.Text)) = 0 Then
        MsgBox "Informe o Código da linha !!!", vbOKOnly + vbCritical, "Aviso"
        StPlano.Tab = 0
        txtCodProd.SetFocus
        Exit Function
     End If
     If cboMes.ListIndex = -1 Then
        MsgBox "Mês inválido !!!", vbOKOnly + vbExclamation, "Aviso"
        StPlano.Tab = 0
        cboMes.SetFocus
        Exit Function
     End If
     If cboAno.ListIndex = -1 Then
        MsgBox "Ano inválido !!!", vbOKOnly + vbExclamation, "Aviso"
        StPlano.Tab = 0
        cboAno.SetFocus
        Exit Function
     End If
     
     If Len(Trim(txtQTDE.Text)) = 0 Then
        MsgBox "Quantidade inválida !!!", vbOKOnly + vbExclamation, "Aviso"
        StPlano.Tab = 0
        txtQTDE.SetFocus
        Exit Function
     End If
     
    '' ------------------------------------
    lngQTDE = 0
    With grdDIASSEM
        If (.Rows - 1) = 0 Then
            MsgBox "Não foi informado as quantidades das Semanas !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Function
        End If
        For I = 1 To (.Rows - 1)
            If CLng(.Cell(flexcpText, I, conCOL_PmS_Ativo)) = 1 Then
                lngQTDEINF = 0
                lngQTDDIAS = 0
                If Len(Trim(.Cell(flexcpText, I, conCOL_PmS_Qtde))) > 0 Then
                    lngQTDE = lngQTDE + CLng(.Cell(flexcpText, I, conCOL_PmS_Qtde))
                    lngQTDEINF = CLng(.Cell(flexcpText, I, conCOL_PmS_Qtde))
                End If
                
                For j = 1 To (grdDIAS.Rows - 1)
                    If grdDIAS.Cell(flexcpText, j, conCOL_PmSDIAS_SemanaIndice) = .Cell(flexcpText, I, conCOL_PmS_Semana) Then
                       If grdDIAS.Cell(flexcpText, j, conCOL_PmSDIAS_Ativo) = 1 Then
                            If Len(Trim(grdDIAS.Cell(flexcpText, j, conCOL_PmSDIAS_QTDE))) > 0 Then lngQTDDIAS = lngQTDDIAS + CLng(grdDIAS.Cell(flexcpText, j, conCOL_PmSDIAS_QTDE))
                       End If
                    End If
                Next j
                
                If lngQTDEINF <> lngQTDDIAS Then
                   MsgBox "A Soma Total das Latas na Semana : " & .Cell(flexcpText, I, conCOL_PmS_Semana) & " e de " & .Cell(flexcpText, I, conCOL_PmS_Qtde) & " e não confere com o total Informado na gride dias da Semana !!!", vbOKOnly + vbExclamation, "Aviso"
                   Exit Function
                End If
            End If
        Next I
    End With
    
    If lngQTDE <> CLng(txtQTDE.Text) Then
        MsgBox "A Soma Total das Latas na Gride Semanas e de " & lngQTDE & " e não confere com o total Informado !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Function
    End If
    '' ------------------------------------
     
     If cTipOper = "I" Then
        If optSimNao.Item(0).Value = True Then
            If objCADPLAMESTRE.QTDPEDIDOS > CCur(txtQTDE.Text) Then
               MsgBox "Atenção a quantidade do plano mestre etá menor que os pedidos no mês !!!" & vbCrLf & "Pedidos no Mês : " & Format(objCADPLAMESTRE.QTDPEDIDOS, "#,####0.0000"), vbOKOnly + vbExclamation, "Aviso"
               StPlano.Tab = 0
               txtQTDE.Text = Format(objCADPLAMESTRE.QTDPEDIDOS, "#0")
               txtQTDE.SetFocus
               Exit Function
            End If
        End If
     End If
     
     If cTipOper = "A" Then
        If optSimNao.Item(0).Value = True Then
            If objCADPLAMESTRE.QTDE <> CCur(txtQTDE.Text) Then
               objCADPLAMESTRE.QTDPEDIDOS = objCADPLAMESTRE.PegaPedidos(cboMes.ItemData(cboMes.ListIndex), cboAno.ItemData(cboAno.ListIndex))
               If objCADPLAMESTRE.QTDPEDIDOS > CCur(txtQTDE.Text) Then
                  MsgBox "Atenção a quantidade do plano mestre etá menor que os pedidos no mês !!!" & vbCrLf & "Pedidos no Mês : " & Format(objCADPLAMESTRE.QTDPEDIDOS, "#,####0.0000"), vbOKOnly + vbExclamation, "Aviso"
                  StPlano.Tab = 0
                  txtQTDE.Text = Format(objCADPLAMESTRE.QTDE, "#0")
                  txtQTDE.SetFocus
                  Exit Function
               End If
            End If
        
            curOF = objCADPLAMESTRE.VerifSaldoOF
        
            If curOF > CCur(txtQTDE.Text) Then
               MsgBox "A quantidade do plano mestre não pode seer menor que as OF's já inclusas !!!", vbOKOnly + vbExclamation, "Aviso"
               txtQTDE.Text = Format(curOF, "#0")
               Exit Function
            End If
        End If
     End If
     
     
     ValidaCampos = True
     
End Function


Private Sub Consulta()

    Dim I As Integer
    Dim j As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
    
    Me.Caption = "Cadastro de plano mestre de produção - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADPLAMESTRE
    objCADPLAMESTRE.CODIGO = iCodigo
    
    objCADPLAMESTRE.PreenchComboMes cboMes
    objCADPLAMESTRE.PreenchComboAno cboAno
    
    Call LimpaCamposLabel
    Call ConfGridPMS
    Call ConfGridPMSDIAS
    
    StPlano.Tab = 0
    optATIVOSN(1).Value = True
    
    If objCADPLAMESTRE.Carrega_campos(strTABELAFILIAL) = True Then
        
       txtCODIGO.Text = objCADPLAMESTRE.CODIGO
       txtCodProd.Text = objCADPLAMESTRE.CODLINHA
       lblDescLinha.Caption = PegaDescrLinha(txtCodProd.Text, "SGI_CODIGO")
       
       '' Mês
       cboMes.ListIndex = -1
       For I = 0 To (cboMes.ListCount - 1)
           If cboMes.ItemData(I) = objCADPLAMESTRE.MES Then cboMes.ListIndex = I
       Next I
       
       '' Ano
       cboAno.ListIndex = -1
       For I = 0 To (cboAno.ListCount - 1)
           If cboAno.ItemData(I) = objCADPLAMESTRE.ANO Then cboAno.ListIndex = I
       Next I
       
       optSimNao.Item(objCADPLAMESTRE.optSimNao).Value = True
       txtQTDE.Text = objCADPLAMESTRE.QTDE
       optATIVOSN(objCADPLAMESTRE.ATIVO).Value = True
       
        Call PopGrd
        Call PopGrdDiasSemana
    End If

End Sub


Private Sub Altera()

    Dim I As Integer
    Dim j As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    
    Me.Caption = "Cadastro de plano mestre de produção - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADPLAMESTRE
    objCADPLAMESTRE.CODIGO = iCodigo
    
    objCADPLAMESTRE.PreenchComboMes cboMes
    objCADPLAMESTRE.PreenchComboAno cboAno
    
    txtCodProd.Enabled = False
    Command2.Enabled = False
    cboMes.Enabled = False
    cboAno.Enabled = False
    
    StPlano.Tab = 0
    optATIVOSN(1).Value = True
    
    Call LimpaCamposLabel
    Call ConfGridPMS
    Call ConfGridPMSDIAS
    
    
    If objCADPLAMESTRE.Carrega_campos(strTABELAFILIAL) = True Then
        
       txtCODIGO.Text = objCADPLAMESTRE.CODIGO
       txtCodProd.Text = objCADPLAMESTRE.CODLINHA
       lblDescLinha.Caption = PegaDescrLinha(txtCodProd.Text, "SGI_CODIGO")
       
       '' Mês
       cboMes.ListIndex = -1
       For I = 0 To (cboMes.ListCount - 1)
           If cboMes.ItemData(I) = objCADPLAMESTRE.MES Then cboMes.ListIndex = I
       Next I
       
       '' Ano
       cboAno.ListIndex = -1
       For I = 0 To (cboAno.ListCount - 1)
           If cboAno.ItemData(I) = objCADPLAMESTRE.ANO Then cboAno.ListIndex = I
       Next I
       
       optSimNao.Item(objCADPLAMESTRE.optSimNao).Value = True
       txtQTDE.Text = objCADPLAMESTRE.QTDE
       
       Call PopGrd
       Call PopGrdDiasSemana
       
    End If

End Sub

Private Sub LimpaCamposLabel()
    lblDescLinha.Caption = ""
End Sub

Private Function PegaDescrLinha(strCODIGO As String, strCampo As String) As String
    
    PegaDescrLinha = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADGRUPLINHA" & strTABELAFILIAL & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And " & strCampo & " = " & strCODIGO
    sSql = sSql & "   And SGI_ATIVO  = 1" & vbCrLf
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
       PegaDescrLinha = BREC!SGI_DESCRI
    End If
    BREC.Close
    
End Function

Private Sub Destroy_Objeto()
       Set objBLBFunc = Nothing
       Set objCADPLAMESTRE = Nothing
       Set objPESQPADRAO = Nothing
End Sub

Private Sub PopGrd()
        Dim I As Integer
        arrITENSPM = objCADPLAMESTRE.ITENSDIAS
        If IsArray(arrITENSPM) Then
            With grdDIASSEM
                For I = 1 To UBound(arrITENSPM)
                    .AddItem arrITENSPM(I, 1) & vbTab & _
                             arrITENSPM(I, 2) & vbTab & _
                             arrITENSPM(I, 3) & vbTab & _
                             arrITENSPM(I, 4)
                Next I
            End With
        End If
End Sub

Private Function CarregaSemana(lngMes As Long, lngANO As Long) As String

    CarregaSemana = ""

    If lngMes = 0 Then Exit Function
    If lngANO = 0 Then Exit Function
    
    CarregaSemana = "01/" & Format(lngMes, "##00") & "/" & Format(lngANO, "####0000")

End Function

Private Function Verif_MesAno() As Boolean

    Verif_MesAno = False
    
    If cboMes.ListIndex = -1 Then
        MsgBox "Informe Primeiro o Mês !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Function
    End If
    If cboAno.ListIndex = -1 Then
        MsgBox "Informe o Ano !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Function
    End If
    If Not IsNumeric(txtQTDE.Text) Then
        MsgBox "ATENÇÂO - Informe a Quantidade !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Function
    End If
    
    Verif_MesAno = True
    
End Function

Private Sub IgualaSemana(dtDATA As Date)

    '' Pegando a Semana Final
    Dim dtFinal         As Date
    Dim dtFinal2        As Date
    Dim dtATUAL         As Date
    Dim lngSemanaIni    As Long
    Dim lngSemanaFin    As Long
    Dim lngTOTSEMANA    As Long
    Dim lngMes          As Long
    Dim lngANO          As Long
    Dim I               As Long
    Dim j               As Long
    Dim lngQTDE         As Long
    Dim lngQTDDIAS      As Long
    Dim lngGERSEMANA    As Long
    Dim dtDATADIA       As Date
    Dim lngSEMANA       As Long
    Dim lngINDICE       As Long
    Dim lngDIFDIAS      As Long
    Dim lngSOMA         As Long
    Dim lngSOMADIAS     As Long
    Dim lngTOTDIAS      As Long
    Dim arrSEMANA       As Variant
    
    
    Call ConfGridPMS
    Call ConfGridPMSDIAS
    
    lngQTDE = CLng(txtQTDE.Text)
    
    mtvSemana.Value = dtDATA
    lngSemanaIni = mtvSemana.Week
    lngSemanaFin = 0
    lngTOTSEMANA = 0
    
    If Month(dtDATA) = 12 Then
        lngMes = 1
        lngANO = (Year(dtDATA) + 1)
    Else
        lngMes = (Month(dtDATA) + 1)
        lngANO = Year(dtDATA)
    End If
    
    dtFinal = CDate("01/" & Format(lngMes, "##00") & "/" & Format(lngANO, "####0000"))
    dtFinal2 = CDate("01/" & Format(lngMes, "##00") & "/" & Format(lngANO, "####0000"))
    If Weekday(dtFinal) = 1 Then dtFinal = (dtFinal - 1)
    
    mtvSemana.Value = dtFinal
    lngSemanaFin = mtvSemana.Week
    
    ReDim arrSEMANA(1 To 7) As String
    arrSEMANA(1) = "Domingo"
    arrSEMANA(2) = "Segunda"
    arrSEMANA(3) = "Terça"
    arrSEMANA(4) = "Quarta"
    arrSEMANA(5) = "Quinta"
    arrSEMANA(6) = "Sexta"
    arrSEMANA(7) = "Sabado"
    
    '' Total de Dias
    lngQTDDIAS = (dtFinal2 - dtDATA)
    
    
    ''lngTOTSEMANA = (lngSemanaFin - lngSemanaIni)
    ''lngGERSEMANA = (lngQTDDIAS / lngTOTSEMANA)
    
    With grdDIASSEM
        For I = 1 To lngQTDDIAS
            dtDATADIA = CDate(Str(I) & "/" & cboMes.ItemData(cboMes.ListIndex) & "/" & cboAno.ItemData(cboAno.ListIndex))
            mtvSemana.Value = dtDATADIA
            lngSEMANA = mtvSemana.Week
            
            '' Pesquizando se já tem a Semana
            lngINDICE = .FindRow(lngSEMANA, , conCOL_PmS_Semana)
            
            If lngINDICE = -1 Then
                .AddItem lngSEMANA & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         1
            
                '' ==========================
                '' Dias da Semana
                For j = 1 To lngQTDDIAS
                   dtDATADIA = CDate(Str(j) & "/" & cboMes.ItemData(cboMes.ListIndex) & "/" & cboAno.ItemData(cboAno.ListIndex))
                   mtvSemana.Value = dtDATADIA
                   
                   If mtvSemana.Week = .Cell(flexcpText, (.Rows - 1), conCOL_PmS_Semana) Then
                       grdDIAS.AddItem mtvSemana.Week & vbTab & _
                                       Format(dtDATADIA, "DD/MM/YYYY") & vbTab & _
                                       Weekday(dtDATADIA) & vbTab & _
                                       arrSEMANA(Weekday(dtDATADIA)) & vbTab & _
                                       "" & vbTab & _
                                       "" & vbTab & _
                                       1
                   End If
                Next j
            End If
    
        Next I
        If (.Rows - 1) > 0 Then .Row = 1
    
        '' =========================
        lngQTDE = CLng(txtQTDE.Text)
        lngGERSEMANA = (.Rows - 1)
        lngTOTSEMANA = (lngQTDE / lngGERSEMANA)
        For I = 1 To lngGERSEMANA
            .Cell(flexcpText, I, conCOL_PmS_Qtde) = lngTOTSEMANA
            lngSOMA = (lngSOMA + lngTOTSEMANA)
            
            lngQTDDIAS = 0
            For j = 1 To (grdDIAS.Rows - 1)
                If grdDIAS.Cell(flexcpText, j, conCOL_PmSDIAS_SemanaIndice) = .Cell(flexcpText, I, conCOL_PmS_Semana) Then lngQTDDIAS = (lngQTDDIAS + 1)
            Next j
            
            '' =========================
            '' Dias da Semana
            lngTOTDIAS = (lngTOTSEMANA / lngQTDDIAS)
            lngSOMADIAS = 0
            For j = 1 To (grdDIAS.Rows - 1)
                If grdDIAS.Cell(flexcpText, j, conCOL_PmSDIAS_SemanaIndice) = .Cell(flexcpText, I, conCOL_PmS_Semana) Then
                    grdDIAS.Cell(flexcpText, j, conCOL_PmSDIAS_QTDE) = lngTOTDIAS
                    lngSOMADIAS = (lngSOMADIAS + lngTOTDIAS)
                End If
            Next j
            '' =========================
        
        Next I
    
        '' =========================
        '' Diferença
        lngDIFDIAS = (lngQTDE - lngSOMA)
        If lngDIFDIAS > 0 Then
            .Cell(flexcpText, 1, conCOL_PmS_Qtde) = CLng(.Cell(flexcpText, 1, conCOL_PmS_Qtde)) + lngDIFDIAS
        Else
            .Cell(flexcpText, 1, conCOL_PmS_Qtde) = CLng(.Cell(flexcpText, 1, conCOL_PmS_Qtde)) + lngDIFDIAS
        End If
        '' =========================
    
        '' =========================
        lngDIFDIAS = 0
        For I = 1 To (.Rows - 1)
            lngQTDE = CLng(.Cell(flexcpText, I, conCOL_PmS_Qtde))
            lngSOMADIAS = 0
            lngTOTDIAS = 0
            For j = 1 To (grdDIAS.Rows - 1)
                If grdDIAS.Cell(flexcpText, j, conCOL_PmSDIAS_SemanaIndice) = .Cell(flexcpText, I, conCOL_PmS_Semana) Then
                    lngTOTDIAS = CLng(grdDIAS.Cell(flexcpText, j, conCOL_PmSDIAS_QTDE))
                    lngSOMADIAS = (lngSOMADIAS + lngTOTDIAS)
                End If
            Next j
            '' Diferença
            lngDIFDIAS = (lngQTDE - lngSOMADIAS)
            lngINDICE = grdDIAS.FindRow(.Cell(flexcpText, I, conCOL_PmS_Semana), , conCOL_PmSDIAS_SemanaIndice)
            If lngINDICE > -1 Then
                grdDIAS.Cell(flexcpText, lngINDICE, conCOL_PmSDIAS_QTDE) = (CLng(grdDIAS.Cell(flexcpText, lngINDICE, conCOL_PmSDIAS_QTDE)) + lngDIFDIAS)
            End If
        Next I
        
        
    End With
    Call MostraDados

End Sub


Private Sub ConfGridPMS()

    With grdDIASSEM
    
       .Cols = conColumnsIn_PmS
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_PmS_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_PmS_Semana) = ""
       .ColDataType(conCOL_PmS_Semana) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_PmS_Qtde) = ""
       .ColDataType(conCOL_PmS_Qtde) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PmS_IDINTERNO) = ""
       .ColDataType(conCOL_PmS_IDINTERNO) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PmS_Ativo) = ""
       .ColDataType(conCOL_PmS_Ativo) = flexDTString
       .ColComboList(conCOL_PmS_Ativo) = objCADPLAMESTRE.PreenchComboAtivo
       
       .ColWidth(conCOL_PmS_Semana) = 1000
       .ColWidth(conCOL_PmS_Qtde) = 1200
       .ColWidth(conCOL_PmS_IDINTERNO) = 0
       .ColWidth(conCOL_PmS_Ativo) = 700
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With

End Sub

Private Sub IncRegGridCotas()
   
    If Len(Trim(txtCodProd.Text)) = 0 Then
        MsgBox "Primeiro informe a linha !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCodProd.SetFocus
        Exit Sub
    End If
    If Len(Trim(txtQTDE.Text)) = 0 Then
        MsgBox "Primeiro informe a Qtde de Latas !!!", vbOKOnly + vbExclamation, "Aviso"
        txtQTDE.SetFocus
        Exit Sub
    End If
    
    If Verif_MesAno = False Then Exit Sub
    
    dtDIAMESANO = CDate(CarregaSemana(cboMes.ItemData(cboMes.ListIndex), cboAno.ItemData(cboAno.ListIndex)))
    Call IgualaSemana(dtDIAMESANO)
    
End Sub



Private Sub ConfGridPMSDIAS()

    With grdDIAS
    
       .Cols = conColumnsIn_PmSDIAS
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_PmSDIAS_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_PmSDIAS_SemanaIndice) = ""
       .ColDataType(conCOL_PmSDIAS_SemanaIndice) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PmSDIAS_DATA) = ""
       .ColDataType(conCOL_PmSDIAS_DATA) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_PmSDIAS_DIASEMANA) = ""
       .ColDataType(conCOL_PmSDIAS_DIASEMANA) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PmSDIAS_DIA) = ""
       .ColDataType(conCOL_PmSDIAS_DIA) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PmSDIAS_QTDE) = ""
       .ColDataType(conCOL_PmSDIAS_QTDE) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PmSDIAS_IDINTERNO) = ""
       .ColDataType(conCOL_PmSDIAS_IDINTERNO) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PmSDIAS_Ativo) = ""
       .ColDataType(conCOL_PmSDIAS_Ativo) = flexDTString
       .ColComboList(conCOL_PmSDIAS_Ativo) = objCADPLAMESTRE.PreenchComboAtivo
       
       .ColWidth(conCOL_PmSDIAS_SemanaIndice) = 0
       .ColWidth(conCOL_PmSDIAS_DATA) = 1200
       .ColWidth(conCOL_PmSDIAS_DIASEMANA) = 0
       .ColWidth(conCOL_PmSDIAS_DIA) = 1000
       .ColWidth(conCOL_PmSDIAS_QTDE) = 1000
       .ColWidth(conCOL_PmSDIAS_IDINTERNO) = 0
       .ColWidth(conCOL_PmSDIAS_Ativo) = 700
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With

End Sub


Private Sub MostraDados()
    
On Error GoTo Err_MostraDados
    
    With grdDIASSEM
        If (.Rows - 1) > 0 And .Row > 0 Then
            Call objBLBFunc.CarregaDadosGrdFilhoSemAction2Do(grdDIAS, conCOL_PmSDIAS_SemanaIndice, .Cell(flexcpText, .Row, conCOL_PmS_Semana))
            Frame6.Caption = "[ Dias da Semana ]  : " & .Cell(flexcpText, .Row, conCOL_PmS_Semana)
        End If
    End With
    
    Exit Sub
    
Err_MostraDados:
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : MostraDados()", Me.Name, "MostraDados()", strCAMARQERRO)
    
End Sub


Private Sub DesabFilho(lngROW As Long)

    Dim I As Integer
    With grdDIAS
        For I = 1 To (.Rows - 1)
            If grdDIASSEM.Cell(flexcpText, lngROW, conCOL_PmS_Semana) = .Cell(flexcpText, I, conCOL_PmSDIAS_SemanaIndice) Then
                .Cell(flexcpText, I, conCOL_PmSDIAS_Ativo) = grdDIASSEM.Cell(flexcpText, lngROW, conCOL_PmS_Ativo)
            End If
        Next I
    End With
    

End Sub

Private Sub PopGrdDiasSemana()
        
        Dim I               As Integer
        Dim arrSEMANA       As Variant
        
        ReDim arrSEMANA(1 To 7) As String
        arrSEMANA(1) = "Domingo"
        arrSEMANA(2) = "Segunda"
        arrSEMANA(3) = "Terça"
        arrSEMANA(4) = "Quarta"
        arrSEMANA(5) = "Quinta"
        arrSEMANA(6) = "Sexta"
        arrSEMANA(7) = "Sabado"
        
        arrITENSDIAS = objCADPLAMESTRE.ITENSDIASSEM
        If IsArray(arrITENSDIAS) Then
            With grdDIAS
                For I = 1 To UBound(arrITENSDIAS)
                    .AddItem arrITENSDIAS(I, 1) & vbTab & _
                             arrITENSDIAS(I, 2) & vbTab & _
                             arrITENSDIAS(I, 3) & vbTab & _
                             arrSEMANA(CInt(arrITENSDIAS(I, 3))) & vbTab & _
                             arrITENSDIAS(I, 4) & vbTab & _
                             arrITENSDIAS(I, 5) & vbTab & _
                             arrITENSDIAS(I, 6)
                Next I
            End With
        End If
        If (grdDIASSEM.Rows - 1) > 0 Then grdDIASSEM.Row = 1
        Call MostraDados

End Sub

Private Sub RecalcCotas(lngSEMANA As Long)

    Dim I               As Long
    Dim lngQTDDIAS      As Long
    Dim intLINHA        As Integer
    Dim lngQTDE         As Long
    Dim lngQTDE2        As Long
    Dim lngDIFERENCA    As Long
    
    lngQTDDIAS = 0
    lngQTDE = 0
    lngQTDE2 = 0
    
    With grdDIAS
        For I = 1 To (.Rows - 1)
            If CLng(.Cell(flexcpText, I, conCOL_PmSDIAS_SemanaIndice)) = lngSEMANA And _
               CLng(.Cell(flexcpText, I, conCOL_PmSDIAS_Ativo)) = 1 Then
               lngQTDDIAS = lngQTDDIAS + 1
            End If
        Next I
    End With

    If lngQTDDIAS > 0 Then
        intLINHA = grdDIASSEM.FindRow(lngSEMANA, , conCOL_PmS_Semana)
        If intLINHA > 0 Then
            lngQTDE = (CLng(grdDIASSEM.Cell(flexcpText, intLINHA, conCOL_PmS_Qtde)) / lngQTDDIAS)
            With grdDIAS
                For I = 1 To (.Rows - 1)
                    If CLng(.Cell(flexcpText, I, conCOL_PmSDIAS_SemanaIndice)) = lngSEMANA And _
                       CLng(.Cell(flexcpText, I, conCOL_PmSDIAS_Ativo)) = 1 Then
                       .Cell(flexcpText, I, conCOL_PmSDIAS_QTDE) = lngQTDE
                       lngQTDE2 = (lngQTDE2 + lngQTDE)
                    End If
                Next I
            End With
            
            lngDIFERENCA = (CLng(grdDIASSEM.Cell(flexcpText, intLINHA, conCOL_PmS_Qtde)) - (lngQTDE * lngQTDDIAS))
                With grdDIAS
                    For I = 1 To (.Rows - 1)
                        If CLng(.Cell(flexcpText, I, conCOL_PmSDIAS_SemanaIndice)) = lngSEMANA And _
                           CLng(.Cell(flexcpText, I, conCOL_PmSDIAS_Ativo)) = 1 Then
                           .Cell(flexcpText, I, conCOL_PmSDIAS_QTDE) = (CLng(.Cell(flexcpText, I, conCOL_PmSDIAS_QTDE)) + lngDIFERENCA)
                           Exit Sub
                        End If
                    Next I
                End With
        End If
    End If

End Sub
