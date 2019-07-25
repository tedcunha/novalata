VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCADFICHATECNICA 
   Caption         =   "Cadastro de ficha técnica de produto"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10515
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   10515
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab stdados 
      Height          =   5415
      Left            =   120
      TabIndex        =   22
      Top             =   3120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   9551
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Parâmetros"
      TabPicture(0)   =   "frmCADFICHATECNICA.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grdParamFicha"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdIncIten"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdExcIten"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Dados &Complementares"
      TabPicture(1)   =   "frmCADFICHATECNICA.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "stDadosGerais"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin TabDlg.SSTab stDadosGerais 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   30
         Top             =   360
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   8705
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Velocidade de Produção"
         TabPicture(0)   =   "frmCADFICHATECNICA.frx":0038
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "grdVelocidade"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "grdUnidades"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Command2"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Command3"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Command4"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Command5"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Setup de Máquina"
         TabPicture(1)   =   "frmCADFICHATECNICA.frx":0054
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame5"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame5 
            Height          =   615
            Left            =   -74880
            TabIndex        =   37
            Top             =   420
            Width           =   9855
            Begin VB.TextBox txtMinutos 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   4080
               TabIndex        =   39
               Text            =   "txtMinutos"
               Top             =   240
               Width           =   2415
            End
            Begin VB.ComboBox cboUidade 
               Height          =   315
               Left            =   3120
               TabIndex        =   38
               Text            =   "cboMAQUINA"
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Unidade"
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
               Index           =   4
               Left            =   2280
               TabIndex        =   41
               Top             =   240
               Width           =   720
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Tempo de Setup em"
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
               Left            =   120
               TabIndex        =   40
               Top             =   240
               Width           =   1710
            End
         End
         Begin VB.CommandButton Command5 
            Height          =   315
            Left            =   9600
            Picture         =   "frmCADFICHATECNICA.frx":0070
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   2460
            Width           =   375
         End
         Begin VB.CommandButton Command4 
            Height          =   315
            Left            =   9600
            Picture         =   "frmCADFICHATECNICA.frx":01AE
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   2820
            Width           =   375
         End
         Begin VB.CommandButton Command3 
            Height          =   315
            Left            =   9600
            Picture         =   "frmCADFICHATECNICA.frx":0738
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   420
            Width           =   375
         End
         Begin VB.CommandButton Command2 
            Height          =   315
            Left            =   9600
            Picture         =   "frmCADFICHATECNICA.frx":0876
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   780
            Width           =   375
         End
         Begin VSFlex8LCtl.VSFlexGrid grdUnidades 
            Height          =   1935
            Left            =   120
            TabIndex        =   31
            Top             =   480
            Width           =   9375
            _cx             =   16536
            _cy             =   3413
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
         Begin VSFlex8LCtl.VSFlexGrid grdVelocidade 
            Height          =   2415
            Left            =   120
            TabIndex        =   34
            Top             =   2460
            Width           =   9375
            _cx             =   16536
            _cy             =   4260
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
      Begin VB.CommandButton cmdExcIten 
         Height          =   315
         Left            =   9840
         Picture         =   "frmCADFICHATECNICA.frx":0E00
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   780
         Width           =   375
      End
      Begin VB.CommandButton cmdIncIten 
         Height          =   315
         Left            =   9840
         Picture         =   "frmCADFICHATECNICA.frx":138A
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   420
         Width           =   375
      End
      Begin VSFlex8LCtl.VSFlexGrid grdParamFicha 
         Height          =   4815
         Left            =   120
         TabIndex        =   18
         Top             =   420
         Width           =   9735
         _cx             =   17171
         _cy             =   8493
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
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   0
      TabIndex        =   26
      Top             =   960
      Width           =   10455
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   3480
         Picture         =   "frmCADFICHATECNICA.frx":14C8
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtCODFAMMAQUINA 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2280
         TabIndex        =   7
         Text            =   "txtCODFAMMAQUINA"
         Top             =   960
         Width           =   1215
      End
      Begin VB.ComboBox cboFAMMAQUINA 
         Height          =   315
         Left            =   3840
         TabIndex        =   8
         Text            =   "cboFAMMAQUINA"
         Top             =   960
         Width           =   6495
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5520
         TabIndex        =   15
         Top             =   1680
         Width           =   2175
         Begin VB.OptionButton optPADRAO 
            Caption         =   "Sim"
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
            TabIndex        =   16
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton optPADRAO 
            Caption         =   "Não"
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
            Left            =   960
            TabIndex        =   17
            Top             =   0
            Width           =   855
         End
      End
      Begin MSMask.MaskEdBox mskData 
         Height          =   285
         Left            =   9120
         TabIndex        =   2
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboProd 
         Height          =   315
         Left            =   3840
         TabIndex        =   5
         Text            =   "cboProd"
         Top             =   600
         Width           =   6495
      End
      Begin VB.CommandButton cmdProd 
         Height          =   315
         Left            =   3480
         Picture         =   "frmCADFICHATECNICA.frx":15CA
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtCodProd 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   4
         Text            =   "txtCodProd"
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox cboMAQUINA 
         Height          =   315
         Left            =   3840
         TabIndex        =   11
         Text            =   "cboMAQUINA"
         Top             =   1320
         Width           =   6495
      End
      Begin VB.TextBox txtCODMAQUINA 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2280
         TabIndex        =   10
         Text            =   "txtCODMAQUINA"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton Command11 
         Height          =   315
         Left            =   3480
         Picture         =   "frmCADFICHATECNICA.frx":16CC
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "txtCodigo"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Familia de Máquina"
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
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1650
      End
      Begin VB.Label lblCORRPRODPADR 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblCORRPRODPADR"
         Height          =   255
         Left            =   2280
         TabIndex        =   13
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Produto Padrão"
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
         Left            =   4080
         TabIndex        =   14
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Correlação do Produto"
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
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   1920
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data"
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
         Left            =   8520
         TabIndex        =   1
         Top             =   240
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cód. Produto"
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
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1125
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Máquina"
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
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   21
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
         Picture         =   "frmCADFICHATECNICA.frx":17CE
         Style           =   1  'Graphical
         TabIndex        =   25
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
         MaskColor       =   &H8000000F&
         Picture         =   "frmCADFICHATECNICA.frx":18D0
         Style           =   1  'Graphical
         TabIndex        =   24
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
         Picture         =   "frmCADFICHATECNICA.frx":19D2
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADFICHATECNICA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho    As String
Public Linha       As Variant
Public cTipOper    As String
Public iCodigo     As Integer
Public FILIAL      As Integer
Public strAcesso   As String
Dim objBLBFunc     As Object
Dim objCADFICHATEC As Object
Dim objPESQPADRAO  As Object
Dim arrGRDFICHATEC As Variant
Dim arrGRDCOEFIC   As Variant
Dim arrFAMMEDIDAS  As Variant
Dim arrEFICIENCIA  As Variant

Const conCOL_SonParam_CodParam                      As Integer = 0
Const conCOL_SonParam_PesqParam                     As Integer = 1
Const conCOL_SonParam_Desc_Param                    As Integer = 2
Const conCOL_SonParam_UniMed                        As Integer = 3
Const conCOL_SonParam_ValParam                      As Integer = 4
Const conCOL_SonParam_ValParamPos                   As Integer = 5
Const conCOL_SonParam_ValParamNeg                   As Integer = 6
Const conCOL_SonParam_FormatString                  As String = "=Cód. Param|...|Descrição Parâmetros|Unidade|Parâmetro|Parâmetro(+)|Parâmetro(-)"
Const conColumnsIn_SonParam                         As Integer = 7

Const conCOL_SonUnidades_CodUnidDe                  As Integer = 0
Const conCOL_SonUnidades_PesqUnidDe                 As Integer = 1
Const conCOL_SonUnidades_UnidDe                     As Integer = 2
Const conCOL_SonUnidades_Desc_UnidDe                As Integer = 3
Const conCOL_SonUnidades_CodUnidPara                As Integer = 4
Const conCOL_SonUnidades_PesqUnidPara               As Integer = 5
Const conCOL_SonUnidades_UnidPara                   As Integer = 6
Const conCOL_SonUnidades_Desc_UnidPara              As Integer = 7
Const conCOL_SonUnidades_Default                    As Integer = 8
Const conCOL_SonUnidades_FamDe                      As Integer = 9
Const conCOL_SonUnidades_FamPara                    As Integer = 10
Const conCOL_SonUnidades_Indice                     As Integer = 11
Const conCOL_SonUnidades_FormatString               As String = "=Cód. De|...|Unid.|Descr. Unid. De|Cód. Para|...|Unid.|Descr. Unid. Para|Default|FamDe|FamPara|Indice"
Const conColumnsIn_SonUnidades                      As Integer = 12

Const conCOL_SonVelicidade_EficMed                  As Integer = 0
Const conCOL_SonVelocidade_ProdTeorica              As Integer = 1
Const conCOL_SonVelocidade_ProdReal                 As Integer = 2
Const conCOL_SonVelocidade_Unidade                  As Integer = 3
Const conCOL_SonVelocidade_Indice                   As Integer = 4
Const conCOL_SonVelocidade_FormatString             As String = "=Eficiência Média %|Produção Teórica|Produção Real|Unidade|Indice"
Const conColumnsIn_SonVelocidade                    As Integer = 5

Private Sub cboFAMMAQUINA_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboFAMMAQUINA, KeyAscii
End Sub

Private Sub cboFAMMAQUINA_Validate(Cancel As Boolean)
    If cboFAMMAQUINA.ListIndex > -1 Then
       txtCODFAMMAQUINA.Text = cboFAMMAQUINA.ItemData(cboFAMMAQUINA.ListIndex)
       Call objCADFICHATEC.PreenchComboMaquina(cboMAQUINA, CLng(txtCODFAMMAQUINA.Text))
    End If
End Sub

Private Sub cboMAQUINA_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboMAQUINA, KeyAscii
End Sub

Private Sub cboMAQUINA_Validate(Cancel As Boolean)
    If cboMAQUINA.ListIndex > -1 Then txtCODMAQUINA.Text = cboMAQUINA.ItemData(cboMAQUINA.ListIndex)
End Sub

Private Sub cboProd_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboProd, KeyAscii
End Sub

Private Sub cboProd_Validate(Cancel As Boolean)
    If cboProd.ListIndex > -1 Then
       txtCodProd.Text = Mid(cboProd.List(cboProd.ListIndex), 1, 10)
       Call PegaDadosDoProd(Trim(txtCodProd.Text))
       Call objCADFICHATEC.PreenchComboFamMaquina(cboFAMMAQUINA, txtCodProd.Text)
    End If
End Sub

Private Sub cboUidade_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboUidade, KeyAscii
End Sub

Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame3.Enabled = False
   
    Me.Caption = "Cadastro de ficha técnica de produto - [ ALTERAÇÃO ]"
    cTipOper = "A"
    
End Sub


Private Sub cmdExcIten_Click()
    If cTipOper = "I" Or cTipOper = "A" Then Call objBLBFunc.ExclLinhaGrid(grdParamFicha, grdParamFicha.Row)
End Sub

Private Sub cmdIncIten_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncRegGrid
End Sub

Private Sub cmdProd_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL      = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1500"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Produtos")
    
    If Len(Trim(varRETORNO)) > 0 Then
       txtCodProd.Text = varRETORNO
       Call PegaDadosDoProd(Trim(txtCodProd.Text))
       Call objCADFICHATEC.PreenchComboFamMaquina(cboFAMMAQUINA, txtCodProd.Text)
    End If
    
    cboProd.ListIndex = -1
    txtCodProd.SetFocus

End Sub

Private Sub CmdSalva_Click()

    Dim I As Integer
    
    If ValidaCampos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADFICHATEC.CODIGO = objCADFICHATEC.Gera_Codigo(Me.Name)

    objCADFICHATEC.Data = CDate(mskData.Text)
    objCADFICHATEC.CODPROD = txtCodProd.Text
    objCADFICHATEC.CODFAMMAQ = txtCODFAMMAQUINA.Text
    objCADFICHATEC.CODMAQ = txtCODMAQUINA.Text
    
    '' =====================================================
    If Len(Trim(txtMinutos.Text)) > 0 Then objCADFICHATEC.MINUTOS = CCur(txtMinutos.Text)
    objCADFICHATEC.UNIDSETTUP = cboUidade.ItemData(cboUidade.ListIndex)
    
    '' =====================================================
    '' Parâmetros da Ficha Técnica
    arrGRDFICHATEC = Empty
    If (grdParamFicha.Rows - 1) > 0 Then
        ReDim arrGRDFICHATEC(1 To (grdParamFicha.Rows - 1), 1 To 5) As Variant
        For I = 1 To (grdParamFicha.Rows - 1)
            arrGRDFICHATEC(I, 1) = grdParamFicha.Cell(flexcpText, I, conCOL_SonParam_CodParam)
            arrGRDFICHATEC(I, 2) = grdParamFicha.Cell(flexcpText, I, conCOL_SonParam_UniMed)
            arrGRDFICHATEC(I, 3) = grdParamFicha.Cell(flexcpText, I, conCOL_SonParam_ValParam)
            arrGRDFICHATEC(I, 4) = grdParamFicha.Cell(flexcpText, I, conCOL_SonParam_ValParamPos)
            arrGRDFICHATEC(I, 5) = grdParamFicha.Cell(flexcpText, I, conCOL_SonParam_ValParamNeg)
        Next I
    End If
    objCADFICHATEC.ITENS = arrGRDFICHATEC
    
    '' =====================================================
    '' Familias de Unidade de Medida
    arrFAMMEDIDAS = Empty
    If (grdUnidades.Rows - 1) > 0 Then
        ReDim arrFAMMEDIDAS(1 To (grdUnidades.Rows - 1), 1 To 4) As Variant
        For I = 1 To (grdUnidades.Rows - 1)
            arrFAMMEDIDAS(I, 1) = grdUnidades.Cell(flexcpText, I, conCOL_SonUnidades_CodUnidDe)
            arrFAMMEDIDAS(I, 2) = grdUnidades.Cell(flexcpText, I, conCOL_SonUnidades_CodUnidPara)
            If grdUnidades.Cell(flexcpTextDisplay, I, conCOL_SonUnidades_Default) = "Sim" Then arrFAMMEDIDAS(I, 3) = 1
            If grdUnidades.Cell(flexcpTextDisplay, I, conCOL_SonUnidades_Default) = "Não" Then arrFAMMEDIDAS(I, 3) = 0
            arrFAMMEDIDAS(I, 4) = grdUnidades.Cell(flexcpTextDisplay, I, conCOL_SonUnidades_Indice)
        Next I
    End If
    objCADFICHATEC.FAMUNIDADE = arrFAMMEDIDAS
    '' =====================================================
    
    '' =====================================================
    '' Eficiência
    arrEFICIENCIA = Empty
    If (grdVelocidade.Rows - 1) > 0 Then
        ReDim arrEFICIENCIA(1 To (grdVelocidade.Rows - 1), 1 To 4) As Variant
        For I = 1 To (grdVelocidade.Rows - 1)
            arrEFICIENCIA(I, 1) = grdVelocidade.Cell(flexcpText, I, conCOL_SonVelicidade_EficMed)
            arrEFICIENCIA(I, 2) = grdVelocidade.Cell(flexcpText, I, conCOL_SonVelocidade_ProdTeorica)
            arrEFICIENCIA(I, 3) = grdVelocidade.Cell(flexcpText, I, conCOL_SonVelocidade_ProdReal)
            arrEFICIENCIA(I, 4) = grdVelocidade.Cell(flexcpText, I, conCOL_SonVelocidade_Indice)
        Next I
    End If
    objCADFICHATEC.EFICIENCIA = arrEFICIENCIA
    '' =====================================================
    
    If objCADFICHATEC.GRAVA(cTipOper) = False Then Exit Sub
    If objCADFICHATEC.Atualiza(cTipOper, Str(objCADFICHATEC.CODIGO), FILIAL, Me.Name) = False Then Exit Sub
          
    MsgBox "A ficha foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
       
    If cTipOper = "I" Then Call Inclui

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objCADFICHATEC = Nothing
    Set objPESQPADRAO = Nothing
    Unload Me
End Sub

Private Sub Command1_Click()

    If Len(Trim(txtCodProd.Text)) = 0 Then Exit Sub
    
    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       FM.* " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO PR" & vbCrLf
    sSql = sSql & "      ,SGI_CADFAMMAQUINAS FM" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       PR.SGI_FILIAL      = " & FILIAL & vbCrLf
    sSql = sSql & "   And PR.SGI_CODIGO      = '" & txtCodProd.Text & "'" & vbCrLf
    sSql = sSql & "   And FM.SGI_FILIAL      = PR.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And FM.SGI_CODIGO      = PR.SGI_CADFAMMAQ "
    
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Familia de Máquinas")
    
    If Len(Trim(varRETORNO)) > 0 Then
       txtCODFAMMAQUINA.Text = varRETORNO
       Call objCADFICHATEC.PreenchComboMaquina(cboMAQUINA, CLng(txtCODFAMMAQUINA.Text))
    End If
    
    cboFAMMAQUINA.ListIndex = -1
    txtCODFAMMAQUINA.SetFocus

End Sub

Private Sub Command11_Click()

    If Len(Trim(txtCODFAMMAQUINA.Text)) = 0 Then Exit Sub
    
    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADMAQUINA " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    sSql = sSql & "   And SGI_CODFAMILIA = " & txtCODFAMMAQUINA.Text
    
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
        
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Maquina")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODMAQUINA.Text = varRETORNO
    
    cboMAQUINA.ListIndex = -1
    txtCODMAQUINA.SetFocus

End Sub

Private Sub Command2_Click()
    If cTipOper = "I" Or cTipOper = "A" Then Call objBLBFunc.ExclLinhaGrid(grdUnidades, grdUnidades.Row)
End Sub

Private Sub Command3_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncFamUnidade
End Sub

Private Sub Command5_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncVelociadade
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADFICHATEC = CreateObject("CADFICHATECNICA.clsCADFICHATECNICA")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
   objCADFICHATEC.FILIAL = FILIAL
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame3.Enabled = False
   
    Me.Caption = "Cadastro de ficha técnica de produto - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADFICHATECNICA
    
    Call InitGridParam
    Call InitGridFamUnidade
    Call InitGridVelocidade
    
    objCADFICHATEC.PreencheComboProd cboProd
    objCADFICHATEC.PreenchComboUnidade cboUidade
    
    lblCORRPRODPADR.Caption = ""
    
    mskData.Text = Format(Date, "DD/MM/YYYY")
    
    If txtCodProd.Enabled = True And txtCodProd.Visible = True Then txtCodProd.SetFocus
    
End Sub

Private Sub InitGridParam()

    With grdParamFicha
    
       .Cols = conColumnsIn_SonParam
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonParam_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonParam_CodParam) = ""
       .ColDataType(conCOL_SonParam_CodParam) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonParam_PesqParam) = ""
       .ColDataType(conCOL_SonParam_PesqParam) = flexDTString
       .ColComboList(conCOL_SonParam_PesqParam) = "..."
       
       .Cell(flexcpData, 0, conCOL_SonParam_Desc_Param) = ""
       .ColDataType(conCOL_SonParam_Desc_Param) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonParam_UniMed) = ""
       .ColDataType(conCOL_SonParam_UniMed) = flexDTString
       
       .ColComboList(conCOL_SonParam_UniMed) = objCADFICHATEC.PreenchComboUnidadeGrid
       
       .Cell(flexcpData, 0, conCOL_SonParam_ValParam) = ""
       .ColDataType(conCOL_SonParam_ValParam) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonParam_ValParamPos) = ""
       .ColDataType(conCOL_SonParam_ValParamPos) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonParam_ValParamNeg) = ""
       .ColDataType(conCOL_SonParam_ValParamNeg) = flexDTCurrency
       
       .ColWidth(conCOL_SonParam_CodParam) = 1000
       .ColWidth(conCOL_SonParam_PesqParam) = 300
       .ColWidth(conCOL_SonParam_Desc_Param) = 3000
       .ColWidth(conCOL_SonParam_UniMed) = 800
       .ColWidth(conCOL_SonParam_ValParam) = 1500
       .ColWidth(conCOL_SonParam_ValParamPos) = 1500
       .ColWidth(conCOL_SonParam_ValParamNeg) = 1500
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub IncRegGrid()
   
    If ExisteLinhaVazia = False Then Exit Sub
    
    grdParamFicha.AddItem "" & vbTab & _
                          "" & vbTab & _
                          "" & vbTab & _
                          "" & vbTab & _
                          "" & vbTab & _
                          "" & vbTab & _
                          ""
                            
End Sub

Private Function ExisteLinhaVazia() As Boolean
    ExisteLinhaVazia = False
    
    Dim I As Integer
    
    For I = 1 To (grdParamFicha.Rows - 1)
        If grdParamFicha.Cell(flexcpText, I, conCOL_SonParam_CodParam) = Empty Then Exit Function
    Next I
    
    ExisteLinhaVazia = True
End Function


Private Function ExisteLinhaVaziaUnidade() As Boolean
    ExisteLinhaVaziaUnidade = False
    
    Dim I As Integer
    
    With grdUnidades
         For I = 1 To (.Rows - 1)
             If .Cell(flexcpText, I, conCOL_SonUnidades_CodUnidDe) = Empty Or _
                .Cell(flexcpText, I, conCOL_SonUnidades_CodUnidPara) = Empty Then Exit Function
         Next I
    End With
    
    ExisteLinhaVaziaUnidade = True
End Function

Private Sub grdParamFicha_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Select Case Col
    Case conCOL_SonParam_CodParam
         grdParamFicha.Col = Col + 3
         grdParamFicha.EditCell
    Case conCOL_SonParam_PesqParam
         grdParamFicha.Col = Col + 2
         grdParamFicha.EditCell
    Case conCOL_SonParam_UniMed, _
         conCOL_SonParam_ValParam, _
         conCOL_SonParam_ValParamPos, _
         conCOL_SonParam_ValParamNeg
         If (grdParamFicha.Cols - 1) <> grdParamFicha.Col Then
            grdParamFicha.Col = Col + 1
            grdParamFicha.EditCell
         End If
    End Select
    Exit Sub
End Sub

Private Sub grdParamFicha_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonParam_Desc_Param
         Cancel = True
    Case conCOL_SonParam_CodParam, _
         conCOL_SonParam_PesqParam, _
         conCOL_SonParam_UniMed, _
         conCOL_SonParam_ValParam, _
         conCOL_SonParam_ValParamNeg, _
         conCOL_SonParam_ValParamPos
         If cTipOper = "C" Then Cancel = True
    Case Else
        grdParamFicha.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub grdParamFicha_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    If cTipOper = "C" Then Exit Sub
    
    If (grdParamFicha.Rows - 1) = 0 Then Exit Sub
    
    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    Select Case Col
        Case conCOL_SonParam_PesqParam
    
            sSql = "Select " & vbCrLf
            sSql = sSql & "       * " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CADPARFICHA " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
            
            arrTABELA(1) = sSql
            
            arrCAMPOS(1, 1) = "SGI_CODIGO"
            arrCAMPOS(1, 2) = "S"
            arrCAMPOS(1, 3) = "Código"
            arrCAMPOS(1, 4) = "1500"
            arrCAMPOS(1, 5) = "SGI_CODIGO"
            
            arrCAMPOS(2, 1) = "SGI_DESCRI"
            arrCAMPOS(2, 2) = "S"
            arrCAMPOS(2, 3) = "Nome"
            arrCAMPOS(2, 4) = "5000"
            arrCAMPOS(2, 5) = "SGI_DESCRI"
            
            varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Usuários")
            
            If Len(Trim(varRETORNO)) > 0 Then
               grdParamFicha.Cell(flexcpText, Row, conCOL_SonParam_CodParam) = varRETORNO
               grdParamFicha.Cell(flexcpText, Row, conCOL_SonParam_Desc_Param) = PegaDescrParametro(CLng(grdParamFicha.Cell(flexcpText, Row, conCOL_SonParam_CodParam)))
            End If
            
            If VerifItensRepetidos(Row, conCOL_SonParam_CodParam, varRETORNO) = False Then
               MsgBox "Este parâmetro já foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
               grdParamFicha.Cell(flexcpText, Row, conCOL_SonParam_CodParam) = Empty
               grdParamFicha.Cell(flexcpText, Row, conCOL_SonParam_Desc_Param) = Empty
               Exit Sub
            End If

    End Select

End Sub

Private Function PegaDescrParametro(lngCodUsuario As Long) As String
    PegaDescrParametro = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPARFICHA " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCodUsuario
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaDescrParametro = BREC!SGI_DESCRI
    BREC.Close
    
End Function

Private Function PegaDescrFamUnidades(lngCodUsuario As Long) As String
    PegaDescrFamUnidades = ""
    
    With grdUnidades
         sSql = "Select " & vbCrLf
         sSql = sSql & "       * " & vbCrLf
         sSql = sSql & "  From " & vbCrLf
         sSql = sSql & "       SGI_CADUNIMED " & vbCrLf
         sSql = sSql & " Where " & vbCrLf
         sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
         sSql = sSql & "   And SGI_CODIGO = " & lngCodUsuario
        
         BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         If Not BREC.EOF Then PegaDescrFamUnidades = BREC!SGI_DESCRICAO
         BREC.Close
    End With
    
End Function

Private Function PegaUnidade(lngCodUsuario As Long) As String
    PegaUnidade = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADUNIMED " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCodUsuario
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaUnidade = BREC!SGI_UNIDADE
    BREC.Close
    
End Function



Private Function VerifItensRepetidos(intRow As Long, intCol As Long, varCampo As Variant) As Boolean
    VerifItensRepetidos = False
    Dim I As Integer
    
    If Not IsNumeric(varCampo) Then varCampo = UCase(Trim(varCampo))
    
    For I = 1 To (grdParamFicha.Rows - 1)
        If I <> intRow And grdParamFicha.Cell(flexcpText, I, intCol) = varCampo Then Exit Function
    Next I
    VerifItensRepetidos = True
End Function

Private Function VerifItensRepetidosUnidades(intRow As Long, intCol As Long, varCampo As Variant) As Boolean
    VerifItensRepetidosUnidades = False
    Dim I As Integer
    
    If Not IsNumeric(varCampo) Then varCampo = UCase(Trim(varCampo))
    
    For I = 1 To (grdUnidades.Rows - 1)
        If I <> intRow And grdUnidades.Cell(flexcpText, I, intCol) = varCampo Then Exit Function
    Next I
    VerifItensRepetidosUnidades = True
End Function


Private Sub grdParamFicha_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdParamFicha
          Select Case Col
                    Case conCOL_SonParam_CodParam
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
                    Case conCOL_SonParam_ValParam, conCOL_SonParam_ValParamPos, conCOL_SonParam_ValParamNeg
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 2, myvarAsCurrency)
          End Select
     End With
End Sub

Private Sub grdParamFicha_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
     With grdParamFicha
          Select Case Col
                 Case conCOL_SonParam_CodParam
                        If .EditText = Empty Then Exit Sub
                        If VerifItensRepetidos(Row, conCOL_SonParam_CodParam, .EditText) = False Then
                           MsgBox "Este parâmetro ja foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                           grdParamFicha.Cell(flexcpText, Row, conCOL_SonParam_CodParam) = Empty
                           grdParamFicha.Cell(flexcpText, Row, conCOL_SonParam_Desc_Param) = Empty
                           Cancel = True
                           Exit Sub
                        End If
                        If Len(Trim(PegaDescrParametro(CLng(.EditText)))) = 0 Then
                           MsgBox "Este parâmetro não existe !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, Col) = Empty
                           .Cell(flexcpText, Row, conCOL_SonParam_Desc_Param) = Empty
                           Cancel = True
                           Exit Sub
                        End If
                        .Cell(flexcpText, Row, conCOL_SonParam_Desc_Param) = PegaDescrParametro(CLng(.EditText))
                 Case conCOL_SonParam_ValParam
                        If (.EditText = Empty And Len(Trim(.Cell(flexcpText, Row, conCOL_SonParam_ValParamPos)))) > 0 Or _
                            (.EditText = Empty And Len(Trim(.Cell(flexcpText, Row, conCOL_SonParam_ValParamNeg)))) > 0 Then
                            MsgBox "Informe o valor do Parâmetro !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        Else
                            If .EditText <> Empty And Len(Trim(.Cell(flexcpText, Row, conCOL_SonParam_ValParamPos))) > 0 Then
                                If CCur(.EditText) > CCur(.Cell(flexcpText, Row, conCOL_SonParam_ValParamPos)) Then
                                    MsgBox "O Valor do Parâmetro não pode ser maior que o valor Parâmetro(+) !!!", vbOKOnly + vbExclamation, "Aviso"
                                    Cancel = True
                                    Exit Sub
                                End If
                            End If
                            If .EditText <> Empty And Len(Trim(.Cell(flexcpText, Row, conCOL_SonParam_ValParamNeg))) > 0 Then
                                If CCur(.EditText) < CCur(.Cell(flexcpText, Row, conCOL_SonParam_ValParamNeg)) Then
                                    MsgBox "O Valor do Parâmetro não pode ser menor que o valor Parâmetro(-) !!!", vbOKOnly + vbExclamation, "Aviso"
                                    Cancel = True
                                    Exit Sub
                                End If
                            End If
                        End If
                 Case conCOL_SonParam_ValParamPos
                        If (.EditText = Empty And Len(Trim(.Cell(flexcpText, Row, conCOL_SonParam_ValParam)))) > 0 Then
                            MsgBox "Informe o valor do Parâmetro(+) !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                        If (.EditText <> Empty And Len(Trim(.Cell(flexcpText, Row, conCOL_SonParam_ValParam)))) > 0 Then
                            If CCur(.EditText) < CCur(.Cell(flexcpText, Row, conCOL_SonParam_ValParam)) Then
                                MsgBox "O Valor do Parâmetro(+) não pode ser menor que o valor Parâmetro !!!", vbOKOnly + vbExclamation, "Aviso"
                                Cancel = True
                                Exit Sub
                            End If
                        End If
                 Case conCOL_SonParam_ValParamNeg
                        If (.EditText = Empty And Len(Trim(.Cell(flexcpText, Row, conCOL_SonParam_ValParam)))) > 0 Then
                            MsgBox "Informe o valor do Parâmetro(-) !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                        If (.EditText <> Empty And Len(Trim(.Cell(flexcpText, Row, conCOL_SonParam_ValParam)))) > 0 Then
                            If CCur(.EditText) > CCur(.Cell(flexcpText, Row, conCOL_SonParam_ValParam)) Then
                                MsgBox "O Valor do Parâmetro(-) não pode ser maior que o valor Parâmetro !!!", vbOKOnly + vbExclamation, "Aviso"
                                Cancel = True
                                Exit Sub
                            End If
                        End If
          End Select
     End With
End Sub

Private Sub grdUnidades_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim I As Integer
    With grdUnidades
         Select Case Col
         Case conCOL_SonUnidades_Default
              If .Cell(flexcpTextDisplay, Row, Col) = "Sim" Then
                 For I = 1 To (.Rows - 1)
                     If .Row <> I Then .Cell(flexcpText, I, conCOL_SonUnidades_Default) = 0
                 Next I
              End If
         End Select
    End With
    Exit Sub
End Sub

Private Sub grdUnidades_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonUnidades_Desc_UnidDe, _
         conCOL_SonUnidades_Desc_UnidPara, _
         conCOL_SonUnidades_UnidDe, _
         conCOL_SonUnidades_UnidPara
         Cancel = True
    Case conCOL_SonUnidades_CodUnidDe, _
         conCOL_SonUnidades_CodUnidPara, _
         conCOL_SonUnidades_PesqUnidDe, _
         conCOL_SonUnidades_PesqUnidPara, _
         conCOL_SonUnidades_Default
         If cTipOper = "C" Then Cancel = True
    Case Else
        grdUnidades.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub grdUnidades_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    Dim I As Integer
    
    If cTipOper = "C" Then Exit Sub
    
    With grdUnidades
         If (.Rows - 1) = 0 Then Exit Sub
        
         ReDim arrCAMPOS(1 To 4, 1 To 5) As String
         ReDim arrTABELA(1 To 1) As String
        
         Select Case Col
                Case conCOL_SonUnidades_PesqUnidDe
                    
                     sSql = "Select " & vbCrLf
                     sSql = sSql & "       UND.SGI_CODIGO    " & vbCrLf
                     sSql = sSql & "      ,UND.SGI_UNIDADE   " & vbCrLf
                     sSql = sSql & "      ,UND.SGI_DESCRICAO " & vbCrLf
                     sSql = sSql & "      ,FAM.SGI_DESCRI    " & vbCrLf
                     sSql = sSql & "  From " & vbCrLf
                     sSql = sSql & "       SGI_CADUNIMED     UND " & vbCrLf
                     sSql = sSql & "      ,SGI_CADFAMUNIDADE FAM " & vbCrLf
                     sSql = sSql & " Where " & vbCrLf
                     sSql = sSql & "       UND.SGI_FILIAL  = " & FILIAL & vbCrLf
                     If Len(Trim(grdUnidades.Cell(flexcpText, Row, conCOL_SonUnidades_FamPara))) <> 0 Then
                        sSql = sSql & "   And UND.SGI_CODFAMUNID <> " & .Cell(flexcpText, Row, conCOL_SonUnidades_FamPara)
                     End If
                     sSql = sSql & "   And FAM.SGI_FILIAL     = UND.SGI_FILIAL " & vbCrLf
                     sSql = sSql & "   And FAM.SGI_CODIGO     = UND.SGI_CODFAMUNID "
                     
                     arrTABELA(1) = sSql
                     
                     arrCAMPOS(1, 1) = "SGI_CODIGO"
                     arrCAMPOS(1, 2) = "N"
                     arrCAMPOS(1, 3) = "Código"
                     arrCAMPOS(1, 4) = "1000"
                     arrCAMPOS(1, 5) = "UND.SGI_CODIGO"
                    
                     arrCAMPOS(2, 1) = "SGI_UNIDADE"
                     arrCAMPOS(2, 2) = "S"
                     arrCAMPOS(2, 3) = "Unidade"
                     arrCAMPOS(2, 4) = "1000"
                     arrCAMPOS(2, 5) = "UND.SGI_UNIDADE"
                     
                     arrCAMPOS(3, 1) = "SGI_DESCRICAO"
                     arrCAMPOS(3, 2) = "S"
                     arrCAMPOS(3, 3) = "Descrição"
                     arrCAMPOS(3, 4) = "2000"
                     arrCAMPOS(3, 5) = "UND.SGI_DESCRICAO"
                     
                     arrCAMPOS(4, 1) = "SGI_DESCRI"
                     arrCAMPOS(4, 2) = "S"
                     arrCAMPOS(4, 3) = "Familia"
                     arrCAMPOS(4, 4) = "3000"
                     arrCAMPOS(4, 5) = "FAM.SGI_DESCRI"
                     
                     varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Familia de Unidades")
                    
                     If Len(Trim(varRETORNO)) > 0 Then
                        .Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidDe) = varRETORNO
                        .Cell(flexcpText, Row, conCOL_SonUnidades_UnidDe) = PegaUnidade(CLng(.Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidDe)))
                        .Cell(flexcpText, Row, conCOL_SonUnidades_FamDe) = PegaCodFamUnidades(CLng(.Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidDe)))
                        .Cell(flexcpText, Row, conCOL_SonUnidades_Desc_UnidDe) = PegaDescrFamUnidades(CLng(.Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidDe)))
                        
                        If VerifItensRepetidosUnidades(Row, conCOL_SonUnidades_Indice, Trim(varRETORNO) & Trim(grdUnidades.Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidPara))) = False Then
                           MsgBox "Esta unidade já foi relacionada na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidPara) = Empty
                           .Cell(flexcpText, Row, conCOL_SonUnidades_UnidPara) = Empty
                           .Cell(flexcpText, Row, conCOL_SonUnidades_Desc_UnidPara) = Empty
                           .Cell(flexcpText, Row, conCOL_SonUnidades_FamDe) = ""
                           .Cell(flexcpText, Row, conCOL_SonUnidades_Indice) = Trim(.Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidPara))
                           Exit Sub
                        End If
                        
                        I = grdVelocidade.FindRow(.Cell(flexcpText, Row, conCOL_SonUnidades_Indice), , conCOL_SonVelocidade_Indice)
                        If I <> -1 Then
                           grdVelocidade.Cell(flexcpText, I, conCOL_SonVelocidade_Indice) = Trim(varRETORNO) & Trim(.Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidPara))
                           grdVelocidade.Cell(flexcpText, I, conCOL_SonVelocidade_Unidade) = Trim(.Cell(flexcpText, Row, conCOL_SonUnidades_UnidDe)) & "/" & Trim(.Cell(flexcpText, Row, conCOL_SonUnidades_UnidPara))
                        End If
                        
                        .Cell(flexcpText, Row, conCOL_SonUnidades_Indice) = Trim(varRETORNO) & Trim(.Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidPara))
                     End If
                    
               Case conCOL_SonUnidades_PesqUnidPara
               
                     sSql = "Select " & vbCrLf
                     sSql = sSql & "       UND.SGI_CODIGO    " & vbCrLf
                     sSql = sSql & "      ,UND.SGI_UNIDADE   " & vbCrLf
                     sSql = sSql & "      ,UND.SGI_DESCRICAO " & vbCrLf
                     sSql = sSql & "      ,FAM.SGI_DESCRI    " & vbCrLf
                     sSql = sSql & "  From " & vbCrLf
                     sSql = sSql & "       SGI_CADUNIMED     UND " & vbCrLf
                     sSql = sSql & "      ,SGI_CADFAMUNIDADE FAM " & vbCrLf
                     sSql = sSql & " Where " & vbCrLf
                     sSql = sSql & "       UND.SGI_FILIAL  = " & FILIAL & vbCrLf
                     If Len(Trim(grdUnidades.Cell(flexcpText, Row, conCOL_SonUnidades_FamDe))) <> 0 Then
                        sSql = sSql & "   And UND.SGI_CODFAMUNID <> " & .Cell(flexcpText, Row, conCOL_SonUnidades_FamDe)
                     End If
                     sSql = sSql & "   And FAM.SGI_FILIAL     = UND.SGI_FILIAL " & vbCrLf
                     sSql = sSql & "   And FAM.SGI_CODIGO     = UND.SGI_CODFAMUNID "
                    
                     arrTABELA(1) = sSql
                    
                     arrCAMPOS(1, 1) = "SGI_CODIGO"
                     arrCAMPOS(1, 2) = "N"
                     arrCAMPOS(1, 3) = "Código"
                     arrCAMPOS(1, 4) = "1000"
                     arrCAMPOS(1, 5) = "SGI_CODIGO"
                    
                     arrCAMPOS(2, 1) = "SGI_UNIDADE"
                     arrCAMPOS(2, 2) = "S"
                     arrCAMPOS(2, 3) = "Unidade"
                     arrCAMPOS(2, 4) = "1000"
                     arrCAMPOS(2, 5) = "SGI_UNIDADE"
                    
                     arrCAMPOS(3, 1) = "SGI_DESCRICAO"
                     arrCAMPOS(3, 2) = "S"
                     arrCAMPOS(3, 3) = "Descrição"
                     arrCAMPOS(3, 4) = "2000"
                     arrCAMPOS(3, 5) = "SGI_DESCRICAO"
                    
                     arrCAMPOS(4, 1) = "SGI_DESCRI"
                     arrCAMPOS(4, 2) = "S"
                     arrCAMPOS(4, 3) = "Familia"
                     arrCAMPOS(4, 4) = "3000"
                     arrCAMPOS(4, 5) = "FAM.SGI_DESCRI"
                    
                     varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Familia de Unidades")
                    
                     If Len(Trim(varRETORNO)) > 0 Then
                        .Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidPara) = varRETORNO
                        .Cell(flexcpText, Row, conCOL_SonUnidades_UnidPara) = PegaUnidade(CLng(.Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidPara)))
                        .Cell(flexcpText, Row, conCOL_SonUnidades_FamPara) = PegaCodFamUnidades(CLng(.Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidPara)))
                        .Cell(flexcpText, Row, conCOL_SonUnidades_Desc_UnidPara) = PegaDescrFamUnidades(CLng(.Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidPara)))
                        
                    
                        If VerifItensRepetidosUnidades(Row, conCOL_SonUnidades_Indice, Trim(.Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidDe)) & Trim(varRETORNO)) = False Then
                           MsgBox "Esta unidade já foi relacionada na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidPara) = Empty
                           .Cell(flexcpText, Row, conCOL_SonUnidades_UnidPara) = Empty
                           .Cell(flexcpText, Row, conCOL_SonUnidades_Desc_UnidPara) = Empty
                           .Cell(flexcpText, Row, conCOL_SonUnidades_FamPara) = ""
                           .Cell(flexcpText, Row, conCOL_SonUnidades_Indice) = Trim(.Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidDe))
                           Exit Sub
                        End If
                        
                        I = grdVelocidade.FindRow(.Cell(flexcpText, Row, conCOL_SonUnidades_Indice), , conCOL_SonVelocidade_Indice)
                        If I <> -1 Then
                           grdVelocidade.Cell(flexcpText, I, conCOL_SonVelocidade_Indice) = Trim(.Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidDe)) & Trim(varRETORNO)
                           grdVelocidade.Cell(flexcpText, I, conCOL_SonVelocidade_Unidade) = Trim(.Cell(flexcpText, Row, conCOL_SonUnidades_UnidDe)) & "/" & Trim(.Cell(flexcpText, Row, conCOL_SonUnidades_UnidPara))
                        End If
                        
                        .Cell(flexcpText, Row, conCOL_SonUnidades_Indice) = Trim(.Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidDe)) & Trim(varRETORNO)
                     End If
               
         End Select
    End With
End Sub

Private Sub grdUnidades_Click()
    If (grdUnidades.Rows - 1) > 0 And grdUnidades.Row > 0 Then Call PosRegGrdEficiencia(grdUnidades.Cell(flexcpText, grdUnidades.Row, conCOL_SonUnidades_Indice))
End Sub

Private Sub grdUnidades_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdUnidades
          Select Case Col
                    Case conCOL_SonUnidades_CodUnidDe, _
                         conCOL_SonUnidades_CodUnidPara
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
          End Select
     End With
End Sub

Private Sub grdUnidades_RowColChange()
    If (grdUnidades.Rows - 1) > 0 And grdUnidades.Row > 0 Then Call PosRegGrdEficiencia(grdUnidades.Cell(flexcpText, grdUnidades.Row, conCOL_SonUnidades_Indice))
End Sub

Private Sub grdUnidades_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
     
     Dim I As Integer
     With grdUnidades
          Select Case Col
                 Case conCOL_SonUnidades_CodUnidDe
                      If .EditText = Empty Then
                         .Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidDe) = Empty
                         .Cell(flexcpText, Row, conCOL_SonUnidades_UnidDe) = Empty
                         .Cell(flexcpText, Row, conCOL_SonUnidades_Desc_UnidDe) = Empty
                         .Cell(flexcpText, Row, conCOL_SonUnidades_FamDe) = ""
                         .Cell(flexcpText, Row, conCOL_SonUnidades_Indice) = Trim(grdUnidades.Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidPara))
                         Exit Sub
                      End If
                      If VerifItensRepetidosUnidades(Row, conCOL_SonUnidades_Indice, Trim(.EditText) & Trim(grdUnidades.Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidPara))) = False Then
                         MsgBox "Esta unidade já foi relacionada na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                         .Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidDe) = Empty
                         .Cell(flexcpText, Row, conCOL_SonUnidades_UnidDe) = Empty
                         .Cell(flexcpText, Row, conCOL_SonUnidades_Desc_UnidDe) = Empty
                         .Cell(flexcpText, Row, conCOL_SonUnidades_FamDe) = ""
                         .Cell(flexcpText, Row, conCOL_SonUnidades_Indice) = Trim(.Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidPara))
                         Cancel = True
                         Exit Sub
                      End If
                      If Len(Trim(PegaDescrFamUnidades(CLng(.EditText)))) = 0 Then
                         MsgBox "Esta unidade não existe !!!", vbOKOnly + vbExclamation, "Aviso"
                         .Cell(flexcpText, Row, Col) = Empty
                         .Cell(flexcpText, Row, conCOL_SonUnidades_UnidDe) = Empty
                         .Cell(flexcpText, Row, conCOL_SonUnidades_Desc_UnidDe) = Empty
                         .Cell(flexcpText, Row, conCOL_SonUnidades_FamDe) = ""
                         .Cell(flexcpText, Row, conCOL_SonUnidades_Indice) = Trim(.Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidPara))
                         Cancel = True
                         Exit Sub
                      End If
                      If PegaCodFamUnidades(CLng(.EditText)) = .Cell(flexcpText, Row, conCOL_SonUnidades_FamPara) Then
                         MsgBox "Esta unidade pretence a mesma familia. !!!", vbOKOnly + vbExclamation, "Aviso"
                         .Cell(flexcpText, Row, Col) = Empty
                         .Cell(flexcpText, Row, conCOL_SonUnidades_UnidDe) = Empty
                         .Cell(flexcpText, Row, conCOL_SonUnidades_Desc_UnidDe) = Empty
                         .Cell(flexcpText, Row, conCOL_SonUnidades_FamDe) = ""
                         .Cell(flexcpText, Row, conCOL_SonUnidades_Indice) = Trim(.Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidPara))
                         Cancel = True
                         Exit Sub
                      End If
                      .Cell(flexcpText, Row, conCOL_SonUnidades_UnidDe) = PegaUnidade(CLng(.EditText))
                      .Cell(flexcpText, Row, conCOL_SonUnidades_Desc_UnidDe) = PegaDescrFamUnidades(CLng(.EditText))
                      .Cell(flexcpText, Row, conCOL_SonUnidades_FamDe) = PegaCodFamUnidades(CLng(.EditText))
                      
                       I = grdVelocidade.FindRow(.Cell(flexcpText, Row, conCOL_SonUnidades_Indice), , conCOL_SonVelocidade_Indice)
                       If I <> -1 Then
                          grdVelocidade.Cell(flexcpText, I, conCOL_SonVelocidade_Indice) = Trim(.EditText) & Trim(.Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidPara))
                          grdVelocidade.Cell(flexcpText, I, conCOL_SonVelocidade_Unidade) = Trim(.Cell(flexcpText, Row, conCOL_SonUnidades_UnidDe)) & "/" & Trim(.Cell(flexcpText, Row, conCOL_SonUnidades_UnidPara))
                       End If
                      .Cell(flexcpText, Row, conCOL_SonUnidades_Indice) = Trim(.EditText) & Trim(.Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidPara))
                      
                 Case conCOL_SonUnidades_CodUnidPara
                      If .EditText = Empty Then
                         .Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidPara) = Empty
                         .Cell(flexcpText, Row, conCOL_SonUnidades_UnidPara) = Empty
                         .Cell(flexcpText, Row, conCOL_SonUnidades_Desc_UnidPara) = Empty
                         .Cell(flexcpText, Row, conCOL_SonUnidades_FamPara) = ""
                         .Cell(flexcpText, Row, conCOL_SonUnidades_Indice) = Trim(.Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidDe))
                         Exit Sub
                      End If
                      If VerifItensRepetidosUnidades(Row, conCOL_SonUnidades_Indice, Trim(.Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidDe)) & Trim(.EditText)) = False Then
                         MsgBox "Esta unidade já foi relacionada na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                         .Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidPara) = Empty
                         .Cell(flexcpText, Row, conCOL_SonUnidades_UnidPara) = Empty
                         .Cell(flexcpText, Row, conCOL_SonUnidades_Desc_UnidPara) = Empty
                         .Cell(flexcpText, Row, conCOL_SonUnidades_FamPara) = ""
                         .Cell(flexcpText, Row, conCOL_SonUnidades_Indice) = Trim(.Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidDe))
                         Cancel = True
                         Exit Sub
                      End If
                      If Len(Trim(PegaDescrFamUnidades(CLng(.EditText)))) = 0 Then
                         MsgBox "Esta unidade não existe !!!", vbOKOnly + vbExclamation, "Aviso"
                         .Cell(flexcpText, Row, Col) = Empty
                         .Cell(flexcpText, Row, conCOL_SonUnidades_UnidPara) = Empty
                         .Cell(flexcpText, Row, conCOL_SonUnidades_Desc_UnidPara) = Empty
                         .Cell(flexcpText, Row, conCOL_SonUnidades_FamPara) = ""
                         .Cell(flexcpText, Row, conCOL_SonUnidades_Indice) = Trim(.Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidDe))
                         Cancel = True
                         Exit Sub
                      End If
                      If PegaCodFamUnidades(CLng(.EditText)) = .Cell(flexcpText, Row, conCOL_SonUnidades_FamDe) Then
                         MsgBox "Esta unidade pretence a mesma familia. !!!", vbOKOnly + vbExclamation, "Aviso"
                         .Cell(flexcpText, Row, Col) = Empty
                         .Cell(flexcpText, Row, conCOL_SonUnidades_UnidPara) = Empty
                         .Cell(flexcpText, Row, conCOL_SonUnidades_Desc_UnidPara) = Empty
                         .Cell(flexcpText, Row, conCOL_SonUnidades_FamPara) = ""
                         .Cell(flexcpText, Row, conCOL_SonUnidades_Indice) = Trim(.Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidDe))
                         Cancel = True
                         Exit Sub
                      End If
                      .Cell(flexcpText, Row, conCOL_SonUnidades_UnidPara) = PegaUnidade(CLng(.EditText))
                      .Cell(flexcpText, Row, conCOL_SonUnidades_Desc_UnidPara) = PegaDescrFamUnidades(CLng(.EditText))
                      .Cell(flexcpText, Row, conCOL_SonUnidades_FamPara) = PegaCodFamUnidades(CLng(.EditText))
                      
                       I = grdVelocidade.FindRow(.Cell(flexcpText, Row, conCOL_SonUnidades_Indice), , conCOL_SonVelocidade_Indice)
                       If I <> -1 Then
                          grdVelocidade.Cell(flexcpText, I, conCOL_SonVelocidade_Indice) = Trim(.Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidDe)) & Trim(.EditText)
                          grdVelocidade.Cell(flexcpText, I, conCOL_SonVelocidade_Unidade) = Trim(.Cell(flexcpText, Row, conCOL_SonUnidades_UnidDe)) & "/" & Trim(.Cell(flexcpText, Row, conCOL_SonUnidades_UnidPara))
                       End If
                      .Cell(flexcpText, Row, conCOL_SonUnidades_Indice) = Trim(.Cell(flexcpText, Row, conCOL_SonUnidades_CodUnidDe)) & Trim(.EditText)
          End Select
     End With
End Sub

Private Sub grdVelocidade_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With grdVelocidade
        Select Case Col
               Case conCOL_SonVelicidade_EficMed
                    If Len(Trim(.Cell(flexcpText, .Row, conCOL_SonVelicidade_EficMed))) > 0 Then .Cell(flexcpText, .Row, conCOL_SonVelicidade_EficMed) = Format(.Cell(flexcpText, .Row, conCOL_SonVelicidade_EficMed), "#,##0.00")
               Case conCOL_SonVelocidade_ProdTeorica
                    If Len(Trim(.Cell(flexcpText, .Row, conCOL_SonVelocidade_ProdTeorica))) > 0 Then .Cell(flexcpText, .Row, conCOL_SonVelocidade_ProdTeorica) = Format(.Cell(flexcpText, .Row, conCOL_SonVelocidade_ProdTeorica), "#,##0.00")
        End Select
        .Cell(flexcpText, .Row, conCOL_SonVelocidade_ProdReal) = Format(CalcPcRealGrid(.Cell(flexcpText, .Row, conCOL_SonVelicidade_EficMed), .Cell(flexcpText, .Row, conCOL_SonVelocidade_ProdTeorica)), "#,##0.00")
    End With
End Sub

Private Sub grdVelocidade_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonVelocidade_ProdReal, conCOL_SonVelocidade_Unidade
         Cancel = True
    End Select
    Exit Sub
End Sub

Private Sub grdVelocidade_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdVelocidade
          Select Case Col
                    Case conCOL_SonVelicidade_EficMed, conCOL_SonVelocidade_ProdTeorica
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 2, myvarAsCurrency)
          End Select
     End With
End Sub

Private Sub grdVelocidade_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

     With grdVelocidade
          Select Case Col
                 Case conCOL_SonVelicidade_EficMed
                      If Len(Trim(.EditText)) > 0 Then
                         If CLng(.EditText) > 100 Then
                            MsgBox "Somente é Permitido até 100% !!!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                         End If
                         If CLng(.EditText) < 0 Then
                            MsgBox "Somente é Permitido valores negativos !!!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                         End If
                      End If
                 Case conCOL_SonVelocidade_ProdTeorica
                     If Len(Trim(.EditText)) > 0 Then
                        If CLng(.EditText) < 0 Then
                           MsgBox "Somente é Permitido valores negativos !!!!", vbOKOnly + vbExclamation, "Aviso"
                           Cancel = True
                        End If
                     End If
          End Select
     End With

End Sub

Private Sub txtCODFAMMAQUINA_GotFocus()
    objBLBFunc.SelecionaCampos txtCODFAMMAQUINA.Name, frmCADFICHATECNICA
End Sub

Private Sub txtCODFAMMAQUINA_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODFAMMAQUINA.Text
End Sub

Private Sub txtCODFAMMAQUINA_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODFAMMAQUINA.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODFAMMAQUINA.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODFAMMAQUINA.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboFAMMAQUINA.ListIndex = -1
    For I = 0 To (cboFAMMAQUINA.ListCount - 1)
        If cboFAMMAQUINA.ItemData(I) = CInt(txtCODFAMMAQUINA.Text) Then cboFAMMAQUINA.ListIndex = I
    Next I
    
    If cboFAMMAQUINA.ListIndex = -1 Then
       MsgBox "Esta Familia de máquina não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODFAMMAQUINA.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Call objCADFICHATEC.PreenchComboMaquina(cboMAQUINA, CLng(txtCODFAMMAQUINA.Text))

End Sub

Private Sub txtCODMAQUINA_GotFocus()
    objBLBFunc.SelecionaCampos txtCODMAQUINA.Name, frmCADFICHATECNICA
End Sub

Private Sub txtCODMAQUINA_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODMAQUINA.Text
End Sub

Private Sub txtCODMAQUINA_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODMAQUINA.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODMAQUINA.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODMAQUINA.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboMAQUINA.ListIndex = -1
    For I = 0 To (cboMAQUINA.ListCount - 1)
        If cboMAQUINA.ItemData(I) = CInt(txtCODMAQUINA.Text) Then cboMAQUINA.ListIndex = I
    Next I
    
    If cboMAQUINA.ListIndex = -1 Then
       MsgBox "Esta máquina não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODMAQUINA.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtCodProd_GotFocus()
    objBLBFunc.SelecionaCampos txtCodProd.Name, frmCADFICHATECNICA
End Sub

Private Sub txtCodProd_Validate(Cancel As Boolean)

   Dim I As Integer

   If Len(Trim(txtCodProd.Text)) = 0 Then Exit Sub
   
   cboProd.ListIndex = -1
   For I = 0 To (cboProd.ListCount - 1)
       If Trim(Mid(cboProd.List(I), 1, 10)) = Trim(txtCodProd.Text) Then cboProd.ListIndex = I
   Next I
    
   If cboProd.ListIndex = -1 Then
      MsgBox "Esta produto não existe !!!", vbOKOnly + vbCritical, "Aviso"
      txtCodProd.Text = ""
      Cancel = True
      Exit Sub
   End If

   Call PegaDadosDoProd(Trim(txtCodProd.Text))
   Call objCADFICHATEC.PreenchComboFamMaquina(cboFAMMAQUINA, txtCodProd.Text)
   
End Sub

Private Function ValidaCampos() As Boolean

     Dim I          As Integer
     Dim boolAchou  As Boolean
     Dim curVALPARM As Currency
     
     ValidaCampos = False
     
     If Not IsDate(mskData.Text) Then
        MsgBox "Data da ficha inválida !!!", vbOKOnly + vbCritical, "Aviso"
        mskData.SetFocus
        Exit Function
     End If
     If Trim(Len(txtCodProd.Text)) = 0 Then
        MsgBox "Informe o código do produto !!!", vbOKOnly + vbCritical, "Aviso"
        txtCodProd.SetFocus
        Exit Function
     End If
     If Trim(Len(txtCODFAMMAQUINA.Text)) = 0 Then
        MsgBox "Informe o código da familia de máquinas !!!", vbOKOnly + vbCritical, "Aviso"
        txtCODFAMMAQUINA.SetFocus
        Exit Function
     End If
     If Trim(Len(txtCODMAQUINA.Text)) = 0 Then
        MsgBox "Informe o código da máquina !!!", vbOKOnly + vbCritical, "Aviso"
        txtCODMAQUINA.SetFocus
        Exit Function
     End If
     
     If (cboUidade.ListIndex) = -1 Then
        MsgBox "Informe a unidade do setup de máquinas !!!", vbOKOnly + vbExclamation, "Aviso"
        cboUidade.SetFocus
        stDadosGerais.Tab = 1
        Exit Function
     End If
     If Len(Trim(txtMinutos.Text)) = 0 Then
        MsgBox "Informe o tempo de Setup deste artigo !!!", vbOKOnly + vbExclamation, "Aviso"
        txtMinutos.SetFocus
        stDadosGerais.Tab = 1
        Exit Function
     End If
     
     If (grdUnidades.Rows - 1) = 0 Then
        MsgBox "Informe a unidade de velocidade !!!", vbOKOnly + vbExclamation, "Aviso"
        stDadosGerais.Tab = 0
        grdUnidades.SetFocus
        Exit Function
     End If
     
     
     boolAchou = False
     With grdUnidades
        For I = 1 To (.Rows - 1)
            If grdUnidades.Cell(flexcpTextDisplay, I, conCOL_SonUnidades_Default) = "Sim" Then
               boolAchou = True
            End If
        Next I
     End With
     If boolAchou = False Then
        MsgBox "Informe o unidade default !!!", vbOKOnly + vbCritical, "Aviso"
        txtCODMAQUINA.SetFocus
        Exit Function
     End If
     
     
     Call objBLBFunc.RemoveLinhaVazia(grdUnidades, conCOL_SonUnidades_CodUnidDe)
     Call objBLBFunc.RemoveLinhaVazia(grdParamFicha, conCOL_SonParam_CodParam)
     
     If (grdParamFicha.Rows - 1) > 0 Then
        With grdParamFicha
            For I = 1 To (.Rows - 1)
                If .Cell(flexcpText, I, conCOL_SonParam_CodParam) = Empty Then
                   MsgBox "Á tabela 'Parâmetros' o registro " & I & " 'Campo Códgo do parâmetro' deve ser preenchido !!!", vbOKOnly + vbExclamation, "Aviso"
                   Exit Function
                End If
                If .Cell(flexcpText, I, conCOL_SonParam_UniMed) = Empty Then
                   MsgBox "Á tabela 'Parâmetros' o registro " & I & " 'Campo Unidade de medida' deve ser preenchido !!!", vbOKOnly + vbExclamation, "Aviso"
                   Exit Function
                End If
                If .Cell(flexcpText, I, conCOL_SonParam_ValParam) = Empty Then
                   MsgBox "Á tabela 'Parâmetros' o registro " & I & " 'Campo parâmetros' deve ser preenchido !!!", vbOKOnly + vbExclamation, "Aviso"
                   Exit Function
                End If
                If .Cell(flexcpText, I, conCOL_SonParam_ValParamPos) = Empty Then
                   MsgBox "Á tabela 'Parâmetros' o registro " & I & " 'Campo parâmetros(+)' deve ser preenchido !!!", vbOKOnly + vbExclamation, "Aviso"
                   Exit Function
                End If
                If .Cell(flexcpText, I, conCOL_SonParam_ValParamNeg) = Empty Then
                   MsgBox "Á tabela 'Parâmetros' o registro " & I & " 'Campo parâmetros(-)' deve ser preenchido !!!", vbOKOnly + vbExclamation, "Aviso"
                   Exit Function
                End If
            Next I
        End With
     End If
     
     ValidaCampos = True
     
End Function

Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
    Frame3.Enabled = False
   
    Me.Caption = "Cadastro de ficha técnica de produtos - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADFICHATECNICA
    
    objCADFICHATEC.CODIGO = iCodigo
    
    Call InitGridParam
    Call InitGridFamUnidade
    Call InitGridVelocidade
    
    objCADFICHATEC.PreencheComboProd cboProd
    objCADFICHATEC.PreenchComboUnidade cboUidade
    
    lblCORRPRODPADR.Caption = ""
    
    If objCADFICHATEC.Carrega_campos = True Then
        
        txtCodigo.Text = objCADFICHATEC.CODIGO
        mskData.Text = Format(objCADFICHATEC.Data, "DD/MM/YYYY")
        
        txtMinutos.Text = objCADFICHATEC.MINUTOS
        
        txtCodProd.Text = objCADFICHATEC.CODPROD
        cboProd.ListIndex = -1
        For I = 0 To (cboProd.ListCount - 1)
            If Trim(Mid(cboProd.List(I), 1, 10)) = Trim(txtCodProd.Text) Then cboProd.ListIndex = I
        Next I
        
        If objCADFICHATEC.CODFAMMAQ > 0 Then
            Call objCADFICHATEC.PreenchComboFamMaquina(cboFAMMAQUINA, txtCodProd.Text)
            txtCODFAMMAQUINA.Text = objCADFICHATEC.CODFAMMAQ
            cboFAMMAQUINA.ListIndex = -1
            For I = 0 To (cboFAMMAQUINA.ListCount - 1)
                If cboFAMMAQUINA.ItemData(I) = objCADFICHATEC.CODFAMMAQ Then cboFAMMAQUINA.ListIndex = I
            Next I
           
            Call objCADFICHATEC.PreenchComboMaquina(cboMAQUINA, objCADFICHATEC.CODFAMMAQ)
            txtCODMAQUINA.Text = objCADFICHATEC.CODMAQ
            
            cboMAQUINA.ListIndex = -1
            For I = 0 To (cboMAQUINA.ListCount - 1)
                If cboMAQUINA.ItemData(I) = objCADFICHATEC.CODMAQ Then cboMAQUINA.ListIndex = I
            Next I
            cboUidade.ListIndex = -1
            For I = 0 To (cboUidade.ListCount - 1)
                If cboUidade.ItemData(I) = objCADFICHATEC.UNIDSETTUP Then cboUidade.ListIndex = I
            Next I
        End If
        
        Call CarregaGrid
        Call CarregaGridFamUnidade
        Call CarregaGridEficiencia
        
        Call PegaDadosDoProd(Trim(txtCodProd.Text))
  
    End If

End Sub

Private Sub CarregaGrid()

    Dim I As Integer
    arrGRDFICHATEC = objCADFICHATEC.ITENS
    
    If IsArray(arrGRDFICHATEC) Then
       For I = 1 To UBound(arrGRDFICHATEC)
           grdParamFicha.AddItem arrGRDFICHATEC(I, 1) & vbTab & _
                                 "" & vbTab & _
                                 PegaDescParametro(CLng(arrGRDFICHATEC(I, 1))) & vbTab & _
                                 arrGRDFICHATEC(I, 2) & vbTab & _
                                 Format(arrGRDFICHATEC(I, 3), "#,##0.00") & vbTab & _
                                 Format(arrGRDFICHATEC(I, 4), "#,##0.00") & vbTab & _
                                 Format(arrGRDFICHATEC(I, 5), "#,##0.00")
       Next I
    End If
    
End Sub

Private Function PegaDescParametro(lngCODIGO As Long) As String

    PegaDescParametro = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPARFICHA " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCODIGO
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaDescParametro = BREC!SGI_DESCRI
    BREC.Close
    
End Function

Private Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame3.Enabled = False
   
    Me.Caption = "Cadastro de ficha técnica de produtos - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADFICHATECNICA
    
    objCADFICHATEC.CODIGO = iCodigo
    objCADFICHATEC.PreenchComboUnidade cboUidade
    
    Call InitGridParam
    Call InitGridFamUnidade
    Call InitGridVelocidade
    
    objCADFICHATEC.PreencheComboProd cboProd
    
    lblCORRPRODPADR.Caption = ""
    
    If objCADFICHATEC.Carrega_campos = True Then
        
        txtCodigo.Text = objCADFICHATEC.CODIGO
        mskData.Text = Format(objCADFICHATEC.Data, "DD/MM/YYYY")
        
        txtCodProd.Text = objCADFICHATEC.CODPROD
        cboProd.ListIndex = -1
        For I = 0 To (cboProd.ListCount - 1)
            If Trim(Mid(cboProd.List(I), 1, 10)) = Trim(txtCodProd.Text) Then cboProd.ListIndex = I
        Next I
        
        If objCADFICHATEC.CODFAMMAQ > 0 Then
            Call objCADFICHATEC.PreenchComboFamMaquina(cboFAMMAQUINA, txtCodProd.Text)
            txtCODFAMMAQUINA.Text = objCADFICHATEC.CODFAMMAQ
            cboFAMMAQUINA.ListIndex = -1
            For I = 0 To (cboFAMMAQUINA.ListCount - 1)
                If cboFAMMAQUINA.ItemData(I) = objCADFICHATEC.CODFAMMAQ Then cboFAMMAQUINA.ListIndex = I
            Next I
           
            Call objCADFICHATEC.PreenchComboMaquina(cboMAQUINA, objCADFICHATEC.CODFAMMAQ)
            txtCODMAQUINA.Text = objCADFICHATEC.CODMAQ
            cboMAQUINA.ListIndex = -1
            For I = 0 To (cboMAQUINA.ListCount - 1)
                If cboMAQUINA.ItemData(I) = objCADFICHATEC.CODMAQ Then cboMAQUINA.ListIndex = I
            Next I
            cboUidade.ListIndex = -1
            For I = 0 To (cboUidade.ListCount - 1)
                If cboUidade.ItemData(I) = objCADFICHATEC.UNIDSETTUP Then cboUidade.ListIndex = I
            Next I
        End If
        
        Call CarregaGrid
        Call CarregaGridFamUnidade
        Call CarregaGridEficiencia
        
        Call PegaDadosDoProd(Trim(txtCodProd.Text))
        
    End If

End Sub


Private Sub PegaDadosDoProd(strProd As String)

    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = '" & Trim(strProd) & "'"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
       optPADRAO(BREC!SGI_PADRAO).Value = True
       
       lblCORRPRODPADR.Caption = Format(0, "#,###0.000")
       If BREC!SGI_PADRAO = 1 Then
          lblCORRPRODPADR.Caption = Format(1, "#,###0.000")
       Else
          If Not IsNull(BREC!SGI_CORRPRODPADR) Then lblCORRPRODPADR.Caption = Format(BREC!SGI_CORRPRODPADR, "#,###0.000")
       End If
       
    End If
    BREC.Close
    
End Sub



Private Sub txtMinutos_GotFocus()
    objBLBFunc.SelecionaCampos txtMinutos.Name, frmCADFICHATECNICA
End Sub

Private Sub txtMinutos_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtMinutos.Text
End Sub

Private Sub txtMinutos_Validate(Cancel As Boolean)

    Dim curSEGUNDOS As Currency
    
    If Len(Trim(txtMinutos.Text)) = 0 Then
       txtMinutos.Text = ""
       Exit Sub
    End If
    
    If Not IsNumeric(txtMinutos.Text) Then
       MsgBox "Somente é permitido números e pontos !!!", vbOKOnly + vbExclamation, "Aviso"
       txtMinutos.Text = ""
       txtMinutos.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    '' Converte em Segundos
    curSEGUNDOS = (60 * CCur(txtMinutos.Text))

End Sub




Private Sub InitGridFamUnidade()

    With grdUnidades
    
       .Cols = conColumnsIn_SonUnidades
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonUnidades_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonUnidades_CodUnidDe) = ""
       .ColDataType(conCOL_SonUnidades_CodUnidDe) = flexDTLong
      
       .Cell(flexcpData, 0, conCOL_SonUnidades_PesqUnidDe) = ""
       .ColDataType(conCOL_SonUnidades_PesqUnidDe) = flexDTString
       .ColComboList(conCOL_SonUnidades_PesqUnidDe) = "..."
       
       .Cell(flexcpData, 0, conCOL_SonUnidades_UnidDe) = ""
       .ColDataType(conCOL_SonUnidades_UnidDe) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonUnidades_Desc_UnidDe) = ""
       .ColDataType(conCOL_SonUnidades_Desc_UnidDe) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonUnidades_CodUnidPara) = ""
       .ColDataType(conCOL_SonUnidades_CodUnidPara) = flexDTLong
      
       .Cell(flexcpData, 0, conCOL_SonUnidades_PesqUnidPara) = ""
       .ColDataType(conCOL_SonUnidades_PesqUnidPara) = flexDTString
       .ColComboList(conCOL_SonUnidades_PesqUnidPara) = "..."
       
       .Cell(flexcpData, 0, conCOL_SonUnidades_UnidPara) = ""
       .ColDataType(conCOL_SonUnidades_UnidPara) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonUnidades_Desc_UnidPara) = ""
       .ColDataType(conCOL_SonUnidades_Desc_UnidPara) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonUnidades_FamDe) = ""
       .ColDataType(conCOL_SonUnidades_FamDe) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonUnidades_FamPara) = ""
       .ColDataType(conCOL_SonUnidades_FamPara) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonUnidades_Indice) = ""
       .ColDataType(conCOL_SonUnidades_Indice) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonUnidades_Default) = ""
       .ColDataType(conCOL_SonUnidades_Default) = flexDTBoolean
       .ColFormat(conCOL_SonUnidades_Default) = "Sim;Não"
       
       .ColWidth(conCOL_SonUnidades_CodUnidDe) = 800
       .ColWidth(conCOL_SonUnidades_PesqUnidDe) = 300
       .ColWidth(conCOL_SonUnidades_UnidDe) = 500
       .ColWidth(conCOL_SonUnidades_Desc_UnidDe) = 2500
       
       .ColWidth(conCOL_SonUnidades_CodUnidPara) = 800
       .ColWidth(conCOL_SonUnidades_PesqUnidPara) = 300
       .ColWidth(conCOL_SonUnidades_UnidPara) = 500
       .ColWidth(conCOL_SonUnidades_Desc_UnidPara) = 2500
       
       .ColWidth(conCOL_SonUnidades_FamDe) = 800
       .ColWidth(conCOL_SonUnidades_FamPara) = 800
       .ColWidth(conCOL_SonUnidades_Indice) = 800
       .ColWidth(conCOL_SonUnidades_Default) = 800
       
       .ColHidden(conCOL_SonUnidades_FamDe) = True
       .ColHidden(conCOL_SonUnidades_FamPara) = True
       .ColHidden(conCOL_SonUnidades_Indice) = True
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub IncFamUnidade()
   
    If ExisteLinhaVaziaUnidade = False Then Exit Sub
    
    grdUnidades.AddItem "" & vbTab & _
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

End Sub


Private Sub CarregaGridFamUnidade()

    Dim I As Integer
    arrFAMMEDIDAS = objCADFICHATEC.FAMUNIDADE
    
    If IsArray(arrFAMMEDIDAS) Then
       For I = 1 To UBound(arrFAMMEDIDAS)
           grdUnidades.AddItem arrFAMMEDIDAS(I, 1) & vbTab & _
                               "" & vbTab & _
                               PegaUnidade(CLng(arrFAMMEDIDAS(I, 1))) & vbTab & _
                               PegaDescrFamUnidades(CLng(arrFAMMEDIDAS(I, 1))) & vbTab & _
                               arrFAMMEDIDAS(I, 2) & vbTab & _
                               "" & vbTab & _
                               PegaUnidade(CLng(arrFAMMEDIDAS(I, 2))) & vbTab & _
                               PegaDescrFamUnidades(CLng(arrFAMMEDIDAS(I, 2))) & vbTab & _
                               arrFAMMEDIDAS(I, 3) & vbTab & _
                               PegaCodFamUnidades(CLng(arrFAMMEDIDAS(I, 1))) & vbTab & _
                               PegaCodFamUnidades(CLng(arrFAMMEDIDAS(I, 2))) & vbTab & _
                               Trim(arrFAMMEDIDAS(I, 1)) & Trim(arrFAMMEDIDAS(I, 2))
                               
       Next I
    End If
    
End Sub


Private Function PegaCodFamUnidades(lngCodUsuario As Long) As String
    PegaCodFamUnidades = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADUNIMED " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCodUsuario
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaCodFamUnidades = BREC!SGI_CODFAMUNID
    BREC.Close
    
End Function


Private Sub InitGridVelocidade()

    With grdVelocidade
    
       .Cols = conColumnsIn_SonVelocidade
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonVelocidade_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonVelicidade_EficMed) = ""
       .ColDataType(conCOL_SonVelicidade_EficMed) = flexDTCurrency
      
       .Cell(flexcpData, 0, conCOL_SonVelocidade_ProdTeorica) = ""
       .ColDataType(conCOL_SonVelocidade_ProdTeorica) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonVelocidade_ProdReal) = ""
       .ColDataType(conCOL_SonVelocidade_ProdReal) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonVelocidade_Unidade) = ""
       .ColDataType(conCOL_SonVelocidade_Unidade) = flexDTString
       
       .ColWidth(conCOL_SonVelicidade_EficMed) = 1500
       .ColWidth(conCOL_SonVelocidade_ProdTeorica) = 1500
       .ColWidth(conCOL_SonVelocidade_ProdReal) = 1500
       .ColWidth(conCOL_SonVelocidade_Unidade) = 1500
       
       .ColHidden(conCOL_SonVelocidade_Indice) = True
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub IncVelociadade()
   
    With grdUnidades
        If (.Rows - 1) = 0 Then
           MsgBox "Primeiro insira a unidade de velocidade !!!", vbOKOnly + vbCritical, "Aviso"
           Exit Sub
        End If
        If (.Row) = 0 Then
           MsgBox "Primeiro Selecione a unidade de velocidade !!!", vbOKOnly + vbCritical, "Aviso"
           Exit Sub
        End If
        If .Cell(flexcpText, .Row, conCOL_SonUnidades_CodUnidDe) = Empty Or _
           .Cell(flexcpText, .Row, conCOL_SonUnidades_CodUnidPara) = Empty Then
           MsgBox "Primeiro informe as unidades de medidas !!!", vbOKOnly + vbCritical, "Aviso"
           Exit Sub
        End If
        
        If ConfLinha(.Cell(flexcpText, .Row, conCOL_SonUnidades_Indice)) > 0 Then Exit Sub
    
        grdVelocidade.AddItem "" & vbTab & _
                              "" & vbTab & _
                              "" & vbTab & _
                              Trim(.Cell(flexcpText, .Row, conCOL_SonUnidades_UnidDe)) & "/" & Trim(.Cell(flexcpText, .Row, conCOL_SonUnidades_UnidPara)) & vbTab & _
                              .Cell(flexcpText, .Row, conCOL_SonUnidades_Indice)
        
        Call PosRegGrdEficiencia(.Cell(flexcpText, .Row, conCOL_SonUnidades_Indice))
                          
    End With
End Sub


Private Function ConfLinha(strINDICE As String) As Long
    ConfLinha = 0
        
    Dim I As Integer
    
    With grdVelocidade
        For I = 1 To (.Rows - 1)
            If .Cell(flexcpText, I, conCOL_SonVelocidade_Indice) = strINDICE Then ConfLinha = (ConfLinha + 1)
        Next I
    End With

End Function

Private Sub PosRegGrdEficiencia(strINDICE As String)
    Dim I As Integer
    With grdVelocidade
         For I = 1 To (.Rows - 1)
             If .Cell(flexcpText, I, conCOL_SonVelocidade_Indice) <> strINDICE Then
                .RowHidden(I) = True
             Else
                .RowHidden(I) = False
             End If
         Next I
    End With
End Sub


Private Function CalcPcRealGrid(strValor1 As String, strValor2 As String) As Currency

    CalcPcRealGrid = 0
    
    If Len(Trim(strValor1)) = 0 Or Len(Trim(strValor2)) = 0 Then Exit Function
    If CCur(strValor1) = 0 Or CCur(strValor2) = 0 Then Exit Function
    
    CalcPcRealGrid = (CCur(strValor1) / 100) * CCur(strValor2)
    
End Function



Private Sub CarregaGridEficiencia()

    Dim I As Integer
    Dim J As Integer
    arrEFICIENCIA = objCADFICHATEC.EFICIENCIA
    
    If IsArray(arrEFICIENCIA) Then
       For I = 1 To UBound(arrEFICIENCIA)
           
           grdVelocidade.AddItem Format(arrEFICIENCIA(I, 1), "#,##0.00") & vbTab & _
                                 Format(arrEFICIENCIA(I, 2), "#,##0.00") & vbTab & _
                                 Format(arrEFICIENCIA(I, 3), "#,##0.00") & vbTab & _
                                 "" & vbTab & _
                                 arrEFICIENCIA(I, 4)
           
           J = grdUnidades.FindRow(arrEFICIENCIA(I, 4), , conCOL_SonUnidades_Indice)
           If J <> -1 Then
              grdVelocidade.Cell(flexcpText, I, conCOL_SonVelocidade_Unidade) = Trim(grdUnidades.Cell(flexcpText, J, conCOL_SonUnidades_UnidDe)) & "/" & Trim(grdUnidades.Cell(flexcpText, J, conCOL_SonUnidades_UnidPara))
           End If
           
       Next I
       
       If (grdUnidades.Rows - 1) > 0 Then
          grdUnidades.Row = 1
          Call grdUnidades_Click
       End If
    End If
    
End Sub

