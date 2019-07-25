VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCADOORDFABP 
   Caption         =   "Cadastro de Ordem de Fabricação"
   ClientHeight    =   9225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17505
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   17505
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   7095
      Left            =   0
      TabIndex        =   17
      Top             =   2040
      Width           =   17415
      Begin VB.Frame Frame4 
         Caption         =   "[ Legenda ]"
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
         Height          =   855
         Left            =   120
         TabIndex        =   22
         Top             =   6120
         Width           =   17175
         Begin VB.Frame Frame5 
            Caption         =   "Seleciona Apenas OP's sem impressão"
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
            Left            =   13320
            TabIndex        =   35
            Top             =   120
            Width           =   3615
            Begin VB.OptionButton optSelOPNaoImpSN 
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
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   1
               Left            =   720
               TabIndex        =   37
               Top             =   240
               Width           =   855
            End
            Begin VB.OptionButton optSelOPNaoImpSN 
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
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   0
               Left            =   1920
               TabIndex        =   36
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Seleciona Todos os Registros"
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
            Left            =   9960
            TabIndex        =   32
            Top             =   120
            Width           =   2775
            Begin VB.OptionButton optSelecTodosSN 
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
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   0
               Left            =   1440
               TabIndex        =   34
               Top             =   240
               Width           =   735
            End
            Begin VB.OptionButton optSelecTodosSN 
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
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   1
               Left            =   480
               TabIndex        =   33
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "[ Tipo de OP ]"
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
            Left            =   6360
            TabIndex        =   29
            Top             =   120
            Width           =   2655
            Begin VB.OptionButton optTPOP 
               Caption         =   "Homologada"
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
               TabIndex        =   31
               Top             =   240
               Width           =   1455
            End
            Begin VB.OptionButton optTPOP 
               Caption         =   "Normal"
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
               TabIndex        =   30
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Label Label3 
            BackColor       =   &H0000FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00008000&
            Height          =   255
            Index           =   2
            Left            =   3480
            TabIndex        =   39
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label4 
            Caption         =   "Alteradas"
            Height          =   255
            Index           =   2
            Left            =   3840
            TabIndex        =   38
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Ainda sem Impressão"
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   26
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00008000&
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   25
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label4 
            Caption         =   "Já Impressa"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   24
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label3 
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00008000&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   255
         End
      End
      Begin TabDlg.SSTab stOrdem 
         Height          =   5895
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   17175
         _ExtentX        =   30295
         _ExtentY        =   10398
         _Version        =   393216
         Style           =   1
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
         TabCaption(0)   =   "Em Aberto"
         TabPicture(0)   =   "frmCADOORDFABP.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "grdOrdeFab"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Parcial"
         TabPicture(1)   =   "frmCADOORDFABP.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "grdPARCIAL"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Finalizado"
         TabPicture(2)   =   "frmCADOORDFABP.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "grdFINALIZADO"
         Tab(2).ControlCount=   1
         Begin VSFlex8LCtl.VSFlexGrid grdPARCIAL 
            Height          =   5415
            Left            =   -74880
            TabIndex        =   21
            Top             =   360
            Width           =   16935
            _cx             =   29871
            _cy             =   9551
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
         Begin VSFlex8LCtl.VSFlexGrid grdOrdeFab 
            Height          =   5415
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   16935
            _cx             =   29871
            _cy             =   9551
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
         Begin VSFlex8LCtl.VSFlexGrid grdFINALIZADO 
            Height          =   5415
            Left            =   -74880
            TabIndex        =   20
            Top             =   360
            Width           =   16935
            _cx             =   29871
            _cy             =   9551
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
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   17415
      Begin VB.TextBox txtCODROTULO 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   15000
         TabIndex        =   5
         Text            =   "txtCODROTULO"
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtRAZAOSOC 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10440
         TabIndex        =   7
         Text            =   "txtRAZAOSOC"
         Top             =   840
         Width           =   6735
      End
      Begin VB.TextBox txtCODCLIE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   11640
         TabIndex        =   4
         Text            =   "txtCODCLIE"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtRotulo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Text            =   "txtRotulo"
         Top             =   840
         Width           =   6375
      End
      Begin VB.Frame Frame8 
         Caption         =   "[ Data da OP ]"
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
         Left            =   5760
         TabIndex        =   42
         Top             =   120
         Width           =   3495
         Begin MSMask.MaskEdBox mskDTINI 
            Height          =   285
            Left            =   480
            TabIndex        =   2
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
         Begin MSMask.MaskEdBox mskDTFIN 
            Height          =   285
            Left            =   2040
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
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "á"
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
            Left            =   1800
            TabIndex        =   44
            Top             =   240
            Width           =   120
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "de"
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
            TabIndex        =   43
            Top             =   240
            Width           =   225
         End
      End
      Begin VB.TextBox txtCODPED 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         TabIndex        =   1
         Text            =   "txtCODPED"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtCODOP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   0
         Text            =   "txtCODOP"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Código do Rótulo"
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
         Index           =   7
         Left            =   13320
         TabIndex        =   48
         Top             =   360
         Width           =   1485
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Razão Social"
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
         Index           =   6
         Left            =   9000
         TabIndex        =   47
         Top             =   840
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Código do Cliente"
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
         Left            =   9840
         TabIndex        =   46
         Top             =   360
         Width           =   1515
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Desc. Rótulo"
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
         Left            =   120
         TabIndex        =   45
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "nº Pedido"
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
         Left            =   2040
         TabIndex        =   41
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "nº OP"
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
         TabIndex        =   40
         Top             =   360
         Width           =   510
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   8
      Top             =   1200
      Width           =   17415
      Begin VB.CommandButton cmdImpressao 
         Caption         =   "Im&prime <F5>"
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
         Index           =   1
         Left            =   5160
         Picture         =   "frmCADOORDFABP.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Imprime Registro"
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Liquida"
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
         Left            =   14160
         Picture         =   "frmCADOORDFABP.frx":0156
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Liquida a Ordem de Produção"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdImpressao 
         Caption         =   "&Visualiza <F5>"
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
         Index           =   0
         Left            =   3600
         Picture         =   "frmCADOORDFABP.frx":0258
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Imprime Registro"
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmdVoltar 
         Caption         =   "&Voltar"
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
         Picture         =   "frmCADOORDFABP.frx":035A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Volta ao Menu Principal"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdInclui 
         Caption         =   "&Inclui"
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
         Picture         =   "frmCADOORDFABP.frx":088C
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Inclui um novo registro"
         Top             =   120
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
         Left            =   1800
         Picture         =   "frmCADOORDFABP.frx":0DBE
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Altera Registro"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdExclui 
         Caption         =   "&Exclui"
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
         Left            =   2640
         Picture         =   "frmCADOORDFABP.frx":0EC0
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Exclui Registro"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdCanFiltro 
         Caption         =   "Desfas"
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
         Left            =   15840
         Picture         =   "frmCADOORDFABP.frx":0FC2
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Desfas Ultima Pesqusa"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdOrden 
         Caption         =   "Ordem"
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
         Left            =   16560
         Picture         =   "frmCADOORDFABP.frx":14F4
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Ordena os Registros"
         Top             =   120
         Width           =   735
      End
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   7320
         Top             =   120
      End
   End
End
Attribute VB_Name = "frmCADOORDFABP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public FILIAL           As Integer
Public strAcesso        As String
Public strUsuario       As String
Public lngCodUsuaro     As Long
Public intFILIALPED     As Integer

Dim lngCodVendedor      As Long
Dim objFuncoes          As Object
Dim objCADOORDFAB       As Object
Dim objRel              As Object
Dim objRelDireto        As Object
Dim iCodigo             As Long
Dim strNOMTABELA1       As String
Dim strNOMTABELA2       As String
Dim strNomModulo        As String
Dim strNOMFILIAL        As String


Const conCOL_OrdFab_Selecionado                     As Integer = 0
Const conCOL_OrdFab_Codigo                          As Integer = 1
Const conCOL_OrdFab_Pedido                          As Integer = 2
Const conCOL_OrdFab_DataOrdem                       As Integer = 3
Const conCOL_OrdFab_Rotulo                          As Integer = 4
Const conCOL_OrdFab_DescRotulo                      As Integer = 5
Const conCOL_OrdFab_CodCliente                      As Integer = 6
Const conCOL_OrdFab_Cliente                         As Integer = 7
Const conCOL_OrdFab_IDProduto                       As Integer = 8
Const conCOL_OrdFab_Tipo                            As Integer = 9
Const conCOL_OrdFab_DescTipo                        As Integer = 10
Const conCOL_OrdFab_TipoDaOP                        As Integer = 11
Const conCOL_OrdFab_JaImpressa                      As Integer = 12
Const conCOL_OrdFab_FilialEmp                       As Integer = 13
Const conCOL_OrdFab_NomeFilialEmp                   As Integer = 14
Const conCOL_OrdFab_FoiAlterado                     As Integer = 15
Const conCOL_OrdFab_FormatString                    As String = "=  |Nº Ordem|Nº Pedido|Dt.Ordem|Rótulo|Desc.Rótulo|Cód.Cliente|Cliente|ID_Produto|Tipo|Desc.Tipo|Cod.Tipo|Já.Imp|CodFilial|Filial|FoiAlterado"
Const conColumnsIn_OrdFab                           As Integer = 16

Const conCOL_OrdFabParc_Selecionado                 As Integer = 0
Const conCOL_OrdFabParc_Codigo                      As Integer = 1
Const conCOL_OrdFabParc_Pedido                      As Integer = 2
Const conCOL_OrdFabParc_DataOrdem                   As Integer = 3
Const conCOL_OrdFabParc_Rotulo                      As Integer = 4
Const conCOL_OrdFabParc_DescRotulo                  As Integer = 5
Const conCOL_OrdFabParc_CodCliente                  As Integer = 6
Const conCOL_OrdFabParc_Cliente                     As Integer = 7
Const conCOL_OrdFabParc_IDProduto                   As Integer = 8
Const conCOL_OrdFabParc_Tipo                        As Integer = 9
Const conCOL_OrdFabParc_DescTipo                    As Integer = 10
Const conCOL_OrdFabParc_TipoDaOP                    As Integer = 11
Const conCOL_OrdFabParc_JaImpressa                  As Integer = 12
Const conCOL_OrdFabParc_FilialEmp                   As Integer = 13
Const conCOL_OrdFabParc_NomeFilialEmp               As Integer = 14
Const conCOL_OrdFabParc_FoiAlterado                 As Integer = 15
Const conCOL_OrdFabParc_FormatString                As String = "=  |Nº Ordem|Nº Pedido|Dt.Ordem|Rótulo|Desc.Rótulo|Cód.Cliente|Cliente|ID_Produto|Tipo|Desc.Tipo|Cod.Tipo|Já.Imp|CodFilial|Filial|FoiAlterado"
Const conColumnsIn_OrdFabParc                       As Integer = 16

Const conCOL_OrdFabFin_Selecionado                  As Integer = 0
Const conCOL_OrdFabFin_Codigo                       As Integer = 1
Const conCOL_OrdFabFin_Pedido                       As Integer = 2
Const conCOL_OrdFabFin_DataOrdem                    As Integer = 3
Const conCOL_OrdFabFin_Rotulo                       As Integer = 4
Const conCOL_OrdFabFin_DescRotulo                   As Integer = 5
Const conCOL_OrdFabFin_CodCliente                   As Integer = 6
Const conCOL_OrdFabFin_Cliente                      As Integer = 7
Const conCOL_OrdFabFin_IDProduto                    As Integer = 8
Const conCOL_OrdFabFin_Tipo                         As Integer = 9
Const conCOL_OrdFabFin_DescTipo                     As Integer = 10
Const conCOL_OrdFabFin_TipoDaOP                     As Integer = 11
Const conCOL_OrdFabFin_JaImpressa                   As Integer = 12
Const conCOL_OrdFabFin_FilialEmp                    As Integer = 13
Const conCOL_OrdFabFin_NomeFilialEmp                As Integer = 14
Const conCOL_OrdFabFin_FoiAlterado                  As Integer = 15
Const conCOL_OrdFabFin_FormatString                 As String = "=  |Nº Ordem|Nº Pedido|Dt.Ordem|Rótulo|Desc.Rótulo|Cód.Cliente|Cliente|ID_Produto|Tipo|Desc.Tipo|Cod.Tipo|Já.Imp|CodFilial|Filial|FoiAlterado"
Const conColumnsIn_OrdFabFin                        As Integer = 16


Private Sub cmdAltera_Click()
    If objFuncoes.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    Call Operacao("A")
End Sub

Private Sub cmdCanFiltro_Click()
   Call ConfGridOrdFab
   Call ConfGridOrdFabParc
   Call ConfGridOrdFabFin
End Sub

Private Sub cmdExclui_Click()

    If stOrdem.Tab = 1 Or _
       stOrdem.Tab = 2 Then
       MsgBox "Não pode ser excluida nesta opção !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Sub
    End If
    
    If objFuncoes.ChecaAcesso2("E", strAcesso) = False Then Exit Sub
    
    If VerificaOrdFat(Str(objCADOORDFAB.CODORDEM)) = True Then
        MsgBox "Existe ordem de faturamento para esta ordem de produção, não pode ser excluida !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    Dim iResp     As Integer
    Dim lngCodLog As Long
    
    iResp = MsgBox("Confirma a exclusão do registro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
    If iResp <> 6 Then Exit Sub
    
    objCADOORDFAB.IdProduto = CLng(grdOrdeFab.Cell(flexcpText, grdOrdeFab.Row, conCOL_OrdFab_IDProduto))
    objCADOORDFAB.CODPEDIDO = CLng(grdOrdeFab.Cell(flexcpText, grdOrdeFab.Row, conCOL_OrdFab_Pedido))
    
    If objCADOORDFAB.GRAVA("E", Trim(strNOMTABELA1), "") = False Then Exit Sub
    lngCodLog = objFuncoes.Gera_Codigo("SGI_LOGMODULO", FILIAL, Linha)
    Call objFuncoes.GravaLogModulo(FILIAL, lngCodLog, Trim(strNomModulo), "E", lngCodUsuaro, Str(objCADOORDFAB.CODORDEM), Linha)
    
    MsgBox "Registro excluso com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
  
    Call AbilitaCampos
    Call ConfGridOrdFab
    Call ConfGridOrdFabParc
    Call ConfGridOrdFabFin

End Sub

Private Sub cmdImpressao_Click(index As Integer)
    
    Dim I           As Long
    Dim lngCOL1     As Long
    Dim lngCOL2     As Long
    Dim lngCOL3     As Long
    Dim grdGENERICA As Variant
    Dim intTime     As Long
    Dim strCODORD   As String
    
    If stOrdem.Tab = 0 Then
        Set grdGENERICA = grdOrdeFab
        lngCOL1 = conCOL_OrdFab_Selecionado
        lngCOL2 = conCOL_OrdFab_JaImpressa
        lngCOL3 = conCOL_OrdFab_Codigo
    ElseIf stOrdem.Tab = 1 Then
        Set grdGENERICA = grdPARCIAL
        lngCOL1 = conCOL_OrdFabParc_Selecionado
        lngCOL2 = conCOL_OrdFabParc_JaImpressa
        lngCOL3 = conCOL_OrdFabParc_Codigo
    ElseIf stOrdem.Tab = 2 Then
        Set grdGENERICA = grdFINALIZADO
        lngCOL1 = conCOL_OrdFabFin_Selecionado
        lngCOL2 = conCOL_OrdFabFin_JaImpressa
        lngCOL3 = conCOL_OrdFabFin_Codigo
    End If
    
    With grdGENERICA
        If (.Rows - 1) = 0 Then
            MsgBox "Não há dados para a impressão !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
        End If
        If .Row = 0 Then
           MsgBox "Selecione um registro !!!", vbOKOnly + vbExclamation, "Aviso"
           Exit Sub
        End If
        If index = 0 Then
            For I = 1 To (.Rows - 1)
                If .Cell(flexcpText, I, lngCOL1) = -1 Then
                    objCADOORDFAB.CODORDEM = CLng(.Cell(flexcpText, I, lngCOL3))
                    strCODORD = strCODORD & Trim(.Cell(flexcpText, I, lngCOL3)) & ","
                    .Cell(flexcpText, I, lngCOL1) = 0
                    .Cell(flexcpText, I, lngCOL2) = "SIM"
                    Call MudaCorCelula(grdGENERICA, 1, I)
                End If
            Next I
            If Len(Trim(strCODORD)) > 0 Then
               strCODORD = Mid(strCODORD, 1, (Len(strCODORD) - 1))
               Call Imprime(index, strCODORD)
            End If
        End If
    End With
    
    Set grdGENERICA = Nothing

    optSelecTodosSN(0).Value = True
    
End Sub

Private Sub cmdInclui_Click()
    If objFuncoes.ChecaAcesso2("I", strAcesso) = False Then Exit Sub
    Call Operacao("I")
End Sub

Private Sub cmdOrden_Click()
    Call Ordem
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    
    Dim grdGENERICO As VSFlexGrid
    Dim lngCOL      As Long
    Dim lngCOL2     As Long
    
    If stOrdem.Tab = 2 Then
        MsgBox "Opção inválida !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    If objFuncoes.ChecaAcesso2("E", strAcesso) = False Then Exit Sub
    Call Operacao("BX")
    
    Exit Sub

End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then cmdVoltar_Click
   If KeyCode = vbKeyF5 Then Call cmdImpressao_Click(0)

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
    Set objCADOORDFAB = CreateObject("CADOORDFAB.clsCADOORDFAB")
    Set objRel = CreateObject("MOSTRAREL.clsMOSTRAREL")
 
    objCADOORDFAB.FILIAL = FILIAL
    objFuncoes.LimpaCampos Me
    
    If adoBanco_Dados.State = 0 Then Set adoBanco_Dados = objFuncoes.Banco_Dados(Linha)
    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    Call ConfGridOrdFab
    Call ConfGridOrdFabParc
    Call ConfGridOrdFabFin
     
    stOrdem.Tab = 0
   
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    
    strNOMFILIAL = ""
    If intFILIALPED = 0 Then
       Me.Caption = Me.Caption & " / NOVALATA"
       strNOMTABELA1 = "SGI_ORDEMPROD"
       strNOMTABELA2 = "SGI_CADPEDVENDH"
       strNomModulo = "frmCADOORDFAB"
    ElseIf intFILIALPED = 1 Then
       strNOMFILIAL = "_STEEL"
       Me.Caption = Me.Caption & " / STEEL ROLL"
       strNOMTABELA1 = "SGI_ORDEMPROD_STEEL"
       strNOMTABELA2 = "SGI_CADPEDVENDH_STEEL"
       strNomModulo = "frmCADOORDFAB_STEEL"
    End If
    
    Call AbilitaCampos
    optTPOP(0).Value = True
    
    optSelecTodosSN(0).Value = True
    optSelOPNaoImpSN(1).Value = True
    

End Sub

Private Sub AbilitaCampos()

    Dim boolAtivoDesativo As Boolean
    
    boolAtivoDesativo = objCADOORDFAB.AtivoDesativo(strNOMTABELA1)
    
    cmdAltera.Enabled = boolAtivoDesativo
    cmdExclui.Enabled = boolAtivoDesativo
    cmdImpressao(0).Enabled = boolAtivoDesativo
    cmdImpressao(1).Enabled = boolAtivoDesativo
    
    Frame1.Enabled = boolAtivoDesativo
    Frame3.Enabled = boolAtivoDesativo

End Sub

Private Sub ConfGridOrdFab()

    With grdOrdeFab
    
       .Cols = conColumnsIn_OrdFab
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_OrdFab_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       
       .Cell(flexcpData, 0, conCOL_OrdFab_Selecionado) = ""
       .ColDataType(conCOL_OrdFab_Selecionado) = flexDTBoolean
       
       .Cell(flexcpData, 0, conCOL_OrdFab_Codigo) = ""
       .ColDataType(conCOL_OrdFab_Codigo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdFab_Pedido) = ""
       .ColDataType(conCOL_OrdFab_Pedido) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdFab_DataOrdem) = ""
       .ColDataType(conCOL_OrdFab_DataOrdem) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_OrdFab_Rotulo) = ""
       .ColDataType(conCOL_OrdFab_Rotulo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_OrdFab_DescRotulo) = ""
       .ColDataType(conCOL_OrdFab_DescRotulo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_OrdFab_CodCliente) = ""
       .ColDataType(conCOL_OrdFab_CodCliente) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdFab_Cliente) = ""
       .ColDataType(conCOL_OrdFab_Cliente) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_OrdFab_IDProduto) = ""
       .ColDataType(conCOL_OrdFab_IDProduto) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdFab_Tipo) = ""
       .ColDataType(conCOL_OrdFab_Tipo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdFab_DescTipo) = ""
       .ColDataType(conCOL_OrdFab_DescTipo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_OrdFab_TipoDaOP) = ""
       .ColDataType(conCOL_OrdFab_TipoDaOP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdFab_JaImpressa) = ""
       .ColDataType(conCOL_OrdFab_JaImpressa) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdFab_FilialEmp) = ""
       .ColDataType(conCOL_OrdFab_FilialEmp) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdFab_NomeFilialEmp) = ""
       .ColDataType(conCOL_OrdFab_NomeFilialEmp) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_OrdFab_FoiAlterado) = ""
       .ColDataType(conCOL_OrdFab_FoiAlterado) = flexDTLong
       
       .ColWidth(conCOL_OrdFab_Selecionado) = 300
       .ColWidth(conCOL_OrdFab_Codigo) = 1000
       .ColWidth(conCOL_OrdFab_Pedido) = 1000
       .ColWidth(conCOL_OrdFab_DataOrdem) = 1000
       .ColWidth(conCOL_OrdFab_Rotulo) = 1200
       .ColWidth(conCOL_OrdFab_DescRotulo) = 4500
       .ColWidth(conCOL_OrdFab_CodCliente) = 1000
       .ColWidth(conCOL_OrdFab_Cliente) = 4500
       .ColWidth(conCOL_OrdFab_IDProduto) = 0
       .ColWidth(conCOL_OrdFab_Tipo) = 0
       .ColWidth(conCOL_OrdFab_DescTipo) = 1400
       .ColWidth(conCOL_OrdFab_TipoDaOP) = 0
       .ColWidth(conCOL_OrdFab_JaImpressa) = 700
       .ColWidth(conCOL_OrdFab_FilialEmp) = 0
       .ColWidth(conCOL_OrdFab_NomeFilialEmp) = 0
       .ColWidth(conCOL_OrdFab_FoiAlterado) = 0
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub Operacao(strOperacao As String)
  
    Dim grdGENERICA As VSFlexGrid
    
    If stOrdem.Tab = 0 Then
        With grdOrdeFab
            If (.Rows - 1) > 0 And .Row > 0 Then iCodigo = CLng(.Cell(flexcpText, .Row, conCOL_OrdFab_Codigo))
        End With
    ElseIf stOrdem.Tab = 1 Then
        If strOperacao = "A" Then
            MsgBox "Esta Ordem Não pode Ser Alterada !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
        End If
        With grdPARCIAL
            If (.Rows - 1) > 0 And .Row > 0 Then iCodigo = CLng(.Cell(flexcpText, .Row, conCOL_OrdFabParc_Codigo))
        End With
    ElseIf stOrdem.Tab = 2 Then
        If strOperacao = "A" Then
            MsgBox "Esta Ordem Não pode Ser Alterada !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Sub
        End If
        With grdFINALIZADO
            If (.Rows - 1) > 0 And .Row > 0 Then iCodigo = CLng(.Cell(flexcpText, .Row, conCOL_OrdFabFin_Codigo))
        End With
    End If
    
    If strOperacao = "A" Then
       If objCADOORDFAB.SomaQtdeItens(Str(iCodigo), strNOMTABELA1) > 1 Then
          MsgBox "ATENÇÃO - Esta Ordem esta desmembrada em mais de uma ordem !!!", vbOKOnly + vbExclamation, "Aviso"
          Exit Sub
       End If
    End If
    
    frmCADOORDFAB.cCaminho = cCaminho
    frmCADOORDFAB.Linha = Linha
    frmCADOORDFAB.iCodigo = iCodigo
    frmCADOORDFAB.cTipOper = strOperacao
    frmCADOORDFAB.FILIAL = FILIAL
    frmCADOORDFAB.strAcesso = strAcesso
    frmCADOORDFAB.strMODPAI = Me.Name
    frmCADOORDFAB.strUsuario = strUsuario
    frmCADOORDFAB.lngCodVendedor = lngCodVendedor
    frmCADOORDFAB.lngCodUsuario = lngCodUsuaro
    frmCADOORDFAB.intFILIALPED = intFILIALPED
    frmCADOORDFAB.Show vbModal
    
    Call AbilitaCampos
    Call ConfGridOrdFab
    Call ConfGridOrdFabParc
    Call ConfGridOrdFabFin

End Sub

Private Sub PreencheGridGerado()

    With grdOrdeFab
    
        sSql = "Select " & vbCrLf
        
        sSql = sSql & "       ORD.SGI_CODIGO " & vbCrLf
        sSql = sSql & "      ,ORD.SGI_CODPED " & vbCrLf
        sSql = sSql & "      ,ORD.SGI_DATAORDEM " & vbCrLf
        sSql = sSql & "      ,ORD.SGI_CODPROD " & vbCrLf
        sSql = sSql & "      ,ORD.SGI_TIPOP " & vbCrLf
        sSql = sSql & "      ,PED.SGI_CODCLI " & vbCrLf
        sSql = sSql & "      ,CLI.SGI_RAZAOSOC " & vbCrLf
        sSql = sSql & "      ,ORD.SGI_IDPRODUTO " & vbCrLf
        sSql = sSql & "      ,PROD.SGI_CODTIPO " & vbCrLf
        sSql = sSql & "      ,ORD.SGI_JAIMPRESSA " & vbCrLf
        
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_ORDEMPROD   ORD " & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDH PED " & vbCrLf
        sSql = sSql & "      ,SGI_CADCLIENTE  CLI " & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO  PROD " & vbCrLf
        
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       ORD.SGI_FILIAL     = " & FILIAL & vbCrLf
        sSql = sSql & "   And ORD.SGI_STATUS     = 0" & vbCrLf
        sSql = sSql & "   And PED.SGI_FILIAL     = ORD.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And PED.SGI_CODIGO     = ORD.SGI_CODPED " & vbCrLf
        sSql = sSql & "   And CLI.SGI_FILIAL     = PED.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And CLI.SGI_CODIGO     = PED.SGI_CODCLI " & vbCrLf
        sSql = sSql & "   And PROD.SGI_FILIAL    = ORD.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And PROD.SGI_IDPRODUTO = ORD.SGI_IDPRODUTO " & vbCrLf
        sSql = sSql & "Order By ORD.SGI_CODIGO "
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        Do While Not BREC.EOF()
        
            .AddItem BREC!SGI_CODIGO & vbTab & _
                     BREC!SGI_CODPED & vbTab & _
                     Format(BREC!SGI_DATAORDEM, "DD/MM/YYYY") & vbTab & _
                     Trim(BREC!SGI_CODPROD) & vbTab & _
                     BREC!SGI_CODCLI & vbTab & _
                     Trim(BREC!SGI_RAZAOSOC) & vbTab & _
                     BREC!SGI_IDPRODUTO & vbTab & _
                     BREC!SGI_CODTIPO & vbTab & _
                     IIf(BREC!SGI_CODTIPO = 1, "NORMAL", "HOMOLOGADA") & vbTab & _
                     BREC!SGI_TIPOP & vbTab & _
                     IIf(BREC!SGI_JAIMPRESSA = 1, "SIM", "NÃO")
        
            Call MudaCorCelula(grdOrdeFab, BREC!SGI_JAIMPRESSA, (.Rows - 1))
            
            BREC.MoveNext
        Loop
        BREC.Close

    End With
End Sub

Private Sub ConfGridOrdFabFin()

    With grdFINALIZADO
    
       .Cols = conColumnsIn_OrdFabFin
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_OrdFabFin_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_OrdFabFin_Selecionado) = ""
       .ColDataType(conCOL_OrdFabFin_Selecionado) = flexDTBoolean
       
       .Cell(flexcpData, 0, conCOL_OrdFabFin_Codigo) = ""
       .ColDataType(conCOL_OrdFabFin_Codigo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdFabFin_Pedido) = ""
       .ColDataType(conCOL_OrdFabFin_Pedido) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdFabFin_DataOrdem) = ""
       .ColDataType(conCOL_OrdFabFin_DataOrdem) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_OrdFabFin_Rotulo) = ""
       .ColDataType(conCOL_OrdFabFin_Rotulo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_OrdFabFin_DescRotulo) = ""
       .ColDataType(conCOL_OrdFabFin_DescRotulo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_OrdFabFin_CodCliente) = ""
       .ColDataType(conCOL_OrdFabFin_CodCliente) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdFabFin_Cliente) = ""
       .ColDataType(conCOL_OrdFabFin_Cliente) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_OrdFabFin_IDProduto) = ""
       .ColDataType(conCOL_OrdFabFin_IDProduto) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdFabFin_Tipo) = ""
       .ColDataType(conCOL_OrdFabFin_Tipo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdFabFin_DescTipo) = ""
       .ColDataType(conCOL_OrdFabFin_DescTipo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_OrdFabFin_TipoDaOP) = ""
       .ColDataType(conCOL_OrdFabFin_TipoDaOP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdFabFin_JaImpressa) = ""
       .ColDataType(conCOL_OrdFabFin_JaImpressa) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdFabFin_FilialEmp) = ""
       .ColDataType(conCOL_OrdFabFin_FilialEmp) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdFabFin_NomeFilialEmp) = ""
       .ColDataType(conCOL_OrdFabFin_NomeFilialEmp) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_OrdFabFin_FoiAlterado) = ""
       .ColDataType(conCOL_OrdFabFin_FoiAlterado) = flexDTLong
       
       .ColWidth(conCOL_OrdFabFin_Selecionado) = 300
       .ColWidth(conCOL_OrdFabFin_Codigo) = 1000
       .ColWidth(conCOL_OrdFabFin_Pedido) = 1000
       .ColWidth(conCOL_OrdFabFin_DataOrdem) = 1000
       .ColWidth(conCOL_OrdFabFin_Rotulo) = 1200
       .ColWidth(conCOL_OrdFabFin_DescRotulo) = 4500
       .ColWidth(conCOL_OrdFabFin_CodCliente) = 1000
       .ColWidth(conCOL_OrdFabFin_Cliente) = 3500
       .ColWidth(conCOL_OrdFabFin_IDProduto) = 0
       .ColWidth(conCOL_OrdFabFin_Tipo) = 0
       .ColWidth(conCOL_OrdFabFin_DescTipo) = 1400
       .ColWidth(conCOL_OrdFabFin_TipoDaOP) = 0
       .ColWidth(conCOL_OrdFabFin_JaImpressa) = 700
       .ColWidth(conCOL_OrdFabFin_FilialEmp) = 0
       .ColWidth(conCOL_OrdFabFin_NomeFilialEmp) = 0
       .ColWidth(conCOL_OrdFabFin_FoiAlterado) = 0
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub Ordem()

    If ConfereData = False Then Exit Sub

    Dim strCampos           As String
    Dim strTIPOOP           As String
    Dim boolTemOPParaImp    As Boolean
    Dim boolPermOPSel       As Boolean
   
    If stOrdem.Tab = 0 Then Call ConfGridOrdFab
    If stOrdem.Tab = 1 Then Call ConfGridOrdFabParc
    If stOrdem.Tab = 2 Then Call ConfGridOrdFabFin
    
    boolTemOPParaImp = False
    boolPermOPSel = PermiteOPQualidade
        
    If BREC.State = 1 Then BREC.Close
    
    sSql = ""
  
    sSql = "Select " & vbCrLf
    
    sSql = sSql & "       ORD.SGI_CODIGO " & vbCrLf
    sSql = sSql & "      ,ORD.SGI_CODPED " & vbCrLf
    sSql = sSql & "      ,ORD.SGI_DATAORDEM " & vbCrLf
    sSql = sSql & "      ,ORD.SGI_CODPROD " & vbCrLf
    sSql = sSql & "      ,ORD.SGI_TIPOP " & vbCrLf
    sSql = sSql & "      ,PED.SGI_CODCLI " & vbCrLf
    sSql = sSql & "      ,CLI.SGI_RAZAOSOC " & vbCrLf
    sSql = sSql & "      ,ORD.SGI_IDPRODUTO " & vbCrLf
    sSql = sSql & "      ,PROD.SGI_CODTIPO " & vbCrLf
    sSql = sSql & "      ,PROD.SGI_DESCRICAO" & vbCrLf
    sSql = sSql & "      ,ORD.SGI_JAIMPRESSA " & vbCrLf
    sSql = sSql & "      ,ORD.SGI_FILIALPED" & vbCrLf
    sSql = sSql & "      ,ORD.SGI_ALTERADO" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & strNOMTABELA1 & " ORD " & vbCrLf
    sSql = sSql & "      ," & strNOMTABELA2 & " PED " & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE  CLI " & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO  PROD " & vbCrLf
    
    '' Permite manipular as OP's
    If boolPermOPSel = True Then
        sSql = sSql & "      ,SGI_CADOPQUALI_IT OPMI" & vbCrLf
    End If
    '' ===============================
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       ORD.SGI_FILIAL = " & FILIAL & vbCrLf
    
    If Len(Trim(txtCODOP.Text)) > 0 Then
        sSql = sSql & "   And ORD.SGI_CODIGO = " & txtCODOP.Text & vbCrLf
    End If
    If Len(Trim(txtCODPED.Text)) > 0 Then
        sSql = sSql & "   And ORD.SGI_CODPED = " & txtCODPED.Text & vbCrLf
    End If
    If Len(Trim(Replace(Replace(mskDTINI.Text, "/", ""), "_", ""))) > 0 And Len(Trim(Replace(Replace(mskDTFIN.Text, "/", ""), "_", ""))) > 0 Then
        sSql = sSql & "   And ORD.SGI_DATAORDEM Between '" & Format(CDate(mskDTINI.Text), "MM/DD/YYYY") & "' And '" & Format(CDate(mskDTFIN.Text), "MM/DD/YYYY") & "'" & vbCrLf
    End If
    If Len(Trim(txtCODROTULO.Text)) > 0 Then
        sSql = sSql & "   And ORD.SGI_CODPROD Like '%" & txtCODROTULO.Text & "%'" & vbCrLf
    End If
    
    If stOrdem.Tab = 0 Then sSql = sSql & "   And ORD.SGI_STATUS = 0" & vbCrLf
    If stOrdem.Tab = 1 Then sSql = sSql & "   And ORD.SGI_STATUS = 1" & vbCrLf
    If stOrdem.Tab = 2 Then sSql = sSql & "   And ORD.SGI_STATUS IN(2,4,6,9)" & vbCrLf
    
    '' ==================================
    '' Permite manipular as OP's
    If boolPermOPSel = True Then
        sSql = sSql & "  And OPMI.SGI_FILIAL    = ORD.SGI_FILIAL" & vbCrLf
        sSql = sSql & "  And OPMI.SGI_CODOP     = ORD.SGI_CODIGO" & vbCrLf
        sSql = sSql & "  And OPMI.SGI_FILIALOP  = " & intFILIALPED & vbCrLf
    End If
    '' ==================================
    
    sSql = sSql & "   And PED.SGI_FILIAL     = ORD.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PED.SGI_CODIGO     = ORD.SGI_CODPED " & vbCrLf
    
    sSql = sSql & "   And CLI.SGI_FILIAL     = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLI.SGI_CODIGO     = PED.SGI_CODCLI " & vbCrLf
    
    If Len(Trim(txtRAZAOSOC.Text)) > 0 Then
        sSql = sSql & "   And CLI.SGI_RAZAOSOC    Like '%" & txtRAZAOSOC.Text & "%'" & vbCrLf
    End If
    
    
    sSql = sSql & "   And PROD.SGI_FILIAL    = ORD.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PROD.SGI_IDPRODUTO = ORD.SGI_IDPRODUTO " & vbCrLf
    
    If Len(Trim(txtCODCLIE.Text)) > 0 Then
        sSql = sSql & "   And PROD.SGI_CODCLIE = " & txtCODCLIE.Text & vbCrLf
    End If
    If Len(Trim(txtRotulo.Text)) > 0 Then
        sSql = sSql & "   And PROD.SGI_DESCRICAO Like '%" & Trim(txtRotulo.Text) & "%'" & vbCrLf
    End If
    
    If intFILIALPED = 0 Then '' Novalata
        If optTPOP(0).Value = True Then
            sSql = sSql & "   And PROD.SGI_CODTIPO = 1" & vbCrLf
        ElseIf optTPOP(1).Value = True Then
            sSql = sSql & "   And PROD.SGI_CODTIPO = 2" & vbCrLf
        End If
    End If
    
    ''If cboFiltro.ListIndex = 0 Then sSql = sSql & "Order by ORD.SGI_CODIGO "
    ''If cboFiltro.ListIndex = 1 Then sSql = sSql & "Order by ORD.SGI_CODPED "
    ''If cboFiltro.ListIndex = 2 Then sSql = sSql & "Order by ORD.SGI_DATAORDEM "
    ''If cboFiltro.ListIndex = 3 Then sSql = sSql & "Order by ORD.SGI_CODPROD "
    ''If cboFiltro.ListIndex = 4 Then sSql = sSql & "Order by PED.SGI_CODCLI "
    ''If cboFiltro.ListIndex = 5 Then sSql = sSql & "Order by CLI.SGI_RAZAOSOC "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    strCampos = ""
    Do While Not BREC.EOF()
    
        strTIPOOP = "NORMAL"
        If BREC!SGI_CODTIPO = 2 Then strTIPOOP = "HOMOLOGADA"
    
        strCampos = 0 & vbTab & _
                    BREC!SGI_CODIGO & vbTab & _
                    BREC!SGI_CODPED & vbTab & _
                    Format(BREC!SGI_DATAORDEM, "DD/MM/YYYY") & vbTab & _
                    Trim(BREC!SGI_CODPROD) & vbTab & _
                    Trim(BREC!SGI_DESCRICAO) & vbTab & _
                    BREC!SGI_CODCLI & vbTab & _
                    Trim(BREC!SGI_RAZAOSOC) & vbTab & _
                    BREC!SGI_IDPRODUTO & vbTab & _
                    BREC!SGI_CODTIPO & vbTab & _
                    strTIPOOP & vbTab & _
                    BREC!SGI_TIPOP & vbTab & _
                    IIf(BREC!SGI_JAIMPRESSA = 1, "SIM", "NÃO") & vbTab & _
                    BREC!SGI_FILIALPED & vbTab & _
                    IIf(BREC!SGI_FILIALPED = 0, "NOVALATA", "STEEL ROW") & vbTab & _
                    BREC!SGI_ALTERADO
    
        If BREC!SGI_JAIMPRESSA = 0 Then boolTemOPParaImp = True
        
        If stOrdem.Tab = 0 Then
            grdOrdeFab.AddItem strCampos
            Call MudaCorCelula(grdOrdeFab, BREC!SGI_JAIMPRESSA, (grdOrdeFab.Rows - 1))
            Call MudaCorCelulaAlterada(grdOrdeFab, BREC!SGI_ALTERADO, grdOrdeFab.Rows - 1)
        ElseIf stOrdem.Tab = 1 Then
            grdPARCIAL.AddItem strCampos
            Call MudaCorCelula(grdPARCIAL, BREC!SGI_JAIMPRESSA, (grdPARCIAL.Rows - 1))
            Call MudaCorCelulaAlterada(grdPARCIAL, BREC!SGI_ALTERADO, grdPARCIAL.Rows - 1)
        ElseIf stOrdem.Tab = 2 Then
            grdFINALIZADO.AddItem strCampos
            Call MudaCorCelula(grdFINALIZADO, BREC!SGI_JAIMPRESSA, (grdFINALIZADO.Rows - 1))
            Call MudaCorCelulaAlterada(grdFINALIZADO, BREC!SGI_ALTERADO, grdFINALIZADO.Rows - 1)
        End If
       
       BREC.MoveNext
    Loop
    BREC.Close

    If boolTemOPParaImp = True Then Call SelTodasSemImp(1)

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call DestroiObjeto
End Sub


Private Sub grdFINALIZADO_Click()
   If (grdFINALIZADO.Rows - 1) > 0 And grdFINALIZADO.Row > 0 Then objCADOORDFAB.CODORDEM = CLng(grdFINALIZADO.Cell(flexcpText, grdFINALIZADO.Row, conCOL_OrdFabFin_Codigo))
End Sub

Private Sub grdFINALIZADO_DblClick()
   If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
   If (grdFINALIZADO.Rows - 1) > 0 Then Operacao "C"
End Sub

Private Sub grdFINALIZADO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
        If (grdFINALIZADO.Rows - 1) > 0 Then Operacao "C"
    End If
End Sub

Private Sub grdFINALIZADO_RowColChange()
   If (grdFINALIZADO.Rows - 1) > 0 And grdFINALIZADO.Row > 0 Then objCADOORDFAB.CODORDEM = CLng(grdFINALIZADO.Cell(flexcpText, grdFINALIZADO.Row, conCOL_OrdFabFin_Codigo))
End Sub

Private Sub grdOrdeFab_Click()
   If (grdOrdeFab.Rows - 1) > 0 And grdOrdeFab.Row > 0 Then objCADOORDFAB.CODORDEM = CLng(grdOrdeFab.Cell(flexcpText, grdOrdeFab.Row, conCOL_OrdFab_Codigo))
End Sub

Private Sub grdOrdeFab_DblClick()
   If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
   If (grdOrdeFab.Rows - 1) > 0 Then Operacao "C"
End Sub

Private Sub grdOrdeFab_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
        If (grdOrdeFab.Rows - 1) > 0 Then Operacao "C"
    End If
End Sub

Private Sub grdOrdeFab_RowColChange()
   If (grdOrdeFab.Rows - 1) > 0 And grdOrdeFab.Row > 0 Then objCADOORDFAB.CODORDEM = CLng(grdOrdeFab.Cell(flexcpText, grdOrdeFab.Row, conCOL_OrdFab_Codigo))
End Sub


Private Sub grdPARCIAL_Click()
    With grdPARCIAL
        If (.Rows - 1) > 0 And .Row > 0 Then objCADOORDFAB.CODORDEM = CLng(.Cell(flexcpText, .Row, conCOL_OrdFabParc_Codigo))
    End With
End Sub

Private Sub grdPARCIAL_DblClick()
   If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
   If (grdPARCIAL.Rows - 1) > 0 Then Operacao "C"
End Sub

Private Sub grdPARCIAL_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
        If (grdPARCIAL.Rows - 1) > 0 Then Operacao "C"
    End If
End Sub

Private Sub grdPARCIAL_RowColChange()
    With grdPARCIAL
        If (.Rows - 1) > 0 And .Row > 0 Then objCADOORDFAB.CODORDEM = CLng(.Cell(flexcpText, .Row, conCOL_OrdFabParc_Codigo))
    End With
End Sub





Private Sub mskDTFIN_GotFocus()
    objFuncoes.SelecionaCampos mskDTFIN.Name, Me
End Sub

Private Sub mskDTINI_GotFocus()
    objFuncoes.SelecionaCampos mskDTINI.Name, Me
End Sub

Private Sub optSelecTodosSN_Click(index As Integer)
    Call Seleciona(index)
End Sub

Private Sub optSelOPNaoImpSN_Click(index As Integer)
    Call SelTodasSemImp(index)
End Sub

Private Sub stOrdem_Click(PreviousTab As Integer)
   If stOrdem.Tab = 0 Then Call ConfGridOrdFab
   If stOrdem.Tab = 1 Then Call ConfGridOrdFabParc
   If stOrdem.Tab = 2 Then Call ConfGridOrdFabFin
End Sub
Private Sub Timer1_Timer()
    Call AbilitaCampos
End Sub


Private Sub Imprime(lngIndice As Integer, strCODORDEM As String)

On Error GoTo Err_Imp
    
    Dim strNomArquivo As String
    Dim grdGENERICA   As Variant
    Dim lngCOLGen     As Long
    Dim lngCOLGen2    As Long
    Dim lngCOLGen3    As Long
    
    ''=====================================
    strNomArquivo = ""
    If stOrdem.Tab = 0 Then
       Set grdGENERICA = grdOrdeFab
       lngCOLGen = conCOL_OrdFab_Tipo
       lngCOLGen2 = conCOL_OrdFab_TipoDaOP
       lngCOLGen3 = conCOL_OrdFab_FilialEmp
    ElseIf stOrdem.Tab = 1 Then
       Set grdGENERICA = grdPARCIAL
       lngCOLGen = conCOL_OrdFabParc_Tipo
       lngCOLGen2 = conCOL_OrdFabParc_TipoDaOP
       lngCOLGen3 = conCOL_OrdFabParc_FilialEmp
    ElseIf stOrdem.Tab = 2 Then
       Set grdGENERICA = grdFINALIZADO
       lngCOLGen = conCOL_OrdFabFin_Tipo
       lngCOLGen2 = conCOL_OrdFabFin_TipoDaOP
       lngCOLGen3 = conCOL_OrdFabFin_FilialEmp
    End If
    
    If (grdGENERICA.Rows - 1) = 0 Then
        MsgBox "Selecione uma OP !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    If (grdGENERICA.Row) <= 0 Then
        MsgBox "Selecione uma OP !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    
    sSql = ""
    
    If intFILIALPED = 0 Then  '' Novalata
        If grdGENERICA.Cell(flexcpText, grdGENERICA.Row, lngCOLGen) = 2 Then
        
            sSql = "Select " & vbCrLf
            sSql = sSql & "        " & strNOMTABELA1 & ".SGI_CODIGO " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_DATAORDEM " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_CODPED " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_CODPROD " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_QTDE " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_SALDO " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_ALTFILM " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_FOTNOVO " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_REPETICAO " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_DATENTREGA " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_CODOPMAE " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_NOMEVEND " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_OBSOP " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_FECHTPFU " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_STATUS " & vbCrLf
            
            sSql = sSql & "       ," & strNOMTABELA2 & ".SGI_CODCLI " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA2 & ".SGI_CODVEND " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA2 & ".SGI_PARAESTOQUE " & vbCrLf
            sSql = sSql & "       ,SGI_CADCLIENTE.SGI_RAZAOSOC " & vbCrLf
            
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_IDPRODUTO " & vbCrLf
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_DESCRICAO " & vbCrLf
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_FechTampaFuro " & vbCrLf
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_VernCorpo " & vbCrLf
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_VernTampa " & vbCrLf
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_VernFundo " & vbCrLf
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_CorpoEspess " & vbCrLf
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_TampaEspess " & vbCrLf
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_FundoEspess " & vbCrLf
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_ArgolaEspess " & vbCrLf
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_Pipeta " & vbCrLf
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_AZELHA " & vbCrLf
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_QTDEPORFOLHA" & vbCrLf
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_CODTIPO" & vbCrLf
            
            sSql = sSql & "  From " & vbCrLf
            
            sSql = sSql & "       SGI_CADCLIENTE SGI_CADCLIENTE " & vbCrLf
            sSql = sSql & "      ," & strNOMTABELA2 & " " & strNOMTABELA2 & vbCrLf
            sSql = sSql & "      ,SGI_CADPRODUTO SGI_CADPRODUTO " & vbCrLf
            sSql = sSql & "      ," & strNOMTABELA1 & " " & strNOMTABELA1 & vbCrLf
            
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       " & strNOMTABELA1 & ".SGI_FILIAL     = " & FILIAL & vbCrLf
            sSql = sSql & "   And " & strNOMTABELA1 & ".SGI_CODIGO     In (" & Trim(strCODORDEM) & ")" & vbCrLf
            
            sSql = sSql & "   And " & strNOMTABELA1 & ".SGI_FILIAL     = " & strNOMTABELA2 & ".SGI_FILIAL " & vbCrLf
            sSql = sSql & "   And " & strNOMTABELA1 & ".SGI_CODPED     = " & strNOMTABELA2 & ".SGI_CODIGO " & vbCrLf
            
            sSql = sSql & "   And " & strNOMTABELA1 & ".SGI_FILIAL     = SGI_CADPRODUTO.SGI_FILIAL " & vbCrLf
            sSql = sSql & "   And " & strNOMTABELA1 & ".SGI_IDPRODUTO  = SGI_CADPRODUTO.SGI_IDPRODUTO " & vbCrLf
            
            sSql = sSql & "   And " & strNOMTABELA2 & ".SGI_FILIAL   = SGI_CADCLIENTE.SGI_FILIAL " & vbCrLf
            sSql = sSql & "   And " & strNOMTABELA2 & ".SGI_CODCLI   = SGI_CADCLIENTE.SGI_CODIGO " & vbCrLf
            
        ElseIf grdGENERICA.Cell(flexcpText, grdGENERICA.Row, lngCOLGen) = 1 Then
        
            sSql = "Select " & vbCrLf
            sSql = sSql & "        " & strNOMTABELA1 & ".SGI_CODIGO " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_DATAORDEM " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_DATAPED " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_CODPED " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_CODPROD " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_QTDE " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_SALDO " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_ALTFILM " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_FOTNOVO " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_REPETICAO " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_DATENTREGA " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_CODOPMAE " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_NOMEVEND " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_OBSOP " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_FECHTPFU " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_STATUS " & vbCrLf
            
            sSql = sSql & "       ," & strNOMTABELA2 & ".SGI_CODCLI " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA2 & ".SGI_CODVEND " & vbCrLf
            sSql = sSql & "       ," & strNOMTABELA2 & ".SGI_PARAESTOQUE " & vbCrLf
            
            sSql = sSql & "       ,SGI_CADCLIENTE.SGI_RAZAOSOC " & vbCrLf
            
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_IDPRODUTO " & vbCrLf
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_DESCRICAO " & vbCrLf
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_FechTampaFuro " & vbCrLf
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_VernCorpo " & vbCrLf
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_VernTampa " & vbCrLf
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_VernFundo " & vbCrLf
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_CorpoEspess " & vbCrLf
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_TampaEspess " & vbCrLf
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_FundoEspess " & vbCrLf
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_ArgolaEspess " & vbCrLf
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_Pipeta " & vbCrLf
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_AZELHA " & vbCrLf
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_VernArgola " & vbCrLf
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_FechTampaFuro" & vbCrLf
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_NECKIN" & vbCrLf
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_DIMPADRAO" & vbCrLf
            
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_DESENV" & vbCrLf
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_ALTURA" & vbCrLf
            
            sSql = sSql & "       ,SGI_CADLINHAPRODUTO.SGI_DESENV" & vbCrLf
            sSql = sSql & "       ,SGI_CADLINHAPRODUTO.SGI_ALTURA" & vbCrLf
            sSql = sSql & "       ,SGI_CADPRODUTO.SGI_CODTIPO" & vbCrLf
            
            sSql = sSql & "  From " & vbCrLf
            
            sSql = sSql & "       SGI_CADCLIENTE SGI_CADCLIENTE " & vbCrLf
            sSql = sSql & "      ,SGI_CADLINHAPRODUTO SGI_CADLINHAPRODUTO " & vbCrLf
            sSql = sSql & "      ," & strNOMTABELA2 & " " & strNOMTABELA2 & vbCrLf
            sSql = sSql & "      ,SGI_CADPRODUTO SGI_CADPRODUTO " & vbCrLf
            sSql = sSql & "      ," & strNOMTABELA1 & " " & strNOMTABELA1 & vbCrLf
            
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       " & strNOMTABELA1 & ".SGI_FILIAL     = " & FILIAL & vbCrLf
            sSql = sSql & "   And " & strNOMTABELA1 & ".SGI_CODIGO     In (" & Trim(strCODORDEM) & ")" & vbCrLf
            
            sSql = sSql & "   And " & strNOMTABELA1 & ".SGI_FILIAL     = " & strNOMTABELA2 & ".SGI_FILIAL " & vbCrLf
            sSql = sSql & "   And " & strNOMTABELA1 & ".SGI_CODPED     = " & strNOMTABELA2 & ".SGI_CODIGO " & vbCrLf
            
            sSql = sSql & "   And " & strNOMTABELA1 & ".SGI_FILIAL     = SGI_CADPRODUTO.SGI_FILIAL " & vbCrLf
            sSql = sSql & "   And " & strNOMTABELA1 & ".SGI_IDPRODUTO  = SGI_CADPRODUTO.SGI_IDPRODUTO " & vbCrLf
            
            sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIAL     = SGI_CADLINHAPRODUTO.SGI_FILIAL " & vbCrLf
            sSql = sSql & "   And SGI_CADPRODUTO.SGI_CODLINPROD = SGI_CADLINHAPRODUTO.SGI_CODLIN " & vbCrLf
            
            sSql = sSql & "   And " & strNOMTABELA2 & ".SGI_FILIAL   = SGI_CADCLIENTE.SGI_FILIAL " & vbCrLf
            sSql = sSql & "   And " & strNOMTABELA2 & ".SGI_CODCLI   = SGI_CADCLIENTE.SGI_CODIGO " & vbCrLf
        
        End If
    ElseIf intFILIALPED = 1 Then
    
        sSql = ""
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "        " & strNOMTABELA1 & ".SGI_CODIGO " & vbCrLf
        sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_DATAORDEM " & vbCrLf
        sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_DATAPED " & vbCrLf
        sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_CODPED " & vbCrLf
        sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_CODPROD " & vbCrLf
        sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_QTDE " & vbCrLf
        sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_SALDO " & vbCrLf
        sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_ALTFILM " & vbCrLf
        sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_FOTNOVO " & vbCrLf
        sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_REPETICAO " & vbCrLf
        sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_DATENTREGA " & vbCrLf
        sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_CODOPMAE " & vbCrLf
        sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_NOMEVEND " & vbCrLf
        sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_OBSOP " & vbCrLf
        sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_FECHTPFU " & vbCrLf
        sSql = sSql & "       ," & strNOMTABELA1 & ".SGI_STATUS " & vbCrLf
        
        sSql = sSql & "       ," & strNOMTABELA2 & ".SGI_CODCLI " & vbCrLf
        sSql = sSql & "       ," & strNOMTABELA2 & ".SGI_CODVEND " & vbCrLf
        sSql = sSql & "       ," & strNOMTABELA2 & ".SGI_PARAESTOQUE " & vbCrLf
        
        sSql = sSql & "       ,SGI_CADCLIENTE.SGI_RAZAOSOC " & vbCrLf
        
        sSql = sSql & "       ,SGI_CADPRODUTO.SGI_IDPRODUTO " & vbCrLf
        sSql = sSql & "       ,SGI_CADPRODUTO.SGI_DESCRICAO " & vbCrLf
        sSql = sSql & "       ,SGI_CADPRODUTO.SGI_FechTampaFuro " & vbCrLf
        sSql = sSql & "       ,SGI_CADPRODUTO.SGI_VernCorpo " & vbCrLf
        sSql = sSql & "       ,SGI_CADPRODUTO.SGI_VernTampa " & vbCrLf
        sSql = sSql & "       ,SGI_CADPRODUTO.SGI_VernFundo " & vbCrLf
        sSql = sSql & "       ,SGI_CADPRODUTO.SGI_CorpoEspess " & vbCrLf
        sSql = sSql & "       ,SGI_CADPRODUTO.SGI_TampaEspess " & vbCrLf
        sSql = sSql & "       ,SGI_CADPRODUTO.SGI_FundoEspess " & vbCrLf
        sSql = sSql & "       ,SGI_CADPRODUTO.SGI_ArgolaEspess " & vbCrLf
        sSql = sSql & "       ,SGI_CADPRODUTO.SGI_Pipeta " & vbCrLf
        sSql = sSql & "       ,SGI_CADPRODUTO.SGI_AZELHA " & vbCrLf
        sSql = sSql & "       ,SGI_CADPRODUTO.SGI_VernArgola " & vbCrLf
        sSql = sSql & "       ,SGI_CADPRODUTO.SGI_FechTampaFuro" & vbCrLf
        sSql = sSql & "       ,SGI_CADPRODUTO.SGI_NECKIN" & vbCrLf
        sSql = sSql & "       ,SGI_CADPRODUTO.SGI_DIMPADRAO" & vbCrLf
        
        sSql = sSql & "       ,SGI_CADPRODUTO.SGI_DESENV" & vbCrLf
        sSql = sSql & "       ,SGI_CADPRODUTO.SGI_ALTURA" & vbCrLf
        sSql = sSql & "       ,SGI_CADPRODUTO.SGI_CODTIPO" & vbCrLf
        
        sSql = sSql & "       ,SGI_CADLINHAPRODUTO.SGI_DESENV" & vbCrLf
        sSql = sSql & "       ,SGI_CADLINHAPRODUTO.SGI_ALTURA" & vbCrLf
        
        sSql = sSql & "  From " & vbCrLf
        
        sSql = sSql & "       SGI_CADCLIENTE SGI_CADCLIENTE " & vbCrLf
        sSql = sSql & "      ,SGI_CADLINHAPRODUTO SGI_CADLINHAPRODUTO " & vbCrLf
        sSql = sSql & "      ," & strNOMTABELA2 & " " & strNOMTABELA2 & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO SGI_CADPRODUTO " & vbCrLf
        sSql = sSql & "      ," & strNOMTABELA1 & " " & strNOMTABELA1 & vbCrLf
        
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       " & strNOMTABELA1 & ".SGI_FILIAL     = " & FILIAL & vbCrLf
        sSql = sSql & "   And " & strNOMTABELA1 & ".SGI_CODIGO     In(" & Trim(strCODORDEM) & ")" & vbCrLf
        
        sSql = sSql & "   And " & strNOMTABELA1 & ".SGI_FILIAL     = " & strNOMTABELA2 & ".SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And " & strNOMTABELA1 & ".SGI_CODPED     = " & strNOMTABELA2 & ".SGI_CODIGO " & vbCrLf
        
        sSql = sSql & "   And " & strNOMTABELA1 & ".SGI_FILIAL     = SGI_CADPRODUTO.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And " & strNOMTABELA1 & ".SGI_IDPRODUTO  = SGI_CADPRODUTO.SGI_IDPRODUTO " & vbCrLf
        
        sSql = sSql & "   And SGI_CADPRODUTO.SGI_FILIAL            = SGI_CADLINHAPRODUTO.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And SGI_CADPRODUTO.SGI_CODLINPROD        = SGI_CADLINHAPRODUTO.SGI_CODLIN " & vbCrLf
        
        sSql = sSql & "   And " & strNOMTABELA2 & ".SGI_FILIAL     = SGI_CADCLIENTE.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And " & strNOMTABELA2 & ".SGI_CODCLI     = SGI_CADCLIENTE.SGI_CODIGO " & vbCrLf
    End If
    
    BREC8.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC8.EOF() Then
        MsgBox "Não há dados para realizar a impressão !!!", vbOKOnly + vbExclamation, "Aviso"
        BREC8.Close
        Exit Sub
    End If
    BREC8.Close
    
    With grdGENERICA
        If .Row = 0 Then
           MsgBox "Selecione uma ordem !!!", vbOKOnly + vbExclamation, "Aviso"
           Exit Sub
        End If
        
        If intFILIALPED = 0 Then
            If optTPOP(0).Value = True Then strNomArquivo = "ORDPRODNORM_NOVALATA.RPT"
            If optTPOP(1).Value = True Then strNomArquivo = "ORDPRODHOMO_NOVALATA.RPT"
        ElseIf intFILIALPED = 1 Then
            strNomArquivo = "ORDPRODNORMSTEEL.RPT"
        End If
    End With
        
        
        
''        If .Cell(flexcpText, .Row, lngCOLGen) = 2 Then
''           If .Cell(flexcpText, .Row, lngCOLGen2) = 0 Then
''              If intFILIALPED = 0 Then strNomArquivo = "ORDPRODHOMO2.RPT"
''           ElseIf .Cell(flexcpText, .Row, lngCOLGen2) = 1 Then
''              If intFILIALPED = 0 Then strNomArquivo = "ORDPRODHOMO2.RPT"
''           End If
''        ElseIf .Cell(flexcpText, .Row, lngCOLGen) = 1 Then
''           If .Cell(flexcpText, .Row, lngCOLGen2) = 0 Then
''              If intFILIALPED = 0 Then strNomArquivo = "ORDPRODHOMO2.RPT"
''              ''If intFILIALPED = 0 Then strNomArquivo = "ORDPRODNORMAL.RPT"
''              If intFILIALPED = 1 Then strNomArquivo = "ORDPRODNORMAL2STEEL.RPT"
''           ElseIf .Cell(flexcpText, .Row, lngCOLGen2) = 1 Then
''              If intFILIALPED = 0 Then strNomArquivo = "ORDPRODHOMO2.RPT"
''              ''If intFILIALPED = 0 Then strNomArquivo = "ORDPRODNORMAL2.RPT"
''              If intFILIALPED = 1 Then strNomArquivo = "ORDPRODNORMAL2STEEL.RPT"
''           End If
''        Else
''           If .Cell(flexcpText, .Row, lngCOLGen2) = 0 Then
''              If intFILIALPED = 0 Then strNomArquivo = "ORDPRODHOMO2.RPT"
''              ''If intFILIALPED = 0 Then strNomArquivo = "ORDPRODNORMAL.RPT"
''              ''If intFILIALPED = 1 Then strNomArquivo = "ORDPRODNORMALSTEEL.RPT"
''              If intFILIALPED = 1 Then strNomArquivo = "ORDPRODNORMAL2STEEL.RPT"
''           ElseIf .Cell(flexcpText, .Row, lngCOLGen2) = 1 Then
''              If intFILIALPED = 0 Then strNomArquivo = "ORDPRODHOMO2.RPT"
''              ''If intFILIALPED = 0 Then strNomArquivo = "ORDPRODNORMAL2.RPT"
''              If intFILIALPED = 1 Then strNomArquivo = "ORDPRODNORMAL2STEEL.RPT"
''           End If
''        End If
''=====================================
    
    If Len(Trim(strNomArquivo)) > 0 Then
        Call objRel.REL(FILIAL, sSql, strCamRelNovo & cCamRelPCP & strNomArquivo, Linha, 1, "", "", False, strAcesso, True)
        objCADOORDFAB.INCODORDS = strCODORDEM
        Call objCADOORDFAB.GRAVA("IMP", strNOMTABELA1, "")
        Call objCADOORDFAB.GRAVA("ALT", strNOMTABELA1, "")
    End If
    
    Exit Sub
    
Err_Imp:

    MsgBox "Erro : " & Err.Number & " - " & Err.Description, vbOKOnly + vbCritical, "Erro"

End Sub

Private Sub MudaCorCelula(grdGENERICA As Variant, lngID As Long, lngROW As Long)
    With grdGENERICA
        If lngID = 1 Then
            .Cell(flexcpBackColor, lngROW, conCOL_OrdFab_Codigo, lngROW, conCOL_OrdFab_NomeFilialEmp) = &H80FF80    '' Verde
        ElseIf lngID = 0 Then
            .Cell(flexcpBackColor, lngROW, conCOL_OrdFab_Codigo, lngROW, conCOL_OrdFab_NomeFilialEmp) = &HC0C0FF    '' Vermelho
        End If
    End With
End Sub

Private Sub Atualiza_Grid()
    
     Dim I              As Long
     Dim bolAchou       As Boolean
     Dim lngCODIGO      As Long
     Dim lngCOL         As Long
     Dim strACAO        As String
     Dim grdGENERICA    As VSFlexGrid
     
     If stOrdem.Tab = 0 Then
        Set grdGENERICA = grdOrdeFab
        lngCOL = conCOL_OrdFab_Codigo
     ElseIf stOrdem.Tab = 1 Then
        Set grdGENERICA = grdPARCIAL
        lngCOL = conCOL_OrdFabParc_Codigo
     ElseIf stOrdem.Tab = 2 Then
        Set grdGENERICA = grdFINALIZADO
        lngCOL = conCOL_OrdFabFin_Codigo
     End If
     
     bolAchou = False
     
     If BRECATU.State = 1 Then BRECATU.Close
     ''If BREC2.State = 1 Then BREC2.Close
     
     With grdGENERICA
         
         sSql = "Select" & vbCrLf
         sSql = sSql & "      * " & vbCrLf
         sSql = sSql & "  From" & vbCrLf
         sSql = sSql & "       SGI_ATUALIZA" & vbCrLf
         sSql = sSql & " Where" & vbCrLf
         sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
         sSql = sSql & "   And SGI_MODULO = 'frmCADOORDFAB'" & vbCrLf
    
         BRECATU.Open sSql, adoBanco_Dados, adOpenDynamic
         If Not BRECATU.EOF Then
            strACAO = Trim(BRECATU!SGI_ACAO)
            lngCODIGO = BRECATU!SGI_CODIGO
         End If
         BRECATU.Close
            
        I = .FindRow(lngCODIGO, , lngCOL)
        If I > 0 Then
           If Trim(strACAO) = "E" Or Trim(strACAO) = "B" Then
              If .Rows = 2 Then .Rows = 1
              If .Rows > 2 Then .RemoveItem I
           ElseIf Trim(lngCODIGO) = "I" Or Trim(strACAO) = "A" Then
              bolAchou = True
           End If
        End If
            
        sSql = "Select " & vbCrLf
        
        sSql = sSql & "       ORD.SGI_CODIGO " & vbCrLf
        sSql = sSql & "      ,ORD.SGI_CODPED " & vbCrLf
        sSql = sSql & "      ,ORD.SGI_DATAORDEM " & vbCrLf
        sSql = sSql & "      ,ORD.SGI_CODPROD " & vbCrLf
        sSql = sSql & "      ,ORD.SGI_TIPOP " & vbCrLf
        sSql = sSql & "      ,PED.SGI_CODCLI " & vbCrLf
        sSql = sSql & "      ,CLI.SGI_RAZAOSOC " & vbCrLf
        sSql = sSql & "      ,ORD.SGI_IDPRODUTO " & vbCrLf
        sSql = sSql & "      ,PROD.SGI_CODTIPO " & vbCrLf
        sSql = sSql & "      ,ORD.SGI_JAIMPRESSA " & vbCrLf
        sSql = sSql & "      ,ORD.SGI_FILIALPED " & vbCrLf
        
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_ORDEMPROD   ORD " & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDH PED " & vbCrLf
        sSql = sSql & "      ,SGI_CADCLIENTE  CLI " & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO  PROD " & vbCrLf
        
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       ORD.SGI_FILIAL     = " & FILIAL & vbCrLf
        sSql = sSql & "   And ORD.SGI_CODIGO     = " & lngCODIGO & vbCrLf
                
        If stOrdem.Tab = 0 Then
            sSql = sSql & "   And ORD.SGI_STATUS     = 0" & vbCrLf
        ElseIf stOrdem.Tab = 1 Then
            sSql = sSql & "   And ORD.SGI_STATUS     = 1" & vbCrLf
        ElseIf stOrdem.Tab = 2 Then
            sSql = sSql & "   And ORD.SGI_STATUS     = 2" & vbCrLf
        End If
        
        sSql = sSql & "   And PED.SGI_FILIAL     = ORD.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And PED.SGI_CODIGO     = ORD.SGI_CODPED " & vbCrLf
        sSql = sSql & "   And CLI.SGI_FILIAL     = PED.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And CLI.SGI_CODIGO     = PED.SGI_CODCLI " & vbCrLf
        sSql = sSql & "   And PROD.SGI_FILIAL    = ORD.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And PROD.SGI_IDPRODUTO = ORD.SGI_IDPRODUTO " & vbCrLf
            
        BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC2.EOF Then
            If bolAchou = False And Trim(strACAO) = "I" Then
            
                .AddItem BREC2!SGI_CODIGO & vbTab & _
                         BREC2!SGI_CODPED & vbTab & _
                         Format(BREC2!SGI_DATAORDEM, "DD/MM/YYYY") & vbTab & _
                         Trim(BREC2!SGI_CODPROD) & vbTab & _
                         BREC2!SGI_CODCLI & vbTab & _
                         Trim(BREC2!SGI_RAZAOSOC) & vbTab & _
                         BREC2!SGI_IDPRODUTO & vbTab & _
                         BREC2!SGI_CODTIPO & vbTab & _
                         IIf(BREC2!SGI_CODTIPO = 1, "NORMAL", "HOMOLOGADA") & vbTab & _
                         BREC2!SGI_TIPOP & vbTab & _
                         IIf(BREC2!SGI_JAIMPRESSA = 1, "SIM", "NÃO") & vbTab & _
                         BREC2!SGI_FILIALPED & vbTab & _
                         IIf(BREC2!SGI_FILIALPED = 0, "NOVALATA", "STEEL ROW")
        
                Call MudaCorCelula(grdGENERICA, BREC2!SGI_JAIMPRESSA, (.Rows - 1))
            
            ElseIf bolAchou = True And Trim(strACAO) = "A" Then
            
               .Cell(flexcpText, I, conCOL_OrdFab_Codigo) = BREC2!SGI_CODIGO
               .Cell(flexcpText, I, conCOL_OrdFab_Pedido) = BREC2!SGI_CODPED
               .Cell(flexcpText, I, conCOL_OrdFab_DataOrdem) = Format(BREC2!SGI_DATAORDEM, "DD/MM/YYYY")
               .Cell(flexcpText, I, conCOL_OrdFab_Rotulo) = Trim(BREC2!SGI_CODPROD)
               .Cell(flexcpText, I, conCOL_OrdFab_CodCliente) = BREC2!SGI_CODCLI
               .Cell(flexcpText, I, conCOL_OrdFab_Cliente) = Trim(BREC2!SGI_RAZAOSOC)
               .Cell(flexcpText, I, conCOL_OrdFab_IDProduto) = BREC2!SGI_IDPRODUTO
               .Cell(flexcpText, I, conCOL_OrdFab_Tipo) = BREC2!SGI_CODTIPO
               .Cell(flexcpText, I, conCOL_OrdFab_DescTipo) = IIf(BREC2!SGI_CODTIPO = 1, "NORMAL", "HOMOLOGADA")
               .Cell(flexcpText, I, conCOL_OrdFab_TipoDaOP) = BREC2!SGI_TIPOP
               .Cell(flexcpText, I, conCOL_OrdFab_JaImpressa) = IIf(BREC2!SGI_JAIMPRESSA = 1, "SIM", "NÃO")
               .Cell(flexcpText, I, conCOL_OrdFab_FilialEmp) = BREC2!SGI_FILIALPED
               .Cell(flexcpText, I, conCOL_OrdFab_NomeFilialEmp) = IIf(BREC2!SGI_FILIALPED = 0, "NOVALATA", "STEEL ROW")
               
                Call MudaCorCelula(grdGENERICA, BREC2!SGI_JAIMPRESSA, I)
            
            End If
        
        
        End If
        BREC2.Close

     End With
      
End Sub


Private Sub ConfGridOrdFabParc()

    With grdPARCIAL
    
       .Cols = conColumnsIn_OrdFabParc
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_OrdFabParc_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_OrdFabParc_Selecionado) = ""
       .ColDataType(conCOL_OrdFabParc_Selecionado) = flexDTBoolean
       
       .Cell(flexcpData, 0, conCOL_OrdFabParc_Codigo) = ""
       .ColDataType(conCOL_OrdFabParc_Codigo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdFabParc_Pedido) = ""
       .ColDataType(conCOL_OrdFabParc_Pedido) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdFabParc_DataOrdem) = ""
       .ColDataType(conCOL_OrdFabParc_DataOrdem) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_OrdFabParc_Rotulo) = ""
       .ColDataType(conCOL_OrdFabParc_Rotulo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_OrdFabParc_DescRotulo) = ""
       .ColDataType(conCOL_OrdFabParc_DescRotulo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_OrdFabParc_CodCliente) = ""
       .ColDataType(conCOL_OrdFabParc_CodCliente) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdFabParc_Cliente) = ""
       .ColDataType(conCOL_OrdFabParc_Cliente) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_OrdFabParc_IDProduto) = ""
       .ColDataType(conCOL_OrdFabParc_IDProduto) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdFabParc_Tipo) = ""
       .ColDataType(conCOL_OrdFabParc_Tipo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdFabParc_DescTipo) = ""
       .ColDataType(conCOL_OrdFabParc_DescTipo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_OrdFabParc_TipoDaOP) = ""
       .ColDataType(conCOL_OrdFabParc_TipoDaOP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdFabParc_JaImpressa) = ""
       .ColDataType(conCOL_OrdFabParc_JaImpressa) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdFabParc_FilialEmp) = ""
       .ColDataType(conCOL_OrdFabParc_FilialEmp) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdFabParc_NomeFilialEmp) = ""
       .ColDataType(conCOL_OrdFabParc_NomeFilialEmp) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_OrdFab_FoiAlterado) = ""
       .ColDataType(conCOL_OrdFab_FoiAlterado) = flexDTLong
       
       .ColWidth(conCOL_OrdFab_Selecionado) = 300
       .ColWidth(conCOL_OrdFabParc_Codigo) = 1000
       .ColWidth(conCOL_OrdFabParc_Pedido) = 1000
       .ColWidth(conCOL_OrdFabParc_DataOrdem) = 1000
       .ColWidth(conCOL_OrdFabParc_Rotulo) = 1200
       .ColWidth(conCOL_OrdFabParc_DescRotulo) = 4500
       .ColWidth(conCOL_OrdFabParc_CodCliente) = 1000
       .ColWidth(conCOL_OrdFabParc_Cliente) = 4500
       .ColWidth(conCOL_OrdFabParc_IDProduto) = 0
       .ColWidth(conCOL_OrdFabParc_Tipo) = 0
       .ColWidth(conCOL_OrdFabParc_DescTipo) = 1400
       .ColWidth(conCOL_OrdFabParc_TipoDaOP) = 0
       .ColWidth(conCOL_OrdFabParc_JaImpressa) = 700
       .ColWidth(conCOL_OrdFabParc_FilialEmp) = 0
       .ColWidth(conCOL_OrdFabParc_NomeFilialEmp) = 0
       .ColWidth(conCOL_OrdFab_FoiAlterado) = 0
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub PreencheGridGeradoParcial()

    With grdPARCIAL
    
        sSql = "Select " & vbCrLf
        
        sSql = sSql & "       ORD.SGI_CODIGO " & vbCrLf
        sSql = sSql & "      ,ORD.SGI_CODPED " & vbCrLf
        sSql = sSql & "      ,ORD.SGI_DATAORDEM " & vbCrLf
        sSql = sSql & "      ,ORD.SGI_CODPROD " & vbCrLf
        sSql = sSql & "      ,ORD.SGI_TIPOP " & vbCrLf
        sSql = sSql & "      ,PED.SGI_CODCLI " & vbCrLf
        sSql = sSql & "      ,CLI.SGI_RAZAOSOC " & vbCrLf
        sSql = sSql & "      ,ORD.SGI_IDPRODUTO " & vbCrLf
        sSql = sSql & "      ,PROD.SGI_CODTIPO " & vbCrLf
        sSql = sSql & "      ,ORD.SGI_JAIMPRESSA " & vbCrLf
        
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_ORDEMPROD   ORD " & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDH PED " & vbCrLf
        sSql = sSql & "      ,SGI_CADCLIENTE  CLI " & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO  PROD " & vbCrLf
        
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       ORD.SGI_FILIAL     = " & FILIAL & vbCrLf
        sSql = sSql & "   And ORD.SGI_STATUS     = 1" & vbCrLf
        sSql = sSql & "   And PED.SGI_FILIAL     = ORD.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And PED.SGI_CODIGO     = ORD.SGI_CODPED " & vbCrLf
        sSql = sSql & "   And CLI.SGI_FILIAL     = PED.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And CLI.SGI_CODIGO     = PED.SGI_CODCLI " & vbCrLf
        sSql = sSql & "   And PROD.SGI_FILIAL    = ORD.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And PROD.SGI_IDPRODUTO = ORD.SGI_IDPRODUTO " & vbCrLf
        sSql = sSql & "Order By ORD.SGI_CODIGO "
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        Do While Not BREC.EOF()
        
            .AddItem BREC!SGI_CODIGO & vbTab & _
                     BREC!SGI_CODPED & vbTab & _
                     Format(BREC!SGI_DATAORDEM, "DD/MM/YYYY") & vbTab & _
                     Trim(BREC!SGI_CODPROD) & vbTab & _
                     BREC!SGI_CODCLI & vbTab & _
                     Trim(BREC!SGI_RAZAOSOC) & vbTab & _
                     BREC!SGI_IDPRODUTO & vbTab & _
                     BREC!SGI_CODTIPO & vbTab & _
                     IIf(BREC!SGI_CODTIPO = 1, "NORMAL", "HOMOLOGADA") & vbTab & _
                     BREC!SGI_TIPOP & vbTab & _
                     IIf(BREC!SGI_JAIMPRESSA = 1, "SIM", "NÃO")
        
            Call MudaCorCelula(grdPARCIAL, BREC!SGI_JAIMPRESSA, (.Rows - 1))
            
            BREC.MoveNext
        Loop
        BREC.Close

    End With
End Sub


Private Sub PreencheGridGeradoFinalizado()

    With grdFINALIZADO
    
        sSql = "Select " & vbCrLf
        
        sSql = sSql & "       ORD.SGI_CODIGO " & vbCrLf
        sSql = sSql & "      ,ORD.SGI_CODPED " & vbCrLf
        sSql = sSql & "      ,ORD.SGI_DATAORDEM " & vbCrLf
        sSql = sSql & "      ,ORD.SGI_CODPROD " & vbCrLf
        sSql = sSql & "      ,ORD.SGI_TIPOP " & vbCrLf
        sSql = sSql & "      ,PED.SGI_CODCLI " & vbCrLf
        sSql = sSql & "      ,CLI.SGI_RAZAOSOC " & vbCrLf
        sSql = sSql & "      ,ORD.SGI_IDPRODUTO " & vbCrLf
        sSql = sSql & "      ,PROD.SGI_CODTIPO " & vbCrLf
        sSql = sSql & "      ,ORD.SGI_JAIMPRESSA " & vbCrLf
        
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_ORDEMPROD   ORD " & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDH PED " & vbCrLf
        sSql = sSql & "      ,SGI_CADCLIENTE  CLI " & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO  PROD " & vbCrLf
        
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       ORD.SGI_FILIAL     = " & FILIAL & vbCrLf
        sSql = sSql & "   And ORD.SGI_STATUS     = 2" & vbCrLf
        sSql = sSql & "   And PED.SGI_FILIAL     = ORD.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And PED.SGI_CODIGO     = ORD.SGI_CODPED " & vbCrLf
        sSql = sSql & "   And CLI.SGI_FILIAL     = PED.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And CLI.SGI_CODIGO     = PED.SGI_CODCLI " & vbCrLf
        sSql = sSql & "   And PROD.SGI_FILIAL    = ORD.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And PROD.SGI_IDPRODUTO = ORD.SGI_IDPRODUTO " & vbCrLf
        sSql = sSql & "Order By ORD.SGI_CODIGO "
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        Do While Not BREC.EOF()
        
            .AddItem BREC!SGI_CODIGO & vbTab & _
                     BREC!SGI_CODPED & vbTab & _
                     Format(BREC!SGI_DATAORDEM, "DD/MM/YYYY") & vbTab & _
                     Trim(BREC!SGI_CODPROD) & vbTab & _
                     BREC!SGI_CODCLI & vbTab & _
                     Trim(BREC!SGI_RAZAOSOC) & vbTab & _
                     BREC!SGI_IDPRODUTO & vbTab & _
                     BREC!SGI_CODTIPO & vbTab & _
                     IIf(BREC!SGI_CODTIPO = 1, "NORMAL", "HOMOLOGADA") & vbTab & _
                     BREC!SGI_TIPOP & vbTab & _
                     IIf(BREC!SGI_JAIMPRESSA = 1, "SIM", "NÃO")
        
            Call MudaCorCelula(grdFINALIZADO, BREC!SGI_JAIMPRESSA, (.Rows - 1))
            
            BREC.MoveNext
        Loop
        BREC.Close

    End With
End Sub

Private Function VerificaOrdFat(strCODORD As String) As Boolean

    VerificaOrdFat = False
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADORDFATI " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODORDFAB = " & Trim(strCODORD)
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then VerificaOrdFat = True
    BREC10.Close
    
End Function

Private Sub VerifOrdFatAberto(strCODPED As String)

    Dim intLinha    As Integer
    Dim arrORDFAT() As String
    
    objCADOORDFAB.ORDFAT = Empty
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADORDFATH" & strNOMFILIAL & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODPED = " & Trim(strCODPED) & vbCrLf
    sSql = sSql & "   And SGI_STATUS = 0"

    BREC7.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC7.EOF() Then
        Do While Not BREC7.EOF()
            intLinha = intLinha + 1
            ReDim Preserve arrORDFAT(1 To intLinha) As String
            arrORDFAT(intLinha) = Trim(Str(BREC7!SGI_CODORD))
            BREC7.MoveNext
        Loop
        objCADOORDFAB.ORDFAT = arrORDFAT
    End If
    BREC7.Close
    
End Sub

Private Function FechaPedido(strCODPEDIDO As String) As Boolean

    
    FechaPedido = False
    
    Dim curQTDEPEDTOT As Currency
    Dim curQTDABERTO  As Currency
    Dim curQTDEJAFAT  As Currency
    Dim curQTDEFECHA  As Currency
    Dim curSALDO      As Currency
    Dim strSTATUS     As String
    
    curQTDEPEDTOT = 0
    curQTDABERTO = 0
    curQTDEJAFAT = 0
    curQTDEFECHA = 0
    curSALDO = 0
    strSTATUS = ""
    
    '' Pega Qtde Geral do Pedido
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_QTDEITENSPEDIDO" & vbCrLf
    sSql = sSql & "      ,SGI_STATUS" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPEDVENDH " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & strCODPEDIDO
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
       curQTDEPEDTOT = BREC!SGI_QTDEITENSPEDIDO
       strSTATUS = BREC!SGI_STATUS
    End If
    BREC.Close
    
    '' Pega OPS em Aberto
    sSql = "Select Sum(SGI_QTDEPED) As SGI_QTDEPED " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "  And  SGI_CODPED = " & strCODPEDIDO & vbCrLf
    sSql = sSql & "  And  SGI_STATUS = 0"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        If Not IsNull(BREC!SGI_QTDEPED) Then curQTDABERTO = BREC!SGI_QTDEPED
    End If
    BREC.Close
    
    '' Pega OPS em Aberto/Parcial
    sSql = "Select Sum(SGI_QTDFAT) As SGI_QTDFAT " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "  And  SGI_CODPED = " & strCODPEDIDO & vbCrLf
    sSql = sSql & "  And  SGI_STATUS = 1"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        If Not IsNull(BREC!SGI_QTDFAT) Then curQTDEJAFAT = BREC!SGI_QTDFAT
    End If
    BREC.Close
    
    '' Pega OPS Fechadas
    sSql = "Select Sum(SGI_QTDFAT) As SGI_QTDFAT " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "  And  SGI_CODPED = " & strCODPEDIDO & vbCrLf
    sSql = sSql & "  And  SGI_STATUS = 2"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        If Not IsNull(BREC!SGI_QTDFAT) Then curQTDEFECHA = BREC!SGI_QTDFAT
    End If
    BREC.Close
    
    
    curSALDO = (curQTDEPEDTOT - (curQTDABERTO + curQTDEJAFAT + curQTDEFECHA))
    If curSALDO <= 0 Then
       FechaPedido = True
    End If
    
End Function

Private Sub DestroiObjeto()
    If adoBanco_Dados.State = 1 Then adoBanco_Dados.Close
    Set objFuncoes = Nothing
    Set objCADOORDFAB = Nothing
    Set objRel = Nothing
    Set objRelDireto = Nothing
End Sub

Private Sub Seleciona(intIndice As Integer)
    Dim I As Integer
    With grdOrdeFab
        For I = 1 To (.Rows - 1)
            If intIndice = 0 Then
               .Cell(flexcpText, I, conCOL_OrdFab_Selecionado) = 0
            ElseIf intIndice = 1 Then
              .Cell(flexcpText, I, conCOL_OrdFab_Selecionado) = -1
            End If
        Next I
    End With
End Sub

Private Sub MudaCorCelulaAlterada(grdGENERICA As VSFlexGrid, lngID As Long, lngROW As Long)
    With grdGENERICA
        If lngID = 1 Then
            .Cell(flexcpBackColor, lngROW, conCOL_OrdFab_Codigo, lngROW, conCOL_OrdFab_NomeFilialEmp) = &H80FFFF       '' Amarelo
        End If
    End With
End Sub

Private Function VerificaPedido(strCODPED As String, strIDPRODUTO As String) As Boolean

    VerificaPedido = False
    
    Dim lngQTDETOTALPEDIDO  As Long
    Dim lngQTDEOPS          As Long
    Dim lngQTDETORULOS      As Long
    Dim intRESP             As Integer
    
    lngQTDETOTALPEDIDO = 0
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       Sum(PEDI.SGI_QTDE) As SGI_QTDE" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPEDVENDI" & strNOMFILIAL & " PEDI" & vbCrLf
    sSql = sSql & "  Where" & vbCrLf
    sSql = sSql & "       PEDI.SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And PEDI.SGI_CODIGO    = " & Trim(strCODPED) & vbCrLf
    ''sSql = sSql & "   And PEDI.SGI_IDPRODUTO = " & Trim(strIDPRODUTO)
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        If Not IsNull(BREC!SGI_QTDE) Then lngQTDETOTALPEDIDO = BREC!SGI_QTDE
    End If
    BREC.Close


    '' Pega Quantas OP's Existem para este pedido e Produto
    lngQTDEOPS = 0
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       Count(*) As SGI_QTDEOP" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD" & strNOMFILIAL & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_IDPRODUTO = " & Trim(strIDPRODUTO) & vbCrLf
    sSql = sSql & "   And SGI_CODPED    = " & Trim(strCODPED)

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        If Not IsNull(BREC!SGI_QTDEOP) Then lngQTDEOPS = BREC!SGI_QTDEOP
    End If
    BREC.Close

    If lngQTDEOPS > 1 Then
       intRESP = MsgBox("ATENÇÃO" & vbCrLf & _
                        "Para este rótulo existe(m) " & Format(lngQTDEOPS, "##00") & " OP's." & vbCrLf & _
                        "Deseja realmente baixar esta OP. ?", vbExclamation + vbYesNo + vbDefaultButton2, "Aviso")
        
       If intRESP = vbNo Then Exit Function
    End If


    '' Verificando se existe Rótulos Diferentes para este pedido
    lngQTDETORULOS = 0
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       Count(*) As SGI_QTDEROTULOS" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPEDVENDI" & strNOMFILIAL & " PEDI" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       PEDI.SGI_FILIAL    =  " & FILIAL & vbCrLf
    sSql = sSql & "   And PEDI.SGI_CODIGO    =  " & Trim(strCODPED) & vbCrLf
    sSql = sSql & "   And PEDI.SGI_IDPRODUTO <> " & Trim(strIDPRODUTO)

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        If Not IsNull(BREC!SGI_QTDEROTULOS) Then lngQTDETORULOS = BREC!SGI_QTDEROTULOS
    End If
    BREC.Close
    
    If lngQTDETORULOS <> 0 Then
       intRESP = MsgBox("ATENÇÃO" & vbCrLf & _
                        "Para este pedido existe(m) " & Format(lngQTDETORULOS, "##00") & " Rótulo(s) que diferem do Rótulo que vc esta querendo baixar." & vbCrLf & _
                        "Deseja realmente baixar esta OP. ?", vbExclamation + vbYesNo + vbDefaultButton2, "Aviso")
        
       If intRESP = vbNo Then Exit Function
    End If


    
    VerificaPedido = True
    
    '' Pegando as OPs para poder Baixar

End Function

Private Sub SelTodasSemImp(intIndex As Integer)

    Dim I As Long
    
    With grdOrdeFab
    
        If intIndex = 1 Then
        
            For I = 1 To (.Rows - 1)
                If .Cell(flexcpText, I, conCOL_OrdFab_JaImpressa) = "NÃO" Then
                   .RowHidden(I) = False
                    If .Cell(flexcpChecked, I, conCOL_OrdFab_Selecionado) = 2 Then .Cell(flexcpChecked, I, conCOL_OrdFab_Selecionado) = -1
                Else
                    .RowHidden(I) = True
                End If
            Next I
        
        ElseIf intIndex = 0 Then
        
            For I = 1 To (.Rows - 1)
                .RowHidden(I) = False
                .Cell(flexcpChecked, I, conCOL_OrdFab_Selecionado) = 2
            Next I
        
        End If
    End With

End Sub


Private Sub txtCODCLIE_GotFocus()
    objFuncoes.SelecionaCampos txtCODCLIE.Name, Me
End Sub

Private Sub txtCODCLIE_KeyPress(KeyAscii As Integer)
    objFuncoes.SoNumeroPonto KeyAscii, txtCODCLIE.Text
End Sub

Private Sub txtCODOP_GotFocus()
    objFuncoes.SelecionaCampos txtCODOP.Name, Me
End Sub

Private Sub txtCODOP_KeyPress(KeyAscii As Integer)
    objFuncoes.SoNumeroPonto KeyAscii, txtCODOP.Text
End Sub

Private Sub txtCODPED_GotFocus()
    objFuncoes.SelecionaCampos txtCODPED.Name, Me
End Sub

Private Sub txtCODPED_KeyPress(KeyAscii As Integer)
    objFuncoes.SoNumeroPonto KeyAscii, txtCODPED.Text
End Sub

Private Sub txtCODROTULO_GotFocus()
    objFuncoes.SelecionaCampos txtCODROTULO.Name, Me
End Sub

Private Sub txtCODROTULO_KeyPress(KeyAscii As Integer)
    objFuncoes.SoNumeroPonto KeyAscii, txtCODROTULO.Text
End Sub

Private Sub txtRAZAOSOC_GotFocus()
    objFuncoes.SelecionaCampos txtRAZAOSOC.Name, Me
End Sub

Private Sub txtRAZAOSOC_KeyPress(KeyAscii As Integer)
    KeyAscii = objFuncoes.Maiuscula(KeyAscii)
End Sub

Private Sub txtRotulo_GotFocus()
    objFuncoes.SelecionaCampos txtRotulo.Name, Me
End Sub

Private Sub txtRotulo_KeyPress(KeyAscii As Integer)
    KeyAscii = objFuncoes.Maiuscula(KeyAscii)
End Sub

Private Function ConfereData() As Boolean
    ConfereData = True
    
    
    If Len(Trim(Replace(Replace(mskDTINI.Text, "/", ""), "_", ""))) > 0 And Len(Trim(Replace(Replace(mskDTFIN.Text, "/", ""), "_", ""))) > 0 Then
        
        If Not IsDate(mskDTINI.Text) Then
            MsgBox "DATA INVÀLIDA !!!", vbOKOnly + vbExclamation, "Aviso"
            mskDTINI.SetFocus
            ConfereData = False
            Exit Function
        End If
        If Not IsDate(mskDTFIN.Text) Then
            MsgBox "DATA INVÀLIDA !!!", vbOKOnly + vbExclamation, "Aviso"
            mskDTFIN.SetFocus
            ConfereData = False
            Exit Function
        End If
        
        If CDate(mskDTINI.Text) > CDate(mskDTFIN.Text) Then
            MsgBox "Data Inicial não pode ser maior que data final !!!", vbOKOnly + vbExclamation, "Aviso"
            mskDTINI.SetFocus
            Exit Function
        End If
    
    End If
            
End Function

Private Function PermiteOPQualidade() As Boolean
    
    PermiteOPQualidade = False
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       SGI_MOP" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_USUARIO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCodUsuaro
    
    BREC12.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC12.EOF() Then
        If BREC12!SGI_MOP = 1 Then PermiteOPQualidade = True
    End If
    BREC12.Close
    
End Function
