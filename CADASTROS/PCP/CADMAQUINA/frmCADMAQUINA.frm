VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCADMAQUINA 
   Caption         =   "Cadastro de Máquinas"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13215
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   13215
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab StMaquinas 
      Height          =   6975
      Left            =   0
      TabIndex        =   24
      Top             =   960
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   12303
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
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
      TabCaption(0)   =   "Dados da Maquina"
      TabPicture(0)   =   "frmCADMAQUINA.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Fonte Alimentação"
      TabPicture(1)   =   "frmCADMAQUINA.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame10"
      Tab(1).Control(1)=   "Frame12"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Turnos"
      TabPicture(2)   =   "frmCADMAQUINA.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame11"
      Tab(2).Control(1)=   "Frame13"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Manutenção"
      TabPicture(3)   =   "frmCADMAQUINA.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "StPecas"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Linha Produtos  ( Capacidade )"
      TabPicture(4)   =   "frmCADMAQUINA.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "grdCAPPROD"
      Tab(4).Control(1)=   "Command27"
      Tab(4).Control(2)=   "Command26"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "Operdores"
      TabPicture(5)   =   "frmCADMAQUINA.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame14"
      Tab(5).Control(1)=   "Frame15"
      Tab(5).ControlCount=   2
      Begin VB.CommandButton Command26 
         Height          =   300
         Left            =   -62280
         Picture         =   "frmCADMAQUINA.frx":00A8
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Exclui a linha da Gride Selecionada"
         Top             =   720
         Width           =   300
      End
      Begin VB.CommandButton Command27 
         Height          =   300
         Left            =   -62280
         Picture         =   "frmCADMAQUINA.frx":01F2
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Inclui uma nova linha na Gride"
         Top             =   360
         Width           =   300
      End
      Begin VSFlex8LCtl.VSFlexGrid grdCAPPROD 
         Height          =   6495
         Left            =   -74880
         TabIndex        =   72
         Top             =   360
         Width           =   12495
         _cx             =   22040
         _cy             =   11456
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
      Begin VB.Frame Frame15 
         Height          =   495
         Left            =   -74880
         TabIndex        =   55
         Top             =   360
         Width           =   12855
         Begin VB.TextBox txtPORUTIL 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7320
            TabIndex        =   64
            Text            =   "txtPORUTIL"
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton Command10 
            Height          =   315
            Left            =   8400
            Picture         =   "frmCADMAQUINA.frx":033C
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton Command9 
            Height          =   315
            Left            =   1920
            Picture         =   "frmCADMAQUINA.frx":043E
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   120
            Width           =   375
         End
         Begin VB.TextBox txtOperadores 
            Height          =   285
            Left            =   840
            MaxLength       =   10
            TabIndex        =   56
            Text            =   "txtOperado"
            Top             =   120
            Width           =   1095
         End
         Begin VB.ComboBox cboOperadores 
            Height          =   315
            Left            =   2280
            TabIndex        =   58
            Text            =   "cboOperadores"
            Top             =   120
            Width           =   4335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "% Uso"
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
            Index           =   20
            Left            =   6720
            TabIndex        =   65
            Top             =   180
            Width           =   540
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
            Index           =   13
            Left            =   120
            TabIndex        =   60
            Top             =   180
            Width           =   660
         End
      End
      Begin VB.Frame Frame14 
         Height          =   6015
         Left            =   -74880
         TabIndex        =   53
         Top             =   840
         Width           =   12855
         Begin MSFlexGridLib.MSFlexGrid flxOPERADORES 
            Height          =   5775
            Left            =   120
            TabIndex        =   54
            Top             =   120
            Width           =   12615
            _ExtentX        =   22251
            _ExtentY        =   10186
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame13 
         Height          =   6015
         Left            =   -74880
         TabIndex        =   49
         Top             =   840
         Width           =   12855
         Begin MSFlexGridLib.MSFlexGrid flxGRIDTURNOS 
            Height          =   5775
            Left            =   120
            TabIndex        =   50
            Top             =   120
            Width           =   12615
            _ExtentX        =   22251
            _ExtentY        =   10186
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame12 
         Height          =   6015
         Left            =   -74880
         TabIndex        =   48
         Top             =   840
         Width           =   12855
         Begin MSFlexGridLib.MSFlexGrid flxFONTEALIM 
            Height          =   5775
            Left            =   120
            TabIndex        =   16
            Top             =   120
            Width           =   12615
            _ExtentX        =   22251
            _ExtentY        =   10186
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame11 
         Height          =   495
         Left            =   -74880
         TabIndex        =   45
         Top             =   360
         Width           =   12855
         Begin VB.TextBox txtPORCUSO 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   8280
            TabIndex        =   63
            Text            =   "txtPORCUSO"
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton Command8 
            Height          =   315
            Left            =   9360
            Picture         =   "frmCADMAQUINA.frx":0540
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton Command7 
            Height          =   315
            Left            =   1920
            Picture         =   "frmCADMAQUINA.frx":0642
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   120
            Width           =   375
         End
         Begin VB.TextBox txtCodTurn 
            Height          =   285
            Left            =   840
            MaxLength       =   10
            TabIndex        =   17
            Text            =   "txtCodTurn"
            Top             =   120
            Width           =   1095
         End
         Begin VB.ComboBox cboCodTurn 
            Height          =   315
            Left            =   2280
            TabIndex        =   18
            Text            =   "cboCodTurn"
            Top             =   120
            Width           =   5175
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "% Uso"
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
            Index           =   19
            Left            =   7560
            TabIndex        =   62
            Top             =   180
            Width           =   540
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
            Index           =   5
            Left            =   120
            TabIndex        =   47
            Top             =   180
            Width           =   660
         End
      End
      Begin VB.Frame Frame10 
         Height          =   495
         Left            =   -74880
         TabIndex        =   42
         Top             =   360
         Width           =   12855
         Begin VB.ComboBox cboFontAlim 
            Height          =   315
            Left            =   2280
            TabIndex        =   14
            Text            =   "cboFontAlim"
            Top             =   120
            Width           =   5175
         End
         Begin VB.TextBox txtFontAlim 
            Height          =   285
            Left            =   840
            MaxLength       =   10
            TabIndex        =   13
            Text            =   "txtFontAli"
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton Command6 
            Height          =   315
            Left            =   1920
            Picture         =   "frmCADMAQUINA.frx":0744
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton Command5 
            Height          =   315
            Left            =   7440
            Picture         =   "frmCADMAQUINA.frx":0846
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   120
            Width           =   375
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
            Index           =   12
            Left            =   120
            TabIndex        =   44
            Top             =   180
            Width           =   660
         End
      End
      Begin VB.Frame Frame2 
         Height          =   6495
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   12855
         Begin VB.Frame Frame4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2520
            TabIndex        =   79
            Top             =   3480
            Width           =   2295
            Begin VB.OptionButton optEmpresa 
               Caption         =   "STEEL"
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
               Left            =   1320
               TabIndex        =   81
               Top             =   0
               Width           =   1215
            End
            Begin VB.OptionButton optEmpresa 
               Caption         =   "NOVALATA"
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
               Left            =   0
               TabIndex        =   80
               Top             =   0
               Width           =   1335
            End
         End
         Begin VB.TextBox txtCADGRLIN 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4920
            TabIndex        =   77
            Text            =   "txtCADGRLIN"
            Top             =   3480
            Width           =   735
         End
         Begin VB.CommandButton Command1 
            Height          =   315
            Left            =   5640
            Picture         =   "frmCADMAQUINA.frx":0948
            Style           =   1  'Graphical
            TabIndex        =   76
            Top             =   3480
            Width           =   375
         End
         Begin VB.TextBox txtFraOpeadores 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   7560
            TabIndex        =   12
            Text            =   "txtFraOpeadores"
            Top             =   3120
            Width           =   1455
         End
         Begin VB.TextBox txtQtdMaquinas 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4920
            TabIndex        =   11
            Text            =   "txtQtdMaquinas"
            Top             =   3120
            Width           =   1215
         End
         Begin VB.CommandButton Command11 
            Height          =   315
            Left            =   5640
            Picture         =   "frmCADMAQUINA.frx":0A4A
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   2040
            Width           =   375
         End
         Begin VB.TextBox txtCODFAMILIA 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4920
            TabIndex        =   6
            Text            =   "txtCODFAMILIA"
            Top             =   2040
            Width           =   735
         End
         Begin VB.ComboBox cboFAMILIA 
            Height          =   315
            Left            =   6000
            TabIndex        =   8
            Text            =   "cboFAMILIA"
            Top             =   2040
            Width           =   4455
         End
         Begin VB.ComboBox cboUnidade 
            Height          =   315
            Left            =   4920
            TabIndex        =   10
            Text            =   "cboUnidade"
            Top             =   2760
            Width           =   855
         End
         Begin VB.TextBox txtCAPMIN 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4920
            TabIndex        =   9
            Text            =   "txtCAPMIN"
            Top             =   2400
            Width           =   1215
         End
         Begin VB.TextBox txtNumAtivo 
            Height          =   285
            Left            =   4920
            MaxLength       =   20
            TabIndex        =   5
            Text            =   "txtNumAtivo"
            Top             =   1680
            Width           =   1215
         End
         Begin VB.TextBox txtAnoFabr 
            Height          =   285
            Left            =   4920
            MaxLength       =   20
            TabIndex        =   4
            Text            =   "txtAnoFabr"
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Frame Frame8 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   4800
            TabIndex        =   32
            Top             =   960
            Width           =   2055
            Begin VB.OptionButton optAtivoSN 
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
               Index           =   1
               Left            =   1080
               TabIndex        =   3
               Top             =   0
               Width           =   855
            End
            Begin VB.OptionButton optAtivoSN 
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
               Index           =   0
               Left            =   120
               TabIndex        =   2
               Top             =   0
               Width           =   735
            End
         End
         Begin VB.TextBox txtDescricao 
            Height          =   285
            Left            =   4920
            MaxLength       =   30
            TabIndex        =   1
            Text            =   "txtDescricao"
            Top             =   600
            Width           =   5295
         End
         Begin VB.Label lblDESGRPMAQ 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDESGRPMAQ"
            Height          =   255
            Left            =   6000
            TabIndex        =   78
            Top             =   3480
            Width           =   5415
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Grupo de Linha de Produto"
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
            Index           =   10
            Left            =   120
            TabIndex        =   75
            Top             =   3480
            Width           =   2310
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Operadores"
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
            Index           =   22
            Left            =   6360
            TabIndex        =   67
            Top             =   3120
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "1 Operador Cuida de quantas máquinas deste tipo"
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
            Index           =   21
            Left            =   120
            TabIndex        =   66
            Top             =   3120
            Width           =   4260
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   18
            Left            =   120
            TabIndex        =   61
            Top             =   2040
            Width           =   1650
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   52
            Top             =   2760
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Capacidade Por Hora"
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
            TabIndex        =   51
            Top             =   2400
            Width           =   1830
         End
         Begin VB.Label lblCodigo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblCodigo"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4920
            TabIndex        =   0
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nº Ativo Fixo:"
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
            Index           =   8
            Left            =   120
            TabIndex        =   31
            Top             =   1680
            Width           =   1185
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ano Fabricação"
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
            Left            =   120
            TabIndex        =   30
            Top             =   1320
            Width           =   1350
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
            TabIndex        =   29
            Top             =   240
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Descrição:"
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
            TabIndex        =   28
            Top             =   600
            Width           =   930
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ativa:"
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
            TabIndex        =   27
            Top             =   960
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   2280
            TabIndex        =   26
            Top             =   960
            Width           =   75
         End
      End
      Begin TabDlg.SSTab StPecas 
         Height          =   6495
         Left            =   -74880
         TabIndex        =   33
         Top             =   360
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   11456
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
         TabCaption(0)   =   "Peças"
         TabPicture(0)   =   "frmCADMAQUINA.frx":0B4C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame6"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame3"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Manutenção Preventiva"
         TabPicture(1)   =   "frmCADMAQUINA.frx":0B68
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Command3"
         Tab(1).Control(1)=   "Command4"
         Tab(1).Control(2)=   "grdAgendaManut"
         Tab(1).Control(3)=   "grdParamManut"
         Tab(1).ControlCount=   4
         Begin VB.CommandButton Command3 
            Height          =   315
            Left            =   -59880
            Picture         =   "frmCADMAQUINA.frx":0B84
            Style           =   1  'Graphical
            TabIndex        =   71
            Top             =   360
            Width           =   375
         End
         Begin VB.CommandButton Command4 
            Height          =   315
            Left            =   -59880
            Picture         =   "frmCADMAQUINA.frx":0CC2
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   720
            Width           =   375
         End
         Begin VSFlex8LCtl.VSFlexGrid grdAgendaManut 
            Height          =   2295
            Left            =   -74880
            TabIndex        =   69
            Top             =   2040
            Width           =   14895
            _cx             =   26273
            _cy             =   4048
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
         Begin VSFlex8LCtl.VSFlexGrid grdParamManut 
            Height          =   1575
            Left            =   -74880
            TabIndex        =   68
            Top             =   360
            Width           =   14895
            _cx             =   26273
            _cy             =   2778
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
         Begin VB.Frame Frame3 
            Height          =   495
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   12615
            Begin VB.CommandButton cmbGravEsp 
               Height          =   315
               Left            =   7320
               Picture         =   "frmCADMAQUINA.frx":124C
               Style           =   1  'Graphical
               TabIndex        =   40
               Top             =   120
               Width           =   375
            End
            Begin VB.CommandButton cmdPesq 
               Height          =   315
               Left            =   1620
               Picture         =   "frmCADMAQUINA.frx":134E
               Style           =   1  'Graphical
               TabIndex        =   39
               Top             =   120
               Width           =   375
            End
            Begin VB.TextBox txtCodPecas 
               Height          =   285
               Left            =   840
               MaxLength       =   10
               TabIndex        =   38
               Text            =   "txtCodPeca"
               Top             =   120
               Width           =   750
            End
            Begin VB.ComboBox cboCodPecas 
               Height          =   315
               Left            =   2040
               TabIndex        =   37
               Text            =   "cboCodPecas"
               Top             =   120
               Width           =   5295
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
               Index           =   9
               Left            =   120
               TabIndex        =   41
               Top             =   180
               Width           =   660
            End
         End
         Begin VB.Frame Frame6 
            Height          =   5535
            Left            =   120
            TabIndex        =   34
            Top             =   840
            Width           =   12615
            Begin MSFlexGridLib.MSFlexGrid flxPECAS 
               Height          =   5295
               Left            =   120
               TabIndex        =   35
               Top             =   120
               Width           =   12375
               _ExtentX        =   21828
               _ExtentY        =   9340
               _Version        =   393216
               FixedCols       =   0
               HighLight       =   2
               SelectionMode   =   1
               Appearance      =   0
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   13095
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
         Picture         =   "frmCADMAQUINA.frx":1450
         Style           =   1  'Graphical
         TabIndex        =   23
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
         Picture         =   "frmCADMAQUINA.frx":1982
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
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
         Picture         =   "frmCADMAQUINA.frx":1A84
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmCADMAQUINA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public cTipOper         As String
Public iCodigo          As Integer
Public FILIAL           As Integer
Public strAcesso        As String

Dim objBLBFunc          As Object
Dim objCADMAQUINA       As Object
Dim objPESQPADRAO       As Object
Dim arrGRDFONTALIM      As Variant
Dim arrGRDTURNOS        As Variant
Dim arrGRDPRODUTOS      As Variant
Dim arrGRDOPERADOR      As Variant
Dim arrPARAMETROS       As Variant
Dim arrDIASPARAMETROS   As Variant
Dim arrLINHACAP         As Variant

Const conCOL_SonParamManut_TipoManut            As Integer = 0
Const conCOL_SonParammanut_DataInici            As Integer = 1
Const conCOL_SonParammanut_DataFinal            As Integer = 2
Const conCOL_SonParamManut_HoraInici            As Integer = 3
Const conCOL_SonParamManut_HoraFinal            As Integer = 4
Const conCOL_SonParamManut_TempoUsado           As Integer = 5
Const conCOL_SonParamManut_MaqUso               As Integer = 6
Const conCOL_SonParamManut_Ativo                As Integer = 7
Const conCOL_SonParamManut_EmConjunto           As Integer = 8
Const conCOL_SonParamManut_Indice               As Integer = 9
Const conCOL_SonParamManut_FormatString         As String = "=Tipo de Manutenção|Data Inicial|Data Final|Hora Inicial|Hora Final|Tempo Usado|Maq. Uso|Ativo|Em Conjunto|Indice"
Const conColumnsIn_SonParamManut                As Integer = 10

Const conCOL_SonDiasManut_DataManut             As Integer = 0
Const conCOL_SonDiasManut_HoraInic              As Integer = 1
Const conCOL_SonDiasManut_HoraFina              As Integer = 2
Const conCOL_SonDiasManut_Tempo                 As Integer = 3
Const conCOL_SonDiasManut_MaqUso                As Integer = 4
Const conCOL_SonDiasManut_Ativo                 As Integer = 5
Const conCOL_SonDiasManut_EmConjunto            As Integer = 6
Const conCOL_SonDiasManut_Pai                   As Integer = 7
Const conCOL_SonDiasManut_Indice                As Integer = 8
Const conCOL_SonDiasManut_DtParametro           As Integer = 9
Const conCOL_SonDiasManut_FormatString          As String = "=Dt. Manutenção|Hora Inicial|Hora Final|Tempo Usado|Maq. Uso|Ativo|Em Conjunto|Pai|Indice|Dtparametro"
Const conColumnsIn_SonDiasManut                 As Integer = 10

Const conCOL_SonCap_Itens                       As Integer = 0
Const conCOL_SonCap_IDLinha                     As Integer = 1
Const conCOL_SonCap_CodLinha                    As Integer = 2
Const conCOL_SonCap_PesqLinha                   As Integer = 3
Const conCOL_SonCap_DescLinha                   As Integer = 4
Const conCOL_SonCap_CodSeqCorte                 As Integer = 5
Const conCOL_SonCap_CodCorte                    As Integer = 6
Const conCOL_SonCap_PesqCorte                   As Integer = 7
Const conCOL_SonCap_DescCorte                   As Integer = 8
Const conCOL_SonCap_INDICE                      As Integer = 9
Const conCOL_SonCap_FormatString                As String = "=Seg|IDLINHA|Código|...|Descrição da Capacidade|Seg.|Cod.Corte|...|Descrição do Corte|INDICE"
Const conColumnsIn_SonCap                       As Integer = 10


Private Sub cboCodPecas_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboCodPecas, KeyAscii
End Sub

Private Sub cboCodTurn_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboCodTurn, KeyAscii
End Sub

Private Sub cboCodTurn_Validate(Cancel As Boolean)
    If cboCodTurn.ListIndex > -1 Then txtCodTurn.Text = cboCodTurn.ItemData(cboCodTurn.ListIndex)
End Sub

Private Sub cboFAMILIA_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboFAMILIA, KeyAscii
End Sub

Private Sub cboFAMILIA_Validate(Cancel As Boolean)
    If cboFAMILIA.ListIndex > -1 Then txtCODFAMILIA.Text = Str(cboFAMILIA.ItemData(cboFAMILIA.ListIndex))
End Sub

Private Sub cboFontAlim_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboFontAlim, KeyAscii
End Sub

Private Sub cboFontAlim_Validate(Cancel As Boolean)
    If cboFontAlim.ListIndex > -1 Then txtFontAlim.Text = cboFontAlim.ItemData(cboFontAlim.ListIndex)
End Sub

Private Sub cboOperadores_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboOperadores, KeyAscii
End Sub

Private Sub cboOperadores_Validate(Cancel As Boolean)
    If cboOperadores.ListIndex > -1 Then txtOperadores.Text = cboOperadores.ItemData(cboOperadores.ListIndex)
End Sub




Private Sub cmdAltera_Click()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame10.Enabled = True
    Frame11.Enabled = True
    Frame3.Enabled = True
    Frame15.Enabled = True
    
    Me.Caption = "Cadastro de Máquinas - [ ALTERAÇÃO ]"
    
    cTipOper = "A"
    
    StMaquinas.Tab = 0
    StPecas.Tab = 0
    
    txtDescricao.SetFocus
    

End Sub

Private Sub CmdSalva_Click()
    
    Dim i As Integer
    
    If ValidaCampos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADMAQUINA.CODIGO = objCADMAQUINA.Gera_Codigo(Me.Name)
    
    objCADMAQUINA.DESCRI = txtDescricao.Text
    If optAtivoSN(0).Value = True Then objCADMAQUINA.OPTATIVO = 0
    If optAtivoSN(1).Value = True Then objCADMAQUINA.OPTATIVO = 1
    objCADMAQUINA.ANOFABR = txtAnoFabr.Text
    objCADMAQUINA.NUMATIVO = txtNumAtivo.Text
    objCADMAQUINA.CAPMINUTO = CCur(txtCAPMIN.Text)
    objCADMAQUINA.UNIMED = cboUnidade.ItemData(cboUnidade.ListIndex)
    objCADMAQUINA.FAMILIA = cboFAMILIA.ItemData(cboFAMILIA.ListIndex)
    
    If Len(Trim(txtQtdMaquinas.Text)) > 0 Then
       objCADMAQUINA.QtdOperadores = CInt(txtQtdMaquinas.Text)
       objCADMAQUINA.FraOperadores = CCur(txtFraOpeadores.Text)
    End If
    
    '' Fontes de Alimentação
    If (flxFONTEALIM.Rows - 1) > 0 Then
       ReDim arrGRDFONTALIM(1 To (flxFONTEALIM.Rows - 1)) As String
       For i = 1 To (flxFONTEALIM.Rows - 1)
           arrGRDFONTALIM(i) = flxFONTEALIM.TextMatrix(i, 1)
       Next i
       objCADMAQUINA.GRDFONTALIM = arrGRDFONTALIM
    Else
       ReDim arrGRDFONTALIM(0) As String
       objCADMAQUINA.GRDFONTALIM = arrGRDFONTALIM
    End If
    
    '' Turnos
    If (flxGRIDTURNOS.Rows - 1) > 0 Then
       ReDim arrGRDTURNOS(1 To (flxGRIDTURNOS.Rows - 1), 1 To 2) As String
       For i = 1 To (flxGRIDTURNOS.Rows - 1)
           arrGRDTURNOS(i, 1) = flxGRIDTURNOS.TextMatrix(i, 1)
           arrGRDTURNOS(i, 2) = flxGRIDTURNOS.TextMatrix(i, 3)
       Next i
       objCADMAQUINA.GRDTURNOS = arrGRDTURNOS
    Else
       ReDim arrGRDTURNOS(0) As String
       objCADMAQUINA.GRDTURNOS = arrGRDTURNOS
    End If
    
    '' Operadores
    If (flxOPERADORES.Rows - 1) > 0 Then
       ReDim arrGRDOPERADOR(1 To (flxOPERADORES.Rows - 1), 1 To 2) As String
       For i = 1 To (flxOPERADORES.Rows - 1)
           arrGRDOPERADOR(i, 1) = flxOPERADORES.TextMatrix(i, 1)
           arrGRDOPERADOR(i, 2) = flxOPERADORES.TextMatrix(i, 3)
       Next i
       objCADMAQUINA.GRDOPERADOR = arrGRDOPERADOR
    Else
       ReDim arrGRDOPERADOR(0) As String
       objCADMAQUINA.GRDOPERADOR = arrGRDOPERADOR
    End If
    
    '' Manutenção
    arrPARAMETROS = Empty
    If (grdParamManut.Rows - 1) > 0 Then
        ReDim arrPARAMETROS(1 To (grdParamManut.Rows - 1), 1 To 9) As String
        For i = 1 To (grdParamManut.Rows - 1)
            arrPARAMETROS(i, 1) = grdParamManut.Cell(flexcpText, i, conCOL_SonParamManut_TipoManut)
            arrPARAMETROS(i, 2) = grdParamManut.Cell(flexcpText, i, conCOL_SonParammanut_DataInici)
            arrPARAMETROS(i, 3) = grdParamManut.Cell(flexcpText, i, conCOL_SonParammanut_DataFinal)
            arrPARAMETROS(i, 4) = grdParamManut.Cell(flexcpText, i, conCOL_SonParamManut_HoraInici)
            arrPARAMETROS(i, 5) = grdParamManut.Cell(flexcpText, i, conCOL_SonParamManut_HoraFinal)
            arrPARAMETROS(i, 6) = grdParamManut.Cell(flexcpText, i, conCOL_SonParamManut_TempoUsado)
            arrPARAMETROS(i, 7) = IIf(grdParamManut.Cell(flexcpTextDisplay, i, conCOL_SonParamManut_MaqUso) = "Sim", 1, 0)
            arrPARAMETROS(i, 8) = IIf(grdParamManut.Cell(flexcpTextDisplay, i, conCOL_SonParamManut_Ativo) = "Sim", 1, 0)
            arrPARAMETROS(i, 9) = IIf(grdParamManut.Cell(flexcpTextDisplay, i, conCOL_SonParamManut_EmConjunto) = "Sim", 1, 0)
        Next i
    End If
    objCADMAQUINA.PARAMETROS = arrPARAMETROS
    
    arrDIASPARAMETROS = Empty
    If (grdAgendaManut.Rows - 1) > 0 Then
       ReDim arrDIASPARAMETROS(1 To (grdAgendaManut.Rows - 1), 1 To 9) As String
       For i = 1 To (grdAgendaManut.Rows - 1)
           arrDIASPARAMETROS(i, 1) = grdAgendaManut.Cell(flexcpText, i, conCOL_SonDiasManut_DataManut)
           arrDIASPARAMETROS(i, 2) = grdAgendaManut.Cell(flexcpText, i, conCOL_SonDiasManut_HoraInic)
           arrDIASPARAMETROS(i, 3) = grdAgendaManut.Cell(flexcpText, i, conCOL_SonDiasManut_HoraFina)
           arrDIASPARAMETROS(i, 4) = grdAgendaManut.Cell(flexcpText, i, conCOL_SonDiasManut_Tempo)
           arrDIASPARAMETROS(i, 5) = IIf(grdAgendaManut.Cell(flexcpTextDisplay, i, conCOL_SonDiasManut_MaqUso) = "Sim", 1, 0)
           arrDIASPARAMETROS(i, 6) = IIf(grdAgendaManut.Cell(flexcpTextDisplay, i, conCOL_SonDiasManut_Ativo) = "Sim", 1, 0)
           arrDIASPARAMETROS(i, 7) = IIf(grdAgendaManut.Cell(flexcpTextDisplay, i, conCOL_SonDiasManut_EmConjunto) = "Sim", 1, 0)
           arrDIASPARAMETROS(i, 8) = grdAgendaManut.Cell(flexcpText, i, conCOL_SonDiasManut_Indice)
           arrDIASPARAMETROS(i, 9) = grdAgendaManut.Cell(flexcpText, i, conCOL_SonDiasManut_DtParametro)
       Next i
    End If
    objCADMAQUINA.DIASPARAMETROS = arrDIASPARAMETROS
    
    '' -------------------------------------
    '' Capacidade (Linhas)
    arrLINHACAP = Empty
    With grdCAPPROD
        If (.Rows - 1) > 0 Then
            ReDim arrLINHACAP(1 To (.Rows - 1), 1 To 6) As String
            For i = 1 To (.Rows - 1)
                arrLINHACAP(i, 1) = .Cell(flexcpText, i, conCOL_SonCap_Itens)
                arrLINHACAP(i, 2) = .Cell(flexcpText, i, conCOL_SonCap_IDLinha)
                arrLINHACAP(i, 3) = .Cell(flexcpText, i, conCOL_SonCap_CodLinha)
                arrLINHACAP(i, 4) = .Cell(flexcpText, i, conCOL_SonCap_CodSeqCorte)
                arrLINHACAP(i, 5) = .Cell(flexcpText, i, conCOL_SonCap_CodCorte)
                arrLINHACAP(i, 6) = .Cell(flexcpText, i, conCOL_SonCap_INDICE)
            Next i
        End If
    End With
    objCADMAQUINA.LINHACAP = arrLINHACAP
    '' -------------------------------------
    
    If objCADMAQUINA.GRAVA(cTipOper) = False Then Exit Sub
    MsgBox "A maquina foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
    
    If cTipOper = "I" Then Call Inclui
    
End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objCADMAQUINA = Nothing
    Set objPESQPADRAO = Nothing
    Unload Me
End Sub


Private Sub Command1_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    
    If optEmpresa(0).Value = True Then
        sSql = sSql & "       SGI_CADGRUPLINHA " & vbCrLf
    Else
        sSql = sSql & "       SGI_CADGRUPLINHA_STEEL" & vbCrLf
    End If
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRI"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "4000"
    arrCAMPOS(2, 5) = "SGI_DESCRI"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Grupo de Linhas")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODFAMILIA.Text = varRETORNO
    
    txtCODFAMILIA.SetFocus
    
End Sub

Private Sub Command10_Click()
    If cTipOper = "I" Or cTipOper = "A" Then Inc_Operadores
End Sub

Private Sub Command11_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADFAMMAQUINAS " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
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
    
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Familia de máquinas")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODFAMILIA.Text = varRETORNO
    
    cboFAMILIA.ListIndex = -1
    txtCODFAMILIA.SetFocus

End Sub


Private Sub Command26_Click()
    If cTipOper = "C" Then Exit Sub
    If cTipOper = "I" Or cTipOper = "A" Then
        With grdCAPPROD
            If (.Rows - 1) = 1 Then .Rows = 1
            If (.Rows - 1) > 1 Then Call objBLBFunc.ExclLinhaGrid(grdCAPPROD, grdCAPPROD.Row)
            Call RefazIndice
        End With
    End If
End Sub

Private Sub Command27_Click()
    If (cTipOper = "I" Or cTipOper = "A") Then Call IncRegGrid
End Sub

Private Sub Command3_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncGrdParam
End Sub


Private Sub Command4_Click()
    Dim strINDICE  As String
    Dim i          As Integer
    If cTipOper = "I" Or cTipOper = "A" Then
VOLTA:
       strINDICE = Trim(grdParamManut.Cell(flexcpText, grdParamManut.Row, conCOL_SonParamManut_Indice))
       For i = 1 To (grdAgendaManut.Rows - 1)
           If Trim(grdAgendaManut.Cell(flexcpText, i, conCOL_SonDiasManut_Pai)) = Trim(strINDICE) Then
              grdAgendaManut.RemoveItem i
              GoTo VOLTA
           End If
       Next i
       Call objBLBFunc.ExclLinhaGrid(grdParamManut, grdParamManut.Row)
    End If
End Sub

Private Sub Command5_Click()
    If cTipOper = "I" Or cTipOper = "A" Then Inc_FontALim
End Sub

Private Sub Command6_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADTIPALIM " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
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
    
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Tipo de Alimentação")
    
    If Len(Trim(varRETORNO)) > 0 Then txtFontAlim.Text = varRETORNO
    
    cboFontAlim.ListIndex = -1
    txtFontAlim.SetFocus

End Sub

Private Sub Command7_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADQTDETURN " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
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
    
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Turnos")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCodTurn.Text = varRETORNO
    
    cboCodTurn.ListIndex = -1
    txtCodTurn.SetFocus

End Sub

Private Sub Command8_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncGridTurn
End Sub

Private Sub Command9_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADOPERADOR " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
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
    
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Operadores")
    
    If Len(Trim(varRETORNO)) > 0 Then txtOperadores.Text = varRETORNO
    
    cboOperadores.ListIndex = -1
    txtOperadores.SetFocus

End Sub

Private Sub flxFONTEALIM_KeyDown(KeyCode As Integer, Shift As Integer)
    If cTipOper = "I" Or cTipOper = "A" Then
       If KeyCode <> vbKeyDelete Then Exit Sub
       If flxFONTEALIM.Rows = 2 Then flxFONTEALIM.Rows = 1
       If flxFONTEALIM.Rows > 2 Then flxFONTEALIM.RemoveItem flxFONTEALIM.RowSel
    End If
End Sub

Private Sub flxGRIDTURNOS_KeyDown(KeyCode As Integer, Shift As Integer)
    If cTipOper = "I" Or cTipOper = "A" Then
       If KeyCode <> vbKeyDelete Then Exit Sub
       If flxGRIDTURNOS.Rows = 2 Then flxGRIDTURNOS.Rows = 1
       If flxGRIDTURNOS.Rows > 2 Then flxGRIDTURNOS.RemoveItem flxGRIDTURNOS.RowSel
    End If
End Sub


Private Sub flxOPERADORES_KeyDown(KeyCode As Integer, Shift As Integer)
    If cTipOper = "I" Or cTipOper = "A" Then
       If KeyCode <> vbKeyDelete Then Exit Sub
       If flxOPERADORES.Rows = 2 Then flxOPERADORES.Rows = 1
       If flxOPERADORES.Rows > 2 Then flxOPERADORES.RemoveItem flxOPERADORES.RowSel
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADMAQUINA = CreateObject("CADMAQUINA.clsCADMAQUINA")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
   objCADMAQUINA.FILIAL = FILIAL
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

End Sub

Private Sub Inclui()

    StMaquinas.Tab = 0
    StPecas.Tab = 0
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame3.Enabled = True
    Frame10.Enabled = True
    Frame11.Enabled = True
    Frame3.Enabled = True
    Frame15.Enabled = True
   
    Me.Caption = "Cadastro de Máquinas - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADMAQUINA
    
    lblCodigo.Caption = ""
    
    ConfGridTipAlim
    ConfGridTurnos
    ConfGridPecas
    CofGridOperadores
    InitGridParametros
    InitGridDiasanutencao
    Call InitGridCapCorte
    Call LimpaCamposLabel

    optAtivoSN.Item(0).Value = True
    optEmpresa(0).Value = True
    
    objCADMAQUINA.PreencheComboTipAlim cboFontAlim
    objCADMAQUINA.PreencheComboTurnos cboCodTurn
    objCADMAQUINA.PreenchComboUnidade cboUnidade
    objCADMAQUINA.PreencheComboOperadores cboOperadores
    objCADMAQUINA.PreenchComboFamilia cboFAMILIA
    
    If txtDescricao.Enabled = True And txtDescricao.Visible = True Then txtDescricao.SetFocus
    
End Sub

Private Sub ConfGridTipAlim()

    flxFONTEALIM.Rows = 1
    flxFONTEALIM.Cols = 3
    
    flxFONTEALIM.TextMatrix(0, 0) = ""
    flxFONTEALIM.TextMatrix(0, 1) = "Código"
    flxFONTEALIM.TextMatrix(0, 2) = "Descrição"
    
    flxFONTEALIM.ColWidth(0) = 0
    flxFONTEALIM.ColWidth(1) = 1000
    flxFONTEALIM.ColWidth(2) = 4000
    
End Sub

Private Sub ConfGridTurnos()

    flxGRIDTURNOS.Rows = 1
    flxGRIDTURNOS.Cols = 4
    
    flxGRIDTURNOS.TextMatrix(0, 0) = ""
    flxGRIDTURNOS.TextMatrix(0, 1) = "Código"
    flxGRIDTURNOS.TextMatrix(0, 2) = "Descrição"
    flxGRIDTURNOS.TextMatrix(0, 3) = "% Uso"
    
    flxGRIDTURNOS.ColWidth(0) = 0
    flxGRIDTURNOS.ColWidth(1) = 1000
    flxGRIDTURNOS.ColWidth(2) = 4000
    flxGRIDTURNOS.ColWidth(3) = 1500
    
End Sub

Private Sub grdAgendaManut_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    Dim strTOTALPERIODO As String
    Dim dtTotalLiquido  As Date
    
    With grdAgendaManut
        Select Case Col
            Case conCOL_SonDiasManut_HoraInic, _
                 conCOL_SonDiasManut_HoraFina
                 If Len(Trim(Replace(.Cell(flexcpText, Row, conCOL_SonDiasManut_HoraInic), ":", ""))) > 0 And _
                    Len(Trim(Replace(.Cell(flexcpText, Row, conCOL_SonDiasManut_HoraFina), ":", ""))) > 0 Then
                    
                    If CDate(.Cell(flexcpText, Row, conCOL_SonDiasManut_HoraInic)) > CDate(.Cell(flexcpText, Row, conCOL_SonDiasManut_HoraFina)) Then
                       MsgBox "Hora de Inicial não pode ser maior que hora final !!!", vbOKOnly + vbExclamation, "Aviso"
                       .Cell(flexcpText, Row, conCOL_SonDiasManut_HoraInic) = ""
                       .Cell(flexcpText, Row, conCOL_SonDiasManut_HoraFina) = ""
                       .Cell(flexcpText, Row, conCOL_SonDiasManut_Tempo) = ""
                       Exit Sub
                    End If
                    
                    strTOTALPERIODO = objBLBFunc.CalcTempo(.Cell(flexcpText, Row, conCOL_SonDiasManut_HoraInic), .Cell(flexcpText, Row, conCOL_SonDiasManut_HoraFina))
                    .Cell(flexcpText, Row, conCOL_SonDiasManut_Tempo) = Format(CDate(strTOTALPERIODO), "HH:MM")
                 End If
        End Select
    End With

End Sub

Private Sub grdAgendaManut_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonDiasManut_Tempo, _
         conCOL_SonDiasManut_DataManut, _
         conCOL_SonDiasManut_HoraInic, _
         conCOL_SonDiasManut_HoraFina
         Cancel = True
    End Select
    Exit Sub
End Sub

Private Sub grdAgendaManut_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    Dim intHoras As Integer
    Dim intMinutos As Integer
    
    With grdAgendaManut
        Select Case Col
               Case conCOL_SonDiasManut_HoraInic, _
                    conCOL_SonDiasManut_HoraFina
                    If .EditText = "  :  " Then
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
        End Select
    End With

End Sub

Private Sub grdCAPPROD_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     Dim i As Integer
     With grdCAPPROD
          Select Case Col
                 Case conCOL_SonCap_CodLinha
          End Select
     End With
End Sub

Private Sub grdCAPPROD_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonCap_Itens, _
         conCOL_SonCap_DescLinha, _
         conCOL_SonCap_CodSeqCorte, _
         conCOL_SonCap_CodCorte, _
         conCOL_SonCap_DescCorte, _
         conCOL_SonCap_IDLinha, _
         conCOL_SonCap_INDICE
         Cancel = True
    Case conCOL_SonCap_CodLinha, _
         conCOL_SonCap_PesqLinha, _
         conCOL_SonCap_PesqCorte
         If cTipOper = "C" Then Cancel = True
    Case Else
        grdCAPPROD.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub grdCAPPROD_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    Dim strDESCPROD As String
    
    With grdCAPPROD
    
        If (.Rows - 1) = 0 Then Exit Sub
        
        Select Case Col
            Case conCOL_SonCap_PesqLinha
        
                If cTipOper = "C" Then Exit Sub
                
                ReDim arrCAMPOS(1 To 2, 1 To 5) As String
                ReDim arrTABELA(1 To 1) As String
                
                sSql = ""
                
                sSql = "Select " & vbCrLf
                sSql = sSql & "       * " & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_CADLINHAPRODUTO " & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       SGI_FILIAL = " & FILIAL
                
                arrTABELA(1) = sSql
                
                arrCAMPOS(1, 1) = "SGI_CODLIN"
                arrCAMPOS(1, 2) = "N"
                arrCAMPOS(1, 3) = "Linha"
                arrCAMPOS(1, 4) = "2000"
                arrCAMPOS(1, 5) = "SGI_CODLIN"
                
                arrCAMPOS(2, 1) = "SGI_DESCRI"
                arrCAMPOS(2, 2) = "S"
                arrCAMPOS(2, 3) = "Descrição"
                arrCAMPOS(2, 4) = "5000"
                arrCAMPOS(2, 5) = "SGI_DESCRI"
                
                varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Capacidade da Lata")
                
                If Len(Trim(varRETORNO)) > 0 Then
                
                   .Cell(flexcpText, Row, conCOL_SonCap_CodLinha) = varRETORNO
                   Call LimpaCamposCorte(Row)
                   If PegaDescrLinha(varRETORNO, Row) = False Then
                        Call LimpaCamposGridLinhaProd(Row)
                        Exit Sub
                   End If
                   
                End If
            Case conCOL_SonCap_PesqCorte
                
                If Len(Trim(.Cell(flexcpText, Row, conCOL_SonCap_CodLinha))) = 0 Then
                    MsgBox "ATENÇÂO Informe primeiro a capacidade !!!", vbOKOnly + vbExclamation, "Aviso"
                    Exit Sub
                End If
                
                If cTipOper = "C" Then Exit Sub
                
                ReDim arrCAMPOS(1 To 3, 1 To 5) As String
                ReDim arrTABELA(1 To 1) As String
                
                
                sSql = ""
                
                sSql = "Select " & vbCrLf
                sSql = sSql & "       CODLIN.SGI_ITENCOR" & vbCrLf
                sSql = sSql & "      ,CODLIN.SGI_CODMEDCORT" & vbCrLf
                sSql = sSql & "      ,DIMCO.SGI_DESCORTE" & vbCrLf
                
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_MEDCORTELINHA CODLIN" & vbCrLf
                sSql = sSql & "      ,SGI_CADDIMCORTE   DIMCO " & vbCrLf

                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       CODLIN.SGI_FILIAL = " & FILIAL & vbCrLf
                sSql = sSql & "   And CODLIN.SGI_CODIGO = " & Trim(.Cell(flexcpText, Row, conCOL_SonCap_IDLinha)) & vbCrLf
                sSql = sSql & "   And DIMCO.SGI_FILIAL  = CODLIN.SGI_FILIAL" & vbCrLf
                sSql = sSql & "   And DIMCO.SGI_CODIGO  = CODLIN.SGI_CODMEDCORT" & vbCrLf
                
                arrTABELA(1) = sSql
                
                arrCAMPOS(1, 1) = "SGI_ITENCOR"
                arrCAMPOS(1, 2) = "N"
                arrCAMPOS(1, 3) = "Seg.Corte"
                arrCAMPOS(1, 4) = "1000"
                arrCAMPOS(1, 5) = "CODLIN.SGI_ITENCOR"
                
                arrCAMPOS(2, 1) = "SGI_CODMEDCORT"
                arrCAMPOS(2, 2) = "N"
                arrCAMPOS(2, 3) = "Cod.Corte"
                arrCAMPOS(2, 4) = "1000"
                arrCAMPOS(2, 5) = "CODLIN.SGI_CODMEDCORT"
                
                arrCAMPOS(3, 1) = "SGI_DESCORTE"
                arrCAMPOS(3, 2) = "S"
                arrCAMPOS(3, 3) = "Descrição"
                arrCAMPOS(3, 4) = "5000"
                arrCAMPOS(3, 5) = "DIMCO.SGI_DESCORTE"
                
                varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Cortes da Lata")
                
                If Len(Trim(varRETORNO)) > 0 Then
                    
                    If PegaDescrCorte(varRETORNO, .Cell(flexcpText, Row, conCOL_SonCap_IDLinha), Row) = False Then Exit Sub
                   
                   If objBLBFunc.FcVerifItensRepetidos(grdCAPPROD, Row, conCOL_SonCap_INDICE, .Cell(flexcpText, Row, conCOL_SonCap_INDICE)) = False Then
                        MsgBox "Esta Sequência de Corte já foi relacionado na Gride !!!", vbOKOnly + vbExclamation
                        Call LimpaCamposCorte(Row)
                        Exit Sub
                   End If
                    
                End If
        End Select
    End With
End Sub

Private Sub grdCAPPROD_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdCAPPROD
          Select Case Col
                    Case conCOL_SonCap_CodLinha
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
          End Select
     End With
End Sub

Private Sub grdCAPPROD_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
     With grdCAPPROD
          Select Case Col
                 Case conCOL_SonCap_CodLinha
                        If .EditText = Empty Then Exit Sub
                        If PegaDescrLinha(.EditText, Row) = False Then
                             Call LimpaCamposGridLinhaProd(Row)
                             Cancel = True
                             Exit Sub
                        End If
          End Select
     End With
End Sub

Private Sub grdParamManut_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    Dim strTOTALPERIODO As String
    Dim dtTotalLiquido  As Date
    
    With grdParamManut
        Select Case Col
            Case conCOL_SonParamManut_HoraInici, _
                 conCOL_SonParamManut_HoraFinal
                 If Len(Trim(Replace(.Cell(flexcpText, Row, conCOL_SonParamManut_HoraInici), ":", ""))) > 0 And _
                    Len(Trim(Replace(.Cell(flexcpText, Row, conCOL_SonParamManut_HoraFinal), ":", ""))) > 0 Then
                    
                    If CDate(.Cell(flexcpText, Row, conCOL_SonParamManut_HoraInici)) > CDate(.Cell(flexcpText, Row, conCOL_SonParamManut_HoraFinal)) Then
                       MsgBox "Hora de Inicial não pode ser maior que hora final !!!", vbOKOnly + vbExclamation, "Aviso"
                       .Cell(flexcpText, Row, conCOL_SonParamManut_HoraInici) = ""
                       .Cell(flexcpText, Row, conCOL_SonParamManut_HoraFinal) = ""
                       .Cell(flexcpText, Row, conCOL_SonParamManut_TempoUsado) = ""
                       Exit Sub
                    End If
                    
                    strTOTALPERIODO = objBLBFunc.CalcTempo(.Cell(flexcpText, Row, conCOL_SonParamManut_HoraInici), .Cell(flexcpText, Row, conCOL_SonParamManut_HoraFinal))
                    .Cell(flexcpText, Row, conCOL_SonParamManut_TempoUsado) = Format(CDate(strTOTALPERIODO), "HH:MM")
                    
                 End If
                    
                 If Col = conCOL_SonParamManut_HoraInici Then Call PopTotHoras(.Cell(flexcpText, Row, conCOL_SonParamManut_HoraInici), conCOL_SonDiasManut_HoraInic, .Cell(flexcpText, Row, conCOL_SonParamManut_Indice))
                 If Col = conCOL_SonParamManut_HoraFinal Then Call PopTotHoras(.Cell(flexcpText, Row, conCOL_SonParamManut_HoraFinal), conCOL_SonDiasManut_HoraFina, .Cell(flexcpText, Row, conCOL_SonParamManut_Indice))
            
            Case conCOL_SonParamManut_TipoManut, conCOL_SonParammanut_DataInici, conCOL_SonParammanut_DataFinal
                 
                 If VerifItensRepetidos(Row, conCOL_SonParamManut_Indice, Trim(.Cell(flexcpText, Row, conCOL_SonParamManut_TipoManut)) & Trim(.Cell(flexcpText, Row, conCOL_SonParammanut_DataInici))) = True Then
                    MsgBox "Este tipo de manutenção já existe !!!", vbOKOnly + vbExclamation, "Aviso"
                    .Cell(flexcpText, Row, conCOL_SonParamManut_TipoManut) = Empty
                    .Cell(flexcpText, Row, conCOL_SonParamManut_Indice) = Empty
                    Exit Sub
                 End If
                 
                 If Col = conCOL_SonParammanut_DataInici Then
                    If .Cell(flexcpText, Row, conCOL_SonParamManut_TipoManut) <> 9 Then .Cell(flexcpText, Row, conCOL_SonParammanut_DataFinal) = .Cell(flexcpText, Row, conCOL_SonParammanut_DataInici)
                 End If
                 
                 .Cell(flexcpText, Row, conCOL_SonParamManut_Indice) = Trim(.Cell(flexcpText, Row, conCOL_SonParamManut_TipoManut)) & Trim(.Cell(flexcpText, Row, conCOL_SonParammanut_DataInici))
                 If .Cell(flexcpText, Row, conCOL_SonParamManut_TipoManut) <> Empty Then Call PopGrdDias(Row, CInt(.Cell(flexcpText, Row, conCOL_SonParamManut_TipoManut)), Trim(.Cell(flexcpText, Row, conCOL_SonParamManut_Indice)))
            
            Case conCOL_SonParamManut_Ativo
                 ''Call MudaAtvo(Row)
                 Call PopSN(.Cell(flexcpText, Row, conCOL_SonParamManut_Ativo), conCOL_SonDiasManut_Ativo, .Cell(flexcpText, Row, conCOL_SonParamManut_Indice))
            
            Case conCOL_SonParamManut_MaqUso
                 Call PopSN(.Cell(flexcpText, Row, conCOL_SonParamManut_MaqUso), conCOL_SonDiasManut_MaqUso, .Cell(flexcpText, Row, conCOL_SonParamManut_Indice))
            
            Case conCOL_SonParamManut_EmConjunto
                 Call PopSN(.Cell(flexcpText, Row, conCOL_SonParamManut_EmConjunto), conCOL_SonDiasManut_EmConjunto, .Cell(flexcpText, Row, conCOL_SonParamManut_Indice))
        
        End Select
    End With

End Sub

Private Sub grdParamManut_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonParamManut_TempoUsado
         Cancel = True
    Case conCOL_SonParamManut_TipoManut, _
         conCOL_SonParamManut_HoraInici, _
         conCOL_SonParamManut_HoraFinal, _
         conCOL_SonParamManut_MaqUso, _
         conCOL_SonParamManut_Ativo
         If cTipOper = "C" Then Cancel = True
    Case conCOL_SonParammanut_DataFinal
         If grdParamManut.Cell(flexcpText, Row, conCOL_SonParamManut_TipoManut) <> 9 Then Cancel = True
    End Select
    Exit Sub
End Sub

Private Sub grdParamManut_Click()
    If (grdParamManut.Rows - 1) > 0 And grdParamManut.Row > 0 Then Call PosRegGrdDias(Trim(grdParamManut.Cell(flexcpText, grdParamManut.Row, conCOL_SonParamManut_TipoManut)) & Trim(grdParamManut.Cell(flexcpText, grdParamManut.Row, conCOL_SonParammanut_DataInici)))
End Sub

Private Sub grdParamManut_RowColChange()
    If (grdParamManut.Rows - 1) > 0 And grdParamManut.Row > 0 Then Call PosRegGrdDias(Trim(grdParamManut.Cell(flexcpText, grdParamManut.Row, conCOL_SonParamManut_TipoManut)) & Trim(grdParamManut.Cell(flexcpText, grdParamManut.Row, conCOL_SonParammanut_DataInici)))
End Sub

Private Sub grdParamManut_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    Dim intHoras As Integer
    Dim intMinutos As Integer
    
    With grdParamManut
        Select Case Col
               Case conCOL_SonParamManut_HoraInici, _
                    conCOL_SonParamManut_HoraFinal
                    If .EditText = "  :  " Then
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
               Case conCOL_SonParamManut_TipoManut, _
                    conCOL_SonParammanut_DataInici, _
                    conCOL_SonParammanut_DataFinal
                    Call DeleteDias(Trim(.Cell(flexcpText, Row, conCOL_SonParamManut_TipoManut)) & Trim(.Cell(flexcpText, Row, conCOL_SonParammanut_DataInici)))
        End Select
    End With

End Sub

Private Sub txtCADGRLIN_GotFocus()
    objBLBFunc.SelecionaCampos txtCADGRLIN.Name, Me
End Sub

Private Sub txtCADGRLIN_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCADGRLIN.Text
End Sub

Private Sub txtCADGRLIN_Validate(Cancel As Boolean)

    Dim i As Integer
    
    If Len(Trim(txtCADGRLIN.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCADGRLIN.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCADGRLIN.Text = ""
       Cancel = True
       Exit Sub
    End If
        
End Sub

Private Sub txtCAPMIN_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCAPMIN.Text
End Sub

Private Sub txtCAPMIN_Validate(Cancel As Boolean)
    
    If Len(Trim(txtCAPMIN.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCAPMIN.Text) Then
       MsgBox "Somente é permitido numeros e pontos !!!", vbOKOnly + vbExclamation, "Aviso"
       txtCAPMIN.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtCAPMIN.Text = Format(txtCAPMIN.Text, "#,####0.0000")
    
End Sub

Private Sub txtCODFAMILIA_GotFocus()
    objBLBFunc.SelecionaCampos txtCODFAMILIA.Name, frmCADMAQUINA
End Sub

Private Sub txtCODFAMILIA_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODFAMILIA.Text
End Sub

Private Sub txtCODFAMILIA_Validate(Cancel As Boolean)

    Dim i As Integer
    
    If Len(Trim(txtCODFAMILIA.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODFAMILIA.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODFAMILIA.Text = ""
       Cancel = True
       Exit Sub
    End If
        
    cboFAMILIA.ListIndex = -1
    For i = 0 To (cboFAMILIA.ListCount - 1)
        If cboFAMILIA.ItemData(i) = Str(Val(txtCODFAMILIA.Text)) Then cboFAMILIA.ListIndex = i
    Next i
    
    If cboFAMILIA.ListIndex = -1 Then
       MsgBox "Esta familia de máquina não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODFAMILIA.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub


Private Sub txtCodTurn_GotFocus()
    objBLBFunc.SelecionaCampos txtCodTurn.Name, frmCADMAQUINA
End Sub

Private Sub txtCodTurn_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCodTurn.Text
End Sub

Private Sub txtCodTurn_Validate(Cancel As Boolean)

    Dim i As Integer
    
    If Len(Trim(txtCodTurn.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCodTurn.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCodTurn.Text = ""
       Cancel = True
       Exit Sub
    End If
   
    cboCodTurn.ListIndex = -1
    For i = 0 To (cboCodTurn.ListCount - 1)
        If cboCodTurn.ItemData(i) = Str(Val(txtCodTurn.Text)) Then cboCodTurn.ListIndex = i
    Next i
    
    If cboCodTurn.ListIndex = -1 Then
       MsgBox "Este turno não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCodTurn.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub





Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtFontAlim_GotFocus()
    objBLBFunc.SelecionaCampos txtFontAlim.Name, frmCADMAQUINA
End Sub

Private Sub txtFontAlim_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtFontAlim.Text
End Sub

Private Sub txtFontAlim_Validate(Cancel As Boolean)

    Dim i As Integer
    
    If Len(Trim(txtFontAlim.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtFontAlim.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtFontAlim.Text = ""
       Cancel = True
       Exit Sub
    End If
        
    cboFontAlim.ListIndex = -1
    For i = 0 To (cboFontAlim.ListCount - 1)
        If cboFontAlim.ItemData(i) = Str(Val(txtFontAlim.Text)) Then cboFontAlim.ListIndex = i
    Next i
    
    If cboFontAlim.ListIndex = -1 Then
       MsgBox "Este tipo de alimentação não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtFontAlim.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub Inc_FontALim()

    Dim i As Integer
    
    If (Len(Trim(txtFontAlim.Text)) = 0) Or (cboFontAlim.ListIndex = -1) Then
       MsgBox "Informe o código da fonte de alimentação !!!", vbOKOnly + vbExclamation, "Aviso"
       txtFontAlim.SetFocus
       Exit Sub
    End If
    
    For i = 1 To (flxFONTEALIM.Rows - 1)
        If Trim(flxFONTEALIM.TextMatrix(i, 1)) = Trim(txtFontAlim.Text) Then
           MsgBox "Este tipo de alimentação já existe !!!", vbOKOnly + vbExclamation, "Aviso"
           txtFontAlim.Text = ""
           cboFontAlim.ListIndex = -1
           txtFontAlim.SetFocus
           Exit Sub
        End If
    Next i
   
    flxFONTEALIM.AddItem vbTab & _
                         txtFontAlim.Text & vbTab & _
                         cboFontAlim.Text
                         
    txtFontAlim.Text = ""
    cboFontAlim.ListIndex = -1
    txtFontAlim.SetFocus
    
End Sub

Private Sub IncGridTurn()

    Dim i As Integer
    
    If (Len(Trim(txtCodTurn.Text)) = 0) Or (cboCodTurn.ListIndex = -1) Then
       MsgBox "Informe o código do turno !!!", vbOKOnly + vbExclamation, "Aviso"
       txtCodTurn.SetFocus
       Exit Sub
    End If
    
    If (Len(Trim(txtPORCUSO.Text)) = 0) Then
       MsgBox "Informe a % de Uso !!!", vbOKOnly + vbExclamation, "Aviso"
       txtPORCUSO.SetFocus
       Exit Sub
    End If
    
    For i = 1 To (flxGRIDTURNOS.Rows - 1)
        If Trim(flxGRIDTURNOS.TextMatrix(i, 1)) = Trim(txtCodTurn.Text) Then
           MsgBox "Este turno já existe !!!", vbOKOnly + vbExclamation, "Aviso"
           txtCodTurn.Text = ""
           cboCodTurn.ListIndex = -1
           txtCodTurn.SetFocus
           Exit Sub
        End If
    Next i
    
    flxGRIDTURNOS.AddItem vbTab & _
                         txtCodTurn.Text & vbTab & _
                         cboCodTurn.Text & vbTab & _
                         Format(txtPORCUSO.Text, "#,##0.00")
                         
    txtCodTurn.Text = ""
    cboCodTurn.ListIndex = -1
    txtPORCUSO.Text = ""
    txtCodTurn.SetFocus

End Sub


Private Sub ConfGridPecas()

    flxPECAS.Rows = 1
    flxPECAS.Cols = 3
    
    flxPECAS.TextMatrix(0, 0) = ""
    flxPECAS.TextMatrix(0, 1) = "Código"
    flxPECAS.TextMatrix(0, 2) = "Descrição"
    
    flxPECAS.ColWidth(0) = 0
    flxPECAS.ColWidth(1) = 1000
    flxPECAS.ColWidth(2) = 3000
    
End Sub


Private Sub CofGridOperadores()

    flxOPERADORES.Rows = 1
    flxOPERADORES.Cols = 4
    
    flxOPERADORES.TextMatrix(0, 0) = ""
    flxOPERADORES.TextMatrix(0, 1) = "Código"
    flxOPERADORES.TextMatrix(0, 2) = "Descrição"
    flxOPERADORES.TextMatrix(0, 3) = "% Utiliza"
    
    flxOPERADORES.ColWidth(0) = 0
    flxOPERADORES.ColWidth(1) = 1000
    flxOPERADORES.ColWidth(2) = 3000
    flxOPERADORES.ColWidth(3) = 1500
    
End Sub

Private Sub txtOperadores_GotFocus()
    objBLBFunc.SelecionaCampos txtOperadores.Name, frmCADMAQUINA
End Sub

Private Sub txtOperadores_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtOperadores.Text
End Sub

Private Sub txtOperadores_Validate(Cancel As Boolean)

    Dim i As Integer
    
    If Len(Trim(txtOperadores.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtOperadores.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtOperadores.Text = ""
       Cancel = True
       Exit Sub
    End If
        
    cboOperadores.ListIndex = -1
    For i = 0 To (cboOperadores.ListCount - 1)
        If cboOperadores.ItemData(i) = Str(Val(txtOperadores.Text)) Then cboOperadores.ListIndex = i
    Next i
    
    If cboOperadores.ListIndex = -1 Then
       MsgBox "Operadore inválido !!!", vbOKOnly + vbCritical, "Aviso"
       txtOperadores.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Public Sub Inc_Operadores()

    Dim i As Integer
    
    If (Len(Trim(txtOperadores.Text)) = 0) Or (cboOperadores.ListIndex = -1) Then
       MsgBox "Informe o código do operador !!!", vbOKOnly + vbExclamation, "Aviso"
       txtOperadores.SetFocus
       Exit Sub
    End If
    If (Len(Trim(txtPORUTIL.Text)) = 0) Then
       MsgBox "Informe a % de Utilização !!!", vbOKOnly + vbExclamation, "Aviso"
       txtPORUTIL.SetFocus
       Exit Sub
    End If
        
    For i = 1 To (flxOPERADORES.Rows - 1)
        If Trim(flxOPERADORES.TextMatrix(i, 1)) = Trim(txtOperadores.Text) Then
           MsgBox "Este tipo de alimentação já existe !!!", vbOKOnly + vbExclamation, "Aviso"
           txtOperadores.Text = ""
           cboOperadores.ListIndex = -1
           txtOperadores.SetFocus
           Exit Sub
        End If
    Next i
   
    flxOPERADORES.AddItem vbTab & _
                          txtOperadores.Text & vbTab & _
                          cboOperadores.Text & vbTab & _
                          txtPORUTIL.Text
                         
    txtOperadores.Text = ""
    cboOperadores.ListIndex = -1
    txtPORUTIL.Text = ""
    txtOperadores.SetFocus

End Sub

Private Function ValidaCampos() As Boolean

     Dim i As Integer
     
     ValidaCampos = False
     
     If Trim(Len(txtDescricao.Text)) = 0 Then
        MsgBox "descrição da máquina inválida !!!", vbOKOnly + vbCritical, "Aviso"
        txtDescricao.SetFocus
        Exit Function
     End If
     If Len(Trim(txtCAPMIN.Text)) = 0 Then
        MsgBox "Informe a capacidade de produção por minuto !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCAPMIN.SetFocus
        Exit Function
     End If
     If cboUnidade.ListIndex = -1 Then
        MsgBox "Informe a undade de produção !!!", vbOKOnly + vbExclamation, "Aviso"
        cboUnidade.SetFocus
        Exit Function
     End If
     If cboFAMILIA.ListIndex = -1 Then
        MsgBox "Informe a familia de máquina !!!", vbOKOnly + vbExclamation, "Aviso"
        cboFAMILIA.SetFocus
        Exit Function
     End If
     
     For i = 1 To (grdParamManut.Rows - 1)
         If grdParamManut.Cell(flexcpText, i, conCOL_SonParamManut_TipoManut) = Empty Then
            MsgBox "Informe o Tipo de Mautenção !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Function
         End If
         If grdParamManut.Cell(flexcpText, i, conCOL_SonParamManut_HoraInici) = Empty Then
            MsgBox "Informe a Hora Inicial !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Function
         End If
         If grdParamManut.Cell(flexcpText, i, conCOL_SonParamManut_HoraFinal) = Empty Then
            MsgBox "Informe a Hora Final !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Function
         End If
     Next i
     
     ValidaCampos = True
     
End Function


Private Sub Consulta()

    Dim i As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
    Frame10.Enabled = False
    Frame11.Enabled = False
    Frame3.Enabled = False
    Frame15.Enabled = False
    
    StMaquinas.Tab = 0
    StPecas.Tab = 0
   
    Me.Caption = "Cadastro de Máquinas - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADMAQUINA
    objCADMAQUINA.CODIGO = iCodigo
    
    ConfGridTipAlim
    ConfGridTurnos
    ConfGridPecas
    CofGridOperadores
    InitGridParametros
    InitGridDiasanutencao
    Call InitGridCapCorte
    Call LimpaCamposLabel
    
    optEmpresa(0).Value = True
    
    objCADMAQUINA.PreencheComboTipAlim cboFontAlim
    objCADMAQUINA.PreencheComboTurnos cboCodTurn
    objCADMAQUINA.PreenchComboUnidade cboUnidade
    objCADMAQUINA.PreencheComboOperadores cboOperadores
    objCADMAQUINA.PreenchComboFamilia cboFAMILIA
    
    If objCADMAQUINA.Carrega_campos = True Then
    
       lblCodigo.Caption = Str(objCADMAQUINA.CODIGO)
       txtDescricao.Text = objCADMAQUINA.DESCRI
       optAtivoSN.Item(objCADMAQUINA.OPTATIVO).Value = True
       txtAnoFabr.Text = objCADMAQUINA.ANOFABR
       txtNumAtivo.Text = objCADMAQUINA.NUMATIVO
       txtCAPMIN.Text = Format(objCADMAQUINA.CAPMINUTO, "#,####0.0000")
       txtQtdMaquinas.Text = objCADMAQUINA.QtdOperadores
       txtFraOpeadores.Text = Format(objCADMAQUINA.FraOperadores, "#,###0.000")
       
       For i = 0 To (cboUnidade.ListCount - 1)
           If cboUnidade.ItemData(i) = objCADMAQUINA.UNIMED Then cboUnidade.ListIndex = i
       Next i
       
       txtCODFAMILIA.Text = objCADMAQUINA.FAMILIA
       For i = 0 To (cboFAMILIA.ListCount - 1)
           If cboFAMILIA.ItemData(i) = objCADMAQUINA.FAMILIA Then cboFAMILIA.ListIndex = i
       Next i
       
       arrGRDFONTALIM = objCADMAQUINA.GRDFONTALIM
       
       '' Fonte de Alimentação
       If IsArray(arrGRDFONTALIM) Then
          For i = 1 To UBound(arrGRDFONTALIM)
          
              sSql = "Select " & vbCrLf
              sSql = sSql & "       *" & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADTIPALIM " & vbCrLf
              sSql = sSql & " Where " & vbCrLf
              sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
              sSql = sSql & "   And SGI_CODIGO = " & arrGRDFONTALIM(i)
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
              
              Do While Not BREC.EOF
                 flxFONTEALIM.AddItem "" & vbTab & _
                                      BREC!SGI_CODIGO & vbTab & _
                                      BREC!SGI_DESCRI
                 BREC.MoveNext
              Loop
              BREC.Close
              
          Next i
       End If
       
       '' Turnos
       arrGRDTURNOS = objCADMAQUINA.GRDTURNOS
       
       If IsArray(arrGRDTURNOS) Then
          For i = 1 To UBound(arrGRDTURNOS)
          
              sSql = "Select " & vbCrLf
              sSql = sSql & "       *" & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADQTDETURN " & vbCrLf
              sSql = sSql & " Where " & vbCrLf
              sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
              sSql = sSql & "   And SGI_CODIGO = " & arrGRDTURNOS(i, 1)
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
              
              Do While Not BREC.EOF
                 flxGRIDTURNOS.AddItem "" & vbTab & _
                                       BREC!SGI_CODIGO & vbTab & _
                                       BREC!SGI_DESCRI & vbTab & _
                                       Format(arrGRDTURNOS(i, 2), "#,##0.00")
                 BREC.MoveNext
              Loop
              BREC.Close
              
          Next i
       End If
       
       '' Operadores
       arrGRDOPERADOR = objCADMAQUINA.GRDOPERADOR
       
       If IsArray(arrGRDOPERADOR) Then
          For i = 1 To UBound(arrGRDOPERADOR)
          
              sSql = "Select " & vbCrLf
              sSql = sSql & "       *" & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADOPERADOR " & vbCrLf
              sSql = sSql & " Where " & vbCrLf
              sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
              sSql = sSql & "   And SGI_CODIGO = '" & arrGRDOPERADOR(i, 1) & "'"
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
              
              Do While Not BREC.EOF
                 flxOPERADORES.AddItem "" & vbTab & _
                                       BREC!SGI_CODIGO & vbTab & _
                                       BREC!SGI_DESCRI & vbTab & _
                                       Format(arrGRDOPERADOR(i, 2), "#,##0.00")
                 BREC.MoveNext
              Loop
              BREC.Close
              
          Next i
       End If
    
       Call PopGrdParamDias
       Call PopGridLinha
    
    End If

End Sub

Private Sub Altera()

    Dim i As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame10.Enabled = True
    Frame11.Enabled = True
    Frame3.Enabled = True
    Frame15.Enabled = True
    
    StMaquinas.Tab = 0
    StPecas.Tab = 0
   
    Me.Caption = "Cadastro de Máquinas - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADMAQUINA
    objCADMAQUINA.CODIGO = iCodigo
    
    ConfGridTipAlim
    ConfGridTurnos
    ConfGridPecas
    CofGridOperadores
    InitGridParametros
    InitGridDiasanutencao
    Call InitGridCapCorte
    Call LimpaCamposLabel
    
    objCADMAQUINA.PreencheComboTipAlim cboFontAlim
    objCADMAQUINA.PreencheComboTurnos cboCodTurn
    objCADMAQUINA.PreenchComboUnidade cboUnidade
    objCADMAQUINA.PreencheComboOperadores cboOperadores
    objCADMAQUINA.PreenchComboFamilia cboFAMILIA
    
    optEmpresa(0).Value = True
    
    If objCADMAQUINA.Carrega_campos = True Then
    
       lblCodigo.Caption = Str(objCADMAQUINA.CODIGO)
       txtDescricao.Text = objCADMAQUINA.DESCRI
       optAtivoSN.Item(objCADMAQUINA.OPTATIVO).Value = True
       txtAnoFabr.Text = objCADMAQUINA.ANOFABR
       txtNumAtivo.Text = objCADMAQUINA.NUMATIVO
       txtCAPMIN.Text = Format(objCADMAQUINA.CAPMINUTO, "#,####0.0000")
       txtQtdMaquinas.Text = objCADMAQUINA.QtdOperadores
       txtFraOpeadores.Text = Format(objCADMAQUINA.FraOperadores, "#,###0.000")
       
       For i = 0 To (cboUnidade.ListCount - 1)
           If cboUnidade.ItemData(i) = objCADMAQUINA.UNIMED Then cboUnidade.ListIndex = i
       Next i
       
       txtCODFAMILIA.Text = objCADMAQUINA.FAMILIA
       For i = 0 To (cboFAMILIA.ListCount - 1)
           If cboFAMILIA.ItemData(i) = objCADMAQUINA.FAMILIA Then cboFAMILIA.ListIndex = i
       Next i
       
       arrGRDFONTALIM = objCADMAQUINA.GRDFONTALIM
       
       '' Fonte de Alimentação
       If IsArray(arrGRDFONTALIM) Then
          For i = 1 To UBound(arrGRDFONTALIM)
          
              sSql = "Select " & vbCrLf
              sSql = sSql & "       *" & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADTIPALIM " & vbCrLf
              sSql = sSql & " Where " & vbCrLf
              sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
              sSql = sSql & "   And SGI_CODIGO = " & arrGRDFONTALIM(i)
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
              
              Do While Not BREC.EOF
                 flxFONTEALIM.AddItem "" & vbTab & _
                                      BREC!SGI_CODIGO & vbTab & _
                                      BREC!SGI_DESCRI & vbTab

                 BREC.MoveNext
              Loop
              BREC.Close
              
          Next i
       End If
       
       '' Turnos
       arrGRDTURNOS = objCADMAQUINA.GRDTURNOS
       
       If IsArray(arrGRDTURNOS) Then
          For i = 1 To UBound(arrGRDTURNOS)
          
              sSql = "Select " & vbCrLf
              sSql = sSql & "       *" & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADQTDETURN " & vbCrLf
              sSql = sSql & " Where " & vbCrLf
              sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
              sSql = sSql & "   And SGI_CODIGO = " & arrGRDTURNOS(i, 1)
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
              
              Do While Not BREC.EOF
                 flxGRIDTURNOS.AddItem "" & vbTab & _
                                       BREC!SGI_CODIGO & vbTab & _
                                       BREC!SGI_DESCRI & vbTab & _
                                       Format(arrGRDTURNOS(i, 2), "#,##0.00")
                 BREC.MoveNext
              Loop
              BREC.Close
              
          Next i
       End If
       
       '' Operadores
       arrGRDOPERADOR = objCADMAQUINA.GRDOPERADOR
       
       If IsArray(arrGRDOPERADOR) Then
          For i = 1 To UBound(arrGRDOPERADOR)
          
              sSql = "Select " & vbCrLf
              sSql = sSql & "       *" & vbCrLf
              sSql = sSql & "  From " & vbCrLf
              sSql = sSql & "       SGI_CADOPERADOR " & vbCrLf
              sSql = sSql & " Where " & vbCrLf
              sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
              sSql = sSql & "   And SGI_CODIGO = '" & arrGRDOPERADOR(i, 1) & "'"
              
              BREC.Open sSql, adoBanco_Dados, adOpenDynamic
              
              Do While Not BREC.EOF
                 flxOPERADORES.AddItem "" & vbTab & _
                                       BREC!SGI_CODIGO & vbTab & _
                                       BREC!SGI_DESCRI & vbTab & _
                                       Format(arrGRDOPERADOR(i, 2), "#,##0.00")
                 BREC.MoveNext
              Loop
              BREC.Close
              
          Next i
       End If
    
       Call PopGrdParamDias
 
    End If

End Sub

Private Sub txtPORCUSO_GotFocus()
    objBLBFunc.SelecionaCampos txtPORCUSO.Name, frmCADMAQUINA
End Sub

Private Sub txtPORCUSO_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtPORCUSO.Text
End Sub

Private Sub txtPORCUSO_Validate(Cancel As Boolean)
    
    If Len(Trim(txtPORCUSO.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtPORCUSO.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbExclamation, "Aviso"
       txtPORCUSO.Text = ""
       txtPORCUSO.SetFocus
       Exit Sub
    End If
    
    txtPORCUSO.Text = Format(txtPORCUSO.Text, "#,##0.00")
    
End Sub

Private Sub txtPORUTIL_GotFocus()
    objBLBFunc.SelecionaCampos txtPORUTIL.Name, frmCADMAQUINA
End Sub

Private Sub txtPORUTIL_Validate(Cancel As Boolean)

    If Len(Trim(txtPORUTIL.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtPORUTIL.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbExclamation, "Aviso"
       txtPORUTIL.Text = ""
       txtPORUTIL.SetFocus
       Exit Sub
    End If
    
    txtPORUTIL.Text = Format(txtPORUTIL.Text, "#,##0.00")

End Sub

Private Sub txtQtdMaquinas_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtQtdMaquinas.Text
End Sub

Private Sub txtQtdMaquinas_Validate(Cancel As Boolean)

    Dim curFraOperadores As Currency
    
    If Len(Trim(txtQtdMaquinas.Text)) = 0 Then
       txtFraOpeadores.Text = ""
       Exit Sub
    End If
    
    If Not IsNumeric(txtQtdMaquinas.Text) Then
       MsgBox "Somente é permitido números e pontos !!!", vbOKOnly + vbExclamation, "Aviso"
       txtQtdMaquinas.Text = ""
       txtFraOpeadores.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    '' Fração de Operadores
    If Len(Trim(txtQtdMaquinas.Text)) > 0 And txtQtdMaquinas.Text <> "0" Then
       curFraOperadores = (1 / CCur(txtQtdMaquinas.Text))
       txtFraOpeadores.Text = Format(curFraOperadores, "#,###0.000")
    End If
End Sub


Private Sub CalcPcMin()
    
    Dim curPecaHora As Currency
    Dim curQtdPcMin As Currency
    Dim curCutoHora As Currency
    Dim curCustoMin As Currency
        
    curPecaHora = 0
    curQtdPcMin = 0
    curCutoHora = 0
    curCustoMin = 0
    
    ''If Len(Trim(txtQtdPHora.Text)) > 0 Then curPecaHora = CCur(txtQtdPHora.Text)
    ''If Len(Trim(txtCustHora.Text)) > 0 Then curCutoHora = CCur(txtCustHora.Text)
    
    If curPecaHora > 0 Then curQtdPcMin = (curPecaHora / 60)
    If curCutoHora > 0 Then curCustoMin = (curCutoHora / 60)
    
    ''If curQtdPcMin > 0 Then txtQtdPcMin.Text = Format(curQtdPcMin, "##00")
    ''If curCustoMin > 0 Then txtCustMin.Text = Format(curCustoMin, "#,##0.00")
    
End Sub

Private Sub InitGridParametros()

    With grdParamManut
    
       .Cols = conColumnsIn_SonParamManut
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonParamManut_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonParamManut_TipoManut) = ""
       .ColDataType(conCOL_SonParamManut_TipoManut) = flexDTString
       .ColComboList(conCOL_SonParamManut_TipoManut) = "|#1;Por Turno|#2;Por Dia|#3;Por Semana|#4;Por Quinzena|#5;Por Mês|#6;Por Trimestre|#7;Por Semestre|#8;Por Ano|#9;Customizado"
       
       .Cell(flexcpData, 0, conCOL_SonParammanut_DataInici) = ""
       .ColDataType(conCOL_SonParammanut_DataInici) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonParammanut_DataFinal) = ""
       .ColDataType(conCOL_SonParammanut_DataFinal) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonParamManut_HoraInici) = ""
       .ColDataType(conCOL_SonParamManut_HoraInici) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonParamManut_HoraFinal) = ""
       .ColDataType(conCOL_SonParamManut_HoraFinal) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonParamManut_TempoUsado) = ""
       .ColDataType(conCOL_SonParamManut_TempoUsado) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonParamManut_MaqUso) = ""
       .ColDataType(conCOL_SonParamManut_MaqUso) = flexDTBoolean
       .ColFormat(conCOL_SonParamManut_MaqUso) = "Sim;Não"
         
       .Cell(flexcpData, 0, conCOL_SonParamManut_Ativo) = ""
       .ColDataType(conCOL_SonParamManut_Ativo) = flexDTBoolean
       .ColFormat(conCOL_SonParamManut_Ativo) = "Sim;Não"
       
       .Cell(flexcpData, 0, conCOL_SonParamManut_EmConjunto) = ""
       .ColDataType(conCOL_SonParamManut_EmConjunto) = flexDTBoolean
       .ColFormat(conCOL_SonParamManut_EmConjunto) = "Sim;Não"
       
       .Cell(flexcpData, 0, conCOL_SonParamManut_Indice) = ""
       .ColDataType(conCOL_SonParamManut_Indice) = flexDTString
       
       .ColWidth(conCOL_SonParamManut_TipoManut) = 2000
       .ColWidth(conCOL_SonParammanut_DataInici) = 1000
       .ColWidth(conCOL_SonParammanut_DataFinal) = 1000
       .ColWidth(conCOL_SonParamManut_HoraInici) = 1000
       .ColWidth(conCOL_SonParamManut_HoraFinal) = 1000
       .ColWidth(conCOL_SonParamManut_TempoUsado) = 1150
       .ColWidth(conCOL_SonParamManut_MaqUso) = 1000
       .ColWidth(conCOL_SonParamManut_Ativo) = 1000
       .ColWidth(conCOL_SonParamManut_EmConjunto) = 1000
       
       .ColHidden(conCOL_SonParamManut_Indice) = True
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub InitGridDiasanutencao()

    With grdAgendaManut
    
       .Cols = conColumnsIn_SonDiasManut
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonDiasManut_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonDiasManut_DataManut) = ""
       .ColDataType(conCOL_SonDiasManut_DataManut) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonDiasManut_HoraInic) = ""
       .ColDataType(conCOL_SonDiasManut_HoraInic) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonDiasManut_HoraFina) = ""
       .ColDataType(conCOL_SonDiasManut_HoraFina) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonDiasManut_Tempo) = ""
       .ColDataType(conCOL_SonDiasManut_Tempo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonDiasManut_MaqUso) = ""
       .ColDataType(conCOL_SonDiasManut_MaqUso) = flexDTBoolean
       .ColFormat(conCOL_SonDiasManut_MaqUso) = "Sim;Não"
         
       .Cell(flexcpData, 0, conCOL_SonDiasManut_Ativo) = ""
       .ColDataType(conCOL_SonDiasManut_Ativo) = flexDTBoolean
       .ColFormat(conCOL_SonDiasManut_Ativo) = "Sim;Não"
       
       .Cell(flexcpData, 0, conCOL_SonDiasManut_EmConjunto) = ""
       .ColDataType(conCOL_SonDiasManut_EmConjunto) = flexDTBoolean
       .ColFormat(conCOL_SonDiasManut_EmConjunto) = "Sim;Não"
       
       .Cell(flexcpData, 0, conCOL_SonDiasManut_Pai) = ""
       .ColDataType(conCOL_SonDiasManut_Pai) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonDiasManut_Indice) = ""
       .ColDataType(conCOL_SonDiasManut_Indice) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonDiasManut_DtParametro) = ""
       .ColDataType(conCOL_SonDiasManut_DtParametro) = flexDTDate
       
       
       .ColWidth(conCOL_SonDiasManut_DataManut) = 1500
       .ColWidth(conCOL_SonDiasManut_HoraInic) = 1000
       .ColWidth(conCOL_SonDiasManut_HoraFina) = 1000
       .ColWidth(conCOL_SonDiasManut_Tempo) = 1150
       .ColWidth(conCOL_SonDiasManut_MaqUso) = 1000
       .ColWidth(conCOL_SonDiasManut_Ativo) = 1000
       .ColWidth(conCOL_SonDiasManut_EmConjunto) = 1000
       
       .ColHidden(conCOL_SonDiasManut_Pai) = True
       .ColHidden(conCOL_SonDiasManut_Indice) = True
       .ColHidden(conCOL_SonDiasManut_DtParametro) = True
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub IncGrdParam()

    If ExisteLinhaVaziaParametros = False Then Exit Sub
    
    With grdParamManut
        .AddItem "" & vbTab & _
                 Format(Date, "DD/MM/YYYY") & vbTab & _
                 Format(Date, "DD/MM/YYYY") & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 ""
                 
        '' Formatado para Campo Hora
        .ColEditMask(conCOL_SonParammanut_DataInici) = "##/##/####"
        .ColEditMask(conCOL_SonParammanut_DataFinal) = "##/##/####"
        .ColEditMask(conCOL_SonParamManut_HoraInici) = "##:##"
        .ColEditMask(conCOL_SonParamManut_HoraFinal) = "##:##"
        .ColEditMask(conCOL_SonParamManut_TempoUsado) = "##:##"
                          
        '' Alinhamento dos Campos
        .ColAlignment(conCOL_SonParamManut_HoraInici) = flexAlignRightCenter
        .ColAlignment(conCOL_SonParamManut_HoraFinal) = flexAlignRightCenter
        .ColAlignment(conCOL_SonParamManut_TempoUsado) = flexAlignRightCenter
        
        Call PosRegGrdDias(grdParamManut.Cell(flexcpText, (grdParamManut.Rows - 1), conCOL_SonParamManut_TipoManut))
        
    End With
                          
End Sub

Private Function ExisteLinhaVaziaParametros() As Boolean
    ExisteLinhaVaziaParametros = False
    
    Dim i As Integer
    
    For i = 1 To (grdParamManut.Rows - 1)
        If grdParamManut.Cell(flexcpText, i, conCOL_SonParamManut_TipoManut) = Empty Then Exit Function
    Next i
    
    ExisteLinhaVaziaParametros = True
End Function



Private Function ExisteLinhaVaziaDiasManut() As Boolean
    ExisteLinhaVaziaDiasManut = False
    
    Dim i As Integer
    
    For i = 1 To (grdAgendaManut.Rows - 1)
        If grdAgendaManut.Cell(flexcpText, i, conCOL_SonDiasManut_DataManut) = Empty Then Exit Function
    Next i
    
    ExisteLinhaVaziaDiasManut = True
End Function

Private Sub PosRegGrdDias(strVALOR As String)
    Dim i As Integer
    With grdAgendaManut
         For i = 1 To (.Rows - 1)
             If Trim(.Cell(flexcpText, i, conCOL_SonDiasManut_Pai)) <> Trim(strVALOR) Then
                .RowHidden(i) = True
             Else
                .RowHidden(i) = False
             End If
         Next i
    End With
End Sub

Private Sub PopGrdDias(lngROW As Long, intTipoMAnut As Integer, strINDICE As String)

    '' - 1/Por Turno
    '' - 2/Por Dia
    '' - 3/Por Semana
    '' - 4/Por Quinzena
    '' - 5/Por Mês
    '' - 6/Por Trimeste
    '' - 7/Por Semestre
    '' - 8/Por Ano
    '' - 9/Customizado

    Dim intTotDias As Integer
    Dim strHORAINI As String
    Dim strHORAFIN As String
    Dim strTOTALUS As String
    Dim i          As Integer
    Dim dtManut    As Date
    Dim dtManutFin As Date
    Dim strATIVO   As String
    Dim strMAQUSO  As String
    Dim strEMCONJ  As String
    
    ''intTotDias = CDate("31/12/" & Year(Date)) - CDate(grdParamManut.Cell(flexcpText, lngRow, conCOL_SonParammanut_DataInici))
    intTotDias = (365 * 3)
    strHORAINI = grdParamManut.Cell(flexcpText, lngROW, conCOL_SonParamManut_HoraInici)
    strHORAFIN = grdParamManut.Cell(flexcpText, lngROW, conCOL_SonParamManut_HoraFinal)
    If Len(Trim(strHORAINI)) > 0 And Len(Trim(strHORAFIN)) > 0 Then strTOTALUS = objBLBFunc.CalcTempo(strHORAINI, strHORAFIN)
    dtManut = CDate(grdParamManut.Cell(flexcpText, lngROW, conCOL_SonParammanut_DataInici))
    dtManutFin = CDate(grdParamManut.Cell(flexcpText, lngROW, conCOL_SonParammanut_DataFinal))
    
    strMAQUSO = Trim(grdParamManut.Cell(flexcpText, lngROW, conCOL_SonParamManut_MaqUso))
    strATIVO = Trim(grdParamManut.Cell(flexcpText, lngROW, conCOL_SonParamManut_Ativo))
    strEMCONJ = Trim(grdParamManut.Cell(flexcpText, lngROW, conCOL_SonParamManut_EmConjunto))
    
    Select Case intTipoMAnut
           Case 1 '' Por Turno
                intTotDias = intTotDias
           Case 2 '' Por Dia
                intTotDias = intTotDias
           Case 3 '' Por Semana
                intTotDias = (intTotDias / 7)
           Case 4 '' Por Quinzena
                intTotDias = (intTotDias / 15)
           Case 5 '' Por Mês
                intTotDias = (intTotDias / 30)
           Case 6 '' Por Trimestre
                intTotDias = (intTotDias / 90)
           Case 7 '' Por Semestre
                intTotDias = (intTotDias / 180)
           Case 8 '' Por Ano
                intTotDias = (intTotDias / 365)
           Case 9 '' Customizado
                intTotDias = (dtManutFin - dtManut)
                If intTotDias = 0 Then intTotDias = 1
           
    End Select

    For i = 1 To intTotDias
        With grdAgendaManut
             If intTipoMAnut = 2 Or _
                intTipoMAnut = 1 Then               '' Por Dia ou Turno
                dtManut = (dtManut + 1)
             ElseIf intTipoMAnut = 3 Then           '' Por Semana
                dtManut = (dtManut + 7)
             ElseIf intTipoMAnut = 4 Then           '' Por Quinzena
                dtManut = (dtManut + 15)
             ElseIf intTipoMAnut = 5 Then           '' Por Mês
                dtManut = (dtManut + 30)
             ElseIf intTipoMAnut = 6 Then           '' Por Trimestre
                dtManut = (dtManut + 90)
             ElseIf intTipoMAnut = 7 Then           '' Por Semestre
                dtManut = (dtManut + 180)
             ElseIf intTipoMAnut = 8 Then           '' Por Ano
                dtManut = (dtManut + 365)
             ElseIf intTipoMAnut = 9 Then           '' Customizado
                dtManut = (dtManut + 1)
             End If
             
             .AddItem Format(dtManut, "DD/MM/YYYY") & vbTab & _
                             Trim(strHORAINI) & vbTab & _
                             Trim(strHORAFIN) & vbTab & _
                             Trim(strTOTALUS) & vbTab & _
                             strMAQUSO & vbTab & _
                             strATIVO & vbTab & _
                             strEMCONJ & vbTab & _
                             Trim(strINDICE) & vbTab & _
                             grdParamManut.Cell(flexcpText, grdParamManut.Row, conCOL_SonParamManut_TipoManut) & vbTab & _
                             grdParamManut.Cell(flexcpText, grdParamManut.Row, conCOL_SonParammanut_DataInici)
            
             '' Formatado para Campo Hora
             .ColEditMask(conCOL_SonDiasManut_DataManut) = "##/##/####"
             .ColEditMask(conCOL_SonDiasManut_HoraInic) = "##:##"
             .ColEditMask(conCOL_SonDiasManut_HoraFina) = "##:##"
             .ColEditMask(conCOL_SonDiasManut_Tempo) = "##:##"
                              
             '' Alinhamento dos Campos
             .ColAlignment(conCOL_SonDiasManut_HoraInic) = flexAlignRightCenter
             .ColAlignment(conCOL_SonDiasManut_HoraFina) = flexAlignRightCenter
             .ColAlignment(conCOL_SonDiasManut_Tempo) = flexAlignRightCenter
        
        End With
    Next i

End Sub

Private Sub DeleteDias(strTipoMAnut As String)
    Dim i As Integer
    With grdAgendaManut
VOLTA:
         For i = 1 To (.Rows - 1)
             If Trim(.Cell(flexcpText, i, conCOL_SonDiasManut_Pai)) = Trim(strTipoMAnut) Then
               If (.Rows - 1) = 1 Then .Rows = 1
               If (.Rows - 1) > 1 Then
                  .RemoveItem i
                  GoTo VOLTA
               End If
             End If
         Next i
    End With
End Sub

Private Function VerifItensRepetidos(intROW As Long, intCol As Long, varCampo As Variant) As Boolean
    VerifItensRepetidos = False
    Dim i As Integer
    
    If Not IsNumeric(varCampo) Then varCampo = UCase(Trim(varCampo))
    For i = 1 To (grdParamManut.Rows - 1)
        If i <> intROW And grdParamManut.Cell(flexcpText, i, intCol) = varCampo Then
           VerifItensRepetidos = True
           Exit Function
        End If
    Next i
    
End Function

Private Sub PopTotHoras(strVALOR As String, lngCOL As Long, strINDICE As String)
    
    Dim i               As Integer
    Dim strTOTALPERIODO As String
    
    With grdAgendaManut
         For i = 1 To (.Rows - 1)
             If Trim(.Cell(flexcpText, i, conCOL_SonDiasManut_Pai)) = Trim(strINDICE) Then
                .Cell(flexcpText, i, lngCOL) = Format(strVALOR, "HH:MM")
                strTOTALPERIODO = objBLBFunc.CalcTempo(.Cell(flexcpText, i, conCOL_SonDiasManut_HoraInic), .Cell(flexcpText, i, conCOL_SonDiasManut_HoraFina))
                .Cell(flexcpText, i, conCOL_SonDiasManut_Tempo) = Format(CDate(strTOTALPERIODO), "HH:MM")
             End If
         Next i
    End With
End Sub

Private Sub PopSN(strVALOR As String, lngCOL As Long, strINDICE As String)
    
    Dim i               As Integer
    With grdAgendaManut
         For i = 1 To (.Rows - 1)
             If Trim(.Cell(flexcpText, i, conCOL_SonDiasManut_Pai)) = Trim(strINDICE) Then
                .Cell(flexcpText, i, lngCOL) = Trim(strVALOR)
             End If
         Next i
    End With
End Sub


Private Sub MudaAtvo(lngROW As Long)
    Dim i As Integer
    Dim j As Integer
    With grdParamManut
        For i = 1 To (.Rows - 1)
            If lngROW <> i Then
               .Cell(flexcpText, i, conCOL_SonParamManut_Ativo) = 0
               For j = 1 To (grdAgendaManut.Rows - 1)
                   If Trim(grdAgendaManut.Cell(flexcpText, j, conCOL_SonDiasManut_Pai)) = Trim(.Cell(flexcpText, i, conCOL_SonParamManut_Indice)) Then
                      grdAgendaManut.Cell(flexcpText, j, conCOL_SonDiasManut_Ativo) = "0"
                   End If
               Next j
            End If
        Next i
    End With
End Sub

Private Sub PopGrdParamDias()

    Dim i As Integer
    
    arrPARAMETROS = objCADMAQUINA.PARAMETROS
    If IsArray(arrPARAMETROS) Then
       For i = 1 To UBound(arrPARAMETROS)
           grdParamManut.AddItem arrPARAMETROS(i, 1) & vbTab & _
                                 arrPARAMETROS(i, 2) & vbTab & _
                                 arrPARAMETROS(i, 3) & vbTab & _
                                 arrPARAMETROS(i, 4) & vbTab & _
                                 arrPARAMETROS(i, 5) & vbTab & _
                                 arrPARAMETROS(i, 6) & vbTab & _
                                 arrPARAMETROS(i, 7) & vbTab & _
                                 arrPARAMETROS(i, 8) & vbTab & _
                                 arrPARAMETROS(i, 9) & vbTab & _
                                 arrPARAMETROS(i, 1) & arrPARAMETROS(i, 2)
       Next i
    End If
    
    arrDIASPARAMETROS = objCADMAQUINA.DIASPARAMETROS
    If IsArray(arrDIASPARAMETROS) Then
       For i = 1 To UBound(arrDIASPARAMETROS)
           grdAgendaManut.AddItem arrDIASPARAMETROS(i, 1) & vbTab & _
                                  arrDIASPARAMETROS(i, 2) & vbTab & _
                                  arrDIASPARAMETROS(i, 3) & vbTab & _
                                  arrDIASPARAMETROS(i, 4) & vbTab & _
                                  arrDIASPARAMETROS(i, 5) & vbTab & _
                                  arrDIASPARAMETROS(i, 6) & vbTab & _
                                  arrDIASPARAMETROS(i, 7) & vbTab & _
                                  arrDIASPARAMETROS(i, 8) & arrDIASPARAMETROS(i, 9) & vbTab & _
                                  arrDIASPARAMETROS(i, 8) & vbTab & _
                                  arrDIASPARAMETROS(i, 9)
       Next i
    End If
End Sub

Private Sub InitGridCapCorte()

    With grdCAPPROD
    
       .Cols = conColumnsIn_SonCap
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonCap_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonCap_Itens) = ""
       .ColDataType(conCOL_SonCap_Itens) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonCap_IDLinha) = ""
       .ColDataType(conCOL_SonCap_IDLinha) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonCap_CodLinha) = ""
       .ColDataType(conCOL_SonCap_CodLinha) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonCap_PesqLinha) = ""
       .ColDataType(conCOL_SonCap_PesqLinha) = flexDTString
       .ColComboList(conCOL_SonCap_PesqLinha) = "..."
       
       .Cell(flexcpData, 0, conCOL_SonCap_DescLinha) = ""
       .ColDataType(conCOL_SonCap_DescLinha) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonCap_CodSeqCorte) = ""
       .ColDataType(conCOL_SonCap_CodSeqCorte) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonCap_CodCorte) = ""
       .ColDataType(conCOL_SonCap_CodCorte) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonCap_PesqCorte) = ""
       .ColDataType(conCOL_SonCap_PesqCorte) = flexDTString
       .ColComboList(conCOL_SonCap_PesqCorte) = "..."
       
       .Cell(flexcpData, 0, conCOL_SonCap_DescCorte) = ""
       .ColDataType(conCOL_SonCap_DescCorte) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonCap_INDICE) = ""
       .ColDataType(conCOL_SonCap_INDICE) = flexDTString
       
       .ColWidth(conCOL_SonCap_Itens) = 500
       .ColWidth(conCOL_SonCap_IDLinha) = 0
       .ColWidth(conCOL_SonCap_CodLinha) = 1000
       .ColWidth(conCOL_SonCap_PesqLinha) = 300
       .ColWidth(conCOL_SonCap_DescLinha) = 4000
       .ColWidth(conCOL_SonCap_CodSeqCorte) = 500
       .ColWidth(conCOL_SonCap_CodCorte) = 1000
       .ColWidth(conCOL_SonCap_PesqCorte) = 300
       .ColWidth(conCOL_SonCap_DescCorte) = 4000
       .ColWidth(conCOL_SonCap_INDICE) = 0
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
    End With
    
End Sub

Private Sub IncRegGrid()
   
    If objBLBFunc.FcExisteLinhaVazia(grdCAPPROD, conCOL_SonCap_CodLinha) = False Then Exit Sub
    
    grdCAPPROD.AddItem "" & vbTab & _
                       "" & vbTab & _
                       "" & vbTab & _
                       "" & vbTab & _
                       "" & vbTab & _
                       "" & vbTab & _
                       "" & vbTab & _
                       "" & vbTab & _
                       ""
    
    Call RefazIndice
    
End Sub


Private Sub RefazIndice()

    Dim i As Integer
    
    With grdCAPPROD
        For i = 1 To (.Rows - 1)
            .Cell(flexcpText, i, conCOL_SonCap_Itens) = i
        Next i
    End With
End Sub


Private Function PegaDescrLinha(strCODLINHA As String, lngLINHA As Long) As Boolean

    PegaDescrLinha = False
    
    With grdCAPPROD
    
        sSql = ""
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       * " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADLINHAPRODUTO " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_CODLIN = " & strCODLINHA
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC.EOF() Then
           .Cell(flexcpText, lngLINHA, conCOL_SonCap_IDLinha) = BREC!SGI_CODIGO
           .Cell(flexcpText, lngLINHA, conCOL_SonCap_DescLinha) = BREC!SGI_DESCRI
           PegaDescrLinha = True
        Else
            MsgBox "Esta Linha de Produto não Existe !!!", vbOKOnly + vbExclamation, "Aviso"
            Call LimpaCamposGridLinhaProd(lngLINHA)
        End If
        BREC.Close
    
    End With
    
End Function

Private Sub LimpaCamposGridLinhaProd(lngLINHA As Long)
    With grdCAPPROD
        .Cell(flexcpText, lngLINHA, conCOL_SonCap_IDLinha, lngLINHA, conCOL_SonCap_CodSeqCorte) = Empty
        .Cell(flexcpText, lngLINHA, conCOL_SonCap_IDLinha, lngLINHA, conCOL_SonCap_CodCorte) = Empty
        .Cell(flexcpText, lngLINHA, conCOL_SonCap_IDLinha, lngLINHA, conCOL_SonCap_DescCorte) = Empty
        .Cell(flexcpText, lngLINHA, conCOL_SonCap_IDLinha, lngLINHA, conCOL_SonCap_CodLinha) = Empty
        .Cell(flexcpText, lngLINHA, conCOL_SonCap_IDLinha, lngLINHA, conCOL_SonCap_DescLinha) = Empty
        .Cell(flexcpText, lngLINHA, conCOL_SonCap_IDLinha, lngLINHA, conCOL_SonCap_INDICE) = Empty
    End With
End Sub


Private Function PegaDescrCorte(strSEGCORTE As String, strCODLINHA As String, lngLINHA As Long) As Boolean

    PegaDescrCorte = False
    
    
    Dim strINDICE As String
    
    With grdCAPPROD
    
        sSql = ""
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       CODLIN.SGI_ITENCOR" & vbCrLf
        sSql = sSql & "      ,CODLIN.SGI_CODMEDCORT" & vbCrLf
        sSql = sSql & "      ,DIMCO.SGI_DESCORTE" & vbCrLf
        
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_MEDCORTELINHA CODLIN" & vbCrLf
        sSql = sSql & "      ,SGI_CADDIMCORTE   DIMCO " & vbCrLf

        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       CODLIN.SGI_FILIAL  = " & FILIAL & vbCrLf
        sSql = sSql & "   And CODLIN.SGI_CODIGO  = " & Trim(strCODLINHA) & vbCrLf
        sSql = sSql & "   And CODLIN.SGI_ITENCOR = " & Trim(strSEGCORTE) & vbCrLf
        sSql = sSql & "   And DIMCO.SGI_FILIAL   = CODLIN.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And DIMCO.SGI_CODIGO   = CODLIN.SGI_CODMEDCORT" & vbCrLf
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC.EOF() Then
           .Cell(flexcpText, lngLINHA, conCOL_SonCap_CodSeqCorte) = BREC!SGI_ITENCOR
           .Cell(flexcpText, lngLINHA, conCOL_SonCap_CodCorte) = BREC!SGI_CODMEDCORT
           .Cell(flexcpText, lngLINHA, conCOL_SonCap_DescCorte) = BREC!SGI_DESCORTE
           
           strINDICE = Trim(.Cell(flexcpText, lngLINHA, conCOL_SonCap_IDLinha)) & _
                       Trim(.Cell(flexcpText, lngLINHA, conCOL_SonCap_CodSeqCorte)) & _
                       Trim(.Cell(flexcpText, lngLINHA, conCOL_SonCap_CodCorte))
           
           .Cell(flexcpText, lngLINHA, conCOL_SonCap_INDICE) = strINDICE
           PegaDescrCorte = True
        Else
            MsgBox "Este Corte não Existe !!!", vbOKOnly + vbExclamation, "Aviso"
            Call LimpaCamposCorte(lngLINHA)
        End If
        BREC.Close
    
    End With
    
End Function


Private Sub LimpaCamposCorte(lngROW As Long)
    With grdCAPPROD
        .Cell(flexcpText, lngROW, conCOL_SonCap_CodSeqCorte) = Empty
        .Cell(flexcpText, lngROW, conCOL_SonCap_CodCorte) = Empty
        .Cell(flexcpText, lngROW, conCOL_SonCap_DescCorte) = Empty
        .Cell(flexcpText, lngROW, conCOL_SonCap_INDICE) = Empty
    End With
End Sub

Private Sub PopGridLinha()
    
    Dim i As Long
    
    arrLINHACAP = objCADMAQUINA.LINHACAP
    
    If IsArray(arrLINHACAP) Then
        With grdCAPPROD
            For i = 1 To UBound(arrLINHACAP)
                .AddItem arrLINHACAP(i, 1) & vbTab & _
                         arrLINHACAP(i, 2) & vbTab & _
                         arrLINHACAP(i, 3) & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         arrLINHACAP(i, 4) & vbTab & _
                         arrLINHACAP(i, 5) & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         arrLINHACAP(i, 6)
                         
                Call PegaDescrLinha(Str(arrLINHACAP(i, 3)), i)
                Call PegaDescrCorte(Str(arrLINHACAP(i, 4)), Str(arrLINHACAP(i, 2)), i)
            
            Next i
        End With
    End If

End Sub

Private Sub LimpaCamposLabel()
    lblDESGRPMAQ.Caption = ""
End Sub
