VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmCADCOTAVENDA 
   Caption         =   "Cadastro de cotação de vendas"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12810
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   12810
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab stOrca 
      Height          =   6855
      Left            =   0
      TabIndex        =   29
      Top             =   1080
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   12091
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
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
      TabCaption(0)   =   "Dados Básico"
      TabPicture(0)   =   "frmCADCOTAVENDA.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame9"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Itens do Orçamento"
      TabPicture(1)   =   "frmCADCOTAVENDA.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label20"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame8"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame7"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtTotalItens"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.Frame Frame9 
         Caption         =   "[ Totais do Pedido ]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   -74880
         TabIndex        =   65
         Top             =   4620
         Width           =   12495
         Begin VB.TextBox txtVLDESCTO 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4560
            TabIndex        =   70
            Text            =   "txtVLDESCTO"
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox txtPDESCTOTAL 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2040
            TabIndex        =   69
            Text            =   "txtPDESCTOTAL"
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox txtFRETE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4560
            TabIndex        =   68
            Text            =   "txtFRETE"
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtALIQICMS 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4560
            TabIndex        =   67
            Text            =   "txtALIQICMS"
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtOutrDesp 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2040
            TabIndex        =   66
            Text            =   "txtOutrDesp"
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Vl.Desconto"
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
            Left            =   3480
            TabIndex        =   85
            Top             =   1320
            Width           =   1050
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Desconto%:"
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
            Left            =   960
            TabIndex        =   84
            Top             =   1320
            Width           =   1020
         End
         Begin VB.Label lblVLDESC 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblVLDESC"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2040
            TabIndex        =   83
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblVLIPI 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblVLIPI"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7320
            TabIndex        =   82
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label lblVLTOTAL 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblVLTOTAL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2040
            TabIndex        =   81
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Vl. Descomto p.Iten:"
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
            TabIndex        =   80
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Total do Orçamento:"
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
            Left            =   120
            TabIndex        =   79
            Top             =   1680
            Width           =   1815
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Vl. do IPI:"
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
            Left            =   6360
            TabIndex        =   78
            Top             =   630
            Width           =   870
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Frete:"
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
            Left            =   3840
            TabIndex        =   77
            Top             =   630
            Width           =   510
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Outras despesas:"
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
            Left            =   360
            TabIndex        =   76
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label lblVLICMS 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblVLICMS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7320
            TabIndex        =   75
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Valor do ICMS:"
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
            Left            =   6000
            TabIndex        =   74
            Top             =   285
            Width           =   1290
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Aliq ICMS%:"
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
            Left            =   3480
            TabIndex        =   73
            Top             =   285
            Width           =   1035
         End
         Begin VB.Label lblBASICMS 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblBASICMS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2040
            TabIndex        =   72
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Base Calculo ICMS:"
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
            TabIndex        =   71
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.TextBox txtTotalItens 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   10560
         Locked          =   -1  'True
         TabIndex        =   61
         Text            =   "txtTotalItens"
         Top             =   6480
         Width           =   1935
      End
      Begin VB.Frame Frame7 
         Height          =   5175
         Left            =   120
         TabIndex        =   50
         Top             =   1200
         Width           =   12495
         Begin VB.Frame fraPedidos 
            Caption         =   "[ Pedidos Gerados ]"
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
            Height          =   3015
            Left            =   9000
            TabIndex        =   86
            Top             =   240
            Width           =   3375
            Begin VSFlex8LCtl.VSFlexGrid grdPedidos 
               Height          =   2655
               Left            =   120
               TabIndex        =   87
               Top             =   240
               Width           =   3135
               _cx             =   5530
               _cy             =   4683
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
         Begin MSFlexGridLib.MSFlexGrid flxGridProd 
            Height          =   4815
            Left            =   120
            TabIndex        =   51
            Top             =   240
            Width           =   12255
            _ExtentX        =   21616
            _ExtentY        =   8493
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame8 
         Height          =   855
         Left            =   120
         TabIndex        =   43
         Top             =   360
         Width           =   12495
         Begin VB.TextBox txtPORCDESC 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   8880
            TabIndex        =   21
            Text            =   "txtPORCDESC"
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton cmdProd 
            Height          =   315
            Left            =   1320
            Picture         =   "frmCADCOTAVENDA.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtCodProd 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            TabIndex        =   18
            Text            =   "txtCodProd"
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtQtdCompra 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6360
            TabIndex        =   19
            Text            =   "txtQtdCompra"
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtVlUnitarioCompra 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7560
            TabIndex        =   20
            Text            =   "txtVlUnitarioCompra"
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtIPICompra 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   9840
            TabIndex        =   22
            Text            =   "txtIPICompra"
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtVlTotalCompra 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   10680
            TabIndex        =   23
            Text            =   "txtVlTotalCompra"
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton cmbGravPagto 
            Height          =   315
            Left            =   12000
            Picture         =   "frmCADCOTAVENDA.frx":013A
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   480
            Width           =   375
         End
         Begin VB.Label lblDescProd 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDescProd"
            Height          =   285
            Left            =   1680
            TabIndex        =   63
            Top             =   480
            Width           =   4575
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "% Desc."
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
            Left            =   9000
            TabIndex        =   52
            Top             =   240
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Codigo Produto"
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
            TabIndex        =   49
            Top             =   240
            Width           =   1320
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Quantide"
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
            Left            =   6600
            TabIndex        =   48
            Top             =   240
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Vl. Unitário"
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
            Left            =   7800
            TabIndex        =   47
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "% IPI"
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
            Left            =   10080
            TabIndex        =   46
            Top             =   240
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Vl. Total"
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
            Index           =   11
            Left            =   11280
            TabIndex        =   45
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   40
         Top             =   2520
         Width           =   12495
         Begin VB.ComboBox cboTelefone 
            Height          =   315
            Left            =   1920
            TabIndex        =   11
            Text            =   "cboTelefone"
            Top             =   1680
            Width           =   3255
         End
         Begin VB.ComboBox cboEMail 
            Height          =   315
            Left            =   1920
            TabIndex        =   10
            Text            =   "cboEMail"
            Top             =   1320
            Width           =   3255
         End
         Begin VB.TextBox txtDepto 
            Height          =   285
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   9
            Text            =   "txtDepto"
            Top             =   960
            Width           =   3255
         End
         Begin VB.ComboBox cboContato 
            Height          =   315
            Left            =   1920
            TabIndex        =   8
            Text            =   "cboContato"
            Top             =   600
            Width           =   3255
         End
         Begin VB.TextBox txtVALORC 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   11640
            TabIndex        =   13
            Text            =   "txtVALORC"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtPRZENTREGA 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5280
            MaxLength       =   20
            TabIndex        =   12
            Text            =   "txtPRZENTREGA"
            Top             =   240
            Width           =   735
         End
         Begin MSMask.MaskEdBox mskDtValidade 
            Height          =   285
            Left            =   8040
            TabIndex        =   7
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskDtEntrega 
            Height          =   285
            Left            =   1920
            TabIndex        =   6
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Data da Entrega:"
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
            TabIndex        =   64
            Top             =   240
            Width           =   1470
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Data da Validade:"
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
            Left            =   6240
            TabIndex        =   60
            Top             =   240
            Width           =   1545
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Telefone:"
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
            TabIndex        =   56
            Top             =   1680
            Width           =   825
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "E-Mail:"
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
            TabIndex        =   55
            Top             =   1380
            Width           =   600
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Departamento:"
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
            Left            =   120
            TabIndex        =   54
            Top             =   1005
            Width           =   1260
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Contato:"
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
            TabIndex        =   53
            Top             =   645
            Width           =   735
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Validade da Proposta:"
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
            Left            =   9600
            TabIndex        =   42
            Top             =   240
            Width           =   1890
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Prazo de Entrega:"
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
            Left            =   3360
            TabIndex        =   41
            Top             =   240
            Width           =   1545
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   30
         Top             =   360
         Width           =   12495
         Begin VB.ComboBox cboTIPORC 
            Height          =   315
            Left            =   2880
            TabIndex        =   17
            Text            =   "cboTIPORC"
            Top             =   960
            Width           =   7455
         End
         Begin VB.CommandButton Command3 
            Height          =   315
            Left            =   2520
            Picture         =   "frmCADCOTAVENDA.frx":023C
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   960
            Width           =   375
         End
         Begin VB.TextBox txtCODTIPORC 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   3
            Text            =   "txtCODTIPO"
            Top             =   975
            Width           =   1215
         End
         Begin VB.ComboBox cboCONDPGTO 
            Height          =   315
            Left            =   2880
            TabIndex        =   16
            Text            =   "cboCONDPGTO"
            Top             =   1680
            Width           =   7455
         End
         Begin VB.CommandButton Command2 
            Height          =   315
            Left            =   2520
            Picture         =   "frmCADCOTAVENDA.frx":033E
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   1680
            Width           =   375
         End
         Begin VB.TextBox txtCODCONDPGT 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   5
            Text            =   "txtCODCOND"
            Top             =   1695
            Width           =   1215
         End
         Begin VB.ComboBox cboVendedor 
            Height          =   315
            Left            =   2880
            TabIndex        =   14
            Text            =   "cboVendedor"
            Top             =   600
            Width           =   7455
         End
         Begin VB.CommandButton Command1 
            Height          =   315
            Left            =   2520
            Picture         =   "frmCADCOTAVENDA.frx":0440
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtCODVEND 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   2
            Text            =   "txtCODVEND"
            Top             =   615
            Width           =   1215
         End
         Begin VB.ComboBox cboCliente 
            Height          =   315
            Left            =   2880
            TabIndex        =   15
            Text            =   "cboCliente"
            Top             =   1320
            Width           =   7455
         End
         Begin VB.CommandButton cmdFornec 
            Height          =   315
            Left            =   2520
            Picture         =   "frmCADCOTAVENDA.frx":0542
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   1320
            Width           =   375
         End
         Begin VB.TextBox txtCODCLI 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   4
            Text            =   "txtCODCLI"
            Top             =   1335
            Width           =   1215
         End
         Begin MSMask.MaskEdBox mskDTPEDIDO 
            Height          =   285
            Left            =   3600
            TabIndex        =   1
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblSTATUS 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblSTATUS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   5760
            TabIndex        =   59
            Top             =   240
            Width           =   1095
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
            Index           =   2
            Left            =   5040
            TabIndex        =   58
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Orçto:"
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
            Left            =   120
            TabIndex        =   39
            Top             =   1005
            Width           =   975
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Cond.Pgto:"
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
            Left            =   120
            TabIndex        =   37
            Top             =   1725
            Width           =   960
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor:"
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
            Left            =   120
            TabIndex        =   35
            Top             =   645
            Width           =   885
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
            TabIndex        =   34
            Top             =   240
            Width           =   660
         End
         Begin VB.Label lblCODIGO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblCODIGO"
            Height          =   255
            Left            =   1320
            TabIndex        =   0
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label3 
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
            Left            =   2880
            TabIndex        =   33
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Cliente:"
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
            Left            =   120
            TabIndex        =   32
            Top             =   1365
            Width           =   660
         End
      End
      Begin VB.Label Label20 
         Caption         =   "Total"
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
         Height          =   255
         Left            =   9960
         TabIndex        =   62
         Top             =   6480
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   12735
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
         Picture         =   "frmCADCOTAVENDA.frx":0644
         Style           =   1  'Graphical
         TabIndex        =   28
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
         Picture         =   "frmCADCOTAVENDA.frx":0B76
         Style           =   1  'Graphical
         TabIndex        =   27
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
         Picture         =   "frmCADCOTAVENDA.frx":0C78
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADCOTAVENDA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho      As String
Public Linha         As Variant
Public cTipOper      As String
Public iCodigo       As Long
Public FILIAL        As Integer
Public strAcesso     As String
Public strMODPAI     As String
Public strUsuario    As String
Public lngCodUsuario As Long
Dim objBLBFunc       As Object
Dim objCADCOTAVENDA  As Object
Dim objPESQPADRAO    As Object
Dim arrTIPDESP       As Variant
Dim arrPRODUTOS      As Variant
Dim arrTIPOSSERV     As Variant
Dim intLinha         As Integer
Dim intColuna        As Integer

Private Sub cboCliente_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboCliente, KeyAscii
End Sub

Private Sub cboCliente_Validate(Cancel As Boolean)
    If cboCliente.ListIndex > -1 Then
       txtCODCLI.Text = cboCliente.ItemData(cboCliente.ListIndex)
       objCADCOTAVENDA.CODCLI = txtCODCLI.Text
       objCADCOTAVENDA.PreencheComboContato cboContato
       objCADCOTAVENDA.PreencheComboEmail cboEMail
       objCADCOTAVENDA.PreencheComboTelefone cboTelefone
    End If
End Sub

Private Sub cboCONDPGTO_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboCONDPGTO, KeyAscii
End Sub

Private Sub cboCONDPGTO_Validate(Cancel As Boolean)
    If cboCONDPGTO.ListIndex > -1 Then
    
       txtCODCONDPGT.Text = cboCONDPGTO.ItemData(cboCONDPGTO.ListIndex)
       ConfGridProd
       
       txtCodProd.Text = ""
       txtQtdCompra.Text = ""
       txtVlUnitarioCompra.Text = ""
       txtIPICompra.Text = ""
       txtVlTotalCompra.Text = ""
       txtPORCDESC.Text = ""
    
    End If
End Sub

Private Sub cboContato_Validate(Cancel As Boolean)
    
    If Len(Trim(cboContato.Text)) > 50 Then
       MsgBox "Somente é permitido 50 digitos !!!", vbOKOnly + vbExclamation, "Aviso"
       Cancel = True
       Exit Sub
    End If
    
End Sub

Private Sub cboEMail_Validate(Cancel As Boolean)

    If Len(Trim(cboEMail.Text)) > 50 Then
       MsgBox "Somente é permitido 50 digitos !!!", vbOKOnly + vbExclamation, "Aviso"
       Cancel = True
       Exit Sub
    End If

End Sub






Private Sub cboTelefone_Validate(Cancel As Boolean)

    If Len(Trim(cboTelefone.Text)) > 50 Then
       MsgBox "Somente é permitido 50 digitos !!!", vbOKOnly + vbExclamation, "Aviso"
       Cancel = True
       Exit Sub
    End If

End Sub
Private Sub cboTIPORC_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboTIPORC, KeyAscii
End Sub

Private Sub cboTIPORC_Validate(Cancel As Boolean)
    If cboTIPORC.ListIndex > -1 Then txtCODTIPORC.Text = cboTIPORC.ItemData(cboTIPORC.ListIndex)
End Sub

Private Sub cboVendedor_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboVendedor, KeyAscii
End Sub

Private Sub cboVendedor_Validate(Cancel As Boolean)
    If cboVendedor.ListIndex > -1 Then txtCODVEND.Text = cboVendedor.ItemData(cboVendedor.ListIndex)
End Sub

Private Sub cmbGravPagto_Click()
    If (cTipOper = "I" Or cTipOper = "A") Then IncProdGridItens
End Sub

Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    If VerifPed = True Then
       MsgBox "Não pode ser alterado existe pedido !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Sub
    End If
    
    stOrca.Tab = 0
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    
    Frame2.Enabled = True
    Frame3.Enabled = True
    Frame8.Enabled = True
    Frame9.Enabled = True
    

    Me.Caption = "Cadastro de orçamentos - [ ALTERAÇÃO ]"
    
    cTipOper = "A"
    
    mskDTPEDIDO.SetFocus

End Sub

Private Sub cmdFornec_Click()

    ReDim arrCAMPOS(1 To 5, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
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
    arrCAMPOS(3, 4) = "3000"
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
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Clientes", "CADCLIENTE.clsCADCLIENTE")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODCLI.Text = varRETORNO
    
    cboCliente.ListIndex = -1
    txtCODCLI.SetFocus

End Sub

Private Sub cmdProd_Click()

   
    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select Case PRO.SGI_PRODUTOTIPO " & vbCrLf
    sSql = sSql & "            When 1 Then replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD))) + '.' + " & vbCrLf
    sSql = sSql & "                        replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE))) + '.' + " & vbCrLf
    sSql = sSql & "                        replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO))) + '.' + " & vbCrLf
    sSql = sSql & "                        (Case " & vbCrLf
    sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Null Then '0' " & vbCrLf
    sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(1),PRO.SGI_DIGVERIF))) End) " & vbCrLf
    sSql = sSql & "            When 0 Then PRO.SGI_CODIGO End As SGI_CODIGO " & vbCrLf
    sSql = sSql & ",SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "         SGI_CADPRODUTO PRO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL     = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "S"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1500"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Produtos")
        
    If Len(Trim(varRETORNO)) > 0 Then txtCodProd.Text = varRETORNO
    
    txtCodProd.Tag = PegaTagProduto(txtCodProd.Text)
    txtQtdCompra.SetFocus

End Sub

Private Sub CmdSalva_Click()

On Error GoTo err_grava
    
    Dim I         As Integer
    Dim j         As Integer
    Dim intResp   As Integer
    Dim strESPTEC As String
    Dim strOUTRAS As String
        
    If ValidaCampos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADCOTAVENDA.CODIGO = objCADCOTAVENDA.Gera_Codigo(Me.Name) & Year(Now)
    
    objCADCOTAVENDA.DTCOTACAO = CDate(mskDTPEDIDO.Text)
    objCADCOTAVENDA.CODCLI = CInt(txtCODCLI.Text)
    objCADCOTAVENDA.CODVEND = CInt(txtCODVEND.Text)
    objCADCOTAVENDA.CODCONDPGT = CInt(txtCODCONDPGT.Text)
    objCADCOTAVENDA.CODTIPORC = CInt(txtCODTIPORC.Text)
    
    
    If Len(Trim(txtPRZENTREGA.Text)) > 0 Then objCADCOTAVENDA.PRZENTREGA = CInt(txtPRZENTREGA.Text)
    If Len(Trim(txtVALORC.Text)) > 0 Then objCADCOTAVENDA.VALCOTACAO = CInt(txtVALORC.Text)
    
    objCADCOTAVENDA.CONTATO = cboContato.Text
    objCADCOTAVENDA.DEPTO = txtDepto.Text
    objCADCOTAVENDA.EMAIL = cboEMail.Text
    objCADCOTAVENDA.TELCLIE = cboTelefone.Text
    
    objCADCOTAVENDA.BASEICMS = 0
    If Len(Trim(lblBASICMS.Caption)) > 0 Then objCADCOTAVENDA.BASEICMS = CCur(lblBASICMS.Caption)
    
    objCADCOTAVENDA.ALIQICMS = 0
    If Len(Trim(txtALIQICMS.Text)) > 0 Then objCADCOTAVENDA.ALIQICMS = CCur(txtALIQICMS.Text)
    
    objCADCOTAVENDA.VLICMS = 0
    If Len(Trim(lblVLICMS.Caption)) > 0 Then objCADCOTAVENDA.VLICMS = CCur(lblVLICMS.Caption)
    
    objCADCOTAVENDA.TOTOUTDESP = 0
    If Len(Trim(txtOutrDesp.Text)) > 0 Then objCADCOTAVENDA.TOTOUTDESP = CCur(txtOutrDesp.Text)
    
    objCADCOTAVENDA.TOTFRETE = 0
    If Len(Trim(txtFRETE.Text)) > 0 Then objCADCOTAVENDA.TOTFRETE = CCur(txtFRETE.Text)
    
    objCADCOTAVENDA.VLIPI = 0
    If Len(Trim(lblVLIPI.Caption)) > 0 Then objCADCOTAVENDA.VLIPI = CCur(lblVLIPI.Caption)
    
    objCADCOTAVENDA.VLDESCTO = 0
    If Len(Trim(lblVLDESC.Caption)) > 0 Then objCADCOTAVENDA.VLDESCTO = CCur(lblVLDESC.Caption)
    
    objCADCOTAVENDA.TOTORCTO = 0
    If Len(Trim(lblVLTOTAL.Caption)) > 0 Then objCADCOTAVENDA.TOTORCTO = CCur(lblVLTOTAL.Caption)
    
    objCADCOTAVENDA.VLPDESCT = 0
    objCADCOTAVENDA.VLTOTDESC = 0
    objCADCOTAVENDA.VLSERVIC = 0
    If Len(Trim(txtPDESCTOTAL.Text)) > 0 Then objCADCOTAVENDA.VLPDESCT = CCur(txtPDESCTOTAL.Text)
    If Len(Trim(txtVLDESCTO.Text)) > 0 Then objCADCOTAVENDA.VLTOTDESC = CCur(txtVLDESCTO.Text)
    
    arrPRODUTOS = Empty
    If flxGridProd.Rows > 1 Then
       ReDim arrPRODUTOS(1 To (flxGridProd.Rows - 1), 1 To 7) As Variant
       For I = 1 To (flxGridProd.Rows - 1)
           
           arrPRODUTOS(I, 7) = flxGridProd.TextMatrix(I, 0)
           arrPRODUTOS(I, 1) = flxGridProd.TextMatrix(I, 1)
           
           arrPRODUTOS(I, 2) = CCur(flxGridProd.TextMatrix(I, 3))
           arrPRODUTOS(I, 3) = CCur(flxGridProd.TextMatrix(I, 4))
           
           arrPRODUTOS(I, 4) = CCur("0")
           If Len(Trim(flxGridProd.TextMatrix(I, 5))) > 0 Then arrPRODUTOS(I, 4) = CCur(flxGridProd.TextMatrix(I, 5))
           
           arrPRODUTOS(I, 5) = CCur("0")
           If Len(Trim(flxGridProd.TextMatrix(I, 6))) > 0 Then arrPRODUTOS(I, 5) = CCur(flxGridProd.TextMatrix(I, 6))
           
           arrPRODUTOS(I, 6) = CCur("0")
           If Len(Trim(flxGridProd.TextMatrix(I, 7))) > 0 Then arrPRODUTOS(I, 6) = CCur(flxGridProd.TextMatrix(I, 7))
           
       Next I
       objCADCOTAVENDA.Produtos = arrPRODUTOS
    End If
    
    objCADCOTAVENDA.TOTALITENS = (flxGridProd.Rows - 1)
    objCADCOTAVENDA.DATVALCOTA = CDate(mskDtValidade.Text)
    objCADCOTAVENDA.DATENTREGA = CDate(mskDtEntrega.Text)

    If objCADCOTAVENDA.GRAVA(cTipOper) = False Then Exit Sub
    If objCADCOTAVENDA.Atualiza(cTipOper, Str(objCADCOTAVENDA.CODIGO), FILIAL, Me.Name) = False Then Exit Sub
          
    MsgBox "A cotação de venda foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", IIf(cTipOper = "L", "liberada", ""))) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
    
    If cTipOper = "I" Then
       intResp = MsgBox("Deseja gerar outra cotação ?", vbYesNo + vbQuestion, "Aviso")
       
       If intResp = 6 Then
          Inclui
          stOrca.Tab = 0
          mskDTPEDIDO.SetFocus
       Else
          Set objBLBFunc = Nothing
          Set objCADCOTAVENDA = Nothing
          Set objPESQPADRAO = Nothing
          Unload Me
       End If
    
    End If
    
    
    Exit Sub
    
err_grava:
    
    MsgBox "Erro nº: " & Err.Number & " Descrição: " & Err.Description, vbOKOnly + vbCritical, "Aviso"
    
End Sub

Private Sub cmdVoltar_Click()
  Set objBLBFunc = Nothing
  Set objCADCOTAVENDA = Nothing
  Set objPESQPADRAO = Nothing
  Unload Me
End Sub

Private Sub Command1_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "        SGI_CODIGO    " & vbCrLf
    sSql = sSql & "       ,SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "        SGI_CADVENDEDOR " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Venderores", "CADVENDEDOR.clsCADVENDEDOR")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODVEND.Text = varRETORNO
    
    cboVendedor.ListIndex = -1
    txtCODVEND.SetFocus

End Sub

Private Sub Command2_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "        SGI_CODIGO    " & vbCrLf
    sSql = sSql & "       ,SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "        SGI_CADCONDPGTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Condição de Pagamento")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODCONDPGT.Text = varRETORNO
        
    cboCONDPGTO.ListIndex = -1
    txtCODCONDPGT.SetFocus

End Sub

Private Sub Command3_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "        SGI_CODIGO    " & vbCrLf
    sSql = sSql & "       ,SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "        SGI_CADESPORCA " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "        SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Tipo de Orçamento", "CADESPORCA.clsCADESPORCA")
    
    If Len(Trim(varRETORNO)) > 0 Then txtCODTIPORC.Text = varRETORNO
        
    cboTIPORC.ListIndex = -1
    txtCODTIPORC.SetFocus

End Sub

Private Sub flxGridProd_Click()

    If flxGridProd.Rows - 1 = 0 Then Exit Sub
    
    If flxGridProd.Rows > 1 Then
       If Len(Trim(flxGridProd.TextMatrix(flxGridProd.Row, 11))) = 0 Then
          flxGridProd.TextMatrix(flxGridProd.Row, 11) = "***"
       End If
    End If
    
End Sub

Private Sub flxGridProd_KeyDown(KeyCode As Integer, Shift As Integer)
    If flxGridProd.Rows = 1 Then Exit Sub
    If cTipOper = "C" Then Exit Sub
    If KeyCode = vbKeyDelete Then
       If flxGridProd.Rows = 2 Then flxGridProd.Rows = 1
       If flxGridProd.Rows > 2 Then flxGridProd.RemoveItem flxGridProd.Row
       CalcTotORc
       txtTotalItens.Text = Format(SomaGrdItens, "#,##0.00")
    End If
End Sub

Private Sub flxGridProd_LeaveCell()
    If (flxGridProd.Rows - 1) > 0 Then
       If Trim(flxGridProd.TextMatrix(flxGridProd.Row, 11)) = "***" Then
          flxGridProd.TextMatrix(flxGridProd.Row, 11) = ""
       End If
    End If
End Sub

Private Sub flxGridProd_RowColChange()
    
    If flxGridProd.Rows - 1 = 0 Then Exit Sub
    
    If flxGridProd.Rows > 1 Then
       If Len(Trim(flxGridProd.TextMatrix(flxGridProd.Row, 11))) = 0 Then
          flxGridProd.TextMatrix(flxGridProd.Row, 11) = "***"
       End If
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
   Set objCADCOTAVENDA = CreateObject("CADCOTAVENDA.clsCADCOTAVENDA")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
   objCADCOTAVENDA.FILIAL = FILIAL
   
   Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

   If adoBanco_Dados.State = 0 Then
      MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
      Exit Sub
   End If
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

   stOrca.Tab = 0
    
    
End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    
    lblSTATUS.Caption = "ABERTO"
    lblSTATUS.ForeColor = &HFF&
    objCADCOTAVENDA.STATUS = "A"
    
    lblDescProd.Caption = ""
        
    fraPedidos.Visible = False '' Mostra Pedidos Gerados
    
    
    Frame2.Enabled = True
    Me.Caption = "Cadastro de orçamentos - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADCOTAVENDA
    
    mskDTPEDIDO.Text = Format(Date, "DD/MM/YYYY")
    mskDtValidade.Text = "__/__/____"
    mskDtEntrega.Text = "__/__/____"
        
    ReDim arrTIPOSSERV(1 To 100, 1 To 100, 1 To 6) As String
    
    objCADCOTAVENDA.PreencheComboVendedor cboVendedor, objBLBFunc.Crypt(Trim(strUsuario))
    objCADCOTAVENDA.PreencheComboTipoOrc cboTIPORC
    
    objCADCOTAVENDA.PreencheComboCliente cboCliente
    objCADCOTAVENDA.PreencheComboCondPgto cboCONDPGTO
    
    ConfGridProd
     
    lblCODIGO.Caption = ""
    
    lblBASICMS.Caption = ""
    lblVLICMS.Caption = ""
    lblVLIPI.Caption = ""
    lblVLTOTAL.Caption = ""
    lblVLDESC.Caption = ""
    txtOutrDesp.Text = ""
    
    Call Pega_Vendedor(lngCodUsuario)
        
    txtPRZENTREGA.Enabled = False
    txtVALORC.Enabled = False
    
End Sub

Private Sub ConfGridProd()
        
    flxGridProd.Rows = 1
    flxGridProd.Cols = 12
    
    flxGridProd.TextMatrix(0, 0) = ""
    flxGridProd.TextMatrix(0, 1) = "Produto"
    flxGridProd.TextMatrix(0, 2) = "Descrição"
    flxGridProd.TextMatrix(0, 3) = "Qtde"
    flxGridProd.TextMatrix(0, 4) = "Vl.Unit."
    flxGridProd.TextMatrix(0, 5) = "% Desc"
    flxGridProd.TextMatrix(0, 6) = "% IPI"
    flxGridProd.TextMatrix(0, 7) = "Vl. Total"
    flxGridProd.TextMatrix(0, 8) = "Qtde. Ped"
    flxGridProd.TextMatrix(0, 9) = "Saldo"
    flxGridProd.TextMatrix(0, 10) = "Status"
    flxGridProd.TextMatrix(0, 11) = "   "
    
    flxGridProd.ColWidth(0) = 0
    flxGridProd.ColWidth(1) = 1500
    flxGridProd.ColWidth(2) = 4000
    flxGridProd.ColWidth(3) = 1000
    flxGridProd.ColWidth(4) = 1000
    flxGridProd.ColWidth(5) = 1500
    flxGridProd.ColWidth(6) = 1000
    flxGridProd.ColWidth(7) = 1000
    flxGridProd.ColWidth(8) = 0
    flxGridProd.ColWidth(9) = 0
    flxGridProd.ColWidth(10) = 0
    flxGridProd.ColWidth(11) = 300
    
End Sub

Private Sub mskDtEntrega_Validate(Cancel As Boolean)
    If mskDtEntrega.Text <> "__/__/____" Then
        If Not IsDate(mskDtEntrega.Text) Then
            MsgBox "Data de Entrega Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
            mskDtEntrega.Text = "__/__/____"
            Cancel = True
        Else
            If CDate(mskDtEntrega.Text) < CDate(mskDTPEDIDO.Text) Then
                MsgBox "A data de Entrega não pode ser menor que a data do Pedido !!!", vbOKOnly + vbCritical, "Aviso"
                mskDtEntrega.Text = "__/__/____"
                Cancel = True
            Else
               txtPRZENTREGA.Text = (CDate(mskDtEntrega.Text) - CDate(mskDTPEDIDO.Text))
            End If
        End If
    End If
End Sub

Private Sub mskDTPEDIDO_GotFocus()
    objBLBFunc.SelecionaCampos mskDTPEDIDO.Name, frmCADCOTAVENDA
End Sub

Private Sub mskDtValidade_Validate(Cancel As Boolean)

    If mskDtValidade.Text <> "__/__/____" Then
        If Not IsDate(mskDtValidade.Text) Then
            MsgBox "Data de validade inválida !!!", vbOKOnly + vbExclamation, "Aviso"
            mskDtValidade.Text = "__/__/____"
            Cancel = True
        Else
            If CDate(mskDtValidade.Text) < CDate(mskDTPEDIDO.Text) Then
                MsgBox "A data de validade não pode ser menor que a data do Pedido !!!", vbOKOnly + vbCritical, "Aviso"
                mskDtValidade.Text = "__/__/____"
                Cancel = True
            Else
               txtVALORC.Text = (CDate(mskDtValidade.Text) - CDate(mskDTPEDIDO.Text))
            End If
        End If
    End If

End Sub

Private Sub stOrca_Click(PreviousTab As Integer)
    If PreviousTab = 0 Then
       If txtCodProd.Enabled = True Then txtCodProd.SetFocus
    End If
End Sub

Private Sub txtALIQICMS_GotFocus()
    objBLBFunc.SelecionaCampos txtALIQICMS.Name, frmCADCOTAVENDA
End Sub

Private Sub txtALIQICMS_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtALIQICMS.Text
End Sub

Private Sub txtALIQICMS_Validate(Cancel As Boolean)
    
    Dim ccurVLICMS As Currency
    
    lblVLICMS.Caption = ""
    
    If Len(Trim(txtALIQICMS.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtALIQICMS.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbExclamation, "Aviso"
       txtALIQICMS.Text = ""
       txtALIQICMS.SetFocus
       Cancel = True
       Exit Sub
    End If
    
    txtALIQICMS.Text = Format(txtALIQICMS.Text, "#,##0.00")
    
    CalcValores
    
End Sub

Private Sub txtCODCLI_GotFocus()
    objBLBFunc.SelecionaCampos txtCODCLI.Name, frmCADCOTAVENDA
End Sub

Private Sub txtCODCLI_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODCLI.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODCLI.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODCLI.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboCliente.ListIndex = -1
    For I = 0 To (cboCliente.ListCount - 1)
        If cboCliente.ItemData(I) = CInt(txtCODCLI.Text) Then cboCliente.ListIndex = I
    Next I
    
    If cboCliente.ListIndex = -1 Then
       MsgBox "Este cliente não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODCLI.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    objCADCOTAVENDA.CODCLI = txtCODCLI.Text
    objCADCOTAVENDA.PreencheComboContato cboContato
    objCADCOTAVENDA.PreencheComboEmail cboEMail
    objCADCOTAVENDA.PreencheComboTelefone cboTelefone

End Sub

Private Sub txtCODCONDPGT_GotFocus()
    objBLBFunc.SelecionaCampos txtCODCONDPGT.Name, frmCADCOTAVENDA
End Sub

Private Sub txtCODCONDPGT_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODCONDPGT.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODCONDPGT.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODCONDPGT.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboCONDPGTO.ListIndex = -1
    For I = 0 To (cboCONDPGTO.ListCount - 1)
        If cboCONDPGTO.ItemData(I) = CInt(txtCODCONDPGT.Text) Then cboCONDPGTO.ListIndex = I
    Next I
    
    If cboCONDPGTO.ListIndex = -1 Then
       MsgBox "Esta condição de pagamento não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODCONDPGT.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    ConfGridProd

    txtCodProd.Text = ""
    txtQtdCompra.Text = ""
    txtVlUnitarioCompra.Text = ""
    txtIPICompra.Text = ""
    txtVlTotalCompra.Text = ""
    txtPORCDESC.Text = ""

End Sub


Private Sub txtCodProd_GotFocus()
    objBLBFunc.SelecionaCampos txtCodProd.Name, frmCADCOTAVENDA
End Sub

Private Sub txtCodProd_Validate(Cancel As Boolean)

   Dim I As Integer

   If Len(Trim(txtCodProd.Text)) = 0 Then Exit Sub
   
   txtCodProd.Tag = PegaTagProduto(txtCodProd.Text)
   txtQtdCompra.SetFocus

End Sub


Private Sub txtCODTIPORC_GotFocus()
    objBLBFunc.SelecionaCampos txtCODTIPORC.Name, frmCADCOTAVENDA
End Sub

Private Sub txtCODTIPORC_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODTIPORC.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODTIPORC.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODTIPORC.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboTIPORC.ListIndex = -1
    For I = 0 To (cboTIPORC.ListCount - 1)
        If cboTIPORC.ItemData(I) = CInt(txtCODTIPORC.Text) Then cboTIPORC.ListIndex = I
    Next I
    
    If cboTIPORC.ListIndex = -1 Then
       MsgBox "Este tipo de orçamento não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODTIPORC.Text = ""
       Cancel = True
       Exit Sub
    End If
    
End Sub

Private Sub txtCODVEND_GotFocus()
    objBLBFunc.SelecionaCampos txtCODVEND.Name, frmCADCOTAVENDA
End Sub

Private Sub txtCODVEND_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODVEND.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODVEND.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODVEND.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboVendedor.ListIndex = -1
    For I = 0 To (cboVendedor.ListCount - 1)
        If cboVendedor.ItemData(I) = CInt(txtCODVEND.Text) Then cboVendedor.ListIndex = I
    Next I
    
    If cboVendedor.ListIndex = -1 Then
       MsgBox "Este vendedor não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODVEND.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtDepto_GotFocus()
    objBLBFunc.SelecionaCampos txtDepto.Name, frmCADCOTAVENDA
End Sub

Private Sub txtDepto_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtFRETE_GotFocus()
    objBLBFunc.SelecionaCampos txtFRETE.Name, frmCADCOTAVENDA
End Sub

Private Sub txtFRETE_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtFRETE.Text
End Sub

Private Sub txtFRETE_Validate(Cancel As Boolean)

    If Len(Trim(txtFRETE.Text)) = 0 Then
       CalcValores
       Exit Sub
    End If
    
    If Not IsNumeric(txtFRETE.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbExclamation, "Aviso"
       txtFRETE.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtFRETE.Text = Format(txtFRETE.Text, "#,##0.00")
    
    CalcValores

End Sub

Private Sub txtIPICompra_GotFocus()
    objBLBFunc.SelecionaCampos txtIPICompra.Name, frmCADCOTAVENDA
End Sub

Private Sub txtIPICompra_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtIPICompra.Text
End Sub

Private Sub txtIPICompra_Validate(Cancel As Boolean)

    If Len(Trim(txtIPICompra.Text)) = 0 Then
       SommaItens
       Exit Sub
    End If

    If Not IsNumeric(txtIPICompra.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       Cancel = True
       Exit Sub
    End If
    
    If Val(txtIPICompra.Text) < 0 Then
       MsgBox "Não é permitido numero negativo !!!", vbOKOnly + vbCritical, "aviso"
       txtIPICompra.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtIPICompra.Text = Format(txtIPICompra.Text, "#,##0.00")
    
    SommaItens

End Sub

Private Sub txtOutrDesp_GotFocus()
    objBLBFunc.SelecionaCampos txtOutrDesp.Name, frmCADCOTAVENDA
End Sub

Private Sub txtOutrDesp_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtALIQICMS.Text
End Sub

Private Sub txtOutrDesp_Validate(Cancel As Boolean)

    If Len(Trim(txtOutrDesp.Text)) = 0 Then
       Call CalcValores
       Exit Sub
    End If
    
    If Not IsNumeric(txtOutrDesp.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbExclamation, "Aviso"
       txtOutrDesp.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtOutrDesp.Text = Format(txtOutrDesp.Text, "#,##0.00")
    
    Call CalcValores

End Sub

Private Sub txtPDESCTOTAL_GotFocus()
    objBLBFunc.SelecionaCampos txtPDESCTOTAL.Name, frmCADCOTAVENDA
End Sub

Private Sub txtPDESCTOTAL_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtPDESCTOTAL.Text
End Sub

Private Sub txtPDESCTOTAL_Validate(Cancel As Boolean)

    Dim ccurVLDESCTO As Currency
    
    txtVLDESCTO.Text = 0
    
    If Len(Trim(txtPDESCTOTAL.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtPDESCTOTAL.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbExclamation, "Aviso"
       txtPDESCTOTAL.Text = ""
       txtPDESCTOTAL.SetFocus
       Cancel = True
       Exit Sub
    End If
    
    
    txtPDESCTOTAL.Text = Format(txtPDESCTOTAL.Text, "#,##0.00")
    
    CalcValores

End Sub

Private Sub txtPORCDESC_GotFocus()
    objBLBFunc.SelecionaCampos txtPORCDESC.Name, frmCADCOTAVENDA
End Sub

Private Sub txtPORCDESC_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtPORCDESC.Text
End Sub

Private Sub txtPORCDESC_Validate(Cancel As Boolean)

    If Len(Trim(txtPORCDESC.Text)) = 0 Then
       SommaItens
       Exit Sub
    End If

    If Not IsNumeric(txtPORCDESC.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       Cancel = True
       Exit Sub
    End If
    
    If Val(txtPORCDESC.Text) < 0 Then
       MsgBox "Não é permitido numero negativo !!!", vbOKOnly + vbCritical, "aviso"
       txtPORCDESC.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtPORCDESC.Text = Format(txtPORCDESC.Text, "#,##0.00")
    
    SommaItens

End Sub

Private Sub txtPRZENTREGA_GotFocus()
    objBLBFunc.SelecionaCampos txtPRZENTREGA.Name, frmCADCOTAVENDA
End Sub

Private Sub txtPRZENTREGA_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtPRZENTREGA_Validate(Cancel As Boolean)

    If Len(Trim(txtPRZENTREGA.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtPRZENTREGA.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtPRZENTREGA.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtQtdCompra_GotFocus()
    objBLBFunc.SelecionaCampos txtQtdCompra.Name, frmCADCOTAVENDA
End Sub

Private Sub txtQtdCompra_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtQtdCompra.Text
End Sub

Private Sub txtQtdCompra_Validate(Cancel As Boolean)

    If Len(Trim(txtQtdCompra.Text)) = 0 Then
       SommaItens
       Exit Sub
    End If

    If Not IsNumeric(txtQtdCompra.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       Cancel = True
       Exit Sub
    End If
    
    If Val(txtQtdCompra.Text) < 0 Then
       MsgBox "Não é permitido numero negativo !!!", vbOKOnly + vbCritical, "aviso"
       txtQtdCompra.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtQtdCompra.Text = Format(txtQtdCompra.Text, "#,###0.000")
    
    SommaItens

End Sub


Private Sub txtVALORC_GotFocus()
    objBLBFunc.SelecionaCampos txtVALORC.Name, frmCADCOTAVENDA
End Sub

Private Sub txtVlUnitarioCompra_GotFocus()
    objBLBFunc.SelecionaCampos txtVlUnitarioCompra.Name, frmCADCOTAVENDA
End Sub

Private Sub txtVlUnitarioCompra_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtVlUnitarioCompra.Text
End Sub

Private Sub txtVlUnitarioCompra_Validate(Cancel As Boolean)

    If Len(Trim(txtVlUnitarioCompra.Text)) = 0 Then
       SommaItens
       Exit Sub
    End If

    If Not IsNumeric(txtVlUnitarioCompra.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       Cancel = True
       Exit Sub
    End If
    
    If Val(txtVlUnitarioCompra.Text) < 0 Then
       MsgBox "Não é permitido numero negativo !!!", vbOKOnly + vbCritical, "aviso"
       txtVlUnitarioCompra.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtVlUnitarioCompra.Text = Format(txtVlUnitarioCompra.Text, "#,##0.00")
    
    SommaItens

End Sub

Private Sub SommaItens()

    Dim curQtde    As Currency
    Dim curVLTot   As Currency
    Dim curVLUNI   As Currency
    Dim curPORCIPI As Currency
    Dim curVLIPI   As Currency
    Dim curVLDESC  As Currency
    Dim curPORDESC As Currency
    Dim curTOTDESC As Currency
    Dim curVLITEM  As Currency
    
    If Len(Trim(txtQtdCompra.Text)) > 0 Then curQtde = CCur(txtQtdCompra.Text)
    If Len(Trim(txtVlUnitarioCompra.Text)) > 0 Then curVLUNI = CCur(txtVlUnitarioCompra.Text)
    If Len(Trim(txtIPICompra.Text)) > 0 Then curPORCIPI = CCur(txtIPICompra.Text)
    If Len(Trim(txtPORCDESC.Text)) > 0 Then curPORDESC = CCur(txtPORCDESC.Text)
    
    curTOTDESC = (curPORDESC * (curQtde * curVLUNI) / 100)
    curTOTDESC = (curTOTDESC * -1)
    
    curVLITEM = (curQtde * curVLUNI)
    curVLITEM = curVLITEM + curTOTDESC
    
    curVLIPI = ((curPORCIPI * curVLITEM) / 100)
    
    curVLTot = curVLITEM + curVLIPI
        
    txtVlTotalCompra.Text = Format(curVLTot, "#,##0.00")
    
End Sub


Private Sub IncProdGridItens()

    Dim I As Integer
    
    '' ---------------------------------------------------------
    If Len(Trim(txtQtdCompra.Text)) = 0 Then
       MsgBox "Informe a quantidade !!!", vbOKOnly + vbExclamation, "Aviso"
       txtQtdCompra.SetFocus
       Exit Sub
    End If
    If CCur(txtQtdCompra.Text) = 0 Then
       MsgBox "Informe a quantidade !!!", vbOKOnly + vbExclamation, "Aviso"
       txtQtdCompra.SetFocus
       Exit Sub
    End If
    
    If Len(Trim(txtVlUnitarioCompra.Text)) = 0 Then
       MsgBox "Informe o valor unitário !!!", vbOKOnly + vbExclamation, "Aviso"
       txtVlUnitarioCompra.SetFocus
       Exit Sub
    End If
    
    For I = 1 To (flxGridProd.Rows - 1)
        If Trim(flxGridProd.TextMatrix(I, 1)) = Trim(txtCodProd.Text) Then
           MsgBox "Este produto já esta relacionado !!!", vbOKOnly + vbExclamation, "aviso"
           txtCodProd.Text = ""
           lblDescProd.Caption = ""
           txtQtdCompra.Text = ""
           txtVlUnitarioCompra.Text = ""
           txtPORCDESC.Text = ""
           txtIPICompra.Text = ""
           txtVlTotalCompra.Text = ""
           txtCodProd.SetFocus
           Exit Sub
        End If
    Next I
    '' ---------------------------------------------------------
    
    flxGridProd.AddItem txtCodProd.Tag & vbTab & _
                        txtCodProd.Text & vbTab & _
                        Trim(lblDescProd.Caption) & vbTab & _
                        txtQtdCompra.Text & vbTab & _
                        txtVlUnitarioCompra.Text & vbTab & _
                        txtPORCDESC.Text & vbTab & _
                        txtIPICompra.Text & vbTab & _
                        txtVlTotalCompra.Text
                     
    
    
    Call CalcTotORc
    txtTotalItens.Text = Format(SomaGrdItens, "#,##0.00")
    
    txtCodProd.Text = ""
    txtQtdCompra.Text = ""
    txtVlUnitarioCompra.Text = ""
    txtIPICompra.Text = ""
    txtVlTotalCompra.Text = ""
    txtPORCDESC.Text = ""
    lblDescProd.Caption = ""
    
    txtCodProd.SetFocus
    
End Sub

Private Sub CalcTotORc()

    Dim I            As Integer
    Dim vlVbaseICMS  As Currency
    Dim VLIPI        As Currency
    Dim vlTotal      As Currency
    Dim curVLOUTDESP As Currency
    Dim curVLFRETE   As Currency
    Dim curVLDESC    As Currency
    Dim curVLITEM    As Currency
    Dim cutTOTDESC   As Currency
    Dim curPORICMS   As Currency
    Dim curTOTVLICMS As Currency
    Dim curVLSERV    As Currency
    Dim curTOTDESC   As Currency
    
    lblBASICMS.Caption = ""
    
    lblVLIPI.Caption = ""
    lblVLTOTAL.Caption = ""
    lblVLDESC.Caption = ""
    
    vlVbaseICMS = 0
    VLIPI = 0
    vlTotal = 0
    curVLOUTDESP = 0
    curVLFRETE = 0
    curVLDESC = 0
    cutTOTDESC = 0
    curVLSERV = 0
    curTOTDESC = 0
        
    For I = 1 To flxGridProd.Rows - 1
    
        curVLITEM = (CCur(flxGridProd.TextMatrix(I, 3)) * CCur(flxGridProd.TextMatrix(I, 4)))
        
        If Len(Trim(flxGridProd.TextMatrix(I, 5))) > 0 Then curVLDESC = ((CCur(flxGridProd.TextMatrix(I, 5)) * curVLITEM) / 100)
        
        vlVbaseICMS = vlVbaseICMS + curVLITEM
        
        If Len(Trim(txtALIQICMS.Text)) > 0 Then curPORICMS = ((CCur(txtALIQICMS.Text) * vlVbaseICMS) / 100)
        If Len(Trim(flxGridProd.TextMatrix(I, 6))) > 0 Then VLIPI = VLIPI + ((CCur(flxGridProd.TextMatrix(I, 6)) * (curVLITEM - curVLDESC)) / 100)
        
        vlTotal = vlTotal + CCur(flxGridProd.TextMatrix(I, 7))
        
        cutTOTDESC = cutTOTDESC + curVLDESC
        curTOTVLICMS = curTOTVLICMS + curPORICMS
    Next I
    
    If vlVbaseICMS > 0 Then lblBASICMS.Caption = Format(vlVbaseICMS, "#,##0.00")
    
    If VLIPI > 0 Then lblVLIPI.Caption = Format(VLIPI, "#,##0.00")
    If Len(Trim(txtOutrDesp.Text)) > 0 Then curVLOUTDESP = CCur(txtOutrDesp.Text)
    If Len(Trim(txtFRETE.Text)) > 0 Then curVLFRETE = CCur(txtFRETE.Text)
    If cutTOTDESC > 0 Then lblVLDESC.Caption = Format(cutTOTDESC, "#,##0.00")
    If curPORICMS > 0 Then lblVLICMS.Caption = Format(curPORICMS, "#,##0.00")
    
    If Len(Trim(txtPDESCTOTAL.Text)) > 0 Then curTOTDESC = ((CCur(txtPDESCTOTAL.Text) * (vlVbaseICMS + VLIPI + curVLOUTDESP + curVLSERV)) / 100)
    If curTOTDESC > 0 Then txtVLDESCTO.Text = Format(curTOTDESC, "#,##0.00")
    
    lblVLTOTAL.Caption = Format((((vlTotal + curVLOUTDESP + curVLSERV) - curTOTDESC)) + curVLFRETE, "#,##0.00")

End Sub

Private Sub CalcValores()

    Dim curVLBASE    As Currency
    Dim curVLIPI     As Currency
    Dim curVLOUTDESP As Currency
    Dim curVLFRETE   As Currency
    Dim curVLTOTAL   As Currency
    Dim ccurVLICMS   As Currency
    Dim ccurDESCONTO As Currency
    Dim ccurVLSERV   As Currency
    
    curVLBASE = 0
    ccurVLICMS = 0
    curVLIPI = 0
    curVLOUTDESP = 0
    curVLFRETE = 0
    ccurDESCONTO = 0
    ccurVLSERV = 0
        
    If Len(Trim(lblBASICMS.Caption)) > 0 Then curVLBASE = CCur(lblBASICMS.Caption)
    If curVLBASE > 0 Then
       If Len(Trim(txtALIQICMS.Text)) > 0 Then ccurVLICMS = (CCur(txtALIQICMS.Text) * CCur(curVLBASE) / 100)
    End If
    If ccurVLICMS > 0 Then
       lblVLICMS.Caption = Format(ccurVLICMS, "#,##0.00")
    End If
    
    If Len(Trim(txtOutrDesp.Text)) > 0 Then curVLOUTDESP = CCur(txtOutrDesp.Text)
    If Len(Trim(lblVLIPI.Caption)) > 0 Then curVLIPI = CCur(lblVLIPI.Caption)
    If Len(Trim(txtFRETE.Text)) > 0 Then curVLFRETE = CCur(txtFRETE.Text)
    
    If Len(Trim(txtPDESCTOTAL.Text)) > 0 Then ccurDESCONTO = (CCur(txtPDESCTOTAL.Text) * (curVLBASE + curVLOUTDESP + ccurVLSERV + curVLIPI) / 100)
    If ccurDESCONTO > 0 Then
       txtVLDESCTO.Text = Format(ccurDESCONTO, "#,##0.00")
    End If
    
    curVLTOTAL = (((curVLBASE + curVLOUTDESP + ccurVLSERV + curVLIPI) - ccurDESCONTO) + curVLFRETE)
    
    lblVLTOTAL.Caption = Format(curVLTOTAL, "#,##0.00")
    
    
End Sub

Private Function PegaPreco(strCODPROD As String) As Double

    PegaPreco = 0
    
    If Len(Trim(txtCODCONDPGT.Text)) = 0 Then Exit Function
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_TABPRECO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODPROD = '" & strCODPROD & "'" & vbCrLf
    sSql = sSql & "   And SGI_CODPGTO = " & txtCODCONDPGT.Text & vbCrLf
    sSql = sSql & "   And SGI_VIGENTE = 'S' "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC.EOF Then
       PegaPreco = BREC!SGI_VLVENDA
       BREC.Close
    Else
       BREC.Close
       
       If Len(Trim(strCODPROD)) > 0 Then
       
             sSql = "Select" & vbCrLf
             sSql = sSql & "       * " & vbCrLf
             sSql = sSql & "  From " & vbCrLf
             sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
             sSql = sSql & " Where " & vbCrLf
             sSql = sSql & "       SGI_FILIAL     = " & FILIAL & vbCrLf
             sSql = sSql & "   And SGI_IDPRODUTO  = " & strCODPROD
            
             BREC.Open sSql, adoBanco_Dados, adOpenDynamic
             If Not BREC.EOF Then
                If Not IsNull(BREC!SGI_PRECOPROD) Then PegaPreco = BREC!SGI_PRECOPROD
             End If
             BREC.Close
             
        End If
    End If
    
End Function

Private Function ValidaCampos() As Boolean

    ValidaCampos = False
    
    If Not IsDate(mskDTPEDIDO.Text) Then
       MsgBox "Data do orçamento inválida !!!", vbOKOnly + vbExclamation, "Aviso"
       stOrca.Tab = 0
       mskDTPEDIDO.SetFocus
       Exit Function
    End If
    If Len(Trim(txtCODCLI.Text)) = 0 Then
       MsgBox "Informe o código do cliente !!!", vbOKOnly + vbExclamation, "Aviso"
       stOrca.Tab = 0
       txtCODCLI.SetFocus
       Exit Function
    End If
    If Len(Trim(txtCODVEND.Text)) = 0 Then
       MsgBox "Informe o código do vendedor !!!", vbOKOnly + vbExclamation, "Aviso"
       stOrca.Tab = 0
       txtCODVEND.SetFocus
       Exit Function
    End If
    If Len(Trim(txtCODCONDPGT.Text)) = 0 Then
       MsgBox "Informe o código da condição de pagamento !!!", vbOKOnly + vbExclamation, "Aviso"
       stOrca.Tab = 0
       txtCODCONDPGT.SetFocus
       Exit Function
    End If
    If Len(Trim(txtCODTIPORC.Text)) = 0 Then
       MsgBox "Informe o código do tipo de orçamento  !!!", vbOKOnly + vbExclamation, "Aviso"
       stOrca.Tab = 0
       txtCODTIPORC.SetFocus
       Exit Function
    End If
    If flxGridProd.Rows = 1 Then
       MsgBox "Informe itens para o orçamento  !!!", vbOKOnly + vbExclamation, "Aviso"
       stOrca.Tab = 1
       txtCodProd.SetFocus
       Exit Function
    End If
    If Len(Trim(txtVALORC.Text)) = 0 Then
       MsgBox "Informe a validade do orçamento  !!!", vbOKOnly + vbExclamation, "Aviso"
       mskDtValidade.SetFocus
       Exit Function
    End If
    If Len(Trim(txtPRZENTREGA.Text)) = 0 Then
       MsgBox "Informe o Prazo de Entrega do orçamento !!!", vbOKOnly + vbExclamation, "Aviso"
       mskDtEntrega.SetFocus
       Exit Function
    End If
    
    ValidaCampos = True

End Function

Private Sub Consulta()

    Dim I          As Integer
    Dim j          As Integer
    Dim arrTIPSERV As Variant
    Dim curQTDORCA As Currency
    Dim curQTDPED  As Currency
    Dim curSALDO   As Currency
    Dim strSTATUS  As String
    
    stOrca.Tab = 0
    ReDim arrTIPOSSERV(1 To 100, 1 To 100, 1 To 6) As String
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    
    fraPedidos.Visible = False '' Mostra Pedidos Gerados
    
    Frame2.Enabled = False
    Frame3.Enabled = False
    Frame8.Enabled = False
    Frame9.Enabled = False

    Me.Caption = "Cadastro de orçamentos - [ CONSULTA ]"
    
    
    objBLBFunc.LimpaCampos frmCADCOTAVENDA
    
    lblCODIGO.Caption = ""
    lblBASICMS.Caption = ""
    lblVLICMS.Caption = ""
    lblVLIPI.Caption = ""
    txtOutrDesp.Text = ""
    lblVLDESC.Caption = ""
    lblVLTOTAL.Caption = ""
    lblDescProd.Caption = ""
    
    mskDtEntrega.Text = "__/__/____"
    mskDtValidade.Text = "__/__/____"
    
    ConfGridProd
    
    objCADCOTAVENDA.CODIGO = iCodigo
    
    objCADCOTAVENDA.PreencheComboCliente cboCliente
    objCADCOTAVENDA.PreencheComboVendedor cboVendedor, strUsuario
    objCADCOTAVENDA.PreencheComboCondPgto cboCONDPGTO
    objCADCOTAVENDA.PreencheComboTipoOrc cboTIPORC
    
    If objCADCOTAVENDA.Carrega_Campos = True Then
    
       If objCADCOTAVENDA.STATUS = "A" Then
          lblSTATUS.Caption = "ABERTO"
          lblSTATUS.ForeColor = &HFF&
       ElseIf objCADCOTAVENDA.STATUS = "B" Then
          lblSTATUS.Caption = "BAIXADO"
          lblSTATUS.ForeColor = &H8000&
       ElseIf objCADCOTAVENDA.STATUS = "P" Then
          lblSTATUS.Caption = "PARCIAL"
          lblSTATUS.ForeColor = &H8080&
       End If
       
       mskDtValidade.Text = Format(objCADCOTAVENDA.DATVALCOTA, "DD/MM/YYYY")
       mskDtEntrega.Text = Format(objCADCOTAVENDA.DATENTREGA, "DD/MM/YYYY")
       
       lblCODIGO.Caption = Mid(Trim(objCADCOTAVENDA.CODIGO), 1, (Len(Trim(objCADCOTAVENDA.CODIGO)) - 4)) & "/" & Right(objCADCOTAVENDA.CODIGO, 4)
       mskDTPEDIDO.Text = Format(objCADCOTAVENDA.DTCOTACAO, "DD/MM/YYYY")
       
       txtCODCLI.Text = objCADCOTAVENDA.CODCLI
       For I = 0 To (cboCliente.ListCount - 1)
           If cboCliente.ItemData(I) = CInt(txtCODCLI.Text) Then
              cboCliente.ListIndex = I
              Exit For
           End If
       Next I
       
       objCADCOTAVENDA.PreencheComboContato cboContato
       
       cboContato.Text = objCADCOTAVENDA.CONTATO
       txtDepto.Text = objCADCOTAVENDA.DEPTO
       cboEMail.Text = objCADCOTAVENDA.EMAIL
       cboTelefone.Text = objCADCOTAVENDA.TELCLIE
       
       txtCODVEND.Text = objCADCOTAVENDA.CODVEND
       For I = 0 To (cboVendedor.ListCount - 1)
           If cboVendedor.ItemData(I) = CInt(txtCODVEND.Text) Then
              cboVendedor.ListIndex = I
              Exit For
           End If
       Next I
       
       txtCODCONDPGT.Text = objCADCOTAVENDA.CODCONDPGT
       For I = 0 To (cboCONDPGTO.ListCount - 1)
           If cboCONDPGTO.ItemData(I) = CInt(txtCODCONDPGT.Text) Then
              cboCONDPGTO.ListIndex = I
              Exit For
           End If
       Next I
       
       txtCODTIPORC.Text = objCADCOTAVENDA.CODTIPORC
       For I = 0 To (cboTIPORC.ListCount - 1)
           If cboTIPORC.ItemData(I) = CInt(txtCODTIPORC.Text) Then
              cboTIPORC.ListIndex = I
              Exit For
           End If
       Next I
       
       txtPRZENTREGA.Text = objCADCOTAVENDA.PRZENTREGA
       txtVALORC.Text = objCADCOTAVENDA.VALCOTACAO
       
       arrPRODUTOS = objCADCOTAVENDA.Produtos
       Call PopGrdItensCota
       
       '' Soma Total do Itens
       txtTotalItens.Text = Format(SomaGrdItens, "#,##0.00")
       
       If MostraSaldos = True Then
          flxGridProd.ColWidth(8) = 1000
          flxGridProd.ColWidth(9) = 1000
          flxGridProd.ColWidth(10) = 1000
       End If
       
       If objCADCOTAVENDA.BASEICMS > 0 Then lblBASICMS.Caption = Format(objCADCOTAVENDA.BASEICMS, "#,##0.00")
       If objCADCOTAVENDA.ALIQICMS > 0 Then txtALIQICMS.Text = Format(objCADCOTAVENDA.ALIQICMS, "#,##0.00")
       If objCADCOTAVENDA.VLICMS > 0 Then lblVLICMS.Caption = Format(objCADCOTAVENDA.VLICMS, "#,##0.00")
       If objCADCOTAVENDA.TOTOUTDESP > 0 Then txtOutrDesp.Text = Format(objCADCOTAVENDA.TOTOUTDESP, "#,##0.00")
       If objCADCOTAVENDA.TOTFRETE > 0 Then txtFRETE.Text = Format(objCADCOTAVENDA.TOTFRETE, "#,##0.00")
       If objCADCOTAVENDA.VLIPI > 0 Then lblVLIPI.Caption = Format(objCADCOTAVENDA.VLIPI, "#,##0.00")
       If objCADCOTAVENDA.VLDESCTO > 0 Then lblVLDESC.Caption = Format(objCADCOTAVENDA.VLDESCTO, "#,##0.00")
      
       If objCADCOTAVENDA.VLPDESCT > 0 Then txtPDESCTOTAL.Text = Format(objCADCOTAVENDA.VLPDESCT, "#,##0.00")
       If objCADCOTAVENDA.VLTOTDESC > 0 Then txtVLDESCTO.Text = Format(objCADCOTAVENDA.VLTOTDESC, "#,##0.00")
       
       lblVLTOTAL.Caption = Format(objCADCOTAVENDA.TOTORCTO, "#,##0.00")
       
    End If

End Sub


Private Sub Altera()

    Dim I          As Integer
    Dim j          As Integer
    Dim arrTIPSERV As Variant
    
    stOrca.Tab = 0
    ReDim arrTIPOSSERV(1 To 100, 1 To 100, 1 To 2) As String
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    
    fraPedidos.Visible = False '' Mostra Pedidos Gerados
    
    Frame2.Enabled = True
    Frame3.Enabled = True
    Frame8.Enabled = True
    Frame9.Enabled = True

    Me.Caption = "Cadastro de orçamentos - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADCOTAVENDA
    
    lblCODIGO.Caption = ""
    lblBASICMS.Caption = ""
    lblVLICMS.Caption = ""
    lblVLIPI.Caption = ""
    txtOutrDesp.Text = ""
    lblVLDESC.Caption = ""
    lblVLTOTAL.Caption = ""
    lblDescProd.Caption = ""
    
    mskDtEntrega.Text = "__/__/____"
    mskDtValidade.Text = "__/__/____"
    
    ConfGridProd
    
    objCADCOTAVENDA.CODIGO = iCodigo
    
    objCADCOTAVENDA.PreencheComboCliente cboCliente
    objCADCOTAVENDA.PreencheComboVendedor cboVendedor, strUsuario
    objCADCOTAVENDA.PreencheComboCondPgto cboCONDPGTO
    objCADCOTAVENDA.PreencheComboTipoOrc cboTIPORC
    
    If objCADCOTAVENDA.Carrega_Campos = True Then
    
       If objCADCOTAVENDA.STATUS = "A" Then
          lblSTATUS.Caption = "ABERTO"
          lblSTATUS.ForeColor = &HFF&
       ElseIf objCADCOTAVENDA.STATUS = "B" Then
          lblSTATUS.Caption = "BAIXADO"
          lblSTATUS.ForeColor = &H8000&
       ElseIf objCADCOTAVENDA.STATUS = "P" Then
          lblSTATUS.Caption = "PARCIAL"
          lblSTATUS.ForeColor = &H8080&
       End If
       
       mskDtValidade.Text = Format(objCADCOTAVENDA.DATVALCOTA, "DD/MM/YYYY")
       mskDtEntrega.Text = Format(objCADCOTAVENDA.DATENTREGA, "DD/MM/YYYY")
       
       lblCODIGO.Caption = Mid(Trim(objCADCOTAVENDA.CODIGO), 1, (Len(Trim(objCADCOTAVENDA.CODIGO)) - 4)) & "/" & Right(objCADCOTAVENDA.CODIGO, 4)
       mskDTPEDIDO.Text = Format(objCADCOTAVENDA.DTCOTACAO, "DD/MM/YYYY")
       
       txtCODCLI.Text = objCADCOTAVENDA.CODCLI
       For I = 0 To (cboCliente.ListCount - 1)
           If cboCliente.ItemData(I) = CInt(txtCODCLI.Text) Then
              cboCliente.ListIndex = I
              Exit For
           End If
       Next I
       
       objCADCOTAVENDA.PreencheComboContato cboContato
       
       cboContato.Text = objCADCOTAVENDA.CONTATO
       txtDepto.Text = objCADCOTAVENDA.DEPTO
       cboEMail.Text = objCADCOTAVENDA.EMAIL
       cboTelefone.Text = objCADCOTAVENDA.TELCLIE
       
       txtCODVEND.Text = objCADCOTAVENDA.CODVEND
       For I = 0 To (cboVendedor.ListCount - 1)
           If cboVendedor.ItemData(I) = CInt(txtCODVEND.Text) Then
              cboVendedor.ListIndex = I
              Exit For
           End If
       Next I
       
       txtCODCONDPGT.Text = objCADCOTAVENDA.CODCONDPGT
       For I = 0 To (cboCONDPGTO.ListCount - 1)
           If cboCONDPGTO.ItemData(I) = CInt(txtCODCONDPGT.Text) Then
              cboCONDPGTO.ListIndex = I
              Exit For
           End If
       Next I
      
       txtCODTIPORC.Text = objCADCOTAVENDA.CODTIPORC
       For I = 0 To (cboTIPORC.ListCount - 1)
           If cboTIPORC.ItemData(I) = CInt(txtCODTIPORC.Text) Then
              cboTIPORC.ListIndex = I
              Exit For
           End If
       Next I
       
       txtPRZENTREGA.Text = objCADCOTAVENDA.PRZENTREGA
       txtVALORC.Text = objCADCOTAVENDA.VALCOTACAO
       
       '' Tipo de Serviços
       arrTIPOSSERV = objCADCOTAVENDA.TIPSERV
       '' -----------------------------
       
       arrPRODUTOS = objCADCOTAVENDA.Produtos
       Call PopGrdItensCota
       
       '' Soma Total do Itens
       txtTotalItens.Text = Format(SomaGrdItens, "#,##0.00")
       
       If objCADCOTAVENDA.BASEICMS > 0 Then lblBASICMS.Caption = Format(objCADCOTAVENDA.BASEICMS, "#,##0.00")
       If objCADCOTAVENDA.ALIQICMS > 0 Then txtALIQICMS.Text = Format(objCADCOTAVENDA.ALIQICMS, "#,##0.00")
       If objCADCOTAVENDA.VLICMS > 0 Then lblVLICMS.Caption = Format(objCADCOTAVENDA.VLICMS, "#,##0.00")
       If objCADCOTAVENDA.TOTOUTDESP > 0 Then txtOutrDesp.Text = Format(objCADCOTAVENDA.TOTOUTDESP, "#,##0.00")
       If objCADCOTAVENDA.TOTFRETE > 0 Then txtFRETE.Text = Format(objCADCOTAVENDA.TOTFRETE, "#,##0.00")
       If objCADCOTAVENDA.VLIPI > 0 Then lblVLIPI.Caption = Format(objCADCOTAVENDA.VLIPI, "#,##0.00")
       If objCADCOTAVENDA.VLDESCTO > 0 Then lblVLDESC.Caption = Format(objCADCOTAVENDA.VLDESCTO, "#,##0.00")
       lblVLTOTAL.Caption = Format(objCADCOTAVENDA.TOTORCTO, "#,##0.00")
       
    End If

End Sub

Private Sub IncRegGrid(intQtdeCampos As Integer, gGrid As MSFlexGrid)

    Dim I          As Integer
    Dim j          As Integer
    Dim strCampos  As String
    Dim intTotNull As Integer
    
    If gGrid.Rows > 1 Then
       For I = 1 To (gGrid.Rows - 1)
           intTotNull = 0
           For j = 1 To (gGrid.Cols - 1)
               If Len(Trim(gGrid.TextMatrix(I, j))) = 0 Then intTotNull = intTotNull + 1
           Next j
           If (gGrid.Cols - 1) = intTotNull Then Exit Sub
       Next I
    End If
    
    For I = 1 To intQtdeCampos
        strCampos = strCampos & ""
        If I < intQtdeCampos Then strCampos = strCampos & vbTab
    Next I
    
    gGrid.AddItem strCampos

End Sub



Private Function MostraSaldos() As Boolean

    Dim I As Integer
    
    MostraSaldos = False
    
    For I = 1 To (flxGridProd.Rows - 1)
        If Len(Trim(flxGridProd.TextMatrix(I, 8))) > 0 Then
           If CCur(flxGridProd.TextMatrix(I, 8)) > 0 Then MostraSaldos = True
        End If
    Next I
    
End Function

Private Function VerifPed() As Boolean

    VerifPed = False
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPEDVENDH " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODCOTA = " & objCADCOTAVENDA.CODIGO
       
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC2.EOF Then VerifPed = True
    BREC2.Close
    '' ---------------------------

End Function

Private Function SomaGrdItens() As Currency
    SomaGrdItens = 0
    
    Dim I As Integer
    
    For I = 1 To (flxGridProd.Rows - 1)
        SomaGrdItens = SomaGrdItens + CCur(flxGridProd.TextMatrix(I, 7))
    Next I
    
End Function

Private Function CalcVlTotOrca(vlBase As Currency) As Currency

End Function

Private Sub Pega_Vendedor(lngCodUsuario As Long)

    Dim I As Integer
    
    For I = 0 To (cboVendedor.ListCount - 1)
        If Trim(cboVendedor.ItemData(I)) = lngCodUsuario Then
           cboVendedor.ListIndex = I
           txtCODVEND.Text = Str(cboVendedor.ItemData(I))
           
           txtCODVEND.Enabled = False
           cboVendedor.Enabled = False
           Command1.Enabled = False
           
           Exit For
        End If
    Next I
       
    '' ===========================================
    '' Pega os Tipos de Orcamentos para o Vendedor
    If Len(Trim(txtCODVEND.Text)) > 0 Then
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       ESPO.* " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_VENDTIPORCA VEND " & vbCrLf
        sSql = sSql & "      ,SGI_CADESPORCA  ESPO " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       VEND.SGI_FILIAL  = " & FILIAL & vbCrLf
        sSql = sSql & "   And VEND.SGI_CODVEND = " & Trim(txtCODVEND.Text) & vbCrLf
        sSql = sSql & "   And ESPO.SGI_FILIAL  = VEND.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And ESPO.SGI_CODIGO  = VEND.SGI_CODTIPORCA "
        
        BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC2.EOF() Then
           cboTIPORC.Clear
           Do While Not BREC2.EOF()
              cboTIPORC.AddItem Trim(BREC2!SGI_DESCRICAO)
              cboTIPORC.ItemData(cboTIPORC.NewIndex) = BREC2!SGI_CODIGO
              BREC2.MoveNext
           Loop
        
            If cboTIPORC.ListCount = 1 Then
               cboTIPORC.ListIndex = 0
               txtCODTIPORC.Text = cboTIPORC.ItemData(0)
               cboTIPORC.Enabled = False
               txtCODTIPORC.Enabled = False
               Command3.Enabled = False
            End If
        
        
        End If
        BREC2.Close
        
        '' ===========================================
        '' Pega os Clientes
        sSql = "Select" & vbCrLf
        sSql = sSql & "       CLIE.*" & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_CADCLIEVEND VEND" & vbCrLf
        sSql = sSql & "      ,SGI_CADCLIENTE  CLIE" & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "       VEND.SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And VEND.SGI_CODIGO = " & Trim(txtCODVEND.Text) & vbCrLf
        sSql = sSql & "   And CLIE.SGI_FILIAL = VEND.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And CLIE.SGI_CODIGO = VEND.SGI_CODCLI"
        
        BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC2.EOF() Then
           cboCliente.Clear
           Do While Not BREC2.EOF()
              cboCliente.AddItem Trim(BREC2!SGI_RAZAOSOC)
              cboCliente.ItemData(cboCliente.NewIndex) = BREC2!SGI_CODIGO
              BREC2.MoveNext
           Loop
        End If
        BREC2.Close
        '' ===========================================
        
    End If
    '' ===========================================
       
    
End Sub


Private Function PegaTagProduto(strPRODUTO As String) As String
    
    PegaTagProduto = ""
    lblDescProd.Caption = ""
    
    sSql = ""
    
    sSql = "Select SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & ",Case PRO.SGI_PRODUTOTIPO " & vbCrLf
    sSql = sSql & "            When 1 Then replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD))) + '.' + " & vbCrLf
    sSql = sSql & "                        replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE))) + '.' + " & vbCrLf
    sSql = sSql & "                        replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO))) + '.' + " & vbCrLf
    sSql = sSql & "                        (Case " & vbCrLf
    sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Null Then '0' " & vbCrLf
    sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(1),PRO.SGI_DIGVERIF))) End) " & vbCrLf
    sSql = sSql & "            When 0 Then PRO.SGI_CODIGO End As SGI_CODIGO " & vbCrLf
    sSql = sSql & ",SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "         SGI_CADPRODUTO PRO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = 1 " & vbCrLf
    sSql = sSql & "   And (Case PRO.SGI_PRODUTOTIPO " & vbCrLf
    sSql = sSql & "            When 1 Then replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD))) + '.' + " & vbCrLf
    sSql = sSql & "                        replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE))) + '.' + " & vbCrLf
    sSql = sSql & "                        replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO))) + '.' + " & vbCrLf
    sSql = sSql & "                        (Case " & vbCrLf
    sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Null Then '0' " & vbCrLf
    sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(1),PRO.SGI_DIGVERIF))) End) " & vbCrLf
    sSql = sSql & "            When 0 Then PRO.SGI_CODIGO End) = '" & Trim(strPRODUTO) & "'"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        PegaTagProduto = Trim(Str(BREC!SGI_IDPRODUTO))
        lblDescProd.Caption = Trim(BREC!SGI_DESCRICAO)
    Else
        MsgBox "Produto Inexistente !!!", vbOKOnly + vbExclamation, "Aviso de Sistema"
    End If
    BREC.Close
    
    txtVlUnitarioCompra.Text = Format(PegaPreco(PegaTagProduto), "#,##0.00")
    
End Function


Private Sub PopGrdItensCota()

        Dim I As Integer
        Dim j As Integer
        Dim curQTDORCA As Currency
        Dim curQTDPED  As Currency
        Dim curSALDO   As Currency
        Dim strSTATUS  As String
       
        If Not IsArray(arrPRODUTOS) Then Exit Sub
       
       For I = 1 To UBound(arrPRODUTOS)
       
            sSql = "Select " & vbCrLf
            sSql = sSql & "SGI_IDPRODUTO" & vbCrLf
            sSql = sSql & ",Case PRO.SGI_PRODUTOTIPO " & vbCrLf
            sSql = sSql & "            When 1 Then replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODLINPROD))) + '.' + " & vbCrLf
            sSql = sSql & "                        replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODCLIE))) + '.' + " & vbCrLf
            sSql = sSql & "                        replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRO.SGI_CODROTULO))) + '.' + " & vbCrLf
            sSql = sSql & "                        (Case " & vbCrLf
            sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Null Then '0' " & vbCrLf
            sSql = sSql & "                              When PRO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(1),PRO.SGI_DIGVERIF))) End) " & vbCrLf
            sSql = sSql & "            When 0 Then PRO.SGI_CODIGO End As SGI_CODIGO " & vbCrLf
            sSql = sSql & ",SGI_DESCRICAO " & vbCrLf
            
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "         SGI_CADPRODUTO PRO " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       SGI_FILIAL     = " & FILIAL & vbCrLf
            sSql = sSql & "   And SGI_IDPRODUTO  = " & arrPRODUTOS(I, 7)
           
           BREC.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC.EOF Then
           
              curQTDORCA = arrPRODUTOS(I, 2)
              curQTDPED = CCur(arrPRODUTOS(I, 8))
              curSALDO = curQTDORCA - curQTDPED
              
              If curSALDO = 0 Then strSTATUS = "Atendido"
              If curSALDO > 0 Then
                 If curQTDPED = 0 Then
                    strSTATUS = "Aberto"
                 ElseIf curQTDORCA > curQTDPED Then
                    strSTATUS = "Parcial"
                 End If
              End If
              
              flxGridProd.AddItem BREC!SGI_IDPRODUTO & vbTab & _
                                  BREC!SGI_CODIGO & vbTab & _
                                  BREC!SGI_DESCRICAO & vbTab & _
                                  Format(arrPRODUTOS(I, 2), "#,###0.000") & vbTab & _
                                  Format(arrPRODUTOS(I, 3), "#,##0.00") & vbTab & _
                                  Format(arrPRODUTOS(I, 4), "#,##0.00") & vbTab & _
                                  Format(arrPRODUTOS(I, 5), "#,##0.00") & vbTab & _
                                  Format(arrPRODUTOS(I, 6), "#,##0.00") & vbTab & _
                                  Format(arrPRODUTOS(I, 8), "#,###0.000") & vbTab & _
                                  Format(curSALDO, "#,###0.000") & vbTab & _
                                  strSTATUS
                                  
              For j = 1 To (flxGridProd.Cols - 1)
                  flxGridProd.Row = (flxGridProd.Rows - 1)
                  flxGridProd.Col = j
                  If strSTATUS = "Aberto" Then
                     flxGridProd.CellForeColor = &HFF&
                  ElseIf strSTATUS = "Atendido" Then
                     flxGridProd.CellForeColor = &H8000&
                  ElseIf strSTATUS = "Parcial" Then
                     flxGridProd.CellForeColor = &H8080&
                  End If
              Next j
                                  
           End If
           BREC.Close
       
       Next I

End Sub
