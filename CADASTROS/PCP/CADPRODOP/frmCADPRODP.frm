VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCADPROD 
   Caption         =   "Cadastro de Produto"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15630
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   15630
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   96
      Top             =   0
      Width           =   15615
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
         Picture         =   "frmCADPRODP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   120
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
         Picture         =   "frmCADPRODP.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   120
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
         Picture         =   "frmCADPRODP.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7695
      Left            =   0
      TabIndex        =   51
      Top             =   840
      Width           =   15615
      Begin TabDlg.SSTab stProd 
         Height          =   7455
         Left            =   120
         TabIndex        =   52
         Top             =   120
         Width           =   15375
         _ExtentX        =   27120
         _ExtentY        =   13150
         _Version        =   393216
         Style           =   1
         Tabs            =   11
         Tab             =   10
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
         TabCaption(0)   =   "Dados do Produto"
         TabPicture(0)   =   "frmCADPRODP.frx":0736
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame3"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Ficha Técnica"
         TabPicture(1)   =   "frmCADPRODP.frx":0752
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "stFichaTec"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Localização Fisica"
         TabPicture(2)   =   "frmCADPRODP.frx":076E
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame16"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Descrição do Produto"
         TabPicture(3)   =   "frmCADPRODP.frx":078A
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame12"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Produção"
         TabPicture(4)   =   "frmCADPRODP.frx":07A6
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Command31"
         Tab(4).Control(1)=   "Command30"
         Tab(4).Control(2)=   "grdFamMaq"
         Tab(4).ControlCount=   3
         TabCaption(5)   =   "Unidades de Conversão"
         TabPicture(5)   =   "frmCADPRODP.frx":07C2
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "cmdexcUnidConv"
         Tab(5).Control(1)=   "cmdIncUnidConv"
         Tab(5).Control(2)=   "grdUnidConv"
         Tab(5).ControlCount=   3
         TabCaption(6)   =   "Desenho"
         TabPicture(6)   =   "frmCADPRODP.frx":07DE
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "Frame43(3)"
         Tab(6).ControlCount=   1
         TabCaption(7)   =   "Estoques"
         TabPicture(7)   =   "frmCADPRODP.frx":07FA
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "SSTab1"
         Tab(7).ControlCount=   1
         TabCaption(8)   =   "Processo - Produtivo"
         TabPicture(8)   =   "frmCADPRODP.frx":0816
         Tab(8).ControlEnabled=   0   'False
         Tab(8).Control(0)=   "Frame9"
         Tab(8).ControlCount=   1
         TabCaption(9)   =   "Dados Inerente a Produção"
         TabPicture(9)   =   "frmCADPRODP.frx":0832
         Tab(9).ControlEnabled=   0   'False
         Tab(9).ControlCount=   0
         TabCaption(10)  =   "Lista de Isumos"
         TabPicture(10)  =   "frmCADPRODP.frx":084E
         Tab(10).ControlEnabled=   -1  'True
         Tab(10).Control(0)=   "Frame19"
         Tab(10).Control(0).Enabled=   0   'False
         Tab(10).ControlCount=   1
         Begin VB.Frame Frame19 
            Height          =   6735
            Left            =   0
            TabIndex        =   257
            Top             =   600
            Width           =   15015
            Begin ComctlLib.TreeView treListaMat 
               Height          =   5295
               Left            =   120
               TabIndex        =   269
               Top             =   240
               Width           =   14175
               _ExtentX        =   25003
               _ExtentY        =   9340
               _Version        =   327682
               Style           =   7
               Appearance      =   1
            End
            Begin VB.CommandButton Command9 
               Height          =   300
               Left            =   14520
               Picture         =   "frmCADPRODP.frx":086A
               Style           =   1  'Graphical
               TabIndex        =   268
               Top             =   6000
               Width           =   300
            End
            Begin VB.CommandButton Command10 
               Height          =   300
               Left            =   14520
               Picture         =   "frmCADPRODP.frx":09B4
               Style           =   1  'Graphical
               TabIndex        =   267
               Top             =   5640
               Width           =   300
            End
            Begin VB.Frame fraDadosList 
               Caption         =   "[ Material ]"
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
               Height          =   975
               Left            =   120
               TabIndex        =   258
               Top             =   5640
               Width           =   14295
               Begin VB.TextBox txtQtdeIns 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  Height          =   315
                  Left            =   4800
                  MaxLength       =   10
                  TabIndex        =   259
                  Text            =   "txtQtdeIns"
                  Top             =   600
                  Width           =   1335
               End
               Begin VB.ComboBox cboUniIns 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Left            =   2040
                  TabIndex        =   266
                  Text            =   "cboUnidade"
                  Top             =   600
                  Width           =   855
               End
               Begin VB.CommandButton Command11 
                  Height          =   315
                  Left            =   3360
                  Picture         =   "frmCADPRODP.frx":0AFE
                  Style           =   1  'Graphical
                  TabIndex        =   262
                  Top             =   240
                  Width           =   375
               End
               Begin VB.TextBox txtCODPROD 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  Height          =   315
                  Left            =   2040
                  MaxLength       =   20
                  TabIndex        =   261
                  Text            =   "txtCODPROD"
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  Caption         =   "Quantidade"
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
                  Left            =   3720
                  TabIndex        =   265
                  Top             =   600
                  Width           =   990
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  Caption         =   "Unidade de Medida"
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
                  TabIndex        =   264
                  Top             =   600
                  Width           =   1665
               End
               Begin VB.Label lblDescProd 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lblDescProd"
                  ForeColor       =   &H80000008&
                  Height          =   285
                  Left            =   3720
                  TabIndex        =   263
                  Top             =   240
                  Width           =   7095
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  Caption         =   "Insumo"
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
                  TabIndex        =   260
                  Top             =   240
                  Width           =   615
               End
            End
         End
         Begin TabDlg.SSTab SSTab1 
            Height          =   6615
            Left            =   -74880
            TabIndex        =   246
            Top             =   720
            Width           =   15015
            _ExtentX        =   26485
            _ExtentY        =   11668
            _Version        =   393216
            Style           =   1
            Tab             =   1
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
            TabCaption(0)   =   "Estoque"
            TabPicture(0)   =   "frmCADPRODP.frx":0C00
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "grdESTOQUE"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Estoque de Litografia"
            TabPicture(1)   =   "frmCADPRODP.frx":0C1C
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "grdEstLit"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Estoque de Latas Montadas"
            TabPicture(2)   =   "frmCADPRODP.frx":0C38
            Tab(2).ControlEnabled=   0   'False
            Tab(2).ControlCount=   0
            Begin VSFlex8LCtl.VSFlexGrid grdEstLit 
               Height          =   6135
               Left            =   120
               TabIndex        =   248
               Top             =   360
               Width           =   14775
               _cx             =   26061
               _cy             =   10821
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
            Begin VSFlex8LCtl.VSFlexGrid grdESTOQUE 
               Height          =   6015
               Left            =   -74880
               TabIndex        =   247
               Top             =   480
               Width           =   10335
               _cx             =   18230
               _cy             =   10610
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
         Begin VB.Frame Frame9 
            Caption         =   "[ Recursos ]"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   5655
            Left            =   -74880
            TabIndex        =   234
            Top             =   600
            Width           =   10575
            Begin VB.CommandButton Command2 
               Height          =   300
               Left            =   10200
               Picture         =   "frmCADPRODP.frx":0C54
               Style           =   1  'Graphical
               TabIndex        =   237
               Top             =   600
               Width           =   300
            End
            Begin VB.CommandButton Command1 
               Height          =   300
               Left            =   10200
               Picture         =   "frmCADPRODP.frx":0D9E
               Style           =   1  'Graphical
               TabIndex        =   236
               Top             =   240
               Width           =   300
            End
            Begin VSFlex8LCtl.VSFlexGrid grdPRODENTR 
               Height          =   5295
               Left            =   120
               TabIndex        =   235
               Top             =   240
               Width           =   9975
               _cx             =   17595
               _cy             =   9340
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
         Begin VB.Frame Frame43 
            Height          =   6735
            Index           =   3
            Left            =   -74880
            TabIndex        =   205
            Top             =   600
            Width           =   10575
            Begin VB.CommandButton Command8 
               Caption         =   "&Grava Inagem"
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
               Left            =   3360
               Picture         =   "frmCADPRODP.frx":0EE8
               Style           =   1  'Graphical
               TabIndex        =   256
               Top             =   6120
               Width           =   1575
            End
            Begin VB.CommandButton Command7 
               Caption         =   "Limpa"
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
               Left            =   1680
               TabIndex        =   255
               Top             =   6120
               Width           =   1695
            End
            Begin VB.CommandButton cmdAbreArq 
               Caption         =   "Abre Imagem"
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
               Left            =   120
               TabIndex        =   206
               Top             =   6120
               Width           =   1575
            End
            Begin MSComDlg.CommonDialog cmoAbreArq 
               Left            =   5880
               Top             =   6120
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.Image Image1 
               Height          =   5895
               Left            =   120
               Stretch         =   -1  'True
               Top             =   120
               Width           =   10335
            End
         End
         Begin VB.CommandButton cmdexcUnidConv 
            Height          =   300
            Left            =   -64800
            Picture         =   "frmCADPRODP.frx":0FEA
            Style           =   1  'Graphical
            TabIndex        =   203
            Top             =   1020
            Width           =   300
         End
         Begin VB.CommandButton cmdIncUnidConv 
            Height          =   300
            Left            =   -64800
            Picture         =   "frmCADPRODP.frx":1134
            Style           =   1  'Graphical
            TabIndex        =   202
            Top             =   660
            Width           =   300
         End
         Begin VB.CommandButton Command31 
            Height          =   300
            Left            =   -64800
            Picture         =   "frmCADPRODP.frx":127E
            Style           =   1  'Graphical
            TabIndex        =   200
            Top             =   660
            Width           =   300
         End
         Begin VB.CommandButton Command30 
            Height          =   300
            Left            =   -64800
            Picture         =   "frmCADPRODP.frx":13C8
            Style           =   1  'Graphical
            TabIndex        =   199
            Top             =   1020
            Width           =   300
         End
         Begin VB.Frame Frame12 
            Height          =   6615
            Left            =   -74880
            TabIndex        =   197
            Top             =   660
            Width           =   10455
            Begin VB.TextBox txtDescEquip 
               Height          =   6255
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   198
               Text            =   "frmCADPRODP.frx":1512
               Top             =   240
               Width           =   10215
            End
         End
         Begin VB.Frame Frame16 
            Height          =   6615
            Left            =   -74880
            TabIndex        =   188
            Top             =   660
            Width           =   10455
            Begin VB.TextBox txtRua 
               Height          =   315
               Left            =   1560
               MaxLength       =   30
               TabIndex        =   192
               Text            =   "txtRua"
               Top             =   240
               Width           =   1575
            End
            Begin VB.TextBox txtBox 
               Height          =   315
               Left            =   1560
               MaxLength       =   30
               TabIndex        =   191
               Text            =   "txtBox"
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox txtPrateleira 
               Height          =   315
               Left            =   1560
               MaxLength       =   30
               TabIndex        =   190
               Text            =   "txtPrateleira"
               Top             =   960
               Width           =   1575
            End
            Begin VB.TextBox txtDistancia 
               Height          =   315
               Left            =   1560
               MaxLength       =   30
               TabIndex        =   189
               Text            =   "txtDistancia"
               Top             =   1320
               Width           =   1575
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Rua"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   2
               Left            =   300
               TabIndex        =   196
               Top             =   240
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Box"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   3
               Left            =   300
               TabIndex        =   195
               Top             =   600
               Width           =   330
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Prateleira"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   4
               Left            =   300
               TabIndex        =   194
               Top             =   960
               Width           =   825
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Distância"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   5
               Left            =   300
               TabIndex        =   193
               Top             =   1320
               Width           =   810
            End
         End
         Begin TabDlg.SSTab stFichaTec 
            Height          =   6615
            Left            =   -74880
            TabIndex        =   139
            Top             =   660
            Width           =   15015
            _ExtentX        =   26485
            _ExtentY        =   11668
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
            TabCaption(0)   =   "Litografia"
            TabPicture(0)   =   "frmCADPRODP.frx":151F
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "fraGeral2(1)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "fraGeral1(0)"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "Folha de Flandres"
            TabPicture(1)   =   "frmCADPRODP.frx":153B
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame13"
            Tab(1).Control(1)=   "Frame10"
            Tab(1).Control(2)=   "Frame8"
            Tab(1).Control(3)=   "Frame7"
            Tab(1).Control(4)=   "txtQtdePassada"
            Tab(1).Control(5)=   "Frame38(3)"
            Tab(1).Control(6)=   "Frame38(0)"
            Tab(1).Control(7)=   "Frame38(1)"
            Tab(1).Control(8)=   "label11(18)"
            Tab(1).Control(9)=   "label11(9)"
            Tab(1).Control(10)=   "label11(8)"
            Tab(1).Control(11)=   "label11(2)"
            Tab(1).Control(12)=   "label11(1)"
            Tab(1).Control(13)=   "label11(0)"
            Tab(1).Control(14)=   "label11(4)"
            Tab(1).ControlCount=   15
            TabCaption(2)   =   "Montagem"
            TabPicture(2)   =   "frmCADPRODP.frx":1557
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Frame43(0)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Fechamento"
            TabPicture(3)   =   "frmCADPRODP.frx":1573
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "Frame43(1)"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "Expedição"
            TabPicture(4)   =   "frmCADPRODP.frx":158F
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "Frame38(7)"
            Tab(4).ControlCount=   1
            TabCaption(5)   =   "Clientes"
            TabPicture(5)   =   "frmCADPRODP.frx":15AB
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "Frame43(2)"
            Tab(5).ControlCount=   1
            Begin VB.Frame Frame13 
               Height          =   615
               Left            =   -69000
               TabIndex        =   241
               Top             =   2280
               Width           =   2175
               Begin VB.TextBox txtQtdePorFolha 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   120
                  TabIndex        =   242
                  Text            =   "txtQtdePorFolha"
                  Top             =   240
                  Width           =   1815
               End
            End
            Begin VB.Frame Frame10 
               Caption         =   "[ Padrão ]"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   615
               Left            =   -71160
               TabIndex        =   238
               Top             =   2280
               Width           =   2055
               Begin VB.OptionButton optQTDCORPPADRAOSN 
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
                  ForeColor       =   &H000000FF&
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   240
                  Top             =   240
                  Width           =   855
               End
               Begin VB.OptionButton optQTDCORPPADRAOSN 
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
                  ForeColor       =   &H000000FF&
                  Height          =   255
                  Index           =   0
                  Left            =   960
                  TabIndex        =   239
                  Top             =   240
                  Width           =   975
               End
            End
            Begin VB.Frame Frame8 
               Height          =   615
               Left            =   -69120
               TabIndex        =   228
               Top             =   3360
               Width           =   4455
               Begin VB.TextBox txtALTURA 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   3480
                  TabIndex        =   232
                  Text            =   "txtALTURA"
                  Top             =   240
                  Width           =   855
               End
               Begin VB.TextBox txtDESENV 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   231
                  Text            =   "txtDESENV"
                  Top             =   240
                  Width           =   855
               End
               Begin VB.Label Label7 
                  Caption         =   "Altura"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   255
                  Left            =   2760
                  TabIndex        =   230
                  Top             =   240
                  Width           =   615
               End
               Begin VB.Label Label3 
                  Caption         =   "Desenvolvimento"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   255
                  Left            =   120
                  TabIndex        =   229
                  Top             =   240
                  Width           =   1455
               End
            End
            Begin VB.Frame Frame7 
               Caption         =   "[ Padrão ]"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   615
               Left            =   -71160
               TabIndex        =   225
               Top             =   3360
               Width           =   1935
               Begin VB.OptionButton optDimCortePAD 
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
                  ForeColor       =   &H000000FF&
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   227
                  Top             =   240
                  Width           =   735
               End
               Begin VB.OptionButton optDimCortePAD 
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
                  ForeColor       =   &H000000FF&
                  Height          =   255
                  Index           =   0
                  Left            =   960
                  TabIndex        =   226
                  Top             =   240
                  Width           =   855
               End
            End
            Begin VB.TextBox txtQtdePassada 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   -71160
               TabIndex        =   212
               Text            =   "txtQtdePassada"
               Top             =   3000
               Width           =   1815
            End
            Begin VB.Frame Frame43 
               Height          =   6015
               Index           =   2
               Left            =   -74880
               TabIndex        =   174
               Top             =   360
               Width           =   10215
               Begin VB.CheckBox chkClientes 
                  Caption         =   "Puxa todos os Clientes"
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
                  Left            =   120
                  TabIndex        =   182
                  Top             =   5640
                  Width           =   2415
               End
               Begin VB.CommandButton cmdExcCli 
                  Height          =   300
                  Left            =   9840
                  Picture         =   "frmCADPRODP.frx":15C7
                  Style           =   1  'Graphical
                  TabIndex        =   83
                  Top             =   600
                  Width           =   300
               End
               Begin VB.CommandButton cmdIncCli 
                  Height          =   300
                  Left            =   9840
                  Picture         =   "frmCADPRODP.frx":1711
                  Style           =   1  'Graphical
                  TabIndex        =   82
                  Top             =   240
                  Width           =   300
               End
               Begin VSFlex8LCtl.VSFlexGrid grdClientes 
                  Height          =   5295
                  Left            =   120
                  TabIndex        =   81
                  Top             =   240
                  Width           =   9615
                  _cx             =   16960
                  _cy             =   9340
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
            Begin VB.Frame Frame38 
               Caption         =   "[ Expedição ]"
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
               Height          =   6015
               Index           =   7
               Left            =   -74880
               TabIndex        =   156
               Top             =   360
               Width           =   10215
               Begin VB.Frame Frame18 
                  BorderStyle     =   0  'None
                  Height          =   255
                  Left            =   6600
                  TabIndex        =   252
                  Top             =   360
                  Width           =   2895
                  Begin VB.OptionButton optUsarPadrPalSN 
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
                     Left            =   0
                     TabIndex        =   254
                     Top             =   0
                     Width           =   855
                  End
                  Begin VB.OptionButton optUsarPadrPalSN 
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
                     Left            =   840
                     TabIndex        =   253
                     Top             =   0
                     Width           =   975
                  End
               End
               Begin VB.Frame Frame4 
                  BorderStyle     =   0  'None
                  Height          =   255
                  Left            =   120
                  TabIndex        =   183
                  Top             =   840
                  Width           =   6495
                  Begin VB.OptionButton optLaudoSN 
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
                     ForeColor       =   &H8000000D&
                     Height          =   255
                     Index           =   1
                     Left            =   1440
                     TabIndex        =   186
                     Top             =   0
                     Width           =   735
                  End
                  Begin VB.OptionButton optLaudoSN 
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
                     ForeColor       =   &H8000000D&
                     Height          =   255
                     Index           =   0
                     Left            =   2160
                     TabIndex        =   185
                     Top             =   0
                     Width           =   735
                  End
                  Begin VB.Label Label1 
                     AutoSize        =   -1  'True
                     Caption         =   "Emite Laudo"
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
                     Index           =   24
                     Left            =   120
                     TabIndex        =   184
                     Top             =   0
                     Width           =   1065
                  End
               End
               Begin VB.Frame Frame38 
                  BorderStyle     =   0  'None
                  Height          =   255
                  Index           =   6
                  Left            =   120
                  TabIndex        =   157
                  Top             =   360
                  Width           =   4935
                  Begin VB.TextBox txtQtdeEmb 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   3480
                     TabIndex        =   80
                     Text            =   "txtQtdeEmb"
                     Top             =   0
                     Width           =   1095
                  End
                  Begin VB.OptionButton optTipPalet 
                     Caption         =   "GRANEL"
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
                     Index           =   1
                     Left            =   1320
                     TabIndex        =   79
                     Top             =   0
                     Width           =   1095
                  End
                  Begin VB.OptionButton optTipPalet 
                     Caption         =   "PALLET"
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
                     Index           =   0
                     Left            =   120
                     TabIndex        =   78
                     Top             =   0
                     Width           =   1095
                  End
                  Begin VB.Label label11 
                     Caption         =   "Qtde"
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
                     Index           =   17
                     Left            =   2880
                     TabIndex        =   158
                     Top             =   0
                     Width           =   615
                  End
               End
               Begin VB.Label Label8 
                  Caption         =   "Usar Padrão"
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
                  Left            =   5280
                  TabIndex        =   251
                  Top             =   360
                  Width           =   1455
               End
            End
            Begin VB.Frame Frame43 
               Height          =   6015
               Index           =   1
               Left            =   -74880
               TabIndex        =   150
               Top             =   360
               Width           =   10215
               Begin VB.Frame frmNeckIn 
                  BorderStyle     =   0  'None
                  Height          =   255
                  Left            =   8400
                  TabIndex        =   218
                  Top             =   360
                  Width           =   1695
                  Begin VB.OptionButton optNeckInSN 
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
                     ForeColor       =   &H000000FF&
                     Height          =   255
                     Index           =   1
                     Left            =   0
                     TabIndex        =   220
                     Top             =   0
                     Width           =   735
                  End
                  Begin VB.OptionButton optNeckInSN 
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
                     ForeColor       =   &H000000FF&
                     Height          =   255
                     Index           =   0
                     Left            =   720
                     TabIndex        =   219
                     Top             =   0
                     Width           =   735
                  End
               End
               Begin VB.Frame Frame41 
                  Height          =   2535
                  Index           =   4
                  Left            =   6840
                  TabIndex        =   173
                  Top             =   840
                  Width           =   1935
                  Begin VB.CommandButton cmdBtnFechaExc 
                     Height          =   300
                     Index           =   3
                     Left            =   1560
                     Picture         =   "frmCADPRODP.frx":185B
                     Style           =   1  'Graphical
                     TabIndex        =   74
                     Top             =   600
                     Width           =   300
                  End
                  Begin VB.CommandButton cmdBtnFechaInc 
                     Height          =   300
                     Index           =   3
                     Left            =   1560
                     Picture         =   "frmCADPRODP.frx":19A5
                     Style           =   1  'Graphical
                     TabIndex        =   73
                     Top             =   240
                     Width           =   300
                  End
                  Begin VSFlex8LCtl.VSFlexGrid grdSomFecha 
                     Height          =   2175
                     Index           =   3
                     Left            =   120
                     TabIndex        =   72
                     Top             =   240
                     Width           =   1335
                     _cx             =   2355
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
               Begin VB.Frame Frame41 
                  Height          =   2535
                  Index           =   3
                  Left            =   4560
                  TabIndex        =   172
                  Top             =   840
                  Width           =   2175
                  Begin VB.CommandButton cmdBtnFechaExc 
                     Height          =   300
                     Index           =   2
                     Left            =   1800
                     Picture         =   "frmCADPRODP.frx":1AEF
                     Style           =   1  'Graphical
                     TabIndex        =   71
                     Top             =   600
                     Width           =   300
                  End
                  Begin VB.CommandButton cmdBtnFechaInc 
                     Height          =   300
                     Index           =   2
                     Left            =   1800
                     Picture         =   "frmCADPRODP.frx":1C39
                     Style           =   1  'Graphical
                     TabIndex        =   70
                     Top             =   240
                     Width           =   300
                  End
                  Begin VSFlex8LCtl.VSFlexGrid grdSomFecha 
                     Height          =   2175
                     Index           =   2
                     Left            =   120
                     TabIndex        =   69
                     Top             =   240
                     Width           =   1575
                     _cx             =   2778
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
               Begin VB.Frame Frame41 
                  Height          =   2535
                  Index           =   2
                  Left            =   2400
                  TabIndex        =   171
                  Top             =   840
                  Width           =   2055
                  Begin VB.CommandButton cmdBtnFechaExc 
                     Height          =   300
                     Index           =   1
                     Left            =   1680
                     Picture         =   "frmCADPRODP.frx":1D83
                     Style           =   1  'Graphical
                     TabIndex        =   68
                     Top             =   600
                     Width           =   300
                  End
                  Begin VB.CommandButton cmdBtnFechaInc 
                     Height          =   300
                     Index           =   1
                     Left            =   1680
                     Picture         =   "frmCADPRODP.frx":1ECD
                     Style           =   1  'Graphical
                     TabIndex        =   67
                     Top             =   240
                     Width           =   300
                  End
                  Begin VSFlex8LCtl.VSFlexGrid grdSomFecha 
                     Height          =   2175
                     Index           =   1
                     Left            =   120
                     TabIndex        =   66
                     Top             =   240
                     Width           =   1455
                     _cx             =   2566
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
               Begin VB.Frame Frame38 
                  Caption         =   "[ Vedante ]"
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
                  Height          =   2535
                  Index           =   5
                  Left            =   120
                  TabIndex        =   160
                  Top             =   3360
                  Width           =   9975
                  Begin VB.CommandButton Command29 
                     Height          =   300
                     Left            =   9600
                     Picture         =   "frmCADPRODP.frx":2017
                     Style           =   1  'Graphical
                     TabIndex        =   76
                     Top             =   240
                     Width           =   300
                  End
                  Begin VB.CommandButton Command28 
                     Height          =   300
                     Left            =   9600
                     Picture         =   "frmCADPRODP.frx":2161
                     Style           =   1  'Graphical
                     TabIndex        =   77
                     Top             =   600
                     Width           =   300
                  End
                  Begin VSFlex8LCtl.VSFlexGrid grdVedanteCompound 
                     Height          =   2175
                     Left            =   120
                     TabIndex        =   75
                     Top             =   240
                     Width           =   9375
                     _cx             =   16536
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
               Begin VB.Frame Frame41 
                  Height          =   2535
                  Index           =   1
                  Left            =   120
                  TabIndex        =   159
                  Top             =   840
                  Width           =   2175
                  Begin VB.CommandButton cmdBtnFechaExc 
                     Height          =   300
                     Index           =   0
                     Left            =   1800
                     Picture         =   "frmCADPRODP.frx":22AB
                     Style           =   1  'Graphical
                     TabIndex        =   65
                     Top             =   600
                     Width           =   300
                  End
                  Begin VB.CommandButton cmdBtnFechaInc 
                     Height          =   300
                     Index           =   0
                     Left            =   1800
                     Picture         =   "frmCADPRODP.frx":23F5
                     Style           =   1  'Graphical
                     TabIndex        =   64
                     Top             =   240
                     Width           =   300
                  End
                  Begin VSFlex8LCtl.VSFlexGrid grdSomFecha 
                     Height          =   2175
                     Index           =   0
                     Left            =   120
                     TabIndex        =   63
                     Top             =   240
                     Width           =   1575
                     _cx             =   2778
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
               Begin VB.Frame Frame38 
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
                  Index           =   4
                  Left            =   3360
                  TabIndex        =   152
                  Top             =   120
                  Width           =   3975
                  Begin VB.ComboBox cboFechTampaFuro 
                     Height          =   315
                     Left            =   1440
                     TabIndex        =   62
                     Text            =   "cboFechTampaFuro"
                     Top             =   240
                     Width           =   2415
                  End
                  Begin VB.Label label11 
                     Caption         =   "Tampa / Furo"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000FF&
                     Height          =   255
                     Index           =   3
                     Left            =   120
                     TabIndex        =   153
                     Top             =   240
                     Width           =   1335
                  End
               End
               Begin VB.Frame Frame41 
                  Height          =   735
                  Index           =   0
                  Left            =   120
                  TabIndex        =   151
                  Top             =   120
                  Width           =   3135
                  Begin VB.ComboBox cboFechSoldaAgraf 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   61
                     Text            =   "cboFechSoldaAgraf"
                     Top             =   240
                     Width           =   2895
                  End
               End
               Begin VB.Label label11 
                  Caption         =   "Neck IN"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   255
                  Index           =   11
                  Left            =   7440
                  TabIndex        =   217
                  Top             =   360
                  Width           =   735
               End
            End
            Begin VB.Frame Frame43 
               Height          =   6015
               Index           =   0
               Left            =   -74880
               TabIndex        =   149
               Top             =   360
               Width           =   10215
               Begin VB.Frame Frame5 
                  BorderStyle     =   0  'None
                  Height          =   255
                  Left            =   2040
                  TabIndex        =   214
                  Top             =   2760
                  Width           =   2415
                  Begin VB.OptionButton optAlcGalSN 
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
                     Left            =   0
                     TabIndex        =   216
                     Top             =   0
                     Width           =   855
                  End
                  Begin VB.OptionButton optAlcGalSN 
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
                     Left            =   840
                     TabIndex        =   215
                     Top             =   0
                     Width           =   735
                  End
               End
               Begin VB.TextBox txtColEsp 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2040
                  MaxLength       =   10
                  TabIndex        =   60
                  Text            =   "txtColEsp"
                  Top             =   2280
                  Width           =   1215
               End
               Begin VB.CommandButton cmdPesqMont 
                  Height          =   315
                  Index           =   1
                  Left            =   3240
                  Picture         =   "frmCADPRODP.frx":253F
                  Style           =   1  'Graphical
                  TabIndex        =   167
                  Top             =   2280
                  Width           =   375
               End
               Begin VB.TextBox txtVerCat 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2040
                  MaxLength       =   10
                  TabIndex        =   59
                  Text            =   "txtVerCat"
                  Top             =   1920
                  Width           =   1215
               End
               Begin VB.CommandButton cmdPesqMont 
                  Height          =   315
                  Index           =   0
                  Left            =   3240
                  Picture         =   "frmCADPRODP.frx":2641
                  Style           =   1  'Graphical
                  TabIndex        =   165
                  Top             =   1920
                  Width           =   375
               End
               Begin VB.Frame Frame39 
                  BorderStyle     =   0  'None
                  Height          =   255
                  Left            =   2040
                  TabIndex        =   164
                  Top             =   240
                  Width           =   2415
                  Begin VB.OptionButton optAzSimNao 
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
                     ForeColor       =   &H8000000D&
                     Height          =   255
                     Index           =   1
                     Left            =   0
                     TabIndex        =   53
                     Top             =   0
                     Width           =   735
                  End
                  Begin VB.OptionButton optAzSimNao 
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
                     ForeColor       =   &H8000000D&
                     Height          =   255
                     Index           =   0
                     Left            =   840
                     TabIndex        =   54
                     Top             =   0
                     Width           =   855
                  End
               End
               Begin VB.ComboBox cboAlcaPlastica 
                  Height          =   315
                  Left            =   2040
                  TabIndex        =   55
                  Text            =   "cboAlcaPlastica"
                  Top             =   600
                  Width           =   2295
               End
               Begin VB.ComboBox cboAlcaFerro 
                  Height          =   315
                  Left            =   2040
                  TabIndex        =   56
                  Text            =   "cboAlcaFerro"
                  Top             =   960
                  Width           =   2295
               End
               Begin VB.Frame Frame40 
                  BorderStyle     =   0  'None
                  Height          =   255
                  Left            =   2040
                  TabIndex        =   154
                  Top             =   1440
                  Width           =   1815
                  Begin VB.OptionButton optPipetSimNao 
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
                     ForeColor       =   &H8000000D&
                     Height          =   255
                     Index           =   0
                     Left            =   840
                     TabIndex        =   58
                     Top             =   0
                     Width           =   735
                  End
                  Begin VB.OptionButton optPipetSimNao 
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
                     ForeColor       =   &H8000000D&
                     Height          =   255
                     Index           =   1
                     Left            =   0
                     TabIndex        =   57
                     Top             =   0
                     Width           =   735
                  End
               End
               Begin VB.Label label11 
                  AutoSize        =   -1  'True
                  Caption         =   "Alça de galão"
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
                  Height          =   315
                  Index           =   10
                  Left            =   240
                  TabIndex        =   213
                  Top             =   2760
                  Width           =   1170
               End
               Begin VB.Label label11 
                  BackColor       =   &H8000000E&
                  BorderStyle     =   1  'Fixed Single
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
                  Index           =   16
                  Left            =   3600
                  TabIndex        =   170
                  Top             =   2280
                  Width           =   4935
               End
               Begin VB.Label label11 
                  BackColor       =   &H8000000E&
                  BorderStyle     =   1  'Fixed Single
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
                  Index           =   15
                  Left            =   3600
                  TabIndex        =   169
                  Top             =   1920
                  Width           =   4935
               End
               Begin VB.Label label11 
                  Caption         =   "Cola Especial"
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
                  Index           =   14
                  Left            =   240
                  TabIndex        =   168
                  Top             =   2280
                  Width           =   1815
               End
               Begin VB.Label label11 
                  Caption         =   "Verniz  Catalizador"
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
                  Index           =   13
                  Left            =   240
                  TabIndex        =   166
                  Top             =   1920
                  Width           =   1815
               End
               Begin VB.Label label11 
                  Caption         =   "Azelha"
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
                  Index           =   12
                  Left            =   240
                  TabIndex        =   163
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.Label label11 
                  Caption         =   "Alça Plastica"
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
                  Index           =   5
                  Left            =   240
                  TabIndex        =   162
                  Top             =   600
                  Width           =   1335
               End
               Begin VB.Label label11 
                  Caption         =   "Alça de Ferro"
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
                  Index           =   6
                  Left            =   240
                  TabIndex        =   161
                  Top             =   960
                  Width           =   1335
               End
               Begin VB.Label label11 
                  Caption         =   "Pipeta"
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
                  Index           =   7
                  Left            =   240
                  TabIndex        =   155
                  Top             =   1440
                  Width           =   735
               End
            End
            Begin VB.Frame Frame38 
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
               ForeColor       =   &H000000FF&
               Height          =   1815
               Index           =   3
               Left            =   -74160
               TabIndex        =   144
               Top             =   360
               Width           =   3615
               Begin VB.ComboBox cboTampaVerniz 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   39
                  Text            =   "cboTampaVerniz"
                  Top             =   600
                  Width           =   3375
               End
               Begin VB.ComboBox cboFundoVerniz 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   43
                  Text            =   "cboFundoVerniz"
                  Top             =   960
                  Width           =   3375
               End
               Begin VB.ComboBox cboArgolaVerniz 
                  Height          =   315
                  ItemData        =   "frmCADPRODP.frx":2743
                  Left            =   120
                  List            =   "frmCADPRODP.frx":2745
                  TabIndex        =   47
                  Text            =   "cboArgolaVerniz"
                  Top             =   1320
                  Width           =   3375
               End
               Begin VB.ComboBox cboCorpoVerniz 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   35
                  Text            =   "cboCorpoVerniz"
                  Top             =   240
                  Width           =   3375
               End
            End
            Begin VB.Frame Frame38 
               Caption         =   "[ Espess. ]"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   1815
               Index           =   0
               Left            =   -70440
               TabIndex        =   143
               Top             =   360
               Width           =   1335
               Begin VB.TextBox txtCorpoEspess 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   120
                  TabIndex        =   36
                  Text            =   "txtCorpoEspess"
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.TextBox txtTampaEspess 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   120
                  TabIndex        =   40
                  Text            =   "txtTampaEspess"
                  Top             =   600
                  Width           =   1095
               End
               Begin VB.TextBox txtFundoEspess 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   120
                  TabIndex        =   44
                  Text            =   "txtFundoEspess"
                  Top             =   960
                  Width           =   1095
               End
               Begin VB.TextBox txtArgolaEspess 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   120
                  TabIndex        =   48
                  Text            =   "txtArgolaEspess"
                  Top             =   1320
                  Width           =   1095
               End
            End
            Begin VB.Frame Frame38 
               Caption         =   "[ Revest. ]"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   1815
               Index           =   1
               Left            =   -69000
               TabIndex        =   142
               Top             =   360
               Width           =   2775
               Begin VB.TextBox txtArgolaRevest2 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1560
                  TabIndex        =   50
                  Text            =   "txtTampaRevest"
                  Top             =   1320
                  Width           =   1095
               End
               Begin VB.TextBox txtFundoRevest2 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1560
                  TabIndex        =   46
                  Text            =   "txtTampaRevest"
                  Top             =   960
                  Width           =   1095
               End
               Begin VB.TextBox txtTampaRevest2 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1560
                  TabIndex        =   42
                  Text            =   "txtTampaRevest"
                  Top             =   600
                  Width           =   1095
               End
               Begin VB.TextBox txtCorpoRevest2 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1560
                  TabIndex        =   38
                  Text            =   "txtCorpoRevest"
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.TextBox txtCorpoRevest 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   120
                  TabIndex        =   37
                  Text            =   "txtCorpoRevest"
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.TextBox txtTampaRevest 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   120
                  TabIndex        =   41
                  Text            =   "txtTampaRevest"
                  Top             =   600
                  Width           =   1095
               End
               Begin VB.TextBox txtFundoRevest 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   120
                  TabIndex        =   45
                  Text            =   "txtFundoRevest"
                  Top             =   960
                  Width           =   1095
               End
               Begin VB.TextBox txtArgolaRevest 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   120
                  TabIndex        =   49
                  Text            =   "txtArgolaRevest"
                  Top             =   1320
                  Width           =   1095
               End
               Begin VB.Line Line3 
                  BorderColor     =   &H80000002&
                  Index           =   3
                  X1              =   1320
                  X2              =   1440
                  Y1              =   1560
                  Y2              =   1320
               End
               Begin VB.Line Line3 
                  BorderColor     =   &H80000002&
                  Index           =   2
                  X1              =   1320
                  X2              =   1440
                  Y1              =   1200
                  Y2              =   960
               End
               Begin VB.Line Line3 
                  BorderColor     =   &H80000002&
                  Index           =   1
                  X1              =   1320
                  X2              =   1440
                  Y1              =   840
                  Y2              =   600
               End
               Begin VB.Line Line3 
                  BorderColor     =   &H80000002&
                  Index           =   0
                  X1              =   1320
                  X2              =   1440
                  Y1              =   480
                  Y2              =   240
               End
            End
            Begin VB.Frame fraGeral1 
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
               Height          =   3375
               Index           =   0
               Left            =   120
               TabIndex        =   141
               Top             =   3120
               Width           =   10215
               Begin VB.CommandButton Command26 
                  Height          =   300
                  Left            =   9720
                  Picture         =   "frmCADPRODP.frx":2747
                  Style           =   1  'Graphical
                  TabIndex        =   33
                  Top             =   600
                  Width           =   300
               End
               Begin VB.CommandButton Command27 
                  Height          =   300
                  Left            =   9720
                  Picture         =   "frmCADPRODP.frx":2891
                  Style           =   1  'Graphical
                  TabIndex        =   32
                  Top             =   240
                  Width           =   300
               End
               Begin VSFlex8LCtl.VSFlexGrid grdCores 
                  Height          =   3015
                  Left            =   120
                  TabIndex        =   31
                  Top             =   240
                  Width           =   9495
                  _cx             =   16748
                  _cy             =   5318
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
            Begin VB.Frame fraGeral2 
               Caption         =   "[Verniz / Esmalte ]"
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
               Height          =   2655
               Index           =   1
               Left            =   120
               TabIndex        =   140
               Top             =   360
               Width           =   10215
               Begin VSFlex8LCtl.VSFlexGrid grdVernizEsm 
                  Height          =   2295
                  Left            =   120
                  TabIndex        =   34
                  Top             =   240
                  Width           =   9975
                  _cx             =   17595
                  _cy             =   4048
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
            Begin VB.Label label11 
               Caption         =   "Dimensão de Corte para desenvolvimento"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Index           =   18
               Left            =   -74880
               TabIndex        =   224
               Top             =   3600
               Width           =   3735
            End
            Begin VB.Label label11 
               Caption         =   "Qtde de Passadas na Litografia"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Index           =   9
               Left            =   -74880
               TabIndex        =   211
               Top             =   3000
               Width           =   2895
            End
            Begin VB.Label label11 
               Caption         =   "Qtde por Folha"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Index           =   8
               Left            =   -74880
               TabIndex        =   187
               Top             =   2400
               Width           =   1455
            End
            Begin VB.Label label11 
               Caption         =   "Argola"
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
               Index           =   2
               Left            =   -74880
               TabIndex        =   148
               Top             =   1680
               Width           =   615
            End
            Begin VB.Label label11 
               Caption         =   "Fundo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Index           =   1
               Left            =   -74880
               TabIndex        =   147
               Top             =   1320
               Width           =   735
            End
            Begin VB.Label label11 
               Caption         =   "Tampa"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Index           =   0
               Left            =   -74880
               TabIndex        =   146
               Top             =   960
               Width           =   735
            End
            Begin VB.Label label11 
               Caption         =   "Corpo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Index           =   4
               Left            =   -74880
               TabIndex        =   145
               Top             =   600
               Width           =   735
            End
         End
         Begin VB.Frame Frame3 
            Height          =   6735
            Left            =   -74880
            TabIndex        =   84
            Top             =   660
            Width           =   15015
            Begin VB.TextBox txtIPI 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   5160
               TabIndex        =   249
               Text            =   "txtIPI"
               Top             =   5280
               Width           =   1095
            End
            Begin VB.Frame Frame17 
               Caption         =   "[ Produto Novo ]"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   615
               Left            =   7680
               TabIndex        =   243
               Top             =   3120
               Width           =   2655
               Begin VB.OptionButton optProdNovoSN 
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
                  ForeColor       =   &H000000FF&
                  Height          =   255
                  Index           =   1
                  Left            =   240
                  TabIndex        =   245
                  Top             =   240
                  Width           =   735
               End
               Begin VB.OptionButton optProdNovoSN 
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
                  ForeColor       =   &H000000FF&
                  Height          =   255
                  Index           =   0
                  Left            =   1080
                  TabIndex        =   244
                  Top             =   240
                  Width           =   735
               End
            End
            Begin VB.TextBox txtComplemento 
               Height          =   285
               Left            =   2040
               MaxLength       =   20
               TabIndex        =   7
               Text            =   "txtComplemento"
               Top             =   2520
               Width           =   3015
            End
            Begin VB.Frame Frame6 
               Caption         =   "[ Filial ]"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   615
               Left            =   7680
               TabIndex        =   221
               Top             =   3720
               Width           =   2655
               Begin VB.OptionButton optFilialPed 
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
                  ForeColor       =   &H000000FF&
                  Height          =   255
                  Index           =   1
                  Left            =   1560
                  TabIndex        =   223
                  Top             =   240
                  Width           =   975
               End
               Begin VB.OptionButton optFilialPed 
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
                  ForeColor       =   &H000000FF&
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   222
                  Top             =   240
                  Width           =   1335
               End
            End
            Begin VB.TextBox txtDigVerif 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1920
               MaxLength       =   10
               TabIndex        =   3
               Text            =   "txtDigVeri"
               Top             =   1200
               Width           =   975
            End
            Begin VB.Frame Frame38 
               BorderStyle     =   0  'None
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
               Height          =   1095
               Index           =   2
               Left            =   1800
               TabIndex        =   132
               Top             =   120
               Width           =   8535
               Begin VB.Frame Frame46 
                  BorderStyle     =   0  'None
                  Height          =   255
                  Left            =   2280
                  TabIndex        =   176
                  Top             =   840
                  Width           =   2415
                  Begin VB.OptionButton optNatSimNao 
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
                     ForeColor       =   &H000000FF&
                     Height          =   255
                     Index           =   1
                     Left            =   0
                     TabIndex        =   178
                     Top             =   0
                     Width           =   735
                  End
                  Begin VB.OptionButton optNatSimNao 
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
                     ForeColor       =   &H000000FF&
                     Height          =   255
                     Index           =   0
                     Left            =   840
                     TabIndex        =   177
                     Top             =   0
                     Width           =   735
                  End
               End
               Begin VB.TextBox txtLinProd 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   120
                  MaxLength       =   10
                  TabIndex        =   0
                  Text            =   "txtLinProd"
                  Top             =   0
                  Width           =   975
               End
               Begin VB.TextBox txtCodCliente 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   120
                  MaxLength       =   10
                  TabIndex        =   1
                  Text            =   "txtCodClie"
                  Top             =   360
                  Width           =   975
               End
               Begin VB.CommandButton cmdLinProd 
                  Height          =   315
                  Left            =   1080
                  Picture         =   "frmCADPRODP.frx":29DB
                  Style           =   1  'Graphical
                  TabIndex        =   134
                  Top             =   0
                  Width           =   375
               End
               Begin VB.CommandButton cmdClie 
                  Height          =   315
                  Left            =   1080
                  Picture         =   "frmCADPRODP.frx":2ADD
                  Style           =   1  'Graphical
                  TabIndex        =   133
                  Top             =   360
                  Width           =   375
               End
               Begin VB.TextBox txtCodRot 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   120
                  MaxLength       =   10
                  TabIndex        =   2
                  Text            =   "txtCodRot"
                  Top             =   720
                  Width           =   975
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Natural ?"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   195
                  Index           =   20
                  Left            =   1320
                  TabIndex        =   175
                  Top             =   840
                  Width           =   795
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  Caption         =   "lblCodProd"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   315
                  Index           =   14
                  Left            =   5880
                  TabIndex        =   138
                  Top             =   360
                  Width           =   2655
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Código do Produto"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   195
                  Index           =   13
                  Left            =   5880
                  TabIndex        =   137
                  Top             =   0
                  Width           =   1590
               End
               Begin VB.Label lblLinhProd 
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lblLinhProd"
                  Height          =   315
                  Left            =   1440
                  TabIndex        =   136
                  Top             =   0
                  Width           =   4215
               End
               Begin VB.Label lblDesclie 
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lblDesclie"
                  Height          =   315
                  Left            =   1440
                  TabIndex        =   135
                  Top             =   360
                  Width           =   4215
               End
            End
            Begin VB.TextBox txtGRAMATURAM2 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   7680
               TabIndex        =   30
               Text            =   "txtGRAMATURAM2"
               Top             =   6360
               Width           =   975
            End
            Begin VB.TextBox txtLARGURA 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   5160
               TabIndex        =   29
               Text            =   "txtLARGURA"
               Top             =   6360
               Width           =   1095
            End
            Begin VB.TextBox txtMETROS 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2520
               TabIndex        =   28
               Text            =   "txtMETROS"
               Top             =   6360
               Width           =   1215
            End
            Begin VB.TextBox txtCUBAGEN 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   7680
               TabIndex        =   27
               Text            =   "txtCUBAGEN"
               Top             =   6000
               Width           =   975
            End
            Begin VB.TextBox txtCodigoEAN 
               Height          =   315
               Left            =   4440
               MaxLength       =   30
               TabIndex        =   16
               Text            =   "txtCodigoEAN"
               Top             =   4920
               Width           =   1575
            End
            Begin VB.TextBox txtCodProdFornec 
               Height          =   315
               Left            =   5880
               MaxLength       =   10
               TabIndex        =   5
               Text            =   "txtCodProdFornec"
               Top             =   1800
               Width           =   1695
            End
            Begin VB.TextBox txtPeso 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4680
               MaxLength       =   10
               TabIndex        =   13
               Text            =   "txtPeso"
               Top             =   4560
               Width           =   1095
            End
            Begin VB.Frame Frame15 
               Caption         =   "Estilo Produto"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   615
               Left            =   7680
               TabIndex        =   114
               Top             =   2520
               Width           =   2655
               Begin VB.OptionButton optEstProduto 
                  Caption         =   "De Clientes"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   255
                  Index           =   5
                  Left            =   2280
                  TabIndex        =   123
                  Top             =   240
                  Width           =   255
               End
               Begin VB.OptionButton optEstProduto 
                  Caption         =   "Bruto"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   255
                  Index           =   4
                  Left            =   1320
                  TabIndex        =   122
                  Top             =   240
                  Width           =   855
               End
               Begin VB.OptionButton optEstProduto 
                  Caption         =   "Barras"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   255
                  Index           =   3
                  Left            =   2280
                  TabIndex        =   121
                  Top             =   240
                  Width           =   210
               End
               Begin VB.OptionButton optEstProduto 
                  Caption         =   "Peças"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   255
                  Index           =   2
                  Left            =   2280
                  TabIndex        =   117
                  Top             =   240
                  Width           =   255
               End
               Begin VB.OptionButton optEstProduto 
                  Caption         =   "Bruto"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   255
                  Index           =   1
                  Left            =   2280
                  TabIndex        =   116
                  Top             =   240
                  Width           =   210
               End
               Begin VB.OptionButton optEstProduto 
                  Caption         =   "Acabado"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   115
                  Top             =   240
                  Width           =   1215
               End
            End
            Begin VB.Frame Frame14 
               Caption         =   "Tipo de Produto"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   735
               Left            =   7680
               TabIndex        =   111
               Top             =   1800
               Width           =   2655
               Begin VB.OptionButton optTipProd 
                  Caption         =   "Rótulo"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   195
                  Index           =   1
                  Left            =   120
                  TabIndex        =   113
                  Top             =   480
                  Width           =   1215
               End
               Begin VB.OptionButton optTipProd 
                  Caption         =   "Materia Prima"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   195
                  Index           =   0
                  Left            =   120
                  TabIndex        =   112
                  Top             =   240
                  Width           =   1695
               End
            End
            Begin VB.TextBox txtPRCFINAL 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   285
               Left            =   7680
               TabIndex        =   24
               Text            =   "txtPRCFINAL"
               Top             =   5640
               Width           =   975
            End
            Begin VB.TextBox txtPOCACRES 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   6840
               TabIndex        =   23
               Text            =   "txtPOCACRES"
               Top             =   5640
               Width           =   615
            End
            Begin VB.TextBox txtCODSUBGRUP 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2040
               MaxLength       =   10
               TabIndex        =   9
               Text            =   "txtCODSUBG"
               Top             =   3240
               Width           =   735
            End
            Begin VB.CommandButton Command6 
               Height          =   315
               Left            =   2760
               Picture         =   "frmCADPRODP.frx":2BDF
               Style           =   1  'Graphical
               TabIndex        =   109
               Top             =   3240
               Width           =   375
            End
            Begin VB.TextBox txtCODGRUP 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2040
               MaxLength       =   10
               TabIndex        =   8
               Text            =   "txtCODGRUP"
               Top             =   2880
               Width           =   735
            End
            Begin VB.CommandButton Command5 
               Height          =   315
               Left            =   2760
               Picture         =   "frmCADPRODP.frx":2CE1
               Style           =   1  'Graphical
               TabIndex        =   107
               Top             =   2880
               Width           =   375
            End
            Begin VB.TextBox txtProcedencia 
               Height          =   285
               Left            =   7680
               MaxLength       =   20
               TabIndex        =   20
               Text            =   "txtProcedencia"
               Top             =   5280
               Width           =   2535
            End
            Begin VB.Frame Frame11 
               BorderStyle     =   0  'None
               Height          =   255
               Left            =   2520
               TabIndex        =   104
               Top             =   5280
               Width           =   1695
               Begin VB.OptionButton optATUAUTOMSIMNAO 
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
                  ForeColor       =   &H00C00000&
                  Height          =   255
                  Index           =   1
                  Left            =   840
                  TabIndex        =   19
                  Top             =   0
                  Width           =   735
               End
               Begin VB.OptionButton optATUAUTOMSIMNAO 
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
                  ForeColor       =   &H00C00000&
                  Height          =   255
                  Index           =   0
                  Left            =   0
                  TabIndex        =   18
                  Top             =   0
                  Width           =   735
               End
            End
            Begin VB.TextBox txtPRECOPRODUTO 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   5160
               TabIndex        =   22
               Text            =   "txtPRECOPRODUTO"
               Top             =   5640
               Width           =   1095
            End
            Begin VB.ComboBox cboClass 
               Height          =   315
               Left            =   7680
               TabIndex        =   14
               Text            =   "cboClass"
               Top             =   4560
               Width           =   2055
            End
            Begin VB.CommandButton Command4 
               Height          =   315
               Left            =   2760
               Picture         =   "frmCADPRODP.frx":2DE3
               Style           =   1  'Graphical
               TabIndex        =   95
               Top             =   3960
               Width           =   375
            End
            Begin VB.CommandButton Command3 
               Height          =   315
               Left            =   2760
               Picture         =   "frmCADPRODP.frx":2EE5
               Style           =   1  'Graphical
               TabIndex        =   94
               Top             =   3600
               Width           =   375
            End
            Begin VB.ComboBox cboUnidade 
               Height          =   315
               Left            =   1680
               TabIndex        =   12
               Text            =   "cboUnidade"
               Top             =   4560
               Width           =   855
            End
            Begin VB.TextBox txtTipo 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2040
               MaxLength       =   10
               TabIndex        =   11
               Text            =   "txtTipo"
               Top             =   3960
               Width           =   735
            End
            Begin VB.TextBox txtEspecie 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2040
               MaxLength       =   10
               TabIndex        =   10
               Text            =   "txtEspecie"
               Top             =   3600
               Width           =   735
            End
            Begin VB.TextBox txtEstMinimo 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   7680
               MaxLength       =   10
               TabIndex        =   17
               Text            =   "txtEstMini"
               Top             =   4920
               Width           =   2055
            End
            Begin VB.TextBox txtSaldo 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1680
               MaxLength       =   10
               TabIndex        =   15
               Text            =   "txtSaldo"
               Top             =   4920
               Width           =   1095
            End
            Begin MSMask.MaskEdBox mskDtCadastro 
               Height          =   315
               Left            =   2520
               TabIndex        =   21
               Top             =   5640
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskDtUltMov 
               Height          =   315
               Left            =   2520
               TabIndex        =   25
               Top             =   6000
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.TextBox txtDescricao 
               Height          =   315
               Left            =   2040
               MaxLength       =   50
               TabIndex        =   6
               Text            =   "txtDescricao"
               Top             =   2160
               Width           =   5535
            End
            Begin VB.TextBox txtCodigo 
               Height          =   315
               Left            =   2040
               MaxLength       =   20
               TabIndex        =   4
               Text            =   "txtCodigo"
               Top             =   1800
               Width           =   1695
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "% IPI:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   2
               Left            =   4560
               TabIndex        =   250
               Top             =   5280
               Width           =   510
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Complemento"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   233
               Top             =   2520
               Width           =   1140
            End
            Begin VB.Label lblDescTipoProd 
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lblDescTipoProd"
               Height          =   285
               Left            =   3120
               TabIndex        =   210
               Top             =   3960
               Width           =   4455
            End
            Begin VB.Label lblDescEspecieProd 
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lblDescEspecieProd"
               Height          =   285
               Left            =   3120
               TabIndex        =   209
               Top             =   3600
               Width           =   4455
            End
            Begin VB.Label lblDescSubFamProd 
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lblDescSubFamProd"
               Height          =   285
               Left            =   3120
               TabIndex        =   208
               Top             =   3240
               Width           =   4455
            End
            Begin VB.Label lblDescFamProd 
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lblDescFamProd"
               Height          =   285
               Left            =   3120
               TabIndex        =   207
               Top             =   2880
               Width           =   4455
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lblSTATUS"
               Height          =   285
               Index           =   23
               Left            =   8400
               TabIndex        =   181
               Top             =   1200
               Width           =   1335
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Status"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   22
               Left            =   7680
               TabIndex        =   180
               Top             =   1200
               Width           =   555
            End
            Begin VB.Line Line2 
               X1              =   120
               X2              =   14880
               Y1              =   1680
               Y2              =   1680
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Digito Verificador:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   195
               Index           =   12
               Left            =   120
               TabIndex        =   131
               Top             =   1320
               Width           =   1545
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Rótulo:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   195
               Index           =   11
               Left            =   120
               TabIndex        =   130
               Top             =   960
               Width           =   630
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Código do Cliente :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   195
               Index           =   10
               Left            =   120
               TabIndex        =   129
               Top             =   600
               Width           =   1635
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Linha do Produto :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   195
               Index           =   9
               Left            =   120
               TabIndex        =   128
               Top             =   240
               Width           =   1590
            End
            Begin VB.Line Line1 
               X1              =   120
               X2              =   14880
               Y1              =   4440
               Y2              =   4440
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Gramatura/m2:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   2
               Left            =   6360
               TabIndex        =   127
               Top             =   6360
               Width           =   1275
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Largura/cm:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   1
               Left            =   3840
               TabIndex        =   126
               Top             =   6360
               Width           =   1050
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Metros:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   125
               Top             =   6360
               Width           =   645
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Cubagem:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   1
               Left            =   6360
               TabIndex        =   124
               Top             =   6000
               Width           =   855
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Código EAN:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   1
               Left            =   3000
               TabIndex        =   120
               Top             =   4950
               Width           =   1095
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Código Fornecedor:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   1
               Left            =   3960
               TabIndex        =   119
               Top             =   1845
               Width           =   1680
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Peso Unitário KG :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   2
               Left            =   3000
               TabIndex        =   118
               Top             =   4605
               Width           =   1590
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   1
               Left            =   6600
               TabIndex        =   110
               Top             =   5640
               Width           =   150
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Sub-Grupo:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   108
               Top             =   3240
               Width           =   975
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Grupo:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   106
               Top             =   2880
               Width           =   585
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Procedência"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   1
               Left            =   6360
               TabIndex        =   105
               Top             =   5280
               Width           =   1080
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Atualiza preço automático:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   16
               Left            =   120
               TabIndex        =   103
               Top             =   5280
               Width           =   2280
            End
            Begin VB.Label lblPRECOMEDIO 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lblPRECOMEDIO"
               Height          =   255
               Left            =   5160
               TabIndex        =   26
               Top             =   6000
               Width           =   1095
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Preço Médio:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   0
               Left            =   3840
               TabIndex        =   102
               Top             =   6000
               Width           =   1140
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Preço Produto:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   0
               Left            =   3840
               TabIndex        =   101
               Top             =   5640
               Width           =   1290
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Class. fiscal"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   1
               Left            =   6480
               TabIndex        =   100
               Top             =   4605
               Width           =   1035
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Estoque Minimo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   0
               Left            =   6120
               TabIndex        =   93
               Top             =   4950
               Width           =   1350
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Saldo:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   92
               Top             =   4950
               Width           =   555
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Unidade:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   91
               Top             =   4605
               Width           =   780
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Tipo:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   195
               Index           =   0
               Left            =   105
               TabIndex        =   90
               Top             =   3990
               Width           =   450
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Espécie:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   195
               Index           =   15
               Left            =   135
               TabIndex        =   89
               Top             =   3645
               Width           =   750
            End
            Begin VB.Label Label2 
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
               ForeColor       =   &H000000FF&
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   88
               Top             =   2205
               Width           =   930
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
               ForeColor       =   &H000000FF&
               Height          =   195
               Index           =   0
               Left            =   135
               TabIndex        =   87
               Top             =   1845
               Width           =   660
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Data Ultima Movimentação:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   18
               Left            =   120
               TabIndex        =   86
               Top             =   6045
               Width           =   2355
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Data do Cadastro:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   255
               Index           =   17
               Left            =   120
               TabIndex        =   85
               Top             =   5670
               Width           =   1560
            End
         End
         Begin VSFlex8LCtl.VSFlexGrid grdFamMaq 
            Height          =   6735
            Left            =   -74880
            TabIndex        =   201
            Top             =   660
            Width           =   9975
            _cx             =   17595
            _cy             =   11880
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
         Begin VSFlex8LCtl.VSFlexGrid grdUnidConv 
            Height          =   6735
            Left            =   -74880
            TabIndex        =   204
            Top             =   660
            Width           =   9975
            _cx             =   17595
            _cy             =   11880
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código do Produto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   21
      Left            =   0
      TabIndex        =   179
      Top             =   0
      Width           =   1590
   End
End
Attribute VB_Name = "frmCADPROD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho     As String
Public Linha        As Variant
Public cTipOper     As String
Public iCodigo      As Long
Public FILIAL       As Integer
Public strAcesso    As String
Public strUSUARIO   As String
Public lngCodUsuario As Long

Dim objBLBFunc      As Object
Dim objCADPRODUTO   As Object
Dim objCADFORN      As Object
Dim objPESQPADRAO   As Object

Dim arrProdLST      As Variant
Dim arrESPTECPROD   As Variant
Dim arrPROCESSO     As Variant
Dim arrFERRAMENTA   As Variant
Dim arrPRODFOR      As Variant
Dim strCodProc      As String
Dim arrPRODTIPORCA  As Variant
Dim arrPRODVAPROC   As Variant
Dim arrPRODVAREC    As Variant
Dim arrPRODVAINSP   As Variant
Dim arrPRODATRPROC  As Variant
Dim arrPRODATRREC   As Variant
Dim arrPRODATRINSP  As Variant
Dim arrCORRPRODPADR As Variant
Dim arrPARPRODPADR  As Variant
Dim arrUNIDCONV     As Variant
Dim arrCORES        As Variant
Dim arrVERNIZ       As Variant
Dim arrVERNIZ02     As Variant
Dim arrPEDIDOS      As Variant
Dim arrVERNIZEXT    As Variant

Dim arrVERNIZACAB   As Variant
Dim arrESMALTE      As Variant
Dim arrVEDANTE      As Variant

Dim arrTAMPAPRES        As Variant
Dim arrBATRETR          As Variant
Dim arrBATPLAST         As Variant
Dim arrTAMPAVISOR       As Variant
Dim arrPRODATECLIE      As Variant
Dim arrFAMMAQ           As Variant
Dim arrPRODENTR         As Variant

Dim strNOMARG           As String
Dim strCAMINHO          As String
Dim strCamImgRotulos    As String
Dim arrFILHO()          As PRODARVPROD


Dim lngINDLIST          As Long

' to handle node dragging
Private Const intHora = 60

Const DRAGTOL = 100                                 ' mouse movement before dragging starts

Const conCOL_SonFamMaq_CodFamMaq                    As Integer = 0
Const conCOL_SonFamMaq_PesqFam                      As Integer = 1
Const conCOL_SonFamMaq_DescFam                      As Integer = 2
Const conCOL_SonFamMaq_FormatString                 As String = "=Cód.Fam|...|Descrição"
Const conColumnsIn_SonFamMaq                        As Integer = 3

Const conCOL_SonProdEntr_IDProduto                  As Integer = 0
Const conCOL_SonProdEntr_CodProd                    As Integer = 1
Const conCOL_SonProdEntr_PesqProd                   As Integer = 2
Const conCOL_SonProdEntr_DescProd                   As Integer = 3
Const conCOL_SonProdEntr_UniMed                     As Integer = 4
Const conCOL_SonProdEntr_Qtde                       As Integer = 5
Const conCOL_SonProdEntr_FormatString               As String = "=IDProd|Cód.Prod|...|Descrição|Un.Medida|Qtde"
Const conColumnsIn_SonProdEntr                      As Integer = 6


Const conCOL_SonCoefic_CodParam                     As Integer = 0
Const conCOL_SonCoefic_Desc_Param                   As Integer = 1
Const conCOL_SonCoefic_UniMed                       As Integer = 2
Const conCOL_SonCoefic_ValParam                     As Integer = 3
Const conCOL_SonCoefic_CodParPad                    As Integer = 4
Const conCOL_SonCoefic_FormatString                 As String = "=Cód. Paramêtro|Descrição parâmetro|Unidade|Valor|"
Const conColumnsIn_SonCoefic                        As Integer = 5

Const conCOL_SonPara_CodParam                       As Integer = 0
Const conCOL_SonPara_Pesq                           As Integer = 1
Const conCOL_SonPara_Desc_Parametro                 As Integer = 2
Const conCOL_SonPara_Unidade                        As Integer = 3
Const conCOL_SonPara_Peso                           As Integer = 4
Const conCOL_SonPara_FormatString                   As String = "=Cód. Paramêtro|...|Descrição parâmetro|Unidade|Peso"
Const conColumnsIn_SonPara                          As Integer = 5

Const conCOL_SonParam_CodParam                      As Integer = 0
Const conCOL_SonParam_PesqParam                     As Integer = 1
Const conCOL_SonParam_Desc_Param                    As Integer = 2
Const conCOL_SonParam_UniMed                        As Integer = 3
Const conCOL_SonParam_ValParam                      As Integer = 4
Const conCOL_SonParam_ValParamPos                   As Integer = 5
Const conCOL_SonParam_ValParamNeg                   As Integer = 6
Const conCOL_SonParam_FormatString                  As String = "=Cód. Param|...|Descrição Parâmetros|Unidade|Parâmetro|Parâmetro(+)|Parâmetro(-)"
Const conColumnsIn_SonParam                         As Integer = 7

Const conCOL_SonUnid_CodUnid                        As Integer = 0
Const conCOL_SonUnid_PesqUnid                       As Integer = 1
Const conCOL_SonUnid_Unidade                        As Integer = 2
Const conCOL_SonUnid_Desc_Unid                      As Integer = 3
Const conCOL_SonUnid_Fator                          As Integer = 4
Const conCOL_SonUnid_FormatString                   As String = "=Cód. Unid|...|Unid.|Descrição Unidade|Fator"
Const conColumnsIn_SonUnid                          As Integer = 5

Const conCOL_SonVerniz_Item                         As Integer = 0
Const conCOL_SonVerniz_CodTipo                      As Integer = 1
Const conCOL_SonVerniz_PesqVerniz                   As Integer = 2
Const conCOL_SonVerniz_DescVerniz                   As Integer = 3
Const conCOL_SonVerniz_Codigo                       As Integer = 4
Const conCOL_SonVerniz_UnidMed                      As Integer = 5
Const conCOL_SonVerniz_Qtde                         As Integer = 6
Const conCOL_SonVerniz_FormatString                 As String = "=Item|Código|...|Descrição|Codigo|Un.Medida|Quantidade"
Const conColumnsIn_SonVerniz                        As Integer = 7

Const conCOL_SonCores_Codigo                        As Integer = 0
Const conCOL_SonCores_CodCor                        As Integer = 1
Const conCOL_SonCores_PesqCor                       As Integer = 2
Const conCOL_SonCores_DescCores                     As Integer = 3
Const conCOL_SonCores_Ordem                         As Integer = 4
Const conCOL_SonCores_UniMed                        As Integer = 5
Const conCOL_SonCores_Qtde                          As Integer = 6
Const conCOL_SonCores_FormatString                  As String = "=Cod|Código|...|Descrição|Ordem|Un.Medida|Qtde"
Const conColumnsIn_SonCores                         As Integer = 7

Const conCOL_SonVedante_Codigo                      As Integer = 0
Const conCOL_SonVedante_CodProd                     As Integer = 1
Const conCOL_SonVedante_PesqProd                    As Integer = 2
Const conCOL_SonVedante_DescProd                    As Integer = 3
Const conCOL_SonVedante_UniMed                      As Integer = 4
Const conCOL_SonVedante_Qtde                        As Integer = 5
Const conCOL_SonVedante_FormatString                As String = "=Cod|Código|...|Descrição|Uni.Medida|Qtde"
Const conColumnsIn_SonVedante                       As Integer = 6

Const conCOL_SomFecha_TampaPressao                  As Integer = 0
Const conCOL_SomFecha_BatoqueRetra                  As Integer = 1
Const conCOL_SomFecha_BatoquePlast                  As Integer = 2
Const conCOL_SomFecha_TampaVisor                    As Integer = 3

Const conCOL_SomFecha_Valor                         As Integer = 0
Const conCOL_SomFecha_TampaPressao_FormatString     As String = "=Tampa Pressão"
Const conCOL_SomFecha_BatoqueRetra_FormatString     As String = "=Batoque Retrátil"
Const conCOL_SomFecha_BatoquePlast_FormatString     As String = "=Batoque Plástico"
Const conCOL_SomFecha_TampaVisor_FormatString       As String = "=Tampa Visor"
Const conColumnsIn_SomFecha                         As Integer = 0

Const conCOL_SonProdAte_CodigoClie                  As Integer = 0
Const conCOL_SonProdAte_PesqClie                    As Integer = 1
Const conCOL_SonProdAte_DescClie                    As Integer = 2
Const conCOL_SonProdAte_FormatString                As String = "=Código|...|Descrição"
Const conColumnsIn_SonProdAte                       As Integer = 3

Const conCOL_SonEstoque_CodClie                     As Integer = 0
Const conCOL_SonEstoque_DescClie                    As Integer = 1
Const conCOL_SonEstoque_Entradas                    As Integer = 2
Const conCOL_SonEstoque_Saidas                      As Integer = 3
Const conCOL_SonEstoque_Saldo                       As Integer = 4
Const conCOL_SonEstoque_FormatString                As String = "=Cód.Cli|Razão Social|Entradas|Saidas|Saldo"
Const conColumnsIn_SonEstoque                       As Integer = 5

Const conCOL_SonEstoqueLit_CodClie                     As Integer = 0
Const conCOL_SonEstoqueLit_DescClie                    As Integer = 1
Const conCOL_SonEstoqueLit_EntradasKG                  As Integer = 2
Const conCOL_SonEstoqueLit_SaidasKG                    As Integer = 3
Const conCOL_SonEstoqueLit_SaldoKG                     As Integer = 4
Const conCOL_SonEstoqueLit_EntradasFolhas              As Integer = 5
Const conCOL_SonEstoqueLit_SaidasFolhas                As Integer = 6
Const conCOL_SonEstoqueLit_SaldoFolhas                 As Integer = 7
Const conCOL_SonEstoqueLit_FormatString                As String = "=Cód.Cli|Razão Social|Entradas KG|Saidas KG|Saldo KG|Entradas de Folhas|Saida de Folhas|Saldo de Folhas"
Const conColumnsIn_SonEstoqueLit                       As Integer = 8

Private Sub cboAlcaFerro_Validate(Cancel As Boolean)
    cboAlcaPlastica.ListIndex = -1
End Sub

Private Sub cboAlcaPlastica_Validate(Cancel As Boolean)
    cboAlcaFerro.ListIndex = -1
End Sub

Private Sub cboClass_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboClass, KeyAscii
End Sub


Private Sub cboUnidade_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboUnidade, KeyAscii
End Sub



Private Sub cboUniIns_Validate(Cancel As Boolean)
    If cboUniIns.ListIndex > -1 And lngINDLIST > 1 Then
        arrPROVARV(lngINDLIST).lngCodUniMed = cboUniIns.ItemData(cboUniIns.ListIndex)
        arrPROVARV(lngINDLIST).strUNIDADE = cboUniIns.Text
    End If
End Sub

Private Sub chkClientes_Click()
    Call InitGrdProdAtend
    Call CarregaClientes
End Sub

Private Sub cmdAbreArq_Click()

On Error GoTo err_Img

    
    cmoAbreArq.FileName = ""
    Call LoadPicture("")
    
    cmoAbreArq.ShowOpen
    
    If Len(Trim(cmoAbreArq.FileName)) = "" Then Exit Sub
    
    Image1.Picture = LoadPicture(cmoAbreArq.FileName)

    Exit Sub

err_Img:

        MsgBox "Erro nº   : " & Err.Number & vbCrLf & _
               "Descrição : " & Err.Description, vbOKOnly + vbExclamation, "Aviso"

End Sub

Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    
    Frame3.Enabled = True
    Frame38(3).Enabled = True
    Frame38(4).Enabled = True
    stProd.Enabled = True
    Frame14.Enabled = False
    Frame38(0).Enabled = True
    Frame38(1).Enabled = True
    Frame38(2).Enabled = False
    Frame38(5).Enabled = True
    Frame43(0).Enabled = True
    Frame43(1).Enabled = True
    Frame43(3).Enabled = True
    Frame6.Enabled = True
    Frame9.Enabled = True
    
    Frame41(1).Enabled = True
    
    stProd.Tab = 0
    
    txtCodigo.Enabled = True
   
    Me.Caption = "Cadastro de produtos - [ ALTERAÇÃO ]"
    
    cTipOper = "A"
    
    txtDescricao.SetFocus
    txtDescEquip.Locked = False
    
    txtCodigo.Enabled = False
    
    If objCADPRODUTO.PRODUTOTIPO = 0 Then
        txtCodigo.Enabled = True
        txtDigVerif.Enabled = False
    End If
    
End Sub

Private Sub cmdBtnFechaExc_Click(Index As Integer)
    If cTipOper = "I" Or cTipOper = "A" Then
       If Index = 0 Then Call objBLBFunc.ExclLinhaGrid(grdSomFecha(conCOL_SomFecha_TampaPressao), grdSomFecha(conCOL_SomFecha_TampaPressao).Row)
       If Index = 1 Then Call objBLBFunc.ExclLinhaGrid(grdSomFecha(conCOL_SomFecha_BatoqueRetra), grdSomFecha(conCOL_SomFecha_BatoqueRetra).Row)
       If Index = 2 Then Call objBLBFunc.ExclLinhaGrid(grdSomFecha(conCOL_SomFecha_BatoquePlast), grdSomFecha(conCOL_SomFecha_BatoquePlast).Row)
       If Index = 3 Then Call objBLBFunc.ExclLinhaGrid(grdSomFecha(conCOL_SomFecha_TampaVisor), grdSomFecha(conCOL_SomFecha_TampaVisor).Row)
    End If
End Sub

Private Sub cmdBtnFechaInc_Click(Index As Integer)
    If cTipOper = "I" Or cTipOper = "A" Then
       If Index = 0 Then Call IncRegSomFecha(conCOL_SomFecha_TampaPressao)
       If Index = 1 Then Call IncRegSomFecha(conCOL_SomFecha_BatoqueRetra)
       If Index = 2 Then Call IncRegSomFecha(conCOL_SomFecha_BatoquePlast)
       If Index = 3 Then Call IncRegSomFecha(conCOL_SomFecha_TampaVisor)
    End If
End Sub

Private Sub cmdClie_Click()

    ReDim arrCAMPOS(1 To 3, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "700"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_RAZAOSOC"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Razão Social"
    arrCAMPOS(2, 4) = "3000"
    arrCAMPOS(2, 5) = "SGI_RAZAOSOC"
    
    arrCAMPOS(3, 1) = "SGI_NOMFANTA"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "Nome Fantasia"
    arrCAMPOS(3, 4) = "3000"
    arrCAMPOS(3, 5) = "SGI_NOMFANTA"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Clientes", "CADCLIENTE.clsCADCLIENTE")
    
    If Len(Trim(varRETORNO)) > 0 Then
       txtCodCliente.Text = varRETORNO
       lblDesclie.Caption = PegaDescClie(CLng(txtCodCliente.Text))
    End If
    
    txtCodCliente.SetFocus

End Sub


Private Sub cmdExcCli_Click()
    If cTipOper = "I" Or cTipOper = "A" Then Call objBLBFunc.ExclLinhaGrid(grdClientes, grdClientes.Row)
End Sub

Private Sub cmdexcUnidConv_Click()
    If cTipOper = "I" Or cTipOper = "A" Then Call objBLBFunc.ExclLinhaGrid(grdUnidConv, grdUnidConv.Row)
End Sub

Private Sub cmdIncCli_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncRegGridProdAtend
End Sub

Private Sub cmdIncUnidConv_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncRegGridUnidades
End Sub

Private Sub cmdLinProd_Click()
    
    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODLIN"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "700"
    arrCAMPOS(1, 5) = "SGI_CODLIN"
    
    arrCAMPOS(2, 1) = "SGI_DESCRI"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "3000"
    arrCAMPOS(2, 5) = "SGI_DESCRI"
    
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Linha de Produto", "CADLINHAPROD.clsCADLINHAPROD")
    
    If Len(Trim(varRETORNO)) > 0 Then
       txtLinProd.Text = varRETORNO
       lblLinhProd.Caption = PegaDescLinProd(CLng(txtLinProd.Text))
       Call ConfGridFecham
       Call ConfComboFech
    End If
    
    txtLinProd.SetFocus

End Sub


Private Sub cmdPesqMont_Click(Index As Integer)

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    Dim strIDProd                   As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL      = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_PRODUTOTIPO = 0"
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "S"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1500"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Nome"
    arrCAMPOS(2, 4) = "5000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Produtos")
    
    If Len(Trim(varRETORNO)) > 0 Then
       strIDProd = PegaIDProduto(varRETORNO)
       If Index = 0 Then
          txtVerCat.Text = varRETORNO
          label11(15).Caption = PegaDescrProduto(varRETORNO)
          objCADPRODUTO.VerCat = IIf(IsNumeric(strIDProd) = True, CLng(strIDProd), 0)
       ElseIf Index = 1 Then
          txtColEsp.Text = varRETORNO
          label11(16).Caption = PegaDescrProduto(varRETORNO)
          objCADPRODUTO.ColEsp = IIf(IsNumeric(strIDProd) = True, CLng(strIDProd), 0)
       End If
    End If


End Sub


Private Sub CmdSalva_Click()

    Dim i           As Integer
    Dim j           As Integer
    Dim intResp     As Integer
    Dim strValor    As String
    Dim sValor      As String
    Dim lngQTDREGS  As Long
    Dim strEMPRESA  As String

    If ValidaCampos = True Then
       
       ''If cTipOper = "I" Then objCADPRODUTO.CODLCTO = objCADPRODUTO.Gera_Codigo("CARDEX")
       If cTipOper = "I" Then objCADPRODUTO.IDProduto = objCADPRODUTO.Gera_Codigo(Me.Name)
       
       
       objCADPRODUTO.IPI = "Null"
       If Len(Trim(txtIPI.Text)) > 0 Then objCADPRODUTO.IPI = txtIPI.Text
        
        If optTipProd(1).Value = True And _
           optEstProduto(0).Value = True Then
             objCADPRODUTO.CodLinProd = CLng(txtLinProd.Text)
             objCADPRODUTO.CodClie = CLng(txtCodCliente.Text)
             
             If cTipOper = "A" Then objCADPRODUTO.CodRotulo = CLng(txtCodRot.Text)
             
             objCADPRODUTO.DigVerif = CLng(txtDigVerif.Text)
             
             objCADPRODUTO.SEGUENCIA = "Null"
        End If
       
       If optTipProd(0).Value = True And _
          (optEstProduto(0).Value = True Or optEstProduto(4).Value = True) Then
           
           ''txtCodigo.Text = Constroi_Codigo_Comprado(txtCODGRUP.Text, txtCODSUBGRUP.Text)
           objCADPRODUTO.CodigoProd = Trim(Replace(Replace(txtCodigo.Text, " ", ""), ",", ""))
       
       End If
       
       objCADPRODUTO.DescriProd = txtDescricao.Text
       objCADPRODUTO.Unidade = cboUnidade.ItemData(cboUnidade.ListIndex)
       objCADPRODUTO.DESCEQUIP = txtDescEquip.Text
       objCADPRODUTO.PROCEDEN = txtProcedencia.Text
       objCADPRODUTO.CODPROFORNEC = txtCodProdFornec.Text
       
       If optQTDCORPPADRAOSN(0).Value = True Then objCADPRODUTO.QTDCORPSPADRAOSN = 0
       If optQTDCORPPADRAOSN(1).Value = True Then objCADPRODUTO.QTDCORPSPADRAOSN = 1
       
       objCADPRODUTO.COMPLEMENTO = "Null"
       If Len(Trim(txtComplemento.Text)) > 0 Then objCADPRODUTO.COMPLEMENTO = "'" & Trim(txtComplemento.Text) & "'"
       
       If optNatSimNao(1).Value = True Then objCADPRODUTO.NATURALSIMNAO = 1
       If optNatSimNao(0).Value = True Then objCADPRODUTO.NATURALSIMNAO = 0
       
       objCADPRODUTO.QTDEPORFOLHA = 0
       If Len(Trim(txtQtdePorFolha.Text)) > 0 Then objCADPRODUTO.QTDEPORFOLHA = CLng(txtQtdePorFolha.Text)
       objCADPRODUTO.QTDPASSADAS = 0
       If Len(Trim(txtQtdePassada.Text)) > 0 Then objCADPRODUTO.QTDPASSADAS = CLng(txtQtdePassada.Text)

       objCADPRODUTO.RUA = txtRua.Text
       objCADPRODUTO.box = txtBox.Text
       objCADPRODUTO.PRATELEIRA = txtPrateleira.Text
       objCADPRODUTO.CODEAN = txtCodigoEAN.Text
       
       If optTipProd(0).Value = True Then objCADPRODUTO.PRODUTOTIPO = 0
       If optTipProd(1).Value = True Then objCADPRODUTO.PRODUTOTIPO = 1
       
       If optEstProduto(0).Value = True Then objCADPRODUTO.ESTILOPROD = 0
       If optEstProduto(1).Value = True Then objCADPRODUTO.ESTILOPROD = 1
       If optEstProduto(2).Value = True Then objCADPRODUTO.ESTILOPROD = 2
       If optEstProduto(3).Value = True Then objCADPRODUTO.ESTILOPROD = 3
       If optEstProduto(4).Value = True Then objCADPRODUTO.ESTILOPROD = 4
       If optEstProduto(5).Value = True Then objCADPRODUTO.ESTILOPROD = 5
       
       If Len(Trim(txtPOCACRES.Text)) > 0 Then objCADPRODUTO.PORCACRES = CCur(txtPOCACRES.Text)
       If Len(Trim(txtPRCFINAL.Text)) > 0 Then objCADPRODUTO.PRCCUSTO = CCur(txtPRCFINAL.Text)
       
       If Len(Trim(txtSaldo.Text)) > 0 Then objCADPRODUTO.Saldo = CDbl(txtSaldo.Text)
       If Len(Trim(txtSaldo.Text)) = 0 Then objCADPRODUTO.Saldo = 0
       
       objCADPRODUTO.PESOUNIT = 0
       If Len(Trim(txtPeso.Text)) > 0 Then objCADPRODUTO.PESOUNIT = CCur(txtPeso.Text)
       
       If Len(Trim(txtEstMinimo.Text)) > 0 Then objCADPRODUTO.EstMin = CDbl(txtEstMinimo.Text)
       If Len(Trim(txtEstMinimo.Text)) = 0 Then objCADPRODUTO.EstMin = 0
       
       objCADPRODUTO.DataCadast = "'" & Format(CDate(mskDtCadastro.Text), "MM/DD/YYYY") & "'"
       If cboClass.ListIndex > -1 Then objCADPRODUTO.CLASFISC = cboClass.ItemData(cboClass.ListIndex)
       
       If optATUAUTOMSIMNAO(0).Value = True Then objCADPRODUTO.ATAUTOMSN = 0
       If optATUAUTOMSIMNAO(1).Value = True Then objCADPRODUTO.ATAUTOMSN = 1
       
       If Len(Trim(txtPRECOPRODUTO.Text)) > 0 Then objCADPRODUTO.PRCPROD = CCur(txtPRECOPRODUTO.Text)
       If Len(Trim(lblPRECOMEDIO.Caption)) > 0 Then objCADPRODUTO.PRCMED = CCur(lblPRECOMEDIO.Caption)
       
       If Len(Trim(txtDistancia.Text)) > 0 Then objCADPRODUTO.DISTANCIA = CCur(txtDistancia.Text)
       If Len(Trim(txtDistancia.Text)) = 0 Then objCADPRODUTO.DISTANCIA = 0
       
       objCADPRODUTO.CUBAGEN = 0
       If Len(Trim(txtCUBAGEN.Text)) > 0 Then objCADPRODUTO.CUBAGEN = CCur(txtCUBAGEN.Text)
       
       objCADPRODUTO.METROS = 0
       objCADPRODUTO.LARGURA = 0
       objCADPRODUTO.GRAMATURA2 = 0
       
       If Len(Trim(txtMETROS.Text)) > 0 Then objCADPRODUTO.METROS = CCur(txtMETROS.Text)
       If Len(Trim(txtLARGURA)) > 0 Then objCADPRODUTO.LARGURA = CCur(txtLARGURA.Text)
       If Len(Trim(txtGRAMATURAM2.Text)) > 0 Then objCADPRODUTO.GRAMATURA2 = CCur(txtGRAMATURAM2.Text)
       
       
       objCADPRODUTO.CODGRUPPROD = CLng(txtCODGRUP.Text)
       objCADPRODUTO.CODSUBGPROD = CLng(txtCODSUBGRUP.Text)
       objCADPRODUTO.EspProduto = CLng(txtEspecie.Text)
       objCADPRODUTO.TIPPRODUTO = CLng(txtTipo.Text)
       
       '' Especificação Tecnica ....
       objCADPRODUTO.ESPTECPROD = Empty
       
       '' Processos ....
       objCADPRODUTO.Processos = Empty
       
       '' Ferramentas ....
       objCADPRODUTO.FERRAMENTA = Empty
       
       '' Tipos de Orçamentos .....
       objCADPRODUTO.PRODTIPORCA = Empty
       
       '' Variante de Controle de Processo ............
       objCADPRODUTO.PRODVAPROC = Empty
       
       '' Variante de Recebimento ........
       objCADPRODUTO.PRODVAREC = Empty
        
       '' Variante de Inspeção Final ........
       objCADPRODUTO.PRODVAINSP = Empty
       '' Atributos de Processo ............
       objCADPRODUTO.PRODATRPROC = Empty
       
       '' Atributo de Recebimento ........
       objCADPRODUTO.PRODATRREC = Empty
       
       '' Atributo de Inspeção Final ........
       objCADPRODUTO.PRODATRINSP = Empty
       
       '' Aba Produção
       objCADPRODUTO.PORCCORR = (objCADPRODUTO.CORRELACAO * 100)
       
       '' ------------------------------------------------------
       '' Correlação ao Produto Padrão
       objCADPRODUTO.CORRPRODPADR = Empty
       '' ------------------------------------------------------
       '' Parametros para a correlação
       objCADPRODUTO.PARPROD = Empty
       
       '' Usar Palhet Padrão
       If optUsarPadrPalSN(0).Value = True Then objCADPRODUTO.PALLHETPADRAO = 0
       If optUsarPadrPalSN(1).Value = True Then objCADPRODUTO.PALLHETPADRAO = 1
       
       '' ------------------------------------------------------
       '' Gravando Unidades de Converção
       arrUNIDCONV = Empty
       If (grdUnidConv.Rows - 1) > 0 Then
           ReDim arrUNIDCONV(1 To (grdUnidConv.Rows - 1), 1 To 2) As Variant
           For i = 1 To (grdUnidConv.Rows - 1)
               arrUNIDCONV(i, 1) = CLng(grdUnidConv.Cell(flexcpText, i, conCOL_SonUnid_CodUnid))
               arrUNIDCONV(i, 2) = CDbl(grdUnidConv.Cell(flexcpText, i, conCOL_SonUnid_Fator))
           Next i
       End If
       objCADPRODUTO.UNIDCONV = arrUNIDCONV
       '' ------------------------------------------------------
       
       '' ------------------------------------------------------
       '' Cores
        Call objBLBFunc.removeLinhaVazia(grdCores, conCOL_SonCores_CodCor)
        arrCORES = Empty
        With grdCores
            If (.Rows - 1) > 0 Then
                ReDim arrCORES(1 To (.Rows - 1), 1 To 4) As String
                For i = 1 To (.Rows - 1)
                     arrCORES(i, 1) = .Cell(flexcpText, i, conCOL_SonCores_Codigo)
                     arrCORES(i, 2) = "Null"
                     If Len(Trim(.Cell(flexcpText, i, conCOL_SonCores_Ordem))) > 0 Then arrCORES(i, 2) = .Cell(flexcpText, i, conCOL_SonCores_Ordem)
                
                     arrCORES(i, 3) = "Null"
                     If Len(Trim(.Cell(flexcpText, i, conCOL_SonCores_UniMed))) > 0 Then arrCORES(i, 3) = Trim(.Cell(flexcpText, i, conCOL_SonCores_UniMed))
                     
                     sValor = "Null"
                     If Len(Trim(.Cell(flexcpText, i, conCOL_SonCores_Qtde))) > 0 Then
                        sValor = Replace(.Cell(flexcpText, i, conCOL_SonCores_Qtde), ".", "")
                        sValor = Replace(sValor, ",", ".")
                     End If
                     arrCORES(i, 4) = sValor
                Next i
            End If
        End With
        objCADPRODUTO.CORES = arrCORES
       
       '' ------------------------------------------------------
       '' Verniz 01
       arrVERNIZ = Empty
       If (grdVernizEsm.Rows - 1) > 0 Then
           If Len(Trim(grdVernizEsm.Cell(flexcpText, 1, conCOL_SonVerniz_Codigo))) > 0 Then
                ReDim arrVERNIZ(1 To 1, 1 To 3) As String
                arrVERNIZ(1, 1) = Trim(grdVernizEsm.Cell(flexcpText, 1, conCOL_SonVerniz_Codigo))
                
                arrVERNIZ(1, 2) = "Null"
                If Len(Trim(grdVernizEsm.Cell(flexcpText, 1, conCOL_SonVerniz_UnidMed))) > 0 Then arrVERNIZ(1, 2) = grdVernizEsm.Cell(flexcpText, 1, conCOL_SonVerniz_UnidMed)
                
                sValor = "Null"
                If Len(Trim(grdVernizEsm.Cell(flexcpText, 1, conCOL_SonVerniz_Qtde))) > 0 Then
                   sValor = Replace(grdVernizEsm.Cell(flexcpText, 1, conCOL_SonVerniz_Qtde), ".", "")
                   sValor = Replace(sValor, ",", ".")
                End If
                arrVERNIZ(1, 3) = sValor
           
           End If
       End If
       objCADPRODUTO.VERNIZ = arrVERNIZ
       '' ------------------------------------------------------
       
       '' Verniz 02
       arrVERNIZ02 = Empty
       If (grdVernizEsm.Rows - 1) > 0 Then
           If Len(Trim(grdVernizEsm.Cell(flexcpText, 2, conCOL_SonVerniz_Codigo))) > 0 Then
                ReDim arrVERNIZ02(1 To 1, 1 To 3) As String
                arrVERNIZ02(1, 1) = Trim(grdVernizEsm.Cell(flexcpText, 2, conCOL_SonVerniz_Codigo))
           
                arrVERNIZ02(1, 2) = "Null"
                If Len(Trim(grdVernizEsm.Cell(flexcpText, 2, conCOL_SonVerniz_UnidMed))) > 0 Then arrVERNIZ02(1, 2) = grdVernizEsm.Cell(flexcpText, 2, conCOL_SonVerniz_UnidMed)
                
                sValor = "Null"
                If Len(Trim(grdVernizEsm.Cell(flexcpText, 2, conCOL_SonVerniz_Qtde))) > 0 Then
                   sValor = Replace(grdVernizEsm.Cell(flexcpText, 2, conCOL_SonVerniz_Qtde), ".", "")
                   sValor = Replace(sValor, ",", ".")
                End If
                arrVERNIZ02(1, 3) = sValor
           End If
       End If
       objCADPRODUTO.VERNIZ02 = arrVERNIZ02
       
       '' ------------------------------------------------------
       '' Verniz Acabamento
       arrVERNIZACAB = Empty
       If (grdVernizEsm.Rows - 1) > 0 Then
           If Len(Trim(grdVernizEsm.Cell(flexcpText, 3, conCOL_SonVerniz_Codigo))) > 0 Then
                ReDim arrVERNIZACAB(1 To 1, 1 To 3) As String
                arrVERNIZACAB(1, 1) = Trim(grdVernizEsm.Cell(flexcpText, 3, conCOL_SonVerniz_Codigo))
           
                arrVERNIZACAB(1, 2) = "Null"
                If Len(Trim(grdVernizEsm.Cell(flexcpText, 3, conCOL_SonVerniz_UnidMed))) > 0 Then arrVERNIZACAB(1, 2) = grdVernizEsm.Cell(flexcpText, 3, conCOL_SonVerniz_UnidMed)
                
                sValor = "Null"
                If Len(Trim(grdVernizEsm.Cell(flexcpText, 3, conCOL_SonVerniz_Qtde))) > 0 Then
                   sValor = Replace(grdVernizEsm.Cell(flexcpText, 3, conCOL_SonVerniz_Qtde), ".", "")
                   sValor = Replace(sValor, ",", ".")
                End If
                arrVERNIZACAB(1, 3) = sValor
           End If
       End If
       objCADPRODUTO.VERNIZACAB = arrVERNIZACAB
       '' ------------------------------------------------------
       
       '' ------------------------------------------------------
       '' Esmalte
       arrESMALTE = Empty
       If (grdVernizEsm.Rows - 1) > 0 Then
           If Len(Trim(grdVernizEsm.Cell(flexcpText, 4, conCOL_SonVerniz_Codigo))) > 0 Then
                ReDim arrESMALTE(1 To 1, 1 To 3) As String
                arrESMALTE(1, 1) = Trim(grdVernizEsm.Cell(flexcpText, 4, conCOL_SonVerniz_Codigo))
                
                arrESMALTE(1, 2) = "Null"
                If Len(Trim(grdVernizEsm.Cell(flexcpText, 4, conCOL_SonVerniz_UnidMed))) > 0 Then arrESMALTE(1, 2) = grdVernizEsm.Cell(flexcpText, 4, conCOL_SonVerniz_UnidMed)
                
                sValor = "Null"
                If Len(Trim(grdVernizEsm.Cell(flexcpText, 4, conCOL_SonVerniz_Qtde))) > 0 Then
                   sValor = Replace(grdVernizEsm.Cell(flexcpText, 4, conCOL_SonVerniz_Qtde), ".", "")
                   sValor = Replace(sValor, ",", ".")
                End If
                arrESMALTE(1, 3) = sValor
           End If
       End If
       objCADPRODUTO.ESMALTE = arrESMALTE
       '' ------------------------------------------------------

       '' ------------------------------------------------------
       '' Verniz Externo
       arrVERNIZEXT = Empty
       If (grdVernizEsm.Rows - 1) > 0 Then
           If Len(Trim(grdVernizEsm.Cell(flexcpText, 5, conCOL_SonVerniz_Codigo))) > 0 Then
                ReDim arrVERNIZEXT(1 To 1, 1 To 3) As String
                arrVERNIZEXT(1, 1) = Trim(grdVernizEsm.Cell(flexcpText, 5, conCOL_SonVerniz_Codigo))
           
                arrVERNIZEXT(1, 2) = "Null"
                If Len(Trim(grdVernizEsm.Cell(flexcpText, 5, conCOL_SonVerniz_UnidMed))) > 0 Then arrVERNIZEXT(1, 2) = grdVernizEsm.Cell(flexcpText, 5, conCOL_SonVerniz_UnidMed)
                
                sValor = "Null"
                If Len(Trim(grdVernizEsm.Cell(flexcpText, 5, conCOL_SonVerniz_Qtde))) > 0 Then
                   sValor = Replace(grdVernizEsm.Cell(flexcpText, 5, conCOL_SonVerniz_Qtde), ".", "")
                   sValor = Replace(sValor, ",", ".")
                End If
                arrVERNIZEXT(1, 3) = sValor
           
           End If
       End If
       objCADPRODUTO.VERNIZEXT = arrVERNIZEXT
       '' ------------------------------------------------------


       objCADPRODUTO.VernCorpo = 0
       objCADPRODUTO.VernTampa = 0
       objCADPRODUTO.VernFundo = 0
       objCADPRODUTO.VernArgola = 0
       
       If cboCorpoVerniz.ListIndex > -1 Then objCADPRODUTO.VernCorpo = cboCorpoVerniz.ItemData(cboCorpoVerniz.ListIndex)
       If cboTampaVerniz.ListIndex > -1 Then objCADPRODUTO.VernTampa = cboTampaVerniz.ItemData(cboTampaVerniz.ListIndex)
       If cboFundoVerniz.ListIndex > -1 Then objCADPRODUTO.VernFundo = cboFundoVerniz.ItemData(cboFundoVerniz.ListIndex)
       If cboArgolaVerniz.ListIndex > -1 Then objCADPRODUTO.VernArgola = cboArgolaVerniz.ItemData(cboArgolaVerniz.ListIndex)
       
       objCADPRODUTO.EspessCorpo = 0
       objCADPRODUTO.EspessTampa = 0
       objCADPRODUTO.EspessFundo = 0
       objCADPRODUTO.EspessArgola = 0
       
       If Len(Trim(txtCorpoEspess.Text)) > 0 Then objCADPRODUTO.EspessCorpo = CLng(txtCorpoEspess.Text)
       If Len(Trim(txtTampaEspess.Text)) > 0 Then objCADPRODUTO.EspessTampa = CLng(txtTampaEspess.Text)
       If Len(Trim(txtFundoEspess.Text)) > 0 Then objCADPRODUTO.EspessFundo = CLng(txtFundoEspess.Text)
       If Len(Trim(txtArgolaEspess.Text)) > 0 Then objCADPRODUTO.EspessArgola = CLng(txtArgolaEspess.Text)
       
       '' =====================================
       objCADPRODUTO.RevestCorpo = 0
       objCADPRODUTO.RevestCorpo2 = 0
       objCADPRODUTO.RevestTampa = 0
       objCADPRODUTO.RevestTampa2 = 0
       objCADPRODUTO.RevestFundo = 0
       objCADPRODUTO.RevestFundo2 = 0
       objCADPRODUTO.RevestArgola = 0
       objCADPRODUTO.RevestArgola2 = 0
       
       If Len(Trim(txtCorpoRevest.Text)) > 0 Then objCADPRODUTO.RevestCorpo = CLng(txtCorpoRevest.Text)
       If Len(Trim(txtCorpoRevest2.Text)) > 0 Then objCADPRODUTO.RevestCorpo2 = CLng(txtCorpoRevest2.Text)
       If Len(Trim(txtTampaRevest.Text)) > 0 Then objCADPRODUTO.RevestTampa = CLng(txtTampaRevest.Text)
       If Len(Trim(txtTampaRevest2.Text)) > 0 Then objCADPRODUTO.RevestTampa2 = CLng(txtTampaRevest2.Text)
       If Len(Trim(txtFundoRevest.Text)) > 0 Then objCADPRODUTO.RevestFundo = CLng(txtFundoRevest.Text)
       If Len(Trim(txtFundoRevest2.Text)) > 0 Then objCADPRODUTO.RevestFundo2 = CLng(txtFundoRevest2.Text)
       If Len(Trim(txtArgolaRevest.Text)) > 0 Then objCADPRODUTO.RevestArgola = CLng(txtArgolaRevest.Text)
       If Len(Trim(txtArgolaRevest2.Text)) > 0 Then objCADPRODUTO.RevestArgola2 = CLng(txtArgolaRevest2.Text)
       
       objCADPRODUTO.FechTampaFuro = -1
       If cboFechTampaFuro.ListIndex > -1 Then objCADPRODUTO.FechTampaFuro = cboFechTampaFuro.ItemData(cboFechTampaFuro.ListIndex)
       
       
       '' =============================================
       '' Montagem
       objCADPRODUTO.AlcaPlastica = 0
       objCADPRODUTO.AlcaFerro = 0
       objCADPRODUTO.Pipeta = 0
       objCADPRODUTO.Azelha = 0
       objCADPRODUTO.VerCat = 0
       objCADPRODUTO.ColEsp = 0
       objCADPRODUTO.FechSoldaAgrafado = -1
       
       If cboAlcaPlastica.ListIndex > -1 Then objCADPRODUTO.AlcaPlastica = cboAlcaPlastica.ItemData(cboAlcaPlastica.ListIndex)
       If cboAlcaFerro.ListIndex > -1 Then objCADPRODUTO.AlcaFerro = cboAlcaFerro.ItemData(cboAlcaFerro.ListIndex)
       If optPipetSimNao(0).Value = True Then objCADPRODUTO.Pipeta = 0
       If optPipetSimNao(1).Value = True Then objCADPRODUTO.Pipeta = 1
       
       If optAzSimNao(0).Value = True Then objCADPRODUTO.Azelha = 0
       If optAzSimNao(1).Value = True Then objCADPRODUTO.Azelha = 1
       
       If cboFechSoldaAgraf.ListIndex > -1 Then objCADPRODUTO.FechSoldaAgrafado = cboFechSoldaAgraf.ItemData(cboFechSoldaAgraf.ListIndex)
       
       If Len(Trim(txtVerCat.Text)) > 0 Then objCADPRODUTO.VerCat = PegaIDProduto(txtVerCat.Text)
       If Len(Trim(txtColEsp.Text)) > 0 Then objCADPRODUTO.ColEsp = PegaIDProduto(txtColEsp.Text)
       
       
       '' =============================================
       '' Fechamento
       
       arrTAMPAPRES = Empty
       objCADPRODUTO.TampaPressao = conCOL_SomFecha_TampaPressao
       If (grdSomFecha(conCOL_SomFecha_TampaPressao).Rows - 1) > 0 Then
           ReDim arrTAMPAPRES(1 To (grdSomFecha(conCOL_SomFecha_TampaPressao).Rows - 1)) As Variant
           For i = 1 To (grdSomFecha(conCOL_SomFecha_TampaPressao).Rows - 1)
                If Len(Trim(grdSomFecha(conCOL_SomFecha_TampaPressao).Cell(flexcpText, i, conCOL_SomFecha_Valor))) > 0 Then
                    arrTAMPAPRES(i) = CCur(grdSomFecha(conCOL_SomFecha_TampaPressao).Cell(flexcpText, i, conCOL_SomFecha_Valor))
                End If
           Next i
       End If
       objCADPRODUTO.TAMPAPRESS = arrTAMPAPRES
        
       arrBATRETR = Empty
       objCADPRODUTO.BatoqueRetratil = conCOL_SomFecha_BatoqueRetra
       If (grdSomFecha(conCOL_SomFecha_BatoqueRetra).Rows - 1) > 0 Then
           ReDim arrBATRETR(1 To (grdSomFecha(conCOL_SomFecha_BatoqueRetra).Rows - 1)) As Variant
           For i = 1 To (grdSomFecha(conCOL_SomFecha_BatoqueRetra).Rows - 1)
                If Len(Trim(grdSomFecha(conCOL_SomFecha_BatoqueRetra).Cell(flexcpText, i, conCOL_SomFecha_Valor))) > 0 Then
                    arrBATRETR(i) = CCur(grdSomFecha(conCOL_SomFecha_BatoqueRetra).Cell(flexcpText, i, conCOL_SomFecha_Valor))
                End If
           Next i
       End If
       objCADPRODUTO.BATRETRATI = arrBATRETR
       
        
       arrBATPLAST = Empty
       objCADPRODUTO.BatoquePlastico = conCOL_SomFecha_BatoquePlast
       If (grdSomFecha(conCOL_SomFecha_BatoquePlast).Rows - 1) > 0 Then
           ReDim arrBATPLAST(1 To (grdSomFecha(conCOL_SomFecha_BatoquePlast).Rows - 1)) As Variant
           For i = 1 To (grdSomFecha(conCOL_SomFecha_BatoquePlast).Rows - 1)
                If Len(Trim(grdSomFecha(conCOL_SomFecha_BatoquePlast).Cell(flexcpText, i, conCOL_SomFecha_Valor))) > 0 Then
                    arrBATPLAST(i) = CCur(grdSomFecha(conCOL_SomFecha_BatoquePlast).Cell(flexcpText, i, conCOL_SomFecha_Valor))
                End If
           Next i
       End If
       objCADPRODUTO.BATPLASTIC = arrBATPLAST
        
       arrTAMPAVISOR = Empty
       objCADPRODUTO.TAMPAVIS = conCOL_SomFecha_TampaVisor
       If (grdSomFecha(conCOL_SomFecha_TampaVisor).Rows - 1) > 0 Then
           ReDim arrTAMPAVISOR(1 To (grdSomFecha(conCOL_SomFecha_TampaVisor).Rows - 1)) As Variant
           For i = 1 To (grdSomFecha(conCOL_SomFecha_TampaVisor).Rows - 1)
                If Len(Trim(grdSomFecha(conCOL_SomFecha_TampaVisor).Cell(flexcpText, i, conCOL_SomFecha_Valor))) > 0 Then
                    arrTAMPAVISOR(i) = CCur(grdSomFecha(conCOL_SomFecha_TampaVisor).Cell(flexcpText, i, conCOL_SomFecha_Valor))
                End If
           Next i
       End If
       objCADPRODUTO.TAMPAVISOR = arrTAMPAVISOR
        
        
       arrVEDANTE = Empty
       If (grdVedanteCompound.Rows - 1) > 0 Then
           ReDim arrVEDANTE(1 To (grdVedanteCompound.Rows - 1), 1 To 3) As String
           For i = 1 To (grdVedanteCompound.Rows - 1)
                If Len(Trim(grdVedanteCompound.Cell(flexcpText, i, conCOL_SonVedante_Codigo))) > 0 Then
                    arrVEDANTE(i, 1) = Trim(grdVedanteCompound.Cell(flexcpText, i, conCOL_SonVedante_Codigo))
                    
                    arrVEDANTE(i, 2) = "Null"
                    If Len(Trim(grdVedanteCompound.Cell(flexcpText, i, conCOL_SonVedante_UniMed))) > 0 Then arrVEDANTE(i, 2) = Trim(grdVedanteCompound.Cell(flexcpText, i, conCOL_SonVedante_UniMed))
                    
                    sValor = "Null"
                    If Len(Trim(grdVedanteCompound.Cell(flexcpText, i, conCOL_SonVedante_Qtde))) > 0 Then
                       sValor = Replace(grdVedanteCompound.Cell(flexcpText, i, conCOL_SonVedante_Qtde), ".", "")
                       sValor = Replace(sValor, ",", ".")
                    End If
                    arrVEDANTE(i, 3) = sValor
                
                End If
           Next i
       End If
       objCADPRODUTO.VEDANTE = arrVEDANTE
       
       '' =============================================
       '' Expedição
       objCADPRODUTO.TipoPalletGranel = 0
       objCADPRODUTO.QtdePalletGranel = 0
       
       If optTipPalet(0).Value = True Then objCADPRODUTO.TipoPalletGranel = 0
       If optTipPalet(1).Value = True Then objCADPRODUTO.TipoPalletGranel = 1
       If Len(Trim(txtQtdeEmb.Text)) > 0 Then objCADPRODUTO.QtdePalletGranel = CDbl(txtQtdeEmb.Text)
       
       '' =============================================
       '' Cliente que este produto atende
       arrPRODATECLIE = Empty
       If (grdClientes.Rows - 1) > 0 Then
           ReDim arrPRODATECLIE(1 To (grdClientes.Rows - 1)) As Variant
           For i = 1 To (grdClientes.Rows - 1)
                If Len(Trim(grdClientes.Cell(flexcpText, i, conCOL_SonProdAte_CodigoClie))) > 0 Then
                    arrPRODATECLIE(i) = CLng(grdClientes.Cell(flexcpText, i, conCOL_SonProdAte_CodigoClie))
                End If
           Next i
       End If
       objCADPRODUTO.PRODATECLIE = arrPRODATECLIE
       '' =============================================
       
       If optLaudoSN(0).Value = True Then objCADPRODUTO.EMITLAUDO = 0
       If optLaudoSN(1).Value = True Then objCADPRODUTO.EMITLAUDO = 1
       '' =============================================
       
       '' Familias de Máquinas
        objCADPRODUTO.FAMMAQ = Empty
        Call objBLBFunc.removeLinhaVazia(grdFamMaq, conCOL_SonFamMaq_CodFamMaq)
       
        With grdFamMaq
             If (.Rows - 1) > 0 Then
             
                 ReDim arrFAMMAQ(1 To (.Rows - 1)) As String
                 For i = 1 To (.Rows - 1)
                     arrFAMMAQ(i) = .Cell(flexcpText, i, conCOL_SonFamMaq_CodFamMaq)
                 Next i
                 objCADPRODUTO.FAMMAQ = arrFAMMAQ
             End If
        End With
       
        '' Alça de Galão
        If optAlcGalSN(0).Value = True Then objCADPRODUTO.ALCAGALAO = 0
        If optAlcGalSN(1).Value = True Then objCADPRODUTO.ALCAGALAO = 1
       
        '' Neck IN
        If optNeckInSN(0).Value = True Then objCADPRODUTO.NECKIN = 0
        If optNeckInSN(1).Value = True Then objCADPRODUTO.NECKIN = 1
       
        If optFilialPed(0).Value = True Then objCADPRODUTO.FILIALPED = 0
        If optFilialPed(1).Value = True Then objCADPRODUTO.FILIALPED = 1
       
        objCADPRODUTO.CAMINHO = "Null"
        If Len(Trim(cmoAbreArq.FileName)) > 0 Then objCADPRODUTO.CAMINHO = "'" & Trim(cmoAbreArq.FileName) & "'"
       
        If optDimCortePAD(0).Value = True Then objCADPRODUTO.DIMPADRAO = 0
        If optDimCortePAD(1).Value = True Then objCADPRODUTO.DIMPADRAO = 1
       
        strValor = "Null"
        If Len(Trim(txtDESENV.Text)) > 0 Then
            strValor = Replace(Format(txtDESENV.Text, "#,##0.00"), ".", "")
            strValor = Replace(strValor, ",", ".")
        End If
        objCADPRODUTO.DESENV = strValor
        
        strValor = "Null"
        If Len(Trim(txtALTURA.Text)) > 0 Then
            strValor = Replace(Format(txtALTURA.Text, "#,##0.00"), ".", "")
            strValor = Replace(strValor, ",", ".")
        End If
        objCADPRODUTO.ALTURA = strValor
       
        If optFilialPed(0).Value = True Then objCADPRODUTO.FILIALPED = 0
        If optFilialPed(1).Value = True Then objCADPRODUTO.FILIALPED = 1
        
       objCADPRODUTO.CODPRODENTR = "Null"
       
       '' -----------------------------------
       '' Recursos de Entrada
       arrPRODENTR = Empty
       With grdPRODENTR
            If (.Rows - 1) > 0 Then
                 ReDim arrPRODENTR(1 To (.Rows - 1), 1 To 3) As String
                 For i = 1 To (.Rows - 1)
                     arrPRODENTR(i, 1) = .Cell(flexcpText, i, conCOL_SonProdEntr_IDProduto)
                     
                     arrPRODENTR(i, 2) = "Null"
                     If Len(Trim(.Cell(flexcpText, i, conCOL_SonProdEntr_UniMed))) > 0 Then arrPRODENTR(i, 2) = .Cell(flexcpText, i, conCOL_SonProdEntr_UniMed)
                 
                     sValor = "Null"
                     If Len(Trim(.Cell(flexcpText, i, conCOL_SonProdEntr_Qtde))) > 0 Then
                        sValor = Replace(.Cell(flexcpText, i, conCOL_SonProdEntr_Qtde), ".", "")
                        sValor = Replace(sValor, ",", ".")
                     End If
                     arrPRODENTR(i, 3) = sValor
                 Next i
            End If
       End With
       objCADPRODUTO.PRODENTR = arrPRODENTR
       '' -----------------------------------
       
       If cTipOper = "I" Then
          If optProdNovoSN(1).Value = True Then
             If optTipProd(0).Value = True Then objCADPRODUTO.STATUS = 1
             If optTipProd(1).Value = True Then objCADPRODUTO.STATUS = 2
             objCADPRODUTO.PRODNOVO = 1
          ElseIf optProdNovoSN(0).Value = True Then
             objCADPRODUTO.STATUS = 1
             objCADPRODUTO.PRODNOVO = 0
          End If
       Else
          If optProdNovoSN(1).Value = True Then objCADPRODUTO.PRODNOVO = 1
          If optProdNovoSN(0).Value = True Then objCADPRODUTO.PRODNOVO = 0
       End If
       
       '' ------------------------------
       strEMPRESA = ""
       If optFilialPed(1).Value = True Then strEMPRESA = "_STEEL"
       
       arrPEDIDOS = Empty
       sSql = ""
       
       sSql = "Select " & vbCrLf
       sSql = sSql & "       SGI_CADPEDVENDH" & strEMPRESA & ".SGI_CODIGO" & vbCrLf
       sSql = sSql & "  From " & vbCrLf
       sSql = sSql & "       SGI_CADPEDVENDH" & strEMPRESA & " SGI_CADPEDVENDH" & strEMPRESA & vbCrLf
       sSql = sSql & "      ,SGI_CADPEDVENDI" & strEMPRESA & " SGI_CADPEDVENDI" & strEMPRESA & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       SGI_CADPEDVENDH" & strEMPRESA & ".SGI_FILIAL    = " & FILIAL & vbCrLf
       sSql = sSql & "   And SGI_CADPEDVENDH" & strEMPRESA & ".SGI_STATUS    = 'V'" & vbCrLf
       sSql = sSql & "   And SGI_CADPEDVENDI" & strEMPRESA & ".SGI_FILIAL    = SGI_CADPEDVENDH" & strEMPRESA & ".SGI_FILIAL" & vbCrLf
       sSql = sSql & "   And SGI_CADPEDVENDI" & strEMPRESA & ".SGI_CODIGO    = SGI_CADPEDVENDH" & strEMPRESA & ".SGI_CODIGO" & vbCrLf
       sSql = sSql & "   And SGI_CADPEDVENDI" & strEMPRESA & ".SGI_IDPRODUTO = " & objCADPRODUTO.IDProduto
       
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
       If Not BREC.EOF() Then
          lngQTDREGS = 0
          Do While Not BREC.EOF()
             lngQTDREGS = lngQTDREGS + 1
             BREC.MoveNext
          Loop
          If lngQTDREGS > 0 Then
                ReDim arrPEDIDOS(1 To lngQTDREGS, 1 To 1) As String
                BREC.MoveFirst
                lngQTDREGS = 1
                Do While Not BREC.EOF()
                    arrPEDIDOS(lngQTDREGS, 1) = BREC!SGI_CODIGO
                    BREC.MoveNext
                Loop
          End If
       End If
       BREC.Close
       objCADPRODUTO.PEDIDOS = arrPEDIDOS
       '' ------------------------------
       
       If cTipOper = "I" Then
            objCADPRODUTO.FOTALTSN = 0  '' Alteração no fotolito
            objCADPRODUTO.PRODNOVO = 1  '' Produto Novo
       End If
       
       If cTipOper <> "I" Then
            '' ===================================
            '' Gerando Codigos Da Estrutura
            Dim lngCODPAI As Long
            For i = 1 To UBound(arrPROVARV)
                 
                 If arrPROVARV(i).lngProdutoID > 0 Then
                 
                 lngCODPAI = 0
                 If i > 1 Then lngCODPAI = arrPROVARV(treListaMat.Nodes(i).Parent.Index).lngCODIGO
                    
                 arrPROVARV(i).lngCODIGO = objCADPRODUTO.Gera_Codigo(Me.Name & "_LISTMAT")
                 arrPROVARV(i).lngCODPAI = lngCODPAI
                 
                 sValor = "Null"
                 If Len(Trim(arrPROVARV(i).strQTDCONS)) > 0 Then
                    sValor = Replace(arrPROVARV(i).strQTDCONS, ".", "")
                    sValor = Replace(sValor, ",", ".")
                 End If
                 arrPROVARV(i).strQTDCONS = sValor
                 
                 End If
                 
            Next i
       End If
        ''For I = 1 To UBound(arrPROVARV)
        ''     For j = 1 To UBound(arrPROVARV)
        ''         If arrPROVARV(j).lngIDPai = arrPROVARV(I).lngID Then arrPROVARV(j).lngCODPAI = arrPROVARV(I).lngCODIGO
        ''     Next j
        ''Next I
        
        '' ===================================
       
       If objCADPRODUTO.GRAVA(cTipOper) = True Then
          
            If optTipProd(1).Value = True And optEstProduto(0).Value = True Then
                 If cTipOper = "I" Then
                    objCADPRODUTO.CodRotulo = objCADPRODUTO.GeraCodRotulo
                    Call objCADPRODUTO.UpdateRotulo(cTipOper)
                 End If
            
                objCADPRODUTO.CodigoProd = Format(objCADPRODUTO.CodLinProd, "###000") & "." & _
                                           Format(objCADPRODUTO.CodClie, "####0000") & "." & _
                                           Format(objCADPRODUTO.CodRotulo, "##00") & "." & _
                                           Format(objCADPRODUTO.DigVerif, "#0")
                                           
                Call objCADPRODUTO.UpdateCodProd(cTipOper, objCADPRODUTO.CodigoProd)
            
            
            End If
          
          
            MsgBox "O Produto foi " & IIf(cTipOper = "I", "incluso", IIf(cTipOper = "A", "alterado", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
            If objCADPRODUTO.Atualiza(cTipOper, Trim(objCADPRODUTO.IDProduto), FILIAL, Me.Name) = False Then Exit Sub
                
                Dim lngCodLog As Long
                lngCodLog = objCADPRODUTO.Gera_Codigo("SGI_LOGMODULO")
                Call objCADPRODUTO.GravaLogModulo(FILIAL, lngCodLog, Me.Name, cTipOper, lngCodUsuario, Str(objCADPRODUTO.IDProduto))
            
                
                '' ==================================================
                '' Imagem
                If Len(Trim(strNOMARG)) > 0 Then
                    ''sSql = "Select SGI_IMAGEM From SGI_CADPRODUTO Where SGI_FILIAL = " & FILIAL & " And SGI_IDPRODUTO = " & objCADPRODUTO.IDProduto
                    ''BREC2.CursorType = adOpenDynamic
                    ''BREC2.LockType = adLockOptimistic
                    ''BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
                    ''If Not BREC2.EOF Then
                    ''   Call objBLBFunc.GravaBlobParaBanco(BREC2, "SGI_IMAGEM", strNOMARG)
                    ''   BREC2.Update
                    ''End If
                    ''BREC2.Close
                    '' ------------------------------
                End If
          
          'If cTipOper = "A" Then
          
          '   If objCADPRODUTO.Update_ListaMat(cTipOper) = False Then
          '      MsgBox "Não foi possivel atualizar as listas deste produto !!!", vbOKOnly + vbExclamation, "Aviso"
          '   Else
          '      If objCADPRODUTO.RecalcEstrutura = False Then
          '         MsgBox "Não foi possivel atualizar as listas deste produto !!!", vbOKOnly + vbExclamation, "Aviso"
          '      End If
          '   End If
          '
          'End If
           Unload Me
          
       End If
       
    End If
    
End Sub


Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncRegGridProdEntr
End Sub

Private Sub Command10_Click()
    If cTipOper = "A" Then Call Inseri_Item(lngINDLIST)
End Sub

Private Sub Command11_Click()

    
    If cTipOper = "C" Then Exit Sub
    
    ReDim arrCAMPOS(1 To 3, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       PRO.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "       ,PRO.SGI_CODIGO" & vbCrLf
    sSql = sSql & "       ,PRO.SGI_CODCLIE" & vbCrLf
    sSql = sSql & "       ,PRO.SGI_DESCRICAO" & vbCrLf
    sSql = sSql & "       ,PRO.SGI_COMPLEMENTO" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO PRO" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       PRO.SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "S"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "1500"
    arrCAMPOS(1, 5) = "PRO.SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_COMPLEMENTO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Complemento"
    arrCAMPOS(2, 4) = "2500"
    arrCAMPOS(2, 5) = "PRO.SGI_COMPLEMENTO"
    
    arrCAMPOS(3, 1) = "SGI_DESCRICAO"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "Descrição"
    arrCAMPOS(3, 4) = "5000"
    arrCAMPOS(3, 5) = "PRO.SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Produtos")
            
    If Len(Trim(varRETORNO)) > 0 Then
    
        lblDescProd.Caption = PegaDescrProduto(varRETORNO)
        txtCODPROD.Text = varRETORNO
        
        Call CriaArray
        
    End If
            
End Sub


Private Sub Command2_Click()
    If cTipOper = "I" Or cTipOper = "A" Then Call objBLBFunc.ExclLinhaGrid(grdPRODENTR, grdPRODENTR.Row)
End Sub

Private Sub Command26_Click()
    If cTipOper = "I" Or cTipOper = "A" Then
       Call objBLBFunc.ExclLinhaGrid(grdCores, grdCores.Row)
       Call Refaz_IndCores
    End If
End Sub

Private Sub Command27_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncRegGridCores
End Sub

Private Sub Command28_Click()
    If cTipOper = "I" Or cTipOper = "A" Then Call objBLBFunc.ExclLinhaGrid(grdVedanteCompound, grdVedanteCompound.Row)
End Sub

Private Sub Command29_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncRegGridVedante
End Sub

Private Sub Command3_Click()

    If Len(Trim(txtCODSUBGRUP.Text)) > 0 Then
    
       ReDim arrCAMPOS(1 To 2, 1 To 5) As String
       ReDim arrTABELA(1 To 1) As String
    
       sSql = "Select " & vbCrLf
       sSql = sSql & "       ESPECI.SGI_CODIGO    " & vbCrLf
       sSql = sSql & "      ,ESPECI.SGI_DESCRICAO " & vbCrLf
       sSql = sSql & "  from " & vbCrLf
       sSql = sSql & "       SGI_SUBGRUPRODITEN SUBGRP " & vbCrLf
       sSql = sSql & "      ,SGI_CADESPPROD     ESPECI " & vbCrLf
       sSql = sSql & "Where " & vbCrLf
       sSql = sSql & "      SUBGRP.SGI_CODIGO = " & txtCODSUBGRUP.Text & vbCrLf
       sSql = sSql & "  And SUBGRP.SGI_FILIAL = " & FILIAL & vbCrLf
       sSql = sSql & "  And ESPECI.SGI_FILIAL = SUBGRP.SGI_FILIAL " & vbCrLf
       sSql = sSql & "  And ESPECI.SGI_CODIGO = SUBGRP.SGI_CODESPECIE "
    
       arrTABELA(1) = sSql
    
       arrCAMPOS(1, 1) = "SGI_CODIGO"
       arrCAMPOS(1, 2) = "N"
       arrCAMPOS(1, 3) = "Código"
       arrCAMPOS(1, 4) = "700"
       arrCAMPOS(1, 5) = "ESPECI.SGI_CODIGO"
    
       arrCAMPOS(2, 1) = "SGI_DESCRICAO"
       arrCAMPOS(2, 2) = "S"
       arrCAMPOS(2, 3) = "Descrição"
       arrCAMPOS(2, 4) = "3000"
       arrCAMPOS(2, 5) = "ESPECI.SGI_DESCRICAO"
    
       varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Especie de produtos", "CADESPPROD.clsCADESPPROD")
    
       If Len(Trim(varRETORNO)) > 0 Then
          txtEspecie.Text = varRETORNO
          Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADESPPROD", varRETORNO, lblDescEspecieProd, True)
          Call ConfCombosVerniz
       End If
       
        txtEspecie.SetFocus
       
    End If


End Sub

Private Sub Command30_Click()
    If cTipOper = "I" Or cTipOper = "A" Then Call objBLBFunc.ExclLinhaGrid(grdFamMaq, grdFamMaq.Row)
End Sub

Private Sub Command31_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncRegGridFamMaq
End Sub

Private Sub Command4_Click()

    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADTIPPROD " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "700"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "3000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Tipos de produtos", "CADTIPPRO.clsCADTIPPRO")
    
    If Len(Trim(varRETORNO)) > 0 Then
       txtTipo.Text = varRETORNO
        Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADTIPPROD", varRETORNO, lblDescTipoProd, True)
    End If
    
    txtTipo.SetFocus

End Sub


Private Sub Command5_Click()

    
    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADGRUPROD " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "700"
    arrCAMPOS(1, 5) = "SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "3000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Familia de produtos", "CADGRUPPRO.clsCADGRUPPRO")
    
    If Len(Trim(varRETORNO)) > 0 Then
       txtCODGRUP.Text = varRETORNO
       Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADGRUPROD", varRETORNO, lblDescFamProd, True)
    End If

    lblDescSubFamProd.Caption = ""
    txtCODSUBGRUP.Text = ""
    
    lblDescEspecieProd.Caption = ""
    txtEspecie.Text = ""

    txtCODGRUP.SetFocus

    If optTipProd(0).Value = True Then txtCodigo.Text = Constroi_Codigo_Comprado(txtCODGRUP.Text, txtCODSUBGRUP.Text)

End Sub

Private Sub Command6_Click()
    
    If Len(Trim(txtCODGRUP.Text)) > 0 Then
    
        ReDim arrCAMPOS(1 To 2, 1 To 5) As String
        ReDim arrTABELA(1 To 1) As String
    
        sSql = "Select " & vbCrLf
        sSql = sSql & "       SUBGR.SGI_CODIGO     " & vbCrLf
        sSql = sSql & "      ,SUBGR.SGI_DESCRICAO  " & vbCrLf
        sSql = sSql & "  from " & vbCrLf
        sSql = sSql & "       SGI_GRUPPRODITEN GRUPO" & vbCrLf
        sSql = sSql & "      ,SGI_CADSUBGRPROD SUBGR" & vbCrLf
        sSql = sSql & "Where " & vbCrLf
        sSql = sSql & "      GRUPO.SGI_CODIGO = " & txtCODGRUP.Text
        sSql = sSql & "  And GRUPO.SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "  And SUBGR.SGI_FILIAL = GRUPO.SGI_FILIAL " & vbCrLf
        sSql = sSql & "  And SUBGR.SGI_CODIGO = GRUPO.SGI_CODESPECIE " & vbCrLf
    
        arrTABELA(1) = sSql
    
        arrCAMPOS(1, 1) = "SGI_CODIGO"
        arrCAMPOS(1, 2) = "N"
        arrCAMPOS(1, 3) = "Código"
        arrCAMPOS(1, 4) = "700"
        arrCAMPOS(1, 5) = "SUBGR.SGI_CODIGO"
    
        arrCAMPOS(2, 1) = "SGI_DESCRICAO"
        arrCAMPOS(2, 2) = "S"
        arrCAMPOS(2, 3) = "Descrição"
        arrCAMPOS(2, 4) = "3000"
        arrCAMPOS(2, 5) = "SUBGR.SGI_DESCRICAO"
    
        varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Sub-familia de produtos", "CADSUBGRPRO.clsCADSUBGRPRO")
    
        If Len(Trim(varRETORNO)) > 0 Then
           txtCODSUBGRUP.Text = varRETORNO
           Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADSUBGRPROD", varRETORNO, lblDescSubFamProd, True)
        End If
        
        lblDescEspecieProd.Caption = ""
        txtEspecie.Text = ""
        txtCODSUBGRUP.SetFocus
    
        If optTipProd(0).Value = True Then txtCodigo.Text = Constroi_Codigo_Comprado(txtCODGRUP.Text, txtCODSUBGRUP.Text)
    
    End If
    
End Sub

Private Sub Command7_Click()
    Image1.Picture = LoadPicture("")
End Sub

Private Sub Command8_Click()

    Dim arrArquivo As Variant
    Dim i As Integer
    
    arrArquivo = Split(cmoAbreArq.FileName, "\")
    
    If IsArray(arrArquivo) Then
        strCAMINHO = ""
        For i = 0 To (UBound(arrArquivo) - 1)
            strCAMINHO = strCAMINHO & Trim(arrArquivo(i)) + "\"
        Next i
        
        If UBound(arrArquivo) > 0 Then strNOMARG = Trim(arrArquivo(UBound(arrArquivo)))
        
        If Len(Trim(strNOMARG)) > 0 Then
        
            sSql = ""
            
            sSql = "Select SGI_IMAGEM From SGI_CADPROD_IMAGEN Where SGI_FILIAL = " & FILIAL & " And SGI_CODIGO = " & objCADPRODUTO.IDProduto
            BREC2.Open sSql, adoBanco_Dados_Imagem, adOpenDynamic, adLockOptimistic
            If Not BREC2.EOF Then
               Call objBLBFunc.GravaBlobParaBanco(BREC2, "SGI_IMAGEM", strCAMINHO & strNOMARG)
               BREC2.Update
            End If
            BREC2.Close
            '' ------------------------------
            
        End If
        
        
    End If

End Sub

Private Sub Command9_Click()
        If cTipOper = "C" Or cTipOper = "I" Then Exit Sub
        If lngINDLIST > 0 And lngINDLIST > 1 Then
            Call Exclui_Item_Arvore
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
   Set objCADPRODUTO = CreateObject("CADPRODU.clsCADPRODU")
   Set objCADFORN = CreateObject("CADFORNEC.clsCADFORNEC")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
   objCADPRODUTO.FILIAL = FILIAL
   
    Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)
    Set adoBanco_Dados_Imagem = objBLBFunc.Banco_Dados_Imagem(Linha)
    
    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
   
    If adoBanco_Dados_Imagem.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
   
    strCamImgRotulos = Right(Linha(8), Len(Trim(Linha(8))) - 7)
   
   optEstProduto(1).Visible = False
   optEstProduto(2).Visible = False
   optEstProduto(3).Visible = False
   optEstProduto(5).Visible = False
       
    lngINDLIST = 0
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame38(3).Enabled = True
    Frame38(4).Enabled = True
    Frame38(0).Enabled = True
    Frame38(1).Enabled = True
    Frame38(2).Enabled = True
    Frame43(0).Enabled = True
    Frame43(1).Enabled = True
    Frame43(3).Enabled = True
    Frame8.Enabled = False
    Frame6.Enabled = True
    Frame9.Enabled = True
    
    Call AbilDesDadosListMat
    
    chkClientes.Value = 0
    
    Frame41(1).Enabled = True
    txtCodRot.Enabled = False
    
    label11(15).Caption = ""
    label11(16).Caption = ""
    
    Label1(23).Caption = ""
    
    txtCodigo.Enabled = True
   
    Me.Caption = "Cadastro de produtos - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADPROD
    txtDescEquip.Locked = False
    LimpaCaptions
    
    objCADPRODUTO.PreenchComboUnidade cboUnidade
    objCADPRODUTO.PreenchComboClaFis cboClass
    objCADPRODUTO.PreenchComboUnidade cboUniIns
    cboUniIns.ListIndex = -1
    
    txtCodigo.Text = ""
    lblPRECOMEDIO.Caption = ""
    
    mskDtCadastro.Text = Format(Date, "DD/MM/YYYY")
    
    cboUnidade.ListIndex = -1
   
    optATUAUTOMSIMNAO(1).Value = True
    optTipProd(1).Value = True
    optEstProduto(0).Value = True
    optDimCortePAD(1).Value = True
    optQTDCORPPADRAOSN(1).Value = True
    
    cboFechSoldaAgraf.ListIndex = -1
    optProdNovoSN(1).Value = True
    
    ConfGridFecham
    
    InitGridUnidConv
    
    Call InitGridVernizEsm
    Call InitGridCores
    Call InitGrdVedanteCompound
    Call ConfCombosVerniz
    Call PreenchComboAlca
    Call InitgrdSomFecha
    Call InitGrdProdAtend
    Call InitGrdFamMaq
    Call ConfEstoque
    Call ConfProdEntr
    Call ConfEstoqueLit
            
        
    Call ConfComboFech
        
        
    stProd.Tab = 0
    
    txtCodigo.Enabled = False
    Frame38(2).Enabled = True
    
    optNatSimNao(0).Value = True
        
    objCADPRODUTO.PRODUTOTIPO = 2
    Label1(23).Caption = "AG.LIBERAÇÃO"
    
    optLaudoSN(0).Value = True
    optAlcGalSN(0).Value = True
    optNeckInSN(0).Value = True

    Call LimpaCamposLabel

    optFilialPed(0).Value = True

    objCADPRODUTO.STATUS = 1

    optUsarPadrPalSN(0).Value = True
    Frame17.Enabled = False
    
    ''Call Cria_Item_Pai
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call DestroiObjeto
End Sub


Private Sub grdClientes_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonProdAte_DescClie
         Cancel = True
    Case conCOL_SonProdAte_CodigoClie, _
         conCOL_SonProdAte_PesqClie
         If cTipOper = "C" Then Cancel = True
    Case Else
        grdClientes.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub grdClientes_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    If cTipOper = "C" Then Exit Sub
    
    If (grdCores.Rows - 1) = 0 Then Exit Sub
    
    ReDim arrCAMPOS(1 To 4, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    Select Case Col
        Case conCOL_SonProdAte_PesqClie
    
            sSql = "Select " & vbCrLf
            sSql = sSql & "       * " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       SGI_FILIAL      = " & FILIAL & vbCrLf
            
            arrTABELA(1) = sSql
            
            arrCAMPOS(1, 1) = "SGI_CODIGO"
            arrCAMPOS(1, 2) = "S"
            arrCAMPOS(1, 3) = "Código"
            arrCAMPOS(1, 4) = "1500"
            arrCAMPOS(1, 5) = "SGI_CODIGO"
            
            arrCAMPOS(2, 1) = "SGI_CPFCNPJ"
            arrCAMPOS(2, 2) = "S"
            arrCAMPOS(2, 3) = "CPF/CNPJ"
            arrCAMPOS(2, 4) = "1500"
            arrCAMPOS(2, 5) = "SGI_CPFCNPJ"
            
            arrCAMPOS(3, 1) = "SGI_NOMFANTA"
            arrCAMPOS(3, 2) = "S"
            arrCAMPOS(3, 3) = "Nome Fantazia"
            arrCAMPOS(3, 4) = "3000"
            arrCAMPOS(3, 5) = "SGI_NOMFANTA"
            
            arrCAMPOS(4, 1) = "SGI_RAZAOSOC"
            arrCAMPOS(4, 2) = "S"
            arrCAMPOS(4, 3) = "Razão Social"
            arrCAMPOS(4, 4) = "5000"
            arrCAMPOS(4, 5) = "SGI_RAZAOSOC"
            
            varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Clientes", "CADCLIENTE.clsCADCLIENTE")
            
            If Len(Trim(varRETORNO)) > 0 Then
               grdClientes.Cell(flexcpText, Row, conCOL_SonProdAte_CodigoClie) = varRETORNO
               grdClientes.Cell(flexcpText, Row, conCOL_SonProdAte_DescClie) = PegaDescrCliente(CLng(varRETORNO))
            End If
            
            If objBLBFunc.FcVerifItensRepetidos(grdClientes, Row, conCOL_SonProdAte_CodigoClie, varRETORNO) = False Then
               MsgBox "Este Cliente ja foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
               grdClientes.Cell(flexcpText, Row, conCOL_SonProdAte_CodigoClie) = ""
               grdClientes.Cell(flexcpText, Row, conCOL_SonProdAte_DescClie) = ""
               Exit Sub
            End If

    End Select

End Sub

Private Sub grdClientes_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdClientes
          Select Case Col
                    Case conCOL_SonProdAte_CodigoClie
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
                         ''KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
          End Select
     End With
End Sub

Private Sub grdClientes_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
     With grdClientes
          Select Case Col
                 Case conCOL_SonProdAte_CodigoClie
                        If .EditText = Empty Then Exit Sub
                        If objBLBFunc.FcVerifItensRepetidos(grdClientes, Row, conCOL_SonProdAte_CodigoClie, .EditText) = False Then
                           MsgBox "Este Cliente ja foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, conCOL_SonProdAte_CodigoClie) = ""
                           .Cell(flexcpText, Row, conCOL_SonProdAte_DescClie) = ""
                           Cancel = True
                           Exit Sub
                        End If
                        If Len(Trim(PegaDescrCliente(CLng(.EditText)))) = 0 Then
                           MsgBox "Este cliente não existe !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, Col) = Empty
                           .Cell(flexcpText, Row, conCOL_SonProdAte_DescClie) = Empty
                           Cancel = True
                           Exit Sub
                        End If
                        .Cell(flexcpText, Row, conCOL_SonProdAte_CodigoClie) = .EditText
                        .Cell(flexcpText, Row, conCOL_SonProdAte_DescClie) = PegaDescrCliente(CLng(.EditText))
          End Select
     End With
End Sub

Private Sub grdCores_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     With grdCores
          Select Case Col
                 Case conCOL_SonCores_Qtde
                      If Len(Trim(.Cell(flexcpText, Row, conCOL_SonCores_Qtde))) > 0 Then .Cell(flexcpText, Row, conCOL_SonCores_Qtde) = Format(.Cell(flexcpText, Row, conCOL_SonCores_Qtde), "#,####0.0000")
          End Select
     End With
End Sub

Private Sub grdCores_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonCores_DescCores, _
         conCOL_SonCores_Ordem
         Cancel = True
    Case conCOL_SonCores_CodCor, _
         conCOL_SonCores_PesqCor, _
         conCOL_SonCores_UniMed, _
         conCOL_SonCores_Qtde
         If cTipOper = "C" Then Cancel = True
    Case Else
        grdCores.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub grdCores_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    If cTipOper = "C" Then Exit Sub
    
    If (grdCores.Rows - 1) = 0 Then Exit Sub
    
    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    Select Case Col
        Case conCOL_SonCores_PesqCor
    
            sSql = "Select " & vbCrLf
            sSql = sSql & "       * " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       SGI_FILIAL      = " & FILIAL & vbCrLf
            sSql = sSql & "   And SGI_PRODUTOTIPO = 0"
            
            arrTABELA(1) = sSql
            
            arrCAMPOS(1, 1) = "SGI_CODIGO"
            arrCAMPOS(1, 2) = "S"
            arrCAMPOS(1, 3) = "Código"
            arrCAMPOS(1, 4) = "1500"
            arrCAMPOS(1, 5) = "SGI_CODIGO"
            
            arrCAMPOS(2, 1) = "SGI_DESCRICAO"
            arrCAMPOS(2, 2) = "S"
            arrCAMPOS(2, 3) = "Nome"
            arrCAMPOS(2, 4) = "5000"
            arrCAMPOS(2, 5) = "SGI_DESCRICAO"
            
            varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Produtos")
            
            If Len(Trim(varRETORNO)) > 0 Then
               grdCores.Cell(flexcpText, Row, conCOL_SonCores_Codigo) = PegaIDProduto(varRETORNO)
               grdCores.Cell(flexcpText, Row, conCOL_SonCores_CodCor) = varRETORNO
               grdCores.Cell(flexcpText, Row, conCOL_SonCores_DescCores) = PegaDescrProduto(varRETORNO)
            End If
            
            If objBLBFunc.FcVerifItensRepetidos(grdCores, Row, conCOL_SonCores_CodCor, varRETORNO) = False Then
               MsgBox "Este Produto ja foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
               grdCores.Cell(flexcpText, Row, conCOL_SonCores_CodCor) = ""
               grdCores.Cell(flexcpText, Row, conCOL_SonCores_DescCores) = ""
               Exit Sub
            End If

    End Select

End Sub

Private Sub grdCores_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdCores
          Select Case Col
                    Case conCOL_SonCores_CodCor
                         ''KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
                         KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
                    Case conCOL_SonCores_Qtde
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 4, myvarAsCurrency)
          End Select
     End With
End Sub

Private Sub grdCores_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
     With grdCores
          Select Case Col
                 Case conCOL_SonCores_CodCor
                        If .EditText = Empty Then Exit Sub
                        If objBLBFunc.FcVerifItensRepetidos(grdCores, Row, conCOL_SonCores_CodCor, .EditText) = False Then
                           MsgBox "Esta Cor ja foi relacionada na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                           grdCores.Cell(flexcpText, Row, conCOL_SonCores_CodCor) = ""
                           grdCores.Cell(flexcpText, Row, conCOL_SonCores_DescCores) = ""
                           Cancel = True
                           Exit Sub
                        End If
                        If Len(Trim(PegaDescrProduto(.EditText))) = 0 Then
                           MsgBox "Este Produto não existe !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, Col) = Empty
                           .Cell(flexcpText, Row, conCOL_SonCores_DescCores) = Empty
                           Cancel = True
                           Exit Sub
                        End If
                        .Cell(flexcpText, Row, conCOL_SonCores_Codigo) = PegaIDProduto(.EditText)
                        .Cell(flexcpText, Row, conCOL_SonCores_DescCores) = PegaDescrProduto(.EditText)
          End Select
     End With
End Sub

Private Sub grdFamMaq_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonFamMaq_DescFam
         Cancel = True
    Case conCOL_SonFamMaq_CodFamMaq, _
         conCOL_SonFamMaq_PesqFam
         If cTipOper = "C" Then Cancel = True
    Case Else
        grdFamMaq.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub grdFamMaq_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    If cTipOper = "C" Then Exit Sub
    
    If (grdFamMaq.Rows - 1) = 0 Then Exit Sub
    
    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    Select Case Col
        Case conCOL_SonFamMaq_PesqFam
    
            sSql = "Select " & vbCrLf
            sSql = sSql & "       *" & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CADFAMMAQUINAS " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       SGI_FILIAL = " & FILIAL
            
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
            
            varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Familia de Máquinas")
                        
            If Len(Trim(varRETORNO)) > 0 Then
               With grdFamMaq
                    .Cell(flexcpText, Row, conCOL_SonFamMaq_CodFamMaq) = varRETORNO
                    .Cell(flexcpText, Row, conCOL_SonFamMaq_DescFam) = PegaDescrFamMaq(varRETORNO)
               End With
            End If
            
            If objBLBFunc.FcVerifItensRepetidos(grdFamMaq, Row, conCOL_SonFamMaq_CodFamMaq, varRETORNO) = False Then
               MsgBox "Esta Familia de máquina já foi relacionada na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
               grdFamMaq.Cell(flexcpText, Row, conCOL_SonFamMaq_CodFamMaq) = Empty
               grdFamMaq.Cell(flexcpText, Row, conCOL_SonFamMaq_DescFam) = Empty
               Exit Sub
            End If

    End Select

End Sub

Private Sub grdFamMaq_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdFamMaq
          Select Case Col
                    Case conCOL_SonFamMaq_CodFamMaq
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
          End Select
     End With
End Sub

Private Sub grdFamMaq_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
     With grdFamMaq
          Select Case Col
                 Case conCOL_SonFamMaq_CodFamMaq
                        If .EditText = Empty Then Exit Sub
                        If objBLBFunc.FcVerifItensRepetidos(grdFamMaq, Row, conCOL_SonCoefic_CodParam, .EditText) = False Then
                           MsgBox "Esta familia foi relacionada na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, conCOL_SonFamMaq_CodFamMaq) = Empty
                           .Cell(flexcpText, Row, conCOL_SonFamMaq_DescFam) = Empty
                           Cancel = True
                           Exit Sub
                        End If
                        If Len(Trim(PegaDescrFamMaq(.EditText))) = 0 Then
                           MsgBox "Esta familia não existe !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, Col) = Empty
                           .Cell(flexcpText, Row, conCOL_SonFamMaq_DescFam) = Empty
                           Cancel = True
                           Exit Sub
                        End If
                        .Cell(flexcpText, Row, conCOL_SonFamMaq_DescFam) = PegaDescrFamMaq(.EditText)
          End Select
     End With
End Sub



Private Sub grdPRODENTR_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Select Case Col
           Case conCOL_SonProdEntr_Qtde
                If Len(Trim(grdPRODENTR.Cell(flexcpText, Row, Col))) > 0 Then grdPRODENTR.Cell(flexcpText, Row, Col) = Format(grdPRODENTR.Cell(flexcpText, Row, Col), "#,####0.0000")
    End Select
    Exit Sub
End Sub

Private Sub grdPRODENTR_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonProdEntr_IDProduto, _
         conCOL_SonProdEntr_DescProd
         Cancel = True
    Case conCOL_SonProdEntr_CodProd, _
         conCOL_SonProdEntr_PesqProd, _
         conCOL_SonProdEntr_UniMed, _
         conCOL_SonProdEntr_Qtde
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
        Case conCOL_SonProdEntr_PesqProd
    
            sSql = "Select " & vbCrLf
            sSql = sSql & "       * " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       SGI_FILIAL      = " & FILIAL & vbCrLf
            sSql = sSql & "   And SGI_PRODUTOTIPO = 0"
            
            arrTABELA(1) = sSql
            
            arrCAMPOS(1, 1) = "SGI_CODIGO"
            arrCAMPOS(1, 2) = "S"
            arrCAMPOS(1, 3) = "Código"
            arrCAMPOS(1, 4) = "1500"
            arrCAMPOS(1, 5) = "SGI_CODIGO"
            
            arrCAMPOS(2, 1) = "SGI_DESCRICAO"
            arrCAMPOS(2, 2) = "S"
            arrCAMPOS(2, 3) = "Descrição"
            arrCAMPOS(2, 4) = "4000"
            arrCAMPOS(2, 5) = "SGI_DESCRICAO"
            
            arrCAMPOS(3, 1) = "SGI_COMPLEMENTO"
            arrCAMPOS(3, 2) = "S"
            arrCAMPOS(3, 3) = "Complemento"
            arrCAMPOS(3, 4) = "3000"
            arrCAMPOS(3, 5) = "SGI_COMPLEMENTO"
            
            varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Produtos")
            
            If Len(Trim(varRETORNO)) > 0 Then
                If objBLBFunc.FcVerifItensRepetidos(grdPRODENTR, Row, conCOL_SonProdEntr_CodProd, varRETORNO) = False Then
                   MsgBox "Este Produto ja foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                   Call LimpaGrdProdEntr(Row)
                   Exit Sub
                End If
                grdPRODENTR.Cell(flexcpText, Row, conCOL_SonProdEntr_CodProd) = varRETORNO
                grdPRODENTR.Cell(flexcpText, Row, conCOL_SonProdEntr_DescProd) = PegaDescrProduto(varRETORNO)
                grdPRODENTR.Cell(flexcpText, Row, conCOL_SonProdEntr_IDProduto) = PegaIDProduto(varRETORNO)
            End If
    End Select

End Sub

Private Sub grdPRODENTR_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdCores
          Select Case Col
                    Case conCOL_SonProdEntr_CodProd
                         KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
                    Case conCOL_SonProdEntr_Qtde
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 4, myvarAsDouble)
          End Select
     End With
End Sub

Private Sub grdPRODENTR_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
     With grdPRODENTR
          Select Case Col
                 Case conCOL_SonProdEntr_CodProd
                        If .EditText = Empty Then Exit Sub
                        If objBLBFunc.FcVerifItensRepetidos(grdCores, Row, conCOL_SonProdEntr_CodProd, .EditText) = False Then
                           MsgBox "Este Produto ja foi relacionada na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                           Cancel = True
                           Exit Sub
                        End If
                        If Len(Trim(PegaDescrProduto(.EditText))) = 0 Then
                            MsgBox "Este Produto não existe !!!", vbOKOnly + vbExclamation, "Aviso"
                            Call LimpaGrdProdEntr(Row)
                            Cancel = True
                            Exit Sub
                        End If
                        .Cell(flexcpText, Row, conCOL_SonProdEntr_IDProduto) = PegaIDProduto(.EditText)
                        .Cell(flexcpText, Row, conCOL_SonProdEntr_DescProd) = PegaDescrProduto(.EditText)
          End Select
     End With
End Sub

Private Sub grdSomFecha_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    With grdSomFecha(Index)
        Select Case Col
        Case conCOL_SomFecha_Valor
             If .Cell(flexcpText, Row, conCOL_SomFecha_Valor) <> Empty Then
                .Cell(flexcpText, Row, conCOL_SomFecha_Valor) = Format(.Cell(flexcpText, Row, conCOL_SomFecha_Valor), "#,##0.00")
             End If
        End Select
    End With
    Exit Sub
End Sub


Private Sub grdSomFecha_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdSomFecha(Index)
          Select Case Col
                    Case conCOL_SomFecha_Valor
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 2, myvarAsCurrency)
          End Select
     End With
End Sub

Private Sub grdUnidConv_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Select Case Col
    Case conCOL_SonUnid_CodUnid
         grdUnidConv.Col = Col + 4
         grdUnidConv.EditCell
    Case conCOL_SonUnid_PesqUnid
         grdUnidConv.Col = Col + 3
         grdUnidConv.EditCell
    Case conCOL_SonUnid_Fator
         If grdUnidConv.Cell(flexcpText, Row, conCOL_SonUnid_Fator) <> Empty Then
            grdUnidConv.Cell(flexcpText, Row, conCOL_SonUnid_Fator) = Format(grdUnidConv.Cell(flexcpText, Row, conCOL_SonUnid_Fator), "#,####0.0000")
         End If
    End Select
    Exit Sub
End Sub

Private Sub grdUnidConv_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonUnid_Desc_Unid, conCOL_SonUnid_Unidade
         Cancel = True
    Case conCOL_SonUnid_CodUnid, _
         conCOL_SonUnid_PesqUnid, _
         conCOL_SonUnid_Fator
         If cTipOper = "C" Then Cancel = True
    Case Else
        grdUnidConv.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub grdUnidConv_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    If cTipOper = "C" Then Exit Sub
    
    With grdUnidConv
         If (.Rows - 1) = 0 Then Exit Sub
        
         ReDim arrCAMPOS(1 To 4, 1 To 5) As String
         ReDim arrTABELA(1 To 1) As String
        
         Select Case Col
                Case conCOL_SonUnid_PesqUnid
                    
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
                        .Cell(flexcpText, Row, conCOL_SonUnid_CodUnid) = varRETORNO
                        .Cell(flexcpText, Row, conCOL_SonUnid_Unidade) = PegaUnidade(CLng(varRETORNO))
                        .Cell(flexcpText, Row, conCOL_SonUnid_Desc_Unid) = PegaDescrUnidade(CLng(varRETORNO))
                        If VerifItensRepetidosuUnidConv(Row, conCOL_SonUnid_CodUnid, Trim(varRETORNO)) = False Then
                           MsgBox "Está unidade já foi relacionada na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, conCOL_SonUnid_CodUnid) = Empty
                           .Cell(flexcpText, Row, conCOL_SonUnid_Unidade) = Empty
                           .Cell(flexcpText, Row, conCOL_SonUnid_Desc_Unid) = Empty
                           .Cell(flexcpText, Row, conCOL_SonUnid_Fator) = Empty
                           Exit Sub
                        End If
                     End If
                    
         End Select
    End With
                    

End Sub

Private Sub grdUnidConv_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdUnidConv
          Select Case Col
                    Case conCOL_SonUnid_CodUnid
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
                    Case conCOL_SonUnid_Fator
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 4, myvarAsDouble)
          End Select
     End With
End Sub

Private Sub grdUnidConv_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
     With grdUnidConv
          Select Case Col
                 Case conCOL_SonUnid_CodUnid
                        If .EditText = Empty Then Exit Sub
                        If VerifItensRepetidosuUnidConv(Row, conCOL_SonUnid_CodUnid, .EditText) = False Then
                           MsgBox "Está unidade ja foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, conCOL_SonUnid_CodUnid) = Empty
                           .Cell(flexcpText, Row, conCOL_SonUnid_Unidade) = Empty
                           .Cell(flexcpText, Row, conCOL_SonUnid_Desc_Unid) = Empty
                           .Cell(flexcpText, Row, conCOL_SonUnid_Fator) = Empty
                           Cancel = True
                           Exit Sub
                        End If
                        If Len(Trim(PegaDescrUnidade(CLng(.EditText)))) = 0 Then
                           MsgBox "Está unidade não existe !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, conCOL_SonUnid_CodUnid) = Empty
                           .Cell(flexcpText, Row, conCOL_SonUnid_Unidade) = Empty
                           .Cell(flexcpText, Row, conCOL_SonUnid_Desc_Unid) = Empty
                           .Cell(flexcpText, Row, conCOL_SonUnid_Fator) = Empty
                           Cancel = True
                           Exit Sub
                        End If
                        .Cell(flexcpText, Row, conCOL_SonUnid_Unidade) = PegaUnidade(CLng(.EditText))
                        .Cell(flexcpText, Row, conCOL_SonUnid_Desc_Unid) = PegaDescrUnidade(CLng(.EditText))
          End Select
     End With
End Sub

Private Sub grdVedanteCompound_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     With grdVedanteCompound
          Select Case Col
                 Case conCOL_SonVedante_Qtde
                      If Len(Trim(.Cell(flexcpText, Row, conCOL_SonVedante_Qtde))) > 0 Then .Cell(flexcpText, Row, conCOL_SonVedante_Qtde) = Format(.Cell(flexcpText, Row, conCOL_SonVedante_Qtde), "#,####0.0000")
          End Select
     End With
End Sub

Private Sub grdVedanteCompound_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonVedante_DescProd
         Cancel = True
    Case conCOL_SonVedante_CodProd, _
         conCOL_SonVedante_PesqProd, _
         conCOL_SonVedante_UniMed, _
         conCOL_SonVedante_Qtde
         If cTipOper = "C" Then Cancel = True
    Case Else
        grdVedanteCompound.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub grdVedanteCompound_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    If cTipOper = "C" Then Exit Sub
    
    If (grdVedanteCompound.Rows - 1) = 0 Then Exit Sub
    
    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    Select Case Col
        Case conCOL_SonVedante_PesqProd
    
            sSql = "Select " & vbCrLf
            sSql = sSql & "       * " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       SGI_FILIAL      = " & FILIAL & vbCrLf
            sSql = sSql & "   And SGI_PRODUTOTIPO = 0"
            
            arrTABELA(1) = sSql
            
            arrCAMPOS(1, 1) = "SGI_CODIGO"
            arrCAMPOS(1, 2) = "S"
            arrCAMPOS(1, 3) = "Código"
            arrCAMPOS(1, 4) = "1500"
            arrCAMPOS(1, 5) = "SGI_CODIGO"
            
            arrCAMPOS(2, 1) = "SGI_DESCRICAO"
            arrCAMPOS(2, 2) = "S"
            arrCAMPOS(2, 3) = "Nome"
            arrCAMPOS(2, 4) = "5000"
            arrCAMPOS(2, 5) = "SGI_DESCRICAO"
            
            varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Produtos")
            
            If Len(Trim(varRETORNO)) > 0 Then
               grdVedanteCompound.Cell(flexcpText, Row, conCOL_SonVedante_Codigo) = PegaIDProduto(varRETORNO)
               grdVedanteCompound.Cell(flexcpText, Row, conCOL_SonVedante_CodProd) = varRETORNO
               grdVedanteCompound.Cell(flexcpText, Row, conCOL_SonVedante_DescProd) = PegaDescrProduto(varRETORNO)
            End If
            
            If objBLBFunc.FcVerifItensRepetidos(grdVedanteCompound, Row, conCOL_SonVedante_CodProd, varRETORNO) = False Then
               MsgBox "Este Produto ja foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
               grdVedanteCompound.Cell(flexcpText, Row, conCOL_SonVedante_CodProd) = ""
               grdVedanteCompound.Cell(flexcpText, Row, conCOL_SonVedante_DescProd) = ""
               Exit Sub
            End If

    End Select

End Sub

Private Sub grdVedanteCompound_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdVedanteCompound
          Select Case Col
                    Case conCOL_SonVedante_CodProd
                         KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
                    Case conCOL_SonVedante_Qtde
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 4, myvarAsCurrency)
          End Select
     End With
End Sub

Private Sub grdVedanteCompound_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

     With grdVedanteCompound
          Select Case Col
                 Case conCOL_SonVedante_CodProd
                        If .EditText = Empty Then Exit Sub
                        If objBLBFunc.FcVerifItensRepetidos(grdVedanteCompound, Row, conCOL_SonVedante_CodProd, .EditText) = False Then
                           MsgBox "Este Produto ja foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, conCOL_SonVedante_CodProd) = ""
                           .Cell(flexcpText, Row, conCOL_SonVedante_DescProd) = ""
                           Cancel = True
                           Exit Sub
                        End If
                        If Len(Trim(PegaDescrProduto(.EditText))) = 0 Then
                           MsgBox "Este Produto não existe !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, Col) = Empty
                           .Cell(flexcpText, Row, conCOL_SonVedante_DescProd) = Empty
                           Cancel = True
                           Exit Sub
                        End If
                        .Cell(flexcpText, Row, conCOL_SonVedante_Codigo) = PegaIDProduto(.EditText)
                        .Cell(flexcpText, Row, conCOL_SonVedante_DescProd) = PegaDescrProduto(.EditText)
          End Select
     End With

End Sub

Private Sub grdVernizEsm_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     With grdVernizEsm
          Select Case Col
                 Case conCOL_SonVerniz_Qtde
                      If Len(Trim(.Cell(flexcpText, Row, conCOL_SonVerniz_Qtde))) > 0 Then .Cell(flexcpText, Row, conCOL_SonVerniz_Qtde) = Format(.Cell(flexcpText, Row, conCOL_SonVerniz_Qtde), "#,####0.0000")
          End Select
     End With
End Sub

Private Sub grdVernizEsm_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonVerniz_DescVerniz
         Cancel = True
    Case conCOL_SonVerniz_CodTipo, _
         conCOL_SonVerniz_PesqVerniz, _
         conCOL_SonVerniz_UnidMed, _
         conCOL_SonVerniz_Qtde
         If cTipOper = "C" Then
            Cancel = True
         Else
            
            If Len(Trim(txtEspecie.Text)) = 0 Then
               If Row = 1 Or Row = 2 Then
                  Cancel = True
                  Exit Sub
               End If
            End If
            
            sSql = ""
            
            sSql = "Select " & vbCrLf
            sSql = sSql & "       SGI_Vern01" & vbCrLf
            sSql = sSql & "      ,SGI_Vern02" & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CADESPPROD" & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
            sSql = sSql & "   And SGI_CODIGO = " & txtEspecie.Text
            
            BREC.Open sSql, adoBanco_Dados, adOpenDynamic
            If Row = 1 Then
               If BREC!SGI_Vern01 = 0 Then Cancel = True
               If BREC!SGI_Vern01 = 1 Then Cancel = False
            ElseIf Row = 2 Then Cancel = True
               If BREC!SGI_Vern02 = 0 Then Cancel = True
               If BREC!SGI_Vern02 = 1 Then Cancel = False
            End If
            BREC.Close
            
         End If
    Case Else
        grdVernizEsm.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub grdVernizEsm_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    If cTipOper = "C" Then Exit Sub
    
    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    Select Case Col
        Case conCOL_SonVerniz_PesqVerniz
    
            sSql = "Select " & vbCrLf
            sSql = sSql & "       * " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       SGI_FILIAL      = " & FILIAL & vbCrLf
            sSql = sSql & "   And SGI_PRODUTOTIPO = 0"
            
            arrTABELA(1) = sSql
            
            arrCAMPOS(1, 1) = "SGI_CODIGO"
            arrCAMPOS(1, 2) = "S"
            arrCAMPOS(1, 3) = "Código"
            arrCAMPOS(1, 4) = "1500"
            arrCAMPOS(1, 5) = "SGI_CODIGO"
            
            arrCAMPOS(2, 1) = "SGI_DESCRICAO"
            arrCAMPOS(2, 2) = "S"
            arrCAMPOS(2, 3) = "Nome"
            arrCAMPOS(2, 4) = "5000"
            arrCAMPOS(2, 5) = "SGI_DESCRICAO"
            
            varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Cadastro de Produtos")
            
            If Len(Trim(varRETORNO)) > 0 Then
               grdVernizEsm.Cell(flexcpText, Row, conCOL_SonVerniz_Codigo) = PegaIDProduto(varRETORNO)
               grdVernizEsm.Cell(flexcpText, Row, conCOL_SonVerniz_CodTipo) = varRETORNO
               grdVernizEsm.Cell(flexcpText, Row, conCOL_SonVerniz_DescVerniz) = PegaDescrProduto(varRETORNO)
            End If
            
            If objBLBFunc.FcVerifItensRepetidos(grdVernizEsm, Row, conCOL_SonVerniz_CodTipo, Trim(varRETORNO)) = False Then
               MsgBox "Este Produto ja foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
               grdVernizEsm.Cell(flexcpText, Row, conCOL_SonVerniz_Codigo) = Empty
               grdVernizEsm.Cell(flexcpText, Row, conCOL_SonVerniz_CodTipo) = Empty
               grdVernizEsm.Cell(flexcpText, Row, conCOL_SonVerniz_DescVerniz) = Empty
               Exit Sub
            End If

    End Select

End Sub

Private Sub grdVernizEsm_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdVernizEsm
          Select Case Col
                    Case conCOL_SonVerniz_CodTipo
                         ''KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
                         KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
                    Case conCOL_SonVerniz_Qtde
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 4, myvarAsDouble)
          End Select
     End With
End Sub

Private Sub grdVernizEsm_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
     With grdVernizEsm
          Select Case Col
                 Case conCOL_SonVerniz_CodTipo
                        If .EditText = Empty Then
                            .Cell(flexcpText, Row, conCOL_SonVerniz_Codigo) = Empty
                            .Cell(flexcpText, Row, conCOL_SonVerniz_DescVerniz) = Empty
                            Exit Sub
                        End If
                        If objBLBFunc.FcVerifItensRepetidos(grdVernizEsm, Row, conCOL_SonVerniz_CodTipo, Trim(.EditText)) = False Then
                           MsgBox "Este Produto ja foi relacionado na Grid. !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, conCOL_SonVerniz_Codigo) = Empty
                           .Cell(flexcpText, Row, conCOL_SonVerniz_CodTipo) = Empty
                           .Cell(flexcpText, Row, conCOL_SonVerniz_DescVerniz) = Empty
                           Cancel = True
                           Exit Sub
                        End If
                        If Len(Trim(PegaDescrProduto(.EditText))) = 0 Then
                           MsgBox "Este Produto não existe !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, conCOL_SonVerniz_Codigo) = Empty
                           .Cell(flexcpText, Row, Col) = Empty
                           .Cell(flexcpText, Row, conCOL_SonVerniz_DescVerniz) = Empty
                           Cancel = True
                           Exit Sub
                        End If
                        .Cell(flexcpText, Row, conCOL_SonVerniz_Codigo) = PegaIDProduto(.EditText)
                        .Cell(flexcpText, Row, conCOL_SonVerniz_DescVerniz) = PegaDescrProduto(.EditText)
          End Select
     End With
End Sub


Private Sub mskDtCadastro_GotFocus()
    objBLBFunc.SelecionaCampos mskDtCadastro.Name, frmCADPROD
End Sub


Private Sub optATUAUTOMSIMNAO_Click(Index As Integer)
    If Index = 0 Then txtPRECOPRODUTO.Enabled = False
    If Index = 1 Then txtPRECOPRODUTO.Enabled = True
End Sub

Private Sub optDimCortePAD_Click(Index As Integer)
    If Index = 1 Then
        Frame8.Enabled = False
        txtDESENV.Text = ""
        txtALTURA.Text = ""
    ElseIf Index = 0 Then
        Frame8.Enabled = True
    End If
End Sub

Private Sub optEstProduto_Click(Index As Integer)
    Call PintaLabel(IIf(optTipProd(0).Value = True, 0, 1), Index)
End Sub

Private Sub optQTDCORPPADRAOSN_Click(Index As Integer)
    If Index = 1 Then
        Frame13.Enabled = False
        txtQtdePorFolha.Text = ""
    ElseIf Index = 0 Then
        Frame13.Enabled = True
    End If
End Sub

Private Sub optTipProd_Click(Index As Integer)
    
    If cTipOper = "I" Then
        txtLinProd.Text = ""
        txtCodCliente.Text = ""
        txtCodRot.Text = ""
        txtDigVerif.Text = ""
        lblLinhProd.Caption = ""
        lblDesclie.Caption = ""
        txtCodigo.Text = ""
    
        If Index = 0 Then
            txtCodigo.Enabled = True
            Frame38(2).Enabled = False
            txtCodigo.SetFocus
        ElseIf Index = 1 Then
            txtCodigo.Enabled = False
            Frame38(2).Enabled = True
            If txtLinProd.Enabled = True And txtLinProd.Visible = True Then txtLinProd.SetFocus
        End If
    
    End If
    
    Call PintaLabel(Index, IIf(optEstProduto(0).Value = True, 0, 1))

End Sub



''Private Sub treListaMat_NodeClick(ByVal Node As MSComctlLib.Node)
''    lngINDLIST = Node.Index
''    If lngINDLIST = 1 Then
''        Call AbilDesCamposListMat(False)
''    ElseIf lngINDLIST > 1 Then
''        If cTipOper = "I" Or cTipOper = "A" Then
''            Call AbilDesCamposListMat(True)
''        Else
''            Call AbilDesCamposListMat(False)
''        End If
''        Call PegaDadosDoArray
''    End If
''End Sub

Private Sub txtALTURA_GotFocus()
    objBLBFunc.SelecionaCampos txtALTURA.Name, frmCADPROD
End Sub

Private Sub txtALTURA_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtALTURA.Text
End Sub

Private Sub txtALTURA_Validate(Cancel As Boolean)

    If Len(Trim(txtALTURA.Text)) = 0 Then Exit Sub
    
    txtALTURA.Text = Format(txtALTURA.Text, "#,##0.00")

End Sub

Private Sub txtArgolaEspess_GotFocus()
    objBLBFunc.SelecionaCampos txtArgolaEspess.Name, frmCADPROD
End Sub

Private Sub txtArgolaEspess_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtArgolaEspess.Text
End Sub

Private Sub txtArgolaRevest_GotFocus()
    objBLBFunc.SelecionaCampos txtArgolaRevest.Name, frmCADPROD
End Sub

Private Sub txtArgolaRevest_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtArgolaRevest.Text
End Sub

Private Sub txtArgolaRevest2_GotFocus()
    objBLBFunc.SelecionaCampos txtArgolaRevest2.Name, frmCADPROD
End Sub

Private Sub txtArgolaRevest2_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtArgolaRevest2.Text
End Sub



Private Sub txtBox_GotFocus()
    objBLBFunc.SelecionaCampos txtBox.Name, frmCADPROD
End Sub

Private Sub txtBox_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub


Private Sub txtCodCliente_GotFocus()
    objBLBFunc.SelecionaCampos txtCodCliente.Name, frmCADPROD
End Sub

Private Sub txtCodCliente_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCodCliente.Text
End Sub

Private Sub txtCodCliente_Validate(Cancel As Boolean)

    Dim i As Integer
    
    If Len(Trim(txtCodCliente.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtCodCliente.Text) = False Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODGRUP.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    lblDesclie.Caption = PegaDescClie(CLng(txtCodCliente.Text))
    If Len(Trim(lblDesclie.Caption)) = 0 Then
        MsgBox "Cliente Não Cadastrado !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCodCliente.Text = ""
        Cancel = True
        Exit Sub
    End If
    
    Call FormaCodProd
    
End Sub



Private Sub txtCODGRUP_GotFocus()
    objBLBFunc.SelecionaCampos txtCODGRUP.Name, frmCADPROD
End Sub

Private Sub txtCODGRUP_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODGRUP.Text
End Sub

Private Sub txtCODGRUP_Validate(Cancel As Boolean)

    If Len(Trim(txtCODGRUP.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtCODGRUP.Text) = False Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODGRUP.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADGRUPROD", txtCODGRUP.Text, lblDescFamProd, True)
    If Len(Trim(lblDescFamProd.Caption)) = 0 Then txtCODGRUP.Text = ""

    txtCODSUBGRUP.Text = ""
    lblDescSubFamProd.Caption = ""
    
    If optTipProd(0).Value = True Then txtCodigo.Text = Constroi_Codigo_Comprado(txtCODGRUP.Text, txtCODSUBGRUP.Text)

End Sub

Private Sub txtCodigo_GotFocus()
    objBLBFunc.SelecionaCampos txtCodigo.Name, frmCADPROD
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub



Private Sub txtCODPROD_GotFocus()
    objBLBFunc.SelecionaCampos txtCODPROD.Name, Me
End Sub

Private Sub txtCODPROD_Validate(Cancel As Boolean)

    Dim i As Integer
    
    If Len(Trim(txtCODPROD.Text)) = 0 Then Exit Sub
    
    lblDescProd.Caption = PegaDescrProduto(txtCODPROD.Text)
    If Len(Trim(lblDescProd.Caption)) = 0 Then
       MsgBox "Produto não cadastrada !!!", vbOKOnly + vbExclamation, "Aviso"
       txtCODPROD.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Call CriaArray
    ''Call PopDadosArray
    
End Sub

Private Sub txtCodProdFornec_GotFocus()
    objBLBFunc.SelecionaCampos txtCodProdFornec.Name, frmCADPROD
End Sub

Private Sub txtCodProdFornec_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtCodRot_GotFocus()
    objBLBFunc.SelecionaCampos txtCodRot.Name, frmCADPROD
End Sub

Private Sub txtCodRot_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCodRot.Text
End Sub


Private Sub txtCODSUBGRUP_GotFocus()
    objBLBFunc.SelecionaCampos txtCODSUBGRUP.Name, frmCADPROD
End Sub

Private Sub txtCODSUBGRUP_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODSUBGRUP.Text
End Sub

Private Sub txtCODSUBGRUP_Validate(Cancel As Boolean)

    If Len(Trim(txtCODGRUP.Text)) = 0 Then
        MsgBox "informe Primeiro a familia de Produto !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCODGRUP.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(txtCODSUBGRUP.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtCODSUBGRUP.Text) = False Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODSUBGRUP.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Call PegaDescTabelasSubFamProd(txtCODGRUP.Text, txtCODSUBGRUP.Text, lblDescSubFamProd, True)
    If Len(Trim(lblDescSubFamProd.Caption)) = 0 Then txtCODSUBGRUP.Text = ""
    
    txtEspecie.Text = ""
    lblDescEspecieProd.Caption = ""
   
    If optTipProd(0).Value = True Then txtCodigo.Text = Constroi_Codigo_Comprado(txtCODGRUP.Text, txtCODSUBGRUP.Text)
    
End Sub



Private Sub txtColEsp_GotFocus()
    objBLBFunc.SelecionaCampos txtColEsp.Name, frmCADPROD
End Sub

Private Sub txtColEsp_Validate(Cancel As Boolean)

   Dim i         As Integer
   Dim strIDProd As String

   If Len(Trim(txtColEsp.Text)) = 0 Then Exit Sub
   
   strIDProd = PegaIDProduto(Trim(txtColEsp.Text))
   
   If Len(Trim(strIDProd)) = 0 Then
      MsgBox "Este Produto Não existe !!!", vbOKOnly + vbExclamation, "Aviso"
      Cancel = True
      Exit Sub
   End If
   
   label11(16).Caption = PegaDescrProduto(Trim(txtColEsp.Text))
   objCADPRODUTO.ColEsp = IIf(IsNumeric(strIDProd) = True, CLng(strIDProd), 0)

End Sub

Private Sub txtComplemento_GotFocus()
    objBLBFunc.SelecionaCampos txtComplemento.Name, frmCADPROD
End Sub

Private Sub txtComplemento_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtCorpoEspess_GotFocus()
    objBLBFunc.SelecionaCampos txtCorpoEspess.Name, frmCADPROD
End Sub

Private Sub txtCorpoEspess_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCorpoEspess.Text
End Sub

Private Sub txtCorpoRevest_GotFocus()
    objBLBFunc.SelecionaCampos txtCorpoRevest.Name, frmCADPROD
End Sub

Private Sub txtCorpoRevest_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCorpoRevest.Text
End Sub

Private Sub txtCorpoRevest2_GotFocus()
    objBLBFunc.SelecionaCampos txtCorpoRevest2.Name, frmCADPROD
End Sub
Private Sub txtCorpoRevest2_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCorpoRevest2.Text
End Sub

Private Sub txtCUBAGEN_GotFocus()
    objBLBFunc.SelecionaCampos txtCUBAGEN.Name, frmCADPROD
End Sub

Private Sub txtCUBAGEN_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCUBAGEN.Text
End Sub

Private Sub txtCUBAGEN_Validate(Cancel As Boolean)

    If Len(Trim(txtCUBAGEN.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtCUBAGEN.Text) = False Then
       MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
       txtCUBAGEN.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtCUBAGEN.Text = Format(txtCUBAGEN.Text, "#,###0.000")

End Sub

Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, frmCADPROD
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub


Private Sub txtDESENV_GotFocus()
    objBLBFunc.SelecionaCampos txtDESENV.Name, frmCADPROD
End Sub

Private Sub txtDESENV_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtDESENV.Text
End Sub

Private Sub txtDESENV_Validate(Cancel As Boolean)

    If Len(Trim(txtDESENV.Text)) = 0 Then Exit Sub
    
    txtDESENV.Text = Format(txtDESENV.Text, "#,##0.00")
    
End Sub

Private Sub txtDigVerif_GotFocus()
    objBLBFunc.SelecionaCampos txtDigVerif.Name, frmCADPROD
End Sub

Private Sub txtDigVerif_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtDigVerif.Text
End Sub

Private Sub txtDigVerif_Validate(Cancel As Boolean)
    
    Dim i As Integer
    
    If Len(Trim(txtDigVerif.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtDigVerif.Text) = False Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtDigVerif.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Call FormaCodProd

End Sub

Private Sub txtDistancia_GotFocus()
    objBLBFunc.SelecionaCampos txtDistancia.Name, frmCADPROD
End Sub

Private Sub txtDistancia_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtDistancia.Text
End Sub

Private Sub txtEspecie_GotFocus()
    objBLBFunc.SelecionaCampos txtEspecie.Name, frmCADPROD
End Sub

Private Sub txtEspecie_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtEspecie.Text
End Sub


Private Sub txtEspecie_Validate(Cancel As Boolean)

    If Len(Trim(txtCODSUBGRUP.Text)) = 0 Then
        MsgBox "Primeiro informe a Sub-Familia de Produtos !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCODSUBGRUP.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(txtEspecie.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtEspecie.Text) = False Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtEspecie.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Call PegaDescTabelasEspProd(txtCODSUBGRUP.Text, txtEspecie.Text, lblDescEspecieProd, True)
    If Len(Trim(lblDescEspecieProd.Caption)) = 0 Then
       txtEspecie.Text = ""
       Exit Sub
    End If
    Call ConfCombosVerniz


End Sub

Private Sub txtEstMinimo_GotFocus()
    objBLBFunc.SelecionaCampos txtEstMinimo.Name, frmCADPROD
End Sub

Private Sub txtEstMinimo_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtEstMinimo.Text
End Sub

Private Sub txtEstMinimo_Validate(Cancel As Boolean)

    If Len(Trim(txtEstMinimo.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtEstMinimo.Text) = False Then
       MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
       txtEstMinimo.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtEstMinimo.Text = Format(txtEstMinimo.Text, "#0")

End Sub

Private Sub txtFundoEspess_GotFocus()
    objBLBFunc.SelecionaCampos txtFundoEspess.Name, frmCADPROD
End Sub

Private Sub txtFundoEspess_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtFundoEspess.Text
End Sub

Private Sub txtFundoRevest_GotFocus()
    objBLBFunc.SelecionaCampos txtFundoRevest.Name, frmCADPROD
End Sub

Private Sub txtFundoRevest_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtFundoRevest.Text
End Sub

Private Sub txtFundoRevest2_GotFocus()
    objBLBFunc.SelecionaCampos txtFundoRevest2.Name, frmCADPROD
End Sub

Private Sub txtFundoRevest2_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtFundoRevest2.Text
End Sub

Private Sub txtGRAMATURAM2_GotFocus()
    objBLBFunc.SelecionaCampos txtGRAMATURAM2.Name, frmCADPROD
End Sub

Private Sub txtGRAMATURAM2_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtGRAMATURAM2.Text
End Sub

Private Sub txtGRAMATURAM2_Validate(Cancel As Boolean)

    If Len(Trim(txtGRAMATURAM2.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtGRAMATURAM2.Text) = False Then
       MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
       txtGRAMATURAM2.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtGRAMATURAM2.Text = Format(txtGRAMATURAM2.Text, "#,###0.000")

End Sub

Private Sub txtIPI_GotFocus()
    objBLBFunc.SelecionaCampos txtIPI.Name, Me
End Sub

Private Sub txtIPI_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtIPI.Text
End Sub

Private Sub txtLARGURA_GotFocus()
    objBLBFunc.SelecionaCampos txtLARGURA.Name, frmCADPROD
End Sub

Private Sub txtLARGURA_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtLARGURA.Text
End Sub

Private Sub txtLARGURA_Validate(Cancel As Boolean)

    If Len(Trim(txtLARGURA.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtLARGURA.Text) = False Then
       MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
       txtLARGURA.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtLARGURA.Text = Format(txtLARGURA.Text, "#,###0.000")

End Sub

Private Sub txtLinProd_GotFocus()
    objBLBFunc.SelecionaCampos txtLinProd.Name, frmCADPROD
End Sub

Private Sub txtLinProd_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtLinProd.Text
End Sub

Private Sub txtLinProd_Validate(Cancel As Boolean)

    Dim i As Integer
    
    If Len(Trim(txtLinProd.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtLinProd.Text) = False Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtLinProd.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    lblLinhProd.Caption = PegaDescLinProd(CLng(txtLinProd.Text))
    If Len(Trim(lblLinhProd.Caption)) = 0 Then
       MsgBox "Linha de Produto não cadastrada !!!", vbOKOnly + vbExclamation, "Aviso"
       txtLinProd.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Call FormaCodProd
    Call ConfGridFecham
    Call ConfComboFech
    
End Sub

Private Sub txtMETROS_GotFocus()
    objBLBFunc.SelecionaCampos txtMETROS.Name, frmCADPROD
End Sub

Private Sub txtMETROS_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtMETROS.Text
End Sub

Private Sub txtMETROS_Validate(Cancel As Boolean)

    If Len(Trim(txtMETROS.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtMETROS.Text) = False Then
       MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
       txtMETROS.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtMETROS.Text = Format(txtMETROS.Text, "#,###0.000")

End Sub

Private Sub txtPeso_GotFocus()
    objBLBFunc.SelecionaCampos txtPeso.Name, frmCADPROD
End Sub

Private Sub txtPeso_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtSaldo.Text
End Sub

Private Sub txtPeso_Validate(Cancel As Boolean)

    If Len(Trim(txtPeso.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtPeso.Text) = False Then
       MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
       txtPeso.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtPeso.Text = Format(txtPeso.Text, "#,####0.0000")

End Sub

Private Sub txtPOCACRES_GotFocus()
    objBLBFunc.SelecionaCampos txtPOCACRES.Name, frmCADPROD
End Sub

Private Sub txtPOCACRES_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtPOCACRES.Text
End Sub

Private Sub txtPOCACRES_Validate(Cancel As Boolean)

    txtPRCFINAL.Text = ""
    
    If Len(Trim(txtPOCACRES.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtPOCACRES.Text) = False Then
       MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
       txtPOCACRES.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtPOCACRES.Text = Format(txtPOCACRES.Text, "#,##0.00")
    
End Sub

Private Sub txtPrateleira_GotFocus()
    objBLBFunc.SelecionaCampos txtPrateleira.Name, frmCADPROD
End Sub

Private Sub txtPrateleira_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtPRECOPRODUTO_GotFocus()
    objBLBFunc.SelecionaCampos txtPRECOPRODUTO.Name, frmCADPROD
End Sub

Private Sub txtPRECOPRODUTO_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtPRECOPRODUTO.Text
End Sub

Private Sub txtPRECOPRODUTO_Validate(Cancel As Boolean)

    txtPRCFINAL.Text = ""
    txtPOCACRES.Text = ""
    
    If Len(Trim(txtPRECOPRODUTO.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtPRECOPRODUTO.Text) = False Then
       MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
       txtPRECOPRODUTO.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtPRECOPRODUTO.Text = Format(txtPRECOPRODUTO.Text, "#,##0.00")
    
End Sub

Private Sub txtProcedencia_GotFocus()
    objBLBFunc.SelecionaCampos txtProcedencia.Name, frmCADPROD
End Sub

Private Sub txtProcedencia_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtQtdeIns_GotFocus()
    objBLBFunc.SelecionaCampos txtQtdeIns.Name, Me
End Sub

Private Sub txtQtdeIns_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtQtdeIns.Text
End Sub

Private Sub txtQtdeIns_Validate(Cancel As Boolean)
    
    If Len(Trim(txtQtdeIns.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtQtdeIns.Text) = False Then
       MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
       txtSaldo.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtQtdeIns.Text = Format(txtQtdeIns.Text, "#,####0.0000")
    If lngINDLIST > 1 Then
        arrPROVARV(lngINDLIST).strQTDCONS = txtQtdeIns.Text
    End If

    ''Call PopDadosArray
    
End Sub

Private Sub txtQtdePassada_GotFocus()
    objBLBFunc.SelecionaCampos txtQtdePassada.Name, frmCADPROD
End Sub

Private Sub txtQtdePassada_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtQtdePassada.Text
End Sub

Private Sub txtQtdePorFolha_GotFocus()
    objBLBFunc.SelecionaCampos txtQtdePorFolha.Name, frmCADPROD
End Sub

Private Sub txtQtdePorFolha_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtQtdePorFolha.Text
End Sub

Private Sub txtRua_GotFocus()
    objBLBFunc.SelecionaCampos txtRua.Name, frmCADPROD
End Sub

Private Sub txtRua_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtSaldo_GotFocus()
    objBLBFunc.SelecionaCampos txtSaldo.Name, frmCADPROD
End Sub
Private Sub txtSaldo_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtSaldo.Text
End Sub

Private Sub txtSaldo_Validate(Cancel As Boolean)
    
    If Len(Trim(txtSaldo.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtSaldo.Text) = False Then
       MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
       txtSaldo.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtSaldo.Text = Format(txtSaldo.Text, "#0")
    
End Sub


Private Sub txtTampaEspess_GotFocus()
    objBLBFunc.SelecionaCampos txtTampaEspess.Name, frmCADPROD
End Sub

Private Sub txtTampaEspess_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtTampaEspess.Text
End Sub

Private Sub txtTampaRevest_GotFocus()
    objBLBFunc.SelecionaCampos txtTampaRevest.Name, frmCADPROD
End Sub

Private Sub txtTampaRevest_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtTampaRevest.Text
End Sub

Private Sub txtTampaRevest2_GotFocus()
    objBLBFunc.SelecionaCampos txtTampaRevest2.Name, frmCADPROD
End Sub

Private Sub txtTampaRevest2_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtTampaRevest2.Text
End Sub



Private Sub txtTipo_GotFocus()
    objBLBFunc.SelecionaCampos txtTipo.Name, frmCADPROD
End Sub

Private Sub txtTipo_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtTipo.Text
End Sub

Private Sub txtTipo_Validate(Cancel As Boolean)

    If Len(Trim(txtTipo.Text)) = 0 Then Exit Sub
    
    If IsNumeric(txtTipo.Text) = False Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtTipo.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADTIPPROD", txtTipo.Text, lblDescTipoProd, True)
    If Len(Trim(lblDescTipoProd.Caption)) = 0 Then txtTipo.Text = ""

End Sub

Private Function ValidaCampos() As Boolean
   
   ValidaCampos = False
   
   Dim intPADRAO    As Integer
   Dim i            As Integer
      
   
   
   If optTipProd(1).Value = True Then
        If cboFechSoldaAgraf.ListIndex = -1 Then
           MsgBox "ATENÇÂO - Favor Informar o Fechamento de Solda Agrafado !!!", vbOKOnly + vbExclamation, "Aviso"
           Exit Function
        End If
        If cboFechTampaFuro.ListIndex = -1 Then
           MsgBox "ATENÇÃO - Favor informar a Tampa/Furo !!!", vbOKOnly + vbExclamation, "Aviso"
           Exit Function
        End If
   End If
   
   '' ==============================================
   '' Caso se For Produzido
   If optTipProd(1).Value = True And _
      optEstProduto(0).Value = True Then
        If Len(Trim(txtLinProd.Text)) = 0 Then
           MsgBox "A Linha do Produto deve ser Informado !!!", vbOKOnly + vbCritical, "Aviso"
           stProd.Tab = 0
           txtLinProd.SetFocus
           Exit Function
        End If
        If Len(Trim(txtCodCliente.Text)) = 0 Then
           MsgBox "O Código do Cliente deve ser Informado !!!", vbOKOnly + vbCritical, "Aviso"
           stProd.Tab = 0
           txtCodCliente.SetFocus
           Exit Function
        End If
        If Len(Trim(txtDigVerif.Text)) = 0 Then
           MsgBox "O digito verificador deve ser Informado !!!", vbOKOnly + vbCritical, "Aviso"
           stProd.Tab = 0
           txtDigVerif.SetFocus
           Exit Function
        End If
        If optDimCortePAD(0).Value = True Then
            If Len(Trim(txtDESENV.Text)) = 0 Then
                MsgBox "ATENÇÂO - Campo Desenvolvimento deve ser informado !!!", vbOKOnly + vbExclamation, "Aviso"
                txtDESENV.SetFocus
                Exit Function
            End If
            If Len(Trim(txtALTURA.Text)) = 0 Then
                MsgBox "ATENÇÂO - Campo Altura deve ser informado !!!", vbOKOnly + vbExclamation, "Aviso"
                txtALTURA.SetFocus
                Exit Function
            End If
        End If
   End If
   '' ==============================================
   
   '' ==============================================
   '' Caso se For Comprado
   If optTipProd(0).Value = True And _
      optEstProduto(0).Value = True Then
        If Len(Trim(txtCodigo.Text)) = 0 Then
           MsgBox "Campo código inválido !!!", vbOKOnly + vbCritical, "Aviso"
           stProd.Tab = 0
           txtCodigo.SetFocus
           Exit Function
        End If
   End If
   '' ==============================================
   
   If Len(Trim(txtDescricao.Text)) = 0 Then
      MsgBox "Campo descrição inválido !!!", vbOKOnly + vbCritical, "Aviso"
      stProd.Tab = 0
      txtDescricao.SetFocus
      Exit Function
   End If
   
   If Len(Trim(txtCODGRUP.Text)) = 0 Then
      MsgBox "Campo Grupo inválido !!!", vbOKOnly + vbCritical, "Aviso"
      stProd.Tab = 0
      txtCODGRUP.SetFocus
      Exit Function
   End If
   If Len(Trim(txtCODSUBGRUP.Text)) = 0 Then
      MsgBox "Campo Sub Grupo inválido !!!", vbOKOnly + vbCritical, "Aviso"
      stProd.Tab = 0
      txtCODSUBGRUP.SetFocus
      Exit Function
   End If
   If Len(Trim(txtEspecie.Text)) = 0 Then
      MsgBox "Campo espécie inválido !!!", vbOKOnly + vbCritical, "Aviso"
      stProd.Tab = 0
      txtEspecie.SetFocus
      Exit Function
   End If
   
   '' Depois Voltar
   ''If cboCODFAMMAQ.ListIndex = -1 Then
   ''   MsgBox "Campo familia de máquina deve ser preenchido !!!", vbOKOnly + vbCritical, "Aviso"
   ''   stProd.Tab = 9
   ''   tabProPadrao.Tab = 0
   ''   cboCODFAMMAQ.SetFocus
   ''   Exit Function
   ''End If
   
   If Len(Trim(txtTipo.Text)) = 0 Then
      MsgBox "Campo Tipo inválido !!!", vbOKOnly + vbCritical, "Aviso"
      stProd.Tab = 0
      txtTipo.SetFocus
      Exit Function
   End If
   
   If cboUnidade.ListIndex = -1 Then
      MsgBox "Campo unidade inválida !!!", vbOKOnly + vbCritical, "Aviso"
      stProd.Tab = 0
      cboUnidade.SetFocus
      Exit Function
   End If
   
   If optATUAUTOMSIMNAO(0).Value = False And optATUAUTOMSIMNAO(1).Value = False Then
      MsgBox "Informe se os preços vai ser atualizado automático (S/N) !!!", vbOKOnly + vbExclamation, "Aviso"
      stProd.Tab = 0
      optATUAUTOMSIMNAO(0).SetFocus
      Exit Function
   End If
   
   If Len(Trim(txtSaldo.Text)) > 0 Then
      ''If Val(txtSaldo.Text) < 0 Then
         ''MsgBox "Não é permitido valores negativos !!!", vbOKOnly + vbCritical, "Aviso"
         ''stProd.Tab = 0
         ''txtSaldo.Text = ""
         ''txtSaldo.SetFocus
         ''Exit Function
      ''End If
   End If
   
   If Len(Trim(txtEstMinimo.Text)) > 0 Then
      If Val(txtEstMinimo.Text) < 0 Then
         MsgBox "Não é permitido valores negativos !!!", vbOKOnly + vbCritical, "Aviso"
         stProd.Tab = 0
         txtEstMinimo.Text = ""
         txtEstMinimo.SetFocus
         Exit Function
      End If
   End If
   
   If IsDate(mskDtCadastro.Text) = False Then
      MsgBox "Data do Cadastro inválido !!!", vbOKOnly + vbCritical, "Aviso"
      stProd.Tab = 0
      mskDtCadastro.SetFocus
      Exit Function
   End If
   
   If (optTipProd(0).Value = False) And (optTipProd(1).Value = False) Then
      MsgBox "Informe o tipo do Produto !!!", vbOKOnly + vbExclamation, "Aviso"
      stProd.Tab = 0
      optTipProd(0).SetFocus
      Exit Function
   End If
   
   If Len(Trim(strCAMINHO)) > 0 Then
        If Trim(UCase(strCAMINHO)) <> Trim(UCase(strCamImgRotulos)) Then
           MsgBox "ATENÇÃO - O local de Abertura do arquivo de imagen é inválido !!!", vbOKOnly + vbExclamation, "Aviso"
           Exit Function
        End If
   End If
   
   ValidaCampos = True

End Function
   
   
Private Sub Altera()

    Dim i   As Integer
    Dim j   As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    
    Frame3.Enabled = True
    Frame38(3).Enabled = True
    Frame38(4).Enabled = True
    stProd.Enabled = True
    Frame14.Enabled = False
    Frame38(0).Enabled = True
    Frame38(1).Enabled = True
    Frame38(2).Enabled = False
    Frame6.Enabled = True
    Frame9.Enabled = True
   
    Frame41(1).Enabled = True
    Frame38(5).Enabled = True
    Frame43(0).Enabled = True
    Frame43(1).Enabled = True
    Frame43(3).Enabled = True
    Frame8.Enabled = False
    
    Call Limpa_Lista_Lata
    
    optFilialPed(0).Value = True
    optProdNovoSN(1).Value = True
    
    txtCodRot.Enabled = False
   
    label11(15).Caption = ""
    label11(16).Caption = ""
    
    
    txtSaldo.Enabled = False
    mskDtCadastro.Enabled = False
    
    stProd.Tab = 0
    
    txtCodigo.Enabled = True
    txtPRECOPRODUTO.Enabled = False
   
    Me.Caption = "Cadastro de produtos - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADPROD
    LimpaCaptions
    txtDescEquip.Locked = False
    
    objCADPRODUTO.PreenchComboUnidade cboUnidade
    objCADPRODUTO.PreenchComboClaFis cboClass
    objCADPRODUTO.PreenchComboUnidade cboUniIns
    cboUniIns.ListIndex = -1
    
    optATUAUTOMSIMNAO(0).Value = False
    optATUAUTOMSIMNAO(1).Value = False
    optAlcGalSN(0).Value = True
    optNeckInSN(0).Value = True
    optDimCortePAD(1).Value = True
    optQTDCORPPADRAOSN(1).Value = True
    
    
    lblPRECOMEDIO.Caption = ""
    
    InitGridUnidConv
    InitGridVernizEsm
    InitGridCores
    InitGrdVedanteCompound
    InitgrdSomFecha
    InitGrdProdAtend
    InitGrdFamMaq
    
    Call ConfCombosVerniz
    Call ConfComboFech
    
    Call PreenchComboAlca
    Call ConfGridFecham
    Call ConfEstoque
    Call ConfProdEntr
    Call ConfEstoqueLit
    
    objCADPRODUTO.IDProduto = iCodigo
    cboFechSoldaAgraf.ListIndex = -1
    
    objCADPRODUTO.TampaPressao = conCOL_SomFecha_TampaPressao
    objCADPRODUTO.BatoqueRetratil = conCOL_SomFecha_BatoqueRetra
    objCADPRODUTO.BatoquePlastico = conCOL_SomFecha_BatoquePlast
    objCADPRODUTO.TAMPAVIS = conCOL_SomFecha_TampaVisor
    
    txtCodigo.Enabled = False
    optUsarPadrPalSN(0).Value = True
    
    Call AbilDesDadosListMat
    Call LimpaCamposLabel
    
    If Len(Trim(objCADPRODUTO.COMPLEMENTO)) > 0 Then txtComplemento.Text = Trim(objCADPRODUTO.COMPLEMENTO)
    
    
    If objCADPRODUTO.Carrega_campos = True Then
    
        txtLinProd.Text = objCADPRODUTO.CodLinProd
        Call ConfGridFecham
        Call ConfComboFech
       
        txtIPI.Text = objCADPRODUTO.IPI
       
        txtCodCliente.Text = objCADPRODUTO.CodClie
        txtCodRot.Text = objCADPRODUTO.CodRotulo
        txtDigVerif.Text = objCADPRODUTO.DigVerif
        optNatSimNao(objCADPRODUTO.NATURALSIMNAO).Value = True
       
       Label1(23).Caption = ""
       If objCADPRODUTO.STATUS = 0 Then Label1(23).Caption = "DESATIVADO"
       If objCADPRODUTO.STATUS = 1 Then Label1(23).Caption = "ATIVO"
       If objCADPRODUTO.STATUS = 2 Then Label1(23).Caption = "AG.LIBERAÇÃO"
       
       optUsarPadrPalSN(objCADPRODUTO.PALLHETPADRAO).Value = True
       
       Call FormaCodProd
       
       lblLinhProd.Caption = PegaDescLinProd(CLng(txtLinProd.Text))
       lblDesclie.Caption = PegaDescClie(CLng(txtCodCliente.Text))
       
       txtCodigo.Text = objCADPRODUTO.CodigoProd
       txtCodProdFornec.Text = objCADPRODUTO.CODPROFORNEC
       txtDescricao.Text = objCADPRODUTO.DescriProd
       txtEspecie.Text = Str(objCADPRODUTO.EspProduto)
       
       If Len(Trim(objCADPRODUTO.COMPLEMENTO)) > 0 Then txtComplemento.Text = Trim(objCADPRODUTO.COMPLEMENTO)
       
       '' Aba Produção
       optLaudoSN(objCADPRODUTO.EMITLAUDO).Value = True
       
       txtPOCACRES.Text = Format(objCADPRODUTO.PORCACRES, "#,##0.00")
       If objCADPRODUTO.PORCACRES > 0 Then txtPOCACRES.Text = Format(objCADPRODUTO.PORCACRES, "#,##0.00")
       If objCADPRODUTO.PRCCUSTO > 0 Then txtPRCFINAL.Text = Format(objCADPRODUTO.PRCCUSTO, "#,##0.00")
       If objCADPRODUTO.PESOUNIT > 0 Then txtPeso.Text = Format(objCADPRODUTO.PESOUNIT, "#,###0.000")
       
       If objCADPRODUTO.CUBAGEN > 0 Then txtCUBAGEN.Text = Format(objCADPRODUTO.CUBAGEN, "#,###0.000")
       If objCADPRODUTO.QTDEPORFOLHA > 0 Then txtQtdePorFolha.Text = objCADPRODUTO.QTDEPORFOLHA
       If objCADPRODUTO.QTDPASSADAS > 0 Then txtQtdePassada.Text = objCADPRODUTO.QTDPASSADAS
       
       If objCADPRODUTO.CODGRUPPROD > 0 Then txtCODGRUP.Text = objCADPRODUTO.CODGRUPPROD
       If objCADPRODUTO.CODSUBGPROD > 0 Then txtCODSUBGRUP.Text = objCADPRODUTO.CODSUBGPROD
       
       txtTipo.Text = Str(objCADPRODUTO.TIPPRODUTO)
       txtSaldo.Text = Format(objCADPRODUTO.Saldo, "#0")
       txtEstMinimo.Text = Format(objCADPRODUTO.EstMin, "#0")
       mskDtCadastro.Text = objCADPRODUTO.DataCadast
       
       txtDescEquip.Text = objCADPRODUTO.DESCEQUIP
       txtProcedencia.Text = objCADPRODUTO.PROCEDEN
       
       txtRua.Text = objCADPRODUTO.RUA
       txtBox.Text = objCADPRODUTO.box
       txtPrateleira.Text = objCADPRODUTO.PRATELEIRA
       txtCodigoEAN.Text = objCADPRODUTO.CODEAN
       If objCADPRODUTO.DISTANCIA > 0 Then txtDistancia.Text = objCADPRODUTO.DISTANCIA
       
       
       If objCADPRODUTO.PRODUTOTIPO = 0 Then
          optTipProd(0).Value = True
          txtCodigo.Enabled = True
          txtDigVerif.Enabled = False
       End If
       If objCADPRODUTO.PRODUTOTIPO = 1 Then optTipProd(1).Value = True
       
       If objCADPRODUTO.ESTILOPROD = 0 Then optEstProduto(0).Value = True
       If objCADPRODUTO.ESTILOPROD = 1 Then optEstProduto(1).Value = True
       If objCADPRODUTO.ESTILOPROD = 2 Then optEstProduto(2).Value = True
       If objCADPRODUTO.ESTILOPROD = 3 Then optEstProduto(3).Value = True
       If objCADPRODUTO.ESTILOPROD = 4 Then optEstProduto(4).Value = True
       If objCADPRODUTO.ESTILOPROD = 5 Then optEstProduto(5).Value = True
              
       
       If objCADPRODUTO.DataUltMov > 0 Then mskDtUltMov.Text = Format(objCADPRODUTO.DataUltMov, "DD/MM/YYYY")
       arrProdLST = objCADPRODUTO.ProdLST
       arrPROCESSO = objCADPRODUTO.Processos
       arrFERRAMENTA = objCADPRODUTO.FERRAMENTA
       
       If objCADPRODUTO.PRCPROD > 0 Then txtPRECOPRODUTO.Text = Format(objCADPRODUTO.PRCPROD, "#,###0.000")
       If objCADPRODUTO.PRCMED > 0 Then lblPRECOMEDIO.Caption = Format(objCADPRODUTO.PRCMED, "#,###0.000")
       
       If objCADPRODUTO.ATAUTOMSN = 0 Then optATUAUTOMSIMNAO(0).Value = True
       If objCADPRODUTO.ATAUTOMSN = 1 Then optATUAUTOMSIMNAO(1).Value = True
       
       For i = 0 To (cboUnidade.ListCount - 1)
           If cboUnidade.ItemData(i) = objCADPRODUTO.Unidade Then cboUnidade.ListIndex = i
       Next i
       
       For i = 0 To (cboClass.ListCount - 1)
           If cboClass.ItemData(i) = objCADPRODUTO.CLASFISC Then cboClass.ListIndex = i
       Next i
       
       
       '' =============================================
       Call ConfCombosVerniz
       
       For i = 0 To (cboCorpoVerniz.ListCount - 1)
           If cboCorpoVerniz.ItemData(i) = objCADPRODUTO.VernCorpo Then cboCorpoVerniz.ListIndex = i
       Next i
       For i = 0 To (cboTampaVerniz.ListCount - 1)
           If cboTampaVerniz.ItemData(i) = objCADPRODUTO.VernTampa Then cboTampaVerniz.ListIndex = i
       Next i
       For i = 0 To (cboFundoVerniz.ListCount - 1)
           If cboFundoVerniz.ItemData(i) = objCADPRODUTO.VernFundo Then cboFundoVerniz.ListIndex = i
       Next i
       For i = 0 To (cboArgolaVerniz.ListCount - 1)
           If cboArgolaVerniz.ItemData(i) = objCADPRODUTO.VernArgola Then cboArgolaVerniz.ListIndex = i
       Next i

       For i = 0 To (cboFechTampaFuro.ListCount - 1)
           If cboFechTampaFuro.ItemData(i) = objCADPRODUTO.FechTampaFuro Then cboFechTampaFuro.ListIndex = i
       Next i
       '' =============================================
       
       If objCADPRODUTO.PRCPROD > 0 Then txtPRECOPRODUTO.Text = Format(objCADPRODUTO.PRCPROD, "#,##0.00")
       
       arrPRODTIPORCA = objCADPRODUTO.PRODTIPORCA
       
       arrPRODVAPROC = objCADPRODUTO.PRODVAPROC
       arrPRODVAREC = objCADPRODUTO.PRODVAREC
       arrPRODVAINSP = objCADPRODUTO.PRODVAINSP
       
       arrPRODATRPROC = objCADPRODUTO.PRODATRPROC
       arrPRODATRREC = objCADPRODUTO.PRODATRREC
       arrPRODATRINSP = objCADPRODUTO.PRODATRINSP
       arrCORES = objCADPRODUTO.CORES
       arrVERNIZ = objCADPRODUTO.VERNIZ
       arrVERNIZACAB = objCADPRODUTO.VERNIZACAB
       arrESMALTE = objCADPRODUTO.ESMALTE
       arrVEDANTE = objCADPRODUTO.VEDANTE
       arrPRODATECLIE = objCADPRODUTO.PRODATECLIE
       arrVERNIZ02 = objCADPRODUTO.VERNIZ02
       
       Call CarregaGridCores
       Call CarregaGridVerniz
       Call CarregaGridVerniz02
       Call CarregaGridVernizAcab
       Call CarregaGridEsmalte
       Call CarregaVernizExterno
       Call CarregaGridVedante
       Call CarregaGridCliente
       
       strCodProc = ""
       
       txtMETROS.Text = Format(objCADPRODUTO.METROS, "#,###0.000")
       txtLARGURA.Text = Format(objCADPRODUTO.LARGURA, "#,###0.000")
       txtGRAMATURAM2.Text = Format(objCADPRODUTO.GRAMATURA2, "#.###0.000")
       
       Call PopUnidConv

       If objCADPRODUTO.EspessCorpo > 0 Then txtCorpoEspess.Text = Format(objCADPRODUTO.EspessCorpo, "#,##0.00")
       If objCADPRODUTO.EspessTampa > 0 Then txtTampaEspess.Text = Format(objCADPRODUTO.EspessTampa, "#,##0.00")
       If objCADPRODUTO.EspessFundo > 0 Then txtFundoEspess.Text = Format(objCADPRODUTO.EspessFundo, "#,##0.00")
       If objCADPRODUTO.EspessArgola > 0 Then txtArgolaEspess.Text = Format(objCADPRODUTO.EspessArgola, "#,##0.00")
       
       If objCADPRODUTO.RevestCorpo > 0 Then txtCorpoRevest.Text = Format(objCADPRODUTO.RevestCorpo, "#,##0.00")
       If objCADPRODUTO.RevestCorpo2 > 0 Then txtCorpoRevest2.Text = Format(objCADPRODUTO.RevestCorpo2, "#,##0.00")
       If objCADPRODUTO.RevestTampa > 0 Then txtTampaRevest.Text = Format(objCADPRODUTO.RevestTampa, "#,##0.00")
       If objCADPRODUTO.RevestTampa2 > 0 Then txtTampaRevest2.Text = Format(objCADPRODUTO.RevestTampa2, "#,##0.00")
       If objCADPRODUTO.RevestFundo > 0 Then txtFundoRevest.Text = Format(objCADPRODUTO.RevestFundo, "#,##0.00")
       If objCADPRODUTO.RevestFundo2 > 0 Then txtFundoRevest2.Text = Format(objCADPRODUTO.RevestFundo2, "#,##0.00")
       If objCADPRODUTO.RevestArgola > 0 Then txtArgolaRevest.Text = Format(objCADPRODUTO.RevestArgola, "#,##0.00")
       If objCADPRODUTO.RevestArgola2 > 0 Then txtArgolaRevest2.Text = Format(objCADPRODUTO.RevestArgola2, "#,##0.00")
    
       '' ======================================================
       '' Montagem
       For i = 0 To (cboAlcaPlastica.ListCount - 1)
           If cboAlcaPlastica.ItemData(i) = objCADPRODUTO.AlcaPlastica Then cboAlcaPlastica.ListIndex = i
       Next i
       For i = 0 To (cboAlcaFerro.ListCount - 1)
           If cboAlcaFerro.ItemData(i) = objCADPRODUTO.AlcaFerro Then cboAlcaFerro.ListIndex = i
       Next i
       optPipetSimNao(objCADPRODUTO.Pipeta).Value = True
       optAzSimNao(objCADPRODUTO.Azelha).Value = True
       
       Call PegaProdVerCol_ColEsp(objCADPRODUTO.VerCat, objCADPRODUTO.ColEsp)
       '' ======================================================
       
       '' ======================================================
       '' Fechamento
       arrTAMPAPRES = objCADPRODUTO.TAMPAPRESS
       Call PopGrdFechamentoCampos(objCADPRODUTO.TampaPressao)
       
       arrBATRETR = objCADPRODUTO.BATRETRATI
       Call PopGrdFechamentoCampos(objCADPRODUTO.BatoqueRetratil)
       
       arrBATPLAST = objCADPRODUTO.BATPLASTIC
       Call PopGrdFechamentoCampos(objCADPRODUTO.BatoquePlastico)
       
       arrTAMPAVISOR = objCADPRODUTO.TAMPAVISOR
       Call PopGrdFechamentoCampos(objCADPRODUTO.TAMPAVIS)
       
       For i = 0 To (cboFechSoldaAgraf.ListCount - 1)
           If cboFechSoldaAgraf.ItemData(i) = objCADPRODUTO.FechSoldaAgrafado Then cboFechSoldaAgraf.ListIndex = i
       Next i
       
       '' ======================================================
    
       '' Expedição
       If objCADPRODUTO.TipoPalletGranel > 0 Then optTipPalet(objCADPRODUTO.TipoPalletGranel).Value = True
       If objCADPRODUTO.QtdePalletGranel > 0 Then txtQtdeEmb.Text = objCADPRODUTO.QtdePalletGranel
       '' ======================================================
    
       Call PopGrdFamMaq
    
        Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADGRUPROD", txtCODGRUP.Text, lblDescFamProd, False)
        Call PegaDescTabelasSubFamProd(txtCODGRUP.Text, txtCODSUBGRUP.Text, lblDescSubFamProd, False)
        Call PegaDescTabelasEspProd(txtCODSUBGRUP.Text, txtEspecie.Text, lblDescEspecieProd, False)
        Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADTIPPROD", txtTipo.Text, lblDescTipoProd, False)
    
        Call PopGrdEstoque(objCADPRODUTO.IDProduto)
        Call PopEstoqueLitografado(objCADPRODUTO.IDProduto)
    
        optAlcGalSN(objCADPRODUTO.ALCAGALAO).Value = True
        optNeckInSN(objCADPRODUTO.NECKIN).Value = True
        optDimCortePAD(objCADPRODUTO.DIMPADRAO).Value = True
    
        If Len(Trim(objCADPRODUTO.DESENV)) > 0 Then txtDESENV.Text = Format(objCADPRODUTO.DESENV, "#,##0.00")
        If Len(Trim(objCADPRODUTO.ALTURA)) > 0 Then txtALTURA.Text = Format(objCADPRODUTO.ALTURA, "#,##0.00")
    
        optFilialPed(objCADPRODUTO.FILIALPED).Value = True
    
        Call CarregaImagem
        Call PopGrdProdEntr
    
        optQTDCORPPADRAOSN(objCADPRODUTO.QTDCORPSPADRAOSN).Value = True
        optProdNovoSN(objCADPRODUTO.PRODNOVO).Value = True
    
    
        Me.Caption = Me.Caption & " / [ Produto : " & Trim(txtCodigo.Text) & " - " & Trim(txtDescricao.Text) & "]"
       
        '' Carrega a Estrutura
        Call Pop_Estrutura
    
    End If

    Frame17.Enabled = False
    
End Sub


Private Function ValidaCoef(strCoef As String) As Boolean
    ValidaCoef = True
    If Not IsNumeric(strCoef) Then ValidaCoef = False
End Function

Private Function ValidaCoefPos(curVal1 As Currency, curVal2 As Currency) As Boolean
    ValidaCoefPos = True
    If curVal2 <= curVal1 Then ValidaCoefPos = False
End Function


Private Function CamposOKGridVarContProc(ByVal COEF As TextBox, ByVal COEFP As TextBox, ByVal COEFN As TextBox, ByVal CODVAR As TextBox) As Boolean
    CamposOKGridVarContProc = False
    
    If Not IsNumeric(CODVAR.Text) Then
       MsgBox "Informe o Código da Variante !!!", vbOKOnly + vbExclamation, "Aviso"
       CODVAR.SetFocus
       Exit Function
    End If
    
    If IsNumeric(COEF.Text) And Not IsNumeric(COEFP.Text) And Not IsNumeric(COEFN.Text) Then
       MsgBox "Informe os Campos 'Tolerância +','Tolerância -' !!!", vbOKOnly + vbExclamation, "Aviso"
       COEF.SetFocus
       Exit Function
    End If
    If IsNumeric(COEF.Text) And Not IsNumeric(COEFP.Text) And IsNumeric(COEFN.Text) Then
       MsgBox "Informe o Campo 'Tolerância +' !!!", vbOKOnly + vbExclamation, "Aviso"
       COEF.SetFocus
       Exit Function
    End If
    If IsNumeric(COEF.Text) And IsNumeric(COEFP.Text) And Not IsNumeric(COEFN.Text) Then
       MsgBox "Informe o Campo 'Tolerância -' !!!", vbOKOnly + vbExclamation, "Aviso"
       COEF.SetFocus
       Exit Function
    End If
    If Not IsNumeric(COEF.Text) And IsNumeric(COEFP.Text) And Not IsNumeric(COEFN.Text) Then
       MsgBox "Informe os Campos 'Coeficiente','Tolerância -' !!!", vbOKOnly + vbExclamation, "Aviso"
       COEF.SetFocus
       Exit Function
    End If
    If Not IsNumeric(COEF.Text) And Not IsNumeric(COEFP.Text) And IsNumeric(COEFN.Text) Then
       MsgBox "Informe os Campos 'Coeficiente','Tolerância +' !!!", vbOKOnly + vbExclamation, "Aviso"
       COEF.SetFocus
       Exit Function
    End If
    If Not IsNumeric(COEF.Text) And Not IsNumeric(COEFP.Text) And Not IsNumeric(COEFN.Text) Then
       MsgBox "Informe os Campos 'Coeficiente','Tolerância +','Tolerância -' !!!", vbOKOnly + vbExclamation, "Aviso"
       COEF.SetFocus
       Exit Function
    End If
    If Not IsNumeric(COEF.Text) And IsNumeric(COEFP.Text) And IsNumeric(COEFN.Text) Then
       MsgBox "Informe os Campos 'Coeficiente'!!!", vbOKOnly + vbExclamation, "Aviso"
       COEF.SetFocus
       Exit Function
    End If
    
    CamposOKGridVarContProc = True
End Function


Private Function CamposOKGridAtributos(ByVal CODIGO As TextBox, ByVal ESPECIFICACAO As TextBox, ByVal CONDICAO As TextBox) As Boolean
    CamposOKGridAtributos = False
    If Not IsNumeric(CODIGO.Text) Then
       MsgBox "Informe o Código do Atributo !!!", vbOKOnly + vbExclamation, "Aviso"
       CODIGO.SetFocus
       Exit Function
    End If
    If Len(Trim(ESPECIFICACAO.Text)) = 0 Then
       MsgBox "Informa a Especificação do Atribuito !!!", vbOKOnly + vbExclamation, "Aviso"
       ESPECIFICACAO.SetFocus
       Exit Function
    End If
    If Len(Trim(CONDICAO.Text)) = 0 Then
       MsgBox "Informa a Condição do Atribuito !!!", vbOKOnly + vbExclamation, "Aviso"
       CONDICAO.SetFocus
       Exit Function
    End If
    CamposOKGridAtributos = True
End Function

Private Function PesqPadrao() As Boolean

    PesqPadrao = False
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO <> '" & objCADPRODUTO.CodigoProd & "'"
    sSql = sSql & "   And SGI_PADRAO  = 1"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PesqPadrao = True
    BREC.Close
    
End Function


Private Function PegaDescrPadrao(lngCodUsuario As Long) As String
    PegaDescrPadrao = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADDESCPADRAO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCodUsuario
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaDescrPadrao = BREC!SGI_DESCRI
    BREC.Close
    
End Function


Private Function PegaCodProdPadrao() As String
    PegaCodProdPadrao = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_CODIGO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_PADRAO = 1"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaCodProdPadrao = BREC!SGI_CODIGO
    BREC.Close
    
End Function

Private Function PegaValParaPadrao(strCODPROD As String, lngCODPARAM As Long) As Currency

    PegaValParaPadrao = 0
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       CE.SGI_QTDE " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPARCOEF      CE " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       CE.SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "   And CE.SGI_CODPROD  = '" & strCODPROD & "'" & vbCrLf
    sSql = sSql & "   And CE.SGI_CODPARAM = " & lngCODPARAM
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaValParaPadrao = BREC!SGI_QTDE
    BREC.Close
End Function

Private Function PegaValPesos(strCODPROD As String, lngCODPARAM As Long) As Currency

    PegaValPesos = 0
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       CE.SGI_PESO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPARPRODPADRAO CE " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       CE.SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "   And CE.SGI_CODPROD  = '" & Trim(strCODPROD) & "'" & vbCrLf
    sSql = sSql & "   And CE.SGI_CODPAR   = " & lngCODPARAM
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaValPesos = BREC!SGI_PESO
    BREC.Close

End Function

Private Sub InitGridUnidConv()

    With grdUnidConv
    
       .Cols = conColumnsIn_SonUnid
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonUnid_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonUnid_CodUnid) = ""
       .ColDataType(conCOL_SonUnid_CodUnid) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonUnid_PesqUnid) = ""
       .ColDataType(conCOL_SonUnid_PesqUnid) = flexDTString
       .ColComboList(conCOL_SonUnid_PesqUnid) = "..."
       
       .Cell(flexcpData, 0, conCOL_SonUnid_Unidade) = ""
       .ColDataType(conCOL_SonUnid_Unidade) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonUnid_Desc_Unid) = ""
       .ColDataType(conCOL_SonUnid_Desc_Unid) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonUnid_Fator) = ""
       .ColDataType(conCOL_SonUnid_Fator) = flexDTCurrency
       
       .ColWidth(conCOL_SonUnid_CodUnid) = 1500
       .ColWidth(conCOL_SonUnid_PesqUnid) = 300
       .ColWidth(conCOL_SonUnid_Unidade) = 800
       .ColWidth(conCOL_SonUnid_Desc_Unid) = 4000
       .ColWidth(conCOL_SonUnid_Fator) = 1500
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Function VerifItensRepetidosuUnidConv(intRow As Long, intCol As Long, varCampo As Variant) As Boolean
    VerifItensRepetidosuUnidConv = False
    Dim i As Integer
    
    If Not IsNumeric(varCampo) Then varCampo = UCase(Trim(varCampo))
    
    For i = 1 To (grdUnidConv.Rows - 1)
        If i <> intRow And grdUnidConv.Cell(flexcpText, i, intCol) = varCampo Then Exit Function
    Next i
    VerifItensRepetidosuUnidConv = True
End Function


Private Function PegaDescrUnidade(lngCodUsuario As Long) As String
    PegaDescrUnidade = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADUNIMED " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCodUsuario
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaDescrUnidade = BREC!SGI_DESCRICAO
    BREC.Close
    
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

Private Sub IncRegGridUnidades()
   
    If ExisteLinhaVaziaUnidade = False Then Exit Sub
    
    grdUnidConv.AddItem "" & vbTab & _
                        "" & vbTab & _
                        "" & vbTab & _
                        "" & vbTab & _
                        ""
                          
End Sub

Private Function ExisteLinhaVaziaUnidade() As Boolean
    ExisteLinhaVaziaUnidade = False
    
    Dim i As Integer
    
    For i = 1 To (grdUnidConv.Rows - 1)
        If grdUnidConv.Cell(flexcpText, i, conCOL_SonUnid_CodUnid) = Empty Then Exit Function
    Next i
    
    ExisteLinhaVaziaUnidade = True
End Function

Private Sub PopUnidConv()

    Dim i As Integer
    
    arrUNIDCONV = objCADPRODUTO.UNIDCONV
    
    If IsArray(arrUNIDCONV) Then
       For i = 1 To UBound(arrUNIDCONV)
           grdUnidConv.AddItem arrUNIDCONV(i, 1) & vbTab & _
                               "" & vbTab & _
                               PegaUnidade(CLng(arrUNIDCONV(i, 1))) & vbTab & _
                               PegaDescrUnidade(CLng(arrUNIDCONV(i, 1))) & vbTab & _
                               Format(arrUNIDCONV(i, 2), "#,####0.0000")
       Next i
    End If
End Sub

Private Sub LimpaCaptions()

    lblLinhProd.Caption = ""
    lblDesclie.Caption = ""
    Label1(14).Caption = ""
    
End Sub

Private Function PegaDescClie(lngCodClie As Long) As String

    PegaDescClie = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
    sSql = sSql & "  Where " & vbCrLf
    sSql = sSql & "        SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "    And SGI_CODIGO = " & lngCodClie
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaDescClie = BREC!SGI_RAZAOSOC
    BREC.Close
    
End Function

Private Function PegaDescLinProd(lngCodLinProd As Long) As String

    PegaDescLinProd = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO " & vbCrLf
    sSql = sSql & "  Where " & vbCrLf
    sSql = sSql & "        SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "    And SGI_CODLIN = " & lngCodLinProd
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
       PegaDescLinProd = BREC!SGI_DESCRI
       ''objCADPRODUTO.FILIALPED = BREC!SGI_FILIALPED
       optFilialPed(BREC!SGI_FILIALPED).Value = True
    End If
    BREC.Close
    
    ''optFilialPed(objCADPRODUTO.FILIALPED).Value = True
    
End Function


Private Sub FormaCodProd()
    
    Dim lngCodLinProd   As Long
    Dim lngCodClie      As Long
    Dim lngCodRotulo    As Long
    Dim lngDigito       As Long
    
    lngCodLinProd = 0
    lngCodClie = 0
    lngCodRotulo = 0
    lngDigito = 0
    
    If Len(Trim(txtLinProd.Text)) > 0 Then lngCodLinProd = CLng(txtLinProd.Text)
    If Len(Trim(txtCodCliente.Text)) > 0 Then lngCodClie = CLng(txtCodCliente.Text)
    If Len(Trim(txtCodRot.Text)) > 0 Then lngCodRotulo = CLng(txtCodRot.Text)
    If Len(Trim(txtDigVerif.Text)) > 0 Then lngDigito = CLng(txtDigVerif.Text)
    
    Label1(14).Caption = Format(lngCodLinProd, "###000")
    Label1(14).Caption = Label1(14).Caption & "." & Format(lngCodClie, "####0000")
    Label1(14).Caption = Label1(14).Caption & "." & Format(lngCodRotulo, "##00")
    Label1(14).Caption = Label1(14).Caption & "." & Format(lngDigito, "#0")
    
    txtCodigo.Text = Label1(14).Caption
    
End Sub


Private Sub InitGridVernizEsm()

    With grdVernizEsm
    
       .Cols = conColumnsIn_SonVerniz
       .Rows = 6
       .FixedCols = 1
       .FormatString = conCOL_SonVerniz_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonVerniz_Item) = ""
       .ColDataType(conCOL_SonVerniz_Item) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonVerniz_CodTipo) = ""
       .ColDataType(conCOL_SonVerniz_CodTipo) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonVerniz_PesqVerniz) = ""
       .ColDataType(conCOL_SonVerniz_PesqVerniz) = flexDTString
       .ColComboList(conCOL_SonVerniz_PesqVerniz) = "..."
       
       .Cell(flexcpData, 0, conCOL_SonVerniz_DescVerniz) = ""
       .ColDataType(conCOL_SonVerniz_DescVerniz) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonVerniz_Codigo) = ""
       .ColDataType(conCOL_SonVerniz_Codigo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonVerniz_UnidMed) = ""
       .ColDataType(conCOL_SonVerniz_UnidMed) = flexDTString
       .ColComboList(conCOL_SonVerniz_UnidMed) = objCADPRODUTO.PreenchComboUnidMedGrid
       
       .Cell(flexcpData, 0, conCOL_SonVerniz_Qtde) = ""
       .ColDataType(conCOL_SonVerniz_Qtde) = flexDTDouble
       
       .ColWidth(conCOL_SonVerniz_Item) = 3000
       .ColWidth(conCOL_SonVerniz_CodTipo) = 1500
       .ColWidth(conCOL_SonVerniz_PesqVerniz) = 300
       .ColWidth(conCOL_SonVerniz_DescVerniz) = 5000
       .ColWidth(conCOL_SonVerniz_Codigo) = 0
       .ColWidth(conCOL_SonVerniz_UnidMed) = 1500
       .ColWidth(conCOL_SonVerniz_Qtde) = 1000
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
       .TextMatrix(1, 0) = "Verniz Interno - 01"
       .TextMatrix(2, 0) = "Verniz Interno - 02"
       .TextMatrix(3, 0) = "Verniz de Acabamento"
       .TextMatrix(4, 0) = "Esmalte"
       .TextMatrix(5, 0) = "Verniz Externo"
       
       .Cell(flexcpData, 1, 0) = 0
       .Cell(flexcpData, 3, 0) = 1
       .Cell(flexcpData, 2, 0) = 2
       .Cell(flexcpData, 4, 0) = 3
       .Cell(flexcpData, 5, 0) = 4
    
    End With
    
End Sub

Private Sub InitGridCores()

    With grdCores
    
       .Cols = conColumnsIn_SonCores
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonCores_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonCores_Codigo) = ""
       .ColDataType(conCOL_SonCores_Codigo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonCores_CodCor) = ""
       .ColDataType(conCOL_SonCores_CodCor) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonCores_PesqCor) = ""
       .ColDataType(conCOL_SonCores_PesqCor) = flexDTString
       .ColComboList(conCOL_SonCores_PesqCor) = "..."
       
       .Cell(flexcpData, 0, conCOL_SonCores_DescCores) = ""
       .ColDataType(conCOL_SonCores_DescCores) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonCores_Ordem) = ""
       .ColDataType(conCOL_SonCores_Ordem) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonCores_UniMed) = ""
       .ColDataType(conCOL_SonCores_UniMed) = flexDTString
       .ColComboList(conCOL_SonCores_UniMed) = objCADPRODUTO.PreenchComboUnidMedGrid
       
       .Cell(flexcpData, 0, conCOL_SonCores_Qtde) = ""
       .ColDataType(conCOL_SonCores_Qtde) = flexDTCurrency
       
       .ColWidth(conCOL_SonCores_Codigo) = 0
       .ColWidth(conCOL_SonCores_CodCor) = 1500
       .ColWidth(conCOL_SonCores_PesqCor) = 300
       .ColWidth(conCOL_SonCores_DescCores) = 5000
       .ColWidth(conCOL_SonCores_Ordem) = 1000
       .ColWidth(conCOL_SonCores_UniMed) = 1500
       .ColWidth(conCOL_SonCores_Qtde) = 1000
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
    End With
    
End Sub

Private Sub IncRegGridCores()
   
    If objBLBFunc.FcExisteLinhaVazia(grdCores, conCOL_SonCores_CodCor) = False Then Exit Sub
    
    ''If ConsisteGrid = False Then Exit Sub
    
    grdCores.AddItem "" & vbTab & _
                     "" & vbTab & _
                     "" & vbTab & _
                     "" & vbTab & _
                     "" & vbTab & _
                     ""
                          
    Call Refaz_IndCores
                            
End Sub


Private Function PegaDescrProduto(lngCodProduto As String) As String
    PegaDescrProduto = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = '" & Trim(UCase(lngCodProduto)) & "'"
    
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC2.EOF Then PegaDescrProduto = BREC2!SGI_DESCRICAO
    BREC2.Close
    
End Function


Private Sub CarregaGridCores()

    Dim i As Integer
    
    If IsArray(arrCORES) = True Then
       For i = 1 To UBound(arrCORES)
       
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  From " & vbCrLf
           sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
           sSql = sSql & "   And SGI_IDPRODUTO = " & arrCORES(i, 1)
           
           BREC.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC.EOF Then
              grdCores.AddItem BREC!SGI_IDPRODUTO & vbTab & _
                               BREC!SGI_CODIGO & vbTab & _
                               "" & vbTab & _
                               BREC!SGI_DESCRICAO & vbTab & _
                               arrCORES(i, 2) & vbTab & _
                               arrCORES(i, 3) & vbTab & _
                               ""
                               
              If Len(Trim(arrCORES(i, 4))) > 0 Then grdCores.Cell(flexcpText, (grdCores.Rows - 1), conCOL_SonCores_Qtde) = Format(arrCORES(i, 4), "#,####0.0000")

           End If
           BREC.Close
       
       Next i
    End If

End Sub

Private Sub ConfCombosVerniz()

    cboCorpoVerniz.Clear
    cboTampaVerniz.Clear
    cboFundoVerniz.Clear
    cboArgolaVerniz.Clear
   
    If Len(Trim(txtEspecie.Text)) = 0 Then Exit Sub
    
    Dim arrVERFECH()    As String
    ReDim arrVERFECH(1 To 4) As String
    
    arrVERFECH(1) = "VEX - Verniz Externo"
    arrVERFECH(2) = "VZ  - Verniz dos 2 Lados"
    arrVERFECH(3) = "NAT - Natural"
    arrVERFECH(4) = "VI - Verniz interno Apenas"
    
    With cboCorpoVerniz
        sSql = ""
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       SGI_COD" & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADESPROD_CORPO" & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_CODIGO = " & Trim(txtEspecie.Text)
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        
        Do While Not BREC.EOF()
            .AddItem arrVERFECH(BREC!SGI_COD)
            .ItemData(.NewIndex) = BREC!SGI_COD
            BREC.MoveNext
        Loop
        BREC.Close
    End With
    '' ===================================================
    
    With cboTampaVerniz
        sSql = ""
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       SGI_COD" & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADESPROD_TAMPA" & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_CODIGO = " & txtEspecie.Text
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        
        Do While Not BREC.EOF()
            .AddItem arrVERFECH(BREC!SGI_COD)
            .ItemData(.NewIndex) = BREC!SGI_COD
            BREC.MoveNext
        Loop
        BREC.Close
    End With
    '' ===================================================
    
    With cboFundoVerniz
        sSql = ""
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       SGI_COD" & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADESPROD_FUNDO" & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_CODIGO = " & txtEspecie.Text
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        
        Do While Not BREC.EOF()
            .AddItem arrVERFECH(BREC!SGI_COD)
            .ItemData(.NewIndex) = BREC!SGI_COD
            BREC.MoveNext
        Loop
        BREC.Close
    End With
    '' ===================================================
    
    With cboArgolaVerniz
        sSql = ""
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       SGI_COD" & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADESPROD_ARGOLA" & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_CODIGO = " & txtEspecie.Text
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        
        Do While Not BREC.EOF()
            .AddItem arrVERFECH(BREC!SGI_COD)
            .ItemData(.NewIndex) = BREC!SGI_COD
            BREC.MoveNext
        Loop
        BREC.Close
    End With
    '' ===================================================
    
End Sub

Private Function PegaIDProduto(strCODPRODUTO As String) As String

    PegaIDProduto = "0"
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_CODIGO = '" & Trim(UCase(strCODPRODUTO)) & "'" & vbCrLf
    sSql = sSql & "   And SGI_FILIAL = " & FILIAL

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaIDProduto = BREC!SGI_IDPRODUTO
    BREC.Close
    
End Function

Private Sub CarregaGridVerniz()

    Dim i As Integer
    
    If IsArray(arrVERNIZ) = True Then
       For i = 1 To UBound(arrVERNIZ)
       
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  From " & vbCrLf
           sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
           sSql = sSql & "   And SGI_IDPRODUTO = " & arrVERNIZ(i, 1)
           
           BREC.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC.EOF Then
              grdVernizEsm.Cell(flexcpText, 1, conCOL_SonVerniz_Codigo) = BREC!SGI_IDPRODUTO
              grdVernizEsm.Cell(flexcpText, 1, conCOL_SonVerniz_CodTipo) = BREC!SGI_CODIGO
              grdVernizEsm.Cell(flexcpText, 1, conCOL_SonVerniz_DescVerniz) = BREC!SGI_DESCRICAO
              grdVernizEsm.Cell(flexcpText, 1, conCOL_SonVerniz_UnidMed) = arrVERNIZ(i, 2)
              If Len(Trim(arrVERNIZ(i, 3))) > 0 Then grdVernizEsm.Cell(flexcpText, 1, conCOL_SonVerniz_Qtde) = Format(arrVERNIZ(i, 3), "#,####0.0000")
           End If
           BREC.Close
       
       Next i
    End If

End Sub


Private Sub CarregaGridVernizAcab()

    Dim i As Integer
    
    If IsArray(arrVERNIZACAB) = True Then
       For i = 1 To UBound(arrVERNIZACAB)
       
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  From " & vbCrLf
           sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
           sSql = sSql & "   And SGI_IDPRODUTO = " & arrVERNIZACAB(i, 1)
           
           BREC.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC.EOF Then
              grdVernizEsm.Cell(flexcpText, 3, conCOL_SonVerniz_Codigo) = BREC!SGI_IDPRODUTO
              grdVernizEsm.Cell(flexcpText, 3, conCOL_SonVerniz_CodTipo) = BREC!SGI_CODIGO
              grdVernizEsm.Cell(flexcpText, 3, conCOL_SonVerniz_DescVerniz) = BREC!SGI_DESCRICAO
              grdVernizEsm.Cell(flexcpText, 3, conCOL_SonVerniz_UnidMed) = arrVERNIZACAB(i, 2)
              If Len(Trim(arrVERNIZACAB(i, 3))) > 0 Then grdVernizEsm.Cell(flexcpText, 3, conCOL_SonVerniz_Qtde) = Format(arrVERNIZACAB(i, 3), "#,####0.0000")
           
           End If
           BREC.Close
       
       Next i
    End If

End Sub

Private Sub CarregaGridEsmalte()

    Dim i As Integer
    
    If IsArray(arrESMALTE) = True Then
       For i = 1 To UBound(arrESMALTE)
       
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  From " & vbCrLf
           sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
           sSql = sSql & "   And SGI_IDPRODUTO = " & arrESMALTE(i, 1)
           
           BREC.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC.EOF Then
              grdVernizEsm.Cell(flexcpText, 4, conCOL_SonVerniz_Codigo) = BREC!SGI_IDPRODUTO
              grdVernizEsm.Cell(flexcpText, 4, conCOL_SonVerniz_CodTipo) = BREC!SGI_CODIGO
              grdVernizEsm.Cell(flexcpText, 4, conCOL_SonVerniz_DescVerniz) = BREC!SGI_DESCRICAO
              grdVernizEsm.Cell(flexcpText, 4, conCOL_SonVerniz_UnidMed) = arrESMALTE(i, 2)
              If Len(Trim(arrESMALTE(i, 3))) > 0 Then grdVernizEsm.Cell(flexcpText, 4, conCOL_SonVerniz_Qtde) = Format(arrESMALTE(i, 3), "#,####0.0000")
           End If
           BREC.Close
       
       Next i
    End If

End Sub

Private Sub PreenchComboAlca()

    cboAlcaPlastica.Clear
    
    cboAlcaPlastica.AddItem "3D"
    cboAlcaPlastica.ItemData(cboAlcaPlastica.NewIndex) = 1
    cboAlcaPlastica.AddItem "4D"
    cboAlcaPlastica.ItemData(cboAlcaPlastica.NewIndex) = 2
    
    cboAlcaFerro.AddItem "3D"
    cboAlcaFerro.ItemData(cboAlcaFerro.NewIndex) = 1
    cboAlcaFerro.AddItem "4D"
    cboAlcaFerro.ItemData(cboAlcaFerro.NewIndex) = 2
    
End Sub

Public Sub ConfGridFecham()
    
    cboFechSoldaAgraf.Clear

    If Len(Trim(txtLinProd.Text)) = 0 Then Exit Sub
    
    Dim i               As Integer
    Dim arrFECHAMENTO() As String
    
    ReDim arrFECHAMENTO(0 To 2) As String
    
    arrFECHAMENTO(0) = "SOLDA"
    arrFECHAMENTO(1) = "AGRAFADO"
    arrFECHAMENTO(2) = "REPUXO"
    
    With cboFechSoldaAgraf
    
        sSql = ""
        
        sSql = sSql & "Select " & vbCrLf
        sSql = sSql & "       FECH.*" & vbCrLf
        
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADLINHAPRODUTO      LPRO" & vbCrLf
        sSql = sSql & "     , SGI_CADLINHAPRODUTO_FECH FECH" & vbCrLf
        
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       LPRO.SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And LPRO.SGI_CODLIN = " & txtLinProd.Text & vbCrLf
        sSql = sSql & "   And FECH.SGI_FILIAL = LPRO.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And FECH.SGI_CODIGO = LPRO.SGI_CODIGO"
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        Do While Not BREC.EOF()
            .AddItem arrFECHAMENTO(BREC!SGI_COD)
            .ItemData(.NewIndex) = BREC!SGI_COD
            BREC.MoveNext
        Loop
        BREC.Close
        
        If .ListCount = 1 Then .ListIndex = 0
    
    End With
End Sub

Private Sub InitGrdVedanteCompound()

    With grdVedanteCompound
    
       .Cols = conColumnsIn_SonVedante
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonVedante_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonVedante_Codigo) = ""
       .ColDataType(conCOL_SonVedante_Codigo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonVedante_CodProd) = ""
       .ColDataType(conCOL_SonVedante_CodProd) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonVedante_PesqProd) = ""
       .ColDataType(conCOL_SonVedante_PesqProd) = flexDTString
       .ColComboList(conCOL_SonVedante_PesqProd) = "..."
       
       .Cell(flexcpData, 0, conCOL_SonVedante_DescProd) = ""
       .ColDataType(conCOL_SonVedante_DescProd) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonVedante_UniMed) = ""
       .ColDataType(conCOL_SonVedante_UniMed) = flexDTString
       .ColComboList(conCOL_SonVedante_UniMed) = objCADPRODUTO.PreenchComboUnidMedGrid
       
       .Cell(flexcpData, 0, conCOL_SonVedante_Qtde) = ""
       .ColDataType(conCOL_SonVedante_Qtde) = flexDTCurrency
       
       .ColWidth(conCOL_SonVedante_Codigo) = 0
       .ColWidth(conCOL_SonVedante_CodProd) = 1500
       .ColWidth(conCOL_SonVedante_PesqProd) = 300
       .ColWidth(conCOL_SonVedante_DescProd) = 5000
       .ColWidth(conCOL_SonVedante_UniMed) = 1500
       .ColWidth(conCOL_SonVedante_Qtde) = 1000
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub IncRegGridVedante()
   
    If objBLBFunc.FcExisteLinhaVazia(grdVedanteCompound, conCOL_SonVedante_CodProd) = False Then Exit Sub
    
    ''If ConsisteGrid = False Then Exit Sub
    
    grdVedanteCompound.AddItem "" & vbTab & _
                               "" & vbTab & _
                               "" & vbTab & _
                               "" & vbTab & _
                               "" & vbTab & _
                               ""
                          
                            
End Sub


Private Sub CarregaGridVedante()

    Dim i As Integer
    
    If IsArray(arrVEDANTE) = True Then
       For i = 1 To UBound(arrVEDANTE)
       
           If Len(Trim(arrVEDANTE(i, 1))) > 0 Then
           
               sSql = "Select " & vbCrLf
               sSql = sSql & "       * " & vbCrLf
               sSql = sSql & "  From " & vbCrLf
               sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
               sSql = sSql & " Where " & vbCrLf
               sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
               sSql = sSql & "   And SGI_IDPRODUTO = " & arrVEDANTE(i, 1)
               
               BREC.Open sSql, adoBanco_Dados, adOpenDynamic
               If Not BREC.EOF Then
                  grdVedanteCompound.AddItem BREC!SGI_IDPRODUTO & vbTab & _
                                             BREC!SGI_CODIGO & vbTab & _
                                             "" & vbTab & _
                                             BREC!SGI_DESCRICAO & vbTab & _
                                             arrVEDANTE(i, 2) & vbTab & _
                                             ""

                 If Len(Trim(arrVEDANTE(i, 3))) > 0 Then grdVedanteCompound.Cell(flexcpText, (grdVedanteCompound.Rows - 1), conCOL_SonVedante_Qtde) = Format(arrVEDANTE(i, 3), "#,####0.0000")
    
               End If
               BREC.Close
           End If
       
       Next i
    End If

End Sub

Private Sub txtVerCat_GotFocus()
    objBLBFunc.SelecionaCampos txtVerCat.Name, frmCADPROD
End Sub

Private Sub txtVerCat_Validate(Cancel As Boolean)

   Dim i         As Integer
   Dim strIDProd As String

   If Len(Trim(txtVerCat.Text)) = 0 Then Exit Sub
   
   strIDProd = PegaIDProduto(Trim(txtVerCat.Text))
   
   If Len(Trim(strIDProd)) = 0 Then
      MsgBox "Este Produto Não existe !!!", vbOKOnly + vbExclamation, "Aviso"
      Cancel = True
      Exit Sub
   End If
   
   label11(15).Caption = PegaDescrProduto(Trim(txtVerCat.Text))
   objCADPRODUTO.VerCat = IIf(IsNumeric(strIDProd) = True, CLng(strIDProd), 0)
   
End Sub

Private Sub PegaProdVerCol_ColEsp(lngVerCol As Long, lngColEsp As Long)

        sSql = "Select " & vbCrLf
        sSql = sSql & "       * " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_IDPRODUTO = " & lngVerCol
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC.EOF() Then
           txtVerCat.Text = Trim(BREC!SGI_CODIGO)
           label11(15).Caption = PegaDescrProduto(Trim(txtVerCat.Text))
        End If
        BREC.Close


        sSql = "Select " & vbCrLf
        sSql = sSql & "       * " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_IDPRODUTO = " & lngColEsp
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC.EOF() Then
           txtColEsp.Text = Trim(BREC!SGI_CODIGO)
           label11(16).Caption = PegaDescrProduto(Trim(txtColEsp.Text))
        End If
        BREC.Close

End Sub


Private Sub InitgrdSomFecha()

    With grdSomFecha(conCOL_SomFecha_TampaPressao)
    
       .Cols = conColumnsIn_SomFecha
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SomFecha_TampaPressao_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SomFecha_Valor) = ""
       .ColDataType(conCOL_SomFecha_Valor) = flexDTString
       
       .ColWidth(conCOL_SomFecha_Valor) = 1400
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
    With grdSomFecha(conCOL_SomFecha_BatoqueRetra)
    
       .Cols = conColumnsIn_SomFecha
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SomFecha_BatoqueRetra_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SomFecha_Valor) = ""
       .ColDataType(conCOL_SomFecha_Valor) = flexDTString
       
       .ColWidth(conCOL_SomFecha_Valor) = 1300
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
    With grdSomFecha(conCOL_SomFecha_BatoquePlast)
    
       .Cols = conColumnsIn_SomFecha
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SomFecha_BatoquePlast_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SomFecha_Valor) = ""
       .ColDataType(conCOL_SomFecha_Valor) = flexDTString
       
       .ColWidth(conCOL_SomFecha_Valor) = 1400
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
    With grdSomFecha(conCOL_SomFecha_TampaVisor)
    
       .Cols = conColumnsIn_SomFecha
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SomFecha_TampaVisor_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SomFecha_Valor) = ""
       .ColDataType(conCOL_SomFecha_Valor) = flexDTString
       
       .ColWidth(conCOL_SomFecha_Valor) = 1200
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub


Private Sub IncRegSomFecha(intIndice As Integer)
   
    If objBLBFunc.FcExisteLinhaVazia(grdSomFecha(intIndice), conCOL_SomFecha_Valor) = False Then Exit Sub
    
    ''If ConsisteGrid = False Then Exit Sub
    
    grdSomFecha(intIndice).AddItem ""
                          
                            
End Sub

Public Sub PopGrdFechamentoCampos(intIndice As Integer)
    Dim i As Integer
    If intIndice = conCOL_SomFecha_TampaPressao Then
        If IsArray(arrTAMPAPRES) Then
            For i = 1 To UBound(arrTAMPAPRES)
                grdSomFecha(intIndice).AddItem Format(arrTAMPAPRES(i), "#,##0.00")
            Next i
        End If
    ElseIf intIndice = conCOL_SomFecha_BatoqueRetra Then
        If IsArray(arrBATRETR) Then
            For i = 1 To UBound(arrBATRETR)
                grdSomFecha(intIndice).AddItem Format(arrBATRETR(i), "#,##0.00")
            Next i
        End If
    ElseIf intIndice = conCOL_SomFecha_BatoquePlast Then
        If IsArray(arrBATPLAST) Then
            For i = 1 To UBound(arrBATPLAST)
                grdSomFecha(intIndice).AddItem Format(arrBATPLAST(i), "#,##0.00")
            Next i
        End If
    ElseIf intIndice = conCOL_SomFecha_TampaVisor Then
        If IsArray(arrTAMPAVISOR) Then
            For i = 1 To UBound(arrTAMPAVISOR)
                grdSomFecha(intIndice).AddItem Format(arrTAMPAVISOR(i), "#,##0.00")
            Next i
        End If
    End If
End Sub


Private Sub InitGrdProdAtend()

    With grdClientes
    
       .Cols = conColumnsIn_SonProdAte
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonProdAte_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonProdAte_CodigoClie) = ""
       .ColDataType(conCOL_SonProdAte_CodigoClie) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProdAte_PesqClie) = ""
       .ColDataType(conCOL_SonProdAte_PesqClie) = flexDTString
       .ColComboList(conCOL_SonProdAte_PesqClie) = "..."
       
       .Cell(flexcpData, 0, conCOL_SonProdAte_DescClie) = ""
       .ColDataType(conCOL_SonProdAte_DescClie) = flexDTString
       
       .ColWidth(conCOL_SonProdAte_CodigoClie) = 1500
       .ColWidth(conCOL_SonProdAte_PesqClie) = 300
       .ColWidth(conCOL_SonProdAte_DescClie) = 5000
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub IncRegGridProdAtend()
   
    If objBLBFunc.FcExisteLinhaVazia(grdClientes, conCOL_SonProdAte_CodigoClie) = False Then Exit Sub
    
    ''If ConsisteGrid = False Then Exit Sub
    
    grdClientes.AddItem "" & vbTab & _
                        "" & vbTab & _
                        ""
                            
End Sub

Private Function PegaDescrCliente(lngCodClie As Long) As String
    PegaDescrCliente = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCodClie
    
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC2.EOF Then PegaDescrCliente = BREC2!SGI_RAZAOSOC
    BREC2.Close
    
End Function


Private Sub CarregaGridCliente()

    Dim i As Integer
    
    If IsArray(arrPRODATECLIE) = True Then
       For i = 1 To UBound(arrPRODATECLIE)
       
           If Len(Trim(arrPRODATECLIE(i))) > 0 Then
           
               sSql = "Select " & vbCrLf
               sSql = sSql & "       * " & vbCrLf
               sSql = sSql & "  From " & vbCrLf
               sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
               sSql = sSql & " Where " & vbCrLf
               sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
               sSql = sSql & "   And SGI_CODIGO = " & arrPRODATECLIE(i)
               
               BREC.Open sSql, adoBanco_Dados, adOpenDynamic
               If Not BREC.EOF Then
                  grdClientes.AddItem BREC!SGI_CODIGO & vbTab & _
                                      "" & vbTab & _
                                      BREC!SGI_RAZAOSOC
    
               End If
               BREC.Close
           End If
       
       Next i
    End If

End Sub

Private Sub CarregaImagem()
    
    Dim strNOMARQ   As String
    Dim strCAMINHO  As String
    
    Image1.Picture = LoadPicture("")
    
    ''If Len(Trim(objCADPRODUTO.CAMINHO)) = 0 Then Exit Sub
    ''If Len(Trim(Dir(objCADPRODUTO.CAMINHO))) = 0 Then Exit Sub
    
    ''Image1.Picture = LoadPicture(objCADPRODUTO.CAMINHO)
    
    sSql = ""
    
    sSql = "Select SGI_IMAGEM from SGI_CADPROD_IMAGEN Where SGI_FILIAL = " & FILIAL & " And SGI_CODIGO = " & objCADPRODUTO.IDProduto
    BREC2.Open sSql, adoBanco_Dados_Imagem, adOpenDynamic, adLockOptimistic
    If Not BREC2.EOF Then
       If Not IsNull(BREC2!SGI_IMAGEM) Then
          strNOMARQ = "ImgProd" & objCADPRODUTO.IDProduto
          strCAMINHO = App.Path & "\"
          Call objBLBFunc.LeCampoBlobDoDB(BREC2, "SGI_IMAGEM", strCAMINHO + strNOMARQ)
          Image1.Picture = LoadPicture(strCAMINHO + strNOMARQ)
       End If
    End If
    BREC2.Close

End Sub

Private Sub CarregaClientes()

    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL =  " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO <> " & Trim(txtCodCliente.Text)
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC.EOF()
        
        With grdClientes
            .AddItem BREC!SGI_CODIGO & vbTab & _
                     "" & vbTab & _
                     Trim(BREC!SGI_RAZAOSOC)
        End With
    
        BREC.MoveNext
    Loop
    BREC.Close
    
End Sub

Private Sub DestroiObjeto()
    Set objBLBFunc = Nothing
    Set objCADPRODUTO = Nothing
    Set objCADFORN = Nothing
    Set objPESQPADRAO = Nothing
End Sub

Private Sub Consulta()

    Dim i As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = True
    
    Frame3.Enabled = True
    Frame38(3).Enabled = True
    Frame38(4).Enabled = True
    stProd.Enabled = True
    Frame14.Enabled = False
    Frame38(0).Enabled = True
    Frame38(1).Enabled = True
    Frame38(2).Enabled = False
    Frame6.Enabled = False
    Frame9.Enabled = False
    optProdNovoSN(1).Value = True
   
   
    Frame41(1).Enabled = True
    Frame38(5).Enabled = True
    Frame43(0).Enabled = True
    Frame43(1).Enabled = True
    Frame43(3).Enabled = True
    Frame8.Enabled = False
    
    txtCodRot.Enabled = False
   
    label11(15).Caption = ""
    label11(16).Caption = ""
    
    optFilialPed(0).Value = True
    
    txtSaldo.Enabled = False
    mskDtCadastro.Enabled = False
    
    stProd.Tab = 0
    
    txtCodigo.Enabled = True
    txtPRECOPRODUTO.Enabled = False
   
    Me.Caption = "Cadastro de produtos - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADPROD
    LimpaCaptions
    txtDescEquip.Locked = False
    
    objCADPRODUTO.PreenchComboUnidade cboUnidade
    objCADPRODUTO.PreenchComboClaFis cboClass
    objCADPRODUTO.PreenchComboUnidade cboUniIns
    cboUniIns.ListIndex = -1
    
    optATUAUTOMSIMNAO(0).Value = False
    optATUAUTOMSIMNAO(1).Value = False
    optAlcGalSN(0).Value = True
    optNeckInSN(0).Value = True
    optDimCortePAD(1).Value = True
    optQTDCORPPADRAOSN(1).Value = True
    
    lblPRECOMEDIO.Caption = ""
    
    Call InitGridUnidConv
    Call InitGridVernizEsm
    Call InitGridCores
    Call InitGrdVedanteCompound
    Call InitgrdSomFecha
    Call InitGrdProdAtend
    Call InitGrdFamMaq
    Call ConfEstoque
    Call ConfProdEntr
    Call ConfEstoqueLit
    
    Call ConfCombosVerniz
    Call ConfComboFech
    
    PreenchComboAlca
    ConfGridFecham
    
    objCADPRODUTO.IDProduto = iCodigo
    cboFechSoldaAgraf.ListIndex = -1
    
    objCADPRODUTO.TampaPressao = conCOL_SomFecha_TampaPressao
    objCADPRODUTO.BatoqueRetratil = conCOL_SomFecha_BatoqueRetra
    objCADPRODUTO.BatoquePlastico = conCOL_SomFecha_BatoquePlast
    objCADPRODUTO.TAMPAVIS = conCOL_SomFecha_TampaVisor
    
    txtCodigo.Enabled = False
    optUsarPadrPalSN(0).Value = True
    
    Call LimpaCamposLabel
    
    If objCADPRODUTO.Carrega_campos = True Then
    
        txtLinProd.Text = objCADPRODUTO.CodLinProd
        Call ConfGridFecham
        Call ConfComboFech
       
       txtIPI.Text = objCADPRODUTO.IPI
       
       txtCodCliente.Text = objCADPRODUTO.CodClie
       txtCodRot.Text = objCADPRODUTO.CodRotulo
       txtDigVerif.Text = objCADPRODUTO.DigVerif
       optNatSimNao(objCADPRODUTO.NATURALSIMNAO).Value = True
       
       Label1(23).Caption = ""
       If objCADPRODUTO.STATUS = 1 Then Label1(23).Caption = "ATIVO"
       If objCADPRODUTO.STATUS = 0 Then Label1(23).Caption = "DESATIVADO"
       If objCADPRODUTO.STATUS = 2 Then Label1(23).Caption = "AG.LIBERAÇÃO"
       
       Call FormaCodProd
       
       lblLinhProd.Caption = PegaDescLinProd(CLng(txtLinProd.Text))
       lblDesclie.Caption = PegaDescClie(CLng(txtCodCliente.Text))
       
       txtCodigo.Text = objCADPRODUTO.CodigoProd
       txtCodProdFornec.Text = objCADPRODUTO.CODPROFORNEC
       txtDescricao.Text = objCADPRODUTO.DescriProd
       txtEspecie.Text = Str(objCADPRODUTO.EspProduto)
       
       If Len(Trim(objCADPRODUTO.COMPLEMENTO)) > 0 Then txtComplemento.Text = Trim(objCADPRODUTO.COMPLEMENTO)
       
       
       '' Aba Produção
       optLaudoSN(objCADPRODUTO.EMITLAUDO).Value = True
       
       txtPOCACRES.Text = Format(objCADPRODUTO.PORCACRES, "#,##0.00")
       If objCADPRODUTO.PORCACRES > 0 Then txtPOCACRES.Text = Format(objCADPRODUTO.PORCACRES, "#,##0.00")
       If objCADPRODUTO.PRCCUSTO > 0 Then txtPRCFINAL.Text = Format(objCADPRODUTO.PRCCUSTO, "#,##0.00")
       If objCADPRODUTO.PESOUNIT > 0 Then txtPeso.Text = Format(objCADPRODUTO.PESOUNIT, "#,###0.000")
       
       If objCADPRODUTO.CUBAGEN > 0 Then txtCUBAGEN.Text = Format(objCADPRODUTO.CUBAGEN, "#,###0.000")
       
       If objCADPRODUTO.QTDEPORFOLHA > 0 Then txtQtdePorFolha.Text = objCADPRODUTO.QTDEPORFOLHA
       If objCADPRODUTO.QTDPASSADAS > 0 Then txtQtdePassada.Text = objCADPRODUTO.QTDPASSADAS
       
       If objCADPRODUTO.CODGRUPPROD > 0 Then txtCODGRUP.Text = objCADPRODUTO.CODGRUPPROD
       If objCADPRODUTO.CODSUBGPROD > 0 Then txtCODSUBGRUP.Text = objCADPRODUTO.CODSUBGPROD
       
        optUsarPadrPalSN(objCADPRODUTO.PALLHETPADRAO).Value = True
       
       
       txtTipo.Text = Str(objCADPRODUTO.TIPPRODUTO)
       txtSaldo.Text = Format(objCADPRODUTO.Saldo, "#0")
       txtEstMinimo.Text = Format(objCADPRODUTO.EstMin, "#0")
       mskDtCadastro.Text = objCADPRODUTO.DataCadast
       
       txtDescEquip.Text = objCADPRODUTO.DESCEQUIP
       txtProcedencia.Text = objCADPRODUTO.PROCEDEN
       
       txtRua.Text = objCADPRODUTO.RUA
       txtBox.Text = objCADPRODUTO.box
       txtPrateleira.Text = objCADPRODUTO.PRATELEIRA
       txtCodigoEAN.Text = objCADPRODUTO.CODEAN
       If objCADPRODUTO.DISTANCIA > 0 Then txtDistancia.Text = objCADPRODUTO.DISTANCIA
       
       
       If objCADPRODUTO.PRODUTOTIPO = 0 Then
          optTipProd(0).Value = True
          txtCodigo.Enabled = True
          txtDigVerif.Enabled = False
       End If
       If objCADPRODUTO.PRODUTOTIPO = 1 Then optTipProd(1).Value = True
       
       If objCADPRODUTO.ESTILOPROD = 0 Then optEstProduto(0).Value = True
       If objCADPRODUTO.ESTILOPROD = 1 Then optEstProduto(1).Value = True
       If objCADPRODUTO.ESTILOPROD = 2 Then optEstProduto(2).Value = True
       If objCADPRODUTO.ESTILOPROD = 3 Then optEstProduto(3).Value = True
       If objCADPRODUTO.ESTILOPROD = 4 Then optEstProduto(4).Value = True
       If objCADPRODUTO.ESTILOPROD = 5 Then optEstProduto(5).Value = True
              
       
       If objCADPRODUTO.DataUltMov > 0 Then mskDtUltMov.Text = Format(objCADPRODUTO.DataUltMov, "DD/MM/YYYY")
       arrProdLST = objCADPRODUTO.ProdLST
       arrPROCESSO = objCADPRODUTO.Processos
       arrFERRAMENTA = objCADPRODUTO.FERRAMENTA
       
       If objCADPRODUTO.PRCPROD > 0 Then txtPRECOPRODUTO.Text = Format(objCADPRODUTO.PRCPROD, "#,###0.000")
       If objCADPRODUTO.PRCMED > 0 Then lblPRECOMEDIO.Caption = Format(objCADPRODUTO.PRCMED, "#,###0.000")
       
       If objCADPRODUTO.ATAUTOMSN = 0 Then optATUAUTOMSIMNAO(0).Value = True
       If objCADPRODUTO.ATAUTOMSN = 1 Then optATUAUTOMSIMNAO(1).Value = True
       
       For i = 0 To (cboUnidade.ListCount - 1)
           If cboUnidade.ItemData(i) = objCADPRODUTO.Unidade Then cboUnidade.ListIndex = i
       Next i
       
       For i = 0 To (cboClass.ListCount - 1)
           If cboClass.ItemData(i) = objCADPRODUTO.CLASFISC Then cboClass.ListIndex = i
       Next i
       
       
       '' =============================================
       Call ConfCombosVerniz
       
       For i = 0 To (cboCorpoVerniz.ListCount - 1)
           If cboCorpoVerniz.ItemData(i) = objCADPRODUTO.VernCorpo Then cboCorpoVerniz.ListIndex = i
       Next i
       For i = 0 To (cboTampaVerniz.ListCount - 1)
           If cboTampaVerniz.ItemData(i) = objCADPRODUTO.VernTampa Then cboTampaVerniz.ListIndex = i
       Next i
       For i = 0 To (cboFundoVerniz.ListCount - 1)
           If cboFundoVerniz.ItemData(i) = objCADPRODUTO.VernFundo Then cboFundoVerniz.ListIndex = i
       Next i
       For i = 0 To (cboArgolaVerniz.ListCount - 1)
           If cboArgolaVerniz.ItemData(i) = objCADPRODUTO.VernArgola Then cboArgolaVerniz.ListIndex = i
       Next i


       For i = 0 To (cboFechTampaFuro.ListCount - 1)
           If cboFechTampaFuro.ItemData(i) = objCADPRODUTO.FechTampaFuro Then cboFechTampaFuro.ListIndex = i
       Next i
       '' =============================================
       
       If objCADPRODUTO.PRCPROD > 0 Then txtPRECOPRODUTO.Text = Format(objCADPRODUTO.PRCPROD, "#,##0.00")
       
       arrPRODTIPORCA = objCADPRODUTO.PRODTIPORCA
       
       arrPRODVAPROC = objCADPRODUTO.PRODVAPROC
       arrPRODVAREC = objCADPRODUTO.PRODVAREC
       arrPRODVAINSP = objCADPRODUTO.PRODVAINSP
       
       arrPRODATRPROC = objCADPRODUTO.PRODATRPROC
       arrPRODATRREC = objCADPRODUTO.PRODATRREC
       arrPRODATRINSP = objCADPRODUTO.PRODATRINSP
       arrCORES = objCADPRODUTO.CORES
       arrVERNIZ = objCADPRODUTO.VERNIZ
       arrVERNIZACAB = objCADPRODUTO.VERNIZACAB
       arrESMALTE = objCADPRODUTO.ESMALTE
       arrVEDANTE = objCADPRODUTO.VEDANTE
       arrPRODATECLIE = objCADPRODUTO.PRODATECLIE
       arrVERNIZ02 = objCADPRODUTO.VERNIZ02
       
       Call CarregaGridCores
       Call CarregaGridVerniz
       Call CarregaGridVerniz02
       Call CarregaGridVernizAcab
       Call CarregaVernizExterno
       Call CarregaGridEsmalte
       Call CarregaGridVedante
       Call CarregaGridCliente
       
       strCodProc = ""
       
       txtMETROS.Text = Format(objCADPRODUTO.METROS, "#,###0.000")
       txtLARGURA.Text = Format(objCADPRODUTO.LARGURA, "#,###0.000")
       txtGRAMATURAM2.Text = Format(objCADPRODUTO.GRAMATURA2, "#.###0.000")
       
       Call PopUnidConv

       If objCADPRODUTO.EspessCorpo > 0 Then txtCorpoEspess.Text = Format(objCADPRODUTO.EspessCorpo, "#,##0.00")
       If objCADPRODUTO.EspessTampa > 0 Then txtTampaEspess.Text = Format(objCADPRODUTO.EspessTampa, "#,##0.00")
       If objCADPRODUTO.EspessFundo > 0 Then txtFundoEspess.Text = Format(objCADPRODUTO.EspessFundo, "#,##0.00")
       If objCADPRODUTO.EspessArgola > 0 Then txtArgolaEspess.Text = Format(objCADPRODUTO.EspessArgola, "#,##0.00")
       
       If objCADPRODUTO.RevestCorpo > 0 Then txtCorpoRevest.Text = Format(objCADPRODUTO.RevestCorpo, "#,##0.00")
       If objCADPRODUTO.RevestCorpo2 > 0 Then txtCorpoRevest2.Text = Format(objCADPRODUTO.RevestCorpo2, "#,##0.00")
       If objCADPRODUTO.RevestTampa > 0 Then txtTampaRevest.Text = Format(objCADPRODUTO.RevestTampa, "#,##0.00")
       If objCADPRODUTO.RevestTampa2 > 0 Then txtTampaRevest2.Text = Format(objCADPRODUTO.RevestTampa2, "#,##0.00")
       If objCADPRODUTO.RevestFundo > 0 Then txtFundoRevest.Text = Format(objCADPRODUTO.RevestFundo, "#,##0.00")
       If objCADPRODUTO.RevestFundo2 > 0 Then txtFundoRevest2.Text = Format(objCADPRODUTO.RevestFundo2, "#,##0.00")
       If objCADPRODUTO.RevestArgola > 0 Then txtArgolaRevest.Text = Format(objCADPRODUTO.RevestArgola, "#,##0.00")
       If objCADPRODUTO.RevestArgola2 > 0 Then txtArgolaRevest2.Text = Format(objCADPRODUTO.RevestArgola2, "#,##0.00")
    
       '' ======================================================
       '' Montagem
       For i = 0 To (cboAlcaPlastica.ListCount - 1)
           If cboAlcaPlastica.ItemData(i) = objCADPRODUTO.AlcaPlastica Then cboAlcaPlastica.ListIndex = i
       Next i
       For i = 0 To (cboAlcaFerro.ListCount - 1)
           If cboAlcaFerro.ItemData(i) = objCADPRODUTO.AlcaFerro Then cboAlcaFerro.ListIndex = i
       Next i
       optPipetSimNao(objCADPRODUTO.Pipeta).Value = True
       optAzSimNao(objCADPRODUTO.Azelha).Value = True
       
       Call PegaProdVerCol_ColEsp(objCADPRODUTO.VerCat, objCADPRODUTO.ColEsp)
       '' ======================================================
       
       '' ======================================================
       '' Fechamento
       arrTAMPAPRES = objCADPRODUTO.TAMPAPRESS
       Call PopGrdFechamentoCampos(objCADPRODUTO.TampaPressao)
       
       arrBATRETR = objCADPRODUTO.BATRETRATI
       Call PopGrdFechamentoCampos(objCADPRODUTO.BatoqueRetratil)
       
       arrBATPLAST = objCADPRODUTO.BATPLASTIC
       Call PopGrdFechamentoCampos(objCADPRODUTO.BatoquePlastico)
       
       arrTAMPAVISOR = objCADPRODUTO.TAMPAVISOR
       Call PopGrdFechamentoCampos(objCADPRODUTO.TAMPAVIS)
       
       For i = 0 To (cboFechSoldaAgraf.ListCount - 1)
           If cboFechSoldaAgraf.ItemData(i) = objCADPRODUTO.FechSoldaAgrafado Then cboFechSoldaAgraf.ListIndex = i
       Next i
       
       '' ======================================================
    
       '' Expedição
       If objCADPRODUTO.TipoPalletGranel > 0 Then optTipPalet(objCADPRODUTO.TipoPalletGranel).Value = True
       If objCADPRODUTO.QtdePalletGranel > 0 Then txtQtdeEmb.Text = objCADPRODUTO.QtdePalletGranel
       '' ======================================================
    
       Call PopGrdFamMaq
       
        Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADGRUPROD", txtCODGRUP.Text, lblDescFamProd, False)
        Call PegaDescTabelasSubFamProd(txtCODGRUP.Text, txtCODSUBGRUP.Text, lblDescSubFamProd, False)
        Call PegaDescTabelasEspProd(txtCODSUBGRUP.Text, txtEspecie.Text, lblDescEspecieProd, False)
        Call PegaDescTabelas("SGI_CODIGO", "SGI_DESCRICAO", "SGI_CADTIPPROD", txtTipo.Text, lblDescTipoProd, False)
    
        Call PopGrdEstoque(objCADPRODUTO.IDProduto)
    
        optAlcGalSN(objCADPRODUTO.ALCAGALAO).Value = True
        optNeckInSN(objCADPRODUTO.NECKIN).Value = True
        optDimCortePAD(objCADPRODUTO.DIMPADRAO).Value = True
   
        If Len(Trim(objCADPRODUTO.DESENV)) > 0 Then txtDESENV.Text = Format(objCADPRODUTO.DESENV, "#,##0.00")
        If Len(Trim(objCADPRODUTO.ALTURA)) > 0 Then txtALTURA.Text = Format(objCADPRODUTO.ALTURA, "#,##0.00")
        
        optFilialPed(objCADPRODUTO.FILIALPED).Value = True
        
        Call CarregaImagem
        Call PopGrdProdEntr
        Call PopEstoqueLitografado(objCADPRODUTO.IDProduto)
    
        optQTDCORPPADRAOSN(objCADPRODUTO.QTDCORPSPADRAOSN).Value = True
        optProdNovoSN(objCADPRODUTO.PRODNOVO).Value = True
    
    
        Me.Caption = Me.Caption & " / [ Produto : " & Trim(txtCodigo.Text) & " - " & Trim(txtDescricao.Text) & "]"
        
        
        '' Carrega a Estrutura
        Call Pop_Estrutura
        
    End If
    
    Frame17.Enabled = False

End Sub

Private Sub InitGrdFamMaq()

    With grdFamMaq
    
       .Cols = conColumnsIn_SonFamMaq
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonFamMaq_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonFamMaq_CodFamMaq) = ""
       .ColDataType(conCOL_SonFamMaq_CodFamMaq) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonFamMaq_PesqFam) = ""
       .ColDataType(conCOL_SonFamMaq_PesqFam) = flexDTString
       .ColComboList(conCOL_SonFamMaq_PesqFam) = "..."
       
       .Cell(flexcpData, 0, conCOL_SonFamMaq_DescFam) = ""
       .ColDataType(conCOL_SonFamMaq_DescFam) = flexDTString
       
       .ColWidth(conCOL_SonFamMaq_CodFamMaq) = 1500
       .ColWidth(conCOL_SonFamMaq_PesqFam) = 300
       .ColWidth(conCOL_SonFamMaq_DescFam) = 5000
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
    End With
    
End Sub

Private Function PegaDescrFamMaq(strCodfamMaq As String) As String
    PegaDescrFamMaq = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADFAMMAQUINAS " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & Trim(strCodfamMaq)
    
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC2.EOF Then PegaDescrFamMaq = BREC2!SGI_DESCRI
    BREC2.Close
    
End Function

Private Sub IncRegGridFamMaq()
   
    If objBLBFunc.FcExisteLinhaVazia(grdFamMaq, conCOL_SonFamMaq_CodFamMaq) = False Then Exit Sub
    
    With grdFamMaq
            .AddItem "" & vbTab & _
                     "" & vbTab & _
                     ""
    End With
                            
End Sub

Private Sub PopGrdFamMaq()

     Dim i As Integer
     
     arrFAMMAQ = objCADPRODUTO.FAMMAQ
     If IsArray(arrFAMMAQ) Then
        With grdFamMaq
            For i = 1 To UBound(arrFAMMAQ)
                .AddItem arrFAMMAQ(i) & vbTab & _
                         "" & vbTab & _
                         PegaDescrFamMaq(Trim(Str(arrFAMMAQ(i))))
            Next i
        End With
     End If

End Sub

Public Sub LimpaCamposLabel()
    lblDescFamProd.Caption = ""
    lblDescSubFamProd.Caption = ""
    lblDescEspecieProd.Caption = ""
    lblDescTipoProd.Caption = ""
    lblDescProd.Caption = ""
End Sub


Private Sub PegaDescTabelas(StrCampoPesq As String, StrCampoRetorno As String, strTabela As String, strCODIGO As String, lblLabel As Label, boolMens As Boolean)

    lblLabel.Caption = ""
    
    If Len(Trim(strCODIGO)) = 0 Then Exit Sub
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoRetorno) & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & Trim(strTabela) & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoPesq) & " = " & Trim(strCODIGO)
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then
       lblLabel.Caption = Trim(BREC10(Trim(StrCampoRetorno)))
    Else
       MsgBox "Registro Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC10.Close
    
End Sub

Private Sub PegaDescTabelasSubFamProd(StrFamProd As String, StrCodSubFam As String, lblLabel As Label, boolMens As Boolean)

    lblLabel.Caption = ""
    
    If Len(Trim(StrCodSubFam)) = 0 Then Exit Sub
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SUBGR.SGI_CODIGO     " & vbCrLf
    sSql = sSql & "      ,SUBGR.SGI_DESCRICAO  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_GRUPPRODITEN GRUPO" & vbCrLf
    sSql = sSql & "      ,SGI_CADSUBGRPROD SUBGR" & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      GRUPO.SGI_CODIGO = " & StrFamProd
    sSql = sSql & "  And GRUPO.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "  And SUBGR.SGI_CODIGO = " & StrCodSubFam & vbCrLf
    sSql = sSql & "  And SUBGR.SGI_FILIAL = GRUPO.SGI_FILIAL " & vbCrLf
    sSql = sSql & "  And SUBGR.SGI_CODIGO = GRUPO.SGI_CODESPECIE " & vbCrLf
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then
       lblLabel.Caption = Trim(BREC10!SGI_DESCRICAO)
    Else
       If boolMens = True Then MsgBox "Registro Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC10.Close
    
End Sub

Private Sub PegaDescTabelasEspProd(StrSubFamProd As String, StrCodEspProd As String, lblLabel As Label, boolMens As Boolean)

    lblLabel.Caption = ""
    
    If Len(Trim(StrCodEspProd)) = 0 Then Exit Sub
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       ESPECI.SGI_CODIGO    " & vbCrLf
    sSql = sSql & "      ,ESPECI.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_SUBGRUPRODITEN SUBGRP " & vbCrLf
    sSql = sSql & "      ,SGI_CADESPPROD     ESPECI " & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      SUBGRP.SGI_CODIGO = " & StrSubFamProd & vbCrLf
    sSql = sSql & "  And SUBGRP.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "  And ESPECI.SGI_CODIGO = " & StrCodEspProd & vbCrLf
    sSql = sSql & "  And ESPECI.SGI_FILIAL = SUBGRP.SGI_FILIAL " & vbCrLf
    sSql = sSql & "  And ESPECI.SGI_CODIGO = SUBGRP.SGI_CODESPECIE "
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then
       lblLabel.Caption = Trim(BREC10!SGI_DESCRICAO)
    Else
       If boolMens = True Then MsgBox "Registro Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC10.Close
    
End Sub


Private Sub ConfEstoque()

    With grdESTOQUE
    
       .Cols = conColumnsIn_SonEstoque
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonEstoque_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonEstoque_CodClie) = ""
       .ColDataType(conCOL_SonEstoque_CodClie) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonEstoque_DescClie) = ""
       .ColDataType(conCOL_SonEstoque_DescClie) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonEstoque_Entradas) = ""
       .ColDataType(conCOL_SonEstoque_Entradas) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonEstoque_Saidas) = ""
       .ColDataType(conCOL_SonEstoque_Saidas) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonEstoque_Saldo) = ""
       .ColDataType(conCOL_SonEstoque_Saldo) = flexDTLong
       
       .ColWidth(conCOL_SonEstoque_CodClie) = 1000
       .ColWidth(conCOL_SonEstoque_DescClie) = 4500
       .ColWidth(conCOL_SonEstoque_Entradas) = 1000
       .ColWidth(conCOL_SonEstoque_Saidas) = 1000
       .ColWidth(conCOL_SonEstoque_Saldo) = 1000
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With

End Sub


Private Sub PopGrdEstoque(strID As String)

    Dim lngLinha    As Long
    Dim dblENTRADAS As Double
    Dim dblSAIDAS   As Double
    Dim dblSALDOS   As Double
    Dim i           As Long
    
    
    sSql = ""
    
    With grdESTOQUE
    
        '' Entradas
        sSql = ""
        
        sSql = "Select" & vbCrLf
        sSql = sSql & "      CABEC.SGI_CODCLIENTE" & vbCrLf
        sSql = sSql & "     ,CLIE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "     ,Sum(IT.SGI_QTD) As SGI_SALDO" & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "      SGI_CADITREQENTRMAT IT" & vbCrLf
        sSql = sSql & "     ,SGI_CADREQENTRMAT   CABEC" & vbCrLf
        sSql = sSql & "     ,SGI_CADCLIENTE      CLIE" & vbCrLf
        sSql = sSql & "  Where" & vbCrLf
        sSql = sSql & "      IT.SGI_FILIAL    = " & FILIAL & vbCrLf
        sSql = sSql & "  And IT.SGI_IDPRODUTO = " & strID & vbCrLf
        sSql = sSql & "  And CABEC.SGI_FILIAL = IT.SGI_FILIAL" & vbCrLf
        sSql = sSql & "  And CABEC.SGI_CODIGO = IT.SGI_CODIGO" & vbCrLf
        sSql = sSql & "  And CLIE.SGI_FILIAL  = CABEC.SGI_FILIAL" & vbCrLf
        sSql = sSql & "  And CLIE.SGI_CODIGO  = CABEC.SGI_CODCLIENTE" & vbCrLf
        sSql = sSql & "Group By CABEC.SGI_CODCLIENTE" & vbCrLf
        sSql = sSql & "        ,CLIE.SGI_RAZAOSOC"
    
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        Do While Not BREC.EOF()
            lngLinha = .FindRow(BREC!SGI_CODCLIENTE, , conCOL_SonEstoque_CodClie)
            If lngLinha = -1 Then
               .AddItem BREC!SGI_CODCLIENTE & vbTab & _
                        BREC!SGI_RAZAOSOC & vbTab & _
                        BREC!SGI_SALDO & vbTab & _
                        0 & vbTab & _
                        0
            Else
               .Cell(flexcpText, lngLinha, conCOL_SonEstoque_CodClie) = BREC!SGI_CODCLIENTE
               .Cell(flexcpText, lngLinha, conCOL_SonEstoque_DescClie) = BREC!SGI_RAZAOSOC
               .Cell(flexcpText, lngLinha, conCOL_SonEstoque_Entradas) = BREC!SGI_SALDO
            End If
            BREC.MoveNext
        Loop
        BREC.Close
    
    
        '' Saidas
        sSql = ""
        
        sSql = "Select" & vbCrLf
        sSql = sSql & "      CABEC.SGI_CODCLIENTE" & vbCrLf
        sSql = sSql & "     ,CLIE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "     ,Sum(IT.SGI_QTD) As SGI_SALDO" & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "      SGI_CADITREQSAIMAT IT" & vbCrLf
        sSql = sSql & "     ,SGI_CADREQSAIMAT   CABEC" & vbCrLf
        sSql = sSql & "     ,SGI_CADCLIENTE      CLIE" & vbCrLf
        sSql = sSql & "  Where" & vbCrLf
        sSql = sSql & "      IT.SGI_FILIAL    = " & FILIAL & vbCrLf
        sSql = sSql & "  And IT.SGI_IDPRODUTO = " & strID & vbCrLf
        sSql = sSql & "  And CABEC.SGI_FILIAL = IT.SGI_FILIAL" & vbCrLf
        sSql = sSql & "  And CABEC.SGI_CODIGO = IT.SGI_CODIGO" & vbCrLf
        sSql = sSql & "  And CLIE.SGI_FILIAL  = CABEC.SGI_FILIAL" & vbCrLf
        sSql = sSql & "  And CLIE.SGI_CODIGO  = CABEC.SGI_CODCLIENTE" & vbCrLf
        sSql = sSql & "Group By CABEC.SGI_CODCLIENTE" & vbCrLf
        sSql = sSql & "        ,CLIE.SGI_RAZAOSOC"
    
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        Do While Not BREC.EOF()
            lngLinha = .FindRow(BREC!SGI_CODCLIENTE, , conCOL_SonEstoque_CodClie)
            If lngLinha = -1 Then
               .AddItem BREC!SGI_CODCLIENTE & vbTab & _
                        BREC!SGI_RAZAOSOC & vbTab & _
                        0 & vbTab & _
                        BREC!SGI_SALDO & vbTab & _
                        0
            Else
               .Cell(flexcpText, lngLinha, conCOL_SonEstoque_CodClie) = BREC!SGI_CODCLIENTE
               .Cell(flexcpText, lngLinha, conCOL_SonEstoque_DescClie) = BREC!SGI_RAZAOSOC
               .Cell(flexcpText, lngLinha, conCOL_SonEstoque_Saidas) = BREC!SGI_SALDO
            End If
            BREC.MoveNext
        Loop
        BREC.Close
        
        For i = 1 To (.Rows - 1)
            dblENTRADAS = CDbl(.Cell(flexcpText, i, conCOL_SonEstoque_Entradas))
            dblSAIDAS = CDbl(.Cell(flexcpText, i, conCOL_SonEstoque_Saidas))
            If dblSAIDAS > 0 Then dblSAIDAS = (dblSAIDAS * -1)
            dblSALDOS = (dblENTRADAS + dblSAIDAS)
            .Cell(flexcpText, i, conCOL_SonEstoque_Saldo) = dblSALDOS
        Next i

    End With
End Sub


Private Sub CarregaGridVerniz02()

    Dim i As Integer
    
    If IsArray(arrVERNIZ02) = True Then
       For i = 1 To UBound(arrVERNIZ02)
       
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  From " & vbCrLf
           sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
           sSql = sSql & "   And SGI_IDPRODUTO = " & arrVERNIZ02(i, 1)
           
           BREC.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC.EOF Then
              grdVernizEsm.Cell(flexcpText, 2, conCOL_SonVerniz_Codigo) = BREC!SGI_IDPRODUTO
              grdVernizEsm.Cell(flexcpText, 2, conCOL_SonVerniz_CodTipo) = BREC!SGI_CODIGO
              grdVernizEsm.Cell(flexcpText, 2, conCOL_SonVerniz_DescVerniz) = BREC!SGI_DESCRICAO
              grdVernizEsm.Cell(flexcpText, 2, conCOL_SonVerniz_UnidMed) = arrVERNIZ02(i, 2)
              If Len(Trim(arrVERNIZ02(i, 3))) > 0 Then grdVernizEsm.Cell(flexcpText, 2, conCOL_SonVerniz_Qtde) = Format(arrVERNIZ02(i, 3), "#,###0.000")
           
           End If
           BREC.Close
       
       Next i
    End If

End Sub


Private Sub PegaDescTabelas2(StrCampoPesq As String, StrCampoRetorno As String, strTabela As String, strCODIGO As String, txtText As TextBox, boolMens As Boolean)

    txtText.Text = ""
    
    If Len(Trim(strCODIGO)) = 0 Then Exit Sub
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoRetorno) & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & Trim(strTabela) & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoPesq) & " = " & Trim(strCODIGO)
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then
       txtText.Text = Trim(BREC10(Trim(StrCampoRetorno)))
    Else
       MsgBox "Registro Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC10.Close
    
End Sub


Private Sub ConfProdEntr()

    With grdPRODENTR
    
       .Cols = conColumnsIn_SonProdEntr
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonProdEntr_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonProdEntr_IDProduto) = ""
       .ColDataType(conCOL_SonProdEntr_IDProduto) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonProdEntr_CodProd) = ""
       .ColDataType(conCOL_SonProdEntr_CodProd) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonProdEntr_PesqProd) = ""
       .ColDataType(conCOL_SonProdEntr_PesqProd) = flexDTString
       .ColComboList(conCOL_SonProdEntr_PesqProd) = "..."
       
       .Cell(flexcpData, 0, conCOL_SonProdEntr_DescProd) = ""
       .ColDataType(conCOL_SonProdEntr_DescProd) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonProdEntr_UniMed) = ""
       .ColDataType(conCOL_SonProdEntr_UniMed) = flexDTString
       .ColComboList(conCOL_SonProdEntr_UniMed) = objCADPRODUTO.PreenchComboUnidMedGrid
       
       .Cell(flexcpData, 0, conCOL_SonProdEntr_Qtde) = ""
       .ColDataType(conCOL_SonProdEntr_Qtde) = flexDTCurrency
       
       .ColWidth(conCOL_SonProdEntr_IDProduto) = 800
       .ColWidth(conCOL_SonProdEntr_CodProd) = 1500
       .ColWidth(conCOL_SonProdEntr_PesqProd) = 300
       .ColWidth(conCOL_SonProdEntr_DescProd) = 5000
       .ColWidth(conCOL_SonProdEntr_UniMed) = 1500
       .ColWidth(conCOL_SonProdEntr_Qtde) = 1500
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With

End Sub

Private Sub LimpaGrdProdEntr(lngROW As Long)
    With grdPRODENTR
        .Cell(flexcpText, lngROW, conCOL_SonProdEntr_IDProduto) = Empty
        .Cell(flexcpText, lngROW, conCOL_SonProdEntr_CodProd) = Empty
        .Cell(flexcpText, lngROW, conCOL_SonProdEntr_DescProd) = Empty
    End With
End Sub

Private Sub IncRegGridProdEntr()
   
    If objBLBFunc.FcExisteLinhaVazia(grdPRODENTR, conCOL_SonProdEntr_CodProd) = False Then Exit Sub
    
    grdPRODENTR.AddItem "" & vbTab & _
                        "" & vbTab & _
                        "" & vbTab & _
                        "" & vbTab & _
                        "" & vbTab & _
                        ""
                        
                            
End Sub



Private Sub PopGrdProdEntr()

    Dim i As Integer
    arrPRODENTR = objCADPRODUTO.PRODENTR

    If IsArray(arrPRODENTR) Then
        With grdPRODENTR
            For i = 1 To UBound(arrPRODENTR)
                .AddItem arrPRODENTR(i, 1) & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         arrPRODENTR(i, 2) & vbTab & _
                         ""
            
                If Len(Trim(arrPRODENTR(i, 3))) > 0 Then .Cell(flexcpText, (.Rows - 1), conCOL_SonProdEntr_Qtde) = Format(arrPRODENTR(i, 3), "#,####0.0000")
            
                sSql = ""
                
                sSql = "Select " & vbCrLf
                sSql = sSql & "       SGI_CODIGO" & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_CADPRODUTO" & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
                sSql = sSql & "   And SGI_IDPRODUTO = " & arrPRODENTR(i, 1)
                
                BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
                If Not BREC10.EOF() Then
                   .Cell(flexcpText, (.Rows - 1), conCOL_SonProdEntr_CodProd) = BREC10!SGI_CODIGO
                   .Cell(flexcpText, (.Rows - 1), conCOL_SonProdEntr_DescProd) = PegaDescrProduto(BREC10!SGI_CODIGO)
                End If
                BREC10.Close
                
                
            
            
            Next i
        End With
    End If

End Sub

Private Sub PintaLabel(intTipo As Integer, intEstilo As Integer)

    If intTipo = 1 And intEstilo = 0 Then
        '' Vermelho
        Label1(9).ForeColor = &HFF&
        Label1(10).ForeColor = &HFF&
        Label1(11).ForeColor = &HFF&
        Label1(12).ForeColor = &HFF&
        Label1(20).ForeColor = &HFF&
        optNatSimNao(0).ForeColor = &HFF&
        optNatSimNao(1).ForeColor = &HFF&
        
        '' Azul
        Label1(0).ForeColor = &HC00000
    Else
        '' Azul
        Label1(9).ForeColor = &HC00000
        Label1(10).ForeColor = &HC00000
        Label1(11).ForeColor = &HC00000
        Label1(12).ForeColor = &HC00000
        Label1(20).ForeColor = &HC00000
        optNatSimNao(0).ForeColor = &HC00000
        optNatSimNao(1).ForeColor = &HC00000
        
        '' Vermelho
        Label1(0).ForeColor = &HFF&
    
    End If
    
    
End Sub

Private Sub Refaz_IndCores()

    Dim i As Integer

    With grdCores
        For i = 1 To (.Rows - 1)
            .Cell(flexcpText, i, conCOL_SonCores_Ordem) = i
        Next i
    End With

End Sub


Private Sub CarregaVernizExterno()

    Dim i As Integer
    
    If IsArray(arrVERNIZEXT) = True Then
       For i = 1 To UBound(arrVERNIZEXT)
       
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  From " & vbCrLf
           sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
           sSql = sSql & "   And SGI_IDPRODUTO = " & arrVERNIZEXT(i, 1)
           
           BREC.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC.EOF Then
              grdVernizEsm.Cell(flexcpText, 5, conCOL_SonVerniz_Codigo) = BREC!SGI_IDPRODUTO
              grdVernizEsm.Cell(flexcpText, 5, conCOL_SonVerniz_CodTipo) = BREC!SGI_CODIGO
              grdVernizEsm.Cell(flexcpText, 5, conCOL_SonVerniz_DescVerniz) = BREC!SGI_DESCRICAO
              grdVernizEsm.Cell(flexcpText, 5, conCOL_SonVerniz_UnidMed) = arrVERNIZEXT(i, 2)
              If Len(Trim(arrVERNIZEXT(i, 3))) > 0 Then grdVernizEsm.Cell(flexcpText, 5, conCOL_SonVerniz_Qtde) = Format(arrVERNIZEXT(i, 3), "#,####0.0000")
           
           End If
           BREC.Close
       
       Next i
    End If

End Sub

Private Sub PopEstoqueLitografado(strID As String)

    Dim lngLinha        As Long
    Dim dblENTRADAS     As Double
    Dim dblSAIDAS       As Double
    Dim dblSALDOS       As Double
    
    Dim dblENTRADASFol  As Double
    Dim dblSAIDASFol    As Double
    Dim dblSALDOSFol    As Double
    
    Dim i               As Long
    
    With grdEstLit
    
        '' ===============================
        '' Estoque de Litografia
        '' Entradas de Litografia ( Na Origrm )
        
        sSql = ""
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "      CABEC.SGI_CODCLIE" & vbCrLf
        sSql = sSql & "     ,CLIE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "     ,Sum(ITEN.SGI_CONFPESO)         As SGI_SALDOPESO" & vbCrLf
        sSql = sSql & "     ,Sum(ITEN.SGI_QTDEFOLHASREC)    As SGI_SALDOFOLHAS" & vbCrLf
        
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADRECROTLIT_IT ITEN" & vbCrLf
        sSql = sSql & "      ,SGI_CADRECROTLIT    CABEC" & vbCrLf
        sSql = sSql & "      ,SGI_CADCLIENTE      CLIE" & vbCrLf
        
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       ITEN.SGI_FILIAL    = " & FILIAL & vbCrLf
        sSql = sSql & "   And ITEN.SGI_IDPRODUTO = " & strID & vbCrLf
        sSql = sSql & "   And ITEN.SGI_STATUS    = 'REC'" & vbCrLf       '' Entrada
        
        sSql = sSql & "   And CABEC.SGI_FILIAL   = ITEN.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And CABEC.SGI_CODIGO   = ITEN.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And CLIE.SGI_FILIAL    = CABEC.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And CLIE.SGI_CODIGO    = CABEC.SGI_CODCLIE" & vbCrLf
        sSql = sSql & "Group By CABEC.SGI_CODCLIE" & vbCrLf
        sSql = sSql & "        ,CLIE.SGI_RAZAOSOC"
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        Do While Not BREC.EOF()
            lngLinha = .FindRow(BREC!SGI_CODCLIE, , conCOL_SonEstoqueLit_CodClie)
            If lngLinha = -1 Then
               .AddItem BREC!SGI_CODCLIE & vbTab & _
                        BREC!SGI_RAZAOSOC & " ( ORIGEM )" & vbTab & _
                        Format(BREC!SGI_SALDOPESO, "#,####0.0000") & vbTab & _
                        0 & vbTab & _
                        0 & vbTab & _
                        Format(BREC!SGI_SALDOFOLHAS, "#0") & vbTab & _
                        0 & vbTab & _
                        0
            Else
               .Cell(flexcpText, lngLinha, conCOL_SonEstoqueLit_CodClie) = BREC!SGI_CODCLIE
               .Cell(flexcpText, lngLinha, conCOL_SonEstoqueLit_DescClie) = BREC!SGI_RAZAOSOC & " ( ORIGEM )"
               .Cell(flexcpText, lngLinha, conCOL_SonEstoqueLit_EntradasKG) = Format(BREC!SGI_SALDOPESO, "#,####0.0000")
               .Cell(flexcpText, lngLinha, conCOL_SonEstoqueLit_EntradasFolhas) = Format(BREC!SGI_SALDOFOLHAS, "#0")
            End If
            BREC.MoveNext
        Loop
        BREC.Close
        '' ===============================
        
        
        '' ===============================
        '' Estoque de Litografia
        '' Entradas de Litografia ( No Destino )
        
        sSql = ""
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "      CABEC.SGI_CODCLIEDEST" & vbCrLf
        sSql = sSql & "     ,CLIE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "     ,Sum(ITEN.SGI_CONFPESO)         As SGI_SALDOPESO" & vbCrLf
        sSql = sSql & "     ,Sum(ITEN.SGI_QTDEFOLHASREC)    As SGI_SALDOFOLHAS" & vbCrLf
        
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADRECROTLIT_IT ITEN" & vbCrLf
        sSql = sSql & "      ,SGI_CADRECROTLIT    CABEC" & vbCrLf
        sSql = sSql & "      ,SGI_CADCLIENTE      CLIE" & vbCrLf
        
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       ITEN.SGI_FILIAL    = " & FILIAL & vbCrLf
        sSql = sSql & "   And ITEN.SGI_IDPRODUTO = " & strID & vbCrLf
        sSql = sSql & "   And ITEN.SGI_STATUS    = 'REC'" & vbCrLf       '' Entrada no Destino
        
        sSql = sSql & "   And CABEC.SGI_FILIAL   = ITEN.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And CABEC.SGI_CODIGO   = ITEN.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And CLIE.SGI_FILIAL    = CABEC.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And CLIE.SGI_CODIGO    = CABEC.SGI_CODCLIEDEST" & vbCrLf
        sSql = sSql & "Group By CABEC.SGI_CODCLIEDEST" & vbCrLf
        sSql = sSql & "        ,CLIE.SGI_RAZAOSOC"
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        Do While Not BREC.EOF()
            lngLinha = .FindRow(BREC!SGI_CODCLIEDEST, , conCOL_SonEstoqueLit_CodClie)
            If lngLinha = -1 Then
               .AddItem BREC!SGI_CODCLIEDEST & vbTab & _
                        BREC!SGI_RAZAOSOC & " ( DESTINO )" & vbTab & _
                        Format(BREC!SGI_SALDOPESO, "#,####0.0000") & vbTab & _
                        0 & vbTab & _
                        0 & vbTab & _
                        Format(BREC!SGI_SALDOFOLHAS, "#0") & vbTab & _
                        0 & vbTab & _
                        0
            Else
               .Cell(flexcpText, lngLinha, conCOL_SonEstoqueLit_CodClie) = BREC!SGI_CODCLIEDEST
               .Cell(flexcpText, lngLinha, conCOL_SonEstoqueLit_DescClie) = BREC!SGI_RAZAOSOC & " ( DESTINO )"
               .Cell(flexcpText, lngLinha, conCOL_SonEstoqueLit_EntradasKG) = Format(BREC!SGI_SALDOPESO, "#,####0.0000")
               .Cell(flexcpText, lngLinha, conCOL_SonEstoqueLit_EntradasFolhas) = Format(BREC!SGI_SALDOFOLHAS, "#0")
            End If
            BREC.MoveNext
        Loop
        BREC.Close
        '' ===============================
        
        
        '' ===============================
        '' Estoque de Litografia
        '' Saidas de Litografia
        '' Saida no Cliente ( Origem )
        
        sSql = ""
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "      CABEC.SGI_CODCLIE" & vbCrLf
        sSql = sSql & "     ,CLIE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "     ,Sum(ITEN.SGI_PESO)       As SGI_SALDOPESO" & vbCrLf
        sSql = sSql & "     ,Sum(ITEN.SGI_QTDEFOLHAS) As SGI_SALDOFOLHAS" & vbCrLf
        
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADENTROTLIT_IT ITEN" & vbCrLf
        sSql = sSql & "      ,SGI_CADENTROTLIT    CABEC" & vbCrLf
        sSql = sSql & "      ,SGI_CADCLIENTE      CLIE" & vbCrLf
        
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       ITEN.SGI_FILIAL    = " & FILIAL & vbCrLf
        sSql = sSql & "   And ITEN.SGI_IDPRODUTO = " & strID & vbCrLf
        sSql = sSql & "   And ITEN.SGI_STATUS    = 'ENV'" & vbCrLf       '' Saida
        
        sSql = sSql & "   And CABEC.SGI_FILIAL   = ITEN.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And CABEC.SGI_CODIGO   = ITEN.SGI_CODIGO" & vbCrLf
        
        sSql = sSql & "   And CLIE.SGI_FILIAL    = CABEC.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And CLIE.SGI_CODIGO    = CABEC.SGI_CODCLIE" & vbCrLf
        sSql = sSql & "Group By CABEC.SGI_CODCLIE" & vbCrLf
        sSql = sSql & "        ,CLIE.SGI_RAZAOSOC"
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        Do While Not BREC.EOF()
            lngLinha = .FindRow(BREC!SGI_CODCLIE, , conCOL_SonEstoqueLit_CodClie)
            If lngLinha = -1 Then
               .AddItem BREC!SGI_CODCLIE & vbTab & _
                        BREC!SGI_RAZAOSOC & " ( ORIGEM )" & vbTab & _
                        0 & vbTab & _
                        Format(BREC!SGI_SALDOPESO, "#,####0.0000") & vbTab & _
                        0 & vbTab & _
                        0 & vbTab & _
                        Format(BREC!SGI_SALDOFOLHAS, "#0") & vbTab & _
                        0
            Else
               .Cell(flexcpText, lngLinha, conCOL_SonEstoqueLit_CodClie) = BREC!SGI_CODCLIE
               .Cell(flexcpText, lngLinha, conCOL_SonEstoqueLit_DescClie) = BREC!SGI_RAZAOSOC & " ( ORIGEM )"
               .Cell(flexcpText, lngLinha, conCOL_SonEstoqueLit_SaidasKG) = Format(BREC!SGI_SALDOPESO, "#,####0.0000")
               .Cell(flexcpText, lngLinha, conCOL_SonEstoqueLit_SaidasFolhas) = Format(BREC!SGI_SALDOFOLHAS, "#0")
            End If
            BREC.MoveNext
        Loop
        BREC.Close
        '' ===============================
        
        
        For i = 1 To (.Rows - 1)
            dblENTRADAS = CDbl(.Cell(flexcpText, i, conCOL_SonEstoqueLit_EntradasKG))
            dblSAIDAS = CDbl(.Cell(flexcpText, i, conCOL_SonEstoqueLit_SaidasKG))
            
            If dblSAIDAS > 0 Then dblSAIDAS = (dblSAIDAS * -1)
            dblSALDOS = (dblENTRADAS + dblSAIDAS)
            .Cell(flexcpText, i, conCOL_SonEstoqueLit_SaldoKG) = Format(dblSALDOS, "#,####0.0000")
            
            '' ===================================
            dblENTRADASFol = CDbl(.Cell(flexcpText, i, conCOL_SonEstoqueLit_EntradasFolhas))
            dblSAIDASFol = CDbl(.Cell(flexcpText, i, conCOL_SonEstoqueLit_SaidasFolhas))
            
            If dblSAIDASFol > 0 Then dblSAIDASFol = (dblSAIDASFol * -1)
            dblSALDOSFol = (dblENTRADASFol + dblSAIDASFol)
            .Cell(flexcpText, i, conCOL_SonEstoqueLit_SaldoFolhas) = Format(dblSALDOSFol, "#0")
            
            
            '' Pinta a Celula
            .Cell(flexcpBackColor, i, conCOL_SonEstoqueLit_EntradasKG) = &HFF00&
            .Cell(flexcpBackColor, i, conCOL_SonEstoqueLit_EntradasFolhas) = &HFF00&

            .Cell(flexcpBackColor, i, conCOL_SonEstoqueLit_SaidasKG) = &HFF&
            .Cell(flexcpBackColor, i, conCOL_SonEstoqueLit_SaidasFolhas) = &HFF&

        Next i

    End With


End Sub

Private Sub ConfEstoqueLit()

    With grdEstLit
    
       .Cols = conColumnsIn_SonEstoqueLit
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonEstoqueLit_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonEstoqueLit_CodClie) = ""
       .ColDataType(conCOL_SonEstoqueLit_CodClie) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonEstoqueLit_DescClie) = ""
       .ColDataType(conCOL_SonEstoqueLit_DescClie) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonEstoqueLit_EntradasKG) = ""
       .ColDataType(conCOL_SonEstoqueLit_EntradasKG) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonEstoqueLit_SaidasKG) = ""
       .ColDataType(conCOL_SonEstoqueLit_SaidasKG) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonEstoqueLit_SaldoKG) = ""
       .ColDataType(conCOL_SonEstoqueLit_SaldoKG) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonEstoqueLit_EntradasFolhas) = ""
       .ColDataType(conCOL_SonEstoqueLit_EntradasFolhas) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_SonEstoqueLit_SaidasFolhas) = ""
       .ColDataType(conCOL_SonEstoqueLit_SaidasFolhas) = flexDTCurrency
       
       .ColWidth(conCOL_SonEstoqueLit_CodClie) = 900
       .ColWidth(conCOL_SonEstoqueLit_DescClie) = 5500
       .ColWidth(conCOL_SonEstoqueLit_EntradasKG) = 1000
       .ColWidth(conCOL_SonEstoqueLit_SaidasKG) = 1000
       .ColWidth(conCOL_SonEstoqueLit_SaldoKG) = 1000
       .ColWidth(conCOL_SonEstoqueLit_EntradasFolhas) = 1600
       .ColWidth(conCOL_SonEstoqueLit_SaidasFolhas) = 1600
       .ColWidth(conCOL_SonEstoqueLit_SaldoFolhas) = 1600
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightAlways
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With

End Sub


Private Sub ConfComboFech()

    ''======================================================
    ''cboFechTampaFuro.Clear
    
    ''Dim I               As Integer
    ''Dim arrFECTP()      As String
    ''ReDim arrFECTP(1 To 25) As String
    
    ''arrFECTP(1) = "Ø24"
    ''arrFECTP(2) = "Ø25"
    ''arrFECTP(3) = "Ø42"
    ''arrFECTP(4) = "Ø45"
    ''arrFECTP(5) = "Ø57"
    ''arrFECTP(6) = "Ø80"
    ''arrFECTP(7) = "Ø130"
    ''arrFECTP(8) = "Ø170"
    ''arrFECTP(9) = "Ø110"
    ''arrFECTP(10) = "Ø170 c/b Ø25"
    ''arrFECTP(11) = "Ø170 c/v Ø57"
    ''arrFECTP(12) = "TP"
    ''arrFECTP(13) = "TP2"
    ''arrFECTP(14) = "TP4"
    ''arrFECTP(15) = "FA"
    ''arrFECTP(16) = "A RECRAVAR"
    ''arrFECTP(17) = "FA - C/Visor"
    ''arrFECTP(18) = "COFRE"
    ''arrFECTP(19) = "Porta Canetas"
    ''arrFECTP(20) = "Ø32 Bico Ret."
    ''arrFECTP(21) = "Repuxo"
    ''arrFECTP(22) = "FA c/b Ø25"
    ''arrFECTP(23) = "Ø24 - Bico de Pato"
    ''arrFECTP(24) = "Ø24 - REU"
    ''arrFECTP(25) = "F/T"
    
    ''With cboFechTampaFuro
    ''    For I = 1 To UBound(arrFECTP)
    ''        .AddItem arrFECTP(I)
    ''        .ItemData(.NewIndex) = I
    ''    Next I
    ''End With
    '' ===================================================

    cboFechTampaFuro.Clear
    If Len(Trim(txtLinProd.Text)) = 0 Then Exit Sub

    With cboFechTampaFuro
    
        sSql = ""
        
        sSql = sSql & "Select " & vbCrLf
        sSql = sSql & "       FECH.SGI_COD" & vbCrLf
        sSql = sSql & "     , FEC.SGI_DESCRI" & vbCrLf
        
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADLINHAPRODUTO          LPRO" & vbCrLf
        sSql = sSql & "     , SGI_CADLINHAPRODUTO_FECHTPFR FECH" & vbCrLf
        sSql = sSql & "     , SGI_CADFECHAM                FEC" & vbCrLf
                
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       LPRO.SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And LPRO.SGI_CODLIN = " & txtLinProd.Text & vbCrLf
        sSql = sSql & "   And FECH.SGI_FILIAL = LPRO.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And FECH.SGI_CODIGO = LPRO.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And FECH.SGI_FILIAL = FEC.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And FECH.SGI_COD    = FEC.SGI_CODIGO" & vbCrLf
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        Do While Not BREC.EOF()
            .AddItem Trim(BREC!SGI_DESCRI)
            .ItemData(.NewIndex) = BREC!SGI_COD
            BREC.MoveNext
        Loop
        BREC.Close
        
        If .ListCount = 1 Then .ListIndex = 0
    
    End With

End Sub

Private Function Constroi_Codigo_Comprado(strCODGRUPO As String, strSUBGRUP As String)

    Constroi_Codigo_Comprado = ""
    
    If cTipOper = "I" Or cTipOper = "A" Then
        
        objCADPRODUTO.SEGUENCIA = "Null"
        
        Dim lngSEQUENCIAL As Long
    
        If Len(Trim(strCODGRUPO)) > 0 Then Constroi_Codigo_Comprado = Constroi_Codigo_Comprado & Format(strCODGRUPO, "##00")
        If Len(Trim(strSUBGRUP)) > 0 Then Constroi_Codigo_Comprado = Constroi_Codigo_Comprado & "." & Format(strSUBGRUP, "##00")
       
        If Len(Trim(strCODGRUPO)) > 0 And Len(Trim(strSUBGRUP)) > 0 Then
            
            sSql = ""
            
            sSql = "Select" & vbCrLf
            sSql = sSql & "       Count(*) as SGI_QTDEREGS" & vbCrLf
            sSql = sSql & "  From" & vbCrLf
            sSql = sSql & "       SGI_CADPRODUTO" & vbCrLf
            sSql = sSql & " Where" & vbCrLf
            sSql = sSql & "       SGI_FILIAL         = " & FILIAL & vbCrLf
            sSql = sSql & "   And SGI_PRODUTOTIPO    = 0" & vbCrLf
            sSql = sSql & "   And (SGI_PRODUTOESTILO = 0 or SGI_PRODUTOESTILO = 4)" & vbCrLf  '' Acabado
            sSql = sSql & "   And SGI_CODGPROD       = " & strCODGRUPO & vbCrLf
            sSql = sSql & "   And SGI_CODSUBGPROD    = " & strSUBGRUP
        
            BREC.Open sSql, adoBanco_Dados, adOpenDynamic
            If Not BREC.EOF() Then
                lngSEQUENCIAL = 1
                If Not IsNull(BREC!SGI_QTDEREGS) Then lngSEQUENCIAL = (BREC!SGI_QTDEREGS + 1)
            End If
            BREC.Close
            
            objCADPRODUTO.SEGUENCIA = Format(lngSEQUENCIAL, "####0000")
            Constroi_Codigo_Comprado = Constroi_Codigo_Comprado & "." & Format(lngSEQUENCIAL, "####0000")
        
        End If
        

    End If
    
End Function


Private Sub Limpa_Lista_Lata()
      treListaMat.Nodes.Clear
End Sub

Private Sub Inseri_Item(lngINDLIST As Long)

    Dim i            As Long

    If lngINDLIST > 0 Then
        
        For i = 1 To treListaMat.Nodes.Count
            If Len(Trim(treListaMat.Nodes.Item(i).Text)) = 0 Then
                MsgBox "ATENÇÂO" & vbCrLf & "Primeiro informe o Insumo para poder adicionar um novo insumo !!!", vbOKOnly + vbExclamation, "Aviso"
                Exit Sub
            End If
        Next i
        
        If ConsisteArvore = False Then Exit Sub
        
        treListaMat.Nodes.Add lngINDLIST, 4, , "Novo"
        treListaMat.Nodes.Item(lngINDLIST).Expanded = True
    
    Else
        MsgBox "Selecione o Cabeçalho !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    
End Sub

Private Sub Cria_Item_Pai()
    Dim strKEY As String
    
    If treListaMat.Nodes.Count = 0 Then
        
        treListaMat.Nodes.Add , , , txtDescricao.Text
        strKEY = Trim(Str(treListaMat.Nodes.Count))
        
        ReDim arrPROVARV(1 To 1) As PRODARVPROD
        arrPROVARV(CLng(strKEY)).intAction2Do = dacEnumUpdateAction_Insert
        arrPROVARV(CLng(strKEY)).lngIDPai = -1
        arrPROVARV(CLng(strKEY)).lngID = CLng(strKEY)
        
        If cTipOper = "A" Or cTipOper = "C" Then
            arrPROVARV(CLng(strKEY)).lngProdutoID = iCodigo
            arrPROVARV(CLng(strKEY)).strPRODUTO = txtDescricao.Text
        Else
            arrPROVARV(CLng(strKEY)).lngProdutoID = -1
        End If
        
        lngINDLIST = treListaMat.Nodes.Count
        treListaMat.Nodes.Item(lngINDLIST).Selected = True
        treListaMat.Nodes.Item(lngINDLIST).Text = arrPROVARV(CLng(strKEY)).strPRODUTO
        
        
    End If

End Sub

Private Sub AbilDesDadosListMat()
    fraDadosList.Enabled = True
End Sub

Private Sub AbilDesCamposListMat(boolChave As Boolean)
    fraDadosList.Enabled = boolChave
    Call LimpaCamposLista
End Sub

Private Sub LimpaCamposLista()
    txtCODPROD.Text = ""
    lblDescProd.Caption = ""
    cboUniIns.ListIndex = -1
    txtQtdeIns.Text = ""
End Sub


Private Sub PegaDadosDoArray()

    Dim i           As Integer

    If lngINDLIST > 1 Then
        If treListaMat.Nodes.Item(lngINDLIST).Text <> "Novo" Then
            If Len(Trim(arrPROVARV(lngINDLIST).strPRODUTO)) = 0 Then Exit Sub
            txtCODPROD.Text = arrPROVARV(lngINDLIST).strPRODUTO
            lblDescProd.Caption = PegaDescrProduto(txtCODPROD.Text)
            
            If Len(Trim(arrPROVARV(lngINDLIST).strQTDCONS)) > 0 Then txtQtdeIns.Text = arrPROVARV(lngINDLIST).strQTDCONS
        
            For i = 0 To (cboUniIns.ListCount - 1)
                If cboUniIns.ItemData(i) = arrPROVARV(lngINDLIST).lngCodUniMed Then
                   cboUniIns.ListIndex = i
                   Exit For
                End If
            Next i
        
            treListaMat.Nodes.Item(lngINDLIST).Text = txtCODPROD.Text & " - " & lblDescProd.Caption & " - " & cboUniIns.Text & " " & txtQtdeIns.Text
        End If
    End If

End Sub


Private Function PegaDesProdID(lngIDProduto As Long) As String
    PegaDesProdID = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_IDPRODUTO = " & lngIDProduto
    
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC2.EOF Then PegaDesProdID = Trim(BREC2!SGI_CODIGO) & " - " & Trim(BREC2!SGI_DESCRICAO)
    BREC2.Close
    
End Function


Private Function PegaCodProdID(lngIDProduto As Long) As String
    PegaCodProdID = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_IDPRODUTO = " & lngIDProduto
    
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC2.EOF Then PegaCodProdID = Trim(BREC2!SGI_CODIGO)
    BREC2.Close
    
End Function


Private Sub Pop_Estrutura()

        Dim i   As Long
        Dim j   As Long
        Dim k   As Long

        '' 07/08/2017
        '' Cria a Estrutura
        If arrPROVARV(1).lngProdutoID = -1 Then
            Call Cria_Item_Pai
        Else
        
            Dim lngIDPai As Long
        
            For i = 1 To UBound(arrPROVARV)
            
                If arrPROVARV(i).lngCODPAI = 0 Then
                    
                    treListaMat.Nodes.Add , , , txtDescricao.Text
                    treListaMat.Nodes.Item(i).Expanded = True
                    arrPROVARV(i).lngIDPai = -1
                    arrPROVARV(i).lngID = i
                    lngIDPai = treListaMat.Nodes.Count
                    
                Else
        
                    
                    '' Pegando o ID do Pai
                    For j = 1 To UBound(arrPROVARV)
                        If arrPROVARV(i).lngCODPAI = arrPROVARV(j).lngCODIGO Then
                            lngIDPai = j
                            arrPROVARV(i).lngIDPai = lngIDPai
                        End If
                    Next j
                    
                    '' Inserino Item no tree view
                    treListaMat.Nodes.Add lngIDPai, 4
                    treListaMat.Nodes.Item(lngIDPai).Expanded = True
    
                    arrPROVARV(i).intAction2Do = dacEnumUpdateAction_Ignore
                    arrPROVARV(i).lngID = i
                    
                    If arrPROVARV(i).lngCodUniMed > 0 Then
                        For k = 0 To (cboUniIns.ListCount - 1)
                            If cboUniIns.ItemData(k) = arrPROVARV(i).lngCodUniMed Then cboUniIns.ListIndex = k
                        Next k
                    End If
                    If Len(Trim(arrPROVARV(i).strQTDCONS)) > 0 Then txtQtdeIns.Text = arrPROVARV(i).strQTDCONS
                    
                    treListaMat.Nodes.Item(i).Text = PegaDesProdID(arrPROVARV(i).lngProdutoID) & " - " & Trim(cboUniIns.Text) & " " & Trim(arrPROVARV(i).strQTDCONS)
                    arrPROVARV(i).strPRODUTO = PegaCodProdID(arrPROVARV(i).lngProdutoID)
                
                End If
            
            Next i
        
        End If
        Call LimpaCamposLista

End Sub

Private Sub CriaArray()

    Dim lngProdIDPai    As Long
    Dim strDADOS        As String

    If treListaMat.Nodes.Item(lngINDLIST).Text = "Novo" Then
        
        ReDim Preserve arrPROVARV(1 To treListaMat.Nodes.Count) As PRODARVPROD
        
        lngProdIDPai = arrPROVARV(treListaMat.Nodes.Item(lngINDLIST).Parent.Index).lngProdutoID
        
        arrPROVARV(treListaMat.Nodes.Count).lngIDPai = treListaMat.Nodes.Item(lngINDLIST).Parent.Index
        arrPROVARV(treListaMat.Nodes.Count).lngID = treListaMat.Nodes.Count
        arrPROVARV(treListaMat.Nodes.Count).strPRODUTO = txtCODPROD.Text
        arrPROVARV(treListaMat.Nodes.Count).lngProdutoID = CLng(PegaIDProduto(txtCODPROD.Text))
        arrPROVARV(treListaMat.Nodes.Count).lngProdutoIDPai = lngProdIDPai
        
        strDADOS = ""
        If cboUniIns.ListIndex > -1 Then
            arrPROVARV(lngINDLIST).lngCodUniMed = cboUniIns.ItemData(cboUniIns.ListIndex)
            arrPROVARV(lngINDLIST).strUNIDADE = cboUniIns.Text
            strDADOS = " - " & arrPROVARV(lngINDLIST).strUNIDADE
        End If
        If Len(Trim(txtQtdeIns.Text)) > 0 Then
            arrPROVARV(lngINDLIST).strQTDCONS = txtQtdeIns.Text
            strDADOS = strDADOS & " " & arrPROVARV(lngINDLIST).strQTDCONS
        End If
        
        treListaMat.Nodes.Item(lngINDLIST).Text = Trim(txtCODPROD.Text) & " - " & Trim(lblDescProd.Caption) & strDADOS
        
    Else
    
        lngProdIDPai = arrPROVARV(treListaMat.Nodes.Item(lngINDLIST).Parent.Index).lngProdutoID
        
        arrPROVARV(treListaMat.Nodes.Count).lngIDPai = treListaMat.Nodes.Item(lngINDLIST).Parent.Index
        arrPROVARV(treListaMat.Nodes.Count).lngID = lngINDLIST
        arrPROVARV(treListaMat.Nodes.Count).strPRODUTO = txtCODPROD.Text
        arrPROVARV(treListaMat.Nodes.Count).lngProdutoID = CLng(PegaIDProduto(txtCODPROD.Text))
        arrPROVARV(treListaMat.Nodes.Count).lngProdutoIDPai = lngProdIDPai
        
        strDADOS = ""
        If cboUniIns.ListIndex > -1 Then
            arrPROVARV(lngINDLIST).lngCodUniMed = cboUniIns.ItemData(cboUniIns.ListIndex)
            arrPROVARV(lngINDLIST).strUNIDADE = cboUniIns.Text
            strDADOS = " - " & arrPROVARV(lngINDLIST).strUNIDADE
        End If
        If Len(Trim(txtQtdeIns.Text)) > 0 Then
            arrPROVARV(lngINDLIST).strQTDCONS = txtQtdeIns.Text
            strDADOS = strDADOS & " " & arrPROVARV(lngINDLIST).strQTDCONS
        End If
        
        treListaMat.Nodes.Item(lngINDLIST).Text = Trim(txtCODPROD.Text) & " - " & Trim(lblDescProd.Caption) & strDADOS
    
    End If
End Sub

Private Function ConsisteArvore() As Boolean
    ConsisteArvore = False
    
    Dim i As Long
    
    For i = 1 To treListaMat.Nodes.Count
    
        If treListaMat.Nodes(i).Text = "Novo" Then
            MsgBox "ATENÇÂO - Informe primeiro o produto para inserir novo Item !!!", vbOKOnly + vbExclamation, "Aviso"
            Exit Function
        End If
    Next i

    ConsisteArvore = True
End Function

Private Sub Exclui_Item_Arvore()

    Dim i               As Long
    Dim lngITEMARRAY    As Long
    Dim lngIDPai        As Long
    Dim arrNOVO()       As PRODARVPROD
    
    '' =====================
    '' Apagando do Array
    arrPROVARV(lngINDLIST).intAction2Do = dacEnumUpdateAction_delete
    
    '' =====================
    '' Verificando se Existe Filhos
    If treListaMat.Nodes.Item(lngINDLIST).Children > 0 Then
        '' Apagando os Filhos (Todos)
        lngIDPai = treListaMat.Nodes.Item(lngINDLIST).Index
        For i = 1 To UBound(arrPROVARV)
            If arrPROVARV(i).lngIDPai = lngIDPai Then arrPROVARV(i).intAction2Do = dacEnumUpdateAction_delete
        Next i
    End If
    '' =====================
    
    
    '' =====================
    '' Recriando o Array
    lngITEMARRAY = 1
    For i = 1 To (treListaMat.Nodes.Count)
    
        ReDim Preserve arrNOVO(1 To lngITEMARRAY) As PRODARVPROD
        
        If i > 1 Then
            If arrPROVARV(i).intAction2Do <> dacEnumUpdateAction_delete Then
                
                lngIDPai = treListaMat.Nodes.Item(i).Parent.Index
                arrNOVO(lngITEMARRAY).intAction2Do = arrPROVARV(i).intAction2Do
                arrNOVO(lngITEMARRAY).lngCODIGO = arrPROVARV(i).lngCODIGO
                arrNOVO(lngITEMARRAY).lngCODPAI = arrPROVARV(i).lngCODPAI
                arrNOVO(lngITEMARRAY).lngCodUniMed = arrPROVARV(i).lngCodUniMed
                
                arrNOVO(lngITEMARRAY).lngID = lngITEMARRAY
                arrNOVO(lngITEMARRAY).lngIDPai = lngIDPai
                
                arrNOVO(lngITEMARRAY).lngProdutoID = arrPROVARV(i).lngProdutoID
                arrNOVO(lngITEMARRAY).lngProdutoIDPai = arrPROVARV(i).lngProdutoIDPai
                arrNOVO(lngITEMARRAY).lngTipo = arrPROVARV(i).lngTipo
                arrNOVO(lngITEMARRAY).strPRODUTO = arrPROVARV(i).strPRODUTO
                arrNOVO(lngITEMARRAY).strProdutoPAI = arrPROVARV(i).strProdutoPAI
                arrNOVO(lngITEMARRAY).strQTDCONS = arrPROVARV(i).strQTDCONS
                arrNOVO(lngITEMARRAY).strUNIDADE = arrPROVARV(i).strUNIDADE
            
            End If
        ElseIf i = 1 Then
            arrNOVO(lngITEMARRAY).intAction2Do = arrPROVARV(i).intAction2Do
            arrNOVO(lngITEMARRAY).lngCODIGO = arrPROVARV(i).lngCODIGO
            arrNOVO(lngITEMARRAY).lngCODPAI = arrPROVARV(i).lngCODPAI
            arrNOVO(lngITEMARRAY).lngCodUniMed = arrPROVARV(i).lngCodUniMed
            
            arrNOVO(lngITEMARRAY).lngID = lngITEMARRAY
            arrNOVO(lngITEMARRAY).lngIDPai = -1
            
            arrNOVO(lngITEMARRAY).lngProdutoID = arrPROVARV(i).lngProdutoID
            arrNOVO(lngITEMARRAY).lngProdutoIDPai = arrPROVARV(i).lngProdutoIDPai
            arrNOVO(lngITEMARRAY).lngTipo = arrPROVARV(i).lngTipo
            arrNOVO(lngITEMARRAY).strPRODUTO = arrPROVARV(i).strPRODUTO
            arrNOVO(lngITEMARRAY).strProdutoPAI = arrPROVARV(i).strProdutoPAI
            arrNOVO(lngITEMARRAY).strQTDCONS = arrPROVARV(i).strQTDCONS
            arrNOVO(lngITEMARRAY).strUNIDADE = arrPROVARV(i).strUNIDADE
        End If
        
        If arrPROVARV(i).intAction2Do <> dacEnumUpdateAction_delete Then
            lngITEMARRAY = (lngITEMARRAY + 1)
        End If
    
    Next i

    
    arrPROVARV = arrNOVO
    
    '' Removendo da arvore
    treListaMat.Nodes.Remove lngINDLIST

End Sub
