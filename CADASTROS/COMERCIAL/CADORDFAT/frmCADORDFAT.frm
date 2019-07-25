VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCADORDFAT 
   Caption         =   "Cadastro de Orderm de Faturamento"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13035
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   13035
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame26 
      Caption         =   "[ Log ]"
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
      Left            =   6240
      TabIndex        =   68
      Top             =   5880
      Width           =   6735
      Begin VSFlex8LCtl.VSFlexGrid grdLogPed 
         Height          =   1935
         Left            =   120
         TabIndex        =   69
         Top             =   240
         Width           =   6495
         _cx             =   11456
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
   Begin VB.Frame fraNF 
      Caption         =   "[ Nota Fiscal Gerada ]"
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
      Height          =   615
      Left            =   6240
      TabIndex        =   57
      Top             =   8280
      Width           =   6735
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código da Confirmação"
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
         Index           =   17
         Left            =   120
         TabIndex        =   61
         Top             =   240
         Width           =   1980
      End
      Begin VB.Label lblCODCONFIRMACAO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblCODCONFIRMACAO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2160
         TabIndex        =   60
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblCODFATURA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblCODFATURA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5160
         TabIndex        =   59
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº da Nota Fiscal"
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
         Index           =   16
         Left            =   3600
         TabIndex        =   58
         Top             =   240
         Width           =   1515
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "[ Totais do Faturamento ]"
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
      Height          =   3015
      Left            =   0
      TabIndex        =   33
      Top             =   5880
      Width           =   6135
      Begin VB.TextBox txtFRETE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   4680
         TabIndex        =   5
         Text            =   "txtFRETE"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtOutrDesp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   1920
         TabIndex        =   4
         Text            =   "txtOutrDesp"
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblVLDESCTOTOT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblVLDESCTOTOT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4680
         TabIndex        =   53
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblPDESCTOTAL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblPDESCTOTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1920
         TabIndex        =   52
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblVLTOTAL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1920
         TabIndex        =   51
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Total da Fatura"
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
         TabIndex        =   50
         Top             =   2055
         Width           =   1320
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
         TabIndex        =   49
         Top             =   1680
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
         Index           =   0
         Left            =   120
         TabIndex        =   48
         Top             =   1680
         Width           =   1020
      End
      Begin VB.Label lblALIQICMS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblALIQICMS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4680
         TabIndex        =   47
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblVLICMS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1920
         TabIndex        =   46
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Valor do ICMS"
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
         TabIndex        =   45
         Top             =   645
         Width           =   1230
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Aliq ICMS"
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
         TabIndex        =   44
         Top             =   285
         Width           =   840
      End
      Begin VB.Label lblVLIPI 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1920
         TabIndex        =   43
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Valor do IPI"
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
         TabIndex        =   42
         Top             =   1350
         Width           =   1020
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Frete"
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
         TabIndex        =   41
         Top             =   990
         Width           =   450
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Outras despesas"
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
         TabIndex        =   40
         Top             =   1005
         Width           =   1425
      End
      Begin VB.Label lblBASICMS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1920
         TabIndex        =   39
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Base Calculo ICMS"
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
         TabIndex        =   38
         Top             =   285
         Width           =   1635
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "[ Itens do Pedido ]"
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
      Height          =   2175
      Left            =   0
      TabIndex        =   26
      Top             =   3720
      Width           =   12975
      Begin VSFlex8LCtl.VSFlexGrid grdITENSPEDIDO 
         Height          =   1455
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   12735
         _cx             =   22463
         _cy             =   2566
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
      Begin VB.Label lblTOTALFAT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblTOTALFAT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   11160
         TabIndex        =   37
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total de Faturamento"
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
         Index           =   14
         Left            =   9240
         TabIndex        =   36
         Top             =   1800
         Width           =   1830
      End
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   0
      TabIndex        =   24
      Top             =   3000
      Width           =   12975
      Begin VB.TextBox txtOBS 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Text            =   "frmCADORDFAT.frx":0000
         Top             =   120
         Width           =   11535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Observação"
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
         TabIndex        =   25
         Top             =   120
         Width           =   1035
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1575
      Left            =   0
      TabIndex        =   16
      Top             =   1440
      Width           =   12975
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Transporte"
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
         Left            =   7080
         TabIndex        =   67
         Top             =   1200
         Width           =   930
      End
      Begin VB.Label lblTRANSPORTE 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblTRANSPORTE"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8160
         TabIndex        =   66
         Top             =   1200
         Width           =   4695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Condiçào de Pagamento"
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
         TabIndex        =   65
         Top             =   1200
         Width           =   2085
      End
      Begin VB.Label lblCONDPGTO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblCONDPGTO"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2280
         TabIndex        =   64
         Top             =   1200
         Width           =   4695
      End
      Begin VB.Label lblBairro 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblBairro"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5880
         TabIndex        =   35
         Top             =   650
         Width           =   3855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bairro"
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
         Left            =   5160
         TabIndex        =   34
         Top             =   650
         Width           =   510
      End
      Begin VB.Label lblTELEFONE 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblTELEFONE"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   32
         Top             =   930
         Width           =   5895
      End
      Begin VB.Label lblESTADO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblESTADO"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   11160
         TabIndex        =   31
         Top             =   650
         Width           =   495
      End
      Begin VB.Label lblCIDADE 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblCIDADE"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   30
         Top             =   650
         Width           =   3855
      End
      Begin VB.Label lblENDERECO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblENDERECO"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   29
         Top             =   370
         Width           =   5895
      End
      Begin VB.Label lblINSCREST 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblINSCREST"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   11160
         TabIndex        =   28
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label lblCNPJCPF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblCNPJCPF"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8040
         TabIndex        =   27
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label lblCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblCliente"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   8
         Top             =   120
         Width           =   5895
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Telefone"
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
         TabIndex        =   23
         Top             =   930
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
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
         Left            =   9960
         TabIndex        =   22
         Top             =   650
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cidade"
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
         TabIndex        =   21
         Top             =   650
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Endereço"
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
         TabIndex        =   20
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inscr.Estadual"
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
         TabIndex        =   19
         Top             =   120
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ/CPF"
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
         Left            =   7080
         TabIndex        =   18
         Top             =   120
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
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
         TabIndex        =   17
         Top             =   120
         Width           =   600
      End
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   0
      TabIndex        =   12
      Top             =   960
      Width           =   12975
      Begin VB.CommandButton cmdPedido 
         Height          =   315
         Left            =   7440
         Picture         =   "frmCADORDFAT.frx":0009
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   120
         Width           =   375
      End
      Begin VB.TextBox txtStatus 
         Appearance      =   0  'Flat
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
         Left            =   11160
         TabIndex        =   55
         Text            =   "txtStatus"
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox txtCodPedido 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6120
         TabIndex        =   2
         Text            =   "txtCodPedido"
         Top             =   120
         Width           =   1335
      End
      Begin MSMask.MaskEdBox mskDataOrdem 
         Height          =   255
         Left            =   3840
         TabIndex        =   1
         Top             =   120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin VB.Label lblCODPED 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblCODPED"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9120
         TabIndex        =   63
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cód.Pedido"
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
         Left            =   8040
         TabIndex        =   62
         Top             =   120
         Width           =   990
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
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   15
         Left            =   10440
         TabIndex        =   54
         Top             =   120
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cód.OP"
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
         Left            =   5400
         TabIndex        =   15
         Top             =   120
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cód.Ordem"
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
         Top             =   150
         Width           =   945
      End
      Begin VB.Label lblCODIGO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblCODIGO"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data da Ordem"
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
         Left            =   2400
         TabIndex        =   13
         Top             =   120
         Width           =   1290
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   12975
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
         Picture         =   "frmCADORDFAT.frx":010B
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
         Picture         =   "frmCADORDFAT.frx":063D
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "frmCADORDFAT.frx":073F
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADORDFAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho             As String
Public Linha                As Variant
Public cTipOper             As String
Public iCodigo              As Long
Public FILIAL               As Integer
Public strAcesso            As String
Public strMODPAI            As String
Public strUSUARIO           As String
Public lngCodVendedor       As Long
Public lngCodUsuario        As Long
Public intFILIALPED         As Integer
Public strNOMMODULO         As String
Public intLIBERASN          As Integer
Public intLIB10PORC         As Integer


Dim arrITENSFAT         As Variant
Dim objBLBFunc          As Object
Dim objCADORDFAT        As Object
Dim objPESQPADRAO       As Object
Dim strLOGMODULO        As String
Dim strNOMTABELA1       As String
Dim strNOMTABELA2       As String
Dim strNOMTABELA3       As String
Dim strNOMTABELA4       As String
Dim strNOMFILIAL        As String

Const conCOL_Produto_IDProduto                As Integer = 0
Const conCOL_Produto_CodProduto               As Integer = 1
Const conCOL_Produto_DescProduto              As Integer = 2
Const conCOL_Produto_QtdeReal                 As Integer = 3
Const conCOL_Produto_QtdeJaFaturada           As Integer = 4
Const conCOL_Produto_QtdeFaturada             As Integer = 5
Const conCOL_Produto_Saldo                    As Integer = 6
Const conCOL_Produto_PorcIPI                  As Integer = 7
Const conCOL_Produto_VlUnit                   As Integer = 8
Const conCOL_Produto_VLReal                   As Integer = 9
Const conCOL_Produto_VLFaturado               As Integer = 10
Const conCOL_Produto_VLIPI                    As Integer = 11
Const conCOL_Produto_Action2Do                As Integer = 12
Const conCOL_Produto_OrdFab                   As Integer = 13
Const conCOL_Produto_CodForn                  As Integer = 14

Const conCOL_Produto_FormatString             As String = "=Código|Produto|Descrição|Qtde. Real|Qtd. Já Faturada|Qtd. a Faturar|Saldo|% IPI|Vl. Unitário|Vl.Total|Vl.Faturado|Vl. do IPI|Action2Do|Ord.Fab|Cod.Forn"
Const conColumnsIn_Produto                    As Integer = 15

'' ========================================================================================
Const conCOL_SonLogPed_Data                     As Integer = 0
Const conCOL_SonLogPed_Hora                     As Integer = 1
Const conCOL_SonLogPed_CodUsuario               As Integer = 2
Const conCOL_SonLogPed_Usuario                  As Integer = 3
Const conCOL_SonLogPed_CodAcao                  As Integer = 4
Const conCOL_SonLogPed_Acao                     As Integer = 5
Const conCOL_SonLogPed_Tipo                     As Integer = 6

Const conCOL_SonLogPed_FormatString             As String = "=Data|Hora|CodUsuario|Usuário|CodAcao|Ação|Tipo"
Const conColumnsIn_SonLogPed                    As Integer = 7


Private Sub cmdAltera_Click()

    If objCADORDFAT.STATUS = 1 Then
       MsgBox "Esta Ordem de Faturmanto já Foi Confirmada !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Sub
    End If
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    
    Me.Caption = "Cadastro de Ordem de Faturamento - [ ALTERAÇÃO ]"
    
    Frame2.Enabled = False
    Frame3.Enabled = False
    Frame8.Enabled = True
    txtOBS.Locked = False

    cTipOper = "A"

End Sub

Private Sub cmdPedido_Click()


    ReDim arrCAMPOS(1 To 6, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       ORD.SGI_CODPED     " & vbCrLf
    sSql = sSql & "      ,ORD.SGI_CODIGO     " & vbCrLf
    sSql = sSql & "      ,ORD.SGI_CODPROD    " & vbCrLf
    sSql = sSql & "      ,PROD.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "      ,PED.SGI_CODCLI     " & vbCrLf
    sSql = sSql & "      ,CLIE.SGI_RAZAOSOC  " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD" & strNOMFILIAL & "   ORD  " & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDI" & strNOMFILIAL & " PEDI " & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH" & strNOMFILIAL & " PED  " & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO  PROD " & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE  CLIE " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       ORD.SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "   And (ORD.SGI_STATUS    = 0 or ORD.SGI_STATUS = 1)" & vbCrLf
''    sSql = sSql & "   And ORD.SGI_OPENVIADA  = 3" & vbCrLf
    sSql = sSql & "   And PEDI.SGI_FILIAL    = ORD.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PEDI.SGI_CODIGO    = ORD.SGI_CODPED " & vbCrLf
    sSql = sSql & "   And PEDI.SGI_IDPRODUTO = ORD.SGI_IDPRODUTO " & vbCrLf
    sSql = sSql & "   And PROD.SGI_FILIAL    = PEDI.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PROD.SGI_IDPRODUTO = PEDI.SGI_IDPRODUTO " & vbCrLf
    sSql = sSql & "   And PED.SGI_FILIAL     = PEDI.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PED.SGI_CODIGO     = PEDI.SGI_CODIGO " & vbCrLf
    sSql = sSql & "   And CLIE.SGI_FILIAL    = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLIE.SGI_CODIGO    = PED.SGI_CODCLI " & vbCrLf
    sSql = sSql & "   And ((PEDI.SGI_QTDE - (Select " & vbCrLf
    sSql = sSql & "                                 Sum (ORDI.SGI_QTDFAT)" & vbCrLf
    sSql = sSql & "                            From" & vbCrLf
    sSql = sSql & "                                 SGI_CADORDFATH" & strNOMFILIAL & " ORDH" & vbCrLf
    sSql = sSql & "                                ,SGI_CADORDFATI" & strNOMFILIAL & " ORDI" & vbCrLf
    sSql = sSql & "                           Where" & vbCrLf
    sSql = sSql & "                                 ORDH.SGI_FILIAL    = ORD.SGI_FILIAL" & vbCrLf
    sSql = sSql & "                             And ORDH.SGI_CODPED    = ORD.SGI_CODPED" & vbCrLf
    sSql = sSql & "                             And ORDI.SGI_FILIAL    = ORDH.SGI_FILIAL" & vbCrLf
    sSql = sSql & "                             And ORDI.SGI_CODORD    = ORDH.SGI_CODORD" & vbCrLf
    sSql = sSql & "                             And ORDI.SGI_IDPRODUTO = ORD.SGI_IDPRODUTO)) > 0" & vbCrLf
    sSql = sSql & "      Or" & vbCrLf
    sSql = sSql & "        (PEDI.SGI_QTDE - (Select" & vbCrLf
    sSql = sSql & "                                 Sum (ORDI.SGI_QTDFAT)" & vbCrLf
    sSql = sSql & "                            From" & vbCrLf
    sSql = sSql & "                                 SGI_CADORDFATH" & strNOMFILIAL & "  ORDH" & vbCrLf
    sSql = sSql & "                                ,SGI_CADORDFATI" & strNOMFILIAL & "  ORDI" & vbCrLf
    sSql = sSql & "                           Where" & vbCrLf
    sSql = sSql & "                                 ORDH.SGI_FILIAL    = ORD.SGI_FILIAL" & vbCrLf
    sSql = sSql & "                             And ORDH.SGI_CODPED    = ORD.SGI_CODPED" & vbCrLf
    sSql = sSql & "                             And ORDI.SGI_FILIAL    = ORDH.SGI_FILIAL" & vbCrLf
    sSql = sSql & "                             And ORDI.SGI_CODORD    = ORDH.SGI_CODORD" & vbCrLf
    sSql = sSql & "                             And ORDI.SGI_IDPRODUTO = ORD.SGI_IDPRODUTO)) Is Null)"
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODIGO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código OP"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "ORD.SGI_CODIGO"
    
    arrCAMPOS(2, 1) = "SGI_CODPED"
    arrCAMPOS(2, 2) = "N"
    arrCAMPOS(2, 3) = "Cód.Pedido"
    arrCAMPOS(2, 4) = "1000"
    arrCAMPOS(2, 5) = "ORD.SGI_CODPED"
    
    arrCAMPOS(3, 1) = "SGI_CODPROD"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "Rótulo"
    arrCAMPOS(3, 4) = "2000"
    arrCAMPOS(3, 5) = "ORD.SGI_CODPROD"
    
    arrCAMPOS(4, 1) = "SGI_DESCRICAO"
    arrCAMPOS(4, 2) = "S"
    arrCAMPOS(4, 3) = "Descrição do Rótulo"
    arrCAMPOS(4, 4) = "3000"
    arrCAMPOS(4, 5) = "PROD.SGI_DESCRICAO"
    
    arrCAMPOS(5, 1) = "SGI_CODCLI"
    arrCAMPOS(5, 2) = "N"
    arrCAMPOS(5, 3) = "Cód.Cliente"
    arrCAMPOS(5, 4) = "1000"
    arrCAMPOS(5, 5) = "PED.SGI_CODCLI"
    
    arrCAMPOS(6, 1) = "SGI_RAZAOSOC"
    arrCAMPOS(6, 2) = "S"
    arrCAMPOS(6, 3) = "Razão Social"
    arrCAMPOS(6, 4) = "4000"
    arrCAMPOS(6, 5) = "CLIE.SGI_RAZAOSOC"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Pesquisa OP", "")
    
    If Len(Trim(varRETORNO)) > 0 Then
       txtCodPedido.Text = varRETORNO
       Call txtCodPedido_Validate(True)
       txtCodPedido.SetFocus
    End If

End Sub

Private Sub CmdSalva_Click()
    
On Error GoTo err_grava
    
    If Not VerifCampos Then Exit Sub
    
    Dim I           As Integer
    Dim intRESP     As Integer
    Dim lngCodLog   As Long
    Dim curQTDEFAT  As Currency
    
    
    If cTipOper = "I" Then objCADORDFAT.CODORD = objBLBFunc.Gera_Codigo(Trim(strNOMMODULO), FILIAL, Linha) & Year(Now)
    
    objCADORDFAT.CODPED = CLng(lblCODPED.Caption)
    objCADORDFAT.DATAORD = CDate(mskDataOrdem.Text)
    objCADORDFAT.OBS = Trim(Replace(txtOBS.Text, "'", ""))
    
    objCADORDFAT.BASEICMS = 0
    objCADORDFAT.ALIQICMS = 0
    objCADORDFAT.VALOICMS = 0
    objCADORDFAT.OUTRASDESP = 0
    objCADORDFAT.FRETE = 0
    objCADORDFAT.VALORIPI = 0
    objCADORDFAT.PORCDESCTO = 0
    objCADORDFAT.VALORDESCT = 0
    objCADORDFAT.VLTOTALFAT = 0
    
    If Len(Trim(lblBASICMS.Caption)) > 0 Then objCADORDFAT.BASEICMS = CCur(lblBASICMS.Caption)
    If Len(Trim(lblALIQICMS.Caption)) > 0 Then objCADORDFAT.ALIQICMS = CCur(lblALIQICMS.Caption)
    If Len(Trim(lblVLICMS.Caption)) > 0 Then objCADORDFAT.VALOICMS = CCur(lblVLICMS.Caption)
    If Len(Trim(txtOutrDesp.Text)) > 0 Then objCADORDFAT.OUTRASDESP = CCur(txtOutrDesp.Text)
    If Len(Trim(txtFRETE.Text)) > 0 Then objCADORDFAT.FRETE = CCur(txtFRETE.Text)
    If Len(Trim(lblVLIPI.Caption)) > 0 Then objCADORDFAT.VALORIPI = CCur(lblVLIPI.Caption)
    If Len(Trim(lblPDESCTOTAL.Caption)) > 0 Then objCADORDFAT.PORCDESCTO = CCur(lblPDESCTOTAL.Caption)
    If Len(Trim(lblVLDESCTOTOT.Caption)) > 0 Then objCADORDFAT.VALORDESCT = CCur(lblVLDESCTOTOT.Caption)
    If Len(Trim(lblVLTOTAL.Caption)) > 0 Then objCADORDFAT.VLTOTALFAT = CCur(lblVLTOTAL.Caption)
    
    '' Itens do Faturamento
    objCADORDFAT.ITENSFAT = Empty
    curQTDEFAT = 0
    With grdITENSPEDIDO
        If (.Rows - 1) > 0 Then
            ReDim arrITENSFAT(1 To (.Rows - 1), 1 To 14) As String
            For I = 1 To (.Rows - 1)
            
                arrITENSFAT(I, 1) = .Cell(flexcpText, I, conCOL_Produto_IDProduto)
                arrITENSFAT(I, 2) = .Cell(flexcpText, I, conCOL_Produto_CodProduto)
                arrITENSFAT(I, 3) = .Cell(flexcpText, I, conCOL_Produto_QtdeReal)
                arrITENSFAT(I, 4) = .Cell(flexcpText, I, conCOL_Produto_QtdeJaFaturada)
                arrITENSFAT(I, 5) = .Cell(flexcpText, I, conCOL_Produto_QtdeFaturada)
                arrITENSFAT(I, 6) = .Cell(flexcpText, I, conCOL_Produto_Saldo)
                arrITENSFAT(I, 7) = .Cell(flexcpText, I, conCOL_Produto_PorcIPI)
                arrITENSFAT(I, 8) = .Cell(flexcpText, I, conCOL_Produto_VlUnit)
                arrITENSFAT(I, 9) = .Cell(flexcpText, I, conCOL_Produto_VLReal)
                arrITENSFAT(I, 10) = .Cell(flexcpText, I, conCOL_Produto_VLFaturado)
                arrITENSFAT(I, 11) = .Cell(flexcpText, I, conCOL_Produto_VLIPI)
                arrITENSFAT(I, 12) = .Cell(flexcpText, I, conCOL_Produto_Action2Do)
                arrITENSFAT(I, 13) = .Cell(flexcpText, I, conCOL_Produto_OrdFab)
            
                If Len(Trim(.Cell(flexcpText, I, conCOL_Produto_CodForn))) > 0 Then
                    arrITENSFAT(I, 14) = Trim(Str(.Cell(flexcpText, I, conCOL_Produto_CodForn)))
                Else
                    arrITENSFAT(I, 14) = "Null"
                End If
                
                If Len(Trim(.Cell(flexcpText, I, conCOL_Produto_QtdeFaturada))) > 0 Then
                    curQTDEFAT = (curQTDEFAT + CCur(.Cell(flexcpText, I, conCOL_Produto_QtdeFaturada)))
                End If
            Next I
        End If
        objCADORDFAT.ITENSFAT = arrITENSFAT
    End With
    objCADORDFAT.QTDETOTALFAT = curQTDEFAT
    
    '' Gravando as Informações no banco
    If objCADORDFAT.GRAVA(cTipOper, intFILIALPED) = False Then Exit Sub
    
    Call objBLBFunc.GravaLogForm(FILIAL, objCADORDFAT.CODORD, lngCodUsuario, cTipOper, objCADORDFAT.NOMFORM)
  
    '' Atualizando os Dados
    If objBLBFunc.Atualiza(cTipOper, Str(objCADORDFAT.CODORD), FILIAL, Me.Name, Linha, Str(intFILIALPED)) = False Then Exit Sub
    
    '' Gerand Log de Sistema
    lngCodLog = objBLBFunc.Gera_Codigo(Trim(strLOGMODULO), FILIAL, Linha)
    Call objBLBFunc.GravaLogModulo(FILIAL, lngCodLog, Trim(strNOMMODULO), cTipOper, lngCodUsuario, Str(objCADORDFAT.CODORD), Linha)
    
    MsgBox "A Ordem de Faturamento nº ( " & Trim(Str(objCADORDFAT.CODORD)) & " ) foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
    
    If cTipOper = "I" Then
       intRESP = MsgBox("Deseja gerar outra Ordem de Faturamento ?", vbYesNo + vbQuestion, "Aviso")
       
       
       If intRESP = 6 Then
          Call Inclui
       Else
          Unload Me
       End If
    
    End If
    
    Exit Sub
    
err_grava:
    
    MsgBox "Erro nº  : " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, vbOKOnly + vbCritical, "Aviso"
    
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
    If KeyCode = vbKeyF5 Then
        intLIB10PORC = 0
        intLIBERASN = 0
        frmLIBERA.Linha = Linha
        frmLIBERA.FILIAL = FILIAL
        frmLIBERA.Show vbModal
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADORDFAT = CreateObject("CADORDFAT.clsCADORDFAT")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
   objCADORDFAT.FILIAL = FILIAL
   
   If adoBanco_Dados.State = 0 Then Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)
   If adoBanco_Dados.State = 0 Then
      MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
      Exit Sub
   End If
   
   If intFILIALPED = 0 Then strLOGMODULO = "SGI_LOGMODULO"
   If intFILIALPED = 1 Then strLOGMODULO = "SGI_LOGMODULO_STEEL"
   
    If intFILIALPED = 0 Then
       Me.Caption = Me.Caption & " / NOVALATA"
       strNOMTABELA1 = "SGI_CADORDFATH"
       strNOMTABELA2 = "SGI_CADPEDVENDH"
       strNOMTABELA3 = "SGI_CADORDFATI"
       strNOMTABELA4 = ""
       strNOMFILIAL = ""
    ElseIf intFILIALPED = 1 Then
       Me.Caption = Me.Caption & " / STEEL ROLL"
       strNOMTABELA1 = "SGI_CADORDFATH_STEEL"
       strNOMTABELA2 = "SGI_CADPEDVENDH_STEEL"
       strNOMTABELA3 = "SGI_CADORDFATI_STEEL"
       strNOMTABELA4 = ""
       strNOMFILIAL = "_STEEL"
    End If
   
   objCADORDFAT.NOMFORM = Trim(Me.Name & strNOMFILIAL)
   
   If cTipOper = "I" Then Call Inclui
   If cTipOper = "A" Then Call Altera
   If cTipOper = "C" Then Call Consulta


    If intFILIALPED = 0 Then Me.Caption = Me.Caption & " / NOVALATA"
    If intFILIALPED = 1 Then Me.Caption = Me.Caption & " / STEEL ROLL"

    Me.Caption = Me.Caption & " - Versão : " & App.Major & "." & App.Minor & "." & App.Revision

    objCADORDFAT.CODUSUARIO = lngCodUsuario

End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    lblCODIGO.Caption = ""
    lblCODPED.Caption = ""
    
    Frame2.Enabled = True
    Frame3.Enabled = True
    Frame8.Enabled = True
    txtOBS.Locked = False
    
    Me.Caption = "Cadastro de Ordem de Faturamento - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADORDFAT
    
    mskDataOrdem.Text = Format(Now, "DD/MM/YYYY")
    Call LimpaCamposLabel
    
    Call ConfGridProdutos
    Call InitGridLogPed
    
    txtStatus.Text = "Em Aberto"
    
    objCADORDFAT.STATUS = 0 '' Em Aberto
    
    Call AbilitNF("")
    
    intLIB10PORC = PegaLib10Porc
    intLIBERASN = PegaLibSN
    
End Sub


Private Sub LimpaCamposLabel()

    lblCliente.Caption = ""
    lblCNPJCPF.Caption = ""
    lblINSCREST.Caption = ""
    lblENDERECO.Caption = ""
    lblCIDADE.Caption = ""
    lblBairro.Caption = ""
    lblESTADO.Caption = ""
    lblTELEFONE.Caption = ""
    lblCONDPGTO.Caption = ""
    lblTRANSPORTE.Caption = ""
    
    lblTOTALFAT.Caption = ""
    lblBASICMS.Caption = ""
    lblALIQICMS.Caption = ""
    lblVLICMS.Caption = ""
    txtOutrDesp.Text = ""
    txtFRETE.Text = ""
    lblPDESCTOTAL.Caption = ""
    lblVLDESCTOTOT.Caption = ""
    lblVLIPI.Caption = ""
    lblVLTOTAL.Caption = ""


End Sub


Private Sub ConfGridProdutos()

    With grdITENSPEDIDO
    
       .Cols = conColumnsIn_Produto
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_Produto_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeNone
       
       .Cell(flexcpData, 0, conCOL_Produto_IDProduto) = ""
       .ColDataType(conCOL_Produto_IDProduto) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Produto_CodProduto) = ""
       .ColDataType(conCOL_Produto_CodProduto) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_Produto_DescProduto) = ""
       .ColDataType(conCOL_Produto_DescProduto) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_Produto_QtdeReal) = ""
       .ColDataType(conCOL_Produto_QtdeReal) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Produto_QtdeJaFaturada) = ""
       .ColDataType(conCOL_Produto_QtdeJaFaturada) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Produto_QtdeFaturada) = ""
       .ColDataType(conCOL_Produto_QtdeFaturada) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Produto_Saldo) = ""
       .ColDataType(conCOL_Produto_Saldo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Produto_PorcIPI) = ""
       .ColDataType(conCOL_Produto_PorcIPI) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_Produto_VlUnit) = ""
       .ColDataType(conCOL_Produto_VlUnit) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_Produto_VLReal) = ""
       .ColDataType(conCOL_Produto_VLReal) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_Produto_VLFaturado) = ""
       .ColDataType(conCOL_Produto_VLFaturado) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_Produto_VLIPI) = ""
       .ColDataType(conCOL_Produto_VLIPI) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_Produto_Action2Do) = ""
       .ColDataType(conCOL_Produto_Action2Do) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Produto_OrdFab) = ""
       .ColDataType(conCOL_Produto_OrdFab) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Produto_CodForn) = ""
       .ColDataType(conCOL_Produto_CodForn) = flexDTLong
       
       .ColWidth(conCOL_Produto_IDProduto) = 0
       .ColWidth(conCOL_Produto_CodProduto) = 1200
       .ColWidth(conCOL_Produto_DescProduto) = 3500
       .ColWidth(conCOL_Produto_QtdeReal) = 1200
       .ColWidth(conCOL_Produto_QtdeJaFaturada) = 1300
       .ColWidth(conCOL_Produto_QtdeFaturada) = 1200
       .ColWidth(conCOL_Produto_Saldo) = 1200
       .ColWidth(conCOL_Produto_PorcIPI) = 600
       .ColWidth(conCOL_Produto_VlUnit) = 1000
       .ColWidth(conCOL_Produto_VLReal) = 1200
       .ColWidth(conCOL_Produto_VLFaturado) = 1200
       .ColWidth(conCOL_Produto_VLIPI) = 1200
       .ColWidth(conCOL_Produto_Action2Do) = 0
       .ColWidth(conCOL_Produto_OrdFab) = 1000
       .ColWidth(conCOL_Produto_CodForn) = 0
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DestroiObjeto
End Sub

Private Sub grdITENSPEDIDO_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     With grdITENSPEDIDO
          Select Case Col
                 Case conCOL_Produto_QtdeFaturada, conCOL_Produto_VlUnit
                      If Col = conCOL_Produto_VlUnit Then
                        .Cell(flexcpText, Row, Col) = Format(.Cell(flexcpText, Row, Col), "#,##0.00")
                      End If
                      
                      .Cell(flexcpText, Row, conCOL_Produto_VLFaturado) = Format(CalcItenGrid(Row), "#,##0.00")
                      .Cell(flexcpText, Row, conCOL_Produto_Saldo) = CalcSaldo(Row)
                      .Cell(flexcpText, Row, conCOL_Produto_Action2Do) = dacEnumUpdateAction_update
                      Call CalcTotFatura
          End Select
     End With
End Sub

Private Sub grdITENSPEDIDO_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_Produto_CodProduto, _
         conCOL_Produto_DescProduto, _
         conCOL_Produto_QtdeReal, _
         conCOL_Produto_QtdeJaFaturada, _
         conCOL_Produto_Saldo, _
         conCOL_Produto_PorcIPI, _
         conCOL_Produto_VLReal, _
         conCOL_Produto_VLFaturado, _
         conCOL_Produto_VLIPI, _
         conCOL_Produto_OrdFab
         Cancel = True
    Case conCOL_Produto_QtdeFaturada, conCOL_Produto_VlUnit
         If cTipOper = "C" Then Cancel = True
    Case Else
        grdITENSPEDIDO.ComboList = ""
    End Select
    Exit Sub
End Sub

Private Sub grdITENSPEDIDO_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdITENSPEDIDO
          Select Case Col
                    Case conCOL_Produto_QtdeFaturada
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
                    Case conCOL_Produto_VlUnit
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 2, myvarAsCurrency)
          End Select
     End With
End Sub

Private Sub grdITENSPEDIDO_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

     Dim curQTDFATURADA As Currency
     
     With grdITENSPEDIDO
          Select Case Col
                Case conCOL_Produto_QtdeFaturada
                    If .EditText = Empty Then
                       MsgBox "ATENÇÂO" & vbCrLf & "Não é permitido valores nulo !!!", vbOKOnly + vbExclamation, "Aviso"
                       Cancel = True
                       Exit Sub
                    End If
                    curQTDFATURADA = CCur(.EditText)
                    Cancel = VerifSaldo(Row, curQTDFATURADA)
                Case conCOL_Produto_VlUnit
                    If .EditText = Empty Then Exit Sub
          End Select
     End With

End Sub


Private Sub mskDataOrdem_GotFocus()
    objBLBFunc.SelecionaCampos mskDataOrdem.Name, frmCADORDFAT
End Sub

Private Sub txtCodPedido_GotFocus()
    objBLBFunc.SelecionaCampos txtCodPedido.Name, frmCADORDFAT
End Sub

Private Sub txtCodPedido_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCodPedido.Text
End Sub

Private Sub txtCodPedido_Validate(Cancel As Boolean)
    
    If Len(Trim(txtCodPedido.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCodPedido.Text) Then
       MsgBox "Somente é Permitido Numeros !!!", vbOKOnly + vbExclamation, "Aviso"
       txtCodPedido.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Cancel = PegaDadosDoPedido(Trim(txtCodPedido.Text))
    
    
End Sub

Private Function PegaDadosDoPedido(strCodPedido As String) As Boolean

    PegaDadosDoPedido = False
    
    Call LimpaCamposLabel
    Call ConfGridProdutos
    txtOBS.Text = ""
    
    Dim curQTDEREAL  As Currency
    Dim curQTDEFAT   As Currency
    Dim curSaldo     As Currency
    Dim curVLAFAT    As Currency
    Dim curVLIPI     As Currency
    Dim curPORCIPI   As Currency
    
    Dim strCODPED       As String
    Dim strDTENTREGA    As String
    Dim strIDPROD       As String
    Dim strCODCLIE      As String
    
    '' ==============================
    '' Pega o Código do Pedido
    '' Data 23/06/2014
    strCODPED = ""
    strDTENTREGA = ""
    strIDPROD = ""
    strCODCLIE = ""
    
    intLIB10PORC = PegaLib10Porc
    intLIBERASN = PegaLibSN
    
    '' Pega Código do Pedido
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       ORDP.SGI_CODPED" & vbCrLf
    sSql = sSql & "      ,ORDP.SGI_DATENTREGA" & vbCrLf
    sSql = sSql & "      ,ORDP.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "      ,PED.SGI_CODCLI" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD" & strNOMFILIAL & " ORDP" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH" & strNOMFILIAL & " PED" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       ORDP.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And ORDP.SGI_CODIGO = " & strCodPedido & vbCrLf
    sSql = sSql & "   And PED.SGI_FILIAL  = ORDP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PED.SGI_CODIGO  = ORDP.SGI_CODPED"
    
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC2.EOF() Then
        strCODPED = BREC2!SGI_CODPED
        strDTENTREGA = "'" & Format(BREC2!SGI_DATENTREGA, "MM/DD/YYYY") & "'"
        strIDPROD = Trim(Str(BREC2!SGI_IDPRODUTO))
        strCODCLIE = Trim(Str(BREC2!SGI_CODCLI))
    End If
    BREC2.Close
    '' ==============================
    
    If Verifica_Credito(strCODCLIE) = "N" Then
       txtCodPedido.Text = ""
       Exit Function
    End If
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    
    sSql = sSql & "        SGI_TOTITENSFAT = (Select" & vbCrLf
    sSql = sSql & "                                  Sum (ORDI.SGI_QTDFAT)" & vbCrLf
    sSql = sSql & "                             From" & vbCrLf
    sSql = sSql & "                                  SGI_CADORDFATH" & strNOMFILIAL & " ORDH" & vbCrLf
    sSql = sSql & "                                 ,SGI_CADORDFATI" & strNOMFILIAL & " ORDI" & vbCrLf
    sSql = sSql & "                            Where" & vbCrLf
    sSql = sSql & "                                  ORDH.SGI_FILIAL    = PED.SGI_FILIAL" & vbCrLf
    sSql = sSql & "                              And ORDH.SGI_CODPED    = PED.SGI_CODIGO" & vbCrLf
    sSql = sSql & "                              And ORDI.SGI_FILIAL    = ORDH.SGI_FILIAL" & vbCrLf
    sSql = sSql & "                              And ORDI.SGI_CODORD    = ORDH.SGI_CODORD" & vbCrLf
    sSql = sSql & "                              And ORDI.SGI_IDPRODUTO = ORD.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "                              And ORDI.SGI_CODORDFAB = ORD.SGI_CODIGO)" & vbCrLf
    
    sSql = sSql & "      ,PED.* " & vbCrLf
    sSql = sSql & "      ,PED.SGI_CODIGO AS SGI_CODPED " & vbCrLf
    sSql = sSql & "      ,CLI.* " & vbCrLf
    sSql = sSql & "      ,PGT.SGI_DESCRICAO AS SGI_DESCPGTO " & vbCrLf
    sSql = sSql & "      ,TRA.SGI_DESCRICAO AS SGI_DESCTRANSP " & vbCrLf
    sSql = sSql & "      ,PED.SGI_OBS As SGI_OBSPED " & vbCrLf
    sSql = sSql & "      ,ORD.SGI_STATUS AS SGI_STATUSOP" & vbCrLf
    sSql = sSql & "      ,ORD.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "      ,ORD.SGI_QTDFAT As SGI_QTDFATORD" & vbCrLf
    sSql = sSql & "      ,ORD.SGI_QTDE   As SGI_QTDEORD" & vbCrLf
    sSql = sSql & "      ,ORD.SGI_OPENVIADA " & vbCrLf
    sSql = sSql & "      ,ORD.SGI_CODIGO As SGI_CODOP" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD" & strNOMFILIAL & "   ORD" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH" & strNOMFILIAL & " PED" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE  CLI" & vbCrLf
    sSql = sSql & "      ,SGI_CADCONDPGTO PGT" & vbCrLf
    sSql = sSql & "      ,SGI_CADTRANSP   TRA" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       ORD.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And ORD.SGI_CODIGO = " & Trim(strCodPedido) & vbCrLf
    sSql = sSql & "   And PED.SGI_FILIAL = ORD.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PED.SGI_CODIGO = ORD.SGI_CODPED " & vbCrLf
    sSql = sSql & "   And CLI.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLI.SGI_CODIGO = PED.SGI_CODCLI " & vbCrLf
    sSql = sSql & "   And PGT.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PGT.SGI_CODIGO = PED.SGI_CODCONDPGT " & vbCrLf
    sSql = sSql & "   And TRA.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And TRA.SGI_CODIGO = PED.SGI_CODTRANSP " & vbCrLf

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        
        If BREC!SGI_STATUSOP = 2 Then '' Finalizada ou Faturada
            MsgBox "ATENÇÂO - Esta Ordem de Produção já está Faturada !!!", vbOKOnly + vbExclamation, "A|viso"
            PegaDadosDoPedido = True
        ElseIf BREC!SGI_STATUSOP = 6 Then '' Bloqueada P.Cota
            MsgBox "ATENÇÂO - Esta Ordem de Produção esta Bloqueada por P.Cota !!!", vbOKOnly + vbExclamation, "A|viso"
            PegaDadosDoPedido = True
        ElseIf BREC!SGI_STATUSOP = 7 Then '' Bloqueada P.Data
            MsgBox "ATENÇÂO - Esta Ordem de Produção esta Bloqueada por P.Data !!!", vbOKOnly + vbExclamation, "A|viso"
            PegaDadosDoPedido = True
        ElseIf BREC!SGI_STATUSOP = 9 Then '' Finalizada Manualmente
            MsgBox "ATENÇÂO - Esta Ordem de Produção foi Finalizada Manualmente pelo PCP !!!", vbOKOnly + vbExclamation, "A|viso"
            PegaDadosDoPedido = True
        ElseIf BREC!SGI_STATUSOP = 3 Or BREC!SGI_STATUSOP = 4 Then
            MsgBox "ATENÇÂO - Esta Ordem de Produção esta Bloqueada favor informar o PCP !!!", vbOKOnly + vbExclamation, "A|viso"
            PegaDadosDoPedido = True
        End If
        
        If PegaDadosDoPedido = True Then
           BREC.Close
           txtCodPedido.Text = ""
           txtCodPedido.SetFocus
           Exit Function
        End If
        '' =====================================================
        
        
        ''If Not IsNull(BREC!SGI_QTDEITENSPEDIDO) Then objCADORDFAT.QTDETOTALPED = BREC!SGI_QTDEITENSPEDIDO
        If Not IsNull(BREC!SGI_QTDEORD) Then objCADORDFAT.QTDETOTALPED = BREC!SGI_QTDEORD
        
        objCADORDFAT.QTDEATENDPED = 0
        If Not IsNull(BREC!SGI_TOTITENSFAT) Then objCADORDFAT.QTDEATENDPED = BREC!SGI_TOTITENSFAT
        
        objCADORDFAT.FILIALEMP = BREC!SGI_FILIALPED
        
        '' =================================================================
        '' Verificando o Saldo a Faturar
        If (objCADORDFAT.QTDETOTALPED - objCADORDFAT.QTDEATENDPED) <= 0 Then
            MsgBox "Não há Mais Saldo na OP Para Gerar Ordem de Faturamento !!!", vbOKOnly + vbExclamation, "Aviso"
            txtCodPedido.Text = ""
            mskDataOrdem.Text = Format(Now, "DD/MM/YYYY")
            PegaDadosDoPedido = True
            BREC.Close
            Exit Function
        End If
        '' =================================================================
        
        lblCODPED.Caption = Trim(BREC!SGI_CODPED)
        lblCliente.Caption = Trim(BREC!SGI_RAZAOSOC)
        lblCNPJCPF.Caption = objBLBFunc.FormataCnpj(BREC!SGI_CPFCNPJ)
        lblINSCREST.Caption = BREC!SGI_RGCGC
        lblENDERECO.Caption = BREC!SGI_ENDNORM
        lblCIDADE.Caption = BREC!SGI_CIDNORM
        lblBairro.Caption = BREC!SGI_BAINROM
        lblCONDPGTO.Caption = BREC!SGI_DESCPGTO
        lblTRANSPORTE.Caption = BREC!SGI_DESCTRANSP
        
        lblTELEFONE.Caption = PegaTelefonesCli(Trim(Str(BREC!SGI_CODCLI)))
        lblESTADO.Caption = PegaEstado(Trim(Str(BREC!SGI_ESTNORM)))
        
        If Not IsNull(BREC!SGI_PORCICMS) Then lblALIQICMS.Caption = Format(BREC!SGI_PORCICMS, "#,##0.00")
        If Not IsNull(BREC!SGI_OUTRASTOT) Then txtOutrDesp.Text = Format(BREC!SGI_OUTRASTOT, "#,##0.00")
        If Not IsNull(BREC!SGI_FRETETOT) Then txtFRETE.Text = Format(BREC!SGI_FRETETOT, "#,##0.00")
        If Not IsNull(BREC!SGI_PORCDESCPED) Then lblPDESCTOTAL.Caption = Format(BREC!SGI_PORCDESCPED, "#,##0.00")
        
        If Not IsNull(BREC!SGI_OBSPED) Then txtOBS.Text = BREC!SGI_OBSPED
        
        ''
        '' Data da Alteração 23/06/2014
        '' Pegando OP's
        Do While Not BREC.EOF()
        
            '' Pega Itens do Pedido
            sSql = "Select " & vbCrLf
            sSql = sSql & "       ITEN.* " & vbCrLf
            sSql = sSql & "      ,PROD.* " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CADPEDVENDI" & strNOMFILIAL & " ITEN " & vbCrLf
            sSql = sSql & "      ,SGI_CADPRODUTO  PROD " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       ITEN.SGI_FILIAL    = " & FILIAL & vbCrLf
            sSql = sSql & "   And ITEN.SGI_CODIGO    = " & BREC!SGI_CODPED & vbCrLf
            sSql = sSql & "   And ITEN.SGI_IDPRODUTO = " & BREC!SGI_IDPRODUTO & vbCrLf
            sSql = sSql & "   And PROD.SGI_FILIAL    = ITEN.SGI_FILIAL " & vbCrLf
            sSql = sSql & "   And PROD.SGI_IDPRODUTO = ITEN.SGI_IDPRODUTO "
            
            BREC5.Open sSql, adoBanco_Dados, adOpenDynamic
            If Not BREC5.EOF() Then
                Do While Not BREC5.EOF()
                
                   With grdITENSPEDIDO
                   
                        curQTDEREAL = BREC!SGI_QTDEORD
                        
                        curQTDEFAT = 0
                        If Not IsNull(BREC!SGI_TOTITENSFAT) Then curQTDEFAT = BREC!SGI_TOTITENSFAT
                        
                        curSaldo = (curQTDEREAL - curQTDEFAT)
                   
                        If curSaldo > 0 Then
                        
                            curPORCIPI = 0
                            If Not IsNull(BREC5!SGI_PRCIPI) Then curPORCIPI = BREC5!SGI_PRCIPI
                            
                            curVLAFAT = CalcVlTotal(curSaldo, BREC5!SGI_VLUNIT)
                            curVLIPI = CalculaIPI(curVLAFAT, curPORCIPI)
                            
                            .AddItem BREC5!SGI_IDPRODUTO & vbTab & _
                                     Trim(BREC5!SGI_CODPROD) & vbTab & _
                                     Trim(BREC5!SGI_DESCRICAO) & vbTab & _
                                     BREC!SGI_QTDEORD & vbTab & _
                                     curQTDEFAT & vbTab & _
                                     curSaldo & vbTab & _
                                     "" & vbTab & _
                                     Format(BREC5!SGI_PRCIPI, "#,##0.00") & vbTab & _
                                     Format(BREC5!SGI_VLUNIT, "#,##0.00") & vbTab & _
                                     Format(BREC5!SGI_VLTOT, "#,##0.00") & vbTab & _
                                     Format((curVLAFAT + curVLIPI), "#,##0.00") & vbTab & _
                                     IIf(curVLIPI = 0, "", Format(curVLIPI, "#,##0.00")) & vbTab & _
                                     dacEnumUpdateAction_Insert & vbTab & _
                                     BREC!SGI_CODOP & vbTab & _
                                     BREC5!SGI_CODFORN
                            
                            .Cell(flexcpText, .Rows - 1, conCOL_Produto_Saldo) = CalcSaldo(.Rows - 1)
                            .Cell(flexcpBackColor, (.Rows - 1), conCOL_Produto_QtdeFaturada) = &H80C0FF
                            .Cell(flexcpBackColor, (.Rows - 1), conCOL_Produto_VLFaturado) = &H80C0FF
                            .Cell(flexcpBackColor, (.Rows - 1), conCOL_Produto_VLIPI) = &H80C0FF
                        
                        End If
                        
                   End With
                    
                   BREC5.MoveNext
                Loop
                Call CalcTotFatura
            End If
            BREC5.Close
            
            BREC.MoveNext
        Loop
    Else
        MsgBox "Esta OP Não Existe !!!", vbOKOnly + vbExclamation, "Aviso"
        PegaDadosDoPedido = True
        txtCodPedido.Text = ""
    End If
    BREC.Close
    
End Function
 

Private Function PegaTelefonesCli(strCODCLIE As String) As String

   PegaTelefonesCli = ""

   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_DADOSGERAIS " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL   = " & FILIAL & vbCrLf
   sSql = sSql & "   And SGI_CODIGO   = " & strCODCLIE & vbCrLf
   sSql = sSql & "   And SGI_MODULO   = 'frmCADCLIENTE'" & vbCrLf
   sSql = sSql & "   And SGI_CONTROLE = 'cboTELNORM'" & vbCrLf
   
   BREC4.Open sSql, adoBanco_Dados, adOpenDynamic
   Do While Not BREC4.EOF()
       PegaTelefonesCli = PegaTelefonesCli & BREC4!SGI_TEXTO
       BREC4.MoveNext
       If Not BREC4.EOF() Then PegaTelefonesCli = PegaTelefonesCli & "/"
   Loop
   BREC4.Close

End Function

Private Function PegaEstado(strCODESTADO As Long)

   If strCODESTADO = 0 Then Exit Function
   
   Dim V_Estado() As String
   ReDim V_Estado(1 To 28) As String
   
   V_Estado(1) = "AM"
   V_Estado(2) = "AC"
   V_Estado(3) = "AL"
   V_Estado(4) = "AP"
   V_Estado(5) = "BA"
   V_Estado(6) = "CE"
   V_Estado(7) = "DF"
   V_Estado(8) = "ES"
   V_Estado(9) = "GO"
   V_Estado(10) = "MA"
   V_Estado(11) = "MG"
   V_Estado(12) = "MT"
   V_Estado(13) = "MS"
   V_Estado(14) = "PE"
   V_Estado(15) = "PA"
   V_Estado(16) = "PB"
   V_Estado(17) = "PI"
   V_Estado(18) = "PR"
   V_Estado(19) = "RJ"
   V_Estado(20) = "RN"
   V_Estado(21) = "RO"
   V_Estado(22) = "RR"
   V_Estado(23) = "RS"
   V_Estado(24) = "SC"
   V_Estado(25) = "SE"
   V_Estado(26) = "SP"
   V_Estado(27) = "TO"
   V_Estado(28) = "EX"
   
   PegaEstado = Trim(V_Estado(CLng(strCODESTADO)))
   
End Function

Private Function CalcVlTotal(curQtdeAFat As Currency, curVlUnitario As Currency) As Currency
    CalcVlTotal = (curQtdeAFat * curVlUnitario)
End Function

Private Function CalculaIPI(VlTotal As Currency, PorcIPI As Currency) As Currency
    CalculaIPI = ((VlTotal * PorcIPI) / 100)
End Function

Private Function CalcItenGrid(lngRow As Long) As Currency
    CalcItenGrid = 0
    
    Dim curQtd_do_Item      As Currency
    Dim curVlUn_do_Item     As Currency
    Dim curDESC_do_Item     As Currency
    Dim curVLDESC_do_Item   As Currency
    Dim curIPI_do_Item      As Currency
    Dim curVlIPI_do_Item    As Currency
    Dim curTotal_do_Iten    As Currency
    
    curQtd_do_Item = 0
    curVlUn_do_Item = 0
    curDESC_do_Item = 0
    curVLDESC_do_Item = 0
    curIPI_do_Item = 0
    curVlIPI_do_Item = 0
    curTotal_do_Iten = 0
    
    With grdITENSPEDIDO
         If Len(Trim(.Cell(flexcpText, lngRow, conCOL_Produto_QtdeFaturada))) > 0 Then curQtd_do_Item = CCur(.Cell(flexcpText, lngRow, conCOL_Produto_QtdeFaturada))
         If Len(Trim(.Cell(flexcpText, lngRow, conCOL_Produto_VlUnit))) > 0 Then curVlUn_do_Item = CCur(.Cell(flexcpText, lngRow, conCOL_Produto_VlUnit))
         If Len(Trim(.Cell(flexcpText, lngRow, conCOL_Produto_PorcIPI))) > 0 Then curIPI_do_Item = CCur(.Cell(flexcpText, lngRow, conCOL_Produto_PorcIPI))
         
         curTotal_do_Iten = (curQtd_do_Item * curVlUn_do_Item)
         .Cell(flexcpText, lngRow, conCOL_Produto_VLFaturado) = Format(curTotal_do_Iten, "#,##0.00")
         
         '' Desconto
         ''curVLDESC_do_Item = ((curTotal_do_Iten * curDESC_do_Item) / 100)
         ''curTotal_do_Iten = (curTotal_do_Iten - curVLDESC_do_Item)
         
         '' IPI
         curVlIPI_do_Item = ((curTotal_do_Iten * curIPI_do_Item) / 100)
         curTotal_do_Iten = (curTotal_do_Iten + curVlIPI_do_Item)
         
         ''.Cell(flexcpText, lngRow, conCOL_SonProd_VlDesc) = Format(curVLDESC_do_Item, "#,##0.00")
         If curVlIPI_do_Item > 0 Then .Cell(flexcpText, lngRow, conCOL_Produto_VLIPI) = Format(curVlIPI_do_Item, "#,##0.00")
    End With
    
    CalcItenGrid = curTotal_do_Iten
End Function


Private Function CalcSaldo(lngRow As Long) As Currency

    CalcSaldo = 0
    
    Dim curQTD_FATURADA     As Currency
    Dim curQtd_Ja_Faturada  As Currency
    Dim curQtd_Real         As Currency
    Dim curSaldo_Calc       As Currency
    
    With grdITENSPEDIDO
    
        If Len(Trim(.Cell(flexcpText, lngRow, conCOL_Produto_QtdeReal))) > 0 Then curQtd_Real = CCur(.Cell(flexcpText, lngRow, conCOL_Produto_QtdeReal))
        If Len(Trim(.Cell(flexcpText, lngRow, conCOL_Produto_QtdeJaFaturada))) > 0 Then curQtd_Ja_Faturada = CCur(.Cell(flexcpText, lngRow, conCOL_Produto_QtdeJaFaturada))
        If Len(Trim(.Cell(flexcpText, lngRow, conCOL_Produto_QtdeFaturada))) > 0 Then curQTD_FATURADA = CCur(.Cell(flexcpText, lngRow, conCOL_Produto_QtdeFaturada))
    
        curSaldo_Calc = (curQtd_Real - curQtd_Ja_Faturada)
        CalcSaldo = (curSaldo_Calc - curQTD_FATURADA)
    
    End With

End Function


Private Function VerifSaldo(lngRow As Long, curQTD_FATURADA As Currency) As Boolean

    VerifSaldo = False
    
    
    Dim curQtd_Ja_Faturada  As Currency
    Dim curQtd_Real         As Currency
    Dim curSaldo_Calc       As Currency
    Dim curTolerancia       As Currency
    Dim curQtdTolerancia    As Currency
    
    With grdITENSPEDIDO
    
        If Len(Trim(.Cell(flexcpText, lngRow, conCOL_Produto_QtdeReal))) > 0 Then curQtd_Real = CCur(.Cell(flexcpText, lngRow, conCOL_Produto_QtdeReal))
        If Len(Trim(.Cell(flexcpText, lngRow, conCOL_Produto_QtdeJaFaturada))) > 0 Then curQtd_Ja_Faturada = CCur(.Cell(flexcpText, lngRow, conCOL_Produto_QtdeJaFaturada))
    
        curSaldo_Calc = (curQtd_Real - curQtd_Ja_Faturada)
         
        ''If PermiteLibFaturamento = True Then Exit Function
        
        '' Novalata
        ''If intFILIALPED = 0 Then
        ''    If lngCodUsuario = 0 Then Exit Function
        ''    If intLIB10PORC = 1 Then Exit Function
        ''End If
        
        If lngCodUsuario = 0 Then Exit Function
        If intLIB10PORC = 1 Then Exit Function
        
        curTolerancia = 0.1
        curQtdTolerancia = (curSaldo_Calc * curTolerancia)
        
        curSaldo_Calc = (curSaldo_Calc + curQtdTolerancia)
        If curQTD_FATURADA > curSaldo_Calc Then
           MsgBox "ATENÇÃO" & "O Usuário não tem permissão para Faturar quantidades maior que 10% do Saldo da OP !!!", vbOKOnly + vbExclamation, "Aviso"
           VerifSaldo = True
        End If
    
    End With

End Function

Private Sub CalcTotFatura()

    Dim I                   As Integer
    Dim curBaseCalculo      As Currency
    Dim curQTDEFAT          As Currency
    Dim curVlUnitario       As Currency
    
    Dim curALIQICMS         As Currency
    Dim curValICMS          As Currency
    Dim curVALORIPI         As Currency
    ''Dim curTotalDescIten    As Currency
    Dim curValOutrDesp      As Currency
    Dim curValFrete         As Currency
    Dim curPercDescPedido   As Currency
    Dim curDescPedido       As Currency
    Dim curTotalFatura      As Currency
    Dim curToTGerFatura     As Currency
    
    lblTOTALFAT.Caption = ""
    lblBASICMS.Caption = ""
    lblVLICMS.Caption = ""
    lblVLDESCTOTOT.Caption = ""
    lblVLIPI.Caption = ""
    lblVLTOTAL.Caption = ""
    
    curBaseCalculo = 0
    curALIQICMS = 0
    curValICMS = 0
    curVALORIPI = 0
    ''curTotalDescIten = 0
    curTotalFatura = 0
    curValOutrDesp = 0
    curValFrete = 0
    curPercDescPedido = 0
    curDescPedido = 0
    curToTGerFatura = 0
    
    With grdITENSPEDIDO
        For I = 1 To (.Rows - 1)
            curQTDEFAT = 0
            If Len(Trim(.Cell(flexcpText, I, conCOL_Produto_QtdeFaturada))) > 0 Then curQTDEFAT = CCur(.Cell(flexcpText, I, conCOL_Produto_QtdeFaturada))
            curVlUnitario = CCur(.Cell(flexcpText, I, conCOL_Produto_VlUnit))
        
            curBaseCalculo = curBaseCalculo + (curQTDEFAT * curVlUnitario)
            ''curTotalDescIten = curTotalDescIten + CCur(grdProduto.Cell(flexcpText, I, conCOL_SonProd_VlDesc))
            
            If Len(Trim(.Cell(flexcpText, I, conCOL_Produto_VLIPI))) > 0 Then curVALORIPI = curVALORIPI + CCur(.Cell(flexcpText, I, conCOL_Produto_VLIPI))
            curTotalFatura = curTotalFatura + CCur(.Cell(flexcpText, I, conCOL_Produto_VLFaturado))
        Next I
    End With
    lblTOTALFAT.Caption = Format(curTotalFatura, "#,##0.00")
    lblBASICMS.Caption = Format(curBaseCalculo, "#,##0.00")
    
    '' Calcula ICMS
    If Len(Trim(lblALIQICMS.Caption)) > 0 Then
        curALIQICMS = CCur(lblALIQICMS.Caption)
        curValICMS = ((curBaseCalculo * curALIQICMS) / 100)
        lblVLICMS.Caption = Format(curValICMS, "#,##0.00")
    End If
    If Len(Trim(lblPDESCTOTAL.Caption)) > 0 Then
        curPercDescPedido = CCur(lblPDESCTOTAL.Caption)
        curDescPedido = ((curBaseCalculo * curPercDescPedido) / 100)
        lblVLDESCTOTOT.Caption = Format(curDescPedido, "#,##0.00")
    End If
    
    ''lblVLDESCONTO.Caption = Format(curTotalDescIten, "#,##0.00")
    If curVALORIPI > 0 Then lblVLIPI.Caption = Format(curVALORIPI, "#,##0.00")
    
    If Len(Trim(txtOutrDesp.Text)) > 0 Then curValOutrDesp = CCur(txtOutrDesp.Text)
    If Len(Trim(txtFRETE.Text)) > 0 Then curValFrete = CCur(txtFRETE.Text)
    
    curToTGerFatura = ((curBaseCalculo + curVALORIPI + curValOutrDesp + curValFrete) - curDescPedido)
    
    lblVLTOTAL.Caption = Format(curToTGerFatura, "#,##0.00")
    
End Sub

Private Sub txtFRETE_GotFocus()
    objBLBFunc.SelecionaCampos txtFRETE.Name, frmCADORDFAT
End Sub

Private Sub txtFRETE_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtFRETE.Text
End Sub

Private Sub txtFRETE_Validate(Cancel As Boolean)
    If Len(Trim(txtFRETE.Text)) > 0 Then txtFRETE.Text = Format(txtFRETE.Text, "#,##0.00")
    Call CalcTotFatura
End Sub

Private Sub txtOutrDesp_GotFocus()
    objBLBFunc.SelecionaCampos txtOutrDesp.Name, frmCADORDFAT
End Sub

Private Sub txtOutrDesp_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtOutrDesp.Text
End Sub

Private Sub txtOutrDesp_Validate(Cancel As Boolean)
    If Len(Trim(txtOutrDesp.Text)) > 0 Then txtOutrDesp.Text = Format(txtOutrDesp.Text, "#,##0.00")
    Call CalcTotFatura
End Sub

Private Function VerifCampos() As Boolean
    
    VerifCampos = False
    
    Dim I                   As Integer
    Dim curQTD_FATURADA     As Currency
    Dim curQTD_TOTFATUR     As Currency
    Dim curValor_TOLERANCIA As Currency
    Dim lngTEM_VAZIO        As Long
    
    If Not IsDate(mskDataOrdem.Text) Then
       MsgBox "Data Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
       mskDataOrdem.SetFocus
       Exit Function
    End If
    If Len(Trim(txtCodPedido.Text)) = 0 Then
       MsgBox "O Código do Pedido Não foi Informado !!!", vbOKOnly + vbExclamation, "Aviso"
       txtCodPedido.SetFocus
       Exit Function
    End If
    
    '' ------------------------------------------
    curQTD_TOTFATUR = 0
    With grdITENSPEDIDO
        
        lngTEM_VAZIO = 0
        For I = 1 To (.Rows - 1)
            If Len(Trim(.Cell(flexcpText, I, conCOL_Produto_QtdeFaturada))) = 0 Then
               lngTEM_VAZIO = (lngTEM_VAZIO + 1)
            Else
                If IsNumeric(.Cell(flexcpText, I, conCOL_Produto_QtdeFaturada)) Then
                    If .Cell(flexcpText, I, conCOL_Produto_QtdeFaturada) = 0 Then lngTEM_VAZIO = (lngTEM_VAZIO + 1)
                End If
            End If
        Next I
        
        
        For I = 1 To (.Rows - 1)
            curQTD_FATURADA = 0
            If Len(Trim(.Cell(flexcpText, I, conCOL_Produto_QtdeFaturada))) > 0 Then curQTD_FATURADA = CCur(.Cell(flexcpText, I, conCOL_Produto_QtdeFaturada))
            curQTD_TOTFATUR = (curQTD_TOTFATUR + curQTD_FATURADA)
        Next I
    
    End With
    
    ''If intFILIALPED = 1 Then
        ''If lngTEM_VAZIO > 0 Then
        ''    MsgBox "ATENÇÃO" & vbCrLf & _
                   "Existe(m) Iten(s) com Quantidade a Faturar vázio !!!", vbOKOnly + vbExclamation, "Aviso"

                   ''"Existe(m) Iten(s) com Quantidade a Faturar vázio !!!" & vbCrLf & _
                   ''"Precisa de Senha de Liberação !!!", vbOKOnly + vbExclamation, "Aviso"
        ''    Exit Function
        ''End If
    ''Else
        ''If lngCodUsuario > 0 Then
        ''    If intLIBERASN = 0 Then
        ''        If lngTEM_VAZIO > 0 Then
        ''            MsgBox "ATENÇÃO" & vbCrLf & _
        ''                   "Existe(m) Iten(s) com Quantidade a Faturar vázio !!!" & vbCrLf & _
        ''                   "Precisa de Senha de Liberação !!!", vbOKOnly + vbExclamation, "Aviso"
        ''            Exit Function
        ''        End If
        ''    End If
        ''End If
    ''End If
    
    ''If lngCodUsuario > 0 Then
    ''    If intLIBERASN = 0 Then
    ''        If curQTD_TOTFATUR = 0 Then
    ''           MsgBox "Não foi informada nenhuma qtde a faturar, a Ordem de Faturamento não pode ser salva !!!", vbOKOnly + vbExclamation, "Aviso"
    ''           Exit Function
    ''        Else
    ''            curValor_TOLERANCIA = (objCADORDFAT.QTDETOTALPED * 0.1)
    ''            If ((curQTD_TOTFATUR + objCADORDFAT.QTDEATENDPED) > (objCADORDFAT.QTDETOTALPED + curValor_TOLERANCIA)) Then
    ''                 MsgBox "A Qtde Faturada não pode ser maior que a Qtde Total do Pedido !!!", vbOKOnly + vbExclamation, "Aviso"
    ''                 Exit Function
    ''            End If
    ''        End If
    ''        '' ------------------------------------------
    ''    End If
    ''End If
    
    VerifCampos = True
    
End Function


Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    
    Me.Caption = "Cadastro de Ordem de Faturamento - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos Me

    Frame2.Enabled = False
    Frame3.Enabled = False
    Frame8.Enabled = False
    txtOBS.Locked = True
    
    objCADORDFAT.CODORD = iCodigo
    
    intLIB10PORC = PegaLib10Porc
    intLIBERASN = PegaLibSN
    
    Call ConfGridProdutos
    Call InitGridLogPed
    
    Call LimpaCamposLabel
    
    Call CarregaCampos
    
End Sub

Private Sub CarregaDadosCliente(strCODPED As String)

    sSql = "Select " & vbCrLf
    sSql = sSql & "       PED.* " & vbCrLf
    sSql = sSql & "      ,PED.SGI_CODIGO AS SGI_CODPED " & vbCrLf
    sSql = sSql & "      ,CLI.* " & vbCrLf
    sSql = sSql & "      ,PGT.SGI_DESCRICAO AS SGI_DESCPGTO " & vbCrLf
    sSql = sSql & "      ,TRA.SGI_DESCRICAO AS SGI_DESCTRANSP " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & strNOMTABELA2 & " PED" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE  CLI" & vbCrLf
    sSql = sSql & "      ,SGI_CADCONDPGTO PGT" & vbCrLf
    sSql = sSql & "      ,SGI_CADTRANSP   TRA" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       PED.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And PED.SGI_CODIGO = " & Trim(strCODPED) & vbCrLf
    sSql = sSql & "   And CLI.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLI.SGI_CODIGO = PED.SGI_CODCLI " & vbCrLf
    sSql = sSql & "   And PGT.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PGT.SGI_CODIGO = PED.SGI_CODCONDPGT " & vbCrLf
    sSql = sSql & "   And TRA.SGI_FILIAL = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And TRA.SGI_CODIGO = PED.SGI_CODTRANSP " & vbCrLf

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
    
        lblCliente.Caption = Trim(BREC!SGI_RAZAOSOC)
        lblCNPJCPF.Caption = objBLBFunc.FormataCnpj(BREC!SGI_CPFCNPJ)
        lblINSCREST.Caption = BREC!SGI_RGCGC
        lblENDERECO.Caption = BREC!SGI_ENDNORM
        lblCIDADE.Caption = BREC!SGI_CIDNORM
        lblBairro.Caption = BREC!SGI_BAINROM
        lblCONDPGTO.Caption = BREC!SGI_DESCPGTO
        lblTRANSPORTE.Caption = BREC!SGI_DESCTRANSP
        
        lblTELEFONE.Caption = PegaTelefonesCli(Trim(Str(BREC!SGI_CODCLI)))
        lblESTADO.Caption = PegaEstado(Trim(Str(BREC!SGI_ESTNORM)))
    
    End If
    BREC.Close

End Sub

Private Sub CarregaCampos()

    If objCADORDFAT.Carrega_Campos(intFILIALPED) = True Then
        
        If objCADORDFAT.STATUS = 0 Then txtStatus.Text = "Em Aberto"
        If objCADORDFAT.STATUS = 1 Then txtStatus.Text = "Faturado"
        
        
        lblCODIGO.Caption = objCADORDFAT.CODORD
        lblCODPED.Caption = objCADORDFAT.CODPED
        mskDataOrdem.Text = Format(objCADORDFAT.DATAORD, "DD/MM/YYYY")
        txtCodPedido.Text = ""
        txtOBS.Text = objCADORDFAT.OBS
        
        Call CarregaDadosCliente(lblCODPED.Caption)
       
        If objCADORDFAT.BASEICMS > 0 Then lblBASICMS.Caption = Format(objCADORDFAT.BASEICMS, "#,##0.00")
        If objCADORDFAT.ALIQICMS > 0 Then lblALIQICMS.Caption = Format(objCADORDFAT.ALIQICMS, "#,##0.00")
        If objCADORDFAT.VALOICMS > 0 Then lblVLICMS.Caption = Format(objCADORDFAT.VALOICMS, "#,##0.00")
        If objCADORDFAT.OUTRASDESP > 0 Then txtOutrDesp.Text = Format(objCADORDFAT.OUTRASDESP, "#,##0.00")
        If objCADORDFAT.FRETE > 0 Then txtFRETE.Text = Format(objCADORDFAT.FRETE, "#,##0.00")
        If objCADORDFAT.VALORIPI > 0 Then lblVLIPI.Caption = Format(objCADORDFAT.VALORIPI, "#,##0.00")
        If objCADORDFAT.PORCDESCTO > 0 Then lblPDESCTOTAL.Caption = Format(objCADORDFAT.PORCDESCTO, "#,##0.00")
        If objCADORDFAT.VALORDESCT > 0 Then lblVLDESCTOTOT.Caption = Format(objCADORDFAT.VALORDESCT, "#,##0.00")
        If objCADORDFAT.VLTOTALFAT > 0 Then lblVLTOTAL.Caption = Format(objCADORDFAT.VLTOTALFAT, "#,##0.00")
    
        Call PopGrdItens
        Call PopLogPedidos
        Call AbilitNF(Str(objCADORDFAT.CODORD))
        
    End If

End Sub

Private Sub PopGrdItens()

    Dim I As Integer
    
    arrITENSFAT = objCADORDFAT.ITENSFAT

    If IsArray(arrITENSFAT) Then
        With grdITENSPEDIDO
            For I = 1 To UBound(arrITENSFAT)
            
                If Len(Trim(arrITENSFAT(I, 5))) > 0 Then
            
                .AddItem arrITENSFAT(I, 1) & vbTab & _
                         arrITENSFAT(I, 2) & vbTab & _
                         PegaDescProd(Str(arrITENSFAT(I, 1))) & vbTab & _
                         arrITENSFAT(I, 3) & vbTab & _
                         arrITENSFAT(I, 4) & vbTab & _
                         arrITENSFAT(I, 5) & vbTab & _
                         Format(arrITENSFAT(I, 6), "#,##0.00") & vbTab & _
                         Format(arrITENSFAT(I, 7), "#,##0.00") & vbTab & _
                         Format(arrITENSFAT(I, 8), "#,##0.00") & vbTab & _
                         Format(arrITENSFAT(I, 9), "#,##0.00") & vbTab & _
                         Format(arrITENSFAT(I, 10), "#,##0.00") & vbTab & _
                         Format(arrITENSFAT(I, 11), "#,##0.00") & vbTab & _
                         arrITENSFAT(I, 12) & vbTab & _
                         arrITENSFAT(I, 13) & vbTab & _
                         arrITENSFAT(I, 14)
                         
                txtCodPedido.Text = arrITENSFAT(I, 13)
                
                .Cell(flexcpBackColor, (.Rows - 1), conCOL_Produto_QtdeFaturada) = &H80C0FF
                .Cell(flexcpBackColor, (.Rows - 1), conCOL_Produto_VLFaturado) = &H80C0FF
                .Cell(flexcpBackColor, (.Rows - 1), conCOL_Produto_VLIPI) = &H80C0FF
                
                .Cell(flexcpText, (.Rows - 1), conCOL_Produto_VLFaturado) = Format(CalcItenGrid((.Rows - 1)), "#,##0.00")
                .Cell(flexcpText, (.Rows - 1), conCOL_Produto_Saldo) = CalcSaldo((.Rows - 1))
            
                End If
            Next I
            
            Call CalcTotFatura
        End With
    End If

End Sub

Private Function PegaDescProd(idProduto As String) As String

    PegaDescProd = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_IDPRODUTO = " & Trim(idProduto)
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PegaDescProd = Trim(BREC!SGI_DESCRICAO)
    BREC.Close
    
End Function

Private Sub Altera()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    
    Me.Caption = "Cadastro de Ordem de Faturamento - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADORDFAT

    Frame2.Enabled = False
    Frame3.Enabled = False
    Frame8.Enabled = True
    txtOBS.Locked = False
    
    intLIB10PORC = PegaLib10Porc
    intLIBERASN = PegaLibSN
    
    objCADORDFAT.CODORD = iCodigo
    
    Call ConfGridProdutos
    Call InitGridLogPed
    
    Call LimpaCamposLabel
    
    Call CarregaCampos
    
    
End Sub


Private Function Calc_QtdeJaFaturada(lngCODPEDIDO As Long, lngIDPRODUTO As Long) As Currency

    Calc_QtdeJaFaturada = 0
    
    Dim curQTDEREAL As Currency
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       ITENS.* " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADORDFATH CABEC" & vbCrLf
    sSql = sSql & "      ,SGI_CADORDFATI ITENS" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       CABEC.SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And CABEC.SGI_CODPED    = " & lngCODPEDIDO & vbCrLf
    sSql = sSql & "   And ITENS.SGI_FILIAL    = CABEC.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And ITENS.SGI_CODORD    = CABEC.SGI_CODORD " & vbCrLf
    sSql = sSql & "   And ITENS.SGI_IDPRODUTO = " & lngIDPRODUTO
    
    BREC6.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC6.EOF()
    
        curQTDEREAL = 0
        If Not IsNull(BREC6!SGI_QTDFAT) Then curQTDEREAL = BREC6!SGI_QTDFAT
        
        Calc_QtdeJaFaturada = Calc_QtdeJaFaturada + curQTDEREAL
        BREC6.MoveNext
    Loop
    BREC6.Close
    
End Function


Private Sub AbilitNF(strCODORD As String)

    fraNF.Visible = False
    
    If Len(Trim(strCODORD)) = 0 Then Exit Sub
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_CODCONF " & vbCrLf
    sSql = sSql & "      ,SGI_CODFATURA " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADORDCONFH" & strNOMFILIAL & vbCrLf
    sSql = sSql & "  Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODORD = " & Trim(strCODORD)
    
    BREC11.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC11.EOF() Then
       fraNF.Visible = True
       lblCODCONFIRMACAO.Caption = BREC11!SGI_CODCONF
       lblCODFATURA.Caption = BREC11!SGI_CODFATURA
    End If
    BREC11.Close
    
End Sub

Private Sub DestroiObjeto()
    Set objBLBFunc = Nothing
    Set objCADORDFAT = Nothing
    Set objPESQPADRAO = Nothing
End Sub

Private Function PegaLib10Porc() As Long

    PegaLib10Porc = 0
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_PERMFAT10POR" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_USUARIO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL       = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO       = " & lngCodUsuario
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PegaLib10Porc = BREC!SGI_PERMFAT10POR
    BREC.Close
    
    
End Function


Private Function PegaLibSN() As Long

    PegaLibSN = 0
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_PERMFATROTDIFSN" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_USUARIO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL       = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO       = " & lngCodUsuario
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PegaLibSN = BREC!SGI_PERMFATROTDIFSN
    BREC.Close
    
    
End Function


Private Sub InitGridLogPed()

    With grdLogPed
    
       .Cols = conColumnsIn_SonLogPed
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonLogPed_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       
       .Cell(flexcpData, 0, conCOL_SonLogPed_Data) = ""
       .ColDataType(conCOL_SonLogPed_Data) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_SonLogPed_Hora) = ""
       .ColDataType(conCOL_SonLogPed_Hora) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonLogPed_CodUsuario) = ""
       .ColDataType(conCOL_SonLogPed_CodUsuario) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonLogPed_Usuario) = ""
       .ColDataType(conCOL_SonLogPed_Usuario) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonLogPed_CodAcao) = ""
       .ColDataType(conCOL_SonLogPed_CodAcao) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonLogPed_Acao) = ""
       .ColDataType(conCOL_SonLogPed_Acao) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonLogPed_Tipo) = ""
       .ColDataType(conCOL_SonLogPed_Tipo) = flexDTString
       
       .ColWidth(conCOL_SonLogPed_Data) = 1000
       .ColWidth(conCOL_SonLogPed_Hora) = 1000
       .ColWidth(conCOL_SonLogPed_CodUsuario) = 0
       .ColWidth(conCOL_SonLogPed_Usuario) = 1000
       .ColWidth(conCOL_SonLogPed_CodAcao) = 0
       .ColWidth(conCOL_SonLogPed_Acao) = 3000
       .ColWidth(conCOL_SonLogPed_Tipo) = 0
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
    End With
    
End Sub


Private Sub PopLogPedidos()

    Dim I           As Integer
    Dim strNNOMUSU  As String
    Dim arrLOG      As Variant
    
    arrLOG = objCADORDFAT.LOG
    
    If IsArray(objCADORDFAT.LOG) Then
    
        With grdLogPed
        
            For I = 1 To UBound(arrLOG)
                
                strNNOMUSU = objBLBFunc.PegaUsuario(CLng(arrLOG(I, 3)), Linha, FILIAL)
                
                .AddItem arrLOG(I, 1) & vbTab & _
                         arrLOG(I, 2) & vbTab & _
                         arrLOG(I, 3) & vbTab & _
                         objBLBFunc.PegaUsuario(CLng(arrLOG(I, 3)), Linha, FILIAL) & vbTab & _
                         arrLOG(I, 4) & vbTab & _
                         "" & vbTab & _
                         ""
                         
                .Cell(flexcpText, (.Rows - 1), conCOL_SonLogPed_Acao) = DescAcao(.Cell(flexcpText, (.Rows - 1), conCOL_SonLogPed_CodAcao))
                         
            Next I
        
        End With
    
    End If

End Sub



Private Function DescAcao(strACAO As String) As String

    DescAcao = ""
    
    If strACAO = "I" Then DescAcao = "Incluso"
    If strACAO = "A" Then DescAcao = "Alterado"
        
End Function

Private Function Verifica_Credito(strCODCLIE As String) As String
    
On Error GoTo Err_Verifica_Credito

    If Len(Trim(strCODCLIE)) = 0 Then Exit Function
    
    If BREC.State = 1 Then BREC.Close
    
    Dim curLIMCRED     As Currency
    Dim strBLOQPED     As String
    Dim curVLTITABERTO As Currency
    Dim intRESP        As Integer
    
    curLIMCRED = 0
    curVLTITABERTO = 0
    strBLOQPED = "N"
    
    ''Verifica_Credito = "B"
    Verifica_Credito = "S"
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_VLLIMCRED    " & vbCrLf
    sSql = sSql & "      ,SGI_CREDSBLQPED  " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & strCODCLIE
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
       curLIMCRED = BREC!SGI_VLLIMCRED
       strBLOQPED = BREC!SGI_CREDSBLQPED
    End If
    BREC.Close
    
    '' -----------------------------------------------------
    sSql = "Select " & vbCrLf
    sSql = sSql & "        SUM(ITENS.SGI_VLDOC) As SGI_SALDO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "        SGI_CONTASHARC CABEC " & vbCrLf
    sSql = sSql & "       ,SGI_CONTASIARC ITENS " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       CABEC.SGI_FILIAL = " & FILIAL
    sSql = sSql & "   And CABEC.SGI_CODCLI = " & strCODCLIE
    sSql = sSql & "   And ITENS.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And ITENS.SGI_CODIGO = CABEC.SGI_CODIGO " & vbCrLf
    sSql = sSql & "   And ITENS.SGI_VLPAGO IS NULL "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
       If Not IsNull(BREC!SGI_SALDO) Then curVLTITABERTO = BREC!SGI_SALDO
    End If
    BREC.Close
    
    '' Regras para para crédito
    ''If (curVLTITABERTO >= curLIMCRED) And curVLTITABERTO > 0 And curLIMCRED > 0 Then
    ''   If strBLOQPED = "N" Then
    ''      intRESP = MsgBox("Atenção este Cliente está com o limite estourado, bloqueia para análize ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
    ''      If intRESP = 6 Then Verifica_Credito = "B"
    ''      If intRESP = 7 Then Verifica_Credito = "L"
    ''   End If
    ''   If strBLOQPED = "S" Then
    ''      MsgBox "Atenção este Cliente está com o limite estourado, pedido será bloqueado para análize", vbOKOnly + vbInformation, "Aviso"
    ''      Verifica_Credito = "B"
    ''   End If
    ''End If
    '' ------------------------
    
    If strBLOQPED = "S" Then
       MsgBox "Atenção este Cliente está bloqueado por motivos financeiros !!!", vbOKOnly + vbInformation, "Aviso"
       Verifica_Credito = "N"
    End If
    
    ''Verifica_Credito = "B"
    
    Exit Function
    
Err_Verifica_Credito:
    
    If BREC.State = 1 Then BREC.Close
    
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Altera()", Me.Name, "Altera()", strCAMARQERRO)
    
End Function

