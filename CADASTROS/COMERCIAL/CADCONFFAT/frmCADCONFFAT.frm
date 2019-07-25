VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCADCONFFAT 
   Caption         =   "Cadastro de Confirmação de Faturamento"
   ClientHeight    =   9960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13875
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9960
   ScaleWidth      =   13875
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
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
      Left            =   6360
      TabIndex        =   60
      Top             =   7440
      Width           =   7455
      Begin VB.TextBox txtCODFATURA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         TabIndex        =   2
         Text            =   "txtCODFATURA"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   61
         Top             =   240
         Width           =   1515
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
      Height          =   2655
      Left            =   0
      TabIndex        =   56
      Top             =   4800
      Width           =   13815
      Begin VSFlex8LCtl.VSFlexGrid grdITENSPEDIDO 
         Height          =   1935
         Left            =   120
         TabIndex        =   57
         Top             =   240
         Width           =   13575
         _cx             =   23945
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
         Index           =   15
         Left            =   9960
         TabIndex        =   59
         Top             =   2280
         Width           =   1830
      End
      Begin VB.Label lblTOTALFAT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
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
         Height          =   255
         Left            =   11880
         TabIndex        =   58
         Top             =   2280
         Width           =   1695
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
      Height          =   2415
      Left            =   0
      TabIndex        =   37
      Top             =   7440
      Width           =   6255
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
         TabIndex        =   39
         Text            =   "txtOutrDesp"
         Top             =   960
         Width           =   1335
      End
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
         TabIndex        =   38
         Text            =   "txtFRETE"
         Top             =   960
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
         TabIndex        =   55
         Top             =   285
         Width           =   1635
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
         TabIndex        =   54
         Top             =   240
         Width           =   1335
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
         TabIndex        =   53
         Top             =   1005
         Width           =   1425
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
         TabIndex        =   52
         Top             =   990
         Width           =   450
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
         TabIndex        =   51
         Top             =   1350
         Width           =   1020
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
         TabIndex        =   50
         Top             =   1320
         Width           =   1335
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
         TabIndex        =   49
         Top             =   285
         Width           =   840
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
         TabIndex        =   48
         Top             =   645
         Width           =   1230
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
         TabIndex        =   47
         Top             =   600
         Width           =   1335
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
         TabIndex        =   46
         Top             =   240
         Width           =   1335
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
         TabIndex        =   45
         Top             =   1680
         Width           =   1020
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
         TabIndex        =   44
         Top             =   1680
         Width           =   1050
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
         TabIndex        =   43
         Top             =   2055
         Width           =   1320
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
         TabIndex        =   42
         Top             =   2040
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
         TabIndex        =   41
         Top             =   1680
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
         TabIndex        =   40
         Top             =   1680
         Width           =   1335
      End
   End
   Begin VB.Frame Frame7 
      Height          =   975
      Left            =   0
      TabIndex        =   34
      Top             =   3840
      Width           =   13815
      Begin VB.TextBox txtOBS 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Text            =   "frmCADCONFFAT.frx":0000
         Top             =   240
         Width           =   12375
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
         Index           =   14
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   0
      TabIndex        =   29
      Top             =   3240
      Width           =   13815
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
         Index           =   12
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   2085
      End
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
         Left            =   7920
         TabIndex        =   32
         Top             =   240
         Width           =   930
      End
      Begin VB.Label lblCONDPGTO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblCONDPGTO"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2280
         TabIndex        =   31
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label lblTRANSPORTE 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblTRANSPORTE"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9000
         TabIndex        =   30
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1695
      Left            =   0
      TabIndex        =   12
      Top             =   1560
      Width           =   13815
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
         Index           =   10
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   600
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
         Left            =   7320
         TabIndex        =   27
         Top             =   240
         Width           =   915
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
         Left            =   10560
         TabIndex        =   26
         Top             =   240
         Width           =   1230
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
         TabIndex        =   25
         Top             =   600
         Width           =   825
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
         TabIndex        =   24
         Top             =   960
         Width           =   600
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
         Left            =   10560
         TabIndex        =   23
         Top             =   960
         Width           =   600
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
         TabIndex        =   22
         Top             =   1320
         Width           =   765
      End
      Begin VB.Label lblCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblCliente"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   21
         Top             =   240
         Width           =   5895
      End
      Begin VB.Label lblCNPJCPF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblCNPJCPF"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8640
         TabIndex        =   20
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblINSCREST 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblINSCREST"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   12000
         TabIndex        =   19
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblENDERECO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblENDERECO"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   18
         Top             =   600
         Width           =   5895
      End
      Begin VB.Label lblCIDADE 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblCIDADE"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Top             =   960
         Width           =   3855
      End
      Begin VB.Label lblESTADO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblESTADO"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   12000
         TabIndex        =   16
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblTELEFONE 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblTELEFONE"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   15
         Top             =   1320
         Width           =   5895
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
         Left            =   5520
         TabIndex        =   14
         Top             =   960
         Width           =   510
      End
      Begin VB.Label lblBairro 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblBairro"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6480
         TabIndex        =   13
         Top             =   960
         Width           =   3855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   13815
      Begin MSMask.MaskEdBox mskDataConf 
         Height          =   255
         Left            =   5280
         TabIndex        =   0
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtCodOrdem 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9720
         TabIndex        =   1
         Text            =   "txtCodOrdem"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblCODPED 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblCODPED"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   12360
         TabIndex        =   63
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pedido No."
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
         Left            =   11280
         TabIndex        =   62
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data da Confirmação"
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
         Left            =   3360
         TabIndex        =   11
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label lblCODIGO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblCODIGO"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2160
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
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
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   270
         Width           =   1980
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código da Ordem de Faturamento"
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
         Left            =   6720
         TabIndex        =   8
         Top             =   240
         Width           =   2835
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   13815
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
         Picture         =   "frmCADCONFFAT.frx":0009
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "frmCADCONFFAT.frx":010B
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "frmCADCONFFAT.frx":020D
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADCONFFAT"
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
Public intFILIALPED     As Integer
Public boolSomenteCons  As Boolean

Dim arrITENSFAT         As Variant
Dim arrOPS_RELACIONADA  As Variant
Dim objBLBFunc          As Object
Dim objCADCONFFAT       As Object
Dim objPESQPADRAO       As Object
Dim lngOPFECHADA        As Long

Dim strSGI_CADORDFATH   As String
Dim strSGI_CADPEDVENDH  As String
Dim strSGI_CADORDFATI   As String
Dim strSGI_ORDEMPROD    As String
Dim strNOMFILIAL        As String
Dim strCAPFILIAL        As String
Dim strfrmCADCONFFAT    As String
Dim strNOMTABELA        As String

Const conCOL_Produto_IDProduto                As Integer = 0
Const conCOL_Produto_CodProduto               As Integer = 1
Const conCOL_Produto_DescProduto              As Integer = 2
Const conCOL_Produto_QtdeFaturada             As Integer = 3
Const conCOL_Produto_PorcIPI                  As Integer = 4
Const conCOL_Produto_VlUnit                   As Integer = 5
Const conCOL_Produto_VLFaturado               As Integer = 6
Const conCOL_Produto_VLIPI                    As Integer = 7
Const conCOL_Produto_Action2Do                As Integer = 8
Const conCOL_Produto_OrdFab                   As Integer = 9
Const conCOL_Produto_QtdTotFat                As Integer = 10
Const conCOL_Produto_QtdOrdem                 As Integer = 11
Const conCOL_Produto_SaldoOrdem               As Integer = 12
Const conCOL_Produto_Status                   As Integer = 13
Const conCOL_Produto_CodForn                  As Integer = 14
Const conCOL_Produto_QtdeOrd                  As Integer = 15

Const conCOL_Produto_FormatString             As String = "=Código|Produto|Descrição|Qtd. Faturada|% IPI|Vl. Unitário|Vl.Total|Vl. do IPI|Action2Do|OrdFab|QtdTotfat|QtdOrdem|SaldoOrd|Status|Cod.Forn|Qtde.Ordem"
Const conColumnsIn_Produto                    As Integer = 16

Private Sub cmdAltera_Click()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    
    Me.Caption = "Cadastro de Confirmação de Faturamento - [ ALTERAÇÃO ]"
    
    Frame2.Enabled = False
    Frame3.Enabled = True
    Frame4.Enabled = True
    Frame8.Enabled = False
    Frame5.Enabled = True
    txtOBS.Locked = False
    
    cTipOper = "A"

End Sub

Private Sub CmdSalva_Click()

On Error GoTo err_grava
    
    If Not VerifCampos Then Exit Sub
    
    Dim I           As Integer
    Dim intRESP     As Integer
    Dim lngCodLog   As Long
    Dim curQTDEFAT  As Currency
    Dim lngqtdeOPS  As Long

    If cTipOper = "I" Then objCADCONFFAT.CODCONF = objBLBFunc.Gera_Codigo(strfrmCADCONFFAT, FILIAL, Linha) & Year(Now)

    objCADCONFFAT.CODORD = CLng(txtCodOrdem.Text)
    objCADCONFFAT.DATACONF = CDate(mskDataConf.Text)
    objCADCONFFAT.OBS = Trim(Replace(txtOBS.Text, "'", ""))
    objCADCONFFAT.CODFAT = CLng(txtCODFATURA.Text)
    objCADCONFFAT.CODPED = CLng(Trim(Replace(lblCODPED.Caption, "/", "")))

    objCADCONFFAT.BASEICMS = 0
    objCADCONFFAT.ALIQICMS = 0
    objCADCONFFAT.VALOICMS = 0
    objCADCONFFAT.OUTRASDESP = 0
    objCADCONFFAT.FRETE = 0
    objCADCONFFAT.VALORIPI = 0
    objCADCONFFAT.PORCDESCTO = 0
    objCADCONFFAT.VALORDESCT = 0
    objCADCONFFAT.VLTOTALFAT = 0
    
    If Len(Trim(lblBASICMS.Caption)) > 0 Then objCADCONFFAT.BASEICMS = CCur(lblBASICMS.Caption)
    If Len(Trim(lblALIQICMS.Caption)) > 0 Then objCADCONFFAT.ALIQICMS = CCur(lblALIQICMS.Caption)
    If Len(Trim(lblVLICMS.Caption)) > 0 Then objCADCONFFAT.VALOICMS = CCur(lblVLICMS.Caption)
    If Len(Trim(txtOutrDesp.Text)) > 0 Then objCADCONFFAT.OUTRASDESP = CCur(txtOutrDesp.Text)
    If Len(Trim(txtFRETE.Text)) > 0 Then objCADCONFFAT.FRETE = CCur(txtFRETE.Text)
    If Len(Trim(lblVLIPI.Caption)) > 0 Then objCADCONFFAT.VALORIPI = CCur(lblVLIPI.Caption)
    If Len(Trim(lblPDESCTOTAL.Caption)) > 0 Then objCADCONFFAT.PORCDESCTO = CCur(lblPDESCTOTAL.Caption)
    If Len(Trim(lblVLDESCTOTOT.Caption)) > 0 Then objCADCONFFAT.VALORDESCT = CCur(lblVLDESCTOTOT.Caption)
    If Len(Trim(lblVLTOTAL.Caption)) > 0 Then objCADCONFFAT.VLTOTALFAT = CCur(lblVLTOTAL.Caption)
    
    '' Itens do Faturamento
    objCADCONFFAT.ITENSFAT = Empty
    curQTDEFAT = 0
    lngOPFECHADA = 0
    Dim lngQTDFAT As Long
    With grdITENSPEDIDO
        If (.Rows - 1) > 0 Then
            ReDim arrITENSFAT(1 To (.Rows - 1), 1 To 15) As String
            For I = 1 To (.Rows - 1)
            
                arrITENSFAT(I, 1) = .Cell(flexcpText, I, conCOL_Produto_IDProduto)
                arrITENSFAT(I, 2) = .Cell(flexcpText, I, conCOL_Produto_CodProduto)
                arrITENSFAT(I, 3) = .Cell(flexcpText, I, conCOL_Produto_QtdeFaturada)
                arrITENSFAT(I, 4) = .Cell(flexcpText, I, conCOL_Produto_PorcIPI)
                arrITENSFAT(I, 5) = .Cell(flexcpText, I, conCOL_Produto_VlUnit)
                arrITENSFAT(I, 6) = .Cell(flexcpText, I, conCOL_Produto_VLFaturado)
                arrITENSFAT(I, 7) = .Cell(flexcpText, I, conCOL_Produto_VLIPI)
                arrITENSFAT(I, 8) = .Cell(flexcpText, I, conCOL_Produto_Action2Do)
                arrITENSFAT(I, 9) = .Cell(flexcpText, I, conCOL_Produto_OrdFab)
                arrITENSFAT(I, 10) = .Cell(flexcpText, I, conCOL_Produto_QtdTotFat)
                arrITENSFAT(I, 11) = .Cell(flexcpText, I, conCOL_Produto_QtdOrdem)
                arrITENSFAT(I, 12) = .Cell(flexcpText, I, conCOL_Produto_SaldoOrdem)
                arrITENSFAT(I, 13) = .Cell(flexcpText, I, conCOL_Produto_Status)
                
                If Len(Trim(.Cell(flexcpText, I, conCOL_Produto_CodForn))) > 0 Then
                    arrITENSFAT(I, 14) = .Cell(flexcpText, I, conCOL_Produto_CodForn)
                Else
                    arrITENSFAT(I, 14) = "Null"
                End If
            
                If Len(Trim(.Cell(flexcpText, I, conCOL_Produto_QtdeFaturada))) > 0 Then
                    curQTDEFAT = (curQTDEFAT + CCur(.Cell(flexcpText, I, conCOL_Produto_QtdeFaturada)))
                    lngOPFECHADA = (lngOPFECHADA + OPFECHADA(Trim(Str(arrITENSFAT(I, 9))), Trim(Str(arrITENSFAT(I, 1))), CCur(.Cell(flexcpText, I, conCOL_Produto_QtdeFaturada))))
                End If
                
            Next I
        End If
        objCADCONFFAT.ITENSFAT = arrITENSFAT
        
    End With
    objCADCONFFAT.QTDEFATCONF = curQTDEFAT
    
    '' Pegando os dados para pode fechar o pedido
    ''Call CalcSaldoFechado(objCADCONFFAT.CODCONF, curQTDEFAT)
    
    Dim lngSALDOPEDIDO As Long
    lngSALDOPEDIDO = SaldoPedido(lblCODPED.Caption, curQTDEFAT)
    If lngSALDOPEDIDO <= 0 Then
       objCADCONFFAT.STATUSPED = "'F'"
    ElseIf lngSALDOPEDIDO > 0 Then
       objCADCONFFAT.STATUSPED = "'P'"
    End If
    
    '' =========================
    '' Pegando OP's Relacionadas
    arrOPS_RELACIONADA = Empty
    If objCADCONFFAT.STATUSPED = "'F'" Then
        If IsArray(arrITENSFAT) Then
            For I = 1 To UBound(arrITENSFAT)
                
                lngqtdeOPS = 0
                
                sSql = ""
                
                sSql = "Select" & vbCrLf
                sSql = sSql & "       Count(*) as Qtde" & vbCrLf
                sSql = sSql & "  From" & vbCrLf
                sSql = sSql & "       SGI_ORDEMPROD" & strNOMFILIAL & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       SGI_FILIAL    =  " & FILIAL & vbCrLf
                sSql = sSql & "   And SGI_CODPED    =  " & objCADCONFFAT.CODPED & vbCrLf
                sSql = sSql & "   And SGI_IDPRODUTO =  " & arrITENSFAT(I, 1) & vbCrLf
                sSql = sSql & "   And SGI_CODIGO    <> " & arrITENSFAT(I, 9) & vbCrLf
                sSql = sSql & "   And (SGI_STATUS   = 0 Or SGI_STATUS    = 1)"
                
                BREC.Open sSql, adoBanco_Dados, adOpenDynamic
                If Not BREC.EOF() Then lngqtdeOPS = BREC!Qtde
                BREC.Close
            
                If lngqtdeOPS > 0 Then
                    ReDim arrOPS_RELACIONADA(1 To lngqtdeOPS, 1 To 1) As String
                
                    lngqtdeOPS = 1
                    sSql = ""
                    
                    sSql = "Select" & vbCrLf
                    sSql = sSql & "       *" & vbCrLf
                    sSql = sSql & "  From" & vbCrLf
                    sSql = sSql & "       SGI_ORDEMPROD" & strNOMFILIAL & vbCrLf
                    sSql = sSql & " Where " & vbCrLf
                    sSql = sSql & "       SGI_FILIAL    =  " & FILIAL & vbCrLf
                    sSql = sSql & "   And SGI_CODPED    =  " & objCADCONFFAT.CODPED & vbCrLf
                    sSql = sSql & "   And SGI_IDPRODUTO =  " & arrITENSFAT(I, 1) & vbCrLf
                    sSql = sSql & "   And SGI_CODIGO    <> " & arrITENSFAT(I, 9) & vbCrLf
                    sSql = sSql & "   And (SGI_STATUS   = 0 Or SGI_STATUS    = 1)"
                
                    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
                    Do While Not BREC.EOF()
                        arrOPS_RELACIONADA(lngqtdeOPS, 1) = Trim(Str(BREC!SGI_CODIGO))
                    
                        lngqtdeOPS = (lngqtdeOPS + 1)
                        BREC.MoveNext
                    Loop
                    BREC.Close
                
                End If
            Next I
        End If
    End If
    objCADCONFFAT.arrARROPSELECIONADA = arrOPS_RELACIONADA
    '' =========================
    
    
    '' Gravando as Informações no banco
    If objCADCONFFAT.GRAVA(cTipOper, intFILIALPED) = False Then Exit Sub
    
    If objCADCONFFAT.Atualiza(Trim(Replace(objCADCONFFAT.STATUSPED, "'", "")), objCADCONFFAT.CODCONF, FILIAL, "frmCADPEDVENDA", intFILIALPED) = False Then Exit Sub
    
    
    '' Gerand Log de Sistema
    lngCodLog = objBLBFunc.Gera_Codigo("SGI_LOGMODULO" & Trim(strNOMFILIAL), FILIAL, Linha)
    Call objBLBFunc.GravaLogModulo(FILIAL, lngCodLog, strfrmCADCONFFAT, cTipOper, lngCodUsuario, Str(objCADCONFFAT.CODCONF), Linha)
    
    MsgBox "A Confirmação da ordem de faturamento nº ( " & Trim(Str(objCADCONFFAT.CODCONF)) & " ) foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
    
    If cTipOper = "I" Then
       intRESP = MsgBox("Deseja gerar outra confirmação da ordem de Faturamento ?", vbYesNo + vbQuestion, "Aviso")
       
       If intRESP = 6 Then
          Call Inclui
       Else
          Call DestroiObjeto
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
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADCONFFAT = CreateObject("CADCONFFAT.clsCADCONFFAT")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
   objCADCONFFAT.FILIAL = FILIAL
   
   If adoBanco_Dados.State = 0 Then Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)
   If adoBanco_Dados.State = 0 Then
      MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
      Exit Sub
   End If
   
    strSGI_CADORDFATH = "SGI_CADORDFATH"
    strSGI_CADPEDVENDH = "SGI_CADPEDVENDH"
    strSGI_CADORDFATI = "SGI_CADORDFATI"
    strSGI_ORDEMPROD = "SGI_ORDEMPROD"
    strfrmCADCONFFAT = Trim(Me.Name)
    
    If intFILIALPED = 0 Then
       strNOMFILIAL = ""
       strCAPFILIAL = " / NOVALATA"
    ElseIf intFILIALPED = 1 Then
       strNOMFILIAL = "_STEEL"
       strCAPFILIAL = " / STEEL ROL"
    End If
   
    strSGI_CADORDFATH = Trim(strSGI_CADORDFATH & strNOMFILIAL)
    strSGI_CADPEDVENDH = Trim(strSGI_CADPEDVENDH & strNOMFILIAL)
    strSGI_CADORDFATI = Trim(strSGI_CADORDFATI & strNOMFILIAL)
    strSGI_ORDEMPROD = Trim(strSGI_ORDEMPROD & strNOMFILIAL)
    strfrmCADCONFFAT = Trim(strfrmCADCONFFAT & strNOMFILIAL)
    
    objCADCONFFAT.CODUSUARIO = lngCodUsuario
    
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    lblCODIGO.Caption = ""
    
    Frame2.Enabled = True
    Frame5.Enabled = True
    Frame8.Enabled = False
    txtOBS.Locked = True
    
    Me.Caption = "Cadastro de Confirmação de Faturamento - [ INCLUSÃO ]" & strCAPFILIAL
    
    objBLBFunc.LimpaCampos frmCADCONFFAT
    
    mskDataConf.Text = Format(Now, "DD/MM/YYYY")
    Call LimpaCamposLabel
    
    Call ConfGridProdutos
    
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
       
       .Cell(flexcpData, 0, conCOL_Produto_QtdeFaturada) = ""
       .ColDataType(conCOL_Produto_QtdeFaturada) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Produto_PorcIPI) = ""
       .ColDataType(conCOL_Produto_PorcIPI) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_Produto_VlUnit) = ""
       .ColDataType(conCOL_Produto_VlUnit) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_Produto_VLFaturado) = ""
       .ColDataType(conCOL_Produto_VLFaturado) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_Produto_VLIPI) = ""
       .ColDataType(conCOL_Produto_VLIPI) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_Produto_Action2Do) = ""
       .ColDataType(conCOL_Produto_Action2Do) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Produto_OrdFab) = ""
       .ColDataType(conCOL_Produto_OrdFab) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Produto_QtdTotFat) = ""
       .ColDataType(conCOL_Produto_QtdTotFat) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_Produto_QtdOrdem) = ""
       .ColDataType(conCOL_Produto_QtdOrdem) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_Produto_SaldoOrdem) = ""
       .ColDataType(conCOL_Produto_SaldoOrdem) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_Produto_Status) = ""
       .ColDataType(conCOL_Produto_Status) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_Produto_CodForn) = ""
       .ColDataType(conCOL_Produto_CodForn) = flexDTLong
       
       .ColWidth(conCOL_Produto_IDProduto) = 0
       .ColWidth(conCOL_Produto_CodProduto) = 1200
       .ColWidth(conCOL_Produto_DescProduto) = 3500
       .ColWidth(conCOL_Produto_QtdeFaturada) = 1200
       .ColWidth(conCOL_Produto_PorcIPI) = 600
       .ColWidth(conCOL_Produto_VlUnit) = 1000
       .ColWidth(conCOL_Produto_VLFaturado) = 1200
       .ColWidth(conCOL_Produto_VLIPI) = 1200
       .ColWidth(conCOL_Produto_Action2Do) = 0
       .ColWidth(conCOL_Produto_OrdFab) = 1000
       
       .ColWidth(conCOL_Produto_QtdTotFat) = 1000
       .ColWidth(conCOL_Produto_QtdOrdem) = 1000
       .ColWidth(conCOL_Produto_SaldoOrdem) = 1000
       .ColWidth(conCOL_Produto_Status) = 1000
       .ColWidth(conCOL_Produto_CodForn) = 0
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
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
    lblCODPED.Caption = ""
    
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

Private Sub Form_Unload(Cancel As Integer)
    Call DestroiObjeto
End Sub





Private Sub mskDataConf_GotFocus()
    objBLBFunc.SelecionaCampos mskDataConf.Name, Me
End Sub

Private Sub txtCodOrdem_GotFocus()
    objBLBFunc.SelecionaCampos txtCodOrdem.Name, frmCADCONFFAT
End Sub

Private Sub txtCodOrdem_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCodOrdem.Text
End Sub

Private Sub txtCodOrdem_Validate(Cancel As Boolean)

    If Len(Trim(txtCodOrdem.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCodOrdem.Text) Then
       MsgBox "Somente é Permitido Numeros !!!", vbOKOnly + vbExclamation, "Aviso"
       txtCodOrdem.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Cancel = PegaDadosDaOrdem(Trim(txtCodOrdem.Text))

End Sub


Private Function PegaDadosDaOrdem(strCodOrdem As String) As Boolean

    PegaDadosDaOrdem = False
    
    Call LimpaCamposLabel
    Call ConfGridProdutos
    
    Dim curQTDORDEM As Currency
    Dim curQTDJAFAT As Currency
    Dim curQTDSALDO As Currency
    txtOBS.Text = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       ORDF.* " & vbCrLf
    sSql = sSql & "      ,PED.*" & vbCrLf
    sSql = sSql & "      ,PED.SGI_CODIGO AS SGI_CODPED " & vbCrLf
    sSql = sSql & "      ,CLI.* " & vbCrLf
    sSql = sSql & "      ,PGT.SGI_DESCRICAO AS SGI_DESCPGTO " & vbCrLf
    sSql = sSql & "      ,TRA.SGI_DESCRICAO AS SGI_DESCTRANSP " & vbCrLf
    sSql = sSql & "      ,ORDF.SGI_OBS      As SGI_OBSCONF " & vbCrLf
    sSql = sSql & "      ,ORDF.SGI_FILIALPED " & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & strSGI_CADORDFATH & " ORDF" & vbCrLf
    sSql = sSql & "      ," & strSGI_CADPEDVENDH & " PED" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE  CLI" & vbCrLf
    sSql = sSql & "      ,SGI_CADCONDPGTO PGT" & vbCrLf
    sSql = sSql & "      ,SGI_CADTRANSP   TRA" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    
    sSql = sSql & "       ORDF.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And ORDF.SGI_CODORD = " & Trim(strCodOrdem) & vbCrLf
    sSql = sSql & "   And ORDF.SGI_STATUS = 0 " & vbCrLf
    sSql = sSql & "   And PED.SGI_FILIAL  = ORDF.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PED.SGI_CODIGO  = ORDF.SGI_CODPED " & vbCrLf
    sSql = sSql & "   And CLI.SGI_FILIAL  = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And CLI.SGI_CODIGO  = PED.SGI_CODCLI " & vbCrLf
    sSql = sSql & "   And PGT.SGI_FILIAL  = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PGT.SGI_CODIGO  = PED.SGI_CODCONDPGT " & vbCrLf
    sSql = sSql & "   And TRA.SGI_FILIAL  = PED.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And TRA.SGI_CODIGO  = PED.SGI_CODTRANSP " & vbCrLf

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
    
        objCADCONFFAT.CODCLIE = BREC!SGI_CODCLI
        lblCliente.Caption = Trim(BREC!SGI_RAZAOSOC)
        lblCNPJCPF.Caption = objBLBFunc.FormataCnpj(BREC!SGI_CPFCNPJ)
        lblINSCREST.Caption = BREC!SGI_RGCGC
        lblENDERECO.Caption = BREC!SGI_ENDNORM
        lblCIDADE.Caption = BREC!SGI_CIDNORM
        lblBairro.Caption = BREC!SGI_BAINROM
        lblCONDPGTO.Caption = BREC!SGI_DESCPGTO
        lblTRANSPORTE.Caption = BREC!SGI_DESCTRANSP
        lblCODPED.Caption = Format(BREC!SGI_CODPED, "#/####")
        
        If Not IsNull(BREC!SGI_OBSCONF) Then txtOBS.Text = BREC!SGI_OBSCONF
        
        lblTELEFONE.Caption = PegaTelefonesCli(Trim(Str(BREC!SGI_CODCLI)))
        lblESTADO.Caption = PegaEstado(Trim(Str(BREC!SGI_ESTNORM)))
        
        If Not IsNull(BREC!SGI_BASEFAT) Then lblBASICMS.Caption = Format(BREC!SGI_BASEFAT, "#,##0.00")
        If Not IsNull(BREC!SGI_ALIQICMS) Then lblALIQICMS.Caption = Format(BREC!SGI_ALIQICMS, "#,##0.00")
        If Not IsNull(BREC!SGI_VALOICMS) Then lblVLICMS.Caption = Format(BREC!SGI_VALOICMS, "#,##0.00")
        If Not IsNull(BREC!SGI_OUTRDESP) Then txtOutrDesp.Text = Format(BREC!SGI_OUTRDESP, "#,##0.00")
        If Not IsNull(BREC!SGI_FRETE) Then txtFRETE.Text = Format(BREC!SGI_FRETE, "#,##0.00")
        If Not IsNull(BREC!SGI_TOTALIPI) Then lblVLIPI.Caption = Format(BREC!SGI_TOTALIPI, "#,##0.00")
        If Not IsNull(BREC!SGI_PERCDESC) Then lblPDESCTOTAL.Caption = Format(BREC!SGI_PERCDESC, "#,##0.00")
        If Not IsNull(BREC!SGI_VALODESC) Then lblVLDESCTOTOT.Caption = Format(BREC!SGI_VALODESC, "#,##0.00")
        If Not IsNull(BREC!SGI_TOTALFAT) Then lblVLTOTAL.Caption = Format(BREC!SGI_TOTALFAT, "#,##0.00")
        
        objCADCONFFAT.FILIALEMP = BREC!SGI_FILIALPED
        
        '' Pega Itens do Pedido
        sSql = "Select " & vbCrLf
        sSql = sSql & "       ITEN.* " & vbCrLf
        sSql = sSql & "      ,PROD.* " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       " & strSGI_CADORDFATI & " ITEN " & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO PROD " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       ITEN.SGI_FILIAL    = " & FILIAL & vbCrLf
        sSql = sSql & "   And ITEN.SGI_CODORD    = " & BREC!SGI_CODORD & vbCrLf
        sSql = sSql & "   And PROD.SGI_FILIAL    = ITEN.SGI_FILIAL " & vbCrLf
        sSql = sSql & "   And PROD.SGI_IDPRODUTO = ITEN.SGI_IDPRODUTO "
        
        BREC5.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC5.EOF() Then
            Do While Not BREC5.EOF()
            
               If Not IsNull(BREC5!SGI_QTDFAT) Then
                    With grdITENSPEDIDO
                    
                         .AddItem BREC5!SGI_IDPRODUTO & vbTab & _
                                  Trim(BREC5!SGI_CODPRODUTO) & vbTab & _
                                  Trim(BREC5!SGI_DESCRICAO) & vbTab & _
                                  BREC5!SGI_QTDFAT & vbTab & _
                                  Format(BREC5!SGI_PORCIPI, "#,##0.00") & vbTab & _
                                  Format(BREC5!SGI_VLUNIT, "#,##0.00") & vbTab & _
                                  Format(BREC5!SGI_VLFATURADO, "#,##0.00") & vbTab & _
                                  Format(BREC5!SGI_VLDOPI, "#,##0.00") & vbTab & _
                                  dacEnumUpdateAction_Ignore & vbTab & _
                                  BREC5!SGI_CODORDFAB & vbTab & _
                                  0 & vbTab & _
                                  0 & vbTab & _
                                  0 & vbTab & _
                                  0 & vbTab & _
                                  BREC5!SGI_CODFORN
                         
                         .Cell(flexcpBackColor, (.Rows - 1), conCOL_Produto_QtdeFaturada) = &H80C0FF
                         
                         '' ----------------------------------------------------------
                         curQTDORDEM = PegaQtdRealOrdem(Str(BREC5!SGI_CODORDFAB))
                         curQTDJAFAT = (PegaQtdJaFaturada(Str(BREC5!SGI_CODORDFAB)) + BREC5!SGI_QTDFAT)
                         curQTDSALDO = (curQTDORDEM - curQTDJAFAT)
                         
                         .Cell(flexcpText, (.Rows - 1), conCOL_Produto_QtdOrdem) = curQTDORDEM
                         .Cell(flexcpText, (.Rows - 1), conCOL_Produto_QtdTotFat) = curQTDJAFAT
                         .Cell(flexcpText, (.Rows - 1), conCOL_Produto_SaldoOrdem) = curQTDSALDO
                         
                         .Cell(flexcpText, (.Rows - 1), conCOL_Produto_Status) = 0
                         
                         '' Se o saldo for maior que 0 esta fechada parcial Status 1
                         If curQTDSALDO > 0 Then .Cell(flexcpText, (.Rows - 1), conCOL_Produto_Status) = 1
                         
                         '' Se o saldo for menor que 0 esta fechada Total Status 2
                         If curQTDSALDO <= 0 Then .Cell(flexcpText, (.Rows - 1), conCOL_Produto_Status) = 2
                         '' ----------------------------------------------------------
                         
                         '' Esta Opção foi tirada por pedido do Sr. Anderson
                         ''If FechamentoOP(curQTDORDEM, curQTDJAFAT) = "F" Then
                         ''   .Cell(flexcpText, (.Rows - 1), conCOL_Produto_Status) = 2
                         ''Else
                         ''   .Cell(flexcpText, (.Rows - 1), conCOL_Produto_Status) = 0
                         ''
                         ''   '' Se o saldo for maior que 0 esta fechada parcial Status 1
                         ''   If curQTDSALDO > 0 Then .Cell(flexcpText, (.Rows - 1), conCOL_Produto_Status) = 1
                         ''
                         ''   '' Se o saldo for menor que 0 esta fechada Total Status 2
                         ''   If curQTDSALDO <= 0 Then .Cell(flexcpText, (.Rows - 1), conCOL_Produto_Status) = 2
                         ''   '' ----------------------------------------------------------
                        ''End If
                    End With
               End If
                
               BREC5.MoveNext
            Loop
        End If
        BREC5.Close
        
        Call CalcTotFatura
    
    Else
        MsgBox "Esta Ordem Não Existe !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCodOrdem.Text = ""
        PegaDadosDaOrdem = True
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


Private Function VerifCampos() As Boolean
    
    VerifCampos = False
    
    Dim I               As Integer
    Dim curQTD_FATURADA As Currency
    Dim curQTD_TOTFATUR As Currency
    
    If Not IsDate(mskDataConf.Text) Then
       MsgBox "Data Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
       mskDataConf.SetFocus
       Exit Function
    End If
    If Len(Trim(txtCodOrdem.Text)) = 0 Then
       MsgBox "O Código do Pedido Não foi Informado !!!", vbOKOnly + vbExclamation, "Aviso"
       txtCodOrdem.SetFocus
       Exit Function
    End If
    If Len(Trim(txtCODFATURA.Text)) = 0 Then
        MsgBox "O Numero da Fatura não foi informado !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCODFATURA.SetFocus
        Exit Function
    End If
    
    VerifCampos = True
    
End Function

Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    
    If boolSomenteCons = True Then
       cmdAltera.Enabled = False
    Else
       cmdAltera.Enabled = True
    End If
    
    Me.Caption = "Cadastro de Confirmação de Faturamento - [ CONSULTA ]" & strCAPFILIAL
    
    objBLBFunc.LimpaCampos frmCADCONFFAT

    Frame2.Enabled = False
    Frame3.Enabled = False
    Frame4.Enabled = False
    Frame8.Enabled = False
    Frame5.Enabled = False
    txtOBS.Locked = True
    
    objCADCONFFAT.CODCONF = iCodigo
    
    Call ConfGridProdutos
    Call LimpaCamposLabel
    
    Call CarregaCampos
    
End Sub

Private Sub CarregaCampos()

    If objCADCONFFAT.Carrega_Campos(intFILIALPED) = True Then
        
        lblCODIGO.Caption = objCADCONFFAT.CODCONF
        mskDataConf.Text = Format(objCADCONFFAT.DATACONF, "DD/MM/YYYY")
        txtCodOrdem.Text = objCADCONFFAT.CODORD
        txtOBS.Text = objCADCONFFAT.OBS
        txtCODFATURA.Text = objCADCONFFAT.CODFAT
        
        Call CarregaDadosCliente(txtCodOrdem.Text)
       
        If objCADCONFFAT.BASEICMS > 0 Then lblBASICMS.Caption = Format(objCADCONFFAT.BASEICMS, "#,##0.00")
        If objCADCONFFAT.ALIQICMS > 0 Then lblALIQICMS.Caption = Format(objCADCONFFAT.ALIQICMS, "#,##0.00")
        If objCADCONFFAT.VALOICMS > 0 Then lblVLICMS.Caption = Format(objCADCONFFAT.VALOICMS, "#,##0.00")
        If objCADCONFFAT.OUTRASDESP > 0 Then txtOutrDesp.Text = Format(objCADCONFFAT.OUTRASDESP, "#,##0.00")
        If objCADCONFFAT.FRETE > 0 Then txtFRETE.Text = Format(objCADCONFFAT.FRETE, "#,##0.00")
        If objCADCONFFAT.VALORIPI > 0 Then lblVLIPI.Caption = Format(objCADCONFFAT.VALORIPI, "#,##0.00")
        If objCADCONFFAT.PORCDESCTO > 0 Then lblPDESCTOTAL.Caption = Format(objCADCONFFAT.PORCDESCTO, "#,##0.00")
        If objCADCONFFAT.VALORDESCT > 0 Then lblVLDESCTOTOT.Caption = Format(objCADCONFFAT.VALORDESCT, "#,##0.00")
        If objCADCONFFAT.VLTOTALFAT > 0 Then lblVLTOTAL.Caption = Format(objCADCONFFAT.VLTOTALFAT, "#,##0.00")
    
        Call PopGrdItens
    
    End If

End Sub

Private Sub PopGrdItens()

    Dim I           As Integer
    Dim curQTDORDEM As Currency
    Dim curQTDJAFAT As Currency
    Dim curQTDSALDO As Currency
    
    arrITENSFAT = objCADCONFFAT.ITENSFAT

    If IsArray(arrITENSFAT) Then
        With grdITENSPEDIDO
            For I = 1 To UBound(arrITENSFAT)
            
                .AddItem arrITENSFAT(I, 1) & vbTab & _
                         arrITENSFAT(I, 2) & vbTab & _
                         PegaDescProd(Str(arrITENSFAT(I, 1))) & vbTab & _
                         arrITENSFAT(I, 3) & vbTab & _
                         Format(arrITENSFAT(I, 4), "#,##0.00") & vbTab & _
                         Format(arrITENSFAT(I, 5), "#,##0.00") & vbTab & _
                         Format(arrITENSFAT(I, 6), "#,##0.00") & vbTab & _
                         Format(arrITENSFAT(I, 7), "#,##0.00") & vbTab & _
                         arrITENSFAT(I, 8) & vbTab & _
                         arrITENSFAT(I, 9) & vbTab & _
                         0 & vbTab & _
                         0 & vbTab & _
                         0 & vbTab & _
                         0 & vbTab & _
                         arrITENSFAT(I, 14)
                         
                .Cell(flexcpBackColor, (.Rows - 1), conCOL_Produto_QtdeFaturada) = &H80C0FF
            
                '' ----------------------------------------------------------
                curQTDORDEM = PegaQtdRealOrdem(Str(arrITENSFAT(I, 9)))
                curQTDJAFAT = PegaQtdJaFaturada(Str(arrITENSFAT(I, 9)))
                curQTDSALDO = (curQTDORDEM - curQTDJAFAT)
                
                .Cell(flexcpText, (.Rows - 1), conCOL_Produto_QtdOrdem) = curQTDORDEM
                .Cell(flexcpText, (.Rows - 1), conCOL_Produto_QtdTotFat) = curQTDJAFAT
                .Cell(flexcpText, (.Rows - 1), conCOL_Produto_SaldoOrdem) = curQTDSALDO
                
                '' Se o saldo for maior que 0 esta fechada parcial Status 1
                If curQTDSALDO > 0 Then .Cell(flexcpText, (.Rows - 1), conCOL_Produto_Status) = 1
                
                '' Se o saldo for menor que 0 esta fechada Total Status 2
                If curQTDSALDO <= 0 Then .Cell(flexcpText, (.Rows - 1), conCOL_Produto_Status) = 2
                '' ----------------------------------------------------------
            
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

Private Sub CarregaDadosCliente(strCODORD As String)

    sSql = "Select " & vbCrLf
    sSql = sSql & "       PED.* " & vbCrLf
    sSql = sSql & "      ,PED.SGI_CODIGO AS SGI_CODPED " & vbCrLf
    sSql = sSql & "      ,CLI.* " & vbCrLf
    sSql = sSql & "      ,PGT.SGI_DESCRICAO AS SGI_DESCPGTO " & vbCrLf
    sSql = sSql & "      ,TRA.SGI_DESCRICAO AS SGI_DESCTRANSP " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & strSGI_CADORDFATH & " ORDF" & vbCrLf
    sSql = sSql & "      ," & strSGI_CADPEDVENDH & " PED" & vbCrLf
    sSql = sSql & "      ,SGI_CADCLIENTE  CLI" & vbCrLf
    sSql = sSql & "      ,SGI_CADCONDPGTO PGT" & vbCrLf
    sSql = sSql & "      ,SGI_CADTRANSP   TRA" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       ORDF.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And ORDF.SGI_CODORD = " & Trim(strCODORD) & vbCrLf
    sSql = sSql & "   And PED.SGI_FILIAL = ORDF.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PED.SGI_CODIGO = ORDF.SGI_CODPED " & vbCrLf
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

Private Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    
    Me.Caption = "Cadastro de Confirmação de Faturamento - [ ALTERAÇÃO ]" & strCAPFILIAL
    
    objBLBFunc.LimpaCampos frmCADCONFFAT

    Frame2.Enabled = False
    Frame3.Enabled = True
    Frame4.Enabled = True
    Frame8.Enabled = False
    Frame5.Enabled = True
    txtOBS.Locked = False
    
    objCADCONFFAT.CODCONF = iCodigo
    
    Call ConfGridProdutos
    Call LimpaCamposLabel
    
    Call CarregaCampos
    
End Sub

Private Sub CalcSaldoFechado(lngCODCONF As Long, curQTDTOTFAT As Currency)
    
    objCADCONFFAT.CODPED = 0
    objCADCONFFAT.SALDOFECHADO = False
    
    Dim curSALDO        As Currency
    Dim curQTDPED       As Currency
    Dim curQTDFAT       As Currency
    Dim curQTDJAFAT     As Currency
    
    Dim lngTOTALOP      As Long
    Dim lngOPFECHADO    As Long
    Dim lngSALDOOP      As Long
    
    curSALDO = 0
    curQTDPED = 0
    curQTDFAT = 0
    curQTDJAFAT = 0
    
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PEDV.SGI_CODIGO " & vbCrLf
    sSql = sSql & "      ,PEDV.SGI_QTDEITENSPEDIDO " & vbCrLf
    sSql = sSql & "      ,PEDV.SGI_QTDEITENSFATURADOS " & vbCrLf
    sSql = sSql & "      ,ORDF.SGI_QTDETOTFAT " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & strSGI_CADORDFATH & " ORDF " & vbCrLf
    sSql = sSql & "      ," & strSGI_CADPEDVENDH & " PEDV " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       ORDF.SGI_FILIAL  = " & FILIAL & vbCrLf
    sSql = sSql & "   And ORDF.SGI_CODORD  = " & objCADCONFFAT.CODORD & vbCrLf
    sSql = sSql & "   And PEDV.SGI_FILIAL  = ORDF.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PEDV.SGI_CODIGO  = ORDF.SGI_CODPED "

    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then
        objCADCONFFAT.CODPED = BREC10!SGI_CODIGO
        If Not IsNull(BREC10!SGI_QTDEITENSPEDIDO) Then curQTDPED = BREC10!SGI_QTDEITENSPEDIDO
        If Not IsNull(BREC10!SGI_QTDEITENSFATURADOS) Then curQTDFAT = BREC10!SGI_QTDEITENSFATURADOS
        
        
        ''curSALDO = (curQTDPED - (curQTDFAT + curQTDTOTFAT))
        ''If curSALDO <= 0 Then objCADCONFFAT.SALDOFECHADO = True
        
        
        lngTOTALOP = 0
        sSql = "Select Count(SGI_CODPED) As SGI_QTDEOPTOTAL" & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       " & strSGI_ORDEMPROD & " " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_CODPED = " & BREC10!SGI_CODIGO
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC.EOF() Then
           If Not IsNull(BREC!SGI_QTDEOPTOTAL) Then lngTOTALOP = BREC!SGI_QTDEOPTOTAL
        End If
        BREC.Close
        
        lngOPFECHADO = 0
        sSql = "Select Count(SGI_CODPED) As SGI_QTDEOPFECHADA" & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       " & strSGI_ORDEMPROD & " " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_CODPED = " & BREC10!SGI_CODIGO & vbCrLf
        sSql = sSql & "   And SGI_STATUS = 2"
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC.EOF() Then
           If Not IsNull(BREC!SGI_QTDEOPFECHADA) Then lngOPFECHADO = BREC!SGI_QTDEOPFECHADA
        End If
        BREC.Close
    
        lngOPFECHADO = (lngOPFECHADO + lngOPFECHADA)
        
        
        lngSALDOOP = (lngTOTALOP - lngOPFECHADO)
        If lngSALDOOP <= 0 Then objCADCONFFAT.SALDOFECHADO = True
    
    End If
    BREC10.Close
    
    objCADCONFFAT.QTDETOTALFAT = (curQTDFAT + curQTDTOTFAT)
    
End Sub


Private Sub CalcTotFatura()

    Dim I                   As Integer
    Dim curTotalFatura      As Currency
    
    lblTOTALFAT.Caption = ""
    curTotalFatura = 0
    
    With grdITENSPEDIDO
        For I = 1 To (.Rows - 1)
            curTotalFatura = curTotalFatura + CCur(.Cell(flexcpText, I, conCOL_Produto_VLFaturado))
        Next I
    End With
    lblTOTALFAT.Caption = Format(curTotalFatura, "#,##0.00")
    
End Sub

Private Function PegaQtdJaFaturada(strCODORDFAB As String) As Currency

    PegaQtdJaFaturada = 0
    
    sSql = "Select " & vbTab
    sSql = sSql & "       ORD.SGI_QTDFAT " & vbTab
    sSql = sSql & "  From " & vbTab
    sSql = sSql & "       " & strSGI_ORDEMPROD & " ORD " & vbTab
    sSql = sSql & " Where " & vbTab
    sSql = sSql & "       ORD.SGI_FILIAL  = " & FILIAL & vbTab
    sSql = sSql & "  And  ORD.SGI_CODIGO  = " & strCODORDFAB
    
    BREC6.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC6.EOF Then
        If Not IsNull(BREC6!SGI_QTDFAT) Then PegaQtdJaFaturada = BREC6!SGI_QTDFAT
    End If
    BREC6.Close
    
End Function

Private Function PegaQtdRealOrdem(strCODORDFAB As String) As Currency

    PegaQtdRealOrdem = 0
    
    sSql = "Select " & vbTab
    sSql = sSql & "       ORD.SGI_QTDE " & vbTab
    sSql = sSql & "  From " & vbTab
    sSql = sSql & "       " & strSGI_ORDEMPROD & " ORD " & vbTab
    sSql = sSql & " Where " & vbTab
    sSql = sSql & "       ORD.SGI_FILIAL  = " & FILIAL & vbTab
    sSql = sSql & "  And  ORD.SGI_CODIGO  = " & strCODORDFAB
    
    BREC6.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC6.EOF Then
        If Not IsNull(BREC6!SGI_QTDE) Then PegaQtdRealOrdem = BREC6!SGI_QTDE
    End If
    BREC6.Close
    
End Function


Private Function OPFECHADA(strCODOP As String, strIDPRODUTO As String, curQTDE As Currency) As Long

    OPFECHADA = 0
    Dim curQTDEOP   As Currency
    Dim curQTDEFAT  As Currency
    Dim curSALDOP   As Currency
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & strSGI_ORDEMPROD & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO    = " & strCODOP & vbCrLf
    sSql = sSql & "   And SGI_IDPRODUTO = " & strIDPRODUTO
    
    BREC4.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC4.EOF() Then
        
        If Not IsNull(BREC4!SGI_QTDEPED) Then curQTDEOP = BREC4!SGI_QTDEPED
        If Not IsNull(BREC4!SGI_QTDFAT) Then curQTDEFAT = BREC4!SGI_QTDFAT
        
        curQTDEFAT = (curQTDEFAT + curQTDE)
        curSALDOP = (curQTDEOP - curQTDEFAT)
        
        If curSALDOP <= 0 Then OPFECHADA = 1
    End If
    BREC4.Close
    
End Function

Private Sub DestroiObjeto()
    Set objBLBFunc = Nothing
    Set objCADCONFFAT = Nothing
    Set objPESQPADRAO = Nothing
End Sub


Private Function SaldoPedido(strCODPED As String, lngQTDEFAT As Currency) As Long

    SaldoPedido = 0

    If Len(Trim(Replace(strCODPED, "/", ""))) = 0 Then Exit Function
    
    Dim strFILIAL As String
    Dim lngQTDETOT_ITENSPED     As Long
    Dim lngQTDETOT_JAFATURADO   As Long
    
    strFILIAL = ""
    If intFILIALPED = 1 Then strFILIAL = "_STEEL"
    
    
    '' Pega a quantidade Total do PEdido (ITENS)
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_QTDEITENSPEDIDO" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPEDVENDH" & strFILIAL & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & Trim(Replace(strCODPED, "/", ""))
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then lngQTDETOT_ITENSPED = BREC!SGI_QTDEITENSPEDIDO
    BREC.Close
    ''---------------------------------------------
    
    '' Pega Quanto já foi faturado
    
    sSql = ""

    sSql = "Select " & vbCrLf
    sSql = sSql & "       Sum(ITE.SGI_QTDFAT) As SGI_QTDFAT" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADORDFATH" & strFILIAL & " CAB" & vbCrLf
    sSql = sSql & "     , SGI_CADORDFATI" & strFILIAL & " ITE" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       CAB.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And CAB.SGI_CODPED = " & Trim(Replace(strCODPED, "/", "")) & vbCrLf
    sSql = sSql & "   And CAB.SGI_STATUS = 1" & vbCrLf
    sSql = sSql & "   And ITE.SGI_FILIAL = CAB.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And ITE.SGI_CODORD = CAB.SGI_CODORD"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        If Not IsNull(BREC!SGI_QTDFAT) Then lngQTDETOT_JAFATURADO = BREC!SGI_QTDFAT
    End If
    BREC.Close
    ''---------------------------------------------
    
    objCADCONFFAT.QTDETOTALFAT = (lngQTDETOT_JAFATURADO + lngQTDEFAT)
    
    SaldoPedido = (lngQTDETOT_ITENSPED - (lngQTDETOT_JAFATURADO + lngQTDEFAT))
    
    
End Function

Private Function FechamentoOP(currQTDOP As Currency, currQTDFAT As Currency) As String

    FechamentoOP = ""

    Dim currTOLERANCIA   As Currency
    Dim currTOTOLETANCIA As Currency
    Dim currTOTOLOP      As Currency

    If currQTDOP <= 29000 Then currTOLERANCIA = 0.15
    If currQTDOP > 29001 Then currTOLERANCIA = 0.1
    
    currTOTOLETANCIA = (currQTDOP * currTOLERANCIA)
    currTOTOLOP = (currQTDOP - currTOTOLETANCIA)

    If currQTDFAT >= currTOTOLOP Then FechamentoOP = "F"
    If currQTDFAT < currTOTOLOP Then FechamentoOP = "A"

End Function

