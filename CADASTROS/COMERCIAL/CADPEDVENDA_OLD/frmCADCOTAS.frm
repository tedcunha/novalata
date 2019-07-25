VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmCADCOTAS 
   Caption         =   "Datas Disponiveis"
   ClientHeight    =   8055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6900
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   6900
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6855
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
         Picture         =   "frmCADCOTAS.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame26 
      Caption         =   "[ Datas Produção ]"
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
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   6855
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "frmCADCOTAS.frx":0102
         Left            =   120
         List            =   "frmCADCOTAS.frx":0104
         TabIndex        =   3
         Text            =   "cboMes"
         Top             =   240
         Width           =   2655
      End
      Begin VB.ComboBox cboAno 
         Height          =   315
         Left            =   2760
         TabIndex        =   2
         Text            =   "cboAno"
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton Command6 
         Height          =   300
         Left            =   5160
         Picture         =   "frmCADCOTAS.frx":0106
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Inclui uma nova linha na Gride"
         Top             =   240
         Width           =   300
      End
      Begin VSFlex8LCtl.VSFlexGrid grdSEMANAS 
         Height          =   2775
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   6615
         _cx             =   11668
         _cy             =   4895
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
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
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "OP's em atrazo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   32
         Top             =   3960
         Width           =   1290
      End
      Begin VB.Label lblDADOS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDADOS"
         Height          =   255
         Index           =   12
         Left            =   3480
         TabIndex        =   31
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label lblDADOS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDADOS"
         Height          =   255
         Index           =   11
         Left            =   3480
         TabIndex        =   30
         Top             =   6360
         Width           =   1455
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Há ser Empenhado"
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
         Left            =   120
         TabIndex        =   29
         Top             =   6360
         Width           =   1605
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         X1              =   120
         X2              =   4920
         Y1              =   5760
         Y2              =   5760
      End
      Begin VB.Label lblDADOS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDADOS"
         Height          =   255
         Index           =   10
         Left            =   3480
         TabIndex        =   28
         Top             =   6600
         Width           =   1455
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Saldo"
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
         TabIndex        =   27
         Top             =   6600
         Width           =   495
      End
      Begin VB.Label lblDADOS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDADOS"
         Height          =   255
         Index           =   9
         Left            =   3480
         TabIndex        =   26
         Top             =   6120
         Width           =   1455
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Empenhado no Dia"
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
         TabIndex        =   25
         Top             =   6120
         Width           =   1620
      End
      Begin VB.Label lblDADOS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDADOS"
         Height          =   255
         Index           =   8
         Left            =   3480
         TabIndex        =   24
         Top             =   5880
         Width           =   1455
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Disponivel"
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
         TabIndex        =   23
         Top             =   5880
         Width           =   1440
      End
      Begin VB.Label lblDADOS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDADOS"
         Height          =   255
         Index           =   7
         Left            =   3480
         TabIndex        =   22
         Top             =   5400
         Width           =   1455
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Total Comrprometido"
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
         Top             =   5400
         Width           =   1755
      End
      Begin VB.Label lblDADOS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDADOS"
         Height          =   255
         Index           =   6
         Left            =   3480
         TabIndex        =   20
         Top             =   5160
         Width           =   1455
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Pedidos Bloqueados P.Cota/P.Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   19
         Top             =   5160
         Width           =   3045
      End
      Begin VB.Label lblDADOS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDADOS"
         Height          =   255
         Index           =   5
         Left            =   3480
         TabIndex        =   18
         Top             =   4920
         Width           =   1455
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Pedidos Bloqueados no Fotolito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   4920
         Width           =   2700
      End
      Begin VB.Label lblDADOS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDADOS"
         Height          =   255
         Index           =   4
         Left            =   3480
         TabIndex        =   16
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Pedidos Bloqueados p/Alteração"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   4680
         Width           =   2790
      End
      Begin VB.Label lblDADOS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDADOS"
         Height          =   255
         Index           =   3
         Left            =   3480
         TabIndex        =   14
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Pedidos Bloqueados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   4440
         Width           =   1740
      End
      Begin VB.Label lblDADOS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDADOS"
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   12
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label lblDADOS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDADOS"
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   11
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label lblDADOS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDADOS"
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   10
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "OP Alocada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   4200
         Width           =   1020
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Cota do Dia"
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
         TabIndex        =   8
         Top             =   3720
         Width           =   1020
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Dia"
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
         Top             =   3480
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmCADCOTAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strIDPRODUTO         As String
Public strNOMFILIAL         As String
Public FILIAL               As Integer
Public mskDTPED             As String
Public cTipOper             As String
Public intALTFILME          As Integer
Public intFOTNOVO           As Integer
Public strRETORNO           As String
Public intAction2Do         As Integer
Public intStatusOP          As Integer
Public strPRODCODLIN        As String
Public lngSALDOQTDENTR      As Long
Public lngCODPED            As Long
Public strGRPCOD            As String
Public arrDIASCOTAS         As Variant
Public intHOMOLOGADO        As Integer
Public strIDINTERNO         As String

Dim objBLBFunc2             As Object
Dim objCADPEDVENDACOTA      As Object
Dim lngSALDOQTDENTR2        As Long
Dim lngTOTALATRAZADO        As Long
Dim lngALOCATRAZADO         As Long

'' ========================================================================================
Const conCOL_SonMes_Semana                     As Integer = 0
Const conCOL_SonMes_Domingo                    As Integer = 1
Const conCOL_SonMes_Segunda                    As Integer = 2
Const conCOL_SonMes_Terca                      As Integer = 3
Const conCOL_SonMes_Quarta                     As Integer = 4
Const conCOL_SonMes_Quinta                     As Integer = 5
Const conCOL_SonMes_Sexta                      As Integer = 6
Const conCOL_SonMes_Sabado                     As Integer = 7
Const conCOL_SonMes_IDINTERNO                  As Integer = 8
Const conCOL_SonMes_FormatString               As String = "=Sem|Dom|Seg|Ter|Qua|Qui|Sex|Sab|IDINTERNO"
Const conColumnsIn_SonMes                      As Integer = 6

Private Sub PegaDataLivre(strMES As String, strANO As String, strIDPRODUTO As String)

    Dim I As Integer
    
    Me.MousePointer = 0
    
    Call InitGridPCP
    
    If Len(Trim(strMES)) = 0 Then Exit Sub
    If Len(Trim(strANO)) = 0 Then Exit Sub
    If Len(Trim(strIDPRODUTO)) = 0 Then Exit Sub
    
    If BREC.State = 1 Then BREC.Close
    
    Me.MousePointer = 11
     
    Dim lngLINHA    As Long
    Dim dtDATA      As Date
    
    dtDATA = Date
         
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "      PMDIA.SGI_SEMANA" & vbCrLf
    sSql = sSql & "     ,PMDIA.SGI_DTSEMANA" & vbCrLf
    sSql = sSql & "     ,PMDIA.SGI_DIASEMANA" & vbCrLf
    sSql = sSql & "     ,Sum(PMDIA.SGI_QTDE) As SGI_QTDE" & vbCrLf
    sSql = sSql & "     ,GRPI.SGI_CODLIN" & vbCrLf
    sSql = sSql & "     ,PROD.SGI_NECKIN" & vbCrLf
    sSql = sSql & "     ,PM.SGI_ATIVO" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_CADPRODUTO            PROD" & vbCrLf
    sSql = sSql & "     ,SGI_CADLINHAPRODUTO       LINP" & vbCrLf
    sSql = sSql & "     ,SGI_CADGRUPLINHAIT" & strNOMFILIAL & "  GRPI" & vbCrLf
    sSql = sSql & "     ,SGI_CADPLANMESTRE" & strNOMFILIAL & "   PM" & vbCrLf
    sSql = sSql & "     ,SGI_CADDIASPMSEMANA" & strNOMFILIAL & " PMDIA" & vbCrLf
  
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      PROD.SGI_FILIAL         = " & FILIAL & vbCrLf
    sSql = sSql & "  And PROD.SGI_IDPRODUTO      = " & strIDPRODUTO & vbCrLf
    
    sSql = sSql & "  And LINP.SGI_FILIAL         = PROD.SGI_FILIAL" & vbCrLf
    sSql = sSql & "  And LINP.SGI_CODLIN         = PROD.SGI_CODLINPROD" & vbCrLf
    
    sSql = sSql & "  And GRPI.SGI_FILIAL         = LINP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "  And GRPI.SGI_CODLIN         = LINP.SGI_CODIGO" & vbCrLf
    sSql = sSql & "  And GRPI.SGI_OPTCOMNECKINSN = PROD.SGI_NECKIN" & vbCrLf
    sSql = sSql & "  And GRPI.SGI_HOMOLOGSN      = " & intHOMOLOGADO & vbCrLf
    
    sSql = sSql & "  And PM.SGI_FILIAL           = GRPI.SGI_FILIAL" & vbCrLf
    sSql = sSql & "  And PM.SGI_CODLINHA         = GRPI.SGI_CODIGO" & vbCrLf
    sSql = sSql & "  And PM.SGI_ATIVO            = 1" & vbCrLf
    sSql = sSql & "  And PM.SGI_MES              = " & strMES & vbCrLf
    sSql = sSql & "  And PM.SGI_ANO              = " & strANO & vbCrLf
    
    sSql = sSql & "  And PMDIA.SGI_FILIAL        = PM.SGI_FILIAL" & vbCrLf
    sSql = sSql & "  And PMDIA.SGI_CODIGO        = PM.SGI_CODIGO" & vbCrLf
    sSql = sSql & "  And PMDIA.SGI_ATIVO         = 1" & vbCrLf
    
    sSql = sSql & "Group By" & vbCrLf
    sSql = sSql & "        PMDIA.SGI_SEMANA" & vbCrLf
    sSql = sSql & "       ,PMDIA.SGI_DTSEMANA" & vbCrLf
    sSql = sSql & "       ,PMDIA.SGI_DIASEMANA" & vbCrLf
    sSql = sSql & "       ,GRPI.SGI_CODLIN" & vbCrLf
    sSql = sSql & "       ,PROD.SGI_NECKIN" & vbCrLf
    sSql = sSql & "       ,PM.SGI_ATIVO" & vbCrLf
    sSql = sSql & "Order By PMDIA.SGI_SEMANA" & vbCrLf
    sSql = sSql & "        ,PMDIA.SGI_DTSEMANA" & vbCrLf
    sSql = sSql & "        ,PMDIA.SGI_DIASEMANA"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        If BREC!SGI_ATIVO = 0 Then
           MsgBox "ATENÇÃO - O Plano mestre de produção esta desativado para este Mês !!!" & vbCrLf & _
                  "FAVOR INFORMAR O PCP", vbOKOnly + vbExclamation, "Aviso"
        Else
            With grdSEMANAS
                Do While Not BREC.EOF()
                    
                    lngLINHA = .FindRow(BREC!SGI_SEMANA, , conCOL_SonMes_Semana)
                    If lngLINHA < 0 Then
                        .AddItem BREC!SGI_SEMANA & vbTab & _
                                 "" & vbTab & _
                                 "" & vbTab & _
                                 "" & vbTab & _
                                 "" & vbTab & _
                                 "" & vbTab & _
                                 "" & vbTab & _
                                 "" & vbTab & _
                                 BREC!SGI_SEMANA
                        
                        .Cell(flexcpData, (.Rows - 1), conCOL_SonMes_Semana) = 0
                        .Cell(flexcpData, (.Rows - 1), conCOL_SonMes_Domingo) = 0
                        .Cell(flexcpData, (.Rows - 1), conCOL_SonMes_Segunda) = 0
                        .Cell(flexcpData, (.Rows - 1), conCOL_SonMes_Terca) = 0
                        .Cell(flexcpData, (.Rows - 1), conCOL_SonMes_Quarta) = 0
                        .Cell(flexcpData, (.Rows - 1), conCOL_SonMes_Quinta) = 0
                        .Cell(flexcpData, (.Rows - 1), conCOL_SonMes_Sexta) = 0
                        .Cell(flexcpData, (.Rows - 1), conCOL_SonMes_Sabado) = 0
                    End If
                    
                    If BREC!SGI_DTSEMANA >= dtDATA Then
                        ''If BREC!SGI_DTSEMANA = dtDATA Then lngALOCATRAZADO = lngTOTALATRAZADO
                        lngLINHA = .FindRow(BREC!SGI_SEMANA, , conCOL_SonMes_Semana)
                        If lngLINHA > 0 Then
                           .Cell(flexcpText, lngLINHA, BREC!SGI_DIASEMANA) = Day(BREC!SGI_DTSEMANA)
                           Call PintaCelula(BREC!SGI_CODLIN, Format(BREC!SGI_DTSEMANA, "DD/MM/YYYY"), lngLINHA, BREC!SGI_DIASEMANA, BREC!SGI_QTDE, BREC!SGI_NECKIN, intHOMOLOGADO)
                        End If
                    End If
    
                    BREC.MoveNext
                Loop
            End With
        End If
    Else
        MsgBox "ATENÇÃO - Não esta cadastrado plano mestre de produção para esta linha !!!" & vbCrLf & _
               "FAVOR INFORMAR O PCP", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC.Close

    Me.MousePointer = 0

End Sub

Private Sub cboAno_Change()
    Call Command6_Click
End Sub

Private Sub cboMes_Change()
    Call Command6_Click
End Sub

Private Sub cmdVoltar_Click()
    strRETORNO = ""
    Unload Me
End Sub

Private Sub Command6_Click()
    If ValidaCampos = False Then Exit Sub
    Call PegaDataLivre(cboMes.ItemData(Str(cboMes.ListIndex)), Str(cboAno.ItemData(cboAno.ListIndex)), strIDPRODUTO)
End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc2 = Nothing
    Set objCADPEDVENDACOTA = Nothing
End Sub

Private Sub Form_Activate()
    Call InitGridPCP
End Sub

Private Sub Form_Load()

    Set objBLBFunc2 = CreateObject("BLBCWS.clsFuncoes")
    Set objCADPEDVENDACOTA = CreateObject("CADPEDVENDA.clsCADPEDVENDA")


    Call objBLBFunc2.Preenche_Mes(cboMes)
    cboMes.ListIndex = (Month(CDate(Now)) - 1)
    
    Call objBLBFunc2.Preenche_Ano(cboAno)
    cboAno.ListIndex = 0

    Call InitGridPCP

    objCADPEDVENDACOTA.FILIAL = FILIAL

    intHOMOLOGADO = objCADPEDVENDACOTA.PegaHOMOLOGADO(strIDPRODUTO)
    lngTOTALATRAZADO = objCADPEDVENDACOTA.PegaAtrazados(strPRODCODLIN, strNOMFILIAL)

    Call LimpaArray

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub


Private Sub grdSEMANAS_DblClick()

On Error GoTo Err_grdSEMANAS_DblClick

    Dim intRESP         As Integer
    Dim I               As Integer
    Dim strDATA         As String
    Dim dtDTLIDTIME     As Date
    Dim lngDIASLIDTIME  As Long
    Dim lngSALDODISP    As Long
    
    With grdSEMANAS
        Select Case .Col
        Case conCOL_SonMes_Semana, _
             conCOL_SonMes_IDINTERNO
             Exit Sub
        Case conCOL_SonMes_Domingo, _
             conCOL_SonMes_Segunda, _
             conCOL_SonMes_Terca, _
             conCOL_SonMes_Quarta, _
             conCOL_SonMes_Quinta, _
             conCOL_SonMes_Sexta, _
             conCOL_SonMes_Sabado
             
             If cTipOper = "C" Then Exit Sub
             
             ''If CLng(lblDADOS(9)) = 0 Then
             ''   MsgBox "ATENÇÂO" & vbCrLf & _
             ''          "Empenho do Dia não pode ser ZERO !!!", vbOKOnly + vbExclamation, "Aviso"
             ''          Exit Sub
             ''End If
             
             If Len(Trim(grdSEMANAS.Cell(flexcpText, .Row, .Col))) = 0 Then
                Exit Sub
             Else
             
                
                strDATA = .Cell(flexcpText, .Row, .Col) & "/" & cboMes.ItemData(cboMes.ListIndex) & "/" & cboAno.ItemData(cboAno.ListIndex)
                
                '' Calcula Lid Time
                dtDTLIDTIME = objCADPEDVENDACOTA.dtLidTime(mskDTPED)
                Dim strDTLIBFOT As String
                Dim boolABORTA  As Boolean
             
                '' Verificando Lid Time de Fotolito
                If Trim(cTipOper) <> "V" Then '' Liberando fotolito
                    If intALTFILME = 1 Or _
                       intFOTNOVO = 1 Then
                        strDTLIBFOT = objCADPEDVENDACOTA.DtEntregaLDTIME(mskDTPED)
                        If CDate(strDATA) < CDate(strDTLIBFOT) Then
                           MsgBox "ATENÇÂO" & vbCrLf & _
                                  "A data escolhida " & Format(CDate(strDATA), "DD/MM/YYYY") & " não obedece os 15 dias de Lid-Time para o fotolito o sistema irá colocar uma possivel data de entrega !!!", vbOKOnly + vbExclamation, "Aviso"
                                    
                           boolABORTA = False
                           Do While boolABORTA = False
                                '' Verifica se já esta estourado a Cota
                                If CotaEstourada(strPRODCODLIN, strDTLIBFOT) = False Then
                                    strRETORNO = strDTLIBFOT
                                    boolABORTA = True
                                End If
                                strDTLIBFOT = (CDate(strDTLIBFOT) + 1)
                           Loop
                           Unload Me
                           Exit Sub
                        End If
                    End If
                End If
                
                '' P.Data
                If CDate(strDATA) < dtDTLIDTIME Then
                    intRESP = MsgBox("ATENÇÃO" & vbCrLf & "O Lid-Time para produção é de 15 Dias , a data escolhida foi " & Format(strDATA, "DD/MM/YYYY") & " e a data de Lid-Time é " & Format(dtDTLIDTIME, "DD/MM/YYYY") & "." & vbCrLf & "Deseja realmente escolher este dia ?", vbYesNo + vbQuestion + vbDefaultButton1, "Aviso")
                    If intRESP = vbNo Then
                        If intAction2Do = dacEnumUpdateAction_Ignore Then intAction2Do = dacEnumUpdateAction_update
                        intStatusOP = 0
                        strRETORNO = ""
                        Exit Sub
                    Else
                        If intAction2Do = dacEnumUpdateAction_Ignore Then intAction2Do = dacEnumUpdateAction_update
                        intStatusOP = 7
                        strRETORNO = strDATA
                        Unload Me
                        Exit Sub
                    End If
                Else
                    If intAction2Do = dacEnumUpdateAction_Ignore Then intAction2Do = dacEnumUpdateAction_update
                    intStatusOP = 0
                    strRETORNO = strDATA
                End If
                
                '' Verifica se já esta estourado a Cota
                If CotaEstourada(strPRODCODLIN, strDATA) = True Then
                    Call Command6_Click
                    Exit Sub
                Else
                    Unload Me
                    Exit Sub
                End If
                
                
                
                
                Call Command6_Click
             End If
             
        Case Else
            .ComboList = ""
        End Select
    End With
    
    Exit Sub

Err_grdSEMANAS_DblClick:

    Call objBLBFunc2.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : grdSEMANAS_DblClick()", Me.Name, "grdSEMANAS_DblClick)", strCAMARQERRO)

End Sub

Private Sub InitGridPCP()

    With grdSEMANAS
    
       .Cols = conColumnsIn_SonMes
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonMes_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonMes_Semana) = ""
       .ColDataType(conCOL_SonMes_Semana) = flexDTLong
       .Cell(flexcpFontSize) = 12
       
       .Cell(flexcpData, 0, conCOL_SonMes_Domingo) = ""
       .ColDataType(conCOL_SonMes_Domingo) = flexDTLong
       .Cell(flexcpFontSize) = 12
       
       .Cell(flexcpData, 0, conCOL_SonMes_Segunda) = ""
       .ColDataType(conCOL_SonMes_Segunda) = flexDTLong
       .Cell(flexcpFontSize) = 12
       
       .Cell(flexcpData, 0, conCOL_SonMes_Terca) = ""
       .ColDataType(conCOL_SonMes_Terca) = flexDTLong
       .Cell(flexcpFontSize) = 12
       
       .Cell(flexcpData, 0, conCOL_SonMes_Quarta) = ""
       .ColDataType(conCOL_SonMes_Quarta) = flexDTLong
       .Cell(flexcpFontSize) = 12
       
       .Cell(flexcpData, 0, conCOL_SonMes_Quinta) = ""
       .ColDataType(conCOL_SonMes_Quinta) = flexDTLong
       .Cell(flexcpFontSize) = 12
       
       .Cell(flexcpData, 0, conCOL_SonMes_Sexta) = ""
       .ColDataType(conCOL_SonMes_Sexta) = flexDTLong
       .Cell(flexcpFontSize) = 12
       
       .Cell(flexcpData, 0, conCOL_SonMes_Sabado) = ""
       .ColDataType(conCOL_SonMes_Sabado) = flexDTLong
       .Cell(flexcpFontSize) = 12
       
       .Cell(flexcpData, 0, conCOL_SonMes_IDINTERNO) = ""
       .ColDataType(conCOL_SonMes_IDINTERNO) = flexDTLong
       .Cell(flexcpFontSize) = 12
       
       .ColWidth(conCOL_SonMes_Semana) = 800
       .ColWidth(conCOL_SonMes_Domingo) = 800
       .ColWidth(conCOL_SonMes_Segunda) = 800
       .ColWidth(conCOL_SonMes_Terca) = 800
       .ColWidth(conCOL_SonMes_Quarta) = 800
       .ColWidth(conCOL_SonMes_Quinta) = 800
       .ColWidth(conCOL_SonMes_Sexta) = 800
       .ColWidth(conCOL_SonMes_Sabado) = 800
       .ColWidth(conCOL_SonMes_IDINTERNO) = 0
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
       .FrozenCols = conCOL_SonMes_Semana
    
    End With
    
End Sub

Private Sub PintaCelula(strCODLIN As String, strDTENTREGA As String, lngLINHA As Long, intDIA As Integer, lngQTDCOTA As Long, lngNECKIN As Long, intHOMOLOGADO As Integer)

    If Len(Trim(strCODLIN)) = 0 Then Exit Sub
    If Len(Trim(strDTENTREGA)) = 0 Then Exit Sub
    
    Dim lngSALDO                As Long
    Dim lngQTDE                 As Long
    Dim lngQTDEOP               As Long
    Dim lngQTDEPEDBLOQ          As Long
    Dim lngQTDEPEDBLOQ2         As Long
    Dim lngQTDEPEDBLOQALT       As Long
    Dim lngQTDEPEDBLOQALT2      As Long
    Dim lngQTDEPEDLBOQFOT       As Long
    Dim lngQTDEPEDLBOQFOT2      As Long
    Dim lngQTDEPEDBLOQPCPD      As Long
    Dim lngQTDEPEDBLOQPCPD2     As Long
    Dim lngQTDEATRAZO           As Long
    Dim arrARRGRPLIN()          As String
    Dim I                       As Long
    
    Dim strCODLIN2              As String
    Dim strGRPCOD               As String
        
    lngSALDO = 0
    lngQTDE = 0
    lngQTDEOP = 0
    lngQTDEPEDBLOQ = 0
    lngQTDEPEDBLOQ2 = 0
    lngQTDEPEDBLOQALT = 0
    lngQTDEPEDBLOQALT2 = 0
    lngQTDEPEDLBOQFOT = 0
    lngQTDEPEDLBOQFOT2 = 0
    lngQTDEPEDBLOQPCPD = 0
    lngQTDEPEDBLOQPCPD2 = 0

    
    strCODLIN2 = ""
    strGRPCOD = ""
    
    '' =========================
    sSql = ""
    
    sSql = "Select Distinct" & vbCrLf
    sSql = sSql & "       GRPI.*" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADGRUPLINHAIT" & strNOMFILIAL & "  GRPI" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       GRPI.SGI_FILIAL         = " & FILIAL & vbCrLf
    sSql = sSql & "   And GRPI.SGI_CODLIN         = " & strCODLIN & vbCrLf
    sSql = sSql & "   And GRPI.SGI_OPTCOMNECKINSN = " & lngNECKIN & vbCrLf
    sSql = sSql & "   And GRPI.SGI_HOMOLOGSN      = " & intHOMOLOGADO & vbCrLf
    
    BREC7.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC7.EOF()
        strGRPCOD = strGRPCOD & BREC7!SGI_CODIGO
        BREC7.MoveNext
        If Not BREC7.EOF() Then strGRPCOD = strGRPCOD & ","
    Loop
    BREC7.Close
    
    lngSALDOQTDENTR2 = PegaQtdeEntrega(Trim(Replace(strGRPCOD, ",", "")), Month(CDate(strDTENTREGA)), Year(CDate(strDTENTREGA)), strDTENTREGA, strIDINTERNO)
    
    '' =========================
    sSql = ""
    
    sSql = "Select Distinct" & vbCrLf
    sSql = sSql & "      LIMP.SGI_CODLIN" & vbCrLf
    sSql = sSql & "     ,(" & objCADPEDVENDACOTA.PegaQueryOPDia("LIMP.SGI_CODLIN", strDTENTREGA, strNOMFILIAL, lngCODPED, lngNECKIN, intHOMOLOGADO) & ") As QTDEOP" & vbCrLf
    sSql = sSql & "     ,(" & objCADPEDVENDACOTA.PegaQueryPedBloq("LIMP.SGI_CODLIN", strDTENTREGA, strNOMFILIAL, lngCODPED, lngNECKIN, intHOMOLOGADO) & ") As QTDEPEDBLOQ" & vbCrLf
    sSql = sSql & "     ,(" & objCADPEDVENDACOTA.PegaPedQueryBloqAlt("LIMP.SGI_CODLIN", strDTENTREGA, strNOMFILIAL, lngCODPED, lngNECKIN, intHOMOLOGADO) & ") As QTDEPEDBLOQALT" & vbCrLf
    sSql = sSql & "     ,(" & objCADPEDVENDACOTA.PegaPedQueryBloqFot("LIMP.SGI_CODLIN", strDTENTREGA, strNOMFILIAL, lngCODPED, lngNECKIN, intHOMOLOGADO) & ") As QTDEPEDLBOQFOT" & vbCrLf
    sSql = sSql & "     ,(" & objCADPEDVENDACOTA.PegaPedQueryBloqPcPd("LIMP.SGI_CODLIN", strDTENTREGA, strNOMFILIAL, lngCODPED, lngNECKIN, intHOMOLOGADO) & ") As QTDEPEDBLOQPCPD" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_CADGRUPLINHAIT" & strNOMFILIAL & " GRPI" & vbCrLf
    sSql = sSql & "     ,SGI_CADLINHAPRODUTO      LIMP" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      GRPI.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "  And GRPI.SGI_CODIGO IN(" & Trim(strGRPCOD) & ")" & vbCrLf
    sSql = sSql & "  And GRPI.SGI_OPTCOMNECKINSN = " & lngNECKIN & vbCrLf
    sSql = sSql & "  And GRPI.SGI_HOMOLOGSN      = " & intHOMOLOGADO & vbCrLf
    
    sSql = sSql & "  And LIMP.SGI_FILIAL = GRPI.SGI_FILIAL" & vbCrLf
    sSql = sSql & "  And LIMP.SGI_CODIGO = GRPI.SGI_CODLIN" & vbCrLf
    BREC7.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC7.EOF() Then
        Do While Not BREC7.EOF()
        
            lngQTDEOP = 0
            lngQTDEPEDBLOQ2 = 0
            lngQTDEPEDBLOQALT2 = 0
            lngQTDEPEDLBOQFOT2 = 0
            lngQTDEPEDBLOQPCPD2 = 0
            
            If Not IsNull(BREC7!QtdeOP) Then lngQTDEOP = BREC7!QtdeOP
            If Not IsNull(BREC7!QTDEPEDBLOQ) Then lngQTDEPEDBLOQ2 = BREC7!QTDEPEDBLOQ
            If Not IsNull(BREC7!QTDEPEDBLOQALT) Then lngQTDEPEDBLOQALT2 = BREC7!QTDEPEDBLOQALT
            If Not IsNull(BREC7!QTDEPEDLBOQFOT) Then lngQTDEPEDLBOQFOT2 = BREC7!QTDEPEDLBOQFOT
            If Not IsNull(BREC7!QTDEPEDBLOQPCPD) Then lngQTDEPEDBLOQPCPD2 = BREC7!QTDEPEDBLOQPCPD
            
            lngQTDE = (lngQTDE + lngQTDEOP)
            lngQTDEPEDBLOQ = (lngQTDEPEDBLOQ + lngQTDEPEDBLOQ2)
            lngQTDEPEDBLOQALT = (lngQTDEPEDBLOQALT + lngQTDEPEDBLOQALT2)
            lngQTDEPEDLBOQFOT = (lngQTDEPEDLBOQFOT + lngQTDEPEDLBOQFOT2)
            lngQTDEPEDBLOQPCPD = (lngQTDEPEDBLOQPCPD + lngQTDEPEDBLOQPCPD2)
            
            BREC7.MoveNext
        Loop
    
        grdSEMANAS.Cell(flexcpData, lngLINHA, intDIA) = Trim(Str(lngQTDCOTA)) & "^" & _
                                                        Trim(Str(lngQTDE)) & "^" & _
                                                        Trim(Str(lngQTDEPEDBLOQ)) & "^" & _
                                                        Trim(Str(lngQTDEPEDBLOQALT)) & "^" & _
                                                        Trim(Str(lngQTDEPEDLBOQFOT)) & "^" & _
                                                        Trim(Str(lngQTDEPEDBLOQPCPD))
    
        lngSALDO = (lngQTDCOTA - (lngQTDE + lngSALDOQTDENTR2 + lngQTDEPEDBLOQ + lngQTDEPEDBLOQALT + lngQTDEPEDLBOQFOT + lngQTDEPEDBLOQPCPD + lngTOTALATRAZADO))
        
        ''If lngSALDO > 0 Then
        ''    lngALOCATRAZADO = (lngALOCATRAZADO - lngSALDO)
        ''    If lngALOCATRAZADO > 0 Then lngSALDO = (lngSALDO - lngSALDO)
        ''End If
        
        If lngSALDO <= 0 Then
            '' Vermelho
            grdSEMANAS.Cell(flexcpBackColor, lngLINHA, intDIA) = &HFF&
            grdSEMANAS.Cell(flexcpData, lngLINHA, intDIA) = grdSEMANAS.Cell(flexcpData, lngLINHA, intDIA) & "^1"
        Else
            '' Verde
            grdSEMANAS.Cell(flexcpBackColor, lngLINHA, intDIA) = &HC000&
            grdSEMANAS.Cell(flexcpData, lngLINHA, intDIA) = grdSEMANAS.Cell(flexcpData, lngLINHA, intDIA) & "^0"
        End If
    
    End If
    BREC7.Close
    
End Sub

Private Function ValidaCampos() As Boolean
    ValidaCampos = False
    If cboAno.ListIndex = -1 Then Exit Function
    If cboMes.ListIndex = -1 Then Exit Function
    If CInt(cboMes.ItemData(cboMes.ListIndex)) < CInt(Month(CDate(mskDTPED))) Then
        If cboAno.ItemData(cboAno.ListIndex) <= CInt(Year(CDate(mskDTPED))) Then
            MsgBox "ATENÇÃO" & vbCrLf & "O Mês não ´pode ser menor que o Mês do Pedido !!!", vbOKOnly + vbExclamation, "Aviso"
            cboMes.ListIndex = (Month(CDate(mskDTPED)) - 1)
            Exit Function
        End If
    End If
    
''    If CDbl(cboMes.ItemData(cboMes.ListIndex) & cboAno.ItemData(cboAno.ListIndex)) < CDbl(Month(CDate(mskDTPED)) & Year(CDate(mskDTPED))) Then
''        MsgBox "ATENÇÃO" & vbCrLf & "O Mês não ´pode ser menor que o Mês do Pedido !!!", vbOKOnly + vbExclamation, "Aviso"
''        cboMes.ListIndex = (Month(CDate(mskDTPED)) - 1)
''        Exit Function
 ''   End If
    ValidaCampos = True
End Function

Private Sub MostraDadosCotas()

    Dim arrDADOS()      As String
    Dim lngTOTCOMPROM   As Long
    Dim lngSALDODISP    As Long
    Dim lngSALDO        As Long
    Dim lngTOTEMPDIA    As Long
    Dim I               As Integer
    Dim strDATA         As String
    Dim lngTOTEMP       As Long
    
    Dim lngTOTDIA       As Long
    
    Call LimpaArray

    If grdSEMANAS.Row = 0 Then Exit Sub
    If Len(Trim(grdSEMANAS.Cell(flexcpData, grdSEMANAS.Row, grdSEMANAS.Col))) > 0 And Len(Trim(grdSEMANAS.Cell(flexcpText, grdSEMANAS.Row, grdSEMANAS.Col))) > 0 Then
       
        strDATA = grdSEMANAS.Cell(flexcpText, grdSEMANAS.Row, grdSEMANAS.Col) & "/" & Format(cboMes.ItemData(cboMes.ListIndex), "00") & "/" & cboAno.ItemData(cboAno.ListIndex)
        
        lngTOTEMPDIA = PegaQtdeEntrega(strGRPCOD, Month(CDate(strDATA)), Year(CDate(strDATA)), strDATA, strIDINTERNO)
        lngTOTEMP = (PegaQtdeEntregaInd(strGRPCOD, Month(CDate(strDATA)), Year(CDate(strDATA)), strDATA, strIDINTERNO) - lngSALDOQTDENTR)
        If lngTOTEMP < 0 Then lngTOTEMP = (lngTOTEMP * -1)
        
        arrDADOS = Split(grdSEMANAS.Cell(flexcpData, grdSEMANAS.Row, grdSEMANAS.Col), "^")
        lblDADOS(0).Caption = grdSEMANAS.Cell(flexcpText, grdSEMANAS.Row, grdSEMANAS.Col)
        If IsArray(arrDADOS) Then
            If UBound(arrDADOS) = 0 Then Exit Sub
            lblDADOS(1).Caption = arrDADOS(0)   '' Cota do Dia
            
            lblDADOS(2).Caption = arrDADOS(1)   ''
            lblDADOS(3).Caption = arrDADOS(2)   ''
            lblDADOS(4).Caption = arrDADOS(3)   ''
            lblDADOS(5).Caption = arrDADOS(4)   ''
            lblDADOS(6).Caption = arrDADOS(5)   ''
            lblDADOS(12).Caption = lngTOTALATRAZADO
        
            lngTOTCOMPROM = (CLng(arrDADOS(1)) + _
                             CLng(arrDADOS(2)) + _
                             CLng(arrDADOS(3)) + _
                             CLng(arrDADOS(4)) + _
                             CLng(arrDADOS(5)) + _
                             lngTOTALATRAZADO)
            
            lblDADOS(7).Caption = lngTOTCOMPROM                         '   ' Comprometido
            
            lngSALDODISP = (CLng(arrDADOS(0)) - lngTOTCOMPROM)
            lblDADOS(8).Caption = lngSALDODISP                              '' Saldo Disponivel
            
            lblDADOS(9).Caption = lngTOTEMPDIA                              '' Empenhado no Dia
            lblDADOS(11).Caption = lngTOTEMP                                '' Há Empenhar no Dia
            
            lngTOTDIA = (lngTOTEMPDIA + lngTOTEMP)
            lngTOTDIA = (lngTOTDIA * -1)
            
            lngSALDO = (lngSALDODISP + lngTOTDIA)
            lblDADOS(10).Caption = lngSALDO                                 '' Saldo
        
        End If
    End If

End Sub


Private Sub grdSEMANAS_Click()
    Call MostraDadosCotas
End Sub


Private Sub grdSEMANAS_RowColChange()
    Call MostraDadosCotas
End Sub


Private Function CotaEstourada(strCODLIN As String, strDTENTREGA As String) As Boolean

    CotaEstourada = True
    
    If Len(Trim(strCODLIN)) = 0 Then Exit Function
    If Len(Trim(strDTENTREGA)) = 0 Then Exit Function
    
    Dim lngSALDO                As Long
    Dim lngQTDE                 As Long
    Dim lngQTDEOP               As Long
    Dim lngQTDEPEDBLOQ          As Long
    Dim lngQTDEPEDBLOQ2         As Long
    Dim lngQTDEPEDBLOQALT       As Long
    Dim lngQTDEPEDBLOQALT2      As Long
    Dim lngQTDEPEDLBOQFOT       As Long
    Dim lngQTDEPEDLBOQFOT2      As Long
    Dim lngQTDEPEDBLOQPCPD      As Long
    Dim lngQTDEPEDBLOQPCPD2     As Long
    Dim lngSALDOQTDENTR3        As Long
    Dim lngSALDOQTDENTR4        As Long
    Dim lngALOCADOPDIA          As Long
    Dim lngQTDCOTA              As Long
    Dim lngTOTALSALDO           As Long
    Dim intRESP                 As Integer
    Dim arrGRPLIN()             As String
    Dim I                       As Integer
    
    Dim strCODLIN2              As String
    Dim strGRPCOD               As String
    Dim lngNECKIN               As Long
    Dim arrDADOS()              As String
        
    lngSALDO = 0
    lngQTDE = 0
    lngQTDEOP = 0
    lngQTDEPEDBLOQ = 0
    lngQTDEPEDBLOQ2 = 0
    lngQTDEPEDBLOQALT = 0
    lngQTDEPEDBLOQALT2 = 0
    lngQTDEPEDLBOQFOT = 0
    lngQTDEPEDLBOQFOT2 = 0
    lngQTDEPEDBLOQPCPD = 0
    lngQTDEPEDBLOQPCPD2 = 0
    lngTOTALSALDO = 0
    
    strCODLIN2 = ""
    strGRPCOD = ""
    
    lngNECKIN = objCADPEDVENDACOTA.PegaNECKIN(strIDPRODUTO)
    intHOMOLOGADO = objCADPEDVENDACOTA.PegaHOMOLOGADO(strIDPRODUTO)
    
    '' =========================
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       GRPI.*" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO       LIMP" & vbCrLf
    sSql = sSql & "      ,SGI_CADGRUPLINHAIT" & strNOMFILIAL & "  GRPI" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       LIMP.SGI_FILIAL         = " & FILIAL & vbCrLf
    sSql = sSql & "   And LIMP.SGI_CODLIN         = " & Trim(strCODLIN) & vbCrLf
    sSql = sSql & "   And GRPI.SGI_FILIAL         = LIMP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And GRPI.SGI_CODLIN         = LIMP.SGI_CODIGO" & vbCrLf
    sSql = sSql & "   And GRPI.SGI_OPTCOMNECKINSN = " & lngNECKIN & vbCrLf
    sSql = sSql & "   And GRPI.SGI_HOMOLOGSN      = " & intHOMOLOGADO & vbCrLf
   
    BREC7.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC7.EOF()
        strGRPCOD = strGRPCOD & BREC7!SGI_CODIGO
        BREC7.MoveNext
        If Not BREC7.EOF() Then strGRPCOD = strGRPCOD & ","
    Loop
    BREC7.Close
    
    lngSALDOQTDENTR3 = PegaQtdeEntrega(Trim(Replace(strGRPCOD, ",", "")), Month(CDate(strDTENTREGA)), Year(CDate(strDTENTREGA)), strDTENTREGA, strIDINTERNO)
    lngSALDOQTDENTR3 = lngSALDOQTDENTR3 + lngSALDOQTDENTR
    
    '' Pega Cota
    lngQTDCOTA = objCADPEDVENDACOTA.PegaCota(Trim(strIDPRODUTO), strDTENTREGA, strNOMFILIAL, intHOMOLOGADO)
    
    '' =========================
    sSql = ""
    
    sSql = "Select Distinct" & vbCrLf
    sSql = sSql & "      LIMP.SGI_CODLIN" & vbCrLf
    sSql = sSql & "     ,(" & objCADPEDVENDACOTA.PegaQueryOPDia("LIMP.SGI_CODLIN", strDTENTREGA, strNOMFILIAL, lngCODPED, lngNECKIN, intHOMOLOGADO) & ") As QTDEOP" & vbCrLf
    sSql = sSql & "     ,(" & objCADPEDVENDACOTA.PegaQueryPedBloq("LIMP.SGI_CODLIN", strDTENTREGA, strNOMFILIAL, lngCODPED, lngNECKIN, intHOMOLOGADO) & ") As QTDEPEDBLOQ" & vbCrLf
    sSql = sSql & "     ,(" & objCADPEDVENDACOTA.PegaPedQueryBloqAlt("LIMP.SGI_CODLIN", strDTENTREGA, strNOMFILIAL, lngCODPED, lngNECKIN, intHOMOLOGADO) & ") As QTDEPEDBLOQALT" & vbCrLf
    sSql = sSql & "     ,(" & objCADPEDVENDACOTA.PegaPedQueryBloqFot("LIMP.SGI_CODLIN", strDTENTREGA, strNOMFILIAL, lngCODPED, lngNECKIN, intHOMOLOGADO) & ") As QTDEPEDLBOQFOT" & vbCrLf
    sSql = sSql & "     ,(" & objCADPEDVENDACOTA.PegaPedQueryBloqPcPd("LIMP.SGI_CODLIN", strDTENTREGA, strNOMFILIAL, lngCODPED, lngNECKIN, intHOMOLOGADO) & ") As QTDEPEDBLOQPCPD" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_CADGRUPLINHAIT" & strNOMFILIAL & " GRPI" & vbCrLf
    sSql = sSql & "     ,SGI_CADLINHAPRODUTO      LIMP" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      GRPI.SGI_FILIAL         = " & FILIAL & vbCrLf
    sSql = sSql & "  And GRPI.SGI_CODIGO IN(" & Trim(strGRPCOD) & ")" & vbCrLf
    sSql = sSql & "  And GRPI.SGI_OPTCOMNECKINSN = " & lngNECKIN & vbCrLf
    sSql = sSql & "  And GRPI.SGI_HOMOLOGSN      = " & intHOMOLOGADO & vbCrLf
    sSql = sSql & "  And LIMP.SGI_FILIAL         = GRPI.SGI_FILIAL" & vbCrLf
    sSql = sSql & "  And LIMP.SGI_CODIGO         = GRPI.SGI_CODLIN"
    
    BREC7.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC7.EOF() Then
        Do While Not BREC7.EOF()
        
            lngQTDEOP = 0
            lngQTDEPEDBLOQ2 = 0
            lngQTDEPEDBLOQALT2 = 0
            lngQTDEPEDLBOQFOT2 = 0
            lngQTDEPEDBLOQPCPD2 = 0
            
            If Not IsNull(BREC7!QtdeOP) Then lngQTDEOP = BREC7!QtdeOP
            If Not IsNull(BREC7!QTDEPEDBLOQ) Then lngQTDEPEDBLOQ2 = BREC7!QTDEPEDBLOQ
            If Not IsNull(BREC7!QTDEPEDBLOQALT) Then lngQTDEPEDBLOQALT2 = BREC7!QTDEPEDBLOQALT
            If Not IsNull(BREC7!QTDEPEDLBOQFOT) Then lngQTDEPEDLBOQFOT2 = BREC7!QTDEPEDLBOQFOT
            If Not IsNull(BREC7!QTDEPEDBLOQPCPD) Then lngQTDEPEDBLOQPCPD2 = BREC7!QTDEPEDBLOQPCPD
            
            lngQTDE = (lngQTDE + lngQTDEOP)
            lngQTDEPEDBLOQ = (lngQTDEPEDBLOQ + lngQTDEPEDBLOQ2)
            lngQTDEPEDBLOQALT = (lngQTDEPEDBLOQALT + lngQTDEPEDBLOQALT2)
            lngQTDEPEDLBOQFOT = (lngQTDEPEDLBOQFOT + lngQTDEPEDLBOQFOT2)
            lngQTDEPEDBLOQPCPD = (lngQTDEPEDBLOQPCPD + lngQTDEPEDBLOQPCPD2)
            
            BREC7.MoveNext
        Loop
        
        lngALOCADOPDIA = (lngQTDE + lngSALDOQTDENTR3 + lngQTDEPEDBLOQ + lngQTDEPEDBLOQALT + lngQTDEPEDLBOQFOT + lngQTDEPEDBLOQPCPD + lngTOTALATRAZADO)
    
    End If
    BREC7.Close

    lngSALDO = (lngQTDCOTA - lngALOCADOPDIA)
    
    If lngSALDO >= 0 Then CotaEstourada = False

    If (intALTFILME = 1 Or intFOTNOVO = 1) Then
        CotaEstourada = False
        Exit Function
    End If
    
    If CotaEstourada Then
        If intStatusOP = 7 Then Exit Function
        intRESP = MsgBox("ATENÇÃO" & vbCrLf & "A Cota para para o dia " & Format(CDate(strDTENTREGA), "DD/MM/YYYY") & " já está estourada." & vbCrLf & _
                         vbCrLf & _
                         "Deseja realmente escolher este dia ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
        
        If intRESP = vbNo Then
            If intAction2Do = dacEnumUpdateAction_Ignore Then intAction2Do = dacEnumUpdateAction_update
            intStatusOP = 0
            strRETORNO = ""
        Else
            If intAction2Do = dacEnumUpdateAction_Ignore Then intAction2Do = dacEnumUpdateAction_update
            intStatusOP = 6
            strRETORNO = strDTENTREGA
            Unload Me
        End If
    Else
        arrDADOS = Split(grdSEMANAS.Cell(flexcpData, grdSEMANAS.Row, grdSEMANAS.Col), "^")
        If arrDADOS(6) = 1 Then
            If intAction2Do = dacEnumUpdateAction_Ignore Then intAction2Do = dacEnumUpdateAction_update
            intStatusOP = 6
            strRETORNO = strDTENTREGA
            Unload Me
        Else
            If intAction2Do = dacEnumUpdateAction_Ignore Then intAction2Do = dacEnumUpdateAction_update
            intStatusOP = 0
            Unload Me
        End If
    End If
    
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub LimpaArray()
    Dim I As Integer
    For I = 0 To 12
        lblDADOS(I).Caption = ""
    Next I
End Sub

Private Function PegaQtdeEntrega(strIDPRODUTO As String, strMES As String, strANO As String, strDia As String, strIDINTERNO As String) As Long
    
On Error GoTo Err_PegaQtdeEntrega
    
    PegaQtdeEntrega = 0
    
    Dim I       As Integer
    
    For I = 1 To UBound(arrDIASCOTAS)
        If Len(Trim(arrDIASCOTAS(I, 1))) > 0 And _
           Len(Trim(arrDIASCOTAS(I, 2))) > 0 Then
           
            If Trim(strIDPRODUTO) = Trim(arrDIASCOTAS(I, 4)) And _
               Trim(strMES) = Trim(Str(Month(CDate(arrDIASCOTAS(I, 1))))) And _
               Trim(strANO) = Trim(Str(Year(CDate(arrDIASCOTAS(I, 1))))) And _
               CDate(arrDIASCOTAS(I, 1)) = CDate(strDia) Then
               PegaQtdeEntrega = PegaQtdeEntrega + CLng(arrDIASCOTAS(I, 2))
            End If
        End If
    Next I
    
    Exit Function
    
Err_PegaQtdeEntrega:
    
    Call objBLBFunc2.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PegaQtdeEntrega()", Me.Name, "PegaQtdeEntrega()", strCAMARQERRO)
    
End Function

Private Function PegaQtdeEntregaInd(strIDPRODUTO As String, strMES As String, strANO As String, strDia As String, strIDINTERNO As String) As Long
    
On Error GoTo Err_PegaQtdeEntregaInd
    
    PegaQtdeEntregaInd = 0
    
    Dim I       As Integer
    
    For I = 1 To UBound(arrDIASCOTAS)
        If Len(Trim(arrDIASCOTAS(I, 1))) > 0 And _
           Len(Trim(arrDIASCOTAS(I, 2))) > 0 Then
           
            If Trim(strIDPRODUTO) = Trim(arrDIASCOTAS(I, 4)) And _
               Trim(strMES) = Trim(Str(Month(CDate(arrDIASCOTAS(I, 1))))) And _
               Trim(strANO) = Trim(Str(Year(CDate(arrDIASCOTAS(I, 1))))) And _
               CDate(arrDIASCOTAS(I, 1)) = CDate(strDia) And _
               Trim(arrDIASCOTAS(I, 5)) = Trim(strIDINTERNO) Then
               PegaQtdeEntregaInd = PegaQtdeEntregaInd + CLng(arrDIASCOTAS(I, 2))
            End If
        End If
    Next I
    
    Exit Function
    
Err_PegaQtdeEntregaInd:
    
    Call objBLBFunc2.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PegaQtdeEntrega()", Me.Name, "PegaQtdeEntrega()", strCAMARQERRO)
    
End Function


