VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCADOPCOMP 
   Caption         =   "Cadastro de Ordem de Fabricação de Componentes"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11835
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   11835
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      Caption         =   "[ Dados Complementares ]"
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
      Height          =   1935
      Left            =   6120
      TabIndex        =   29
      Top             =   5520
      Width           =   5535
      Begin VB.Label lblQTDEFOLHAS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1920
         TabIndex        =   33
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Qtde de Folhas"
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
         TabIndex        =   32
         Top             =   720
         Width           =   1305
      End
      Begin VB.Label lblDIMCORTE 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1920
         TabIndex        =   31
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dimensão de Corte"
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
         Top             =   360
         Width           =   1605
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "[ Observação ]"
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
      Height          =   2055
      Left            =   120
      TabIndex        =   23
      Top             =   5400
      Width           =   5895
      Begin VB.TextBox txtOBS 
         Height          =   1695
         Left            =   120
         MaxLength       =   300
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Text            =   "frmCADOPCOMP.frx":0000
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "[ Dados da OP ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   0
      TabIndex        =   20
      Top             =   960
      Width           =   11775
      Begin VB.CommandButton cmdLinProd 
         Height          =   315
         Left            =   2640
         Picture         =   "frmCADOPCOMP.frx":0009
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtLinProd 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   9
         Text            =   "txtLinProd"
         Top             =   600
         Width           =   1335
      End
      Begin VB.Frame Frame3 
         Caption         =   "[ Materiais Usados - Padrão ]"
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
         Height          =   2895
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   11535
         Begin VSFlex8LCtl.VSFlexGrid grdProdInsumos 
            Height          =   2535
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   11295
            _cx             =   19923
            _cy             =   4471
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
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   2640
         Picture         =   "frmCADOPCOMP.frx":010B
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtCODPROD 
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Text            =   "txtCODPROD"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtQTDEOP 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6720
         TabIndex        =   5
         Text            =   "txtQTDEOP"
         Top             =   240
         Width           =   1455
      End
      Begin MSMask.MaskEdBox mskDTOP 
         Height          =   285
         Left            =   4200
         TabIndex        =   3
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtCODOP 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Text            =   "txtCODOP"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblLinhProd 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblLinhProd"
         Height          =   315
         Left            =   3000
         TabIndex        =   11
         Top             =   600
         Width           =   5535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Capacidade"
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
         Index           =   8
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1020
      End
      Begin VB.Label lblDescStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   9000
         TabIndex        =   7
         Top             =   240
         Width           =   2535
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
         Index           =   4
         Left            =   8280
         TabIndex        =   6
         Top             =   240
         Width           =   555
      End
      Begin VB.Label lblDescProd 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescProd"
         Height          =   285
         Left            =   3000
         TabIndex        =   15
         Top             =   1080
         Width           =   5535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Produto"
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
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Qtde da OP"
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
         Left            =   5640
         TabIndex        =   4
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data da OP"
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
         Left            =   3000
         TabIndex        =   2
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
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
         TabIndex        =   0
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   11775
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
         Picture         =   "frmCADOPCOMP.frx":020D
         Style           =   1  'Graphical
         TabIndex        =   19
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
         Picture         =   "frmCADOPCOMP.frx":030F
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Picture         =   "frmCADOPCOMP.frx":0411
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label lblDTCRIA 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblDTCRIA"
      Height          =   285
      Left            =   4440
      TabIndex        =   28
      Top             =   7560
      Width           =   2655
   End
   Begin VB.Label Label1 
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
      Index           =   6
      Left            =   3840
      TabIndex        =   27
      Top             =   7560
      Width           =   480
   End
   Begin VB.Label lblEmitidoPor 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblEmitidoPor"
      Height          =   285
      Left            =   1320
      TabIndex        =   26
      Top             =   7560
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Emitido Por:"
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
      TabIndex        =   25
      Top             =   7560
      Width           =   1035
   End
End
Attribute VB_Name = "frmCADOPCOMP"
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
Public lngCODUSUARIO    As Long
Public intFILIALPED     As Integer


Dim strCAPTION          As String
Dim objBLBFunc          As Object
Dim objCADOPCOMP        As Object
Dim objPESQPADRAO       As Object
Dim strNOMEFILIAL       As String
Dim strNOMETABELA       As String

Const conCOL_OrdFabC_IDProduto                       As Integer = 0
Const conCOL_OrdFabC_CodProd                         As Integer = 1
Const conCOL_OrdFabC_PesqProd                        As Integer = 2
Const conCOL_OrdFabC_DescProd                        As Integer = 3
Const conCOL_OrdFabC_AbilitSN                        As Integer = 4
Const conCOL_OrdFabC_FormatString                    As String = "=ID|Cod.Prod|...|Descrição do Produto|Abilitado S/N"
Const conColumnsIn_OrdFabC                           As Integer = 5


Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    cTipOper = "A"
    
    Call objBLBFunc.MudaBotoes(CmdSalva, cmdAltera, cTipOper)
    Call objBLBFunc.MudaCaption(Me, strCAPTION, cTipOper)
    Call DesabilitaCampos(cTipOper)

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
    End If
    If cTipOper = "A" Then
        txtCODPROD.Text = ""
        txtCODPROD.Tag = ""
        lblDescProd.Caption = ""
    End If
    txtCODPROD.SetFocus

End Sub

Private Sub CmdSalva_Click()

    If Valida_Campos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADOPCOMP.CODIGO = objBLBFunc.Gera_Codigo(Trim(Me.Name) & strNOMETABELA, FILIAL, Linha) & Year(Now)
    
    objCADOPCOMP.DATAOP = "'" & Format(CDate(mskDTOP.Text), "MM/DD/YYYY") & "'"
    objCADOPCOMP.QTDOP = CLng(txtQTDEOP.Text)
    objCADOPCOMP.CODCAPACID = CLng(txtLinProd.Text)
    objCADOPCOMP.IDPRODUTO = CLng(txtCODPROD.Tag)
    
    objCADOPCOMP.OBS = "Null"
    If Len(Trim(txtOBS.Text)) > 0 Then objCADOPCOMP.OBS = "'" & Replace(txtOBS.Text, ",", " ") & "'"

    objCADOPCOMP.CODUSUARIO = lngCODUSUARIO
    objCADOPCOMP.NOMUSUARIO = "'" & strUsuario & "'"
    
    objCADOPCOMP.DTCRIA = "'" & Format(Now, "MM/DD/YYYY HH:MM:SS") & "'"
    
    If cTipOper = "I" Then objCADOPCOMP.STATUS = 0

    If objCADOPCOMP.GRAVA(cTipOper, strNOMETABELA) = False Then Exit Sub

    MsgBox "A OP [ " & objCADOPCOMP.CODIGO & " ] foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
    If cTipOper = "I" Then Unload Me
    If cTipOper = "A" Then
        Call ConfGridInsumos
        Call CarregaCampos
    End If


End Sub

Private Sub Command2_Click()

    If Len(Trim(txtLinProd.Text)) = 0 Then
        MsgBox "ATENÇÃO - Primeiro Informe a Capacidade Produtiva !!!"
        Exit Sub
    End If
    
    ReDim arrCAMPOS(1 To 5, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PROD.* " & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO LINHA " & vbCrLf
    sSql = sSql & "      ,SGI_PRODUSDLINHA    PRDLIN" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO      PROD" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       LINHA.SGI_FILIAL       = " & FILIAL & vbCrLf
    sSql = sSql & "   And LINHA.SGI_CODLIN       = " & Trim(txtLinProd.Text) & vbCrLf
    sSql = sSql & "   And PRDLIN.SGI_FILIAL      = LINHA.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PRDLIN.SGI_CODIGO      = LINHA.SGI_CODIGO" & vbCrLf
    sSql = sSql & "   And PROD.SGI_FILIAL        = PRDLIN.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PROD.SGI_IDPRODUTO     = PRDLIN.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "   And PROD.SGI_PRODUTOTIPO   = 0"
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_IDPRODUTO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "ID"
    arrCAMPOS(1, 4) = "1000"
    arrCAMPOS(1, 5) = "PROD.SGI_IDPRODUTO"
    
    arrCAMPOS(2, 1) = "SGI_CODIGO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Código"
    arrCAMPOS(2, 4) = "1500"
    arrCAMPOS(2, 5) = "PROD.SGI_CODIGO"
    
    arrCAMPOS(3, 1) = "SGI_DESCRICAO"
    arrCAMPOS(3, 2) = "S"
    arrCAMPOS(3, 3) = "Descrição"
    arrCAMPOS(3, 4) = "4000"
    arrCAMPOS(3, 5) = "PROD.SGI_DESCRICAO"
    
    arrCAMPOS(4, 1) = "SGI_CODPRODFORN"
    arrCAMPOS(4, 2) = "S"
    arrCAMPOS(4, 3) = "Cód.Fornecedor"
    arrCAMPOS(4, 4) = "1500"
    arrCAMPOS(4, 5) = "PROD.SGI_CODPRODFORN"
    
    arrCAMPOS(5, 1) = "SGI_COMPLEMENTO"
    arrCAMPOS(5, 2) = "S"
    arrCAMPOS(5, 3) = "Complemento"
    arrCAMPOS(5, 4) = "2000"
    arrCAMPOS(5, 5) = "PROD.SGI_COMPLEMENTO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Produtos")
    
    If Len(Trim(varRETORNO)) > 0 Then
       txtCODPROD.Tag = varRETORNO
       lblDescProd.Caption = PegaDescProduto(varRETORNO, 1)
       txtCODPROD.Text = PegaCodProduto(varRETORNO, 1)
       
       Call PegaDimCorte(varRETORNO)
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Form_Load()

   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADOPCOMP = CreateObject("CADOPCOMP.clsCADOPCOMP")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
   objCADOPCOMP.FILIAL = FILIAL
   
   If adoBanco_Dados.State = 0 Then Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)
   If adoBanco_Dados.State = 0 Then
      MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
      Exit Sub
   End If
   
   strNOMEFILIAL = "NOVALATA "
   strNOMETABELA = ""
   If intFILIALPED = 1 Then
        strNOMEFILIAL = "STEEL-ROL "
        strNOMETABELA = "_STEEL"
   End If
   
   strCAPTION = Me.Caption & " - " & strNOMEFILIAL
   
   Call IniciaForm
   
End Sub

Private Sub DestroiObjetos()
    Set objBLBFunc = Nothing
    Set objCADOPCOMP = Nothing
    Set objPESQPADRAO = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DestroiObjetos
End Sub

Private Sub IniciaForm()
    
    Call objBLBFunc.MudaBotoes(CmdSalva, cmdAltera, cTipOper)
    Call objBLBFunc.MudaCaption(Me, strCAPTION, cTipOper)
    Call objBLBFunc.LimpaCampos(Me)
    
    Call DesabilitaCampos(Trim(cTipOper))
    Call LimpaCamposLabel
    
    objCADOPCOMP.CODIGO = iCodigo
    If cTipOper = "I" Then mskDTOP = Format(Now, "DD/MM/YYYY")
    
    Call ConfGridInsumos
    Call CarregaCampos
    Call PegaStatus
    
End Sub


Private Sub DesabilitaCampos(strTipOper As String)
    If strTipOper = "I" Or strTipOper = "A" Then
        Frame2.Enabled = True
    ElseIf strTipOper = "C" Then
        Frame2.Enabled = False
    End If
End Sub

Private Sub LimpaCamposLabel()
    lblLinhProd.Caption = ""
    lblDescProd.Caption = ""
    lblDescStatus.Caption = ""
    lblEmitidoPor.Caption = ""
    lblDTCRIA.Caption = ""
    lblDIMCORTE.Caption = ""
    lblQTDEFOLHAS.Caption = ""
End Sub

Private Sub mskDTOP_GotFocus()
    objBLBFunc.SelecionaCampos mskDTOP.Name, Me
End Sub

Private Sub txtCODPROD_GotFocus()
    objBLBFunc.SelecionaCampos txtCODPROD.Name, Me
End Sub


Private Sub txtLinProd_GotFocus()
    objBLBFunc.SelecionaCampos txtLinProd.Name, Me
End Sub

Private Sub txtLinProd_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtLinProd.Text
End Sub

Private Sub txtLinProd_Validate(Cancel As Boolean)

    Dim I As Integer
    
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

    If cTipOper = "A" Then
        txtCODPROD.Text = ""
        txtCODPROD.Tag = ""
        lblDescProd.Caption = ""
    End If

End Sub

Private Sub txtQTDEOP_GotFocus()
    objBLBFunc.SelecionaCampos txtQTDEOP.Name, Me
End Sub

Private Function PegaDescProduto(strIDPRODUTO As String, intTipo As Integer) As String

    PegaDescProduto = ""

    If Len(Trim(strIDPRODUTO)) = 0 Then Exit Function
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
    
    If intTipo = 1 Then sSql = sSql & "   And SGI_IDPRODUTO = " & strIDPRODUTO
    If intTipo = 2 Then sSql = sSql & "   And SGI_CODIGO    = '" & Trim(strIDPRODUTO) & "'"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PegaDescProduto = Trim(BREC!SGI_DESCRICAO)
    BREC.Close
    
End Function


Private Function PegaCodProduto(strIDPRODUTO As String, intTipo As Integer) As String

    PegaCodProduto = ""

    If Len(Trim(strIDPRODUTO)) = 0 Then Exit Function
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
    
    If intTipo = 1 Then sSql = sSql & "   And SGI_IDPRODUTO = " & strIDPRODUTO
    If intTipo = 2 Then sSql = sSql & "   And SGI_CODIGO    = '" & Trim(strIDPRODUTO) & "'"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        If intTipo = 1 Then PegaCodProduto = Trim(BREC!SGI_CODIGO)
        If intTipo = 2 Then PegaCodProduto = Trim(BREC!SGI_IDPRODUTO)
    End If
    BREC.Close
    
End Function


Private Sub ConfGridInsumos()

    With grdProdInsumos
    
       .Cols = conColumnsIn_OrdFabC
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_OrdFabC_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_OrdFabC_IDProduto) = ""
       .ColDataType(conCOL_OrdFabC_IDProduto) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_OrdFabC_CodProd) = ""
       .ColDataType(conCOL_OrdFabC_CodProd) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_OrdFabC_PesqProd) = ""
       .ColDataType(conCOL_OrdFabC_PesqProd) = flexDTString
       .ColComboList(conCOL_OrdFabC_PesqProd) = "..."
       
       .Cell(flexcpData, 0, conCOL_OrdFabC_DescProd) = ""
       .ColDataType(conCOL_OrdFabC_DescProd) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_OrdFabC_AbilitSN) = ""
       .ColDataType(conCOL_OrdFabC_AbilitSN) = flexDTString
       .ColComboList(conCOL_OrdFabC_AbilitSN) = objCADOPCOMP.PreenchComboAtivo
       
       .ColWidth(conCOL_OrdFabC_IDProduto) = 0
       .ColWidth(conCOL_OrdFabC_CodProd) = 1500
       .ColWidth(conCOL_OrdFabC_PesqProd) = 300
       .ColWidth(conCOL_OrdFabC_DescProd) = 5000
       .ColWidth(conCOL_OrdFabC_AbilitSN) = 1500
       
       .Editable = flexEDNone
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
    
    End With
    
End Sub

Private Function Valida_Campos() As Boolean

On Error GoTo Err_Valida_Campos
     
     Dim lngLINHA   As Long
     Dim I          As Long
     
     Valida_Campos = False
     
     If Not IsDate(mskDTOP.Text) Then
        MsgBox "ATENÇÂO - Data da OP Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
        mskDTOP.SetFocus
        Exit Function
     End If
     If Len(Trim(txtQTDEOP.Text)) = 0 Then
        MsgBox "ATENÇÂO - Informe a quantidade da OP !!!", vbOKOnly + vbExclamation, "Aviso"
        txtQTDEOP.SetFocus
        Exit Function
     End If
     If Len(Trim(txtLinProd.Text)) = 0 Then
        MsgBox "ATENÇÂO - Informe o código da Capacidade !!!", vbOKOnly + vbExclamation, "Aviso"
        txtLinProd.SetFocus
        Exit Function
     End If
     If Len(Trim(txtCODPROD.Text)) = 0 Then
        MsgBox "ATENÇÂO - Informe o código do produto !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCODPROD.SetFocus
        Exit Function
     End If
     If Not IsNumeric(txtQTDEOP.Text) Then
        MsgBox "ATENÇÂO - quantidade da OP Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
        txtQTDEOP.SetFocus
        Exit Function
     End If
     If Not IsNumeric(txtCODPROD.Tag) Then
        MsgBox "ATENÇÂO - código do produto inválido !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCODPROD.SetFocus
        Exit Function
     End If
     
     Valida_Campos = True

    Exit Function
    
Err_Valida_Campos:

    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : Valida_Campos()", Me.Name, "Valida_Campos()", strCAMARQERRO)

End Function


Private Sub CarregaCampos()
    If objCADOPCOMP.Carrega_Campos(strNOMETABELA) = True Then
    
        txtCODOP.Text = objCADOPCOMP.CODIGO
        mskDTOP.Text = objCADOPCOMP.DATAOP
        txtQTDEOP.Text = objCADOPCOMP.QTDOP
        txtCODPROD.Tag = objCADOPCOMP.IDPRODUTO
        txtOBS.Text = objCADOPCOMP.OBS
        lblEmitidoPor.Caption = objCADOPCOMP.NOMUSUARIO
        lblDTCRIA.Caption = objCADOPCOMP.DTCRIA
            
        txtLinProd.Text = objCADOPCOMP.CODCAPACID
        lblLinhProd.Caption = PegaDescLinProd(CLng(txtLinProd.Text))
        
        lblDescProd.Caption = PegaDescProduto(Str(objCADOPCOMP.IDPRODUTO), 1)
        txtCODPROD.Text = PegaCodProduto(Str(objCADOPCOMP.IDPRODUTO), 1)
            
        Call PegaDimCorte(Str(objCADOPCOMP.IDPRODUTO))
    
    End If
End Sub

Private Sub PegaStatus()
    If objCADOPCOMP.STATUS = 0 Then lblDescStatus.Caption = "ABERTA"
End Sub

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
    If Not BREC.EOF Then PegaDescLinProd = BREC!SGI_DESCRI
    BREC.Close
    
End Function


Private Sub PegaDimCorte(strIDPROD As String)


    Dim dblQTDEFOLHA As Double

    
    lblDIMCORTE.Caption = ""
    lblQTDEFOLHAS.Caption = ""
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       PROD.SGI_DIMPADRAO" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_DESENV" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_ALTURA" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_QTDEPORFOLHA" & vbCrLf
    
    sSql = sSql & "      ,LINH.SGI_DESENV As SGI_VALDESENV" & vbCrLf
    sSql = sSql & "      ,LINH.SGI_ALTURA As SGI_VALALTURA" & vbCrLf
    sSql = sSql & "      ,LINH.SGI_PERDPROC" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO      PROD" & vbCrLf
    sSql = sSql & "      ,SGI_PRODUSDLINHA    LINP" & vbCrLf
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO LINH" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       PROD.SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And PROD.SGI_IDPRODUTO = " & strIDPROD & vbCrLf
    sSql = sSql & "   And LINP.SGI_FILIAL    = PROD.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And LINP.SGI_IDPRODUTO = PROD.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "   And LINH.SGI_FILIAL    = LINP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And LINH.SGI_CODIGO    = LINP.SGI_CODIGO"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        If BREC!SGI_DIMPADRAO = 1 Then      '' Padrão da Linha
            lblDIMCORTE.Caption = Format(BREC!SGI_VALDESENV, "#,###0.000") & " x " & Format(BREC!SGI_VALALTURA, "#,###0.000")
        ElseIf BREC!SGI_DIMPADRAO = 0 Then  '' Não é Padrão
            lblDIMCORTE.Caption = Format(BREC!SGI_DESENV, "#,###0.000") & " x " & Format(BREC!SGI_ALTURA, "#,###0.000")
        End If
        
        dblQTDEFOLHA = 0
        If Not IsNull(BREC!SGI_QTDEPORFOLHA) And Not IsNull(BREC!SGI_PERDPROC) Then
            dblQTDEFOLHA = ((CDbl(txtQTDEOP.Text) / BREC!SGI_QTDEPORFOLHA) * BREC!SGI_PERDPROC)
        End If
        If dblQTDEFOLHA > 0 Then lblQTDEFOLHAS.Caption = Format(dblQTDEFOLHA, "#0")
        
    End If
    BREC.Close


End Sub
