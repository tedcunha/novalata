VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCADCLASSFISC 
   Caption         =   "Cadastro de classificação fiscal"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   6000
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   0
      TabIndex        =   13
      Top             =   2760
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6165
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
      TabCaption(0)   =   "Familia de Produtos"
      TabPicture(0)   =   "frmCADCLASSFISC.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grdFamProdutos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdIncUnidConv"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdexcUnidConv"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Natureza de Operação"
      TabPicture(1)   =   "frmCADCLASSFISC.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "grdNATOPER"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.CommandButton Command2 
         Height          =   300
         Left            =   -69480
         Picture         =   "frmCADCLASSFISC.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   720
         Width           =   300
      End
      Begin VB.CommandButton Command1 
         Height          =   300
         Left            =   -69480
         Picture         =   "frmCADCLASSFISC.frx":0182
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   360
         Width           =   300
      End
      Begin VSFlex8LCtl.VSFlexGrid grdNATOPER 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   17
         Top             =   360
         Width           =   5295
         _cx             =   9340
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
      Begin VB.CommandButton cmdexcUnidConv 
         Height          =   300
         Left            =   5520
         Picture         =   "frmCADCLASSFISC.frx":02CC
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   720
         Width           =   300
      End
      Begin VB.CommandButton cmdIncUnidConv 
         Height          =   300
         Left            =   5520
         Picture         =   "frmCADCLASSFISC.frx":0416
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   300
      End
      Begin VSFlex8LCtl.VSFlexGrid grdFamProdutos 
         Height          =   3015
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   5295
         _cx             =   9340
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
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   5895
      Begin VB.TextBox txtAliqII 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3840
         TabIndex        =   27
         Text            =   "txtAliquota"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtIPITRANSF 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   25
         Text            =   "txtIPITRANSF"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2640
         TabIndex        =   20
         Top             =   960
         Width           =   3015
         Begin VB.OptionButton optTEMSTSN 
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
            Left            =   1200
            TabIndex        =   23
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton optTEMSTSN 
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
            Left            =   2040
            TabIndex        =   22
            Top             =   0
            Width           =   735
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tem ST"
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
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Width           =   675
         End
      End
      Begin VB.TextBox txtAliquota 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   12
         Text            =   "txtAliquota"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "txtCodigo"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtLetra 
         Height          =   285
         Left            =   1200
         MaxLength       =   1
         TabIndex        =   1
         Text            =   "txtLetra"
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox txtNomeclatura 
         Height          =   285
         Left            =   3840
         MaxLength       =   15
         TabIndex        =   2
         Text            =   "txtNomeclatura"
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Aliq II"
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
         Left            =   2640
         TabIndex        =   26
         Top             =   1320
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "IPI Transf."
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
         Index           =   2
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Aliquota IPI"
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
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ID"
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
         Left            =   180
         TabIndex        =   10
         Top             =   240
         Width           =   210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Letra"
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
         Left            =   150
         TabIndex        =   9
         Top             =   645
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nomeclatura"
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
         Left            =   2640
         TabIndex        =   8
         Top             =   645
         Width           =   1080
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5895
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
         Picture         =   "frmCADCLASSFISC.frx":0560
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmCADCLASSFISC.frx":0A92
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "frmCADCLASSFISC.frx":0B94
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADCLASSFISC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho      As String
Public Linha         As Variant
Public cTipOper      As String
Public iCodigo       As Integer
Public FILIAL        As Integer
Public strAcesso     As String
Public strMODPAI     As String
Dim objBLBFunc       As Object
Dim objCADCLASSFISC  As Object
Dim objPESQPADRAO    As Object
Dim arrFAMPRODUTOS   As Variant
Dim arrNATOPER       As Variant

'' ========================================================================================
Const conCOL_SonClaFis_Codigo                   As Integer = 0
Const conCOL_SonClaFis_PesqFam                  As Integer = 1
Const conCOL_SonClaFis_DescFam                  As Integer = 2

Const conCOL_SonClaFis_FormatString             As String = "=Código|...|Descrição"
Const conColumnsIn_SonClaFis                    As Integer = 3

'' ========================================================================================
Const conCOL_SonNatOper_Codigo                  As Integer = 0
Const conCOL_SonNatOper_Pesq                    As Integer = 1
Const conCOL_SonNatOper_Desc                    As Integer = 2
Const conCOL_SonNatOper_CFOP                    As Integer = 3

Const conCOL_SonNatOper_FormatString            As String = "=Código|...|Descrição|CFOP"
Const conColumnsIn_SonNatOper                   As Integer = 4


Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de classificação fiscal - [ ALTERACAO ]"
    cTipOper = "A"
    
    txtLetra.SetFocus

End Sub

Private Sub cmdexcUnidConv_Click()
    Call objBLBFunc.ExclLinhaGrid(grdFamProdutos, grdFamProdutos.Row)
End Sub

Private Sub cmdIncUnidConv_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncRegGridFamProdutos
End Sub

Private Sub CmdSalva_Click()

    Dim I        As Integer
    Dim strValor As String
    
    If ValidaCampos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADCLASSFISC.CODIGO = objCADCLASSFISC.Gera_Codigo(Me.Name)
    objCADCLASSFISC.LETRA = txtLetra.Text
    objCADCLASSFISC.NOMECLA = txtNomeclatura.Text
    If optTEMSTSN(1).Value = True Then objCADCLASSFISC.TEMST = -1   '' Sim
    If optTEMSTSN(0).Value = True Then objCADCLASSFISC.TEMST = 0    '' Não
    
    If Len(Trim(txtAliquota.Text)) > 0 Then objCADCLASSFISC.ALIQUOTA = CCur(txtAliquota.Text)
    
    objCADCLASSFISC.ALIQII = "Null"
    If Len(Trim(txtAliqII.Text)) > 0 Then
        strValor = Replace(txtAliqII.Text, ".", "")
        strValor = Replace(strValor, ",", ".")
        objCADCLASSFISC.ALIQII = strValor
    End If
    
    objCADCLASSFISC.IPITRANSF = "Null"
    If Len(Trim(txtIPITRANSF.Text)) > 0 Then
        strValor = Replace(txtIPITRANSF.Text, ".", "")
        strValor = Replace(strValor, ",", ".")
        objCADCLASSFISC.IPITRANSF = strValor
    End If
    
    '' ---------------------------
    '' Familia de Produtos
    Call objBLBFunc.RemoveLinhaVazia(grdFamProdutos, conCOL_SonClaFis_Codigo)
    
    arrFAMPRODUTOS = Empty
    If (grdFamProdutos.Rows - 1) > 0 Then
        ReDim arrFAMPRODUTOS(1 To (grdFamProdutos.Rows - 1)) As String
        For I = 1 To (grdFamProdutos.Rows - 1)
            arrFAMPRODUTOS(I) = grdFamProdutos.Cell(flexcpText, I, conCOL_SonClaFis_Codigo)
        Next I
    End If
    objCADCLASSFISC.FAMPROD = arrFAMPRODUTOS
    '' ---------------------------
    
    '' ---------------------------
    '' Natureza de Operação
    Call objBLBFunc.RemoveLinhaVazia(grdNATOPER, conCOL_SonNatOper_Codigo)
    
    arrNATOPER = Empty
    With grdNATOPER
        If (.Rows - 1) > 0 Then
            ReDim arrNATOPER(1 To (.Rows - 1)) As String
            For I = 1 To (.Rows - 1)
                arrNATOPER(I) = .Cell(flexcpText, I, conCOL_SonNatOper_Codigo)
            Next I
        End If
    End With
    objCADCLASSFISC.NATOPER = arrNATOPER
    '' ---------------------------
    
    If objCADCLASSFISC.GRAVA(cTipOper) = False Then Exit Sub
          
    MsgBox "A Classificação foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
    If objCADCLASSFISC.Atualiza(cTipOper, Str(objCADCLASSFISC.CODIGO), FILIAL, Me.Name) = False Then Exit Sub
    
    If cTipOper = "I" Then
       Set objBLBFunc = Nothing
       Set objCADCLASSFISC = Nothing
       Unload Me
    End If

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objCADCLASSFISC = Nothing
    Set objPESQPADRAO = Nothing
    Unload Me
End Sub

Private Sub Command1_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncRegGridNatOper
End Sub

Private Sub Command2_Click()
    Call objBLBFunc.ExclLinhaGrid(grdNATOPER, grdNATOPER.Row)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADCLASSFISC = CreateObject("CADCLASSFISC.clsCADCLASSFISC")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
      
   objCADCLASSFISC.FILIAL = FILIAL
   
   Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

   If adoBanco_Dados.State = 0 Then
      MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
      Exit Sub
   End If
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

End Sub


Private Sub grdFamProdutos_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonClaFis_DescFam
         Cancel = True
    Case conCOL_SonClaFis_Codigo, _
         conCOL_SonClaFis_PesqFam
         If cTipOper = "C" Then Cancel = True
    Case Else
        grdFamProdutos.ComboList = ""
    End Select
    Exit Sub
End Sub


Private Sub grdFamProdutos_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    If cTipOper = "C" Then Exit Sub
    
    With grdFamProdutos
         If (.Rows - 1) = 0 Then Exit Sub
        
         ReDim arrCAMPOS(1 To 2, 1 To 5) As String
         ReDim arrTABELA(1 To 1) As String
        
         Select Case Col
                Case conCOL_SonClaFis_PesqFam
                    
                     sSql = "Select " & vbCrLf
                     sSql = sSql & "       * " & vbCrLf
                     sSql = sSql & "  From " & vbCrLf
                     sSql = sSql & "       SGI_CADGRUPROD " & vbCrLf
                     sSql = sSql & " Where " & vbCrLf
                     sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
                     
                     arrTABELA(1) = sSql
                     
                     arrCAMPOS(1, 1) = "SGI_CODIGO"
                     arrCAMPOS(1, 2) = "N"
                     arrCAMPOS(1, 3) = "Código"
                     arrCAMPOS(1, 4) = "1000"
                     arrCAMPOS(1, 5) = "SGI_CODIGO"
                    
                     arrCAMPOS(2, 1) = "SGI_DESCRICAO"
                     arrCAMPOS(2, 2) = "S"
                     arrCAMPOS(2, 3) = "Descrição"
                     arrCAMPOS(2, 4) = "2500"
                     arrCAMPOS(2, 5) = "SGI_DESCRICAO"
                     
                     varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Familia de Produtos")
                    
                     If Len(Trim(varRETORNO)) > 0 Then
                        
                        If objBLBFunc.FcVerifItensRepetidos(grdFamProdutos, Row, conCOL_SonClaFis_Codigo, varRETORNO) = False Then
                           MsgBox "A Familia de Produto Já está relacionada na Gride !!!", vbOKOnly + vbExclamation, "Aviso"
                           Exit Sub
                        End If

                        .Cell(flexcpText, Row, conCOL_SonClaFis_Codigo) = varRETORNO
                        .Cell(flexcpText, Row, conCOL_SonClaFis_DescFam) = Trim(PegaDescFamProd(varRETORNO))
                        
                     End If
                    
         End Select
    End With

End Sub

Private Sub grdFamProdutos_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdFamProdutos
          Select Case Col
                    Case conCOL_SonClaFis_Codigo
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
          End Select
     End With
End Sub

Private Sub grdFamProdutos_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

     With grdFamProdutos
          Select Case Col
                 Case conCOL_SonClaFis_Codigo
                        If .EditText = Empty Then Exit Sub
                        If objBLBFunc.FcVerifItensRepetidos(grdFamProdutos, Row, conCOL_SonClaFis_Codigo, .EditText) = False Then
                           MsgBox "A Familia de Produto Já está relacionada na Gride !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, conCOL_SonClaFis_Codigo) = Empty
                           .Cell(flexcpText, Row, conCOL_SonClaFis_DescFam) = Empty
                           Cancel = True
                           Exit Sub
                        End If
                        
                        If Len(Trim(PegaDescFamProd(.EditText))) = 0 Then
                           MsgBox "Esta Familia de Produto não existe !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, conCOL_SonClaFis_Codigo) = Empty
                           .Cell(flexcpText, Row, conCOL_SonClaFis_DescFam) = Empty
                           Cancel = True
                           Exit Sub
                        End If
                        .Cell(flexcpText, Row, conCOL_SonClaFis_Codigo) = .EditText
                        .Cell(flexcpText, Row, conCOL_SonClaFis_DescFam) = PegaDescFamProd(.EditText)
          End Select
     End With

End Sub

Private Sub grdNATOPER_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case conCOL_SonNatOper_Desc, _
         conCOL_SonNatOper_CFOP
         Cancel = True
    Case conCOL_SonNatOper_Codigo, _
         conCOL_SonNatOper_Pesq
         If cTipOper = "C" Then Cancel = True
    Case Else
        grdNATOPER.ComboList = ""
    End Select
    Exit Sub
End Sub


Private Sub grdNATOPER_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    If cTipOper = "C" Then Exit Sub
    
    With grdNATOPER
         If (.Rows - 1) = 0 Then Exit Sub
        
         ReDim arrCAMPOS(1 To 3, 1 To 5) As String
         ReDim arrTABELA(1 To 1) As String
        
         Select Case Col
                Case conCOL_SonClaFis_PesqFam
                    
                     sSql = "Select " & vbCrLf
                     sSql = sSql & "       * " & vbCrLf
                     sSql = sSql & "  From " & vbCrLf
                     sSql = sSql & "       SGI_CADNATOPERACAO " & vbCrLf
                     sSql = sSql & " Where " & vbCrLf
                     sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
                     
                     arrTABELA(1) = sSql
                     
                     arrCAMPOS(1, 1) = "SGI_CODIGO"
                     arrCAMPOS(1, 2) = "N"
                     arrCAMPOS(1, 3) = "Código"
                     arrCAMPOS(1, 4) = "1000"
                     arrCAMPOS(1, 5) = "SGI_CODIGO"
                    
                     arrCAMPOS(2, 1) = "SGI_NOMECLCOD"
                     arrCAMPOS(2, 2) = "S"
                     arrCAMPOS(2, 3) = "CFOP"
                     arrCAMPOS(2, 4) = "1000"
                     arrCAMPOS(2, 5) = "SGI_NOMECLCOD"
                    
                     arrCAMPOS(3, 1) = "SGI_DESCRICAO"
                     arrCAMPOS(3, 2) = "S"
                     arrCAMPOS(3, 3) = "Descrição"
                     arrCAMPOS(3, 4) = "2500"
                     arrCAMPOS(3, 5) = "SGI_DESCRICAO"
                     
                     varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strAcesso, V_Usuario, arrCAMPOS, arrTABELA, "Natureza de Operação")
                    
                     If Len(Trim(varRETORNO)) > 0 Then
                        
                        If objBLBFunc.FcVerifItensRepetidos(grdNATOPER, Row, conCOL_SonNatOper_Codigo, varRETORNO) = False Then
                           MsgBox "A Natureza de Operação Já está relacionada na Gride !!!", vbOKOnly + vbExclamation, "Aviso"
                           Exit Sub
                        End If

                        .Cell(flexcpText, Row, conCOL_SonNatOper_Codigo) = varRETORNO
                        .Cell(flexcpText, Row, conCOL_SonNatOper_Desc) = Trim(PegaDescNatOper(varRETORNO, Row))
                        
                     End If
                    
         End Select
    End With

End Sub

Private Sub grdNATOPER_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdNATOPER
          Select Case Col
                    Case conCOL_SonNatOper_Codigo
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
          End Select
     End With
End Sub

Private Sub grdNATOPER_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

     With grdNATOPER
          Select Case Col
                 Case conCOL_SonClaFis_Codigo
                        If .EditText = Empty Then Exit Sub
                        If objBLBFunc.FcVerifItensRepetidos(grdNATOPER, Row, conCOL_SonNatOper_Codigo, .EditText) = False Then
                           MsgBox "A Natureza de Operação Já está relacionada na Gride !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, conCOL_SonNatOper_Codigo) = Empty
                           .Cell(flexcpText, Row, conCOL_SonNatOper_Desc) = Empty
                           Cancel = True
                           Exit Sub
                        End If
                        
                        If Len(Trim(PegaDescNatOper(.EditText, Row))) = 0 Then
                           MsgBox "Esta natureza de operação não existe !!!", vbOKOnly + vbExclamation, "Aviso"
                           .Cell(flexcpText, Row, conCOL_SonNatOper_Codigo) = Empty
                           .Cell(flexcpText, Row, conCOL_SonNatOper_Desc) = Empty
                           Cancel = True
                           Exit Sub
                        End If
                        .Cell(flexcpText, Row, conCOL_SonNatOper_Codigo) = .EditText
                        .Cell(flexcpText, Row, conCOL_SonNatOper_Desc) = PegaDescNatOper(.EditText, Row)
          End Select
     End With

End Sub

Private Sub txtAliqII_GotFocus()
    objBLBFunc.SelecionaCampos txtAliqII.Name, frmCADCLASSFISC
End Sub

Private Sub txtAliqII_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.SoNumeroPonto(KeyAscii, txtAliqII.Text)
End Sub

Private Sub txtAliqII_Validate(Cancel As Boolean)
    If Len(Trim(txtAliqII.Text)) = 0 Then Exit Sub
    txtAliqII.Text = Format(txtAliqII.Text, "#,##0.00")
End Sub

Private Sub txtAliquota_GotFocus()
    objBLBFunc.SelecionaCampos txtAliquota.Name, frmCADCLASSFISC
End Sub

Private Sub txtAliquota_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.SoNumeroPonto(KeyAscii, txtAliquota.Text)
End Sub

Private Sub txtAliquota_Validate(Cancel As Boolean)
    If Len(Trim(txtAliquota.Text)) = 0 Then Exit Sub
    txtAliquota.Text = Format(txtAliquota.Text, "#,##0.00")
End Sub

Private Sub txtIPITRANSF_GotFocus()
    objBLBFunc.SelecionaCampos txtIPITRANSF.Name, frmCADCLASSFISC
End Sub

Private Sub txtIPITRANSF_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.SoNumeroPonto(KeyAscii, txtIPITRANSF.Text)
End Sub

Private Sub txtIPITRANSF_Validate(Cancel As Boolean)
    If Len(Trim(txtAliquota.Text)) = 0 Then Exit Sub
    txtIPITRANSF.Text = Format(txtIPITRANSF.Text, "#,##0.00")
End Sub

Private Sub txtLetra_GotFocus()
    objBLBFunc.SelecionaCampos txtLetra.Name, frmCADCLASSFISC
End Sub

Private Sub txtLetra_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub txtNomeclatura_GotFocus()
    objBLBFunc.SelecionaCampos txtNomeclatura.Name, frmCADCLASSFISC
End Sub

Private Sub txtNomeclatura_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
   
    Me.Caption = "Cadastro de classificação fiscal - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADCLASSFISC
    
    txtCodigo.Text = ""
   
    Call InitGridProd
    Call InitGridNatOper
    
    optTEMSTSN(0).Value = True
    
End Sub

Private Function ValidaCampos() As Boolean

     ValidaCampos = False
     
     If Trim(Len(txtLetra.Text)) = 0 Then
        MsgBox "Letra não pode ser nulo !!!", vbOKOnly + vbCritical, "Aviso"
        txtLetra.SetFocus
        Exit Function
     End If
     
     If Trim(Len(txtNomeclatura.Text)) = 0 Then
        MsgBox "Nomeclatura não pode ser nulo !!!", vbOKOnly + vbCritical, "Aviso"
        txtNomeclatura.SetFocus
        Exit Function
     End If
     
     If cTipOper = "I" Then
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       * " & vbCrLf
        sSql = sSql & "  from " & vbCrLf
        sSql = sSql & "       SGI_CADCLASSFIS  " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_CODCLASS = '" & txtLetra.Text & "'" & vbCrLf
        sSql = sSql & "   And SGI_FILIAL = " & FILIAL
        
        BREC.Open sSql, adoBanco_Dados
        
        If Not BREC.EOF Then
           MsgBox "Está letra já existe !!!", vbOKOnly + vbCritical, "Aviso"
           txtLetra.SetFocus
           BREC.Close
           Exit Function
        End If
        
        BREC.Close
     
     
        sSql = "Select " & vbCrLf
        sSql = sSql & "       * " & vbCrLf
        sSql = sSql & "  from " & vbCrLf
        sSql = sSql & "       SGI_CADCLASSFIS  " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_NOMECLA = '" & txtNomeclatura.Text & "'" & vbCrLf
        sSql = sSql & "   And SGI_FILIAL = " & FILIAL
        
        BREC.Open sSql, adoBanco_Dados
        
        If Not BREC.EOF Then
           MsgBox "Está nomeclatura já existe !!!", vbOKOnly + vbCritical, "Aviso"
           txtNomeclatura.SetFocus
           BREC.Close
           Exit Function
        End If
        
        BREC.Close
     
     End If
     
     If cTipOper = "A" Then
     
        If objCADCLASSFISC.LETRA <> txtLetra.Text Then
        
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  from " & vbCrLf
           sSql = sSql & "       SGI_CADCLASSFIS  " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_CODCLASS = '" & txtLetra.Text & "'" & vbCrLf
           sSql = sSql & "   And SGI_FILIAL   = " & FILIAL
           
           BREC.Open sSql, adoBanco_Dados
        
           If Not BREC.EOF Then
              MsgBox "Este banco já existe !!!", vbOKOnly + vbCritical, "Aviso"
              txtLetra.Text = objCADCLASSFISC.LETRA
              txtLetra.SetFocus
              BREC.Close
              Exit Function
           End If
        
           BREC.Close
        
        End If
        
        If objCADCLASSFISC.NOMECLA <> txtNomeclatura.Text Then
        
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  from " & vbCrLf
           sSql = sSql & "       SGI_CADCLASSFIS  " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_NOMECLA = '" & txtNomeclatura.Text & "'" & vbCrLf
           sSql = sSql & "   And SGI_FILIAL   = " & FILIAL
           
           BREC.Open sSql, adoBanco_Dados
        
           If Not BREC.EOF Then
              MsgBox "Esta nomeclatura já existe !!!", vbOKOnly + vbCritical, "Aviso"
              txtNomeclatura.Text = objCADCLASSFISC.NOMECLA
              txtNomeclatura.SetFocus
              BREC.Close
              Exit Function
           End If
        
           BREC.Close
        
        End If
     
     End If
     
     ValidaCampos = True
     
End Function

Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
   
    Me.Caption = "Cadastro de classificação fiscal - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADCLASSFISC
    
    objCADCLASSFISC.CODIGO = iCodigo
    
    Call InitGridProd
    Call InitGridNatOper
    
    If objCADCLASSFISC.Carrega_campos = True Then
       txtCodigo.Text = Str(objCADCLASSFISC.CODIGO)
       txtLetra.Text = objCADCLASSFISC.LETRA
       txtNomeclatura.Text = objCADCLASSFISC.NOMECLA
       
       If objCADCLASSFISC.TEMST = 0 Then optTEMSTSN(0).Value = True
       If objCADCLASSFISC.TEMST = -1 Then optTEMSTSN(1).Value = True
    
       If objCADCLASSFISC.ALIQUOTA > 0 Then txtAliquota.Text = Format(objCADCLASSFISC.ALIQUOTA, "#,##0.00")
       If Len(Trim(objCADCLASSFISC.IPITRANSF)) > 0 Then txtIPITRANSF.Text = objCADCLASSFISC.IPITRANSF
       If Len(Trim(objCADCLASSFISC.ALIQII)) > 0 Then txtAliqII.Text = objCADCLASSFISC.ALIQII
       
       Call CarregaFamiliaProdutos
       Call CarregaNatOpers
    End If

End Sub

Public Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    
    Me.Caption = "Cadastro de classificação fiscal - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADCLASSFISC
        
    objCADCLASSFISC.CODIGO = iCodigo
    
    Call InitGridProd
    Call InitGridNatOper
    
    If objCADCLASSFISC.Carrega_campos = True Then
       txtCodigo.Text = Str(objCADCLASSFISC.CODIGO)
       txtLetra.Text = objCADCLASSFISC.LETRA
       txtNomeclatura.Text = objCADCLASSFISC.NOMECLA
       
       If objCADCLASSFISC.TEMST = 0 Then optTEMSTSN(0).Value = True
       If objCADCLASSFISC.TEMST = -1 Then optTEMSTSN(1).Value = True
    
       If objCADCLASSFISC.ALIQUOTA > 0 Then txtAliquota.Text = Format(objCADCLASSFISC.ALIQUOTA, "#,##0.00")
       If Len(Trim(objCADCLASSFISC.IPITRANSF)) > 0 Then txtIPITRANSF.Text = objCADCLASSFISC.IPITRANSF
       If Len(Trim(objCADCLASSFISC.ALIQII)) > 0 Then txtAliqII.Text = objCADCLASSFISC.ALIQII
       
       Call CarregaFamiliaProdutos
       Call CarregaNatOpers
    End If
    
End Sub

Private Sub InitGridProd()

    With grdFamProdutos
    
       .Cols = conColumnsIn_SonClaFis
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonClaFis_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonClaFis_Codigo) = ""
       .ColDataType(conCOL_SonClaFis_Codigo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonClaFis_PesqFam) = ""
       .ColDataType(conCOL_SonClaFis_PesqFam) = flexDTString
       .ColComboList(conCOL_SonClaFis_PesqFam) = "..."
       
       .Cell(flexcpData, 0, conCOL_SonClaFis_DescFam) = ""
       .ColDataType(conCOL_SonClaFis_DescFam) = flexDTString
       
       .ColWidth(conCOL_SonClaFis_Codigo) = 1000
       .ColWidth(conCOL_SonClaFis_PesqFam) = 300
       .ColWidth(conCOL_SonClaFis_DescFam) = 3500
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
       
    End With
    
End Sub

Private Sub IncRegGridFamProdutos()
   
    If objBLBFunc.TemLinhaVazia(grdFamProdutos, conCOL_SonClaFis_Codigo) = True Then Exit Sub
    
    grdFamProdutos.AddItem "" & vbTab & _
                           "" & vbTab & _
                           ""
                          
End Sub

Private Function PegaDescFamProd(strCodigo As String) As String

    PegaDescFamProd = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADGRUPROD " & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & strCodigo
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PegaDescFamProd = Trim(BREC!SGI_DESCRICAO)
    BREC.Close
    
End Function

Private Sub CarregaFamiliaProdutos()

    Dim I               As Integer
    Dim strDescGrdProd  As String
    
    arrFAMPRODUTOS = objCADCLASSFISC.FAMPROD
    
    If IsArray(arrFAMPRODUTOS) Then
        For I = 1 To UBound(arrFAMPRODUTOS)
        
            strDescGrdProd = PegaDescFamProd(Str(arrFAMPRODUTOS(I)))
            
            grdFamProdutos.AddItem arrFAMPRODUTOS(I) & vbTab & _
                                   "" & vbTab & _
                                   Trim(strDescGrdProd)
        Next I
    End If

End Sub


Private Sub InitGridNatOper()

    With grdNATOPER
    
       .Cols = conColumnsIn_SonNatOper
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_SonNatOper_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeBoth
       
       .Cell(flexcpData, 0, conCOL_SonNatOper_Codigo) = ""
       .ColDataType(conCOL_SonNatOper_Codigo) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_SonNatOper_Pesq) = ""
       .ColDataType(conCOL_SonNatOper_Pesq) = flexDTString
       .ColComboList(conCOL_SonNatOper_Pesq) = "..."
       
       .Cell(flexcpData, 0, conCOL_SonNatOper_Desc) = ""
       .ColDataType(conCOL_SonNatOper_Desc) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_SonNatOper_CFOP) = ""
       .ColDataType(conCOL_SonNatOper_CFOP) = flexDTString
       
       .ColWidth(conCOL_SonNatOper_Codigo) = 800
       .ColWidth(conCOL_SonNatOper_Pesq) = 300
       .ColWidth(conCOL_SonNatOper_Desc) = 2000
       .ColWidth(conCOL_SonNatOper_CFOP) = 1000
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
       
    End With
    
End Sub


Private Function PegaDescNatOper(strCodigo As String, lngRow As Long) As String

    PegaDescNatOper = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADNATOPERACAO " & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & strCodigo
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
       PegaDescNatOper = Trim(BREC!SGI_DESCRICAO)
       grdNATOPER.Cell(flexcpText, lngRow, conCOL_SonNatOper_CFOP) = BREC!SGI_NOMECLCOD
    End If
    BREC.Close
    
End Function


Private Sub IncRegGridNatOper()
   
    If objBLBFunc.TemLinhaVazia(grdNATOPER, conCOL_SonNatOper_Codigo) = True Then Exit Sub
    
    grdNATOPER.AddItem "" & vbTab & _
                       "" & vbTab & _
                       "" & vbTab & _
                       ""
                          
End Sub



Private Sub CarregaNatOpers()

    Dim I               As Integer
    Dim strDescGrdProd  As String
    
    arrNATOPER = objCADCLASSFISC.NATOPER
    
    If IsArray(arrNATOPER) Then
        With grdNATOPER
            For I = 1 To UBound(arrNATOPER)
                
                .AddItem arrNATOPER(I) & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         ""
                
                strDescGrdProd = PegaDescNatOper(Str(arrNATOPER(I)), (.Rows - 1))
                .Cell(flexcpText, (.Rows - 1), conCOL_SonNatOper_Desc) = Trim(strDescGrdProd)
            
            Next I
        End With
    End If

End Sub

