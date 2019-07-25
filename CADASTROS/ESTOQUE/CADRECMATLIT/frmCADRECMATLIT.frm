VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCADRECMATLIT 
   Caption         =   "Recebimento de Material (Folhas Litografadas)"
   ClientHeight    =   8355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18915
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   18915
   StartUpPosition =   1  'CenterOwner
   Begin VSFlex8LCtl.VSFlexGrid grdLANCTOS 
      Height          =   5775
      Left            =   0
      TabIndex        =   18
      Top             =   2400
      Width           =   18855
      _cx             =   33258
      _cy             =   10186
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
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   0
      TabIndex        =   8
      Top             =   960
      Width           =   18855
      Begin VB.TextBox txtCODENV 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6000
         TabIndex        =   2
         Text            =   "txtCODENV"
         Top             =   240
         Width           =   1215
      End
      Begin MSMask.MaskEdBox mskDTENTRADA 
         Height          =   285
         Left            =   3120
         TabIndex        =   1
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
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "txtCodigo"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtCIDCLIE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   10
         Text            =   "txtCIDCLIE"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtCLIEDEST 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   9
         Text            =   "txtCLIEDES"
         Top             =   960
         Width           =   1215
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   735
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
         Index           =   1
         Left            =   2520
         TabIndex        =   16
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente Origem"
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
         Left            =   4440
         TabIndex        =   15
         Top             =   600
         Width           =   1245
      End
      Begin VB.Label lblDescCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescCliente"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7245
         TabIndex        =   14
         Top             =   600
         Width           =   8055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente Destino"
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
         Left            =   4440
         TabIndex        =   13
         Top             =   960
         Width           =   1305
      End
      Begin VB.Label lblDescClienteDest 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDescClienteDest"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7245
         TabIndex        =   12
         Top             =   960
         Width           =   8055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código do Envio"
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
         Left            =   4440
         TabIndex        =   11
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   18855
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
         Picture         =   "frmCADRECMATLIT.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "frmCADRECMATLIT.frx":0532
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
         Picture         =   "frmCADRECMATLIT.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdImpressao 
         Caption         =   "Im&prime"
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
         Left            =   18000
         Picture         =   "frmCADRECMATLIT.frx":0736
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Imprime Registro"
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmCADRECMATLIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public cTipOper         As String
Public iCodigo          As Long
Public strMODPAI        As String
Public FILIAL           As Integer
Public strAcesso        As String
Public strUsuario       As String
Public lngCodUsuario    As Long

Dim objBLBFunc          As Object
Dim objCADRECMATLIT     As Object
Dim objPESQPADRAO       As Object
Dim objRel              As Object
Dim strCAPTION          As String
Dim arrENTROTLIT        As Variant


Const conCOL_ENTROTLIT_ID                      As Integer = 0
Const conCOL_ENTROTLIT_OP                      As Integer = 1
Const conCOL_ENTROTLIT_PEDIDO                  As Integer = 2
Const conCOL_ENTROTLIT_IDPRODUTO               As Integer = 3
Const conCOL_ENTROTLIT_CODIGO                  As Integer = 4
Const conCOL_ENTROTLIT_CODCAPAC                As Integer = 5
Const conCOL_ENTROTLIT_CAPAC                   As Integer = 6
Const conCOL_ENTROTLIT_LITSTEEL                As Integer = 7
Const conCOL_ENTROTLIT_PADRAO                  As Integer = 8
Const conCOL_ENTROTLIT_CODFOLHAUSADA           As Integer = 9
Const conCOL_ENTROTLIT_PESQFOLHAUSADA          As Integer = 10
Const conCOL_ENTROTLIT_FOLHAUSADA              As Integer = 11
Const conCOL_ENTROTLIT_ESPESS                  As Integer = 12
Const conCOL_ENTROTLIT_LARG                    As Integer = 13
Const conCOL_ENTROTLIT_COMP                    As Integer = 14
Const conCOL_ENTROTLIT_QTDECORP                As Integer = 15
Const conCOL_ENTROTLIT_PERDPRODC               As Integer = 16
Const conCOL_ENTROTLIT_QTDEFOLHAS              As Integer = 17
Const conCOL_ENTROTLIT_PESO                    As Integer = 18
Const conCOL_ENTROTLIT_UNID                    As Integer = 19
Const conCOL_ENTROTLIT_QTDELATAS               As Integer = 20
Const conCOL_ENTROTLIT_QTDEFARDOS              As Integer = 21
Const conCOL_ENTROTLIT_CONFPESO                As Integer = 22
Const conCOL_ENTROTLIT_CONFFARDO               As Integer = 23
Const conCOL_ENTROTLIT_STATUS                  As Integer = 24
Const conCOL_ENTROTLIT_QTDEFOLHASREC           As Integer = 25
Const conCOL_ENTROTLIT_QTDELATASREC            As Integer = 26
Const conCOL_ENTROTLIT_FormatString            As String = "=ID|Cód.OP|Pedido|IDPRODUTO|Código|Cod.Capac|Capac.|Descrição Litografia|Padrão|Cod.Folha|...|Folha.Usada|Espessura|Largura|Comprimento|Qtde.Corpos|Perd.Proc|Qtde.Folhas|Peso|Unid.|Qtde.Latas|Qtde.Fardos|Conf. Peso|Conf. Fardo|STATUS|Qtde.Folhas.Rec|Qtde.Latas.Rec"
Const conColumnsIn_ENTROTLIT                   As Integer = 27

Private Sub cmdAltera_Click()

    cTipOper = "A"
    If objBLBFunc.ChecaAcesso2(cTipOper, strAcesso) = False Then Exit Sub
    
    Call objBLBFunc.MudaBotoes(CmdSalva, cmdAltera, cTipOper)
    Call objBLBFunc.MudaCaption(Me, strCAPTION, cTipOper)
    Call DesabilitaCampos(Trim(cTipOper))

End Sub

Private Sub CmdSalva_Click()

    Dim i                   As Long
    Dim sValor              As String
    
    If ValidaCampos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADRECMATLIT.CODIGO = objBLBFunc.Gera_Codigo(Trim(Me.Name), FILIAL, Linha)
    
    objCADRECMATLIT.STATUS = "'ENV'"
    If Conf_Status = True Then objCADRECMATLIT.STATUS = "'REC'"
    
    objCADRECMATLIT.DTENTRADA = "'" & Format(CDate(mskDTENTRADA.Text), "MM/DD/YYYY") & "'"
    objCADRECMATLIT.CODCLIE = Trim(txtCIDCLIE.Text)
    objCADRECMATLIT.CODCLIEDEST = Trim(txtCLIEDEST.Text)
    objCADRECMATLIT.CODIGOENV = CLng(txtCODENV.Text)
    
    arrENTROTLIT = Empty
    With grdLANCTOS
        ReDim arrENTROTLIT(1 To (.Rows - 1), 1 To 23) As String
        For i = 1 To (.Rows - 1)
            
            arrENTROTLIT(i, 1) = objBLBFunc.Gera_Codigo(Trim(Me.Name & "_CODITEN"), FILIAL, Linha)
            arrENTROTLIT(i, 2) = Trim(Str(.Cell(flexcpText, i, conCOL_ENTROTLIT_OP)))
            arrENTROTLIT(i, 3) = Trim(Str(.Cell(flexcpText, i, conCOL_ENTROTLIT_PEDIDO)))
            arrENTROTLIT(i, 4) = Trim(Str(.Cell(flexcpText, i, conCOL_ENTROTLIT_IDPRODUTO)))
            arrENTROTLIT(i, 5) = "'" & Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_CODIGO)) & "'"
            arrENTROTLIT(i, 6) = Trim(Str(.Cell(flexcpText, i, conCOL_ENTROTLIT_CODCAPAC)))
            arrENTROTLIT(i, 7) = Trim(Str(.Cell(flexcpText, i, conCOL_ENTROTLIT_PADRAO)))
            
            sValor = "Null"
            If Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_ESPESS))) > 0 Then
               sValor = Replace(.Cell(flexcpText, i, conCOL_ENTROTLIT_ESPESS), ".", "")
               sValor = Replace(sValor, ",", ".")
            End If
            arrENTROTLIT(i, 8) = sValor
            
            sValor = "Null"
            If Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_LARG))) > 0 Then
               sValor = Replace(.Cell(flexcpText, i, conCOL_ENTROTLIT_LARG), ".", "")
               sValor = Replace(sValor, ",", ".")
            End If
            arrENTROTLIT(i, 9) = sValor
            
            sValor = "Null"
            If Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_COMP))) > 0 Then
               sValor = Replace(.Cell(flexcpText, i, conCOL_ENTROTLIT_COMP), ".", "")
               sValor = Replace(sValor, ",", ".")
            End If
            arrENTROTLIT(i, 10) = sValor
            
            sValor = "Null"
            If Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_QTDECORP))) > 0 Then
               sValor = Replace(.Cell(flexcpText, i, conCOL_ENTROTLIT_QTDECORP), ".", "")
               sValor = Replace(sValor, ",", ".")
            End If
            arrENTROTLIT(i, 11) = sValor
            
            sValor = "Null"
            If Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_PERDPRODC))) > 0 Then
               sValor = Replace(.Cell(flexcpText, i, conCOL_ENTROTLIT_PERDPRODC), ".", "")
               sValor = Replace(sValor, ",", ".")
            End If
            arrENTROTLIT(i, 12) = sValor
            
            sValor = "Null"
            If Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_QTDEFOLHAS))) > 0 Then
               sValor = Replace(.Cell(flexcpText, i, conCOL_ENTROTLIT_QTDEFOLHAS), ".", "")
               sValor = Replace(sValor, ",", ".")
            End If
            arrENTROTLIT(i, 13) = sValor
            
            sValor = "Null"
            If Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_PESO))) > 0 Then
               sValor = Replace(.Cell(flexcpText, i, conCOL_ENTROTLIT_PESO), ".", "")
               sValor = Replace(sValor, ",", ".")
            End If
            arrENTROTLIT(i, 14) = sValor
            
            arrENTROTLIT(i, 15) = "Null"
            If Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_UNID))) > 0 Then arrENTROTLIT(i, 15) = "'" & Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_UNID)) & "'"
            
            sValor = "Null"
            If Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_QTDELATAS))) > 0 Then
               sValor = Replace(.Cell(flexcpText, i, conCOL_ENTROTLIT_QTDELATAS), ".", "")
               sValor = Replace(sValor, ",", ".")
            End If
            arrENTROTLIT(i, 16) = sValor
            
            sValor = "Null"
            If Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_QTDEFARDOS))) > 0 Then
               sValor = Replace(.Cell(flexcpText, i, conCOL_ENTROTLIT_QTDEFARDOS), ".", "")
               sValor = Replace(sValor, ",", ".")
            End If
            arrENTROTLIT(i, 17) = sValor
            
            arrENTROTLIT(i, 18) = .Cell(flexcpText, i, conCOL_ENTROTLIT_CODFOLHAUSADA)
            
            sValor = "Null"
            If Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_CONFPESO))) > 0 Then
               sValor = Replace(.Cell(flexcpText, i, conCOL_ENTROTLIT_CONFPESO), ".", "")
               sValor = Replace(sValor, ",", ".")
            End If
            arrENTROTLIT(i, 19) = sValor
        
            sValor = "Null"
            If Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_CONFFARDO))) > 0 Then
               sValor = Replace(.Cell(flexcpText, i, conCOL_ENTROTLIT_CONFFARDO), ".", "")
               sValor = Replace(sValor, ",", ".")
            End If
            arrENTROTLIT(i, 20) = sValor
            
            arrENTROTLIT(i, 21) = "Null"
            If Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_STATUS))) > 0 Then arrENTROTLIT(i, 21) = "'" & Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_STATUS)) & "'"
        
            sValor = "Null"
            If Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_QTDEFOLHASREC))) > 0 Then
               sValor = Replace(.Cell(flexcpText, i, conCOL_ENTROTLIT_QTDEFOLHASREC), ".", "")
               sValor = Replace(sValor, ",", ".")
            End If
            arrENTROTLIT(i, 22) = sValor
        
            sValor = "Null"
            If Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_QTDELATASREC))) > 0 Then
               sValor = Replace(.Cell(flexcpText, i, conCOL_ENTROTLIT_QTDELATASREC), ".", "")
               sValor = Replace(sValor, ",", ".")
            End If
            arrENTROTLIT(i, 23) = sValor
        
        Next i
    End With
    objCADRECMATLIT.LANCTOS = arrENTROTLIT
    
    If objCADRECMATLIT.GRAVA(cTipOper) = False Then Exit Sub
          
    MsgBox "A Movimentação de Litografia foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
    If cTipOper = "I" Then Unload Me

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

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objCADRECMATLIT = Nothing
    Set objPESQPADRAO = Nothing
    Set objRel = Nothing
End Sub

Private Sub Form_Load()

    Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
    Set objCADRECMATLIT = CreateObject("CADRECMATLIT.clsCADRECMATLIT")
    Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
    Set objRel = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    If adoBanco_Dados.State = 0 Then Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)
    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
   
    strCAPTION = "Recebimento de Material (Folhas Litografadas)"
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
   
    objCADRECMATLIT.FILIAL = FILIAL
   
    Call IniciaForm

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

Private Sub IniciaForm()
    
    Call objBLBFunc.MudaBotoes(CmdSalva, cmdAltera, cTipOper)
    Call objBLBFunc.MudaCaption(Me, strCAPTION, cTipOper)
    Call objBLBFunc.LimpaCampos(Me)
    
    Call DesabilitaCampos(Trim(cTipOper))
    
    Call ConfGrd
    Call LimpaCamposLabel
    
    If cTipOper = "I" Then iCodigo = 0
    objCADRECMATLIT.CODIGO = iCodigo
    
    mskDTENTRADA.Text = Format(Now, "DD/MM/YYYY")
    
    Call CarregaCampos
    
End Sub


Private Sub LimpaCamposLabel()
    lblDescCliente.Caption = ""
    lblDescClienteDest.Caption = ""
End Sub

Private Sub DesabilitaCampos(strTipOper As String)
    If strTipOper = "I" Or strTipOper = "A" Then
        Frame2.Enabled = True
        cmdImpressao.Enabled = True
        txtCODENV.Enabled = True
        If strTipOper = "A" Then txtCODENV.Enabled = False
        If strTipOper = "I" Then cmdImpressao.Enabled = False
    ElseIf strTipOper = "C" Then
        Frame2.Enabled = False
        cmdImpressao.Enabled = True
    End If
End Sub

Private Sub grdLANCTOS_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    With grdLANCTOS
        If (.Rows - 1) = 0 Then Exit Sub
        If Row = 0 Then Exit Sub
        Select Case Col
               Case conCOL_ENTROTLIT_CONFPESO
                    If Len(Trim(.Cell(flexcpText, Row, Col))) > 0 Then .Cell(flexcpText, Row, Col) = Format(.Cell(flexcpText, Row, Col), "#,####0.0000")
               Case conCOL_ENTROTLIT_CONFFARDO
                    If Row < (.Rows - 1) Then
                        .Row = (Row + 1)
                    End If
               Case Else
                    .ComboList = ""
        End Select
        Call Conf_Recebimento
    End With

End Sub

Private Sub grdLANCTOS_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    With grdLANCTOS
        Select Case Col
               Case conCOL_ENTROTLIT_ID, _
                    conCOL_ENTROTLIT_PEDIDO, _
                    conCOL_ENTROTLIT_IDPRODUTO, _
                    conCOL_ENTROTLIT_CODIGO, _
                    conCOL_ENTROTLIT_CODCAPAC, _
                    conCOL_ENTROTLIT_CAPAC, _
                    conCOL_ENTROTLIT_LITSTEEL, _
                    conCOL_ENTROTLIT_QTDECORP, _
                    conCOL_ENTROTLIT_PERDPRODC, _
                    conCOL_ENTROTLIT_QTDEFOLHAS, _
                    conCOL_ENTROTLIT_QTDELATAS, _
                    conCOL_ENTROTLIT_FOLHAUSADA, _
                    conCOL_ENTROTLIT_OP, _
                    conCOL_ENTROTLIT_PESO, _
                    conCOL_ENTROTLIT_UNID, _
                    conCOL_ENTROTLIT_QTDEFARDOS, _
                    conCOL_ENTROTLIT_CODFOLHAUSADA, _
                    conCOL_ENTROTLIT_PESQFOLHAUSADA, _
                    conCOL_ENTROTLIT_ESPESS, _
                    conCOL_ENTROTLIT_LARG, _
                    conCOL_ENTROTLIT_COMP, _
                    conCOL_ENTROTLIT_STATUS, _
                    conCOL_ENTROTLIT_QTDEFOLHASREC, _
                    conCOL_ENTROTLIT_QTDELATASREC
                    Cancel = True
               Case conCOL_ENTROTLIT_CONFPESO, _
                    conCOL_ENTROTLIT_CONFFARDO
                    If cTipOper = "C" Then Cancel = True
               Case Else
                   .ComboList = ""
               End Select
    End With
    
    Exit Sub

End Sub

Private Sub grdLANCTOS_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdLANCTOS
          Select Case Col
                    Case conCOL_ENTROTLIT_CONFFARDO
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
                    Case conCOL_ENTROTLIT_CONFPESO
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 4, myvarAsDouble)
          End Select
     End With

End Sub

Private Sub grdLANCTOS_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)


     Dim curSALDOPESO   As Currency
     Dim lngSALDOFARDO  As Long

     With grdLANCTOS
          Select Case Col
                        
                 Case conCOL_ENTROTLIT_CONFPESO
                        If .EditText = Empty Then Exit Sub
                        
                        If Not IsNumeric(.EditText) Then
                            MsgBox "ATENÇÃO - O Peso está Inválido !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                        
                        curSALDOPESO = (CCur(.Cell(flexcpText, Row, conCOL_ENTROTLIT_PESO)) - CCur(.EditText))
                        If curSALDOPESO < 0 Then
                            MsgBox "ATENÇÂO" & vbCrLf & _
                                   "Vc esta tentando receber mais peso do que foi enviado !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
           
                        .Cell(flexcpText, Row, conCOL_ENTROTLIT_QTDEFOLHASREC) = QtdeFolhas(Row, CCur(.EditText))
                        .Cell(flexcpText, Row, conCOL_ENTROTLIT_QTDELATASREC) = QtdeLatas(Row)
                 
                        
                        Call PosColCapac(conCOL_ENTROTLIT_CONFFARDO, Row)
                 
                 Case conCOL_ENTROTLIT_CONFFARDO
                        If .EditText = Empty Then
                            Exit Sub
                        End If
                        
                        If Not IsNumeric(.EditText) Then
                            MsgBox "ATENÇÃO - Qtde. Fardos inválido !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                        
                        lngSALDOFARDO = CLng(.Cell(flexcpText, Row, conCOL_ENTROTLIT_QTDEFARDOS)) - CLng(.EditText)
                        If lngSALDOFARDO < 0 Then
                            MsgBox "ATENÇÂO" & vbCrLf & _
                                   "Vc esta tentando receber mais fardos do que foi enviado !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                        
                        If Row < (.Rows - 1) Then
                            Call PosColCapac(conCOL_ENTROTLIT_CONFPESO, Row)
                        End If
          
          End Select
     End With

End Sub

Private Sub mskDTENTRADA_GotFocus()
    objBLBFunc.SelecionaCampos mskDTENTRADA.Name, Me
End Sub


Private Sub txtCODENV_GotFocus()
    objBLBFunc.SelecionaCampos txtCODENV.Name, Me
End Sub

Private Sub txtCODENV_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtCODENV.Text
End Sub

Private Sub txtCODENV_Validate(Cancel As Boolean)

    If Len(Trim(txtCODENV.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODENV.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODENV.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    Call CarregaCampos_Envio
    
End Sub

Private Sub ConfGrd()

    With grdLANCTOS

       .Cols = conColumnsIn_ENTROTLIT
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_ENTROTLIT_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeNone
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_ID) = ""
       .ColDataType(conCOL_ENTROTLIT_ID) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_OP) = ""
       .ColDataType(conCOL_ENTROTLIT_OP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_PEDIDO) = ""
       .ColDataType(conCOL_ENTROTLIT_PEDIDO) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_IDPRODUTO) = ""
       .ColDataType(conCOL_ENTROTLIT_IDPRODUTO) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_CODIGO) = ""
       .ColDataType(conCOL_ENTROTLIT_CODIGO) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_CODCAPAC) = ""
       .ColDataType(conCOL_ENTROTLIT_CODCAPAC) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_CAPAC) = ""
       .ColDataType(conCOL_ENTROTLIT_CAPAC) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_LITSTEEL) = ""
       .ColDataType(conCOL_ENTROTLIT_LITSTEEL) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_PADRAO) = ""
       .ColDataType(conCOL_ENTROTLIT_PADRAO) = flexDTString
       .ColComboList(conCOL_ENTROTLIT_PADRAO) = "|#1;Sim|#0;Não"
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_ESPESS) = ""
       .ColDataType(conCOL_ENTROTLIT_ESPESS) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_LARG) = ""
       .ColDataType(conCOL_ENTROTLIT_LARG) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_COMP) = ""
       .ColDataType(conCOL_ENTROTLIT_COMP) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_QTDECORP) = ""
       .ColDataType(conCOL_ENTROTLIT_QTDECORP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_PERDPRODC) = ""
       .ColDataType(conCOL_ENTROTLIT_PERDPRODC) = flexDTDouble
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_QTDEFOLHAS) = ""
       .ColDataType(conCOL_ENTROTLIT_QTDEFOLHAS) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_PESO) = ""
       .ColDataType(conCOL_ENTROTLIT_PESO) = flexDTDouble
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_QTDELATAS) = ""
       .ColDataType(conCOL_ENTROTLIT_QTDELATAS) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_QTDEFARDOS) = ""
       .ColDataType(conCOL_ENTROTLIT_QTDEFARDOS) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_CODFOLHAUSADA) = ""
       .ColDataType(conCOL_ENTROTLIT_CODFOLHAUSADA) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_FOLHAUSADA) = ""
       .ColDataType(conCOL_ENTROTLIT_FOLHAUSADA) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_PESQFOLHAUSADA) = ""
       .ColDataType(conCOL_ENTROTLIT_PESQFOLHAUSADA) = flexDTString
       .ColComboList(conCOL_ENTROTLIT_PESQFOLHAUSADA) = "..."
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_CONFPESO) = ""
       .ColDataType(conCOL_ENTROTLIT_CONFPESO) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_CONFFARDO) = ""
       .ColDataType(conCOL_ENTROTLIT_CONFFARDO) = flexDTCurrency
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_STATUS) = ""
       .ColDataType(conCOL_ENTROTLIT_STATUS) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_QTDEFOLHASREC) = ""
       .ColDataType(conCOL_ENTROTLIT_QTDEFOLHASREC) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_ENTROTLIT_QTDELATASREC) = ""
       .ColDataType(conCOL_ENTROTLIT_QTDELATASREC) = flexDTLong
       
       .ColWidth(conCOL_ENTROTLIT_ID) = 0
       .ColWidth(conCOL_ENTROTLIT_OP) = 1100
       .ColWidth(conCOL_ENTROTLIT_PEDIDO) = 1100
       .ColWidth(conCOL_ENTROTLIT_IDPRODUTO) = 0
       .ColWidth(conCOL_ENTROTLIT_CODIGO) = 1500
       .ColWidth(conCOL_ENTROTLIT_CODCAPAC) = 0
       .ColWidth(conCOL_ENTROTLIT_CAPAC) = 0
       .ColWidth(conCOL_ENTROTLIT_LITSTEEL) = 5000
       .ColWidth(conCOL_ENTROTLIT_ESPESS) = 1000
       .ColWidth(conCOL_ENTROTLIT_LARG) = 1000
       .ColWidth(conCOL_ENTROTLIT_COMP) = 1000
       .ColWidth(conCOL_ENTROTLIT_QTDECORP) = 1000
       .ColWidth(conCOL_ENTROTLIT_PERDPRODC) = 0
       .ColWidth(conCOL_ENTROTLIT_QTDEFOLHAS) = 1000
       .ColWidth(conCOL_ENTROTLIT_PESO) = 1000
       .ColWidth(conCOL_ENTROTLIT_UNID) = 0
       .ColWidth(conCOL_ENTROTLIT_QTDELATAS) = 1000
       .ColWidth(conCOL_ENTROTLIT_QTDEFARDOS) = 1000
       .ColWidth(conCOL_ENTROTLIT_PADRAO) = 0
       
       .ColWidth(conCOL_ENTROTLIT_CODFOLHAUSADA) = 1000
       .ColWidth(conCOL_ENTROTLIT_PESQFOLHAUSADA) = 300
       .ColWidth(conCOL_ENTROTLIT_FOLHAUSADA) = 1500
       
       .ColWidth(conCOL_ENTROTLIT_CONFPESO) = 1500
       .ColWidth(conCOL_ENTROTLIT_CONFFARDO) = 1500
       .ColWidth(conCOL_ENTROTLIT_STATUS) = 0
       
       .ColWidth(conCOL_ENTROTLIT_QTDEFOLHASREC) = 1300
       .ColWidth(conCOL_ENTROTLIT_QTDELATASREC) = 1300
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       
       
    End With
    
End Sub


Private Sub CarregaCampos_Envio()

On Error GoTo Err_CarregaCampos_Envio
    
    objCADRECMATLIT.CODIGOENV = CLng(txtCODENV.Text)
    
    If objCADRECMATLIT.Carrega_campos_Envio = True Then
        txtCIDCLIE.Text = objCADRECMATLIT.CODCLIE
        txtCLIEDEST.Text = objCADRECMATLIT.CODCLIEDEST
        Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", txtCIDCLIE.Text, lblDescCliente, "CLIE")
        Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", txtCLIEDEST, lblDescClienteDest, "CLIE")
    
        Call PopGrdLancto_Env
    Else
        MsgBox "ATENÇÂO" & vbCrLf & "Este Código de Envio não existe !!!", vbOKOnly + vbExclamation, "Aviso"
        txtCODENV.Text = ""
        objCADRECMATLIT.CODIGOENV = 0
        txtCODENV.SetFocus
    End If
    
    Exit Sub

Err_CarregaCampos_Envio:
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : CarregaCampos_Envio", Me.Name, "CarregaCampos_Envio")
    
End Sub


Private Sub PegaDescTabelas(StrCampoPesq As String, StrCampoRetorno As String, strTabela As String, strCodigo As String, lblLabel As Label, strTIPO As String)

    lblLabel.Caption = ""
    
    If Len(Trim(strCodigo)) = 0 Then Exit Sub
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoRetorno) & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & Trim(strTabela) & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       " & Trim(StrCampoPesq) & " = " & Trim(strCodigo)
    
    If strTIPO = "CLIE" Then sSql = sSql & "   And SGI_VISTELENTEST = 1 " & vbCrLf
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then
       lblLabel.Caption = Trim(BREC10(Trim(StrCampoRetorno)))
    Else
       MsgBox "Registro Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
    End If
    BREC10.Close
    
End Sub


Private Sub PopGrdLancto_Env()

    Dim i As Integer
    
    arrENTROTLIT = objCADRECMATLIT.LANCTOS
    If IsArray(arrENTROTLIT) Then
        With grdLANCTOS
            For i = 1 To UBound(arrENTROTLIT)
                .AddItem arrENTROTLIT(i, 1) & vbTab & _
                         arrENTROTLIT(i, 2) & vbTab & _
                         arrENTROTLIT(i, 3) & vbTab & _
                         arrENTROTLIT(i, 4) & vbTab & _
                         arrENTROTLIT(i, 5) & vbTab & _
                         arrENTROTLIT(i, 6) & vbTab & _
                         PegaDescCapac(Str(arrENTROTLIT(i, 6))) & vbTab & _
                         PegaDescProd(Str(arrENTROTLIT(i, 4))) & vbTab & _
                         arrENTROTLIT(i, 7) & vbTab & _
                         arrENTROTLIT(i, 18) & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         arrENTROTLIT(i, 8) & vbTab & _
                         arrENTROTLIT(i, 9) & vbTab & _
                         arrENTROTLIT(i, 10) & vbTab & _
                         arrENTROTLIT(i, 11) & vbTab & _
                         arrENTROTLIT(i, 12) & vbTab & _
                         arrENTROTLIT(i, 13) & vbTab & _
                         arrENTROTLIT(i, 14) & vbTab & _
                         arrENTROTLIT(i, 15) & vbTab & _
                         arrENTROTLIT(i, 16) & vbTab & _
                         arrENTROTLIT(i, 17) & vbTab & _
                         "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & ""
                         
                         
                Call PegaDadosFF(Str(arrENTROTLIT(i, 6)), Str(arrENTROTLIT(i, 18)), (.Rows - 1))
                Call PintaCelula(i)
            
            Next i
        End With
    End If

End Sub


Private Function PegaDescCapac(strCODCAPAC As String) As String

    PegaDescCapac = ""
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       SGI_DESCRI" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODLIN = " & strCODCAPAC
    
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC2.EOF() Then PegaDescCapac = BREC2!SGI_DESCRI
    BREC2.Close

End Function

Private Function PegaDescProd(strIDPROD As String) As String

    PegaDescProd = ""
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       SGI_DESCRICAO" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_IDPRODUTO = " & strIDPROD
    
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC2.EOF() Then PegaDescProd = BREC2!SGI_DESCRICAO
    BREC2.Close

End Function


Private Function PegaDadosFF(strCODLINHA As String, strCODFF As String, lngLINHA As Long) As Boolean

    PegaDadosFF = True
    
    If Len(Trim(strCODLINHA)) = 0 Then Exit Function
    If Len(Trim(strCODFF)) = 0 Then Exit Function
    If lngLINHA <= 0 Then Exit Function
    
    With grdLANCTOS
    
        sSql = ""
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "        MEDC.*" & vbCrLf
        sSql = sSql & "      , DIMC.SGI_DESCORTE" & vbCrLf
        
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "        SGI_CADLINHAPRODUTO LINH" & vbCrLf
        sSql = sSql & "      , SGI_MEDCORTELINHA   MEDC" & vbCrLf
        sSql = sSql & "      , SGI_CADDIMCORTE     DIMC" & vbCrLf
         
        sSql = sSql & "  Where" & vbCrLf
        sSql = sSql & "        LINH.SGI_FILIAL     = " & FILIAL & vbCrLf
        sSql = sSql & "    And LINH.SGI_CODLIN     = " & strCODLINHA & vbCrLf
        
        sSql = sSql & "    And MEDC.SGI_FILIAL     = LINH.SGI_FILIAL" & vbCrLf
        sSql = sSql & "    And MEDC.SGI_CODIGO     = LINH.SGI_CODIGO" & vbCrLf
        sSql = sSql & "    And MEDC.SGI_CODMEDCORT = " & strCODFF & vbCrLf
        sSql = sSql & "    And DIMC.SGI_FILIAL     = MEDC.SGI_FILIAL" & vbCrLf
        sSql = sSql & "    And DIMC.SGI_CODIGO     = MEDC.SGI_CODMEDCORT"
        
        
        BREC11.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC11.EOF() Then
        
            .Cell(flexcpText, lngLINHA, conCOL_ENTROTLIT_FOLHAUSADA) = Trim(BREC11!SGI_DESCORTE)
            .Cell(flexcpText, lngLINHA, conCOL_ENTROTLIT_ESPESS) = Format(BREC11!SGI_EXPESS, "#,####0.0000")
            .Cell(flexcpText, lngLINHA, conCOL_ENTROTLIT_LARG) = Format(BREC11!SGI_LARGUR, "#,####0.0000")
            .Cell(flexcpText, lngLINHA, conCOL_ENTROTLIT_COMP) = Format(BREC11!SGI_COMPRI, "#,####0.0000")
            .Cell(flexcpText, lngLINHA, conCOL_ENTROTLIT_QTDECORP) = BREC11!SGI_QTDECORPOS
            .Cell(flexcpText, lngLINHA, conCOL_ENTROTLIT_PERDPRODC) = Format(BREC11!SGI_PERDPROC, "#,####0.0000")
            PegaDadosFF = False
        
        Else
            MsgBox "ATENÇÂO" & vbCrLf & "Esta FF não existe para esta linha !!!", vbOKOnly + vbExclamation, "Aviso"
        End If
        BREC11.Close

    End With
End Function

Private Sub PintaCelula(intROW As Integer)
    With grdLANCTOS
        .Cell(flexcpBackColor, intROW, conCOL_ENTROTLIT_OP) = &H80C0FF
        .Cell(flexcpBackColor, intROW, conCOL_ENTROTLIT_PESO) = &H80C0FF
        .Cell(flexcpBackColor, intROW, conCOL_ENTROTLIT_QTDEFARDOS) = &H80C0FF
        
        .Cell(flexcpBackColor, intROW, conCOL_ENTROTLIT_CONFPESO) = &HFFFF80
        .Cell(flexcpBackColor, intROW, conCOL_ENTROTLIT_CONFFARDO) = &HFFFF80
    
    End With
End Sub


Private Sub Conf_Recebimento()
    Dim i As Long
    With grdLANCTOS
        For i = 1 To (.Rows - 1)
            .Cell(flexcpText, i, conCOL_ENTROTLIT_STATUS) = ""
            If Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_CONFPESO))) > 0 And _
               Len(Trim(.Cell(flexcpText, i, conCOL_ENTROTLIT_CONFFARDO))) > 0 Then
                .Cell(flexcpText, i, conCOL_ENTROTLIT_STATUS) = "REC"
            End If
        Next i
    End With
End Sub

Private Function ValidaCampos() As Boolean

        ValidaCampos = False
     
        Dim i As Integer
                
        If Len(Trim(Replace(mskDTENTRADA.Text, "/", ""))) = 0 Then
            MsgBox "ATENÇÃO" & vbCrLf & _
                   "Campo Data de Entrada não pode ser vázio !!!", vbOKOnly + vbExclamation, "Acviso"
                   mskDTENTRADA.SetFocus
                   Exit Function
        End If
     
        If Not IsDate(mskDTENTRADA.Text) Then
            MsgBox "ATENÇÃO" & vbCrLf & _
                   "Campo Data inválido !!!", vbOKOnly + vbExclamation, "Acviso"
                   mskDTENTRADA.SetFocus
                   Exit Function
        End If
        
     
        ValidaCampos = True
     
End Function

Private Function Conf_Status() As Boolean
    
    Conf_Status = False
    
    Dim i As Long
    Dim lngQTDLCTOS As Long
    Dim lngQTDREC   As Long
    
    With grdLANCTOS
        lngQTDLCTOS = (.Rows - 1)
        lngQTDREC = 0
        For i = 1 To (.Rows - 1)
            If .Cell(flexcpText, i, conCOL_ENTROTLIT_STATUS) = "REC" Then lngQTDREC = (lngQTDREC + 1)
        Next i
        If lngQTDREC = lngQTDLCTOS Then Conf_Status = True
    End With
End Function

Private Sub CarregaCampos()

On Error GoTo Err_CarregaCampos
    
    If objCADRECMATLIT.Carrega_campos = True Then
        txtCodigo.Text = objCADRECMATLIT.CODIGO
        mskDTENTRADA.Text = objCADRECMATLIT.DTENTRADA
        txtCIDCLIE.Text = objCADRECMATLIT.CODCLIE
        txtCLIEDEST.Text = objCADRECMATLIT.CODCLIEDEST
        txtCODENV.Text = objCADRECMATLIT.CODIGOENV
        
        Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", txtCIDCLIE.Text, lblDescCliente, "CLIE")
        Call PegaDescTabelas("SGI_CODIGO", "SGI_RAZAOSOC", "SGI_CADCLIENTE", txtCLIEDEST, lblDescClienteDest, "CLIE")
    
        Call PopGrdLancto
    
    End If
    
    Exit Sub

Err_CarregaCampos:
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : CarregaCampos", Me.Name, "CarregaCampos")
    
End Sub


Private Sub PopGrdLancto()

    Dim i As Integer
    
    arrENTROTLIT = objCADRECMATLIT.LANCTOS
    If IsArray(arrENTROTLIT) Then
        With grdLANCTOS
            For i = 1 To UBound(arrENTROTLIT)
                .AddItem arrENTROTLIT(i, 1) & vbTab & _
                         arrENTROTLIT(i, 2) & vbTab & _
                         arrENTROTLIT(i, 3) & vbTab & _
                         arrENTROTLIT(i, 4) & vbTab & _
                         arrENTROTLIT(i, 5) & vbTab & _
                         arrENTROTLIT(i, 6) & vbTab & _
                         PegaDescCapac(Str(arrENTROTLIT(i, 6))) & vbTab & _
                         PegaDescProd(Str(arrENTROTLIT(i, 4))) & vbTab & _
                         arrENTROTLIT(i, 7) & vbTab & _
                         arrENTROTLIT(i, 18) & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         arrENTROTLIT(i, 8) & vbTab & _
                         arrENTROTLIT(i, 9) & vbTab & _
                         arrENTROTLIT(i, 10) & vbTab & _
                         arrENTROTLIT(i, 11) & vbTab & _
                         arrENTROTLIT(i, 12) & vbTab & _
                         arrENTROTLIT(i, 13) & vbTab & _
                         arrENTROTLIT(i, 14) & vbTab & _
                         arrENTROTLIT(i, 15) & vbTab & _
                         arrENTROTLIT(i, 16) & vbTab & _
                         arrENTROTLIT(i, 17) & vbTab & _
                         arrENTROTLIT(i, 19) & vbTab & arrENTROTLIT(i, 20) & vbTab & arrENTROTLIT(i, 21) & vbTab & arrENTROTLIT(i, 22) & vbTab & arrENTROTLIT(i, 23)
                
                         
                Call PegaDadosFF(Str(arrENTROTLIT(i, 6)), Str(arrENTROTLIT(i, 18)), (.Rows - 1))
                Call PintaCelula(i)
            
            Next i
        End With
    End If

End Sub


Private Function QtdeFolhas(lngROW As Long, currPeso As Currency) As Long

    QtdeFolhas = 0

    Dim curPESO        As Currency
    Dim curExpessura   As Currency
    Dim curLargura     As Currency
    Dim curComprimento As Currency

    With grdLANCTOS
        curPESO = currPeso
        curExpessura = 0
        curLargura = 0
        curComprimento = 0
    
        If Len(Trim(.Cell(flexcpText, lngROW, conCOL_ENTROTLIT_ESPESS))) > 0 Then curExpessura = CCur(.Cell(flexcpText, lngROW, conCOL_ENTROTLIT_ESPESS))
        If Len(Trim(.Cell(flexcpText, lngROW, conCOL_ENTROTLIT_LARG))) > 0 Then curLargura = CCur(.Cell(flexcpText, lngROW, conCOL_ENTROTLIT_LARG))
        If Len(Trim(.Cell(flexcpText, lngROW, conCOL_ENTROTLIT_COMP))) > 0 Then curComprimento = CCur(.Cell(flexcpText, lngROW, conCOL_ENTROTLIT_COMP))
    
    End With

    If curExpessura > 0 And curLargura > 0 And curComprimento > 0 Then
        QtdeFolhas = (curPESO / (curExpessura * (curLargura / 1000) * (curComprimento / 1000) * 7.85))
    End If

End Function

Private Function QtdeLatas(lngROW As Long) As Long

    QtdeLatas = 0
    
    Dim lngQTDEFOLHAS As Long
    Dim lngQTDECORPOS As Long
    
    lngQTDEFOLHAS = 0
    lngQTDECORPOS = 0
    
    With grdLANCTOS
        If Len(Trim(.Cell(flexcpText, lngROW, conCOL_ENTROTLIT_QTDEFOLHASREC))) > 0 Then lngQTDEFOLHAS = CLng(.Cell(flexcpText, lngROW, conCOL_ENTROTLIT_QTDEFOLHASREC))
        If Len(Trim(.Cell(flexcpText, lngROW, conCOL_ENTROTLIT_QTDECORP))) > 0 Then lngQTDECORPOS = CLng(.Cell(flexcpText, lngROW, conCOL_ENTROTLIT_QTDECORP))
    End With
    
    QtdeLatas = (lngQTDEFOLHAS * lngQTDECORPOS)
    
End Function

Private Sub PosColCapac(lngPOSCOL As Long, lngPOSROL As Long)
    
On Error GoTo Err_PosCol
    
    With grdLANCTOS
        .SetFocus
        .Row = lngPOSROL
        .Col = lngPOSCOL
    End With
    
    Exit Sub
    
Err_PosCol:
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PosCol()", Me.Name, "PosCol()", strCAMARQERRO)
    
End Sub

