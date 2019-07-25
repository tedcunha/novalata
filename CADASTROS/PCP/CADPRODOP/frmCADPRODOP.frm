VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCADPRODOP 
   Caption         =   "Controle de OP Programadas"
   ClientHeight    =   7950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   18405
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command9 
      Height          =   300
      Left            =   18120
      Picture         =   "frmCADPRODOP.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Inclui uma nova linha na Gride"
      Top             =   1680
      Width           =   300
   End
   Begin VB.CommandButton Command3 
      Height          =   300
      Left            =   18120
      Picture         =   "frmCADPRODOP.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2040
      Width           =   300
   End
   Begin VSFlex8LCtl.VSFlexGrid grdOPAPOT 
      Height          =   6255
      Left            =   4200
      TabIndex        =   10
      Top             =   1680
      Width           =   13815
      _cx             =   24368
      _cy             =   11033
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
      SelectionMode   =   1
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
   Begin VSFlex8LCtl.VSFlexGrid grdAPONT 
      Height          =   6255
      Left            =   0
      TabIndex        =   9
      Top             =   1680
      Width           =   4095
      _cx             =   7223
      _cy             =   11033
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
      SelectionMode   =   1
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
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   18375
      Begin VB.CommandButton cmdCARREGA 
         Caption         =   "&Carrega Gride - <F5>"
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
         Left            =   6720
         TabIndex        =   13
         Top             =   240
         Width           =   2055
      End
      Begin MSMask.MaskEdBox mskDtDoc 
         Height          =   285
         Left            =   5400
         TabIndex        =   8
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Text            =   "txtCodigo"
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Data do Apontamento"
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
         Left            =   3360
         TabIndex        =   7
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "N° Apontamento"
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
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18375
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
         Picture         =   "frmCADPRODOP.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "frmCADPRODOP.frx":07C6
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Picture         =   "frmCADPRODOP.frx":08C8
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADPRODOP"
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
Public strUSUARIO       As String
Public lngCodVendedor   As Long
Public lngCodUsuario    As Long
Public strEMPRESA       As String
Public strTABELA        As String

Dim lngCodLog           As Long
Dim strCAPTION          As String
Dim strNOMTABELA        As String
Dim strModulo           As String

Dim objBLBFunc          As New clsFuncoes
Dim objCADPRODOP        As New clsCADPRODOP
Dim objPESQPADRAO       As New clsPESQPADRAO
Dim arrPROGRAMADO       As Variant

'' -----------------------------------------------------------------------------------
Const conCOL_PRODOP_CODLINHA                    As Integer = 0
Const conCOL_PRODOP_DESCLINHA                   As Integer = 1
Const conCOL_PRODOP_QTDEAPONT                   As Integer = 2
Const conCOL_PRODOP_IDLINHA                     As Integer = 3
Const conCOL_PRODOP_IDINTERNO                   As Integer = 4
Const conCOL_PRODOP_Action2Do                   As Integer = 5
Const conCOL_PRODOP_FormatString                As String = "=Cod.Linha|Linha|Qntd.Apontada|ID.LINHA|IDINTERNO|Action2Do"
Const conColumnsIn_PRODOP                       As Integer = 6

'' -----------------------------------------------------------------------------------
Const conCOL_PRODOPAP_DTPROG                    As Integer = 0
Const conCOL_PRODOPAP_CODOP                     As Integer = 1
Const conCOL_PRODOPAP_CODROT                    As Integer = 2
Const conCOL_PRODOPAP_DESCROT                   As Integer = 3
Const conCOL_PRODOPAP_NECK                      As Integer = 4
Const conCOL_PRODOPAP_FECHAMENTO                As Integer = 5
Const conCOL_PRODOPAP_COMPONENTES               As Integer = 6
Const conCOL_PRODOPAP_QTDEPED                   As Integer = 7
Const conCOL_PRODOPAP_QTDEPROD                  As Integer = 8
Const conCOL_PRODOPAP_QTDFOLHAS                 As Integer = 9
Const conCOL_PRODOPAP_PESO                      As Integer = 10
Const conCOL_PRODOPAP_CODLINHA                  As Integer = 11
Const conCOL_PRODOPAP_IDLINHA                   As Integer = 12
Const conCOL_PRODOPAP_IDINTPROG                 As Integer = 13
Const conCOL_PRODOPAP_IDINTOP                   As Integer = 14
Const conCOL_PRODOPAP_IDPRODUTO                 As Integer = 15
Const conCOL_PRODOPAP_IDINTERNO                 As Integer = 16
Const conCOL_PRODOPAP_Action2Do                 As Integer = 17
Const conCOL_PRODOPAP_PERDAMONT                 As Integer = 18
Const conCOL_PRODOPAP_HORAINI                   As Integer = 19
Const conCOL_PRODOPAP_HORAFIN                   As Integer = 20
Const conCOL_PRODOPAP_StatusApont               As Integer = 21
Const conCOL_PRODOPAP_CODFOLHA                  As Integer = 22
Const conCOL_PRODOPAP_ESPESS                    As Integer = 23
Const conCOL_PRODOPAP_LARG                      As Integer = 24
Const conCOL_PRODOPAP_COMP                      As Integer = 25
Const conCOL_PRODOPAP_QTDECORP                  As Integer = 26
Const conCOL_PRODOPAP_PERDPROD                  As Integer = 27
Const conCOL_PRODOPAP_CODPED                    As Integer = 28
Const conCOL_PRODOPAP_Marca                     As Integer = 29
Const conCOL_PRODOPAP_QTDETOTAPONT              As Integer = 30
Const conCOL_PRODOPAP_STATUSPROGFIN             As Integer = 31
Const conCOL_PRODOPAP_TOTALHORAS                As Integer = 32
Const conCOL_PRODOPAP_FormatString              As String = "=Dt.Progr.|Cod.OP|Rótulo|Descrição do Rótulo|NECK|FECH|COMP|Qntd.Pedido|Qntd.Prod.|Qtde.Folhas|Peso|Cod.Linha|ID.LINHA|IDINTPROG|IDINTOP|IDPRODUTO|IDINTERNO|Action2Do|Perda|Hor.Ini|Hor.Fin|Status|CODFOLHA|EXPESS|LARG|COMP|QTDECORP|PERDPROC|CODPED|  |QTDEAPONTTOTAL|STATUSAPONTFEC|TOT.HORAS"
Const conColumnsIn_PRODOPAP                     As Integer = 33

Private Sub cmdAltera_Click()

    cTipOper = "A"
    
    Call objBLBFunc.MudaBotoes(CmdSalva, cmdAltera, cTipOper)
    Call objBLBFunc.MudaCaption(frmCADPRODOP, strCAPTION, cTipOper)
    
End Sub

Private Sub cmdCARREGA_Click()
    If ConsisteData = False Then Exit Sub
    
    If cTipOper = "I" Then
        If ConsisteLanc("'" & Format(CDate(mskDtDoc.Text), "MM/DD/YYYY") & "'") = True Then
            MsgBox "ATENÇÂO" & vbCrLf & _
                   "Este dia de Apontamento já está lançado entre com a opção de Alteração !!!", vbOKOnly + vbExclamation, "Aviso"
                   Exit Sub
        End If
        Call PopGrdApont
    End If
    
End Sub

Private Sub CmdSalva_Click()
    
On Error GoTo errGrava
    
    Dim I               As Integer
    Dim sValor          As String
    Dim lngROWLINATU    As Long
    
    
    lngROWLINATU = grdAPONT.Row
    
    Call objBLBFunc.RemoveLinhaVazia(grdOPAPOT, conCOL_PRODOPAP_CODOP)
    
    If Verif_Campos = False Then Exit Sub

    If cTipOper = "I" Then objCADPRODOP.Codigo = objBLBFunc.Gera_Codigo(Me.Name & strNOMTABELA, FILIAL, Linha)
    
    objCADPRODOP.DATLACTO = "'" & Format(CDate(mskDtDoc.Text), "MM/DD/YYYY") & "'"
    
    '' Verifica Para fechar o Apontamento
    ''Call FechaProg

    '' ===========================
    '' Gravando o Array
    arrPROGRAMADO = Empty
    With grdOPAPOT
        If (.Rows - 1) > 0 Then
            ReDim arrPROGRAMADO(1 To (.Rows - 1), 1 To 20) As String
            For I = 1 To (.Rows - 1)
                
                arrPROGRAMADO(I, 1) = .Cell(flexcpText, I, conCOL_PRODOPAP_CODOP)
                arrPROGRAMADO(I, 2) = .Cell(flexcpText, I, conCOL_PRODOPAP_QTDEPED)
                arrPROGRAMADO(I, 3) = .Cell(flexcpText, I, conCOL_PRODOPAP_QTDEPROD)
                arrPROGRAMADO(I, 4) = .Cell(flexcpText, I, conCOL_PRODOPAP_IDINTPROG)
                arrPROGRAMADO(I, 5) = .Cell(flexcpText, I, conCOL_PRODOPAP_IDINTOP)
                arrPROGRAMADO(I, 6) = .Cell(flexcpText, I, conCOL_PRODOPAP_IDLINHA)
                arrPROGRAMADO(I, 7) = .Cell(flexcpText, I, conCOL_PRODOPAP_CODLINHA)
                arrPROGRAMADO(I, 8) = .Cell(flexcpText, I, conCOL_PRODOPAP_IDPRODUTO)
                arrPROGRAMADO(I, 9) = .Cell(flexcpText, I, conCOL_PRODOPAP_Action2Do)
                
                If .Cell(flexcpText, I, conCOL_PRODOPAP_Action2Do) = dacEnumUpdateAction_Insert Then
                    arrPROGRAMADO(I, 10) = objBLBFunc.Gera_Codigo(Me.Name & strNOMTABELA & "_APONTP", FILIAL, Linha)
                Else
                    arrPROGRAMADO(I, 10) = .Cell(flexcpText, I, conCOL_PRODOPAP_IDINTERNO)
                End If
            
                arrPROGRAMADO(I, 11) = Trim(.Cell(flexcpText, I, conCOL_PRODOPAP_StatusApont))
                
                arrPROGRAMADO(I, 12) = "Null"
                If Len(Trim(Trim(.Cell(flexcpText, I, conCOL_PRODOPAP_PERDAMONT)))) > 0 Then
                    arrPROGRAMADO(I, 12) = Trim(.Cell(flexcpText, I, conCOL_PRODOPAP_PERDAMONT))
                End If
                
                arrPROGRAMADO(I, 13) = "Null"
                If Len(Trim(Trim(.Cell(flexcpText, I, conCOL_PRODOPAP_QTDFOLHAS)))) > 0 Then arrPROGRAMADO(I, 13) = Trim(.Cell(flexcpText, I, conCOL_PRODOPAP_QTDFOLHAS))
                    
                sValor = "Null"
                If Len(Trim(.Cell(flexcpText, I, conCOL_PRODOPAP_PESO))) > 0 Then
                   sValor = Replace(.Cell(flexcpText, I, conCOL_PRODOPAP_PESO), ".", "")
                   sValor = Replace(sValor, ",", ".")
                End If
                arrPROGRAMADO(I, 14) = sValor
                    
                arrPROGRAMADO(I, 15) = .Cell(flexcpText, I, conCOL_PRODOPAP_CODPED)
                arrPROGRAMADO(I, 16) = .Cell(flexcpText, I, conCOL_PRODOPAP_DTPROG)
                arrPROGRAMADO(I, 17) = .Cell(flexcpText, I, conCOL_PRODOPAP_QTDETOTAPONT)
                
                arrPROGRAMADO(I, 18) = "Null"
                arrPROGRAMADO(I, 19) = "Null"
                arrPROGRAMADO(I, 20) = "Null"
                
                If Len(Trim(.Cell(flexcpText, I, conCOL_PRODOPAP_HORAINI))) > 0 Then arrPROGRAMADO(I, 18) = "'" & Trim(.Cell(flexcpText, I, conCOL_PRODOPAP_HORAINI)) & "'"
                If Len(Trim(.Cell(flexcpText, I, conCOL_PRODOPAP_HORAFIN))) > 0 Then arrPROGRAMADO(I, 19) = "'" & Trim(.Cell(flexcpText, I, conCOL_PRODOPAP_HORAFIN)) & "'"
                If Len(Trim(.Cell(flexcpText, I, conCOL_PRODOPAP_TOTALHORAS))) > 0 Then arrPROGRAMADO(I, 20) = "'" & Trim(.Cell(flexcpText, I, conCOL_PRODOPAP_TOTALHORAS)) & "'"
            
            Next I
        End If
    End With
    objCADPRODOP.PROGRAMADO = arrPROGRAMADO
    '' ===========================

    '' Gravando as Informações no banco
    If objCADPRODOP.GRAVA(cTipOper, strTABELA) = False Then Exit Sub
    
    '' Atualizando os Dados
    ''If objBLBFunc.Atualiza(cTipOper, Str(objCADPRODOP.Codigo), FILIAL, Me.Name & strNOMTABELA, Linha) = False Then Exit Sub
    
    '' Gerando Log de Sistema
    ''lngCodLog = objBLBFunc.Gera_Codigo("SGI_LOGMODULO", FILIAL, Linha)
    ''Call objBLBFunc.GravaLogModulo(FILIAL, lngCodLog, Me.Name, cTipOper, lngCodUsuario, Str(objCADPRODOP.Codigo), Linha)
    
    MsgBox "As OP's ( " & Trim(Str(objCADPRODOP.Codigo)) & " ) foram " & IIf(cTipOper = "I", "inclusa(s)", IIf(cTipOper = "A", "alterada(s)", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
    
    If cTipOper = "I" Then cTipOper = "C"
    
    If objCADPRODOP.TemDados(strTABELA, Format(CDate(mskDtDoc.Text), "MM/DD/YYYY")) = False Then
        Unload Me
        Exit Sub
    Else
        iCodigo = objCADPRODOP.Codigo
        Call IniciaForm
    End If
    
    grdAPONT.Row = lngROWLINATU
    grdAPONT.RowSel = lngROWLINATU
    
    Exit Sub

errGrava:

    MsgBox "Erro Nº : " & Err.Number & vbCrLf & _
           "Descr.  : " & Err.Description, vbOKOnly + vbExclamation, "Aviso"



End Sub

Private Sub Command3_Click()
        
        Dim I       As Integer
        Dim intRESP As Integer
        
        If grdOPAPOT.Row <= 0 Then
            MsgBox "ATENÇÂO" & vbCrLf & _
                   "Selecione uma Linha !!!", vbOKOnly + vbExclamation, "viso"
            Exit Sub
        ElseIf (grdOPAPOT.Rows - 1) = 0 Then Exit Sub
            MsgBox "ATENÇÂO" & vbCrLf & _
                   "Não há dados de Linha !!!", vbOKOnly + vbExclamation, "viso"
            Exit Sub
        End If
        
        If cTipOper = "C" Then Exit Sub
        
        
        intRESP = MsgBox("ATENÇÃO" & vbCrLf & _
                         "Deseja Realmente excluir as OP's Relacionada(s) !!!", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
        
        
        If intRESP = vbNo Then Exit Sub
        
        With grdOPAPOT
VOLTA:
            For I = 1 To (.Rows - 1)
                If .Cell(flexcpText, I, conCOL_PRODOPAP_Action2Do) <> dacEnumUpdateAction_delete Then
                    If .Cell(flexcpChecked, I, conCOL_PRODOPAP_Marca) = 1 Then
                        If .Cell(flexcpText, I, conCOL_PRODOPAP_Action2Do) = dacEnumUpdateAction_Insert Then
                            Call objBLBFunc.ExclLinhaGrid(grdOPAPOT, I)
                            GoTo VOLTA
                        ElseIf .Cell(flexcpText, I, conCOL_PRODOPAP_Action2Do) = dacEnumUpdateAction_Ignore Or _
                               .Cell(flexcpText, I, conCOL_PRODOPAP_Action2Do) = dacEnumUpdateAction_update Then
                            .Cell(flexcpText, I, conCOL_PRODOPAP_Action2Do) = dacEnumUpdateAction_delete
                        End If
                    End If
                End If
            Next I
            Call MostraDados
            Call CalcTotApont
        End With

End Sub

Private Sub Command9_Click()
    If cTipOper = "I" Or cTipOper = "A" Then Call IncRegGrid
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
    If KeyCode = vbKeyF5 Then cmdCARREGA_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    objCADPRODOP.FILIAL = FILIAL
    objCADPRODOP.CODUSUARIO = lngCodUsuario
   
    If adoBanco_Dados.State = 0 Then Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)
    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    strNOMTABELA = strTABELA
    
    strCAPTION = Me.Caption & "/" & strEMPRESA & " - "
    
    Call IniciaForm
   
End Sub

Private Sub Destroy_Objeto()
    Set objBLBFunc = Nothing
    Set objCADPRODOP = Nothing
    Set objPESQPADRAO = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Destroy_Objeto
End Sub

Private Function Verif_Campos() As Boolean
    
    Dim I       As Integer
    Dim lngQTDE As Long
    
    Verif_Campos = False
    
    With grdOPAPOT
        For I = 1 To (.Rows - 1)
            If .Cell(flexcpText, I, conCOL_PRODOPAP_Action2Do) <> dacEnumUpdateAction_delete Then
               If Len(Trim(.Cell(flexcpText, I, conCOL_PRODOPAP_QTDEPROD))) = 0 Then
                    MsgBox "ATENÇÂO" & vbCrLf & "Existe(m) Qtde(s).Programada(s) Vázia(s) !!!", vbOKOnly + vbExclamation, "Aviso"
                    Exit Function
                End If
                If Len(Trim(.Cell(flexcpText, I, conCOL_PRODOPAP_StatusApont))) = 0 Then
                    MsgBox "ATENÇÂO" & vbCrLf & "Existe(m) OP's. sem o status de Apontamento !!!", vbOKOnly + vbExclamation, "Aviso"
                    Exit Function
                End If
                
                If Len(Trim(.Cell(flexcpText, I, conCOL_PRODOPAP_HORAINI))) > 0 And _
                   Len(Trim(.Cell(flexcpText, I, conCOL_PRODOPAP_HORAFIN))) = 0 Then
                   MsgBox "ATENÇÂO" & vbCrLf & _
                          "A OP. " & .Cell(flexcpText, I, conCOL_PRODOPAP_CODOP) & " está com a data final vázia !!!", vbOKOnly + vbExclamation, "Aviso"
                   Exit Function
                End If
                If Len(Trim(.Cell(flexcpText, I, conCOL_PRODOPAP_HORAINI))) = 0 And _
                   Len(Trim(.Cell(flexcpText, I, conCOL_PRODOPAP_HORAFIN))) > 0 Then
                   MsgBox "ATENÇÂO" & vbCrLf & _
                          "A OP. " & .Cell(flexcpText, I, conCOL_PRODOPAP_CODOP) & " está com a data inicial vázia !!!", vbOKOnly + vbExclamation, "Aviso"
                   Exit Function
                End If
                
                ''If Len(Trim(.Cell(flexcpText, I, conCOL_PRODOPAP_QTDFOLHAS))) = 0 Then
                ''    MsgBox "ATENÇÂO" & vbCrLf & "Existe(m) OP's. Qtde(s) de Folha(s) Vázia(s) !!!", vbOKOnly + vbExclamation, "Aviso"
                ''    Exit Function
                ''End If
                ''If Len(Trim(.Cell(flexcpText, I, conCOL_PRODOPAP_PERDAMONT))) = 0 Then
                ''    MsgBox "ATENÇÂO" & vbCrLf & "Existe(m) OP's. que a perda não foi informada !!!", vbOKOnly + vbExclamation, "Aviso"
                ''    Exit Function
                ''End If
            
            End If
        Next I
    End With
    
    
    Verif_Campos = True
End Function

Private Sub CarregaCampos()

    If objCADPRODOP.Codigo = 0 Then Exit Sub
    
    If objCADPRODOP.Carrega_Campos(strTABELA) = True Then
        txtCodigo.Text = objCADPRODOP.Codigo
        mskDtDoc.Text = objCADPRODOP.DATLACTO
        arrPROGRAMADO = objCADPRODOP.PROGRAMADO
    End If

    If cTipOper = "C" Or cTipOper = "A" Then
        Call PopGrdApont
        Call PopGrdOPApont
        
        With grdAPONT
            If (.Rows - 1) > 0 Then
                .Row = 1
                Call MostraDados
            End If
        End With
        
        Call CalcTotApont
    End If
    
End Sub

Private Sub IniciaForm()
    
    Call objBLBFunc.MudaBotoes(CmdSalva, cmdAltera, cTipOper)
    Call objBLBFunc.MudaCaption(Me, strCAPTION, cTipOper)
    Call objBLBFunc.LimpaCampos(Me)
    
    Call DesabilitaCampos(Trim(cTipOper))
    
    objCADPRODOP.Codigo = iCodigo
    If cTipOper = "I" Then mskDtDoc.Text = Format(Now, "DD/MM/YYYY")
    
    Call ConfGrd
    Call ConfGrdOPAPONT
    
    Call CarregaCampos
    
End Sub

Private Sub DesabilitaCampos(strTipOper As String)
    If strTipOper = "I" Then
        Frame2.Enabled = True
        mskDtDoc.Enabled = True
        cmdCARREGA.Enabled = True
    ElseIf strTipOper = "C" Or strTipOper = "A" Then
        Frame2.Enabled = True
        mskDtDoc.Enabled = False
        cmdCARREGA.Enabled = False
    End If
End Sub

Private Sub ConfGrd()

    With grdAPONT

       .Cols = conColumnsIn_PRODOP
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_PRODOP_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeNone
       
       .Cell(flexcpData, 0, conCOL_PRODOP_CODLINHA) = ""
       .ColDataType(conCOL_PRODOP_CODLINHA) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PRODOP_DESCLINHA) = ""
       .ColDataType(conCOL_PRODOP_DESCLINHA) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PRODOP_QTDEAPONT) = ""
       .ColDataType(conCOL_PRODOP_QTDEAPONT) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PRODOP_IDLINHA) = ""
       .ColDataType(conCOL_PRODOP_IDLINHA) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PRODOP_IDINTERNO) = ""
       .ColDataType(conCOL_PRODOP_IDINTERNO) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PRODOP_Action2Do) = ""
       .ColDataType(conCOL_PRODOP_Action2Do) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PRODOP_QTDEAPONT) = ""
       .ColDataType(conCOL_PRODOP_QTDEAPONT) = flexDTLong
       
       .ColWidth(conCOL_PRODOP_CODLINHA) = 0
       .ColWidth(conCOL_PRODOP_DESCLINHA) = 2500
       .ColWidth(conCOL_PRODOP_IDLINHA) = 0
       .ColWidth(conCOL_PRODOP_IDINTERNO) = 0
       .ColWidth(conCOL_PRODOP_Action2Do) = 0
       .ColWidth(conCOL_PRODOP_QTDEAPONT) = 1200
       
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       .FontName = "Arial"
       .FontSize = 7
       .FontBold = True

    End With
    
End Sub

Private Sub PopGrdApont()

    Call ConfGrd
    Call ConfGrdOPAPONT
    
    If Len(Trim(Replace(Replace(mskDtDoc.Text, "/", ""), "_", ""))) = 0 Then Exit Sub
    
    With grdAPONT
    
        sSql = ""
    
        sSql = "Select Distinct" & vbCrLf
        sSql = sSql & "                LINH.SGI_DESCRI     As SGI_DESCRLINHA" & vbCrLf
        sSql = sSql & "               ,LINH.SGI_CODLIN     As SGI_CODLINHA" & vbCrLf
        sSql = sSql & "               ,LINH.SGI_CODIGO     " & vbCrLf
        
        sSql = sSql & "           From" & vbCrLf
        sSql = sSql & "                SGI_CADMOVPCP" & strNOMTABELA & " MOVC" & vbCrLf
        sSql = sSql & "               ,SGI_CADPRODUTO      PROD" & vbCrLf
        sSql = sSql & "               ,SGI_CADLINHAPRODUTO LINH" & vbCrLf
        sSql = sSql & "          Where" & vbCrLf
        sSql = sSql & "                MOVC.SGI_FILIAL           = " & FILIAL & vbCrLf
        sSql = sSql & "           And  Month(MOVC.SGI_DATAPROG)  = " & Month(CDate(mskDtDoc.Text)) & vbCrLf
        sSql = sSql & "           And  Year(MOVC.SGI_DATAPROG)   = " & Year(CDate(mskDtDoc.Text)) & vbCrLf
        sSql = sSql & "           And  PROD.SGI_FILIAL           = MOVC.SGI_FILIAL" & vbCrLf
        sSql = sSql & "           And  PROD.SGI_IDPRODUTO        = MOVC.SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "           And  LINH.SGI_FILIAL           = PROD.SGI_FILIAL" & vbCrLf
        sSql = sSql & "           And  LINH.SGI_CODLIN           = PROD.SGI_CODLINPROD"
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC.EOF() Then
            Do While Not BREC.EOF()
                .AddItem BREC!SGI_CODLINHA & vbTab & _
                         BREC!SGI_DESCRLINHA & vbTab & _
                         0 & vbTab & _
                         BREC!SGI_CODIGO & vbTab & _
                         "" & vbTab & _
                         dacEnumUpdateAction_Ignore
    
               BREC.MoveNext
            Loop
        Else
            MsgBox "ATENÇÂO" & vbCrLf & "A Programação do Mês " & Format(Month(CDate(mskDtDoc.Text)), "##00") & "/" & Year(CDate(mskDtDoc.Text)) & ", ainda não foi lançado !!!", vbOKOnly + vbExclamation, "Aviso"
        End If
        BREC.Close
    
    End With
    
End Sub

Private Sub grdAPONT_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    With grdAPONT
        
        If (.Rows - 1) = 0 Then Exit Sub
        If Row = 0 Then Exit Sub
        
        Select Case Col
               Case conCOL_PRODOP_CODLINHA
        End Select
    
    End With

End Sub

Private Sub grdAPONT_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    Call MarcaLinha(grdAPONT, NewRowSel)
End Sub

Private Sub grdAPONT_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    With grdAPONT
        Select Case Col
               Case conCOL_PRODOP_CODLINHA, _
                    conCOL_PRODOP_DESCLINHA, _
                    conCOL_PRODOP_QTDEAPONT, _
                    conCOL_PRODOP_IDLINHA, _
                    conCOL_PRODOP_IDINTERNO, _
                    conCOL_PRODOP_Action2Do
                    Cancel = True
               Case Else
                   .ComboList = ""
               End Select
    End With
    Exit Sub

End Sub

Private Sub ConfGrdOPAPONT()

    With grdOPAPOT

       .Cols = conColumnsIn_PRODOPAP
       .Rows = 1
       .FixedCols = 0
       .FormatString = conCOL_PRODOPAP_FormatString
       
       .AutoSizeMouse = False

       .AllowUserResizing = flexResizeNone
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_DTPROG) = ""
       .ColDataType(conCOL_PRODOPAP_DTPROG) = flexDTDate
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_CODOP) = ""
       .ColDataType(conCOL_PRODOPAP_CODOP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_CODROT) = ""
       .ColDataType(conCOL_PRODOPAP_CODROT) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_DESCROT) = ""
       .ColDataType(conCOL_PRODOPAP_DESCROT) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_NECK) = ""
       .ColDataType(conCOL_PRODOPAP_NECK) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_FECHAMENTO) = ""
       .ColDataType(conCOL_PRODOPAP_FECHAMENTO) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_COMPONENTES) = ""
       .ColDataType(conCOL_PRODOPAP_COMPONENTES) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_QTDEPED) = ""
       .ColDataType(conCOL_PRODOPAP_QTDEPED) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_QTDEPROD) = ""
       .ColDataType(conCOL_PRODOPAP_QTDEPROD) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_QTDFOLHAS) = ""
       .ColDataType(conCOL_PRODOPAP_QTDFOLHAS) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_PESO) = ""
       .ColDataType(conCOL_PRODOPAP_PESO) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_CODLINHA) = ""
       .ColDataType(conCOL_PRODOPAP_CODLINHA) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_IDLINHA) = ""
       .ColDataType(conCOL_PRODOPAP_IDLINHA) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_IDINTPROG) = ""
       .ColDataType(conCOL_PRODOPAP_IDINTPROG) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_IDINTOP) = ""
       .ColDataType(conCOL_PRODOPAP_IDINTOP) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_IDPRODUTO) = ""
       .ColDataType(conCOL_PRODOPAP_IDPRODUTO) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_IDINTERNO) = ""
       .ColDataType(conCOL_PRODOPAP_IDINTERNO) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_Action2Do) = ""
       .ColDataType(conCOL_PRODOPAP_Action2Do) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_PERDAMONT) = ""
       .ColDataType(conCOL_PRODOPAP_PERDAMONT) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_StatusApont) = ""
       .ColDataType(conCOL_PRODOPAP_StatusApont) = flexDTString
       .ColComboList(conCOL_PRODOPAP_StatusApont) = objCADPRODOP.PreenchComboStatus
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_CODFOLHA) = ""
       .ColDataType(conCOL_PRODOPAP_CODFOLHA) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_ESPESS) = ""
       .ColDataType(conCOL_PRODOPAP_ESPESS) = flexDTDouble
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_LARG) = ""
       .ColDataType(conCOL_PRODOPAP_LARG) = flexDTDouble
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_COMP) = ""
       .ColDataType(conCOL_PRODOPAP_COMP) = flexDTDouble
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_QTDECORP) = ""
       .ColDataType(conCOL_PRODOPAP_QTDECORP) = flexDTDouble
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_PERDPROD) = ""
       .ColDataType(conCOL_PRODOPAP_PERDPROD) = flexDTDouble
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_CODPED) = ""
       .ColDataType(conCOL_PRODOPAP_CODPED) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_Marca) = ""
       .ColDataType(conCOL_PRODOPAP_Marca) = flexDTBoolean
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_QTDETOTAPONT) = ""
       .ColDataType(conCOL_PRODOPAP_QTDETOTAPONT) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_STATUSPROGFIN) = ""
       .ColDataType(conCOL_PRODOPAP_STATUSPROGFIN) = flexDTLong
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_HORAINI) = ""
       .ColDataType(conCOL_PRODOPAP_HORAINI) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_HORAFIN) = ""
       .ColDataType(conCOL_PRODOPAP_HORAFIN) = flexDTString
       
       .Cell(flexcpData, 0, conCOL_PRODOPAP_TOTALHORAS) = ""
       .ColDataType(conCOL_PRODOPAP_TOTALHORAS) = flexDTString
       
       .ColWidth(conCOL_PRODOPAP_DTPROG) = 800
       .ColWidth(conCOL_PRODOPAP_CODOP) = 900
       .ColWidth(conCOL_PRODOPAP_CODROT) = 1000
       .ColWidth(conCOL_PRODOPAP_DESCROT) = 4000
       .ColWidth(conCOL_PRODOPAP_NECK) = 500
       .ColWidth(conCOL_PRODOPAP_FECHAMENTO) = 500
       .ColWidth(conCOL_PRODOPAP_COMPONENTES) = 500
       .ColWidth(conCOL_PRODOPAP_QTDEPED) = 950
       .ColWidth(conCOL_PRODOPAP_QTDEPROD) = 900
       .ColWidth(conCOL_PRODOPAP_QTDFOLHAS) = 0
       .ColWidth(conCOL_PRODOPAP_PESO) = 0
       .ColWidth(conCOL_PRODOPAP_CODLINHA) = 0
       .ColWidth(conCOL_PRODOPAP_IDLINHA) = 0
       .ColWidth(conCOL_PRODOPAP_IDINTPROG) = 0
       .ColWidth(conCOL_PRODOPAP_IDINTOP) = 0
       .ColWidth(conCOL_PRODOPAP_IDPRODUTO) = 0
       .ColWidth(conCOL_PRODOPAP_IDINTERNO) = 0
       .ColWidth(conCOL_PRODOPAP_Action2Do) = 0
       .ColWidth(conCOL_PRODOPAP_PERDAMONT) = 700
       .ColWidth(conCOL_PRODOPAP_StatusApont) = 1000
       
       .ColWidth(conCOL_PRODOPAP_CODFOLHA) = 0
       .ColWidth(conCOL_PRODOPAP_ESPESS) = 0
       .ColWidth(conCOL_PRODOPAP_LARG) = 0
       .ColWidth(conCOL_PRODOPAP_COMP) = 0
       .ColWidth(conCOL_PRODOPAP_QTDECORP) = 0
       .ColWidth(conCOL_PRODOPAP_PERDPROD) = 0
       .ColWidth(conCOL_PRODOPAP_CODPED) = 0
       .ColWidth(conCOL_PRODOPAP_Marca) = 200
       .ColWidth(conCOL_PRODOPAP_QTDETOTAPONT) = 0
       .ColWidth(conCOL_PRODOPAP_STATUSPROGFIN) = 0
       .ColWidth(conCOL_PRODOPAP_HORAINI) = 700
       .ColWidth(conCOL_PRODOPAP_HORAFIN) = 700
       .ColWidth(conCOL_PRODOPAP_TOTALHORAS) = 0
      
       .Editable = flexEDKbdMouse
       .AllowSelection = False
       .HighLight = flexHighlightWithFocus
       .SelectionMode = flexSelectionByRow
       .BackColor = &H80000018
       .ForeColor = vbBlack
       .FontName = "Arial"
       .FontBold = True
       .FontSize = 7

    End With
    
End Sub

Private Sub grdAPONT_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
        Call Desmarca(grdAPONT, OldRowSel)
End Sub

Private Sub grdAPONT_Click()
    Call MostraDados
End Sub

Private Sub grdAPONT_RowColChange()
    Call MostraDados
End Sub

Private Sub grdOPAPOT_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    Dim lngSALDOOP      As Long
    Dim strTOTALPERIODO As String
    
    With grdOPAPOT
        
        If (.Rows - 1) = 0 Then Exit Sub
        If Row = 0 Then Exit Sub
        
        Select Case Col
               Case conCOL_PRODOPAP_CODOP
               Case conCOL_PRODOPAP_QTDEPROD
                    Call CalcTotApont
                    
                    lngSALDOOP = PegaSaldoOP(.Cell(flexcpText, Row, conCOL_PRODOPAP_IDINTPROG), .Cell(flexcpText, Row, conCOL_PRODOPAP_QTDEPED))
                    If lngSALDOOP <= 0 Then
                        Call MudandoStatusOPs(lngSALDOOP, .Cell(flexcpText, Row, conCOL_PRODOPAP_IDINTPROG))
                    End If
                    
               Case conCOL_PRODOPAP_QTDFOLHAS
                    Call CalcTotApont
               Case conCOL_PRODOPAP_PESO
                    Call CalcTotApont
                    If Len(Trim(.Cell(flexcpText, Row, Col))) > 0 Then .Cell(flexcpText, Row, Col) = Format(CDbl(.Cell(flexcpText, Row, Col)), "#,####0.0000")
               Case conCOL_PRODOPAP_PERDAMONT
                    Call CalcTotApont
               Case conCOL_PRODOPAP_HORAINI
                    
                    strTOTALPERIODO = objBLBFunc.CalcTempo(.Cell(flexcpText, Row, conCOL_PRODOPAP_HORAINI), .Cell(flexcpText, Row, conCOL_PRODOPAP_HORAFIN))
                    If strTOTALPERIODO = "00:00" Then
                        .Cell(flexcpText, Row, conCOL_PRODOPAP_TOTALHORAS) = Empty
                    Else
                        .Cell(flexcpText, Row, conCOL_PRODOPAP_TOTALHORAS) = Format(CDate(strTOTALPERIODO), "HH:MM")
                    End If
               
               Case conCOL_PRODOPAP_HORAFIN
               
                    strTOTALPERIODO = objBLBFunc.CalcTempo(.Cell(flexcpText, Row, conCOL_PRODOPAP_HORAINI), .Cell(flexcpText, Row, conCOL_PRODOPAP_HORAFIN))
                    If strTOTALPERIODO = "00:00" Then
                        .Cell(flexcpText, Row, conCOL_PRODOPAP_TOTALHORAS) = Empty
                    Else
                        .Cell(flexcpText, Row, conCOL_PRODOPAP_TOTALHORAS) = Format(CDate(strTOTALPERIODO), "HH:MM")
                    End If
               
                    Call CalcTotApont
                    If Row < (.Rows - 1) Then
                        Call Command9_Click
                        .Row = (Row + 1)
                    Else
                        Call Command9_Click
                        .Row = (Row + 1)
                    End If
        End Select
    
        
    
    End With

End Sub

Private Sub grdOPAPOT_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    With grdOPAPOT
        Select Case Col
               Case conCOL_PRODOPAP_CODROT, _
                    conCOL_PRODOPAP_DESCROT, _
                    conCOL_PRODOPAP_NECK, _
                    conCOL_PRODOPAP_COMPONENTES, _
                    conCOL_PRODOPAP_FECHAMENTO, _
                    conCOL_PRODOPAP_QTDEPED, _
                    conCOL_PRODOPAP_CODLINHA, _
                    conCOL_PRODOPAP_IDLINHA, _
                    conCOL_PRODOPAP_IDINTPROG, _
                    conCOL_PRODOPAP_IDINTOP, _
                    conCOL_PRODOPAP_IDPRODUTO, _
                    conCOL_PRODOPAP_IDINTERNO, _
                    conCOL_PRODOPAP_Action2Do, _
                    conCOL_PRODOPAP_StatusApont, _
                    conCOL_PRODOPAP_CODPED, _
                    conCOL_PRODOPAP_QTDETOTAPONT, _
                    conCOL_PRODOPAP_STATUSPROGFIN, _
                    conCOL_PRODOPAP_TOTALHORAS
                    Cancel = True
               Case conCOL_PRODOPAP_DTPROG, _
                    conCOL_PRODOPAP_CODOP, _
                    conCOL_PRODOPAP_QTDEPROD, _
                    conCOL_PRODOPAP_QTDFOLHAS, _
                    conCOL_PRODOPAP_PESO, _
                    conCOL_PRODOPAP_PERDAMONT, _
                    conCOL_PRODOPAP_Marca
                    If cTipOper = "C" Then
                        Cancel = True
                    Else
                        If Col = conCOL_PRODOPAP_StatusApont Then
                            If Len(Trim(.Cell(flexcpText, Row, conCOL_PRODOPAP_QTDEPROD))) = 0 Then
                                Cancel = True
                            ElseIf .Cell(flexcpText, Row, conCOL_PRODOPAP_QTDEPROD) = 0 Then
                               Cancel = True
                            ElseIf .Cell(flexcpText, Row, conCOL_PRODOPAP_QTDFOLHAS) = 0 Then
                                Cancel = True
                            ElseIf Len(Trim(.Cell(flexcpText, Row, conCOL_PRODOPAP_QTDFOLHAS))) = 0 Then
                                Cancel = True
                            ElseIf .Cell(flexcpText, Row, conCOL_PRODOPAP_PESO) = 0 Then
                                Cancel = True
                            ElseIf Len(Trim(.Cell(flexcpText, Row, conCOL_PRODOPAP_PESO))) = 0 Then
                                Cancel = True
                            End If
                        ElseIf Col = conCOL_PRODOPAP_CODOP Or _
                               Col = conCOL_PRODOPAP_QTDEPROD Or _
                               Col = conCOL_PRODOPAP_PERDAMONT Or _
                               Col = conCOL_PRODOPAP_HORAINI Or _
                               Col = conCOL_PRODOPAP_HORAFIN Then
                                If Len(Trim(Replace(Replace(.Cell(flexcpText, Row, conCOL_PRODOPAP_DTPROG), "/", ""), "_", ""))) = 0 Then
                                    Cancel = True
                                    Exit Sub
                                ElseIf Len(Trim(Replace(Replace(.Cell(flexcpText, Row, conCOL_PRODOPAP_DTPROG), "/", ""), "_", ""))) < 8 Then
                                    Cancel = True
                                    Exit Sub
                                End If
                        End If
                    End If
               Case Else
                   .ComboList = ""
               End Select
    End With
    Exit Sub

End Sub


Private Sub IncRegGrid()
   
    If objBLBFunc.FcExisteLinhaVazia(grdOPAPOT, conCOL_PRODOPAP_DTPROG) = False Then
        Call objBLBFunc.RemoveLinhaVazia(grdOPAPOT, conCOL_PRODOPAP_DTPROG)
    End If
    If objBLBFunc.FcExisteLinhaVazia(grdOPAPOT, conCOL_PRODOPAP_CODOP) = False Then
        Call objBLBFunc.RemoveLinhaVazia(grdOPAPOT, conCOL_PRODOPAP_CODOP)
    End If
    
    If (grdAPONT.Rows - 1) = 0 Then
        MsgBox "ATENÇÂO - Não existe dados na Gride de Linha !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    If (grdAPONT.Row) = 0 Then
        MsgBox "ATENÇÂO - Não foi selecionado uma Linha de Produtos na Gride de Linha !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    
    With grdOPAPOT
        
        .AddItem "" & vbTab & "" & vbTab & "" & vbTab & _
                 "" & vbTab & "" & vbTab & _
                 "" & vbTab & "" & vbTab & _
                 "" & vbTab & "" & vbTab & _
                 "" & vbTab & "" & vbTab & _
                 grdAPONT.Cell(flexcpText, grdAPONT.Row, conCOL_PRODOP_CODLINHA) & vbTab & grdAPONT.Cell(flexcpText, grdAPONT.Row, conCOL_PRODOP_IDLINHA) & vbTab & _
                 "" & vbTab & "" & vbTab & _
                 "" & vbTab & "" & vbTab & _
                 dacEnumUpdateAction_Insert & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 "" & vbTab & _
                 0 & vbTab & _
                 "" & vbTab & _
                 0 & vbTab & _
                 ""
        
        .Cell(flexcpFontSize, (.Rows - 1), conCOL_PRODOPAP_DESCROT) = 7
        .Cell(flexcpText, (.Rows - 1), conCOL_PRODOPAP_IDINTERNO) = ((.Rows - 1) * -1)
        
        Call PintaColunasEditaveis
    
        .SetFocus
        .Row = (.Rows - 1)
        .Col = conCOL_PRODOPAP_DTPROG
        .Select .Row, .Col
        .EditCell
    
    End With

End Sub


Private Sub grdOPAPOT_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     With grdOPAPOT
          Select Case Col
                    Case conCOL_PRODOPAP_CODOP, _
                         conCOL_PRODOPAP_QTDEPROD, _
                         conCOL_PRODOPAP_QTDFOLHAS, _
                         conCOL_PRODOPAP_PERDAMONT
                         KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsLong)
                    Case conCOL_PRODOPAP_PESO
                        KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsDouble)
                    Case conCOL_PRODOPAP_DTPROG
                        ''KeyAscii = objBLBFunc.MaskNumber(.EditText, KeyAscii, 0, myvarAsDate)
          
          End Select
     End With
End Sub

Private Sub MostraDados()
    
    If (grdAPONT.Rows - 1) = 0 Then Exit Sub
    If (grdAPONT.Row) <= 0 Then Exit Sub

    Call objBLBFunc.CarregaDadosGrdFilho(grdOPAPOT, conCOL_PRODOPAP_Action2Do, conCOL_PRODOPAP_CODLINHA, grdAPONT.Cell(flexcpText, grdAPONT.Row, conCOL_PRODOP_CODLINHA))

End Sub

Private Sub grdOPAPOT_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

     Dim lngQTDJA_APONT     As Long
     Dim lngQTDJA_APONT_01  As Long
     Dim lngSALDOOP         As Long
     Dim lngTOTALGAPNT      As Long
     Dim dtDTBAIXAPROG      As Date
     Dim lngIDINT           As Long
     
     With grdOPAPOT
          Select Case Col
                 Case conCOL_PRODOPAP_DTPROG
                 
                        If .EditText = Empty Then Exit Sub
                        If Len(Trim(Replace(Replace(.EditText, "/", ""), "_", ""))) = 0 Then Exit Sub
                        If Len(Trim(Replace(Replace(.EditText, "/", ""), "_", ""))) < 8 Then
                            MsgBox "DATA inválida !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                        
                        If Len(Trim(.EditText)) = 8 Then .EditText = Format(.EditText, "##/##/####")
                        If Not IsDate(.EditText) Then
                            MsgBox "DATA inválida !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                        
                        Call PosColCapac(conCOL_PRODOPAP_CODOP, Row)
                        
                 Case conCOL_PRODOPAP_CODOP
                 
                        If .EditText = Empty Then Exit Sub
                        If Len(Trim(.EditText)) = 0 Then Exit Sub
                        If Not IsNumeric(.EditText) Then Exit Sub
                        
                        If PegaOPProg(.EditText, Row, .Cell(flexcpText, Row, conCOL_PRODOPAP_DTPROG)) = False Then
                            Cancel = True
                            Exit Sub
                        End If
                        
                        If VerifOPLanc(.Cell(flexcpText, Row, conCOL_PRODOPAP_IDINTPROG), CDate(mskDtDoc.Text)) = True Then
                            Call LimpaCamposGrid(Row)
                            Cancel = True
                            Exit Sub
                        End If
                        
                        Call PosColCapac(conCOL_PRODOPAP_QTDEPROD, Row)
                        
                 Case conCOL_PRODOPAP_QTDEPROD
                        
                        If .EditText = Empty Then Exit Sub
                        If Len(Trim(.EditText)) = 0 Then Exit Sub
                        If Not IsNumeric(.EditText) Then Exit Sub
                        
                        dtDTBAIXAPROG = CDate(mskDtDoc.Text)
                        lngIDINT = .Cell(flexcpText, Row, conCOL_PRODOPAP_IDINTERNO)
                        lngQTDJA_APONT = PegaQTD_JAAPONT(.Cell(flexcpText, Row, conCOL_PRODOPAP_IDINTPROG), CLng(.EditText), .Cell(flexcpText, Row, conCOL_PRODOPAP_IDINTOP))
                        lngQTDJA_APONT_01 = PegaJaProg_02(.Cell(flexcpText, Row, conCOL_PRODOPAP_IDINTOP), .Cell(flexcpText, Row, conCOL_PRODOPAP_IDINTPROG), dtDTBAIXAPROG)

                        
                        lngTOTALGAPNT = (lngQTDJA_APONT + lngQTDJA_APONT_01)
                        '' Verificando os 10%
                        If Calcula10porc(Row, lngTOTALGAPNT) = False Then
                            MsgBox "ATENÇÃO" & vbCrLf & _
                                   "Está tentando apontar a mais de 10% da quantidade da OP !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                        
                        If Len(Trim(.Cell(flexcpText, Row, conCOL_PRODOPAP_QTDEPROD))) > 0 Then
                            If CLng(.Cell(flexcpText, Row, conCOL_PRODOPAP_QTDEPROD)) <> CLng(.EditText) Then
                                If .Cell(flexcpText, Row, conCOL_PRODOPAP_Action2Do) = dacEnumUpdateAction_Ignore Then
                                   .Cell(flexcpText, Row, conCOL_PRODOPAP_Action2Do) = dacEnumUpdateAction_update
                                End If
                            End If
                        End If
                        
                        
                        .Cell(flexcpText, Row, conCOL_PRODOPAP_QTDETOTAPONT) = lngTOTALGAPNT
                        
                        If lngTOTALGAPNT >= CLng(.Cell(flexcpText, Row, conCOL_PRODOPAP_QTDEPED)) Then
                            .Cell(flexcpText, Row, conCOL_PRODOPAP_StatusApont) = 1
                        ElseIf lngTOTALGAPNT < CLng(.Cell(flexcpText, Row, conCOL_PRODOPAP_QTDEPED)) Then
                            .Cell(flexcpText, Row, conCOL_PRODOPAP_StatusApont) = 2
                        End If
                        
                        Call PosColCapac(conCOL_PRODOPAP_PERDAMONT, Row)
                        
                Case conCOL_PRODOPAP_QTDFOLHAS
          
                        If .EditText = Empty Then Exit Sub
                        If Len(Trim(.EditText)) = 0 Then Exit Sub
                        If Not IsNumeric(.EditText) Then Exit Sub
                        
                        If IsNumeric(.Cell(flexcpText, Row, conCOL_PRODOPAP_QTDFOLHAS)) = True Then
                            If CLng(.Cell(flexcpText, Row, conCOL_PRODOPAP_QTDFOLHAS)) <> CLng(.EditText) Then
                                If .Cell(flexcpText, Row, conCOL_PRODOPAP_Action2Do) = dacEnumUpdateAction_Ignore Then
                                   .Cell(flexcpText, Row, conCOL_PRODOPAP_Action2Do) = dacEnumUpdateAction_update
                                End If
                            End If
                        Else
                            If .Cell(flexcpText, Row, conCOL_PRODOPAP_Action2Do) = dacEnumUpdateAction_Ignore Then
                               .Cell(flexcpText, Row, conCOL_PRODOPAP_Action2Do) = dacEnumUpdateAction_update
                            End If
                        End If
                        
                        
                        Call PosColCapac(conCOL_PRODOPAP_PESO, Row)
          
                Case conCOL_PRODOPAP_PESO
                        
                        If .EditText = Empty Then Exit Sub
                        If Len(Trim(.EditText)) = 0 Then Exit Sub
                        If Not IsNumeric(.EditText) Then Exit Sub
                        
                        If IsNumeric(.Cell(flexcpText, Row, conCOL_PRODOPAP_PESO)) = True Then
                            If CLng(.Cell(flexcpText, Row, conCOL_PRODOPAP_PESO)) <> CLng(.EditText) Then
                                If .Cell(flexcpText, Row, conCOL_PRODOPAP_Action2Do) = dacEnumUpdateAction_Ignore Then
                                   .Cell(flexcpText, Row, conCOL_PRODOPAP_Action2Do) = dacEnumUpdateAction_update
                                End If
                            End If
                        Else
                            If .Cell(flexcpText, Row, conCOL_PRODOPAP_Action2Do) = dacEnumUpdateAction_Ignore Then
                               .Cell(flexcpText, Row, conCOL_PRODOPAP_Action2Do) = dacEnumUpdateAction_update
                            End If
                        End If
                        
                        Call PosColCapac(conCOL_PRODOPAP_PERDAMONT, Row)
          
                Case conCOL_PRODOPAP_PERDAMONT
                        
                        If .EditText = Empty Then Exit Sub
                        If Len(Trim(.EditText)) = 0 Then Exit Sub
                        If Not IsNumeric(.EditText) Then Exit Sub
                        
                        If IsNumeric(.Cell(flexcpText, Row, conCOL_PRODOPAP_PERDAMONT)) = True Then
                            If CLng(.Cell(flexcpText, Row, conCOL_PRODOPAP_PERDAMONT)) <> CLng(.EditText) Then
                                If .Cell(flexcpText, Row, conCOL_PRODOPAP_Action2Do) = dacEnumUpdateAction_Ignore Then
                                   .Cell(flexcpText, Row, conCOL_PRODOPAP_Action2Do) = dacEnumUpdateAction_update
                                End If
                            End If
                        Else
                            If .Cell(flexcpText, Row, conCOL_PRODOPAP_PERDAMONT) <> .EditText Then
                                If .Cell(flexcpText, Row, conCOL_PRODOPAP_Action2Do) = dacEnumUpdateAction_Ignore Then
                                   .Cell(flexcpText, Row, conCOL_PRODOPAP_Action2Do) = dacEnumUpdateAction_update
                                End If
                            End If
                        End If
                        
                        Call PosColCapac(conCOL_PRODOPAP_HORAINI, Row)
                        
                Case conCOL_PRODOPAP_HORAINI
                
                        If .EditText = Empty Then
                            .Cell(flexcpText, Row, conCOL_PRODOPAP_HORAFIN) = Empty
                            .Cell(flexcpText, Row, conCOL_PRODOPAP_TOTALHORAS) = Empty
                            Exit Sub
                        End If
                        If Len(Trim(Replace(Replace(.EditText, ":", ""), "_", ""))) = 0 Then Exit Sub
                        If Len(Trim(Replace(Replace(.EditText, ":", ""), "_", ""))) < 4 Then Exit Sub
                        
                        If Not IsDate(.EditText) Then
                            MsgBox "ATENÇÃO" & vbCrLf & _
                                   "Hora Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                        
                        If .Cell(flexcpText, Row, conCOL_PRODOPAP_HORAINI) <> .EditText Then
                            If .Cell(flexcpText, Row, conCOL_PRODOPAP_Action2Do) = dacEnumUpdateAction_Ignore Then
                               .Cell(flexcpText, Row, conCOL_PRODOPAP_Action2Do) = dacEnumUpdateAction_update
                            End If
                        End If
                        
                        
                        Call PosColCapac(conCOL_PRODOPAP_HORAFIN, Row)
          
                Case conCOL_PRODOPAP_HORAFIN
                        
                        If .EditText = Empty Then
                            .Cell(flexcpText, Row, conCOL_PRODOPAP_HORAINI) = Empty
                            .Cell(flexcpText, Row, conCOL_PRODOPAP_TOTALHORAS) = Empty
                            Exit Sub
                        End If
                        If Len(Trim(Replace(Replace(.EditText, ":", ""), "_", ""))) = 0 Then Exit Sub
                        If Len(Trim(Replace(Replace(.EditText, ":", ""), "_", ""))) < 4 Then Exit Sub
                        
                        If Not IsDate(.EditText) Then
                            MsgBox "ATENÇÃO" & vbCrLf & _
                                   "Hora Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                        
                        If Len(Trim(Replace(.Cell(flexcpText, Row, conCOL_PRODOPAP_HORAINI), ":", ""))) = 0 Then
                            MsgBox "ATENÇÃO" & vbCrLf & _
                                   "Informe a Hora Inicial !!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                        End If
                        
                        
                        If CDate(.Cell(flexcpText, Row, conCOL_PRODOPAP_HORAINI)) > CDate(.EditText) Then
                            MsgBox "ATENÇÃO" & vbCrLf & _
                                   "A Hora Inicial não pode ser maior que a Hora Final!!!", vbOKOnly + vbExclamation, "Aviso"
                            Cancel = True
                            Exit Sub
                            .Cell(flexcpText, Row, conCOL_PRODOPAP_HORAINI) = Empty
                            .Cell(flexcpText, Row, conCOL_PRODOPAP_TOTALHORAS) = Empty
                        End If
                        
                        If .Cell(flexcpText, Row, conCOL_PRODOPAP_HORAFIN) <> .EditText Then
                            If .Cell(flexcpText, Row, conCOL_PRODOPAP_Action2Do) = dacEnumUpdateAction_Ignore Then
                               .Cell(flexcpText, Row, conCOL_PRODOPAP_Action2Do) = dacEnumUpdateAction_update
                            End If
                        End If
                        
                        
                        Call PosColCapac(conCOL_PRODOPAP_DTPROG, Row)
          End Select
     End With

End Sub

Private Function PegaOPProg(strCODOP As String, lngLINHA As Long, dtDTPROG As Date) As Boolean

    PegaOPProg = False
    
    If Len(Trim(strCODOP)) = 0 Then Exit Function
    If Len(Trim(lngLINHA)) <= 0 Then Exit Function
    
    Dim lngPESQ             As Long
    Dim lngSALDOOP          As Long
    Dim lngIDINT            As Long
    Dim lngQTDJA_APONT      As Long
    Dim lngQTDJA_APONT_01   As Long
    Dim lngTOTALGAPNT       As Long
    
    Dim boolNAOINC          As Boolean
    Dim dtDTBAIXAPROG       As Date
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       VPCP.SGI_CODINTENO" & vbCrLf
    sSql = sSql & "      ,VPCP.SGI_IDINTERNO" & vbCrLf
    sSql = sSql & "      ,VPCP.SGI_QTDEPROD" & vbCrLf
    sSql = sSql & "      ,VPCP.SGI_CODPED" & vbCrLf
    
    sSql = sSql & "      ,PROD.SGI_CODIGO       As SGI_CODPROD" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_DESCRICAO    As SGI_DESCPROD" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_CODLINPROD   As SGI_CODLINHA" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_IDPRODUTO    " & vbCrLf
    sSql = sSql & "      ,PROD.SGI_NECKIN       " & vbCrLf
    sSql = sSql & "      ,PROD.SGI_VernTampa    " & vbCrLf
    sSql = sSql & "      ,ORDP.SGI_FECHTPFU     " & vbCrLf
    sSql = sSql & "      ,ORDP.SGI_QTDE         As SGI_QTDEOP" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADMOVPCP" & strNOMTABELA & " VPCP" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO      PROD" & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMTABELA & " ORDP" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       VPCP.SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And VPCP.SGI_CODOP     = " & Trim(Replace(Replace(strCODOP, ",", ""), ".", "")) & vbCrLf
    sSql = sSql & "   And VPCP.SGI_DATAPROG  = '" & Format(dtDTPROG, "MM/DD/YYYY") & "'" & vbCrLf
    
    sSql = sSql & "   And PROD.SGI_FILIAL    = VPCP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PROD.SGI_IDPRODUTO = VPCP.SGI_IDPRODUTO" & vbCrLf
    
    sSql = sSql & "   And ORDP.SGI_FILIAL    = VPCP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And ORDP.SGI_CODIGO    = VPCP.SGI_CODOP" & vbCrLf
    sSql = sSql & "   And ORDP.SGI_IDPRODUTO = VPCP.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "   And ORDP.SGI_IDPAI     = VPCP.SGI_IDINTERNO"
    
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC2.EOF() Then
        With grdOPAPOT
            If BREC2!SGI_CODLINHA <> CLng(.Cell(flexcpText, lngLINHA, conCOL_PRODOPAP_CODLINHA)) Then
                MsgBox "ATENÇÂO" & vbCrLf & "Esta OP não é desta linha !!!", vbOKOnly + vbExclamation, "Aviso"
                Call LimpaCamposGrid(lngLINHA)
            Else
                
                boolNAOINC = True
                dtDTBAIXAPROG = CDate(mskDtDoc.Text)
                lngIDINT = .Cell(flexcpText, lngLINHA, conCOL_PRODOPAP_IDINTERNO)
                
                lngQTDJA_APONT = PegaQTD_JAAPONT(BREC2!SGI_CODINTENO, 0, BREC2!SGI_IDINTERNO)
                lngQTDJA_APONT_01 = PegaJaProg_02(BREC2!SGI_IDINTERNO, BREC2!SGI_CODINTENO, dtDTBAIXAPROG)
                lngTOTALGAPNT = (lngQTDJA_APONT + lngQTDJA_APONT_01)
                
                lngSALDOOP = (BREC2!SGI_QTDEPROD - lngTOTALGAPNT)
                
                If lngSALDOOP <= 0 Then
                   MsgBox "ATENÇÃO" & vbCrLf & "A OP " & Trim(strCODOP) & " Já foi toda apontada !!!", vbOKOnly + vbExclamation, "Aviso"
                   Call LimpaCamposGrid(lngLINHA)
                   boolNAOINC = False
                End If
                
                If boolNAOINC = True Then
                    PegaOPProg = True
                    .Cell(flexcpText, lngLINHA, conCOL_PRODOPAP_CODROT) = BREC2!SGI_CODPROD
                    .Cell(flexcpText, lngLINHA, conCOL_PRODOPAP_DESCROT) = BREC2!SGI_DESCPROD
                    .Cell(flexcpText, lngLINHA, conCOL_PRODOPAP_QTDEPED) = (BREC2!SGI_QTDEPROD)
                    .Cell(flexcpText, lngLINHA, conCOL_PRODOPAP_CODPED) = BREC2!SGI_CODPED
                    
                    .Cell(flexcpText, lngLINHA, conCOL_PRODOPAP_IDINTPROG) = BREC2!SGI_CODINTENO
                    .Cell(flexcpText, lngLINHA, conCOL_PRODOPAP_IDINTOP) = BREC2!SGI_IDINTERNO
                    .Cell(flexcpText, lngLINHA, conCOL_PRODOPAP_IDPRODUTO) = BREC2!SGI_IDPRODUTO
                    .Cell(flexcpText, lngLINHA, conCOL_PRODOPAP_NECK) = IIf(BREC2!SGI_NECKIN = 1, "Sim", "Não")
                    .Cell(flexcpText, lngLINHA, conCOL_PRODOPAP_FECHAMENTO) = PegaFechamentoTampaFuro(BREC2!SGI_FECHTPFU)
                    .Cell(flexcpText, lngLINHA, conCOL_PRODOPAP_COMPONENTES) = IIf(IsNull(BREC2!SGI_VernTampa) = False, PegaComp(BREC2!SGI_VernTampa), "")
                End If
                
            End If
        End With
    Else
        MsgBox "ATENÇÂO" & vbCrLf & "Esta OP não foi Programada neste dia, verifique !!!", vbOKOnly + vbExclamation, "Aviso"
        Call LimpaCamposGrid(lngLINHA)
    End If
    BREC2.Close
    
    
End Function

Private Sub LimpaCamposGrid(lngLINHA As Long)

    If lngLINHA = 0 Then Exit Sub
    If lngLINHA <= 0 Then Exit Sub
    
    With grdOPAPOT
        .Cell(flexcpText, lngLINHA, conCOL_PRODOPAP_DTPROG) = Empty
        .Cell(flexcpText, lngLINHA, conCOL_PRODOPAP_CODROT) = Empty
        .Cell(flexcpText, lngLINHA, conCOL_PRODOPAP_DESCROT) = Empty
        .Cell(flexcpText, lngLINHA, conCOL_PRODOPAP_NECK) = Empty
        .Cell(flexcpText, lngLINHA, conCOL_PRODOPAP_COMPONENTES) = Empty
        .Cell(flexcpText, lngLINHA, conCOL_PRODOPAP_FECHAMENTO) = Empty
        .Cell(flexcpText, lngLINHA, conCOL_PRODOPAP_QTDEPED) = Empty
        .Cell(flexcpText, lngLINHA, conCOL_PRODOPAP_QTDEPROD) = Empty
        .Cell(flexcpText, lngLINHA, conCOL_PRODOPAP_IDINTPROG) = Empty
        .Cell(flexcpText, lngLINHA, conCOL_PRODOPAP_IDINTOP) = Empty
        .Cell(flexcpText, lngLINHA, conCOL_PRODOPAP_IDPRODUTO) = Empty
        .Cell(flexcpText, lngLINHA, conCOL_PRODOPAP_PERDAMONT) = Empty
        .Cell(flexcpText, lngLINHA, conCOL_PRODOPAP_HORAINI) = Empty
        .Cell(flexcpText, lngLINHA, conCOL_PRODOPAP_HORAFIN) = Empty
        .Cell(flexcpText, lngLINHA, conCOL_PRODOPAP_TOTALHORAS) = Empty
    End With

End Sub

Private Sub CalcTotApont()

    If (grdOPAPOT.Rows - 1) = 0 Then Exit Sub
    
    Dim I               As Long
    Dim j               As Long
    Dim strCODLINHA     As String
    Dim lngQTDAPONT     As Long
    Dim lngTOTAPONT     As Long
    
    With grdAPONT
        For I = 1 To (.Rows - 1)
            strCODLINHA = Trim(.Cell(flexcpText, I, conCOL_PRODOP_CODLINHA))
            lngQTDAPONT = 0
            lngTOTAPONT = 0
            
            For j = 1 To (grdOPAPOT.Rows - 1)
                If strCODLINHA = Trim(grdOPAPOT.Cell(flexcpText, j, conCOL_PRODOPAP_CODLINHA)) Then
                    If grdOPAPOT.Cell(flexcpText, j, conCOL_PRODOPAP_Action2Do) <> dacEnumUpdateAction_delete Then
                        If Len(Trim(grdOPAPOT.Cell(flexcpText, j, conCOL_PRODOPAP_QTDEPROD))) > 0 Then lngQTDAPONT = CLng(grdOPAPOT.Cell(flexcpText, j, conCOL_PRODOPAP_QTDEPROD))
                        lngTOTAPONT = (lngTOTAPONT + lngQTDAPONT)
                    End If
                End If
            Next j
            
            .Cell(flexcpText, I, conCOL_PRODOP_QTDEAPONT) = lngTOTAPONT
            
        Next I
    End With

End Sub

Private Sub PopGrdOPApont()

    Dim I               As Integer
    Dim lngQTDJA_APONT  As Long
    
    If IsArray(arrPROGRAMADO) Then
        With grdOPAPOT
            For I = 1 To UBound(arrPROGRAMADO)
                
                .AddItem Format(CDate(arrPROGRAMADO(I, 15)), "DD/MM/YYYY") & vbTab & arrPROGRAMADO(I, 1) & vbTab & _
                         "" & vbTab & "" & vbTab & _
                         "" & vbTab & "" & vbTab & _
                         "" & vbTab & arrPROGRAMADO(I, 2) & vbTab & _
                         arrPROGRAMADO(I, 3) & vbTab & arrPROGRAMADO(I, 12) & vbTab & _
                         "" & vbTab & arrPROGRAMADO(I, 7) & vbTab & _
                         arrPROGRAMADO(I, 6) & vbTab & arrPROGRAMADO(I, 4) & vbTab & _
                         arrPROGRAMADO(I, 5) & vbTab & arrPROGRAMADO(I, 8) & vbTab & _
                         arrPROGRAMADO(I, 9) & vbTab & _
                         dacEnumUpdateAction_Ignore & vbTab & _
                         arrPROGRAMADO(I, 11) & vbTab & _
                         arrPROGRAMADO(I, 16) & vbTab & _
                         arrPROGRAMADO(I, 17) & vbTab & _
                         arrPROGRAMADO(I, 10) & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         "" & vbTab & _
                         arrPROGRAMADO(I, 14) & vbTab & _
                         0 & vbTab & _
                         "" & vbTab & _
                         0 & vbTab & _
                         arrPROGRAMADO(I, 18)
                         
                         
                
                lngQTDJA_APONT = PegaQTD_JAAPONT(.Cell(flexcpText, (.Rows - 1), conCOL_PRODOPAP_IDINTPROG), 0, .Cell(flexcpText, (.Rows - 1), conCOL_PRODOPAP_IDINTOP))
                .Cell(flexcpText, (.Rows - 1), conCOL_PRODOPAP_QTDETOTAPONT) = lngQTDJA_APONT
                
                '' ===================================
                '' Pegando o Produto
                sSql = ""
                
                sSql = "Select " & vbCrLf
                sSql = sSql & "       PROD.SGI_CODIGO       As SGI_CODPROD" & vbCrLf
                sSql = sSql & "      ,PROD.SGI_DESCRICAO    As SGI_DESCPROD" & vbCrLf
                sSql = sSql & "      ,PROD.SGI_NECKIN       " & vbCrLf
                sSql = sSql & "      ,PROD.SGI_VernTampa    " & vbCrLf
                
                sSql = sSql & "      ,ORDP.SGI_FECHTPFU     " & vbCrLf
                 
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_CADPRODUTO PROD" & vbCrLf
                sSql = sSql & "     , SGI_ORDEMPROD" & strTABELA & " ORDP" & vbCrLf
                 
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       PROD.SGI_FILIAL    = " & FILIAL & vbCrLf
                sSql = sSql & "   And PROD.SGI_IDPRODUTO = " & arrPROGRAMADO(I, 8) & vbCrLf
                sSql = sSql & "   And ORDP.SGI_FILIAL    = PROD.SGI_FILIAL" & vbCrLf
                sSql = sSql & "   And ORDP.SGI_IDPRODUTO = PROD.SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "   And ORDP.SGI_CODIGO    = " & arrPROGRAMADO(I, 1)
                
                BREC4.Open sSql, adoBanco_Dados, adOpenDynamic
                If Not BREC4.EOF() Then
                    .Cell(flexcpText, (.Rows - 1), conCOL_PRODOPAP_CODROT) = Trim(BREC4!SGI_CODPROD)
                    .Cell(flexcpText, (.Rows - 1), conCOL_PRODOPAP_DESCROT) = Trim(BREC4!SGI_DESCPROD)
                    .Cell(flexcpText, (.Rows - 1), conCOL_PRODOPAP_NECK) = IIf(BREC4!SGI_NECKIN = 1, "Sim", "Não")
                    .Cell(flexcpText, (.Rows - 1), conCOL_PRODOPAP_FECHAMENTO) = PegaFechamentoTampaFuro(Str(BREC4!SGI_FECHTPFU))
                    .Cell(flexcpText, (.Rows - 1), conCOL_PRODOPAP_COMPONENTES) = IIf(IsNull(BREC4!SGI_VernTampa) = False, PegaComp(BREC4!SGI_VernTampa), "")
                End If
                BREC4.Close
                '' ===================================
            
            Next I
        End With
        Call PintaColunasEditaveis
    End If

End Sub


Private Sub mskDtDoc_GotFocus()
    objBLBFunc.SelecionaCampos mskDtDoc.Name, Me
End Sub

Private Function ConsisteData() As Boolean
    ConsisteData = False
    
    If Len(Trim(Replace(Replace(mskDtDoc.Text, "/", ""), "_", ""))) < 8 Then
        MsgBox "ATENÇÂO" & vbCrLf & "Data Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Function
    ElseIf Not IsDate(mskDtDoc.Text) Then
        MsgBox "ATENÇÂO" & vbCrLf & "Data Inválida !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Function
    End If
    If CDate(mskDtDoc.Text) > Now Then
        MsgBox "ATENÇÂO" & vbCrLf & "Data de apontamento não pode ser maior que a Data do Sistema !!!", vbOKOnly + vbExclamation, "Aviso"
        mskDtDoc.Text = Format(Now, "DD/MM/YYYY")
        Exit Function
    End If
    
    ConsisteData = True
End Function

Private Function ConsisteLanc(strDTPESQ As String) As Boolean

    ConsisteLanc = False
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADAPONTPROG" & strNOMTABELA & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_DATAPONT = " & strDTPESQ
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then ConsisteLanc = True
    BREC10.Close

End Function

Private Function VerifOPLanc(lngIDINTPROG As String, dtDTAPONT As Date) As Boolean

    VerifOPLanc = False
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADAPONTPROG" & strNOMTABELA & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_FILIAL       = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_IDINTPROG    = " & lngIDINTPROG & vbCrLf
    sSql = sSql & "   And SGI_DATAPONT    <> '" & Format(dtDTAPONT, "MM/DD/YYYY") & "'" & vbCrLf
    sSql = sSql & "   And SGI_STATUSAPONT  = 1"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        MsgBox "ATENÇÃO" & vbCrLf & "Esta OP já esta lançada no dia(s) !!!", vbOKOnly + vbExclamation, "Aviso"
        VerifOPLanc = True
    End If
    BREC.Close

End Function

Private Sub PosColCapac(lngPOSCOL As Long, lngPOSROL As Long)
    
On Error GoTo Err_PosCol
    
    With grdOPAPOT
        .SetFocus
        .Row = lngPOSROL
        .Col = lngPOSCOL
    End With
    
    Exit Sub
    
Err_PosCol:
    Call objBLBFunc.Sub_DescErro(Str(Err.Number), Err.Description, cTipOper, "Função : PosCol()", Me.Name, "PosCol()", strCAMARQERRO)
    
End Sub


Private Sub PintaColunasEditaveis()
    Dim I   As Long
    With grdOPAPOT
        For I = 1 To (.Rows - 1)
            .Cell(flexcpBackColor, I, conCOL_PRODOPAP_CODOP, I, conCOL_PRODOPAP_DTPROG) = &H80FF&
            .Cell(flexcpBackColor, I, conCOL_PRODOPAP_CODOP, I, conCOL_PRODOPAP_CODOP) = &H80FF&
            .Cell(flexcpBackColor, I, conCOL_PRODOPAP_QTDEPROD, I, conCOL_PRODOPAP_QTDEPROD) = &H80FF&
            .Cell(flexcpBackColor, I, conCOL_PRODOPAP_QTDFOLHAS, I, conCOL_PRODOPAP_QTDFOLHAS) = &H80FF&
            .Cell(flexcpBackColor, I, conCOL_PRODOPAP_PESO, I, conCOL_PRODOPAP_PESO) = &H80FF&
            .Cell(flexcpBackColor, I, conCOL_PRODOPAP_PERDAMONT, I, conCOL_PRODOPAP_PERDAMONT) = &H80FF&
            .Cell(flexcpBackColor, I, conCOL_PRODOPAP_PERDAMONT, I, conCOL_PRODOPAP_HORAINI) = &H80FF&
            .Cell(flexcpBackColor, I, conCOL_PRODOPAP_PERDAMONT, I, conCOL_PRODOPAP_HORAFIN) = &H80FF&
            
            .Cell(flexcpForeColor, I, conCOL_PRODOPAP_CODOP, I, conCOL_PRODOPAP_DTPROG) = vbWhite
            .Cell(flexcpForeColor, I, conCOL_PRODOPAP_CODOP, I, conCOL_PRODOPAP_CODOP) = vbWhite
            .Cell(flexcpForeColor, I, conCOL_PRODOPAP_QTDEPROD, I, conCOL_PRODOPAP_QTDEPROD) = vbWhite
            .Cell(flexcpForeColor, I, conCOL_PRODOPAP_QTDFOLHAS, I, conCOL_PRODOPAP_QTDFOLHAS) = vbWhite
            .Cell(flexcpForeColor, I, conCOL_PRODOPAP_PESO, I, conCOL_PRODOPAP_PESO) = vbWhite
            .Cell(flexcpForeColor, I, conCOL_PRODOPAP_PERDAMONT, I, conCOL_PRODOPAP_PERDAMONT) = vbWhite
            .Cell(flexcpForeColor, I, conCOL_PRODOPAP_PERDAMONT, I, conCOL_PRODOPAP_HORAINI) = vbWhite
            .Cell(flexcpForeColor, I, conCOL_PRODOPAP_PERDAMONT, I, conCOL_PRODOPAP_HORAFIN) = vbWhite
        Next I
    End With
End Sub

Private Sub MarcaLinha(grdGenerica As Variant, lngRowSel As Long)
    With grdGenerica
        If lngRowSel > 0 And (.Rows - 1) > 0 Then
            .Cell(flexcpBackColor, .RowSel, 0, .RowSel, (.Cols - 1)) = &H8000000D
            .Cell(flexcpForeColor, .RowSel, 0, .RowSel, (.Cols - 1)) = &H8000000E
        End If
    End With
End Sub

Private Sub Desmarca(grdGenerica As Variant, lngRowSel As Long)
    With grdGenerica
        If lngRowSel > 0 And (.Rows - 1) > 0 Then
            .Cell(flexcpBackColor, .RowSel, 0, .RowSel, (.Cols - 1)) = &H8000000E
            .Cell(flexcpForeColor, .RowSel, 0, .RowSel, (.Cols - 1)) = &H80000008
        End If
    End With
End Sub


Public Function PegaFechamentoTampaFuro(strCODFECH As String) As String

    If BREC10.State = 1 Then BREC10.Close
    
    PegaFechamentoTampaFuro = ""
    
    If Len(Trim(strCODFECH)) = 0 Then Exit Function
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       FECH.SGI_CODIGO" & vbCrLf
    sSql = sSql & "      ,FECH.SGI_DESCRI" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADFECHAM                 FECH" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    
    sSql = sSql & "       FECH.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And FECH.SGI_CODIGO = " & Trim(strCODFECH)
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then PegaFechamentoTampaFuro = Trim(BREC10!SGI_DESCRI)
    BREC10.Close
    
End Function

Private Function PegaComp(lngCOMP As Long) As String
    PegaComp = ""
    If lngCOMP = 1 Then PegaComp = "VEX"
    If lngCOMP = 2 Then PegaComp = "VZ"
    If lngCOMP = 3 Then PegaComp = "NAT"
    If lngCOMP = 4 Then PegaComp = "VI"
End Function

Private Function Calcula10porc(lngROW As Long, lngQTDPROG As Long) As Boolean
    
    Calcula10porc = True
    
    Dim lngVALPORC  As Long
    Dim lngQTDOP    As Long
    Dim lngVLTOTAL  As Long
    
    With grdOPAPOT
        
        lngQTDOP = CLng(.Cell(flexcpText, lngROW, conCOL_PRODOPAP_QTDEPED))
        lngVALPORC = (lngQTDOP * 0.1)
        lngVLTOTAL = (lngQTDOP + lngVALPORC)
        
        If lngQTDPROG > lngVLTOTAL Then Calcula10porc = False
    
    End With
End Function

Private Function PegaSaldoOP(lngCODINTENO As Long, lngQTDOP As Long) As Long

    PegaSaldoOP = 0
    
    Dim I           As Long
    Dim lngQTDAPONT As Long
    
    
    lngQTDAPONT = 0
    With grdOPAPOT
        For I = 1 To (.Rows - 1)
            If .Cell(flexcpText, I, conCOL_PRODOPAP_IDINTPROG) = lngCODINTENO And _
               .Cell(flexcpText, I, conCOL_PRODOPAP_Action2Do) <> dacEnumUpdateAction_delete Then
               If IsNumeric(.Cell(flexcpText, I, conCOL_PRODOPAP_QTDEPROD)) Then lngQTDAPONT = lngQTDAPONT + CLng(.Cell(flexcpText, I, conCOL_PRODOPAP_QTDEPROD))
            End If
        Next I
    End With

    PegaSaldoOP = (lngQTDOP - lngQTDAPONT)

End Function

Private Function PegaQTD_JAAPONT(lngPROGINTENO As Long, lngQTDAPONT As Long, lngIDINTOP As Long) As Long

    PegaQTD_JAAPONT = 0
    
    Dim I                   As Long
    Dim lngQTDJA_APONT      As Long
    Dim lngQTDJA_APONT_01   As Long
    Dim lngQTDPROD          As Long
    Dim dtDTBAIXAPROG       As Date
    
    dtDTBAIXAPROG = CDate(mskDtDoc.Text)
    
    lngQTDJA_APONT = 0
    With grdOPAPOT
        For I = 1 To (.Rows - 1)
            If .Cell(flexcpText, I, conCOL_PRODOPAP_IDINTPROG) = lngPROGINTENO And _
               .Cell(flexcpText, I, conCOL_PRODOPAP_Action2Do) <> dacEnumUpdateAction_delete Then
               lngQTDPROD = 0
               If IsNumeric(.Cell(flexcpText, I, conCOL_PRODOPAP_QTDEPROD)) Then lngQTDPROD = CLng(.Cell(flexcpText, I, conCOL_PRODOPAP_QTDEPROD))
               lngQTDJA_APONT = (lngQTDJA_APONT + lngQTDPROD)
            End If
        Next I
    End With

    
    PegaQTD_JAAPONT = (lngQTDJA_APONT + lngQTDAPONT)


End Function

Private Sub MudandoStatusOPs(lngSALDOOP As Long, lngCODINTENO As Long)
    
    Dim lngStatusOP     As Long
    Dim I               As Long
    
    If lngSALDOOP > 0 Then lngStatusOP = 2
    If lngSALDOOP <= 0 Then lngStatusOP = 1
    
    With grdOPAPOT
        For I = 1 To (.Rows - 1)
            If .Cell(flexcpText, I, conCOL_PRODOPAP_IDINTPROG) = lngCODINTENO And _
               .Cell(flexcpText, I, conCOL_PRODOPAP_Action2Do) <> dacEnumUpdateAction_delete Then
                .Cell(flexcpText, I, conCOL_PRODOPAP_StatusApont) = lngStatusOP
                If .Cell(flexcpText, I, conCOL_PRODOPAP_Action2Do) = dacEnumUpdateAction_Ignore Then
                   .Cell(flexcpText, I, conCOL_PRODOPAP_Action2Do) = dacEnumUpdateAction_update
                End If
            End If
        Next I
    End With
    
End Sub

Private Sub FechaProg()

    Dim I                       As Long
    Dim lngCODINTPROG           As Long
    Dim dtDTBAIXAPROG           As Date
    Dim lngJAPROG_OUTRASDATAS   As Long
    Dim lngQTDE_JAAPONTADA      As Long
    Dim lngQTDOP                As Long
    Dim lngTOTAL_JAAPONT        As Long
    Dim lngSALDOPROG            As Long
    
    
    dtDTBAIXAPROG = CDate(mskDtDoc.Text)
    
    With grdOPAPOT
        For I = 1 To (.Rows - 1)
            
            If .Cell(flexcpText, I, conCOL_PRODOPAP_Action2Do) <> dacEnumUpdateAction_delete And _
               .Cell(flexcpText, I, conCOL_PRODOPAP_Action2Do) <> dacEnumUpdateAction_Ignore Then
               
                lngCODINTPROG = .Cell(flexcpText, I, conCOL_PRODOPAP_IDINTPROG)
                lngQTDOP = PegaQTDEOP(.Cell(flexcpText, I, conCOL_PRODOPAP_IDINTOP))
                
                lngJAPROG_OUTRASDATAS = PegaJaProg_OutrasDatas(.Cell(flexcpText, I, conCOL_PRODOPAP_IDINTOP), .Cell(flexcpText, I, conCOL_PRODOPAP_IDINTPROG), dtDTBAIXAPROG)
                lngQTDE_JAAPONTADA = PegaQTD_JAAPONT(.Cell(flexcpText, I, conCOL_PRODOPAP_IDINTPROG), 0, .Cell(flexcpText, I, conCOL_PRODOPAP_IDINTOP))
                lngTOTAL_JAAPONT = (lngJAPROG_OUTRASDATAS + lngQTDE_JAAPONTADA)
                
                lngSALDOPROG = (lngQTDOP - lngTOTAL_JAAPONT)
                If lngSALDOPROG <= 0 Then
                
                End If
                
            End If
            
        Next I
    End With
End Sub

Private Function PegaJaProg_OutrasDatas(lngIDINTOP As Long, lngIDINTPROG As Long, dtBAIXAPROG As Date) As Long
    
    PegaJaProg_OutrasDatas = 0
    
    Dim lngJAAOPNT_01 As Long
    Dim lngJAAOPNT_02 As Long
    
    
    '' ===========================================
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       Sum(SGI_QTDEPROD) As SGI_QTDEPROD" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADAPONTPROG" & strNOMTABELA & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_FILIAL    =  " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_IDINTOP   =  " & lngIDINTOP & vbCrLf
    sSql = sSql & "   And SGI_IDINTPROG = " & lngIDINTPROG & vbCrLf
    sSql = sSql & "   And SGI_DATAPONT  <> '" & Format(dtBAIXAPROG, "MM/DD/YYYY") & "'"
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then
        If Not IsNull(BREC10!SGI_QTDEPROD) Then lngJAAOPNT_02 = BREC10!SGI_QTDEPROD
    End If
    BREC10.Close
    
    PegaJaProg_OutrasDatas = (lngJAAOPNT_01 + lngJAAOPNT_02)
    
End Function

Private Function PegaQTDEOP(lngIDOP As Long)

    PegaQTDEOP = 0


    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD" & strNOMTABELA & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_IDPAI  = " & lngIDOP
    
    BREC11.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC11.EOF() Then PegaQTDEOP = BREC11!SGI_QTDE
    BREC11.Close
    
End Function


Private Function PegaJaProg_02(lngIDINTOP As Long, lngIDINTPROG As Long, dtBAIXAPROG As Date) As Long
    
    PegaJaProg_02 = 0
    
    Dim lngJAAOPNT_02 As Long
    
    '' ===========================================
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       Sum(SGI_QTDEPROD) As SGI_QTDEPROD" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADAPONTPROG" & strNOMTABELA & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_FILIAL    =  " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_IDINTOP   =  " & lngIDINTOP & vbCrLf
    sSql = sSql & "   And SGI_IDINTPROG = " & lngIDINTPROG & vbCrLf
    sSql = sSql & "   And SGI_DATAPONT  <> '" & Format(dtBAIXAPROG, "MM/DD/YYYY") & "'"
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC10.EOF() Then
        If Not IsNull(BREC10!SGI_QTDEPROD) Then lngJAAOPNT_02 = BREC10!SGI_QTDEPROD
    End If
    BREC10.Close
    
    PegaJaProg_02 = lngJAAOPNT_02
    
End Function

