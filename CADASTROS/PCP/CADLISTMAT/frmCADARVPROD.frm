VERSION 5.00
Object = "{69ECBBD3-5C2A-4A84-ABEC-23937DBF1B54}#1.4#0"; "FlowChartPro.dll"
Begin VB.Form frmCADARVPROD 
   BackColor       =   &H8000000A&
   Caption         =   "Cadastro de Estrutura de Produtos"
   ClientHeight    =   9075
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   13335
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   13335
   StartUpPosition =   1  'CenterOwner
   Begin FLOWCHARTLibCtl.FlowChart fc 
      Height          =   7935
      Left            =   0
      TabIndex        =   18
      Top             =   960
      Width           =   13335
      _cx             =   23521
      _cy             =   13996
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BoxFillColor    =   16777215
      BoxFrameColor   =   0
      ArrowColor      =   16711680
      AlignToGrid     =   -1  'True
      ShowGrid        =   0   'False
      GridColor       =   9866380
      GridSize        =   16
      BoxStyle        =   2
      BackColor       =   13150890
      ShowShadows     =   -1  'True
      ShadowColor     =   9203310
      TextStyle       =   37
      PicturePos      =   0
      TextColor       =   0
      ArrowStyle      =   1
      ArrowSegments   =   1
      StaticMode      =   0   'False
      ScrollX         =   0
      ScrollY         =   0
      ZoomFactor      =   100
      ShowToolTips    =   0   'False
      ToolTipStyle    =   0
      AllowRefLinks   =   -1  'True
      PenStyle        =   0
      PenWidth        =   1
      Behavior        =   0
      ArrowHead       =   4
      SelectAfterCreate=   -1  'True
      BoxCustomDraw   =   0
      ShadowOffsetX   =   3
      ShadowOffsetY   =   3
      RestrObjsToDoc  =   1
      DynamicArrows   =   0   'False
      AutoScroll      =   -1  'True
      TableFillColor  =   10526900
      TableFrameColor =   0
      TableRowsCount  =   4
      TableColumnsCount=   2
      MultiSelStyle   =   0
      ModificationStart=   0
      TableColWidth   =   50
      TableRowHeight  =   22
      TableCaptionHeight=   22
      TableCaption    =   "Table"
      PrpArrowStartOrnt=   0
      TableCellBorders=   2
      UndoDepth       =   0
      SelectAfterPaste=   -1  'True
      DragDropMode    =   3
      KbdActive       =   -1  'True
      ActiveMnpColor  =   16777215
      SelMnpColor     =   11184810
      DisabledMnpColor=   200
      BoxSelStyle     =   1
      TableSelStyle   =   2
      ArrowHeadSize   =   14
      ShowDisabledHandles=   -1  'True
      ExpandOnIncoming=   0   'False
      BoxesExpandable =   -1  'True
      TablesScrollable=   0   'False
      RecursiveExpand =   -1  'True
      ShadowsStyle    =   0
      ArrowEndsMovable=   0   'False
      BoxIncmAnchor   =   0
      BoxOutgAnchor   =   0
      ArrowBase       =   0
      ArrowBaseSize   =   20
      IntermArrowHead =   0
      IntermHeadSize  =   12
      ExtrnDragDrop   =   0   'False
      SnapStyle       =   0
      SnapDistance    =   20
      BoxFillStyle    =   0
      BoxFillColor2   =   16768220
      ArrowFillColor  =   12632064
      AutoSizeDoc     =   0   'False
      FeedbackColor   =   25800
      FeedbackPenStyle=   0
      FeedbackPenWidth=   3
      LayoutGap       =   14
      LayoutStyle     =   0
      FeedbackOnDragOver=   -1  'True
      InplaceEditAllowed=   0   'False
      ShowFocusFrame  =   0   'False
      ExpandBtnPos    =   0
      ArrowsSplittable=   0   'False
      KbdBehavior     =   0
      FireMouseMove   =   0   'False
      AxControlId     =   ""
      AllowMultiSel   =   -1  'True
      AllowLinksRepeat=   -1  'True
      GridStyle       =   0
      ShapeOrientation=   0
      ShowAnchors     =   2
      SnapToAnchor    =   0
      IconTextWidth   =   100
      ArrowCrossings  =   0
      CrossRadius     =   8
      TableLinkStyle  =   0
      RouteArrows     =   0   'False
      IconTextHeight  =   50
      HighSpeedRouting=   0   'False
      HostedAxActivation=   0
      EnableStyledText=   0   'False
      AllowUnconnectedArrows=   0   'False
      ScrollRate      =   1
      BoxFillColorAlpha=   255
      ArrowFillColorAlpha=   255
      TableFillColorAlpha=   255
      ShadowColorAlpha=   150
      TableStyle      =   0
      RerouteArrows   =   1
      RoutingGridSize =   16
      MinimizeRouteSegments=   0   'False
      ArrowText       =   ""
      BoxText         =   ""
      ArrowsSnapToNodeBorders=   0   'False
      ArrowTextStyle  =   0
      ScrollZoneSize  =   0
      HitTestPriority =   1
      ArrowSelStyle   =   1
      BoxWindowFrame  =   2
      AllowUnanchoredArrows=   -1  'True
      MeasureUnit     =   2
      SelHandleSize   =   9
      SelectionOnTop  =   -1  'True
      ToolTipDelay    =   500
      BoxPicturePos   =   2
      MergeThreshold  =   0
      ControlPadding  =   4
      MiddleButtonAction=   0
      RoundedArrows   =   0   'False
      RoundedArrowsRadius=   40
      DisableNoScroll =   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "A&justa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12480
      Picture         =   "frmCADARVPROD.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   120
      Width           =   735
   End
   Begin VB.Frame fraZoom 
      Caption         =   "[ Zoom ]"
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
      Height          =   855
      Left            =   3480
      TabIndex        =   12
      Top             =   0
      Width           =   1815
      Begin VB.OptionButton optZoon 
         Caption         =   "75 %"
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
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton optZoon 
         Caption         =   "25 %"
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
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optZoon 
         Caption         =   "100 %"
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
         Index           =   1
         Left            =   840
         TabIndex        =   14
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton optZoon 
         Caption         =   "50 %"
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
         Index           =   0
         Left            =   840
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   3495
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
         Left            =   2640
         Picture         =   "frmCADARVPROD.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Exclui Empresa"
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
         Picture         =   "frmCADARVPROD.frx":040C
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Picture         =   "frmCADARVPROD.frx":050E
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
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
         Picture         =   "frmCADARVPROD.frx":0610
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame fraTipos 
      Caption         =   "[ Tipos de Produtos ]"
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
      Height          =   855
      Left            =   8040
      TabIndex        =   3
      Top             =   0
      Width           =   4335
      Begin VB.PictureBox pb2 
         AutoSize        =   -1  'True
         Height          =   540
         Left            =   2280
         OLEDragMode     =   1  'Automatic
         Picture         =   "frmCADARVPROD.frx":0B42
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   5
         Top             =   240
         Width           =   540
      End
      Begin VB.PictureBox pb1 
         AutoSize        =   -1  'True
         Height          =   540
         Left            =   120
         Picture         =   "frmCADARVPROD.frx":1B84
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   4
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Prod.Comprado"
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
         Left            =   2880
         TabIndex        =   7
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Prod.Processado"
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
         Left            =   720
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame fraLinha 
      Caption         =   "[ Alinhamento ]"
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
      Height          =   855
      Left            =   5400
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin VB.OptionButton optAlinh 
         Caption         =   "De Cima Para Baixo"
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
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2175
      End
      Begin VB.OptionButton optAlinh 
         Caption         =   "Da Esq. para a direita"
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
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmCADARVPROD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cCaminho         As String
Public Linha            As Variant
Public cTipOper         As String
Public iCodigo          As Long
Public strCODPROD       As String
Public FILIAL           As Integer
Public strAcesso        As String
Public strMODPAI        As String

Dim objBLBFunc          As Object
Dim objCADARVPROD       As Object
Dim lngTotReg           As Long
Dim strITENSDELETED     As String

Dim currNode            As box
Dim root                As box
Dim company             As Group
Dim uniqueID            As Long
Dim lngINDICE           As Long
''Dim Overview            As frmVIEWARVPROD

Private Sub layoutTree(dir As ETreeLayoutDirection)
    Dim tl As TreeLayout
    Set tl = New TreeLayout
    
    tl.root = root
    tl.Type = tltCentralized
    tl.Direction = dir
    tl.ArrowStyle = tlaStraight
    tl.NodeSpacing = 15
    tl.LevelSpacing = 30
    tl.KeepRootPos = True
    tl.reversedArrows = False
    tl.KeepGroupLayout = False
    
    fc.ArrangeDiagram tl
End Sub


Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    
    fraLinha.Enabled = True
    fraZoom.Enabled = True
    
    fraTipos.Enabled = True
    
    Me.Caption = "Cadastro de arvore de produtos - [ ALTERAÇÃO ]"
    
    cTipOper = "A"
    
End Sub

Private Sub cmdImpressao_Click()
    fc.PreviewDiagram
End Sub

Private Sub CmdSalva_Click()

    Dim arrDELETADOS As Variant
    
    If ValidaCampos = False Then Exit Sub
    
    objCADARVPROD.CODIGO = iCodigo
        
    objCADARVPROD.DELETADOS = Empty
    If Len(Trim(strITENSDELETED)) > 0 Then
       arrDELETADOS = Split(strITENSDELETED, "|")
       objCADARVPROD.DELETADOS = arrDELETADOS
    End If
    
    If cTipOper = "A" Then
       If objCADARVPROD.GRAVA("E") = False Then Exit Sub
    End If
    
    '' Criando Os Códigos
    lngIDPai = arrPROVARV(0).lngIDPai
    lngIDNOVO = arrPROVARV(0).lngCODPAI
    For i = 0 To UBound(arrPROVARV)
        If Len(Trim(arrPROVARV(i).strPRODUTO)) > 0 Then
           lngCODIGO = objCADARVPROD.Gera_Codigo("frmCADARVPROD")
           If lngIDPai <> arrPROVARV(i).lngIDPai Then
              lngIDPai = arrPROVARV(i).lngIDPai
              lngIDNOVO = arrPROVARV(arrPROVARV(i).lngIDPai).lngCODIGO
           End If
           arrPROVARV(i).lngCODPAI = lngIDNOVO
           arrPROVARV(i).lngCODIGO = lngCODIGO
        End If
    Next i
    
    If objCADARVPROD.GRAVA(cTipOper) = False Then Exit Sub
    
    MsgBox "A Estrtura foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
       
    strITENSDELETED = ""
       
    If cTipOper = "I" Then
       Set objBLBFunc = Nothing
       Set objCADARVPROD = Nothing
       Unload Me
    End If

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objCADARVPROD = Nothing
    Unload Me
End Sub

Private Sub Command1_Click()
    ''Set Overview = New frmVIEWARVPROD
    ''Overview.SetDocument fc
    ''Overview.SetParent Me
    
    ''Overview.Show vbModal, Me
End Sub




Private Sub fc_BoxCollapsed(ByVal box As FLOWCHARTLibCtl.IBoxItem)
    If optAlinh(0).Value = True Then layoutTree tldLeftToRight
    If optAlinh(1).Value = True Then layoutTree tldTopToBottom
End Sub

Private Sub fc_BoxDblClicked(ByVal box As FLOWCHARTLibCtl.IBoxItem, ByVal Button As FLOWCHARTLibCtl.EMouseButton, ByVal x As Long, ByVal y As Long)
    
    frmCADPRODLISTMAT.cTipOper = cTipOper
    frmCADPRODLISTMAT.FILIAL = FILIAL
    frmCADPRODLISTMAT.cCaminho = cCaminho
    frmCADPRODLISTMAT.Linha = Linha
    frmCADPRODLISTMAT.strAcesso = strAcesso
    frmCADPRODLISTMAT.lngINDICE = box.Tag
    
    frmCADPRODLISTMAT.Show vbModal
    
    If Trim(box.Text) <> Trim(arrPROVARV(box.Tag).strPRODUTO) Then
       
       box.Text = arrPROVARV(box.Tag).strPRODUTO
       If arrPROVARV(box.Tag).intAction2Do = dacEnumUpdateAction_update Then
          Call BoxAlterado(box)
          Call RefazIndice
       End If
       
       '' Continua o arvore se caso tiver ramificação
       fc.ClearAll
       
       ''Call CarregaListaAlterada(Trim(box.Text), CLng(box.Tag))
       Call CarregaDoArray
       
    End If
    
End Sub


Private Sub fc_BoxExpanded(ByVal box As FLOWCHARTLibCtl.IBoxItem)
    If optAlinh(0).Value = True Then layoutTree tldLeftToRight
    If optAlinh(1).Value = True Then layoutTree tldTopToBottom
End Sub

Private Sub fc_DragOverBoxVB(ByVal box As FLOWCHARTLibCtl.IBoxItem, ByVal dataObj As FLOWCHARTLibCtl.IVBDataObject, ByVal docX As Long, ByVal docY As Long, ByVal keyState As Long, effect As Long)
    If box.Visible Then
        effect = vbDropEffectCopy
    Else
        effect = vbDropEffectNone
    End If
End Sub

Private Sub fc_DragOverDocVB(ByVal dataObj As FLOWCHARTLibCtl.IVBDataObject, ByVal docX As Long, ByVal docY As Long, ByVal keyState As Long, effect As Long)
    effect = vbDropEffectNone
End Sub

Private Sub fc_DropInBoxVB(ByVal box As FLOWCHARTLibCtl.IBoxItem, ByVal dataObj As FLOWCHARTLibCtl.IVBDataObject, ByVal docX As Long, ByVal docY As Long, ByVal keyState As Long, effect As Long)
   
    If ConsisteProd(box) = False Then Exit Sub
    
    Dim newNode     As FLOWCHARTLibCtl.box
    Set newNode = addChild(box)
    If newNode Is Nothing Then Exit Sub
    
    'mostra o tag do nó
    newNode.Text = ""
    
    'Se a picture for dragged
    If dataObj.GetFormat(vbCFDIB) Then
        'fixe como um ícone de nó
        newNode.PicturePos = picCenterLeft
        newNode.Picture = dataObj.GetData(vbCFDIB)
        
        'e move o texto à direita
        newNode.TextStyle = tsRight
    End If
End Sub

Private Sub fc_DropInDocVB(ByVal dataObj As FLOWCHARTLibCtl.IVBDataObject, ByVal docX As Long, ByVal docY As Long, ByVal keyState As Long, effect As Long)
    If Not fc.ActiveBox Is Nothing Then
        addChild fc.ActiveBox
    End If
End Sub

Private Sub fc_KeyDown(ByVal KeyCode As Long, ByVal Shift As Long)

    If KeyCode = vbKeyReturn Then
    
        If fc.SelectedBoxes.Count > 0 Then
        
            frmCADPRODLISTMAT.cTipOper = cTipOper
            frmCADPRODLISTMAT.FILIAL = FILIAL
            frmCADPRODLISTMAT.cCaminho = cCaminho
            frmCADPRODLISTMAT.Linha = Linha
            frmCADPRODLISTMAT.strAcesso = strAcesso
            frmCADPRODLISTMAT.lngINDICE = fc.ActiveBox.Tag
            
            frmCADPRODLISTMAT.Show vbModal
            
            If Trim(fc.ActiveBox.Text) <> Trim(arrPROVARV(fc.ActiveBox.Tag).strPRODUTO) Then
               
               fc.ActiveBox.Text = arrPROVARV(fc.ActiveBox.Tag).strPRODUTO
               If arrPROVARV(fc.ActiveBox.Tag).intAction2Do = dacEnumUpdateAction_update Then
                  Call BoxAlterado(fc.ActiveBox) '' box
                  Call RefazIndice
               End If
               
               '' Continua o arvore se caso tiver ramificação
               fc.ClearAll
               
               ''Call CarregaListaAlterada(Trim(box.Text), CLng(box.Tag))
               Call CarregaDoArray
               
            End If
        
        End If
        
    End If

End Sub

Private Sub fc_RequestDeleteArrow(ByVal arrow As FLOWCHARTLibCtl.IArrowItem, pbDelete As Boolean)
    'não permita o usuário apagar as setas
    pbDelete = False
End Sub

Private Sub fc_RequestDeleteBox(ByVal box As FLOWCHARTLibCtl.IBoxItem, pbDelete As Boolean)
    If cTipOper = "C" Then
       pbDelete = False
       Exit Sub
    End If
    If box.Tag = 0 Then pbDelete = False
    If box.Tag > 0 Then
       If cTipOper = "I" Or cTipOper = "A" Then
          pbDelete = TemSetasSaida(box)
          Call RefazIndice
       End If
    End If
End Sub

Private Sub fc_RequestSelectArrow(ByVal arrow As FLOWCHARTLibCtl.IArrowItem, pbSelect As Boolean)
    'não permita os usuários selecionar as setas
    pbSelect = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()
    
   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objCADARVPROD = CreateObject("CADLISTMAT.clsCADLISTMAT")
   Set colPRODUTOS = New Collection
      
   objCADARVPROD.FILIAL = FILIAL
   
   Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

   If adoBanco_Dados.State = 0 Then
      MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
      Exit Sub
   End If
   
   strITENSDELETED = ""
   
   fc.Graphics.StartUp geGdiPlus
    
   'abilitando drag'n'drop
   fc.RegisterDragDrop
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta
   If cTipOper = "AL" Then Call CarregaListaAlterando(strCODPROD)
    
   'tira a selecão da raiz
   fc.ClearSelection
   fc.ExpandOnIncoming = False
    
   optAlinh(0).Value = True
   optZoon(3).Value = True
   
   ''fc.ZoomFactor = 70
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'limpa
    fc.RevokeDragDrop
    fc.Graphics.ShutDown
End Sub


Private Sub optAlinh_Click(Index As Integer)
    If Index = 0 Then layoutTree tldLeftToRight
    If Index = 1 Then layoutTree tldTopToBottom
End Sub


Private Sub optZoon_Click(Index As Integer)
    If Index = 0 Then fc.ZoomFactor = 50
    If Index = 1 Then fc.ZoomFactor = 100
    If Index = 2 Then fc.ZoomFactor = 25
    If Index = 3 Then fc.ZoomFactor = 75
End Sub

Private Sub pb1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    pb1.OLEDrag
End Sub

Private Sub pb1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    AllowedEffects = vbDropEffectCopy
    Data.SetData pb1.Picture, vbCFDIB
End Sub

Private Function addChild(node As FLOWCHARTLibCtl.box) As box
    
    Dim child           As box
    Dim link            As arrow
    Dim division        As Group
    Dim parentDiv       As Group
    
    'adquire o grupo do nó do pai
    Set parentDiv = node.SubordinateGroup
    
    'cria o nó novo e acrescenta ao grupo do pai
    'assim o nó moverá se seu pai é movido
    Set child = fc.CreateBox(0, 0, 150, 40)
    child.Transparent = False
    
    parentDiv.AttachToCorner child, 0

    'crie uma seta entre o nó de pai e o Filho
    Set link = fc.CreateArrow(node, child)
    
    'atualiza o contador
    uniqueID = uniqueID + 1
    child.Tag = uniqueID
    
    ReDim Preserve arrPROVARV(0 To uniqueID) As PRODARVPROD
    arrPROVARV(uniqueID).lngID = uniqueID
    arrPROVARV(uniqueID).lngIDPai = node.Tag
    arrPROVARV(uniqueID).lngTipo = 0
    arrPROVARV(uniqueID).strPRODUTO = ""
    
    arrPROVARV(uniqueID).curQTDCONS = 0
    arrPROVARV(uniqueID).strProdutoPAI = node.Text
    arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Insert
    
    arrPROVARV(uniqueID).lngCODPAI = arrPROVARV(node.Tag).lngCODIGO
    arrPROVARV(uniqueID).lngCODIGO = (uniqueID * -1)
    arrPROVARV(uniqueID).lngProdutoIDPai = arrPROVARV(node.Tag).lngProdutoID
    
    'começa um grupo novo ao qual serão acrescentados os filhos de nó
    Set division = fc.CreateGroup(child)
    
    'arruma a árvore inteira
    If optAlinh(0).Value = True Then layoutTree tldLeftToRight
    If optAlinh(1).Value = True Then layoutTree tldTopToBottom
            
    'selecione o novo nó
    fc.ClearSelection
    fc.AddToSelection child
    
    Set addChild = child
End Function

Private Function TemSetasSaida(ByVal box As FLOWCHARTLibCtl.IBoxItem) As Boolean
    
    Dim i                   As Integer
    Dim intDestino          As Integer
    
    Dim boxHorig            As FLOWCHARTLibCtl.IBoxItem
    Dim boxAnt              As FLOWCHARTLibCtl.IBoxItem
    
    Set boxHorig = box
    
    Dim SetaDeSaida         As FLOWCHARTLibCtl.IArrows
    Set SetaDeSaida = box.OutgoingArrows
    
    If SetaDeSaida.Count <= 0 Then
       TemSetasSaida = True
       arrPROVARV(box.Tag).intAction2Do = dacEnumUpdateAction_delete
    End If
    
    If SetaDeSaida.Count > 0 Then
       TemSetasSaida = True
    
VOLTA:
       For i = 0 To (SetaDeSaida.Count - 1)
           Set box = SetaDeSaida.Item(i).DestinationBox
           Set SetaDeSaida = box.OutgoingArrows
           If SetaDeSaida.Count = 0 Then
              arrPROVARV(box.Tag).intAction2Do = dacEnumUpdateAction_delete
              strITENSDELETED = strITENSDELETED & arrPROVARV(box.Tag).lngCODIGO & "|"
              fc.DeleteItem box
              Set box = boxHorig
              Set SetaDeSaida = box.OutgoingArrows
           End If
           GoTo VOLTA
       Next i
       arrPROVARV(box.Tag).intAction2Do = dacEnumUpdateAction_delete
       strITENSDELETED = strITENSDELETED & arrPROVARV(box.Tag).lngCODIGO & "|"
    End If
    
End Function

Private Function TemSetasEntrada(ByVal box As FLOWCHARTLibCtl.IBoxItem) As Boolean
    TemSetasEntrada = True
    Dim SetaDeEntrada As FLOWCHARTLibCtl.IArrows
    Set SetaDeEntrada = box.IncomingArrows
    If SetaDeEntrada.Count = 0 Then TemSetasEntrada = False
End Function


Private Function PegaIdPai(ByVal box As FLOWCHARTLibCtl.IBoxItem) As Long

    PegaIdPai = -1
    
    Dim i As Integer
    Dim SetaDeEntrada As FLOWCHARTLibCtl.IArrows
    Set SetaDeEntrada = box.IncomingArrows
    If SetaDeEntrada.Count = 0 Then Exit Function
    
    For i = 0 To SetaDeEntrada.Count - 1
        PegaIdPai = SetaDeEntrada.Item(i).OriginBox.Tag
    Next i
End Function

Private Function PegaIdFilhos(ByVal box As FLOWCHARTLibCtl.IBoxItem) As String

    PegaIdFilhos = ""
    
    Dim i As Integer
    Dim SetaDeSaida As FLOWCHARTLibCtl.IArrows
    Set SetaDeSaida = box.OutgoingArrows
    
    For i = 0 To SetaDeSaida.Count - 1
        PegaIdFilhos = PegaIdFilhos & SetaDeSaida.Item(i).DestinationBox.Tag & vbCrLf
    Next i
End Function

Private Sub Novo()
    
    If Len(Trim(strCODPROD)) = 0 Then Exit Sub
    
    
    'criando a raiz da árvore
    Set root = fc.CreateBox(100, 300, 150, 40)
    root.Picture = pb1.Picture
    root.Transparent = False
    root.Text = ""
    
        
    root.PicturePos = picCenterLeft
    root.Picture = pb1.Picture
    
    '' ========================================
    '' Criando a matrix para começar a arvore
    uniqueID = 0
    root.Tag = uniqueID
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_IDPRODUTO = " & Trim(strCODPROD) & vbCrLf
    sSql = sSql & "   And SGI_FILIAL = " & FILIAL
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
       
       ReDim Preserve arrPROVARV(0 To lngINDICE) As PRODARVPROD
       arrPROVARV(lngINDICE).lngID = uniqueID
       arrPROVARV(lngINDICE).lngIDPai = -1
       
       arrPROVARV(lngINDICE).lngProdutoIDPai = -1
       arrPROVARV(lngINDICE).lngProdutoID = BREC!SGI_IDPRODUTO
       
       arrPROVARV(lngINDICE).lngTipo = BREC!SGI_PRODUTOTIPO
       arrPROVARV(lngINDICE).lngCodUniMed = BREC!SGI_UNIDMEDIDA
       
       If BREC!SGI_PRODUTOTIPO = 1 And BREC!SGI_PRODUTOESTILO = 0 Then
          arrPROVARV(lngINDICE).strPRODUTO = Format(IIf(IsNull(BREC!SGI_CODLINPROD), 0, BREC!SGI_CODLINPROD), "###000") & "." & _
                                             Format(IIf(IsNull(BREC!SGI_CodClie), 0, BREC!SGI_CodClie), "####0000") & "." & _
                                             Format(IIf(IsNull(BREC!SGI_CODROTULO), 0, BREC!SGI_CODROTULO), "##00") & "." & _
                                             Format(IIf(IsNull(BREC!SGI_DIGVERIF), 0, BREC!SGI_DIGVERIF), "#0")

       ElseIf BREC!SGI_PRODUTOTIPO = 0 And BREC!SGI_PRODUTOESTILO = 0 Then
          arrPROVARV(lngINDICE).strPRODUTO = Trim(BREC!SGI_CODIGO)
       End If
       
       arrPROVARV(lngINDICE).curQTDCONS = 0
       arrPROVARV(lngINDICE).intAction2Do = dacEnumUpdateAction_Insert
       arrPROVARV(lngINDICE).strProdutoPAI = ""
       
       arrPROVARV(lngINDICE).lngCODPAI = -1
       arrPROVARV(lngINDICE).lngCODIGO = (uniqueID * -1)
       '' ========================================
       
       If BREC!SGI_PRODUTOTIPO = 1 And BREC!SGI_PRODUTOESTILO = 0 Then
          root.Text = Format(IIf(IsNull(BREC!SGI_CODLINPROD), 0, BREC!SGI_CODLINPROD), "###000") & "." & _
                      Format(IIf(IsNull(BREC!SGI_CodClie), 0, BREC!SGI_CodClie), "####0000") & "." & _
                      Format(IIf(IsNull(BREC!SGI_CODROTULO), 0, BREC!SGI_CODROTULO), "##00") & "." & _
                      Format(IIf(IsNull(BREC!SGI_DIGVERIF), 0, BREC!SGI_DIGVERIF), "#0")
                      
       ElseIf BREC!SGI_PRODUTOTIPO = 0 And BREC!SGI_PRODUTOESTILO = 0 Then
          root.Text = Trim(BREC!SGI_CODIGO)
       End If
       
    End If
    BREC.Close
    
    'mostra o tag do nó
    root.TextStyle = tsRight
    
    'cria um grupo hierárquico, de forma que quando
    'a raiz é movida todos os Filhos também moverão
    Set company = fc.CreateGroup(root)

End Sub

Private Function addChildFromArray(node As FLOWCHARTLibCtl.box, lngID As Long) As box
    
    Dim child       As box
    Dim link        As arrow
    Dim division    As Group
    Dim parentDiv   As Group
    
    'adquire o grupo do nó do pai
    Set parentDiv = node.SubordinateGroup
    
    'cria o nó novo e acrescenta ao grupo do pai
    'assim o nó moverá se seu pai é movido
    Set child = fc.CreateBox(0, 0, 150, 40)
    child.Transparent = False
    child.Tag = lngID
    
    parentDiv.AttachToCorner child, 0
        
    'crie uma seta entre o nó de pai e o Filho
    Set link = fc.CreateArrow(node, child)
    
    'começa um grupo novo ao qual serão acrescentados os filhos de nó
    Set division = fc.CreateGroup(child)
    
    'selecione o novo nó
    fc.ClearSelection
    fc.AddToSelection child
        
    Set addChildFromArray = child
End Function

Private Sub Consulta()

    Dim i As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    
    fraLinha.Enabled = True
    fraZoom.Enabled = True
    
    fraTipos.Enabled = False
  
    Me.Caption = "Cadastro de arvore de produtos - [ CONSULTA ]"
    
    fc.ClearAll
    
    Call CarregaLista(iCodigo, FILIAL)
    Call CarregaDoArray
    
End Sub

Public Sub CarregaLista(Optional lngPRODUTO As Long, Optional lngFilial As Integer)
    
    Dim lngIDPai     As Long
    Dim lngNivel0    As Long
    Dim lngNivel1    As Long
    Dim lngNivel2    As Long
    Dim lngNivel3    As Long
    Dim lngNivel4    As Long
    Dim lngNivel5    As Long
    Dim lngNivel6    As Long
    Dim lngNivel7    As Long
    Dim lngNivel8    As Long
    Dim lngNivel9    As Long
    Dim lngNivel10   As Long
    
    sSql = "Select " & vbCrLf
    
    sSql = sSql & "Case PRDOUTO.SGI_PRODUTOTIPO" & vbCrLf
    sSql = sSql & "        When 1 Then replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODLINPROD))) + '.' +"
    sSql = sSql & "                    replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODCLIE))) + '.' +"
    sSql = sSql & "                    replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODROTULO))) + '.' +"
    sSql = sSql & "                    (Case"
    sSql = sSql & "                          When PRDOUTO.SGI_DIGVERIF Is Null Then '0'"
    sSql = sSql & "                          When PRDOUTO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(1),PRDOUTO.SGI_DIGVERIF))) End)"
    sSql = sSql & "        When 0 Then PRDOUTO.SGI_CODIGO End As SGI_CODIGO"

    sSql = sSql & "      ,PRDOUTO.SGI_DESCRICAO   " & vbCrLf
    sSql = sSql & "      ,PRDOUTO.SGI_PRCCUSTO    " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_PRODLST     " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_QTDE        " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_UNIDCONS    " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_CUSTOUNIT   " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_CUSTOTOTAL  " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_PRODUTO     " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_CODIGO      AS SGI_CODID " & vbCrLf
    
    sSql = sSql & "      ,LISTMAT.SGI_IDPRODUTO   " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_IDPRODLST   " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_CODUNIMED   " & vbCrLf
    
    sSql = sSql & "      ,PRDOUTO.SGI_PRODUTOTIPO " & vbCrLf
    sSql = sSql & "      ,PRDOUTO.SGI_PRECOPROD   " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_LISTAMATPROD   LISTMAT " & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO     PRDOUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       LISTMAT.SGI_FILIAL  = " & lngFilial & vbCrLf
    ''sSql = sSql & "   And LISTMAT.SGI_PRODLST = '" & Trim(strPRODUTO) & "'" & vbCrLf
    
    sSql = sSql & "   And LISTMAT.SGI_CODIGO     = " & lngPRODUTO & vbCrLf
    sSql = sSql & "   And PRDOUTO.SGI_FILIAL     = LISTMAT.SGI_FILIAL       " & vbCrLf
    sSql = sSql & "   And PRDOUTO.SGI_IDPRODUTO  = LISTMAT.SGI_IDPRODUTO    " & vbCrLf
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC.EOF() Then
    
       '' ========================================
       '' Criando a matrix para começar a arvore
       uniqueID = 0
       lngIDPai = -1
       
       '' ========================================
       ReDim arrPROVARV(0 To uniqueID) As PRODARVPROD
       arrPROVARV(uniqueID).lngID = uniqueID
       arrPROVARV(uniqueID).lngIDPai = lngIDPai
       arrPROVARV(uniqueID).strPRODUTO = Trim(BREC!SGI_PRODLST)
       arrPROVARV(uniqueID).strProdutoPAI = ""
       arrPROVARV(uniqueID).lngTipo = BREC!SGI_PRODUTOTIPO
       arrPROVARV(uniqueID).curQTDCONS = BREC!SGI_QTDE
       arrPROVARV(uniqueID).strUNIDADE = BREC!SGI_UNIDCONS
       arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Ignore
       
       arrPROVARV(uniqueID).lngCODPAI = -1
       arrPROVARV(uniqueID).lngCODIGO = BREC!SGI_CODID
       
       arrPROVARV(uniqueID).lngProdutoID = BREC!SGI_IDPRODUTO
       arrPROVARV(uniqueID).lngProdutoIDPai = BREC!SGI_IDPRODLST
       arrPROVARV(uniqueID).lngCodUniMed = BREC!SGI_CODUNIMED
       
    End If
    BREC.Close
    '' =============================================================================
    
    sSql = "Select " & vbCrLf
    
    sSql = sSql & "Case PRDOUTO.SGI_PRODUTOTIPO" & vbCrLf
    sSql = sSql & "        When 1 Then replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODLINPROD))) + '.' +"
    sSql = sSql & "                    replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODCLIE))) + '.' +"
    sSql = sSql & "                    replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODROTULO))) + '.' +"
    sSql = sSql & "                    (Case"
    sSql = sSql & "                          When PRDOUTO.SGI_DIGVERIF Is Null Then '0'"
    sSql = sSql & "                          When PRDOUTO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(1),PRDOUTO.SGI_DIGVERIF))) End)"
    sSql = sSql & "        When 0 Then PRDOUTO.SGI_CODIGO End As SGI_CODIGO"
    
    sSql = sSql & "      ,PRDOUTO.SGI_DESCRICAO   " & vbCrLf
    sSql = sSql & "      ,PRDOUTO.SGI_PRCCUSTO    " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_PRODLST     " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_QTDE        " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_UNIDCONS    " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_CUSTOUNIT   " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_CUSTOTOTAL  " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_PRODUTO     " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_CODIGO      AS SGI_CODID " & vbCrLf
    
    sSql = sSql & "      ,LISTMAT.SGI_IDPRODUTO   " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_IDPRODLST   " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_CODUNIMED   " & vbCrLf
    
    sSql = sSql & "      ,PRDOUTO.SGI_PRODUTOTIPO " & vbCrLf
    sSql = sSql & "      ,PRDOUTO.SGI_PRECOPROD   " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_LISTAMATPROD   LISTMAT " & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO PRDOUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       LISTMAT.SGI_FILIAL     = " & lngFilial & vbCrLf
    sSql = sSql & "   And LISTMAT.SGI_CODPAI     = " & arrPROVARV(uniqueID).lngCODIGO & vbCrLf
    sSql = sSql & "   And PRDOUTO.SGI_FILIAL     = LISTMAT.SGI_FILIAL          " & vbCrLf
    sSql = sSql & "   And PRDOUTO.SGI_IDPRODUTO  = LISTMAT.SGI_IDPRODUTO    " & vbCrLf
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC.EOF() Then
       
       lngNivel0 = uniqueID
       
       Do While Not BREC.EOF()
          
          uniqueID = (uniqueID + 1)
          
          ReDim Preserve arrPROVARV(0 To uniqueID) As PRODARVPROD
          arrPROVARV(uniqueID).lngID = uniqueID
          arrPROVARV(uniqueID).lngIDPai = lngNivel0
          arrPROVARV(uniqueID).strPRODUTO = Trim(BREC!SGI_CODIGO)
          arrPROVARV(uniqueID).strProdutoPAI = Trim(BREC!SGI_PRODUTO)
          arrPROVARV(uniqueID).lngTipo = BREC!SGI_PRODUTOTIPO
          arrPROVARV(uniqueID).curQTDCONS = BREC!SGI_QTDE
          arrPROVARV(uniqueID).strUNIDADE = BREC!SGI_UNIDCONS
          arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Ignore
          
          arrPROVARV(uniqueID).lngCODPAI = arrPROVARV(lngNivel0).lngCODIGO
          arrPROVARV(uniqueID).lngCODIGO = BREC!SGI_CODID
          
            arrPROVARV(uniqueID).lngProdutoID = BREC!SGI_IDPRODUTO
            arrPROVARV(uniqueID).lngProdutoIDPai = BREC!SGI_IDPRODLST
            arrPROVARV(uniqueID).lngCodUniMed = BREC!SGI_CODUNIMED
          
          '' ==================
          '' Nivel 1
          lngNivel1 = uniqueID
          
          sSql = "Select " & vbCrLf
          
            sSql = sSql & "Case PRDOUTO.SGI_PRODUTOTIPO" & vbCrLf
            sSql = sSql & "        When 1 Then replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODLINPROD))) + '.' +"
            sSql = sSql & "                    replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODCLIE))) + '.' +"
            sSql = sSql & "                    replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODROTULO))) + '.' +"
            sSql = sSql & "                    (Case"
            sSql = sSql & "                          When PRDOUTO.SGI_DIGVERIF Is Null Then '0'"
            sSql = sSql & "                          When PRDOUTO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(1),PRDOUTO.SGI_DIGVERIF))) End)"
            sSql = sSql & "        When 0 Then PRDOUTO.SGI_CODIGO End As SGI_CODIGO"
          
          sSql = sSql & "      ,PRDOUTO.SGI_DESCRICAO   " & vbCrLf
          sSql = sSql & "      ,PRDOUTO.SGI_PRCCUSTO    " & vbCrLf
          sSql = sSql & "      ,LISTMAT.SGI_PRODLST     " & vbCrLf
          sSql = sSql & "      ,LISTMAT.SGI_QTDE        " & vbCrLf
          sSql = sSql & "      ,LISTMAT.SGI_UNIDCONS    " & vbCrLf
          sSql = sSql & "      ,LISTMAT.SGI_CUSTOUNIT   " & vbCrLf
          sSql = sSql & "      ,LISTMAT.SGI_CUSTOTOTAL  " & vbCrLf
          sSql = sSql & "      ,LISTMAT.SGI_PRODUTO     " & vbCrLf
          sSql = sSql & "      ,LISTMAT.SGI_CODIGO      AS SGI_CODID " & vbCrLf
          
            sSql = sSql & "      ,LISTMAT.SGI_IDPRODUTO   " & vbCrLf
            sSql = sSql & "      ,LISTMAT.SGI_IDPRODLST   " & vbCrLf
            sSql = sSql & "      ,LISTMAT.SGI_CODUNIMED   " & vbCrLf
          
          sSql = sSql & "      ,PRDOUTO.SGI_PRODUTOTIPO " & vbCrLf
          sSql = sSql & "      ,PRDOUTO.SGI_PRECOPROD   " & vbCrLf
          sSql = sSql & "  From " & vbCrLf
          sSql = sSql & "       SGI_LISTAMATPROD   LISTMAT " & vbCrLf
          sSql = sSql & "      ,SGI_CADPRODUTO PRDOUTO " & vbCrLf
          sSql = sSql & " Where " & vbCrLf
          sSql = sSql & "       LISTMAT.SGI_FILIAL  = " & lngFilial & vbCrLf
          ''sSql = sSql & "   And LISTMAT.SGI_PRODUTO = '" & Trim(BREC!SGI_PRODLST) & "'" & vbCrLf
          sSql = sSql & "   And LISTMAT.SGI_CODPAI  = " & BREC!SGI_CODID & vbCrLf
          sSql = sSql & "   And PRDOUTO.SGI_FILIAL  = LISTMAT.SGI_FILIAL      " & vbCrLf
          sSql = sSql & "   And PRDOUTO.SGI_IDPRODUTO  = LISTMAT.SGI_IDPRODUTO     " & vbCrLf
          
          BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
          Do While Not BREC2.EOF()
             
             uniqueID = (uniqueID + 1)
             
             ReDim Preserve arrPROVARV(0 To uniqueID) As PRODARVPROD
             arrPROVARV(uniqueID).lngID = uniqueID
             arrPROVARV(uniqueID).lngIDPai = lngNivel1
             arrPROVARV(uniqueID).strPRODUTO = Trim(BREC2!SGI_CODIGO)
             arrPROVARV(uniqueID).strProdutoPAI = Trim(BREC!SGI_PRODLST)
             arrPROVARV(uniqueID).lngTipo = BREC2!SGI_PRODUTOTIPO
             arrPROVARV(uniqueID).curQTDCONS = BREC2!SGI_QTDE
             arrPROVARV(uniqueID).strUNIDADE = BREC2!SGI_UNIDCONS
             arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Ignore
             
             arrPROVARV(uniqueID).lngCODPAI = arrPROVARV(lngNivel1).lngCODIGO
             arrPROVARV(uniqueID).lngCODIGO = BREC2!SGI_CODID
             
                arrPROVARV(uniqueID).lngProdutoID = BREC2!SGI_IDPRODUTO
                arrPROVARV(uniqueID).lngProdutoIDPai = BREC2!SGI_IDPRODLST
                arrPROVARV(uniqueID).lngCodUniMed = BREC2!SGI_CODUNIMED
             
             '' ==================
             '' Nivel 2
             lngNivel2 = uniqueID
             
             sSql = "Select " & vbCrLf
             
                sSql = sSql & "Case PRDOUTO.SGI_PRODUTOTIPO" & vbCrLf
                sSql = sSql & "        When 1 Then replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODLINPROD))) + '.' +"
                sSql = sSql & "                    replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODCLIE))) + '.' +"
                sSql = sSql & "                    replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODROTULO))) + '.' +"
                sSql = sSql & "                    (Case"
                sSql = sSql & "                          When PRDOUTO.SGI_DIGVERIF Is Null Then '0'"
                sSql = sSql & "                          When PRDOUTO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(1),PRDOUTO.SGI_DIGVERIF))) End)"
                sSql = sSql & "        When 0 Then PRDOUTO.SGI_CODIGO End As SGI_CODIGO"
             
             sSql = sSql & "      ,PRDOUTO.SGI_DESCRICAO   " & vbCrLf
             sSql = sSql & "      ,PRDOUTO.SGI_PRCCUSTO    " & vbCrLf
             sSql = sSql & "      ,LISTMAT.SGI_PRODLST     " & vbCrLf
             sSql = sSql & "      ,LISTMAT.SGI_QTDE        " & vbCrLf
             sSql = sSql & "      ,LISTMAT.SGI_UNIDCONS    " & vbCrLf
             sSql = sSql & "      ,LISTMAT.SGI_CUSTOUNIT   " & vbCrLf
             sSql = sSql & "      ,LISTMAT.SGI_CUSTOTOTAL  " & vbCrLf
             sSql = sSql & "      ,LISTMAT.SGI_CODIGO      AS SGI_CODID " & vbCrLf
             
                sSql = sSql & "      ,LISTMAT.SGI_IDPRODUTO   " & vbCrLf
                sSql = sSql & "      ,LISTMAT.SGI_IDPRODLST   " & vbCrLf
                sSql = sSql & "      ,LISTMAT.SGI_CODUNIMED   " & vbCrLf
             
             sSql = sSql & "      ,PRDOUTO.SGI_PRODUTOTIPO " & vbCrLf
             sSql = sSql & "      ,PRDOUTO.SGI_PRECOPROD   " & vbCrLf
             sSql = sSql & "  From " & vbCrLf
             sSql = sSql & "       SGI_LISTAMATPROD   LISTMAT " & vbCrLf
             sSql = sSql & "      ,SGI_CADPRODUTO PRDOUTO " & vbCrLf
             sSql = sSql & " Where " & vbCrLf
             sSql = sSql & "       LISTMAT.SGI_FILIAL  = " & lngFilial & vbCrLf
             ''sSql = sSql & "   And LISTMAT.SGI_PRODUTO = '" & Trim(BREC2!SGI_PRODLST) & "'" & vbCrLf
             sSql = sSql & "   And LISTMAT.SGI_CODPAI  = " & BREC2!SGI_CODID & vbCrLf
             sSql = sSql & "   And PRDOUTO.SGI_FILIAL  = LISTMAT.SGI_FILIAL      " & vbCrLf
             sSql = sSql & "   And PRDOUTO.SGI_IDPRODUTO  = LISTMAT.SGI_IDPRODUTO     " & vbCrLf
             BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
             Do While Not BREC3.EOF()
                uniqueID = (uniqueID + 1)
                
                ReDim Preserve arrPROVARV(0 To uniqueID) As PRODARVPROD
                arrPROVARV(uniqueID).lngID = uniqueID
                arrPROVARV(uniqueID).lngIDPai = lngNivel2
                arrPROVARV(uniqueID).strPRODUTO = Trim(BREC3!SGI_CODIGO)
                arrPROVARV(uniqueID).strProdutoPAI = Trim(BREC2!SGI_PRODLST)
                arrPROVARV(uniqueID).lngTipo = BREC3!SGI_PRODUTOTIPO
                arrPROVARV(uniqueID).curQTDCONS = BREC3!SGI_QTDE
                arrPROVARV(uniqueID).strUNIDADE = BREC3!SGI_UNIDCONS
                arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Ignore
                
                arrPROVARV(uniqueID).lngCODPAI = arrPROVARV(lngNivel2).lngCODIGO
                arrPROVARV(uniqueID).lngCODIGO = BREC3!SGI_CODID
                
                arrPROVARV(uniqueID).lngProdutoID = BREC3!SGI_IDPRODUTO
                arrPROVARV(uniqueID).lngProdutoIDPai = BREC3!SGI_IDPRODLST
                arrPROVARV(uniqueID).lngCodUniMed = BREC3!SGI_CODUNIMED
                
                '' ==================
                '' Nivel 3
                lngNivel3 = uniqueID
                
                sSql = "Select " & vbCrLf
                
                sSql = sSql & "Case PRDOUTO.SGI_PRODUTOTIPO" & vbCrLf
                sSql = sSql & "        When 1 Then replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODLINPROD))) + '.' +"
                sSql = sSql & "                    replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODCLIE))) + '.' +"
                sSql = sSql & "                    replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODROTULO))) + '.' +"
                sSql = sSql & "                    (Case"
                sSql = sSql & "                          When PRDOUTO.SGI_DIGVERIF Is Null Then '0'"
                sSql = sSql & "                          When PRDOUTO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(1),PRDOUTO.SGI_DIGVERIF))) End)"
                sSql = sSql & "        When 0 Then PRDOUTO.SGI_CODIGO End As SGI_CODIGO"
                
                sSql = sSql & "      ,PRDOUTO.SGI_DESCRICAO   " & vbCrLf
                sSql = sSql & "      ,PRDOUTO.SGI_PRCCUSTO    " & vbCrLf
                sSql = sSql & "      ,LISTMAT.SGI_PRODLST     " & vbCrLf
                sSql = sSql & "      ,LISTMAT.SGI_QTDE        " & vbCrLf
                sSql = sSql & "      ,LISTMAT.SGI_UNIDCONS    " & vbCrLf
                sSql = sSql & "      ,LISTMAT.SGI_CUSTOUNIT   " & vbCrLf
                sSql = sSql & "      ,LISTMAT.SGI_CUSTOTOTAL  " & vbCrLf
                sSql = sSql & "      ,LISTMAT.SGI_CODIGO      AS SGI_CODID " & vbCrLf
                
                sSql = sSql & "      ,LISTMAT.SGI_IDPRODUTO   " & vbCrLf
                sSql = sSql & "      ,LISTMAT.SGI_IDPRODLST   " & vbCrLf
                sSql = sSql & "      ,LISTMAT.SGI_CODUNIMED   " & vbCrLf
                
                sSql = sSql & "      ,PRDOUTO.SGI_PRODUTOTIPO " & vbCrLf
                sSql = sSql & "      ,PRDOUTO.SGI_PRECOPROD   " & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_LISTAMATPROD   LISTMAT " & vbCrLf
                sSql = sSql & "      ,SGI_CADPRODUTO PRDOUTO " & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       LISTMAT.SGI_FILIAL  = " & lngFilial & vbCrLf
                ''sSql = sSql & "   And LISTMAT.SGI_PRODUTO = '" & Trim(BREC3!SGI_PRODLST) & "'" & vbCrLf
                sSql = sSql & "   And LISTMAT.SGI_CODPAI  = " & BREC3!SGI_CODID & vbCrLf
                sSql = sSql & "   And PRDOUTO.SGI_FILIAL  = LISTMAT.SGI_FILIAL      " & vbCrLf
                sSql = sSql & "   And PRDOUTO.SGI_IDPRODUTO  = LISTMAT.SGI_IDPRODUTO     " & vbCrLf
                
                BREC4.Open sSql, adoBanco_Dados, adOpenDynamic
                Do While Not BREC4.EOF()
                   uniqueID = (uniqueID + 1)
                    
                   ReDim Preserve arrPROVARV(0 To uniqueID) As PRODARVPROD
                   arrPROVARV(uniqueID).lngID = uniqueID
                   arrPROVARV(uniqueID).lngIDPai = lngNivel3
                   arrPROVARV(uniqueID).strPRODUTO = Trim(BREC4!SGI_CODIGO)
                   arrPROVARV(uniqueID).strProdutoPAI = Trim(BREC3!SGI_PRODLST)
                   arrPROVARV(uniqueID).lngTipo = BREC4!SGI_PRODUTOTIPO
                   arrPROVARV(uniqueID).curQTDCONS = BREC4!SGI_QTDE
                   arrPROVARV(uniqueID).strUNIDADE = BREC4!SGI_UNIDCONS
                   arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Ignore
                
                   arrPROVARV(uniqueID).lngCODPAI = arrPROVARV(lngNivel3).lngCODIGO
                   arrPROVARV(uniqueID).lngCODIGO = BREC4!SGI_CODID
                   
                    arrPROVARV(uniqueID).lngProdutoID = BREC4!SGI_IDPRODUTO
                    arrPROVARV(uniqueID).lngProdutoIDPai = BREC4!SGI_IDPRODLST
                    arrPROVARV(uniqueID).lngCodUniMed = BREC4!SGI_CODUNIMED
                   
                   '' ==================
                   '' Nivel 4
                   lngNivel4 = uniqueID
                   
                   sSql = "Select " & vbCrLf
                   
                    sSql = sSql & "Case PRDOUTO.SGI_PRODUTOTIPO" & vbCrLf
                    sSql = sSql & "        When 1 Then replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODLINPROD))) + '.' +"
                    sSql = sSql & "                    replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODCLIE))) + '.' +"
                    sSql = sSql & "                    replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODROTULO))) + '.' +"
                    sSql = sSql & "                    (Case"
                    sSql = sSql & "                          When PRDOUTO.SGI_DIGVERIF Is Null Then '0'"
                    sSql = sSql & "                          When PRDOUTO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(1),PRDOUTO.SGI_DIGVERIF))) End)"
                    sSql = sSql & "        When 0 Then PRDOUTO.SGI_CODIGO End As SGI_CODIGO"
                   
                   sSql = sSql & "      ,PRDOUTO.SGI_DESCRICAO   " & vbCrLf
                   sSql = sSql & "      ,PRDOUTO.SGI_PRCCUSTO    " & vbCrLf
                   sSql = sSql & "      ,LISTMAT.SGI_PRODLST     " & vbCrLf
                   sSql = sSql & "      ,LISTMAT.SGI_QTDE        " & vbCrLf
                   sSql = sSql & "      ,LISTMAT.SGI_UNIDCONS    " & vbCrLf
                   sSql = sSql & "      ,LISTMAT.SGI_CUSTOUNIT   " & vbCrLf
                   sSql = sSql & "      ,LISTMAT.SGI_CUSTOTOTAL  " & vbCrLf
                   sSql = sSql & "      ,LISTMAT.SGI_CODIGO      AS SGI_CODID " & vbCrLf
                   
                    sSql = sSql & "      ,LISTMAT.SGI_IDPRODUTO   " & vbCrLf
                    sSql = sSql & "      ,LISTMAT.SGI_IDPRODLST   " & vbCrLf
                    sSql = sSql & "      ,LISTMAT.SGI_CODUNIMED   " & vbCrLf
                   
                   sSql = sSql & "      ,PRDOUTO.SGI_PRODUTOTIPO " & vbCrLf
                   sSql = sSql & "      ,PRDOUTO.SGI_PRECOPROD   " & vbCrLf
                   sSql = sSql & "  From " & vbCrLf
                   sSql = sSql & "       SGI_LISTAMATPROD   LISTMAT " & vbCrLf
                   sSql = sSql & "      ,SGI_CADPRODUTO PRDOUTO " & vbCrLf
                   sSql = sSql & " Where " & vbCrLf
                   sSql = sSql & "       LISTMAT.SGI_FILIAL  = " & lngFilial & vbCrLf
                   ''sSql = sSql & "   And LISTMAT.SGI_PRODUTO = '" & Trim(BREC4!SGI_PRODLST) & "'" & vbCrLf
                   sSql = sSql & "   And LISTMAT.SGI_CODPAI  = " & BREC4!SGI_CODID & vbCrLf
                   sSql = sSql & "   And PRDOUTO.SGI_FILIAL  = LISTMAT.SGI_FILIAL      " & vbCrLf
                   sSql = sSql & "   And PRDOUTO.SGI_IDPRODUTO  = LISTMAT.SGI_IDPRODUTO     " & vbCrLf
                   
                   BREC5.Open sSql, adoBanco_Dados, adOpenDynamic
                   Do While Not BREC5.EOF()
                      uniqueID = (uniqueID + 1)
                    
                      ReDim Preserve arrPROVARV(0 To uniqueID) As PRODARVPROD
                      arrPROVARV(uniqueID).lngID = uniqueID
                      arrPROVARV(uniqueID).lngIDPai = lngNivel4
                      arrPROVARV(uniqueID).strPRODUTO = Trim(BREC5!SGI_CODIGO)
                      arrPROVARV(uniqueID).strProdutoPAI = Trim(BREC4!SGI_PRODLST)
                      arrPROVARV(uniqueID).lngTipo = BREC5!SGI_PRODUTOTIPO
                      arrPROVARV(uniqueID).curQTDCONS = BREC5!SGI_QTDE
                      arrPROVARV(uniqueID).strUNIDADE = BREC5!SGI_UNIDCONS
                      arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Ignore
                      
                      arrPROVARV(uniqueID).lngCODPAI = arrPROVARV(lngNivel4).lngCODIGO
                      arrPROVARV(uniqueID).lngCODIGO = BREC5!SGI_CODID
                      
                        arrPROVARV(uniqueID).lngProdutoID = BREC5!SGI_IDPRODUTO
                        arrPROVARV(uniqueID).lngProdutoIDPai = BREC5!SGI_IDPRODLST
                        arrPROVARV(uniqueID).lngCodUniMed = BREC5!SGI_CODUNIMED
                      
                      '' ==================
                      '' Nivel 5
                      lngNivel5 = uniqueID
                      
                      sSql = "Select " & vbCrLf
                      
                        sSql = sSql & "Case PRDOUTO.SGI_PRODUTOTIPO" & vbCrLf
                        sSql = sSql & "        When 1 Then replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODLINPROD))) + '.' +"
                        sSql = sSql & "                    replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODCLIE))) + '.' +"
                        sSql = sSql & "                    replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODROTULO))) + '.' +"
                        sSql = sSql & "                    (Case"
                        sSql = sSql & "                          When PRDOUTO.SGI_DIGVERIF Is Null Then '0'"
                        sSql = sSql & "                          When PRDOUTO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(1),PRDOUTO.SGI_DIGVERIF))) End)"
                        sSql = sSql & "        When 0 Then PRDOUTO.SGI_CODIGO End As SGI_CODIGO"
                      
                      sSql = sSql & "      ,PRDOUTO.SGI_DESCRICAO   " & vbCrLf
                      sSql = sSql & "      ,PRDOUTO.SGI_PRCCUSTO    " & vbCrLf
                      sSql = sSql & "      ,LISTMAT.SGI_PRODLST     " & vbCrLf
                      sSql = sSql & "      ,LISTMAT.SGI_QTDE        " & vbCrLf
                      sSql = sSql & "      ,LISTMAT.SGI_UNIDCONS    " & vbCrLf
                      sSql = sSql & "      ,LISTMAT.SGI_CUSTOUNIT   " & vbCrLf
                      sSql = sSql & "      ,LISTMAT.SGI_CUSTOTOTAL  " & vbCrLf
                      sSql = sSql & "      ,LISTMAT.SGI_CODIGO      AS SGI_CODID " & vbCrLf
                      
                        sSql = sSql & "      ,LISTMAT.SGI_IDPRODUTO   " & vbCrLf
                        sSql = sSql & "      ,LISTMAT.SGI_IDPRODLST   " & vbCrLf
                        sSql = sSql & "      ,LISTMAT.SGI_CODUNIMED   " & vbCrLf
                      
                      sSql = sSql & "      ,PRDOUTO.SGI_PRODUTOTIPO " & vbCrLf
                      sSql = sSql & "      ,PRDOUTO.SGI_PRECOPROD   " & vbCrLf
                      sSql = sSql & "  From " & vbCrLf
                      sSql = sSql & "       SGI_LISTAMATPROD   LISTMAT " & vbCrLf
                      sSql = sSql & "      ,SGI_CADPRODUTO PRDOUTO " & vbCrLf
                      sSql = sSql & " Where " & vbCrLf
                      sSql = sSql & "       LISTMAT.SGI_FILIAL  = " & lngFilial & vbCrLf
                      ''sSql = sSql & "   And LISTMAT.SGI_PRODUTO = '" & Trim(BREC5!SGI_PRODLST) & "'" & vbCrLf
                      sSql = sSql & "   And LISTMAT.SGI_CODPAI  = " & BREC5!SGI_CODID & vbCrLf
                      sSql = sSql & "   And PRDOUTO.SGI_FILIAL  = LISTMAT.SGI_FILIAL      " & vbCrLf
                      sSql = sSql & "   And PRDOUTO.SGI_IDPRODUTO  = LISTMAT.SGI_IDPRODUTO     " & vbCrLf
                      BREC6.Open sSql, adoBanco_Dados, adOpenDynamic
                      
                      Do While Not BREC6.EOF()
                         uniqueID = (uniqueID + 1)
                    
                         ReDim Preserve arrPROVARV(0 To uniqueID) As PRODARVPROD
                         arrPROVARV(uniqueID).lngID = uniqueID
                         arrPROVARV(uniqueID).lngIDPai = lngNivel5
                         arrPROVARV(uniqueID).strPRODUTO = Trim(BREC6!SGI_CODIGO)
                         arrPROVARV(uniqueID).strProdutoPAI = Trim(BREC5!SGI_PRODLST)
                         arrPROVARV(uniqueID).lngTipo = BREC6!SGI_PRODUTOTIPO
                         arrPROVARV(uniqueID).curQTDCONS = BREC6!SGI_QTDE
                         arrPROVARV(uniqueID).strUNIDADE = BREC6!SGI_UNIDCONS
                         arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Ignore
                      
                         arrPROVARV(uniqueID).lngCODPAI = arrPROVARV(lngNivel5).lngCODIGO
                         arrPROVARV(uniqueID).lngCODIGO = BREC6!SGI_CODID
                         
                            arrPROVARV(uniqueID).lngProdutoID = BREC6!SGI_IDPRODUTO
                            arrPROVARV(uniqueID).lngProdutoIDPai = BREC6!SGI_IDPRODLST
                            arrPROVARV(uniqueID).lngCodUniMed = BREC6!SGI_CODUNIMED
                         
                         '' ==================
                         '' Nivel 6
                         lngNivel6 = uniqueID
                         
                         sSql = "Select " & vbCrLf
                         
                            sSql = sSql & "Case PRDOUTO.SGI_PRODUTOTIPO" & vbCrLf
                            sSql = sSql & "        When 1 Then replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODLINPROD))) + '.' +"
                            sSql = sSql & "                    replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODCLIE))) + '.' +"
                            sSql = sSql & "                    replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODROTULO))) + '.' +"
                            sSql = sSql & "                    (Case"
                            sSql = sSql & "                          When PRDOUTO.SGI_DIGVERIF Is Null Then '0'"
                            sSql = sSql & "                          When PRDOUTO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(1),PRDOUTO.SGI_DIGVERIF))) End)"
                            sSql = sSql & "        When 0 Then PRDOUTO.SGI_CODIGO End As SGI_CODIGO"
                         
                         sSql = sSql & "      ,PRDOUTO.SGI_DESCRICAO   " & vbCrLf
                         sSql = sSql & "      ,PRDOUTO.SGI_PRCCUSTO    " & vbCrLf
                         sSql = sSql & "      ,LISTMAT.SGI_PRODLST     " & vbCrLf
                         sSql = sSql & "      ,LISTMAT.SGI_QTDE        " & vbCrLf
                         sSql = sSql & "      ,LISTMAT.SGI_UNIDCONS    " & vbCrLf
                         sSql = sSql & "      ,LISTMAT.SGI_CUSTOUNIT   " & vbCrLf
                         sSql = sSql & "      ,LISTMAT.SGI_CUSTOTOTAL  " & vbCrLf
                         sSql = sSql & "      ,LISTMAT.SGI_CODIGO      AS SGI_CODID " & vbCrLf
                         
                            sSql = sSql & "      ,LISTMAT.SGI_IDPRODUTO   " & vbCrLf
                            sSql = sSql & "      ,LISTMAT.SGI_IDPRODLST   " & vbCrLf
                            sSql = sSql & "      ,LISTMAT.SGI_CODUNIMED   " & vbCrLf
                         
                         sSql = sSql & "      ,PRDOUTO.SGI_PRODUTOTIPO " & vbCrLf
                         sSql = sSql & "      ,PRDOUTO.SGI_PRECOPROD   " & vbCrLf
                         sSql = sSql & "  From " & vbCrLf
                         sSql = sSql & "       SGI_LISTAMATPROD   LISTMAT " & vbCrLf
                         sSql = sSql & "      ,SGI_CADPRODUTO PRDOUTO " & vbCrLf
                         sSql = sSql & " Where " & vbCrLf
                         sSql = sSql & "       LISTMAT.SGI_FILIAL  = " & lngFilial & vbCrLf
                         ''sSql = sSql & "   And LISTMAT.SGI_PRODUTO = '" & Trim(BREC6!SGI_PRODLST) & "'" & vbCrLf
                         sSql = sSql & "   And LISTMAT.SGI_CODPAI  = " & BREC6!SGI_CODID & vbCrLf
                         sSql = sSql & "   And PRDOUTO.SGI_FILIAL  = LISTMAT.SGI_FILIAL      " & vbCrLf
                         sSql = sSql & "   And PRDOUTO.SGI_IDPRODUTO  = LISTMAT.SGI_IDPRODUTO     " & vbCrLf
                         BREC7.Open sSql, adoBanco_Dados, adOpenDynamic
                         Do While Not BREC7.EOF()
                            uniqueID = (uniqueID + 1)
                    
                            ReDim Preserve arrPROVARV(0 To uniqueID) As PRODARVPROD
                            arrPROVARV(uniqueID).lngID = uniqueID
                            arrPROVARV(uniqueID).lngIDPai = lngNivel6
                            arrPROVARV(uniqueID).strPRODUTO = Trim(BREC7!SGI_CODIGO)
                            arrPROVARV(uniqueID).strProdutoPAI = Trim(BREC6!SGI_PRODLST)
                            arrPROVARV(uniqueID).lngTipo = BREC7!SGI_PRODUTOTIPO
                            arrPROVARV(uniqueID).curQTDCONS = BREC7!SGI_QTDE
                            arrPROVARV(uniqueID).strUNIDADE = BREC7!SGI_UNIDCONS
                            arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Ignore
                            
                            arrPROVARV(uniqueID).lngCODPAI = arrPROVARV(lngNivel6).lngCODIGO
                            arrPROVARV(uniqueID).lngCODIGO = BREC7!SGI_CODID
                            
                            arrPROVARV(uniqueID).lngProdutoID = BREC7!SGI_IDPRODUTO
                            arrPROVARV(uniqueID).lngProdutoIDPai = BREC7!SGI_IDPRODLST
                            arrPROVARV(uniqueID).lngCodUniMed = BREC7!SGI_CODUNIMED
                            
                            '' ==================
                            '' Nivel 7
                            lngNivel7 = uniqueID
                            
                            sSql = "Select " & vbCrLf
                            
                            sSql = sSql & "Case PRDOUTO.SGI_PRODUTOTIPO" & vbCrLf
                            sSql = sSql & "        When 1 Then replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODLINPROD))) + '.' +"
                            sSql = sSql & "                    replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODCLIE))) + '.' +"
                            sSql = sSql & "                    replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODROTULO))) + '.' +"
                            sSql = sSql & "                    (Case"
                            sSql = sSql & "                          When PRDOUTO.SGI_DIGVERIF Is Null Then '0'"
                            sSql = sSql & "                          When PRDOUTO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(1),PRDOUTO.SGI_DIGVERIF))) End)"
                            sSql = sSql & "        When 0 Then PRDOUTO.SGI_CODIGO End As SGI_CODIGO"
                            
                            sSql = sSql & "      ,PRDOUTO.SGI_DESCRICAO   " & vbCrLf
                            sSql = sSql & "      ,PRDOUTO.SGI_PRCCUSTO    " & vbCrLf
                            sSql = sSql & "      ,LISTMAT.SGI_PRODLST     " & vbCrLf
                            sSql = sSql & "      ,LISTMAT.SGI_QTDE        " & vbCrLf
                            sSql = sSql & "      ,LISTMAT.SGI_UNIDCONS    " & vbCrLf
                            sSql = sSql & "      ,LISTMAT.SGI_CUSTOUNIT   " & vbCrLf
                            sSql = sSql & "      ,LISTMAT.SGI_CUSTOTOTAL  " & vbCrLf
                            sSql = sSql & "      ,LISTMAT.SGI_CODIGO      AS SGI_CODID " & vbCrLf
                            
                            sSql = sSql & "      ,LISTMAT.SGI_IDPRODUTO   " & vbCrLf
                            sSql = sSql & "      ,LISTMAT.SGI_IDPRODLST   " & vbCrLf
                            sSql = sSql & "      ,LISTMAT.SGI_CODUNIMED   " & vbCrLf
                            
                            sSql = sSql & "      ,PRDOUTO.SGI_PRODUTOTIPO " & vbCrLf
                            sSql = sSql & "      ,PRDOUTO.SGI_PRECOPROD   " & vbCrLf
                            sSql = sSql & "  From " & vbCrLf
                            sSql = sSql & "       SGI_LISTAMATPROD   LISTMAT " & vbCrLf
                            sSql = sSql & "      ,SGI_CADPRODUTO PRDOUTO " & vbCrLf
                            sSql = sSql & " Where " & vbCrLf
                            sSql = sSql & "       LISTMAT.SGI_FILIAL  = " & lngFilial & vbCrLf
                            ''sSql = sSql & "   And LISTMAT.SGI_PRODUTO = '" & Trim(BREC7!SGI_PRODLST) & "'" & vbCrLf
                            sSql = sSql & "   And LISTMAT.SGI_CODPAI  = " & BREC7!SGI_CODID & vbCrLf
                            sSql = sSql & "   And PRDOUTO.SGI_FILIAL  = LISTMAT.SGI_FILIAL      " & vbCrLf
                            sSql = sSql & "   And PRDOUTO.SGI_IDPRODUTO  = LISTMAT.SGI_IDPRODUTO     " & vbCrLf
                            BREC8.Open sSql, adoBanco_Dados, adOpenDynamic
                            Do While Not BREC8.EOF()
                               uniqueID = (uniqueID + 1)
                    
                               ReDim Preserve arrPROVARV(0 To uniqueID) As PRODARVPROD
                               arrPROVARV(uniqueID).lngID = uniqueID
                               arrPROVARV(uniqueID).lngIDPai = lngNivel7
                               arrPROVARV(uniqueID).strPRODUTO = Trim(BREC8!SGI_CODIGO)
                               arrPROVARV(uniqueID).strProdutoPAI = Trim(BREC7!SGI_PRODLST)
                               arrPROVARV(uniqueID).lngTipo = BREC8!SGI_PRODUTOTIPO
                               arrPROVARV(uniqueID).curQTDCONS = BREC8!SGI_QTDE
                               arrPROVARV(uniqueID).strUNIDADE = BREC8!SGI_UNIDCONS
                               arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Ignore
                            
                               arrPROVARV(uniqueID).lngCODPAI = arrPROVARV(lngNivel7).lngCODIGO
                               arrPROVARV(uniqueID).lngCODIGO = BREC8!SGI_CODID
                               
                                arrPROVARV(uniqueID).lngProdutoID = BREC8!SGI_IDPRODUTO
                                arrPROVARV(uniqueID).lngProdutoIDPai = BREC8!SGI_IDPRODLST
                                arrPROVARV(uniqueID).lngCodUniMed = BREC8!SGI_CODUNIMED
                               
                               '' ==================
                               '' Nivel 8
                               lngNivel8 = uniqueID
                               
                               sSql = "Select " & vbCrLf
                               
                                sSql = sSql & "Case PRDOUTO.SGI_PRODUTOTIPO" & vbCrLf
                                sSql = sSql & "        When 1 Then replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODLINPROD))) + '.' +"
                                sSql = sSql & "                    replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODCLIE))) + '.' +"
                                sSql = sSql & "                    replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODROTULO))) + '.' +"
                                sSql = sSql & "                    (Case"
                                sSql = sSql & "                          When PRDOUTO.SGI_DIGVERIF Is Null Then '0'"
                                sSql = sSql & "                          When PRDOUTO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(1),PRDOUTO.SGI_DIGVERIF))) End)"
                                sSql = sSql & "        When 0 Then PRDOUTO.SGI_CODIGO End As SGI_CODIGO"
                               
                               sSql = sSql & "      ,PRDOUTO.SGI_DESCRICAO   " & vbCrLf
                               sSql = sSql & "      ,PRDOUTO.SGI_PRCCUSTO    " & vbCrLf
                               sSql = sSql & "      ,LISTMAT.SGI_PRODLST     " & vbCrLf
                               sSql = sSql & "      ,LISTMAT.SGI_QTDE        " & vbCrLf
                               sSql = sSql & "      ,LISTMAT.SGI_UNIDCONS    " & vbCrLf
                               sSql = sSql & "      ,LISTMAT.SGI_CUSTOUNIT   " & vbCrLf
                               sSql = sSql & "      ,LISTMAT.SGI_CUSTOTOTAL  " & vbCrLf
                               sSql = sSql & "      ,LISTMAT.SGI_CODIGO      AS SGI_CODID " & vbCrLf
                               
                                sSql = sSql & "      ,LISTMAT.SGI_IDPRODUTO   " & vbCrLf
                                sSql = sSql & "      ,LISTMAT.SGI_IDPRODLST   " & vbCrLf
                                sSql = sSql & "      ,LISTMAT.SGI_CODUNIMED   " & vbCrLf
                               
                               sSql = sSql & "      ,PRDOUTO.SGI_PRODUTOTIPO " & vbCrLf
                               sSql = sSql & "      ,PRDOUTO.SGI_PRECOPROD   " & vbCrLf
                               sSql = sSql & "  From " & vbCrLf
                               sSql = sSql & "       SGI_LISTAMATPROD   LISTMAT " & vbCrLf
                               sSql = sSql & "      ,SGI_CADPRODUTO PRDOUTO " & vbCrLf
                               sSql = sSql & " Where " & vbCrLf
                               sSql = sSql & "       LISTMAT.SGI_FILIAL  = " & lngFilial & vbCrLf
                               ''sSql = sSql & "   And LISTMAT.SGI_PRODUTO = '" & Trim(BREC8!SGI_PRODLST) & "'" & vbCrLf
                               sSql = sSql & "   And LISTMAT.SGI_CODPAI  = " & BREC8!SGI_CODID & vbCrLf
                               sSql = sSql & "   And PRDOUTO.SGI_FILIAL  = LISTMAT.SGI_FILIAL      " & vbCrLf
                               sSql = sSql & "   And PRDOUTO.SGI_IDPRODUTO  = LISTMAT.SGI_IDPRODUTO     " & vbCrLf
                               
                               BREC9.Open sSql, adoBanco_Dados, adOpenDynamic
                               Do While Not BREC9.EOF()
                                  uniqueID = (uniqueID + 1)
                    
                                  ReDim Preserve arrPROVARV(0 To uniqueID) As PRODARVPROD
                                  arrPROVARV(uniqueID).lngID = uniqueID
                                  arrPROVARV(uniqueID).lngIDPai = lngNivel8
                                  arrPROVARV(uniqueID).strPRODUTO = Trim(BREC9!SGI_CODIGO)
                                  arrPROVARV(uniqueID).strProdutoPAI = Trim(BREC8!SGI_PRODLST)
                                  arrPROVARV(uniqueID).lngTipo = BREC9!SGI_PRODUTOTIPO
                                  arrPROVARV(uniqueID).curQTDCONS = BREC9!SGI_QTDE
                                  arrPROVARV(uniqueID).strUNIDADE = BREC9!SGI_UNIDCONS
                                  arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Ignore
                               
                                  arrPROVARV(uniqueID).lngCODPAI = arrPROVARV(lngNivel8).lngCODIGO
                                  arrPROVARV(uniqueID).lngCODIGO = BREC9!SGI_CODID
                                  
                                    arrPROVARV(uniqueID).lngProdutoID = BREC9!SGI_IDPRODUTO
                                    arrPROVARV(uniqueID).lngProdutoIDPai = BREC9!SGI_IDPRODLST
                                    arrPROVARV(uniqueID).lngCodUniMed = BREC9!SGI_CODUNIMED
                                  
                                  '' ==================
                                  '' Nivel 9
                                  lngNivel9 = uniqueID
                               
                                  sSql = "Select " & vbCrLf
                                  
                                    sSql = sSql & "Case PRDOUTO.SGI_PRODUTOTIPO" & vbCrLf
                                    sSql = sSql & "        When 1 Then replicate('0',(3 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODLINPROD)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODLINPROD))) + '.' +"
                                    sSql = sSql & "                    replicate('0',(4 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODCLIE)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODCLIE))) + '.' +"
                                    sSql = sSql & "                    replicate('0',(2 - len(Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODROTULO)))))) + Ltrim(Rtrim(Convert(Char(10),PRDOUTO.SGI_CODROTULO))) + '.' +"
                                    sSql = sSql & "                    (Case"
                                    sSql = sSql & "                          When PRDOUTO.SGI_DIGVERIF Is Null Then '0'"
                                    sSql = sSql & "                          When PRDOUTO.SGI_DIGVERIF Is Not Null Then Ltrim(Rtrim(Convert(Char(1),PRDOUTO.SGI_DIGVERIF))) End)"
                                    sSql = sSql & "        When 0 Then PRDOUTO.SGI_CODIGO End As SGI_CODIGO"
                                  
                                  sSql = sSql & "      ,PRDOUTO.SGI_DESCRICAO   " & vbCrLf
                                  sSql = sSql & "      ,PRDOUTO.SGI_PRCCUSTO    " & vbCrLf
                                  sSql = sSql & "      ,LISTMAT.SGI_PRODLST     " & vbCrLf
                                  sSql = sSql & "      ,LISTMAT.SGI_QTDE        " & vbCrLf
                                  sSql = sSql & "      ,LISTMAT.SGI_UNIDCONS    " & vbCrLf
                                  sSql = sSql & "      ,LISTMAT.SGI_CUSTOUNIT   " & vbCrLf
                                  sSql = sSql & "      ,LISTMAT.SGI_CUSTOTOTAL  " & vbCrLf
                                  sSql = sSql & "      ,LISTMAT.SGI_CODIGO      AS SGI_CODID " & vbCrLf
                                  
                                    sSql = sSql & "      ,LISTMAT.SGI_IDPRODUTO   " & vbCrLf
                                    sSql = sSql & "      ,LISTMAT.SGI_IDPRODLST   " & vbCrLf
                                    sSql = sSql & "      ,LISTMAT.SGI_CODUNIMED   " & vbCrLf
                                  
                                  sSql = sSql & "      ,PRDOUTO.SGI_PRODUTOTIPO " & vbCrLf
                                  sSql = sSql & "      ,PRDOUTO.SGI_PRECOPROD   " & vbCrLf
                                  sSql = sSql & "  From " & vbCrLf
                                  sSql = sSql & "       SGI_LISTAMATPROD   LISTMAT " & vbCrLf
                                  sSql = sSql & "      ,SGI_CADPRODUTO PRDOUTO " & vbCrLf
                                  sSql = sSql & " Where " & vbCrLf
                                  sSql = sSql & "       LISTMAT.SGI_FILIAL  = " & lngFilial & vbCrLf
                                  ''sSql = sSql & "   And LISTMAT.SGI_PRODUTO = '" & Trim(BREC9!SGI_PRODLST) & "'" & vbCrLf
                                  sSql = sSql & "   And LISTMAT.SGI_CODPAI  = " & BREC9!SGI_CODID & vbCrLf
                                  sSql = sSql & "   And PRDOUTO.SGI_FILIAL  = LISTMAT.SGI_FILIAL      " & vbCrLf
                                  sSql = sSql & "   And PRDOUTO.SGI_IDPRODUTO  = LISTMAT.SGI_IDPRODUTO     " & vbCrLf
                                  
                                  BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
                                  Do While Not BREC10.EOF()
                                     uniqueID = (uniqueID + 1)
                    
                                     ReDim Preserve arrPROVARV(0 To uniqueID) As PRODARVPROD
                                     arrPROVARV(uniqueID).lngID = uniqueID
                                     arrPROVARV(uniqueID).lngIDPai = lngNivel9
                                     arrPROVARV(uniqueID).strPRODUTO = Trim(BREC10!SGI_CODIGO)
                                     arrPROVARV(uniqueID).strProdutoPAI = Trim(BREC9!SGI_PRODLST)
                                     arrPROVARV(uniqueID).lngTipo = BREC10!SGI_PRODUTOTIPO
                                     arrPROVARV(uniqueID).curQTDCONS = BREC10!SGI_QTDE
                                     arrPROVARV(uniqueID).strUNIDADE = BREC10!SGI_UNIDCONS
                                     arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Ignore
                                  
                                     arrPROVARV(uniqueID).lngCODPAI = arrPROVARV(lngNivel9).lngCODIGO
                                     arrPROVARV(uniqueID).lngCODIGO = BREC10!SGI_CODID
                                     
                                        arrPROVARV(uniqueID).lngProdutoID = BREC10!SGI_IDPRODUTO
                                        arrPROVARV(uniqueID).lngProdutoIDPai = BREC10!SGI_IDPRODLST
                                        arrPROVARV(uniqueID).lngCodUniMed = BREC10!SGI_CODUNIMED
                                     
                                     BREC10.MoveNext
                                  Loop
                                  BREC10.Close
                                  
                                  '' Fin Nivel 9
                                  '' ==================
                                  
                                  BREC9.MoveNext
                               Loop
                               BREC9.Close
                               '' Fin Nivel 8
                               '' ==================
                               
                               BREC8.MoveNext
                            Loop
                            BREC8.Close
                            '' Fin Nivel 7
                            '' ==================
                            
                            BREC7.MoveNext
                         Loop
                         BREC7.Close
                         '' Fin Nivel 6
                         '' ==================
                         
                         BREC6.MoveNext
                      Loop
                      BREC6.Close
                      '' Fin Nivel 5
                      '' ==================
                      
                      
                      BREC5.MoveNext
                   Loop
                   BREC5.Close
                   '' Fin Nivel 4
                   '' ==================
                   
                   BREC4.MoveNext
                Loop
                BREC4.Close
                '' Fin Nivel 3
                '' ==================
                
                BREC3.MoveNext
             Loop
             BREC3.Close
             '' Fin Nivel 2
             '' ==================
             
             BREC2.MoveNext
          Loop
          BREC2.Close
          '' Fin Nivel 1
          '' ==================
          
          BREC.MoveNext
       Loop
    End If
    BREC.Close
    
End Sub

Private Sub CarregaDoArray()

    Dim i As Long
    Dim j As Long
    
    fc.ClearAll
    
    'abilitando drag'n'drop
    fc.RegisterDragDrop
     
    'Seleciona para Poder criar outro box
    fc.ClearSelection
    fc.ExpandOnIncoming = False
     
    '' Carregando o Flow a partid da Array de Dados
    Dim Pai     As box
    Dim Filho   As box
    
    For i = 0 To UBound(arrPROVARV)
        If arrPROVARV(i).intAction2Do = dacEnumUpdateAction_Insert Or _
           arrPROVARV(i).intAction2Do = dacEnumUpdateAction_Ignore Or arrPROVARV(i).intAction2Do = dacEnumUpdateAction_update Then
            If arrPROVARV(i).lngIDPai = -1 Then
               '' Criando o Pai da Estrutura
               Set root = fc.CreateBox(100, 500, 150, 40)
               root.Tag = arrPROVARV(i).lngID
               root.Picture = pb1.Picture
               root.Text = Trim(arrPROVARV(0).strPRODUTO)
               
               root.TextStyle = tsRight
               root.PicturePos = picCenterLeft

               Set company = fc.CreateGroup(root)
            Else
               Set Pai = fc.boxes(arrPROVARV(i).lngIDPai)
               Set Filho = addChildFromArray(Pai, i)
                   
               'fixa como um ícone de nó
               Filho.PicturePos = picCenterLeft
               If arrPROVARV(i).lngTipo = 1 Then Filho.Picture = pb1.Picture
               If arrPROVARV(i).lngTipo = 0 Then Filho.Picture = pb2.Picture
                    
               'mostra o tag do nó
               Filho.Text = Trim(arrPROVARV(i).strPRODUTO)
               Filho.TextStyle = tsRight
                
               '' Call layoutTree(tldLeftToRight)
            End If
        End If
    Next i
    
    optAlinh(0).Value = True
    If optAlinh(0).Value = True Then Call layoutTree(tldLeftToRight)
    
End Sub

Private Sub Inclui()

    Dim i As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    
    fraLinha.Enabled = True
    fraZoom.Enabled = True
    
    fraTipos.Enabled = True
  
    Me.Caption = "Cadastro de arvore de produtos - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADARVPROD
    
    fc.ClearAll
    
    Call Novo
    
End Sub

Private Function ConsisteProd(ByVal box As FLOWCHARTLibCtl.IBoxItem) As Boolean
    ConsisteProd = False
    
    If Len(Trim(arrPROVARV(box.Tag).strPRODUTO)) = 0 Then
       MsgBox "Informe primeiro o Produto Pai !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Function
    End If
    
    ConsisteProd = True
End Function


Private Sub Altera()

    Dim i As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    
    fraLinha.Enabled = True
    fraZoom.Enabled = True
    
    fraTipos.Enabled = True
  
    Me.Caption = "Cadastro de arvore de produtos - [ ALTERAÇÃO ]"
    
    fc.ClearAll
    
    Call CarregaLista(iCodigo, FILIAL)
    Call CarregaDoArray
    
End Sub

Private Function ValidaCampos() As Boolean

     Dim i As Integer
     Dim j As Integer
     ValidaCampos = False
     
     For i = 1 To (UBound(arrPROVARV))
        If Len(Trim(arrPROVARV(i).strPRODUTO)) = 0 Then
           MsgBox "Faltam produtos a serem informados !!!", vbOKOnly + vbExclamation, "Aviso"
           Exit Function
        End If
        If Len(Trim(arrPROVARV(i).strUNIDADE)) = 0 Then
           MsgBox "O Produto " & arrPROVARV(i).strPRODUTO & " falta informar a unidade de medida !!!", vbOKOnly + vbExclamation, "Aviso"
           Exit Function
        End If
     Next i
     
     ValidaCampos = True
End Function

Private Sub CarregaListaAlterada(strPRODUTO As String, lngINDEX As Long)
    
    Dim lngIDPai     As Long
    Dim lngNivel0    As Long
    Dim lngNivel1    As Long
    Dim lngNivel2    As Long
    Dim lngNivel3    As Long
    Dim lngNivel4    As Long
    Dim lngNivel5    As Long
    Dim lngNivel6    As Long
    Dim lngNivel7    As Long
    Dim lngNivel8    As Long
    Dim lngNivel9    As Long
    Dim lngNivel10   As Long
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PRDOUTO.SGI_CODIGO      " & vbCrLf
    sSql = sSql & "      ,PRDOUTO.SGI_DESCRICAO   " & vbCrLf
    sSql = sSql & "      ,PRDOUTO.SGI_PRCCUSTO    " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_PRODLST     " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_QTDE        " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_UNIDCONS    " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_CUSTOUNIT   " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_CUSTOTOTAL  " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_CODIGO      AS SGI_CODID " & vbCrLf
    sSql = sSql & "      ,PRDOUTO.SGI_PRODUTOTIPO " & vbCrLf
    sSql = sSql & "      ,PRDOUTO.SGI_PRECOPROD   " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_LISTAMAT   LISTMAT " & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO PRDOUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       LISTMAT.SGI_FILIAL  = " & FILIAL & vbCrLf
    ''sSql = sSql & "   And LISTMAT.SGI_PRODUTO = '" & Trim(arrPROVARV(lngINDEX).strPRODUTO) & "'" & vbCrLf
    sSql = sSql & "   And LISTMAT.SGI_CODIGO  = " & Trim(arrPROVARV(lngINDEX).lngCODIGO) & vbCrLf
    sSql = sSql & "   And PRDOUTO.SGI_FILIAL  = LISTMAT.SGI_FILIAL     " & vbCrLf
    sSql = sSql & "   And PRDOUTO.SGI_CODIGO  = LISTMAT.SGI_PRODLST    " & vbCrLf
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC.EOF() Then
      
       lngNivel0 = arrPROVARV(lngINDEX).lngID
       
       Do While Not BREC.EOF()
          
          uniqueID = (uniqueID + 1)
          
          ReDim Preserve arrPROVARV(0 To uniqueID) As PRODARVPROD
          arrPROVARV(uniqueID).lngID = uniqueID
          arrPROVARV(uniqueID).lngIDPai = lngNivel0
          arrPROVARV(uniqueID).strPRODUTO = Trim(BREC!SGI_CODIGO)
          arrPROVARV(uniqueID).lngTipo = BREC!SGI_PRODUTOTIPO
          arrPROVARV(uniqueID).curQTDCONS = BREC!SGI_QTDE
          arrPROVARV(uniqueID).strProdutoPAI = Trim(arrPROVARV(lngINDEX).strPRODUTO)
          arrPROVARV(uniqueID).strUNIDADE = Trim(BREC!SGI_UNIDCONS)
          arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Insert
          
          arrPROVARV(uniqueID).lngCODPAI = arrPROVARV(lngNivel0).lngCODIGO
          arrPROVARV(uniqueID).lngCODIGO = BREC!SGI_CODID
      
          '' ==================
          '' Nivel 1
          lngNivel1 = uniqueID
          
          sSql = "Select " & vbCrLf
          sSql = sSql & "       PRDOUTO.SGI_CODIGO      " & vbCrLf
          sSql = sSql & "      ,PRDOUTO.SGI_DESCRICAO   " & vbCrLf
          sSql = sSql & "      ,PRDOUTO.SGI_PRCCUSTO    " & vbCrLf
          sSql = sSql & "      ,LISTMAT.SGI_PRODLST     " & vbCrLf
          sSql = sSql & "      ,LISTMAT.SGI_QTDE        " & vbCrLf
          sSql = sSql & "      ,LISTMAT.SGI_UNIDCONS    " & vbCrLf
          sSql = sSql & "      ,LISTMAT.SGI_CUSTOUNIT   " & vbCrLf
          sSql = sSql & "      ,LISTMAT.SGI_CUSTOTOTAL  " & vbCrLf
          sSql = sSql & "      ,LISTMAT.SGI_CODIGO      AS SGI_CODID " & vbCrLf
          sSql = sSql & "      ,PRDOUTO.SGI_PRODUTOTIPO " & vbCrLf
          sSql = sSql & "      ,PRDOUTO.SGI_PRECOPROD   " & vbCrLf
          sSql = sSql & "  From " & vbCrLf
          sSql = sSql & "       SGI_LISTAMAT   LISTMAT " & vbCrLf
          sSql = sSql & "      ,SGI_CADPRODUTO PRDOUTO " & vbCrLf
          sSql = sSql & " Where " & vbCrLf
          sSql = sSql & "       LISTMAT.SGI_FILIAL  = " & FILIAL & vbCrLf
          ''sSql = sSql & "   And LISTMAT.SGI_PRODUTO = '" & BREC!SGI_PRODLST & "'" & vbCrLf
          sSql = sSql & "   And LISTMAT.SGI_CODIGO = " & BREC!SGI_CODID & vbCrLf
          sSql = sSql & "   And PRDOUTO.SGI_FILIAL  = LISTMAT.SGI_FILIAL      " & vbCrLf
          sSql = sSql & "   And PRDOUTO.SGI_CODIGO  = LISTMAT.SGI_PRODLST     " & vbCrLf
          
          BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
          Do While Not BREC2.EOF()
             
             uniqueID = (uniqueID + 1)
             
             ReDim Preserve arrPROVARV(0 To uniqueID) As PRODARVPROD
             arrPROVARV(uniqueID).lngID = uniqueID
             arrPROVARV(uniqueID).lngIDPai = lngNivel1
             arrPROVARV(uniqueID).strPRODUTO = Trim(BREC2!SGI_CODIGO)
             arrPROVARV(uniqueID).strProdutoPAI = Trim(BREC!SGI_PRODLST)
             arrPROVARV(uniqueID).lngTipo = BREC2!SGI_PRODUTOTIPO
             arrPROVARV(uniqueID).curQTDCONS = BREC2!SGI_QTDE
             arrPROVARV(uniqueID).strUNIDADE = Trim(BREC2!SGI_UNIDCONS)
             arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Insert
             
             arrPROVARV(uniqueID).lngCODPAI = arrPROVARV(lngNivel1).lngCODIGO
             arrPROVARV(uniqueID).lngCODIGO = BREC2!SGI_CODID
             
             '' ==================
             '' Nivel 2
             lngNivel2 = uniqueID
             
             sSql = "Select " & vbCrLf
             sSql = sSql & "       PRDOUTO.SGI_CODIGO      " & vbCrLf
             sSql = sSql & "      ,PRDOUTO.SGI_DESCRICAO   " & vbCrLf
             sSql = sSql & "      ,PRDOUTO.SGI_PRCCUSTO    " & vbCrLf
             sSql = sSql & "      ,LISTMAT.SGI_PRODLST     " & vbCrLf
             sSql = sSql & "      ,LISTMAT.SGI_QTDE        " & vbCrLf
             sSql = sSql & "      ,LISTMAT.SGI_UNIDCONS    " & vbCrLf
             sSql = sSql & "      ,LISTMAT.SGI_CUSTOUNIT   " & vbCrLf
             sSql = sSql & "      ,LISTMAT.SGI_CUSTOTOTAL  " & vbCrLf
             sSql = sSql & "      ,LISTMAT.SGI_CODIGO      AS SGI_CODID " & vbCrLf
             sSql = sSql & "      ,PRDOUTO.SGI_PRODUTOTIPO " & vbCrLf
             sSql = sSql & "      ,PRDOUTO.SGI_PRECOPROD   " & vbCrLf
             sSql = sSql & "  From " & vbCrLf2
             sSql = sSql & "       SGI_LISTAMAT   LISTMAT " & vbCrLf
             sSql = sSql & "      ,SGI_CADPRODUTO PRDOUTO " & vbCrLf
             sSql = sSql & " Where " & vbCrLf
             sSql = sSql & "       LISTMAT.SGI_FILIAL  = " & FILIAL & vbCrLf
             ''sSql = sSql & "   And LISTMAT.SGI_PRODUTO = '" & BREC2!SGI_PRODLST & "'" & vbCrLf
             sSql = sSql & "   And LISTMAT.SGI_CODIGO = " & BREC2!SGI_CODID & vbCrLf
             sSql = sSql & "   And PRDOUTO.SGI_FILIAL  = LISTMAT.SGI_FILIAL      " & vbCrLf
             sSql = sSql & "   And PRDOUTO.SGI_CODIGO  = LISTMAT.SGI_PRODLST     " & vbCrLf
             BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
             Do While Not BREC3.EOF()
                uniqueID = (uniqueID + 1)
                
                ReDim Preserve arrPROVARV(0 To uniqueID) As PRODARVPROD
                arrPROVARV(uniqueID).lngID = uniqueID
                arrPROVARV(uniqueID).lngIDPai = lngNivel2
                arrPROVARV(uniqueID).strPRODUTO = Trim(BREC3!SGI_CODIGO)
                arrPROVARV(uniqueID).strProdutoPAI = Trim(BREC2!SGI_PRODLST)
                arrPROVARV(uniqueID).lngTipo = BREC3!SGI_PRODUTOTIPO
                arrPROVARV(uniqueID).curQTDCONS = BREC3!SGI_QTDE
                arrPROVARV(uniqueID).strUNIDADE = Trim(BREC3!SGI_UNIDCONS)
                arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Insert
                
                arrPROVARV(uniqueID).lngCODPAI = arrPROVARV(lngNivel2).lngCODIGO
                arrPROVARV(uniqueID).lngCODIGO = BREC3!SGI_CODID
                
                '' ==================
                '' Nivel 3
                lngNivel3 = uniqueID
                
                sSql = "Select " & vbCrLf
                sSql = sSql & "       PRDOUTO.SGI_CODIGO      " & vbCrLf
                sSql = sSql & "      ,PRDOUTO.SGI_DESCRICAO   " & vbCrLf
                sSql = sSql & "      ,PRDOUTO.SGI_PRCCUSTO    " & vbCrLf
                sSql = sSql & "      ,LISTMAT.SGI_PRODLST     " & vbCrLf
                sSql = sSql & "      ,LISTMAT.SGI_QTDE        " & vbCrLf
                sSql = sSql & "      ,LISTMAT.SGI_UNIDCONS    " & vbCrLf
                sSql = sSql & "      ,LISTMAT.SGI_CUSTOUNIT   " & vbCrLf
                sSql = sSql & "      ,LISTMAT.SGI_CUSTOTOTAL  " & vbCrLf
                sSql = sSql & "      ,LISTMAT.SGI_CODIGO      AS SGI_CODID " & vbCrLf
                sSql = sSql & "      ,PRDOUTO.SGI_PRODUTOTIPO " & vbCrLf
                sSql = sSql & "      ,PRDOUTO.SGI_PRECOPROD   " & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_LISTAMAT   LISTMAT " & vbCrLf
                sSql = sSql & "      ,SGI_CADPRODUTO PRDOUTO " & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       LISTMAT.SGI_FILIAL  = " & FILIAL & vbCrLf
                ''sSql = sSql & "   And LISTMAT.SGI_PRODUTO = '" & BREC3!SGI_PRODLST & "'" & vbCrLf
                sSql = sSql & "   And LISTMAT.SGI_CODIGO = " & BREC3!SGI_CODID & vbCrLf
                sSql = sSql & "   And PRDOUTO.SGI_FILIAL  = LISTMAT.SGI_FILIAL      " & vbCrLf
                sSql = sSql & "   And PRDOUTO.SGI_CODIGO  = LISTMAT.SGI_PRODLST     " & vbCrLf
                
                BREC4.Open sSql, adoBanco_Dados, adOpenDynamic
                Do While Not BREC4.EOF()
                   uniqueID = (uniqueID + 1)
                    
                   ReDim Preserve arrPROVARV(0 To uniqueID) As PRODARVPROD
                   arrPROVARV(uniqueID).lngID = uniqueID
                   arrPROVARV(uniqueID).lngIDPai = lngNivel3
                   arrPROVARV(uniqueID).strPRODUTO = Trim(BREC4!SGI_CODIGO)
                   arrPROVARV(uniqueID).strProdutoPAI = Trim(BREC3!SGI_PRODLST)
                   arrPROVARV(uniqueID).lngTipo = BREC4!SGI_PRODUTOTIPO
                   arrPROVARV(uniqueID).curQTDCONS = BREC4!SGI_QTDE
                   arrPROVARV(uniqueID).strUNIDADE = Trim(BREC4!SGI_UNIDCONS)
                   arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Insert
                
                   arrPROVARV(uniqueID).lngCODPAI = arrPROVARV(lngNivel3).lngCODIGO
                   arrPROVARV(uniqueID).lngCODIGO = BREC4!SGI_CODID
                   
                   '' ==================
                   '' Nivel 4
                   lngNivel4 = uniqueID
                   
                   sSql = "Select " & vbCrLf
                   sSql = sSql & "       PRDOUTO.SGI_CODIGO      " & vbCrLf
                   sSql = sSql & "      ,PRDOUTO.SGI_DESCRICAO   " & vbCrLf
                   sSql = sSql & "      ,PRDOUTO.SGI_PRCCUSTO    " & vbCrLf
                   sSql = sSql & "      ,LISTMAT.SGI_PRODLST     " & vbCrLf
                   sSql = sSql & "      ,LISTMAT.SGI_QTDE        " & vbCrLf
                   sSql = sSql & "      ,LISTMAT.SGI_UNIDCONS    " & vbCrLf
                   sSql = sSql & "      ,LISTMAT.SGI_CUSTOUNIT   " & vbCrLf
                   sSql = sSql & "      ,LISTMAT.SGI_CUSTOTOTAL  " & vbCrLf
                   sSql = sSql & "      ,LISTMAT.SGI_CODIGO      AS SGI_CODID " & vbCrLf
                   sSql = sSql & "      ,PRDOUTO.SGI_PRODUTOTIPO " & vbCrLf
                   sSql = sSql & "      ,PRDOUTO.SGI_PRECOPROD   " & vbCrLf
                   sSql = sSql & "  From " & vbCrLf
                   sSql = sSql & "       SGI_LISTAMAT   LISTMAT " & vbCrLf
                   sSql = sSql & "      ,SGI_CADPRODUTO PRDOUTO " & vbCrLf
                   sSql = sSql & " Where " & vbCrLf
                   sSql = sSql & "       LISTMAT.SGI_FILIAL  = " & FILIAL & vbCrLf
                   ''sSql = sSql & "   And LISTMAT.SGI_PRODUTO = '" & BREC4!SGI_PRODLST & "'" & vbCrLf
                   sSql = sSql & "   And LISTMAT.SGI_CODIGO  = " & BREC4!SGI_CODID & vbCrLf
                   sSql = sSql & "   And PRDOUTO.SGI_FILIAL  = LISTMAT.SGI_FILIAL      " & vbCrLf
                   sSql = sSql & "   And PRDOUTO.SGI_CODIGO  = LISTMAT.SGI_PRODLST     " & vbCrLf
                   
                   BREC5.Open sSql, adoBanco_Dados, adOpenDynamic
                   Do While Not BREC5.EOF()
                      uniqueID = (uniqueID + 1)
                    
                      ReDim Preserve arrPROVARV(0 To uniqueID) As PRODARVPROD
                      arrPROVARV(uniqueID).lngID = uniqueID
                      arrPROVARV(uniqueID).lngIDPai = lngNivel4
                      arrPROVARV(uniqueID).strPRODUTO = Trim(BREC5!SGI_CODIGO)
                      arrPROVARV(uniqueID).strProdutoPAI = Trim(BREC4!SGI_PRODLST)
                      arrPROVARV(uniqueID).lngTipo = BREC5!SGI_PRODUTOTIPO
                      arrPROVARV(uniqueID).curQTDCONS = BREC5!SGI_QTDE
                      arrPROVARV(uniqueID).strUNIDADE = Trim(BREC5!SGI_UNIDCONS)
                      arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Insert
                      
                      arrPROVARV(uniqueID).lngCODPAI = arrPROVARV(lngNivel4).lngCODIGO
                      arrPROVARV(uniqueID).lngCODIGO = BREC5!SGI_CODID
                      
                      '' ==================
                      '' Nivel 5
                      lngNivel5 = uniqueID
                      
                      sSql = "Select " & vbCrLf
                      sSql = sSql & "       PRDOUTO.SGI_CODIGO      " & vbCrLf
                      sSql = sSql & "      ,PRDOUTO.SGI_DESCRICAO   " & vbCrLf
                      sSql = sSql & "      ,PRDOUTO.SGI_PRCCUSTO    " & vbCrLf
                      sSql = sSql & "      ,LISTMAT.SGI_PRODLST     " & vbCrLf
                      sSql = sSql & "      ,LISTMAT.SGI_QTDE        " & vbCrLf
                      sSql = sSql & "      ,LISTMAT.SGI_UNIDCONS    " & vbCrLf
                      sSql = sSql & "      ,LISTMAT.SGI_CUSTOUNIT   " & vbCrLf
                      sSql = sSql & "      ,LISTMAT.SGI_CUSTOTOTAL  " & vbCrLf
                      sSql = sSql & "      ,LISTMAT.SGI_CODIGO      AS SGI_CODID " & vbCrLf
                      sSql = sSql & "      ,PRDOUTO.SGI_PRODUTOTIPO " & vbCrLf
                      sSql = sSql & "      ,PRDOUTO.SGI_PRECOPROD   " & vbCrLf
                      sSql = sSql & "  From " & vbCrLf
                      sSql = sSql & "       SGI_LISTAMAT   LISTMAT " & vbCrLf
                      sSql = sSql & "      ,SGI_CADPRODUTO PRDOUTO " & vbCrLf
                      sSql = sSql & " Where " & vbCrLf
                      sSql = sSql & "       LISTMAT.SGI_FILIAL  = " & FILIAL & vbCrLf
                      ''sSql = sSql & "   And LISTMAT.SGI_PRODUTO = '" & BREC5!SGI_PRODLST & "'" & vbCrLf
                      sSql = sSql & "   And LISTMAT.SGI_CODIGO  = " & BREC5!SGI_CODID & vbCrLf
                      sSql = sSql & "   And PRDOUTO.SGI_FILIAL  = LISTMAT.SGI_FILIAL      " & vbCrLf
                      sSql = sSql & "   And PRDOUTO.SGI_CODIGO  = LISTMAT.SGI_PRODLST     " & vbCrLf
                      BREC6.Open sSql, adoBanco_Dados, adOpenDynamic
                      
                      Do While Not BREC6.EOF()
                         uniqueID = (uniqueID + 1)
                    
                         ReDim Preserve arrPROVARV(0 To uniqueID) As PRODARVPROD
                         arrPROVARV(uniqueID).lngID = uniqueID
                         arrPROVARV(uniqueID).lngIDPai = lngNivel5
                         arrPROVARV(uniqueID).strPRODUTO = Trim(BREC6!SGI_CODIGO)
                         arrPROVARV(uniqueID).strProdutoPAI = Trim(BREC5!SGI_PRODLST)
                         arrPROVARV(uniqueID).lngTipo = BREC6!SGI_PRODUTOTIPO
                         arrPROVARV(uniqueID).curQTDCONS = BREC6!SGI_QTDE
                         arrPROVARV(uniqueID).strUNIDADE = Trim(BREC6!SGI_UNIDCONS)
                         arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Insert
                      
                         arrPROVARV(uniqueID).lngCODPAI = arrPROVARV(lngNivel5).lngCODIGO
                         arrPROVARV(uniqueID).lngCODIGO = BREC6!SGI_CODID
                         
                         '' ==================
                         '' Nivel 6
                         lngNivel6 = uniqueID
                         
                         sSql = "Select " & vbCrLf
                         sSql = sSql & "       PRDOUTO.SGI_CODIGO      " & vbCrLf
                         sSql = sSql & "      ,PRDOUTO.SGI_DESCRICAO   " & vbCrLf
                         sSql = sSql & "      ,PRDOUTO.SGI_PRCCUSTO    " & vbCrLf
                         sSql = sSql & "      ,LISTMAT.SGI_PRODLST     " & vbCrLf
                         sSql = sSql & "      ,LISTMAT.SGI_QTDE        " & vbCrLf
                         sSql = sSql & "      ,LISTMAT.SGI_UNIDCONS    " & vbCrLf
                         sSql = sSql & "      ,LISTMAT.SGI_CUSTOUNIT   " & vbCrLf
                         sSql = sSql & "      ,LISTMAT.SGI_CUSTOTOTAL  " & vbCrLf
                         sSql = sSql & "      ,LISTMAT.SGI_CODIGO      AS SGI_CODID " & vbCrLf
                         sSql = sSql & "      ,PRDOUTO.SGI_PRODUTOTIPO " & vbCrLf
                         sSql = sSql & "      ,PRDOUTO.SGI_PRECOPROD   " & vbCrLf
                         sSql = sSql & "  From " & vbCrLf
                         sSql = sSql & "       SGI_LISTAMAT   LISTMAT " & vbCrLf
                         sSql = sSql & "      ,SGI_CADPRODUTO PRDOUTO " & vbCrLf
                         sSql = sSql & " Where " & vbCrLf
                         sSql = sSql & "       LISTMAT.SGI_FILIAL  = " & FILIAL & vbCrLf
                         ''sSql = sSql & "   And LISTMAT.SGI_PRODUTO = '" & BREC6!SGI_PRODLST & "'" & vbCrLf
                         sSql = sSql & "   And LISTMAT.SGI_CODIGO = " & BREC6!SGI_CODID & vbCrLf
                         sSql = sSql & "   And PRDOUTO.SGI_FILIAL  = LISTMAT.SGI_FILIAL      " & vbCrLf
                         sSql = sSql & "   And PRDOUTO.SGI_CODIGO  = LISTMAT.SGI_PRODLST     " & vbCrLf
                         BREC7.Open sSql, adoBanco_Dados, adOpenDynamic
                         Do While Not BREC7.EOF()
                            uniqueID = (uniqueID + 1)
                    
                            ReDim Preserve arrPROVARV(0 To uniqueID) As PRODARVPROD
                            arrPROVARV(uniqueID).lngID = uniqueID
                            arrPROVARV(uniqueID).lngIDPai = lngNivel6
                            arrPROVARV(uniqueID).strPRODUTO = Trim(BREC7!SGI_CODIGO)
                            arrPROVARV(uniqueID).strProdutoPAI = Trim(BREC6!SGI_PRODLST)
                            arrPROVARV(uniqueID).lngTipo = BREC7!SGI_PRODUTOTIPO
                            arrPROVARV(uniqueID).curQTDCONS = BREC7!SGI_QTDE
                            arrPROVARV(uniqueID).strUNIDADE = Trim(BREC7!SGI_UNIDCONS)
                            arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Insert
                               
                            arrPROVARV(uniqueID).lngCODPAI = arrPROVARV(lngNivel6).lngCODIGO
                            arrPROVARV(uniqueID).lngCODIGO = BREC7!SGI_CODID
                            
                            '' ==================
                            '' Nivel 7
                            lngNivel7 = uniqueID
                            
                            sSql = "Select " & vbCrLf
                            sSql = sSql & "       PRDOUTO.SGI_CODIGO      " & vbCrLf
                            sSql = sSql & "      ,PRDOUTO.SGI_DESCRICAO   " & vbCrLf
                            sSql = sSql & "      ,PRDOUTO.SGI_PRCCUSTO    " & vbCrLf
                            sSql = sSql & "      ,LISTMAT.SGI_PRODLST     " & vbCrLf
                            sSql = sSql & "      ,LISTMAT.SGI_QTDE        " & vbCrLf
                            sSql = sSql & "      ,LISTMAT.SGI_UNIDCONS    " & vbCrLf
                            sSql = sSql & "      ,LISTMAT.SGI_CUSTOUNIT   " & vbCrLf
                            sSql = sSql & "      ,LISTMAT.SGI_CUSTOTOTAL  " & vbCrLf
                            sSql = sSql & "      ,LISTMAT.SGI_CODIGO      AS SGI_CODID " & vbCrLf
                            sSql = sSql & "      ,PRDOUTO.SGI_PRODUTOTIPO " & vbCrLf
                            sSql = sSql & "      ,PRDOUTO.SGI_PRECOPROD   " & vbCrLf
                            sSql = sSql & "  From " & vbCrLf
                            sSql = sSql & "       SGI_LISTAMAT   LISTMAT " & vbCrLf
                            sSql = sSql & "      ,SGI_CADPRODUTO PRDOUTO " & vbCrLf
                            sSql = sSql & " Where " & vbCrLf
                            sSql = sSql & "       LISTMAT.SGI_FILIAL  = " & FILIAL & vbCrLf
                            ''sSql = sSql & "   And LISTMAT.SGI_PRODUTO = '" & BREC7!SGI_PRODLST & "'" & vbCrLf
                            sSql = sSql & "   And LISTMAT.SGI_CODIGO = " & BREC7!SGI_CODID & vbCrLf
                            sSql = sSql & "   And PRDOUTO.SGI_FILIAL  = LISTMAT.SGI_FILIAL      " & vbCrLf
                            sSql = sSql & "   And PRDOUTO.SGI_CODIGO  = LISTMAT.SGI_PRODLST     " & vbCrLf
                            BREC8.Open sSql, adoBanco_Dados, adOpenDynamic
                            Do While Not BREC8.EOF()
                               uniqueID = (uniqueID + 1)
                    
                               ReDim Preserve arrPROVARV(0 To uniqueID) As PRODARVPROD
                               arrPROVARV(uniqueID).lngID = uniqueID
                               arrPROVARV(uniqueID).lngIDPai = lngNivel7
                               arrPROVARV(uniqueID).strPRODUTO = Trim(BREC8!SGI_CODIGO)
                               arrPROVARV(uniqueID).strProdutoPAI = Trim(BREC7!SGI_PRODLST)
                               arrPROVARV(uniqueID).lngTipo = BREC8!SGI_PRODUTOTIPO
                               arrPROVARV(uniqueID).curQTDCONS = BREC8!SGI_QTDE
                               arrPROVARV(uniqueID).strUNIDADE = Trim(BREC8!SGI_UNIDCONS)
                               arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Insert
                            
                               arrPROVARV(uniqueID).lngCODPAI = arrPROVARV(lngNivel7).lngCODIGO
                               arrPROVARV(uniqueID).lngCODIGO = BREC8!SGI_CODID
                               
                               '' ==================
                               '' Nivel 8
                               lngNivel8 = uniqueID
                               
                               sSql = "Select " & vbCrLf
                               sSql = sSql & "       PRDOUTO.SGI_CODIGO      " & vbCrLf
                               sSql = sSql & "      ,PRDOUTO.SGI_DESCRICAO   " & vbCrLf
                               sSql = sSql & "      ,PRDOUTO.SGI_PRCCUSTO    " & vbCrLf
                               sSql = sSql & "      ,LISTMAT.SGI_PRODLST     " & vbCrLf
                               sSql = sSql & "      ,LISTMAT.SGI_QTDE        " & vbCrLf
                               sSql = sSql & "      ,LISTMAT.SGI_UNIDCONS    " & vbCrLf
                               sSql = sSql & "      ,LISTMAT.SGI_CUSTOUNIT   " & vbCrLf
                               sSql = sSql & "      ,LISTMAT.SGI_CUSTOTOTAL  " & vbCrLf
                               sSql = sSql & "      ,LISTMAT.SGI_CODIGO      AS SGI_CODID " & vbCrLf
                               sSql = sSql & "      ,PRDOUTO.SGI_PRODUTOTIPO " & vbCrLf
                               sSql = sSql & "      ,PRDOUTO.SGI_PRECOPROD   " & vbCrLf
                               sSql = sSql & "  From " & vbCrLf
                               sSql = sSql & "       SGI_LISTAMAT   LISTMAT " & vbCrLf
                               sSql = sSql & "      ,SGI_CADPRODUTO PRDOUTO " & vbCrLf
                               sSql = sSql & " Where " & vbCrLf
                               sSql = sSql & "       LISTMAT.SGI_FILIAL  = " & FILIAL & vbCrLf
                               ''sSql = sSql & "   And LISTMAT.SGI_PRODUTO = '" & BREC8!SGI_PRODLST & "'" & vbCrLf
                               sSql = sSql & "   And LISTMAT.SGI_CODIGO = " & BREC8!SGI_CODID & vbCrLf
                               sSql = sSql & "   And PRDOUTO.SGI_FILIAL  = LISTMAT.SGI_FILIAL      " & vbCrLf
                               sSql = sSql & "   And PRDOUTO.SGI_CODIGO  = LISTMAT.SGI_PRODLST     " & vbCrLf
                               
                               BREC9.Open sSql, adoBanco_Dados, adOpenDynamic
                               Do While Not BREC9.EOF()
                                  uniqueID = (uniqueID + 1)
                    
                                  ReDim Preserve arrPROVARV(0 To uniqueID) As PRODARVPROD
                                  arrPROVARV(uniqueID).lngID = uniqueID
                                  arrPROVARV(uniqueID).lngIDPai = lngNivel8
                                  arrPROVARV(uniqueID).strPRODUTO = Trim(BREC9!SGI_CODIGO)
                                  arrPROVARV(uniqueID).strProdutoPAI = Trim(BREC8!SGI_PRODLST)
                                  arrPROVARV(uniqueID).lngTipo = BREC9!SGI_PRODUTOTIPO
                                  arrPROVARV(uniqueID).curQTDCONS = BREC9!SGI_QTDE
                                  arrPROVARV(uniqueID).strUNIDADE = Trim(BREC9!SGI_UNIDCONS)
                                  arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Insert
                               
                                  arrPROVARV(uniqueID).lngCODPAI = arrPROVARV(lngNivel8).lngCODIGO
                                  arrPROVARV(uniqueID).lngCODIGO = BREC9!SGI_CODID
                                  
                                  '' ==================
                                  '' Nivel 9
                                  lngNivel9 = uniqueID
                               
                                  sSql = "Select " & vbCrLf
                                  sSql = sSql & "       PRDOUTO.SGI_CODIGO      " & vbCrLf
                                  sSql = sSql & "      ,PRDOUTO.SGI_DESCRICAO   " & vbCrLf
                                  sSql = sSql & "      ,PRDOUTO.SGI_PRCCUSTO    " & vbCrLf
                                  sSql = sSql & "      ,LISTMAT.SGI_PRODLST     " & vbCrLf
                                  sSql = sSql & "      ,LISTMAT.SGI_QTDE        " & vbCrLf
                                  sSql = sSql & "      ,LISTMAT.SGI_UNIDCONS    " & vbCrLf
                                  sSql = sSql & "      ,LISTMAT.SGI_CUSTOUNIT   " & vbCrLf
                                  sSql = sSql & "      ,LISTMAT.SGI_CUSTOTOTAL  " & vbCrLf
                                  sSql = sSql & "      ,LISTMAT.SGI_CODIGO      AS SGI_CODID " & vbCrLf
                                  sSql = sSql & "      ,PRDOUTO.SGI_PRODUTOTIPO " & vbCrLf
                                  sSql = sSql & "      ,PRDOUTO.SGI_PRECOPROD   " & vbCrLf
                                  sSql = sSql & "  From " & vbCrLf
                                  sSql = sSql & "       SGI_LISTAMAT   LISTMAT " & vbCrLf
                                  sSql = sSql & "      ,SGI_CADPRODUTO PRDOUTO " & vbCrLf
                                  sSql = sSql & " Where " & vbCrLf
                                  sSql = sSql & "       LISTMAT.SGI_FILIAL  = " & FILIAL & vbCrLf
                                  ''sSql = sSql & "   And LISTMAT.SGI_PRODUTO = '" & BREC9!SGI_PRODLST & "'" & vbCrLf
                                  sSql = sSql & "   And LISTMAT.SGI_CODIGO = " & BREC9!SGI_CODID & vbCrLf
                                  sSql = sSql & "   And PRDOUTO.SGI_FILIAL  = LISTMAT.SGI_FILIAL      " & vbCrLf
                                  sSql = sSql & "   And PRDOUTO.SGI_CODIGO  = LISTMAT.SGI_PRODLST     " & vbCrLf
                                  
                                  BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
                                  Do While Not BREC10.EOF()
                                     uniqueID = (uniqueID + 1)
                    
                                     ReDim Preserve arrPROVARV(0 To uniqueID) As PRODARVPROD
                                     arrPROVARV(uniqueID).lngID = uniqueID
                                     arrPROVARV(uniqueID).lngIDPai = lngNivel9
                                     arrPROVARV(uniqueID).strPRODUTO = Trim(BREC10!SGI_CODIGO)
                                     arrPROVARV(uniqueID).strProdutoPAI = Trim(BREC9!SGI_PRODLST)
                                     arrPROVARV(uniqueID).lngTipo = BREC10!SGI_PRODUTOTIPO
                                     arrPROVARV(uniqueID).curQTDCONS = BREC10!SGI_QTDE
                                     arrPROVARV(uniqueID).strUNIDADE = Trim(BREC10!SGI_UNIDCONS)
                                     arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Insert
                                  
                                     arrPROVARV(uniqueID).lngCODPAI = arrPROVARV(lngNivel9).lngCODIGO
                                     arrPROVARV(uniqueID).lngCODIGO = BREC10!SGI_CODID
                                  
                                     BREC10.MoveNext
                                  Loop
                                  BREC10.Close
                                  
                                  '' Fin Nivel 9
                                  '' ==================
                                  
                                  BREC9.MoveNext
                               Loop
                               BREC9.Close
                               '' Fin Nivel 8
                               '' ==================
                               
                               BREC8.MoveNext
                            Loop
                            BREC8.Close
                            '' Fin Nivel 7
                            '' ==================
                            
                            BREC7.MoveNext
                         Loop
                         BREC7.Close
                         '' Fin Nivel 6
                         '' ==================
                         
                         BREC6.MoveNext
                      Loop
                      BREC6.Close
                      '' Fin Nivel 5
                      '' ==================
                      
                      
                      BREC5.MoveNext
                   Loop
                   BREC5.Close
                   '' Fin Nivel 4
                   '' ==================
                   
                   BREC4.MoveNext
                Loop
                BREC4.Close
                '' Fin Nivel 3
                '' ==================
                
                BREC3.MoveNext
             Loop
             BREC3.Close
             '' Fin Nivel 2
             '' ==================
             
             BREC2.MoveNext
          Loop
          BREC2.Close
          '' Fin Nivel 1
          '' ==================
          
          BREC.MoveNext
       Loop
        
    End If
    BREC.Close
    
End Sub

Private Sub RefazIndice()
    '' Refazendo os Indices
    Dim i           As Integer
    Dim lngINDICE   As Integer
    Dim arrTEMP()   As PRODARVPROD
    
    lngINDICE = 0
    For i = 0 To (fc.boxes.Count - 1)
        If arrPROVARV(fc.boxes(i).Tag).intAction2Do = dacEnumUpdateAction_Ignore Or _
           arrPROVARV(fc.boxes(i).Tag).intAction2Do = dacEnumUpdateAction_Insert Or arrPROVARV(fc.boxes(i).Tag).intAction2Do = dacEnumUpdateAction_update Then
           
           ReDim Preserve arrTEMP(0 To lngINDICE) As PRODARVPROD
           arrTEMP(lngINDICE).intAction2Do = arrPROVARV(fc.boxes(i).Tag).intAction2Do
           arrTEMP(lngINDICE).strPRODUTO = arrPROVARV(fc.boxes(i).Tag).strPRODUTO
           arrTEMP(lngINDICE).strUNIDADE = arrPROVARV(fc.boxes(i).Tag).strUNIDADE
           arrTEMP(lngINDICE).curQTDCONS = arrPROVARV(fc.boxes(i).Tag).curQTDCONS
           arrTEMP(lngINDICE).lngTipo = arrPROVARV(fc.boxes(i).Tag).lngTipo
           arrTEMP(lngINDICE).strProdutoPAI = arrPROVARV(fc.boxes(i).Tag).strProdutoPAI
           arrTEMP(lngINDICE).lngCODPAI = arrPROVARV(fc.boxes(i).Tag).lngCODPAI
           arrTEMP(lngINDICE).lngCODIGO = arrPROVARV(fc.boxes(i).Tag).lngCODIGO
           arrTEMP(lngINDICE).lngProdutoID = arrPROVARV(fc.boxes(i).Tag).lngProdutoID
           arrTEMP(lngINDICE).lngProdutoIDPai = arrPROVARV(fc.boxes(i).Tag).lngProdutoIDPai
           arrTEMP(lngINDICE).lngCodUniMed = arrPROVARV(fc.boxes(i).Tag).lngCodUniMed
           
           If arrPROVARV(fc.boxes(i).Tag).lngIDPai <> -1 Then
              arrPROVARV(fc.boxes(i).Tag).lngIDPai = fc.boxes(i).IncomingArrows(0).OriginBox.Tag
              arrTEMP(lngINDICE).lngIDPai = arrPROVARV(fc.boxes(i).Tag).lngIDPai
           Else
              arrPROVARV(fc.boxes(i).Tag).lngIDPai = -1
              arrTEMP(lngINDICE).lngIDPai = arrPROVARV(fc.boxes(i).Tag).lngIDPai
           End If
              
           arrPROVARV(fc.boxes(i).Tag).lngID = lngINDICE
           arrTEMP(lngINDICE).lngID = arrPROVARV(fc.boxes(i).Tag).lngID
           
           fc.boxes(i).Tag = lngINDICE
           
           uniqueID = lngINDICE
           lngINDICE = lngINDICE + 1
        End If
    Next i
    
    arrPROVARV = arrTEMP
End Sub

Private Sub BoxAlterado(ByVal box As FLOWCHARTLibCtl.IBoxItem)
    
    Dim i                   As Integer
    Dim intDestino          As Integer
    
    Dim boxHorig            As FLOWCHARTLibCtl.IBoxItem
    Dim boxAnt              As FLOWCHARTLibCtl.IBoxItem
    
    Set boxHorig = box
    
    Dim SetaDeSaida         As FLOWCHARTLibCtl.IArrows
    Set SetaDeSaida = box.OutgoingArrows
    
    If SetaDeSaida.Count > 0 Then
VOLTA:
          For i = 0 To (SetaDeSaida.Count - 1)
              Set box = SetaDeSaida.Item(i).DestinationBox
              Set SetaDeSaida = box.OutgoingArrows
              If SetaDeSaida.Count = 0 Then
                 arrPROVARV(box.Tag).intAction2Do = dacEnumUpdateAction_delete
                 strITENSDELETED = strITENSDELETED & arrPROVARV(box.Tag).lngCODIGO & "|"
                 fc.DeleteItem box
                 Set box = boxHorig
                 Set SetaDeSaida = box.OutgoingArrows
              End If
              GoTo VOLTA
          Next i
    End If
End Sub

Public Sub CarregaListaAlterando(Optional strPRODUTO As String)
    
    Dim lngIDPai     As Long
    Dim lngNivel0    As Long
    Dim lngNivel1    As Long
    Dim lngNivel2    As Long
    Dim lngNivel3    As Long
    Dim lngNivel4    As Long
    Dim lngNivel5    As Long
    Dim lngNivel6    As Long
    Dim lngNivel7    As Long
    Dim lngNivel8    As Long
    Dim lngNivel9    As Long
    Dim lngNivel10   As Long
    
    '' ========================================
    '' Criando a matrix para começar a arvore
    uniqueID = 0
    lngIDPai = -1
    
    '' ========================================
    ReDim arrPROVARV(0 To uniqueID) As PRODARVPROD
    arrPROVARV(uniqueID).lngID = uniqueID
    arrPROVARV(uniqueID).lngIDPai = lngIDPai
    arrPROVARV(uniqueID).strPRODUTO = Trim(strPRODUTO)
    arrPROVARV(uniqueID).strProdutoPAI = ""
    arrPROVARV(uniqueID).lngTipo = 1
    arrPROVARV(uniqueID).curQTDCONS = 0
    arrPROVARV(uniqueID).strUNIDADE = ""
    arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Ignore
    
    arrPROVARV(uniqueID).lngCODPAI = -1
    arrPROVARV(uniqueID).lngCODIGO = objCADARVPROD.Gera_Codigo(Me.Name)
       
    sSql = "Select " & vbCrLf
    sSql = sSql & "       PRDOUTO.SGI_CODIGO      " & vbCrLf
    sSql = sSql & "      ,PRDOUTO.SGI_DESCRICAO   " & vbCrLf
    sSql = sSql & "      ,PRDOUTO.SGI_PRCCUSTO    " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_PRODLST     " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_QTDE        " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_UNIDCONS    " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_CUSTOUNIT   " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_CUSTOTOTAL  " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_PRODUTO     " & vbCrLf
    sSql = sSql & "      ,LISTMAT.SGI_CODIGO      AS SGI_CODID " & vbCrLf
    sSql = sSql & "      ,PRDOUTO.SGI_PRODUTOTIPO " & vbCrLf
    sSql = sSql & "      ,PRDOUTO.SGI_PRECOPROD   " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_LISTAMAT   LISTMAT " & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO PRDOUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       LISTMAT.SGI_FILIAL  = " & FILIAL & vbCrLf
    sSql = sSql & "   And LISTMAT.SGI_PRODUTO  = '" & Trim(strPRODUTO) & "'" & vbCrLf
    sSql = sSql & "   And PRDOUTO.SGI_FILIAL  = LISTMAT.SGI_FILIAL     " & vbCrLf
    sSql = sSql & "   And PRDOUTO.SGI_CODIGO  = LISTMAT.SGI_PRODLST    " & vbCrLf
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC.EOF() Then
       
       lngNivel0 = uniqueID
       
       Do While Not BREC.EOF()
          
          uniqueID = (uniqueID + 1)
          
          ReDim Preserve arrPROVARV(0 To uniqueID) As PRODARVPROD
          arrPROVARV(uniqueID).lngID = uniqueID
          arrPROVARV(uniqueID).lngIDPai = lngNivel0
          arrPROVARV(uniqueID).strPRODUTO = Trim(BREC!SGI_CODIGO)
          arrPROVARV(uniqueID).strProdutoPAI = Trim(BREC!SGI_PRODUTO)
          arrPROVARV(uniqueID).lngTipo = BREC!SGI_PRODUTOTIPO
          arrPROVARV(uniqueID).curQTDCONS = BREC!SGI_QTDE
          arrPROVARV(uniqueID).strUNIDADE = BREC!SGI_UNIDCONS
          arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Ignore
          
          arrPROVARV(uniqueID).lngCODPAI = arrPROVARV(lngNivel0).lngCODIGO
          If Not IsNull(BREC!SGI_CODID) Then
             arrPROVARV(uniqueID).lngCODIGO = BREC!SGI_CODID
          Else
             arrPROVARV(uniqueID).lngCODIGO = objCADARVPROD.Gera_Codigo(Me.Name)
          End If
          
          '' ==================
          '' Nivel 1
          lngNivel1 = uniqueID
          
          sSql = "Select " & vbCrLf
          sSql = sSql & "       PRDOUTO.SGI_CODIGO      " & vbCrLf
          sSql = sSql & "      ,PRDOUTO.SGI_DESCRICAO   " & vbCrLf
          sSql = sSql & "      ,PRDOUTO.SGI_PRCCUSTO    " & vbCrLf
          sSql = sSql & "      ,LISTMAT.SGI_PRODLST     " & vbCrLf
          sSql = sSql & "      ,LISTMAT.SGI_QTDE        " & vbCrLf
          sSql = sSql & "      ,LISTMAT.SGI_UNIDCONS    " & vbCrLf
          sSql = sSql & "      ,LISTMAT.SGI_CUSTOUNIT   " & vbCrLf
          sSql = sSql & "      ,LISTMAT.SGI_CUSTOTOTAL  " & vbCrLf
          sSql = sSql & "      ,LISTMAT.SGI_PRODUTO     " & vbCrLf
          sSql = sSql & "      ,LISTMAT.SGI_CODIGO      AS SGI_CODID " & vbCrLf
          sSql = sSql & "      ,PRDOUTO.SGI_PRODUTOTIPO " & vbCrLf
          sSql = sSql & "      ,PRDOUTO.SGI_PRECOPROD   " & vbCrLf
          sSql = sSql & "  From " & vbCrLf
          sSql = sSql & "       SGI_LISTAMAT   LISTMAT " & vbCrLf
          sSql = sSql & "      ,SGI_CADPRODUTO PRDOUTO " & vbCrLf
          sSql = sSql & " Where " & vbCrLf
          sSql = sSql & "       LISTMAT.SGI_FILIAL  = " & FILIAL & vbCrLf
          sSql = sSql & "   And LISTMAT.SGI_PRODUTO = '" & Trim(BREC!SGI_PRODLST) & "'" & vbCrLf
          ''sSql = sSql & "   And LISTMAT.SGI_CODPAI  = " & BREC!SGI_CODID & vbCrLf
          sSql = sSql & "   And PRDOUTO.SGI_FILIAL  = LISTMAT.SGI_FILIAL      " & vbCrLf
          sSql = sSql & "   And PRDOUTO.SGI_CODIGO  = LISTMAT.SGI_PRODLST     " & vbCrLf
          
          BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
          Do While Not BREC2.EOF()
             
             uniqueID = (uniqueID + 1)
             
             ReDim Preserve arrPROVARV(0 To uniqueID) As PRODARVPROD
             arrPROVARV(uniqueID).lngID = uniqueID
             arrPROVARV(uniqueID).lngIDPai = lngNivel1
             arrPROVARV(uniqueID).strPRODUTO = Trim(BREC2!SGI_CODIGO)
             arrPROVARV(uniqueID).strProdutoPAI = Trim(BREC!SGI_PRODLST)
             arrPROVARV(uniqueID).lngTipo = BREC2!SGI_PRODUTOTIPO
             arrPROVARV(uniqueID).curQTDCONS = BREC2!SGI_QTDE
             arrPROVARV(uniqueID).strUNIDADE = BREC2!SGI_UNIDCONS
             arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Ignore
             
             arrPROVARV(uniqueID).lngCODPAI = arrPROVARV(lngNivel1).lngCODIGO
             If Not IsNull(BREC2!SGI_CODID) Then
                arrPROVARV(uniqueID).lngCODIGO = BREC2!SGI_CODID
             Else
                arrPROVARV(uniqueID).lngCODIGO = objCADARVPROD.Gera_Codigo(Me.Name)
             End If
             
             '' ==================
             '' Nivel 2
             lngNivel2 = uniqueID
             
             sSql = "Select " & vbCrLf
             sSql = sSql & "       PRDOUTO.SGI_CODIGO      " & vbCrLf
             sSql = sSql & "      ,PRDOUTO.SGI_DESCRICAO   " & vbCrLf
             sSql = sSql & "      ,PRDOUTO.SGI_PRCCUSTO    " & vbCrLf
             sSql = sSql & "      ,LISTMAT.SGI_PRODLST     " & vbCrLf
             sSql = sSql & "      ,LISTMAT.SGI_QTDE        " & vbCrLf
             sSql = sSql & "      ,LISTMAT.SGI_UNIDCONS    " & vbCrLf
             sSql = sSql & "      ,LISTMAT.SGI_CUSTOUNIT   " & vbCrLf
             sSql = sSql & "      ,LISTMAT.SGI_CUSTOTOTAL  " & vbCrLf
             sSql = sSql & "      ,LISTMAT.SGI_CODIGO      AS SGI_CODID " & vbCrLf
             sSql = sSql & "      ,PRDOUTO.SGI_PRODUTOTIPO " & vbCrLf
             sSql = sSql & "      ,PRDOUTO.SGI_PRECOPROD   " & vbCrLf
             sSql = sSql & "  From " & vbCrLf
             sSql = sSql & "       SGI_LISTAMAT   LISTMAT " & vbCrLf
             sSql = sSql & "      ,SGI_CADPRODUTO PRDOUTO " & vbCrLf
             sSql = sSql & " Where " & vbCrLf
             sSql = sSql & "       LISTMAT.SGI_FILIAL  = " & FILIAL & vbCrLf
             sSql = sSql & "   And LISTMAT.SGI_PRODUTO = '" & Trim(BREC2!SGI_PRODLST) & "'" & vbCrLf
             ''sSql = sSql & "   And LISTMAT.SGI_CODPAI  = " & BREC2!SGI_CODID & vbCrLf
             sSql = sSql & "   And PRDOUTO.SGI_FILIAL  = LISTMAT.SGI_FILIAL      " & vbCrLf
             sSql = sSql & "   And PRDOUTO.SGI_CODIGO  = LISTMAT.SGI_PRODLST     " & vbCrLf
             BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
             Do While Not BREC3.EOF()
                uniqueID = (uniqueID + 1)
                
                ReDim Preserve arrPROVARV(0 To uniqueID) As PRODARVPROD
                arrPROVARV(uniqueID).lngID = uniqueID
                arrPROVARV(uniqueID).lngIDPai = lngNivel2
                arrPROVARV(uniqueID).strPRODUTO = Trim(BREC3!SGI_CODIGO)
                arrPROVARV(uniqueID).strProdutoPAI = Trim(BREC2!SGI_PRODLST)
                arrPROVARV(uniqueID).lngTipo = BREC3!SGI_PRODUTOTIPO
                arrPROVARV(uniqueID).curQTDCONS = BREC3!SGI_QTDE
                arrPROVARV(uniqueID).strUNIDADE = BREC3!SGI_UNIDCONS
                arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Ignore
                
                arrPROVARV(uniqueID).lngCODPAI = arrPROVARV(lngNivel2).lngCODIGO
                If Not IsNull(BREC3!SGI_CODID) Then
                   arrPROVARV(uniqueID).lngCODIGO = BREC3!SGI_CODID
                Else
                   arrPROVARV(uniqueID).lngCODIGO = objCADARVPROD.Gera_Codigo(Me.Name)
                End If
                
                '' ==================
                '' Nivel 3
                lngNivel3 = uniqueID
                
                sSql = "Select " & vbCrLf
                sSql = sSql & "       PRDOUTO.SGI_CODIGO      " & vbCrLf
                sSql = sSql & "      ,PRDOUTO.SGI_DESCRICAO   " & vbCrLf
                sSql = sSql & "      ,PRDOUTO.SGI_PRCCUSTO    " & vbCrLf
                sSql = sSql & "      ,LISTMAT.SGI_PRODLST     " & vbCrLf
                sSql = sSql & "      ,LISTMAT.SGI_QTDE        " & vbCrLf
                sSql = sSql & "      ,LISTMAT.SGI_UNIDCONS    " & vbCrLf
                sSql = sSql & "      ,LISTMAT.SGI_CUSTOUNIT   " & vbCrLf
                sSql = sSql & "      ,LISTMAT.SGI_CUSTOTOTAL  " & vbCrLf
                sSql = sSql & "      ,LISTMAT.SGI_CODIGO      AS SGI_CODID " & vbCrLf
                sSql = sSql & "      ,PRDOUTO.SGI_PRODUTOTIPO " & vbCrLf
                sSql = sSql & "      ,PRDOUTO.SGI_PRECOPROD   " & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_LISTAMAT   LISTMAT " & vbCrLf
                sSql = sSql & "      ,SGI_CADPRODUTO PRDOUTO " & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       LISTMAT.SGI_FILIAL  = " & FILIAL & vbCrLf
                sSql = sSql & "   And LISTMAT.SGI_PRODUTO = '" & Trim(BREC3!SGI_PRODLST) & "'" & vbCrLf
                ''sSql = sSql & "   And LISTMAT.SGI_CODPAI  = " & BREC3!SGI_CODID & vbCrLf
                sSql = sSql & "   And PRDOUTO.SGI_FILIAL  = LISTMAT.SGI_FILIAL      " & vbCrLf
                sSql = sSql & "   And PRDOUTO.SGI_CODIGO  = LISTMAT.SGI_PRODLST     " & vbCrLf
                
                BREC4.Open sSql, adoBanco_Dados, adOpenDynamic
                Do While Not BREC4.EOF()
                   uniqueID = (uniqueID + 1)
                    
                   ReDim Preserve arrPROVARV(0 To uniqueID) As PRODARVPROD
                   arrPROVARV(uniqueID).lngID = uniqueID
                   arrPROVARV(uniqueID).lngIDPai = lngNivel3
                   arrPROVARV(uniqueID).strPRODUTO = Trim(BREC4!SGI_CODIGO)
                   arrPROVARV(uniqueID).strProdutoPAI = Trim(BREC3!SGI_PRODLST)
                   arrPROVARV(uniqueID).lngTipo = BREC4!SGI_PRODUTOTIPO
                   arrPROVARV(uniqueID).curQTDCONS = BREC4!SGI_QTDE
                   arrPROVARV(uniqueID).strUNIDADE = BREC4!SGI_UNIDCONS
                   arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Ignore
                
                   arrPROVARV(uniqueID).lngCODPAI = arrPROVARV(lngNivel3).lngCODIGO
                   If Not IsNull(BREC4!SGI_CODID) Then
                      arrPROVARV(uniqueID).lngCODIGO = BREC4!SGI_CODID
                   Else
                      arrPROVARV(uniqueID).lngCODIGO = objCADARVPROD.Gera_Codigo(Me.Name)
                   End If
                   
                   '' ==================
                   '' Nivel 4
                   lngNivel4 = uniqueID
                   
                   sSql = "Select " & vbCrLf
                   sSql = sSql & "       PRDOUTO.SGI_CODIGO      " & vbCrLf
                   sSql = sSql & "      ,PRDOUTO.SGI_DESCRICAO   " & vbCrLf
                   sSql = sSql & "      ,PRDOUTO.SGI_PRCCUSTO    " & vbCrLf
                   sSql = sSql & "      ,LISTMAT.SGI_PRODLST     " & vbCrLf
                   sSql = sSql & "      ,LISTMAT.SGI_QTDE        " & vbCrLf
                   sSql = sSql & "      ,LISTMAT.SGI_UNIDCONS    " & vbCrLf
                   sSql = sSql & "      ,LISTMAT.SGI_CUSTOUNIT   " & vbCrLf
                   sSql = sSql & "      ,LISTMAT.SGI_CUSTOTOTAL  " & vbCrLf
                   sSql = sSql & "      ,LISTMAT.SGI_CODIGO      AS SGI_CODID " & vbCrLf
                   sSql = sSql & "      ,PRDOUTO.SGI_PRODUTOTIPO " & vbCrLf
                   sSql = sSql & "      ,PRDOUTO.SGI_PRECOPROD   " & vbCrLf
                   sSql = sSql & "  From " & vbCrLf
                   sSql = sSql & "       SGI_LISTAMAT   LISTMAT " & vbCrLf
                   sSql = sSql & "      ,SGI_CADPRODUTO PRDOUTO " & vbCrLf
                   sSql = sSql & " Where " & vbCrLf
                   sSql = sSql & "       LISTMAT.SGI_FILIAL  = " & FILIAL & vbCrLf
                   sSql = sSql & "   And LISTMAT.SGI_PRODUTO = '" & Trim(BREC4!SGI_PRODLST) & "'" & vbCrLf
                   ''sSql = sSql & "   And LISTMAT.SGI_CODPAI  = " & BREC4!SGI_CODID & vbCrLf
                   sSql = sSql & "   And PRDOUTO.SGI_FILIAL  = LISTMAT.SGI_FILIAL      " & vbCrLf
                   sSql = sSql & "   And PRDOUTO.SGI_CODIGO  = LISTMAT.SGI_PRODLST     " & vbCrLf
                   
                   BREC5.Open sSql, adoBanco_Dados, adOpenDynamic
                   Do While Not BREC5.EOF()
                      uniqueID = (uniqueID + 1)
                    
                      ReDim Preserve arrPROVARV(0 To uniqueID) As PRODARVPROD
                      arrPROVARV(uniqueID).lngID = uniqueID
                      arrPROVARV(uniqueID).lngIDPai = lngNivel4
                      arrPROVARV(uniqueID).strPRODUTO = Trim(BREC5!SGI_CODIGO)
                      arrPROVARV(uniqueID).strProdutoPAI = Trim(BREC4!SGI_PRODLST)
                      arrPROVARV(uniqueID).lngTipo = BREC5!SGI_PRODUTOTIPO
                      arrPROVARV(uniqueID).curQTDCONS = BREC5!SGI_QTDE
                      arrPROVARV(uniqueID).strUNIDADE = BREC5!SGI_UNIDCONS
                      arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Ignore
                      
                      arrPROVARV(uniqueID).lngCODPAI = arrPROVARV(lngNivel4).lngCODIGO
                      If Not IsNull(BREC5!SGI_CODID) Then
                         arrPROVARV(uniqueID).lngCODIGO = BREC5!SGI_CODID
                      Else
                         arrPROVARV(uniqueID).lngCODIGO = objCADARVPROD.Gera_Codigo(Me.Name)
                      End If
                      
                      '' ==================
                      '' Nivel 5
                      lngNivel5 = uniqueID
                      
                      sSql = "Select " & vbCrLf
                      sSql = sSql & "       PRDOUTO.SGI_CODIGO      " & vbCrLf
                      sSql = sSql & "      ,PRDOUTO.SGI_DESCRICAO   " & vbCrLf
                      sSql = sSql & "      ,PRDOUTO.SGI_PRCCUSTO    " & vbCrLf
                      sSql = sSql & "      ,LISTMAT.SGI_PRODLST     " & vbCrLf
                      sSql = sSql & "      ,LISTMAT.SGI_QTDE        " & vbCrLf
                      sSql = sSql & "      ,LISTMAT.SGI_UNIDCONS    " & vbCrLf
                      sSql = sSql & "      ,LISTMAT.SGI_CUSTOUNIT   " & vbCrLf
                      sSql = sSql & "      ,LISTMAT.SGI_CUSTOTOTAL  " & vbCrLf
                      sSql = sSql & "      ,LISTMAT.SGI_CODIGO      AS SGI_CODID " & vbCrLf
                      sSql = sSql & "      ,PRDOUTO.SGI_PRODUTOTIPO " & vbCrLf
                      sSql = sSql & "      ,PRDOUTO.SGI_PRECOPROD   " & vbCrLf
                      sSql = sSql & "  From " & vbCrLf
                      sSql = sSql & "       SGI_LISTAMAT   LISTMAT " & vbCrLf
                      sSql = sSql & "      ,SGI_CADPRODUTO PRDOUTO " & vbCrLf
                      sSql = sSql & " Where " & vbCrLf
                      sSql = sSql & "       LISTMAT.SGI_FILIAL  = " & FILIAL & vbCrLf
                      sSql = sSql & "   And LISTMAT.SGI_PRODUTO = '" & Trim(BREC5!SGI_PRODLST) & "'" & vbCrLf
                      ''sSql = sSql & "   And LISTMAT.SGI_CODPAI  = " & BREC5!SGI_CODID & vbCrLf
                      sSql = sSql & "   And PRDOUTO.SGI_FILIAL  = LISTMAT.SGI_FILIAL      " & vbCrLf
                      sSql = sSql & "   And PRDOUTO.SGI_CODIGO  = LISTMAT.SGI_PRODLST     " & vbCrLf
                      BREC6.Open sSql, adoBanco_Dados, adOpenDynamic
                      
                      Do While Not BREC6.EOF()
                         uniqueID = (uniqueID + 1)
                    
                         ReDim Preserve arrPROVARV(0 To uniqueID) As PRODARVPROD
                         arrPROVARV(uniqueID).lngID = uniqueID
                         arrPROVARV(uniqueID).lngIDPai = lngNivel5
                         arrPROVARV(uniqueID).strPRODUTO = Trim(BREC6!SGI_CODIGO)
                         arrPROVARV(uniqueID).strProdutoPAI = Trim(BREC5!SGI_PRODLST)
                         arrPROVARV(uniqueID).lngTipo = BREC6!SGI_PRODUTOTIPO
                         arrPROVARV(uniqueID).curQTDCONS = BREC6!SGI_QTDE
                         arrPROVARV(uniqueID).strUNIDADE = BREC6!SGI_UNIDCONS
                         arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Ignore
                      
                         arrPROVARV(uniqueID).lngCODPAI = arrPROVARV(lngNivel5).lngCODIGO
                         If Not IsNull(BREC6!SGI_CODID) Then
                            arrPROVARV(uniqueID).lngCODIGO = BREC6!SGI_CODID
                         Else
                            arrPROVARV(uniqueID).lngCODIGO = objCADARVPROD.Gera_Codigo(Me.Name)
                         End If
                         
                         '' ==================
                         '' Nivel 6
                         lngNivel6 = uniqueID
                         
                         sSql = "Select " & vbCrLf
                         sSql = sSql & "       PRDOUTO.SGI_CODIGO      " & vbCrLf
                         sSql = sSql & "      ,PRDOUTO.SGI_DESCRICAO   " & vbCrLf
                         sSql = sSql & "      ,PRDOUTO.SGI_PRCCUSTO    " & vbCrLf
                         sSql = sSql & "      ,LISTMAT.SGI_PRODLST     " & vbCrLf
                         sSql = sSql & "      ,LISTMAT.SGI_QTDE        " & vbCrLf
                         sSql = sSql & "      ,LISTMAT.SGI_UNIDCONS    " & vbCrLf
                         sSql = sSql & "      ,LISTMAT.SGI_CUSTOUNIT   " & vbCrLf
                         sSql = sSql & "      ,LISTMAT.SGI_CUSTOTOTAL  " & vbCrLf
                         sSql = sSql & "      ,LISTMAT.SGI_CODIGO      AS SGI_CODID " & vbCrLf
                         sSql = sSql & "      ,PRDOUTO.SGI_PRODUTOTIPO " & vbCrLf
                         sSql = sSql & "      ,PRDOUTO.SGI_PRECOPROD   " & vbCrLf
                         sSql = sSql & "  From " & vbCrLf
                         sSql = sSql & "       SGI_LISTAMAT   LISTMAT " & vbCrLf
                         sSql = sSql & "      ,SGI_CADPRODUTO PRDOUTO " & vbCrLf
                         sSql = sSql & " Where " & vbCrLf
                         sSql = sSql & "       LISTMAT.SGI_FILIAL  = " & FILIAL & vbCrLf
                         sSql = sSql & "   And LISTMAT.SGI_PRODUTO = '" & Trim(BREC6!SGI_PRODLST) & "'" & vbCrLf
                         ''sSql = sSql & "   And LISTMAT.SGI_CODPAI  = " & BREC6!SGI_CODID & vbCrLf
                         sSql = sSql & "   And PRDOUTO.SGI_FILIAL  = LISTMAT.SGI_FILIAL      " & vbCrLf
                         sSql = sSql & "   And PRDOUTO.SGI_CODIGO  = LISTMAT.SGI_PRODLST     " & vbCrLf
                         BREC7.Open sSql, adoBanco_Dados, adOpenDynamic
                         Do While Not BREC7.EOF()
                            uniqueID = (uniqueID + 1)
                    
                            ReDim Preserve arrPROVARV(0 To uniqueID) As PRODARVPROD
                            arrPROVARV(uniqueID).lngID = uniqueID
                            arrPROVARV(uniqueID).lngIDPai = lngNivel6
                            arrPROVARV(uniqueID).strPRODUTO = Trim(BREC7!SGI_CODIGO)
                            arrPROVARV(uniqueID).strProdutoPAI = Trim(BREC6!SGI_PRODLST)
                            arrPROVARV(uniqueID).lngTipo = BREC7!SGI_PRODUTOTIPO
                            arrPROVARV(uniqueID).curQTDCONS = BREC7!SGI_QTDE
                            arrPROVARV(uniqueID).strUNIDADE = BREC7!SGI_UNIDCONS
                            arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Ignore
                            
                            arrPROVARV(uniqueID).lngCODPAI = arrPROVARV(lngNivel6).lngCODIGO
                            If Not IsNull(BREC7!SGI_CODID) Then
                               arrPROVARV(uniqueID).lngCODIGO = BREC7!SGI_CODID
                            Else
                               arrPROVARV(uniqueID).lngCODIGO = objCADARVPROD.Gera_Codigo(Me.Name)
                            End If
                            
                            '' ==================
                            '' Nivel 7
                            lngNivel7 = uniqueID
                            
                            sSql = "Select " & vbCrLf
                            sSql = sSql & "       PRDOUTO.SGI_CODIGO      " & vbCrLf
                            sSql = sSql & "      ,PRDOUTO.SGI_DESCRICAO   " & vbCrLf
                            sSql = sSql & "      ,PRDOUTO.SGI_PRCCUSTO    " & vbCrLf
                            sSql = sSql & "      ,LISTMAT.SGI_PRODLST     " & vbCrLf
                            sSql = sSql & "      ,LISTMAT.SGI_QTDE        " & vbCrLf
                            sSql = sSql & "      ,LISTMAT.SGI_UNIDCONS    " & vbCrLf
                            sSql = sSql & "      ,LISTMAT.SGI_CUSTOUNIT   " & vbCrLf
                            sSql = sSql & "      ,LISTMAT.SGI_CUSTOTOTAL  " & vbCrLf
                            sSql = sSql & "      ,LISTMAT.SGI_CODIGO      AS SGI_CODID " & vbCrLf
                            sSql = sSql & "      ,PRDOUTO.SGI_PRODUTOTIPO " & vbCrLf
                            sSql = sSql & "      ,PRDOUTO.SGI_PRECOPROD   " & vbCrLf
                            sSql = sSql & "  From " & vbCrLf
                            sSql = sSql & "       SGI_LISTAMAT   LISTMAT " & vbCrLf
                            sSql = sSql & "      ,SGI_CADPRODUTO PRDOUTO " & vbCrLf
                            sSql = sSql & " Where " & vbCrLf
                            sSql = sSql & "       LISTMAT.SGI_FILIAL  = " & FILIAL & vbCrLf
                            sSql = sSql & "   And LISTMAT.SGI_PRODUTO = '" & Trim(BREC7!SGI_PRODLST) & "'" & vbCrLf
                            ''sSql = sSql & "   And LISTMAT.SGI_CODPAI  = " & BREC7!SGI_CODID & vbCrLf
                            sSql = sSql & "   And PRDOUTO.SGI_FILIAL  = LISTMAT.SGI_FILIAL      " & vbCrLf
                            sSql = sSql & "   And PRDOUTO.SGI_CODIGO  = LISTMAT.SGI_PRODLST     " & vbCrLf
                            BREC8.Open sSql, adoBanco_Dados, adOpenDynamic
                            Do While Not BREC8.EOF()
                               uniqueID = (uniqueID + 1)
                    
                               ReDim Preserve arrPROVARV(0 To uniqueID) As PRODARVPROD
                               arrPROVARV(uniqueID).lngID = uniqueID
                               arrPROVARV(uniqueID).lngIDPai = lngNivel7
                               arrPROVARV(uniqueID).strPRODUTO = Trim(BREC8!SGI_CODIGO)
                               arrPROVARV(uniqueID).strProdutoPAI = Trim(BREC7!SGI_PRODLST)
                               arrPROVARV(uniqueID).lngTipo = BREC8!SGI_PRODUTOTIPO
                               arrPROVARV(uniqueID).curQTDCONS = BREC8!SGI_QTDE
                               arrPROVARV(uniqueID).strUNIDADE = BREC8!SGI_UNIDCONS
                               arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Ignore
                            
                               arrPROVARV(uniqueID).lngCODPAI = arrPROVARV(lngNivel7).lngCODIGO
                               If Not IsNull(BREC8!SGI_CODID) Then
                                  arrPROVARV(uniqueID).lngCODIGO = BREC8!SGI_CODID
                               Else
                                  arrPROVARV(uniqueID).lngCODIGO = objCADARVPROD.Gera_Codigo(Me.Name)
                               End If
                               
                               '' ==================
                               '' Nivel 8
                               lngNivel8 = uniqueID
                               
                               sSql = "Select " & vbCrLf
                               sSql = sSql & "       PRDOUTO.SGI_CODIGO      " & vbCrLf
                               sSql = sSql & "      ,PRDOUTO.SGI_DESCRICAO   " & vbCrLf
                               sSql = sSql & "      ,PRDOUTO.SGI_PRCCUSTO    " & vbCrLf
                               sSql = sSql & "      ,LISTMAT.SGI_PRODLST     " & vbCrLf
                               sSql = sSql & "      ,LISTMAT.SGI_QTDE        " & vbCrLf
                               sSql = sSql & "      ,LISTMAT.SGI_UNIDCONS    " & vbCrLf
                               sSql = sSql & "      ,LISTMAT.SGI_CUSTOUNIT   " & vbCrLf
                               sSql = sSql & "      ,LISTMAT.SGI_CUSTOTOTAL  " & vbCrLf
                               sSql = sSql & "      ,LISTMAT.SGI_CODIGO      AS SGI_CODID " & vbCrLf
                               sSql = sSql & "      ,PRDOUTO.SGI_PRODUTOTIPO " & vbCrLf
                               sSql = sSql & "      ,PRDOUTO.SGI_PRECOPROD   " & vbCrLf
                               sSql = sSql & "  From " & vbCrLf
                               sSql = sSql & "       SGI_LISTAMAT   LISTMAT " & vbCrLf
                               sSql = sSql & "      ,SGI_CADPRODUTO PRDOUTO " & vbCrLf
                               sSql = sSql & " Where " & vbCrLf
                               sSql = sSql & "       LISTMAT.SGI_FILIAL  = " & FILIAL & vbCrLf
                               sSql = sSql & "   And LISTMAT.SGI_PRODUTO = '" & Trim(BREC8!SGI_PRODLST) & "'" & vbCrLf
                               ''sSql = sSql & "   And LISTMAT.SGI_CODPAI  = " & BREC8!SGI_CODID & vbCrLf
                               sSql = sSql & "   And PRDOUTO.SGI_FILIAL  = LISTMAT.SGI_FILIAL      " & vbCrLf
                               sSql = sSql & "   And PRDOUTO.SGI_CODIGO  = LISTMAT.SGI_PRODLST     " & vbCrLf
                               
                               BREC9.Open sSql, adoBanco_Dados, adOpenDynamic
                               Do While Not BREC9.EOF()
                                  uniqueID = (uniqueID + 1)
                    
                                  ReDim Preserve arrPROVARV(0 To uniqueID) As PRODARVPROD
                                  arrPROVARV(uniqueID).lngID = uniqueID
                                  arrPROVARV(uniqueID).lngIDPai = lngNivel8
                                  arrPROVARV(uniqueID).strPRODUTO = Trim(BREC9!SGI_CODIGO)
                                  arrPROVARV(uniqueID).strProdutoPAI = Trim(BREC8!SGI_PRODLST)
                                  arrPROVARV(uniqueID).lngTipo = BREC9!SGI_PRODUTOTIPO
                                  arrPROVARV(uniqueID).curQTDCONS = BREC9!SGI_QTDE
                                  arrPROVARV(uniqueID).strUNIDADE = BREC9!SGI_UNIDCONS
                                  arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Ignore
                               
                                  arrPROVARV(uniqueID).lngCODPAI = arrPROVARV(lngNivel8).lngCODIGO
                                  If Not IsNull(BREC9!SGI_CODID) Then
                                     arrPROVARV(uniqueID).lngCODIGO = BREC9!SGI_CODID
                                  Else
                                     arrPROVARV(uniqueID).lngCODIGO = objCADARVPROD.Gera_Codigo(Me.Name)
                                  End If
                                  
                                  '' ==================
                                  '' Nivel 9
                                  lngNivel9 = uniqueID
                               
                                  sSql = "Select " & vbCrLf
                                  sSql = sSql & "       PRDOUTO.SGI_CODIGO      " & vbCrLf
                                  sSql = sSql & "      ,PRDOUTO.SGI_DESCRICAO   " & vbCrLf
                                  sSql = sSql & "      ,PRDOUTO.SGI_PRCCUSTO    " & vbCrLf
                                  sSql = sSql & "      ,LISTMAT.SGI_PRODLST     " & vbCrLf
                                  sSql = sSql & "      ,LISTMAT.SGI_QTDE        " & vbCrLf
                                  sSql = sSql & "      ,LISTMAT.SGI_UNIDCONS    " & vbCrLf
                                  sSql = sSql & "      ,LISTMAT.SGI_CUSTOUNIT   " & vbCrLf
                                  sSql = sSql & "      ,LISTMAT.SGI_CUSTOTOTAL  " & vbCrLf
                                  sSql = sSql & "      ,LISTMAT.SGI_CODIGO      AS SGI_CODID " & vbCrLf
                                  sSql = sSql & "      ,PRDOUTO.SGI_PRODUTOTIPO " & vbCrLf
                                  sSql = sSql & "      ,PRDOUTO.SGI_PRECOPROD   " & vbCrLf
                                  sSql = sSql & "  From " & vbCrLf
                                  sSql = sSql & "       SGI_LISTAMAT   LISTMAT " & vbCrLf
                                  sSql = sSql & "      ,SGI_CADPRODUTO PRDOUTO " & vbCrLf
                                  sSql = sSql & " Where " & vbCrLf
                                  sSql = sSql & "       LISTMAT.SGI_FILIAL  = " & FILIAL & vbCrLf
                                  sSql = sSql & "   And LISTMAT.SGI_PRODUTO = '" & Trim(BREC9!SGI_PRODLST) & "'" & vbCrLf
                                  ''sSql = sSql & "   And LISTMAT.SGI_CODPAI  = " & BREC9!SGI_CODID & vbCrLf
                                  sSql = sSql & "   And PRDOUTO.SGI_FILIAL  = LISTMAT.SGI_FILIAL      " & vbCrLf
                                  sSql = sSql & "   And PRDOUTO.SGI_CODIGO  = LISTMAT.SGI_PRODLST     " & vbCrLf
                                  
                                  BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
                                  Do While Not BREC10.EOF()
                                     uniqueID = (uniqueID + 1)
                    
                                     ReDim Preserve arrPROVARV(0 To uniqueID) As PRODARVPROD
                                     arrPROVARV(uniqueID).lngID = uniqueID
                                     arrPROVARV(uniqueID).lngIDPai = lngNivel9
                                     arrPROVARV(uniqueID).strPRODUTO = Trim(BREC10!SGI_CODIGO)
                                     arrPROVARV(uniqueID).strProdutoPAI = Trim(BREC9!SGI_PRODLST)
                                     arrPROVARV(uniqueID).lngTipo = BREC10!SGI_PRODUTOTIPO
                                     arrPROVARV(uniqueID).curQTDCONS = BREC10!SGI_QTDE
                                     arrPROVARV(uniqueID).strUNIDADE = BREC10!SGI_UNIDCONS
                                     arrPROVARV(uniqueID).intAction2Do = dacEnumUpdateAction_Ignore
                                  
                                     arrPROVARV(uniqueID).lngCODPAI = arrPROVARV(lngNivel9).lngCODIGO
                                     If Not IsNull(BREC10!SGI_CODID) Then
                                        arrPROVARV(uniqueID).lngCODIGO = BREC10!SGI_CODID
                                     Else
                                        arrPROVARV(uniqueID).lngCODIGO = objCADARVPROD.Gera_Codigo(Me.Name)
                                     End If
                                     
                                     BREC10.MoveNext
                                  Loop
                                  BREC10.Close
                                  
                                  '' Fin Nivel 9
                                  '' ==================
                                  
                                  BREC9.MoveNext
                               Loop
                               BREC9.Close
                               '' Fin Nivel 8
                               '' ==================
                               
                               BREC8.MoveNext
                            Loop
                            BREC8.Close
                            '' Fin Nivel 7
                            '' ==================
                            
                            BREC7.MoveNext
                         Loop
                         BREC7.Close
                         '' Fin Nivel 6
                         '' ==================
                         
                         BREC6.MoveNext
                      Loop
                      BREC6.Close
                      '' Fin Nivel 5
                      '' ==================
                      
                      
                      BREC5.MoveNext
                   Loop
                   BREC5.Close
                   '' Fin Nivel 4
                   '' ==================
                   
                   BREC4.MoveNext
                Loop
                BREC4.Close
                '' Fin Nivel 3
                '' ==================
                
                BREC3.MoveNext
             Loop
             BREC3.Close
             '' Fin Nivel 2
             '' ==================
             
             BREC2.MoveNext
          Loop
          BREC2.Close
          '' Fin Nivel 1
          '' ==================
          
          BREC.MoveNext
       Loop
    End If
    BREC.Close
    
    Call objCADARVPROD.GRAVA("AL")
    
End Sub

