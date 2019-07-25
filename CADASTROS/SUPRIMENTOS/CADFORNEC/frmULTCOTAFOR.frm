VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmULTCOTAFOR 
   Caption         =   "Ultimas Cotações/Pedidos"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10380
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   10380
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab stCotaPedidos 
      Height          =   3855
      Left            =   0
      TabIndex        =   7
      Top             =   1920
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   6800
      _Version        =   393216
      Style           =   1
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
      TabCaption(0)   =   "Cotações"
      TabPicture(0)   =   "frmULTCOTAFOR.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "flxGridCotMes"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "flxCotaGeral"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Pedidos"
      TabPicture(1)   =   "frmULTCOTAFOR.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "flxPedidoGeral"
      Tab(1).Control(1)=   "flxGridPedMes"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Resumo"
      TabPicture(2)   =   "frmULTCOTAFOR.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(1)=   "Frame3"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame4 
         Caption         =   "[ Pedidos no Ano ]"
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
         Height          =   1695
         Left            =   -74880
         TabIndex        =   13
         Top             =   2040
         Width           =   10095
         Begin MSFlexGridLib.MSFlexGrid flxResPedidoAnos 
            Height          =   1335
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   2355
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "[ Cotações no Ano ]"
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
         Height          =   1575
         Left            =   -74880
         TabIndex        =   12
         Top             =   360
         Width           =   10095
         Begin MSFlexGridLib.MSFlexGrid flxResCotaAnos 
            Height          =   1215
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   2143
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin MSFlexGridLib.MSFlexGrid flxCotaGeral 
         Height          =   1695
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   2990
         _Version        =   393216
         FixedCols       =   0
         HighLight       =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid flxGridCotMes 
         Height          =   1575
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   2778
         _Version        =   393216
         FixedCols       =   0
         HighLight       =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid flxGridPedMes 
         Height          =   1575
         Left            =   -74880
         TabIndex        =   9
         Top             =   360
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   2778
         _Version        =   393216
         FixedCols       =   0
         HighLight       =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid flxPedidoGeral 
         Height          =   1695
         Left            =   -74880
         TabIndex        =   11
         Top             =   2040
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   2990
         _Version        =   393216
         FixedCols       =   0
         HighLight       =   2
         SelectionMode   =   1
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "[ Dados Estatisticos ]"
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
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   10335
      Begin VB.Label Label24 
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Fornecedor"
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
         Left            =   120
         TabIndex        =   5
         Top             =   540
         Width           =   975
      End
      Begin VB.Label lblProduto 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblProduto"
         Height          =   255
         Left            =   1800
         TabIndex        =   4
         Top             =   240
         Width           =   8415
      End
      Begin VB.Label lblFornecedor 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblFornecedor"
         Height          =   255
         Left            =   1800
         TabIndex        =   3
         Top             =   540
         Width           =   8415
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10335
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
         Picture         =   "frmULTCOTAFOR.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmULTCOTAFOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho    As String
Public Linha       As Variant
Public cTipOper    As String
Public iCodigo     As Long
Public lngCODCOTA  As Long
Public FILIAL      As Integer
Public strAcesso   As String
Public strMODPAI   As String
Public strUSUARIO  As String
Public strFORNEC   As String
Public strPROD     As String
Dim objBLBFun      As Object
Dim objCADCOTA     As Object
Dim arrMES         As Variant

Private Sub cmdVoltar_Click()
    Set objBLBFun = Nothing
    Set objCADCOTA = Nothing
    Unload Me
End Sub

Private Sub flxGridCotMes_Click()
    If (flxGridCotMes.Rows - 1) > 0 Then Call PopGrdCotacoes(CLng(flxGridCotMes.TextMatrix(flxGridCotMes.Row, 1)), CLng(flxGridCotMes.TextMatrix(flxGridCotMes.Row, 2)))
End Sub

Private Sub flxGridCotMes_RowColChange()
    If (flxGridCotMes.Rows - 1) > 0 Then Call PopGrdCotacoes(CLng(flxGridCotMes.TextMatrix(flxGridCotMes.Row, 1)), CLng(flxGridCotMes.TextMatrix(flxGridCotMes.Row, 2)))
End Sub

Private Sub flxGridPedMes_Click()
    If (flxGridPedMes.Rows - 1) > 0 Then Call PopGrdPedidos(CLng(flxGridPedMes.TextMatrix(flxGridPedMes.Row, 1)), CLng(flxGridPedMes.TextMatrix(flxGridPedMes.Row, 2)))
End Sub

Private Sub flxGridPedMes_RowColChange()
    If (flxGridPedMes.Rows - 1) > 0 Then Call PopGrdPedidos(CLng(flxGridPedMes.TextMatrix(flxGridPedMes.Row, 1)), CLng(flxGridPedMes.TextMatrix(flxGridPedMes.Row, 2)))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

   Set objBLBFun = CreateObject("BLBCWS.clsFuncoes")
   Set objCADCOTA = CreateObject("CADCOTACAO.clsCADCOTACAO")
   
   objCADCOTA.FILIAL = FILIAL
   
   Set adoBanco_Dados = objBLBFun.Banco_Dados(Linha)

   If adoBanco_Dados.State = 0 Then
      MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
      Exit Sub
   End If
   
   PopCampos
   
   ReDim arrMES(1 To 12) As String
   arrMES(1) = "Janeiro"
   arrMES(2) = "Fevereiro"
   arrMES(3) = "Março"
   arrMES(4) = "Abril"
   arrMES(5) = "Maio"
   arrMES(6) = "Junho"
   arrMES(7) = "Julho"
   arrMES(8) = "Agosto"
   arrMES(9) = "Setembto"
   arrMES(10) = "Outubro"
   arrMES(11) = "Novembro"
   arrMES(11) = "Dezembro"
   
   Call ConfGirdCotMes
   Call ConfGirdPedMes
   Call ConfGridCotGeral
   Call ConfGridPedGeral
   Call ConfGridResCotAno
   Call ConfGridResPedidosAno
   
   Call PopGridCotacoesMes
   Call PopGridPedidosMes
   Call PopGridCotacoesAno
   Call PopGridPedidosAno
   
   stCotaPedidos.Tab = 0
   
End Sub

Private Sub PopCampos()

    lblProduto.Caption = ""
    lblFornecedor.Caption = ""
    
    '' Produtos
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = '" & Trim(strPROD) & "'"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then lblProduto.Caption = Trim(BREC!SGI_CODIGO) & " - " & Trim(BREC!SGI_DESCRICAO)
    BREC.Close

    '' Fornecedor
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADFORNEC " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & strFORNEC
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then lblFornecedor.Caption = Trim(BREC!SGI_CODIGO) & " - " & Trim(BREC!SGI_RAZAOSOC)
    BREC.Close

End Sub


Private Sub PopGridCotacoesMes()

    sSql = "Select " & vbCrLf
    sSql = sSql & "       Year(SGI_COTAHEADER.SGI_DATA)  As ANO " & vbCrLf
    sSql = sSql & "      ,MONTH(SGI_COTAHEADER.SGI_DATA) As MES " & vbCrLf
    sSql = sSql & "      ,AVG(SGI_COTAITENS.SGI_VLUNIT)  AS SGI_VLUNIT " & vbCrLf
    sSql = sSql & "      ,AVG(SGI_COTAITENS.SGI_PRZENTR) AS SGI_PRZENTR  " & vbCrLf
    sSql = sSql & "      ,Count(*)                       AS SGI_QTDCOTA " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_COTAITENS SGI_COTAITENS   " & vbCrLf
    sSql = sSql & "      ,SGI_COTAHEADER SGI_COTAHEADER " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    
    sSql = sSql & "       SGI_COTAITENS.SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_COTAITENS.SGI_CODFOR   = " & strFORNEC & vbCrLf
    sSql = sSql & "   And SGI_COTAITENS.SGI_PRODUTO  = '" & Trim(strPROD) & "'" & vbCrLf
    sSql = sSql & "   And SGI_COTAITENS.SGI_VLUNIT IS NOT NULL " & vbCrLf
    
    sSql = sSql & "   And SGI_COTAHEADER.SGI_FILIAL = SGI_COTAITENS.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And SGI_COTAHEADER.SGI_CODIGO = SGI_COTAITENS.SGI_CODIGO " & vbCrLf
    sSql = sSql & " Group By " & vbCrLf
    sSql = sSql & "          Year(SGI_COTAHEADER.SGI_DATA) " & vbCrLf
    sSql = sSql & "         ,MONTH (SGI_COTAHEADER.SGI_DATA) " & vbCrLf
    sSql = sSql & " Order By " & vbCrLf
    sSql = sSql & "          Year (SGI_COTAHEADER.SGI_DATA) DESC" & vbCrLf
    sSql = sSql & "         ,MONTH(SGI_COTAHEADER.SGI_DATA) DESC"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       flxGridCotMes.AddItem "" & vbTab & _
                             BREC!ANO & vbTab & _
                             BREC!MES & vbTab & _
                             arrMES(BREC!MES) & "/" & Str(BREC!ANO) & vbTab & _
                             Format(BREC!SGI_VLUNIT, "#,##0.00") & vbTab & _
                             Format(BREC!SGI_QTDCOTA, "##00") & vbTab & _
                             Format(BREC!SGI_PRZENTR, "##00")

       BREC.MoveNext
    Loop
    BREC.Close

End Sub

Private Sub ConfGirdCotMes()

    flxGridCotMes.Rows = 1
    flxGridCotMes.Cols = 7
    
    flxGridCotMes.TextMatrix(0, 0) = ""
    flxGridCotMes.TextMatrix(0, 1) = "Ano"
    flxGridCotMes.TextMatrix(0, 2) = "Mes"
    flxGridCotMes.TextMatrix(0, 3) = "Mês/Ano"
    flxGridCotMes.TextMatrix(0, 4) = "Prc.Médio"
    flxGridCotMes.TextMatrix(0, 5) = "Qtde.Cotações"
    flxGridCotMes.TextMatrix(0, 6) = "Mèdia Prz.Entr"
    
    flxGridCotMes.ColWidth(0) = 0
    flxGridCotMes.ColWidth(1) = 0
    flxGridCotMes.ColWidth(2) = 0
    
    flxGridCotMes.ColWidth(3) = 1100
    flxGridCotMes.ColWidth(4) = 1000
    flxGridCotMes.ColWidth(5) = 1200
    flxGridCotMes.ColWidth(6) = 1200
    
End Sub

Private Sub ConfGridCotGeral()

    flxCotaGeral.Rows = 1
    flxCotaGeral.Cols = 6
    
    flxCotaGeral.TextMatrix(0, 0) = ""
    flxCotaGeral.TextMatrix(0, 1) = "Cod.Cotação"
    flxCotaGeral.TextMatrix(0, 2) = "Data"
    flxCotaGeral.TextMatrix(0, 3) = "Vlr.Unitário"
    flxCotaGeral.TextMatrix(0, 4) = "Qtde"
    flxCotaGeral.TextMatrix(0, 5) = "Prz.Entrega"
    
    flxCotaGeral.ColWidth(0) = 0
    flxCotaGeral.ColWidth(1) = 1100
    flxCotaGeral.ColWidth(2) = 1000
    flxCotaGeral.ColWidth(3) = 1000
    flxCotaGeral.ColWidth(4) = 900
    flxCotaGeral.ColWidth(5) = 1000
    
End Sub


Private Sub PopGrdCotacoes(lngANO As Long, lngMES As Long)

    Call ConfGridCotGeral
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_COTAITENS.SGI_CODIGO " & vbCrLf
    sSql = sSql & "      ,SGI_COTAHEADER.SGI_DATA " & vbCrLf
    sSql = sSql & "      ,SGI_COTAITENS.SGI_VLUNIT " & vbCrLf
    sSql = sSql & "      ,SGI_COTAITENS.SGI_QTD " & vbCrLf
    sSql = sSql & "      ,SGI_COTAITENS.SGI_PRZENTR " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_COTAITENS SGI_COTAITENS " & vbCrLf
    sSql = sSql & "      ,SGI_COTAHEADER SGI_COTAHEADER " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_COTAITENS.SGI_FILIAL       = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_COTAITENS.SGI_CODFOR       = " & strFORNEC & vbCrLf
    sSql = sSql & "   And SGI_COTAITENS.SGI_PRODUTO      = '" & Trim(strPROD) & "'" & vbCrLf
    sSql = sSql & "   And SGI_COTAITENS.SGI_VLUNIT IS NOT NULL " & vbCrLf
    sSql = sSql & "   And Year(SGI_COTAHEADER.SGI_DATA)  = " & lngANO & vbCrLf
    sSql = sSql & "   And MONTH(SGI_COTAHEADER.SGI_DATA) = " & lngMES & vbCrLf
    sSql = sSql & "   And SGI_COTAHEADER.SGI_FILIAL = SGI_COTAITENS.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And SGI_COTAHEADER.SGI_CODIGO = SGI_COTAITENS.SGI_CODIGO " & vbCrLf
    sSql = sSql & " Order by " & vbCrLf
    sSql = sSql & "          SGI_COTAITENS.SGI_CODIGO DESC"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC.EOF
       flxCotaGeral.AddItem "" & vbTab & _
                            Format(BREC!SGI_CODIGO, "#/####") & vbTab & _
                            Format(BREC!SGI_DATA, "DD/MM/YYYY") & vbTab & _
                            Format(BREC!SGI_VLUNIT, "#,##0.00") & vbTab & _
                            Format(BREC!SGI_QTD, "#,###0.000") & vbTab & _
                            Format(BREC!SGI_PRZENTR, "##00")
                            
       BREC.MoveNext
    Loop
    BREC.Close
End Sub


Private Sub ConfGirdPedMes()

    flxGridPedMes.Rows = 1
    flxGridPedMes.Cols = 7
    
    flxGridPedMes.TextMatrix(0, 0) = ""
    flxGridPedMes.TextMatrix(0, 1) = "Ano"
    flxGridPedMes.TextMatrix(0, 2) = "Mes"
    flxGridPedMes.TextMatrix(0, 3) = "Mês/Ano"
    flxGridPedMes.TextMatrix(0, 4) = "Prc.Médio"
    flxGridPedMes.TextMatrix(0, 5) = "Qtde.Pedidos"
    flxGridPedMes.TextMatrix(0, 6) = "Mèdia Prz.Entr"
    
    flxGridPedMes.ColWidth(0) = 0
    flxGridPedMes.ColWidth(1) = 0
    flxGridPedMes.ColWidth(2) = 0
    
    flxGridPedMes.ColWidth(3) = 1100
    flxGridPedMes.ColWidth(4) = 1000
    flxGridPedMes.ColWidth(5) = 1200
    flxGridPedMes.ColWidth(6) = 1200
    
End Sub

Private Sub PopGridPedidosMes()

    sSql = "Select " & vbCrLf
    sSql = sSql & "       Year(SGI_PEDIDOHEADER.SGI_DATAPEDIDO)  As ANO " & vbCrLf
    sSql = sSql & "      ,MONTH(SGI_PEDIDOHEADER.SGI_DATAPEDIDO) As MES " & vbCrLf
    sSql = sSql & "      ,AVG(SGI_PEDIDOITENS.SGI_VLUNIT)        AS SGI_VLUNIT " & vbCrLf
    sSql = sSql & "      ,AVG(SGI_PEDIDOITENS.SGI_PRZENTR)       AS SGI_PRZENTR  " & vbCrLf
    sSql = sSql & "      ,Count(*)                               AS SGI_QTDCOTA " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_PEDIDOITENS SGI_PEDIDOITENS   " & vbCrLf
    sSql = sSql & "      ,SGI_PEDIDOHEADER SGI_PEDIDOHEADER " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    
    sSql = sSql & "       SGI_PEDIDOITENS.SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_PEDIDOITENS.SGI_PRODUTO  = '" & Trim(strPROD) & "'" & vbCrLf
    sSql = sSql & "   And SGI_PEDIDOITENS.SGI_VLUNIT IS NOT NULL " & vbCrLf
    
    sSql = sSql & "   And SGI_PEDIDOHEADER.SGI_CODFOR   = " & strFORNEC & vbCrLf
    sSql = sSql & "   And SGI_PEDIDOHEADER.SGI_FILIAL = SGI_PEDIDOITENS.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And SGI_PEDIDOHEADER.SGI_CODIGO = SGI_PEDIDOITENS.SGI_CODIGO " & vbCrLf
    sSql = sSql & " Group By " & vbCrLf
    sSql = sSql & "          Year(SGI_PEDIDOHEADER.SGI_DATAPEDIDO) " & vbCrLf
    sSql = sSql & "         ,MONTH (SGI_PEDIDOHEADER.SGI_DATAPEDIDO) " & vbCrLf
    sSql = sSql & " Order By " & vbCrLf
    sSql = sSql & "          Year (SGI_PEDIDOHEADER.SGI_DATAPEDIDO) DESC" & vbCrLf
    sSql = sSql & "         ,MONTH(SGI_PEDIDOHEADER.SGI_DATAPEDIDO) DESC"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       flxGridPedMes.AddItem "" & vbTab & _
                             BREC!ANO & vbTab & _
                             BREC!MES & vbTab & _
                             arrMES(BREC!MES) & "/" & Str(BREC!ANO) & vbTab & _
                             Format(BREC!SGI_VLUNIT, "#,##0.00") & vbTab & _
                             Format(BREC!SGI_QTDCOTA, "##00") & vbTab & _
                             Format(BREC!SGI_PRZENTR, "##00")

       BREC.MoveNext
    Loop
    BREC.Close

End Sub


Private Sub PopGrdPedidos(lngANO As Long, lngMES As Long)

    Call ConfGridPedGeral
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_PEDIDOITENS.SGI_CODIGO " & vbCrLf
    sSql = sSql & "      ,SGI_PEDIDOHEADER.SGI_DATAPEDIDO " & vbCrLf
    sSql = sSql & "      ,SGI_PEDIDOITENS.SGI_VLUNIT " & vbCrLf
    sSql = sSql & "      ,SGI_PEDIDOITENS.SGI_QTD " & vbCrLf
    sSql = sSql & "      ,SGI_PEDIDOITENS.SGI_PRZENTR " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_PEDIDOITENS SGI_PEDIDOITENS " & vbCrLf
    sSql = sSql & "      ,SGI_PEDIDOHEADER SGI_PEDIDOHEADER " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    
    sSql = sSql & "       SGI_PEDIDOITENS.SGI_FILIAL       = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_PEDIDOITENS.SGI_PRODUTO      = '" & Trim(strPROD) & "'" & vbCrLf
    sSql = sSql & "   And SGI_PEDIDOITENS.SGI_VLUNIT IS NOT NULL " & vbCrLf
    
    sSql = sSql & "   And SGI_PEDIDOHEADER.SGI_CODFOR       = " & strFORNEC & vbCrLf
    sSql = sSql & "   And Year(SGI_PEDIDOHEADER.SGI_DATAPEDIDO)  = " & lngANO & vbCrLf
    sSql = sSql & "   And MONTH(SGI_PEDIDOHEADER.SGI_DATAPEDIDO) = " & lngMES & vbCrLf
    sSql = sSql & "   And SGI_PEDIDOHEADER.SGI_FILIAL = SGI_PEDIDOITENS.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And SGI_PEDIDOHEADER.SGI_CODIGO = SGI_PEDIDOITENS.SGI_CODIGO " & vbCrLf
    sSql = sSql & " Order by " & vbCrLf
    sSql = sSql & "          SGI_PEDIDOITENS.SGI_CODIGO DESC"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC.EOF
       flxPedidoGeral.AddItem "" & vbTab & _
                               Format(BREC!SGI_CODIGO, "#/####") & vbTab & _
                               Format(BREC!SGI_DATAPEDIDO, "DD/MM/YYYY") & vbTab & _
                               Format(BREC!SGI_VLUNIT, "#,##0.00") & vbTab & _
                               Format(BREC!SGI_QTD, "#,###0.000") & vbTab & _
                               Format(BREC!SGI_PRZENTR, "##00")
                            
       BREC.MoveNext
    Loop
    BREC.Close
    
End Sub


Private Sub ConfGridPedGeral()

    flxPedidoGeral.Rows = 1
    flxPedidoGeral.Cols = 6
    
    flxPedidoGeral.TextMatrix(0, 0) = ""
    flxPedidoGeral.TextMatrix(0, 1) = "Cod.Pedido"
    flxPedidoGeral.TextMatrix(0, 2) = "Data"
    flxPedidoGeral.TextMatrix(0, 3) = "Vlr.Unitário"
    flxPedidoGeral.TextMatrix(0, 4) = "Qtde"
    flxPedidoGeral.TextMatrix(0, 5) = "Prz.Entrega"
    
    flxPedidoGeral.ColWidth(0) = 0
    flxPedidoGeral.ColWidth(1) = 1100
    flxPedidoGeral.ColWidth(2) = 1000
    flxPedidoGeral.ColWidth(3) = 1000
    flxPedidoGeral.ColWidth(4) = 900
    flxPedidoGeral.ColWidth(5) = 1000
    
End Sub

Private Sub ConfGridResCotAno()

    flxResCotaAnos.Rows = 1
    flxResCotaAnos.Cols = 7
    
    flxResCotaAnos.TextMatrix(0, 0) = ""
    flxResCotaAnos.TextMatrix(0, 1) = "Ano"
    flxResCotaAnos.TextMatrix(0, 2) = "Mes"
    flxResCotaAnos.TextMatrix(0, 3) = "Ano"
    flxResCotaAnos.TextMatrix(0, 4) = "Prc.Médio"
    flxResCotaAnos.TextMatrix(0, 5) = "Qtde.Cotações"
    flxResCotaAnos.TextMatrix(0, 6) = "Mèdia Prz.Entr"
    
    flxResCotaAnos.ColWidth(0) = 0
    flxResCotaAnos.ColWidth(1) = 0
    flxResCotaAnos.ColWidth(2) = 0
    
    flxResCotaAnos.ColWidth(3) = 1100
    flxResCotaAnos.ColWidth(4) = 1000
    flxResCotaAnos.ColWidth(5) = 1200
    flxResCotaAnos.ColWidth(6) = 1200
    
End Sub

Private Sub ConfGridResPedidosAno()

    flxResPedidoAnos.Rows = 1
    flxResPedidoAnos.Cols = 7
    
    flxResPedidoAnos.TextMatrix(0, 0) = ""
    flxResPedidoAnos.TextMatrix(0, 1) = "Ano"
    flxResPedidoAnos.TextMatrix(0, 2) = "Mes"
    flxResPedidoAnos.TextMatrix(0, 3) = "Ano"
    flxResPedidoAnos.TextMatrix(0, 4) = "Prc.Médio"
    flxResPedidoAnos.TextMatrix(0, 5) = "Qtde.Pedidos"
    flxResPedidoAnos.TextMatrix(0, 6) = "Mèdia Prz.Entr"
    
    flxResPedidoAnos.ColWidth(0) = 0
    flxResPedidoAnos.ColWidth(1) = 0
    flxResPedidoAnos.ColWidth(2) = 0
    
    flxResPedidoAnos.ColWidth(3) = 1100
    flxResPedidoAnos.ColWidth(4) = 1000
    flxResPedidoAnos.ColWidth(5) = 1200
    flxResPedidoAnos.ColWidth(6) = 1200
    
End Sub

Private Sub PopGridCotacoesAno()

    sSql = "Select " & vbCrLf
    sSql = sSql & "       Year(SGI_COTAHEADER.SGI_DATA)  As ANO " & vbCrLf
    sSql = sSql & "      ,AVG(SGI_COTAITENS.SGI_VLUNIT)  AS SGI_VLUNIT " & vbCrLf
    sSql = sSql & "      ,AVG(SGI_COTAITENS.SGI_PRZENTR) AS SGI_PRZENTR  " & vbCrLf
    sSql = sSql & "      ,Count(*)                       AS SGI_QTDCOTA " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_COTAITENS SGI_COTAITENS   " & vbCrLf
    sSql = sSql & "      ,SGI_COTAHEADER SGI_COTAHEADER " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    
    sSql = sSql & "       SGI_COTAITENS.SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_COTAITENS.SGI_CODFOR   = " & strFORNEC & vbCrLf
    sSql = sSql & "   And SGI_COTAITENS.SGI_PRODUTO  = '" & Trim(strPROD) & "'" & vbCrLf
    sSql = sSql & "   And SGI_COTAITENS.SGI_VLUNIT IS NOT NULL " & vbCrLf
    
    sSql = sSql & "   And SGI_COTAHEADER.SGI_FILIAL = SGI_COTAITENS.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And SGI_COTAHEADER.SGI_CODIGO = SGI_COTAITENS.SGI_CODIGO " & vbCrLf
    sSql = sSql & " Group By " & vbCrLf
    sSql = sSql & "          Year(SGI_COTAHEADER.SGI_DATA) " & vbCrLf
    sSql = sSql & " Order By " & vbCrLf
    sSql = sSql & "          Year (SGI_COTAHEADER.SGI_DATA) DESC" & vbCrLf
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       flxResCotaAnos.AddItem "" & vbTab & _
                             BREC!ANO & vbTab & _
                             "" & vbTab & _
                             Str(BREC!ANO) & vbTab & _
                             Format(BREC!SGI_VLUNIT, "#,##0.00") & vbTab & _
                             Format(BREC!SGI_QTDCOTA, "##00") & vbTab & _
                             Format(BREC!SGI_PRZENTR, "##00")

       BREC.MoveNext
    Loop
    BREC.Close

End Sub


Private Sub PopGridPedidosAno()

    sSql = "Select " & vbCrLf
    sSql = sSql & "       Year(SGI_PEDIDOHEADER.SGI_DATAPEDIDO)  As ANO " & vbCrLf
    sSql = sSql & "      ,AVG(SGI_PEDIDOITENS.SGI_VLUNIT)        AS SGI_VLUNIT " & vbCrLf
    sSql = sSql & "      ,AVG(SGI_PEDIDOITENS.SGI_PRZENTR)       AS SGI_PRZENTR  " & vbCrLf
    sSql = sSql & "      ,Count(*)                               AS SGI_QTDCOTA " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_PEDIDOITENS SGI_PEDIDOITENS   " & vbCrLf
    sSql = sSql & "      ,SGI_PEDIDOHEADER SGI_PEDIDOHEADER " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    
    sSql = sSql & "       SGI_PEDIDOITENS.SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_PEDIDOITENS.SGI_PRODUTO  = '" & Trim(strPROD) & "'" & vbCrLf
    sSql = sSql & "   And SGI_PEDIDOITENS.SGI_VLUNIT IS NOT NULL " & vbCrLf
    
    sSql = sSql & "   And SGI_PEDIDOHEADER.SGI_CODFOR   = " & strFORNEC & vbCrLf
    sSql = sSql & "   And SGI_PEDIDOHEADER.SGI_FILIAL = SGI_PEDIDOITENS.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And SGI_PEDIDOHEADER.SGI_CODIGO = SGI_PEDIDOITENS.SGI_CODIGO " & vbCrLf
    sSql = sSql & " Group By " & vbCrLf
    sSql = sSql & "          Year(SGI_PEDIDOHEADER.SGI_DATAPEDIDO) " & vbCrLf
    sSql = sSql & " Order By " & vbCrLf
    sSql = sSql & "          Year (SGI_PEDIDOHEADER.SGI_DATAPEDIDO) DESC" & vbCrLf
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       flxResPedidoAnos.AddItem "" & vbTab & _
                             BREC!ANO & vbTab & _
                             "" & vbTab & _
                             Str(BREC!ANO) & vbTab & _
                             Format(BREC!SGI_VLUNIT, "#,##0.00") & vbTab & _
                             Format(BREC!SGI_QTDCOTA, "##00") & vbTab & _
                             Format(BREC!SGI_PRZENTR, "##00")

       BREC.MoveNext
    Loop
    BREC.Close

End Sub


