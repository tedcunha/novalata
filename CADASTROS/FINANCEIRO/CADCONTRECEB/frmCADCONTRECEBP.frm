VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCADCONTRECEBP 
   Caption         =   "Cadastro de Contas a Receber Manual"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   9855
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   4455
      Left            =   0
      TabIndex        =   14
      Top             =   1560
      Width           =   9855
      Begin TabDlg.SSTab StTitulos 
         Height          =   4095
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   7223
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
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
         TabCaption(0)   =   "Titulos em Aberto"
         TabPicture(0)   =   "frmCADCONTRECEBP.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "flxContasaReceb"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Titulos Baixados"
         TabPicture(1)   =   "frmCADCONTRECEBP.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "flxTitBaixados"
         Tab(1).ControlCount=   1
         Begin MSFlexGridLib.MSFlexGrid flxTitBaixados 
            Height          =   3615
            Left            =   -74880
            TabIndex        =   17
            Top             =   360
            Width           =   9375
            _ExtentX        =   16536
            _ExtentY        =   6376
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin MSFlexGridLib.MSFlexGrid flxContasaReceb 
            Height          =   3615
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   9375
            _ExtentX        =   16536
            _ExtentY        =   6376
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            SelectionMode   =   1
            Appearance      =   0
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9855
      Begin VB.ComboBox cboFiltro 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   200
         Width           =   1695
      End
      Begin VB.TextBox txtCampos 
         Height          =   285
         Left            =   4440
         MaxLength       =   50
         TabIndex        =   10
         Text            =   "txtCampos"
         Top             =   200
         Width           =   5295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filtro:"
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
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Campo:"
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
         Left            =   3720
         TabIndex        =   12
         Top             =   240
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   9855
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   5280
         Top             =   240
      End
      Begin VB.CommandButton cmdVoltar 
         Caption         =   "&Voltar"
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
         Left            =   120
         Picture         =   "frmCADCONTRECEBP.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Volta ao Menu Principal"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdInclui 
         Caption         =   "&Inclui"
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
         Left            =   960
         Picture         =   "frmCADCONTRECEBP.frx":056A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Inclui uma nova condição de pagamento"
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
         Height          =   735
         Left            =   1800
         Picture         =   "frmCADCONTRECEBP.frx":0A9C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Altera a condição de pagamento"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdExclui 
         Caption         =   "&Exclui"
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
         Left            =   2640
         Picture         =   "frmCADCONTRECEBP.frx":0B9E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Exclui a condição de pagamento"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdCanFiltro 
         Caption         =   "Desfas"
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
         Left            =   8040
         Picture         =   "frmCADCONTRECEBP.frx":0CA0
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdOrden 
         Caption         =   "Ordem"
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
         Left            =   8880
         Picture         =   "frmCADCONTRECEBP.frx":11D2
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdBaixa 
         Caption         =   "&Baixa"
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
         Left            =   3480
         Picture         =   "frmCADCONTRECEBP.frx":12D4
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Exclui a condição de pagamento"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdExtorna 
         Caption         =   "E&xtorna"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4320
         Picture         =   "frmCADCONTRECEBP.frx":1716
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Exclui Empresa"
         Top             =   120
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADCONTRECEBP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho     As String
Public Linha        As Variant
Public FILIAL       As Integer
Public strAcesso    As String
Public strUSUARIO   As String
Dim objFuncoes      As Object
Dim objCADCONTRECEB As Object
Dim iCodigo         As Integer

Private Sub cmdAltera_Click()
    
    If objFuncoes.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    If StTitulos.Tab = 0 Then
       If VerifBaixados = False Then Exit Sub
       Operacao "A"
    End If
    
    If StTitulos.Tab = 1 Then
       If VerificaCaixa = True Then Exit Sub
       Operacao "AB"
    End If
    
End Sub

Private Sub cmdBaixa_Click()
    If StTitulos.Tab = 1 Then Exit Sub
    Operacao "B"
End Sub

Private Sub cmdCanFiltro_Click()

    AbilitaCampos
    ConfGrid
    ConfGridBaixados
    PreencheGrid
    PopGridBaixados
    
    cboFiltro.Clear
    cboFiltro.AddItem "Nº Doc-NF"
    cboFiltro.AddItem "Cliente"
    cboFiltro.AddItem "Dt. Vencto"
    
    txtCampos.Text = ""
    cboFiltro.ListIndex = 0
    StTitulos.Tab = 0

End Sub

Private Sub cmdExclui_Click()

  If objFuncoes.ChecaAcesso2("E", strAcesso) = False Then Exit Sub
  
  If Verif_reg = True Then Exit Sub
  
  Dim iResp As Integer
  
  If StTitulos.Tab = 1 Then Exit Sub
  If VerifBaixados = False Then Exit Sub
  
  iResp = MsgBox("Confirma a exclusão do registro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
  If iResp <> 6 Then Exit Sub
  
  If objCADCONTRECEB.GRAVA("E") = False Then Exit Sub
  
  MsgBox "Registro excluso com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
  
  Atualiza_Grid
  AbilitaCampos

End Sub

Private Sub cmdExtorna_Click()

  Dim iResp As Integer
  
  If StTitulos.Tab = 0 Then Exit Sub
  
  If VerificaCaixa = True Then Exit Sub
  
  If (flxTitBaixados.Rows - 1) = 0 Then Exit Sub
  
  iResp = MsgBox("Confirma o extorno do titulo ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
  If iResp <> 6 Then Exit Sub
  
  objCADCONTRECEB.PARCPGTO = CInt(flxTitBaixados.TextMatrix(flxTitBaixados.Row, 9))
  objCADCONTRECEB.NUMDOC = flxTitBaixados.TextMatrix(flxTitBaixados.Row, 2)
  objCADCONTRECEB.FILIAL = FILIAL
  If objCADCONTRECEB.GRAVA("X") = False Then Exit Sub
  If objCADCONTRECEB.Atualiza("I", Str(objCADCONTRECEB.CODPGTO), FILIAL, "frmCADCONTRECEB") = False Then Exit Sub
  
  MsgBox "Registro extornado com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
  
  Atualiza_Grid
  AbilitaCampos

End Sub

Private Sub cmdInclui_Click()
    If objFuncoes.ChecaAcesso2("I", strAcesso) = False Then Exit Sub
    Operacao "I"
End Sub

Private Sub cmdOrden_Click()
    Call Ordem
End Sub

Private Sub cmdVoltar_Click()
    Set objFuncoes = Nothing
    Set objCADCONTRECEB = Nothing
    Unload Me
End Sub

Private Sub flxContasaReceb_Click()
    If flxContasaReceb.Rows > 1 Then objCADCONTRECEB.CODPGTO = CInt(flxContasaReceb.TextMatrix(flxContasaReceb.RowSel, 1))
End Sub

Private Sub flxContasaReceb_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    If flxContasaReceb.Rows > 1 Then Operacao "C"
End Sub

Private Sub flxContasaReceb_RowColChange()
    If flxContasaReceb.Rows > 1 Then objCADCONTRECEB.CODPGTO = CInt(flxContasaReceb.TextMatrix(flxContasaReceb.RowSel, 1))
End Sub

Private Sub flxTitBaixados_Click()
    
    If flxTitBaixados.Rows > 1 Then objCADCONTRECEB.CODPGTO = CInt(flxTitBaixados.TextMatrix(flxTitBaixados.RowSel, 1))
End Sub

Private Sub flxTitBaixados_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    If flxTitBaixados.Rows > 1 Then Operacao "CB"
End Sub

Private Sub flxTitBaixados_RowColChange()
    If flxTitBaixados.Rows > 1 Then objCADCONTRECEB.CODPGTO = CInt(flxTitBaixados.TextMatrix(flxTitBaixados.RowSel, 1))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
    Set objCADCONTRECEB = CreateObject("CADCONTRECEB.clsCADCONTRECEB")
    
    objCADCONTRECEB.FILIAL = FILIAL
    
    objFuncoes.LimpaCampos frmCADCONTRECEBP
    
    Set adoBanco_Dados = objFuncoes.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    AbilitaCampos
    ConfGrid
    ConfGridBaixados
    PreencheGrid
    PopGridBaixados
    
    cboFiltro.AddItem "Nº Doc-NF"
    cboFiltro.AddItem "Cliente"
    cboFiltro.AddItem "Dt. Vencto"
    
    cboFiltro.ListIndex = 0
    StTitulos.Tab = 0
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)

End Sub

Private Sub AbilitaCampos()
    
    If objCADCONTRECEB.Pesq_CadContasARC = False Then
       cmdAltera.Enabled = False
       cmdExclui.Enabled = False
       cmdBaixa.Enabled = False
       cmdExtorna.Enabled = False
       Frame1.Enabled = False
       Frame3.Enabled = False
    Else
       cmdAltera.Enabled = True
       cmdExclui.Enabled = True
       cmdBaixa.Enabled = True
       cmdExtorna.Enabled = True
       Frame1.Enabled = True
       Frame3.Enabled = True
    End If

End Sub

Private Sub ConfGrid()
        
    flxContasaReceb.Rows = 1
    flxContasaReceb.Cols = 9
    
    flxContasaReceb.TextMatrix(0, 0) = ""
    flxContasaReceb.TextMatrix(0, 1) = "Código"
    flxContasaReceb.TextMatrix(0, 2) = "Nº Doc."
    flxContasaReceb.TextMatrix(0, 3) = "Vencto"
    flxContasaReceb.TextMatrix(0, 4) = "Parcela"
    flxContasaReceb.TextMatrix(0, 5) = "Valor"
    flxContasaReceb.TextMatrix(0, 6) = "Cod. Cliente"
    flxContasaReceb.TextMatrix(0, 7) = "Razão Social"
    flxContasaReceb.TextMatrix(0, 8) = "Parcela"
    
    flxContasaReceb.ColWidth(0) = 0
    flxContasaReceb.ColWidth(1) = 1000
    flxContasaReceb.ColWidth(2) = 1000
    flxContasaReceb.ColWidth(3) = 1000
    flxContasaReceb.ColWidth(4) = 1000
    flxContasaReceb.ColWidth(5) = 1000
    flxContasaReceb.ColWidth(6) = 1000
    flxContasaReceb.ColWidth(7) = 5000
    flxContasaReceb.ColWidth(8) = 0
    
End Sub

Private Sub PreencheGrid()

    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
    sSql = sSql & "     ,CABEC.SGI_CODCLI " & vbCrLf
    sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
    sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_CONTASIARC ITENS" & vbCrLf
    sSql = sSql & "     ,SGI_CONTASHARC CABEC" & vbCrLf
    sSql = sSql & "     ,SGI_CADCLIENTE FORNE" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "  And ITENS.SGI_VLPAGO IS NULL" & vbCrLf
    sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
    sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
    sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
    sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODCLI " & vbCrLf
    sSql = sSql & "Order By" & vbCrLf
    sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
    sSql = sSql & "       ,CABEC.SGI_CODCLI" & vbCrLf
    sSql = sSql & "       ,ITENS.SGI_PARCELA"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       flxContasaReceb.AddItem "" & vbTab & _
                               BREC!SGI_CODIGO & vbTab & _
                               BREC!SGI_NUMDOC & vbTab & _
                               Format(BREC!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                               Format(BREC!SGI_PARCELA, "##00") & "/" & Format(BREC!SGI_QTDPARC, "##00") & vbTab & _
                               Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                               BREC!SGI_CODCLI & vbTab & _
                               BREC!SGI_RAZAOSOC & vbTab & _
                               BREC!SGI_PARCELA
                               
       BREC.MoveNext
    Loop
    
    PosGrid
    
    BREC.Close
    
End Sub
Private Sub PosGrid()

    If iCodigo > 0 Then
        Dim I As Integer
               
        For I = 1 To (flxContasaReceb.Rows - 1)
            
            If flxContasaReceb.TextMatrix(I, 1) = iCodigo Then
               flxContasaReceb.Row = I
               flxContasaReceb.Col = 1
               
               Exit For
            End If
         
        Next I
    End If

End Sub

Private Sub Operacao(strOperacao As String)
 
  Dim Pesquisa As String
  
  If StTitulos.Tab = 0 Then
     If flxContasaReceb.Rows > 1 Then iCodigo = CInt(flxContasaReceb.TextMatrix(flxContasaReceb.RowSel, 1))
  End If
  If StTitulos.Tab = 1 Then
     If flxTitBaixados.Rows > 1 Then iCodigo = CInt(flxTitBaixados.TextMatrix(flxTitBaixados.RowSel, 1))
  End If
    
  frmCADCONTRECEB.cCaminho = cCaminho
  frmCADCONTRECEB.Linha = Linha
  frmCADCONTRECEB.iCodigo = iCodigo
  frmCADCONTRECEB.cTipOper = strOperacao
  frmCADCONTRECEB.FILIAL = FILIAL
  frmCADCONTRECEB.strAcesso = strAcesso
  frmCADCONTRECEB.strMODPAI = Me.Name
  frmCADCONTRECEB.strUSUARIO = strUSUARIO
  
  If StTitulos.Tab = 0 Then
     If flxContasaReceb.Rows > 1 Then frmCADCONTRECEB.iParcela = CInt(flxContasaReceb.TextMatrix(flxContasaReceb.RowSel, 8))
  End If
  If StTitulos.Tab = 1 Then
     If flxTitBaixados.Rows > 1 Then frmCADCONTRECEB.iParcela = CInt(flxTitBaixados.TextMatrix(flxTitBaixados.RowSel, 9))
  End If
  
  frmCADCONTRECEB.Show vbModal
  
  Atualiza_Grid
  AbilitaCampos

End Sub

Private Function VerifBaixados() As Boolean

    VerifBaixados = False
    
    '' Verifica se há baixados
    sSql = "Select" & vbCrLf
    sSql = sSql & "      * " & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_CONTASIARC " & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "  And SGI_CODIGO = " & objCADCONTRECEB.CODPGTO & vbCrLf
    sSql = sSql & "  And SGI_VLPAGO IS NOT NULL "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
       MsgBox "Há titulos já baixado !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Function
    End If
    BREC.Close
    '' ------------------------------
    
    VerifBaixados = True

End Function

Public Sub ConfGridBaixados()

    flxTitBaixados.Rows = 1
    flxTitBaixados.Cols = 10
    
    flxTitBaixados.TextMatrix(0, 0) = ""
    flxTitBaixados.TextMatrix(0, 1) = "Código"
    flxTitBaixados.TextMatrix(0, 2) = "Nº Doc."
    flxTitBaixados.TextMatrix(0, 3) = "Vencto"
    flxTitBaixados.TextMatrix(0, 4) = "Pgto."
    flxTitBaixados.TextMatrix(0, 5) = "Parcela"
    flxTitBaixados.TextMatrix(0, 6) = "Valor"
    flxTitBaixados.TextMatrix(0, 7) = "Cod. Cli"
    flxTitBaixados.TextMatrix(0, 8) = "Razão Social"
    flxTitBaixados.TextMatrix(0, 9) = "Parcela"
    
    flxTitBaixados.ColWidth(0) = 0
    flxTitBaixados.ColWidth(1) = 1000
    flxTitBaixados.ColWidth(2) = 1000
    flxTitBaixados.ColWidth(3) = 1000
    flxTitBaixados.ColWidth(4) = 1000
    flxTitBaixados.ColWidth(5) = 1000
    flxTitBaixados.ColWidth(6) = 1000
    flxTitBaixados.ColWidth(7) = 1000
    flxTitBaixados.ColWidth(8) = 5000
    flxTitBaixados.ColWidth(9) = 0

End Sub

Private Sub PopGridBaixados()

    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
    sSql = sSql & "     ,CABEC.SGI_CODCLI " & vbCrLf
    sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_DTPGTO" & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
    sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_CONTASIARC ITENS" & vbCrLf
    sSql = sSql & "     ,SGI_CONTASHARC CABEC" & vbCrLf
    sSql = sSql & "     ,SGI_CADCLIENTE FORNE" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "  And ITENS.SGI_VLPAGO IS NOT NULL" & vbCrLf
    sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
    sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
    sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
    sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODCLI " & vbCrLf
    sSql = sSql & "Order By" & vbCrLf
    sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
    sSql = sSql & "       ,CABEC.SGI_CODCLI" & vbCrLf
    sSql = sSql & "       ,ITENS.SGI_PARCELA"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       flxTitBaixados.AddItem "" & vbTab & _
                              BREC!SGI_CODIGO & vbTab & _
                              BREC!SGI_NUMDOC & vbTab & _
                              Format(BREC!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                              Format(BREC!SGI_DTPGTO, "DD/MM/YYYY") & vbTab & _
                              Format(BREC!SGI_PARCELA, "##00") & "/" & Format(BREC!SGI_QTDPARC, "##00") & vbTab & _
                              Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                              BREC!SGI_CODCLI & vbTab & _
                              BREC!SGI_RAZAOSOC & vbTab & _
                              BREC!SGI_PARCELA
                               
       BREC.MoveNext
    Loop
    
    PosGridBaixado
    
    BREC.Close

End Sub


Private Sub PosGridBaixado()

    If iCodigo > 0 Then
        Dim I As Integer
               
        For I = 1 To (flxTitBaixados.Rows - 1)
            
            If flxTitBaixados.TextMatrix(I, 1) = iCodigo Then
               flxTitBaixados.Row = I
               flxTitBaixados.Col = 1
               
               Exit For
            End If
         
        Next I
    End If

End Sub


Private Function VerificaCaixa() As Boolean
    
    VerificaCaixa = False
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CONTASIARC " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & objCADCONTRECEB.CODPGTO & vbCrLf
    sSql = sSql & "   And SGI_CADCAIXA IS NOT NULL"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
       MsgBox "Já foi criado fluxo de caixa !!!", vbOKOnly + vbExclamation, "Aviso"
       VerificaCaixa = True
    End If
    BREC.Close
    
    
End Function

Private Sub StTitulos_Click(PreviousTab As Integer)

    cboFiltro.Clear
    If StTitulos.Tab = 0 Then
       cboFiltro.AddItem "Nº Doc-NF"
       cboFiltro.AddItem "Cliente"
       cboFiltro.AddItem "Dt. Vencto"
    ElseIf StTitulos.Tab = 1 Then
       cboFiltro.AddItem "Nº Doc-NF"
       cboFiltro.AddItem "Cliente"
       cboFiltro.AddItem "Dt. Vencto"
       cboFiltro.AddItem "Dt. Pgto."
    End If
    cboFiltro.ListIndex = 0

End Sub

Private Sub Timer1_Timer()
    AbilitaCampos
    Atualiza_Grid
End Sub


Private Sub Atualiza_Grid()
    
     Dim I        As Integer
     Dim bolAchou As Boolean
      
      bolAchou = False
      
     sSql = "Select" & vbCrLf
     sSql = sSql & "      * " & vbCrLf
     sSql = sSql & "  From" & vbCrLf
     sSql = sSql & "       SGI_ATUALIZA" & vbCrLf
     sSql = sSql & " Where" & vbCrLf
     sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & "   And SGI_MODULO = 'frmCADCONTRECEB'" & vbCrLf

     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     If Not BREC.EOF Then
        If StTitulos.Tab = 0 Then
           For I = 1 To (flxContasaReceb.Rows - 1)
               If Trim(BREC!SGI_ACAO) = "E" Or Trim(BREC!SGI_ACAO) = "B" Then
                  If flxContasaReceb.TextMatrix(I, 1) = Trim(BREC!SGI_CODIGO) Then
                     If flxContasaReceb.Rows = 2 Then flxContasaReceb.Rows = 1
                     If flxContasaReceb.Rows > 2 Then flxContasaReceb.RemoveItem I
                     Exit For
                  End If
               ElseIf Trim(BREC!SGI_ACAO) = "I" Or Trim(BREC!SGI_ACAO) = "A" Then
                  If Trim(BREC!SGI_CODIGO) = Trim(flxContasaReceb.TextMatrix(I, 1)) Then
                     bolAchou = True
                     Exit For
                  End If
               End If
           Next I
        ElseIf StTitulos.Tab = 1 Then
           For I = 1 To (flxTitBaixados.Rows - 1)
               If Trim(BREC!SGI_ACAO) = "E" Or Trim(BREC!SGI_ACAO) = "I" Then
                  If flxTitBaixados.TextMatrix(I, 1) = Trim(BREC!SGI_CODIGO) Then
                     If flxTitBaixados.Rows = 2 Then flxTitBaixados.Rows = 1
                     If flxTitBaixados.Rows > 2 Then flxTitBaixados.RemoveItem I
                     Exit For
                  End If
               ElseIf Trim(BREC!SGI_ACAO) = "B" Or Trim(BREC!SGI_ACAO) = "A" Then
                  If Trim(BREC!SGI_CODIGO) = Trim(flxTitBaixados.TextMatrix(I, 1)) Then
                     bolAchou = True
                     Exit For
                  End If
               End If
           Next I
        End If
        
        If StTitulos.Tab = 0 Then
        
           If bolAchou = False And Trim(BREC!SGI_ACAO) = "I" Then
           
              sSql = "Select" & vbCrLf
              sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
              sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
              sSql = sSql & "     ,CABEC.SGI_CODCLI " & vbCrLf
              sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
              sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
              sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
              sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
              sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
              sSql = sSql & "  From" & vbCrLf
              sSql = sSql & "      SGI_CONTASIARC ITENS" & vbCrLf
              sSql = sSql & "     ,SGI_CONTASHARC CABEC" & vbCrLf
              sSql = sSql & "     ,SGI_CADCLIENTE FORNE" & vbCrLf
              sSql = sSql & " Where" & vbCrLf
              sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
              sSql = sSql & "  And ITENS.SGI_VLPAGO IS NULL" & vbCrLf
              sSql = sSql & "  And ITENS.SGI_CODIGO = " & Trim(BREC!SGI_CODIGO) & vbCrLf
              sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
              sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
              sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
              sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODCLI " & vbCrLf
              sSql = sSql & "Order By" & vbCrLf
              sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
              sSql = sSql & "       ,CABEC.SGI_CODCLI" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_PARCELA"
               
              BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
              If Not BREC2.EOF Then
              
                 flxContasaReceb.AddItem "" & vbTab & _
                                   BREC2!SGI_CODIGO & vbTab & _
                                   BREC2!SGI_NUMDOC & vbTab & _
                                   Format(BREC2!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                                   Format(BREC2!SGI_PARCELA, "##00") & "/" & Format(BREC2!SGI_QTDPARC, "##00") & vbTab & _
                                   Format(BREC2!SGI_VLDOC, "#,##0.00") & vbTab & _
                                   BREC2!SGI_CODCLI & vbTab & _
                                   BREC2!SGI_RAZAOSOC & vbTab & _
                                   BREC2!SGI_PARCELA
              End If
              BREC2.Close
              
           ElseIf bolAchou = False And Trim(BREC!SGI_ACAO) = "B" Then
               BREC.Close
               Exit Sub
           ElseIf bolAchou = True And BREC!SGI_ACAO = "A" Then
                
              sSql = "Select" & vbCrLf
              sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
              sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
              sSql = sSql & "     ,CABEC.SGI_CODCLI " & vbCrLf
              sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
              sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
              sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
              sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
              sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
              sSql = sSql & "  From" & vbCrLf
              sSql = sSql & "      SGI_CONTASIARC ITENS" & vbCrLf
              sSql = sSql & "     ,SGI_CONTASHARC CABEC" & vbCrLf
              sSql = sSql & "     ,SGI_CADCLIENTE FORNE" & vbCrLf
              sSql = sSql & " Where" & vbCrLf
              sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
              sSql = sSql & "  And ITENS.SGI_VLPAGO IS NULL" & vbCrLf
              sSql = sSql & "  And ITENS.SGI_CODIGO = " & Trim(BREC!SGI_CODIGO)
              sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
              sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
              sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
              sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODCLI " & vbCrLf
              sSql = sSql & "Order By" & vbCrLf
              sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
              sSql = sSql & "       ,CABEC.SGI_CODCLI" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_PARCELA"
                
              BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
              If Not BREC2.EOF Then
                 flxContasaReceb.TextMatrix(I, 0) = ""
                 flxContasaReceb.TextMatrix(I, 1) = BREC2!SGI_CODIGO
                 flxContasaReceb.TextMatrix(I, 2) = BREC2!SGI_NUMDOC
                 flxContasaReceb.TextMatrix(I, 3) = Format(BREC2!SGI_DATAVENC, "DD/MM/YYYY")
                 flxContasaReceb.TextMatrix(I, 4) = Format(BREC2!SGI_PARCELA, "##00") & "/" & Format(BREC2!SGI_QTDPARC, "##00")
                 flxContasaReceb.TextMatrix(I, 5) = Format(BREC2!SGI_VLDOC, "#,##0.00")
                 flxContasaReceb.TextMatrix(I, 6) = BREC2!SGI_CODCLI
                 flxContasaReceb.TextMatrix(I, 7) = BREC2!SGI_RAZAOSOC
                 flxContasaReceb.TextMatrix(I, 8) = BREC2!SGI_PARCELA
              End If
              BREC2.Close
           
           End If
            
        ElseIf StTitulos.Tab = 1 Then
        
           If bolAchou = False And BREC!SGI_ACAO = "B" Then
           
               sSql = "Select" & vbCrLf
               sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
               sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
               sSql = sSql & "     ,CABEC.SGI_CODCLI " & vbCrLf
               sSql = sSql & "     ,FORNE.SGI_RAZAOSOC " & vbCrLf
               sSql = sSql & "     ,ITENS.SGI_DATAVENC " & vbCrLf
               sSql = sSql & "     ,ITENS.SGI_DTPGTO " & vbCrLf
               sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
               sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
               sSql = sSql & "     ,ITENS.SGI_VLDOC " & vbCrLf
               sSql = sSql & "  From" & vbCrLf
               sSql = sSql & "      SGI_CONTASIARC ITENS" & vbCrLf
               sSql = sSql & "     ,SGI_CONTASHARC CABEC" & vbCrLf
               sSql = sSql & "     ,SGI_CADCLIENTE FORNE" & vbCrLf
               sSql = sSql & " Where" & vbCrLf
               sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
               sSql = sSql & "  And ITENS.SGI_VLPAGO IS NOT NULL" & vbCrLf
               sSql = sSql & "  And ITENS.SGI_CODIGO = " & Trim(BREC!SGI_CODIGO)
               sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
               sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
               sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
               sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODCLI " & vbCrLf
               sSql = sSql & "Order By" & vbCrLf
               sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
               sSql = sSql & "       ,CABEC.SGI_CODCLI" & vbCrLf
               sSql = sSql & "       ,ITENS.SGI_PARCELA"
           
               BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
               If Not BREC2.EOF Then
                  flxTitBaixados.AddItem "" & vbTab & _
                                         BREC2!SGI_CODIGO & vbTab & _
                                         BREC2!SGI_NUMDOC & vbTab & _
                                         Format(BREC2!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                                         Format(BREC2!SGI_DTPGTO, "DD/MM/YYYY") & vbTab & _
                                         Format(BREC2!SGI_PARCELA, "##00") & "/" & Format(BREC2!SGI_QTDPARC, "##00") & vbTab & _
                                         Format(BREC2!SGI_VLDOC, "#,##0.00") & vbTab & _
                                         BREC2!SGI_CODCLI & vbTab & _
                                         BREC2!SGI_RAZAOSOC & vbTab & _
                                         BREC2!SGI_PARCELA
               End If
               BREC2.Close
           
           ElseIf bolAchou = False And BREC!SGI_ACAO = "I" Then
              BREC.Close
              Exit Sub
           
           ElseIf bolAchou = True And BREC!SGI_ACAO = "B" Then
           
              sSql = "Select" & vbCrLf
              sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
              sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
              sSql = sSql & "     ,CABEC.SGI_CODCLI " & vbCrLf
              sSql = sSql & "     ,FORNE.SGI_RAZAOSOC " & vbCrLf
              sSql = sSql & "     ,ITENS.SGI_DATAVENC " & vbCrLf
              sSql = sSql & "     ,ITENS.SGI_DTPGTO " & vbCrLf
              sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
              sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
              sSql = sSql & "     ,ITENS.SGI_VLDOC " & vbCrLf
              sSql = sSql & "  From" & vbCrLf
              sSql = sSql & "      SGI_CONTASIARC ITENS" & vbCrLf
              sSql = sSql & "     ,SGI_CONTASHARC CABEC" & vbCrLf
              sSql = sSql & "     ,SGI_CADCLIENTE FORNE" & vbCrLf
              sSql = sSql & " Where" & vbCrLf
              sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
              sSql = sSql & "  And ITENS.SGI_VLPAGO IS NOT NULL" & vbCrLf
              sSql = sSql & "  And ITENS.SGI_CODIGO = " & Trim(BREC!SGI_CODIGO)
              sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
              sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
              sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
              sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODCLI " & vbCrLf
              sSql = sSql & "Order By" & vbCrLf
              sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
              sSql = sSql & "       ,CABEC.SGI_CODCLI" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_PARCELA"
              
              BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
              If Not BREC2.EOF Then
                 flxTitBaixados.TextMatrix(I, 0) = ""
                 flxTitBaixados.TextMatrix(I, 1) = BREC2!SGI_CODIGO
                 flxTitBaixados.TextMatrix(I, 2) = BREC2!SGI_NUMDOC
                 flxTitBaixados.TextMatrix(I, 3) = Format(BREC2!SGI_DATAVENC, "DD/MM/YYYY")
                 flxTitBaixados.TextMatrix(I, 4) = Format(BREC2!SGI_DTPGTO, "DD/MM/YYYY")
                 flxTitBaixados.TextMatrix(I, 5) = Format(BREC2!SGI_PARCELA, "##00") & "/" & Format(BREC2!SGI_QTDPARC, "##00")
                 flxTitBaixados.TextMatrix(I, 6) = Format(BREC2!SGI_VLDOC, "#,##0.00")
                 flxTitBaixados.TextMatrix(I, 7) = BREC2!SGI_CODCLI
                 flxTitBaixados.TextMatrix(I, 8) = BREC2!SGI_RAZAOSOC
                 flxTitBaixados.TextMatrix(I, 9) = BREC2!SGI_PARCELA
              End If
              BREC2.Close
           End If
           
        End If
        
     End If
     BREC.Close
      
End Sub


Private Sub Ordem()

  Dim intQtdFiltros As Integer
  
  txtCampos.Text = ""
  
  sSql = ""
  
  intQtdFiltros = cboFiltro.ListIndex
  
  If StTitulos.Tab = 0 Then ConfGrid
  If StTitulos.Tab = 1 Then ConfGridBaixados
  
  If intQtdFiltros = 0 Then
     If StTitulos.Tab = 0 Then
        sSql = "Select" & vbCrLf
        sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
        sSql = sSql & "     ,CABEC.SGI_CODCLI " & vbCrLf
        sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
        sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "      SGI_CONTASIARC ITENS" & vbCrLf
        sSql = sSql & "     ,SGI_CONTASHARC CABEC" & vbCrLf
        sSql = sSql & "     ,SGI_CADCLIENTE FORNE" & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "  And ITENS.SGI_VLPAGO IS NULL" & vbCrLf
        sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
        sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
        sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
        sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODCLI " & vbCrLf
        sSql = sSql & "Order By" & vbCrLf
        sSql = sSql & "        ITENS.SGI_NUMDOC" & vbCrLf
        sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "       ,ITENS.SGI_PARCELA" & vbCrLf
        sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
        
     ElseIf StTitulos.Tab = 1 Then
     
        sSql = "Select" & vbCrLf
        sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
        sSql = sSql & "     ,CABEC.SGI_CODCLI " & vbCrLf
        sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_DTPGTO" & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
        sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "      SGI_CONTASIARC ITENS" & vbCrLf
        sSql = sSql & "     ,SGI_CONTASHARC CABEC" & vbCrLf
        sSql = sSql & "     ,SGI_CADCLIENTE FORNE" & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "  And ITENS.SGI_VLPAGO IS NOT NULL" & vbCrLf
        sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
        sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
        sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
        sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODCLI " & vbCrLf
        sSql = sSql & "Order By" & vbCrLf
        sSql = sSql & "        ITENS.SGI_NUMDOC "
        sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "       ,ITENS.SGI_PARCELA" & vbCrLf
        sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
     
     End If
  ElseIf intQtdFiltros = 1 Then
     If StTitulos.Tab = 0 Then
        sSql = "Select" & vbCrLf
        sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
        sSql = sSql & "     ,CABEC.SGI_CODCLI " & vbCrLf
        sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
        sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "      SGI_CONTASIARC ITENS" & vbCrLf
        sSql = sSql & "     ,SGI_CONTASHARC CABEC" & vbCrLf
        sSql = sSql & "     ,SGI_CADCLIENTE FORNE" & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "  And ITENS.SGI_VLPAGO IS NULL" & vbCrLf
        sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
        sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
        sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
        sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODCLI " & vbCrLf
        sSql = sSql & "Order By" & vbCrLf
        sSql = sSql & "        FORNE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "       ,ITENS.SGI_NUMDOC" & vbCrLf
        sSql = sSql & "       ,ITENS.SGI_PARCELA" & vbCrLf
        sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
     ElseIf StTitulos.Tab = 1 Then
     
        sSql = "Select" & vbCrLf
        sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
        sSql = sSql & "     ,CABEC.SGI_CODCLI " & vbCrLf
        sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_DTPGTO" & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
        sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "      SGI_CONTASIARC ITENS" & vbCrLf
        sSql = sSql & "     ,SGI_CONTASHARC CABEC" & vbCrLf
        sSql = sSql & "     ,SGI_CADCLIENTE FORNE" & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "  And ITENS.SGI_VLPAGO IS NOT NULL" & vbCrLf
        sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
        sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
        sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
        sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODCLI " & vbCrLf
        sSql = sSql & "Order By" & vbCrLf
        sSql = sSql & "        FORNE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "       ,ITENS.SGI_NUMDOC " & vbCrLf
        sSql = sSql & "       ,ITENS.SGI_PARCELA" & vbCrLf
        sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
     
     End If
  ElseIf intQtdFiltros = 2 Then
     If StTitulos.Tab = 0 Then
        sSql = "Select" & vbCrLf
        sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
        sSql = sSql & "     ,CABEC.SGI_CODCLI " & vbCrLf
        sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
        sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "      SGI_CONTASIARC ITENS" & vbCrLf
        sSql = sSql & "     ,SGI_CONTASHARC CABEC" & vbCrLf
        sSql = sSql & "     ,SGI_CADCLIENTE FORNE" & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "  And ITENS.SGI_VLPAGO IS NULL" & vbCrLf
        sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
        sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
        sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
        sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODCLI " & vbCrLf
        sSql = sSql & "Order By" & vbCrLf
        sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
        sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "       ,ITENS.SGI_NUMDOC" & vbCrLf
        sSql = sSql & "       ,ITENS.SGI_PARCELA" & vbCrLf
        
     ElseIf StTitulos.Tab = 1 Then
     
        sSql = "Select" & vbCrLf
        sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
        sSql = sSql & "     ,CABEC.SGI_CODCLI " & vbCrLf
        sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_DTPGTO" & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
        sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "      SGI_CONTASIARC ITENS" & vbCrLf
        sSql = sSql & "     ,SGI_CONTASHARC CABEC" & vbCrLf
        sSql = sSql & "     ,SGI_CADCLIENTE FORNE" & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "  And ITENS.SGI_VLPAGO IS NOT NULL" & vbCrLf
        sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
        sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
        sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
        sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODCLI " & vbCrLf
        sSql = sSql & "Order By" & vbCrLf
        sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
        sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "       ,ITENS.SGI_NUMDOC " & vbCrLf
        sSql = sSql & "       ,ITENS.SGI_PARCELA" & vbCrLf
     
     End If
  ElseIf intQtdFiltros = 3 Then
     If StTitulos.Tab = 1 Then
        
        sSql = "Select" & vbCrLf
        sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
        sSql = sSql & "     ,CABEC.SGI_CODCLI " & vbCrLf
        sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_DTPGTO" & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_DTPGTO" & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
        sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
        sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "      SGI_CONTASIARC ITENS" & vbCrLf
        sSql = sSql & "     ,SGI_CONTASHARC CABEC" & vbCrLf
        sSql = sSql & "     ,SGI_CADCLIENTE FORNE" & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "  And ITENS.SGI_VLPAGO IS NOT NULL" & vbCrLf
        sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
        sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
        sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
        sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODCLI " & vbCrLf
        sSql = sSql & "Order By" & vbCrLf
        sSql = sSql & "        ITENS.SGI_DTPGTO" & vbCrLf
        sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "       ,ITENS.SGI_NUMDOC " & vbCrLf
        sSql = sSql & "       ,ITENS.SGI_PARCELA" & vbCrLf
     
     End If
  End If
  
  BREC.Open sSql, adoBanco_Dados
    
  Do While Not BREC.EOF
     If StTitulos.Tab = 0 Then
        flxContasaReceb.AddItem "" & vbTab & _
                                BREC!SGI_CODIGO & vbTab & _
                                BREC!SGI_NUMDOC & vbTab & _
                                Format(BREC!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                                Format(BREC!SGI_PARCELA, "##00") & "/" & Format(BREC!SGI_QTDPARC, "##00") & vbTab & _
                                Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                                BREC!SGI_CODCLI & vbTab & _
                                BREC!SGI_RAZAOSOC & vbTab & _
                                BREC!SGI_PARCELA
     ElseIf StTitulos.Tab = 1 Then
       flxTitBaixados.AddItem "" & vbTab & _
                              BREC!SGI_CODIGO & vbTab & _
                              BREC!SGI_NUMDOC & vbTab & _
                              Format(BREC!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                              Format(BREC!SGI_DTPGTO, "DD/MM/YYYY") & vbTab & _
                              Format(BREC!SGI_PARCELA, "##00") & "/" & Format(BREC!SGI_QTDPARC, "##00") & vbTab & _
                              Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                              BREC!SGI_CODCLI & vbTab & _
                              BREC!SGI_RAZAOSOC & vbTab & _
                              BREC!SGI_PARCELA
     End If
     BREC.MoveNext
  Loop
  
  BREC.Close

End Sub
Private Sub txtCampos_GotFocus()
    objFuncoes.SelecionaCampos txtCampos.Name, frmCADCONTRECEBP
End Sub

Private Sub txtCampos_KeyPress(KeyAscii As Integer)
    KeyAscii = objFuncoes.Maiuscula(KeyAscii)
End Sub

Private Sub txtCampos_Validate(Cancel As Boolean)


    If Len(Trim(txtCampos.Text)) = 0 Then Exit Sub
    
    sSql = ""
    
    If StTitulos.Tab = 0 Then ConfGrid
    If StTitulos.Tab = 1 Then ConfGridBaixados
    
    If cboFiltro.ListIndex = 0 Then
       
       If IsNumeric(txtCampos.Text) = False Then
          MsgBox "Somente é permitido numero !!!", vbOKOnly + vbCritical, "Aviso"
          txtCampos.Text = ""
          txtCampos.SetFocus
          Exit Sub
       End If
       
       If StTitulos.Tab = 0 Then
          
          sSql = "Select" & vbCrLf
          sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
          sSql = sSql & "     ,CABEC.SGI_CODCLI " & vbCrLf
          sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
          sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
          sSql = sSql & "  From" & vbCrLf
          sSql = sSql & "      SGI_CONTASIARC ITENS" & vbCrLf
          sSql = sSql & "     ,SGI_CONTASHARC CABEC" & vbCrLf
          sSql = sSql & "     ,SGI_CADCLIENTE FORNE" & vbCrLf
          sSql = sSql & " Where" & vbCrLf
          sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
          sSql = sSql & "  And ITENS.SGI_VLPAGO IS NULL" & vbCrLf
          sSql = sSql & "  And ITENS.SGI_NUMDOC = '" & Trim(txtCampos.Text) & "'" & vbCrLf
          sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
          sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
          sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
          sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODCLI " & vbCrLf
          sSql = sSql & "Order By" & vbCrLf
          sSql = sSql & "        ITENS.SGI_NUMDOC " & vbCrLf
          sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
          sSql = sSql & "       ,ITENS.SGI_PARCELA" & vbCrLf
          sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
       
       ElseIf StTitulos.Tab = 1 Then
          
          sSql = "Select" & vbCrLf
          sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
          sSql = sSql & "     ,CABEC.SGI_CODCLI " & vbCrLf
          sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_DTPGTO" & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
          sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
          sSql = sSql & "  From" & vbCrLf
          sSql = sSql & "      SGI_CONTASIARC ITENS" & vbCrLf
          sSql = sSql & "     ,SGI_CONTASHARC CABEC" & vbCrLf
          sSql = sSql & "     ,SGI_CADCLIENTE FORNE" & vbCrLf
          sSql = sSql & " Where" & vbCrLf
          sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
          sSql = sSql & "  And ITENS.SGI_VLPAGO IS NOT NULL" & vbCrLf
          sSql = sSql & "  And ITENS.SGI_NUMDOC = '" & Trim(txtCampos.Text) & "'" & vbCrLf
          sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
          sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
          sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
          sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODCLI " & vbCrLf
          sSql = sSql & "Order By" & vbCrLf
          sSql = sSql & "        ITENS.SGI_NUMDOC " & vbCrLf
          sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
          sSql = sSql & "       ,ITENS.SGI_PARCELA" & vbCrLf
          sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
       
       End If
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
             
          Do While Not BREC.EOF
          
             If StTitulos.Tab = 0 Then
                flxContasaReceb.AddItem "" & vbTab & _
                                        BREC!SGI_CODIGO & vbTab & _
                                        BREC!SGI_NUMDOC & vbTab & _
                                        Format(BREC!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                                        Format(BREC!SGI_PARCELA, "##00") & "/" & Format(BREC!SGI_QTDPARC, "##00") & vbTab & _
                                        Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                                        BREC!SGI_CODCLI & vbTab & _
                                        BREC!SGI_RAZAOSOC & vbTab & _
                                        BREC!SGI_PARCELA
             ElseIf StTitulos.Tab = 1 Then
                flxTitBaixados.AddItem "" & vbTab & _
                                       BREC!SGI_CODIGO & vbTab & _
                                       BREC!SGI_NUMDOC & vbTab & _
                                       Format(BREC!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                                       Format(BREC!SGI_DTPGTO, "DD/MM/YYYY") & vbTab & _
                                       Format(BREC!SGI_PARCELA, "##00") & "/" & Format(BREC!SGI_QTDPARC, "##00") & vbTab & _
                                       Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                                       BREC!SGI_CODCLI & vbTab & _
                                       BREC!SGI_RAZAOSOC & vbTab & _
                                       BREC!SGI_PARCELA
             End If
             BREC.MoveNext
          Loop
          BREC.Close
          
          Exit Sub
          
       End If
                           
    ElseIf cboFiltro.ListIndex = 1 Then
    
       If StTitulos.Tab = 0 Then
          sSql = "Select" & vbCrLf
          sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
          sSql = sSql & "     ,CABEC.SGI_CODCLI " & vbCrLf
          sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
          sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
          sSql = sSql & "  From" & vbCrLf
          sSql = sSql & "      SGI_CONTASIARC ITENS" & vbCrLf
          sSql = sSql & "     ,SGI_CONTASHARC CABEC" & vbCrLf
          sSql = sSql & "     ,SGI_CADCLIENTE FORNE" & vbCrLf
          sSql = sSql & " Where" & vbCrLf
          sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
          sSql = sSql & "  And ITENS.SGI_VLPAGO IS NULL" & vbCrLf
          sSql = sSql & "  And CABEC.SGI_FILIAL   = ITENS.SGI_FILIAL " & vbCrLf
          sSql = sSql & "  And CABEC.SGI_CODIGO   = ITENS.SGI_CODIGO " & vbCrLf
          sSql = sSql & "  And FORNE.SGI_FILIAL   = CABEC.SGI_FILIAL " & vbCrLf
          sSql = sSql & "  And FORNE.SGI_CODIGO   = CABEC.SGI_CODCLI " & vbCrLf
          sSql = sSql & "  And FORNE.SGI_RAZAOSOC Like '" & Trim(txtCampos.Text) & "%'" & vbCrLf
          sSql = sSql & "Order By" & vbCrLf
          sSql = sSql & "        FORNE.SGI_RAZAOSOC" & vbCrLf
          sSql = sSql & "       ,ITENS.SGI_NUMDOC " & vbCrLf
          sSql = sSql & "       ,ITENS.SGI_PARCELA" & vbCrLf
          sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
       ElseIf StTitulos.Tab = 1 Then
          sSql = "Select" & vbCrLf
          sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
          sSql = sSql & "     ,CABEC.SGI_CODCLI " & vbCrLf
          sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_DTPGTO" & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
          sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
          sSql = sSql & "  From" & vbCrLf
          sSql = sSql & "      SGI_CONTASIARC ITENS" & vbCrLf
          sSql = sSql & "     ,SGI_CONTASHARC CABEC" & vbCrLf
          sSql = sSql & "     ,SGI_CADCLIENTE FORNE" & vbCrLf
          sSql = sSql & " Where" & vbCrLf
          sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
          sSql = sSql & "  And ITENS.SGI_VLPAGO IS NOT NULL" & vbCrLf
          sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
          sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
          sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
          sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODCLI " & vbCrLf
          sSql = sSql & "  And FORNE.SGI_RAZAOSOC Like '" & Trim(txtCampos.Text) & "%'" & vbCrLf
          sSql = sSql & "Order By" & vbCrLf
          sSql = sSql & "        FORNE.SGI_RAZAOSOC" & vbCrLf
          sSql = sSql & "       ,ITENS.SGI_NUMDOC " & vbCrLf
          sSql = sSql & "       ,ITENS.SGI_PARCELA" & vbCrLf
          sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
       End If
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
              
          Do While Not BREC.EOF
             If StTitulos.Tab = 0 Then
                flxContasaReceb.AddItem "" & vbTab & _
                                        BREC!SGI_CODIGO & vbTab & _
                                        BREC!SGI_NUMDOC & vbTab & _
                                        Format(BREC!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                                        Format(BREC!SGI_PARCELA, "##00") & "/" & Format(BREC!SGI_QTDPARC, "##00") & vbTab & _
                                        Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                                        BREC!SGI_CODCLI & vbTab & _
                                        BREC!SGI_RAZAOSOC & vbTab & _
                                        BREC!SGI_PARCELA
             ElseIf StTitulos.Tab = 1 Then
                flxTitBaixados.AddItem "" & vbTab & _
                                       BREC!SGI_CODIGO & vbTab & _
                                       BREC!SGI_NUMDOC & vbTab & _
                                       Format(BREC!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                                       Format(BREC!SGI_DTPGTO, "DD/MM/YYYY") & vbTab & _
                                       Format(BREC!SGI_PARCELA, "##00") & "/" & Format(BREC!SGI_QTDPARC, "##00") & vbTab & _
                                       Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                                       BREC!SGI_CODCLI & vbTab & _
                                       BREC!SGI_RAZAOSOC & vbTab & _
                                       BREC!SGI_PARCELA
             End If
             BREC.MoveNext
          Loop
              
          BREC.Close
          Exit Sub
          
       End If
    
    ElseIf cboFiltro.ListIndex = 2 Then
    
        If Not IsDate(txtCampos.Text) Then
          MsgBox "Somente é permitido Datas !!!", vbOKOnly + vbCritical, "Aviso"
          txtCampos.Text = ""
          txtCampos.SetFocus
          Exit Sub
       End If
      
       If StTitulos.Tab = 0 Then
          sSql = "Select" & vbCrLf
          sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
          sSql = sSql & "     ,CABEC.SGI_CODCLI " & vbCrLf
          sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
          sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
          sSql = sSql & "  From" & vbCrLf
          sSql = sSql & "      SGI_CONTASIARC ITENS" & vbCrLf
          sSql = sSql & "     ,SGI_CONTASHARC CABEC" & vbCrLf
          sSql = sSql & "     ,SGI_CADCLIENTE FORNE" & vbCrLf
          sSql = sSql & " Where" & vbCrLf
          sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
          sSql = sSql & "  And ITENS.SGI_VLPAGO IS NULL" & vbCrLf
          sSql = sSql & "  And ITENS.SGI_DATAVENC = '" & Format(CDate(txtCampos.Text), "MM/DD/YYYY") & "'" & vbCrLf
          sSql = sSql & "  And CABEC.SGI_FILIAL   = ITENS.SGI_FILIAL " & vbCrLf
          sSql = sSql & "  And CABEC.SGI_CODIGO   = ITENS.SGI_CODIGO " & vbCrLf
          sSql = sSql & "  And FORNE.SGI_FILIAL   = CABEC.SGI_FILIAL " & vbCrLf
          sSql = sSql & "  And FORNE.SGI_CODIGO   = CABEC.SGI_CODCLI " & vbCrLf
          sSql = sSql & "Order By" & vbCrLf
          sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
          sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
          sSql = sSql & "       ,ITENS.SGI_NUMDOC " & vbCrLf
          sSql = sSql & "       ,ITENS.SGI_PARCELA" & vbCrLf
       ElseIf StTitulos.Tab = 1 Then
          sSql = "Select" & vbCrLf
          sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
          sSql = sSql & "     ,CABEC.SGI_CODCLI " & vbCrLf
          sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_DTPGTO" & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
          sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
          sSql = sSql & "  From" & vbCrLf
          sSql = sSql & "      SGI_CONTASIARC ITENS" & vbCrLf
          sSql = sSql & "     ,SGI_CONTASHARC CABEC" & vbCrLf
          sSql = sSql & "     ,SGI_CADCLIENTE FORNE" & vbCrLf
          sSql = sSql & " Where" & vbCrLf
          sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
          sSql = sSql & "  And ITENS.SGI_VLPAGO IS NOT NULL" & vbCrLf
          sSql = sSql & "  And ITENS.SGI_DATAVENC = '" & Format(CDate(txtCampos.Text), "MM/DD/YYYY") & "'" & vbCrLf
          sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
          sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
          sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
          sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODCLI " & vbCrLf
          sSql = sSql & "Order By" & vbCrLf
          sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
          sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
          sSql = sSql & "       ,ITENS.SGI_NUMDOC " & vbCrLf
          sSql = sSql & "       ,ITENS.SGI_PARCELA" & vbCrLf
       End If
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
            
          Do While Not BREC.EOF
             If StTitulos.Tab = 0 Then
                flxContasaReceb.AddItem "" & vbTab & _
                                        BREC!SGI_CODIGO & vbTab & _
                                        BREC!SGI_NUMDOC & vbTab & _
                                        Format(BREC!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                                        Format(BREC!SGI_PARCELA, "##00") & "/" & Format(BREC!SGI_QTDPARC, "##00") & vbTab & _
                                        Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                                        BREC!SGI_CODCLI & vbTab & _
                                        BREC!SGI_RAZAOSOC & vbTab & _
                                        BREC!SGI_PARCELA
             ElseIf StTitulos.Tab = 1 Then
                flxTitBaixados.AddItem "" & vbTab & _
                                       BREC!SGI_CODIGO & vbTab & _
                                       BREC!SGI_NUMDOC & vbTab & _
                                       Format(BREC!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                                       Format(BREC!SGI_DTPGTO, "DD/MM/YYYY") & vbTab & _
                                       Format(BREC!SGI_PARCELA, "##00") & "/" & Format(BREC!SGI_QTDPARC, "##00") & vbTab & _
                                       Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                                       BREC!SGI_CODCLI & vbTab & _
                                       BREC!SGI_RAZAOSOC & vbTab & _
                                       BREC!SGI_PARCELA
             End If
             BREC.MoveNext
          Loop
              
          BREC.Close
          Exit Sub
          
       End If
    
    ElseIf cboFiltro.ListIndex = 3 Then
       If StTitulos.Tab = 1 Then
          sSql = "Select" & vbCrLf
          sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
          sSql = sSql & "     ,CABEC.SGI_CODCLI " & vbCrLf
          sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_DTPGTO" & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
          sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
          sSql = sSql & "  From" & vbCrLf
          sSql = sSql & "      SGI_CONTASIARC ITENS" & vbCrLf
          sSql = sSql & "     ,SGI_CONTASHARC CABEC" & vbCrLf
          sSql = sSql & "     ,SGI_CADCLIENTE FORNE" & vbCrLf
          sSql = sSql & " Where" & vbCrLf
          sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
          sSql = sSql & "  And ITENS.SGI_VLPAGO IS NOT NULL" & vbCrLf
          sSql = sSql & "  And ITENS.SGI_DATAVENC = '" & Format(CDate(txtCampos.Text), "MM/DD/YYYY") & "'" & vbCrLf
          sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
          sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
          sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
          sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODCLI " & vbCrLf
          sSql = sSql & "Order By" & vbCrLf
          sSql = sSql & "        ITENS.SGI_DTPGTO" & vbCrLf
          sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
          sSql = sSql & "       ,ITENS.SGI_NUMDOC " & vbCrLf
          sSql = sSql & "       ,ITENS.SGI_PARCELA" & vbCrLf
       End If
         
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
              
          Do While Not BREC.EOF
             If StTitulos.Tab = 1 Then
                flxTitBaixados.AddItem "" & vbTab & _
                                       BREC!SGI_CODIGO & vbTab & _
                                       BREC!SGI_NUMDOC & vbTab & _
                                       Format(BREC!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                                       Format(BREC!SGI_DTPGTO, "DD/MM/YYYY") & vbTab & _
                                       Format(BREC!SGI_PARCELA, "##00") & "/" & Format(BREC!SGI_QTDPARC, "##00") & vbTab & _
                                       Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                                       BREC!SGI_CODCLI & vbTab & _
                                       BREC!SGI_RAZAOSOC & vbTab & _
                                       BREC!SGI_PARCELA
             End If
             BREC.MoveNext
          Loop
              
          BREC.Close
          Exit Sub
          
       End If
    
    End If
    BREC.Close
    
    If StTitulos.Tab = 0 Then flxContasaReceb.SetFocus
    If StTitulos.Tab = 1 Then flxTitBaixados.SetFocus
    

End Sub

Private Function Verif_reg() As Boolean

    Verif_reg = False
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "      * " & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_CONTASIARC ITENS" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "  And ITENS.SGI_CODIGO = " & objCADCONTRECEB.CODPGTO & vbCrLf
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF Then
       MsgBox "Este registro foi Excluso !!!", vbOKOnly + vbExclamation, "Aviso"
       Verif_reg = True
    End If
    BREC.Close

End Function

