VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCADCONTASAPGP 
   Caption         =   "Cadastro de Contas a Pagar"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11460
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   11460
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   4815
      Left            =   0
      TabIndex        =   12
      Top             =   1560
      Width           =   11415
      Begin TabDlg.SSTab StTitulos 
         Height          =   4575
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   8070
         _Version        =   393216
         Style           =   1
         Tabs            =   5
         TabsPerRow      =   5
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
         TabPicture(0)   =   "frmCADCONTASAPGP.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "flxCADCONTASAPG"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Titulos Baixados"
         TabPicture(1)   =   "frmCADCONTASAPGP.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "flxTitBaixados"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Lotes Gerados"
         TabPicture(2)   =   "frmCADCONTASAPGP.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label3(0)"
         Tab(2).Control(1)=   "Label3(1)"
         Tab(2).Control(2)=   "Label3(2)"
         Tab(2).Control(3)=   "Label3(3)"
         Tab(2).Control(4)=   "Label4"
         Tab(2).Control(5)=   "flxLote"
         Tab(2).ControlCount=   6
         TabCaption(3)   =   "Lotes Liberados"
         TabPicture(3)   =   "frmCADCONTASAPGP.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "flxLoteLiberado"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Lotes Baixados"
         TabPicture(4)   =   "frmCADCONTASAPGP.frx":0070
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "flxLoteBaixados"
         Tab(4).ControlCount=   1
         Begin MSFlexGridLib.MSFlexGrid flxLoteBaixados 
            Height          =   4095
            Left            =   -74880
            TabIndex        =   29
            Top             =   360
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   7223
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin MSFlexGridLib.MSFlexGrid flxLoteLiberado 
            Height          =   4095
            Left            =   -74880
            TabIndex        =   22
            Top             =   360
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   7223
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin MSFlexGridLib.MSFlexGrid flxLote 
            Height          =   3375
            Left            =   -74880
            TabIndex        =   19
            Top             =   360
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   5953
            _Version        =   393216
            FixedCols       =   0
            AllowBigSelection=   0   'False
            HighLight       =   2
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin MSFlexGridLib.MSFlexGrid flxTitBaixados 
            Height          =   4095
            Left            =   -74880
            TabIndex        =   16
            Top             =   360
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   7223
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin MSFlexGridLib.MSFlexGrid flxCADCONTASAPG 
            Height          =   4095
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   7223
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
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
            Left            =   -71520
            TabIndex        =   30
            Top             =   3960
            Width           =   75
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Index           =   3
            Left            =   -73200
            TabIndex        =   26
            Top             =   4200
            Width           =   1365
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Index           =   2
            Left            =   -73200
            TabIndex        =   25
            Top             =   3840
            Width           =   1365
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Total Selecionado:"
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
            Left            =   -74880
            TabIndex        =   24
            Top             =   4200
            Width           =   1620
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Total do Lotes:"
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
            Left            =   -74550
            TabIndex        =   23
            Top             =   3840
            Width           =   1305
         End
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   11415
      Begin VB.CommandButton cmdCheque 
         Caption         =   "&Cheques"
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
         Left            =   6960
         Picture         =   "frmCADCONTASAPGP.frx":008C
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Baixa titulos ou lote de pagamento"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdImpLotes 
         Caption         =   "Im&prime"
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
         Left            =   6120
         Picture         =   "frmCADCONTASAPGP.frx":0396
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Baixa titulos ou lote de pagamento"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdDeslib 
         Caption         =   "&Deslibera"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5160
         Picture         =   "frmCADCONTASAPGP.frx":06A0
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Libera lote"
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdLibera 
         Caption         =   "Libe&ra"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4320
         Picture         =   "frmCADCONTASAPGP.frx":0AC6
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Libera lote"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdLote 
         Caption         =   "&Lote"
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
         Picture         =   "frmCADCONTASAPGP.frx":0EF3
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Gera lote de pagamento"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdExtorna 
         Caption         =   "E&xtorna"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8640
         Picture         =   "frmCADCONTASAPGP.frx":1335
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Extorna titulos ou lote de pagamento"
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
         Left            =   7800
         Picture         =   "frmCADCONTASAPGP.frx":1437
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Baixa titulos ou lote de pagamento"
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
         Left            =   10440
         Picture         =   "frmCADCONTASAPGP.frx":1D01
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Left            =   9600
         Picture         =   "frmCADCONTASAPGP.frx":1E03
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Picture         =   "frmCADCONTASAPGP.frx":2335
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Exclui um titulo de pagamento"
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
         Picture         =   "frmCADCONTASAPGP.frx":2437
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Altera o titulo de pagamento"
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
         Picture         =   "frmCADCONTASAPGP.frx":2539
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Inclui um novo titulo de pagamento"
         Top             =   120
         Width           =   855
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
         Picture         =   "frmCADCONTASAPGP.frx":2A6B
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Volta ao Menu Principal"
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11415
      Begin VB.Timer Timer1 
         Interval        =   50000
         Left            =   9000
         Top             =   120
      End
      Begin VB.TextBox txtCampos 
         Height          =   285
         Left            =   3360
         MaxLength       =   50
         TabIndex        =   2
         Text            =   "txtCampos"
         Top             =   200
         Width           =   7935
      End
      Begin VB.ComboBox cboFiltro 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   200
         Width           =   1695
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
         Left            =   2640
         TabIndex        =   4
         Top             =   240
         Width           =   645
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
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmCADCONTASAPGP"
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
Dim objCADCONTASAPG As Object
Dim objREL          As Object
Dim iCodigo         As Integer
Dim cCamRel         As String
Dim strTitulo       As String
Dim strCABEC2       As String


Private Sub cmdAltera_Click()
    
    If objFuncoes.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    If StTitulos.Tab = 0 Then
       If VerifBaixados = False Then Exit Sub
       Operacao "A"
    End If
    
    If StTitulos.Tab = 1 And flxTitBaixados.Rows > 1 Then
       If VerifLote = False Then Exit Sub
       If VerificaCaixa = True Then Exit Sub
       Operacao "AB"
    End If
    
    If StTitulos.Tab = 2 Then Operacao "A"
    
End Sub

Private Sub cmdBaixa_Click()
    
    If StTitulos.Tab = 0 Then
       If flxCADCONTASAPG.TextMatrix(flxCADCONTASAPG.RowSel, 9) = "L" Then
          MsgBox "Este titulo esta incluso no lote e aguarda baixa !!!", vbOKOnly + vbExclamation, "Aviso"
          Exit Sub
       End If
       Operacao "B"
    End If
    If StTitulos.Tab = 3 Then BaixaLiberados
    
    
End Sub

Private Sub cmdCanFiltro_Click()

    StTitulos.Tab = 0
    
    cboFiltro.Clear
    cboFiltro.AddItem "Código"
    cboFiltro.AddItem "Nº Doc"
    cboFiltro.AddItem "Fornecedor"
    cboFiltro.AddItem "Dt. Vencto"
    cboFiltro.ListIndex = 0
    
    txtCampos.Text = ""
    
    AbilitaCampos
    ConfGrid
    ConfGridBaixados
    ConfGridLote
    ConfGridLoteLiberdo
    ConfGridLoteBaixado
    PreencheGrid
    PopGridBaixados
    PreenchGridLote
    PreenchGridLoteLiberado
    PreenchGridLoteBaixado

End Sub

Private Sub cmdDeslib_Click()
    If StTitulos.Tab = 3 Then DesLibera
End Sub

Private Sub cmdExclui_Click()

  If objFuncoes.ChecaAcesso2("E", strAcesso) = False Then Exit Sub
  
  If Verif_reg = True Then Exit Sub
  
  Dim iresp As Integer
  
  If StTitulos.Tab = 1 Or StTitulos.Tab = 2 Then Exit Sub
  If VerifBaixados = False Then Exit Sub
  
  iresp = MsgBox("Confirma a exclusão do registro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
  If iresp <> 6 Then Exit Sub
  If objCADCONTASAPG.GRAVA("E") = False Then Exit Sub
  If objCADCONTASAPG.Atualiza("E", Str(objCADCONTASAPG.CODPGTO), FILIAL, "frmCADCONTASAPG") = False Then Exit Sub
  
  MsgBox "Registro excluso com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
  
  AbilitaCampos
  Atualiza_Grid

End Sub

Private Sub cmdExtorna_Click()

  Dim iresp As Integer
  
  If StTitulos.Tab = 1 Then
     
     If Verif_reg = True Then Exit Sub
     
     If (flxTitBaixados.Rows - 1) = 0 Then Exit Sub
     
     If VerifLote = False Then Exit Sub
     If VerificaCaixa = True Then Exit Sub
  
     iresp = MsgBox("Confirma o extorno do titulo ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
     If iresp <> 6 Then Exit Sub
  
     objCADCONTASAPG.PARCPGTO = CInt(flxTitBaixados.TextMatrix(flxTitBaixados.Row, 9))
     objCADCONTASAPG.NUMDOC = flxTitBaixados.TextMatrix(flxTitBaixados.Row, 2)
     objCADCONTASAPG.FILIAL = FILIAL
     
     If objCADCONTASAPG.GRAVA("X") = False Then Exit Sub
  
     MsgBox "Registro extornado com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
     
     If objCADCONTASAPG.Atualiza("X", Str(objCADCONTASAPG.CODPGTO), FILIAL, "frmCADCONTASAPG") = False Then Exit Sub
     
  End If
  
  If StTitulos.Tab = 2 Then
  
     If (flxLote.Rows - 1) = 0 Then Exit Sub
    
     If flxLote.TextMatrix(flxLote.RowSel, 7) = "L" Then
        MsgBox "Este lote vá foi liberado, favor desliberar !!!", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
     End If
     
     iresp = MsgBox("Confirma o extorno do lote ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
  
     If iresp <> 6 Then Exit Sub
  
     objCADCONTASAPG.FILIAL = FILIAL
     If objCADCONTASAPG.GRAVALOTE("X") = False Then Exit Sub
     If objCADCONTASAPG.Atualiza("X", Str(objCADCONTASAPG.CODPGTO), FILIAL, "LOTE") = False Then Exit Sub
  
     MsgBox "lote desliberado com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
  
  End If
  
  If StTitulos.Tab = 4 Then ExtornaLotePago
  
  AbilitaCampos
  Atualiza_Grid

End Sub

Private Sub cmdImpLotes_Click()
    If StTitulos.Tab = 2 And flxLote.Rows > 1 Then Implote
    If StTitulos.Tab = 3 And flxLoteLiberado.Rows > 1 Then Implote
    If StTitulos.Tab = 4 And flxLoteBaixados.Rows > 1 Then Implote
End Sub

Private Sub cmdInclui_Click()
    If objFuncoes.ChecaAcesso2("I", strAcesso) = False Then Exit Sub
    Operacao "I"
End Sub

Private Sub cmdLibera_Click()
   If StTitulos.Tab = 2 Then Libera
End Sub

Private Sub cmdLote_Click()
  
  If StTitulos.Tab <> 0 Then Exit Sub
  
  If (flxCADCONTASAPG.Rows) = 1 Then
     MsgBox "Não há titulos para gerar lote !!!", vbOKOnly + vbExclamation, "Aviso"
     Exit Sub
  End If
  
  frmCADLOTE.cCaminho = cCaminho
  frmCADLOTE.Linha = Linha
  frmCADLOTE.iCodigo = iCodigo
  frmCADLOTE.cTipOper = "I"
  frmCADLOTE.FILIAL = FILIAL
  frmCADLOTE.strAcesso = strAcesso
  frmCADLOTE.strMODPAI = Me.Name
  frmCADLOTE.strUSUARIO = strUSUARIO
  
  frmCADLOTE.Show vbModal
  
  StTitulos.Tab = 2
  AbilitaCampos
  Atualiza_Grid

End Sub

Private Sub cmdOrden_Click()
    Call Ordem
End Sub

Private Sub cmdVoltar_Click()
    Set objFuncoes = Nothing
    Set objCADCONTASAPG = Nothing
    Set objREL = Nothing
    Unload Me
End Sub
Private Sub flxCADCONTASAPG_Click()
    If flxCADCONTASAPG.Rows > 1 Then objCADCONTASAPG.CODPGTO = CInt(flxCADCONTASAPG.TextMatrix(flxCADCONTASAPG.RowSel, 1))
End Sub

Private Sub flxCADCONTASAPG_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    If flxCADCONTASAPG.Rows > 1 Then Operacao "C"
End Sub

Private Sub flxCADCONTASAPG_RowColChange()
    If flxCADCONTASAPG.Rows > 1 Then objCADCONTASAPG.CODPGTO = CInt(flxCADCONTASAPG.TextMatrix(flxCADCONTASAPG.RowSel, 1))
End Sub

Private Sub flxLote_Click()
    If flxLote.Rows > 1 Then objCADCONTASAPG.CODLOTE = CInt(flxLote.TextMatrix(flxLote.RowSel, 1))
End Sub

Private Sub flxLote_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    If flxLote.Rows > 1 Then Operacao "C"
End Sub

Private Sub flxLote_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
       If flxLote.Rows = 1 Then Exit Sub
       
       If flxLote.TextMatrix(flxLote.Row, 9) = "*" Then
          flxLote.TextMatrix(flxLote.Row, 9) = ""
       ElseIf flxLote.TextMatrix(flxLote.Row, 9) = "" Then
          flxLote.TextMatrix(flxLote.Row, 9) = "*"
       End If
       SomaTitSelec
       
    End If
End Sub

Private Sub flxLote_RowColChange()
    If flxLote.Rows > 1 Then objCADCONTASAPG.CODLOTE = CInt(flxLote.TextMatrix(flxLote.RowSel, 1))
End Sub

Private Sub flxLoteBaixados_Click()
    If flxLoteBaixados.Rows > 1 Then objCADCONTASAPG.CODLOTE = CInt(flxLoteBaixados.TextMatrix(flxLoteBaixados.RowSel, 1))
End Sub

Private Sub flxLoteBaixados_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    If flxLoteBaixados.Rows > 1 Then Operacao "C"
End Sub

Private Sub flxLoteBaixados_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeySpace Then
       If flxLoteBaixados.Rows = 1 Then Exit Sub
       
       If flxLoteBaixados.TextMatrix(flxLoteBaixados.Row, 9) = "*" Then
          flxLoteBaixados.TextMatrix(flxLoteBaixados.Row, 9) = ""
       ElseIf flxLoteBaixados.TextMatrix(flxLoteBaixados.Row, 9) = "" Then
          flxLoteBaixados.TextMatrix(flxLoteBaixados.Row, 9) = "*"
       End If
       
    End If

End Sub

Private Sub flxLoteBaixados_RowColChange()
    If flxLoteBaixados.Rows > 1 Then objCADCONTASAPG.CODLOTE = CInt(flxLoteBaixados.TextMatrix(flxLoteBaixados.RowSel, 1))
End Sub

Private Sub flxLoteLiberado_Click()
    If flxLoteLiberado.Rows > 1 Then objCADCONTASAPG.CODLOTE = CInt(flxLoteLiberado.TextMatrix(flxLoteLiberado.RowSel, 1))
End Sub

Private Sub flxLoteLiberado_DblClick()
    If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
    If flxLoteLiberado.Rows > 1 Then Operacao "C"
End Sub

Private Sub flxLoteLiberado_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeySpace Then
       If flxLoteLiberado.Rows = 1 Then Exit Sub
       
       If flxLoteLiberado.TextMatrix(flxLoteLiberado.Row, 9) = "*" Then
          flxLoteLiberado.TextMatrix(flxLoteLiberado.Row, 9) = ""
       ElseIf flxLoteLiberado.TextMatrix(flxLoteLiberado.Row, 9) = "" Then
          flxLoteLiberado.TextMatrix(flxLoteLiberado.Row, 9) = "*"
       End If
       
    End If

End Sub

Private Sub flxLoteLiberado_RowColChange()
    If flxLoteLiberado.Rows > 1 Then objCADCONTASAPG.CODLOTE = CInt(flxLoteLiberado.TextMatrix(flxLoteLiberado.RowSel, 1))
End Sub

Private Sub flxTitBaixados_Click()
   If flxTitBaixados.Rows > 1 Then objCADCONTASAPG.CODPGTO = CInt(flxTitBaixados.TextMatrix(flxTitBaixados.RowSel, 1))
End Sub

Private Sub flxTitBaixados_DblClick()
   If objFuncoes.ChecaAcesso2("C", strAcesso) = False Then Exit Sub
   If flxTitBaixados.Rows > 1 Then Operacao "CB"
End Sub

Private Sub flxTitBaixados_RowColChange()
   If flxTitBaixados.Rows > 1 Then objCADCONTASAPG.CODPGTO = CInt(flxTitBaixados.TextMatrix(flxTitBaixados.RowSel, 1))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
    Set objCADCONTASAPG = CreateObject("CADCONTASAPG.clsCADCONTASAPG")
    Set objREL = CreateObject("MOSTRAREL.clsMOSTRAREL")
    
    objCADCONTASAPG.FILIAL = FILIAL
    
    objFuncoes.LimpaCampos frmCADCONTASAPGP
    
    Set adoBanco_Dados = objFuncoes.Banco_Dados(Linha)

    If adoBanco_Dados.State = 0 Then
       MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    AbilitaCampos
    ConfGrid
    ConfGridBaixados
    ConfGridLote
    ConfGridLoteLiberdo
    ConfGridLoteBaixado
    PreencheGrid
    PopGridBaixados
    PreenchGridLote
    PreenchGridLoteLiberado
    PreenchGridLoteBaixado
        
    cboFiltro.AddItem "Código"
    cboFiltro.AddItem "Nº Doc"
    cboFiltro.AddItem "Fornecedor"
    cboFiltro.AddItem "Dt. Vencto"
    
    cboFiltro.ListIndex = 0
    StTitulos.Tab = 0
    
    
    Label4.Caption = "Para selecionar um ou mais lotes, selecione o lote com click " & vbCrLf
    Label4.Caption = Label4.Caption & "e depois pressione barra de espaço."
    
    strCamRelNovo = Right(Linha(7), Len(Trim(Linha(7))) - 7)
    
End Sub

Private Sub AbilitaCampos()
    
    If objCADCONTASAPG.Pesq_CadContasAPG = False Then
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
    
    flxCADCONTASAPG.Rows = 1
    flxCADCONTASAPG.Cols = 10
    
    flxCADCONTASAPG.TextMatrix(0, 0) = ""
    flxCADCONTASAPG.TextMatrix(0, 1) = "Código"
    flxCADCONTASAPG.TextMatrix(0, 2) = "Nº Doc."
    flxCADCONTASAPG.TextMatrix(0, 3) = "Vencto"
    flxCADCONTASAPG.TextMatrix(0, 4) = "Parcela"
    flxCADCONTASAPG.TextMatrix(0, 5) = "Valor"
    flxCADCONTASAPG.TextMatrix(0, 6) = "Cod. Forn"
    flxCADCONTASAPG.TextMatrix(0, 7) = "Razão Social"
    flxCADCONTASAPG.TextMatrix(0, 8) = "Parcela"
    flxCADCONTASAPG.TextMatrix(0, 9) = "Status"
    
    flxCADCONTASAPG.ColWidth(0) = 0
    flxCADCONTASAPG.ColWidth(1) = 1000
    flxCADCONTASAPG.ColWidth(2) = 1000
    flxCADCONTASAPG.ColWidth(3) = 1000
    flxCADCONTASAPG.ColWidth(4) = 1000
    flxCADCONTASAPG.ColWidth(5) = 1000
    flxCADCONTASAPG.ColWidth(6) = 1000
    flxCADCONTASAPG.ColWidth(7) = 4000
    flxCADCONTASAPG.ColWidth(8) = 0
    flxCADCONTASAPG.ColWidth(9) = 0
    
End Sub

Private Sub PreencheGrid()

    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
    sSql = sSql & "     ,CABEC.SGI_CODFOR " & vbCrLf
    sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
    sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_STATUS  " & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_CONTASIAPG ITENS" & vbCrLf
    sSql = sSql & "     ,SGI_CONTASHAPG CABEC" & vbCrLf
    sSql = sSql & "     ,SGI_CADFORNEC  FORNE" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "  And ITENS.SGI_VLPAGO IS NULL" & vbCrLf
    sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
    sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
    sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
    sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODFOR " & vbCrLf
    sSql = sSql & "Order By" & vbCrLf
    sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
    sSql = sSql & "       ,CABEC.SGI_CODFOR" & vbCrLf
    sSql = sSql & "       ,ITENS.SGI_PARCELA"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       flxCADCONTASAPG.AddItem "" & vbTab & _
                               BREC!SGI_CODIGO & vbTab & _
                               BREC!SGI_NUMDOC & vbTab & _
                               Format(BREC!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                               Format(BREC!SGI_PARCELA, "##00") & "/" & Format(BREC!SGI_QTDPARC, "##00") & vbTab & _
                               Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                               BREC!SGI_CODFOR & vbTab & _
                               BREC!SGI_RAZAOSOC & vbTab & _
                               BREC!SGI_PARCELA & vbTab & _
                               BREC!SGI_STATUS
                               
       BREC.MoveNext
    Loop
    
    PosGrid
    
    BREC.Close
    
End Sub

Private Sub PosGrid()

    If iCodigo > 0 Then
        Dim I As Integer
               
        For I = 1 To (flxCADCONTASAPG.Rows - 1)
            
            If flxCADCONTASAPG.TextMatrix(I, 1) = iCodigo Then
               flxCADCONTASAPG.Row = I
               flxCADCONTASAPG.Col = 1
               
               Exit For
            End If
         
        Next I
    End If

End Sub

Private Sub Operacao(strOperacao As String)
 
  If strOperacao = "A" Or _
     strOperacao = "C" Or _
     strOperacao = "B" Or _
     strOperacao = "AB" Then
     If Verif_reg = True Then Exit Sub
  End If
  
  Dim Pesquisa As String
  
  If StTitulos.Tab = 0 Then
     If flxCADCONTASAPG.Rows > 1 Then iCodigo = CInt(flxCADCONTASAPG.TextMatrix(flxCADCONTASAPG.RowSel, 1))
  End If
  If StTitulos.Tab = 1 Then
     If flxTitBaixados.Rows > 1 Then iCodigo = CInt(flxTitBaixados.TextMatrix(flxTitBaixados.RowSel, 1))
  End If
  If StTitulos.Tab = 2 Then
     If flxLote.Rows > 1 Then iCodigo = CInt(flxLote.TextMatrix(flxLote.RowSel, 1))
  End If
  If StTitulos.Tab = 3 Then
     If flxLoteLiberado.Rows > 1 Then iCodigo = CInt(flxLoteLiberado.TextMatrix(flxLoteLiberado.RowSel, 1))
  End If
  If StTitulos.Tab = 4 Then
     If flxLoteBaixados.Rows > 1 Then iCodigo = CInt(flxLoteBaixados.TextMatrix(flxLoteBaixados.RowSel, 1))
  End If
  
  If StTitulos.Tab = 0 Or StTitulos.Tab = 1 Then
     frmCADCONTASAPG.cCaminho = cCaminho
     frmCADCONTASAPG.Linha = Linha
     frmCADCONTASAPG.iCodigo = iCodigo
     frmCADCONTASAPG.cTipOper = strOperacao
     frmCADCONTASAPG.FILIAL = FILIAL
     frmCADCONTASAPG.strAcesso = strAcesso
     frmCADCONTASAPG.strMODPAI = Me.Name
     frmCADCONTASAPG.strUSUARIO = strUSUARIO
  End If
  
  If StTitulos.Tab = 2 Or _
     StTitulos.Tab = 3 Or _
     StTitulos.Tab = 4 Then
     frmCADLOTE.cCaminho = cCaminho
     frmCADLOTE.Linha = Linha
     frmCADLOTE.iCodigo = iCodigo
     frmCADLOTE.cTipOper = strOperacao
     frmCADLOTE.strAcesso = strAcesso
     frmCADLOTE.strMODPAI = Me.Name
     frmCADLOTE.strUSUARIO = strUSUARIO
     frmCADLOTE.FILIAL = FILIAL
  End If
  
  If StTitulos.Tab = 0 Then
     If flxCADCONTASAPG.Rows > 1 Then frmCADCONTASAPG.iParcela = CInt(flxCADCONTASAPG.TextMatrix(flxCADCONTASAPG.RowSel, 8))
  End If
  If StTitulos.Tab = 1 Then
     If flxTitBaixados.Rows > 1 Then frmCADCONTASAPG.iParcela = CInt(flxTitBaixados.TextMatrix(flxTitBaixados.RowSel, 9))
  End If
  
  If StTitulos.Tab = 0 Or StTitulos.Tab = 1 Then frmCADCONTASAPG.Show vbModal
  If StTitulos.Tab = 2 Or StTitulos.Tab = 3 Or StTitulos.Tab = 4 Then frmCADLOTE.Show vbModal
  
  AbilitaCampos
  Atualiza_Grid
  
End Sub


Public Sub ConfGridBaixados()

    flxTitBaixados.Rows = 1
    flxTitBaixados.Cols = 11
    
    flxTitBaixados.TextMatrix(0, 0) = ""
    flxTitBaixados.TextMatrix(0, 1) = "Código"
    flxTitBaixados.TextMatrix(0, 2) = "Nº Doc."
    flxTitBaixados.TextMatrix(0, 3) = "Vencto"
    flxTitBaixados.TextMatrix(0, 4) = "Pagto"
    flxTitBaixados.TextMatrix(0, 5) = "Parcela"
    flxTitBaixados.TextMatrix(0, 6) = "Valor"
    flxTitBaixados.TextMatrix(0, 7) = "Cod. Forn"
    flxTitBaixados.TextMatrix(0, 8) = "Razão Social"
    flxTitBaixados.TextMatrix(0, 9) = "Parcela"
    flxTitBaixados.TextMatrix(0, 10) = "Status"
    
    flxTitBaixados.ColWidth(0) = 0
    flxTitBaixados.ColWidth(1) = 1000
    flxTitBaixados.ColWidth(2) = 1000
    flxTitBaixados.ColWidth(3) = 1000
    flxTitBaixados.ColWidth(4) = 1000
    flxTitBaixados.ColWidth(5) = 1000
    flxTitBaixados.ColWidth(6) = 1000
    flxTitBaixados.ColWidth(7) = 1000
    flxTitBaixados.ColWidth(8) = 4000
    flxTitBaixados.ColWidth(9) = 0
    flxTitBaixados.ColWidth(10) = 0

End Sub

Private Sub PopGridBaixados()

    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
    sSql = sSql & "     ,CABEC.SGI_CODFOR " & vbCrLf
    sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_DTPGTO" & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
    sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_STATUS  " & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_CONTASIAPG ITENS" & vbCrLf
    sSql = sSql & "     ,SGI_CONTASHAPG CABEC" & vbCrLf
    sSql = sSql & "     ,SGI_CADFORNEC  FORNE" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "  And ITENS.SGI_VLPAGO IS NOT NULL" & vbCrLf
    sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
    sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
    sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
    sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODFOR " & vbCrLf
    sSql = sSql & "Order By" & vbCrLf
    sSql = sSql & "        ITENS.SGI_DTPGTO" & vbCrLf
    sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
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
                              BREC!SGI_CODFOR & vbTab & _
                              BREC!SGI_RAZAOSOC & vbTab & _
                              BREC!SGI_PARCELA & vbTab & _
                              BREC!SGI_STATUS
                               
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

Private Function VerifBaixados() As Boolean

    VerifBaixados = False
    
    '' Verifica se há baixados
    sSql = "Select" & vbCrLf
    sSql = sSql & "      * " & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_CONTASIAPG " & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "  And SGI_CODIGO = " & objCADCONTASAPG.CODPGTO & vbCrLf
    sSql = sSql & "  And SGI_VLPAGO IS NOT NULL "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
       MsgBox "Há titulos já baixado !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Function
    End If
    BREC.Close
    '' ------------------------------
    
    If flxCADCONTASAPG.TextMatrix(flxCADCONTASAPG.RowSel, 9) = "L" Then
       MsgBox "Este titulo esta incluso no lote e aguarda baixa !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Function
    Else
       
       sSql = "Select " & vbCrLf
       sSql = sSql & "       * " & vbCrLf
       sSql = sSql & "  From " & vbCrLf
       sSql = sSql & "       SGI_CADLOTEITENS " & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
       sSql = sSql & "   And SGI_CODTIT = " & objCADCONTASAPG.CODPGTO
       
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
       If Not BREC.EOF Then
          MsgBox "Existem titulos que estão inclusos em um lote !!!", vbOKOnly + vbExclamation, "Aviso"
          BREC.Close
          Exit Function
       End If
       BREC.Close
       
    End If
    
    '' ----------------------------------------------------------------
    '' Verifica se existe NF
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_NFENTRADACABEC " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODCONTAPG = " & objCADCONTASAPG.CODPGTO
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
       MsgBox "Este titulo está ligado a uma NF !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Function
    End If
    BREC.Close
    '' ----------------------------------------------------------------
    
    VerifBaixados = True

End Function

Public Sub ConfGridLote()

    flxLote.Rows = 1
    flxLote.Cols = 10
    
    flxLote.TextMatrix(0, 0) = ""
    flxLote.TextMatrix(0, 1) = "Nº Lote"
    flxLote.TextMatrix(0, 2) = "Nº Doc"
    flxLote.TextMatrix(0, 3) = "Dt. Lote"
    flxLote.TextMatrix(0, 4) = "Valor"
    flxLote.TextMatrix(0, 5) = "Tipo Pagamento"
    flxLote.TextMatrix(0, 6) = "Origen Documento"
    flxLote.TextMatrix(0, 7) = "Status"
    flxLote.TextMatrix(0, 8) = "Status"
    flxLote.TextMatrix(0, 9) = " "
    
    flxLote.ColWidth(0) = 0
    flxLote.ColWidth(1) = 1000
    flxLote.ColWidth(2) = 1000
    flxLote.ColWidth(3) = 1000
    flxLote.ColWidth(4) = 1000
    flxLote.ColWidth(5) = 2500
    flxLote.ColWidth(6) = 2500
    flxLote.ColWidth(7) = 0
    flxLote.ColWidth(8) = 750
    flxLote.ColWidth(9) = 500
    
    Label3(3).Caption = Format(0, "#,##0.00")
    Label3(2).Caption = Format(0, "#,##0.00")

End Sub

Private Sub PreenchGridLote()

    Dim strLoteDescOrig As String

    sSql = "Select " & vbCrLf
    sSql = sSql & "       LOTE.SGI_CODIGO    " & vbCrLf
    sSql = sSql & "      ,LOTE.SGI_NUMDOC    " & vbCrLf
    sSql = sSql & "      ,LOTE.SGI_DATALOTE  " & vbCrLf
    sSql = sSql & "      ,LOTE.SGI_VLDOC     " & vbCrLf
    sSql = sSql & "      ,TPGT.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "      ,LOTE.SGI_CODBCO    " & vbCrLf
    sSql = sSql & "      ,LOTE.SGI_STATUS    " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADLOTEHEADER LOTE " & vbCrLf
    sSql = sSql & "      ,SGI_CADTIPOPGTO   TPGT " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       LOTE.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And LOTE.SGI_STATUS = 'A'" & vbCrLf
    sSql = sSql & "   And TPGT.SGI_FILIAL = LOTE.SGI_FILIAL  " & vbCrLf
    sSql = sSql & "   And TPGT.SGI_CODIGO = LOTE.SGI_TIPPGTO " & vbCrLf
    sSql = sSql & "Order by LOTE.SGI_DATALOTE "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    
    Label3(2).Caption = Format(0, "#,##0.00")
    
    Do While Not BREC.EOF
    
       strLoteDescOrig = ""
       
       '' ------------------------
       sSql = "Select " & vbCrLf
       sSql = sSql & "       * " & vbCrLf
       sSql = sSql & "  From " & vbCrLf
       sSql = sSql & "       SGI_CADBANCOS " & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
       sSql = sSql & "   And SGI_CODIGO = " & BREC!SGI_CODBCO
       
       BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
       If Not BREC2.EOF Then strLoteDescOrig = BREC2!SGI_DESCRICAO
       BREC2.Close
       '' ------------------------
       
       flxLote.AddItem "" & vbTab & _
                       BREC!SGI_CODIGO & vbTab & _
                       BREC!SGI_NUMDOC & vbTab & _
                       Format(BREC!SGI_DATALOTE, "DD/MM/YYYY") & vbTab & _
                       Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                       BREC!SGI_DESCRICAO & vbTab & _
                       strLoteDescOrig & vbTab & _
                       BREC!SGI_STATUS & vbTab & _
                       IIf(BREC!SGI_STATUS = "A", "ABERTO", "")
                       
       Label3(2).Caption = Format((CCur(Label3(2).Caption) + BREC!SGI_VLDOC), "#,##0.00")
       
       BREC.MoveNext
    Loop
    
    BREC.Close

End Sub

Public Sub ConfGridLoteLiberdo()

    flxLoteLiberado.Rows = 1
    flxLoteLiberado.Cols = 10
    
    flxLoteLiberado.TextMatrix(0, 0) = ""
    flxLoteLiberado.TextMatrix(0, 1) = "Nº Lote"
    flxLoteLiberado.TextMatrix(0, 2) = "Nº Doc"
    flxLoteLiberado.TextMatrix(0, 3) = "Dt. Lote"
    flxLoteLiberado.TextMatrix(0, 4) = "Valor"
    flxLoteLiberado.TextMatrix(0, 5) = "Tipo Pagamento"
    flxLoteLiberado.TextMatrix(0, 6) = "Origen Documento"
    flxLoteLiberado.TextMatrix(0, 7) = "Status"
    flxLoteLiberado.TextMatrix(0, 8) = "Status"
    flxLoteLiberado.TextMatrix(0, 9) = ""
    
    flxLoteLiberado.ColWidth(0) = 0
    flxLoteLiberado.ColWidth(1) = 1000
    flxLoteLiberado.ColWidth(2) = 1000
    flxLoteLiberado.ColWidth(3) = 1000
    flxLoteLiberado.ColWidth(4) = 1000
    flxLoteLiberado.ColWidth(5) = 2500
    flxLoteLiberado.ColWidth(6) = 2500
    flxLoteLiberado.ColWidth(7) = 0
    flxLoteLiberado.ColWidth(8) = 900
    flxLoteLiberado.ColWidth(9) = 300

End Sub


Private Sub SomaTitSelec()

    Dim I As Integer
    
    Label3(3).Caption = Format(0, "#,##0.00")
    For I = 1 To (flxLote.Rows - 1)
        If flxLote.TextMatrix(I, 9) = "*" Then Label3(3).Caption = Format(CCur(Label3(3).Caption) + CCur(flxLote.TextMatrix(I, 4)), "#,##0.00")
    Next I

End Sub

Private Sub DesLibera()

    Dim I      As Integer
    Dim iresp  As Integer
    Dim blAcou As Boolean
    
    If flxLoteLiberado.Rows = 1 Then Exit Sub
    
    blAcou = False
    For I = 1 To (flxLoteLiberado.Rows - 1)
        If flxLoteLiberado.TextMatrix(I, 9) = "*" Then blAcou = True
    Next I
    If blAcou = False Then
       MsgBox "Não foi selecionado nenhum lote para desliberar !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    iresp = MsgBox("Deseja realmente desliberar estes lotes ?", vbYesNo + vbQuestion + vbDefaultButton1, "Aviso")
    
    If iresp = 7 Then Exit Sub
    
    objCADCONTASAPG.FILIAL = FILIAL
    
    For I = 1 To (flxLoteLiberado.Rows - 1)
        If flxLoteLiberado.TextMatrix(I, 9) = "*" Then
           objCADCONTASAPG.CODLOTE = CLng(flxLoteLiberado.TextMatrix(I, 1))
           If objCADCONTASAPG.GRAVALOTE("D") = False Then Exit Sub
           If objCADCONTASAPG.Atualiza("D", Str(objCADCONTASAPG.CODLOTE), FILIAL, "LOTE") = False Then Exit Sub
        End If
    Next I
    
    Atualiza_Grid
    SomaTitSelec

End Sub

Private Sub PreenchGridLoteLiberado()

    Dim strLoteDescOrig As String

    sSql = "Select " & vbCrLf
    sSql = sSql & "       LOTE.SGI_CODIGO    " & vbCrLf
    sSql = sSql & "      ,LOTE.SGI_NUMDOC    " & vbCrLf
    sSql = sSql & "      ,LOTE.SGI_DATALOTE  " & vbCrLf
    sSql = sSql & "      ,LOTE.SGI_VLDOC     " & vbCrLf
    sSql = sSql & "      ,TPGT.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "      ,LOTE.SGI_CODBCO    " & vbCrLf
    sSql = sSql & "      ,LOTE.SGI_STATUS    " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADLOTEHEADER LOTE " & vbCrLf
    sSql = sSql & "      ,SGI_CADTIPOPGTO   TPGT " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       LOTE.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And LOTE.SGI_STATUS = 'L'" & vbCrLf
    sSql = sSql & "   And TPGT.SGI_FILIAL = LOTE.SGI_FILIAL  " & vbCrLf
    sSql = sSql & "   And TPGT.SGI_CODIGO = LOTE.SGI_TIPPGTO " & vbCrLf
    sSql = sSql & "Order by LOTE.SGI_DATALOTE "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
    
       strLoteDescOrig = ""
       
       '' ------------------------
       sSql = "Select " & vbCrLf
       sSql = sSql & "       * " & vbCrLf
       sSql = sSql & "  From " & vbCrLf
       sSql = sSql & "       SGI_CADBANCOS " & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
       sSql = sSql & "   And SGI_CODIGO = " & BREC!SGI_CODBCO
       
       BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
       If Not BREC2.EOF Then strLoteDescOrig = BREC2!SGI_DESCRICAO
       BREC2.Close
       '' ------------------------
       
       flxLoteLiberado.AddItem "" & vbTab & _
                               BREC!SGI_CODIGO & vbTab & _
                               BREC!SGI_NUMDOC & vbTab & _
                               Format(BREC!SGI_DATALOTE, "DD/MM/YYYY") & vbTab & _
                               Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                               BREC!SGI_DESCRICAO & vbTab & _
                               strLoteDescOrig & vbTab & _
                               BREC!SGI_STATUS & vbTab & _
                               IIf(BREC!SGI_STATUS = "L", "LIBERADO", "")
                       
       BREC.MoveNext
    Loop
    
    BREC.Close

End Sub

Private Sub Libera()

    Dim I      As Integer
    Dim iresp  As Integer
    Dim blAcou As Boolean
    
    If flxLote.Rows = 1 Then Exit Sub
    
    blAcou = False
    For I = 1 To (flxLote.Rows - 1)
        If flxLote.TextMatrix(I, 9) = "*" Then blAcou = True
    Next I
    If blAcou = False Then
       MsgBox "Não foi selecionado nenhum lote para liberar !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    iresp = MsgBox("Deseja realmente liberar estes lotes ?", vbYesNo + vbQuestion + vbDefaultButton1, "Aviso")
    
    If iresp = 7 Then Exit Sub
    
    objCADCONTASAPG.FILIAL = FILIAL
    
    For I = 1 To (flxLote.Rows - 1)
        If flxLote.TextMatrix(I, 9) = "*" Then
           objCADCONTASAPG.CODLOTE = CLng(flxLote.TextMatrix(I, 1))
           If objCADCONTASAPG.GRAVALOTE("L") = False Then Exit Sub
           If objCADCONTASAPG.Atualiza("L", Str(objCADCONTASAPG.CODLOTE), FILIAL, "LOTE") = False Then Exit Sub
        End If
    Next I
    
    Atualiza_Grid
    SomaTitSelec

End Sub

Private Sub BaixaLiberados()

    Dim I      As Integer
    Dim iresp  As Integer
    Dim blAcou As Boolean
    
    If flxLoteLiberado.Rows = 1 Then Exit Sub
    
    blAcou = False
    For I = 1 To (flxLoteLiberado.Rows - 1)
        If flxLoteLiberado.TextMatrix(I, 9) = "*" Then blAcou = True
    Next I
    If blAcou = False Then
       MsgBox "Não foi selecionado nenhum lote para baixar !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    iresp = MsgBox("Deseja realmente baixar estes lotes ?", vbYesNo + vbQuestion + vbDefaultButton1, "Aviso")
    
    If iresp = 7 Then Exit Sub
    
    objCADCONTASAPG.FILIAL = FILIAL
    
    For I = 1 To (flxLoteLiberado.Rows - 1)
        If flxLoteLiberado.TextMatrix(I, 9) = "*" Then
           objCADCONTASAPG.CODLOTE = CLng(flxLoteLiberado.TextMatrix(I, 1))
           If objCADCONTASAPG.GRAVALOTE("BL") = False Then Exit Sub
           If objCADCONTASAPG.Atualiza("B", Str(objCADCONTASAPG.CODLOTE), FILIAL, "LOTE") = False Then Exit Sub
        End If
    Next I
    
    ConfGrid
    ConfGridBaixados
    ConfGridLote
    ConfGridLoteLiberdo
    ConfGridLoteBaixado
    PreencheGrid
    PopGridBaixados
    PreenchGridLote
    PreenchGridLoteLiberado
    PreenchGridLoteBaixado
    SomaTitSelec

End Sub

Private Function VerifLote() As Boolean

       VerifLote = False
       
       sSql = "Select " & vbCrLf
       sSql = sSql & "       * " & vbCrLf
       sSql = sSql & "  From " & vbCrLf
       sSql = sSql & "       SGI_CADLOTEITENS " & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
       sSql = sSql & "   And SGI_CODTIT = " & objCADCONTASAPG.CODPGTO
       
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
       If Not BREC.EOF Then
          MsgBox "Existem titulos que estão inclusos em um lote !!!", vbOKOnly + vbExclamation, "Aviso"
          BREC.Close
          Exit Function
       End If
       BREC.Close
       
       VerifLote = True

End Function

Public Sub ConfGridLoteBaixado()

    flxLoteBaixados.Rows = 1
    flxLoteBaixados.Cols = 10
    
    flxLoteBaixados.TextMatrix(0, 0) = ""
    flxLoteBaixados.TextMatrix(0, 1) = "Nº Lote"
    flxLoteBaixados.TextMatrix(0, 2) = "Nº Doc"
    flxLoteBaixados.TextMatrix(0, 3) = "Dt. Lote"
    flxLoteBaixados.TextMatrix(0, 4) = "Valor"
    flxLoteBaixados.TextMatrix(0, 5) = "Tipo Pagamento"
    flxLoteBaixados.TextMatrix(0, 6) = "Origen Documento"
    flxLoteBaixados.TextMatrix(0, 7) = "Status"
    flxLoteBaixados.TextMatrix(0, 8) = "Status"
    flxLoteBaixados.TextMatrix(0, 9) = ""
    
    flxLoteBaixados.ColWidth(0) = 0
    flxLoteBaixados.ColWidth(1) = 1000
    flxLoteBaixados.ColWidth(2) = 1000
    flxLoteBaixados.ColWidth(3) = 1000
    flxLoteBaixados.ColWidth(4) = 1000
    flxLoteBaixados.ColWidth(5) = 2500
    flxLoteBaixados.ColWidth(6) = 2500
    flxLoteBaixados.ColWidth(7) = 0
    flxLoteBaixados.ColWidth(8) = 900
    flxLoteBaixados.ColWidth(9) = 300

End Sub

Private Sub PreenchGridLoteBaixado()

    Dim strLoteDescOrig As String

    sSql = "Select " & vbCrLf
    sSql = sSql & "       LOTE.SGI_CODIGO    " & vbCrLf
    sSql = sSql & "      ,LOTE.SGI_NUMDOC    " & vbCrLf
    sSql = sSql & "      ,LOTE.SGI_DATALOTE  " & vbCrLf
    sSql = sSql & "      ,LOTE.SGI_VLDOC     " & vbCrLf
    sSql = sSql & "      ,TPGT.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "      ,LOTE.SGI_CODBCO    " & vbCrLf
    sSql = sSql & "      ,LOTE.SGI_STATUS    " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADLOTEHEADER LOTE " & vbCrLf
    sSql = sSql & "      ,SGI_CADTIPOPGTO   TPGT " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       LOTE.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And LOTE.SGI_STATUS = 'B'" & vbCrLf
    sSql = sSql & "   And TPGT.SGI_FILIAL = LOTE.SGI_FILIAL  " & vbCrLf
    sSql = sSql & "   And TPGT.SGI_CODIGO = LOTE.SGI_TIPPGTO " & vbCrLf
    sSql = sSql & "Order by LOTE.SGI_DATALOTE "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
    
       strLoteDescOrig = ""
       
       '' ------------------------
       sSql = "Select " & vbCrLf
       sSql = sSql & "       * " & vbCrLf
       sSql = sSql & "  From " & vbCrLf
       sSql = sSql & "       SGI_CADBANCOS " & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
       sSql = sSql & "   And SGI_CODIGO = " & BREC!SGI_CODBCO
       
       BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
       If Not BREC2.EOF Then strLoteDescOrig = BREC2!SGI_DESCRICAO
       BREC2.Close
       '' ------------------------
       
       flxLoteBaixados.AddItem "" & vbTab & _
                               BREC!SGI_CODIGO & vbTab & _
                               BREC!SGI_NUMDOC & vbTab & _
                               Format(BREC!SGI_DATALOTE, "DD/MM/YYYY") & vbTab & _
                               Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                               BREC!SGI_DESCRICAO & vbTab & _
                               strLoteDescOrig & vbTab & _
                               BREC!SGI_STATUS & vbTab & _
                               IIf(BREC!SGI_STATUS = "B", "BAIXADO", "")
                       
       BREC.MoveNext
    Loop
    
    BREC.Close

End Sub

Private Sub ExtornaLotePago()

    Dim I      As Integer
    Dim iresp  As Integer
    Dim blAcou As Boolean
    
    If flxLoteBaixados.Rows = 1 Then Exit Sub
    
    If VerificaCaixa = True Then Exit Sub
    
    blAcou = False
    For I = 1 To (flxLoteBaixados.Rows - 1)
        If flxLoteBaixados.TextMatrix(I, 9) = "*" Then blAcou = True
    Next I
    If blAcou = False Then
       MsgBox "Não foi selecionado nenhum lote para extornar !!!", vbOKOnly + vbCritical, "Aviso"
       Exit Sub
    End If
    
    iresp = MsgBox("Deseja realmente extornar estes lotes ?", vbYesNo + vbQuestion + vbDefaultButton1, "Aviso")
    
    If iresp = 7 Then Exit Sub
    
    objCADCONTASAPG.FILIAL = FILIAL
    
    For I = 1 To (flxLoteBaixados.Rows - 1)
        If flxLoteBaixados.TextMatrix(I, 9) = "*" Then
           objCADCONTASAPG.CODLOTE = CLng(flxLoteBaixados.TextMatrix(I, 1))
           If objCADCONTASAPG.GRAVALOTE("EL") = False Then Exit Sub
        End If
    Next I
    
    ConfGrid
    ConfGridBaixados
    ConfGridLote
    ConfGridLoteLiberdo
    ConfGridLoteBaixado
    PreencheGrid
    PopGridBaixados
    PreenchGridLote
    PreenchGridLoteLiberado
    PreenchGridLoteBaixado
    SomaTitSelec

End Sub

Private Function VerificaCaixa() As Boolean
    
    VerificaCaixa = False
    
    If StTitulos.Tab = 4 Then
    
       sSql = "Select " & vbCrLf
       sSql = sSql & "       * " & vbCrLf
       sSql = sSql & "  From " & vbCrLf
       sSql = sSql & "       SGI_CADLOTEHEADER " & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
       sSql = sSql & "   And SGI_CODIGO = " & objCADCONTASAPG.CODLOTE & vbCrLf
       sSql = sSql & "   And SGI_CODCAIXA IS NOT NULL"
    
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
       If Not BREC.EOF Then
          MsgBox "Já foi criado fluxo de caixa !!!", vbOKOnly + vbExclamation, "Aviso"
          VerificaCaixa = True
       End If
       BREC.Close
    
    End If
    
    If StTitulos.Tab = 1 Then
       sSql = "Select " & vbCrLf
       sSql = sSql & "       * " & vbCrLf
       sSql = sSql & "  From " & vbCrLf
       sSql = sSql & "       SGI_CONTASIAPG " & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
       sSql = sSql & "   And SGI_CODIGO = " & objCADCONTASAPG.CODPGTO & vbCrLf
       sSql = sSql & "   And SGI_CODCAIXA IS NOT NULL"
     
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
       If Not BREC.EOF Then
          MsgBox "Já foi criado fluxo de caixa !!!", vbOKOnly + vbExclamation, "Aviso"
          VerificaCaixa = True
       End If
       BREC.Close
    End If
    
End Function

Private Sub Implote()
    
    sSql = "Select "
    sSql = sSql & "       SGI_CONTASIAPG.SGI_NLOTE        "
    sSql = sSql & "      ,SGI_CADLOTEHEADER.SGI_DATALOTE  "
    sSql = sSql & "      ,SGI_CADLOTEHEADER.SGI_STATUS    "
    sSql = sSql & "      ,SGI_CONTASIAPG.SGI_NUMDOC       "
    sSql = sSql & "      ,SGI_CONTASIAPG.SGI_DATAVENC     "
    sSql = sSql & "      ,SGI_CONTASIAPG.SGI_VLDOC        "
    sSql = sSql & "      ,SGI_CONTASIAPG.SGI_PARCELA      "
    sSql = sSql & "      ,SGI_CONTASHAPG.SGI_CODIGO       "
    sSql = sSql & "      ,SGI_CONTASHAPG.SGI_QTDPARC      "
    sSql = sSql & "      ,SGI_CADFORNEC.SGI_RAZAOSOC      "
    
    sSql = sSql & "  From "
    sSql = sSql & "      SGI_CONTASIAPG    "
    sSql = sSql & "     ,SGI_CADLOTEHEADER "
    sSql = sSql & "     ,SGI_CONTASHAPG    "
    sSql = sSql & "     ,SGI_CADFORNEC     "
    
    sSql = sSql & " Where "
    sSql = sSql & "      SGI_CONTASIAPG.SGI_FILIAL    = " & FILIAL
    sSql = sSql & "  And SGI_CONTASIAPG.SGI_NLOTE     = " & objCADCONTASAPG.CODLOTE
    sSql = sSql & "  And SGI_CADLOTEHEADER.SGI_FILIAL = SGI_CONTASIAPG.SGI_FILIAL     "
    sSql = sSql & "  And SGI_CADLOTEHEADER.SGI_CODIGO = SGI_CONTASIAPG.SGI_NLOTE      "
    sSql = sSql & "  And SGI_CONTASHAPG.SGI_FILIAL    = SGI_CONTASIAPG.SGI_FILIAL     "
    sSql = sSql & "  And SGI_CONTASHAPG.SGI_CODIGO    = SGI_CONTASIAPG.SGI_CODIGO     "
    sSql = sSql & "  And SGI_CADFORNEC.SGI_FILIAL     = SGI_CONTASHAPG.SGI_FILIAL     "
    sSql = sSql & "  And SGI_CADFORNEC.SGI_CODIGO     = SGI_CONTASHAPG.SGI_CODFOR     "
    
    sSql = sSql & " Order by SGI_CONTASIAPG.SGI_DATAVENC "

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF Then
       MsgBox "Este lote não existe !!!", vbOKOnly + vbExclamation, "Aviso"
       BREC.Close
       Exit Sub
    End If
    BREC.Close
    
    strTitulo = "** LOTE DE PAGAMENTOS **"
    strCABEC2 = ""
    objREL.REL FILIAL, sSql, strCamRelNovo & cCamRelContasAPGLote & "RELLOTEAPG2.rpt", Linha, 1, strTitulo, strCABEC2, False

End Sub

Private Sub Atualiza_Grid()
    
     Dim I               As Integer
     Dim J               As Integer
     Dim bolAchou        As Boolean
     Dim strLoteDescOrig As String
      
     bolAchou = False
      
     sSql = "Select" & vbCrLf
     sSql = sSql & "      * " & vbCrLf
     sSql = sSql & "  From" & vbCrLf
     sSql = sSql & "       SGI_ATUALIZA" & vbCrLf
     sSql = sSql & " Where" & vbCrLf
     sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
     
     If StTitulos.Tab = 0 Then sSql = sSql & "   And SGI_MODULO = 'frmCADCONTASAPG'" & vbCrLf
     If StTitulos.Tab = 1 Then sSql = sSql & "   And SGI_MODULO = 'frmCADCONTASAPG'" & vbCrLf
     If StTitulos.Tab = 2 Then sSql = sSql & "   And SGI_MODULO = 'LOTE'" & vbCrLf
     If StTitulos.Tab = 3 Then sSql = sSql & "   And SGI_MODULO = 'LOTE'" & vbCrLf
     If StTitulos.Tab = 4 Then sSql = sSql & "   And SGI_MODULO = 'LOTE'" & vbCrLf
     
'     If StTitulos.Tab = 0 Then ConfGrid
'     If StTitulos.Tab = 1 Then ConfGridBaixados
'     If StTitulos.Tab = 2 Then ConfGridLote
'     If StTitulos.Tab = 3 Then ConfGridLoteLiberdo
'     If StTitulos.Tab = 4 Then ConfGridLoteBaixado
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     If Not BREC.EOF Then
        
        If StTitulos.Tab = 0 Then
           For I = 1 To (flxCADCONTASAPG.Rows - 1)
               If Trim(BREC!SGI_ACAO) = "E" Or Trim(BREC!SGI_ACAO) = "B" Then
                  If flxCADCONTASAPG.TextMatrix(I, 1) = Trim(BREC!SGI_CODIGO) Then
                     If flxCADCONTASAPG.Rows = 2 Then flxCADCONTASAPG.Rows = 1
                     If flxCADCONTASAPG.Rows > 2 Then flxCADCONTASAPG.RemoveItem I
                     Exit For
                  End If
               ElseIf Trim(BREC!SGI_ACAO) = "I" Or Trim(BREC!SGI_ACAO) = "A" Or Trim(BREC!SGI_ACAO) = "X" Then
                  If Trim(BREC!SGI_CODIGO) = Trim(flxCADCONTASAPG.TextMatrix(I, 1)) Then
                     bolAchou = True
                     Exit For
                  End If
               End If
           Next I
        
        ElseIf StTitulos.Tab = 1 Then
           
           For I = 1 To (flxTitBaixados.Rows - 1)
               If Trim(BREC!SGI_ACAO) = "E" Or Trim(BREC!SGI_ACAO) = "X" Then
                  If flxTitBaixados.TextMatrix(I, 1) = Trim(BREC!SGI_CODIGO) Then
                     If flxTitBaixados.Rows = 2 Then flxTitBaixados.Rows = 1
                     If flxTitBaixados.Rows > 2 Then flxTitBaixados.RemoveItem I
                     Exit For
                  End If
               ElseIf Trim(BREC!SGI_ACAO) = "I" Or Trim(BREC!SGI_ACAO) = "A" Or Trim(BREC!SGI_ACAO) = "B" Then
                  If Trim(BREC!SGI_CODIGO) = Trim(flxTitBaixados.TextMatrix(I, 1)) Then
                     bolAchou = True
                     Exit For
                  End If
               End If
           Next I
        
        ElseIf StTitulos.Tab = 2 Then
            
           For I = 1 To (flxLote.Rows - 1)
               If Trim(BREC!SGI_ACAO) = "L" Then
                  If flxLote.TextMatrix(I, 1) = Trim(BREC!SGI_CODIGO) Then
                     If flxLote.Rows = 2 Then flxLote.Rows = 1
                     If flxLote.Rows > 2 Then flxLote.RemoveItem I
                     Exit For
                  End If
               ElseIf Trim(BREC!SGI_ACAO) = "I" Or Trim(BREC!SGI_ACAO) = "A" Or Trim(BREC!SGI_ACAO) = "D" Then
                  If Trim(BREC!SGI_CODIGO) = Trim(flxLote.TextMatrix(I, 1)) Then
                     bolAchou = True
                     Exit For
                  End If
               End If
           Next I
        
        ElseIf StTitulos.Tab = 3 Then
            
           For I = 1 To (flxLoteLiberado.Rows - 1)
               If Trim(BREC!SGI_ACAO) = "D" Then
                  If flxLoteLiberado.TextMatrix(I, 1) = Trim(BREC!SGI_CODIGO) Then
                     If flxLoteLiberado.Rows = 2 Then flxLoteLiberado.Rows = 1
                     If flxLoteLiberado.Rows > 2 Then flxLoteLiberado.RemoveItem I
                     Exit For
                  End If
               ElseIf Trim(BREC!SGI_ACAO) = "L" Or Trim(BREC!SGI_ACAO) = "A" Then
                  If Trim(BREC!SGI_CODIGO) = Trim(flxLoteLiberado.TextMatrix(I, 1)) Then
                     bolAchou = True
                     Exit For
                  End If
               End If
           Next I
        
        ElseIf StTitulos.Tab = 4 Then
            
           For I = 1 To (flxLoteBaixados.Rows - 1)
               If Trim(BREC!SGI_ACAO) = "E" Then
                  If flxLoteBaixados.TextMatrix(I, 1) = Trim(BREC!SGI_CODIGO) Then
                     If flxLoteBaixados.Rows = 2 Then flxLoteBaixados.Rows = 1
                     If flxLoteBaixados.Rows > 2 Then flxLoteBaixados.RemoveItem I
                     Exit For
                  End If
               ElseIf Trim(BREC!SGI_ACAO) = "B" Or Trim(BREC!SGI_ACAO) = "A" Then
                  If Trim(BREC!SGI_CODIGO) = Trim(flxLoteBaixados.TextMatrix(I, 1)) Then
                     bolAchou = True
                     Exit For
                  End If
               End If
           Next I
        
        End If
        
        If (bolAchou = False And Trim(BREC!SGI_ACAO) = "I") Or _
           (bolAchou = False And Trim(BREC!SGI_ACAO) = "B") Or _
           (bolAchou = False And Trim(BREC!SGI_ACAO) = "X") Or _
           (bolAchou = False And Trim(BREC!SGI_ACAO) = "D") Or _
           (bolAchou = False And Trim(BREC!SGI_ACAO) = "L") Then
           
           If StTitulos.Tab = 0 Then
              If Trim(BREC!SGI_ACAO) = "I" Or Trim(BREC!SGI_ACAO) = "X" Then
                 sSql = "Select" & vbCrLf
                 sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
                 sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
                 sSql = sSql & "     ,CABEC.SGI_CODFOR " & vbCrLf
                 sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
                 sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
                 sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
                 sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
                 sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
                 sSql = sSql & "     ,ITENS.SGI_STATUS  " & vbCrLf
                 sSql = sSql & "  From" & vbCrLf
                 sSql = sSql & "      SGI_CONTASIAPG ITENS" & vbCrLf
                 sSql = sSql & "     ,SGI_CONTASHAPG CABEC" & vbCrLf
                 sSql = sSql & "     ,SGI_CADFORNEC  FORNE" & vbCrLf
                 sSql = sSql & " Where" & vbCrLf
                 sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
                 sSql = sSql & "  And ITENS.SGI_VLPAGO IS NULL" & vbCrLf
                 sSql = sSql & "  And ITENS.SGI_CODIGO = " & Trim(BREC!SGI_CODIGO) & vbCrLf
                 sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
                 sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
                 sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
                 sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODFOR " & vbCrLf
                 sSql = sSql & "Order By" & vbCrLf
                 If cboFiltro.ListIndex = 0 Then
                    sSql = sSql & "        ITENS.SGI_CODIGO" & vbCrLf
                    sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
                    sSql = sSql & "       ,ITENS.SGI_PARCELA"
                 ElseIf cboFiltro.ListIndex = 1 Then
                    sSql = sSql & "        ITENS.SGI_NUMDOC" & vbCrLf
                    sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
                    sSql = sSql & "       ,ITENS.SGI_PARCELA"
                 ElseIf cboFiltro.ListIndex = 2 Then
                    sSql = sSql & "        FORNE.SGI_RAZAOSOC" & vbCrLf
                    sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
                    sSql = sSql & "       ,ITENS.SGI_PARCELA"
                 ElseIf cboFiltro.ListIndex = 3 Then
                    sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
                    sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
                    sSql = sSql & "       ,ITENS.SGI_PARCELA"
                 End If
              ElseIf Trim(BREC!SGI_ACAO) = "B" Then
                 sSql = "Select" & vbCrLf
                 sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
                 sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
                 sSql = sSql & "     ,CABEC.SGI_CODFOR " & vbCrLf
                 sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
                 sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
                 sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
                 sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
                 sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
                 sSql = sSql & "     ,ITENS.SGI_STATUS  " & vbCrLf
                 sSql = sSql & "  From" & vbCrLf
                 sSql = sSql & "      SGI_CONTASIAPG ITENS" & vbCrLf
                 sSql = sSql & "     ,SGI_CONTASHAPG CABEC" & vbCrLf
                 sSql = sSql & "     ,SGI_CADFORNEC  FORNE" & vbCrLf
                 sSql = sSql & " Where" & vbCrLf
                 sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
                 sSql = sSql & "  And ITENS.SGI_VLPAGO IS NULL" & vbCrLf
                 sSql = sSql & "  And ITENS.SGI_CODIGO = " & Trim(BREC!SGI_CODIGO) & vbCrLf
                 sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
                 sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
                 sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
                 sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODFOR " & vbCrLf
                 sSql = sSql & "Order By" & vbCrLf
                 If cboFiltro.ListIndex = 0 Then
                    sSql = sSql & "        ITENS.SGI_CODIGO" & vbCrLf
                    sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
                    sSql = sSql & "       ,ITENS.SGI_PARCELA"
                 ElseIf cboFiltro.ListIndex = 1 Then
                    sSql = sSql & "        ITENS.SGI_NUMDOC" & vbCrLf
                    sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
                    sSql = sSql & "       ,ITENS.SGI_PARCELA"
                 ElseIf cboFiltro.ListIndex = 2 Then
                    sSql = sSql & "        FORNE.SGI_RAZAOSOC" & vbCrLf
                    sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
                    sSql = sSql & "       ,ITENS.SGI_PARCELA"
                 ElseIf cboFiltro.ListIndex = 3 Then
                    sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
                    sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
                    sSql = sSql & "       ,ITENS.SGI_PARCELA"
                 End If
              End If
           
           ElseIf StTitulos.Tab = 1 Then
              If Trim(BREC!SGI_ACAO) = "I" Or Trim(BREC!SGI_ACAO) = "X" Then
           
                  sSql = "Select" & vbCrLf
                  sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
                  sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
                  sSql = sSql & "     ,CABEC.SGI_CODFOR " & vbCrLf
                  sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
                  sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
                  sSql = sSql & "     ,ITENS.SGI_DTPGTO" & vbCrLf
                  sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
                  sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
                  sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
                  sSql = sSql & "     ,ITENS.SGI_STATUS  " & vbCrLf
                  sSql = sSql & "  From" & vbCrLf
                  sSql = sSql & "      SGI_CONTASIAPG ITENS" & vbCrLf
                  sSql = sSql & "     ,SGI_CONTASHAPG CABEC" & vbCrLf
                  sSql = sSql & "     ,SGI_CADFORNEC  FORNE" & vbCrLf
                  sSql = sSql & " Where" & vbCrLf
                  sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
                  sSql = sSql & "  And ITENS.SGI_VLPAGO IS NULL" & vbCrLf
                  sSql = sSql & "  And ITENS.SGI_CODIGO = " & Trim(BREC!SGI_CODIGO) & vbCrLf
                  sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
                  sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
                  sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
                  sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODFOR " & vbCrLf
                  sSql = sSql & "Order By" & vbCrLf
                  If cboFiltro.ListIndex = 0 Then
                     sSql = sSql & "        ITENS.SGI_CODIGO" & vbCrLf
                     sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
                     sSql = sSql & "       ,ITENS.SGI_PARCELA"
                     sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
                  ElseIf cboFiltro.ListIndex = 1 Then
                     sSql = sSql & "        ITENS.SGI_NUMDOC" & vbCrLf
                     sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
                     sSql = sSql & "       ,ITENS.SGI_PARCELA"
                     sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
                  ElseIf cboFiltro.ListIndex = 2 Then
                     sSql = sSql & "        FORNE.SGI_RAZAOSOC" & vbCrLf
                     sSql = sSql & "       ,ITENS.SGI_PARCELA"
                     sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
                  ElseIf cboFiltro.ListIndex = 3 Then
                     sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
                     sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
                     sSql = sSql & "       ,ITENS.SGI_PARCELA"
                  ElseIf cboFiltro.ListIndex = 4 Then
                     sSql = sSql & "        ITENS.SGI_DTPGTO" & vbCrLf
                     sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
                     sSql = sSql & "       ,ITENS.SGI_PARCELA"
                  End If
           
              ElseIf Trim(BREC!SGI_ACAO) = "B" Then
           
                  sSql = "Select" & vbCrLf
                  sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
                  sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
                  sSql = sSql & "     ,CABEC.SGI_CODFOR " & vbCrLf
                  sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
                  sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
                  sSql = sSql & "     ,ITENS.SGI_DTPGTO" & vbCrLf
                  sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
                  sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
                  sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
                  sSql = sSql & "  From" & vbCrLf
                  sSql = sSql & "      SGI_CONTASIAPG ITENS" & vbCrLf
                  sSql = sSql & "     ,SGI_CONTASHAPG CABEC" & vbCrLf
                  sSql = sSql & "     ,SGI_CADFORNEC  FORNE" & vbCrLf
                  sSql = sSql & " Where" & vbCrLf
                  sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
                  sSql = sSql & "  And ITENS.SGI_VLPAGO IS NOT NULL" & vbCrLf
                  sSql = sSql & "  aND ITENS.SGI_CODIGO = " & Trim(BREC!SGI_CODIGO) & vbCrLf
                  sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
                  sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
                  sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
                  sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODFOR " & vbCrLf
                  sSql = sSql & "Order By" & vbCrLf
                  If cboFiltro.ListIndex = 0 Then
                     sSql = sSql & "        ITENS.SGI_CODIGO" & vbCrLf
                     sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
                     sSql = sSql & "       ,ITENS.SGI_PARCELA"
                     sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
                  ElseIf cboFiltro.ListIndex = 1 Then
                     sSql = sSql & "        ITENS.SGI_NUMDOC" & vbCrLf
                     sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
                     sSql = sSql & "       ,ITENS.SGI_PARCELA"
                     sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
                  ElseIf cboFiltro.ListIndex = 2 Then
                     sSql = sSql & "        FORNE.SGI_RAZAOSOC" & vbCrLf
                     sSql = sSql & "       ,ITENS.SGI_PARCELA"
                     sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
                  ElseIf cboFiltro.ListIndex = 3 Then
                     sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
                     sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
                     sSql = sSql & "       ,ITENS.SGI_PARCELA"
                  ElseIf cboFiltro.ListIndex = 4 Then
                     sSql = sSql & "        ITENS.SGI_DTPGTO" & vbCrLf
                     sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
                     sSql = sSql & "       ,ITENS.SGI_PARCELA"
                  End If
           
              End If
           ElseIf StTitulos.Tab = 2 Or StTitulos.Tab = 3 Then
              If Trim(BREC!SGI_ACAO) = "I" Or Trim(BREC!SGI_ACAO) = "D" Then
    
                 sSql = "Select " & vbCrLf
                 sSql = sSql & "       LOTE.SGI_CODIGO    " & vbCrLf
                 sSql = sSql & "      ,LOTE.SGI_NUMDOC    " & vbCrLf
                 sSql = sSql & "      ,LOTE.SGI_DATALOTE  " & vbCrLf
                 sSql = sSql & "      ,LOTE.SGI_VLDOC     " & vbCrLf
                 sSql = sSql & "      ,TPGT.SGI_DESCRICAO " & vbCrLf
                 sSql = sSql & "      ,LOTE.SGI_CODBCO    " & vbCrLf
                 sSql = sSql & "      ,LOTE.SGI_STATUS    " & vbCrLf
                 sSql = sSql & "  From " & vbCrLf
                 sSql = sSql & "       SGI_CADLOTEHEADER LOTE " & vbCrLf
                 sSql = sSql & "      ,SGI_CADTIPOPGTO   TPGT " & vbCrLf
                 sSql = sSql & " Where " & vbCrLf
                 sSql = sSql & "       LOTE.SGI_FILIAL = " & FILIAL & vbCrLf
                 sSql = sSql & "   And LOTE.SGI_STATUS = 'A'" & vbCrLf
                 sSql = sSql & "   And LOTE.SGI_CODIGO = " & Trim(BREC!SGI_CODIGO) & vbCrLf
                 sSql = sSql & "   And TPGT.SGI_FILIAL = LOTE.SGI_FILIAL  " & vbCrLf
                 sSql = sSql & "   And TPGT.SGI_CODIGO = LOTE.SGI_TIPPGTO " & vbCrLf
                 sSql = sSql & "Order by " & vbCrLf
                 If cboFiltro.ListIndex = 0 Then
                    sSql = sSql & "         LOTE.SGI_CODIGO "
                 ElseIf cboFiltro.ListIndex = 1 Then
                    sSql = sSql & "         LOTE.SGI_NUMDOC "
                 ElseIf cboFiltro.ListIndex = 2 Then
                    sSql = sSql & "         LOTE.SGI_DATALOTE "
                 End If
           
              ElseIf Trim(BREC!SGI_ACAO) = "L" Then
              
                 sSql = "Select " & vbCrLf
                 sSql = sSql & "       LOTE.SGI_CODIGO    " & vbCrLf
                 sSql = sSql & "      ,LOTE.SGI_NUMDOC    " & vbCrLf
                 sSql = sSql & "      ,LOTE.SGI_DATALOTE  " & vbCrLf
                 sSql = sSql & "      ,LOTE.SGI_VLDOC     " & vbCrLf
                 sSql = sSql & "      ,TPGT.SGI_DESCRICAO " & vbCrLf
                 sSql = sSql & "      ,LOTE.SGI_CODBCO    " & vbCrLf
                 sSql = sSql & "      ,LOTE.SGI_STATUS    " & vbCrLf
                 sSql = sSql & "  From " & vbCrLf
                 sSql = sSql & "       SGI_CADLOTEHEADER LOTE " & vbCrLf
                 sSql = sSql & "      ,SGI_CADTIPOPGTO   TPGT " & vbCrLf
                 sSql = sSql & " Where " & vbCrLf
                 sSql = sSql & "       LOTE.SGI_FILIAL = " & FILIAL & vbCrLf
                 sSql = sSql & "   And LOTE.SGI_STATUS = 'L'" & vbCrLf
                 sSql = sSql & "   And LOTE.SGI_CODIGO = " & Trim(BREC!SGI_CODIGO) & vbCrLf
                 sSql = sSql & "   And TPGT.SGI_FILIAL = LOTE.SGI_FILIAL  " & vbCrLf
                 sSql = sSql & "   And TPGT.SGI_CODIGO = LOTE.SGI_TIPPGTO " & vbCrLf
                 sSql = sSql & "Order by " & vbCrLf
                 If cboFiltro.ListIndex = 0 Then
                    sSql = sSql & "         LOTE.SGI_CODIGO "
                 ElseIf cboFiltro.ListIndex = 1 Then
                    sSql = sSql & "         LOTE.SGI_NUMDOC "
                 ElseIf cboFiltro.ListIndex = 2 Then
                    sSql = sSql & "         LOTE.SGI_DATALOTE "
                 End If
              
              End If
           End If
           
           BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC2.EOF Then
                                  
              If StTitulos.Tab = 0 Then
                 If Trim(BREC!SGI_ACAO) = "I" Or Trim(BREC!SGI_ACAO) = "X" Then
                    flxCADCONTASAPG.AddItem "" & vbTab & _
                                            BREC2!SGI_CODIGO & vbTab & _
                                            BREC2!SGI_NUMDOC & vbTab & _
                                            Format(BREC2!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                                            Format(BREC2!SGI_PARCELA, "##00") & "/" & Format(BREC2!SGI_QTDPARC, "##00") & vbTab & _
                                            Format(BREC2!SGI_VLDOC, "#,##0.00") & vbTab & _
                                            BREC2!SGI_CODFOR & vbTab & _
                                            BREC2!SGI_RAZAOSOC & vbTab & _
                                            BREC2!SGI_PARCELA & vbTab & _
                                            BREC2!SGI_STATUS
                 
                 ElseIf Trim(BREC!SGI_ACAO) = "B" Then
                    flxCADCONTASAPG.AddItem "" & vbTab & _
                                            BREC2!SGI_CODIGO & vbTab & _
                                            BREC2!SGI_NUMDOC & vbTab & _
                                            Format(BREC2!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                                            Format(BREC2!SGI_PARCELA, "##00") & "/" & Format(BREC2!SGI_QTDPARC, "##00") & vbTab & _
                                            Format(BREC2!SGI_VLDOC, "#,##0.00") & vbTab & _
                                            BREC2!SGI_CODFOR & vbTab & _
                                            BREC2!SGI_RAZAOSOC & vbTab & _
                                            BREC2!SGI_PARCELA & vbTab & _
                                            BREC2!SGI_STATUS
                 End If
              ElseIf StTitulos.Tab = 1 Then
              
                 If Trim(BREC!SGI_ACAO) = "I" Or Trim(BREC!SGI_ACAO) = "X" Then
               
                    flxTitBaixados.AddItem "" & vbTab & _
                                   BREC2!SGI_CODIGO & vbTab & _
                                   BREC2!SGI_NUMDOC & vbTab & _
                                   Format(BREC2!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                                   Format(BREC2!SGI_DTPGTO, "DD/MM/YYYY") & vbTab & _
                                   Format(BREC2!SGI_PARCELA, "##00") & "/" & Format(BREC2!SGI_QTDPARC, "##00") & vbTab & _
                                   Format(BREC2!SGI_VLDOC, "#,##0.00") & vbTab & _
                                   BREC2!SGI_CODFOR & vbTab & _
                                   BREC2!SGI_RAZAOSOC & vbTab & _
                                   BREC2!SGI_PARCELA & vbTab & _
                                   BREC2!SGI_STATUS
                                  
                 ElseIf Trim(BREC!SGI_ACAO) = "B" Then
              
                    flxTitBaixados.AddItem "" & vbTab & _
                                   BREC2!SGI_CODIGO & vbTab & _
                                   BREC2!SGI_NUMDOC & vbTab & _
                                   Format(BREC2!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                                   Format(BREC2!SGI_DTPGTO, "DD/MM/YYYY") & vbTab & _
                                   Format(BREC2!SGI_PARCELA, "##00") & "/" & Format(BREC2!SGI_QTDPARC, "##00") & vbTab & _
                                   Format(BREC2!SGI_VLDOC, "#,##0.00") & vbTab & _
                                   BREC2!SGI_CODFOR & vbTab & _
                                   BREC2!SGI_RAZAOSOC & vbTab & _
                                   BREC2!SGI_PARCELA & vbTab & _
                                   ""
              
                 End If
              ElseIf StTitulos.Tab = 2 Or StTitulos.Tab = 3 Then
              
                  If Trim(BREC!SGI_ACAO) = "I" Or Trim(BREC!SGI_ACAO) = "D" Then
                  
                     Label3(2).Caption = Format(0, "#,##0.00")

                     Do While Not BREC2.EOF
    
                        strLoteDescOrig = ""
       
                        '' ------------------------
                        sSql = "Select " & vbCrLf
                        sSql = sSql & "       * " & vbCrLf
                        sSql = sSql & "  From " & vbCrLf
                        sSql = sSql & "       SGI_CADBANCOS " & vbCrLf
                        sSql = sSql & " Where " & vbCrLf
                        sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
                        sSql = sSql & "   And SGI_CODIGO = " & BREC2!SGI_CODBCO
       
                        BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
                        If Not BREC3.EOF Then strLoteDescOrig = BREC3!SGI_DESCRICAO
                        BREC3.Close
                        '' ------------------------
       
                        flxLote.AddItem "" & vbTab & _
                                        BREC2!SGI_CODIGO & vbTab & _
                                        BREC2!SGI_NUMDOC & vbTab & _
                                        Format(BREC2!SGI_DATALOTE, "DD/MM/YYYY") & vbTab & _
                                        Format(BREC2!SGI_VLDOC, "#,##0.00") & vbTab & _
                                        BREC2!SGI_DESCRICAO & vbTab & _
                                        strLoteDescOrig & vbTab & _
                                        BREC2!SGI_STATUS & vbTab & _
                                        IIf(BREC2!SGI_STATUS = "A", "ABERTO", "")
                       
                        Label3(2).Caption = Format((CCur(Label3(2).Caption) + BREC2!SGI_VLDOC), "#,##0.00")
       
                        BREC2.MoveNext
                     
                     Loop
                  
                  ElseIf Trim(BREC!SGI_ACAO) = "L" Then
                  
                     Do While Not BREC2.EOF
    
                        strLoteDescOrig = ""
       
                        '' ------------------------
                        sSql = "Select " & vbCrLf
                        sSql = sSql & "       * " & vbCrLf
                        sSql = sSql & "  From " & vbCrLf
                        sSql = sSql & "       SGI_CADBANCOS " & vbCrLf
                        sSql = sSql & " Where " & vbCrLf
                        sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
                        sSql = sSql & "   And SGI_CODIGO = " & BREC2!SGI_CODBCO
       
                        BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
                        If Not BREC3.EOF Then strLoteDescOrig = BREC3!SGI_DESCRICAO
                        BREC3.Close
                        '' ------------------------
       
                        flxLoteLiberado.AddItem "" & vbTab & _
                                        BREC2!SGI_CODIGO & vbTab & _
                                        BREC2!SGI_NUMDOC & vbTab & _
                                        Format(BREC2!SGI_DATALOTE, "DD/MM/YYYY") & vbTab & _
                                        Format(BREC2!SGI_VLDOC, "#,##0.00") & vbTab & _
                                        BREC2!SGI_DESCRICAO & vbTab & _
                                        strLoteDescOrig & vbTab & _
                                        BREC2!SGI_STATUS & vbTab & _
                                        IIf(BREC2!SGI_STATUS = "L", "LIBERADO", "")
                       
                        BREC2.MoveNext
                     Loop
                  
                  End If
                  
              End If
              
           End If
           BREC2.Close
        
        ElseIf (bolAchou = True And BREC!SGI_ACAO = "A") Or _
               (bolAchou = True And BREC!SGI_ACAO = "B") Or _
               (bolAchou = True And BREC!SGI_ACAO = "X") Or _
               (bolAchou = True And BREC!SGI_ACAO = "I") Or _
               (bolAchou = True And BREC!SGI_ACAO = "L") Or _
               (bolAchou = True And BREC!SGI_ACAO = "D") Then
        
            If StTitulos.Tab = 0 Then
           
               sSql = "Select" & vbCrLf
               sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
               sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
               sSql = sSql & "     ,CABEC.SGI_CODFOR " & vbCrLf
               sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
               sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
               sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
               sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
               sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
               sSql = sSql & "     ,ITENS.SGI_STATUS  " & vbCrLf
               sSql = sSql & "  From" & vbCrLf
               sSql = sSql & "      SGI_CONTASIAPG ITENS" & vbCrLf
               sSql = sSql & "     ,SGI_CONTASHAPG CABEC" & vbCrLf
               sSql = sSql & "     ,SGI_CADFORNEC  FORNE" & vbCrLf
               sSql = sSql & " Where" & vbCrLf
               sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
               sSql = sSql & "  And ITENS.SGI_VLPAGO IS NULL" & vbCrLf
               sSql = sSql & "  And ITENS.SGI_CODIGO = " & Trim(BREC!SGI_CODIGO) & vbCrLf
               sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
               sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
               sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
               sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODFOR " & vbCrLf
               sSql = sSql & "Order By" & vbCrLf
               If cboFiltro.ListIndex = 0 Then
                  sSql = sSql & "        ITENS.SGI_CODIGO" & vbCrLf
                  sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
                  sSql = sSql & "       ,ITENS.SGI_PARCELA"
               ElseIf cboFiltro.ListIndex = 1 Then
                  sSql = sSql & "        ITENS.SGI_NUMDOC" & vbCrLf
                  sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
                  sSql = sSql & "       ,ITENS.SGI_PARCELA"
               ElseIf cboFiltro.ListIndex = 2 Then
                  sSql = sSql & "        FORNE.SGI_RAZAOSOC" & vbCrLf
                  sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
                  sSql = sSql & "       ,ITENS.SGI_PARCELA"
               ElseIf cboFiltro.ListIndex = 3 Then
                  sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
                  sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
                  sSql = sSql & "       ,ITENS.SGI_PARCELA"
               End If
           
            ElseIf StTitulos.Tab = 1 Then
           
               sSql = "Select" & vbCrLf
               sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
               sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
               sSql = sSql & "     ,CABEC.SGI_CODFOR " & vbCrLf
               sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
               sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
               sSql = sSql & "     ,ITENS.SGI_DTPGTO" & vbCrLf
               sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
               sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
               sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
               sSql = sSql & "     ,ITENS.SGI_STATUS  " & vbCrLf
               sSql = sSql & "  From" & vbCrLf
               sSql = sSql & "      SGI_CONTASIAPG ITENS" & vbCrLf
               sSql = sSql & "     ,SGI_CONTASHAPG CABEC" & vbCrLf
               sSql = sSql & "     ,SGI_CADFORNEC  FORNE" & vbCrLf
               sSql = sSql & " Where" & vbCrLf
               sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
               sSql = sSql & "  And ITENS.SGI_VLPAGO IS NOT NULL" & vbCrLf
               sSql = sSql & "  aND ITENS.SGI_CODIGO = " & Trim(BREC!SGI_CODIGO) & vbCrLf
               sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
               sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
               sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
               sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODFOR " & vbCrLf
               sSql = sSql & "Order By" & vbCrLf
                If cboFiltro.ListIndex = 0 Then
                   sSql = sSql & "        ITENS.SGI_CODIGO" & vbCrLf
                   sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
                   sSql = sSql & "       ,ITENS.SGI_PARCELA"
                   sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
                ElseIf cboFiltro.ListIndex = 1 Then
                   sSql = sSql & "        ITENS.SGI_NUMDOC" & vbCrLf
                   sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
                   sSql = sSql & "       ,ITENS.SGI_PARCELA"
                   sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
                ElseIf cboFiltro.ListIndex = 2 Then
                   sSql = sSql & "        FORNE.SGI_RAZAOSOC" & vbCrLf
                   sSql = sSql & "       ,ITENS.SGI_PARCELA"
                   sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
                ElseIf cboFiltro.ListIndex = 3 Then
                   sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
                   sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
                   sSql = sSql & "       ,ITENS.SGI_PARCELA"
                ElseIf cboFiltro.ListIndex = 4 Then
                   sSql = sSql & "        ITENS.SGI_DTPGTO" & vbCrLf
                   sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
                   sSql = sSql & "       ,ITENS.SGI_PARCELA"
                End If
            
            ElseIf StTitulos.Tab = 2 Then
            
               sSql = "Select " & vbCrLf
               sSql = sSql & "       LOTE.SGI_CODIGO    " & vbCrLf
               sSql = sSql & "      ,LOTE.SGI_NUMDOC    " & vbCrLf
               sSql = sSql & "      ,LOTE.SGI_DATALOTE  " & vbCrLf
               sSql = sSql & "      ,LOTE.SGI_VLDOC     " & vbCrLf
               sSql = sSql & "      ,TPGT.SGI_DESCRICAO " & vbCrLf
               sSql = sSql & "      ,LOTE.SGI_CODBCO    " & vbCrLf
               sSql = sSql & "      ,LOTE.SGI_STATUS    " & vbCrLf
               sSql = sSql & "  From " & vbCrLf
               sSql = sSql & "       SGI_CADLOTEHEADER LOTE " & vbCrLf
               sSql = sSql & "      ,SGI_CADTIPOPGTO   TPGT " & vbCrLf
               sSql = sSql & " Where " & vbCrLf
               sSql = sSql & "       LOTE.SGI_FILIAL = " & FILIAL & vbCrLf
               sSql = sSql & "   And LOTE.SGI_STATUS = 'A'" & vbCrLf
               sSql = sSql & "   And LOTE.SGI_CODIGO = " & Trim(BREC!SGI_CODIGO) & vbCrLf
               sSql = sSql & "   And TPGT.SGI_FILIAL = LOTE.SGI_FILIAL  " & vbCrLf
               sSql = sSql & "   And TPGT.SGI_CODIGO = LOTE.SGI_TIPPGTO " & vbCrLf
               sSql = sSql & "Order by " & vbCrLf
               If cboFiltro.ListIndex = 0 Then
                  sSql = sSql & "         LOTE.SGI_CODIGO "
               ElseIf cboFiltro.ListIndex = 1 Then
                  sSql = sSql & "         LOTE.SGI_NUMDOC "
               ElseIf cboFiltro.ListIndex = 2 Then
                  sSql = sSql & "         LOTE.SGI_DATALOTE "
               End If
            
            ElseIf StTitulos.Tab = 3 Then
            
               sSql = "Select " & vbCrLf
               sSql = sSql & "       LOTE.SGI_CODIGO    " & vbCrLf
               sSql = sSql & "      ,LOTE.SGI_NUMDOC    " & vbCrLf
               sSql = sSql & "      ,LOTE.SGI_DATALOTE  " & vbCrLf
               sSql = sSql & "      ,LOTE.SGI_VLDOC     " & vbCrLf
               sSql = sSql & "      ,TPGT.SGI_DESCRICAO " & vbCrLf
               sSql = sSql & "      ,LOTE.SGI_CODBCO    " & vbCrLf
               sSql = sSql & "      ,LOTE.SGI_STATUS    " & vbCrLf
               sSql = sSql & "  From " & vbCrLf
               sSql = sSql & "       SGI_CADLOTEHEADER LOTE " & vbCrLf
               sSql = sSql & "      ,SGI_CADTIPOPGTO   TPGT " & vbCrLf
               sSql = sSql & " Where " & vbCrLf
               sSql = sSql & "       LOTE.SGI_FILIAL = " & FILIAL & vbCrLf
               sSql = sSql & "   And LOTE.SGI_STATUS = 'L'" & vbCrLf
               sSql = sSql & "   And LOTE.SGI_CODIGO = " & Trim(BREC!SGI_CODIGO) & vbCrLf
               sSql = sSql & "   And TPGT.SGI_FILIAL = LOTE.SGI_FILIAL  " & vbCrLf
               sSql = sSql & "   And TPGT.SGI_CODIGO = LOTE.SGI_TIPPGTO " & vbCrLf
               sSql = sSql & "Order by " & vbCrLf
               If cboFiltro.ListIndex = 0 Then
                  sSql = sSql & "         LOTE.SGI_CODIGO "
               ElseIf cboFiltro.ListIndex = 1 Then
                  sSql = sSql & "         LOTE.SGI_NUMDOC "
               ElseIf cboFiltro.ListIndex = 2 Then
                  sSql = sSql & "         LOTE.SGI_DATALOTE "
               End If
            
            ElseIf StTitulos.Tab = 4 Then
            
               sSql = "Select " & vbCrLf
               sSql = sSql & "       LOTE.SGI_CODIGO    " & vbCrLf
               sSql = sSql & "      ,LOTE.SGI_NUMDOC    " & vbCrLf
               sSql = sSql & "      ,LOTE.SGI_DATALOTE  " & vbCrLf
               sSql = sSql & "      ,LOTE.SGI_VLDOC     " & vbCrLf
               sSql = sSql & "      ,TPGT.SGI_DESCRICAO " & vbCrLf
               sSql = sSql & "      ,LOTE.SGI_CODBCO    " & vbCrLf
               sSql = sSql & "      ,LOTE.SGI_STATUS    " & vbCrLf
               sSql = sSql & "  From " & vbCrLf
               sSql = sSql & "       SGI_CADLOTEHEADER LOTE " & vbCrLf
               sSql = sSql & "      ,SGI_CADTIPOPGTO   TPGT " & vbCrLf
               sSql = sSql & " Where " & vbCrLf
               sSql = sSql & "       LOTE.SGI_FILIAL = " & FILIAL & vbCrLf
               sSql = sSql & "   And LOTE.SGI_STATUS = 'B'" & vbCrLf
               sSql = sSql & "   And LOTE.SGI_CODIGO = " & Trim(BREC!SGI_CODIGO) & vbCrLf
               sSql = sSql & "   And TPGT.SGI_FILIAL = LOTE.SGI_FILIAL  " & vbCrLf
               sSql = sSql & "   And TPGT.SGI_CODIGO = LOTE.SGI_TIPPGTO " & vbCrLf
               sSql = sSql & "Order by " & vbCrLf
               If cboFiltro.ListIndex = 0 Then
                  sSql = sSql & "         LOTE.SGI_CODIGO "
               ElseIf cboFiltro.ListIndex = 1 Then
                  sSql = sSql & "         LOTE.SGI_NUMDOC "
               ElseIf cboFiltro.ListIndex = 2 Then
                  sSql = sSql & "         LOTE.SGI_DATALOTE "
               End If
            
            End If
           
            BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
            If Not BREC2.EOF Then
             
               If StTitulos.Tab = 0 Then
                  
                  flxCADCONTASAPG.TextMatrix(I, 0) = ""
                  flxCADCONTASAPG.TextMatrix(I, 1) = BREC2!SGI_CODIGO
                  flxCADCONTASAPG.TextMatrix(I, 2) = BREC2!SGI_NUMDOC
                  flxCADCONTASAPG.TextMatrix(I, 3) = Format(BREC2!SGI_DATAVENC, "DD/MM/YYYY")
                  flxCADCONTASAPG.TextMatrix(I, 4) = Format(BREC2!SGI_PARCELA, "##00") & "/" & Format(BREC2!SGI_QTDPARC, "##00")
                  flxCADCONTASAPG.TextMatrix(I, 5) = Format(BREC2!SGI_VLDOC, "#,##0.00")
                  flxCADCONTASAPG.TextMatrix(I, 6) = BREC2!SGI_CODFOR
                  flxCADCONTASAPG.TextMatrix(I, 7) = BREC2!SGI_RAZAOSOC
                  flxCADCONTASAPG.TextMatrix(I, 8) = BREC2!SGI_PARCELA
                  flxCADCONTASAPG.TextMatrix(I, 9) = BREC2!SGI_STATUS
               
               ElseIf StTitulos.Tab = 1 Then
               
                  flxTitBaixados.TextMatrix(I, 0) = ""
                  flxTitBaixados.TextMatrix(I, 1) = BREC2!SGI_CODIGO
                  flxTitBaixados.TextMatrix(I, 2) = BREC2!SGI_NUMDOC
                  flxTitBaixados.TextMatrix(I, 3) = Format(BREC2!SGI_DATAVENC, "DD/MM/YYYY")
                  flxTitBaixados.TextMatrix(I, 4) = Format(BREC2!SGI_DTPGTO, "DD/MM/YYYY")
                  flxTitBaixados.TextMatrix(I, 5) = Format(BREC2!SGI_PARCELA, "##00") & "/" & Format(BREC2!SGI_QTDPARC, "##00")
                  flxTitBaixados.TextMatrix(I, 6) = Format(BREC2!SGI_VLDOC, "#,##0.00")
                  flxTitBaixados.TextMatrix(I, 7) = BREC2!SGI_CODFOR
                  flxTitBaixados.TextMatrix(I, 8) = BREC2!SGI_RAZAOSOC
                  flxTitBaixados.TextMatrix(I, 9) = BREC2!SGI_PARCELA
                  flxTitBaixados.TextMatrix(I, 10) = BREC2!SGI_STATUS
               
               ElseIf StTitulos.Tab = 2 Then
               
                  Label3(2).Caption = Format(0, "#,##0.00")
                  
                  Do While Not BREC2.EOF
    
                     strLoteDescOrig = ""
       
                     '' ------------------------
                     sSql = "Select " & vbCrLf
                     sSql = sSql & "       * " & vbCrLf
                     sSql = sSql & "  From " & vbCrLf
                     sSql = sSql & "       SGI_CADBANCOS " & vbCrLf
                     sSql = sSql & " Where " & vbCrLf
                     sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
                     sSql = sSql & "   And SGI_CODIGO = " & BREC2!SGI_CODBCO
       
                     BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
                     If Not BREC3.EOF Then strLoteDescOrig = BREC3!SGI_DESCRICAO
                     BREC3.Close
                     '' ------------------------
       
                     flxLote.TextMatrix(I, 0) = ""
                     flxLote.TextMatrix(I, 1) = BREC2!SGI_CODIGO
                     flxLote.TextMatrix(I, 2) = BREC2!SGI_NUMDOC
                     flxLote.TextMatrix(I, 3) = Format(BREC2!SGI_DATALOTE, "DD/MM/YYYY")
                     flxLote.TextMatrix(I, 4) = Format(BREC2!SGI_VLDOC, "#,##0.00")
                     flxLote.TextMatrix(I, 5) = BREC2!SGI_DESCRICAO
                     flxLote.TextMatrix(I, 6) = strLoteDescOrig
                     flxLote.TextMatrix(I, 7) = BREC2!SGI_STATUS
                     flxLote.TextMatrix(I, 8) = IIf(BREC2!SGI_STATUS = "A", "ABERTO", "")
                       
                     Label3(2).Caption = Format((CCur(Label3(2).Caption) + BREC2!SGI_VLDOC), "#,##0.00")
       
                     BREC2.MoveNext
                  
                  Loop
               
               ElseIf StTitulos.Tab = 3 Then
               
                  Do While Not BREC2.EOF
    
                     strLoteDescOrig = ""
       
                     '' ------------------------
                     sSql = "Select " & vbCrLf
                     sSql = sSql & "       * " & vbCrLf
                     sSql = sSql & "  From " & vbCrLf
                     sSql = sSql & "       SGI_CADBANCOS " & vbCrLf
                     sSql = sSql & " Where " & vbCrLf
                     sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
                     sSql = sSql & "   And SGI_CODIGO = " & BREC2!SGI_CODBCO
       
                     BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
                     If Not BREC3.EOF Then strLoteDescOrig = BREC3!SGI_DESCRICAO
                     BREC3.Close
                     '' ------------------------
       
                     flxLoteLiberado.TextMatrix(I, 0) = ""
                     flxLoteLiberado.TextMatrix(I, 1) = BREC2!SGI_CODIGO
                     flxLoteLiberado.TextMatrix(I, 2) = BREC2!SGI_NUMDOC
                     flxLoteLiberado.TextMatrix(I, 3) = Format(BREC2!SGI_DATALOTE, "DD/MM/YYYY")
                     flxLoteLiberado.TextMatrix(I, 4) = Format(BREC2!SGI_VLDOC, "#,##0.00")
                     flxLoteLiberado.TextMatrix(I, 5) = BREC2!SGI_DESCRICAO
                     flxLoteLiberado.TextMatrix(I, 6) = strLoteDescOrig
                     flxLoteLiberado.TextMatrix(I, 7) = BREC2!SGI_STATUS
                     flxLoteLiberado.TextMatrix(I, 8) = IIf(BREC2!SGI_STATUS = "L", "LIBERADO", "")
                       
                     BREC2.MoveNext
                  Loop
               
               ElseIf StTitulos.Tab = 4 Then
               
                     strLoteDescOrig = ""
       
                     '' ------------------------
                     sSql = "Select " & vbCrLf
                     sSql = sSql & "       * " & vbCrLf
                     sSql = sSql & "  From " & vbCrLf
                     sSql = sSql & "       SGI_CADBANCOS " & vbCrLf
                     sSql = sSql & " Where " & vbCrLf
                     sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
                     sSql = sSql & "   And SGI_CODIGO = " & BREC2!SGI_CODBCO
       
                     BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
                     If Not BREC3.EOF Then strLoteDescOrig = BREC3!SGI_DESCRICAO
                     BREC3.Close
       
       
                     flxLoteBaixados.TextMatrix(I, 0) = ""
                     flxLoteBaixados.TextMatrix(I, 1) = BREC2!SGI_CODIGO
                     flxLoteBaixados.TextMatrix(I, 2) = BREC2!SGI_NUMDOC
                     flxLoteBaixados.TextMatrix(I, 3) = Format(BREC2!SGI_DATALOTE, "DD/MM/YYYY")
                     flxLoteBaixados.TextMatrix(I, 4) = Format(BREC2!SGI_VLDOC, "#,##0.00")
                     flxLoteBaixados.TextMatrix(I, 5) = BREC2!SGI_DESCRICAO
                     flxLoteBaixados.TextMatrix(I, 6) = strLoteDescOrig
                     flxLoteBaixados.TextMatrix(I, 7) = BREC2!SGI_STATUS
                     flxLoteBaixados.TextMatrix(I, 8) = IIf(BREC2!SGI_STATUS = "B", "BAIXADO", "")
               
               End If
             
           End If
           BREC2.Close
        
        End If
        
     End If
     BREC.Close
      
End Sub

Private Function Verif_reg() As Boolean

    Verif_reg = False
    
    If StTitulos.Tab = 0 Then
       sSql = "Select " & vbCrLf
       sSql = sSql & "       * " & vbCrLf
       sSql = sSql & "  From " & vbCrLf
       sSql = sSql & "       SGI_CONTASHAPG " & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
       sSql = sSql & "   And SGI_CODIGO = " & objCADCONTASAPG.CODPGTO
    
    ElseIf StTitulos.Tab = 1 Then
    
       sSql = "Select" & vbCrLf
       sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
       sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
       sSql = sSql & "     ,CABEC.SGI_CODFOR " & vbCrLf
       sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
       sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
       sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
       sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
       sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
       sSql = sSql & "  From" & vbCrLf
       sSql = sSql & "      SGI_CONTASIAPG ITENS" & vbCrLf
       sSql = sSql & "     ,SGI_CONTASHAPG CABEC" & vbCrLf
       sSql = sSql & "     ,SGI_CADFORNEC  FORNE" & vbCrLf
       sSql = sSql & " Where" & vbCrLf
       sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
       sSql = sSql & "  And ITENS.SGI_CODIGO = " & objCADCONTASAPG.CODPGTO & vbCrLf
       sSql = sSql & "  And ITENS.SGI_VLPAGO IS NOT NULL" & vbCrLf
       sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
       sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
       sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
       sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODFOR " & vbCrLf
       sSql = sSql & "Order By" & vbCrLf
       sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
       sSql = sSql & "       ,CABEC.SGI_CODFOR" & vbCrLf
       sSql = sSql & "       ,ITENS.SGI_PARCELA"
    
    ElseIf StTitulos.Tab = 2 Then
    
       sSql = "Select " & vbCrLf
       sSql = sSql & "       LOTE.SGI_CODIGO    " & vbCrLf
       sSql = sSql & "      ,LOTE.SGI_NUMDOC    " & vbCrLf
       sSql = sSql & "      ,LOTE.SGI_DATALOTE  " & vbCrLf
       sSql = sSql & "      ,LOTE.SGI_VLDOC     " & vbCrLf
       sSql = sSql & "      ,TPGT.SGI_DESCRICAO " & vbCrLf
       sSql = sSql & "      ,LOTE.SGI_CODBCO    " & vbCrLf
       sSql = sSql & "      ,LOTE.SGI_STATUS    " & vbCrLf
       sSql = sSql & "  From " & vbCrLf
       sSql = sSql & "       SGI_CADLOTEHEADER LOTE " & vbCrLf
       sSql = sSql & "      ,SGI_CADTIPOPGTO   TPGT " & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       LOTE.SGI_FILIAL = " & FILIAL & vbCrLf
       sSql = sSql & "   And LOTE.SGI_CODIGO = " & objCADCONTASAPG.CODLOTE & vbCrLf
       sSql = sSql & "   And LOTE.SGI_STATUS = 'A'" & vbCrLf
       sSql = sSql & "   And TPGT.SGI_FILIAL = LOTE.SGI_FILIAL  " & vbCrLf
       sSql = sSql & "   And TPGT.SGI_CODIGO = LOTE.SGI_TIPPGTO " & vbCrLf
       sSql = sSql & "Order by LOTE.SGI_DATALOTE "
       
    ElseIf StTitulos.Tab = 3 Then
    
       sSql = "Select " & vbCrLf
       sSql = sSql & "       LOTE.SGI_CODIGO    " & vbCrLf
       sSql = sSql & "      ,LOTE.SGI_NUMDOC    " & vbCrLf
       sSql = sSql & "      ,LOTE.SGI_DATALOTE  " & vbCrLf
       sSql = sSql & "      ,LOTE.SGI_VLDOC     " & vbCrLf
       sSql = sSql & "      ,TPGT.SGI_DESCRICAO " & vbCrLf
       sSql = sSql & "      ,LOTE.SGI_CODBCO    " & vbCrLf
       sSql = sSql & "      ,LOTE.SGI_STATUS    " & vbCrLf
       sSql = sSql & "  From " & vbCrLf
       sSql = sSql & "       SGI_CADLOTEHEADER LOTE " & vbCrLf
       sSql = sSql & "      ,SGI_CADTIPOPGTO   TPGT " & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       LOTE.SGI_FILIAL = " & FILIAL & vbCrLf
       sSql = sSql & "   And LOTE.SGI_CODIGO = " & objCADCONTASAPG.CODLOTE & vbCrLf
       sSql = sSql & "   And LOTE.SGI_STATUS = 'L'" & vbCrLf
       sSql = sSql & "   And TPGT.SGI_FILIAL = LOTE.SGI_FILIAL  " & vbCrLf
       sSql = sSql & "   And TPGT.SGI_CODIGO = LOTE.SGI_TIPPGTO " & vbCrLf
       sSql = sSql & "Order by LOTE.SGI_DATALOTE "
    
    ElseIf StTitulos.Tab = 4 Then
    
        sSql = "Select " & vbCrLf
        sSql = sSql & "       LOTE.SGI_CODIGO    " & vbCrLf
        sSql = sSql & "      ,LOTE.SGI_NUMDOC    " & vbCrLf
        sSql = sSql & "      ,LOTE.SGI_DATALOTE  " & vbCrLf
        sSql = sSql & "      ,LOTE.SGI_VLDOC     " & vbCrLf
        sSql = sSql & "      ,TPGT.SGI_DESCRICAO " & vbCrLf
        sSql = sSql & "      ,LOTE.SGI_CODBCO    " & vbCrLf
        sSql = sSql & "      ,LOTE.SGI_STATUS    " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADLOTEHEADER LOTE " & vbCrLf
        sSql = sSql & "      ,SGI_CADTIPOPGTO   TPGT " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       LOTE.SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And LOTE.SGI_CODIGO = " & objCADCONTASAPG.CODLOTE & vbCrLf
        sSql = sSql & "   And LOTE.SGI_STATUS = 'B'" & vbCrLf
        sSql = sSql & "   And TPGT.SGI_FILIAL = LOTE.SGI_FILIAL  " & vbCrLf
        sSql = sSql & "   And TPGT.SGI_CODIGO = LOTE.SGI_TIPPGTO " & vbCrLf
        sSql = sSql & "Order by LOTE.SGI_DATALOTE "
    
    End If
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF Then
       MsgBox "Este registro foi excluso !!!", vbOKOnly + vbExclamation, "Aviso"
       Verif_reg = True
    End If
    BREC.Close

End Function

Private Sub StTitulos_Click(PreviousTab As Integer)

    cboFiltro.Clear
    If StTitulos.Tab = 0 Then
       cboFiltro.AddItem "Código"
       cboFiltro.AddItem "Nº Doc"
       cboFiltro.AddItem "Fornecedor"
       cboFiltro.AddItem "Dt. Vencto"
    ElseIf StTitulos.Tab = 1 Then
       cboFiltro.AddItem "Código"
       cboFiltro.AddItem "Nº Doc"
       cboFiltro.AddItem "Fornecedor"
       cboFiltro.AddItem "Dt. Vencto"
       cboFiltro.AddItem "Dt. Pagto"
    ElseIf StTitulos.Tab = 2 Then
       cboFiltro.AddItem "Código"
       cboFiltro.AddItem "Nº Lote"
       cboFiltro.AddItem "Dt. Lote"
    ElseIf StTitulos.Tab = 3 Then
       cboFiltro.AddItem "Código"
       cboFiltro.AddItem "Nº Lote"
       cboFiltro.AddItem "Dt. Lote"
    ElseIf StTitulos.Tab = 4 Then
       cboFiltro.AddItem "Código"
       cboFiltro.AddItem "Nº Lote"
       cboFiltro.AddItem "Dt. Lote"
    End If
    cboFiltro.ListIndex = 0
    
End Sub

Private Sub Timer1_Timer()
    AbilitaCampos
    Atualiza_Grid
End Sub

Private Sub Ordem()
  
  txtCampos.Text = ""
  Dim strLoteDescOrig As String
  
  sSql = ""
  
    sSql = "Select" & vbCrLf
    sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
    sSql = sSql & "     ,CABEC.SGI_CODFOR " & vbCrLf
    sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_DTPGTO" & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
    sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
    sSql = sSql & "     ,ITENS.SGI_STATUS  " & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_CONTASIAPG ITENS" & vbCrLf
    sSql = sSql & "     ,SGI_CONTASHAPG CABEC" & vbCrLf
    sSql = sSql & "     ,SGI_CADFORNEC  FORNE" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
    
    If StTitulos.Tab = 0 Then
        sSql = sSql & "  And ITENS.SGI_VLPAGO IS NULL" & vbCrLf
    ElseIf StTitulos.Tab = 1 Then
        sSql = sSql & "  And ITENS.SGI_VLPAGO IS NOT NULL" & vbCrLf
    End If
    
    sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
    sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
    sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
    sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODFOR " & vbCrLf
    sSql = sSql & "Order By" & vbCrLf
  
  
  If StTitulos.Tab = 0 Then
     
     ConfGrid
     
     If cboFiltro.ListIndex = 0 Then
        sSql = sSql & "        ITENS.SGI_CODIGO" & vbCrLf
        sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
        sSql = sSql & "       ,ITENS.SGI_PARCELA"
     ElseIf cboFiltro.ListIndex = 1 Then
        sSql = sSql & "        ITENS.SGI_NUMDOC" & vbCrLf
        sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
        sSql = sSql & "       ,ITENS.SGI_PARCELA"
     ElseIf cboFiltro.ListIndex = 2 Then
        sSql = sSql & "        FORNE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
        sSql = sSql & "       ,ITENS.SGI_PARCELA"
     ElseIf cboFiltro.ListIndex = 3 Then
        sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
        sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "       ,ITENS.SGI_PARCELA"
     End If
  ElseIf StTitulos.Tab = 1 Then
     
     ConfGridBaixados
     
     If cboFiltro.ListIndex = 0 Then
        sSql = sSql & "        ITENS.SGI_CODIGO" & vbCrLf
        sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "       ,ITENS.SGI_PARCELA"
        sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
     ElseIf cboFiltro.ListIndex = 1 Then
        sSql = sSql & "        ITENS.SGI_NUMDOC" & vbCrLf
        sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "       ,ITENS.SGI_PARCELA"
        sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
     ElseIf cboFiltro.ListIndex = 2 Then
        sSql = sSql & "        FORNE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "       ,ITENS.SGI_PARCELA"
        sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
     ElseIf cboFiltro.ListIndex = 3 Then
        sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
        sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "       ,ITENS.SGI_PARCELA"
     ElseIf cboFiltro.ListIndex = 4 Then
        sSql = sSql & "        ITENS.SGI_DTPGTO" & vbCrLf
        sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
        sSql = sSql & "       ,ITENS.SGI_PARCELA"
     End If
  ElseIf StTitulos.Tab = 2 Then
     
     ConfGridLote
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       LOTE.SGI_CODIGO    " & vbCrLf
     sSql = sSql & "      ,LOTE.SGI_NUMDOC    " & vbCrLf
     sSql = sSql & "      ,LOTE.SGI_DATALOTE  " & vbCrLf
     sSql = sSql & "      ,LOTE.SGI_VLDOC     " & vbCrLf
     sSql = sSql & "      ,TPGT.SGI_DESCRICAO " & vbCrLf
     sSql = sSql & "      ,LOTE.SGI_CODBCO    " & vbCrLf
     sSql = sSql & "      ,LOTE.SGI_STATUS    " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADLOTEHEADER LOTE " & vbCrLf
     sSql = sSql & "      ,SGI_CADTIPOPGTO   TPGT " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       LOTE.SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & "   And LOTE.SGI_STATUS = 'A'" & vbCrLf
     sSql = sSql & "   And TPGT.SGI_FILIAL = LOTE.SGI_FILIAL  " & vbCrLf
     sSql = sSql & "   And TPGT.SGI_CODIGO = LOTE.SGI_TIPPGTO " & vbCrLf
     sSql = sSql & "Order by " & vbCrLf
     If cboFiltro.ListIndex = 0 Then
        sSql = sSql & "         LOTE.SGI_CODIGO "
     ElseIf cboFiltro.ListIndex = 1 Then
        sSql = sSql & "         LOTE.SGI_NUMDOC "
     ElseIf cboFiltro.ListIndex = 2 Then
        sSql = sSql & "         LOTE.SGI_DATALOTE "
     End If
     
     Label3(2).Caption = Format(0, "#,##0.00")
     
  ElseIf StTitulos.Tab = 3 Then
     ConfGridLoteLiberdo
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       LOTE.SGI_CODIGO    " & vbCrLf
     sSql = sSql & "      ,LOTE.SGI_NUMDOC    " & vbCrLf
     sSql = sSql & "      ,LOTE.SGI_DATALOTE  " & vbCrLf
     sSql = sSql & "      ,LOTE.SGI_VLDOC     " & vbCrLf
     sSql = sSql & "      ,TPGT.SGI_DESCRICAO " & vbCrLf
     sSql = sSql & "      ,LOTE.SGI_CODBCO    " & vbCrLf
     sSql = sSql & "      ,LOTE.SGI_STATUS    " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADLOTEHEADER LOTE " & vbCrLf
     sSql = sSql & "      ,SGI_CADTIPOPGTO   TPGT " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       LOTE.SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & "   And LOTE.SGI_STATUS = 'L'" & vbCrLf
     sSql = sSql & "   And TPGT.SGI_FILIAL = LOTE.SGI_FILIAL  " & vbCrLf
     sSql = sSql & "   And TPGT.SGI_CODIGO = LOTE.SGI_TIPPGTO " & vbCrLf
     sSql = sSql & "Order by " & vbCrLf
     If cboFiltro.ListIndex = 0 Then
        sSql = sSql & "         LOTE.SGI_CODIGO "
     ElseIf cboFiltro.ListIndex = 1 Then
        sSql = sSql & "         LOTE.SGI_NUMDOC "
     ElseIf cboFiltro.ListIndex = 2 Then
        sSql = sSql & "         LOTE.SGI_DATALOTE "
     End If
     
  ElseIf StTitulos.Tab = 4 Then
     ConfGridLoteBaixado
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       LOTE.SGI_CODIGO    " & vbCrLf
     sSql = sSql & "      ,LOTE.SGI_NUMDOC    " & vbCrLf
     sSql = sSql & "      ,LOTE.SGI_DATALOTE  " & vbCrLf
     sSql = sSql & "      ,LOTE.SGI_VLDOC     " & vbCrLf
     sSql = sSql & "      ,TPGT.SGI_DESCRICAO " & vbCrLf
     sSql = sSql & "      ,LOTE.SGI_CODBCO    " & vbCrLf
     sSql = sSql & "      ,LOTE.SGI_STATUS    " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADLOTEHEADER LOTE " & vbCrLf
     sSql = sSql & "      ,SGI_CADTIPOPGTO   TPGT " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       LOTE.SGI_FILIAL = " & FILIAL & vbCrLf
     sSql = sSql & "   And LOTE.SGI_STATUS = 'B'" & vbCrLf
     sSql = sSql & "   And TPGT.SGI_FILIAL = LOTE.SGI_FILIAL  " & vbCrLf
     sSql = sSql & "   And TPGT.SGI_CODIGO = LOTE.SGI_TIPPGTO " & vbCrLf
     sSql = sSql & "Order by " & vbCrLf
     
     If cboFiltro.ListIndex = 0 Then
        sSql = sSql & "         LOTE.SGI_CODIGO "
     ElseIf cboFiltro.ListIndex = 1 Then
        sSql = sSql & "         LOTE.SGI_NUMDOC "
     ElseIf cboFiltro.ListIndex = 2 Then
        sSql = sSql & "         LOTE.SGI_DATALOTE "
     End If
     
  End If
  
  BREC.Open sSql, adoBanco_Dados
    
  Do While Not BREC.EOF
     If StTitulos.Tab = 0 Then
        flxCADCONTASAPG.AddItem "" & vbTab & _
                                BREC!SGI_CODIGO & vbTab & _
                                BREC!SGI_NUMDOC & vbTab & _
                                Format(BREC!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                                Format(BREC!SGI_PARCELA, "##00") & "/" & Format(BREC!SGI_QTDPARC, "##00") & vbTab & _
                                Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                                BREC!SGI_CODFOR & vbTab & _
                                BREC!SGI_RAZAOSOC & vbTab & _
                                BREC!SGI_PARCELA & vbTab & _
                                BREC!SGI_STATUS
                                
     ElseIf StTitulos.Tab = 1 Then
        flxTitBaixados.AddItem "" & vbTab & _
                               BREC!SGI_CODIGO & vbTab & _
                               BREC!SGI_NUMDOC & vbTab & _
                               Format(BREC!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                               Format(BREC!SGI_DTPGTO, "DD/MM/YYYY") & vbTab & _
                               Format(BREC!SGI_PARCELA, "##00") & "/" & Format(BREC!SGI_QTDPARC, "##00") & vbTab & _
                               Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                               BREC!SGI_CODFOR & vbTab & _
                               BREC!SGI_RAZAOSOC & vbTab & _
                               BREC!SGI_PARCELA & vbTab & _
                               BREC!SGI_STATUS
                              
     ElseIf StTitulos.Tab = 2 Then
      
        strLoteDescOrig = ""
       
        '' ------------------------
        sSql = "Select " & vbCrLf
        sSql = sSql & "       * " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADBANCOS " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_CODIGO = " & BREC!SGI_CODBCO
        
        BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC2.EOF Then strLoteDescOrig = BREC2!SGI_DESCRICAO
        BREC2.Close
        '' ------------------------
        
        flxLote.AddItem "" & vbTab & _
                        BREC!SGI_CODIGO & vbTab & _
                        BREC!SGI_NUMDOC & vbTab & _
                        Format(BREC!SGI_DATALOTE, "DD/MM/YYYY") & vbTab & _
                        Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                        BREC!SGI_DESCRICAO & vbTab & _
                        strLoteDescOrig & vbTab & _
                        BREC!SGI_STATUS & vbTab & _
                        IIf(BREC!SGI_STATUS = "A", "ABERTO", "")
                       
        Label3(2).Caption = Format((CCur(Label3(2).Caption) + BREC!SGI_VLDOC), "#,##0.00")
       
     ElseIf StTitulos.Tab = 3 Then
     
        strLoteDescOrig = ""
       
        '' ------------------------
        sSql = "Select " & vbCrLf
        sSql = sSql & "       * " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADBANCOS " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_CODIGO = " & BREC!SGI_CODBCO
       
        BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC2.EOF Then strLoteDescOrig = BREC2!SGI_DESCRICAO
        BREC2.Close
        '' ------------------------
     
        flxLoteLiberado.AddItem "" & vbTab & _
                                BREC!SGI_CODIGO & vbTab & _
                                BREC!SGI_NUMDOC & vbTab & _
                                Format(BREC!SGI_DATALOTE, "DD/MM/YYYY") & vbTab & _
                                Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                                BREC!SGI_DESCRICAO & vbTab & _
                                strLoteDescOrig & vbTab & _
                                BREC!SGI_STATUS & vbTab & _
                                IIf(BREC!SGI_STATUS = "L", "LIBERADO", "")
                                
     ElseIf StTitulos.Tab = 4 Then
        
         strLoteDescOrig = ""
       
         '' ------------------------
         sSql = "Select " & vbCrLf
         sSql = sSql & "       * " & vbCrLf
         sSql = sSql & "  From " & vbCrLf
         sSql = sSql & "       SGI_CADBANCOS " & vbCrLf
         sSql = sSql & " Where " & vbCrLf
         sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
         sSql = sSql & "   And SGI_CODIGO = " & BREC!SGI_CODBCO
       
         BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
         If Not BREC2.EOF Then strLoteDescOrig = BREC2!SGI_DESCRICAO
         BREC2.Close
         '' ------------------------
       
         flxLoteBaixados.AddItem "" & vbTab & _
                                 BREC!SGI_CODIGO & vbTab & _
                                 BREC!SGI_NUMDOC & vbTab & _
                                 Format(BREC!SGI_DATALOTE, "DD/MM/YYYY") & vbTab & _
                                 Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                                 BREC!SGI_DESCRICAO & vbTab & _
                                 strLoteDescOrig & vbTab & _
                                 BREC!SGI_STATUS & vbTab & _
                                 IIf(BREC!SGI_STATUS = "B", "BAIXADO", "")
     End If
     BREC.MoveNext
     
  Loop
  BREC.Close

End Sub


Private Sub txtCampos_GotFocus()
    objFuncoes.SelecionaCampos txtCampos.Name, frmCADCONTASAPGP
End Sub

Private Sub txtCampos_KeyPress(KeyAscii As Integer)
    KeyAscii = objFuncoes.Maiuscula(KeyAscii)
End Sub

Private Sub txtCampos_Validate(Cancel As Boolean)

    Dim intQtdFiltros As Integer
    Dim strLoteDescOrig As String
    
    If Len(Trim(txtCampos.Text)) = 0 Then Exit Sub
    
    sSql = ""
    
    intQtdFiltros = cboFiltro.ListIndex
    
    If intQtdFiltros = 0 Then
       
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
          sSql = sSql & "     ,CABEC.SGI_CODFOR " & vbCrLf
          sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
          sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_STATUS  " & vbCrLf
          sSql = sSql & "  From" & vbCrLf
          sSql = sSql & "      SGI_CONTASIAPG ITENS" & vbCrLf
          sSql = sSql & "     ,SGI_CONTASHAPG CABEC" & vbCrLf
          sSql = sSql & "     ,SGI_CADFORNEC  FORNE" & vbCrLf
          sSql = sSql & " Where" & vbCrLf
          sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
          sSql = sSql & "  And ITENS.SGI_VLPAGO IS NULL" & vbCrLf
          sSql = sSql & "  And ITENS.SGI_CODIGO = " & Trim(txtCampos.Text) & vbCrLf
          sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
          sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
          sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
          sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODFOR " & vbCrLf
          sSql = sSql & "Order By" & vbCrLf
          If intQtdFiltros = 0 Then
             sSql = sSql & "        ITENS.SGI_CODIGO" & vbCrLf
             sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
             sSql = sSql & "       ,ITENS.SGI_PARCELA"
          ElseIf intQtdFiltros = 1 Then
             sSql = sSql & "        ITENS.SGI_NUMDOC" & vbCrLf
             sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
             sSql = sSql & "       ,ITENS.SGI_PARCELA"
          ElseIf intQtdFiltros = 2 Then
             sSql = sSql & "        FORNE.SGI_RAZAOSOC" & vbCrLf
             sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
             sSql = sSql & "       ,ITENS.SGI_PARCELA"
          ElseIf intQtdFiltros = 3 Then
             sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
             sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
             sSql = sSql & "       ,ITENS.SGI_PARCELA"
          End If
       ElseIf StTitulos.Tab = 1 Then
           sSql = "Select" & vbCrLf
           sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
           sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
           sSql = sSql & "     ,CABEC.SGI_CODFOR " & vbCrLf
           sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
           sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
           sSql = sSql & "     ,ITENS.SGI_DTPGTO" & vbCrLf
           sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
           sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
           sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
           sSql = sSql & "     ,ITENS.SGI_STATUS  " & vbCrLf
           sSql = sSql & "  From" & vbCrLf
           sSql = sSql & "      SGI_CONTASIAPG ITENS" & vbCrLf
           sSql = sSql & "     ,SGI_CONTASHAPG CABEC" & vbCrLf
           sSql = sSql & "     ,SGI_CADFORNEC  FORNE" & vbCrLf
           sSql = sSql & " Where" & vbCrLf
           sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
           sSql = sSql & "  And ITENS.SGI_VLPAGO IS NOT NULL" & vbCrLf
           sSql = sSql & "  And ITENS.SGI_CODIGO = " & Trim(txtCampos.Text) & vbCrLf
           sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
           sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
           sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
           sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODFOR " & vbCrLf
           sSql = sSql & "Order By" & vbCrLf
           If intQtdFiltros = 0 Then
              sSql = sSql & "        ITENS.SGI_CODIGO" & vbCrLf
              sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_PARCELA"
              sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
           ElseIf intQtdFiltros = 1 Then
              sSql = sSql & "        ITENS.SGI_NUMDOC" & vbCrLf
              sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_PARCELA"
              sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
           ElseIf intQtdFiltros = 2 Then
              sSql = sSql & "        FORNE.SGI_RAZAOSOC" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_PARCELA"
              sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
           ElseIf intQtdFiltros = 3 Then
              sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
              sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_PARCELA"
           ElseIf intQtdFiltros = 4 Then
              sSql = sSql & "        ITENS.SGI_DTPGTO" & vbCrLf
              sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_PARCELA"
           End If
       ElseIf StTitulos.Tab = 2 Then
       
           sSql = "Select " & vbCrLf
           sSql = sSql & "       LOTE.SGI_CODIGO    " & vbCrLf
           sSql = sSql & "      ,LOTE.SGI_NUMDOC    " & vbCrLf
           sSql = sSql & "      ,LOTE.SGI_DATALOTE  " & vbCrLf
           sSql = sSql & "      ,LOTE.SGI_VLDOC     " & vbCrLf
           sSql = sSql & "      ,TPGT.SGI_DESCRICAO " & vbCrLf
           sSql = sSql & "      ,LOTE.SGI_CODBCO    " & vbCrLf
           sSql = sSql & "      ,LOTE.SGI_STATUS    " & vbCrLf
           sSql = sSql & "  From " & vbCrLf
           sSql = sSql & "       SGI_CADLOTEHEADER LOTE " & vbCrLf
           sSql = sSql & "      ,SGI_CADTIPOPGTO   TPGT " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       LOTE.SGI_FILIAL = " & FILIAL & vbCrLf
           sSql = sSql & "   And LOTE.SGI_STATUS = 'A'" & vbCrLf
           sSql = sSql & "   And LOTE.SGI_CODIGO = " & Trim(txtCampos.Text) & vbCrLf
           sSql = sSql & "   And TPGT.SGI_FILIAL = LOTE.SGI_FILIAL  " & vbCrLf
           sSql = sSql & "   And TPGT.SGI_CODIGO = LOTE.SGI_TIPPGTO " & vbCrLf
           sSql = sSql & "Order by " & vbCrLf
           If cboFiltro.ListIndex = 0 Then
              sSql = sSql & "         LOTE.SGI_CODIGO "
           ElseIf cboFiltro.ListIndex = 1 Then
              sSql = sSql & "         LOTE.SGI_NUMDOC "
           ElseIf cboFiltro.ListIndex = 2 Then
              sSql = sSql & "         TPGT.SGI_DESCRICAO "
           ElseIf cboFiltro.ListIndex = 3 Then
              sSql = sSql & "         LOTE.SGI_DATALOTE "
           End If
     
           Label3(2).Caption = Format(0, "#,##0.00")
       
       ElseIf StTitulos.Tab = 3 Then
       ElseIf StTitulos.Tab = 4 Then
       End If
       
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
            
          If StTitulos.Tab = 0 Then ConfGrid
          If StTitulos.Tab = 1 Then ConfGridBaixados
          If StTitulos.Tab = 2 Then ConfGridLote
              
          Do While Not BREC.EOF
             If StTitulos.Tab = 0 Then
                flxCADCONTASAPG.AddItem "" & vbTab & _
                                        BREC!SGI_CODIGO & vbTab & _
                                        BREC!SGI_NUMDOC & vbTab & _
                                        Format(BREC!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                                        Format(BREC!SGI_PARCELA, "##00") & "/" & Format(BREC!SGI_QTDPARC, "##00") & vbTab & _
                                        Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                                        BREC!SGI_CODFOR & vbTab & _
                                        BREC!SGI_RAZAOSOC & vbTab & _
                                        BREC!SGI_PARCELA & vbTab & _
                                        BREC!SGI_STATUS
             
             ElseIf StTitulos.Tab = 1 Then
                flxTitBaixados.AddItem "" & vbTab & _
                                       BREC!SGI_CODIGO & vbTab & _
                                       BREC!SGI_NUMDOC & vbTab & _
                                       Format(BREC!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                                       Format(BREC!SGI_DTPGTO, "DD/MM/YYYY") & vbTab & _
                                       Format(BREC!SGI_PARCELA, "##00") & "/" & Format(BREC!SGI_QTDPARC, "##00") & vbTab & _
                                       Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                                       BREC!SGI_CODFOR & vbTab & _
                                       BREC!SGI_RAZAOSOC & vbTab & _
                                       BREC!SGI_PARCELA & vbTab & _
                                       BREC!SGI_STATUS
             
             ElseIf StTitulos.Tab = 2 Then
             
                strLoteDescOrig = ""
                
                '' ------------------------
                sSql = "Select " & vbCrLf
                sSql = sSql & "       * " & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_CADBANCOS " & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
                sSql = sSql & "   And SGI_CODIGO = " & BREC!SGI_CODBCO
                 
                BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
                If Not BREC2.EOF Then strLoteDescOrig = BREC2!SGI_DESCRICAO
                BREC2.Close
                '' ------------------------
                 
                flxLote.AddItem "" & vbTab & _
                                BREC!SGI_CODIGO & vbTab & _
                                BREC!SGI_NUMDOC & vbTab & _
                                Format(BREC!SGI_DATALOTE, "DD/MM/YYYY") & vbTab & _
                                Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                                BREC!SGI_DESCRICAO & vbTab & _
                                strLoteDescOrig & vbTab & _
                                BREC!SGI_STATUS & vbTab & _
                                IIf(BREC!SGI_STATUS = "A", "ABERTO", "")
                               
                Label3(2).Caption = Format((CCur(Label3(2).Caption) + BREC!SGI_VLDOC), "#,##0.00")
                          
             ElseIf StTitulos.Tab = 3 Then
             ElseIf StTitulos.Tab = 4 Then
             End If
             BREC.MoveNext
          Loop
              
          BREC.Close
          
          If StTitulos.Tab = 0 Then flxCADCONTASAPG.SetFocus
          If StTitulos.Tab = 1 Then flxTitBaixados.SetFocus
          If StTitulos.Tab = 2 Then flxLote.SetFocus
          If StTitulos.Tab = 3 Then flxLoteLiberado.SetFocus
          If StTitulos.Tab = 4 Then flxLoteBaixados.SetFocus
          
          Exit Sub
          
       End If
                           
    ElseIf intQtdFiltros = 1 Then
    
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
          sSql = sSql & "     ,CABEC.SGI_CODFOR " & vbCrLf
          sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
          sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_STATUS  " & vbCrLf
          sSql = sSql & "  From" & vbCrLf
          sSql = sSql & "      SGI_CONTASIAPG ITENS" & vbCrLf
          sSql = sSql & "     ,SGI_CONTASHAPG CABEC" & vbCrLf
          sSql = sSql & "     ,SGI_CADFORNEC  FORNE" & vbCrLf
          sSql = sSql & " Where" & vbCrLf
          sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
          sSql = sSql & "  And ITENS.SGI_VLPAGO IS NULL" & vbCrLf
          sSql = sSql & "  And ITENS.SGI_NUMDOC = '" & Trim(txtCampos.Text) & "'" & vbCrLf
          sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
          sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
          sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
          sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODFOR " & vbCrLf
          sSql = sSql & "Order By" & vbCrLf
          If intQtdFiltros = 0 Then
             sSql = sSql & "        ITENS.SGI_CODIGO" & vbCrLf
             sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
             sSql = sSql & "       ,ITENS.SGI_PARCELA"
          ElseIf intQtdFiltros = 1 Then
             sSql = sSql & "        ITENS.SGI_NUMDOC" & vbCrLf
             sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
             sSql = sSql & "       ,ITENS.SGI_PARCELA"
          ElseIf intQtdFiltros = 2 Then
             sSql = sSql & "        FORNE.SGI_RAZAOSOC" & vbCrLf
             sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
             sSql = sSql & "       ,ITENS.SGI_PARCELA"
          ElseIf intQtdFiltros = 3 Then
             sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
             sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
             sSql = sSql & "       ,ITENS.SGI_PARCELA"
          End If
       ElseIf StTitulos.Tab = 1 Then
           sSql = "Select" & vbCrLf
           sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
           sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
           sSql = sSql & "     ,CABEC.SGI_CODFOR " & vbCrLf
           sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
           sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
           sSql = sSql & "     ,ITENS.SGI_DTPGTO" & vbCrLf
           sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
           sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
           sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
           sSql = sSql & "     ,ITENS.SGI_STATUS  " & vbCrLf
           sSql = sSql & "  From" & vbCrLf
           sSql = sSql & "      SGI_CONTASIAPG ITENS" & vbCrLf
           sSql = sSql & "     ,SGI_CONTASHAPG CABEC" & vbCrLf
           sSql = sSql & "     ,SGI_CADFORNEC  FORNE" & vbCrLf
           sSql = sSql & " Where" & vbCrLf
           sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
           sSql = sSql & "  And ITENS.SGI_VLPAGO IS NOT NULL" & vbCrLf
           sSql = sSql & "  And ITENS.SGI_NUMDOC = '" & Trim(txtCampos.Text) & "'" & vbCrLf
           sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
           sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
           sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
           sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODFOR " & vbCrLf
           sSql = sSql & "Order By" & vbCrLf
           If intQtdFiltros = 0 Then
              sSql = sSql & "        ITENS.SGI_CODIGO" & vbCrLf
              sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_PARCELA"
              sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
           ElseIf intQtdFiltros = 1 Then
              sSql = sSql & "        ITENS.SGI_NUMDOC" & vbCrLf
              sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_PARCELA"
              sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
           ElseIf intQtdFiltros = 2 Then
              sSql = sSql & "        FORNE.SGI_RAZAOSOC" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_PARCELA"
              sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
           ElseIf intQtdFiltros = 3 Then
              sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
              sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_PARCELA"
           ElseIf intQtdFiltros = 4 Then
              sSql = sSql & "        ITENS.SGI_DTPGTO" & vbCrLf
              sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_PARCELA"
           End If
       ElseIf StTitulos.Tab = 2 Then
       
           sSql = "Select " & vbCrLf
           sSql = sSql & "       LOTE.SGI_CODIGO    " & vbCrLf
           sSql = sSql & "      ,LOTE.SGI_NUMDOC    " & vbCrLf
           sSql = sSql & "      ,LOTE.SGI_DATALOTE  " & vbCrLf
           sSql = sSql & "      ,LOTE.SGI_VLDOC     " & vbCrLf
           sSql = sSql & "      ,TPGT.SGI_DESCRICAO " & vbCrLf
           sSql = sSql & "      ,LOTE.SGI_CODBCO    " & vbCrLf
           sSql = sSql & "      ,LOTE.SGI_STATUS    " & vbCrLf
           sSql = sSql & "  From " & vbCrLf
           sSql = sSql & "       SGI_CADLOTEHEADER LOTE " & vbCrLf
           sSql = sSql & "      ,SGI_CADTIPOPGTO   TPGT " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       LOTE.SGI_FILIAL = " & FILIAL & vbCrLf
           sSql = sSql & "   And LOTE.SGI_STATUS = 'A'" & vbCrLf
           sSql = sSql & "   And LOTE.SGI_NUMDOC = " & Trim(txtCampos.Text) & vbCrLf
           sSql = sSql & "   And TPGT.SGI_FILIAL = LOTE.SGI_FILIAL  " & vbCrLf
           sSql = sSql & "   And TPGT.SGI_CODIGO = LOTE.SGI_TIPPGTO " & vbCrLf
           sSql = sSql & "Order by " & vbCrLf
           If cboFiltro.ListIndex = 0 Then
              sSql = sSql & "         LOTE.SGI_CODIGO "
           ElseIf cboFiltro.ListIndex = 1 Then
              sSql = sSql & "         LOTE.SGI_NUMDOC "
           ElseIf cboFiltro.ListIndex = 2 Then
              sSql = sSql & "         TPGT.SGI_DESCRICAO "
           ElseIf cboFiltro.ListIndex = 3 Then
              sSql = sSql & "         LOTE.SGI_DATALOTE "
           End If
     
           Label3(2).Caption = Format(0, "#,##0.00")
       
       
       ElseIf StTitulos.Tab = 3 Then
       ElseIf StTitulos.Tab = 4 Then
       End If
       
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
          
          If StTitulos.Tab = 0 Then ConfGrid
          If StTitulos.Tab = 1 Then ConfGridBaixados
          If StTitulos.Tab = 2 Then ConfGridLote
          
          Do While Not BREC.EOF
          
             If StTitulos.Tab = 0 Then
                flxCADCONTASAPG.AddItem "" & vbTab & _
                                        BREC!SGI_CODIGO & vbTab & _
                                        BREC!SGI_NUMDOC & vbTab & _
                                        Format(BREC!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                                        Format(BREC!SGI_PARCELA, "##00") & "/" & Format(BREC!SGI_QTDPARC, "##00") & vbTab & _
                                        Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                                        BREC!SGI_CODFOR & vbTab & _
                                        BREC!SGI_RAZAOSOC & vbTab & _
                                        BREC!SGI_PARCELA & vbTab & _
                                        BREC!SGI_STATUS
             
             ElseIf StTitulos.Tab = 1 Then
                flxTitBaixados.AddItem "" & vbTab & _
                                       BREC!SGI_CODIGO & vbTab & _
                                       BREC!SGI_NUMDOC & vbTab & _
                                       Format(BREC!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                                       Format(BREC!SGI_DTPGTO, "DD/MM/YYYY") & vbTab & _
                                       Format(BREC!SGI_PARCELA, "##00") & "/" & Format(BREC!SGI_QTDPARC, "##00") & vbTab & _
                                       Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                                       BREC!SGI_CODFOR & vbTab & _
                                       BREC!SGI_RAZAOSOC & vbTab & _
                                       BREC!SGI_PARCELA & vbTab & _
                                       BREC!SGI_STATUS
             
             ElseIf StTitulos.Tab = 2 Then
             
                strLoteDescOrig = ""
                
                '' ------------------------
                sSql = "Select " & vbCrLf
                sSql = sSql & "       * " & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_CADBANCOS " & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
                sSql = sSql & "   And SGI_CODIGO = " & BREC!SGI_CODBCO
                 
                BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
                If Not BREC2.EOF Then strLoteDescOrig = BREC2!SGI_DESCRICAO
                BREC2.Close
                '' ------------------------
                 
                flxLote.AddItem "" & vbTab & _
                                BREC!SGI_CODIGO & vbTab & _
                                BREC!SGI_NUMDOC & vbTab & _
                                Format(BREC!SGI_DATALOTE, "DD/MM/YYYY") & vbTab & _
                                Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                                BREC!SGI_DESCRICAO & vbTab & _
                                strLoteDescOrig & vbTab & _
                                BREC!SGI_STATUS & vbTab & _
                                IIf(BREC!SGI_STATUS = "A", "ABERTO", "")
                               
                Label3(2).Caption = Format((CCur(Label3(2).Caption) + BREC!SGI_VLDOC), "#,##0.00")
             
             ElseIf StTitulos.Tab = 3 Then
             ElseIf StTitulos.Tab = 4 Then
             End If
             
             BREC.MoveNext
          Loop
              
          BREC.Close
          
          If StTitulos.Tab = 0 Then flxCADCONTASAPG.SetFocus
          If StTitulos.Tab = 1 Then flxTitBaixados.SetFocus
          If StTitulos.Tab = 2 Then flxLote.SetFocus
          If StTitulos.Tab = 3 Then flxLoteLiberado.SetFocus
          If StTitulos.Tab = 4 Then flxLoteBaixados.SetFocus
          
          Exit Sub
          
       End If
    
    ElseIf intQtdFiltros = 2 Then
    
       If StTitulos.Tab = 0 Then
          sSql = "Select" & vbCrLf
          sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
          sSql = sSql & "     ,CABEC.SGI_CODFOR " & vbCrLf
          sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
          sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_STATUS  " & vbCrLf
          sSql = sSql & "  From" & vbCrLf
          sSql = sSql & "      SGI_CONTASIAPG ITENS" & vbCrLf
          sSql = sSql & "     ,SGI_CONTASHAPG CABEC" & vbCrLf
          sSql = sSql & "     ,SGI_CADFORNEC  FORNE" & vbCrLf
          sSql = sSql & " Where" & vbCrLf
          sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
          sSql = sSql & "  And ITENS.SGI_VLPAGO IS NULL" & vbCrLf
          sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
          sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
          sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
          sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODFOR " & vbCrLf
          sSql = sSql & "  And FORNE.SGI_RAZAOSOC like '" & Trim(txtCampos.Text) & "%'"
          sSql = sSql & "Order By" & vbCrLf
          If intQtdFiltros = 0 Then
             sSql = sSql & "        ITENS.SGI_CODIGO" & vbCrLf
             sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
             sSql = sSql & "       ,ITENS.SGI_PARCELA"
          ElseIf intQtdFiltros = 1 Then
             sSql = sSql & "        ITENS.SGI_NUMDOC" & vbCrLf
             sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
             sSql = sSql & "       ,ITENS.SGI_PARCELA"
          ElseIf intQtdFiltros = 2 Then
             sSql = sSql & "        FORNE.SGI_RAZAOSOC" & vbCrLf
             sSql = sSql & "       ,ITENS.SGI_NUMDOC" & vbCrLf
             sSql = sSql & "       ,ITENS.SGI_PARCELA" & vbCrLf
             sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
          ElseIf intQtdFiltros = 3 Then
             sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
             sSql = sSql & "       ,ITENS.SGI_NUMDOC" & vbCrLf
             sSql = sSql & "       ,ITENS.SGI_PARCELA"
             sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
          End If
       ElseIf StTitulos.Tab = 1 Then
           sSql = "Select" & vbCrLf
           sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
           sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
           sSql = sSql & "     ,CABEC.SGI_CODFOR " & vbCrLf
           sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
           sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
           sSql = sSql & "     ,ITENS.SGI_DTPGTO" & vbCrLf
           sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
           sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
           sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
           sSql = sSql & "     ,ITENS.SGI_STATUS  " & vbCrLf
           sSql = sSql & "  From" & vbCrLf
           sSql = sSql & "      SGI_CONTASIAPG ITENS" & vbCrLf
           sSql = sSql & "     ,SGI_CONTASHAPG CABEC" & vbCrLf
           sSql = sSql & "     ,SGI_CADFORNEC  FORNE" & vbCrLf
           sSql = sSql & " Where" & vbCrLf
           sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
           sSql = sSql & "  And ITENS.SGI_VLPAGO IS NOT NULL" & vbCrLf
           sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
           sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
           sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
           sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODFOR " & vbCrLf
           sSql = sSql & "  And FORNE.SGI_RAZAOSOC Like '" & Trim(txtCampos.Text) & "%'" & vbCrLf
           sSql = sSql & "Order By" & vbCrLf
           If intQtdFiltros = 0 Then
              sSql = sSql & "        ITENS.SGI_CODIGO" & vbCrLf
              sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_PARCELA"
              sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
           ElseIf intQtdFiltros = 1 Then
              sSql = sSql & "        ITENS.SGI_NUMDOC" & vbCrLf
              sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_PARCELA"
              sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
           ElseIf intQtdFiltros = 2 Then
              sSql = sSql & "        FORNE.SGI_RAZAOSOC" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_NUMDOC" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_PARCELA"
              sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
           ElseIf intQtdFiltros = 3 Then
              sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
              sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_PARCELA"
           ElseIf intQtdFiltros = 4 Then
              sSql = sSql & "        ITENS.SGI_DTPGTO" & vbCrLf
              sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_PARCELA"
           End If
       ElseIf StTitulos.Tab = 2 Or StTitulos.Tab = 3 Or StTitulos.Tab = 4 Then
       
           If Not IsDate(txtCampos.Text) Then
              MsgBox "Somente é permitido datas !!!", vbOKOnly + vbCritical, "Aviso"
              txtCampos.Text = ""
              txtCampos.SetFocus
              Exit Sub
           End If
           
           sSql = "Select " & vbCrLf
           sSql = sSql & "       LOTE.SGI_CODIGO    " & vbCrLf
           sSql = sSql & "      ,LOTE.SGI_NUMDOC    " & vbCrLf
           sSql = sSql & "      ,LOTE.SGI_DATALOTE  " & vbCrLf
           sSql = sSql & "      ,LOTE.SGI_VLDOC     " & vbCrLf
           sSql = sSql & "      ,TPGT.SGI_DESCRICAO " & vbCrLf
           sSql = sSql & "      ,LOTE.SGI_CODBCO    " & vbCrLf
           sSql = sSql & "      ,LOTE.SGI_STATUS    " & vbCrLf
           sSql = sSql & "  From " & vbCrLf
           sSql = sSql & "       SGI_CADLOTEHEADER LOTE " & vbCrLf
           sSql = sSql & "      ,SGI_CADTIPOPGTO   TPGT " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       LOTE.SGI_FILIAL = " & FILIAL & vbCrLf
           If StTitulos.Tab = 2 Then sSql = sSql & "   And LOTE.SGI_STATUS = 'A'" & vbCrLf
           If StTitulos.Tab = 3 Then sSql = sSql & "   And LOTE.SGI_STATUS = 'L'" & vbCrLf
           If StTitulos.Tab = 4 Then sSql = sSql & "   And LOTE.SGI_STATUS = 'B'" & vbCrLf
           sSql = sSql & "   And LOTE.SGI_DATALOTE = '" & Format(CDate(txtCampos.Text), "MM/DD/YYYY") & "'" & vbCrLf
           sSql = sSql & "   And TPGT.SGI_FILIAL   = LOTE.SGI_FILIAL  " & vbCrLf
           sSql = sSql & "   And TPGT.SGI_CODIGO   = LOTE.SGI_TIPPGTO " & vbCrLf
           sSql = sSql & "Order by " & vbCrLf
           If cboFiltro.ListIndex = 0 Then
              sSql = sSql & "         LOTE.SGI_CODIGO "
           ElseIf cboFiltro.ListIndex = 1 Then
              sSql = sSql & "         LOTE.SGI_NUMDOC "
           ElseIf cboFiltro.ListIndex = 2 Then
              sSql = sSql & "         LOTE.SGI_DATALOTE "
           End If
        
           Label3(2).Caption = Format(0, "#,##0.00")
       
       End If
    
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
          
          If StTitulos.Tab = 0 Then ConfGrid
          If StTitulos.Tab = 1 Then ConfGridBaixados
          If StTitulos.Tab = 2 Then ConfGridLote
          If StTitulos.Tab = 3 Then ConfGridLoteLiberdo
          If StTitulos.Tab = 4 Then ConfGridLoteBaixado
    
          Do While Not BREC.EOF
          
             If StTitulos.Tab = 0 Then
                flxCADCONTASAPG.AddItem "" & vbTab & _
                                        BREC!SGI_CODIGO & vbTab & _
                                        BREC!SGI_NUMDOC & vbTab & _
                                        Format(BREC!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                                        Format(BREC!SGI_PARCELA, "##00") & "/" & Format(BREC!SGI_QTDPARC, "##00") & vbTab & _
                                        Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                                        BREC!SGI_CODFOR & vbTab & _
                                        BREC!SGI_RAZAOSOC & vbTab & _
                                        BREC!SGI_PARCELA & vbTab & _
                                        BREC!SGI_STATUS
             
             ElseIf StTitulos.Tab = 1 Then
                flxTitBaixados.AddItem "" & vbTab & _
                                       BREC!SGI_CODIGO & vbTab & _
                                       BREC!SGI_NUMDOC & vbTab & _
                                       Format(BREC!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                                       Format(BREC!SGI_DTPGTO, "DD/MM/YYYY") & vbTab & _
                                       Format(BREC!SGI_PARCELA, "##00") & "/" & Format(BREC!SGI_QTDPARC, "##00") & vbTab & _
                                       Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                                       BREC!SGI_CODFOR & vbTab & _
                                       BREC!SGI_RAZAOSOC & vbTab & _
                                       BREC!SGI_PARCELA & vbTab & _
                                       BREC!SGI_STATUS
             
             ElseIf StTitulos.Tab = 2 Then
                strLoteDescOrig = ""
       
                '' ------------------------
                sSql = "Select " & vbCrLf
                sSql = sSql & "       * " & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_CADBANCOS " & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
                sSql = sSql & "   And SGI_CODIGO = " & BREC!SGI_CODBCO
                
                BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
                If Not BREC2.EOF Then strLoteDescOrig = BREC2!SGI_DESCRICAO
                BREC2.Close
                '' ------------------------
                
                flxLote.AddItem "" & vbTab & _
                                BREC!SGI_CODIGO & vbTab & _
                                BREC!SGI_NUMDOC & vbTab & _
                                Format(BREC!SGI_DATALOTE, "DD/MM/YYYY") & vbTab & _
                                Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                                BREC!SGI_DESCRICAO & vbTab & _
                                strLoteDescOrig & vbTab & _
                                BREC!SGI_STATUS & vbTab & _
                                IIf(BREC!SGI_STATUS = "A", "ABERTO", "")
                               
                Label3(2).Caption = Format((CCur(Label3(2).Caption) + BREC!SGI_VLDOC), "#,##0.00")
       
             ElseIf StTitulos.Tab = 3 Then
     
                strLoteDescOrig = ""
              
                '' ------------------------
                sSql = "Select " & vbCrLf
                sSql = sSql & "       * " & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_CADBANCOS " & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
                sSql = sSql & "   And SGI_CODIGO = " & BREC!SGI_CODBCO
              
                BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
                If Not BREC2.EOF Then strLoteDescOrig = BREC2!SGI_DESCRICAO
                BREC2.Close
                '' ------------------------
            
                flxLoteLiberado.AddItem "" & vbTab & _
                                        BREC!SGI_CODIGO & vbTab & _
                                        BREC!SGI_NUMDOC & vbTab & _
                                        Format(BREC!SGI_DATALOTE, "DD/MM/YYYY") & vbTab & _
                                        Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                                        BREC!SGI_DESCRICAO & vbTab & _
                                        strLoteDescOrig & vbTab & _
                                        BREC!SGI_STATUS & vbTab & _
                                        IIf(BREC!SGI_STATUS = "L", "LIBERADO", "")
                                
             ElseIf StTitulos.Tab = 4 Then
        
                 strLoteDescOrig = ""
       
                 '' ------------------------
                 sSql = "Select " & vbCrLf
                 sSql = sSql & "       * " & vbCrLf
                 sSql = sSql & "  From " & vbCrLf
                 sSql = sSql & "       SGI_CADBANCOS " & vbCrLf
                 sSql = sSql & " Where " & vbCrLf
                 sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
                 sSql = sSql & "   And SGI_CODIGO = " & BREC!SGI_CODBCO
       
                 BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
                 If Not BREC2.EOF Then strLoteDescOrig = BREC2!SGI_DESCRICAO
                 BREC2.Close
                 '' ------------------------
       
                 flxLoteBaixados.AddItem "" & vbTab & _
                                         BREC!SGI_CODIGO & vbTab & _
                                         BREC!SGI_NUMDOC & vbTab & _
                                         Format(BREC!SGI_DATALOTE, "DD/MM/YYYY") & vbTab & _
                                         Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                                         BREC!SGI_DESCRICAO & vbTab & _
                                         strLoteDescOrig & vbTab & _
                                         BREC!SGI_STATUS & vbTab & _
                                         IIf(BREC!SGI_STATUS = "B", "BAIXADO", "")
             
             End If
             
             BREC.MoveNext
          Loop
              
          BREC.Close
          
          If StTitulos.Tab = 0 Then flxCADCONTASAPG.SetFocus
          If StTitulos.Tab = 1 Then flxTitBaixados.SetFocus
          If StTitulos.Tab = 2 Then flxLote.SetFocus
          If StTitulos.Tab = 3 Then flxLoteLiberado.SetFocus
          If StTitulos.Tab = 4 Then flxLoteBaixados.SetFocus
          
          Exit Sub
       
       End If
    
    ElseIf intQtdFiltros = 3 Then
    
       If Not IsDate(txtCampos.Text) Then
          MsgBox "Somente é permitido datas !!!", vbOKOnly + vbCritical, "Aviso"
          txtCampos.Text = ""
          txtCampos.SetFocus
          Exit Sub
       End If
    
       If StTitulos.Tab = 0 Then
          sSql = "Select" & vbCrLf
          sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
          sSql = sSql & "     ,CABEC.SGI_CODFOR " & vbCrLf
          sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
          sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
          sSql = sSql & "     ,ITENS.SGI_STATUS  " & vbCrLf
          sSql = sSql & "  From" & vbCrLf
          sSql = sSql & "      SGI_CONTASIAPG ITENS" & vbCrLf
          sSql = sSql & "     ,SGI_CONTASHAPG CABEC" & vbCrLf
          sSql = sSql & "     ,SGI_CADFORNEC  FORNE" & vbCrLf
          sSql = sSql & " Where" & vbCrLf
          sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
          sSql = sSql & "  And ITENS.SGI_VLPAGO IS NULL" & vbCrLf
          sSql = sSql & "  and ITENS.SGI_DATAVENC ='" & Format(CDate(txtCampos.Text), "MM/DD/YYYY") & "'" & vbCrLf
          sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
          sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
          sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
          sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODFOR " & vbCrLf
          sSql = sSql & "Order By" & vbCrLf
          If intQtdFiltros = 0 Then
             sSql = sSql & "        ITENS.SGI_CODIGO" & vbCrLf
             sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
             sSql = sSql & "       ,ITENS.SGI_PARCELA"
          ElseIf intQtdFiltros = 1 Then
             sSql = sSql & "        ITENS.SGI_NUMDOC" & vbCrLf
             sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
             sSql = sSql & "       ,ITENS.SGI_PARCELA"
          ElseIf intQtdFiltros = 2 Then
             sSql = sSql & "        FORNE.SGI_RAZAOSOC" & vbCrLf
             sSql = sSql & "       ,ITENS.SGI_NUMDOC" & vbCrLf
             sSql = sSql & "       ,ITENS.SGI_PARCELA" & vbCrLf
             sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
          ElseIf intQtdFiltros = 3 Then
             sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
             sSql = sSql & "       ,ITENS.SGI_NUMDOC" & vbCrLf
             sSql = sSql & "       ,ITENS.SGI_PARCELA"
             sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
          End If
          
       ElseIf StTitulos.Tab = 1 Then
       
           sSql = "Select" & vbCrLf
           sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
           sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
           sSql = sSql & "     ,CABEC.SGI_CODFOR " & vbCrLf
           sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
           sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
           sSql = sSql & "     ,ITENS.SGI_DTPGTO" & vbCrLf
           sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
           sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
           sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
           sSql = sSql & "     ,ITENS.SGI_STATUS  " & vbCrLf
           sSql = sSql & "  From" & vbCrLf
           sSql = sSql & "      SGI_CONTASIAPG ITENS" & vbCrLf
           sSql = sSql & "     ,SGI_CONTASHAPG CABEC" & vbCrLf
           sSql = sSql & "     ,SGI_CADFORNEC  FORNE" & vbCrLf
           sSql = sSql & " Where" & vbCrLf
           sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
           sSql = sSql & "  And ITENS.SGI_VLPAGO IS NOT NULL" & vbCrLf
           sSql = sSql & "  And ITENS.SGI_DATAVENC = '" & Format(CDate(txtCampos.Text), "MM/DD/YYYY") & "'" & vbCrLf
           sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
           sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
           sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
           sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODFOR " & vbCrLf
           sSql = sSql & "Order By" & vbCrLf
           If intQtdFiltros = 0 Then
              sSql = sSql & "        ITENS.SGI_CODIGO" & vbCrLf
              sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_PARCELA"
              sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
           ElseIf intQtdFiltros = 1 Then
              sSql = sSql & "        ITENS.SGI_NUMDOC" & vbCrLf
              sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_PARCELA"
              sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
           ElseIf intQtdFiltros = 2 Then
              sSql = sSql & "        FORNE.SGI_RAZAOSOC" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_NUMDOC" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_PARCELA"
              sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
           ElseIf intQtdFiltros = 3 Then
              sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_NUMDOC" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_PARCELA"
              sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
           ElseIf intQtdFiltros = 4 Then
              sSql = sSql & "        ITENS.SGI_DTPGTO" & vbCrLf
              sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_PARCELA"
           End If
       ElseIf StTitulos.Tab = 2 Then
       ElseIf StTitulos.Tab = 3 Then
       ElseIf StTitulos.Tab = 4 Then
       End If
    
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
          
          If StTitulos.Tab = 0 Then ConfGrid
          If StTitulos.Tab = 1 Then ConfGridBaixados
    
          Do While Not BREC.EOF
          
             If StTitulos.Tab = 0 Then
                flxCADCONTASAPG.AddItem "" & vbTab & _
                                        BREC!SGI_CODIGO & vbTab & _
                                        BREC!SGI_NUMDOC & vbTab & _
                                        Format(BREC!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                                        Format(BREC!SGI_PARCELA, "##00") & "/" & Format(BREC!SGI_QTDPARC, "##00") & vbTab & _
                                        Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                                        BREC!SGI_CODFOR & vbTab & _
                                        BREC!SGI_RAZAOSOC & vbTab & _
                                        BREC!SGI_PARCELA & vbTab & _
                                        BREC!SGI_STATUS
             
             ElseIf StTitulos.Tab = 1 Then
                flxTitBaixados.AddItem "" & vbTab & _
                                       BREC!SGI_CODIGO & vbTab & _
                                       BREC!SGI_NUMDOC & vbTab & _
                                       Format(BREC!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                                       Format(BREC!SGI_DTPGTO, "DD/MM/YYYY") & vbTab & _
                                       Format(BREC!SGI_PARCELA, "##00") & "/" & Format(BREC!SGI_QTDPARC, "##00") & vbTab & _
                                       Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                                       BREC!SGI_CODFOR & vbTab & _
                                       BREC!SGI_RAZAOSOC & vbTab & _
                                       BREC!SGI_PARCELA & vbTab & _
                                       BREC!SGI_STATUS
             
             ElseIf StTitulos.Tab = 2 Then
             ElseIf StTitulos.Tab = 3 Then
             ElseIf StTitulos.Tab = 4 Then
             End If
             
             BREC.MoveNext
          Loop
              
          BREC.Close
          
          If StTitulos.Tab = 0 Then flxCADCONTASAPG.SetFocus
          If StTitulos.Tab = 1 Then flxTitBaixados.SetFocus
          If StTitulos.Tab = 2 Then flxLote.SetFocus
          If StTitulos.Tab = 3 Then flxLoteLiberado.SetFocus
          If StTitulos.Tab = 4 Then flxLoteBaixados.SetFocus
          
          Exit Sub
       End If
    
    ElseIf intQtdFiltros = 4 Then
    
       If Not IsDate(txtCampos.Text) Then
          MsgBox "Somente é permitido datas !!!", vbOKOnly + vbCritical, "Aviso"
          txtCampos.Text = ""
          txtCampos.SetFocus
          Exit Sub
       End If
       
          
       If StTitulos.Tab = 1 Then
       
           sSql = "Select" & vbCrLf
           sSql = sSql & "      ITENS.SGI_CODIGO " & vbCrLf
           sSql = sSql & "     ,ITENS.SGI_NUMDOC " & vbCrLf
           sSql = sSql & "     ,CABEC.SGI_CODFOR " & vbCrLf
           sSql = sSql & "     ,FORNE.SGI_RAZAOSOC" & vbCrLf
           sSql = sSql & "     ,ITENS.SGI_DATAVENC" & vbCrLf
           sSql = sSql & "     ,ITENS.SGI_DTPGTO" & vbCrLf
           sSql = sSql & "     ,ITENS.SGI_PARCELA " & vbCrLf
           sSql = sSql & "     ,CABEC.SGI_QTDPARC " & vbCrLf
           sSql = sSql & "     ,ITENS.SGI_VLDOC   " & vbCrLf
           sSql = sSql & "     ,ITENS.SGI_STATUS  " & vbCrLf
           sSql = sSql & "  From" & vbCrLf
           sSql = sSql & "      SGI_CONTASIAPG ITENS" & vbCrLf
           sSql = sSql & "     ,SGI_CONTASHAPG CABEC" & vbCrLf
           sSql = sSql & "     ,SGI_CADFORNEC  FORNE" & vbCrLf
           sSql = sSql & " Where" & vbCrLf
           sSql = sSql & "      ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
           sSql = sSql & "  And ITENS.SGI_VLPAGO IS NOT NULL" & vbCrLf
           sSql = sSql & "  And ITENS.SGI_DTPGTO = '" & Format(CDate(txtCampos.Text), "MM/DD/YYYY") & "'" & vbCrLf
           sSql = sSql & "  And CABEC.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
           sSql = sSql & "  And CABEC.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
           sSql = sSql & "  And FORNE.SGI_FILIAL = CABEC.SGI_FILIAL " & vbCrLf
           sSql = sSql & "  And FORNE.SGI_CODIGO = CABEC.SGI_CODFOR " & vbCrLf
           sSql = sSql & "Order By" & vbCrLf
           If intQtdFiltros = 0 Then
              sSql = sSql & "        ITENS.SGI_CODIGO" & vbCrLf
              sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_PARCELA"
              sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
           ElseIf intQtdFiltros = 1 Then
              sSql = sSql & "        ITENS.SGI_NUMDOC" & vbCrLf
              sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_PARCELA"
              sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
           ElseIf intQtdFiltros = 2 Then
              sSql = sSql & "        FORNE.SGI_RAZAOSOC" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_NUMDOC" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_PARCELA"
              sSql = sSql & "       ,ITENS.SGI_DATAVENC" & vbCrLf
           ElseIf intQtdFiltros = 3 Then
              sSql = sSql & "        ITENS.SGI_DATAVENC" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_NUMDOC" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_PARCELA" & vbCrLf
              sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
           ElseIf intQtdFiltros = 4 Then
              sSql = sSql & "        ITENS.SGI_DTPGTO" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_NUMDOC" & vbCrLf
              sSql = sSql & "       ,ITENS.SGI_PARCELA" & vbCrLf
              sSql = sSql & "       ,FORNE.SGI_RAZAOSOC" & vbCrLf
              
           End If
       ElseIf StTitulos.Tab = 2 Then
       ElseIf StTitulos.Tab = 3 Then
       ElseIf StTitulos.Tab = 4 Then
       End If
       
       
       BREC.Open sSql, adoBanco_Dados, adOpenDynamic
         
       If Not BREC.EOF Then
          
          If StTitulos.Tab = 0 Then ConfGrid
          If StTitulos.Tab = 1 Then ConfGridBaixados
    
          Do While Not BREC.EOF
          
             If StTitulos.Tab = 1 Then
                flxTitBaixados.AddItem "" & vbTab & _
                                       BREC!SGI_CODIGO & vbTab & _
                                       BREC!SGI_NUMDOC & vbTab & _
                                       Format(BREC!SGI_DATAVENC, "DD/MM/YYYY") & vbTab & _
                                       Format(BREC!SGI_DTPGTO, "DD/MM/YYYY") & vbTab & _
                                       Format(BREC!SGI_PARCELA, "##00") & "/" & Format(BREC!SGI_QTDPARC, "##00") & vbTab & _
                                       Format(BREC!SGI_VLDOC, "#,##0.00") & vbTab & _
                                       BREC!SGI_CODFOR & vbTab & _
                                       BREC!SGI_RAZAOSOC & vbTab & _
                                       BREC!SGI_PARCELA & vbTab & _
                                       BREC!SGI_STATUS
             
             ElseIf StTitulos.Tab = 2 Then
             ElseIf StTitulos.Tab = 3 Then
             ElseIf StTitulos.Tab = 4 Then
             End If
             
             BREC.MoveNext
          Loop
              
          BREC.Close
          
          If StTitulos.Tab = 1 Then flxTitBaixados.SetFocus
          If StTitulos.Tab = 2 Then flxLote.SetFocus
          If StTitulos.Tab = 3 Then flxLoteLiberado.SetFocus
          If StTitulos.Tab = 4 Then flxLoteBaixados.SetFocus
          
          Exit Sub
       End If
    End If

    BREC.Close
    

End Sub
