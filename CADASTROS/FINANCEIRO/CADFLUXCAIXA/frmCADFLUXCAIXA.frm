VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCADFLUXCAIXA 
   Caption         =   "Fluxo de Caixa"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   11280
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab STFluxo 
      Height          =   4575
      Left            =   0
      TabIndex        =   19
      Top             =   1560
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   8070
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
      TabCaption(0)   =   "Lançamentos"
      TabPicture(0)   =   "frmCADFLUXCAIXA.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Fechamento"
      TabPicture(1)   =   "frmCADFLUXCAIXA.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   1095
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   11055
         Begin VB.TextBox txtNUMDOC 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7680
            MaxLength       =   10
            TabIndex        =   4
            Text            =   "txtNUMDOC"
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmbGravPGT 
            Height          =   315
            Left            =   6960
            Picture         =   "frmCADFLUXCAIXA.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtValor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   9720
            TabIndex        =   5
            Text            =   "txtValor"
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdPesq 
            Height          =   315
            Left            =   2520
            Picture         =   "frmCADFLUXCAIXA.frx":013A
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   600
            Width           =   375
         End
         Begin VB.ComboBox cboBanco 
            Height          =   315
            Left            =   2925
            TabIndex        =   7
            Text            =   "cboBanco"
            Top             =   600
            Width           =   3975
         End
         Begin VB.TextBox txtBanco 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   6
            Text            =   "txtBanco"
            Top             =   600
            Width           =   735
         End
         Begin VB.CommandButton Command1 
            Height          =   315
            Left            =   2535
            Picture         =   "frmCADFLUXCAIXA.frx":023C
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   210
            Width           =   375
         End
         Begin VB.ComboBox cboTIPOPGTO 
            Height          =   315
            Left            =   2925
            TabIndex        =   3
            Text            =   "cboTIPOPGTO"
            Top             =   210
            Width           =   3975
         End
         Begin VB.TextBox txtCODTIPPGTO 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   2
            Text            =   "txtCODFORNEC"
            Top             =   210
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nº Doc:"
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
            Index           =   7
            Left            =   6960
            TabIndex        =   30
            Top             =   240
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor:"
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
            Index           =   6
            Left            =   9120
            TabIndex        =   29
            Top             =   240
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
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
            Index           =   4
            Left            =   1080
            TabIndex        =   27
            Top             =   630
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Lançamento:"
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
            Left            =   155
            TabIndex        =   25
            Top             =   240
            Width           =   1545
         End
      End
      Begin VB.Frame Frame5 
         Height          =   3015
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   11055
         Begin MSFlexGridLib.MSFlexGrid flxCadLctos 
            Height          =   2655
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   4683
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame4 
         Height          =   4095
         Left            =   -74880
         TabIndex        =   20
         Top             =   360
         Width           =   11055
         Begin MSFlexGridLib.MSFlexGrid flxFluxoCaixa 
            Height          =   3735
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   6588
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            SelectionMode   =   1
            Appearance      =   0
         End
      End
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   0
      TabIndex        =   15
      Top             =   960
      Width           =   11295
      Begin VB.CommandButton cmdCarrega 
         Height          =   315
         Left            =   4320
         Picture         =   "frmCADFLUXCAIXA.frx":033E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskDTFLUXCAIXA 
         Height          =   285
         Left            =   3120
         TabIndex        =   0
         Top             =   150
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   9
         Top             =   170
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código.:"
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
         TabIndex        =   18
         Top             =   195
         Width           =   720
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   6000
         TabIndex        =   10
         Top             =   165
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Status.:"
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
         Left            =   5160
         TabIndex        =   17
         Top             =   195
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data.:"
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
         Left            =   2400
         TabIndex        =   16
         Top             =   195
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11295
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
         Picture         =   "frmCADFLUXCAIXA.frx":076E
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   975
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
         Left            =   1920
         Picture         =   "frmCADFLUXCAIXA.frx":0870
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Left            =   1080
         Picture         =   "frmCADFLUXCAIXA.frx":0972
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADFLUXCAIXA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho      As String
Public Linha         As Variant
Public cTipOper      As String
Public iCodigo       As Integer
Public iParcela      As Integer
Public FILIAL        As Integer
Public strAcesso     As String
Public strMODPAI     As String
Public strUsuario    As String
Dim ccurSALDOANT     As Currency
Dim dtDTSALDOANT     As Date
Dim objBLBFunc       As Object
Dim objPESQPADRAO    As Object
Dim objCADFLXCAIXA   As Object
Dim arrLANCAMENTOS   As Variant
Dim arrFLUXOCAIXA    As Variant

Private Sub cboBanco_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboBanco, KeyAscii
End Sub

Private Sub cboBanco_Validate(Cancel As Boolean)
    If cboBanco.ListIndex > -1 Then txtBanco.Text = cboBanco.ItemData(cboBanco.ListIndex)
End Sub

Private Sub cboTIPOPGTO_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboTIPOPGTO, KeyAscii
End Sub

Private Sub cboTIPOPGTO_Validate(Cancel As Boolean)
    If cboTIPOPGTO.ListIndex > -1 Then txtCODTIPPGTO.Text = cboTIPOPGTO.ItemData(cboTIPOPGTO.ListIndex)
End Sub


Private Sub cmbGravPGT_Click()
    InseriGridTipoPgto
    PintaCelulaLcto
    PintaCelula
End Sub

Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strAcesso) = False Then Exit Sub
    
    If objCADFLXCAIXA.STATUS = "FECHADO" Then
       MsgBox "Caixa fechado !!!", vbOKOnly + vbExclamation, "aviso"
       Exit Sub
    End If
    
    cmdVoltar.Enabled = True
    cmdAltera.Enabled = False
    CmdSalva.Enabled = True
    
    Me.Caption = "Fluxo de Caixa - [ ALTERAÇÃO ]"
    
    Frame3.Enabled = True
    STFluxo.Tab = 0
    
    cTipOper = "A"
    
    txtCODTIPPGTO.SetFocus
    
End Sub

Private Sub cmdCarrega_Click()
    
    Dim intResp As Integer
    Dim intDIAS As Integer
    
    If Not IsDate(mskDTFLUXCAIXA.Text) Then
       MsgBox "Data inválida !!!", vbOKOnly + vbExclamation, "Aviso"
       mskDTFLUXCAIXA.Text = Format(Now, "DD/MM/YYYY")
       mskDTFLUXCAIXA.SetFocus
       Exit Sub
    End If
    
    If VerificaCaixa = True Then
       MsgBox "Este caixa já foi criado !!!", vbOKOnly + vbExclamation, "Aviso"
       mskDTFLUXCAIXA.Text = Format(Now, "DD/MM/YYYY")
       mskDTFLUXCAIXA.SetFocus
       Exit Sub
    End If
    
    If CaixaFechado = True Then Exit Sub
    
    ccurSALDOANT = PegaSaldoAnt
    
    If CarregaCaixa = False Then
       
       MsgBox "Não existe movimentos para este dia !!!", vbOKOnly + vbExclamation, "Aviso"
       intResp = MsgBox("Deseja abrir este dia mesmo assim ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
       
       If intResp = 7 Then
          mskDTFLUXCAIXA.SetFocus
          Exit Sub
       End If
       
       STFluxo.Enabled = True
       txtCODTIPPGTO.SetFocus
       
       Exit Sub
    End If

End Sub

Private Sub CmdSalva_Click()
    
    Dim I       As Integer
    Dim intResp As Integer
    
    If ValidaCampos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADFLXCAIXA.CODFLXCAIXA = objCADFLXCAIXA.Gera_Codigo(Me.Name)
    objCADFLXCAIXA.DATAFLXCAIXA = CDate(mskDTFLUXCAIXA.Text)
    objCADFLXCAIXA.STATUS = lblStatus(2).Caption
    objCADFLXCAIXA.SALDOANT = CCur(flxFluxoCaixa.TextMatrix(1, 3))
    objCADFLXCAIXA.SALDOATU = CCur(flxFluxoCaixa.TextMatrix(flxFluxoCaixa.Rows - 1, 3))
    
    '' Lançamentos
    If flxCadLctos.Rows > 1 Then
       ReDim arrLANCAMENTOS(1 To (flxCadLctos.Rows - 1), 1 To 5)
       For I = 1 To (flxCadLctos.Rows - 1)
           
           arrLANCAMENTOS(I, 1) = flxCadLctos.TextMatrix(I, 7)        '' Tipo de Pagamento
           
           arrLANCAMENTOS(I, 2) = 0
           If Len(Trim(flxCadLctos.TextMatrix(I, 8))) > 0 Then arrLANCAMENTOS(I, 2) = flxCadLctos.TextMatrix(I, 8) '' Banco
           
           arrLANCAMENTOS(I, 3) = CDate(flxCadLctos.TextMatrix(I, 3)) '' Data
           arrLANCAMENTOS(I, 4) = flxCadLctos.TextMatrix(I, 4)        '' Nº Documento
           arrLANCAMENTOS(I, 5) = CCur(flxCadLctos.TextMatrix(I, 5))  '' Valor
           
       Next I
       objCADFLXCAIXA.LCTOFLXCAIXA = arrLANCAMENTOS
    End If
    
    '' Saldo de Caixa
    If flxFluxoCaixa.Rows > 1 Then
       ReDim arrFLUXOCAIXA(1 To (flxFluxoCaixa.Rows - 1), 1 To 4)
       For I = 1 To (flxFluxoCaixa.Rows - 1)
           If Len(Trim(flxFluxoCaixa.TextMatrix(I, 2))) > 0 Then
              arrFLUXOCAIXA(I, 1) = flxFluxoCaixa.TextMatrix(I, 1)
              arrFLUXOCAIXA(I, 2) = CDate(flxFluxoCaixa.TextMatrix(I, 2))
              arrFLUXOCAIXA(I, 3) = CCur(flxFluxoCaixa.TextMatrix(I, 3))
              arrFLUXOCAIXA(I, 4) = CLng(flxFluxoCaixa.TextMatrix(I, 5))
           End If
       Next I
       objCADFLXCAIXA.FLUXOCAIXA = arrFLUXOCAIXA
    End If
    
    If objCADFLXCAIXA.GRAVA(cTipOper) = False Then Exit Sub
          
    MsgBox "O caixa foi " & IIf(cTipOper = "I", "incluso", IIf(cTipOper = "A", "alterado", IIf(cTipOper = "B", "baixado", IIf(cTipOper = "AB", "alterado", "")))) & " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
    If cTipOper = "I" Then
       intResp = MsgBox("Deseja incluir novo caixa ?", vbYesNo + vbQuestion + vbDefaultButton1, "Aviso")
       If intResp = 7 Then
          Set objBLBFunc = Nothing
          Set objPESQPADRAO = Nothing
          Set objCADFLXCAIXA = Nothing
          Unload Me
       Else
          Inclui
       End If
    End If
    
    
End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objPESQPADRAO = Nothing
    Set objCADFLXCAIXA = Nothing
    Unload Me
End Sub


Private Sub flxCadLctos_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If cTipOper = "I" Or cTipOper = "A" Then
       Dim lngBCO As Long
       lngBCO = 0
       If flxCadLctos.Rows > 1 And Len(Trim(flxCadLctos.TextMatrix(flxCadLctos.Row, 8))) > 0 Then lngBCO = CLng(flxCadLctos.TextMatrix(flxCadLctos.Row, 8))
    
       If KeyCode = vbKeyDelete And flxCadLctos.Rows > 1 Then AbateCaixa CCur(flxCadLctos.TextMatrix(flxCadLctos.Row, 5)), lngBCO
       If KeyCode = vbKeyDelete And flxCadLctos.Rows = 2 Then flxCadLctos.Rows = 1
       If KeyCode = vbKeyDelete And flxCadLctos.Rows > 2 Then flxCadLctos.RemoveItem flxCadLctos.Row
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdVoltar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()
   
   Set objBLBFunc = CreateObject("BLBCWS.clsFuncoes")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   Set objCADFLXCAIXA = CreateObject("CADFLUXCAIXA.clsCADFLUXCAIXA")
   
   objCADFLXCAIXA.FILIAL = FILIAL
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

    STFluxo.Tab = 0
    
End Sub

Private Sub Inclui()

    cmdVoltar.Enabled = True
    cmdAltera.Enabled = False
    CmdSalva.Enabled = True
    
    Me.Caption = "Fluxo de Caixa - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADFLUXCAIXA
    
    mskDTFLUXCAIXA.Text = "__/__/____"
    mskDTFLUXCAIXA.Text = Format(Now, "DD/MM/YYYY")
    
    lblStatus(2).Caption = "ABERTO"
    
    STFluxo.Tab = 0
    
    ConfGridFluxCaixa
    ConfGridLctos
    
    objCADFLXCAIXA.PreencheComboTipoPgto cboTIPOPGTO
    objCADFLXCAIXA.PreenchComboBancos cboBanco
    
    STFluxo.Enabled = False
    Frame2.Enabled = True
    Frame3.Enabled = True

End Sub

Private Sub mskDTFLUXCAIXA_GotFocus()
    ConfGridFluxCaixa
    ConfGridLctos
    STFluxo.Enabled = False
    objBLBFunc.SelecionaCampos mskDTFLUXCAIXA.Name, frmCADFLUXCAIXA
End Sub

Private Function CarregaCaixa() As Boolean

    Dim blContasaPagar   As Boolean
    Dim blContasaReceber As Boolean
    Dim intDIAS          As Integer
    
    CarregaCaixa = False
    
    ConfGridFluxCaixa
    ConfGridLctos
   
    '' Saldo Anterior
    flxFluxoCaixa.AddItem "" & vbTab & _
                          "Saldo Anterior dia " & Format(dtDTSALDOANT, "DD/MM/YYYY") & vbTab & _
                          "" & vbTab & _
                          Format(ccurSALDOANT, "#,##0.00")
                          
    
    blContasaReceber = PegaLctoContasaReceber '' Pega Contas á Receber
    blContasaPagar = PegaLctoContasaPagar     '' Pega Contas á Pagar
    
    flxFluxoCaixa.AddItem "" & vbTab & _
                          "Saldo Atual - " & mskDTFLUXCAIXA.Text & vbTab & _
                          "" & vbTab & _
                          Format(CalcSaldo, "#,##0.00")
    
    PintaCelula
    
    If blContasaReceber = False And _
       blContasaPagar = False Then
       Exit Function
    End If
    
    STFluxo.Enabled = True
    txtCODTIPPGTO.SetFocus
    CarregaCaixa = True
    
End Function

Private Sub ConfGridFluxCaixa()

    flxFluxoCaixa.Rows = 1
    flxFluxoCaixa.Cols = 9
    
    flxFluxoCaixa.TextMatrix(0, 0) = "Nº Doc"
    flxFluxoCaixa.TextMatrix(0, 1) = "Histórico de movimentos"
    flxFluxoCaixa.TextMatrix(0, 2) = "Data Lcto."
    flxFluxoCaixa.TextMatrix(0, 3) = "Valor"
    flxFluxoCaixa.TextMatrix(0, 4) = "Especie"
    flxFluxoCaixa.TextMatrix(0, 5) = "CodBco"
    flxFluxoCaixa.TextMatrix(0, 6) = "Operacao"
    flxFluxoCaixa.TextMatrix(0, 7) = "Saldo"
    flxFluxoCaixa.TextMatrix(0, 8) = "Sub.Total"
    
    flxFluxoCaixa.ColWidth(0) = 0
    flxFluxoCaixa.ColWidth(1) = 6000
    flxFluxoCaixa.ColWidth(2) = 1000
    flxFluxoCaixa.ColWidth(3) = 1500
    flxFluxoCaixa.ColWidth(4) = 0
    flxFluxoCaixa.ColWidth(5) = 0
    flxFluxoCaixa.ColWidth(6) = 0
    flxFluxoCaixa.ColWidth(7) = 1000
    flxFluxoCaixa.ColWidth(8) = 0

End Sub

Private Function PegaLctoContasaPagar() As Boolean
    
    PegaLctoContasaPagar = False
    
    Dim strDESCLCTO As String
    Dim ccurValor   As Currency
    Dim blAchou     As Boolean
    Dim I           As Integer
    Dim ccurTOTAL   As Currency
    
    '' Pegando Lotes
    sSql = "Select " & vbCrLf
    sSql = sSql & "       APGHADER.SGI_CODIGO   " & vbCrLf
    sSql = sSql & "      ,APGHADER.SGI_CODBCO   " & vbCrLf
    sSql = sSql & "      ,APGHADER.SGI_DATALOTE " & vbCrLf
    sSql = sSql & "      ,Sum(APGHADER.SGI_VLDOC) As SGI_VLDOC " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADLOTEHEADER APGHADER " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       APGHADER.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And APGHADER.SGI_DATALOTE = '" & Format(CDate(mskDTFLUXCAIXA.Text), "MM/DD/YYYY") & "'" & vbCrLf
    sSql = sSql & "   And APGHADER.SGI_STATUS   = 'B'" & vbCrLf
    sSql = sSql & " Group By " & vbCrLf
    sSql = sSql & "          APGHADER.SGI_CODIGO   " & vbCrLf
    sSql = sSql & "         ,APGHADER.SGI_CODBCO   " & vbCrLf
    sSql = sSql & "         ,APGHADER.SGI_DATALOTE " & vbCrLf
    sSql = sSql & " Order By " & vbCrLf
    sSql = sSql & "          APGHADER.SGI_DATALOTE " & vbCrLf
    sSql = sSql & "         ,APGHADER.SGI_CODBCO   "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    ccurTOTAL = 0
    Do While Not BREC.EOF
    
       strDESCLCTO = ""
       If BREC!SGI_CODBCO = 0 Then strDESCLCTO = "PAGAMENTOS DIVERSOS"
       
       '' ------------------------------------------------------
       '' Pega bancos
       sSql = "Select " & vbCrLf
       sSql = sSql & "       * " & vbCrLf
       sSql = sSql & "  From " & vbCrLf
       sSql = sSql & "       SGI_CADBANCOS " & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
       sSql = sSql & "   And SGI_CODIGO = " & BREC!SGI_CODBCO
       
       BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
       If Not BREC2.EOF Then strDESCLCTO = "PAGAMENTOS NO " & BREC2!SGI_DESCRICAO & " - AG: " & BREC2!SGI_AGENCIA & " - C/C : " & BREC2!SGI_CC
       BREC2.Close
       '' ------------------------------------------------------
       
       flxFluxoCaixa.AddItem BREC!SGI_CODIGO & vbTab & _
                             strDESCLCTO & vbTab & _
                             Format(BREC!SGI_DATALOTE, "DD/MM/YYYY") & vbTab & _
                             Format((BREC!SGI_VLDOC * -1), "#,##0.00") & vbTab & _
                             "BEBITO" & vbTab & _
                             BREC!SGI_CODBCO & vbTab & _
                             "1"
    
       
       ccurTOTAL = ccurTOTAL + BREC!SGI_VLDOC
       
       BREC.MoveNext
       PegaLctoContasaPagar = True
    Loop
    
    BREC.Close
    
    '' Baixas sem Lotes
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_DTPGTO " & vbCrLf
    sSql = sSql & "      ,SUM(SGI_VLPAGO) as SGI_VLPAGO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CONTASIAPG " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_STATUS = 'B'  " & vbCrLf
    sSql = sSql & "   And SGI_NLOTE is NUll " & vbCrLf
    sSql = sSql & "   And SGI_DTPGTO = '" & Format(CDate(mskDTFLUXCAIXA.Text), "MM/DD/YYYY") & "'" & vbCrLf
    sSql = sSql & " Group by SGI_DTPGTO "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    ccurValor = 0
    Do While Not BREC.EOF
       ccurValor = ccurValor + BREC!SGI_VLPAGO
       ccurTOTAL = ccurTOTAL + BREC!SGI_VLPAGO
       PegaLctoContasaPagar = True
       BREC.MoveNext
    Loop
    BREC.Close
    If ccurValor > 0 Then ccurValor = (ccurValor * -1)
    
    blAchou = False
    For I = 1 To (flxFluxoCaixa.Rows - 1)
        If flxFluxoCaixa.TextMatrix(I, 1) = "PAGAMENTOS DIVERSOS" Then
           flxFluxoCaixa.TextMatrix(I, 3) = Format((CCur(flxFluxoCaixa.TextMatrix(I, 3)) + ccurValor), "#,##0.00")
           blAchou = True
           Exit For
        End If
    Next I
    If blAchou = False And ccurValor < 0 Then
       
       flxFluxoCaixa.AddItem "" & vbTab & _
                             "PAGAMENTOS DIVERSOS" & vbTab & _
                             mskDTFLUXCAIXA.Text & vbTab & _
                             Format(ccurValor, "#,##0.00") & vbTab & _
                             "DEBITO" & vbTab & _
                             "0" & vbTab & _
                             "1"
    
    End If
    
    If ccurTOTAL > 0 Then
       flxFluxoCaixa.TextMatrix((flxFluxoCaixa.Rows - 1), 7) = Format((ccurTOTAL * -1), "#,##0.00")
    End If
    
    
End Function
 
Private Function PegaLctoContasaReceber() As Boolean
    
    PegaLctoContasaReceber = False
    
    Dim strDESCLCTO As String
    Dim ccurTOTAL   As Currency
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       HEADE.SGI_CODBCO " & vbCrLf
    sSql = sSql & "      ,ITENS.SGI_DTPGTO " & vbCrLf
    sSql = sSql & "      ,SUM(ITENS.SGI_VLPAGO) as SGI_VLPAGO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CONTASIARC ITENS " & vbCrLf
    sSql = sSql & "      ,SGI_CONTASHARC HEADE " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       ITENS.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And ITENS.SGI_DTPGTO = '" & Format(CDate(mskDTFLUXCAIXA.Text), "MM/DD/YYYY") & "'" & vbCrLf
    sSql = sSql & "   And HEADE.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And HEADE.SGI_CODIGO = ITENS.SGI_CODIGO " & vbCrLf
    sSql = sSql & " Group By " & vbCrLf
    sSql = sSql & "       HEADE.SGI_CODBCO " & vbCrLf
    sSql = sSql & "      ,ITENS.SGI_DTPGTO " & vbCrLf
    sSql = sSql & "Order By " & vbCrLf
    sSql = sSql & "         HEADE.SGI_CODBCO " & vbCrLf
    sSql = sSql & "        ,ITENS.SGI_DTPGTO "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    ccurTOTAL = 0
    Do While Not BREC.EOF
    
       strDESCLCTO = ""
       If BREC!SGI_CODBCO = 0 Then strDESCLCTO = "RECEBIMENTOS DIVERSOS"
       
       '' ------------------------------------------------------
       '' Pega bancos
       sSql = "Select " & vbCrLf
       sSql = sSql & "       * " & vbCrLf
       sSql = sSql & "  From " & vbCrLf
       sSql = sSql & "       SGI_CADBANCOS " & vbCrLf
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
       sSql = sSql & "   And SGI_CODIGO = " & BREC!SGI_CODBCO
       
       BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
       If Not BREC2.EOF Then strDESCLCTO = "RECEBIMENTOS NO " & BREC2!SGI_DESCRICAO & " - AG: " & BREC2!SGI_AGENCIA & " - C/C : " & BREC2!SGI_CC
       BREC2.Close
       '' ------------------------------------------------------
       
       flxFluxoCaixa.AddItem "" & vbTab & _
                             strDESCLCTO & vbTab & _
                             Format(BREC!SGI_DTPGTO, "DD/MM/YYYY") & vbTab & _
                             Format(BREC!SGI_VLPAGO, "#,##0.00") & vbTab & _
                             "CREDITO" & vbTab & _
                             BREC!SGI_CODBCO & vbTab & _
                             "2"
       
       
       ccurTOTAL = ccurTOTAL + BREC!SGI_VLPAGO
       
       PegaLctoContasaReceber = True
       BREC.MoveNext
       If BREC.EOF Then
          flxFluxoCaixa.TextMatrix((flxFluxoCaixa.Rows - 1), 7) = Format(ccurTOTAL, "#,##0.00")
       End If
       
    Loop
    
    BREC.Close
    
    

End Function

Private Function CalcSaldo() As Currency
    
    Dim I As Integer
    
    CalcSaldo = 0
    
    For I = 1 To (flxFluxoCaixa.Rows - 1)
        CalcSaldo = CalcSaldo + CCur(flxFluxoCaixa.TextMatrix(I, 3))
        flxFluxoCaixa.TextMatrix(I, 8) = Format(CalcSaldo, "#,##0.00")
    Next I
    
End Function

Private Sub PintaCelula()

    Dim I As Integer
    
    For I = 1 To (flxFluxoCaixa.Rows - 1)
        flxFluxoCaixa.Row = I
        flxFluxoCaixa.Col = 3
        If CCur(flxFluxoCaixa.TextMatrix(I, 3)) < 0 Then flxFluxoCaixa.CellForeColor = vbRed
        If CCur(flxFluxoCaixa.TextMatrix(I, 3)) >= 0 Then flxFluxoCaixa.CellForeColor = vbBlue
        
        flxFluxoCaixa.Col = 7
        If Len(Trim(flxFluxoCaixa.TextMatrix(I, 7))) > 0 Then
           If CCur(flxFluxoCaixa.TextMatrix(I, 7)) < 0 Then flxFluxoCaixa.CellForeColor = vbRed
           If CCur(flxFluxoCaixa.TextMatrix(I, 7)) >= 0 Then flxFluxoCaixa.CellForeColor = vbBlue
        End If
        
    Next I
    
    flxFluxoCaixa.Row = 1
    flxFluxoCaixa.Col = 1

End Sub

Private Sub ConfGridLctos()

    flxCadLctos.Rows = 1
    flxCadLctos.Cols = 9
    
    flxCadLctos.TextMatrix(0, 0) = ""
    flxCadLctos.TextMatrix(0, 1) = "Tipo Lcto."
    flxCadLctos.TextMatrix(0, 2) = "Bco."
    flxCadLctos.TextMatrix(0, 3) = "Data"
    flxCadLctos.TextMatrix(0, 4) = " Nº Doc."
    flxCadLctos.TextMatrix(0, 5) = "Valor"
    flxCadLctos.TextMatrix(0, 6) = "Sinal"
    flxCadLctos.TextMatrix(0, 7) = "Operacao"
    flxCadLctos.TextMatrix(0, 8) = "Banco"
    
    flxCadLctos.ColWidth(0) = 0
    flxCadLctos.ColWidth(1) = 3500
    flxCadLctos.ColWidth(2) = 3500
    flxCadLctos.ColWidth(3) = 1000
    flxCadLctos.ColWidth(4) = 1000
    flxCadLctos.ColWidth(5) = 1500
    flxCadLctos.ColWidth(6) = 0
    flxCadLctos.ColWidth(7) = 0
    flxCadLctos.ColWidth(8) = 0
    
    flxCadLctos.ColAlignment(1) = vbLeftJustify
    flxCadLctos.ColAlignment(2) = vbLeftJustify
    
End Sub

Private Sub txtBanco_GotFocus()
    objBLBFunc.SelecionaCampos txtBanco.Name, frmCADFLUXCAIXA
End Sub

Private Sub txtBanco_Validate(Cancel As Boolean)

    Dim I       As Integer
    Dim blAchou As Boolean
    
    If Len(Trim(txtBanco.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtBanco.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       txtBanco.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    blAchou = False
    For I = 0 To (cboBanco.ListCount - 1)
        If CInt(txtBanco.Text) = cboBanco.ItemData(I) Then
           blAchou = True
           cboBanco.ListIndex = I
        End If
    Next I
    
    If blAchou = False Then
       MsgBox "Este banco não existe !!!", vbOKOnly + vbCritical, "aviso"
       txtBanco.Text = ""
       cboBanco.ListIndex = -1
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub txtCODTIPPGTO_GotFocus()
    objBLBFunc.SelecionaCampos txtCODTIPPGTO.Name, frmCADFLUXCAIXA
End Sub

Private Sub txtCODTIPPGTO_Validate(Cancel As Boolean)

    Dim I As Integer
    
    If Len(Trim(txtCODTIPPGTO.Text)) = 0 Then Exit Sub
    
    If Not IsNumeric(txtCODTIPPGTO.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODTIPPGTO.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    cboTIPOPGTO.ListIndex = -1
    For I = 0 To (cboTIPOPGTO.ListCount - 1)
        If cboTIPOPGTO.ItemData(I) = Str(Val(txtCODTIPPGTO.Text)) Then cboTIPOPGTO.ListIndex = I
    Next I
    
    If cboTIPOPGTO.ListIndex = -1 Then
       MsgBox "Esta Condição de pagamento não existe !!!", vbOKOnly + vbCritical, "Aviso"
       txtCODTIPPGTO.Text = ""
       Cancel = True
       Exit Sub
    End If

End Sub

Private Sub InseriGridTipoPgto()
    
    Dim strSinal    As String
    Dim lngCODBANCO As Long
    
    If Len(Trim(txtCODTIPPGTO.Text)) = 0 Then
       MsgBox "Informe o tipo de operação !!!", vbOKOnly + vbExclamation, "Aviso"
       txtCODTIPPGTO.SetFocus
       Exit Sub
    End If
    If Not IsNumeric(txtCODTIPPGTO.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbExclamation, "Aviso"
       txtCODTIPPGTO.SetFocus
       Exit Sub
    End If
    If Len(Trim(txtValor.Text)) = 0 Then
       MsgBox "Informe o valor !!!", vbOKOnly + vbExclamation, "Aviso"
       txtValor.SetFocus
       Exit Sub
    End If
    If Not IsNumeric(txtValor.Text) Then
       MsgBox "Somente e permitido valores !!!", vbOKOnly + vbExclamation, "Aviso"
       txtValor.SetFocus
       Exit Sub
    End If
    
    '' Pega Sinal da Operação
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADTIPOPGTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & txtCODTIPPGTO.Text
    
    strSinal = ""
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then strSinal = BREC!SGI_SINAL
    BREC.Close
    
    flxCadLctos.AddItem "" & vbTab & _
                        Trim(cboTIPOPGTO.Text) & vbTab & _
                        Trim(cboBanco.Text) & vbTab & _
                        mskDTFLUXCAIXA.Text & vbTab & _
                        txtNUMDOC.Text & vbTab & _
                        IIf(strSinal = "-", strSinal & txtValor.Text, txtValor.Text) & vbTab & _
                        strSinal & vbTab & _
                        txtCODTIPPGTO.Text & vbTab & _
                        txtBanco.Text
                           
    
    lngCODBANCO = 0
    If Len(Trim(txtBanco.Text)) > 0 Then lngCODBANCO = CLng(txtBanco.Text)
    CalcLctosCaixa CCur(strSinal & txtValor.Text), lngCODBANCO
    
    flxFluxoCaixa.AddItem "" & vbTab & _
                          "Saldo Atual - " & mskDTFLUXCAIXA.Text & vbTab & _
                          "" & vbTab & _
                          Format(CalcSaldo, "#,##0.00")
    
    
    txtCODTIPPGTO.Text = ""
    cboTIPOPGTO.ListIndex = -1
    txtNUMDOC.Text = ""
    txtValor.Text = ""
    txtBanco.Text = ""
    cboBanco.ListIndex = -1
    
    txtCODTIPPGTO.SetFocus

    
End Sub

Private Sub txtValor_GotFocus()
    objBLBFunc.SelecionaCampos txtValor.Name, frmCADFLUXCAIXA
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    objBLBFunc.SoNumeroPonto KeyAscii, txtValor.Text
End Sub

Private Sub txtValor_Validate(Cancel As Boolean)

    If Len(Trim(txtValor.Text)) = 0 Then Exit Sub

    If Not IsNumeric(txtValor.Text) Then
       MsgBox "Somente é permitido numeros !!!", vbOKOnly + vbCritical, "aviso"
       Cancel = True
       Exit Sub
    End If
    
    If Val(txtValor.Text) < 0 Then
       MsgBox "Não é permitido numero negativo !!!", vbOKOnly + vbCritical, "aviso"
       txtValor.Text = ""
       Cancel = True
       Exit Sub
    End If
    
    txtValor.Text = Format(txtValor.Text, "#,##0.00")

End Sub

Private Sub PintaCelulaLcto()

    Dim I As Integer
    
    For I = 1 To (flxCadLctos.Rows - 1)
        flxCadLctos.Row = I
        flxCadLctos.Col = 5
        If Len(Trim(flxCadLctos.TextMatrix(I, 5))) > 0 Then
           If CCur(flxCadLctos.TextMatrix(I, 5)) < 0 Then flxCadLctos.CellForeColor = vbRed
           If CCur(flxCadLctos.TextMatrix(I, 5)) >= 0 Then flxCadLctos.CellForeColor = vbBlue
        End If
    Next I
    
    If flxCadLctos.Rows > 1 Then
       flxCadLctos.Row = 1
       flxCadLctos.Col = 1
    End If

End Sub

Private Sub CalcLctosCaixa(ccurCalor As Currency, lngCODBANCO As Long)
    
    Dim I               As Integer
    Dim blAchouRecebDiv As Boolean
    Dim blAchouRecebBco As Boolean
    Dim blAchouPagto    As Boolean
    Dim blAchouPagtoDiv As Boolean
    Dim strBancos       As String
    
    blAchouRecebDiv = False
    blAchouRecebBco = False
    For I = 1 To (flxFluxoCaixa.Rows - 1)
        If ccurCalor > 0 Then
           If lngCODBANCO > 0 And CCur(flxFluxoCaixa.TextMatrix(I, 3)) > 0 Then
              If Len(Trim(flxFluxoCaixa.TextMatrix(I, 5))) > 0 Then
                 If CLng(flxFluxoCaixa.TextMatrix(I, 5)) = lngCODBANCO Then
                    flxFluxoCaixa.TextMatrix(I, 3) = Format(CCur(flxFluxoCaixa.TextMatrix(I, 3)) + ccurCalor, "#,##0.00")
                    blAchouRecebBco = True
                    Exit For
                 End If
              End If
           Else
              If flxFluxoCaixa.TextMatrix(I, 1) = "RECEBIMENTOS DIVERSOS" Then
                 flxFluxoCaixa.TextMatrix(I, 3) = Format(CCur(flxFluxoCaixa.TextMatrix(I, 3)) + ccurCalor, "#,##0.00")
                 blAchouRecebDiv = True
                 Exit For
              End If
           End If
        Else
           If lngCODBANCO > 0 And CCur(flxFluxoCaixa.TextMatrix(I, 3)) < 0 Then
              If Len(Trim(flxFluxoCaixa.TextMatrix(I, 5))) > 0 Then
                 If CLng(flxFluxoCaixa.TextMatrix(I, 5)) = lngCODBANCO Then
                    flxFluxoCaixa.TextMatrix(I, 3) = Format(CCur(flxFluxoCaixa.TextMatrix(I, 3)) + ccurCalor, "#,##0.00")
                    blAchouPagtoDiv = True
                    Exit For
                 End If
              End If
           Else
              If flxFluxoCaixa.TextMatrix(I, 1) = "PAGAMENTOS DIVERSOS" Then
                 flxFluxoCaixa.TextMatrix(I, 3) = Format(CCur(flxFluxoCaixa.TextMatrix(I, 3)) + ccurCalor, "#,##0.00")
                 blAchouPagto = True
                 Exit For
              End If
           End If
        End If
    Next I
    
    flxFluxoCaixa.RemoveItem (flxFluxoCaixa.Rows - 1)
    
    If ccurCalor > 0 Then
       If blAchouRecebDiv = False And lngCODBANCO = 0 Then
          flxFluxoCaixa.AddItem "" & vbTab & _
                                "RECEBIMENTOS DIVERSOS" & vbTab & _
                                mskDTFLUXCAIXA.Text & vbTab & _
                                Format(ccurCalor, "#,##0.00") & vbTab & _
                                "CREDITO" & vbTab & _
                                "0" & vbTab & _
                                "2"
       End If
    
       If blAchouRecebBco = False And lngCODBANCO > 0 Then
    
          '' ------------------------------------------------------
          '' Bancos
          sSql = "Select" & vbCrLf
          sSql = sSql & "       *" & vbCrLf
          sSql = sSql & "  From" & vbCrLf
          sSql = sSql & "       SGI_CADBANCOS" & vbCrLf
          sSql = sSql & " Where" & vbCrLf
          sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
          sSql = sSql & "   And SGI_CODIGO = " & lngCODBANCO
       
          BREC.Open sSql, adoBanco_Dados, adOpenDynamic
          If Not BREC.EOF Then strBancos = "RECEBIMENTOS NO " & BREC!SGI_DESCRICAO & " - AG: " & BREC!SGI_AGENCIA & " - C/C : " & BREC!SGI_CC
          BREC.Close
       
          flxFluxoCaixa.AddItem "" & vbTab & _
                                strBancos & vbTab & _
                                mskDTFLUXCAIXA.Text & vbTab & _
                                Format(ccurCalor, "#,##0.00") & vbTab & _
                                "CREDITO" & vbTab & _
                                txtBanco.Text & vbTab & _
                                "2"
          '' ------------------------------------------------------
                             
       End If
    End If
    
    If ccurCalor < 0 Then
       If blAchouPagto = False And lngCODBANCO = 0 Then
          flxFluxoCaixa.AddItem "" & vbTab & _
                                "PAGAMENTOS DIVERSOS" & vbTab & _
                                mskDTFLUXCAIXA.Text & vbTab & _
                                Format(ccurCalor, "#,##0.00") & vbTab & _
                                "DEBITO" & vbTab & _
                                "0" & vbTab & _
                                "1"
       End If
       If blAchouPagtoDiv = False And lngCODBANCO > 0 Then
          
          '' ------------------------------------------------------
          '' Bancos
          sSql = "Select" & vbCrLf
          sSql = sSql & "       *" & vbCrLf
          sSql = sSql & "  From" & vbCrLf
          sSql = sSql & "       SGI_CADBANCOS" & vbCrLf
          sSql = sSql & " Where" & vbCrLf
          sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
          sSql = sSql & "   And SGI_CODIGO = " & lngCODBANCO
       
          BREC.Open sSql, adoBanco_Dados, adOpenDynamic
          If Not BREC.EOF Then strBancos = "PAGAMENTOS NO " & BREC!SGI_DESCRICAO & " - AG: " & BREC!SGI_AGENCIA & " - C/C : " & BREC!SGI_CC
          BREC.Close
       
          flxFluxoCaixa.AddItem "" & vbTab & _
                                strBancos & vbTab & _
                                mskDTFLUXCAIXA.Text & vbTab & _
                                Format(ccurCalor, "#,##0.00") & vbTab & _
                                "DEBITO" & vbTab & _
                                txtBanco.Text & vbTab & _
                                "1"
          '' ------------------------------------------------------
       
       End If
    End If
    
    
End Sub

Private Sub AbateCaixa(ccurCalor As Currency, lngCODBANCO As Long)

    Dim I               As Integer
    
    For I = 1 To (flxFluxoCaixa.Rows - 1)
        If ccurCalor > 0 Then
           If lngCODBANCO > 0 And CCur(flxFluxoCaixa.TextMatrix(I, 3)) > 0 Then
              If Len(Trim(flxFluxoCaixa.TextMatrix(I, 5))) > 0 Then
                 If CLng(flxFluxoCaixa.TextMatrix(I, 5)) = lngCODBANCO Then
                    flxFluxoCaixa.TextMatrix(I, 3) = Format(CCur(flxFluxoCaixa.TextMatrix(I, 3)) + (ccurCalor * -1), "#,##0.00")
                    Exit For
                 End If
              End If
           Else
              If flxFluxoCaixa.TextMatrix(I, 1) = "RECEBIMENTOS DIVERSOS" Then
                 flxFluxoCaixa.TextMatrix(I, 3) = Format(CCur(flxFluxoCaixa.TextMatrix(I, 3)) + (ccurCalor * -1), "#,##0.00")
                 Exit For
              End If
           End If
        Else
           If lngCODBANCO > 0 And CCur(flxFluxoCaixa.TextMatrix(I, 3)) < 0 Then
              If Len(Trim(flxFluxoCaixa.TextMatrix(I, 5))) > 0 Then
                 If CLng(flxFluxoCaixa.TextMatrix(I, 5)) = lngCODBANCO Then
                    flxFluxoCaixa.TextMatrix(I, 3) = Format(CCur(flxFluxoCaixa.TextMatrix(I, 3)) + (ccurCalor * -1), "#,##0.00")
                    Exit For
                 End If
              End If
           Else
              If flxFluxoCaixa.TextMatrix(I, 1) = "PAGAMENTOS DIVERSOS" Then
                 flxFluxoCaixa.TextMatrix(I, 3) = Format(CCur(flxFluxoCaixa.TextMatrix(I, 3)) + (ccurCalor * -1), "#,##0.00")
                 Exit For
              End If
           End If
        End If
    Next I
    
    '' ----------------------------------------------------------------------------
    flxFluxoCaixa.RemoveItem (flxFluxoCaixa.Rows - 1)
    For I = 1 To (flxFluxoCaixa.Rows - 1)
        If Len(Trim(flxFluxoCaixa.TextMatrix(I, 4))) > 0 And CCur(flxFluxoCaixa.TextMatrix(I, 3)) = 0 Then
           flxFluxoCaixa.RemoveItem I
           Exit For
        End If
    Next I
    
    flxFluxoCaixa.AddItem "" & vbTab & _
                          "Saldo Atual - " & mskDTFLUXCAIXA.Text & vbTab & _
                          "" & vbTab & _
                          Format(CalcSaldo, "#,##0.00")
    PintaCelula
    '' ----------------------------------------------------------------------------

End Sub

Private Function ValidaCampos() As Boolean

     ValidaCampos = False
     
     If flxFluxoCaixa.Rows = 1 Then
        MsgBox "Não foi calculado o Caixa !!!", vbOKOnly + vbExclamation, "Aviso"
        mskDTFLUXCAIXA.SetFocus
        Exit Function
     End If
     
     ValidaCampos = True
     
End Function


Private Function PegaSaldoAnt() As Currency
    
    PegaSaldoAnt = 0
    
    sSql = "Select " & vbTab
    sSql = sSql & "       SGI_SALDATU " & vbTab
    sSql = sSql & "      ,SGI_DATA     " & vbTab
    sSql = sSql & "  From " & vbTab
    sSql = sSql & "       SGI_CADFLXCXHEADER " & vbTab
    sSql = sSql & " Where " & vbTab
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbTab
    sSql = sSql & " Order by SGI_DATA DESC"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    dtDTSALDOANT = Now
    
    If Not BREC.EOF Then
       dtDTSALDOANT = BREC!SGI_DATA
       PegaSaldoAnt = BREC!SGI_SALDATU
    End If
    BREC.Close
    
End Function

Private Function VerificaCaixa() As Boolean

    VerificaCaixa = False
    
    sSql = "Select " & vbTab
    sSql = sSql & "       * " & vbTab
    sSql = sSql & "  From " & vbTab
    sSql = sSql & "       SGI_CADFLXCXHEADER " & vbTab
    sSql = sSql & " Where " & vbTab
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbTab
    sSql = sSql & "   And SGI_DATA   = '" & Format(CDate(mskDTFLUXCAIXA.Text), "MM/DD/YYYY") & "'"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then VerificaCaixa = True
    BREC.Close
    
End Function

Private Function CaixaFechado() As Boolean

    CaixaFechado = False
    
    Dim dtDATAANT As Date
    Dim intDIAS   As Integer
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "      * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "      SGI_CADFLXCXHEADER " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "  And SGI_STATUS = 'ABERTO'" & vbCrLf
    sSql = sSql & "Order by SGI_DATA DESC "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
       MsgBox "Existe caixa aberto favor fechar !!!", vbOKOnly + vbExclamation, "Aviso"
       mskDTFLUXCAIXA.Text = Format(Now, "DD/MM/YYYY")
       mskDTFLUXCAIXA.SetFocus
       CaixaFechado = True
    End If
    BREC.Close
    
    If CaixaFechado = True Then Exit Function

    
    sSql = "Select " & vbCrLf
    sSql = sSql & "      * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "      SGI_CADFLXCXHEADER " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "  And SGI_STATUS = 'FECHADO'" & vbCrLf
    sSql = sSql & "Order by SGI_DATA DESC "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    intDIAS = 0
    dtDATAANT = Now
    If Not BREC.EOF Then
       dtDATAANT = BREC!SGI_DATA
       
       If dtDATAANT < CDate(mskDTFLUXCAIXA.Text) Then intDIAS = CDate(mskDTFLUXCAIXA.Text) - dtDATAANT
       If dtDATAANT > CDate(mskDTFLUXCAIXA.Text) Then intDIAS = CDate(mskDTFLUXCAIXA.Text) - dtDATAANT
       
       If intDIAS >= 2 Then
          MsgBox "O caixa do dia " & Format(((CDate(mskDTFLUXCAIXA.Text) - intDIAS) + 1), "DD/MM/YYYY") & " não foi calculado !!!", vbOKOnly + vbExclamation, "Aviso"
          mskDTFLUXCAIXA.Text = Format(Now, "DD/MM/YYYY")
          mskDTFLUXCAIXA.SetFocus
          CaixaFechado = True
       End If
       
       If intDIAS < 0 Then
          MsgBox "A data digitada é retroativa !!!", vbOKOnly + vbExclamation, "Aviso"
          mskDTFLUXCAIXA.Text = Format(Now, "DD/MM/YYYY")
          mskDTFLUXCAIXA.SetFocus
          CaixaFechado = True
       End If
       
    End If
    BREC.Close
    
    If CaixaFechado = True Then Exit Function
    
End Function

Private Sub Consulta()

    cmdVoltar.Enabled = True
    cmdAltera.Enabled = True
    CmdSalva.Enabled = False
    
    Dim I         As Integer
    Dim J         As Integer
    Dim strLcto1  As String
    Dim strLcto2  As String
    
    Me.Caption = "Fluxo de Caixa - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADFLUXCAIXA
    lblStatus(0).Caption = ""
    lblStatus(2).Caption = ""
    
    
    mskDTFLUXCAIXA.Text = "__/__/____"
    ConfGridFluxCaixa
    ConfGridLctos
    
    objCADFLXCAIXA.PreencheComboTipoPgto cboTIPOPGTO
    objCADFLXCAIXA.PreenchComboBancos cboBanco
    
    STFluxo.Enabled = True
    Frame2.Enabled = False
    Frame3.Enabled = False
    
    objCADFLXCAIXA.CODFLXCAIXA = iCodigo
    
    If objCADFLXCAIXA.Carrega_campos = True Then
    
       lblStatus(0).Caption = objCADFLXCAIXA.CODFLXCAIXA
       lblStatus(2).Caption = objCADFLXCAIXA.STATUS
       
       mskDTFLUXCAIXA.Text = Format(objCADFLXCAIXA.DATAFLXCAIXA, "DD/MM/YYYY")
       
       arrLANCAMENTOS = objCADFLXCAIXA.LCTOFLXCAIXA
       arrFLUXOCAIXA = objCADFLXCAIXA.FLUXOCAIXA
       
       ccurSALDOANT = objCADFLXCAIXA.SALDOANT
       dtDTSALDOANT = (objCADFLXCAIXA.DATAFLXCAIXA - 1)
       
       '' Saldo Anterior
       flxFluxoCaixa.AddItem "" & vbTab & _
                             "Saldo Anterior dia " & Format(dtDTSALDOANT, "DD/MM/YYYY") & vbTab & _
                             "" & vbTab & _
                             Format(ccurSALDOANT, "#,##0.00")
       
       '' Grid Lançamentos
       If IsArray(arrLANCAMENTOS) Then
          For I = 1 To UBound(arrLANCAMENTOS)
          
              strLcto1 = ""
              For J = 0 To (cboTIPOPGTO.ListCount - 1)
                  If cboTIPOPGTO.ItemData(J) = arrLANCAMENTOS(I, 1) Then strLcto1 = cboTIPOPGTO.List(J)
              Next J
              
              strLcto2 = ""
              For J = 0 To (cboBanco.ListCount - 1)
                  If cboBanco.ItemData(J) = arrLANCAMENTOS(I, 2) Then strLcto2 = cboBanco.List(J)
              Next J
              
              flxCadLctos.AddItem "" & vbTab & _
                                  Trim(strLcto1) & vbTab & _
                                  Trim(strLcto2) & vbTab & _
                                  Format(arrLANCAMENTOS(I, 3), "DD/MM/YYYY") & vbTab & _
                                  arrLANCAMENTOS(I, 4) & vbTab & _
                                  Format(arrLANCAMENTOS(I, 5), "#,##0.00") & vbTab & _
                                  IIf(CCur(arrLANCAMENTOS(I, 5)) > 0, "+", "-") & vbTab & _
                                  arrLANCAMENTOS(I, 1) & vbTab & _
                                  arrLANCAMENTOS(I, 2)
          
          Next I
       End If
       
       '' Grid Caixa
       If IsArray(arrFLUXOCAIXA) Then
          For I = 1 To UBound(arrFLUXOCAIXA)
          
              flxFluxoCaixa.AddItem "" & vbTab & _
                                    arrFLUXOCAIXA(I, 1) & vbTab & _
                                    Format(arrFLUXOCAIXA(I, 2), "DD/MM/YYYY") & vbTab & _
                                    Format(arrFLUXOCAIXA(I, 3), "#,##0.00") & vbTab & _
                                    IIf(arrFLUXOCAIXA(I, 3) > 0, "CREDITO", "DEBITO") & vbTab & _
                                    arrFLUXOCAIXA(I, 4) & vbTab & _
                                    IIf(arrFLUXOCAIXA(I, 3) > 0, "2", "1")
                                    
          
          Next I
       End If
       
       flxFluxoCaixa.AddItem "" & vbTab & _
                             "Saldo Atual - " & mskDTFLUXCAIXA.Text & vbTab & _
                             "" & vbTab & _
                             Format(CalcSaldo, "#,##0.00")
       PintaCelula
       PintaCelulaLcto
       
    End If

End Sub

Private Sub Altera()

    cmdVoltar.Enabled = True
    cmdAltera.Enabled = False
    CmdSalva.Enabled = True
    
    Dim I         As Integer
    Dim J         As Integer
    Dim strLcto1  As String
    Dim strLcto2  As String
    
    Me.Caption = "Fluxo de Caixa - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADFLUXCAIXA
    lblStatus(0).Caption = ""
    lblStatus(2).Caption = ""
    
    
    mskDTFLUXCAIXA.Text = "__/__/____"
    ConfGridFluxCaixa
    ConfGridLctos
    
    objCADFLXCAIXA.PreencheComboTipoPgto cboTIPOPGTO
    objCADFLXCAIXA.PreenchComboBancos cboBanco
    
    STFluxo.Enabled = True
    Frame2.Enabled = False
    Frame3.Enabled = True
    
    objCADFLXCAIXA.CODFLXCAIXA = iCodigo
    
    If objCADFLXCAIXA.Carrega_campos = True Then
    
       lblStatus(0).Caption = objCADFLXCAIXA.CODFLXCAIXA
       lblStatus(2).Caption = objCADFLXCAIXA.STATUS
       
       mskDTFLUXCAIXA.Text = Format(objCADFLXCAIXA.DATAFLXCAIXA, "DD/MM/YYYY")
       
       arrLANCAMENTOS = objCADFLXCAIXA.LCTOFLXCAIXA
       arrFLUXOCAIXA = objCADFLXCAIXA.FLUXOCAIXA
       
       ccurSALDOANT = objCADFLXCAIXA.SALDOANT
       dtDTSALDOANT = (objCADFLXCAIXA.DATAFLXCAIXA - 1)
       
       '' Saldo Anterior
       flxFluxoCaixa.AddItem "" & vbTab & _
                             "Saldo Anterior dia " & Format(dtDTSALDOANT, "DD/MM/YYYY") & vbTab & _
                             "" & vbTab & _
                             Format(ccurSALDOANT, "#,##0.00")
       
       '' Grid Lançamentos
       If IsArray(arrLANCAMENTOS) Then
          For I = 1 To UBound(arrLANCAMENTOS)
          
              strLcto1 = ""
              For J = 0 To (cboTIPOPGTO.ListCount - 1)
                  If cboTIPOPGTO.ItemData(J) = arrLANCAMENTOS(I, 1) Then strLcto1 = cboTIPOPGTO.List(J)
              Next J
              
              strLcto2 = ""
              For J = 0 To (cboBanco.ListCount - 1)
                  If cboBanco.ItemData(J) = arrLANCAMENTOS(I, 2) Then strLcto2 = cboBanco.List(J)
              Next J
              
              flxCadLctos.AddItem "" & vbTab & _
                                  Trim(strLcto1) & vbTab & _
                                  Trim(strLcto2) & vbTab & _
                                  Format(arrLANCAMENTOS(I, 3), "DD/MM/YYYY") & vbTab & _
                                  arrLANCAMENTOS(I, 4) & vbTab & _
                                  Format(arrLANCAMENTOS(I, 5), "#,##0.00") & vbTab & _
                                  IIf(CCur(arrLANCAMENTOS(I, 5)) > 0, "+", "-") & vbTab & _
                                  arrLANCAMENTOS(I, 1) & vbTab & _
                                  arrLANCAMENTOS(I, 2)
          
          
          Next I
       End If
       
       '' Grid Caixa
       If IsArray(arrFLUXOCAIXA) Then
          For I = 1 To UBound(arrFLUXOCAIXA)
          
              flxFluxoCaixa.AddItem "" & vbTab & _
                                    arrFLUXOCAIXA(I, 1) & vbTab & _
                                    Format(arrFLUXOCAIXA(I, 2), "DD/MM/YYYY") & vbTab & _
                                    Format(arrFLUXOCAIXA(I, 3), "#,##0.00") & vbTab & _
                                    IIf(arrFLUXOCAIXA(I, 3) > 0, "CREDITO", "DEBITO") & vbTab & _
                                    arrFLUXOCAIXA(I, 4) & vbTab & _
                                    IIf(arrFLUXOCAIXA(I, 3) > 0, "2", "1")
                                    
          
          Next I
       End If
       
       flxFluxoCaixa.AddItem "" & vbTab & _
                             "Saldo Atual - " & mskDTFLUXCAIXA.Text & vbTab & _
                             "" & vbTab & _
                             Format(CalcSaldo, "#,##0.00")
       PintaCelula
       PintaCelulaLcto
       
    End If


End Sub
