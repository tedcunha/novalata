VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCADZONAGEO 
   Caption         =   "Cadastro de Zona Geografica"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8250
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   8250
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   7858
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
      TabCaption(0)   =   "Dados"
      TabPicture(0)   =   "frmCADZONAGEO.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Clientes"
      TabPicture(1)   =   "frmCADZONAGEO.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame5 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   15
         Top             =   360
         Width           =   7935
         Begin MSFlexGridLib.MSFlexGrid flxCLIENTES 
            Height          =   3615
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   6376
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2055
         Left            =   120
         TabIndex        =   13
         Top             =   2280
         Width           =   7935
         Begin MSFlexGridLib.MSFlexGrid flxCidade 
            Height          =   1695
            Left            =   3960
            TabIndex        =   18
            Top             =   240
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   2990
            _Version        =   393216
            FixedCols       =   0
            Appearance      =   0
         End
         Begin MSFlexGridLib.MSFlexGrid flxESTADOS 
            Height          =   1695
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   2990
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   7935
         Begin VB.CommandButton cmdIncTransp 
            Height          =   315
            Left            =   2280
            Picture         =   "frmCADZONAGEO.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   240
            Width           =   375
         End
         Begin VB.ComboBox cboESTADO 
            Height          =   315
            Left            =   1560
            TabIndex        =   8
            Text            =   "cboESTADO"
            Top             =   240
            Width           =   750
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
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
            TabIndex        =   7
            Top             =   300
            Width           =   600
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   7935
         Begin VB.TextBox txtDescricao 
            Height          =   285
            Left            =   1560
            MaxLength       =   30
            TabIndex        =   5
            Text            =   "txtDescricao"
            Top             =   600
            Width           =   5775
         End
         Begin VB.TextBox txtCodigo 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   3
            Text            =   "txtCodigo"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Descrição"
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
            Left            =   120
            TabIndex        =   4
            Top             =   600
            Width           =   870
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
            TabIndex        =   2
            Top             =   240
            Width           =   600
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   8175
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
         Picture         =   "frmCADZONAGEO.frx":013A
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Picture         =   "frmCADZONAGEO.frx":066C
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Picture         =   "frmCADZONAGEO.frx":076E
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCADZONAGEO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho         As String
Public Linha            As Variant
Public cTipOper         As String
Public iCodigo          As Integer
Public FILIAL           As Integer
Public strAcesso        As String
Public strMODPAI        As String
Dim objBLBFunc          As Object
Dim objCADZONAGEO       As Object
Dim objPESQPADRAO       As Object
Dim arrESTADOS          As Variant

Private Sub cboESTADO_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboESTADO, KeyAscii
End Sub

Private Sub cmdAltera_Click()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame3.Enabled = True
    
    Me.Caption = "Cadastro de Zona Geografica - [ ALTERAÇÃO ]"
    
    txtDescricao.SetFocus
    
    cTipOper = "A"

End Sub

Private Sub cmdIncTransp_Click()
    If cTipOper = "I" Or cTipOper = "A" Then IncGridEstado
End Sub

Private Sub CmdSalva_Click()

    Dim I As Integer
    
    If ValidaCampos = False Then Exit Sub
    
    If cTipOper = "I" Then objCADZONAGEO.CODIGO = objCADZONAGEO.Gera_Codigo(Me.Name)
    objCADZONAGEO.DESCRI = Trim(txtDescricao.Text)
    
    If (flxESTADOS.Rows - 1) > 0 Then
       ReDim arrESTADOS(1 To (flxESTADOS.Rows - 1))
       For I = 1 To (flxESTADOS.Rows - 1)
           arrESTADOS(I) = flxESTADOS.TextMatrix(I, 0)
       Next I
       objCADZONAGEO.ESTADOS = arrESTADOS
    Else
       arrESTADOS = Empty
       objCADZONAGEO.ESTADOS = arrESTADOS
    End If
        
    If objCADZONAGEO.GRAVA(cTipOper) = False Then Exit Sub
    MsgBox "A zona geografica foi " & IIf(cTipOper = "I", "inclusa", IIf(cTipOper = "A", "alterada", "")) + " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
    If objCADZONAGEO.Atualiza(cTipOper, Str(objCADZONAGEO.CODIGO), FILIAL, Me.Name) = False Then Exit Sub
       
    If cTipOper = "I" Then
       Set objBLBFunc = Nothing
       Set objCADZONAGEO = Nothing
       Set objPESQPADRAO = Nothing
       Unload Me
    End If

End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objCADZONAGEO = Nothing
    Set objPESQPADRAO = Nothing
    Unload Me
End Sub
Private Sub flxCidade_DblClick()
    If (flxCidade.Rows - 1) > 0 Then
       SSTab1.Tab = 1
       Call PopGridClientesCidade(flxCidade.TextMatrix(flxCidade.Row, 1), CLng(flxESTADOS.TextMatrix(flxESTADOS.Row, 0)))
    End If
End Sub

Private Sub flxESTADOS_DblClick()
    If (flxESTADOS.Rows - 1) > 0 Then
       SSTab1.Tab = 1
       Call PopGridClientes(CLng(flxESTADOS.TextMatrix(flxESTADOS.Row, 0)))
    End If
End Sub

Private Sub flxESTADOS_KeyDown(KeyCode As Integer, Shift As Integer)
    If (flxESTADOS.Rows - 1) > 0 Then
       If (flxESTADOS.Rows - 1) = 2 Then flxESTADOS.Rows = 1
       If (flxESTADOS.Rows - 1) > 2 Then flxESTADOS.RemoveItem flxESTADOS.Row
    End If
End Sub

Private Sub flxESTADOS_RowColChange()
    If (flxESTADOS.Rows - 1) > 0 Then
       Call PopGridClientes(CLng(flxESTADOS.TextMatrix(flxESTADOS.Row, 0)))
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
   Set objCADZONAGEO = CreateObject("CADZONAGEO.clsCADZONAGEO")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
      
   objCADZONAGEO.FILIAL = FILIAL
   
   Set adoBanco_Dados = objBLBFunc.Banco_Dados(Linha)

   If adoBanco_Dados.State = 0 Then
      MsgBox "Nâo foi possivel abrir o banco de dados !!!", vbOKOnly + vbCritical, "Aviso"
      Exit Sub
   End If
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta

End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame3.Enabled = True
   
    Me.Caption = "Cadastro de Zona Geografica - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADZONAGEO
    objBLBFunc.Preenche_Estado cboESTADO
    
    txtCodigo.Text = ""
    
    ConfGridEstado
    ConfGridCliente
    ConfGridCidade
    
End Sub

Private Sub SSTab1_DblClick()
    If (flxESTADOS.Rows - 1) > 0 Then
       Call PopGridClientes(CLng(flxESTADOS.TextMatrix(flxESTADOS.RowSel, 0)))
    End If
End Sub

Private Sub txtDescricao_GotFocus()
    objBLBFunc.SelecionaCampos txtDescricao.Name, frmCADZONAGEO
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = objBLBFunc.Maiuscula(KeyAscii)
End Sub

Private Function ValidaCampos() As Boolean

     Dim I As Integer
     
     ValidaCampos = False
     
     If Trim(Len(txtDescricao.Text)) = 0 Then
        MsgBox "Nome não pode ser vázio !!!", vbOKOnly + vbCritical, "Aviso"
        txtDescricao.SetFocus
        Exit Function
     End If
     
     If cTipOper = "I" Then
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       * " & vbCrLf
        sSql = sSql & "  from " & vbCrLf
        sSql = sSql & "       SGI_CADZONAGEO  " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_DESCRI = '" & Trim(txtDescricao) & "'" & vbCrLf
        sSql = sSql & "   And SGI_FILIAL = " & FILIAL
        
        BREC.Open sSql, adoBanco_Dados
        
        If Not BREC.EOF Then
           MsgBox "Esta zona geografica já existe !!!", vbOKOnly + vbCritical, "Aviso"
           txtDescricao.SetFocus
           BREC.Close
           Exit Function
        End If
        BREC.Close
        
        '' Verifica se já existe este estado cadastrado
        For I = 1 To (flxESTADOS.Rows - 1)
        
            sSql = "Select " & vbCrLf
            sSql = sSql & "       * " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CLIZONAGEO " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
            sSql = sSql & "   And SGI_CODESTADO = " & flxESTADOS.TextMatrix(I, 0)
            
            BREC.Open sSql, adoBanco_Dados, adOpenDynamic
            If Not BREC.EOF Then
               MsgBox "O estado " & Trim(flxESTADOS.TextMatrix(I, 1)) & " Já esta relacionado para outra area geografica !!!", vbOKOnly + vbExclamation, "Aviso"
               BREC.Close
               Exit Function
            End If
            BREC.Close
        Next I
     
     End If
     
     If cTipOper = "A" Then
        If objCADZONAGEO.DESCRI <> txtDescricao.Text Then
        
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  from " & vbCrLf
           sSql = sSql & "       SGI_CADZONAGEO " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_DESCRI = '" & Trim(txtDescricao.Text) & "'" & vbCrLf
           sSql = sSql & "   And SGI_FILIAL = " & FILIAL
           
           BREC.Open sSql, adoBanco_Dados
           If Not BREC.EOF Then
              MsgBox "Esta zona geografia já existe !!!", vbOKOnly + vbCritical, "Aviso"
              txtDescricao.Text = objCADZONAGEO.DESCRI
              txtDescricao.SetFocus
              BREC.Close
              Exit Function
           End If
           BREC.Close
        
        End If
     
        '' Verifica se já existe este estado cadastrado
        For I = 1 To (flxESTADOS.Rows - 1)
        
            sSql = "Select " & vbCrLf
            sSql = sSql & "       * " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_CLIZONAGEO " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
            sSql = sSql & "   And SGI_CODESTADO = " & flxESTADOS.TextMatrix(I, 0)
            
            BREC.Open sSql, adoBanco_Dados, adOpenDynamic
            Do While Not BREC.EOF
               If BREC!SGI_CODIGO <> CLng(txtCodigo.Text) Then
                  MsgBox "O estado " & Trim(flxESTADOS.TextMatrix(I, 1)) & " Já esta relacionado para outra area geografica !!!", vbOKOnly + vbExclamation, "Aviso"
                  BREC.Close
                  Exit Function
               End If
               BREC.MoveNext
            Loop
            BREC.Close
        Next I
     
     
     End If
     ValidaCampos = True
     
End Function

Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    Frame2.Enabled = False
    Frame3.Enabled = False
    
    Me.Caption = "Cadastro de Zona Geografica - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADZONAGEO
    
    objCADZONAGEO.CODIGO = iCodigo
    
    objBLBFunc.LimpaCampos frmCADZONAGEO
    objBLBFunc.Preenche_Estado cboESTADO
    
    ConfGridEstado
    ConfGridCliente
    ConfGridCidade
    
    If objCADZONAGEO.Carrega_campos = True Then
       
       txtCodigo.Text = objCADZONAGEO.CODIGO
       txtDescricao.Text = objCADZONAGEO.DESCRI
       
       arrESTADOS = objCADZONAGEO.ESTADOS
       If IsArray(arrESTADOS) Then
          For I = 1 To UBound(arrESTADOS)
              flxESTADOS.AddItem arrESTADOS(I) & vbTab & _
                                 CarregaEstado(arrESTADOS(I))
          Next I
       End If
    End If

End Sub

Private Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    Frame2.Enabled = True
    Frame3.Enabled = True
    
    Me.Caption = "Cadastro de Zona Geografica - [ ALTERAÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADZONAGEO
    
    objCADZONAGEO.CODIGO = iCodigo
    
    objBLBFunc.LimpaCampos frmCADZONAGEO
    objBLBFunc.Preenche_Estado cboESTADO
    
    ConfGridEstado
    ConfGridCliente
    ConfGridCidade
    
    If objCADZONAGEO.Carrega_campos = True Then
       
       txtCodigo.Text = objCADZONAGEO.CODIGO
       txtDescricao.Text = objCADZONAGEO.DESCRI
    
       arrESTADOS = objCADZONAGEO.ESTADOS
       If IsArray(arrESTADOS) Then
          For I = 1 To UBound(arrESTADOS)
              flxESTADOS.AddItem arrESTADOS(I) & vbTab & _
                                 CarregaEstado(arrESTADOS(I))
          Next I
       End If
    End If

End Sub

Private Sub ConfGridEstado()

    flxESTADOS.Rows = 1
    flxESTADOS.Cols = 2
    
    flxESTADOS.TextMatrix(0, 0) = ""
    flxESTADOS.TextMatrix(0, 1) = "Estado"

    flxESTADOS.ColWidth(0) = 0
    flxESTADOS.ColWidth(1) = 1000

End Sub

Private Sub ConfGridCliente()

    flxCLIENTES.Rows = 1
    flxCLIENTES.Cols = 4
    
    flxCLIENTES.TextMatrix(0, 0) = ""
    flxCLIENTES.TextMatrix(0, 1) = "Código"
    flxCLIENTES.TextMatrix(0, 2) = "Razão Social"
    flxCLIENTES.TextMatrix(0, 3) = "Cidade"

    flxCLIENTES.ColWidth(0) = 0
    flxCLIENTES.ColWidth(1) = 1000
    flxCLIENTES.ColWidth(2) = 5000
    flxCLIENTES.ColWidth(3) = 3000

End Sub

Public Sub IncGridEstado()

    Dim I As Integer

    If Len(Trim(cboESTADO.Text)) = 0 Then
       MsgBox "O Estado deve ser incluso !!", vbOKOnly + vbExclamation, "Aviso"
       cboESTADO.Text = ""
       cboESTADO.SetFocus
       Exit Sub
    End If
    
    For I = 1 To (flxESTADOS.Rows - 1)
        If cboESTADO.ItemData(cboESTADO.ListIndex) = flxESTADOS.TextMatrix(I, 0) Then
           MsgBox "Este Estado já foi informado !!!", vbOKOnly + vbExclamation, "Aviso"
           cboESTADO.ListIndex = -1
           cboESTADO.SetFocus
           Exit Sub
        End If
    Next I
        
    flxESTADOS.AddItem cboESTADO.ItemData(cboESTADO.ListIndex) & vbTab & _
                       Trim(cboESTADO.Text)
                       
                       
    cboESTADO.ListIndex = -1
    cboESTADO.SetFocus
    
End Sub

Private Function CarregaEstado(lngCODGIESTADO As Variant) As String
    
    Dim I As Integer
    
    CarregaEstado = ""
    
    For I = 0 To (cboESTADO.ListCount - 1)
        If cboESTADO.ItemData(I) = lngCODGIESTADO Then CarregaEstado = cboESTADO.List(I)
    Next I
    
End Function

Private Sub PopGridClientes(lngESTADO As Long)

    ConfGridCliente
    ConfGridCidade
    
    
    Frame5.Caption = ""
    If (flxESTADOS.Rows - 1) > 0 Then
        Frame5.Caption = "[ Estado : " & Trim(flxESTADOS.TextMatrix(flxESTADOS.Row, 1)) & " ]"
    End If
    
    '' ---------------------------------------------------------------

    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_ESTNORM = " & lngESTADO

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       flxCLIENTES.AddItem "" & vbTab & _
                           BREC!SGI_CODIGO & vbTab & _
                           Trim(BREC!SGI_RAZAOSOC) & vbTab & _
                           Trim(BREC!SGI_CIDNORM)
                           
       BREC.MoveNext
    Loop
    BREC.Close
    
    '' ---------------------------------------------------------------
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       Distinct SGI_CIDNORM " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_ESTNORM = " & lngESTADO & vbCrLf
    sSql = sSql & " Order by SGI_CIDNORM "

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       flxCidade.AddItem "" & vbTab & Trim(BREC!SGI_CIDNORM)
       BREC.MoveNext
    Loop
    BREC.Close
    
End Sub

Private Sub ConfGridCidade()

    flxCidade.Rows = 1
    flxCidade.Cols = 2
    
    flxCidade.TextMatrix(0, 0) = ""
    flxCidade.TextMatrix(0, 1) = "Cidade"

    flxCidade.ColWidth(0) = 0
    flxCidade.ColWidth(1) = 2500

End Sub

Private Sub PopGridClientesCidade(strCIDADE As String, lngESTADO As Long)

    ConfGridCliente
    
    Frame5.Caption = ""
    If (flxESTADOS.Rows - 1) > 0 Then
        Frame5.Caption = "[ Cidade : " & Trim(strCIDADE) & " ]"
    End If
    
    '' ---------------------------------------------------------------

    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADCLIENTE " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL  = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CIDNORM = '" & Trim(strCIDADE) & "'" & vbTab
    sSql = sSql & "   And SGI_ESTNORM = " & lngESTADO

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       flxCLIENTES.AddItem "" & vbTab & _
                           BREC!SGI_CODIGO & vbTab & _
                           Trim(BREC!SGI_RAZAOSOC) & vbTab & _
                           Trim(BREC!SGI_CIDNORM)
                           
       BREC.MoveNext
    Loop
    BREC.Close
    
End Sub


