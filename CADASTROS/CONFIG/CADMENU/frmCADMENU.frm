VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCADMENU 
   Caption         =   "Cadastro de Niveis de Menu"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7665
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   7665
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab sstMenu 
      Height          =   4215
      Left            =   0
      TabIndex        =   15
      Top             =   1080
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   7435
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
      TabCaption(0)   =   "Dados de Acesso"
      TabPicture(0)   =   "frmCADMENU.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Arvore de Menu"
      TabPicture(1)   =   "frmCADMENU.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame4 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   21
         Top             =   360
         Width           =   7335
         Begin VB.CommandButton cmd_Cancela 
            Caption         =   "Cancela"
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
            Left            =   1800
            TabIndex        =   22
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmdRecarrega 
            Caption         =   "Padrão"
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
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1575
         End
         Begin MSComctlLib.TreeView treMenu 
            Height          =   3015
            Left            =   120
            TabIndex        =   10
            Top             =   600
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   5318
            _Version        =   393217
            LabelEdit       =   1
            LineStyle       =   1
            PathSeparator   =   ","
            Style           =   7
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3735
         Left            =   120
         TabIndex        =   16
         Top             =   420
         Width           =   7335
         Begin VB.TextBox txtCodigo 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1800
            TabIndex        =   11
            Text            =   "txtCodigo"
            Top             =   240
            Width           =   975
         End
         Begin VB.ComboBox cboNivel 
            Height          =   315
            ItemData        =   "frmCADMENU.frx":0038
            Left            =   1800
            List            =   "frmCADMENU.frx":003A
            TabIndex        =   0
            Text            =   "cboNivel"
            Top             =   600
            Width           =   1815
         End
         Begin VB.ComboBox cboDepto 
            Height          =   315
            Left            =   2160
            TabIndex        =   1
            Text            =   "cboDepto"
            Top             =   960
            Width           =   4455
         End
         Begin VB.CommandButton cmdPesquisa 
            Height          =   315
            Left            =   1800
            Picture         =   "frmCADMENU.frx":003C
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   960
            Width           =   375
         End
         Begin VB.Frame Frame3 
            Caption         =   "Tipo de Acesso"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2295
            Left            =   120
            TabIndex        =   17
            Top             =   1320
            Width           =   7095
            Begin VB.OptionButton Option2 
               Caption         =   "Limpa Tipo de Acesso"
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
               Left            =   4080
               TabIndex        =   8
               Top             =   960
               Width           =   2295
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Acesso Total"
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
               Left            =   4080
               TabIndex        =   7
               Top             =   600
               Width           =   1455
            End
            Begin VB.ComboBox cboTipoAcesso 
               Height          =   315
               Left            =   240
               TabIndex        =   2
               Text            =   "cboTipoAcesso"
               Top             =   240
               Width           =   2055
            End
            Begin VB.CommandButton cmdGravaNiveis 
               Height          =   315
               Left            =   2280
               Picture         =   "frmCADMENU.frx":013E
               Style           =   1  'Graphical
               TabIndex        =   3
               Top             =   240
               Width           =   375
            End
            Begin MSFlexGridLib.MSFlexGrid flxAcesso 
               Height          =   1575
               Left            =   240
               TabIndex        =   6
               Top             =   600
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   2778
               _Version        =   393216
               FixedCols       =   0
               HighLight       =   2
               SelectionMode   =   1
               Appearance      =   0
            End
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código :"
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
            Left            =   840
            TabIndex        =   20
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nivel :"
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
            Left            =   975
            TabIndex        =   19
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Departamento :"
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
            Left            =   225
            TabIndex        =   18
            Top             =   960
            Width           =   1335
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   7575
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
         Picture         =   "frmCADMENU.frx":0240
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   735
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
         Left            =   1560
         Picture         =   "frmCADMENU.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   735
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
         Left            =   840
         Picture         =   "frmCADMENU.frx":0444
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmCADMENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCaminho    As String
Public Linha       As Variant
Public cTipOper    As String
Public iCodigo     As Integer
Public FILIAL      As Integer
Public strACESSO   As String
Dim objBLBFunc     As Object
Dim objMENU        As Object
Dim nodX           As Node
Dim Indice         As Integer
''Dim arrMenu        As Variant
Dim arrTipAces     As Variant
Dim objPESQPADRAO  As Object

Private Sub cboDepto_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboDepto, KeyAscii
End Sub

Private Sub cboNivel_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboNivel, KeyAscii
End Sub

Private Sub cboTipoAcesso_KeyPress(KeyAscii As Integer)
    objBLBFunc.ComboMagico cboTipoAcesso, KeyAscii
End Sub

Private Sub cmd_Cancela_Click()
    CarregaMenuAlteracao
End Sub

Private Sub cmdAltera_Click()

    If objBLBFunc.ChecaAcesso2("A", strACESSO) = False Then Exit Sub
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    cmd_Cancela.Visible = True
    
    Frame2.Enabled = True
    Frame4.Enabled = True
   
    Me.Caption = "Cadastro de Niveis de Menu - [ ALTERÇÃO ]"

End Sub

Private Sub cmdGravaNiveis_Click()
    
    Dim I As Integer
    
    If Len(Trim(cboTipoAcesso.Text)) = 0 Then Exit Sub
    
    For I = 0 To (flxAcesso.Rows - 1)
        If Trim(flxAcesso.TextMatrix(I, 1)) = Trim(cboTipoAcesso.Text) Then Exit Sub
    Next I
    
    flxAcesso.AddItem Mid(cboTipoAcesso.Text, 1, 1) & vbTab & cboTipoAcesso.Text
    
End Sub

Private Sub cmdPesquisa_Click()

    Dim I As Integer
    ReDim arrCAMPOS(1 To 2, 1 To 5) As String
    ReDim arrTABELA(1 To 1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADDEPTO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    arrTABELA(1) = sSql
    
    arrCAMPOS(1, 1) = "SGI_CODDEPTO"
    arrCAMPOS(1, 2) = "N"
    arrCAMPOS(1, 3) = "Código"
    arrCAMPOS(1, 4) = "700"
    arrCAMPOS(1, 5) = "SGI_CODDEPTO"
    
    arrCAMPOS(2, 1) = "SGI_DESCRICAO"
    arrCAMPOS(2, 2) = "S"
    arrCAMPOS(2, 3) = "Descrição"
    arrCAMPOS(2, 4) = "3000"
    arrCAMPOS(2, 5) = "SGI_DESCRICAO"
    
    varRETORNO = objPESQPADRAO.cConnect(cCaminho, Linha, FILIAL, strACESSO, V_Usuario, arrCAMPOS, arrTABELA, "Departamentos")
    
    cboDepto.ListIndex = -1
    
    If Len(Trim(varRETORNO)) > 0 Then
       For I = 0 To (cboDepto.ListCount - 1)
           If cboDepto.ItemData(I) = varRETORNO Then cboDepto.ListIndex = I
       Next I
    End If
    
    cboDepto.SetFocus

End Sub

Private Sub cmdRecarrega_Click()
    CarregaMenu
End Sub

Private Sub CmdSalva_Click()
    
    Dim arrGrid      As Variant
    Dim I            As Integer
    Dim j            As Integer
    Dim strCHAVE     As String
    Dim strMODULOS   As String
    Dim arrMODULOS() As String
    Dim arrCAMPOS()  As String
    
    If Valida_Campos = True Then
    
       
       If cTipOper = "I" Then objMENU.MENCODIGO = objMENU.Gera_Codigo(Me.Name)
       
       objMENU.MENNIVEL = cboNivel.ItemData(cboNivel.ListIndex)
       
       '' Montando os Acessos
       ReDim arrGrid(1 To (flxAcesso.Rows - 1)) As String
       For I = 1 To UBound(arrGrid)
           arrGrid(I) = flxAcesso.TextMatrix(I, 1)
       Next I
       objMENU.MENTIPACS = arrGrid
       
       '' Pegando Os itens que ten ainda no grid
       '' FAzendo Somente do Root
       
       Dim intREGSROOT      As Integer
       Dim intREGSTIPOS     As Integer
       Dim intREGSTIPOM     As Integer
       
       intREGSROOT = 0
       intREGSTIPOS = 0
       intREGSTIPOM = 0
       For I = 1 To treMenu.Nodes.Count
           
           If intREGSTIPOS > 0 Then
           
                sSql = "Select " & vbCrLf
                sSql = sSql & "       * " & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_MENUP " & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       SGI_FILIAL = 0" & vbCrLf
                sSql = sSql & "   And SGI_TIPO   = 'M'" & vbCrLf
                sSql = sSql & "   And SGI_CIGLA  = '" & Trim(arrMenu(intREGSROOT).arrNIVEL_S(intREGSTIPOS).strCIGLA2) & "'" & vbCrLf
                sSql = sSql & "   And SGI_CIGLA2 = '" & Trim(treMenu.Nodes.Item(I).Key) & "'"
                
                BREC.Open sSql, adoBanco_Dados, adOpenDynamic
                If Not BREC.EOF() Then
                   
                   intREGSTIPOM = (intREGSTIPOM + 1)
                   ReDim Preserve arrMenu(intREGSROOT).arrNIVEL_S(intREGSTIPOS).arrNIVEL_M(1 To intREGSTIPOM) As Menu_Niveis_TipoM
                   
                   arrMenu(intREGSROOT).arrNIVEL_S(intREGSTIPOS).arrNIVEL_M(intREGSTIPOM).intCODIGO = BREC!SGI_CODIGO
                   arrMenu(intREGSROOT).arrNIVEL_S(intREGSTIPOS).arrNIVEL_M(intREGSTIPOM).strTEXTO = BREC!SGI_TEXTO
                   arrMenu(intREGSROOT).arrNIVEL_S(intREGSTIPOS).arrNIVEL_M(intREGSTIPOM).strTIPO = BREC!SGI_TIPO
                   arrMenu(intREGSROOT).arrNIVEL_S(intREGSTIPOS).arrNIVEL_M(intREGSTIPOM).strCIGLA = BREC!SGI_CIGLA
                   arrMenu(intREGSROOT).arrNIVEL_S(intREGSTIPOS).arrNIVEL_M(intREGSTIPOM).strCIGLA2 = IIf(IsNull(BREC!SGI_CIGLA2), "", BREC!SGI_CIGLA2)
                   arrMenu(intREGSROOT).arrNIVEL_S(intREGSTIPOS).arrNIVEL_M(intREGSTIPOM).strModulo = IIf(IsNull(BREC!SGI_MODULO), "", BREC!SGI_MODULO)
                   arrMenu(intREGSROOT).arrNIVEL_S(intREGSTIPOS).intQTDNIVEL = intREGSTIPOM
                   
                End If
                BREC.Close
           
           End If
           
           If intREGSROOT > 0 Then
           
                sSql = "Select " & vbCrLf
                sSql = sSql & "       * " & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_MENUP " & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       SGI_FILIAL = 0" & vbCrLf
                sSql = sSql & "   And SGI_TIPO   = 'S'" & vbCrLf
                sSql = sSql & "   And SGI_CIGLA  = '" & Trim(arrMenu(intREGSROOT).strCIGLA) & "'" & vbCrLf
                sSql = sSql & "   And SGI_CIGLA2 = '" & Trim(treMenu.Nodes.Item(I).Key) & "'"
                
                BREC.Open sSql, adoBanco_Dados, adOpenDynamic
                If Not BREC.EOF Then
                      
                      intREGSTIPOM = 0
                      intREGSTIPOS = (intREGSTIPOS + 1)
                      ReDim Preserve arrMenu(intREGSROOT).arrNIVEL_S(1 To intREGSTIPOS) As Menu_Niveis_TipoS
                   
                      arrMenu(intREGSROOT).arrNIVEL_S(intREGSTIPOS).intCODIGO = BREC!SGI_CODIGO
                      arrMenu(intREGSROOT).arrNIVEL_S(intREGSTIPOS).strTEXTO = BREC!SGI_TEXTO
                      arrMenu(intREGSROOT).arrNIVEL_S(intREGSTIPOS).strTIPO = BREC!SGI_TIPO
                      arrMenu(intREGSROOT).arrNIVEL_S(intREGSTIPOS).strCIGLA = BREC!SGI_CIGLA
                      arrMenu(intREGSROOT).arrNIVEL_S(intREGSTIPOS).strCIGLA2 = IIf(IsNull(BREC!SGI_CIGLA2), "", BREC!SGI_CIGLA2)
                      arrMenu(intREGSROOT).arrNIVEL_S(intREGSTIPOS).strModulo = IIf(IsNull(BREC!SGI_MODULO), "", BREC!SGI_MODULO)
                      arrMenu(intREGSROOT).intQTDNIVEL = intREGSTIPOS
                
                End If
                BREC.Close
           
           End If
           
           sSql = "Select " & vbCrLf
           sSql = sSql & "       * " & vbCrLf
           sSql = sSql & "  From " & vbCrLf
           sSql = sSql & "       SGI_MENUP " & vbCrLf
           sSql = sSql & " Where " & vbCrLf
           sSql = sSql & "       SGI_FILIAL = 0" & vbCrLf
           sSql = sSql & "   And SGI_TIPO   = 'P' " & vbCrLf
           sSql = sSql & "   And SGI_CIGLA  = '" & Trim(treMenu.Nodes.Item(I).Key) & "'"
           
           BREC.Open sSql, adoBanco_Dados, adOpenDynamic
           If Not BREC.EOF() Then
              
              intREGSTIPOS = 0
              intREGSTIPOM = 0
              intREGSROOT = (intREGSROOT + 1)
              ReDim Preserve arrMenu(1 To intREGSROOT) As Menu_Niveis
              
              arrMenu(intREGSROOT).intCODIGO = BREC!SGI_CODIGO
              arrMenu(intREGSROOT).strTEXTO = BREC!SGI_TEXTO
              arrMenu(intREGSROOT).strTIPO = BREC!SGI_TIPO
              arrMenu(intREGSROOT).strCIGLA = BREC!SGI_CIGLA
              arrMenu(intREGSROOT).strCIGLA2 = IIf(IsNull(BREC!SGI_CIGLA2), "", BREC!SGI_CIGLA2)
              arrMenu(intREGSROOT).strModulo = IIf(IsNull(BREC!SGI_MODULO), "", BREC!SGI_MODULO)
              arrMenu(intREGSROOT).intQTDNIVEL = intREGSTIPOS
                            
           End If
           BREC.Close
           
                       
       Next I
       
       '' ---------
       objMENU.MENDESNIV = cboNivel.Text
       objMENU.MENDEPTO = cboDepto.ItemData(cboDepto.ListIndex)
        
       If objMENU.GRAVA(cTipOper) = True Then
          
          MsgBox "O nivel de acesso foi " & IIf(cTipOper = "I", "incluso", IIf(cTipOper = "A", "Alterado", "")) & " com sucesso !!!", vbOKOnly + vbInformation, "Aviso"
          
          If cTipOper = "I" Then
             Set objBLBFunc = Nothing
             Set objMENU = Nothing
             Unload Me
          End If
          
          If cTipOper = "A" Then sstMenu.Tab = 0
          
       End If
    
    End If
    
End Sub

Private Sub cmdVoltar_Click()
    Set objBLBFunc = Nothing
    Set objMENU = Nothing
    Unload Me
End Sub

Private Sub flxAcesso_KeyDown(KeyCode As Integer, Shift As Integer)

 If KeyCode = vbKeyDelete Then
    If flxAcesso.Rows > 2 Then
       flxAcesso.RemoveItem flxAcesso.RowSel
    ElseIf flxAcesso.Rows = 2 Then
       flxAcesso.Rows = 1
    End If
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
   Set objMENU = CreateObject("CADMENU.clsCADMENU")
   Set objPESQPADRAO = CreateObject("PESQPADRAO.clsPESQPADRAO")
   
   objMENU.FILIAL = FILIAL
   
   If cTipOper = "I" Then Inclui
   If cTipOper = "A" Then Altera
   If cTipOper = "C" Then Consulta
   
   objBLBFunc.ChecaAcesso frmCADMENU, strACESSO

End Sub

Private Sub Inclui()

    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    cmd_Cancela.Visible = False
    sstMenu.Tab = 0
    
    Frame2.Enabled = True
    Frame4.Enabled = True
   
    Me.Caption = "Cadastro de Niveis de Menu - [ INCLUSÃO ]"
    
    objBLBFunc.LimpaCampos frmCADMENU
    
    txtCodigo.Text = ""

    ConfGrid
    
    '' ----------------------------------------------------------------------
    '' Inclui nivel de acesso
    cboNivel.Clear
    cboNivel.AddItem "DIRETORIA"
    cboNivel.ItemData(cboNivel.NewIndex) = 1
    
    cboNivel.AddItem "GERENCIA"
    cboNivel.ItemData(cboNivel.NewIndex) = 2
    
    cboNivel.AddItem "USUÁRIO"
    cboNivel.ItemData(cboNivel.NewIndex) = 3
    '' ----------------------------------------------------------------------
    
    Call CarregaComboAcesso
    Call CarregaMenu
    
    PreencheComboDepto
    
    cboNivel.TabIndex = 0
    
End Sub

Private Sub ConfGrid()

    flxAcesso.Rows = 1
    flxAcesso.Cols = 2
    
    flxAcesso.TextMatrix(0, 1) = "Tipo Acesso"
    
    flxAcesso.ColWidth(0) = 0
    flxAcesso.ColWidth(1) = 2000
    
End Sub

Private Sub Option1_Click()

    Dim I As Integer
    
    flxAcesso.Rows = 1
    
    For I = 0 To (cboTipoAcesso.ListCount - 1)
        flxAcesso.AddItem Mid(cboTipoAcesso.List(I), 1, 1) & vbTab & cboTipoAcesso.List(I)
    Next I

End Sub

Private Sub Option2_Click()
    flxAcesso.Rows = 1
End Sub

Private Function Valida_Campos() As Boolean

    Valida_Campos = False
    
    If cboNivel.ListIndex = -1 Then
       MsgBox "Informe o nivel de acesso !!!", vbOKOnly + vbCritical, "Aviso"
       cboNivel.SetFocus
       Exit Function
    End If
    
    If cboDepto.ListIndex = -1 Then
       MsgBox "Informe o departamento !!!", vbOKOnly + vbCritical, "Aviso"
       cboDepto.SetFocus
       Exit Function
    End If
    
    If flxAcesso.Rows = 1 Then
       MsgBox "Informe o tipo de Acesso !!!", vbOKOnly + vbCritical, "Aviso"
       cboTipoAcesso.SetFocus
       Exit Function
    End If
    
    If treMenu.Nodes.Count = 0 Then
       MsgBox "Informe a arvore de menu !!!", vbOKOnly + vbCritical, "Aviso"
       sstMenu.Tab = 1
       Exit Function
    End If
    
    Valida_Campos = True
    
End Function

Private Sub CarregaMenu()

   Dim iTotMenu As Integer
   Dim iTotReg  As Integer
  
   ' Set Treeview control properties.
   treMenu.Nodes.Clear
   treMenu.LineStyle = tvwRootLines  ' Linestyle 1
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_MENUP " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = 0 " & vbCrLf
   sSql = sSql & "   And SGI_CODGER = 0" & vbCrLf  '' & objMENU.MENCODIGO & vbCrLf
   sSql = sSql & "   And SGI_TIPO = 'P' " & vbCrLf
   sSql = sSql & " Order by SGI_CODIGO"
   
   BREC.Open sSql, adoBanco_Dados, adOpenDynamic

   If Not BREC.EOF() Then
   
      '' -------------------------------------------------------------
      iTotReg = 1
      Do While Not BREC.EOF
         iTotReg = iTotReg + 1
         BREC.MoveNext
      Loop
      '' -------------------------------------------------------------
      
      iTotReg = 1
      BREC.MoveFirst
      Do While Not BREC.EOF
         
         If UCase(BREC!SGI_TIPO) = "P" Then
     
            Set nodX = treMenu.Nodes.Add(, , Trim(BREC!SGI_CIGLA), Trim(BREC!SGI_TEXTO))
            treMenu.Nodes.Item(treMenu.Nodes.Count).Checked = True
            
            If Trim(BREC!SGI_TEXTO) = "Configurações" Then
            
                sSql = "Select " & vbCrLf
                sSql = sSql & "       *" & vbCrLf
                sSql = sSql & "  From " & vbCrLf
                sSql = sSql & "       SGI_MENUP " & vbCrLf
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       SGI_FILIAL = 0 " & vbCrLf
                sSql = sSql & "   And SGI_TIPO   = 'M'" & vbCrLf
                sSql = sSql & "   And SGI_CIGLA  = '" & Trim(BREC!SGI_CIGLA) & "'" & vbCrLf
                sSql = sSql & " Order by SGI_CODIGO"
                
                BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
                Do While Not BREC3.EOF()
                
                   Set nodX = treMenu.Nodes.Add(Trim(BREC3!SGI_CIGLA), 4, Trim(BREC3!SGI_CIGLA2), Trim(BREC3!SGI_TEXTO))
                   treMenu.Nodes.Item(treMenu.Nodes.Count).Checked = True
                   
                   BREC3.MoveNext
                Loop
                BREC3.Close
            
            Else
            
               sSql = " Select " & vbCrLf
               sSql = sSql & "        * " & vbCrLf
               sSql = sSql & "   From " & vbCrLf
               sSql = sSql & "        SGI_MENUP " & vbCrLf
               sSql = sSql & "  Where " & vbCrLf
               sSql = sSql & "        SGI_FILIAL = 0 " & vbCrLf
               sSql = sSql & "    And SGI_TIPO   = 'S' " & vbCrLf
               sSql = sSql & "    And SGI_CIGLA  = '" & Trim(BREC!SGI_CIGLA) & "'" & vbCrLf
               sSql = sSql & "Order by SGI_CODIGO "
            
               BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
               Do While Not BREC2.EOF()
                   
                   Set nodX = treMenu.Nodes.Add(Trim(BREC2!SGI_CIGLA), 4, Trim(BREC2!SGI_CIGLA2), Trim(BREC2!SGI_TEXTO))
                   treMenu.Nodes.Item(treMenu.Nodes.Count).Checked = True
                   
                   sSql = "Select " & vbCrLf
                   sSql = sSql & "       *" & vbCrLf
                   sSql = sSql & "  From " & vbCrLf
                   sSql = sSql & "       SGI_MENUP " & vbCrLf
                   sSql = sSql & " Where " & vbCrLf
                   sSql = sSql & "       SGI_FILIAL = 0 " & vbCrLf
                   sSql = sSql & "   And SGI_TIPO   = 'M'" & vbCrLf
                   sSql = sSql & "   And SGI_CIGLA  = '" & Trim(BREC2!SGI_CIGLA2) & "'" & vbCrLf
                   sSql = sSql & " Order by SGI_CODIGO"
                   
                   BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
                   Do While Not BREC3.EOF()
                   
                      Set nodX = treMenu.Nodes.Add(Trim(BREC3!SGI_CIGLA), 4, Trim(BREC3!SGI_CIGLA2), Trim(BREC3!SGI_TEXTO))
                      treMenu.Nodes.Item(treMenu.Nodes.Count).Checked = True
                      
                      BREC3.MoveNext
                   Loop
                   BREC3.Close
                   
               
                   BREC2.MoveNext
               Loop
               BREC2.Close
         
            End If
            
         End If
       
         iTotReg = iTotReg + 1
         BREC.MoveNext
         
      Loop
      
   End If
   BREC.Close

End Sub
Private Sub treMenu_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'Dim I           As Integer
    'Dim j           As Integer
    'Dim L           As Integer
    
    'Dim strPAI        As String
    'Dim strATUAL      As String
    'Dim intFilhos     As Integer
    
    'Dim sConfiguracao As String
    'Dim sTexto        As String
    'Dim sFullPath     As String
    'Dim sFullPath2    As String
    'Dim sChld         As String
    'Dim sChave        As String
    'Dim sChave2       As String
    'Dim sChave3       As String
    'Dim arrCaminho    As Variant
    
    'If KeyCode = vbKeyDelete Then
       
    '   arrCaminho = Split(treMenu.Nodes.Item(Indice).FullPath, ",") '' -- Array do Caminho
    '   sConfiguracao = treMenu.Nodes.Item(Indice).FullPath
    '   strATUAL = treMenu.Nodes.Item(Indice).Text
    '
    '   For I = 0 To UBound(arrCaminho)
    '       For j = 1 To UBound(arrMenu)
    '           If (arrMenu(j, 2) = "P") And (arrMenu(j, 1) = arrCaminho(I)) Then
    '              sChave = arrMenu(j, 3)
    '           ElseIf (arrMenu(j, 2) = "S") And (arrMenu(j, 1) = arrCaminho(I)) Then
    '              sChave = arrMenu(j, 3)
    '           ElseIf (arrMenu(j, 2) = "M") And (arrMenu(j, 1) = arrCaminho(I)) Then
    '              sChave = arrMenu(j, 4)
    '           End If
    '       Next j
    '   Next I

    '   For I = 1 To UBound(arrMenu)
    '       If (arrMenu(I, 2) = "P") And (arrMenu(I, 3) = sChave) Then
    '          arrMenu(I, 1) = ""
    '          arrMenu(I, 2) = ""
    '          arrMenu(I, 3) = ""
    '          arrMenu(I, 4) = ""
    '          arrMenu(I, 5) = ""
    '          arrMenu(I, 6) = ""
    '          For j = 1 To UBound(arrMenu)
    '              If (arrMenu(j, 2) = "S") And (arrMenu(j, 3) = sChave) Then
    '                 sChave2 = arrMenu(j, 4)
    '                 arrMenu(j, 1) = ""
    '                 arrMenu(j, 2) = ""
    '                 arrMenu(j, 3) = ""
    '                 arrMenu(j, 4) = ""
    '                 arrMenu(j, 5) = ""
    '                 arrMenu(j, 6) = ""
    '                 For L = 1 To UBound(arrMenu)
    '                     If (arrMenu(L, 2) = "M") And (arrMenu(L, 3) = sChave2) Then
    '                        sChave3 = arrMenu(L, 4)
    '                        arrMenu(L, 1) = ""
    '                        arrMenu(L, 2) = ""
    '                        arrMenu(L, 3) = ""
    '                        arrMenu(L, 4) = ""
    '                        arrMenu(L, 5) = ""
    '                        arrMenu(L, 6) = ""
    '                     End If
    '                 Next L
    '              End If
    '          Next j
    '       End If
    '   Next I
      
    '   For I = 1 To UBound(arrMenu)
    '       If (arrMenu(I, 2) = "S") And (arrMenu(I, 3) = sChave) Then
    '           sChave2 = arrMenu(I, 4)
    '           arrMenu(I, 1) = ""
    '           arrMenu(I, 2) = ""
    '           arrMenu(I, 3) = ""
    '           arrMenu(I, 4) = ""
    '           arrMenu(I, 5) = ""
    '           arrMenu(I, 6) = ""
    '           For j = 1 To UBound(arrMenu)
    '               If (arrMenu(j, 2) = "M") And (arrMenu(j, 3) = sChave2) Then
    '                  sChave3 = arrMenu(j, 4)
    '                  arrMenu(j, 1) = ""
    '                  arrMenu(j, 2) = ""
    '                  arrMenu(j, 3) = ""
    '                  arrMenu(j, 4) = ""
    '                  arrMenu(j, 5) = ""
    '                  arrMenu(j, 6) = ""
    '                End If
    '           Next j
    '       End If
    '   Next I

    '   For I = 1 To UBound(arrMenu)
    '       If (arrMenu(I, 2) = "M") And (arrMenu(I, 4) = sChave) Then
    '           arrMenu(I, 1) = ""
    '           arrMenu(I, 2) = ""
    '           arrMenu(I, 3) = ""
    '           arrMenu(I, 4) = ""
    '           arrMenu(I, 5) = ""
    '           arrMenu(I, 6) = ""
    '       End If
    '   Next I
    '
    '   '' Para apagar opções de configuração
    '   If sConfiguracao = "Configurações" Then
    '      For I = 1 To UBound(arrMenu)
    '          If (arrMenu(I, 2) = "M") And (arrMenu(I, 3) = sChave) Then
    '              arrMenu(I, 1) = ""
    '              arrMenu(I, 2) = ""
    '              arrMenu(I, 3) = ""
    '              arrMenu(I, 4) = ""
    '              arrMenu(I, 5) = ""
    '              arrMenu(I, 6) = ""
    '          End If
    '      Next I
    '   End If
    '
       If Indice > 0 Then treMenu.Nodes.Remove Indice
       If treMenu.Nodes.Count = 0 Then Indice = -1
    '
    'End If
   
End Sub

Private Sub treMenu_NodeClick(ByVal Node As MSComctlLib.Node)
    Indice = Node.Index
End Sub

Private Sub Consulta()

    Dim I As Integer
    
    CmdSalva.Enabled = False
    cmdAltera.Enabled = True
    cmd_Cancela.Visible = False
    
    Frame2.Enabled = False
    Frame4.Enabled = False
   
    Me.Caption = "Cadastro de Niveis de Menu - [ CONSULTA ]"
    
    objBLBFunc.LimpaCampos frmCADMENU
    
    ConfGrid
    
    objMENU.MENCODIGO = iCodigo
    
    '' ----------------------------------------------------------------------
    '' Inclui nivel de acesso
    cboNivel.Clear
    cboNivel.AddItem "DIRETORIA"
    cboNivel.ItemData(cboNivel.NewIndex) = 1
    
    cboNivel.AddItem "GERENCIA"
    cboNivel.ItemData(cboNivel.NewIndex) = 2
    
    cboNivel.AddItem "USUÁRIO"
    cboNivel.ItemData(cboNivel.NewIndex) = 3
    '' ----------------------------------------------------------------------
    
    Call CarregaComboAcesso
    
    If objMENU.Carrega_campos = True Then

       txtCodigo.Text = Str(objMENU.MENCODIGO)
       
       For I = 0 To (cboNivel.ListCount - 1)
           If cboNivel.ItemData(I) = objMENU.MENNIVEL Then cboNivel.ListIndex = I
       Next I
       
       If IsArray(objMENU.MENTIPACS) = True Then
          arrTipAces = objMENU.MENTIPACS
          For I = 1 To UBound(arrTipAces)
              flxAcesso.AddItem "" & vbTab & arrTipAces(I)
          Next I
       End If
       
       CarregaMenuAlteracao
       
       PreencheComboDepto
       For I = 0 To (cboDepto.ListCount - 1)
           If cboDepto.ItemData(I) = objMENU.MENDEPTO Then cboDepto.ListIndex = I
       Next I
       
    End If
    
    sstMenu.Tab = 0

End Sub

Private Sub Altera()

    Dim I As Integer
    
    CmdSalva.Enabled = True
    cmdAltera.Enabled = False
    cmd_Cancela.Visible = True
    
    Frame2.Enabled = True
    Frame4.Enabled = True
   
    Me.Caption = "Cadastro de Niveis de Menu - [ ALTERÇÃO ]"
    
    objBLBFunc.LimpaCampos frmCADMENU
    
    ConfGrid
    
    objMENU.MENCODIGO = iCodigo
    
    '' ----------------------------------------------------------------------
    '' Inclui nivel de acesso
    cboNivel.Clear
    cboNivel.AddItem "DIRETORIA"
    cboNivel.ItemData(cboNivel.NewIndex) = 1
    
    cboNivel.AddItem "GERENCIA"
    cboNivel.ItemData(cboNivel.NewIndex) = 2
    
    cboNivel.AddItem "USUÁRIO"
    cboNivel.ItemData(cboNivel.NewIndex) = 3
    '' ----------------------------------------------------------------------
    
    Call CarregaComboAcesso
       
    If objMENU.Carrega_campos = True Then

       txtCodigo.Text = Str(objMENU.MENCODIGO)
       
       For I = 0 To (cboNivel.ListCount - 1)
           If cboNivel.ItemData(I) = objMENU.MENNIVEL Then cboNivel.ListIndex = I
       Next I
       
       If IsArray(objMENU.MENTIPACS) = True Then
          arrTipAces = objMENU.MENTIPACS
          For I = 1 To UBound(arrTipAces)
              flxAcesso.AddItem "" & vbTab & arrTipAces(I)
          Next I
       End If
       
       CarregaMenuAlteracao
       
       PreencheComboDepto
       For I = 0 To (cboDepto.ListCount - 1)
           If cboDepto.ItemData(I) = objMENU.MENDEPTO Then cboDepto.ListIndex = I
       Next I
       
    End If
    
    sstMenu.Tab = 0

End Sub


Private Sub CarregaMenuAlteracao()

   Dim iTotMenu As Integer
   Dim iTotReg  As Integer
  
   ' Set Treeview control properties.
   treMenu.Nodes.Clear
   treMenu.LineStyle = tvwRootLines  ' Linestyle 1
   
   sSql = "Select " & vbTab
   sSql = sSql & "       * " & vbTab
   sSql = sSql & "  From " & vbTab
   sSql = sSql & "       SGI_MENUP " & vbTab
   sSql = sSql & " Where " & vbTab
   sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbTab
   sSql = sSql & "   And SGI_CODGER = " & iCodigo & vbTab
   sSql = sSql & " Order by SGI_CODIGO"
   
   BREC.Open sSql, adoBanco_Dados, adOpenDynamic

   If Not BREC.EOF Then
   
      '' -------------------------------------------------------------
      iTotReg = 1
      Do While Not BREC.EOF
         iTotReg = iTotReg + 1
         BREC.MoveNext
      Loop
      '' -------------------------------------------------------------
      
      ''ReDim Preserve arrMenu(1 To iTotReg) As Menu_Niveis
      
      iTotReg = 1
      BREC.MoveFirst
      Do While Not BREC.EOF
         
         If UCase(BREC!SGI_TIPO) = "P" Then
            Set nodX = treMenu.Nodes.Add(, , Trim(BREC!SGI_CIGLA), Trim(BREC!SGI_TEXTO))
         ElseIf UCase(BREC!SGI_TIPO) = "S" Then
            Set nodX = treMenu.Nodes.Add(Trim(BREC!SGI_CIGLA), 4, Trim(BREC!SGI_CIGLA2), Trim(BREC!SGI_TEXTO))
         ElseIf UCase(BREC!SGI_TIPO) = "M" Then
            Set nodX = treMenu.Nodes.Add(Trim(BREC!SGI_CIGLA), 4, Trim(BREC!SGI_CIGLA2) + Trim(Str(iTotMenu)), Trim(BREC!SGI_TEXTO))
         End If
         
         iTotReg = iTotReg + 1
         BREC.MoveNext
      Loop
      
   End If
   BREC.Close

End Sub


Private Sub PreencheComboDepto()

    cboDepto.Clear
    
    sSql = "Select * from SGI_CADDEPTO Where SGI_FILIAL = " & FILIAL
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       cboDepto.AddItem Format(BREC!SGI_CODDEPTO, "00") & " - " & BREC!SGI_DESCRICAO
       cboDepto.ItemData(cboDepto.NewIndex) = BREC!SGI_CODDEPTO
       BREC.MoveNext
    Loop
    
    BREC.Close
    
End Sub


Private Sub CarregaComboAcesso()

    '' ----------------------------------------------------------------------
    '' Inclui Tipo de Acesso
    cboTipoAcesso.Clear
    cboTipoAcesso.AddItem "I-INCLUSÃO"
    cboTipoAcesso.ItemData(cboTipoAcesso.NewIndex) = 1
    
    cboTipoAcesso.AddItem "A-ALTERAÇÃO"
    cboTipoAcesso.ItemData(cboTipoAcesso.NewIndex) = 2
    
    cboTipoAcesso.AddItem "E-EXCLUSÃO"
    cboTipoAcesso.ItemData(cboTipoAcesso.NewIndex) = 3
    
    cboTipoAcesso.AddItem "C-CONSULTA"
    cboTipoAcesso.ItemData(cboTipoAcesso.NewIndex) = 4
    
    cboTipoAcesso.AddItem "R-RELATÓRIO"
    cboTipoAcesso.ItemData(cboTipoAcesso.NewIndex) = 5
    
    cboTipoAcesso.AddItem "P-IMPRESSÃO"
    cboTipoAcesso.ItemData(cboTipoAcesso.NewIndex) = 6
    
    cboTipoAcesso.AddItem "L-LIBERA"
    cboTipoAcesso.ItemData(cboTipoAcesso.NewIndex) = 7
    
    cboTipoAcesso.AddItem "B-BLOQUEIA"
    cboTipoAcesso.ItemData(cboTipoAcesso.NewIndex) = 8
    
    cboTipoAcesso.AddItem "V-REPROVA"
    cboTipoAcesso.ItemData(cboTipoAcesso.NewIndex) = 9
    '' ----------------------------------------------------------------------

End Sub
