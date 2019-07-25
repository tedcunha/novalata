VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSIGE 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SIGE - Sistema Integrado de Gestão Empresarial"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   17055
   ForeColor       =   &H8000000D&
   Icon            =   "frmSIGE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   579
   ScaleMode       =   0  'User
   ScaleWidth      =   1.33328e6
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar TbrMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   17055
      _ExtentX        =   30083
      _ExtentY        =   741
      Appearance      =   1
      _Version        =   327682
   End
   Begin ComctlLib.StatusBar stbMensagem 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   8310
      Width           =   17055
      _ExtentX        =   30083
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   22331
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            AutoSize        =   2
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrHora 
      Interval        =   1000
      Left            =   120
      Top             =   7800
   End
   Begin VSFlex8LCtl.VSFlexGrid grdMENU 
      Height          =   8895
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   11655
      _cx             =   20558
      _cy             =   15690
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
   Begin ComctlLib.ImageList ImageList1 
      Left            =   11880
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Image imgSair 
      Height          =   360
      Left            =   12000
      Picture         =   "frmSIGE.frx":08CA
      Top             =   2280
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgItem 
      Height          =   225
      Left            =   12240
      Picture         =   "frmSIGE.frx":194C
      Top             =   1920
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgFolder 
      Height          =   225
      Left            =   12000
      Picture         =   "frmSIGE.frx":1A46
      Top             =   1920
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Menu mnFinanceiro 
      Caption         =   "Financeiro"
   End
   Begin VB.Menu mnComercial 
      Caption         =   "Comercial"
   End
   Begin VB.Menu mnExpedicao 
      Caption         =   "Expedição"
   End
   Begin VB.Menu mnEstoque 
      Caption         =   "Estoque"
   End
   Begin VB.Menu mnSuprimentos 
      Caption         =   ""
   End
   Begin VB.Menu mnControleDoc 
      Caption         =   ""
   End
   Begin VB.Menu mnQualidade 
      Caption         =   ""
   End
   Begin VB.Menu mnPCP 
      Caption         =   ""
   End
   Begin VB.Menu mnMetrologia 
      Caption         =   ""
   End
   Begin VB.Menu mnConfiguracao 
      Caption         =   ""
   End
   Begin VB.Menu mnSair 
      Caption         =   "&Sair"
   End
End
Attribute VB_Name = "frmSIGE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Linha        As Variant
Public strLINHA     As String
Public strUSUARIO   As String
Public iAcesso      As Integer
Public strEmpresa   As String
Public intFilial    As Integer
Public intNOVO      As Integer
Public lngCODACESSO As Long


Dim nodX            As Node
Dim intregs         As Integer
Dim IndMenu         As Integer
Dim StrImage        As String
Dim strACESSO       As String
Dim boolTEMMENU     As Boolean
Dim strVERSAO       As String
Dim boolChamaTela   As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
       If Sair = True Then Call FechaCampo
    End If
''    If KeyCode = vbKeyF2 Then Logon
End Sub

Private Sub Form_Load()
  
  '' --------------------
  V_Usuario = strUSUARIO
  
  Dim intLarg As Double
  Dim intAlt  As Double
  
  IndMenu = 1
  boolTEMMENU = True
  Me.Caption = CabecForm("CWS - " & strEmpresa & " - Versão : " & App.Major & "." & App.Minor & "." & App.Revision)
   
  strVERSAO = App.Major & "." & App.Minor & "." & App.Revision
   
  stbMensagem.Panels(1).Text = CabecForm("Menssagem : <F2> Troca Usuário / <F3> Pedidos Novalata / <F4> Pedidos Steel ")
  stbMensagem.Panels(2).Text = CabecForm("Usuário : ") & strUSUARIO
  
  '' --------------------
  '' Acesso ao Menu
  Call AbBanco(strNOVACONECT)
  
  sSql = "Select * From SGI_MENUP"
  
  BREC.Open sSql, BD, adOpenDynamic
  If BREC.EOF() Then boolTEMMENU = False
  BREC.Close
  Call FcBanco
  
  If boolTEMMENU = False Then
     MsgBox "ATENÇÃO - Menu de Usuário Inexistente !!!", vbOKOnly + vbExclamation, "Aviso"
     End
  End If
 
  Call ConfGridMenu
  Call CarregaMenuGride
  
  ''Call CriaToolBar(ImageList1, TbrMenu)
  strNOMMAQUINA = Trim(UCase(GetIPHostName()))

End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If strUSUARIO = "SGI" And Button = 2 Then ChamaConfigBanco
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Sair Then
        Call FechaCampo
    Else
        Cancel = True
    End If
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If strUSUARIO = "CWS" And Button = 2 Then ChamaConfigBanco
End Sub

Private Sub grdMENU_DblClick()
    Call ChamaTelas
    If boolChamaTela = False Then Call ExpNoExpGride
End Sub

Private Sub grdMENU_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call ChamaTelas
        If boolChamaTela = False Then Call ExpNoExpGride
    ElseIf KeyCode = vbKeyLeft Then
        If grdMENU.GetNode(grdMENU.Row).Expanded = False Then Exit Sub
        Call ExpNoExpGride
    ElseIf KeyCode = vbKeyRight Then
        Call ExpNoExpGride
    End If
End Sub

Private Sub mnSair_Click()
    If Sair = True Then Call FechaCampo
End Sub

Private Sub stbMensagem_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If strUSUARIO = "SGI" And Button = 2 Then ChamaConfigBanco
End Sub

''Private Sub TbrMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
      ''Call Configuracao(Button.ToolTipText, Button.Key)
''End Sub

Private Sub tmrHora_Timer()
  stbMensagem.Panels(4).Text = CabecForm("Hora : ") & Format(Time, "HH:MM:SS")
  stbMensagem.Panels(3).Text = CabecForm("Data : ") & DateSerial(Year(Date), Month(Date), Day(Date))
End Sub

Private Sub Configuracao(Menu As String, Cigla As String)
  
On Error GoTo err_descr
  
  Dim StrImage     As String
  
  intFILIALPED01 = -1
  
  If Len(Trim(Cigla)) = 0 Then Exit Sub
 
  sSql = ""
  
  sSql = "Select " & vbCrLf
  sSql = sSql & "       * " & vbCrLf
  sSql = sSql & "  From " & vbCrLf
  sSql = sSql & "       SGI_MENUP " & vbCrLf
  sSql = sSql & " Where " & vbCrLf
  If strUSUARIO = "CWS" Then
     sSql = sSql & "       SGI_FILIAL = 0"
  Else
     sSql = sSql & "       SGI_FILIAL =  " & iFilial
  End If
  sSql = sSql & "   And SGI_TEXTO  = '" + Menu + "'" & vbCrLf
  sSql = sSql & "   And SGI_CIGLA2 = '" + Cigla + "'"
 
  Call AbBanco(strLINHA)
  BREC.Open sSql, BD, adOpenDynamic

  If Not BREC.EOF Then
     
     If BREC!SGI_TIPO <> "M" Then
        BREC.Close
        Exit Sub
     End If
     
     If Trim(BREC!SGI_MODULO) > 0 Then
        
        Dim objCham As Object
        Set objCham = CreateObject(Trim(BREC!SGI_MODULO))
        
        If Cigla = "CV02" Then intFILIALPED01 = 0 '' Pedidos Novalata
        If Cigla = "CV03" Then intFILIALPED01 = 1 '' Pedidos Stil Row
        
        If Cigla = "CV04" Then intFILIALPED01 = 0  ''Ordem de Faturamento Novalata
        If Cigla = "CV06" Then intFILIALPED01 = 1  ''Ordem de Faturamento Stil Row
        
        If Cigla = "CV05" Then intFILIALPED01 = 0  ''Confirma Ordem de Faturamento Novalata
        If Cigla = "CV07" Then intFILIALPED01 = 1 '' Confirma Ordem de Faturamento Stil Row
        
        If Cigla = "MZ1" Then intFILIALPED01 = 0  ''OP Novalata
        If Cigla = "MZ07" Then intFILIALPED01 = 1 ''OP Stil Row
        
        If Cigla = "MZ04" Then intFILIALPED01 = 0  ''OP Novalata
        If Cigla = "MZ08" Then intFILIALPED01 = 1  ''OP Steel Rol
        
        If Cigla = "MZ06" Then intFILIALPED01 = 0  ''Apontamento Novalata
        If Cigla = "MZ09" Then intFILIALPED01 = 1  ''Apontamento Steel Rol
        
        If Cigla = "PZ011" Then intFILIALPED01 = 0  ''Plano Mestre NOVALATA
        If Cigla = "PZ019" Then intFILIALPED01 = 1  'Plano Mestre Steel Rol
        
        If Cigla = "PZ020" Then intFILIALPED01 = 0  ''Plano Mestre NOVALATA
        If Cigla = "PZ021" Then intFILIALPED01 = 1  'Plano Mestre Steel Rol
        
        If Cigla = "MZ05" Then intFILIALPED01 = 0  ''OP Enviada Montagem NOVALATA
        If Cigla = "MZ09" Then intFILIALPED01 = 1  ''OP Enviada Montagem STEEL
        
        If Cigla = "MZ10" Then intFILIALPED01 = 1  ''Programação de Litografia STEEL
        ''If Cigla = "MZ11" Then intFILIALPED01 = 0  ''Programação de Litografia NOVALATA
        
        If Cigla = "MZ11" Then intFILIALPED01 = 1  ''Ordem de Fabricação de Componentes STEEl-ROL
        If Cigla = "MZ03" Then intFILIALPED01 = 0  ''Ordem de Fabricação de Componentes NOVALATA
        
        If Cigla = "ME4" Then intFILIALPED01 = 1   ''Entrada de Produtos ( Litografia )
        If Cigla = "ME5" Then intFILIALPED01 = 0   ''Saida de Produtos ( Litografia )
        
        
        If (Cigla = "MZ1" Or Cigla = "MZ07" Or _
            Cigla = "CV04" Or Cigla = "CV05" Or _
            Cigla = "CV06" Or Cigla = "CV07" Or _
            Cigla = "MZ04" Or Cigla = "MZ08" Or _
            Cigla = "MZ06" Or Cigla = "MZ09" Or _
            Cigla = "PZ011" Or Cigla = "PZ019" Or _
            Cigla = "PZ020" Or Cigla = "PZ021" Or _
            Cigla = "MZ05" Or Cigla = "MZ09" Or _
            Cigla = "MZ10" Or Cigla = "MZ11" Or _
            Cigla = "MZ03" Or Cigla = "ME4" Or _
            Cigla = "ME5") Then
           objCham.cConnect App.Path, Linha, iFilial, strACESSO, strUSUARIO, V_UsuarioId, intFILIALPED01
        ElseIf (Cigla = "CV02" Or Cigla = "CV03") Then
           ''objCham.cConnect App.Path, Linha, iFilial, strACESSO, strUSUARIO, V_UsuarioId, intFILIALPED01, strVERSAO, strNOMMAQUINA
           objCham.cConnect App.Path, Linha, iFilial, strACESSO, strUSUARIO, V_UsuarioId, intFILIALPED01, strVERSAO
        Else
           objCham.cConnect App.Path, Linha, iFilial, strACESSO, strUSUARIO, V_UsuarioId
        End If
        
        Set objCham = Nothing
    
        ''Open App.Path & "\" & "SIGE.txt" For Input As #1
        ''intregs = 0
      
        ''Do While Not EOF(1)
     
        ''   ReDim Preserve Linha(intregs)
      
        ''   Input #1, Linha(intregs)
        ''   intregs = intregs + 1
         
        ''Loop
   
        ''Close #1
        ' --------------------------------------
  
        ''objFuncoes.CarregaCor Linha,
        
        
        ''If Len(Trim(Trim(Mid(Linha(6), 7, 200)))) > 0 Then
        ''    StrImage = Trim(Mid(Linha(5), 8, 200)) & Trim(Mid(Linha(6), 7, 200))
            ''Image1.Picture = LoadPicture(StrImage)
        ''End If
     Else
        MsgBox "Módulo ainda não está disponivel !!!", vbOKOnly + vbCritical, "Aviso"
     End If
  
  End If
  BREC.Close
  
  Call FcBanco
  
  Exit Sub
  
err_descr:

    If Err.Number = -2147024770 Then
       MsgBox "Módulo ainda não está disponivel !!!", vbOKOnly + vbCritical, "Aviso"
    Else
       MsgBox "Erro nº   : " & Err.Number & vbCrLf & _
              "Erro Desc : " & Err.Description, vbOKOnly + vbCritical, "Aviso"
    End If
    
    If BREC.State = 1 Then BREC.Close
    Set objCham = Nothing

End Sub

Private Function Sair() As Boolean

    Dim iResp As Integer
    
    iResp = MsgBox("Deseja realmente sair do sistema ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
      
    If iResp = 6 Then Sair = True
   
End Function

Private Sub ChamaConfigBanco()
    Dim objConectBanco As Object
    Set objConectBanco = CreateObject("CONFIG.clsConfig")
    objConectBanco.cConnect App.Path, Linha, iFilial, strACESSO, strUSUARIO
End Sub

Private Sub Logon()

    frmAcessoLogon.Show vbModal
    Form_Load

End Sub

Private Function GeraMenuScript() As Boolean

On Error GoTo Err_GravaMenu

    GeraMenuScript = False

    adoBanco_Dados.BeginTrans
    BGRV.ActiveConnection = adoBanco_Dados
    
    sSql = "Insert into SGI_MENUP (SGI_FILIAL,SGI_CODIGO,SGI_TEXTO,SGI_TIPO,SGI_CIGLA,SGI_CIGLA2,SGI_MODULO,SGI_ACESSO,SGI_DEPTO,SGI_NIVEL,SGI_CODGER,SGI_DESNIV)"
    sSql = sSql & "        Values (0,1,'Financeiro','P','F','NULL','NULL','IAECR',0,0,0,Null)"
    BGRV.CommandText = sSql
    BGRV.Execute
    
    sSql = "Insert into SGI_MENUP (SGI_FILIAL,SGI_CODIGO,SGI_TEXTO,SGI_TIPO,SGI_CIGLA,SGI_CIGLA2,SGI_MODULO,SGI_ACESSO,SGI_DEPTO,SGI_NIVEL,SGI_CODGER,SGI_DESNIV)"
    sSql = sSql & "        Values (0,2,'Cadastros','S','F','FC','NULL','IAECR',0,0,0,Null)"
    BGRV.CommandText = sSql
    BGRV.Execute

    sSql = "Insert into SGI_MENUP (SGI_FILIAL,SGI_CODIGO,SGI_TEXTO,SGI_TIPO,SGI_CIGLA,SGI_CIGLA2,SGI_MODULO,SGI_ACESSO,SGI_DEPTO,SGI_NIVEL,SGI_CODGER,SGI_DESNIV)"
    sSql = sSql & "        Values (0,3,'Movimentos','S','F','FM','NULL','IAECR',0,0,0,Null)"
    BGRV.CommandText = sSql
    BGRV.Execute

    sSql = "Insert into SGI_MENUP (SGI_FILIAL,SGI_CODIGO,SGI_TEXTO,SGI_TIPO,SGI_CIGLA,SGI_CIGLA2,SGI_MODULO,SGI_ACESSO,SGI_DEPTO,SGI_NIVEL,SGI_CODGER,SGI_DESNIV)"
    sSql = sSql & "        Values (0,4,'Relatórios','S','F','FR','NULL','IAECR',0,0,0,Null)"
    BGRV.CommandText = sSql
    BGRV.Execute

    sSql = "Insert into SGI_MENUP (SGI_FILIAL,SGI_CODIGO,SGI_TEXTO,SGI_TIPO,SGI_CIGLA,SGI_CIGLA2,SGI_MODULO,SGI_ACESSO,SGI_DEPTO,SGI_NIVEL,SGI_CODGER,SGI_DESNIV)"
    sSql = sSql & "        Values (0,5,'Condição de Pagamento','M','FC','FP1','CADCONDPAGTO.clsCADCONDPAGTO','IAECR',0,0,0,Null)"
    BGRV.CommandText = sSql
    BGRV.Execute

    sSql = "Insert into SGI_MENUP (SGI_FILIAL,SGI_CODIGO,SGI_TEXTO,SGI_TIPO,SGI_CIGLA,SGI_CIGLA2,SGI_MODULO,SGI_ACESSO,SGI_DEPTO,SGI_NIVEL,SGI_CODGER,SGI_DESNIV)"
    sSql = sSql & "        Values (0,6,'Bancos','M','FC','FP2','CADBANCOS.clsCADBANCOS','IAECR',0,0,0,Null)"
    BGRV.CommandText = sSql
    BGRV.Execute

    sSql = "Insert into SGI_MENUP (SGI_FILIAL,SGI_CODIGO,SGI_TEXTO,SGI_TIPO,SGI_CIGLA,SGI_CIGLA2,SGI_MODULO,SGI_ACESSO,SGI_DEPTO,SGI_NIVEL,SGI_CODGER,SGI_DESNIV)"
    sSql = sSql & "  Values (0,7,'Comercial','P','C','NULL','NULL','IAECR',0,0,0,Null)"
    BGRV.CommandText = sSql
    BGRV.Execute
    
    sSql = "Insert into SGI_MENUP (SGI_FILIAL,SGI_CODIGO,SGI_TEXTO,SGI_TIPO,SGI_CIGLA,SGI_CIGLA2,SGI_MODULO,SGI_ACESSO,SGI_DEPTO,SGI_NIVEL,SGI_CODGER,SGI_DESNIV)"
    sSql = sSql & "        Values (0,8,'Cadastros','S','C','CC','NULL','IAECR',0,0,0,Null)"
    BGRV.CommandText = sSql
    BGRV.Execute

    sSql = "Insert into SGI_MENUP (SGI_FILIAL,SGI_CODIGO,SGI_TEXTO,SGI_TIPO,SGI_CIGLA,SGI_CIGLA2,SGI_MODULO,SGI_ACESSO,SGI_DEPTO,SGI_NIVEL,SGI_CODGER,SGI_DESNIV)"
    sSql = sSql & "        Values (0,9,'Movimentos','S','C','CM','NULL','IAECR',0,0,0,Null)"
    BGRV.CommandText = sSql
    BGRV.Execute

    sSql = "Insert into SGI_MENUP (SGI_FILIAL,SGI_CODIGO,SGI_TEXTO,SGI_TIPO,SGI_CIGLA,SGI_CIGLA2,SGI_MODULO,SGI_ACESSO,SGI_DEPTO,SGI_NIVEL,SGI_CODGER,SGI_DESNIV)"
    sSql = sSql & "        Values (0,10,'Relatórios','S','C','CR','NULL','IAECR',0,0,0,Null)"
    BGRV.CommandText = sSql
    BGRV.Execute
    
    sSql = "Insert into SGI_MENUP (SGI_FILIAL,SGI_CODIGO,SGI_TEXTO,SGI_TIPO,SGI_CIGLA,SGI_CIGLA2,SGI_MODULO,SGI_ACESSO,SGI_DEPTO,SGI_NIVEL,SGI_CODGER,SGI_DESNIV)"
    sSql = sSql & "        Values (0,11,'Clientes','M','CC','CO1','CADCLIENTE.clsCADCLIENTE','IAECR',0,0,0,Null)"
    BGRV.CommandText = sSql
    BGRV.Execute

    sSql = "Insert into SGI_MENUP (SGI_FILIAL,SGI_CODIGO,SGI_TEXTO,SGI_TIPO,SGI_CIGLA,SGI_CIGLA2,SGI_MODULO,SGI_ACESSO,SGI_DEPTO,SGI_NIVEL,SGI_CODGER,SGI_DESNIV)"
    sSql = sSql & "        Values (0,12,'Segmento do Cliente','M','CC','CO2','CADSEGMENTO.clsCADSEGMENTO','IAECR',0,0,0,Null)"
    BGRV.CommandText = sSql
    BGRV.Execute

    sSql = "Insert into SGI_MENUP (SGI_FILIAL,SGI_CODIGO,SGI_TEXTO,SGI_TIPO,SGI_CIGLA,SGI_CIGLA2,SGI_MODULO,SGI_ACESSO,SGI_DEPTO,SGI_NIVEL,SGI_CODGER,SGI_DESNIV)"
    sSql = sSql & "        Values (0,19,'Cadastro de Vendedores','M','CC','CO3','CADVENDEDOR.clsCADVENDEDOR','IAECR',0,0,0,Null)"
    BGRV.CommandText = sSql
    BGRV.Execute

    sSql = "Insert into SGI_MENUP (SGI_FILIAL,SGI_CODIGO,SGI_TEXTO,SGI_TIPO,SGI_CIGLA,SGI_CIGLA2,SGI_MODULO,SGI_ACESSO,SGI_DEPTO,SGI_NIVEL,SGI_CODGER,SGI_DESNIV)"
    sSql = sSql & "        Values (0,20,'Cadastro de Tabela de Gastos','M','CC','CO4','CADTABGASTOS.clsCADTABGASTOS','IAECR',0,0,0,Null)"
    BGRV.CommandText = sSql
    BGRV.Execute

    sSql = "Insert into SGI_MENUP (SGI_FILIAL,SGI_CODIGO,SGI_TEXTO,SGI_TIPO,SGI_CIGLA,SGI_CIGLA2,SGI_MODULO,SGI_ACESSO,SGI_DEPTO,SGI_NIVEL,SGI_CODGER,SGI_DESNIV)"
    sSql = sSql & "        Values (0,21,'Tipo de Orçamentos','M','CC','CO5','CADESPORCA.clsCADESPORCA','IAECR',0,0,0,Null)"
    BGRV.CommandText = sSql
    BGRV.Execute
    
    sSql = "Insert into SGI_MENUP (SGI_FILIAL,SGI_CODIGO,SGI_TEXTO,SGI_TIPO,SGI_CIGLA,SGI_CIGLA2,SGI_MODULO,SGI_ACESSO,SGI_DEPTO,SGI_NIVEL,SGI_CODGER,SGI_DESNIV)"
    sSql = sSql & "        Values (0,22,'Tabela de serviços','M','CC','CO6','CADSERVICOS.clsCADSERVICOS','IAECR',0,0,0,Null)"
    BGRV.CommandText = sSql
    BGRV.Execute

    sSql = "Insert into SGI_MENUP (SGI_FILIAL,SGI_CODIGO,SGI_TEXTO,SGI_TIPO,SGI_CIGLA,SGI_CIGLA2,SGI_MODULO,SGI_ACESSO,SGI_DEPTO,SGI_NIVEL,SGI_CODGER,SGI_DESNIV)"
    sSql = sSql & "        Values (0,50,'Suprimentos','P','S','NULL','NULL','IAECR',0,0,0,Null)"
    BGRV.CommandText = sSql
    BGRV.Execute
    
    adoBanco_Dados.CommitTrans
    
    GeraMenuScript = True
    
    Exit Function
    
Err_GravaMenu:

    MsgBox "Erro Nº: " & Err.Number & vbCrLf & _
           "Descrição : " & Err.Description & vbCrLf & _
           "Função : GeraMenuScript " & vbCrLf & _
           "Módulo : " & Me.Name, vbOKOnly + vbCritical, "Erro"
           
    adoBanco_Dados.RollbackTrans
    
End Function

Private Sub Jobs()

     Dim dtDATA As Date
     
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       * " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADPARAMPEDCOT " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_ATIVO  = 1" & vbCrLf
     sSql = sSql & "   And SGI_FILIAL = " & intFilial
    
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     If Not BREC.EOF Then
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "        SGI_DATACOTA " & vbCrLf
        sSql = sSql & "       ,SGI_CODIGO   " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_CADCOTAVENDH " & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "   And SGI_STATUS = 'A'"
        
        BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
        Do While Not BREC2.EOF
            
            If (Date - BREC2!SGI_DATACOTA) > BREC!SGI_DIASCOT Then
                '' Muda o status das cotações para não Atendida
                sSql = "Update SGI_CADCOTAVENDH Set SGI_STATUS = 'N'" & vbCrLf
                sSql = sSql & "                      Where " & vbCrLf
                sSql = sSql & "                            SGI_FILIAL = " & intFilial & vbCrLf
                sSql = sSql & "                        And SGI_CODIGO = " & BREC2!SGI_CODIGO
                
                BGRV.CommandText = sSql
                BGRV.Execute
            End If
            
            BREC2.MoveNext
        Loop
        BREC2.Close
        
     End If
     BREC.Close
     
     adoBanco_Dados.CommitTrans
     
     Exit Sub
     
err_grava:
     
     adoBanco_Dados.RollbackTrans
    
     Dim objErro    As Object
     Set objErro = CreateObject("BLBCWS.clsFuncoes")
     Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, "", sSql)
     Set objErro = Nothing

End Sub

Private Sub FechaCampo()
    If GravaLogSaida(strLINHA, intFilial, lngCODACESSO) = False Then Exit Sub
    End
End Sub

Public Sub CriaToolBar(ilsList As ImageList, TbrMenu As Toolbar)
   
   ''Dim btn As MSComctlLib.Button
   
   ''Set TbrMenu.ImageList = ilsList
   ''Set btn = TbrMenu.Buttons.Add(, "PED_NOVALATA", , , "PED_NOVALATA")
   ''btn.ToolTipText = "Pedidos Novalata"
   
   
   ''Set btn = TbrMenu.Buttons.Add(, "PED_NOVALATA", , , "PED_NOVALATA")
   ''btn.ToolTipText = "Pedidos Novalata"
   
   ''Set btn = TbrMenu.Buttons.Add(, "PED_STEEL", , , "PED_NOVALATA")
   ''btn.ToolTipText = "Pedidos Steel"
   
   ''Set btn = TbrMenu.Buttons.Add(, , , tbrSeparator)
   ''Set btn = TbrMenu.Buttons.Add(, , , tbrSeparator)
   
   ''Set btn = TbrMenu.Buttons.Add(, "Incluir", , , "Key01")
   ''btn.ToolTipText = "Incluir Dados"
   
   ''Set btn = TbrMenu.Buttons.Add(, "Alterar", , , "Key01")
   ''btn.ToolTipText = "Alterar dados"

   ''Set btn = TbrMenu.Buttons.Add(, "Excluir", , , "Key01")
   ''btn.ToolTipText = "Excluir Dados"

   ''Set btn = TbrMenu.Buttons.Add(, , , tbrSeparator)
   ''Set btn = TbrMenu.Buttons.Add(, , , tbrSeparator)
   
   ''Set btn = TbrMenu.Buttons.Add(, "Pesquizar", , , "Key01")
   ''btn.ToolTipText = "Pesquisa Dados"

End Sub

Sub SetDefaults(fa As VSFlexGrid)
    
    With fa
        .BindToArray Null
        .Rows = 0
        .Cols = 0
        .ScrollTrack = False
        .ExplorerBar = flexExNone
        .AutoSearch = flexSearchNone
        .Editable = False
        .AllowUserResizing = flexResizeNone
        .SelectionMode = flexSelectionFree
        .OutlineBar = flexOutlineBarNone
        .OLEDragMode = flexOLEDragManual
        .OLEDropMode = flexOLEDropNone
        .ScrollTips = False
        .ToolTipText = ""
    End With
    
End Sub


Private Sub ConfGridMenu()
    
    ' reset the control
    SetDefaults grdMENU
    
    With grdMENU
    
        
        .Redraw = True
        
        ' set the properties we want
        .Rows = 1
        .FixedRows = 1
        .Cols = 5
        .AllowUserResizing = flexResizeBoth
        .ExtendLastCol = True
        .OutlineBar = flexOutlineBarComplete
        .OutlineCol = 0
        .SubtotalPosition = flexSTAbove
        
        .GridLines = flexGridNone
        .FontName = "Arial"
        .FontBold = True
        .FontSize = 11
        
        ' fill the control with data
        .Cell(flexcpText, 0, 0) = "Descrição"
        .Cell(flexcpText, 0, 1) = "Cigla"
        .Cell(flexcpText, 0, 2) = "Tipo"
        .Cell(flexcpText, 0, 3) = "Classe"
        .Cell(flexcpText, 0, 4) = "Acesso"
        
        '' Tamanho da Coluna
        .ColWidth(0) = 5000
        
        '' Coluna Visivel
        .ColHidden(1) = True
        .ColHidden(2) = True
        .ColHidden(3) = True
        .ColHidden(4) = True
        
        '' Linha Visivel
        .RowHidden(0) = True
    
    End With

    
End Sub

Private Sub CarregaMenuGride()

   Dim lngLINHA_PAI         As Long
   Dim lngLINHA_FILHO       As Long
   
   ''Set TbrMenu.ImageList = ImageList1
   ''Dim btn      As MSComctlLib.Button
   
   sSql = ""
   
   If strUSUARIO = "CWS" Then
      sSql = "Select * " & vbCrLf
      sSql = sSql & "       from " & vbCrLf
      sSql = sSql & "            SGI_MENUP " & vbCrLf
      sSql = sSql & "      Where " & vbCrLf
      sSql = sSql & "            SGI_FILIAL = 0 " & vbCrLf
      sSql = sSql & "        And SGI_CODGER = 0 " & vbCrLf
      sSql = sSql & "        And SGI_TIPO   = 'P'" & vbCrLf
      sSql = sSql & " Order by SGI_CODIGO" & vbCrLf
   Else
      sSql = "Select * from SGI_MENUP " & vbCrLf
      sSql = sSql & "       Where " & vbCrLf
      sSql = sSql & "             SGI_FILIAL = " & iFilial & vbCrLf
      If intNOVO = 1 Then
         sSql = sSql & "         And SGI_CODGER = " & iAcesso & vbCrLf
      ElseIf intNOVO = 0 Then
         sSql = sSql & "         And SGI_CODUSUARIO = " & iCodUsu & vbCrLf
         sSql = sSql & "         And SGI_ATIVO      = 1" & vbCrLf
         sSql = sSql & "         And SGI_TIPO       = 'P'" & vbCrLf
      End If
      sSql = sSql & " Order by SGI_CODIGO"
   End If
   
   Call AbBanco(strNOVACONECT)
   BREC.Open sSql, BD, adOpenDynamic
   
   If Not BREC.EOF Then
   
        mnFinanceiro.Visible = False
        mnComercial.Visible = False
        mnExpedicao.Visible = False
        mnEstoque.Visible = False
        mnSuprimentos.Visible = False
        mnControleDoc.Visible = False
        mnQualidade.Visible = False
        mnPCP.Visible = False
        mnConfiguracao.Visible = False
        
        lngLINHA_PAI = 0
        Do While Not BREC.EOF
          
            '' Criando Gride
            '' Itens Principais (Pai)
            grdMENU.AddItem Trim(BREC!SGI_TEXTO) & vbTab & _
                            Trim(BREC!SGI_CIGLA) & vbTab & _
                            Trim(BREC!SGI_TIPO) & vbTab & _
                            "" & vbTab & _
                            ""

            grdMENU.IsSubtotal(grdMENU.Rows - 1) = True
            grdMENU.RowOutlineLevel((grdMENU.Rows - 1)) = 0
            lngLINHA_PAI = (grdMENU.Rows - 1)
            
            If Trim(BREC!SGI_TEXTO) = "Financeiro" Then
               mnFinanceiro.Visible = True
               mnFinanceiro.Caption = Trim(BREC!SGI_TEXTO)
            ElseIf Trim(BREC!SGI_TEXTO) = "Comercial" Then
               mnComercial.Visible = True
               mnComercial.Caption = Trim(BREC!SGI_TEXTO)
            ElseIf Trim(BREC!SGI_TEXTO) = "Expedição" Then
               mnExpedicao.Visible = True
               mnExpedicao.Caption = Trim(BREC!SGI_TEXTO)
            ElseIf Trim(BREC!SGI_TEXTO) = "Estoque" Then
               mnEstoque.Visible = True
               mnEstoque.Caption = Trim(BREC!SGI_TEXTO)
            ElseIf Trim(BREC!SGI_TEXTO) = "Suprimentos" Then
               mnSuprimentos.Visible = True
               mnSuprimentos.Caption = Trim(BREC!SGI_TEXTO)
            ElseIf Trim(BREC!SGI_TEXTO) = "Controle de Documentos" Then
               mnControleDoc.Visible = True
               mnControleDoc.Caption = Trim(BREC!SGI_TEXTO)
            ElseIf Trim(BREC!SGI_TEXTO) = "Qualidade" Then
               mnQualidade.Visible = True
               mnQualidade.Caption = Trim(BREC!SGI_TEXTO)
            ElseIf Trim(BREC!SGI_TEXTO) = "PCP" Then
               mnPCP.Visible = True
               mnPCP.Caption = Trim(BREC!SGI_TEXTO)
            ElseIf Trim(BREC!SGI_TEXTO) = "Metrologia" Then
               mnMetrologia.Visible = True
               mnMetrologia.Caption = Trim(BREC!SGI_TEXTO)
            ElseIf Trim(BREC!SGI_TEXTO) = "Configurações" Then
               mnConfiguracao.Visible = True
               mnConfiguracao.Caption = Trim(BREC!SGI_TEXTO)
            End If
             
            If Trim(BREC!SGI_CIGLA) <> "G" Then
            
                sSql = ""
                sSql = "Select" & vbCrLf
                sSql = sSql & "      *" & vbCrLf
                sSql = sSql & "  From" & vbCrLf
                sSql = sSql & "       SGI_MENUP" & vbCrLf
                sSql = sSql & " Where" & vbCrLf
                sSql = sSql & "       SGI_FILIAL = " & BREC!SGI_FILIAL & vbCrLf
                sSql = sSql & "   And SGI_TIPO   = 'S'" & vbCrLf
                If intNOVO = 1 Then
                    sSql = sSql & "   And SGI_CODGER = " & iAcesso & vbCrLf
                ElseIf intNOVO = 0 Then
                    sSql = sSql & "   And SGI_CODUSUARIO = " & BREC!SGI_CODUSUARIO & vbCrLf
                    sSql = sSql & "   And SGI_ATIVO      = 1" & vbCrLf
                    sSql = sSql & "   And SGI_CIGLA      = '" & Trim(BREC!SGI_CIGLA) & "'" & vbCrLf
                End If
                sSql = sSql & "Order by SGI_CODIGO"
                
                BREC10.Open sSql, BD, adOpenDynamic
                Do While Not BREC10.EOF()
                
                    grdMENU.AddItem Trim(BREC10!SGI_TEXTO) & vbTab & _
                                    Trim(BREC10!SGI_CIGLA2) & vbTab & _
                                    Trim(BREC10!SGI_TIPO) & vbTab & _
                                    "" & vbTab & _
                                    ""
                                    
                    grdMENU.IsSubtotal(grdMENU.Rows - 1) = True
                    grdMENU.RowOutlineLevel((grdMENU.Rows - 1)) = 1
                    grdMENU.Cell(flexcpPicture, (grdMENU.Rows - 1), 0) = imgItem.Picture
                    
                    lngLINHA_FILHO = (grdMENU.Rows - 1)
                    
                    sSql = ""
                    sSql = "Select" & vbCrLf
                    sSql = sSql & "      *" & vbCrLf
                    sSql = sSql & "  From" & vbCrLf
                    sSql = sSql & "       SGI_MENUP" & vbCrLf
                    sSql = sSql & " Where" & vbCrLf
                    sSql = sSql & "       SGI_FILIAL = " & BREC10!SGI_FILIAL & vbCrLf
                    sSql = sSql & "   And SGI_TIPO   = 'M'" & vbCrLf
                    If intNOVO = 1 Then
                        sSql = sSql & "   And SGI_CODGER = " & iAcesso & vbCrLf
                    ElseIf intNOVO = 0 Then
                        sSql = sSql & "   And SGI_CODUSUARIO = " & BREC10!SGI_CODUSUARIO & vbCrLf
                        sSql = sSql & "   And SGI_ATIVO      = 1" & vbCrLf
                        sSql = sSql & "   And SGI_CIGLA      = '" & Trim(BREC10!SGI_CIGLA2) & "'" & vbCrLf
                    End If
                    sSql = sSql & "Order by SGI_CODIGO"
                    
                    BREC11.Open sSql, BD, adOpenDynamic
                    Do While Not BREC11.EOF()
                    
                        grdMENU.AddItem Trim(BREC11!SGI_TEXTO) & vbTab & _
                                        Trim(BREC11!SGI_CIGLA2) & vbTab & _
                                        Trim(BREC11!SGI_TIPO) & vbTab & _
                                        Trim(BREC11!SGI_MODULO) & vbTab & _
                                        Trim(BREC11!SGI_ACESSO)
                                        
                        grdMENU.IsSubtotal(grdMENU.Rows - 1) = False
                        grdMENU.RowOutlineLevel((grdMENU.Rows - 1)) = 2
                        
                        '' ----------------------------------------
                        '' Ciglas
                        '' ----------------------------------------
                        '' CV02 -- Pedidos Novalata
                        '' CV03 -- Pedidos Steel
                        '' ----------------------------------------
                        '' CV04 -- Ordem de Faturamento Novalata
                        '' CV06 -- Ordem de Faturamento Steel
                        '' ----------------------------------------
                        '' CV05 -- Confirma Ordem de Faturmento Novalata
                        '' CV07 -- Confirma Ordem de Faturmento Steel
                        '' ----------------------------------------
                        '' MZ1  -- Ordem de Produção Novalata
                        '' MZ07 -- Ordem de Produção Novalata
                        '' ----------------------------------------
                        If Trim(BREC11!SGI_CIGLA2) = "CV02" Then          '' CV02 -- Pedidos Novalata
                            ''Set btn = TbrMenu.Buttons.Add(, Trim(BREC11!SGI_CIGLA2), , , "Pict9")
                            ''btn.ToolTipText = Trim(BREC11!SGI_TEXTO)
                        ElseIf Trim(BREC11!SGI_CIGLA2) = "CV03" Then      '' CV03 -- Pedidos Steel
                            ''Set btn = TbrMenu.Buttons.Add(, Trim(BREC11!SGI_CIGLA2), , , "Pict9")
                            ''btn.ToolTipText = Trim(BREC11!SGI_TEXTO)
                        ElseIf Trim(BREC11!SGI_CIGLA2) = "CV04" Then      '' CV04 -- Ordem de Faturamento Novalata
                            ''Set btn = TbrMenu.Buttons.Add(, , , tbrSeparator)
                            ''Set btn = TbrMenu.Buttons.Add(, , , tbrSeparator)
                            ''Set btn = TbrMenu.Buttons.Add(, Trim(BREC11!SGI_CIGLA2), , , "Pict9")
                            ''btn.ToolTipText = Trim(BREC11!SGI_TEXTO)
                        ElseIf Trim(BREC11!SGI_CIGLA2) = "CV06" Then      '' CV06 -- Ordem de Faturmento Steel
                            ''Set btn = TbrMenu.Buttons.Add(, Trim(BREC11!SGI_CIGLA2), , , "Pict9")
                            ''btn.ToolTipText = Trim(BREC11!SGI_TEXTO)
                        ElseIf Trim(BREC!SGI_CIGLA2) = "CV05" Then      '' CV05 -- Confirma Ordem de Faturmento Novalata
                            ''Set btn = TbrMenu.Buttons.Add(, , , tbrSeparator)
                            ''Set btn = TbrMenu.Buttons.Add(, , , tbrSeparator)
                            ''Set btn = TbrMenu.Buttons.Add(, Trim(BREC11!SGI_CIGLA2), , , "Pict9")
                            ''btn.ToolTipText = Trim(BREC11!SGI_TEXTO)
                        ElseIf Trim(BREC11!SGI_CIGLA2) = "CV07" Then      '' CV07 -- Confirma Ordem de Faturmento Novalata
                            ''Set btn = TbrMenu.Buttons.Add(, Trim(BREC11!SGI_CIGLA2), , , "Pict9")
                            ''btn.ToolTipText = Trim(BREC11!SGI_TEXTO)
                        ElseIf Trim(BREC11!SGI_CIGLA2) = "MZ1" Then       '' MZ1 -- Ordem de Produção Novalata
                            ''Set btn = TbrMenu.Buttons.Add(, , , tbrSeparator)
                            ''Set btn = TbrMenu.Buttons.Add(, , , tbrSeparator)
                            ''Set btn = TbrMenu.Buttons.Add(, Trim(BREC11!SGI_CIGLA2), , , "Pict9")
                            ''btn.ToolTipText = Trim(BREC11!SGI_TEXTO)
                        ElseIf Trim(BREC11!SGI_CIGLA2) = "MZ07" Then      '' MZ07 -- Ordem de Produção Steel
                            ''Set btn = TbrMenu.Buttons.Add(, Trim(BREC11!SGI_CIGLA2), , , "Pict9")
                            ''btn.ToolTipText = Trim(BREC11!SGI_TEXTO)
                        End If
                        
                        BREC11.MoveNext
                    Loop
                    BREC11.Close
                    
                    grdMENU.GetNode(lngLINHA_FILHO).Expanded = False
                    
                    BREC10.MoveNext
                Loop
                BREC10.Close
            
            ElseIf Trim(BREC!SGI_CIGLA) = "G" Then
            
                sSql = ""
                sSql = "Select" & vbCrLf
                sSql = sSql & "      *" & vbCrLf
                sSql = sSql & "  From" & vbCrLf
                sSql = sSql & "       SGI_MENUP" & vbCrLf
                sSql = sSql & " Where" & vbCrLf
                sSql = sSql & "       SGI_FILIAL = " & BREC!SGI_FILIAL & vbCrLf
                sSql = sSql & "   And SGI_TIPO   = 'M'" & vbCrLf
                If intNOVO = 1 Then
                    sSql = sSql & "   And SGI_CODGER = " & iAcesso & vbCrLf
                ElseIf intNOVO = 0 Then
                    sSql = sSql & "   And SGI_CODUSUARIO = " & BREC!SGI_CODUSUARIO & vbCrLf
                    sSql = sSql & "   And SGI_ATIVO      = 1" & vbCrLf
                    sSql = sSql & "   And SGI_CIGLA      = '" & Trim(BREC!SGI_CIGLA) & "'" & vbCrLf
                End If
                sSql = sSql & "Order by SGI_CODIGO"
                
                BREC10.Open sSql, BD, adOpenDynamic
                Do While Not BREC10.EOF()
                
                    grdMENU.AddItem Trim(BREC10!SGI_TEXTO) & vbTab & _
                                    Trim(BREC10!SGI_CIGLA2) & vbTab & _
                                    Trim(BREC10!SGI_TIPO) & vbTab & _
                                    Trim(BREC10!SGI_MODULO) & vbTab & _
                                    Trim(BREC10!SGI_ACESSO)
                                    
                    grdMENU.IsSubtotal(grdMENU.Rows - 1) = False
                    grdMENU.RowOutlineLevel((grdMENU.Rows - 1)) = 1
                    lngLINHA_FILHO = (grdMENU.Rows - 1)
                    
                    grdMENU.GetNode(lngLINHA_FILHO).Expanded = False
                    
                    BREC10.MoveNext
                Loop
                BREC10.Close
            
            
            End If
            
            grdMENU.GetNode(lngLINHA_PAI).Expanded = False
            
            BREC.MoveNext
        Loop
    
        
        '' Sair do Sistema
        grdMENU.AddItem "Sair do Sistema" & vbTab & _
                        "SAIR" & vbTab & _
                        "SAIR" & vbTab & _
                        "" & vbTab & _
                        ""
        grdMENU.IsSubtotal(grdMENU.Rows - 1) = True
        grdMENU.RowOutlineLevel((grdMENU.Rows - 1)) = 0
    
    End If
    
    BREC.Close
    Call FcBanco

    Call Carrega_Imagem_Menu

End Sub


Private Sub CarregaForm(strCIGLA As String, strModulo As String, strPERM As String)
  
On Error GoTo err_descr
  
    Dim StrImage                    As String
    Dim objCham                     As Object
    Dim arrCIGLAS_NOVALATA()        As String
    Dim arrCIGLAS_STEEL()           As String
    Dim i                           As Long
    
    Set objCham = CreateObject(Trim(strModulo))
    intFILIALPED01 = -1
        
    '' Array de Ciglas Novalata
    ReDim arrCIGLAS_NOVALATA(1 To 11) As String
    arrCIGLAS_NOVALATA(1) = "CV02"
    arrCIGLAS_NOVALATA(2) = "CV04"
    arrCIGLAS_NOVALATA(3) = "CV05"
    arrCIGLAS_NOVALATA(4) = "MZ1"
    arrCIGLAS_NOVALATA(5) = "MZ04"
    arrCIGLAS_NOVALATA(6) = "MZ06"
    arrCIGLAS_NOVALATA(7) = "PZ011"
    arrCIGLAS_NOVALATA(8) = "PZ020"
    arrCIGLAS_NOVALATA(9) = "MZ05"
    arrCIGLAS_NOVALATA(11) = "MZ03"
    
    
    For i = 1 To UBound(arrCIGLAS_NOVALATA)
        If strCIGLA = arrCIGLAS_NOVALATA(i) Then
            intFILIALPED01 = 0 '' Novalata
            Exit For
        End If
    Next i
        
    '' Array de Ciglas Steel
    ReDim arrCIGLAS_STEEL(1 To 11) As String
    arrCIGLAS_STEEL(1) = "CV03"
    arrCIGLAS_STEEL(2) = "CV06"
    arrCIGLAS_STEEL(3) = "CV07"
    arrCIGLAS_STEEL(4) = "MZ07"
    arrCIGLAS_STEEL(5) = "MZ08"
    arrCIGLAS_STEEL(6) = "MZ09"
    arrCIGLAS_STEEL(7) = "PZ019"
    arrCIGLAS_STEEL(8) = "PZ021"
    arrCIGLAS_STEEL(9) = "MZ09"
    arrCIGLAS_STEEL(10) = "MZ10"
    arrCIGLAS_STEEL(11) = "MZ11"
    
    
    For i = 1 To UBound(arrCIGLAS_STEEL)
        If strCIGLA = arrCIGLAS_STEEL(i) Then
            intFILIALPED01 = 1 '' Steel
            Exit For
        End If
    Next i
        
    If intFILIALPED01 > -1 Then objCham.cConnect App.Path, Linha, iFilial, strPERM, strUSUARIO, V_UsuarioId, intFILIALPED01
    If intFILIALPED01 < 0 Then objCham.cConnect App.Path, Linha, iFilial, strPERM, strUSUARIO, V_UsuarioId
        
    Set objCham = Nothing
    
    boolChamaTela = True
    
    Exit Sub
  
err_descr:

    Set objCham = Nothing
    If Err.Number = -2147024770 Or Err.Number = 429 Then
       MsgBox "Módulo ainda não está disponivel !!!", vbOKOnly + vbCritical, "Aviso"
    Else
       MsgBox "Erro nº   : " & Err.Number & vbCrLf & _
              "Erro Desc : " & Err.Description, vbOKOnly + vbCritical, "Aviso"
    End If
    

End Sub

Private Sub ChamaTelas()
    With grdMENU
        If (.Rows - 1) = 0 Then Exit Sub
        If (.Row) <= 0 Then Exit Sub
        boolChamaTela = False
        If Trim(.Cell(flexcpText, .Row, 2)) = "M" Then
            Call CarregaForm(.Cell(flexcpText, .Row, 1), .Cell(flexcpText, .Row, 3), .Cell(flexcpText, .Row, 4))
        ElseIf Trim(.Cell(flexcpText, .Row, 2)) = "SAIR" Then
            If Sair = True Then Call FechaCampo
        End If
    End With
End Sub

Private Sub ExpNoExpGride()
    With grdMENU
        If (.Rows - 1) = 0 Then Exit Sub
        If (.Row) <= 0 Then Exit Sub
        If .GetNode(.Row).Expanded = False Then
            .GetNode(.Row).Expanded = True
        ElseIf .GetNode(.Row).Expanded = True Then
            .GetNode(.Row).Expanded = False
        End If
    End With
End Sub


Private Sub Carrega_Imagem_Menu()
    '' Imagem
    Dim i   As Long
    
    With grdMENU
        For i = 1 To (.Rows - 1)
            If .Cell(flexcpText, i, 2) = "P" Then '' Financeiro
                If .Cell(flexcpText, i, 1) = "F" Then '' Financeiro
                    ''.Cell(flexcpPicture, i, 0) = ImageList1.ListImages(21).Picture
                ElseIf .Cell(flexcpText, i, 1) = "C" Then '' Comercial
                    ''.Cell(flexcpPicture, i, 0) = ImageList1.ListImages(22).Picture
                Else
                    .Cell(flexcpPicture, i, 0) = imgFolder
                End If
            ElseIf .Cell(flexcpText, i, 2) = "SAIR" Then
                .Cell(flexcpPicture, i, 0) = imgSair.Picture
            End If
            '' ========================
        Next i
    End With
End Sub

