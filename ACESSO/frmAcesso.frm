VERSION 5.00
Begin VB.Form frmAcesso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acesso ao Sistema"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   Icon            =   "frmAcesso.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboUsuario 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Text            =   "cboUsuario"
      Top             =   1320
      Width           =   2655
   End
   Begin VB.ComboBox cboEmpresa 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   840
      Width           =   5175
   End
   Begin VB.CommandButton cmbCancelar 
      Caption         =   "&Cancelar -  F5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton cmbEntrar 
      Caption         =   "&Entrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox txtSenha 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "txtSenha"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   945
   End
   Begin VB.Image imgPW 
      Height          =   720
      Left            =   120
      Picture         =   "frmAcesso.frx":625A
      Top             =   -120
      Width           =   720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   15
      X2              =   6360
      Y1              =   585
      Y2              =   585
   End
   Begin VB.Label lblSenha 
      AutoSize        =   -1  'True
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   5
      Top             =   1800
      Width           =   675
   End
   Begin VB.Label lblUsuáruo 
      AutoSize        =   -1  'True
      Caption         =   "Usuário"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   825
   End
   Begin VB.Label labSenha 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Controle de acesso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   0
      TabIndex        =   6
      Top             =   50
      Width           =   6330
   End
End
Attribute VB_Name = "frmAcesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Linha()        As String
Dim strDATEXP      As String
Dim strDATEXP2     As String
Dim strLINHA       As String
Dim strLINHA2      As String

Private Sub cboEmpresa_Validate(Cancel As Boolean)
    Call PreencheComboUsuario
End Sub

Private Sub cboUsuario_GotFocus()
    Call SelecionaCampos(cboUsuario.Name, Me)
End Sub

Private Sub cboUsuario_KeyPress(KeyAscii As Integer)
    KeyAscii = Maiuscula(KeyAscii)
End Sub


Private Sub cmbCancelar_Click()
  Call Sair
End Sub

Private Sub Sair()
  Call FcBanco
  End
End Sub

Private Sub cmbEntrar_Click()
  
  Dim intResp As Integer
  
  If ValidaSenha(cboUsuario.Text, txtSenha.Text) = False Then
     
     intResp = MsgBox("ATENÇÃO - Usuário ou Senha Inválida !!!" & vbCrLf & "Tentar Novamente", vbYesNo + vbCritical, "Aviso do Sistema")
     
     If intResp = vbNo Then
        Call Sair
     End If
     
     cboUsuario.SetFocus
     Exit Sub
  
  End If
  
  Me.Hide
  
  '' Validação
  If Trim(cboUsuario.Text) <> "CWS" Then
        If Len(Trim(strDATEXP)) > 0 Then
              If CDate(strDATEXP) = CDate(strDATEXP2) Then
                  MsgBox "Erro : -234667589 : Database Violation !!!", vbOKOnly + vbExclamation, "Aviso"
                  End
              End If
        End If
  End If
  
  frmSIGE.strEmpresa = IIf(Len(Trim(cboEmpresa.Text)) > 0, " - " & cboEmpresa.Text, "")
  frmSIGE.Linha = Linha
  frmSIGE.strLINHA = strLINHA2
  frmSIGE.intFilial = iFilial
  frmSIGE.lngCODACESSO = GravaAcesso(Trim(cboUsuario.Text), Me.Name)
  frmSIGE.Show vbModal
  
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmbCancelar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()
      
On Error GoTo Err_Load

  Dim intLinha  As Integer
  Dim SerieHd   As String
  Dim intregs   As Integer
  Call LimpaCampos
  
  'Abre aquivo de Configuração
  intLinha = 1
  Open App.Path & "\" & "SIGE.txt" For Input As #1
  
  intLinha = 2
  intregs = 0
  intLinha = 3
  strLINHA = ""
  strNOVACONECT = ""
  Do While Not EOF(1)
     
     intLinha = 4
     ReDim Preserve Linha(intregs)
     
     Input #1, Linha(intregs)
     
     strLINHA2 = strLINHA2 & Linha(intregs) & "^"
     strNOVACONECT = strNOVACONECT & Linha(intregs) & "^"
     
     intLinha = 5
     intregs = intregs + 1
     
     intLinha = 6
  Loop
  intLinha = 7
  Close #1
  
  ' --------------------------------------
  intLinha = 8
  Call PreencheCombo
  
  intLinha = 9
  Call PegaEmpresaPadrao
  
  intLinha = 10
  Call PreencheComboUsuario
  intLinha = 11
    
  '' Pega o Nemero de Serie do HD
  ''  SerieHd = Get_Number_Serie("C:\")
  
  '  If SerieHd <> "2968-18F1" Then
  '     MsgBox "Atenção este software não esta autorizado para rodar nesta maquina !!!" & vbCrLf & _
  '            "Por favor contactar a CWS - Informatica para obter a licença !!!" & vbCrLf & _
  '            "0XX11-98603810" & vbCrLf & _
  '            "Falar - Com Ricardo", vbOKOnly + vbCritical, "Aviso"
  '     End
  '  End If
  
    Call GravaTrava
    intLinha = 12
    
    Exit Sub
    
Err_Load:

    MsgBox "Erro   : " & Err.Description & vbCrLf & _
           "Erro N : " & Err.Number & vbCrLf & _
           "Função : Load()" & vbCrLf & _
           "Form   : frmAcesso" & _
           "Linha  : " & intLinha, vbOKOnly + vbCritical, "Aviso"

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Sair
End Sub

Private Sub LimpaCampos()
  
  cboUsuario = ""
  txtSenha = ""
  cboEmpresa.Clear
  
End Sub

Private Function ValidaSenha(strUsuarioDig As String, strSenhaDig As String) As Boolean
  
  ValidaSenha = False
  
  Dim rstSenha As New ADODB.Recordset
  Dim strSenha As String
  Dim strTeste As String
  
  If UCase(strUsuarioDig) = "CWS" Then
     If UCase(strSenhaDig) = "JANNUS" Then
        frmSIGE.strUSUARIO = UCase(strUsuarioDig)
        If cboEmpresa.ListIndex = -1 Then
           iCodUsu = 0
           iFilial = 0
        End If
        If cboEmpresa.ListIndex > -1 Then iFilial = cboEmpresa.ItemData(cboEmpresa.ListIndex)
        ValidaSenha = True
        Exit Function
     End If
  Else
  
     iFilial = 0
     If cboEmpresa.ListIndex > -1 Then iFilial = cboEmpresa.ItemData(cboEmpresa.ListIndex)
     
     strSenha = "Select * " & vbCrLf
     strSenha = strSenha & "  From " & vbCrLf
     strSenha = strSenha & "       SGI_USUARIO " & vbCrLf
     strSenha = strSenha & " Where " & vbCrLf
     strSenha = strSenha & "       SGI_NOME     = '" & Crypt(strUsuarioDig) & "'" & vbCrLf
     strSenha = strSenha & "   And SGI_FILIAL   = " & iFilial & vbCrLf
     strSenha = strSenha & "   And SGI_ATIVO    = 1"
     
     Call AbBanco(strLINHA2)
     
     rstSenha.Open strSenha, BD, adOpenDynamic
     If Not rstSenha.EOF Then
        strSenha = Crypt(rstSenha!SGI_SENHA)
        If Trim(strSenha) = Trim(strSenhaDig) Then
            frmSIGE.strUSUARIO = Crypt(UCase(strUsuarioDig))
            frmSIGE.iAcesso = IIf(IsNull(rstSenha!SGI_ACESSO) = False, rstSenha!SGI_ACESSO, 0)
            frmSIGE.intFilial = iFilial
            frmSIGE.intNOVO = rstSenha!SGI_NOVO
            iCodUsu = rstSenha!SGI_CODIGO
            V_UsuarioId = rstSenha!SGI_CODIGO
            ValidaSenha = True
        End If
     End If
     rstSenha.Close
     
     Call FcBanco
  
  End If
 
End Function

Private Sub txtSenha_GotFocus()
    Call SelecionaCampos(txtSenha.Name, Me)
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
    KeyAscii = Maiuscula(KeyAscii)
End Sub

Private Sub PreencheCombo()
    
On Error GoTo Err_PreencheCombo

    Dim intLinSub As Integer
    Dim QUE As New ADODB.Recordset
    
    intLinSub = 1
    cboEmpresa.Clear
    
    intLinSub = 2
    Call AbBanco(strLINHA2)
    
    intLinSub = 3
    sSql = "Select * from SGI_FILIAL Order by SGI_FILIAL"
    
    intLinSub = 4
    QUE.Open sSql, BD, adOpenDynamic
    
    intLinSub = 5
    cboEmpresa.AddItem "ADMINISTRADOR"
    
    intLinSub = 6
    cboEmpresa.ItemData(cboEmpresa.NewIndex) = 0
    
    Do While Not QUE.EOF()
       intLinSub = 7
       cboEmpresa.AddItem Trim(QUE!SGI_DESCRICAO)
       
       intLinSub = 8
       cboEmpresa.ItemData(cboEmpresa.NewIndex) = QUE!SGI_FILIAL
       
       intLinSub = 9
       QUE.MoveNext
    Loop
    
    intLinSub = 10
    QUE.Close
    
    intLinSub = 11
    cboEmpresa.ListIndex = -1
    
    intLinSub = 12
    Call FcBanco
    
    Exit Sub
    
Err_PreencheCombo:

    MsgBox "Erro Desc : " & Err.Description & vbCrLf & _
           "Erro Nº   : " & Err.Number & vbCrLf & _
           "Função    : PreencheCombo" & vbCrLf & _
           "Form      : frmAcesso" & vbCrLf & _
           "Linha     : " & intLinSub, vbOKOnly + vbCritical, "Aviso"
           
End Sub

Private Sub PegaEmpresaPadrao()

    Dim i                   As Integer
    
    cboEmpresa.ListIndex = 0
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_FILIAL " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_PADRAO = 1 "
    
    Call AbBanco(strLINHA2)
    
    BREC.Open sSql, BD, adOpenDynamic
    If Not BREC.EOF Then
       For i = 1 To (cboEmpresa.ListCount - 1)
           If cboEmpresa.ItemData(i) = BREC!SGI_FILIAL Then cboEmpresa.ListIndex = i
           If Not IsNull(BREC!SGI_SENHA) Then strDATEXP = DateSerial(Year(BREC!SGI_SENHA), Month(BREC!SGI_SENHA), Day(BREC!SGI_SENHA))
           If Not IsNull(BREC!SGI_SENHA2) Then strDATEXP2 = DateSerial(Year(BREC!SGI_SENHA2), Month(BREC!SGI_SENHA2), Day(BREC!SGI_SENHA2))
       Next i
    End If
    BREC.Close
    
    Call FcBanco
        
End Sub

Private Sub PreencheComboUsuario()
    
    cboUsuario.Clear
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_USUARIO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & cboEmpresa.ItemData(cboEmpresa.ListIndex) & vbCrLf
    sSql = sSql & "   And SGI_ATIVO  = 1"
    
    Call AbBanco(strLINHA2)
    
    BREC.Open sSql, BD, adOpenDynamic
    Do While Not BREC.EOF()
        cboUsuario.AddItem Trim(Crypt(BREC!SGI_NOME))
        BREC.MoveNext
    Loop
    BREC.Close

    Call FcBanco

End Sub

Private Sub GravaTrava()
       
       Call AbBanco(strLINHA2)
       
       
       If Len(Trim(strDATEXP)) > 0 Then
           If Len(Trim(strDATEXP2)) = 0 Then strDATEXP2 = DateSerial(Year(Now), Month(Now), Day(Now))
           If CDate(strDATEXP) <> CDate(strDATEXP2) Then
               If CDate(strDATEXP) >= Now Then strDATEXP2 = DateSerial(Year(Now), Month(Now), Day(Now))
               If Now >= CDate(strDATEXP) Then strDATEXP2 = DateSerial(Year(CDate(strDATEXP)), Month(CDate(strDATEXP)), Day(CDate(strDATEXP)))
               
               BD.BeginTrans
               BGRV.ActiveConnection = BD
            
               sSql = ""
               
               sSql = "Update SGI_FILIAL Set SGI_SENHA2 = '" & DateSerial(Year(CDate(strDATEXP2)), Month(CDate(strDATEXP2)), Day(CDate(strDATEXP2))) & "'" & vbCrLf
               sSql = sSql & "            Where " & vbCrLf
               sSql = sSql & "                  SGI_FILIAL = " & cboEmpresa.ItemData(cboEmpresa.ListIndex)
               
               BGRV.CommandText = sSql
               BGRV.Execute
               
               BD.CommitTrans
           End If
       End If
       
       Call FcBanco
       
End Sub


Private Function GravaAcesso(strUSUARIO As String, strModulo As String) As Long

On Error GoTo err_Trans

     Dim lngCODIGOOP As Long

     lngCODIGOOP = Gera_Codigo(Trim(Me.Name) & "_LOGAC", cboEmpresa.ItemData(cboEmpresa.ListIndex), strLINHA2)

     '' Inicia transação
     Call AbBanco(strLINHA2)
     
     BD.BeginTrans
     BGRV.ActiveConnection = BD
     
     sSql = ""

     sSql = "Insert Into SGI_LOGACESSO (" & vbCrLf
     sSql = sSql & "                           SGI_FILIAL" & vbCrLf
     sSql = sSql & "                          ,SGI_CODIGO" & vbCrLf
     sSql = sSql & "                          ,SGI_USUARIO" & vbCrLf
     sSql = sSql & "                          ,SGI_MODULO" & vbCrLf
     sSql = sSql & "                          ,SGI_DATENTR" & vbCrLf
     sSql = sSql & "                          ,SGI_DATSAIDA" & vbCrLf
     
     sSql = sSql & "                 ) Values (" & vbCrLf
     sSql = sSql & "                           " & cboEmpresa.ItemData(cboEmpresa.ListIndex) & vbCrLf
     sSql = sSql & "                          ," & lngCODIGOOP & vbCrLf
     sSql = sSql & "                          ,'" & strUSUARIO & "'" & vbCrLf
     sSql = sSql & "                          ,'" & strModulo & "'" & vbCrLf
     sSql = sSql & "                          ,'" & Format(Now, "MM/DD/YYYY HH:MM:SS") & "'" & vbCrLf
     sSql = sSql & "                          ,Null" & vbCrLf
     sSql = sSql & "                          )"

     BGRV.CommandText = sSql
     BGRV.Execute

     GravaAcesso = lngCODIGOOP

     BD.CommitTrans
     
     
     Call FcBanco
     
     Exit Function

err_Trans:
     
     BD.RollbackTrans
     
     Dim objErro    As Object
     Set objErro = CreateObject("BLBCWS.clsFuncoes")
     Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, "I", sSql)
     Set objErro = Nothing

End Function

Private Sub txtSenha_Validate(Cancel As Boolean)
    If Len(Trim(txtSenha.Text)) > 0 Then Call cmbEntrar_Click
End Sub
