VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFuncoes"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim chunk() As Byte
Const conChunkSize = 256

Public Function CabecForm(strCabec As String) As String
  CabecForm = strCabec
End Function

Public Function Banco_Dados(Conecxao As Variant) As ADODB.Connection

On Error GoTo err_sys
  
  Dim Linha()     As Variant
  Dim intregs     As Integer
  Dim i           As Integer
  
  Dim strNoBanco  As String
  Dim strProvider As String
  Dim strConnt    As String
  Dim strLinReg()   As String
    
  '"Microsoft.Jet.OLEDB.4.0"
  '------------------------------------------------------
  'Abre uma conexção com o banco
 
  Set Banco_Dados = New ADODB.Connection
  
  If Right(Conecxao(0), 1) = 1 Then
     Banco_Dados.Provider = Mid(Conecxao(2), 11, 50) '' Drive do banco em Uso
     Banco_Dados.ConnectionString = Mid(Conecxao(1), 9, 50) & Mid(Conecxao(3), 11, 50) '' Caminho do Banco
     Banco_Dados.CommandTimeout = 1200 '' 20 minutos
  End If
  
  If Right(Conecxao(0), 1) = 2 Then
     Banco_Dados.ConnectionString = Conecxao(1) & Conecxao(2) & strPASSWORD
  End If
  
  Banco_Dados.Open
    
  Exit Function
err_sys:
  
  If Err.Number = 53 Then
     MsgBox "ATENÇÃO - Erro : " & Err.Number & " Arquivo de Configuração não encontrado" & vbCrLf & _
            "Função = Banco_Dados" & vbCrLf & _
            "Biblioteca = BLBCWS.clsFuncoes" & vbCrLf & _
            "Caminho = " & App.Path, vbOKOnly + vbCritical, "Aviso"
  Else
     MsgBox "ATENÇÃO - " & Err.Number & " - " & Err.Description, vbOKOnly + vbCritical, "Aviso"
  End If
  
End Function

Public Function Sair(nomForm As Variant) As String
  Unload nomForm
End Function

Public Sub LimpaCampos(frmForm As Variant)
  
  Dim vControl

  'Apagando todos os text boxes de um form:
  '----------------------------------------

  For Each vControl In frmForm.Controls
      If TypeOf vControl Is TextBox Then
         vControl.Text = ""
      'ElseIf TypeOf vControl Is MaskEdBox Then
      '   vControl.Text = "__/__/____"
      ElseIf TypeOf vControl Is OptionButton Then
         vControl.Value = 0
      ElseIf TypeOf vControl Is ComboBox Then
         If vControl.Style = 0 Then
            vControl.Text = ""
            vControl.Clear
         ElseIf vControl.Style = 2 Then
            vControl.Clear
         End If
      ElseIf TypeOf vControl Is ListBox Then
         vControl.Clear
      End If
   Next

End Sub


Public Sub SelecionaCampos(nomecampo As String, frmfornmulario As Variant)
  
  Dim i As Integer
  
  For i = 0 To frmfornmulario.Count - 1
      
      If UCase(frmfornmulario.Controls(i).Name) = UCase(nomecampo) Then
         frmfornmulario.Controls(i).SelStart = 0
         frmfornmulario.Controls(i).SelLength = Len(frmfornmulario.Controls(i).Text)
      End If
  Next
  
End Sub


'Public Function Banco_Dados_Externo(strProvider As String, strConn As String) As ADODB.Connection

'On Error GoTo err_sys
  
  '"Microsoft.Jet.OLEDB.4.0"
  
  '------------------------------------------------------
  'Abre uma conexção com o banco
 
'  Set Banco_Dados_Externo = New ADODB.Connection
  
'  Banco_Dados_Externo.Provider = strProvider '' Drive do banco em Uso
'  Banco_Dados_Externo.ConnectionString = strConn '' Caminho do Banco
  
'  Banco_Dados_Externo.Open
    
'  Exit Function
  
'err_sys:
  
'  If Err.Number = 53 Then
     
'     MsgBox "ATENÇÃO - Erro : " & Err.Number & " Arquivo de Configuração não encontrado" & vbCrLf & _
'            "Função = Banco_Dados_Externo" & vbCrLf & _
'            "Biblioteca = BLBCWS.clsFuncoes", vbOKOnly + vbCritical, "Aviso"
'  Else
  
'     MsgBox "ATENÇÃO - " & Err.Number & " - " & Err.Description & vbCrLf & _
'            " Função = Banco_Dados_Externo" & vbCrLf & _
'            " Biblioteca = BLBCWS.clsFuncoes", vbOKOnly + vbCritical, "Aviso"
     
'  End If
  
'End Function

Public Function Banco_Dados_Externo(strConn As String) As ADODB.Connection

On Error GoTo err_sys
  
  '"Microsoft.Jet.OLEDB.4.0"
  
  '------------------------------------------------------
  'Abre uma conexção com o banco
 
  Set Banco_Dados_Externo = New ADODB.Connection
  
  Banco_Dados_Externo.ConnectionString = strConn '' Caminho do Banco
  Banco_Dados_Externo.Open
    
  Exit Function
  
err_sys:
  
  If Err.Number = 53 Then
     
     MsgBox "ATENÇÃO - Erro : " & Err.Number & " Arquivo de Configuração não encontrado" & vbCrLf & _
            "Função = Banco_Dados_Externo" & vbCrLf & _
            "Biblioteca = BLBCWS.clsFuncoes", vbOKOnly + vbCritical, "Aviso"
  Else
  
     MsgBox "ATENÇÃO - " & Err.Number & " - " & Err.Description & vbCrLf & _
            " Função = Banco_Dados_Externo" & vbCrLf & _
            " Biblioteca = BLBCWS.clsFuncoes", vbOKOnly + vbCritical, "Aviso"
     
  End If
  
End Function



Public Function CarregaCor(Linha As Variant, Formulario As Variant) As String

    Dim i      As Integer
    Dim strCor As Long
    Dim vControl
    
    For i = 0 To UBound(Linha)
        If UCase(Mid(Linha(i), 1, 3)) = UCase("cor") Then
           strCor = CLng(UCase(Mid(Linha(i), 5, 1000)))
        End If
    Next
    
    If Len(strCor) = 0 Then
       strCor = Formulario.BackColor
    End If

    Formulario.BackColor = strCor
    For Each vControl In Formulario.Controls
        If TypeOf vControl Is TextBox Then
              
        ElseIf TypeOf vControl Is OptionButton Then
              
        ElseIf TypeOf vControl Is ComboBox Then
           
        ElseIf TypeOf vControl Is Label Then
               vControl.BackColor = strCor
        ElseIf TypeOf vControl Is Frame Then
               vControl.BackColor = strCor
        ElseIf TypeOf vControl Is OptionButton Then
               vControl.BackColor = strCor
        ElseIf TypeOf vControl Is CommandButton Then
               vControl.BackColor = strCor
        End If
    Next
    
    CarregaCor = strCor
    
End Function

Public Function Maiuscula(CodAscII As Integer) As Integer
    Maiuscula = asc(UCase(Chr(CodAscII)))
End Function

Public Function ViewCPF(CPF As String) As Boolean
    
    Dim s As Integer
    Dim r As Integer
    Dim i As Integer
    
    s = 0
    For i = 1 To 9
        s = s + Val(Mid$(CPF, i, 1)) * (11 - i)
    Next i
    r = 11 - (s - (Int(s / 11) * 11))
    If r = 10 Or r = 11 Then r = 0
    If r <> Val(Mid$(CPF, 10, 1)) Then
        ViewCPF = False
        Exit Function
    End If
        
    s = 0
    For i = 1 To 10
        s = s + Val(Mid$(CPF, i, 1)) * (12 - i)
    Next i
    r = 11 - (s - (Int(s / 11) * 11))
    If r = 10 Or r = 11 Then r = 0
    If r <> Val(Mid$(CPF, 11, 1)) Then
        ViewCPF = False
        Exit Function
    End If
    
    ViewCPF = True

End Function


Public Function ViewCGC(CGC As String) As Boolean
    
    Dim i As Byte
    Dim c As Byte
    
    If Modulo11(Left(CGC, 12)) <> Mid(CGC, 13, 1) Then
        ViewCGC = False
        Exit Function
    End If
    
    If Modulo11(Left(CGC, 13)) <> Mid(CGC, 14, 1) Then
        ViewCGC = False
        Exit Function
    End If
    
    For i = 1 To 14
        If Mid(CGC, i, 1) = 0 Then
            c = c + 1
        End If
    Next
    If c = 14 Then
        ViewCGC = False
        Exit Function
    End If
    ViewCGC = True
    
End Function

Private Function Modulo11(Number As String) As String
    
    Dim i As Integer
    Dim p As Integer
    Dim M As Integer
    Dim d As Integer
    
    If Not IsNumeric(Number) Then
        Modulo11 = ""
        Exit Function
    End If
    
    M = 2
    For i = Len(Number) To 1 Step -1
        p = p + Val(Mid(Number, i, 1)) * M
        M = IIf(M = 9, 2, M + 1)
    Next
    
    d = 11 - Int(p Mod 11)
    d = IIf(d = 10 Or d = 11, 0, d)
    Modulo11 = Trim(Str(d))
    
End Function

Public Function Preenche_Estado(ByVal Lista As Object)
   
   Dim V_Estado As Variant
   Dim Indice As Integer
   Dim Codigo As Integer
   
   V_Estado = Array("AM", "AC", "AL", "AP", "BA", "CE", "DF", "ES", _
                    "GO", "MA", "MG", "MT", "MS", "PE", "PA", "PB", "PI", "PR", "RJ", _
                    "RN", "RO", "RR", "RS", "SC", "SE", "SP", "TO", "EX")
   Indice = 0
   Codigo = 1
   Lista.Clear
   
   Do While Indice <= UBound(V_Estado)
      Lista.AddItem V_Estado(Indice)
      Lista.ItemData(Lista.NewIndex) = Codigo
      Indice = Indice + 1
      Codigo = Codigo + 1
   Loop
   
End Function

Public Function ComboMagico(CmbMagico As Object, KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Or _
       KeyAscii = vbKeyTab Or _
       KeyAscii = vbKeyEscape Then
       Exit Function
    End If

    Dim sBuffer As String, lRetVal As Long

    With CmbMagico
         sBuffer = Left(.Text, .SelStart) & Chr(KeyAscii)
         lRetVal = SendMessage(.hWnd, CB_FINDSTRING, -1, ByVal sBuffer)
         If lRetVal >= 0 Then
           .ListIndex = lRetVal
           .Text = CmbMagico.List(lRetVal)
           .SelStart = Len(sBuffer)
           .SelLength = Len(.Text)
         Else
           .Text = Chr(KeyAscii)
            lRetVal = SendMessage(.hWnd, CB_FINDSTRING, -1, ByVal sBuffer)
            If lRetVal >= 0 Then
              .ListIndex = lRetVal
              .Text = CmbMagico.List(lRetVal)
              .SelStart = 1
              .SelLength = Len(.Text)
            Else
              .Text = Empty
            End If
         End If
         KeyAscii = Empty
    End With
    
End Function

Public Function Seleciona(Controle As Object, PosInicial As Integer, PosFinal As Integer)
       Controle.SelStart = PosInicial
       Controle.SelLength = PosFinal
End Function

Public Function SoNumeroPonto(KeyAscii As Integer, Texto As String) As Integer

    If KeyAscii = 46 Or KeyAscii = 44 Then
       
       If KeyAscii = 46 Then
          KeyAscii = 44
       End If
       If InStr(1, Texto, ",") Then
          SoNumeroPonto = 0
       Else
          SoNumeroPonto = KeyAscii
       End If
       
       Exit Function
    
    End If

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
       SoNumeroPonto = 0
    Else
       SoNumeroPonto = KeyAscii
    End If
    
End Function

Public Sub ChecaAcesso(frmForm As Variant, strACESSO As String)
  
  Dim i As Integer
  
  Dim vControl

  'Apagando todos os text boxes de um form:
  '----------------------------------------

   For Each vControl In frmForm.Controls
       If TypeOf vControl Is CommandButton Then
          If UCase(vControl.Name) = UCase("cmdInclui") Then
             vControl.Enabled = False
          ElseIf UCase(vControl.Name) = UCase("cmdAltera") Then
             vControl.Enabled = False
          ElseIf UCase(vControl.Name) = UCase("cmdExclui") Then
             vControl.Enabled = False
          End If
       End If
  Next
  
  For i = 1 To Len(Trim(strACESSO))
      For Each vControl In frmForm.Controls
          If TypeOf vControl Is CommandButton Then
             If UCase(vControl.Name) = UCase("cmdInclui") Then
                If Mid(strACESSO, i, 1) = "I" Then vControl.Enabled = True
             ElseIf UCase(vControl.Name) = UCase("cmdAltera") Then
                If Mid(strACESSO, i, 1) = "A" Then vControl.Enabled = True
             ElseIf UCase(vControl.Name) = UCase("cmdExclui") Then
                If Mid(strACESSO, i, 1) = "E" Then vControl.Enabled = True
             End If
          End If
      Next
  Next i

End Sub

Public Function Crypt(Text As String) As String

    Dim i           As Integer
    Dim strTempChar As String
    
    For i = 1 To Len(Text)

        If asc(Mid$(Text, i, 1)) < 128 Then
           strTempChar = asc(Mid$(Text, i, 1)) + 128
        ElseIf asc(Mid$(Text, i, 1)) > 128 Then
           strTempChar = asc(Mid$(Text, i, 1)) - 128
        End If

        Mid$(Text, i, 1) = Chr(strTempChar)

    Next i

    Crypt = Text

End Function


Public Function PegaEstados(intESTADO) As String

On Error GoTo Err_Estado
       
       Dim V_Estado As Variant
       ReDim V_Estado(1 To 28) As String
       
       PegaEstados = ""
       
       V_Estado(1) = "AM"
       V_Estado(2) = "AC"
       V_Estado(3) = "AL"
       V_Estado(4) = "AP"
       V_Estado(5) = "BA"
       V_Estado(6) = "CE"
       V_Estado(7) = "DF"
       V_Estado(8) = "ES"
       V_Estado(9) = "GO"
       V_Estado(10) = "MA"
       V_Estado(11) = "MG"
       V_Estado(12) = "MT"
       V_Estado(13) = "MS"
       V_Estado(14) = "PE"
       V_Estado(15) = "PA"
       V_Estado(16) = "PB"
       V_Estado(17) = "PI"
       V_Estado(18) = "PR"
       V_Estado(19) = "RJ"
       V_Estado(20) = "RN"
       V_Estado(21) = "RO"
       V_Estado(22) = "RR"
       V_Estado(23) = "RS"
       V_Estado(24) = "SC"
       V_Estado(25) = "SE"
       V_Estado(26) = "SP"
       V_Estado(27) = "TO"
       V_Estado(28) = "EX"
                        
       PegaEstados = V_Estado(intESTADO)
       
       Exit Function

Err_Estado:
    PegaEstados = ""

End Function

Public Sub LogSistema(strUSUARIO As String, strOperacao As String, strRegistro As String, strFormulario As String)



End Sub

Public Function Extenso(strValor As String, intMoeda As Integer) As String

    Extenso = ""
    
    If Len(Trim(strValor)) > 14 Then Exit Function

    Dim i              As Integer
    Dim intMilhao      As Integer
    Dim intMil         As Integer
    Dim intCentena     As Integer
    Dim intCentavos    As Integer
    
    Dim strExtMilhao   As String
    Dim strExtMil      As String
    Dim strExtCentenas As String
    Dim strExtCentavos As String
    
    Dim arrMOEDAS      As Variant
    Dim arrUNIDADE     As Variant
    Dim arrDEZENA      As Variant
    Dim arrDEZENA2     As Variant
    Dim arrCENTENA     As Variant

    ReDim arrMOEDAS(1 To 7) As String
    ReDim arrUNIDADE(1 To 9) As String
    ReDim arrDEZENA(1 To 9) As String
    ReDim arrDEZENA2(11 To 19) As String
    ReDim arrCENTENA(1 To 10) As String

    If intMoeda = 1 Then
       arrMOEDAS(1) = "Real"
       arrMOEDAS(2) = "Reais"
       arrMOEDAS(3) = "Centavo"
       arrMOEDAS(4) = "Centavos"
       arrMOEDAS(5) = "Mil"
       arrMOEDAS(6) = "Milhão"
       arrMOEDAS(7) = "Milhões"
    ElseIf intMoeda = 2 Then
       arrMOEDAS(1) = "Dolar"
       arrMOEDAS(2) = "Dolares"
       arrMOEDAS(3) = "Cent"
       arrMOEDAS(4) = "Cents"
       arrMOEDAS(5) = "Mil"
       arrMOEDAS(6) = "Milhão"
       arrMOEDAS(7) = "Milhões"
    End If

    '' Unidades
    arrUNIDADE(1) = "Um"
    arrUNIDADE(2) = "Dois"
    arrUNIDADE(3) = "Três"
    arrUNIDADE(4) = "Quatro"
    arrUNIDADE(5) = "Cinco"
    arrUNIDADE(6) = "Seis"
    arrUNIDADE(7) = "Sete"
    arrUNIDADE(8) = "Oito"
    arrUNIDADE(9) = "Nove"

    '' Dezenas
    arrDEZENA(1) = "Dez"
    arrDEZENA(2) = "Vinte"
    arrDEZENA(3) = "Trinta"
    arrDEZENA(4) = "Quarenta"
    arrDEZENA(5) = "Cinquenta"
    arrDEZENA(6) = "Sessenta"
    arrDEZENA(7) = "Setenta"
    arrDEZENA(8) = "Oitenta"
    arrDEZENA(9) = "Noventa"
    
    arrDEZENA2(11) = "Onze"
    arrDEZENA2(12) = "Doze"
    arrDEZENA2(13) = "Treze"
    arrDEZENA2(14) = "Quatorze"
    arrDEZENA2(15) = "Quinze"
    arrDEZENA2(16) = "Dezesseis"
    arrDEZENA2(17) = "Dezessete"
    arrDEZENA2(18) = "Dezoito"
    arrDEZENA2(19) = "Dezenove"

    '' Centenas
    arrCENTENA(1) = "Cem"
    arrCENTENA(2) = "Duzentos"
    arrCENTENA(3) = "Trezentos"
    arrCENTENA(4) = "Quatrocentos"
    arrCENTENA(5) = "Quiunhentos"
    arrCENTENA(6) = "Seissentos"
    arrCENTENA(7) = "Setecentos"
    arrCENTENA(8) = "Oitocentos"
    arrCENTENA(9) = "Novecentos"
    arrCENTENA(10) = "Cento"
    
    '' -----------------------------------------------------------------------------------
    '' Casa de Milhão
    If Len(Trim(strValor)) >= 12 And Len(Trim(strValor)) <= 14 Then
       If Len(Trim(strValor)) = 12 Then intMilhao = CInt(Mid(strValor, 1, 1))
       If Len(Trim(strValor)) = 13 Then intMilhao = CInt(Mid(strValor, 1, 2))
       If Len(Trim(strValor)) = 14 Then intMilhao = CInt(Mid(strValor, 1, 3))
    End If
    
    If intMilhao >= 1 And intMilhao <= 9 Then
       If intMilhao = 1 Then strExtMilhao = arrUNIDADE(intMilhao) & " " & arrMOEDAS(6)
       If intMilhao > 1 Then strExtMilhao = arrUNIDADE(intMilhao) & " " & arrMOEDAS(7)
    End If
    If intMilhao >= 10 And intMilhao < 20 Then
       If intMilhao = 10 Then strExtMilhao = arrDEZENA(1) & " " & arrMOEDAS(7)
       If intMilhao > 10 And intMilhao <= 19 Then strExtMilhao = arrDEZENA2(intMilhao) & " " & arrMOEDAS(7)
    ElseIf intMilhao >= 20 And intMilhao <= 99 Then
       If Val(Mid(Trim(Str(intMilhao)), 2, 1)) > 0 Then
          strExtMilhao = arrDEZENA(Mid(Trim(Str(intMilhao)), 1, 1)) & " e " & arrUNIDADE(Mid(Trim(Str(intMilhao)), 2, 1)) & " " & arrMOEDAS(7)
       Else
          strExtMilhao = arrDEZENA(Mid(Trim(Str(intMilhao)), 1, 1)) & " " & arrMOEDAS(7)
       End If
    ElseIf intMilhao > 99 And intMilhao <= 999 Then
       If intMilhao = 100 Then strExtMilhao = arrCENTENA(1) & " " & arrMOEDAS(7)
       If intMilhao > 100 And intMilhao <= 999 Then
          If Mid(Trim(Str(intMilhao)), 1, 1) = 1 Then strExtMilhao = arrCENTENA(10)
          If Mid(Trim(Str(intMilhao)), 1, 1) > 1 Then strExtMilhao = arrCENTENA(CInt(Mid(Trim(Str(intMilhao)), 1, 1)))
          If CInt(Mid(Trim(Str(intMilhao)), 2, 2)) >= 1 And CInt(Mid(Trim(Str(intMilhao)), 2, 2)) <= 9 Then
             strExtMilhao = strExtMilhao & " e " & arrUNIDADE(CInt(Mid(Trim(Str(intMilhao)), 2, 2)))
          End If
          If CInt(Mid(Trim(Str(intMilhao)), 2, 2)) >= 10 And CInt(Mid(Trim(Str(intMilhao)), 2, 2)) < 20 Then
             If CInt(Mid(Trim(Str(intMilhao)), 2, 2)) = 10 Then strExtMilhao = strExtMilhao & " e " & arrDEZENA(1) & " " & arrMOEDAS(7)
             If CInt(Mid(Trim(Str(intMilhao)), 2, 2)) > 10 Then strExtMilhao = strExtMilhao & " e " & arrDEZENA2(CInt(Mid(Trim(Str(intMilhao)), 2, 2))) & " " & arrMOEDAS(7)
          ElseIf CInt(Mid(Trim(Str(intMilhao)), 2, 2)) >= 20 And CInt(Mid(Trim(Str(intMilhao)), 2, 2)) <= 99 Then
             If CInt(Mid(Trim(Str(intMilhao)), 3, 1)) > 0 Then
                strExtMilhao = strExtMilhao & " e " & arrDEZENA(Mid(Trim(Str(intMilhao)), 2, 1)) & " e " & arrUNIDADE(Mid(Trim(Str(intMilhao)), 3, 1)) & " " & arrMOEDAS(7)
             Else
                strExtMilhao = strExtMilhao & " e " & arrDEZENA(Mid(Trim(Str(intMilhao)), 2, 1)) & " " & arrMOEDAS(7)
             End If
          Else
             strExtMilhao = strExtMilhao & " " & arrMOEDAS(7)
          End If
       End If
    End If
    '' Fim
    '' -----------------------------------------------------------------------------------
    
    '' -----------------------------------------------------------------------------------
    '' Casa de Mil
    If Len(Trim(strValor)) >= 8 And Len(Trim(strValor)) <= 10 Then
       If Len(Trim(strValor)) = 8 Then intMil = CInt(Mid(strValor, 1, 1))
       If Len(Trim(strValor)) = 9 Then intMil = CInt(Mid(strValor, 1, 2))
       If Len(Trim(strValor)) = 10 Then intMil = CInt(Mid(strValor, 1, 3))
    ElseIf Len(Trim(strValor)) > 10 Then
       For i = 1 To Len(Trim(strValor))
           If Mid(Trim(strValor), i, 1) = "." Then
              intMil = CInt(Mid(strValor, (i + 1), 3))
              Exit For
           End If
       Next
    End If
    
    If intMil >= 1 And intMil <= 9 Then strExtMil = arrUNIDADE(intMil) & " " & arrMOEDAS(5)
    If intMil >= 10 And intMil < 20 Then
       If intMil = 10 Then strExtMil = arrDEZENA(1) & " " & arrMOEDAS(5)
       If intMil > 10 And intMil <= 19 Then strExtMil = arrDEZENA2(intMil) & " " & arrMOEDAS(5)
    ElseIf intMil >= 20 And intMil <= 99 Then
       If Val(Mid(Trim(Str(intMil)), 2, 1)) > 0 Then
          strExtMil = arrDEZENA(Mid(Trim(Str(intMil)), 1, 1)) & " e " & arrUNIDADE(Mid(Trim(Str(intMil)), 2, 1)) & " " & arrMOEDAS(5)
       Else
          strExtMil = arrDEZENA(Mid(Trim(Str(intMil)), 1, 1)) & " " & arrMOEDAS(5)
       End If
    ElseIf intMil > 99 And intMil <= 999 Then
       If intMil = 100 Then strExtMil = arrCENTENA(1) & " " & arrMOEDAS(5)
       If intMil > 100 And intMil <= 999 Then
          If Mid(Trim(Str(intMil)), 1, 1) = 1 Then strExtMil = arrCENTENA(10)
          If Mid(Trim(Str(intMil)), 1, 1) > 1 Then strExtMil = arrCENTENA(CInt(Mid(Trim(Str(intMil)), 1, 1)))
          If CInt(Mid(Trim(Str(intMil)), 2, 2)) >= 1 And CInt(Mid(Trim(Str(intMil)), 2, 2)) <= 9 Then
             strExtMil = strExtMil & " e " & arrUNIDADE(CInt(Mid(Trim(Str(intMil)), 2, 2)))
          End If
          If CInt(Mid(Trim(Str(intMil)), 2, 2)) >= 10 And CInt(Mid(Trim(Str(intMil)), 2, 2)) < 20 Then
             If CInt(Mid(Trim(Str(intMil)), 2, 2)) = 10 Then strExtMil = strExtMil & " e " & arrDEZENA(1) & " " & arrMOEDAS(5)
             If CInt(Mid(Trim(Str(intMil)), 2, 2)) > 10 Then strExtMil = strExtMil & " e " & arrDEZENA2(CInt(Mid(Trim(Str(intMil)), 2, 2))) & " " & arrMOEDAS(5)
          ElseIf CInt(Mid(Trim(Str(intMil)), 2, 2)) >= 20 And CInt(Mid(Trim(Str(intMil)), 2, 2)) <= 99 Then
             If CInt(Mid(Trim(Str(intMil)), 3, 1)) > 0 Then
                strExtMil = strExtMil & " e " & arrDEZENA(Mid(Trim(Str(intMil)), 2, 1)) & " e " & arrUNIDADE(Mid(Trim(Str(intMil)), 3, 1)) & " " & arrMOEDAS(5)
             Else
                strExtMil = strExtMil & " e " & arrDEZENA(Mid(Trim(Str(intMil)), 2, 1)) & " " & arrMOEDAS(5)
             End If
          Else
             strExtMil = strExtMil & " " & arrMOEDAS(5)
          End If
       End If
    End If
    ''Fim
    '' -----------------------------------------------------------------------------------
    
    '' -----------------------------------------------------------------------------------
    '' Centena
    If Len(Trim(strValor)) >= 4 And Len(Trim(strValor)) <= 6 Then
       If Len(Trim(strValor)) = 4 Then intCentena = CInt(Mid(strValor, 1, 1))
       If Len(Trim(strValor)) = 5 Then intCentena = CInt(Mid(strValor, 1, 2))
       If Len(Trim(strValor)) = 6 Then intCentena = CInt(Mid(strValor, 1, 3))
    ElseIf Len(Trim(strValor)) > 6 Then
       For i = 1 To Len(Trim(strValor))
           If Mid(Trim(strValor), i, 1) = "." Then intCentena = CInt(Mid(strValor, (i + 1), 3))
       Next
    End If
    
    If intCentena >= 1 And intCentena <= 9 Then
       If intCentena = 1 Then strExtCentenas = arrUNIDADE(intCentena) & " " & arrMOEDAS(1)
       If intCentena > 1 And intCentena <= 9 Then strExtCentenas = arrUNIDADE(intCentena) & " " & arrMOEDAS(2)
    ElseIf intCentena >= 10 And intCentena < 20 Then
       If intCentena = 10 Then strExtCentenas = arrDEZENA(1) & " " & arrMOEDAS(2)
       If intCentena > 10 And intCentena <= 19 Then strExtCentenas = arrDEZENA2(intCentena) & " " & arrMOEDAS(2)
    ElseIf intCentena >= 20 And intCentena <= 99 Then
       If Val(Mid(Trim(Str(intCentena)), 2, 1)) > 0 Then
          strExtCentenas = arrDEZENA(Mid(Trim(Str(intCentena)), 1, 1)) & " e " & arrUNIDADE(Mid(Trim(Str(intCentena)), 2, 1)) & " " & arrMOEDAS(2)
       Else
          strExtCentenas = arrDEZENA(Mid(Trim(Str(intCentena)), 1, 1)) & " " & arrMOEDAS(2)
       End If
    ElseIf intCentena > 99 And intCentena <= 999 Then
       If intCentena = 100 Then strExtCentenas = arrCENTENA(1) & " " & arrMOEDAS(2)
       If intCentena > 100 And intCentena <= 999 Then
          If Mid(Trim(Str(intCentena)), 1, 1) = 1 Then strExtCentenas = arrCENTENA(10)
          If Mid(Trim(Str(intCentena)), 1, 1) > 1 Then strExtCentenas = arrCENTENA(CInt(Mid(Trim(Str(intCentena)), 1, 1)))
          If CInt(Mid(Trim(Str(intCentena)), 2, 2)) >= 1 And CInt(Mid(Trim(Str(intCentena)), 2, 2)) <= 9 Then
             strExtCentenas = strExtCentenas & " e " & arrUNIDADE(CInt(Mid(Trim(Str(intCentena)), 2, 2)))
          End If
          If CInt(Mid(Trim(Str(intCentena)), 2, 2)) >= 10 And CInt(Mid(Trim(Str(intCentena)), 2, 2)) < 20 Then
             If CInt(Mid(Trim(Str(intCentena)), 2, 2)) = 10 Then strExtCentenas = strExtCentenas & " e " & arrDEZENA(1) & " " & arrMOEDAS(2)
             If CInt(Mid(Trim(Str(intCentena)), 2, 2)) > 10 Then strExtCentenas = strExtCentenas & " e " & arrDEZENA2(CInt(Mid(Trim(Str(intCentena)), 2, 2))) & " " & arrMOEDAS(2)
          ElseIf CInt(Mid(Trim(Str(intCentena)), 2, 2)) >= 20 And CInt(Mid(Trim(Str(intCentena)), 2, 2)) <= 99 Then
             If CInt(Mid(Trim(Str(intCentena)), 3, 1)) > 0 Then
                strExtCentenas = strExtCentenas & " e " & arrDEZENA(Mid(Trim(Str(intCentena)), 2, 1)) & " e " & arrUNIDADE(Mid(Trim(Str(intCentena)), 3, 1)) & " " & arrMOEDAS(2)
             Else
                strExtCentenas = strExtCentenas & " e " & arrDEZENA(Mid(Trim(Str(intCentena)), 2, 1)) & " " & arrMOEDAS(2)
             End If
          Else
             strExtCentenas = strExtCentenas & " " & arrMOEDAS(2)
          End If
       End If
    End If
    '' Fim
    '' -----------------------------------------------------------------------------------
    
    '' -----------------------------------------------------------------------------------
    '' Centavos
    intCentavos = CInt(Right(strValor, 2))
    If intCentavos >= 1 And intCentavos <= 9 Then
       If intCentavos = 1 Then strExtCentavos = arrUNIDADE(intCentavos) & " " & arrMOEDAS(3)
       If intCentavos >= 2 Then strExtCentavos = arrUNIDADE(intCentavos) & " " & arrMOEDAS(4)
    End If
        
    If intCentavos >= 10 And intCentavos < 20 Then
       If intCentavos = 10 Then
          strExtCentavos = arrDEZENA(1) & " " & arrMOEDAS(4)
       ElseIf intCentavos > 10 And intCentavos <= 19 Then
          strExtCentavos = arrDEZENA2(intCentavos) & " " & arrMOEDAS(4)
       End If
    ElseIf intCentavos >= 20 And intCentavos <= 99 Then
       If Val(Mid(Trim(Str(intCentavos)), 2, 1)) > 0 Then
          strExtCentavos = arrDEZENA(Mid(Trim(Str(intCentavos)), 1, 1)) & " e " & arrUNIDADE(Mid(Trim(Str(intCentavos)), 2, 1)) & " " & arrMOEDAS(4)
       Else
          strExtCentavos = arrDEZENA(Mid(Trim(Str(intCentavos)), 1, 1)) & " " & arrMOEDAS(4)
       End If
    End If
    '' Fim
    '' -----------------------------------------------------------------------------------
    
    If Len(Trim(strExtMilhao)) > 0 And Len(Trim(strExtMil)) = 0 And Len(Trim(strExtCentenas)) = 0 And Len(Trim(strExtCentavos)) = 0 Then Extenso = strExtMilhao & " de " & arrMOEDAS(2)
    If Len(Trim(strExtMilhao)) > 0 And Len(Trim(strExtMil)) > 0 And Len(Trim(strExtCentenas)) > 0 And Len(Trim(strExtCentavos)) > 0 Then Extenso = strExtMilhao & " e " & Trim(strExtMil) & " e " & Trim(strExtCentenas) & " e " & strExtCentavos
    If Len(Trim(strExtMilhao)) > 0 And Len(Trim(strExtMil)) > 0 And Len(Trim(strExtCentenas)) > 0 And Len(Trim(strExtCentavos)) = 0 Then Extenso = strExtMilhao & " e " & Trim(strExtMil) & " e " & Trim(strExtCentenas)
    If Len(Trim(strExtMilhao)) > 0 And Len(Trim(strExtMil)) > 0 And Len(Trim(strExtCentenas)) = 0 And Len(Trim(strExtCentavos)) = 0 Then Extenso = strExtMilhao & " e " & Trim(strExtMil) & " " & arrMOEDAS(2)
    If Len(Trim(strExtMilhao)) > 0 And Len(Trim(strExtMil)) = 0 And Len(Trim(strExtCentenas)) > 0 And Len(Trim(strExtCentavos)) > 0 Then Extenso = strExtMilhao & " e " & Trim(strExtCentenas) & " e " & strExtCentavos
    If Len(Trim(strExtMilhao)) > 0 And Len(Trim(strExtMil)) = 0 And Len(Trim(strExtCentenas)) = 0 And Len(Trim(strExtCentavos)) > 0 Then Extenso = strExtMilhao & " de " & arrMOEDAS(2) & " e " & strExtCentavos
    If Len(Trim(strExtMilhao)) > 0 And Len(Trim(strExtMil)) = 0 And Len(Trim(strExtCentenas)) > 0 And Len(Trim(strExtCentavos)) = 0 Then Extenso = strExtMilhao & " e " & Trim(strExtCentenas)
    If Len(Trim(strExtMilhao)) > 0 And Len(Trim(strExtMil)) > 0 And Len(Trim(strExtCentenas)) = 0 And Len(Trim(strExtCentavos)) > 0 Then Extenso = strExtMilhao & " e " & Trim(strExtMil) & " " & arrMOEDAS(2) & " e " & strExtCentavos
    
    If Len(Trim(strExtMilhao)) = 0 And Len(Trim(strExtMil)) > 0 And Len(Trim(strExtCentavos)) > 0 And Len(Trim(strExtCentenas)) = 0 Then Extenso = strExtMil & " " & arrMOEDAS(2) & " e " & Trim(strExtCentavos)
    If Len(Trim(strExtMilhao)) = 0 And Len(Trim(strExtMil)) > 0 And Len(Trim(strExtCentavos)) = 0 And Len(Trim(strExtCentenas)) = 0 Then Extenso = strExtMil & " " & arrMOEDAS(2)
    If Len(Trim(strExtMilhao)) = 0 And Len(Trim(strExtMil)) > 0 And Len(Trim(strExtCentavos)) > 0 And Len(Trim(strExtCentenas)) > 0 Then Extenso = strExtMil & " e " & strExtCentenas & " e " & strExtCentavos
    If Len(Trim(strExtMilhao)) = 0 And Len(Trim(strExtMil)) = 0 And Len(Trim(strExtCentavos)) > 0 And Len(Trim(strExtCentenas)) > 0 Then Extenso = strExtCentenas & " e " & strExtCentavos
    If Len(Trim(strExtMilhao)) = 0 And Len(Trim(strExtMil)) = 0 And Len(Trim(strExtCentenas)) = 0 And Len(Trim(strExtCentavos)) > 0 Then Extenso = strExtCentavos
    If Len(Trim(strExtMilhao)) = 0 And Len(Trim(strExtMil)) > 0 And Len(Trim(strExtCentenas)) > 0 And Len(Trim(strExtCentavos)) = 0 Then Extenso = strExtMil & " e " & strExtCentenas
    If Len(Trim(strExtMilhao)) = 0 And Len(Trim(strExtMil)) = 0 And Len(Trim(strExtCentavos)) = 0 And Len(Trim(strExtCentenas)) > 0 Then Extenso = Trim(strExtCentenas)
    
    Extenso = "(" & Extenso & ".)"
    
End Function

Public Function PegaDescCombo(Combo As Object, intCODIGO As Integer) As String

    Dim i As Integer
    
    PegaDescCombo = ""
    
    For i = 0 To (Combo.ListCount - 1)
        If Combo.ItemData(i) = intCODIGO Then
           PegaDescCombo = Combo.List(i)
           Exit For
        End If
    Next i

End Function

Public Function PegaListIndexCombo(Combo As Object, intCODIGO As Integer) As Integer

    Dim i As Integer
    
    PegaListIndexCombo = -1
    
    For i = 0 To (Combo.ListCount - 1)
        If Combo.ItemData(i) = intCODIGO Then
           PegaListIndexCombo = i
           Exit For
        End If
    Next i

End Function


Public Function ChecaAcesso2(strTIPOPERACAO As String, strACESSO As String, Optional boolExibeMens As Boolean) As Boolean

  Dim i As Integer
  
  ChecaAcesso2 = False
  
  For i = 1 To Len(Trim(strACESSO))
      If Mid(strACESSO, i, 1) = strTIPOPERACAO Then ChecaAcesso2 = True
  Next i
  
  If boolExibeMens = True Then Exit Function
  
  If ChecaAcesso2 = False And strTIPOPERACAO = "I" Then
     MsgBox "Você não tem permissão para incluir !!!", vbOKOnly + vbExclamation, "Aviso"
  End If
  If ChecaAcesso2 = False And strTIPOPERACAO = "A" Then
     MsgBox "Você não tem permissão para alterar !!!", vbOKOnly + vbExclamation, "Aviso"
  End If
  If ChecaAcesso2 = False And strTIPOPERACAO = "E" Then
     MsgBox "Você não tem permissão para excluir !!!", vbOKOnly + vbExclamation, "Aviso"
  End If
  If ChecaAcesso2 = False And strTIPOPERACAO = "C" Then
     MsgBox "Você não tem permissão para consultar !!!", vbOKOnly + vbExclamation, "Aviso"
  End If
  If ChecaAcesso2 = False And strTIPOPERACAO = "R" Then
     MsgBox "Você não tem permissão para tirar relatórios !!!", vbOKOnly + vbExclamation, "Aviso"
  End If
  If ChecaAcesso2 = False And strTIPOPERACAO = "P" Then
     MsgBox "Você não tem permissão para imprimir !!!", vbOKOnly + vbExclamation, "Aviso"
  End If
  If ChecaAcesso2 = False And strTIPOPERACAO = "L" Then
     MsgBox "Você não tem permissão para liberar !!!", vbOKOnly + vbExclamation, "Aviso"
  End If
  If ChecaAcesso2 = False And strTIPOPERACAO = "B" Then
     MsgBox "Você não tem permissão para bloquear !!!", vbOKOnly + vbExclamation, "Aviso"
  End If
  If ChecaAcesso2 = False And strTIPOPERACAO = "V" Then
     MsgBox "Você não tem permissão para reprovar !!!", vbOKOnly + vbExclamation, "Aviso"
  End If

End Function

''Public Sub Sub_DescErro(strErroNumber As String, strErrDesc As String, strOperacao As String, strErro As String)

''    If strOperacao = "I" Then MsgBox "Foi impossivel incluir o registro !!!" & vbCrLf & _
''                                 "Erro : " & strErroNumber & vbCrLf & _
''                                 "Desc : " & strErrDesc & vbCrLf & _
''                                 "" & vbCrLf & _
''                                 strErro, vbOKOnly + vbCritical, "Aviso Sistema"
    
''    If strOperacao = "A" Then MsgBox "Foi impossivel altera o registro !!!" & vbCrLf & _
''                                 "Erro : " & strErroNumber & vbCrLf & _
''                                 "Desc : " & strErrDesc & vbCrLf & _
''                                 "" & vbCrLf & _
''                                 strErro, vbOKOnly + vbCritical, "Aviso Sistema"
''
''    If strOperacao = "E" Then MsgBox "Foi impossivel excluir o registro !!!" & vbCrLf & _
''                                 "Erro : " & strErroNumber & vbCrLf & _
''                                 "Desc : " & strErrDesc & vbCrLf & _
''                                 "" & vbCrLf & _
''                                 strErro, vbOKOnly + vbCritical, "Aviso Sistema"

''    If strOperacao = "AL" Then MsgBox "Foi impossivel gerar a estrutura !!!" & vbCrLf & _
''                                 "Erro : " & strErroNumber & vbCrLf & _
''                                "Desc : " & strErrDesc & vbCrLf & _
''                                 "" & vbCrLf & _
''                                 strErro, vbOKOnly + vbCritical, "Aviso Sistema"

''End Sub

'Public Function WriteBlobToDB(rsDoc As ADODB.Recordset, fld As String, fname As String)

'Dim lngOffset As Long
'Dim nHandle
'Dim lngNumreroItapas As Long
'Dim lSize As Long
'Dim I As Integer
'
'    If Trim(fname) = "" Then ' Caso caminho da imagem seja nulo limpa imagem do banco
'        rsDoc.Fields(fld) = Null
'        Exit Function
'    End If
'
'    nHandle = FreeFile
'
'    Open fname For Binary As nHandle
'    lSize = LOF(nHandle)
'    If nHandle = 0 Then
'        Close nHandle
'    End If
'    lngNumreroItapas = lSize \ conChunkSize
'    lngOffset = lSize Mod conChunkSize
'    If lngOffset > 0 Then
'        ReDim chunk(lngOffset - 1)
'        Get nHandle, , chunk()
'        rsDoc.Fields(fld).AppendChunk chunk()
'    End If
'    ReDim chunk(conChunkSize - 1)
'
'    For I = 1 To lngNumreroItapas
'        Get nHandle, , chunk()
'        rsDoc.Fields(fld).AppendChunk chunk()
'    Next
'
'    Close nHandle
'End Function

Public Function GravaBlobParaBanco(rsDoc As ADODB.Recordset, fld As String, fname As String) As Variant

    Dim lngOffset           As Long
    Dim nHandle
    Dim lngNumreroItapas    As Long
    Dim lSize               As Long
    Dim i                   As Integer
    Dim btConteudoArq       As Byte

    GravaBlobParaBanco = Empty
    
    If Trim(fname) = "" Then ' Caso caminho da imagem seja nulo limpa imagem do banco
       rsDoc.Fields(fld) = Null
       Exit Function
    End If
    
    nHandle = FreeFile
    Open fname For Binary As nHandle
    lSize = LOF(nHandle)
    If nHandle = 0 Then
       Close nHandle
       Exit Function
    End If

    lngNumreroItapas = lSize \ conChunkSize
    lngOffset = lSize Mod conChunkSize
    If lngOffset > 0 Then
       ReDim chunk(lngOffset - 1)
       Get nHandle, , chunk()
       rsDoc.Fields(fld).AppendChunk chunk()
    End If
    ReDim chunk(conChunkSize - 1)

    For i = 1 To lngNumreroItapas
        Get nHandle, , chunk()
        rsDoc.Fields(fld).AppendChunk chunk()
        GravaBlobParaBanco = rsDoc.Fields(fld)
    Next
    Close nHandle
    
End Function


Public Function LeCampoBlobDoDB(rsDoc As ADODB.Recordset, fld As String, fname As String) As String
    
    Dim i As Integer
    Dim intNumreroItapas As Integer
    Dim fileno As Integer
    Dim lngDocsize As Long
    Dim lngOffset As Long
    Dim strfname As String
    Dim strDir As String
    
    
    ''strDir = cArquivos ' Temporary directory - CHUMBADO

    strfname = strDir & fname
    If rsDoc.EOF() Then
        LeCampoBlobDoDB = ""
        Exit Function
    End If
    
    lngDocsize = rsDoc.Fields(fld).ActualSize
    intNumreroItapas = lngDocsize \ conChunkSize
    lngOffset = lngDocsize Mod conChunkSize
    
    
    fileno = FreeFile
    Open strfname For Binary As fileno
    
    If intNumreroItapas > 0 Then
        ReDim chunk(conChunkSize - 1)
        For i = 1 To intNumreroItapas
            chunk = rsDoc.Fields(fld).GetChunk(conChunkSize)
            Put #fileno, , chunk
        Next
    End If
    
    If lngOffset > 0 Then
        ReDim chunk(lngOffset - 1)
        chunk = rsDoc.Fields(fld).GetChunk(lngOffset)
        Put #fileno, , chunk
    End If
    Close fileno
    LeCampoBlobDoDB = strfname
End Function

Public Sub PreenchComboStatus(Combo As Variant)

    Combo.Clear
    
    Combo.AddItem "Cadastrado"
    Combo.ItemData(Combo.NewIndex) = 0
    
    Combo.AddItem "Em Análise"
    Combo.ItemData(Combo.NewIndex) = 1
    
    Combo.AddItem "Liberado"
    Combo.ItemData(Combo.NewIndex) = 2

    Combo.AddItem "Reprovado"
    Combo.ItemData(Combo.NewIndex) = 3

End Sub

Public Sub PreenchComboClassificacao(Combo As Variant)

    Combo.Clear
    
    Combo.AddItem "Ótimo"
    Combo.ItemData(Combo.NewIndex) = 0
    
    Combo.AddItem "Bom"
    Combo.ItemData(Combo.NewIndex) = 1
    
    Combo.AddItem "Médio"
    Combo.ItemData(Combo.NewIndex) = 2

    Combo.AddItem "Ruim"
    Combo.ItemData(Combo.NewIndex) = 3

End Sub

Public Sub ExclLinhaGrid(grdGenerica As Variant, intRow As Integer, Optional intIndice As Integer)
    If (grdGenerica.Rows - 1) = 1 Then grdGenerica.Rows = 1
    If (grdGenerica.Rows - 1) > 1 Then grdGenerica.RemoveItem intRow
End Sub

Public Function FormataCnpj(strCNPJ As String)
       FormataCnpj = strCNPJ
       If Len(Trim(FormataCnpj)) >= 14 Then
          FormataCnpj = Mid(strCNPJ, 1, 2) & "." & Mid(strCNPJ, 3, 3) & "." & Mid(strCNPJ, 6, 3) & "/" & Mid(strCNPJ, 9, 4) & "-" & Mid(strCNPJ, 13, 2)
       End If
End Function

Public Function TamanhCelula(strColuna As String, lngTmCol As Long) As Boolean
    TamanhCelula = True
    If Len(Trim(strColuna)) > lngTmCol Then
       MsgBox "Somente é Permitido " & Str(lngTmCol) & " Digito(s) !!!", vbOKOnly + vbExclamation, "Aviso"
       TamanhCelula = False
    End If
End Function

Public Sub RemoveLinhaGrid(grdSon As Variant, intRow As Integer)
    If grdSon.Row = 0 Then Exit Sub
    If (grdSon.Rows - 1) = 1 Then grdSon.Rows = 1
    If (grdSon.Rows - 1) > 1 Then grdSon.RemoveItem intRow
End Sub

Public Function MaskNumber(ByVal strText As String, ByVal asc As Integer, ByVal intDecimals As Integer, ByVal Tipo As Long, Optional maxValue As Variant, Optional boolAllowNegative As Boolean = False) As Integer
Dim i As Double
Dim isNumber As Boolean
Dim s As String
Dim j As Long
    
    If (asc = 27) Or (asc = 13) Or (asc = 8) Then 'Esc or Enter or Backspace
      MaskNumber = asc
      Exit Function
    End If
    
    If Tipo = myvarAsSmallDate Or Tipo = myvarAsDate Then
        If Not (asc = 47 Or (asc >= 48 And asc <= 57)) Then ' equal to  / or 1..9
            MaskNumber = 0
            Exit Function
        End If
        
        Dim strParts() As String
        strParts = Split(strText & Chr(asc), "/")
        If UBound(strParts) > 2 Then ' More than 2 / exists
            MaskNumber = 0
            Exit Function
        Else
            If UBound(strParts) = 0 Then ' Only days digited
                If Not IsNumeric(strParts(0)) Or Len(strParts(0)) > 2 Then
                    MaskNumber = 0
                    Exit Function
                ElseIf strParts(0) > 31 Then
                    MaskNumber = 0
                    Exit Function
                End If
            ElseIf UBound(strParts) = 1 Then ' Only days and months digited
                If Not IsNumeric(strParts(0)) Or Len(strParts(0)) > 2 Then
                    MaskNumber = 0
                    Exit Function
                ElseIf strParts(0) > 31 Then
                    MaskNumber = 0
                    Exit Function
                ElseIf IsNumeric(strParts(1)) Then
                    If strParts(1) > 12 Then
                        MaskNumber = 0
                        Exit Function
                    End If
                    Select Case CInt(strParts(1))
                    Case 1, 3, 5, 7, 8, 10, 12
                        If strParts(0) > 31 Then
                            MaskNumber = 0
                            Exit Function
                        End If
                    Case 2
                        If strParts(0) > 29 Then
                            MaskNumber = 0
                            Exit Function
                        End If
                    Case Else
                        If strParts(0) > 30 Then
                            MaskNumber = 0
                            Exit Function
                        End If
                    End Select
                ElseIf strParts(1) <> "" Then
                    MaskNumber = 0
                    Exit Function
                End If
            Else ' All are digited
                If Not IsNumeric(strParts(0)) Or Len(strParts(0)) > 2 Then
                    MaskNumber = 0
                    Exit Function
                ElseIf strParts(0) > 31 Then
                    MaskNumber = 0
                    Exit Function
                ElseIf Not IsNumeric(strParts(1)) Or Len(strParts(1)) > 2 Then
                    MaskNumber = 0
                    Exit Function
                ElseIf CLng(strParts(1)) > 12 Then
                    MaskNumber = 0
                    Exit Function
                ElseIf IsNumeric(strParts(2)) Then
                    If strParts(2) > 2078 Then
                        MaskNumber = 0
                        Exit Function
                    End If
                    Select Case CInt(strParts(1))
                    Case 1, 3, 5, 7, 8, 10, 12
                        If strParts(0) > 31 Then
                            MaskNumber = 0
                            Exit Function
                        End If
                    Case 2
                        If strParts(0) > 29 Then
                            MaskNumber = 0
                            Exit Function
                        End If
                    Case Else
                        If strParts(0) > 30 Then
                            MaskNumber = 0
                            Exit Function
                        End If
                    End Select
                ElseIf strParts(2) <> "" Then
                        MaskNumber = 0
                        Exit Function
                End If
            End If
        End If
        MaskNumber = asc
        Exit Function
    End If
    
    

    If asc = 45 And Not boolAllowNegative Then  ' -
       MaskNumber = 0
       Exit Function
    End If
    
    If Tipo = myvarAsString Then
       If (Len(strText) - 1) >= intDecimals Then
           MaskNumber = 0
       Else
           MaskNumber = asc
       End If
       Exit Function
       
    End If
    
     
    If ((asc = 43) Or (asc = 45)) And (Len(Trim(strText)) > 0) Then '+ OR - SIGN
     MaskNumber = 0
     Exit Function
    End If
    If asc = 44 Then ' asc=44 ","
       If CSng("1,00") > 1 Then ' i.e "," is thousand separator in Locale
           asc = 46
       End If
    End If
    
    If asc = 46 Then 'asc=46 "."
       If CSng("1.00") > 1 Then ' i.e "." is thousand separator in Locale
           asc = 44
       End If
    End If

 If IsMissing(maxValue) Then
    Select Case Tipo
      Case myvarAsByte
           maxValue = 255
      Case myvarAsInteger
           maxValue = 32767
      Case myvarAsLong
           maxValue = 2147483647
      Case myvarAsDouble
           maxValue = 1.79769313486231E+308
      Case myvarAsCurrency
           maxValue = 922337203685477#
      Case myvarAsDate
           maxValue = #12/31/9999#
      Case myvarAsSmallDate
           maxValue = #6/6/2079#
      Case myvarAsSingle
           maxValue = 3.402823E+38
    End Select
 End If
  
 MaskNumber = asc
 ' backspace or Enter
 If (asc = 8) Or (asc = 13) Then
    Exit Function
 End If
 
 strText = strText & Chr(asc)
 If (Trim(strText) = "-") Or (Trim(strText) = "+") Then
    MaskNumber = asc
    Exit Function
 End If
 
' If IsNumeric(strText) Then
'    If Len(strText) >= 2 Then
        'If (Mid(strText, 1, 1) <> "+" And Mid(strText, 1, 1) <> "-") And Mid(strText, 1, 1) = "0" And (Mid(strText, 2, 1) <> "." Or Mid(strText, 2, 1) <> ",") Then
         '   MaskNumber = 0
            'Exit Function
        'End If
    'End If
    'If Len(strText) >= 3 Then
        'If (Mid(strText, 1, 1) = "+" Or Mid(strText, 1, 1) = "-") And _
            'Mid(strText, 2, 1) = "0" And _
            '(Mid(strText, 3, 1) <> "." Or Mid(strText, 3, 1) <> ",") Then
            'MaskNumber = 0
            'Exit Function
        'End If
    'End If
 'End If
 
 If IsNumeric(strText) Then
   isNumber = True
   i = CDbl(strText)
 Else
   isNumber = False
   i = 0
 End If

 j = InStr(strText, ".")
 If j <> 0 Then
   s = Mid(strText, j) 's=decimal part
   j = Len(s) - 1 ' j=Len of decimal part part
 Else
   s = ""
   j = 0
 End If
 
 Select Case Tipo
 Case myvarAsLong
     If isNumber Then
       If (i > CDbl(maxValue)) Then
          MaskNumber = 0
        ElseIf (i - Fix(i)) <> 0 Then
           MaskNumber = 0
        ElseIf asc = vbKeyDecimal Then
           MaskNumber = 0
       End If
     Else
        MaskNumber = 0
     End If
 Case myvarAsInteger
     If isNumber Then
        If (i > CDbl(maxValue)) Then
          MaskNumber = 0
        ElseIf (i - Fix(i)) <> 0 Then
           MaskNumber = 0
        ElseIf asc = vbKeyDecimal Then
           MaskNumber = 0
        End If
     Else
        MaskNumber = 0
     End If
 Case myvarAsDouble
     If isNumber Then
        If (i > CDbl(maxValue)) Then
           MaskNumber = 0
        ElseIf j > intDecimals Then
           MaskNumber = 0
        End If
     Else
        MaskNumber = 0
     End If
     
 Case myvarAsSingle
     If isNumber Then
        If (i > CSng(maxValue)) Then
           MaskNumber = 0
        ElseIf j > intDecimals Then
           MaskNumber = 0
        End If
     Else
        MaskNumber = 0
     End If
 Case myvarAsCurrency
     If isNumber Then
        If (i > CDbl(maxValue)) Then
          MaskNumber = 0
        ElseIf j > intDecimals Then
          MaskNumber = 0
        End If
     Else
        MaskNumber = 0
     End If
     
  Case myvarAsByte
     If isNumber Then
        If (i > CDbl(maxValue)) Then
          MaskNumber = 0
        ElseIf (i - Fix(i)) <> 0 Then
           MaskNumber = 0
        End If
     ElseIf asc = vbKeyDecimal Then
           MaskNumber = 0
     Else
        MaskNumber = 0
     End If
     
  End Select
  
End Function

Public Function Date2Str(dtValue As String) As String
    If IsDate(dtValue) Then
        If Trim(dtValue) = "00:00:00" Then
            Date2Str = ""
        Else
            Date2Str = dtValue
        End If
    Else
        Date2Str = ""
    End If
End Function

Public Function Currency2Str(sngValue As Currency, Optional DecimalPoint As Integer = 2) As String
Dim i As Byte
Dim s As String
    If sngValue = -1 Then
        Currency2Str = ""
    Else
        For i = 1 To DecimalPoint
            s = s & "0"
        Next
        If s <> "" Then
            s = "0." & s
        Else
            s = "0"
        End If
            
        Currency2Str = Format(sngValue, s)
    End If
End Function

Public Function Double2Str(sngValue As Double, Optional DecimalPoint As Integer = 2) As String
Dim i As Byte
Dim s As String
    If sngValue = -1 Then
        Double2Str = ""
    Else
        For i = 1 To DecimalPoint
            s = s & "0"
        Next
        If s <> "" Then
            s = "0." & s
        Else
            s = "0"
        End If
            
        Double2Str = Format(sngValue, s)
    End If
End Function

Public Function Long2Str(lngValue As Long) As String
    If lngValue = -1 Then
        Long2Str = ""
    Else
        Long2Str = lngValue
    End If
End Function

Public Function Integer2Str(intValue As Integer) As String
    If intValue = -1 Then
        Integer2Str = ""
    Else
        Integer2Str = intValue
    End If
End Function

Public Function Byte2Str(bytValue As Byte) As String
    If bytValue = 255 Then
        Byte2Str = ""
    Else
        Byte2Str = bytValue
    End If
End Function
Public Function Bool2Str(boolValue As String) As Boolean
    If Trim(boolValue) = "-1" Or Trim(boolValue) = "1" Then
        Bool2Str = True
    Else
        Bool2Str = False
    End If
End Function

Public Sub RemoveLinhaVazia(grdGRID As Variant, lngCOL As Long)
    Dim i As Integer
VOLTA:
    For i = 1 To (grdGRID.Rows - 1)
        If grdGRID.Cell(flexcpText, i, lngCOL) = Empty Then
           grdGRID.RemoveItem i
           GoTo VOLTA
        End If
    Next i
End Sub

Public Function CalcTempo(strHORINI As String, strHORFIN As String) As String
    
    CalcTempo = ""
    
    Dim dtTotalHora         As Date
    Dim lngMinutosIni       As Long
    Dim lngMinutosFin       As Long
    Dim lngTotMinutos       As Long
    
    If Trim(strHORINI) <> ":" And Trim(strHORFIN) <> ":" Then
        If Len(Trim(strHORINI)) > 0 And Len(Trim(strHORFIN)) > 0 Then
           If CDate(Replace(strHORINI, ";", ":")) > CDate(Replace(strHORFIN, ";", ":")) Then
              
              lngMinutosIni = CONVHRMIN(strHORINI)
              lngMinutosFin = CONVHRMIN("23:59")
              lngTotMinutos = ((lngMinutosFin - lngMinutosIni) + 1)
              
              lngMinutosFin = CONVHRMIN(strHORFIN)
              lngTotMinutos = (lngTotMinutos + lngMinutosFin)
              
           Else
              
              lngMinutosIni = CONVHRMIN(strHORINI)
              lngMinutosFin = CONVHRMIN(strHORFIN)
              lngTotMinutos = (lngMinutosFin - lngMinutosIni)
           
           End If
        End If
    End If
    
    CalcTempo = CONVMINHR(lngTotMinutos)
    
    If Len(Trim(CalcTempo)) = 0 Then CalcTempo = "00:00"
    
End Function

Public Function CONVHRMIN(strHORA As String) As Long
    
    CONVHRMIN = 0
    
    If Len(Trim(strHORA)) = 0 Then Exit Function
    
    Dim HORAS       As Long
    Dim MINUTOS     As Long
    Dim TOTMINUTOS  As Long
    
    HORAS = Hour(CDate(Replace(strHORA, ";", ":")))
    MINUTOS = Minute(CDate(Replace(strHORA, ";", ":")))
    TOTMINUTOS = ((HORAS * 60) + MINUTOS)
    
    CONVHRMIN = TOTMINUTOS

End Function

Public Function CONVMINHR(lngMinutos As Long) As String
    
    CONVMINHR = ""
    
    If lngMinutos = 0 Then Exit Function
    
    Dim TOTMINUTOS  As Double
    Dim HORA        As Long
    Dim MINUTO      As Long
    Dim strHORAS    As String
    Dim arrHRMN()   As String
    
    TOTMINUTOS = Round((lngMinutos / 60), 2)
    strHORAS = Format(TOTMINUTOS, "###,##000.00")
    arrHRMN = Split(strHORAS, ",")
    
    HORA = CLng(arrHRMN(0))
    MINUTO = (CLng(arrHRMN(1)) * (0.6))
    
    CONVMINHR = Trim(Format(HORA, "##00") & ":" & Format(MINUTO, "##00") & ":" & "00")

End Function


Public Function FcExisteLinhaVazia(grdGenerica As Variant, Col As Long) As Boolean
    FcExisteLinhaVazia = False
    Dim i As Integer
    For i = 1 To (grdGenerica.Rows - 1)
        If grdGenerica.Cell(flexcpText, i, Col) = Empty Then Exit Function
    Next i
    FcExisteLinhaVazia = True
End Function

Public Function FcVerifItensRepetidos(grdGenerica As Variant, intRow As Long, intCol As Long, varCampo As Variant) As Boolean
    FcVerifItensRepetidos = False
    Dim i As Integer
    
    If Not IsNumeric(varCampo) Then varCampo = UCase(Trim(varCampo))
    
    For i = 1 To (grdGenerica.Rows - 1)
        If i <> intRow And grdGenerica.Cell(flexcpText, i, intCol) = varCampo Then Exit Function
    Next i
    FcVerifItensRepetidos = True
End Function

Public Function FcExisteLinhaVaziaFilho(grdGenerica As Variant, Col As Long, ColCod As Long, ColActio2Do As Long, intCodPai As String) As Boolean
    FcExisteLinhaVaziaFilho = False
    Dim i As Integer
    With grdGenerica
        For i = 1 To (.Rows - 1)
            If .Cell(flexcpText, i, ColCod) = intCodPai And .Cell(flexcpText, i, ColActio2Do) = dacEnumUpdateAction_Insert Then
                If .Cell(flexcpText, i, Col) = Empty Then Exit Function
            End If
        Next i
    End With
    FcExisteLinhaVaziaFilho = True
End Function

Public Sub CarregaDadosGrdFilho(grdGenerica As Variant, ColActio2Do As Long, ColCod As Long, CodPai As Variant)
    Dim i As Integer
    Dim varCODVARIF As Variant
    
    If IsNumeric(CodPai) Then CodPai = CDbl(CodPai)
    
    With grdGenerica
        For i = 1 To (.Rows - 1)
            
            If IsNumeric(.Cell(flexcpText, i, ColCod)) Then
                varCODVARIF = CDbl(.Cell(flexcpText, i, ColCod))
            Else
                varCODVARIF = Trim(.Cell(flexcpText, i, ColCod))
            End If
            
            If varCODVARIF = CodPai Then
                If .Cell(flexcpText, i, ColActio2Do) <> dacEnumUpdateAction_delete Then
                   .RowHidden(i) = False
                Else
                   .RowHidden(i) = True
                End If
            Else
                .RowHidden(i) = True
            End If
        Next i
    End With
End Sub

Public Sub CarregaDadosGrdFilhoSemAction2Do(grdGenerica As Variant, ColCod As Long, CodPai As Variant)
    Dim i As Integer
    Dim varCODVARIF As Variant
    
    If IsNumeric(CodPai) Then CodPai = CDbl(CodPai)
    
    With grdGenerica
        For i = 1 To (.Rows - 1)
            
            If IsNumeric(.Cell(flexcpText, i, ColCod)) Then
                varCODVARIF = CDbl(.Cell(flexcpText, i, ColCod))
            Else
                varCODVARIF = Trim(.Cell(flexcpText, i, ColCod))
            End If
            
            If varCODVARIF = CodPai Then
                .RowHidden(i) = False
            Else
                .RowHidden(i) = True
            End If
        Next i
    End With
End Sub

Public Sub GravaLogModulo(intFilial_Atu As Integer, lngCodLog As Long, strModulo As String, strACAO As String, lngCodVendedor As Long, strCodReg As String, Linha As Variant)

On Error GoTo err_Trans
    
    '' Inicia transação
    
    If adoBanco_Dados.State = 0 Then Set adoBanco_Dados = Banco_Dados(Linha)
    If adoBanco_Dados.State = 0 Then
       MsgBox "Não foi possivel acessar o banco !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Sub
    End If
    
    adoBanco_Dados.BeginTrans
    BGRV.ActiveConnection = adoBanco_Dados

    sSql = "Insert Into SGI_LOGMODULO (SGI_FILIAL" & vbCrLf
    sSql = sSql & "                   ,SGI_CODIGO" & vbCrLf
    sSql = sSql & "                   ,SGI_MODULO" & vbCrLf
    sSql = sSql & "                   ,SGI_ACAO" & vbCrLf
    sSql = sSql & "                   ,SGI_CODUSUARIO" & vbCrLf
    sSql = sSql & "                   ,SGI_CODREG" & vbCrLf
    sSql = sSql & "                   ,SGI_DATA)" & vbCrLf
    sSql = sSql & "           Values (" & vbCrLf
    sSql = sSql & "                   " & intFilial_Atu & vbCrLf
    sSql = sSql & "                 ," & lngCodLog & vbCrLf
    sSql = sSql & "                 ,'" & Trim(strModulo) & "'" & vbCrLf
    sSql = sSql & "                 ,'" & Trim(strACAO) & "'" & vbCrLf
    sSql = sSql & "                 ," & lngCodVendedor & vbCrLf
    sSql = sSql & "                 ," & Trim(strCodReg) & vbCrLf
    sSql = sSql & "                 ,'" & Format(Now, "MM/DD/YYYY HH:MM:SS") & "'" & vbCrLf
    sSql = sSql & "                   )"
    
    BGRV.CommandText = sSql
    BGRV.Execute

    adoBanco_Dados.CommitTrans
     
    Exit Sub

err_Trans:
    
    adoBanco_Dados.RollbackTrans

    ''Dim objErro    As Object
    ''Set objErro = CreateObject("BLBCWS.clsFuncoes")
    ''Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, strAcao, sSql)
    ''Set objErro = Nothing

End Sub


Public Function Atualiza(strACAO As String, lngCODIGO As Long, lngFilial_Atu As Integer, strModulo As String, Linha As Variant, Optional strFILIALPED As String) As Boolean
    
On Error GoTo Erro_Atualiza

    Atualiza = False
    
    If adoBanco_Dados.State = 0 Then Set adoBanco_Dados = Banco_Dados(Linha)
    If adoBanco_Dados.State = 0 Then
       MsgBox "Não foi possivel acessar o banco !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Function
    End If
    
    '' Inicia transação
    adoBanco_Dados.BeginTrans
    BGRV.ActiveConnection = adoBanco_Dados
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "      *" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_ATUALIZA" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      SGI_FILIAL    = " & lngFilial_Atu & vbCrLf
    sSql = sSql & "  And SGI_MODULO    = '" & Trim(strModulo) & "'" & vbCrLf
    If Len(Trim(strFILIALPED)) > 0 Then sSql = sSql & "  And SGI_FILIALPED = " & Trim(strFILIALPED)
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF Then
       
        sSql = "Insert Into SGI_ATUALIZA (SGI_FILIAL" & vbCrLf
        sSql = sSql & "                  ,SGI_MODULO" & vbCrLf
        sSql = sSql & "                  ,SGI_ACAO" & vbCrLf
        sSql = sSql & "                  ,SGI_CODIGO" & vbCrLf
        
        If Len(Trim(strFILIALPED)) > 0 Then
            sSql = sSql & "                  ,SGI_FILIALPED" & vbCrLf
        End If
        
        sSql = sSql & "         ) Values (" & vbCrLf
        sSql = sSql & "                   " & lngFilial_Atu & vbCrLf
        sSql = sSql & "                 ,'" & Trim(strModulo) & "'" & vbCrLf
        sSql = sSql & "                 ,'" & Trim(strACAO) & "'" & vbCrLf
        sSql = sSql & "                 ,'" & Trim(Str(lngCODIGO)) & "'" & vbCrLf
        
        If Len(Trim(strFILIALPED)) > 0 Then
            sSql = sSql & "                 ," & strFILIALPED & vbCrLf
        End If
        
        sSql = sSql & "                   )"
    
        BGRV.CommandText = sSql
        BGRV.Execute
    
    Else
    
        sSql = sSql & "Update SGI_ATUALIZA Set" & vbCrLf
        sSql = sSql & "           SGI_ACAO   = '" & Trim(strACAO) & "'" & vbCrLf
        sSql = sSql & "          ,SGI_CODIGO = '" & Trim(Str(lngCODIGO)) & "'" & vbCrLf
        sSql = sSql & "Where " & vbCrLf
        sSql = sSql & "      SGI_FILIAL = " & lngFilial_Atu & vbCrLf
        sSql = sSql & "  And SGI_MODULO = '" & Trim(strModulo) & "'" & vbCrLf
        If Len(Trim(strFILIALPED)) > 0 Then sSql = sSql & "  And SGI_FILIALPED = " & Trim(strFILIALPED)
        
        BGRV.CommandText = sSql
        BGRV.Execute
    
    End If
    BREC.Close
    
    Atualiza = True
    
    adoBanco_Dados.CommitTrans
    
    Exit Function
        
Erro_Atualiza:
    
     adoBanco_Dados.RollbackTrans
     
     ''Dim objErro    As Object
     ''Set objErro = CreateObject("BLBCWS.clsFuncoes")
     ''Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, strAcao, sSql)
     ''Set objErro = Nothing
    
End Function


Public Function Gera_Codigo(sModulo As String, intFILIAL As Integer, Linha As Variant) As Long

    Gera_Codigo = 1
    
    If adoBanco_Dados.State = 0 Then Set adoBanco_Dados = Banco_Dados(Linha)
    If adoBanco_Dados.State = 0 Then
       MsgBox "Não foi possivel acessar o banco !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Function
    End If
    BGRV.ActiveConnection = adoBanco_Dados
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql + "       (Max(SGI_NUMERO) + 1) As SGI_NUMERO " & vbCrLf
    sSql = sSql + "  From " & vbCrLf
    sSql = sSql + "       SGI_NUMERO " & vbCrLf
    sSql = sSql + " Where " & vbCrLf
    sSql = sSql + "       SGI_MODULO = '" & sModulo & "'"
    sSql = sSql + "   And SGI_FILIAL = " & intFILIAL
    
    BREC.Open sSql, adoBanco_Dados
    
    If Not BREC.EOF Then
    
       If IsNull(BREC!SGI_NUMERO) = True Then
          
          Gera_Codigo = 1
          
          sSql = "Insert into SGI_NUMERO (SGI_FILIAL,SGI_NUMERO,SGI_MODULO) Values(" & vbCrLf
          sSql = sSql + "                                              " & intFILIAL & vbCrLf
          sSql = sSql + "                                            ,1" & vbCrLf
          sSql = sSql + "                                            ,'" & sModulo & "'" & vbCrLf
          sSql = sSql + "                                          )" & vbCrLf
          
       ElseIf BREC!SGI_NUMERO > 1 Then
       
          Gera_Codigo = BREC!SGI_NUMERO
          
          sSql = "Update SGI_NUMERO Set " & vbCrLf
          sSql = sSql + "           SGI_NUMERO = " & BREC!SGI_NUMERO & vbCrLf
          sSql = sSql + "         Where " & vbCrLf
          sSql = sSql + "               SGI_MODULO = '" & sModulo & "'" & vbCrLf
          sSql = sSql + "           And SGI_FILIAL =  " & intFILIAL
       
       End If
       
       BGRV.CommandText = sSql
       BGRV.Execute
       
       
    End If
    
    BREC.Close
    
End Function

Public Sub MudaBotoes(CmdSalva As Object, cmdAltera As Object, strTipOper As String)

    If (strTipOper = "I" Or strTipOper = "A") Then
        CmdSalva.Enabled = True
        cmdAltera.Enabled = False
    ElseIf strTipOper = "C" Then
        CmdSalva.Enabled = False
        cmdAltera.Enabled = True
    End If

End Sub

Public Sub MudaCaption(frmFormulario As Object, strNomeForm As String, strACAO As String)

    Dim strDescAcao As String
    
    If strACAO = "I" Then strDescAcao = "[ INCLUSÃO ]"
    If strACAO = "A" Then strDescAcao = "[ ALTERAÇÃO ]"
    If strACAO = "C" Then strDescAcao = "[ CONSULTA ]"
    
    frmFormulario.Caption = strNomeForm & strDescAcao

End Sub

Public Sub ExclLinhaGridAction2Do(grdGenerica As Variant, intRow As Integer, lngCOLAction2Do As Long)
    Dim i As Integer
    With grdGenerica
        If (.Rows - 1) > 0 Then
            For i = 1 To (.Rows - 1)
                .RowHidden(i) = False
                If .Cell(flexcpText, i, lngCOLAction2Do) = dacEnumUpdateAction_delete Then .RowHidden(i) = True
            Next i
        End If
    End With
End Sub

Public Function FcVerifItensRepetidosAct2Do(grdGenerica As Variant, intRow As Long, intCol As Long, varCampo As Variant, lngColAct2Do As Long) As Boolean
    FcVerifItensRepetidosAct2Do = False
    Dim i As Integer
    
    If Not IsNumeric(varCampo) Then varCampo = UCase(Trim(varCampo))
    
    For i = 1 To (grdGenerica.Rows - 1)
        If i <> intRow And _
                grdGenerica.Cell(flexcpText, i, intCol) = varCampo And _
                grdGenerica.Cell(flexcpText, i, lngColAct2Do) <> dacEnumUpdateAction_delete Then
                Exit Function
        End If
    Next i
    FcVerifItensRepetidosAct2Do = True
End Function

Public Sub ExcLinhaGrdFilhoAct2Do(grdGRIDFILHO As Variant, lngCOLFILHO As Long, strIDPai As String, lngColAct2Do As Long)
    Dim i As Integer
    With grdGRIDFILHO
        For i = 1 To (.Rows - 1)
            If Trim(.Cell(flexcpText, i, lngCOLFILHO)) = Trim(strIDPai) Then
               .Cell(flexcpText, i, lngColAct2Do) = dacEnumUpdateAction_delete
               .RowHidden(i) = True
               If .Cell(flexcpText, i, lngColAct2Do) = dacEnumUpdateAction_delete Then .RowHidden(i) = True
            End If
        Next i
    End With
End Sub

Public Sub TrocaAction2Do(grdGENETICA As Variant, lngROW As Long, lngCOLAction2Do As Long, strDADOSORIG As String, strDADOSDIG As String)
    If grdGENETICA.Cell(flexcpText, lngROW, lngCOLAction2Do) = dacEnumUpdateAction_Ignore Then
        If Trim(strDADOSORIG) <> Trim(strDADOSDIG) Then grdGENETICA.Cell(flexcpText, lngROW, lngCOLAction2Do) = dacEnumUpdateAction_update
    End If
End Sub

Public Sub ExcLinhaGrdFilho(grdGRIDFILHO As Variant, lngCOLFILHO As Long, strIDPai As String)
    Dim i As Integer
VOLTA:
    For i = 1 To (grdGRIDFILHO.Rows - 1)
        If Trim(grdGRIDFILHO.Cell(flexcpText, i, lngCOLFILHO)) = Trim(strIDPai) Then
           grdGRIDFILHO.RemoveItem i
           GoTo VOLTA
        End If
    Next i
End Sub

Public Sub Sub_DescErro(strErroNumber As String, strErrDesc As String, strOperacao As String, strErro As String, Optional strModulo As String, Optional strFUNCAO As String, Optional strCAMARQERRO As String, Optional strDescSQL As String, Optional strUSUARIO As String)
    
    
On Error GoTo Err_Erro

    Dim strDTERRO As String
    Dim strHRERRO As String
    
    strDTERRO = Trim(Replace(Format(Now, "DD/MM/YYYY"), "/", ""))
    strHRERRO = Trim(Replace(Format(Now, "HH:MM:SS"), ":", ""))
    
    If CLng(strErroNumber) = 429 Then strErrDesc = "ATENÇÃO - Não Foi possivel criar o Objeto favor sair do sistema e reiniciar !!!"
    If CLng(strErroNumber) = 440 Then strErrDesc = "ATENÇÃO - Erro de Automação favor sair do sistema e reiniciar !!!"
    
    
    
    If Len(Trim(strOperacao)) = 0 Then MsgBox "Erro : " & strErroNumber & vbCrLf & _
                                              "Desc : " & strErrDesc & vbCrLf & _
                                              "" & vbCrLf & _
                                              strErro, vbOKOnly + vbCritical, "Aviso Sistema"
    
    If strOperacao = "I" Then MsgBox "Foi impossivel incluir o registro !!!" & vbCrLf & _
                                 "Erro : " & strErroNumber & vbCrLf & _
                                 "Desc : " & strErrDesc & vbCrLf & _
                                 "" & vbCrLf & _
                                 strErro, vbOKOnly + vbCritical, "Aviso Sistema"
    
    If strOperacao = "A" Then MsgBox "Foi impossivel altera o registro !!!" & vbCrLf & _
                                 "Erro : " & strErroNumber & vbCrLf & _
                                 "Desc : " & strErrDesc & vbCrLf & _
                                 "" & vbCrLf & _
                                 strErro, vbOKOnly + vbCritical, "Aviso Sistema"
    
    If strOperacao = "E" Then MsgBox "Foi impossivel excluir o registro !!!" & vbCrLf & _
                                 "Erro : " & strErroNumber & vbCrLf & _
                                 "Desc : " & strErrDesc & vbCrLf & _
                                 "" & vbCrLf & _
                                 strErro, vbOKOnly + vbCritical, "Aviso Sistema"

    If strOperacao = "AL" Then MsgBox "Foi impossivel gerar a estrutura !!!" & vbCrLf & _
                                 "Erro : " & strErroNumber & vbCrLf & _
                                 "Desc : " & strErrDesc & vbCrLf & _
                                 "" & vbCrLf & _
                                 strErro, vbOKOnly + vbCritical, "Aviso Sistema"


    If Len(Trim(strCAMARQERRO)) > 0 Then

        Open strCAMARQERRO & "ERR_" & strDTERRO & "_" & strHRERRO & ".txt" For Output As #1
        
        If strOperacao = "I" Then Print #1, "Foi impossivel Incluir o registro !!!"
        If strOperacao = "A" Then Print #1, "Foi impossivel Alterar o registro !!!"
        If strOperacao = "E" Then Print #1, "Foi impossivel Excluir o registro !!!"
        
        Print #1, "Módulo         : " & strModulo
        Print #1, "Função         : " & strFUNCAO
        Print #1, "Usuário        : " & strUSUARIO
        Print #1, " "
        Print #1, "Erro Numero    : " & strErroNumber
        Print #1, "Erro Descrição : " & strErrDesc
        Print #1, " "
        Print #1, " "
        Print #1, strErro
        Print #1, " "
        Print #1, " "
        Print #1, strDescSQL
        
        
        Close #1
        
    End If

    Exit Sub

Err_Erro:

        Open App.Path & "\" & "ERR_" & strDTERRO & "_" & strHRERRO & ".txt" For Output As #2
        
        If strOperacao = "I" Then Print #2, "Foi impossivel Incluir o registro !!!"
        If strOperacao = "A" Then Print #2, "Foi impossivel Alterar o registro !!!"
        If strOperacao = "E" Then Print #2, "Foi impossivel Excluir o registro !!!"
        
        Print #2, "Módulo         : " & strModulo
        Print #2, "Função         : " & strFUNCAO
        Print #2, "Usuário        : " & strUSUARIO
        Print #2, " "
        Print #2, "Erro Numero    : " & strErroNumber
        Print #2, "Erro Descrição : " & strErrDesc
        Print #2, " "
        Print #2, " "
        Print #2, strErro
        
        Close #2

End Sub

Public Function Preenche_Mes(ByVal Lista As Object)
   
   Dim Indice As Integer
   
   Lista.Clear
   For Indice = 1 To 12
      If Indice = 1 Then Lista.AddItem "JANEIRO"
      If Indice = 2 Then Lista.AddItem "FEVEREIRO"
      If Indice = 3 Then Lista.AddItem "MARCO"
      If Indice = 4 Then Lista.AddItem "ABRIL"
      If Indice = 5 Then Lista.AddItem "MAIO"
      If Indice = 6 Then Lista.AddItem "JUNHO"
      If Indice = 7 Then Lista.AddItem "JULHO"
      If Indice = 8 Then Lista.AddItem "AGOSTO"
      If Indice = 9 Then Lista.AddItem "SETEMBRO"
      If Indice = 10 Then Lista.AddItem "OUTUBRO"
      If Indice = 11 Then Lista.AddItem "NOVEMBRO"
      If Indice = 12 Then Lista.AddItem "DEZEMBRO"
      Lista.ItemData(Lista.NewIndex) = Indice
   Next Indice
   
End Function


Public Function Preenche_Ano(ByVal Lista As Object)
   
   Dim Indice       As Integer
   Dim intANO       As Integer
   Dim intANOFINAL  As Integer
   
   intANO = Year(Now)
   intANOFINAL = intANO
   
   Lista.Clear
   For Indice = 1 To 40
      Lista.AddItem intANOFINAL
      Lista.ItemData(Lista.NewIndex) = intANOFINAL
      intANOFINAL = (intANO + Indice)
   Next Indice
   
End Function

Public Sub LOG_ACAO(intFILIAL As Integer, Linha As Variant, strModulo As String, strUSUARIO As String, strNOMCOMP As String, strVERSAO As String, strACAO As String, lngCODREG As String, strCAMPO As String, strCONTANT As String, strCONTNOV As String)

On Error GoTo err_Trans

    Dim lngCOEREGLOG As Long
    
    lngCOEREGLOG = Gera_Codigo("LOG_ACAO", intFILIAL, Linha)
    
    '' Inicia transação
    adoBanco_Dados.BeginTrans
    BGRV.ActiveConnection = adoBanco_Dados
    
    sSql = ""
    
    sSql = "Insert into SGI_LOG_ACAO (" & vbCrLf
    sSql = sSql & "                          SGI_FILIAL" & vbCrLf
    sSql = sSql & "                         ,SGI_CODIGO" & vbCrLf
    sSql = sSql & "                         ,SGI_MODULO" & vbCrLf
    sSql = sSql & "                         ,SGI_USUARIO" & vbCrLf
    sSql = sSql & "                         ,SGI_NOMCOMP" & vbCrLf
    sSql = sSql & "                         ,SGI_VERSAO" & vbCrLf
    sSql = sSql & "                         ,SGI_ACAO" & vbCrLf
    sSql = sSql & "                         ,SGI_CODREG" & vbCrLf
    sSql = sSql & "                         ,SGI_DATA" & vbCrLf
    sSql = sSql & "                         ,SGI_CAMPO" & vbCrLf
    sSql = sSql & "                         ,SGI_CONT_ANT" & vbCrLf
    sSql = sSql & "                         ,SGI_CONT_NOV" & vbCrLf
    sSql = sSql & "                ) Values (" & vbCrLf
    sSql = sSql & "                          " & intFILIAL & vbCrLf
    sSql = sSql & "                         ," & lngCOEREGLOG & vbCrLf
    sSql = sSql & "                         ," & "'" & strModulo & "'" & vbCrLf
    sSql = sSql & "                         ," & "'" & strUSUARIO & "'" & vbCrLf
    sSql = sSql & "                         ," & "'" & strNOMCOMP & "'" & vbCrLf
    sSql = sSql & "                         ," & "'" & strVERSAO & "'" & vbCrLf
    sSql = sSql & "                         ," & "'" & strACAO & "'" & vbCrLf
    sSql = sSql & "                         ," & lngCODREG & vbCrLf
    sSql = sSql & "                         ,'" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf
    sSql = sSql & "                         ," & "'" & strCAMPO & "'" & vbCrLf
    sSql = sSql & "                         ," & "'" & strCONTANT & "'" & vbCrLf
    sSql = sSql & "                         ," & "'" & strCONTNOV & "'" & vbCrLf
    sSql = sSql & "                         )"

    BGRV.CommandText = sSql
    BGRV.Execute
    
    adoBanco_Dados.CommitTrans
    
    Exit Sub

err_Trans:
    
    adoBanco_Dados.RollbackTrans

End Sub

Public Sub GravaLogForm(intFILIAL As Integer, lngCODIGO As Long, lngCodUsuario As Long, strOPER As String, strFORM As String)

On Error GoTo err_GravaLogForm
    
    '' Inicia transação
    adoBanco_Dados.BeginTrans
    BGRV.ActiveConnection = adoBanco_Dados
    
    '' =====================================
    '' Grava Log
    sSql = ""
    
    sSql = "Insert Into SGI_LOGFORM (" & vbCrLf
    sSql = sSql & "                              SGI_FILIAL" & vbCrLf
    sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
    sSql = sSql & "                             ,SGI_DATA" & vbCrLf
    sSql = sSql & "                             ,SGI_HORA" & vbCrLf
    sSql = sSql & "                             ,SGI_CODUSUARIO" & vbCrLf
    sSql = sSql & "                             ,SGI_ACAO" & vbCrLf
    sSql = sSql & "                             ,SGI_FORM" & vbCrLf
    sSql = sSql & "                    ) Values (" & vbCrLf
    sSql = sSql & "                              " & intFILIAL & vbCrLf
    sSql = sSql & "                             ," & lngCODIGO & vbCrLf
    sSql = sSql & "                             ,'" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf
    sSql = sSql & "                             ,'" & Format(Now, "HH:MM:SS") & "'" & vbCrLf
    sSql = sSql & "                             ," & lngCodUsuario & vbCrLf
    sSql = sSql & "                             ,'" & strOPER & "'" & vbCrLf
    sSql = sSql & "                             ,'" & strFORM & "'" & vbCrLf
    sSql = sSql & "                             )"
    
    BGRV.CommandText = sSql
    BGRV.Execute
    '' =====================================
        
    adoBanco_Dados.CommitTrans
        
    Exit Sub

err_GravaLogForm:
    
    If adoBanco_Dados.State = 1 Then adoBanco_Dados.RollbackTrans

End Sub


Public Function PegaUsuario(lngCodUsu As Long, Linha As Variant, intFILIAL As Integer) As String

On Error GoTo Err_PegaUsuario
    
    PegaUsuario = ""
    
    If adoBanco_Dados.State = 0 Then Set adoBanco_Dados = Banco_Dados(Linha)
    If adoBanco_Dados.State = 0 Then
       MsgBox "Não foi possivel acessar o banco !!!", vbOKOnly + vbExclamation, "Aviso"
       Exit Function
    End If
    
    
    If lngCodUsu = 0 Then
        PegaUsuario = "CWS"
        Exit Function
    End If

    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_NOME " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_USUARIO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCodUsu
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PegaUsuario = Crypt(BREC!SGI_NOME)
    BREC.Close
    
    Exit Function

Err_PegaUsuario:

    If BREC.State = 1 Then BREC.Close

End Function


Public Function Banco_Dados_Imagem(Conecxao As Variant) As ADODB.Connection

On Error GoTo err_sys
  
  Dim Linha()     As Variant
  Dim intregs     As Integer
  Dim i           As Integer
  
  Dim strNoBanco  As String
  Dim strProvider As String
  Dim strConnt    As String
  Dim strLinReg() As String
    
  '"Microsoft.Jet.OLEDB.4.0"
  '------------------------------------------------------
  'Abre uma conexção com o banco
 
  Set Banco_Dados_Imagem = New ADODB.Connection
  
  If Right(Conecxao(0), 1) = 1 Then
     Banco_Dados_Imagem.Provider = Mid(Conecxao(2), 11, 50) '' Drive do banco em Uso
     Banco_Dados_Imagem.ConnectionString = Mid(Conecxao(1), 9, 50) & Mid(Conecxao(3), 11, 50) '' Caminho do Banco
     Banco_Dados_Imagem.CommandTimeout = 1200 '' 20 minutos
  End If
  
  If Right(Conecxao(0), 1) = 2 Then
     Banco_Dados_Imagem.ConnectionString = Conecxao(11) & Conecxao(2) & Conecxao(3)
  End If
  
  Banco_Dados_Imagem.Open
    
  Exit Function
err_sys:
  
  If Err.Number = 53 Then
     MsgBox "ATENÇÃO - Erro : " & Err.Number & " Arquivo de Configuração não encontrado" & vbCrLf & _
            "Função = Banco_Dados_Imagem" & vbCrLf & _
            "Biblioteca = BLBCWS.clsFuncoes" & vbCrLf & _
            "Caminho = " & App.Path, vbOKOnly + vbCritical, "Aviso"
  Else
     MsgBox "ATENÇÃO - " & Err.Number & " - " & Err.Description, vbOKOnly + vbCritical, "Aviso"
  End If
  
End Function


