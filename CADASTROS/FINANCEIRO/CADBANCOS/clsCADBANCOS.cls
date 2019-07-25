VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADBANCOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim intFilial      As Integer
Dim intCADANCOCOD  As Integer
Dim strCADANCOAGE  As String
Dim strCADANCOCOR  As String
Dim strCADANCODES  As String

Public Property Get CADANCODES() As String
    CADANCODES = strCADANCODES
End Property

Public Property Let CADANCODES(ByVal vData As String)
    strCADANCODES = vData
End Property

Public Property Get CADANCOCOR() As String
    CADANCOCOR = strCADANCOCOR
End Property

Public Property Let CADANCOCOR(ByVal vData As String)
    strCADANCOCOR = vData
End Property

Public Property Get CADANCOAGE() As String
    CADANCOAGE = strCADANCOAGE
End Property

Public Property Let CADANCOAGE(ByVal vData As String)
    strCADANCOAGE = vData
End Property

Public Property Get CADANCOCOD() As Integer
    CADANCOCOD = intCADANCOCOD
End Property

Public Property Let CADANCOCOD(ByVal vData As Integer)
    intCADANCOCOD = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADBANCOSP.cCaminho = cCaminho
    frmCADBANCOSP.Linha = Linha
    frmCADBANCOSP.FILIAL = iFilial
    frmCADBANCOSP.strAcesso = strAcesso
    frmCADBANCOSP.strUSUARIO = strUSUARIO
    frmCADBANCOSP.Show vbModal
End Sub

Public Sub cConnectNovo(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String)
    frmCADBANCOS.cCaminho = cCaminho
    frmCADBANCOS.Linha = Linha
    frmCADBANCOS.FILIAL = iFilial
    frmCADBANCOS.strAcesso = strAcesso
    frmCADBANCOS.strUSUARIO = strUSUARIO
    frmCADBANCOS.cTipOper = "I"
    frmCADBANCOS.Show vbModal
End Sub

Public Function Pesq_CadBancos() As Boolean

   Pesq_CadBancos = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       * " & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_CADBANCOS " & vbCrLf
   sSql = sSql + " Where " & vbCrLf
   sSql = sSql + "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_CadBancos = True
   BREC.Close

End Function

Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_grava

     GRAVA = False
         
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Then
        
        '' Grava Bancos
        sSql = "Insert Into SGI_CADBANCOS (SGI_FILIAL,SGI_CODIGO,SGI_AGENCIA,SGI_CC,SGI_DESCRICAO) Values (" & vbCrLf
        sSql = sSql & "                            " & intFilial
        sSql = sSql & "                           ," & intCADANCOCOD
        sSql = sSql & "                           ,'" & strCADANCOAGE & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strCADANCOCOR & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strCADANCODES & "'" & vbCrLf
        sSql = sSql & "                           )"
     
     ElseIf strOPER = "A" Then
     
        '' Altera Bancos
        sSql = "Update SGI_CADBANCOS Set " & vbCrLf
        sSql = sSql & "               SGI_AGENCIA   = '" & strCADANCOAGE & "'" & vbCrLf
        sSql = sSql & "              ,SGI_CC        = '" & strCADANCOCOR & "'" & vbCrLf
        sSql = sSql & "              ,SGI_DESCRICAO = '" & strCADANCODES & "'" & vbCrLf
        sSql = sSql & "         Where " & vbCrLf
        sSql = sSql & "               SGI_CODIGO    = " & intCADANCOCOD & vbCrLf
        sSql = sSql & "           And SGI_FILIAL    = " & intFilial
     
     ElseIf strOPER = "E" Then
     
        '' Apaga especie de orçamento
        sSql = "Delete from SGI_CADBANCOS " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intCADANCOCOD & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
     
     End If
     
     BGRV.CommandText = sSql
     BGRV.Execute
     
     adoBanco_Dados.CommitTrans
     
     GRAVA = True
     
     Exit Function
     
err_grava:
     
     adoBanco_Dados.RollbackTrans
    
     Dim objErro    As Object
     Set objErro = CreateObject("BLBCWS.clsFuncoes")
     Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, strOPER, sSql)
     Set objErro = Nothing
     
End Function

Public Function Gera_Codigo(sModulo As String) As Integer

    Gera_Codigo = 1
    
    BGRV.ActiveConnection = adoBanco_Dados
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql + "       (Max(SGI_NUMERO) + 1) As SGI_NUMERO " & vbCrLf
    sSql = sSql + "  From " & vbCrLf
    sSql = sSql + "       SGI_NUMERO " & vbCrLf
    sSql = sSql + " Where " & vbCrLf
    sSql = sSql + "       SGI_MODULO = '" & sModulo & "'"
    sSql = sSql + "   And SGI_FILIAL = " & intFilial
    
    BREC.Open sSql, adoBanco_Dados
    
    If Not BREC.EOF Then
    
       If IsNull(BREC!SGI_NUMERO) = True Then
          
          Gera_Codigo = 1
          
          sSql = "Insert into SGI_NUMERO (SGI_FILIAL,SGI_NUMERO,SGI_MODULO) Values(" & vbCrLf
          sSql = sSql + "                                              " & intFilial & vbCrLf
          sSql = sSql + "                                            ,1" & vbCrLf
          sSql = sSql + "                                            ,'" & sModulo & "'" & vbCrLf
          sSql = sSql + "                                          )" & vbCrLf
          
       ElseIf BREC!SGI_NUMERO > 1 Then
       
          Gera_Codigo = BREC!SGI_NUMERO
          
          sSql = "Update SGI_NUMERO Set " & vbCrLf
          sSql = sSql + "           SGI_NUMERO = " & BREC!SGI_NUMERO & vbCrLf
          sSql = sSql + "         Where " & vbCrLf
          sSql = sSql + "               SGI_MODULO = '" & sModulo & "'" & vbCrLf
          sSql = sSql + "           And SGI_FILIAL =  " & intFilial
       
       End If
       
       BGRV.CommandText = sSql
       BGRV.Execute
       
       
    End If
    
    BREC.Close
    
End Function

Public Function Carrega_campos() As Boolean

   Carrega_campos = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADBANCOS " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & intCADANCOCOD & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   
   If Not BREC.EOF Then
      
      intCADANCOCOD = BREC!SGI_CODIGO
      strCADANCOAGE = BREC!SGI_AGENCIA
      strCADANCOCOR = BREC!SGI_CC
      strCADANCODES = BREC!SGI_DESCRICAO
      
      Carrega_campos = True
   End If
   
   BREC.Close

End Function

Public Function Atualiza(strAcao As String, lngCODIGO As Long, lngFilial_Atu As Integer, strModulo As String) As Boolean
    
On Error GoTo Erro_Atualiza

    Atualiza = False
    
    '' Inicia transação
    adoBanco_Dados.BeginTrans
    BGRV.ActiveConnection = adoBanco_Dados
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "      *" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_ATUALIZA" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & lngFilial_Atu & vbCrLf
    sSql = sSql & "  And SGI_MODULO = '" & Trim(strModulo) & "'"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF Then
       
        sSql = "Insert Into SGI_ATUALIZA (SGI_FILIAL" & vbCrLf
        sSql = sSql & "                  ,SGI_MODULO" & vbCrLf
        sSql = sSql & "                  ,SGI_ACAO" & vbCrLf
        sSql = sSql & "                  ,SGI_CODIGO)" & vbCrLf
        sSql = sSql & "           Values (" & vbCrLf
        sSql = sSql & "                   " & lngFilial_Atu & vbCrLf
        sSql = sSql & "                 ,'" & Trim(strModulo) & "'" & vbCrLf
        sSql = sSql & "                 ,'" & Trim(strAcao) & "'" & vbCrLf
        sSql = sSql & "                 ,'" & Trim(Str(lngCODIGO)) & "'" & vbCrLf
        sSql = sSql & "                   )"
    
        BGRV.CommandText = sSql
        BGRV.Execute
    
    Else
    
        sSql = sSql & "Update SGI_ATUALIZA Set" & vbCrLf
        sSql = sSql & "           SGI_ACAO   = '" & Trim(strAcao) & "'" & vbCrLf
        sSql = sSql & "          ,SGI_CODIGO = '" & Trim(Str(lngCODIGO)) & "'" & vbCrLf
        sSql = sSql & "Where " & vbCrLf
        sSql = sSql & "      SGI_FILIAL = " & lngFilial_Atu & vbCrLf
        sSql = sSql & "  And SGI_MODULO = '" & Trim(strModulo) & "'"
        
        BGRV.CommandText = sSql
        BGRV.Execute
    
    End If
    BREC.Close
    
    Atualiza = True
    
    adoBanco_Dados.CommitTrans
    
    Exit Function
        
Erro_Atualiza:
    
     adoBanco_Dados.RollbackTrans
     
     Dim objErro    As Object
     Set objErro = CreateObject("BLBCWS.clsFuncoes")
     Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, strAcao, sSql)
     Set objErro = Nothing
    
End Function

