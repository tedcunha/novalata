VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADSAIMAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim intFilial        As Integer
Dim lngCADREQSAICOD  As Long
Dim lngCADDEPCOD     As Long
Dim lngCADUSUCOD     As Long
Dim dtCADDTREQ       As Date
Dim lngCADREQPROD    As Long
Dim varrITENSREQ     As Variant
Dim varrITENSREQBACK As Variant
Dim lngCODLCTO       As Long
Dim objFuncoes       As Object
Dim lngCODMOTIVOS    As Long
Dim strCODNF         As String
Dim lngCODCLIE       As Long

Public Property Get CODCLIE() As Long
    CODCLIE = lngCODCLIE
End Property

Public Property Let CODCLIE(ByVal vData As Long)
    lngCODCLIE = vData
End Property

Public Property Get CODNF() As String
    CODNF = strCODNF
End Property

Public Property Let CODNF(ByVal vData As String)
    strCODNF = vData
End Property

Public Property Get CODMOTIVOS() As Long
    CODMOTIVOS = lngCODMOTIVOS
End Property

Public Property Let CODMOTIVOS(ByVal vData As Long)
    lngCODMOTIVOS = vData
End Property

Public Property Get ITENSREQBACK() As Variant
    ITENSREQBACK = varrITENSREQBACK
End Property

Public Property Let ITENSREQBACK(ByVal vData As Variant)
    varrITENSREQBACK = vData
End Property

Public Property Get CODLCTO() As Long
    CODLCTO = lngCODLCTO
End Property

Public Property Let CODLCTO(ByVal vData As Long)
    lngCODLCTO = vData
End Property

Public Property Get ITENSREQ() As Variant
    ITENSREQ = varrITENSREQ
End Property

Public Property Let ITENSREQ(ByVal vData As Variant)
    varrITENSREQ = vData
End Property

Public Property Get CADREQPROD() As Long
    CADREQPROD = lngCADREQPROD
End Property

Public Property Let CADREQPROD(ByVal vData As Long)
    lngCADREQPROD = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Property Get CADREQSAICOD() As Long
    CADREQSAICOD = lngCADREQSAICOD
End Property

Public Property Let CADREQSAICOD(ByVal vData As Long)
    lngCADREQSAICOD = vData
End Property

Public Property Get CADDEPCOD() As Long
    CADDEPCOD = lngCADDEPCOD
End Property

Public Property Let CADDEPCOD(ByVal vData As Long)
    lngCADDEPCOD = vData
End Property

Public Property Get CADUSUCOD() As Long
    CADUSUCOD = lngCADUSUCOD
End Property

Public Property Let CADUSUCOD(ByVal vData As Long)
    lngCADUSUCOD = vData
End Property

Public Property Get CADDTREQ() As Date
    CADDTREQ = dtCADDTREQ
End Property

Public Property Let CADDTREQ(ByVal vData As Date)
    dtCADDTREQ = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADSAIMATP.cCaminho = cCaminho
    frmCADSAIMATP.Linha = Linha
    frmCADSAIMATP.FILIAL = iFilial
    frmCADSAIMATP.strAcesso = strAcesso
    frmCADSAIMATP.strUSUARIO = strUSUARIO
    frmCADSAIMATP.lngCodUsuario = lngIDUsuario
    frmCADSAIMATP.Show vbModal
End Sub

Public Function Pesq_CadSaiReqMat() As Boolean

   Pesq_CadSaiReqMat = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       *" & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_CADREQSAIMAT " & vbCrLf
   sSql = sSql + " Where "
   sSql = sSql + "       SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_CadSaiReqMat = True
   BREC.Close

End Function

Public Sub PreencheCombo(Combo As Variant, strTabela As String)

     Combo.Clear
     
     If strTabela = "SGI_USUARIO" Then
        Set objFuncoes = CreateObject("BLBCWS.clsFuncoes")
     End If
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       * " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       " & strTabela & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL  = " & intFilial & vbCrLf
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        If strTabela = "SGI_USUARIO" Then
           Combo.AddItem Trim(BREC.Fields(1)) & Space(10 - Len(Trim(BREC.Fields(1)))) & " - " & objFuncoes.Crypt(Trim(BREC.Fields(2)))
        Else
           Combo.AddItem Trim(BREC.Fields(1)) & Space(10 - Len(Trim(BREC.Fields(1)))) & " - " & Trim(BREC.Fields(2))
        End If
        Combo.ItemData(Combo.NewIndex) = BREC.Fields(1)
        BREC.MoveNext
     Loop
    
     BREC.Close
    
     If strTabela = "SGI_USUARIO" Then
        Set objFuncoes = Nothing
     End If
    
End Sub

Public Sub PreencheComboProd(Combo As Variant)

     Combo.Clear
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       PRODUT.* " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADPRODUTO  PRODUT " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       PRODUT.SGI_FILIAL  = " & intFilial & vbCrLf
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_CODIGO) & Space(10 - Len(Trim(BREC!SGI_CODIGO))) & " - " & Trim(BREC!SGI_DESCRICAO)
        BREC.MoveNext
     Loop
    
     ''If Combo.ListCount > 0 Then Combo.ListIndex = 0
    
     BREC.Close
    
End Sub

Public Function Gera_Codigo(sModulo As String) As Long

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


Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_Trans
     
     GRAVA = False
         
     Dim sValor As String
     Dim I      As Integer
     
     '' Inicia transação
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Then '' Inclui
        
Retorno:
        
        '' Grava requisição HEADER(CABEÇALHO)
        sSql = "Insert Into SGI_CADREQSAIMAT (SGI_FILIAL" & vbCrLf
        sSql = sSql & "                      ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                      ,SGI_DATREQ" & vbCrLf
        sSql = sSql & "                      ,SGI_CODDEPTO" & vbCrLf
        sSql = sSql & "                      ,SGI_CODUSUAR" & vbCrLf
        sSql = sSql & "                      ,SGI_CODNF" & vbCrLf
        sSql = sSql & "                      ,SGI_CODMOTIVO" & vbCrLf
        sSql = sSql & "                      ,SGI_CODCLIENTE" & vbCrLf
        sSql = sSql & "             ) Values (" & vbCrLf
        sSql = sSql & "                        " & intFilial & vbCrLf
        sSql = sSql & "                      , " & lngCADREQSAICOD & vbCrLf
        sSql = sSql & "                      ,'" & Format(dtCADDTREQ, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                      , " & lngCADDEPCOD & vbCrLf
        sSql = sSql & "                      , " & lngCADUSUCOD & vbCrLf
        sSql = sSql & "                      , '" & strCODNF & "'" & vbCrLf
        sSql = sSql & "                      , " & lngCODMOTIVOS & vbCrLf
        sSql = sSql & "                      ," & lngCODCLIE & vbCrLf
        sSql = sSql & "                      )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Grava os Itens da Requisição
        If IsArray(varrITENSREQ) = True Then
           For I = 1 To UBound(varrITENSREQ)
               
               sSql = "Insert Into SGI_CADITREQSAIMAT  ("
               sSql = sSql & "                          SGI_FILIAL" & vbCrLf
               sSql = sSql & "                         ,SGI_CODIGO" & vbCrLf
               sSql = sSql & "                         ,SGI_IDPRODUTO" & vbCrLf
               sSql = sSql & "                         ,SGI_PRODUTO" & vbCrLf
               sSql = sSql & "                         ,SGI_QTD" & vbCrLf
               sSql = sSql & "                         ,SGI_CODREQ" & vbCrLf
               sSql = sSql & "                         ,SGI_QTDEKG" & vbCrLf
               sSql = sSql & "                         ,SGI_CODLOTE" & vbCrLf
               sSql = sSql & "                         ,SGI_INDICE" & vbCrLf
               sSql = sSql & "                         ,SGI_DATREQ" & vbCrLf
               sSql = sSql & "             ) Values (" & vbCrLf
               
               sSql = sSql & "                       " & intFilial & vbCrLf
               sSql = sSql & "                      ," & lngCADREQSAICOD & vbCrLf
               sSql = sSql & "                      ," & varrITENSREQ(I, 1) & vbCrLf
               sSql = sSql & "                      ,'" & Trim(varrITENSREQ(I, 2)) & "'" & vbCrLf
               sSql = sSql & "                      ," & varrITENSREQ(I, 3) & vbCrLf
               sSql = sSql & "                      ," & lngCADREQPROD & vbCrLf
               sSql = sSql & "                      ," & varrITENSREQ(I, 5) & vbCrLf
               sSql = sSql & "                      ," & varrITENSREQ(I, 6) & vbCrLf
               sSql = sSql & "                      ," & varrITENSREQ(I, 7) & vbCrLf
               sSql = sSql & "                      ," & varrITENSREQ(I, 8) & vbCrLf
               sSql = sSql & "                      )" & vbCrLf
               
               BGRV.CommandText = sSql
               BGRV.Execute
               
               ''If varrITENSREQ(I, 4) = 0 Then
               ''
               ''     sSql = "Insert Into SGI_PRODSALDOS (" & vbCrLf
               ''     sSql = sSql & "                            SGI_FILIAL" & vbCrLf
               ''     sSql = sSql & "                           ,SGI_IDPRODUTO" & vbCrLf
               ''     sSql = sSql & "                           ,SGI_CODCLIENTE" & vbCrLf
               ''     sSql = sSql & "                           ,SGI_SALDO" & vbCrLf
               ''     sSql = sSql & "                  ) Values (" & vbCrLf
               ''     sSql = sSql & "                            " & intFilial & vbCrLf
               ''     sSql = sSql & "                           ," & varrITENSREQ(I, 1) & vbCrLf
               ''     sSql = sSql & "                           ," & lngCODCLIE & vbCrLf
               ''     sSql = sSql & "                           ," & varrITENSREQ(I, 3) & vbCrLf
               ''     sSql = sSql & "                           )"
                    
               ''     BGRV.CommandText = sSql
               ''     BGRV.Execute
               
               ''ElseIf varrITENSREQ(I, 4) = 1 Then
               
                    '' --------------------------------------------------
                    '' Dá Entrada no Estoque
                ''    sSql = "Update SGI_PRODSALDOS " & vbCrLf
                ''    sSql = sSql & "               Set SGI_SALDO    = (SGI_SALDO - " & varrITENSREQ(I, 3) & ")" & vbCrLf
                ''    sSql = sSql & "Where " & vbCrLf
                ''    sSql = sSql & "      SGI_FILIAL     = " & FILIAL & vbCrLf
                ''    sSql = sSql & "  And SGI_IDPRODUTO  = " & varrITENSREQ(I, 1) & vbCrLf
                ''    sSql = sSql & "  And SGI_CODCLIENTE = " & lngCODCLIE
                    
                ''    BGRV.CommandText = sSql
                ''    BGRV.Execute
               
                ''End If
               
               '' --------------------------------------------------
               '' Dá Baixa no Estoque
               
               sSql = "Update SGI_CADPRODUTO Set SGI_SALDO    = (SGI_SALDO - " & varrITENSREQ(I, 3) & ")" & vbCrLf
               sSql = sSql & "                  ,SGI_DTULTMOV = '" & Format(Date, "MM/DD/YYYY") & "'" & vbCrLf
               sSql = sSql & "Where " & vbCrLf
               sSql = sSql & "      SGI_FILIAL    = " & FILIAL & vbCrLf
               sSql = sSql & "  And SGI_IDPRODUTO = " & varrITENSREQ(I, 1)
               
               BGRV.CommandText = sSql
               BGRV.Execute

           Next I
        End If
           
     ElseIf strOPER = "A" Or strOPER = "E" Then '' Altera/Exclui
           
        '' Extorna Qtde de Estoque para o produto
        If IsArray(varrITENSREQBACK) = True Then
           For I = 1 To UBound(varrITENSREQBACK)
              
                 '' --------------------------------------------------
                 '' Dá Entrada no Estoque
                 ''sSql = "Update SGI_PRODSALDOS " & vbCrLf
                 ''sSql = sSql & "               Set SGI_SALDO    = (SGI_SALDO - " & varrITENSREQBACK(I, 3) & ")" & vbCrLf
                 ''sSql = sSql & "Where " & vbCrLf
                 ''sSql = sSql & "      SGI_FILIAL     = " & FILIAL & vbCrLf
                 ''sSql = sSql & "  And SGI_IDPRODUTO  = " & varrITENSREQBACK(I, 1) & vbCrLf
                 ''sSql = sSql & "  And SGI_CODCLIENTE = " & lngCODCLIE
                 
                 ''BGRV.CommandText = sSql
                 ''BGRV.Execute
               
                 '' Dá Baixa no Estoque
                 sSql = "Update SGI_CADPRODUTO Set SGI_SALDO = (SGI_SALDO + " & varrITENSREQBACK(I, 3) & ")" & vbCrLf
                 sSql = sSql & "                  ,SGI_DTULTMOV = '" & Format(Date, "MM/DD/YYYY") & "'" & vbCrLf
                 sSql = sSql & "Where " & vbCrLf
                 sSql = sSql & "      SGI_FILIAL    = " & FILIAL & vbCrLf
                 sSql = sSql & "  And SGI_IDPRODUTO = " & varrITENSREQ(I, 1)
                 
                 BGRV.CommandText = sSql
                 BGRV.Execute
                           
                  
           Next I
        End If
        
        '' Apagando o cabeçalho da Requisição
        sSql = "Delete From SGI_CADITREQSAIMAT " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCADREQSAICOD
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Apagando os itens da Requisição
        sSql = "Delete From SGI_CADREQSAIMAT " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCADREQSAICOD
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        
        If strOPER = "A" Then GoTo Retorno
         
     End If
     
     adoBanco_Dados.CommitTrans
     
     GRAVA = True
     
     Exit Function

err_Trans:
    
    adoBanco_Dados.RollbackTrans

    Dim objErro    As Object
    Set objErro = CreateObject("BLBCWS.clsFuncoes")
    Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, strOPER, sSql)
    Set objErro = Nothing

End Function

Public Function Carrega_campos() As Boolean

   Carrega_campos = False
   
   Dim intNREQ As Integer
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADREQSAIMAT " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & lngCADREQSAICOD & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   
   If Not BREC.EOF Then
      
      lngCADREQSAICOD = BREC!SGI_CODIGO
      lngCADDEPCOD = BREC!SGI_CODDEPTO
      lngCADUSUCOD = BREC!SGI_CODUSUAR
      dtCADDTREQ = BREC!SGI_DATREQ
      strCODNF = BREC!SGI_CODNF
      lngCODMOTIVOS = BREC!SGI_CODMOTIVO
      If Not IsNull(BREC!SGI_CODCLIENTE) Then lngCODCLIE = BREC!SGI_CODCLIENTE
      
      Carrega_campos = True
   End If
   
   BREC.Close
   
   '' Carrega Itens --
   sSql = "Select " & vbCrLf
   sSql = sSql & "       REQ.* " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADITREQSAIMAT REQ" & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       REQ.SGI_FILIAL  = " & intFilial & vbCrLf
   sSql = sSql & "   And REQ.SGI_CODIGO  = " & lngCADREQSAICOD
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   
   intNREQ = 0
   Do While Not BREC2.EOF()
      intNREQ = intNREQ + 1
      BREC2.MoveNext
   Loop
   
   If intNREQ > 0 Then
      BREC2.MoveFirst
      ReDim varrITENSREQ(1 To intNREQ, 1 To 6) As String
      ReDim varrITENSREQBACK(1 To intNREQ, 1 To 6) As String
      
      intNREQ = 1
      Do While Not BREC2.EOF
         varrITENSREQ(intNREQ, 1) = BREC2!SGI_IDPRODUTO
         varrITENSREQ(intNREQ, 2) = BREC2!SGI_PRODUTO
         varrITENSREQ(intNREQ, 3) = BREC2!SGI_QTD
         If Not IsNull(BREC2!SGI_QTDEKG) Then varrITENSREQ(intNREQ, 4) = Format(BREC2!SGI_QTDEKG, "#,####0.0000")
         If Not IsNull(BREC2!SGI_CODLOTE) Then varrITENSREQ(intNREQ, 5) = BREC2!SGI_CODLOTE
         If Not IsNull(BREC2!SGI_INDICE) Then varrITENSREQ(intNREQ, 6) = BREC2!SGI_INDICE
         
         varrITENSREQBACK(intNREQ, 1) = BREC2!SGI_IDPRODUTO
         varrITENSREQBACK(intNREQ, 2) = BREC2!SGI_PRODUTO
         varrITENSREQBACK(intNREQ, 3) = BREC2!SGI_QTD
         If Not IsNull(BREC2!SGI_QTDEKG) Then varrITENSREQBACK(intNREQ, 4) = Format(BREC2!SGI_QTDEKG, "#,####0.0000")
         If Not IsNull(BREC2!SGI_CODLOTE) Then varrITENSREQBACK(intNREQ, 5) = BREC2!SGI_CODLOTE
         If Not IsNull(BREC2!SGI_INDICE) Then varrITENSREQBACK(intNREQ, 6) = BREC2!SGI_INDICE
         
         intNREQ = intNREQ + 1
         BREC2.MoveNext
      Loop
      
   End If
   
   BREC2.Close

End Function

Public Sub PreencheMotivoSaida(Combo As Variant)

     Combo.Clear
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       * " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADMOTIVOS" & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL  = " & intFilial & vbCrLf
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_DESCRI)
        Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
        BREC.MoveNext
     Loop
     BREC.Close
    
End Sub

