VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADREQMAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim intFilial     As Integer
Dim lngCADREQCOD  As Long
Dim lngCADDEPCOD  As Long
Dim lngCADUSUCOD  As Long
Dim dtCADDTREQ    As Date
Dim varrITENSREQ  As Variant
Dim objFuncoes    As Object

Public Property Get ITENSREQ() As Variant
    ITENSREQ = varrITENSREQ
End Property

Public Property Let ITENSREQ(ByVal vData As Variant)
    varrITENSREQ = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Property Get CADREQCOD() As Long
    CADREQCOD = lngCADREQCOD
End Property

Public Property Let CADREQCOD(ByVal vData As Long)
    lngCADREQCOD = vData
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
    frmCADREQMATP.cCaminho = cCaminho
    frmCADREQMATP.Linha = Linha
    frmCADREQMATP.FILIAL = iFilial
    frmCADREQMATP.strAcesso = strAcesso
    frmCADREQMATP.strUSUARIO = strUSUARIO
    frmCADREQMATP.Show vbModal
End Sub


Public Function Pesq_CadReqMat() As Boolean

   Pesq_CadReqMat = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       *" & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_CADREQMAT " & vbCrLf
   sSql = sSql + " Where "
   sSql = sSql + "       SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_CadReqMat = True
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
    
     ''If Combo.ListCount > 0 Then Combo.ListIndex = 0
    
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
        sSql = "Insert Into SGI_CADREQMAT (SGI_FILIAL" & vbCrLf
        sSql = sSql & "                   ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                   ,SGI_DATREQ" & vbCrLf
        sSql = sSql & "                   ,SGI_CODDEPTO" & vbCrLf
        sSql = sSql & "                   ,SGI_CODUSUAR)" & vbCrLf
        sSql = sSql & "                Values (" & vbCrLf
        sSql = sSql & "                            " & intFilial & vbCrLf
        sSql = sSql & "                           ," & lngCADREQCOD & vbCrLf
        sSql = sSql & "                           ,'" & Format(dtCADDTREQ, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                           , " & lngCADDEPCOD & vbCrLf
        sSql = sSql & "                           , " & lngCADUSUCOD & vbCrLf
        sSql = sSql & "                           )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Grava os Itens da Requisição
        If IsArray(varrITENSREQ) = True Then
           For I = 1 To UBound(varrITENSREQ)
               
               sSql = "Insert Into SGI_CADITREQMAT  (SGI_FILIAL" & vbCrLf
               sSql = sSql & "                      ,SGI_CODIGO" & vbCrLf
               sSql = sSql & "                      ,SGI_PRODUTO" & vbCrLf
               sSql = sSql & "                      ,SGI_QTD" & vbCrLf
               sSql = sSql & "                      ,SGI_QTDATEND)" & vbCrLf
               sSql = sSql & "               Values (" & vbCrLf
               sSql = sSql & "                       " & intFilial & vbCrLf
               sSql = sSql & "                      ," & lngCADREQCOD & vbCrLf
               sSql = sSql & "                      ,'" & varrITENSREQ(I, 1) & "'" & vbCrLf
                              
               sValor = Replace(varrITENSREQ(I, 4), ".", "")
               sValor = Replace(Trim(sValor), ",", ".")
               sSql = sSql & "                            ," & sValor & vbCrLf
               
               sSql = sSql & "                            ,0)" & vbCrLf
               
               BGRV.CommandText = sSql
               BGRV.Execute
               
           Next I
        End If
           
     ElseIf strOPER = "A" Then '' Altera
        
        '' Apagando os itens da Requisição
        sSql = "Delete From SGI_CADREQMAT " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCADREQCOD
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Apagando o cabeçalho da Requisição
        sSql = "Delete From SGI_CADITREQMAT " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCADREQCOD
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        GoTo Retorno
         
     ElseIf strOPER = "E" Then '' Exclui
     
        '' Apagando o cabeçalho da Requisição
        sSql = "Delete From SGI_CADITREQMAT " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCADREQCOD
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Apagando os itens da Requisição
        sSql = "Delete From SGI_CADREQMAT " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCADREQCOD
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
     
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
   sSql = sSql & "       SGI_CADREQMAT " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & lngCADREQCOD & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   
   If Not BREC.EOF Then
      
      lngCADREQCOD = BREC!SGI_CODIGO
      lngCADDEPCOD = BREC!SGI_CODDEPTO
      lngCADUSUCOD = BREC!SGI_CODUSUAR
      dtCADDTREQ = BREC!SGI_DATREQ
      
      Carrega_campos = True
   End If
   
   BREC.Close
   
   '' Carrega Itens --
   sSql = "Select " & vbCrLf
   sSql = sSql & "       REQ.* " & vbCrLf
   sSql = sSql & "      ,PRO.SGI_DESCRICAO " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADITREQMAT REQ" & vbCrLf
   sSql = sSql & "      ,SGI_CADPRODUTO  PRO" & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       REQ.SGI_FILIAL  = " & intFilial & vbCrLf
   sSql = sSql & "   And REQ.SGI_CODIGO  = " & lngCADREQCOD
   sSql = sSql & "   And PRO.SGI_FILIAL  = REQ.SGI_FILIAL "
   sSql = sSql & "   And PRO.SGI_CODIGO  = REQ.SGI_PRODUTO"
   
   BREC2.Open sSql, adoBanco_Dados
   
   intNREQ = 0
   Do While Not BREC2.EOF
      intNREQ = intNREQ + 1
      BREC2.MoveNext
   Loop
   
   If intNREQ > 0 Then
         
      BREC2.MoveFirst
      ReDim varrITENSREQ(1 To intNREQ, 1 To 6) As Variant
      intNREQ = 1
      Do While Not BREC2.EOF
         varrITENSREQ(intNREQ, 1) = BREC2!SGI_PRODUTO
         varrITENSREQ(intNREQ, 2) = BREC2!SGI_PRODUTO
         varrITENSREQ(intNREQ, 3) = BREC2!SGI_DESCRICAO
         varrITENSREQ(intNREQ, 4) = BREC2!SGI_QTD
         varrITENSREQ(intNREQ, 5) = BREC2!SGI_QTDATEND
         varrITENSREQ(intNREQ, 6) = (BREC2!SGI_QTD - BREC2!SGI_QTDATEND)
         intNREQ = intNREQ + 1
         BREC2.MoveNext
      Loop
      
   End If
   
   BREC2.Close

End Function


