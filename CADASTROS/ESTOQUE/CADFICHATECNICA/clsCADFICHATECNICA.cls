VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADFICHATECNICA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial       As Integer
Dim intCODIGO       As Integer
Dim lngCODMAQ       As Long
Dim strCODPROD      As String
Dim dtData          As Date
Dim arrITENS        As Variant
Dim lngCODFAMMAQ    As Long
Dim curEFCMEDIA     As Currency
Dim curPRODPECTEOR  As Currency
Dim curPRODPECREAL  As Currency
Dim curMINUTOS      As Currency
Dim curSEGUNDOS     As Currency
Dim arrFAMUNIDADE   As Variant
Dim arrEFICIENCIA   As Variant
Dim lngUNIDSETTUP   As Long

Public Property Get UNIDSETTUP() As Long
    UNIDSETTUP = lngUNIDSETTUP
End Property
Public Property Let UNIDSETTUP(ByVal vData As Long)
    lngUNIDSETTUP = vData
End Property

Public Property Get EFICIENCIA() As Variant
    EFICIENCIA = arrEFICIENCIA
End Property
Public Property Let EFICIENCIA(ByVal vData As Variant)
    arrEFICIENCIA = vData
End Property

Public Property Get FAMUNIDADE() As Variant
    FAMUNIDADE = arrFAMUNIDADE
End Property
Public Property Let FAMUNIDADE(ByVal vData As Variant)
    arrFAMUNIDADE = vData
End Property

Public Property Get SEGUNDOS() As Currency
    SEGUNDOS = curSEGUNDOS
End Property
Public Property Let SEGUNDOS(ByVal vData As Currency)
    curSEGUNDOS = vData
End Property

Public Property Get MINUTOS() As Currency
    MINUTOS = curMINUTOS
End Property
Public Property Let MINUTOS(ByVal vData As Currency)
    curMINUTOS = vData
End Property

Public Property Get PRODPECREAL() As Currency
    PRODPECREAL = curPRODPECREAL
End Property
Public Property Let PRODPECREAL(ByVal vData As Currency)
    curPRODPECREAL = vData
End Property

Public Property Get PRODPECTEOR() As Currency
    PRODPECTEOR = curPRODPECTEOR
End Property
Public Property Let PRODPECTEOR(ByVal vData As Currency)
    curPRODPECTEOR = vData
End Property

Public Property Get EFCMEDIA() As Currency
    EFCMEDIA = curEFCMEDIA
End Property
Public Property Let EFCMEDIA(ByVal vData As Currency)
    curEFCMEDIA = vData
End Property

Public Property Get CODFAMMAQ() As Variant
    CODFAMMAQ = lngCODFAMMAQ
End Property
Public Property Let CODFAMMAQ(ByVal vData As Variant)
    lngCODFAMMAQ = vData
End Property

Public Property Get ITENS() As Variant
    ITENS = arrITENS
End Property
Public Property Let ITENS(ByVal vData As Variant)
    arrITENS = vData
End Property

Public Property Get Data() As Date
    Data = dtData
End Property
Public Property Let Data(ByVal vData As Date)
    dtData = vData
End Property

Public Property Get CODPROD() As String
    CODPROD = strCODPROD
End Property
Public Property Let CODPROD(ByVal vData As String)
    strCODPROD = vData
End Property

Public Property Get CODMAQ() As Long
    CODMAQ = lngCODMAQ
End Property
Public Property Let CODMAQ(ByVal vData As Long)
    lngCODMAQ = vData
End Property

Public Property Get CODIGO() As Integer
    CODIGO = intCODIGO
End Property
Public Property Let CODIGO(ByVal vData As Integer)
    intCODIGO = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property
Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADFICHATECNICAP.cCaminho = cCaminho
    frmCADFICHATECNICAP.Linha = Linha
    frmCADFICHATECNICAP.FILIAL = iFilial
    frmCADFICHATECNICAP.strAcesso = strAcesso
    frmCADFICHATECNICAP.Show vbModal
End Sub

Public Function Pesq_CadFichaTec() As Boolean

   Pesq_CadFichaTec = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       * " & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_CADFICHATECHEAD " & vbCrLf
   sSql = sSql + " Where " & vbCrLf
   sSql = sSql + "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_CadFichaTec = True
   BREC.Close

End Function

Public Sub PreenchComboMaquina(Combo As Variant, lngCODFAM As Long)

    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADMAQUINA " & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "  And SGI_CODFAMILIA = " & lngCODFAM & vbCrLf
    sSql = sSql & "Order by SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       Combo.AddItem Trim(BREC!SGI_DESCRI)
       Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
       BREC.MoveNext
    Loop
    
    BREC.Close

End Sub

Public Sub PreencheComboProd(Combo As Variant)

     Combo.Clear
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       PRODUTO.SGI_CODIGO " & vbCrLf
     sSql = sSql & "      ,PRODUTO.SGI_DESCRICAO " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADPRODUTO PRODUTO " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       PRODUTO.SGI_FILIAL        = " & intFilial & vbCrLf
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_CODIGO) & Space(10 - Len(Trim(BREC!SGI_CODIGO))) & " - " & Trim(BREC!SGI_DESCRICAO)
        BREC.MoveNext
     Loop
    
     BREC.Close
    
End Sub

Public Function PreenchComboUnidadeGrid() As String
        
    PreenchComboUnidadeGrid = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADUNIMED " & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "Order by SGI_CODIGO"
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC.EOF
       PreenchComboUnidadeGrid = PreenchComboUnidadeGrid & "|#" & BREC!SGI_CODIGO & ";" & Trim(BREC!SGI_UNIDADE)
       BREC.MoveNext
    Loop
    BREC.Close
End Function

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

Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_grava

     GRAVA = False
     
     Dim sValor As String
     Dim I      As Integer
         
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     If strOPER = "I" Then
        
Volta:
        '' Grava Familia de Máquinas
        sSql = "Insert Into SGI_CADFICHATECHEAD (SGI_FILIAL,SGI_CODIGO,SGI_DATA,SGI_CODMAQ,SGI_CDPROD,SGI_CADFAMMAQ,SGI_EFCMEDIA,SGI_PRODPECTEOR,SGI_PRODPECREAL,SGI_SETUPMIN,SGI_SETUPSEG,SGI_CODUNIDMED) Values (" & vbCrLf
        sSql = sSql & "                            " & intFilial & vbCrLf
        sSql = sSql & "                           ," & intCODIGO & vbCrLf
        sSql = sSql & "                           ,'" & Format(dtData, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                           ," & lngCODMAQ & vbCrLf
        sSql = sSql & "                           ,'" & Trim(strCODPROD) & "'" & vbCrLf
        sSql = sSql & "                           ," & lngCODFAMMAQ & vbCrLf
        
        sValor = Replace(curEFCMEDIA, ".", "")
        sValor = Replace(Trim(sValor), ",", ".")
        sSql = sSql & "                           , " & sValor & vbCrLf

        sValor = Replace(curPRODPECTEOR, ".", "")
        sValor = Replace(Trim(sValor), ",", ".")
        sSql = sSql & "                           , " & sValor & vbCrLf
        
        sValor = Replace(curPRODPECREAL, ".", "")
        sValor = Replace(Trim(sValor), ",", ".")
        sSql = sSql & "                           , " & sValor & vbCrLf
        
        sValor = Replace(curMINUTOS, ".", "")
        sValor = Replace(Trim(sValor), ",", ".")
        sSql = sSql & "                           , " & sValor & vbCrLf
        
        sValor = Replace(curSEGUNDOS, ".", "")
        sValor = Replace(Trim(sValor), ",", ".")
        sSql = sSql & "                           , " & sValor & vbCrLf
        sSql = sSql & "                           , " & lngUNIDSETTUP & vbCrLf
        
        sSql = sSql & "                           )"
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Parâmetros da Ficha Técnica
        If IsArray(arrITENS) Then
            For I = 1 To UBound(arrITENS)
                sSql = "Insert Into SGI_CADFICHATECITEN (" & vbCrLf
                sSql = sSql & "                                 SGI_FILIAL" & vbCrLf
                sSql = sSql & "                                ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                                ,SGI_CODPAR" & vbCrLf
                sSql = sSql & "                                ,SGI_CODUNIMED" & vbCrLf
                sSql = sSql & "                                ,SGI_VALPAD" & vbCrLf
                sSql = sSql & "                                ,SGI_VALPADPOS" & vbCrLf
                sSql = sSql & "                                ,SGI_VALPADNEG" & vbCrLf
                sSql = sSql & "                                )" & vbCrLf
                sSql = sSql & "                         Values (" & vbCrLf
                sSql = sSql & "                                 " & intFilial & vbCrLf
                sSql = sSql & "                                ," & intCODIGO & vbCrLf
                sSql = sSql & "                                ," & arrITENS(I, 1) & vbCrLf
                sSql = sSql & "                                ," & arrITENS(I, 2) & vbCrLf
                
                sValor = Replace(arrITENS(I, 3), ".", "")
                sValor = Replace(Trim(sValor), ",", ".")
                sSql = sSql & "                           , " & sValor & vbCrLf
                
                sValor = Replace(arrITENS(I, 4), ".", "")
                sValor = Replace(Trim(sValor), ",", ".")
                sSql = sSql & "                           , " & sValor & vbCrLf
                
                sValor = Replace(arrITENS(I, 5), ".", "")
                sValor = Replace(Trim(sValor), ",", ".")
                sSql = sSql & "                           , " & sValor & vbCrLf
                
                sSql = sSql & "                                )"
            
                BGRV.CommandText = sSql
                BGRV.Execute
            
            Next I
        End If
        
        '' Familia de Unidade de Medidas
        If IsArray(arrFAMUNIDADE) Then
           For I = 1 To UBound(arrFAMUNIDADE)
                sSql = "Insert Into SGI_CADFICHAFAMUNID (" & vbCrLf
                sSql = sSql & "                                 SGI_FILIAL" & vbCrLf
                sSql = sSql & "                                ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                                ,SGI_CODUNIDDE" & vbCrLf
                sSql = sSql & "                                ,SGI_CODUNIDPARA" & vbCrLf
                sSql = sSql & "                                ,SGI_DAFAULT" & vbCrLf
                sSql = sSql & "                                ,SGI_INDICE)" & vbCrLf
                sSql = sSql & "                         Values (" & vbCrLf
                sSql = sSql & "                                 " & intFilial & vbCrLf
                sSql = sSql & "                                ," & intCODIGO & vbCrLf
                sSql = sSql & "                                ," & arrFAMUNIDADE(I, 1) & vbCrLf
                sSql = sSql & "                                ," & arrFAMUNIDADE(I, 2) & vbCrLf
                sSql = sSql & "                                ," & arrFAMUNIDADE(I, 3) & vbCrLf
                sSql = sSql & "                                ," & arrFAMUNIDADE(I, 4) & vbCrLf
                sSql = sSql & "                                )"
            
                BGRV.CommandText = sSql
                BGRV.Execute
           Next I
        End If
        
        '' Eficiências
        If IsArray(arrEFICIENCIA) Then
           For I = 1 To UBound(arrEFICIENCIA)
                sSql = "Insert Into SGI_CADFICHAEFIC (" & vbCrLf
                sSql = sSql & "                                 SGI_FILIAL" & vbCrLf
                sSql = sSql & "                                ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                                ,SGI_INDICE" & vbCrLf
                sSql = sSql & "                                ,SGI_EFICMED" & vbCrLf
                sSql = sSql & "                                ,SGI_PRODTEO" & vbCrLf
                sSql = sSql & "                                ,SGI_PRODREA)" & vbCrLf
                sSql = sSql & "                         Values (" & vbCrLf
                sSql = sSql & "                                 " & intFilial & vbCrLf
                sSql = sSql & "                                ," & intCODIGO & vbCrLf
                sSql = sSql & "                                ," & arrEFICIENCIA(I, 4) & vbCrLf
                
                sValor = Replace(arrEFICIENCIA(I, 1), ".", "")
                sValor = Replace(Trim(sValor), ",", ".")
                sSql = sSql & "                                ," & sValor & vbCrLf
                
                sValor = Replace(arrEFICIENCIA(I, 2), ".", "")
                sValor = Replace(Trim(sValor), ",", ".")
                sSql = sSql & "                                ," & sValor & vbCrLf
                
                sValor = Replace(arrEFICIENCIA(I, 3), ".", "")
                sValor = Replace(Trim(sValor), ",", ".")
                sSql = sSql & "                                ," & sValor & vbCrLf
                sSql = sSql & "                                )"
            
                BGRV.CommandText = sSql
                BGRV.Execute
           Next I
        End If
        
     ElseIf strOPER = "E" Or strOPER = "A" Then
        
        '' Eficiencia
        sSql = "Delete From SGI_CADFICHAEFIC " & vbCrLf
        sSql = sSql & "            Where SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "              And SGI_CODIGO = " & intCODIGO

        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Familia de Unidade de Medidas
        sSql = "Delete From SGI_CADFICHAFAMUNID " & vbCrLf
        sSql = sSql & "            Where SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "              And SGI_CODIGO = " & intCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Apaga Ficha Técnica
        sSql = "Delete from SGI_CADFICHATECITEN " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
                
        sSql = "Delete from SGI_CADFICHATECHEAD " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        If strOPER = "A" Then GoTo Volta
     
     End If
     
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

Public Function Carrega_campos() As Boolean

   Dim lngTOTREG As Long
   
   Carrega_campos = False
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADFICHATECHEAD " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & intCODIGO & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then
      
      intCODIGO = BREC!SGI_CODIGO
      dtData = BREC!SGI_DATA
      lngCODMAQ = BREC!SGI_CODMAQ
      strCODPROD = BREC!SGI_CDPROD
      lngCODFAMMAQ = BREC!SGI_CADFAMMAQ
      curEFCMEDIA = BREC!SGI_EFCMEDIA
      curPRODPECTEOR = BREC!SGI_PRODPECTEOR
      curPRODPECREAL = BREC!SGI_PRODPECREAL
      
      If Not IsNull(BREC!SGI_SETUPMIN) Then curMINUTOS = BREC!SGI_SETUPMIN
      If Not IsNull(BREC!SGI_SETUPSEG) Then curSEGUNDOS = BREC!SGI_SETUPSEG
      If Not IsNull(BREC!SGI_CODUNIDMED) Then lngUNIDSETTUP = BREC!SGI_CODUNIDMED
      
      Carrega_campos = True
   End If
   BREC.Close
   
   '' ===================================================
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADFICHATECITEN " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & intCODIGO & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then
      lngTOTREG = 0
      Do While Not BREC.EOF
         lngTOTREG = lngTOTREG + 1
         BREC.MoveNext
      Loop
      
      ReDim arrITENS(1 To lngTOTREG, 1 To 5) As Variant
      
      lngTOTREG = 1
      BREC.MoveFirst
      Do While Not BREC.EOF
         arrITENS(lngTOTREG, 1) = BREC!SGI_CODPAR
         arrITENS(lngTOTREG, 2) = BREC!SGI_CODUNIMED
         arrITENS(lngTOTREG, 3) = BREC!SGI_VALPAD
         arrITENS(lngTOTREG, 4) = BREC!SGI_VALPADPOS
         arrITENS(lngTOTREG, 5) = BREC!SGI_VALPADNEG
         lngTOTREG = lngTOTREG + 1
         BREC.MoveNext
      Loop
   End If
   BREC.Close
   
   
   '' ===================================================
   '' Familia de Unidades
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADFICHAFAMUNID " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & intCODIGO & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then
      lngTOTREG = 0
      Do While Not BREC.EOF
         lngTOTREG = lngTOTREG + 1
         BREC.MoveNext
      Loop
      
      ReDim arrFAMUNIDADE(1 To lngTOTREG, 1 To 3) As Variant
      
      lngTOTREG = 1
      BREC.MoveFirst
      Do While Not BREC.EOF
         arrFAMUNIDADE(lngTOTREG, 1) = BREC!SGI_CODUNIDDE
         arrFAMUNIDADE(lngTOTREG, 2) = BREC!SGI_CODUNIDPARA
         arrFAMUNIDADE(lngTOTREG, 3) = BREC!SGI_DAFAULT
         lngTOTREG = lngTOTREG + 1
         BREC.MoveNext
      Loop
   End If
   BREC.Close
   
   
   '' ===================================================
   '' Familia de Unidades
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADFICHAEFIC " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & intCODIGO & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then
      lngTOTREG = 0
      Do While Not BREC.EOF
         lngTOTREG = lngTOTREG + 1
         BREC.MoveNext
      Loop
      
      ReDim arrEFICIENCIA(1 To lngTOTREG, 1 To 4) As Variant
      
      lngTOTREG = 1
      BREC.MoveFirst
      Do While Not BREC.EOF
         arrEFICIENCIA(lngTOTREG, 1) = BREC!SGI_EFICMED
         arrEFICIENCIA(lngTOTREG, 2) = BREC!SGI_PRODTEO
         arrEFICIENCIA(lngTOTREG, 3) = BREC!SGI_PRODREA
         arrEFICIENCIA(lngTOTREG, 4) = BREC!SGI_INDICE
         lngTOTREG = lngTOTREG + 1
         BREC.MoveNext
      Loop
   End If
   BREC.Close
   
End Function

Public Sub PreenchComboFamMaquina(Combo As Variant, strCODPROD As String)

    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       FM.*  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO PR " & vbCrLf
    sSql = sSql & "      ,SGI_CADFAMMAQUINAS FM " & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      PR.SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "  And PR.SGI_CODIGO = '" & strCODPROD & "'" & vbCrLf
    sSql = sSql & "  And FM.SGI_FILIAL = PR.SGI_FILIAL " & vbCrLf
    sSql = sSql & "  And FM.SGI_CODIGO = PR.SGI_CADFAMMAQ "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       Combo.AddItem Trim(BREC!SGI_DESCRI)
       Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
       BREC.MoveNext
    Loop
    
    BREC.Close

End Sub

Public Sub PreenchComboUnidade(Combo As Variant)
    Combo.Clear
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADUNIMED " & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "Order by SGI_CODIGO"
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC.EOF
       Combo.AddItem Trim(BREC!SGI_UNIDADE)
       Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
       BREC.MoveNext
    Loop
    '' If Combo.ListCount > 0 Then Combo.ListIndex = 0
    BREC.Close
End Sub

