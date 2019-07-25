VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADPEDVENDA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFILIAL           As Integer
Dim lngCODPEDIDO        As Long
Dim lngCODCOTA          As Long
Dim dtDATAPED           As Date
Dim dtDATCALIBR         As Date
Dim strSTATUS           As String
Dim intCODCLIE          As Integer
Dim intCODCONDPGTO      As Integer
Dim intCODVEND          As Integer
Dim intTIPPED           As Integer
Dim strENDENTR          As String
Dim strBAIENTR          As String
Dim strCIDENTR          As String
Dim strTELENTR          As String
Dim strFAXENTR          As String
Dim intESTENTREGA       As Integer
Dim strCEPENTREGA       As String
Dim strENDCOBRA         As String
Dim strBAICOBRA         As String
Dim strCIDCOBRA         As String
Dim intESTCOBRA         As Integer
Dim strCEPCOBRA         As String
Dim strTELCOBRA         As String
Dim strFAXCOBRA         As String
Dim arrPRODUTOS         As Variant
Dim arrPRODCOPIA        As Variant
Dim arrDESPESAS         As Variant
Dim arrSERVICOS         As Variant
Dim arrTIPREPROVA       As Variant
Dim intPRZENTREGA       As Integer
Dim intVALPROPOSTA      As Integer
Dim intIPIISENTO        As Integer
Dim intICMSISENTO       As Integer
Dim intFRETE            As Integer
Dim curVALBASICMS       As Currency
Dim curAliqICMS         As Currency
Dim curVLICMS           As Currency
Dim curOUTRDESPESAS     As Currency
Dim curVLFRETE          As Currency
Dim curVLIPI            As Currency
Dim curVLDESCTO         As Currency
Dim curTOTORCTO         As Currency
Dim curVALDESC          As Currency
Dim curPORDESC          As Currency
Dim intTOTALITENS       As Integer
Dim strLIBERADO         As String
Dim dtDTHORALIB         As Date
Dim strOBS              As String
Dim strOBSCOM           As String
Dim strREPROVADO        As String
Dim dtDTHORAREPRO       As Date
Dim lngQTDTOTCOTA       As Long
Dim lngQTDPEDATEND      As Long
Dim boolTEMPEDIDO       As Boolean
Dim dtDATAENTREGA       As Date
Dim lngCODTRANSP        As Long
Dim strORDCOMPCLI       As String
Dim strCONTATO          As String
Dim strDEPARTAMENTO     As String
Dim strEMAIL            As String
Dim intCODTECNICO       As Integer
Dim strOBSERVACAO       As String
Dim lngCODTIPODOC       As Long
Dim lngCODBANCO         As Long
Dim arrPROGENTREGAS     As Variant
Dim curQTDITENSPROD     As Currency
Dim arrOPS              As Variant
Dim strNomeVend         As String
Dim strOBS2             As String
Dim intFILIALPED        As Integer
Dim intESPECIAL         As Integer
Dim intPARAESTOQUE      As Integer
Dim boolFOIREPROVADO    As Boolean
Dim intCONFERIDO        As Integer
Dim strVERSAO           As String
Dim strNOMCOMP          As String
Dim lngCodUsuario       As Long
Dim arrLOGPEDIDOS       As Variant
Dim strCODMOTLIQ        As String
Dim strOBSLIQ           As String
Dim intPERMITEFECHOP    As Integer

Public Property Get PERMITEFECHOP() As Integer
    PERMITEFECHOP = intPERMITEFECHOP
End Property

Public Property Let PERMITEFECHOP(ByVal vData As Integer)
    intPERMITEFECHOP = vData
End Property

Public Property Get OBSLIQ() As String
    OBSLIQ = strOBSLIQ
End Property

Public Property Let OBSLIQ(ByVal vData As String)
    strOBSLIQ = vData
End Property

Public Property Get CODMOTLIQ() As String
    CODMOTLIQ = strCODMOTLIQ
End Property

Public Property Let CODMOTLIQ(ByVal vData As String)
    strCODMOTLIQ = vData
End Property

Public Property Get LOGPEDIDOS() As Variant
    LOGPEDIDOS = arrLOGPEDIDOS
End Property

Public Property Let LOGPEDIDOS(ByVal vData As Variant)
    arrLOGPEDIDOS = vData
End Property

Public Property Get CODUSUARIO() As Long
    CODUSUARIO = lngCodUsuario
End Property

Public Property Let CODUSUARIO(ByVal vData As Long)
    lngCodUsuario = vData
End Property

Public Property Get NOMCOMP() As String
    NOMCOMP = strNOMCOMP
End Property

Public Property Let NOMCOMP(ByVal vData As String)
    strNOMCOMP = vData
End Property

Public Property Get VERSAO() As String
    VERSAO = strVERSAO
End Property

Public Property Let VERSAO(ByVal vData As String)
    strVERSAO = vData
End Property

Public Property Get CONFERIDO() As Integer
    CONFERIDO = intCONFERIDO
End Property

Public Property Let CONFERIDO(ByVal vData As Integer)
    intCONFERIDO = vData
End Property

Public Property Get FOIREPROVADO() As Boolean
    FOIREPROVADO = boolFOIREPROVADO
End Property

Public Property Let FOIREPROVADO(ByVal vData As Boolean)
    boolFOIREPROVADO = vData
End Property

Public Property Get PARAESTOQUE() As Integer
    PARAESTOQUE = intPARAESTOQUE
End Property

Public Property Let PARAESTOQUE(ByVal vData As Integer)
    intPARAESTOQUE = vData
End Property

Public Property Get ESPECIAL() As Integer
    ESPECIAL = intESPECIAL
End Property

Public Property Let ESPECIAL(ByVal vData As Integer)
    intESPECIAL = vData
End Property

Public Property Get FILIALPED() As Integer
    FILIALPED = intFILIALPED
End Property

Public Property Let FILIALPED(ByVal vData As Integer)
    intFILIALPED = vData
End Property

Public Property Get OBS2() As String
    OBS2 = strOBS2
End Property

Public Property Let OBS2(ByVal vData As String)
    strOBS2 = vData
End Property

Public Property Get NOMEVEND() As String
    NOMEVEND = strNomeVend
End Property

Public Property Let NOMEVEND(ByVal vData As String)
    strNomeVend = vData
End Property

Public Property Get OPS() As Variant
    OPS = arrOPS
End Property

Public Property Let OPS(ByVal vData As Variant)
    arrOPS = vData
End Property

Public Property Get QTDITENSPROD() As Currency
    QTDITENSPROD = curQTDITENSPROD
End Property

Public Property Let QTDITENSPROD(ByVal vData As Currency)
    curQTDITENSPROD = vData
End Property

Public Property Get PROGENTREGAS() As Variant
    PROGENTREGAS = arrPROGENTREGAS
End Property

Public Property Let PROGENTREGAS(ByVal vData As Variant)
    arrPROGENTREGAS = vData
End Property

Public Property Get PORDESC() As Currency
    PORDESC = curPORDESC
End Property

Public Property Let PORDESC(ByVal vData As Currency)
    curPORDESC = vData
End Property

Public Property Get VALDESC() As Currency
    VALDESC = curVALDESC
End Property

Public Property Let VALDESC(ByVal vData As Currency)
    curVALDESC = vData
End Property

Public Property Get CODBANCO() As Long
    CODBANCO = lngCODBANCO
End Property

Public Property Let CODBANCO(ByVal vData As Long)
    lngCODBANCO = vData
End Property

Public Property Get CODTIPODOC() As Long
    CODTIPODOC = lngCODTIPODOC
End Property

Public Property Let CODTIPODOC(ByVal vData As Long)
    lngCODTIPODOC = vData
End Property


Public Property Get OBSERVACAO() As String
    OBSERVACAO = strOBSERVACAO
End Property

Public Property Let OBSERVACAO(ByVal vData As String)
    strOBSERVACAO = vData
End Property

Public Property Get OBSCOMERCIAL() As String
    OBSCOMERCIAL = strOBSCOM
End Property

Public Property Let OBSCOMERCIAL(ByVal vData As String)
    strOBSCOM = vData
End Property

Public Property Get CODTECNICO() As Integer
    CODTECNICO = intCODTECNICO
End Property

Public Property Let CODTECNICO(ByVal vData As Integer)
    intCODTECNICO = vData
End Property

Public Property Get EMAIL() As String
    EMAIL = strEMAIL
End Property

Public Property Let EMAIL(ByVal vData As String)
    strEMAIL = vData
End Property

Public Property Get DATCALIBR() As Date
    DATCALIBR = dtDATCALIBR
End Property

Public Property Let DATCALIBR(ByVal vData As Date)
    dtDATCALIBR = vData
End Property

Public Property Get DEPARTAMENTO() As String
    DEPARTAMENTO = strDEPARTAMENTO
End Property

Public Property Let DEPARTAMENTO(ByVal vData As String)
    strDEPARTAMENTO = vData
End Property

Public Property Get CONTATO() As String
    CONTATO = strCONTATO
End Property

Public Property Let CONTATO(ByVal vData As String)
    strCONTATO = vData
End Property

Public Property Get ORDCOMPCLI() As String
    ORDCOMPCLI = strORDCOMPCLI
End Property

Public Property Let ORDCOMPCLI(ByVal vData As String)
    strORDCOMPCLI = vData
End Property

Public Property Get TELCOBRA() As String
    TELCOBRA = strTELCOBRA
End Property

Public Property Let TELCOBRA(ByVal vData As String)
    strTELCOBRA = vData
End Property

Public Property Get FAXCOBRA() As String
    FAXCOBRA = strFAXCOBRA
End Property

Public Property Let FAXCOBRA(ByVal vData As String)
    strFAXCOBRA = vData
End Property

Public Property Get TELENTR() As String
    TELENTR = strTELENTR
End Property

Public Property Let TELENTR(ByVal vData As String)
    strTELENTR = vData
End Property

Public Property Get FAXENTR() As String
    FAXENTR = strFAXENTR
End Property

Public Property Let FAXENTR(ByVal vData As String)
    strFAXENTR = vData
End Property

Public Property Get CODTRANSP() As Long
    CODTRANSP = lngCODTRANSP
End Property

Public Property Let CODTRANSP(ByVal vData As Long)
    lngCODTRANSP = vData
End Property

Public Property Get DATAENTREGA() As Date
    DATAENTREGA = dtDATAENTREGA
End Property

Public Property Let DATAENTREGA(ByVal vData As Date)
    dtDATAENTREGA = vData
End Property

Public Property Get TEMPEDIDO() As Boolean
    TEMPEDIDO = boolTEMPEDIDO
End Property

Public Property Let TEMPEDIDO(ByVal vData As Boolean)
    boolTEMPEDIDO = vData
End Property

Public Property Get QTDPEDATEND() As Long
    QTDPEDATEND = lngQTDPEDATEND
End Property

Public Property Let QTDPEDATEND(ByVal vData As Long)
    lngQTDPEDATEND = vData
End Property

Public Property Get QTDTOTCOTA() As Long
    QTDTOTCOTA = lngQTDTOTCOTA
End Property

Public Property Let QTDTOTCOTA(ByVal vData As Long)
    lngQTDTOTCOTA = vData
End Property

Public Property Get DTHORAREPRO() As Date
    DTHORAREPRO = dtDTHORAREPRO
End Property

Public Property Let DTHORAREPRO(ByVal vData As Date)
    dtDTHORAREPRO = vData
End Property

Public Property Get REPROVADO() As String
    REPROVADO = strREPROVADO
End Property

Public Property Let REPROVADO(ByVal vData As String)
    strREPROVADO = vData
End Property

Public Property Get TIPREPROVA() As Variant
    TIPREPROVA = arrTIPREPROVA
End Property

Public Property Let TIPREPROVA(ByVal vData As Variant)
    arrTIPREPROVA = vData
End Property

Public Property Get OBS() As String
    OBS = strOBS
End Property

Public Property Let OBS(ByVal vData As String)
    strOBS = vData
End Property

Public Property Get DTHORALIB() As Date
    DTHORALIB = dtDTHORALIB
End Property

Public Property Let DTHORALIB(ByVal vData As Date)
    dtDTHORALIB = vData
End Property

Public Property Get LIBERADO() As String
    LIBERADO = strLIBERADO
End Property

Public Property Let LIBERADO(ByVal vData As String)
    strLIBERADO = vData
End Property

Public Property Get PRODCOPIA() As Variant
    PRODCOPIA = arrPRODCOPIA
End Property

Public Property Let PRODCOPIA(ByVal vData As Variant)
    arrPRODCOPIA = vData
End Property

Public Property Get CODCONDPGTO() As Integer
    CODCONDPGTO = intCODCONDPGTO
End Property

Public Property Let CODCONDPGTO(ByVal vData As Integer)
    intCODCONDPGTO = vData
End Property

Public Property Get TOTALITENS() As Integer
    TOTALITENS = intTOTALITENS
End Property

Public Property Let TOTALITENS(ByVal vData As Integer)
    intTOTALITENS = vData
End Property

Public Property Get STATUS() As String
    STATUS = strSTATUS
End Property

Public Property Let STATUS(ByVal vData As String)
    strSTATUS = vData
End Property

Public Property Get DATAPED() As Date
    DATAPED = dtDATAPED
End Property

Public Property Let DATAPED(ByVal vData As Date)
    dtDATAPED = vData
End Property

Public Property Get CODCOTA() As Long
    CODCOTA = lngCODCOTA
End Property

Public Property Let CODCOTA(ByVal vData As Long)
    lngCODCOTA = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFILIAL
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFILIAL = vData
End Property

Public Property Get CODPEDIDO() As Long
    CODPEDIDO = lngCODPEDIDO
End Property

Public Property Let CODPEDIDO(ByVal vData As Long)
    lngCODPEDIDO = vData
End Property

Public Property Get CODCLIE() As Integer
    CODCLIE = intCODCLIE
End Property

Public Property Let CODCLIE(ByVal vData As Integer)
    intCODCLIE = vData
End Property

Public Property Get CODVEND() As Integer
    CODVEND = intCODVEND
End Property

Public Property Let CODVEND(ByVal vData As Integer)
    intCODVEND = vData
End Property

Public Property Get TIPPED() As Integer
    TIPPED = intTIPPED
End Property

Public Property Let TIPPED(ByVal vData As Integer)
    intTIPPED = vData
End Property

Public Property Get ENDENTR() As String
    ENDENTR = strENDENTR
End Property

Public Property Let ENDENTR(ByVal vData As String)
    strENDENTR = vData
End Property

Public Property Get BAIENTR() As String
    BAIENTR = strBAIENTR
End Property

Public Property Let BAIENTR(ByVal vData As String)
    strBAIENTR = vData
End Property

Public Property Get CIDENTR() As String
    CIDENTR = strCIDENTR
End Property

Public Property Let CIDENTR(ByVal vData As String)
    strCIDENTR = vData
End Property

Public Property Get ESTENTREGA() As Integer
    ESTENTREGA = intESTENTREGA
End Property

Public Property Let ESTENTREGA(ByVal vData As Integer)
    intESTENTREGA = vData
End Property

Public Property Get CEPENTREGA() As String
    CEPENTREGA = strCEPENTREGA
End Property

Public Property Let CEPENTREGA(ByVal vData As String)
    strCEPENTREGA = vData
End Property

Public Property Get ENDCOBRA() As String
    ENDCOBRA = strENDCOBRA
End Property

Public Property Let ENDCOBRA(ByVal vData As String)
    strENDCOBRA = vData
End Property

Public Property Get BAICOBRA() As String
    BAICOBRA = strBAICOBRA
End Property

Public Property Let BAICOBRA(ByVal vData As String)
    strBAICOBRA = vData
End Property

Public Property Get ESTCOBRA() As Integer
    ESTCOBRA = intESTCOBRA
End Property

Public Property Let ESTCOBRA(ByVal vData As Integer)
    intESTCOBRA = vData
End Property

Public Property Get CIDCOBRA() As String
    CIDCOBRA = strCIDCOBRA
End Property

Public Property Let CIDCOBRA(ByVal vData As String)
    strCIDCOBRA = vData
End Property

Public Property Get CEPCOBRA() As String
    CEPCOBRA = strCEPCOBRA
End Property

Public Property Let CEPCOBRA(ByVal vData As String)
    strCEPCOBRA = vData
End Property


Public Property Get PRODUTOS() As Variant
    PRODUTOS = arrPRODUTOS
End Property

Public Property Let PRODUTOS(ByVal vData As Variant)
    arrPRODUTOS = vData
End Property

Public Property Get DESPESAS() As Variant
    DESPESAS = arrDESPESAS
End Property

Public Property Let DESPESAS(ByVal vData As Variant)
    arrDESPESAS = vData
End Property

Public Property Get SERVICOS() As Variant
    SERVICOS = arrSERVICOS
End Property

Public Property Let SERVICOS(ByVal vData As Variant)
    arrSERVICOS = vData
End Property

Public Property Let VALPROPOSTA(ByVal vData As Integer)
    intVALPROPOSTA = vData
End Property

Public Property Get VALPROPOSTA() As Integer
    VALPROPOSTA = intVALPROPOSTA
End Property

Public Property Get PRZENTREGA() As Integer
    PRZENTREGA = intPRZENTREGA
End Property

Public Property Let PRZENTREGA(ByVal vData As Integer)
    intPRZENTREGA = vData
End Property

Public Property Get IPIISENTO() As Integer
    IPIISENTO = intIPIISENTO
End Property

Public Property Let IPIISENTO(ByVal vData As Integer)
    intIPIISENTO = vData
End Property

Public Property Get ICMSISENTO() As Integer
    ICMSISENTO = intICMSISENTO
End Property

Public Property Let ICMSISENTO(ByVal vData As Integer)
    intICMSISENTO = vData
End Property

Public Property Get FRETE() As Integer
    FRETE = intFRETE
End Property

Public Property Let FRETE(ByVal vData As Integer)
    intFRETE = vData
End Property

Public Property Get VALBASICMS() As Currency
    VALBASICMS = curVALBASICMS
End Property

Public Property Let VALBASICMS(ByVal vData As Currency)
    curVALBASICMS = vData
End Property

Public Property Get ALIQICMS() As Currency
    ALIQICMS = curAliqICMS
End Property

Public Property Let ALIQICMS(ByVal vData As Currency)
    curAliqICMS = vData
End Property

Public Property Get VLICMS() As Currency
    VLICMS = curVLICMS
End Property

Public Property Let VLICMS(ByVal vData As Currency)
    curVLICMS = vData
End Property

Public Property Get OUTRDESPESAS() As Currency
    OUTRDESPESAS = curOUTRDESPESAS
End Property

Public Property Let OUTRDESPESAS(ByVal vData As Currency)
    curOUTRDESPESAS = vData
End Property

Public Property Get VLFRETE() As Currency
    VLFRETE = curVLFRETE
End Property

Public Property Let VLFRETE(ByVal vData As Currency)
    curVLFRETE = vData
End Property

Public Property Get VLIPI() As Currency
    VLIPI = curVLIPI
End Property

Public Property Let VLIPI(ByVal vData As Currency)
    curVLIPI = vData
End Property

Public Property Get VLDESCTO() As Currency
    VLDESCTO = curVLDESCTO
End Property

Public Property Let VLDESCTO(ByVal vData As Currency)
    curVLDESCTO = vData
End Property

Public Property Get TOTORCTO() As Currency
    TOTORCTO = curTOTORCTO
End Property

Public Property Let TOTORCTO(ByVal vData As Currency)
    curTOTORCTO = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strACESSO As String, strUSUARIO As String, lngIDUsuario As Long, Optional intFILIAL_PED As Integer, Optional strVERSAO As String, Optional strNOMCOMP As String)
    frmCADPEDVENDAP.cCaminho = cCaminho
    frmCADPEDVENDAP.Linha = Linha
    frmCADPEDVENDAP.FILIAL = iFilial
    frmCADPEDVENDAP.strACESSO = strACESSO
    frmCADPEDVENDAP.strUSUARIO = strUSUARIO
    frmCADPEDVENDAP.lngCodUsuaro = lngIDUsuario
    frmCADPEDVENDAP.intFILIALPED = intFILIAL_PED
    frmCADPEDVENDAP.strVERSAO = strVERSAO
    frmCADPEDVENDAP.strNOMCOMP = strNOMCOMP
    frmCADPEDVENDAP.Show vbModal
End Sub

Public Sub cConnectPesq(cCaminho As String, Linha As Variant, iFilial As Integer, strACESSO As String, strUSUARIO As String, lngIDUsuario As Long, Optional intFILIAL_PED As Integer, Optional boolSomenteConsulta As Boolean)
    frmCADPEDVENDA.cCaminho = cCaminho
    frmCADPEDVENDA.Linha = Linha
    frmCADPEDVENDA.iCodigo = lngCODPEDIDO
    frmCADPEDVENDA.cTipOper = "C"
    frmCADPEDVENDA.FILIAL = iFilial
    frmCADPEDVENDA.strACESSO = strACESSO
    frmCADPEDVENDA.strMODPAI = "frmCADCLIENTE"
    frmCADPEDVENDA.strUSUARIO = strUSUARIO
    frmCADPEDVENDA.lngCodVendedor = 0
    frmCADPEDVENDA.lngCodUsuario = lngIDUsuario
    frmCADPEDVENDA.intFILIALPED = intFILIAL_PED
    frmCADPEDVENDA.Show vbModal
End Sub

Public Function Pesq_CadPedido() As Boolean

   Pesq_CadPedido = False
   
   If BREC.State = 1 Then BREC.Close
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       * " & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_CADPEDVENDH " & vbCrLf
   sSql = sSql + " Where " & vbCrLf
   sSql = sSql + "       SGI_FILIAL = " & intFILIAL & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_CadPedido = True
   BREC.Close

End Function


Public Sub PreencheComboCliente(Combo As Variant)

     Combo.Clear
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       * " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADCLIENTE  " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL  = " & intFILIAL & vbCrLf
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_RAZAOSOC)
        Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
        BREC.MoveNext
     Loop
    
     BREC.Close
    
End Sub

Public Sub PreencheComboCondPgto(Combo As Variant)

     Combo.Clear
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       PGTO.* " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADCONDPGTO  PGTO " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       PGTO.SGI_FILIAL  = " & intFILIAL & vbCrLf
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_DESCRICAO)
        Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
        BREC.MoveNext
     Loop
    
     BREC.Close
    
End Sub

Public Sub PreencheComboVendedor(Combo As Variant)

     Combo.Clear
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       * " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADVENDEDOR " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL  = " & intFILIAL & vbCrLf
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_DESCRICAO)
        Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
        BREC.MoveNext
     Loop
    
     BREC.Close
    
End Sub


Public Sub PreencheComboTipoOrc(Combo As Variant)

     Combo.Clear
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       * " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADESPORCA " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL  = " & intFILIAL & vbCrLf
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_DESCRICAO)
        Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
        BREC.MoveNext
     Loop
    
     BREC.Close
    
End Sub

Public Sub PreencheComboProd(Combo As Variant, intCODTIPORCA As Integer)

     Combo.Clear
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       PRODUTO.SGI_CODIGO " & vbCrLf
     sSql = sSql & "      ,PRODUTO.SGI_DESCRICAO " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADPRODUTO PRODUTO " & vbCrLf
     sSql = sSql & "      ,SGI_CADPRODTIPORCA TIPORCA " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       TIPORCA.SGI_FILIAL        = " & intFILIAL & vbCrLf
     sSql = sSql & "   And TIPORCA.SGI_CODTIPORCA    = " & intCODTIPORCA & vbCrLf
     sSql = sSql & "   And TIPORCA.SGI_FILIAL        = PRODUTO.SGI_FILIAL " & vbCrLf
     sSql = sSql & "   And TIPORCA.SGI_CODPROD       = PRODUTO.SGI_CODIGO  " & vbCrLf
     
     BREC5.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC5.EOF
        Combo.AddItem Trim(BREC5!SGI_CODIGO) & Space(10 - Len(Trim(BREC5!SGI_CODIGO))) & " - " & Trim(BREC5!SGI_DESCRICAO)
        BREC5.MoveNext
     Loop
    
     BREC5.Close
    
End Sub

Public Sub PreencheComboTipoDespesa(Combo As Variant, lngGrpGastos As Long)

     Combo.Clear
     
    sSql = "Select " & vbCrLf
    sSql = sSql & "       GRPDE.SGI_CODIGO    " & vbCrLf
    sSql = sSql & "      ,GRPDE.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "      ,GRPDE.SGI_VALOR     " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADGRPITENS ITENS " & vbCrLf
    sSql = sSql & "      ,SGI_CADTIPDESP  GRPDE " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       ITENS.SGI_FILIAL = " & intFILIAL & vbCrLf
    sSql = sSql & "   And ITENS.SGI_CODIGO = " & lngGrpGastos & vbCrLf
    sSql = sSql & "   And GRPDE.SGI_FILIAL = ITENS.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And GRPDE.SGI_CODIGO = ITENS.SGI_CODGASTO "
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_DESCRICAO)
        Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
        BREC.MoveNext
     Loop
    
     BREC.Close
    
End Sub

Public Sub PreencheComboTipoServicos(Combo As Variant, strCODPROD As String)

     Combo.Clear
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       SGI_CADSERVICOS.* " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_SERVPROD    SGI_SERVPROD " & vbCrLf
     sSql = sSql & "      ,SGI_CADSERVICOS SGI_CADSERVICOS " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_SERVPROD.SGI_FILIAL     = " & intFILIAL & vbCrLf
     sSql = sSql & "   And SGI_SERVPROD.SGI_PRODUTO    = '" & Trim(strCODPROD) & "'" & vbCrLf
     sSql = sSql & "   And SGI_CADSERVICOS.SGI_FILIAL  = SGI_SERVPROD.SGI_FILIAL " & vbCrLf
     sSql = sSql & "   And SGI_CADSERVICOS.SGI_CODSERV = SGI_SERVPROD.SGI_CODIGO " & vbCrLf
     
     BREC5.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC5.EOF
        Combo.AddItem Trim(BREC5!SGI_DESCSERV)
        Combo.ItemData(Combo.NewIndex) = BREC5!SGI_CODSERV
        BREC5.MoveNext
     Loop
    
     BREC5.Close
    
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

Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_Trans
     
     GRAVA = False
         
     Dim i           As Integer
     Dim j           As Integer
     Dim sValor      As String
     Dim strTIPODESP As String
     Dim lngQtdPed   As Long
     
     '' Inicia transação
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
      
     If strOPER = "I" Then '' Inclui
        
        '' Grava Cotação HEADER(CABEÇALHO)
        sSql = "Insert Into SGI_CADPEDVENDH (SGI_FILIAL" & vbCrLf
        sSql = sSql & "                      ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                      ,SGI_DATAPED" & vbCrLf
        sSql = sSql & "                      ,SGI_DTCALIBRA" & vbCrLf
        sSql = sSql & "                      ,SGI_CODCLI" & vbCrLf
        sSql = sSql & "                      ,SGI_CODVEND" & vbCrLf
        sSql = sSql & "                      ,SGI_CODCONDPGT" & vbCrLf
        sSql = sSql & "                      ,SGI_CODTIPORC" & vbCrLf
        sSql = sSql & "                      ,SGI_PRZENTR" & vbCrLf
        sSql = sSql & "                      ,SGI_VALPROP" & vbCrLf
        sSql = sSql & "                      ,SGI_IPIISENTO" & vbCrLf
        sSql = sSql & "                      ,SGI_ICMSISENTO" & vbCrLf
        sSql = sSql & "                      ,SGI_TIPFRETE" & vbCrLf
        sSql = sSql & "                      ,SGI_QTDITENS" & vbCrLf
        sSql = sSql & "                      ,SGI_VLTOT" & vbCrLf
        sSql = sSql & "                      ,SGI_OUTRASTOT" & vbCrLf
        sSql = sSql & "                      ,SGI_FRETETOT" & vbCrLf
        sSql = sSql & "                      ,SGI_BASEICMS" & vbCrLf
        sSql = sSql & "                      ,SGI_PORCICMS" & vbCrLf
        sSql = sSql & "                      ,SGI_VLICMS" & vbCrLf
        sSql = sSql & "                      ,SGI_VLIPI" & vbCrLf
        sSql = sSql & "                      ,SGI_VLDESCTO" & vbCrLf
        sSql = sSql & "                      ,SGI_STATUS" & vbCrLf
        sSql = sSql & "                      ,SGI_ENDENTRE" & vbCrLf
        sSql = sSql & "                      ,SGI_BAIENTRE" & vbCrLf
        sSql = sSql & "                      ,SGI_CIDENTRE" & vbCrLf
        sSql = sSql & "                      ,SGI_ESTENTRE" & vbCrLf
        sSql = sSql & "                      ,SGI_CEPENTRE" & vbCrLf
        sSql = sSql & "                      ,SGI_ENDCOBR" & vbCrLf
        sSql = sSql & "                      ,SGI_BAICOBR" & vbCrLf
        sSql = sSql & "                      ,SGI_CIDCOBR" & vbCrLf
        sSql = sSql & "                      ,SGI_ESTCOBR" & vbCrLf
        sSql = sSql & "                      ,SGI_CEPCOBR" & vbCrLf
        sSql = sSql & "                      ,SGI_CODCOTA" & vbCrLf
        sSql = sSql & "                      ,SGI_DTENTREGA" & vbCrLf
        sSql = sSql & "                      ,SGI_TELENTR" & vbCrLf
        sSql = sSql & "                      ,SGI_FAXENTR" & vbCrLf
        sSql = sSql & "                      ,SGI_TELCOBR" & vbCrLf
        sSql = sSql & "                      ,SGI_FAXCOBR" & vbCrLf
        sSql = sSql & "                      ,SGI_CONTATO" & vbCrLf
        sSql = sSql & "                      ,SGI_DEPTO" & vbCrLf
        sSql = sSql & "                      ,SGI_ORDCOCLI" & vbCrLf
        sSql = sSql & "                      ,SGI_CODTRANSP" & vbCrLf
        sSql = sSql & "                      ,SGI_EMAIL" & vbCrLf
        sSql = sSql & "                      ,SGI_CODTEC" & vbCrLf
        sSql = sSql & "                      ,SGI_OBS" & vbCrLf
        sSql = sSql & "                      ,SGI_CODTIPODOC" & vbCrLf
        sSql = sSql & "                      ,SGI_CODBANCO" & vbCrLf
        sSql = sSql & "                      ,SGI_PORCDESCPED" & vbCrLf
        sSql = sSql & "                      ,SGI_VALDESCPED" & vbCrLf
        sSql = sSql & "                      ,SGI_QTDEITENSPEDIDO" & vbCrLf
        sSql = sSql & "                      ,SGI_OBSQUALID" & vbCrLf
        sSql = sSql & "                      ,SGI_FILIALPED" & vbCrLf
        sSql = sSql & "                      ,SGI_ESPECIAL" & vbCrLf
        sSql = sSql & "                      ,SGI_PARAESTOQUE" & vbCrLf
        sSql = sSql & "                      ,SGI_CONFERIDO" & vbCrLf
        sSql = sSql & "                      ,SGI_VERSAO" & vbCrLf
        sSql = sSql & "                      ,SGI_NOMCOMP" & vbCrLf
        
        sSql = sSql & "                      )" & vbCrLf
        
        sSql = sSql & "               Values (" & vbCrLf
        sSql = sSql & "                       " & intFILIAL & vbCrLf
        sSql = sSql & "                      ," & lngCODPEDIDO & vbCrLf
        sSql = sSql & "                      ,'" & Format(dtDATAPED, "MM/DD/YYYY") & "'" & vbCrLf
        If IsDate(dtDATCALIBR) Then
           sSql = sSql & "                      ,'" & Format(dtDATCALIBR, "MM/DD/YYYY") & "'" & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
        sSql = sSql & "                      ," & intCODCLIE & vbCrLf
        sSql = sSql & "                      ," & intCODVEND & vbCrLf
        sSql = sSql & "                      ," & intCODCONDPGTO & vbCrLf
        sSql = sSql & "                      ," & intTIPPED & vbCrLf
        sSql = sSql & "                      ," & intPRZENTREGA & vbCrLf
        sSql = sSql & "                      ," & intVALPROPOSTA & vbCrLf
        sSql = sSql & "                      ," & intIPIISENTO & vbCrLf
        sSql = sSql & "                      ," & intICMSISENTO & vbCrLf
        sSql = sSql & "                      ," & intFRETE & vbCrLf
        sSql = sSql & "                      ," & intTOTALITENS & vbCrLf
        
        sValor = Replace(curTOTORCTO, ".", "")
        sValor = Replace(sValor, ",", ".")
        sSql = sSql & "                      ," & sValor & vbCrLf
        
        If curOUTRDESPESAS > 0 Then
           sValor = Replace(curOUTRDESPESAS, ".", "")
           sValor = Replace(sValor, ",", ".")
           sSql = sSql & "                      ," & sValor & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
        
        If curVLFRETE > 0 Then
           sValor = Replace(curVLFRETE, ".", "")
           sValor = Replace(sValor, ",", ".")
           sSql = sSql & "                      ," & sValor & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
        
        If curVALBASICMS > 0 Then
           sValor = Replace(curVALBASICMS, ".", "")
           sValor = Replace(sValor, ",", ".")
           sSql = sSql & "                      ," & sValor & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
        
        If curAliqICMS > 0 Then
           sValor = Replace(curAliqICMS, ".", "")
           sValor = Replace(sValor, ",", ".")
           sSql = sSql & "                      ," & sValor & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
        
        If curVLICMS > 0 Then
           sValor = Replace(curVLICMS, ".", "")
           sValor = Replace(sValor, ",", ".")
           sSql = sSql & "                      ," & sValor & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If

        If curVLIPI > 0 Then
           sValor = Replace(curVLIPI, ".", "")
           sValor = Replace(sValor, ",", ".")
           sSql = sSql & "                      ," & sValor & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If

        If curVLDESCTO > 0 Then
           sValor = Replace(curVLDESCTO, ".", "")
           sValor = Replace(sValor, ",", ".")
           sSql = sSql & "                      ," & sValor & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
        
        sSql = sSql & "                      ,'" & strSTATUS & "'" & vbCrLf
        sSql = sSql & "                      ,'" & strENDENTR & "'" & vbCrLf
        sSql = sSql & "                      ,'" & strBAIENTR & "'" & vbCrLf
        sSql = sSql & "                      ,'" & strCIDENTR & "'" & vbCrLf
        sSql = sSql & "                      ," & intESTENTREGA & vbCrLf
        sSql = sSql & "                      ,'" & strCEPENTREGA & "'" & vbCrLf
        sSql = sSql & "                      ,'" & strENDCOBRA & "'" & vbCrLf
        sSql = sSql & "                      ,'" & strBAICOBRA & "'" & vbCrLf
        sSql = sSql & "                      ,'" & strCIDCOBRA & "'" & vbCrLf
        sSql = sSql & "                      ," & intESTCOBRA & vbCrLf
        sSql = sSql & "                      ,'" & strCEPCOBRA & "'" & vbCrLf
        
        If lngCODCOTA > 0 Then
           sSql = sSql & "                      ," & lngCODCOTA & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
        
        sSql = sSql & "                      ,'" & Format(dtDATAENTREGA, "MM/DD/YYYY") & "'" & vbCrLf
        
        sSql = sSql & "                      ,'" & strTELENTR & "'" & vbCrLf
        sSql = sSql & "                      ,'" & strFAXENTR & "'" & vbCrLf
        sSql = sSql & "                      ,'" & strTELCOBRA & "'" & vbCrLf
        sSql = sSql & "                      ,'" & strFAXCOBRA & "'" & vbCrLf
        sSql = sSql & "                      ,'" & strCONTATO & "'" & vbCrLf
        sSql = sSql & "                      ,'" & strDEPARTAMENTO & "'" & vbCrLf
        sSql = sSql & "                      ,'" & strORDCOMPCLI & "'" & vbCrLf
        sSql = sSql & "                      ," & lngCODTRANSP & vbCrLf
        sSql = sSql & "                      ,'" & strEMAIL & "'" & vbCrLf
        If intCODTECNICO > 0 Then
           sSql = sSql & "                      ," & intCODTECNICO & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
        sSql = sSql & "                      ,'" & strOBSERVACAO & "'" & vbCrLf
        
        If lngCODTIPODOC > 0 Then
           sSql = sSql & "                      ," & lngCODTIPODOC & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
        If lngCODBANCO > 0 Then
           sSql = sSql & "                      ," & lngCODBANCO & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
        
        If curPORDESC > 0 Then
           sValor = Replace(curPORDESC, ".", "")
           sValor = Replace(sValor, ",", ".")
           sSql = sSql & "                      ," & sValor & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
        
        If curVALDESC > 0 Then
           sValor = Replace(curVALDESC, ".", "")
           sValor = Replace(sValor, ",", ".")
           sSql = sSql & "                      ," & sValor & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
        
        sValor = "Null"
        If curQTDITENSPROD > 0 Then
           sValor = Replace(curQTDITENSPROD, ".", "")
           sValor = Replace(sValor, ",", ".")
        End If
        sSql = sSql & "                      ," & curQTDITENSPROD & vbCrLf
        sSql = sSql & "                      ," & Trim(strOBS2) & vbCrLf
        sSql = sSql & "                      ," & intFILIALPED & vbCrLf
        sSql = sSql & "                      ," & intESPECIAL & vbCrLf
        sSql = sSql & "                      ," & intPARAESTOQUE & vbCrLf
        sSql = sSql & "                      ," & intCONFERIDO & vbCrLf
        sSql = sSql & "                      ," & strVERSAO & vbCrLf
        sSql = sSql & "                      ," & strNOMCOMP & vbCrLf
        
        sSql = sSql & "                           )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        
        '' Grava os Produtos
        lngQtdPed = 0
        If IsArray(arrPRODUTOS) = True Then
           For i = 1 To UBound(arrPRODUTOS)
               
                     sSql = "Insert Into SGI_CADPEDVENDI (SGI_FILIAL" & vbCrLf
                     sSql = sSql & "                      ,SGI_CODIGO" & vbCrLf
                     sSql = sSql & "                      ,SGI_IDPRODUTO" & vbCrLf
                     sSql = sSql & "                      ,SGI_CODPROD" & vbCrLf
                     sSql = sSql & "                      ,SGI_QTDE" & vbCrLf
                     sSql = sSql & "                      ,SGI_VLUNIT" & vbCrLf
                     sSql = sSql & "                      ,SGI_PRCDESC" & vbCrLf
                     sSql = sSql & "                      ,SGI_PRCIPI" & vbCrLf
                     sSql = sSql & "                      ,SGI_VLTOT" & vbCrLf
                     sSql = sSql & "                      ,SGI_VLIPI" & vbCrLf
                     sSql = sSql & "                      ,SGI_VLDESC" & vbCrLf
                     sSql = sSql & "                      ,SGI_ALTFILM" & vbCrLf
                     sSql = sSql & "                      ,SGI_FOTNOVO" & vbCrLf
                     sSql = sSql & "                      ,SGI_REPETICAO" & vbCrLf
                     sSql = sSql & "                      ,SGI_DATAPED" & vbCrLf
                     sSql = sSql & "                      ,SGI_CODLINPROD" & vbCrLf
                     sSql = sSql & "                      ,SGI_CODFORN" & vbCrLf
                     sSql = sSql & "                      ,SGI_OBSOP" & vbCrLf
                     sSql = sSql & "                      ,SGI_FECHTPFU" & vbCrLf
                     sSql = sSql & "                      ,SGI_FILIALPED" & vbCrLf
                     sSql = sSql & "                      ,SGI_QTDEPORPALHETS" & vbCrLf
                     sSql = sSql & "                      ,SGI_PALHETS" & vbCrLf
                     sSql = sSql & "                      ,SGI_CONFERIDO" & vbCrLf
                     sSql = sSql & "                      ,SGI_USARPADRPALSN" & vbCrLf
                     
                     sSql = sSql & "              ) Values (" & vbCrLf
                     
                     sSql = sSql & "                       " & intFILIAL & vbCrLf
                     sSql = sSql & "                      ," & lngCODPEDIDO & vbCrLf
                     sSql = sSql & "                      ," & Trim(arrPRODUTOS(i, 9)) & vbCrLf
                     sSql = sSql & "                      ,'" & Trim(arrPRODUTOS(i, 1)) & "'" & vbCrLf
                     sSql = sSql & "                      ," & arrPRODUTOS(i, 2) & vbCrLf
                     lngQtdPed = lngQtdPed + CLng(arrPRODUTOS(i, 2))
                     sSql = sSql & "                      ," & arrPRODUTOS(i, 3) & vbCrLf
                     sSql = sSql & "                      ," & arrPRODUTOS(i, 4) & vbCrLf
                     sSql = sSql & "                      ," & arrPRODUTOS(i, 5) & vbCrLf
                     sSql = sSql & "                      ," & arrPRODUTOS(i, 6) & vbCrLf
                     sSql = sSql & "                      ," & arrPRODUTOS(i, 7) & vbCrLf
                     sSql = sSql & "                      ," & arrPRODUTOS(i, 8) & vbCrLf
                     
                     sSql = sSql & "                      ," & Trim(arrPRODUTOS(i, 10)) & vbCrLf
                     sSql = sSql & "                      ," & Trim(arrPRODUTOS(i, 11)) & vbCrLf
                     sSql = sSql & "                      ," & Trim(arrPRODUTOS(i, 12)) & vbCrLf
                     sSql = sSql & "                      ,'" & Format(dtDATAPED, "MM/DD/YYYY") & "'" & vbCrLf
                     sSql = sSql & "                      ," & arrPRODUTOS(i, 13) & vbCrLf
                     sSql = sSql & "                      ," & arrPRODUTOS(i, 14) & vbCrLf
                     sSql = sSql & "                      ,'" & Trim(arrPRODUTOS(i, 15)) & "'" & vbCrLf
                     sSql = sSql & "                      ," & Trim(arrPRODUTOS(i, 16)) & vbCrLf
                     sSql = sSql & "                      ," & intFILIALPED & vbCrLf
                     sSql = sSql & "                      ," & arrPRODUTOS(i, 19) & vbCrLf
                     sSql = sSql & "                      ," & arrPRODUTOS(i, 20) & vbCrLf
                     sSql = sSql & "                      ," & arrPRODUTOS(i, 21) & vbCrLf
                     sSql = sSql & "                      ," & arrPRODUTOS(i, 22) & vbCrLf
                     
                     sSql = sSql & "                           )"
                
                     BGRV.CommandText = sSql
                     BGRV.Execute
                     
                     If (arrPRODUTOS(i, 10) = 1 Or arrPRODUTOS(i, 11) = 1) Then
                           
                        sSql = "Update SGI_CADPRODUTO" & vbCrLf
                        sSql = sSql & "                      Set SGI_STATUS = 2" & vbCrLf
                        sSql = sSql & "       Where" & vbCrLf
                        sSql = sSql & "             SGI_FILIAL    = " & intFILIAL & vbCrLf
                        sSql = sSql & "         And SGI_IDPRODUTO = " & Trim(arrPRODUTOS(i, 9))
                        
                        BGRV.CommandText = sSql
                        BGRV.Execute
                        
                     End If
                     
           
           Next i
        End If
        
        
        '' =====================================
        '' Programação de Entregas
        If IsArray(arrPROGENTREGAS) Then
           For i = 1 To UBound(arrPROGENTREGAS)
                   
                   sSql = "Insert Into SGI_PROGENTRPROD (" & vbCrLf
                   sSql = sSql & "                       SGI_FILIAL" & vbCrLf
                   sSql = sSql & "                      ,SGI_CODPED" & vbCrLf
                   sSql = sSql & "                      ,SGI_IDPRODUTO" & vbCrLf
                   sSql = sSql & "                      ,SGI_INDICE" & vbCrLf
                   sSql = sSql & "                      ,SGI_QTDE" & vbCrLf
                   sSql = sSql & "                      ,SGI_DATENTREGA" & vbCrLf
                   sSql = sSql & "                      ,SGI_OBSOP" & vbCrLf
                   sSql = sSql & "                      ,SGI_FECHTPFU" & vbCrLf
                   sSql = sSql & "                      ,SGI_FILIALPED" & vbCrLf
                   sSql = sSql & "                      ,SGI_IDINTERNO" & vbCrLf
                   sSql = sSql & "                      ,SGI_CODOP" & vbCrLf
                   sSql = sSql & "                      ,SGI_STATUS" & vbCrLf
                   sSql = sSql & "                      ,SGI_QTDEPORPALHETS" & vbCrLf
                   sSql = sSql & "                      ,SGI_PALHETS" & vbCrLf
                   
                   sSql = sSql & "                      ) Values ( " & vbCrLf
                   sSql = sSql & "                               " & intFILIAL & vbCrLf
                   sSql = sSql & "                              ," & lngCODPEDIDO & vbCrLf
                   sSql = sSql & "                              ," & arrPROGENTREGAS(i, 0) & vbCrLf
                   sSql = sSql & "                              ," & i & vbCrLf
                   
                   sValor = Replace(arrPROGENTREGAS(i, 1), ".", "")
                   sValor = Replace(sValor, ",", ".")
                   sSql = sSql & "                              ," & sValor & vbCrLf
                   
                   sSql = sSql & "                              ," & arrPROGENTREGAS(i, 2) & vbCrLf
                   sSql = sSql & "                              ,'" & Trim(arrPROGENTREGAS(i, 4)) & "'" & vbCrLf
                   sSql = sSql & "                              ," & Trim(arrPROGENTREGAS(i, 5)) & vbCrLf
                   sSql = sSql & "                              ," & intFILIALPED & vbCrLf
                   sSql = sSql & "                              ," & arrPROGENTREGAS(i, 16) & vbCrLf
                   sSql = sSql & "                              ,Null" & vbCrLf
                   sSql = sSql & "                              ," & arrPROGENTREGAS(i, 17) & vbCrLf
                   sSql = sSql & "                              ," & arrPROGENTREGAS(i, 18) & vbCrLf
                   sSql = sSql & "                              ," & arrPROGENTREGAS(i, 19) & vbCrLf
                   
                   sSql = sSql & "                              )"
                   
                   BGRV.CommandText = sSql
                   BGRV.Execute
                   
           Next i
        End If
        '' =====================================
        
        '' =====================================
        '' Grava Log
        sSql = ""
        
        sSql = "Insert Into SGI_LOGPED (" & vbCrLf
        sSql = sSql & "                              SGI_FILIAL" & vbCrLf
        sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                             ,SGI_DATA" & vbCrLf
        sSql = sSql & "                             ,SGI_HORA" & vbCrLf
        sSql = sSql & "                             ,SGI_CODUSUARIO" & vbCrLf
        sSql = sSql & "                             ,SGI_ACAO" & vbCrLf
        sSql = sSql & "                    ) Values (" & vbCrLf
        sSql = sSql & "                              " & FILIAL & vbCrLf
        sSql = sSql & "                             ," & lngCODPEDIDO & vbCrLf
        sSql = sSql & "                             ,'" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                             ,'" & Format(Now, "HH:MM:SS") & "'" & vbCrLf
        sSql = sSql & "                             ," & lngCodUsuario & vbCrLf
        sSql = sSql & "                             ,'" & strOPER & "'" & vbCrLf
        sSql = sSql & "                             )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' =====================================
     
     ElseIf strOPER = "A" Then '' Altera
     
         '' Apagando Tipo de Reprovação
         sSql = "Delete from SGI_TIPREPRPED " & vbCrLf
         sSql = sSql & "      Where " & vbCrLf
         sSql = sSql & "           SGI_FILIAL = " & intFILIAL & vbCrLf
         sSql = sSql & "       And SGI_CODPED = " & lngCODPEDIDO
         
         BGRV.CommandText = sSql
         BGRV.Execute
         
         '' Pedidos Steel
         sSql = "Update SGI_CADPEDVENDH Set" & vbCrLf
         
         sSql = sSql & "                           SGI_DATAPED    = '" & Format(dtDATAPED, "MM/DD/YYYY") & "'" & vbCrLf
         sSql = sSql & "                          ,SGI_CODVEND    = " & intCODVEND & vbCrLf
         sSql = sSql & "                          ,SGI_CODCONDPGT = " & intCODCONDPGTO & vbCrLf
         sSql = sSql & "                          ,SGI_CODTIPORC  = " & intTIPPED & vbCrLf
         sSql = sSql & "                          ,SGI_CODTRANSP  = " & lngCODTRANSP & vbCrLf
         sSql = sSql & "                          ,SGI_CONFERIDO  = " & intCONFERIDO & vbCrLf
         
         sValor = "Null"
         If curVALBASICMS > 0 Then
            sValor = Replace(curVALBASICMS, ".", "")
            sValor = Replace(sValor, ",", ".")
         End If
         sSql = sSql & "                      ,SGI_BASEICMS = " & sValor & vbCrLf
         
         sValor = "Null"
         If curAliqICMS > 0 Then
            sValor = Replace(curAliqICMS, ".", "")
            sValor = Replace(sValor, ",", ".")
         End If
         sSql = sSql & "                      ,SGI_PORCICMS = " & sValor & vbCrLf
         
         sValor = "Null"
         If curVLICMS > 0 Then
            sValor = Replace(curVLICMS, ".", "")
            sValor = Replace(sValor, ",", ".")
         End If
         sSql = sSql & "                      ,SGI_VLICMS = " & sValor & vbCrLf
         
         sValor = "Null"
         If curVLIPI > 0 Then
            sValor = Replace(curVLIPI, ".", "")
            sValor = Replace(sValor, ",", ".")
         End If
         sSql = sSql & "                      ,SGI_VLIPI = " & sValor & vbCrLf
         
         sValor = "Null"
         If curOUTRDESPESAS > 0 Then
            sValor = Replace(curOUTRDESPESAS, ".", "")
            sValor = Replace(sValor, ",", ".")
         End If
         sSql = sSql & "                      ,SGI_OUTRASTOT = " & sValor & vbCrLf
         
         sValor = "Null"
         If curVLFRETE > 0 Then
            sValor = Replace(curVLFRETE, ".", "")
            sValor = Replace(sValor, ",", ".")
         End If
         sSql = sSql & "                      ,SGI_FRETETOT = " & sValor & vbCrLf
         
         sValor = "Null"
         If curVLDESCTO > 0 Then
            sValor = Replace(curVLDESCTO, ".", "")
            sValor = Replace(sValor, ",", ".")
         End If
         sSql = sSql & "                      ,SGI_VLDESCTO = " & sValor & vbCrLf
         
         sValor = "Null"
         If curTOTORCTO > 0 Then
            sValor = Replace(curTOTORCTO, ".", "")
            sValor = Replace(sValor, ",", ".")
         End If
         sSql = sSql & "                      ,SGI_VLTOT = " & sValor & vbCrLf
         
         sSql = sSql & "                      ,SGI_QTDITENS     = " & intTOTALITENS & vbCrLf
         sSql = sSql & "                      ,SGI_ESPECIAL     = " & intESPECIAL & vbCrLf
         sSql = sSql & "                      ,SGI_PARAESTOQUE  = " & intPARAESTOQUE & vbCrLf
         
         sSql = sSql & "                      ,SGI_ENDENTRE     = '" & strENDENTR & "'" & vbCrLf
         sSql = sSql & "                      ,SGI_BAIENTRE     = '" & strBAIENTR & "'" & vbCrLf
         sSql = sSql & "                      ,SGI_CIDENTRE     = '" & strCIDENTR & "'" & vbCrLf
         sSql = sSql & "                      ,SGI_CEPENTRE     = '" & strCEPENTREGA & "'" & vbCrLf
         sSql = sSql & "                      ,SGI_ESTENTRE     = " & intESTENTREGA & vbCrLf
         sSql = sSql & "                      ,SGI_TELENTR      = '" & strTELENTR & "'" & vbCrLf
         sSql = sSql & "                      ,SGI_FAXENTR      = '" & strFAXENTR & "'" & vbCrLf
    
         sSql = sSql & "                      ,SGI_ENDCOBR      = '" & strENDCOBRA & "'" & vbCrLf
         sSql = sSql & "                      ,SGI_BAICOBR      = '" & strBAICOBRA & "'" & vbCrLf
         sSql = sSql & "                      ,SGI_CIDCOBR      = '" & strCIDCOBRA & "'" & vbCrLf
         sSql = sSql & "                      ,SGI_CEPCOBR      = '" & strCEPCOBRA & "'" & vbCrLf
         sSql = sSql & "                      ,SGI_ESTCOBR      = " & intESTCOBRA & vbCrLf
         sSql = sSql & "                      ,SGI_TELCOBR      = '" & strTELCOBRA & "'" & vbCrLf
         sSql = sSql & "                      ,SGI_FAXCOBR      = '" & strFAXCOBRA & "'" & vbCrLf
         
         sSql = sSql & "                      ,SGI_CONTATO      = '" & strCONTATO & "'" & vbCrLf
         sSql = sSql & "                      ,SGI_DEPTO        = '" & strDEPARTAMENTO & "'" & vbCrLf
         sSql = sSql & "                      ,SGI_ORDCOCLI     = '" & strORDCOMPCLI & "'" & vbCrLf
         sSql = sSql & "                      ,SGI_EMAIL        = '" & strEMAIL & "'" & vbCrLf
         sSql = sSql & "                      ,SGI_OBS          = '" & strOBSERVACAO & "'" & vbCrLf
         sSql = sSql & "                      ,SGI_OBSQUALID    = " & strOBS2 & vbCrLf
         sSql = sSql & "                      ,SGI_STATUS       = '" & strSTATUS & "'" & vbCrLf
         sSql = sSql & "                      ,SGI_NOMCOMP      = '" & strNOMCOMP & "'" & vbCrLf
        
         sSql = sSql & " Where " & vbCrLf
         sSql = sSql & "       SGI_CODIGO = " & lngCODPEDIDO
         
         BGRV.CommandText = sSql
         BGRV.Execute
     
         
        '' =====================================
        '' Programação de Entregas
        If IsArray(arrPROGENTREGAS) Then
           For i = 1 To UBound(arrPROGENTREGAS)
                If arrPROGENTREGAS(i, 3) = dacEnumUpdateAction_delete Then
                
                    If Len(Trim(arrPROGENTREGAS(i, 9))) > 0 Then
                    
                        sSql = ""
                        
                        sSql = "Delete SGI_ORDEMPROD " & vbCrLf
                        sSql = sSql & "                         Where " & vbCrLf
                        sSql = sSql & "                               SGI_FILIAL     = " & intFILIAL & vbCrLf
                        sSql = sSql & "                           And SGI_CODIGO     = " & Trim(arrPROGENTREGAS(i, 9)) & vbCrLf
                        sSql = sSql & "                           And SGI_IDPRODUTO  = " & arrPROGENTREGAS(i, 0) & vbCrLf
                        sSql = sSql & "                           And SGI_IDPAI      = " & arrPROGENTREGAS(i, 16)
                        
                        ''sSql = sSql & "                           And SGI_DATENTREGA = " & arrPROGENTREGAS(I, 10)
                    
                        BGRV.CommandText = sSql
                        BGRV.Execute
                    
                    End If
                    
                    
                    sSql = ""
                    
                    sSql = "Delete SGI_PROGENTRPROD " & vbCrLf
                    sSql = sSql & "                           Where " & vbCrLf
                    sSql = sSql & "                                 SGI_FILIAL    = " & intFILIAL & vbCrLf
                    sSql = sSql & "                             And SGI_CODPED    = " & lngCODPEDIDO & vbCrLf
                    sSql = sSql & "                             And SGI_IDPRODUTO = " & arrPROGENTREGAS(i, 0) & vbCrLf
                    sSql = sSql & "                             And SGI_IDINTERNO = " & arrPROGENTREGAS(i, 16) & vbCrLf
                
                    BGRV.CommandText = sSql
                    BGRV.Execute
                
                
                End If
           Next i
           
           For i = 1 To UBound(arrPROGENTREGAS)
                If arrPROGENTREGAS(i, 3) = dacEnumUpdateAction_Insert Then
                
                   sSql = "Insert Into SGI_PROGENTRPROD (" & vbCrLf
                   sSql = sSql & "                       SGI_FILIAL" & vbCrLf
                   sSql = sSql & "                      ,SGI_CODPED" & vbCrLf
                   sSql = sSql & "                      ,SGI_IDPRODUTO" & vbCrLf
                   sSql = sSql & "                      ,SGI_INDICE" & vbCrLf
                   sSql = sSql & "                      ,SGI_QTDE" & vbCrLf
                   sSql = sSql & "                      ,SGI_DATENTREGA" & vbCrLf
                   sSql = sSql & "                      ,SGI_OBSOP" & vbCrLf
                   sSql = sSql & "                      ,SGI_FECHTPFU" & vbCrLf
                   sSql = sSql & "                      ,SGI_FILIALPED" & vbCrLf
                   sSql = sSql & "                      ,SGI_IDINTERNO" & vbCrLf
                   sSql = sSql & "                      ,SGI_QTDEPORPALHETS" & vbCrLf
                   sSql = sSql & "                      ,SGI_PALHETS" & vbCrLf
                   
                   sSql = sSql & "                      ) Values ( " & vbCrLf
                   sSql = sSql & "                               " & intFILIAL & vbCrLf
                   sSql = sSql & "                              ," & lngCODPEDIDO & vbCrLf
                   sSql = sSql & "                              ," & arrPROGENTREGAS(i, 0) & vbCrLf
                   sSql = sSql & "                              ," & i & vbCrLf
                   
                   sValor = Replace(arrPROGENTREGAS(i, 1), ".", "")
                   sValor = Replace(sValor, ",", ".")
                   sSql = sSql & "                              ," & sValor & vbCrLf
                   
                   sSql = sSql & "                              ," & arrPROGENTREGAS(i, 2) & vbCrLf
                   sSql = sSql & "                              ,'" & Trim(arrPROGENTREGAS(i, 4)) & "'" & vbCrLf
                   sSql = sSql & "                              ," & Trim(arrPROGENTREGAS(i, 5)) & vbCrLf
                   sSql = sSql & "                              ," & intFILIALPED & vbCrLf & vbCrLf
                   sSql = sSql & "                              ," & arrPROGENTREGAS(i, 16) & vbCrLf
                   sSql = sSql & "                              ," & arrPROGENTREGAS(i, 18) & vbCrLf
                   sSql = sSql & "                              ," & arrPROGENTREGAS(i, 19) & vbCrLf
                   
                   sSql = sSql & "                              )"
                   
                   BGRV.CommandText = sSql
                   BGRV.Execute
                
                ElseIf arrPROGENTREGAS(i, 3) = dacEnumUpdateAction_update Then
                
                    sSql = ""
                    
                    sSql = "Update SGI_PROGENTRPROD Set " & vbCrLf
                    sSql = sSql & "                                   SGI_QTDE              = " & arrPROGENTREGAS(i, 1) & vbCrLf
                    sSql = sSql & "                                  ,SGI_INDICE            = " & arrPROGENTREGAS(i, 7) & vbCrLf
                    sSql = sSql & "                                  ,SGI_DATENTREGA        = " & arrPROGENTREGAS(i, 2) & vbCrLf
                    sSql = sSql & "                                  ,SGI_OBSOP             = '" & arrPROGENTREGAS(i, 4) & "'" & vbCrLf
                    sSql = sSql & "                                  ,SGI_FECHTPFU          = " & arrPROGENTREGAS(i, 5) & vbCrLf
                    sSql = sSql & "                                  ,SGI_STATUS            = " & arrPROGENTREGAS(i, 17) & vbCrLf
                    sSql = sSql & "                                  ,SGI_QTDEPORPALHETS    = " & arrPROGENTREGAS(i, 18) & vbCrLf
                    sSql = sSql & "                                  ,SGI_PALHETS           = " & arrPROGENTREGAS(i, 19) & vbCrLf
                    
                    sSql = sSql & "                           Where " & vbCrLf
                    sSql = sSql & "                                 SGI_FILIAL    = " & intFILIAL & vbCrLf
                    sSql = sSql & "                             And SGI_CODPED    = " & lngCODPEDIDO & vbCrLf
                    sSql = sSql & "                             And SGI_IDPRODUTO = " & arrPROGENTREGAS(i, 0) & vbCrLf
                    sSql = sSql & "                             And SGI_IDINTERNO = " & arrPROGENTREGAS(i, 16) & vbCrLf
                
                    BGRV.CommandText = sSql
                    BGRV.Execute
                
                    If Len(Trim(arrPROGENTREGAS(i, 9))) > 0 Then
                    
                        sSql = ""
                        
                        sSql = "Update SGI_ORDEMPROD Set " & vbCrLf
                        sSql = sSql & "                               SGI_QTDEPED           = " & arrPROGENTREGAS(i, 1) & vbCrLf
                        sSql = sSql & "                              ,SGI_QTDE              = " & arrPROGENTREGAS(i, 1) & vbCrLf
                        sSql = sSql & "                              ,SGI_DATENTREGA        = " & arrPROGENTREGAS(i, 2) & vbCrLf
                        sSql = sSql & "                              ,SGI_FECHTPFU          = " & arrPROGENTREGAS(i, 5) & vbCrLf
                        sSql = sSql & "                              ,SGI_OBSOP             = '" & arrPROGENTREGAS(i, 4) & "'" & vbCrLf
                        sSql = sSql & "                              ,SGI_JAIMPRESSA        = 0" & vbCrLf
                        sSql = sSql & "                              ,SGI_ALTERADO          = 1" & vbCrLf
                        sSql = sSql & "                              ,SGI_QTDEPORPALHETS    = " & arrPROGENTREGAS(i, 18) & vbCrLf
                        sSql = sSql & "                              ,SGI_PALHETS           = " & arrPROGENTREGAS(i, 19) & vbCrLf
                        
                        sSql = sSql & "                         Where " & vbCrLf
                        sSql = sSql & "                               SGI_FILIAL     = " & intFILIAL & vbCrLf
                        sSql = sSql & "                           And SGI_CODIGO     = " & Trim(arrPROGENTREGAS(i, 9)) & vbCrLf
                        sSql = sSql & "                           And SGI_IDPRODUTO  = " & arrPROGENTREGAS(i, 0) & vbCrLf
                        sSql = sSql & "                           And SGI_IDPAI      = " & arrPROGENTREGAS(i, 16) & vbCrLf
                    
                        BGRV.CommandText = sSql
                        BGRV.Execute
                    
                    End If
                End If
           Next i
        End If
        '' =====================================
         
         
        '' Itens do pedido
        If IsArray(arrPRODUTOS) = True Then
           For i = 1 To UBound(arrPRODUTOS)
               If arrPRODUTOS(i, 17) = dacEnumUpdateAction_update Then '' Altera
               
                  sSql = ""
                  
                  sSql = "Update SGI_CADPEDVENDI Set " & vbCrLf
                  sSql = sSql & "                                 SGI_QTDE           = " & arrPRODUTOS(i, 2) & vbCrLf
                  sSql = sSql & "                                ,SGI_VLUNIT         = " & arrPRODUTOS(i, 3) & vbCrLf
                  sSql = sSql & "                                ,SGI_PRCDESC        = " & arrPRODUTOS(i, 4) & vbCrLf
                  sSql = sSql & "                                ,SGI_PRCIPI         = " & arrPRODUTOS(i, 5) & vbCrLf
                  sSql = sSql & "                                ,SGI_VLTOT          = " & arrPRODUTOS(i, 6) & vbCrLf
                  sSql = sSql & "                                ,SGI_VLIPI          = " & arrPRODUTOS(i, 7) & vbCrLf
                  sSql = sSql & "                                ,SGI_VLDESC         = " & arrPRODUTOS(i, 8) & vbCrLf
                  sSql = sSql & "                                ,SGI_ALTFILM        = " & arrPRODUTOS(i, 10) & vbCrLf
                  sSql = sSql & "                                ,SGI_FOTNOVO        = " & arrPRODUTOS(i, 11) & vbCrLf
                  sSql = sSql & "                                ,SGI_REPETICAO      = " & arrPRODUTOS(i, 12) & vbCrLf
                  sSql = sSql & "                                ,SGI_CODLINPROD     = " & arrPRODUTOS(i, 13) & vbCrLf
                  sSql = sSql & "                                ,SGI_OBSOP          = '" & arrPRODUTOS(i, 15) & "'" & vbCrLf
                  sSql = sSql & "                                ,SGI_FECHTPFU       = " & arrPRODUTOS(i, 16) & vbCrLf
                  sSql = sSql & "                                ,SGI_CONFERIDO      = " & arrPRODUTOS(i, 21) & vbCrLf
                  sSql = sSql & "                                ,SGI_USARPADRPALSN  = " & arrPRODUTOS(i, 22) & vbCrLf
                  
                  sSql = sSql & "Where " & vbCrLf
                  sSql = sSql & "      SGI_FILIAL    = " & intFILIAL & vbCrLf
                  sSql = sSql & "  And SGI_CODIGO    = " & lngCODPEDIDO & vbCrLf
                  sSql = sSql & "  And SGI_IDPRODUTO = " & arrPRODUTOS(i, 9)
                  
                  BGRV.CommandText = sSql
                  BGRV.Execute
                  
                  '' Alteração no Filme
                  '' Fotolito Novo
                  If arrPRODUTOS(i, 10) = 1 Or _
                     arrPRODUTOS(i, 11) = 1 Then
                        
                        sSql = ""
                        
                        sSql = "Update SGI_CADPRODUTO Set SGI_STATUS = 2 " & vbCrLf
                        sSql = sSql & " Where " & vbCrLf
                        sSql = sSql & "       SGI_FILIAL    = " & intFILIAL & vbCrLf
                        sSql = sSql & "   And SGI_IDPRODUTO = " & arrPRODUTOS(i, 9)
                           
                        BGRV.CommandText = sSql
                        BGRV.Execute
                  
                  End If
                  
                  
               ElseIf arrPRODUTOS(i, 17) = dacEnumUpdateAction_delete Then '' Exclui

                  sSql = ""
                  
                  sSql = "Delete SGI_CADPEDVENDI " & vbCrLf
                  sSql = sSql & "Where " & vbCrLf
                  sSql = sSql & "      SGI_FILIAL    = " & intFILIAL & vbCrLf
                  sSql = sSql & "  And SGI_CODIGO    = " & lngCODPEDIDO & vbCrLf
                  sSql = sSql & "  And SGI_IDPRODUTO = " & arrPRODUTOS(i, 9)
                  
                  BGRV.CommandText = sSql
                  BGRV.Execute

               End If
           Next i
        End If
     
        '' =====================================
        '' Grava Log
        sSql = ""
        
        sSql = "Insert Into SGI_LOGPED (" & vbCrLf
        sSql = sSql & "                              SGI_FILIAL" & vbCrLf
        sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                             ,SGI_DATA" & vbCrLf
        sSql = sSql & "                             ,SGI_HORA" & vbCrLf
        sSql = sSql & "                             ,SGI_CODUSUARIO" & vbCrLf
        sSql = sSql & "                             ,SGI_ACAO" & vbCrLf
        sSql = sSql & "                    ) Values (" & vbCrLf
        sSql = sSql & "                              " & FILIAL & vbCrLf
        sSql = sSql & "                             ," & lngCODPEDIDO & vbCrLf
        sSql = sSql & "                             ,'" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                             ,'" & Format(Now, "HH:MM:SS") & "'" & vbCrLf
        sSql = sSql & "                             ," & lngCodUsuario & vbCrLf
        sSql = sSql & "                             ,'" & strOPER & "'" & vbCrLf
        sSql = sSql & "                             )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' =====================================
     
     ElseIf strOPER = "E" Then '' Exclui
     
         arrPRODCOPIA = arrPRODUTOS
     
         '' Apagando Tipo de Reprovação
         sSql = "Delete from SGI_TIPREPRPED " & vbCrLf
         sSql = sSql & "      Where " & vbCrLf
         sSql = sSql & "           SGI_FILIAL = " & intFILIAL & vbCrLf
         sSql = sSql & "       And SGI_CODPED = " & lngCODPEDIDO
         
         BGRV.CommandText = sSql
         BGRV.Execute
         
         '' Apagando itens de Programacao
         sSql = "Delete from SGI_PROGENTRPROD " & vbCrLf
         sSql = sSql & "        Where SGI_FILIAL = " & intFILIAL & vbCrLf
         sSql = sSql & "          And SGI_CODPED = " & lngCODPEDIDO
         
         BGRV.CommandText = sSql
         BGRV.Execute
         
         '' Apagando itens do pedido
         sSql = "Delete from SGI_CADPEDVENDI " & vbCrLf
         sSql = sSql & "        Where SGI_FILIAL = " & intFILIAL & vbCrLf
         sSql = sSql & "          And SGI_CODIGO = " & lngCODPEDIDO
         
         BGRV.CommandText = sSql
         BGRV.Execute
     
         '' Apagando Cabeçalho do pedido
         sSql = "Delete from SGI_CADPEDVENDH " & vbCrLf
         sSql = sSql & "        Where SGI_FILIAL = " & intFILIAL & vbCrLf
         sSql = sSql & "          And SGI_CODIGO = " & lngCODPEDIDO
         
         BGRV.CommandText = sSql
         BGRV.Execute
         
     ElseIf strOPER = "L" Or strOPER = "N" Then '' Libera Pedido Bloqueado
        
        sSql = "Update SGI_CADPEDVENDH Set " & vbCrLf
        sSql = sSql & "                           SGI_LIBERADO    = '" & strLIBERADO & "'" & vbCrLf
        sSql = sSql & "                          ,SGI_LIBDATAHORA = '" & Format(dtDTHORALIB, "MM/DD/YYYY HH:MM:SS") & "'" & vbCrLf
        sSql = sSql & "                          ,SGI_OBSLIBCOM   = '" & strOBSCOM & "'" & vbCrLf
        If strOPER = "N" Then
            sSql = sSql & "                          ,SGI_STATUS      = '" & Trim(strOPER) & "'" & vbCrLf
        Else
            sSql = sSql & "                          ,SGI_STATUS      = '" & Trim(strSTATUS) & "'" & vbCrLf
        End If
        sSql = sSql & "                          ,SGI_CONFERIDO   = " & intCONFERIDO & vbCrLf
        
        sSql = sSql & "                     Where " & vbCrLf & vbCrLf
        sSql = sSql & "                           SGI_FILIAL = " & intFILIAL & vbCrLf
        sSql = sSql & "                       And SGI_CODIGO = " & lngCODPEDIDO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        If strOPER = "L" Then
            
            '' Gravando Itens da Ordem
            If IsArray(arrPROGENTREGAS) Then
                For i = 1 To UBound(arrPROGENTREGAS)
                
                    If Len(Trim(arrPROGENTREGAS(i, 0))) > 0 Then
                    
                        sSql = "Insert Into SGI_ORDEMPROD (" & vbCrLf
                        sSql = sSql & "                       SGI_FILIAL" & vbCrLf
                        sSql = sSql & "                      ,SGI_CODIGO" & vbCrLf
                        sSql = sSql & "                      ,SGI_DATAORDEM" & vbCrLf
                        sSql = sSql & "                      ,SGI_IDPRODUTO" & vbCrLf
                        sSql = sSql & "                      ,SGI_CODPROD" & vbCrLf
                        sSql = sSql & "                      ,SGI_CODPED" & vbCrLf
                        sSql = sSql & "                      ,SGI_DATAPED" & vbCrLf
                        
                        sSql = sSql & "                      ,SGI_QTDEPED" & vbCrLf
                        sSql = sSql & "                      ,SGI_QTDEJAPROG" & vbCrLf
                        sSql = sSql & "                      ,SGI_QTDE" & vbCrLf
                        sSql = sSql & "                      ,SGI_SALDO" & vbCrLf
                        ''sSql = sSql & "                      ,SGI_STATUS" & vbCrLf
                        sSql = sSql & "                      ,SGI_TIPO" & vbCrLf
                        
                        sSql = sSql & "                      ,SGI_ALTFILM" & vbCrLf
                        sSql = sSql & "                      ,SGI_FOTNOVO" & vbCrLf
                        sSql = sSql & "                      ,SGI_REPETICAO" & vbCrLf
                        sSql = sSql & "                      ,SGI_DATENTREGA" & vbCrLf
                        sSql = sSql & "                      ,SGI_TIPOP" & vbCrLf
                        sSql = sSql & "                      ,SGI_JAIMPRESSA" & vbCrLf
                        sSql = sSql & "                      ,SGI_OPCOMPL" & vbCrLf
                        sSql = sSql & "                      ,SGI_NOMEVEND" & vbCrLf
                        sSql = sSql & "                      ,SGI_OBSOP" & vbCrLf
                        sSql = sSql & "                      ,SGI_FECHTPFU" & vbCrLf
                        sSql = sSql & "                      ,SGI_FILIALPED" & vbCrLf
                        sSql = sSql & "                      ,SGI_ALTERADO" & vbCrLf
                        sSql = sSql & "                      ,SGI_IDPAI" & vbCrLf
                        sSql = sSql & "                      ,SGI_QTDEPORPALHETS" & vbCrLf
                        sSql = sSql & "                      ,SGI_PALHETS" & vbCrLf
                        sSql = sSql & "                      ,SGI_STATUS" & vbCrLf
                        
                        sSql = sSql & "                    ) Values (" & vbCrLf
                        
                        sSql = sSql & "                       " & intFILIAL & vbCrLf
                        sSql = sSql & "                      ," & arrPROGENTREGAS(i, 4) & vbCrLf
                        sSql = sSql & "                      ,'" & Format(CDate(arrPROGENTREGAS(i, 5)), "MM/DD/YYYY") & "'" & vbCrLf
                        sSql = sSql & "                      ," & arrPROGENTREGAS(i, 0) & vbCrLf
                        sSql = sSql & "                      ,'" & arrPROGENTREGAS(i, 6) & "'" & vbCrLf
                        sSql = sSql & "                      ," & lngCODPEDIDO & vbCrLf
                        sSql = sSql & "                      ,'" & Format(CDate(dtDATAPED), "MM/DD/YYYY") & "'" & vbCrLf
                        
                        sValor = "Null"
                        If Len(Trim(arrPROGENTREGAS(i, 1))) > 0 And arrPROGENTREGAS(i, 1) > 0 Then
                            sValor = Replace(arrPROGENTREGAS(i, 1), ".", "")
                            sValor = Replace(sValor, ",", ".")
                        End If
                        sSql = sSql & "                      ," & sValor & vbCrLf
                        
                        sValor = "Null"
                        sSql = sSql & "                      ," & sValor & vbCrLf
                        
                        sValor = "Null"
                        If Len(Trim(arrPROGENTREGAS(i, 1))) > 0 And arrPROGENTREGAS(i, 1) > 0 Then
                            sValor = Replace(arrPROGENTREGAS(i, 1), ".", "")
                            sValor = Replace(sValor, ",", ".")
                        End If
                        sSql = sSql & "                      ," & sValor & vbCrLf
                        
                        sValor = "Null"
                        sSql = sSql & "                      ," & sValor & vbCrLf
                        ''sSql = sSql & "                      ,0" & vbCrLf
                        sSql = sSql & "                      ,1" & vbCrLf
                        
                        sSql = sSql & "                      ," & arrPROGENTREGAS(i, 7) & vbCrLf
                        sSql = sSql & "                      ," & arrPROGENTREGAS(i, 8) & vbCrLf
                        sSql = sSql & "                      ," & arrPROGENTREGAS(i, 9) & vbCrLf
                        
                        sSql = sSql & "                      ,'" & Format(arrPROGENTREGAS(i, 2), "MM/DD/YYYY") & "'" & vbCrLf
                        sSql = sSql & "                      ,1" & vbCrLf
                        sSql = sSql & "                      ,0" & vbCrLf
                        sSql = sSql & "                      ,0" & vbCrLf
                        sSql = sSql & "                      ,'" & Trim(strNomeVend) & "'" & vbCrLf
                        sSql = sSql & "                      ,'" & Trim(arrPROGENTREGAS(i, 10)) & "'" & vbCrLf
                        sSql = sSql & "                      ," & Trim(arrPROGENTREGAS(i, 11)) & vbCrLf
                        sSql = sSql & "                      ," & intFILIALPED & vbCrLf
                        sSql = sSql & "                      ,0" & vbCrLf
                        sSql = sSql & "                      ," & Trim(arrPROGENTREGAS(i, 14)) & vbCrLf
                        sSql = sSql & "                      ," & arrPROGENTREGAS(i, 15) & vbCrLf
                        sSql = sSql & "                      ," & arrPROGENTREGAS(i, 16) & vbCrLf
                        sSql = sSql & "                      ," & arrPROGENTREGAS(i, 17) & vbCrLf
                        
                        sSql = sSql & "                    )"
                        
                        BGRV.CommandText = sSql
                        BGRV.Execute
                    
                        '' =============================
                        '' Dia 19/04/2015
                        '' Dando Update na Tabela do Item do Pedido ( Progrmação de Entrega"
                        
                        '' =====================================
                        '' Programação de Entregas
                        sSql = ""
                        
                        sSql = "Update SGI_PROGENTRPROD Set" & vbCrLf
                        sSql = sSql & "                       SGI_STATUS = " & arrPROGENTREGAS(i, 17) & vbCrLf
                        sSql = sSql & "                     Where " & vbCrLf
                        sSql = sSql & "                           SGI_FILIAL    = " & intFILIAL & vbCrLf
                        sSql = sSql & "                       And SGI_CODPED    = " & lngCODPEDIDO & vbCrLf
                        sSql = sSql & "                       And SGI_IDPRODUTO = " & arrPROGENTREGAS(i, 0) & vbCrLf
                        sSql = sSql & "                       And SGI_IDINTERNO = " & arrPROGENTREGAS(i, 14) & vbCrLf
                        
                        BGRV.CommandText = sSql
                        BGRV.Execute
                    
                    Else
                    
                        If boolFOIREPROVADO = True Then '' Foi reprovado e agora esta voltando a OP
                            If Len(Trim(arrPROGENTREGAS(i, 9))) > 0 Then
                                
                                sSql = ""
                                
                                sSql = "Update SGI_PROGENTRPROD Set " & vbCrLf
                                sSql = sSql & "                               SGI_STATUS = 0" & vbCrLf
                                sSql = sSql & "      Where " & vbCrLf
                                sSql = sSql & "            SGI_FILIAL    = " & intFILIAL & vbCrLf
                                sSql = sSql & "        And SGI_IDINTERNO = " & arrPROGENTREGAS(i, 16)
                                
                                BGRV.CommandText = sSql
                                BGRV.Execute
                                
                                sSql = ""
                                
                                sSql = "Update SGI_ORDEMPROD " & vbCrLf
                                sSql = sSql & "                           Set SGI_STATUS = 0 " & vbCrLf
                                sSql = sSql & "                     Where " & vbCrLf
                                sSql = sSql & "                           SGI_FILIAL = " & intFILIAL & vbCrLf
                                sSql = sSql & "                       And SGI_CODIGO = " & arrPROGENTREGAS(i, 9)
                                
                                BGRV.CommandText = sSql
                                BGRV.Execute
                                
                            End If
                        End If
                    
                    End If
                
                Next i
            End If
            '' =====================================================================
        Else
        
            '' Atualizando a Confirmação dos Itens
            If IsArray(arrPRODUTOS) Then
                For i = 1 To UBound(arrPRODUTOS)
                    sSql = ""
                    
                    sSql = "Update SGI_CADPEDVENDI Set " & vbCrLf
                    sSql = sSql & "                           SGI_CONFERIDO = " & arrPRODUTOS(i, 21) & vbCrLf
                    
                    sSql = sSql & "Where" & vbCrLf
                    sSql = sSql & "      SGI_FILIAL    = " & intFILIAL & vbCrLf
                    sSql = sSql & "  And SGI_CODIGO    = " & lngCODPEDIDO & vbCrLf
                    sSql = sSql & "  And SGI_IDPRODUTO = " & arrPRODUTOS(i, 9)
            
                    BGRV.CommandText = sSql
                    BGRV.Execute
            
                Next i
            End If
            
            
            '' Gravando Itens da Ordem
            If IsArray(arrPROGENTREGAS) Then
                For i = 1 To UBound(arrPROGENTREGAS)
                    
                        If boolFOIREPROVADO = True Then '' Foi reprovado e agora esta voltando a OP
                            If Len(Trim(arrPROGENTREGAS(i, 9))) > 0 Then
                                
                                sSql = ""
                                
                                sSql = "Update SGI_PROGENTRPROD_STEEL Set " & vbCrLf
                                sSql = sSql & "                               SGI_STATUS = 0" & vbCrLf
                                sSql = sSql & "      Where " & vbCrLf
                                sSql = sSql & "            SGI_FILIAL    = " & intFILIAL & vbCrLf
                                sSql = sSql & "        And SGI_IDINTERNO = " & arrPROGENTREGAS(i, 16)
                                
                                BGRV.CommandText = sSql
                                BGRV.Execute
                                
                                sSql = ""
                                
                                sSql = "Update SGI_ORDEMPROD " & vbCrLf
                                sSql = sSql & "                           Set SGI_STATUS = 0 " & vbCrLf
                                sSql = sSql & "                     Where " & vbCrLf
                                sSql = sSql & "                           SGI_FILIAL = " & intFILIAL & vbCrLf
                                sSql = sSql & "                       And SGI_CODIGO = " & arrPROGENTREGAS(i, 9)
                                
                                BGRV.CommandText = sSql
                                BGRV.Execute
                                
                            End If
                        End If
                
                Next i
            End If
            '' =====================================================================
        
        
        End If
     
        '' =====================================
        '' Grava Log
        sSql = ""
        
        sSql = "Insert Into SGI_LOGPED (" & vbCrLf
        sSql = sSql & "                              SGI_FILIAL" & vbCrLf
        sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                             ,SGI_DATA" & vbCrLf
        sSql = sSql & "                             ,SGI_HORA" & vbCrLf
        sSql = sSql & "                             ,SGI_CODUSUARIO" & vbCrLf
        sSql = sSql & "                             ,SGI_ACAO" & vbCrLf
        sSql = sSql & "                    ) Values (" & vbCrLf
        sSql = sSql & "                              " & FILIAL & vbCrLf
        sSql = sSql & "                             ," & lngCODPEDIDO & vbCrLf
        sSql = sSql & "                             ,'" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                             ,'" & Format(Now, "HH:MM:SS") & "'" & vbCrLf
        sSql = sSql & "                             ," & lngCodUsuario & vbCrLf
        sSql = sSql & "                             ,'" & strOPER & "'" & vbCrLf
        sSql = sSql & "                             )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' =====================================
     
     ElseIf strOPER = "R" Then '' Reprova Pedido Bloqueado
        
        sSql = "Delete from SGI_TIPREPRPED " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "           SGI_FILIAL = " & intFILIAL & vbCrLf
        sSql = sSql & "       And SGI_CODPED = " & lngCODPEDIDO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        sSql = "Update SGI_CADPEDVENDH Set " & vbCrLf
        sSql = sSql & "                           SGI_REPROCADO   = '" & strLIBERADO & "'" & vbCrLf
        sSql = sSql & "                          ,SGI_REPDATAHORA = '" & Format(dtDTHORALIB, "MM/DD/YYYY HH:MM:SS") & "'" & vbCrLf
        sSql = sSql & "                          ,SGI_OBSLIBCOM   = '" & strOBSCOM & "'" & vbCrLf
        sSql = sSql & "                          ,SGI_STATUS      = 'R'" & vbCrLf
        sSql = sSql & "                          ,SGI_CONFERIDO   = " & intCONFERIDO & vbCrLf
        
        sSql = sSql & "                     Where " & vbCrLf & vbCrLf
        sSql = sSql & "                           SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "                       And SGI_CODIGO = " & lngCODPEDIDO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' ------------------------------------------------------
        '' Gravando Tipo de Reprovação se Olver
        If IsArray(arrTIPREPROVA) = True Then
           For i = 1 To UBound(arrTIPREPROVA)
        
               sSql = "Insert into SGI_TIPREPRPED (" & vbCrLf
               sSql = sSql & "                     SGI_FILIAL" & vbCrLf
               sSql = sSql & "                    ,SGI_CODPED" & vbCrLf
               sSql = sSql & "                    ,SGI_CODREP" & vbCrLf
               sSql = sSql & "                    ,SGI_FILIALPED" & vbCrLf
               sSql = sSql & "                 ) Values (" & vbCrLf
               sSql = sSql & "                            " & intFILIAL & vbCrLf
               sSql = sSql & "                           ," & lngCODPEDIDO & vbCrLf
               sSql = sSql & "                           ," & arrTIPREPROVA(i) & vbCrLf
               sSql = sSql & "                           ," & intFILIALPED & vbCrLf
               sSql = sSql & "                            )"
        
               BGRV.CommandText = sSql
               BGRV.Execute
               
           Next i
        End If
        '' ------------------------------------------------------
        
        '' --------------------------------------------
        '' Mudando o Status da OP Geradas Pelo Sistema para Reprovada
        If IsArray(arrPROGENTREGAS) Then
            For i = 1 To UBound(arrPROGENTREGAS)
                If Len(Trim(Replace(arrPROGENTREGAS(i, 6), "/", ""))) > 0 Then
                    
                    sSql = ""
                    
                    sSql = "Update SGI_PROGENTRPROD Set " & vbCrLf
                    sSql = sSql & "                               SGI_STATUS = 3" & vbCrLf
                    sSql = sSql & "      Where " & vbCrLf
                    sSql = sSql & "            SGI_FILIAL    = " & intFILIAL & vbCrLf
                    sSql = sSql & "        And SGI_IDINTERNO = " & arrPROGENTREGAS(i, 16)
                    
                    BGRV.CommandText = sSql
                    BGRV.Execute
                    
                    
                    
                    sSql = ""
            
                    sSql = "Update SGI_ORDEMPROD Set " & vbCrLf
                    sSql = sSql & "                               SGI_STATUS = 3"
                    sSql = sSql & "      Where " & vbCrLf
                    sSql = sSql & "            SGI_FILIAL = " & intFILIAL & vbCrLf
                    sSql = sSql & "        And SGI_CODIGO = " & Trim(Replace(arrPROGENTREGAS(i, 6), "/", ""))
                    
                    BGRV.CommandText = sSql
                    BGRV.Execute
                    '' --------------------------------------------
                End If
            Next i
        End If
        
        '' --------------------------------------------
        '' Apagando as OP Geradas Pelo Sistema
        ''sSql = "Delete From SGI_ORDEMPROD " & vbCrLf
        ''sSql = sSql & "      Where " & vbCrLf
        ''sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        ''sSql = sSql & "        And SGI_CODPED = " & lngCODPEDIDO
        
        ''BGRV.CommandText = sSql
        ''BGRV.Execute
        '' --------------------------------------------
     
        '' =====================================
        '' Grava Log
        sSql = ""
        
        sSql = "Insert Into SGI_LOGPED (" & vbCrLf
        sSql = sSql & "                              SGI_FILIAL" & vbCrLf
        sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                             ,SGI_DATA" & vbCrLf
        sSql = sSql & "                             ,SGI_HORA" & vbCrLf
        sSql = sSql & "                             ,SGI_CODUSUARIO" & vbCrLf
        sSql = sSql & "                             ,SGI_ACAO" & vbCrLf
        sSql = sSql & "                    ) Values (" & vbCrLf
        sSql = sSql & "                              " & FILIAL & vbCrLf
        sSql = sSql & "                             ," & lngCODPEDIDO & vbCrLf
        sSql = sSql & "                             ,'" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                             ,'" & Format(Now, "HH:MM:SS") & "'" & vbCrLf
        sSql = sSql & "                             ," & lngCodUsuario & vbCrLf
        sSql = sSql & "                             ,'" & strOPER & "'" & vbCrLf
        sSql = sSql & "                             )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' =====================================
     
     ElseIf strOPER = "D" Then '' Bloqueia o pedido manualmente
        
        sSql = "Update SGI_CADPEDVENDH Set " & vbCrLf
        sSql = sSql & "                           SGI_REPROCADO   = '" & strLIBERADO & "'" & vbCrLf
        sSql = sSql & "                          ,SGI_REPDATAHORA = '" & Format(dtDTHORALIB, "MM/DD/YYYY HH:MM:SS") & "'" & vbCrLf
        sSql = sSql & "                          ,SGI_OBSLIBCOM   = '" & strOBSCOM & "'" & vbCrLf
        sSql = sSql & "                          ,SGI_STATUS      = 'S'" & vbCrLf
        
        sSql = sSql & "                     Where " & vbCrLf & vbCrLf
        sSql = sSql & "                           SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "                       And SGI_CODIGO = " & lngCODPEDIDO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        ''  --------------------------------------------
     
        '' --------------------------------------------
        '' Mudando o Status da OP Geradas Pelo Sistema para Reprovada
        If IsArray(arrPROGENTREGAS) Then
            For i = 1 To UBound(arrPROGENTREGAS)
        
                If Len(Trim(Replace(arrPROGENTREGAS(i, 6), "/", ""))) > 0 Then
        
                    
                    sSql = ""
            
                    sSql = "Update SGI_PROGENTRPROD Set " & vbCrLf
                    sSql = sSql & "                               SGI_STATUS = 4" & vbCrLf
                    sSql = sSql & "      Where " & vbCrLf
                    sSql = sSql & "            SGI_FILIAL    = " & intFILIAL & vbCrLf
                    sSql = sSql & "        And SGI_IDINTERNO = " & arrPROGENTREGAS(i, 16)
                    
                    BGRV.CommandText = sSql
                    BGRV.Execute
                    
                    sSql = ""
            
                    sSql = "Update SGI_ORDEMPROD Set " & vbCrLf
                    sSql = sSql & "                               SGI_STATUS = 4" & vbCrLf
                    sSql = sSql & "      Where " & vbCrLf
                    sSql = sSql & "            SGI_FILIAL = " & intFILIAL & vbCrLf
                    sSql = sSql & "        And SGI_CODIGO = " & Trim(Replace(arrPROGENTREGAS(i, 6), "/", ""))
                    
                    BGRV.CommandText = sSql
                    BGRV.Execute
                    '' --------------------------------------------
                
                End If
            Next i
        End If
     
     
        '' =====================================
        '' Grava Log
        sSql = ""
        
        sSql = "Insert Into SGI_LOGPED (" & vbCrLf
        sSql = sSql & "                              SGI_FILIAL" & vbCrLf
        sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                             ,SGI_DATA" & vbCrLf
        sSql = sSql & "                             ,SGI_HORA" & vbCrLf
        sSql = sSql & "                             ,SGI_CODUSUARIO" & vbCrLf
        sSql = sSql & "                             ,SGI_ACAO" & vbCrLf
        sSql = sSql & "                    ) Values (" & vbCrLf
        sSql = sSql & "                              " & FILIAL & vbCrLf
        sSql = sSql & "                             ," & lngCODPEDIDO & vbCrLf
        sSql = sSql & "                             ,'" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                             ,'" & Format(Now, "HH:MM:SS") & "'" & vbCrLf
        sSql = sSql & "                             ," & lngCodUsuario & vbCrLf
        sSql = sSql & "                             ,'" & strOPER & "'" & vbCrLf
        sSql = sSql & "                             )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' =====================================
     
     ElseIf strOPER = "T" Then
     
        If IsArray(arrSERVICOS) Then
            For i = 1 To UBound(arrSERVICOS)
            
                '' Apagando Confirmação de Faturamento Itens
                sSql = "Delete from SGI_CADORDCONFI " & vbCrLf
                sSql = sSql & "      Where " & vbCrLf
                sSql = sSql & "           SGI_FILIAL  = " & intFILIAL & vbCrLf
                sSql = sSql & "       And SGI_CODCONF = " & arrSERVICOS(i)
            
                BGRV.CommandText = sSql
                BGRV.Execute
                
                '' Apagando Confirmação de Faturamento Cabec
                sSql = "Delete from SGI_CADORDCONFH " & vbCrLf
                sSql = sSql & "      Where " & vbCrLf
                sSql = sSql & "           SGI_FILIAL  = " & intFILIAL & vbCrLf
                sSql = sSql & "       And SGI_CODCONF = " & arrSERVICOS(i)

                BGRV.CommandText = sSql
                BGRV.Execute
            
            Next i
        End If
        
        If IsArray(arrPROGENTREGAS) Then
            For i = 1 To UBound(arrPROGENTREGAS)
            
                '' Apagando Ordem de Faturamento Itens
                sSql = "Delete from SGI_CADORDFATI " & vbCrLf
                sSql = sSql & "      Where " & vbCrLf
                sSql = sSql & "           SGI_FILIAL = " & intFILIAL & vbCrLf
                sSql = sSql & "       And SGI_CODORD = " & arrPROGENTREGAS(i)
            
                BGRV.CommandText = sSql
                BGRV.Execute
                
                '' Apagando Ordem de Faturamento Cabec
                sSql = "Delete from SGI_CADORDFATH " & vbCrLf
                sSql = sSql & "      Where " & vbCrLf
                sSql = sSql & "           SGI_FILIAL = " & intFILIAL & vbCrLf
                sSql = sSql & "       And SGI_CODORD = " & arrPROGENTREGAS(i)
            
                BGRV.CommandText = sSql
                BGRV.Execute
            
            Next i
        End If
        
        
        If IsArray(arrPRODUTOS) Then
        
            For i = 1 To UBound(arrPRODUTOS)
        
                '' Apagando Ordem de Produção
                sSql = "Delete from SGI_ORDEMPROD " & vbCrLf
                sSql = sSql & "      Where " & vbCrLf
                sSql = sSql & "           SGI_FILIAL = " & intFILIAL & vbCrLf
                sSql = sSql & "       And SGI_CODPED = " & arrPRODUTOS(i)
            
                BGRV.CommandText = sSql
                BGRV.Execute
                
                '' Apagando Programação de Entegras
                sSql = "Delete from SGI_PROGENTRPROD " & vbCrLf
                sSql = sSql & "      Where " & vbCrLf
                sSql = sSql & "           SGI_FILIAL = " & intFILIAL & vbCrLf
                sSql = sSql & "       And SGI_CODPED = " & arrPRODUTOS(i)
            
                BGRV.CommandText = sSql
                BGRV.Execute
                
                '' Apagando Itens de Pedido
                sSql = "Delete from SGI_CADPEDVENDI " & vbCrLf
                sSql = sSql & "      Where " & vbCrLf
                sSql = sSql & "           SGI_FILIAL = " & intFILIAL & vbCrLf
                sSql = sSql & "       And SGI_CODIGO = " & arrPRODUTOS(i)
            
                BGRV.CommandText = sSql
                BGRV.Execute
                
                '' Apagando Pedido
                sSql = "Delete from SGI_CADPEDVENDH " & vbCrLf
                sSql = sSql & "      Where " & vbCrLf
                sSql = sSql & "           SGI_FILIAL = " & intFILIAL & vbCrLf
                sSql = sSql & "       And SGI_CODIGO = " & arrPRODUTOS(i)
                
                BGRV.CommandText = sSql
                BGRV.Execute
            
            
            Next i
            
        End If
        
     ElseIf strOPER = "M" Then
        
        If strSTATUS = "L" Or strSTATUS = "P" Then
            sSql = "Update SGI_CADPEDVENDH Set " & vbCrLf
            sSql = sSql & "                           SGI_STATUS      = 'M'" & vbCrLf
            
            sSql = sSql & "                     Where " & vbCrLf & vbCrLf
            sSql = sSql & "                           SGI_FILIAL = " & FILIAL & vbCrLf
            sSql = sSql & "                       And SGI_CODIGO = " & lngCODPEDIDO
            
            BGRV.CommandText = sSql
            BGRV.Execute
        End If
        '' Dando Update Nas OPS
        If IsArray(arrOPS) Then
            For i = 1 To UBound(arrOPS)
                sSql = "Update SGI_ORDEMPROD Set " & vbCrLf
                sSql = sSql & "                           SGI_STATUS      = 2" & vbCrLf
                
                sSql = sSql & "                     Where " & vbCrLf & vbCrLf
                sSql = sSql & "                           SGI_FILIAL = " & FILIAL & vbCrLf
                sSql = sSql & "                       And SGI_CODIGO = " & arrOPS(i)
                
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
     
     
        '' =====================================
        '' Grava Log
        sSql = ""
        
        sSql = "Insert Into SGI_LOGPED (" & vbCrLf
        sSql = sSql & "                              SGI_FILIAL" & vbCrLf
        sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                             ,SGI_DATA" & vbCrLf
        sSql = sSql & "                             ,SGI_HORA" & vbCrLf
        sSql = sSql & "                             ,SGI_CODUSUARIO" & vbCrLf
        sSql = sSql & "                             ,SGI_ACAO" & vbCrLf
        sSql = sSql & "                    ) Values (" & vbCrLf
        sSql = sSql & "                              " & FILIAL & vbCrLf
        sSql = sSql & "                             ," & lngCODPEDIDO & vbCrLf
        sSql = sSql & "                             ,'" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                             ,'" & Format(Now, "HH:MM:SS") & "'" & vbCrLf
        sSql = sSql & "                             ," & lngCodUsuario & vbCrLf
        sSql = sSql & "                             ,'" & strOPER & "'" & vbCrLf
        sSql = sSql & "                             )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' =====================================
     
     ElseIf strOPER = "S" Then
        '' Liberando da Alteração
        sSql = "Update SGI_CADPEDVENDH Set " & vbCrLf
        sSql = sSql & "                           SGI_STATUS      = 'B'" & vbCrLf
        sSql = sSql & "                     Where " & vbCrLf & vbCrLf
        sSql = sSql & "                           SGI_FILIAL = " & intFILIAL & vbCrLf
        sSql = sSql & "                       And SGI_CODIGO = " & lngCODPEDIDO
        
        BGRV.CommandText = sSql
        BGRV.Execute
     
     
     ElseIf strOPER = "V" Or strOPER = "LC" Then
     
        '' Programação de Entregas
        If IsArray(arrPROGENTREGAS) Then
            For i = 1 To UBound(arrPROGENTREGAS)
            
                sSql = ""
                
                sSql = "Update SGI_PROGENTRPROD Set" & vbCrLf
                sSql = sSql & "                                 SGI_STATUS     = " & arrPROGENTREGAS(i, 17) & vbCrLf
                sSql = sSql & "                                ,SGI_DATENTREGA = " & arrPROGENTREGAS(i, 2) & vbCrLf
                sSql = sSql & " Where" & vbCrLf
                sSql = sSql & "       SGI_FILIAL    = " & intFILIAL & vbCrLf
                sSql = sSql & "   And SGI_CODPED    = " & lngCODPEDIDO
                sSql = sSql & "   And SGI_IDINTERNO = " & arrPROGENTREGAS(i, 16) & vbCrLf
            
                BGRV.CommandText = sSql
                BGRV.Execute
                
                sSql = ""
                
                sSql = "Update SGI_ORDEMPROD Set " & vbCrLf
                sSql = sSql & "                           SGI_DATENTREGA      = " & arrPROGENTREGAS(i, 2) & vbCrLf
                sSql = sSql & "                          ,SGI_STATUS          = " & arrPROGENTREGAS(i, 17) & vbCrLf
                
                sSql = sSql & "                     Where " & vbCrLf
                sSql = sSql & "                           SGI_FILIAL    = " & FILIAL & vbCrLf
                sSql = sSql & "                       And SGI_IDPRODUTO = " & arrPROGENTREGAS(i, 0) & vbCrLf
                sSql = sSql & "                       And SGI_IDPAI     = " & arrPROGENTREGAS(i, 16) & vbCrLf
                
                BGRV.CommandText = sSql
                BGRV.Execute
                
            
                '' Liberando Fotolito
                sSql = "Update SGI_CADPEDVENDH Set " & vbCrLf
                
                If strOPER = "V" Then
                    If (arrPROGENTREGAS(i, 17) = 7 Or arrPROGENTREGAS(i, 17) = 6) Then
                       sSql = sSql & "                           SGI_STATUS      = '" & strSTATUS & "'" & vbCrLf
                    Else
                       sSql = sSql & "                           SGI_STATUS      = 'B'" & vbCrLf
                    End If
                ElseIf strOPER = "LC" Then
                    strSTATUS = "L"
                    sSql = sSql & "                           SGI_STATUS      = '" & strSTATUS & "'" & vbCrLf
                End If
                
                sSql = sSql & "                     Where " & vbCrLf & vbCrLf
                sSql = sSql & "                           SGI_FILIAL = " & intFILIAL & vbCrLf
                sSql = sSql & "                       And SGI_CODIGO = " & lngCODPEDIDO
                
                BGRV.CommandText = sSql
                BGRV.Execute
            
            
            Next i
        End If
        
        '' Liberando da Fotolito
        ''sSql = "Update SGI_CADPEDVENDH Set " & vbCrLf
        
        ''If strOPER = "V" Then
        ''    sSql = sSql & "                           SGI_STATUS      = 'B'" & vbCrLf
        ''ElseIf strOPER = "LC" Then
        ''    sSql = sSql & "                           SGI_STATUS      = '" & strSTATUS & "'" & vbCrLf
        ''End If
        
        ''sSql = sSql & "                     Where " & vbCrLf & vbCrLf
        ''sSql = sSql & "                           SGI_FILIAL = " & intFilial & vbCrLf
        ''sSql = sSql & "                       And SGI_CODIGO = " & lngCODPEDIDO
        
        ''BGRV.CommandText = sSql
        ''BGRV.Execute
     
        '' =====================================
        '' Grava Log
        sSql = ""
        
        sSql = "Insert Into SGI_LOGPED (" & vbCrLf
        sSql = sSql & "                              SGI_FILIAL" & vbCrLf
        sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                             ,SGI_DATA" & vbCrLf
        sSql = sSql & "                             ,SGI_HORA" & vbCrLf
        sSql = sSql & "                             ,SGI_CODUSUARIO" & vbCrLf
        sSql = sSql & "                             ,SGI_ACAO" & vbCrLf
        sSql = sSql & "                    ) Values (" & vbCrLf
        sSql = sSql & "                              " & FILIAL & vbCrLf
        sSql = sSql & "                             ," & lngCODPEDIDO & vbCrLf
        sSql = sSql & "                             ,'" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                             ,'" & Format(Now, "HH:MM:SS") & "'" & vbCrLf
        sSql = sSql & "                             ," & lngCodUsuario & vbCrLf
        sSql = sSql & "                             ,'" & strOPER & "'" & vbCrLf
        sSql = sSql & "                             )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' =====================================
     
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

Public Function Carrega_Campos() As Boolean

    Carrega_Campos = False

    Dim i         As Integer
    Dim intQTDREG As Integer
    Dim strNOMTAB As String
    
    If intFILIALPED = 0 Then
        strNOMTAB = "SGI_CADPEDVENDH"
    ElseIf intFILIALPED = 1 Then
        strNOMTAB = "SGI_CADPEDVENDH_STEEL"
    End If
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & strNOMTAB & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCODPEDIDO
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       BREC.Close
       Exit Function
    End If
    
    lngCODPEDIDO = BREC!SGI_CODIGO
    dtDATAPED = BREC!SGI_DATAPED
    intCODCLIE = BREC!SGI_CODCLI
    strSTATUS = BREC!SGI_STATUS
    
    boolFOIREPROVADO = False
    If BREC!SGI_STATUS = "R" Then boolFOIREPROVADO = True
    
    
    dtDATAENTREGA = BREC!SGI_DTENTREGA
    If Not IsNull(BREC!SGI_DTCALIBRA) Then dtDATCALIBR = BREC!SGI_DTCALIBRA
    If Not IsNull(BREC!SGI_CODTEC) Then intCODTECNICO = BREC!SGI_CODTEC
    
    intCODVEND = BREC!SGI_CODVEND
    intCODCONDPGTO = BREC!SGI_CODCONDPGT
    intTIPPED = BREC!SGI_CODTIPORC
    
    If Not IsNull(BREC!SGI_ENDENTRE) Then strENDENTR = BREC!SGI_ENDENTRE
    If Not IsNull(BREC!SGI_BAIENTRE) Then strBAIENTR = BREC!SGI_BAIENTRE
    If Not IsNull(BREC!SGI_CIDENTRE) Then strCIDENTR = BREC!SGI_CIDENTRE
    If Not IsNull(BREC!SGI_CEPENTRE) Then strCEPENTREGA = BREC!SGI_CEPENTRE
    If Not IsNull(BREC!SGI_ESTENTRE) Then intESTENTREGA = BREC!SGI_ESTENTRE
    If Not IsNull(BREC!SGI_TELENTR) Then strTELENTR = BREC!SGI_TELENTR
    If Not IsNull(BREC!SGI_FAXENTR) Then strFAXENTR = BREC!SGI_FAXENTR
    
    If Not IsNull(BREC!SGI_ENDCOBR) Then strENDCOBRA = BREC!SGI_ENDCOBR
    If Not IsNull(BREC!SGI_BAICOBR) Then strBAICOBRA = BREC!SGI_BAICOBR
    If Not IsNull(BREC!SGI_CIDCOBR) Then strCIDCOBRA = BREC!SGI_CIDCOBR
    If Not IsNull(BREC!SGI_CEPCOBR) Then strCEPCOBRA = BREC!SGI_CEPCOBR
    If Not IsNull(BREC!SGI_ESTCOBR) Then intESTCOBRA = BREC!SGI_ESTCOBR
    If Not IsNull(BREC!SGI_TELCOBR) Then strTELCOBRA = BREC!SGI_TELCOBR
    If Not IsNull(BREC!SGI_FAXCOBR) Then strFAXCOBRA = BREC!SGI_FAXCOBR
    
    If Not IsNull(BREC!SGI_CONTATO) Then strCONTATO = BREC!SGI_CONTATO
    If Not IsNull(BREC!SGI_DEPTO) Then strDEPARTAMENTO = BREC!SGI_DEPTO
    If Not IsNull(BREC!SGI_ORDCOCLI) Then strORDCOMPCLI = BREC!SGI_ORDCOCLI
    If Not IsNull(BREC!SGI_CODTRANSP) Then lngCODTRANSP = BREC!SGI_CODTRANSP
    
    If Not IsNull(BREC!SGI_BASEICMS) Then curVALBASICMS = BREC!SGI_BASEICMS
    If Not IsNull(BREC!SGI_PORCICMS) Then curAliqICMS = BREC!SGI_PORCICMS
    If Not IsNull(BREC!SGI_VLICMS) Then curVLICMS = BREC!SGI_VLICMS
    If Not IsNull(BREC!SGI_OUTRASTOT) Then curOUTRDESPESAS = BREC!SGI_OUTRASTOT
    If Not IsNull(BREC!SGI_FRETETOT) Then curVLFRETE = BREC!SGI_FRETETOT
    If Not IsNull(BREC!SGI_VLIPI) Then curVLIPI = BREC!SGI_VLIPI
    If Not IsNull(BREC!SGI_VLDESCTO) Then curVLDESCTO = BREC!SGI_VLDESCTO
    
    If Not IsNull(BREC!SGI_PORCDESCPED) Then curPORDESC = BREC!SGI_PORCDESCPED
    If Not IsNull(BREC!SGI_VALDESCPED) Then curVALDESC = BREC!SGI_VALDESCPED
    
    If Not IsNull(BREC!SGI_VLTOT) Then curTOTORCTO = BREC!SGI_VLTOT
        
    If Not IsNull(BREC!SGI_PRZENTR) Then intPRZENTREGA = BREC!SGI_PRZENTR
    If Not IsNull(BREC!SGI_VALPROP) Then intVALPROPOSTA = BREC!SGI_VALPROP
    If Not IsNull(BREC!SGI_IPIISENTO) Then intIPIISENTO = BREC!SGI_IPIISENTO
    If Not IsNull(BREC!SGI_ICMSISENTO) Then intICMSISENTO = BREC!SGI_ICMSISENTO
    If Not IsNull(BREC!SGI_TIPFRETE) Then intFRETE = BREC!SGI_TIPFRETE
    
    If Not IsNull(BREC!SGI_CODCOTA) Then lngCODCOTA = BREC!SGI_CODCOTA
    
    If Not IsNull(BREC!SGI_LIBERADO) Then strLIBERADO = BREC!SGI_LIBERADO
    If Not IsNull(BREC!SGI_LIBDATAHORA) Then dtDTHORALIB = BREC!SGI_LIBDATAHORA
        
    If Not IsNull(BREC!SGI_REPROCADO) Then strREPROVADO = BREC!SGI_REPROCADO
    If Not IsNull(BREC!SGI_REPDATAHORA) Then dtDTHORAREPRO = BREC!SGI_REPDATAHORA
    
    If Not IsNull(BREC!SGI_EMAIL) Then strEMAIL = BREC!SGI_EMAIL
    If Not IsNull(BREC!SGI_OBS) Then strOBSERVACAO = BREC!SGI_OBS
        
    If Not IsNull(BREC!SGI_CODTIPODOC) Then lngCODTIPODOC = BREC!SGI_CODTIPODOC
    If Not IsNull(BREC!SGI_CODBANCO) Then lngCODBANCO = BREC!SGI_CODBANCO
    
    If Not IsNull(BREC!SGI_OBSLIBCOM) Then strOBSCOM = BREC!SGI_OBSLIBCOM
    If Not IsNull(BREC!SGI_OBSQUALID) Then strOBS2 = Trim(BREC!SGI_OBSQUALID)
    If Not IsNull(BREC!SGI_ESPECIAL) Then intESPECIAL = BREC!SGI_ESPECIAL
    If Not IsNull(BREC!SGI_PARAESTOQUE) Then intPARAESTOQUE = BREC!SGI_PARAESTOQUE
    
    If Not IsNull(BREC!SGI_CONFERIDO) Then intCONFERIDO = BREC!SGI_CONFERIDO
    
    If Not IsNull(BREC!SGI_CODMOTLIQ) Then strCODMOTLIQ = Trim(Str(BREC!SGI_CODMOTLIQ))
    If Not IsNull(BREC!SGI_OBSLIQ) Then strOBSLIQ = Trim(BREC!SGI_OBSLIQ)
    
    intPERMITEFECHOP = BREC!SGI_PERMFECHOP
    
    BREC.Close
    
    
    If intFILIALPED = 0 Then
        strNOMTAB = "SGI_CADPEDVENDI"
    ElseIf intFILIALPED = 1 Then
        strNOMTAB = "SGI_CADPEDVENDI_STEEL"
    End If
    
    '' -----------------------------------------------------------
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       ITEN.* " & vbCrLf
    sSql = sSql & "      ,PROD.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "      ,PROD.SGI_FechSoldaAgrafado" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_VernCorpo" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_VernTampa" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_VernFundo" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_VernArgola" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_FechTampaFuro" & vbCrLf
    sSql = sSql & "      ,PROD.SGI_STATUS" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & strNOMTAB & " ITEN " & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO  PROD " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       ITEN.SGI_FILIAL    = " & intFILIAL & vbCrLf
    sSql = sSql & "   And ITEN.SGI_CODIGO    = " & lngCODPEDIDO & vbCrLf
    sSql = sSql & "   And PROD.SGI_FILIAL    = ITEN.SGI_FILIAL  " & vbCrLf
    sSql = sSql & "   And PROD.SGI_IDPRODUTO = ITEN.SGI_IDPRODUTO " & vbCrLf
    
    BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC3.EOF Then
    
       intQTDREG = 0
       Do While Not BREC3.EOF
          intQTDREG = intQTDREG + 1
          BREC3.MoveNext
       Loop
    
       ReDim arrPRODUTOS(1 To intQTDREG, 1 To 27) As Variant
       
       intQTDREG = 1
       BREC3.MoveFirst
       
       Do While Not BREC3.EOF()
          
          arrPRODUTOS(intQTDREG, 1) = BREC3!SGI_CODPROD
          arrPRODUTOS(intQTDREG, 2) = BREC3!SGI_QTDE
          arrPRODUTOS(intQTDREG, 3) = BREC3!SGI_VLUNIT
          arrPRODUTOS(intQTDREG, 4) = BREC3!SGI_PRCDESC
          arrPRODUTOS(intQTDREG, 5) = BREC3!SGI_PRCIPI
          arrPRODUTOS(intQTDREG, 6) = BREC3!SGI_VLTOT
          arrPRODUTOS(intQTDREG, 7) = BREC3!SGI_VLDESC
          arrPRODUTOS(intQTDREG, 8) = BREC3!SGI_VLIPI
          arrPRODUTOS(intQTDREG, 9) = BREC3!SGI_DESCRICAO
          arrPRODUTOS(intQTDREG, 10) = BREC3!SGI_IDPRODUTO
          
          arrPRODUTOS(intQTDREG, 11) = 0
          arrPRODUTOS(intQTDREG, 12) = 0
          arrPRODUTOS(intQTDREG, 13) = 0
          arrPRODUTOS(intQTDREG, 14) = 0
          arrPRODUTOS(intQTDREG, 15) = 0
          arrPRODUTOS(intQTDREG, 16) = 0
          
          If Not IsNull(BREC3!SGI_FechSoldaAgrafado) Then arrPRODUTOS(intQTDREG, 11) = BREC3!SGI_FechSoldaAgrafado
          If Not IsNull(BREC3!SGI_VernCorpo) Then arrPRODUTOS(intQTDREG, 12) = BREC3!SGI_VernCorpo
          If Not IsNull(BREC3!SGI_VernTampa) Then arrPRODUTOS(intQTDREG, 13) = BREC3!SGI_VernTampa
          If Not IsNull(BREC3!SGI_VernFundo) Then arrPRODUTOS(intQTDREG, 14) = BREC3!SGI_VernFundo
          If Not IsNull(BREC3!SGI_VernArgola) Then arrPRODUTOS(intQTDREG, 15) = BREC3!SGI_VernArgola
          If Not IsNull(BREC3!SGI_FECHTPFU) Then arrPRODUTOS(intQTDREG, 16) = BREC3!SGI_FECHTPFU
          
          arrPRODUTOS(intQTDREG, 17) = "0"
          arrPRODUTOS(intQTDREG, 18) = "0"
          arrPRODUTOS(intQTDREG, 19) = "0"
          If Not IsNull(BREC3!SGI_ALTFILM) Then arrPRODUTOS(intQTDREG, 17) = Trim(Str(BREC3!SGI_ALTFILM))
          If Not IsNull(BREC3!SGI_FOTNOVO) Then arrPRODUTOS(intQTDREG, 18) = Trim(Str(BREC3!SGI_FOTNOVO))
          If Not IsNull(BREC3!SGI_REPETICAO) Then arrPRODUTOS(intQTDREG, 19) = Trim(Str(BREC3!SGI_REPETICAO))
          arrPRODUTOS(intQTDREG, 20) = BREC3!SGI_CODLINPROD
          
          If Not IsNull(BREC3!SGI_CODFORN) Then arrPRODUTOS(intQTDREG, 21) = BREC3!SGI_CODFORN
          
          arrPRODUTOS(intQTDREG, 22) = ""
          If Not IsNull(BREC3!SGI_OBSOP) Then arrPRODUTOS(intQTDREG, 22) = Trim(Replace(Replace(BREC3!SGI_OBSOP, vbTab, ""), vbCrLf, ""))
          arrPRODUTOS(intQTDREG, 23) = BREC3!SGI_STATUS
          
          arrPRODUTOS(intQTDREG, 24) = ""
          arrPRODUTOS(intQTDREG, 25) = ""
          If Not IsNull(BREC3!SGI_QTDEPORPALHETS) Then arrPRODUTOS(intQTDREG, 24) = BREC3!SGI_QTDEPORPALHETS
          If Not IsNull(BREC3!SGI_PALHETS) Then arrPRODUTOS(intQTDREG, 25) = BREC3!SGI_PALHETS
          
          arrPRODUTOS(intQTDREG, 26) = BREC3!SGI_CONFERIDO
          
          arrPRODUTOS(intQTDREG, 27) = 0
          If Not IsNull(BREC3!SGI_USARPADRPALSN) Then arrPRODUTOS(intQTDREG, 27) = BREC3!SGI_USARPADRPALSN
          
          intQTDREG = intQTDREG + 1
          
          BREC3.MoveNext
       Loop
    
    End If
    BREC3.Close
    
    '' ==================================================================
    '' Programacao de entrega
    
    If intFILIALPED = 0 Then
        strNOMTAB = "SGI_PROGENTRPROD"
    ElseIf intFILIALPED = 1 Then
        strNOMTAB = "SGI_PROGENTRPROD_STEEL"
    End If
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & strNOMTAB & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODPED = " & lngCODPEDIDO
    
    BREC6.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC6.EOF() Then
        intQTDREG = 0
        Do While Not BREC6.EOF()
            intQTDREG = (intQTDREG + 1)
            BREC6.MoveNext
        Loop
        
        ReDim arrPROGENTREGAS(1 To intQTDREG, 0 To 10) As String
        
        intQTDREG = 1
        BREC6.MoveFirst
        
        Do While Not BREC6.EOF()
            arrPROGENTREGAS(intQTDREG, 0) = BREC6!SGI_IDPRODUTO
            arrPROGENTREGAS(intQTDREG, 1) = BREC6!SGI_QTDE
            If Not IsNull(BREC6!SGI_DATENTREGA) Then arrPROGENTREGAS(intQTDREG, 2) = BREC6!SGI_DATENTREGA
            arrPROGENTREGAS(intQTDREG, 3) = dacEnumUpdateAction_Ignore
            If Not IsNull(BREC6!SGI_OBSOP) Then arrPROGENTREGAS(intQTDREG, 4) = Trim(Replace(BREC6!SGI_OBSOP, vbTab, ""))
            arrPROGENTREGAS(intQTDREG, 5) = BREC6!SGI_FECHTPFU
            arrPROGENTREGAS(intQTDREG, 6) = BREC6!SGI_INDICE
            If Not IsNull(BREC6!SGI_IDINTERNO) Then arrPROGENTREGAS(intQTDREG, 7) = BREC6!SGI_IDINTERNO
            
            arrPROGENTREGAS(intQTDREG, 8) = 0
            If Not IsNull(BREC6!SGI_STATUS) Then arrPROGENTREGAS(intQTDREG, 8) = BREC6!SGI_STATUS
            
            arrPROGENTREGAS(intQTDREG, 9) = ""
            arrPROGENTREGAS(intQTDREG, 10) = ""
            If Not IsNull(BREC6!SGI_QTDEPORPALHETS) Then arrPROGENTREGAS(intQTDREG, 9) = BREC6!SGI_QTDEPORPALHETS
            If Not IsNull(BREC6!SGI_PALHETS) Then arrPROGENTREGAS(intQTDREG, 10) = BREC6!SGI_PALHETS
            
            intQTDREG = (intQTDREG + 1)
            BREC6.MoveNext
        Loop
    End If
    BREC6.Close
    '' ==================================================================
    
    '' Tipos de Reprovação
    If intFILIALPED = 0 Then
        strNOMTAB = "SGI_TIPREPRPED"
    ElseIf intFILIALPED = 1 Then
        strNOMTAB = "SGI_TIPREPRPED_STEEL"
    End If
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & strNOMTAB & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & intFILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODPED    = " & lngCODPEDIDO & vbCrLf
    
    BREC4.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC4.EOF() Then
        intQTDREG = 0
        Do While Not BREC4.EOF()
            intQTDREG = intQTDREG + 1
            BREC4.MoveNext
        Loop
    
        ReDim arrTIPREPROVA(1 To intQTDREG) As Variant
        
        intQTDREG = 1
        BREC4.MoveFirst
    
        Do While Not BREC4.EOF()
            arrTIPREPROVA(intQTDREG) = BREC4!SGI_CODREP
            intQTDREG = intQTDREG + 1
            BREC4.MoveNext
        Loop
    End If
    BREC4.Close
    
    
    strNOMTAB = "SGI_LOGPED"
    If intFILIALPED = 1 Then strNOMTAB = "SGI_LOGPED_STEEL"
    
    '' Log Ação
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & strNOMTAB & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & intFILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO    = " & lngCODPEDIDO & vbCrLf
    sSql = sSql & "Order By SGI_DATA,SGI_HORA"
    
    BREC4.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC4.EOF() Then
        intQTDREG = 0
        Do While Not BREC4.EOF()
            intQTDREG = intQTDREG + 1
            BREC4.MoveNext
        Loop
    
        ReDim arrLOGPEDIDOS(1 To intQTDREG, 1 To 5) As Variant
        
        intQTDREG = 1
        BREC4.MoveFirst
    
        Do While Not BREC4.EOF()
            arrLOGPEDIDOS(intQTDREG, 1) = Format(BREC4!SGI_DATA, "DD/MM/YYYY")
            arrLOGPEDIDOS(intQTDREG, 2) = Format(BREC4!SGI_HORA, "HH:MM:SS")
            arrLOGPEDIDOS(intQTDREG, 3) = BREC4!SGI_CODUSUARIO
            arrLOGPEDIDOS(intQTDREG, 4) = BREC4!SGI_ACAO
            If Not IsNull(BREC4!SGI_DESCACAO) Then arrLOGPEDIDOS(intQTDREG, 5) = BREC4!SGI_DESCACAO
            
            intQTDREG = intQTDREG + 1
            BREC4.MoveNext
        Loop
    End If
    BREC4.Close
    
    
    
    Carrega_Campos = True
    ' ----------------------------------------------------------
    
End Function

Public Sub PreencheComboTipoReprova(Combo As Variant)

     Combo.Clear
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       * " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADTIPREP " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL  = " & intFILIAL & vbCrLf
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_DESCRI)
        Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
        BREC.MoveNext
     Loop
    
     BREC.Close
    
End Sub

Public Sub PreencheComboTransportadoras(Combo As Variant)

     Combo.Clear
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       * " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADTRANSP " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL  = " & intFILIAL & vbCrLf
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_DESCRICAO)
        Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
        BREC.MoveNext
     Loop
    
     BREC.Close
    
End Sub

Public Function Atualiza(strACAO As String, lngCODIGO As Long, lngFilial_Atu As Integer, strModulo As String, intFILIALPEDG As Integer) As Boolean
    
On Error GoTo Erro_Atualiza

    Atualiza = False
    
    '' Inicia transação
    adoBanco_Dados.BeginTrans
    BGRV.ActiveConnection = adoBanco_Dados
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "      *" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_ATUALIZA" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      SGI_FILIAL    = " & lngFilial_Atu & vbCrLf
    sSql = sSql & "  And SGI_MODULO    = '" & Trim(strModulo) & "'" & vbCrLf
    sSql = sSql & "  And SGI_FILIALPED = " & intFILIALPEDG
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF Then
       
        sSql = ""
       
        sSql = "Insert Into SGI_ATUALIZA (SGI_FILIAL" & vbCrLf
        sSql = sSql & "                  ,SGI_MODULO" & vbCrLf
        sSql = sSql & "                  ,SGI_ACAO" & vbCrLf
        sSql = sSql & "                  ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                  ,SGI_FILIALPED" & vbCrLf
        sSql = sSql & "         ) Values (" & vbCrLf
        sSql = sSql & "                   " & lngFilial_Atu & vbCrLf
        sSql = sSql & "                 ,'" & Trim(strModulo) & "'" & vbCrLf
        sSql = sSql & "                 ,'" & Trim(strACAO) & "'" & vbCrLf
        sSql = sSql & "                 ,'" & Trim(Str(lngCODIGO)) & "'" & vbCrLf
        sSql = sSql & "                 ," & intFILIALPEDG & vbCrLf
        sSql = sSql & "                   )"
    
        BGRV.CommandText = sSql
        BGRV.Execute
    
    Else
    
        sSql = ""
    
        sSql = sSql & "Update SGI_ATUALIZA Set" & vbCrLf
        sSql = sSql & "           SGI_ACAO   = '" & Trim(strACAO) & "'" & vbCrLf
        sSql = sSql & "          ,SGI_CODIGO = '" & Trim(Str(lngCODIGO)) & "'" & vbCrLf
        sSql = sSql & "Where " & vbCrLf
        sSql = sSql & "      SGI_FILIAL    = " & lngFilial_Atu & vbCrLf
        sSql = sSql & "  And SGI_MODULO    = '" & Trim(strModulo) & "'" & vbCrLf
        sSql = sSql & "  And SGI_FILIALPED = " & intFILIALPEDG
        
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
     Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, strACAO, sSql)
     Set objErro = Nothing
    
End Function

Public Sub PreencheComboTecnico(Combo As Variant)

     Combo.Clear
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       TEC.*" & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADTECNICO TEC" & vbCrLf
     
     BREC4.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC4.EOF
        Combo.AddItem Trim(BREC4!SGI_NOME)
        Combo.ItemData(Combo.NewIndex) = BREC4!SGI_CODIGO
        BREC4.MoveNext
     Loop
    
     BREC4.Close
    
End Sub


Public Sub PreencheCombo(Combo As Variant, strTabela As String)

     Combo.Clear
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       * " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       " & strTabela & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL  = " & intFILIAL & vbCrLf
     If UCase(Trim(strTabela)) = "SGI_CADTIPOPGTO" Then
        sSql = sSql & "   And SGI_OPERACAO  = 2" & vbCrLf
     End If
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_DESCRICAO)
        Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
        BREC.MoveNext
     Loop
    
     BREC.Close
    
End Sub

Public Sub PreenchComboBancos(Combo As Variant)

    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADBANCOS " & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & intFILIAL & vbCrLf
    sSql = sSql & "Order by SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       Combo.AddItem Trim(BREC!SGI_AGENCIA) & Space(10 - Len(Trim(BREC!SGI_AGENCIA))) & " - " & Trim(BREC!SGI_CC) & Space(20 - Len(Trim(BREC!SGI_CC))) & " - " & Trim(BREC!SGI_DESCRICAO)
       Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
       BREC.MoveNext
    Loop
    
    BREC.Close

End Sub

Public Sub PreenchComboGrpGastos(Combo As Variant)

    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADGRUPOGASTOS " & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & intFILIAL & vbCrLf
    sSql = sSql & "Order by SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       Combo.AddItem Trim(BREC!SGI_DESCRICAO)
       Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
       BREC.MoveNext
    Loop
    
    BREC.Close

End Sub

Public Function PegaQtdTotItPedido(lngCODCOTA As Long) As Long

    PegaQtdTotItPedido = 0
    
''    sSql = "Select " & vbCrLf
''    sSql = sSql & "       PED.* " & vbCrLf
''    sSql = sSql & "  From " & vbCrLf
''    sSql = sSql & "       SGI_CADPEDVENDI PED " & vbCrLf
''    sSql = sSql & " Where " & vbCrLf
''    sSql = sSql & "       PED.SGI_FILIAL = " & intFilial & vbCrLf
''    sSql = sSql & "   And PED.SGI_CODIGO = " & lngCODPEDIDO & vbCrLf
    
    sSql = "Select sum(ITE.SGI_QTDE) As SGI_QTDE" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_CADPEDVENDH PED" & vbCrLf
    sSql = sSql & "     ,SGI_CADPEDVENDI ITE" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      PED.SGI_FILIAL  = " & intFILIAL & vbCrLf
    sSql = sSql & "  And PED.SGI_CODCOTA = " & lngCODCOTA & vbCrLf
    sSql = sSql & "  And ITE.SGI_FILIAL  = PED.SGI_FILIAL" & vbCrLf
    sSql = sSql & "  And ITE.SGI_CODIGO  = PED.SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
       If Not IsNull(BREC!SGI_QTDE) Then PegaQtdTotItPedido = PegaQtdTotItPedido + BREC!SGI_QTDE
    End If
    BREC.Close

End Function


Public Function PegaQtdTotItPedido_Exc(lngCODPED As Long) As Long

    PegaQtdTotItPedido_Exc = 0
    
''    sSql = "Select " & vbCrLf
''    sSql = sSql & "       PED.* " & vbCrLf
''    sSql = sSql & "  From " & vbCrLf
''    sSql = sSql & "       SGI_CADPEDVENDI PED " & vbCrLf
''    sSql = sSql & " Where " & vbCrLf
''    sSql = sSql & "       PED.SGI_FILIAL = " & intFilial & vbCrLf
''    sSql = sSql & "   And PED.SGI_CODIGO = " & lngCODPEDIDO & vbCrLf
    
    sSql = "Select Sum(ITE.SGI_QTDE) As SGI_QTDE" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "      SGI_CADPEDVENDH PED" & vbCrLf
    sSql = sSql & "     ,SGI_CADPEDVENDI ITE" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "      PED.SGI_FILIAL = " & intFILIAL & vbCrLf
    sSql = sSql & "  And PED.SGI_CODIGO = " & lngCODPED & vbCrLf
    sSql = sSql & "  And ITE.SGI_FILIAL  = PED.SGI_FILIAL" & vbCrLf
    sSql = sSql & "  And ITE.SGI_CODIGO  = PED.SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then PegaQtdTotItPedido_Exc = PegaQtdTotItPedido_Exc + BREC!SGI_QTDE
    BREC.Close

End Function




Public Function PegaQtdTotItCota(lngCODCOTA As Long) As Long

    PegaQtdTotItCota = 0
    
    If BREC.State = 1 Then BREC.Close
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       Sum(SGI_QTDE) As SGI_QTDE " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADCOTAVENDI " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCODCOTA
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
       If Not IsNull(BREC!SGI_QTDE) Then PegaQtdTotItCota = BREC!SGI_QTDE
    End If
    BREC.Close

End Function

Public Function PreenchComboFechamentoGrd() As String
        
    PreenchComboFechamentoGrd = ""
    PreenchComboFechamentoGrd = PreenchComboFechamentoGrd & "|#" & 1 & ";VEX - Verniz Externo"
    PreenchComboFechamentoGrd = PreenchComboFechamentoGrd & "|#" & 2 & ";VZ - Verniz dos 2 Lados"
    PreenchComboFechamentoGrd = PreenchComboFechamentoGrd & "|#" & 3 & ";NAT - Natural"
    PreenchComboFechamentoGrd = PreenchComboFechamentoGrd & "|#" & 4 & ";VI - Verniz interno Apenas"

End Function


Public Function PreenchComboFechamentoGrdSA() As String
        
    PreenchComboFechamentoGrdSA = ""
    PreenchComboFechamentoGrdSA = PreenchComboFechamentoGrdSA & "|#" & 0 & ";SOLDA"
    PreenchComboFechamentoGrdSA = PreenchComboFechamentoGrdSA & "|#" & 1 & ";AGRAFADO"
    PreenchComboFechamentoGrdSA = PreenchComboFechamentoGrdSA & "|#" & 2 & ";REPUXO"

End Function


Public Function PreenchComboFechamentoTampaFuro(strIDPROD As String) As String

    If BREC10.State = 1 Then BREC10.Close
    
    PreenchComboFechamentoTampaFuro = ""
    
    If Len(Trim(strIDPROD)) = 0 Then Exit Function
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       LFTF.SGI_COD" & vbCrLf
    sSql = sSql & "      ,FECH.SGI_DESCRI" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO                PROD" & vbCrLf
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO           LIMP" & vbCrLf
    sSql = sSql & "      ,SGI_CADLINHAPRODUTO_FECHTPFR  LFTF" & vbCrLf
    sSql = sSql & "      ,SGI_CADFECHAM                 FECH" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       PROD.SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "   And PROD.SGI_IDPRODUTO  = " & strIDPROD & vbCrLf
    
    sSql = sSql & "   And PROD.SGI_FILIAL     = LIMP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PROD.SGI_CODLINPROD = LIMP.SGI_CODLIN" & vbCrLf
    
    sSql = sSql & "   And LIMP.SGI_FILIAL     = LFTF.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And LIMP.SGI_CODIGO     = LFTF.SGI_CODIGO" & vbCrLf
    
    sSql = sSql & "   And LFTF.SGI_FILIAL     = FECH.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And LFTF.SGI_COD        = FECH.SGI_CODIGO" & vbCrLf
    
    sSql = sSql & "Order By LFTF.SGI_COD"
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC10.EOF()
       PreenchComboFechamentoTampaFuro = PreenchComboFechamentoTampaFuro & "|#" & BREC10!SGI_COD & ";" & Trim(BREC10!SGI_DESCRI)
       BREC10.MoveNext
    Loop
    BREC10.Close
    
    ''PreenchComboFechamentoTampaFuro = PreenchComboFechamentoTampaFuro & "|#" & 1 & ";Ø24"
    ''PreenchComboFechamentoTampaFuro = PreenchComboFechamentoTampaFuro & "|#" & 2 & ";Ø25"
    ''PreenchComboFechamentoTampaFuro = PreenchComboFechamentoTampaFuro & "|#" & 3 & ";Ø42"
    ''PreenchComboFechamentoTampaFuro = PreenchComboFechamentoTampaFuro & "|#" & 4 & ";Ø45"
    ''PreenchComboFechamentoTampaFuro = PreenchComboFechamentoTampaFuro & "|#" & 5 & ";Ø57"
    ''PreenchComboFechamentoTampaFuro = PreenchComboFechamentoTampaFuro & "|#" & 6 & ";Ø80"
    ''PreenchComboFechamentoTampaFuro = PreenchComboFechamentoTampaFuro & "|#" & 7 & ";Ø130"
    ''PreenchComboFechamentoTampaFuro = PreenchComboFechamentoTampaFuro & "|#" & 8 & ";Ø170"
    ''PreenchComboFechamentoTampaFuro = PreenchComboFechamentoTampaFuro & "|#" & 9 & ";Ø110"
    ''PreenchComboFechamentoTampaFuro = PreenchComboFechamentoTampaFuro & "|#" & 10 & ";Ø170 c/b Ø25"
    ''PreenchComboFechamentoTampaFuro = PreenchComboFechamentoTampaFuro & "|#" & 11 & ";Ø170 c/v Ø57"
    ''PreenchComboFechamentoTampaFuro = PreenchComboFechamentoTampaFuro & "|#" & 12 & ";TP"
    ''PreenchComboFechamentoTampaFuro = PreenchComboFechamentoTampaFuro & "|#" & 13 & ";TP2"
    ''PreenchComboFechamentoTampaFuro = PreenchComboFechamentoTampaFuro & "|#" & 14 & ";TP4"
    ''PreenchComboFechamentoTampaFuro = PreenchComboFechamentoTampaFuro & "|#" & 15 & ";FA"
    ''PreenchComboFechamentoTampaFuro = PreenchComboFechamentoTampaFuro & "|#" & 16 & ";A RECRAVAR"
    ''PreenchComboFechamentoTampaFuro = PreenchComboFechamentoTampaFuro & "|#" & 17 & ";FA - C/Visor"
    ''PreenchComboFechamentoTampaFuro = PreenchComboFechamentoTampaFuro & "|#" & 18 & ";COFRE"
    ''PreenchComboFechamentoTampaFuro = PreenchComboFechamentoTampaFuro & "|#" & 19 & ";Porta Canetas"
    ''PreenchComboFechamentoTampaFuro = PreenchComboFechamentoTampaFuro & "|#" & 20 & ";Ø32 Bico Ret."
    ''PreenchComboFechamentoTampaFuro = PreenchComboFechamentoTampaFuro & "|#" & 21 & ";Repuxo"
    ''PreenchComboFechamentoTampaFuro = PreenchComboFechamentoTampaFuro & "|#" & 22 & ";FA c/b Ø25"
    ''PreenchComboFechamentoTampaFuro = PreenchComboFechamentoTampaFuro & "|#" & 23 & ";Ø24-Bico de Pato"
    ''PreenchComboFechamentoTampaFuro = PreenchComboFechamentoTampaFuro & "|#" & 24 & ";Ø24-REU"
    ''PreenchComboFechamentoTampaFuro = PreenchComboFechamentoTampaFuro & "|#" & 25 & ";F/T"
    
End Function

Public Sub GravaLogModulo(intFilial_Atu As Integer, lngCodLog As Long, strModulo As String, strACAO As String, lngCodVendedor As Long, strCodReg As String)

On Error GoTo err_Trans
    
    '' Inicia transação
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

    Dim objErro    As Object
    Set objErro = CreateObject("BLBCWS.clsFuncoes")
    Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, strACAO, sSql)
    Set objErro = Nothing

End Sub


Public Function GRAVASTEEL(strOPER As String) As Boolean

On Error GoTo err_Trans
     
     GRAVASTEEL = False
         
     Dim i              As Integer
     Dim j              As Integer
     Dim sValor         As String
     Dim strTIPODESP    As String
     Dim lngQtdPed      As Long
     Dim strNOTABFILIAL As String
     
     
     strNOTABFILIAL = ""
     If intFILIALPED = 1 Then strNOTABFILIAL = "_STEEL"
     
     '' Inicia transação
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
      
     If strOPER = "I" Then '' Inclui
        
        '' Grava Cotação HEADER(CABEÇALHO)
        sSql = "Insert Into SGI_CADPEDVENDH" & strNOTABFILIAL & "  (SGI_FILIAL" & vbCrLf
        sSql = sSql & "                      ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                      ,SGI_DATAPED" & vbCrLf
        sSql = sSql & "                      ,SGI_DTCALIBRA" & vbCrLf
        sSql = sSql & "                      ,SGI_CODCLI" & vbCrLf
        sSql = sSql & "                      ,SGI_CODVEND" & vbCrLf
        sSql = sSql & "                      ,SGI_CODCONDPGT" & vbCrLf
        sSql = sSql & "                      ,SGI_CODTIPORC" & vbCrLf
        sSql = sSql & "                      ,SGI_PRZENTR" & vbCrLf
        sSql = sSql & "                      ,SGI_VALPROP" & vbCrLf
        sSql = sSql & "                      ,SGI_IPIISENTO" & vbCrLf
        sSql = sSql & "                      ,SGI_ICMSISENTO" & vbCrLf
        sSql = sSql & "                      ,SGI_TIPFRETE" & vbCrLf
        sSql = sSql & "                      ,SGI_QTDITENS" & vbCrLf
        sSql = sSql & "                      ,SGI_VLTOT" & vbCrLf
        sSql = sSql & "                      ,SGI_OUTRASTOT" & vbCrLf
        sSql = sSql & "                      ,SGI_FRETETOT" & vbCrLf
        sSql = sSql & "                      ,SGI_BASEICMS" & vbCrLf
        sSql = sSql & "                      ,SGI_PORCICMS" & vbCrLf
        sSql = sSql & "                      ,SGI_VLICMS" & vbCrLf
        sSql = sSql & "                      ,SGI_VLIPI" & vbCrLf
        sSql = sSql & "                      ,SGI_VLDESCTO" & vbCrLf
        sSql = sSql & "                      ,SGI_STATUS" & vbCrLf
        sSql = sSql & "                      ,SGI_ENDENTRE" & vbCrLf
        sSql = sSql & "                      ,SGI_BAIENTRE" & vbCrLf
        sSql = sSql & "                      ,SGI_CIDENTRE" & vbCrLf
        sSql = sSql & "                      ,SGI_ESTENTRE" & vbCrLf
        sSql = sSql & "                      ,SGI_CEPENTRE" & vbCrLf
        sSql = sSql & "                      ,SGI_ENDCOBR" & vbCrLf
        sSql = sSql & "                      ,SGI_BAICOBR" & vbCrLf
        sSql = sSql & "                      ,SGI_CIDCOBR" & vbCrLf
        sSql = sSql & "                      ,SGI_ESTCOBR" & vbCrLf
        sSql = sSql & "                      ,SGI_CEPCOBR" & vbCrLf
        sSql = sSql & "                      ,SGI_CODCOTA" & vbCrLf
        sSql = sSql & "                      ,SGI_DTENTREGA" & vbCrLf
        sSql = sSql & "                      ,SGI_TELENTR" & vbCrLf
        sSql = sSql & "                      ,SGI_FAXENTR" & vbCrLf
        sSql = sSql & "                      ,SGI_TELCOBR" & vbCrLf
        sSql = sSql & "                      ,SGI_FAXCOBR" & vbCrLf
        sSql = sSql & "                      ,SGI_CONTATO" & vbCrLf
        sSql = sSql & "                      ,SGI_DEPTO" & vbCrLf
        sSql = sSql & "                      ,SGI_ORDCOCLI" & vbCrLf
        sSql = sSql & "                      ,SGI_CODTRANSP" & vbCrLf
        sSql = sSql & "                      ,SGI_EMAIL" & vbCrLf
        sSql = sSql & "                      ,SGI_CODTEC" & vbCrLf
        sSql = sSql & "                      ,SGI_OBS" & vbCrLf
        sSql = sSql & "                      ,SGI_CODTIPODOC" & vbCrLf
        sSql = sSql & "                      ,SGI_CODBANCO" & vbCrLf
        sSql = sSql & "                      ,SGI_PORCDESCPED" & vbCrLf
        sSql = sSql & "                      ,SGI_VALDESCPED" & vbCrLf
        sSql = sSql & "                      ,SGI_QTDEITENSPEDIDO" & vbCrLf
        sSql = sSql & "                      ,SGI_OBSQUALID" & vbCrLf
        sSql = sSql & "                      ,SGI_FILIALPED" & vbCrLf
        sSql = sSql & "                      ,SGI_ESPECIAL" & vbCrLf
        sSql = sSql & "                      ,SGI_PARAESTOQUE" & vbCrLf
        sSql = sSql & "                      ,SGI_CONFERIDO" & vbCrLf
        sSql = sSql & "                      ,SGI_VERSAO" & vbCrLf
        sSql = sSql & "                      ,SGI_NOMCOMP" & vbCrLf
        sSql = sSql & "                      ,SGI_PERMFECHOP" & vbCrLf
        
        sSql = sSql & "                      )" & vbCrLf
        
        sSql = sSql & "               Values (" & vbCrLf
        sSql = sSql & "                       " & intFILIAL & vbCrLf
        sSql = sSql & "                      ," & lngCODPEDIDO & vbCrLf
        sSql = sSql & "                      ,'" & Format(dtDATAPED, "MM/DD/YYYY") & "'" & vbCrLf
        If IsDate(dtDATCALIBR) Then
           sSql = sSql & "                      ,'" & Format(dtDATCALIBR, "MM/DD/YYYY") & "'" & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
        sSql = sSql & "                      ," & intCODCLIE & vbCrLf
        sSql = sSql & "                      ," & intCODVEND & vbCrLf
        sSql = sSql & "                      ," & intCODCONDPGTO & vbCrLf
        sSql = sSql & "                      ," & intTIPPED & vbCrLf
        sSql = sSql & "                      ," & intPRZENTREGA & vbCrLf
        sSql = sSql & "                      ," & intVALPROPOSTA & vbCrLf
        sSql = sSql & "                      ," & intIPIISENTO & vbCrLf
        sSql = sSql & "                      ," & intICMSISENTO & vbCrLf
        sSql = sSql & "                      ," & intFRETE & vbCrLf
        sSql = sSql & "                      ," & intTOTALITENS & vbCrLf
        
        sValor = Replace(curTOTORCTO, ".", "")
        sValor = Replace(sValor, ",", ".")
        sSql = sSql & "                      ," & sValor & vbCrLf
        
        If curOUTRDESPESAS > 0 Then
           sValor = Replace(curOUTRDESPESAS, ".", "")
           sValor = Replace(sValor, ",", ".")
           sSql = sSql & "                      ," & sValor & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
        
        If curVLFRETE > 0 Then
           sValor = Replace(curVLFRETE, ".", "")
           sValor = Replace(sValor, ",", ".")
           sSql = sSql & "                      ," & sValor & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
        
        If curVALBASICMS > 0 Then
           sValor = Replace(curVALBASICMS, ".", "")
           sValor = Replace(sValor, ",", ".")
           sSql = sSql & "                      ," & sValor & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
        
        If curAliqICMS > 0 Then
           sValor = Replace(curAliqICMS, ".", "")
           sValor = Replace(sValor, ",", ".")
           sSql = sSql & "                      ," & sValor & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
        
        If curVLICMS > 0 Then
           sValor = Replace(curVLICMS, ".", "")
           sValor = Replace(sValor, ",", ".")
           sSql = sSql & "                      ," & sValor & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If

        If curVLIPI > 0 Then
           sValor = Replace(curVLIPI, ".", "")
           sValor = Replace(sValor, ",", ".")
           sSql = sSql & "                      ," & sValor & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If

        If curVLDESCTO > 0 Then
           sValor = Replace(curVLDESCTO, ".", "")
           sValor = Replace(sValor, ",", ".")
           sSql = sSql & "                      ," & sValor & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
        
        sSql = sSql & "                      ,'" & strSTATUS & "'" & vbCrLf
        sSql = sSql & "                      ,'" & strENDENTR & "'" & vbCrLf
        sSql = sSql & "                      ,'" & strBAIENTR & "'" & vbCrLf
        sSql = sSql & "                      ,'" & strCIDENTR & "'" & vbCrLf
        sSql = sSql & "                      ," & intESTENTREGA & vbCrLf
        sSql = sSql & "                      ,'" & strCEPENTREGA & "'" & vbCrLf
        sSql = sSql & "                      ,'" & strENDCOBRA & "'" & vbCrLf
        sSql = sSql & "                      ,'" & strBAICOBRA & "'" & vbCrLf
        sSql = sSql & "                      ,'" & strCIDCOBRA & "'" & vbCrLf
        sSql = sSql & "                      ," & intESTCOBRA & vbCrLf
        sSql = sSql & "                      ,'" & strCEPCOBRA & "'" & vbCrLf
        
        If lngCODCOTA > 0 Then
           sSql = sSql & "                      ," & lngCODCOTA & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
        
        sSql = sSql & "                      ,'" & Format(dtDATAENTREGA, "MM/DD/YYYY") & "'" & vbCrLf
        
        sSql = sSql & "                      ,'" & strTELENTR & "'" & vbCrLf
        sSql = sSql & "                      ,'" & strFAXENTR & "'" & vbCrLf
        sSql = sSql & "                      ,'" & strTELCOBRA & "'" & vbCrLf
        sSql = sSql & "                      ,'" & strFAXCOBRA & "'" & vbCrLf
        sSql = sSql & "                      ,'" & strCONTATO & "'" & vbCrLf
        sSql = sSql & "                      ,'" & strDEPARTAMENTO & "'" & vbCrLf
        sSql = sSql & "                      ,'" & strORDCOMPCLI & "'" & vbCrLf
        sSql = sSql & "                      ," & lngCODTRANSP & vbCrLf
        sSql = sSql & "                      ,'" & strEMAIL & "'" & vbCrLf
        If intCODTECNICO > 0 Then
           sSql = sSql & "                      ," & intCODTECNICO & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
        sSql = sSql & "                      ,'" & strOBSERVACAO & "'" & vbCrLf
        
        If lngCODTIPODOC > 0 Then
           sSql = sSql & "                      ," & lngCODTIPODOC & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
        If lngCODBANCO > 0 Then
           sSql = sSql & "                      ," & lngCODBANCO & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
        
        If curPORDESC > 0 Then
           sValor = Replace(curPORDESC, ".", "")
           sValor = Replace(sValor, ",", ".")
           sSql = sSql & "                      ," & sValor & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
        
        If curVALDESC > 0 Then
           sValor = Replace(curVALDESC, ".", "")
           sValor = Replace(sValor, ",", ".")
           sSql = sSql & "                      ," & sValor & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
        
        sValor = "Null"
        If curQTDITENSPROD > 0 Then
           sValor = Replace(curQTDITENSPROD, ".", "")
           sValor = Replace(sValor, ",", ".")
        End If
        sSql = sSql & "                      ," & curQTDITENSPROD & vbCrLf
        sSql = sSql & "                      ," & Trim(strOBS2) & vbCrLf
        sSql = sSql & "                      ," & intFILIALPED & vbCrLf
        sSql = sSql & "                      ," & intESPECIAL & vbCrLf
        sSql = sSql & "                      ," & intPARAESTOQUE & vbCrLf
        sSql = sSql & "                      ," & intCONFERIDO & vbCrLf
        sSql = sSql & "                      ," & strVERSAO & vbCrLf
        sSql = sSql & "                      ," & strNOMCOMP & vbCrLf
        sSql = sSql & "                      ," & intPERMITEFECHOP & vbCrLf
        
        sSql = sSql & "                           )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        
        '' Grava os Produtos
        lngQtdPed = 0
        If IsArray(arrPRODUTOS) = True Then
           For i = 1 To UBound(arrPRODUTOS)
               
                     sSql = ""
                     
                     sSql = "Insert Into SGI_CADPEDVENDI" & strNOTABFILIAL & "  (SGI_FILIAL" & vbCrLf
                     sSql = sSql & "                      ,SGI_CODIGO" & vbCrLf
                     sSql = sSql & "                      ,SGI_IDPRODUTO" & vbCrLf
                     sSql = sSql & "                      ,SGI_CODPROD" & vbCrLf
                     sSql = sSql & "                      ,SGI_QTDE" & vbCrLf
                     sSql = sSql & "                      ,SGI_VLUNIT" & vbCrLf
                     sSql = sSql & "                      ,SGI_PRCDESC" & vbCrLf
                     sSql = sSql & "                      ,SGI_PRCIPI" & vbCrLf
                     sSql = sSql & "                      ,SGI_VLTOT" & vbCrLf
                     sSql = sSql & "                      ,SGI_VLIPI" & vbCrLf
                     sSql = sSql & "                      ,SGI_VLDESC" & vbCrLf
                     sSql = sSql & "                      ,SGI_ALTFILM" & vbCrLf
                     sSql = sSql & "                      ,SGI_FOTNOVO" & vbCrLf
                     sSql = sSql & "                      ,SGI_REPETICAO" & vbCrLf
                     sSql = sSql & "                      ,SGI_DATAPED" & vbCrLf
                     sSql = sSql & "                      ,SGI_CODLINPROD" & vbCrLf
                     sSql = sSql & "                      ,SGI_CODFORN" & vbCrLf
                     sSql = sSql & "                      ,SGI_OBSOP" & vbCrLf
                     sSql = sSql & "                      ,SGI_FECHTPFU" & vbCrLf
                     sSql = sSql & "                      ,SGI_FILIALPED" & vbCrLf
                     sSql = sSql & "                      ,SGI_QTDEPORPALHETS" & vbCrLf
                     sSql = sSql & "                      ,SGI_PALHETS" & vbCrLf
                     sSql = sSql & "                      ,SGI_CONFERIDO" & vbCrLf
                     sSql = sSql & "                      ,SGI_USARPADRPALSN" & vbCrLf
                     sSql = sSql & "                      ,SGI_DESCPROD" & vbCrLf
                     
                     sSql = sSql & "                      ,SGI_FechSoldaAgrafado" & vbCrLf
                     sSql = sSql & "                      ,SGI_VernCorpo" & vbCrLf
                     sSql = sSql & "                      ,SGI_VernTampa" & vbCrLf
                     sSql = sSql & "                      ,SGI_VernFundo" & vbCrLf
                     sSql = sSql & "                      ,SGI_VernArgola" & vbCrLf
                     
                     sSql = sSql & "              ) Values (" & vbCrLf
                     
                     sSql = sSql & "                       " & intFILIAL & vbCrLf
                     sSql = sSql & "                      ," & lngCODPEDIDO & vbCrLf
                     sSql = sSql & "                      ," & Trim(arrPRODUTOS(i, 9)) & vbCrLf
                     sSql = sSql & "                      ,'" & Trim(arrPRODUTOS(i, 1)) & "'" & vbCrLf
                     sSql = sSql & "                      ," & arrPRODUTOS(i, 2) & vbCrLf
                     lngQtdPed = lngQtdPed + CLng(arrPRODUTOS(i, 2))
                     
                     sSql = sSql & "                      ," & arrPRODUTOS(i, 3) & vbCrLf
                     sSql = sSql & "                      ," & arrPRODUTOS(i, 4) & vbCrLf
                     sSql = sSql & "                      ," & arrPRODUTOS(i, 5) & vbCrLf
                     sSql = sSql & "                      ," & arrPRODUTOS(i, 6) & vbCrLf
                     sSql = sSql & "                      ," & arrPRODUTOS(i, 7) & vbCrLf
                     sSql = sSql & "                      ," & arrPRODUTOS(i, 8) & vbCrLf
                     
                     sSql = sSql & "                      ," & Trim(arrPRODUTOS(i, 10)) & vbCrLf
                     sSql = sSql & "                      ," & Trim(arrPRODUTOS(i, 11)) & vbCrLf
                     sSql = sSql & "                      ," & Trim(arrPRODUTOS(i, 12)) & vbCrLf
                     sSql = sSql & "                      ,'" & Format(dtDATAPED, "MM/DD/YYYY") & "'" & vbCrLf
                     sSql = sSql & "                      ," & arrPRODUTOS(i, 13) & vbCrLf
                     sSql = sSql & "                      ," & arrPRODUTOS(i, 14) & vbCrLf
                     sSql = sSql & "                      ,'" & Trim(arrPRODUTOS(i, 15)) & "'" & vbCrLf
                     sSql = sSql & "                      ," & Trim(arrPRODUTOS(i, 16)) & vbCrLf
                     sSql = sSql & "                      ," & intFILIALPED & vbCrLf
                     sSql = sSql & "                      ," & arrPRODUTOS(i, 19) & vbCrLf
                     sSql = sSql & "                      ," & arrPRODUTOS(i, 20) & vbCrLf
                     sSql = sSql & "                      ," & arrPRODUTOS(i, 21) & vbCrLf
                     sSql = sSql & "                      ," & arrPRODUTOS(i, 22) & vbCrLf
                     sSql = sSql & "                      ," & arrPRODUTOS(i, 23) & vbCrLf
                     
                     sSql = sSql & "                      ," & arrPRODUTOS(i, 24) & vbCrLf
                     sSql = sSql & "                      ," & arrPRODUTOS(i, 25) & vbCrLf
                     sSql = sSql & "                      ," & arrPRODUTOS(i, 26) & vbCrLf
                     sSql = sSql & "                      ," & arrPRODUTOS(i, 27) & vbCrLf
                     sSql = sSql & "                      ," & arrPRODUTOS(i, 28) & vbCrLf
                     
                     sSql = sSql & "                           )"
                
                     BGRV.CommandText = sSql
                     BGRV.Execute
                     
                     
                     If (arrPRODUTOS(i, 10) = 1 Or arrPRODUTOS(i, 11) = 1) Then
                           
                        '' arrPRODUTOS(I, 10) Alteração do Fotolito (1 = sim / 0 = não)
                        '' arrPRODUTOS(I, 11) Fotolito Novo  (1 = sim / 0 = não)
                           
                        sSql = "Update SGI_CADPRODUTO Set" & vbCrLf
                        
                        sSql = sSql & "                      SGI_STATUS     = 2" & vbCrLf
                        sSql = sSql & "                     ,SGI_ALTFOTOLIT = " & Trim(arrPRODUTOS(i, 10)) & vbCrLf
                        sSql = sSql & "                     ,SGI_PRODNOVO   = " & Trim(arrPRODUTOS(i, 11)) & vbCrLf
                        
                        sSql = sSql & "       Where" & vbCrLf
                        sSql = sSql & "             SGI_FILIAL    = " & intFILIAL & vbCrLf
                        sSql = sSql & "         And SGI_IDPRODUTO = " & Trim(arrPRODUTOS(i, 9))
                        
                        BGRV.CommandText = sSql
                        BGRV.Execute
                        
                     End If
           
           Next i
        End If
        
        
        '' =====================================
        '' Programação de Entregas
        If IsArray(arrPROGENTREGAS) Then
           For i = 1 To UBound(arrPROGENTREGAS)
                   
                   sSql = "Insert Into SGI_PROGENTRPROD" & strNOTABFILIAL & " (" & vbCrLf
                   sSql = sSql & "                       SGI_FILIAL" & vbCrLf
                   sSql = sSql & "                      ,SGI_CODPED" & vbCrLf
                   sSql = sSql & "                      ,SGI_IDPRODUTO" & vbCrLf
                   sSql = sSql & "                      ,SGI_INDICE" & vbCrLf
                   sSql = sSql & "                      ,SGI_QTDE" & vbCrLf
                   sSql = sSql & "                      ,SGI_DATENTREGA" & vbCrLf
                   sSql = sSql & "                      ,SGI_OBSOP" & vbCrLf
                   sSql = sSql & "                      ,SGI_FECHTPFU" & vbCrLf
                   sSql = sSql & "                      ,SGI_FILIALPED" & vbCrLf
                   sSql = sSql & "                      ,SGI_IDINTERNO" & vbCrLf
                   sSql = sSql & "                      ,SGI_STATUS" & vbCrLf
                   sSql = sSql & "                      ,SGI_QTDEPORPALHETS" & vbCrLf
                   sSql = sSql & "                      ,SGI_PALHETS" & vbCrLf
                   
                   sSql = sSql & "                      ) Values ( " & vbCrLf
                   sSql = sSql & "                               " & intFILIAL & vbCrLf
                   sSql = sSql & "                              ," & lngCODPEDIDO & vbCrLf
                   sSql = sSql & "                              ," & arrPROGENTREGAS(i, 0) & vbCrLf
                   sSql = sSql & "                              ," & i & vbCrLf
                   
                   sValor = Replace(arrPROGENTREGAS(i, 1), ".", "")
                   sValor = Replace(sValor, ",", ".")
                   sSql = sSql & "                              ," & sValor & vbCrLf
                   
                   sSql = sSql & "                              ," & arrPROGENTREGAS(i, 2) & vbCrLf
                   sSql = sSql & "                              ,'" & Trim(arrPROGENTREGAS(i, 4)) & "'" & vbCrLf
                   sSql = sSql & "                              ," & Trim(arrPROGENTREGAS(i, 5)) & vbCrLf
                   sSql = sSql & "                              ," & intFILIALPED & vbCrLf
                   sSql = sSql & "                              ," & arrPROGENTREGAS(i, 16) & vbCrLf
                   sSql = sSql & "                              ," & arrPROGENTREGAS(i, 17) & vbCrLf '' Status
                   sSql = sSql & "                              ," & arrPROGENTREGAS(i, 18) & vbCrLf
                   sSql = sSql & "                              ," & arrPROGENTREGAS(i, 19) & vbCrLf
                   
                   sSql = sSql & "                              )"
                   
                   BGRV.CommandText = sSql
                   BGRV.Execute
                   
           Next i
        End If
        '' =====================================
        
        
        '' =====================================
        '' Grava Log
        sSql = ""
        
        sSql = "Insert Into SGI_LOGPED" & strNOTABFILIAL & " (" & vbCrLf
        sSql = sSql & "                              SGI_FILIAL" & vbCrLf
        sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                             ,SGI_DATA" & vbCrLf
        sSql = sSql & "                             ,SGI_HORA" & vbCrLf
        sSql = sSql & "                             ,SGI_CODUSUARIO" & vbCrLf
        sSql = sSql & "                             ,SGI_ACAO" & vbCrLf
        sSql = sSql & "                    ) Values (" & vbCrLf
        sSql = sSql & "                              " & FILIAL & vbCrLf
        sSql = sSql & "                             ," & lngCODPEDIDO & vbCrLf
        sSql = sSql & "                             ,'" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                             ,'" & Format(Now, "HH:MM:SS") & "'" & vbCrLf
        sSql = sSql & "                             ," & lngCodUsuario & vbCrLf
        sSql = sSql & "                             ,'" & strOPER & "'" & vbCrLf
        sSql = sSql & "                             )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' =====================================
        
        
     
     ElseIf strOPER = "A" Then '' Altera
     
         '' Apagando Tipo de Reprovação
         sSql = "Delete from SGI_TIPREPRPED" & strNOTABFILIAL & vbCrLf
         sSql = sSql & "      Where " & vbCrLf
         sSql = sSql & "           SGI_FILIAL = " & intFILIAL & vbCrLf
         sSql = sSql & "       And SGI_CODPED = " & lngCODPEDIDO
         
         BGRV.CommandText = sSql
         BGRV.Execute
         
         '' Pedidos Steel
         sSql = "Update SGI_CADPEDVENDH" & strNOTABFILIAL & " Set" & vbCrLf
         
         sSql = sSql & "                           SGI_DATAPED    = '" & Format(dtDATAPED, "MM/DD/YYYY") & "'" & vbCrLf
         sSql = sSql & "                          ,SGI_CODVEND    = " & intCODVEND & vbCrLf
         sSql = sSql & "                          ,SGI_CODCONDPGT = " & intCODCONDPGTO & vbCrLf
         sSql = sSql & "                          ,SGI_CODTIPORC  = " & intTIPPED & vbCrLf
         sSql = sSql & "                          ,SGI_CODTRANSP  = " & lngCODTRANSP & vbCrLf
         
         sValor = "Null"
         If curVALBASICMS > 0 Then
            sValor = Replace(curVALBASICMS, ".", "")
            sValor = Replace(sValor, ",", ".")
         End If
         sSql = sSql & "                      ,SGI_BASEICMS = " & sValor & vbCrLf
         
         sValor = "Null"
         If curAliqICMS > 0 Then
            sValor = Replace(curAliqICMS, ".", "")
            sValor = Replace(sValor, ",", ".")
         End If
         sSql = sSql & "                      ,SGI_PORCICMS = " & sValor & vbCrLf
         
         sValor = "Null"
         If curVLICMS > 0 Then
            sValor = Replace(curVLICMS, ".", "")
            sValor = Replace(sValor, ",", ".")
         End If
         sSql = sSql & "                      ,SGI_VLICMS = " & sValor & vbCrLf
         
         sValor = "Null"
         If curVLIPI > 0 Then
            sValor = Replace(curVLIPI, ".", "")
            sValor = Replace(sValor, ",", ".")
         End If
         sSql = sSql & "                      ,SGI_VLIPI = " & sValor & vbCrLf
         
         sValor = "Null"
         If curOUTRDESPESAS > 0 Then
            sValor = Replace(curOUTRDESPESAS, ".", "")
            sValor = Replace(sValor, ",", ".")
         End If
         sSql = sSql & "                      ,SGI_OUTRASTOT = " & sValor & vbCrLf
         
         sValor = "Null"
         If curVLFRETE > 0 Then
            sValor = Replace(curVLFRETE, ".", "")
            sValor = Replace(sValor, ",", ".")
         End If
         sSql = sSql & "                      ,SGI_FRETETOT = " & sValor & vbCrLf
         
         sValor = "Null"
         If curVLDESCTO > 0 Then
            sValor = Replace(curVLDESCTO, ".", "")
            sValor = Replace(sValor, ",", ".")
         End If
         sSql = sSql & "                      ,SGI_VLDESCTO = " & sValor & vbCrLf
         
         sValor = "Null"
         If curTOTORCTO > 0 Then
            sValor = Replace(curTOTORCTO, ".", "")
            sValor = Replace(sValor, ",", ".")
         End If
         sSql = sSql & "                      ,SGI_VLTOT = " & sValor & vbCrLf
         
         sSql = sSql & "                      ,SGI_QTDITENS     = " & intTOTALITENS & vbCrLf
         sSql = sSql & "                      ,SGI_ESPECIAL     = " & intESPECIAL & vbCrLf
         sSql = sSql & "                      ,SGI_PARAESTOQUE  = " & intPARAESTOQUE & vbCrLf
         
         sSql = sSql & "                      ,SGI_ENDENTRE     = '" & strENDENTR & "'" & vbCrLf
         sSql = sSql & "                      ,SGI_BAIENTRE     = '" & strBAIENTR & "'" & vbCrLf
         sSql = sSql & "                      ,SGI_CIDENTRE     = '" & strCIDENTR & "'" & vbCrLf
         sSql = sSql & "                      ,SGI_CEPENTRE     = '" & strCEPENTREGA & "'" & vbCrLf
         sSql = sSql & "                      ,SGI_ESTENTRE     = " & intESTENTREGA & vbCrLf
         sSql = sSql & "                      ,SGI_TELENTR      = '" & strTELENTR & "'" & vbCrLf
         sSql = sSql & "                      ,SGI_FAXENTR      = '" & strFAXENTR & "'" & vbCrLf
    
         sSql = sSql & "                      ,SGI_ENDCOBR      = '" & strENDCOBRA & "'" & vbCrLf
         sSql = sSql & "                      ,SGI_BAICOBR      = '" & strBAICOBRA & "'" & vbCrLf
         sSql = sSql & "                      ,SGI_CIDCOBR      = '" & strCIDCOBRA & "'" & vbCrLf
         sSql = sSql & "                      ,SGI_CEPCOBR      = '" & strCEPCOBRA & "'" & vbCrLf
         sSql = sSql & "                      ,SGI_ESTCOBR      = " & intESTCOBRA & vbCrLf
         sSql = sSql & "                      ,SGI_TELCOBR      = '" & strTELCOBRA & "'" & vbCrLf
         sSql = sSql & "                      ,SGI_FAXCOBR      = '" & strFAXCOBRA & "'" & vbCrLf
         
         sSql = sSql & "                      ,SGI_CONTATO      = '" & strCONTATO & "'" & vbCrLf
         sSql = sSql & "                      ,SGI_DEPTO        = '" & strDEPARTAMENTO & "'" & vbCrLf
         sSql = sSql & "                      ,SGI_ORDCOCLI     = '" & strORDCOMPCLI & "'" & vbCrLf
         sSql = sSql & "                      ,SGI_EMAIL        = '" & strEMAIL & "'" & vbCrLf
         sSql = sSql & "                      ,SGI_OBS          = '" & strOBSERVACAO & "'" & vbCrLf
         sSql = sSql & "                      ,SGI_OBSQUALID    = " & strOBS2 & vbCrLf
         sSql = sSql & "                      ,SGI_STATUS       = '" & strSTATUS & "'" & vbCrLf '' Status
         sSql = sSql & "                      ,SGI_CONFERIDO    = " & intCONFERIDO & vbCrLf
         sSql = sSql & "                      ,SGI_NOMCOMP      = '" & strNOMCOMP & "'" & vbCrLf
         sSql = sSql & "                      ,SGI_PERMFECHOP   = " & intPERMITEFECHOP & vbCrLf
        
         sSql = sSql & " Where " & vbCrLf
         sSql = sSql & "       SGI_CODIGO = " & lngCODPEDIDO
         
         BGRV.CommandText = sSql
         BGRV.Execute
     
         
        '' =====================================
        '' Programação de Entregas
        If IsArray(arrPROGENTREGAS) Then
           For i = 1 To UBound(arrPROGENTREGAS)
                If arrPROGENTREGAS(i, 3) = dacEnumUpdateAction_delete Then
                
                    If Len(Trim(arrPROGENTREGAS(i, 9))) > 0 Then
                    
                        sSql = ""
                        
                        sSql = "Delete SGI_ORDEMPROD" & strNOTABFILIAL & vbCrLf
                        sSql = sSql & "                         Where " & vbCrLf
                        sSql = sSql & "                               SGI_FILIAL     = " & intFILIAL & vbCrLf
                        sSql = sSql & "                           And SGI_CODIGO     = " & Trim(arrPROGENTREGAS(i, 9)) & vbCrLf
                        sSql = sSql & "                           And SGI_IDPRODUTO  = " & arrPROGENTREGAS(i, 0) & vbCrLf
                        sSql = sSql & "                           And SGI_IDPAI      = " & arrPROGENTREGAS(i, 16) & vbCrLf
                    
                        BGRV.CommandText = sSql
                        BGRV.Execute
                    
                    End If
                    
                    
                    sSql = ""
                    
                    sSql = "Delete SGI_PROGENTRPROD" & strNOTABFILIAL & vbCrLf
                    sSql = sSql & "                           Where " & vbCrLf
                    sSql = sSql & "                                 SGI_FILIAL    = " & intFILIAL & vbCrLf
                    sSql = sSql & "                             And SGI_CODPED    = " & lngCODPEDIDO & vbCrLf
                    sSql = sSql & "                             And SGI_IDPRODUTO = " & arrPROGENTREGAS(i, 0) & vbCrLf
                    sSql = sSql & "                             And SGI_IDINTERNO = " & arrPROGENTREGAS(i, 16) & vbCrLf
                    
                    ''sSql = sSql & "                             And SGI_INDICE    = " & arrPROGENTREGAS(I, 8)
                
                    BGRV.CommandText = sSql
                    BGRV.Execute
                
                
                End If
           Next i
           
           For i = 1 To UBound(arrPROGENTREGAS)
                If arrPROGENTREGAS(i, 3) = dacEnumUpdateAction_Insert Then
                
                   sSql = "Insert Into SGI_PROGENTRPROD" & strNOTABFILIAL & " (" & vbCrLf
                   sSql = sSql & "                       SGI_FILIAL" & vbCrLf
                   sSql = sSql & "                      ,SGI_CODPED" & vbCrLf
                   sSql = sSql & "                      ,SGI_IDPRODUTO" & vbCrLf
                   sSql = sSql & "                      ,SGI_INDICE" & vbCrLf
                   sSql = sSql & "                      ,SGI_QTDE" & vbCrLf
                   sSql = sSql & "                      ,SGI_DATENTREGA" & vbCrLf
                   sSql = sSql & "                      ,SGI_OBSOP" & vbCrLf
                   sSql = sSql & "                      ,SGI_FECHTPFU" & vbCrLf
                   sSql = sSql & "                      ,SGI_FILIALPED" & vbCrLf
                   sSql = sSql & "                      ,SGI_IDINTERNO" & vbCrLf
                   sSql = sSql & "                      ,SGI_STATUS" & vbCrLf
                   sSql = sSql & "                      ,SGI_QTDEPORPALHETS" & vbCrLf
                   sSql = sSql & "                      ,SGI_PALHETS" & vbCrLf
                   
                   sSql = sSql & "                      ) Values ( " & vbCrLf
                   sSql = sSql & "                               " & intFILIAL & vbCrLf
                   sSql = sSql & "                              ," & lngCODPEDIDO & vbCrLf
                   sSql = sSql & "                              ," & arrPROGENTREGAS(i, 0) & vbCrLf
                   sSql = sSql & "                              ," & i & vbCrLf
                   
                   sValor = Replace(arrPROGENTREGAS(i, 1), ".", "")
                   sValor = Replace(sValor, ",", ".")
                   sSql = sSql & "                              ," & sValor & vbCrLf
                   
                   sSql = sSql & "                              ," & arrPROGENTREGAS(i, 2) & vbCrLf
                   sSql = sSql & "                              ,'" & Trim(arrPROGENTREGAS(i, 4)) & "'" & vbCrLf
                   sSql = sSql & "                              ," & Trim(arrPROGENTREGAS(i, 5)) & vbCrLf
                   sSql = sSql & "                              ," & intFILIALPED & vbCrLf
                   sSql = sSql & "                              ," & arrPROGENTREGAS(i, 16) & vbCrLf
                   sSql = sSql & "                              ," & arrPROGENTREGAS(i, 17) & vbCrLf
                   sSql = sSql & "                              ," & arrPROGENTREGAS(i, 18) & vbCrLf
                   sSql = sSql & "                              ," & arrPROGENTREGAS(i, 19) & vbCrLf
                   
                   sSql = sSql & "                              )"
                   
                   BGRV.CommandText = sSql
                   BGRV.Execute
                
                ElseIf arrPROGENTREGAS(i, 3) = dacEnumUpdateAction_update Then
                
                    sSql = ""
                    
                    sSql = "Update SGI_PROGENTRPROD" & strNOTABFILIAL & " Set " & vbCrLf
                    sSql = sSql & "                                   SGI_QTDE              = " & arrPROGENTREGAS(i, 1) & vbCrLf
                    sSql = sSql & "                                  ,SGI_INDICE            = " & arrPROGENTREGAS(i, 7) & vbCrLf
                    sSql = sSql & "                                  ,SGI_DATENTREGA        = " & arrPROGENTREGAS(i, 2) & vbCrLf
                    sSql = sSql & "                                  ,SGI_OBSOP             = '" & arrPROGENTREGAS(i, 4) & "'" & vbCrLf
                    sSql = sSql & "                                  ,SGI_FECHTPFU          = " & arrPROGENTREGAS(i, 5) & vbCrLf
                    sSql = sSql & "                                  ,SGI_STATUS            = " & arrPROGENTREGAS(i, 17) & vbCrLf '' Status da OP
                    sSql = sSql & "                                  ,SGI_QTDEPORPALHETS    = " & arrPROGENTREGAS(i, 18) & vbCrLf
                    sSql = sSql & "                                  ,SGI_PALHETS           = " & arrPROGENTREGAS(i, 19) & vbCrLf
                    
                    sSql = sSql & "                           Where " & vbCrLf
                    sSql = sSql & "                                 SGI_FILIAL    = " & intFILIAL & vbCrLf
                    sSql = sSql & "                             And SGI_CODPED    = " & lngCODPEDIDO & vbCrLf
                    sSql = sSql & "                             And SGI_IDPRODUTO = " & arrPROGENTREGAS(i, 0) & vbCrLf
                    sSql = sSql & "                             And SGI_IDINTERNO = " & arrPROGENTREGAS(i, 16) & vbCrLf
                
                    BGRV.CommandText = sSql
                    BGRV.Execute
                
                    If Len(Trim(arrPROGENTREGAS(i, 9))) > 0 Then
                    
                        sSql = ""
                        
                        sSql = "Update SGI_ORDEMPROD" & strNOTABFILIAL & " Set " & vbCrLf
                        sSql = sSql & "                               SGI_QTDEPED               = " & arrPROGENTREGAS(i, 1) & vbCrLf
                        sSql = sSql & "                              ,SGI_QTDE                  = " & arrPROGENTREGAS(i, 1) & vbCrLf
                        sSql = sSql & "                              ,SGI_DATENTREGA            = " & arrPROGENTREGAS(i, 2) & vbCrLf
                        sSql = sSql & "                              ,SGI_FECHTPFU              = " & arrPROGENTREGAS(i, 5) & vbCrLf
                        sSql = sSql & "                              ,SGI_OBSOP                 = '" & arrPROGENTREGAS(i, 4) & "'" & vbCrLf
                        sSql = sSql & "                              ,SGI_JAIMPRESSA            = 0" & vbCrLf
                        sSql = sSql & "                              ,SGI_ALTERADO              = 1" & vbCrLf
                        sSql = sSql & "                              ,SGI_ALTFILM               = " & arrPROGENTREGAS(i, 12) & vbCrLf
                        sSql = sSql & "                              ,SGI_FOTNOVO               = " & arrPROGENTREGAS(i, 13) & vbCrLf
                        sSql = sSql & "                              ,SGI_REPETICAO             = " & arrPROGENTREGAS(i, 14) & vbCrLf
                        sSql = sSql & "                              ,SGI_QTDEPORPALHETS        = " & arrPROGENTREGAS(i, 18) & vbCrLf
                        sSql = sSql & "                              ,SGI_PALHETS               = " & arrPROGENTREGAS(i, 19) & vbCrLf
                        sSql = sSql & "                              ,SGI_STATUS                = " & arrPROGENTREGAS(i, 17) & vbCrLf
                        
                        sSql = sSql & "                         Where " & vbCrLf
                        sSql = sSql & "                               SGI_FILIAL     = " & intFILIAL & vbCrLf
                        sSql = sSql & "                           And SGI_CODIGO     = " & Trim(arrPROGENTREGAS(i, 9)) & vbCrLf
                        sSql = sSql & "                           And SGI_IDPRODUTO  = " & arrPROGENTREGAS(i, 0) & vbCrLf
                        sSql = sSql & "                           And SGI_IDPAI      = " & Trim(arrPROGENTREGAS(i, 16))
                    
                        BGRV.CommandText = sSql
                        BGRV.Execute
                    
                    End If
                End If
           Next i
        End If
        '' =====================================
         
         
         '' Itens do pedido
        If IsArray(arrPRODUTOS) = True Then
           For i = 1 To UBound(arrPRODUTOS)
               If arrPRODUTOS(i, 17) = dacEnumUpdateAction_update Then '' Altera
               
                  sSql = ""
                  
                  sSql = "Update SGI_CADPEDVENDI" & strNOTABFILIAL & " Set " & vbCrLf
                  sSql = sSql & "                                 SGI_QTDE              = " & arrPRODUTOS(i, 2) & vbCrLf
                  sSql = sSql & "                                ,SGI_VLUNIT            = " & arrPRODUTOS(i, 3) & vbCrLf
                  sSql = sSql & "                                ,SGI_PRCDESC           = " & arrPRODUTOS(i, 4) & vbCrLf
                  sSql = sSql & "                                ,SGI_PRCIPI            = " & arrPRODUTOS(i, 5) & vbCrLf
                  sSql = sSql & "                                ,SGI_VLTOT             = " & arrPRODUTOS(i, 6) & vbCrLf
                  sSql = sSql & "                                ,SGI_VLIPI             = " & arrPRODUTOS(i, 7) & vbCrLf
                  sSql = sSql & "                                ,SGI_VLDESC            = " & arrPRODUTOS(i, 8) & vbCrLf
                  sSql = sSql & "                                ,SGI_ALTFILM           = " & arrPRODUTOS(i, 10) & vbCrLf
                  sSql = sSql & "                                ,SGI_FOTNOVO           = " & arrPRODUTOS(i, 11) & vbCrLf
                  sSql = sSql & "                                ,SGI_REPETICAO         = " & arrPRODUTOS(i, 12) & vbCrLf
                  sSql = sSql & "                                ,SGI_CODLINPROD        = " & arrPRODUTOS(i, 13) & vbCrLf
                  sSql = sSql & "                                ,SGI_OBSOP             = '" & arrPRODUTOS(i, 15) & "'" & vbCrLf
                  sSql = sSql & "                                ,SGI_FECHTPFU          = " & arrPRODUTOS(i, 16) & vbCrLf
                  sSql = sSql & "                                ,SGI_QTDEPORPALHETS    = " & arrPRODUTOS(i, 19) & vbCrLf
                  sSql = sSql & "                                ,SGI_PALHETS           = " & arrPRODUTOS(i, 20) & vbCrLf
                  sSql = sSql & "                                ,SGI_CONFERIDO         = " & arrPRODUTOS(i, 21) & vbCrLf
                  sSql = sSql & "                                ,SGI_USARPADRPALSN     = " & arrPRODUTOS(i, 22) & vbCrLf
                  sSql = sSql & "                                ,SGI_DESCPROD          = " & arrPRODUTOS(i, 23) & vbCrLf
                  
                  sSql = sSql & "                                ,SGI_FechSoldaAgrafado = " & arrPRODUTOS(i, 24) & vbCrLf
                  sSql = sSql & "                                ,SGI_VernCorpo         = " & arrPRODUTOS(i, 25) & vbCrLf
                  sSql = sSql & "                                ,SGI_VernTampa         = " & arrPRODUTOS(i, 26) & vbCrLf
                  sSql = sSql & "                                ,SGI_VernFundo         = " & arrPRODUTOS(i, 27) & vbCrLf
                  sSql = sSql & "                                ,SGI_VernArgola         = " & arrPRODUTOS(i, 28) & vbCrLf
                  
                  sSql = sSql & "Where " & vbCrLf
                  sSql = sSql & "      SGI_FILIAL    = " & intFILIAL & vbCrLf
                  sSql = sSql & "  And SGI_CODIGO    = " & lngCODPEDIDO & vbCrLf
                  sSql = sSql & "  And SGI_IDPRODUTO = " & arrPRODUTOS(i, 9)
                  
                  BGRV.CommandText = sSql
                  BGRV.Execute
                  
                  '' Alteração no Filme
                  '' Fotolito Novo
                  If arrPRODUTOS(i, 10) = 1 Or _
                     arrPRODUTOS(i, 11) = 1 Then
                        
                        sSql = ""
                        
                        sSql = "Update SGI_CADPRODUTO Set SGI_STATUS = 2 " & vbCrLf
                        sSql = sSql & " Where " & vbCrLf
                        sSql = sSql & "       SGI_FILIAL    = " & intFILIAL & vbCrLf
                        sSql = sSql & "   And SGI_IDPRODUTO = " & arrPRODUTOS(i, 9)
                           
                        BGRV.CommandText = sSql
                        BGRV.Execute
                  
                  End If
                  
               ElseIf arrPRODUTOS(i, 17) = dacEnumUpdateAction_delete Then '' Exclui

                  sSql = ""
                  
                  sSql = "Delete SGI_CADPEDVENDI" & strNOTABFILIAL & vbCrLf
                  sSql = sSql & "Where " & vbCrLf
                  sSql = sSql & "      SGI_FILIAL    = " & intFILIAL & vbCrLf
                  sSql = sSql & "  And SGI_CODIGO    = " & lngCODPEDIDO & vbCrLf
                  sSql = sSql & "  And SGI_IDPRODUTO = " & arrPRODUTOS(i, 9)
                  
                  BGRV.CommandText = sSql
                  BGRV.Execute

               End If
           Next i
        End If
     
        '' =====================================
        '' Grava Log
        sSql = ""
        
        sSql = "Insert Into SGI_LOGPED" & strNOTABFILIAL & " (" & vbCrLf
        sSql = sSql & "                              SGI_FILIAL" & vbCrLf
        sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                             ,SGI_DATA" & vbCrLf
        sSql = sSql & "                             ,SGI_HORA" & vbCrLf
        sSql = sSql & "                             ,SGI_CODUSUARIO" & vbCrLf
        sSql = sSql & "                             ,SGI_ACAO" & vbCrLf
        sSql = sSql & "                    ) Values (" & vbCrLf
        sSql = sSql & "                              " & FILIAL & vbCrLf
        sSql = sSql & "                             ," & lngCODPEDIDO & vbCrLf
        sSql = sSql & "                             ,'" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                             ,'" & Format(Now, "HH:MM:SS") & "'" & vbCrLf
        sSql = sSql & "                             ," & lngCodUsuario & vbCrLf
        sSql = sSql & "                             ,'" & strOPER & "'" & vbCrLf
        sSql = sSql & "                             )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        
        If IsArray(arrPROGENTREGAS) Then
            For i = 1 To UBound(arrPROGENTREGAS)
                If arrPROGENTREGAS(i, 20) = dacEnumUpdateAction_update Then
                
                    '' =====================================
                    '' Grava Log
                    sSql = ""
                    
                    sSql = "Insert Into SGI_LOGPED" & strNOTABFILIAL & " (" & vbCrLf
                    sSql = sSql & "                              SGI_FILIAL" & vbCrLf
                    sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
                    sSql = sSql & "                             ,SGI_DATA" & vbCrLf
                    sSql = sSql & "                             ,SGI_HORA" & vbCrLf
                    sSql = sSql & "                             ,SGI_CODUSUARIO" & vbCrLf
                    sSql = sSql & "                             ,SGI_ACAO" & vbCrLf
                    sSql = sSql & "                             ,SGI_DESCACAO" & vbCrLf
                    sSql = sSql & "                    ) Values (" & vbCrLf
                    sSql = sSql & "                              " & FILIAL & vbCrLf
                    sSql = sSql & "                             ," & lngCODPEDIDO & vbCrLf
                    sSql = sSql & "                             ,'" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf
                    sSql = sSql & "                             ,'" & Format(Now, "HH:MM:SS") & "'" & vbCrLf
                    sSql = sSql & "                             ," & lngCodUsuario & vbCrLf
                    sSql = sSql & "                             ,'AD'" & vbCrLf
                    sSql = sSql & "                             ,'" & arrPROGENTREGAS(i, 22) & " para " & arrPROGENTREGAS(i, 21) & "'"
                    sSql = sSql & "                             )"
                    
                    BGRV.CommandText = sSql
                    BGRV.Execute
                
                End If
            Next i
        End If
        
        '' =====================================
             
             
     
     
     
     ElseIf strOPER = "E" Then '' Exclui
     
         arrPRODCOPIA = arrPRODUTOS
     
         '' Apagando Tipo de Reprovação
         sSql = "Delete from SGI_TIPREPRPED" & strNOTABFILIAL & vbCrLf
         sSql = sSql & "      Where " & vbCrLf
         sSql = sSql & "           SGI_FILIAL = " & intFILIAL & vbCrLf
         sSql = sSql & "       And SGI_CODPED = " & lngCODPEDIDO
         
         BGRV.CommandText = sSql
         BGRV.Execute
         
         '' Apagando itens de Programacao
         sSql = "Delete from SGI_PROGENTRPROD" & strNOTABFILIAL & vbCrLf
         sSql = sSql & "        Where SGI_FILIAL = " & intFILIAL & vbCrLf
         sSql = sSql & "          And SGI_CODPED = " & lngCODPEDIDO
         
         BGRV.CommandText = sSql
         BGRV.Execute
         
         '' Apagando itens do pedido
         sSql = "Delete from SGI_CADPEDVENDI" & strNOTABFILIAL & vbCrLf
         sSql = sSql & "        Where SGI_FILIAL = " & intFILIAL & vbCrLf
         sSql = sSql & "          And SGI_CODIGO = " & lngCODPEDIDO
         
         BGRV.CommandText = sSql
         BGRV.Execute
     
         '' Apagando Cabeçalho do pedido
         sSql = "Delete from SGI_CADPEDVENDH" & strNOTABFILIAL & vbCrLf
         sSql = sSql & "        Where SGI_FILIAL = " & intFILIAL & vbCrLf
         sSql = sSql & "          And SGI_CODIGO = " & lngCODPEDIDO
         
         BGRV.CommandText = sSql
         BGRV.Execute
         
     ElseIf strOPER = "LF" Or strOPER = "LN" Then
        '' Libera Pedido Bloqueado LN = Comercial
        '' Libera Pedido Bloqueado LF = Financeiro
     
        sSql = "Update SGI_CADPEDVENDH" & strNOTABFILIAL & " Set " & vbCrLf
        sSql = sSql & "                           SGI_LIBERADO    = '" & strLIBERADO & "'" & vbCrLf
        sSql = sSql & "                          ,SGI_LIBDATAHORA = '" & Format(dtDTHORALIB, "MM/DD/YYYY HH:MM:SS") & "'" & vbCrLf
        sSql = sSql & "                          ,SGI_OBSLIBCOM   = '" & strOBSCOM & "'" & vbCrLf
        sSql = sSql & "                          ,SGI_STATUS      = '" & Trim(strSTATUS) & "'" & vbCrLf
        sSql = sSql & "                          ,SGI_CONFERIDO   = " & intCONFERIDO & vbCrLf
        
        sSql = sSql & "                     Where " & vbCrLf & vbCrLf
        sSql = sSql & "                           SGI_FILIAL = " & intFILIAL & vbCrLf
        sSql = sSql & "                       And SGI_CODIGO = " & lngCODPEDIDO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        If strOPER = "LF" Then
            '' Gravando Itens da Ordem
            If IsArray(arrPROGENTREGAS) Then
                For i = 1 To UBound(arrPROGENTREGAS)
                
                    If Len(Trim(arrPROGENTREGAS(i, 0))) > 0 Then
                    
                        sSql = "Insert Into SGI_ORDEMPROD" & strNOTABFILIAL & " (" & vbCrLf
                        sSql = sSql & "                       SGI_FILIAL" & vbCrLf
                        sSql = sSql & "                      ,SGI_CODIGO" & vbCrLf
                        sSql = sSql & "                      ,SGI_DATAORDEM" & vbCrLf
                        sSql = sSql & "                      ,SGI_IDPRODUTO" & vbCrLf
                        sSql = sSql & "                      ,SGI_CODPROD" & vbCrLf
                        sSql = sSql & "                      ,SGI_CODPED" & vbCrLf
                        sSql = sSql & "                      ,SGI_DATAPED" & vbCrLf
                        
                        sSql = sSql & "                      ,SGI_QTDEPED" & vbCrLf
                        sSql = sSql & "                      ,SGI_QTDEJAPROG" & vbCrLf
                        sSql = sSql & "                      ,SGI_QTDE" & vbCrLf
                        sSql = sSql & "                      ,SGI_SALDO" & vbCrLf
                        sSql = sSql & "                      ,SGI_TIPO" & vbCrLf
                        
                        sSql = sSql & "                      ,SGI_ALTFILM" & vbCrLf
                        sSql = sSql & "                      ,SGI_FOTNOVO" & vbCrLf
                        sSql = sSql & "                      ,SGI_REPETICAO" & vbCrLf
                        sSql = sSql & "                      ,SGI_DATENTREGA" & vbCrLf
                        sSql = sSql & "                      ,SGI_TIPOP" & vbCrLf
                        sSql = sSql & "                      ,SGI_JAIMPRESSA" & vbCrLf
                        sSql = sSql & "                      ,SGI_OPCOMPL" & vbCrLf
                        sSql = sSql & "                      ,SGI_NOMEVEND" & vbCrLf
                        sSql = sSql & "                      ,SGI_OBSOP" & vbCrLf
                        sSql = sSql & "                      ,SGI_FECHTPFU" & vbCrLf
                        sSql = sSql & "                      ,SGI_FILIALPED" & vbCrLf
                        sSql = sSql & "                      ,SGI_ALTERADO" & vbCrLf
                        sSql = sSql & "                      ,SGI_IDPAI" & vbCrLf
                        sSql = sSql & "                      ,SGI_QTDEPORPALHETS" & vbCrLf
                        sSql = sSql & "                      ,SGI_PALHETS" & vbCrLf
                        sSql = sSql & "                      ,SGI_STATUS" & vbCrLf
                        sSql = sSql & "                      ,SGI_PROGRAMADO" & vbCrLf
                        
                        
                        sSql = sSql & "                    ) Values (" & vbCrLf
                        
                        sSql = sSql & "                       " & intFILIAL & vbCrLf
                        sSql = sSql & "                      ," & arrPROGENTREGAS(i, 4) & vbCrLf
                        sSql = sSql & "                      ,'" & Format(CDate(arrPROGENTREGAS(i, 5)), "MM/DD/YYYY") & "'" & vbCrLf
                        sSql = sSql & "                      ," & arrPROGENTREGAS(i, 0) & vbCrLf
                        sSql = sSql & "                      ,'" & arrPROGENTREGAS(i, 6) & "'" & vbCrLf
                        sSql = sSql & "                      ," & lngCODPEDIDO & vbCrLf
                        sSql = sSql & "                      ,'" & Format(CDate(dtDATAPED), "MM/DD/YYYY") & "'" & vbCrLf
                        
                        sValor = "Null"
                        If Len(Trim(arrPROGENTREGAS(i, 1))) > 0 And arrPROGENTREGAS(i, 1) > 0 Then
                            sValor = Replace(arrPROGENTREGAS(i, 1), ".", "")
                            sValor = Replace(sValor, ",", ".")
                        End If
                        sSql = sSql & "                      ," & sValor & vbCrLf
                        
                        sValor = "Null"
                        sSql = sSql & "                      ," & sValor & vbCrLf
                        
                        sValor = "Null"
                        If Len(Trim(arrPROGENTREGAS(i, 1))) > 0 And arrPROGENTREGAS(i, 1) > 0 Then
                            sValor = Replace(arrPROGENTREGAS(i, 1), ".", "")
                            sValor = Replace(sValor, ",", ".")
                        End If
                        sSql = sSql & "                      ," & sValor & vbCrLf
                        
                        sValor = "Null"
                        sSql = sSql & "                      ," & sValor & vbCrLf
                        sSql = sSql & "                      ,1" & vbCrLf
                        
                        sSql = sSql & "                      ," & arrPROGENTREGAS(i, 7) & vbCrLf
                        sSql = sSql & "                      ," & arrPROGENTREGAS(i, 8) & vbCrLf
                        sSql = sSql & "                      ," & arrPROGENTREGAS(i, 9) & vbCrLf
                        
                        sSql = sSql & "                      ,'" & Format(arrPROGENTREGAS(i, 2), "MM/DD/YYYY") & "'" & vbCrLf
                        sSql = sSql & "                      ,1" & vbCrLf
                        sSql = sSql & "                      ,0" & vbCrLf
                        sSql = sSql & "                      ,0" & vbCrLf
                        sSql = sSql & "                      ,'" & Trim(strNomeVend) & "'" & vbCrLf
                        sSql = sSql & "                      ,'" & Trim(arrPROGENTREGAS(i, 10)) & "'" & vbCrLf
                        sSql = sSql & "                      ," & Trim(arrPROGENTREGAS(i, 11)) & vbCrLf
                        sSql = sSql & "                      ," & intFILIALPED & vbCrLf
                        sSql = sSql & "                      ,0" & vbCrLf
                        sSql = sSql & "                      ," & Trim(arrPROGENTREGAS(i, 14)) & vbCrLf
                        sSql = sSql & "                      ," & arrPROGENTREGAS(i, 15) & vbCrLf
                        sSql = sSql & "                      ," & arrPROGENTREGAS(i, 16) & vbCrLf
                        sSql = sSql & "                      ," & arrPROGENTREGAS(i, 17) & vbCrLf '' Status
                        sSql = sSql & "                      ,0" & vbCrLf
                        
                        sSql = sSql & "                    )"
                        
                        BGRV.CommandText = sSql
                        BGRV.Execute
                                        
                                        
                                        
                        '' =============================
                        '' Dia 19/04/2015
                        '' Dando Update na Tabela do Item do Pedido ( Progrmação de Entrega"
                        
                        '' =====================================
                        '' Programação de Entregas
                        sSql = ""
                        
                        sSql = "Update SGI_PROGENTRPROD" & strNOTABFILIAL & " Set" & vbCrLf
                        sSql = sSql & "                       SGI_STATUS = " & arrPROGENTREGAS(i, 17) & vbCrLf
                        sSql = sSql & "                     Where " & vbCrLf
                        sSql = sSql & "                           SGI_FILIAL    = " & intFILIAL & vbCrLf
                        sSql = sSql & "                       And SGI_CODPED    = " & lngCODPEDIDO & vbCrLf
                        sSql = sSql & "                       And SGI_IDPRODUTO = " & arrPROGENTREGAS(i, 0) & vbCrLf
                        sSql = sSql & "                       And SGI_IDINTERNO = " & arrPROGENTREGAS(i, 14) & vbCrLf
                        
                        BGRV.CommandText = sSql
                        BGRV.Execute
                                   
                                        
                    Else
                        If boolFOIREPROVADO = True Then '' Foi reprovado e agora esta voltando a OP
                            
                            sSql = ""
                            
                            sSql = "Update SGI_PROGENTRPROD" & strNOTABFILIAL & " Set " & vbCrLf
                            sSql = sSql & "                               SGI_STATUS = 0" & vbCrLf
                            sSql = sSql & "      Where " & vbCrLf
                            sSql = sSql & "            SGI_FILIAL    = " & intFILIAL & vbCrLf
                            sSql = sSql & "        And SGI_CODPED    = " & lngCODPEDIDO & vbCrLf
                            sSql = sSql & "        And SGI_IDINTERNO = " & arrPROGENTREGAS(i, 16)
                            
                            BGRV.CommandText = sSql
                            BGRV.Execute
                            
                            If Len(Trim(arrPROGENTREGAS(i, 9))) > 0 Then
                                
                                sSql = ""
                                
                                sSql = "Update SGI_ORDEMPROD" & strNOTABFILIAL & vbCrLf
                                sSql = sSql & "                           Set SGI_STATUS = 0 " & vbCrLf
                                sSql = sSql & "                     Where " & vbCrLf
                                sSql = sSql & "                           SGI_FILIAL = " & intFILIAL & vbCrLf
                                sSql = sSql & "                       And SGI_CODIGO = " & arrPROGENTREGAS(i, 9)
                                
                                BGRV.CommandText = sSql
                                BGRV.Execute
                                
                            End If
                        Else
                        
                        
                            If Trim(strSTATUS) = "N" Or Trim(strSTATUS) = "L" Then
                        
                                sSql = ""
                                
                                sSql = "Update SGI_PROGENTRPROD" & strNOTABFILIAL & " Set " & vbCrLf
                                sSql = sSql & "                               SGI_STATUS = 0" & vbCrLf
                                sSql = sSql & "      Where " & vbCrLf
                                sSql = sSql & "            SGI_FILIAL    = " & intFILIAL & vbCrLf
                                sSql = sSql & "        And SGI_CODPED    = " & lngCODPEDIDO & vbCrLf
                                sSql = sSql & "        And SGI_IDINTERNO = " & arrPROGENTREGAS(i, 16) & vbCrLf
                                sSql = sSql & "        And SGI_IDPRODUTO = " & arrPROGENTREGAS(i, 18)
                                
                                BGRV.CommandText = sSql
                                BGRV.Execute
                                
                                    
                                sSql = ""
                                
                                sSql = "Update SGI_ORDEMPROD" & strNOTABFILIAL & vbCrLf
                                sSql = sSql & "                           Set SGI_STATUS = 0 " & vbCrLf
                                sSql = sSql & "                     Where " & vbCrLf
                                sSql = sSql & "                           SGI_FILIAL     = " & intFILIAL & vbCrLf
                                sSql = sSql & "                       And SGI_CODIGO     = " & arrPROGENTREGAS(i, 6) & vbCrLf
                                sSql = sSql & "                       And SGI_IDPAI      = " & arrPROGENTREGAS(i, 16) & vbCrLf
                                sSql = sSql & "                       And SGI_IDPRODUTO  = " & arrPROGENTREGAS(i, 18)
                                
                                BGRV.CommandText = sSql
                                BGRV.Execute
                            
                            End If
                                
                        
                        End If
                    End If
                
                Next i
            End If
            '' =====================================================================
        Else
        
            '' Alterando Conferido
            If IsArray(arrPRODUTOS) Then
                For i = 1 To UBound(arrPRODUTOS)
                
                    sSql = ""
                    
                    sSql = "Update SGI_CADPEDVENDI" & strNOTABFILIAL & " Set" & vbCrLf
                    sSql = sSql & "                                 SGI_CONFERIDO = " & arrPRODUTOS(i, 21) & vbCrLf
                    
                    sSql = sSql & "Where" & vbCrLf
                    sSql = sSql & "      SGI_FILIAL    = " & intFILIAL & vbCrLf
                    sSql = sSql & "  And SGI_CODIGO    = " & lngCODPEDIDO & vbCrLf
                    sSql = sSql & "  And SGI_IDPRODUTO = " & arrPRODUTOS(i, 9)
            
                    BGRV.CommandText = sSql
                    BGRV.Execute
                
                Next i
            End If
            
            '' Gravando Itens da Ordem
            If IsArray(arrPROGENTREGAS) Then
                For i = 1 To UBound(arrPROGENTREGAS)
                    If boolFOIREPROVADO = True Then '' Foi reprovado e agora esta voltando a OP
                        
                        sSql = ""
                        
                        sSql = "Update SGI_PROGENTRPROD" & strNOTABFILIAL & " Set " & vbCrLf
                        sSql = sSql & "                               SGI_STATUS     = " & arrPROGENTREGAS(i, 17) & vbCrLf
                        sSql = sSql & "                              ,SGI_DATENTREGA = " & arrPROGENTREGAS(i, 2) & vbCrLf
                        sSql = sSql & "      Where " & vbCrLf
                        sSql = sSql & "            SGI_FILIAL    = " & intFILIAL & vbCrLf
                        sSql = sSql & "        And SGI_CODPED    = " & lngCODPEDIDO & vbCrLf
                        sSql = sSql & "        And SGI_IDINTERNO = " & arrPROGENTREGAS(i, 16)
                        
                        BGRV.CommandText = sSql
                        BGRV.Execute
                        
                        If Len(Trim(arrPROGENTREGAS(i, 9))) > 0 Then
                            
                            sSql = ""
                            
                            sSql = "Update SGI_ORDEMPROD" & strNOTABFILIAL & vbCrLf
                            sSql = sSql & "                           Set SGI_STATUS     = " & arrPROGENTREGAS(i, 17) & vbCrLf
                            sSql = sSql & "                              ,SGI_DATENTREGA = " & arrPROGENTREGAS(i, 2) & vbCrLf
                            
                            sSql = sSql & "                     Where " & vbCrLf
                            sSql = sSql & "                           SGI_FILIAL = " & intFILIAL & vbCrLf
                            sSql = sSql & "                       And SGI_CODIGO = " & arrPROGENTREGAS(i, 9) & vbCrLf
                            sSql = sSql & "                       And SGI_IDPAI  = " & arrPROGENTREGAS(i, 16)
                            
                            BGRV.CommandText = sSql
                            BGRV.Execute
                            
                        End If
                    End If
                Next i
            End If
            '' =====================================================================
        
        End If
     
     
        '' =====================================
        '' Grava Log
        sSql = ""
        
        sSql = "Insert Into SGI_LOGPED" & strNOTABFILIAL & " (" & vbCrLf
        sSql = sSql & "                              SGI_FILIAL" & vbCrLf
        sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                             ,SGI_DATA" & vbCrLf
        sSql = sSql & "                             ,SGI_HORA" & vbCrLf
        sSql = sSql & "                             ,SGI_CODUSUARIO" & vbCrLf
        sSql = sSql & "                             ,SGI_ACAO" & vbCrLf
        sSql = sSql & "                    ) Values (" & vbCrLf
        sSql = sSql & "                              " & FILIAL & vbCrLf
        sSql = sSql & "                             ," & lngCODPEDIDO & vbCrLf
        sSql = sSql & "                             ,'" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                             ,'" & Format(Now, "HH:MM:SS") & "'" & vbCrLf
        sSql = sSql & "                             ," & lngCodUsuario & vbCrLf
        sSql = sSql & "                             ,'" & strOPER & "'" & vbCrLf
        sSql = sSql & "                             )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' =====================================
     
     
     ElseIf strOPER = "R" Then '' Reprova Pedido
        
        sSql = "Delete from SGI_TIPREPRPED" & strNOTABFILIAL & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "           SGI_FILIAL = " & intFILIAL & vbCrLf
        sSql = sSql & "       And SGI_CODPED = " & lngCODPEDIDO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        sSql = "Update SGI_CADPEDVENDH" & strNOTABFILIAL & " Set " & vbCrLf
        sSql = sSql & "                           SGI_REPROCADO   = '" & strLIBERADO & "'" & vbCrLf
        sSql = sSql & "                          ,SGI_REPDATAHORA = '" & Format(dtDTHORALIB, "MM/DD/YYYY HH:MM:SS") & "'" & vbCrLf
        sSql = sSql & "                          ,SGI_OBSLIBCOM   = '" & strOBSCOM & "'" & vbCrLf
        sSql = sSql & "                          ,SGI_STATUS      = '" & strSTATUS & "'" & vbCrLf
        
        sSql = sSql & "                     Where " & vbCrLf & vbCrLf
        sSql = sSql & "                           SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "                       And SGI_CODIGO = " & lngCODPEDIDO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' ------------------------------------------------------
        '' Gravando Tipo de Reprovação se Olver
        If IsArray(arrTIPREPROVA) = True Then
           For i = 1 To UBound(arrTIPREPROVA)
        
               sSql = "Insert into SGI_TIPREPRPED" & strNOTABFILIAL & " (" & vbCrLf
               sSql = sSql & "                     SGI_FILIAL" & vbCrLf
               sSql = sSql & "                    ,SGI_CODPED" & vbCrLf
               sSql = sSql & "                    ,SGI_CODREP" & vbCrLf
               sSql = sSql & "                    ,SGI_FILIALPED" & vbCrLf
               sSql = sSql & "                 ) Values (" & vbCrLf
               sSql = sSql & "                            " & intFILIAL & vbCrLf
               sSql = sSql & "                           ," & lngCODPEDIDO & vbCrLf
               sSql = sSql & "                           ," & arrTIPREPROVA(i) & vbCrLf
               sSql = sSql & "                           ," & intFILIALPED & vbCrLf
               sSql = sSql & "                            )"
        
               BGRV.CommandText = sSql
               BGRV.Execute
               
           Next i
        End If
        '' ------------------------------------------------------
        
        '' --------------------------------------------
        '' Mudando o Status da OP Geradas Pelo Sistema para Reprovada
        If IsArray(arrPROGENTREGAS) Then
            For i = 1 To UBound(arrPROGENTREGAS)
                If Len(Trim(Replace(arrPROGENTREGAS(i, 6), "/", ""))) > 0 Then
                    
                    sSql = ""
            
                    sSql = "Update SGI_PROGENTRPROD" & strNOTABFILIAL & " Set " & vbCrLf
                    sSql = sSql & "                               SGI_STATUS = 3" & vbCrLf
                    sSql = sSql & "      Where " & vbCrLf
                    sSql = sSql & "            SGI_FILIAL    = " & intFILIAL & vbCrLf
                    sSql = sSql & "        And SGI_CODPED    = " & lngCODPEDIDO & vbCrLf
                    sSql = sSql & "        And SGI_IDPRODUTO = " & arrPROGENTREGAS(i, 0) & vbCrLf
                    sSql = sSql & "        And SGI_IDINTERNO = " & arrPROGENTREGAS(i, 16)
                    
                    BGRV.CommandText = sSql
                    BGRV.Execute
                    
                    sSql = ""
            
                    sSql = "Update SGI_ORDEMPROD" & strNOTABFILIAL & " Set " & vbCrLf
                    sSql = sSql & "                               SGI_STATUS = 3"
                    sSql = sSql & "      Where " & vbCrLf
                    sSql = sSql & "            SGI_FILIAL    = " & intFILIAL & vbCrLf
                    sSql = sSql & "        And SGI_CODIGO    = " & Trim(Replace(arrPROGENTREGAS(i, 6), "/", "")) & vbCrLf
                    sSql = sSql & "        And SGI_IDPRODUTO = " & arrPROGENTREGAS(i, 0) & vbCrLf
                    sSql = sSql & "        And SGI_IDPAI     = " & arrPROGENTREGAS(i, 16)
                    
                    BGRV.CommandText = sSql
                    BGRV.Execute
                    '' --------------------------------------------
                
                End If
            
            Next i
        End If
        
        '' =====================================
        '' Grava Log
        sSql = ""
        
        sSql = "Insert Into SGI_LOGPED" & strNOTABFILIAL & " (" & vbCrLf
        sSql = sSql & "                              SGI_FILIAL" & vbCrLf
        sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                             ,SGI_DATA" & vbCrLf
        sSql = sSql & "                             ,SGI_HORA" & vbCrLf
        sSql = sSql & "                             ,SGI_CODUSUARIO" & vbCrLf
        sSql = sSql & "                             ,SGI_ACAO" & vbCrLf
        sSql = sSql & "                    ) Values (" & vbCrLf
        sSql = sSql & "                              " & FILIAL & vbCrLf
        sSql = sSql & "                             ," & lngCODPEDIDO & vbCrLf
        sSql = sSql & "                             ,'" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                             ,'" & Format(Now, "HH:MM:SS") & "'" & vbCrLf
        sSql = sSql & "                             ," & lngCodUsuario & vbCrLf
        sSql = sSql & "                             ,'" & strOPER & "'" & vbCrLf
        sSql = sSql & "                             )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' =====================================
     
     
     ElseIf strOPER = "D" Then '' Bloqueia o pedido manualmente
        
        sSql = "Update SGI_CADPEDVENDH" & strNOTABFILIAL & " Set " & vbCrLf
        sSql = sSql & "                           SGI_REPROCADO   = '" & strLIBERADO & "'" & vbCrLf
        sSql = sSql & "                          ,SGI_REPDATAHORA = '" & Format(dtDTHORALIB, "MM/DD/YYYY HH:MM:SS") & "'" & vbCrLf
        sSql = sSql & "                          ,SGI_OBSLIBCOM   = '" & strOBSCOM & "'" & vbCrLf
        sSql = sSql & "                          ,SGI_STATUS      = '" & strSTATUS & "'" & vbCrLf
        
        sSql = sSql & "                     Where " & vbCrLf & vbCrLf
        sSql = sSql & "                           SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "                       And SGI_CODIGO = " & lngCODPEDIDO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        ''  --------------------------------------------
     
     
        '' --------------------------------------------
        '' Mudando o Status da OP Geradas Pelo Sistema para Reprovada
        If IsArray(arrPROGENTREGAS) Then
            For i = 1 To UBound(arrPROGENTREGAS)
        
                sSql = ""
        
                sSql = "Update SGI_PROGENTRPROD" & strNOTABFILIAL & " Set " & vbCrLf
                sSql = sSql & "                               SGI_STATUS = 4" & vbCrLf
                sSql = sSql & "      Where " & vbCrLf
                sSql = sSql & "            SGI_FILIAL    = " & intFILIAL & vbCrLf
                sSql = sSql & "        And SGI_CODPED    = " & lngCODPEDIDO & vbCrLf
                sSql = sSql & "        And SGI_IDPRODUTO = " & arrPROGENTREGAS(i, 0) & vbCrLf
                sSql = sSql & "        And SGI_IDINTERNO = " & arrPROGENTREGAS(i, 16)
                
                BGRV.CommandText = sSql
                BGRV.Execute
                
                If Len(Trim(Replace(arrPROGENTREGAS(i, 6), "/", ""))) > 0 Then
                    
                    sSql = ""
            
                    sSql = "Update SGI_ORDEMPROD" & strNOTABFILIAL & " Set " & vbCrLf
                    sSql = sSql & "                               SGI_STATUS = 4" & vbCrLf
                    sSql = sSql & "      Where " & vbCrLf
                    sSql = sSql & "            SGI_FILIAL       = " & intFILIAL & vbCrLf
                    sSql = sSql & "        And SGI_CODIGO       = " & Trim(Replace(arrPROGENTREGAS(i, 6), "/", "")) & vbCrLf
                    sSql = sSql & "        And SGI_IDPRODUTO    = " & arrPROGENTREGAS(i, 0) & vbCrLf
                    sSql = sSql & "        And SGI_IDPAI        = " & arrPROGENTREGAS(i, 16)
                    
                    BGRV.CommandText = sSql
                    BGRV.Execute
                    '' --------------------------------------------
                
                End If
            Next i
        End If
     
     
        '' =====================================
        '' Grava Log
        sSql = ""
        
        sSql = "Insert Into SGI_LOGPED" & strNOTABFILIAL & " (" & vbCrLf
        sSql = sSql & "                              SGI_FILIAL" & vbCrLf
        sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                             ,SGI_DATA" & vbCrLf
        sSql = sSql & "                             ,SGI_HORA" & vbCrLf
        sSql = sSql & "                             ,SGI_CODUSUARIO" & vbCrLf
        sSql = sSql & "                             ,SGI_ACAO" & vbCrLf
        sSql = sSql & "                    ) Values (" & vbCrLf
        sSql = sSql & "                              " & FILIAL & vbCrLf
        sSql = sSql & "                             ," & lngCODPEDIDO & vbCrLf
        sSql = sSql & "                             ,'" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                             ,'" & Format(Now, "HH:MM:SS") & "'" & vbCrLf
        sSql = sSql & "                             ," & lngCodUsuario & vbCrLf
        sSql = sSql & "                             ,'" & strOPER & "'" & vbCrLf
        sSql = sSql & "                             )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' =====================================
     
     
     ElseIf strOPER = "T" Then
     
        If IsArray(arrSERVICOS) Then
            For i = 1 To UBound(arrSERVICOS)
            
                '' Apagando Confirmação de Faturamento Itens
                sSql = "Delete from SGI_CADORDCONFI" & strNOTABFILIAL & vbCrLf
                sSql = sSql & "      Where " & vbCrLf
                sSql = sSql & "           SGI_FILIAL  = " & intFILIAL & vbCrLf
                sSql = sSql & "       And SGI_CODCONF = " & arrSERVICOS(i)
            
                BGRV.CommandText = sSql
                BGRV.Execute
                
                '' Apagando Confirmação de Faturamento Cabec
                sSql = "Delete from SGI_CADORDCONFH" & strNOTABFILIAL & vbCrLf
                sSql = sSql & "      Where " & vbCrLf
                sSql = sSql & "           SGI_FILIAL  = " & intFILIAL & vbCrLf
                sSql = sSql & "       And SGI_CODCONF = " & arrSERVICOS(i)

                BGRV.CommandText = sSql
                BGRV.Execute
            
            Next i
        End If
        
        If IsArray(arrPROGENTREGAS) Then
            For i = 1 To UBound(arrPROGENTREGAS)
            
                '' Apagando Ordem de Faturamento Itens
                sSql = "Delete from SGI_CADORDFATI" & strNOTABFILIAL & vbCrLf
                sSql = sSql & "      Where " & vbCrLf
                sSql = sSql & "           SGI_FILIAL = " & intFILIAL & vbCrLf
                sSql = sSql & "       And SGI_CODORD = " & arrPROGENTREGAS(i)
            
                BGRV.CommandText = sSql
                BGRV.Execute
                
                '' Apagando Ordem de Faturamento Cabec
                sSql = "Delete from SGI_CADORDFATH" & strNOTABFILIAL & vbCrLf
                sSql = sSql & "      Where " & vbCrLf
                sSql = sSql & "           SGI_FILIAL = " & intFILIAL & vbCrLf
                sSql = sSql & "       And SGI_CODORD = " & arrPROGENTREGAS(i)
            
                BGRV.CommandText = sSql
                BGRV.Execute
            
            Next i
        End If
        
        
        If IsArray(arrPRODUTOS) Then
        
            For i = 1 To UBound(arrPRODUTOS)
        
                '' Apagando Ordem de Produção
                sSql = "Delete from SGI_ORDEMPROD" & strNOTABFILIAL & vbCrLf
                sSql = sSql & "      Where " & vbCrLf
                sSql = sSql & "           SGI_FILIAL = " & intFILIAL & vbCrLf
                sSql = sSql & "       And SGI_CODPED = " & arrPRODUTOS(i)
            
                BGRV.CommandText = sSql
                BGRV.Execute
                
                '' Apagando Programação de Entegras
                sSql = "Delete from SGI_PROGENTRPROD" & strNOTABFILIAL & vbCrLf
                sSql = sSql & "      Where " & vbCrLf
                sSql = sSql & "           SGI_FILIAL = " & intFILIAL & vbCrLf
                sSql = sSql & "       And SGI_CODPED = " & arrPRODUTOS(i)
            
                BGRV.CommandText = sSql
                BGRV.Execute
                
                '' Apagando Itens de Pedido
                sSql = "Delete from SGI_CADPEDVENDI" & strNOTABFILIAL & vbCrLf
                sSql = sSql & "      Where " & vbCrLf
                sSql = sSql & "           SGI_FILIAL = " & intFILIAL & vbCrLf
                sSql = sSql & "       And SGI_CODIGO = " & arrPRODUTOS(i)
            
                BGRV.CommandText = sSql
                BGRV.Execute
                
                '' Apagando Pedido
                sSql = "Delete from SGI_CADPEDVENDH" & strNOTABFILIAL & vbCrLf
                sSql = sSql & "      Where " & vbCrLf
                sSql = sSql & "           SGI_FILIAL = " & intFILIAL & vbCrLf
                sSql = sSql & "       And SGI_CODIGO = " & arrPRODUTOS(i)
                
                BGRV.CommandText = sSql
                BGRV.Execute
            
            
            Next i
            
        End If
        
     ElseIf strOPER = "M" Then
        
        If strSTATUS = "L" Or strSTATUS = "P" Then
            
            sSql = ""
            
            sSql = "Update SGI_CADPEDVENDH" & strNOTABFILIAL & " Set " & vbCrLf
            sSql = sSql & "                           SGI_STATUS      = 'M'" & vbCrLf
            sSql = sSql & "                          ,SGI_CODMOTLIQ   = " & strCODMOTLIQ & vbCrLf
            sSql = sSql & "                          ,SGI_OBSLIQ      = " & strOBSLIQ & vbCrLf
            
            sSql = sSql & "                     Where " & vbCrLf & vbCrLf
            sSql = sSql & "                           SGI_FILIAL = " & FILIAL & vbCrLf
            sSql = sSql & "                       And SGI_CODIGO = " & lngCODPEDIDO
            
            BGRV.CommandText = sSql
            BGRV.Execute
        End If
        
        '' Dando Update Nas OPS
        If IsArray(arrOPS) Then
            For i = 1 To UBound(arrOPS)
            
                sSql = "Update SGI_ORDEMPROD" & strNOTABFILIAL & " Set " & vbCrLf
                sSql = sSql & "                           SGI_STATUS      = 2" & vbCrLf
                
                sSql = sSql & "                     Where " & vbCrLf & vbCrLf
                sSql = sSql & "                           SGI_FILIAL = " & FILIAL & vbCrLf
                sSql = sSql & "                       And SGI_CODIGO = " & arrOPS(i)
                
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
     
        '' =====================================
        '' Grava Log
        sSql = ""
        
        sSql = "Insert Into SGI_LOGPED" & strNOTABFILIAL & " (" & vbCrLf
        sSql = sSql & "                              SGI_FILIAL" & vbCrLf
        sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                             ,SGI_DATA" & vbCrLf
        sSql = sSql & "                             ,SGI_HORA" & vbCrLf
        sSql = sSql & "                             ,SGI_CODUSUARIO" & vbCrLf
        sSql = sSql & "                             ,SGI_ACAO" & vbCrLf
        sSql = sSql & "                    ) Values (" & vbCrLf
        sSql = sSql & "                              " & FILIAL & vbCrLf
        sSql = sSql & "                             ," & lngCODPEDIDO & vbCrLf
        sSql = sSql & "                             ,'" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                             ,'" & Format(Now, "HH:MM:SS") & "'" & vbCrLf
        sSql = sSql & "                             ," & lngCodUsuario & vbCrLf
        sSql = sSql & "                             ,'" & strOPER & "'" & vbCrLf
        sSql = sSql & "                             )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' =====================================
     
     
     ElseIf strOPER = "LS" Then
        
        '' Liberando da Alteração
        sSql = "Update SGI_CADPEDVENDH" & strNOTABFILIAL & " Set " & vbCrLf
        sSql = sSql & "                           SGI_STATUS      = '" & strSTATUS & "'" & vbCrLf
        sSql = sSql & "                     Where " & vbCrLf & vbCrLf
        sSql = sSql & "                           SGI_FILIAL = " & intFILIAL & vbCrLf
        sSql = sSql & "                       And SGI_CODIGO = " & lngCODPEDIDO
        
        BGRV.CommandText = sSql
        BGRV.Execute
     
        '' --------------------------------------------
        '' Mudando o Status da OP Geradas Pelo Sistema para Reprovada
        If IsArray(arrPROGENTREGAS) Then
            For i = 1 To UBound(arrPROGENTREGAS)
        
                sSql = ""
        
                sSql = "Update SGI_PROGENTRPROD" & strNOTABFILIAL & " Set " & vbCrLf
                sSql = sSql & "                               SGI_STATUS = " & arrPROGENTREGAS(i, 17) & vbCrLf
                sSql = sSql & "      Where " & vbCrLf
                sSql = sSql & "            SGI_FILIAL    = " & intFILIAL & vbCrLf
                sSql = sSql & "        And SGI_CODPED    = " & lngCODPEDIDO & vbCrLf
                sSql = sSql & "        And SGI_IDINTERNO = " & arrPROGENTREGAS(i, 16)
                
                BGRV.CommandText = sSql
                BGRV.Execute
                
                If Len(Trim(Replace(arrPROGENTREGAS(i, 6), "/", ""))) > 0 Then
                    
                    sSql = ""
            
                    sSql = "Update SGI_ORDEMPROD" & strNOTABFILIAL & " Set " & vbCrLf
                    sSql = sSql & "                               SGI_STATUS = " & arrPROGENTREGAS(i, 17) & vbCrLf
                    sSql = sSql & "      Where " & vbCrLf
                    sSql = sSql & "            SGI_FILIAL = " & intFILIAL & vbCrLf
                    sSql = sSql & "        And SGI_CODIGO = " & Trim(Replace(arrPROGENTREGAS(i, 6), "/", "")) & vbCrLf
                    sSql = sSql & "        And SGI_IDPAI  = " & arrPROGENTREGAS(i, 16)
                    
                    BGRV.CommandText = sSql
                    BGRV.Execute
                    '' --------------------------------------------
                
                End If
            Next i
        End If
     
        '' =====================================
        '' Grava Log
        sSql = ""
        
        sSql = "Insert Into SGI_LOGPED" & strNOTABFILIAL & " (" & vbCrLf
        sSql = sSql & "                              SGI_FILIAL" & vbCrLf
        sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                             ,SGI_DATA" & vbCrLf
        sSql = sSql & "                             ,SGI_HORA" & vbCrLf
        sSql = sSql & "                             ,SGI_CODUSUARIO" & vbCrLf
        sSql = sSql & "                             ,SGI_ACAO" & vbCrLf
        sSql = sSql & "                    ) Values (" & vbCrLf
        sSql = sSql & "                              " & FILIAL & vbCrLf
        sSql = sSql & "                             ," & lngCODPEDIDO & vbCrLf
        sSql = sSql & "                             ,'" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                             ,'" & Format(Now, "HH:MM:SS") & "'" & vbCrLf
        sSql = sSql & "                             ," & lngCodUsuario & vbCrLf
        sSql = sSql & "                             ,'" & strOPER & "'" & vbCrLf
        sSql = sSql & "                             )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' =====================================
     
     ElseIf strOPER = "LV" Or strOPER = "LC" Then
        '' Liberando ( Aguardando Liberação de Fotolito ) = LV
         ' Liberando ( Aguardando Liberação P.Cota/P.Data ) = LC
        '' Programação de Entregas
        If IsArray(arrPROGENTREGAS) Then
            For i = 1 To UBound(arrPROGENTREGAS)
            
                sSql = ""
                
                sSql = "Update SGI_PROGENTRPROD" & strNOTABFILIAL & " Set" & vbCrLf
                sSql = sSql & "                                 SGI_STATUS     = " & arrPROGENTREGAS(i, 17) & vbCrLf '' Status
                sSql = sSql & "                                ,SGI_DATENTREGA = " & arrPROGENTREGAS(i, 2) & vbCrLf
                sSql = sSql & " Where" & vbCrLf
                sSql = sSql & "       SGI_FILIAL    = " & intFILIAL & vbCrLf
                sSql = sSql & "   And SGI_CODPED    = " & lngCODPEDIDO
                sSql = sSql & "   And SGI_IDINTERNO = " & arrPROGENTREGAS(i, 16) & vbCrLf
            
                BGRV.CommandText = sSql
                BGRV.Execute
                
                If Len(Trim(Replace(arrPROGENTREGAS(i, 9), "/", ""))) > 0 Then
                    sSql = ""
                    
                    sSql = "Update SGI_ORDEMPROD" & strNOTABFILIAL & " Set " & vbCrLf
                    sSql = sSql & "                           SGI_DATENTREGA      = " & arrPROGENTREGAS(i, 2) & vbCrLf
                    sSql = sSql & "                          ,SGI_STATUS          = " & arrPROGENTREGAS(i, 17) & vbCrLf '' Status
                    
                    sSql = sSql & "                     Where " & vbCrLf
                    sSql = sSql & "                           SGI_FILIAL    = " & FILIAL & vbCrLf
                    sSql = sSql & "                       And SGI_IDPRODUTO = " & arrPROGENTREGAS(i, 0) & vbCrLf
                    sSql = sSql & "                       And SGI_CODIGO    = " & arrPROGENTREGAS(i, 9) & vbCrLf
                    sSql = sSql & "                       And SGI_IDPAI     = " & arrPROGENTREGAS(i, 16) & vbCrLf
                    
                    BGRV.CommandText = sSql
                    BGRV.Execute
                End If
            Next i
        End If
     
        
        '' Liberando Fotolito
        sSql = "Update SGI_CADPEDVENDH" & strNOTABFILIAL & " Set " & vbCrLf
        sSql = sSql & "                           SGI_STATUS      = '" & strSTATUS & "'" & vbCrLf '' Status
        
        sSql = sSql & "                     Where " & vbCrLf & vbCrLf
        sSql = sSql & "                           SGI_FILIAL = " & intFILIAL & vbCrLf
        sSql = sSql & "                       And SGI_CODIGO = " & lngCODPEDIDO
        
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' =====================================
        '' Grava Log
        sSql = ""
        
        sSql = "Insert Into SGI_LOGPED" & strNOTABFILIAL & " (" & vbCrLf
        sSql = sSql & "                              SGI_FILIAL" & vbCrLf
        sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                             ,SGI_DATA" & vbCrLf
        sSql = sSql & "                             ,SGI_HORA" & vbCrLf
        sSql = sSql & "                             ,SGI_CODUSUARIO" & vbCrLf
        sSql = sSql & "                             ,SGI_ACAO" & vbCrLf
        sSql = sSql & "                    ) Values (" & vbCrLf
        sSql = sSql & "                              " & FILIAL & vbCrLf
        sSql = sSql & "                             ," & lngCODPEDIDO & vbCrLf
        sSql = sSql & "                             ,'" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                             ,'" & Format(Now, "HH:MM:SS") & "'" & vbCrLf
        sSql = sSql & "                             ," & lngCodUsuario & vbCrLf
        sSql = sSql & "                             ,'" & strOPER & "'" & vbCrLf
        sSql = sSql & "                             )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' =====================================
     
     End If
     
     adoBanco_Dados.CommitTrans
     
     GRAVASTEEL = True
     
     Exit Function

err_Trans:
    
    adoBanco_Dados.RollbackTrans

    Dim objErro    As Object
    Set objErro = CreateObject("BLBCWS.clsFuncoes")
    Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, strOPER, sSql)
    Set objErro = Nothing

End Function

Public Function PegaQueryOPDia(strCODLIN As String, strDTENTREGA As String, strNOMFILIAL As String, lngCODPED As Long, lngNECKIN As Long, intHOMOLOGSN As Integer) As Long

        PegaQueryOPDia = 0
        
        If Len(Trim(strCODLIN)) = 0 Then Exit Function
        If Len(Trim(strDTENTREGA)) = 0 Then Exit Function
        
        sSql = ""
        
        sSql = "Select" & vbCrLf
        sSql = sSql & "       Sum(OP.SGI_QTDE) As SGI_QTDE" & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_CADPRODUTO      PROD" & vbCrLf
        sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMFILIAL & " OP" & vbCrLf
        sSql = sSql & "      ,SGI_CADTIPPROD      TP" & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "       PROD.SGI_FILIAL     = " & FILIAL & vbCrLf
        sSql = sSql & "   And PROD.SGI_CODLINPROD = " & strCODLIN & vbCrLf
        ''sSql = sSql & "   And PROD.SGI_NECKIN     = " & lngNECKIN & vbCrLf
        sSql = sSql & "   And OP.SGI_FILIAL       = PROD.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And OP.SGI_IDPRODUTO    = PROD.SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "   And OP.SGI_DATENTREGA   = '" & Format(CDate(strDTENTREGA), "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "   And OP.SGI_STATUS       = 0" & vbCrLf
        sSql = sSql & "   And OP.SGI_CODPED       <> " & lngCODPED & vbCrLf
        sSql = sSql & "   And TP.SGI_FILIAL       = PROD.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And TP.SGI_CODIGO       = PROD.SGI_CODTIPO" & vbCrLf
        sSql = sSql & "   And TP.SGI_HOMOLOGSN    = " & intHOMOLOGSN
        
        BREC8.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC8.EOF() Then
            If Not IsNull(BREC8!SGI_QTDE) Then PegaQueryOPDia = BREC8!SGI_QTDE
        End If
        BREC8.Close
        
End Function

Public Function PegaQueryPedBloq(strCODLIN As String, strDTENTREGA As String, strNOMFILIAL As String, lngCODPED As Long, lngNECKIN As Long, intHOMOLOGSN As Integer) As String

        PegaQueryPedBloq = ""

        If Len(Trim(strCODLIN)) = 0 Then Exit Function
        If Len(Trim(strDTENTREGA)) = 0 Then Exit Function

        sSql = ""
        
        sSql = "Select" & vbCrLf
        sSql = sSql & "       Sum(PROGE.SGI_QTDE) As SGI_QTDE" & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_CADPEDVENDI" & strNOMFILIAL & "  PEDVI" & vbCrLf
        sSql = sSql & "      ,SGI_PROGENTRPROD" & strNOMFILIAL & " PROGE" & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDH" & strNOMFILIAL & " PEDVH" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO PROD" & vbCrLf
        sSql = sSql & "      ,SGI_CADTIPPROD TP" & vbCrLf
        
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "       PEDVI.SGI_FILIAL     = " & FILIAL & vbCrLf
        sSql = sSql & "   And PEDVI.SGI_CODLINPROD = " & Trim(strCODLIN) & vbCrLf
        sSql = sSql & "   And PEDVI.SGI_CODIGO     <> " & lngCODPED & vbCrLf
        sSql = sSql & "   And PROGE.SGI_FILIAL     = PEDVI.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_CODPED     = PEDVI.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_IDPRODUTO  = PEDVI.SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_DATENTREGA = '" & Format(CDate(strDTENTREGA), "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "   And PEDVH.SGI_FILIAL     = PEDVI.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PEDVH.SGI_CODIGO     = PEDVI.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And (PEDVH.SGI_STATUS    = 'B' Or PEDVH.SGI_STATUS = 'N')" & vbCrLf
        
        sSql = sSql & "   And PROD.SGI_FILIAL      = PROGE.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PROD.SGI_IDPRODUTO   = PROGE.SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "   And PROD.SGI_NECKIN      = " & lngNECKIN & vbCrLf
        
        sSql = sSql & "   And TP.SGI_FILIAL        = PROD.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And TP.SGI_CODIGO        = PROD.SGI_CODTIPO" & vbCrLf
        sSql = sSql & "   And TP.SGI_HOMOLOGSN     = " & intHOMOLOGSN
        
        PegaQueryPedBloq = sSql
        
End Function

Public Function PegaPedQueryBloqAlt(strCODLIN As String, strDTENTREGA As String, strNOMFILIAL As String, lngCODPED As Long, lngNECKIN As Long, intHOLOGSN As Integer) As String

        PegaPedQueryBloqAlt = ""

        If Len(Trim(strCODLIN)) = 0 Then Exit Function
        If Len(Trim(strDTENTREGA)) = 0 Then Exit Function

        sSql = ""
        
        sSql = "Select" & vbCrLf
        sSql = sSql & "       Sum(PROGE.SGI_QTDE) As SGI_QTDE" & vbCrLf
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_CADPEDVENDI" & strNOMFILIAL & "  PEDVI" & vbCrLf
        sSql = sSql & "      ,SGI_PROGENTRPROD" & strNOMFILIAL & " PROGE" & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDH" & strNOMFILIAL & " PEDVH" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO PROD" & vbCrLf
        sSql = sSql & "      ,SGI_CADTIPPROD TP" & vbCrLf
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "       PEDVI.SGI_FILIAL     = " & FILIAL & vbCrLf
        sSql = sSql & "   And PEDVI.SGI_CODLINPROD = " & Trim(strCODLIN) & vbCrLf
        sSql = sSql & "   And PEDVI.SGI_CODIGO     <> " & lngCODPED & vbCrLf
        
        sSql = sSql & "   And PROGE.SGI_FILIAL     = PEDVI.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_CODPED     = PEDVI.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_IDPRODUTO  = PEDVI.SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_DATENTREGA = '" & Format(CDate(strDTENTREGA), "MM/DD/YYYY") & "'" & vbCrLf
        
        sSql = sSql & "   And PEDVH.SGI_FILIAL     = PEDVI.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PEDVH.SGI_CODIGO     = PEDVI.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And PEDVH.SGI_STATUS     = 'S'" & vbCrLf
        
        sSql = sSql & "   And PROD.SGI_FILIAL      = PROGE.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PROD.SGI_IDPRODUTO   = PROGE.SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "   And PROD.SGI_NECKIN      = " & lngNECKIN & vbCrLf
        
        sSql = sSql & "   And TP.SGI_FILIAL        = PROD.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And TP.SGI_CODIGO        = PROD.SGI_CODTIPO" & vbCrLf
        sSql = sSql & "   And TP.SGI_HOMOLOGSN     = " & intHOLOGSN
        
        PegaPedQueryBloqAlt = sSql

End Function

Public Function PegaPedQueryBloqFot(strCODLIN As String, strDTENTREGA As String, strNOMFILIAL As String, lngCODPED As Long, lngNECKIN As Long, intHOMOLOGSN As Integer) As String

        PegaPedQueryBloqFot = ""

        If Len(Trim(strCODLIN)) = 0 Then Exit Function
        If Len(Trim(strDTENTREGA)) = 0 Then Exit Function

        sSql = ""
        
        sSql = "Select" & vbCrLf
        sSql = sSql & "       Sum(PROGE.SGI_QTDE) As SGI_QTDE" & vbCrLf
        
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_CADPEDVENDI" & strNOMFILIAL & "  PEDVI" & vbCrLf
        sSql = sSql & "      ,SGI_PROGENTRPROD" & strNOMFILIAL & " PROGE" & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDH" & strNOMFILIAL & "  PEDVH" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO PROD" & vbCrLf
        sSql = sSql & "      ,SGI_CADTIPPROD TP" & vbCrLf
        
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "       PEDVI.SGI_FILIAL     = " & FILIAL & vbCrLf
        sSql = sSql & "   And PEDVI.SGI_CODLINPROD = " & Trim(strCODLIN) & vbCrLf
        sSql = sSql & "   And PEDVI.SGI_CODIGO     <> " & lngCODPED & vbCrLf
        sSql = sSql & "   And PROGE.SGI_FILIAL     = PEDVI.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_CODPED     = PEDVI.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_IDPRODUTO  = PEDVI.SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_DATENTREGA = '" & Format(CDate(strDTENTREGA), "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "   And PEDVH.SGI_FILIAL     = PEDVI.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PEDVH.SGI_CODIGO     = PEDVI.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And PEDVH.SGI_STATUS     = 'V'" & vbCrLf
        
        sSql = sSql & "   And PROD.SGI_FILIAL      = PROGE.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PROD.SGI_IDPRODUTO   = PROGE.SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "   And PROD.SGI_NECKIN      = " & lngNECKIN & vbCrLf
        
        sSql = sSql & "   And TP.SGI_FILIAL        = PROD.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And TP.SGI_CODIGO        = PROD.SGI_CODTIPO" & vbCrLf
        sSql = sSql & "   And TP.SGI_HOMOLOGSN     = " & intHOMOLOGSN
        
        PegaPedQueryBloqFot = sSql
        
        
End Function

Public Function PegaPedQueryBloqPcPd(strCODLIN As String, strDTENTREGA As String, strNOMFILIAL As String, lngCODPED As Long, lngNECKIN As Long, intHOMOLOGSN As Integer) As Long

        PegaPedQueryBloqPcPd = 0

        If Len(Trim(strCODLIN)) = 0 Then Exit Function
        If Len(Trim(strDTENTREGA)) = 0 Then Exit Function

        sSql = ""
        
        sSql = "Select Distinct" & vbCrLf
        sSql = sSql & "       PROGE.SGI_QTDE" & vbCrLf
        
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "       SGI_CADPEDVENDI" & strNOMFILIAL & "  PEDVI" & vbCrLf
        sSql = sSql & "      ,SGI_PROGENTRPROD" & strNOMFILIAL & " PROGE" & vbCrLf
        sSql = sSql & "      ,SGI_CADPEDVENDH" & strNOMFILIAL & "  PEDVH" & vbCrLf
        sSql = sSql & "      ,SGI_CADPRODUTO PROD" & vbCrLf
        sSql = sSql & "      ,SGI_CADTIPPROD TP" & vbCrLf
        
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "       PEDVI.SGI_FILIAL     = " & FILIAL & vbCrLf
        sSql = sSql & "   And PEDVI.SGI_CODLINPROD = " & strCODLIN & vbCrLf
        sSql = sSql & "   And PEDVI.SGI_CODIGO     <> " & lngCODPED & vbCrLf
        sSql = sSql & "   And PROGE.SGI_FILIAL     = PEDVI.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_CODPED     = PEDVI.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_IDPRODUTO  = PEDVI.SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "   And PROGE.SGI_DATENTREGA = '" & Format(CDate(strDTENTREGA), "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "   And PEDVH.SGI_FILIAL     = PEDVI.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PEDVH.SGI_CODIGO     = PEDVI.SGI_CODIGO" & vbCrLf
        sSql = sSql & "   And (PEDVH.SGI_STATUS    = 'C' or PEDVH.SGI_STATUS   = '4')" & vbCrLf
        
        sSql = sSql & "   And PROD.SGI_FILIAL      = PROGE.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And PROD.SGI_IDPRODUTO   = PROGE.SGI_IDPRODUTO" & vbCrLf
        ''sSql = sSql & "   And PROD.SGI_NECKIN      = " & lngNECKIN & vbCrLf
        
        sSql = sSql & "   And TP.SGI_FILIAL        = PROD.SGI_FILIAL" & vbCrLf
        sSql = sSql & "   And TP.SGI_CODIGO        = PROD.SGI_CODTIPO" & vbCrLf
        sSql = sSql & "   And TP.SGI_HOMOLOGSN     = " & intHOMOLOGSN
       
        BREC8.Open sSql, adoBanco_Dados, adOpenDynamic
        Do While Not BREC8.EOF()
            PegaPedQueryBloqPcPd = (PegaPedQueryBloqPcPd + BREC8!SGI_QTDE)
            BREC8.MoveNext
        Loop
        BREC8.Close
        
End Function

Public Function PegaCota(strIDPRODUTO As String, strDTENTREGA As String, strNOMFILIAL As String, intHOMOLOGADO As Integer) As Long

        PegaCota = 0
        
        If Len(Trim(strIDPRODUTO)) = 0 Then Exit Function
        If Len(Trim(strDTENTREGA)) = 0 Then Exit Function
        
        sSql = ""
        
        sSql = "Select" & vbCrLf
        sSql = sSql & "      PMDIA.SGI_SEMANA" & vbCrLf
        sSql = sSql & "     ,PMDIA.SGI_DTSEMANA" & vbCrLf
        sSql = sSql & "     ,PMDIA.SGI_DIASEMANA" & vbCrLf
        sSql = sSql & "     ,Sum(PMDIA.SGI_QTDE) As SGI_QTDE" & vbCrLf
        sSql = sSql & "     ,GRPI.SGI_CODLIN" & vbCrLf
        
        sSql = sSql & "  From" & vbCrLf
        sSql = sSql & "      SGI_CADPRODUTO            PROD" & vbCrLf
        sSql = sSql & "     ,SGI_CADLINHAPRODUTO       LINP" & vbCrLf
        sSql = sSql & "     ,SGI_CADGRUPLINHAIT" & strNOMFILIAL & "  GRPI" & vbCrLf
        sSql = sSql & "     ,SGI_CADPLANMESTRE" & strNOMFILIAL & "   PM" & vbCrLf
        sSql = sSql & "     ,SGI_CADDIASPMSEMANA" & strNOMFILIAL & " PMDIA" & vbCrLf
        
        sSql = sSql & " Where" & vbCrLf
        sSql = sSql & "      PROD.SGI_FILIAL         = " & FILIAL & vbCrLf
        sSql = sSql & "  And PROD.SGI_IDPRODUTO      = " & strIDPRODUTO & vbCrLf
        
        sSql = sSql & "  And LINP.SGI_FILIAL         = PROD.SGI_FILIAL" & vbCrLf
        sSql = sSql & "  And LINP.SGI_CODLIN         = PROD.SGI_CODLINPROD" & vbCrLf
        
        sSql = sSql & "  And GRPI.SGI_FILIAL         = LINP.SGI_FILIAL" & vbCrLf
        sSql = sSql & "  And GRPI.SGI_CODLIN         = LINP.SGI_CODIGO" & vbCrLf
''        sSql = sSql & "  And GRPI.SGI_OPTCOMNECKINSN = PROD.SGI_NECKIN" & vbCrLf
        sSql = sSql & "  And GRPI.SGI_HOMOLOGSN      = " & intHOMOLOGADO & vbCrLf
        
        sSql = sSql & "  And PM.SGI_FILIAL           = GRPI.SGI_FILIAL" & vbCrLf
        sSql = sSql & "  And PM.SGI_CODLINHA         = GRPI.SGI_CODIGO" & vbCrLf
        sSql = sSql & "  And PM.SGI_ATIVO            = 1" & vbCrLf
        sSql = sSql & "  And PM.SGI_MES              = " & Month(CDate(strDTENTREGA)) & vbCrLf
        sSql = sSql & "  And PM.SGI_ANO              = " & Year(CDate(strDTENTREGA)) & vbCrLf
        
        sSql = sSql & "  And PMDIA.SGI_FILIAL        = PM.SGI_FILIAL" & vbCrLf
        sSql = sSql & "  And PMDIA.SGI_CODIGO        = PM.SGI_CODIGO" & vbCrLf
        sSql = sSql & "  And PMDIA.SGI_ATIVO         = 1" & vbCrLf
        sSql = sSql & "  And PMDIA.SGI_DTSEMANA      = '" & Format(CDate(strDTENTREGA), "MM/DD/YYYY") & "'" & vbCrLf
        
        sSql = sSql & "Group By" & vbCrLf
        sSql = sSql & "        PMDIA.SGI_SEMANA" & vbCrLf
        sSql = sSql & "       ,PMDIA.SGI_DTSEMANA" & vbCrLf
        sSql = sSql & "       ,PMDIA.SGI_DIASEMANA" & vbCrLf
        sSql = sSql & "       ,GRPI.SGI_CODLIN" & vbCrLf
        sSql = sSql & "Order By PMDIA.SGI_DTSEMANA"
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC.EOF() Then
           If Not IsNull(BREC!SGI_QTDE) Then PegaCota = BREC!SGI_QTDE
        End If
        BREC.Close

End Function

Public Function DtEntregaLDTIME(strDATPED As String) As Date
    
    Dim lngDIASLIDTIME As Long
    Dim intDOMINGOS    As Integer
    
    DtEntregaLDTIME = Empty

    '' Calcula Lid Time
    lngDIASLIDTIME = 0
    intDOMINGOS = 0
    DtEntregaLDTIME = CDate(strDATPED)
    ''Do While lngDIASLIDTIME <= 20
    ''       lngDIASLIDTIME = lngDIASLIDTIME + 1
    ''    DtEntregaLDTIME = (DtEntregaLDTIME + 1)
    ''Loop

    ''If Weekday(DtEntregaLDTIME) = 1 Then DtEntregaLDTIME = (DtEntregaLDTIME + 20)
    DtEntregaLDTIME = (DtEntregaLDTIME + 27)

End Function

Public Function dtLidTime(mskDTPED As String) As Date

    Dim lngDIASLIDTIME As Long
    
    dtLidTime = Empty

    '' Calcula Lid Time
    lngDIASLIDTIME = 1
    dtLidTime = CDate(mskDTPED)
    Do While lngDIASLIDTIME <= 15
       If Weekday(dtLidTime) <> 1 Then
          lngDIASLIDTIME = lngDIASLIDTIME + 1
       End If
       dtLidTime = (dtLidTime + 1)
    Loop

End Function


Public Function PegaHOMOLOGADO(strIDPRODUTO As String) As Long
    
    PegaHOMOLOGADO = 0
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       SGI_CODTIPO" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_IDPRODUTO = " & strIDPRODUTO
    
    BREC8.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC8.EOF() Then
       If BREC8!SGI_CODTIPO = 2 Then
          PegaHOMOLOGADO = 1
       Else
          PegaHOMOLOGADO = 0
       End If
    End If
    BREC8.Close
    
End Function


Public Function PegaNECKIN(strIDPRODUTO As String) As Long
    
    PegaNECKIN = 0
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       SGI_NECKIN" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_IDPRODUTO = " & strIDPRODUTO
    
    BREC8.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC8.EOF() Then PegaNECKIN = BREC8!SGI_NECKIN
    BREC8.Close
    
End Function


Public Function PegaAtrazados(strCODCAPCIDADE As String, strNOMETAB As String) As Long

    PegaAtrazados = 0

    Dim lngQTDEABERTO As Long
    Dim lngQTDESALDO  As Long
    
    lngQTDEABERTO = 0
    lngQTDESALDO = 0
    
    '' Em aberto
    sSql = ""
    
    sSql = "Select Distinct" & vbCrLf
    sSql = sSql & "       ORDP.SGI_QTDEPED" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO       LINHA" & vbCrLf
    sSql = sSql & "      ,SGI_CADGRUPLINHAIT" & strNOMETAB & "  GRPLI" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO            PROD" & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMETAB & "       ORDP" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH" & strNOMETAB & "   PED" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       LINHA.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And LINHA.SGI_CODLIN    = " & Trim(strCODCAPCIDADE) & vbCrLf
    
    sSql = sSql & "   And GRPLI.SGI_FILIAL    = LINHA.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And GRPLI.SGI_CODLIN    = LINHA.SGI_CODIGO" & vbCrLf
    
    sSql = sSql & "   And PROD.SGI_FILIAL     = LINHA.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PROD.SGI_CODLINPROD = LINHA.SGI_CODLIN" & vbCrLf
    sSql = sSql & "   And PROD.SGI_NECKIN     = GRPLI.SGI_OPTCOMNECKINSN" & vbCrLf
    
    sSql = sSql & "   And ORDP.SGI_FILIAL     = PROD.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And ORDP.SGI_IDPRODUTO  = PROD.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "   And ORDP.SGI_STATUS     = 0" & vbCrLf
    sSql = sSql & "   And ORDP.SGI_DATENTREGA < '" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf
    
    sSql = sSql & "   And PED.SGI_FILIAL      = ORDP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PED.SGI_CODIGO      = ORDP.SGI_CODPED" & vbCrLf
    sSql = sSql & "   And PED.SGI_STATUS      IN('L')"
    

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC.EOF()
        lngQTDEABERTO = (lngQTDEABERTO + BREC!SGI_QTDEPED)
        BREC.MoveNext
    Loop
    BREC.Close

    '' parcial
    sSql = ""
    
    sSql = "Select Distinct" & vbCrLf
    sSql = sSql & "       ORDP.SGI_SALDO" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO       LINHA" & vbCrLf
    sSql = sSql & "      ,SGI_CADGRUPLINHAIT" & strNOMETAB & "  GRPLI" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO            PROD" & vbCrLf
    sSql = sSql & "      ,SGI_ORDEMPROD" & strNOMETAB & "       ORDP" & vbCrLf
    sSql = sSql & "      ,SGI_CADPEDVENDH" & strNOMETAB & "   PED" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       LINHA.SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And LINHA.SGI_CODLIN    = " & Trim(strCODCAPCIDADE) & vbCrLf
    
    sSql = sSql & "   And GRPLI.SGI_FILIAL    = LINHA.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And GRPLI.SGI_CODLIN    = LINHA.SGI_CODIGO" & vbCrLf
    
    sSql = sSql & "   And PROD.SGI_FILIAL     = LINHA.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PROD.SGI_CODLINPROD = LINHA.SGI_CODLIN" & vbCrLf
    sSql = sSql & "   And PROD.SGI_NECKIN     = GRPLI.SGI_OPTCOMNECKINSN" & vbCrLf
    
    sSql = sSql & "   And ORDP.SGI_FILIAL     = PROD.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And ORDP.SGI_IDPRODUTO  = PROD.SGI_IDPRODUTO" & vbCrLf
    sSql = sSql & "   And ORDP.SGI_STATUS     = 1" & vbCrLf
    sSql = sSql & "   And ORDP.SGI_DATENTREGA < '" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf
    sSql = sSql & "   And ORDP.SGI_SALDO      > 0" & vbCrLf

    sSql = sSql & "   And PED.SGI_FILIAL      = ORDP.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And PED.SGI_CODIGO      = ORDP.SGI_CODPED" & vbCrLf
    sSql = sSql & "   And PED.SGI_STATUS      IN('P')"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC.EOF()
        lngQTDESALDO = (lngQTDESALDO + BREC!SGI_SALDO)
        BREC.MoveNext
    Loop
    BREC.Close

    PegaAtrazados = (lngQTDEABERTO + lngQTDESALDO)

End Function

Public Function PegaStatusFotolitoNovo(strIDPRODUTO As String) As Boolean
    
    PegaStatusFotolitoNovo = False
    
    If Len(Trim(strIDPRODUTO)) = 0 Then Exit Function
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       SGI_PRODNOVO" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_IDPRODUTO = " & Trim(strIDPRODUTO) & vbCrLf
    sSql = sSql & "   And SGI_PRODNOVO  = 1 "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PegaStatusFotolitoNovo = True
    BREC.Close
    
End Function

Public Function PegaStatusFotolitoAlteracao(strIDPRODUTO As String) As Boolean
    
    PegaStatusFotolitoAlteracao = False
    
    If Len(Trim(strIDPRODUTO)) = 0 Then Exit Function
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       SGI_PRODNOVO" & vbCrLf
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO" & vbCrLf
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_FILIAL        = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_IDPRODUTO     = " & Trim(strIDPRODUTO) & vbCrLf
    sSql = sSql & "   And SGI_ALTFOTOLIT    = 1 "
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PegaStatusFotolitoAlteracao = True
    BREC.Close
    
End Function


Public Function PermiteLibComercial(lngCodUsuario As Long) As Boolean

    PermiteLibComercial = False
    
    If lngCodUsuario = 0 Then
       PermiteLibComercial = True
       Exit Function
    End If
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_LIBCOMSN" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_USUARIO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO   = " & lngCodUsuario & vbCrLf
    sSql = sSql & "   And SGI_LIBCOMSN = 1"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PermiteLibComercial = True
    BREC.Close

End Function

Public Function PermiteLibFinanceiro(lngCodUsuario As Long) As Boolean

    PermiteLibFinanceiro = False
    
    If lngCodUsuario = 0 Then
       PermiteLibFinanceiro = True
       Exit Function
    End If
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_LIBFINSN" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_USUARIO" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL   = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO   = " & lngCodUsuario & vbCrLf
    sSql = sSql & "   And SGI_LIBFINSN = 1"

    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PermiteLibFinanceiro = True
    BREC.Close

End Function

