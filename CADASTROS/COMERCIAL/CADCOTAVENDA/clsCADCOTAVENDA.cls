VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADCOTAVENDA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial      As Integer
Dim intCODIGO      As Long
Dim dtDTCOTACAO    As Date
Dim intCODCLI      As Integer
Dim intCODVEND     As Integer
Dim intCODCONDPGT  As Integer
Dim intCODTIPORC   As Integer
Dim intPRZENTREGA  As Integer
Dim intVALCOTACAO  As Integer
Dim intIPIISENTO   As Integer
Dim intICMSISENTO  As Integer
Dim intFRETECIFFOB As Integer
Dim curBASEICMS    As Currency
Dim curALIQICMS    As Currency
Dim curVLICMS      As Currency
Dim curTOTOUTDESP  As Currency
Dim curTOTFRETE    As Currency
Dim curVLIPI       As Currency
Dim curVLDESCTO    As Currency
Dim curTOTORCTO    As Currency
Dim arrTIPODESP    As Variant
Dim arrPRODUTOS    As Variant
Dim arrTIPSERV     As Variant
Dim intTOTALITENS  As Integer
Dim strCONTATO     As String
Dim strDEPTO       As String
Dim strEMAIL       As String
Dim strTELCLIE     As String
Dim strTIPODESP    As String
Dim strEXTENSO     As String
Dim dtDATCALIBRAC  As Date
Dim strSTATUS      As String
Dim lngCODPED      As Long
Dim dtDATVALCOTA   As Date
Dim dtDATENTREGA   As Date
Dim curVLPDESCT    As Currency
Dim curVLTOTDESC   As Currency
Dim curVLSERVIC    As Currency
Dim intCODTECNICO  As Integer
Dim lngCODTIPODOC  As Long
Dim lngCODBANCO    As Long

Public Property Get DATENTREGA() As Date
    DATENTREGA = dtDATENTREGA
End Property

Public Property Let DATENTREGA(ByVal vData As Date)
    dtDATENTREGA = vData
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

Public Property Get CODTECNICO() As Integer
    CODTECNICO = intCODTECNICO
End Property

Public Property Let CODTECNICO(ByVal vData As Integer)
    intCODTECNICO = vData
End Property

Public Property Get VLSERVIC() As Currency
    VLSERVIC = curVLSERVIC
End Property

Public Property Let VLSERVIC(ByVal vData As Currency)
    curVLSERVIC = vData
End Property

Public Property Get VLTOTDESC() As Currency
    VLTOTDESC = curVLTOTDESC
End Property

Public Property Let VLTOTDESC(ByVal vData As Currency)
    curVLTOTDESC = vData
End Property

Public Property Get VLPDESCT() As Currency
    VLPDESCT = curVLPDESCT
End Property

Public Property Let VLPDESCT(ByVal vData As Currency)
    curVLPDESCT = vData
End Property

Public Property Get DATVALCOTA() As Date
    DATVALCOTA = dtDATVALCOTA
End Property

Public Property Let DATVALCOTA(ByVal vData As Date)
    dtDATVALCOTA = vData
End Property

Public Property Get CODPED() As Long
    CODPED = lngCODPED
End Property

Public Property Let CODPED(ByVal vData As Long)
    lngCODPED = vData
End Property

Public Property Get STATUS() As String
    STATUS = strSTATUS
End Property

Public Property Let STATUS(ByVal vData As String)
    strSTATUS = vData
End Property

Public Property Get DATCALIBRAC() As Date
    DATCALIBRAC = dtDATCALIBRAC
End Property

Public Property Let DATCALIBRAC(ByVal vData As Date)
    dtDATCALIBRAC = vData
End Property

Public Property Get EXTENSO() As String
    EXTENSO = strEXTENSO
End Property

Public Property Let EXTENSO(ByVal vData As String)
    strEXTENSO = vData
End Property

Public Property Get TELCLIE() As String
    TELCLIE = strTELCLIE
End Property

Public Property Let TELCLIE(ByVal vData As String)
    strTELCLIE = vData
End Property

Public Property Get EMAIL() As String
    EMAIL = strEMAIL
End Property

Public Property Let EMAIL(ByVal vData As String)
    strEMAIL = vData
End Property

Public Property Get DEPTO() As String
    DEPTO = strDEPTO
End Property

Public Property Let DEPTO(ByVal vData As String)
    strDEPTO = vData
End Property

Public Property Get CONTATO() As String
    CONTATO = strCONTATO
End Property

Public Property Let CONTATO(ByVal vData As String)
    strCONTATO = vData
End Property

Public Property Get TOTALITENS() As Integer
    TOTALITENS = intTOTALITENS
End Property

Public Property Let TOTALITENS(ByVal vData As Integer)
    intTOTALITENS = vData
End Property

Public Property Get PRODUTOS() As Variant
    PRODUTOS = arrPRODUTOS
End Property

Public Property Let PRODUTOS(ByVal vData As Variant)
    arrPRODUTOS = vData
End Property

Public Property Get TIPODESP() As Variant
    TIPODESP = arrTIPODESP
End Property

Public Property Let TIPODESP(ByVal vData As Variant)
    arrTIPODESP = vData
End Property

Public Property Get TIPSERV() As Variant
    TIPSERV = arrTIPSERV
End Property

Public Property Let TIPSERV(ByVal vData As Variant)
    arrTIPSERV = vData
End Property

Public Property Get TOTORCTO() As Currency
    TOTORCTO = curTOTORCTO
End Property

Public Property Let TOTORCTO(ByVal vData As Currency)
    curTOTORCTO = vData
End Property

Public Property Get VLDESCTO() As Currency
    VLDESCTO = curVLDESCTO
End Property

Public Property Let VLDESCTO(ByVal vData As Currency)
    curVLDESCTO = vData
End Property

Public Property Get VLIPI() As Currency
    VLIPI = curVLIPI
End Property

Public Property Let VLIPI(ByVal vData As Currency)
    curVLIPI = vData
End Property

Public Property Get TOTFRETE() As Currency
    TOTFRETE = curTOTFRETE
End Property

Public Property Let TOTFRETE(ByVal vData As Currency)
    curTOTFRETE = vData
End Property

Public Property Get TOTOUTDESP() As Currency
    TOTOUTDESP = curTOTOUTDESP
End Property

Public Property Let TOTOUTDESP(ByVal vData As Currency)
    curTOTOUTDESP = vData
End Property

Public Property Get VLICMS() As Currency
    VLICMS = curVLICMS
End Property

Public Property Let VLICMS(ByVal vData As Currency)
    curVLICMS = vData
End Property

Public Property Get ALIQICMS() As Currency
    ALIQICMS = curALIQICMS
End Property

Public Property Let ALIQICMS(ByVal vData As Currency)
    curALIQICMS = vData
End Property

Public Property Get BASEICMS() As Currency
    BASEICMS = curBASEICMS
End Property

Public Property Let BASEICMS(ByVal vData As Currency)
    curBASEICMS = vData
End Property

Public Property Get FRETECIFFOB() As Integer
    FRETECIFFOB = intFRETECIFFOB
End Property

Public Property Let FRETECIFFOB(ByVal vData As Integer)
    intFRETECIFFOB = vData
End Property

Public Property Get ICMSISENTO() As Integer
    ICMSISENTO = intICMSISENTO
End Property

Public Property Let ICMSISENTO(ByVal vData As Integer)
    intICMSISENTO = vData
End Property

Public Property Get IPIISENTO() As Integer
    IPIISENTO = intIPIISENTO
End Property

Public Property Let IPIISENTO(ByVal vData As Integer)
    intIPIISENTO = vData
End Property

Public Property Get VALCOTACAO() As Integer
    VALCOTACAO = intVALCOTACAO
End Property

Public Property Let VALCOTACAO(ByVal vData As Integer)
    intVALCOTACAO = vData
End Property

Public Property Get PRZENTREGA() As Integer
    PRZENTREGA = intPRZENTREGA
End Property

Public Property Let PRZENTREGA(ByVal vData As Integer)
    intPRZENTREGA = vData
End Property

Public Property Get CODTIPORC() As Integer
    CODTIPORC = intCODTIPORC
End Property

Public Property Let CODTIPORC(ByVal vData As Integer)
    intCODTIPORC = vData
End Property

Public Property Get CODCONDPGT() As Integer
    CODCONDPGT = intCODCONDPGT
End Property

Public Property Let CODCONDPGT(ByVal vData As Integer)
    intCODCONDPGT = vData
End Property

Public Property Get CODVEND() As Integer
    CODVEND = intCODVEND
End Property

Public Property Let CODVEND(ByVal vData As Integer)
    intCODVEND = vData
End Property

Public Property Get CODCLI() As Integer
    CODCLI = intCODCLI
End Property

Public Property Let CODCLI(ByVal vData As Integer)
    intCODCLI = vData
End Property

Public Property Get DTCOTACAO() As Date
    DTCOTACAO = dtDTCOTACAO
End Property

Public Property Let DTCOTACAO(ByVal vData As Date)
    dtDTCOTACAO = vData
End Property

Public Property Get CODIGO() As Long
    CODIGO = intCODIGO
End Property

Public Property Let CODIGO(ByVal vData As Long)
    intCODIGO = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUsuario As String, lngIDUsuario As Long)
    frmCADCOTAVENDAP.cCaminho = cCaminho
    frmCADCOTAVENDAP.Linha = Linha
    frmCADCOTAVENDAP.FILIAL = iFilial
    frmCADCOTAVENDAP.strAcesso = strAcesso
    frmCADCOTAVENDAP.strUsuario = strUsuario
    frmCADCOTAVENDAP.Show vbModal
End Sub

Public Function Pesq_CadOrcVenda() As Boolean

   Pesq_CadOrcVenda = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       * " & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_CADCOTAVENDH " & vbCrLf
   sSql = sSql + " Where " & vbCrLf
   sSql = sSql + "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_CadOrcVenda = True
   BREC.Close

End Function


Public Sub PreencheComboCliente(Combo As Variant)

     Combo.Clear
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       * " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADCLIENTE  " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL  = " & intFilial & vbCrLf
     
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
     sSql = sSql & "       PGTO.SGI_FILIAL  = " & intFilial & vbCrLf
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_DESCRICAO)
        Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
        BREC.MoveNext
     Loop
    
     BREC.Close
    
End Sub

Public Sub PreencheComboVendedor(Combo As Variant, strUsuario As String)

     Combo.Clear
     
     Dim I As Integer
     
     If Len(Trim(strUsuario)) = 0 Then Exit Sub
     
    sSql = "Select " & vbCrLf
    sSql = sSql & "       VEN.* " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_USUARIO      USU" & vbCrLf
    sSql = sSql & "      ,SGI_CADVENDEDOR  VEN" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       USU.SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "   And USU.SGI_NOME       = '" & strUsuario & "'" & vbCrLf
    sSql = sSql & "   And VEN.SGI_FILIAL     = USU.SGI_FILIAL" & vbCrLf
    sSql = sSql & "   And VEN.SGI_CODUSUARIO = USU.SGI_CODIGO"
     
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        Do While Not BREC.EOF()
           Combo.AddItem Trim(BREC!SGI_DESCRICAO)
           Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
           BREC.MoveNext
        Loop
        
        
    Else
         sSql = "Select " & vbCrLf
         sSql = sSql & "       * " & vbCrLf
         sSql = sSql & "  From " & vbCrLf
         sSql = sSql & "       SGI_CADVENDEDOR " & vbCrLf
         sSql = sSql & " Where " & vbCrLf
         sSql = sSql & "       SGI_FILIAL  = " & intFilial & vbCrLf
         
         BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
         Do While Not BREC2.EOF()
            Combo.AddItem Trim(BREC2!SGI_DESCRICAO)
            Combo.ItemData(Combo.NewIndex) = BREC2!SGI_CODIGO
            BREC2.MoveNext
         Loop
        
         BREC2.Close
    End If
    BREC.Close
    
End Sub


Public Sub PreencheComboTipoOrc(Combo As Variant)

     Combo.Clear
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       * " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADESPORCA " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL  = " & intFilial & vbCrLf
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_DESCRICAO)
        Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
        BREC.MoveNext
     Loop
    
     BREC.Close
    
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
    sSql = sSql & "       ITENS.SGI_FILIAL = " & intFilial & vbCrLf
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
         
     Dim I      As Integer
     Dim j      As Integer
     Dim sValor As String
     
     '' Inicia transação
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Then '' Inclui
        
ATALHO_ALTERACAO:
       
        '' Grava Cotação HEADER(CABEÇALHO)
        sSql = "Insert Into SGI_CADCOTAVENDH (SGI_FILIAL" & vbCrLf
        sSql = sSql & "                      ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                      ,SGI_DATACOTA" & vbCrLf
        sSql = sSql & "                      ,SGI_CODCLI" & vbCrLf
        sSql = sSql & "                      ,SGI_CODVEND" & vbCrLf
        sSql = sSql & "                      ,SGI_CODCONDPGT" & vbCrLf
        sSql = sSql & "                      ,SGI_CODTIPORC" & vbCrLf
        sSql = sSql & "                      ,SGI_PRZENTR" & vbCrLf
        sSql = sSql & "                      ,SGI_VALPROP" & vbCrLf
        sSql = sSql & "                      ,SGI_IPIISENTO" & vbCrLf
        sSql = sSql & "                      ,SGI_ICMSISENTO" & vbCrLf
        sSql = sSql & "                      ,SGI_TIPFRETE" & vbCrLf
        sSql = sSql & "                      ,SGI_CONTATO" & vbCrLf
        sSql = sSql & "                      ,SGI_DEPTO" & vbCrLf
        sSql = sSql & "                      ,SGI_EMAIL" & vbCrLf
        sSql = sSql & "                      ,SGI_TEL" & vbCrLf
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
        sSql = sSql & "                      ,SGI_ITENSDESPASTEC" & vbCrLf
        sSql = sSql & "                      ,SGI_EXTENSO " & vbCrLf
        sSql = sSql & "                      ,SGI_DTCALIBRA" & vbCrLf
        sSql = sSql & "                      ,SGI_DATVALCOT" & vbCrLf
        sSql = sSql & "                      ,SGI_VLSERVICOS" & vbCrLf
        sSql = sSql & "                      ,SGI_PROCDESCTOT" & vbCrLf
        sSql = sSql & "                      ,SGI_VLDESCTOTAL" & vbCrLf
        sSql = sSql & "                      ,SGI_CODTEC" & vbCrLf
        sSql = sSql & "                      ,SGI_CODTIPODOC" & vbCrLf
        sSql = sSql & "                      ,SGI_CODBANCO" & vbCrLf
        sSql = sSql & "                      ,SGI_DATENTREGA" & vbCrLf
        
        sSql = sSql & "                       ) " & vbCrLf
        sSql = sSql & "               Values (" & vbCrLf
        sSql = sSql & "                       " & intFilial & vbCrLf
        sSql = sSql & "                      ," & intCODIGO & vbCrLf
        sSql = sSql & "                      ,'" & Format(dtDTCOTACAO, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                      ," & intCODCLI & vbCrLf
        sSql = sSql & "                      ," & intCODVEND & vbCrLf
        sSql = sSql & "                      ," & intCODCONDPGT & vbCrLf
        sSql = sSql & "                      ," & intCODTIPORC & vbCrLf
        sSql = sSql & "                      ," & intPRZENTREGA & vbCrLf
        sSql = sSql & "                      ," & intVALCOTACAO & vbCrLf
        sSql = sSql & "                      ," & intIPIISENTO & vbCrLf
        sSql = sSql & "                      ," & intICMSISENTO & vbCrLf
        sSql = sSql & "                      ," & intFRETECIFFOB & vbCrLf
        sSql = sSql & "                      ,'" & strCONTATO & "'" & vbCrLf
        sSql = sSql & "                      ,'" & strDEPTO & "'" & vbCrLf
        sSql = sSql & "                      ,'" & strEMAIL & "'" & vbCrLf
        sSql = sSql & "                      ,'" & strTELCLIE & "'" & vbCrLf
        sSql = sSql & "                      ," & intTOTALITENS & vbCrLf
        
        sValor = Replace(curTOTORCTO, ".", "")
        sValor = Replace(sValor, ",", ".")
        sSql = sSql & "                      ," & sValor & vbCrLf
        
        If curTOTOUTDESP > 0 Then
           sValor = Replace(curTOTOUTDESP, ".", "")
           sValor = Replace(sValor, ",", ".")
           sSql = sSql & "                      ," & sValor & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
        
        If curTOTFRETE > 0 Then
           sValor = Replace(curTOTFRETE, ".", "")
           sValor = Replace(sValor, ",", ".")
           sSql = sSql & "                      ," & sValor & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
        
        If curBASEICMS > 0 Then
           sValor = Replace(curBASEICMS, ".", "")
           sValor = Replace(sValor, ",", ".")
           sSql = sSql & "                      ," & sValor & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
        
        If curALIQICMS > 0 Then
           sValor = Replace(curALIQICMS, ".", "")
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
        
        If Len(Trim(strTIPODESP)) > 0 Then
           sSql = sSql & "                      ,'" & strTIPODESP & "'" & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
        
        If intCODTIPORC = 2 Then
           sSql = sSql & "                      ,'" & strEXTENSO & "'" & vbCrLf
           If IsDate(dtDATCALIBRAC) Then
              If dtDATCALIBRAC <> "00:00:00" Then
                 sSql = sSql & "                      ,'" & Format(dtDATCALIBRAC, "MM/DD/YYYY") & "'" & vbCrLf
              Else
                 sSql = sSql & "                      ,Null" & vbCrLf
              End If
           Else
              sSql = sSql & "                      ,Null" & vbCrLf
           End If
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
        
        sSql = sSql & "                      ,'" & Format(dtDATVALCOTA, "MM/DD/YYYY") & "'" & vbCrLf
        
        If curVLSERVIC > 0 Then
           sValor = Replace(curVLSERVIC, ".", "")
           sValor = Replace(sValor, ",", ".")
           sSql = sSql & "                      ," & sValor & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
        
        If curVLPDESCT > 0 Then
           sValor = Replace(curVLPDESCT, ".", "")
           sValor = Replace(sValor, ",", ".")
           sSql = sSql & "                      ," & sValor & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
        
        If curVLTOTDESC > 0 Then
           sValor = Replace(curVLTOTDESC, ".", "")
           sValor = Replace(sValor, ",", ".")
           sSql = sSql & "                      ," & sValor & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
        
        If intCODTECNICO > 0 Then
           sSql = sSql & "                      ," & intCODTECNICO & vbCrLf
        Else
           sSql = sSql & "                      ,Null" & vbCrLf
        End If
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
        
        sSql = sSql & "                      ,'" & Format(dtDATENTREGA, "MM/DD/YYYY") & "'" & vbCrLf
        
        sSql = sSql & "                           )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Grava os Produtos
        If IsArray(arrPRODUTOS) = True Then
           For I = 1 To UBound(arrPRODUTOS)
                     
               sSql = "Insert Into SGI_CADCOTAVENDI (SGI_FILIAL" & vbCrLf
               sSql = sSql & "                      ,SGI_CODIGO" & vbCrLf
               sSql = sSql & "                      ,SGI_IDPRODUTO" & vbCrLf
               sSql = sSql & "                      ,SGI_CODPROD" & vbCrLf
               sSql = sSql & "                      ,SGI_QTDE" & vbCrLf
               sSql = sSql & "                      ,SGI_VLUNIT" & vbCrLf
               sSql = sSql & "                      ,SGI_PRCDESC" & vbCrLf
               sSql = sSql & "                      ,SGI_PRCIPI" & vbCrLf
               sSql = sSql & "                      ,SGI_VLTOT" & vbCrLf
               sSql = sSql & "                      ,SGI_ESPTEC" & vbCrLf
               sSql = sSql & "                      ,SGI_QTDPED)" & vbCrLf
               sSql = sSql & "               Values (" & vbCrLf
               sSql = sSql & "                       " & intFilial & vbCrLf
               sSql = sSql & "                      ," & intCODIGO & vbCrLf
               sSql = sSql & "                      ," & Trim(arrPRODUTOS(I, 7)) & vbCrLf
               sSql = sSql & "                      ,'" & Trim(arrPRODUTOS(I, 1)) & "'" & vbCrLf
                
               sValor = Replace(arrPRODUTOS(I, 2), ".", "")
               sValor = Replace(sValor, ",", ".")
               sSql = sSql & "                      ," & sValor & vbCrLf
               
               sValor = Replace(arrPRODUTOS(I, 3), ".", "")
               sValor = Replace(sValor, ",", ".")
               sSql = sSql & "                      ," & sValor & vbCrLf
               
               If arrPRODUTOS(I, 4) > 0 Then
                  sValor = Replace(arrPRODUTOS(I, 4), ".", "")
                  sValor = Replace(sValor, ",", ".")
                  sSql = sSql & "                      ," & sValor & vbCrLf
               Else
                  sSql = sSql & "                      ,Null" & vbCrLf
               End If
               
               If arrPRODUTOS(I, 5) > 0 Then
                  sValor = Replace(arrPRODUTOS(I, 5), ".", "")
                  sValor = Replace(sValor, ",", ".")
                  sSql = sSql & "                      ," & sValor & vbCrLf
               Else
                  sSql = sSql & "                      ,Null" & vbCrLf
               End If
               
               sValor = Replace(arrPRODUTOS(I, 6), ".", "")
               sValor = Replace(sValor, ",", ".")
               sSql = sSql & "                      ," & sValor & vbCrLf
               
               sSql = sSql & "                         ,''" & vbCrLf
               sSql = sSql & "                         ,0" & vbCrLf
               sSql = sSql & "                           )"
          
               BGRV.CommandText = sSql
               BGRV.Execute
           
           Next I
        End If
        
     ElseIf strOPER = "A" Then '' Altera
     
         '' Apagando tipo itens ta cotação
         sSql = "Delete from SGI_CADCOTAVENDI " & vbCrLf
         sSql = sSql & "        Where SGI_FILIAL = " & intFilial & vbCrLf
         sSql = sSql & "          And SGI_CODIGO = " & intCODIGO
         
         BGRV.CommandText = sSql
         BGRV.Execute
     
         '' Apagando tipo Cabeçalho da cotação
         sSql = "Delete from SGI_CADCOTAVENDH " & vbCrLf
         sSql = sSql & "        Where SGI_FILIAL = " & intFilial & vbCrLf
         sSql = sSql & "          And SGI_CODIGO = " & intCODIGO
         
         BGRV.CommandText = sSql
         BGRV.Execute
     
         GoTo ATALHO_ALTERACAO
     
     ElseIf strOPER = "E" Then '' Exclui
     
     
     
         '' Apagando tipo itens ta cotação
         sSql = "Delete from SGI_CADCOTAVENDI " & vbCrLf
         sSql = sSql & "        Where SGI_FILIAL = " & intFilial & vbCrLf
         sSql = sSql & "          And SGI_CODIGO = " & intCODIGO
         
         BGRV.CommandText = sSql
         BGRV.Execute
     
         '' Apagando tipo Cabeçalho da cotação
         sSql = "Delete From SGI_CADCOTAVENDH " & vbCrLf
         sSql = sSql & "        Where SGI_FILIAL = " & intFilial & vbCrLf
         sSql = sSql & "          And SGI_CODIGO = " & intCODIGO
         
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

Public Function Carrega_Campos() As Boolean

    Carrega_Campos = False

    Dim I         As Integer
    Dim intQTDREG As Integer
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADCOTAVENDH " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & intCODIGO
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       BREC.Close
       Exit Function
    End If
    
    intCODIGO = BREC!SGI_CODIGO
    DTCOTACAO = BREC!SGI_DATACOTA
    intCODCLI = BREC!SGI_CODCLI
    intCODVEND = BREC!SGI_CODVEND
    intCODCONDPGT = BREC!SGI_CODCONDPGT
    intCODTIPORC = BREC!SGI_CODTIPORC
    intPRZENTREGA = BREC!SGI_PRZENTR
    intVALCOTACAO = BREC!SGI_VALPROP
    intIPIISENTO = BREC!SGI_IPIISENTO
    intICMSISENTO = BREC!SGI_ICMSISENTO
    intFRETECIFFOB = BREC!SGI_TIPFRETE
    strSTATUS = BREC!SGI_STATUS
    dtDATVALCOTA = BREC!SGI_DATVALCOT
    dtDATENTREGA = BREC!SGI_DATENTREGA
    If Not IsNull(BREC!SGI_CODTEC) Then intCODTECNICO = BREC!SGI_CODTEC
    
    If Not IsNull(BREC!SGI_CODPED) Then lngCODPED = BREC!SGI_CODPED
    
    If Not IsNull(BREC!SGI_DTCALIBRA) Then dtDATCALIBRAC = BREC!SGI_DTCALIBRA
    If Not IsNull(BREC!SGI_CONTATO) Then strCONTATO = BREC!SGI_CONTATO
    If Not IsNull(BREC!SGI_DEPTO) Then strDEPTO = BREC!SGI_DEPTO
    If Not IsNull(BREC!SGI_EMAIL) Then strEMAIL = BREC!SGI_EMAIL
    If Not IsNull(BREC!SGI_TEL) Then strTELCLIE = BREC!SGI_TEL
    
    If Not IsNull(BREC!SGI_BASEICMS) Then curBASEICMS = BREC!SGI_BASEICMS
    If Not IsNull(BREC!SGI_PORCICMS) Then curALIQICMS = BREC!SGI_PORCICMS
    If Not IsNull(BREC!SGI_VLICMS) Then curVLICMS = BREC!SGI_VLICMS
    If Not IsNull(BREC!SGI_OUTRASTOT) Then curTOTOUTDESP = BREC!SGI_OUTRASTOT
    If Not IsNull(BREC!SGI_FRETETOT) Then curTOTFRETE = BREC!SGI_FRETETOT
    If Not IsNull(BREC!SGI_VLIPI) Then curVLIPI = BREC!SGI_VLIPI
    If Not IsNull(BREC!SGI_VLDESCTO) Then curVLDESCTO = BREC!SGI_VLDESCTO
    curTOTORCTO = BREC!SGI_VLTOT
    
    If Not IsNull(BREC!SGI_VLSERVICOS) Then curVLSERVIC = BREC!SGI_VLSERVICOS
    If Not IsNull(BREC!SGI_PROCDESCTOT) Then curVLPDESCT = BREC!SGI_PROCDESCTOT
    If Not IsNull(BREC!SGI_VLDESCTOTAL) Then curVLTOTDESC = BREC!SGI_VLDESCTOTAL
    
    If Not IsNull(BREC!SGI_CODTIPODOC) Then lngCODTIPODOC = BREC!SGI_CODTIPODOC
    If Not IsNull(BREC!SGI_CODBANCO) Then lngCODBANCO = BREC!SGI_CODBANCO
    
    BREC.Close
        
    '' ---------------------------------------------------------
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADCOTAVENDI " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & intCODIGO
    
    BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC3.EOF Then
    
       intQTDREG = 0
       Do While Not BREC3.EOF
          intQTDREG = intQTDREG + 1
          BREC3.MoveNext
       Loop
    
       ReDim arrPRODUTOS(1 To intQTDREG, 1 To 8) As Variant
       
       intQTDREG = 1
       BREC3.MoveFirst
       
       Do While Not BREC3.EOF
          
          arrPRODUTOS(intQTDREG, 1) = BREC3!SGI_CODPROD
          arrPRODUTOS(intQTDREG, 2) = BREC3!SGI_QTDE
          arrPRODUTOS(intQTDREG, 3) = BREC3!SGI_VLUNIT
          arrPRODUTOS(intQTDREG, 4) = BREC3!SGI_PRCDESC
          arrPRODUTOS(intQTDREG, 5) = BREC3!SGI_PRCIPI
          arrPRODUTOS(intQTDREG, 6) = BREC3!SGI_VLTOT
          arrPRODUTOS(intQTDREG, 7) = BREC3!SGI_IDPRODUTO
          If Not IsNull(BREC3!SGI_QTDPED) Then
             arrPRODUTOS(intQTDREG, 8) = Format(BREC3!SGI_QTDPED, "#,##0.00")
          Else
             arrPRODUTOS(intQTDREG, 8) = Format(0, "#,##0.00")
          End If
          
          intQTDREG = intQTDREG + 1
          
          BREC3.MoveNext
       Loop
    
    End If
    BREC3.Close
    
    Carrega_Campos = True
    ' ----------------------------------------------------------
    
End Function

Public Sub PreencheComboContato(Combo As Variant)

     Combo.Clear
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       *" & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_DADOSGERAIS " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL   = " & intFilial & vbCrLf
     sSql = sSql & "   And SGI_CODIGO   = " & intCODCLI & vbCrLf
     sSql = sSql & "   And SGI_MODULO   = 'frmCADCLIENTE'" & vbCrLf
     sSql = sSql & "   And SGI_CONTROLE = 'cboCONTNORM'" & vbCrLf
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_TEXTO)
        BREC.MoveNext
     Loop
    
     BREC.Close
    
End Sub

Public Sub PreencheComboEmail(Combo As Variant)

     Combo.Clear
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       *" & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_DADOSGERAIS " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL   = " & intFilial & vbCrLf
     sSql = sSql & "   And SGI_CODIGO   = " & intCODCLI & vbCrLf
     sSql = sSql & "   And SGI_MODULO   = 'frmCADCLIENTE'" & vbCrLf
     sSql = sSql & "   And SGI_CONTROLE = 'cboEMAILNORM'" & vbCrLf
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_TEXTO)
        BREC.MoveNext
     Loop
    
     BREC.Close
    
End Sub

Public Sub PreencheComboTelefone(Combo As Variant)

     Combo.Clear
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       *" & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_DADOSGERAIS " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL   = " & intFilial & vbCrLf
     sSql = sSql & "   And SGI_CODIGO   = " & intCODCLI & vbCrLf
     sSql = sSql & "   And SGI_MODULO   = 'frmCADCLIENTE'" & vbCrLf
     sSql = sSql & "   And SGI_CONTROLE = 'cboTELNORM'" & vbCrLf
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_TEXTO)
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
     sSql = sSql & "       SGI_SERVPROD.SGI_FILIAL     = " & intFilial & vbCrLf
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

Public Sub PreencheComboTecnico(Combo As Variant)

     Combo.Clear
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       TEC.*" & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADTECNICO TEC" & vbCrLf
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_NOME)
        Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
        BREC.MoveNext
     Loop
    
     BREC.Close
    
End Sub

Public Sub PreencheCombo(Combo As Variant, strTabela As String)

     Combo.Clear
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       * " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       " & strTabela & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL  = " & intFilial & vbCrLf
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
    sSql = sSql & "      SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "Order by SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       Combo.AddItem Trim(BREC!SGI_AGENCIA) & Space(10 - Len(Trim(BREC!SGI_AGENCIA))) & " - " & Trim(BREC!SGI_CC) & Space(20 - Len(Trim(BREC!SGI_CC))) & " - " & Trim(BREC!SGI_DESCRICAO)
       Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
       BREC.MoveNext
    Loop
    
    BREC.Close

End Sub


