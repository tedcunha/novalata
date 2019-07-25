VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADCONFFAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial               As Integer
Dim lngCODCONF              As Long
Dim lngCODORD               As Long
Dim lngCODPED               As Long
Dim dtDATACONF              As Date
Dim strOBS                  As String
Dim curBASEICMS             As Currency
Dim curALIQICMS             As Currency
Dim curVALOICMS             As Currency
Dim curOUTRASDESP           As Currency
Dim curFRETE                As Currency
Dim curVALORIPI             As Currency
Dim curPORCDESCTO           As Currency
Dim curVALORDESCT           As Currency
Dim curVLTOTALFAT           As Currency
Dim arrITENSFAT             As Variant
Dim curQTDETOTALFAT         As Long
Dim curQTDEFATCONF          As Currency
Dim curQTDETOTALPED         As Currency
Dim curQTDEATENDPED         As Currency
Dim sValor                  As String
Dim lngCODFAT               As Long
Dim boolSALDOFECHADO        As Boolean
Dim intFILIALEMP            As Integer
Dim strSTATUSPED            As String
Dim lngCODCLIE              As Long
Dim arrARR_OPSELECIONADA    As Variant
Dim lngCodUsuario           As Long

Public Property Get CODUSUARIO() As Long
    CODUSUARIO = lngCodUsuario
End Property

Public Property Let CODUSUARIO(ByVal vData As Long)
    lngCodUsuario = vData
End Property


Public Property Get arrARROPSELECIONADA() As Variant
    arrARROPSELECIONADA = arrARR_OPSELECIONADA
End Property

Public Property Let arrARROPSELECIONADA(ByVal vData As Variant)
    arrARR_OPSELECIONADA = vData
End Property

Public Property Get CODCLIE() As Long
    CODCLIE = lngCODCLIE
End Property

Public Property Let CODCLIE(ByVal vData As Long)
    lngCODCLIE = vData
End Property

Public Property Get STATUSPED() As String
    STATUSPED = strSTATUSPED
End Property

Public Property Let STATUSPED(ByVal vData As String)
    strSTATUSPED = vData
End Property

Public Property Get FILIALEMP() As Integer
    FILIALEMP = intFILIALEMP
End Property

Public Property Let FILIALEMP(ByVal vData As Integer)
    intFILIALEMP = vData
End Property

Public Property Get QTDEFATCONF() As Currency
    QTDEFATCONF = curQTDEFATCONF
End Property

Public Property Let QTDEFATCONF(ByVal vData As Currency)
    curQTDEFATCONF = vData
End Property

Public Property Get SALDOFECHADO() As Boolean
    SALDOFECHADO = boolSALDOFECHADO
End Property

Public Property Let SALDOFECHADO(ByVal vData As Boolean)
    boolSALDOFECHADO = vData
End Property

Public Property Get CODPED() As Long
    CODPED = lngCODPED
End Property

Public Property Let CODPED(ByVal vData As Long)
    lngCODPED = vData
End Property

Public Property Get CODFAT() As Long
    CODFAT = lngCODFAT
End Property

Public Property Let CODFAT(ByVal vData As Long)
    lngCODFAT = vData
End Property

Public Property Get QTDEATENDPED() As Currency
    QTDEATENDPED = curQTDEATENDPED
End Property

Public Property Let QTDEATENDPED(ByVal vData As Currency)
    curQTDEATENDPED = vData
End Property

Public Property Get QTDETOTALPED() As Currency
    QTDETOTALPED = curQTDETOTALPED
End Property

Public Property Let QTDETOTALPED(ByVal vData As Currency)
    curQTDETOTALPED = vData
End Property

Public Property Get QTDETOTALFAT() As Long
    QTDETOTALFAT = curQTDETOTALFAT
End Property

Public Property Let QTDETOTALFAT(ByVal vData As Long)
    curQTDETOTALFAT = vData
End Property

Public Property Get ITENSFAT() As Variant
    ITENSFAT = arrITENSFAT
End Property

Public Property Let ITENSFAT(ByVal vData As Variant)
    arrITENSFAT = vData
End Property

Public Property Get VLTOTALFAT() As Currency
    VLTOTALFAT = curVLTOTALFAT
End Property

Public Property Let VLTOTALFAT(ByVal vData As Currency)
    curVLTOTALFAT = vData
End Property

Public Property Get VALORDESCT() As Currency
    VALORDESCT = curVALORDESCT
End Property

Public Property Let VALORDESCT(ByVal vData As Currency)
    curVALORDESCT = vData
End Property

Public Property Get PORCDESCTO() As Currency
    PORCDESCTO = curPORCDESCTO
End Property

Public Property Let PORCDESCTO(ByVal vData As Currency)
    curPORCDESCTO = vData
End Property

Public Property Get VALORIPI() As Currency
    VALORIPI = curVALORIPI
End Property

Public Property Let VALORIPI(ByVal vData As Currency)
    curVALORIPI = vData
End Property

Public Property Get FRETE() As Currency
    FRETE = curFRETE
End Property

Public Property Let FRETE(ByVal vData As Currency)
    curFRETE = vData
End Property

Public Property Get OUTRASDESP() As Currency
    OUTRASDESP = curOUTRASDESP
End Property

Public Property Let OUTRASDESP(ByVal vData As Currency)
    curOUTRASDESP = vData
End Property

Public Property Get VALOICMS() As Currency
    VALOICMS = curVALOICMS
End Property

Public Property Let VALOICMS(ByVal vData As Currency)
    curVALOICMS = vData
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

Public Property Get OBS() As String
    OBS = strOBS
End Property

Public Property Let OBS(ByVal vData As String)
    strOBS = vData
End Property

Public Property Get CODORD() As Long
    CODORD = lngCODORD
End Property

Public Property Let CODORD(ByVal vData As Long)
    lngCODORD = vData
End Property

Public Property Get DATACONF() As Date
    DATACONF = dtDATACONF
End Property

Public Property Let DATACONF(ByVal vData As Date)
    dtDATACONF = vData
End Property

Public Property Get CODCONF() As Long
    CODCONF = lngCODCONF
End Property

Public Property Let CODCONF(ByVal vData As Long)
    lngCODCONF = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUsuario As String, lngIDUsuario As Long, Optional intFILIAL_PED As Integer)
    frmCADCONFFATP.cCaminho = cCaminho
    frmCADCONFFATP.Linha = Linha
    frmCADCONFFATP.FILIAL = iFilial
    frmCADCONFFATP.strAcesso = strAcesso
    frmCADCONFFATP.strUsuario = strUsuario
    frmCADCONFFATP.lngCodUsuaro = lngIDUsuario
    frmCADCONFFATP.intFILIALPED = intFILIAL_PED
    frmCADCONFFAT.boolSomenteCons = False
    frmCADCONFFATP.Show vbModal
End Sub

Public Sub cConnectPesq(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUsuario As String, lngIDUsuario As Long, Optional intFILIAL_PED As Integer, Optional boolSomenteCons As Boolean)
    frmCADCONFFAT.cCaminho = cCaminho
    frmCADCONFFAT.Linha = Linha
    frmCADCONFFAT.iCodigo = lngCODCONF
    frmCADCONFFAT.cTipOper = "C"
    frmCADCONFFAT.FILIAL = iFilial
    frmCADCONFFAT.strAcesso = strAcesso
    frmCADCONFFAT.strMODPAI = "frmCADCONFFATP"
    frmCADCONFFAT.strUsuario = strUsuario
    frmCADCONFFAT.lngCodVendedor = 0
    frmCADCONFFAT.lngCodUsuario = lngIDUsuario
    frmCADCONFFAT.intFILIALPED = intFILIAL_PED
    frmCADCONFFAT.boolSomenteCons = boolSomenteCons
    frmCADCONFFAT.Show vbModal
End Sub


Public Function AtivoDesativo(strNOMTABELA As String) As Boolean

   AtivoDesativo = False
   
   If BREC.State = 1 Then BREC.Close
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       * " & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       " & strNOMTABELA & " " & vbCrLf
   sSql = sSql + " Where " & vbCrLf
   sSql = sSql + "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then AtivoDesativo = True
   BREC.Close

End Function



Public Function GRAVA(strOPER As String, intFILIALPED As Integer) As Boolean

On Error GoTo err_Trans
     
     GRAVA = False
     
     Dim I            As Integer
     Dim strNOMFILIAL As String
     
     strNOMFILIAL = ""
     If intFILIALPED = 1 Then strNOMFILIAL = "_STEEL"
     
     '' Inicia transação
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
      
     If strOPER = "I" Then '' Inclui
        
        '' Grava Confirmação HEADER(CABEÇALHO)
        sSql = "Insert Into SGI_CADORDCONFH" & Trim(strNOMFILIAL) & " (SGI_FILIAL" & vbCrLf
        sSql = sSql & "                      ,SGI_CODCONF" & vbCrLf
        sSql = sSql & "                      ,SGI_CODORD" & vbCrLf
        sSql = sSql & "                      ,SGI_DATACONF" & vbCrLf
        sSql = sSql & "                      ,SGI_OBS" & vbCrLf
        sSql = sSql & "                      ,SGI_CODFATURA" & vbCrLf
        sSql = sSql & "                      ,SGI_QTDETOTFAT" & vbCrLf
        sSql = sSql & "                      ,SGI_VALTOTFAT" & vbCrLf
        sSql = sSql & "                      ,SGI_BASEFAT" & vbCrLf
        sSql = sSql & "                      ,SGI_ALIQICMS" & vbCrLf
        sSql = sSql & "                      ,SGI_VALOICMS" & vbCrLf
        sSql = sSql & "                      ,SGI_OUTRDESP" & vbCrLf
        sSql = sSql & "                      ,SGI_FRETE" & vbCrLf
        sSql = sSql & "                      ,SGI_TOTALIPI" & vbCrLf
        sSql = sSql & "                      ,SGI_PERCDESC" & vbCrLf
        sSql = sSql & "                      ,SGI_VALODESC" & vbCrLf
        sSql = sSql & "                      ,SGI_TOTALFAT" & vbCrLf
        sSql = sSql & "                      ,SGI_FILIALPED" & vbCrLf
        
        sSql = sSql & "                      )" & vbCrLf
        
        sSql = sSql & "               Values (" & vbCrLf
        sSql = sSql & "                       " & intFilial & vbCrLf
        sSql = sSql & "                      ," & lngCODCONF & vbCrLf
        sSql = sSql & "                      ," & lngCODORD & vbCrLf
        sSql = sSql & "                      ,'" & Format(dtDATACONF, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                      ,'" & strOBS & "'" & vbCrLf
        sSql = sSql & "                      ," & lngCODFAT & vbCrLf
        
        sValor = Replace(curQTDEFATCONF, ".", "")
        sValor = Replace(sValor, ",", ".")
        sSql = sSql & "                      ," & sValor & vbCrLf
        
        sValor = Replace(curVLTOTALFAT, ".", "")
        sValor = Replace(sValor, ",", ".")
        sSql = sSql & "                      ," & sValor & vbCrLf
        
        sValor = Replace(curBASEICMS, ".", "")
        sValor = Replace(sValor, ",", ".")
        sSql = sSql & "                      ," & sValor & vbCrLf
        
        sValor = "Null"
        If curALIQICMS > 0 Then
            sValor = Replace(curALIQICMS, ".", "")
            sValor = Replace(sValor, ",", ".")
        End If
        sSql = sSql & "                      ," & sValor & vbCrLf
        
        sValor = "Null"
        If curVALOICMS > 0 Then
            sValor = Replace(curVALOICMS, ".", "")
            sValor = Replace(sValor, ",", ".")
        End If
        sSql = sSql & "                      ," & sValor & vbCrLf
        
        sValor = "Null"
        If curOUTRASDESP > 0 Then
            sValor = Replace(curOUTRASDESP, ".", "")
            sValor = Replace(sValor, ",", ".")
        End If
        sSql = sSql & "                      ," & sValor & vbCrLf
        
        sValor = "Null"
        If curFRETE > 0 Then
            sValor = Replace(curFRETE, ".", "")
            sValor = Replace(sValor, ",", ".")
        End If
        sSql = sSql & "                      ," & sValor & vbCrLf
        
        sValor = "Null"
        If curVALORIPI > 0 Then
            sValor = Replace(curVALORIPI, ".", "")
            sValor = Replace(sValor, ",", ".")
        End If
        sSql = sSql & "                      ," & sValor & vbCrLf
        
        sValor = "Null"
        If curPORCDESCTO > 0 Then
            sValor = Replace(curPORCDESCTO, ".", "")
            sValor = Replace(sValor, ",", ".")
        End If
        sSql = sSql & "                      ," & sValor & vbCrLf
        
        sValor = "Null"
        If curVALORDESCT > 0 Then
            sValor = Replace(curVALORDESCT, ".", "")
            sValor = Replace(sValor, ",", ".")
        End If
        sSql = sSql & "                      ," & sValor & vbCrLf
        
        sValor = Replace(curVLTOTALFAT, ".", "")
        sValor = Replace(sValor, ",", ".")
        sSql = sSql & "                      ," & sValor & vbCrLf
        sSql = sSql & "                      ," & intFILIALEMP & vbCrLf
        
        sSql = sSql & "                           )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
     
        '' Gravando Itens da Confirmação
        If IsArray(arrITENSFAT) Then
            For I = 1 To UBound(arrITENSFAT)
                
                sSql = "Insert Into SGI_CADORDCONFI" & Trim(strNOMFILIAL) & "  (" & vbCrLf
                
                sSql = sSql & "                       SGI_FILIAL" & vbCrLf
                sSql = sSql & "                      ,SGI_CODCONF" & vbCrLf
                sSql = sSql & "                      ,SGI_INDICE" & vbCrLf
                sSql = sSql & "                      ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                      ,SGI_CODPRODUTO" & vbCrLf
                sSql = sSql & "                      ,SGI_QTDREAL" & vbCrLf
                sSql = sSql & "                      ,SGI_PORCIPI" & vbCrLf
                sSql = sSql & "                      ,SGI_VLUNIT" & vbCrLf
                sSql = sSql & "                      ,SGI_VLTOTAL" & vbCrLf
                sSql = sSql & "                      ,SGI_VLDOPI" & vbCrLf
                sSql = sSql & "                      ,SGI_CODORDPROD" & vbCrLf
                sSql = sSql & "                      ,SGI_CODFORN" & vbCrLf
                sSql = sSql & "                      ,SGI_FILIALPED" & vbCrLf
                
                sSql = sSql & "                     ) Values (" & vbCrLf
                
                sSql = sSql & "                       " & intFilial & vbCrLf
                sSql = sSql & "                      ," & lngCODCONF & vbCrLf
                sSql = sSql & "                      ," & I & vbCrLf
                sSql = sSql & "                      ," & arrITENSFAT(I, 1) & vbCrLf
                sSql = sSql & "                      ,'" & Trim(arrITENSFAT(I, 2)) & "'" & vbCrLf
                
                sValor = Replace(arrITENSFAT(I, 3), ".", "")
                sValor = Replace(sValor, ",", ".")
                sSql = sSql & "                      ," & sValor & vbCrLf
                
                sValor = "Null"
                If Len(Trim(arrITENSFAT(I, 4))) > 0 And arrITENSFAT(I, 4) > 0 Then
                    sValor = Replace(arrITENSFAT(I, 4), ".", "")
                    sValor = Replace(sValor, ",", ".")
                End If
                sSql = sSql & "                      ," & sValor & vbCrLf
                
                sValor = "Null"
                If Len(Trim(arrITENSFAT(I, 5))) > 0 And arrITENSFAT(I, 5) > 0 Then
                    sValor = Replace(arrITENSFAT(I, 5), ".", "")
                    sValor = Replace(sValor, ",", ".")
                End If
                sSql = sSql & "                      ," & sValor & vbCrLf
                
                sValor = "Null"
                If Len(Trim(arrITENSFAT(I, 6))) > 0 And arrITENSFAT(I, 6) > 0 Then
                    sValor = Replace(arrITENSFAT(I, 6), ".", "")
                    sValor = Replace(sValor, ",", ".")
                End If
                sSql = sSql & "                      ," & sValor & vbCrLf
                
                sValor = "Null"
                If Len(Trim(arrITENSFAT(I, 7))) > 0 And arrITENSFAT(I, 7) > 0 Then
                    sValor = Replace(arrITENSFAT(I, 7), ".", "")
                    sValor = Replace(sValor, ",", ".")
                End If
                sSql = sSql & "                      ," & sValor & vbCrLf
                sSql = sSql & "                      ," & arrITENSFAT(I, 9) & vbCrLf
                sSql = sSql & "                      ," & arrITENSFAT(I, 14) & vbCrLf
                sSql = sSql & "                      ," & intFILIALEMP & vbCrLf
                
                sSql = sSql & "                    )"
                
                
                BGRV.CommandText = sSql
                BGRV.Execute
                
                '' ----------------------------------------
                '' Dando Baixa Na Ordem Baixa
                sSql = "Update SGI_ORDEMPROD" & Trim(strNOMFILIAL) & " Set " & vbCrLf
                
                sValor = Replace(arrITENSFAT(I, 10), ".", "")
                sValor = Replace(sValor, ",", ".")
                sSql = sSql & "                  SGI_QTDFAT = " & sValor & vbCrLf
                
                sValor = Replace(arrITENSFAT(I, 12), ".", "")
                sValor = Replace(sValor, ",", ".")
                sSql = sSql & "                 ,SGI_SALDO  = " & sValor & vbCrLf
                
                sSql = sSql & "                 ,SGI_STATUS = " & arrITENSFAT(I, 13) & vbCrLf
                
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       SGI_FILIAL    = " & intFilial & vbCrLf
                sSql = sSql & "   And SGI_CODIGO    = " & Trim(arrITENSFAT(I, 9)) & vbCrLf
                sSql = sSql & "   And SGI_IDPRODUTO = " & Trim(arrITENSFAT(I, 1))
            
                BGRV.CommandText = sSql
                BGRV.Execute
                '' ----------------------------------------
            
                '' =====================================
                '' Grava Log Pedido
                sSql = ""
                
                sSql = "Insert Into SGI_LOGPED" & Trim(strNOMFILIAL) & " (" & vbCrLf
                sSql = sSql & "                              SGI_FILIAL" & vbCrLf
                sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                             ,SGI_DATA" & vbCrLf
                sSql = sSql & "                             ,SGI_HORA" & vbCrLf
                sSql = sSql & "                             ,SGI_CODUSUARIO" & vbCrLf
                sSql = sSql & "                             ,SGI_ACAO" & vbCrLf
                sSql = sSql & "                             ,SGI_DESCACAO" & vbCrLf
                
                sSql = sSql & "                    ) Values (" & vbCrLf
                sSql = sSql & "                              " & FILIAL & vbCrLf
                sSql = sSql & "                             ," & lngCODPED & vbCrLf
                sSql = sSql & "                             ,'" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf
                sSql = sSql & "                             ,'" & Format(Now, "HH:MM:SS") & "'" & vbCrLf
                sSql = sSql & "                             ," & lngCodUsuario & vbCrLf
                sSql = sSql & "                             ,'CF'" & vbCrLf
                sSql = sSql & "                             ,'" & arrITENSFAT(I, 9) & "'" & vbCrLf
                sSql = sSql & "                             )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
                '' =====================================
            
            
            Next I
        End If
        
        '' =====================================================================
        '' Atualizando Ordem de Faturamento
        sSql = "Update SGI_CADORDFATH" & Trim(strNOMFILIAL) & " Set SGI_STATUS = 1" & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_CODORD = " & lngCODORD
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' =====================================================================
        
        '' =====================================================================
        '' Fechando o Pedido Caso ja Seja Saldo 0
        sSql = "Update SGI_CADPEDVENDH" & Trim(strNOMFILIAL) & " Set " & vbCrLf
        sSql = sSql & "                       SGI_STATUS = " & strSTATUSPED & vbCrLf
        
        sValor = Replace(curQTDETOTALFAT, ".", "")
        sValor = Replace(sValor, ",", ".")
        sSql = sSql & "                          ,SGI_QTDEITENSFATURADOS = " & Trim(sValor) & vbCrLf
        
        sSql = sSql & "                 Where " & vbCrLf
        sSql = sSql & "                       SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "                   And SGI_CODIGO = " & lngCODPED
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' =====================================================================
     
        '' Atualizando o Cliente com a Data de Faturamento
        sSql = ""
        
        If Len(Trim(strNOMFILIAL)) > 0 Then
            sSql = "Update SGI_CADCLIENTE Set SGI_DATULTFATSTEEL = '" & Format(dtDATACONF, "MM/DD/YYYY") & "'" & vbCrLf
        Else
            sSql = "Update SGI_CADCLIENTE Set SGI_DATULTFATNOVA  = '" & Format(dtDATACONF, "MM/DD/YYYY") & "'" & vbCrLf
        End If
        
        sSql = sSql & "Where " & vbCrLf
        sSql = sSql & "      SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "  And SGI_CODIGO = " & lngCODCLIE
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' =====================================================================
     
        '' OP's Selecionadas Quando O Pedido Foi Totalmente Fechado
        ''If IsArray(arrARROPSELECIONADA) Then
        
        ''    For I = 1 To UBound(arrARROPSELECIONADA)
            
        ''        sSql = ""
                
        ''        sSql = "Update SGI_ORDEMPROD" & Trim(strNOMFILIAL) & " Set SGI_STATUS = 2" & vbCrLf
        ''        sSql = sSql & "Where" & vbCrLf
        ''        sSql = sSql & "      SGI_FILIAL = " & FILIAL & vbCrLf
        ''        sSql = sSql & "  And SGI_CODIGO = " & arrARROPSELECIONADA(I, 1)
            
        ''        BGRV.CommandText = sSql
        ''        BGRV.Execute
            
        ''    Next I
        
        ''End If
     
     ElseIf strOPER = "A" Then '' Altera
     
        '' Alterando o Cabeçalho da Confirmação da Ordem de Faturamento
        sSql = "Update SGI_CADORDCONFH" & Trim(strNOMFILIAL) & " Set " & vbCrLf
        sSql = sSql & "            SGI_OBS = '" & Trim(strOBS) & "'" & vbCrLf
        sSql = sSql & "           ,SGI_CODFATURA = " & lngCODFAT & vbCrLf
        
        sSql = sSql & "     Where " & vbCrLf
        sSql = sSql & "           SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "       And SGI_CODCONF = " & lngCODCONF
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
     ElseIf strOPER = "E" Then '' Exclui
        
        '' =====================================================================
        '' Atualizando Ordem de Faturamento
        sSql = "Update SGI_CADORDFATH" & Trim(strNOMFILIAL) & " Set SGI_STATUS = 0" & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_CODORD = " & lngCODORD
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' =====================================================================
         
        '' =====================================================================
        sSql = "Update SGI_CADPEDVENDH" & Trim(strNOMFILIAL) & " Set " & vbCrLf
        If boolSALDOFECHADO = False Then
            sSql = sSql & "                       SGI_STATUS             = 'P'" & vbCrLf
        ElseIf boolSALDOFECHADO = True Then
            sSql = sSql & "                       SGI_STATUS             = 'L'" & vbCrLf
        End If
        
        sValor = "Null"
        If curQTDEATENDPED > 0 Then
            sValor = Replace(curQTDEATENDPED, ".", "")
            sValor = Replace(sValor, ",", ".")
        End If
        sSql = sSql & "                          ,SGI_QTDEITENSFATURADOS = " & Trim(sValor) & vbCrLf
        
        sSql = sSql & "                 Where " & vbCrLf
        sSql = sSql & "                       SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "                   And SGI_CODIGO = " & lngCODPED
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' =====================================================================
        '' Dando Update na Ordem de Producao
        If IsArray(arrITENSFAT) Then
            For I = 1 To UBound(arrITENSFAT)
        
                sSql = "Update SGI_ORDEMPROD" & Trim(strNOMFILIAL) & " Set " & vbCrLf
                
                sValor = "Null"
                If arrITENSFAT(I, 10) > 0 Then
                    sValor = Replace(arrITENSFAT(I, 10), ".", "")
                    sValor = Replace(sValor, ",", ".")
                End If
                sSql = sSql & "                  SGI_QTDFAT = " & sValor & vbCrLf
                
                sValor = "Null"
                If arrITENSFAT(I, 12) > 0 Then
                    sValor = Replace(arrITENSFAT(I, 12), ".", "")
                    sValor = Replace(sValor, ",", ".")
                End If
                sSql = sSql & "                 ,SGI_SALDO  = " & sValor & vbCrLf
                sSql = sSql & "                 ,SGI_STATUS = " & arrITENSFAT(I, 13) & vbCrLf
                
                sSql = sSql & " Where " & vbCrLf
                sSql = sSql & "       SGI_FILIAL    = " & intFilial & vbCrLf
                sSql = sSql & "   And SGI_CODIGO    = " & Trim(arrITENSFAT(I, 9)) & vbCrLf
                sSql = sSql & "   And SGI_IDPRODUTO = " & Trim(arrITENSFAT(I, 1))
            
                BGRV.CommandText = sSql
                BGRV.Execute
            
            Next I
        End If
        
        '' =====================================================================
        '' Apagando Itens da Confirmação da Ordem de Faturamento
        sSql = "Delete from SGI_CADORDCONFI" & Trim(strNOMFILIAL) & vbCrLf
        sSql = sSql & "        Where SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "          And SGI_CODCONF = " & lngCODCONF
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Apagando Cabeçalho da Confirmação da Ordem de Faturamento
        sSql = "Delete from SGI_CADORDCONFH" & Trim(strNOMFILIAL) & vbCrLf
        sSql = sSql & "        Where SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "          And SGI_CODCONF = " & lngCODCONF
        
        BGRV.CommandText = sSql
        BGRV.Execute
         
     
        '' =================
        '' OP's Selecionadas
        ''If IsArray(arrARROPSELECIONADA) Then
        ''    For I = 1 To UBound(arrARROPSELECIONADA)
            
        ''        sSql = ""
                
        ''        sSql = "Update SGI_ORDEMPROD" & Trim(strNOMFILIAL) & " Set SGI_STATUS = 0" & vbCrLf
        ''        sSql = sSql & "Where" & vbCrLf
        ''        sSql = sSql & "      SGI_FILIAL = " & FILIAL & vbCrLf
        ''        sSql = sSql & "  And SGI_CODIGO = " & arrARROPSELECIONADA(I, 1)
            
        ''        BGRV.CommandText = sSql
        ''        BGRV.Execute
            
        ''    Next I
        ''End If
        '' =================
     
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


Public Function Carrega_Campos(intFILIALPED As Integer) As Boolean

    Carrega_Campos = False

    Dim I               As Integer
    Dim intQTDREG       As Integer
    Dim curQTDORDEM     As Currency
    Dim curQTDJAFAT     As Currency
    Dim curQTDSALDO     As Currency
    Dim strNOMFILIAL    As String
    
    strNOMFILIAL = ""
    If intFILIALPED = 1 Then strNOMFILIAL = "_STEEL"
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADORDCONFH" & Trim(strNOMFILIAL) & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODCONF = " & lngCODCONF
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       BREC.Close
       Exit Function
    End If
    
    lngCODCONF = BREC!SGI_CODCONF
    lngCODORD = BREC!SGI_CODORD
    dtDATACONF = BREC!SGI_DATACONF
    lngCODFAT = BREC!SGI_CODFATURA
    
    If Not IsNull(BREC!SGI_OBS) Then strOBS = BREC!SGI_OBS
    
    If Not IsNull(BREC!SGI_BASEFAT) Then curBASEICMS = BREC!SGI_BASEFAT
    If Not IsNull(BREC!SGI_ALIQICMS) Then curALIQICMS = BREC!SGI_ALIQICMS
    If Not IsNull(BREC!SGI_VALOICMS) Then curVALOICMS = BREC!SGI_VALOICMS
    If Not IsNull(BREC!SGI_OUTRDESP) Then curOUTRASDESP = BREC!SGI_OUTRDESP
    If Not IsNull(BREC!SGI_FRETE) Then curFRETE = BREC!SGI_FRETE
    If Not IsNull(BREC!SGI_TOTALIPI) Then curVALORIPI = BREC!SGI_TOTALIPI
    If Not IsNull(BREC!SGI_PERCDESC) Then curPORCDESCTO = BREC!SGI_PERCDESC
    If Not IsNull(BREC!SGI_VALODESC) Then curVALORDESCT = BREC!SGI_VALODESC
    If Not IsNull(BREC!SGI_TOTALFAT) Then curVLTOTALFAT = BREC!SGI_TOTALFAT
        
    intFILIALEMP = BREC!SGI_FILIALPED
    
    BREC.Close
    
    Carrega_Campos = True
    ' ----------------------------------------------------------
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADORDCONFI" & Trim(strNOMFILIAL) & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL  = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODCONF = " & lngCODCONF
    sSql = sSql & " Order By SGI_INDICE"
    
    BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC3.EOF() Then
    
       intQTDREG = 0
       Do While Not BREC3.EOF
          intQTDREG = intQTDREG + 1
          BREC3.MoveNext
       Loop
    
       ReDim arrITENSFAT(1 To intQTDREG, 1 To 14) As Variant
       
       intQTDREG = 1
       BREC3.MoveFirst
       
       Do While Not BREC3.EOF()
       
            arrITENSFAT(intQTDREG, 1) = BREC3!SGI_IDPRODUTO
            arrITENSFAT(intQTDREG, 2) = BREC3!SGI_CODPRODUTO
            
            arrITENSFAT(intQTDREG, 3) = BREC3!SGI_QTDREAL
            If Not IsNull(BREC3!SGI_PORCIPI) Then arrITENSFAT(intQTDREG, 4) = BREC3!SGI_PORCIPI
            arrITENSFAT(intQTDREG, 5) = BREC3!SGI_VLUNIT
            arrITENSFAT(intQTDREG, 6) = BREC3!SGI_VLTOTAL
            If Not IsNull(BREC3!SGI_VLDOPI) Then arrITENSFAT(intQTDREG, 7) = BREC3!SGI_VLDOPI
            arrITENSFAT(intQTDREG, 8) = dacEnumUpdateAction_Ignore
            arrITENSFAT(intQTDREG, 9) = BREC3!SGI_CODORDPROD
            
            arrITENSFAT(intQTDREG, 10) = 0
            arrITENSFAT(intQTDREG, 11) = 0
            arrITENSFAT(intQTDREG, 12) = 0
            arrITENSFAT(intQTDREG, 13) = 0
            If Not IsNull(BREC3!SGI_CODFORN) Then arrITENSFAT(intQTDREG, 14) = BREC3!SGI_CODFORN
            
          
            '' -----------------------------------------
            '' Qtde Real da Ordem
            sSql = "Select " & vbCrLf
            sSql = sSql & "       ORD.SGI_QTDE" & vbCrLf
            sSql = sSql & "      ,ORD.SGI_QTDFAT" & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_ORDEMPROD" & Trim(strNOMFILIAL) & " ORD " & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       ORD.SGI_FILIAL  = " & intFilial & vbCrLf
            sSql = sSql & "  And  ORD.SGI_CODIGO  = " & BREC3!SGI_CODORDPROD
          
            BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
            If Not BREC10.EOF() Then
                arrITENSFAT(intQTDREG, 10) = BREC10!SGI_QTDFAT
                arrITENSFAT(intQTDREG, 11) = BREC10!SGI_QTDE
            End If
            BREC10.Close
            '' -----------------------------------------
       
            '' ----------------------------------------------------------
            curQTDORDEM = arrITENSFAT(intQTDREG, 11)
            curQTDJAFAT = (arrITENSFAT(intQTDREG, 10) - arrITENSFAT(intQTDREG, 3))
            curQTDSALDO = (curQTDORDEM - curQTDJAFAT)
            
            arrITENSFAT(intQTDREG, 10) = curQTDJAFAT
            If curQTDSALDO < curQTDORDEM Then arrITENSFAT(intQTDREG, 12) = curQTDSALDO
            If curQTDSALDO = curQTDORDEM Then arrITENSFAT(intQTDREG, 12) = 0
            
            '' Se o saldo for maior que 0 esta fechada parcial Status 1
            If curQTDJAFAT > 0 Then arrITENSFAT(intQTDREG, 13) = 1
            If curQTDJAFAT <= 0 Then arrITENSFAT(intQTDREG, 13) = 0
       
          curQTDETOTALFAT = curQTDETOTALFAT + BREC3!SGI_QTDREAL
          
          BREC3.MoveNext
          intQTDREG = (intQTDREG + 1)
       Loop
    End If
    BREC3.Close
    
End Function

Public Function Atualiza(strACAO As String, lngCODIGO As Long, lngFilial_Atu As Integer, strMODULO As String, intFILIALPEDG As Integer) As Boolean
    
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
    sSql = sSql & "  And SGI_MODULO    = '" & Trim(strMODULO) & "'" & vbCrLf
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
        sSql = sSql & "                 ,'" & Trim(strMODULO) & "'" & vbCrLf
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
        sSql = sSql & "  And SGI_MODULO    = '" & Trim(strMODULO) & "'" & vbCrLf
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

