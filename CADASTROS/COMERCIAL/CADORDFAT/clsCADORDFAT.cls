VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADORDFAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial       As Integer
Dim lngCODORD       As Long
Dim lngCODPED       As Long
Dim dtDATAORD       As Date
Dim strOBS          As String
Dim curBASEICMS     As Currency
Dim curALIQICMS     As Currency
Dim curVALOICMS     As Currency
Dim curOUTRASDESP   As Currency
Dim curFRETE        As Currency
Dim curVALORIPI     As Currency
Dim curPORCDESCTO   As Currency
Dim curVALORDESCT   As Currency
Dim curVLTOTALFAT   As Currency
Dim arrITENSFAT     As Variant
Dim curQTDETOTALFAT As Currency
Dim curQTDETOTALPED As Currency
Dim curQTDEATENDPED As Currency
Dim sValor          As String
Dim intSTATUS       As Integer
Dim arrORDENS       As Variant
Dim intFILIALEMP    As Integer
Dim arrLOG          As Variant
Dim strNOMFORM      As String
Dim lngCodUsuario   As Long

Public Property Get CODUSUARIO() As Long
    CODUSUARIO = lngCodUsuario
End Property

Public Property Let CODUSUARIO(ByVal vData As Long)
    lngCodUsuario = vData
End Property

Public Property Get NOMFORM() As String
    NOMFORM = strNOMFORM
End Property

Public Property Let NOMFORM(ByVal vData As String)
    strNOMFORM = vData
End Property

Public Property Get LOG() As Variant
    LOG = arrLOG
End Property

Public Property Let LOG(ByVal vData As Variant)
    arrLOG = vData
End Property

Public Property Get FILIALEMP() As Integer
    FILIALEMP = intFILIALEMP
End Property

Public Property Let FILIALEMP(ByVal vData As Integer)
    intFILIALEMP = vData
End Property

Public Property Get ORDENS() As Variant
    ORDENS = arrORDENS
End Property

Public Property Let ORDENS(ByVal vData As Variant)
    arrORDENS = vData
End Property

Public Property Get STATUS() As Integer
    STATUS = intSTATUS
End Property

Public Property Let STATUS(ByVal vData As Integer)
    intSTATUS = vData
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

Public Property Get QTDETOTALFAT() As Currency
    QTDETOTALFAT = curQTDETOTALFAT
End Property

Public Property Let QTDETOTALFAT(ByVal vData As Currency)
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

Public Property Get CODPED() As Long
    CODPED = lngCODPED
End Property

Public Property Let CODPED(ByVal vData As Long)
    lngCODPED = vData
End Property

Public Property Get DATAORD() As Date
    DATAORD = dtDATAORD
End Property

Public Property Let DATAORD(ByVal vData As Date)
    dtDATAORD = vData
End Property

Public Property Get CODORD() As Long
    CODORD = lngCODORD
End Property

Public Property Let CODORD(ByVal vData As Long)
    lngCODORD = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long, Optional intFILIAL_PED As Integer)
    frmCADORDFATP.cCaminho = cCaminho
    frmCADORDFATP.Linha = Linha
    frmCADORDFATP.FILIAL = iFilial
    frmCADORDFATP.strAcesso = strAcesso
    frmCADORDFATP.strUSUARIO = strUSUARIO
    frmCADORDFATP.lngCodUsuaro = lngIDUsuario
    frmCADORDFATP.intFILIALPED = intFILIAL_PED
    frmCADORDFATP.Show vbModal
End Sub

Public Function AtivoDesativo(intFILIALPED As Integer) As Boolean

   AtivoDesativo = False
   
   If BREC.State = 1 Then BREC.Close
   
   Dim strFILIALPED    As String
   
   If intFILIALPED = 0 Then strFILIALPED = ""
   If intFILIALPED = 1 Then strFILIALPED = "_STEEL"
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       * " & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_CADORDFATH" & strFILIALPED & vbCrLf
   sSql = sSql + " Where " & vbCrLf
   sSql = sSql + "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then AtivoDesativo = True
   BREC.Close

End Function

Public Function GRAVA(strOPER As String, intFILIALPED As Integer) As Boolean

On Error GoTo err_Trans
     
     GRAVA = False
     
     Dim strFILIALPED As String
     
     If intFILIALPED = 0 Then strFILIALPED = ""
     If intFILIALPED = 1 Then strFILIALPED = "_STEEL"
     
     Dim I  As Integer
     
     '' Inicia transação
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
      
     If strOPER = "I" Then '' Inclui
        
        '' Grava Ordem HEADER(CABEÇALHO)
        sSql = "Insert Into SGI_CADORDFATH" & strFILIALPED & "  (SGI_FILIAL" & vbCrLf
        sSql = sSql & "                      ,SGI_CODORD" & vbCrLf
        sSql = sSql & "                      ,SGI_CODPED" & vbCrLf
        sSql = sSql & "                      ,SGI_DATAORDEM" & vbCrLf
        sSql = sSql & "                      ,SGI_OBS" & vbCrLf
        sSql = sSql & "                      ,SGI_STATUS" & vbCrLf
        
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
        sSql = sSql & "                      ," & lngCODORD & vbCrLf
        sSql = sSql & "                      ," & lngCODPED & vbCrLf
        sSql = sSql & "                      ,'" & Format(dtDATAORD, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                      ,'" & strOBS & "'" & vbCrLf
        sSql = sSql & "                      ," & intSTATUS & vbCrLf
        
        sValor = Replace(curQTDETOTALFAT, ".", "")
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
     
        '' Gravando Itens da Ordem
        If IsArray(arrITENSFAT) Then
            For I = 1 To UBound(arrITENSFAT)
            
                If Len(Trim(arrITENSFAT(I, 5))) > 0 Then
                
                    sSql = "Insert Into SGI_CADORDFATI" & strFILIALPED & " (" & vbCrLf
                    sSql = sSql & "                       SGI_FILIAL" & vbCrLf
                    sSql = sSql & "                      ,SGI_CODORD" & vbCrLf
                    sSql = sSql & "                      ,SGI_INDICE" & vbCrLf
                    sSql = sSql & "                      ,SGI_IDPRODUTO" & vbCrLf
                    sSql = sSql & "                      ,SGI_CODPRODUTO" & vbCrLf
                    
                    sSql = sSql & "                      ,SGI_QTDREAL" & vbCrLf
                    sSql = sSql & "                      ,SGI_QTDJAFAT" & vbCrLf
                    sSql = sSql & "                      ,SGI_QTDFAT" & vbCrLf
                    sSql = sSql & "                      ,SGI_SALDO" & vbCrLf
                    sSql = sSql & "                      ,SGI_PORCIPI" & vbCrLf
                    sSql = sSql & "                      ,SGI_VLUNIT" & vbCrLf
                    sSql = sSql & "                      ,SGI_VLTOTAL" & vbCrLf
                    sSql = sSql & "                      ,SGI_VLFATURADO" & vbCrLf
                    sSql = sSql & "                      ,SGI_VLDOPI" & vbCrLf
                    sSql = sSql & "                      ,SGI_CODORDFAB" & vbCrLf
                    sSql = sSql & "                      ,SGI_CODFORN" & vbCrLf
                    sSql = sSql & "                      ,SGI_FILIALPED" & vbCrLf
                    
                    sSql = sSql & "                    ) Values (" & vbCrLf
                    
                    sSql = sSql & "                       " & intFilial & vbCrLf
                    sSql = sSql & "                      ," & lngCODORD & vbCrLf
                    sSql = sSql & "                      ," & I & vbCrLf
                    sSql = sSql & "                      ," & arrITENSFAT(I, 1) & vbCrLf
                    sSql = sSql & "                      ,'" & Trim(arrITENSFAT(I, 2)) & "'" & vbCrLf
                    
                    sValor = Replace(arrITENSFAT(I, 3), ".", "") And arrITENSFAT(I, 3) > 0
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
                    
                    sValor = "Null"
                    If Len(Trim(arrITENSFAT(I, 8))) > 0 And arrITENSFAT(I, 8) > 0 Then
                        sValor = Replace(arrITENSFAT(I, 8), ".", "")
                        sValor = Replace(sValor, ",", ".")
                    End If
                    sSql = sSql & "                      ," & sValor & vbCrLf
                    
                    sValor = "Null"
                    If Len(Trim(arrITENSFAT(I, 9))) > 0 And arrITENSFAT(I, 9) > 0 Then
                        sValor = Replace(arrITENSFAT(I, 9), ".", "")
                        sValor = Replace(sValor, ",", ".")
                    End If
                    sSql = sSql & "                      ," & sValor & vbCrLf
                    
                    sValor = "Null"
                    If Len(Trim(arrITENSFAT(I, 10))) > 0 And arrITENSFAT(I, 10) > 0 Then
                        sValor = Replace(arrITENSFAT(I, 10), ".", "")
                        sValor = Replace(sValor, ",", ".")
                    End If
                    sSql = sSql & "                      ," & sValor & vbCrLf
                    
                    sValor = "Null"
                    If Len(Trim(arrITENSFAT(I, 11))) > 0 And arrITENSFAT(I, 11) > 0 Then
                        sValor = Replace(arrITENSFAT(I, 11), ".", "")
                        sValor = Replace(sValor, ",", ".")
                    End If
                    sSql = sSql & "                      ," & sValor & vbCrLf
                    
                    If arrITENSFAT(I, 13) > 0 Then
                        sSql = sSql & "                      ," & arrITENSFAT(I, 13) & vbCrLf
                    Else
                        sSql = sSql & "                      ,Null" & vbCrLf
                    End If
                    sSql = sSql & "                      ," & arrITENSFAT(I, 14) & vbCrLf
                    sSql = sSql & "                      ," & intFILIALEMP & vbCrLf
                    
                    sSql = sSql & "                    )"
                    
                    BGRV.CommandText = sSql
                    BGRV.Execute
                
                
                    '' =====================================
                    '' Grava Log Pedido
                    sSql = ""
                    
                    sSql = "Insert Into SGI_LOGPED" & strFILIALPED & " (" & vbCrLf
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
                    sSql = sSql & "                             ,'OF'" & vbCrLf
                    sSql = sSql & "                             ,'" & arrITENSFAT(I, 13) & "'" & vbCrLf
                    sSql = sSql & "                             )"
                    
                    BGRV.CommandText = sSql
                    BGRV.Execute
                    '' =====================================
                
                End If
            
            Next I
        End If
        
        
        
        
        '' =====================================================================
        '' Atualizando Pedido de Venda
        
        ''sValor = Replace((curQTDEATENDPED + curQTDETOTALFAT), ".", "")
        ''sValor = Replace(sValor, ",", ".")
        
        ''sSql = "Update SGI_CADPEDVENDH Set SGI_QTDEITENSFATURADOS = " & sValor & vbCrLf
        ''sSql = sSql & " Where " & vbCrLf
        ''sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
        ''sSql = sSql & "   And SGI_CODIGO = " & lngCODPED
        
        ''BGRV.CommandText = sSql
        ''BGRV.Execute
        '' =====================================================================
     
     ElseIf strOPER = "A" Then '' Altera
     
        '' =====================================================================
        '' Atualizando Pedido de Venda
        sValor = "Null"
        If (curQTDETOTALFAT - curQTDEATENDPED) > 0 Then
            sValor = Replace((curQTDETOTALFAT - curQTDEATENDPED), ".", "")
            sValor = Replace(sValor, ",", ".")
        End If
     
        sSql = "Update SGI_CADPEDVENDH" & strFILIALPED & " Set SGI_QTDEITENSFATURADOS = " & sValor & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "   And SGI_CODIGO = " & lngCODPED
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' =====================================================================
        
        
        '' Alterando o Cabeçalho da Ordem de Faturamento
        sSql = "Update SGI_CADORDFATH" & strFILIALPED & " Set " & vbCrLf
        sSql = sSql & "            SGI_OBS = '" & Trim(strOBS) & "'" & vbCrLf
        
        sValor = Replace(curQTDETOTALFAT, ".", "")
        sValor = Replace(sValor, ",", ".")
        sSql = sSql & "           ,SGI_QTDETOTFAT = " & sValor & vbCrLf
        
        sValor = Replace(curVLTOTALFAT, ".", "")
        sValor = Replace(sValor, ",", ".")
        sSql = sSql & "           ,SGI_VALTOTFAT = " & sValor & vbCrLf
        
        sValor = Replace(curBASEICMS, ".", "")
        sValor = Replace(sValor, ",", ".")
        sSql = sSql & "           ,SGI_BASEFAT = " & sValor & vbCrLf
        
        sValor = "Null"
        If curALIQICMS > 0 Then
            sValor = Replace(curALIQICMS, ".", "")
            sValor = Replace(sValor, ",", ".")
        End If
        sSql = sSql & "           ,SGI_ALIQICMS = " & sValor & vbCrLf
        
        sValor = "Null"
        If curVALOICMS > 0 Then
            sValor = Replace(curVALOICMS, ".", "")
            sValor = Replace(sValor, ",", ".")
        End If
        sSql = sSql & "           ,SGI_VALOICMS = " & sValor & vbCrLf
        
        sValor = "Null"
        If curOUTRASDESP > 0 Then
            sValor = Replace(curOUTRASDESP, ".", "")
            sValor = Replace(sValor, ",", ".")
        End If
        sSql = sSql & "           ,SGI_OUTRDESP = " & sValor & vbCrLf
        
        sValor = "Null"
        If curFRETE > 0 Then
            sValor = Replace(curFRETE, ".", "")
            sValor = Replace(sValor, ",", ".")
        End If
        sSql = sSql & "           ,SGI_FRETE = " & sValor & vbCrLf
        
        sValor = "Null"
        If curVALORIPI > 0 Then
            sValor = Replace(curVALORIPI, ".", "")
            sValor = Replace(sValor, ",", ".")
        End If
        sSql = sSql & "           ,SGI_TOTALIPI = " & sValor & vbCrLf
        
        sValor = "Null"
        If curPORCDESCTO > 0 Then
            sValor = Replace(curPORCDESCTO, ".", "")
            sValor = Replace(sValor, ",", ".")
        End If
        sSql = sSql & "           ,SGI_PERCDESC = " & sValor & vbCrLf
        
        sValor = "Null"
        If curVALORDESCT > 0 Then
            sValor = Replace(curVALORDESCT, ".", "")
            sValor = Replace(sValor, ",", ".")
        End If
        sSql = sSql & "           ,SGI_VALODESC = " & sValor & vbCrLf
        
        sValor = Replace(curVLTOTALFAT, ".", "")
        sValor = Replace(sValor, ",", ".")
        sSql = sSql & "           ,SGI_TOTALFAT = " & sValor & vbCrLf
         
        sSql = sSql & "           ,SGI_FILIALPED = " & intFILIALEMP & vbCrLf
        
        sSql = sSql & "     Where " & vbCrLf
        sSql = sSql & "           SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "       And SGI_CODORD = " & lngCODORD
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
     
        '' Gravando Itens da Ordem
        If IsArray(arrITENSFAT) Then
            For I = 1 To UBound(arrITENSFAT)
                If arrITENSFAT(I, 12) = dacEnumUpdateAction_update Then
                
                    sSql = "Update SGI_CADORDFATI" & strFILIALPED & "  Set " & vbCrLf
                    
                    sValor = "Null"
                    If Len(Trim(arrITENSFAT(I, 3))) > 0 And arrITENSFAT(I, 3) > 0 Then
                        sValor = Replace(arrITENSFAT(I, 3), ".", "")
                        sValor = Replace(sValor, ",", ".")
                    End If
                    sSql = sSql & "                      SGI_QTDREAL = " & sValor & vbCrLf
                    
                    sValor = "Null"
                    If Len(Trim(arrITENSFAT(I, 4))) > 0 And arrITENSFAT(I, 4) > 0 Then
                        sValor = Replace(arrITENSFAT(I, 4), ".", "")
                        sValor = Replace(sValor, ",", ".")
                    End If
                    sSql = sSql & "                      ,SGI_QTDJAFAT = " & sValor & vbCrLf
                    
                    sValor = "Null"
                    If Len(Trim(arrITENSFAT(I, 5))) > 0 And arrITENSFAT(I, 5) > 0 Then
                        sValor = Replace(arrITENSFAT(I, 5), ".", "")
                        sValor = Replace(sValor, ",", ".")
                    End If
                    sSql = sSql & "                      ,SGI_QTDFAT = " & sValor & vbCrLf
                    
                    sValor = "Null"
                    If Len(Trim(arrITENSFAT(I, 6))) > 0 And arrITENSFAT(I, 6) > 0 Then
                        sValor = Replace(arrITENSFAT(I, 6), ".", "")
                        sValor = Replace(sValor, ",", ".")
                    End If
                    sSql = sSql & "                      ,SGI_SALDO = " & sValor & vbCrLf
                    
                    sValor = "Null"
                    If Len(Trim(arrITENSFAT(I, 7))) > 0 And arrITENSFAT(I, 7) > 0 Then
                        sValor = Replace(arrITENSFAT(I, 7), ".", "")
                        sValor = Replace(sValor, ",", ".")
                    End If
                    sSql = sSql & "                      ,SGI_PORCIPI = " & sValor & vbCrLf
                    
                    sValor = "Null"
                    If Len(Trim(arrITENSFAT(I, 8))) > 0 And arrITENSFAT(I, 8) > 0 Then
                        sValor = Replace(arrITENSFAT(I, 8), ".", "")
                        sValor = Replace(sValor, ",", ".")
                    End If
                    sSql = sSql & "                      ,SGI_VLUNIT = " & sValor & vbCrLf
                    
                    sValor = "Null"
                    If Len(Trim(arrITENSFAT(I, 9))) > 0 And arrITENSFAT(I, 9) > 0 Then
                        sValor = Replace(arrITENSFAT(I, 9), ".", "")
                        sValor = Replace(sValor, ",", ".")
                    End If
                    sSql = sSql & "                      ,SGI_VLTOTAL = " & sValor & vbCrLf
                    
                    sValor = "Null"
                    If Len(Trim(arrITENSFAT(I, 10))) > 0 And arrITENSFAT(I, 10) > 0 Then
                        sValor = Replace(arrITENSFAT(I, 10), ".", "")
                        sValor = Replace(sValor, ",", ".")
                    End If
                    sSql = sSql & "                      ,SGI_VLFATURADO = " & sValor & vbCrLf
                    
                    sValor = "Null"
                    If Len(Trim(arrITENSFAT(I, 11))) > 0 And arrITENSFAT(I, 11) > 0 Then
                        sValor = Replace(arrITENSFAT(I, 11), ".", "")
                        sValor = Replace(sValor, ",", ".")
                    End If
                    sSql = sSql & "                      ,SGI_VLDOPI = " & sValor & vbCrLf
                    
                    sSql = sSql & "                      ,SGI_CODFORN   = " & arrITENSFAT(I, 14) & vbCrLf
                    sSql = sSql & "                      ,SGI_FILIALPED = " & intFILIALEMP & vbCrLf
                    
                    sSql = sSql & "  Where " & vbCrLf
                    sSql = sSql & "        SGI_FILIAL     = " & intFilial & vbCrLf
                    sSql = sSql & "    And SGI_IDPRODUTO  = " & arrITENSFAT(I, 1) & vbCrLf
                    sSql = sSql & "    And SGI_CODORD     = " & lngCODORD & vbCrLf
                
                    BGRV.CommandText = sSql
                    BGRV.Execute
                
                
                    '' =====================================
                    '' Grava Log
                    sSql = ""
                    
                    sSql = "Insert Into SGI_LOGPED" & strFILIALPED & " (" & vbCrLf
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
                    sSql = sSql & "                             ,'OFA'" & vbCrLf
                    sSql = sSql & "                             ,'" & arrITENSFAT(I, 13) & "'" & vbCrLf
                    sSql = sSql & "                             )"
                    
                    BGRV.CommandText = sSql
                    BGRV.Execute
                    '' =====================================
                
                
                
                End If
                
            Next I
        End If
     
     
     
     ElseIf strOPER = "E" Then '' Exclui
        
        '' =====================================================================
        '' Atualizando Pedido de Venda
        ''sValor = "Null"
        ''If (curQTDEATENDPED - curQTDETOTALFAT) > 0 Then
        ''    sValor = Replace((curQTDEATENDPED - curQTDETOTALFAT), ".", "")
        ''    sValor = Replace(sValor, ",", ".")
        ''End If
        
        ''sSql = "Update SGI_CADPEDVENDH Set SGI_QTDEITENSFATURADOS = " & sValor & vbCrLf
        ''sSql = sSql & " Where " & vbCrLf
        ''sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
        ''sSql = sSql & "   And SGI_CODIGO = " & lngCODPED
        
        ''BGRV.CommandText = sSql
        ''BGRV.Execute
        '' =====================================================================
         
        If IsArray(arrORDENS) Then
        
            For I = 1 To UBound(arrORDENS)
        
                '' Apagando Itens da Ordem de Faturamento
                sSql = "Delete from SGI_CADORDFATI" & strFILIALPED & " " & vbCrLf
                sSql = sSql & "        Where SGI_FILIAL = " & intFilial & vbCrLf
                sSql = sSql & "          And SGI_CODORD = " & arrORDENS(I)
                
                BGRV.CommandText = sSql
                BGRV.Execute
                
                '' Apagando Cabeçalho da Ordem de Faturamento
                sSql = "Delete from SGI_CADORDFATH" & strFILIALPED & " " & vbCrLf
                sSql = sSql & "        Where SGI_FILIAL = " & intFilial & vbCrLf
                sSql = sSql & "          And SGI_CODORD = " & arrORDENS(I)
                
                BGRV.CommandText = sSql
                BGRV.Execute
         
            Next I
        End If
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
    Dim strFILIALPED    As String
    
    If intFILIALPED = 0 Then strFILIALPED = ""
    If intFILIALPED = 1 Then strFILIALPED = "_STEEL"
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADORDFATH" & strFILIALPED & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODORD = " & lngCODORD
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       BREC.Close
       Exit Function
    End If
    
    lngCODORD = BREC!SGI_CODORD
    lngCODPED = BREC!SGI_CODPED
    dtDATAORD = BREC!SGI_DATAORDEM
    intSTATUS = BREC!SGI_STATUS
    intFILIALEMP = BREC!SGI_FILIALPED
    
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
    
    BREC.Close
    
    Carrega_Campos = True
    ' ----------------------------------------------------------
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADORDFATI" & strFILIALPED & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODORD = " & lngCODORD
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
          If Not IsNull(BREC3!SGI_QTDJAFAT) Then arrITENSFAT(intQTDREG, 4) = BREC3!SGI_QTDJAFAT
          If Not IsNull(BREC3!SGI_QTDFAT) Then arrITENSFAT(intQTDREG, 5) = BREC3!SGI_QTDFAT
          If Not IsNull(BREC3!SGI_SALDO) Then arrITENSFAT(intQTDREG, 6) = BREC3!SGI_SALDO
          If Not IsNull(BREC3!SGI_PORCIPI) Then arrITENSFAT(intQTDREG, 7) = BREC3!SGI_PORCIPI
          arrITENSFAT(intQTDREG, 8) = BREC3!SGI_VLUNIT
          arrITENSFAT(intQTDREG, 9) = BREC3!SGI_VLTOTAL
          If Not IsNull(BREC3!SGI_VLFATURADO) Then arrITENSFAT(intQTDREG, 10) = BREC3!SGI_VLFATURADO
          If Not IsNull(BREC3!SGI_VLDOPI) Then arrITENSFAT(intQTDREG, 11) = BREC3!SGI_VLDOPI
          arrITENSFAT(intQTDREG, 12) = dacEnumUpdateAction_Ignore
          If Not IsNull(BREC3!SGI_CODORDFAB) Then arrITENSFAT(intQTDREG, 13) = BREC3!SGI_CODORDFAB
          If Not IsNull(BREC3!SGI_CODFORN) Then arrITENSFAT(intQTDREG, 14) = BREC3!SGI_CODFORN
       
          curQTDETOTALPED = curQTDETOTALPED + BREC3!SGI_QTDREAL
          
          BREC3.MoveNext
          intQTDREG = (intQTDREG + 1)
       Loop
    End If
    BREC3.Close

    '' Log Ação
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_LOGFORM" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO    = " & lngCODORD & vbCrLf
    sSql = sSql & "   And SGI_FORM      = '" & strNOMFORM & "'" & vbCrLf
    sSql = sSql & "Order By SGI_DATA,SGI_HORA"
    
    BREC4.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC4.EOF() Then
        intQTDREG = 0
        Do While Not BREC4.EOF()
            intQTDREG = intQTDREG + 1
            BREC4.MoveNext
        Loop
    
        ReDim arrLOG(1 To intQTDREG, 1 To 4) As Variant
        
        intQTDREG = 1
        BREC4.MoveFirst
    
        Do While Not BREC4.EOF()
            arrLOG(intQTDREG, 1) = Format(BREC4!SGI_DATA, "DD/MM/YYYY")
            arrLOG(intQTDREG, 2) = Format(BREC4!SGI_HORA, "HH:MM:SS")
            arrLOG(intQTDREG, 3) = BREC4!SGI_CODUSUARIO
            arrLOG(intQTDREG, 4) = BREC4!SGI_ACAO
            
            intQTDREG = intQTDREG + 1
            BREC4.MoveNext
        Loop
    End If
    BREC4.Close


End Function

