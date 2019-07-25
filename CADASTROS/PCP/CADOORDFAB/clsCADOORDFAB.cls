VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADOORDFAB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial        As Integer
Dim lngCODORDEM      As Long
Dim lngCODPEDIDO     As Long
Dim lngIDPRODUTO     As Long
Dim lngIDINTERNO     As Long
Dim dtDATAPED        As Date
Dim arrPRODUTOS      As Variant
Dim arrLOG           As Variant
Dim intSTATUS        As Integer
Dim intTipo          As Integer
Dim arrORDFAT        As Variant
Dim booFechaPED      As Boolean
Dim arrORDFATGERADAS As Variant
Dim strINCODORDS     As String
Dim strOBSLIQ        As String
Dim strCODMOTLIQOP   As String

Public Property Get LOG() As Variant
    LOG = arrLOG
End Property

Public Property Let LOG(ByVal vData As Variant)
    arrLOG = vData
End Property

Public Property Get IDINTERNO() As Long
    IDINTERNO = lngIDINTERNO
End Property

Public Property Let IDINTERNO(ByVal vData As Long)
    lngIDINTERNO = vData
End Property

Public Property Get CODMOTLIQOP() As String
    CODMOTLIQOP = strCODMOTLIQOP
End Property

Public Property Let CODMOTLIQOP(ByVal vData As String)
    strCODMOTLIQOP = vData
End Property

Public Property Get OBSLIQ() As String
    OBSLIQ = strOBSLIQ
End Property

Public Property Let OBSLIQ(ByVal vData As String)
    strOBSLIQ = vData
End Property

Public Property Get INCODORDS() As String
    INCODORDS = strINCODORDS
End Property

Public Property Let INCODORDS(ByVal vData As String)
    strINCODORDS = vData
End Property

Public Property Get ORDFATGERADAS() As Variant
    ORDFATGERADAS = arrORDFATGERADAS
End Property

Public Property Let ORDFATGERADAS(ByVal vData As Variant)
    arrORDFATGERADAS = vData
End Property

Public Property Get FechaPED() As Boolean
    FechaPED = booFechaPED
End Property

Public Property Let FechaPED(ByVal vData As Boolean)
    booFechaPED = vData
End Property

Public Property Get ORDFAT() As Variant
    ORDFAT = arrORDFAT
End Property

Public Property Let ORDFAT(ByVal vData As Variant)
    arrORDFAT = vData
End Property

Public Property Get IdProduto() As Long
    IdProduto = lngIDPRODUTO
End Property

Public Property Let IdProduto(ByVal vData As Long)
    lngIDPRODUTO = vData
End Property

Public Property Get TIPO() As Integer
    TIPO = intTipo
End Property

Public Property Let TIPO(ByVal vData As Integer)
    intTipo = vData
End Property

Public Property Get STATUS() As Integer
    STATUS = intSTATUS
End Property

Public Property Let STATUS(ByVal vData As Integer)
    intSTATUS = vData
End Property

Public Property Get CODORDEM() As Long
    CODORDEM = lngCODORDEM
End Property

Public Property Let CODORDEM(ByVal vData As Long)
    lngCODORDEM = vData
End Property

Public Property Get DATAPED() As Date
    DATAPED = dtDATAPED
End Property

Public Property Let DATAPED(ByVal vData As Date)
    dtDATAPED = vData
End Property

Public Property Get PRODUTOS() As Variant
    PRODUTOS = arrPRODUTOS
End Property

Public Property Let PRODUTOS(ByVal vData As Variant)
    arrPRODUTOS = vData
End Property

Public Property Get CODPEDIDO() As Long
    CODPEDIDO = lngCODPEDIDO
End Property

Public Property Let CODPEDIDO(ByVal vData As Long)
    lngCODPEDIDO = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUsuario As String, lngIDUsuario As Long, Optional intFILIAL_PED As Integer)
    frmCADOORDFABP.cCaminho = cCaminho
    frmCADOORDFABP.Linha = Linha
    frmCADOORDFABP.FILIAL = iFilial
    frmCADOORDFABP.strAcesso = strAcesso
    frmCADOORDFABP.strUsuario = strUsuario
    frmCADOORDFABP.lngCodUsuaro = lngIDUsuario
    frmCADOORDFABP.intFILIALPED = intFILIAL_PED
    frmCADOORDFABP.Show vbModal
End Sub

Public Function AtivoDesativo(strTabela As String) As Boolean

   AtivoDesativo = False
   
   If Len(Trim(strTabela)) = 0 Then Exit Function
   
   sSql = ""
   
   If BREC.State = 1 Then BREC.Close
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       " & Trim(strTabela) & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then AtivoDesativo = True
   BREC.Close

End Function


Public Function GRAVA(strOPER As String, strTabela1 As String, strNOMEFILIAL As String) As Boolean

On Error GoTo err_Trans
     
     GRAVA = False
     
     '' Parâmetro
     '' strTabela1 = SGI_ORDEMPROD
     
     Dim I      As Integer
     Dim sValor As String
     
     '' Inicia transação
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
      
     If strOPER = "I" Then '' Inclui
        
        '' Gravando Itens da Ordem
        If IsArray(arrPRODUTOS) Then
            For I = 1 To UBound(arrPRODUTOS)
            
                If Len(Trim(arrPRODUTOS(I, 2))) > 0 Then
                
                    sSql = "Insert Into " & strTabela1 & " (" & vbCrLf
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
                    sSql = sSql & "                      ,SGI_STATUS" & vbCrLf
                    sSql = sSql & "                      ,SGI_TIPO" & vbCrLf
                    
                    sSql = sSql & "                      ,SGI_ALTFILM" & vbCrLf
                    sSql = sSql & "                      ,SGI_FOTNOVO" & vbCrLf
                    sSql = sSql & "                      ,SGI_REPETICAO" & vbCrLf
                    sSql = sSql & "                      ,SGI_DATENTREGA" & vbCrLf
                    sSql = sSql & "                      ,SGI_TIPOP" & vbCrLf
                    sSql = sSql & "                      ,SGI_JAIMPRESSA" & vbCrLf
                    sSql = sSql & "                      ,SGI_OPCOMPL" & vbCrLf
                    sSql = sSql & "                      ,SGI_CODOPMAE" & vbCrLf
                    
                    sSql = sSql & "                    ) Values (" & vbCrLf
                    
                    sSql = sSql & "                       " & intFilial & vbCrLf
                    sSql = sSql & "                      ," & arrPRODUTOS(I, 4) & vbCrLf
                    sSql = sSql & "                      ,'" & Format(CDate(arrPRODUTOS(I, 5)), "MM/DD/YYYY") & "'" & vbCrLf
                    sSql = sSql & "                      ," & arrPRODUTOS(I, 0) & vbCrLf
                    sSql = sSql & "                      ,'" & arrPRODUTOS(I, 6) & "'" & vbCrLf
                    sSql = sSql & "                      ," & arrPRODUTOS(I, 10) & vbCrLf
                    sSql = sSql & "                      ,'" & Format(CDate(dtDATAPED), "MM/DD/YYYY") & "'" & vbCrLf
                    
                    
                    sValor = "Null"
                    If Len(Trim(arrPRODUTOS(I, 1))) > 0 And arrPRODUTOS(I, 1) > 0 Then
                        sValor = Replace(arrPRODUTOS(I, 1), ".", "")
                        sValor = Replace(sValor, ",", ".")
                    End If
                    sSql = sSql & "                      ," & sValor & vbCrLf
                    
                    sValor = "Null"
                    ''If Len(Trim(arrPRODUTOS(I, 5))) > 0 And arrPRODUTOS(I, 5) > 0 Then
                    ''    sValor = Replace(arrPRODUTOS(I, 5), ".", "")
                    ''    sValor = Replace(sValor, ",", ".")
                    ''End If
                    sSql = sSql & "                      ," & sValor & vbCrLf
                    
                    sValor = "Null"
                    If Len(Trim(arrPRODUTOS(I, 1))) > 0 And arrPRODUTOS(I, 1) > 0 Then
                        sValor = Replace(arrPRODUTOS(I, 1), ".", "")
                        sValor = Replace(sValor, ",", ".")
                    End If
                    sSql = sSql & "                      ," & sValor & vbCrLf
                    
                    sValor = "Null"
                    ''If Len(Trim(arrPRODUTOS(I, 7))) > 0 And arrPRODUTOS(I, 7) > 0 Then
                    ''    sValor = Replace(arrPRODUTOS(I, 7), ".", "")
                    ''    sValor = Replace(sValor, ",", ".")
                    ''End If
                    sSql = sSql & "                      ," & sValor & vbCrLf
                    sSql = sSql & "                      ," & intSTATUS & vbCrLf
                    sSql = sSql & "                      ," & intTipo & vbCrLf
                    
                    sSql = sSql & "                      ,0" & vbCrLf
                    sSql = sSql & "                      ,0" & vbCrLf
                    sSql = sSql & "                      ,1" & vbCrLf
                    sSql = sSql & "                      ,'" & Format(CDate(arrPRODUTOS(I, 2)), "MM/DD/YYYY") & "'" & vbCrLf
                    sSql = sSql & "                      ,1" & vbCrLf
                    sSql = sSql & "                      ,0" & vbCrLf
                    sSql = sSql & "                      ,0" & vbCrLf
                    
                    If Len(Trim(arrPRODUTOS(I, 11))) > 0 Then
                        sSql = sSql & "                      ," & arrPRODUTOS(I, 11) & vbCrLf
                    Else
                        sSql = sSql & "                      ,Null" & vbCrLf
                    End If
                    
                    sSql = sSql & "                    )"
                    
                    BGRV.CommandText = sSql
                    BGRV.Execute
                
                    '' Fecha a OP mãe
                    ''If arrPRODUTOS(I, 10) = 1 Then
                        '' Fecha a Op Mãe
                    
                        sSql = "Update " & strTabela1 & " Set " & vbCrLf
                        sSql = sSql & "                         SGI_STATUS = 2" & vbCrLf
                        sSql = sSql & "       Where " & vbCrLf
                        sSql = sSql & "             SGI_FILIAL = " & intFilial & vbCrLf
                        sSql = sSql & "         And SGI_CODIGO = " & arrPRODUTOS(I, 11)
                    
                        BGRV.CommandText = sSql
                        BGRV.Execute
                    
                    ''End If
                
                
                End If
            
            Next I
        End If
        '' =====================================================================
     
     ElseIf strOPER = "A" Then '' Altera
     
     
     ElseIf strOPER = "E" Then '' Exclui
         
        '' Apagando Ordem de Fabricação
        sSql = "Delete from " & strTabela1 & vbCrLf
        sSql = sSql & "        Where SGI_FILIAL    = " & intFilial & vbCrLf
        sSql = sSql & "          And SGI_CODPED    = " & lngCODPEDIDO & vbCrLf
        sSql = sSql & "          And SGI_IDPRODUTO = " & lngIDPRODUTO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
     ElseIf strOPER = "IMP" Then '' Atualiza Impressão
        
        '' Atualizando Impressão
        sSql = "Update " & strTabela1 & " Set SGI_JAIMPRESSA = 1" & vbCrLf
        sSql = sSql & "        Where SGI_FILIAL    = " & intFilial & vbCrLf
        sSql = sSql & "          And SGI_CODIGO    in( " & Trim(strINCODORDS) & ")" & vbCrLf
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
     ElseIf strOPER = "ALT" Then '' Atualiza Flag Alterado
        
        '' Atualiza Flag Alterado
        sSql = "Update " & strTabela1 & " Set SGI_ALTERADO = 0" & vbCrLf
        sSql = sSql & "        Where SGI_FILIAL    = " & intFilial & vbCrLf
        sSql = sSql & "          And SGI_CODIGO    in( " & Trim(strINCODORDS) & ")" & vbCrLf
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
     ElseIf strOPER = "BX" Then '' Baixa a OP Manualmente
     
        '' Fecha a Ordem de Fabricação Manualmente
        sSql = ""
        
        sSql = "Update " & strTabela1 & " Set SGI_STATUS = 9" & vbCrLf
        sSql = sSql & "                     , SGI_OBSLIQ      = " & strOBSLIQ & vbCrLf
        sSql = sSql & "                     , SGI_CODMOTLIQOP = " & strCODMOTLIQOP & vbCrLf
        
        sSql = sSql & "        Where SGI_FILIAL    = " & intFilial & vbCrLf
        sSql = sSql & "          And SGI_CODIGO    = " & lngCODORDEM & vbCrLf
        sSql = sSql & "          And SGI_IDPAI     = " & lngIDINTERNO
        
        BGRV.CommandText = sSql
        BGRV.Execute
     
        If IsArray(arrORDFAT) Then
            For I = 1 To UBound(arrORDFAT)
            
                '' Fecha a Ordem de faturamento
                sSql = "Update SGI_CADORDFATH" & strNOMEFILIAL & " Set SGI_STATUS = 1" & vbCrLf
                sSql = sSql & "        Where SGI_FILIAL    = " & intFilial & vbCrLf
                sSql = sSql & "          And SGI_CODORD    = " & Trim(arrORDFAT(I)) & vbCrLf
                
                BGRV.CommandText = sSql
                BGRV.Execute
            
            Next I
        End If
        
        sSql = ""
        
        sSql = "Update SGI_CADPEDVENDH" & strNOMEFILIAL & " Set " & vbCrLf
        sSql = sSql & "                                         SGI_STATUS = 'M'" & vbCrLf

        sSql = sSql & "Where " & vbCrLf
        sSql = sSql & "      SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "  And SGI_CODIGO = " & lngCODPEDIDO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        
        '' Atualizando Programação de Entrega
        sSql = ""
        
        sSql = "Update SGI_PROGENTRPROD" & strNOMEFILIAL & " Set " & vbCrLf
        sSql = sSql & "                                                SGI_STATUS = 9" & vbCrLf

        sSql = sSql & "Where " & vbCrLf
        sSql = sSql & "      SGI_FILIAL    = " & intFilial & vbCrLf
        sSql = sSql & "  And SGI_CODPED    = " & lngCODPEDIDO & vbCrLf
        sSql = sSql & "  And SGI_IDINTERNO = " & lngIDINTERNO
        
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
    sSql = sSql & "       SGI_ORDEMPROD  " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCODORDEM
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       BREC.Close
       Exit Function
    End If
    
    lngCODORDEM = BREC!SGI_CODIGO
    lngCODPEDIDO = BREC!SGI_CODPED
    lngIDPRODUTO = BREC!SGI_IDPRODUTO
    lngIDINTERNO = BREC!SGI_IDPAI
    dtDATAPED = BREC!SGI_DATAPED
    intSTATUS = BREC!SGI_STATUS
    intTipo = BREC!SGI_TIPO
    
    If Not IsNull(BREC!SGI_CODMOTLIQOP) Then strCODMOTLIQOP = Trim(Str(BREC!SGI_CODMOTLIQOP))
    If Not IsNull(BREC!SGI_OBSLIQ) Then strOBSLIQ = Trim(BREC!SGI_OBSLIQ)
    
    ' ----------------------------------------------------------
    intQTDREG = 0
    Do While Not BREC.EOF
       intQTDREG = intQTDREG + 1
       BREC.MoveNext
    Loop
    
    ReDim arrPRODUTOS(1 To intQTDREG, 1 To 13) As Variant
       
    intQTDREG = 1
    BREC.MoveFirst
       
    Do While Not BREC.EOF()
    
        arrPRODUTOS(intQTDREG, 1) = BREC!SGI_CODIGO
        arrPRODUTOS(intQTDREG, 2) = BREC!SGI_IDPRODUTO
        arrPRODUTOS(intQTDREG, 3) = BREC!SGI_CODPROD
        arrPRODUTOS(intQTDREG, 4) = BREC!SGI_QTDEPED
        If Not IsNull(BREC!SGI_QTDFAT) Then arrPRODUTOS(intQTDREG, 5) = BREC!SGI_QTDFAT
        arrPRODUTOS(intQTDREG, 6) = BREC!SGI_QTDE
        If Not IsNull(BREC!SGI_SALDO) Then arrPRODUTOS(intQTDREG, 7) = BREC!SGI_SALDO
        arrPRODUTOS(intQTDREG, 8) = Format(BREC!SGI_DATAORDEM, "DD/MM/YYYY")
        arrPRODUTOS(intQTDREG, 9) = dacEnumUpdateAction_Ignore
        
        arrPRODUTOS(intQTDREG, 10) = IIf(IsNull(BREC!SGI_DATENTREGA) = True, Format(BREC!SGI_DATAORDEM, "DD/MM/YYYY"), Format(BREC!SGI_DATENTREGA, "DD/MM/YYYY"))
        
        arrPRODUTOS(intQTDREG, 11) = BREC!SGI_OPCOMPL
        arrPRODUTOS(intQTDREG, 12) = BREC!SGI_CODOPMAE
        arrPRODUTOS(intQTDREG, 13) = BREC!SGI_CODPED
    
        BREC.MoveNext
        intQTDREG = (intQTDREG + 1)
    Loop
    BREC.Close
    
    
    '' Mostra as OPS Filhas
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL   = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODOPMAE = " & lngCODORDEM

    BREC8.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC8.EOF() Then
    intQTDREG = 0
    Do While Not BREC8.EOF
       intQTDREG = intQTDREG + 1
       BREC8.MoveNext
    Loop

    ReDim arrORDFATGERADAS(1 To intQTDREG, 1 To 5) As Variant
       
    intQTDREG = 1
    BREC8.MoveFirst
    
    Do While Not BREC8.EOF()
    
        arrORDFATGERADAS(intQTDREG, 1) = BREC8!SGI_IDPRODUTO
        arrORDFATGERADAS(intQTDREG, 2) = Format(BREC8!SGI_DATENTREGA, "DD/MM/YYYY")
        arrORDFATGERADAS(intQTDREG, 3) = BREC8!SGI_QTDE
        arrORDFATGERADAS(intQTDREG, 4) = BREC8!SGI_CODIGO
        arrORDFATGERADAS(intQTDREG, 5) = dacEnumUpdateAction_Ignore
    
        BREC8.MoveNext
        intQTDREG = (intQTDREG + 1)
    Loop
    End If
    BREC8.Close
    
    
    '' Log Ação
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_LOGFORM" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO    = " & lngCODORDEM & vbCrLf
    sSql = sSql & "   And SGI_FORM      = 'frmCADOORDFAB'" & vbCrLf
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
    
    
    Carrega_Campos = True

End Function

Public Function SomaQtdeItens(strCodido As String, strTabela As String) As Long

    SomaQtdeItens = 0
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       " & Trim(strTabela) & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & strCodido
    
    BREC11.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC11.EOF() Then
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       Count(*) As SGI_QTDE " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       " & Trim(strTabela) & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL   = " & intFilial & vbCrLf
        sSql = sSql & "  And SGI_CODPED    = " & BREC11!SGI_CODPED & vbCrLf
        sSql = sSql & "  And SGI_IDPRODUTO = " & BREC11!SGI_IDPRODUTO
        
        BREC12.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC12.EOF() Then SomaQtdeItens = BREC12!SGI_QTDE
        BREC12.Close

    End If
    BREC11.Close

End Function


Public Function Carrega_CamposSteel() As Boolean

    Carrega_CamposSteel = False

    Dim I         As Integer
    Dim intQTDREG As Integer
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD_STEEL  " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCODORDEM
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If BREC.EOF Then
       BREC.Close
       Exit Function
    End If
    
    lngCODORDEM = BREC!SGI_CODIGO
    lngCODPEDIDO = BREC!SGI_CODPED
    lngIDPRODUTO = BREC!SGI_IDPRODUTO
    lngIDINTERNO = BREC!SGI_IDPAI
    dtDATAPED = BREC!SGI_DATAPED
    intSTATUS = BREC!SGI_STATUS
    intTipo = BREC!SGI_TIPO
    
    If Not IsNull(BREC!SGI_CODMOTLIQOP) Then strCODMOTLIQOP = Trim(Str(BREC!SGI_CODMOTLIQOP))
    If Not IsNull(BREC!SGI_OBSLIQ) Then strOBSLIQ = Trim(BREC!SGI_OBSLIQ)
    
    ' ----------------------------------------------------------
    intQTDREG = 0
    Do While Not BREC.EOF
       intQTDREG = intQTDREG + 1
       BREC.MoveNext
    Loop
    
    ReDim arrPRODUTOS(1 To intQTDREG, 1 To 13) As Variant
       
    intQTDREG = 1
    BREC.MoveFirst
       
    Do While Not BREC.EOF()
    
        arrPRODUTOS(intQTDREG, 1) = BREC!SGI_CODIGO
        arrPRODUTOS(intQTDREG, 2) = BREC!SGI_IDPRODUTO
        arrPRODUTOS(intQTDREG, 3) = BREC!SGI_CODPROD
        arrPRODUTOS(intQTDREG, 4) = BREC!SGI_QTDEPED
        If Not IsNull(BREC!SGI_QTDFAT) Then arrPRODUTOS(intQTDREG, 5) = BREC!SGI_QTDFAT
        arrPRODUTOS(intQTDREG, 6) = BREC!SGI_QTDE
        If Not IsNull(BREC!SGI_SALDO) Then arrPRODUTOS(intQTDREG, 7) = BREC!SGI_SALDO
        arrPRODUTOS(intQTDREG, 8) = Format(BREC!SGI_DATAORDEM, "DD/MM/YYYY")
        arrPRODUTOS(intQTDREG, 9) = dacEnumUpdateAction_Ignore
        
        arrPRODUTOS(intQTDREG, 10) = IIf(IsNull(BREC!SGI_DATENTREGA) = True, Format(BREC!SGI_DATAORDEM, "DD/MM/YYYY"), Format(BREC!SGI_DATENTREGA, "DD/MM/YYYY"))
        
        arrPRODUTOS(intQTDREG, 11) = BREC!SGI_OPCOMPL
        arrPRODUTOS(intQTDREG, 12) = BREC!SGI_CODOPMAE
        arrPRODUTOS(intQTDREG, 13) = BREC!SGI_CODPED
    
        BREC.MoveNext
        intQTDREG = (intQTDREG + 1)
    Loop
    BREC.Close
    
    
    '' Mostra as OPS Filhas
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD_STEEL " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL   = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODOPMAE = " & lngCODORDEM

    BREC8.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC8.EOF() Then
    intQTDREG = 0
    Do While Not BREC8.EOF
       intQTDREG = intQTDREG + 1
       BREC8.MoveNext
    Loop

    ReDim arrORDFATGERADAS(1 To intQTDREG, 1 To 5) As Variant
       
    intQTDREG = 1
    BREC8.MoveFirst
    
    Do While Not BREC8.EOF()
    
        arrORDFATGERADAS(intQTDREG, 1) = BREC8!SGI_IDPRODUTO
        arrORDFATGERADAS(intQTDREG, 2) = Format(BREC8!SGI_DATENTREGA, "DD/MM/YYYY")
        arrORDFATGERADAS(intQTDREG, 3) = BREC8!SGI_QTDE
        arrORDFATGERADAS(intQTDREG, 4) = BREC8!SGI_CODIGO
        arrORDFATGERADAS(intQTDREG, 5) = dacEnumUpdateAction_Ignore
    
        BREC8.MoveNext
        intQTDREG = (intQTDREG + 1)
    Loop
    End If
    BREC8.Close
    
    
    '' Log Ação
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_LOGFORM" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO    = " & lngCODORDEM & vbCrLf
    sSql = sSql & "   And SGI_FORM      = 'frmCADOORDFAB_STEEL'" & vbCrLf
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
    
    Carrega_CamposSteel = True

End Function

