VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADGRUPLINHA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial           As Integer
Dim lngCODIGO           As Long
Dim strDESCRI           As String
Dim arrCAPACIDADELIN    As Variant
Dim intATIVO            As Integer
Dim strCAPACTOTLINHA    As String
Dim arrMESANO           As Variant
Dim arrCAPACPROD        As Variant
Dim strQTDPORHORA       As String

Public Property Get QTDPORHORA() As String
    QTDPORHORA = strQTDPORHORA
End Property

Public Property Let QTDPORHORA(ByVal vData As String)
    strQTDPORHORA = vData
End Property

Public Property Get CAPACPROD() As Variant
    CAPACPROD = arrCAPACPROD
End Property

Public Property Let CAPACPROD(ByVal vData As Variant)
    arrCAPACPROD = vData
End Property

Public Property Get MESANO() As Variant
    MESANO = arrMESANO
End Property

Public Property Let MESANO(ByVal vData As Variant)
    arrMESANO = vData
End Property

Public Property Get CAPACTOTLINHA() As String
    CAPACTOTLINHA = strCAPACTOTLINHA
End Property

Public Property Let CAPACTOTLINHA(ByVal vData As String)
    strCAPACTOTLINHA = vData
End Property

Public Property Get ATIVO() As Integer
    ATIVO = intATIVO
End Property

Public Property Let ATIVO(ByVal vData As Integer)
    intATIVO = vData
End Property

Public Property Get CAPACIDADELIN() As Variant
    CAPACIDADELIN = arrCAPACIDADELIN
End Property

Public Property Let CAPACIDADELIN(ByVal vData As Variant)
    arrCAPACIDADELIN = vData
End Property

Public Property Get DESCRI() As String
    DESCRI = strDESCRI
End Property

Public Property Let DESCRI(ByVal vData As String)
    strDESCRI = vData
End Property

Public Property Get CODIGO() As Long
    CODIGO = lngCODIGO
End Property

Public Property Let CODIGO(ByVal vData As Long)
    lngCODIGO = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUsuario As String, lngIDUsuario As Long, intEMPRESA As Integer)
    frmCADGRUPLINHAP.cCaminho = cCaminho
    frmCADGRUPLINHAP.Linha = Linha
    frmCADGRUPLINHAP.FILIAL = iFilial
    frmCADGRUPLINHAP.strAcesso = strAcesso
    frmCADGRUPLINHAP.strUsuario = strUsuario
    frmCADGRUPLINHAP.lngCodUsuaro = lngIDUsuario
    frmCADGRUPLINHAP.intFILIALPED = 0
    frmCADGRUPLINHAP.intEMPRESA = intEMPRESA
    frmCADGRUPLINHAP.Show vbModal
End Sub

Public Function AtivoDesativo(strNOMETABELA As String) As Boolean

   AtivoDesativo = False
   
   sSql = ""
   
   If BREC.State = 1 Then BREC.Close
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADGRUPLINHA" & strNOMETABELA & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then AtivoDesativo = True
   BREC.Close

End Function

Public Function GRAVA(strOPER As String, strNOMETABELA As String) As Boolean

On Error GoTo err_Trans
     
     GRAVA = False
         
     Dim I As Long
     
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Then
        
        
        '' Grava Cabeçalho de Grupo de Cores
        sSql = "Insert Into SGI_CADGRUPLINHA" & strNOMETABELA & " (" & vbCrLf
        
        sSql = sSql & "                        SGI_FILIAL" & vbCrLf
        sSql = sSql & "                       ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                       ,SGI_DESCRI" & vbCrLf
        sSql = sSql & "                       ,SGI_ATIVO" & vbCrLf
        sSql = sSql & "                       ,SGI_QTDETOTAL" & vbCrLf
        sSql = sSql & "                       ,SGI_QTDEPORHORA" & vbCrLf

        sSql = sSql & "              ) Values (" & vbCrLf
        
        sSql = sSql & "                        " & intFilial & vbCrLf
        sSql = sSql & "                       ," & lngCODIGO & vbCrLf
        sSql = sSql & "                       ," & strDESCRI & vbCrLf
        sSql = sSql & "                       ," & intATIVO & vbCrLf
        sSql = sSql & "                       ," & strCAPACTOTLINHA & vbCrLf
        sSql = sSql & "                       ," & strQTDPORHORA & vbCrLf
        
        sSql = sSql & "                                )"
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        
Altera_Registro:
     
        If IsArray(arrCAPACIDADELIN) Then
            For I = 1 To UBound(arrCAPACIDADELIN)
        
                sSql = ""
                
                sSql = "Insert Into SGI_CADGRUPLINHAIT" & strNOMETABELA & " (" & vbCrLf
                sSql = sSql & "                                                     SGI_FILIAL" & vbCrLf
                sSql = sSql & "                                                    ,SGI_CODIGO" & vbCrLf
                
                sSql = sSql & "                                                    ,SGI_CODLIN" & vbCrLf         '' 1
                sSql = sSql & "                                                    ,SGI_OPTCOMNECKINSN" & vbCrLf '' 2
                sSql = sSql & "                                                    ,SGI_HOMOLOGSN" & vbCrLf      '' 3
                sSql = sSql & "                                                    ,SGI_INDICE" & vbCrLf         '' 4
                sSql = sSql & "                                                    ,SGI_IDINTERNO" & vbCrLf      '' 5
                
                sSql = sSql & "                                           ) Values (" & vbCrLf
                sSql = sSql & "                                                     " & intFilial & vbCrLf
                sSql = sSql & "                                                    ," & lngCODIGO & vbCrLf
                sSql = sSql & "                                                    ," & arrCAPACIDADELIN(I, 1) & vbCrLf
                sSql = sSql & "                                                    ," & arrCAPACIDADELIN(I, 2) & vbCrLf
                sSql = sSql & "                                                    ," & arrCAPACIDADELIN(I, 3) & vbCrLf
                sSql = sSql & "                                                    ," & arrCAPACIDADELIN(I, 4) & vbCrLf
                sSql = sSql & "                                                    ," & arrCAPACIDADELIN(I, 5) & vbCrLf
                
                sSql = sSql & "                                                     )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
                
            Next I
        End If
     
        '' -------------------------------
        '' Mês/Ano
        If IsArray(arrMESANO) Then
            For I = 1 To UBound(arrMESANO)
                sSql = ""
                
                sSql = "Insert Into SGI_MAQULIN_MESANO" & strNOMETABELA & " (" & vbCrLf
                sSql = sSql & "                                SGI_FILIAL" & vbCrLf
                sSql = sSql & "                               ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                               ,SGI_MES" & vbCrLf
                sSql = sSql & "                               ,SGI_ANO" & vbCrLf
                sSql = sSql & "                               ,SGI_INDICE" & vbCrLf
                sSql = sSql & "                               ,SGI_QTDECAPACIDADE " & vbCrLf
                sSql = sSql & "                               ,SGI_IDPAI" & vbCrLf
                
                sSql = sSql & "                     ) Values (" & vbCrLf
                sSql = sSql & "                               " & intFilial & vbCrLf
                sSql = sSql & "                              ," & lngCODIGO & vbCrLf
                
                sSql = sSql & "                              ," & arrMESANO(I, 1) & vbCrLf
                sSql = sSql & "                              ," & arrMESANO(I, 2) & vbCrLf
                sSql = sSql & "                              ,'" & arrMESANO(I, 3) & "'" & vbCrLf
                sSql = sSql & "                              ," & arrMESANO(I, 4) & vbCrLf
                sSql = sSql & "                              ," & arrMESANO(I, 5) & vbCrLf
                
                sSql = sSql & "                              )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
            Next I
        End If
        '' -------------------------------
     
     
        '' -------------------------------
        '' Capacidade Produtiva
        If IsArray(arrCAPACPROD) Then
            For I = 1 To UBound(arrCAPACPROD)
                sSql = ""
                
                sSql = "Insert Into SGI_MAQULIN_CAPAC" & strNOMETABELA & " (" & vbCrLf
                sSql = sSql & "                                SGI_FILIAL" & vbCrLf
                sSql = sSql & "                               ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                               ,SGI_DATA" & vbCrLf
                sSql = sSql & "                               ,SGI_HORAINI" & vbCrLf
                sSql = sSql & "                               ,SGI_HORAPARADA" & vbCrLf
                sSql = sSql & "                               ,SGI_HORAFIN" & vbCrLf
                sSql = sSql & "                               ,SGI_HORATOTAL" & vbCrLf
                sSql = sSql & "                               ,SGI_TOTALPECAS" & vbCrLf
                sSql = sSql & "                               ,SGI_CODINDICE" & vbCrLf
                sSql = sSql & "                               ,SGI_DIASEM" & vbCrLf
                sSql = sSql & "                               ,SGI_ATIVO" & vbCrLf
                sSql = sSql & "                               ,SGI_INDICE" & vbCrLf
                sSql = sSql & "                               ,SGI_IDPAI" & vbCrLf
                
                sSql = sSql & "                     ) Values (" & vbCrLf
                
                sSql = sSql & "                               " & intFilial & vbCrLf
                sSql = sSql & "                              ," & lngCODIGO & vbCrLf
                
                sSql = sSql & "                              ," & arrCAPACPROD(I, 1) & vbCrLf   '' 01 - SGI_DATA
                sSql = sSql & "                              ," & arrCAPACPROD(I, 2) & vbCrLf   '' 02 - SGI_HORAINI
                sSql = sSql & "                              ," & arrCAPACPROD(I, 3) & vbCrLf   '' 03 - SGI_HORAPARADA
                sSql = sSql & "                              ," & arrCAPACPROD(I, 4) & vbCrLf   '' 04 - SGI_HORAFIN
                sSql = sSql & "                              ," & arrCAPACPROD(I, 5) & vbCrLf   '' 05 - SGI_HORATOTAL
                sSql = sSql & "                              ," & arrCAPACPROD(I, 6) & vbCrLf   '' 06 - SGI_TOTALPECAS
                sSql = sSql & "                              ," & arrCAPACPROD(I, 7) & vbCrLf   '' 07 - SGI_CODINDICE
                sSql = sSql & "                              ," & arrCAPACPROD(I, 8) & vbCrLf   '' 08 - SGI_DIASEM
                sSql = sSql & "                              ," & arrCAPACPROD(I, 9) & vbCrLf   '' 09 - SGI_ATIVO
                sSql = sSql & "                              ," & arrCAPACPROD(I, 10) & vbCrLf  '' 10 - SGI_INDICE
                sSql = sSql & "                              ," & arrCAPACPROD(I, 11) & vbCrLf  '' 11 - SGI_IDPAI
                
                sSql = sSql & "                              )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
            Next I
        End If
        '' -------------------------------
     
     
     ElseIf strOPER = "A" Then
     
        
        
        sSql = ""
        
        sSql = "Delete from SGI_CADGRUPLINHAIT" & strNOMETABELA & vbCrLf
        sSql = sSql & "      Where SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODIGO
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' ==================
        
        '' ---------------------------------
        '' Mês/Ano
        sSql = ""
                        
        sSql = sSql & "Delete From SGI_MAQULIN_MESANO" & strNOMETABELA & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' ---------------------------------
        
        '' ---------------------------------
        '' Capacidade Produtiva
        sSql = ""
                        
        sSql = sSql & "Delete From SGI_MAQULIN_CAPAC" & strNOMETABELA & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' ---------------------------------
        
        
        
        sSql = ""
        
        sSql = "Update SGI_CADGRUPLINHA" & strNOMETABELA & " Set" & vbCrLf
        sSql = sSql & "                           SGI_DESCRI      = " & strDESCRI & vbCrLf
        sSql = sSql & "                          ,SGI_ATIVO       = " & intATIVO & vbCrLf
        sSql = sSql & "                          ,SGI_QTDETOTAL   = " & strCAPACTOTLINHA & vbCrLf
        sSql = sSql & "                          ,SGI_QTDEPORHORA = " & strQTDPORHORA & vbCrLf
        
        sSql = sSql & "Where" & vbCrLf
        sSql = sSql & "      SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "  And SGI_CODIGO = " & lngCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        If strOPER = "A" Then GoTo Altera_Registro
     
     ElseIf strOPER = "E" Then
     
        sSql = ""
        
        sSql = "Delete from SGI_CADGRUPLINHAIT" & strNOMETABELA & vbCrLf
        sSql = sSql & "      Where SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODIGO
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        
        '' ---------------------------------
        '' Mês/Ano
        sSql = ""
                        
        sSql = sSql & "Delete From SGI_MAQULIN_MESANO" & strNOMETABELA & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' ---------------------------------
        
        '' ---------------------------------
        '' Capacidade Produtiva
        sSql = ""
                        
        sSql = sSql & "Delete From SGI_MAQULIN_CAPAC" & strNOMETABELA & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' ---------------------------------
        
        '' Cabeçalho
        '' ==================
        sSql = ""
        
        sSql = "Delete from SGI_CADGRUPLINHA" & strNOMETABELA & vbCrLf
        sSql = sSql & "      Where SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODIGO
     
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

Public Function Carrega_campos(strNOMETABELA As String) As Boolean

On Error GoTo Err_Carrega_campos
    
    Carrega_campos = False
    
    Dim intReg      As Integer
    Dim intQTDREG   As Integer
    Dim I           As Integer
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADGRUPLINHA" & strNOMETABELA & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_CODIGO = " & lngCODIGO & vbCrLf
    sSql = sSql & "   And SGI_FILIAL   = " & intFilial
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
        
       lngCODIGO = BREC!SGI_CODIGO
       strDESCRI = BREC!SGI_DESCRI
       intATIVO = BREC!SGI_ATIVO
       If Not IsNull(BREC!SGI_QTDETOTAL) Then strCAPACTOTLINHA = Trim(Str(BREC!SGI_QTDETOTAL))
       If Not IsNull(BREC!SGI_QTDEPORHORA) Then strQTDPORHORA = Trim(Str(BREC!SGI_QTDEPORHORA))
           
       Carrega_campos = True
    End If
    BREC.Close
   
    If Carrega_campos = False Then Exit Function
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADGRUPLINHAIT" & strNOMETABELA & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO    = " & lngCODIGO & vbCrLf
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
      intReg = 0
      Do While Not BREC.EOF
         intReg = intReg + 1
         BREC.MoveNext
      Loop
      
      ReDim arrCAPACIDADELIN(1 To intReg, 1 To 5) As String
      BREC.MoveFirst
      intReg = 1
      Do While Not BREC.EOF
        
         arrCAPACIDADELIN(intReg, 1) = BREC!SGI_CODLIN
         arrCAPACIDADELIN(intReg, 2) = BREC!SGI_OPTCOMNECKINSN
         arrCAPACIDADELIN(intReg, 3) = BREC!SGI_HOMOLOGSN
         If Not IsNull(BREC!SGI_INDICE) Then arrCAPACIDADELIN(intReg, 4) = BREC!SGI_INDICE
         If Not IsNull(BREC!SGI_IDINTERNO) Then arrCAPACIDADELIN(intReg, 5) = BREC!SGI_IDINTERNO
         
         intReg = intReg + 1
         BREC.MoveNext
      Loop
    End If
    BREC.Close
    
    
    '' ==================================================================
    '' Mês/Ano
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_MAQULIN_MESANO" & strNOMETABELA & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCODIGO & vbCrLf
    sSql = sSql & "Order By SGI_Mes,SGI_Ano"
    
    BREC6.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC6.EOF() Then
        intQTDREG = 0
        Do While Not BREC6.EOF()
            intQTDREG = (intQTDREG + 1)
            BREC6.MoveNext
        Loop
        
        ReDim arrMESANO(1 To intQTDREG, 1 To 5) As String
        
        intQTDREG = 1
        BREC6.MoveFirst
        
        Do While Not BREC6.EOF()
            arrMESANO(intQTDREG, 1) = BREC6!SGI_Mes
            arrMESANO(intQTDREG, 2) = BREC6!SGI_Ano
            arrMESANO(intQTDREG, 3) = BREC6!SGI_INDICE
            If Not IsNull(BREC6!SGI_QTDECAPACIDADE) Then arrMESANO(intQTDREG, 4) = BREC6!SGI_QTDECAPACIDADE
            If Not IsNull(BREC6!SGI_IDPAI) Then arrMESANO(intQTDREG, 5) = BREC6!SGI_IDPAI
            BREC6.MoveNext
            intQTDREG = (intQTDREG + 1)
        Loop
    End If
    BREC6.Close
    '' ==================================================================


    '' ==================================================================
    '' Capacidade Produtiva
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_MAQULIN_CAPAC" & strNOMETABELA & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCODIGO & vbCrLf
    sSql = sSql & "Order By SGI_DATA"
    
    BREC6.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC6.EOF() Then
        
        intQTDREG = 0
        Do While Not BREC6.EOF()
            intQTDREG = (intQTDREG + 1)
            BREC6.MoveNext
        Loop
        
        ReDim arrCAPACPROD(1 To intQTDREG, 1 To 11) As String
        
        intQTDREG = 1
        BREC6.MoveFirst
        
        Do While Not BREC6.EOF()
        
            arrCAPACPROD(intQTDREG, 1) = Format(BREC6!SGI_DATA, "DD/MM/YYYY")
            
            If Not IsNull(BREC6!SGI_HORAINI) Then arrCAPACPROD(intQTDREG, 2) = BREC6!SGI_HORAINI
            If Not IsNull(BREC6!SGI_HORAPARADA) Then arrCAPACPROD(intQTDREG, 3) = BREC6!SGI_HORAPARADA
            If Not IsNull(BREC6!SGI_HORAFIN) Then arrCAPACPROD(intQTDREG, 4) = BREC6!SGI_HORAFIN
            If Not IsNull(BREC6!SGI_HORATOTAL) Then arrCAPACPROD(intQTDREG, 5) = BREC6!SGI_HORATOTAL
            If Not IsNull(BREC6!SGI_TOTALPECAS) Then arrCAPACPROD(intQTDREG, 6) = BREC6!SGI_TOTALPECAS
            
            arrCAPACPROD(intQTDREG, 7) = BREC6!SGI_CODINDICE
            arrCAPACPROD(intQTDREG, 8) = BREC6!SGI_DIASEM
            arrCAPACPROD(intQTDREG, 9) = BREC6!SGI_ATIVO
            arrCAPACPROD(intQTDREG, 10) = BREC6!SGI_INDICE
            If Not IsNull(BREC6!SGI_IDPAI) Then arrCAPACPROD(intQTDREG, 11) = BREC6!SGI_IDPAI
            
            BREC6.MoveNext
            intQTDREG = (intQTDREG + 1)
        Loop
    End If
    BREC6.Close
    '' ==================================================================
    
    Exit Function
    
Err_Carrega_campos:

    Carrega_campos = False
    If BREC.State = 1 Then BREC.Close
    
    Dim objErro    As Object
    Set objErro = CreateObject("BLBCWS.clsFuncoes")
    Call objErro.Sub_DescErro(Str(Err.Number), Err.Description & "Iten : " & I, "C", sSql, "frmCADMOVCXGER", "Carrega_campos")
    Set objErro = Nothing

End Function

Public Function PreenchComboMes() As String
        
    PreenchComboMes = ""
    PreenchComboMes = PreenchComboMes & "|#" & 1 & ";Janeiro"
    PreenchComboMes = PreenchComboMes & "|#" & 2 & ";Fevereiro"
    PreenchComboMes = PreenchComboMes & "|#" & 3 & ";Março"
    PreenchComboMes = PreenchComboMes & "|#" & 4 & ";Abril"
    PreenchComboMes = PreenchComboMes & "|#" & 5 & ";Maio"
    PreenchComboMes = PreenchComboMes & "|#" & 6 & ";Junho"
    PreenchComboMes = PreenchComboMes & "|#" & 7 & ";Julho"
    PreenchComboMes = PreenchComboMes & "|#" & 8 & ";Agosto"
    PreenchComboMes = PreenchComboMes & "|#" & 9 & ";Setembro"
    PreenchComboMes = PreenchComboMes & "|#" & 10 & ";Outubro"
    PreenchComboMes = PreenchComboMes & "|#" & 11 & ";Novenbro"
    PreenchComboMes = PreenchComboMes & "|#" & 12 & ";Dezembro"

End Function

Public Function PreenchComboAno() As String
        
    PreenchComboAno = ""
    PreenchComboAno = PreenchComboAno & "|#" & 2016 & ";2016"
    PreenchComboAno = PreenchComboAno & "|#" & 2017 & ";2017"
    PreenchComboAno = PreenchComboAno & "|#" & 2018 & ";2018"
    PreenchComboAno = PreenchComboAno & "|#" & 2019 & ";2019"
    PreenchComboAno = PreenchComboAno & "|#" & 2020 & ";2020"
    PreenchComboAno = PreenchComboAno & "|#" & 2021 & ";2021"
    PreenchComboAno = PreenchComboAno & "|#" & 2022 & ";2022"
    PreenchComboAno = PreenchComboAno & "|#" & 2023 & ";2023"
    PreenchComboAno = PreenchComboAno & "|#" & 2024 & ";2024"
    PreenchComboAno = PreenchComboAno & "|#" & 2025 & ";2025"
    PreenchComboAno = PreenchComboAno & "|#" & 2026 & ";2026"
    
End Function

