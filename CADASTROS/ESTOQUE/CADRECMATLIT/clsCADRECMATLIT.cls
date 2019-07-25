VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADRECMATLIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial           As Integer
Dim lngCODIGO           As Long
Dim lngCODIGOENV        As Long
Dim strDTENTRADA        As String
Dim strCODCLIE          As String
Dim strCODCLIEDEST      As String
Dim arrLANCTOS          As Variant
Dim strSTATUS           As String

Public Property Get STATUS() As String
    STATUS = strSTATUS
End Property

Public Property Let STATUS(ByVal vData As String)
    strSTATUS = vData
End Property

Public Property Get LANCTOS() As Variant
    LANCTOS = arrLANCTOS
End Property

Public Property Let LANCTOS(ByVal vData As Variant)
    arrLANCTOS = vData
End Property

Public Property Get CODCLIEDEST() As String
    CODCLIEDEST = strCODCLIEDEST
End Property

Public Property Let CODCLIEDEST(ByVal vData As String)
    strCODCLIEDEST = vData
End Property

Public Property Get CODCLIE() As String
    CODCLIE = strCODCLIE
End Property

Public Property Let CODCLIE(ByVal vData As String)
    strCODCLIE = vData
End Property

Public Property Get DTENTRADA() As String
    DTENTRADA = strDTENTRADA
End Property

Public Property Let DTENTRADA(ByVal vData As String)
    strDTENTRADA = vData
End Property

Public Property Get CODIGO() As Long
    CODIGO = lngCODIGO
End Property

Public Property Let CODIGO(ByVal vData As Long)
    lngCODIGO = vData
End Property

Public Property Get CODIGOENV() As Long
    CODIGOENV = lngCODIGOENV
End Property

Public Property Let CODIGOENV(ByVal vData As Long)
    lngCODIGOENV = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUsuario As String, lngIDUsuario As Long)
    frmCADRECMATLITP.cCaminho = cCaminho
    frmCADRECMATLITP.Linha = Linha
    frmCADRECMATLITP.FILIAL = iFilial
    frmCADRECMATLITP.strAcesso = strAcesso
    frmCADRECMATLITP.strUsuario = strUsuario
    frmCADRECMATLITP.lngCodUsuaro = lngIDUsuario
    frmCADRECMATLITP.Show vbModal
End Sub

Public Function AtivoDesativo(strNOMETABELA As String) As Boolean

   AtivoDesativo = False
   
   If Len(Trim(strNOMETABELA)) = 0 Then Exit Function
   
   sSql = ""
   
   If BREC.State = 1 Then BREC.Close
   
   sSql = ""
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       " & strNOMETABELA & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then AtivoDesativo = True
   BREC.Close

End Function


Public Function Carrega_campos_Envio() As Boolean

On Error GoTo Err_Carrega_campos_Envio
    
    Carrega_campos_Envio = False
    
    Dim intReg As Integer
    Dim i      As Integer
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADENTROTLIT" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_CODIGO = " & lngCODIGOENV & vbCrLf
    sSql = sSql & "   And SGI_FILIAL   = " & intFilial
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
       strCODCLIE = Trim(Str(BREC!SGI_CODCLIE))
       strCODCLIEDEST = Trim(Str(BREC!SGI_CODCLIEDEST))
       Carrega_campos_Envio = True
    End If
    BREC.Close
   
    If Carrega_campos_Envio = False Then Exit Function
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADENTROTLIT_IT" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO    = " & lngCODIGOENV & vbCrLf
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
      intReg = 0
      Do While Not BREC.EOF
         intReg = intReg + 1
         BREC.MoveNext
      Loop
      
      ReDim arrLANCTOS(1 To intReg, 1 To 18) As String
      BREC.MoveFirst
      intReg = 1
      Do While Not BREC.EOF
        
         arrLANCTOS(intReg, 1) = BREC!SGI_CODIT
         arrLANCTOS(intReg, 2) = BREC!SGI_CODOP
         arrLANCTOS(intReg, 3) = BREC!SGI_CODPED
         arrLANCTOS(intReg, 4) = BREC!SGI_IDPRODUTO
         arrLANCTOS(intReg, 5) = BREC!SGI_PRODUTO
         arrLANCTOS(intReg, 6) = BREC!SGI_CODCAPAC
         arrLANCTOS(intReg, 7) = BREC!SGI_PADRAO
         arrLANCTOS(intReg, 8) = Format(BREC!SGI_EXPESS, "#,####0.0000")
         arrLANCTOS(intReg, 9) = Format(BREC!SGI_LARGUR, "#,####0.0000")
         arrLANCTOS(intReg, 10) = Format(BREC!SGI_COMPRI, "#,####0.0000")
         arrLANCTOS(intReg, 11) = BREC!SGI_QTDECORP
         arrLANCTOS(intReg, 12) = Format(BREC!SGI_PERDPRODC, "#,####0.0000")
         arrLANCTOS(intReg, 13) = BREC!SGI_QTDEFOLHAS
         arrLANCTOS(intReg, 14) = Format(BREC!SGI_PESO, "#,####0.0000")
         If Not IsNull(BREC!SGI_CODUNID) Then arrLANCTOS(intReg, 15) = BREC!SGI_CODUNID
         arrLANCTOS(intReg, 16) = BREC!SGI_QTDELATAS
         arrLANCTOS(intReg, 17) = BREC!SGI_QTDEFARDOS
         If Not IsNull(BREC!SGI_CODCODTE) Then arrLANCTOS(intReg, 18) = BREC!SGI_CODCODTE
         
         intReg = intReg + 1
         BREC.MoveNext
      Loop
    End If
    BREC.Close
    
    Exit Function
    
Err_Carrega_campos_Envio:

    Carrega_campos_Envio = False
    If BREC.State = 1 Then BREC.Close
    
    Dim objErro    As Object
    Set objErro = CreateObject("BLBCWS.clsFuncoes")
    Call objErro.Sub_DescErro(Str(Err.Number), Err.Description & "Iten : " & i, "C", sSql, "frmCADMOVCXGER", "Carrega_campos_Envio")
    Set objErro = Nothing

End Function


Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_Trans
     
     GRAVA = False
         
     Dim i As Long
     
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Then
        
        
        '' Grava Cabeçalho de Grupo de Cores
        sSql = "Insert Into SGI_CADRECROTLIT (" & vbCrLf
        
        sSql = sSql & "                        SGI_FILIAL" & vbCrLf
        sSql = sSql & "                       ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                       ,SGI_DTENTRADA" & vbCrLf
        sSql = sSql & "                       ,SGI_CODCLIE" & vbCrLf
        sSql = sSql & "                       ,SGI_CODCLIEDEST" & vbCrLf
        sSql = sSql & "                       ,SGI_CODENV" & vbCrLf
        
        sSql = sSql & "              ) Values (" & vbCrLf
        
        sSql = sSql & "                        " & intFilial & vbCrLf
        sSql = sSql & "                       ," & lngCODIGO & vbCrLf
        sSql = sSql & "                       ," & strDTENTRADA & vbCrLf
        sSql = sSql & "                       ," & strCODCLIE & vbCrLf
        sSql = sSql & "                       ," & strCODCLIEDEST & vbCrLf
        sSql = sSql & "                       ," & lngCODIGOENV & vbCrLf
        
        sSql = sSql & "                                )"
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        
Altera_Registro:
     
        If IsArray(arrLANCTOS) Then
            For i = 1 To UBound(arrLANCTOS)
        
                sSql = ""
                
                sSql = "Insert Into SGI_CADRECROTLIT_IT" & " (" & vbCrLf
                sSql = sSql & "                                                     SGI_FILIAL" & vbCrLf
                sSql = sSql & "                                                    ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                                                    ,SGI_CODIT" & vbCrLf
                sSql = sSql & "                                                    ,SGI_CODOP" & vbCrLf
                sSql = sSql & "                                                    ,SGI_CODPED" & vbCrLf
                sSql = sSql & "                                                    ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                                                    ,SGI_PRODUTO" & vbCrLf
                sSql = sSql & "                                                    ,SGI_CODCAPAC" & vbCrLf
                sSql = sSql & "                                                    ,SGI_PADRAO" & vbCrLf
                sSql = sSql & "                                                    ,SGI_EXPESS" & vbCrLf
                sSql = sSql & "                                                    ,SGI_LARGUR" & vbCrLf
                sSql = sSql & "                                                    ,SGI_COMPRI" & vbCrLf
                sSql = sSql & "                                                    ,SGI_QTDECORP" & vbCrLf
                sSql = sSql & "                                                    ,SGI_PERDPRODC" & vbCrLf
                sSql = sSql & "                                                    ,SGI_QTDEFOLHAS" & vbCrLf
                sSql = sSql & "                                                    ,SGI_PESO" & vbCrLf
                sSql = sSql & "                                                    ,SGI_CODUNID" & vbCrLf
                sSql = sSql & "                                                    ,SGI_QTDELATAS" & vbCrLf
                sSql = sSql & "                                                    ,SGI_QTDEFARDOS" & vbCrLf
                sSql = sSql & "                                                    ,SGI_CODCODTE" & vbCrLf
                sSql = sSql & "                                                    ,SGI_CONFPESO" & vbCrLf
                sSql = sSql & "                                                    ,SGI_CONFFARDOS" & vbCrLf
                sSql = sSql & "                                                    ,SGI_STATUS" & vbCrLf
                sSql = sSql & "                                                    ,SGI_QTDEFOLHASREC" & vbCrLf
                sSql = sSql & "                                                    ,SGI_QTDELATASREC" & vbCrLf
                
                
                sSql = sSql & "                                           ) Values (" & vbCrLf
                sSql = sSql & "                                                     " & intFilial & vbCrLf
                sSql = sSql & "                                                    ," & lngCODIGO & vbCrLf
                sSql = sSql & "                                                    ," & arrLANCTOS(i, 1) & vbCrLf
                sSql = sSql & "                                                    ," & arrLANCTOS(i, 2) & vbCrLf
                sSql = sSql & "                                                    ," & arrLANCTOS(i, 3) & vbCrLf
                sSql = sSql & "                                                    ," & arrLANCTOS(i, 4) & vbCrLf
                sSql = sSql & "                                                    ," & arrLANCTOS(i, 5) & vbCrLf
                sSql = sSql & "                                                    ," & arrLANCTOS(i, 6) & vbCrLf
                sSql = sSql & "                                                    ," & arrLANCTOS(i, 7) & vbCrLf
                sSql = sSql & "                                                    ," & arrLANCTOS(i, 8) & vbCrLf
                sSql = sSql & "                                                    ," & arrLANCTOS(i, 9) & vbCrLf
                sSql = sSql & "                                                    ," & arrLANCTOS(i, 10) & vbCrLf
                sSql = sSql & "                                                    ," & arrLANCTOS(i, 11) & vbCrLf
                sSql = sSql & "                                                    ," & arrLANCTOS(i, 12) & vbCrLf
                sSql = sSql & "                                                    ," & arrLANCTOS(i, 13) & vbCrLf
                sSql = sSql & "                                                    ," & arrLANCTOS(i, 14) & vbCrLf
                sSql = sSql & "                                                    ," & arrLANCTOS(i, 15) & vbCrLf
                sSql = sSql & "                                                    ," & arrLANCTOS(i, 16) & vbCrLf
                sSql = sSql & "                                                    ," & arrLANCTOS(i, 17) & vbCrLf
                sSql = sSql & "                                                    ," & arrLANCTOS(i, 18) & vbCrLf
                sSql = sSql & "                                                    ," & arrLANCTOS(i, 19) & vbCrLf
                sSql = sSql & "                                                    ," & arrLANCTOS(i, 20) & vbCrLf
                sSql = sSql & "                                                    ," & arrLANCTOS(i, 21) & vbCrLf
                sSql = sSql & "                                                    ," & arrLANCTOS(i, 22) & vbCrLf
                sSql = sSql & "                                                    ," & arrLANCTOS(i, 23) & vbCrLf
                
                sSql = sSql & "                                                     )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
                
            Next i
        End If
        
        '' Dando Update No Lançamento
        sSql = ""
        sSql = sSql & "Update SGI_CADENTROTLIT Set" & vbCrLf
        sSql = sSql & "                            SGI_STATUS = " & strSTATUS & vbCrLf
        sSql = sSql & "Where" & vbCrLf
        sSql = sSql & "      SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "  And SGI_CODIGO = " & lngCODIGOENV
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
     ElseIf strOPER = "A" Then
     
        
        '' Itens
        '' ==================
        sSql = ""
        
        sSql = "Delete from SGI_CADRECROTLIT_IT" & vbCrLf
        sSql = sSql & "      Where SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODIGO
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        If strOPER = "A" Then GoTo Altera_Registro
     
     ElseIf strOPER = "E" Then
     
        
        '' Itens
        '' ==================
        sSql = ""
        
        sSql = "Delete from SGI_CADRECROTLIT_IT" & vbCrLf
        sSql = sSql & "      Where SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODIGO
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Cabeçalho
        '' ==================
        sSql = ""
        
        sSql = "Delete from SGI_CADRECROTLIT" & vbCrLf
        sSql = sSql & "      Where SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODIGO
     
        BGRV.CommandText = sSql
        BGRV.Execute
     
     
        '' Dando Update No Lançamento
        sSql = ""
        sSql = sSql & "Update SGI_CADENTROTLIT Set" & vbCrLf
        sSql = sSql & "                            SGI_STATUS = " & strSTATUS & vbCrLf
        sSql = sSql & "Where" & vbCrLf
        sSql = sSql & "      SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "  And SGI_CODIGO = " & lngCODIGOENV
        
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

On Error GoTo Err_Carrega_campos
    
    Carrega_campos = False
    
    Dim intReg As Integer
    Dim i      As Integer
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADRECROTLIT" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_CODIGO = " & lngCODIGO & vbCrLf
    sSql = sSql & "   And SGI_FILIAL   = " & intFilial
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
       lngCODIGO = BREC!SGI_CODIGO
       lngCODIGOENV = BREC!SGI_CODENV
       strDTENTRADA = Format(BREC!SGI_DTENTRADA, "DD/MM/YYYY")
       strCODCLIE = Trim(Str(BREC!SGI_CODCLIE))
       strCODCLIEDEST = Trim(Str(BREC!SGI_CODCLIEDEST))
       Carrega_campos = True
    End If
    BREC.Close
   
    If Carrega_campos = False Then Exit Function
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADRECROTLIT_IT" & vbCrLf
    
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
      
      ReDim arrLANCTOS(1 To intReg, 1 To 23) As String
      BREC.MoveFirst
      intReg = 1
      Do While Not BREC.EOF
        
         arrLANCTOS(intReg, 1) = BREC!SGI_CODIT
         arrLANCTOS(intReg, 2) = BREC!SGI_CODOP
         arrLANCTOS(intReg, 3) = BREC!SGI_CODPED
         arrLANCTOS(intReg, 4) = BREC!SGI_IDPRODUTO
         arrLANCTOS(intReg, 5) = BREC!SGI_PRODUTO
         arrLANCTOS(intReg, 6) = BREC!SGI_CODCAPAC
         arrLANCTOS(intReg, 7) = BREC!SGI_PADRAO
         arrLANCTOS(intReg, 8) = Format(BREC!SGI_EXPESS, "#,####0.0000")
         arrLANCTOS(intReg, 9) = Format(BREC!SGI_LARGUR, "#,####0.0000")
         arrLANCTOS(intReg, 10) = Format(BREC!SGI_COMPRI, "#,####0.0000")
         arrLANCTOS(intReg, 11) = BREC!SGI_QTDECORP
         arrLANCTOS(intReg, 12) = Format(BREC!SGI_PERDPRODC, "#,####0.0000")
         arrLANCTOS(intReg, 13) = BREC!SGI_QTDEFOLHAS
         arrLANCTOS(intReg, 14) = Format(BREC!SGI_PESO, "#,####0.0000")
         If Not IsNull(BREC!SGI_CODUNID) Then arrLANCTOS(intReg, 15) = BREC!SGI_CODUNID
         arrLANCTOS(intReg, 16) = BREC!SGI_QTDELATAS
         arrLANCTOS(intReg, 17) = BREC!SGI_QTDEFARDOS
         If Not IsNull(BREC!SGI_CODCODTE) Then arrLANCTOS(intReg, 18) = BREC!SGI_CODCODTE
         
         If Not IsNull(BREC!SGI_CONFPESO) Then arrLANCTOS(intReg, 19) = Format(BREC!SGI_CONFPESO, "#,####0.0000")
         If Not IsNull(BREC!SGI_CONFFARDOS) Then arrLANCTOS(intReg, 20) = BREC!SGI_CONFFARDOS
         If Not IsNull(BREC!SGI_STATUS) Then arrLANCTOS(intReg, 21) = BREC!SGI_STATUS
         
         If Not IsNull(BREC!SGI_QTDEFOLHASREC) Then arrLANCTOS(intReg, 22) = BREC!SGI_QTDEFOLHASREC
         If Not IsNull(BREC!SGI_QTDELATASREC) Then arrLANCTOS(intReg, 23) = BREC!SGI_QTDELATASREC
         
         intReg = intReg + 1
         BREC.MoveNext
      Loop
    End If
    BREC.Close
    
    Exit Function
    
Err_Carrega_campos:

    Carrega_campos = False
    If BREC.State = 1 Then BREC.Close
    
    Dim objErro    As Object
    Set objErro = CreateObject("BLBCWS.clsFuncoes")
    Call objErro.Sub_DescErro(Str(Err.Number), Err.Description & "Iten : " & i, "C", sSql, "frmCADMOVCXGER", "Carrega_campos")
    Set objErro = Nothing

End Function

