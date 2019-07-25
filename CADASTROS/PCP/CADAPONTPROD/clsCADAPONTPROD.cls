VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADAPONTPROD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial           As Integer
Dim lngCODIGO           As Long
Dim strCODMAQ           As String
Dim strCODTURN          As String
Dim strCODOPER          As String
Dim strDTLANCT          As String
Dim arrAPONT            As Variant
Dim arrPARADAS          As Variant

Public Property Get PARADAS() As Variant
    PARADAS = arrPARADAS
End Property

Public Property Let PARADAS(ByVal vData As Variant)
    arrPARADAS = vData
End Property

Public Property Get APONT() As Variant
    APONT = arrAPONT
End Property

Public Property Let APONT(ByVal vData As Variant)
    arrAPONT = vData
End Property

Public Property Get DTLANCT() As String
    DTLANCT = strDTLANCT
End Property

Public Property Let DTLANCT(ByVal vData As String)
    strDTLANCT = vData
End Property

Public Property Get CODOPER() As String
    CODOPER = strCODOPER
End Property

Public Property Let CODOPER(ByVal vData As String)
    strCODOPER = vData
End Property

Public Property Get CODTURN() As String
    CODTURN = strCODTURN
End Property

Public Property Let CODTURN(ByVal vData As String)
    strCODTURN = vData
End Property

Public Property Get CODMAQ() As String
    CODMAQ = strCODMAQ
End Property

Public Property Let CODMAQ(ByVal vData As String)
    strCODMAQ = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Property Get CODIGO() As Long
    CODIGO = lngCODIGO
End Property

Public Property Let CODIGO(ByVal vData As Long)
    lngCODIGO = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUsuario As String, lngIDUsuario As Long, Optional intFILIAL_PED As Integer)
    frmCADAPONTPRODP.cCaminho = cCaminho
    frmCADAPONTPRODP.Linha = Linha
    frmCADAPONTPRODP.FILIAL = iFilial
    frmCADAPONTPRODP.strAcesso = strAcesso
    frmCADAPONTPRODP.strUsuario = strUsuario
    frmCADAPONTPRODP.lngCodUsuaro = lngIDUsuario
    frmCADAPONTPRODP.intFILIALPED = intFILIAL_PED
    frmCADAPONTPRODP.Show vbModal
End Sub

Public Function AtivoDesativo(strNOMFILIAL As String) As Boolean

   AtivoDesativo = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADAPPRODUCAO" & strNOMFILIAL & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC.EOF Then AtivoDesativo = True
   BREC.Close

End Function



Public Function GRAVA(strOPER As String, strNOMFILIAL As String) As Boolean

On Error GoTo grav_err
     
     GRAVA = False
     
     Dim I As Integer
     
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Or strOPER = "A" Then
     
        If strOPER = "I" Then
            
            sSql = ""
            
            sSql = "Insert Into SGI_CADAPPRODUCAO (" & vbCrLf
            sSql = sSql & "                               SGI_FILIAL" & vbCrLf
            sSql = sSql & "                              ,SGI_CODIGO" & vbCrLf
            sSql = sSql & "                              ,SGI_DTLANC" & vbCrLf
            sSql = sSql & "                              ,SGI_CODMAQ" & vbCrLf
            sSql = sSql & "                              ,SGI_CODOPE" & vbCrLf
            sSql = sSql & "                              ,SGI_CODTUN" & vbCrLf
            sSql = sSql & "                     ) Values (" & vbCrLf
            sSql = sSql & "                               " & intFilial & vbCrLf
            sSql = sSql & "                              ," & lngCODIGO & vbCrLf
            sSql = sSql & "                              ," & strDTLANCT & vbCrLf
            sSql = sSql & "                              ," & strCODMAQ & vbCrLf
            sSql = sSql & "                              ," & strCODOPER & vbCrLf
            sSql = sSql & "                              ," & strCODTURN & vbCrLf
            sSql = sSql & "                              )"
            
            BGRV.CommandText = sSql
            BGRV.Execute
        
        ElseIf strOPER = "A" Then
             
            sSql = ""
            
            sSql = "Update SGI_CADAPPRODUCAO Set " & vbCrLf
            sSql = sSql & "                               SGI_DTLANC = " & strDTLANCT & vbCrLf
            sSql = sSql & "                              ,SGI_CODMAQ = " & strCODMAQ & vbCrLf
            sSql = sSql & "                              ,SGI_CODOPE = " & strCODOPER & vbCrLf
            sSql = sSql & "                              ,SGI_CODTUN = " & strCODTURN & vbCrLf
            sSql = sSql & "                     Where" & vbCrLf
            sSql = sSql & "                           SGI_FILIAL = " & intFilial & vbCrLf
            sSql = sSql & "                       And SGI_CODIGO = " & lngCODIGO
            
            BGRV.CommandText = sSql
            BGRV.Execute
         
        End If
        
        '' Movimentos
        
        If IsArray(arrAPONT) Then
            '' Apagando
            For I = 1 To UBound(arrAPONT)
                If arrAPONT(I, 13) = dacEnumUpdateAction_delete Then
                
                    sSql = ""
                    
                    sSql = "Delete From SGI_CADAPMOVIM" & vbCrLf
                    sSql = sSql & "       Where" & vbCrLf
                    sSql = sSql & "             SGI_FILIAL    = " & intFilial & vbCrLf
                    sSql = sSql & "         And SGI_CODIGO    = " & lngCODIGO & vbCrLf
                    sSql = sSql & "         And SGI_IDINTERNO = " & arrAPONT(I, 12)
                
                    BGRV.CommandText = sSql
                    BGRV.Execute
                
                End If
            Next I
            
            '' Inserindo e Alterando
            For I = 1 To UBound(arrAPONT)
                If arrAPONT(I, 13) = dacEnumUpdateAction_Insert Then
                
                    sSql = ""
                    
                    sSql = "Insert Into  SGI_CADAPMOVIM (" & vbCrLf
                    
                    sSql = sSql & "                             SGI_FILIAL" & vbCrLf
                    sSql = sSql & "                            ,SGI_CODIGO" & vbCrLf
                    
                    sSql = sSql & "                            ,SGI_CODOP" & vbCrLf
                    sSql = sSql & "                            ,SGI_CODTIPO" & vbCrLf
                    sSql = sSql & "                            ,SGI_APLICLOTE" & vbCrLf
                    sSql = sSql & "                            ,SGI_LOTE" & vbCrLf
                    sSql = sSql & "                            ,SGI_QTDEENTR" & vbCrLf
                    sSql = sSql & "                            ,SGI_QTDESAIDA" & vbCrLf
                    sSql = sSql & "                            ,SGI_RETRABALHO" & vbCrLf
                    sSql = sSql & "                            ,SGI_HORINI" & vbCrLf
                    sSql = sSql & "                            ,SGI_HORFIN" & vbCrLf
                    sSql = sSql & "                            ,SGI_TOTLIQ" & vbCrLf
                    sSql = sSql & "                            ,SGI_INDICE" & vbCrLf
                    sSql = sSql & "                            ,SGI_IDINTERNO" & vbCrLf
                    sSql = sSql & "                            ,SGI_FILIALPED" & vbCrLf
                    
                    sSql = sSql & "                     ) Values (" & vbCrLf
                    sSql = sSql & "                               " & intFilial & vbCrLf
                    sSql = sSql & "                              ," & lngCODIGO & vbCrLf
                    
                    sSql = sSql & "                              ," & arrAPONT(I, 1) & vbCrLf
                    sSql = sSql & "                              ," & arrAPONT(I, 2) & vbCrLf
                    sSql = sSql & "                              ," & arrAPONT(I, 3) & vbCrLf
                    sSql = sSql & "                              ," & arrAPONT(I, 4) & vbCrLf
                    sSql = sSql & "                              ," & arrAPONT(I, 5) & vbCrLf
                    sSql = sSql & "                              ," & arrAPONT(I, 6) & vbCrLf
                    sSql = sSql & "                              ," & arrAPONT(I, 7) & vbCrLf
                    sSql = sSql & "                              ," & arrAPONT(I, 8) & vbCrLf
                    sSql = sSql & "                              ," & arrAPONT(I, 9) & vbCrLf
                    sSql = sSql & "                              ," & arrAPONT(I, 10) & vbCrLf
                    sSql = sSql & "                              ," & arrAPONT(I, 11) & vbCrLf
                    sSql = sSql & "                              ," & arrAPONT(I, 12) & vbCrLf
                    sSql = sSql & "                              ," & arrAPONT(I, 14) & vbCrLf
                    
                    sSql = sSql & "                              )"
                    
                    BGRV.CommandText = sSql
                    BGRV.Execute
                
                ElseIf arrAPONT(I, 13) = dacEnumUpdateAction_update Then
                
                    sSql = ""
                    
                    sSql = "Update SGI_CADAPMOVIM Set " & vbCrLf
                    sSql = sSql & "                                              SGI_CODOP      = " & arrAPONT(I, 1) & vbCrLf
                    sSql = sSql & "                                             ,SGI_CODTIPO    = " & arrAPONT(I, 2) & vbCrLf
                    sSql = sSql & "                                             ,SGI_APLICLOTE  = " & arrAPONT(I, 3) & vbCrLf
                    sSql = sSql & "                                             ,SGI_LOTE       = " & arrAPONT(I, 4) & vbCrLf
                    sSql = sSql & "                                             ,SGI_QTDEENTR   = " & arrAPONT(I, 5) & vbCrLf
                    sSql = sSql & "                                             ,SGI_QTDESAIDA  = " & arrAPONT(I, 6) & vbCrLf
                    sSql = sSql & "                                             ,SGI_RETRABALHO = " & arrAPONT(I, 7) & vbCrLf
                    sSql = sSql & "                                             ,SGI_HORINI     = " & arrAPONT(I, 8) & vbCrLf
                    sSql = sSql & "                                             ,SGI_HORFIN     = " & arrAPONT(I, 9) & vbCrLf
                    sSql = sSql & "                                             ,SGI_TOTLIQ     = " & arrAPONT(I, 10) & vbCrLf
                    sSql = sSql & "                                             ,SGI_INDICE     = " & arrAPONT(I, 11) & vbCrLf
                    
                    sSql = sSql & "                  Where " & vbCrLf
                    sSql = sSql & "                        SGI_FILIAL    = " & intFilial & vbCrLf
                    sSql = sSql & "                    And SGI_CODIGO    = " & lngCODIGO & vbCrLf
                    sSql = sSql & "                    And SGI_IDINTERNO = " & arrAPONT(I, 12)
                
                    BGRV.CommandText = sSql
                    BGRV.Execute
                
                End If
            Next I
        End If
        
        '' Paradas
        If IsArray(arrPARADAS) Then
            '' Apagando
            For I = 1 To UBound(arrPARADAS)
                If arrPARADAS(I, 5) = dacEnumUpdateAction_delete Then
                
                    sSql = ""
                    
                    sSql = "Delete From SGI_CADAPPARADAS" & vbCrLf
                    sSql = sSql & "       Where" & vbCrLf
                    sSql = sSql & "             SGI_FILIAL    = " & intFilial & vbCrLf
                    sSql = sSql & "         And SGI_CODIGO    = " & lngCODIGO & vbCrLf
                    sSql = sSql & "         And SGI_IDINTERNO = " & arrPARADAS(I, 6)
                
                    BGRV.CommandText = sSql
                    BGRV.Execute
                
                End If
            Next I
            
            '' Incluindo e Alterando
            For I = 1 To UBound(arrPARADAS)
                If arrPARADAS(I, 5) = dacEnumUpdateAction_Insert Then
                
                    sSql = ""
                    
                    sSql = "Insert Into SGI_CADAPPARADAS (" & vbCrLf
                    sSql = sSql & "                              SGI_FILIAL" & vbCrLf
                    sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
                    sSql = sSql & "                             ,SGI_CODINTPARADA" & vbCrLf
                    sSql = sSql & "                             ,SGI_IDINTERNO" & vbCrLf
                    sSql = sSql & "                             ,SGI_HORINI" & vbCrLf
                    sSql = sSql & "                             ,SGI_HORFIN" & vbCrLf
                    sSql = sSql & "                             ,SGI_TOTLIQ" & vbCrLf
                    sSql = sSql & "                             ,SGI_IDPAI" & vbCrLf
                    sSql = sSql & "                             ,SGI_INDICE" & vbCrLf
                    sSql = sSql & "                             ,SGI_FILIALPED" & vbCrLf
                    
                    sSql = sSql & "                    ) Values (" & vbCrLf
                    sSql = sSql & "                              " & intFilial & vbCrLf
                    sSql = sSql & "                             ," & lngCODIGO & vbCrLf
                    sSql = sSql & "                             ," & arrPARADAS(I, 1) & vbCrLf
                    sSql = sSql & "                             ," & arrPARADAS(I, 6) & vbCrLf
                    sSql = sSql & "                             ," & arrPARADAS(I, 2) & vbCrLf
                    sSql = sSql & "                             ," & arrPARADAS(I, 3) & vbCrLf
                    sSql = sSql & "                             ," & arrPARADAS(I, 4) & vbCrLf
                    sSql = sSql & "                             ," & arrPARADAS(I, 7) & vbCrLf
                    sSql = sSql & "                             ," & arrPARADAS(I, 8) & vbCrLf
                    sSql = sSql & "                             ," & arrPARADAS(I, 9) & vbCrLf
                    sSql = sSql & "                             )"
                
                    BGRV.CommandText = sSql
                    BGRV.Execute
                    
                ElseIf arrPARADAS(I, 5) = dacEnumUpdateAction_update Then
            
                    sSql = ""
                    
                    sSql = "Update SGI_CADAPPARADAS Set" & vbCrLf
                    sSql = sSql & "       SGI_CODINTPARADA = " & arrPARADAS(I, 1) & vbCrLf
                    sSql = sSql & "      ,SGI_HORINI       = " & arrPARADAS(I, 2) & vbCrLf
                    sSql = sSql & "      ,SGI_HORFIN       = " & arrPARADAS(I, 3) & vbCrLf
                    sSql = sSql & "      ,SGI_TOTLIQ       = " & arrPARADAS(I, 4) & vbCrLf
                    sSql = sSql & "      ,SGI_INDICE       = " & arrPARADAS(I, 8) & vbCrLf
                    sSql = sSql & "      ,SGI_IDPAI        = " & arrPARADAS(I, 7) & vbCrLf
                    sSql = sSql & " Where " & vbCrLf
                    sSql = sSql & "       SGI_FILIAL    = " & intFilial & vbCrLf
                    sSql = sSql & "   And SGI_CODIGO    = " & lngCODIGO & vbCrLf
                    sSql = sSql & "   And SGI_IDINTERNO = " & arrPARADAS(I, 6)
            
                    BGRV.CommandText = sSql
                    BGRV.Execute
            
                End If
            Next I
        End If
        
     ElseIf strOPER = "E" Then
        
        '' Paradas
        sSql = ""
        
        sSql = "Delete From SGI_CADAPPARADAS" & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL    = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO    = " & lngCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
               
     
        '' Apontamento
        sSql = ""
        
        sSql = "Delete From SGI_CADAPMOVIM" & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL    = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO    = " & lngCODIGO & vbCrLf
        
        BGRV.CommandText = sSql
        BGRV.Execute
               
        '' Cabeçalho
        sSql = ""
        
        sSql = "Delete From SGI_CADAPPRODUCAO" & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "   And SGI_CODIGO = " & lngCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
     
     End If
    
     adoBanco_Dados.CommitTrans
     GRAVA = True
     
     Exit Function
     
grav_err:
     
     adoBanco_Dados.RollbackTrans
    
     Dim objErro    As Object
     Set objErro = CreateObject("BLBCWS.clsFuncoes")
     Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, strOPER, sSql)
     Set objErro = Nothing
    
End Function

Public Function Carrega_Campos(strNOMFILIAL As String) As Boolean

    Carrega_Campos = False

    Dim I          As Integer
    Dim intQTDREG  As Integer
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADAPPRODUCAO" & strNOMFILIAL & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCODIGO
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC.EOF Then
        
        strCODMAQ = Str(BREC!SGI_CODMAQ)
        strCODTURN = Str(BREC!SGI_CODTUN)
        strCODOPER = Str(BREC!SGI_CODOPE)
        strDTLANCT = Format(BREC!SGI_DTLANC)
        
        Carrega_Campos = True
    End If
    BREC.Close
    ' ----------------------------------------------------------
    
    If Carrega_Campos = False Then Exit Function
    
    '' -----------------------------------------------------------
    '' Movimento
    sSql = ""
    arrAPONT = Empty
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADAPMOVIM" & strNOMFILIAL & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCODIGO
    
    BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC3.EOF Then
    
       intQTDREG = 0
       Do While Not BREC3.EOF
          intQTDREG = intQTDREG + 1
          BREC3.MoveNext
       Loop
    
       ReDim arrAPONT(1 To intQTDREG, 1 To 13) As Variant
       
       intQTDREG = 1
       BREC3.MoveFirst
       
       Do While Not BREC3.EOF
          
          arrAPONT(intQTDREG, 1) = BREC3!SGI_CODOP
          arrAPONT(intQTDREG, 2) = BREC3!SGI_CODTIPO
          If Not IsNull(BREC3!SGI_APLICLOTE) Then arrAPONT(intQTDREG, 3) = BREC3!SGI_APLICLOTE
          If Not IsNull(BREC3!SGI_LOTE) Then arrAPONT(intQTDREG, 4) = BREC3!SGI_LOTE
          arrAPONT(intQTDREG, 5) = BREC3!SGI_QTDEENTR
          arrAPONT(intQTDREG, 6) = BREC3!SGI_QTDESAIDA
          If Not IsNull(BREC3!SGI_RETRABALHO) Then arrAPONT(intQTDREG, 7) = BREC3!SGI_RETRABALHO
          arrAPONT(intQTDREG, 8) = BREC3!SGI_HORINI
          arrAPONT(intQTDREG, 9) = BREC3!SGI_HORFIN
          arrAPONT(intQTDREG, 10) = BREC3!SGI_TOTLIQ
          arrAPONT(intQTDREG, 11) = BREC3!SGI_IDINTERNO
          arrAPONT(intQTDREG, 12) = BREC3!SGI_INDICE
          arrAPONT(intQTDREG, 13) = BREC3!SGI_FILIALPED
          
          intQTDREG = intQTDREG + 1
          BREC3.MoveNext
       Loop
    
    End If
    BREC3.Close
    
    
    '' -----------------------------------------------------------
    '' Paradas
    sSql = ""
    arrPARADAS = Empty
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADAPPARADAS" & strNOMFILIAL & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCODIGO
    
    BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC3.EOF Then
    
       intQTDREG = 0
       Do While Not BREC3.EOF
          intQTDREG = intQTDREG + 1
          BREC3.MoveNext
       Loop
    
       ReDim arrPARADAS(1 To intQTDREG, 1 To 8) As Variant
       
       intQTDREG = 1
       BREC3.MoveFirst
       
       Do While Not BREC3.EOF
          
          arrPARADAS(intQTDREG, 1) = BREC3!SGI_INDICE
          arrPARADAS(intQTDREG, 2) = BREC3!SGI_CODINTPARADA
          arrPARADAS(intQTDREG, 3) = BREC3!SGI_IDINTERNO
          arrPARADAS(intQTDREG, 4) = BREC3!SGI_HORINI
          arrPARADAS(intQTDREG, 5) = BREC3!SGI_HORFIN
          arrPARADAS(intQTDREG, 6) = BREC3!SGI_TOTLIQ
          arrPARADAS(intQTDREG, 7) = BREC3!SGI_IDPAI
          arrPARADAS(intQTDREG, 8) = BREC3!SGI_FILIALPED
          
          intQTDREG = intQTDREG + 1
          BREC3.MoveNext
       Loop
    
    End If
    BREC3.Close
    
End Function

