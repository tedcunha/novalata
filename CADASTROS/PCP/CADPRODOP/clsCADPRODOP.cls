VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADPRODOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFILIAL           As Integer
Dim lngCODIGO           As Long
Dim strDATLACTO         As String
Dim arrPROGRAMADO       As Variant
Dim arrPROGRAMADOBKP    As Variant
Dim lngCodUsuario       As Long

Public Property Get CODUSUARIO() As Long
    CODUSUARIO = lngCodUsuario
End Property

Public Property Let CODUSUARIO(ByVal vData As Long)
    lngCodUsuario = vData
End Property

Public Property Get PROGRAMADOBKP() As Variant
    PROGRAMADOBKP = arrPROGRAMADOBKP
End Property

Public Property Let PROGRAMADOBKP(ByVal vData As Variant)
    arrPROGRAMADOBKP = vData
End Property

Public Property Get PROGRAMADO() As Variant
    PROGRAMADO = arrPROGRAMADO
End Property

Public Property Let PROGRAMADO(ByVal vData As Variant)
    arrPROGRAMADO = vData
End Property

Public Property Get DATLACTO() As String
    DATLACTO = strDATLACTO
End Property

Public Property Let DATLACTO(ByVal vData As String)
    strDATLACTO = vData
End Property

Public Property Get Codigo() As Long
    Codigo = lngCODIGO
End Property

Public Property Let Codigo(ByVal vData As Long)
    lngCODIGO = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFILIAL
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFILIAL = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long, lngEMPRESA As Integer)
    frmCADPRODOPP.cCaminho = cCaminho
    frmCADPRODOPP.Linha = Linha
    frmCADPRODOPP.FILIAL = iFilial
    frmCADPRODOPP.strAcesso = strAcesso
    frmCADPRODOPP.strUSUARIO = strUSUARIO
    frmCADPRODOPP.lngCodUsuaro = lngIDUsuario
    frmCADPRODOPP.lngEMPRESA = lngEMPRESA
    frmCADPRODOPP.Show vbModal
End Sub

Public Function AtivoDesativo(strTABELA As String) As Boolean

   AtivoDesativo = False
   
   sSql = ""
   
   If BREC.State = 1 Then BREC.Close
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       Count(SGI_CODIGO) as QtdeRegs " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADAPONTPROG" & strTABELA & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFILIAL & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then
      If BREC!QtdeRegs > 0 Then AtivoDesativo = True
   End If
   BREC.Close

End Function

Public Function TemDados(strTABELA As String, strDTMOV As String) As Boolean

   TemDados = False
   
   sSql = ""
   
   If BREC.State = 1 Then BREC.Close
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       Count(SGI_CODIGO) as QtdeRegs " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADAPONTPROG" & strTABELA & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL   = " & intFILIAL & vbCrLf
   sSql = sSql & "   And SGI_DATAPONT = '" & Trim(strDTMOV) & "'" & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then
      If BREC!QtdeRegs > 0 Then TemDados = True
   End If
   BREC.Close

End Function

Public Function GRAVA(strOPER As String, strTABELA As String) As Boolean

On Error GoTo err_Trans
     
     GRAVA = False
     
     Dim I  As Integer
     
     '' Inicia transação
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If IsArray(arrPROGRAMADO) Then
        For I = 1 To UBound(arrPROGRAMADO)
            If arrPROGRAMADO(I, 9) = dacEnumUpdateAction_Insert Then
            
                sSql = ""
                
                sSql = "Insert Into SGI_CADAPONTPROG" & strTABELA & " (" & vbCrLf
                sSql = sSql & "             SGI_FILIAL" & vbCrLf
                sSql = sSql & "            ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "            ,SGI_DATAPONT" & vbCrLf
                sSql = sSql & "            ,SGI_CODOP" & vbCrLf
                sSql = sSql & "            ,SGI_QTDEPED" & vbCrLf
                sSql = sSql & "            ,SGI_QTDEPROD" & vbCrLf
                sSql = sSql & "            ,SGI_IDINTPROG" & vbCrLf
                sSql = sSql & "            ,SGI_IDINTOP" & vbCrLf
                sSql = sSql & "            ,SGI_IDLINHA" & vbCrLf
                sSql = sSql & "            ,SGI_CODLINHA" & vbCrLf
                sSql = sSql & "            ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "            ,SGI_IDINTERNO" & vbCrLf
                sSql = sSql & "            ,SGI_STATUSAPONT" & vbCrLf
                sSql = sSql & "            ,SGI_QTDEPERMONT" & vbCrLf
                sSql = sSql & "            ,SGI_QTDEFOLHAS" & vbCrLf
                sSql = sSql & "            ,SGI_PESO" & vbCrLf
                sSql = sSql & "            ,SGI_CODPED" & vbCrLf
                sSql = sSql & "            ,SGI_DTPROG" & vbCrLf
                sSql = sSql & "            ,SGI_HORINI" & vbCrLf
                sSql = sSql & "            ,SGI_HORFIN" & vbCrLf
                sSql = sSql & "            ,SGI_TOTHOR" & vbCrLf
                
                sSql = sSql & "   ) Values (" & vbCrLf
                sSql = sSql & "             " & intFILIAL & vbCrLf
                sSql = sSql & "            ," & lngCODIGO & vbCrLf
                sSql = sSql & "            ," & strDATLACTO & vbCrLf
                sSql = sSql & "            ," & arrPROGRAMADO(I, 1) & vbCrLf
                sSql = sSql & "            ," & arrPROGRAMADO(I, 2) & vbCrLf
                sSql = sSql & "            ," & arrPROGRAMADO(I, 3) & vbCrLf
                sSql = sSql & "            ," & arrPROGRAMADO(I, 4) & vbCrLf
                sSql = sSql & "            ," & arrPROGRAMADO(I, 5) & vbCrLf
                sSql = sSql & "            ," & arrPROGRAMADO(I, 6) & vbCrLf
                sSql = sSql & "            ," & arrPROGRAMADO(I, 7) & vbCrLf
                sSql = sSql & "            ," & arrPROGRAMADO(I, 8) & vbCrLf
                sSql = sSql & "            ," & arrPROGRAMADO(I, 10) & vbCrLf
                sSql = sSql & "            ," & arrPROGRAMADO(I, 11) & vbCrLf
                sSql = sSql & "            ," & arrPROGRAMADO(I, 12) & vbCrLf
                sSql = sSql & "            ," & arrPROGRAMADO(I, 13) & vbCrLf
                sSql = sSql & "            ," & arrPROGRAMADO(I, 14) & vbCrLf
                sSql = sSql & "            ," & arrPROGRAMADO(I, 15) & vbCrLf
                sSql = sSql & "            ,'" & Format(CDate(arrPROGRAMADO(I, 16)), "MM/DD/YYYY") & "'" & vbCrLf
                
                sSql = sSql & "            ," & arrPROGRAMADO(I, 18) & vbCrLf
                sSql = sSql & "            ," & arrPROGRAMADO(I, 19) & vbCrLf
                sSql = sSql & "            ," & arrPROGRAMADO(I, 20) & vbCrLf
                sSql = sSql & "            )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
                
                '' =========================================
                '' Dando Update Na Programação
                sSql = ""
                
                sSql = "Update SGI_CADMOVPCP" & strTABELA & " Set" & vbCrLf
                sSql = sSql & "                                         SGI_STATUSAPONT  = " & Trim(arrPROGRAMADO(I, 11)) & vbCrLf
                sSql = sSql & "                                       , SGI_QTDEAPONTADA = " & arrPROGRAMADO(I, 17) & vbCrLf
                sSql = sSql & "Where" & vbCrLf
                sSql = sSql & "      SGI_FILIAL    = " & intFILIAL & vbCrLf
                sSql = sSql & "  And SGI_CODINTENO = " & arrPROGRAMADO(I, 4) & vbCrLf
                sSql = sSql & "  And SGI_IDINTERNO = " & arrPROGRAMADO(I, 5) & vbCrLf
                sSql = sSql & "  And SGI_CODOP     = " & arrPROGRAMADO(I, 1)
                
                BGRV.CommandText = sSql
                BGRV.Execute
                '' =========================================
                
                '' =========================================
                '' Grava Log Pedido
                sSql = ""
                
                sSql = "Insert Into SGI_LOGPED" & strTABELA & " (" & vbCrLf
                sSql = sSql & "                              SGI_FILIAL" & vbCrLf
                sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                             ,SGI_DATA" & vbCrLf
                sSql = sSql & "                             ,SGI_HORA" & vbCrLf
                sSql = sSql & "                             ,SGI_CODUSUARIO" & vbCrLf
                sSql = sSql & "                             ,SGI_ACAO" & vbCrLf
                sSql = sSql & "                             ,SGI_DESCACAO" & vbCrLf
                sSql = sSql & "                    ) Values (" & vbCrLf
                sSql = sSql & "                              " & intFILIAL & vbCrLf                             '' Filial
                sSql = sSql & "                             ," & arrPROGRAMADO(I, 15) & vbCrLf                  '' Código da Ped
                sSql = sSql & "                             ,'" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf      '' Data da Ação
                sSql = sSql & "                             ,'" & Format(Now, "HH:MM:SS") & "'" & vbCrLf        '' Hora da Ação
                sSql = sSql & "                             ," & lngCodUsuario & vbCrLf                         '' Código do Usuário
                sSql = sSql & "                             ,'OPB'" & vbCrLf                                    '' Código da Ação
                sSql = sSql & "                             ,'" & arrPROGRAMADO(I, 1) & "'"                     '' Descrição da Ação
                sSql = sSql & "                             )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
                '' =========================================
                
            
            ElseIf arrPROGRAMADO(I, 9) = dacEnumUpdateAction_update Then
            
                sSql = ""
                
                sSql = "Update SGI_CADAPONTPROG" & strTABELA & " Set" & vbCrLf
                sSql = sSql & "            SGI_STATUSAPONT  = " & arrPROGRAMADO(I, 11) & vbCrLf
                sSql = sSql & "          , SGI_QTDEPROD     = " & arrPROGRAMADO(I, 3) & vbCrLf
                sSql = sSql & "          , SGI_QTDEPERMONT  = " & arrPROGRAMADO(I, 12) & vbCrLf
                sSql = sSql & "          , SGI_QTDEFOLHAS   = " & arrPROGRAMADO(I, 13) & vbCrLf
                sSql = sSql & "          , SGI_PESO         = " & arrPROGRAMADO(I, 14) & vbCrLf
                sSql = sSql & "          , SGI_HORINI       = " & arrPROGRAMADO(I, 18) & vbCrLf
                sSql = sSql & "          , SGI_HORFIN       = " & arrPROGRAMADO(I, 19) & vbCrLf
                sSql = sSql & "          , SGI_TOTHOR       = " & arrPROGRAMADO(I, 20) & vbCrLf
                
                sSql = sSql & "      Where" & vbCrLf
                sSql = sSql & "            SGI_FILIAL    = " & intFILIAL & vbCrLf
                sSql = sSql & "        And SGI_CODIGO    = " & lngCODIGO & vbCrLf
                sSql = sSql & "        And SGI_IDINTERNO = " & arrPROGRAMADO(I, 10)
            
                BGRV.CommandText = sSql
                BGRV.Execute
            
            
                '' =========================================
                '' Dando Update Na Programação
                sSql = ""
                
                sSql = "Update SGI_CADMOVPCP" & strTABELA & " Set" & vbCrLf
                sSql = sSql & "                                         SGI_STATUSAPONT  = " & Trim(arrPROGRAMADO(I, 11)) & vbCrLf
                sSql = sSql & "                                       , SGI_QTDEAPONTADA = " & arrPROGRAMADO(I, 17) & vbCrLf
                sSql = sSql & "Where" & vbCrLf
                sSql = sSql & "      SGI_FILIAL    = " & intFILIAL & vbCrLf
                sSql = sSql & "  And SGI_CODINTENO = " & arrPROGRAMADO(I, 4) & vbCrLf
                sSql = sSql & "  And SGI_IDINTERNO = " & arrPROGRAMADO(I, 5) & vbCrLf
                sSql = sSql & "  And SGI_CODOP     = " & arrPROGRAMADO(I, 1)
                
                BGRV.CommandText = sSql
                BGRV.Execute
                '' =========================================
            
            ElseIf arrPROGRAMADO(I, 9) = dacEnumUpdateAction_delete Then
            
                sSql = ""
                
                sSql = "Delete From SGI_CADAPONTPROG" & strTABELA & vbCrLf
                sSql = sSql & "      Where" & vbCrLf
                sSql = sSql & "            SGI_FILIAL    = " & intFILIAL & vbCrLf
                sSql = sSql & "        And SGI_CODIGO    = " & lngCODIGO & vbCrLf
                sSql = sSql & "        And SGI_IDINTERNO = " & arrPROGRAMADO(I, 10)
            
                BGRV.CommandText = sSql
                BGRV.Execute
            
                '' =========================================
                '' Dando Update Na Programação / Removendo
                sSql = ""
                
                sSql = "Update SGI_CADMOVPCP" & strTABELA & " Set" & vbCrLf
                sSql = sSql & "                                         SGI_STATUSAPONT = Null" & vbCrLf
                sSql = sSql & "Where" & vbCrLf
                sSql = sSql & "      SGI_FILIAL    = " & intFILIAL & vbCrLf
                sSql = sSql & "  And SGI_CODINTENO = " & arrPROGRAMADO(I, 4) & vbCrLf
                sSql = sSql & "  And SGI_IDINTERNO = " & arrPROGRAMADO(I, 5) & vbCrLf
                sSql = sSql & "  And SGI_CODOP     = " & arrPROGRAMADO(I, 1)
                
                BGRV.CommandText = sSql
                BGRV.Execute
                '' =========================================
            
                '' =========================================
                '' Grava Log Pedido
                sSql = ""
                
                sSql = "Insert Into SGI_LOGPED" & strTABELA & " (" & vbCrLf
                sSql = sSql & "                              SGI_FILIAL" & vbCrLf
                sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                             ,SGI_DATA" & vbCrLf
                sSql = sSql & "                             ,SGI_HORA" & vbCrLf
                sSql = sSql & "                             ,SGI_CODUSUARIO" & vbCrLf
                sSql = sSql & "                             ,SGI_ACAO" & vbCrLf
                sSql = sSql & "                             ,SGI_DESCACAO" & vbCrLf
                sSql = sSql & "                    ) Values (" & vbCrLf
                sSql = sSql & "                              " & intFILIAL & vbCrLf                             '' Filial
                sSql = sSql & "                             ," & arrPROGRAMADO(I, 15) & vbCrLf                  '' Código da Ped
                sSql = sSql & "                             ,'" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf      '' Data da Ação
                sSql = sSql & "                             ,'" & Format(Now, "HH:MM:SS") & "'" & vbCrLf        '' Hora da Ação
                sSql = sSql & "                             ," & lngCodUsuario & vbCrLf                         '' Código do Usuário
                sSql = sSql & "                             ,'OAE'" & vbCrLf                                    '' Código da Ação
                sSql = sSql & "                             ,'" & arrPROGRAMADO(I, 1) & "'"                     '' Descrição da Ação
                sSql = sSql & "                             )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
                '' =========================================
            
            End If
        Next I
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


Public Function Carrega_Campos(strTABELA As String) As Boolean

    Carrega_Campos = False

    Dim I         As Integer
    Dim intQTDREG As Integer
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       OPENV.*" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADAPONTPROG" & strTABELA & " OPENV" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       OPENV.SGI_FILIAL = " & intFILIAL & vbCrLf
    sSql = sSql & "   And OPENV.SGI_CODIGO = " & lngCODIGO
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then
        lngCODIGO = BREC!SGI_CODIGO
        strDATLACTO = Format(BREC!SGI_DATAPONT, "DD/MM/YYYY")
    End If
    BREC.Close
    
    Carrega_Campos = True
    
    '' ==============================================================
    '' Pegando os Apontamentos
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADAPONTPROG" & strTABELA & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & intFILIAL & vbCrLf
    sSql = sSql & "   And SGI_CODIGO    = " & lngCODIGO & vbCrLf
    sSql = sSql & "Order By SGI_CODLINHA"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
      intQTDREG = 0
      Do While Not BREC.EOF
         intQTDREG = intQTDREG + 1
         BREC.MoveNext
      Loop
      
      ReDim arrPROGRAMADO(1 To intQTDREG, 1 To 18) As String
      BREC.MoveFirst
      intQTDREG = 1
      Do While Not BREC.EOF
        
         arrPROGRAMADO(intQTDREG, 1) = BREC!SGI_CODOP
         arrPROGRAMADO(intQTDREG, 2) = BREC!SGI_QTDEPED
         arrPROGRAMADO(intQTDREG, 3) = BREC!SGI_QTDEPROD
         arrPROGRAMADO(intQTDREG, 4) = BREC!SGI_IDINTPROG
         arrPROGRAMADO(intQTDREG, 5) = BREC!SGI_IDINTOP
         arrPROGRAMADO(intQTDREG, 6) = BREC!SGI_IDLINHA
         arrPROGRAMADO(intQTDREG, 7) = BREC!SGI_CODLINHA
         arrPROGRAMADO(intQTDREG, 8) = BREC!SGI_IDPRODUTO
         arrPROGRAMADO(intQTDREG, 9) = BREC!SGI_IDINTERNO
         arrPROGRAMADO(intQTDREG, 10) = BREC!SGI_STATUSAPONT
         
         If Not IsNull(BREC!SGI_QTDEPERMONT) Then arrPROGRAMADO(intQTDREG, 11) = BREC!SGI_QTDEPERMONT
         If Not IsNull(BREC!SGI_QTDEFOLHAS) Then arrPROGRAMADO(intQTDREG, 12) = BREC!SGI_QTDEFOLHAS
         If Not IsNull(BREC!SGI_PESO) Then arrPROGRAMADO(intQTDREG, 13) = BREC!SGI_PESO
         
         arrPROGRAMADO(intQTDREG, 14) = BREC!SGI_CODPED
         arrPROGRAMADO(intQTDREG, 15) = BREC!SGI_DTPROG
         
         If Not IsNull(BREC!SGI_HORINI) Then arrPROGRAMADO(intQTDREG, 16) = BREC!SGI_HORINI
         If Not IsNull(BREC!SGI_HORFIN) Then arrPROGRAMADO(intQTDREG, 17) = BREC!SGI_HORFIN
         If Not IsNull(BREC!SGI_TOTHOR) Then arrPROGRAMADO(intQTDREG, 18) = BREC!SGI_TOTHOR
         
         intQTDREG = intQTDREG + 1
         BREC.MoveNext
      Loop
    End If
    BREC.Close
    
    
End Function

Public Function PreenchComboStatus() As String
        
    PreenchComboStatus = ""
    PreenchComboStatus = PreenchComboStatus & "|#" & 1 & ";Concluido"
    PreenchComboStatus = PreenchComboStatus & "|#" & 2 & ";Parcial"
    PreenchComboStatus = PreenchComboStatus & "|#" & 3 & ";Em Montagem"
    
End Function


