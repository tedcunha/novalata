VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADPROGOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial           As Integer
Dim lngCODIGO           As Long
Dim strDTPROGRAMA       As String
Dim arrPROGRAMADO       As Variant
Dim arrPROGRAMADODEL    As Variant
Dim arrPROGRAMADOBKP    As Variant
Dim arrMUDASTATUSOP     As Variant
Dim arrGERAFOLHAS       As Variant
Dim lngCodUsuario       As Long

Public Property Get GERAFOLHAS() As Variant
    GERAFOLHAS = arrGERAFOLHAS
End Property

Public Property Let GERAFOLHAS(ByVal vData As Variant)
    arrGERAFOLHAS = vData
End Property

Public Property Get PROGRAMADODEL() As Variant
    PROGRAMADODEL = arrPROGRAMADODEL
End Property

Public Property Let PROGRAMADODEL(ByVal vData As Variant)
    arrPROGRAMADODEL = vData
End Property

Public Property Get CODUSUARIO() As Long
    CODUSUARIO = lngCodUsuario
End Property

Public Property Let CODUSUARIO(ByVal vData As Long)
    lngCodUsuario = vData
End Property

Public Property Get MUDASTATUSOP() As Variant
    MUDASTATUSOP = arrMUDASTATUSOP
End Property

Public Property Let MUDASTATUSOP(ByVal vData As Variant)
    arrMUDASTATUSOP = vData
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

Public Property Get DTPROGRAMA() As String
    DTPROGRAMA = strDTPROGRAMA
End Property

Public Property Let DTPROGRAMA(ByVal vData As String)
    strDTPROGRAMA = vData
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

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUsuario As String, lngIDUsuario As Long, intFILIALPED As Integer)
    frmCADPROGOPP.cCaminho = cCaminho
    frmCADPROGOPP.Linha = Linha
    frmCADPROGOPP.FILIAL = iFilial
    frmCADPROGOPP.strAcesso = strAcesso
    frmCADPROGOPP.strUsuario = strUsuario
    frmCADPROGOPP.lngCodUsuaro = lngIDUsuario
    frmCADPROGOPP.intFILIALPED = intFILIALPED
    frmCADPROGOPP.Show vbModal
End Sub

Public Function AtivoDesativo(strNomeModulo As String) As Boolean

   AtivoDesativo = False
   
   sSql = ""
   
   If BREC.State = 1 Then BREC.Close
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADMOVPCP" & strNomeModulo & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then AtivoDesativo = True
   BREC.Close

End Function


Public Function GRAVA(strOPER As String, strNOMFILIAL As String) As Boolean

On Error GoTo grav_err
     
    GRAVA = False
    
    Dim i As Integer
     
    adoBanco_Dados.BeginTrans
    BGRV.ActiveConnection = adoBanco_Dados
    
    '' ===============================
    '' Apagando as Folhas Selecionadas
    sSql = ""
    sSql = "Delete From SGI_CADMOVPCP_FOLHAS" & strNOMFILIAL & vbCrLf
    sSql = sSql & "Where" & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "  And SGI_CODIGO = " & lngCODIGO
    
    BGRV.CommandText = sSql
    BGRV.Execute
    '' ===============================
    
    '' ===============================
    '' Excluindo PAra refazer
    If IsArray(arrPROGRAMADODEL) Then
        For i = 1 To UBound(arrPROGRAMADODEL)
            If arrPROGRAMADODEL(i, 8) = dacEnumUpdateAction_delete Then
            
                '' ===============================
                '' Update no OP
                sSql = ""
                
                sSql = "Update SGI_ORDEMPROD" & strNOMFILIAL & " Set" & vbCrLf
                
                sSql = sSql & "                                       SGI_PROGRAMADO = 0" & vbCrLf
                sSql = sSql & "                                      ,SGI_DATENTREGA = '" & Format(CDate(arrPROGRAMADODEL(i, 11)), "MM/DD/YYYY") & "'" & vbCrLf
                sSql = sSql & "                                      ,SGI_ALTERADO   = 1" & vbCrLf
                sSql = sSql & "                                      ,SGI_JAIMPRESSA = 0" & vbCrLf
                
                sSql = sSql & "Where " & vbCrLf
                sSql = sSql & "      SGI_FILIAl     = " & intFilial & vbCrLf
                sSql = sSql & "  And SGI_CODIGO     = " & arrPROGRAMADODEL(i, 20) & vbCrLf
                sSql = sSql & "  And SGI_IDPRODUTO  = " & arrPROGRAMADODEL(i, 18) & vbCrLf
                sSql = sSql & "  And SGI_IDPAI      = " & arrPROGRAMADODEL(i, 19)
                
                BGRV.CommandText = sSql
                BGRV.Execute
                '' ===============================
            
            
                '' ===============================
                '' Programação de Entrtega
                sSql = ""
                            
                sSql = sSql & "Update SGI_PROGENTRPROD" & strNOMFILIAL & " Set" & vbCrLf
                sSql = sSql & "                                                 SGI_DATENTREGA = '" & Format(CDate(arrPROGRAMADODEL(i, 11)), "MM/DD/YYYY") & "'" & vbCrLf
                
                sSql = sSql & "       Where" & vbCrLf
                sSql = sSql & "             SGI_FILIAl    = " & intFilial & vbCrLf
                sSql = sSql & "         And SGI_IDINTERNO = " & arrPROGRAMADODEL(i, 19) & vbCrLf
                sSql = sSql & "         And SGI_CODPED    = " & arrPROGRAMADODEL(i, 17) & vbCrLf
                sSql = sSql & "         And SGI_IDPRODUTO = " & arrPROGRAMADODEL(i, 18)
                
                BGRV.CommandText = sSql
                BGRV.Execute
                '' ===============================
            
                '' =========================================
                '' Grava Log Pedido
                sSql = ""
                
                sSql = "Insert Into SGI_LOGPED" & strNOMFILIAL & " (" & vbCrLf
                sSql = sSql & "                              SGI_FILIAL" & vbCrLf
                sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                             ,SGI_DATA" & vbCrLf
                sSql = sSql & "                             ,SGI_HORA" & vbCrLf
                sSql = sSql & "                             ,SGI_CODUSUARIO" & vbCrLf
                sSql = sSql & "                             ,SGI_ACAO" & vbCrLf
                sSql = sSql & "                             ,SGI_DESCACAO" & vbCrLf
                sSql = sSql & "                    ) Values (" & vbCrLf
                sSql = sSql & "                              " & intFilial & vbCrLf                             '' Filial
                sSql = sSql & "                             ," & arrPROGRAMADODEL(i, 17) & vbCrLf               '' Código da Ped
                sSql = sSql & "                             ,'" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf      '' Data da Ação
                sSql = sSql & "                             ,'" & Format(Now, "HH:MM:SS") & "'" & vbCrLf        '' Hora da Ação
                sSql = sSql & "                             ," & lngCodUsuario & vbCrLf                         '' Código do Usuário
                sSql = sSql & "                             ,'PPE'" & vbCrLf                                    '' Código da Ação
                sSql = sSql & "                             ,'" & arrPROGRAMADODEL(i, 20) & "'"                 '' Descrição da Ação
                sSql = sSql & "                             )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
                '' =========================================
            
                '' Removendo da Programação
                sSql = ""
                
                sSql = "Delete From SGI_CADMOVPCP" & strNOMFILIAL & vbCrLf
                sSql = sSql & "      Where" & vbCrLf
                sSql = sSql & "            SGI_FILIAL    = " & intFilial & vbCrLf
                ''sSql = sSql & "        And SGI_CODIGO    = " & lngCODIGO & vbCrLf
                sSql = sSql & "        And SGI_CODINTENO = " & arrPROGRAMADODEL(i, 10)
            
                BGRV.CommandText = sSql
                BGRV.Execute
                '' =========================================
            End If
        Next i
    End If
    
    '' ===============================
    '' Gravando
    If IsArray(arrPROGRAMADO) Then
        For i = 1 To UBound(arrPROGRAMADO)
            
            If arrPROGRAMADO(i, 8) = dacEnumUpdateAction_delete Then    '' Excluindo
            
                '' ===============================
                '' Update no OP
                sSql = ""
                
                sSql = "Update SGI_ORDEMPROD" & strNOMFILIAL & " Set" & vbCrLf
                
                sSql = sSql & "                                       SGI_PROGRAMADO = 0" & vbCrLf
                sSql = sSql & "                                      ,SGI_DATENTREGA = '" & Format(CDate(arrPROGRAMADO(i, 11)), "MM/DD/YYYY") & "'" & vbCrLf
                sSql = sSql & "                                      ,SGI_ALTERADO   = 1" & vbCrLf
                sSql = sSql & "                                      ,SGI_JAIMPRESSA = 0" & vbCrLf
                
                sSql = sSql & "Where " & vbCrLf
                sSql = sSql & "      SGI_FILIAl     = " & intFilial & vbCrLf
                sSql = sSql & "  And SGI_CODIGO     = " & arrPROGRAMADO(i, 20) & vbCrLf
                sSql = sSql & "  And SGI_IDPRODUTO  = " & arrPROGRAMADO(i, 18) & vbCrLf
                sSql = sSql & "  And SGI_IDPAI      = " & arrPROGRAMADO(i, 19)
                
                BGRV.CommandText = sSql
                BGRV.Execute
                '' ===============================
            
            
                '' ===============================
                '' Programação de Entrtega
                sSql = ""
                            
                sSql = sSql & "Update SGI_PROGENTRPROD" & strNOMFILIAL & " Set" & vbCrLf
                sSql = sSql & "                                                 SGI_DATENTREGA = '" & Format(CDate(arrPROGRAMADO(i, 11)), "MM/DD/YYYY") & "'" & vbCrLf
                
                sSql = sSql & "       Where" & vbCrLf
                sSql = sSql & "             SGI_FILIAl    = " & intFilial & vbCrLf
                sSql = sSql & "         And SGI_IDINTERNO = " & arrPROGRAMADO(i, 19) & vbCrLf
                sSql = sSql & "         And SGI_CODPED    = " & arrPROGRAMADO(i, 17) & vbCrLf
                sSql = sSql & "         And SGI_IDPRODUTO = " & arrPROGRAMADO(i, 18)
                
                BGRV.CommandText = sSql
                BGRV.Execute
                '' ===============================
            
                '' =========================================
                '' Grava Log Pedido
                sSql = ""
                
                sSql = "Insert Into SGI_LOGPED" & strNOMFILIAL & " (" & vbCrLf
                sSql = sSql & "                              SGI_FILIAL" & vbCrLf
                sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                             ,SGI_DATA" & vbCrLf
                sSql = sSql & "                             ,SGI_HORA" & vbCrLf
                sSql = sSql & "                             ,SGI_CODUSUARIO" & vbCrLf
                sSql = sSql & "                             ,SGI_ACAO" & vbCrLf
                sSql = sSql & "                             ,SGI_DESCACAO" & vbCrLf
                sSql = sSql & "                    ) Values (" & vbCrLf
                sSql = sSql & "                              " & intFilial & vbCrLf                             '' Filial
                sSql = sSql & "                             ," & arrPROGRAMADO(i, 17) & vbCrLf                  '' Código da Ped
                sSql = sSql & "                             ,'" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf      '' Data da Ação
                sSql = sSql & "                             ,'" & Format(Now, "HH:MM:SS") & "'" & vbCrLf        '' Hora da Ação
                sSql = sSql & "                             ," & lngCodUsuario & vbCrLf                         '' Código do Usuário
                sSql = sSql & "                             ,'PPE'" & vbCrLf                                    '' Código da Ação
                sSql = sSql & "                             ,'" & arrPROGRAMADO(i, 20) & "'"                    '' Descrição da Ação
                sSql = sSql & "                             )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
                '' =========================================
            
                '' Removendo da Programação
                sSql = ""
                
                sSql = "Delete From SGI_CADMOVPCP" & strNOMFILIAL & vbCrLf
                sSql = sSql & "      Where" & vbCrLf
                sSql = sSql & "            SGI_FILIAL    = " & intFilial & vbCrLf
                sSql = sSql & "        And SGI_CODIGO    = " & lngCODIGO & vbCrLf
                sSql = sSql & "        And SGI_CODINTENO = " & arrPROGRAMADO(i, 10)
            
                BGRV.CommandText = sSql
                BGRV.Execute
                '' =========================================
            
            ElseIf arrPROGRAMADO(i, 8) = dacEnumUpdateAction_update Then    '' Altarando
            
                '' ===============================
                '' Update no OP
                sSql = ""
                sSql = "Update SGI_ORDEMPROD" & strNOMFILIAL & " Set" & vbCrLf
                sSql = sSql & "                                       SGI_PROGRAMADO = 0" & vbCrLf
                sSql = sSql & "                                      ,SGI_DATENTREGA = '" & Format(CDate(arrPROGRAMADO(i, 3)), "MM/DD/YYYY") & "'" & vbCrLf
                sSql = sSql & "                                      ,SGI_ALTERADO   = 1" & vbCrLf
                sSql = sSql & "                                      ,SGI_JAIMPRESSA = 0" & vbCrLf
                sSql = sSql & "Where " & vbCrLf
                sSql = sSql & "      SGI_FILIAl     = " & intFilial & vbCrLf
                sSql = sSql & "  And SGI_CODIGO     = " & arrPROGRAMADO(i, 20) & vbCrLf
                sSql = sSql & "  And SGI_IDPRODUTO  = " & arrPROGRAMADO(i, 18) & vbCrLf
                sSql = sSql & "  And SGI_IDPAI      = " & arrPROGRAMADO(i, 19)
                
                BGRV.CommandText = sSql
                BGRV.Execute
                '' ===============================
            
                '' ===============================
                '' Programação de Entrtega
                sSql = ""
                sSql = sSql & "Update SGI_PROGENTRPROD" & strNOMFILIAL & " Set" & vbCrLf
                sSql = sSql & "                                                 SGI_DATENTREGA = '" & Format(CDate(arrPROGRAMADO(i, 3)), "MM/DD/YYYY") & "'" & vbCrLf
                sSql = sSql & "       Where" & vbCrLf
                sSql = sSql & "             SGI_FILIAl    = " & intFilial & vbCrLf
                sSql = sSql & "         And SGI_IDINTERNO = " & arrPROGRAMADO(i, 19) & vbCrLf
                sSql = sSql & "         And SGI_CODPED    = " & arrPROGRAMADO(i, 17) & vbCrLf
                sSql = sSql & "         And SGI_IDPRODUTO = " & arrPROGRAMADO(i, 18)
                
                BGRV.CommandText = sSql
                BGRV.Execute
                '' ===============================
            
                '' Removendo da Programação
                sSql = ""
                sSql = "Update SGI_CADMOVPCP" & strNOMFILIAL & " Set" & vbCrLf
                sSql = sSql & "                                              SGI_DATAENTR    = '" & Format(CDate(arrPROGRAMADO(i, 3)), "MM/DD/YYYY") & "'" & vbCrLf
                sSql = sSql & "                                             ,SGI_DATAENTRANT = '" & Format(CDate(arrPROGRAMADO(i, 11)), "MM/DD/YYYY") & "'" & vbCrLf
                sSql = sSql & "                                             ,SGI_ORDEM       = " & arrPROGRAMADO(i, 21) & vbCrLf
                sSql = sSql & "                                             ,SGI_QTDEPROD    = " & arrPROGRAMADO(i, 9) & vbCrLf
                sSql = sSql & "                                             ,SGI_DATAPROG    = '" & Format(arrPROGRAMADO(i, 1), "MM/DD/YYYY") & "'" & vbCrLf
                
                sSql = sSql & "      Where" & vbCrLf
                sSql = sSql & "            SGI_FILIAL    = " & intFilial & vbCrLf
                sSql = sSql & "        And SGI_CODIGO    = " & lngCODIGO & vbCrLf
                sSql = sSql & "        And SGI_CODINTENO = " & arrPROGRAMADO(i, 10)
            
                BGRV.CommandText = sSql
                BGRV.Execute
                '' =========================================
                
                '' =========================================
                '' Caso a Data de Entrega For Alterada
                If arrPROGRAMADO(i, 3) <> arrPROGRAMADO(i, 11) Then
                
                    '' =========================================
                    '' Grava Log Pedido
                    sSql = ""
                    
                    sSql = "Insert Into SGI_LOGPED" & strNOMFILIAL & " (" & vbCrLf
                    sSql = sSql & "                              SGI_FILIAL" & vbCrLf
                    sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
                    sSql = sSql & "                             ,SGI_DATA" & vbCrLf
                    sSql = sSql & "                             ,SGI_HORA" & vbCrLf
                    sSql = sSql & "                             ,SGI_CODUSUARIO" & vbCrLf
                    sSql = sSql & "                             ,SGI_ACAO" & vbCrLf
                    sSql = sSql & "                             ,SGI_DESCACAO" & vbCrLf
                    sSql = sSql & "                    ) Values (" & vbCrLf
                    sSql = sSql & "                              " & intFilial & vbCrLf                                             '' Filial
                    sSql = sSql & "                             ," & arrPROGRAMADO(i, 17) & vbCrLf                                  '' Código do Pedido
                    sSql = sSql & "                             ,'" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf                      '' Data da Ação
                    sSql = sSql & "                             ,'" & Format(Now, "HH:MM:SS") & "'" & vbCrLf                        '' Hora da Ação
                    sSql = sSql & "                             ," & lngCodUsuario & vbCrLf                                         '' Código do Usuário
                    sSql = sSql & "                             ,'POP'" & vbCrLf                                                    '' Código da Ação
                    sSql = sSql & "                             ,'" & arrPROGRAMADO(i, 11) & " Para " & arrPROGRAMADO(i, 3) & "'"   '' Descrição da Ação
                    sSql = sSql & "                             )"
                    
                    BGRV.CommandText = sSql
                    BGRV.Execute
                    '' =========================================
                
                End If
                '' =========================================
            
            ElseIf arrPROGRAMADO(i, 8) = dacEnumUpdateAction_Insert Then    '' Incluindo
            
''                If arrPROGRAMADO(i, 2) = 337012016 Then
''                    MsgBox "Parar"
''                End If
                
                sSql = ""
                
                sSql = "Insert Into SGI_CADMOVPCP" & strNOMFILIAL & " (" & vbCrLf
                sSql = sSql & "                               SGI_FILIAL" & vbCrLf
                sSql = sSql & "                              ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                              ,SGI_DATAPROG" & vbCrLf
                sSql = sSql & "                              ,SGI_CODOP" & vbCrLf
                sSql = sSql & "                              ,SGI_DATAENTR" & vbCrLf
                sSql = sSql & "                              ,SGI_CODSTATUS" & vbCrLf
                sSql = sSql & "                              ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                              ,SGI_IDINTERNO" & vbCrLf
                sSql = sSql & "                              ,SGI_CODPED" & vbCrLf
                sSql = sSql & "                              ,SGI_QTDEPROD" & vbCrLf
                sSql = sSql & "                              ,SGI_CODINTENO" & vbCrLf
                sSql = sSql & "                              ,SGI_DATAENTRANT" & vbCrLf
                sSql = sSql & "                              ,SGI_STATUSORIG" & vbCrLf
                sSql = sSql & "                              ,SGI_FRACIONADA" & vbCrLf
                sSql = sSql & "                              ,SGI_IDLINHA" & vbCrLf
                sSql = sSql & "                              ,SGI_CODLIN" & vbCrLf
                sSql = sSql & "                              ,SGI_CODGRPLIN" & vbCrLf
                sSql = sSql & "                              ,SGI_ORDEM" & vbCrLf
                
                sSql = sSql & "                     ) Values (" & vbCrLf
                sSql = sSql & "                               " & intFilial & vbCrLf                                                    '' Filial
                sSql = sSql & "                              ," & lngCODIGO & vbCrLf                                                    '' Código da Programação
                sSql = sSql & "                              ,'" & Format(CDate(arrPROGRAMADO(i, 1)), "MM/DD/YYYY") & "'" & vbCrLf      '' Data da Programação
                sSql = sSql & "                              ," & arrPROGRAMADO(i, 2) & vbCrLf                                          '' Código da OP
                sSql = sSql & "                              ,'" & Format(CDate(arrPROGRAMADO(i, 3)), "MM/DD/YYYY") & "'" & vbCrLf      '' Data da Entrega
                sSql = sSql & "                              ," & arrPROGRAMADO(i, 4) & vbCrLf                                          '' Código do Status
                sSql = sSql & "                              ," & arrPROGRAMADO(i, 5) & vbCrLf                                          '' ID do Produto
                sSql = sSql & "                              ," & arrPROGRAMADO(i, 6) & vbCrLf                                          '' ID do Pai
                sSql = sSql & "                              ," & arrPROGRAMADO(i, 7) & vbCrLf                                          '' Código do Pedido
                sSql = sSql & "                              ," & arrPROGRAMADO(i, 9) & vbCrLf                                          '' Qtde Real Programada
                sSql = sSql & "                              ," & arrPROGRAMADO(i, 10) & vbCrLf                                         '' Código de Controle Interno
                sSql = sSql & "                              ,'" & Format(CDate(arrPROGRAMADO(i, 11)), "MM/DD/YYYY") & "'" & vbCrLf     '' Data da Entrega Original
                sSql = sSql & "                              ," & arrPROGRAMADO(i, 12) & vbCrLf                                         '' Código do Status Original
                sSql = sSql & "                              ," & arrPROGRAMADO(i, 13) & vbCrLf                                         '' Se esta OP foi fracionada em Outro Dia
                sSql = sSql & "                              ," & arrPROGRAMADO(i, 14) & vbCrLf                                         '' Id da Linha
                sSql = sSql & "                              ," & arrPROGRAMADO(i, 15) & vbCrLf                                         '' Codigo da Linha
                sSql = sSql & "                              ," & arrPROGRAMADO(i, 16) & vbCrLf                                         '' Codigo do Grupo de Linha
                sSql = sSql & "                              ," & arrPROGRAMADO(i, 21) & vbCrLf                                         '' Ordem de Inserção
                
                sSql = sSql & "                               )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
                
                '' ===============================
                '' Update no OP
                sSql = ""
                
                sSql = "Update SGI_ORDEMPROD" & strNOMFILIAL & " Set" & vbCrLf
                
                sSql = sSql & "                                       SGI_PROGRAMADO = 1" & vbCrLf
                sSql = sSql & "                                      ,SGI_DATENTREGA = '" & Format(CDate(arrPROGRAMADO(i, 3)), "MM/DD/YYYY") & "'" & vbCrLf
                sSql = sSql & "                                      ,SGI_ALTERADO   = 1" & vbCrLf
                sSql = sSql & "                                      ,SGI_JAIMPRESSA = 0" & vbCrLf
                
                sSql = sSql & "Where " & vbCrLf
                sSql = sSql & "      SGI_FILIAl     = " & intFilial & vbCrLf
                sSql = sSql & "  And SGI_CODIGO     = " & arrPROGRAMADO(i, 2) & vbCrLf
                sSql = sSql & "  And SGI_IDPRODUTO  = " & arrPROGRAMADO(i, 5) & vbCrLf
                sSql = sSql & "  And SGI_IDPAI      = " & arrPROGRAMADO(i, 6)
                
                BGRV.CommandText = sSql
                BGRV.Execute
                '' ===============================
                
                '' ===============================
                '' Programação de Entrtega
                sSql = ""
                            
                sSql = sSql & "Update SGI_PROGENTRPROD" & strNOMFILIAL & " Set" & vbCrLf
                sSql = sSql & "                                                 SGI_DATENTREGA = '" & Format(CDate(arrPROGRAMADO(i, 3)), "MM/DD/YYYY") & "'" & vbCrLf
                
                sSql = sSql & "       Where" & vbCrLf
                sSql = sSql & "             SGI_FILIAl    = " & intFilial & vbCrLf
                sSql = sSql & "         And SGI_IDINTERNO = " & arrPROGRAMADO(i, 6) & vbCrLf
                sSql = sSql & "         And SGI_CODPED    = " & arrPROGRAMADO(i, 7) & vbCrLf
                sSql = sSql & "         And SGI_IDPRODUTO = " & arrPROGRAMADO(i, 5)
                
                BGRV.CommandText = sSql
                BGRV.Execute
                '' ===============================
            
                '' =========================================
                '' Grava Log Pedido
                sSql = ""
                
                sSql = "Insert Into SGI_LOGPED" & strNOMFILIAL & " (" & vbCrLf
                sSql = sSql & "                              SGI_FILIAL" & vbCrLf
                sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                             ,SGI_DATA" & vbCrLf
                sSql = sSql & "                             ,SGI_HORA" & vbCrLf
                sSql = sSql & "                             ,SGI_CODUSUARIO" & vbCrLf
                sSql = sSql & "                             ,SGI_ACAO" & vbCrLf
                sSql = sSql & "                             ,SGI_DESCACAO" & vbCrLf
                sSql = sSql & "                    ) Values (" & vbCrLf
                sSql = sSql & "                              " & intFilial & vbCrLf                             '' Filial
                sSql = sSql & "                             ," & arrPROGRAMADO(i, 7) & vbCrLf                   '' Código do Pedido
                sSql = sSql & "                             ,'" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf      '' Data da Ação
                sSql = sSql & "                             ,'" & Format(Now, "HH:MM:SS") & "'" & vbCrLf        '' Hora da Ação
                sSql = sSql & "                             ," & lngCodUsuario & vbCrLf                         '' Código do Usuário
                sSql = sSql & "                             ,'PPR'" & vbCrLf                                    '' Código da Ação
                sSql = sSql & "                             ,'" & arrPROGRAMADO(i, 2) & "'"                     '' Descrição da Ação
                sSql = sSql & "                             )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
                '' =========================================
                
                '' =========================================
                '' Caso a Data de Entrega For Alterada
                If arrPROGRAMADO(i, 3) <> arrPROGRAMADO(i, 11) Then
                
                    '' =========================================
                    '' Grava Log Pedido
                    sSql = ""
                    
                    sSql = "Insert Into SGI_LOGPED" & strNOMFILIAL & " (" & vbCrLf
                    sSql = sSql & "                              SGI_FILIAL" & vbCrLf
                    sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
                    sSql = sSql & "                             ,SGI_DATA" & vbCrLf
                    sSql = sSql & "                             ,SGI_HORA" & vbCrLf
                    sSql = sSql & "                             ,SGI_CODUSUARIO" & vbCrLf
                    sSql = sSql & "                             ,SGI_ACAO" & vbCrLf
                    sSql = sSql & "                             ,SGI_DESCACAO" & vbCrLf
                    sSql = sSql & "                    ) Values (" & vbCrLf
                    sSql = sSql & "                              " & intFilial & vbCrLf                                             '' Filial
                    sSql = sSql & "                             ," & arrPROGRAMADO(i, 17) & vbCrLf                                  '' Código do Pedido
                    sSql = sSql & "                             ,'" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf                      '' Data da Ação
                    sSql = sSql & "                             ,'" & Format(Now, "HH:MM:SS") & "'" & vbCrLf                        '' Hora da Ação
                    sSql = sSql & "                             ," & lngCodUsuario & vbCrLf                                         '' Código do Usuário
                    sSql = sSql & "                             ,'POP'" & vbCrLf                                                    '' Código da Ação
                    sSql = sSql & "                             ,'" & arrPROGRAMADO(i, 11) & " Para " & arrPROGRAMADO(i, 3) & "'"   '' Descrição da Ação
                    sSql = sSql & "                             )"
                    
                    BGRV.CommandText = sSql
                    BGRV.Execute
                    '' =========================================
                
                End If
                '' =========================================
                
                
                
                '' Liberando P.Cota/P.Data
                If arrPROGRAMADO(i, 4) = 7 Or arrPROGRAMADO(i, 4) = 6 Then
                    
                    '' =====================================
                    '' Programação de Entrega
                    sSql = ""
                    
                    sSql = "Update SGI_PROGENTRPROD" & strNOMFILIAL & " Set" & vbCrLf
                    sSql = sSql & "                                 SGI_STATUS     = 0" & vbCrLf '' Status
                    sSql = sSql & " Where" & vbCrLf
                    sSql = sSql & "             SGI_FILIAl    = " & intFilial & vbCrLf
                    sSql = sSql & "         And SGI_IDINTERNO = " & arrPROGRAMADO(i, 6) & vbCrLf
                    sSql = sSql & "         And SGI_CODPED    = " & arrPROGRAMADO(i, 7) & vbCrLf
                    sSql = sSql & "         And SGI_IDPRODUTO = " & arrPROGRAMADO(i, 5)
                    
                    BGRV.CommandText = sSql
                    BGRV.Execute
                
                    '' =====================================
                    '' Ordem de Produção
                    sSql = ""
                    
                    sSql = "Update SGI_ORDEMPROD" & strNOMFILIAL & " Set " & vbCrLf
                    sSql = sSql & "                          SGI_STATUS          = 0" & vbCrLf '' Status
                    
                    sSql = sSql & "                     Where " & vbCrLf
                    sSql = sSql & "      SGI_FILIAl     = " & intFilial & vbCrLf
                    sSql = sSql & "  And SGI_CODIGO     = " & arrPROGRAMADO(i, 2) & vbCrLf
                    sSql = sSql & "  And SGI_IDPRODUTO  = " & arrPROGRAMADO(i, 5) & vbCrLf
                    sSql = sSql & "  And SGI_IDPAI      = " & arrPROGRAMADO(i, 6)
                    
                    BGRV.CommandText = sSql
                    BGRV.Execute
                    
                    
                    '' =====================================
                    '' Ordem de Produção
                    sSql = ""
                    
                    sSql = "Update SGI_CADPEDVENDH" & strNOMFILIAL & " Set " & vbCrLf
                    sSql = sSql & "                          SGI_STATUS          = 'L'" & vbCrLf '' Status
                    
                    sSql = sSql & "                     Where " & vbCrLf
                    sSql = sSql & "      SGI_FILIAl     = " & intFilial & vbCrLf
                    sSql = sSql & "  And SGI_CODIGO     = " & arrPROGRAMADO(i, 7) & vbCrLf
                    
                    BGRV.CommandText = sSql
                    BGRV.Execute
                    
                    '' =====================================
                    '' Grava Log
                    sSql = ""
                    
                    sSql = "Insert Into SGI_LOGPED" & strNOMFILIAL & " (" & vbCrLf
                    sSql = sSql & "                              SGI_FILIAL" & vbCrLf
                    sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
                    sSql = sSql & "                             ,SGI_DATA" & vbCrLf
                    sSql = sSql & "                             ,SGI_HORA" & vbCrLf
                    sSql = sSql & "                             ,SGI_CODUSUARIO" & vbCrLf
                    sSql = sSql & "                             ,SGI_ACAO" & vbCrLf
                    sSql = sSql & "                    ) Values (" & vbCrLf
                    sSql = sSql & "                              " & intFilial & vbCrLf
                    sSql = sSql & "                             ," & arrPROGRAMADO(i, 7) & vbCrLf
                    sSql = sSql & "                             ,'" & Format(Now, "MM/DD/YYYY") & "'" & vbCrLf
                    sSql = sSql & "                             ,'" & Format(Now, "HH:MM:SS") & "'" & vbCrLf
                    sSql = sSql & "                             ," & lngCodUsuario & vbCrLf
                    sSql = sSql & "                             ,'LC'" & vbCrLf
                    sSql = sSql & "                             )"
                    
                    BGRV.CommandText = sSql
                    BGRV.Execute
                    '' =====================================
                End If
                
            End If
        Next i
    End If
    
    '' =====================================
    '' Incluindo as Folhas
    'If IsArray(arrGERAFOLHAS) Then
    '    For I = 1 To UBound(arrGERAFOLHAS)
    '
    '        sSql = ""
    '        sSql = sSql & "Insert Into SGI_CADMOVPCP_FOLHAS" & strNOMFILIAL & " (" & vbCrLf
    '        sSql = sSql & "                                SGI_FILIAL" & vbCrLf
    '        sSql = sSql & "                               ,SGI_CODIGO" & vbCrLf
    '        sSql = sSql & "                               ,SGI_IDPRODUTO" & vbCrLf
    '        sSql = sSql & "                               ,SGI_CODFOLHA" & vbCrLf
    '        sSql = sSql & "                               ,SGI_CODOP" & vbCrLf
    '        sSql = sSql & "                               ,SGI_INDICE" & vbCrLf
    '        sSql = sSql & "                      ) Values (" & vbCrLf
    '        sSql = sSql & "                                " & intFilial & vbCrLf
    '        sSql = sSql & "                               ," & lngCODIGO & vbCrLf
    '        sSql = sSql & "                               ," & arrGERAFOLHAS(I, 1) & vbCrLf
    '        sSql = sSql & "                               ," & arrGERAFOLHAS(I, 2) & vbCrLf
    '        sSql = sSql & "                               ," & arrGERAFOLHAS(I, 3) & vbCrLf
    '        sSql = sSql & "                               ," & arrGERAFOLHAS(I, 4) & vbCrLf
    '        sSql = sSql & "                               )"
    '
    '        BGRV.CommandText = sSql
    '        BGRV.Execute
    '
    '    Next I
    'End If
    '' =====================================
     
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

Public Function Carrega_Campos(strNomTab As String) As Boolean

    Carrega_Campos = False

    Dim i          As Integer
    Dim intQTDREG  As Integer
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADMOVPCP" & strNomTab & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCODIGO & vbCrLf
    sSql = sSql & "Order By SGI_ORDEM"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC.EOF Then
    
        strDTPROGRAMA = Format(BREC!SGI_DATAPROG, "DD/MM/YYYY")
    
        intQTDREG = 0
        Do While Not BREC.EOF()
            intQTDREG = (intQTDREG + 1)
            BREC.MoveNext
        Loop
        BREC.MoveFirst
        
        ReDim arrPROGRAMADO(1 To intQTDREG, 1 To 14) As String
        ReDim arrPROGRAMADOBKP(1 To intQTDREG, 1 To 14) As String
        intQTDREG = 0
        Do While Not BREC.EOF()
            intQTDREG = (intQTDREG + 1)
            
            arrPROGRAMADO(intQTDREG, 1) = BREC!SGI_CODINTENO
            arrPROGRAMADO(intQTDREG, 2) = BREC!SGI_CODOP
            arrPROGRAMADO(intQTDREG, 3) = Format(BREC!SGI_DATAENTR, "DD/MM/YYYY")
            arrPROGRAMADO(intQTDREG, 4) = BREC!SGI_CODSTATUS
            arrPROGRAMADO(intQTDREG, 5) = BREC!SGI_IDPRODUTO
            arrPROGRAMADO(intQTDREG, 6) = BREC!SGI_IDINTERNO
            arrPROGRAMADO(intQTDREG, 7) = BREC!SGI_CODPED
            arrPROGRAMADO(intQTDREG, 8) = BREC!SGI_QTDEPROD
            arrPROGRAMADO(intQTDREG, 9) = Format(BREC!SGI_DATAPROG, "DD/MM/YYYY")
            arrPROGRAMADO(intQTDREG, 10) = Format(BREC!SGI_DATAENTRANT, "DD/MM/YYYY")
            arrPROGRAMADO(intQTDREG, 11) = BREC!SGI_CODSTATUS
            arrPROGRAMADO(intQTDREG, 12) = BREC!SGI_STATUSORIG
            If Not IsNull(BREC!SGI_ORDEM) Then arrPROGRAMADO(intQTDREG, 13) = BREC!SGI_ORDEM
            If Not IsNull(BREC!SGI_TimeStamp) Then
                arrPROGRAMADO(intQTDREG, 14) = BREC!SGI_TimeStamp
            Else
                arrPROGRAMADO(intQTDREG, 14) = -1
            End If
            
            arrPROGRAMADOBKP(intQTDREG, 1) = BREC!SGI_CODINTENO
            arrPROGRAMADOBKP(intQTDREG, 2) = BREC!SGI_CODOP
            arrPROGRAMADOBKP(intQTDREG, 3) = "'" & Format(BREC!SGI_DATAENTR, "MM/DD/YYYY") & "'"
            arrPROGRAMADOBKP(intQTDREG, 4) = BREC!SGI_CODSTATUS
            arrPROGRAMADOBKP(intQTDREG, 5) = BREC!SGI_IDPRODUTO
            arrPROGRAMADOBKP(intQTDREG, 6) = BREC!SGI_IDINTERNO
            arrPROGRAMADOBKP(intQTDREG, 7) = BREC!SGI_CODPED
            arrPROGRAMADOBKP(intQTDREG, 8) = BREC!SGI_QTDEPROD
            arrPROGRAMADOBKP(intQTDREG, 9) = "'" & Format(BREC!SGI_DATAPROG, "MM/DD/YYYY") & "'"
            arrPROGRAMADOBKP(intQTDREG, 10) = "'" & Format(BREC!SGI_DATAENTRANT, "MM/DD/YYYY") & "'"
            arrPROGRAMADOBKP(intQTDREG, 11) = BREC!SGI_CODSTATUS
            arrPROGRAMADOBKP(intQTDREG, 12) = BREC!SGI_STATUSORIG
            If Not IsNull(BREC!SGI_ORDEM) Then arrPROGRAMADOBKP(intQTDREG, 13) = BREC!SGI_ORDEM
            If Not IsNull(BREC!SGI_TimeStamp) Then
                arrPROGRAMADOBKP(intQTDREG, 14) = BREC!SGI_TimeStamp
            Else
                arrPROGRAMADOBKP(intQTDREG, 14) = -1
            End If
            
            BREC.MoveNext
        Loop
        
        Carrega_Campos = True
    End If
    BREC.Close
    ' ----------------------------------------------------------
    
End Function

