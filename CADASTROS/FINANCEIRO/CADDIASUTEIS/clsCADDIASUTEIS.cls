VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADDIASUTEIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial    As Integer
Dim intCODIGO    As Integer
Dim strDESCRI    As String
Dim dtDIASEVENTO As Date
Dim dtDATAFINAL  As Date
Dim lngTOTDIAS   As Long
Dim lngTOTHORAS  As Long
Dim lngTOTMIN    As Long

Public Property Get TOTMIN() As Long
    TOTMIN = lngTOTMIN
End Property

Public Property Let TOTMIN(ByVal vData As Long)
    lngTOTMIN = vData
End Property

Public Property Get TOTHORAS() As Long
    TOTHORAS = lngTOTHORAS
End Property

Public Property Let TOTHORAS(ByVal vData As Long)
    lngTOTHORAS = vData
End Property

Public Property Get totDias() As Long
    totDias = lngTOTDIAS
End Property

Public Property Let totDias(ByVal vData As Long)
    lngTOTDIAS = vData
End Property

Public Property Get DATAFINAL() As Date
    DATAFINAL = dtDATAFINAL
End Property

Public Property Let DATAFINAL(ByVal vData As Date)
    dtDATAFINAL = vData
End Property

Public Property Get DIASEVENTO() As Date
    DIASEVENTO = dtDIASEVENTO
End Property

Public Property Let DIASEVENTO(ByVal vData As Date)
    dtDIASEVENTO = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Property Get CODIGO() As Integer
    CODIGO = intCODIGO
End Property

Public Property Let CODIGO(ByVal vData As Integer)
    intCODIGO = vData
End Property

Public Property Get DESCRI() As String
    DESCRI = strDESCRI
End Property

Public Property Let DESCRI(ByVal vData As String)
    strDESCRI = vData
End Property


Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADDIASUTEISP.cCaminho = cCaminho
    frmCADDIASUTEISP.Linha = Linha
    frmCADDIASUTEISP.FILIAL = iFilial
    frmCADDIASUTEISP.strAcesso = strAcesso
    frmCADDIASUTEISP.Show vbModal
End Sub

Public Function Pesq_Banco() As Boolean

   Pesq_Banco = False
   
   sSql = ""
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADDIASUTEIS " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_Banco = True
   BREC.Close

End Function


Public Function Gera_Codigo(sModulo As String) As Integer

    Gera_Codigo = 1
    
    BGRV.ActiveConnection = adoBanco_Dados
    
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql + "       (Max(SGI_NUMERO) + 1) As SGI_NUMERO " & vbCrLf
    sSql = sSql + "  From " & vbCrLf
    sSql = sSql + "       SGI_NUMERO " & vbCrLf
    sSql = sSql + " Where " & vbCrLf
    sSql = sSql + "       SGI_MODULO = '" & sModulo & "'"
    sSql = sSql + "  And  SGI_FILIAL =  " & intFilial
    
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
         
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Then
    
        '' Grava Eventos
        sSql = "Insert Into SGI_CADDIASUTEIS (SGI_FILIAL" & vbCrLf
        sSql = sSql & "                      ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                      ,SGI_DESCRI" & vbCrLf
        sSql = sSql & "                      ,SGI_DATENVENTO" & vbCrLf
        sSql = sSql & "                      ,SGI_DATAFINAL" & vbCrLf
        sSql = sSql & "                      ,SGI_TOTDIAS" & vbCrLf
        sSql = sSql & "                      ,SGI_TOTHORAS" & vbCrLf
        sSql = sSql & "                      ,SGI_TOTMIN" & vbCrLf
        sSql = sSql & "                       ) Values (" & vbCrLf
        sSql = sSql & "                                 " & intFilial & vbCrLf
        sSql = sSql & "                                ," & intCODIGO & vbCrLf
        sSql = sSql & "                               ,'" & strDESCRI & "'" & vbCrLf
        sSql = sSql & "                               ,'" & Format(dtDIASEVENTO, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                               ,'" & Format(dtDATAFINAL, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                                ," & lngTOTDIAS & vbCrLf
        sSql = sSql & "                                ," & lngTOTHORAS & vbCrLf
        sSql = sSql & "                                ," & lngTOTMIN & vbCrLf
        sSql = sSql & "                                )"
     
     ElseIf strOPER = "A" Then
     
        '' Grava Eventos
        sSql = "Update SGI_CADDIASUTEIS Set " & vbCrLf
        sSql = sSql & "                     SGI_DESCRI     = '" & strDESCRI & "'" & vbCrLf
        sSql = sSql & "                    ,SGI_DATENVENTO = '" & Format(dtDIASEVENTO, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                    ,SGI_DATAFINAL  = '" & Format(dtDATAFINAL, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                    ,SGI_TOTDIAS    = " & lngTOTDIAS & vbCrLf
        sSql = sSql & "                    ,SGI_TOTHORAS   = " & lngTOTHORAS & vbCrLf
        sSql = sSql & "                    ,SGI_TOTMIN     = " & lngTOTMIN & vbCrLf
        sSql = sSql & "              Where " & vbCrLf
        sSql = sSql & "                    SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "                And SGI_CODIGO = " & intCODIGO
     
     ElseIf strOPER = "E" Then
     
        sSql = "Delete from SGI_CADDIASUTEIS " & vbCrLf
        sSql = sSql & "      Where SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCODIGO
     
     End If
     
     BGRV.CommandText = sSql
     BGRV.Execute
     
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

   Carrega_campos = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       *" & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_CADDIASUTEIS " & vbCrLf
   sSql = sSql + " Where " & vbCrLf
   sSql = sSql + "       SGI_FILIAL   = " & intFilial
   sSql = sSql + "   And SGI_CODIGO   = " & intCODIGO
   
   BREC.Open sSql, adoBanco_Dados
   
   If Not BREC.EOF Then
      
      intCODIGO = BREC!SGI_CODIGO
      strDESCRI = BREC!SGI_DESCRI
      dtDIASEVENTO = BREC!SGI_DATENVENTO
      dtDATAFINAL = BREC!SGI_DATAFINAL
   
      Carrega_campos = True
   End If
   
   BREC.Close

End Function

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

