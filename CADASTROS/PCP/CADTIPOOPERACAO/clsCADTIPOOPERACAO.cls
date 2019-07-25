VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADTIPOOPERACAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial       As Integer
Dim lngCODIGO       As Long
Dim strDESCRI       As String
Dim strCODFAMMAQ    As String

Public Property Get CODFAMMAQ() As String
    CODFAMMAQ = strCODFAMMAQ
End Property

Public Property Let CODFAMMAQ(ByVal vData As String)
    strCODFAMMAQ = vData
End Property

Public Property Get DESCRI() As String
    DESCRI = strDESCRI
End Property

Public Property Let DESCRI(ByVal vData As String)
    strDESCRI = vData
End Property

Public Property Get Filial() As Integer
    Filial = intFilial
End Property
Public Property Let Filial(ByVal vData As Integer)
    intFilial = vData
End Property
Public Property Get CODIGO() As Long
    CODIGO = lngCODIGO
End Property
Public Property Let CODIGO(ByVal vData As Long)
    lngCODIGO = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADTIPOOPERACAOP.cCaminho = cCaminho
    frmCADTIPOOPERACAOP.Linha = Linha
    frmCADTIPOOPERACAOP.Filial = iFilial
    frmCADTIPOOPERACAOP.strAcesso = strAcesso
    frmCADTIPOOPERACAOP.strUSUARIO = strUSUARIO
    frmCADTIPOOPERACAOP.Show vbModal
End Sub

Public Function Pesq_CadTipOper() As Boolean

   Pesq_CadTipOper = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_TIPOPERACAO " & vbCrLf
   sSql = sSql & " Where "
   sSql = sSql & "       SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_CadTipOper = True
   BREC.Close

End Function

Public Function Carrega_campos() As Boolean

   Carrega_campos = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_TIPOPERACAO " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & lngCODIGO & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then
      lngCODIGO = BREC!SGI_CODIGO
      strDESCRI = BREC!SGI_DESCRI
      If Not IsNull(BREC!SGI_CODFAMMAQ) Then strCODFAMMAQ = Trim(Str(BREC!SGI_CODFAMMAQ))
      Carrega_campos = True
   End If
   BREC.Close

End Function

Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_grava

     GRAVA = False
         
     Dim I As Integer
     
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Then
        
        sSql = "Insert Into SGI_TIPOPERACAO (SGI_FILIAL,SGI_CODIGO,SGI_DESCRI,SGI_CODFAMMAQ) Values (" & vbCrLf
        sSql = sSql & "                            " & intFilial
        sSql = sSql & "                           ," & lngCODIGO
        sSql = sSql & "                           ,'" & strDESCRI & "'" & vbCrLf
        sSql = sSql & "                           ," & Trim(strCODFAMMAQ) & vbCrLf
        sSql = sSql & "                           )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
     
     ElseIf strOPER = "A" Then
     
        sSql = "Update SGI_TIPOPERACAO Set " & vbCrLf
        sSql = sSql & "                    SGI_DESCRI    = '" & Trim(strDESCRI) & "'" & vbCrLf
        sSql = sSql & "                   ,SGI_CODFAMMAQ = " & Trim(strCODFAMMAQ) & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & lngCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
     
     ElseIf strOPER = "E" Then
     
        sSql = "Delete From SGI_TIPOPERACAO " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & lngCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
     End If
     
     adoBanco_Dados.CommitTrans
     
     GRAVA = True
     
     Exit Function
     
err_grava:
     
     adoBanco_Dados.RollbackTrans
    
     Dim objErro    As Object
     Set objErro = CreateObject("BLBCWS.clsFuncoes")
     Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, strOPER, sSql)
     Set objErro = Nothing
     
End Function

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

