VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADFERRA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial    As Long
Dim intCODIGO    As Long
Dim strDESCRI    As String
Dim curQTDEST    As Currency
Dim curESTMIN    As Currency
Dim curCAPACI    As Currency
Dim strARMARIO   As String
Dim strPRATELE   As String
Dim strBOXCAIXA  As String
Dim intUNIDADE   As Integer

Public Property Get UNIDADE() As Integer
    UNIDADE = intUNIDADE
End Property

Public Property Let UNIDADE(ByVal vData As Integer)
    intUNIDADE = vData
End Property

Public Property Get BOXCAIXA() As String
    BOXCAIXA = strBOXCAIXA
End Property

Public Property Let BOXCAIXA(ByVal vData As String)
    strBOXCAIXA = vData
End Property

Public Property Get PRATELE() As String
    PRATELE = strPRATELE
End Property

Public Property Let PRATELE(ByVal vData As String)
    strPRATELE = vData
End Property

Public Property Get ARMARIO() As String
    ARMARIO = strARMARIO
End Property

Public Property Let ARMARIO(ByVal vData As String)
    strARMARIO = vData
End Property

Public Property Get CAPACI() As Currency
    CAPACI = curCAPACI
End Property

Public Property Let CAPACI(ByVal vData As Currency)
    curCAPACI = vData
End Property

Public Property Get ESTMIN() As Currency
    ESTMIN = curESTMIN
End Property

Public Property Let ESTMIN(ByVal vData As Currency)
    curESTMIN = vData
End Property

Public Property Get QTDEST() As Currency
    QTDEST = curQTDEST
End Property

Public Property Let QTDEST(ByVal vData As Currency)
    curQTDEST = vData
End Property

Public Property Get DESCRI() As String
    DESCRI = strDESCRI
End Property

Public Property Let DESCRI(ByVal vData As String)
    strDESCRI = vData
End Property

Public Property Get CODIGO() As Long
    CODIGO = intCODIGO
End Property

Public Property Let CODIGO(ByVal vData As Long)
    intCODIGO = vData
End Property

Public Property Get Filial() As Long
    Filial = intFilial
End Property

Public Property Let Filial(ByVal vData As Long)
    intFilial = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADFERRAP.cCaminho = cCaminho
    frmCADFERRAP.Linha = Linha
    frmCADFERRAP.Filial = iFilial
    frmCADFERRAP.strAcesso = strAcesso
    frmCADFERRAP.Show vbModal
End Sub

Public Function Pesq_CadFerramenta() As Boolean

   Pesq_CadFerramenta = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       *" & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_CADFERRA " & vbCrLf
   sSql = sSql + " Where "
   sSql = sSql + "       SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_CadFerramenta = True
   BREC.Close

End Function

Public Sub PreenchComboUnidade(Combo As Variant)

    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADUNIMED " & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "Order by SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       Combo.AddItem Trim(BREC!SGI_UNIDADE)
       Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
       BREC.MoveNext
    Loop
    
    BREC.Close

End Sub

Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_grava
     
     Dim strVALOR As String
     
     GRAVA = False
         
     BGRV.ActiveConnection = adoBanco_Dados
     adoBanco_Dados.BeginTrans
     
     If strOPER = "I" Then
        
ALTERACAO:

        '' Grava especificação técnica do produto
        sSql = "Insert Into SGI_CADFERRA (SGI_FILIAL" & vbCrLf
        sSql = sSql & "                  ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                  ,SGI_DESCRI" & vbCrLf
        sSql = sSql & "                  ,SGI_QTDEST" & vbCrLf
        sSql = sSql & "                  ,SGI_QTDESTMIN" & vbCrLf
        sSql = sSql & "                  ,SGI_QTDPROD" & vbCrLf
        sSql = sSql & "                  ,SGI_ARMARIO" & vbCrLf
        sSql = sSql & "                  ,SGI_PRATELERIA" & vbCrLf
        sSql = sSql & "                  ,SGI_BOXCAIXA" & vbCrLf
        sSql = sSql & "                  ,SGI_CODUNIMED" & vbCrLf
        sSql = sSql & "                   ) Values (" & vbCrLf
        sSql = sSql & "                            " & intFilial & vbCrLf
        sSql = sSql & "                           ," & intCODIGO & vbCrLf
        sSql = sSql & "                           ,'" & strDESCRI & "'" & vbCrLf
        
        strVALOR = Replace(Format(curQTDEST, "#,###0.000"), ".", "")
        strVALOR = Replace(strVALOR, ",", ".")
        sSql = sSql & "                           ," & Trim(strVALOR) & vbCrLf
        
        strVALOR = Replace(Format(curESTMIN, "#,###0.000"), ".", "")
        strVALOR = Replace(strVALOR, ",", ".")
        sSql = sSql & "                           ," & Trim(strVALOR) & vbCrLf
        
        strVALOR = Replace(Format(curCAPACI, "#,###0.000"), ".", "")
        strVALOR = Replace(strVALOR, ",", ".")
        sSql = sSql & "                           ," & Trim(strVALOR) & vbCrLf
        
        sSql = sSql & "                           ,'" & strARMARIO & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strPRATELE & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strBOXCAIXA & "'" & vbCrLf
        sSql = sSql & "                           ," & intUNIDADE & vbCrLf
        
        sSql = sSql & "                           )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
     
     ElseIf strOPER = "A" Or strOPER = "E" Then
     
        '' Exclui Ferramenta
        sSql = "Delete from SGI_CADFERRA " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        If strOPER = "A" Then GoTo ALTERACAO
     
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

Public Function Carrega_campos() As Boolean

   Carrega_campos = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADFERRA " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & intCODIGO & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   
   If Not BREC.EOF Then
      
      strDESCRI = BREC!SGI_DESCRI
      curQTDEST = BREC!SGI_QTDEST
      curESTMIN = BREC!SGI_QTDESTMIN
      curCAPACI = BREC!SGI_QTDPROD
      strARMARIO = BREC!SGI_ARMARIO
      strPRATELE = BREC!SGI_PRATELERIA
      strBOXCAIXA = BREC!SGI_BOXCAIXA
      intUNIDADE = BREC!SGI_CODUNIMED
      
      Carrega_campos = True
   End If
   
   BREC.Close

End Function

Public Function Atualiza(strAcao As String, lngCODIGO As Long, lngFilial_Atu As Long, strModulo As String) As Boolean
    
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

