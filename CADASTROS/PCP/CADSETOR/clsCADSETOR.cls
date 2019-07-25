VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADSETOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial       As Integer
Dim intCODIGO       As Integer
Dim strDESCRI       As String
Dim arrSECAO        As Variant

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

Public Property Get SECAO() As Variant
    SECAO = arrSECAO
End Property

Public Property Let SECAO(ByVal vData As Variant)
    arrSECAO = vData
End Property

Public Function Carrega_CadSetor() As Boolean

   Carrega_CadSetor = False
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADSETOR " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Carrega_CadSetor = True
   BREC.Close

End Function

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADSETORP.cCaminho = cCaminho
    frmCADSETORP.Linha = Linha
    frmCADSETORP.FILIAL = iFilial
    frmCADSETORP.strAcesso = strAcesso
    frmCADSETORP.strUSUARIO = strUSUARIO
    frmCADSETORP.Show vbModal
End Sub

Public Sub PreencheComboSecao(Combo As Variant)

     Combo.Clear
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       * " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADSECAO " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL  = " & intFilial & vbCrLf
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_DESCRI)
        Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
        BREC.MoveNext
     Loop
    
     BREC.Close
    
End Sub

Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_Trans
     
     GRAVA = False
         
     Dim I As Integer
     
     '' Inicia transação
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Then
        
ponto_volta:
        
        '' Grava Sub-Seção
        sSql = "Insert Into SGI_CADSETOR (SGI_FILIAL,SGI_CODIGO,SGI_DESCRI) Values (" & vbCrLf
        sSql = sSql & "                   " & intFilial
        sSql = sSql & "                  ," & intCODIGO
        sSql = sSql & "                  ,'" & Trim(strDESCRI) & "'" & vbCrLf
        sSql = sSql & "                   )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Grava os itens da Sub-Seção
        If IsArray(arrSECAO) = True Then
           For I = 1 To UBound(arrSECAO)
               
               sSql = "Insert Into SGI_CADITESET (SGI_FILIAL,SGI_CODIGO,SGI_CODSECAO) Values (" & vbCrLf
               sSql = sSql & "                    " & intFilial
               sSql = sSql & "                   ," & intCODIGO
               sSql = sSql & "                   ," & arrSECAO(I) & vbCrLf
               sSql = sSql & "                   )"
           
               BGRV.CommandText = sSql
               BGRV.Execute
           
           Next I
           
        End If
     
     ElseIf (strOPER = "A") Or (strOPER = "E") Then
     
        '' Apaga Itens da Setor
        sSql = "Delete from SGI_CADITESET " & vbCrLf
        sSql = sSql & "       Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "         And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Apaga Setor
        sSql = "Delete from SGI_CADSETOR " & vbCrLf
        sSql = sSql & "       Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "         And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        If strOPER = "A" Then GoTo ponto_volta
        
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
   
   Dim intReg As Integer
   Dim I      As Integer
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADSETOR " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & intCODIGO & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC.EOF Then
      
      intCODIGO = BREC!SGI_CODIGO
      strDESCRI = BREC!SGI_DESCRI
      
      Carrega_campos = True
   End If
   
   BREC.Close
   
   If Carrega_campos = False Then Exit Function
   
   '' ---------------
   
   BREC2.ActiveConnection = adoBanco_Dados
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADITESET " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial
   sSql = sSql & "   And SGI_CODIGO = " & intCODIGO & vbCrLf
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC2.EOF Then
   
      intReg = 0
      Do While Not BREC2.EOF
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
      
      ReDim arrSECAO(1 To intReg)
      
      intReg = 1
      BREC2.MoveFirst
      Do While Not BREC2.EOF
         arrSECAO(intReg) = BREC2!SGI_CODSECAO
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
   
   End If
   
   BREC2.Close

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

