VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADLISTMAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial           As Integer
Dim strCODIGO           As String
Dim arrDELETADOS        As Variant

Public Property Get DELETADOS() As Variant
    DELETADOS = arrDELETADOS
End Property

Public Property Let DELETADOS(ByVal vData As Variant)
    arrDELETADOS = vData
End Property

Public Property Get CODIGO() As String
    CODIGO = strCODIGO
End Property

Public Property Let CODIGO(ByVal vData As String)
    strCODIGO = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property
Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADLISTMATP.cCaminho = cCaminho
    frmCADLISTMATP.Linha = Linha
    frmCADLISTMATP.FILIAL = iFilial
    frmCADLISTMATP.strAcesso = strAcesso
    frmCADLISTMATP.Show vbModal
End Sub

Public Function Pesq_CadLstProd() As Boolean

   Pesq_CadLstProd = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_CadLstProd = True
   BREC.Close

End Function

Public Sub PreencheComboProd(Combo As Variant)

     Combo.Clear
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       PRO.* " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADTIPPROD  TIP " & vbCrLf
     sSql = sSql & "      ,SGI_CADPRODUTO  PRO " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       TIP.SGI_COMPLST = 1 "
     sSql = sSql & "   And TIP.SGI_FILIAL  = " & intFilial & vbCrLf
     sSql = sSql & "   And PRO.SGI_FILIAL  = TIP.SGI_FILIAL " & vbCrLf
     sSql = sSql & "   And PRO.SGI_CODTIPO = TIP.SGI_CODIGO " & vbCrLf
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
    Do While Not BREC.EOF()
       If BREC!SGI_PRODUTOTIPO = 1 And BREC!SGI_PRODUTOESTILO = 0 Then
          Combo.AddItem Format(IIf(IsNull(BREC!SGI_CODLINPROD), 0, BREC!SGI_CODLINPROD), "###000") & "." & _
                        Format(IIf(IsNull(BREC!SGI_CodClie), 0, BREC!SGI_CodClie), "####0000") & "." & _
                        Format(IIf(IsNull(BREC!SGI_CODROTULO), 0, BREC!SGI_CODROTULO), "##00") & "." & _
                        Format(IIf(IsNull(BREC!SGI_DIGVERIF), 0, BREC!SGI_DIGVERIF), "#0") & " - " & Trim(BREC!SGI_DESCRICAO)
       ElseIf BREC!SGI_PRODUTOTIPO = 0 And BREC!SGI_PRODUTOESTILO = 0 Then
          Combo.AddItem Trim(BREC!SGI_CODIGO) & Space(10 - Len(Trim(BREC!SGI_CODIGO))) & " - " & Trim(BREC!SGI_DESCRICAO)
       End If
       
       Combo.ItemData(Combo.NewIndex) = BREC!SGI_IDPRODUTO
       
       BREC.MoveNext
    Loop
    
    BREC.Close

    
End Sub

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

Public Function Atualiza(strAcao As String, strCODIGO As String, lngFilial_Atu As Integer, strModulo As String) As Boolean
    
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
        sSql = sSql & "                 ,'" & Trim(strCODIGO) & "'" & vbCrLf
        sSql = sSql & "                   )"
    
        BGRV.CommandText = sSql
        BGRV.Execute
    
    Else
    
        sSql = sSql & "Update SGI_ATUALIZA Set" & vbCrLf
        sSql = sSql & "           SGI_ACAO   = '" & Trim(strAcao) & "'" & vbCrLf
        sSql = sSql & "          ,SGI_CODIGO = '" & Trim(strCODIGO) & "'" & vbCrLf
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


Public Function GRAVA(strOPER As String) As Boolean
On Error GoTo err_Trans
     
     GRAVA = False
     
     Dim I            As Integer
     Dim lngIDPai     As Long
     Dim lngIDNOVO    As Long
     Dim strValor     As String
     Dim lngCODIGO    As Long
     
     '' Inicia transação
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Then
        
GRAVA_ALTERACAO:
        
        '' Incluindo lista de material
        For I = 0 To UBound(arrPROVARV)
         
            '' Alterando o Produto
            sSql = "Update SGI_CADPRODUTO Set SGI_TEMARVORE = 'S'" & vbCrLf
            sSql = sSql & " Where SGI_FILIAL    = " & intFilial & vbCrLf
            sSql = sSql & "   And SGI_IDPRODUTO = " & arrPROVARV(0).lngProdutoID
            
            BGRV.CommandText = sSql
            BGRV.Execute
            
            If Len(Trim(arrPROVARV(I).strPRODUTO)) > 0 Then
               
               sSql = "Insert Into SGI_LISTAMATPROD (" & vbCrLf
               sSql = sSql & "                   SGI_FILIAL" & vbCrLf
               sSql = sSql & "                  ,SGI_PRODUTO" & vbCrLf
               sSql = sSql & "                  ,SGI_PRODLST" & vbCrLf
               sSql = sSql & "                  ,SGI_QTDE" & vbCrLf
               sSql = sSql & "                  ,SGI_UNIDCONS" & vbCrLf
               sSql = sSql & "                  ,SGI_CODTABCONV" & vbCrLf
               sSql = sSql & "                  ,SGI_CUSTOUNIT" & vbCrLf
               sSql = sSql & "                  ,SGI_CUSTOTOTAL" & vbCrLf
               sSql = sSql & "                  ,SGI_CODIGO" & vbCrLf
               sSql = sSql & "                  ,SGI_CODPAI" & vbCrLf
               sSql = sSql & "                  ,SGI_IDPRODUTO" & vbCrLf
               sSql = sSql & "                  ,SGI_IDPRODLST" & vbCrLf
               sSql = sSql & "                  ,SGI_CODUNIMED" & vbCrLf
               
               sSql = sSql & "                  )" & vbCrLf
               
               sSql = sSql & "           Values (" & vbCrLf
               sSql = sSql & "                   " & intFilial & vbCrLf
               sSql = sSql & "                 ,'" & arrPROVARV(I).strProdutoPAI & "'" & vbCrLf
               sSql = sSql & "                 ,'" & arrPROVARV(I).strPRODUTO & "'" & vbCrLf
             
               strValor = Replace(arrPROVARV(I).curQTDCONS, ".", "")
               strValor = Replace(strValor, ",", ".")
               sSql = sSql & "                  ," & Trim(strValor) & vbCrLf
               sSql = sSql & "                  ,'" & Trim(arrPROVARV(I).strUNIDADE) & "'" & vbCrLf
               sSql = sSql & "                  ,Null" & vbCrLf
            
               sSql = sSql & "                  ,0" & vbCrLf
               sSql = sSql & "                  ,0" & vbCrLf
               
               sSql = sSql & "                  ," & arrPROVARV(I).lngCODIGO & vbCrLf
               sSql = sSql & "                  ," & arrPROVARV(I).lngCODPAI & vbCrLf
               
               sSql = sSql & "                  ," & arrPROVARV(I).lngProdutoID & vbCrLf
               sSql = sSql & "                  ," & arrPROVARV(I).lngProdutoIDPai & vbCrLf
               sSql = sSql & "                  ," & arrPROVARV(I).lngCodUniMed & vbCrLf
               
               sSql = sSql & "                   )"
            
               BGRV.CommandText = sSql
               BGRV.Execute
               
            End If
            
        Next I
           
     ElseIf strOPER = "E" Or strOPER = "A" Then
     
        If IsArray(arrDELETADOS) Then
           For I = 0 To UBound(arrDELETADOS)
               If Len(Trim(arrDELETADOS(I))) > 0 Then
                  sSql = "Delete from SGI_LISTAMAT " & vbCrLf
                  sSql = sSql & "       Where SGI_CODIGO = " & CLng(Trim(arrDELETADOS(I))) & vbCrLf
                  sSql = sSql & "         And SGI_FILIAL  = " & intFilial
                
                  BGRV.CommandText = sSql
                  BGRV.Execute
               End If
           Next I
        End If
        
        '' Alterando o Produto
        sSql = "Update SGI_CADPRODUTO Set SGI_TEMARVORE = 'N'" & vbCrLf
        sSql = sSql & " Where SGI_FILIAL    = " & intFilial & vbCrLf
        sSql = sSql & "   And SGI_IDPRODUTO = " & arrPROVARV(0).lngProdutoID
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------
        '' Apaga Lista de material
        For I = 0 To UBound(arrPROVARV)
            sSql = "Delete from SGI_LISTAMATPROD " & vbCrLf
            sSql = sSql & "       Where SGI_CODIGO = " & CLng(Trim(arrPROVARV(I).lngCODIGO)) & vbCrLf
            sSql = sSql & "         And SGI_FILIAL  = " & intFilial
            
            BGRV.CommandText = sSql
            BGRV.Execute
        Next I
        
        If strOPER = "A" Then GoTo GRAVA_ALTERACAO
        
     ElseIf strOPER = "AL" Then
        
        '' --------------------------
        '' Apaga Lista de material
        For I = 0 To UBound(arrPROVARV)
            sSql = "Delete from SGI_LISTAMATPROD " & vbCrLf
            sSql = sSql & "       Where SGI_PRODUTO = '" & Trim(arrPROVARV(I).strPRODUTO) & "'" & vbCrLf
            sSql = sSql & "         And SGI_FILIAL  = " & intFilial
            
            BGRV.CommandText = sSql
            BGRV.Execute
        Next I
        
        '' Incluindo lista de material
        For I = 0 To UBound(arrPROVARV)
         
            If Len(Trim(arrPROVARV(I).strPRODUTO)) > 0 Then
            
               sSql = "Insert Into SGI_LISTAMATPROD (" & vbCrLf
               sSql = sSql & "                   SGI_FILIAL" & vbCrLf
               sSql = sSql & "                  ,SGI_PRODUTO" & vbCrLf
               sSql = sSql & "                  ,SGI_PRODLST" & vbCrLf
               sSql = sSql & "                  ,SGI_QTDE" & vbCrLf
               sSql = sSql & "                  ,SGI_UNIDCONS" & vbCrLf
               sSql = sSql & "                  ,SGI_CODTABCONV" & vbCrLf
               sSql = sSql & "                  ,SGI_CUSTOUNIT" & vbCrLf
               sSql = sSql & "                  ,SGI_CUSTOTOTAL" & vbCrLf
               sSql = sSql & "                  ,SGI_CODIGO" & vbCrLf
               sSql = sSql & "                  ,SGI_CODPAI)" & vbCrLf
               
               sSql = sSql & "           Values (" & vbCrLf
               sSql = sSql & "                   " & intFilial & vbCrLf
               sSql = sSql & "                 ,'" & arrPROVARV(I).strProdutoPAI & "'" & vbCrLf
               sSql = sSql & "                 ,'" & arrPROVARV(I).strPRODUTO & "'" & vbCrLf
             
               strValor = Replace(arrPROVARV(I).curQTDCONS, ".", "")
               strValor = Replace(strValor, ",", ".")
               sSql = sSql & "                  ," & Trim(strValor) & vbCrLf
               sSql = sSql & "                  ,'" & Trim(arrPROVARV(I).strUNIDADE) & "'" & vbCrLf
               sSql = sSql & "                  ,Null" & vbCrLf
            
               sSql = sSql & "                  ,0" & vbCrLf
               sSql = sSql & "                  ,0" & vbCrLf
               
               sSql = sSql & "                  ," & arrPROVARV(I).lngCODIGO & vbCrLf
               sSql = sSql & "                  ," & arrPROVARV(I).lngCODPAI & vbCrLf
               
               sSql = sSql & "                   )"
            
               BGRV.CommandText = sSql
               BGRV.Execute
               
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
    
    BREC12.Open sSql, adoBanco_Dados
    
    If Not BREC12.EOF Then
    
       If IsNull(BREC12!SGI_NUMERO) = True Then
          
          Gera_Codigo = 1
          
          sSql = "Insert into SGI_NUMERO (SGI_FILIAL,SGI_NUMERO,SGI_MODULO) Values(" & vbCrLf
          sSql = sSql + "                                              " & intFilial & vbCrLf
          sSql = sSql + "                                            ,1" & vbCrLf
          sSql = sSql + "                                            ,'" & sModulo & "'" & vbCrLf
          sSql = sSql + "                                          )" & vbCrLf
          
       ElseIf BREC12!SGI_NUMERO > 1 Then
       
          Gera_Codigo = BREC12!SGI_NUMERO
          
          sSql = "Update SGI_NUMERO Set " & vbCrLf
          sSql = sSql + "           SGI_NUMERO = " & BREC12!SGI_NUMERO & vbCrLf
          sSql = sSql + "         Where " & vbCrLf
          sSql = sSql + "               SGI_MODULO = '" & sModulo & "'" & vbCrLf
          sSql = sSql + "           And SGI_FILIAL =  " & intFilial
       
       End If
       
       BGRV.CommandText = sSql
       BGRV.Execute
       
    End If
    
    BREC12.Close
    
End Function

