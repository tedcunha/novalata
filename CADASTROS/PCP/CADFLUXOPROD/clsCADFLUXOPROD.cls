VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADFLUXOPROD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial       As Integer
Dim intCODIGO       As Integer
Dim strCODPROD      As String
Dim strIDProduto    As String
Dim arrPROCESSOS    As Variant
Dim curMELHOR       As Currency
Dim curPIOR         As Currency
Dim curCORRMELHOR   As Currency
Dim curCORRPIOR     As Currency

Public Property Get IDPRODUTO() As String
    IDPRODUTO = strIDProduto
End Property
Public Property Let IDPRODUTO(ByVal vData As String)
    strIDProduto = vData
End Property

Public Property Get CORRPIOR() As Currency
    CORRPIOR = curCORRPIOR
End Property
Public Property Let CORRPIOR(ByVal vData As Currency)
    curCORRPIOR = vData
End Property

Public Property Get CORRMELHOR() As Currency
    CORRMELHOR = curCORRMELHOR
End Property
Public Property Let CORRMELHOR(ByVal vData As Currency)
    curCORRMELHOR = vData
End Property

Public Property Get PIOR() As Currency
    PIOR = curPIOR
End Property
Public Property Let PIOR(ByVal vData As Currency)
    curPIOR = vData
End Property

Public Property Get MELHOR() As Currency
    MELHOR = curMELHOR
End Property
Public Property Let MELHOR(ByVal vData As Currency)
    curMELHOR = vData
End Property

Public Property Get Processos() As Variant
    Processos = arrPROCESSOS
End Property
Public Property Let Processos(ByVal vData As Variant)
    arrPROCESSOS = vData
End Property

Public Property Get CODPROD() As String
    CODPROD = strCODPROD
End Property
Public Property Let CODPROD(ByVal vData As String)
    strCODPROD = vData
End Property

Public Property Get CODIGO() As Integer
    CODIGO = intCODIGO
End Property
Public Property Let CODIGO(ByVal vData As Integer)
    intCODIGO = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property
Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADFLUXOPRODP.cCaminho = cCaminho
    frmCADFLUXOPRODP.Linha = Linha
    frmCADFLUXOPRODP.FILIAL = iFilial
    frmCADFLUXOPRODP.strAcesso = strAcesso
    frmCADFLUXOPRODP.Show vbModal
End Sub

Public Function Pesq_CadFlxProd() As Boolean

   Pesq_CadFlxProd = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADFLUXPROD " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_CadFlxProd = True
   BREC.Close

End Function

Public Function PreenchComboUnidade() As String
        
    PreenchComboUnidade = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADUNIMED " & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "Order by SGI_CODIGO"
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC.EOF
       PreenchComboUnidade = PreenchComboUnidade & "|#" & BREC!SGI_CODIGO & ";" & Trim(BREC!SGI_UNIDADE)
       BREC.MoveNext
    Loop
    BREC.Close
End Function

Public Sub PreenchComboProcesso(Combo As Variant)

    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADPROCESSO " & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "Order by SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       Combo.AddItem Trim(BREC!SGI_DESCRI)
       Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
       BREC.MoveNext
    Loop
    
    BREC.Close

End Sub

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

Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_grava

     GRAVA = False
     
     Dim strValor As String
     Dim I        As Integer
     Dim j        As Integer
     Dim k        As Integer
         
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     If strOPER = "I" Then
     
VOLTA:
        '' Grava Fluxo Produtivo
        sSql = "Insert Into SGI_CADFLUXPROD (SGI_FILIAL"
        sSql = sSql & "                     ,SGI_CODIGO"
        sSql = sSql & "                     ,SGI_IDPRODUTO"
        sSql = sSql & "                     ,SGI_CODPROD"
        sSql = sSql & "                     ,SGI_MELHOR"
        sSql = sSql & "                     ,SGI_PIOR"
        sSql = sSql & "                     ,SGI_CORRMELHOR"
        sSql = sSql & "                     ,SGI_CORRPIOR"
        sSql = sSql & "            ) Values (" & vbCrLf
        sSql = sSql & "                      " & intFilial & vbCrLf
        sSql = sSql & "                     ," & intCODIGO & vbCrLf
        sSql = sSql & "                     ," & Trim(strIDProduto)
        sSql = sSql & "                     ,'" & Trim(strCODPROD) & "'" & vbCrLf
            
        strValor = Replace(curMELHOR, ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                             ," & strValor & vbCrLf
        
        strValor = Replace(curPIOR, ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                             ," & strValor & vbCrLf
        
        strValor = Replace(curCORRMELHOR, ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                             ," & strValor & vbCrLf
        
        strValor = Replace(curCORRPIOR, ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                             ," & strValor & vbCrLf
        
        sSql = sSql & "                             )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Grava Fluxo Produtivo (Estrutura)
        For I = 0 To UBound(arrPROCESSO)
        
            '' Produtos
            For j = 1 To arrPROCESSO(I).lngQTDPRODUTOS
                
                sSql = "Insert Into SGI_CADFLXPRODPROCESSO (" & vbCrLf
                sSql = sSql & "                                    SGI_FILIAL" & vbCrLf
                sSql = sSql & "                                   ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                                   ,SGI_CODPROC" & vbCrLf
                sSql = sSql & "                                   ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                                   ,SGI_CODPROD" & vbCrLf
                sSql = sSql & "                                   ,SGI_INDICE" & vbCrLf
                sSql = sSql & "                                   ,SGI_UNIDMED" & vbCrLf
                sSql = sSql & "                                   ,SGI_MELHOR" & vbCrLf
                sSql = sSql & "                                   ,SGI_PIOR)" & vbCrLf
                sSql = sSql & "                            Values (" & vbCrLf
                sSql = sSql & "                                    " & intFilial & vbCrLf
                sSql = sSql & "                                   ," & intCODIGO & vbCrLf
                sSql = sSql & "                                   ," & arrPROCESSO(I).lngCODIGO & vbCrLf
                sSql = sSql & "                                   ," & arrPROCESSO(I).typProdutos(j).lngIDPRODUTO & vbCrLf
                sSql = sSql & "                                   ,'" & arrPROCESSO(I).typProdutos(j).strPRODUTO & "'" & vbCrLf
                sSql = sSql & "                                   ," & I & vbCrLf
                sSql = sSql & "                                   ," & arrPROCESSO(I).typProdutos(j).lngCodUniMed & vbCrLf
                
                
                strValor = Replace(arrPROCESSO(I).typProdutos(j).curMELHORCENARIO, ".", "")
                strValor = Replace(strValor, ",", ".")
                sSql = sSql & "                                   ," & strValor & vbCrLf
                
                strValor = Replace(arrPROCESSO(I).typProdutos(j).curPIORCENARIO, ".", "")
                strValor = Replace(strValor, ",", ".")
                sSql = sSql & "                                   ," & strValor & vbCrLf
                sSql = sSql & "                                   )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
                
                '' Maquinas
                For k = 1 To arrPROCESSO(I).typProdutos(j).lngTOTMAQUINAS
                    sSql = "Insert Into SGI_CADFLXMAQUINA (" & vbCrLf
                    sSql = sSql & "                               SGI_FILIAL" & vbCrLf
                    sSql = sSql & "                              ,SGI_CODIGO" & vbCrLf
                    sSql = sSql & "                              ,SGI_CODPROC" & vbCrLf
                    sSql = sSql & "                              ,SGI_PRODUTO" & vbCrLf
                    sSql = sSql & "                              ,SGI_CODMAQUINA" & vbCrLf
                    sSql = sSql & "                              ,SGI_CODFAMMAQ" & vbCrLf
                    sSql = sSql & "                              ,SGI_INDICE" & vbCrLf
                    sSql = sSql & "                              ,SGI_UNIDMED" & vbCrLf
                    sSql = sSql & "                              ,SGI_QTDPORMN" & vbCrLf
                    sSql = sSql & "                              ,SGI_TEMPOPROD" & vbCrLf
                    sSql = sSql & "                              )" & vbCrLf
                    sSql = sSql & "                       Values ("
                    sSql = sSql & "                               " & intFilial & vbCrLf
                    sSql = sSql & "                              ," & intCODIGO & vbCrLf
                    sSql = sSql & "                              ," & arrPROCESSO(I).lngCODIGO & vbCrLf
                    sSql = sSql & "                              ,'" & Trim(arrPROCESSO(I).typProdutos(j).typMaquinas(k).strPAI) & "'" & vbCrLf
                    sSql = sSql & "                              ," & arrPROCESSO(I).typProdutos(j).typMaquinas(k).lngCODMAQ & vbCrLf
                    sSql = sSql & "                              ," & arrPROCESSO(I).typProdutos(j).lngCODFAMMAQ & vbCrLf
                    sSql = sSql & "                              ," & I & vbCrLf
                    sSql = sSql & "                              ," & arrPROCESSO(I).typProdutos(j).lngCodUniMed & vbCrLf
                    
                    strValor = Replace(arrPROCESSO(I).typProdutos(j).typMaquinas(k).lngQTDPCMIN, ".", "")
                    strValor = Replace(strValor, ",", ".")
                    sSql = sSql & "                              ," & strValor & vbCrLf
                
                    strValor = Replace(arrPROCESSO(I).typProdutos(j).typMaquinas(k).curTEMPPROD, ".", "")
                    strValor = Replace(strValor, ",", ".")
                    sSql = sSql & "                              ," & strValor & vbCrLf
                    
                    sSql = sSql & "                              )"
                
                    BGRV.CommandText = sSql
                    BGRV.Execute
                Next k
                
                '' Na Novalata não vai usar
                '' Produtos de Entrada
                For k = 1 To arrPROCESSO(I).typProdutos(j).lngTOTPRODENTRADA
                    sSql = "Insert Into SGI_CADFLXPRODENTR (" & vbCrLf
                    sSql = sSql & "                                SGI_FILIAL" & vbCrLf
                    sSql = sSql & "                               ,SGI_CODIGO" & vbCrLf
                    sSql = sSql & "                               ,SGI_CODPROC" & vbCrLf
                    sSql = sSql & "                               ,SGI_PRODPROC" & vbCrLf
                    sSql = sSql & "                               ,SGI_PRODENTR" & vbCrLf
                    sSql = sSql & "                               ,SGI_INDICE" & vbCrLf
                    sSql = sSql & "                               ,SGI_UNIMED" & vbCrLf
                    sSql = sSql & "                               ,SGI_ESTOQUE" & vbCrLf
                    sSql = sSql & "                               ,SGI_CADENCIA" & vbCrLf
                    sSql = sSql & "                               )" & vbCrLf
                    sSql = sSql & "                        Values (" & vbCrLf
                    sSql = sSql & "                                " & intFilial & vbCrLf
                    sSql = sSql & "                               ," & intCODIGO & vbCrLf
                    sSql = sSql & "                               ," & arrPROCESSO(I).lngCODIGO & vbCrLf
                    sSql = sSql & "                               ,'" & Trim(arrPROCESSO(I).typProdutos(j).typProdEntrada(k).strPAI) & "'" & vbCrLf
                    sSql = sSql & "                               ,'" & Trim(arrPROCESSO(I).typProdutos(j).typProdEntrada(k).strCODPROD) & "'" & vbCrLf
                    sSql = sSql & "                               ," & I & vbCrLf
                    sSql = sSql & "                               ," & arrPROCESSO(I).typProdutos(j).typProdEntrada(k).lngUNIDMED & vbCrLf
                    
                    strValor = Replace(arrPROCESSO(I).typProdutos(j).typProdEntrada(k).curQTDESTOQUE, ".", "")
                    strValor = Replace(strValor, ",", ".")
                    sSql = sSql & "                              ," & strValor & vbCrLf
                    
                    sSql = sSql & "                               ," & arrPROCESSO(I).typProdutos(j).typProdEntrada(k).intCADENCIA & vbCrLf
                    sSql = sSql & "                                )"
                    
                    BGRV.CommandText = sSql
                    BGRV.Execute
                Next k
                
                '' Produtos de Saida
                For k = 1 To arrPROCESSO(I).typProdutos(j).lngTOTPRODSAIDA
                    sSql = "Insert Into SGI_CADFLXPRODSAID (" & vbCrLf
                    sSql = sSql & "                                SGI_FILIAL" & vbCrLf
                    sSql = sSql & "                               ,SGI_CODIGO" & vbCrLf
                    sSql = sSql & "                               ,SGI_CODPROC" & vbCrLf
                    sSql = sSql & "                               ,SGI_PRODPROC" & vbCrLf
                    sSql = sSql & "                               ,SGI_PRODSAID" & vbCrLf
                    sSql = sSql & "                               ,SGI_INDICE" & vbCrLf
                    sSql = sSql & "                               ,SGI_UNIDMED" & vbCrLf
                    sSql = sSql & "                               ,SGI_ESTOQUE" & vbCrLf
                    sSql = sSql & "                               )" & vbCrLf
                    sSql = sSql & "                        Values (" & vbCrLf
                    sSql = sSql & "                                " & intFilial & vbCrLf
                    sSql = sSql & "                               ," & intCODIGO & vbCrLf
                    sSql = sSql & "                               ," & arrPROCESSO(I).lngCODIGO & vbCrLf
                    sSql = sSql & "                               ,'" & Trim(arrPROCESSO(I).typProdutos(j).typProdSaida(k).strPAI) & "'" & vbCrLf
                    sSql = sSql & "                               ,'" & Trim(arrPROCESSO(I).typProdutos(j).typProdSaida(k).strCODPROD) & "'" & vbCrLf
                    sSql = sSql & "                               ," & I & vbCrLf
                    sSql = sSql & "                               ," & arrPROCESSO(I).typProdutos(j).typProdSaida(k).lngUNIDMED & vbCrLf
                
                    strValor = Replace(arrPROCESSO(I).typProdutos(j).typProdSaida(k).curQTDESTOQUE, ".", "")
                    strValor = Replace(strValor, ",", ".")
                    sSql = sSql & "                              ," & strValor & vbCrLf
                    
                    sSql = sSql & "                               )"
                    
                    BGRV.CommandText = sSql
                    BGRV.Execute
                
                Next k
            Next j
        Next I
        '' ===================================================================================
        
        '' Grava Fluxo Produtivo do Produto
        For I = 0 To UBound(arrPROCESSO)
                
            sSql = "Insert Into SGI_CADPROPROC (SGI_FILIAL" & vbCrLf
            sSql = sSql & "                    ,SGI_IDPRODUTO" & vbCrLf
            sSql = sSql & "                    ,SGI_CODPROD" & vbCrLf
            sSql = sSql & "                    ,SGI_CODPROC" & vbCrLf
            sSql = sSql & "                    ,SGI_TEMPO" & vbCrLf
            sSql = sSql & "                    ,SGI_VALORHORA" & vbCrLf
            sSql = sSql & "                    ,SGI_TOTVALOR" & vbCrLf
            sSql = sSql & "           ) Values (" & intFilial & vbCrLf
            sSql = sSql & "                  ," & strIDProduto & vbCrLf
            sSql = sSql & "                  ,'" & Trim(strCODPROD) & "'" & vbCrLf
            sSql = sSql & "                  ," & arrPROCESSO(I).lngCODIGO & vbCrLf
            
            strValor = Replace(arrPROCESSO(I).curMELHORCENARIO, ".", "")
            strValor = Replace(strValor, ",", ".")
            sSql = sSql & "                  ," & strValor & vbCrLf
            sSql = sSql & "                  ,0" & vbCrLf
            sSql = sSql & "                  ,0" & vbCrLf
            sSql = sSql & "                   )"
    
            BGRV.CommandText = sSql
            BGRV.Execute
            
        Next I
        
     ElseIf strOPER = "E" Or strOPER = "A" Then
        
        '' Apaga Fluxo de Produtos
        sSql = "Delete from SGI_CADPROPROC " & vbCrLf
        sSql = sSql & "              Where SGI_CODPROD = '" & Trim(strCODPROD) & "'" & vbCrLf
        sSql = sSql & "                And SGI_FILIAL  = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
                
        '' Apaga Produto de Saida
        sSql = "Delete from SGI_CADFLXPRODSAID " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Apaga Produto de Entrada
        sSql = "Delete from SGI_CADFLXPRODENTR " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Apaga Estrutura de Maquinas
        sSql = "Delete from SGI_CADFLXMAQUINA " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Apaga Estrutura de Fluxo
        sSql = "Delete from SGI_CADFLXPRODPROCESSO " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Apaga Header do Fluxo
        sSql = "Delete from SGI_CADFLUXPROD " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        If strOPER = "A" Then GoTo VOLTA
        
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

Public Function Carrega_campos() As Boolean

   Dim lngTOTREG        As Long
   Dim lngREGSPROD      As Long
   Dim lngREGSMAQ       As Long
   Dim lngREGSPRDENTR   As Long
   Dim lngREGSPRDSAID   As Long
   
   Dim arrPRODS()       As PRODUTOS
   Dim arrPRODSV()      As PRODUTOS
   
   Dim arrMAQUINAS()    As Maquinas
   Dim arrMAQUINASV()   As Maquinas
   
   Dim arrPRODENTR()    As ProdEntrada
   Dim arrPRODENTRV()   As ProdEntrada
   
   Dim arrPRODSAID()    As ProdSaida
   Dim arrPRODSAIDV()   As ProdSaida
   
   Carrega_campos = False
   
   sSql = ""
   
   '' ==============================================================
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADFLUXPROD " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & intCODIGO & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then
      
      intCODIGO = BREC!SGI_CODIGO
      strIDProduto = BREC!SGI_IDPRODUTO
      strCODPROD = BREC!SGI_CODPROD
      curMELHOR = BREC!SGI_MELHOR
      curPIOR = BREC!SGI_PIOR
      curCORRMELHOR = BREC!SGI_CORRMELHOR
      curCORRPIOR = BREC!SGI_CORRPIOR
      
      Carrega_campos = True
   End If
   BREC.Close
   
   '' ==============================================================
   sSql = "Select " & vbCrLf
   sSql = sSql & "        SGI_INDICE  " & vbCrLf
   sSql = sSql & "       ,SGI_CODPROC " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADFLXPRODPROCESSO " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_CODIGO = " & intCODIGO & vbCrLf
   sSql = sSql & " Group By SGI_INDICE,SGI_CODPROC " & vbCrLf
   sSql = sSql & " Order By SGI_INDICE,SGI_CODPROC "
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   
   lngTOTREG = 0
   Do While Not BREC2.EOF
   
      ReDim Preserve arrPROCESSO(0 To lngTOTREG) As Processos
      arrPROCESSO(lngTOTREG).intORCEM = BREC2!SGI_INDICE
      arrPROCESSO(lngTOTREG).intTipo = 0
      arrPROCESSO(lngTOTREG).lngCODIGO = BREC2!SGI_CODPROC
      arrPROCESSO(lngTOTREG).strDESCRI = PesqDescProc(BREC2!SGI_CODPROC)
      
      '' =================================================================
      '' Processo
      sSql = "Select " & vbCrLf
      sSql = sSql & "       * " & vbCrLf
      sSql = sSql & "  From " & vbCrLf
      sSql = sSql & "       SGI_CADFLXPRODPROCESSO " & vbCrLf
      sSql = sSql & " Where " & vbCrLf
      sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
      sSql = sSql & "   And SGI_CODIGO = " & intCODIGO & vbCrLf
      sSql = sSql & "   And SGI_INDICE = " & BREC2!SGI_INDICE
      
      BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
      
      lngREGSPROD = 0
      Do While Not BREC3.EOF
         lngREGSPROD = lngREGSPROD + 1
         ReDim Preserve arrPRODS(1 To lngREGSPROD) As PRODUTOS
         arrPRODS(lngREGSPROD).intTipo = 0
         arrPRODS(lngREGSPROD).strPRODUTO = BREC3!SGI_CODPROD
         arrPRODS(lngREGSPROD).lngIDPRODUTO = BREC3!SGI_IDPRODUTO
         arrPRODS(lngREGSPROD).lngCodUniMed = 0
         arrPRODS(lngREGSPROD).lngCODFAMMAQ = PegaCodFamMaq(Trim(BREC3!SGI_CODPROD))
         arrPRODS(lngREGSPROD).curMELHORCENARIO = BREC3!SGI_MELHOR
         arrPRODS(lngREGSPROD).curPIORCENARIO = BREC3!SGI_PIOR
         
         '' =================================================================
         '' Maquinas
         sSql = "Select " & vbCrLf
         sSql = sSql & "       * " & vbCrLf
         sSql = sSql & "  From " & vbCrLf
         sSql = sSql & "       SGI_CADFLXMAQUINA " & vbCrLf
         sSql = sSql & " Where " & vbCrLf
         sSql = sSql & "       SGI_FILIAL  = " & intFilial & vbCrLf
         sSql = sSql & "   And SGI_CODIGO  = " & intCODIGO & vbCrLf
         sSql = sSql & "   And SGI_PRODUTO = '" & BREC3!SGI_IDPRODUTO & "'" & vbCrLf
         sSql = sSql & "   And SGI_INDICE  = " & BREC2!SGI_INDICE
         
         BREC4.Open sSql, adoBanco_Dados, adOpenDynamic
         
         lngREGSMAQ = 0
         Do While Not BREC4.EOF
            lngREGSMAQ = lngREGSMAQ + 1
            ReDim Preserve arrMAQUINAS(1 To lngREGSMAQ) As Maquinas
            arrMAQUINAS(lngREGSMAQ).intTipo = 0
            arrMAQUINAS(lngREGSMAQ).lngCODMAQ = BREC4!SGI_CODMAQUINA
            arrMAQUINAS(lngREGSMAQ).strINDICE = Trim(Str(BREC4!SGI_CODMAQUINA)) & Trim(BREC4!SGI_PRODUTO)
            arrMAQUINAS(lngREGSMAQ).strPAI = BREC4!SGI_PRODUTO
            arrMAQUINAS(lngREGSMAQ).lngQTDPCMIN = BREC4!SGI_QTDPORMN
            arrMAQUINAS(lngREGSMAQ).curTEMPPROD = BREC4!SGI_TEMPOPROD
            BREC4.MoveNext
         Loop
         BREC4.Close
         arrPRODS(lngREGSPROD).lngTOTMAQUINAS = lngREGSMAQ
         arrPRODS(lngREGSPROD).typMaquinas = arrMAQUINAS
         '' =================================================================
         
         '' =================================================================
         '' Produto de Entrada
         sSql = "Select " & vbCrLf
         sSql = sSql & "       * " & vbCrLf
         sSql = sSql & "  From " & vbCrLf
         sSql = sSql & "       SGI_CADFLXPRODENTR " & vbCrLf
         sSql = sSql & " Where " & vbCrLf
         sSql = sSql & "       SGI_FILIAL   = " & intFilial & vbCrLf
         sSql = sSql & "   And SGI_CODIGO   = " & intCODIGO & vbCrLf
         sSql = sSql & "   And SGI_PRODPROC = '" & BREC3!SGI_CODPROD & "'" & vbCrLf
         sSql = sSql & "   And SGI_INDICE   = " & BREC2!SGI_INDICE
         
         BREC5.Open sSql, adoBanco_Dados, adOpenDynamic
         
         lngREGSPRDENTR = 0
         Do While Not BREC5.EOF
            lngREGSPRDENTR = lngREGSPRDENTR + 1
            ReDim Preserve arrPRODENTR(1 To lngREGSPRDENTR) As ProdEntrada
            arrPRODENTR(lngREGSPRDENTR).intTipo = 0
            arrPRODENTR(lngREGSPRDENTR).strCODPROD = BREC5!SGI_PRODENTR
            arrPRODENTR(lngREGSPRDENTR).strPAI = BREC5!SGI_PRODPROC
            arrPRODENTR(lngREGSPRDENTR).lngUNIDMED = BREC5!SGI_UNIMED
            arrPRODENTR(lngREGSPRDENTR).curQTDESTOQUE = BREC5!SGI_ESTOQUE
            arrPRODENTR(lngREGSPRDENTR).intCADENCIA = BREC5!SGI_CADENCIA
            BREC5.MoveNext
         Loop
         BREC5.Close
         arrPRODS(lngREGSPROD).lngTOTPRODENTRADA = lngREGSPRDENTR
         arrPRODS(lngREGSPROD).typProdEntrada = arrPRODENTR
         '' =================================================================
         
         
         '' =================================================================
         '' Produto de Saida
         sSql = "Select " & vbCrLf
         sSql = sSql & "       * " & vbCrLf
         sSql = sSql & "  From " & vbCrLf
         sSql = sSql & "       SGI_CADFLXPRODSAID " & vbCrLf
         sSql = sSql & " Where " & vbCrLf
         sSql = sSql & "       SGI_FILIAL   = " & intFilial & vbCrLf
         sSql = sSql & "   And SGI_CODIGO   = " & intCODIGO & vbCrLf
         sSql = sSql & "   And SGI_PRODPROC = '" & BREC3!SGI_CODPROD & "'" & vbCrLf
         sSql = sSql & "   And SGI_INDICE   = " & BREC2!SGI_INDICE
         
         BREC6.Open sSql, adoBanco_Dados, adOpenDynamic
         
         lngREGSPRDSAID = 0
         Do While Not BREC6.EOF
            lngREGSPRDSAID = lngREGSPRDSAID + 1
            ReDim Preserve arrPRODSAID(1 To lngREGSPRDSAID) As ProdSaida
            arrPRODSAID(lngREGSPRDSAID).intTipo = 0
            arrPRODSAID(lngREGSPRDSAID).strPAI = BREC6!SGI_PRODPROC
            arrPRODSAID(lngREGSPRDSAID).strCODPROD = BREC6!SGI_PRODSAID
            arrPRODSAID(lngREGSPRDSAID).lngUNIDMED = BREC6!SGI_UNIDMED
            arrPRODSAID(lngREGSPRDSAID).curQTDESTOQUE = BREC6!SGI_ESTOQUE
            BREC6.MoveNext
         Loop
         BREC6.Close
         arrPRODS(lngREGSPROD).lngTOTPRODSAIDA = lngREGSPRDSAID
         arrPRODS(lngREGSPROD).typProdSaida = arrPRODSAID
         '' =================================================================
         
         BREC3.MoveNext
      Loop
      BREC3.Close
      '' =================================================================
      arrPROCESSO(lngTOTREG).lngQTDPRODUTOS = lngREGSPROD
      arrPROCESSO(lngTOTREG).typProdutos = arrPRODS
      ''arrPROCESSO(lngTOTREG).curMELHORCENARIO = BREC2!SGI_MELHOR
      ''arrPROCESSO(lngTOTREG).curPIORCENARIO = BREC2!SGI_PIOR
      
      lngTOTREG = (lngTOTREG + 1)
      BREC2.MoveNext
   Loop
   BREC2.Close
   
End Function

Private Function PesqDescProc(lngCODPROC As Long) As String
    
    PesqDescProc = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPROCESSO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & lngCODPROC
    
    BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC3.EOF Then PesqDescProc = BREC3!SGI_DESCRI
    BREC3.Close
    
End Function

Private Function PegaCodFamMaq(strCODPROD As String) As Long

    PegaCodFamMaq = 0
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_TIPOPERACAO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & Trim(strCODPROD)
    
    BREC4.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC4.EOF Then PegaCodFamMaq = BREC4!SGI_CODFAMMAQ
    BREC4.Close
    
End Function
