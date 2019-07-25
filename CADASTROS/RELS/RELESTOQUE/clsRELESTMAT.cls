VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRELESTMAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial       As Integer
Dim arrPRODPAI      As Variant
Dim arrPROCPROD01   As Variant
Dim arrPROCPROD02   As Variant
Dim arrPROCPROD03   As Variant
Dim arrPROCPROD04   As Variant
Dim arrPROCPROD05   As Variant
Dim arrPROCPROD06   As Variant
Dim arrPROCPROD07   As Variant
Dim arrPROCPROD08   As Variant
Dim arrPROCPROD09   As Variant
Dim arrNIVEL01      As Variant
Dim arrNIVEL02      As Variant
Dim arrNIVEL03      As Variant
Dim arrNIVEL04      As Variant
Dim arrNIVEL05      As Variant
Dim arrNIVEL06      As Variant
Dim arrNIVEL07      As Variant
Dim arrNIVEL08      As Variant
Dim arrNIVEL09      As Variant
Dim arrPROCPAI      As Variant

Public Property Get PROCPROD01() As Variant
    PROCPROD01 = arrPROCPROD01
End Property

Public Property Let PROCPROD01(ByVal vData As Variant)
    arrPROCPROD01 = vData
End Property

Public Property Get PROCPROD02() As Variant
    PROCPROD02 = arrPROCPROD02
End Property

Public Property Let PROCPROD02(ByVal vData As Variant)
    arrPROCPROD02 = vData
End Property

Public Property Get PROCPROD03() As Variant
    PROCPROD03 = arrPROCPROD03
End Property

Public Property Let PROCPROD03(ByVal vData As Variant)
    arrPROCPROD03 = vData
End Property

Public Property Get PROCPROD04() As Variant
    PROCPROD04 = arrPROCPROD04
End Property

Public Property Let PROCPROD04(ByVal vData As Variant)
    arrPROCPROD04 = vData
End Property

Public Property Get PROCPROD05() As Variant
    PROCPROD05 = arrPROCPROD05
End Property

Public Property Let PROCPROD05(ByVal vData As Variant)
    arrPROCPROD05 = vData
End Property

Public Property Get PROCPROD06() As Variant
    PROCPROD06 = arrPROCPROD06
End Property

Public Property Let PROCPROD06(ByVal vData As Variant)
    arrPROCPROD06 = vData
End Property

Public Property Get PROCPROD07() As Variant
    PROCPROD07 = arrPROCPROD07
End Property

Public Property Let PROCPROD07(ByVal vData As Variant)
    arrPROCPROD07 = vData
End Property

Public Property Get PROCPROD08() As Variant
    PROCPROD08 = arrPROCPROD08
End Property

Public Property Let PROCPROD08(ByVal vData As Variant)
    arrPROCPROD08 = vData
End Property

Public Property Get PROCPROD09() As Variant
    PROCPROD09 = arrPROCPROD09
End Property

Public Property Let PROCPROD09(ByVal vData As Variant)
    arrPROCPROD09 = vData
End Property

Public Property Get NIVEL01() As Variant
    NIVEL01 = arrNIVEL01
End Property

Public Property Let NIVEL01(ByVal vData As Variant)
    arrNIVEL01 = vData
End Property

Public Property Get NIVEL02() As Variant
    NIVEL02 = arrNIVEL02
End Property

Public Property Let NIVEL02(ByVal vData As Variant)
    arrNIVEL02 = vData
End Property

Public Property Get NIVEL03() As Variant
    NIVEL03 = arrNIVEL03
End Property

Public Property Let NIVEL03(ByVal vData As Variant)
    arrNIVEL03 = vData
End Property

Public Property Get NIVEL04() As Variant
    NIVEL04 = arrNIVEL04
End Property

Public Property Let NIVEL04(ByVal vData As Variant)
    arrNIVEL04 = vData
End Property

Public Property Get NIVEL05() As Variant
    NIVEL05 = arrNIVEL05
End Property

Public Property Let NIVEL05(ByVal vData As Variant)
    arrNIVEL05 = vData
End Property

Public Property Get NIVEL06() As Variant
    NIVEL06 = arrNIVEL06
End Property

Public Property Let NIVEL06(ByVal vData As Variant)
    arrNIVEL06 = vData
End Property

Public Property Get NIVEL07() As Variant
    NIVEL07 = arrNIVEL07
End Property

Public Property Let NIVEL07(ByVal vData As Variant)
    arrNIVEL07 = vData
End Property

Public Property Get NIVEL08() As Variant
    NIVEL08 = arrNIVEL08
End Property

Public Property Let NIVEL08(ByVal vData As Variant)
    arrNIVEL08 = vData
End Property

Public Property Get NIVEL09() As Variant
    NIVEL09 = arrNIVEL09
End Property

Public Property Let NIVEL09(ByVal vData As Variant)
    arrNIVEL09 = vData
End Property

Public Property Get PRODPAI() As Variant
    PRODPAI = arrPRODPAI
End Property

Public Property Let PRODPAI(ByVal vData As Variant)
    arrPRODPAI = vData
End Property

Public Property Get PROCPAI() As Variant
    PROCPAI = arrPROCPAI
End Property

Public Property Let PROCPAI(ByVal vData As Variant)
    arrPROCPAI = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String)
    frmRELESTMAT.cCaminho = cCaminho
    frmRELESTMAT.Linha = Linha
    frmRELESTMAT.FILIAL = iFilial
    frmRELESTMAT.strAcesso = strAcesso
    frmRELESTMAT.Show vbModal
End Sub

Public Sub PreencheComboProd(Combo As Variant, intEstilo As Integer)
     sSql = "Select " & vbCrLf
     sSql = sSql & "       * " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL        = " & intFilial & vbCrLf
     sSql = sSql & "   And SGI_PRODUTOESTILO = " & intEstilo
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
    Do While Not BREC.EOF
       Combo.AddItem Trim(BREC!SGI_CODIGO) & Space(10 - Len(Trim(BREC!SGI_CODIGO))) & " - " & Trim(BREC!SGI_DESCRICAO)
       BREC.MoveNext
    Loop
    BREC.Close
End Sub

Public Sub PreenchComboGrupProduto(Combo As Variant)

    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADGRUPROD " & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "Order by SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       Combo.AddItem Trim(BREC!SGI_DESCRICAO)
       Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
       BREC.MoveNext
    Loop
    
    BREC.Close

End Sub

Public Sub PreenchComboSubGrupProduto(Combo As Variant)

    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADSUBGRPROD " & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "Order by SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       Combo.AddItem Trim(BREC!SGI_DESCRICAO)
       Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
       BREC.MoveNext
    Loop
    
    BREC.Close

End Sub

Public Sub PreenchComboEspecie(Combo As Variant)

    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       SGI_CODIGO    " & vbCrLf
    sSql = sSql & "      ,SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADESPPROD " & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & FILIAL
    
    sSql = sSql & "Order by SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       Combo.AddItem Trim(BREC!SGI_DESCRICAO)
       Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
       BREC.MoveNext
    Loop
    
    BREC.Close

End Sub

Public Sub PreenchComboTipo(Combo As Variant)

    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADTIPPROD " & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "Order by SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       Combo.AddItem Trim(BREC!SGI_DESCRICAO)
       Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
       BREC.MoveNext
    Loop
    
    BREC.Close

End Sub

Public Sub PreenchComboEspTecnica(Combo As Variant)

    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADESPTEC " & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "Order by SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       Combo.AddItem Trim(BREC!SGI_DESCESPTEC)
       Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
       BREC.MoveNext
    Loop
   
    BREC.Close

End Sub

Public Sub PreencheComboFornec(Combo As Variant)

     sSql = "Select " & vbCrLf
     sSql = sSql & "       FORN.* " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADFORNEC  FORN " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       FORN.SGI_FILIAL  = " & intFilial & vbCrLf
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_CODIGO) & Space(10 - Len(Trim(BREC!SGI_CODIGO))) & " - " & Trim(BREC!SGI_RAZAOSOC)
        Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
        BREC.MoveNext
     Loop
    
     BREC.Close
    
End Sub


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

Public Function Grava_ListMatTemp()

On Error GoTo err_Trans

     Grava_ListMatTemp = False
     
     Dim I As Integer
     Dim strValor As String
         
     BGRV.ActiveConnection = adoBanco_Dados
     adoBanco_Dados.BeginTrans
     
     '' Apagando a tabela Temporária
     sSql = "Delete from SGI_LISTAMAT_TEMP Where SGI_FILIAL = " & intFilial
     BGRV.CommandText = sSql
     BGRV.Execute
     
     '' Apagando Lista de Processos
     sSql = "Delete from SGI_CADPROPROC_TEMP Where SGI_FILIAL = " & intFilial
     BGRV.CommandText = sSql
     BGRV.Execute
     
     If IsArray(arrPRODPAI) Then
       
        '' Gravando Produto Pai
        sSql = "Insert Into SGI_LISTAMAT_TEMP (SGI_FILIAL" & vbCrLf
        sSql = sSql & "                       ,SGI_PRODUTO0" & vbCrLf
        sSql = sSql & "                       ,SGI_DESCRICAO" & vbCrLf
        sSql = sSql & "                       ,SGI_QTDE" & vbCrLf
        sSql = sSql & "                       ,SGI_CUSTOUNI" & vbCrLf
        sSql = sSql & "                       ,SGI_CUSTOTOTAL" & vbCrLf
        sSql = sSql & "                       ,SGI_UNIDCONS" & vbCrLf
        sSql = sSql & "                       ,SGI_CODTABCONV)" & vbCrLf
        sSql = sSql & "                       Values (" & intFilial & vbCrLf
        sSql = sSql & "                              ,'" & Trim(arrPRODPAI(1, 1)) & "'" & vbCrLf
        sSql = sSql & "                              ,'" & Trim(arrPRODPAI(1, 2)) & "'" & vbCrLf
        sSql = sSql & "                              ," & arrPRODPAI(1, 3) & vbCrLf
        
        strValor = Replace(arrPRODPAI(1, 4), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                              ," & strValor & vbCrLf
        
        strValor = Replace(arrPRODPAI(1, 5), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                              ," & strValor & vbCrLf
        
        sSql = sSql & "                              ,'" & arrPRODPAI(1, 6) & "'" & vbCrLf
        
        If Not IsNull(arrPRODPAI(1, 7)) Then
           sSql = sSql & "                              ," & arrPRODPAI(1, 7) & vbCrLf
        Else
           sSql = sSql & "                              ,Null" & vbCrLf
        End If
        
        sSql = sSql & "                              )" & vbCrLf
        
        BGRV.CommandText = sSql
        BGRV.Execute
       
    
    End If
        
    '' -----------------------------------
    '' Gravando Filho 01
    If IsArray(arrNIVEL01) Then
       For I = 1 To UBound(arrNIVEL01)
     
           sSql = "Insert Into SGI_LISTAMAT_TEMP (SGI_FILIAL" & vbCrLf
           sSql = sSql & "                       ,SGI_PRODUTO1" & vbCrLf
           sSql = sSql & "                       ,SGI_DESCRICAO" & vbCrLf
           sSql = sSql & "                       ,SGI_QTDE" & vbCrLf
           sSql = sSql & "                       ,SGI_CUSTOUNI" & vbCrLf
           sSql = sSql & "                       ,SGI_CUSTOTOTAL" & vbCrLf
           sSql = sSql & "                       ,SGI_UNIDCONS" & vbCrLf
           sSql = sSql & "                       ,SGI_CODTABCONV)" & vbCrLf
           sSql = sSql & "                       Values (" & intFilial & vbCrLf
           sSql = sSql & "                              ,'" & Trim(arrNIVEL01(I, 1)) & "'" & vbCrLf
           sSql = sSql & "                              ,'" & Trim(arrNIVEL01(I, 2)) & "'" & vbCrLf
           sSql = sSql & "                              ," & arrNIVEL01(I, 3) & vbCrLf
        
           strValor = Replace(arrNIVEL01(I, 4), ".", "")
           strValor = Replace(strValor, ",", ".")
           sSql = sSql & "                              ," & strValor & vbCrLf
           strValor = Replace(arrNIVEL01(I, 5), ".", "")
           strValor = Replace(strValor, ",", ".")
           sSql = sSql & "                              ," & strValor & vbCrLf
           sSql = sSql & "                              ,'" & arrNIVEL01(I, 6) & "'" & vbCrLf
           If Not IsNull(arrNIVEL01(I, 7)) And arrNIVEL01(I, 7) > 0 Then
              sSql = sSql & "                              ," & arrNIVEL01(I, 7) & vbCrLf
           Else
              sSql = sSql & "                              ,Null" & vbCrLf
           End If
           sSql = sSql & "                              )" & vbCrLf
        
           BGRV.CommandText = sSql
           BGRV.Execute
       Next I
    End If
   
    '' -----------------------------------
    '' Gravando Filho 02
    If IsArray(arrNIVEL02) Then
       For I = 1 To UBound(arrNIVEL02)
     
           sSql = "Insert Into SGI_LISTAMAT_TEMP (SGI_FILIAL" & vbCrLf
           sSql = sSql & "                       ,SGI_PRODUTO2" & vbCrLf
           sSql = sSql & "                       ,SGI_DESCRICAO" & vbCrLf
           sSql = sSql & "                       ,SGI_QTDE" & vbCrLf
           sSql = sSql & "                       ,SGI_CUSTOUNI" & vbCrLf
           sSql = sSql & "                       ,SGI_CUSTOTOTAL" & vbCrLf
           sSql = sSql & "                       ,SGI_UNIDCONS" & vbCrLf
           sSql = sSql & "                       ,SGI_CODTABCONV)" & vbCrLf
           sSql = sSql & "                       Values (" & intFilial & vbCrLf
           sSql = sSql & "                              ,'" & Trim(arrNIVEL02(I, 1)) & "'" & vbCrLf
           sSql = sSql & "                              ,'" & Trim(arrNIVEL02(I, 2)) & "'" & vbCrLf
           sSql = sSql & "                              ," & arrNIVEL02(I, 3) & vbCrLf
        
           strValor = Replace(arrNIVEL02(I, 4), ".", "")
           strValor = Replace(strValor, ",", ".")
           sSql = sSql & "                              ," & strValor & vbCrLf
           strValor = Replace(arrNIVEL02(I, 5), ".", "")
           strValor = Replace(strValor, ",", ".")
           sSql = sSql & "                              ," & strValor & vbCrLf
           sSql = sSql & "                              ,'" & arrNIVEL02(I, 6) & "'" & vbCrLf
           If Not IsNull(arrNIVEL02(I, 7)) And arrNIVEL02(I, 7) > 0 Then
              sSql = sSql & "                              ," & arrNIVEL02(I, 7) & vbCrLf
           Else
              sSql = sSql & "                              ,Null" & vbCrLf
           End If
           sSql = sSql & "                              )" & vbCrLf
        
           BGRV.CommandText = sSql
           BGRV.Execute
       Next I
    End If
    
    adoBanco_Dados.CommitTrans
    Grava_ListMatTemp = True
     
    Exit Function
     
err_Trans:
     
     adoBanco_Dados.RollbackTrans
     
     Dim objErro    As Object
     Set objErro = CreateObject("BLBCWS.clsFuncoes")
     Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, "I", sSql)
     Set objErro = Nothing

End Function



Public Function Grava_ListMatTemp_ProdPai()

On Error GoTo err_Trans

     Grava_ListMatTemp_ProdPai = False
     
     Dim I As Integer
     Dim strValor As String
         
     BGRV.ActiveConnection = adoBanco_Dados
     adoBanco_Dados.BeginTrans
     
     '' Apagando a tabela Temporária
     sSql = "Delete from SGI_LISTAMAT_TEMP Where SGI_FILIAL = " & intFilial
     BGRV.CommandText = sSql
     BGRV.Execute
     
     '' Apagando Processo Temporário
     sSql = "Delete From SGI_CADPROPROC_TEMP Where SGI_FILIAL = " & intFilial
     BGRV.CommandText = sSql
     BGRV.Execute
     
     If IsArray(arrPRODPAI) Then
       
        '' Gravando Produto Pai
        sSql = "Insert Into SGI_LISTAMAT_TEMP (SGI_FILIAL" & vbCrLf
        sSql = sSql & "                       ,SGI_PRODUTO0" & vbCrLf
        sSql = sSql & "                       ,SGI_DESCRICAO" & vbCrLf
        sSql = sSql & "                       ,SGI_QTDE" & vbCrLf
        sSql = sSql & "                       ,SGI_CUSTOUNI" & vbCrLf
        sSql = sSql & "                       ,SGI_CUSTOTOTAL" & vbCrLf
        sSql = sSql & "                       ,SGI_UNIDCONS" & vbCrLf
        sSql = sSql & "                       ,SGI_CODTABCONV" & vbCrLf
        sSql = sSql & "                       ,SGI_PARENTESCO)" & vbCrLf
        
        sSql = sSql & "                       Values (" & intFilial & vbCrLf
        sSql = sSql & "                              ,'" & Trim(arrPRODPAI(1, 1)) & "'" & vbCrLf
        sSql = sSql & "                              ,'" & Trim(arrPRODPAI(1, 2)) & "'" & vbCrLf
        sSql = sSql & "                              ," & arrPRODPAI(1, 3) & vbCrLf
        
        strValor = Replace(arrPRODPAI(1, 4), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                              ," & strValor & vbCrLf
        
        strValor = Replace(arrPRODPAI(1, 5), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                              ," & strValor & vbCrLf
        
        sSql = sSql & "                              ,'" & arrPRODPAI(1, 6) & "'" & vbCrLf
        
        If Not IsNull(arrPRODPAI(1, 7)) Then
           sSql = sSql & "                              ," & arrPRODPAI(1, 7) & vbCrLf
        Else
           sSql = sSql & "                              ,Null" & vbCrLf
        End If
        sSql = sSql & "                              ,'" & Trim(arrPRODPAI(1, 8)) & "'" & vbCrLf
        
        sSql = sSql & "                              )" & vbCrLf
        
        BGRV.CommandText = sSql
        BGRV.Execute
       
    
        ''Gravando Processo do Pai
        If IsArray(arrPROCPAI) Then
           For I = 1 To UBound(arrPROCPAI)
               sSql = "Insert into SGI_CADPROPROC_TEMP (SGI_FILIAL" & vbCrLf
               sSql = sSql & "                         ,SGI_PARENTESCO" & vbCrLf
               sSql = sSql & "                         ,SGI_CODPROC0" & vbCrLf
               sSql = sSql & "                         ,SGI_DESCRICAOPROC0" & vbCrLf
               sSql = sSql & "                         ,SGI_TEMPO" & vbCrLf
               sSql = sSql & "                         ,SGI_VALORHORA" & vbCrLf
               sSql = sSql & "                         ,SGI_TOTVALOR)" & vbCrLf
               sSql = sSql & "                  Values (" & vbCrLf
               sSql = sSql & "                          " & intFilial & vbCrLf
               sSql = sSql & "                         ,'" & Trim(arrPROCPAI(I, 1)) & "'" & vbCrLf
               sSql = sSql & "                         ," & arrPROCPAI(I, 2) & vbCrLf
               sSql = sSql & "                         ,'" & Trim(arrPROCPAI(I, 6)) & "'" & vbCrLf
               
               strValor = Replace(arrPROCPAI(I, 3), ".", "")
               strValor = Replace(strValor, ",", ".")
               sSql = sSql & "                         ," & strValor & vbCrLf
               
               strValor = Replace(arrPROCPAI(I, 4), ".", "")
               strValor = Replace(strValor, ",", ".")
               sSql = sSql & "                         ," & strValor & vbCrLf
               
               strValor = Replace(arrPROCPAI(I, 5), ".", "")
               strValor = Replace(strValor, ",", ".")
               sSql = sSql & "                         ," & strValor & vbCrLf
               
               sSql = sSql & "                         )"
               BGRV.CommandText = sSql
               BGRV.Execute
           Next I
        End If
    End If
        
    
    adoBanco_Dados.CommitTrans
    Grava_ListMatTemp_ProdPai = True
     
    Exit Function
     
err_Trans:
     
     adoBanco_Dados.RollbackTrans
     
     Dim objErro    As Object
     Set objErro = CreateObject("BLBCWS.clsFuncoes")
     Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, "I", sSql)
     Set objErro = Nothing

End Function



Public Function Grava_ListMatTemp_Filho1(intReg As Integer)

On Error GoTo err_Trans

     Grava_ListMatTemp_Filho1 = False
     
     Dim I As Integer
     Dim strValor As String
         
     BGRV.ActiveConnection = adoBanco_Dados
     adoBanco_Dados.BeginTrans
        
    '' -----------------------------------
    '' Gravando Filho 01
    If IsArray(arrNIVEL01) Then
     
           sSql = "Insert Into SGI_LISTAMAT_TEMP (SGI_FILIAL" & vbCrLf
           sSql = sSql & "                       ,SGI_PRODUTO1" & vbCrLf
           sSql = sSql & "                       ,SGI_DESCRICAO" & vbCrLf
           sSql = sSql & "                       ,SGI_QTDE" & vbCrLf
           sSql = sSql & "                       ,SGI_CUSTOUNI" & vbCrLf
           sSql = sSql & "                       ,SGI_CUSTOTOTAL" & vbCrLf
           sSql = sSql & "                       ,SGI_UNIDCONS" & vbCrLf
           sSql = sSql & "                       ,SGI_CODTABCONV" & vbCrLf
           sSql = sSql & "                       ,SGI_PARENTESCO)" & vbCrLf
           
           sSql = sSql & "                       Values (" & intFilial & vbCrLf
           sSql = sSql & "                              ,'" & Trim(arrNIVEL01(intReg, 1)) & "'" & vbCrLf
           sSql = sSql & "                              ,'" & Trim(arrNIVEL01(intReg, 2)) & "'" & vbCrLf
           sSql = sSql & "                              ," & arrNIVEL01(intReg, 3) & vbCrLf
        
           strValor = Replace(arrNIVEL01(intReg, 4), ".", "")
           strValor = Replace(strValor, ",", ".")
           sSql = sSql & "                              ," & strValor & vbCrLf
           strValor = Replace(arrNIVEL01(intReg, 5), ".", "")
           strValor = Replace(strValor, ",", ".")
           sSql = sSql & "                              ," & strValor & vbCrLf
           sSql = sSql & "                              ,'" & arrNIVEL01(intReg, 6) & "'" & vbCrLf
           If Not IsNull(arrNIVEL01(intReg, 7)) And arrNIVEL01(intReg, 7) > 0 Then
              sSql = sSql & "                              ," & arrNIVEL01(intReg, 7) & vbCrLf
           Else
              sSql = sSql & "                              ,Null" & vbCrLf
           End If
           
           sSql = sSql & "                              ,'" & Trim(arrNIVEL01(intReg, 8)) & "'" & vbCrLf
           sSql = sSql & "                              )" & vbCrLf
        
           BGRV.CommandText = sSql
           BGRV.Execute
           
           ''Gravando Processo do Nivel 01
           If IsArray(arrPROCPROD01) Then
              For I = 1 To UBound(arrPROCPROD01)
                  sSql = "Insert into SGI_CADPROPROC_TEMP (SGI_FILIAL" & vbCrLf
                  sSql = sSql & "                         ,SGI_PARENTESCO" & vbCrLf
                  sSql = sSql & "                         ,SGI_CODPROC01" & vbCrLf
                  sSql = sSql & "                         ,SGI_DESCRICAOPROC01" & vbCrLf
                  sSql = sSql & "                         ,SGI_TEMPO" & vbCrLf
                  sSql = sSql & "                         ,SGI_VALORHORA" & vbCrLf
                  sSql = sSql & "                         ,SGI_TOTVALOR)" & vbCrLf
                  sSql = sSql & "                  Values (" & vbCrLf
                  sSql = sSql & "                          " & intFilial & vbCrLf
                  sSql = sSql & "                         ,'" & Trim(arrPROCPROD01(I, 1)) & "'" & vbCrLf
                  sSql = sSql & "                         ," & arrPROCPROD01(I, 2) & vbCrLf
                  sSql = sSql & "                         ,'" & Trim(arrPROCPROD01(I, 6)) & "'" & vbCrLf
               
                  strValor = Replace(arrPROCPROD01(I, 3), ".", "")
                  strValor = Replace(strValor, ",", ".")
                  sSql = sSql & "                         ," & strValor & vbCrLf
               
                  strValor = Replace(arrPROCPROD01(I, 4), ".", "")
                  strValor = Replace(strValor, ",", ".")
                  sSql = sSql & "                         ," & strValor & vbCrLf
               
                  strValor = Replace(arrPROCPROD01(I, 5), ".", "")
                  strValor = Replace(strValor, ",", ".")
                  sSql = sSql & "                         ," & strValor & vbCrLf
               
                  sSql = sSql & "                         )"
                  BGRV.CommandText = sSql
                  BGRV.Execute
              Next I
           End If
           
           
    End If
       
    adoBanco_Dados.CommitTrans
    Grava_ListMatTemp_Filho1 = True
     
    Exit Function
     
err_Trans:
     
     adoBanco_Dados.RollbackTrans
     
     Dim objErro    As Object
     Set objErro = CreateObject("BLBCWS.clsFuncoes")
     Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, "I", sSql)
     Set objErro = Nothing

End Function


Public Function Grava_ListMatTemp_Filho2(intReg As Integer)

On Error GoTo err_Trans

     Grava_ListMatTemp_Filho2 = False
     
     Dim I As Integer
     Dim strValor As String
         
     BGRV.ActiveConnection = adoBanco_Dados
     adoBanco_Dados.BeginTrans
        
    '' -----------------------------------
    '' Gravando Filho 02
    If IsArray(arrNIVEL02) Then
     
           sSql = "Insert Into SGI_LISTAMAT_TEMP (SGI_FILIAL" & vbCrLf
           sSql = sSql & "                       ,SGI_PRODUTO2" & vbCrLf
           sSql = sSql & "                       ,SGI_DESCRICAO" & vbCrLf
           sSql = sSql & "                       ,SGI_QTDE" & vbCrLf
           sSql = sSql & "                       ,SGI_CUSTOUNI" & vbCrLf
           sSql = sSql & "                       ,SGI_CUSTOTOTAL" & vbCrLf
           sSql = sSql & "                       ,SGI_UNIDCONS" & vbCrLf
           sSql = sSql & "                       ,SGI_CODTABCONV" & vbCrLf
           sSql = sSql & "                       ,SGI_PARENTESCO)" & vbCrLf
           sSql = sSql & "                       Values (" & intFilial & vbCrLf
           sSql = sSql & "                              ,'" & Trim(arrNIVEL02(intReg, 1)) & "'" & vbCrLf
           sSql = sSql & "                              ,'" & Trim(arrNIVEL02(intReg, 2)) & "'" & vbCrLf
           sSql = sSql & "                              ," & arrNIVEL02(intReg, 3) & vbCrLf
        
           strValor = Replace(arrNIVEL02(intReg, 4), ".", "")
           strValor = Replace(strValor, ",", ".")
           sSql = sSql & "                              ," & strValor & vbCrLf
           strValor = Replace(arrNIVEL02(intReg, 5), ".", "")
           strValor = Replace(strValor, ",", ".")
           sSql = sSql & "                              ," & strValor & vbCrLf
           sSql = sSql & "                              ,'" & arrNIVEL02(intReg, 6) & "'" & vbCrLf
           If Not IsNull(arrNIVEL02(intReg, 7)) And arrNIVEL02(intReg, 7) > 0 Then
              sSql = sSql & "                              ," & arrNIVEL02(intReg, 7) & vbCrLf
           Else
              sSql = sSql & "                              ,Null" & vbCrLf
           End If
           sSql = sSql & "                              ,'" & Trim(arrNIVEL02(intReg, 8)) & "'" & vbCrLf
           sSql = sSql & "                              )" & vbCrLf
           
           BGRV.CommandText = sSql
           BGRV.Execute
    End If
       
    adoBanco_Dados.CommitTrans
    Grava_ListMatTemp_Filho2 = True
     
    Exit Function
     
err_Trans:
     
     adoBanco_Dados.RollbackTrans
     
     Dim objErro    As Object
     Set objErro = CreateObject("BLBCWS.clsFuncoes")
     Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, "I", sSql)
     Set objErro = Nothing

End Function

Public Function Grava_ListMatTemp_Filho3(intReg As Integer)

On Error GoTo err_Trans

     Grava_ListMatTemp_Filho3 = False
     
     Dim I As Integer
     Dim strValor As String
         
     BGRV.ActiveConnection = adoBanco_Dados
     adoBanco_Dados.BeginTrans
        
    '' -----------------------------------
    '' Gravando Filho 03
    If IsArray(arrNIVEL03) Then
     
           sSql = "Insert Into SGI_LISTAMAT_TEMP (SGI_FILIAL" & vbCrLf
           sSql = sSql & "                       ,SGI_PRODUTO3" & vbCrLf
           sSql = sSql & "                       ,SGI_DESCRICAO" & vbCrLf
           sSql = sSql & "                       ,SGI_QTDE" & vbCrLf
           sSql = sSql & "                       ,SGI_CUSTOUNI" & vbCrLf
           sSql = sSql & "                       ,SGI_CUSTOTOTAL" & vbCrLf
           sSql = sSql & "                       ,SGI_UNIDCONS" & vbCrLf
           sSql = sSql & "                       ,SGI_CODTABCONV"
           sSql = sSql & "                       ,SGI_PARENTESCO)" & vbCrLf
           sSql = sSql & "                       Values (" & intFilial & vbCrLf
           sSql = sSql & "                              ,'" & Trim(arrNIVEL03(intReg, 1)) & "'" & vbCrLf
           sSql = sSql & "                              ,'" & Trim(arrNIVEL03(intReg, 2)) & "'" & vbCrLf
           sSql = sSql & "                              ," & arrNIVEL03(intReg, 3) & vbCrLf
        
           strValor = Replace(arrNIVEL03(intReg, 4), ".", "")
           strValor = Replace(strValor, ",", ".")
           sSql = sSql & "                              ," & strValor & vbCrLf
           strValor = Replace(arrNIVEL03(intReg, 5), ".", "")
           strValor = Replace(strValor, ",", ".")
           sSql = sSql & "                              ," & strValor & vbCrLf
           sSql = sSql & "                              ,'" & arrNIVEL03(intReg, 6) & "'" & vbCrLf
           If Not IsNull(arrNIVEL03(intReg, 7)) And arrNIVEL03(intReg, 7) > 0 Then
              sSql = sSql & "                              ," & arrNIVEL03(intReg, 7) & vbCrLf
           Else
              sSql = sSql & "                              ,Null" & vbCrLf
           End If
           
           sSql = sSql & "                              ,'" & Trim(arrNIVEL03(intReg, 8)) & "'" & vbCrLf
           sSql = sSql & "                              )" & vbCrLf
        
           BGRV.CommandText = sSql
           BGRV.Execute
    End If
       
    adoBanco_Dados.CommitTrans
    Grava_ListMatTemp_Filho3 = True
     
    Exit Function
     
err_Trans:
     
     adoBanco_Dados.RollbackTrans
     
     Dim objErro    As Object
     Set objErro = CreateObject("BLBCWS.clsFuncoes")
     Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, "I", sSql)
     Set objErro = Nothing

End Function

Public Function Grava_ListMatTemp_Filho4(intReg As Integer)

On Error GoTo err_Trans

     Grava_ListMatTemp_Filho4 = False
     
     Dim I As Integer
     Dim strValor As String
         
     BGRV.ActiveConnection = adoBanco_Dados
     adoBanco_Dados.BeginTrans
        
    '' -----------------------------------
    '' Gravando Filho 04
    If IsArray(arrNIVEL04) Then
     
           sSql = "Insert Into SGI_LISTAMAT_TEMP (SGI_FILIAL" & vbCrLf
           sSql = sSql & "                       ,SGI_PRODUTO4" & vbCrLf
           sSql = sSql & "                       ,SGI_DESCRICAO" & vbCrLf
           sSql = sSql & "                       ,SGI_QTDE" & vbCrLf
           sSql = sSql & "                       ,SGI_CUSTOUNI" & vbCrLf
           sSql = sSql & "                       ,SGI_CUSTOTOTAL" & vbCrLf
           sSql = sSql & "                       ,SGI_UNIDCONS" & vbCrLf
           sSql = sSql & "                       ,SGI_CODTABCONV" & vbCrLf
           sSql = sSql & "                       ,SGI_PARENTESCO)" & vbCrLf
           sSql = sSql & "                       Values (" & intFilial & vbCrLf
           sSql = sSql & "                              ,'" & Trim(arrNIVEL04(intReg, 1)) & "'" & vbCrLf
           sSql = sSql & "                              ,'" & Trim(arrNIVEL04(intReg, 2)) & "'" & vbCrLf
           sSql = sSql & "                              ," & arrNIVEL04(intReg, 3) & vbCrLf
        
           strValor = Replace(arrNIVEL04(intReg, 4), ".", "")
           strValor = Replace(strValor, ",", ".")
           sSql = sSql & "                              ," & strValor & vbCrLf
           strValor = Replace(arrNIVEL04(intReg, 5), ".", "")
           strValor = Replace(strValor, ",", ".")
           sSql = sSql & "                              ," & strValor & vbCrLf
           sSql = sSql & "                              ,'" & arrNIVEL04(intReg, 6) & "'" & vbCrLf
           If Not IsNull(arrNIVEL04(intReg, 7)) And arrNIVEL04(intReg, 7) > 0 Then
              sSql = sSql & "                              ," & arrNIVEL04(intReg, 7) & vbCrLf
           Else
              sSql = sSql & "                              ,Null" & vbCrLf
           End If
           
           sSql = sSql & "                              ,'" & Trim(arrNIVEL04(intReg, 8)) & "'" & vbCrLf
           sSql = sSql & "                              )" & vbCrLf
        
           BGRV.CommandText = sSql
           BGRV.Execute
    End If
       
    adoBanco_Dados.CommitTrans
    Grava_ListMatTemp_Filho4 = True
     
    Exit Function
     
err_Trans:
     
     adoBanco_Dados.RollbackTrans
     
     Dim objErro    As Object
     Set objErro = CreateObject("BLBCWS.clsFuncoes")
     Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, "I", sSql)
     Set objErro = Nothing

End Function

Public Function Grava_ListMatTemp_Filho5(intReg As Integer)

On Error GoTo err_Trans

     Grava_ListMatTemp_Filho5 = False
     
     Dim I As Integer
     Dim strValor As String
         
     BGRV.ActiveConnection = adoBanco_Dados
     adoBanco_Dados.BeginTrans
        
    '' -----------------------------------
    '' Gravando Filho 05
    If IsArray(arrNIVEL05) Then
     
           sSql = "Insert Into SGI_LISTAMAT_TEMP (SGI_FILIAL" & vbCrLf
           sSql = sSql & "                       ,SGI_PRODUTO5" & vbCrLf
           sSql = sSql & "                       ,SGI_DESCRICAO" & vbCrLf
           sSql = sSql & "                       ,SGI_QTDE" & vbCrLf
           sSql = sSql & "                       ,SGI_CUSTOUNI" & vbCrLf
           sSql = sSql & "                       ,SGI_CUSTOTOTAL" & vbCrLf
           sSql = sSql & "                       ,SGI_UNIDCONS" & vbCrLf
           sSql = sSql & "                       ,SGI_CODTABCONV" & vbCrLf
           sSql = sSql & "                       ,SGI_PARENTESCO)" & vbCrLf
           sSql = sSql & "                       Values (" & intFilial & vbCrLf
           sSql = sSql & "                              ,'" & Trim(arrNIVEL05(intReg, 1)) & "'" & vbCrLf
           sSql = sSql & "                              ,'" & Trim(arrNIVEL05(intReg, 2)) & "'" & vbCrLf
           sSql = sSql & "                              ," & arrNIVEL05(intReg, 3) & vbCrLf
        
           strValor = Replace(arrNIVEL05(intReg, 4), ".", "")
           strValor = Replace(strValor, ",", ".")
           sSql = sSql & "                              ," & strValor & vbCrLf
           strValor = Replace(arrNIVEL05(intReg, 5), ".", "")
           strValor = Replace(strValor, ",", ".")
           sSql = sSql & "                              ," & strValor & vbCrLf
           sSql = sSql & "                              ,'" & arrNIVEL05(intReg, 6) & "'" & vbCrLf
           If Not IsNull(arrNIVEL05(intReg, 7)) And arrNIVEL05(intReg, 7) > 0 Then
              sSql = sSql & "                              ," & arrNIVEL05(intReg, 7) & vbCrLf
           Else
              sSql = sSql & "                              ,Null" & vbCrLf
           End If
           
           sSql = sSql & "                              ,'" & Trim(arrNIVEL05(intReg, 8)) & "'" & vbCrLf
           sSql = sSql & "                              )" & vbCrLf
        
           BGRV.CommandText = sSql
           BGRV.Execute
    End If
       
    adoBanco_Dados.CommitTrans
    Grava_ListMatTemp_Filho5 = True
     
    Exit Function
     
err_Trans:
     
     adoBanco_Dados.RollbackTrans
     
     Dim objErro    As Object
     Set objErro = CreateObject("BLBCWS.clsFuncoes")
     Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, "I", sSql)
     Set objErro = Nothing

End Function


Public Function Grava_ListMatTemp_Filho6(intReg As Integer)

On Error GoTo err_Trans

     Grava_ListMatTemp_Filho6 = False
     
     Dim I As Integer
     Dim strValor As String
         
     BGRV.ActiveConnection = adoBanco_Dados
     adoBanco_Dados.BeginTrans
        
    '' -----------------------------------
    '' Gravando Filho 06
    If IsArray(arrNIVEL06) Then
     
           sSql = "Insert Into SGI_LISTAMAT_TEMP (SGI_FILIAL" & vbCrLf
           sSql = sSql & "                       ,SGI_PRODUTO5" & vbCrLf
           sSql = sSql & "                       ,SGI_DESCRICAO" & vbCrLf
           sSql = sSql & "                       ,SGI_QTDE" & vbCrLf
           sSql = sSql & "                       ,SGI_CUSTOUNI" & vbCrLf
           sSql = sSql & "                       ,SGI_CUSTOTOTAL" & vbCrLf
           sSql = sSql & "                       ,SGI_UNIDCONS" & vbCrLf
           sSql = sSql & "                       ,SGI_CODTABCONV" & vbCrLf
           sSql = sSql & "                       ,SGI_PARENTESCO)" & vbCrLf
           sSql = sSql & "                       Values (" & intFilial & vbCrLf
           sSql = sSql & "                              ,'" & Trim(arrNIVEL06(intReg, 1)) & "'" & vbCrLf
           sSql = sSql & "                              ,'" & Trim(arrNIVEL06(intReg, 2)) & "'" & vbCrLf
           sSql = sSql & "                              ," & arrNIVEL06(intReg, 3) & vbCrLf
        
           strValor = Replace(arrNIVEL06(intReg, 4), ".", "")
           strValor = Replace(strValor, ",", ".")
           sSql = sSql & "                              ," & strValor & vbCrLf
           strValor = Replace(arrNIVEL06(intReg, 5), ".", "")
           strValor = Replace(strValor, ",", ".")
           sSql = sSql & "                              ," & strValor & vbCrLf
           sSql = sSql & "                              ,'" & arrNIVEL06(intReg, 6) & "'" & vbCrLf
           If Not IsNull(arrNIVEL06(intReg, 7)) And arrNIVEL06(intReg, 7) > 0 Then
              sSql = sSql & "                              ," & arrNIVEL06(intReg, 7) & vbCrLf
           Else
              sSql = sSql & "                              ,Null" & vbCrLf
           End If
           
           sSql = sSql & "                              ,'" & Trim(arrNIVEL06(intReg, 8)) & "'" & vbCrLf
           sSql = sSql & "                              )" & vbCrLf
        
           BGRV.CommandText = sSql
           BGRV.Execute
    End If
       
    adoBanco_Dados.CommitTrans
    Grava_ListMatTemp_Filho6 = True
     
    Exit Function
     
err_Trans:
     
     adoBanco_Dados.RollbackTrans
     
     Dim objErro    As Object
     Set objErro = CreateObject("BLBCWS.clsFuncoes")
     Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, "I", sSql)
     Set objErro = Nothing

End Function


Public Function Grava_ListMatTemp_Filho7(intReg As Integer)

On Error GoTo err_Trans

     Grava_ListMatTemp_Filho7 = False
     
     Dim I As Integer
     Dim strValor As String
         
     BGRV.ActiveConnection = adoBanco_Dados
     adoBanco_Dados.BeginTrans
        
    '' -----------------------------------
    '' Gravando Filho 07
    If IsArray(arrNIVEL07) Then
     
           sSql = "Insert Into SGI_LISTAMAT_TEMP (SGI_FILIAL" & vbCrLf
           sSql = sSql & "                       ,SGI_PRODUTO5" & vbCrLf
           sSql = sSql & "                       ,SGI_DESCRICAO" & vbCrLf
           sSql = sSql & "                       ,SGI_QTDE" & vbCrLf
           sSql = sSql & "                       ,SGI_CUSTOUNI" & vbCrLf
           sSql = sSql & "                       ,SGI_CUSTOTOTAL" & vbCrLf
           sSql = sSql & "                       ,SGI_UNIDCONS" & vbCrLf
           sSql = sSql & "                       ,SGI_CODTABCONV" & vbCrLf
           sSql = sSql & "                       ,SGI_PARENTESCO)" & vbCrLf
           sSql = sSql & "                       Values (" & intFilial & vbCrLf
           sSql = sSql & "                              ,'" & Trim(arrNIVEL07(intReg, 1)) & "'" & vbCrLf
           sSql = sSql & "                              ,'" & Trim(arrNIVEL07(intReg, 2)) & "'" & vbCrLf
           sSql = sSql & "                              ," & arrNIVEL07(intReg, 3) & vbCrLf
        
           strValor = Replace(arrNIVEL07(intReg, 4), ".", "")
           strValor = Replace(strValor, ",", ".")
           sSql = sSql & "                              ," & strValor & vbCrLf
           strValor = Replace(arrNIVEL07(intReg, 5), ".", "")
           strValor = Replace(strValor, ",", ".")
           sSql = sSql & "                              ," & strValor & vbCrLf
           sSql = sSql & "                              ,'" & arrNIVEL07(intReg, 6) & "'" & vbCrLf
           If Not IsNull(arrNIVEL07(intReg, 7)) And arrNIVEL07(intReg, 7) > 0 Then
              sSql = sSql & "                              ," & arrNIVEL07(intReg, 7) & vbCrLf
           Else
              sSql = sSql & "                              ,Null" & vbCrLf
           End If
           
           sSql = sSql & "                              ,'" & Trim(arrNIVEL07(intReg, 8)) & "'" & vbCrLf
           sSql = sSql & "                              )" & vbCrLf
        
           BGRV.CommandText = sSql
           BGRV.Execute
    End If
       
    adoBanco_Dados.CommitTrans
    Grava_ListMatTemp_Filho7 = True
     
    Exit Function
     
err_Trans:
     
     adoBanco_Dados.RollbackTrans
     
     Dim objErro    As Object
     Set objErro = CreateObject("BLBCWS.clsFuncoes")
     Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, "I", sSql)
     Set objErro = Nothing

End Function


Public Function Grava_ListMatTemp_Filho8(intReg As Integer)

On Error GoTo err_Trans

     Grava_ListMatTemp_Filho8 = False
     
     Dim I As Integer
     Dim strValor As String
         
     BGRV.ActiveConnection = adoBanco_Dados
     adoBanco_Dados.BeginTrans
        
    '' -----------------------------------
    '' Gravando Filho 08
    If IsArray(arrNIVEL08) Then
     
           sSql = "Insert Into SGI_LISTAMAT_TEMP (SGI_FILIAL" & vbCrLf
           sSql = sSql & "                       ,SGI_PRODUTO5" & vbCrLf
           sSql = sSql & "                       ,SGI_DESCRICAO" & vbCrLf
           sSql = sSql & "                       ,SGI_QTDE" & vbCrLf
           sSql = sSql & "                       ,SGI_CUSTOUNI" & vbCrLf
           sSql = sSql & "                       ,SGI_CUSTOTOTAL" & vbCrLf
           sSql = sSql & "                       ,SGI_UNIDCONS" & vbCrLf
           sSql = sSql & "                       ,SGI_CODTABCONV" & vbCrLf
           sSql = sSql & "                       ,SGI_PARENTESCO)" & vbCrLf
           sSql = sSql & "                       Values (" & intFilial & vbCrLf
           sSql = sSql & "                              ,'" & Trim(arrNIVEL08(intReg, 1)) & "'" & vbCrLf
           sSql = sSql & "                              ,'" & Trim(arrNIVEL08(intReg, 2)) & "'" & vbCrLf
           sSql = sSql & "                              ," & arrNIVEL08(intReg, 3) & vbCrLf
        
           strValor = Replace(arrNIVEL08(intReg, 4), ".", "")
           strValor = Replace(strValor, ",", ".")
           sSql = sSql & "                              ," & strValor & vbCrLf
           strValor = Replace(arrNIVEL08(intReg, 5), ".", "")
           strValor = Replace(strValor, ",", ".")
           sSql = sSql & "                              ," & strValor & vbCrLf
           sSql = sSql & "                              ,'" & arrNIVEL08(intReg, 6) & "'" & vbCrLf
           If Not IsNull(arrNIVEL08(intReg, 7)) And arrNIVEL08(intReg, 7) > 0 Then
              sSql = sSql & "                              ," & arrNIVEL08(intReg, 7) & vbCrLf
           Else
              sSql = sSql & "                              ,Null" & vbCrLf
           End If
           
           sSql = sSql & "                              ,'" & Trim(arrNIVEL08(intReg, 8)) & "'" & vbCrLf
           sSql = sSql & "                              )" & vbCrLf
        
           BGRV.CommandText = sSql
           BGRV.Execute
    End If
       
    adoBanco_Dados.CommitTrans
    Grava_ListMatTemp_Filho8 = True
     
    Exit Function
     
err_Trans:
     
     adoBanco_Dados.RollbackTrans
     
     Dim objErro    As Object
     Set objErro = CreateObject("BLBCWS.clsFuncoes")
     Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, "I", sSql)
     Set objErro = Nothing

End Function


Public Function Grava_ListMatTemp_Filho9(intReg As Integer)

On Error GoTo err_Trans

     Grava_ListMatTemp_Filho9 = False
     
     Dim I As Integer
     Dim strValor As String
         
     BGRV.ActiveConnection = adoBanco_Dados
     adoBanco_Dados.BeginTrans
        
    '' -----------------------------------
    '' Gravando Filho 09
    If IsArray(arrNIVEL09) Then
     
           sSql = "Insert Into SGI_LISTAMAT_TEMP (SGI_FILIAL" & vbCrLf
           sSql = sSql & "                       ,SGI_PRODUTO5" & vbCrLf
           sSql = sSql & "                       ,SGI_DESCRICAO" & vbCrLf
           sSql = sSql & "                       ,SGI_QTDE" & vbCrLf
           sSql = sSql & "                       ,SGI_CUSTOUNI" & vbCrLf
           sSql = sSql & "                       ,SGI_CUSTOTOTAL" & vbCrLf
           sSql = sSql & "                       ,SGI_UNIDCONS" & vbCrLf
           sSql = sSql & "                       ,SGI_CODTABCONV" & vbCrLf
           sSql = sSql & "                       ,SGI_PARENTESCO)" & vbCrLf
           sSql = sSql & "                       Values (" & intFilial & vbCrLf
           sSql = sSql & "                              ,'" & Trim(arrNIVEL09(intReg, 1)) & "'" & vbCrLf
           sSql = sSql & "                              ,'" & Trim(arrNIVEL09(intReg, 2)) & "'" & vbCrLf
           sSql = sSql & "                              ," & arrNIVEL09(intReg, 3) & vbCrLf
        
           strValor = Replace(arrNIVEL09(intReg, 4), ".", "")
           strValor = Replace(strValor, ",", ".")
           sSql = sSql & "                              ," & strValor & vbCrLf
           strValor = Replace(arrNIVEL09(intReg, 5), ".", "")
           strValor = Replace(strValor, ",", ".")
           sSql = sSql & "                              ," & strValor & vbCrLf
           sSql = sSql & "                              ,'" & arrNIVEL09(intReg, 6) & "'" & vbCrLf
           If Not IsNull(arrNIVEL09(intReg, 7)) And arrNIVEL09(intReg, 7) > 0 Then
              sSql = sSql & "                              ," & arrNIVEL09(intReg, 7) & vbCrLf
           Else
              sSql = sSql & "                              ,Null" & vbCrLf
           End If
           
           sSql = sSql & "                              ,'" & Trim(arrNIVEL09(intReg, 8)) & "'" & vbCrLf
           sSql = sSql & "                              )" & vbCrLf
        
           BGRV.CommandText = sSql
           BGRV.Execute
    End If
       
    adoBanco_Dados.CommitTrans
    Grava_ListMatTemp_Filho9 = True
     
    Exit Function
     
err_Trans:
     
     adoBanco_Dados.RollbackTrans
     
     Dim objErro    As Object
     Set objErro = CreateObject("BLBCWS.clsFuncoes")
     Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, "I", sSql)
     Set objErro = Nothing

End Function

