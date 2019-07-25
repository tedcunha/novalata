VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADFLUXCAIXA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim intFilial       As Long
Dim lngCODFLXCAIXA  As Long
Dim dtDATAFLXCAIXA  As Date
Dim strSTATUS       As String
Dim ccurSALDOANT    As Currency
Dim ccurSALDOATU    As Currency
Dim arrLCTOFLXCAIXA As Variant
Dim arrFLUXOCAIXA   As Variant

Public Property Get SALDOATU() As Currency
    SALDOATU = ccurSALDOATU
End Property
Public Property Let SALDOATU(ByVal vData As Currency)
    ccurSALDOATU = vData
End Property

Public Property Get SALDOANT() As Currency
    SALDOANT = ccurSALDOANT
End Property
Public Property Let SALDOANT(ByVal vData As Currency)
    ccurSALDOANT = vData
End Property

Public Property Get FLUXOCAIXA() As Variant
    FLUXOCAIXA = arrFLUXOCAIXA
End Property
Public Property Let FLUXOCAIXA(ByVal vData As Variant)
    arrFLUXOCAIXA = vData
End Property
Public Property Get LCTOFLXCAIXA() As Variant
    LCTOFLXCAIXA = arrLCTOFLXCAIXA
End Property
Public Property Let LCTOFLXCAIXA(ByVal vData As Variant)
    arrLCTOFLXCAIXA = vData
End Property
Public Property Get STATUS() As String
    STATUS = strSTATUS
End Property
Public Property Let STATUS(ByVal vData As String)
    strSTATUS = vData
End Property
Public Property Get DATAFLXCAIXA() As Date
    DATAFLXCAIXA = dtDATAFLXCAIXA
End Property
Public Property Let DATAFLXCAIXA(ByVal vData As Date)
    dtDATAFLXCAIXA = vData
End Property

Public Property Get CODFLXCAIXA() As Long
    CODFLXCAIXA = lngCODFLXCAIXA
End Property

Public Property Let CODFLXCAIXA(ByVal vData As Long)
    lngCODFLXCAIXA = vData
End Property

Public Property Get FILIAL() As Long
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Long)
    intFilial = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADFLUXCAIXAP.cCaminho = cCaminho
    frmCADFLUXCAIXAP.Linha = Linha
    frmCADFLUXCAIXAP.FILIAL = iFilial
    frmCADFLUXCAIXAP.strAcesso = strAcesso
    frmCADFLUXCAIXAP.Show vbModal
End Sub


Public Function Pesq_CadFluxCaixa() As Boolean

   Pesq_CadFluxCaixa = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       *" & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_CADFLXCXHEADER " & vbCrLf
   sSql = sSql + " Where "
   sSql = sSql + "       SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_CadFluxCaixa = True
   BREC.Close

End Function

Public Sub PreencheComboTipoPgto(Combo As Variant)

     Combo.Clear
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       * " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADTIPOPGTO " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL  = " & intFilial & vbCrLf
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_CODIGO) & Space(10 - Len(Trim(BREC!SGI_CODIGO))) & " - " & Trim(BREC!SGI_DESCRICAO) & " - " & IIf(BREC!SGI_OPERACAO = 1, "DEBITO", "CRÉDITO")
        Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
        BREC.MoveNext
     Loop
    
     BREC.Close
    
End Sub

Public Sub PreenchComboBancos(Combo As Variant)

    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADBANCOS " & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "Order by SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       Combo.AddItem Space(5 - Len(Trim(Str(BREC!SGI_CODIGO)))) & Trim(Str(BREC!SGI_CODIGO)) & " - " & Trim(BREC!SGI_AGENCIA) & Space(10 - Len(Trim(BREC!SGI_AGENCIA))) & " - " & Trim(BREC!SGI_CC) & Space(20 - Len(Trim(BREC!SGI_CC))) & " - " & Trim(BREC!SGI_DESCRICAO)
       Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
       BREC.MoveNext
    Loop
    
    BREC.Close

End Sub

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

On Error GoTo grav_err
     
     GRAVA = False
     
     Dim I      As Integer
     Dim sValor As String
     
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Then
        
Altera:
        
        ''' -------------------------------------------------
        ''' Grava Fluxo de Caixa Cabeçalho
        sSql = "Insert Into SGI_CADFLXCXHEADER (" & vbCrLf
        sSql = sSql & "                                 SGI_FILIAL " & vbCrLf
        sSql = sSql & "                                ,SGI_CODIGO " & vbCrLf
        sSql = sSql & "                                ,SGI_DATA   " & vbCrLf
        sSql = sSql & "                                ,SGI_STATUS " & vbCrLf
        sSql = sSql & "                                ,SGI_SALDANT" & vbCrLf
        sSql = sSql & "                                ,SGI_SALDATU" & vbCrLf
        sSql = sSql & "                               )" & vbCrLf
        sSql = sSql & "                        Values (" & vbCrLf
        sSql = sSql & "                                 " & intFilial & vbCrLf
        sSql = sSql & "                                ," & lngCODFLXCAIXA & vbCrLf
        sSql = sSql & "                                ,'" & Format(dtDATAFLXCAIXA, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                                ,'" & strSTATUS & "'" & vbCrLf
                   
        sValor = Replace(ccurSALDOANT, ".", "")
        sValor = Replace(Trim(sValor), ",", ".")
        sSql = sSql & "                                ," & sValor & vbCrLf
        
        sValor = Replace(ccurSALDOATU, ".", "")
        sValor = Replace(Trim(sValor), ",", ".")
        sSql = sSql & "                                ," & sValor & vbCrLf
        
        sSql = sSql & "                               )"
         
        BGRV.CommandText = sSql
        BGRV.Execute
        
        ''' -------------------------------------------------
        ''' Grava Lançamentos
        If IsArray(arrLCTOFLXCAIXA) = True Then
           For I = 1 To UBound(arrLCTOFLXCAIXA)
                            
               sSql = "Insert into SGI_CADFLXCXLCTOS (" & vbCrLf
               sSql = sSql & "                                SGI_FILIAL    " & vbCrLf
               sSql = sSql & "                               ,SGI_CODIGO    " & vbCrLf
               sSql = sSql & "                               ,SGI_DATA      " & vbCrLf
               sSql = sSql & "                               ,SGI_CODTPPGTO " & vbCrLf
               sSql = sSql & "                               ,SGI_CODBCO    " & vbCrLf
               sSql = sSql & "                               ,SGI_NDOC      " & vbCrLf
               sSql = sSql & "                               ,SGI_VALOR     " & vbCrLf
               sSql = sSql & "                              )" & vbCrLf
               sSql = sSql & "                       Values (" & vbCrLf
               sSql = sSql & "                                " & intFilial & vbCrLf
               sSql = sSql & "                               ," & lngCODFLXCAIXA & vbCrLf
               sSql = sSql & "                               ,'" & Format(arrLCTOFLXCAIXA(I, 3), "MM/DD/YYYY") & "'" & vbCrLf
               sSql = sSql & "                               ," & arrLCTOFLXCAIXA(I, 1) & vbCrLf
               sSql = sSql & "                               ," & arrLCTOFLXCAIXA(I, 2) & vbCrLf
               sSql = sSql & "                               ,'" & arrLCTOFLXCAIXA(I, 4) & "'" & vbCrLf
               
               sValor = Replace(arrLCTOFLXCAIXA(I, 5), ".", "")
               sValor = Replace(Trim(sValor), ",", ".")
               
               sSql = sSql & "                               ," & sValor & vbCrLf
               sSql = sSql & "                              )"
               
               BGRV.CommandText = sSql
               BGRV.Execute
               
           Next I
        End If
        
        ''' -------------------------------------------------
        ''' Grava Fluxo de Caixa
        If IsArray(arrFLUXOCAIXA) = True Then
           For I = 1 To UBound(arrFLUXOCAIXA)
           
                If Len(Trim(arrFLUXOCAIXA(I, 1))) > 0 Then
                    
                   sSql = "Insert into SGI_CADFLXCAIXA (" & vbCrLf
                   sSql = sSql & "                              SGI_FILIAL    " & vbCrLf
                   sSql = sSql & "                             ,SGI_CODIGO    " & vbCrLf
                   sSql = sSql & "                             ,SGI_DATA      " & vbCrLf
                   sSql = sSql & "                             ,SGI_HISTORICO " & vbCrLf
                   sSql = sSql & "                             ,SGI_VALOR     " & vbCrLf
                   sSql = sSql & "                             ,SGI_CODBCO    " & vbCrLf
                   sSql = sSql & "                            )" & vbCrLf
                   sSql = sSql & "                     Values (" & vbCrLf
                   sSql = sSql & "                              " & intFilial & vbCrLf
                   sSql = sSql & "                             ," & lngCODFLXCAIXA & vbCrLf
                   sSql = sSql & "                             ,'" & Format(arrFLUXOCAIXA(I, 2), "MM/DD/YYYY") & "'" & vbCrLf
                   sSql = sSql & "                             ,'" & arrFLUXOCAIXA(I, 1) & "'" & vbCrLf
                   
                   sValor = Replace(arrFLUXOCAIXA(I, 3), ".", "")
                   sValor = Replace(Trim(sValor), ",", ".")
                   
                   sSql = sSql & "                             ," & sValor & vbCrLf
                   sSql = sSql & "                             ," & arrFLUXOCAIXA(I, 4) & vbCrLf
                   sSql = sSql & "                            )"
               
                   BGRV.CommandText = sSql
                   BGRV.Execute
                   
                End If
               
           Next I
        End If
        
        '' ----------------------------------------------------------------------------
        '' Relacionando o Documento ao Caixa
        '' Contas a receber
        sSql = "Update SGI_CONTASIARC Set SGI_CADCAIXA = " & lngCODFLXCAIXA & vbCrLf
        sSql = sSql & "                    Where " & vbCrLf
        sSql = sSql & "                          SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "                      And SGI_DTPGTO = '" & Format(dtDATAFLXCAIXA, "MM/DD/YYYY") & "'"
        
        BGRV.CommandText = sSql
        BGRV.Execute
     
        '' Contas a Pagar Lotes
        sSql = "Update SGI_CADLOTEHEADER Set SGI_CODCAIXA = " & lngCODFLXCAIXA & vbCrLf
        sSql = sSql & "                             Where " & vbCrLf
        sSql = sSql & "                                   SGI_FILIAL   = " & FILIAL & vbCrLf
        sSql = sSql & "                               And SGI_DATALOTE = '" & Format(dtDATAFLXCAIXA, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                               And SGI_STATUS   = 'B'"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Contas a Pagar Sem Lotes
        sSql = "Update SGI_CONTASIAPG Set SGI_CODCAIXA = " & lngCODFLXCAIXA & vbCrLf
        sSql = sSql & "                                 Where " & vbCrLf
        sSql = sSql & "                                       SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "                                   And SGI_STATUS = 'B' " & vbCrLf
        sSql = sSql & "                                   And SGI_NLOTE is NUll" & vbCrLf
        sSql = sSql & "                                   And SGI_DTPGTO = '" & Format(dtDATAFLXCAIXA, "MM/DD/YYYY") & "'"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' ----------------------------------------------------------------------------
     
     ElseIf strOPER = "A" Then
        
        '' ----------------------------------------------------------------------------
        '' Extornando
        '' Contas a receber
        sSql = "Update SGI_CONTASIARC Set SGI_CADCAIXA = NULL" & vbCrLf
        sSql = sSql & "                    Where " & vbCrLf
        sSql = sSql & "                          SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "                      And SGI_DTPGTO = '" & Format(dtDATAFLXCAIXA, "MM/DD/YYYY") & "'"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Contas a Pagar Lotes
        sSql = "Update SGI_CADLOTEHEADER Set SGI_CODCAIXA = NULL" & vbCrLf
        sSql = sSql & "                             Where " & vbCrLf
        sSql = sSql & "                                   SGI_FILIAL   = " & FILIAL & vbCrLf
        sSql = sSql & "                               And SGI_DATALOTE = '" & Format(dtDATAFLXCAIXA, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                               And SGI_STATUS   = 'B'"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        
        '' Contas a Pagar Sem Lotes
        sSql = "Update SGI_CONTASIAPG Set SGI_CODCAIXA = NULL" & vbCrLf
        sSql = sSql & "                                 Where " & vbCrLf
        sSql = sSql & "                                       SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "                                   And SGI_STATUS = 'B' " & vbCrLf
        sSql = sSql & "                                   And SGI_NLOTE is NUll" & vbCrLf
        sSql = sSql & "                                   And SGI_DTPGTO = '" & Format(dtDATAFLXCAIXA, "MM/DD/YYYY") & "'"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' ----------------------------------------------------------------------------
        
        '' Apaga Caixa
        sSql = "Delete From SGI_CADFLXCAIXA " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODFLXCAIXA
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        
        '' Apaga Lançamentos
        sSql = "Delete From SGI_CADFLXCXLCTOS " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODFLXCAIXA
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Apaga Lançamentos
        sSql = "Delete From SGI_CADFLXCXHEADER " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODFLXCAIXA
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        GoTo Altera
     
     ElseIf strOPER = "E" Then
     
        '' ----------------------------------------------------------------------------
        '' Extornando
        '' Contas a receber
        sSql = "Update SGI_CONTASIARC Set SGI_CADCAIXA = NULL" & vbCrLf
        sSql = sSql & "                    Where " & vbCrLf
        sSql = sSql & "                          SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "                      And SGI_DTPGTO = '" & Format(dtDATAFLXCAIXA, "MM/DD/YYYY") & "'"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Contas a Pagar Lotes
        sSql = "Update SGI_CADLOTEHEADER Set SGI_CODCAIXA = NULL" & vbCrLf
        sSql = sSql & "                             Where " & vbCrLf
        sSql = sSql & "                                   SGI_FILIAL   = " & FILIAL & vbCrLf
        sSql = sSql & "                               And SGI_DATALOTE = '" & Format(dtDATAFLXCAIXA, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                               And SGI_STATUS   = 'B'"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        
        '' Contas a Pagar Sem Lotes
        sSql = "Update SGI_CONTASIAPG Set SGI_CODCAIXA = NULL" & vbCrLf
        sSql = sSql & "                                 Where " & vbCrLf
        sSql = sSql & "                                       SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "                                   And SGI_STATUS = 'B' " & vbCrLf
        sSql = sSql & "                                   And SGI_NLOTE is NUll" & vbCrLf
        sSql = sSql & "                                   And SGI_DTPGTO = '" & Format(dtDATAFLXCAIXA, "MM/DD/YYYY") & "'"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' ----------------------------------------------------------------------------
        
        '' Apaga Caixa
        sSql = "Delete From SGI_CADFLXCAIXA " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODFLXCAIXA
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        
        '' Apaga Lançamentos
        sSql = "Delete From SGI_CADFLXCXLCTOS " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODFLXCAIXA
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Apaga Lançamentos
        sSql = "Delete From SGI_CADFLXCXHEADER " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODFLXCAIXA
        
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

Public Function Carrega_campos() As Boolean

   Carrega_campos = False
   
   Dim intReg As Integer
   Dim I      As Integer
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADFLXCXHEADER " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & lngCODFLXCAIXA & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC.EOF Then
      
      lngCODFLXCAIXA = BREC!SGI_CODIGO
      dtDATAFLXCAIXA = BREC!SGI_DATA
      strSTATUS = BREC!SGI_STATUS
      ccurSALDOANT = BREC!SGI_SALDANT
      ccurSALDOATU = BREC!SGI_SALDATU
      
      Carrega_campos = True
   End If
   
   BREC.Close
   
   If Carrega_campos = False Then Exit Function
   
   '' -----------------------------------------------------
   BREC2.ActiveConnection = adoBanco_Dados
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADFLXCXLCTOS " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial
   sSql = sSql & "   And SGI_CODIGO = " & lngCODFLXCAIXA & vbCrLf
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC2.EOF Then
   
      intReg = 0
      Do While Not BREC2.EOF
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
      
      ReDim arrLCTOFLXCAIXA(1 To intReg, 1 To 5)
      
      intReg = 1
      BREC2.MoveFirst
      Do While Not BREC2.EOF
         arrLCTOFLXCAIXA(intReg, 1) = BREC2!SGI_CODTPPGTO
         arrLCTOFLXCAIXA(intReg, 2) = BREC2!SGI_CODBCO
         arrLCTOFLXCAIXA(intReg, 3) = BREC2!SGI_DATA
         arrLCTOFLXCAIXA(intReg, 4) = BREC2!SGI_NDOC
         arrLCTOFLXCAIXA(intReg, 5) = BREC2!SGI_VALOR
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
   
   End If
   
   BREC2.Close
   
   '' -----------------------------------------------------
   '' Movimentação de Caixa
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADFLXCAIXA " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial
   sSql = sSql & "   And SGI_CODIGO = " & lngCODFLXCAIXA & vbCrLf
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC2.EOF Then
   
      intReg = 0
      Do While Not BREC2.EOF
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
      
      ReDim arrFLUXOCAIXA(1 To intReg, 1 To 4)
      
      intReg = 1
      BREC2.MoveFirst
      Do While Not BREC2.EOF
         arrFLUXOCAIXA(intReg, 1) = BREC2!SGI_HISTORICO
         arrFLUXOCAIXA(intReg, 2) = BREC2!SGI_DATA
         arrFLUXOCAIXA(intReg, 3) = BREC2!SGI_VALOR
         arrFLUXOCAIXA(intReg, 4) = BREC2!SGI_CODBCO
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
   
   End If
   
   BREC2.Close
   
   
   
   

End Function

