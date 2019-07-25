VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADCONTASAPG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim intFilial       As Integer
Dim lngCODPGTO      As Long
Dim dtDATALCTO      As Date
Dim lngCODFORN      As Long
Dim lngCODCONDPGTO  As Long
Dim lngCODGRPDESP   As Long
Dim curVLTOTLCTO    As Currency
Dim strDOCPAI       As String
Dim arrDOCPGTO      As Variant
Dim curVLPGTO       As Currency
Dim dtDTPAGTO       As Date
Dim lngTIPOPGTO     As Long
Dim curDESCONTOPGTO As Currency
Dim curPORCDESC     As Currency
Dim curACRESCPGTO   As Currency
Dim curPORCACRES    As Currency
Dim intPARCPGTO     As Integer
Dim strNUMDOC       As String
Dim lngCODLOTE      As Long
Dim dtDATALOTE      As Date
Dim lngCODBANCO     As Long
Dim strSTATUS       As String

Public Property Get STATUS() As String
    STATUS = strSTATUS
End Property

Public Property Let STATUS(ByVal vData As String)
    strSTATUS = vData
End Property

Public Property Get CODLOTE() As Long
    CODLOTE = lngCODLOTE
End Property

Public Property Let CODLOTE(ByVal vData As Long)
    lngCODLOTE = vData
End Property

Public Property Get DATALOTE() As Date
    DATALOTE = dtDATALOTE
End Property

Public Property Let DATALOTE(ByVal vData As Date)
    dtDATALOTE = vData
End Property

Public Property Get CODBANCO() As Long
    CODBANCO = lngCODBANCO
End Property

Public Property Let CODBANCO(ByVal vData As Long)
    lngCODBANCO = vData
End Property

Public Property Get NUMDOC() As String
    NUMDOC = strNUMDOC
End Property

Public Property Let NUMDOC(ByVal vData As String)
    strNUMDOC = vData
End Property

Public Property Get PARCPGTO() As Integer
    PARCPGTO = intPARCPGTO
End Property

Public Property Let PARCPGTO(ByVal vData As Integer)
    intPARCPGTO = vData
End Property

Public Property Get PORCACRES() As Currency
    PORCACRES = curPORCACRES
End Property

Public Property Let PORCACRES(ByVal vData As Currency)
    curPORCACRES = vData
End Property

Public Property Get ACRESCPGTO() As Currency
    ACRESCPGTO = curACRESCPGTO
End Property

Public Property Let ACRESCPGTO(ByVal vData As Currency)
    curACRESCPGTO = vData
End Property

Public Property Get PORCDESC() As Currency
    PORCDESC = curPORCDESC
End Property

Public Property Let PORCDESC(ByVal vData As Currency)
    curPORCDESC = vData
End Property

Public Property Get DESCONTOPGTO() As Currency
    DESCONTOPGTO = curDESCONTOPGTO
End Property

Public Property Let DESCONTOPGTO(ByVal vData As Currency)
    curDESCONTOPGTO = vData
End Property

Public Property Get TIPOPGTO() As Long
    TIPOPGTO = lngTIPOPGTO
End Property

Public Property Let TIPOPGTO(ByVal vData As Long)
    lngTIPOPGTO = vData
End Property

Public Property Get DTPAGTO() As Date
    DTPAGTO = dtDTPAGTO
End Property

Public Property Let DTPAGTO(ByVal vData As Date)
    dtDTPAGTO = vData
End Property

Public Property Get VLPGTO() As Currency
    VLPGTO = curVLPGTO
End Property

Public Property Let VLPGTO(ByVal vData As Currency)
    curVLPGTO = vData
End Property

Public Property Get DOCPGTO() As Variant
    DOCPGTO = arrDOCPGTO
End Property

Public Property Let DOCPGTO(ByVal vData As Variant)
    arrDOCPGTO = vData
End Property

Public Property Get DOCPAI() As String
    DOCPAI = strDOCPAI
End Property

Public Property Let DOCPAI(ByVal vData As String)
    strDOCPAI = vData
End Property

Public Property Get VLTOTLCTO() As Currency
    VLTOTLCTO = curVLTOTLCTO
End Property

Public Property Let VLTOTLCTO(ByVal vData As Currency)
    curVLTOTLCTO = vData
End Property

Public Property Get DATALCTO() As Date
    DATALCTO = dtDATALCTO
End Property

Public Property Let DATALCTO(ByVal vData As Date)
    dtDATALCTO = vData
End Property

Public Property Get CODGRPDESP() As Long
    CODGRPDESP = lngCODGRPDESP
End Property

Public Property Let CODGRPDESP(ByVal vData As Long)
    lngCODGRPDESP = vData
End Property

Public Property Get CODCONDPGTO() As Long
    CODCONDPGTO = lngCODCONDPGTO
End Property

Public Property Let CODCONDPGTO(ByVal vData As Long)
    lngCODCONDPGTO = vData
End Property

Public Property Get CODFORN() As Long
    CODFORN = lngCODFORN
End Property

Public Property Let CODFORN(ByVal vData As Long)
    lngCODFORN = vData
End Property

Public Property Get CODPGTO() As Long
    CODPGTO = lngCODPGTO
End Property

Public Property Let CODPGTO(ByVal vData As Long)
    lngCODPGTO = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADCONTASAPGP.cCaminho = cCaminho
    frmCADCONTASAPGP.Linha = Linha
    frmCADCONTASAPGP.FILIAL = iFilial
    frmCADCONTASAPGP.strAcesso = strAcesso
    frmCADCONTASAPGP.strUSUARIO = strUSUARIO
    frmCADCONTASAPGP.Show vbModal
End Sub

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Function Pesq_CadContasAPG() As Boolean

   If BREC.State = 1 Then BREC.Close
   
   Pesq_CadContasAPG = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       * " & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_CONTASHAPG " & vbCrLf
   sSql = sSql + " Where " & vbCrLf
   sSql = sSql + "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_CadContasAPG = True
   BREC.Close

End Function

Public Sub PreencheComboFornec(Combo As Variant)

     sSql = "Select " & vbCrLf
     sSql = sSql & "       FORN.* " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADFORNEC  FORN " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       FORN.SGI_FILIAL  = " & intFilial & vbCrLf
     sSql = sSql & "   And FORN.SGI_STATUS  IN(1,2)" & vbCrLf
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_CODIGO) & Space(10 - Len(Trim(BREC!SGI_CODIGO))) & " - " & Trim(BREC!SGI_RAZAOSOC)
        Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
        BREC.MoveNext
     Loop
    
     ''If Combo.ListCount > 0 Then Combo.ListIndex = 0
    
     BREC.Close
    
End Sub

Public Sub PreencheComboCondPgto(Combo As Variant)

     sSql = "Select " & vbCrLf
     sSql = sSql & "       PGTO.* " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADCONDPGTO  PGTO " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       PGTO.SGI_FILIAL  = " & intFilial & vbCrLf
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_CODIGO) & Space(10 - Len(Trim(BREC!SGI_CODIGO))) & " - " & Trim(BREC!SGI_DESCRICAO)
        Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
        BREC.MoveNext
     Loop
    
     ''If Combo.ListCount > 0 Then Combo.ListIndex = 0
    
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
     
     Dim I         As Integer
     Dim boolAchou As Boolean
     
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Then
        
Altera:
        
        ''' -------------------------------------------------
        ''' Grava Contas a Pagar Cabeçalho do titulo
        sSql = "Insert Into SGI_CONTASHAPG (SGI_FILIAL" & vbCrLf
        sSql = sSql & "                    ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                    ,SGI_DATALANC" & vbCrLf
        sSql = sSql & "                    ,SGI_CODFOR" & vbCrLf
        sSql = sSql & "                    ,SGI_CODCONDPGT" & vbCrLf
        sSql = sSql & "                    ,SGI_GRPDESP" & vbCrLf
        sSql = sSql & "                    ,SGI_VLTOTDOC" & vbCrLf
        sSql = sSql & "                    ,SGI_DOCPAI" & vbCrLf
        sSql = sSql & "                    ,SGI_QTDPARC)" & vbCrLf
        sSql = sSql & "             Values (" & intFilial & vbCrLf
        sSql = sSql & "                    ," & lngCODPGTO & vbCrLf
        sSql = sSql & "                    ,'" & Format(dtDATALCTO, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                    ," & lngCODFORN & vbCrLf
        sSql = sSql & "                    ," & lngCODCONDPGTO & vbCrLf
        sSql = sSql & "                    ," & lngCODGRPDESP & vbCrLf
        
        sValor = Replace(curVLTOTLCTO, ".", "")
        sValor = Replace(Trim(sValor), ",", ".")
        sSql = sSql & "                    ," & sValor & vbCrLf
        
        sSql = sSql & "                    ,'" & strDOCPAI & "'" & vbCrLf
        sSql = sSql & "                    ," & UBound(arrDOCPGTO) & ")"
         
        BGRV.CommandText = sSql
        BGRV.Execute
        
        ''' -------------------------------------------------
        ''' Grava Titulos para ser pagos
        If IsArray(arrDOCPGTO) = True Then
           For I = 1 To UBound(arrDOCPGTO)
            
               sSql = "Insert Into SGI_CONTASIAPG (SGI_FILIAL" & vbCrLf
               sSql = sSql & "                           ,SGI_CODIGO" & vbCrLf
               sSql = sSql & "                           ,SGI_NUMDOC" & vbCrLf
               sSql = sSql & "                           ,SGI_DATAVENC" & vbCrLf
               sSql = sSql & "                           ,SGI_PARCELA " & vbCrLf
               sSql = sSql & "                           ,SGI_VLDOC   " & vbCrLf
               sSql = sSql & "                           ,SGI_TIPDOC  " & vbCrLf
               sSql = sSql & "                           ,SGI_STATUS)" & vbCrLf
               sSql = sSql & "                    Values (" & intFilial & vbCrLf
               sSql = sSql & "                           ," & lngCODPGTO & vbCrLf
               sSql = sSql & "                           ,'" & arrDOCPGTO(I, 1) & "'" & vbCrLf
               sSql = sSql & "                           ,'" & Format(arrDOCPGTO(I, 2), "MM/DD/YYYY") & "'" & vbCrLf
               sSql = sSql & "                           ," & I & vbCrLf
               
               sValor = Replace(arrDOCPGTO(I, 3), ".", "")
               sValor = Replace(Trim(sValor), ",", ".")
               sSql = sSql & "                           ," & sValor & vbCrLf
               
               sSql = sSql & "                           ," & IIf(Len(Trim(arrDOCPGTO(I, 5))) = 0, "Null", arrDOCPGTO(I, 5)) & vbCrLf
               sSql = sSql & "                           ,'A')" & vbCrLf
               
               BGRV.CommandText = sSql
               BGRV.Execute
               
           Next I
        End If
        
        ''' -------------------------------------------------
        ''' Inseri o Grupo de Despesa
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       * " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "       SGI_GRPDESPFORN" & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL    = " & intFilial & vbCrLf
        sSql = sSql & "   And SGI_CODFORN   = " & lngCODFORN & vbCrLf
        sSql = sSql & "   And SGI_CODGRPDSP = " & lngCODGRPDESP
        
        boolAchou = False
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        If Not BREC.EOF Then boolAchou = True
        BREC.Close
        
        If boolAchou = False Then
           
           sSql = "Insert Into SGI_GRPDESPFORN (SGI_FILIAL,SGI_CODGRPDSP,SGI_CODFORN) " & vbCrLf
           sSql = sSql & "                     Values (" & vbCrLf
           sSql = sSql & "                             " & intFilial & vbCrLf
           sSql = sSql & "                            ," & lngCODGRPDESP & vbCrLf
           sSql = sSql & "                            ," & lngCODFORN & vbCrLf
           sSql = sSql & "                             )"
               
           BGRV.CommandText = sSql
           BGRV.Execute
           
        End If
        ''' -------------------------------------------------
        
     ElseIf strOPER = "A" Then
        
        '' Apaga Titulos Itens
        sSql = "Delete From SGI_CONTASIAPG " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODPGTO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Apaga Titulos Cabeçalho
        sSql = "Delete From SGI_CONTASHAPG " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODPGTO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        GoTo Altera
     
     ElseIf strOPER = "E" Then
     
        '' Apaga Titulos Itens
        sSql = "Delete From SGI_CONTASIAPG " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODPGTO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Apaga Titulos Cabeçalho
        sSql = "Delete From SGI_CONTASHAPG " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODPGTO
        
        BGRV.CommandText = sSql
        BGRV.Execute
     
     ElseIf strOPER = "B" Or strOPER = "AB" Then
        
        '' Baixando Titulo ou Alterando o Titulo
        sSql = "Update SGI_CONTASIAPG Set " & vbCrLf
        
        sValor = Replace(curVLPGTO, ".", "")
        sValor = Replace(Trim(sValor), ",", ".")
        sSql = sSql & "                      SGI_VLPAGO     = " & sValor & vbCrLf
        
        sSql = sSql & "                     ,SGI_DTPGTO     = '" & Format(dtDTPAGTO, "MM/DD/YYYY") & "'" & vbCrLf
        If lngTIPOPGTO = 0 Then
           sSql = sSql & "                     ,SGI_TIPOPGTO   = Null" & vbCrLf
        Else
           sSql = sSql & "                     ,SGI_TIPOPGTO   = " & lngTIPOPGTO & vbCrLf
        End If
        
        If curDESCONTOPGTO > 0 Then
           sValor = Replace(curDESCONTOPGTO, ".", "")
           sValor = Replace(Trim(sValor), ",", ".")
           sSql = sSql & "                     ,SGI_VLDESC     = " & sValor & vbCrLf
        Else
           sSql = sSql & "                     ,SGI_VLDESC     = NuLL" & vbCrLf
        End If
        
        If curPORCDESC > 0 Then
           sValor = Replace(curPORCDESC, ".", "")
           sValor = Replace(Trim(sValor), ",", ".")
           sSql = sSql & "                     ,SGI_PORCDESC   = " & sValor & vbCrLf
        Else
           sSql = sSql & "                     ,SGI_PORCDESC   = Null" & vbCrLf
        End If
        
        If curACRESCPGTO > 0 Then
           sValor = Replace(curACRESCPGTO, ".", "")
           sValor = Replace(Trim(sValor), ",", ".")
           sSql = sSql & "                     ,SGI_VLACRES    = " & sValor & vbCrLf
        Else
           sSql = sSql & "                     ,SGI_VLACRES    = Null" & vbCrLf
        End If
        
        If curPORCACRES > 0 Then
           sValor = Replace(curPORCACRES, ".", "")
           sValor = Replace(Trim(sValor), ",", ".")
           sSql = sSql & "                     ,SGI_PORCACRES  = " & sValor & vbCrLf
        Else
           sSql = sSql & "                     ,SGI_PORCACRES  = Null" & vbCrLf
        End If
        
        sSql = sSql & "                     ,SGI_STATUS  = 'B'" & vbCrLf
        sSql = sSql & "Where " & vbCrLf
        sSql = sSql & "      SGI_FILIAL  = " & FILIAL & vbCrLf
        sSql = sSql & "  And SGI_CODIGO  = " & lngCODPGTO & vbCrLf
        sSql = sSql & "  And SGI_NUMDOC  = '" & strNUMDOC & "'" & vbCrLf
        sSql = sSql & "  And SGI_PARCELA = " & intPARCPGTO
        
        BGRV.CommandText = sSql
        BGRV.Execute
     
     ElseIf strOPER = "X" Then
        
        '' Extorno
        sSql = "Update SGI_CONTASIAPG Set " & vbCrLf
        sSql = sSql & "                      SGI_VLPAGO     = Null" & vbCrLf
        sSql = sSql & "                     ,SGI_DTPGTO     = Null" & vbCrLf
        sSql = sSql & "                     ,SGI_TIPOPGTO   = Null" & vbCrLf
        sSql = sSql & "                     ,SGI_VLDESC     = NuLL" & vbCrLf
        sSql = sSql & "                     ,SGI_PORCDESC   = Null" & vbCrLf
        sSql = sSql & "                     ,SGI_VLACRES    = Null" & vbCrLf
        sSql = sSql & "                     ,SGI_PORCACRES  = Null" & vbCrLf
        sSql = sSql & "                     ,SGI_STATUS     = 'A'" & vbCrLf
        sSql = sSql & "Where " & vbCrLf
        sSql = sSql & "      SGI_FILIAL  = " & FILIAL & vbCrLf
        sSql = sSql & "  And SGI_CODIGO  = " & lngCODPGTO & vbCrLf
        sSql = sSql & "  And SGI_NUMDOC  = '" & strNUMDOC & "'" & vbCrLf
        sSql = sSql & "  And SGI_PARCELA = " & intPARCPGTO
        
        BGRV.CommandText = sSql
        BGRV.Execute
     
     End If
    
     adoBanco_Dados.CommitTrans
     GRAVA = True
     
     Exit Function
     
grav_err:
    
    MsgBox "Erro Nº : " & Err.Number & vbCrLf & " Descrição : " & Err.Description
    
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
   sSql = sSql & "       SGI_CONTASHAPG " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & lngCODPGTO & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC.EOF Then
      
      lngCODPGTO = BREC!SGI_CODIGO
      dtDATALCTO = BREC!SGI_DATALANC
      lngCODFORN = BREC!SGI_CODFOR
      lngCODCONDPGTO = BREC!SGI_CODCONDPGT
      lngCODGRPDESP = BREC!SGI_GRPDESP
      curVLTOTLCTO = BREC!SGI_VLTOTDOC
      strDOCPAI = BREC!SGI_DOCPAI
      
      Carrega_campos = True
   End If
   
   BREC.Close
   
   If Carrega_campos = False Then Exit Function
   
   '' ---------------
   
   BREC2.ActiveConnection = adoBanco_Dados
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CONTASIAPG " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial
   sSql = sSql & "   And SGI_CODIGO = " & lngCODPGTO & vbCrLf
   sSql = sSql & " Order by SGI_PARCELA"
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC2.EOF Then
   
      intReg = 0
      Do While Not BREC2.EOF
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
      
      ReDim arrDOCPGTO(1 To intReg, 1 To 5)
      
      intReg = 1
      BREC2.MoveFirst
      Do While Not BREC2.EOF
         arrDOCPGTO(intReg, 1) = BREC2!SGI_NUMDOC
         arrDOCPGTO(intReg, 2) = BREC2!SGI_DATAVENC
         arrDOCPGTO(intReg, 3) = BREC2!SGI_PARCELA
         arrDOCPGTO(intReg, 4) = BREC2!SGI_VLDOC
         arrDOCPGTO(intReg, 5) = BREC2!SGI_TIPDOC
         
         If intPARCPGTO = intReg Then
         
            If Not IsNull(BREC2!SGI_VLPAGO) Then curVLPGTO = BREC2!SGI_VLPAGO
            If Not IsNull(BREC2!SGI_DTPGTO) Then dtDTPAGTO = BREC2!SGI_DTPGTO
            If Not IsNull(BREC2!SGI_TIPOPGTO) Then lngTIPOPGTO = BREC2!SGI_TIPOPGTO
            If Not IsNull(BREC2!SGI_VLDESC) Then curDESCONTOPGTO = BREC2!SGI_VLDESC
            If Not IsNull(BREC2!SGI_PORCDESC) Then curPORCDESC = BREC2!SGI_PORCDESC
            If Not IsNull(BREC2!SGI_VLACRES) Then curACRESCPGTO = BREC2!SGI_VLACRES
            If Not IsNull(BREC2!SGI_PORCACRES) Then curPORCACRES = BREC2!SGI_PORCACRES
            strNUMDOC = BREC2!SGI_NUMDOC
         
         End If
         
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
   
   End If
   
   BREC2.Close

End Function

Public Sub PreencheCombo(Combo As Variant, strTabela As String)

     Combo.Clear
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       * " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       " & strTabela & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL  = " & intFilial & vbCrLf
     If UCase(Trim(strTabela)) = "SGI_CADTIPOPGTO" Then
        sSql = sSql & "   And SGI_SINAL  = '-'" & vbCrLf
     End If
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_CODIGO) & Space(10 - Len(Trim(BREC!SGI_CODIGO))) & " - " & Trim(BREC!SGI_DESCRICAO)
        Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
        BREC.MoveNext
     Loop
    
     ''If Combo.ListCount > 0 Then Combo.ListIndex = 0
    
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


Public Function GRAVALOTE(strOPER As String) As Boolean

On Error GoTo grav_err
     
     GRAVALOTE = False
     
     Dim I As Integer
     
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Then
        
Altera:
        
        ''' -------------------------------------------------
        ''' GravaCabeçalho do Lote
        sSql = "Insert Into SGI_CADLOTEHEADER (SGI_FILIAL   " & vbCrLf
        sSql = sSql & "                       ,SGI_CODIGO   " & vbCrLf
        sSql = sSql & "                       ,SGI_NUMDOC   " & vbCrLf
        sSql = sSql & "                       ,SGI_DATALOTE " & vbCrLf
        sSql = sSql & "                       ,SGI_TIPPGTO  " & vbCrLf
        sSql = sSql & "                       ,SGI_CODBCO   " & vbCrLf
        sSql = sSql & "                       ,SGI_VLDOC    " & vbCrLf
        sSql = sSql & "                       ,SGI_STATUS)  " & vbCrLf
        sSql = sSql & "             Values (" & intFilial & vbCrLf
        sSql = sSql & "                    ," & lngCODLOTE & vbCrLf
        sSql = sSql & "                    ,'" & strNUMDOC & "'" & vbCrLf
        sSql = sSql & "                    ,'" & Format(dtDATALOTE, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                    ," & lngTIPOPGTO & vbCrLf
        sSql = sSql & "                    ," & lngCODBANCO & vbCrLf
        
        sValor = Replace(curVLTOTLCTO, ".", "")
        sValor = Replace(Trim(sValor), ",", ".")
        sSql = sSql & "                    ," & sValor & vbCrLf
         
        sSql = sSql & "                    ,'A')"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        ''' -------------------------------------------------
        ''' Grava Lote de Titulos para serem pagos
        If IsArray(arrDOCPGTO) = True Then
           For I = 1 To UBound(arrDOCPGTO)
           
               '' Inclui Itens do Lote
               sSql = "Insert into SGI_CADLOTEITENS (" & vbCrLf
               sSql = sSql & "                               SGI_FILIAL  " & vbCrLf
               sSql = sSql & "                              ,SGI_CODIGO  " & vbCrLf
               sSql = sSql & "                              ,SGI_CODTIT  " & vbCrLf
               sSql = sSql & "                              ,SGI_PARCELA)" & vbCrLf
               sSql = sSql & "               Values (" & vbCrLf
               sSql = sSql & "                              " & intFilial & vbCrLf
               sSql = sSql & "                             ," & lngCODLOTE & vbCrLf
               sSql = sSql & "                             ," & arrDOCPGTO(I, 1) & vbCrLf
               sSql = sSql & "                             ," & arrDOCPGTO(I, 4) & ")"
               
               BGRV.CommandText = sSql
               BGRV.Execute
            
               '' Atualizabdo tabela de titulos
               sSql = "Update SGI_CONTASIAPG Set " & vbCrLf
               sSql = sSql & "                          SGI_NLOTE  = " & lngCODLOTE & vbCrLf
               sSql = sSql & "                         ,SGI_STATUS = 'L'" & vbCrLf
               sSql = sSql & "                    Where " & vbCrLf
               sSql = sSql & "                          SGI_FILIAL   = " & intFilial & vbCrLf
               sSql = sSql & "                      And SGI_CODIGO   = " & arrDOCPGTO(I, 1) & vbCrLf
               sSql = sSql & "                      And SGI_NUMDOC   = '" & arrDOCPGTO(I, 2) & "'" & vbCrLf
               sSql = sSql & "                      And SGI_DATAVENC = '" & Format(CDate(arrDOCPGTO(I, 3)), "MM/DD/YYYY") & "'" & vbCrLf
               sSql = sSql & "                      And SGI_PARCELA  = " & arrDOCPGTO(I, 4) & vbCrLf
               
               BGRV.CommandText = sSql
               BGRV.Execute
               
           Next I
        End If
     
     ElseIf strOPER = "A" Then
        
        '' Apaga Titulos Lote Itens
        sSql = "Delete From SGI_CADLOTEITENS " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODLOTE
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        
        '' Apaga Titulos Lote Cabeçalho
        sSql = "Delete From SGI_CADLOTEHEADER " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODLOTE
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Extorna Titulo
        sSql = "Update SGI_CONTASIAPG Set " & vbCrLf
        sSql = sSql & "                          SGI_NLOTE  = Null" & vbCrLf
        sSql = sSql & "                         ,SGI_STATUS = 'A' " & vbCrLf
        sSql = sSql & "                    Where " & vbCrLf
        sSql = sSql & "                          SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "                      And SGI_NLOTE  = " & lngCODLOTE

        BGRV.CommandText = sSql
        BGRV.Execute
        
        GoTo Altera
     
     ElseIf strOPER = "E" Then
     
        '' Apaga Titulos Itens
        sSql = "Delete From SGI_CONTASIAPG " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODPGTO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        
        '' Apaga Titulos Cabeçalho
        sSql = "Delete From SGI_CONTASHAPG " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODPGTO
        
        BGRV.CommandText = sSql
        BGRV.Execute
     
     ElseIf strOPER = "BL" Then
        
        
        sSql = "Update SGI_CADLOTEHEADER Set SGI_STATUS = 'B'" & vbCrLf
        sSql = sSql & "                   Where " & vbCrLf
        sSql = sSql & "                         SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "                     And SGI_CODIGO = " & lngCODLOTE
           
        BGRV.CommandText = sSql
        BGRV.Execute
        
        If BREC.State = 1 Then BREC.Close
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       IAPG.*   " & vbCrLf
        sSql = sSql & "      ,ITENS.*  " & vbCrLf
        sSql = sSql & "      ,HEADER.SGI_DATALOTE " & vbCrLf
        sSql = sSql & "      ,HEADER.SGI_TIPPGTO  " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "      SGI_CADLOTEITENS  ITENS " & vbCrLf
        sSql = sSql & "     ,SGI_CONTASIAPG    IAPG  " & vbCrLf
        sSql = sSql & "     ,SGI_CADLOTEHEADER HEADER" & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       ITENS.SGI_FILIAL  = " & FILIAL & vbCrLf
        sSql = sSql & "   And ITENS.SGI_CODIGO  = " & lngCODLOTE & vbCrLf
        sSql = sSql & "   And IAPG.SGI_FILIAL   = ITENS.SGI_FILIAL  " & vbCrLf
        sSql = sSql & "   And IAPG.SGI_CODIGO   = ITENS.SGI_CODTIT  " & vbCrLf
        sSql = sSql & "   And IAPG.SGI_PARCELA  = ITENS.SGI_PARCELA " & vbCrLf
        sSql = sSql & "   And HEADER.SGI_FILIAL = ITENS.SGI_FILIAL  " & vbCrLf
        sSql = sSql & "   And HEADER.SGI_CODIGO = ITENS.SGI_CODIGO  "
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        
        Do While Not BREC.EOF
        
            '' Baixando Titulo ou Alterando o Titulo
            sValor = Replace(BREC!SGI_VLDOC, ".", "")
            sValor = Replace(Trim(sValor), ",", ".")
            
            sSql = "Update SGI_CONTASIAPG Set " & vbCrLf
            sSql = sSql & "                      SGI_VLPAGO     = " & sValor & vbCrLf
            sSql = sSql & "                     ,SGI_DTPGTO     = '" & Format(BREC!SGI_DATALOTE, "MM/DD/YYYY") & "'" & vbCrLf
            sSql = sSql & "                     ,SGI_TIPOPGTO   = " & BREC!SGI_TIPPGTO & vbCrLf
            sSql = sSql & "                     ,SGI_VLDESC     = NuLL" & vbCrLf
            sSql = sSql & "                     ,SGI_PORCDESC   = Null" & vbCrLf
            sSql = sSql & "                     ,SGI_VLACRES    = Null" & vbCrLf
            sSql = sSql & "                     ,SGI_PORCACRES  = Null" & vbCrLf
            sSql = sSql & "                     ,SGI_STATUS  = 'B'" & vbCrLf
            sSql = sSql & "Where " & vbCrLf
            sSql = sSql & "      SGI_FILIAL  = " & FILIAL & vbCrLf
            sSql = sSql & "  And SGI_CODIGO  = " & BREC!SGI_CODTIT & vbCrLf
            sSql = sSql & "  And SGI_NUMDOC  = '" & BREC!SGI_NUMDOC & "'" & vbCrLf
            sSql = sSql & "  And SGI_PARCELA = " & BREC!SGI_PARCELA
        
            BGRV.CommandText = sSql
            BGRV.Execute
           
           BREC.MoveNext
        Loop
        BREC.Close
     
     ElseIf strOPER = "X" Then
        
        '' Apaga Itens
        sSql = "Delete From SGI_CADLOTEITENS " & vbCrLf
        sSql = sSql & "Where " & vbCrLf
        sSql = sSql & "      SGI_FILIAL  = " & FILIAL & vbCrLf
        sSql = sSql & "  And SGI_CODIGO  = " & lngCODLOTE & vbCrLf
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Extorno -- Apaga Cabeçalho
        sSql = "Delete From SGI_CADLOTEHEADER " & vbCrLf
        sSql = sSql & "Where " & vbCrLf
        sSql = sSql & "      SGI_FILIAL  = " & FILIAL & vbCrLf
        sSql = sSql & "  And SGI_CODIGO  = " & lngCODLOTE & vbCrLf
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Extorna Titulo
        sSql = "Update SGI_CONTASIAPG Set " & vbCrLf
        sSql = sSql & "                          SGI_NLOTE  = Null" & vbCrLf
        sSql = sSql & "                         ,SGI_STATUS = 'A' " & vbCrLf
        sSql = sSql & "                    Where " & vbCrLf
        sSql = sSql & "                          SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "                      And SGI_NLOTE  = " & lngCODLOTE

        BGRV.CommandText = sSql
        BGRV.Execute
     
     ElseIf strOPER = "L" Then
     
        sSql = "Update SGI_CADLOTEHEADER Set SGI_STATUS = 'L'" & vbCrLf
        sSql = sSql & "                   Where " & vbCrLf
        sSql = sSql & "                         SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "                     And SGI_CODIGO = " & lngCODLOTE
           
        BGRV.CommandText = sSql
        BGRV.Execute
     
     ElseIf strOPER = "D" Then
     
        sSql = "Update SGI_CADLOTEHEADER Set SGI_STATUS = 'A'" & vbCrLf
        sSql = sSql & "                   Where " & vbCrLf
        sSql = sSql & "                         SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "                     And SGI_CODIGO = " & lngCODLOTE
           
        BGRV.CommandText = sSql
        BGRV.Execute
     
     ElseIf strOPER = "EL" Then
     
        sSql = "Update SGI_CADLOTEHEADER Set SGI_STATUS = 'L'" & vbCrLf
        sSql = sSql & "                   Where " & vbCrLf
        sSql = sSql & "                         SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "                     And SGI_CODIGO = " & lngCODLOTE
           
        BGRV.CommandText = sSql
        BGRV.Execute
        
        
        sSql = "Select " & vbCrLf
        sSql = sSql & "       IAPG.*   " & vbCrLf
        sSql = sSql & "      ,ITENS.*  " & vbCrLf
        sSql = sSql & "      ,HEADER.SGI_DATALOTE " & vbCrLf
        sSql = sSql & "      ,HEADER.SGI_TIPPGTO  " & vbCrLf
        sSql = sSql & "  From " & vbCrLf
        sSql = sSql & "      SGI_CADLOTEITENS  ITENS " & vbCrLf
        sSql = sSql & "     ,SGI_CONTASIAPG    IAPG  " & vbCrLf
        sSql = sSql & "     ,SGI_CADLOTEHEADER HEADER" & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       ITENS.SGI_FILIAL  = " & FILIAL & vbCrLf
        sSql = sSql & "   And ITENS.SGI_CODIGO  = " & lngCODLOTE & vbCrLf
        sSql = sSql & "   And IAPG.SGI_FILIAL   = ITENS.SGI_FILIAL  " & vbCrLf
        sSql = sSql & "   And IAPG.SGI_CODIGO   = ITENS.SGI_CODTIT  " & vbCrLf
        sSql = sSql & "   And IAPG.SGI_PARCELA  = ITENS.SGI_PARCELA " & vbCrLf
        sSql = sSql & "   And HEADER.SGI_FILIAL = ITENS.SGI_FILIAL  " & vbCrLf
        sSql = sSql & "   And HEADER.SGI_CODIGO = ITENS.SGI_CODIGO  "
        
        BREC.Open sSql, adoBanco_Dados, adOpenDynamic
        
        Do While Not BREC.EOF
        
            '' Baixando Titulo ou Alterando o Titulo
            sSql = "Update SGI_CONTASIAPG Set " & vbCrLf
            sSql = sSql & "                      SGI_VLPAGO     = NULL" & vbCrLf
            sSql = sSql & "                     ,SGI_DTPGTO     = NULL" & vbCrLf
            sSql = sSql & "                     ,SGI_TIPOPGTO   = NULL" & vbCrLf
            sSql = sSql & "                     ,SGI_VLDESC     = NuLL" & vbCrLf
            sSql = sSql & "                     ,SGI_PORCDESC   = Null" & vbCrLf
            sSql = sSql & "                     ,SGI_VLACRES    = Null" & vbCrLf
            sSql = sSql & "                     ,SGI_PORCACRES  = Null" & vbCrLf
            sSql = sSql & "                     ,SGI_STATUS  = 'L'" & vbCrLf
            sSql = sSql & "Where " & vbCrLf
            sSql = sSql & "      SGI_FILIAL  = " & FILIAL & vbCrLf
            sSql = sSql & "  And SGI_CODIGO  = " & BREC!SGI_CODTIT & vbCrLf
            sSql = sSql & "  And SGI_NUMDOC  = '" & BREC!SGI_NUMDOC & "'" & vbCrLf
            sSql = sSql & "  And SGI_PARCELA = " & BREC!SGI_PARCELA
        
            BGRV.CommandText = sSql
            BGRV.Execute
           
            BREC.MoveNext
        Loop
        BREC.Close
     
     
     End If
    
     adoBanco_Dados.CommitTrans
     GRAVALOTE = True
     
     Exit Function
     
grav_err:
    MsgBox "Erro Nº : " & Err.Number & vbCrLf & " Descrição : " & Err.Description
    adoBanco_Dados.RollbackTrans
    
End Function

Public Function Carrega_CamposLote() As Boolean

   Carrega_CamposLote = False
   
   Dim intReg As Integer
   Dim I      As Integer
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADLOTEHEADER " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & lngCODLOTE & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC.EOF Then
      
      lngCODLOTE = BREC!SGI_CODIGO
      dtDATALOTE = BREC!SGI_DATALOTE
      strNUMDOC = BREC!SGI_NUMDOC
      curVLTOTLCTO = BREC!SGI_VLDOC
      lngTIPOPGTO = BREC!SGI_TIPPGTO
      lngCODBANCO = BREC!SGI_CODBCO
      strSTATUS = BREC!SGI_STATUS
      
      Carrega_CamposLote = True
   End If
   
   BREC.Close
   
   If Carrega_CamposLote = False Then Exit Function
   
   '' ---------------
   
   BREC2.ActiveConnection = adoBanco_Dados
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADLOTEITENS " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial
   sSql = sSql & "   And SGI_CODIGO = " & lngCODLOTE & vbCrLf
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC2.EOF Then
   
      intReg = 0
      Do While Not BREC2.EOF
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
      
      ReDim arrDOCPGTO(1 To intReg, 1 To 5)
      
      intReg = 1
      BREC2.MoveFirst
      Do While Not BREC2.EOF
         arrDOCPGTO(intReg, 1) = BREC2!SGI_CODTIT
         arrDOCPGTO(intReg, 2) = BREC2!SGI_PARCELA
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

