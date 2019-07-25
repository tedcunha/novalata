VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADCONTRECEB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim intFilial       As Integer
Dim lngCODPGTO      As Long
Dim dtDATALCTO      As Date
Dim lngCODFORN      As Long
Dim lngCODCONDPGTO  As Long
Dim lngCODTIPDOC    As Long
Dim lngCODBANCO     As Long
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
Dim intGRPRECEB     As Integer

Public Property Get GRPRECEB() As Integer
    GRPRECEB = intGRPRECEB
End Property

Public Property Let GRPRECEB(ByVal vData As Integer)
    intGRPRECEB = vData
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

Public Property Get CODTIPDOC() As Long
    CODTIPDOC = lngCODTIPDOC
End Property

Public Property Let CODTIPDOC(ByVal vData As Long)
    lngCODTIPDOC = vData
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

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADCONTRECEBP.cCaminho = cCaminho
    frmCADCONTRECEBP.Linha = Linha
    frmCADCONTRECEBP.FILIAL = iFilial
    frmCADCONTRECEBP.strAcesso = strAcesso
    frmCADCONTRECEBP.strUSUARIO = strUSUARIO
    frmCADCONTRECEBP.Show vbModal
End Sub

Public Function Pesq_CadContasARC() As Boolean

   Pesq_CadContasARC = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       * " & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_CONTASHARC " & vbCrLf
   sSql = sSql + " Where " & vbCrLf
   sSql = sSql + "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_CadContasARC = True
   BREC.Close

End Function

Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo grav_err
     
     GRAVA = False
     
     Dim I As Integer
     
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Then
        
Altera:
        
        ''' -------------------------------------------------
        ''' Grava Contas a Pagar Cabeçalho do titulo
        sSql = "Insert Into SGI_CONTASHARC (SGI_FILIAL" & vbCrLf
        sSql = sSql & "                    ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                    ,SGI_DATALANC" & vbCrLf
        sSql = sSql & "                    ,SGI_CODCLI" & vbCrLf
        sSql = sSql & "                    ,SGI_CODCONDPGT" & vbCrLf
        sSql = sSql & "                    ,SGI_TIPDOC" & vbCrLf
        sSql = sSql & "                    ,SGI_CODBCO" & vbCrLf
        sSql = sSql & "                    ,SGI_VLTOTDOC" & vbCrLf
        sSql = sSql & "                    ,SGI_DOCPAI" & vbCrLf
        sSql = sSql & "                    ,SGI_QTDPARC" & vbCrLf
        sSql = sSql & "                    ,SGI_CODGRPRECEB)" & vbCrLf
        sSql = sSql & "             Values (" & intFilial & vbCrLf
        sSql = sSql & "                    ," & lngCODPGTO & vbCrLf
        sSql = sSql & "                    ,'" & Format(dtDATALCTO, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                    ," & lngCODFORN & vbCrLf
        sSql = sSql & "                    ," & lngCODCONDPGTO & vbCrLf
        sSql = sSql & "                    ," & lngCODTIPDOC & vbCrLf
        sSql = sSql & "                    ," & lngCODBANCO & vbCrLf
        
        sValor = Replace(curVLTOTLCTO, ".", "")
        sValor = Replace(Trim(sValor), ",", ".")
        sSql = sSql & "                    ," & sValor & vbCrLf
        
        sSql = sSql & "                    ,'" & strDOCPAI & "'" & vbCrLf
        sSql = sSql & "                    ," & UBound(arrDOCPGTO) & vbCrLf
        sSql = sSql & "                    ," & intGRPRECEB & ")"
         
        BGRV.CommandText = sSql
        BGRV.Execute
        
        ''' -------------------------------------------------
        ''' Grava Titulos para ser pagos
        If IsArray(arrDOCPGTO) = True Then
           For I = 1 To UBound(arrDOCPGTO)
            
               sSql = "Insert Into SGI_CONTASIARC (SGI_FILIAL" & vbCrLf
               sSql = sSql & "                           ,SGI_CODIGO" & vbCrLf
               sSql = sSql & "                           ,SGI_NUMDOC" & vbCrLf
               sSql = sSql & "                           ,SGI_DATAVENC" & vbCrLf
               sSql = sSql & "                           ,SGI_PARCELA" & vbCrLf
               sSql = sSql & "                           ,SGI_VLDOC " & vbCrLf
               sSql = sSql & "                           ,SGI_TIPDOC)" & vbCrLf
               sSql = sSql & "                    Values (" & intFilial & vbCrLf
               sSql = sSql & "                           ," & lngCODPGTO & vbCrLf
               sSql = sSql & "                           ,'" & arrDOCPGTO(I, 1) & "'" & vbCrLf
               sSql = sSql & "                           ,'" & Format(arrDOCPGTO(I, 2), "MM/DD/YYYY") & "'" & vbCrLf
               sSql = sSql & "                           ," & I & vbCrLf
               
               sValor = Replace(arrDOCPGTO(I, 3), ".", "")
               sValor = Replace(Trim(sValor), ",", ".")
               sSql = sSql & "                           ," & sValor & vbCrLf
               
               sSql = sSql & "                           ," & lngCODTIPDOC & ")" & vbCrLf
               
               BGRV.CommandText = sSql
               BGRV.Execute
               
           Next I
        End If
     
     ElseIf strOPER = "A" Then
        
        '' Apaga Titulos Itens
        sSql = "Delete From SGI_CONTASIARC " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODPGTO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        
        '' Apaga Titulos Cabeçalho
        sSql = "Delete From SGI_CONTASHARC " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODPGTO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        GoTo Altera
     
     ElseIf strOPER = "E" Then
     
        '' Apaga Titulos Itens
        sSql = "Delete From SGI_CONTASIARC " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODPGTO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        
        '' Apaga Titulos Cabeçalho
        sSql = "Delete From SGI_CONTASHARC " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & lngCODPGTO
        
        BGRV.CommandText = sSql
        BGRV.Execute
     
     ElseIf strOPER = "B" Or strOPER = "AB" Then
        
        '' Baixando Titulo ou Alterando o Titulo
        sSql = "Update SGI_CONTASIARC Set " & vbCrLf
        
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
        
        sSql = sSql & "Where " & vbCrLf
        sSql = sSql & "      SGI_FILIAL  = " & FILIAL & vbCrLf
        sSql = sSql & "  And SGI_CODIGO  = " & lngCODPGTO & vbCrLf
        sSql = sSql & "  And SGI_NUMDOC  = '" & strNUMDOC & "'" & vbCrLf
        sSql = sSql & "  And SGI_PARCELA = " & intPARCPGTO
        
        BGRV.CommandText = sSql
        BGRV.Execute
     
     ElseIf strOPER = "X" Then
        
        '' Extorno
        sSql = "Update SGI_CONTASIARC Set " & vbCrLf
        sSql = sSql & "                      SGI_VLPAGO     = Null" & vbCrLf
        sSql = sSql & "                     ,SGI_DTPGTO     = Null" & vbCrLf
        sSql = sSql & "                     ,SGI_TIPOPGTO   = Null" & vbCrLf
        sSql = sSql & "                     ,SGI_VLDESC     = NuLL" & vbCrLf
        sSql = sSql & "                     ,SGI_PORCDESC   = Null" & vbCrLf
        sSql = sSql & "                     ,SGI_VLACRES    = Null" & vbCrLf
        sSql = sSql & "                     ,SGI_PORCACRES  = Null" & vbCrLf
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
   sSql = sSql & "       SGI_CONTASHARC " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & lngCODPGTO & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC.EOF Then
      
      lngCODPGTO = BREC!SGI_CODIGO
      dtDATALCTO = BREC!SGI_DATALANC
      lngCODFORN = BREC!SGI_CODCLI
      lngCODCONDPGTO = BREC!SGI_CODCONDPGT
      lngCODTIPDOC = BREC!SGI_TIPDOC
      lngCODBANCO = BREC!SGI_CODBCO
      curVLTOTLCTO = BREC!SGI_VLTOTDOC
      strDOCPAI = BREC!SGI_DOCPAI
      If Not IsNull(BREC!SGI_CODGRPRECEB) Then intGRPRECEB = BREC!SGI_CODGRPRECEB
      
      Carrega_campos = True
   End If
   
   BREC.Close
   
   If Carrega_campos = False Then Exit Function
   
   '' ---------------
   
   BREC2.ActiveConnection = adoBanco_Dados
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CONTASIARC " & vbCrLf
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
