VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADPLAMESTRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial       As Integer
Dim lngCODIGO       As Long
Dim lngCODLINHA     As Long
Dim intMES          As Integer
Dim lngANO          As Long
Dim intOPTSIMNAO    As Integer
Dim curQTDE         As Currency
Dim curQTDPEDIDOS   As Currency
Dim lngQTDBATIDAS   As Long
Dim arrITENSDIAS    As Variant
Dim intATIVO        As Integer
Dim arrITENSDIASSEM As Variant

Public Property Get ITENSDIASSEM() As Variant
    ITENSDIASSEM = arrITENSDIASSEM
End Property
Public Property Let ITENSDIASSEM(ByVal vData As Variant)
    arrITENSDIASSEM = vData
End Property

Public Property Get ATIVO() As Integer
    ATIVO = intATIVO
End Property
Public Property Let ATIVO(ByVal vData As Integer)
    intATIVO = vData
End Property

Public Property Get ITENSDIAS() As Variant
    ITENSDIAS = arrITENSDIAS
End Property
Public Property Let ITENSDIAS(ByVal vData As Variant)
    arrITENSDIAS = vData
End Property

Public Property Get QTDBATIDAS() As Long
    QTDBATIDAS = lngQTDBATIDAS
End Property
Public Property Let QTDBATIDAS(ByVal vData As Long)
    lngQTDBATIDAS = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property
Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Property Get CODIGO() As Long
    CODIGO = lngCODIGO
End Property
Public Property Let CODIGO(ByVal vData As Long)
    lngCODIGO = vData
End Property
Public Property Get CODLINHA() As Long
    CODLINHA = lngCODLINHA
End Property
Public Property Let CODLINHA(ByVal vData As Long)
    lngCODLINHA = vData
End Property
Public Property Get MES() As Integer
    MES = intMES
End Property
Public Property Let MES(ByVal vData As Integer)
    intMES = vData
End Property
Public Property Get ANO() As Long
    ANO = lngANO
End Property
Public Property Let ANO(ByVal vData As Long)
    lngANO = vData
End Property
Public Property Get optSimNao() As Integer
    optSimNao = intOPTSIMNAO
End Property
Public Property Let optSimNao(ByVal vData As Integer)
    intOPTSIMNAO = vData
End Property

Public Property Get QTDE() As Currency
    QTDE = curQTDE
End Property
Public Property Let QTDE(ByVal vData As Currency)
    curQTDE = vData
End Property

Public Property Get QTDPEDIDOS() As Currency
    QTDPEDIDOS = curQTDPEDIDOS
End Property
Public Property Let QTDPEDIDOS(ByVal vData As Currency)
    curQTDPEDIDOS = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long, Optional intFILIAL_PED As Integer)
    frmCADPLAMESTREP.cCaminho = cCaminho
    frmCADPLAMESTREP.Linha = Linha
    frmCADPLAMESTREP.FILIAL = iFilial
    frmCADPLAMESTREP.strAcesso = strAcesso
    frmCADPLAMESTREP.lngCodUsuario = lngIDUsuario
    frmCADPLAMESTREP.intFILIALTAB = intFILIAL_PED
    frmCADPLAMESTREP.Show vbModal
End Sub

Public Function Pesq_CadPlanMestre(strNOMEFILIAL As String) As Boolean

   Pesq_CadPlanMestre = False
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       SGI_CODIGO  " & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_CADPLANMESTRE" & strNOMEFILIAL & vbCrLf
   sSql = sSql + " Where " & vbCrLf
   sSql = sSql + "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_CadPlanMestre = True
   BREC.Close

End Function


Public Sub PreenchComboMes(Combo As Variant)

    Combo.Clear
    
    Combo.AddItem "Janeiro"
    Combo.ItemData(Combo.NewIndex) = 1
    Combo.AddItem "Fevereiro"
    Combo.ItemData(Combo.NewIndex) = 2
    Combo.AddItem "Março"
    Combo.ItemData(Combo.NewIndex) = 3
    Combo.AddItem "Abril"
    Combo.ItemData(Combo.NewIndex) = 4
    Combo.AddItem "Maio"
    Combo.ItemData(Combo.NewIndex) = 5
    Combo.AddItem "Junho"
    Combo.ItemData(Combo.NewIndex) = 6
    Combo.AddItem "Julho"
    Combo.ItemData(Combo.NewIndex) = 7
    Combo.AddItem "Agosto"
    Combo.ItemData(Combo.NewIndex) = 8
    Combo.AddItem "Setembro"
    Combo.ItemData(Combo.NewIndex) = 9
    Combo.AddItem "Outubro"
    Combo.ItemData(Combo.NewIndex) = 10
    Combo.AddItem "Novembro"
    Combo.ItemData(Combo.NewIndex) = 11
    Combo.AddItem "Dezembro"
    Combo.ItemData(Combo.NewIndex) = 12
    
End Sub


Public Sub PreenchComboAno(Combo As Variant)

    Combo.Clear
    
    Dim I As Long
    
    For I = 2000 To 3000
        Combo.AddItem Trim(Str(I))
        Combo.ItemData(Combo.NewIndex) = I
    Next I
    
End Sub


Public Sub PreencheComboProd(Combo As Variant)

     Combo.Clear
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       * " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL  = " & intFilial & vbCrLf
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_IDPRODUTO) & " - " & Trim(BREC!SGI_DESCRICAO)
        BREC.MoveNext
     Loop
    
     BREC.Close
    
End Sub

Public Function PegaPedidos(intMESINF As Integer, lngANOINF As Long) As Currency

    PegaPedidos = 0
    
    Dim OpEmAberto As Currency
    
    
    OpEmAberto = 0
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       Sum(ORDP.SGI_QTDE) as SGI_QTDE " & vbCrLf
    
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_ORDEMPROD  ORDP" & vbCrLf
    sSql = sSql & "      ,SGI_CADPRODUTO PROD" & vbCrLf
    
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       ORDP.SGI_FILIAL            = " & intFilial & vbCrLf
    sSql = sSql & "   And Month(ORDP.SGI_DATENTREGA) = " & intMESINF & vbCrLf
    sSql = sSql & "   And Year(ORDP.SGI_DATENTREGA)  = " & lngANOINF & vbCrLf
    sSql = sSql & "   And ORDP.SGI_STATUS            = 0" & vbCrLf
    sSql = sSql & "   And PROD.SGI_FILIAL            = ORDP.SGI_FILIAL " & vbCrLf
    sSql = sSql & "   And PROD.SGI_IDPRODUTO         = ORDP.SGI_IDPRODUTO " & vbCrLf
    sSql = sSql & "   And PROD.SGI_CODLINPROD        = " & lngCODLINHA & vbCrLf
    
    BREC.Open sSql, adoBanco_Dados
    If Not BREC.EOF Then
       If Not IsNull(BREC!SGI_QTDE) Then OpEmAberto = BREC!SGI_QTDE
    End If
    BREC.Close
    
    PegaPedidos = OpEmAberto
    
End Function

Public Function GRAVA(strOPER As String, strNOMETABELA As String) As Boolean

On Error GoTo err_grava
     
     GRAVA = False
     
     Dim I      As Integer
     Dim sValor As String
         
     BGRV.ActiveConnection = adoBanco_Dados
     adoBanco_Dados.BeginTrans
     
     If strOPER = "I" Then
       
RETORNO:
        
        '' Grava turnos
        sSql = "Insert Into SGI_CADPLANMESTRE" & strNOMETABELA & " (" & vbCrLf
        sSql = sSql & "                        SGI_FILIAL" & vbCrLf
        sSql = sSql & "                       ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                       ,SGI_CODLINHA" & vbCrLf
        sSql = sSql & "                       ,SGI_MES" & vbCrLf
        sSql = sSql & "                       ,SGI_ANO" & vbCrLf
        sSql = sSql & "                       ,SGI_CONSPED" & vbCrLf
        sSql = sSql & "                       ,SGI_QTDE" & vbCrLf
        sSql = sSql & "                       ,SGI_ATIVO" & vbCrLf
        sSql = sSql & "              ) Values (" & vbCrLf
        sSql = sSql & "                            " & intFilial & vbCrLf
        sSql = sSql & "                           ," & lngCODIGO & vbCrLf
        sSql = sSql & "                           ," & lngCODLINHA & vbCrLf
        sSql = sSql & "                           ," & intMES & vbCrLf
        sSql = sSql & "                           ," & lngANO & vbCrLf
        sSql = sSql & "                           ," & intOPTSIMNAO & vbCrLf
        
        sValor = Replace(curQTDE, ".", "")
        sValor = Replace(Trim(sValor), ",", ".")
        
        sSql = sSql & "                           ," & sValor & vbCrLf
        
        sSql = sSql & "                           ," & intATIVO & vbCrLf
        
        
        sSql = sSql & "                           )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' ======================
        '' Semana
        If IsArray(arrITENSDIAS) Then
            For I = 1 To UBound(arrITENSDIAS)
                
                sSql = ""
                
                sSql = "Insert Into SGI_CADDIASPM" & strNOMETABELA & " (" & vbCrLf
                sSql = sSql & "                           SGI_FILIAL" & vbCrLf
                sSql = sSql & "                          ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                          ,SGI_SEMANA" & vbCrLf
                sSql = sSql & "                          ,SGI_QTDE" & vbCrLf
                sSql = sSql & "                          ,SGI_IDINTERNO" & vbCrLf
                sSql = sSql & "                          ,SGI_ATIVO" & vbCrLf
                
                sSql = sSql & "                 ) Values (" & vbCrLf
                sSql = sSql & "                          " & intFilial & vbCrLf
                sSql = sSql & "                         ," & lngCODIGO & vbCrLf
                sSql = sSql & "                         ," & arrITENSDIAS(I, 1) & vbCrLf
                sSql = sSql & "                         ," & arrITENSDIAS(I, 2) & vbCrLf
                sSql = sSql & "                         ," & arrITENSDIAS(I, 3) & vbCrLf
                sSql = sSql & "                         ," & arrITENSDIAS(I, 4) & vbCrLf
                
                sSql = sSql & "                         )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
                
            Next I
        End If
        
        '' ======================
        '' Dias das Semanas
        If IsArray(arrITENSDIASSEM) Then
            For I = 1 To UBound(arrITENSDIASSEM)
                
                sSql = ""
                
                sSql = "Insert Into SGI_CADDIASPMSEMANA" & strNOMETABELA & " (" & vbCrLf
                sSql = sSql & "                           SGI_FILIAL" & vbCrLf
                sSql = sSql & "                          ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                          ,SGI_SEMANA" & vbCrLf
                sSql = sSql & "                          ,SGI_DTSEMANA" & vbCrLf
                sSql = sSql & "                          ,SGI_DIASEMANA" & vbCrLf
                sSql = sSql & "                          ,SGI_QTDE" & vbCrLf
                sSql = sSql & "                          ,SGI_IDINTERNO" & vbCrLf
                sSql = sSql & "                          ,SGI_ATIVO" & vbCrLf
                
                sSql = sSql & "                 ) Values (" & vbCrLf
                sSql = sSql & "                          " & intFilial & vbCrLf
                sSql = sSql & "                         ," & lngCODIGO & vbCrLf
                sSql = sSql & "                         ," & arrITENSDIASSEM(I, 1) & vbCrLf
                sSql = sSql & "                         ," & arrITENSDIASSEM(I, 2) & vbCrLf
                sSql = sSql & "                         ," & arrITENSDIASSEM(I, 3) & vbCrLf
                sSql = sSql & "                         ," & arrITENSDIASSEM(I, 4) & vbCrLf
                sSql = sSql & "                         ," & arrITENSDIASSEM(I, 5) & vbCrLf
                sSql = sSql & "                         ," & arrITENSDIASSEM(I, 6) & vbCrLf
                
                sSql = sSql & "                         )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
                
            Next I
        End If
        
        If intATIVO = 1 Then
            sSql = ""
            
            sSql = "Update SGI_CADPLANMESTRE Set SGI_ATIVO = 0" & vbCrLf
            sSql = sSql & "Where " & vbCrLf
            sSql = sSql & "      SGI_FILIAL   = " & intFilial & vbCrLf
            sSql = sSql & "  And SGI_CODIGO   <>" & lngCODIGO & vbCrLf
            sSql = sSql & "  And SGI_CODLINHA = " & lngCODLINHA & vbCrLf
            sSql = sSql & "  And SGI_MES      = " & intMES & vbCrLf
            sSql = sSql & "  And SGI_ANO      = " & lngANO
        
            BGRV.CommandText = sSql
            BGRV.Execute
        End If
        
     ElseIf strOPER = "A" Or strOPER = "E" Then
     
        '' Exclui Dias do Plano Mestre
        sSql = "Delete From SGI_CADDIASPMSEMANA" & strNOMETABELA & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO    = " & lngCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL    = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        sSql = "Delete From SGI_CADDIASPM" & strNOMETABELA & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO    = " & lngCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL    = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Exclui Plano Mestre
        sSql = "Delete from SGI_CADPLANMESTRE" & strNOMETABELA & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & lngCODIGO & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        If strOPER = "A" Then GoTo RETORNO
        
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
    sSql = sSql & "       (Max(SGI_NUMERO) + 1) As SGI_NUMERO " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_NUMERO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_MODULO = '" & sModulo & "'"
    sSql = sSql & "   And SGI_FILIAL = " & intFilial
    
    BREC.Open sSql, adoBanco_Dados
    
    If Not BREC.EOF Then
    
       If IsNull(BREC!SGI_NUMERO) = True Then
          
          Gera_Codigo = 1
          
          sSql = "Insert into SGI_NUMERO (SGI_FILIAL,SGI_NUMERO,SGI_MODULO) Values(" & vbCrLf
          sSql = sSql & "                                              " & intFilial & vbCrLf
          sSql = sSql & "                                            ,1" & vbCrLf
          sSql = sSql & "                                            ,'" & sModulo & "'" & vbCrLf
          sSql = sSql & "                                          )" & vbCrLf
          
       ElseIf BREC!SGI_NUMERO > 1 Then
       
          Gera_Codigo = BREC!SGI_NUMERO
          
          sSql = "Update SGI_NUMERO Set " & vbCrLf
          sSql = sSql & "           SGI_NUMERO = " & BREC!SGI_NUMERO & vbCrLf
          sSql = sSql & "         Where " & vbCrLf
          sSql = sSql & "               SGI_MODULO = '" & sModulo & "'" & vbCrLf
          sSql = sSql & "           And SGI_FILIAL =  " & intFilial
       
       End If
       
       BGRV.CommandText = sSql
       BGRV.Execute
       
       
    End If
    
    BREC.Close
    
End Function


Public Function CapacidadeProdMaq() As Currency

    CapacidadeProdMaq = 0
    
    ''sSql = "Select " & vbCrLf
    ''sSql = sSql & "       SUM(MAQN.SGI_CAPMINU) As SGI_CAPMINU " & vbCrLf
    ''sSql = sSql & "  From " & vbCrLf
    ''sSql = sSql & "       SGI_CADMAQPROD PROD " & vbCrLf
    ''sSql = sSql & "      ,SGI_CADMAQUINA MAQN " & vbCrLf
    ''sSql = sSql & " Where " & vbCrLf
    ''sSql = sSql & "       PROD.SGI_FILIAL  = " & FILIAL & vbCrLf
    ''sSql = sSql & "   And PROD.SGI_CODPROD = '" & strPRODUTO & "'" & vbCrLf
    ''sSql = sSql & "   And MAQN.SGI_FILIAL  = PROD.SGI_FILIAL " & vbCrLf
    ''sSql = sSql & "   And MAQN.SGI_CODIGO  = PROD.SGI_CODIGO " & vbCrLf
    ''sSql = sSql & "   And MAQN.SGI_ATIVA   = 0 "
    
    ''BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    ''If Not BREC.EOF Then CapacidadeProdMaq = BREC!SGI_CAPMINU
    ''BREC.Close

End Function

Public Function Carrega_campos(strNOMEFILIAL As String) As Boolean

   Carrega_campos = False
   
   Dim nreg     As Integer
   Dim intNREQ  As Integer
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADPLANMESTRE" & strNOMEFILIAL & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & lngCODIGO & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados
   
   If Not BREC.EOF Then
      
      lngCODIGO = BREC!SGI_CODIGO
      lngCODLINHA = BREC!SGI_CODLINHA
      intMES = BREC!SGI_MES
      lngANO = BREC!SGI_ANO
      intOPTSIMNAO = BREC!SGI_CONSPED
      curQTDE = BREC!SGI_QTDE
      intATIVO = BREC!SGI_ATIVO
       
      Carrega_campos = True
   End If
   BREC.Close
   
   
   '' Carrega Semanas --
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADDIASPM" & strNOMEFILIAL & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL  = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_CODIGO  = " & lngCODIGO
   
   BREC2.Open sSql, adoBanco_Dados
   
   intNREQ = 0
   Do While Not BREC2.EOF
      intNREQ = intNREQ + 1
      BREC2.MoveNext
   Loop
   
   If intNREQ > 0 Then
         
      BREC2.MoveFirst
      ReDim arrITENSDIAS(1 To intNREQ, 1 To 4) As String
      intNREQ = 1
      Do While Not BREC2.EOF
         arrITENSDIAS(intNREQ, 1) = BREC2!SGI_SEMANA
         If Not IsNull(BREC2!SGI_QTDE) Then arrITENSDIAS(intNREQ, 2) = BREC2!SGI_QTDE
         arrITENSDIAS(intNREQ, 3) = BREC2!SGI_IDINTERNO
         arrITENSDIAS(intNREQ, 4) = BREC2!SGI_ATIVO
         
         intNREQ = intNREQ + 1
         BREC2.MoveNext
      Loop
      
   End If
   
   BREC2.Close
   
   
   
   '' Carrega Dias Semana --
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADDIASPMSEMANA" & strNOMEFILIAL & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL  = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_CODIGO  = " & lngCODIGO
   
   BREC2.Open sSql, adoBanco_Dados
   
   intNREQ = 0
   Do While Not BREC2.EOF
      intNREQ = intNREQ + 1
      BREC2.MoveNext
   Loop
   
   If intNREQ > 0 Then
      BREC2.MoveFirst
      ReDim arrITENSDIASSEM(1 To intNREQ, 1 To 6) As String
      intNREQ = 1
      Do While Not BREC2.EOF
         arrITENSDIASSEM(intNREQ, 1) = BREC2!SGI_SEMANA
         arrITENSDIASSEM(intNREQ, 2) = BREC2!SGI_DTSEMANA
         arrITENSDIASSEM(intNREQ, 3) = BREC2!SGI_DIASEMANA
         If Not IsNull(BREC2!SGI_QTDE) Then arrITENSDIASSEM(intNREQ, 4) = BREC2!SGI_QTDE
         arrITENSDIASSEM(intNREQ, 5) = BREC2!SGI_IDINTERNO
         arrITENSDIASSEM(intNREQ, 6) = BREC2!SGI_ATIVO
         intNREQ = intNREQ + 1
         BREC2.MoveNext
      Loop
      
   End If
   
   BREC2.Close
   
End Function

Public Function VerifSaldoOF() As Currency

    VerifSaldoOF = 0
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       Sum(SGI_QTDE) as SGI_QTDE " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADOF " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_CODPMSTR = " & lngCODIGO
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not IsNull(BREC!SGI_QTDE) Then VerifSaldoOF = BREC!SGI_QTDE
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

Public Function PreenchComboAtivo() As String
    PreenchComboAtivo = PreenchComboAtivo & "|#" & 1 & ";SIM"
    PreenchComboAtivo = PreenchComboAtivo & "|#" & 0 & ";NÃO"
End Function

