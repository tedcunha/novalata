VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADFORNEC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial      As Integer
Dim intCodigoFOR   As Integer
Dim strCNPJCPFFOR  As String
Dim strRAZSOCFOR   As String
Dim strNOMFANTFOR  As String
Dim strENDFOR      As String
Dim strCIDFOR      As String
Dim strBAIFOR      As String
Dim intESTFOR      As Integer
Dim strCEPFOR      As String
Dim dtDTCADFOR     As Date
Dim strENDRETI     As String
Dim strBAIRETI     As String
Dim strCIDRETI     As String
Dim intESTRETI     As Integer
Dim strCEPRETI     As String
Dim arrBANCOS      As Variant
Dim arrTELFOR      As Variant
Dim arrCONTATOFOR  As Variant
Dim arrEMAILFOR    As Variant
Dim arrSITEFOR     As Variant
Dim arrPRODUTO     As Variant
Dim arrNATOPER     As Variant
Dim strModulo      As String
Dim strCONTROLE    As String
Dim lngIQF         As Long
Dim arrTRANSP      As Variant
Dim intAVISAIGF    As Integer
Dim intLIBIGF      As Integer
Dim intTRAVAIGF    As Integer
Dim intGRPDESP     As Integer
Dim intGRPDESPBKP  As Integer
Dim lngCODNIVRISCO As Long
Dim intSTATUS      As Integer
Dim arrPRODNAOINSP As Variant
Dim intREQINSPQUA  As Integer

Public Property Get REQINSPQUA() As Integer
    REQINSPQUA = intREQINSPQUA
End Property

Public Property Let REQINSPQUA(ByVal vData As Integer)
    intREQINSPQUA = vData
End Property

Public Property Get PRODNAOINSP() As Variant
    PRODNAOINSP = arrPRODNAOINSP
End Property

Public Property Let PRODNAOINSP(ByVal vData As Variant)
    arrPRODNAOINSP = vData
End Property

Public Property Get STATUS() As Integer
    STATUS = intSTATUS
End Property

Public Property Let STATUS(ByVal vData As Integer)
    intSTATUS = vData
End Property

Public Property Get CODNIVRISCO() As Long
    CODNIVRISCO = lngCODNIVRISCO
End Property

Public Property Let CODNIVRISCO(ByVal vData As Long)
    lngCODNIVRISCO = vData
End Property

Public Property Get GRPDESPBKP() As Integer
    GRPDESPBKP = intGRPDESPBKP
End Property

Public Property Let GRPDESPBKP(ByVal vData As Integer)
    intGRPDESPBKP = vData
End Property

Public Property Get GRPDESP() As Integer
    GRPDESP = intGRPDESP
End Property

Public Property Let GRPDESP(ByVal vData As Integer)
    intGRPDESP = vData
End Property

Public Property Get TRAVAIGF() As Integer
    TRAVAIGF = intTRAVAIGF
End Property

Public Property Let TRAVAIGF(ByVal vData As Integer)
    intTRAVAIGF = vData
End Property

Public Property Get LIBIGF() As Integer
    LIBIGF = intLIBIGF
End Property

Public Property Let LIBIGF(ByVal vData As Integer)
    intLIBIGF = vData
End Property

Public Property Get AVISAIGF() As Integer
    AVISAIGF = intAVISAIGF
End Property

Public Property Let AVISAIGF(ByVal vData As Integer)
    intAVISAIGF = vData
End Property

Public Property Get TRANSP() As Variant
    TRANSP = arrTRANSP
End Property

Public Property Let TRANSP(ByVal vData As Variant)
    arrTRANSP = vData
End Property

Public Property Get IQF() As Long
    IQF = lngIQF
End Property

Public Property Let IQF(ByVal vData As Long)
    lngIQF = vData
End Property

Public Property Get NATOPER() As Variant
    NATOPER = arrNATOPER
End Property

Public Property Let NATOPER(ByVal vData As Variant)
    arrNATOPER = vData
End Property

Public Property Get BANCOS() As Variant
    BANCOS = arrBANCOS
End Property

Public Property Let BANCOS(ByVal vData As Variant)
    arrBANCOS = vData
End Property

Public Property Get CEPRETI() As String
    CEPRETI = strCEPRETI
End Property

Public Property Let CEPRETI(ByVal vData As String)
    strCEPRETI = vData
End Property

Public Property Get ESTRETI() As Integer
    ESTRETI = intESTRETI
End Property

Public Property Let ESTRETI(ByVal vData As Integer)
    intESTRETI = vData
End Property

Public Property Get CIDRETI() As String
    CIDRETI = strCIDRETI
End Property

Public Property Let CIDRETI(ByVal vData As String)
    strCIDRETI = vData
End Property

Public Property Get BAIRETI() As String
    BAIRETI = strBAIRETI
End Property

Public Property Let BAIRETI(ByVal vData As String)
    strBAIRETI = vData
End Property

Public Property Get ENDRETI() As String
    ENDRETI = strENDRETI
End Property

Public Property Let ENDRETI(ByVal vData As String)
    strENDRETI = vData
End Property

Public Property Get CONTROLE() As String
    CONTROLE = strCONTROLE
End Property

Public Property Let CONTROLE(ByVal vData As String)
    strCONTROLE = vData
End Property

Public Property Get MODULO() As String
    MODULO = strModulo
End Property

Public Property Let MODULO(ByVal vData As String)
    strModulo = vData
End Property

Public Property Get PRODUTO() As Variant
    PRODUTO = arrPRODUTO
End Property

Public Property Let PRODUTO(ByVal vData As Variant)
    arrPRODUTO = vData
End Property

Public Property Get SITEFOR() As Variant
    SITEFOR = arrSITEFOR
End Property

Public Property Let SITEFOR(ByVal vData As Variant)
    arrSITEFOR = vData
End Property

Public Property Get EMAILFOR() As Variant
    EMAILFOR = arrEMAILFOR
End Property

Public Property Let EMAILFOR(ByVal vData As Variant)
    arrEMAILFOR = vData
End Property

Public Property Get CONTATOFOR() As Variant
    CONTATOFOR = arrCONTATOFOR
End Property

Public Property Let CONTATOFOR(ByVal vData As Variant)
    arrCONTATOFOR = vData
End Property

Public Property Get TELFOR() As Variant
    TELFOR = arrTELFOR
End Property

Public Property Let TELFOR(ByVal vData As Variant)
    arrTELFOR = vData
End Property

Public Property Get DTCADFOR() As Date
    DTCADFOR = dtDTCADFOR
End Property

Public Property Let DTCADFOR(ByVal vData As Date)
    dtDTCADFOR = vData
End Property

Public Property Get CEPFOR() As String
    CEPFOR = strCEPFOR
End Property

Public Property Let CEPFOR(ByVal vData As String)
    strCEPFOR = vData
End Property

Public Property Get ESTFOR() As Integer
    ESTFOR = intESTFOR
End Property

Public Property Let ESTFOR(ByVal vData As Integer)
    intESTFOR = vData
End Property

Public Property Get BAIFOR() As String
    BAIFOR = strBAIFOR
End Property

Public Property Let BAIFOR(ByVal vData As String)
    strBAIFOR = vData
End Property

Public Property Get CIDFOR() As String
    CIDFOR = strCIDFOR
End Property

Public Property Let CIDFOR(ByVal vData As String)
    strCIDFOR = vData
End Property

Public Property Get ENDFOR() As String
    ENDFOR = strENDFOR
End Property

Public Property Let ENDFOR(ByVal vData As String)
    strENDFOR = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Property Get CodigoFOR() As Integer
    CodigoFOR = intCodigoFOR
End Property

Public Property Let CodigoFOR(ByVal vData As Integer)
    intCodigoFOR = vData
End Property

Public Property Get CNPJCPFFOR() As String
    CNPJCPFFOR = strCNPJCPFFOR
End Property

Public Property Let CNPJCPFFOR(ByVal vData As String)
    strCNPJCPFFOR = vData
End Property

Public Property Get RAZSOCFOR() As String
    RAZSOCFOR = strRAZSOCFOR
End Property

Public Property Let RAZSOCFOR(ByVal vData As String)
    strRAZSOCFOR = vData
End Property

Public Property Get NOMFANTFOR() As String
    NOMFANTFOR = strNOMFANTFOR
End Property

Public Property Let NOMFANTFOR(ByVal vData As String)
    strNOMFANTFOR = vData
End Property


Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADFORNECP.cCaminho = cCaminho
    frmCADFORNECP.Linha = Linha
    frmCADFORNECP.FILIAL = iFilial
    frmCADFORNECP.strAcesso = strAcesso
    frmCADFORNECP.Show vbModal
End Sub

Public Sub cConnectConsulta(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, intCODIGO As Integer)
    frmCADFORNEC.cCaminho = cCaminho
    frmCADFORNEC.Linha = Linha
    frmCADFORNEC.FILIAL = iFilial
    frmCADFORNEC.strAcesso = strAcesso
    frmCADFORNEC.cTipOper = "C"
    frmCADFORNEC.iCodigo = intCODIGO
    frmCADFORNEC.strMODPAI = ""
    frmCADFORNEC.Show vbModal
End Sub


Public Function Pesq_CadFornec() As Boolean

   Pesq_CadFornec = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql + "       * " & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_CADFORNEC " & vbCrLf
   sSql = sSql + " Where " & vbCrLf
   sSql = sSql + "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_CadFornec = True
   BREC.Close

End Function

Public Sub PreencheComboProd(Combo As Variant)

     sSql = "Select " & vbCrLf
     sSql = sSql & "       PROD.* " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADPRODUTO PROD" & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       PROD.SGI_FILIAL      = " & intFilial & vbCrLf
     sSql = sSql & "   And PROD.SGI_PRODUTOTIPO = 0"
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_CODIGO) & " - " & Trim(BREC!SGI_DESCRICAO)
        Combo.ItemData(Combo.NewIndex) = BREC!SGI_IDPRODUTO
        BREC.MoveNext
     Loop
    
     If Combo.ListCount > 0 Then Combo.ListIndex = 0
    
     BREC.Close
    
End Sub

Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_Trans
     
     GRAVA = False
         
     Dim I As Integer
     Dim arrCamposGerais As Variant
     
     '' Inicia transação
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Then
        
        '' Grava Fornecedor
        sSql = "Insert Into SGI_CADFORNEC (SGI_FILIAL,SGI_CODIGO,SGI_CPFCNPJ,SGI_RAZAOSOC,SGI_NOMFANTA,SGI_ENDLOC,"
        sSql = sSql & "                    SGI_CIDADE,SGI_BAIRRO,SGI_ESTADO,SGI_CEP,SGI_DTCADSTRO,SGI_ENDRETI,SGI_BAIRETI,"
        sSql = sSql & "                    SGI_CIDRETI,SGI_ESTRETI,SGI_CEPRETI,SGI_IQF,SGI_IGFAVISA,SGI_IGFLIBERA,SGI_IGFTRAVA,SGI_STATUS,SGI_REQINSPQUA) Values (" & vbCrLf
        
        sSql = sSql & "                            " & intFilial & vbCrLf
        sSql = sSql & "                           ," & intCodigoFOR & vbCrLf
        sSql = sSql & "                           ,'" & strCNPJCPFFOR & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strRAZSOCFOR & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strNOMFANTFOR & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strENDFOR & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strCIDFOR & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strBAIFOR & "'" & vbCrLf
        sSql = sSql & "                           ," & intESTFOR & vbCrLf
        sSql = sSql & "                           ,'" & strCEPFOR & "'" & vbCrLf
        sSql = sSql & "                           ,'" & Format(dtDTCADFOR, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strENDRETI & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strBAIRETI & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strCIDRETI & "'" & vbCrLf
        sSql = sSql & "                           ,'" & intESTRETI & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strCEPRETI & "'" & vbCrLf
        sSql = sSql & "                           ," & lngIQF & vbCrLf
        
        If intAVISAIGF > 0 Then
           sSql = sSql & "                           ," & intAVISAIGF & vbCrLf
        Else
           sSql = sSql & "                           ,Null" & vbCrLf
        End If
        
        If intLIBIGF > 0 Then
           sSql = sSql & "                           ," & intLIBIGF & vbCrLf
        Else
           sSql = sSql & "                           ,Null" & vbCrLf
        End If
        
        If intTRAVAIGF > 0 Then
           sSql = sSql & "                           ," & intTRAVAIGF & vbCrLf
        Else
           sSql = sSql & "                           ,Null" & vbCrLf
        End If
        
        sSql = sSql & "                           ," & intSTATUS & vbCrLf
         
        sSql = sSql & "                           ," & intREQINSPQUA & vbCrLf
        
        sSql = sSql & "                           )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        ' Gravando Telefone
        If IsArray(arrTELFOR) = True Then
           For I = 0 To UBound(arrTELFOR)
                  arrCamposGerais = Split(arrTELFOR(I), "^")
                  sSql = "Insert Into SGI_DADOSGERAIS (SGI_FILIAL,SGI_CODIGO,SGI_MODULO,SGI_CONTROLE,SGI_TEXTO) Values (" & vbCrLf
                  sSql = sSql & "                            " & intFilial & vbCrLf
                  sSql = sSql & "                           ," & intCodigoFOR & vbCrLf
                  sSql = sSql & "                           ,'" & Trim(strModulo) & "'" & vbCrLf
                  sSql = sSql & "                           ,'" & Trim(arrCamposGerais(1)) & "'" & vbCrLf
                  sSql = sSql & "                           ,'" & Trim(arrCamposGerais(0)) & "'" & vbCrLf
                  sSql = sSql & "                           )"
               
                  BGRV.CommandText = sSql
                  BGRV.Execute
           Next I
        End If
        ' Gravando Contato
        If IsArray(arrCONTATOFOR) = True Then
           For I = 0 To UBound(arrCONTATOFOR)
                  arrCamposGerais = Split(arrCONTATOFOR(I), "^")
                  sSql = "Insert Into SGI_DADOSGERAIS (SGI_FILIAL,SGI_CODIGO,SGI_MODULO,SGI_CONTROLE,SGI_TEXTO) Values (" & vbCrLf
                  sSql = sSql & "                            " & intFilial & vbCrLf
                  sSql = sSql & "                           ," & intCodigoFOR & vbCrLf
                  sSql = sSql & "                           ,'" & Trim(strModulo) & "'" & vbCrLf
                  sSql = sSql & "                           ,'" & Trim(arrCamposGerais(1)) & "'" & vbCrLf
                  sSql = sSql & "                           ,'" & Trim(arrCamposGerais(0)) & "'" & vbCrLf
                  sSql = sSql & "                           )"
               
                  BGRV.CommandText = sSql
                  BGRV.Execute
           Next I
        End If
        ' Gravando E-MAIL
        If IsArray(arrEMAILFOR) = True Then
           For I = 0 To UBound(arrEMAILFOR)
                  arrCamposGerais = Split(arrEMAILFOR(I), "^")
                  sSql = "Insert Into SGI_DADOSGERAIS (SGI_FILIAL,SGI_CODIGO,SGI_MODULO,SGI_CONTROLE,SGI_TEXTO) Values (" & vbCrLf
                  sSql = sSql & "                            " & intFilial & vbCrLf
                  sSql = sSql & "                           ," & intCodigoFOR & vbCrLf
                  sSql = sSql & "                           ,'" & Trim(strModulo) & "'" & vbCrLf
                  sSql = sSql & "                           ,'" & Trim(arrCamposGerais(1)) & "'" & vbCrLf
                  sSql = sSql & "                           ,'" & Trim(arrCamposGerais(0)) & "'" & vbCrLf
                  sSql = sSql & "                           )"
               
                  BGRV.CommandText = sSql
                  BGRV.Execute
               
           Next I
        End If
        ' Gravando SITE
        If IsArray(arrSITEFOR) = True Then
           For I = 0 To UBound(arrSITEFOR)
                  arrCamposGerais = Split(arrSITEFOR(I), "^")
                  sSql = "Insert Into SGI_DADOSGERAIS (SGI_FILIAL,SGI_CODIGO,SGI_MODULO,SGI_CONTROLE,SGI_TEXTO) Values (" & vbCrLf
                  sSql = sSql & "                            " & intFilial & vbCrLf
                  sSql = sSql & "                           ," & intCodigoFOR & vbCrLf
                  sSql = sSql & "                           ,'" & Trim(strModulo) & "'" & vbCrLf
                  sSql = sSql & "                           ,'" & Trim(arrCamposGerais(1)) & "'" & vbCrLf
                  sSql = sSql & "                           ,'" & Trim(arrCamposGerais(0)) & "'" & vbCrLf
                  sSql = sSql & "                           )"
               
                  BGRV.CommandText = sSql
                  BGRV.Execute
           Next I
        End If
        ' Gravando Bancos
        If IsArray(arrBANCOS) = True Then
           For I = 1 To UBound(arrBANCOS)
               sSql = "Insert Into SGI_CADBCOFOR (SGI_FILIAL,SGI_CODIGO,SGI_CODBANCO) Values (" & vbCrLf
               sSql = sSql & "                            " & intFilial & vbCrLf
               sSql = sSql & "                           ," & intCodigoFOR & vbCrLf
               sSql = sSql & "                           ," & arrBANCOS(I) & vbCrLf
               sSql = sSql & "                           )"
               
               BGRV.CommandText = sSql
               BGRV.Execute
           Next I
        End If
        ' Gravando PRODUTO
        If IsArray(arrPRODUTO) = True Then
           For I = 1 To UBound(arrPRODUTO)
               
               sSql = "Insert Into SGI_PRODFOR (SGI_FILIAL,SGI_CODIGO,SGI_PRODUTO,SGI_PRODUTOFORN,SGI_PORCINSP) Values (" & vbCrLf
               sSql = sSql & "                            " & intFilial & vbCrLf
               sSql = sSql & "                           ," & intCodigoFOR & vbCrLf
               sSql = sSql & "                           ," & Trim(arrPRODUTO(I, 1)) & vbCrLf
               
               If Len(Trim(arrPRODUTO(I, 2))) > 0 Then
                    sSql = sSql & "                           ,'" & Trim(arrPRODUTO(I, 2)) & "'" & vbCrLf
               Else
                    sSql = sSql & "                           ,Null" & vbCrLf
               End If
               
               If Len(Trim(arrPRODUTO(I, 3))) > 0 Then
                    sSql = sSql & "                           ," & Trim(arrPRODUTO(I, 3)) & vbCrLf
               Else
                    sSql = sSql & "                           ,Null" & vbCrLf
               End If
               
               sSql = sSql & "                           )"
               
               BGRV.CommandText = sSql
               BGRV.Execute
           Next I
        End If
        ' Gravando NATUREZA DE OPERAÇÃO
        If IsArray(arrNATOPER) = True Then
           For I = 1 To UBound(arrNATOPER)
               sSql = "Insert Into SGI_FORNATOPER (SGI_FILIAL,SGI_CODIGO,SGI_CODNATOPER) Values (" & vbCrLf
               sSql = sSql & "                            " & intFilial & vbCrLf
               sSql = sSql & "                           ," & intCodigoFOR & vbCrLf
               sSql = sSql & "                           ,'" & Trim(arrNATOPER(I)) & "'" & vbCrLf
               sSql = sSql & "                           )"
               
               BGRV.CommandText = sSql
               BGRV.Execute
           Next I
        End If
        '' Grava Transportadoras
        If IsArray(arrTRANSP) = True Then
           For I = 1 To UBound(arrTRANSP)
               sSql = "Insert Into SGI_FORTRANSP (SGI_FILIAL,SGI_CODIGO,SGI_CODTRANSP) Values (" & vbCrLf
               sSql = sSql & "                            " & intFilial & vbCrLf
               sSql = sSql & "                           ," & intCodigoFOR & vbCrLf
               sSql = sSql & "                           ," & arrTRANSP(I) & vbCrLf
               sSql = sSql & "                           )"
               
               BGRV.CommandText = sSql
               BGRV.Execute
           Next I
        End If
        
        '' Grava Grupo de Despesas
        If intGRPDESP > 0 Then
           sSql = "Insert Into SGI_GRPDESPFORN (SGI_FILIAL,SGI_CODGRPDSP,SGI_CODFORN) " & vbCrLf
           sSql = sSql & "                     Values ( " & vbCrLf
           sSql = sSql & "                              " & intFilial & vbCrLf
           sSql = sSql & "                             ," & intGRPDESP & vbCrLf
           sSql = sSql & "                             ," & intCodigoFOR & vbCrLf
           sSql = sSql & "                            ) "
               
           BGRV.CommandText = sSql
           BGRV.Execute
        End If
        
        '' Gravando Nivel de Risco
        If lngCODNIVRISCO > 0 Then
           
           sSql = "Insert Into SGI_CADRISCOFORNEC (SGI_FILIAL,SGI_CODIGO,SGI_CODFORNEC) " & vbCrLf
           sSql = sSql & "                        Values ( "
           sSql = sSql & "                                 " & intFilial & vbCrLf
           sSql = sSql & "                                ," & lngCODNIVRISCO & vbCrLf
           sSql = sSql & "                                ," & intCodigoFOR & vbCrLf
           sSql = sSql & "                            ) "
           
           BGRV.CommandText = sSql
           BGRV.Execute
           
        End If
        
        '' Gravando Produtos Não Inspecionados
        If IsArray(arrPRODNAOINSP) Then
            For I = 1 To UBound(arrPRODNAOINSP)
                sSql = "Insert Into SGI_PRODNAOINSP (SGI_FILIAL,SGI_CODIGO,SGI_PRODUTO) " & vbCrLf
                sSql = sSql & "                        Values ( "
                sSql = sSql & "                                 " & intFilial & vbCrLf
                sSql = sSql & "                                ," & intCodigoFOR & vbCrLf
                sSql = sSql & "                                ,'" & Trim(arrPRODNAOINSP(I)) & "'" & vbCrLf
                sSql = sSql & "                            ) "
                
                BGRV.CommandText = sSql
                BGRV.Execute
            Next I
        End If
     
     ElseIf strOPER = "A" Then
        
        '' --------------------------------------------------------
        '' Apagando Produtos não Inspecionados pelo Fornecedor
        sSql = "Delete from SGI_PRODNAOINSP " & vbCrLf
        sSql = sSql & "      Where SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCodigoFOR & vbCrLf
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------------------------------------
        '' Apagando dados do fornecedor
        sSql = "Delete from SGI_DADOSGERAIS " & vbCrLf
        sSql = sSql & "      Where SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCodigoFOR & vbCrLf
        sSql = sSql & "        And sgi_MODULO = '" & strModulo & "'"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------------------------------------
        '' Apagando Bancos do fornecedor
        sSql = "Delete from SGI_CADBCOFOR " & vbCrLf
        sSql = sSql & "      Where SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCodigoFOR & vbCrLf
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------------------------------------
        '' Apagando Produtos do fornecedor
        sSql = "Delete from SGI_PRODFOR " & vbCrLf
        sSql = sSql & "      Where SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCodigoFOR & vbCrLf
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------------------------------------
        sSql = "Delete from SGI_FORNATOPER " & vbCrLf
        sSql = sSql & "      Where SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCodigoFOR & vbCrLf
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------------------------------------
        sSql = "Delete from SGI_FORTRANSP " & vbCrLf
        sSql = sSql & "      Where SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCodigoFOR & vbCrLf
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        
        '' Deleta Grupo de Despesas
        sSql = "Delete From SGI_GRPDESPFORN " & vbCrLf
        sSql = sSql & "                      Where " & vbCrLf
        sSql = sSql & "                            SGI_FILIAL  = " & intFilial & vbCrLf
        sSql = sSql & "                        And SGI_CODFORN = " & intCodigoFOR
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------------------------------------
        '' Grava Fornecedor
        sSql = "Update SGI_CADFORNEC Set " & vbCrLf
        sSql = sSql & "                   SGI_CPFCNPJ   = '" & strCNPJCPFFOR & "'" & vbCrLf
        sSql = sSql & "                  ,SGI_RAZAOSOC  = '" & strRAZSOCFOR & "'" & vbCrLf
        sSql = sSql & "                  ,SGI_NOMFANTA  = '" & strNOMFANTFOR & "'" & vbCrLf
        sSql = sSql & "                  ,SGI_ENDLOC    = '" & strENDFOR & "'" & vbCrLf
        sSql = sSql & "                  ,SGI_CIDADE    = '" & strCIDFOR & "'" & vbCrLf
        sSql = sSql & "                  ,SGI_BAIRRO    = '" & strBAIFOR & "'" & vbCrLf
        sSql = sSql & "                  ,SGI_ESTADO    = " & intESTFOR & vbCrLf
        sSql = sSql & "                  ,SGI_CEP       = '" & strCEPFOR & "'" & vbCrLf
        sSql = sSql & "                  ,SGI_DTCADSTRO = '" & Format(dtDTCADFOR, "MM/DD/YYYY") & "'" & vbCrLf
        sSql = sSql & "                  ,SGI_ENDRETI   = '" & strENDRETI & "'" & vbCrLf
        sSql = sSql & "                  ,SGI_BAIRETI   = '" & strBAIRETI & "'" & vbCrLf
        sSql = sSql & "                  ,SGI_CIDRETI   = '" & strCIDRETI & "'" & vbCrLf
        sSql = sSql & "                  ,SGI_ESTRETI   = '" & intESTRETI & "'" & vbCrLf
        sSql = sSql & "                  ,SGI_CEPRETI   = '" & strCEPRETI & "'" & vbCrLf
        sSql = sSql & "                  ,SGI_IQF       = " & lngIQF & vbCrLf
        
        If intAVISAIGF > 0 Then
           sSql = sSql & "                  ,SGI_IGFAVISA  = " & intAVISAIGF & vbCrLf
        Else
           sSql = sSql & "                  ,SGI_IGFAVISA  = Null" & vbCrLf
        End If
        If intLIBIGF > 0 Then
           sSql = sSql & "                  ,SGI_IGFLIBERA = " & intLIBIGF & vbCrLf
        Else
           sSql = sSql & "                  ,SGI_IGFLIBERA = Null" & vbCrLf
        End If
        If intTRAVAIGF > 0 Then
           sSql = sSql & "                  ,SGI_IGFTRAVA  = " & intTRAVAIGF & vbCrLf
        Else
           sSql = sSql & "                  ,SGI_IGFTRAVA  = Null" & vbCrLf
        End If
        
        sSql = sSql & "                  ,SGI_REQINSPQUA = " & intREQINSPQUA & vbCrLf
        
        sSql = sSql & "            Where " & vbCrLf
        sSql = sSql & "                  SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "              And SGI_CODIGO = " & intCodigoFOR
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------------------------------------
        ' Gravando Telefone
        If IsArray(arrTELFOR) = True Then
              For I = 0 To UBound(arrTELFOR)
                  arrCamposGerais = Split(arrTELFOR(I), "^")
                  sSql = "Insert Into SGI_DADOSGERAIS (SGI_FILIAL,SGI_CODIGO,SGI_MODULO,SGI_CONTROLE,SGI_TEXTO) Values (" & vbCrLf
                  sSql = sSql & "                            " & intFilial & vbCrLf
                  sSql = sSql & "                           ," & intCodigoFOR & vbCrLf
                  sSql = sSql & "                           ,'" & Trim(strModulo) & "'" & vbCrLf
                  sSql = sSql & "                           ,'" & Trim(arrCamposGerais(1)) & "'" & vbCrLf
                  sSql = sSql & "                           ,'" & Trim(arrCamposGerais(0)) & "'" & vbCrLf
                  sSql = sSql & "                           )"
               
                  BGRV.CommandText = sSql
                  BGRV.Execute
              Next I
        End If
        '' --------------------------------------------------------
        ' Gravando Contato
        If IsArray(arrCONTATOFOR) = True Then
              For I = 0 To UBound(arrCONTATOFOR)
                  arrCamposGerais = Split(arrCONTATOFOR(I), "^")
                  sSql = "Insert Into SGI_DADOSGERAIS (SGI_FILIAL,SGI_CODIGO,SGI_MODULO,SGI_CONTROLE,SGI_TEXTO) Values (" & vbCrLf
                  sSql = sSql & "                            " & intFilial & vbCrLf
                  sSql = sSql & "                           ," & intCodigoFOR & vbCrLf
                  sSql = sSql & "                           ,'" & Trim(strModulo) & "'" & vbCrLf
                  sSql = sSql & "                           ,'" & Trim(arrCamposGerais(1)) & "'" & vbCrLf
                  sSql = sSql & "                           ,'" & Trim(arrCamposGerais(0)) & "'" & vbCrLf
                  sSql = sSql & "                           )"
               
                  BGRV.CommandText = sSql
                  BGRV.Execute
              Next I
        End If
        '' --------------------------------------------------------
        ' Gravando E-MAIL
        If IsArray(arrEMAILFOR) = True Then
              For I = 0 To UBound(arrEMAILFOR)
                  arrCamposGerais = Split(arrEMAILFOR(I), "^")
                  sSql = "Insert Into SGI_DADOSGERAIS (SGI_FILIAL,SGI_CODIGO,SGI_MODULO,SGI_CONTROLE,SGI_TEXTO) Values (" & vbCrLf
                  sSql = sSql & "                            " & intFilial & vbCrLf
                  sSql = sSql & "                           ," & intCodigoFOR & vbCrLf
                  sSql = sSql & "                           ,'" & Trim(strModulo) & "'" & vbCrLf
                  sSql = sSql & "                           ,'" & Trim(arrCamposGerais(1)) & "'" & vbCrLf
                  sSql = sSql & "                           ,'" & Trim(arrCamposGerais(0)) & "'" & vbCrLf
                  sSql = sSql & "                           )"
               
                  BGRV.CommandText = sSql
                  BGRV.Execute
              Next I
        End If
        '' --------------------------------------------------------
        ' Gravando SITE
        If IsArray(arrSITEFOR) = True Then
              For I = 0 To UBound(arrSITEFOR)
                  arrCamposGerais = Split(arrSITEFOR(I), "^")
                  sSql = "Insert Into SGI_DADOSGERAIS (SGI_FILIAL,SGI_CODIGO,SGI_MODULO,SGI_CONTROLE,SGI_TEXTO) Values (" & vbCrLf
                  sSql = sSql & "                            " & intFilial & vbCrLf
                  sSql = sSql & "                           ," & intCodigoFOR & vbCrLf
                  sSql = sSql & "                           ,'" & Trim(strModulo) & "'" & vbCrLf
                  sSql = sSql & "                           ,'" & Trim(arrCamposGerais(1)) & "'" & vbCrLf
                  sSql = sSql & "                           ,'" & Trim(arrCamposGerais(0)) & "'" & vbCrLf
                  sSql = sSql & "                           )"
               
                  BGRV.CommandText = sSql
                  BGRV.Execute
              Next I
        End If
        '' --------------------------------------------------------
        ' Gravando Bancos
        If IsArray(arrBANCOS) = True Then
           If UBound(arrBANCOS) > 0 Then
              For I = 1 To UBound(arrBANCOS)
                  sSql = "Insert Into SGI_CADBCOFOR (SGI_FILIAL,SGI_CODIGO,SGI_CODBANCO) Values (" & vbCrLf
                  sSql = sSql & "                            " & intFilial & vbCrLf
                  sSql = sSql & "                           ," & intCodigoFOR & vbCrLf
                  sSql = sSql & "                           ," & arrBANCOS(I) & vbCrLf
                  sSql = sSql & "                           )"
               
                  BGRV.CommandText = sSql
                  BGRV.Execute
              Next I
           End If
        End If
        '' --------------------------------------------------------
        ' Gravando PRODUTO
        If IsArray(arrPRODUTO) = True Then
           If UBound(arrPRODUTO) > 0 Then
              For I = 1 To UBound(arrPRODUTO)
                  sSql = "Insert Into SGI_PRODFOR (SGI_FILIAL,SGI_CODIGO,SGI_PRODUTO,SGI_PRODUTOFORN,SGI_PORCINSP) Values (" & vbCrLf
                  sSql = sSql & "                            " & intFilial & vbCrLf
                  sSql = sSql & "                           ," & intCodigoFOR & vbCrLf
                  sSql = sSql & "                           ,'" & Trim(arrPRODUTO(I, 1)) & "'" & vbCrLf
                  sSql = sSql & "                           ,'" & Trim(arrPRODUTO(I, 2)) & "'" & vbCrLf
                  
                  If Len(Trim(arrPRODUTO(I, 3))) = 0 Then
                     sSql = sSql & "                           ,Null" & vbCrLf
                  Else
                     sSql = sSql & "                           ," & Trim(arrPRODUTO(I, 3)) & vbCrLf
                  End If
                  
                  sSql = sSql & "                           )"
        
                  BGRV.CommandText = sSql
                  BGRV.Execute
              Next I
           End If
        End If
        '' --------------------------------------------------------
        '' Gravando NATUREZA DE OPERAÇÃO
        If IsArray(arrNATOPER) = True Then
           If UBound(arrNATOPER) > 0 Then
              For I = 1 To UBound(arrNATOPER)
                  sSql = "Insert Into SGI_FORNATOPER (SGI_FILIAL,SGI_CODIGO,SGI_CODNATOPER) Values (" & vbCrLf
                  sSql = sSql & "                            " & intFilial & vbCrLf
                  sSql = sSql & "                           ," & intCodigoFOR & vbCrLf
                  sSql = sSql & "                           ,'" & Trim(arrNATOPER(I)) & "'" & vbCrLf
                  sSql = sSql & "                           )"
               
                  BGRV.CommandText = sSql
                  BGRV.Execute
              Next I
           End If
        End If
        '' --------------------------------------------------------
        '' Grava Transportadoras
        If IsArray(arrTRANSP) = True Then
           If UBound(arrTRANSP) > 0 Then
              For I = 1 To UBound(arrTRANSP)
                  sSql = "Insert Into SGI_FORTRANSP (SGI_FILIAL,SGI_CODIGO,SGI_CODTRANSP) Values (" & vbCrLf
                  sSql = sSql & "                            " & intFilial & vbCrLf
                  sSql = sSql & "                           ," & intCodigoFOR & vbCrLf
                  sSql = sSql & "                           ," & arrTRANSP(I) & vbCrLf
                  sSql = sSql & "                           )"
               
                  BGRV.CommandText = sSql
                  BGRV.Execute
              Next I
           End If
        End If
        '' --------------------------------------------------------
        '' Grava Grupo de Despesas
        If intGRPDESP > 0 Then
        
           sSql = "Insert Into SGI_GRPDESPFORN (SGI_FILIAL,SGI_CODGRPDSP,SGI_CODFORN) " & vbCrLf
           sSql = sSql & "                     Values ( " & vbCrLf
           sSql = sSql & "                              " & intFilial & vbCrLf
           sSql = sSql & "                             ," & intGRPDESP & vbCrLf
           sSql = sSql & "                             ," & intCodigoFOR & vbCrLf
           sSql = sSql & "                            ) "
               
           BGRV.CommandText = sSql
           BGRV.Execute
           
           If intGRPDESP <> intGRPDESPBKP Then
           
              sSql = "Update SGI_CONTASHAPG Set SGI_GRPDESP = " & intGRPDESP & vbCrLf
              sSql = sSql & "                                Where " & vbCrLf
              sSql = sSql & "                                      SGI_FILIAL  = " & intFilial & vbCrLf
              sSql = sSql & "                                  And SGI_GRPDESP = " & intGRPDESPBKP & vbCrLf
              sSql = sSql & "                                  And SGI_CODFOR  = " & intCodigoFOR
              
              BGRV.CommandText = sSql
              BGRV.Execute
           
           End If
           
        End If
    
        '' Apagando Nivel de Risco do Fornecedor
        sSql = "Delete From  SGI_CADRISCOFORNEC " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL    = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODFORNEC = " & intCodigoFOR & vbCrLf
           
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Gravando Nivel de Risco
        If lngCODNIVRISCO > 0 Then
           
           sSql = "Insert Into SGI_CADRISCOFORNEC (SGI_FILIAL,SGI_CODIGO,SGI_CODFORNEC) " & vbCrLf
           sSql = sSql & "                        Values ( "
           sSql = sSql & "                                 " & intFilial & vbCrLf
           sSql = sSql & "                                ," & lngCODNIVRISCO & vbCrLf
           sSql = sSql & "                                ," & intCodigoFOR & vbCrLf
           sSql = sSql & "                            ) "
           
           BGRV.CommandText = sSql
           BGRV.Execute
           
        End If
        
        '' Gravando Produtos Não Inspecionados
        If IsArray(arrPRODNAOINSP) Then
            For I = 1 To UBound(arrPRODNAOINSP)
                sSql = "Insert Into SGI_PRODNAOINSP (SGI_FILIAL,SGI_CODIGO,SGI_PRODUTO) " & vbCrLf
                sSql = sSql & "                        Values ( "
                sSql = sSql & "                                 " & intFilial & vbCrLf
                sSql = sSql & "                                ," & intCodigoFOR & vbCrLf
                sSql = sSql & "                                ,'" & Trim(arrPRODNAOINSP(I)) & "'" & vbCrLf
                sSql = sSql & "                            ) "
                
                BGRV.CommandText = sSql
                BGRV.Execute
            Next I
        End If
     
     ElseIf strOPER = "E" Then
     
        '' --------------------------------------------------------
        '' Apagando Produtos não Inspecionados pelo Fornecedor
        sSql = "Delete from SGI_PRODNAOINSP " & vbCrLf
        sSql = sSql & "      Where SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCodigoFOR & vbCrLf
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------------------------------------
        '' Apagando Transportadora do Fornecedor
        sSql = "Delete from SGI_FORTRANSP " & vbCrLf
        sSql = sSql & "      Where SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCodigoFOR & vbCrLf
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------------------------------------
        '' Apagando dados do fornecedor
        sSql = "Delete from SGI_DADOSGERAIS " & vbCrLf
        sSql = sSql & "      Where SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCodigoFOR & vbCrLf
        sSql = sSql & "        And sgi_MODULO = '" & strModulo & "'"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------------------------------------
        '' Apagando Bancos do fornecedor
        sSql = "Delete from SGI_CADBCOFOR " & vbCrLf
        sSql = sSql & "      Where SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCodigoFOR & vbCrLf
        
        '' --------------------------------------------------------
        '' Apagando Produtos do fornecedor
        sSql = "Delete from SGI_PRODFOR " & vbCrLf
        sSql = sSql & "      Where SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCodigoFOR & vbCrLf
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------------------------------------
        '' Apagando Fornecedor
        sSql = "Delete from SGI_CADFORNEC " & vbCrLf
        sSql = sSql & "      Where SGI_FILIAL = " & FILIAL & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCodigoFOR & vbCrLf
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------------------------------------
        '' Apagando NATUREZA DE OPERAÇÃO
        sSql = "Delete from SGI_FORNATOPER " & vbCrLf
        sSql = sSql & "      Where SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCodigoFOR & vbCrLf
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
       
        '' --------------------------------------------------------
        '' Deleta Grupo de Despesas
        sSql = "Delete From SGI_GRPDESPFORN " & vbCrLf
        sSql = sSql & "                      Where " & vbCrLf
        sSql = sSql & "                            SGI_FILIAL  = " & intFilial & vbCrLf
        sSql = sSql & "                        And SGI_CODFORN = " & intCodigoFOR
        
        BGRV.CommandText = sSql
        BGRV.Execute
     
        
        '' --------------------------------------------------------
        '' Deleta Grupo Nivel de Risco do Fornecedor
        sSql = "Delete From  SGI_CADRISCOFORNEC " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL    = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODFORNEC = " & intCodigoFOR & vbCrLf
           
        BGRV.CommandText = sSql
        BGRV.Execute
     
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
   sSql = sSql & "       SGI_CADFORNEC " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & intCodigoFOR & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC.EOF Then
       
      intCodigoFOR = BREC!SGI_CODIGO
      strCNPJCPFFOR = BREC!SGI_CPFCNPJ
      strRAZSOCFOR = BREC!SGI_RAZAOSOC
      strNOMFANTFOR = BREC!SGI_NOMFANTA
      strENDFOR = BREC!SGI_ENDLOC
      strCIDFOR = BREC!SGI_CIDADE
      strBAIFOR = BREC!SGI_BAIRRO
      intESTFOR = BREC!SGI_ESTADO
      strCEPFOR = BREC!SGI_CEP
      dtDTCADFOR = BREC!SGI_DTCADSTRO
      strENDRETI = BREC!SGI_ENDRETI
      strBAIRETI = BREC!SGI_BAIRETI
      strCIDRETI = BREC!SGI_CIDRETI
      intESTRETI = BREC!SGI_ESTRETI
      strCEPRETI = BREC!SGI_CEPRETI
      If Not IsNull(BREC!SGI_IQF) Then lngIQF = BREC!SGI_IQF
      If Not IsNull(BREC!SGI_IGFAVISA) Then intAVISAIGF = BREC!SGI_IGFAVISA
      If Not IsNull(BREC!SGI_IGFLIBERA) Then intLIBIGF = BREC!SGI_IGFLIBERA
      If Not IsNull(BREC!SGI_IGFTRAVA) Then intTRAVAIGF = BREC!SGI_IGFTRAVA
      intSTATUS = BREC!SGI_STATUS
      intREQINSPQUA = BREC!SGI_REQINSPQUA
     
      Carrega_campos = True
   End If
   
   BREC.Close

   If Carrega_campos = False Then Exit Function
   
   '' ------------------------------------------------------------
   '' Telefone
   
   BREC2.ActiveConnection = adoBanco_Dados
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_DADOSGERAIS " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL   = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_CODIGO   = " & intCodigoFOR & vbCrLf
   sSql = sSql & "   And SGI_MODULO   = '" & strModulo & "'" & vbCrLf
   sSql = sSql & "   And SGI_CONTROLE = 'cboTel'" & vbCrLf
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC2.EOF Then
   
      intReg = 0
      Do While Not BREC2.EOF
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
      
      ReDim arrTELFOR(1 To intReg)
      
      intReg = 1
      BREC2.MoveFirst
      Do While Not BREC2.EOF
         arrTELFOR(intReg) = BREC2!SGI_TEXTO
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
   
   End If
   
   BREC2.Close
   
   '' ------------------------------------------------------------
   '' Contato
   
   BREC2.ActiveConnection = adoBanco_Dados
      
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_DADOSGERAIS " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL   = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_CODIGO   = " & intCodigoFOR & vbCrLf
   sSql = sSql & "   And SGI_MODULO   = '" & strModulo & "'" & vbCrLf
   sSql = sSql & "   And SGI_CONTROLE = 'cboContato'" & vbCrLf
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC2.EOF Then
   
      intReg = 0
      Do While Not BREC2.EOF
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
      
      ReDim arrCONTATOFOR(1 To intReg)
      
      intReg = 1
      BREC2.MoveFirst
      Do While Not BREC2.EOF
         arrCONTATOFOR(intReg) = BREC2!SGI_TEXTO
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
   
   End If
   
   BREC2.Close
   
   '' ------------------------------------------------------------
   '' E-MAIL
   
   BREC2.ActiveConnection = adoBanco_Dados
      
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_DADOSGERAIS " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL   = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_CODIGO   = " & intCodigoFOR & vbCrLf
   sSql = sSql & "   And SGI_MODULO   = '" & strModulo & "'" & vbCrLf
   sSql = sSql & "   And SGI_CONTROLE = 'cboEmail'" & vbCrLf
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC2.EOF Then
   
      intReg = 0
      Do While Not BREC2.EOF
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
      
      ReDim arrEMAILFOR(1 To intReg)
      
      intReg = 1
      BREC2.MoveFirst
      Do While Not BREC2.EOF
         arrEMAILFOR(intReg) = BREC2!SGI_TEXTO
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
   
   End If
   
   BREC2.Close
   
   '' ------------------------------------------------------------
   '' SITE
   
   BREC2.ActiveConnection = adoBanco_Dados
      
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_DADOSGERAIS " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL   = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_CODIGO   = " & intCodigoFOR & vbCrLf
   sSql = sSql & "   And SGI_MODULO   = '" & strModulo & "'" & vbCrLf
   sSql = sSql & "   And SGI_CONTROLE = 'cboSite'" & vbCrLf
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC2.EOF Then
   
      intReg = 0
      Do While Not BREC2.EOF
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
      
      ReDim arrSITEFOR(1 To intReg)
      
      intReg = 1
      BREC2.MoveFirst
      Do While Not BREC2.EOF
         arrSITEFOR(intReg) = BREC2!SGI_TEXTO
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
   
   End If
   
   BREC2.Close
   
   '' ------------------------------------------------------------
   '' Produto
   BREC2.ActiveConnection = adoBanco_Dados
      
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_PRODFOR " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL   = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_CODIGO   = " & intCodigoFOR & vbCrLf
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC2.EOF Then
   
      intReg = 0
      Do While Not BREC2.EOF
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
      
      ReDim arrPRODUTO(1 To intReg, 1 To 3)
      
      intReg = 1
      BREC2.MoveFirst
      Do While Not BREC2.EOF
         arrPRODUTO(intReg, 1) = Trim(BREC2!SGI_PRODUTO)
         arrPRODUTO(intReg, 2) = Trim(BREC2!SGI_PRODUTOFORN)
         
         If Not IsNull(BREC2!SGI_PORCINSP) Then arrPRODUTO(intReg, 3) = Str(BREC2!SGI_PORCINSP)
         If IsNull(BREC2!SGI_PORCINSP) Then arrPRODUTO(intReg, 3) = ""
         
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
   
   End If
   
   BREC2.Close
   
   '' ------------------------------------------------------------
   '' Bancos
   BREC2.ActiveConnection = adoBanco_Dados
      
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADBCOFOR " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL   = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_CODIGO   = " & intCodigoFOR & vbCrLf
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC2.EOF Then
   
      intReg = 0
      Do While Not BREC2.EOF
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
      
      ReDim arrBANCOS(1 To intReg) As String
      
      intReg = 1
      BREC2.MoveFirst
      Do While Not BREC2.EOF
         arrBANCOS(intReg) = Str(BREC2!SGI_CODBANCO)
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
   
   End If
   
   BREC2.Close

   '' ------------------------------------------------------------
   '' Natureza de Operação
   BREC2.ActiveConnection = adoBanco_Dados
      
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_FORNATOPER " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL   = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_CODIGO   = " & intCodigoFOR & vbCrLf
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC2.EOF Then
   
      intReg = 0
      Do While Not BREC2.EOF
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
      
      ReDim arrNATOPER(1 To intReg) As String
      
      intReg = 1
      BREC2.MoveFirst
      Do While Not BREC2.EOF
         arrNATOPER(intReg) = BREC2!SGI_CODNATOPER
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
   
   End If
   
   BREC2.Close

   '' ------------------------------------------------------------
   '' Transportadoras
   BREC2.ActiveConnection = adoBanco_Dados
      
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_FORTRANSP " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL   = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_CODIGO   = " & intCodigoFOR & vbCrLf
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC2.EOF Then
   
      intReg = 0
      Do While Not BREC2.EOF
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
      
      ReDim arrTRANSP(1 To intReg) As Long
      
      intReg = 1
      BREC2.MoveFirst
      Do While Not BREC2.EOF
         arrTRANSP(intReg) = BREC2!SGI_CODTRANSP
         intReg = intReg + 1
         BREC2.MoveNext
      Loop
   
   End If
   
   BREC2.Close
   
   '' ----------------------------------------------------------
   '' Grupo de Despesas
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_GRPDESPFORN " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL  = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_CODFORN = " & intCodigoFOR
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC2.EOF Then
      intGRPDESP = BREC2!SGI_CODGRPDSP
      intGRPDESPBKP = BREC2!SGI_CODGRPDSP
   End If
   BREC2.Close
   '' ----------------------------------------------------------
   
   
   '' Pega Qual o Nivel de Risco do Fornecedor
   BREC3.ActiveConnection = adoBanco_Dados
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADRISCOFORNEC " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL    = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_CODFORNEC = " & intCodigoFOR
   
   BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC3.EOF Then lngCODNIVRISCO = BREC3!SGI_CODIGO
   BREC3.Close
   

   '' ------------------------------------------------------------
   '' Produto
   BREC4.ActiveConnection = adoBanco_Dados
      
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_PRODNAOINSP " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL   = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_CODIGO   = " & intCodigoFOR & vbCrLf
   
   BREC4.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC4.EOF Then
   
      intReg = 0
      Do While Not BREC4.EOF
         intReg = intReg + 1
         BREC4.MoveNext
      Loop
      
      ReDim arrPRODNAOINSP(1 To intReg)
      
      intReg = 1
      BREC4.MoveFirst
      Do While Not BREC4.EOF
         arrPRODNAOINSP(intReg) = Trim(BREC4!SGI_PRODUTO)
         intReg = intReg + 1
         BREC4.MoveNext
      Loop
   
   End If
   BREC4.Close
   '' ------------------------------------------------------------
   
End Function

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
    
    If Combo.ListCount > 0 Then Combo.ListIndex = 0
    
    BREC.Close

End Sub


Public Sub PreencheComboNatOper(Combo As Variant)

     Combo.Clear
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       * " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADNATOPERACAO " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL  = " & intFilial & vbCrLf
     sSql = sSql & "   And SGI_ENTSAI  = 0"
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_CODIGO) & Space(5 - Len(Trim(BREC!SGI_CODIGO))) & " - " & Trim(BREC!SGI_DESCRICAO)
        BREC.MoveNext
     Loop
    
     BREC.Close
    
End Sub

Public Sub PreenchComboTransportadoras(Combo As Variant)

    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADTRANSP " & vbCrLf
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

Public Sub PreenchComboNivelRiscos(Combo As Variant)

    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADRISCO " & vbCrLf
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

Public Sub PreenchComboGrpDespesas(Combo As Variant)

    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADGRUPDESP " & vbCrLf
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

