VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADPRODU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim strCodigoProd    As String
Dim strCodigoProdANT As String
Dim strDescriProd    As String
Dim intEspProduto    As Integer
Dim intCODGRUPPROD   As Integer
Dim intCODSUBGPROD   As Integer
Dim intTIPPRODUTO    As Integer
Dim intUnidade       As Integer
Dim dblSaldo         As Double
Dim dblEstMin        As Double
Dim dtDataCadast     As String
Dim dtDataUltMov     As Date
Dim intFilial        As Integer
Dim arrProdLST       As Variant
Dim arrESPTECPROD    As Variant
Dim arrPROCESSOS     As Variant
Dim arrFERRAMENTA    As Variant
Dim lngCODLCTO       As Long
Dim lngCLASFISC      As Long
Dim intATAUTOMSN     As Integer
Dim dblPRCPROD       As Double
Dim dblPRCMED        As Double
Dim strDESCEQUIP     As String
Dim strPROCEDEN      As String
Dim curPORCACRES     As Currency
Dim curPRCCUSTO      As Currency
Dim intPRODUTOTIPO   As Integer
Dim intESTILOPROD    As Integer
Dim curPESOUNIT      As Currency
Dim lngTABCONV       As Long
Dim strCODPROFORNEC  As String
Dim strRUA           As String
Dim strBOX           As String
Dim strPRATELEIRA    As String
Dim strCODEAN        As String
Dim arrPRODFOR       As Variant
Dim arrPRODTIPORCA   As Variant
Dim arrPRODVAPROC    As Variant
Dim arrPRODVAREC     As Variant
Dim arrPRODVAINSP    As Variant
Dim arrPRODATRPROC   As Variant
Dim arrPRODATRREC    As Variant
Dim arrPRODATRINSP   As Variant
Dim curDISTANCIA     As Currency
Dim intPADRAO        As Integer
Dim curCORRELACAO    As Currency
Dim lngFAMMAQUINA    As Long
Dim arrCORRPRODPADR         As Variant
Dim arrPARPROD              As Variant
Dim curPORCCORR             As Currency
Dim curCUBAGEN              As Currency
Dim arrPRODATECLIE          As Variant
Dim curMetros               As Currency
Dim curLargura              As Currency
Dim curGRAMATURA2           As Currency
Dim arrUNIDCONV             As Variant
Dim lngCodLinProd           As Long
Dim lngCodClie              As Long
Dim lngCodRotulo            As Long
Dim lngDigVerif             As Long
Dim lngIDProduto            As Long
Dim arrCORES                As Variant
Dim arrVERNIZ               As Variant
Dim arrVERNIZ02             As Variant
Dim arrVERNIZACAB           As Variant
Dim arrESMALTE              As Variant
Dim arrVERNIZEXT            As Variant
Dim lngVernCorpo            As Long
Dim lngVernTampa            As Long
Dim lngVernFundo            As Long
Dim lngVernArgola           As Long
Dim lngEspessCorpo          As Long
Dim lngEspessTampa          As Long
Dim lngEspessFundo          As Long
Dim lngEspessArgola         As Long
Dim lngRevestCorpo          As Long
Dim lngRevestCorpo2         As Long
Dim lngRevestTampa          As Long
Dim lngRevestTampa2         As Long
Dim lngRevestFundo          As Long
Dim lngRevestFundo2         As Long
Dim lngRevestArgola         As Long
Dim lngRevestArgola2        As Long
Dim lngAlcaPlastica         As Long
Dim lngAlcaFerro            As Long
Dim intPipeta               As Integer
Dim intTampaPressao         As Integer
Dim intBatoqueRetratil      As Integer
Dim intBatoquePlastico      As Integer
Dim intTampaVisor           As Integer
Dim arrVEDANTE              As Variant
Dim arrTAMPAPRESS           As Variant
Dim arrBATRETRATI           As Variant
Dim arrBATPLASTIC           As Variant
Dim arrTAMPAVISOR           As Variant
Dim intTipoPalletGranel     As Integer
Dim lngQtdePalletGranel     As Double
Dim lngFechSoldaAgrafado    As Long
Dim intFechSA               As Integer
Dim lngFechTampaFuro        As Long
Dim intAzelha               As Integer
Dim lngVerCat               As Long
Dim lngColEsp               As Long
Dim strNOMEARQ              As String
Dim strCAMINHO              As String
Dim imgIMAGEM               As Variant
Dim intNATURALSIMNAO        As Integer
Dim intStatus               As Integer
Dim intEMITLAUDO            As Integer
Dim lngQTDEPORFOLHA         As Long
Dim arrFAMMAQ               As Variant
Dim lngQtdePassadas         As Long
Dim intALCAGALAO            As Integer
Dim intNECKIN               As Integer
Dim intFILIALPED            As Integer
Dim intDIMPADRAO            As Integer
Dim strDESENV               As String
Dim strALTURA               As String
Dim strCOMPLEMENTO          As String
Dim strCODPRODENTR          As String
Dim arrPRODENTR             As Variant
Dim intQTDCORPSPADRAOSN     As Integer
Dim intPRODNOVO             As Integer
Dim arrPEDIDOS              As Variant
Dim strIPI                  As String
Dim intPALLHETPADRAO        As Integer
Dim intFOTALTSN             As Integer
Dim strSEGUENCIA            As String
Dim arrPRODARV              As Variant

Public Property Get PRODARV() As Variant
    PRODARV = arrPRODARV
End Property

Public Property Let PRODARV(ByVal vData As Variant)
    arrPRODARV = vData
End Property



Public Property Get SEGUENCIA() As String
    SEGUENCIA = strSEGUENCIA
End Property

Public Property Let SEGUENCIA(ByVal vData As String)
    strSEGUENCIA = vData
End Property

Public Property Get FOTALTSN() As Integer
    FOTALTSN = intFOTALTSN
End Property

Public Property Let FOTALTSN(ByVal vData As Integer)
    intFOTALTSN = vData
End Property

Public Property Get PALLHETPADRAO() As Integer
    PALLHETPADRAO = intPALLHETPADRAO
End Property

Public Property Let PALLHETPADRAO(ByVal vData As Integer)
    intPALLHETPADRAO = vData
End Property

Public Property Get IPI() As String
    IPI = strIPI
End Property

Public Property Let IPI(ByVal vData As String)
    strIPI = vData
End Property

Public Property Get VERNIZEXT() As Variant
    VERNIZEXT = arrVERNIZEXT
End Property

Public Property Let VERNIZEXT(ByVal vData As Variant)
    arrVERNIZEXT = vData
End Property

Public Property Get PEDIDOS() As Variant
    PEDIDOS = arrPEDIDOS
End Property

Public Property Let PEDIDOS(ByVal vData As Variant)
    arrPEDIDOS = vData
End Property

Public Property Get PRODNOVO() As Integer
    PRODNOVO = intPRODNOVO
End Property

Public Property Let PRODNOVO(ByVal vData As Integer)
    intPRODNOVO = vData
End Property

Public Property Get QTDCORPSPADRAOSN() As Variant
    QTDCORPSPADRAOSN = intQTDCORPSPADRAOSN
End Property

Public Property Let QTDCORPSPADRAOSN(ByVal vData As Variant)
    intQTDCORPSPADRAOSN = vData
End Property

Public Property Get PRODENTR() As Variant
    PRODENTR = arrPRODENTR
End Property

Public Property Let PRODENTR(ByVal vData As Variant)
    arrPRODENTR = vData
End Property

Public Property Get CODPRODENTR() As String
    CODPRODENTR = strCODPRODENTR
End Property

Public Property Let CODPRODENTR(ByVal vData As String)
    strCODPRODENTR = vData
End Property

Public Property Get VERNIZ02() As Variant
    VERNIZ02 = arrVERNIZ02
End Property

Public Property Let VERNIZ02(ByVal vData As Variant)
    arrVERNIZ02 = vData
End Property

Public Property Get COMPLEMENTO() As String
    COMPLEMENTO = strCOMPLEMENTO
End Property

Public Property Let COMPLEMENTO(ByVal vData As String)
    strCOMPLEMENTO = vData
End Property

Public Property Get DESENV() As String
    DESENV = strDESENV
End Property

Public Property Let DESENV(ByVal vData As String)
    strDESENV = vData
End Property

Public Property Get ALTURA() As String
    ALTURA = strALTURA
End Property

Public Property Let ALTURA(ByVal vData As String)
    strALTURA = vData
End Property

Public Property Get DIMPADRAO() As Integer
    DIMPADRAO = intDIMPADRAO
End Property

Public Property Let DIMPADRAO(ByVal vData As Integer)
    intDIMPADRAO = vData
End Property

Public Property Get FILIALPED() As Integer
    FILIALPED = intFILIALPED
End Property

Public Property Let FILIALPED(ByVal vData As Integer)
    intFILIALPED = vData
End Property

Public Property Get NECKIN() As Integer
    NECKIN = intNECKIN
End Property

Public Property Let NECKIN(ByVal vData As Integer)
    intNECKIN = vData
End Property


Public Property Get ALCAGALAO() As Integer
    ALCAGALAO = intALCAGALAO
End Property

Public Property Let ALCAGALAO(ByVal vData As Integer)
    intALCAGALAO = vData
End Property


Public Property Get QTDPASSADAS() As Long
    QTDPASSADAS = lngQtdePassadas
End Property

Public Property Let QTDPASSADAS(ByVal vData As Long)
    lngQtdePassadas = vData
End Property

Public Property Get FAMMAQ() As Variant
    FAMMAQ = arrFAMMAQ
End Property

Public Property Let FAMMAQ(ByVal vData As Variant)
    arrFAMMAQ = vData
End Property

Public Property Get QTDEPORFOLHA() As Long
    QTDEPORFOLHA = lngQTDEPORFOLHA
End Property

Public Property Let QTDEPORFOLHA(ByVal vData As Long)
    lngQTDEPORFOLHA = vData
End Property

Public Property Get EMITLAUDO() As Integer
    EMITLAUDO = intEMITLAUDO
End Property

Public Property Let EMITLAUDO(ByVal vData As Integer)
    intEMITLAUDO = vData
End Property

Public Property Get STATUS() As Integer
    STATUS = intStatus
End Property

Public Property Let STATUS(ByVal vData As Integer)
    intStatus = vData
End Property

Public Property Get NATURALSIMNAO() As Integer
    NATURALSIMNAO = intNATURALSIMNAO
End Property

Public Property Let NATURALSIMNAO(ByVal vData As Integer)
    intNATURALSIMNAO = vData
End Property

Public Property Get IMAGEM() As Variant
    IMAGEM = imgIMAGEM
End Property

Public Property Let IMAGEM(ByVal vData As Variant)
    imgIMAGEM = vData
End Property

Public Property Get CAMINHO() As String
    CAMINHO = strCAMINHO
End Property

Public Property Let CAMINHO(ByVal vData As String)
    strCAMINHO = vData
End Property

Public Property Get NOMEARQ() As String
    NOMEARQ = strNOMEARQ
End Property

Public Property Let NOMEARQ(ByVal vData As String)
    strNOMEARQ = vData
End Property


Public Property Get PRODATECLIE() As Variant
    PRODATECLIE = arrPRODATECLIE
End Property
Public Property Let PRODATECLIE(ByVal vData As Variant)
    arrPRODATECLIE = vData
End Property

Public Property Get TAMPAVISOR() As Variant
    TAMPAVISOR = arrTAMPAVISOR
End Property
Public Property Let TAMPAVISOR(ByVal vData As Variant)
    arrTAMPAVISOR = vData
End Property

Public Property Get BATPLASTIC() As Variant
    BATPLASTIC = arrBATPLASTIC
End Property
Public Property Let BATPLASTIC(ByVal vData As Variant)
    arrBATPLASTIC = vData
End Property

Public Property Get BATRETRATI() As Variant
    BATRETRATI = arrBATRETRATI
End Property
Public Property Let BATRETRATI(ByVal vData As Variant)
    arrBATRETRATI = vData
End Property

Public Property Get TAMPAPRESS() As Variant
    TAMPAPRESS = arrTAMPAPRESS
End Property
Public Property Let TAMPAPRESS(ByVal vData As Variant)
    arrTAMPAPRESS = vData
End Property

Public Property Get ColEsp() As Long
    ColEsp = lngColEsp
End Property
Public Property Let ColEsp(ByVal vData As Long)
    lngColEsp = vData
End Property

Public Property Get VerCat() As Long
    VerCat = lngVerCat
End Property
Public Property Let VerCat(ByVal vData As Long)
    lngVerCat = vData
End Property

Public Property Get Azelha() As Integer
    Azelha = intAzelha
End Property
Public Property Let Azelha(ByVal vData As Integer)
    intAzelha = vData
End Property

Public Property Get VERNIZACAB() As Variant
    VERNIZACAB = arrVERNIZACAB
End Property
Public Property Let VERNIZACAB(ByVal vData As Variant)
    arrVERNIZACAB = vData
End Property

Public Property Get QtdePalletGranel() As Double
    QtdePalletGranel = lngQtdePalletGranel
End Property
Public Property Let QtdePalletGranel(ByVal vData As Double)
    lngQtdePalletGranel = vData
End Property

Public Property Get TipoPalletGranel() As Integer
    TipoPalletGranel = intTipoPalletGranel
End Property
Public Property Let TipoPalletGranel(ByVal vData As Integer)
    intTipoPalletGranel = vData
End Property

Public Property Get VEDANTE() As Variant
    VEDANTE = arrVEDANTE
End Property
Public Property Let VEDANTE(ByVal vData As Variant)
    arrVEDANTE = vData
End Property

Public Property Get TAMPAVIS() As Integer
    TAMPAVIS = intTampaVisor
End Property
Public Property Let TAMPAVIS(ByVal vData As Integer)
    intTampaVisor = vData
End Property

Public Property Get BatoquePlastico() As Integer
    BatoquePlastico = intBatoquePlastico
End Property
Public Property Let BatoquePlastico(ByVal vData As Integer)
    intBatoquePlastico = vData
End Property

Public Property Get BatoqueRetratil() As Integer
    BatoqueRetratil = intBatoqueRetratil
End Property
Public Property Let BatoqueRetratil(ByVal vData As Integer)
    intBatoqueRetratil = vData
End Property

Public Property Get TampaPressao() As Integer
    TampaPressao = intTampaPressao
End Property
Public Property Let TampaPressao(ByVal vData As Integer)
    intTampaPressao = vData
End Property

Public Property Get FechSA() As Integer
    FechSA = intFechSA
End Property
Public Property Let FechSA(ByVal vData As Integer)
    intFechSA = vData
End Property

Public Property Get FechSoldaAgrafado() As Long
    FechSoldaAgrafado = lngFechSoldaAgrafado
End Property
Public Property Let FechSoldaAgrafado(ByVal vData As Long)
    lngFechSoldaAgrafado = vData
End Property

Public Property Get Pipeta() As Integer
    Pipeta = intPipeta
End Property
Public Property Let Pipeta(ByVal vData As Integer)
    intPipeta = vData
End Property

Public Property Get AlcaFerro() As Long
    AlcaFerro = lngAlcaFerro
End Property
Public Property Let AlcaFerro(ByVal vData As Long)
    lngAlcaFerro = vData
End Property

Public Property Get AlcaPlastica() As Long
    AlcaPlastica = lngAlcaPlastica
End Property
Public Property Let AlcaPlastica(ByVal vData As Long)
    lngAlcaPlastica = vData
End Property

Public Property Get RevestArgola2() As Long
    RevestArgola2 = lngRevestArgola2
End Property
Public Property Let RevestArgola2(ByVal vData As Long)
    lngRevestArgola2 = vData
End Property

Public Property Get RevestArgola() As Long
    RevestArgola = lngRevestArgola
End Property
Public Property Let RevestArgola(ByVal vData As Long)
    lngRevestArgola = vData
End Property

Public Property Get RevestFundo2() As Long
    RevestFundo2 = lngRevestFundo2
End Property
Public Property Let RevestFundo2(ByVal vData As Long)
    lngRevestFundo2 = vData
End Property

Public Property Get RevestFundo() As Long
    RevestFundo = lngRevestFundo
End Property
Public Property Let RevestFundo(ByVal vData As Long)
    lngRevestFundo = vData
End Property

Public Property Get RevestTampa2() As Long
    RevestTampa2 = lngRevestTampa2
End Property
Public Property Let RevestTampa2(ByVal vData As Long)
    lngRevestTampa2 = vData
End Property

Public Property Get RevestTampa() As Long
    RevestTampa = lngRevestTampa
End Property
Public Property Let RevestTampa(ByVal vData As Long)
    lngRevestTampa = vData
End Property

Public Property Get RevestCorpo2() As Long
    RevestCorpo2 = lngRevestCorpo2
End Property
Public Property Let RevestCorpo2(ByVal vData As Long)
    lngRevestCorpo2 = vData
End Property

Public Property Get RevestCorpo() As Long
    RevestCorpo = lngRevestCorpo
End Property
Public Property Let RevestCorpo(ByVal vData As Long)
    lngRevestCorpo = vData
End Property

Public Property Get EspessArgola() As Long
    EspessArgola = lngEspessArgola
End Property
Public Property Let EspessArgola(ByVal vData As Long)
    lngEspessArgola = vData
End Property

Public Property Get EspessFundo() As Long
    EspessFundo = lngEspessFundo
End Property
Public Property Let EspessFundo(ByVal vData As Long)
    lngEspessFundo = vData
End Property

Public Property Get EspessTampa() As Long
    EspessTampa = lngEspessTampa
End Property
Public Property Let EspessTampa(ByVal vData As Long)
    lngEspessTampa = vData
End Property

Public Property Get EspessCorpo() As Long
    EspessCorpo = lngEspessCorpo
End Property
Public Property Let EspessCorpo(ByVal vData As Long)
    lngEspessCorpo = vData
End Property

Public Property Get FechTampaFuro() As Long
    FechTampaFuro = lngFechTampaFuro
End Property
Public Property Let FechTampaFuro(ByVal vData As Long)
    lngFechTampaFuro = vData
End Property

Public Property Get VernCorpo() As Long
    VernCorpo = lngVernCorpo
End Property
Public Property Let VernCorpo(ByVal vData As Long)
    lngVernCorpo = vData
End Property

Public Property Get VernArgola() As Long
    VernArgola = lngVernArgola
End Property
Public Property Let VernArgola(ByVal vData As Long)
    lngVernArgola = vData
End Property

Public Property Get VernFundo() As Long
    VernFundo = lngVernFundo
End Property
Public Property Let VernFundo(ByVal vData As Long)
    lngVernFundo = vData
End Property

Public Property Get VernTampa() As Long
    VernTampa = lngVernTampa
End Property
Public Property Let VernTampa(ByVal vData As Long)
    lngVernTampa = vData
End Property

Public Property Get CORES() As Variant
    CORES = arrCORES
End Property
Public Property Let CORES(ByVal vData As Variant)
    arrCORES = vData
End Property

Public Property Get VERNIZ() As Variant
    VERNIZ = arrVERNIZ
End Property
Public Property Let VERNIZ(ByVal vData As Variant)
    arrVERNIZ = vData
End Property
Public Property Get ESMALTE() As Variant
    ESMALTE = arrESMALTE
End Property
Public Property Let ESMALTE(ByVal vData As Variant)
    arrESMALTE = vData
End Property

Public Property Get IDProduto() As Variant
    IDProduto = lngIDProduto
End Property
Public Property Let IDProduto(ByVal vData As Variant)
    lngIDProduto = vData
End Property

Public Property Get CodLinProd() As Variant
    CodLinProd = lngCodLinProd
End Property
Public Property Let CodLinProd(ByVal vData As Variant)
    lngCodLinProd = vData
End Property

Public Property Get CodClie() As Variant
    CodClie = lngCodClie
End Property
Public Property Let CodClie(ByVal vData As Variant)
    lngCodClie = vData
End Property

Public Property Get CodRotulo() As Variant
    CodRotulo = lngCodRotulo
End Property
Public Property Let CodRotulo(ByVal vData As Variant)
    lngCodRotulo = vData
End Property

Public Property Get DigVerif() As Variant
    DigVerif = lngDigVerif
End Property
Public Property Let DigVerif(ByVal vData As Variant)
    lngDigVerif = vData
End Property

Public Property Get UNIDCONV() As Variant
    UNIDCONV = arrUNIDCONV
End Property
Public Property Let UNIDCONV(ByVal vData As Variant)
    arrUNIDCONV = vData
End Property

Public Property Get GRAMATURA2() As Currency
    GRAMATURA2 = curGRAMATURA2
End Property
Public Property Let GRAMATURA2(ByVal vData As Currency)
    curGRAMATURA2 = vData
End Property

Public Property Get LARGURA() As Currency
    LARGURA = curLargura
End Property
Public Property Let LARGURA(ByVal vData As Currency)
    curLargura = vData
End Property

Public Property Get METROS() As Currency
    METROS = curMetros
End Property
Public Property Let METROS(ByVal vData As Currency)
    curMetros = vData
End Property

Public Property Get CUBAGEN() As Currency
    CUBAGEN = curCUBAGEN
End Property
Public Property Let CUBAGEN(ByVal vData As Currency)
    curCUBAGEN = vData
End Property

Public Property Get PORCCORR() As Variant
    PORCCORR = curPORCCORR
End Property
Public Property Let PORCCORR(ByVal vData As Variant)
    curPORCCORR = vData
End Property

Public Property Get PARPROD() As Variant
    PARPROD = arrPARPROD
End Property
Public Property Let PARPROD(ByVal vData As Variant)
    arrPARPROD = vData
End Property

Public Property Get CORRPRODPADR() As Variant
    CORRPRODPADR = arrCORRPRODPADR
End Property
Public Property Let CORRPRODPADR(ByVal vData As Variant)
    arrCORRPRODPADR = vData
End Property

Public Property Get FAMMAQUINA() As Long
    FAMMAQUINA = lngFAMMAQUINA
End Property
Public Property Let FAMMAQUINA(ByVal vData As Long)
    lngFAMMAQUINA = vData
End Property

Public Property Get CORRELACAO() As Currency
    CORRELACAO = curCORRELACAO
End Property
Public Property Let CORRELACAO(ByVal vData As Currency)
    curCORRELACAO = vData
End Property

Public Property Get PADRAO() As Integer
    PADRAO = intPADRAO
End Property
Public Property Let PADRAO(ByVal vData As Integer)
    intPADRAO = vData
End Property

Public Property Get DISTANCIA() As Currency
    DISTANCIA = curDISTANCIA
End Property
Public Property Let DISTANCIA(ByVal vData As Currency)
    curDISTANCIA = vData
End Property

Public Property Get PRODATRINSP() As Variant
    PRODATRINSP = arrPRODATRINSP
End Property
Public Property Let PRODATRINSP(ByVal vData As Variant)
    arrPRODATRINSP = vData
End Property

Public Property Get PRODATRREC() As Variant
    PRODATRREC = arrPRODATRREC
End Property
Public Property Let PRODATRREC(ByVal vData As Variant)
    arrPRODATRREC = vData
End Property

Public Property Get PRODATRPROC() As Variant
    PRODATRPROC = arrPRODATRPROC
End Property
Public Property Let PRODATRPROC(ByVal vData As Variant)
    arrPRODATRPROC = vData
End Property

Public Property Get PRODVAPROC() As Variant
    PRODVAPROC = arrPRODVAPROC
End Property
Public Property Let PRODVAPROC(ByVal vData As Variant)
    arrPRODVAPROC = vData
End Property

Public Property Get PRODVAINSP() As Variant
    PRODVAINSP = arrPRODVAINSP
End Property
Public Property Let PRODVAINSP(ByVal vData As Variant)
    arrPRODVAINSP = vData
End Property

Public Property Get PRODVAREC() As Variant
    PRODVAREC = arrPRODVAREC
End Property
Public Property Let PRODVAREC(ByVal vData As Variant)
    arrPRODVAREC = vData
End Property

Public Property Get PRODTIPORCA() As Variant
    PRODTIPORCA = arrPRODTIPORCA
End Property
Public Property Let PRODTIPORCA(ByVal vData As Variant)
    arrPRODTIPORCA = vData
End Property

Public Property Get PRODFOR() As Variant
    PRODFOR = arrPRODFOR
End Property
Public Property Let PRODFOR(ByVal vData As Variant)
    arrPRODFOR = vData
End Property

Public Property Get CODEAN() As String
    CODEAN = strCODEAN
End Property
Public Property Let CODEAN(ByVal vData As String)
    strCODEAN = vData
End Property
Public Property Get PRATELEIRA() As String
    PRATELEIRA = strPRATELEIRA
End Property
Public Property Let PRATELEIRA(ByVal vData As String)
    strPRATELEIRA = vData
End Property
Public Property Get box() As String
    box = strBOX
End Property
Public Property Let box(ByVal vData As String)
    strBOX = vData
End Property
Public Property Get RUA() As String
    RUA = strRUA
End Property
Public Property Let RUA(ByVal vData As String)
    strRUA = vData
End Property
Public Property Get CODPROFORNEC() As String
    CODPROFORNEC = strCODPROFORNEC
End Property
Public Property Let CODPROFORNEC(ByVal vData As String)
    strCODPROFORNEC = vData
End Property
Public Property Get TABCONV() As Long
    TABCONV = lngTABCONV
End Property

Public Property Let TABCONV(ByVal vData As Long)
    lngTABCONV = vData
End Property
Public Property Get PESOUNIT() As Currency
    PESOUNIT = curPESOUNIT
End Property
Public Property Let PESOUNIT(ByVal vData As Currency)
    curPESOUNIT = vData
End Property

Public Property Get ESTILOPROD() As Integer
    ESTILOPROD = intESTILOPROD
End Property

Public Property Let ESTILOPROD(ByVal vData As Integer)
    intESTILOPROD = vData
End Property
Public Property Get PRODUTOTIPO() As Integer
    PRODUTOTIPO = intPRODUTOTIPO
End Property
Public Property Let PRODUTOTIPO(ByVal vData As Integer)
    intPRODUTOTIPO = vData
End Property
Public Property Get PRCCUSTO() As Currency
    PRCCUSTO = curPRCCUSTO
End Property
Public Property Let PRCCUSTO(ByVal vData As Currency)
    curPRCCUSTO = vData
End Property
Public Property Get PORCACRES() As Currency
    PORCACRES = curPORCACRES
End Property
Public Property Let PORCACRES(ByVal vData As Currency)
    curPORCACRES = vData
End Property
Public Property Get Processos() As Variant
    Processos = arrPROCESSOS
End Property
Public Property Let Processos(ByVal vData As Variant)
    arrPROCESSOS = vData
End Property
Public Property Get CODSUBGPROD() As Integer
    CODSUBGPROD = intCODSUBGPROD
End Property
Public Property Let CODSUBGPROD(ByVal vData As Integer)
    intCODSUBGPROD = vData
End Property
Public Property Get CODGRUPPROD() As Integer
    CODGRUPPROD = intCODGRUPPROD
End Property
Public Property Let CODGRUPPROD(ByVal vData As Integer)
    intCODGRUPPROD = vData
End Property
Public Property Get PROCEDEN() As String
    PROCEDEN = strPROCEDEN
End Property
Public Property Let PROCEDEN(ByVal vData As String)
    strPROCEDEN = vData
End Property
Public Property Get DESCEQUIP() As String
    DESCEQUIP = strDESCEQUIP
End Property
Public Property Let DESCEQUIP(ByVal vData As String)
    strDESCEQUIP = vData
End Property
Public Property Get ESPTECPROD() As Variant
    ESPTECPROD = arrESPTECPROD
End Property
Public Property Let ESPTECPROD(ByVal vData As Variant)
    arrESPTECPROD = vData
End Property
Public Property Get PRCMED() As Double
    PRCMED = dblPRCMED
End Property
Public Property Let PRCMED(ByVal vData As Double)
    dblPRCMED = vData
End Property
Public Property Get PRCPROD() As Double
    PRCPROD = dblPRCPROD
End Property
Public Property Let PRCPROD(ByVal vData As Double)
    dblPRCPROD = vData
End Property
Public Property Get ATAUTOMSN() As Integer
    ATAUTOMSN = intATAUTOMSN
End Property
Public Property Let ATAUTOMSN(ByVal vData As Integer)
    intATAUTOMSN = vData
End Property
Public Property Get CLASFISC() As Long
    CLASFISC = lngCLASFISC
End Property
Public Property Let CLASFISC(ByVal vData As Long)
    lngCLASFISC = vData
End Property
Public Property Get DataUltMov() As Date
    DataUltMov = dtDataUltMov
End Property
Public Property Let DataUltMov(ByVal vData As Date)
    dtDataUltMov = vData
End Property
Public Property Get CODLCTO() As Long
    CODLCTO = lngCODLCTO
End Property
Public Property Let CODLCTO(ByVal vData As Long)
    lngCODLCTO = vData
End Property
Public Property Get ProdLST() As Variant
    ProdLST = arrProdLST
End Property
Public Property Let ProdLST(ByVal vData As Variant)
    arrProdLST = vData
End Property
Public Property Get CodigoProd() As String
    CodigoProd = strCodigoProd
End Property
Public Property Let CodigoProd(ByVal vData As String)
    strCodigoProd = vData
End Property
Public Property Get CodigoProdANT() As String
    CodigoProdANT = strCodigoProdANT
End Property
Public Property Let CodigoProdANT(ByVal vData As String)
    strCodigoProdANT = vData
End Property
Public Property Get DescriProd() As String
    DescriProd = strDescriProd
End Property
Public Property Let DescriProd(ByVal vData As String)
    strDescriProd = vData
End Property
Public Property Get EspProduto() As Integer
    EspProduto = intEspProduto
End Property
Public Property Let EspProduto(ByVal vData As Integer)
    intEspProduto = vData
End Property
Public Property Get TIPPRODUTO() As Integer
    TIPPRODUTO = intTIPPRODUTO
End Property
Public Property Let TIPPRODUTO(ByVal vData As Integer)
    intTIPPRODUTO = vData
End Property
Public Property Get Unidade() As Integer
    Unidade = intUnidade
End Property
Public Property Let Unidade(ByVal vData As Integer)
    intUnidade = vData
End Property
Public Property Get Saldo() As Double
    Saldo = dblSaldo
End Property
Public Property Let Saldo(ByVal vData As Double)
    dblSaldo = vData
End Property
Public Property Get EstMin() As Double
    EstMin = dblEstMin
End Property
Public Property Let EstMin(ByVal vData As Double)
    dblEstMin = vData
End Property
Public Property Get DataCadast() As String
    DataCadast = dtDataCadast
End Property
Public Property Let DataCadast(ByVal vData As String)
    dtDataCadast = vData
End Property
Public Property Get FERRAMENTA() As Variant
    FERRAMENTA = arrFERRAMENTA
End Property
Public Property Let FERRAMENTA(ByVal vData As Variant)
    arrFERRAMENTA = vData
End Property
Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADPRODP.cCaminho = cCaminho
    frmCADPRODP.Linha = Linha
    frmCADPRODP.FILIAL = iFilial
    frmCADPRODP.strAcesso = strAcesso
    frmCADPRODP.lngCodUsuario = lngIDUsuario
    frmCADPRODP.Show vbModal
End Sub

Public Sub cConnectNovo(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String)
    frmCADPROD.cCaminho = cCaminho
    frmCADPROD.Linha = Linha
    frmCADPROD.FILIAL = iFilial
    frmCADPROD.strAcesso = strAcesso
    frmCADPROD.strUSUARIO = strUSUARIO
    frmCADPROD.cTipOper = "I"
    frmCADPROD.Show vbModal
End Sub

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property
Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property
Public Function Pesq_CadProduto() As Boolean
   Pesq_CadProduto = False
   
   If BREC.State = 1 Then BREC.Close
   
   sSql = ""
   sSql = "Select " & vbCrLf
   sSql = sSql + "       *     " & vbCrLf
   sSql = sSql + "  From " & vbCrLf
   sSql = sSql + "       SGI_CADPRODUTO " & vbCrLf
   sSql = sSql + " Where " & vbCrLf
   sSql = sSql + "       SGI_FILIAL = " & intFilial & vbCrLf
   BREC.Open sSql, adoBanco_Dados
   If Not BREC.EOF Then Pesq_CadProduto = True
   BREC.Close
End Function
Public Sub PreenchComboEspecie(Combo As Variant, intCODSUBGRU As Integer)
    Combo.Clear
    sSql = "Select " & vbCrLf
    sSql = sSql & "       ESPECI.SGI_CODIGO    " & vbCrLf
    sSql = sSql & "      ,ESPECI.SGI_DESCRICAO " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_SUBGRUPRODITEN SUBGRP " & vbCrLf
    sSql = sSql & "      ,SGI_CADESPPROD     ESPECI " & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      SUBGRP.SGI_CODIGO = " & intCODSUBGRU & vbCrLf
    sSql = sSql & "  And SUBGRP.SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "  And ESPECI.SGI_FILIAL = SUBGRP.SGI_FILIAL " & vbCrLf
    sSql = sSql & "  And ESPECI.SGI_CODIGO = SUBGRP.SGI_CODESPECIE "
    sSql = sSql & "Order by ESPECI.SGI_CODIGO"
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
    If Combo.ListCount > 0 Then Combo.ListIndex = 0
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
    If Combo.ListCount > 0 Then Combo.ListIndex = 0
    BREC.Close
End Sub
Public Function GRAVA(strOPER As String) As Boolean
On Error GoTo err_Trans
     
     GRAVA = False
     
     Dim i            As Integer
     Dim intCodRotulo As Integer
     Dim strValor     As String
     Dim strEMPRESA   As String
     
     '' Inicia transação
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Then
     
        '' Grava Produto
        sSql = "Insert Into SGI_CADPRODUTO (SGI_FILIAL" & vbCrLf
        sSql = sSql & "                    ,SGI_IDPRODUTO" & vbCrLf
        sSql = sSql & "                    ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                    ,SGI_DESCRICAO" & vbCrLf
        sSql = sSql & "                    ,SGI_CODESPECIE" & vbCrLf
        sSql = sSql & "                    ,SGI_CODTIPO" & vbCrLf
        sSql = sSql & "                    ,SGI_UNIDMEDIDA" & vbCrLf
        sSql = sSql & "                    ,SGI_CODCLASS" & vbCrLf
        sSql = sSql & "                    ,SGI_SALDO" & vbCrLf
        sSql = sSql & "                    ,SGI_ESTMINIMO" & vbCrLf
        sSql = sSql & "                    ,SGI_DTCADASTRO" & vbCrLf
        sSql = sSql & "                    ,SGI_PRECOPROD" & vbCrLf
        sSql = sSql & "                    ,SGI_PRECOMED" & vbCrLf
        sSql = sSql & "                    ,SGI_ATUAUTOM" & vbCrLf
        sSql = sSql & "                    ,SGI_DESCPROD" & vbCrLf
        sSql = sSql & "                    ,SGI_PROCEDENCIA" & vbCrLf
        sSql = sSql & "                    ,SGI_CODGPROD" & vbCrLf
        sSql = sSql & "                    ,SGI_CODSUBGPROD" & vbCrLf
        sSql = sSql & "                    ,SGI_PROCACRES" & vbCrLf
        sSql = sSql & "                    ,SGI_PRCCUSTO" & vbCrLf
        sSql = sSql & "                    ,SGI_PRODUTOTIPO" & vbCrLf
        sSql = sSql & "                    ,SGI_PRODUTOESTILO" & vbCrLf
        sSql = sSql & "                    ,SGI_PESOUNIT" & vbCrLf
        sSql = sSql & "                    ,SGI_CODTABCONV" & vbCrLf
        sSql = sSql & "                    ,SGI_CODPRODFORN" & vbCrLf
        sSql = sSql & "                    ,SGI_RUA" & vbCrLf
        sSql = sSql & "                    ,SGI_BOX" & vbCrLf
        sSql = sSql & "                    ,SGI_PRATELEIRA" & vbCrLf
        sSql = sSql & "                    ,SGI_CODEAN" & vbCrLf
        sSql = sSql & "                    ,SGI_DISTESTOQUE" & vbCrLf
        sSql = sSql & "                    ,SGI_PADRAO" & vbCrLf
        sSql = sSql & "                    ,SGI_CADFAMMAQ" & vbCrLf
        sSql = sSql & "                    ,SGI_CORRPRODPADR" & vbCrLf
        sSql = sSql & "                    ,SGI_PORCCORR" & vbCrLf
        sSql = sSql & "                    ,SGI_CUBAGEN" & vbCrLf
        sSql = sSql & "                    ,SGI_METROS" & vbCrLf
        sSql = sSql & "                    ,SGI_LARGURA" & vbCrLf
        sSql = sSql & "                    ,SGI_GRAMATURA2" & vbCrLf
        sSql = sSql & "                    ,SGI_CODLINPROD" & vbCrLf
        sSql = sSql & "                    ,SGI_CODCLIE" & vbCrLf
        sSql = sSql & "                    ,SGI_CODROTULO" & vbCrLf
        sSql = sSql & "                    ,SGI_DIGVERIF" & vbCrLf
        
        sSql = sSql & "                    ,SGI_VernCorpo" & vbCrLf
        sSql = sSql & "                    ,SGI_VernTampa" & vbCrLf
        sSql = sSql & "                    ,SGI_VernFundo" & vbCrLf
        sSql = sSql & "                    ,SGI_VernArgola" & vbCrLf
        sSql = sSql & "                    ,SGI_FechTampaFuro" & vbCrLf
        
        sSql = sSql & "                    ,SGI_CorpoEspess" & vbCrLf
        sSql = sSql & "                    ,SGI_TampaEspess" & vbCrLf
        sSql = sSql & "                    ,SGI_FundoEspess" & vbCrLf
        sSql = sSql & "                    ,SGI_ArgolaEspess" & vbCrLf
        
        sSql = sSql & "                    ,SGI_CorpoRevest" & vbCrLf
        sSql = sSql & "                    ,SGI_TampaRevest" & vbCrLf
        sSql = sSql & "                    ,SGI_FundoRevest" & vbCrLf
        sSql = sSql & "                    ,SGI_ArgolaRevest" & vbCrLf
        
        sSql = sSql & "                    ,SGI_CorpoRevest2" & vbCrLf
        sSql = sSql & "                    ,SGI_TampaRevest2" & vbCrLf
        sSql = sSql & "                    ,SGI_FundoRevest2" & vbCrLf
        sSql = sSql & "                    ,SGI_ArgolaRevest2" & vbCrLf
        
        sSql = sSql & "                    ,SGI_AlcaPlastica" & vbCrLf
        sSql = sSql & "                    ,SGI_AlcaFerro" & vbCrLf
        sSql = sSql & "                    ,SGI_Pipeta" & vbCrLf
        
        sSql = sSql & "                    ,SGI_FechSoldaAgrafado" & vbCrLf
        sSql = sSql & "                    ,SGI_FechSA" & vbCrLf
        
        sSql = sSql & "                    ,SGI_FechamTampaPress" & vbCrLf
        sSql = sSql & "                    ,SGI_FechamBatoqRetra" & vbCrLf
        sSql = sSql & "                    ,SGI_FechamBatoqPlast" & vbCrLf
        sSql = sSql & "                    ,SGI_FechamTampaVisor" & vbCrLf
        
        sSql = sSql & "                    ,SGI_TIPOEMPILHAMENTO" & vbCrLf
        sSql = sSql & "                    ,SGI_QTDEEMPILHAMENTO" & vbCrLf
        sSql = sSql & "                    ,SGI_AZELHA" & vbCrLf
        sSql = sSql & "                    ,SGI_VERCAT" & vbCrLf
        sSql = sSql & "                    ,SGI_COLESP" & vbCrLf
        sSql = sSql & "                    ,SGI_TEMARVORE" & vbCrLf
        sSql = sSql & "                    ,SGI_NATURALSN" & vbCrLf
        sSql = sSql & "                    ,SGI_STATUS" & vbCrLf
        sSql = sSql & "                    ,SGI_EMITLAUDO" & vbCrLf
        sSql = sSql & "                    ,SGI_QTDEPORFOLHA" & vbCrLf
        sSql = sSql & "                    ,SGI_QTDEPASSADA" & vbCrLf
        sSql = sSql & "                    ,SGI_ALCAGALAO" & vbCrLf
        sSql = sSql & "                    ,SGI_NECKIN" & vbCrLf
        sSql = sSql & "                    ,SGI_FILIALPED" & vbCrLf
        sSql = sSql & "                    ,SGI_ARQUIVOIMG" & vbCrLf
        sSql = sSql & "                    ,SGI_DIMPADRAO" & vbCrLf
        sSql = sSql & "                    ,SGI_DESENV" & vbCrLf
        sSql = sSql & "                    ,SGI_ALTURA" & vbCrLf
        sSql = sSql & "                    ,SGI_COMPLEMENTO" & vbCrLf
        sSql = sSql & "                    ,SGI_CODPRODENTR" & vbCrLf
        sSql = sSql & "                    ,SGI_PRODNOVO" & vbCrLf
        sSql = sSql & "                    ,SGI_IPI" & vbCrLf
        sSql = sSql & "                    ,SGI_USARPADRPALSN" & vbCrLf
        sSql = sSql & "                    ,SGI_ALTFOTOLIT" & vbCrLf
        sSql = sSql & "                    ,SGI_SEGUENCIA" & vbCrLf
        
        sSql = sSql & "                    ) Values (" & vbCrLf
        sSql = sSql & "                     " & intFilial & vbCrLf
        sSql = sSql & "                    ," & lngIDProduto & vbCrLf
        If Len(Trim(strCodigoProd)) > 0 Then
            sSql = sSql & "                    ,'" & strCodigoProd & "'" & vbCrLf
        Else
            sSql = sSql & "                    ,Null" & vbCrLf
        End If
        sSql = sSql & "                    ,'" & strDescriProd & "'" & vbCrLf
        sSql = sSql & "                    ," & intEspProduto & vbCrLf
        sSql = sSql & "                    ," & intTIPPRODUTO & vbCrLf
        sSql = sSql & "                    ," & intUnidade & vbCrLf
        sSql = sSql & "                    ," & lngCLASFISC & vbCrLf
                
        strValor = Replace(Format(dblSaldo, "#,###0.000"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                           ," & Trim(strValor) & vbCrLf
        
        strValor = Replace(Format(dblEstMin, "#,###0.000"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                           ," & Trim(strValor) & vbCrLf
        
        sSql = sSql & "                           ," & dtDataCadast & vbCrLf
        
        strValor = Replace(Format(dblPRCPROD, "#,##0.00"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                           ," & Trim(strValor) & vbCrLf
        
        strValor = Replace(Format(dblPRCMED, "#,##0.00"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                           ," & Trim(strValor) & vbCrLf
        
        sSql = sSql & "                           ," & intATAUTOMSN & vbCrLf
        
        sSql = sSql & "                           ,'" & strDESCEQUIP & "'" & vbCrLf
        
        sSql = sSql & "                           ,'" & strPROCEDEN & "'" & vbCrLf
        sSql = sSql & "                           ," & intCODGRUPPROD & vbCrLf
        sSql = sSql & "                           ," & intCODSUBGPROD & vbCrLf
        
        strValor = Replace(Format(curPORCACRES, "#,##0.00"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                           ," & strValor & vbCrLf
        
        strValor = Replace(Format(curPRCCUSTO, "#,##0.00"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                           ," & strValor & vbCrLf
        
        sSql = sSql & "                           ," & intPRODUTOTIPO & vbCrLf
        sSql = sSql & "                           ," & intESTILOPROD & vbCrLf
        
        strValor = Replace(Format(curPESOUNIT, "#,###0.000"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                           ," & strValor & vbCrLf
        
        If lngTABCONV > 0 Then
           sSql = sSql & "                           ," & lngTABCONV & vbCrLf
        Else
           sSql = sSql & "                           ,Null" & vbCrLf
        End If
        
        sSql = sSql & "                           ,'" & strCODPROFORNEC & "'" & vbCrLf
        
        sSql = sSql & "                           ,'" & strRUA & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strBOX & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strPRATELEIRA & "'" & vbCrLf
        sSql = sSql & "                           ,'" & strCODEAN & "'" & vbCrLf
        
        strValor = Replace(Format(curDISTANCIA, "#,###0.000"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                           ," & Trim(strValor) & vbCrLf
        sSql = sSql & "                           ," & intPADRAO & vbCrLf
        
        If lngFAMMAQUINA > 0 Then
            sSql = sSql & "                           ," & lngFAMMAQUINA & vbCrLf
        Else
            sSql = sSql & "                           ,Null" & vbCrLf
        End If
        
        strValor = Replace(Format(curCORRELACAO, "#,###0.000"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                           ," & strValor & vbCrLf
        
        strValor = Replace(Format(curPORCCORR, "#,##0.00"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                           ," & strValor & vbCrLf
                
        strValor = Replace(Format(curCUBAGEN, "#,###0.000"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                           ," & strValor & vbCrLf
        
        strValor = Replace(Format(curMetros, "#,###0.000"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                           ," & strValor & vbCrLf
        
        strValor = Replace(Format(curLargura, "#,###0.000"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                           ," & strValor & vbCrLf
        
        strValor = Replace(Format(curGRAMATURA2, "#,###0.000"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                           ," & strValor & vbCrLf
        
        If lngCodLinProd > 0 Then
            sSql = sSql & "                           ," & lngCodLinProd & vbCrLf
        Else
            sSql = sSql & "                           ,Null" & vbCrLf
        End If
        
        If lngCodClie > 0 Then
            sSql = sSql & "                           ," & lngCodClie & vbCrLf
        Else
            sSql = sSql & "                           ,Null" & vbCrLf
        End If
        
        ''If lngCodRotulo > 0 Then
        ''    sSql = sSql & "                           ," & lngCodRotulo & vbCrLf
        ''Else
           sSql = sSql & "                           ,Null" & vbCrLf
        ''End If
        
        If lngDigVerif > 0 Then
            sSql = sSql & "                           ," & lngDigVerif & vbCrLf
        Else
            sSql = sSql & "                           ,Null" & vbCrLf
        End If
        
        
        If lngVernCorpo > 0 Then
            sSql = sSql & "                           ," & lngVernCorpo & vbCrLf
        Else
            sSql = sSql & "                           ,Null" & vbCrLf
        End If
        If lngVernTampa > 0 Then
            sSql = sSql & "                           ," & lngVernTampa & vbCrLf
        Else
            sSql = sSql & "                           ,Null" & vbCrLf
        End If
        If lngVernFundo > 0 Then
            sSql = sSql & "                           ," & lngVernFundo & vbCrLf
        Else
            sSql = sSql & "                           ,Null" & vbCrLf
        End If
        If lngVernArgola > 0 Then
            sSql = sSql & "                           ," & lngVernArgola & vbCrLf
        Else
            sSql = sSql & "                           ,Null" & vbCrLf
        End If
        
        If lngFechTampaFuro > -1 Then
            sSql = sSql & "                           ," & lngFechTampaFuro & vbCrLf
        Else
            sSql = sSql & "                           ,Null" & vbCrLf
        End If
        
        
        strValor = Replace(Format(lngEspessCorpo, "#,###0.000"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                           ," & Trim(strValor) & vbCrLf
        
        strValor = Replace(Format(lngEspessTampa, "#,###0.000"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                           ," & Trim(strValor) & vbCrLf

        strValor = Replace(Format(lngEspessFundo, "#,###0.000"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                           ," & Trim(strValor) & vbCrLf

        strValor = Replace(Format(lngEspessArgola, "#,###0.000"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                           ," & Trim(strValor) & vbCrLf
        
        strValor = Replace(Format(lngRevestCorpo, "#,###0.000"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                           ," & Trim(strValor) & vbCrLf
        
        strValor = Replace(Format(lngRevestTampa, "#,###0.000"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                           ," & Trim(strValor) & vbCrLf
        
        strValor = Replace(Format(lngRevestFundo, "#,###0.000"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                           ," & Trim(strValor) & vbCrLf
        
        strValor = Replace(Format(lngRevestArgola, "#,###0.000"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                           ," & Trim(strValor) & vbCrLf
        
        If lngRevestCorpo2 > 0 Then
            sSql = sSql & "                           ," & lngRevestCorpo2 & vbCrLf
        Else
            sSql = sSql & "                           ,Null" & vbCrLf
        End If
        If lngRevestTampa2 > 0 Then
            sSql = sSql & "                           ," & lngRevestTampa2 & vbCrLf
        Else
            sSql = sSql & "                           ,Null" & vbCrLf
        End If
        If lngRevestFundo2 > 0 Then
            sSql = sSql & "                           ," & lngRevestFundo2 & vbCrLf
        Else
            sSql = sSql & "                           ,Null" & vbCrLf
        End If
        If lngRevestArgola2 > 0 Then
            sSql = sSql & "                           ," & lngRevestArgola2 & vbCrLf
        Else
            sSql = sSql & "                           ,Null" & vbCrLf
        End If
        
        
        If lngAlcaPlastica > 0 Then
            sSql = sSql & "                           ," & lngAlcaPlastica & vbCrLf
        Else
            sSql = sSql & "                           ,Null" & vbCrLf
        End If
        If lngAlcaFerro > 0 Then
            sSql = sSql & "                           ," & lngAlcaFerro & vbCrLf
        Else
            sSql = sSql & "                           ,Null" & vbCrLf
        End If
        sSql = sSql & "                           ," & intPipeta & vbCrLf
        
        If lngFechSoldaAgrafado > -1 Then
            sSql = sSql & "                           ," & lngFechSoldaAgrafado & vbCrLf
        Else
            sSql = sSql & "                           ,Null" & vbCrLf
        End If
        sSql = sSql & "                           ," & intFechSA & vbCrLf
        
        sSql = sSql & "                           ,Null" & vbCrLf
        sSql = sSql & "                           ,Null" & vbCrLf
        sSql = sSql & "                           ,Null" & vbCrLf
        sSql = sSql & "                           ,Null" & vbCrLf
        
        sSql = sSql & "                           ," & intTipoPalletGranel & vbCrLf
        strValor = Replace(Format(lngQtdePalletGranel, "#,###0.000"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                           ," & Trim(strValor) & vbCrLf
        
        sSql = sSql & "                           ," & intAzelha & vbCrLf
        
        If lngVerCat > 0 Then
           sSql = sSql & "                           ," & lngVerCat & vbCrLf
        Else
           sSql = sSql & "                           ,Null" & vbCrLf
        End If
        If lngColEsp > 0 Then
           sSql = sSql & "                           ," & lngColEsp & vbCrLf
        Else
           sSql = sSql & "                           ,Null" & vbCrLf
        End If
        
        sSql = sSql & "                           ,'N'" & vbCrLf
        sSql = sSql & "                           ," & intNATURALSIMNAO & vbCrLf
        sSql = sSql & "                           ," & intStatus & vbCrLf
        sSql = sSql & "                           ," & intEMITLAUDO & vbCrLf
        
        If lngQTDEPORFOLHA > 0 Then
            sSql = sSql & "                           ," & lngQTDEPORFOLHA & vbCrLf
        Else
            sSql = sSql & "                           ,Null" & vbCrLf
        End If
        
        If lngQtdePassadas > 0 Then
            sSql = sSql & "                           ," & lngQtdePassadas & vbCrLf
        Else
            sSql = sSql & "                           ,Null" & vbCrLf
        End If
        
        sSql = sSql & "                           ," & intALCAGALAO & vbCrLf
        sSql = sSql & "                           ," & intNECKIN & vbCrLf
        sSql = sSql & "                           ," & intFILIALPED & vbCrLf
        sSql = sSql & "                           ," & strCAMINHO & vbCrLf
        sSql = sSql & "                           ," & intDIMPADRAO & vbCrLf
        sSql = sSql & "                           ," & strDESENV & vbCrLf
        sSql = sSql & "                           ," & strALTURA & vbCrLf
        sSql = sSql & "                           ," & strCOMPLEMENTO & vbCrLf
        sSql = sSql & "                           ," & strCODPRODENTR & vbCrLf
        sSql = sSql & "                           ," & intPRODNOVO & vbCrLf
        sSql = sSql & "                           ," & strIPI & vbCrLf
        sSql = sSql & "                           ," & intPALLHETPADRAO & vbCrLf
        sSql = sSql & "                           ," & intFOTALTSN & vbCrLf
        sSql = sSql & "                           ," & strSEGUENCIA & vbCrLf
        
        sSql = sSql & "                           )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
                
        '' Especificação Técnica
        If IsArray(arrESPTECPROD) = True Then
        
           For i = 1 To UBound(arrESPTECPROD)
           
               sSql = "Insert Into SGI_PRODESPTEC (SGI_FILIAL" & vbCrLf
               sSql = sSql & "                    ,SGI_IDPRODUTO" & vbCrLf
               sSql = sSql & "                    ,SGI_CODESPTEC)" & vbCrLf
               sSql = sSql & "             Values ( " & vbCrLf
               sSql = sSql & "                      " & intFilial & vbCrLf
               sSql = sSql & "                     ," & lngIDProduto & vbCrLf
               sSql = sSql & "                     ," & arrESPTECPROD(i) & vbCrLf
               sSql = sSql & "                     )"
               
               BGRV.CommandText = sSql
               BGRV.Execute
               
           Next i
        End If
                
        
        '' Processos .........
        If IsArray(arrPROCESSOS) = True Then
        
           For i = 1 To UBound(arrPROCESSOS)
           
               sSql = "Insert Into SGI_CADPROPROC (SGI_FILIAL,SGI_CODPROD,SGI_CODPROC,SGI_TEMPO)" & vbCrLf
               sSql = sSql & "             Values ( " & vbCrLf
               sSql = sSql & "                      " & intFilial & vbCrLf
               sSql = sSql & "                     ,'" & strCodigoProd & "'" & vbCrLf
               sSql = sSql & "                     ," & arrPROCESSOS(i, 1) & vbCrLf
               
               strValor = Replace(arrPROCESSOS(i, 2), ".", "")
               strValor = Replace(strValor, ",", ".")
               sSql = sSql & "                     ," & strValor & vbCrLf
               
               sSql = sSql & "                     )"
               
               BGRV.CommandText = sSql
               BGRV.Execute
               
           Next i
        
        End If
                
        '' Ferramentas .........
        If IsArray(arrFERRAMENTA) = True Then
        
           For i = 1 To UBound(arrFERRAMENTA)
           
               sSql = "Insert Into SGI_CADFERRPROC (SGI_FILIAL,SGI_CODPROD,SGI_CODFERRA,SGI_CODPROCESS)" & vbCrLf
               sSql = sSql & "             Values ( " & vbCrLf
               sSql = sSql & "                      " & intFilial & vbCrLf
               sSql = sSql & "                     ,'" & strCodigoProd & "'" & vbCrLf
               sSql = sSql & "                     ," & arrFERRAMENTA(i, 1) & vbCrLf
               sSql = sSql & "                     ," & arrFERRAMENTA(i, 2) & vbCrLf
               sSql = sSql & "                     )"
               
               BGRV.CommandText = sSql
               BGRV.Execute
               
           Next i
        
        End If
        
        '' Tipos de Orçamentos .........
        If IsArray(arrPRODTIPORCA) = True Then
        
           For i = 1 To UBound(arrPRODTIPORCA)
           
               sSql = "Insert Into SGI_CADPRODTIPORCA (SGI_FILIAL,SGI_IDPRODUTO,SGI_CODTIPORCA)" & vbCrLf
               sSql = sSql & "             Values ( " & vbCrLf
               sSql = sSql & "                      " & intFilial & vbCrLf
               sSql = sSql & "                     ," & lngIDProduto & vbCrLf
               sSql = sSql & "                     ," & arrPRODTIPORCA(i, 1) & vbCrLf
               sSql = sSql & "                     )"
               
               BGRV.CommandText = sSql
               BGRV.Execute
               
           Next i
        
        End If
                
        '' Variavel de Recebimento .........
        If IsArray(arrPRODVAREC) = True Then
        
           For i = 1 To UBound(arrPRODVAREC)
           
               sSql = "Insert Into SGI_CADVARECPROD (SGI_FILIAL,SGI_IDPRODUTO,SGI_CODVAR,SGI_COEF,SGI_COEFP,SGI_COEFN,SGI_TIPOINSP)" & vbCrLf
               sSql = sSql & "             Values ( " & vbCrLf
               sSql = sSql & "                      " & intFilial & vbCrLf
               sSql = sSql & "                     ," & lngIDProduto & vbCrLf
               sSql = sSql & "                     ," & arrPRODVAREC(i, 1) & vbCrLf
               
               strValor = Replace(arrPRODVAREC(i, 2), ".", "")
               strValor = Replace(strValor, ",", ".")
               sSql = sSql & "                     ," & strValor & vbCrLf
               
               strValor = Replace(arrPRODVAREC(i, 3), ".", "")
               strValor = Replace(strValor, ",", ".")
               sSql = sSql & "                     ," & strValor & vbCrLf
               
               strValor = Replace(arrPRODVAREC(i, 4), ".", "")
               strValor = Replace(strValor, ",", ".")
               sSql = sSql & "                     ," & strValor & vbCrLf
               sSql = sSql & "                     ,0" & vbCrLf
               sSql = sSql & "                     )"
               
               BGRV.CommandText = sSql
               BGRV.Execute
               
           Next i
        
        End If

        If dblSaldo > 0 Then
           '' ----------------------------------------------------
           '' Inclui Movimentação Cardex
           sSql = "Insert into SGI_CARDEXEST (SGI_FILIAL" & vbCrLf
           sSql = sSql & "                          ,SGI_CODIGO" & vbCrLf
           sSql = sSql & "                          ,SGI_IDPRODUTO" & vbCrLf
           sSql = sSql & "                          ,SGI_CODLCTO" & vbCrLf
           sSql = sSql & "                          ,SGI_dtTrans" & vbCrLf
           sSql = sSql & "                          ,SGI_QTDE" & vbCrLf
           sSql = sSql & "                          ,SGI_CODIGODOC)" & vbCrLf
           sSql = sSql & "                   Values (" & intFilial & vbCrLf
           sSql = sSql & "                          ," & lngCODLCTO & vbCrLf
           sSql = sSql & "                          ," & lngIDProduto & vbCrLf
           sSql = sSql & "                          ,1" & vbCrLf
           sSql = sSql & "                          ," & dtDataCadast & vbCrLf
        
           strValor = Replace(Format(dblSaldo, "#,###0.000"), ".", "")
           strValor = Replace(strValor, ",", ".")
           sSql = sSql & "                          ," & strValor & vbCrLf
        
           sSql = sSql & "                          ," & lngCODLCTO & ")"
        
           BGRV.CommandText = sSql
           BGRV.Execute
           '' ----------------------------------------------------
        End If
        
        '' ----------------------------------------------------
        '' Unidades de Converção
        If IsArray(arrUNIDCONV) Then
           For i = 1 To UBound(arrUNIDCONV)
               sSql = "Insert Into SGI_CADPRODUNIDCONV (SGI_FILIAL" & vbCrLf
               sSql = sSql & "                                ,SGI_IDPRODUTO" & vbCrLf
               sSql = sSql & "                                ,SGI_CODUNIDADE" & vbCrLf
               sSql = sSql & "                                ,SGI_FATOR)" & vbCrLf
               sSql = sSql & "                         Values ("
               sSql = sSql & "                                 " & intFilial & vbCrLf
               sSql = sSql & "                               ," & lngIDProduto & vbCrLf
               sSql = sSql & "                               ," & arrUNIDCONV(i, 1) & vbCrLf
               
               strValor = Replace(arrUNIDCONV(i, 2), ".", "")
               strValor = Replace(Trim(strValor), ",", ".")
               sSql = sSql & "                               ," & strValor & vbCrLf
               sSql = sSql & "                           )"
               
               BGRV.CommandText = sSql
               BGRV.Execute
           Next i
        End If
        
        '' ----------------------------------------------------
        '' Cores
        If IsArray(arrCORES) Then
            For i = 1 To UBound(arrCORES)
                sSql = "Insert Into SGI_CORESPROD (" & vbCrLf
                sSql = sSql & "                          SGI_FILIAL" & vbCrLf
                sSql = sSql & "                         ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                         ,SGI_CODCOR" & vbCrLf
                sSql = sSql & "                         ,SGI_ITEN" & vbCrLf
                sSql = sSql & "                         ,SGI_CODUNIMED" & vbCrLf
                sSql = sSql & "                         ,SGI_QTDE" & vbCrLf
                
                sSql = sSql & "                 ) Values (" & vbCrLf
                sSql = sSql & "                           " & intFilial & vbCrLf
                sSql = sSql & "                          ," & lngIDProduto & vbCrLf
                sSql = sSql & "                          ," & arrCORES(i, 1) & vbCrLf
                sSql = sSql & "                          ," & arrCORES(i, 2) & vbCrLf
                sSql = sSql & "                          ," & arrCORES(i, 3) & vbCrLf
                sSql = sSql & "                          ," & arrCORES(i, 4) & vbCrLf
                
                sSql = sSql & "                           )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
     
        '' ----------------------------------------------------
        '' Verniz
        If IsArray(arrVERNIZ) Then
            For i = 1 To UBound(arrVERNIZ)
                sSql = "Insert Into SGI_VERNIZPROD (" & vbCrLf
                sSql = sSql & "                          SGI_FILIAL" & vbCrLf
                sSql = sSql & "                         ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                         ,SGI_PRODUTO" & vbCrLf
                sSql = sSql & "                         ,SGI_CODUNIMED" & vbCrLf
                sSql = sSql & "                         ,SGI_QTDE" & vbCrLf
                
                sSql = sSql & "                 ) Values (" & vbCrLf
                sSql = sSql & "                           " & intFilial & vbCrLf
                sSql = sSql & "                          ," & lngIDProduto & vbCrLf
                sSql = sSql & "                          ," & arrVERNIZ(i, 1) & vbCrLf
                sSql = sSql & "                          ," & arrVERNIZ(i, 2) & vbCrLf
                sSql = sSql & "                          ," & arrVERNIZ(i, 3) & vbCrLf
                
                sSql = sSql & "                           )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
        
        '' Verniz 02
        If IsArray(arrVERNIZ02) Then
            For i = 1 To UBound(arrVERNIZ02)
                sSql = "Insert Into SGI_VERNIZPROD02 (" & vbCrLf
                sSql = sSql & "                          SGI_FILIAL" & vbCrLf
                sSql = sSql & "                         ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                         ,SGI_PRODUTO" & vbCrLf
                sSql = sSql & "                         ,SGI_CODUNIMED" & vbCrLf
                sSql = sSql & "                         ,SGI_QTDE" & vbCrLf
                
                sSql = sSql & "                 ) Values (" & vbCrLf
                sSql = sSql & "                           " & intFilial & vbCrLf
                sSql = sSql & "                          ," & lngIDProduto & vbCrLf
                sSql = sSql & "                          ," & arrVERNIZ02(i, 1) & vbCrLf
                sSql = sSql & "                          ," & arrVERNIZ02(i, 2) & vbCrLf
                sSql = sSql & "                          ," & arrVERNIZ02(i, 3) & vbCrLf
                
                sSql = sSql & "                           )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
     
        '' ----------------------------------------------------
        '' Verniz Acabamento
        If IsArray(arrVERNIZACAB) Then
            For i = 1 To UBound(arrVERNIZACAB)
                sSql = "Insert Into SGI_VERNIZPRODACAB (" & vbCrLf
                sSql = sSql & "                          SGI_FILIAL" & vbCrLf
                sSql = sSql & "                         ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                         ,SGI_PRODUTO" & vbCrLf
                sSql = sSql & "                         ,SGI_CODUNIMED" & vbCrLf
                sSql = sSql & "                         ,SGI_QTDE" & vbCrLf
                
                sSql = sSql & "                 ) Values (" & vbCrLf
                sSql = sSql & "                           " & intFilial & vbCrLf
                sSql = sSql & "                          ," & lngIDProduto & vbCrLf
                sSql = sSql & "                          ," & arrVERNIZACAB(i, 1) & vbCrLf
                sSql = sSql & "                          ," & arrVERNIZACAB(i, 2) & vbCrLf
                sSql = sSql & "                          ," & arrVERNIZACAB(i, 3) & vbCrLf
                
                sSql = sSql & "                           )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
        
        '' ----------------------------------------------------
        '' Esmalte
        If IsArray(arrESMALTE) Then
            For i = 1 To UBound(arrESMALTE)
                sSql = "Insert Into SGI_ESMALTEPROD (" & vbCrLf
                sSql = sSql & "                          SGI_FILIAL" & vbCrLf
                sSql = sSql & "                         ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                         ,SGI_PRODUTO" & vbCrLf
                sSql = sSql & "                         ,SGI_CODUNIMED" & vbCrLf
                sSql = sSql & "                         ,SGI_QTDE" & vbCrLf
                
                sSql = sSql & "                 ) Values (" & vbCrLf
                sSql = sSql & "                           " & intFilial & vbCrLf
                sSql = sSql & "                          ," & lngIDProduto & vbCrLf
                sSql = sSql & "                          ," & arrESMALTE(i, 1) & vbCrLf
                sSql = sSql & "                          ," & arrESMALTE(i, 2) & vbCrLf
                sSql = sSql & "                          ," & arrESMALTE(i, 3) & vbCrLf
                
                sSql = sSql & "                           )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
        
        
        '' Verniz Externo
        If IsArray(arrVERNIZEXT) Then
            For i = 1 To UBound(arrVERNIZEXT)
                sSql = "Insert Into SGI_VERNIZEXT (" & vbCrLf
                sSql = sSql & "                          SGI_FILIAL" & vbCrLf
                sSql = sSql & "                         ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                         ,SGI_PRODUTO" & vbCrLf
                sSql = sSql & "                         ,SGI_CODUNIMED" & vbCrLf
                sSql = sSql & "                         ,SGI_QTDE" & vbCrLf
                
                sSql = sSql & "                 ) Values (" & vbCrLf
                sSql = sSql & "                           " & intFilial & vbCrLf
                sSql = sSql & "                          ," & lngIDProduto & vbCrLf
                sSql = sSql & "                          ," & arrVERNIZEXT(i, 1) & vbCrLf
                sSql = sSql & "                          ," & arrVERNIZEXT(i, 2) & vbCrLf
                sSql = sSql & "                          ," & arrVERNIZEXT(i, 3) & vbCrLf
                
                sSql = sSql & "                           )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
        
        
        '' ----------------------------------------------------
        '' Vedante Fechamento
        If IsArray(arrVEDANTE) Then
            For i = 1 To UBound(arrVEDANTE)
                sSql = "Insert Into SGI_VEDANTEPROD (" & vbCrLf
                sSql = sSql & "                             SGI_FILIAL" & vbCrLf
                sSql = sSql & "                            ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                            ,SGI_PRODUTO" & vbCrLf
                sSql = sSql & "                            ,SGI_CODUNIMED" & vbCrLf
                sSql = sSql & "                            ,SGI_QTDE" & vbCrLf
                
                sSql = sSql & "                   ) Values (" & vbCrLf
                sSql = sSql & "                             " & intFilial & vbCrLf
                sSql = sSql & "                            ," & lngIDProduto & vbCrLf
                sSql = sSql & "                            ," & arrVEDANTE(i, 1) & vbCrLf
                sSql = sSql & "                            ," & arrVEDANTE(i, 2) & vbCrLf
                sSql = sSql & "                            ," & arrVEDANTE(i, 3) & vbCrLf
                
                sSql = sSql & "                            )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
        '' ----------------------------------------------------
        
        '' ----------------------------------------------------
        '' Tampa Pressão
        If IsArray(arrTAMPAPRESS) Then
            For i = 1 To UBound(arrTAMPAPRESS)
                sSql = "Insert Into SGI_FECHAMCAMPOS (" & vbCrLf
                sSql = sSql & "                              SGI_FILIAL" & vbCrLf
                sSql = sSql & "                             ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                             ,SGI_CAMPO" & vbCrLf
                sSql = sSql & "                             ,SGI_INDICE" & vbCrLf
                sSql = sSql & "                             ,SGI_VALOR" & vbCrLf
                sSql = sSql & "                    ) Values (" & vbCrLf
                sSql = sSql & "                              " & intFilial & vbCrLf
                sSql = sSql & "                             ," & lngIDProduto & vbCrLf
                sSql = sSql & "                             ," & intTampaPressao
                sSql = sSql & "                             ," & i & vbCrLf
                
                strValor = Replace(arrTAMPAPRESS(i), ".", "")
                strValor = Replace(Trim(strValor), ",", ".")
                sSql = sSql & "                             ," & strValor & vbCrLf
                
                sSql = sSql & "                             )"
            
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
     
        '' ----------------------------------------------------
        '' Batoque Retratil
        If IsArray(arrBATRETRATI) Then
            For i = 1 To UBound(arrBATRETRATI)
                sSql = "Insert Into SGI_FECHAMCAMPOS (" & vbCrLf
                sSql = sSql & "                              SGI_FILIAL" & vbCrLf
                sSql = sSql & "                             ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                             ,SGI_CAMPO" & vbCrLf
                sSql = sSql & "                             ,SGI_INDICE" & vbCrLf
                sSql = sSql & "                             ,SGI_VALOR" & vbCrLf
                sSql = sSql & "                    ) Values (" & vbCrLf
                sSql = sSql & "                              " & intFilial & vbCrLf
                sSql = sSql & "                             ," & lngIDProduto & vbCrLf
                sSql = sSql & "                             ," & intBatoqueRetratil
                sSql = sSql & "                             ," & i & vbCrLf
                
                strValor = Replace(arrBATRETRATI(i), ".", "")
                strValor = Replace(Trim(strValor), ",", ".")
                sSql = sSql & "                             ," & strValor & vbCrLf
                
                sSql = sSql & "                             )"
            
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
     
        '' ----------------------------------------------------
        '' Batoque Plastico
        If IsArray(arrBATPLASTIC) Then
            For i = 1 To UBound(arrBATPLASTIC)
                sSql = "Insert Into SGI_FECHAMCAMPOS (" & vbCrLf
                sSql = sSql & "                              SGI_FILIAL" & vbCrLf
                sSql = sSql & "                             ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                             ,SGI_CAMPO" & vbCrLf
                sSql = sSql & "                             ,SGI_INDICE" & vbCrLf
                sSql = sSql & "                             ,SGI_VALOR" & vbCrLf
                sSql = sSql & "                    ) Values (" & vbCrLf
                sSql = sSql & "                              " & intFilial & vbCrLf
                sSql = sSql & "                             ," & lngIDProduto & vbCrLf
                sSql = sSql & "                             ," & intBatoquePlastico
                sSql = sSql & "                             ," & i & vbCrLf
                
                strValor = Replace(arrBATPLASTIC(i), ".", "")
                strValor = Replace(Trim(strValor), ",", ".")
                sSql = sSql & "                             ," & strValor & vbCrLf
                
                sSql = sSql & "                             )"
            
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
     
        '' ----------------------------------------------------
        '' Tampa Visor
        If IsArray(arrTAMPAVISOR) Then
            For i = 1 To UBound(arrTAMPAVISOR)
                sSql = "Insert Into SGI_FECHAMCAMPOS (" & vbCrLf
                sSql = sSql & "                              SGI_FILIAL" & vbCrLf
                sSql = sSql & "                             ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                             ,SGI_CAMPO" & vbCrLf
                sSql = sSql & "                             ,SGI_INDICE" & vbCrLf
                sSql = sSql & "                             ,SGI_VALOR" & vbCrLf
                sSql = sSql & "                    ) Values (" & vbCrLf
                sSql = sSql & "                              " & intFilial & vbCrLf
                sSql = sSql & "                             ," & lngIDProduto & vbCrLf
                sSql = sSql & "                             ," & intTampaVisor
                sSql = sSql & "                             ," & i & vbCrLf
                
                strValor = Replace(arrTAMPAVISOR(i), ".", "")
                strValor = Replace(Trim(strValor), ",", ".")
                sSql = sSql & "                             ," & strValor & vbCrLf
                
                sSql = sSql & "                             )"
            
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
        
        '' Clientes que este produto atende
        If IsArray(arrPRODATECLIE) Then
            For i = 1 To UBound(arrPRODATECLIE)
                sSql = "Insert Into SGI_PRODATECLIE (" & vbCrLf
                sSql = sSql & "                              SGI_FILIAL" & vbCrLf
                sSql = sSql & "                             ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                             ,SGI_IDCLIENTE" & vbCrLf
                sSql = sSql & "                    ) Values (" & vbCrLf
                sSql = sSql & "                              " & intFilial & vbCrLf
                sSql = sSql & "                             ," & lngIDProduto & vbCrLf
                sSql = sSql & "                             ," & arrPRODATECLIE(i)
                sSql = sSql & "                             )"
            
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
        
        '' ------------------------
        '' Familias de Máquinas
        If IsArray(arrFAMMAQ) Then
            For i = 1 To UBound(arrFAMMAQ)
                sSql = "Insert Into SGI_PRODFAMMAQ (" & vbCrLf
                sSql = sSql & "                              SGI_FILIAL" & vbCrLf
                sSql = sSql & "                             ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                             ,SGI_CODFAMMAQ" & vbCrLf
                sSql = sSql & "                    ) Values (" & vbCrLf
                sSql = sSql & "                              " & intFilial & vbCrLf
                sSql = sSql & "                             ," & lngIDProduto & vbCrLf
                sSql = sSql & "                             ," & arrFAMMAQ(i)
                sSql = sSql & "                             )"
            
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
        
        '' Produto de ENtrada
        If IsArray(arrPRODENTR) Then
            For i = 1 To UBound(arrPRODENTR)
                sSql = "Insert Into SGI_PRODENTRADA (" & vbCrLf
                sSql = sSql & "                              SGI_FILIAL" & vbCrLf
                sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                             ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                             ,SGI_CODUNIMED" & vbCrLf
                sSql = sSql & "                             ,SGI_QTDE" & vbCrLf
                
                sSql = sSql & "                    ) Values (" & vbCrLf
                sSql = sSql & "                              " & intFilial & vbCrLf
                sSql = sSql & "                             ," & lngIDProduto & vbCrLf
                sSql = sSql & "                             ," & arrPRODENTR(i, 1) & vbCrLf
                sSql = sSql & "                             ," & arrPRODENTR(i, 2) & vbCrLf
                sSql = sSql & "                             ," & arrPRODENTR(i, 3) & vbCrLf
                
                sSql = sSql & "                             )"
            
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
        
        
        
        '' Grava A Extrutura do Produto
        ''For i = 1 To UBound(arrPROVARV)
        
        ''    If arrPROVARV(i).lngProdutoID > 0 Then
            
        ''        sSql = ""
                
        ''        sSql = "Insert Into SGI_LISTAMATPROD (" & vbCrLf
        ''        sSql = sSql & "                              SGI_FILIAL" & vbCrLf
        ''        sSql = sSql & "                             ,SGI_CODSTRU" & vbCrLf
        ''        sSql = sSql & "                             ,SGI_CODPAI" & vbCrLf
        ''        sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
        ''        sSql = sSql & "                             ,SGI_IDPRODUTO" & vbCrLf
        ''        sSql = sSql & "                             ,SGI_IDPRODPAI" & vbCrLf
        ''        sSql = sSql & "                             ,SGI_CODUNIMED" & vbCrLf
        ''        sSql = sSql & "                             ,SGI_QTDE" & vbCrLf
                
        ''        sSql = sSql & "                ) Values (" & vbCrLf
        ''        sSql = sSql & "                          " & intFilial & vbCrLf
        ''        sSql = sSql & "                         ," & lngIDProduto & vbCrLf
        ''        sSql = sSql & "                         ," & arrPROVARV(i).lngCODPAI & vbCrLf
        ''        sSql = sSql & "                         ," & arrPROVARV(i).lngCODIGO & vbCrLf
        ''        sSql = sSql & "                         ," & arrPROVARV(i).lngProdutoID & vbCrLf
        ''        sSql = sSql & "                         ," & arrPROVARV(i).lngProdutoIDPai & vbCrLf
        ''        sSql = sSql & "                         ," & arrPROVARV(i).lngCodUniMed & vbCrLf
        ''        sSql = sSql & "                         ," & arrPROVARV(i).strQTDCONS & vbCrLf
                
        ''        sSql = sSql & "                         )"
        
        ''        BGRV.CommandText = sSql
        ''        BGRV.Execute
            
        ''    End If
        
        ''Next i
        
     
     ElseIf strOPER = "A" Then
     
        '' --------------------------
        '' Apaga Tabela de Converção
        sSql = "Delete From SGI_CADPRODUNIDCONV " & vbCrLf
        sSql = sSql & "                          Where " & vbCrLf
        sSql = sSql & "                                SGI_FILIAL    = " & intFilial & vbCrLf
        sSql = sSql & "                            And SGI_IDPRODUTO = " & lngIDProduto
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------
        '' Apaga Especificação técnica
        sSql = "Delete from SGI_PRODESPTEC " & vbCrLf
        sSql = sSql & "       Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "         And SGI_FILIAL    = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------
        '' Apaga Tipos de Orçamentos
        sSql = "Delete from SGI_CADPRODTIPORCA " & vbCrLf
        sSql = sSql & "      Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------
        '' Apaga Parâmetros de Coeficiente
        sSql = "Delete From SGI_CADPARPRODPADRAO " & vbCrLf
        sSql = sSql & "      Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "        And SGI_FILIAL  = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------
        '' Apaga Coêficiente para produto padrão
        sSql = "Delete from SGI_CADPARCOEF " & vbCrLf
        sSql = sSql & "      Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        
        '' --------------------------
        '' Cores
        sSql = "Delete from SGI_CORESPROD " & vbCrLf
        sSql = sSql & "      Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------
        '' Verniz
        sSql = "Delete From SGI_VERNIZPROD " & vbCrLf
        sSql = sSql & "      Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' Verniz  01
        sSql = "Delete From SGI_VERNIZPROD02 " & vbCrLf
        sSql = sSql & "      Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------
        '' Verniz Acabamento
        sSql = "Delete From SGI_VERNIZPRODACAB " & vbCrLf
        sSql = sSql & "      Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------
        '' Esmalte
        sSql = "Delete From SGI_ESMALTEPROD " & vbCrLf
        sSql = sSql & "      Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------
        '' Verniz Externo
        sSql = "Delete From SGI_VERNIZEXT " & vbCrLf
        sSql = sSql & "      Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------
        '' Vedante Fechamento
        sSql = "Delete From SGI_VEDANTEPROD " & vbCrLf
        sSql = sSql & "      Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        
        '' --------------------------
        '' Fechamento
        sSql = "Delete From SGI_FECHAMCAMPOS " & vbCrLf
        sSql = sSql & "      Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------
        '' Cliente que este produto atende
        sSql = "Delete From SGI_PRODATECLIE " & vbCrLf
        sSql = sSql & "      Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------
        '' Familias de Máquinas
        sSql = "Delete From SGI_PRODFAMMAQ " & vbCrLf
        sSql = sSql & "      Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------
        '' Produtos de Entrada
        sSql = "Delete From SGI_PRODENTRADA " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & lngIDProduto & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        
        '' --------------------------
        '' Grava Produto
        sSql = ""
        
        sSql = "Update SGI_CADPRODUTO Set " & vbCrLf
        
        sSql = sSql & "                   SGI_CODIGO     = '" & Trim(strCodigoProd) & "'" & vbCrLf
        sSql = sSql & "                  ,SGI_DESCRICAO  = '" & strDescriProd & "'" & vbCrLf
        sSql = sSql & "                  ,SGI_CODESPECIE =  " & intEspProduto & vbCrLf
        sSql = sSql & "                  ,SGI_CODTIPO    =  " & intTIPPRODUTO & vbCrLf
        sSql = sSql & "                  ,SGI_UNIDMEDIDA =  " & intUnidade & vbCrLf
        sSql = sSql & "                  ,SGI_CODCLASS   =  " & lngCLASFISC & vbCrLf
        
        strValor = Replace(Format(dblSaldo, "#,###0.000"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                  ,SGI_SALDO      =  " & Trim(strValor) & vbCrLf
        
        strValor = Replace(Format(dblEstMin, "#,###0.000"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                  ,SGI_ESTMINIMO  =  " & Trim(strValor) & vbCrLf
        
        sSql = sSql & "                  ,SGI_DTCADASTRO = " & dtDataCadast & vbCrLf
        
        strValor = Replace(Format(dblPRCPROD, "#,##0.00"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                  ,SGI_PRECOPROD  = " & Trim(strValor) & vbCrLf
        
        strValor = Replace(Format(dblPRCMED, "#,##0.00"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                  ,SGI_PRECOMED   = " & Trim(strValor) & vbCrLf
        
        sSql = sSql & "                  ,SGI_ATUAUTOM   = " & intATAUTOMSN & vbCrLf
        sSql = sSql & "                  ,SGI_DESCPROD   = '" & strDESCEQUIP & "'" & vbCrLf
        sSql = sSql & "                  ,SGI_PROCEDENCIA  = '" & strPROCEDEN & "'" & vbCrLf
        sSql = sSql & "                  ,SGI_CODGPROD     = " & intCODGRUPPROD & vbCrLf
        sSql = sSql & "                  ,SGI_CODSUBGPROD  = " & intCODSUBGPROD & vbCrLf
        
        strValor = Replace(Format(curPORCACRES, "#,##0.00"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                  ,SGI_PROCACRES    = " & strValor & vbCrLf
        
        strValor = Replace(Format(curPRCCUSTO, "#,##0.00"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                  ,SGI_PRCCUSTO     = " & strValor & vbCrLf
        
        sSql = sSql & "                  ,SGI_PRODUTOTIPO   = " & intPRODUTOTIPO & vbCrLf
        sSql = sSql & "                  ,SGI_PRODUTOESTILO = " & intESTILOPROD & vbCrLf
        
        strValor = Replace(Format(curPESOUNIT, "#,###0.000"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                  ,SGI_PESOUNIT      = " & strValor & vbCrLf
        If lngTABCONV > 0 Then
           sSql = sSql & "                  ,SGI_CODTABCONV    = " & lngTABCONV & vbCrLf
        Else
           sSql = sSql & "                  ,SGI_CODTABCONV    = NULL" & vbCrLf
        End If
        sSql = sSql & "                  ,SGI_CODPRODFORN   = '" & strCODPROFORNEC & "'" & vbCrLf
        
        sSql = sSql & "                  ,SGI_RUA           = '" & strRUA & "'" & vbCrLf
        sSql = sSql & "                  ,SGI_BOX           = '" & strBOX & "'" & vbCrLf
        sSql = sSql & "                  ,SGI_PRATELEIRA    = '" & strPRATELEIRA & "'" & vbCrLf
        sSql = sSql & "                  ,SGI_CODEAN        = '" & strCODEAN & "'" & vbCrLf
        
        strValor = Replace(Format(curDISTANCIA, "#,###0.000"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                  ,SGI_DISTESTOQUE      = " & strValor & vbCrLf
        sSql = sSql & "                  ,SGI_PADRAO           = " & intPADRAO & vbCrLf
        
        If lngFAMMAQUINA > 0 Then
            sSql = sSql & "                  ,SGI_CADFAMMAQ        = " & lngFAMMAQUINA & vbCrLf
        Else
            sSql = sSql & "                  ,SGI_CADFAMMAQ        = Null" & vbCrLf
        End If
        
        strValor = Replace(Format(curCORRELACAO, "#,###0.000"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                  ,SGI_CORRPRODPADR     = " & strValor & vbCrLf
        
        strValor = Replace(Format(curPORCCORR, "#,##0.00"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                  ,SGI_PORCCORR         = " & strValor & vbCrLf
        
        strValor = Replace(Format(curCUBAGEN, "#,###0.000"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                  ,SGI_CUBAGEN          = " & strValor & vbCrLf
        
        strValor = Replace(Format(curMetros, "#,###0.000"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                  ,SGI_METROS          = " & strValor & vbCrLf
        
        strValor = Replace(Format(curLargura, "#,###0.000"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                  ,SGI_LARGURA         = " & strValor & vbCrLf
        
        strValor = Replace(Format(curGRAMATURA2, "#,###0.000"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                  ,SGI_GRAMATURA2      = " & strValor & vbCrLf
        
        If lngVernCorpo > 0 Then
            sSql = sSql & "                           ,SGI_VernCorpo = " & lngVernCorpo & vbCrLf
        Else
            sSql = sSql & "                           ,SGI_VernCorpo = Null" & vbCrLf
        End If
        If lngVernTampa > 0 Then
            sSql = sSql & "                           ,SGI_VernTampa = " & lngVernTampa & vbCrLf
        Else
            sSql = sSql & "                           ,SGI_VernTampa = Null" & vbCrLf
        End If
        If lngVernFundo > 0 Then
            sSql = sSql & "                           ,SGI_VernFundo = " & lngVernFundo & vbCrLf
        Else
            sSql = sSql & "                           ,SGI_VernFundo = Null" & vbCrLf
        End If
        If lngVernArgola > 0 Then
            sSql = sSql & "                           ,SGI_VernArgola = " & lngVernArgola & vbCrLf
        Else
            sSql = sSql & "                           ,SGI_VernArgola = Null" & vbCrLf
        End If
        
        
        If lngFechTampaFuro > -1 Then
            sSql = sSql & "                           ,SGI_FechTampaFuro = " & lngFechTampaFuro & vbCrLf
        Else
            sSql = sSql & "                           ,SGI_FechTampaFuro = Null" & vbCrLf
        End If
        
        
        If lngEspessCorpo > 0 Then
            sSql = sSql & "                           ,SGI_CorpoEspess = " & lngEspessCorpo & vbCrLf
        Else
            sSql = sSql & "                           ,SGI_CorpoEspess = Null" & vbCrLf
        End If
        If lngEspessTampa > 0 Then
            sSql = sSql & "                           ,SGI_TampaEspess = " & lngEspessTampa & vbCrLf
        Else
            sSql = sSql & "                           ,SGI_TampaEspess = Null" & vbCrLf
        End If
        If lngEspessFundo > 0 Then
            sSql = sSql & "                           ,SGI_FundoEspess = " & lngEspessFundo & vbCrLf
        Else
            sSql = sSql & "                           ,SGI_FundoEspess = Null" & vbCrLf
        End If
        If lngEspessArgola > 0 Then
            sSql = sSql & "                           ,SGI_ArgolaEspess = " & lngEspessArgola & vbCrLf
        Else
            sSql = sSql & "                           ,SGI_ArgolaEspess = Null" & vbCrLf
        End If
        
        
        If lngRevestCorpo > 0 Then
            sSql = sSql & "                           ,SGI_CorpoRevest = " & lngRevestCorpo & vbCrLf
        Else
            sSql = sSql & "                           ,SGI_CorpoRevest = Null" & vbCrLf
        End If
        If lngRevestTampa > 0 Then
            sSql = sSql & "                           ,SGI_TampaRevest = " & lngRevestTampa & vbCrLf
        Else
            sSql = sSql & "                           ,SGI_TampaRevest = Null" & vbCrLf
        End If
        If lngRevestFundo > 0 Then
            sSql = sSql & "                           ,SGI_FundoRevest = " & lngRevestFundo & vbCrLf
        Else
            sSql = sSql & "                           ,SGI_FundoRevest = Null" & vbCrLf
        End If
        If lngRevestArgola > 0 Then
            sSql = sSql & "                           ,SGI_ArgolaRevest = " & lngRevestArgola & vbCrLf
        Else
            sSql = sSql & "                           ,SGI_ArgolaRevest = Null" & vbCrLf
        End If
        
        
        If lngRevestCorpo2 > 0 Then
            sSql = sSql & "                           ,SGI_CorpoRevest2 = " & lngRevestCorpo2 & vbCrLf
        Else
            sSql = sSql & "                           ,SGI_CorpoRevest2 = Null" & vbCrLf
        End If
        If lngRevestTampa2 > 0 Then
            sSql = sSql & "                           ,SGI_TampaRevest2 = " & lngRevestTampa2 & vbCrLf
        Else
            sSql = sSql & "                           ,SGI_TampaRevest2 = Null" & vbCrLf
        End If
        If lngRevestFundo2 > 0 Then
            sSql = sSql & "                           ,SGI_FundoRevest2 = " & lngRevestFundo2 & vbCrLf
        Else
            sSql = sSql & "                           ,SGI_FundoRevest2 = Null" & vbCrLf
        End If
        If lngRevestArgola2 > 0 Then
            sSql = sSql & "                           ,SGI_ArgolaRevest2 = " & lngRevestArgola2 & vbCrLf
        Else
            sSql = sSql & "                           ,SGI_ArgolaRevest2 = Null" & vbCrLf
        End If
        
        
        If lngAlcaPlastica > 0 Then
            sSql = sSql & "                           ,SGI_AlcaPlastica = " & lngAlcaPlastica & vbCrLf
        Else
            sSql = sSql & "                           ,SGI_AlcaPlastica = Null" & vbCrLf
        End If
        If lngAlcaFerro > 0 Then
            sSql = sSql & "                           ,SGI_AlcaFerro = " & lngAlcaFerro & vbCrLf
        Else
            sSql = sSql & "                           ,SGI_AlcaFerro = Null" & vbCrLf
        End If
        
        sSql = sSql & "                           ,SGI_Pipeta = " & intPipeta & vbCrLf
        
        If lngFechSoldaAgrafado > -1 Then
            sSql = sSql & "                           ,SGI_FechSoldaAgrafado = " & lngFechSoldaAgrafado & vbCrLf
        Else
            sSql = sSql & "                           ,SGI_FechSoldaAgrafado = Null" & vbCrLf
        End If
        sSql = sSql & "                           ,SGI_FechSA = " & intFechSA
        
        
        sSql = sSql & "                  ,SGI_FechamTampaPress         = Null" & vbCrLf
        sSql = sSql & "                  ,SGI_FechamBatoqRetra         = Null" & vbCrLf
        sSql = sSql & "                  ,SGI_FechamBatoqPlast         = Null" & vbCrLf
        sSql = sSql & "                  ,SGI_FechamTampaVisor         = Null" & vbCrLf
        
        
        sSql = sSql & "                  ,SGI_TIPOEMPILHAMENTO         = " & intTipoPalletGranel & vbCrLf
        
        strValor = Replace(Format(lngQtdePalletGranel, "#,###0.000"), ".", "")
        strValor = Replace(strValor, ",", ".")
        sSql = sSql & "                  ,SGI_QTDEEMPILHAMENTO         = " & Trim(strValor) & vbCrLf
         
        sSql = sSql & "                  ,SGI_AZELHA                   = " & intAzelha & vbCrLf
        
        If lngVerCat > 0 Then
           sSql = sSql & "                  ,SGI_VERCAT         = " & lngVerCat & vbCrLf
        Else
           sSql = sSql & "                  ,SGI_VERCAT         = Null" & vbCrLf
        End If
        
        If lngColEsp > 0 Then
           sSql = sSql & "                  ,SGI_COLESP         = " & lngColEsp & vbCrLf
        Else
           sSql = sSql & "                  ,SGI_COLESP         = Null" & vbCrLf
        End If
        
        If lngDigVerif > 0 Then
            sSql = sSql & "                    ,SGI_DIGVERIF = " & lngDigVerif & vbCrLf
        Else
            sSql = sSql & "                    ,SGI_DIGVERIF = Null" & vbCrLf
        End If
        sSql = sSql & "                    ,SGI_EMITLAUDO = " & intEMITLAUDO & vbCrLf
        
        If lngQTDEPORFOLHA > 0 Then
            sSql = sSql & "                           ,SGI_QTDEPORFOLHA = " & lngQTDEPORFOLHA & vbCrLf
        Else
            sSql = sSql & "                           ,SGI_QTDEPORFOLHA = Null" & vbCrLf
        End If
        If lngQtdePassadas > 0 Then
            sSql = sSql & "                           ,SGI_QTDEPASSADA = " & lngQtdePassadas & vbCrLf
        Else
            sSql = sSql & "                           ,SGI_QTDEPASSADA = Null" & vbCrLf
        End If
        
        sSql = sSql & "                        ,SGI_ALCAGALAO        = " & intALCAGALAO & vbCrLf
        sSql = sSql & "                        ,SGI_NECKIN           = " & intNECKIN & vbCrLf
        sSql = sSql & "                        ,SGI_FILIALPED        = " & intFILIALPED & vbCrLf
        sSql = sSql & "                        ,SGI_ARQUIVOIMG       = " & strCAMINHO & vbCrLf
        sSql = sSql & "                        ,SGI_DIMPADRAO        = " & intDIMPADRAO & vbCrLf
        sSql = sSql & "                        ,SGI_DESENV           = " & strDESENV & vbCrLf
        sSql = sSql & "                        ,SGI_ALTURA           = " & strALTURA & vbCrLf
        sSql = sSql & "                        ,SGI_COMPLEMENTO      = " & strCOMPLEMENTO & vbCrLf
        sSql = sSql & "                        ,SGI_CODPRODENTR      = " & strCODPRODENTR & vbCrLf
        sSql = sSql & "                        ,SGI_QTDCORPSPADRAOSN = " & intQTDCORPSPADRAOSN & vbCrLf
        sSql = sSql & "                        ,SGI_PRODNOVO         = " & intPRODNOVO & vbCrLf
        sSql = sSql & "                        ,SGI_IPI              = " & strIPI & vbCrLf
        sSql = sSql & "                        ,SGI_USARPADRPALSN    = " & intPALLHETPADRAO & vbCrLf
        sSql = sSql & "                        ,SGI_SEGUENCIA        = " & strSEGUENCIA & vbCrLf
        
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL =  " & intFilial & vbCrLf
        sSql = sSql & "   And SGI_IDPRODUTO = " & lngIDProduto
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' Especificação Técnica
        If IsArray(arrESPTECPROD) = True Then
        
           For i = 1 To UBound(arrESPTECPROD)
           
               sSql = "Insert Into SGI_PRODESPTEC (SGI_FILIAL,SGI_IDPRODUTO,SGI_CODESPTEC)"
               sSql = sSql & "             Values ( "
               sSql = sSql & "                      " & intFilial & vbCrLf
               sSql = sSql & "                     ," & lngIDProduto & vbCrLf
               sSql = sSql & "                     ," & arrESPTECPROD(i) & vbCrLf
               sSql = sSql & "                     )"
               
               BGRV.CommandText = sSql
               BGRV.Execute
               
           Next i
        End If
        
        '' Processos
        ''If IsArray(arrPROCESSOS) = True Then
        
        ''   For I = 1 To UBound(arrPROCESSOS)
           
        ''       If Len(Trim(arrPROCESSOS(I, 3))) > 0 Then
               
        ''          sSql = "Insert Into SGI_CADPROPROC (SGI_FILIAL,SGI_CODPROD,SGI_CODPROC,SGI_TEMPO,SGI_VALORHORA,SGI_TOTVALOR)" & vbCrLf
        ''          sSql = sSql & "             Values ( " & vbCrLf
        ''          sSql = sSql & "                      " & intFilial & vbCrLf
        ''          sSql = sSql & "                     ,'" & Trim(strCodigoProd) & "'" & vbCrLf
        ''          sSql = sSql & "                     ," & arrPROCESSOS(I, 1) & vbCrLf
               
        ''          strValor = Replace(arrPROCESSOS(I, 2), ".", "")
        ''          strValor = Replace(strValor, ",", ".")
        ''          If Len(Trim(strValor)) > 0 Then
        ''             sSql = sSql & "                     ," & strValor & vbCrLf
        ''          Else
        ''             sSql = sSql & "                     ,Null" & vbCrLf
        ''          End If
               
        ''          strValor = Replace(arrPROCESSOS(I, 3), ".", "")
        ''          strValor = Replace(strValor, ",", ".")
        ''          If Len(Trim(strValor)) > 0 Then
        ''             sSql = sSql & "                     ," & strValor & vbCrLf
        ''          Else
        ''             sSql = sSql & "                     ,Null" & vbCrLf
        ''          End If
               
        ''          strValor = Replace(arrPROCESSOS(I, 4), ".", "")
        ''          strValor = Replace(strValor, ",", ".")
        ''          If Len(Trim(strValor)) > 0 Then
        ''             sSql = sSql & "                     ," & strValor & vbCrLf
        ''          Else
        ''             sSql = sSql & "                     ,Null" & vbCrLf
        ''          End If
               
        ''          sSql = sSql & "                     )"
               
        ''          BGRV.CommandText = sSql
        ''          BGRV.Execute
               
        ''       End If
               
        ''   Next I
       
        ''End If
        
        '' Ferramentas .........
        ''If IsArray(arrFERRAMENTA) = True Then
     
        ''   For I = 1 To UBound(arrFERRAMENTA)
       
        ''       sSql = "Insert Into SGI_CADFERRPROC (SGI_FILIAL,SGI_CODPROD,SGI_CODFERRA,SGI_CODPROCESS)" & vbCrLf
        ''       sSql = sSql & "             Values ( " & vbCrLf
        ''       sSql = sSql & "                      " & intFilial & vbCrLf
        ''       sSql = sSql & "                     ,'" & strCodigoProd & "'" & vbCrLf
        ''       sSql = sSql & "                     ," & arrFERRAMENTA(I, 1) & vbCrLf
        ''       sSql = sSql & "                     ," & arrFERRAMENTA(I, 2) & vbCrLf
        ''       sSql = sSql & "                     )"
           
        ''       BGRV.CommandText = sSql
        ''       BGRV.Execute
           
        ''   Next I
    
        ''End If
        
        '' Tipos de Orçamentos .........
        If IsArray(arrPRODTIPORCA) = True Then
        
           For i = 1 To UBound(arrPRODTIPORCA)
           
               sSql = "Insert Into SGI_CADPRODTIPORCA (SGI_FILIAL,SGI_IDPRODUTO,SGI_CODTIPORCA)" & vbCrLf
               sSql = sSql & "             Values ( " & vbCrLf
               sSql = sSql & "                      " & intFilial & vbCrLf
               sSql = sSql & "                     ," & lngIDProduto & vbCrLf
               sSql = sSql & "                     ," & arrPRODTIPORCA(i) & vbCrLf
               sSql = sSql & "                     )"
               
               BGRV.CommandText = sSql
               BGRV.Execute
               
           Next i
        
        End If
        
        '' Variavel de Recebimento .........
        ''If IsArray(arrPRODVAREC) = True Then
        
        ''   For i = 1 To UBound(arrPRODVAREC)
           
        ''       sSql = "Insert Into SGI_CADVARECPROD (SGI_FILIAL,SGI_IDPRODUTO,SGI_CODVAR,SGI_COEF,SGI_COEFP,SGI_COEFN,SGI_TIPOINSP)" & vbCrLf
        ''       sSql = sSql & "             Values ( " & vbCrLf
        ''       sSql = sSql & "                      " & intFilial & vbCrLf
        ''       sSql = sSql & "                     ," & lngIDPRODUTO & vbCrLf
        ''       sSql = sSql & "                     ," & Trim(arrPRODVAREC(i, 1)) & vbCrLf
        ''
        ''       strValor = Replace(arrPRODVAREC(i, 2), ".", "")
        ''       strValor = Replace(strValor, ",", ".")
        ''       sSql = sSql & "                     ," & strValor & vbCrLf
        ''
        ''       strValor = Replace(arrPRODVAREC(i, 3), ".", "")
        ''       strValor = Replace(strValor, ",", ".")
        ''       sSql = sSql & "                     ," & strValor & vbCrLf
        ''
        ''       strValor = Replace(arrPRODVAREC(i, 4), ".", "")
        ''       strValor = Replace(strValor, ",", ".")
        ''       sSql = sSql & "                     ," & strValor & vbCrLf
        ''       sSql = sSql & "                     ,0" & vbCrLf
         ''      sSql = sSql & "                     )"
        ''
        ''       BGRV.CommandText = sSql
        ''       BGRV.Execute
        ''
        ''   Next i
      ''
       '' End If

        '' Atributo de Recebimento .........
        
        'If IsArray(arrPRODATRREC) = True Then
        '
        '   For i = 1 To UBound(arrPRODATRREC)
        '
        '       sSql = "Insert Into SGI_CADVARECPROD (SGI_FILIAL,SGI_IDPRODUTO,SGI_CODVAR,SGI_ESPATRIB,SGI_CONATRIB,SGI_TIPOINSP)" & vbCrLf
        '       sSql = sSql & "             Values ( " & vbCrLf
        '       sSql = sSql & "                      " & intFilial & vbCrLf
        '       sSql = sSql & "                     ," & lngIDPRODUTO & vbCrLf
        '       sSql = sSql & "                     ," & arrPRODATRREC(i, 1) & vbCrLf
        '       sSql = sSql & "                     ,'" & arrPRODATRREC(i, 2) & "'" & vbCrLf
        '       sSql = sSql & "                     ,'" & arrPRODATRREC(i, 3) & "'" & vbCrLf
        '       sSql = sSql & "                     ,1" & vbCrLf
        '       sSql = sSql & "                     )"
        '
        '       BGRV.CommandText = sSql
        '       BGRV.Execute
        '
        '   Next i
       '
       ' End If
        
        '' Variante de Processo .........
        'If IsArray(arrPRODVAPROC) = True Then
       '
       '    For i = 1 To UBound(arrPRODVAPROC)
       '
       '        sSql = "Insert Into SGI_CADVAINSPPROC (SGI_FILIAL,SGI_IDPRODUTO,SGI_CODPROC,SGI_CODVAR,SGI_COEF,SGI_COEFP,SGI_COEFN,SGI_TIPOINSP)" & vbCrLf
       '        sSql = sSql & "             Values ( " & vbCrLf
       '        sSql = sSql & "                      " & intFilial & vbCrLf
       '        sSql = sSql & "                     ," & lngIDPRODUTO & vbCrLf
       '        sSql = sSql & "                     ," & Trim(arrPRODVAPROC(i, 5)) & vbCrLf
       '        sSql = sSql & "                     ," & Trim(arrPRODVAPROC(i, 1)) & vbCrLf
       '
       '        strValor = Replace(arrPRODVAPROC(i, 2), ".", "")
       '        strValor = Replace(strValor, ",", ".")
       '        sSql = sSql & "                     ," & strValor & vbCrLf
       '
       '        strValor = Replace(arrPRODVAPROC(i, 3), ".", "")
       '        strValor = Replace(strValor, ",", ".")
       '        sSql = sSql & "                     ," & strValor & vbCrLf
       '
       '        strValor = Replace(arrPRODVAPROC(i, 4), ".", "")
       '        strValor = Replace(strValor, ",", ".")
       '        sSql = sSql & "                     ," & strValor & vbCrLf
       '        sSql = sSql & "                     ,0" & vbCrLf
       '        sSql = sSql & "                     )"
       '
       '        BGRV.CommandText = sSql
       '        BGRV.Execute
       '
       '    Next i
       '
       ' End If
        
        '' Atributo de Processo .........
        'If IsArray(arrPRODATRPROC) = True Then
       '
       '    For i = 1 To UBound(arrPRODATRPROC)
       '
       '        sSql = "Insert Into SGI_CADVAINSPPROC (SGI_FILIAL,SGI_IDPRODUTO,SGI_CODPROC,SGI_CODVAR,SGI_ESPATRIB,SGI_CONATRIB,SGI_TIPOINSP)" & vbCrLf
       '        sSql = sSql & "             Values ( " & vbCrLf
       '        sSql = sSql & "                      " & intFilial & vbCrLf
       '        sSql = sSql & "                     ," & lngIDPRODUTO & vbCrLf
       '        sSql = sSql & "                     ," & Trim(arrPRODATRPROC(i, 4)) & vbCrLf
       '        sSql = sSql & "                     ," & Trim(arrPRODATRPROC(i, 1)) & vbCrLf
       '        sSql = sSql & "                     ,'" & arrPRODATRPROC(i, 2) & "'" & vbCrLf
       '        sSql = sSql & "                     ,'" & arrPRODATRPROC(i, 3) & "'" & vbCrLf
       '        sSql = sSql & "                     ,1" & vbCrLf
       '        sSql = sSql & "                     )"
       '
       '        BGRV.CommandText = sSql
       '        BGRV.Execute
       '
       '    Next i
       '
       ' End If
       
        '' Variante de Inspeção Final .........
        'If IsArray(arrPRODVAINSP) = True Then
       '
       '    For i = 1 To UBound(arrPRODVAINSP)
       '
       '        sSql = "Insert Into SGI_CADVAINSPPROD (SGI_FILIAL,SGI_IDPRODUTO,SGI_CODVAR,SGI_COEF,SGI_COEFP,SGI_COEFN,SGI_TIPOINSP)" & vbCrLf
       '        sSql = sSql & "             Values ( " & vbCrLf
       '        sSql = sSql & "                      " & intFilial & vbCrLf
       '        sSql = sSql & "                     ," & lngIDPRODUTO & vbCrLf
       '        sSql = sSql & "                     ," & arrPRODVAINSP(i, 1) & vbCrLf
       '
       '        strValor = Replace(arrPRODVAINSP(i, 2), ".", "")
       '        strValor = Replace(strValor, ",", ".")
       '        sSql = sSql & "                     ," & strValor & vbCrLf
       '
       '        strValor = Replace(arrPRODVAINSP(i, 3), ".", "")
       '        strValor = Replace(strValor, ",", ".")
       '        sSql = sSql & "                     ," & strValor & vbCrLf
       '
       '        strValor = Replace(arrPRODVAINSP(i, 4), ".", "")
       '        strValor = Replace(strValor, ",", ".")
       '        sSql = sSql & "                     ," & strValor & vbCrLf
       '        sSql = sSql & "                     ,0" & vbCrLf
       '        sSql = sSql & "                     )"
       '
       '        BGRV.CommandText = sSql
       '        BGRV.Execute
       '
       '    Next i
       '
       ' End If
        
        '' Atributo de Inspeção Final .........
        'If IsArray(arrPRODATRINSP) = True Then
       '
       '    For i = 1 To UBound(arrPRODATRINSP)
       '
       '        sSql = "Insert Into SGI_CADVAINSPPROD (SGI_FILIAL,SGI_IDPRODUTO,SGI_CODVAR,SGI_ESPATRIB,SGI_CONATRIB,SGI_TIPOINSP)" & vbCrLf
       '        sSql = sSql & "             Values ( " & vbCrLf
       '        sSql = sSql & "                      " & intFilial & vbCrLf
       '        sSql = sSql & "                     ," & lngIDPRODUTO & vbCrLf
       '        sSql = sSql & "                     ," & arrPRODATRINSP(i, 1) & vbCrLf
       '        sSql = sSql & "                     ,'" & arrPRODATRINSP(i, 2) & "'" & vbCrLf
       '        sSql = sSql & "                     ,'" & arrPRODATRINSP(i, 3) & "'" & vbCrLf
       '        sSql = sSql & "                     ,1" & vbCrLf
       '        sSql = sSql & "                     )"
       '
       '        BGRV.CommandText = sSql
       '        BGRV.Execute
       '
       '    Next i
       '
       ' End If
    '
        '' Gravando os Parâmetros .........
        'If IsArray(arrCORRPRODPADR) Then
        '   For i = 1 To UBound(arrCORRPRODPADR)
        '       sSql = " Insert Into SGI_CADPARPRODPADRAO (" & vbCrLf
        '       sSql = sSql & "                                SGI_FILIAL" & vbCrLf
        '       sSql = sSql & "                               ,SGI_IDPRODUTO" & vbCrLf
        '       sSql = sSql & "                               ,SGI_CODPAR" & vbCrLf
        '       sSql = sSql & "                               ,SGI_CODUNI" & vbCrLf
        '       sSql = sSql & "                               ,SGI_PESO" & vbCrLf
        '       sSql = sSql & "                               )" & vbCrLf
        '       sSql = sSql & "                        Values (" & vbCrLf
        '       sSql = sSql & "                                " & intFilial & vbCrLf
        '       sSql = sSql & "                               ," & lngIDPRODUTO & vbCrLf
        '       sSql = sSql & "                               ," & arrCORRPRODPADR(i, 1) & vbCrLf
        '       sSql = sSql & "                               ," & arrCORRPRODPADR(i, 2) & vbCrLf
        '
        '       strValor = Replace(Format(arrCORRPRODPADR(i, 3), "#,###0.000"), ".", "")
        '       strValor = Replace(strValor, ",", ".")
        '       sSql = sSql & "                              ," & strValor & vbCrLf
        '
        '       sSql = sSql & "                               )" & vbCrLf
        '
        '       BGRV.CommandText = sSql
        '       BGRV.Execute
        '
        '   Next i
        'End If
     
        '' Parâmetros para o Coêficiente para produto padrão
        If IsArray(arrPARPROD) Then
           For i = 1 To UBound(arrPARPROD)
           
               sSql = "Insert Into SGI_CADPARCOEF (" & vbCrLf
               sSql = sSql & "                            SGI_FILIAL" & vbCrLf
               sSql = sSql & "                           ,SGI_IDPRODUTO" & vbCrLf
               sSql = sSql & "                           ,SGI_CODPARAM" & vbCrLf
               sSql = sSql & "                           ,SGI_UNID" & vbCrLf
               sSql = sSql & "                           ,SGI_QTDE" & vbCrLf
               sSql = sSql & "                           )" & vbCrLf
               sSql = sSql & "                    Values (" & vbCrLf
               sSql = sSql & "                            " & intFilial & vbCrLf
               sSql = sSql & "                           ," & lngIDProduto & vbCrLf
               sSql = sSql & "                           ," & arrPARPROD(i, 1) & vbCrLf
               sSql = sSql & "                           ," & arrPARPROD(i, 2) & vbCrLf
               
               strValor = Replace(arrPARPROD(i, 3), ".", "")
               strValor = Replace(Trim(strValor), ",", ".")
               sSql = sSql & "                           ," & strValor & vbCrLf
               
               sSql = sSql & "                           )"
               
               BGRV.CommandText = sSql
               BGRV.Execute
               
           Next i
        End If
        
        '' Unidades de Converção
        If IsArray(arrUNIDCONV) Then
           For i = 1 To UBound(arrUNIDCONV)
               sSql = "Insert Into SGI_CADPRODUNIDCONV (SGI_FILIAL" & vbCrLf
               sSql = sSql & "                                ,SGI_IDPRODUTO" & vbCrLf
               sSql = sSql & "                                ,SGI_CODUNIDADE" & vbCrLf
               sSql = sSql & "                                ,SGI_FATOR)" & vbCrLf
               sSql = sSql & "                         Values ("
               sSql = sSql & "                                 " & intFilial & vbCrLf
               sSql = sSql & "                               ," & lngIDProduto & vbCrLf
               sSql = sSql & "                               ," & arrUNIDCONV(i, 1) & vbCrLf
               
               strValor = Replace(arrUNIDCONV(i, 2), ".", "")
               strValor = Replace(Trim(strValor), ",", ".")
               sSql = sSql & "                               ," & strValor & vbCrLf
               sSql = sSql & "                           )"
               
               BGRV.CommandText = sSql
               BGRV.Execute
           Next i
        End If
        
        
        '' ----------------------------------------------------
        '' Cores
        If IsArray(arrCORES) Then
            For i = 1 To UBound(arrCORES)
                sSql = "Insert Into SGI_CORESPROD (" & vbCrLf
                sSql = sSql & "                          SGI_FILIAL" & vbCrLf
                sSql = sSql & "                         ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                         ,SGI_CODCOR" & vbCrLf
                sSql = sSql & "                         ,SGI_ITEN" & vbCrLf
                sSql = sSql & "                         ,SGI_CODUNIMED" & vbCrLf
                sSql = sSql & "                         ,SGI_QTDE" & vbCrLf
                
                sSql = sSql & "                 ) Values (" & vbCrLf
                sSql = sSql & "                           " & intFilial & vbCrLf
                sSql = sSql & "                          ," & lngIDProduto & vbCrLf
                sSql = sSql & "                          ," & arrCORES(i, 1) & vbCrLf
                sSql = sSql & "                          ," & arrCORES(i, 2) & vbCrLf
                sSql = sSql & "                          ," & arrCORES(i, 3) & vbCrLf
                sSql = sSql & "                          ," & arrCORES(i, 4) & vbCrLf
                
                sSql = sSql & "                           )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
     
        '' ----------------------------------------------------
        '' Verniz
        If IsArray(arrVERNIZ) Then
            For i = 1 To UBound(arrVERNIZ)
                sSql = "Insert Into SGI_VERNIZPROD (" & vbCrLf
                sSql = sSql & "                          SGI_FILIAL" & vbCrLf
                sSql = sSql & "                         ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                         ,SGI_PRODUTO" & vbCrLf
                sSql = sSql & "                         ,SGI_CODUNIMED" & vbCrLf
                sSql = sSql & "                         ,SGI_QTDE" & vbCrLf
                
                sSql = sSql & "                 ) Values (" & vbCrLf
                sSql = sSql & "                           " & intFilial & vbCrLf
                sSql = sSql & "                          ," & lngIDProduto & vbCrLf
                sSql = sSql & "                          ," & arrVERNIZ(i, 1) & vbCrLf
                sSql = sSql & "                          ," & arrVERNIZ(i, 2) & vbCrLf
                sSql = sSql & "                          ," & arrVERNIZ(i, 3) & vbCrLf
                
                sSql = sSql & "                           )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
        '' ----------------------------------------------------
        '' Verniz 02
        If IsArray(arrVERNIZ02) Then
            For i = 1 To UBound(arrVERNIZ02)
                sSql = "Insert Into SGI_VERNIZPROD02 (" & vbCrLf
                sSql = sSql & "                          SGI_FILIAL" & vbCrLf
                sSql = sSql & "                         ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                         ,SGI_PRODUTO" & vbCrLf
                sSql = sSql & "                         ,SGI_CODUNIMED" & vbCrLf
                sSql = sSql & "                         ,SGI_QTDE" & vbCrLf
                
                sSql = sSql & "                 ) Values (" & vbCrLf
                sSql = sSql & "                           " & intFilial & vbCrLf
                sSql = sSql & "                          ," & lngIDProduto & vbCrLf
                sSql = sSql & "                          ," & arrVERNIZ02(i, 1) & vbCrLf
                sSql = sSql & "                          ," & arrVERNIZ02(i, 2) & vbCrLf
                sSql = sSql & "                          ," & arrVERNIZ02(i, 3) & vbCrLf
                
                sSql = sSql & "                           )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
     
        
        '' ----------------------------------------------------
        '' Verniz Acabamento
        If IsArray(arrVERNIZACAB) Then
            For i = 1 To UBound(arrVERNIZACAB)
                sSql = "Insert Into SGI_VERNIZPRODACAB (" & vbCrLf
                sSql = sSql & "                          SGI_FILIAL" & vbCrLf
                sSql = sSql & "                         ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                         ,SGI_PRODUTO" & vbCrLf
                sSql = sSql & "                         ,SGI_CODUNIMED" & vbCrLf
                sSql = sSql & "                         ,SGI_QTDE" & vbCrLf
                
                sSql = sSql & "                 ) Values (" & vbCrLf
                sSql = sSql & "                           " & intFilial & vbCrLf
                sSql = sSql & "                          ," & lngIDProduto & vbCrLf
                sSql = sSql & "                          ," & arrVERNIZACAB(i, 1) & vbCrLf
                sSql = sSql & "                          ," & arrVERNIZACAB(i, 2) & vbCrLf
                sSql = sSql & "                          ," & arrVERNIZACAB(i, 3) & vbCrLf
                
                sSql = sSql & "                           )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
        
        '' ----------------------------------------------------
        '' Esmalte
        If IsArray(arrESMALTE) Then
            For i = 1 To UBound(arrESMALTE)
                sSql = "Insert Into SGI_ESMALTEPROD (" & vbCrLf
                sSql = sSql & "                          SGI_FILIAL" & vbCrLf
                sSql = sSql & "                         ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                         ,SGI_PRODUTO" & vbCrLf
                sSql = sSql & "                         ,SGI_CODUNIMED" & vbCrLf
                sSql = sSql & "                         ,SGI_QTDE" & vbCrLf
                
                sSql = sSql & "                 ) Values (" & vbCrLf
                sSql = sSql & "                           " & intFilial & vbCrLf
                sSql = sSql & "                          ," & lngIDProduto & vbCrLf
                sSql = sSql & "                          ," & arrESMALTE(i, 1) & vbCrLf
                sSql = sSql & "                          ," & arrESMALTE(i, 2) & vbCrLf
                sSql = sSql & "                          ," & arrESMALTE(i, 3) & vbCrLf
                
                sSql = sSql & "                           )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
        '' ----------------------------------------------------
     
        '' Verniz Externo
        If IsArray(arrVERNIZEXT) Then
            For i = 1 To UBound(arrVERNIZEXT)
                sSql = "Insert Into SGI_VERNIZEXT (" & vbCrLf
                sSql = sSql & "                          SGI_FILIAL" & vbCrLf
                sSql = sSql & "                         ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                         ,SGI_PRODUTO" & vbCrLf
                sSql = sSql & "                         ,SGI_CODUNIMED" & vbCrLf
                sSql = sSql & "                         ,SGI_QTDE" & vbCrLf
                
                sSql = sSql & "                 ) Values (" & vbCrLf
                sSql = sSql & "                           " & intFilial & vbCrLf
                sSql = sSql & "                          ," & lngIDProduto & vbCrLf
                sSql = sSql & "                          ," & arrVERNIZEXT(i, 1) & vbCrLf
                sSql = sSql & "                          ," & arrVERNIZEXT(i, 2) & vbCrLf
                sSql = sSql & "                          ," & arrVERNIZEXT(i, 3) & vbCrLf
                
                sSql = sSql & "                           )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
        
        
        '' ----------------------------------------------------
        '' Vedante Fechamento
        If IsArray(arrVEDANTE) Then
            For i = 1 To UBound(arrVEDANTE)
                sSql = "Insert Into SGI_VEDANTEPROD (" & vbCrLf
                sSql = sSql & "                             SGI_FILIAL" & vbCrLf
                sSql = sSql & "                            ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                            ,SGI_PRODUTO" & vbCrLf
                sSql = sSql & "                            ,SGI_CODUNIMED" & vbCrLf
                sSql = sSql & "                            ,SGI_QTDE" & vbCrLf
                
                sSql = sSql & "                   ) Values (" & vbCrLf
                sSql = sSql & "                             " & intFilial & vbCrLf
                sSql = sSql & "                            ," & lngIDProduto & vbCrLf
                sSql = sSql & "                            ," & arrVEDANTE(i, 1) & vbCrLf
                sSql = sSql & "                            ," & arrVEDANTE(i, 2) & vbCrLf
                sSql = sSql & "                            ," & arrVEDANTE(i, 3) & vbCrLf
                
                sSql = sSql & "                            )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
        '' ----------------------------------------------------
     
        '' ----------------------------------------------------
        '' Tampa Pressão
        If IsArray(arrTAMPAPRESS) Then
            For i = 1 To UBound(arrTAMPAPRESS)
                sSql = "Insert Into SGI_FECHAMCAMPOS (" & vbCrLf
                sSql = sSql & "                              SGI_FILIAL" & vbCrLf
                sSql = sSql & "                             ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                             ,SGI_CAMPO" & vbCrLf
                sSql = sSql & "                             ,SGI_INDICE" & vbCrLf
                sSql = sSql & "                             ,SGI_VALOR" & vbCrLf
                sSql = sSql & "                    ) Values (" & vbCrLf
                sSql = sSql & "                              " & intFilial & vbCrLf
                sSql = sSql & "                             ," & lngIDProduto & vbCrLf
                sSql = sSql & "                             ," & intTampaPressao
                sSql = sSql & "                             ," & i & vbCrLf
                
                strValor = Replace(arrTAMPAPRESS(i), ".", "")
                strValor = Replace(Trim(strValor), ",", ".")
                sSql = sSql & "                             ," & strValor & vbCrLf
                
                sSql = sSql & "                             )"
            
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
     
        '' ----------------------------------------------------
        '' Batoque Retratil
        If IsArray(arrBATRETRATI) Then
            For i = 1 To UBound(arrBATRETRATI)
                sSql = "Insert Into SGI_FECHAMCAMPOS (" & vbCrLf
                sSql = sSql & "                              SGI_FILIAL" & vbCrLf
                sSql = sSql & "                             ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                             ,SGI_CAMPO" & vbCrLf
                sSql = sSql & "                             ,SGI_INDICE" & vbCrLf
                sSql = sSql & "                             ,SGI_VALOR" & vbCrLf
                sSql = sSql & "                    ) Values (" & vbCrLf
                sSql = sSql & "                              " & intFilial & vbCrLf
                sSql = sSql & "                             ," & lngIDProduto & vbCrLf
                sSql = sSql & "                             ," & intBatoqueRetratil
                sSql = sSql & "                             ," & i & vbCrLf
                
                strValor = Replace(arrBATRETRATI(i), ".", "")
                strValor = Replace(Trim(strValor), ",", ".")
                sSql = sSql & "                             ," & strValor & vbCrLf
                
                sSql = sSql & "                             )"
            
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
     
        '' ----------------------------------------------------
        '' Batoque Plastico
        If IsArray(arrBATPLASTIC) Then
            For i = 1 To UBound(arrBATPLASTIC)
                sSql = "Insert Into SGI_FECHAMCAMPOS (" & vbCrLf
                sSql = sSql & "                              SGI_FILIAL" & vbCrLf
                sSql = sSql & "                             ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                             ,SGI_CAMPO" & vbCrLf
                sSql = sSql & "                             ,SGI_INDICE" & vbCrLf
                sSql = sSql & "                             ,SGI_VALOR" & vbCrLf
                sSql = sSql & "                    ) Values (" & vbCrLf
                sSql = sSql & "                              " & intFilial & vbCrLf
                sSql = sSql & "                             ," & lngIDProduto & vbCrLf
                sSql = sSql & "                             ," & intBatoquePlastico
                sSql = sSql & "                             ," & i & vbCrLf
                
                strValor = Replace(arrBATPLASTIC(i), ".", "")
                strValor = Replace(Trim(strValor), ",", ".")
                sSql = sSql & "                             ," & strValor & vbCrLf
                
                sSql = sSql & "                             )"
            
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
     
     
        '' ----------------------------------------------------
        '' Tampa Visor
        If IsArray(arrTAMPAVISOR) Then
            For i = 1 To UBound(arrTAMPAVISOR)
                sSql = "Insert Into SGI_FECHAMCAMPOS (" & vbCrLf
                sSql = sSql & "                              SGI_FILIAL" & vbCrLf
                sSql = sSql & "                             ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                             ,SGI_CAMPO" & vbCrLf
                sSql = sSql & "                             ,SGI_INDICE" & vbCrLf
                sSql = sSql & "                             ,SGI_VALOR" & vbCrLf
                sSql = sSql & "                    ) Values (" & vbCrLf
                sSql = sSql & "                              " & intFilial & vbCrLf
                sSql = sSql & "                             ," & lngIDProduto & vbCrLf
                sSql = sSql & "                             ," & intTampaVisor
                sSql = sSql & "                             ," & i & vbCrLf
                
                strValor = Replace(arrTAMPAVISOR(i), ".", "")
                strValor = Replace(Trim(strValor), ",", ".")
                sSql = sSql & "                             ," & strValor & vbCrLf
                
                sSql = sSql & "                             )"
            
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
     
        '' ----------------------------------------------------
        '' Clientes que este produto atende
        If IsArray(arrPRODATECLIE) Then
            For i = 1 To UBound(arrPRODATECLIE)
                sSql = "Insert Into SGI_PRODATECLIE (" & vbCrLf
                sSql = sSql & "                              SGI_FILIAL" & vbCrLf
                sSql = sSql & "                             ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                             ,SGI_IDCLIENTE" & vbCrLf
                sSql = sSql & "                    ) Values (" & vbCrLf
                sSql = sSql & "                              " & intFilial & vbCrLf
                sSql = sSql & "                             ," & lngIDProduto & vbCrLf
                sSql = sSql & "                             ," & arrPRODATECLIE(i)
                sSql = sSql & "                             )"
            
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
        
        
        '' ------------------------
        '' Familias de Máquinas
        If IsArray(arrFAMMAQ) Then
            For i = 1 To UBound(arrFAMMAQ)
                sSql = "Insert Into SGI_PRODFAMMAQ (" & vbCrLf
                sSql = sSql & "                              SGI_FILIAL" & vbCrLf
                sSql = sSql & "                             ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                             ,SGI_CODFAMMAQ" & vbCrLf
                sSql = sSql & "                    ) Values (" & vbCrLf
                sSql = sSql & "                              " & intFilial & vbCrLf
                sSql = sSql & "                             ," & lngIDProduto & vbCrLf
                sSql = sSql & "                             ," & arrFAMMAQ(i)
                sSql = sSql & "                             )"
            
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
     
     
        '' Produto de ENtrada
        If IsArray(arrPRODENTR) Then
            For i = 1 To UBound(arrPRODENTR)
                sSql = ""
            
                sSql = "Insert Into SGI_PRODENTRADA (" & vbCrLf
                sSql = sSql & "                              SGI_FILIAL" & vbCrLf
                sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                             ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                             ,SGI_CODUNIMED" & vbCrLf
                sSql = sSql & "                             ,SGI_QTDE" & vbCrLf
                
                sSql = sSql & "                    ) Values (" & vbCrLf
                sSql = sSql & "                              " & intFilial & vbCrLf
                sSql = sSql & "                             ," & lngIDProduto & vbCrLf
                sSql = sSql & "                             ," & arrPRODENTR(i, 1) & vbCrLf
                sSql = sSql & "                             ," & arrPRODENTR(i, 2) & vbCrLf
                sSql = sSql & "                             ," & arrPRODENTR(i, 3) & vbCrLf
                
                sSql = sSql & "                             )"
            
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
     
                
        strEMPRESA = ""
        If intFILIALPED = 1 Then strEMPRESA = "_STEEL"
        
        If IsArray(arrPEDIDOS) Then
            For i = 1 To UBound(arrPEDIDOS)
               If Len(Trim(arrPEDIDOS(i, 1))) > 0 Then
                   sSql = ""
                
                   sSql = "Update SGI_CADPEDVENDI" & strEMPRESA & " Set SGI_CODPROD = '" & Format(lngCodLinProd, "###000") & "." & Format(lngCodClie, "####0000") & "." & Format(lngCodRotulo, "##00") & "." & Trim(Str(lngDigVerif)) & "'" & vbCrLf
                   sSql = sSql & "                     Where" & vbCrLf
                   sSql = sSql & "                           SGI_FILIAL    = " & intFilial & vbCrLf
                   sSql = sSql & "                       And SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
                   sSql = sSql & "                       And SGI_CODIGO    = " & arrPEDIDOS(i, 1)
                
                   BGRV.CommandText = sSql
                   BGRV.Execute
               End If
           Next i
        End If
        
        
        '' Dia 07/08/2017
        '' Estrutura de Produtos
        '' Grava A Extrutura do Produto
        
        sSql = ""
        
        sSql = "Delete From SGI_LISTAMATPROD" & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL  = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODSTRU = " & lngIDProduto
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        For i = 1 To UBound(arrPROVARV)
        
            If arrPROVARV(i).lngProdutoID > 0 Then
            
                sSql = ""
                
                sSql = "Insert Into SGI_LISTAMATPROD (" & vbCrLf
                sSql = sSql & "                              SGI_FILIAL" & vbCrLf
                sSql = sSql & "                             ,SGI_CODSTRU" & vbCrLf
                sSql = sSql & "                             ,SGI_CODPAI" & vbCrLf
                sSql = sSql & "                             ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                             ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                             ,SGI_IDPRODPAI" & vbCrLf
                sSql = sSql & "                             ,SGI_CODUNIMED" & vbCrLf
                sSql = sSql & "                             ,SGI_QTDE" & vbCrLf
                
                sSql = sSql & "                ) Values (" & vbCrLf
                sSql = sSql & "                          " & intFilial & vbCrLf
                sSql = sSql & "                         ," & lngIDProduto & vbCrLf
                sSql = sSql & "                         ," & arrPROVARV(i).lngCODPAI & vbCrLf
                sSql = sSql & "                         ," & arrPROVARV(i).lngCODIGO & vbCrLf
                sSql = sSql & "                         ," & arrPROVARV(i).lngProdutoID & vbCrLf
                sSql = sSql & "                         ," & arrPROVARV(i).lngProdutoIDPai & vbCrLf
                sSql = sSql & "                         ," & arrPROVARV(i).lngCodUniMed & vbCrLf
                sSql = sSql & "                         ," & arrPROVARV(i).strQTDCONS & vbCrLf
                
                sSql = sSql & "                         )"
        
                BGRV.CommandText = sSql
                BGRV.Execute
            
            End If
        
        Next i
        
        
     
     ElseIf strOPER = "E" Then
        
        '' Apaga Tabela de Converção
        sSql = "Delete From SGI_CADPRODUNIDCONV " & vbCrLf
        sSql = sSql & "                          Where " & vbCrLf
        sSql = sSql & "                                SGI_FILIAL  = " & intFilial & vbCrLf
        sSql = sSql & "                            And SGI_IDPRODUTO = " & lngIDProduto
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        
        '' --------------------------
        '' Apaga o Especificação Técnica
        sSql = "Delete from SGI_PRODESPTEC " & vbCrLf
        sSql = sSql & "       Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "         And SGI_FILIAL  = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------
        '' Apaga o Cardex
        sSql = "Delete from SGI_CARDEXEST"
        sSql = sSql & "       Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "         And SGI_FILIAL  = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        
        '' --------------------------
        '' Apaga Processos
''        sSql = "Delete from SGI_CADPROPROC " & vbCrLf
''        sSql = sSql & "              Where SGI_CODPROD = '" & strCodigoProd & "'" & vbCrLf
''        sSql = sSql & "                And SGI_FILIAL  = " & intFilial
        
''        BGRV.CommandText = sSql
''        BGRV.Execute
        
        '' --------------------------
        '' Apaga Ferramentas
''        sSql = "Delete from SGI_CADFERRPROC " & vbCrLf
''        sSql = sSql & "      Where SGI_CODPROD = '" & strCodigoProd & "'" & vbCrLf
''        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
''        BGRV.CommandText = sSql
''        BGRV.Execute
        
        '' --------------------------
        '' Apaga Tipos de Orçamentos
        sSql = "Delete from SGI_CADPRODTIPORCA " & vbCrLf
        sSql = sSql & "      Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------
        '' Apaga Variante de Inspeção de Processo
        sSql = "Delete from SGI_CADVAINSPPROC " & vbCrLf
        sSql = sSql & "      Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------
        '' Apaga Variante de Recebimento
        sSql = "Delete from SGI_CADVARECPROD " & vbCrLf
        sSql = sSql & "      Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------
        '' Apaga Variante de Inspeção Final
        sSql = "Delete from SGI_CADVAINSPPROD " & vbCrLf
        sSql = sSql & "      Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------
        '' Apaga Pesos de Coeficiente
        sSql = "Delete From SGI_CADPARPRODPADRAO " & vbCrLf
        sSql = sSql & "      Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "        And SGI_FILIAL  = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------
        '' Apaga Parâmetros de Coeficiente
        sSql = "Delete From SGI_CADPARCOEF " & vbCrLf
        sSql = sSql & "      Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "        And SGI_FILIAL  = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------
        '' Apaga Cores
        sSql = "Delete From SGI_CORESPROD " & vbCrLf
        sSql = sSql & "      Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "        And SGI_FILIAL  = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------
        '' Verniz
        sSql = "Delete From SGI_VERNIZPROD " & vbCrLf
        sSql = sSql & "      Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------
        '' Verniz 02
        sSql = "Delete From SGI_VERNIZPROD02 " & vbCrLf
        sSql = sSql & "      Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------
        '' Verniz Acabamento
        sSql = "Delete From SGI_VERNIZPRODACAB " & vbCrLf
        sSql = sSql & "      Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------
        '' Esmalte
        sSql = "Delete From SGI_ESMALTEPROD " & vbCrLf
        sSql = sSql & "      Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        
        '' --------------------------
        '' Verniz Externo
        sSql = "Delete From SGI_VERNIZEXT " & vbCrLf
        sSql = sSql & "      Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------
        '' Vedante Fechamento
        sSql = "Delete From SGI_VEDANTEPROD " & vbCrLf
        sSql = sSql & "      Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------
        '' Fechamento
        sSql = "Delete From SGI_FECHAMCAMPOS " & vbCrLf
        sSql = sSql & "      Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
       
        '' --------------------------
        '' Cliente que este produto atende
        sSql = "Delete From SGI_PRODATECLIE " & vbCrLf
        sSql = sSql & "      Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------
        '' Familias de Máquinas
        sSql = "Delete From SGI_PRODFAMMAQ " & vbCrLf
        sSql = sSql & "      Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------
        '' Produtos de Entrada
        sSql = "Delete From SGI_PRODENTRADA " & vbCrLf
        sSql = sSql & "      Where SGI_CODIGO = " & lngIDProduto & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' --------------------------
        '' Apaga Produto
        sSql = "Delete from SGI_CADPRODUTO " & vbCrLf
        sSql = sSql & "      Where SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
        sSql = sSql & "        And SGI_FILIAL = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
          
     ElseIf strOPER = "AS" Then
     
        sSql = "Update SGI_CADPRODUTO Set " & vbCrLf
        
        sSql = sSql & "                          SGI_STATUS     = " & intStatus & vbCrLf
        sSql = sSql & "                         ,SGI_PRODNOVO   = " & intPRODNOVO & vbCrLf
        sSql = sSql & "                         ,SGI_ALTFOTOLIT = " & intFOTALTSN & vbCrLf
        
        sSql = sSql & "              Where " & vbCrLf
        sSql = sSql & "                    SGI_FILIAL    = " & intFilial & vbCrLf
        sSql = sSql & "                And SGI_IDPRODUTO = " & lngIDProduto
     
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


Public Sub PreencheComboProd(Combo As Variant)

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
     
    Do While Not BREC.EOF
       Combo.AddItem Trim(BREC!SGI_CODIGO) & Space(10 - Len(Trim(BREC!SGI_CODIGO))) & " - " & Trim(BREC!SGI_DESCRICAO)
       BREC.MoveNext
    Loop
    
    If Combo.ListCount > 0 Then Combo.ListIndex = 0
    
    BREC.Close

    
End Sub

Public Function Carrega_campos() As Boolean

   Carrega_campos = False
   
   Dim intReG As Integer
   Dim i      As Integer
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC.EOF Then
      
      
      strIPI = ""
      If Not IsNull(BREC!SGI_IPI) Then strIPI = BREC!SGI_IPI
      
      If Not IsNull(BREC!SGI_ARQUIVOIMG) Then strCAMINHO = Trim(BREC!SGI_ARQUIVOIMG)
      If Not IsNull(BREC!SGI_CODLINPROD) Then lngCodLinProd = BREC!SGI_CODLINPROD
      If Not IsNull(BREC!SGI_CODCLIE) Then lngCodClie = BREC!SGI_CODCLIE
      If Not IsNull(BREC!SGI_CODROTULO) Then lngCodRotulo = BREC!SGI_CODROTULO
      If Not IsNull(BREC!SGI_DIGVERIF) Then lngDigVerif = BREC!SGI_DIGVERIF
      If Not IsNull(BREC!SGI_NATURALSN) Then intNATURALSIMNAO = BREC!SGI_NATURALSN
      If Not IsNull(BREC!SGI_STATUS) Then intStatus = BREC!SGI_STATUS
      If IsNull(BREC!SGI_STATUS) Then intStatus = 1
      If Not IsNull(BREC!SGI_ALTFOTOLIT) Then intFOTALTSN = BREC!SGI_ALTFOTOLIT
        
      lngIDProduto = BREC!SGI_IDPRODUTO
      imgIMAGEM = BREC!SGI_IMAGEM
      
      If Not IsNull(BREC!SGI_CODIGO) Then strCodigoProd = BREC!SGI_CODIGO
      ''strCodigoProdANT = BREC!SGI_CODIGO
      strDescriProd = BREC!SGI_DESCRICAO
      intEspProduto = BREC!SGI_CODESPECIE
      If Not IsNull(BREC!SGI_CODGPROD) Then intCODGRUPPROD = BREC!SGI_CODGPROD
      If Not IsNull(BREC!SGI_CODSUBGPROD) Then intCODSUBGPROD = BREC!SGI_CODSUBGPROD
      intTIPPRODUTO = BREC!SGI_CODTIPO
      intUnidade = BREC!SGI_UNIDMEDIDA
      
      If Not IsNull(BREC!SGI_DESCPROD) Then strDESCEQUIP = BREC!SGI_DESCPROD
      If Not IsNull(BREC!SGI_PROCEDENCIA) Then strPROCEDEN = BREC!SGI_PROCEDENCIA
      
      If Not IsNull(BREC!SGI_CODCLASS) Then lngCLASFISC = BREC!SGI_CODCLASS
      dblSaldo = IIf(IsNull(BREC!SGI_SALDO) = False, BREC!SGI_SALDO, 0)
      dblEstMin = IIf(IsNull(BREC!SGI_ESTMINIMO) = False, BREC!SGI_ESTMINIMO, 0)
      
      dtDataCadast = Format(BREC!SGI_DTCADASTRO, "DD/MM/YYYY")
      If Not IsNull(BREC!SGI_DTULTMOV) Then dtDataUltMov = BREC!SGI_DTULTMOV
      
      If Not IsNull(BREC!SGI_PRECOPROD) Then dblPRCPROD = BREC!SGI_PRECOPROD
      If Not IsNull(BREC!SGI_PRECOMED) Then dblPRCMED = BREC!SGI_PRECOMED
      If Not IsNull(BREC!SGI_ATUAUTOM) Then intATAUTOMSN = BREC!SGI_ATUAUTOM
      
      If Not IsNull(BREC!SGI_PROCACRES) Then curPORCACRES = BREC!SGI_PROCACRES
      If Not IsNull(BREC!SGI_PRCCUSTO) Then curPRCCUSTO = BREC!SGI_PRCCUSTO
      
      If Not IsNull(BREC!SGI_PRODUTOTIPO) Then intPRODUTOTIPO = BREC!SGI_PRODUTOTIPO
      If Not IsNull(BREC!SGI_PRODUTOESTILO) Then intESTILOPROD = BREC!SGI_PRODUTOESTILO
      
      If Not IsNull(BREC!SGI_PESOUNIT) Then curPESOUNIT = BREC!SGI_PESOUNIT
      
      lngTABCONV = 0
      If Not IsNull(BREC!SGI_CODTABCONV) Then lngTABCONV = BREC!SGI_CODTABCONV
      
      If Not IsNull(BREC!SGI_CODPRODFORN) Then strCODPROFORNEC = Trim(BREC!SGI_CODPRODFORN)
      
      If Not IsNull(BREC!SGI_RUA) Then strRUA = BREC!SGI_RUA
      If Not IsNull(BREC!SGI_BOX) Then strBOX = BREC!SGI_BOX
      If Not IsNull(BREC!SGI_PRATELEIRA) Then strPRATELEIRA = BREC!SGI_PRATELEIRA
      If Not IsNull(BREC!SGI_CODEAN) Then strCODEAN = BREC!SGI_CODEAN
      If Not IsNull(BREC!SGI_DISTESTOQUE) Then curDISTANCIA = BREC!SGI_DISTESTOQUE
      intPADRAO = BREC!SGI_PADRAO
      
      If Not IsNull(BREC!SGI_CORRPRODPADR) Then curCORRELACAO = BREC!SGI_CORRPRODPADR
      If Not IsNull(BREC!SGI_PORCCORR) Then curPORCCORR = BREC!SGI_PORCCORR
      If Not IsNull(BREC!SGI_CADFAMMAQ) Then lngFAMMAQUINA = BREC!SGI_CADFAMMAQ
      If Not IsNull(BREC!SGI_QTDEPORFOLHA) Then lngQTDEPORFOLHA = BREC!SGI_QTDEPORFOLHA
      If Not IsNull(BREC!SGI_QTDEPASSADA) Then lngQtdePassadas = BREC!SGI_QTDEPASSADA
      
      strSEGUENCIA = "Null"
      If Not IsNull(BREC!SGI_SEGUENCIA) Then strSEGUENCIA = Format(BREC!SGI_SEGUENCIA, "####0000")
      
      
      curCUBAGEN = BREC!SGI_CUBAGEN
      
      curMetros = BREC!SGI_METROS
      curLargura = BREC!SGI_LARGURA
      curGRAMATURA2 = BREC!SGI_GRAMATURA2
      
      
      lngVernCorpo = 0
      lngVernTampa = 0
      lngVernFundo = 0
      lngVernArgola = 0
      lngFechTampaFuro = -1
      If Not IsNull(BREC!SGI_VernCorpo) Then lngVernCorpo = BREC!SGI_VernCorpo
      If Not IsNull(BREC!SGI_VernTampa) Then lngVernTampa = BREC!SGI_VernTampa
      If Not IsNull(BREC!SGI_VernFundo) Then lngVernFundo = BREC!SGI_VernFundo
      If Not IsNull(BREC!SGI_VernArgola) Then lngVernArgola = BREC!SGI_VernArgola
      If Not IsNull(BREC!SGI_FechTampaFuro) Then lngFechTampaFuro = BREC!SGI_FechTampaFuro
      
      lngEspessCorpo = 0
      lngEspessTampa = 0
      lngEspessFundo = 0
      lngEspessArgola = 0
      If Not IsNull(BREC!SGI_CorpoEspess) Then lngEspessCorpo = BREC!SGI_CorpoEspess
      If Not IsNull(BREC!SGI_TampaEspess) Then lngEspessTampa = BREC!SGI_TampaEspess
      If Not IsNull(BREC!SGI_FundoEspess) Then lngEspessFundo = BREC!SGI_FundoEspess
      If Not IsNull(BREC!SGI_ArgolaEspess) Then lngEspessArgola = BREC!SGI_ArgolaEspess
      
      lngRevestCorpo = 0
      lngRevestCorpo2 = 0
      lngRevestTampa = 0
      lngRevestTampa2 = 0
      lngRevestFundo = 0
      lngRevestFundo2 = 0
      lngRevestArgola = 0
      lngRevestArgola2 = 0
      
      If Not IsNull(BREC!SGI_CorpoRevest) Then lngRevestCorpo = BREC!SGI_CorpoRevest
      If Not IsNull(BREC!SGI_CorpoRevest2) Then lngRevestCorpo2 = BREC!SGI_CorpoRevest2
      If Not IsNull(BREC!SGI_TampaRevest) Then lngRevestTampa = BREC!SGI_TampaRevest
      If Not IsNull(BREC!SGI_TampaRevest2) Then lngRevestTampa2 = BREC!SGI_TampaRevest2
      If Not IsNull(BREC!SGI_FundoRevest) Then lngRevestFundo = BREC!SGI_FundoRevest
      If Not IsNull(BREC!SGI_FundoRevest2) Then lngRevestFundo2 = BREC!SGI_FundoRevest2
      If Not IsNull(BREC!SGI_ArgolaRevest) Then lngRevestArgola = BREC!SGI_ArgolaRevest
      If Not IsNull(BREC!SGI_ArgolaRevest2) Then lngRevestArgola2 = BREC!SGI_ArgolaRevest2
      
      '' ========================================
      '' Montagem
      lngAlcaPlastica = 0
      lngAlcaFerro = 0
      intPipeta = 0
      intAzelha = 0
      lngVerCat = 0
      lngColEsp = 0
      If Not IsNull(BREC!SGI_AlcaPlastica) Then lngAlcaPlastica = BREC!SGI_AlcaPlastica
      If Not IsNull(BREC!SGI_AlcaFerro) Then lngAlcaFerro = BREC!SGI_AlcaFerro
      If Not IsNull(BREC!SGI_Pipeta) Then intPipeta = BREC!SGI_Pipeta
      If Not IsNull(BREC!SGI_AZELHA) Then intAzelha = BREC!SGI_AZELHA
      If Not IsNull(BREC!SGI_VERCAT) Then lngVerCat = BREC!SGI_VERCAT
      If Not IsNull(BREC!SGI_COLESP) Then lngColEsp = BREC!SGI_COLESP
      '' ========================================
      
      lngFechSoldaAgrafado = -1
      intFechSA = 0
      If Not IsNull(BREC!SGI_FechSoldaAgrafado) Then lngFechSoldaAgrafado = BREC!SGI_FechSoldaAgrafado
      If Not IsNull(BREC!SGI_FechSA) Then intFechSA = BREC!SGI_FechSA
      
      intTipoPalletGranel = 0
      lngQtdePalletGranel = 0
      If Not IsNull(BREC!SGI_TIPOEMPILHAMENTO) Then intTipoPalletGranel = BREC!SGI_TIPOEMPILHAMENTO
      If Not IsNull(BREC!SGI_QTDEEMPILHAMENTO) Then lngQtdePalletGranel = BREC!SGI_QTDEEMPILHAMENTO
      
      If Not IsNull(BREC!SGI_EMITLAUDO) Then intEMITLAUDO = BREC!SGI_EMITLAUDO
      
      intALCAGALAO = 0
      If Not IsNull(BREC!SGI_ALCAGALAO) Then intALCAGALAO = BREC!SGI_ALCAGALAO
      
      intNECKIN = 0
      If Not IsNull(BREC!SGI_NECKIN) Then intNECKIN = BREC!SGI_NECKIN
      If Not IsNull(BREC!SGI_FILIALPED) Then intFILIALPED = BREC!SGI_FILIALPED
      
      intDIMPADRAO = BREC!SGI_DIMPADRAO
      If Not IsNull(BREC!SGI_DESENV) Then strDESENV = BREC!SGI_DESENV
      If Not IsNull(BREC!SGI_ALTURA) Then strALTURA = BREC!SGI_ALTURA
      
      If Not IsNull(BREC!SGI_COMPLEMENTO) Then strCOMPLEMENTO = Trim(BREC!SGI_COMPLEMENTO)
      
      If Not IsNull(BREC!SGI_CODPRODENTR) Then strCODPRODENTR = Trim(BREC!SGI_CODPRODENTR)
      If Not IsNull(BREC!SGI_QTDCORPSPADRAOSN) Then intQTDCORPSPADRAOSN = BREC!SGI_QTDCORPSPADRAOSN
      
      intPRODNOVO = 0
      If Not IsNull(BREC!SGI_PRODNOVO) Then intPRODNOVO = BREC!SGI_PRODNOVO
      
      If Not IsNull(BREC!SGI_USARPADRPALSN) Then intPALLHETPADRAO = BREC!SGI_USARPADRPALSN
      
      Carrega_campos = True
   End If
   
   BREC.Close
   
   If Carrega_campos = False Then Exit Function
   
   '' ---------------
   
   BREC3.ActiveConnection = adoBanco_Dados
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_PRODESPTEC " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial
   sSql = sSql & "   And SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
   
   BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC3.EOF Then
   
      intReG = 0
      Do While Not BREC3.EOF
         intReG = intReG + 1
         BREC3.MoveNext
      Loop
      
      
      ReDim arrESPTECPROD(1 To intReG) As Variant
      
      intReG = 1
      BREC3.MoveFirst
      Do While Not BREC3.EOF
         arrESPTECPROD(intReG) = BREC3!SGI_CODESPTEC
         intReG = intReG + 1
         BREC3.MoveNext
      Loop
   
   End If
   
   BREC3.Close

   '' Carrega Lista de Material
   BREC2.ActiveConnection = adoBanco_Dados
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_LISTAMAT " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
   
   BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC2.EOF Then
   
      intReG = 0
      Do While Not BREC2.EOF
         intReG = intReG + 1
         BREC2.MoveNext
      Loop
      
      ReDim arrProdLST(1 To intReG, 1 To 6) As String
      
      intReG = 1
      BREC2.MoveFirst
      Do While Not BREC2.EOF
      
         arrProdLST(intReG, 1) = BREC2!SGI_PRODLST
         arrProdLST(intReG, 2) = BREC2!SGI_QTDE
         
         arrProdLST(intReG, 3) = ""
         If Not IsNull(BREC2!SGI_UNIDCONS) Then arrProdLST(intReG, 3) = BREC2!SGI_UNIDCONS
         arrProdLST(intReG, 4) = ""
         If Not IsNull(BREC2!SGI_CODTABCONV) Then arrProdLST(intReG, 4) = IIf(BREC2!SGI_CODTABCONV <> 0, BREC2!SGI_CODTABCONV, "")
         
         If Not IsNull(BREC2!SGI_CUSTOUNIT) Then arrProdLST(intReG, 5) = BREC2!SGI_CUSTOUNIT
         If Not IsNull(BREC2!SGI_CUSTOTOTAL) Then arrProdLST(intReG, 6) = BREC2!SGI_CUSTOTOTAL
         
         intReG = intReG + 1
         BREC2.MoveNext
      Loop
   
   End If
   
   BREC2.Close
   
   '' Pega Processo
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADPROPROC " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_CODPROD = '" & strCodigoProd & "'" & vbCrLf
   
   BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC3.EOF Then
   
      intReG = 0
      Do While Not BREC3.EOF
         intReG = intReG + 1
         BREC3.MoveNext
      Loop
      
      ReDim arrPROCESSOS(1 To intReG, 1 To 4) As String
      
      BREC3.MoveFirst
      intReG = 1
      Do While Not BREC3.EOF
         arrPROCESSOS(intReG, 1) = Trim(Str(BREC3!SGI_CODPROC))
         arrPROCESSOS(intReG, 2) = CInt(BREC3!SGI_TEMPO)
         arrPROCESSOS(intReG, 3) = Format(BREC3!SGI_VALORHORA, "#,##0.00")
         arrPROCESSOS(intReG, 4) = Format(BREC3!SGI_TOTVALOR, "#,##0.00")
         intReG = intReG + 1
         BREC3.MoveNext
      Loop
   
   End If
   BREC3.Close


   '' Pega Tipos de Orçamentos
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADPRODTIPORCA " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
   
   BREC5.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC5.EOF Then
   
      intReG = 0
      Do While Not BREC5.EOF
         intReG = intReG + 1
         BREC5.MoveNext
      Loop
      
      ReDim arrPRODTIPORCA(1 To intReG) As String
      
      BREC5.MoveFirst
      intReG = 1
      Do While Not BREC5.EOF
         arrPRODTIPORCA(intReG) = Trim(Str(BREC5!SGI_CODTIPORCA))
         intReG = intReG + 1
         BREC5.MoveNext
      Loop
   
   End If
   BREC5.Close

   '' Pega Variaveis de Recebimento
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADVARECPROD " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL   = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_IDPRODUTO  = " & lngIDProduto & vbCrLf
   sSql = sSql & "   And SGI_TIPOINSP = 0" & vbCrLf
   
   BREC6.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC6.EOF Then
   
      intReG = 0
      Do While Not BREC6.EOF
         intReG = intReG + 1
         BREC6.MoveNext
      Loop
      
      ReDim arrPRODVAREC(1 To intReG, 1 To 4) As String
      
      BREC6.MoveFirst
      intReG = 1
      Do While Not BREC6.EOF
         arrPRODVAREC(intReG, 1) = Trim(Str(BREC6!SGI_CODVAR))
         arrPRODVAREC(intReG, 2) = Trim(Str(BREC6!SGI_COEF))
         arrPRODVAREC(intReG, 3) = Trim(Str(BREC6!SGI_COEFP))
         arrPRODVAREC(intReG, 4) = Trim(Str(BREC6!SGI_COEFN))
         intReG = intReG + 1
         BREC6.MoveNext
      Loop
   
   End If
   BREC6.Close


   '' Pega Atributos de Recebimento
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADVARECPROD " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL   = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_IDPRODUTO  = " & lngIDProduto & vbCrLf
   sSql = sSql & "   And SGI_TIPOINSP = 1" & vbCrLf
   
   BREC7.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC7.EOF Then
   
      intReG = 0
      Do While Not BREC7.EOF
         intReG = intReG + 1
         BREC7.MoveNext
      Loop
      
      ReDim arrPRODATRREC(1 To intReG, 1 To 3) As String
      
      BREC7.MoveFirst
      intReG = 1
      Do While Not BREC7.EOF
         arrPRODATRREC(intReG, 1) = Trim(Str(BREC7!SGI_CODVAR))
         arrPRODATRREC(intReG, 2) = Trim(BREC7!SGI_CONATRIB)
         arrPRODATRREC(intReG, 3) = Trim(BREC7!SGI_ESPATRIB)
         intReG = intReG + 1
         BREC7.MoveNext
      Loop
   
   End If
   BREC7.Close
      
   '' Pega Variantes de Inspeção
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADVAINSPPROD " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
   sSql = sSql & "   And SGI_TIPOINSP = 0" & vbCrLf
   
   BREC8.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC8.EOF Then
   
      intReG = 0
      Do While Not BREC8.EOF
         intReG = intReG + 1
         BREC8.MoveNext
      Loop
      
      ReDim arrPRODVAINSP(1 To intReG, 1 To 4) As String
      
      BREC8.MoveFirst
      intReG = 1
      Do While Not BREC8.EOF
         arrPRODVAINSP(intReG, 1) = Trim(Str(BREC8!SGI_CODVAR))
         arrPRODVAINSP(intReG, 2) = Trim(Str(BREC8!SGI_COEF))
         arrPRODVAINSP(intReG, 3) = Trim(Str(BREC8!SGI_COEFP))
         arrPRODVAINSP(intReG, 4) = Trim(Str(BREC8!SGI_COEFN))
         intReG = intReG + 1
         BREC8.MoveNext
      Loop
   
   End If
   BREC8.Close
   

   '' Pega Variantes de Inspeção de Processos
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADVAINSPPROC " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
   sSql = sSql & "   And SGI_TIPOINSP = 0" & vbCrLf
   
   BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC10.EOF Then
   
      intReG = 0
      Do While Not BREC10.EOF
         intReG = intReG + 1
         BREC10.MoveNext
      Loop
      
      ReDim arrPRODVAPROC(1 To intReG, 1 To 5) As String
      
      BREC10.MoveFirst
      intReG = 1
      Do While Not BREC10.EOF
         arrPRODVAPROC(intReG, 1) = Trim(Str(BREC10!SGI_CODVAR))
         arrPRODVAPROC(intReG, 2) = Trim(Str(BREC10!SGI_COEF))
         arrPRODVAPROC(intReG, 3) = Trim(Str(BREC10!SGI_COEFP))
         arrPRODVAPROC(intReG, 4) = Trim(Str(BREC10!SGI_COEFN))
         arrPRODVAPROC(intReG, 5) = Trim(Str(BREC10!SGI_CODPROC))
         intReG = intReG + 1
         BREC10.MoveNext
      Loop
   
   End If
   BREC10.Close
   
   '' Carrega os Padrao
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADPARPRODPADRAO " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL  = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
   
   BREC12.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC12.EOF Then
      intReG = 0
      Do While Not BREC12.EOF
         intReG = intReG + 1
         BREC12.MoveNext
      Loop
      
      ReDim arrCORRPRODPADR(1 To intReG, 1 To 3) As Variant
      BREC12.MoveFirst
      intReG = 1
      Do While Not BREC12.EOF
         
         arrCORRPRODPADR(intReG, 1) = BREC12!SGI_CODPAR
         arrCORRPRODPADR(intReG, 2) = BREC12!SGI_CODUNI
         arrCORRPRODPADR(intReG, 3) = BREC12!SGI_PESO
         
         intReG = (intReG + 1)
         BREC12.MoveNext
      Loop
   End If
   BREC12.Close
   
   
   '' Carrega Unidades de Conversão
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADPRODUNIDCONV " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL  = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
   
   BREC12.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC12.EOF Then
      intReG = 0
      Do While Not BREC12.EOF
         intReG = intReG + 1
         BREC12.MoveNext
      Loop
      
      ReDim arrUNIDCONV(1 To intReG, 1 To 2) As Variant
      BREC12.MoveFirst
      intReG = 1
      Do While Not BREC12.EOF
         
         arrUNIDCONV(intReG, 1) = BREC12!SGI_CODUNIDADE
         arrUNIDCONV(intReG, 2) = BREC12!SGI_FATOR
         
         intReG = (intReG + 1)
         BREC12.MoveNext
      Loop
   End If
   BREC12.Close
   
   
   '' ----------------------------------------------
   '' Cores
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From" & vbCrLf
   sSql = sSql & "       SGI_CORESPROD" & vbCrLf
   sSql = sSql & " Where" & vbCrLf
   sSql = sSql & "       SGI_FILIAL    = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
   sSql = sSql & "Order By SGI_ITEN"
   
   BREC12.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC12.EOF Then
      intReG = 0
      Do While Not BREC12.EOF
         intReG = intReG + 1
         BREC12.MoveNext
      Loop
   
      ReDim arrCORES(1 To intReG, 1 To 4) As Variant
      BREC12.MoveFirst
      intReG = 1
      Do While Not BREC12.EOF
         arrCORES(intReG, 1) = BREC12!SGI_CODCOR
         If Not IsNull(BREC12!SGI_ITEN) Then arrCORES(intReG, 2) = BREC12!SGI_ITEN
         If Not IsNull(BREC12!SGI_CODUNIMED) Then arrCORES(intReG, 3) = BREC12!SGI_CODUNIMED
         If Not IsNull(BREC12!SGI_QTDE) Then arrCORES(intReG, 4) = BREC12!SGI_QTDE
         intReG = (intReG + 1)
         BREC12.MoveNext
      Loop
   End If
   BREC12.Close
   
   '' ----------------------------------------------
   '' Verniz
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From" & vbCrLf
   sSql = sSql & "       SGI_VERNIZPROD" & vbCrLf
   sSql = sSql & " Where" & vbCrLf
   sSql = sSql & "       SGI_FILIAL    = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_IDPRODUTO = " & lngIDProduto
   
   BREC12.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC12.EOF Then
      intReG = 0
      Do While Not BREC12.EOF
         intReG = intReG + 1
         BREC12.MoveNext
      Loop
   
      ReDim arrVERNIZ(1 To intReG, 1 To 3) As Variant
      BREC12.MoveFirst
      intReG = 1
      Do While Not BREC12.EOF
         arrVERNIZ(intReG, 1) = BREC12!SGI_PRODUTO
         If Not IsNull(BREC12!SGI_CODUNIMED) Then arrVERNIZ(intReG, 2) = BREC12!SGI_CODUNIMED
         If Not IsNull(BREC12!SGI_QTDE) Then arrVERNIZ(intReG, 3) = BREC12!SGI_QTDE
         
         intReG = (intReG + 1)
         BREC12.MoveNext
      Loop
   End If
   BREC12.Close
   
   '' ----------------------------------------------
   '' Verniz 02
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From" & vbCrLf
   sSql = sSql & "       SGI_VERNIZPROD02" & vbCrLf
   sSql = sSql & " Where" & vbCrLf
   sSql = sSql & "       SGI_FILIAL    = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_IDPRODUTO = " & lngIDProduto
   
   BREC12.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC12.EOF Then
      intReG = 0
      Do While Not BREC12.EOF
         intReG = intReG + 1
         BREC12.MoveNext
      Loop
   
      ReDim arrVERNIZ02(1 To intReG, 1 To 3) As Variant
      BREC12.MoveFirst
      intReG = 1
      Do While Not BREC12.EOF
         arrVERNIZ02(intReG, 1) = BREC12!SGI_PRODUTO
         If Not IsNull(BREC12!SGI_CODUNIMED) Then arrVERNIZ02(intReG, 2) = BREC12!SGI_CODUNIMED
         If Not IsNull(BREC12!SGI_QTDE) Then arrVERNIZ02(intReG, 3) = BREC12!SGI_QTDE
         
         intReG = (intReG + 1)
         BREC12.MoveNext
      Loop
   End If
   BREC12.Close
   
   '' ----------------------------------------------
   '' Verniz Acabamento
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From" & vbCrLf
   sSql = sSql & "       SGI_VERNIZPRODACAB" & vbCrLf
   sSql = sSql & " Where" & vbCrLf
   sSql = sSql & "       SGI_FILIAL    = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_IDPRODUTO = " & lngIDProduto
   
   BREC12.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC12.EOF Then
      intReG = 0
      Do While Not BREC12.EOF
         intReG = intReG + 1
         BREC12.MoveNext
      Loop
   
      ReDim arrVERNIZACAB(1 To intReG, 1 To 3) As Variant
      BREC12.MoveFirst
      intReG = 1
      Do While Not BREC12.EOF
         arrVERNIZACAB(intReG, 1) = BREC12!SGI_PRODUTO
         If Not IsNull(BREC12!SGI_CODUNIMED) Then arrVERNIZACAB(intReG, 2) = BREC12!SGI_CODUNIMED
         If Not IsNull(BREC12!SGI_QTDE) Then arrVERNIZACAB(intReG, 3) = BREC12!SGI_QTDE
         
         intReG = (intReG + 1)
         BREC12.MoveNext
      Loop
   End If
   BREC12.Close
   
   '' ----------------------------------------------
   '' Esmalte
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From" & vbCrLf
   sSql = sSql & "       SGI_ESMALTEPROD" & vbCrLf
   sSql = sSql & " Where" & vbCrLf
   sSql = sSql & "       SGI_FILIAL    = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_IDPRODUTO = " & lngIDProduto
   
   BREC12.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC12.EOF Then
      intReG = 0
      Do While Not BREC12.EOF
         intReG = intReG + 1
         BREC12.MoveNext
      Loop
   
      ReDim arrESMALTE(1 To intReG, 1 To 3) As Variant
      BREC12.MoveFirst
      intReG = 1
      Do While Not BREC12.EOF
         arrESMALTE(intReG, 1) = BREC12!SGI_PRODUTO
         If Not IsNull(BREC12!SGI_CODUNIMED) Then arrESMALTE(intReG, 2) = BREC12!SGI_CODUNIMED
         If Not IsNull(BREC12!SGI_QTDE) Then arrESMALTE(intReG, 3) = BREC12!SGI_QTDE
         
         intReG = (intReG + 1)
         BREC12.MoveNext
      Loop
   End If
   BREC12.Close
      
   '' ----------------------------------------------
   '' Verniz Externo
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From" & vbCrLf
   sSql = sSql & "       SGI_VERNIZEXT" & vbCrLf
   sSql = sSql & " Where" & vbCrLf
   sSql = sSql & "       SGI_FILIAL    = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_IDPRODUTO = " & lngIDProduto
   
   BREC12.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC12.EOF Then
      intReG = 0
      Do While Not BREC12.EOF
         intReG = intReG + 1
         BREC12.MoveNext
      Loop
   
      ReDim arrVERNIZEXT(1 To intReG, 1 To 3) As Variant
      BREC12.MoveFirst
      intReG = 1
      Do While Not BREC12.EOF
         arrVERNIZEXT(intReG, 1) = BREC12!SGI_PRODUTO
         If Not IsNull(BREC12!SGI_CODUNIMED) Then arrVERNIZEXT(intReG, 2) = BREC12!SGI_CODUNIMED
         If Not IsNull(BREC12!SGI_QTDE) Then arrVERNIZEXT(intReG, 3) = BREC12!SGI_QTDE
         
         intReG = (intReG + 1)
         BREC12.MoveNext
      Loop
   End If
   BREC12.Close
      
   '' ----------------------------------------------
   '' Vedante Fechamento
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From" & vbCrLf
   sSql = sSql & "       SGI_VEDANTEPROD" & vbCrLf
   sSql = sSql & " Where" & vbCrLf
   sSql = sSql & "       SGI_FILIAL    = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_IDPRODUTO = " & lngIDProduto
   
   BREC12.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC12.EOF Then
      intReG = 0
      Do While Not BREC12.EOF
         intReG = intReG + 1
         BREC12.MoveNext
      Loop
   
      ReDim arrVEDANTE(1 To intReG, 1 To 3) As Variant
      BREC12.MoveFirst
      intReG = 1
      Do While Not BREC12.EOF
         arrVEDANTE(intReG, 1) = BREC12!SGI_PRODUTO
         If Not IsNull(BREC12!SGI_CODUNIMED) Then arrVEDANTE(intReG, 2) = BREC12!SGI_CODUNIMED
         If Not IsNull(BREC12!SGI_QTDE) Then arrVEDANTE(intReG, 3) = BREC12!SGI_QTDE
         intReG = (intReG + 1)
         BREC12.MoveNext
      Loop
   End If
   BREC12.Close
   
   
   '' ----------------------------------------------
   '' Tampa Pressão
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From" & vbCrLf
   sSql = sSql & "       SGI_FECHAMCAMPOS" & vbCrLf
   sSql = sSql & " Where" & vbCrLf
   sSql = sSql & "       SGI_FILIAL    = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
   sSql = sSql & "   And SGI_CAMPO     = " & intTampaPressao & vbCrLf
   sSql = sSql & " Order By " & vbCrLf
   sSql = sSql & "         SGI_INDICE "
   
   BREC12.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC12.EOF Then
      intReG = 0
      Do While Not BREC12.EOF
         intReG = intReG + 1
         BREC12.MoveNext
      Loop
   
      ReDim arrTAMPAPRESS(1 To intReG) As Variant
      BREC12.MoveFirst
      intReG = 1
      Do While Not BREC12.EOF
         arrTAMPAPRESS(intReG) = BREC12!SGI_VALOR
         intReG = (intReG + 1)
         BREC12.MoveNext
      Loop
   End If
   BREC12.Close
   
   '' ----------------------------------------------
   '' Batoque Retratil
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From" & vbCrLf
   sSql = sSql & "       SGI_FECHAMCAMPOS" & vbCrLf
   sSql = sSql & " Where" & vbCrLf
   sSql = sSql & "       SGI_FILIAL    = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
   sSql = sSql & "   And SGI_CAMPO     = " & intBatoqueRetratil & vbCrLf
   sSql = sSql & " Order By " & vbCrLf
   sSql = sSql & "         SGI_INDICE "
   
   BREC12.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC12.EOF Then
      intReG = 0
      Do While Not BREC12.EOF
         intReG = intReG + 1
         BREC12.MoveNext
      Loop
   
      ReDim arrBATRETRATI(1 To intReG) As Variant
      BREC12.MoveFirst
      intReG = 1
      Do While Not BREC12.EOF
         arrBATRETRATI(intReG) = BREC12!SGI_VALOR
         intReG = (intReG + 1)
         BREC12.MoveNext
      Loop
   End If
   BREC12.Close
   
   
   '' ----------------------------------------------
   '' Batoque Plastico
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From" & vbCrLf
   sSql = sSql & "       SGI_FECHAMCAMPOS" & vbCrLf
   sSql = sSql & " Where" & vbCrLf
   sSql = sSql & "       SGI_FILIAL    = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
   sSql = sSql & "   And SGI_CAMPO     = " & intBatoquePlastico & vbCrLf
   sSql = sSql & " Order By " & vbCrLf
   sSql = sSql & "         SGI_INDICE "
   
   BREC12.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC12.EOF Then
      intReG = 0
      Do While Not BREC12.EOF
         intReG = intReG + 1
         BREC12.MoveNext
      Loop
   
      ReDim arrBATPLASTIC(1 To intReG) As Variant
      BREC12.MoveFirst
      intReG = 1
      Do While Not BREC12.EOF
         arrBATPLASTIC(intReG) = BREC12!SGI_VALOR
         intReG = (intReG + 1)
         BREC12.MoveNext
      Loop
   End If
   BREC12.Close
   
   '' ----------------------------------------------
   '' Tampa Visor
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From" & vbCrLf
   sSql = sSql & "       SGI_FECHAMCAMPOS" & vbCrLf
   sSql = sSql & " Where" & vbCrLf
   sSql = sSql & "       SGI_FILIAL    = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
   sSql = sSql & "   And SGI_CAMPO     = " & intTampaVisor & vbCrLf
   sSql = sSql & " Order By " & vbCrLf
   sSql = sSql & "         SGI_INDICE "
   
   BREC12.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC12.EOF Then
      intReG = 0
      Do While Not BREC12.EOF
         intReG = intReG + 1
         BREC12.MoveNext
      Loop
   
      ReDim arrTAMPAVISOR(1 To intReG) As Variant
      BREC12.MoveFirst
      intReG = 1
      Do While Not BREC12.EOF
         arrTAMPAVISOR(intReG) = BREC12!SGI_VALOR
         intReG = (intReG + 1)
         BREC12.MoveNext
      Loop
   End If
   BREC12.Close
   
   
   '' ----------------------------------------------
   '' Cliente que este Produto Atende
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From" & vbCrLf
   sSql = sSql & "       SGI_PRODATECLIE" & vbCrLf
   sSql = sSql & " Where" & vbCrLf
   sSql = sSql & "       SGI_FILIAL    = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
   
   BREC12.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC12.EOF Then
      intReG = 0
      Do While Not BREC12.EOF
         intReG = intReG + 1
         BREC12.MoveNext
      Loop
   
      ReDim arrPRODATECLIE(1 To intReG) As Variant
      BREC12.MoveFirst
      intReG = 1
      Do While Not BREC12.EOF
         arrPRODATECLIE(intReG) = BREC12!SGI_IDCLIENTE
         intReG = (intReG + 1)
         BREC12.MoveNext
      Loop
   End If
   BREC12.Close
   
   
   '' Familia de Máquinas
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_PRODFAMMAQ " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial
   sSql = sSql & "   And SGI_IDPRODUTO = " & lngIDProduto & vbCrLf
   
   BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC3.EOF Then
   
      intReG = 0
      Do While Not BREC3.EOF
         intReG = intReG + 1
         BREC3.MoveNext
      Loop
      
      ReDim arrFAMMAQ(1 To intReG) As Variant
      
      intReG = 1
      BREC3.MoveFirst
      Do While Not BREC3.EOF
         arrFAMMAQ(intReG) = BREC3!SGI_CODFAMMAQ
         intReG = intReG + 1
         BREC3.MoveNext
      Loop
   
   End If
   BREC3.Close

   '' ------------------------
   '' Produto de Entrada
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_PRODENTRADA " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial
   sSql = sSql & "   And SGI_CODIGO = " & lngIDProduto & vbCrLf
   
   BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC3.EOF Then
   
      intReG = 0
      Do While Not BREC3.EOF
         intReG = intReG + 1
         BREC3.MoveNext
      Loop
      
      ReDim arrPRODENTR(1 To intReG, 1 To 3) As Variant
      
      intReG = 1
      BREC3.MoveFirst
      Do While Not BREC3.EOF
         arrPRODENTR(intReG, 1) = BREC3!SGI_IDPRODUTO
         If Not IsNull(BREC3!SGI_CODUNIMED) Then arrPRODENTR(intReG, 2) = BREC3!SGI_CODUNIMED
         If Not IsNull(BREC3!SGI_QTDE) Then arrPRODENTR(intReG, 3) = BREC3!SGI_QTDE
         intReG = intReG + 1
         BREC3.MoveNext
      Loop
   
   End If
   BREC3.Close
   
   
   '' Carrega Lista de Material
   '' 07/08/2017
   
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_LISTAMATPROD " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL  = " & intFilial & vbCrLf
   sSql = sSql & "   And SGI_CODSTRU = " & lngIDProduto & vbCrLf
   
   BREC4.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC4.EOF() Then
      
      intReG = 0
      Do While Not BREC4.EOF()
         intReG = intReG + 1
         BREC4.MoveNext
      Loop
      
      ReDim arrPROVARV(1 To intReG) As PRODARVPROD
      
      intReG = 1
      BREC4.MoveFirst
      Do While Not BREC4.EOF()
      
         arrPROVARV(intReG).intAction2Do = dacEnumUpdateAction_Ignore
         arrPROVARV(intReG).lngCODIGO = BREC4!SGI_CODIGO
         arrPROVARV(intReG).lngCODPAI = BREC4!SGI_CODPAI
         arrPROVARV(intReG).lngProdutoID = BREC4!SGI_IDPRODUTO
         arrPROVARV(intReG).lngProdutoIDPai = BREC4!SGI_IDPRODPAI
         arrPROVARV(intReG).lngCodUniMed = BREC4!SGI_CODUNIMED
         If Not IsNull(BREC4!SGI_QTDE) Then arrPROVARV(intReG).strQTDCONS = Format(BREC4!SGI_QTDE, "#,####0.0000")
         
         intReG = (intReG + 1)
         BREC4.MoveNext
      Loop
   Else
      ReDim arrPROVARV(1 To 1) As PRODARVPROD
      arrPROVARV(1).lngProdutoID = -1
   End If
   BREC4.Close

End Function


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
    
    If Combo.ListCount > 0 Then Combo.ListIndex = 0
    
    BREC.Close

End Sub

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

Public Sub PreenchComboClaFis(Combo As Variant)

    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADCLASSFIS " & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "Order by SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       Combo.AddItem Trim(BREC!SGI_NOMECLA)
       Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
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


Public Sub PreenchComboSubGrupProduto(Combo As Variant, intGRUPPROD As Integer)

    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       GRUPO.SGI_CODESPECIE " & vbCrLf
    sSql = sSql & "      ,SUBGR.SGI_DESCRICAO  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_GRUPPRODITEN GRUPO" & vbCrLf
    sSql = sSql & "      ,SGI_CADSUBGRPROD SUBGR" & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      GRUPO.SGI_CODIGO = " & intGRUPPROD & vbCrLf
    sSql = sSql & "  And GRUPO.SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "  And SUBGR.SGI_FILIAL = GRUPO.SGI_FILIAL " & vbCrLf
    sSql = sSql & "  And SUBGR.SGI_CODIGO = GRUPO.SGI_CODESPECIE " & vbCrLf
    sSql = sSql & "Order by GRUPO.SGI_CODESPECIE"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       Combo.AddItem Trim(BREC!SGI_DESCRICAO)
       Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODESPECIE
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

Public Sub PreenchComboTabConv(Combo As Variant)

    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADTABCONV " & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "Order by SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       Combo.AddItem Trim(Str(BREC!SGI_CODIGO)) & Space(7 - Len(Trim(Str(BREC!SGI_CODIGO)))) & " - Unidade de " & PegaUniMed(BREC!SGI_CODUNIDDE) & " Para " & PegaUniMed(BREC!SGI_CODUNIDPARA)
       Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
       BREC.MoveNext
    Loop
    
    BREC.Close

End Sub

Private Function PegaUniMed(intCODUNIMED As Integer) As String

    PegaUniMed = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADUNIMED " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & FILIAL & vbCrLf
    sSql = sSql & "  And SGI_CODIGO = " & intCODUNIMED
    
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC2.EOF Then PegaUniMed = BREC2!SGI_UNIDADE
    BREC2.Close
    
End Function

Public Function Atualiza(strACAO As String, strCODIGO As String, lngFilial_Atu As Integer, strModulo As String) As Boolean
    
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
    
    If BREC.State = 1 Then BREC.Close
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC.EOF Then
       
        sSql = "Insert Into SGI_ATUALIZA (SGI_FILIAL" & vbCrLf
        sSql = sSql & "                  ,SGI_MODULO" & vbCrLf
        sSql = sSql & "                  ,SGI_ACAO" & vbCrLf
        sSql = sSql & "                  ,SGI_CODIGO)" & vbCrLf
        sSql = sSql & "           Values (" & vbCrLf
        sSql = sSql & "                   " & lngFilial_Atu & vbCrLf
        sSql = sSql & "                 ,'" & Trim(strModulo) & "'" & vbCrLf
        sSql = sSql & "                 ,'" & Trim(strACAO) & "'" & vbCrLf
        sSql = sSql & "                 ,'" & Trim(strCODIGO) & "'" & vbCrLf
        sSql = sSql & "                   )"
    
        BGRV.CommandText = sSql
        BGRV.Execute
    
    Else
    
        sSql = sSql & "Update SGI_ATUALIZA Set" & vbCrLf
        sSql = sSql & "           SGI_ACAO   = '" & Trim(strACAO) & "'" & vbCrLf
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
     Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, strACAO, sSql)
     Set objErro = Nothing
    
End Function

Public Sub PreenchComboFerramenta(Combo As Variant)

    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADFERRA " & vbCrLf
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


Public Function Update_ListaMat(strOPER As String) As Boolean

On Error GoTo err_Trans
        
    Update_ListaMat = False
    
    Dim curCustoTotal As Currency
    Dim strValor      As String
    
    '' Inicia transação
    adoBanco_Dados.BeginTrans
    BGRV.ActiveConnection = adoBanco_Dados
    
    '' Alterando Possiveis componentes em outras listas
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_LISTAMAT " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL  = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_PRODLST = '" & Trim(strCodigoProd) & "'"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
    
       '' -------------------------------------------------------
       curCustoTotal = (BREC!SGI_QTDE * curPRCCUSTO)
    
       sSql = "Update SGI_LISTAMAT " & vbCrLf
       sSql = sSql & "                Set " & vbCrLf
        
       strValor = Replace(Format(curPRCCUSTO, "#,##0.00"), ".", "")
       strValor = Replace(strValor, ",", ".")
       sSql = sSql & "                    SGI_CUSTOUNIT  = " & strValor & vbCrLf
       
       strValor = Replace(Format(curCustoTotal, "#,##0.00"), ".", "")
       strValor = Replace(strValor, ",", ".")
       sSql = sSql & "                   ,SGI_CUSTOTOTAL = " & strValor & vbCrLf
    
       sSql = sSql & " Where " & vbCrLf
       sSql = sSql & "       SGI_FILIAL  = " & intFilial & vbCrLf
       sSql = sSql & "   And SGI_PRODLST = '" & Trim(strCodigoProd) & "'"
       
       BGRV.CommandText = sSql
       BGRV.Execute
       '' -------------------------------------------------------
    
       BREC.MoveNext
    Loop
    BREC.Close
    
    Update_ListaMat = True
    
    adoBanco_Dados.CommitTrans
        
    Exit Function
        
err_Trans:

    If adoBanco_Dados.State = 1 Then adoBanco_Dados.RollbackTrans
    
    Dim objErro As Object
    Set objErro = CreateObject("BLBCWS.clsFuncoes")
    Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, strOPER, sSql)
    Set objErro = Nothing
        

End Function

Public Function Update_CodProd() As Boolean

On Error GoTo err_Trans
        
    Update_CodProd = False
    
    '' Inicia transação
    adoBanco_Dados.BeginTrans
    BGRV.ActiveConnection = adoBanco_Dados
    
    '' Se caso o Código do Produto For Alterado
    If Trim(strCodigoProd) <> Trim(strCodigoProdANT) Then
       
       '' Cadastro do Produto
       sSql = "Update SGI_CADPRODUTO Set SGI_CODIGO = '" & Trim(strCodigoProd) & "'" & vbCrLf
       sSql = sSql & "             Where " & vbCrLf
       sSql = sSql & "                   SGI_FILIAL = " & intFilial & vbCrLf
       sSql = sSql & "               And SGI_CODIGO = '" & Trim(strCodigoProdANT) & "'"
       
       BGRV.CommandText = sSql
       BGRV.Execute
       
       '' ----------------------
       
       '' Lista de material
       sSql = "Update SGI_LISTAMAT Set SGI_PRODLST = '" & Trim(strCodigoProd) & "'" & vbCrLf
       sSql = sSql & "                  Where " & vbCrLf
       sSql = sSql & "                        SGI_FILIAL  = " & intFilial & vbCrLf
       sSql = sSql & "                    And SGI_PRODLST = '" & Trim(strCodigoProdANT) & "'"
       
       BGRV.CommandText = sSql
       BGRV.Execute
       
       '' ----------------------
       
       '' Lista de material
       sSql = "Update SGI_LISTAMAT Set SGI_PRODUTO = '" & Trim(strCodigoProd) & "'" & vbCrLf
       sSql = sSql & "                  Where " & vbCrLf
       sSql = sSql & "                        SGI_FILIAL  = " & intFilial & vbCrLf
       sSql = sSql & "                    And SGI_PRODUTO = '" & Trim(strCodigoProdANT) & "'"
       
       BGRV.CommandText = sSql
       BGRV.Execute
       
       '' ----------------------
       
       '' Especificação Técnica
       sSql = "Update SGI_PRODESPTEC Set SGI_PRODUTO = '" & Trim(strCodigoProd) & "'" & vbCrLf
       sSql = sSql & "                    Where "
       sSql = sSql & "                          SGI_FILIAL  =  " & intFilial
       sSql = sSql & "                      And SGI_PRODUTO = '" & Trim(strCodigoProdANT) & "'"
              
       BGRV.CommandText = sSql
       BGRV.Execute
       
       '' ----------------------
       
       '' Processos
       sSql = "Update SGI_CADPROPROC Set SGI_CODPROD = '" & Trim(strCodigoProd) & "'" & vbCrLf
       sSql = sSql & "                    Where " & vbCrLf
       sSql = sSql & "                          SGI_FILIAL  = " & intFilial & vbCrLf
       sSql = sSql & "                      And SGI_CODPROD = '" & Trim(strCodigoProdANT) & "'"
       
       BGRV.CommandText = sSql
       BGRV.Execute
       
       '' ----------------------
       
       '' Ferramentas
       sSql = "Update SGI_CADFERRPROC Set SGI_CODPROD = '" & Trim(strCodigoProd) & "'" & vbCrLf
       sSql = sSql & "                    Where " & vbCrLf
       sSql = sSql & "                          SGI_FILIAL  = " & intFilial & vbCrLf
       sSql = sSql & "                      And SGI_CODPROD = '" & Trim(strCodigoProdANT) & "'"
      
       BGRV.CommandText = sSql
       BGRV.Execute
       
       '' ----------------------
       
       '' Cardex
       sSql = "Update SGI_CARDEXEST Set SGI_PRODUTO = '" & Trim(strCodigoProd) & "'" & vbCrLf
       sSql = sSql & "                    Where " & vbCrLf
       sSql = sSql & "                          SGI_FILIAL  = " & intFilial & vbCrLf
       sSql = sSql & "                      And SGI_PRODUTO = '" & Trim(strCodigoProdANT) & "'"
       
       BGRV.CommandText = sSql
       BGRV.Execute
       
       '' ----------------------
       
       '' Itens de Cotação de Compras
       sSql = "Update SGI_COTAITENS Set SGI_PRODUTO = '" & Trim(strCodigoProd) & "'" & vbCrLf
       sSql = sSql & "                    Where " & vbCrLf
       sSql = sSql & "                          SGI_FILIAL  = " & intFilial & vbCrLf
       sSql = sSql & "                      And SGI_PRODUTO = '" & Trim(strCodigoProdANT) & "'"
    
       BGRV.CommandText = sSql
       BGRV.Execute
       '' ----------------------
       
       '' Itens do Tipo de Liberação
       sSql = "Update SGI_COTATIPLIB Set SGI_PRODUTO = '" & Trim(strCodigoProd) & "'" & vbCrLf
       sSql = sSql & "                    Where " & vbCrLf
       sSql = sSql & "                          SGI_FILIAL  = " & intFilial & vbCrLf
       sSql = sSql & "                      And SGI_PRODUTO = '" & Trim(strCodigoProdANT) & "'"
       
       BGRV.CommandText = sSql
       BGRV.Execute
       
       '' ----------------------
       
       '' Itens de Cotação de Vendas
       sSql = "Update SGI_CADCOTAVENDI Set SGI_CODPROD = '" & Trim(strCodigoProd) & "'" & vbCrLf
       sSql = sSql & "                    Where " & vbCrLf
       sSql = sSql & "                          SGI_FILIAL  = " & intFilial & vbCrLf
       sSql = sSql & "                      And SGI_CODPROD = '" & Trim(strCodigoProdANT) & "'"
       
       BGRV.CommandText = sSql
       BGRV.Execute
       
       '' ----------------------
       
       '' Intens de Requisição de Entrada
       sSql = "Update SGI_CADITREQENTRMAT Set SGI_PRODUTO = '" & Trim(strCodigoProd) & "'" & vbCrLf
       sSql = sSql & "                    Where " & vbCrLf
       sSql = sSql & "                          SGI_FILIAL  = " & intFilial & vbCrLf
       sSql = sSql & "                      And SGI_PRODUTO = '" & Trim(strCodigoProdANT) & "'"
       
       BGRV.CommandText = sSql
       BGRV.Execute
       
       '' ----------------------
       '' Produto do Fornecedor
       sSql = "Update SGI_PRODFOR Set SGI_PRODUTO = '" & Trim(strCodigoProd) & "'" & vbCrLf
       sSql = sSql & "                    Where " & vbCrLf
       sSql = sSql & "                          SGI_FILIAL  = " & intFilial & vbCrLf
       sSql = sSql & "                      And SGI_PRODUTO = '" & Trim(strCodigoProdANT) & "'"
       
       BGRV.CommandText = sSql
       BGRV.Execute
       
       '' ----------------------
       
       '' Produtos da Maquina
       sSql = "Update SGI_CADMAQPROD Set SGI_CODPROD = '" & Trim(strCodigoProd) & "'" & vbCrLf
       sSql = sSql & "                    Where " & vbCrLf
       sSql = sSql & "                          SGI_FILIAL  = " & intFilial & vbCrLf
       sSql = sSql & "                      And SGI_CODPROD = '" & Trim(strCodigoProdANT) & "'"
       
       BGRV.CommandText = sSql
       BGRV.Execute
       
       '' ----------------------
       
       '' NF de Entrada Itens
       sSql = "Update SGI_NFENTRADAITENS Set SGI_PRODUTO = '" & Trim(strCodigoProd) & "'" & vbCrLf
       sSql = sSql & "                    Where " & vbCrLf
       sSql = sSql & "                          SGI_FILIAL  = " & intFilial & vbCrLf
       sSql = sSql & "                      And SGI_PRODUTO = '" & Trim(strCodigoProdANT) & "'"
       
       BGRV.CommandText = sSql
       BGRV.Execute
       
       '' ----------------------
       
       '' NF de Saida Itens
       sSql = "Update SGI_NFSAIDAITENS Set SGI_PRODUTO = '" & Trim(strCodigoProd) & "'" & vbCrLf
       sSql = sSql & "                    Where " & vbCrLf
       sSql = sSql & "                          SGI_FILIAL  = " & intFilial & vbCrLf
       sSql = sSql & "                      And SGI_PRODUTO = '" & Trim(strCodigoProdANT) & "'"
       
       BGRV.CommandText = sSql
       BGRV.Execute
       
       '' ----------------------
       
       '' OF
       sSql = "Update SGI_CADOF Set SGI_CODPROD = '" & Trim(strCodigoProd) & "'" & vbCrLf
       sSql = sSql & "                    Where " & vbCrLf
       sSql = sSql & "                          SGI_FILIAL  = " & intFilial & vbCrLf
       sSql = sSql & "                      And SGI_CODPROD = '" & Trim(strCodigoProdANT) & "'"
       
       BGRV.CommandText = sSql
       BGRV.Execute
       
       '' ----------------------
       
       '' Lista de Material da OF
       sSql = "Update SGI_CADLISTAPODOF Set SGI_CODPROD = '" & Trim(strCodigoProd) & "'" & vbCrLf
       sSql = sSql & "                    Where " & vbCrLf
       sSql = sSql & "                          SGI_FILIAL  = " & intFilial & vbCrLf
       sSql = sSql & "                      And SGI_CODPROD = '" & Trim(strCodigoProdANT) & "'"
       
       BGRV.CommandText = sSql
       BGRV.Execute
        
       '' Lista de Material da OF
       sSql = "Update SGI_CADLISTAPODOF Set SGI_CODPRODLST = '" & Trim(strCodigoProd) & "'" & vbCrLf
       sSql = sSql & "                    Where " & vbCrLf
       sSql = sSql & "                          SGI_FILIAL     = " & intFilial & vbCrLf
       sSql = sSql & "                      And SGI_CODPRODLST = '" & Trim(strCodigoProdANT) & "'"
       
       BGRV.CommandText = sSql
       BGRV.Execute
       
       '' ----------------------
       
       '' Itens do Pedido de Compras
       sSql = "Update SGI_PEDIDOITENS Set SGI_PRODUTO = '" & Trim(strCodigoProd) & "'" & vbCrLf
       sSql = sSql & "                    Where " & vbCrLf
       sSql = sSql & "                          SGI_FILIAL     = " & intFilial & vbCrLf
       sSql = sSql & "                      And SGI_PRODUTO = '" & Trim(strCodigoProdANT) & "'"
       
       BGRV.CommandText = sSql
       BGRV.Execute
       
       '' ----------------------
       
       '' Itens de Pedidos de Vendas
       sSql = "Update SGI_CADPEDVENDI Set SGI_CODPROD = '" & Trim(strCodigoProd) & "'" & vbCrLf
       sSql = sSql & "                    Where " & vbCrLf
       sSql = sSql & "                          SGI_FILIAL  = " & intFilial & vbCrLf
       sSql = sSql & "                      And SGI_CODPROD = '" & Trim(strCodigoProdANT) & "'"
       
       BGRV.CommandText = sSql
       BGRV.Execute
       
       '' ----------------------
       
       '' Plano Mestre
       sSql = "Update SGI_CADPLANMESTRE Set SGI_CODPROD = '" & Trim(strCodigoProd) & "'" & vbCrLf
       sSql = sSql & "                    Where " & vbCrLf
       sSql = sSql & "                          SGI_FILIAL  = " & intFilial & vbCrLf
       sSql = sSql & "                      And SGI_CODPROD = '" & Trim(strCodigoProdANT) & "'"
       
       BGRV.CommandText = sSql
       BGRV.Execute
       
       '' ----------------------
       
       '' Requisição de Material Itens
       sSql = "Update SGI_CADITREQMAT Set SGI_PRODUTO = '" & Trim(strCodigoProd) & "'" & vbCrLf
       sSql = sSql & "                    Where " & vbCrLf
       sSql = sSql & "                          SGI_FILIAL  = " & intFilial & vbCrLf
       sSql = sSql & "                      And SGI_PRODUTO = '" & Trim(strCodigoProdANT) & "'"
       
       BGRV.CommandText = sSql
       BGRV.Execute
       
       '' ----------------------
       
       '' Requisição de Saida de Material Itens
       sSql = "Update SGI_CADITREQSAIMAT Set SGI_PRODUTO = '" & Trim(strCodigoProd) & "'" & vbCrLf
       sSql = sSql & "                    Where " & vbCrLf
       sSql = sSql & "                          SGI_FILIAL  = " & intFilial & vbCrLf
       sSql = sSql & "                      And SGI_PRODUTO = '" & Trim(strCodigoProdANT) & "'"
       
       BGRV.CommandText = sSql
       BGRV.Execute
       
       '' ----------------------
       
       '' Tabela de Preços
       sSql = "Update SGI_TABPRECO Set SGI_CODPROD = '" & Trim(strCodigoProd) & "'" & vbCrLf
       sSql = sSql & "                    Where " & vbCrLf
       sSql = sSql & "                          SGI_FILIAL  = " & intFilial & vbCrLf
       sSql = sSql & "                      And SGI_CODPROD = '" & Trim(strCodigoProdANT) & "'"
       
       BGRV.CommandText = sSql
       BGRV.Execute
       
       '' ----------------------
       
    End If
    
    Update_CodProd = True
    
    adoBanco_Dados.CommitTrans
        
    Exit Function
        
err_Trans:

    adoBanco_Dados.RollbackTrans
    
    Dim objErro As Object
    Set objErro = CreateObject("BLBCWS.clsFuncoes")
    Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, "A", sSql)
    Set objErro = Nothing
        

End Function

Public Function RecalcEstrutura() As Boolean
    
On Error GoTo err_Trans
    
    Dim strPRODUTO  As String
    Dim strPAI      As String
    Dim strValor    As String
    Dim curTotCusto As Currency
    Dim curCustProc As Currency
    Dim curCustGer  As Currency
    Dim curPorcAcr  As Currency
    Dim curCustPorc As Currency
    Dim curCustoPro As Currency
    Dim arrALTLIST  As Variant
    Dim arrCAMPOS   As Variant
    Dim intItens    As Integer
    Dim i           As Integer
    
    strPRODUTO = ""
    intItens = 0
    
    RecalcEstrutura = False
    
    ReDim arrALTLIST(1) As String
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_LISTAMAT " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_PRODLST = '" & Trim(strCodigoProd) & "'" & vbCrLf
    sSql = sSql & "   And SGI_FILIAL  = " & FILIAL
    
    BREC2.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC2.EOF
    
       '' Volta para continuar recalculando estrutura
Volta_Princ:

       intItens = intItens + 1
    
       If Len(Trim(strPRODUTO)) = 0 Then
          strPRODUTO = BREC2!SGI_PRODUTO
       Else
          sSql = "Select " & vbCrLf
          sSql = sSql & "       * " & vbCrLf
          sSql = sSql & "  From " & vbCrLf
          sSql = sSql & "       SGI_LISTAMAT " & vbCrLf
          sSql = sSql & " Where " & vbCrLf
          sSql = sSql & "       SGI_PRODLST = '" & Trim(strPRODUTO) & "'" & vbCrLf
          sSql = sSql & "   And SGI_FILIAL  = " & FILIAL
    
          strPRODUTO = ""
          BREC.Open sSql, adoBanco_Dados, adOpenDynamic
          If Not BREC.EOF Then strPRODUTO = BREC!SGI_PRODUTO
          BREC.Close
       End If
    
       If Len(Trim(strPRODUTO)) > 0 Then
    
          curTotCusto = 0
          curCustGer = 0
          curPorcAcr = 0
          curCustPorc = 0
          curCustoPro = 0
       
          sSql = "Select " & vbCrLf
          sSql = sSql & "       Sum(SGI_CUSTOTOTAL) As SGI_CUSTOTOTAL" & vbCrLf
          sSql = sSql & "  From " & vbCrLf
          sSql = sSql & "       SGI_LISTAMAT " & vbCrLf
          sSql = sSql & " Where " & vbCrLf
          sSql = sSql & "       SGI_PRODUTO = '" & Trim(strPRODUTO) & "'" & vbCrLf
          sSql = sSql & "   And SGI_FILIAL  = " & FILIAL
    
          BREC.Open sSql, adoBanco_Dados, adOpenDynamic
          If BREC!SGI_CUSTOTOTAL > 0 Then curTotCusto = BREC!SGI_CUSTOTOTAL
          BREC.Close
    
          If curTotCusto > 0 Then
       
             curPorcAcr = PegaPorcAcres(strPRODUTO) '' Pega % de Acrescimo
             curCustProc = PegaCustProc(strPRODUTO) '' Pega Custo do Processo
          
             curCustoPro = (curTotCusto + curCustProc) '' Custo do Prod + Processo
          
             curCustPorc = ((curCustoPro * curPorcAcr) / 100) '' Calc. % de Acresc.
          
             curCustGer = (curCustoPro + curCustPorc)
          
          ReDim Preserve arrALTLIST(1 To intItens) As String
          
             arrALTLIST(intItens) = Trim(strPRODUTO) & "|" & Format(curCustGer, "#,##0.00") & "|" & Format(curCustoPro, "#,##0.00")
       
          End If
    
          GoTo Volta_Princ
       
       End If
    
       BREC2.MoveNext
    Loop
    BREC2.Close
    
    '' Inicia transação
    adoBanco_Dados.BeginTrans
    BGRV.ActiveConnection = adoBanco_Dados
   
    For i = 1 To UBound(arrALTLIST)
    
        If Len(Trim(arrALTLIST(i))) > 0 Then
        
           arrCAMPOS = Split(arrALTLIST(i), "|")
        
           '' -----------------------------------
           strValor = Replace(arrCAMPOS(1), ".", "")
           strValor = Replace(strValor, ",", ".")
           '' -----------------------------------
        
           sSql = "Update SGI_LISTAMAT Set " & vbCrLf
           sSql = sSql & "           SGI_CUSTOUNIT = " & strValor & vbCrLf
           sSql = sSql & "          ,SGI_CUSTOTOTAL = (SGI_QTDE * " & strValor & ")" & vbCrLf
           sSql = sSql & "     Where " & vbCrLf
           sSql = sSql & "           SGI_PRODLST = '" & Trim(arrCAMPOS(0)) & "'" & vbCrLf
           sSql = sSql & "       And SGI_FILIAL  = " & intFilial
        
           BGRV.CommandText = sSql
           BGRV.Execute
         
           '' -----------------------------------------------
           '' Altera Valor do Produto
           sSql = "Update SGI_CADPRODUTO Set " & vbCrLf
          
           strValor = Replace(arrCAMPOS(2), ".", "")
           strValor = Replace(strValor, ",", ".")
           sSql = sSql & "                          SGI_PRECOPROD = " & strValor & vbCrLf
       
           strValor = Replace(arrCAMPOS(1), ".", "")
           strValor = Replace(strValor, ",", ".")
           sSql = sSql & "                         ,SGI_PRCCUSTO  = " & strValor & vbCrLf
          
           sSql = sSql & "                    Where " & vbCrLf
           sSql = sSql & "                          SGI_FILIAL  = " & intFilial & vbCrLf
           sSql = sSql & "                      And SGI_CODIGO  = '" & Trim(arrCAMPOS(0)) & "'"
          
           BGRV.CommandText = sSql
           BGRV.Execute
          
           '' -----------------------------------------------
        
        End If
        
    Next i
    
    adoBanco_Dados.CommitTrans
        
    RecalcEstrutura = True
    
    Exit Function
        
err_Trans:

    adoBanco_Dados.RollbackTrans
    
    Dim objErro As Object
    Set objErro = CreateObject("BLBCWS.clsFuncoes")
    Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, "A", sSql)
    Set objErro = Nothing
    
End Function
Private Function PegaCustProc(strPRODUTO As String) As Currency

    PegaCustProc = 0
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       Sum(SGI_TOTVALOR) As SGI_TOTVALOR " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPROPROC " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_CODPROD = '" & Trim(strPRODUTO) & "'" & vbCrLf
    sSql = sSql & "   And SGI_FILIAL  = " & intFilial
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If BREC!SGI_TOTVALOR > 0 Then PegaCustProc = BREC!SGI_TOTVALOR
    BREC.Close

End Function

Private Function PegaPorcAcres(strPRODUTO As String) As Currency
       
    PegaPorcAcres = 0
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADPRODUTO " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = '" & Trim(strPRODUTO) & "'"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF Then
       If Not IsNull(BREC!SGI_PROCACRES) Then PegaPorcAcres = BREC!SGI_PROCACRES
    End If
    BREC.Close
    
End Function
Public Sub PreenchComboTipoOrca(Combo As Variant)
    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADESPORCA " & vbCrLf
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

Public Sub PreenchComboVarInsp(Combo As Variant)
    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADVARINSP " & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      SGI_FILIAL    = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_INSPESAO = 0 " & vbCrLf
    sSql = sSql & "Order by SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       Combo.AddItem Trim(BREC!SGI_DESCRI)
       Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
       BREC.MoveNext
    Loop
    BREC.Close
    
End Sub

Public Sub PreenchComboAtrInsp(Combo As Variant)
    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADVARINSP " & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      SGI_FILIAL    = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_INSPESAO = 1 " & vbCrLf
    sSql = sSql & "Order by SGI_CODIGO"
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
    Do While Not BREC.EOF
       Combo.AddItem Trim(BREC!SGI_DESCRI)
       Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
       BREC.MoveNext
    Loop
    BREC.Close
    
End Sub

Public Function PreenchComboUnidadeGrd() As String
        
    PreenchComboUnidadeGrd = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADUNIMED " & vbCrLf
    sSql = sSql & "Where " & vbCrLf
    sSql = sSql & "      SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "Order by SGI_CODIGO"
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC.EOF
       PreenchComboUnidadeGrd = PreenchComboUnidadeGrd & "|#" & BREC!SGI_CODIGO & ";" & Trim(BREC!SGI_UNIDADE)
       BREC.MoveNext
    Loop
    BREC.Close
End Function

Public Sub PreenchComboFamilia(Combo As Variant)
    
    Combo.Clear
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *  " & vbCrLf
    sSql = sSql & "  from " & vbCrLf
    sSql = sSql & "       SGI_CADFAMMAQUINAS " & vbCrLf
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


Public Function GeraCodRotulo() As Long

    GeraCodRotulo = 1
    
    BGRV.ActiveConnection = adoBanco_Dados
    
    sSql = ""
    
    If intNATURALSIMNAO = 0 Then
    
            sSql = "Select " & vbCrLf
            sSql = sSql & "       (Max(SGI_CODROTULO) + 1) As SGI_CODROTULO " & vbCrLf
            sSql = sSql & "  From " & vbCrLf
            sSql = sSql & "       SGI_PRODESTAMPA" & vbCrLf
            sSql = sSql & " Where " & vbCrLf
            sSql = sSql & "       SGI_FILIAL     = " & intFilial & vbCrLf
            sSql = sSql & "   And SGI_CODCLIE    = " & lngCodClie & vbCrLf
            sSql = sSql & "   And SGI_CODLINPROD = " & lngCodLinProd
            
            BREC.Open sSql, adoBanco_Dados
            
            If Not BREC.EOF Then
            
               If IsNull(BREC!SGI_CODROTULO) = True Then
                  
                  sSql = "Insert into SGI_PRODESTAMPA (" & vbCrLf
                  sSql = sSql & "                       SGI_FILIAL" & vbCrLf
                  sSql = sSql & "                      ,SGI_CODIGO" & vbCrLf
                  sSql = sSql & "                      ,SGI_CODCLIE" & vbCrLf
                  sSql = sSql & "                      ,SGI_CODROTULO" & vbCrLf
                  sSql = sSql & "                      ,SGI_CODLINPROD) Values(" & vbCrLf
                  sSql = sSql & "                                              " & intFilial & vbCrLf
                  sSql = sSql & "                                            ," & lngIDProduto & vbCrLf
                  sSql = sSql & "                                            ," & lngCodClie & vbCrLf
                  sSql = sSql & "                                            ,1" & vbCrLf
                  sSql = sSql & "                                            ," & lngCodLinProd & vbCrLf
                  sSql = sSql & "                                          )" & vbCrLf
                  
               ElseIf BREC!SGI_CODROTULO > 1 Then
               
                  GeraCodRotulo = BREC!SGI_CODROTULO
                  
                  sSql = "Insert into SGI_PRODESTAMPA (" & vbCrLf
                  sSql = sSql & "                       SGI_FILIAL" & vbCrLf
                  sSql = sSql & "                      ,SGI_CODIGO" & vbCrLf
                  sSql = sSql & "                      ,SGI_CODCLIE" & vbCrLf
                  sSql = sSql & "                      ,SGI_CODROTULO" & vbCrLf
                  sSql = sSql & "                      ,SGI_CODLINPROD) Values(" & vbCrLf
                  sSql = sSql & "                                              " & intFilial & vbCrLf
                  sSql = sSql & "                                            ," & lngIDProduto & vbCrLf
                  sSql = sSql & "                                            ," & lngCodClie & vbCrLf
                  sSql = sSql & "                                            ," & GeraCodRotulo & vbCrLf
                  sSql = sSql & "                                            ," & lngCodLinProd & vbCrLf
                  sSql = sSql & "                                          )" & vbCrLf
               
               End If
               
               BGRV.CommandText = sSql
               BGRV.Execute
               
               
            End If
            
            BREC.Close
    
    Else
    
        GeraCodRotulo = 0
        
    End If
    
End Function


Public Function UpdateRotulo(strOPER As String) As Boolean
On Error GoTo err_Trans
     
     UpdateRotulo = False
     
     Dim i            As Integer
     Dim intCodRotulo As Integer
     Dim strValor     As String
     
     '' Inicia transação
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Then
     
        sSql = "Update SGI_CADPRODUTO Set SGI_CODROTULO = " & lngCodRotulo & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL       = " & intFilial & vbCrLf
        sSql = sSql & "   And SGI_IDPRODUTO    = " & lngIDProduto & vbCrLf
     
        BGRV.CommandText = sSql
        BGRV.Execute
     
     
     End If
     
     adoBanco_Dados.CommitTrans
     UpdateRotulo = True
     Exit Function

err_Trans:
    
    adoBanco_Dados.RollbackTrans
    Dim objErro    As Object
    Set objErro = CreateObject("BLBCWS.clsFuncoes")
    Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, strOPER, sSql)
    Set objErro = Nothing
    
End Function


Public Sub GravaLogModulo(intFilial_Atu As Integer, lngCodLog As Long, strModulo As String, strACAO As String, lngCodVendedor As Long, strCodReg As String)

On Error GoTo err_Trans
    
    '' Inicia transação
    adoBanco_Dados.BeginTrans
    BGRV.ActiveConnection = adoBanco_Dados

    sSql = "Insert Into SGI_LOGMODULO (SGI_FILIAL" & vbCrLf
    sSql = sSql & "                   ,SGI_CODIGO" & vbCrLf
    sSql = sSql & "                   ,SGI_MODULO" & vbCrLf
    sSql = sSql & "                   ,SGI_ACAO" & vbCrLf
    sSql = sSql & "                   ,SGI_CODUSUARIO" & vbCrLf
    sSql = sSql & "                   ,SGI_CODREG" & vbCrLf
    sSql = sSql & "                   ,SGI_DATA)" & vbCrLf
    sSql = sSql & "           Values (" & vbCrLf
    sSql = sSql & "                   " & intFilial_Atu & vbCrLf
    sSql = sSql & "                 ," & lngCodLog & vbCrLf
    sSql = sSql & "                 ,'" & Trim(strModulo) & "'" & vbCrLf
    sSql = sSql & "                 ,'" & Trim(strACAO) & "'" & vbCrLf
    sSql = sSql & "                 ," & lngCodVendedor & vbCrLf
    sSql = sSql & "                 ," & Trim(strCodReg) & vbCrLf
    sSql = sSql & "                 ,'" & Format(Now, "MM/DD/YYYY HH:MM:SS") & "'" & vbCrLf
    sSql = sSql & "                   )"
    
    BGRV.CommandText = sSql
    BGRV.Execute

    adoBanco_Dados.CommitTrans
     
    Exit Sub

err_Trans:
    
    adoBanco_Dados.RollbackTrans

    ''Dim objErro    As Object
    ''Set objErro = CreateObject("BLBCWS.clsFuncoes")
    ''Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, strOPER, sSql)
    ''Set objErro = Nothing

End Sub


Public Function PreenchComboUnidMedGrid() As String
        
    
    PreenchComboUnidMedGrid = ""
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADUNIMED" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFilial
    
    BREC8.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC8.EOF()
        PreenchComboUnidMedGrid = PreenchComboUnidMedGrid & "|#" & BREC8!SGI_CODIGO & ";" & Trim(BREC8!SGI_UNIDADE) & "-" & Trim(BREC8!SGI_DESCRICAO)
        BREC8.MoveNext
    Loop
    BREC8.Close
    
End Function

Public Function UpdateCodProd(strOPER As String, strCODPROD As String) As Boolean
On Error GoTo err_Trans
     
     UpdateCodProd = False
     
     Dim i            As Integer
     Dim intCodRotulo As Integer
     Dim strValor     As String
     
     '' Inicia transação
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Or strOPER = "A" Then
     
        sSql = "Update SGI_CADPRODUTO Set SGI_CODIGO = '" & Trim(strCODPROD) & "'" & vbCrLf
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL       = " & intFilial & vbCrLf
        sSql = sSql & "   And SGI_IDPRODUTO    = " & lngIDProduto & vbCrLf
     
        BGRV.CommandText = sSql
        BGRV.Execute
     
     
     End If
     
     adoBanco_Dados.CommitTrans
     UpdateCodProd = True
     Exit Function

err_Trans:
    
    adoBanco_Dados.RollbackTrans
    Dim objErro    As Object
    Set objErro = CreateObject("BLBCWS.clsFuncoes")
    Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, strOPER, sSql)
    Set objErro = Nothing
    
End Function

