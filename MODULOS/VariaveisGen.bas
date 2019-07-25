Attribute VB_Name = "VariaveisGen"
Option Explicit
Public BD               As New ADODB.Connection
Public BREC             As New ADODB.Recordset
Public BREC2            As New ADODB.Recordset
Public BREC3            As New ADODB.Recordset
Public BREC4            As New ADODB.Recordset
Public BREC5            As New ADODB.Recordset
Public BREC6            As New ADODB.Recordset
Public BREC7            As New ADODB.Recordset
Public BREC8            As New ADODB.Recordset
Public BREC9            As New ADODB.Recordset
Public BREC10           As New ADODB.Recordset
Public BREC11           As New ADODB.Recordset
Public BREC12           As New ADODB.Recordset
Public BRECATU          As New ADODB.Recordset
Public BGRV             As New ADODB.Command
Public sSql             As String
Public Stado            As String
Public vCod             As Variant
Public iCodUsu          As Integer
Public iFilial          As Integer
Public intFILIALPED01   As Integer
Public varRETORNO       As String
Public objPrinter       As Printer
Public colPRODUTOS      As Collection
Public strCAMARQERRO    As String
Public strNOMMAQUINA    As String

Public Const ConecStr = "DATABASE=NOVALATA;SERVER=RICARDO-HP;UID=sa;PWD=;"
Public Const ConecStr2 = "DATABASE=NOVALATA;SERVER=RICARDO-HP;UID=sa;PWD=connor220519;"
Public Const Prov = "SQLOLEDB"
Public Const EstadoFilial = "SP"

Public Const myvarAsByte        As Long = 0
Public Const myvarAsInteger     As Long = 1
Public Const myvarAsLong        As Long = 2
Public Const myvarAsDouble      As Long = 3
Public Const myvarAsCurrency    As Long = 4
Public Const myvarAsDate        As Long = 5
Public Const myvarAsSmallDate   As Long = 6
Public Const myvarAsSingle      As Long = 7
Public Const myvarAsString      As Long = 8

Public strCamRelNovo            As String
Public strCamImgRotulos         As String

Public Const strCamRels         As String = "\\SERVIDORNOVA\RPT\"
''Public Const strCamRels         As String = "C:\RICARDO\SGI\NOVALATA\RELATORIOS\MOSTRAREL\RPT\"

Public Const strCamArgs         As String = "\\SERVIDORNOVA\"
''Public Const strCamArgs         As String = "C:\RICARDO\SGI\NOVALATA\"

Public Const strPASSWORD = "PWD=;"

Public Const cCamRelEstoque = "ESTOQUE\"
Public Const cCamRelCotacaoCompras = "COTACAOCOMPRAS\"
Public Const cCamRelCotacaoVendas = "COTACAOVENDAS\"
Public Const cCamRelPedidoCompras = "PEDIDOCOMPRAS\"
Public Const cCamRelPedidoVendas = "PEDIDOVENDAS\"
Public Const cCamRelContasARC = "CONTASARC\"
Public Const cCamRelContasAPG = "CONTASAPG\"
Public Const cCamRelContasGRAF = "RELMAPACONTAREC\"
Public Const cCamRelRegMat = "ESTOQUE\"
Public Const cCamRelSupri = "RELSUPRI\"
Public Const cCamRelComercial = "RELCOMERCIAL\"
Public Const cCamRelQualidade = "RELQUALIDADE\"
Public Const cCamRelContasAPGLote = "LOTESCAPG\"
Public Const cCamRelPCP = "PCP\"
Public Const cCamRelPCP2 = "RELPCP\"
Public Const cArquivos = "ARQUIVOS\"

Public Const intAlinhEsquerda = 0
Public Const intAlinhDireita = 1
Public Const intAlinhCetro = 2

Public Const intCamnpoTextBox = 1
Public Const intCamnpoComboBox = 2
Public arrNIVEL1_PROD() As Nivel_Prod
Public arrNIVEL2_PROD() As Nivel_Prod
Public arrNIVEL3_PROD() As Nivel_Prod
Public arrNIVEL4_PROD() As Nivel_Prod
Public arrNIVEL5_PROD() As Nivel_Prod
Public arrNIVEL6_PROD() As Nivel_Prod
Public arrMenu()        As Menu_Niveis


Public Type Menu_Niveis_TipoM
       intCODIGO    As Integer
       strTEXTO     As String
       strTIPO      As String
       strCIGLA     As String
       strCIGLA2    As String
       strModulo    As String
End Type

Public Type Menu_Niveis_TipoS
       intCODIGO    As Integer
       strTEXTO     As String
       strTIPO      As String
       strCIGLA     As String
       strCIGLA2    As String
       strModulo    As String
       intQTDNIVEL  As Integer
       arrNIVEL_M() As Menu_Niveis_TipoM
End Type

Public Type Menu_Niveis
       intCODIGO    As Integer
       strTEXTO     As String
       strTIPO      As String
       strCIGLA     As String
       strCIGLA2    As String
       strModulo    As String
       intQTDNIVEL  As Integer
       arrNIVEL_S() As Menu_Niveis_TipoS
End Type

Public Type Nivel_Prod
    strCODPROD      As String
    strProdPai      As String
    strUnidadeCons  As String
    intTabConv      As Integer
    curConsUnit     As Currency
    curConsTotal    As Currency
    curCustUnitario As Currency
    curCustTotal    As Currency
    lngTIPOPRODUTO  As Long
    lngCODSETOR     As Long
End Type

'' Variaveis do Processo Produtivo

Public arrPROCESSO() As Processos

Public Type ProdSaida
    intTipo                   As Integer
    strCODPROD                As String
    strPAI                    As String
    lngUNIDMED                As Long
    curQTDESTOQUE             As Currency
End Type

Public Type ProdEntrada
    intTipo                   As Integer
    strCODPROD                As String
    strPAI                    As String
    lngUNIDMED                As Long
    curQTDESTOQUE             As Currency
    intCADENCIA               As Integer
End Type

Public Type Maquinas
    intTipo                   As Integer
    lngCODMAQ                 As Long
    lngQTDPCMIN               As Currency
    curTEMPPROD               As Currency
    curMELHORCENARIO          As Currency
    curPIORCENARIO            As Currency
    strINDICE                 As String
    strPAI                    As String
End Type

Public Type PRODUTOS
    intTipo                   As Integer
    lngIDProduto              As Long
    strPRODUTO                As String
    lngCodUniMed              As Long
    lngCODFAMMAQ              As Long
    lngTOTMAQUINAS            As Long
    lngTOTPRODENTRADA         As Long
    lngTOTPRODSAIDA           As Long
    curMELHORCENARIO          As Currency
    curPIORCENARIO            As Currency
    typMaquinas()             As Maquinas
    typProdEntrada()          As ProdEntrada
    typProdSaida()            As ProdSaida
End Type

Public Type Processos
    intORCEM                  As Long
    lngCODIGO                 As Long
    strDESCRI                 As String
    intTipo                   As Integer
    lngQTDPRODUTOS            As Long
    curMELHORCENARIO          As Currency
    curPIORCENARIO            As Currency
    typProdutos()             As PRODUTOS
End Type

'' Variaveis da arvore de produtos
Public arrPROVARV()           As PRODARVPROD

Public Type PRODARVPROD
       lngID                As Long
       lngIDPAI             As Long
       lngCODIGO            As Long
       lngCODPAI            As Long
       lngProdutoID         As Long
       lngProdutoIDPai      As Long
       lngTipo              As Long
       strPRODUTO           As String
       strProdutoPAI        As String
       strUNIDADE           As String
       lngCodUniMed         As Long
       strQTDCONS           As String
       intAction2Do         As Integer
End Type

Public Type IDEAL
       SGI_DESCRI               As String
       SGI_EFCMEDIA             As Currency
       SGI_PRODPECTEOR          As Currency
       SGI_PRODPECREAL          As Currency
       SGI_PRODDIA              As Currency
       lngPROCESSO              As Long
       lngCODMAQUINA            As Long
End Type

Public Type PRODDIA
       lngCODMAQ                As Long
       curPRODREAL              As Currency
       lngPROCESSO              As Long
End Type


'' Váriaveis para a espinha de Peixe
'' =================================
'' Nova Estrutura

Public arrESPPEIXENOVA()        As ESPINHAPAI
Public arrEFEITOS()             As EFEITOESPINHA
Public arrCAUSAS()              As CAUSASESPINHA

Public Type CAUSASESPINHA
       lngCODIGO                As Long
       lngORDEM                 As Integer
       strDESCRICAUSA           As String
       strDESCRIACAO            As String
       strDESCRIRESP            As String
       dtPrazo                  As Date
       intAction2Do             As Integer
End Type

Public Type EFEITOESPINHA
       lngID                    As Long
       lngCODIGO                As Long
       lngTipo                  As Long
       strDESCREFEITO           As String
       lngTOTCAUSAS             As Long
       intAction2Do             As Integer
       typCAUSAS()              As CAUSASESPINHA
End Type

Public Type ESPINHAPAI
       lngID                    As Long
       lngCODIGO                As Long
       lngTipo                  As Long
       strDESCRPAI              As String
       lngQTDEFEITOS            As Long
       intAction2Do             As Integer
       typEFEITOS()             As EFEITOESPINHA
End Type
'' =================================

'' Cadeia de Valor
Public Type DadosProcPadrao
    lngCodProcesso            As Long
    curQtdProdDia             As Currency
End Type


Public Type DadosProdPadrao
    curProducaoPadrao         As Currency
    curPesoPadrao             As Currency
    curMetros                 As Currency
    curGramMetro2             As Currency
    curLargura                As Currency
    curHorasPorDia            As Currency
    lngMinPorTurno            As Long
    curSegundosPorTurno       As Currency
    lngQtdPorTurno            As Long
    lngFamMaquinas            As Long
End Type


Public Enum dacEnumUpdateAction '
    dacEnumUpdateAction_Ignore = 0              ' Não faz nada
    dacEnumUpdateAction_Insert = 1              ' Insere o registro no banco
    dacEnumUpdateAction_update = 2              ' Atualiza o registro no banco
    dacEnumUpdateAction_delete = 3              ' Apaga o registro do banco
    dacEnumUpdateAction_updateTimeStamp = 4     ' Atualiza somente o timestamp (versão serial numerica) do registro
    dacEnumUpdateAction_Nulo = 5                ' Usando quando não existe registro carregado
End Enum

'' Curva Abc Cliente Quantidade
Public Type MESES
    lngQTDE                     As Long
    curVALOR                    As Currency
End Type

Public Type CAPACIDADE
    lngCODIGO                   As Long
    lngCodLinha                 As Long
    strDESCLINHA                As String
    arrMESES()                  As MESES
    lngTOTALCAPAC               As Long
    lngMEDIACAPAC               As Long
End Type

Public Type ABCCLIENTES
    lngCodClie                As Long
    strRAZAOSOC               As String
    strCNPJ                   As String
    arrCAPACIDADE()           As CAPACIDADE
    lngQTDECAPAC              As Long
End Type

Public Type VENDEDORES
    lngCODVEND                As Long
    strDESCRICAO              As String
    arrMESES()                As MESES
    curTOTALVALOR             As Currency
    curMEDIAVALOR             As Currency
End Type

Public Type ABCCLIVALOR
    lngCodClie                As Long
    strRAZAOSOC               As String
    strCNPJ                   As String
    arrVENDEDOR()             As VENDEDORES
    lngQTDVENDEDOR            As Long
End Type

'' Comparação de Preços
Public Type CPANOS
    lngANO               As Long
    strPRECO             As String
End Type

Public Type CPCAPACIDADE
    lngCodLinProd        As Long
    strDESCLIN           As String
    arrANOS()            As CPANOS
    lngQTDANOS           As Long
End Type

Public Type CPCLIENTES
    lngCodClie                As Long
    strRAZAOSOC               As String
    arrCAPACIDADE()           As CPCAPACIDADE
    lngQTDECAPACIDADE         As Long
End Type

'' Preços e Prazos
Public Type PPPRODUTOS
    lngIDProduto        As Long
    strCODPRODUTO       As String
    strDESCPROD         As String
    strPRECO            As String
End Type

Public Type PPCLIENTES
    lngCodClie                  As Long
    strRAZAOSOC                 As String
    DtUltimaCompra              As String
    lngCodCondPgto              As Long
    strDescPgto                 As String
    lngCODPEDIDO                As Long
    arrPRODUTOS()               As PPPRODUTOS
    lngQTDITENS                 As Long
End Type

'' Lista de Materiais
Public Type ListaMatCabec
    lngCodClie                  As Long
    strRAZAOSOC                 As String
    DtUltimaCompra              As String
    lngCodCondPgto              As Long
    strDescPgto                 As String
    lngCODPEDIDO                As Long
    lngQTDITENS                 As Long
End Type


'' Grupos de Linhas
Public Type GRPLinha
    lngCodGRPLinha                  As Long
    strCODLINHA                     As String
    strDESCGRPLINHA                 As String
    lngQTDTOTALGRPLINHA             As Long
    dtDATA                          As Date
    lngTOTALPECADDIA                As Long
    lngQTDTOTREGGRPLINHA            As Long
    dtDATAENTREGA                   As Date
    dtDATAENTREGAORIG               As Date
    lngCODOP                        As Long
    lngCODPED                       As Long
    lngCODOPBKP                     As Long
    lngIDProduto                    As Long
    lngIDPAI                        As Long
    strCODROTULO                    As String
    strDESCROT                      As String
    lngNECKIN                       As Long
    lngQTDOP                        As Long
    lngQTDREALPROOGOP               As Long
    lngQTDREALPROOGOPORIG           As Long
    lngQTDEREALPRODUZIDA            As Long
    lngTOTLOPS                      As Long
    lngAction2Do                    As Long
    lngLINHADOARRAY                 As Long
    lngSTATUS                       As Long
    lngCODSTATUSAPONT               As Long
    lngSTATUSORIG                   As Long
    intPROGRAMADO                   As Integer
    lngCODINTERNO                   As Long
    strCODBLOCO                     As String
    intFRACIONADO                   As Integer
    lngID_LINHA                     As Long
End Type


Public Type OPS_INCLUSAS_FOLHAS_USADAS
    lngSELECIONADA                      As Long
    lngCODLANC                          As Long
    lngCODOP                            As Long
    lngIDINTRENO                        As Long
    lngFOLHAUSADA                       As Long
    lngIDPROD                           As Long
    strCODPROD                          As String
    lngCODLIN                           As Long
    lngIDLIN                            As Long
    lngCODFOLHAUSADA                    As Long
    strDESCFOLHAUSADA                   As String
    dblESPESS                           As Double
    dblLARG                             As Double
    dblCOMP                             As Double
    lngQTDECORP                         As Long
    dblPERDPRODC                        As Double
    lngQTDEFOLHAS                       As Long
    dblPESO                             As Double
    lngQTDELATAS                        As Long
    strINDICE                           As String
    lngNECEFOLHAS                       As Long
    lngLinha                            As Long
End Type

Public Type OPS_INCLUSAS
    lngCODLINA                      As Long
    lngCIDGRPLIN                    As Long
    dtDATAPROG                      As Date
    strBLOCOOP                      As String
    lngCODOP                        As Long
    lngCODOPBKP                     As Long
    lngIDOP                         As Long
    lngCODPED                       As Long
    lngIDProduto                    As Long
    strCODROTULO                    As String
    strDESCROTULO                   As String
    dtDATAENTREGA                   As Date
    dtDATAENTREGAORIGINAL           As Date
    strTIPO                         As String
    lngAction2Do                    As Long
    intSELECIONADO                  As Integer
    lngCODSTATUS                    As Long
    lngCODSTATUSORIGINAL            As Long
    lngQTDOPORIGINAL                As Long
    lngQTDOPPROGRAMADA              As Long
    lngQTDOAPONTADAORIGINAL         As Long
    lngIDINTERNO                    As Long
    lngID_LINHA                     As Long
    intNECK                         As Integer
    lngIDARRAYLINHA                 As Long
    lngIDARRAYDIA                   As Long
    lngIDARRAYOP                    As Long
    lngCODSTATAPONT                 As Long
    intFRACIONADA                   As Integer
    lngEXPLIN                       As Long
    lngQTDFOLHASUSADAS              As Long
    strINDICE                       As String
    arrFOLHAS_USADAS()              As OPS_INCLUSAS_FOLHAS_USADAS
End Type


Public Type DIAS_LINHAS
    lngCodLinha                     As Long
    lngCODGRPLIN                    As Long
    dtDATAPROG                      As Date
    lngTOTALPECAS                   As Long
    lngTOTPROG                      As Long
    lngTOTDISP                      As Long
    lngID_PAI                       As Long
    lngQTDOPS                       As Long
    lngIDLINHA                      As Long
    strBLOCOOP                      As String
    strTIPO                         As String
    lngEXPLIN                       As Long
    arrOPS_INCLUSAS()               As OPS_INCLUSAS
End Type

Public Type LINHAS
    lngCodLinha                     As Long
    strDESCGRPLINHA                 As String
    lngMES                          As Long
    lngANO                          As Long
    lngCODGRPLIN                    As Long
    lngID_INTERNO                   As Long
    lngQTDECAPACIDADE               As Long
    lngQTDLINHAS                    As Long
    strBLOCOOP                      As String
    lngEXPLIN                       As Long
    arrDIAS_LINHA()                 As DIAS_LINHAS
End Type

Public Type arrAUX
    lngCODOP        As Long
    lngQTDOP        As Long
End Type



Public Function AbreBanco() As Boolean

On Error GoTo ErroAbre

    BD.Provider = Prov
    BD.ConnectionString = ConecStr
    If BD.State = 0 Then BD.Open
    
    AbreBanco = True
    Exit Function
    
ErroAbre:
    AbreBanco = False

End Function


