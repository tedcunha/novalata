VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADLINHAPROD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial       As Integer
Dim intCODIGO       As Integer
Dim lngCODIGOLIN    As Long
Dim strDESCRI       As String
Dim intFILIALPED    As Integer
Dim strDESENV       As String
Dim strALTURA       As String
Dim strPERDPROC     As String
Dim arrDIMCORTE     As Variant
Dim arrSEQCORTE     As Variant
Dim arrPRODUTOS     As Variant
Dim arrFECHAMENTOS  As Variant
Dim arrFECHTPFR     As Variant
Dim arrARMAZ        As Variant
Dim strQTDECORPOS   As String
Dim arrMAQUINA      As Variant
Dim arrMESANO       As Variant
Dim arrCAPACPROD    As Variant

Public Property Get CAPACPROD() As Variant
    CAPACPROD = arrCAPACPROD
End Property

Public Property Let CAPACPROD(ByVal vData As Variant)
    arrCAPACPROD = vData
End Property

Public Property Get MESANO() As Variant
    MESANO = arrMESANO
End Property

Public Property Let MESANO(ByVal vData As Variant)
    arrMESANO = vData
End Property

Public Property Get MAQUINA() As Variant
    MAQUINA = arrMAQUINA
End Property

Public Property Let MAQUINA(ByVal vData As Variant)
    arrMAQUINA = vData
End Property

Public Property Get ARMAZ() As Variant
    ARMAZ = arrARMAZ
End Property

Public Property Let ARMAZ(ByVal vData As Variant)
    arrARMAZ = vData
End Property

Public Property Get FECHTPFR() As Variant
    FECHTPFR = arrFECHTPFR
End Property

Public Property Let FECHTPFR(ByVal vData As Variant)
    arrFECHTPFR = vData
End Property

Public Property Get FECHAMENTOS() As Variant
    FECHAMENTOS = arrFECHAMENTOS
End Property

Public Property Let FECHAMENTOS(ByVal vData As Variant)
    arrFECHAMENTOS = vData
End Property

Public Property Get PRODUTOS() As Variant
    PRODUTOS = arrPRODUTOS
End Property

Public Property Let PRODUTOS(ByVal vData As Variant)
    arrPRODUTOS = vData
End Property

Public Property Get QTDECORPOS() As String
    QTDECORPOS = strQTDECORPOS
End Property

Public Property Let QTDECORPOS(ByVal vData As String)
    strQTDECORPOS = vData
End Property

Public Property Get SEQCORTE() As Variant
    SEQCORTE = arrSEQCORTE
End Property

Public Property Let SEQCORTE(ByVal vData As Variant)
    arrSEQCORTE = vData
End Property


Public Property Get DIMCORTE() As Variant
    DIMCORTE = arrDIMCORTE
End Property

Public Property Let DIMCORTE(ByVal vData As Variant)
    arrDIMCORTE = vData
End Property

Public Property Get PERDPROC() As String
    PERDPROC = strPERDPROC
End Property

Public Property Let PERDPROC(ByVal vData As String)
    strPERDPROC = vData
End Property

Public Property Get ALTURA() As String
    ALTURA = strALTURA
End Property

Public Property Let ALTURA(ByVal vData As String)
    strALTURA = vData
End Property

Public Property Get DESENV() As String
    DESENV = strDESENV
End Property

Public Property Let DESENV(ByVal vData As String)
    strDESENV = vData
End Property

Public Property Get FILIALPED() As Integer
    FILIALPED = intFILIALPED
End Property

Public Property Let FILIALPED(ByVal vData As Integer)
    intFILIALPED = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property
Public Property Get CODIGO() As Integer
    CODIGO = intCODIGO
End Property

Public Property Let CODIGO(ByVal vData As Integer)
    intCODIGO = vData
End Property

Public Property Get CODIGOLIN() As Long
    CODIGOLIN = lngCODIGOLIN
End Property

Public Property Let CODIGOLIN(ByVal vData As Long)
    lngCODIGOLIN = vData
End Property

Public Property Get DESCRI() As String
    DESCRI = strDESCRI
End Property

Public Property Let DESCRI(ByVal vData As String)
    strDESCRI = vData
End Property


Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmCADLINHAPRODP.cCaminho = cCaminho
    frmCADLINHAPRODP.Linha = Linha
    frmCADLINHAPRODP.FILIAL = iFilial
    frmCADLINHAPRODP.strAcesso = strAcesso
    frmCADLINHAPRODP.strUSUARIO = strUSUARIO
    frmCADLINHAPRODP.Show vbModal
End Sub

Public Sub cConnectNovo(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String)
    frmCADLINHAPROD.cCaminho = cCaminho
    frmCADLINHAPROD.Linha = Linha
    frmCADLINHAPROD.FILIAL = iFilial
    frmCADLINHAPROD.strAcesso = strAcesso
    frmCADLINHAPROD.strUSUARIO = strUSUARIO
    frmCADLINHAPROD.cTipOper = "I"
    frmCADLINHAPROD.Show vbModal
End Sub


Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_Trans
     
     GRAVA = False
         
     Dim i As Integer
     
     '' Inicia transação
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Then

        sSql = ""
        
        '' Grava Sub-Seção
        sSql = "Insert Into SGI_CADLINHAPRODUTO (" & vbCrLf
        sSql = sSql & "                           SGI_FILIAL" & vbCrLf
        sSql = sSql & "                          ,SGI_CODIGO" & vbCrLf
        sSql = sSql & "                          ,SGI_CODLIN" & vbCrLf
        sSql = sSql & "                          ,SGI_DESCRI" & vbCrLf
        sSql = sSql & "                          ,SGI_FILIALPED" & vbCrLf
        sSql = sSql & "                          ,SGI_DESENV" & vbCrLf
        sSql = sSql & "                          ,SGI_ALTURA" & vbCrLf
        sSql = sSql & "                          ,SGI_PERDPROC" & vbCrLf
        sSql = sSql & "                          ,SGI_QTDECORPOS" & vbCrLf
        sSql = sSql & "                ) Values (" & vbCrLf
        sSql = sSql & "                   " & intFilial & vbCrLf
        sSql = sSql & "                  ," & intCODIGO & vbCrLf
        sSql = sSql & "                  ," & lngCODIGOLIN & vbCrLf
        sSql = sSql & "                  ,'" & Trim(strDESCRI) & "'" & vbCrLf
        sSql = sSql & "                  ," & intFILIALPED & vbCrLf
        sSql = sSql & "                  ," & strDESENV & vbCrLf
        sSql = sSql & "                  ," & strALTURA & vbCrLf
        sSql = sSql & "                  ," & strPERDPROC & vbCrLf
        sSql = sSql & "                  ," & strQTDECORPOS & vbCrLf
        
        sSql = sSql & "                   )"
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
ponto_volta:
        
        '' -------------------------------
        '' Seguencia de Corte
        If IsArray(arrSEQCORTE) Then
            For i = 1 To UBound(arrSEQCORTE)
                sSql = ""
                
                sSql = "Insert Into SGI_SEQCORTELINHA (" & vbCrLf
                sSql = sSql & "                                SGI_FILIAL" & vbCrLf
                sSql = sSql & "                               ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                               ,SGI_SEQUENCIA" & vbCrLf
                sSql = sSql & "                     ) Values (" & vbCrLf
                sSql = sSql & "                               " & intFilial & vbCrLf
                sSql = sSql & "                              ," & intCODIGO & vbCrLf
                sSql = sSql & "                              ," & arrSEQCORTE(i, 1) & vbCrLf
                sSql = sSql & "                              )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
        
        '' -------------------------------
        '' Dimensão de Corte
        If IsArray(arrDIMCORTE) Then
            For i = 1 To UBound(arrDIMCORTE)
                sSql = ""
                                
                sSql = "Insert Into SGI_MEDCORTELINHA (" & vbCrLf
                sSql = sSql & "                               SGI_FILIAL" & vbCrLf
                sSql = sSql & "                              ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                              ,SGI_ITENCOR" & vbCrLf
                sSql = sSql & "                              ,SGI_CODMEDCORT" & vbCrLf
                sSql = sSql & "                              ,SGI_INDICE" & vbCrLf
                sSql = sSql & "                              ,SGI_PADRAO" & vbCrLf
                sSql = sSql & "                              ,SGI_EXPESS" & vbCrLf
                sSql = sSql & "                              ,SGI_LARGUR" & vbCrLf
                sSql = sSql & "                              ,SGI_COMPRI" & vbCrLf
                sSql = sSql & "                              ,SGI_QTDECORPOS" & vbCrLf
                sSql = sSql & "                              ,SGI_PERDPROC" & vbCrLf
                
                sSql = sSql & "                     ) Values (" & vbCrLf
                sSql = sSql & "                               " & intFilial & vbCrLf
                sSql = sSql & "                              ," & intCODIGO & vbCrLf
                sSql = sSql & "                              ," & arrDIMCORTE(i, 1) & vbCrLf
                sSql = sSql & "                              ," & arrDIMCORTE(i, 2) & vbCrLf
                sSql = sSql & "                              ," & arrDIMCORTE(i, 3) & vbCrLf
                sSql = sSql & "                              ," & arrDIMCORTE(i, 4) & vbCrLf
                sSql = sSql & "                              ," & arrDIMCORTE(i, 5) & vbCrLf
                sSql = sSql & "                              ," & arrDIMCORTE(i, 6) & vbCrLf
                sSql = sSql & "                              ," & arrDIMCORTE(i, 7) & vbCrLf
                sSql = sSql & "                              ," & arrDIMCORTE(i, 8) & vbCrLf
                sSql = sSql & "                              ," & arrDIMCORTE(i, 9) & vbCrLf
                
                sSql = sSql & "                              )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
        '' -------------------------------
     
        '' -------------------------------
        '' Produtos
        If IsArray(arrPRODUTOS) Then
            For i = 1 To UBound(arrPRODUTOS)
                sSql = ""
                
                sSql = "Insert Into SGI_PRODUSDLINHA (" & vbCrLf
                sSql = sSql & "                                SGI_FILIAL" & vbCrLf
                sSql = sSql & "                               ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                               ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "                     ) Values (" & vbCrLf
                sSql = sSql & "                               " & intFilial & vbCrLf
                sSql = sSql & "                              ," & intCODIGO & vbCrLf
                sSql = sSql & "                              ," & arrPRODUTOS(i, 1) & vbCrLf
                sSql = sSql & "                              )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
     
     
        '' -------------------------------
        '' Fechamentos
        If IsArray(arrFECHAMENTOS) Then
            For i = 1 To UBound(arrFECHAMENTOS)
                sSql = ""
                
                sSql = "Insert Into SGI_CADLINHAPRODUTO_FECH (" & vbCrLf
                sSql = sSql & "                                SGI_FILIAL" & vbCrLf
                sSql = sSql & "                               ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                               ,SGI_COD" & vbCrLf
                sSql = sSql & "                     ) Values (" & vbCrLf
                sSql = sSql & "                               " & intFilial & vbCrLf
                sSql = sSql & "                              ," & intCODIGO & vbCrLf
                sSql = sSql & "                              ," & arrFECHAMENTOS(i) & vbCrLf
                sSql = sSql & "                              )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
     
        '' -------------------------------
        '' Fechamentos Tampa Furo
        If IsArray(arrFECHTPFR) Then
            For i = 1 To UBound(arrFECHTPFR)
                sSql = ""
                
                sSql = "Insert Into SGI_CADLINHAPRODUTO_FECHTPFR (" & vbCrLf
                sSql = sSql & "                                SGI_FILIAL" & vbCrLf
                sSql = sSql & "                               ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                               ,SGI_COD" & vbCrLf
                sSql = sSql & "                     ) Values (" & vbCrLf
                sSql = sSql & "                               " & intFilial & vbCrLf
                sSql = sSql & "                              ," & intCODIGO & vbCrLf
                sSql = sSql & "                              ," & arrFECHTPFR(i) & vbCrLf
                sSql = sSql & "                              )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
     
        '' -------------------------------
        '' Fechamentos de Armazenamento
        If IsArray(arrARMAZ) Then
            For i = 1 To UBound(arrARMAZ)
                sSql = ""
                
                sSql = "Insert Into SGI_CADLINHAPRODUTO_FECHTPFRARM (" & vbCrLf
                sSql = sSql & "                                SGI_FILIAL" & vbCrLf
                sSql = sSql & "                               ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                               ,SGI_COD" & vbCrLf
                sSql = sSql & "                               ,SGI_QTDELATAS" & vbCrLf
                sSql = sSql & "                     ) Values (" & vbCrLf
                sSql = sSql & "                               " & intFilial & vbCrLf
                sSql = sSql & "                              ," & intCODIGO & vbCrLf
                sSql = sSql & "                              ," & i & vbCrLf
                sSql = sSql & "                              ," & Trim(arrARMAZ(i)) & vbCrLf
                sSql = sSql & "                              )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
        
        '' -------------------------------
        '' Máquinas
        If IsArray(arrMAQUINA) Then
            For i = 1 To UBound(arrMAQUINA)
                sSql = ""
                
                sSql = "Insert Into SGI_MAQULIN (" & vbCrLf
                sSql = sSql & "                                SGI_FILIAL" & vbCrLf
                sSql = sSql & "                               ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                               ,SGI_CODLIN" & vbCrLf
                sSql = sSql & "                               ,SGI_CODMAQ" & vbCrLf
                sSql = sSql & "                               ,SGI_QTDEPORHORA" & vbCrLf
                
                sSql = sSql & "                     ) Values (" & vbCrLf
                sSql = sSql & "                               " & intFilial & vbCrLf
                sSql = sSql & "                              ," & intCODIGO & vbCrLf
                sSql = sSql & "                              ," & lngCODIGOLIN & vbCrLf
                sSql = sSql & "                              ," & arrMAQUINA(i, 1) & vbCrLf
                sSql = sSql & "                              ," & arrMAQUINA(i, 2) & vbCrLf
                
                sSql = sSql & "                              )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
        '' -------------------------------
        
     
        '' -------------------------------
        '' Mês/Ano
        If IsArray(arrMESANO) Then
            For i = 1 To UBound(arrMESANO)
                sSql = ""
                
                sSql = "Insert Into SGI_MAQULIN_MESANO (" & vbCrLf
                sSql = sSql & "                                SGI_FILIAL" & vbCrLf
                sSql = sSql & "                               ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                               ,SGI_CODLIN" & vbCrLf
                sSql = sSql & "                               ,SGI_CODMAQ" & vbCrLf
                sSql = sSql & "                               ,SGI_MES" & vbCrLf
                sSql = sSql & "                               ,SGI_ANO" & vbCrLf
                sSql = sSql & "                               ,SGI_INDICE" & vbCrLf
                sSql = sSql & "                               ,SGI_QTDECAPACIDADE " & vbCrLf
                
                sSql = sSql & "                     ) Values (" & vbCrLf
                sSql = sSql & "                               " & intFilial & vbCrLf
                sSql = sSql & "                              ," & intCODIGO & vbCrLf
                sSql = sSql & "                              ," & lngCODIGOLIN & vbCrLf
                sSql = sSql & "                              ," & arrMESANO(i, 3) & vbCrLf
                sSql = sSql & "                              ," & arrMESANO(i, 1) & vbCrLf
                sSql = sSql & "                              ," & arrMESANO(i, 2) & vbCrLf
                sSql = sSql & "                              ,'" & arrMESANO(i, 4) & "'" & vbCrLf
                sSql = sSql & "                              ," & arrMESANO(i, 5) & vbCrLf
                
                sSql = sSql & "                              )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
        '' -------------------------------
     
     
        '' -------------------------------
        '' Capacidade Produtiva
        If IsArray(arrCAPACPROD) Then
            For i = 1 To UBound(arrCAPACPROD)
                sSql = ""
                
                sSql = "Insert Into SGI_MAQULIN_CAPAC (" & vbCrLf
                sSql = sSql & "                                SGI_FILIAL" & vbCrLf
                sSql = sSql & "                               ,SGI_CODIGO" & vbCrLf
                sSql = sSql & "                               ,SGI_CODLIN" & vbCrLf
                sSql = sSql & "                               ,SGI_CODMAQ" & vbCrLf
                sSql = sSql & "                               ,SGI_DATA" & vbCrLf
                sSql = sSql & "                               ,SGI_HORAINI" & vbCrLf
                sSql = sSql & "                               ,SGI_HORAPARADA" & vbCrLf
                sSql = sSql & "                               ,SGI_HORAFIN" & vbCrLf
                sSql = sSql & "                               ,SGI_HORATOTAL" & vbCrLf
                sSql = sSql & "                               ,SGI_TOTALPECAS" & vbCrLf
                sSql = sSql & "                               ,SGI_CODINDICE" & vbCrLf
                sSql = sSql & "                               ,SGI_DIASEM" & vbCrLf
                sSql = sSql & "                               ,SGI_ATIVO" & vbCrLf
                sSql = sSql & "                               ,SGI_INDICE" & vbCrLf
                
                sSql = sSql & "                     ) Values (" & vbCrLf
                
                sSql = sSql & "                               " & intFilial & vbCrLf
                sSql = sSql & "                              ," & intCODIGO & vbCrLf
                sSql = sSql & "                              ," & lngCODIGOLIN & vbCrLf
                sSql = sSql & "                              ," & arrCAPACPROD(i, 11) & vbCrLf
                sSql = sSql & "                              ," & arrCAPACPROD(i, 1) & vbCrLf
                sSql = sSql & "                              ," & arrCAPACPROD(i, 2) & vbCrLf
                sSql = sSql & "                              ," & arrCAPACPROD(i, 3) & vbCrLf
                sSql = sSql & "                              ," & arrCAPACPROD(i, 4) & vbCrLf
                sSql = sSql & "                              ," & arrCAPACPROD(i, 5) & vbCrLf
                sSql = sSql & "                              ," & arrCAPACPROD(i, 6) & vbCrLf
                sSql = sSql & "                              ," & arrCAPACPROD(i, 7) & vbCrLf
                sSql = sSql & "                              ," & arrCAPACPROD(i, 8) & vbCrLf
                sSql = sSql & "                              ," & arrCAPACPROD(i, 9) & vbCrLf
                sSql = sSql & "                              ," & arrCAPACPROD(i, 10) & vbCrLf
                
                sSql = sSql & "                              )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
            Next i
        End If
        '' -------------------------------
     
     ElseIf strOPER = "A" Then
     
        sSql = ""
                        
        sSql = sSql & "Delete From SGI_MEDCORTELINHA " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        sSql = ""
                        
        sSql = sSql & "Delete From SGI_SEQCORTELINHA " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        sSql = ""
                        
        sSql = sSql & "Delete From SGI_PRODUSDLINHA " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        sSql = ""
                        
        sSql = sSql & "Delete From SGI_CADLINHAPRODUTO_FECH " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        sSql = ""
                        
        sSql = sSql & "Delete From SGI_CADLINHAPRODUTO_FECHTPFR " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        sSql = ""
                        
        sSql = sSql & "Delete From SGI_CADLINHAPRODUTO_FECHTPFRARM " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' ---------------------------------
        '' Máquina
        sSql = ""
                        
        sSql = sSql & "Delete From SGI_MAQULIN " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' ---------------------------------
        
        '' ---------------------------------
        '' Mês/Ano
        sSql = ""
                        
        sSql = sSql & "Delete From SGI_MAQULIN_MESANO " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' ---------------------------------
        
        '' ---------------------------------
        '' Capacidade Produtiva
        sSql = ""
                        
        sSql = sSql & "Delete From SGI_MAQULIN_CAPAC " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' ---------------------------------
        
        sSql = ""
        
        sSql = "Update  SGI_CADLINHAPRODUTO Set " & vbCrLf
        sSql = sSql & "                                SGI_CODLIN     = " & lngCODIGOLIN & vbCrLf
        sSql = sSql & "                               ,SGI_DESCRI     = '" & Trim(strDESCRI) & "'" & vbCrLf
        sSql = sSql & "                               ,SGI_DESENV     = " & strDESENV & vbCrLf
        sSql = sSql & "                               ,SGI_ALTURA     = " & strALTURA & vbCrLf
        sSql = sSql & "                               ,SGI_PERDPROC   = " & strPERDPROC & vbCrLf
        sSql = sSql & "                               ,SGI_FILIALPED  = " & intFILIALPED & vbCrLf
        sSql = sSql & "                               ,SGI_QTDECORPOS = " & strQTDECORPOS & vbCrLf
        
        sSql = sSql & " Where " & vbCrLf
        sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "   And SGI_CODIGO = " & intCODIGO
     
        BGRV.CommandText = sSql
        BGRV.Execute
        
        GoTo ponto_volta
     
     ElseIf strOPER = "E" Then
     
        sSql = ""
                        
        sSql = sSql & "Delete From SGI_MEDCORTELINHA " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        
        sSql = ""
                        
        sSql = sSql & "Delete From SGI_SEQCORTELINHA " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        sSql = ""
                        
        sSql = sSql & "Delete From SGI_CADLINHAPRODUTO_FECH " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        sSql = ""
                        
        sSql = sSql & "Delete From SGI_CADLINHAPRODUTO_FECHFR " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        sSql = ""
                        
        sSql = sSql & "Delete From SGI_CADLINHAPRODUTO_FECHTPFRARM " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        
        '' ---------------------------------
        '' Máquina
        sSql = ""
                        
        sSql = sSql & "Delete From SGI_MAQULIN " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' ---------------------------------
        
        '' ---------------------------------
        '' Mês/Ano
        sSql = ""
                        
        sSql = sSql & "Delete From SGI_MAQULIN_MESANO " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' ---------------------------------
        
        '' ---------------------------------
        '' Capacidade Produtiva
        sSql = ""
                        
        sSql = sSql & "Delete From SGI_MAQULIN_CAPAC " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL = " & intFilial & vbCrLf
        sSql = sSql & "        And SGI_CODIGO = " & intCODIGO
        
        BGRV.CommandText = sSql
        BGRV.Execute
        '' ---------------------------------
        
        sSql = ""
        
        sSql = "Delete from SGI_CADLINHAPRODUTO " & vbCrLf
        sSql = sSql & "       Where " & vbCrLf
        sSql = sSql & "             SGI_CODIGO = " & intCODIGO & vbCrLf
        sSql = sSql & "         And SGI_FILIAL = " & intFilial
        
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
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    
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
   
   Dim intReg       As Integer
   Dim i            As Integer
   Dim intQTDREG    As Integer
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADLINHAPRODUTO " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_CODIGO = " & intCODIGO & vbCrLf
   sSql = sSql & "   And SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados, adOpenDynamic
   
   If Not BREC.EOF Then
      
      intCODIGO = BREC!SGI_CODIGO
      lngCODIGOLIN = BREC!SGI_CODLIN
      strDESCRI = BREC!SGI_DESCRI
      intFILIALPED = BREC!SGI_FILIALPED
      If Not IsNull(BREC!SGI_DESENV) Then strDESENV = BREC!SGI_DESENV
      If Not IsNull(BREC!SGI_ALTURA) Then strALTURA = BREC!SGI_ALTURA
      If Not IsNull(BREC!SGI_PERDPROC) Then strPERDPROC = BREC!SGI_PERDPROC
      If Not IsNull(BREC!SGI_QTDECORPOS) Then strQTDECORPOS = BREC!SGI_QTDECORPOS
      
      Carrega_campos = True
   End If
   BREC.Close

    '' -----------------------------------------------------------
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_SEQCORTELINHA " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO    = " & intCODIGO & vbCrLf
    sSql = sSql & "Order By SGI_SEQUENCIA"
    
    BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC3.EOF Then
    
       intQTDREG = 0
       Do While Not BREC3.EOF
          intQTDREG = intQTDREG + 1
          BREC3.MoveNext
       Loop
    
       ReDim arrSEQCORTE(1 To intQTDREG, 1 To 1) As Variant
       
       intQTDREG = 1
       BREC3.MoveFirst
       
       Do While Not BREC3.EOF
          arrSEQCORTE(intQTDREG, 1) = BREC3!SGI_SEQUENCIA
          intQTDREG = (intQTDREG + 1)
          BREC3.MoveNext
       Loop
    
    End If
    BREC3.Close

    
    '' -----------------------------------------------------------
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_MEDCORTELINHA " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO    = " & intCODIGO & vbCrLf
    sSql = sSql & "Order By SGI_ITENCOR"
    
    BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC3.EOF Then
    
       intQTDREG = 0
       Do While Not BREC3.EOF
          intQTDREG = intQTDREG + 1
          BREC3.MoveNext
       Loop
    
       ReDim arrDIMCORTE(1 To intQTDREG, 1 To 8) As Variant
       
       intQTDREG = 1
       BREC3.MoveFirst
       
       Do While Not BREC3.EOF
          arrDIMCORTE(intQTDREG, 1) = BREC3!SGI_ITENCOR
          arrDIMCORTE(intQTDREG, 2) = BREC3!SGI_CODMEDCORT
          arrDIMCORTE(intQTDREG, 3) = BREC3!SGI_PADRAO
          
          If Not IsNull(BREC3!SGI_EXPESS) Then arrDIMCORTE(intQTDREG, 4) = Format(BREC3!SGI_EXPESS, "#,##0.00")
          If Not IsNull(BREC3!SGI_LARGUR) Then arrDIMCORTE(intQTDREG, 5) = Format(BREC3!SGI_LARGUR, "#,##0.00")
          If Not IsNull(BREC3!SGI_COMPRI) Then arrDIMCORTE(intQTDREG, 6) = Format(BREC3!SGI_COMPRI, "#,##0.00")
          If Not IsNull(BREC3!SGI_QTDECORPOS) Then arrDIMCORTE(intQTDREG, 7) = Format(BREC3!SGI_QTDECORPOS, "#,##0.00")
          If Not IsNull(BREC3!SGI_PERDPROC) Then arrDIMCORTE(intQTDREG, 8) = Format(BREC3!SGI_PERDPROC, "#,##0.00")
          
          intQTDREG = (intQTDREG + 1)
          BREC3.MoveNext
       Loop
    
    End If
    BREC3.Close


    '' -----------------------------------------------------------
    '' Produtos
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       *" & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_PRODUSDLINHA " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL    = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO    = " & intCODIGO & vbCrLf
    
    BREC3.Open sSql, adoBanco_Dados, adOpenDynamic
    
    If Not BREC3.EOF Then
    
       intQTDREG = 0
       Do While Not BREC3.EOF
          intQTDREG = intQTDREG + 1
          BREC3.MoveNext
       Loop
    
       ReDim arrPRODUTOS(1 To intQTDREG, 1 To 1) As Variant
       
       intQTDREG = 1
       BREC3.MoveFirst
       
       Do While Not BREC3.EOF
          arrPRODUTOS(intQTDREG, 1) = BREC3!SGI_IDPRODUTO
          intQTDREG = (intQTDREG + 1)
          BREC3.MoveNext
       Loop
    
    End If
    BREC3.Close


    '' ==================================================================
    '' Fechamentos
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO_FECH" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & intCODIGO
    
    BREC6.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC6.EOF() Then
        intQTDREG = 0
        Do While Not BREC6.EOF()
            intQTDREG = (intQTDREG + 1)
            BREC6.MoveNext
        Loop
        
        ReDim arrFECHAMENTOS(1 To intQTDREG) As String
        
        intQTDREG = 1
        BREC6.MoveFirst
        
        Do While Not BREC6.EOF()
            arrFECHAMENTOS(intQTDREG) = BREC6!SGI_COD
            BREC6.MoveNext
            intQTDREG = (intQTDREG + 1)
        Loop
    End If
    BREC6.Close
    '' ==================================================================


    '' ==================================================================
    '' Fechamentos Tampa Furo
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO_FECHTPFR" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & intCODIGO
    
    BREC6.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC6.EOF() Then
        intQTDREG = 0
        Do While Not BREC6.EOF()
            intQTDREG = (intQTDREG + 1)
            BREC6.MoveNext
        Loop
        
        ReDim arrFECHTPFR(1 To intQTDREG) As String
        
        intQTDREG = 1
        BREC6.MoveFirst
        
        Do While Not BREC6.EOF()
            arrFECHTPFR(intQTDREG) = BREC6!SGI_COD
            BREC6.MoveNext
            intQTDREG = (intQTDREG + 1)
        Loop
    End If
    BREC6.Close
    '' ==================================================================

    '' ==================================================================
    '' Fechamentos Tampa Furo Armazenamento
    sSql = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADLINHAPRODUTO_FECHTPFRARM" & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & intCODIGO
    
    BREC6.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC6.EOF() Then
        intQTDREG = 0
        Do While Not BREC6.EOF()
            intQTDREG = (intQTDREG + 1)
            BREC6.MoveNext
        Loop
        
        ReDim arrARMAZ(1 To intQTDREG) As String
        
        intQTDREG = 1
        BREC6.MoveFirst
        
        Do While Not BREC6.EOF()
            arrARMAZ(intQTDREG) = BREC6!SGI_QTDELATAS
            BREC6.MoveNext
            intQTDREG = (intQTDREG + 1)
        Loop
    End If
    BREC6.Close
    '' ==================================================================

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


Public Function Carrega_CADLINHAPRODUTO() As Boolean

   Carrega_CADLINHAPRODUTO = False
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       *" & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADLINHAPRODUTO " & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial
   
   BREC.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC.EOF Then Carrega_CADLINHAPRODUTO = True
   BREC.Close

End Function

Public Function PreenchComboFechamentoTampaFuro() As String

    If BREC10.State = 1 Then BREC10.Close
    
    PreenchComboFechamentoTampaFuro = ""
    
    sSql = ""
    
    sSql = "Select" & vbCrLf
    sSql = sSql & "       SGI_CODIGO" & vbCrLf
    sSql = sSql & "      ,SGI_DESCRI" & vbCrLf
    
    sSql = sSql & "  From" & vbCrLf
    sSql = sSql & "       SGI_CADFECHAM" & vbCrLf
    
    sSql = sSql & " Where" & vbCrLf
    sSql = sSql & "       SGI_FILIAL     = " & FILIAL & vbCrLf
    sSql = sSql & "Order By SGI_CODIGO"
    
    BREC10.Open sSql, adoBanco_Dados, adOpenDynamic
    Do While Not BREC10.EOF()
       PreenchComboFechamentoTampaFuro = PreenchComboFechamentoTampaFuro & "|#" & BREC10!SGI_CODIGO & ";" & Trim(BREC10!SGI_DESCRI)
       BREC10.MoveNext
    Loop
    BREC10.Close
    
End Function

Public Function PreenchComboMes() As String
        
    PreenchComboMes = ""
    PreenchComboMes = PreenchComboMes & "|#" & 1 & ";Janeiro"
    PreenchComboMes = PreenchComboMes & "|#" & 2 & ";Fevereiro"
    PreenchComboMes = PreenchComboMes & "|#" & 3 & ";Março"
    PreenchComboMes = PreenchComboMes & "|#" & 4 & ";Abril"
    PreenchComboMes = PreenchComboMes & "|#" & 5 & ";Maio"
    PreenchComboMes = PreenchComboMes & "|#" & 6 & ";Junho"
    PreenchComboMes = PreenchComboMes & "|#" & 7 & ";Julho"
    PreenchComboMes = PreenchComboMes & "|#" & 8 & ";Agosto"
    PreenchComboMes = PreenchComboMes & "|#" & 9 & ";Setembro"
    PreenchComboMes = PreenchComboMes & "|#" & 10 & ";Outubro"
    PreenchComboMes = PreenchComboMes & "|#" & 11 & ";Novenbro"
    PreenchComboMes = PreenchComboMes & "|#" & 12 & ";Dezembro"

End Function

Public Function PreenchComboAno() As String
        
    PreenchComboAno = ""
    PreenchComboAno = PreenchComboAno & "|#" & 2016 & ";2016"
    PreenchComboAno = PreenchComboAno & "|#" & 2017 & ";2017"
    PreenchComboAno = PreenchComboAno & "|#" & 2018 & ";2018"
    PreenchComboAno = PreenchComboAno & "|#" & 2019 & ";2019"
    PreenchComboAno = PreenchComboAno & "|#" & 2020 & ";2020"
    PreenchComboAno = PreenchComboAno & "|#" & 2021 & ";2021"
    PreenchComboAno = PreenchComboAno & "|#" & 2022 & ";2022"
    PreenchComboAno = PreenchComboAno & "|#" & 2023 & ";2023"
    PreenchComboAno = PreenchComboAno & "|#" & 2024 & ";2024"
    PreenchComboAno = PreenchComboAno & "|#" & 2025 & ";2025"
    PreenchComboAno = PreenchComboAno & "|#" & 2026 & ";2026"
    
End Function

