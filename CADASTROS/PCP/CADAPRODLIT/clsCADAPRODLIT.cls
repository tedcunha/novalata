VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADAPRODLIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial           As Integer
Dim lngCODIGO           As Long
Dim strCODMAQ           As String
Dim strCODTURN          As String
Dim strCODOPER          As String
Dim strDTLANCT          As String
Dim arrAPONT            As Variant
Dim arrPARADAS          As Variant

Public Property Get PARADAS() As Variant
    PARADAS = arrPARADAS
End Property

Public Property Let PARADAS(ByVal vData As Variant)
    arrPARADAS = vData
End Property

Public Property Get APONT() As Variant
    APONT = arrAPONT
End Property

Public Property Let APONT(ByVal vData As Variant)
    arrAPONT = vData
End Property

Public Property Get DTLANCT() As String
    DTLANCT = strDTLANCT
End Property

Public Property Let DTLANCT(ByVal vData As String)
    strDTLANCT = vData
End Property

Public Property Get CODOPER() As String
    CODOPER = strCODOPER
End Property

Public Property Let CODOPER(ByVal vData As String)
    strCODOPER = vData
End Property

Public Property Get CODTURN() As String
    CODTURN = strCODTURN
End Property

Public Property Let CODTURN(ByVal vData As String)
    strCODTURN = vData
End Property

Public Property Get CODMAQ() As String
    CODMAQ = strCODMAQ
End Property

Public Property Let CODMAQ(ByVal vData As String)
    strCODMAQ = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Property Get Codigo() As Long
    Codigo = lngCODIGO
End Property

Public Property Let Codigo(ByVal vData As Long)
    lngCODIGO = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strACESSO As String, strUsuario As String, lngIDUsuario As Long, Optional intFILIAL_PED As Integer)
    frmCADAPRODLITP.cCaminho = cCaminho
    frmCADAPRODLITP.Linha = Linha
    frmCADAPRODLITP.FILIAL = iFilial
    frmCADAPRODLITP.strACESSO = strACESSO
    frmCADAPRODLITP.strUsuario = strUsuario
    frmCADAPRODLITP.lngCodUsuaro = lngIDUsuario
    frmCADAPRODLITP.intFILIALPED = intFILIAL_PED
    frmCADAPRODLITP.Show vbModal
End Sub

Public Function AtivoDesativo(strNOMFILIAL As String) As Boolean

   AtivoDesativo = False
   
   sSql = ""
   
   sSql = "Select " & vbCrLf
   sSql = sSql & "       * " & vbCrLf
   sSql = sSql & "  From " & vbCrLf
   sSql = sSql & "       SGI_CADAPPRODUCAO" & strNOMFILIAL & vbCrLf
   sSql = sSql & " Where " & vbCrLf
   sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
   
   BREC.Open sSql, adoBanco_Dados, adOpenDynamic
   If Not BREC.EOF Then AtivoDesativo = True
   BREC.Close

End Function

