VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRELVENDEDORES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial  As Integer

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmRELVENDEDORES.cCaminho = cCaminho
    frmRELVENDEDORES.Linha = Linha
    frmRELVENDEDORES.FILIAL = iFilial
    frmRELVENDEDORES.strAcesso = strAcesso
    frmRELVENDEDORES.lngCodUsuario = lngIDUsuario
    frmRELVENDEDORES.Show vbModal
End Sub


Public Function PegaVendedor(strCodVendedor As Long) As String

    PegaVendedor = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADVENDEDOR " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODIGO = " & strCodVendedor
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PegaVendedor = BREC!SGI_DESCRICAO
    BREC.Close

End Function

Public Function PegaIDVendedor(strCodUsuario As Long) As String

    PegaIDVendedor = ""
    
    sSql = "Select " & vbCrLf
    sSql = sSql & "       * " & vbCrLf
    sSql = sSql & "  From " & vbCrLf
    sSql = sSql & "       SGI_CADVENDEDOR " & vbCrLf
    sSql = sSql & " Where " & vbCrLf
    sSql = sSql & "       SGI_FILIAL     = " & intFilial & vbCrLf
    sSql = sSql & "   And SGI_CODUSUARIO = " & Trim(strCodUsuario)
    
    BREC.Open sSql, adoBanco_Dados, adOpenDynamic
    If Not BREC.EOF() Then PegaIDVendedor = Trim(Str(BREC!SGI_CODIGO))
    BREC.Close

End Function

