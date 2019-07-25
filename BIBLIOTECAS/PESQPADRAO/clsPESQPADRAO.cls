VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPESQPADRAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim lngUNIMEDCOD As Long
Dim intFILIAL    As Integer

Public Property Get FILIAL() As Integer
    FILIAL = intFILIAL
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFILIAL = vData
End Property

Public Property Get UNIMEDCOD() As Long
    UNIMEDCOD = lngUNIMEDCOD
End Property

Public Property Let UNIMEDCOD(ByVal vData As Long)
    lngUNIMEDCOD = vData
End Property

Public Function cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, arrCAMPOS As Variant, arrTABELA As Variant, strCABEC As String, Optional strForm As String, Optional arrTABELA2 As Variant, Optional boolPesqUsuario As Boolean) As String
    
    cConnect = ""
    
    frmPESQPADRAO.cCaminho = cCaminho
    frmPESQPADRAO.Linha = Linha
    frmPESQPADRAO.FILIAL = iFilial
    frmPESQPADRAO.strAcesso = strAcesso
    frmPESQPADRAO.arrCAMPOS = arrCAMPOS
    frmPESQPADRAO.arrTABELA = arrTABELA
    frmPESQPADRAO.arrTABELA2 = arrTABELA2
    frmPESQPADRAO.strCABEC = strCABEC
    frmPESQPADRAO.strForm = strForm
    frmPESQPADRAO.boolPesqUsuario = boolPesqUsuario
    frmPESQPADRAO.Show vbModal
    
    cConnect = varRETORNO
    
End Function

