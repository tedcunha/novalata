VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMOVESTOQUE"
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
    frmMOVESTOQUE.cCaminho = cCaminho
    frmMOVESTOQUE.Linha = Linha
    frmMOVESTOQUE.FILIAL = iFilial
    frmMOVESTOQUE.strAcesso = strAcesso
    frmMOVESTOQUE.lngCodUsuario = lngIDUsuario
    frmMOVESTOQUE.Show vbModal
End Sub


