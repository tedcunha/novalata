VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADCARGACONJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial           As Integer
Dim lngCODIGO           As Long

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

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUsuario As String, lngIDUsuario As Long, Optional intFILIAL_PED As Integer)
    frmCADCARGACONJP.cCaminho = cCaminho
    frmCADCARGACONJP.Linha = Linha
    frmCADCARGACONJP.FILIAL = iFilial
    frmCADCARGACONJP.strAcesso = strAcesso
    frmCADCARGACONJP.strUsuario = strUsuario
    frmCADCARGACONJP.lngCodUsuaro = lngIDUsuario
    frmCADCARGACONJP.intFILIALPED = intFILIAL_PED
    frmCADCARGACONJP.Show vbModal
End Sub

