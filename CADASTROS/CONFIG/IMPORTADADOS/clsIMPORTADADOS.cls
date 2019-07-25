VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIMPORTADADOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmIMPORTADADOS.cCaminho = cCaminho
    frmIMPORTADADOS.Linha = Linha
    frmIMPORTADADOS.FILIAL = iFilial
    frmIMPORTADADOS.strAcesso = strAcesso
    frmIMPORTADADOS.Show vbModal
End Sub

