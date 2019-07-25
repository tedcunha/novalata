VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCADPROGLITO"
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

Public Property Get Codigo() As Long
    Codigo = lngCODIGO
End Property

Public Property Let Codigo(ByVal vData As Long)
    lngCODIGO = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strACESSO As String, strUsuario As String, lngIDUsuario As Long, Optional intFILIAL_PED As Integer)
    frmCADPROGLITOP.cCaminho = cCaminho
    frmCADPROGLITOP.Linha = Linha
    frmCADPROGLITOP.FILIAL = iFilial
    frmCADPROGLITOP.strACESSO = strACESSO
    frmCADPROGLITOP.strUsuario = strUsuario
    frmCADPROGLITOP.lngCodUsuaro = lngIDUsuario
    frmCADPROGLITOP.intFILIALPED = intFILIAL_PED
    frmCADPROGLITOP.Show vbModal
End Sub

Public Function AtivoDesativo(strNOMFILIAL As String) As Boolean

   AtivoDesativo = False
   
   Exit Function
   
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


Public Function PreenchComboFechamentoGrdSA() As String
        
    PreenchComboFechamentoGrdSA = ""
    PreenchComboFechamentoGrdSA = PreenchComboFechamentoGrdSA & "|#" & 0 & ";SOLDA"
    PreenchComboFechamentoGrdSA = PreenchComboFechamentoGrdSA & "|#" & 1 & ";AGRAFADO"
    PreenchComboFechamentoGrdSA = PreenchComboFechamentoGrdSA & "|#" & 2 & ";REPUXO"

End Function

