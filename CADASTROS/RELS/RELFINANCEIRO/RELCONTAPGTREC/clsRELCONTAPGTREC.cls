VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRELCONTAPGTREC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial       As Integer

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmRELCONTAPGREC.cCaminho = cCaminho
    frmRELCONTAPGREC.Linha = Linha
    frmRELCONTAPGREC.FILIAL = iFilial
    frmRELCONTAPGREC.strAcesso = strAcesso
    frmRELCONTAPGREC.Show vbModal
End Sub

Public Sub PreencheComboFornec(Combo As Variant)

     sSql = "Select " & vbCrLf
     sSql = sSql & "       FORN.* " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       SGI_CADFORNEC  FORN " & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       FORN.SGI_FILIAL  = " & intFilial & vbCrLf
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_CODIGO) & Space(10 - Len(Trim(BREC!SGI_CODIGO))) & " - " & Trim(BREC!SGI_RAZAOSOC)
        Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
        BREC.MoveNext
     Loop
    
     BREC.Close
    
End Sub

Public Sub PreencheCombo(Combo As Variant, strTabela As String)

     Combo.Clear
     
     sSql = "Select " & vbCrLf
     sSql = sSql & "       * " & vbCrLf
     sSql = sSql & "  From " & vbCrLf
     sSql = sSql & "       " & strTabela & vbCrLf
     sSql = sSql & " Where " & vbCrLf
     sSql = sSql & "       SGI_FILIAL  = " & intFilial & vbCrLf
     If UCase(Trim(strTabela)) = "SGI_CADTIPOPGTO" Then
        sSql = sSql & "   And SGI_SINAL  = '-'" & vbCrLf
     End If
     
     BREC.Open sSql, adoBanco_Dados, adOpenDynamic
     
     Do While Not BREC.EOF
        Combo.AddItem Trim(BREC!SGI_CODIGO) & Space(10 - Len(Trim(BREC!SGI_CODIGO))) & " - " & Trim(BREC!SGI_DESCRICAO)
        Combo.ItemData(Combo.NewIndex) = BREC!SGI_CODIGO
        BREC.MoveNext
     Loop
    
     ''If Combo.ListCount > 0 Then Combo.ListIndex = 0
    
     BREC.Close
    
End Sub

