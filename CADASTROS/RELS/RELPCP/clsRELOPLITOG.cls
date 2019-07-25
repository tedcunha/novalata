VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRELOPLITOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial  As Integer
Dim arrRELOP   As Variant
Dim lngCODIGO  As Long

Public Property Get CODIGO() As Long
    CODIGO = lngCODIGO
End Property

Public Property Let CODIGO(ByVal vData As Long)
    lngCODIGO = vData
End Property

Public Property Get RELOP() As Variant
    RELOP = arrRELOP
End Property

Public Property Let RELOP(ByVal vData As Variant)
    arrRELOP = vData
End Property

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmRELOPLITOG.cCaminho = cCaminho
    frmRELOPLITOG.Linha = Linha
    frmRELOPLITOG.FILIAL = iFilial
    frmRELOPLITOG.strAcesso = strAcesso
    frmRELOPLITOG.lngCodUsuario = lngIDUsuario
    frmRELOPLITOG.Show vbModal
End Sub

Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_grava

     GRAVA = False
     
     Dim I As Long
         
     adoBanco_Dados.BeginTrans
     BGRV.ActiveConnection = adoBanco_Dados
     
     If strOPER = "I" Then
     
        If IsArray(arrRELOP) Then
            For I = 1 To UBound(arrRELOP)
        
                sSql = "Insert Into SGI_RELOPDTENTLITO (" & vbCrLf
                sSql = sSql & "                                SGI_FILIAL " & vbCrLf
                sSql = sSql & "                               ,SGI_CODIGO " & vbCrLf
                sSql = sSql & "                               ,SGI_CODOP " & vbCrLf
                sSql = sSql & "                               ,SGI_DATENTREGA " & vbCrLf
                sSql = sSql & "                               ,SGI_CODPED " & vbCrLf
                sSql = sSql & "                               ,SGI_CODPROD " & vbCrLf
                sSql = sSql & "                               ,SGI_QTDE " & vbCrLf
                sSql = sSql & "                               ,SGI_IDPRODUTO " & vbCrLf
                sSql = sSql & "                               ,SGI_CODLINPROD " & vbCrLf
                sSql = sSql & "                               ,SGI_CODCLI " & vbCrLf
                sSql = sSql & "                               ,SGI_CODCOR " & vbCrLf
                sSql = sSql & "                               ,SGI_DESCRCOR " & vbCrLf
                sSql = sSql & "                     ) Values (" & vbCrLf
                sSql = sSql & "                               " & intFilial & vbCrLf
                sSql = sSql & "                              ," & lngCODIGO & vbCrLf
                sSql = sSql & "                              ," & arrRELOP(I, 1) & vbCrLf
                sSql = sSql & "                              ," & arrRELOP(I, 2) & vbCrLf
                sSql = sSql & "                              ," & arrRELOP(I, 3) & vbCrLf
                sSql = sSql & "                              ," & arrRELOP(I, 4) & vbCrLf
                sSql = sSql & "                              ," & arrRELOP(I, 5) & vbCrLf
                sSql = sSql & "                              ," & arrRELOP(I, 6) & vbCrLf
                sSql = sSql & "                              ," & arrRELOP(I, 7) & vbCrLf
                sSql = sSql & "                              ," & arrRELOP(I, 8) & vbCrLf
                sSql = sSql & "                              ," & arrRELOP(I, 9) & vbCrLf
                sSql = sSql & "                              ," & arrRELOP(I, 10) & vbCrLf
                sSql = sSql & "                              )"
        
                BGRV.CommandText = sSql
                BGRV.Execute
        
            Next I
        End If
     End If
     
     BGRV.CommandText = sSql
     BGRV.Execute
     
     adoBanco_Dados.CommitTrans
     
     GRAVA = True
     
     Exit Function
     
err_grava:
     
     adoBanco_Dados.RollbackTrans
    
     Dim objErro    As Object
     Set objErro = CreateObject("BLBCWS.clsFuncoes")
     Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, strOPER, sSql)
     Set objErro = Nothing
     
End Function

