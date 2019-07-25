VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRELROTCORES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intFilial           As Integer
Dim arrRELROTSEMCORES   As Variant

Public Property Get FILIAL() As Integer
    FILIAL = intFilial
End Property

Public Property Let FILIAL(ByVal vData As Integer)
    intFilial = vData
End Property

Public Property Get RELROTSEMCORES() As Variant
    RELROTSEMCORES = arrRELROTSEMCORES
End Property

Public Property Let RELROTSEMCORES(ByVal vData As Variant)
    arrRELROTSEMCORES = vData
End Property

Public Sub cConnect(cCaminho As String, Linha As Variant, iFilial As Integer, strAcesso As String, strUSUARIO As String, lngIDUsuario As Long)
    frmRELROTCORES.cCaminho = cCaminho
    frmRELROTCORES.Linha = Linha
    frmRELROTCORES.FILIAL = iFilial
    frmRELROTCORES.strAcesso = strAcesso
    frmRELROTCORES.lngCodUsuario = lngIDUsuario
    frmRELROTCORES.Show vbModal
End Sub


Public Function GRAVA(strOPER As String) As Boolean

On Error GoTo err_grava
     
     GRAVA = False
     
     Dim I      As Integer
     Dim sValor As String
         
     BGRV.ActiveConnection = adoBanco_Dados
     adoBanco_Dados.BeginTrans
     
     If strOPER = "I" Then
       
        
        If IsArray(arrRELROTSEMCORES) Then
            For I = 1 To UBound(arrRELROTSEMCORES)
            
                '' Grava
                sSql = "Insert Into SGI_RELROTULOS (" & vbCrLf
                sSql = sSql & "                        SGI_FILIAL" & vbCrLf
                sSql = sSql & "                       ,SGI_IDPRODUTO" & vbCrLf
                sSql = sSql & "              ) Values (" & vbCrLf
                sSql = sSql & "                            " & intFilial & vbCrLf
                sSql = sSql & "                           ," & arrRELROTSEMCORES(I) & vbCrLf
                
                sSql = sSql & "                           )"
                
                BGRV.CommandText = sSql
                BGRV.Execute
            
            Next I
        End If
        
     ElseIf strOPER = "E" Then
     
        '' Exclui
        sSql = "Delete From SGI_RELROTULOS " & vbCrLf
        sSql = sSql & "      Where " & vbCrLf
        sSql = sSql & "            SGI_FILIAL    = " & intFilial
        
        BGRV.CommandText = sSql
        BGRV.Execute
     
     
     End If
     
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

