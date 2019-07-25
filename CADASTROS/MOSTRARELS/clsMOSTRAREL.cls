VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMOSTRAREL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Sub REL(intFilial As Integer, _
               strQUERYSQL As String, _
               strRELNOME As String, _
               variLinha As Variant, _
               intORI As Integer, _
               strCABEC1 As String, _
               strCABEC2 As String, _
               Optional booArvoreSN As Boolean, _
               Optional strConn As String, _
               Optional boolView As Boolean = True)
             ''                  01,                    02,                   03,                   04,                05,                  06,                  07,                              08,                         09,                                  10
On Error GoTo Err_Desc

    frmRELSERV.intFilial = intFilial
    frmRELSERV.strQUERYSQL = strQUERYSQL
    frmRELSERV.strRELNOME = strRELNOME
    frmRELSERV.Linha = variLinha
    frmRELSERV.intORIENTATION = intORI
    frmRELSERV.strCABEC1 = strCABEC1
    frmRELSERV.strCABEC2 = strCABEC2
    frmRELSERV.boolArvoreSN = booArvoreSN
    frmRELSERV.strConn = strConn
    frmRELSERV.boolView = boolView
    
    frmRELSERV.Show vbModal
    
    Exit Sub
    
Err_Desc:
    
    Dim objErro    As Object
    Set objErro = CreateObject("BLBCWS.clsFuncoes")
    Call objErro.Sub_DescErro(Str(Err.Number), Err.Description, "", sSql)
    Set objErro = Nothing
    
End Sub



