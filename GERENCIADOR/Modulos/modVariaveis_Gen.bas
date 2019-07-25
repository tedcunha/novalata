Attribute VB_Name = "modVariaveis_Gen"
Option Explicit
Public BD               As New ADODB.Connection
Public BGRV             As New ADODB.Command

Public BREC             As New ADODB.Recordset
Public BREC2            As New ADODB.Recordset
Public BREC3            As New ADODB.Recordset
Public BREC4            As New ADODB.Recordset
Public BREC5            As New ADODB.Recordset
Public BREC6            As New ADODB.Recordset
Public BREC7            As New ADODB.Recordset
Public BREC8            As New ADODB.Recordset
Public BREC9            As New ADODB.Recordset
Public BREC10           As New ADODB.Recordset
Public BREC11           As New ADODB.Recordset
Public BREC12           As New ADODB.Recordset
Public BRECATU          As New ADODB.Recordset

Public sSql             As String
Public strSTRCONNECT    As String
Public strNOMCOMP       As String
Public lngCODVENDEDOR   As Long
Public strNOMVENDEDOR   As String
Public intFILIAL        As Integer

