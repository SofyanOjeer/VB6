Attribute VB_Name = "srvZGUIRC10"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsADO As ADODB.Recordset
'!!!!!!!!!!!!!!!!!!!!!!!!!!  ENREGISTREMENT PARTIEL
'======================================================
Type typeYGUIRC10
 
      GUIRC1ETA     As Long
      GUIRC1AGE     As Long
      GUIRC1SER     As String * 2
      GUIRC1SSE     As String * 2
      GUIRC1OPE     As String * 3
      GUIRC1DOS     As Long
      GUIRC1DCR     As Long
      
      GUIRC1NAT     As String * 3
      GUIRC1DE1     As String * 3
      GUIRC1CP2     As String * 20
      GUIRC1MO2    As Currency

End Type

Public xYGUIRC10 As typeYGUIRC10
Public Function srvYGUIRC10_GetBuffer_ODBC(rsADO As ADODB.Recordset, lYGUIRC10 As typeYGUIRC10)
On Error GoTo Error_Handler
srvYGUIRC10_GetBuffer_ODBC = Null

lYGUIRC10.GUIRC1ETA = rsADO("GUIRC1ETA")
lYGUIRC10.GUIRC1AGE = rsADO("GUIRC1AGE")
lYGUIRC10.GUIRC1SER = rsADO("GUIRC1SER")
lYGUIRC10.GUIRC1SSE = rsADO("GUIRC1SSE")
lYGUIRC10.GUIRC1OPE = rsADO("GUIRC1OPE")
lYGUIRC10.GUIRC1DOS = rsADO("GUIRC1DOS")
lYGUIRC10.GUIRC1DCR = rsADO("GUIRC1DCR")
lYGUIRC10.GUIRC1NAT = rsADO("GUIRC1NAT")
lYGUIRC10.GUIRC1DE1 = rsADO("GUIRC1DE1")
lYGUIRC10.GUIRC1CP2 = rsADO("GUIRC1CP2")
lYGUIRC10.GUIRC1MO2 = rsADO("GUIRC1MO2")

Exit Function
Error_Handler:
srvYGUIRC10_GetBuffer_ODBC = Error


End Function


Public Function sqlZGUIRC10_Read(cnADO As ADODB.Connection, rsADO As ADODB.Recordset, lYGUIRC10 As typeYGUIRC10)
Dim xSql As String
On Error GoTo Error_Handler
sqlZGUIRC10_Read = Null

xSql = "select *  from " & paramIBM_Library_SAB & ".ZGUIRC10  " _
    & " where GUIRC1ETA = " & lYGUIRC10.GUIRC1ETA _
    & " and GUIRC1AGE = " & lYGUIRC10.GUIRC1AGE _
    & " and GUIRC1SER = '" & lYGUIRC10.GUIRC1SER & "'" _
    & " and GUIRC1SSE = '" & lYGUIRC10.GUIRC1SSE & "'" _
    & " and GUIRC1OPE = '" & lYGUIRC10.GUIRC1OPE & "'" _
    & " and GUIRC1DOS = " & lYGUIRC10.GUIRC1DOS _

Set rsADO = cnADO.Execute(xSql)
If Not rsADO.EOF Then
    srvYGUIRC10_GetBuffer_ODBC rsADO, lYGUIRC10
Else
''    srvYGUIRC10_Init lYGUIRC10
    sqlZGUIRC10_Read = "? inconnu"
End If
    
    Exit Function
Error_Handler:
    sqlZGUIRC10_Read = Error

End Function
Public Function srvYGUIRC10_Init(lYGUIRC10 As typeYGUIRC10)
lYGUIRC10.GUIRC1ETA = 0
lYGUIRC10.GUIRC1AGE = 0
lYGUIRC10.GUIRC1SER = ""
lYGUIRC10.GUIRC1SSE = ""
lYGUIRC10.GUIRC1OPE = ""
lYGUIRC10.GUIRC1DOS = 0
lYGUIRC10.GUIRC1DCR = 0
lYGUIRC10.GUIRC1DE1 = 0
lYGUIRC10.GUIRC1CP2 = 0
lYGUIRC10.GUIRC1MO2 = 0

End Function

