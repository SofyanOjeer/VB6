Attribute VB_Name = "rsZGUIRC10"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
'!!!!!!!!!!!!!!!!!!!!!!!!!!  ENREGISTREMENT PARTIEL
'======================================================
Type typeZGUIRC10
 
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

Public Function rsZGUIRC10_Read(lZGUIRC10 As typeZGUIRC10)
Dim xSQL As String
On Error GoTo Error_Handler
rsZGUIRC10_Read = Null

xSQL = "select *  from " & paramIBM_Library_SAB & ".ZGUIRC10  " _
    & " where GUIRC1ETA = " & lZGUIRC10.GUIRC1ETA _
    & " and GUIRC1AGE = " & lZGUIRC10.GUIRC1AGE _
    & " and GUIRC1SER = '" & lZGUIRC10.GUIRC1SER & "'" _
    & " and GUIRC1SSE = '" & lZGUIRC10.GUIRC1SSE & "'" _
    & " and GUIRC1OPE = '" & lZGUIRC10.GUIRC1OPE & "'" _
    & " and GUIRC1DOS = " & lZGUIRC10.GUIRC1DOS _

Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    rsZGUIRC10_GetBuffer rsSab, lZGUIRC10
Else
''    srvZGUIRC10_Init lZGUIRC10
    rsZGUIRC10_Read = "? inconnu"
End If
    
    Exit Function
Error_Handler:
    rsZGUIRC10_Read = Error

End Function

Public Function rsZGUIRC10_GetBuffer(rsAdo As ADODB.Recordset, rsZGUIRC10 As typeZGUIRC10)
On Error GoTo Error_Handler
rsZGUIRC10_GetBuffer = Null
rsZGUIRC10.GUIRC1ETA = rsAdo("GUIRC1ETA")
rsZGUIRC10.GUIRC1AGE = rsAdo("GUIRC1AGE")
rsZGUIRC10.GUIRC1SER = rsAdo("GUIRC1SER")
rsZGUIRC10.GUIRC1SSE = rsAdo("GUIRC1SSE")
rsZGUIRC10.GUIRC1OPE = rsAdo("GUIRC1OPE")
rsZGUIRC10.GUIRC1DOS = rsAdo("GUIRC1DOS")
rsZGUIRC10.GUIRC1DCR = rsAdo("GUIRC1DCR")
rsZGUIRC10.GUIRC1NAT = rsAdo("GUIRC1NAT")
rsZGUIRC10.GUIRC1DE1 = rsAdo("GUIRC1DE1")
rsZGUIRC10.GUIRC1CP2 = rsAdo("GUIRC1CP2")
rsZGUIRC10.GUIRC1MO2 = rsAdo("GUIRC1MO2")

Exit Function
Error_Handler:
rsZGUIRC10_GetBuffer = Error
End Function

