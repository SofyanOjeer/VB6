Attribute VB_Name = "sqlYBIARELV"
Option Explicit

Public Function sqlYBIARELV_Insert(newY As typeYBIARELV)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYBIARELV_Insert = Null

xSet = " ("
xValues = " values("

' Détecter les modifications
'===================================================================================

xSet = xSet & ",BIARELCOM": xValues = xValues & " ,'" & Trim(newY.BIARELCOM) & "'"
xSet = xSet & ",BIARELREL": xValues = xValues & " ,'" & newY.BIARELREL & "'"
xSet = xSet & ",BIARELID": xValues = xValues & " ," & newY.BIARELID
xSet = xSet & ",BIARELNUM": xValues = xValues & " ," & newY.BIARELNUM
xSet = xSet & ",BIARELSD0": xValues = xValues & " ," & cur_P(newY.BIARELSD0)
xSet = xSet & ",BIARELD0": xValues = xValues & " ," & newY.BIARELD0
xSet = xSet & ",BIARELSD1": xValues = xValues & " ," & cur_P(newY.BIARELSD1)
xSet = xSet & ",BIARELD1": xValues = xValues & " ," & newY.BIARELD1
xSet = xSet & ",BIAOLDCOM": xValues = xValues & " ,'" & newY.BIAOLDCOM & "'"
xSet = xSet & ",BIAOLDDEV": xValues = xValues & " ,'" & newY.BIAOLDDEV & "'"

Mid$(xSet, 3, 1) = " "
Mid$(xValues, 10, 1) = " "
Call FEU_ROUGE
xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YBIARELV" & xSet & ")" & xValues & ")"

Set rsSab_Update = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYBIARELV_Insert = "Erreur màj : " & newY.BIARELD0
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYBIARELV_Insert = Error
End Function
Public Function sqlYBIARELV_Update(newY As typeYBIARELV, oldY As typeYBIARELV)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlYBIARELV_Update = Null

xWhere = " where BIARELCOM = '" & oldY.BIARELCOM & "'" _
         & " and BIARELREL = '" & oldY.BIARELREL & "'" _
         & " and BIARELID  = " & oldY.BIARELID
         
xSet = xSet & " set"

' Détecter les modifications
'===================================================================================
If newY.BIARELCOM <> oldY.BIARELCOM Then xSet = xSet & " , BIARELCOM = '" & Trim(newY.BIARELCOM) & "'"
If newY.BIARELREL <> oldY.BIARELREL Then xSet = xSet & " , BIARELREL = '" & newY.BIARELREL & "'"
If newY.BIARELID <> oldY.BIARELID Then xSet = xSet & " , BIARELID = " & newY.BIARELID
If newY.BIARELNUM <> oldY.BIARELNUM Then xSet = xSet & " , BIARELNUM = " & newY.BIARELNUM
If newY.BIARELSD0 <> oldY.BIARELSD0 Then xSet = xSet & " , BIARELSD0 = " & cur_P(newY.BIARELSD0)
If newY.BIARELD0 <> oldY.BIARELD0 Then xSet = xSet & " , BIARELD0 = " & newY.BIARELD0
If newY.BIARELSD1 <> oldY.BIARELSD1 Then xSet = xSet & " , BIARELSD1 = " & cur_P(newY.BIARELSD1)
If newY.BIARELD1 <> oldY.BIARELD1 Then xSet = xSet & " , BIARELD1 = " & newY.BIARELD1
If newY.BIAOLDCOM <> oldY.BIAOLDCOM Then xSet = xSet & " , BIAOLDCOM = '" & newY.BIAOLDCOM & "'"
If newY.BIAOLDDEV <> oldY.BIAOLDDEV Then xSet = xSet & " , BIAOLDDEV = '" & newY.BIAOLDDEV & "'"


Mid$(xSet, 6, 1) = " "
xSQL = "update " & paramIBM_Library_SABSPE & ".YBIARELV" & xSet & xWhere
Call FEU_ROUGE
Set rsSab_Update = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYBIARELV_Update = "Erreur màj : " & newY.BIARELD0
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYBIARELV_Update = Error
End Function


