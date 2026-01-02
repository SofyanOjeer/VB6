Attribute VB_Name = "sqlZCOMREF0"
Option Explicit

Public Function sqlZCOMREF0_Insert(newY As typeZCOMREF0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlZCOMREF0_Insert = Null

xSet = " ("
xValues = " values("

' Détecter les modifications
'===================================================================================

xSet = xSet & ",COMREFETA": xValues = xValues & " ," & newY.COMREFETA
xSet = xSet & ",COMREFPLA": xValues = xValues & " ," & newY.COMREFPLA
xSet = xSet & ",COMREFCOM": xValues = xValues & " ,'" & Trim(newY.COMREFCOM) & "'"
xSet = xSet & ",COMREFCOR": xValues = xValues & " ,'" & newY.COMREFCOR & "'"
xSet = xSet & ",COMREFREF": xValues = xValues & " ,'" & Trim(newY.COMREFREF) & "'"

Mid$(xSet, 3, 1) = " "
Mid$(xValues, 10, 1) = " "

Call FEU_ROUGE
xSQL = "Insert into " & paramIBM_Library_SAB & ".ZCOMREF0" & xSet & ")" & xValues & ")"

Set rsSab_Update = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlZCOMREF0_Insert = "Erreur màj : " & newY.COMREFCOM
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlZCOMREF0_Insert = Error
End Function
Public Function sqlZCOMREF0_Update(newY As typeZCOMREF0, oldY As typeZCOMREF0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlZCOMREF0_Update = Null

xWhere = " where COMREFCOM = '" & oldY.COMREFCOM & "'" _
         & " and COMREFcor = '" & oldY.COMREFCOR & "'" _
         & " and COMREFREF = '" & oldY.COMREFREF & "'" _
         & " and COMREFpla  = " & oldY.COMREFPLA & "'" _
         & " and COMREFETA  = " & oldY.COMREFETA
         
xSet = xSet & " set"

' Détecter les modifications
'===================================================================================
If newY.COMREFETA <> oldY.COMREFETA Then xSet = xSet & " , COMREFETA = " & newY.COMREFETA
If newY.COMREFPLA <> oldY.COMREFPLA Then xSet = xSet & " , COMREFPLA = " & newY.COMREFPLA
If newY.COMREFCOM <> oldY.COMREFCOM Then xSet = xSet & " , COMREFCOM = '" & newY.COMREFCOM & "'"
If newY.COMREFCOR <> oldY.COMREFCOR Then xSet = xSet & " , COMREFCOR = '" & newY.COMREFCOR & "'"
If newY.COMREFREF <> oldY.COMREFREF Then xSet = xSet & " , COMREFREF = '" & newY.COMREFREF & "'"


Mid$(xSet, 6, 1) = " "
xSQL = "update " & paramIBM_Library_SAB & ".ZCOMREF0" & xSet & xWhere
Call FEU_ROUGE
Set rsSab_Update = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlZCOMREF0_Update = "Erreur màj : " & newY.COMREFCOM
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlZCOMREF0_Update = Error
End Function




