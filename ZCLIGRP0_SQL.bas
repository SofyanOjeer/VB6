Attribute VB_Name = "sqlZCLIGRP0"
Option Explicit

Public Function sqlZCLIGRP0_Insert(newY As typeZCLIGRP0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlZCLIGRP0_Insert = Null

xSet = " ("
xValues = " values("

' Détecter les modifications
'===================================================================================

xSet = xSet & ",CLIGRPETB": xValues = xValues & " ," & newY.CLIGRPETB
xSet = xSet & ",CLIGRPCLI": xValues = xValues & " ,'" & newY.CLIGRPCLI & "'"
xSet = xSet & ",CLIGRPREG": xValues = xValues & " ,'" & Trim(newY.CLIGRPREG) & "'"
xSet = xSet & ",CLIGRPREl": xValues = xValues & " ,'" & newY.CLIGRPREL & "'"

Mid$(xSet, 3, 1) = " "
Mid$(xValues, 10, 1) = " "
Call FEU_ROUGE
xSQL = "Insert into " & paramIBM_Library_SAB & ".ZCLIGRP0" & xSet & ")" & xValues & ")"

Set rsSab_Update = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlZCLIGRP0_Insert = "Erreur màj : " & newY.CLIGRPREG
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlZCLIGRP0_Insert = Error
End Function
Public Function sqlZCLIGRP0_Delete(newY As typeZCLIGRP0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String

On Error GoTo Error_Handler
sqlZCLIGRP0_Delete = Null


xSQL = "delete from " & paramIBM_Library_SAB & ".ZCLIGRP0" _
    & " where CLIGRPCLI = '" & newY.CLIGRPCLI & "'" _
    & " and CLIGRPREG = '" & newY.CLIGRPREG & "'" _
    & " and CLIGRPETB = " & newY.CLIGRPETB
Call FEU_ROUGE
Set rsSab_Update = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlZCLIGRP0_Delete = "Erreur màj : " & xSQL
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlZCLIGRP0_Delete = Error
End Function

Public Function sqlZCLIGRP0_Update(newY As typeZCLIGRP0, oldY As typeZCLIGRP0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlZCLIGRP0_Update = Null

xWhere = " where CLIGRPREG = '" & oldY.CLIGRPREG & "'" _
         & " and CLIGRPCLI  = " & oldY.CLIGRPCLI & "'" _
         & " and CLIGRPETB  = " & oldY.CLIGRPETB
         
xSet = " set"

' Détecter les modifications
'===================================================================================
If newY.CLIGRPETB <> oldY.CLIGRPETB Then xSet = xSet & " , CLIGRPETB = " & newY.CLIGRPETB
If newY.CLIGRPCLI <> oldY.CLIGRPCLI Then xSet = xSet & " , CLIGRPCLI = '" & newY.CLIGRPCLI & "'"
If newY.CLIGRPREG <> oldY.CLIGRPREG Then xSet = xSet & " , CLIGRPREG = '" & newY.CLIGRPREG & "'"
If newY.CLIGRPREL <> oldY.CLIGRPREL Then xSet = xSet & " , CLIGRPREl = '" & newY.CLIGRPREL & "'"

If xSet = " set" Then Exit Function

Mid$(xSet, 6, 1) = " "
xSQL = "update " & paramIBM_Library_SAB & ".ZCLIGRP0" & xSet & xWhere
Call FEU_ROUGE
Set rsSab_Update = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlZCLIGRP0_Update = "Erreur màj : " & newY.CLIGRPREG
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlZCLIGRP0_Update = Error
End Function





