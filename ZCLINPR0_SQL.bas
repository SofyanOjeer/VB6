Attribute VB_Name = "sqlZCLINPR0"
Option Explicit

Public Function sqlZCLINPR0_Insert(newY As typeZCLINPR0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlZCLINPR0_Insert = Null

xSet = " ("
xValues = " values("

' Détecter les modifications
'===================================================================================

If newY.CLINPRETA <> 0 Then xSet = xSet & ",CLINPRETA": xValues = xValues & " ," & newY.CLINPRETA

X = Trim(newY.CLINPRCLI): If X <> "" Then xSet = xSet & ",CLINPRCLI": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLINPRTYP): If X <> "" Then xSet = xSet & ",CLINPRTYP": xValues = xValues & " ,'" & X & "'"
X = Replace(Trim(newY.CLINPRNUM), "'", "''"): If X <> "" Then xSet = xSet & ",CLINPRNUM": xValues = xValues & " ,'" & X & "'"

Mid$(xSet, 3, 1) = " "
Mid$(xValues, 10, 1) = " "

xSQL = "Insert into " & paramIBM_Library_SAB & ".ZCLINPR0" & xSet & ")" & xValues & ")"
Call FEU_ROUGE
Set rsSab_Update = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlZCLINPR0_Insert = "Erreur màj : " & newY.CLINPRCLI
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlZCLINPR0_Insert = Error
End Function
Public Function sqlZCLINPR0_Update(newY As typeZCLINPR0, oldY As typeZCLINPR0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlZCLINPR0_Update = Null

xWhere = " where CLINPRCLI = '" & oldY.CLINPRCLI & "'" _
         & " and CLINPRETA  = " & oldY.CLINPRETA
         
xSet = " set"

' Détecter les modifications
'===================================================================================
If newY.CLINPRETA <> oldY.CLINPRETA Then xSet = xSet & " , CLINPRETA = " & newY.CLINPRETA

If newY.CLINPRCLI <> oldY.CLINPRCLI Then xSet = xSet & " , CLINPRCLI = '" & newY.CLINPRCLI & "'"
If newY.CLINPRTYP <> oldY.CLINPRTYP Then xSet = xSet & " , CLINPRTYP = '" & newY.CLINPRTYP & "'"
If newY.CLINPRNUM <> oldY.CLINPRNUM Then xSet = xSet & " , CLINPRNUM = '" & Replace(Trim(newY.CLINPRNUM), "'", "''") & "'"

If xSet = " set" Then Exit Function

Mid$(xSet, 6, 1) = " "
xSQL = "update " & paramIBM_Library_SAB & ".ZCLINPR0" & xSet & xWhere
Call FEU_ROUGE
Set rsSab_Update = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlZCLINPR0_Update = "Erreur màj : " & newY.CLINPRCLI
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlZCLINPR0_Update = Error
End Function










