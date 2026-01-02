Attribute VB_Name = "sqlZCLIENB0"
Option Explicit

Public Function sqlZCLIENB0_Insert(newY As typeZCLIENB0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlZCLIENB0_Insert = Null

xSet = " ("
xValues = " values("

' Détecter les modifications
'===================================================================================

If newY.CLIENBETB <> 0 Then xSet = xSet & ",CLIENBETB": xValues = xValues & " ," & newY.CLIENBETB
If newY.CLIENBCRT <> 0 Then xSet = xSet & ",CLIENBCRT": xValues = xValues & " ," & newY.CLIENBCRT
If newY.CLIENBBIL <> 0 Then xSet = xSet & ",CLIENBBIL": xValues = xValues & " ," & newY.CLIENBBIL
If newY.CLIENBEFC <> 0 Then xSet = xSet & ",CLIENBEFC": xValues = xValues & " ," & newY.CLIENBEFC
If newY.CLIENBCH1 <> 0 Then xSet = xSet & ",CLIENBCH1": xValues = xValues & " ," & newY.CLIENBCH1
If newY.CLIENBCH2 <> 0 Then xSet = xSet & ",CLIENBCH2": xValues = xValues & " ," & newY.CLIENBCH2
If newY.CLIENBCH3 <> 0 Then xSet = xSet & ",CLIENBCH3": xValues = xValues & " ," & newY.CLIENBCH3
If newY.CLIENBCP1 <> 0 Then xSet = xSet & ",CLIENBCP1": xValues = xValues & " ," & newY.CLIENBCP1
If newY.CLIENBMD1 <> 0 Then xSet = xSet & ",CLIENBMD1": xValues = xValues & " ," & newY.CLIENBMD1
If newY.CLIENBMUT <> 0 Then xSet = xSet & ",CLIENBMUT": xValues = xValues & " ," & newY.CLIENBMUT
If newY.CLIENBDEC <> 0 Then xSet = xSet & ",CLIENBDEC": xValues = xValues & " ," & newY.CLIENBDEC

X = Trim(newY.CLIENBCLI): If X <> "" Then xSet = xSet & ",CLIENBCLI": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENBAF1): If X <> "" Then xSet = xSet & ",CLIENBAF1": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENBAF2): If X <> "" Then xSet = xSet & ",CLIENBAF2": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENBAF3): If X <> "" Then xSet = xSet & ",CLIENBAF3": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENBNAS): If X <> "" Then xSet = xSet & ",CLIENBNAS": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENBINS): If X <> "" Then xSet = xSet & ",CLIENBINS": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENBCOM): If X <> "" Then xSet = xSet & ",CLIENBCOM": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENBLIE): If X <> "" Then xSet = xSet & ",CLIENBLIE": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENBTER): If X <> "" Then xSet = xSet & ",CLIENBTER": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENBPER): If X <> "" Then xSet = xSet & ",CLIENBPER": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENBMAR): If X <> "" Then xSet = xSet & ",CLIENBMAR": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENBJUR): If X <> "" Then xSet = xSet & ",CLIENBJUR": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENBCAP): If X <> "" Then xSet = xSet & ",CLIENBCAP": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENBBAN): If X <> "" Then xSet = xSet & ",CLIENBBAN": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENBLIB): If X <> "" Then xSet = xSet & ",CLIENBLIB": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENBDED): If X <> "" Then xSet = xSet & ",CLIENBDED": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENBSER): If X <> "" Then xSet = xSet & ",CLIENBSER": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENBSEP): If X <> "" Then xSet = xSet & ",CLIENBSEP": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENBCTL): If X <> "" Then xSet = xSet & ",CLIENBCTL": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENBCIN): If X <> "" Then xSet = xSet & ",CLIENBCIN": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENBTOP): If X <> "" Then xSet = xSet & ",CLIENBTOP": xValues = xValues & " ,'" & X & "'"

Mid$(xSet, 3, 1) = " "
Mid$(xValues, 10, 1) = " "
Call FEU_ROUGE
xSQL = "Insert into " & paramIBM_Library_SAB & ".ZCLIENB0" & xSet & ")" & xValues & ")"

Set rsSab_Update = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlZCLIENB0_Insert = "Erreur màj : " & newY.CLIENBCLI
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlZCLIENB0_Insert = Error
End Function
Public Function sqlZCLIENB0_Update(newY As typeZCLIENB0, oldY As typeZCLIENB0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlZCLIENB0_Update = Null

xWhere = " where CLIENBCLI = '" & oldY.CLIENBCLI & "'" _
         & " and CLIENBETB  = " & oldY.CLIENBETB
         
xSet = " set"

' Détecter les modifications
'===================================================================================
If newY.CLIENBETB <> oldY.CLIENBETB Then xSet = xSet & " , CLIENBETB = " & newY.CLIENBETB
If newY.CLIENBCRT <> oldY.CLIENBCRT Then xSet = xSet & " , CLIENBCRT = " & newY.CLIENBCRT
If newY.CLIENBBIL <> oldY.CLIENBBIL Then xSet = xSet & " , CLIENBBIL = " & newY.CLIENBBIL
If newY.CLIENBEFC <> oldY.CLIENBEFC Then xSet = xSet & " , CLIENBEFC = " & newY.CLIENBEFC
If newY.CLIENBCH1 <> oldY.CLIENBCH1 Then xSet = xSet & " , CLIENBCH1 = " & newY.CLIENBCH1
If newY.CLIENBCH2 <> oldY.CLIENBCH2 Then xSet = xSet & " , CLIENBCH2 = " & newY.CLIENBCH2
If newY.CLIENBCH3 <> oldY.CLIENBCH3 Then xSet = xSet & " , CLIENBCH3 = " & newY.CLIENBCH3
If newY.CLIENBCP1 <> oldY.CLIENBCP1 Then xSet = xSet & " , CLIENBCP1 = " & newY.CLIENBCP1
If newY.CLIENBMD1 <> oldY.CLIENBMD1 Then xSet = xSet & " , CLIENBMD1 = " & newY.CLIENBMD1
If newY.CLIENBMUT <> oldY.CLIENBMUT Then xSet = xSet & " , CLIENBMUT = " & newY.CLIENBMUT
If newY.CLIENBDEC <> oldY.CLIENBDEC Then xSet = xSet & " , CLIENBDEC = " & newY.CLIENBDEC

If newY.CLIENBCLI <> oldY.CLIENBCLI Then xSet = xSet & " , CLIENBCLI = '" & newY.CLIENBCLI & "'"
If newY.CLIENBAF1 <> oldY.CLIENBAF1 Then xSet = xSet & " , CLIENBAF1 = '" & newY.CLIENBAF1 & "'"
If newY.CLIENBAF2 <> oldY.CLIENBAF2 Then xSet = xSet & " , CLIENBAF2 = '" & newY.CLIENBAF2 & "'"
If newY.CLIENBAF3 <> oldY.CLIENBAF3 Then xSet = xSet & " , CLIENBAF3 = '" & newY.CLIENBAF3 & "'"
If newY.CLIENBNAS <> oldY.CLIENBNAS Then xSet = xSet & " , CLIENBNAS = '" & newY.CLIENBNAS & "'"
If newY.CLIENBINS <> oldY.CLIENBINS Then xSet = xSet & " , CLIENBINS = '" & newY.CLIENBINS & "'"
If newY.CLIENBCOM <> oldY.CLIENBCOM Then xSet = xSet & " , CLIENBCOM = '" & newY.CLIENBCOM & "'"
If newY.CLIENBLIE <> oldY.CLIENBLIE Then xSet = xSet & " , CLIENBLIE = '" & newY.CLIENBLIE & "'"
If newY.CLIENBTER <> oldY.CLIENBTER Then xSet = xSet & " , CLIENBTER = '" & newY.CLIENBTER & "'"
If newY.CLIENBPER <> oldY.CLIENBPER Then xSet = xSet & " , CLIENBPER = '" & newY.CLIENBPER & "'"
If newY.CLIENBMAR <> oldY.CLIENBMAR Then xSet = xSet & " , CLIENBMAR = '" & newY.CLIENBMAR & "'"
If newY.CLIENBJUR <> oldY.CLIENBJUR Then xSet = xSet & " , CLIENBJUR = '" & newY.CLIENBJUR & "'"
If newY.CLIENBCAP <> oldY.CLIENBCAP Then xSet = xSet & " , CLIENBCAP = '" & newY.CLIENBCAP & "'"
If newY.CLIENBBAN <> oldY.CLIENBBAN Then xSet = xSet & " , CLIENBBAN = '" & newY.CLIENBBAN & "'"
If newY.CLIENBLIB <> oldY.CLIENBLIB Then xSet = xSet & " , CLIENBLIB = '" & newY.CLIENBLIB & "'"
If newY.CLIENBDED <> oldY.CLIENBDED Then xSet = xSet & " , CLIENBDED = '" & newY.CLIENBDED & "'"
If newY.CLIENBSER <> oldY.CLIENBSER Then xSet = xSet & " , CLIENBSER = '" & newY.CLIENBSER & "'"
If newY.CLIENBSEP <> oldY.CLIENBSEP Then xSet = xSet & " , CLIENBSEP = '" & newY.CLIENBSEP & "'"
If newY.CLIENBCTL <> oldY.CLIENBCTL Then xSet = xSet & " , CLIENBCTL = '" & newY.CLIENBCTL & "'"
If newY.CLIENBCIN <> oldY.CLIENBCIN Then xSet = xSet & " , CLIENBCIN = '" & newY.CLIENBCIN & "'"
If newY.CLIENBTOP <> oldY.CLIENBTOP Then xSet = xSet & " , CLIENBTOP = '" & newY.CLIENBTOP & "'"

If xSet = " set" Then Exit Function

Mid$(xSet, 6, 1) = " "
xSQL = "update " & paramIBM_Library_SAB & ".ZCLIENB0" & xSet & xWhere
Call FEU_ROUGE
Set rsSab_Update = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlZCLIENB0_Update = "Erreur màj : " & newY.CLIENBCLI
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlZCLIENB0_Update = Error
End Function











