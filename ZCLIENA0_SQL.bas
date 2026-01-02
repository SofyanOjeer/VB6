Attribute VB_Name = "sqlZCLIENA0"
Option Explicit

Public Function sqlZCLIENA0_Insert(newY As typeZCLIENA0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlZCLIENA0_Insert = Null

xSet = " ("
xValues = " values("

' Détecter les modifications
'===================================================================================

If newY.CLIENAETB <> 0 Then xSet = xSet & ",CLIENAETB": xValues = xValues & " ," & newY.CLIENAETB
If newY.CLIENAAGE <> 0 Then xSet = xSet & ",CLIENAAGE": xValues = xValues & " ," & newY.CLIENAAGE
If newY.CLIENASRT <> 0 Then xSet = xSet & ",CLIENASRT": xValues = xValues & " ," & newY.CLIENASRT
If newY.CLIENADNA <> 0 Then xSet = xSet & ",CLIENADNA": xValues = xValues & " ," & newY.CLIENADNA
If newY.CLIENAATR <> 0 Then xSet = xSet & ",CLIENAATR": xValues = xValues & " ," & newY.CLIENAATR
If newY.CLIENABIL <> 0 Then xSet = xSet & ",CLIENABIL": xValues = xValues & " ," & newY.CLIENABIL
If newY.CLIENADAT <> 0 Then xSet = xSet & ",CLIENADAT": xValues = xValues & " ," & newY.CLIENADAT
If newY.CLIENAPAY <> 0 Then xSet = xSet & ",CLIENAPAY": xValues = xValues & " ," & newY.CLIENAPAY
If newY.CLIENABIM <> 0 Then xSet = xSet & ",CLIENABIM": xValues = xValues & " ," & newY.CLIENABIM
If newY.CLIENACRE <> 0 Then xSet = xSet & ",CLIENACRE": xValues = xValues & " ," & newY.CLIENACRE

X = Trim(newY.CLIENACLI): If X <> "" Then xSet = xSet & ",CLIENACLI": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENAETA): If X <> "" Then xSet = xSet & ",CLIENAETA": xValues = xValues & " ,'" & X & "'"
X = Replace(Trim(newY.CLIENARA1), "'", "''"): If X <> "" Then xSet = xSet & ",CLIENARA1": xValues = xValues & " ,'" & X & "'"
X = Replace(Trim(newY.CLIENARA2), "'", "''"): If X <> "" Then xSet = xSet & ",CLIENARA2": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENASIG): If X <> "" Then xSet = xSet & ",CLIENASIG": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENASRN): If X <> "" Then xSet = xSet & ",CLIENASRN": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENAREG): If X <> "" Then xSet = xSet & ",CLIENAREG": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENANAT): If X <> "" Then xSet = xSet & ",CLIENANAT": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENARSD): If X <> "" Then xSet = xSet & ",CLIENARSD": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENARES): If X <> "" Then xSet = xSet & ",CLIENARES": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENAECO): If X <> "" Then xSet = xSet & ",CLIENAECO": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENAACT): If X <> "" Then xSet = xSet & ",CLIENAACT": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENAPAI): If X <> "" Then xSet = xSet & ",CLIENAPAI": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENACRD): If X <> "" Then xSet = xSet & ",CLIENACRD": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENAADM): If X <> "" Then xSet = xSet & ",CLIENAADM": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENACAT): If X <> "" Then xSet = xSet & ",CLIENACAT": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENACOT): If X <> "" Then xSet = xSet & ",CLIENACOT": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENACHQ): If X <> "" Then xSet = xSet & ",CLIENACHQ": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENASAC): If X <> "" Then xSet = xSet & ",CLIENASAC": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENAGEO): If X <> "" Then xSet = xSet & ",CLIENAGEO": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENAENT): If X <> "" Then xSet = xSet & ",CLIENAENT": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENAMES): If X <> "" Then xSet = xSet & ",CLIENAMES": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENAFIL): If X <> "" Then xSet = xSet & ",CLIENAFIL": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENADOU): If X <> "" Then xSet = xSet & ",CLIENADOU": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENALI1): If X <> "" Then xSet = xSet & ",CLIENALI1": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENALI2): If X <> "" Then xSet = xSet & ",CLIENALI2": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENAEXT): If X <> "" Then xSet = xSet & ",CLIENAEXT": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENACOL): If X <> "" Then xSet = xSet & ",CLIENACOL": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENASEL): If X <> "" Then xSet = xSet & ",CLIENASEL": xValues = xValues & " ,'" & X & "'"
X = Trim(newY.CLIENAPCS): If X <> "" Then xSet = xSet & ",CLIENAPCS": xValues = xValues & " ,'" & X & "'"

Mid$(xSet, 3, 1) = " "
Mid$(xValues, 10, 1) = " "
Call FEU_ROUGE
xSQL = "Insert into " & paramIBM_Library_SAB & ".ZCLIENA0" & xSet & ")" & xValues & ")"

Set rsSab_Update = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlZCLIENA0_Insert = "Erreur màj : " & newY.CLIENACLI
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlZCLIENA0_Insert = Error
End Function
Public Function sqlZCLIENA0_Update(newY As typeZCLIENA0, oldY As typeZCLIENA0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlZCLIENA0_Update = Null

xWhere = " where CLIENACLI = '" & oldY.CLIENACLI & "'" _
         & " and CLIENAETB  = " & oldY.CLIENAETB
         
xSet = " set"

' Détecter les modifications
'===================================================================================
If newY.CLIENAETB <> oldY.CLIENAETB Then xSet = xSet & " , CLIENAETB = " & newY.CLIENAETB
If newY.CLIENAAGE <> oldY.CLIENAAGE Then xSet = xSet & " , CLIENAAGE = " & newY.CLIENAAGE
If newY.CLIENASRT <> oldY.CLIENASRT Then xSet = xSet & " , CLIENASRT = " & newY.CLIENASRT
If newY.CLIENADNA <> oldY.CLIENADNA Then xSet = xSet & " , CLIENADNA = " & newY.CLIENADNA
If newY.CLIENAATR <> oldY.CLIENAATR Then xSet = xSet & " , CLIENAATR = " & newY.CLIENAATR
If newY.CLIENABIL <> oldY.CLIENABIL Then xSet = xSet & " , CLIENABIL = " & newY.CLIENABIL
If newY.CLIENADAT <> oldY.CLIENADAT Then xSet = xSet & " , CLIENADAT = " & newY.CLIENADAT
If newY.CLIENAPAY <> oldY.CLIENAPAY Then xSet = xSet & " , CLIENAPAY = " & newY.CLIENAPAY
If newY.CLIENABIM <> oldY.CLIENABIM Then xSet = xSet & " , CLIENABIM = " & newY.CLIENABIM
If newY.CLIENACRE <> oldY.CLIENACRE Then xSet = xSet & " , CLIENACRE = " & newY.CLIENACRE

If newY.CLIENACLI <> oldY.CLIENACLI Then xSet = xSet & " , CLIENACLI = '" & newY.CLIENACLI & "'"
If newY.CLIENAETA <> oldY.CLIENAETA Then xSet = xSet & " , CLIENAETA = '" & newY.CLIENAETA & "'"
If newY.CLIENARA1 <> oldY.CLIENARA1 Then xSet = xSet & " , CLIENARA1 = '" & Replace(Trim(newY.CLIENARA1), "'", "''") & "'"
If newY.CLIENARA2 <> oldY.CLIENARA2 Then xSet = xSet & " , CLIENARA2 = '" & Replace(Trim(newY.CLIENARA2), "'", "''") & "'"
If newY.CLIENASIG <> oldY.CLIENASIG Then xSet = xSet & " , CLIENASIG = '" & newY.CLIENASIG & "'"
If newY.CLIENASRN <> oldY.CLIENASRN Then xSet = xSet & " , CLIENASRN = '" & newY.CLIENASRN & "'"
If newY.CLIENAREG <> oldY.CLIENAREG Then xSet = xSet & " , CLIENAREG = '" & newY.CLIENAREG & "'"
If newY.CLIENANAT <> oldY.CLIENANAT Then xSet = xSet & " , CLIENANAT = '" & newY.CLIENANAT & "'"
If newY.CLIENARSD <> oldY.CLIENARSD Then xSet = xSet & " , CLIENARSD = '" & newY.CLIENARSD & "'"
If newY.CLIENARES <> oldY.CLIENARES Then xSet = xSet & " , CLIENARES = '" & newY.CLIENARES & "'"
If newY.CLIENAECO <> oldY.CLIENAECO Then xSet = xSet & " , CLIENAECO = '" & newY.CLIENAECO & "'"
If newY.CLIENAACT <> oldY.CLIENAACT Then xSet = xSet & " , CLIENAACT = '" & newY.CLIENAACT & "'"
If newY.CLIENAPAI <> oldY.CLIENAPAI Then xSet = xSet & " , CLIENAPAI = '" & newY.CLIENAPAI & "'"
If newY.CLIENACRD <> oldY.CLIENACRD Then xSet = xSet & " , CLIENACRD = '" & newY.CLIENACRD & "'"
If newY.CLIENAADM <> oldY.CLIENAADM Then xSet = xSet & " , CLIENAADM = '" & newY.CLIENAADM & "'"
If newY.CLIENACAT <> oldY.CLIENACAT Then xSet = xSet & " , CLIENACAT = '" & newY.CLIENACAT & "'"
If newY.CLIENACOT <> oldY.CLIENACOT Then xSet = xSet & " , CLIENACOT = '" & newY.CLIENACOT & "'"
If newY.CLIENACHQ <> oldY.CLIENACHQ Then xSet = xSet & " , CLIENACHQ = '" & newY.CLIENACHQ & "'"
If newY.CLIENASAC <> oldY.CLIENASAC Then xSet = xSet & " , CLIENASAC = '" & newY.CLIENASAC & "'"
If newY.CLIENAGEO <> oldY.CLIENAGEO Then xSet = xSet & " , CLIENAGEO = '" & newY.CLIENAGEO & "'"
If newY.CLIENAENT <> oldY.CLIENAENT Then xSet = xSet & " , CLIENAENT = '" & newY.CLIENAENT & "'"
If newY.CLIENAMES <> oldY.CLIENAMES Then xSet = xSet & " , CLIENAMES = '" & newY.CLIENAMES & "'"
If newY.CLIENAFIL <> oldY.CLIENAFIL Then xSet = xSet & " , CLIENAFIL = '" & newY.CLIENAFIL & "'"
If newY.CLIENADOU <> oldY.CLIENADOU Then xSet = xSet & " , CLIENADOU = '" & newY.CLIENADOU & "'"
If newY.CLIENALI1 <> oldY.CLIENALI1 Then xSet = xSet & " , CLIENALI1 = '" & newY.CLIENALI1 & "'"
If newY.CLIENALI2 <> oldY.CLIENALI2 Then xSet = xSet & " , CLIENALI2 = '" & newY.CLIENALI2 & "'"
If newY.CLIENAEXT <> oldY.CLIENAEXT Then xSet = xSet & " , CLIENAEXT = '" & newY.CLIENAEXT & "'"
If newY.CLIENACOL <> oldY.CLIENACOL Then xSet = xSet & " , CLIENACOL = '" & newY.CLIENACOL & "'"
If newY.CLIENASEL <> oldY.CLIENASEL Then xSet = xSet & " , CLIENASEL = '" & newY.CLIENASEL & "'"
If newY.CLIENAPCS <> oldY.CLIENAPCS Then xSet = xSet & " , CLIENAPCS = '" & newY.CLIENAPCS & "'"

If xSet = " set" Then Exit Function

Mid$(xSet, 6, 1) = " "
xSQL = "update " & paramIBM_Library_SAB & ".ZCLIENA0" & xSet & xWhere
Call FEU_ROUGE
Set rsSab_Update = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlZCLIENA0_Update = "Erreur màj : " & newY.CLIENACLI
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlZCLIENA0_Update = Error
End Function









