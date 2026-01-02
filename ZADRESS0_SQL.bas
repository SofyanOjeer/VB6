Attribute VB_Name = "sqlZADRESS0"
Option Explicit

Public Function sqlZADRESS0_Insert(newY As typeZADRESS0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlZADRESS0_Insert = Null

xSet = " ("
xValues = " values("

' Détecter les modifications
'===================================================================================

xSet = xSet & ",ADRESSETA": xValues = xValues & " ," & newY.ADRESSETA
xSet = xSet & ",ADRESSTYP": xValues = xValues & " ,'" & newY.ADRESSTYP & "'"
xSet = xSet & ",ADRESSPLA": xValues = xValues & " ," & newY.ADRESSPLA
xSet = xSet & ",ADRESSNUM": xValues = xValues & " ,'" & newY.ADRESSNUM & "'"
xSet = xSet & ",ADRESSCOA": xValues = xValues & " ,'" & newY.ADRESSCOA & "'"
If paramIBM_AS400_ID = "I5A7" Then
    xSet = xSet & ",ADRESSRA1": xValues = xValues & " ,'" & Replace(Trim(newY.ADRESSRA1), "'", "''") & "'"
Else
    X = Replace(Trim(newY.ADRESSRA1), "'", "''")
    xSet = xSet & ",ADRESSRA11": xValues = xValues & " ,'" & Mid$(X, 1, 10) & "'"
    xSet = xSet & ",ADRESSRA12": xValues = xValues & " ,'" & Mid$(X, 11, 15) & "'"
    xSet = xSet & ",ADRESSRA13": xValues = xValues & " ,'" & Mid$(X, 26, 7) & "'"
End If

xSet = xSet & ",ADRESSRA2": xValues = xValues & " ,'" & Replace(Trim(newY.ADRESSRA2), "'", "''") & "'"
xSet = xSet & ",ADRESSAD1": xValues = xValues & " ,'" & Replace(Trim(newY.ADRESSAD1), "'", "''") & "'"
xSet = xSet & ",ADRESSAD2": xValues = xValues & " ,'" & Replace(Trim(newY.ADRESSAD2), "'", "''") & "'"
xSet = xSet & ",ADRESSAD3": xValues = xValues & " ,'" & Replace(Trim(newY.ADRESSAD3), "'", "''") & "'"
xSet = xSet & ",ADRESSCOP": xValues = xValues & " ,'" & Replace(Trim(newY.ADRESSCOP), "'", "''") & "'"
xSet = xSet & ",ADRESSVIL": xValues = xValues & " ,'" & Replace(Trim(newY.ADRESSVIL), "'", "''") & "'"
xSet = xSet & ",ADRESSPAY": xValues = xValues & " ,'" & Replace(Trim(newY.ADRESSPAY), "'", "''") & "'"
Mid$(xSet, 3, 1) = " "
Mid$(xValues, 10, 1) = " "
Call FEU_ROUGE
xSQL = "Insert into " & paramIBM_Library_SAB & ".ZADRESS0" & xSet & ")" & xValues & ")"

Set rsSab_Update = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlZADRESS0_Insert = "Erreur màj : " & newY.ADRESSPLA
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlZADRESS0_Insert = Error
End Function
Public Function sqlZADRESS0_Update(newY As typeZADRESS0, oldY As typeZADRESS0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlZADRESS0_Update = Null

xWhere = " where ADRESSNUM = '" & oldY.ADRESSNUM & "'" _
         & " and ADRESSTYP  = '" & oldY.ADRESSTYP & "'" _
         & " and ADRESSPLA  = " & oldY.ADRESSPLA _
         & " and ADRESSCOA  = '" & oldY.ADRESSCOA & "'" _
         & " and ADRESSETA  = " & oldY.ADRESSETA
         
xSet = " set"

' Détecter les modifications
'===================================================================================
If newY.ADRESSETA <> oldY.ADRESSETA Then xSet = xSet & " , ADRESSETA = " & newY.ADRESSETA
If newY.ADRESSTYP <> oldY.ADRESSTYP Then xSet = xSet & " , ADRESSTYP = '" & newY.ADRESSTYP & "'"
If newY.ADRESSPLA <> oldY.ADRESSPLA Then xSet = xSet & " , ADRESSPLA = " & newY.ADRESSPLA
If newY.ADRESSNUM <> oldY.ADRESSNUM Then xSet = xSet & " , ADRESSNUM = '" & newY.ADRESSNUM & "'"
If newY.ADRESSCOA <> oldY.ADRESSCOA Then xSet = xSet & " , ADRESSCOA = '" & newY.ADRESSCOA & "'"
If newY.ADRESSRA1 <> oldY.ADRESSRA1 Then
    If paramIBM_AS400_ID = "I5A7" Then
        xSet = xSet & " , ADRESSRA1 = '" & Replace(Trim(newY.ADRESSRA1), "'", "''") & "'"
    Else
        X = Replace(Trim(newY.ADRESSRA1), "'", "''")
        xSet = xSet & " , ADRESSRA11 = '" & Mid$(X, 1, 10) & "'"
        xSet = xSet & " , ADRESSRA12 = '" & Mid$(X, 11, 15) & "'"
        xSet = xSet & " , ADRESSRA13 = '" & Mid$(X, 26, 7) & "'"
    End If
End If
If newY.ADRESSRA2 <> oldY.ADRESSRA2 Then xSet = xSet & " , ADRESSRA2 = '" & Replace(Trim(newY.ADRESSRA2), "'", "''") & "'"
If newY.ADRESSAD1 <> oldY.ADRESSAD1 Then xSet = xSet & " , ADRESSAD1 = '" & Replace(Trim(newY.ADRESSAD1), "'", "''") & "'"
If newY.ADRESSAD2 <> oldY.ADRESSAD2 Then xSet = xSet & " , ADRESSAD2 = '" & Replace(Trim(newY.ADRESSAD2), "'", "''") & "'"
If newY.ADRESSAD3 <> oldY.ADRESSAD3 Then xSet = xSet & " , ADRESSAD3 = '" & Replace(Trim(newY.ADRESSAD3), "'", "''") & "'"
If newY.ADRESSCOP <> oldY.ADRESSCOP Then xSet = xSet & " , ADRESSCOP = '" & Replace(Trim(newY.ADRESSCOP), "'", "''") & "'"
If newY.ADRESSVIL <> oldY.ADRESSVIL Then xSet = xSet & " , ADRESSVIL = '" & Replace(Trim(newY.ADRESSVIL), "'", "''") & "'"
If newY.ADRESSPAY <> oldY.ADRESSPAY Then xSet = xSet & " , ADRESSPAY = '" & Replace(Trim(newY.ADRESSPAY), "'", "''") & "'"

If xSet = " set" Then Exit Function

Mid$(xSet, 6, 1) = " "
xSQL = "update " & paramIBM_Library_SAB & ".ZADRESS0" & xSet & xWhere
Call FEU_ROUGE
Set rsSab_Update = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlZADRESS0_Update = "Erreur màj : " & newY.ADRESSPLA
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlZADRESS0_Update = Error
End Function







