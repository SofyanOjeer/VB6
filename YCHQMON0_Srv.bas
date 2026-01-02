Attribute VB_Name = "srvYCHQMON0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsADO As ADODB.Recordset
 
Type typeYCHQMON0
 
      CHQRC1ETA     As Long
      CHQRC1AGE     As Long
      CHQRC1SER     As String * 2
      CHQRC1SSE     As String * 2
      CHQRC1OPE     As String * 3
      CHQRC1DOS     As Long
      CHQRC1DCR     As Long
      CHQDATE       As Long
      CHQCOMPTE     As String * 20
      CHQCREM       As String * 8
      CHQDEVISE     As String * 3
      CHQMONTANT    As Currency
      CHQNB         As Long
      CHQMONSTA     As String * 1   '
      CHQMONUPDS    As Long

End Type
Public xYCHQMON0 As typeYCHQMON0
Public Function sqlYCHQMON0_Insert(newY As typeYCHQMON0, cnADO As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYCHQMON0_Insert = Null

xSet = " (CHQRC1ETA"
xValues = " values(" & newY.CHQRC1ETA

' Détecter les modifications
'===================================================================================
If newY.CHQRC1AGE <> 0 Then xSet = xSet & ",CHQRC1AGE": xValues = xValues & " ," & newY.CHQRC1AGE
If newY.CHQRC1SER <> "" Then xSet = xSet & ",CHQRC1SER": xValues = xValues & " ,'" & newY.CHQRC1SER & "'"
If newY.CHQRC1SSE <> "" Then xSet = xSet & ",CHQRC1SSE": xValues = xValues & " ,'" & newY.CHQRC1SSE & "'"
If newY.CHQRC1OPE <> "" Then xSet = xSet & ",CHQRC1OPE": xValues = xValues & " ,'" & newY.CHQRC1OPE & "'"
If newY.CHQRC1DOS <> 0 Then xSet = xSet & ",CHQRC1DOS": xValues = xValues & " ," & newY.CHQRC1DOS
If newY.CHQRC1DCR <> 0 Then xSet = xSet & ",CHQRC1DCR": xValues = xValues & " ," & newY.CHQRC1DCR
If newY.CHQDATE <> 0 Then xSet = xSet & ",CHQDATE": xValues = xValues & " ," & newY.CHQDATE
If newY.CHQCOMPTE <> "" Then xSet = xSet & ",CHQCOMPTE": xValues = xValues & " ,'" & Trim(newY.CHQCOMPTE) & "'"
If newY.CHQCREM <> "" Then xSet = xSet & ",CHQCREM": xValues = xValues & " ,'" & newY.CHQCREM & "'"
If newY.CHQDEVISE <> "" Then xSet = xSet & ",CHQDEVISE": xValues = xValues & " ,'" & newY.CHQDEVISE & "'"
If newY.CHQMONTANT <> 0 Then xSet = xSet & ",CHQMONTANT": xValues = xValues & " ," & cur_P(newY.CHQMONTANT)
If newY.CHQNB <> 0 Then xSet = xSet & ",CHQNB": xValues = xValues & " ," & newY.CHQNB
If Trim(newY.CHQMONSTA) <> "" Then xSet = xSet & ",CHQMONSTA": xValues = xValues & " ,'" & newY.CHQMONSTA & "'"

xSql = "Insert into " & paramIBM_Library_SABSPE & ".YCHQMON0" & xSet & ")" & xValues & ")"

Set rsADO = cnADO.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYCHQMON0_Insert = "Erreur màj : " & newY.CHQRC1ETA
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYCHQMON0_Insert = Error
End Function

Public Function sqlYCHQMON0_Update(newY As typeYCHQMON0, oldY As typeYCHQMON0, cnADO As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.CHQRC1ETA = newY.CHQRC1ETA _
And oldY.CHQRC1AGE = newY.CHQRC1AGE _
And oldY.CHQRC1SER = newY.CHQRC1SER _
And oldY.CHQRC1SSE = newY.CHQRC1SSE _
And oldY.CHQRC1OPE = newY.CHQRC1OPE _
And oldY.CHQRC1DOS = newY.CHQRC1DOS Then
    sqlYCHQMON0_Update = Null
Else
    sqlYCHQMON0_Update = "Erreur CHQRC1DOS: " & newY.CHQRC1DOS & " / " & oldY.CHQRC1DOS
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where CHQRC1ETA = " & oldY.CHQRC1ETA _
& " and CHQRC1AGE = " & oldY.CHQRC1AGE _
& " and CHQRC1SER = '" & oldY.CHQRC1SER & "'" _
& " and CHQRC1SSE = '" & oldY.CHQRC1SSE & "'" _
& " and CHQRC1OPE = '" & oldY.CHQRC1OPE & "'" _
& " and CHQRC1DOS = " & oldY.CHQRC1DOS _
& " and CHQMONUPDS = " & oldY.CHQMONUPDS

newY.CHQMONUPDS = newY.CHQMONUPDS + 1
xSet = xSet & " set CHQMONUPDS = " & newY.CHQMONUPDS

' Détecter les modifications !!!!!!!!!!!!! uniquement données en provenance de CHEQUE.mdb
'===================================================================================
If newY.CHQDATE <> oldY.CHQDATE Then xSet = xSet & " , CHQDATE = " & newY.CHQDATE
If newY.CHQCOMPTE <> oldY.CHQCOMPTE Then xSet = xSet & " , CHQCOMPTE = '" & Trim(newY.CHQCOMPTE) & "'"
If newY.CHQCREM <> oldY.CHQCREM Then xSet = xSet & " , CHQCREM = '" & newY.CHQCREM & "'"
If newY.CHQDEVISE <> oldY.CHQDEVISE Then xSet = xSet & " , CHQDEVISE = '" & newY.CHQDEVISE & "'"
If newY.CHQMONTANT <> oldY.CHQMONTANT Then xSet = xSet & " , CHQMONTANT = " & cur_P(newY.CHQMONTANT)
If newY.CHQNB <> oldY.CHQNB Then xSet = xSet & " , CHQNB = " & newY.CHQNB
If newY.CHQMONSTA <> oldY.CHQMONSTA Then xSet = xSet & " , CHQMONSTA = '" & newY.CHQMONSTA & "'"

xSql = "update " & paramIBM_Library_SABSPE & ".YCHQMON0" & xSet & xWhere

Set rsADO = cnADO.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYCHQMON0_Update = "Erreur màj : " & newY.CHQRC1DOS
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYCHQMON0_Update = Error
End Function
Public Function sqlYCHQMON0_Delete(oldY As typeYCHQMON0, cnADO As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlYCHQMON0_Delete = Null

' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where CHQRC1ETA = " & oldY.CHQRC1ETA _
& " and CHQRC1AGE = " & oldY.CHQRC1AGE _
& " and CHQRC1SER = '" & oldY.CHQRC1SER & "'" _
& " and CHQRC1SSE = '" & oldY.CHQRC1SSE & "'" _
& " and CHQRC1OPE = '" & oldY.CHQRC1OPE & "'" _
& " and CHQRC1DOS = " & oldY.CHQRC1DOS _
& " and CHQMONUPDS = " & oldY.CHQMONUPDS


xSql = "delete  from " & paramIBM_Library_SABSPE & ".YCHQMON0" & xWhere

Set rsADO = cnADO.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYCHQMON0_Delete = "Erreur màj : " & oldY.CHQRC1DOS
    Exit Function
End If
Exit Function
Error_Handler:
    sqlYCHQMON0_Delete = Error
End Function

Public Function srvYCHQMON0_GetBuffer_ODBC(rsADO As ADODB.Recordset, lYCHQMON0 As typeYCHQMON0)
On Error GoTo Error_Handler
srvYCHQMON0_GetBuffer_ODBC = Null
lYCHQMON0.CHQRC1ETA = rsADO("CHQRC1ETA")
lYCHQMON0.CHQRC1AGE = rsADO("CHQRC1AGE")
lYCHQMON0.CHQRC1SER = rsADO("CHQRC1SER")
lYCHQMON0.CHQRC1SSE = rsADO("CHQRC1SSE")
lYCHQMON0.CHQRC1OPE = rsADO("CHQRC1OPE")
lYCHQMON0.CHQRC1DOS = rsADO("CHQRC1DOS")
lYCHQMON0.CHQRC1DCR = rsADO("CHQRC1DCR")
lYCHQMON0.CHQDATE = rsADO("CHQDATE")
lYCHQMON0.CHQCOMPTE = rsADO("CHQCOMPTE")
lYCHQMON0.CHQCREM = rsADO("CHQCREM")
lYCHQMON0.CHQDEVISE = rsADO("CHQDEVISE")
lYCHQMON0.CHQMONTANT = rsADO("CHQMONTANT")
lYCHQMON0.CHQNB = rsADO("CHQNB")
lYCHQMON0.CHQMONSTA = rsADO("CHQMONSTA")
lYCHQMON0.CHQMONUPDS = rsADO("CHQMONUPDS")

Exit Function
Error_Handler:
srvYCHQMON0_GetBuffer_ODBC = Error


End Function

Public Function srvYCHQMON0_Init(lYCHQMON0 As typeYCHQMON0)
lYCHQMON0.CHQRC1ETA = 0
lYCHQMON0.CHQRC1AGE = 0
lYCHQMON0.CHQRC1SER = ""
lYCHQMON0.CHQRC1SSE = ""
lYCHQMON0.CHQRC1OPE = ""
lYCHQMON0.CHQRC1DOS = 0
lYCHQMON0.CHQRC1DCR = 0
lYCHQMON0.CHQDATE = 0
lYCHQMON0.CHQCOMPTE = ""
lYCHQMON0.CHQCREM = ""
lYCHQMON0.CHQDEVISE = ""
lYCHQMON0.CHQMONTANT = 0
lYCHQMON0.CHQNB = 0
lYCHQMON0.CHQMONSTA = ""
lYCHQMON0.CHQMONUPDS = 0

End Function

Public Sub srvYCHQMON0_fgDisplay(lYCHQMON0 As typeYCHQMON0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 2
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "CHQRC1DOS   9P"
fgDisplay.Col = 1: fgDisplay = "Identification"
fgDisplay.Col = 2: fgDisplay = lYCHQMON0.CHQRC1DOS
End Sub



Public Function srvYCHQMON0_Read(oldY As typeYCHQMON0, cnADO As ADODB.Connection)
Dim xWhere As String, xSql As String
Dim Nb As Long
srvYCHQMON0_Read = "?"

xWhere = " where CHQRC1ETA = " & oldY.CHQRC1ETA _
& " and CHQRC1AGE = " & oldY.CHQRC1AGE _
& " and CHQRC1SER = '" & oldY.CHQRC1SER & "'" _
& " and CHQRC1SSE = '" & oldY.CHQRC1SSE & "'" _
& " and CHQRC1OPE = '" & oldY.CHQRC1OPE & "'" _
& " and CHQRC1DOS = " & oldY.CHQRC1DOS _
& " and CHQMONUPDS = " & oldY.CHQMONUPDS

xSql = "Select * from " & paramIBM_Library_SABSPE & ".YCHQMON0" & xWhere

Set rsADO = cnADO.Execute(xSql, Nb)

'===================================================================================

If Not rsADO.EOF Then
    srvYCHQMON0_Read = srvYCHQMON0_GetBuffer_ODBC(rsADO, oldY)
Else
    srvYCHQMON0_Read = "?Inconnu"
End If

End Function
