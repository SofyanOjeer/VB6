Attribute VB_Name = "srvDAUTPIB"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsado As ADODB.Recordset
 
Type typeDAUTPIB
 
      DAUTSTA     As String * 1   ' STATUT
      DAUTVER     As Integer      ' No VERSION
      DAUTPER     As Long         ' PERIODE TRAITEMENT
      DAUTETB     As String * 2   ' CODE ETABLISSEMENT
      DAUTCLI     As Long         ' CODE MATRICULE
      DAUTAUT     As String * 20  ' CODE AUTORISATION
      DAUTDEV     As String * 3   ' CODE DEVISE
      DAUTMON     As Currency     ' MONTANT AUTORISATION
      DAUTECH     As Long         ' DATE ECHEANCE
      DAUTMAJ     As Long         ' Sequence mise à jour

End Type
Public xDAUTPIB As typeDAUTPIB

Public Function sqlDAUTPIB_Insert(newY As typeDAUTPIB, cnADO As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlDAUTPIB_Insert = Null

xSet = " (DAUTVER"
xValues = " values(" & newY.DAUTVER

' Insertion : Normalement tout est à créer :
'===================================================================================
If Trim(newY.DAUTSTA) <> "" Then xSet = xSet & ",DAUTSTA": xValues = xValues & ", '" & newY.DAUTSTA & "'"
If newY.DAUTPER <> 0 Then xSet = xSet & ",DAUTPER": xValues = xValues & ", " & newY.DAUTPER
If Trim(newY.DAUTETB) <> "" Then xSet = xSet & ",DAUTETB": xValues = xValues & ", '" & newY.DAUTETB & "'"
If newY.DAUTCLI <> 0 Then xSet = xSet & ",DAUTCLI": xValues = xValues & ", " & newY.DAUTCLI
If newY.DAUTAUT <> "" Then xSet = xSet & ",DAUTAUT": xValues = xValues & ", '" & newY.DAUTAUT & "'"
If newY.DAUTDEV <> "" Then xSet = xSet & ",DAUTDEV": xValues = xValues & ", '" & newY.DAUTDEV & "'"
If newY.DAUTMON <> 0 Then xSet = xSet & ",DAUTMON": xValues = xValues & ", " & cur_P(newY.DAUTMON)
If newY.DAUTECH <> 0 Then xSet = xSet & ",DAUTECH": xValues = xValues & ", " & newY.DAUTECH

xSql = "Insert into BODWH.DAUTPIB" & xSet & ")" & xValues & ")"

Set rsado = cnADO.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlDAUTPIB_Insert = "Erreur màj : " & newY.DAUTVER & newY.DAUTPER & newY.DAUTETB & newY.DAUTCLI & newY.DAUTAUT & newY.DAUTDEV
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlDAUTPIB_Insert = Error
End Function

Public Function sqlDAUTPIB_Delete(newY As typeDAUTPIB, oldY As typeDAUTPIB, cnADO As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlDAUTPIB_Delete = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.DAUTVER <> newY.DAUTVER Or oldY.DAUTPER <> newY.DAUTPER Or oldY.DAUTETB <> newY.DAUTETB Or _
oldY.DAUTCLI <> newY.DAUTCLI Or oldY.DAUTAUT <> newY.DAUTAUT Or oldY.DAUTDEV <> newY.DAUTDEV Then
    sqlDAUTPIB_Delete = "Clé erronnée lors de la suppression !"
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Delete'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where DAUTVER = " & oldY.DAUTVER & " and DAUTPER = " & oldY.DAUTPER _
         & " and DAUTETB = '" & oldY.DAUTETB & "' and DAUTCLI = " & oldY.DAUTCLI _
         & " and DAUTAUT = '" & oldY.DAUTAUT & "' and DAUTDEV = '" & oldY.DAUTDEV & "'" _
         & " and DAUTMAJ = " & oldY.DAUTMAJ

' Suppression physique
'===================================================================================

xSql = "Delete from " & paramIBM_Library_BODWH & ".DAUTPIB" & xWhere

Set rsado = cnADO.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlDAUTPIB_Delete = "Erreur SUPP : " & oldY.DAUTVER & oldY.DAUTPER & oldY.DAUTETB & oldY.DAUTCLI & oldY.DAUTAUT & oldY.DAUTDEV
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlDAUTPIB_Delete = Error
End Function

Public Function sqlDAUTPIB_Update(newY As typeDAUTPIB, oldY As typeDAUTPIB, cnADO As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlDAUTPIB_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.DAUTVER <> newY.DAUTVER Or oldY.DAUTPER <> newY.DAUTPER Or oldY.DAUTETB <> newY.DAUTETB Or _
oldY.DAUTCLI <> newY.DAUTCLI Or oldY.DAUTAUT <> newY.DAUTAUT Or oldY.DAUTDEV <> newY.DAUTDEV Then
    sqlDAUTPIB_Update = "Clé erronnée lors mise à jour !"
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where DAUTVER = " & oldY.DAUTVER & " and DAUTPER = " & oldY.DAUTPER _
         & " and DAUTETB = '" & oldY.DAUTETB & "' and DAUTCLI = " & oldY.DAUTCLI _
         & " and DAUTAUT = '" & oldY.DAUTAUT & "' and DAUTDEV = '" & oldY.DAUTDEV & "'" _
         & " and DAUTMAJ = " & oldY.DAUTMAJ

newY.DAUTMAJ = newY.DAUTMAJ + 1
xSet = xSet & " set DAUTMAJ = " & newY.DAUTMAJ

' Détecter les modifications
'===================================================================================
If Trim(newY.DAUTSTA) <> "" Then xSet = xSet & ", DAUTSTA='" & newY.DAUTSTA & "'"
If newY.DAUTMON <> 0 Then xSet = xSet & ", DAUTMON=" & cur_P(newY.DAUTMON)
If newY.DAUTECH <> 0 Then xSet = xSet & ", DAUTECH=" & newY.DAUTECH

xSql = "update BODWH.DAUTPIB" & xSet & xWhere

Set rsado = cnADO.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlDAUTPIB_Update = "Erreur màj : " & newY.DAUTVER & newY.DAUTPER & newY.DAUTETB & newY.DAUTCLI & newY.DAUTAUT & newY.DAUTDEV
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlDAUTPIB_Update = Error
End Function


Public Function srvDAUTPIB_GetBuffer_ODBC(rsado As ADODB.Recordset, lDAUTPIB As typeDAUTPIB)

On Error GoTo Error_Handler

srvDAUTPIB_GetBuffer_ODBC = Null

lDAUTPIB.DAUTSTA = rsado("DAUTSTA")
lDAUTPIB.DAUTVER = rsado("DAUTVER")
lDAUTPIB.DAUTPER = rsado("DAUTPER")
lDAUTPIB.DAUTETB = rsado("DAUTETB")
lDAUTPIB.DAUTCLI = rsado("DAUTCLI")
lDAUTPIB.DAUTAUT = rsado("DAUTAUT")
lDAUTPIB.DAUTDEV = rsado("DAUTDEV")
lDAUTPIB.DAUTMON = rsado("DAUTMON")
lDAUTPIB.DAUTECH = rsado("DAUTECH")
lDAUTPIB.DAUTMAJ = rsado("DAUTMAJ")

Exit Function
Error_Handler:
srvDAUTPIB_GetBuffer_ODBC = Error

End Function

Public Function srvDAUTPIB_Init(lDAUTPIB As typeDAUTPIB)

lDAUTPIB.DAUTSTA = ""
lDAUTPIB.DAUTVER = 0
lDAUTPIB.DAUTPER = 0
lDAUTPIB.DAUTETB = ""
lDAUTPIB.DAUTCLI = 0
lDAUTPIB.DAUTAUT = ""
lDAUTPIB.DAUTDEV = ""
lDAUTPIB.DAUTMON = 0
lDAUTPIB.DAUTECH = 0
lDAUTPIB.DAUTMAJ = 0

End Function

Public Sub srvDAUTPIB_fgDisplay(lDAUTPIB As typeDAUTPIB, fgDisplay As MSFlexGrid)

fgDisplay.Rows = 11

fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "DAUTSTA     1A"
fgDisplay.Col = 1: fgDisplay = "Statut"
fgDisplay.Col = 2: fgDisplay = lDAUTPIB.DAUTSTA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "DAUTVER     1S0"
fgDisplay.Col = 1: fgDisplay = "Version"
fgDisplay.Col = 2: fgDisplay = lDAUTPIB.DAUTVER
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "DAUTPER     8S0"
fgDisplay.Col = 1: fgDisplay = "Période de traitement"
fgDisplay.Col = 2: fgDisplay = lDAUTPIB.DAUTPER
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "DAUTETB     2A"
fgDisplay.Col = 1: fgDisplay = "Code établissement"
fgDisplay.Col = 2: fgDisplay = lDAUTPIB.DAUTETB
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "DAUTCLI    7S0"
fgDisplay.Col = 1: fgDisplay = "No Matricule"
fgDisplay.Col = 2: fgDisplay = lDAUTPIB.DAUTCLI
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "DAUTAUT    20A"
fgDisplay.Col = 1: fgDisplay = "Code autorisation"
fgDisplay.Col = 2: fgDisplay = lDAUTPIB.DAUTAUT
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "DAUTDEV     3A"
fgDisplay.Col = 1: fgDisplay = "Code devise"
fgDisplay.Col = 2: fgDisplay = lDAUTPIB.DAUTDEV
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "DAUTMON    18P2"
fgDisplay.Col = 1: fgDisplay = "Montant autorisation"
fgDisplay.Col = 2: fgDisplay = lDAUTPIB.DAUTMON
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "DAUTECH     8S0"
fgDisplay.Col = 1: fgDisplay = "Date échéance"
fgDisplay.Col = 2: fgDisplay = lDAUTPIB.DAUTECH
fgDisplay.Row = 10
fgDisplay.Col = 0: fgDisplay = "DAUTMAJ     5S0"
fgDisplay.Col = 1: fgDisplay = "Séquence mise à jour"
fgDisplay.Col = 2: fgDisplay = lDAUTPIB.DAUTMAJ

End Sub


