Attribute VB_Name = "srvDAUTLIB0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsADO As ADODB.Recordset
 
Type typeDAUTLIB0
 
      DAUTLIBCOD     As String * 20
      DAUTLIBTXT     As String * 64
      DAUTLIBRGP     As String * 20
      DAUTLIBELM     As String * 3
      DAUTLIBAMO    As String * 3

End Type
Public xDAUTLIB0 As typeDAUTLIB0
Public Function sqlDAUTLIB0_Insert(newY As typeDAUTLIB0, cnAdo As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlDAUTLIB0_Insert = Null

xSet = " (DAUTLIBCOD"
xValues = " values('" & Replace(Trim(newY.DAUTLIBCOD), "'", "''") & "'"

' Insertion :
'===================================================================================
If Trim(newY.DAUTLIBTXT) <> "" Then xSet = xSet & ",DAUTLIBTXT": xValues = xValues & ", '" & Replace(Trim(newY.DAUTLIBTXT), "'", "''") & "'"
If Trim(newY.DAUTLIBRGP) <> "" Then xSet = xSet & ",DAUTLIBRGP": xValues = xValues & ", '" & Replace(Trim(newY.DAUTLIBRGP), "'", "''") & "'"
If Trim(newY.DAUTLIBELM) <> "" Then xSet = xSet & ",DAUTLIBELM": xValues = xValues & ", '" & Replace(Trim(newY.DAUTLIBELM), "'", "''") & "'"
If Trim(newY.DAUTLIBAMO) <> "" Then xSet = xSet & ",DAUTLIBAMO": xValues = xValues & ", '" & Replace(Trim(newY.DAUTLIBAMO), "'", "''") & "'"

xSql = "Insert into " & paramIBM_Library_BODWH & ".DAUTLIB0" & xSet & ")" & xValues & ")"

Set rsADO = cnAdo.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlDAUTLIB0_Insert = "Erreur màj : " & newY.DAUTLIBCOD
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlDAUTLIB0_Insert = Error
End Function

Public Function sqlDAUTLIB0_Delete(oldY As typeDAUTLIB0, cnAdo As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlDAUTLIB0_Delete = Null


xWhere = " where DAUTLIBCOD = '" & Trim(oldY.DAUTLIBCOD) & "'"

' Suppression physique
'===================================================================================

xSql = "Delete from " & paramIBM_Library_BODWH & ".DAUTLIB0" & xWhere

Set rsADO = cnAdo.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlDAUTLIB0_Delete = "Erreur SUP : " & oldY.DAUTLIBCOD
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlDAUTLIB0_Delete = Error
End Function

Public Function sqlDAUTLIB0_Read(oldY As typeDAUTLIB0, cnAdo As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim V

On Error GoTo Error_Handler
sqlDAUTLIB0_Read = Null
xSql = "select * from " & paramIBM_Library_BODWH & ".DAUTLIB0 where DAUTLIBCOD ='" & Trim(oldY.DAUTLIBCOD) & "'"
Set rsADO = cnAdo.Execute(xSql)

If rsADO.EOF Then
    sqlDAUTLIB0_Read = "? inconnu"
Else
    V = srvDAUTLIB0_GetBuffer_ODBC(rsADO, oldY)
    If Not IsNull(V) Then sqlDAUTLIB0_Read = "? srvDAUTLIB0_GetBuffer_ODBC"
End If
 
Exit Function
Error_Handler:
    sqlDAUTLIB0_Read = Error
End Function

Public Function sqlDAUTLIB0_Update(newY As typeDAUTLIB0, oldY As typeDAUTLIB0, cnAdo As ADODB.Connection)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlDAUTLIB0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.DAUTLIBCOD <> newY.DAUTLIBCOD Then
    sqlDAUTLIB0_Update = "Clé erronnée lors mise à jour !"
    Exit Function
End If

' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================
xWhere = " where DAUTLIBCOD = '" & Replace(Trim(newY.DAUTLIBCOD), "'", "''") & "'"


xSet = " set DAUTLIBTXT = '" & Replace(Trim(newY.DAUTLIBTXT), "'", "''") & "'"

' Détecter les modifications
'===================================================================================
If Trim(newY.DAUTLIBRGP) <> Trim(oldY.DAUTLIBRGP) Then xSet = xSet & ",DAUTLIBRGP='" & Replace(Trim(newY.DAUTLIBRGP), "'", "''") & "'"
If Trim(newY.DAUTLIBELM) <> Trim(oldY.DAUTLIBELM) Then xSet = xSet & ",DAUTLIBELM='" & Replace(Trim(newY.DAUTLIBELM), "'", "''") & "'"
If Trim(newY.DAUTLIBAMO) <> Trim(oldY.DAUTLIBAMO) Then xSet = xSet & ",DAUTLIBAMO='" & Replace(Trim(newY.DAUTLIBAMO), "'", "''") & "'"

xSql = "update " & paramIBM_Library_BODWH & ".DAUTLIB0" & xSet & xWhere

Set rsADO = cnAdo.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlDAUTLIB0_Update = "Erreur màj : " & newY.DAUTLIBCOD

    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlDAUTLIB0_Update = Error
End Function


Public Function srvDAUTLIB0_GetBuffer_ODBC(rsADO As ADODB.Recordset, lDAUTLIB0 As typeDAUTLIB0)

On Error GoTo Error_Handler

srvDAUTLIB0_GetBuffer_ODBC = Null

lDAUTLIB0.DAUTLIBCOD = rsADO("DAUTLIBCOD")
lDAUTLIB0.DAUTLIBTXT = rsADO("DAUTLIBTXT")
lDAUTLIB0.DAUTLIBRGP = rsADO("DAUTLIBRGP")
lDAUTLIB0.DAUTLIBELM = rsADO("DAUTLIBELM")
lDAUTLIB0.DAUTLIBAMO = rsADO("DAUTLIBAMO")

Exit Function
Error_Handler:
srvDAUTLIB0_GetBuffer_ODBC = Error

End Function

Public Function srvDAUTLIB0_Init(lDAUTLIB0 As typeDAUTLIB0)

lDAUTLIB0.DAUTLIBCOD = ""
lDAUTLIB0.DAUTLIBTXT = ""
lDAUTLIB0.DAUTLIBRGP = ""
lDAUTLIB0.DAUTLIBELM = ""
lDAUTLIB0.DAUTLIBAMO = ""

End Function

Public Sub srvDAUTLIB0_fgDisplay(lDAUTLIB0 As typeDAUTLIB0, fgDisplay As MSFlexGrid)

fgDisplay.Rows = 6

fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "DAUTLIBCOD    20A"
fgDisplay.Col = 1: fgDisplay = "Code"
fgDisplay.Col = 2: fgDisplay = lDAUTLIB0.DAUTLIBCOD
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "DAUTLIBTXT    64A"
fgDisplay.Col = 1: fgDisplay = "Libellé"
fgDisplay.Col = 2: fgDisplay = lDAUTLIB0.DAUTLIBTXT
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "DAUTLIBRGP     20A"
fgDisplay.Col = 1: fgDisplay = "Regroupement"
fgDisplay.Col = 2: fgDisplay = lDAUTLIB0.DAUTLIBRGP
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "DAUTLIBELM     3A"
fgDisplay.Col = 1: fgDisplay = "Elémentaire"
fgDisplay.Col = 2: fgDisplay = lDAUTLIB0.DAUTLIBELM
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "DAUTLIBAMO      3A"
fgDisplay.Col = 1: fgDisplay = "Amortissable"
fgDisplay.Col = 2: fgDisplay = lDAUTLIB0.DAUTLIBAMO

End Sub



