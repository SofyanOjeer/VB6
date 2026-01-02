Attribute VB_Name = "sqlDAUTLIB0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsADO As ADODB.Recordset
 
Type typeDAUTLIB0
 
      DAUTLIBCOD     As String * 20
      DAUTLIBTXT     As String * 64
      DAUTLIBRGP     As String * 20
      DAUTLIBELM     As String * 3
      DAUTLIBAMO     As String * 3
      
      
End Type
Public xDAUTLIB0 As typeDAUTLIB0

Public Function sqlDAUTLIB0_Insert(newY As typeDAUTLIB0)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlDAUTLIB0_Insert = Null

xSet = " (DAUTLIBCOD"
xValues = " values('" & newY.DAUTLIBCOD & "'"

' Détecter les modifications
'===================================================================================
If Trim(newY.DAUTLIBTXT) <> "" Then xSet = xSet & ",DAUTLIBTXT": xValues = xValues & " ,'" & Replace(Trim(newY.DAUTLIBTXT), "'", "''") & "'"
If Trim(newY.DAUTLIBRGP) <> "" Then xSet = xSet & ",DAUTLIBRGP": xValues = xValues & " ,'" & newY.DAUTLIBRGP & "'"
If Trim(newY.DAUTLIBELM) <> "" Then xSet = xSet & ",DAUTLIBELM": xValues = xValues & " ,'" & newY.DAUTLIBELM & "'"
If Trim(newY.DAUTLIBAMO) <> "" Then xSet = xSet & ",DAUTLIBAMO": xValues = xValues & " ,'" & newY.DAUTLIBAMO & "'"


xSql = "Insert into BODWH.DAUTLIB0" & xSet & ")" & xValues & ")"

Set rsADO = cnSab_Update.Execute(xSql, Nb)

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

Public Function rsDAUTLIB0_GetBuffer(rsADO As ADODB.Recordset, lDAUTLIB0 As typeDAUTLIB0)
On Error GoTo Error_Handler
rsDAUTLIB0_GetBuffer = Null

lDAUTLIB0.DAUTLIBCOD = rsADO("DAUTLIBCOD")
lDAUTLIB0.DAUTLIBTXT = rsADO("DAUTLIBTXT")
lDAUTLIB0.DAUTLIBRGP = rsADO("DAUTLIBRGP")
lDAUTLIB0.DAUTLIBELM = rsADO("DAUTLIBELM")
lDAUTLIB0.DAUTLIBAMO = rsADO("DAUTLIBAMO")

Exit Function
Error_Handler:
rsDAUTLIB0_GetBuffer = Error


End Function





