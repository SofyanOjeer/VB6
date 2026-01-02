Attribute VB_Name = "sqlZTCHCOR0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeZTCHCOR0
 
      TCHCORETB     As Long
      TCHCORCOD     As String * 3
      TCHCORAGS    As Long
      TCHCORRGP    As String * 1
      TCHCORTYP     As String * 1
      TCHCORCLI     As String * 7
      TCHCORBIC     As String * 12
      TCHCORDEV     As String * 3
      TCHCORTY1     As String * 1
      TCHCORCL1     As String * 7
      TCHCORBI1     As String * 12
      TCHCORMTR     As String * 1
      
      
End Type
Public xZTCHCOR0 As typeZTCHCOR0

Public Function sqlZTCHCOR0_Insert(newY As typeZTCHCOR0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlZTCHCOR0_Insert = Null

xSet = " (TCHCORETB"
xValues = " values(" & newY.TCHCORETB

' Détecter les modifications
'===================================================================================
If newY.TCHCORAGS <> 0 Then xSet = xSet & ",TCHCORAGS": xValues = xValues & " ," & cur_P(newY.TCHCORAGS)

If Trim(newY.TCHCORCOD) <> "" Then xSet = xSet & ",TCHCORCOD": xValues = xValues & " ,'" & newY.TCHCORCOD & "'"
If Trim(newY.TCHCORRGP) <> "" Then xSet = xSet & ",TCHCORRGP": xValues = xValues & " ,'" & newY.TCHCORRGP & "'"
If Trim(newY.TCHCORTYP) <> "" Then xSet = xSet & ",TCHCORTYP": xValues = xValues & " ,'" & Replace(Trim(newY.TCHCORTYP), "'", "''") & "'"
If Trim(newY.TCHCORCLI) <> "" Then xSet = xSet & ",TCHCORCLI": xValues = xValues & " ,'" & newY.TCHCORCLI & "'"
If Trim(newY.TCHCORBIC) <> "" Then xSet = xSet & ",TCHCORBIC": xValues = xValues & " ,'" & newY.TCHCORBIC & "'"
If Trim(newY.TCHCORDEV) <> "" Then xSet = xSet & ",TCHCORDEV": xValues = xValues & " ,'" & newY.TCHCORDEV & "'"
If Trim(newY.TCHCORTY1) <> "" Then xSet = xSet & ",TCHCORTY1": xValues = xValues & " ,'" & newY.TCHCORTY1 & "'"
If Trim(newY.TCHCORCL1) <> "" Then xSet = xSet & ",TCHCORCL1": xValues = xValues & " ,'" & newY.TCHCORCL1 & "'"
If Trim(newY.TCHCORBI1) <> "" Then xSet = xSet & ",TCHCORBI1": xValues = xValues & " ,'" & newY.TCHCORBI1 & "'"
If Trim(newY.TCHCORMTR) <> "" Then xSet = xSet & ",TCHCORMTR": xValues = xValues & " ,'" & newY.TCHCORMTR & "'"

Call FEU_ROUGE
xSQL = "Insert into " & paramIBM_Library_SAB & ".ZTCHCOR0" & xSet & ")" & xValues & ")"

Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlZTCHCOR0_Insert = "Erreur màj : " & newY.TCHCORBIC & " / " & newY.TCHCORBI1
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlZTCHCOR0_Insert = Error
End Function

Public Function rsZTCHCOR0_GetBuffer(rsAdo As ADODB.Recordset, lZTCHCOR0 As typeZTCHCOR0)
On Error GoTo Error_Handler
rsZTCHCOR0_GetBuffer = Null

lZTCHCOR0.TCHCORETB = rsAdo("TCHCORETB")
lZTCHCOR0.TCHCORCOD = rsAdo("TCHCORCOD")
lZTCHCOR0.TCHCORAGS = rsAdo("TCHCORAGS")
lZTCHCOR0.TCHCORRGP = rsAdo("TCHCORRGP")
lZTCHCOR0.TCHCORTYP = rsAdo("TCHCORTYP")
lZTCHCOR0.TCHCORCLI = rsAdo("TCHCORCLI")
lZTCHCOR0.TCHCORBIC = rsAdo("TCHCORBIC")
lZTCHCOR0.TCHCORDEV = rsAdo("TCHCORDEV")
lZTCHCOR0.TCHCORTY1 = rsAdo("TCHCORTY1")
lZTCHCOR0.TCHCORCL1 = rsAdo("TCHCORCL1")
lZTCHCOR0.TCHCORBI1 = rsAdo("TCHCORBI1")
lZTCHCOR0.TCHCORMTR = rsAdo("TCHCORMTR")

Exit Function
Error_Handler:
rsZTCHCOR0_GetBuffer = Error


End Function

Public Function rsZTCHCOR0_Init(lZTCHCOR0 As typeZTCHCOR0)

lZTCHCOR0.TCHCORETB = 1
lZTCHCOR0.TCHCORCOD = "001"
lZTCHCOR0.TCHCORAGS = 1
lZTCHCOR0.TCHCORRGP = "T"
lZTCHCOR0.TCHCORTYP = "B"
lZTCHCOR0.TCHCORCLI = ""
lZTCHCOR0.TCHCORBIC = ""
lZTCHCOR0.TCHCORDEV = ""
lZTCHCOR0.TCHCORTY1 = "B"
lZTCHCOR0.TCHCORCL1 = ""
lZTCHCOR0.TCHCORBI1 = ""
lZTCHCOR0.TCHCORMTR = "1"
End Function





