Attribute VB_Name = "sqlZBASTXX0"
Option Explicit

Type typeYBASTXX0
    BASTXXUAMJ      As String * 8
    BASTXXUHMS      As String * 6
    BASTXXUSEQ      As Long
    
    BASTXXNUM       As Long                           ' NUMERO TABLE
    BASTXXDEV       As String * 3                     ' devise
    BASTXXTAU       As String * 6                     ' code taux
    BASTXXAMJ       As Long                           ' date      1aammjj
    BASTXXVAL       As Double                         ' valeur 5v9s
End Type

Public Function sqlYBASTXX0_Insert(newY As typeYBASTXX0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYBASTXX0_Insert = Null

xSet = " ("
xValues = " values("

'===================================================================================

If Trim(newY.BASTXXUAMJ) <> "" Then xSet = xSet & ",BASTXXUAMJ": xValues = xValues & " ,'" & Trim(newY.BASTXXUAMJ) & "'"
If Trim(newY.BASTXXUHMS) <> "" Then xSet = xSet & ",BASTXXUHMS": xValues = xValues & " ,'" & Trim(newY.BASTXXUHMS) & "'"
xSet = xSet & ",BASTXXUSEQ": xValues = xValues & " ," & newY.BASTXXUSEQ
xSet = xSet & ",BASTXXNUM": xValues = xValues & " ," & newY.BASTXXNUM
If Trim(newY.BASTXXDEV) <> "" Then xSet = xSet & ",BASTXXDEV": xValues = xValues & " ,'" & Trim(newY.BASTXXDEV) & "'"
If Trim(newY.BASTXXTAU) <> "" Then xSet = xSet & ",BASTXXTAU": xValues = xValues & " ,'" & Trim(newY.BASTXXTAU) & "'"
xSet = xSet & ",BASTXXAMJ": xValues = xValues & " ," & newY.BASTXXAMJ
X = newY.BASTXXVAL
xSet = xSet & ",BASTXXVAL": xValues = xValues & " ," & Replace(X, ",", ".")

Mid$(xSet, 3, 1) = " "
Mid$(xValues, 10, 1) = " "
Call FEU_ROUGE
xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YBASTXX0" & xSet & ")" & xValues & ")"

Set rsSab_Update = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYBASTXX0_Insert = "Erreur màj : " & newY.BASTXXDEV
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYBASTXX0_Insert = Error
End Function
Public Sub recYBASTXX0_Init(lYBASTXX0 As typeYBASTXX0)

lYBASTXX0.BASTXXUAMJ = ""
lYBASTXX0.BASTXXUHMS = ""
lYBASTXX0.BASTXXUSEQ = 0

lYBASTXX0.BASTXXNUM = 0
lYBASTXX0.BASTXXDEV = ""
lYBASTXX0.BASTXXTAU = ""
lYBASTXX0.BASTXXAMJ = 0
lYBASTXX0.BASTXXVAL = 0

End Sub
