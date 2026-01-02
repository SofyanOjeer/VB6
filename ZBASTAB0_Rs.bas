Attribute VB_Name = "rsZBASTAB0"
'---------------------------------------------------------
Option Explicit
Type typeZBASTAB0
    BASTABETA       As Integer
    BASTABNUM       As Integer
    BASTABARG       As String * 16
    BASTABLO1       As String * 12
    BASTABLO2       As String * 12
    BASTABDON       As String * 212
    
    
End Type

Dim xZBASTAB0 As typeZBASTAB0


Type typePays
    Id          As String * 2
    Nom         As String * 30
    Fiscal      As String * 1
End Type

Public sabPays() As typePays, sabPays_NB As Integer
Public sabPays_FR As Integer, sabPays_DZ As Integer, sabPays_LY As Integer, sabPays_US As Integer


'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZBASTAB0_GetBuffer(rsAdo As ADODB.Recordset, rsZBASTAB0 As typeZBASTAB0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZBASTAB0_GetBuffer = Null

rsZBASTAB0.BASTABETA = rsAdo("BASTABETA")
rsZBASTAB0.BASTABNUM = rsAdo("BASTABNUM")
rsZBASTAB0.BASTABARG = rsAdo("BASTABARG")
rsZBASTAB0.BASTABLO1 = rsAdo("BASTABLO1")
rsZBASTAB0.BASTABLO2 = rsAdo("BASTABLO2")
rsZBASTAB0.BASTABDON = rsAdo("BASTABDON")
Exit Function

Error_Handler:

rsZBASTAB0_GetBuffer = Error

End Function

Public Sub rsZBASTAB0_cboK2(lId As Integer, cbo As Control, lAnd As String)
Dim X As String, K1 As Integer
Dim V

cbo.Clear
X = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0 " _
    & " where BASTABNUM = " & lId & lAnd
    
Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    V = rsZBASTAB0_GetBuffer(rsSab, xZBASTAB0)
    If IsNull(V) Then
        Select Case lId
            Case 5: cbo.AddItem Mid$(xZBASTAB0.BASTABARG, 4, 4) & " : " & Trim(xZBASTAB0.BASTABLO2) & Mid$(rsSab("BASTABDON"), 1, 30)
            Case 14: cbo.AddItem Trim(xZBASTAB0.BASTABARG) & " : " & Trim(xZBASTAB0.BASTABLO2) & Mid$(rsSab("BASTABDON"), 1, 30)
            Case 23: cbo.AddItem Trim(xZBASTAB0.BASTABLO1) & Trim(xZBASTAB0.BASTABARG) & " : " & Mid$(rsSab("BASTABDON"), 1, 30)
            Case 44: cbo.AddItem Mid$(xZBASTAB0.BASTABARG, 1, 6) & " * " & Mid$(rsSab("BASTABDON"), 31, 1) & " : " & Mid$(rsSab("BASTABDON"), 1, 30)
            Case 58: cbo.AddItem Trim(xZBASTAB0.BASTABARG) & " : " & Mid$(rsSab("BASTABDON"), 1, 30)
       End Select
    End If
    rsSab.MoveNext
Loop

End Sub


Public Sub rsZBASTAB0_Pays(lPays() As typePays, lPays_NB As Integer)
Dim X As String, K1 As Integer
Dim rsSab As New ADODB.Recordset
Dim V
X = "select count(*) as Tally from  " & paramIBM_Library_SAB & ".ZBASTAB0 " _
    & " where BASTABNUM = 11 "
Set rsSab = cnsab.Execute(X)
lPays_NB = rsSab("Tally")
ReDim lPays(lPays_NB)
K1 = 0

X = "select * from  " & paramIBM_Library_SAB & ".ZBASTAB0 " _
    & " where BASTABNUM = 11 order by BASTABARG"
Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    lPays(K1).Id = Mid$(rsSab("BASTABARG"), 4, 2)
    X = rsSab("BASTABDON")
    lPays(K1).Fiscal = Mid$(X, 26, 1)
    lPays(K1).Nom = Mid$(rsSab("BASTABLO2"), 4, 16) & Mid$(X, 1, 25)
    K1 = K1 + 1
    rsSab.MoveNext
Loop

End Sub

Public Function rsZBASTAB0_Cours37(lDev As String, lAMJ As Long) As Double
Static wCours As Double, wCours_AMJ As Long, wCours_Dev As String
Dim rsSabX As New ADODB.Recordset
Dim xWhere As String, xSQL As String, wAMJ_Min As Long
On Error GoTo Error_Handler
If wCours_AMJ <> lAMJ Or wCours_Dev <> lDev Then
    wCours = 0
    wCours_AMJ = lAMJ
    wCours_Dev = lDev
    Set rsSabX = Nothing
    xWhere = " where BASTABNUM = 37 and substring(BASTABARG , 1 , 3) = '" & wCours_Dev & "'" _
           & " and substring(BASTABARG , 4 , 4) = x'" & lAMJ & "F'"
                
    xSQL = "select BASTABDON from " & paramIBM_Library_SAB & ".ZBASTAB0 " & xWhere
    Set rsSabX = cnsab.Execute(xSQL)
    If Not rsSabX.EOF Then
        wCours = CDbl(convX2P(Mid$(rsSabX("BASTABDON"), 1, 8))) / 1000000000#
    Else
    End If
    
End If
rsZBASTAB0_Cours37 = wCours

Exit Function
Error_Handler:
    wCours_AMJ = 0
    wCours = 0
End Function
Public Function rsZBASTAB0_Cours37_avant_20120319(lDev As String, lAMJ As Long) As Double
Static wCours As Double, wCours_AMJ As Long, wCours_Dev As String
Dim rsSabX As New ADODB.Recordset
Dim xWhere As String, xSQL As String, wAMJ_Min As Long
On Error GoTo Error_Handler
If wCours_AMJ <> lAMJ Or wCours_Dev <> lDev Then
    wCours = 0
    wCours_AMJ = lAMJ
    wCours_Dev = lDev
    Set rsSabX = Nothing
    xWhere = " where BASTABNUM = 37 and BASTABARG = '" & wCours_Dev & convP2X(wCours_AMJ, 7) & "'"
                
    xSQL = "select BASTABDON from " & paramIBM_Library_SAB & ".ZBASTAB0 " & xWhere
    Set rsSabX = cnsab.Execute(xSQL)
    If Not rsSabX.EOF Then
        wCours = CDbl(convX2P(Mid$(rsSabX("BASTABDON"), 1, 8))) / 1000000000#
    Else
    End If
    
End If
rsZBASTAB0_Cours37_avant_20120319 = wCours

Exit Function
Error_Handler:
wAMJ_Min = dateElp("FinDeMoisP", 0, CStr(wCours_AMJ + 19000000))
xWhere = " where BASTABNUM = 37 and BASTABARG < '" & wCours_Dev & convP2X(wCours_AMJ + 11, 7) & "'" _
       & " and BASTABARG > '" & wCours_Dev & convP2X(wAMJ_Min - 19000000, 7) & "'" _
       & " order by BASTABARG"

xSQL = "select BASTABARG,BASTABDON from " & paramIBM_Library_SAB & ".ZBASTAB0 " & xWhere
Set rsSabX = cnsab.Execute(xSQL)
Do Until rsSabX.EOF
    If convX2P(Mid$(rsSabX("BASTABARG"), 4, 4)) > wCours_AMJ Then Exit Function
    wCours = CDbl(convX2P(Mid$(rsSabX("BASTABDON"), 1, 8))) / 1000000000#
    'Debug.Print convX2P(mId$(rsSabX("BASTABARG"), 4, 4)), wCours
    rsSabX.MoveNext
Loop
'Else
'End If

End Function


