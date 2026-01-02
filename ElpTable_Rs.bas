Attribute VB_Name = "rsElpTable"
Option Explicit

Type typeElpTable
    Id           As String * 12
    K1           As String * 12
    K2           As String * 12
    SNN          As Long
    SNP          As Long
    SN           As Long
    Chrono       As Long
    Name         As String * 36
    Dmin         As String * 8
    Dmax         As String * 8
    Memo         As Variant

End Type

Public Sub rsElpTable_Init(rsElpTable As typeElpTable)
rsElpTable.Id = ""
rsElpTable.K1 = ""
rsElpTable.K2 = ""
rsElpTable.SNN = 0
rsElpTable.SNP = 0
rsElpTable.SN = 0
rsElpTable.SNN = 0
rsElpTable.Name = ""
rsElpTable.Dmin = "00000000"
rsElpTable.Dmax = "00000000"
rsElpTable.Memo = ""
End Sub

'---------------------------------------------------------
Public Sub rsElpTable_GetBuffer(rsADO As ADODB.Recordset, rsElpTable As typeElpTable)
'---------------------------------------------------------

rsElpTable.Id = rsADO("Id")
rsElpTable.K1 = rsADO("K1")
rsElpTable.K2 = rsADO("K2")

rsElpTable.SNN = rsADO("SNN")
rsElpTable.SNP = rsADO("SNP")
rsElpTable.SN = rsADO("SN")
rsElpTable.Chrono = rsADO("Chrono")
rsElpTable.Name = rsADO("Name")
rsElpTable.Dmin = rsADO("DMin")
rsElpTable.Dmax = rsADO("DMax")
rsElpTable.Memo = rsADO("Memo")

End Sub
'---------------------------------------------------------
Public Function rsElpTable_Read(lId As String, lK1 As String, lK2 As String, lName As String, lMemo As String)
'---------------------------------------------------------
Dim xElpTable As typeElpTable
Dim X As String, V
Dim rsMDB As New ADODB.Recordset
On Error GoTo Error_Handler

rsElpTable_Read = Null
lName = ""
lMemo = ""

X = "select Name,Memo from ElpTable where SNN = 0" _
    & " and id = '" & lId & "'" _
    & " and K1 = '" & lK1 & "'" _
    & " and K2 = '" & lK2 & "'"
    
Set rsMDB = cnMDB.Execute(X)
If Not rsMDB.EOF Then
    lName = rsMDB("Name")
    V = rsMDB("Memo")
    If Not IsNull(V) Then lMemo = Trim(V)
    If Trim(lK1) = "PasswordX" Then lMemo = ElpCipher_D(lMemo, paramElpCypher)

Else
    rsElpTable_Read = "? rsElpTable_Read : " & lId & "_" & lK1 & "_" & lK2
End If
Exit Function

Error_Handler:
'-------------
    rsElpTable_Read = " rsElpTable_Read : " & Error & " " & Now
End Function

'---------------------------------------------------------
Public Function rsElpTable_Memo(lId As String, lK1 As String, lK2 As String, lName As String, lMemo As String)
'---------------------------------------------------------
Dim xElpTable As typeElpTable
Dim X As String, V
Dim rsMDB As New ADODB.Recordset
On Error GoTo Error_Handler

rsElpTable_Memo = Null
lName = ""
lMemo = ""

X = "select Name,Memo from ElpTable where SNN = 0" _
    & " and id = '" & lId & "'" _
    & " and K1 = '" & lK1 & "'" _
    & " and K2 = '" & lK2 & "'"
    
Set rsMDB = cnMDB.Execute(X)
If Not rsMDB.EOF Then
    lName = rsMDB("Name")
    V = rsMDB("Memo")
    If Not IsNull(V) Then lMemo = CStr(V)
    If Trim(lK1) = "PasswordX" Then lMemo = ElpCipher_D(lMemo, paramElpCypher)

Else
    rsElpTable_Memo = "? rsElpTable_Memo : " & lId & "_" & lK1 & "_" & lK2
End If
Exit Function

Error_Handler:
'-------------
    rsElpTable_Memo = " rsElpTable_Read : " & Error
End Function



