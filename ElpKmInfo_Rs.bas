Attribute VB_Name = "rsElpKmInfo"
'---------------------------------------------------------
Option Explicit
Type typeElpKmInfo
    ElpKMSrc_Id             As Long
    Id                      As String * 20
    Description             As String * 40
    Pass                    As Long
    Memo                    As Variant
    
End Type

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsElpKmInfo_GetBuffer(rsAdo As ADODB.Recordset, rsElpKmInfo As typeElpKmInfo)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsElpKmInfo_GetBuffer = Null

rsElpKmInfo.ElpKMSrc_Id = rsAdo("ElpKMSrc_Id")
rsElpKmInfo.Id = rsAdo("ID")
rsElpKmInfo.Description = rsAdo("Description")
rsElpKmInfo.Pass = rsAdo("Pass")
rsElpKmInfo.Memo = rsAdo("Memo")

Exit Function

Error_Handler:

rsElpKmInfo_GetBuffer = Error

End Function

'---------------------------------------------------------
Public Function rsElpKmInfo_Read(lElpKmInfo As typeElpKmInfo)
'---------------------------------------------------------
Dim X As String, V
Dim rsMDB As New ADODB.Recordset
On Error GoTo Error_Handler

rsElpKmInfo_Read = Null
lElpKmInfo.Memo = ""
lElpKmInfo.Description = ""

X = "select * from ElpKmInfo " _
    & " where ElpKMSrc_Id = " & lElpKmInfo.ElpKMSrc_Id _
    & " and ID = '" & lElpKmInfo.Id & "'"
    
Set rsMDB = cnMDB.Execute(X)
If Not rsMDB.EOF Then
    rsElpKmInfo_Read = rsElpKmInfo_GetBuffer(rsMDB, lElpKmInfo)
Else
    rsElpKmInfo_Read = "? rsElpKmInfo_Read : " & lElpKmInfo.ElpKMSrc_Id & "_" & lElpKmInfo.Id
End If
Exit Function

Error_Handler:
'-------------
    rsElpKmInfo_Read = " rsElpKmInfo_Read : " & Error
End Function


