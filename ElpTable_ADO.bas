Attribute VB_Name = "adoElpTable"
Option Explicit

'---------------------------------------------------------
Public Function rsElpTable_PutBuffer(rsADO As ADODB.Recordset, rsElpTable As typeElpTable)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsElpTable_PutBuffer = Null

rsADO("id") = rsElpTable.Id
rsADO("K1") = rsElpTable.K1
rsADO("K2") = rsElpTable.K2
rsADO("SNN") = rsElpTable.SNN
rsADO("SNP") = rsElpTable.SNP
rsADO("SN") = rsElpTable.SN
rsADO("Chrono") = rsElpTable.Chrono
rsADO("Name") = rsElpTable.Name

rsADO("DMin") = rsElpTable.Dmin
rsADO("DMax") = rsElpTable.Dmax
rsADO("Memo") = rsElpTable.Memo
Exit Function

Error_Handler:

rsElpTable_PutBuffer = Error
End Function

'---------------------------------------------------------
Public Function adoElpTable_AddNew(rsADO As ADODB.Recordset, rsElpTable As typeElpTable)
'---------------------------------------------------------
On Error GoTo Error_Handler

If rsMDB.State = adStateOpen Then rsMDB.Close
rsMDB.Open "select * from ElpTable", cnMDB, , adLockOptimistic
rsADO.AddNew
adoElpTable_AddNew = rsElpTable_PutBuffer(rsADO, rsElpTable)
rsADO.Update

Exit Function

Error_Handler:

adoElpTable_AddNew = Error

End Function

'---------------------------------------------------------
Public Function adoElpTable_Update(rsADO As ADODB.Recordset, rsElpTable As typeElpTable)
'---------------------------------------------------------
Dim xSql As String

On Error GoTo Error_Handler
xSql = "select * from ElpTable where SNN = " & rsElpTable.SNN _
    & " and id = '" & rsElpTable.Id & "'" _
    & " and K1 = '" & rsElpTable.K1 & "'" _
    & " and K2 = '" & rsElpTable.K2 & "'"

If rsMDB.State = adStateOpen Then rsMDB.Close

rsMDB.Open xSql, cnMDB, , adLockOptimistic
adoElpTable_Update = rsElpTable_PutBuffer(rsADO, rsElpTable)

rsADO.Update

Exit Function

Error_Handler:

adoElpTable_Update = Error

End Function

Public Function adoElpTable_Delete(rsADO As ADODB.Recordset, rsElpTable As typeElpTable)
'---------------------------------------------------------
Dim xSql As String

On Error GoTo Error_Handler
adoElpTable_Delete = Null

xSql = "delete * from ElpTable where SNN = " & rsElpTable.SNN _
    & " and id = '" & rsElpTable.Id & "'" _
    & " and K1 = '" & rsElpTable.K1 & "'" _
    & " and K2 = '" & rsElpTable.K2 & "'"
Call FEU_ROUGE
Set rsADO = cnMDB.Execute(xSql)
Call FEU_VERT
Exit Function

Error_Handler:

adoElpTable_Delete = Error

End Function


