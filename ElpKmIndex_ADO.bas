Attribute VB_Name = "adoElpKmIndex"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsElpKmIndex_PutBuffer(rsADO As ADODB.Recordset, rsElpKmIndex As typeElpKmIndex)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsElpKmIndex_PutBuffer = Null
rsADO("ID") = rsElpKmIndex.ID
rsADO("Classe") = rsElpKmIndex.Classe
rsADO("ElpKMSrc_Id") = rsElpKmIndex.ElpKMSrc_Id
rsADO("Memo") = rsElpKmIndex.Memo

    
Exit Function

Error_Handler:

rsElpKmIndex_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoElpKmIndex_AddNew(rsADO As ADODB.Recordset, rsElpKmIndex As typeElpKmIndex)
'---------------------------------------------------------
On Error GoTo Error_Handler

If rsMDB.State = adStateOpen Then rsMDB.Close
rsMDB.Open "select * from ElpKmIndex", cnMDB, , adLockOptimistic
rsADO.AddNew
adoElpKmIndex_AddNew = rsElpKmIndex_PutBuffer(rsADO, rsElpKmIndex)
rsADO.Update

Exit Function

Error_Handler:

adoElpKmIndex_AddNew = Error

End Function
'---------------------------------------------------------
Public Function adoElpKmIndex_Update(rsADO As ADODB.Recordset, rsElpKmIndex As typeElpKmIndex)
'---------------------------------------------------------
Dim xSql As String

On Error GoTo Error_Handler
xSql = "select * from ElpKmIndex" _
    & " where id = '" & rsElpKmIndex.ID & "'" _
    & " and Classe = " & rsElpKmIndex.Classe _
    & " and ElpKMSrc_Id = " & rsElpKmIndex.ElpKMSrc_Id

If rsMDB.State = adStateOpen Then rsMDB.Close

rsMDB.Open xSql, cnMDB, , adLockOptimistic
adoElpKmIndex_Update = rsElpKmIndex_PutBuffer(rsADO, rsElpKmIndex)

rsADO.Update

Exit Function

Error_Handler:

adoElpKmIndex_Update = Error

End Function


