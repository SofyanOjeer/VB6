Attribute VB_Name = "adoElpKmInfo"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsElpKmInfo_PutBuffer(rsADO As ADODB.Recordset, rsElpKmInfo As typeElpKmInfo)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsElpKmInfo_PutBuffer = Null

rsADO("ElpKMSrc_Id") = rsElpKmInfo.ElpKMSrc_Id
rsADO("ID") = rsElpKmInfo.ID
rsADO("Description") = rsElpKmInfo.Description
rsADO("Pass") = rsElpKmInfo.Pass
rsADO("Memo") = rsElpKmInfo.Memo

    
Exit Function

Error_Handler:

rsElpKmInfo_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoElpKmInfo_AddNew(rsADO As ADODB.Recordset, rsElpKmInfo As typeElpKmInfo)
'---------------------------------------------------------
On Error GoTo Error_Handler

If rsMDB.State = adStateOpen Then rsMDB.Close
rsMDB.Open "select * from ElpKmInfo", cnMDB, , adLockOptimistic
rsADO.AddNew
adoElpKmInfo_AddNew = rsElpKmInfo_PutBuffer(rsADO, rsElpKmInfo)
rsADO.Update

Exit Function

Error_Handler:

adoElpKmInfo_AddNew = Error

End Function
