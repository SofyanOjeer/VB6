Attribute VB_Name = "adoWorldCheck"
Option Explicit

'---------------------------------------------------------
Public Function rsWC_Data_PutBuffer(rsADO As ADODB.Recordset, rsWC_Data As typeWC_Data)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsWC_Data_PutBuffer = Null

rsADO("WC_Id") = rsWC_Data.WC_Id
rsADO("WC_UpdD") = rsWC_Data.WC_UpdD
rsADO("WC_UpdH") = rsWC_Data.WC_UpdH
rsADO("WC_Sta") = rsWC_Data.WC_Sta
rsADO("WC_LastName") = rsWC_Data.WC_LastName
rsADO("WC_FirstName") = rsWC_Data.WC_FirstName
rsADO("WC_Memo") = rsWC_Data.WC_Memo


Exit Function

Error_Handler:

rsWC_Data_PutBuffer = Error
End Function

'---------------------------------------------------------
Public Function adoWC_Data_AddNew(rsADO As ADODB.Recordset, rsWC_Data As typeWC_Data)
'---------------------------------------------------------
On Error GoTo Error_Handler

If rsWC.State = adStateOpen Then rsWC.Close
rsWC.Open "select * from WC_Data", cnWC, , adLockOptimistic
rsADO.AddNew
adoWC_Data_AddNew = rsWC_Data_PutBuffer(rsADO, rsWC_Data)
rsADO.Update

Exit Function

Error_Handler:

adoWC_Data_AddNew = Error

End Function

'---------------------------------------------------------
Public Function adoWC_Data_Update(rsADO As ADODB.Recordset, rsWC_Data As typeWC_Data)
'---------------------------------------------------------
Dim xSql As String

On Error GoTo Error_Handler
xSql = "select * from WC_Data where WC_ID = " & rsWC_Data.WC_Id

If rsWC.State = adStateOpen Then rsWC.Close

rsWC.Open xSql, cnWC, , adLockOptimistic
adoWC_Data_Update = rsWC_Data_PutBuffer(rsADO, rsWC_Data)

rsADO.Update

Exit Function

Error_Handler:

adoWC_Data_Update = Error

End Function

Public Function adoWC_Data_Delete(rsADO As ADODB.Recordset, rsWC_Data As typeWC_Data)
'---------------------------------------------------------
Dim xSql As String

On Error GoTo Error_Handler
adoWC_Data_Delete = Null

xSql = "delete * from WC_Data where WC_ID = " & rsWC_Data.WC_Id
Call FEU_ROUGE
Set rsADO = cnWC.Execute(xSql)
Call FEU_VERT
Exit Function

Error_Handler:

adoWC_Data_Delete = Error

End Function



