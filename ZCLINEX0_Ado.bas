Attribute VB_Name = "adoZCLINEX0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZCLINEX0_PutBuffer(rsADO As ADODB.Recordset, rsZCLINEX0 As typeZCLINEX0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZCLINEX0_PutBuffer = Null
rsADO("CLINEXETB") = rsZCLINEX0.CLINEXETB
rsADO("CLINEXCLI") = rsZCLINEX0.CLINEXCLI
rsADO("CLINEXORG") = rsZCLINEX0.CLINEXORG
rsADO("CLINEXDNO") = rsZCLINEX0.CLINEXDNO
rsADO("CLINEXDCR") = rsZCLINEX0.CLINEXDCR
rsADO("CLINEXDRE") = rsZCLINEX0.CLINEXDRE
rsADO("CLINEXNO1") = rsZCLINEX0.CLINEXNO1
rsADO("CLINEXNO2") = rsZCLINEX0.CLINEXNO2
rsADO("CLINEXDSA") = rsZCLINEX0.CLINEXDSA
rsADO("CLINEXUSR") = rsZCLINEX0.CLINEXUSR

    
Exit Function

Error_Handler:

rsZCLINEX0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoZCLINEX0_AddNew(rsADO As ADODB.Recordset, rsZCLINEX0 As typeZCLINEX0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoZCLINEX0_AddNew = Null
rsADO.AddNew
adoZCLINEX0_AddNew = rsZCLINEX0_PutBuffer(rsADO, rsZCLINEX0)
rsADO.Update

Exit Function

Error_Handler:

adoZCLINEX0_AddNew = Error

End Function

