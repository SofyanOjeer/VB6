Attribute VB_Name = "adoZCLIGRP0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZCLIGRP0_PutBuffer(rsADO As ADODB.Recordset, rsZCLIGRP0 As typeZCLIGRP0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZCLIGRP0_PutBuffer = Null
rsADO("CLIGRPETB") = rsZCLIGRP0.CLIGRPETB
rsADO("CLIGRPCLI") = rsZCLIGRP0.CLIGRPCLI
rsADO("CLIGRPREG") = rsZCLIGRP0.CLIGRPREG
rsADO("CLIGRPREL") = rsZCLIGRP0.CLIGRPREL
rsADO("CLIGRPCOM") = rsZCLIGRP0.CLIGRPCOM
rsADO("CLIGRPAUT") = rsZCLIGRP0.CLIGRPAUT
rsADO("CLIGRPRAT") = rsZCLIGRP0.CLIGRPRAT
rsADO("CLIGRPTAU") = rsZCLIGRP0.CLIGRPTAU
rsADO("CLIGRPPAR") = rsZCLIGRP0.CLIGRPPAR

    
Exit Function

Error_Handler:

rsZCLIGRP0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoZCLIGRP0_AddNew(rsADO As ADODB.Recordset, rsZCLIGRP0 As typeZCLIGRP0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoZCLIGRP0_AddNew = Null
rsADO.AddNew
adoZCLIGRP0_AddNew = rsZCLIGRP0_PutBuffer(rsADO, rsZCLIGRP0)
rsADO.Update

Exit Function

Error_Handler:

adoZCLIGRP0_AddNew = Error

End Function
