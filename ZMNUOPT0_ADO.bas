Attribute VB_Name = "adoZMNUOPT0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZMNUOPT0_PutBuffer(rsADO As ADODB.Recordset, rsZMNUOPT0 As typeZMNUOPT0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZMNUOPT0_PutBuffer = Null

rsADO("MNUOPTCOD") = rsZMNUOPT0.MNUOPTCOD
rsADO("MNUOPTCLI") = rsZMNUOPT0.MNUOPTCLI
rsADO("MNUOPTLIB") = rsZMNUOPT0.MNUOPTLIB
rsADO("MNUOPTENS") = rsZMNUOPT0.MNUOPTENS
rsADO("MNUOPTENT") = rsZMNUOPT0.MNUOPTENT
rsADO("MNUOPTSTR") = rsZMNUOPT0.MNUOPTSTR
rsADO("MNUOPTARE") = rsZMNUOPT0.MNUOPTARE
rsADO("MNUOPTBAT") = rsZMNUOPT0.MNUOPTBAT
rsADO("MNUOPTVAL") = rsZMNUOPT0.MNUOPTVAL
rsADO("MNUOPTSUP") = rsZMNUOPT0.MNUOPTSUP
rsADO("MNUOPTOIA") = rsZMNUOPT0.MNUOPTOIA
rsADO("MNUOPTGES") = rsZMNUOPT0.MNUOPTGES

    
Exit Function

Error_Handler:

rsZMNUOPT0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoZMNUOPT0_AddNew(rsADO As ADODB.Recordset, rsZMNUOPT0 As typeZMNUOPT0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoZMNUOPT0_AddNew = Null
rsADO.AddNew
adoZMNUOPT0_AddNew = rsZMNUOPT0_PutBuffer(rsADO, rsZMNUOPT0)
rsADO.Update

Exit Function

Error_Handler:

adoZMNUOPT0_AddNew = Error

End Function
