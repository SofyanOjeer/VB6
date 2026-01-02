Attribute VB_Name = "adoZMNURUT0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZMNURUT0_PutBuffer(rsADO As ADODB.Recordset, rsZMNURUT0 As typeZMNURUT0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZMNURUT0_PutBuffer = Null

rsADO("MNURUTUTI") = rsZMNURUT0.MNURUTUTI
rsADO("MNURUTNOM") = rsZMNURUT0.MNURUTNOM
rsADO("MNURUTETB") = rsZMNURUT0.MNURUTETB
rsADO("MNURUTCUT") = rsZMNURUT0.MNURUTCUT
rsADO("MNURUTLOG") = rsZMNURUT0.MNURUTLOG

    
Exit Function

Error_Handler:

rsZMNURUT0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoZMNURUT0_AddNew(rsADO As ADODB.Recordset, rsZMNURUT0 As typeZMNURUT0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoZMNURUT0_AddNew = Null
rsADO.AddNew
adoZMNURUT0_AddNew = rsZMNURUT0_PutBuffer(rsADO, rsZMNURUT0)
rsADO.Update

Exit Function

Error_Handler:

adoZMNURUT0_AddNew = Error

End Function
