Attribute VB_Name = "adoZMNUUTP0"
Option Explicit

'---------------------------------------------------------
Public Function adoZMNUUTP0_AddNew(rsADO As ADODB.Recordset, rsZMNUUTP0 As typeZMNUUTP0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoZMNUUTP0_AddNew = Null
rsADO.AddNew
adoZMNUUTP0_AddNew = rsZMNUUTP0_PutBuffer(rsADO, rsZMNUUTP0)
rsADO.Update

Exit Function

Error_Handler:

adoZMNUUTP0_AddNew = Error

End Function

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZMNUUTP0_PutBuffer(rsADO As ADODB.Recordset, rsZMNUUTP0 As typeZMNUUTP0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZMNUUTP0_PutBuffer = Null
rsADO("MNUUTPETB") = rsZMNUUTP0.MNUUTPETB
rsADO("MNUUTPREF") = rsZMNUUTP0.MNUUTPREF
rsADO("MNUUTPGRP") = rsZMNUUTP0.MNUUTPGRP
rsADO("MNUUTPAGE") = rsZMNUUTP0.MNUUTPAGE
rsADO("MNUUTPOIA") = rsZMNUUTP0.MNUUTPOIA
rsADO("MNUUTPCLA") = rsZMNUUTP0.MNUUTPCLA

    
Exit Function

Error_Handler:

rsZMNUUTP0_PutBuffer = Error

End Function



