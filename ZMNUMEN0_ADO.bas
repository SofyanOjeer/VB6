Attribute VB_Name = "adoZMNUMEN0"
'---------------------------------------------------------
Option Explicit

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZMNUMEN0_PutBuffer(rsADO As ADODB.Recordset, rsZMNUMEN0 As typeZMNUMEN0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZMNUMEN0_PutBuffer = Null
rsADO("MNUMENETB") = rsZMNUMEN0.MNUMENETB
rsADO("MNUMENREF") = rsZMNUMEN0.MNUMENREF
rsADO("MNUMENGRP") = rsZMNUMEN0.MNUMENGRP
rsADO("MNUMENPRE") = rsZMNUMEN0.MNUMENPRE
rsADO("MNUMENORD") = rsZMNUMEN0.MNUMENORD
rsADO("MNUMENCOD") = rsZMNUMEN0.MNUMENCOD
rsADO("MNUMENOIA") = rsZMNUMEN0.MNUMENOIA
rsADO("MNUMENJOQ") = rsZMNUMEN0.MNUMENJOQ

    
Exit Function

Error_Handler:

rsZMNUMEN0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoZMNUMEN0_AddNew(rsADO As ADODB.Recordset, rsZMNUMEN0 As typeZMNUMEN0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoZMNUMEN0_AddNew = Null
rsADO.AddNew
adoZMNUMEN0_AddNew = rsZMNUMEN0_PutBuffer(rsADO, rsZMNUMEN0)
rsADO.Update

Exit Function

Error_Handler:

adoZMNUMEN0_AddNew = Error

End Function
