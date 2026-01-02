Attribute VB_Name = "adoZMNUUTI0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZMNUUTI0_PutBuffer(rsADO As ADODB.Recordset, rsZMNUUTI0 As typeZMNUUTI0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZMNUUTI0_PutBuffer = Null

rsADO("MNUUTIETB") = rsZMNUUTI0.MNUUTIETB
rsADO("MNUUTIREF") = rsZMNUUTI0.MNUUTIREF
rsADO("MNUUTICUT") = rsZMNUUTI0.MNUUTICUT
rsADO("MNUUTIGR2") = rsZMNUUTI0.MNUUTIGR2
rsADO("MNUUTIGR3") = rsZMNUUTI0.MNUUTIGR3
rsADO("MNUUTIGR4") = rsZMNUUTI0.MNUUTIGR4
rsADO("MNUUTIOUT") = rsZMNUUTI0.MNUUTIOUT
rsADO("MNUUTILAN") = rsZMNUUTI0.MNUUTILAN
rsADO("MNUUTIMSE") = rsZMNUUTI0.MNUUTIMSE
rsADO("MNUUTIAGE") = rsZMNUUTI0.MNUUTIAGE
rsADO("MNUUTISER") = rsZMNUUTI0.MNUUTISER
rsADO("MNUUTISRV") = rsZMNUUTI0.MNUUTISRV
rsADO("MNUUTIGRS") = rsZMNUUTI0.MNUUTIGRS
rsADO("MNUUTIGEN") = rsZMNUUTI0.MNUUTIGEN
rsADO("MNUUTIPOS") = rsZMNUUTI0.MNUUTIPOS
rsADO("MNUUTIMAI") = rsZMNUUTI0.MNUUTIMAI
    
Exit Function

Error_Handler:

rsZMNUUTI0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoZMNUUTI0_AddNew(rsADO As ADODB.Recordset, rsZMNUUTI0 As typeZMNUUTI0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoZMNUUTI0_AddNew = Null
rsADO.AddNew
adoZMNUUTI0_AddNew = rsZMNUUTI0_PutBuffer(rsADO, rsZMNUUTI0)
rsADO.Update

Exit Function

Error_Handler:

adoZMNUUTI0_AddNew = Error

End Function
