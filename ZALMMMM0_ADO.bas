Attribute VB_Name = "adoZALMMMM0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZALMMM0_PutBuffer(rsado As ADODB.Recordset, rsZALMMM0 As typeZALMMM0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZALMMM0_PutBuffer = Null
rsado("ALMMMREC") = rsZALMMM0.ALMMMREC
rsado("ALMMMDAT") = rsZALMMM0.ALMMMDAT
rsado("ALMMMNBR") = rsZALMMM0.ALMMMNBR
    
Exit Function

Error_Handler:

rsZALMMM0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoZALMMM0_AddNew(rsado As ADODB.Recordset, rsZALMMM0 As typeZALMMM0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoZALMMM0_AddNew = Null
rsado.AddNew
adoZALMMM0_AddNew = rsZALMMM0_PutBuffer(rsado, rsZALMMM0)
rsado.Update

Exit Function

Error_Handler:

adoZALMMM0_AddNew = Error

End Function

