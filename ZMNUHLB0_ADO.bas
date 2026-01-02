Attribute VB_Name = "adoZMNUHLB0"
Option Explicit

'---------------------------------------------------------
Public Function rsZMNUHLB0_PutBuffer(rsADO As ADODB.Recordset, rsZMNUHLB0 As typeZMNUHLB0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZMNUHLB0_PutBuffer = Null
rsADO("MNUHLBETB") = rsZMNUHLB0.MNUHLBETB
rsADO("MNUHLBREF") = rsZMNUHLB0.MNUHLBREF
rsADO("MNUHLBCLA") = rsZMNUHLB0.MNUHLBCLA
rsADO("MNUHLBNOM") = rsZMNUHLB0.MNUHLBNOM
rsADO("MNUHLBVAL") = rsZMNUHLB0.MNUHLBVAL
rsADO("MNUHLBDBD") = rsZMNUHLB0.MNUHLBDBD
rsADO("MNUHLBDBH") = rsZMNUHLB0.MNUHLBDBH
rsADO("MNUHLBSUS") = rsZMNUHLB0.MNUHLBSUS
rsADO("MNUHLBFID") = rsZMNUHLB0.MNUHLBFID
rsADO("MNUHLBFIH") = rsZMNUHLB0.MNUHLBFIH
rsADO("MNUHLBSDT") = rsZMNUHLB0.MNUHLBSDT
rsADO("MNUHLBSHE") = rsZMNUHLB0.MNUHLBSHE

    
Exit Function

Error_Handler:

rsZMNUHLB0_PutBuffer = Error

End Function
'---------------------------------------------------------
Public Function adoZMNUHLB0_AddNew(rsADO As ADODB.Recordset, rsZMNUHLB0 As typeZMNUHLB0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoZMNUHLB0_AddNew = Null
rsADO.AddNew
adoZMNUHLB0_AddNew = rsZMNUHLB0_PutBuffer(rsADO, rsZMNUHLB0)
rsADO.Update

Exit Function

Error_Handler:

adoZMNUHLB0_AddNew = Error

End Function






