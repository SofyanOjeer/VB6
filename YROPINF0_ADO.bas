Attribute VB_Name = "adoYROPINF0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsYROPINF0_PutBuffer(rsADO As ADODB.Recordset, rsYROPINF0 As typeYROPINF0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYROPINF0_PutBuffer = Null
rsADO("ROPINFID") = rsYROPINF0.ROPINFID
rsADO("ROPINFIDP") = rsYROPINF0.ROPINFIDP
rsADO("ROPINFIDT") = rsYROPINF0.ROPINFIDT
rsADO("ROPINFIDT2") = rsYROPINF0.ROPINFIDT2
rsADO("ROPINFIDTL") = rsYROPINF0.ROPINFIDTL
rsADO("ROPINFMAIL") = rsYROPINF0.ROPINFMAIL
rsADO("ROPINFSTA") = rsYROPINF0.ROPINFSTA
rsADO("ROPINFSTAK") = rsYROPINF0.ROPINFSTAK
rsADO("ROPINFUUSR") = rsYROPINF0.ROPINFUUSR
rsADO("ROPINFUAMJ") = rsYROPINF0.ROPINFUAMJ
rsADO("ROPINFUHMS") = rsYROPINF0.ROPINFUHMS
rsADO("ROPINFUVER") = rsYROPINF0.ROPINFUVER
rsADO("ROPINFGUO") = rsYROPINF0.ROPINFGUO
rsADO("ROPINFGECH") = rsYROPINF0.ROPINFGECH
rsADO("ROPINFGUSR") = rsYROPINF0.ROPINFGUSR
rsADO("ROPINFGSRV") = rsYROPINF0.ROPINFGSRV
rsADO("ROPINFGNAT") = rsYROPINF0.ROPINFGNAT
rsADO("ROPINFGPRV") = rsYROPINF0.ROPINFGPRV
rsADO("ROPINFGTXT") = rsYROPINF0.ROPINFGTXT
    
Exit Function

Error_Handler:

rsYROPINF0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoYROPINF0_AddNew(rsADO As ADODB.Recordset, rsYROPINF0 As typeYROPINF0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoYROPINF0_AddNew = Null
rsADO.AddNew
adoYROPINF0_AddNew = rsYROPINF0_PutBuffer(rsADO, rsYROPINF0)
rsADO.Update

Exit Function

Error_Handler:

adoYROPINF0_AddNew = Error

End Function





