Attribute VB_Name = "adoZAUTHST0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZAUTHST0_PutBuffer(rsADO As ADODB.Recordset, rsZAUTHST0 As typeZAUTHST0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZAUTHST0_PutBuffer = Null

rsADO("AUTHSTETA") = rsZAUTHST0.AUTHSTETA
rsADO("AUTHSTGPE") = rsZAUTHST0.AUTHSTGPE
rsADO("AUTHSTCLI") = rsZAUTHST0.AUTHSTCLI
rsADO("AUTHSTTYP") = rsZAUTHST0.AUTHSTTYP
rsADO("AUTHSTAUT") = rsZAUTHST0.AUTHSTAUT
rsADO("AUTHSTMOD") = rsZAUTHST0.AUTHSTMOD
rsADO("AUTHSTSEQ") = rsZAUTHST0.AUTHSTSEQ
rsADO("AUTHSTEFF") = rsZAUTHST0.AUTHSTEFF
rsADO("AUTHSTINT") = rsZAUTHST0.AUTHSTINT
rsADO("AUTHSTPRO") = rsZAUTHST0.AUTHSTPRO
rsADO("AUTHSTDEB") = rsZAUTHST0.AUTHSTDEB
rsADO("AUTHSTFIN") = rsZAUTHST0.AUTHSTFIN
rsADO("AUTHSTMON") = rsZAUTHST0.AUTHSTMON
rsADO("AUTHSTBLO") = rsZAUTHST0.AUTHSTBLO
rsADO("AUTHSTTAU") = rsZAUTHST0.AUTHSTTAU
rsADO("AUTHSTDUR") = rsZAUTHST0.AUTHSTDUR
rsADO("AUTHSTCON") = rsZAUTHST0.AUTHSTCON
rsADO("AUTHSTDEV") = rsZAUTHST0.AUTHSTDEV
rsADO("AUTHSTCUT") = rsZAUTHST0.AUTHSTCUT
rsADO("AUTHSTUCR") = rsZAUTHST0.AUTHSTUCR
rsADO("AUTHSTUVL") = rsZAUTHST0.AUTHSTUVL
rsADO("AUTHSTUMO") = rsZAUTHST0.AUTHSTUMO
rsADO("AUTHSTDCR") = rsZAUTHST0.AUTHSTDCR
rsADO("AUTHSTDVL") = rsZAUTHST0.AUTHSTDVL
rsADO("AUTHSTDMO") = rsZAUTHST0.AUTHSTDMO
Exit Function

Error_Handler:

rsZAUTHST0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoZAUTHST0_AddNew(rsADO As ADODB.Recordset, rsZAUTHST0 As typeZAUTHST0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoZAUTHST0_AddNew = Null
rsADO.AddNew
adoZAUTHST0_AddNew = rsZAUTHST0_PutBuffer(rsADO, rsZAUTHST0)
rsADO.Update

Exit Function

Error_Handler:

adoZAUTHST0_AddNew = Error

End Function

