Attribute VB_Name = "adoYCHQMON0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsYCHQMON0_PutBuffer(rsADO As ADODB.Recordset, rsYCHQMON0 As typeYCHQMON0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYCHQMON0_PutBuffer = Null

rsADO("CHQRC1ETA") = rsYCHQMON0.CHQRC1ETA
rsADO("CHQRC1AGE") = rsYCHQMON0.CHQRC1AGE
rsADO("CHQRC1SER") = rsYCHQMON0.CHQRC1SER
rsADO("CHQRC1SSE") = rsYCHQMON0.CHQRC1SSE
rsADO("CHQRC1OPE") = rsYCHQMON0.CHQRC1OPE
rsADO("CHQRC1DOS") = rsYCHQMON0.CHQRC1DOS
rsADO("CHQRC1DCR") = rsYCHQMON0.CHQRC1DCR
rsADO("CHQDATE") = rsYCHQMON0.CHQDATE
rsADO("CHQCOMPTE") = rsYCHQMON0.CHQCOMPTE
rsADO("CHQCREM") = rsYCHQMON0.CHQCREM
rsADO("CHQDEVISE") = rsYCHQMON0.CHQDEVISE
rsADO("CHQMONTANT") = rsYCHQMON0.CHQMONTANT
rsADO("CHQNB") = rsYCHQMON0.CHQNB
rsADO("CHQMONSTA") = rsYCHQMON0.CHQMONSTA
rsADO("CHQMONUPDS") = rsYCHQMON0.CHQMONUPDS

    
Exit Function

Error_Handler:

rsYCHQMON0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoYCHQMON0_AddNew(rsADO As ADODB.Recordset, rsYCHQMON0 As typeYCHQMON0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoYCHQMON0_AddNew = Null
rsADO.AddNew
adoYCHQMON0_AddNew = rsYCHQMON0_PutBuffer(rsADO, rsYCHQMON0)
rsADO.Update

Exit Function

Error_Handler:

adoYCHQMON0_AddNew = Error

End Function


