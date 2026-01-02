Attribute VB_Name = "adoZCGSMM30"
Option Explicit

'---------------------------------------------------------
Public Function rsZCGSMM30_PutBuffer(rsado As ADODB.Recordset, rsZCGSMM30 As typeZCGSMM30)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZCGSMM30_PutBuffer = Null

rsado("CGSMM3ETA") = rsZCGSMM30.CGSMM3ETA
rsado("CGSMM3AGE") = rsZCGSMM30.CGSMM3AGE
rsado("CGSMM3SER") = rsZCGSMM30.CGSMM3SER
rsado("CGSMM3SES") = rsZCGSMM30.CGSMM3SES
rsado("CGSMM3OPE") = rsZCGSMM30.CGSMM3OPE
rsado("CGSMM3NAT") = rsZCGSMM30.CGSMM3NAT
rsado("CGSMM3NUM") = rsZCGSMM30.CGSMM3NUM
rsado("CGSMM3SEN") = rsZCGSMM30.CGSMM3SEN
rsado("CGSMM3SEQ") = rsZCGSMM30.CGSMM3SEQ
rsado("CGSMM3DEV") = rsZCGSMM30.CGSMM3DEV
rsado("CGSMM3REF") = rsZCGSMM30.CGSMM3REF
rsado("CGSMM3APP") = rsZCGSMM30.CGSMM3APP
rsado("CGSMM3TAU") = rsZCGSMM30.CGSMM3TAU
rsado("CGSMM3MAR") = rsZCGSMM30.CGSMM3MAR
rsado("CGSMM3MRC") = rsZCGSMM30.CGSMM3MRC
rsado("CGSMM3DVA") = rsZCGSMM30.CGSMM3DVA
rsado("CGSMM3DTR") = rsZCGSMM30.CGSMM3DTR
rsado("CGSMM3DRG") = rsZCGSMM30.CGSMM3DRG
rsado("CGSMM3INT") = rsZCGSMM30.CGSMM3INT
rsado("CGSMM3COU") = rsZCGSMM30.CGSMM3COU
rsado("CGSMM3DEB") = rsZCGSMM30.CGSMM3DEB
rsado("CGSMM3FIN") = rsZCGSMM30.CGSMM3FIN
rsado("CGSMM3ASS") = rsZCGSMM30.CGSMM3ASS
rsado("CGSMM3NBJ") = rsZCGSMM30.CGSMM3NBJ
rsado("CGSMM3NBP") = rsZCGSMM30.CGSMM3NBP
rsado("CGSMM3BAS") = rsZCGSMM30.CGSMM3BAS
rsado("CGSMM3MAC") = rsZCGSMM30.CGSMM3MAC
rsado("CGSMM3MIN") = rsZCGSMM30.CGSMM3MIN
rsado("CGSMM3TXA") = rsZCGSMM30.CGSMM3TXA
Exit Function

Error_Handler:

rsZCGSMM30_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoZCGSMM30_AddNew(rsado As ADODB.Recordset, rsZCGSMM30 As typeZCGSMM30)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoZCGSMM30_AddNew = Null
rsado.AddNew
adoZCGSMM30_AddNew = rsZCGSMM30_PutBuffer(rsado, rsZCGSMM30)
rsado.Update

Exit Function

Error_Handler:

adoZCGSMM30_AddNew = Error

End Function



