Attribute VB_Name = "adoZCGSMM10"
Option Explicit

'---------------------------------------------------------
Public Function rsZCGSMM10_PutBuffer(rsado As ADODB.Recordset, rsZCGSMM10 As typeZCGSMM10)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZCGSMM10_PutBuffer = Null

rsado("CGSMM1ETA") = rsZCGSMM10.CGSMM1ETA
rsado("CGSMM1AGE") = rsZCGSMM10.CGSMM1AGE
rsado("CGSMM1SER") = rsZCGSMM10.CGSMM1SER
rsado("CGSMM1SES") = rsZCGSMM10.CGSMM1SES
rsado("CGSMM1OPE") = rsZCGSMM10.CGSMM1OPE
rsado("CGSMM1NAT") = rsZCGSMM10.CGSMM1NAT
rsado("CGSMM1NUM") = rsZCGSMM10.CGSMM1NUM
rsado("CGSMM1MON") = rsZCGSMM10.CGSMM1MON
rsado("CGSMM1NBR") = rsZCGSMM10.CGSMM1NBR
rsado("CGSMM1DEV") = rsZCGSMM10.CGSMM1DEV
rsado("CGSMM1CLI") = rsZCGSMM10.CGSMM1CLI
rsado("CGSMM1COM") = rsZCGSMM10.CGSMM1COM
rsado("CGSMM1ENG") = rsZCGSMM10.CGSMM1ENG
rsado("CGSMM1DEB") = rsZCGSMM10.CGSMM1DEB
rsado("CGSMM1FIN") = rsZCGSMM10.CGSMM1FIN
rsado("CGSMM1DUR") = rsZCGSMM10.CGSMM1DUR
rsado("CGSMM1TYP") = rsZCGSMM10.CGSMM1TYP
rsado("CGSMM1AUT") = rsZCGSMM10.CGSMM1AUT
rsado("CGSMM1CVL") = rsZCGSMM10.CGSMM1CVL
rsado("CGSMM1NLO") = rsZCGSMM10.CGSMM1NLO
Exit Function

Error_Handler:

rsZCGSMM10_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoZCGSMM10_AddNew(rsado As ADODB.Recordset, rsZCGSMM10 As typeZCGSMM10)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoZCGSMM10_AddNew = Null
rsado.AddNew
adoZCGSMM10_AddNew = rsZCGSMM10_PutBuffer(rsado, rsZCGSMM10)
rsado.Update

Exit Function

Error_Handler:

adoZCGSMM10_AddNew = Error

End Function




