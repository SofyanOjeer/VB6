Attribute VB_Name = "adoZADRESS0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZADRESS0_PutBuffer(rsADO As ADODB.Recordset, rsZADRESS0 As typeZADRESS0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZADRESS0_PutBuffer = Null
rsADO("ADRESSETA") = rsZADRESS0.ADRESSETA
rsADO("ADRESSTYP") = rsZADRESS0.ADRESSTYP
rsADO("ADRESSPLA") = rsZADRESS0.ADRESSPLA
rsADO("ADRESSNUM") = rsZADRESS0.ADRESSNUM
rsADO("ADRESSCOA") = rsZADRESS0.ADRESSCOA
rsADO("ADRESSDLI") = rsZADRESS0.ADRESSDLI
rsADO("ADRESSDDE") = rsZADRESS0.ADRESSDDE
rsADO("ADRESSRA1") = rsZADRESS0.ADRESSRA1
rsADO("ADRESSRA2") = rsZADRESS0.ADRESSRA2
rsADO("ADRESSAD1") = rsZADRESS0.ADRESSAD1
rsADO("ADRESSAD2") = rsZADRESS0.ADRESSAD2
rsADO("ADRESSAD3") = rsZADRESS0.ADRESSAD3
rsADO("ADRESSCOP") = rsZADRESS0.ADRESSCOP
rsADO("ADRESSVIL") = rsZADRESS0.ADRESSVIL
rsADO("ADRESSPAY") = rsZADRESS0.ADRESSPAY
rsADO("ADRESSTEL") = rsZADRESS0.ADRESSTEL
rsADO("ADRESSFAX") = rsZADRESS0.ADRESSFAX
rsADO("ADRESSTEX") = rsZADRESS0.ADRESSTEX
    
Exit Function

Error_Handler:

rsZADRESS0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoZADRESS0_AddNew(rsADO As ADODB.Recordset, rsZADRESS0 As typeZADRESS0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoZADRESS0_AddNew = Null
rsADO.AddNew
adoZADRESS0_AddNew = rsZADRESS0_PutBuffer(rsADO, rsZADRESS0)
rsADO.Update

Exit Function

Error_Handler:

adoZADRESS0_AddNew = Error

End Function
