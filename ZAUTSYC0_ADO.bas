Attribute VB_Name = "adoZAUTSYC0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZAUTSYC0_PutBuffer(rsADO As ADODB.Recordset, rsZAUTSYC0 As typeZAUTSYC0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZAUTSYC0_PutBuffer = Null

rsADO("AUTSYCETA") = rsZAUTSYC0.AUTSYCETA
rsADO("AUTSYCGPE") = rsZAUTSYC0.AUTSYCGPE
rsADO("AUTSYCCLI") = rsZAUTSYC0.AUTSYCCLI
rsADO("AUTSYCADR") = rsZAUTSYC0.AUTSYCADR
rsADO("AUTSYCTYP") = rsZAUTSYC0.AUTSYCTYP
rsADO("AUTSYCAUT") = rsZAUTSYC0.AUTSYCAUT
rsADO("AUTSYCPER") = rsZAUTSYC0.AUTSYCPER
rsADO("AUTSYCSUI") = rsZAUTSYC0.AUTSYCSUI
rsADO("AUTSYCELM") = rsZAUTSYC0.AUTSYCELM
rsADO("AUTSYCNIV") = rsZAUTSYC0.AUTSYCNIV
rsADO("AUTSYCINT") = rsZAUTSYC0.AUTSYCINT
rsADO("AUTSYCEFF") = rsZAUTSYC0.AUTSYCEFF
rsADO("AUTSYCPRO") = rsZAUTSYC0.AUTSYCPRO
rsADO("AUTSYCDEB") = rsZAUTSYC0.AUTSYCDEB
rsADO("AUTSYCFIN") = rsZAUTSYC0.AUTSYCFIN
rsADO("AUTSYCMON") = rsZAUTSYC0.AUTSYCMON
rsADO("AUTSYCDEV") = rsZAUTSYC0.AUTSYCDEV
rsADO("AUTSYCBLO") = rsZAUTSYC0.AUTSYCBLO
rsADO("AUTSYCAMO") = rsZAUTSYC0.AUTSYCAMO
rsADO("AUTSYCGRP") = rsZAUTSYC0.AUTSYCGRP
rsADO("AUTSYCRES") = rsZAUTSYC0.AUTSYCRES
rsADO("AUTSYCTAU") = rsZAUTSYC0.AUTSYCTAU
rsADO("AUTSYCDUR") = rsZAUTSYC0.AUTSYCDUR
rsADO("AUTSYCCON") = rsZAUTSYC0.AUTSYCCON
rsADO("AUTSYCCET") = rsZAUTSYC0.AUTSYCCET
rsADO("AUTSYCCUT") = rsZAUTSYC0.AUTSYCCUT
rsADO("AUTSYCUCR") = rsZAUTSYC0.AUTSYCUCR
rsADO("AUTSYCUVL") = rsZAUTSYC0.AUTSYCUVL
rsADO("AUTSYCUMO") = rsZAUTSYC0.AUTSYCUMO
rsADO("AUTSYCDCR") = rsZAUTSYC0.AUTSYCDCR
rsADO("AUTSYCDVL") = rsZAUTSYC0.AUTSYCDVL
rsADO("AUTSYCDMO") = rsZAUTSYC0.AUTSYCDMO

Exit Function

Error_Handler:

rsZAUTSYC0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoZAUTSYC0_AddNew(rsADO As ADODB.Recordset, rsZAUTSYC0 As typeZAUTSYC0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoZAUTSYC0_AddNew = Null
rsADO.AddNew
adoZAUTSYC0_AddNew = rsZAUTSYC0_PutBuffer(rsADO, rsZAUTSYC0)
rsADO.Update

Exit Function

Error_Handler:

adoZAUTSYC0_AddNew = Error

End Function

