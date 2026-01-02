Attribute VB_Name = "adoYROPDOS0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsYROPDOS0_PutBuffer(rsADO As ADODB.Recordset, rsYROPDOS0 As typeYROPDOS0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYROPDOS0_PutBuffer = Null
rsADO("ROPDOSID") = rsYROPDOS0.ROPDOSID
rsADO("ROPDOSSTA") = rsYROPDOS0.ROPDOSSTA
rsADO("ROPDOSSTAK") = rsYROPDOS0.ROPDOSSTAK
rsADO("ROPDOSUUSR") = rsYROPDOS0.ROPDOSUUSR
rsADO("ROPDOSUAMJ") = rsYROPDOS0.ROPDOSUAMJ
rsADO("ROPDOSUHMS") = rsYROPDOS0.ROPDOSUHMS
rsADO("ROPDOSUVER") = rsYROPDOS0.ROPDOSUVER
rsADO("ROPDOSGECH") = rsYROPDOS0.ROPDOSGECH
rsADO("ROPDOSGUSR") = rsYROPDOS0.ROPDOSGUSR
rsADO("ROPDOSGSRV") = rsYROPDOS0.ROPDOSGSRV
rsADO("ROPDOSGNAT") = rsYROPDOS0.ROPDOSGNAT
rsADO("ROPDOSGPRV") = rsYROPDOS0.ROPDOSGPRV
rsADO("ROPDOSGGRA") = rsYROPDOS0.ROPDOSGGRA
rsADO("ROPDOSGPRI") = rsYROPDOS0.ROPDOSGPRI
rsADO("ROPDOSGCOU") = rsYROPDOS0.ROPDOSGCOU
rsADO("ROPDOSIAMJ") = rsYROPDOS0.ROPDOSIAMJ
rsADO("ROPDOSISRV") = rsYROPDOS0.ROPDOSISRV
rsADO("ROPDOSIUSR") = rsYROPDOS0.ROPDOSIUSR
rsADO("ROPDOSIREF") = rsYROPDOS0.ROPDOSIREF
rsADO("ROPDOSXDOM") = rsYROPDOS0.ROPDOSXDOM
rsADO("ROPDOSXAPP") = rsYROPDOS0.ROPDOSXAPP
rsADO("ROPDOSXID") = rsYROPDOS0.ROPDOSXID
rsADO("ROPDOSQUAL") = rsYROPDOS0.ROPDOSQUAL
    
Exit Function

Error_Handler:

rsYROPDOS0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoYROPDOS0_AddNew(rsADO As ADODB.Recordset, rsYROPDOS0 As typeYROPDOS0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoYROPDOS0_AddNew = Null
rsADO.AddNew
adoYROPDOS0_AddNew = rsYROPDOS0_PutBuffer(rsADO, rsYROPDOS0)
rsADO.Update

Exit Function

Error_Handler:

adoYROPDOS0_AddNew = Error

End Function




