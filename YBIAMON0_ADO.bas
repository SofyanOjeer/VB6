Attribute VB_Name = "adoYBIAMON0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsYBIAMON0_PutBuffer(rsado As ADODB.Recordset, rsYBIAMON0 As typeYBIAMON0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYBIAMON0_PutBuffer = Null

rsado("MONAPP") = rsYBIAMON0.MONAPP
rsado("MONFLUX") = rsYBIAMON0.MONFLUX
rsado("MONSTATUS") = rsYBIAMON0.MONSTATUS
rsado("MONNUM") = rsYBIAMON0.MONNUM
rsado("MONJOB") = rsYBIAMON0.MONJOB
rsado("MONPGM") = rsYBIAMON0.MONPGM
rsado("MONUSR") = rsYBIAMON0.MONUSR
rsado("MONAMJ") = rsYBIAMON0.MONAMJ
rsado("MONHMS") = rsYBIAMON0.MONHMS
rsado("MONFILE") = rsYBIAMON0.MONFILE

    
Exit Function

Error_Handler:

rsYBIAMON0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoYBIAMON0_AddNew(rsado As ADODB.Recordset, rsYBIAMON0 As typeYBIAMON0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoYBIAMON0_AddNew = Null
rsado.AddNew
adoYBIAMON0_AddNew = rsYBIAMON0_PutBuffer(rsado, rsYBIAMON0)
rsado.Update

Exit Function

Error_Handler:

adoYBIAMON0_AddNew = Error

End Function


