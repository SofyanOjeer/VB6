Attribute VB_Name = "adoSideEUPLAB0"
Option Explicit

'---------------------------------------------------------
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsSideEUPLAB0_PutBuffer(rsADO As ADODB.Recordset, rsSideEUPLAB0 As typeSideEUPLAB0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsSideEUPLAB0_PutBuffer = Null

rsADO("EUPLABID") = rsSideEUPLAB0.EUPLABID
rsADO("EUPLABBICE") = rsSideEUPLAB0.EUPLABBICE
rsADO("EUPLABNOME") = rsSideEUPLAB0.EUPLABNOME
rsADO("EUPLABLIB") = rsSideEUPLAB0.EUPLABLIB
rsADO("EUPLABMONT") = rsSideEUPLAB0.EUPLABMONT
rsADO("EUPLABDEVI") = rsSideEUPLAB0.EUPLABDEVI
rsADO("EUPLABSTAI") = rsSideEUPLAB0.EUPLABSTAI
rsADO("EUPLABSTAS1") = rsSideEUPLAB0.EUPLABSTAS1
rsADO("EUPLABSTAS2") = rsSideEUPLAB0.EUPLABSTAS2
rsADO("EUPLABSTAS3") = rsSideEUPLAB0.EUPLABSTAS3
rsADO("EUPLABSTAS4") = rsSideEUPLAB0.EUPLABSTAS4
rsADO("EUPLABSTAS5") = rsSideEUPLAB0.EUPLABSTAS5
rsADO("EUPLABSTAS6") = rsSideEUPLAB0.EUPLABSTAS6
rsADO("EUPLABSTAS7") = rsSideEUPLAB0.EUPLABSTAS7
rsADO("EUPLABSTAS8") = rsSideEUPLAB0.EUPLABSTAS8
   
Exit Function

Error_Handler:

rsSideEUPLAB0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoSideEUPLAB0_AddNew(rsADO As ADODB.Recordset, rsSideEUPLAB0 As typeSideEUPLAB0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoSideEUPLAB0_AddNew = Null
rsADO.AddNew
adoSideEUPLAB0_AddNew = rsSideEUPLAB0_PutBuffer(rsADO, rsSideEUPLAB0)
rsADO.Update

Exit Function

Error_Handler:

adoSideEUPLAB0_AddNew = Error

End Function



