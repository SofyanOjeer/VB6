Attribute VB_Name = "adoYEUPMON0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsYEUPMON0_PutBuffer(rsADO As ADODB.Recordset, rsYEUPMON0 As typeYEUPMON0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYEUPMON0_PutBuffer = Null

rsADO("EUPG2AOPE") = rsYEUPMON0.EUPG2AOPE
rsADO("EUPG2ANUM") = rsYEUPMON0.EUPG2ANUM
rsADO("EUPG2ACRE") = rsYEUPMON0.EUPG2ACRE
rsADO("EUPG2ANEC") = rsYEUPMON0.EUPG2ANEC
rsADO("EUPMONID") = rsYEUPMON0.EUPMONID
rsADO("EUPMONSTA") = rsYEUPMON0.EUPMONSTA
rsADO("EUPMONDMO") = rsYEUPMON0.EUPMONDMO
rsADO("EUPMONHMO") = rsYEUPMON0.EUPMONHMO
rsADO("EUPMONDSW") = rsYEUPMON0.EUPMONDSW
rsADO("EUPMONHSW") = rsYEUPMON0.EUPMONHSW
rsADO("EUPMONTIC") = rsYEUPMON0.EUPMONTIC
rsADO("EUPMONDID") = rsYEUPMON0.EUPMONDID
rsADO("EUPMONBIC") = rsYEUPMON0.EUPMONBIC
rsADO("EUPMONNOM") = rsYEUPMON0.EUPMONNOM
rsADO("EUPMONLIB") = rsYEUPMON0.EUPMONLIB
rsADO("EUPMONMON") = rsYEUPMON0.EUPMONMON
rsADO("EUPMONDEV") = rsYEUPMON0.EUPMONDEV
rsADO("EUPMONECH") = rsYEUPMON0.EUPMONECH
rsADO("EUPMONPRI") = rsYEUPMON0.EUPMONPRI

    
Exit Function

Error_Handler:

rsYEUPMON0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoYEUPMON0_AddNew(rsADO As ADODB.Recordset, rsYEUPMON0 As typeYEUPMON0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoYEUPMON0_AddNew = Null
rsADO.AddNew
adoYEUPMON0_AddNew = rsYEUPMON0_PutBuffer(rsADO, rsYEUPMON0)
rsADO.Update

Exit Function

Error_Handler:

adoYEUPMON0_AddNew = Error

End Function



