Attribute VB_Name = "adoZCOMPTE0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZCOMPTE0_PutBuffer(rsADO As ADODB.Recordset, rsZCOMPTE0 As typeZCOMPTE0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZCOMPTE0_PutBuffer = Null
rsADO("COMPTEETA") = rsZCOMPTE0.COMPTEETA
rsADO("COMPTEPLA") = rsZCOMPTE0.COMPTEPLA
rsADO("COMPTECOM") = rsZCOMPTE0.COMPTECOM
rsADO("COMPTEOBL") = rsZCOMPTE0.COMPTEOBL
rsADO("COMPTEINT") = rsZCOMPTE0.COMPTEINT
rsADO("COMPTEAGE") = rsZCOMPTE0.COMPTEAGE
rsADO("COMPTEDEV") = rsZCOMPTE0.COMPTEDEV
rsADO("COMPTEOUV") = rsZCOMPTE0.COMPTEOUV
rsADO("COMPTECLO") = rsZCOMPTE0.COMPTECLO
rsADO("COMPTELOR") = rsZCOMPTE0.COMPTELOR
rsADO("COMPTESUC") = rsZCOMPTE0.COMPTESUC
rsADO("COMPTECLA") = rsZCOMPTE0.COMPTECLA
rsADO("COMPTEFON") = rsZCOMPTE0.COMPTEFON
rsADO("COMPTEBLO") = rsZCOMPTE0.COMPTEBLO
rsADO("COMPTEMOT") = rsZCOMPTE0.COMPTEMOT
rsADO("COMPTESEN") = rsZCOMPTE0.COMPTESEN
rsADO("COMPTEMOD") = rsZCOMPTE0.COMPTEMOD

    
Exit Function

Error_Handler:

rsZCOMPTE0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoZCOMPTE0_AddNew(rsADO As ADODB.Recordset, rsZCOMPTE0 As typeZCOMPTE0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoZCOMPTE0_AddNew = Null
rsADO.AddNew
adoZCOMPTE0_AddNew = rsZCOMPTE0_PutBuffer(rsADO, rsZCOMPTE0)
rsADO.Update

Exit Function

Error_Handler:

adoZCOMPTE0_AddNew = Error

End Function
