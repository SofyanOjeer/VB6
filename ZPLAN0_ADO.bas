Attribute VB_Name = "adoZPLAN0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZPLAN0_PutBuffer(rsADO As ADODB.Recordset, rsZPLAN0 As typeZPLAN0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZPLAN0_PutBuffer = Null

 
rsADO("PLANETABL") = rsZPLAN0.PLANETABL
rsADO("PLANPLAN") = rsZPLAN0.PLANPLAN
rsADO("PLANCOOBL") = rsZPLAN0.PLANCOOBL
rsADO("PLANINTIT") = rsZPLAN0.PLANINTIT
rsADO("PLANCOPRO") = rsZPLAN0.PLANCOPRO
rsADO("PLANCLASS") = rsZPLAN0.PLANCLASS
rsADO("PLANFONCT") = rsZPLAN0.PLANFONCT
rsADO("PLANSESOL") = rsZPLAN0.PLANSESOL
rsADO("PLANGEDEP") = rsZPLAN0.PLANGEDEP
rsADO("PLANTIERS") = rsZPLAN0.PLANTIERS
rsADO("PLANFICOB") = rsZPLAN0.PLANFICOB
rsADO("PLANCARAC") = rsZPLAN0.PLANCARAC
rsADO("PLANPESTO") = rsZPLAN0.PLANPESTO
rsADO("PLANNBPER") = rsZPLAN0.PLANNBPER
rsADO("PLANNBMOU") = rsZPLAN0.PLANNBMOU
rsADO("PLANINEXT") = rsZPLAN0.PLANINEXT
rsADO("PLANPROGR") = rsZPLAN0.PLANPROGR
   
Exit Function

Error_Handler:

rsZPLAN0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoZPLAN0_AddNew(rsADO As ADODB.Recordset, rsZPLAN0 As typeZPLAN0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoZPLAN0_AddNew = Null
rsADO.AddNew
adoZPLAN0_AddNew = rsZPLAN0_PutBuffer(rsADO, rsZPLAN0)
rsADO.Update

Exit Function

Error_Handler:

adoZPLAN0_AddNew = Error

End Function
