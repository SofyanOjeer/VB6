Attribute VB_Name = "adoZRELEVE0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZRELEVE0_PutBuffer(rsADO As ADODB.Recordset, rsZRELEVE0 As typeZRELEVE0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZRELEVE0_PutBuffer = Null

rsADO("RELEVEETA") = rsZRELEVE0.RELEVEETA
rsADO("RELEVEPLA") = rsZRELEVE0.RELEVEPLA
rsADO("RELEVECOM") = rsZRELEVE0.RELEVECOM
rsADO("RELEVEREL") = rsZRELEVE0.RELEVEREL
rsADO("RELEVETYP") = rsZRELEVE0.RELEVETYP
rsADO("RELEVENUM") = rsZRELEVE0.RELEVENUM
rsADO("RELEVEADR") = rsZRELEVE0.RELEVEADR
rsADO("RELEVEGES") = rsZRELEVE0.RELEVEGES
rsADO("RELEVEDER") = rsZRELEVE0.RELEVEDER
rsADO("RELEVEEXT") = rsZRELEVE0.RELEVEEXT

Exit Function

Error_Handler:

rsZRELEVE0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoZRELEVE0_AddNew(rsADO As ADODB.Recordset, rsZRELEVE0 As typeZRELEVE0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoZRELEVE0_AddNew = Null
rsADO.AddNew
adoZRELEVE0_AddNew = rsZRELEVE0_PutBuffer(rsADO, rsZRELEVE0)
rsADO.Update

Exit Function

Error_Handler:

adoZRELEVE0_AddNew = Error

End Function

