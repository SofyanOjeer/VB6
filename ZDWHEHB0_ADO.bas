Attribute VB_Name = "adoZDWHEHB0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZDWHEHB0_PutBuffer(rsADO As ADODB.Recordset, rsZDWHEHB0 As typeZDWHEHB0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZDWHEHB0_PutBuffer = Null
rsADO("DWHEHBDTX") = rsZDWHEHB0.DWHEHBDTX
rsADO("DWHEHBETA") = rsZDWHEHB0.DWHEHBETA
rsADO("DWHEHBAGE") = rsZDWHEHB0.DWHEHBAGE
rsADO("DWHEHBSER") = rsZDWHEHB0.DWHEHBSER
rsADO("DWHEHBSSE") = rsZDWHEHB0.DWHEHBSSE
rsADO("DWHEHBOPE") = rsZDWHEHB0.DWHEHBOPE
rsADO("DWHEHBNAT") = rsZDWHEHB0.DWHEHBNAT
rsADO("DWHEHBNDO") = rsZDWHEHB0.DWHEHBNDO
rsADO("DWHEHBPOO") = rsZDWHEHB0.DWHEHBPOO
rsADO("DWHEHBPOU") = rsZDWHEHB0.DWHEHBPOU
rsADO("DWHEHBMBE") = rsZDWHEHB0.DWHEHBMBE
rsADO("DWHEHBMNE") = rsZDWHEHB0.DWHEHBMNE
rsADO("DWHEHBNUM") = rsZDWHEHB0.DWHEHBNUM
rsADO("DWHEHBAUT") = rsZDWHEHB0.DWHEHBAUT
rsADO("DWHEHBRUB") = rsZDWHEHB0.DWHEHBRUB
rsADO("DWHEHBOBJ") = rsZDWHEHB0.DWHEHBOBJ
rsADO("DWHEHBDSY") = rsZDWHEHB0.DWHEHBDSY
    
Exit Function

Error_Handler:

rsZDWHEHB0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoZDWHEHB0_AddNew(rsADO As ADODB.Recordset, rsZDWHEHB0 As typeZDWHEHB0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoZDWHEHB0_AddNew = Null
rsADO.AddNew
adoZDWHEHB0_AddNew = rsZDWHEHB0_PutBuffer(rsADO, rsZDWHEHB0)
rsADO.Update

Exit Function

Error_Handler:

adoZDWHEHB0_AddNew = Error

End Function



