Attribute VB_Name = "adoZSOLDE0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZSOLDE0_PutBuffer(rsADO As ADODB.Recordset, rsZSOLDE0 As typeZSOLDE0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZSOLDE0_PutBuffer = Null

rsADO("SOLDEETA") = rsZSOLDE0.SOLDEETA
rsADO("SOLDEPLA") = rsZSOLDE0.SOLDEPLA
rsADO("SOLDECOM") = rsZSOLDE0.SOLDECOM
rsADO("SOLDEDMO") = rsZSOLDE0.SOLDEDMO
rsADO("SOLDEDAN") = rsZSOLDE0.SOLDEDAN
rsADO("SOLDECEN") = rsZSOLDE0.SOLDECEN
rsADO("SOLDECAN") = rsZSOLDE0.SOLDECAN
rsADO("SOLDEC01") = rsZSOLDE0.SOLDEC01
rsADO("SOLDEC02") = rsZSOLDE0.SOLDEC02
rsADO("SOLDEC03") = rsZSOLDE0.SOLDEC03
rsADO("SOLDEC04") = rsZSOLDE0.SOLDEC04
rsADO("SOLDEC05") = rsZSOLDE0.SOLDEC05
rsADO("SOLDEC06") = rsZSOLDE0.SOLDEC06
rsADO("SOLDEC07") = rsZSOLDE0.SOLDEC07
rsADO("SOLDEC08") = rsZSOLDE0.SOLDEC08
rsADO("SOLDEC09") = rsZSOLDE0.SOLDEC09
rsADO("SOLDEC10") = rsZSOLDE0.SOLDEC10
rsADO("SOLDEC11") = rsZSOLDE0.SOLDEC11
rsADO("SOLDEC12") = rsZSOLDE0.SOLDEC12
rsADO("SOLDEVEN") = rsZSOLDE0.SOLDEVEN
rsADO("SOLDEVAN") = rsZSOLDE0.SOLDEVAN
rsADO("SOLDEV01") = rsZSOLDE0.SOLDEV01
rsADO("SOLDEV02") = rsZSOLDE0.SOLDEV02
rsADO("SOLDEV03") = rsZSOLDE0.SOLDEV03
rsADO("SOLDEV04") = rsZSOLDE0.SOLDEV04
rsADO("SOLDEV05") = rsZSOLDE0.SOLDEV05
rsADO("SOLDEV06") = rsZSOLDE0.SOLDEV06
rsADO("SOLDEV07") = rsZSOLDE0.SOLDEV07
rsADO("SOLDEV08") = rsZSOLDE0.SOLDEV08
rsADO("SOLDEV09") = rsZSOLDE0.SOLDEV09
rsADO("SOLDEV10") = rsZSOLDE0.SOLDEV10
rsADO("SOLDEV11") = rsZSOLDE0.SOLDEV11
rsADO("SOLDEV12") = rsZSOLDE0.SOLDEV12

    
Exit Function

Error_Handler:

rsZSOLDE0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoZSOLDE0_AddNew(rsADO As ADODB.Recordset, rsZSOLDE0 As typeZSOLDE0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoZSOLDE0_AddNew = Null
rsADO.AddNew
adoZSOLDE0_AddNew = rsZSOLDE0_PutBuffer(rsADO, rsZSOLDE0)
rsADO.Update

Exit Function

Error_Handler:

adoZSOLDE0_AddNew = Error

End Function
