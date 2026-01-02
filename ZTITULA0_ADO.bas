Attribute VB_Name = "adoZTITULA0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZTITULA0_PutBuffer(rsADO As ADODB.Recordset, rsZTITULA0 As typeZTITULA0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZTITULA0_PutBuffer = Null
rsADO("TITULAETA") = rsZTITULA0.TITULAETA
rsADO("TITULAPLA") = rsZTITULA0.TITULAPLA
rsADO("TITULACOM") = rsZTITULA0.TITULACOM
rsADO("TITULACLI") = rsZTITULA0.TITULACLI
rsADO("TITULAPRI") = rsZTITULA0.TITULAPRI
rsADO("TITULATPR") = rsZTITULA0.TITULATPR

    
Exit Function

Error_Handler:

rsZTITULA0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoZTITULA0_AddNew(rsADO As ADODB.Recordset, rsZTITULA0 As typeZTITULA0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoZTITULA0_AddNew = Null
rsADO.AddNew
adoZTITULA0_AddNew = rsZTITULA0_PutBuffer(rsADO, rsZTITULA0)
rsADO.Update

Exit Function

Error_Handler:

adoZTITULA0_AddNew = Error

End Function
