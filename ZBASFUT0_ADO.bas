Attribute VB_Name = "adoZBASFUT0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZBASFUT0_PutBuffer(rsADO As ADODB.Recordset, rsZBASFUT0 As typeZBASFUT0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZBASFUT0_PutBuffer = Null

rsADO("BASFUTETA") = rsZBASFUT0.BASFUTETA
rsADO("BASFUTOPE") = rsZBASFUT0.BASFUTOPE
rsADO("BASFUTAGE") = rsZBASFUT0.BASFUTAGE
rsADO("BASFUTSER") = rsZBASFUT0.BASFUTSER
rsADO("BASFUTSSE") = rsZBASFUT0.BASFUTSSE
rsADO("BASFUTDOS") = rsZBASFUT0.BASFUTDOS
rsADO("BASFUTDTE") = rsZBASFUT0.BASFUTDTE
rsADO("BASFUTEVE") = rsZBASFUT0.BASFUTEVE
rsADO("BASFUTNUM") = rsZBASFUT0.BASFUTNUM
rsADO("BASFUTTYP") = rsZBASFUT0.BASFUTTYP
rsADO("BASFUTNAT") = rsZBASFUT0.BASFUTNAT
rsADO("BASFUTDVA") = rsZBASFUT0.BASFUTDVA
rsADO("BASFUTMON") = rsZBASFUT0.BASFUTMON
rsADO("BASFUTSEN") = rsZBASFUT0.BASFUTSEN
rsADO("BASFUTDEV") = rsZBASFUT0.BASFUTDEV
rsADO("BASFUTCPT") = rsZBASFUT0.BASFUTCPT
rsADO("BASFUTTCL") = rsZBASFUT0.BASFUTTCL
rsADO("BASFUTCLI") = rsZBASFUT0.BASFUTCLI
rsADO("BASFUTTAU") = rsZBASFUT0.BASFUTTAU
rsADO("BASFUTNAG") = rsZBASFUT0.BASFUTNAG
rsADO("BASFUTNSE") = rsZBASFUT0.BASFUTNSE
rsADO("BASFUTNSS") = rsZBASFUT0.BASFUTNSS
rsADO("BASFUTNDO") = rsZBASFUT0.BASFUTNDO
rsADO("BASFUTLIB") = rsZBASFUT0.BASFUTLIB
Exit Function

Error_Handler:

rsZBASFUT0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoZBASFUT0_AddNew(rsADO As ADODB.Recordset, rsZBASFUT0 As typeZBASFUT0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoZBASFUT0_AddNew = Null
rsADO.AddNew
adoZBASFUT0_AddNew = rsZBASFUT0_PutBuffer(rsADO, rsZBASFUT0)
rsADO.Update

Exit Function

Error_Handler:

adoZBASFUT0_AddNew = Error

End Function

