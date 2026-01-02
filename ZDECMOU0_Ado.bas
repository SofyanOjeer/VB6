Attribute VB_Name = "adoZDECMOU0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZDECMOU0_PutBuffer(rsADO As ADODB.Recordset, rsZDECMOU0 As typeZDECMOU0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZDECMOU0_PutBuffer = Null

rsADO("DECMOUETA") = rsZDECMOU0.DECMOUETA
rsADO("DECMOUCOM") = rsZDECMOU0.DECMOUCOM
rsADO("DECMOUDTR") = rsZDECMOU0.DECMOUDTR
rsADO("DECMOUAGE") = rsZDECMOU0.DECMOUAGE
rsADO("DECMOUSER") = rsZDECMOU0.DECMOUSER
rsADO("DECMOUSSE") = rsZDECMOU0.DECMOUSSE
rsADO("DECMOUCOP") = rsZDECMOU0.DECMOUCOP
rsADO("DECMOUNOP") = rsZDECMOU0.DECMOUNOP
rsADO("DECMOUDRE") = rsZDECMOU0.DECMOUDRE
rsADO("DECMOUDLR") = rsZDECMOU0.DECMOUDLR
rsADO("DECMOUUIN") = rsZDECMOU0.DECMOUUIN
rsADO("DECMOUDCR") = rsZDECMOU0.DECMOUDCR
rsADO("DECMOUDUT") = rsZDECMOU0.DECMOUDUT
rsADO("DECMOUUTI") = rsZDECMOU0.DECMOUUTI
rsADO("DECMOUREA") = rsZDECMOU0.DECMOUREA
rsADO("DECMOUNSQ") = rsZDECMOU0.DECMOUNSQ
rsADO("DECMOUFUT") = rsZDECMOU0.DECMOUFUT

rsADO("DECMOUORI") = rsZDECMOU0.DECMOUORI
rsADO("DECMOUNAT") = rsZDECMOU0.DECMOUNAT
rsADO("DECMOUMRE") = rsZDECMOU0.DECMOUMRE
rsADO("DECMOUREQ") = rsZDECMOU0.DECMOUREQ
rsADO("DECMOUAPS") = rsZDECMOU0.DECMOUAPS
rsADO("DECMOUMOS") = rsZDECMOU0.DECMOUMOS
rsADO("DECMOUFIL") = rsZDECMOU0.DECMOUFIL
Exit Function

Error_Handler:

rsZDECMOU0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoZDECMOU0_AddNew(rsADO As ADODB.Recordset, rsZDECMOU0 As typeZDECMOU0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoZDECMOU0_AddNew = Null
rsADO.AddNew
adoZDECMOU0_AddNew = rsZDECMOU0_PutBuffer(rsADO, rsZDECMOU0)
rsADO.Update

Exit Function

Error_Handler:

adoZDECMOU0_AddNew = Error

End Function



