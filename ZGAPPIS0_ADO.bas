Attribute VB_Name = "adoZGAPPIS0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZGAPPIS0_PutBuffer(rsado As ADODB.Recordset, rsZGAPPIS0 As typeZGAPPIS0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZGAPPIS0_PutBuffer = Null

rsado("GAPPISTAB") = rsZGAPPIS0.GAPPISTAB
rsado("GAPPISECH") = rsZGAPPIS0.GAPPISECH
rsado("GAPPISCLA") = rsZGAPPIS0.GAPPISCLA
rsado("GAPPISETA") = rsZGAPPIS0.GAPPISETA
rsado("GAPPISAGE") = rsZGAPPIS0.GAPPISAGE
rsado("GAPPISSER") = rsZGAPPIS0.GAPPISSER
rsado("GAPPISSSE") = rsZGAPPIS0.GAPPISSSE
rsado("GAPPISOPE") = rsZGAPPIS0.GAPPISOPE
rsado("GAPPISNAT") = rsZGAPPIS0.GAPPISNAT
rsado("GAPPISNUO") = rsZGAPPIS0.GAPPISNUO
rsado("GAPPISDEV") = rsZGAPPIS0.GAPPISDEV
rsado("GAPPISSEN") = rsZGAPPIS0.GAPPISSEN
rsado("GAPPISDEC") = rsZGAPPIS0.GAPPISDEC
rsado("GAPPISRUB") = rsZGAPPIS0.GAPPISRUB
rsado("GAPPISTPR") = rsZGAPPIS0.GAPPISTPR
rsado("GAPPISCLI") = rsZGAPPIS0.GAPPISCLI
rsado("GAPPISMON") = rsZGAPPIS0.GAPPISMON
rsado("GAPPISTTI") = rsZGAPPIS0.GAPPISTTI
rsado("GAPPISTTE") = rsZGAPPIS0.GAPPISTTE
rsado("GAPPISRTV") = rsZGAPPIS0.GAPPISRTV
rsado("GAPPISTAU") = rsZGAPPIS0.GAPPISTAU
rsado("GAPPISSOL") = rsZGAPPIS0.GAPPISSOL
rsado("GAPPISPOU") = rsZGAPPIS0.GAPPISPOU
rsado("GAPPISSIG") = rsZGAPPIS0.GAPPISSIG
rsado("GAPPISVIL") = rsZGAPPIS0.GAPPISVIL
   
Exit Function

Error_Handler:

rsZGAPPIS0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoZGAPPIS0_AddNew(rsado As ADODB.Recordset, rsZGAPPIS0 As typeZGAPPIS0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoZGAPPIS0_AddNew = Null
rsado.AddNew
adoZGAPPIS0_AddNew = rsZGAPPIS0_PutBuffer(rsado, rsZGAPPIS0)
rsado.Update

Exit Function

Error_Handler:

adoZGAPPIS0_AddNew = Error

End Function

