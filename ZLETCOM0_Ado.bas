Attribute VB_Name = "adoZLETCOM0"
Option Explicit

'---------------------------------------------------------
Public Function rsZLETCOM0_PutBuffer(rsado As ADODB.Recordset, rsZLETCOM0 As typeZLETCOM0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZLETCOM0_PutBuffer = Null

rsado("LETCOMETA") = rsZLETCOM0.LETCOMETA
rsado("LETCOMPLA") = rsZLETCOM0.LETCOMPLA
rsado("LETCOMCOM") = rsZLETCOM0.LETCOMCOM
rsado("LETCOMAGR") = rsZLETCOM0.LETCOMAGR
rsado("LETCOMSER") = rsZLETCOM0.LETCOMSER
rsado("LETCOMSSR") = rsZLETCOM0.LETCOMSSR
rsado("LETCOMDDE") = rsZLETCOM0.LETCOMDDE
rsado("LETCOMDDR") = rsZLETCOM0.LETCOMDDR
rsado("LETCOMDPR") = rsZLETCOM0.LETCOMDPR
rsado("LETCOMPER") = rsZLETCOM0.LETCOMPER
rsado("LETCOMNBP") = rsZLETCOM0.LETCOMNBP
rsado("LETCOMDTR") = rsZLETCOM0.LETCOMDTR
rsado("LETCOMPIE") = rsZLETCOM0.LETCOMPIE
rsado("LETCOMECR") = rsZLETCOM0.LETCOMECR
rsado("LETCOMOUV") = rsZLETCOM0.LETCOMOUV
rsado("LETCOMCLO") = rsZLETCOM0.LETCOMCLO
rsado("LETCOMDMC") = rsZLETCOM0.LETCOMDMC
rsado("LETCOMMON") = rsZLETCOM0.LETCOMMON
rsado("LETCOMDVA") = rsZLETCOM0.LETCOMDVA
rsado("LETCOMDOP") = rsZLETCOM0.LETCOMDOP
rsado("LETCOMOPE") = rsZLETCOM0.LETCOMOPE
rsado("LETCOMNU1") = rsZLETCOM0.LETCOMNU1
rsado("LETCOMPO1") = rsZLETCOM0.LETCOMPO1
rsado("LETCOMLO1") = rsZLETCOM0.LETCOMLO1
rsado("LETCOMNU2") = rsZLETCOM0.LETCOMNU2
rsado("LETCOMPO2") = rsZLETCOM0.LETCOMPO2
rsado("LETCOMLO2") = rsZLETCOM0.LETCOMLO2
rsado("LETCOMAGO") = rsZLETCOM0.LETCOMAGO
rsado("LETCOMSEO") = rsZLETCOM0.LETCOMSEO
rsado("LETCOMSSO") = rsZLETCOM0.LETCOMSSO
rsado("LETCOMCHE") = rsZLETCOM0.LETCOMCHE
rsado("LETCOMANA") = rsZLETCOM0.LETCOMANA
   
Exit Function

Error_Handler:

rsZLETCOM0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoZLETCOM0_AddNew(rsado As ADODB.Recordset, rsZLETCOM0 As typeZLETCOM0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoZLETCOM0_AddNew = Null
rsado.AddNew
adoZLETCOM0_AddNew = rsZLETCOM0_PutBuffer(rsado, rsZLETCOM0)
rsado.Update

Exit Function

Error_Handler:

adoZLETCOM0_AddNew = Error

End Function


