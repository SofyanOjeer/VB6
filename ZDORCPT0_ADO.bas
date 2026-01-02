Attribute VB_Name = "adoZDORCPT0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZDORCPT0_PutBuffer(rsado As ADODB.Recordset, rsZDORCPT0 As typeZDORCPT0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZDORCPT0_PutBuffer = Null

rsado("DORCPTETA") = rsZDORCPT0.DORCPTETA
rsado("DORCPTPLA") = rsZDORCPT0.DORCPTPLA
rsado("DORCPTCOM") = rsZDORCPT0.DORCPTCOM
rsado("DORCPTDOR") = rsZDORCPT0.DORCPTDOR
rsado("DORCPTDDO") = rsZDORCPT0.DORCPTDDO
rsado("DORCPTDMV") = rsZDORCPT0.DORCPTDMV
rsado("DORCPTDDE") = rsZDORCPT0.DORCPTDDE
rsado("DORCPTDPR") = rsZDORCPT0.DORCPTDPR
rsado("DORCPTCOD") = rsZDORCPT0.DORCPTCOD
rsado("DORCPTDMO") = rsZDORCPT0.DORCPTDMO
rsado("DORCPTDRE") = rsZDORCPT0.DORCPTDRE
rsado("DORCPTMAJ") = rsZDORCPT0.DORCPTMAJ
Exit Function

Error_Handler:

rsZDORCPT0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoZDORCPT0_AddNew(rsado As ADODB.Recordset, rsZDORCPT0 As typeZDORCPT0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoZDORCPT0_AddNew = Null
rsado.AddNew
adoZDORCPT0_AddNew = rsZDORCPT0_PutBuffer(rsado, rsZDORCPT0)
rsado.Update

Exit Function

Error_Handler:

adoZDORCPT0_AddNew = Error

End Function

