Attribute VB_Name = "adoFICBALP0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsFICBALP0_PutBuffer(rsado As ADODB.Recordset, rsFICBALP0 As typeFICBALP0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsFICBALP0_PutBuffer = Null
rsado("COMPTEDEV") = rsFICBALP0.COMPTEDEV

rsado("COMPTEOBL") = rsFICBALP0.COMPTEOBL
rsado("CLASSE") = rsFICBALP0.Classe
rsado("BIL_HBL") = rsFICBALP0.BIL_HBL
rsado("COMPTECOM") = rsFICBALP0.COMPTECOM
rsado("COMPTEINT") = rsFICBALP0.COMPTEINT
rsado("SOLDE_W") = rsFICBALP0.SOLDE_W
rsado("SOLDECVL") = rsFICBALP0.SOLDECVL
    
Exit Function

Error_Handler:

rsFICBALP0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoFICBALP0_AddNew(rsado As ADODB.Recordset, rsFICBALP0 As typeFICBALP0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoFICBALP0_AddNew = Null
rsado.AddNew
adoFICBALP0_AddNew = rsFICBALP0_PutBuffer(rsado, rsFICBALP0)
rsado.Update

Exit Function

Error_Handler:

adoFICBALP0_AddNew = Error

End Function


