Attribute VB_Name = "rsZMNUUTP0"
Option Explicit

Type typeZMNUUTP0
    
    MNUUTPETB       As Integer                        ' ETABLISSEMENT
    MNUUTPREF       As Long                           ' REFERENCE LOT
    MNUUTPGRP       As String * 10                    ' GROUPE MENU
    MNUUTPAGE       As Integer                        ' agence
    MNUUTPOIA       As String * 1                     ' INTER-AGENCE
    MNUUTPCLA       As String * 99                    ' CLASSE MAJ /INT
    
End Type


'---------------------------------------------------------
Public Function rsZMNUUTP0_GetBuffer(rsAdo As ADODB.Recordset, rsZMNUUTP0 As typeZMNUUTP0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZMNUUTP0_GetBuffer = Null

rsZMNUUTP0.MNUUTPETB = rsAdo("MNUUTPETB")
rsZMNUUTP0.MNUUTPREF = rsAdo("MNUUTPREF")
rsZMNUUTP0.MNUUTPGRP = rsAdo("MNUUTPGRP")
rsZMNUUTP0.MNUUTPAGE = rsAdo("MNUUTPAGE")
rsZMNUUTP0.MNUUTPOIA = rsAdo("MNUUTPOIA")
rsZMNUUTP0.MNUUTPCLA = rsAdo("MNUUTPCLA")


Exit Function

Error_Handler:

rsZMNUUTP0_GetBuffer = Error

End Function


