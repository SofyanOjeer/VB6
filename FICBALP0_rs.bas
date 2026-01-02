Attribute VB_Name = "rsFICBALP0"
'---------------------------------------------------------
Option Explicit
Type typeFICBALP0

    COMPTEDEV       As String * 3                      '
    COMPTEOBL       As String * 10                    '
    Classe          As String * 1                    '
    BIL_HBL         As String * 3                    '
    COMPTECOM       As String * 20                  '
    COMPTEINT       As String * 32                  '
    SOLDE_W         As Currency                        '
    SOLDECVL      As Currency                        '


End Type

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsFICBALP0_GetBuffer(rsSab As ADODB.Recordset, rsFICBALP0 As typeFICBALP0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsFICBALP0_GetBuffer = Null

rsFICBALP0.COMPTEDEV = rsSab("COMPTEDEV")
rsFICBALP0.COMPTEOBL = rsSab("COMPTEOBL")
rsFICBALP0.Classe = rsSab("CLASSE")
rsFICBALP0.BIL_HBL = rsSab("BIL_HBL")
rsFICBALP0.COMPTECOM = rsSab("COMPTECOM")
rsFICBALP0.COMPTEINT = rsSab("COMPTEINT")
rsFICBALP0.SOLDE_W = rsSab("SOLDE_W")
rsFICBALP0.SOLDECVL = rsSab("SOLDECVL")

Exit Function

Error_Handler:

rsFICBALP0_GetBuffer = Error

End Function

'---------------------------------------------------------
Public Sub rsFICBALP0_Init(rsFICBALP0 As typeFICBALP0)
'---------------------------------------------------------
rsFICBALP0.COMPTEDEV = ""
rsFICBALP0.COMPTEOBL = ""
rsFICBALP0.Classe = ""
rsFICBALP0.BIL_HBL = ""
rsFICBALP0.COMPTECOM = ""
rsFICBALP0.COMPTEINT = ""
rsFICBALP0.SOLDE_W = 0
rsFICBALP0.SOLDECVL = 0

End Sub


'










