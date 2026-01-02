Attribute VB_Name = "srvZAUTHST0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const constZAUTHST0 = "ZAUTHST0"

Type typeZAUTHST0

    AUTHSTETA       As Integer                        ' ETABLISSEMENT
    AUTHSTGPE       As String * 1                     ' GROUPE
    AUTHSTCLI       As String * 7                     ' N CLIENT
    AUTHSTTYP       As String * 1                     ' TYPE AUTO:1,2,3
    AUTHSTAUT       As String * 20                    ' CODE AUTO
    AUTHSTMOD       As Long                           ' DATE MODIFICATION
    AUTHSTSEQ       As Long                           ' SEQUENCE
    AUTHSTEFF       As Long                           ' DATE EFFET
    AUTHSTINT       As Long                           ' DATE ECH. INTER.
    AUTHSTPRO       As String * 3                     ' CODE PROFIL
    AUTHSTDEB       As Long                           ' DATE DEBUT
    AUTHSTFIN       As Long                           ' DATE FIN
    AUTHSTMON       As Long                           ' MONTANT
    AUTHSTBLO       As String * 1                     ' CODE BLOCAGE
    AUTHSTTAU       As Double                         ' TAUX DEPAS
    AUTHSTDUR       As Long                           ' DUREE
    AUTHSTCON       As String * 1                     ' CREDIT CONFIRME
    AUTHSTDEV       As String * 3                     ' DEVISE
    AUTHSTCUT       As Integer                        ' CODE UTILISATEUR
    AUTHSTUCR       As Integer                        ' COD U.CREATION
    AUTHSTUVL       As Integer                        ' COD U.VALIDATION
    AUTHSTUMO       As Integer                        ' COD U.MODIFICAT.
    AUTHSTDCR       As Long                           ' DAT CREATION
    AUTHSTDVL       As Long                           ' DAT VALIDATION
    AUTHSTDMO       As Long                           ' DAT MODIFICATION

End Type

Public Sub srvZAUTHST_Init(recZAUTHST As typeZAUTHST0)
'recZAUTHST.Obj = "ZAUTHST"
'recZAUTHST.Method = ""
'recZAUTHST.Err = ""

recZAUTHST.AUTHSTETA = 0
recZAUTHST.AUTHSTGPE = ""
recZAUTHST.AUTHSTCLI = ""
recZAUTHST.AUTHSTTYP = ""
recZAUTHST.AUTHSTAUT = ""
recZAUTHST.AUTHSTMOD = 0
recZAUTHST.AUTHSTSEQ = 0
recZAUTHST.AUTHSTEFF = 0
recZAUTHST.AUTHSTINT = 0
recZAUTHST.AUTHSTPRO = ""
recZAUTHST.AUTHSTDEB = 0
recZAUTHST.AUTHSTFIN = 0
recZAUTHST.AUTHSTMON = 0
recZAUTHST.AUTHSTBLO = ""
recZAUTHST.AUTHSTTAU = 0
recZAUTHST.AUTHSTDUR = 0
recZAUTHST.AUTHSTCON = ""
recZAUTHST.AUTHSTDEV = ""
recZAUTHST.AUTHSTCUT = 0
recZAUTHST.AUTHSTUCR = 0
recZAUTHST.AUTHSTUVL = 0
recZAUTHST.AUTHSTUMO = 0
recZAUTHST.AUTHSTDCR = 0
recZAUTHST.AUTHSTDVL = 0
recZAUTHST.AUTHSTDMO = 0

End Sub
Public Function srvZAUTHST0_GetBuffer_ODBC(rsADO As ADODB.Recordset, recZAUTHST0 As typeZAUTHST0)
On Error GoTo Error_Handler
srvZAUTHST0_GetBuffer_ODBC = Null

recZAUTHST0.AUTHSTETA = rsADO("AUTHSTETA")
recZAUTHST0.AUTHSTGPE = rsADO("AUTHSTGPE")
recZAUTHST0.AUTHSTCLI = rsADO("AUTHSTCLI")
recZAUTHST0.AUTHSTTYP = rsADO("AUTHSTTYP")
recZAUTHST0.AUTHSTAUT = rsADO("AUTHSTAUT")
recZAUTHST0.AUTHSTMOD = rsADO("AUTHSTMOD")
recZAUTHST0.AUTHSTSEQ = rsADO("AUTHSTSEQ")
recZAUTHST0.AUTHSTEFF = rsADO("AUTHSTEFF")
recZAUTHST0.AUTHSTINT = rsADO("AUTHSTINT")
recZAUTHST0.AUTHSTPRO = rsADO("AUTHSTPRO")
recZAUTHST0.AUTHSTDEB = rsADO("AUTHSTDEB")
recZAUTHST0.AUTHSTFIN = rsADO("AUTHSTFIN")
recZAUTHST0.AUTHSTMON = rsADO("AUTHSTMON")
recZAUTHST0.AUTHSTBLO = rsADO("AUTHSTBLO")
recZAUTHST0.AUTHSTTAU = rsADO("AUTHSTTAU")
recZAUTHST0.AUTHSTDUR = rsADO("AUTHSTDUR")
recZAUTHST0.AUTHSTCON = rsADO("AUTHSTCON")
recZAUTHST0.AUTHSTDEV = rsADO("AUTHSTDEV")
recZAUTHST0.AUTHSTCUT = rsADO("AUTHSTCUT")
recZAUTHST0.AUTHSTUCR = rsADO("AUTHSTUCR")
recZAUTHST0.AUTHSTUVL = rsADO("AUTHSTUVL")
recZAUTHST0.AUTHSTUMO = rsADO("AUTHSTUMO")
recZAUTHST0.AUTHSTDCR = rsADO("AUTHSTDCR")
recZAUTHST0.AUTHSTDVL = rsADO("AUTHSTDVL")
recZAUTHST0.AUTHSTDMO = rsADO("AUTHSTDMO")

Exit Function
Error_Handler:
srvZAUTHST0_GetBuffer_ODBC = Error
End Function
Public Sub srvYAUTHST0_ElpDisplay(recZAUTHST0 As typeZAUTHST0)

frmElpDisplay.fgData.Rows = 26

frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTHSTETA    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recZAUTHST0.AUTHSTETA
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTHSTGPE    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "GROUPE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recZAUTHST0.AUTHSTGPE
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTHSTCLI    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N CLIENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recZAUTHST0.AUTHSTCLI
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTHSTTYP    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TYPE AUTO:1,2,3"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recZAUTHST0.AUTHSTTYP
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTHSTAUT   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE AUTO"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recZAUTHST0.AUTHSTAUT
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTHSTMOD    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE MODIFICATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recZAUTHST0.AUTHSTMOD
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTHSTSEQ    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SEQUENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recZAUTHST0.AUTHSTSEQ
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTHSTEFF    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE EFFET"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recZAUTHST0.AUTHSTEFF
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTHSTINT    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE ECH. INTER."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recZAUTHST0.AUTHSTINT
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTHSTPRO    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE PROFIL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recZAUTHST0.AUTHSTPRO
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTHSTDEB    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE DEBUT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recZAUTHST0.AUTHSTDEB
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTHSTFIN    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE FIN"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recZAUTHST0.AUTHSTFIN
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTHSTMON   15P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recZAUTHST0.AUTHSTMON
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTHSTBLO    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE BLOCAGE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recZAUTHST0.AUTHSTBLO
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTHSTTAU  6.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TAUX DEPAS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recZAUTHST0.AUTHSTTAU
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTHSTDUR    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DUREE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recZAUTHST0.AUTHSTDUR
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTHSTCON    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CREDIT CONFIRME"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recZAUTHST0.AUTHSTCON
frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTHSTDEV    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DEVISE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recZAUTHST0.AUTHSTDEV
frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTHSTCUT    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE UTILISATEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recZAUTHST0.AUTHSTCUT
frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTHSTUCR    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COD U.CREATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recZAUTHST0.AUTHSTUCR
frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTHSTUVL    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COD U.VALIDATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recZAUTHST0.AUTHSTUVL
frmElpDisplay.fgData.Row = 22
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTHSTUMO    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COD U.MODIFICAT."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recZAUTHST0.AUTHSTUMO
frmElpDisplay.fgData.Row = 23
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTHSTDCR    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DAT CREATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recZAUTHST0.AUTHSTDCR
frmElpDisplay.fgData.Row = 24
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTHSTDVL    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DAT VALIDATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recZAUTHST0.AUTHSTDVL
frmElpDisplay.fgData.Row = 25
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTHSTDMO    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DAT MODIFICATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recZAUTHST0.AUTHSTDMO

frmElpDisplay.Show vbModal
End Sub


