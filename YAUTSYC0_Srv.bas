Attribute VB_Name = "srvYAUTSYC0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const constYAUTSYC0 = "YAUTSYC0"
Type typeYAUTSYC0
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    AUTSYCETA       As Integer                        ' ETABLISSEMENT
    AUTSYCGPE       As String * 1                     ' GROUPE
    AUTSYCCLI       As String * 7                     ' N CLIENT
    AUTSYCADR       As Long                           ' ADRESSE
    AUTSYCTYP       As String * 1                     ' TYPE AUTO:1,2,3
    AUTSYCAUT       As String * 20                    ' CODE AUTO
    AUTSYCPER       As Long                           ' CODE PERE
    AUTSYCSUI       As Long                           ' ADRESSE SUIVANTE
    AUTSYCELM       As String * 1                     ' ELEMENTAIRE
    AUTSYCNIV       As Long                           ' NIVEAU
    AUTSYCINT       As Long                           ' DATE ECH. INTER.
    AUTSYCEFF       As Long                           ' DATE EFFET
    AUTSYCPRO       As String * 3                     ' CODE PROFIL
    AUTSYCDEB       As Long                           ' DATE DEBUT
    AUTSYCFIN       As Long                           ' DATE FIN
    AUTSYCMON       As Long                           ' MONTANT
    AUTSYCDEV       As String * 3                     ' DEVISE
    AUTSYCBLO       As String * 1                     ' CODE BLOCAGE
    AUTSYCAMO       As String * 1                     ' AMORTISSABLE
    AUTSYCGRP       As String * 7                     ' CODE GROUPE
    AUTSYCRES       As String * 3                     ' CODE RESPONSABLE
    AUTSYCTAU       As Double                         ' TAUX DEPAS
    AUTSYCDUR       As Long                           ' DUREE
    AUTSYCCON       As String * 1                     ' CREDIT CONFIRME
    AUTSYCCET       As Long                           ' CODE ETAT
    AUTSYCCUT       As Integer                        ' CODE UTILISATEUR
    AUTSYCUCR       As Integer                        ' COD U.CREATION
    AUTSYCUVL       As Integer                        ' COD U.VALIDATION
    AUTSYCUMO       As Integer                        ' COD U.MODIFICAT.
    AUTSYCDCR       As Long                           ' DAT CREATION
    AUTSYCDVL       As Long                           ' DAT VALIDATION
    AUTSYCDMO       As Long                           ' DAT MODIFICATION

End Type
Public Sub srvYAUTSYC0_Init(recYAUTSYC0 As typeYAUTSYC0)
recYAUTSYC0.Obj = "YAUTSYC0"
recYAUTSYC0.Method = ""
recYAUTSYC0.Err = ""
recYAUTSYC0.AUTSYCETA = 0
recYAUTSYC0.AUTSYCGPE = ""
recYAUTSYC0.AUTSYCCLI = ""
recYAUTSYC0.AUTSYCADR = 0
recYAUTSYC0.AUTSYCTYP = ""
recYAUTSYC0.AUTSYCAUT = ""
recYAUTSYC0.AUTSYCPER = 0
recYAUTSYC0.AUTSYCSUI = 0
recYAUTSYC0.AUTSYCELM = ""
recYAUTSYC0.AUTSYCNIV = 0
recYAUTSYC0.AUTSYCINT = 0
recYAUTSYC0.AUTSYCEFF = 0
recYAUTSYC0.AUTSYCPRO = ""
recYAUTSYC0.AUTSYCDEB = 0
recYAUTSYC0.AUTSYCFIN = 0
recYAUTSYC0.AUTSYCMON = 0
recYAUTSYC0.AUTSYCDEV = ""
recYAUTSYC0.AUTSYCBLO = ""
recYAUTSYC0.AUTSYCAMO = ""
recYAUTSYC0.AUTSYCGRP = ""
recYAUTSYC0.AUTSYCRES = ""
recYAUTSYC0.AUTSYCTAU = 0
recYAUTSYC0.AUTSYCDUR = 0
recYAUTSYC0.AUTSYCCON = ""
recYAUTSYC0.AUTSYCCET = 0
recYAUTSYC0.AUTSYCCUT = 0
recYAUTSYC0.AUTSYCUCR = 0
recYAUTSYC0.AUTSYCUVL = 0
recYAUTSYC0.AUTSYCUMO = 0
recYAUTSYC0.AUTSYCDCR = 0
recYAUTSYC0.AUTSYCDVL = 0
recYAUTSYC0.AUTSYCDMO = 0
End Sub
Public Function srvYAUTSYC0_GetBuffer_ODBC(rsADO As ADODB.Recordset, recYAUTSYC0 As typeYAUTSYC0)
On Error GoTo Error_Handler
srvYAUTSYC0_GetBuffer_ODBC = Null
recYAUTSYC0.AUTSYCETA = rsADO("AUTSYCETA")
recYAUTSYC0.AUTSYCGPE = rsADO("AUTSYCGPE")
recYAUTSYC0.AUTSYCCLI = rsADO("AUTSYCCLI")
recYAUTSYC0.AUTSYCADR = rsADO("AUTSYCADR")
recYAUTSYC0.AUTSYCTYP = rsADO("AUTSYCTYP")
recYAUTSYC0.AUTSYCAUT = rsADO("AUTSYCAUT")
recYAUTSYC0.AUTSYCPER = rsADO("AUTSYCPER")
recYAUTSYC0.AUTSYCSUI = rsADO("AUTSYCSUI")
recYAUTSYC0.AUTSYCELM = rsADO("AUTSYCELM")
recYAUTSYC0.AUTSYCNIV = rsADO("AUTSYCNIV")
recYAUTSYC0.AUTSYCINT = rsADO("AUTSYCINT")
recYAUTSYC0.AUTSYCEFF = rsADO("AUTSYCEFF")
recYAUTSYC0.AUTSYCPRO = rsADO("AUTSYCPRO")
recYAUTSYC0.AUTSYCDEB = rsADO("AUTSYCDEB")
recYAUTSYC0.AUTSYCFIN = rsADO("AUTSYCFIN")
recYAUTSYC0.AUTSYCMON = rsADO("AUTSYCMON")
recYAUTSYC0.AUTSYCDEV = rsADO("AUTSYCDEV")
recYAUTSYC0.AUTSYCBLO = rsADO("AUTSYCBLO")
recYAUTSYC0.AUTSYCAMO = rsADO("AUTSYCAMO")
recYAUTSYC0.AUTSYCGRP = rsADO("AUTSYCGRP")
recYAUTSYC0.AUTSYCRES = rsADO("AUTSYCRES")
recYAUTSYC0.AUTSYCTAU = rsADO("AUTSYCTAU")
recYAUTSYC0.AUTSYCDUR = rsADO("AUTSYCDUR")
recYAUTSYC0.AUTSYCCON = rsADO("AUTSYCCON")
recYAUTSYC0.AUTSYCCET = rsADO("AUTSYCCET")
recYAUTSYC0.AUTSYCCUT = rsADO("AUTSYCCUT")
recYAUTSYC0.AUTSYCUCR = rsADO("AUTSYCUCR")
recYAUTSYC0.AUTSYCUVL = rsADO("AUTSYCUVL")
recYAUTSYC0.AUTSYCUMO = rsADO("AUTSYCUMO")
recYAUTSYC0.AUTSYCDCR = rsADO("AUTSYCDCR")
recYAUTSYC0.AUTSYCDVL = rsADO("AUTSYCDVL")
recYAUTSYC0.AUTSYCDMO = rsADO("AUTSYCDMO")
Exit Function
Error_Handler:
srvYAUTSYC0_GetBuffer_ODBC = Error
End Function
Public Sub srvYAUTSYC0_ElpDisplay(recYAUTSYC0 As typeYAUTSYC0)
frmElpDisplay.fgData.Rows = 33
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTSYCETA    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTSYC0.AUTSYCETA
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTSYCGPE    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "GROUPE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTSYC0.AUTSYCGPE
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTSYCCLI    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N CLIENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTSYC0.AUTSYCCLI
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTSYCADR    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ADRESSE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTSYC0.AUTSYCADR
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTSYCTYP    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TYPE AUTO:1,2,3"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTSYC0.AUTSYCTYP
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTSYCAUT   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE AUTO"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTSYC0.AUTSYCAUT
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTSYCPER    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE PERE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTSYC0.AUTSYCPER
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTSYCSUI    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ADRESSE SUIVANTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTSYC0.AUTSYCSUI
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTSYCELM    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ELEMENTAIRE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTSYC0.AUTSYCELM
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTSYCNIV    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NIVEAU"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTSYC0.AUTSYCNIV
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTSYCINT    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE ECH. INTER."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTSYC0.AUTSYCINT
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTSYCEFF    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE EFFET"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTSYC0.AUTSYCEFF
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTSYCPRO    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE PROFIL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTSYC0.AUTSYCPRO
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTSYCDEB    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE DEBUT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTSYC0.AUTSYCDEB
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTSYCFIN    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE FIN"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTSYC0.AUTSYCFIN
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTSYCMON   15P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTSYC0.AUTSYCMON
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTSYCDEV    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DEVISE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTSYC0.AUTSYCDEV
frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTSYCBLO    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE BLOCAGE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTSYC0.AUTSYCBLO
frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTSYCAMO    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AMORTISSABLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTSYC0.AUTSYCAMO
frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTSYCGRP    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE GROUPE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTSYC0.AUTSYCGRP
frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTSYCRES    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE RESPONSABLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTSYC0.AUTSYCRES
frmElpDisplay.fgData.Row = 22
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTSYCTAU  6.3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TAUX DEPAS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTSYC0.AUTSYCTAU
frmElpDisplay.fgData.Row = 23
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTSYCDUR    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DUREE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTSYC0.AUTSYCDUR
frmElpDisplay.fgData.Row = 24
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTSYCCON    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CREDIT CONFIRME"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTSYC0.AUTSYCCON
frmElpDisplay.fgData.Row = 25
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTSYCCET    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE ETAT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTSYC0.AUTSYCCET
frmElpDisplay.fgData.Row = 26
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTSYCCUT    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE UTILISATEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTSYC0.AUTSYCCUT
frmElpDisplay.fgData.Row = 27
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTSYCUCR    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COD U.CREATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTSYC0.AUTSYCUCR
frmElpDisplay.fgData.Row = 28
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTSYCUVL    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COD U.VALIDATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTSYC0.AUTSYCUVL
frmElpDisplay.fgData.Row = 29
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTSYCUMO    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COD U.MODIFICAT."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTSYC0.AUTSYCUMO
frmElpDisplay.fgData.Row = 30
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTSYCDCR    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DAT CREATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTSYC0.AUTSYCDCR
frmElpDisplay.fgData.Row = 31
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTSYCDVL    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DAT VALIDATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTSYC0.AUTSYCDVL
frmElpDisplay.fgData.Row = 32
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AUTSYCDMO    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DAT MODIFICATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYAUTSYC0.AUTSYCDMO
frmElpDisplay.Show vbModal
End Sub
