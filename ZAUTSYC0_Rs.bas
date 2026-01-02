Attribute VB_Name = "rsZAUTSYC0"
'---------------------------------------------------------
Option Explicit
Type typeZAUTSYC0
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

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZAUTSYC0_GetBuffer(rsAdo As ADODB.Recordset, rsZAUTSYC0 As typeZAUTSYC0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZAUTSYC0_GetBuffer = Null

rsZAUTSYC0.AUTSYCETA = rsAdo("AUTSYCETA")
rsZAUTSYC0.AUTSYCGPE = rsAdo("AUTSYCGPE")
rsZAUTSYC0.AUTSYCCLI = rsAdo("AUTSYCCLI")
rsZAUTSYC0.AUTSYCADR = rsAdo("AUTSYCADR")
rsZAUTSYC0.AUTSYCTYP = rsAdo("AUTSYCTYP")
rsZAUTSYC0.AUTSYCAUT = rsAdo("AUTSYCAUT")
rsZAUTSYC0.AUTSYCPER = rsAdo("AUTSYCPER")
rsZAUTSYC0.AUTSYCSUI = rsAdo("AUTSYCSUI")
rsZAUTSYC0.AUTSYCELM = rsAdo("AUTSYCELM")
rsZAUTSYC0.AUTSYCNIV = rsAdo("AUTSYCNIV")
rsZAUTSYC0.AUTSYCINT = rsAdo("AUTSYCINT")
rsZAUTSYC0.AUTSYCEFF = rsAdo("AUTSYCEFF")
rsZAUTSYC0.AUTSYCPRO = rsAdo("AUTSYCPRO")
rsZAUTSYC0.AUTSYCDEB = rsAdo("AUTSYCDEB")
rsZAUTSYC0.AUTSYCFIN = rsAdo("AUTSYCFIN")
rsZAUTSYC0.AUTSYCMON = rsAdo("AUTSYCMON")
rsZAUTSYC0.AUTSYCDEV = rsAdo("AUTSYCDEV")
rsZAUTSYC0.AUTSYCBLO = rsAdo("AUTSYCBLO")
rsZAUTSYC0.AUTSYCAMO = rsAdo("AUTSYCAMO")
rsZAUTSYC0.AUTSYCGRP = rsAdo("AUTSYCGRP")
rsZAUTSYC0.AUTSYCRES = rsAdo("AUTSYCRES")
rsZAUTSYC0.AUTSYCTAU = rsAdo("AUTSYCTAU")
rsZAUTSYC0.AUTSYCDUR = rsAdo("AUTSYCDUR")
rsZAUTSYC0.AUTSYCCON = rsAdo("AUTSYCCON")
rsZAUTSYC0.AUTSYCCET = rsAdo("AUTSYCCET")
rsZAUTSYC0.AUTSYCCUT = rsAdo("AUTSYCCUT")
rsZAUTSYC0.AUTSYCUCR = rsAdo("AUTSYCUCR")
rsZAUTSYC0.AUTSYCUVL = rsAdo("AUTSYCUVL")
rsZAUTSYC0.AUTSYCUMO = rsAdo("AUTSYCUMO")
rsZAUTSYC0.AUTSYCDCR = rsAdo("AUTSYCDCR")
rsZAUTSYC0.AUTSYCDVL = rsAdo("AUTSYCDVL")
rsZAUTSYC0.AUTSYCDMO = rsAdo("AUTSYCDMO")
Exit Function

Error_Handler:

rsZAUTSYC0_GetBuffer = Error

End Function

'---------------------------------------------------------
Public Sub rsZAUTSYC0_Init(rsZAUTSYC0 As typeZAUTSYC0)
'---------------------------------------------------------
rsZAUTSYC0.AUTSYCETA = 0
rsZAUTSYC0.AUTSYCGPE = ""
rsZAUTSYC0.AUTSYCCLI = ""
rsZAUTSYC0.AUTSYCADR = 0
rsZAUTSYC0.AUTSYCTYP = ""
rsZAUTSYC0.AUTSYCAUT = ""
rsZAUTSYC0.AUTSYCPER = 0
rsZAUTSYC0.AUTSYCSUI = 0
rsZAUTSYC0.AUTSYCELM = ""
rsZAUTSYC0.AUTSYCNIV = 0
rsZAUTSYC0.AUTSYCINT = 0
rsZAUTSYC0.AUTSYCEFF = 0
rsZAUTSYC0.AUTSYCPRO = ""
rsZAUTSYC0.AUTSYCDEB = 0
rsZAUTSYC0.AUTSYCFIN = 0
rsZAUTSYC0.AUTSYCMON = 0
rsZAUTSYC0.AUTSYCDEV = ""
rsZAUTSYC0.AUTSYCBLO = ""
rsZAUTSYC0.AUTSYCAMO = ""
rsZAUTSYC0.AUTSYCGRP = ""
rsZAUTSYC0.AUTSYCRES = ""
rsZAUTSYC0.AUTSYCTAU = 0
rsZAUTSYC0.AUTSYCDUR = 0
rsZAUTSYC0.AUTSYCCON = ""
rsZAUTSYC0.AUTSYCCET = 0
rsZAUTSYC0.AUTSYCCUT = 0
rsZAUTSYC0.AUTSYCUCR = 0
rsZAUTSYC0.AUTSYCUVL = 0
rsZAUTSYC0.AUTSYCUMO = 0
rsZAUTSYC0.AUTSYCDCR = 0
rsZAUTSYC0.AUTSYCDVL = 0
rsZAUTSYC0.AUTSYCDMO = 0
End Sub


'








