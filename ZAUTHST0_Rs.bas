Attribute VB_Name = "rsZAUTHST0"
'---------------------------------------------------------
Option Explicit
Type typeZAUTHST0
    AUTHSTETA       As Integer                        ' ETABLISSEMENT
    AUTHSTGPE       As String * 1                     ' GROUPE
    AUTHSTCLI       As String * 7                     ' N CLIENT
    AUTHSTTYP       As String * 1                     ' TYPE AUTO:1,2,3
    AUTHSTAUT       As String * 20                    ' CODE AUTO
    AUTHSTMOD       As Long                           ' DATE MODIF.
    AUTHSTSEQ       As Long                           ' N SEQUENCE
    AUTHSTEFF       As Long                           ' DATE EFFET
    AUTHSTINT       As Long                           ' DATE ECH. INTER.
    AUTHSTPRO       As String * 3                     ' CODE PROFIL
    AUTHSTDEB       As Long                           ' DATE DEBUT AUTO
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

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZAUTHST0_GetBuffer(rsAdo As ADODB.Recordset, rsZAUTHST0 As typeZAUTHST0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZAUTHST0_GetBuffer = Null

rsZAUTHST0.AUTHSTETA = rsAdo("AUTHSTETA")
rsZAUTHST0.AUTHSTGPE = rsAdo("AUTHSTGPE")
rsZAUTHST0.AUTHSTCLI = rsAdo("AUTHSTCLI")
rsZAUTHST0.AUTHSTTYP = rsAdo("AUTHSTTYP")
rsZAUTHST0.AUTHSTAUT = rsAdo("AUTHSTAUT")
rsZAUTHST0.AUTHSTMOD = rsAdo("AUTHSTMOD")
rsZAUTHST0.AUTHSTSEQ = rsAdo("AUTHSTSEQ")
rsZAUTHST0.AUTHSTEFF = rsAdo("AUTHSTEFF")
rsZAUTHST0.AUTHSTINT = rsAdo("AUTHSTINT")
rsZAUTHST0.AUTHSTPRO = rsAdo("AUTHSTPRO")
rsZAUTHST0.AUTHSTDEB = rsAdo("AUTHSTDEB")
rsZAUTHST0.AUTHSTFIN = rsAdo("AUTHSTFIN")
rsZAUTHST0.AUTHSTMON = rsAdo("AUTHSTMON")
rsZAUTHST0.AUTHSTBLO = rsAdo("AUTHSTBLO")
rsZAUTHST0.AUTHSTTAU = rsAdo("AUTHSTTAU")
rsZAUTHST0.AUTHSTDUR = rsAdo("AUTHSTDUR")
rsZAUTHST0.AUTHSTCON = rsAdo("AUTHSTCON")
rsZAUTHST0.AUTHSTDEV = rsAdo("AUTHSTDEV")
rsZAUTHST0.AUTHSTCUT = rsAdo("AUTHSTCUT")
rsZAUTHST0.AUTHSTUCR = rsAdo("AUTHSTUCR")
rsZAUTHST0.AUTHSTUVL = rsAdo("AUTHSTUVL")
rsZAUTHST0.AUTHSTUMO = rsAdo("AUTHSTUMO")
rsZAUTHST0.AUTHSTDCR = rsAdo("AUTHSTDCR")
rsZAUTHST0.AUTHSTDVL = rsAdo("AUTHSTDVL")
rsZAUTHST0.AUTHSTDMO = rsAdo("AUTHSTDMO")
Exit Function

Error_Handler:

rsZAUTHST0_GetBuffer = Error

End Function

'---------------------------------------------------------
Public Sub rsZAUTHST0_Init(rsZAUTHST0 As typeZAUTHST0)
'---------------------------------------------------------
rsZAUTHST0.AUTHSTETA = 0
rsZAUTHST0.AUTHSTGPE = ""
rsZAUTHST0.AUTHSTCLI = ""
rsZAUTHST0.AUTHSTTYP = ""
rsZAUTHST0.AUTHSTAUT = ""
rsZAUTHST0.AUTHSTMOD = 0
rsZAUTHST0.AUTHSTSEQ = 0
rsZAUTHST0.AUTHSTEFF = 0
rsZAUTHST0.AUTHSTINT = 0
rsZAUTHST0.AUTHSTPRO = ""
rsZAUTHST0.AUTHSTDEB = 0
rsZAUTHST0.AUTHSTFIN = 0
rsZAUTHST0.AUTHSTMON = 0
rsZAUTHST0.AUTHSTBLO = ""
rsZAUTHST0.AUTHSTTAU = 0
rsZAUTHST0.AUTHSTDUR = 0
rsZAUTHST0.AUTHSTCON = ""
rsZAUTHST0.AUTHSTDEV = ""
rsZAUTHST0.AUTHSTCUT = 0
rsZAUTHST0.AUTHSTUCR = 0
rsZAUTHST0.AUTHSTUVL = 0
rsZAUTHST0.AUTHSTUMO = 0
rsZAUTHST0.AUTHSTDCR = 0
rsZAUTHST0.AUTHSTDVL = 0
rsZAUTHST0.AUTHSTDMO = 0
End Sub


'








