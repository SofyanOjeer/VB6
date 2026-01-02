Attribute VB_Name = "rsZCHGDET0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeZCHGDET0
    CHGDETETA       As Integer                        ' ETABLISSEMENT
    CHGDETAGE       As Integer                        ' AGENCE
    CHGDETSER       As String * 2                     ' SERVICE
    CHGDETSSE       As String * 2                     ' S/SERVICE
    CHGDETOPE       As String * 3                     ' CODE OPERATION
    CHGDETDOS       As Long                           ' DOSSIER
    CHGDETTYP       As String * 1                     ' P_Principal
    CHGDETORD       As String * 2                     '  1er caractère   :       P = A
    CHGDETAG1       As Integer                        ' AGENCE
    CHGDETDE1       As String * 3                     ' DEVISE
    CHGDETCP1       As String * 20                    ' COMPTE
    CHGDETCL1       As String * 7                     ' CLIENT
    CHGDETAG2       As Integer                        ' AGENCE
    CHGDETDE2       As String * 3                     ' DEVISE
    CHGDETCP2       As String * 20                    ' COMPTE
    CHGDETCL2       As String * 7                     ' CLIENT
    CHGDETCOM       As String * 6                     ' CODE COMMISSION
    CHGDETREG       As String * 3                     ' CODE REGLEMENT
    CHGDETDTE       As Long                           ' DATE VALEUR
    CHGDETMES       As String * 1                     ' TOP MESSAGE
    CHGDETSEN       As String * 1                     ' SENS OPERATION
    CHGDETMON       As Double                         ' MONTANT
    CHGDETFRF       As Double                         ' MONTANT
    CHGDETEXO       As String * 1                     ' EXONERATION  O/N
    CHGDETTAX       As String * 1                     ' TAXABLE
    CHGDETTVA       As Double                         ' TAUX TVA
    CHGDETDFA       As Long                           ' DATE FACTURATION
    CHGDETFAC       As String * 1                     ' FACTURABLE
    CHGDETGLO       As Double                         ' MONTANT GLOBALISÉ
    CHGDETGFR       As Double                         ' MONTANT GLOBALISÉ

End Type
Public Sub srvZCHGDET0_Init(rszCHGDET0 As typeZCHGDET0)
rszCHGDET0.CHGDETETA = 0
rszCHGDET0.CHGDETAGE = 0
rszCHGDET0.CHGDETSER = ""
rszCHGDET0.CHGDETSSE = ""
rszCHGDET0.CHGDETOPE = ""
rszCHGDET0.CHGDETDOS = 0
rszCHGDET0.CHGDETTYP = ""
rszCHGDET0.CHGDETORD = ""
rszCHGDET0.CHGDETAG1 = 0
rszCHGDET0.CHGDETDE1 = ""
rszCHGDET0.CHGDETCP1 = ""
rszCHGDET0.CHGDETCL1 = ""
rszCHGDET0.CHGDETAG2 = 0
rszCHGDET0.CHGDETDE2 = ""
rszCHGDET0.CHGDETCP2 = ""
rszCHGDET0.CHGDETCL2 = ""
rszCHGDET0.CHGDETCOM = ""
rszCHGDET0.CHGDETREG = ""
rszCHGDET0.CHGDETDTE = 0
rszCHGDET0.CHGDETMES = ""
rszCHGDET0.CHGDETSEN = ""
rszCHGDET0.CHGDETMON = 0
rszCHGDET0.CHGDETFRF = 0
rszCHGDET0.CHGDETEXO = ""
rszCHGDET0.CHGDETTAX = ""
rszCHGDET0.CHGDETTVA = 0
rszCHGDET0.CHGDETDFA = 0
rszCHGDET0.CHGDETFAC = ""
rszCHGDET0.CHGDETGLO = 0
rszCHGDET0.CHGDETGFR = 0
End Sub
Public Function rsZCHGDET0_GetBuffer(rsADO As ADODB.Recordset, rszCHGDET0 As typeZCHGDET0)
On Error GoTo Error_Handler
rsZCHGDET0_GetBuffer = Null
rszCHGDET0.CHGDETETA = rsADO("CHGDETETA")
rszCHGDET0.CHGDETAGE = rsADO("CHGDETAGE")
rszCHGDET0.CHGDETSER = rsADO("CHGDETSER")
rszCHGDET0.CHGDETSSE = rsADO("CHGDETSSE")
rszCHGDET0.CHGDETOPE = rsADO("CHGDETOPE")
rszCHGDET0.CHGDETDOS = rsADO("CHGDETDOS")
rszCHGDET0.CHGDETTYP = rsADO("CHGDETTYP")
rszCHGDET0.CHGDETORD = rsADO("CHGDETORD")
rszCHGDET0.CHGDETAG1 = rsADO("CHGDETAG1")
rszCHGDET0.CHGDETDE1 = rsADO("CHGDETDE1")
rszCHGDET0.CHGDETCP1 = rsADO("CHGDETCP1")
rszCHGDET0.CHGDETCL1 = rsADO("CHGDETCL1")
rszCHGDET0.CHGDETAG2 = rsADO("CHGDETAG2")
rszCHGDET0.CHGDETDE2 = rsADO("CHGDETDE2")
rszCHGDET0.CHGDETCP2 = rsADO("CHGDETCP2")
rszCHGDET0.CHGDETCL2 = rsADO("CHGDETCL2")
rszCHGDET0.CHGDETCOM = rsADO("CHGDETCOM")
rszCHGDET0.CHGDETREG = rsADO("CHGDETREG")
rszCHGDET0.CHGDETDTE = rsADO("CHGDETDTE")
rszCHGDET0.CHGDETMES = rsADO("CHGDETMES")
rszCHGDET0.CHGDETSEN = rsADO("CHGDETSEN")
rszCHGDET0.CHGDETMON = rsADO("CHGDETMON")
rszCHGDET0.CHGDETFRF = rsADO("CHGDETFRF")
rszCHGDET0.CHGDETEXO = rsADO("CHGDETEXO")
rszCHGDET0.CHGDETTAX = rsADO("CHGDETTAX")
rszCHGDET0.CHGDETTVA = rsADO("CHGDETTVA")
rszCHGDET0.CHGDETDFA = rsADO("CHGDETDFA")
rszCHGDET0.CHGDETFAC = rsADO("CHGDETFAC")
rszCHGDET0.CHGDETGLO = rsADO("CHGDETGLO")
rszCHGDET0.CHGDETGFR = rsADO("CHGDETGFR")
Exit Function
Error_Handler:
rsZCHGDET0_GetBuffer = Error
End Function
