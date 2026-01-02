Attribute VB_Name = "rsZTREOPE0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeZTREOPE0
    TREOPEETB       As Integer                        ' ETABLISSEMENT
    TREOPEAGE       As Integer                        ' AGENCE
    TREOPESER       As String * 2                     ' SERVICE
    TREOPESES       As String * 2                     ' SOUS-SERVICE
    TREOPEOPR       As String * 3                     ' CODE OPERATION
    TREOPENUM       As Long                           ' NUMERO OPERATION
    TREOPENAT       As String * 3                     ' NATURE DE L OPERATI
    TREOPENEG       As Long                           ' DATE DE NEGOCIATION
    TREOPEDIS       As Long                           ' DATE MISE A DISPOSIT
    TREOPECLI       As String * 7                     ' N°CLIEN CONTREPARTIE
    TREOPEDEV       As String * 3                     ' DEVISE
    TREOPEMNT       As Currency                       ' MONTANT EN DEVISE
    TREOPEECH       As Long                           ' DATE ECHEANCE PREVUE
    TREOPEREE       As Long                           ' DATE ECHEANCE REELLE
    TREOPEPRC       As Long                           ' DATE PROCHAIN ECHEAN
    TREOPETIE       As String * 1                     ' TIERS COURTIER
    TREOPECOU       As String * 7                     ' COURTIER
    TREOPEABR       As String * 12                    ' ABREGE COURTIER
    TREOPEDAT       As Long                           ' DATE 1ére ECHEANCE
    TREOPEBAS       As Long                           ' BASE CAL.INTERE(0/5)
    TREOPEPRE       As String * 1                     ' PERIODE DE PREAVIS
    TREOPEPER       As Long                           ' NBR PERIODE PREAVIS
    TREOPELEV       As Long                           ' DATE LEVEE DE PREAVI
    TREOPEREN       As Long                           ' RENOUVE.PAR OPERA.N°
    TREOPEROP       As Long                           ' RENOUVELLE OPERA. N°
    TREOPEGAR       As String * 1                     ' GARANTIE CONTRAT O/N
    TREOPEPOS       As String * 1                     ' PRE-POSTCOMPTE (D/T)
    TREOPEBAN       As String * 1                     ' NBR JOUR REEL/BANCAI
    TREOPECIV       As String * 1                     ' PERIOD CIVIL/ANNIVER
    TREOPEDUR       As String * 1                     ' PERIO.DURE J,M,T,S,A
    TREOPENBP       As Long                           ' NOMBRE DE PERIODE
    TREOPECOM       As String * 1                     ' COMPARAISON (< OU >)
    TREOPETRA       As String * 3                     ' TYPE TRANSACTION
    TREOPEFRF       As String * 1                     ' FRANCS - DEVISE
    TREOPEREP       As String * 1                     ' REPORT (S/P/ )
    TREOPECNF       As String * 1                     ' AVIS.OPER/FICH.CONFI
    TREOPEAUT       As String * 12                    ' CODE AUTORISATION
    TREOPECAM       As String * 3                     ' CODE CAMBISTE
    TREOPEEFF       As String * 1                     ' EFFET BLOQUES/LIVRES
    TREOPEPAP       As String * 3                     ' NATURE PAPIER
    TREOPENAN       As String * 1                     ' NATURE PAPIER
    TREOPECMT       As String * 20                    ' COMMENTAIRE
    TREOPEARR       As String * 1                     ' ARRONDI (I/S/ )
    TREOPEETA       As String * 1                     ' CODE ETAT
    TREOPESAI       As Long                           ' DATE SAISIE
    TREOPEUT1       As Integer                        ' UTILISATEUR 1
    TREOPEUT2       As Integer                        ' UTILISATEUR 2
    TREOPEORI       As String * 1                     ' COURUS COURS ORIGINE
    TREOPEPTF       As String * 6                     ' CODE PORTEFEUILLE
    TREOPETIC       As String * 15                    ' TICKET DE SAISIE
    TREOPENET       As String * 1                     ' TOP NETTING
    TREOPERBT       As String * 1                     ' TOP DEMANDE DE RBT
    TREOPEDRB       As Long                           ' DATE RBT/REMISE
    TREOPEADO       As String * 1                     ' TOP ADOSSEMENT

End Type
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZTREOPE0_GetBuffer(rsAdo As ADODB.Recordset, rsZTREOPE0 As typeZTREOPE0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZTREOPE0_GetBuffer = Null

rsZTREOPE0.TREOPEETB = rsAdo("TREOPEETB")
rsZTREOPE0.TREOPEAGE = rsAdo("TREOPEAGE")
rsZTREOPE0.TREOPESER = rsAdo("TREOPESER")
rsZTREOPE0.TREOPESES = rsAdo("TREOPESES")
rsZTREOPE0.TREOPEOPR = rsAdo("TREOPEOPR")
rsZTREOPE0.TREOPENUM = rsAdo("TREOPENUM")
rsZTREOPE0.TREOPENAT = rsAdo("TREOPENAT")
rsZTREOPE0.TREOPENEG = rsAdo("TREOPENEG")
rsZTREOPE0.TREOPEDIS = rsAdo("TREOPEDIS")
rsZTREOPE0.TREOPECLI = rsAdo("TREOPECLI")
rsZTREOPE0.TREOPEDEV = rsAdo("TREOPEDEV")
rsZTREOPE0.TREOPEMNT = rsAdo("TREOPEMNT")
rsZTREOPE0.TREOPEECH = rsAdo("TREOPEECH")
rsZTREOPE0.TREOPEREE = rsAdo("TREOPEREE")
rsZTREOPE0.TREOPEPRC = rsAdo("TREOPEPRC")
rsZTREOPE0.TREOPETIE = rsAdo("TREOPETIE")
rsZTREOPE0.TREOPECOU = rsAdo("TREOPECOU")
rsZTREOPE0.TREOPEABR = rsAdo("TREOPEABR")
rsZTREOPE0.TREOPEDAT = rsAdo("TREOPEDAT")
rsZTREOPE0.TREOPEBAS = rsAdo("TREOPEBAS")
rsZTREOPE0.TREOPEPRE = rsAdo("TREOPEPRE")
rsZTREOPE0.TREOPEPER = rsAdo("TREOPEPER")
rsZTREOPE0.TREOPELEV = rsAdo("TREOPELEV")
rsZTREOPE0.TREOPEREN = rsAdo("TREOPEREN")
rsZTREOPE0.TREOPEROP = rsAdo("TREOPEROP")
rsZTREOPE0.TREOPEGAR = rsAdo("TREOPEGAR")
rsZTREOPE0.TREOPEPOS = rsAdo("TREOPEPOS")
rsZTREOPE0.TREOPEBAN = rsAdo("TREOPEBAN")
rsZTREOPE0.TREOPECIV = rsAdo("TREOPECIV")
rsZTREOPE0.TREOPEDUR = rsAdo("TREOPEDUR")
rsZTREOPE0.TREOPENBP = rsAdo("TREOPENBP")
rsZTREOPE0.TREOPECOM = rsAdo("TREOPECOM")
rsZTREOPE0.TREOPETRA = rsAdo("TREOPETRA")
rsZTREOPE0.TREOPEFRF = rsAdo("TREOPEFRF")
rsZTREOPE0.TREOPEREP = rsAdo("TREOPEREP")
rsZTREOPE0.TREOPECNF = rsAdo("TREOPECNF")
rsZTREOPE0.TREOPEAUT = rsAdo("TREOPEAUT")
rsZTREOPE0.TREOPECAM = rsAdo("TREOPECAM")
rsZTREOPE0.TREOPEEFF = rsAdo("TREOPEEFF")
rsZTREOPE0.TREOPEPAP = rsAdo("TREOPEPAP")
rsZTREOPE0.TREOPENAN = rsAdo("TREOPENAN")
rsZTREOPE0.TREOPECMT = rsAdo("TREOPECMT")
rsZTREOPE0.TREOPEARR = rsAdo("TREOPEARR")
rsZTREOPE0.TREOPEETA = rsAdo("TREOPEETA")
rsZTREOPE0.TREOPESAI = rsAdo("TREOPESAI")
rsZTREOPE0.TREOPEUT1 = rsAdo("TREOPEUT1")
rsZTREOPE0.TREOPEUT2 = rsAdo("TREOPEUT2")
rsZTREOPE0.TREOPEORI = rsAdo("TREOPEORI")
rsZTREOPE0.TREOPEPTF = rsAdo("TREOPEPTF")
rsZTREOPE0.TREOPETIC = rsAdo("TREOPETIC")
rsZTREOPE0.TREOPENET = rsAdo("TREOPENET")
rsZTREOPE0.TREOPERBT = rsAdo("TREOPERBT")
rsZTREOPE0.TREOPEDRB = rsAdo("TREOPEDRB")
rsZTREOPE0.TREOPEADO = rsAdo("TREOPEADO")
Exit Function

Error_Handler:

rsZTREOPE0_GetBuffer = Error

End Function

'---------------------------------------------------------
Public Sub rsZTREOPE0_Init(rsZTREOPE0 As typeZTREOPE0)
'---------------------------------------------------------
rsZTREOPE0.TREOPEETB = 0
rsZTREOPE0.TREOPEAGE = 0
rsZTREOPE0.TREOPESER = ""
rsZTREOPE0.TREOPESES = ""
rsZTREOPE0.TREOPEOPR = ""
rsZTREOPE0.TREOPENUM = 0
rsZTREOPE0.TREOPENAT = ""
rsZTREOPE0.TREOPENEG = 0
rsZTREOPE0.TREOPEDIS = 0
rsZTREOPE0.TREOPECLI = ""
rsZTREOPE0.TREOPEDEV = ""
rsZTREOPE0.TREOPEMNT = 0
rsZTREOPE0.TREOPEECH = 0
rsZTREOPE0.TREOPEREE = 0
rsZTREOPE0.TREOPEPRC = 0
rsZTREOPE0.TREOPETIE = ""
rsZTREOPE0.TREOPECOU = ""
rsZTREOPE0.TREOPEABR = ""
rsZTREOPE0.TREOPEDAT = 0
rsZTREOPE0.TREOPEBAS = 0
rsZTREOPE0.TREOPEPRE = ""
rsZTREOPE0.TREOPEPER = 0
rsZTREOPE0.TREOPELEV = 0
rsZTREOPE0.TREOPEREN = 0
rsZTREOPE0.TREOPEROP = 0
rsZTREOPE0.TREOPEGAR = ""
rsZTREOPE0.TREOPEPOS = ""
rsZTREOPE0.TREOPEBAN = ""
rsZTREOPE0.TREOPECIV = ""
rsZTREOPE0.TREOPEDUR = ""
rsZTREOPE0.TREOPENBP = 0
rsZTREOPE0.TREOPECOM = ""
rsZTREOPE0.TREOPETRA = ""
rsZTREOPE0.TREOPEFRF = ""
rsZTREOPE0.TREOPEREP = ""
rsZTREOPE0.TREOPECNF = ""
rsZTREOPE0.TREOPEAUT = ""
rsZTREOPE0.TREOPECAM = ""
rsZTREOPE0.TREOPEEFF = ""
rsZTREOPE0.TREOPEPAP = ""
rsZTREOPE0.TREOPENAN = ""
rsZTREOPE0.TREOPECMT = ""
rsZTREOPE0.TREOPEARR = ""
rsZTREOPE0.TREOPEETA = ""
rsZTREOPE0.TREOPESAI = 0
rsZTREOPE0.TREOPEUT1 = 0
rsZTREOPE0.TREOPEUT2 = 0
rsZTREOPE0.TREOPEORI = ""
rsZTREOPE0.TREOPEPTF = ""
rsZTREOPE0.TREOPETIC = ""
rsZTREOPE0.TREOPENET = ""
rsZTREOPE0.TREOPERBT = ""
rsZTREOPE0.TREOPEDRB = 0
rsZTREOPE0.TREOPEADO = ""
End Sub


'








