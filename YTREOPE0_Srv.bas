Attribute VB_Name = "srvYTREOPE0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const constYTREOPE0 = "YTREOPE0"
Type typeYTREOPE0
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
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
Public Sub srvYTREOPE0_Init(recYTREOPE0 As typeYTREOPE0)
recYTREOPE0.Obj = "YTREOPE0"
recYTREOPE0.Method = ""
recYTREOPE0.Err = ""
recYTREOPE0.TREOPEETB = 0
recYTREOPE0.TREOPEAGE = 0
recYTREOPE0.TREOPESER = ""
recYTREOPE0.TREOPESES = ""
recYTREOPE0.TREOPEOPR = ""
recYTREOPE0.TREOPENUM = 0
recYTREOPE0.TREOPENAT = ""
recYTREOPE0.TREOPENEG = 0
recYTREOPE0.TREOPEDIS = 0
recYTREOPE0.TREOPECLI = ""
recYTREOPE0.TREOPEDEV = ""
recYTREOPE0.TREOPEMNT = 0
recYTREOPE0.TREOPEECH = 0
recYTREOPE0.TREOPEREE = 0
recYTREOPE0.TREOPEPRC = 0
recYTREOPE0.TREOPETIE = ""
recYTREOPE0.TREOPECOU = ""
recYTREOPE0.TREOPEABR = ""
recYTREOPE0.TREOPEDAT = 0
recYTREOPE0.TREOPEBAS = 0
recYTREOPE0.TREOPEPRE = ""
recYTREOPE0.TREOPEPER = 0
recYTREOPE0.TREOPELEV = 0
recYTREOPE0.TREOPEREN = 0
recYTREOPE0.TREOPEROP = 0
recYTREOPE0.TREOPEGAR = ""
recYTREOPE0.TREOPEPOS = ""
recYTREOPE0.TREOPEBAN = ""
recYTREOPE0.TREOPECIV = ""
recYTREOPE0.TREOPEDUR = ""
recYTREOPE0.TREOPENBP = 0
recYTREOPE0.TREOPECOM = ""
recYTREOPE0.TREOPETRA = ""
recYTREOPE0.TREOPEFRF = ""
recYTREOPE0.TREOPEREP = ""
recYTREOPE0.TREOPECNF = ""
recYTREOPE0.TREOPEAUT = ""
recYTREOPE0.TREOPECAM = ""
recYTREOPE0.TREOPEEFF = ""
recYTREOPE0.TREOPEPAP = ""
recYTREOPE0.TREOPENAN = ""
recYTREOPE0.TREOPECMT = ""
recYTREOPE0.TREOPEARR = ""
recYTREOPE0.TREOPEETA = ""
recYTREOPE0.TREOPESAI = 0
recYTREOPE0.TREOPEUT1 = 0
recYTREOPE0.TREOPEUT2 = 0
recYTREOPE0.TREOPEORI = ""
recYTREOPE0.TREOPEPTF = ""
recYTREOPE0.TREOPETIC = ""
recYTREOPE0.TREOPENET = ""
recYTREOPE0.TREOPERBT = ""
recYTREOPE0.TREOPEDRB = 0
recYTREOPE0.TREOPEADO = ""
End Sub
Public Function srvYTREOPE0_GetBuffer_ODBC(rsADO As ADODB.Recordset, recYTREOPE0 As typeYTREOPE0)
On Error GoTo Error_Handler
srvYTREOPE0_GetBuffer_ODBC = Null
recYTREOPE0.TREOPEETB = rsADO("TREOPEETB")
recYTREOPE0.TREOPEAGE = rsADO("TREOPEAGE")
recYTREOPE0.TREOPESER = rsADO("TREOPESER")
recYTREOPE0.TREOPESES = rsADO("TREOPESES")
recYTREOPE0.TREOPEOPR = rsADO("TREOPEOPR")
recYTREOPE0.TREOPENUM = rsADO("TREOPENUM")
recYTREOPE0.TREOPENAT = rsADO("TREOPENAT")
recYTREOPE0.TREOPENEG = rsADO("TREOPENEG")
recYTREOPE0.TREOPEDIS = rsADO("TREOPEDIS")
recYTREOPE0.TREOPECLI = rsADO("TREOPECLI")
recYTREOPE0.TREOPEDEV = rsADO("TREOPEDEV")
recYTREOPE0.TREOPEMNT = rsADO("TREOPEMNT")
recYTREOPE0.TREOPEECH = rsADO("TREOPEECH")
recYTREOPE0.TREOPEREE = rsADO("TREOPEREE")
recYTREOPE0.TREOPEPRC = rsADO("TREOPEPRC")
recYTREOPE0.TREOPETIE = rsADO("TREOPETIE")
recYTREOPE0.TREOPECOU = rsADO("TREOPECOU")
recYTREOPE0.TREOPEABR = rsADO("TREOPEABR")
recYTREOPE0.TREOPEDAT = rsADO("TREOPEDAT")
recYTREOPE0.TREOPEBAS = rsADO("TREOPEBAS")
recYTREOPE0.TREOPEPRE = rsADO("TREOPEPRE")
recYTREOPE0.TREOPEPER = rsADO("TREOPEPER")
recYTREOPE0.TREOPELEV = rsADO("TREOPELEV")
recYTREOPE0.TREOPEREN = rsADO("TREOPEREN")
recYTREOPE0.TREOPEROP = rsADO("TREOPEROP")
recYTREOPE0.TREOPEGAR = rsADO("TREOPEGAR")
recYTREOPE0.TREOPEPOS = rsADO("TREOPEPOS")
recYTREOPE0.TREOPEBAN = rsADO("TREOPEBAN")
recYTREOPE0.TREOPECIV = rsADO("TREOPECIV")
recYTREOPE0.TREOPEDUR = rsADO("TREOPEDUR")
recYTREOPE0.TREOPENBP = rsADO("TREOPENBP")
recYTREOPE0.TREOPECOM = rsADO("TREOPECOM")
recYTREOPE0.TREOPETRA = rsADO("TREOPETRA")
recYTREOPE0.TREOPEFRF = rsADO("TREOPEFRF")
recYTREOPE0.TREOPEREP = rsADO("TREOPEREP")
recYTREOPE0.TREOPECNF = rsADO("TREOPECNF")
recYTREOPE0.TREOPEAUT = rsADO("TREOPEAUT")
recYTREOPE0.TREOPECAM = rsADO("TREOPECAM")
recYTREOPE0.TREOPEEFF = rsADO("TREOPEEFF")
recYTREOPE0.TREOPEPAP = rsADO("TREOPEPAP")
recYTREOPE0.TREOPENAN = rsADO("TREOPENAN")
recYTREOPE0.TREOPECMT = rsADO("TREOPECMT")
recYTREOPE0.TREOPEARR = rsADO("TREOPEARR")
recYTREOPE0.TREOPEETA = rsADO("TREOPEETA")
recYTREOPE0.TREOPESAI = rsADO("TREOPESAI")
recYTREOPE0.TREOPEUT1 = rsADO("TREOPEUT1")
recYTREOPE0.TREOPEUT2 = rsADO("TREOPEUT2")
recYTREOPE0.TREOPEORI = rsADO("TREOPEORI")
recYTREOPE0.TREOPEPTF = rsADO("TREOPEPTF")
recYTREOPE0.TREOPETIC = rsADO("TREOPETIC")
recYTREOPE0.TREOPENET = rsADO("TREOPENET")
recYTREOPE0.TREOPERBT = rsADO("TREOPERBT")
recYTREOPE0.TREOPEDRB = rsADO("TREOPEDRB")
recYTREOPE0.TREOPEADO = rsADO("TREOPEADO")
Exit Function
Error_Handler:
srvYTREOPE0_GetBuffer_ODBC = Error
End Function
Public Sub srvYTREOPE0_ElpDisplay(recYTREOPE0 As typeYTREOPE0)
frmElpDisplay.fgData.Rows = 55
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEETB    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEETB
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEAGE    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AGENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEAGE
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPESER    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPESER
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPESES    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOUS-SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPESES
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEOPR    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE OPERATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEOPR
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPENUM    9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO OPERATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPENUM
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPENAT    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NATURE DE L OPERATI"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPENAT
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPENEG    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE DE NEGOCIATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPENEG
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEDIS    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE MISE A DISPOSIT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEDIS
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPECLI    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N°CLIEN CONTREPARTIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPECLI
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEDEV    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DEVISE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEDEV
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEMNT 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT EN DEVISE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEMNT
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEECH    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE ECHEANCE PREVUE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEECH
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEREE    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE ECHEANCE REELLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEREE
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEPRC    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE PROCHAIN ECHEAN"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEPRC
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPETIE    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TIERS COURTIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPETIE
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPECOU    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COURTIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPECOU
frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEABR   12A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ABREGE COURTIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEABR
frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEDAT    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE 1ére ECHEANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEDAT
frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEBAS    1P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BASE CAL.INTERE(0/5)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEBAS
frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEPRE    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PERIODE DE PREAVIS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEPRE
frmElpDisplay.fgData.Row = 22
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEPER    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NBR PERIODE PREAVIS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEPER
frmElpDisplay.fgData.Row = 23
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPELEV    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE LEVEE DE PREAVI"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPELEV
frmElpDisplay.fgData.Row = 24
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEREN    9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "RENOUVE.PAR OPERA.N°"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEREN
frmElpDisplay.fgData.Row = 25
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEROP    9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "RENOUVELLE OPERA. N°"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEROP
frmElpDisplay.fgData.Row = 26
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEGAR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "GARANTIE CONTRAT O/N"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEGAR
frmElpDisplay.fgData.Row = 27
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEPOS    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PRE-POSTCOMPTE (D/T)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEPOS
frmElpDisplay.fgData.Row = 28
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEBAN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NBR JOUR REEL/BANCAI"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEBAN
frmElpDisplay.fgData.Row = 29
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPECIV    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PERIOD CIVIL/ANNIVER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPECIV
frmElpDisplay.fgData.Row = 30
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEDUR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PERIO.DURE J,M,T,S,A"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEDUR
frmElpDisplay.fgData.Row = 31
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPENBP    2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NOMBRE DE PERIODE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPENBP
frmElpDisplay.fgData.Row = 32
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPECOM    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COMPARAISON (< OU >)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPECOM
frmElpDisplay.fgData.Row = 33
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPETRA    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TYPE TRANSACTION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPETRA
frmElpDisplay.fgData.Row = 34
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEFRF    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "FRANCS - DEVISE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEFRF
frmElpDisplay.fgData.Row = 35
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEREP    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "REPORT (S/P/ )"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEREP
frmElpDisplay.fgData.Row = 36
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPECNF    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AVIS.OPER/FICH.CONFI"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPECNF
frmElpDisplay.fgData.Row = 37
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEAUT   12A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE AUTORISATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEAUT
frmElpDisplay.fgData.Row = 38
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPECAM    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE CAMBISTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPECAM
frmElpDisplay.fgData.Row = 39
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEEFF    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "EFFET BLOQUES/LIVRES"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEEFF
frmElpDisplay.fgData.Row = 40
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEPAP    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NATURE PAPIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEPAP
frmElpDisplay.fgData.Row = 41
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPENAN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NATURE PAPIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPENAN
frmElpDisplay.fgData.Row = 42
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPECMT   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COMMENTAIRE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPECMT
frmElpDisplay.fgData.Row = 43
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEARR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ARRONDI (I/S/ )"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEARR
frmElpDisplay.fgData.Row = 44
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEETA    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE ETAT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEETA
frmElpDisplay.fgData.Row = 45
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPESAI    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE SAISIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPESAI
frmElpDisplay.fgData.Row = 46
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEUT1    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "UTILISATEUR 1"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEUT1
frmElpDisplay.fgData.Row = 47
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEUT2    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "UTILISATEUR 2"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEUT2
frmElpDisplay.fgData.Row = 48
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEORI    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COURUS COURS ORIGINE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEORI
frmElpDisplay.fgData.Row = 49
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEPTF    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE PORTEFEUILLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEPTF
frmElpDisplay.fgData.Row = 50
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPETIC   15A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TICKET DE SAISIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPETIC
frmElpDisplay.fgData.Row = 51
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPENET    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TOP NETTING"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPENET
frmElpDisplay.fgData.Row = 52
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPERBT    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TOP DEMANDE DE RBT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPERBT
frmElpDisplay.fgData.Row = 53
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEDRB    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE RBT/REMISE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEDRB
frmElpDisplay.fgData.Row = 54
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TREOPEADO    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TOP ADOSSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYTREOPE0.TREOPEADO
frmElpDisplay.Show vbModal
End Sub

