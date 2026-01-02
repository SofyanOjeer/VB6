Attribute VB_Name = "rsYAUTE1I0"
'---------------------------------------------------------
Option Explicit
Type typeYAUTE1I0
    AUTE1IETA       As Integer                        ' ETABLISSEMENT
    AUTE1IGRP       As String * 7                     ' GROUPE
    AUTE1ICLI       As String * 7                     ' CLIENT
    AUTE1ITYP       As String * 1                     ' TYPE 1,2,3
    AUTE1IAUT       As String * 20                    ' CODE AUTORISATION
    AUTE1IDEV       As String * 3                     ' DEVISE
    AUTE1IAGE       As Integer                        ' AGENCE
    AUTE1ISER       As String * 2                     ' SERVICE
    AUTE1ISRV       As String * 2                     ' SOUS SERVICE
    AUTE1ICOP       As String * 3                     ' CODE OPERATION
    AUTE1INOP       As Long                           ' NUMERO OPERATION
    AUTE1IOR1       As Long                           ' ORDRE 1
    AUTE1IOR2       As Long                           ' ORDRE 2
    AUTE1IOR3       As Long                           ' ORDRE 3
    AUTE1IOR4       As Long                           ' ORDRE 4
    AUTE1IDBA       As String * 3                     ' DEVISE DE BASE
    AUTE1IMDB       As Long                           ' MONTANT DEBIT
    AUTE1IMCR       As Long                           ' MONTANT CREDIT
    AUTE1IBDB       As Long                           ' MONTANT DEVBAS DB
    AUTE1IBCR       As Long                           ' MONTANT DEVBAS CR
    AUTE1IRDB       As Long                           ' MONTANT REPOR. DB
    AUTE1IRCR       As Long                           ' MONTANT REPOR. CR
    AUTE1IMAU       As Long                           ' MONTANT AUTO
    AUTE1IDAD       As Long                           ' DATE DEBUT AUTO.
    AUTE1IDAF       As Long                           ' DATE FIN AUTO.
    AUTE1IINT       As String * 1                     ' INTITULITE
    AUTE1IDMO       As Long                           ' DATE DERN.MOUV.
    AUTE1IRA1       As String * 32                    ' RAISON SOCIALE
    AUTE1IRA2       As String * 32                    ' RAISON SOCIALE 2
    AUTE1ISAC       As String * 6                     ' SECTEUR ACTIVITE
    AUTE1IREG       As String * 6                     ' SECTEUR ACT. REG.
    AUTE1ISRN       As String * 9                     ' NUMERO SIREN
    AUTE1IRES       As String * 3                     ' RESPON/EXPLOIT
    AUTE1IECO       As String * 3                     ' QUALITE/AG.ECONO
    AUTE1ICOT       As String * 3                     ' COTATION INTERNE
    AUTE1IBDF       As String * 4                     ' CODE BDF
    AUTE1IDOU       As String * 1                     ' DOUTEUX  O/N
    AUTE1IICH       As String * 1                     ' INTERDIT CHQ  O/N
    AUTE1ICET       As String * 4                     ' CODE ETAT
    AUTE1ISIG       As String * 12                    ' SIGLE
    AUTE1IRAG       As String * 32                    ' RAISON SOC GROUPE
    AUTE1IELM       As String * 1                     ' CODE ELEM. O/N
    AUTE1INIV       As Long                           ' NIVEAU
    AUTE1IBLO       As String * 1                     ' CODE BLOCAGE1,2,3
    AUTE1ICOM       As String * 1                     ' COMPENSATION O/N
    AUTE1ILAU       As String * 30                    ' LIBEL AUTO OU GAR
    AUTE1IECI       As Long                           ' ECHEANCE INTERNE
    AUTE1IDEP       As String * 1                     ' CODE DEPASSEMENT
    AUTE1IMTD       As Long                           ' MONTANT DEPASSEM
    AUTE1IDPD       As Long                           ' DATE 1ER DEPAS.
    AUTE1IDTD       As Long                           ' DEPASSEM  DEPUIS
    AUTE1IC1A       As String * 1                     ' C1AUT POUR DEPASS
    AUTE1IDEB       As Long                           ' DATE DEBUT OPERA
    AUTE1IFIN       As Long                           ' DATE FIN OPERA
    AUTE1ILIB       As String * 32                    ' LIBELLE OPERATION
    AUTE1IRAT       As String * 1                     ' RATTAC GROUPE O/N
    AUTE1IATR       As String * 1                     ' AUTO GROUPE(1à9)
    AUTE1IREL       As String * 3                     ' RELATION CLI-GRP
    AUTE1IRUB       As String * 10                    ' RUBRIQUE COMPT.
    AUTE1IAGC       As Integer                        ' AGENCE CLIENT
    AUTE1ISEG       As String * 3                     ' SEGEMENT DE RESULTAT
    AUTE1ISEP       As String * 3                     ' SEGEMENT POTENTIEL
    AUTE1IFUT       As String * 150                   ' ZONE FUTURE
    
    
End Type

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsYAUTE1I0_GetBuffer(rsAdo As ADODB.Recordset, rsYAUTE1I0 As typeYAUTE1I0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYAUTE1I0_GetBuffer = Null

rsYAUTE1I0.AUTE1IETA = rsAdo("AUTE1IETA")
rsYAUTE1I0.AUTE1IGRP = rsAdo("AUTE1IGRP")
rsYAUTE1I0.AUTE1ICLI = rsAdo("AUTE1ICLI")
rsYAUTE1I0.AUTE1ITYP = rsAdo("AUTE1ITYP")
rsYAUTE1I0.AUTE1IAUT = rsAdo("AUTE1IAUT")
rsYAUTE1I0.AUTE1IDEV = rsAdo("AUTE1IDEV")
rsYAUTE1I0.AUTE1IAGE = rsAdo("AUTE1IAGE")
rsYAUTE1I0.AUTE1ISER = rsAdo("AUTE1ISER")
rsYAUTE1I0.AUTE1ISRV = rsAdo("AUTE1ISRV")
rsYAUTE1I0.AUTE1ICOP = rsAdo("AUTE1ICOP")
rsYAUTE1I0.AUTE1INOP = rsAdo("AUTE1INOP")
rsYAUTE1I0.AUTE1IOR1 = rsAdo("AUTE1IOR1")
rsYAUTE1I0.AUTE1IOR2 = rsAdo("AUTE1IOR2")
rsYAUTE1I0.AUTE1IOR3 = rsAdo("AUTE1IOR3")
rsYAUTE1I0.AUTE1IOR4 = rsAdo("AUTE1IOR4")
rsYAUTE1I0.AUTE1IDBA = rsAdo("AUTE1IDBA")
rsYAUTE1I0.AUTE1IMDB = rsAdo("AUTE1IMDB")
rsYAUTE1I0.AUTE1IMCR = rsAdo("AUTE1IMCR")
rsYAUTE1I0.AUTE1IBDB = rsAdo("AUTE1IBBDB")
rsYAUTE1I0.AUTE1IBCR = rsAdo("AUTE1IBCR")
rsYAUTE1I0.AUTE1IRDB = rsAdo("AUTE1IRDB")
rsYAUTE1I0.AUTE1IRCR = rsAdo("AUTE1IRCR")
rsYAUTE1I0.AUTE1IMAU = rsAdo("AUTE1IMAU")
rsYAUTE1I0.AUTE1IDAD = rsAdo("AUTE1IDAD")
rsYAUTE1I0.AUTE1IDAF = rsAdo("AUTE1IDAF")
rsYAUTE1I0.AUTE1IINT = rsAdo("AUTE1IINT")
rsYAUTE1I0.AUTE1IDMO = rsAdo("AUTE1IDMO")
rsYAUTE1I0.AUTE1IRA1 = rsAdo("AUTE1IRA1")
rsYAUTE1I0.AUTE1IRA2 = rsAdo("AUTE1IRA2")
rsYAUTE1I0.AUTE1ISAC = rsAdo("AUTE1ISAC")
rsYAUTE1I0.AUTE1IREG = rsAdo("AUTE1IREG")
rsYAUTE1I0.AUTE1ISRN = rsAdo("AUTE1ISRN")
rsYAUTE1I0.AUTE1IRES = rsAdo("AUTE1IRES")
rsYAUTE1I0.AUTE1IECO = rsAdo("AUTE1IECO")
rsYAUTE1I0.AUTE1ICOT = rsAdo("AUTE1ICOT")
rsYAUTE1I0.AUTE1IBDF = rsAdo("AUTE1IBDF")
rsYAUTE1I0.AUTE1IDOU = rsAdo("AUTE1IDOU")
rsYAUTE1I0.AUTE1IICH = rsAdo("AUTE1IICH")
rsYAUTE1I0.AUTE1ICET = rsAdo("AUTE1ICET")
rsYAUTE1I0.AUTE1ISIG = rsAdo("AUTE1ISIG")
rsYAUTE1I0.AUTE1IRAG = rsAdo("AUTE1IRAG")
rsYAUTE1I0.AUTE1IELM = rsAdo("AUTE1IELM")
rsYAUTE1I0.AUTE1INIV = rsAdo("AUTE1INIV")
rsYAUTE1I0.AUTE1IBLO = rsAdo("AUTE1IBLO")
rsYAUTE1I0.AUTE1ICOM = rsAdo("AUTE1ICOM")
rsYAUTE1I0.AUTE1ILAU = rsAdo("AUTE1ILAU")
rsYAUTE1I0.AUTE1IECI = rsAdo("AUTE1IECI")
rsYAUTE1I0.AUTE1IDEP = rsAdo("AUTE1IDEP")
rsYAUTE1I0.AUTE1IMTD = rsAdo("AUTE1IMTD")
rsYAUTE1I0.AUTE1IDPD = rsAdo("AUTE1IDPD")
rsYAUTE1I0.AUTE1IDTD = rsAdo("AUTE1IDTD")
rsYAUTE1I0.AUTE1IC1A = rsAdo("AUTE1IC1A")
rsYAUTE1I0.AUTE1IDEB = rsAdo("AUTE1IDEB")
rsYAUTE1I0.AUTE1IFIN = rsAdo("AUTE1IFIN")
rsYAUTE1I0.AUTE1ILIB = rsAdo("AUTE1ILIB")
rsYAUTE1I0.AUTE1IRAT = rsAdo("AUTE1IRAT")
rsYAUTE1I0.AUTE1IATR = rsAdo("AUTE1IATR")
rsYAUTE1I0.AUTE1IREL = rsAdo("AUTE1IREL")
rsYAUTE1I0.AUTE1IRUB = rsAdo("AUTE1IRUB")
rsYAUTE1I0.AUTE1IAGC = rsAdo("AUTE1IAGC")
rsYAUTE1I0.AUTE1ISEG = rsAdo("AUTE1ISEG")
rsYAUTE1I0.AUTE1ISEP = rsAdo("AUTE1ISEP")
rsYAUTE1I0.AUTE1IFUT = rsAdo("AUTE1IFUT")

Exit Function

Error_Handler:

rsYAUTE1I0_GetBuffer = Error

End Function


'







