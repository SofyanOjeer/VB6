Attribute VB_Name = "rsZCHGOPE0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeZCHGOPE0
    CHGOPEETA       As Integer                        ' ETABLISSEMENT
    CHGOPEAGE       As Integer                        ' AGENCE
    CHGOPESER       As String * 2                     ' SERVICE
    CHGOPESSE       As String * 2                     ' S/SERVICE
    CHGOPEOPE       As String * 3                     ' CODE OPERATION
    CHGOPEDOS       As Long                           ' DOSSIER
    CHGOPENAT       As String * 3                     ' NATURE
    CHGOPECAM       As String * 3                     ' CAMBISTE
    CHGOPECON       As String * 7                     ' N° CONTREPARTIE
    CHGOPEPA1       As String * 1                     ' Passage uniquement  transfert e
    CHGOPECOU       As String * 7                     ' Non utilisé pour le      transf
    CHGOPEPA2       As String * 1                     ' 1=AVIS 2=CONFIRMA
    CHGOPEBDF       As String                         ' CODE BANQUE DE FR
    CHGOPECRE       As Long                           ' DTE CREATION OP.
    CHGOPECO1       As Double                         ' Cours comptant ou        strike
    CHGOPECO2       As Double                         ' Cours terme ou prime
    CHGOPESEN       As String * 1                     ' A_Achat  V_Vente         E_Emmi
    CHGOPEENG       As Long                           ' DATE ENGAGEMENT
    CHGOPEDT1       As Long                           ' DATE FIN/ECHEANCE
    CHGOPEDT2       As Long                           ' Date d'échéance comptant Swap o
    CHGOPEDTP       As Long                           ' Date=999999 compta.terminée
    CHGOPEDE1       As String * 3                     ' DEVISE 1
    CHGOPEMO1       As Double                         ' MONTANT 1
    CHGOPEDE2       As String * 3                     ' DEVISE 2
    CHGOPEMO2       As Double                         ' MONTANT 2
    CHGOPEMO3       As Double                         ' MONTANT TERME SWAP
    CHGOPEMO4       As Double                         ' Pour CAL/PUT/TER/SWP
    CHGOPEAOP       As String * 1                     ' Américaine ou européenne (A/E)
    CHGOPEEXE       As String * 1                     ' EXERCICE O/N
    CHGOPEDEX       As Long                           ' DATE.EXERCICE
    CHGOPEDRE       As Long                           ' Non utilisé
    CHGOPEFI3       As Double                         ' Résultat/prime
    CHGOPEFI4       As Double                         ' Option exercée
    CHGOPEMT3       As Double                         ' Fin exercice
    CHGOPEMT4       As Double                         ' Fin exercice
    CHGOPEMT5       As Double                         ' Option exercée
    CHGOPEMT6       As Double                         ' SEC-mont.dev2 clas7
    CHGOPETRA       As String * 3                     ' CODE TRANSACTION
    CHGOPEDCA       As Long                           ' DTE DERNIER CALCUL
    CHGOPEMOP       As Double                         ' En FRF si TER/SWP       En DEV
    CHGOPEMOF       As Double                         ' MONT PROVISIONNE
    CHGOPEENC       As String * 1                     ' TOPAGE EN COURS.
    CHGOPETE1       As Double                         ' COURS TERME DEV 1
    CHGOPETE2       As Double                         ' COURS TERME DEV 2
    CHGOPEMT1       As Double                         ' MONT.EN FRF DEV 1
    CHGOPEMT2       As Double                         ' MONT.EN FRF DEV 2
    CHGOPEFI1       As Double                         ' COURS FIXING DEV1
    CHGOPEFI2       As Double                         ' COURS FIXING DEV2
    CHGOPEEVE       As String * 2                     ' PT/PP/LT/LP/RV    prorogation/l
    CHGOPEINI       As Long                           ' DOSSIER INITIAL
    CHGOPEPRE       As Long                           ' DOSSIER PRECEDENT
    CHGOPESUI       As Long                           ' DOSSIER SUIVANT
    CHGOPEVAL       As String * 1                     ' VALIDATION-COMPTA
    CHGOPEANN       As String * 1                     ' ANNULATION
    CHGOPEAVI       As String * 1                     ' CONFIRMATION/AVIS
    CHGOPECOD       As Integer                        ' CODE UTILISATEUR
    CHGOPEEEN       As String * 3                     ' EVENE. ENGAGEMENT
    CHGOPEAN1       As String * 1                     ' EVENE.ENGAG.ANNUL
    CHGOPEMDA       As String * 3                     ' EVENE.MISE A DISPO
    CHGOPEAN2       As String * 1                     ' EVENE.M A D ANNUL
    CHGOPEECH       As String * 3                     ' EVENE.ECHEANCE-PRIME
    CHGOPEAN3       As String * 1                     ' EVENE.ECH/PRIM ANNUL
    CHGOPEPRO       As String * 3                     ' EVENE.PRORO-LEVEE
    CHGOPEAN4       As String * 1                     ' EVENE.PRO.LE.ANNUL
    CHGOPEREE       As String * 3                     ' EVENE. REESCOMPTE
    CHGOPEAN5       As String * 1                     ' EVENE.REESCOMP.ANNUL
    CHGOPEREV       As String * 3                     ' BLANC/A  /AER
    CHGOPEAN6       As String * 1                     ' MOTIF ANNUL.CHB
    CHGOPEGEO       As String * 3                     ' CODE GEOGRAPHIQUE
    CHGOPEFIL       As String * 1                     ' MOUV.DEB.ATTENTE
    CHGOPEDRP       As String * 3                     ' DEVISE REP/DEP
    CHGOPEMO5       As Double                         ' MONT.ARRONDI
    CHGOPEMO6       As Double                         ' MONT.ARRONDI SWAP
    CHGOPECO3       As Double                         ' C.COMPTANT DIVISE
    CHGOPECO4       As Double                         ' C.TERME    DIVISE
    CHGOPEUTI       As Integer                        ' CODE UTILIS.VALD.
    CHGOPEMT7       As Double                         ' intérêts devise 1
    CHGOPEMT8       As Double                         ' intérêts devise 2
    CHGOPEMT9       As Double                         ' intérêts devise 1
    CHGOPEM10       As Double                         ' intérêts devise 2
    CHGOPECHQ       As Long                           ' N° cheque
    CHGOPEM11       As Double                         ' CONTR.REP/DEP DEV
    CHGOPEDT3       As Long                           ' date  non utilise
    CHGOPEFIE       As String * 30                    ' FILLER
    CHGOPEMO7       As Double                         ' 2EME MONTANT SWAP
    CHGOPECO7       As Double                         ' CONTR.2E MONT SWP
    CHGOPEGRI       As Long                           ' N° GRILLE
    CHGOPEMAR       As Double                         ' MARGE
    CHGOPEDAN       As Long                           ' DATE ANNULATION
    CHGOPESWP       As String * 1                     ' SWAP 2 MONT O/N
    CHGOPEPOR       As String * 6                     ' PORTEFEUILLE
    CHGOPETIC       As String * 15                    ' TICKET SAISIE
    CHGOPEZN1       As String * 1                     ' AUTORISAT.A/B/S
    CHGOPEZN2       As String * 1                     ' EDITION FAX=O
    CHGOPEZN3       As String * 1                     ' AUTORISAT.A/B/S
    CHGOPEZN4       As String * 1                     ' CONTROLE MVT ATT.
    CHGOPEZN5       As String * 1                     ' NON UTILISE
    CHGOPEZN6       As String * 1                     ' NON UTILISE
    CHGOPEZN7       As String * 1                     ' NON UTILISE
    CHGOPEZN8       As String * 1                     ' NON UTILISE
    CHGOPEDVO       As String * 3                     ' DEVISE  ORIGINE
    CHGOPEMTO       As Double                         ' MONTANT ORIGINE
    CHGOPESCO       As String * 1                     ' SENS COURS / X
    CHGOPECI1       As Double                         ' COURS 1 INCERTAIN
    CHGOPECI2       As Double                         ' COURS 2 INCERTAIN
    CHGOPEDBA       As String * 3                     ' DEVISE DE BASE

End Type
Public Sub rsZCHGOPE0_Init(rsZCHGOPE0 As typeZCHGOPE0)
rsZCHGOPE0.CHGOPEETA = 0
rsZCHGOPE0.CHGOPEAGE = 0
rsZCHGOPE0.CHGOPESER = ""
rsZCHGOPE0.CHGOPESSE = ""
rsZCHGOPE0.CHGOPEOPE = ""
rsZCHGOPE0.CHGOPEDOS = 0
rsZCHGOPE0.CHGOPENAT = ""
rsZCHGOPE0.CHGOPECAM = ""
rsZCHGOPE0.CHGOPECON = ""
rsZCHGOPE0.CHGOPEPA1 = ""
rsZCHGOPE0.CHGOPECOU = ""
rsZCHGOPE0.CHGOPEPA2 = ""
rsZCHGOPE0.CHGOPEBDF = ""
rsZCHGOPE0.CHGOPECRE = 0
rsZCHGOPE0.CHGOPECO1 = 0
rsZCHGOPE0.CHGOPECO2 = 0
rsZCHGOPE0.CHGOPESEN = ""
rsZCHGOPE0.CHGOPEENG = 0
rsZCHGOPE0.CHGOPEDT1 = 0
rsZCHGOPE0.CHGOPEDT2 = 0
rsZCHGOPE0.CHGOPEDTP = 0
rsZCHGOPE0.CHGOPEDE1 = ""
rsZCHGOPE0.CHGOPEMO1 = 0
rsZCHGOPE0.CHGOPEDE2 = ""
rsZCHGOPE0.CHGOPEMO2 = 0
rsZCHGOPE0.CHGOPEMO3 = 0
rsZCHGOPE0.CHGOPEMO4 = 0
rsZCHGOPE0.CHGOPEAOP = ""
rsZCHGOPE0.CHGOPEEXE = ""
rsZCHGOPE0.CHGOPEDEX = 0
rsZCHGOPE0.CHGOPEDRE = 0
rsZCHGOPE0.CHGOPEFI3 = 0
rsZCHGOPE0.CHGOPEFI4 = 0
rsZCHGOPE0.CHGOPEMT3 = 0
rsZCHGOPE0.CHGOPEMT4 = 0
rsZCHGOPE0.CHGOPEMT5 = 0
rsZCHGOPE0.CHGOPEMT6 = 0
rsZCHGOPE0.CHGOPETRA = ""
rsZCHGOPE0.CHGOPEDCA = 0
rsZCHGOPE0.CHGOPEMOP = 0
rsZCHGOPE0.CHGOPEMOF = 0
rsZCHGOPE0.CHGOPEENC = ""
rsZCHGOPE0.CHGOPETE1 = 0
rsZCHGOPE0.CHGOPETE2 = 0
rsZCHGOPE0.CHGOPEMT1 = 0
rsZCHGOPE0.CHGOPEMT2 = 0
rsZCHGOPE0.CHGOPEFI1 = 0
rsZCHGOPE0.CHGOPEFI2 = 0
rsZCHGOPE0.CHGOPEEVE = ""
rsZCHGOPE0.CHGOPEINI = 0
rsZCHGOPE0.CHGOPEPRE = 0
rsZCHGOPE0.CHGOPESUI = 0
rsZCHGOPE0.CHGOPEVAL = ""
rsZCHGOPE0.CHGOPEANN = ""
rsZCHGOPE0.CHGOPEAVI = ""
rsZCHGOPE0.CHGOPECOD = 0
rsZCHGOPE0.CHGOPEEEN = ""
rsZCHGOPE0.CHGOPEAN1 = ""
rsZCHGOPE0.CHGOPEMDA = ""
rsZCHGOPE0.CHGOPEAN2 = ""
rsZCHGOPE0.CHGOPEECH = ""
rsZCHGOPE0.CHGOPEAN3 = ""
rsZCHGOPE0.CHGOPEPRO = ""
rsZCHGOPE0.CHGOPEAN4 = ""
rsZCHGOPE0.CHGOPEREE = ""
rsZCHGOPE0.CHGOPEAN5 = ""
rsZCHGOPE0.CHGOPEREV = ""
rsZCHGOPE0.CHGOPEAN6 = ""
rsZCHGOPE0.CHGOPEGEO = ""
rsZCHGOPE0.CHGOPEFIL = ""
rsZCHGOPE0.CHGOPEDRP = ""
rsZCHGOPE0.CHGOPEMO5 = 0
rsZCHGOPE0.CHGOPEMO6 = 0
rsZCHGOPE0.CHGOPECO3 = 0
rsZCHGOPE0.CHGOPECO4 = 0
rsZCHGOPE0.CHGOPEUTI = 0
rsZCHGOPE0.CHGOPEMT7 = 0
rsZCHGOPE0.CHGOPEMT8 = 0
rsZCHGOPE0.CHGOPEMT9 = 0
rsZCHGOPE0.CHGOPEM10 = 0
rsZCHGOPE0.CHGOPECHQ = 0
rsZCHGOPE0.CHGOPEM11 = 0
rsZCHGOPE0.CHGOPEDT3 = 0
rsZCHGOPE0.CHGOPEFIE = ""
rsZCHGOPE0.CHGOPEMO7 = 0
rsZCHGOPE0.CHGOPECO7 = 0
rsZCHGOPE0.CHGOPEGRI = 0
rsZCHGOPE0.CHGOPEMAR = 0
rsZCHGOPE0.CHGOPEDAN = 0
rsZCHGOPE0.CHGOPESWP = ""
rsZCHGOPE0.CHGOPEPOR = ""
rsZCHGOPE0.CHGOPETIC = ""
rsZCHGOPE0.CHGOPEZN1 = ""
rsZCHGOPE0.CHGOPEZN2 = ""
rsZCHGOPE0.CHGOPEZN3 = ""
rsZCHGOPE0.CHGOPEZN4 = ""
rsZCHGOPE0.CHGOPEZN5 = ""
rsZCHGOPE0.CHGOPEZN6 = ""
rsZCHGOPE0.CHGOPEZN7 = ""
rsZCHGOPE0.CHGOPEZN8 = ""
rsZCHGOPE0.CHGOPEDVO = ""
rsZCHGOPE0.CHGOPEMTO = 0
rsZCHGOPE0.CHGOPESCO = ""
rsZCHGOPE0.CHGOPECI1 = 0
rsZCHGOPE0.CHGOPECI2 = 0
rsZCHGOPE0.CHGOPEDBA = ""
End Sub
Public Function rsZCHGOPE0_GetBuffer(rsAdo As ADODB.Recordset, rsZCHGOPE0 As typeZCHGOPE0)
On Error GoTo Error_Handler
rsZCHGOPE0_GetBuffer = Null
rsZCHGOPE0.CHGOPEETA = rsAdo("CHGOPEETA")
rsZCHGOPE0.CHGOPEAGE = rsAdo("CHGOPEAGE")
rsZCHGOPE0.CHGOPESER = rsAdo("CHGOPESER")
rsZCHGOPE0.CHGOPESSE = rsAdo("CHGOPESSE")
rsZCHGOPE0.CHGOPEOPE = rsAdo("CHGOPEOPE")
rsZCHGOPE0.CHGOPEDOS = rsAdo("CHGOPEDOS")
rsZCHGOPE0.CHGOPENAT = rsAdo("CHGOPENAT")
rsZCHGOPE0.CHGOPECAM = rsAdo("CHGOPECAM")
rsZCHGOPE0.CHGOPECON = rsAdo("CHGOPECON")
rsZCHGOPE0.CHGOPEPA1 = rsAdo("CHGOPEPA1")
rsZCHGOPE0.CHGOPECOU = rsAdo("CHGOPECOU")
rsZCHGOPE0.CHGOPEPA2 = rsAdo("CHGOPEPA2")
rsZCHGOPE0.CHGOPEBDF = rsAdo("CHGOPEBDF")
rsZCHGOPE0.CHGOPECRE = rsAdo("CHGOPECRE")
rsZCHGOPE0.CHGOPECO1 = rsAdo("CHGOPECO1")
rsZCHGOPE0.CHGOPECO2 = rsAdo("CHGOPECO2")
rsZCHGOPE0.CHGOPESEN = rsAdo("CHGOPESEN")
rsZCHGOPE0.CHGOPEENG = rsAdo("CHGOPEENG")
rsZCHGOPE0.CHGOPEDT1 = rsAdo("CHGOPEDT1")
rsZCHGOPE0.CHGOPEDT2 = rsAdo("CHGOPEDT2")
rsZCHGOPE0.CHGOPEDTP = rsAdo("CHGOPEDTP")
rsZCHGOPE0.CHGOPEDE1 = rsAdo("CHGOPEDE1")
rsZCHGOPE0.CHGOPEMO1 = rsAdo("CHGOPEMO1")
rsZCHGOPE0.CHGOPEDE2 = rsAdo("CHGOPEDE2")
rsZCHGOPE0.CHGOPEMO2 = rsAdo("CHGOPEMO2")
rsZCHGOPE0.CHGOPEMO3 = rsAdo("CHGOPEMO3")
rsZCHGOPE0.CHGOPEMO4 = rsAdo("CHGOPEMO4")
rsZCHGOPE0.CHGOPEAOP = rsAdo("CHGOPEAOP")
rsZCHGOPE0.CHGOPEEXE = rsAdo("CHGOPEEXE")
rsZCHGOPE0.CHGOPEDEX = rsAdo("CHGOPEDEX")
rsZCHGOPE0.CHGOPEDRE = rsAdo("CHGOPEDRE")
rsZCHGOPE0.CHGOPEFI3 = rsAdo("CHGOPEFI3")
rsZCHGOPE0.CHGOPEFI4 = rsAdo("CHGOPEFI4")
rsZCHGOPE0.CHGOPEMT3 = rsAdo("CHGOPEMT3")
rsZCHGOPE0.CHGOPEMT4 = rsAdo("CHGOPEMT4")
rsZCHGOPE0.CHGOPEMT5 = rsAdo("CHGOPEMT5")
rsZCHGOPE0.CHGOPEMT6 = rsAdo("CHGOPEMT6")
rsZCHGOPE0.CHGOPETRA = rsAdo("CHGOPETRA")
rsZCHGOPE0.CHGOPEDCA = rsAdo("CHGOPEDCA")
rsZCHGOPE0.CHGOPEMOP = rsAdo("CHGOPEMOP")
rsZCHGOPE0.CHGOPEMOF = rsAdo("CHGOPEMOF")
rsZCHGOPE0.CHGOPEENC = rsAdo("CHGOPEENC")
rsZCHGOPE0.CHGOPETE1 = rsAdo("CHGOPETE1")
rsZCHGOPE0.CHGOPETE2 = rsAdo("CHGOPETE2")
rsZCHGOPE0.CHGOPEMT1 = rsAdo("CHGOPEMT1")
rsZCHGOPE0.CHGOPEMT2 = rsAdo("CHGOPEMT2")
rsZCHGOPE0.CHGOPEFI1 = rsAdo("CHGOPEFI1")
rsZCHGOPE0.CHGOPEFI2 = rsAdo("CHGOPEFI2")
rsZCHGOPE0.CHGOPEEVE = rsAdo("CHGOPEEVE")
rsZCHGOPE0.CHGOPEINI = rsAdo("CHGOPEINI")
rsZCHGOPE0.CHGOPEPRE = rsAdo("CHGOPEPRE")
rsZCHGOPE0.CHGOPESUI = rsAdo("CHGOPESUI")
rsZCHGOPE0.CHGOPEVAL = rsAdo("CHGOPEVAL")
rsZCHGOPE0.CHGOPEANN = rsAdo("CHGOPEANN")
rsZCHGOPE0.CHGOPEAVI = rsAdo("CHGOPEAVI")
rsZCHGOPE0.CHGOPECOD = rsAdo("CHGOPECOD")
rsZCHGOPE0.CHGOPEEEN = rsAdo("CHGOPEEEN")
rsZCHGOPE0.CHGOPEAN1 = rsAdo("CHGOPEAN1")
rsZCHGOPE0.CHGOPEMDA = rsAdo("CHGOPEMDA")
rsZCHGOPE0.CHGOPEAN2 = rsAdo("CHGOPEAN2")
rsZCHGOPE0.CHGOPEECH = rsAdo("CHGOPEECH")
rsZCHGOPE0.CHGOPEAN3 = rsAdo("CHGOPEAN3")
rsZCHGOPE0.CHGOPEPRO = rsAdo("CHGOPEPRO")
rsZCHGOPE0.CHGOPEAN4 = rsAdo("CHGOPEAN4")
rsZCHGOPE0.CHGOPEREE = rsAdo("CHGOPEREE")
rsZCHGOPE0.CHGOPEAN5 = rsAdo("CHGOPEAN5")
rsZCHGOPE0.CHGOPEREV = rsAdo("CHGOPEREV")
rsZCHGOPE0.CHGOPEAN6 = rsAdo("CHGOPEAN6")
rsZCHGOPE0.CHGOPEGEO = rsAdo("CHGOPEGEO")
rsZCHGOPE0.CHGOPEFIL = rsAdo("CHGOPEFIL")
rsZCHGOPE0.CHGOPEDRP = rsAdo("CHGOPEDRP")
rsZCHGOPE0.CHGOPEMO5 = rsAdo("CHGOPEMO5")
rsZCHGOPE0.CHGOPEMO6 = rsAdo("CHGOPEMO6")
rsZCHGOPE0.CHGOPECO3 = rsAdo("CHGOPECO3")
rsZCHGOPE0.CHGOPECO4 = rsAdo("CHGOPECO4")
rsZCHGOPE0.CHGOPEUTI = rsAdo("CHGOPEUTI")
rsZCHGOPE0.CHGOPEMT7 = rsAdo("CHGOPEMT7")
rsZCHGOPE0.CHGOPEMT8 = rsAdo("CHGOPEMT8")
rsZCHGOPE0.CHGOPEMT9 = rsAdo("CHGOPEMT9")
rsZCHGOPE0.CHGOPEM10 = rsAdo("CHGOPEM10")
rsZCHGOPE0.CHGOPECHQ = rsAdo("CHGOPECHQ")
rsZCHGOPE0.CHGOPEM11 = rsAdo("CHGOPEM11")
rsZCHGOPE0.CHGOPEDT3 = rsAdo("CHGOPEDT3")
rsZCHGOPE0.CHGOPEFIE = rsAdo("CHGOPEFIE")
rsZCHGOPE0.CHGOPEMO7 = rsAdo("CHGOPEMO7")
rsZCHGOPE0.CHGOPECO7 = rsAdo("CHGOPECO7")
rsZCHGOPE0.CHGOPEGRI = rsAdo("CHGOPEGRI")
rsZCHGOPE0.CHGOPEMAR = rsAdo("CHGOPEMAR")
rsZCHGOPE0.CHGOPEDAN = rsAdo("CHGOPEDAN")
rsZCHGOPE0.CHGOPESWP = rsAdo("CHGOPESWP")
rsZCHGOPE0.CHGOPEPOR = rsAdo("CHGOPEPOR")
rsZCHGOPE0.CHGOPETIC = rsAdo("CHGOPETIC")
rsZCHGOPE0.CHGOPEZN1 = rsAdo("CHGOPEZN1")
rsZCHGOPE0.CHGOPEZN2 = rsAdo("CHGOPEZN2")
rsZCHGOPE0.CHGOPEZN3 = rsAdo("CHGOPEZN3")
rsZCHGOPE0.CHGOPEZN4 = rsAdo("CHGOPEZN4")
rsZCHGOPE0.CHGOPEZN5 = rsAdo("CHGOPEZN5")
rsZCHGOPE0.CHGOPEZN6 = rsAdo("CHGOPEZN6")
rsZCHGOPE0.CHGOPEZN7 = rsAdo("CHGOPEZN7")
rsZCHGOPE0.CHGOPEZN8 = rsAdo("CHGOPEZN8")
rsZCHGOPE0.CHGOPEDVO = rsAdo("CHGOPEDVO")
rsZCHGOPE0.CHGOPEMTO = rsAdo("CHGOPEMTO")
rsZCHGOPE0.CHGOPESCO = rsAdo("CHGOPESCO")
rsZCHGOPE0.CHGOPECI1 = rsAdo("CHGOPECI1")
rsZCHGOPE0.CHGOPECI2 = rsAdo("CHGOPECI2")
rsZCHGOPE0.CHGOPEDBA = rsAdo("CHGOPEDBA")
Exit Function
Error_Handler:
rsZCHGOPE0_GetBuffer = Error
End Function

