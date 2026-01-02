Attribute VB_Name = "rsZCREPRE0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeZCREPRE0
    CREPREETA       As Integer                        ' ETABLISSEMENT
    CREPREAGE       As Integer                        ' AGENCE
    CREPRESCE       As String * 2                     ' SERVICE
    CREPRESSE       As String * 2                     ' SOUS SERVICE
    CREPREDOS       As Long                           ' NUMERO DOSSIER
    CREPREPRE       As Long                           ' NUMERO PRET
    CREPRENAT       As String * 3                     ' NATURE PRET
    CREPREDEV       As String * 3                     ' DEVISE PRET
    CREPREDAE       As Long                           ' DERN ECH K TRAITEE
    CREPREDPE       As Long                           ' PROC ECH K A TRAITER
    CREPRECAL       As Long                           ' DTE DERN CALCUL INT
    CREPREDAI       As Long                           ' DERN ECH I TRAITEE
    CREPREDPI       As Long                           ' PROC ECH I A TRAITER
    CREPREDET       As Long                           ' DATE CODE ETAT
    CREPREOUV       As Long                           ' DATE OUVER PRET
    CREPREDER       As Long                           ' DATE DER. MODIF
    CREPREDIC       As Long                           ' DTE DERN CALCUL IC
    CREPREMON       As Currency                       ' MONTANT PRET
    CREPRECAP       As Currency                       ' CAPITAL RESTANT DU
    CREPREINT       As Currency                       ' INTERETS REPORTES
    CREPREICO       As Currency                       ' INTERETS COURUS
    CREPREICV       As Currency                       ' INTERETS COURUS CTVL
    CREPRECRS       As Double                         ' COURS
    CREPRECTA       As Long                           ' CODE ETAT
    CREPREPLA       As Long                           ' NUMERO PLAN
    CREPREPAL       As Long                           ' NUMERO PALIER
    CREPREECH       As Long                           ' NUMERO ECHEANCE
    CREPREAVI       As String * 1                     ' AVIS ECHEANCE
    CREPRETYR       As String * 1                     ' ANC.TYPE REPORT TYR
    CREPREINR       As String * 1                     ' ANC.INTEGRATION INR
    CREPREBAS       As Long                           ' ANC.BASE NBJ/AN BAS
    CREPREREA       As String * 1                     ' ANC.JOURS REELS REA
    CREPREPRC       As String * 1                     ' INTERETS PRECOMPTES
    CREPRESUP       As Long                           ' JOURS SUPPLEMENTS
    CREPRECOM       As Long                           ' ITC-COMMI-ASSUR
    CREPREAUT       As String * 12                    ' AUTORISATION
    CREPREUTI       As Integer                        ' UTILISATEUR
    CREPREOBJ       As String * 6                     ' OBJET PRET
    CREPREBAR       As String * 6                     ' BAREME
    CREPREREM       As String * 6                     ' REMBT ANTICIPE
    CREPREIMP       As String * 6                     ' GESTION IMPAYE
    CREPREFNC       As Double                         ' TAUX REFINANCEMENT
    CREPREINC       As Integer                        ' INCREMENT ECHEANCE
    CREPRETDO       As String * 1                     ' TYPE DE RECOUVREMENT
    CREPRESUS       As String * 1                     ' SUSPENSION DE RECOUV
    CREPREEXI       As String * 1                     ' EXISTENCE DS IMPAYES
    CREPREAGI       As String * 1                     ' AGIOS COMPENSES
    CREPRERGL       As String * 1                     ' COMPTE CREDITEUR
    CREPRECOD       As Integer                        ' UTILISAT. DERN MODIF
    CREPREOPT       As Long                           ' OPTION DERNIER MODIF

End Type
Public Sub rsZCREPRE0_Init(rsYCREPRE0 As typeZCREPRE0)
rsYCREPRE0.CREPREETA = 0
rsYCREPRE0.CREPREAGE = 0
rsYCREPRE0.CREPRESCE = ""
rsYCREPRE0.CREPRESSE = ""
rsYCREPRE0.CREPREDOS = 0
rsYCREPRE0.CREPREPRE = 0
rsYCREPRE0.CREPRENAT = ""
rsYCREPRE0.CREPREDEV = ""
rsYCREPRE0.CREPREDAE = 0
rsYCREPRE0.CREPREDPE = 0
rsYCREPRE0.CREPRECAL = 0
rsYCREPRE0.CREPREDAI = 0
rsYCREPRE0.CREPREDPI = 0
rsYCREPRE0.CREPREDET = 0
rsYCREPRE0.CREPREOUV = 0
rsYCREPRE0.CREPREDER = 0
rsYCREPRE0.CREPREDIC = 0
rsYCREPRE0.CREPREMON = 0
rsYCREPRE0.CREPRECAP = 0
rsYCREPRE0.CREPREINT = 0
rsYCREPRE0.CREPREICO = 0
rsYCREPRE0.CREPREICV = 0
rsYCREPRE0.CREPRECRS = 0
rsYCREPRE0.CREPRECTA = 0
rsYCREPRE0.CREPREPLA = 0
rsYCREPRE0.CREPREPAL = 0
rsYCREPRE0.CREPREECH = 0
rsYCREPRE0.CREPREAVI = ""
rsYCREPRE0.CREPRETYR = ""
rsYCREPRE0.CREPREINR = ""
rsYCREPRE0.CREPREBAS = 0
rsYCREPRE0.CREPREREA = ""
rsYCREPRE0.CREPREPRC = ""
rsYCREPRE0.CREPRESUP = 0
rsYCREPRE0.CREPRECOM = 0
rsYCREPRE0.CREPREAUT = ""
rsYCREPRE0.CREPREUTI = 0
rsYCREPRE0.CREPREOBJ = ""
rsYCREPRE0.CREPREBAR = ""
rsYCREPRE0.CREPREREM = ""
rsYCREPRE0.CREPREIMP = ""
rsYCREPRE0.CREPREFNC = 0
rsYCREPRE0.CREPREINC = 0
rsYCREPRE0.CREPRETDO = ""
rsYCREPRE0.CREPRESUS = ""
rsYCREPRE0.CREPREEXI = ""
rsYCREPRE0.CREPREAGI = ""
rsYCREPRE0.CREPRERGL = ""
rsYCREPRE0.CREPRECOD = 0
rsYCREPRE0.CREPREOPT = 0
End Sub
Public Function rsZCREPRE0_GetBuffer(rsAdo As ADODB.Recordset, rsZCREPRE0 As typeZCREPRE0)
On Error GoTo Error_Handler
rsZCREPRE0_GetBuffer = Null
rsZCREPRE0.CREPREETA = rsAdo("CREPREETA")
rsZCREPRE0.CREPREAGE = rsAdo("CREPREAGE")
rsZCREPRE0.CREPRESCE = rsAdo("CREPRESCE")
rsZCREPRE0.CREPRESSE = rsAdo("CREPRESSE")
rsZCREPRE0.CREPREDOS = rsAdo("CREPREDOS")
rsZCREPRE0.CREPREPRE = rsAdo("CREPREPRE")
rsZCREPRE0.CREPRENAT = rsAdo("CREPRENAT")
rsZCREPRE0.CREPREDEV = rsAdo("CREPREDEV")
rsZCREPRE0.CREPREDAE = rsAdo("CREPREDAE")
rsZCREPRE0.CREPREDPE = rsAdo("CREPREDPE")
rsZCREPRE0.CREPRECAL = rsAdo("CREPRECAL")
rsZCREPRE0.CREPREDAI = rsAdo("CREPREDAI")
rsZCREPRE0.CREPREDPI = rsAdo("CREPREDPI")
rsZCREPRE0.CREPREDET = rsAdo("CREPREDET")
rsZCREPRE0.CREPREOUV = rsAdo("CREPREOUV")
rsZCREPRE0.CREPREDER = rsAdo("CREPREDER")
rsZCREPRE0.CREPREDIC = rsAdo("CREPREDIC")
rsZCREPRE0.CREPREMON = rsAdo("CREPREMON")
rsZCREPRE0.CREPRECAP = rsAdo("CREPRECAP")
rsZCREPRE0.CREPREINT = rsAdo("CREPREINT")
rsZCREPRE0.CREPREICO = rsAdo("CREPREICO")
rsZCREPRE0.CREPREICV = rsAdo("CREPREICV")
rsZCREPRE0.CREPRECRS = rsAdo("CREPRECRS")
rsZCREPRE0.CREPRECTA = rsAdo("CREPRECTA")
rsZCREPRE0.CREPREPLA = rsAdo("CREPREPLA")
rsZCREPRE0.CREPREPAL = rsAdo("CREPREPAL")
rsZCREPRE0.CREPREECH = rsAdo("CREPREECH")
rsZCREPRE0.CREPREAVI = rsAdo("CREPREAVI")
rsZCREPRE0.CREPRETYR = rsAdo("CREPRETYR")
rsZCREPRE0.CREPREINR = rsAdo("CREPREINR")
rsZCREPRE0.CREPREBAS = rsAdo("CREPREBAS")
'rsZCREPRE0.CREPREREA = rsADO("CREPREREA")
rsZCREPRE0.CREPREPRC = rsAdo("CREPREPRC")
rsZCREPRE0.CREPRESUP = rsAdo("CREPRESUP")
rsZCREPRE0.CREPRECOM = rsAdo("CREPRECOM")
rsZCREPRE0.CREPREAUT = rsAdo("CREPREAUT")
rsZCREPRE0.CREPREUTI = rsAdo("CREPREUTI")
rsZCREPRE0.CREPREOBJ = rsAdo("CREPREOBJ")
rsZCREPRE0.CREPREBAR = rsAdo("CREPREBAR")
rsZCREPRE0.CREPREREM = rsAdo("CREPREREM")
rsZCREPRE0.CREPREIMP = rsAdo("CREPREIMP")
rsZCREPRE0.CREPREFNC = rsAdo("CREPREFNC")
rsZCREPRE0.CREPREINC = rsAdo("CREPREINC")
rsZCREPRE0.CREPRETDO = rsAdo("CREPRETDO")
rsZCREPRE0.CREPRESUS = rsAdo("CREPRESUS")
rsZCREPRE0.CREPREEXI = rsAdo("CREPREEXI")
rsZCREPRE0.CREPREAGI = rsAdo("CREPREAGI")
rsZCREPRE0.CREPRERGL = rsAdo("CREPRERGL")
rsZCREPRE0.CREPRECOD = rsAdo("CREPRECOD")
rsZCREPRE0.CREPREOPT = rsAdo("CREPREOPT")
Exit Function
Error_Handler:
rsZCREPRE0_GetBuffer = Error
End Function

