Attribute VB_Name = "rsZCDOUTI0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeZCDOUTI0
    CDOUTIETB       As Integer                        ' CODE ETABLISSEMENT
    CDOUTIAGE       As Integer                        ' AGENCE
    CDOUTISER       As String * 2                     ' SERVICE
    CDOUTISSE       As String * 2                     ' SOUS-SERVICE
    CDOUTICOP       As String * 3                     ' CODE OPERATION
    CDOUTIDOS       As Long                           ' NUMERO DOSSIER
    CDOUTINUR       As Long                           ' N° RENOUVELLEMENT
    CDOUTIUTI       As Long                           ' N° UTILISATION
    CDOUTITMO       As String * 1                     ' C/N/D
    CDOUTIMON       As Currency                       ' MONTANT UTILISATION
    CDOUTIMAD       As Currency                       ' MONTANT ADDITIONNEL
    CDOUTIMTO       As Currency                       ' MONTANT TOTAL
    CDOUTIMDO       As Currency                       ' MONTANT DOCUMENTS
    CDOUTIMPA       As Currency                       ' MONTANT A PAYER
    CDOUTIPRE       As Long                           ' DATE PREVUE UTILIS.
    CDOUTIDAR       As Long                           ' DATE REFUS DOCUMEN.
    CDOUTIOBJ       As String * 6                     ' OBJET UTILISATION
    CDOUTIMVU       As Currency                       ' MONTANT A VUE
    CDOUTIMCA       As Currency                       ' MONTANT ACCEPTATION
    CDOUTIMDI       As Currency                       ' MONTANT DIFFERE
    CDOUTICTR       As String * 1                     ' REMETTANT (C/T)
    CDOUTIREM       As String * 7                     ' REMETTANT
    CDOUTIRER       As String * 16                    ' REFE.REMETTANT
    CDOUTIDRE       As Long                           ' DATE REMISE (EXP)
    CDOUTIDCO       As String * 1                     ' DOC.CONFORMES (O/N)
    CDOUTIDCE       As String * 1                     ' DOCUMENTS ENVOYES
    CDOUTIRET       As String * 1                     ' RESERVES TRANSMISES
    CDOUTIDAC       As String * 1                     ' DEMANDE ACCORD
    CDOUTIPAR       As String * 1                     ' PAY.SOUS RESERVES
    CDOUTIIRR       As String * 6                     ' IRREGULARITES
    CDOUTIPOR       As String * 1                     ' PORTEFEUILLE
    CDOUTIREF       As String * 1                     ' REFINANCEMENT
    CDOUTIESC       As String * 1                     ' ESCOMPTE
    CDOUTIBEC       As String * 1                     ' BENEF PAY.COMMIS°
    CDOUTIVA1       As Integer                        ' 1ER VALIDEUR
    CDOUTIVA2       As Integer                        ' 2EME VALIDEUR
    CDOUTIEVE       As String * 2                     ' EVENEMENT
    CDOUTIATT       As String * 2                     ' ATTENTE
    CDOUTIETA       As String * 2                     ' ETAT UTILISATION

End Type
Public Sub rsZCDOUTI0_Init(rsZCDOUTI0 As typeZCDOUTI0)
rsZCDOUTI0.CDOUTIETB = 0
rsZCDOUTI0.CDOUTIAGE = 0
rsZCDOUTI0.CDOUTISER = ""
rsZCDOUTI0.CDOUTISSE = ""
rsZCDOUTI0.CDOUTICOP = ""
rsZCDOUTI0.CDOUTIDOS = 0
rsZCDOUTI0.CDOUTINUR = 0
rsZCDOUTI0.CDOUTIUTI = 0
rsZCDOUTI0.CDOUTITMO = ""
rsZCDOUTI0.CDOUTIMON = 0
rsZCDOUTI0.CDOUTIMAD = 0
rsZCDOUTI0.CDOUTIMTO = 0
rsZCDOUTI0.CDOUTIMDO = 0
rsZCDOUTI0.CDOUTIMPA = 0
rsZCDOUTI0.CDOUTIPRE = 0
rsZCDOUTI0.CDOUTIDAR = 0
rsZCDOUTI0.CDOUTIOBJ = ""
rsZCDOUTI0.CDOUTIMVU = 0
rsZCDOUTI0.CDOUTIMCA = 0
rsZCDOUTI0.CDOUTIMDI = 0
rsZCDOUTI0.CDOUTICTR = ""
rsZCDOUTI0.CDOUTIREM = ""
rsZCDOUTI0.CDOUTIRER = ""
rsZCDOUTI0.CDOUTIDRE = 0
rsZCDOUTI0.CDOUTIDCO = ""
rsZCDOUTI0.CDOUTIDCE = ""
rsZCDOUTI0.CDOUTIRET = ""
rsZCDOUTI0.CDOUTIDAC = ""
rsZCDOUTI0.CDOUTIPAR = ""
rsZCDOUTI0.CDOUTIIRR = ""
rsZCDOUTI0.CDOUTIPOR = ""
rsZCDOUTI0.CDOUTIREF = ""
rsZCDOUTI0.CDOUTIESC = ""
rsZCDOUTI0.CDOUTIBEC = ""
rsZCDOUTI0.CDOUTIVA1 = 0
rsZCDOUTI0.CDOUTIVA2 = 0
rsZCDOUTI0.CDOUTIEVE = ""
rsZCDOUTI0.CDOUTIATT = ""
rsZCDOUTI0.CDOUTIETA = ""
End Sub
Public Function rsZCDOUTI0_GetBuffer(rsAdo As ADODB.Recordset, rsZCDOUTI0 As typeZCDOUTI0)
On Error GoTo Error_Handler
rsZCDOUTI0_GetBuffer = Null
rsZCDOUTI0.CDOUTIETB = rsAdo("CDOUTIETB")
rsZCDOUTI0.CDOUTIAGE = rsAdo("CDOUTIAGE")
rsZCDOUTI0.CDOUTISER = rsAdo("CDOUTISER")
rsZCDOUTI0.CDOUTISSE = rsAdo("CDOUTISSE")
rsZCDOUTI0.CDOUTICOP = rsAdo("CDOUTICOP")
rsZCDOUTI0.CDOUTIDOS = rsAdo("CDOUTIDOS")
rsZCDOUTI0.CDOUTINUR = rsAdo("CDOUTINUR")
rsZCDOUTI0.CDOUTIUTI = rsAdo("CDOUTIUTI")
rsZCDOUTI0.CDOUTITMO = rsAdo("CDOUTITMO")
rsZCDOUTI0.CDOUTIMON = rsAdo("CDOUTIMON")
rsZCDOUTI0.CDOUTIMAD = rsAdo("CDOUTIMAD")
rsZCDOUTI0.CDOUTIMTO = rsAdo("CDOUTIMTO")
rsZCDOUTI0.CDOUTIMDO = rsAdo("CDOUTIMDO")
rsZCDOUTI0.CDOUTIMPA = rsAdo("CDOUTIMPA")
rsZCDOUTI0.CDOUTIPRE = rsAdo("CDOUTIPRE")
rsZCDOUTI0.CDOUTIDAR = rsAdo("CDOUTIDAR")
rsZCDOUTI0.CDOUTIOBJ = rsAdo("CDOUTIOBJ")
rsZCDOUTI0.CDOUTIMVU = rsAdo("CDOUTIMVU")
rsZCDOUTI0.CDOUTIMCA = rsAdo("CDOUTIMCA")
rsZCDOUTI0.CDOUTIMDI = rsAdo("CDOUTIMDI")
rsZCDOUTI0.CDOUTICTR = rsAdo("CDOUTICTR")
rsZCDOUTI0.CDOUTIREM = rsAdo("CDOUTIREM")
rsZCDOUTI0.CDOUTIRER = rsAdo("CDOUTIRER")
rsZCDOUTI0.CDOUTIDRE = rsAdo("CDOUTIDRE")
rsZCDOUTI0.CDOUTIDCO = rsAdo("CDOUTIDCO")
rsZCDOUTI0.CDOUTIDCE = rsAdo("CDOUTIDCE")
rsZCDOUTI0.CDOUTIRET = rsAdo("CDOUTIRET")
rsZCDOUTI0.CDOUTIDAC = rsAdo("CDOUTIDAC")
rsZCDOUTI0.CDOUTIPAR = rsAdo("CDOUTIPAR")
rsZCDOUTI0.CDOUTIIRR = rsAdo("CDOUTIIRR")
rsZCDOUTI0.CDOUTIPOR = rsAdo("CDOUTIPOR")
rsZCDOUTI0.CDOUTIREF = rsAdo("CDOUTIREF")
rsZCDOUTI0.CDOUTIESC = rsAdo("CDOUTIESC")
rsZCDOUTI0.CDOUTIBEC = rsAdo("CDOUTIBEC")
rsZCDOUTI0.CDOUTIVA1 = rsAdo("CDOUTIVA1")
rsZCDOUTI0.CDOUTIVA2 = rsAdo("CDOUTIVA2")
rsZCDOUTI0.CDOUTIEVE = rsAdo("CDOUTIEVE")
rsZCDOUTI0.CDOUTIATT = rsAdo("CDOUTIATT")
rsZCDOUTI0.CDOUTIETA = rsAdo("CDOUTIETA")
Exit Function
Error_Handler:
rsZCDOUTI0_GetBuffer = Error
End Function

