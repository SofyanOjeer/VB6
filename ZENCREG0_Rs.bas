Attribute VB_Name = "rszENCREG0"
Option Explicit

Type typeZENCREG0
    ENCREGETA       As Integer                        'code établissement
    ENCREGAGE       As Integer                        ' AGENCE
    ENCREGSER       As String * 2                     ' SERVICE
    ENCREGSSE       As String * 2                     ' SOUS-SERVICE
    ENCREGCOP       As String * 3                     ' CODE OPERATION
    ENCREGDOS       As Long                           ' NUMERO DOSSIER
    ENCREGREG       As Long                           'Num. RGL ou dem.COM
    ENCREGSEN       As String * 1                     ' Sens D/C
    ENCREGTYP       As String * 1                     ' R/C
    ENCREGPAI       As String * 3                     ' Type de paiement (VUE...)
    ENCREGNLI       As Long                           ' Type N° de ligne
    ENCREGMOR       As String * 3                     ' Mode de règlement (SWF/EBA...)
    ENCREGDEV       As String * 3                     ' Devise
    ENCREGCOU       As Currency                       ' Cours de la devise
    ENCREGCOB       As Currency                       ' Cours de la devise de base
    ENCREGMON       As Currency                       ' Montant du règlement
    ENCREGMOD       As Currency                       ' Montant en devise du dossier
    ENCREGMOB       As Currency                       ' Montant en devise de base
    ENCREGDEN       As Long                           ' Date de l'engagement
    ENCREGDRE       As Long                           ' Date du règlement
    ENCREGDCR       As Long                           ' Date du règlement au comptant
    ENCREGDVA       As Long                           ' Date de valeur
    ENCREGDAN       As Long                           ' Date de l'annulation
    ENCREGDAE       As Long                           ' date d'échéance
    ENCREGDAI       As Long                           ' Date INIT CHANGE
    ENCREGCL1       As String * 8                     ' Code Payeur (débit)
    ENCREGBCL       As String * 12                    ' BIC du payeur (débit)
    ENCREGICL       As String * 1                     ' controle IBAN du payeur (débit)
    ENCREGLCL       As String * 34                    ' Zone libre payeur (débit)
    ENCREGCOM       As String * 20                    ' N° de compte payeur (débit)
    ENCREGCL2       As String * 8                     ' Banque recevant les fonds
    ENCREGBC2       As String * 12                    ' BIC
    ENCREGIC2       As String * 1                     ' controle IBAN du bénéficiaire
    ENCREGLC2       As String * 34                    ' Zone libre bénéficiaire
    ENCREGCL3       As String * 8                     ' Banque remise des fonds
    ENCREGBC3       As String * 12                    ' BIC bénéficiaire
    ENCREGIC3       As String * 1                     ' controle IBAN du remettant
    ENCREGLC3       As String * 34                    ' Zone libre remettant
    ENCREGCL4       As String * 8                     ' Banque remettant
    ENCREGBC4       As String * 12                    ' BIC remettant
    ENCREGIC4       As String * 1                     ' controle IBAN du destinataire
    ENCREGLC4       As String * 34                    ' Zone libre bénéficiaire
    ENCREGCL5       As String * 8                     ' Code destinataire (crédit)
    ENCREGBC5       As String * 12                    ' BIC destinataire (crédit)
    ENCREGIC5       As String * 1                     ' contrôle IBAN destinataire (crédit)
    ENCREGLC5       As String * 34                    ' Zone libre destinataire (crédit)
    ENCREGBDF       As String * 3                     ' Code BDF (débit)
    ENCREGPAY       As String * 3                     ' Code pays (débit)
    ENCREGCRP       As String * 1                     ' Type de CRP
    ENCREGBAN       As Long                           ' Code Banque
    ENCREGGUI       As Long                           ' Code guichet
    ENCREGCHQ       As Long                           ' N° de chèque
    ENCREGNDE       As String * 24                    ' Nom du destinataire (du chèque)
    ENCREGEVO       As String * 1                     ' ETAT VIREMT ODC (O/N)
    ENCREGRIB       As Long                           ' RIB code banque
    ENCREGRIG       As Long                           ' RIB code guichet
    ENCREGRIC       As String * 20                    ' N° de compte
    ENCREGESC       As String * 1                     ' Code escompte
    ENCREGRES       As Long                           ' Référence de l'escompte
    ENCREGPOR       As String * 6                     ' Code portefeuille
    ENCREGINF       As String * 140                   ' Informations supplémentaires
    ENCREGDET       As String * 140                   ' Détail du paiement
    ENCREGEFF       As String * 1                     ' Effet en protefeuille
    ENCREGUEN       As Integer                        ' Code utilisateur saisie
    ENCREGUL1       As Integer                        ' Code 1er utilisateur validation
    ENCREGUL2       As Integer                        ' Code 2eme utilisateur validation
    ENCREGCET       As String * 2                     ' Code ETAT
    ENCREGCHA       As String * 1                     ' Charge (B...)
    ENCREGATG       As String * 1                     ' Attente gestion
    ENCREGEUP       As String * 1                     ' Code ETAT virement EUP
End Type

Public Function rsZENCREG0_GetBuffer(rsADO As ADODB.Recordset, rsZENCREG0 As typeZENCREG0)

    rsZENCREG0_GetBuffer = Null
    rsZENCREG0.ENCREGETA = rsADO("ENCREGETA")
    rsZENCREG0.ENCREGAGE = rsADO("ENCREGAGE")
    rsZENCREG0.ENCREGSER = rsADO("ENCREGSER")
    rsZENCREG0.ENCREGSSE = rsADO("ENCREGSSE")
    rsZENCREG0.ENCREGCOP = rsADO("ENCREGCOP")
    rsZENCREG0.ENCREGDOS = rsADO("ENCREGDOS")
    rsZENCREG0.ENCREGREG = rsADO("ENCREGREG")
    rsZENCREG0.ENCREGSEN = rsADO("ENCREGSEN")
    rsZENCREG0.ENCREGTYP = rsADO("ENCREGTYP")
    rsZENCREG0.ENCREGPAI = rsADO("ENCREGPAI")
    rsZENCREG0.ENCREGNLI = rsADO("ENCREGNLI")
    rsZENCREG0.ENCREGMOR = rsADO("ENCREGMOR")
    rsZENCREG0.ENCREGDEV = rsADO("ENCREGDEV")
    rsZENCREG0.ENCREGCOU = rsADO("ENCREGCOU")
    rsZENCREG0.ENCREGCOB = rsADO("ENCREGCOB")
    rsZENCREG0.ENCREGMON = rsADO("ENCREGMON")
    rsZENCREG0.ENCREGMOD = rsADO("ENCREGMOD")
    rsZENCREG0.ENCREGMOB = rsADO("ENCREGMOB")
    rsZENCREG0.ENCREGDEN = rsADO("ENCREGDEN")
    rsZENCREG0.ENCREGDRE = rsADO("ENCREGDRE")
    rsZENCREG0.ENCREGDCR = rsADO("ENCREGDCR")
    rsZENCREG0.ENCREGDVA = rsADO("ENCREGDVA")
    rsZENCREG0.ENCREGDAN = rsADO("ENCREGDAN")
    rsZENCREG0.ENCREGDAE = rsADO("ENCREGDAE")
    rsZENCREG0.ENCREGDAI = rsADO("ENCREGDAI")
    rsZENCREG0.ENCREGCL1 = rsADO("ENCREGCL1")
    rsZENCREG0.ENCREGBCL = rsADO("ENCREGBCL")
    rsZENCREG0.ENCREGICL = rsADO("ENCREGICL")
    rsZENCREG0.ENCREGLCL = rsADO("ENCREGLCL")
    rsZENCREG0.ENCREGCOM = rsADO("ENCREGCOM")
    rsZENCREG0.ENCREGCL2 = rsADO("ENCREGCL2")
    rsZENCREG0.ENCREGBC2 = rsADO("ENCREGBC2")
    rsZENCREG0.ENCREGIC2 = rsADO("ENCREGIC2")
    rsZENCREG0.ENCREGLC2 = rsADO("ENCREGLC2")
    rsZENCREG0.ENCREGCL3 = rsADO("ENCREGCL3")
    rsZENCREG0.ENCREGBC3 = rsADO("ENCREGBC3")
    rsZENCREG0.ENCREGIC3 = rsADO("ENCREGIC3")
    rsZENCREG0.ENCREGLC3 = rsADO("ENCREGLC3")
    rsZENCREG0.ENCREGCL4 = rsADO("ENCREGCL4")
    rsZENCREG0.ENCREGBC4 = rsADO("ENCREGBC4")
    rsZENCREG0.ENCREGIC4 = rsADO("ENCREGIC4")
    rsZENCREG0.ENCREGLC4 = rsADO("ENCREGLC4")
    rsZENCREG0.ENCREGCL5 = rsADO("ENCREGCL5")
    rsZENCREG0.ENCREGBC5 = rsADO("ENCREGBC5")
    rsZENCREG0.ENCREGIC5 = rsADO("ENCREGIC5")
    rsZENCREG0.ENCREGLC5 = rsADO("ENCREGLC5")
    rsZENCREG0.ENCREGBDF = rsADO("ENCREGBDF")
    rsZENCREG0.ENCREGPAY = rsADO("ENCREGPAY")
    rsZENCREG0.ENCREGCRP = rsADO("ENCREGCRP")
    rsZENCREG0.ENCREGBAN = rsADO("ENCREGBAN")
    rsZENCREG0.ENCREGGUI = rsADO("ENCREGGUI")
    rsZENCREG0.ENCREGCHQ = rsADO("ENCREGCHQ")
    rsZENCREG0.ENCREGNDE = rsADO("ENCREGNDE")
    rsZENCREG0.ENCREGEVO = rsADO("ENCREGEVO")
    rsZENCREG0.ENCREGRIB = rsADO("ENCREGRIB")
    rsZENCREG0.ENCREGRIG = rsADO("ENCREGRIG")
    rsZENCREG0.ENCREGRIC = rsADO("ENCREGRIC")
    rsZENCREG0.ENCREGESC = rsADO("ENCREGESC")
    rsZENCREG0.ENCREGRES = rsADO("ENCREGRES")
    rsZENCREG0.ENCREGPOR = rsADO("ENCREGPOR")
    rsZENCREG0.ENCREGINF = rsADO("ENCREGINF")
    rsZENCREG0.ENCREGDET = rsADO("ENCREGDET")
    rsZENCREG0.ENCREGEFF = rsADO("ENCREGEFF")
    rsZENCREG0.ENCREGUEN = rsADO("ENCREGUEN")
    rsZENCREG0.ENCREGUL1 = rsADO("ENCREGUL1")
    rsZENCREG0.ENCREGUL2 = rsADO("ENCREGUL2")
    rsZENCREG0.ENCREGCET = rsADO("ENCREGCET")
    rsZENCREG0.ENCREGCHA = rsADO("ENCREGCHA")
    rsZENCREG0.ENCREGATG = rsADO("ENCREGATG")
    rsZENCREG0.ENCREGEUP = rsADO("ENCREGEUP")
    Exit Function
Error_Handler:
    rsZENCREG0_GetBuffer = Error

End Function


Public Sub rsZENCREG0_Init(rsZENCREG0 As typeZENCREG0)

    rsZENCREG0.ENCREGETA = 0
    rsZENCREG0.ENCREGAGE = 0
    rsZENCREG0.ENCREGSER = ""
    rsZENCREG0.ENCREGSSE = ""
    rsZENCREG0.ENCREGCOP = ""
    rsZENCREG0.ENCREGDOS = 0
    rsZENCREG0.ENCREGREG = 0
    rsZENCREG0.ENCREGSEN = ""
    rsZENCREG0.ENCREGTYP = ""
    rsZENCREG0.ENCREGPAI = ""
    rsZENCREG0.ENCREGNLI = 0
    rsZENCREG0.ENCREGMOR = ""
    rsZENCREG0.ENCREGDEV = ""
    rsZENCREG0.ENCREGCOU = 0
    rsZENCREG0.ENCREGCOB = 0
    rsZENCREG0.ENCREGMON = 0
    rsZENCREG0.ENCREGMOD = 0
    rsZENCREG0.ENCREGMOB = 0
    rsZENCREG0.ENCREGDEN = 0
    rsZENCREG0.ENCREGDRE = 0
    rsZENCREG0.ENCREGDCR = 0
    rsZENCREG0.ENCREGDVA = 0
    rsZENCREG0.ENCREGDAN = 0
    rsZENCREG0.ENCREGDAE = 0
    rsZENCREG0.ENCREGDAI = 0
    rsZENCREG0.ENCREGCL1 = ""
    rsZENCREG0.ENCREGBCL = ""
    rsZENCREG0.ENCREGICL = ""
    rsZENCREG0.ENCREGLCL = ""
    rsZENCREG0.ENCREGCOM = ""
    rsZENCREG0.ENCREGCL2 = ""
    rsZENCREG0.ENCREGBC2 = ""
    rsZENCREG0.ENCREGIC2 = ""
    rsZENCREG0.ENCREGLC2 = ""
    rsZENCREG0.ENCREGCL3 = ""
    rsZENCREG0.ENCREGBC3 = ""
    rsZENCREG0.ENCREGIC3 = ""
    rsZENCREG0.ENCREGLC3 = ""
    rsZENCREG0.ENCREGCL4 = ""
    rsZENCREG0.ENCREGBC4 = ""
    rsZENCREG0.ENCREGIC4 = ""
    rsZENCREG0.ENCREGLC4 = ""
    rsZENCREG0.ENCREGCL5 = ""
    rsZENCREG0.ENCREGBC5 = ""
    rsZENCREG0.ENCREGIC5 = ""
    rsZENCREG0.ENCREGLC5 = ""
    rsZENCREG0.ENCREGBDF = ""
    rsZENCREG0.ENCREGPAY = ""
    rsZENCREG0.ENCREGCRP = ""
    rsZENCREG0.ENCREGBAN = 0
    rsZENCREG0.ENCREGGUI = 0
    rsZENCREG0.ENCREGCHQ = 0
    rsZENCREG0.ENCREGNDE = ""
    rsZENCREG0.ENCREGEVO = ""
    rsZENCREG0.ENCREGRIB = 0
    rsZENCREG0.ENCREGRIG = 0
    rsZENCREG0.ENCREGRIC = ""
    rsZENCREG0.ENCREGESC = ""
    rsZENCREG0.ENCREGRES = 0
    rsZENCREG0.ENCREGPOR = ""
    rsZENCREG0.ENCREGINF = ""
    rsZENCREG0.ENCREGDET = ""
    rsZENCREG0.ENCREGEFF = ""
    rsZENCREG0.ENCREGUEN = 0
    rsZENCREG0.ENCREGUL1 = 0
    rsZENCREG0.ENCREGUL2 = 0
    rsZENCREG0.ENCREGCET = ""
    rsZENCREG0.ENCREGCHA = ""
    rsZENCREG0.ENCREGATG = ""
    rsZENCREG0.ENCREGEUP = ""

End Sub


