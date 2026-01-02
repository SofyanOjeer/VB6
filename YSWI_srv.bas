Attribute VB_Name = "rsZSWI"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
 
Type typeZSWIACR0
 
        SWIACRBIC               As String * 11                    ' CODE BIC
        SWIACRSNP               As String * 11                    ' CODE BIC SNP
        SWIACRTBF               As String * 11                    ' CODE BIC TBF
        SWIACRSIE               As Long                           ' CODE SIEGE
        SWIACRBQE               As String * 5                     ' CODE BANQUE
        SWIACRGUI               As String * 5                     ' CODE GUICHET
        SWIACRLI1               As String * 35                    ' LIBELLE 1
        SWIACRLI2               As String * 35                    ' LIBELLE 2
        SWIACRLI3               As String * 35                    ' LIBELLE 3
        SWIACRRSN               As String * 23                    ' RIB SNP
        SWIACRRTB               As String * 23                    ' RIB TBF

End Type
Public xZSWIACR0 As typeZSWIACR0
 
Type typeZSWIALI0
 
        SWIALIETA               As Integer                        ' ETABLISSEMENT
        SWIALIAGE               As Integer                        ' AGENCE
        SWIALISER               As String * 2                     ' SERVICE
        SWIALISSE               As String * 2                     ' SERVICE
        SWIALIMES               As String * 3                     ' TYPE MESSAGE
        SWIALINUM               As Long                           ' NUMERO INTERNE
        SWIALINEN               As String * 1                     ' NUMER ENVOI
        SWIALINLI               As Long                           ' NUMERO DE LIGNE
        SWIALIDON               As String * 512                   ' DONNE MESSAGE
        SWIALIOK                As String * 1                     ' PASSE SI OK

End Type
Public xZSWIALI0 As typeZSWIALI0
 
Type typeZSWIBUF0
 
        SWIBUFETA               As Integer                        ' ETABLISSEMENT
        SWIBUFAGE               As Integer                        ' AGENCE
        SWIBUFSER               As String * 2                     ' SERVICE
        SWIBUFSSE               As String * 2                     ' SOUS-SERVICE
        SWIBUFREF               As String * 19                    ' REF.GLO.OPERATION
        SWIBUFNLI               As Long                           ' NUMERO DE LIGNE
        SWIBUFDON               As String * 990                   ' DONNEE

End Type
Public xZSWIBUF0 As typeZSWIBUF0
 
Type typeZSWICCI0
 
        SWICCIETA               As Integer                        ' ETABLISSEMENT
        SWICCIAGE               As Integer                        ' AGENCE
        SWICCISER               As String * 2                     ' SERVICE
        SWICCISSE               As String * 2                     ' SERVICE
        SWICCIMES               As String * 3                     ' TYPE MESSAGE
        SWICCINUM               As Long                           ' NUMERO INTERNE
        SWICCINEN               As String * 1                     ' NUMER ENVOI
        SWICCINLI               As Long                           ' NUMERO DE LIGNE
        SWICCIDON               As String * 512                   ' DONNE MESSAGE
        SWICCIOK                As String * 1                     ' PASSE SI OK

End Type
Public xZSWICCI0 As typeZSWICCI0
 
Type typeZSWICLA0
 
        SWICLAETA               As Integer                        ' ETABLISSEMENT
        SWICLAAGE               As Integer                        ' AGENCE
        SWICLASER               As String * 2                     ' SERVICE
        SWICLASES               As String * 2                     ' SOUS-SERVICE
        SWICLAOPR               As String * 3                     ' CODE OPERATION
        SWICLANUM               As Long                           ' NUMERO OPERATION
        SWICLACLA               As String * 2                     ' CLASSE
        SWICLAMES               As String * 3                     ' TYPE MESSAGE
        SWICLACRI               As Long                           ' NUMRO ACCES
        SWICLAREF               As String * 16                    ' IDENTIFICATEUR
        SWICLAINT               As Long                           ' NUMERO INTERNE
        SWICLANEN               As String * 1                     ' NUMERO REENVOI

End Type
Public xZSWICLA0 As typeZSWICLA0
 
Type typeZSWICRI0
 
        SWICRIETA               As Integer                        ' ETABLISSEMENT
        SWICRICRI               As String * 25                    ' CRITERE DE SELEC.
        SWICRISEQ               As Long                           ' SEQUE DE MESSAGE
        SWICRIPRI               As String * 1                     ' PRIORITE 1-9
        SWICRINPR               As Long                           ' NUM PAR COD.PRIOR
        SWICRITYP               As String * 1                     ' TYPE LIG.1-2-3
        SWICRIDON               As String * 97                    ' DONNE DIFF.TYPES

End Type
Public xZSWICRI0 As typeZSWICRI0
 
Type typeZSWIECA0
 
        SWIECAETA               As Integer                        ' ETABLISSEMENT
        SWIECAREF               As String * 16                    ' REFERNECE
        SWIECAMES               As String * 3                     ' TYPE MESSAGE
        SWIECAPRI               As String * 2                     ' PRIORITE MESSAGE
        SWIECAEME               As String * 12                    ' EMETTEUR
        SWIECADVA               As Long                           ' DATE VALEUR
        SWIECADE1               As String * 3                     ' DEVISE 1
        SWIECAMON               As Currency                       ' DATE VALEUR
        SWIECADRE               As Long                           ' DATE RECEPTION
        SWIECAHRE               As Long                           ' HEURE RECEPTION
        SWIECAINT               As Long                           ' NUMERO INTERNE
        SWIECACET               As String * 1                     ' CODE ETAT
        SWIECAAGE               As Integer                        ' AGENCE
        SWIECASER               As String * 2                     ' SERVICE
        SWIECASSE               As String * 2                     ' SOUS SERVICE
        SWIECAUTI               As String * 10                    ' UTILISATEUR

End Type
Public xZSWIECA0 As typeZSWIECA0
 
Type typeZSWIECB0
 
        SWIECBETA               As Integer                        ' ETABLISSEMENT
        SWIECBNUM               As Long                           ' NUMERO INTERNE
        SWIECBNOR               As Long                           ' ORDRE
        SWIECBCHA               As Long                           ' CHAMP
        SWIECBIND               As String * 2                     ' INDICE
        SWIECBZON               As Long                           ' ZONE
        SWIECBSZO               As Long                           ' SOUS ZONE
        SWIECBINR               As String * 1                     ' INDICE REELL
        SWIECBVAL               As String * 65                    ' VALEUR ZONE

End Type
Public xZSWIECB0 As typeZSWIECB0
 
Type typeZSWIEHA0
 
        SWIEHAETA               As Integer                        ' ETABLISSEMENT
        SWIEHANUM               As Long                           ' NUMERO INTERNE
        SWIEHAREF               As String * 16                    ' REFERENCE
        SWIEHAMES               As String * 3                     ' TYPE MESSAGE
        SWIEHAEME               As String * 12                    ' EMETTEUR
        SWIEHADRE               As Long                           ' DATE RECEPTION
        SWIEHAHRE               As Long                           ' HEURE RECEPTION
        SWIEHAAGE               As Integer                        ' AGENCE
        SWIEHASER               As String * 2                     ' SERVICE
        SWIEHASSE               As String * 2                     ' SOUS SERVICE
        SWIEHAUTI               As String * 10                    ' UTILISATEUR
        SWIEHADTR               As Long                           ' DATE TRAITEMENT
        SWIEHAAVI               As String * 1                     ' AVIS EDITE O/N
        SWIEHADVA               As Long                           ' DATE VALEUR
        SWIEHADEV               As String * 3                     ' DEVISE 1
        SWIEHAMON               As Currency                       ' MONTANT

End Type
Public xZSWIEHA0 As typeZSWIEHA0
 
Type typeZSWIEHB0
 
        SWIEHBETA               As Integer                        ' ETABLISSEMENT
        SWIEHBNUM               As Long                           ' NUMERO INTERNE
        SWIEHBNOR               As Long                           ' ORDRE
        SWIEHBCHA               As Long                           ' CHAMP
        SWIEHBIND               As String * 2                     ' INDICE
        SWIEHBZON               As Long                           ' ZONE
        SWIEHBSZO               As Long                           ' SOUS ZONE
        SWIEHBINR               As String * 1                     ' INDICE REELL
        SWIEHBVAL               As String * 65                    ' VALEUR ZONE

End Type
Public xZSWIEHB0 As typeZSWIEHB0
 
Type typeZSWIENA0
 
        SWIENAETA               As Integer                        ' ETABLISSEMENT
        SWIENAREF               As String * 16                    ' REFERNECE
        SWIENAMES               As String * 3                     ' TYPE MESSAGE
        SWIENAPRI               As String * 2                     ' PRIORITE MESSAGE
        SWIENAEME               As String * 12                    ' EMETTEUR
        SWIENADVA               As Long                           ' DATE VALEUR
        SWIENADE1               As String * 3                     ' DEVISE 1
        SWIENAMON               As Currency                       ' DATE VALEUR
        SWIENADRE               As Long                           ' DATE RECEPTION
        SWIENAHRE               As Long                           ' HEURE RECEPTION
        SWIENAINT               As Long                           ' NUMERO INTERNE
        SWIENACET               As String * 1                     ' CODE ETAT
        SWIENAAGE               As Integer                        ' AGENCE
        SWIENASER               As String * 2                     ' SERVICE
        SWIENASSE               As String * 2                     ' SOUS SERVICE
        SWIENAUTI               As String * 10                    ' UTILISATEUR

End Type
Public xZSWIENA0 As typeZSWIENA0
 
Type typeZSWIENB0
        SWIENBETA               As Integer                        ' ETABLISSEMENT
        SWIENBNUM               As Long                           ' NUMERO INTERNE
        SWIENBNOR               As Long                           ' ORDRE
        SWIENBCHA               As Long                           ' CHAMP
        SWIENBIND               As String * 2                     ' INDICE
        SWIENBZON               As Long                           ' ZONE
        SWIENBSZO               As Long                           ' SOUS ZONE
        SWIENBINR               As String * 1                     ' INDICE REELL
        SWIENBVAL               As String * 65                    ' VALEUR ZONE

End Type
Public xZSWIENB0 As typeZSWIENB0
 
Type typeZSWIENI0
 
        SWIENIETA               As Integer                        ' ETABLISSEMENT
        SWIENIAGE               As Integer                        ' AGENCE
        SWIENISER               As String * 2                     ' SERVICE
        SWIENISSE               As String * 2                     ' SERVICE
        SWIENIMES               As String * 3                     ' TYPE MESSAGE
        SWIENINUM               As Long                           ' NUMERO INTERNE
        SWIENINEN               As String * 1                     ' NUMER ENVOI
        SWIENINLI               As Long                           ' NUMERO DE LIGNE
        SWIENIDON               As String * 250                   ' DONNE MESSAGE
        SWIENIOK                As String * 1                     ' PASSE SI OK

End Type
Public xZSWIENI0 As typeZSWIENI0
 
Type typeZSWIEVC0
 
        SWIEVCDON               As String * 512                   ' DONNE MESSAGE

End Type
Public xZSWIEVC0 As typeZSWIEVC0
 
Type typeZSWIEVI0
 
        SWIEVIETA               As Integer                        ' ETABLISSEMENT
        SWIEVIAGE               As Integer                        ' AGENCE
        SWIEVISER               As String * 2                     ' SERVICE
        SWIEVISSE               As String * 2                     ' SERVICE
        SWIEVIMES               As String * 3                     ' TYPE MESSAGE
        SWIEVINUM               As Long                           ' NUMERO INTERNE
        SWIEVINEN               As String * 1                     ' NUMER ENVOI
        SWIEVINLI               As Long                           ' NUMERO DE LIGNE
        SWIEVIDON               As String * 512                   ' DONNE MESSAGE
        SWIEVIOK                As String * 1                     ' PASSE SI OK

End Type
Public xZSWIEVI0 As typeZSWIEVI0
 
Type typeZSWIFTA0
 
        SWIFTAETA               As Integer                        ' ETABLISSEMENT
        SWIFTAREF               As String * 16                    ' REFERNECE
        SWIFTANEN               As String * 1                     ' NUMERO DE RENVOI
        SWIFTAPRI               As String * 2                     ' CODE PROIRITE
        SWIFTAMES               As String * 3                     ' TYPE MESSAGE
        SWIFTADOR               As String * 12                    ' DONNEUR ORDRE
        SWIFTADES               As String * 12                    ' DESTINATAIRE
        SWIFTADVA               As Long                           ' DATE VALEUR
        SWIFTADE1               As String * 3                     ' DEVISE 1
        SWIFTAMON               As Currency                       ' DATE VALEUR
        SWIFTADE2               As String * 3                     ' DEVISE 2
        SWIFTADEN               As Long                           ' DATE ENVOI
        SWIFTAHEN               As Long                           ' HEURE ENVOI
        SWIFTACOM               As String * 1                     ' COMPLET
        SWIFTATES               As String * 1                     ' TEST OU REEL
        SWIFTASUP               As String * 1                     ' SUPPRIME
        SWIFTAVAL               As String * 1                     ' TOP VALIDATION
        SWIFTAAGE               As Integer                        ' AGENCE
        SWIFTASER               As String * 2                     ' SERVICE
        SWIFTASSE               As String * 2                     ' SOUS SERVICE
        SWIFTAUTI               As String * 10                    ' UTILISATEUR
        SWIFTANUM               As Long                           ' NUMERO INTERNE
        SWIFTAUT1               As String * 10                    ' UTILISA SAISIE
        SWIFTAPVA               As String * 1                     ' 1ERE VALIDATION
        SWIFTAUT2               As String * 10                    ' UTILISA 1ER VALID

End Type
Public xZSWIFTA0 As typeZSWIFTA0
 
Type typeZSWIFTB0
 
        SWIFTBETA               As Integer                        ' ETABLISSEMENT
        SWIFTBNUM               As Long                           ' NUMERO INTERNE
        SWIFTBNEN               As Long                           ' NUMERO ENVOI
        SWIFTBNLI               As Long                           ' NUMERO LIGNE
        SWIFTBDET               As String * 70                    ' DETAIL

End Type
Public xZSWIFTB0 As typeZSWIFTB0
 
Type typeZSWIFTC0

        SWIFTCETA               As Integer                        ' ETABLISSEMENT
        SWIFTCNUM               As Long                           ' NUMERO INTERNE
        SWIFTCNEN               As Long                           ' NUMERO ENVOI
        SWIFTCNLI               As Long                           ' NUMERO LIGNE
        SWIFTCNSE               As String * 40                    ' NUMERO SEQUENCE
        SWIFTCSOC               As Long                           ' NUM OCC SEQUE
        SWIFTCNCH               As Long                           ' NUMERO CHAMP
        SWIFTCCOC               As Long                           ' NUM OCC CHAMP
        SWIFTCNLC               As Long                           ' NUMERO LIGNE CHAM
        SWIFTCSEQ               As String * 2                     ' DESCRIP SEQUENCE
        SWIFTCCHA               As String * 2                     ' DESCRIP CHAMP
        SWIFTCILI               As String * 1                     ' INDICATEUR DEB
        SWIFTCFAC               As String * 1                     ' FACULTATIF
        SWIFTCSIG               As String * 1                     ' SIGNE COMPLET
        SWIFTCSMA               As Long                           ' OCCUR SEQ MAXIMUM
        SWIFTCCMA               As Long                           ' OCCUR CHA MAXIMUM
        SWIFTCSMI               As Long                           ' OCCUR SEQ MINIMUM
        SWIFTCCMI               As Long                           ' OCCUR CHA MINIMUM

End Type
Public xZSWIFTC0 As typeZSWIFTC0
 
Type typeZSWIGRN0
 
        SWIGRNETA               As Integer                        ' ETABLISSEMENT
        SWIGRNGRP               As String * 6                     ' GRP NATURE
        SWIGRNORD               As Long                           ' NUMERO PAR NATURE
        SWIGRNNAT               As String * 6                     ' NATURE

End Type
Public xZSWIGRN0 As typeZSWIGRN0
 
Type typeZSWIHIA0
 
        SWIHIAETA               As Integer                        ' ETABLISSEMENT
        SWIHIAREF               As String * 16                    ' REFERNECE
        SWIHIANEN               As String * 1                     ' NUMERO DE RENVOI
        SWIHIAPRI               As String * 2                     ' CODE PROIRITE
        SWIHIAMES               As String * 3                     ' TYPE MESSAGE
        SWIHIADOR               As String * 12                    ' DONNEUR ORDRE
        SWIHIADES               As String * 12                    ' DESTINATAIRE
        SWIHIADVA               As Long                           ' DATE VALEUR
        SWIHIADE1               As String * 3                     ' DEVISE 1
        SWIHIAMON               As Currency                       ' MONTANT
        SWIHIADE2               As String * 3                     ' DEVISE 2
        SWIHIADEN               As Long                           ' DATE ENVOI
        SWIHIAHEN               As Long                           ' HEURE ENVOI
        SWIHIACOM               As String * 1                     ' COMPLET
        SWIHIATES               As String * 1                     ' TEST OU REEL
        SWIHIASUP               As String * 1                     ' SUPPRIME
        SWIHIAVAL               As String * 1                     ' TOP VALIDATION
        SWIHIAAGE               As Integer                        ' AGENCE
        SWIHIASER               As String * 2                     ' SERVICE
        SWIHIASSE               As String * 2                     ' SOUS SERVICE
        SWIHIAUTI               As String * 10                    ' UTILISATEUR
        SWIHIANUM               As Long                           ' NUMERO INTERNE
        SWIHIAUT1               As String * 10                    ' UTILISA SAISIE
        SWIHIAPVA               As String * 1                     ' 1ERE VALIDATION
        SWIHIAUT2               As String * 10                    ' UTILISA 1ER VALID

End Type
Public xZSWIHIA0 As typeZSWIHIA0
 
Type typeZSWIHIB0
 
        SWIHIBETA               As Integer                        ' ETABLISSEMENT
        SWIHIBNUM               As Long                           ' NUMERO INTERNE
        SWIHIBNEN               As Long                           ' NUMERO ENVOI
        SWIHIBNLI               As Long                           ' NUMERO LIGNE
        SWIHIBDET               As String * 70                    ' DETAIL

End Type
Public xZSWIHIB0 As typeZSWIHIB0
 
Type typeZSWIHIC0

        SWIHICETA               As Integer                        ' ETABLISSEMENT
        SWIHICNUM               As Long                           ' NUMERO INTERNE
        SWIHICNEN               As Long                           ' NUMERO ENVOI
        SWIHICNLI               As Long                           ' NUMERO LIGNE
        SWIHICNSE               As String * 40                    ' NUMERO SEQUENCE
        SWIHICSOC               As Long                           ' NUM OCC SEQUE
        SWIHICNCH               As Long                           ' NUMERO CHAMP
        SWIHICCOC               As Long                           ' NUM OCC CHAMP
        SWIHICNLC               As Long                           ' NUMERO LIGNE CHAM
        SWIHICSEQ               As String * 2                     ' DESCRIP SEQUENCE
        SWIHICCHA               As String * 2                     ' DESCRIP CHAMP
        SWIHICILI               As String * 1                     ' INDICATEUR DEB
        SWIHICFAC               As String * 1                     ' FACULTATIF
        SWIHICSIG               As String * 1                     ' SIGNE COMPLET
        SWIHICSMA               As Long                           ' OCCUR SEQ MAXIMUM
        SWIHICCMA               As Long                           ' OCCUR CHA MAXIMUM
        SWIHICSMI               As Long                           ' OCCUR SEQ MINIMUM
        SWIHICCMI               As Long                           ' OCCUR CHA MINIMUM

End Type
Public xZSWIHIC0 As typeZSWIHIC0
 
Type typeZSWIHIT0
 
        SWIHITETA               As Integer                        ' ETABLISSEMENT
        SWIHITNUM               As Long                           ' NUMERO INTERNE
        SWIHITNEN               As Long                           ' NUMERO ENVOI
        SWIHITNSE               As String * 40                    ' NUMERO SEQUENCE
        SWIHITSEQ               As String * 2                     ' SEQUENCE
        SWIHITOSE               As Long                           ' OCCURENCE SEQUE.
        SWIHITCHA               As Long                           ' CHAMP
        SWIHITOCH               As Long                           ' OCCURENCE CHAMP
        SWIHITIND               As String * 2                     ' INDICE
        SWIHITZON               As Long                           ' ZONE
        SWIHITOZO               As Long                           ' OCCURENCE ZONE
        SWIHITSZO               As Long                           ' SOUS ZONE
        SWIHITOSZ               As Long                           ' OCCURENCE S-ZONE
        SWIHITCON               As Long                           ' COMPTEUR ENREGIS
        SWIHITCOM               As String * 1                     ' COMPLET
        SWIHITVAL               As String * 65                    ' VALEUR ZONE

End Type
Public xZSWIHIT0 As typeZSWIHIT0
 
Type typeZSWIJOB0

        SWIJOBETA               As Integer                        ' ETABLISSEMENT
        SWIJOBPRO               As String * 10                    ' PROCEDURE
        SWIJOBJOB               As String * 10                    ' TRAVAIL
        SWIJOBUSR               As String * 10                    ' UTILISATEUR
        SWIJOBNBR               As String * 6                     ' N° TRAVAIL
        SWIJOBDLA               As Long                           ' DATE  LANCEMENT
        SWIJOBHLA               As Long                           ' HEURE LANCEMENT
        SWIJOBENV               As Long                           ' NBR MESS.A EVOYER
        SWIJOBTER               As Long                           ' HEURE TERMINAISON
        SWIJOBACT               As String * 1                     ' ACTIF O/N

End Type
Public xZSWIJOB0 As typeZSWIJOB0
 
Type typeZSWIMEA0
  
        SWIMEAETA               As Integer                        ' ETABLISSEMENT
        SWIMEANUM               As Long                           ' NUMERO INTERNE
        SWIMEAREF               As String * 16                    ' REFERENCE
        SWIMEAMES               As String * 3                     ' TYPE MESSAGE
        SWIMEAEME               As String * 12                    ' EMETTEUR
        SWIMEADRE               As Long                           ' DATE RECEPTION
        SWIMEAHRE               As Long                           ' HEURE RECEPTION
        SWIMEAAGE               As Integer                        ' AGENCE
        SWIMEASER               As String * 2                     ' SERVICE
        SWIMEASSE               As String * 2                     ' SOUS SERVICE
        SWIMEAUTI               As String * 10                    ' UTILISATEUR
        SWIMEADTR               As Long                           ' DATE TRAITEMENT
        SWIMEAAVI               As String * 1                     ' AVIS EDITE O/N
        SWIMEADVA               As Long                           ' DATE VALEUR
        SWIMEADEV               As String * 3                     ' DEVISE 1
        SWIMEAMON               As Currency                       ' MONTANT

End Type
Public xZSWIMEA0 As typeZSWIMEA0
 
Type typeZSWIMEB0
 
        SWIMEBETA               As Integer                        ' ETABLISSEMENT
        SWIMEBNUM               As Long                           ' NUMERO INTERNE
        SWIMEBNOR               As Long                           ' ORDRE
        SWIMEBCHA               As Long                           ' CHAMP
        SWIMEBIND               As String * 2                     ' INDICE
        SWIMEBZON               As Long                           ' ZONE
        SWIMEBSZO               As Long                           ' SOUS ZONE
        SWIENBINR               As String * 1                     ' INDICE REELL
        SWIMEBVAL               As String * 65                    ' VALEUR ZONE

End Type
Public xZSWIMEB0 As typeZSWIMEB0
 
Type typeZSWIMEC0
 
        SWIMECETA               As Integer                        ' ETABLISSEMENT
        SWIMECNUM               As Long                           ' NUMERO INTERNE
        SWIMECNOR               As Long                           ' N° ORDRE DU CHAMP
        SWIMECCOM               As String * 50                    ' COMMENTA. AJOUTEE

End Type
Public xZSWIMEC0 As typeZSWIMEC0
 
Type typeZSWIMEM0

        SWIMEMETA               As Integer                        ' ETABLISSEMENT
        SWIMEMNUM               As Long                           ' NUMERO INTERNE
        SWIMEMNOR               As Long                           ' N° ORDRE DU CHAMP
        SWIMEMMEG               As String * 3                     ' MESSAGE GENERE
        SWIMEMREG               As String * 16                    ' REFERENCE GENEREE

End Type
Public xZSWIMEM0 As typeZSWIMEM0
 
Type typeZSWIMEO0
 
        SWIMEOETA               As Integer                        ' ETABLISSEMENT
        SWIMEONUM               As Long                           ' NUMERO INTERNE
        SWIMEONOR               As Long                           ' N° ORDRE DU CHAMP
        SWIMEOOPR               As String * 6                     ' TYPE OPERATION
        SWIMEONAT               As String * 6                     ' NATURE OPERATION
        SWIMEONOP               As Long                           ' NUMERO OPERATION

End Type
Public xZSWIMEO0 As typeZSWIMEO0
 
Type typeZSWIRAL0
 
        SWIRALDON               As String * 512                   ' DONNE MESSAGE
        SWIRALETA               As Integer                        '
        SWIRALMES               As String * 3                     '

End Type
Public xZSWIRAL0 As typeZSWIRAL0
 
Type typeZSWIRDE0
 
        SWIRDEETA               As Integer                        ' ETABLISSEMENT
        SWIRDEBIC               As String * 12                    ' BIC
        SWIRDENUM               As Long                           ' REFERENCE CLIENT
        SWIRDECOM               As String * 35                    ' COMPTE
        SWIRDEDAT               As Long                           ' DATE DERNIERE EMI
        SWIRDECOU               As Long                           ' COMPTEUR QUOTIDIE

End Type
Public xZSWIRDE0 As typeZSWIRDE0
 
Type typeZSWIREC0
 
        SWIRECETA               As Integer                        ' ETABLISSEMENT
        SWIRECNUM               As Long                           ' NUMERO INTERNE
        SWIRECNLI               As Long                           ' NUMERO DE LIGNE
        SWIRECMES               As String * 3                     ' NUMERO DE MESSAGE
        SWIRECDON               As String * 250                   ' DONNE MESSAGE

End Type
Public xZSWIREC0 As typeZSWIREC0
 
Type typeZSWIRED0
  
        SWIREDETA               As Integer                        ' ETABLISSEMENT
        SWIREDAGE               As Integer                        ' AGENCE
        SWIREDSER               As String * 2                     ' SERVICE
        SWIREDSSE               As String * 2                     ' SOUS SERVICE
        SWIREDREF               As String * 16                    ' REFERENCE
        SWIREDME1               As String * 3                     ' TYPE MESSAGE (1)
        SWIREDME2               As String * 3                     ' TYPE MESSAGE (2)
        SWIREDEM1               As String * 12                    ' EMETTEUR (1)
        SWIREDEM2               As String * 12                    ' EMETTEUR (2)
        SWIREDNU1               As Long                           ' NUMERO INTERNE(1)
        SWIREDNU2               As Long                           ' NUMERO INTERNE(2)
        SWIREDDAT               As Long                           ' DATE TRAITEMENT
        SWIREDAVI               As String * 1                     ' EDIT OU NON

End Type
Public xZSWIRED0 As typeZSWIRED0
 
Type typeZSWIRST0
 
        SWIRSTDON               As String * 93                    ' DONNE MESSAGE
        SWIRSTETA               As Integer                        '
        SWIRSTMES               As String * 3                     '

End Type
Public xZSWIRST0 As typeZSWIRST0
 
Type typeZSWISCA0
 
        SWISCAETA               As Integer                        ' ETABLISSEMENT
        SWISCAREF               As String * 16                    ' REFERNECE
        SWISCANEN               As String * 1                     ' NUMERO DE RENVOI
        SWISCAPRI               As String * 2                     ' CODE PROIRITE
        SWISCAMES               As String * 3                     ' TYPE MESSAGE
        SWISCADOR               As String * 12                    ' DONNEUR ORDRE
        SWISCADES               As String * 12                    ' DESTINATAIRE
        SWISCADVA               As Long                           ' DATE VALEUR
        SWISCADE1               As String * 3                     ' DEVISE 1
        SWISCAMON               As Currency                       ' MONTANT
        SWISCADE2               As String * 3                     ' DEVISE 2
        SWISCADEN               As Long                           ' DATE ENVOI
        SWISCAHEN               As Long                           ' HEURE ENVOI
        SWISCACOM               As String * 1                     ' COMPLET
        SWISCATES               As String * 1                     ' TEST OU REEL
        SWISCASUP               As String * 1                     ' SUPPRIME
        SWISCAVAL               As String * 1                     ' TOP VALIDATION
        SWISCAAGE               As Integer                        ' AGENCE
        SWISCASER               As String * 2                     ' SERVICE
        SWISCASSE               As String * 2                     ' SOUS SERVICE
        SWISCAUTI               As String * 10                    ' UTILISATEUR
        SWISCANUM               As Long                           ' NUMERO INTERNE
        SWISCAUT1               As String * 10                    ' UTILISA SAISIE
        SWISCAPVA               As String * 1                     ' 1ERE VALIDATION
        SWISCAUT2               As String * 10                    ' UTILISA 1ER VALID

End Type
Public xZSWISCA0 As typeZSWISCA0
 
Type typeZSWISCB0
 
        SWISCBETA               As Integer                        ' ETABLISSEMENT
        SWISCBNUM               As Long                           ' NUMERO INTERNE
        SWISCBNEN               As Long                           ' NUMERO ENVOI
        SWISCBNLI               As Long                           ' NUMERO LIGNE
        SWISCBDET               As String * 70                    ' DETAIL

End Type
Public xZSWISCB0 As typeZSWISCB0
 
Type typeZSWISCC0
 
        SWISCCETA               As Integer                        ' ETABLISSEMENT
        SWISCCNUM               As Long                           ' NUMERO INTERNE
        SWISCCNEN               As Long                           ' NUMERO ENVOI
        SWISCCNLI               As Long                           ' NUMERO LIGNE
        SWISCCNSE               As String * 40                    ' NUMERO SEQUENCE
        SWISCCSOC               As Long                           ' NUM OCC SEQUE
        SWISCCNCH               As Long                           ' NUMERO CHAMP
        SWISCCCOC               As Long                           ' NUM OCC CHAMP
        SWISCCNLC               As Long                           ' NUMERO LIGNE CHAM
        SWISCCSEQ               As String * 2                     ' DESCRIP SEQUENCE
        SWISCCCHA               As String * 2                     ' DESCRIP CHAMP
        SWISCCILI               As String * 1                     ' INDICATEUR DEB
        SWISCCFAC               As String * 1                     ' FACULTATIF
        SWISCCSIG               As String * 1                     ' SIGNE COMPLET
        SWISCCSMA               As Long                           ' OCCUR SEQ MAXIMUM
        SWISCCCMA               As Long                           ' OCCUR CHA MAXIMUM
        SWISCCSMI               As Long                           ' OCCUR SEQ MINIMUM
        SWISCCCMI               As Long                           ' OCCUR CHA MINIMUM

End Type
Public xZSWISCC0 As typeZSWISCC0
 
Type typeZSWISCT0
 
        SWISCTETA               As Integer                        ' ETABLISSEMENT
        SWISCTNUM               As Long                           ' NUMERO INTERNE
        SWISCTNEN               As Long                           ' NUMERO ENVOI
        SWISCTNSE               As String * 40                    ' NUMERO SEQUENCE
        SWISCTSEQ               As String * 2                     ' SEQUENCE
        SWISCTOSE               As Long                           ' OCCURENCE SEQUE.
        SWISCTCHA               As Long                           ' CHAMP
        SWISCTOCH               As Long                           ' OCCURENCE CHAMP
        SWISCTIND               As String * 2                     ' INDICE
        SWISCTZON               As Long                           ' ZONE
        SWISCTOZO               As Long                           ' OCCURENCE ZONE
        SWISCTSZO               As Long                           ' SOUS ZONE
        SWISCTOSZ               As Long                           ' OCCURENCE S-ZONE
        SWISCTCON               As Long                           ' COMPTEUR ENREGIS
        SWISCTCOM               As String * 1                     ' COMPLET
        SWISCTVAL               As String * 65                    ' VALEUR ZONE

End Type
Public xZSWISCT0 As typeZSWISCT0
 
Type typeZSWISHA0
 
        SWISHAETA               As Integer                        ' ETABLISSEMENT
        SWISHAREF               As String * 16                    ' REFERNECE
        SWISHANEN               As String * 1                     ' NUMERO DE RENVOI
        SWISHAPRI               As String * 2                     ' CODE PROIRITE
        SWISHAMES               As String * 3                     ' TYPE MESSAGE
        SWISHADOR               As String * 12                    ' DONNEUR ORDRE
        SWISHADES               As String * 12                    ' DESTINATAIRE
        SWISHADVA               As Long                           ' DATE VALEUR
        SWISHADE1               As String * 3                     ' DEVISE 1
        SWISHAMON               As Currency                       ' MONTANT
        SWISHADE2               As String * 3                     ' DEVISE 2
        SWISHADEN               As Long                           ' DATE ENVOI
        SWISHAHEN               As Long                           ' HEURE ENVOI
        SWISHACOM               As String * 1                     ' COMPLET
        SWISHATES               As String * 1                     ' TEST OU REEL
        SWISHASUP               As String * 1                     ' SUPPRIME
        SWISHAVAL               As String * 1                     ' TOP VALIDATION
        SWISHAAGE               As Integer                        ' AGENCE
        SWISHASER               As String * 2                     ' SERVICE
        SWISHASSE               As String * 2                     ' SOUS SERVICE
        SWISHAUTI               As String * 10                    ' UTILISATEUR
        SWISHANUM               As Long                           ' NUMERO INTERNE
        SWISHAUT1               As String * 10                    ' UTILISA SAISIE
        SWISHAPVA               As String * 1                     ' 1ERE VALIDATION
        SWISHAUT2               As String * 10                    ' UTILISA 1ER VALID

End Type
Public xZSWISHA0 As typeZSWISHA0
 
Type typeZSWISHB0
 
        SWISHBETA               As Integer                        ' ETABLISSEMENT
        SWISHBNUM               As Long                           ' NUMERO INTERNE
        SWISHBNEN               As Long                           ' NUMERO ENVOI
        SWISHBNLI               As Long                           ' NUMERO LIGNE
        SWISHBDET               As String * 70                    ' DETAIL

End Type
Public xZSWISHB0 As typeZSWISHB0
 
Type typeZSWISHC0
 
        SWISHCETA               As Integer                        ' ETABLISSEMENT
        SWISHCNUM               As Long                           ' NUMERO INTERNE
        SWISHCNEN               As Long                           ' NUMERO ENVOI
        SWISHCNLI               As Long                           ' NUMERO LIGNE
        SWISHCNSE               As String * 40                    ' NUMERO SEQUENCE
        SWISHCSOC               As Long                           ' NUM OCC SEQUE
        SWISHCNCH               As Long                           ' NUMERO CHAMP
        SWISHCCOC               As Long                           ' NUM OCC CHAMP
        SWISHCNLC               As Long                           ' NUMERO LIGNE CHAM
        SWISHCSEQ               As String * 2                     ' DESCRIP SEQUENCE
        SWISHCCHA               As String * 2                     ' DESCRIP CHAMP
        SWISHCILI               As String * 1                     ' INDICATEUR DEB
        SWISHCFAC               As String * 1                     ' FACULTATIF
        SWISHCSIG               As String * 1                     ' SIGNE COMPLET
        SWISHCSMA               As Long                           ' OCCUR SEQ MAXIMUM
        SWISHCCMA               As Long                           ' OCCUR CHA MAXIMUM
        SWISHCSMI               As Long                           ' OCCUR SEQ MINIMUM
        SWISHCCMI               As Long                           ' OCCUR CHA MINIMUM

End Type
Public xZSWISHC0 As typeZSWISHC0
 
Type typeZSWISHT0
 
        SWISHTETA               As Integer                        ' ETABLISSEMENT
        SWISHTNUM               As Long                           ' NUMERO INTERNE
        SWISHTNEN               As Long                           ' NUMERO ENVOI
        SWISHTNSE               As String * 40                    ' NUMERO SEQUENCE
        SWISHTSEQ               As String * 2                     ' SEQUENCE
        SWISHTOSE               As Long                           ' OCCURENCE SEQUE.
        SWISHTCHA               As Long                           ' CHAMP
        SWISHTOCH               As Long                           ' OCCURENCE CHAMP
        SWISHTIND               As String * 2                     ' INDICE
        SWISHTZON               As Long                           ' ZONE
        SWISHTOZO               As Long                           ' OCCURENCE ZONE
        SWISHTSZO               As Long                           ' SOUS ZONE
        SWISHTOSZ               As Long                           ' OCCURENCE S-ZONE
        SWISHTCON               As Long                           ' COMPTEUR ENREGIS
        SWISHTCOM               As String * 1                     ' COMPLET
        SWISHTVAL               As String * 65                    ' VALEUR ZONE

End Type
Public xZSWISHT0 As typeZSWISHT0
 
Type typeZSWISRC0
 
        AIDCLICOD               As Currency                       ' 111
        AIDCLIPRG               As Long                           ' 222
        AIDCLIFMT               As String * 80                    ' 333

End Type
Public xZSWISRC0 As typeZSWISRC0
 
Type typeZSWITAB0
 
        SWITABETA               As Integer                        ' ETABLISSEMENT
        SWITABNUM               As Long                           ' NUMERO TABLE
        SWITABARG               As String * 25                    ' ARGUMENT
        SWITABLO1               As String * 12                    ' LOGIQUE 1
        SWITABLO2               As String * 12                    ' LOGIQUE 2
        SWITABDON               As String * 300                   ' DONNEES

End Type
Public xZSWITAB0 As typeZSWITAB0
 
Type typeZSWITEM0
 
        SWITEMETA               As Integer                        ' ETABLISSEMENT
        SWITEMNUM               As Long                           ' NUMERO INTERNE
        SWITEMNEN               As Long                           ' NUMERO ENVOI
        SWITEMNSE               As String * 40                    ' NUMERO SEQUENCE
        SWITEMSEQ               As String * 2                     ' SEQUENCE
        SWITEMOSE               As Long                           ' OCCURENCE SEQUE.
        SWITEMCHA               As Long                           ' CHAMP
        SWITEMOCH               As Long                           ' OCCURENCE CHAMP
        SWITEMIND               As String * 2                     ' INDICE
        SWITEMZON               As Long                           ' ZONE
        SWITEMOZO               As Long                           ' OCCURENCE ZONE
        SWITEMSZO               As Long                           ' SOUS ZONE
        SWITEMOSZ               As Long                           ' OCCURENCE S-ZONE
        SWITEMCON               As Long                           ' COMPTEUR ENREGIS
        SWITEMCOM               As String * 1                     ' COMPLET
        SWITEMVAL               As String * 65                    ' VALEUR ZONE

End Type
Public xZSWITEM0 As typeZSWITEM0
Public Function srvYSWIFTB0_Sql_Insert(lYSWIFTB0 As typeZSWIFTB0, lSql As String)
On Error GoTo Error_Handler
srvYSWIFTB0_Sql_Insert = Null
lSql = "Insert into " & paramIBM_Library_SAB & ".ZSWIFTB0 " _
        & "(SWIFTBETA,SWIFTBNUM,SWIFTBNEN,SWIFTBNLI,SWIFTBDET) values ( " _
        & lYSWIFTB0.SWIFTBETA _
        & "," & lYSWIFTB0.SWIFTBNUM _
        & "," & lYSWIFTB0.SWIFTBNEN _
        & "," & lYSWIFTB0.SWIFTBNLI _
        & ",'" & lYSWIFTB0.SWIFTBDET & "')"
        
Exit Function
Error_Handler:
srvYSWIFTB0_Sql_Insert = Error
End Function

Public Function srvYSWIHIC0_Sql_Sauvegarde(lYSWIFTC0 As typeZSWIFTC0, lSql As String)
On Error GoTo Error_Handler
srvYSWIHIC0_Sql_Sauvegarde = Null
lSql = "Insert into " & paramIBM_Library_SAB & ".ZSWIHIC0 " _
          & "(SWIHICETA, SWIHICNUM , SWIHICNEN, SWIHICNLI, SWIHICNSE, SWIHICSOC, SWIHICNCH, SWIHICCOC, SWIHICNLC," _
          & " SWIHICSEQ, SWIHICCHA, SWIHICILI, SWIHICFAC, SWIHICSIG, SWIHICSMA, SWIHICCMA, SWIHICSMI, SWIHICCMI" _
          & " ) values ( " _
        & lYSWIFTC0.SWIFTCETA _
        & "," & lYSWIFTC0.SWIFTCNUM _
        & "," & lYSWIFTC0.SWIFTCNEN _
        & "," & lYSWIFTC0.SWIFTCNLI _
        & ",'" & lYSWIFTC0.SWIFTCNSE & "'" _
        & "," & lYSWIFTC0.SWIFTCSOC _
        & "," & lYSWIFTC0.SWIFTCNCH _
        & "," & lYSWIFTC0.SWIFTCCOC _
        & "," & lYSWIFTC0.SWIFTCNLC _
        & ",'" & lYSWIFTC0.SWIFTCSEQ & "'" _
        & ",'" & lYSWIFTC0.SWIFTCCHA & "'" _
        & ",'" & lYSWIFTC0.SWIFTCILI & "'" _
        & ",'" & lYSWIFTC0.SWIFTCFAC & "'" _
        & ",'" & lYSWIFTC0.SWIFTCSIG & "'" _
        & "," & lYSWIFTC0.SWIFTCSMA _
        & "," & lYSWIFTC0.SWIFTCCMA _
        & "," & lYSWIFTC0.SWIFTCSMI _
        & "," & lYSWIFTC0.SWIFTCCMI & ")"

Exit Function
Error_Handler:
srvYSWIHIC0_Sql_Sauvegarde = Error
End Function

Public Function srvYSWIHIA0_Sql_Sauvegarde(lYSWIFTA0 As typeZSWIFTA0, lSql As String)
On Error GoTo Error_Handler
srvYSWIHIA0_Sql_Sauvegarde = Null
lSql = "Insert into " & paramIBM_Library_SAB & ".ZSWIHIA0 " _
          & "(SWIHIAETA, SWIHIAREF , SWIHIANEN, SWIHIAPRI, SWIHIAMES, SWIHIADOR, SWIHIADES, SWIHIADVA, SWIHIADE1," _
          & " SWIHIAMON, SWIHIADE2, SWIHIADEN, SWIHIAHEN, SWIHIACOM, SWIHIATES, SWIHIASUP, SWIHIAVAL, SWIHIAAGE," _
          & " SWIHIASER, SWIHIASSE, SWIHIAUTI, SWIHIANUM, SWIHIAUT1, SWIHIAPVA, SWIHIAUT2 ) values ( " _
        & lYSWIFTA0.SWIFTAETA _
        & ",'" & lYSWIFTA0.SWIFTAREF & "','" & lYSWIFTA0.SWIFTANEN & "'" _
        & ",'" & lYSWIFTA0.SWIFTAPRI & "','" & lYSWIFTA0.SWIFTAMES & "'" _
        & ",'" & lYSWIFTA0.SWIFTADOR & "','" & lYSWIFTA0.SWIFTADES & "'" _
        & "," & lYSWIFTA0.SWIFTADVA & ",'" & lYSWIFTA0.SWIFTADE1 & "'" _
        & "," & lYSWIFTA0.SWIFTAMON & ",'" & lYSWIFTA0.SWIFTADE2 & "'" _
        & "," & lYSWIFTA0.SWIFTADEN _
        & "," & lYSWIFTA0.SWIFTAHEN _
        & ",'" & lYSWIFTA0.SWIFTACOM & "'" _
        & ",'" & lYSWIFTA0.SWIFTATES & "'" _
        & ",'" & lYSWIFTA0.SWIFTASUP & "'" _
        & ",'" & lYSWIFTA0.SWIFTAVAL & "'" _
        & "," & lYSWIFTA0.SWIFTAAGE _
        & ",'" & lYSWIFTA0.SWIFTASER & "'" _
        & ",'" & lYSWIFTA0.SWIFTASSE & "'" _
        & ",'" & lYSWIFTA0.SWIFTAUTI & "'" _
        & "," & lYSWIFTA0.SWIFTANUM _
        & ",'" & lYSWIFTA0.SWIFTAUT1 & "'" _
        & ",'" & lYSWIFTA0.SWIFTAPVA & "'" _
        & ",'" & lYSWIFTA0.SWIFTAUT2 & "')"

Exit Function
Error_Handler:
srvYSWIHIA0_Sql_Sauvegarde = Error


End Function

Public Function srvYSWIHIT0_Sql_Sauvegarde(lYSWITEM0 As typeZSWITEM0, lSql As String)
On Error GoTo Error_Handler
srvYSWIHIT0_Sql_Sauvegarde = Null
Dim K As Integer
' Supprimer le charactère ' dans la requête SQL
Do
    K = InStr(lYSWITEM0.SWITEMVAL, "'")
    If K > 0 Then Mid$(lYSWITEM0.SWITEMVAL, K, 1) = " "
Loop Until K = 0

lSql = "Insert into " & paramIBM_Library_SAB & ".ZSWIHIT0 " _
          & "(SWIHITETA, SWIHITNUM , SWIHITNEN,  SWIHITNSE, SWIHITSEQ, SWIHITOSE, SWIHITCHA, SWIHITOCH," _
          & " SWIHITIND, SWIHITZON, SWIHITOZO, SWIHITSZO, SWIHITOSZ, SWIHITCON, SWIHITCOM, SWIHITVAL" _
          & " ) values ( " _
        & lYSWITEM0.SWITEMETA _
        & "," & lYSWITEM0.SWITEMNUM _
        & "," & lYSWITEM0.SWITEMNEN _
        & ",'" & lYSWITEM0.SWITEMNSE & "'" _
        & ",'" & lYSWITEM0.SWITEMSEQ & "'" _
        & "," & lYSWITEM0.SWITEMOSE _
        & "," & lYSWITEM0.SWITEMCHA _
        & "," & lYSWITEM0.SWITEMOCH _
        & ",'" & lYSWITEM0.SWITEMIND & "'" _
        & "," & lYSWITEM0.SWITEMZON _
        & "," & lYSWITEM0.SWITEMOZO _
        & "," & lYSWITEM0.SWITEMSZO _
        & "," & lYSWITEM0.SWITEMOSZ _
        & "," & lYSWITEM0.SWITEMCON _
        & ",'" & lYSWITEM0.SWITEMCOM & "'" _
        & ",'" & lYSWITEM0.SWITEMVAL & "')"

Exit Function
Error_Handler:
srvYSWIHIT0_Sql_Sauvegarde = Error
End Function


Public Function srvYSWIFTA0_Sql_Restauration(lYSWIHIA0 As typeZSWIHIA0, lSql As String)
On Error GoTo Error_Handler
srvYSWIFTA0_Sql_Restauration = Null
lSql = "Insert into " & paramIBM_Library_SAB & ".ZSWIFTA0 " _
          & "(SWIFTAETA, SWIFTAREF , SWIFTANEN, SWIFTAPRI, SWIFTAMES, SWIFTADOR, SWIFTADES, SWIFTADVA, SWIFTADE1," _
          & " SWIFTAMON, SWIFTADE2, SWIFTADEN, SWIFTAHEN, SWIFTACOM, SWIFTATES, SWIFTASUP, SWIFTAVAL, SWIFTAAGE," _
          & " SWIFTASER, SWIFTASSE, SWIFTAUTI, SWIFTANUM, SWIFTAUT1, SWIFTAPVA, SWIFTAUT2 ) values ( " _
        & lYSWIHIA0.SWIHIAETA _
        & ",'" & lYSWIHIA0.SWIHIAREF & "','" & lYSWIHIA0.SWIHIANEN & "'" _
        & ",'" & lYSWIHIA0.SWIHIAPRI & "','" & lYSWIHIA0.SWIHIAMES & "'" _
        & ",'" & lYSWIHIA0.SWIHIADOR & "','" & lYSWIHIA0.SWIHIADES & "'" _
        & "," & lYSWIHIA0.SWIHIADVA & ",'" & lYSWIHIA0.SWIHIADE1 & "'" _
        & "," & lYSWIHIA0.SWIHIAMON & ",'" & lYSWIHIA0.SWIHIADE2 & "'" _
        & "," & lYSWIHIA0.SWIHIADEN _
        & "," & lYSWIHIA0.SWIHIAHEN _
        & ",'" & lYSWIHIA0.SWIHIACOM & "'" _
        & ",'" & lYSWIHIA0.SWIHIATES & "'" _
        & ",'" & lYSWIHIA0.SWIHIASUP & "'" _
        & ",'" & lYSWIHIA0.SWIHIAVAL & "'" _
        & "," & lYSWIHIA0.SWIHIAAGE _
        & ",'" & lYSWIHIA0.SWIHIASER & "'" _
        & ",'" & lYSWIHIA0.SWIHIASSE & "'" _
        & ",'" & lYSWIHIA0.SWIHIAUTI & "'" _
        & "," & lYSWIHIA0.SWIHIANUM _
        & ",'" & lYSWIHIA0.SWIHIAUT1 & "'" _
        & ",'" & lYSWIHIA0.SWIHIAPVA & "'" _
        & ",'" & lYSWIHIA0.SWIHIAUT2 & "')"

Exit Function
Error_Handler:
srvYSWIFTA0_Sql_Restauration = Error


End Function

Public Function srvYSWIFTB0_Sql_Restauration(lYSWIHIB0 As typeZSWIHIB0, lSql As String)
On Error GoTo Error_Handler
Dim K As Integer
' Supprimer le charactère ' dans la requête SQL
Do
    K = InStr(lYSWIHIB0.SWIHIBDET, "'")
    If K > 0 Then Mid$(lYSWIHIB0.SWIHIBDET, K, 1) = " "
Loop Until K = 0

srvYSWIFTB0_Sql_Restauration = Null
lSql = "Insert into " & paramIBM_Library_SAB & ".ZSWIFTB0 " _
        & "(SWIFTBETA,SWIFTBNUM,SWIFTBNEN,SWIFTBNLI,SWIFTBDET) values ( " _
        & lYSWIHIB0.SWIHIBETA _
        & "," & lYSWIHIB0.SWIHIBNUM _
        & "," & lYSWIHIB0.SWIHIBNEN _
        & "," & lYSWIHIB0.SWIHIBNLI _
        & ",'" & lYSWIHIB0.SWIHIBDET & "')"
        
Exit Function
Error_Handler:
srvYSWIFTB0_Sql_Restauration = Error
End Function



Public Function srvYSWIFTC0_Sql_Restauration(lYSWIHIC0 As typeZSWIHIC0, lSql As String)
On Error GoTo Error_Handler
srvYSWIFTC0_Sql_Restauration = Null
lSql = "Insert into " & paramIBM_Library_SAB & ".ZSWIFTC0 " _
          & "(SWIFTCETA, SWIFTCNUM , SWIFTCNEN, SWIFTCNLI, SWIFTCNSE, SWIFTCSOC, SWIFTCNCH, SWIFTCCOC, SWIFTCNLC," _
          & " SWIFTCSEQ, SWIFTCCHA, SWIFTCILI, SWIFTCFAC, SWIFTCSIG, SWIFTCSMA, SWIFTCCMA, SWIFTCSMI, SWIFTCCMI" _
          & " ) values ( " _
        & lYSWIHIC0.SWIHICETA _
        & "," & lYSWIHIC0.SWIHICNUM _
        & "," & lYSWIHIC0.SWIHICNEN _
        & "," & lYSWIHIC0.SWIHICNLI _
        & ",'" & lYSWIHIC0.SWIHICNSE & "'" _
        & "," & lYSWIHIC0.SWIHICSOC _
        & "," & lYSWIHIC0.SWIHICNCH _
        & "," & lYSWIHIC0.SWIHICCOC _
        & "," & lYSWIHIC0.SWIHICNLC _
        & ",'" & lYSWIHIC0.SWIHICSEQ & "'" _
        & ",'" & lYSWIHIC0.SWIHICCHA & "'" _
        & ",'" & lYSWIHIC0.SWIHICILI & "'" _
        & ",'" & lYSWIHIC0.SWIHICFAC & "'" _
        & ",'" & lYSWIHIC0.SWIHICSIG & "'" _
        & "," & lYSWIHIC0.SWIHICSMA _
        & "," & lYSWIHIC0.SWIHICCMA _
        & "," & lYSWIHIC0.SWIHICSMI _
        & "," & lYSWIHIC0.SWIHICCMI & ")"

Exit Function
Error_Handler:
srvYSWIFTC0_Sql_Restauration = Error
End Function

Public Function srvYSWITEM0_Sql_Restauration(lYSWIHIT0 As typeZSWIHIT0, lSql As String)
On Error GoTo Error_Handler
srvYSWITEM0_Sql_Restauration = Null
Dim K As Integer
' Supprimer le charactère ' dans la requête SQL
Do
    K = InStr(lYSWIHIT0.SWIHITVAL, "'")
    If K > 0 Then Mid$(lYSWIHIT0.SWIHITVAL, K, 1) = " "
Loop Until K = 0

lSql = "Insert into " & paramIBM_Library_SAB & ".ZSWITEM0 " _
          & "(SWITEMETA, SWITEMNUM , SWITEMNEN,  SWITEMNSE, SWITEMSEQ, SWITEMOSE, SWITEMCHA, SWITEMOCH," _
          & " SWITEMIND, SWITEMZON, SWITEMOZO, SWITEMSZO, SWITEMOSZ, SWITEMCON, SWITEMCOM, SWITEMVAL" _
          & " ) values ( " _
        & lYSWIHIT0.SWIHITETA _
        & "," & lYSWIHIT0.SWIHITNUM _
        & "," & lYSWIHIT0.SWIHITNEN _
        & ",'" & lYSWIHIT0.SWIHITNSE & "'" _
        & ",'" & lYSWIHIT0.SWIHITSEQ & "'" _
        & "," & lYSWIHIT0.SWIHITOSE _
        & "," & lYSWIHIT0.SWIHITCHA _
        & "," & lYSWIHIT0.SWIHITOCH _
        & ",'" & lYSWIHIT0.SWIHITIND & "'" _
        & "," & lYSWIHIT0.SWIHITZON _
        & "," & lYSWIHIT0.SWIHITOZO _
        & "," & lYSWIHIT0.SWIHITSZO _
        & "," & lYSWIHIT0.SWIHITOSZ _
        & "," & lYSWIHIT0.SWIHITCON _
        & ",'" & lYSWIHIT0.SWIHITCOM & "'" _
        & ",'" & lYSWIHIT0.SWIHITVAL & "')"

Exit Function
Error_Handler:
srvYSWITEM0_Sql_Restauration = Error
End Function




Public Function srvYSWIHIB0_Sql_Sauvegarde(lYSWIFTB0 As typeZSWIFTB0, lSql As String)
On Error GoTo Error_Handler
Dim K As Integer
' Supprimer le charactère ' dans la requête SQL

Do
    K = InStr(lYSWIFTB0.SWIFTBDET, "'")
    If K > 0 Then Mid$(lYSWIFTB0.SWIFTBDET, K, 1) = " "
Loop Until K = 0

srvYSWIHIB0_Sql_Sauvegarde = Null
lSql = "Insert into " & paramIBM_Library_SAB & ".ZSWIHIB0 " _
        & "(SWIHIBETA,SWIHIBNUM,SWIHIBNEN,SWIHIBNLI,SWIHIBDET) values ( " _
        & lYSWIFTB0.SWIFTBETA _
        & "," & lYSWIFTB0.SWIFTBNUM _
        & "," & lYSWIFTB0.SWIFTBNEN _
        & "," & lYSWIFTB0.SWIFTBNLI _
        & ",'" & lYSWIFTB0.SWIFTBDET & "')"
        
Exit Function
Error_Handler:
srvYSWIHIB0_Sql_Sauvegarde = Error
End Function

Public Sub srvYSWIFTA0_Init(lYSWIFTA0 As typeZSWIFTA0)
lYSWIFTA0.SWIFTAETA = 0
lYSWIFTA0.SWIFTAREF = ""
lYSWIFTA0.SWIFTANEN = ""
lYSWIFTA0.SWIFTAPRI = ""
lYSWIFTA0.SWIFTAMES = ""
lYSWIFTA0.SWIFTADOR = ""
lYSWIFTA0.SWIFTADES = ""
lYSWIFTA0.SWIFTADVA = 0
lYSWIFTA0.SWIFTADE1 = ""
lYSWIFTA0.SWIFTAMON = 0
lYSWIFTA0.SWIFTADE2 = ""
lYSWIFTA0.SWIFTADEN = 0
lYSWIFTA0.SWIFTAHEN = 0
lYSWIFTA0.SWIFTACOM = ""
lYSWIFTA0.SWIFTATES = ""
lYSWIFTA0.SWIFTASUP = ""
lYSWIFTA0.SWIFTAVAL = ""
lYSWIFTA0.SWIFTAAGE = 0
lYSWIFTA0.SWIFTASER = ""
lYSWIFTA0.SWIFTASSE = ""
lYSWIFTA0.SWIFTAUTI = ""
lYSWIFTA0.SWIFTANUM = 0
lYSWIFTA0.SWIFTAUT1 = ""
lYSWIFTA0.SWIFTAPVA = ""
lYSWIFTA0.SWIFTAUT2 = ""
End Sub

Public Function srvYSWIACR0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWIACR0 As typeZSWIACR0)
On Error GoTo Error_Handler
srvYSWIACR0_GetBuffer_ODBC = Null
lYSWIACR0.SWIACRBIC = rsado("SWIACRBIC")
lYSWIACR0.SWIACRSNP = rsado("SWIACRSNP")
lYSWIACR0.SWIACRTBF = rsado("SWIACRTBF")
lYSWIACR0.SWIACRSIE = rsado("SWIACRSIE")
lYSWIACR0.SWIACRBQE = rsado("SWIACRBQE")
lYSWIACR0.SWIACRGUI = rsado("SWIACRGUI")
lYSWIACR0.SWIACRLI1 = rsado("SWIACRLI1")
lYSWIACR0.SWIACRLI2 = rsado("SWIACRLI2")
lYSWIACR0.SWIACRLI3 = rsado("SWIACRLI3")
lYSWIACR0.SWIACRRSN = rsado("SWIACRRSN")
lYSWIACR0.SWIACRRTB = rsado("SWIACRRTB")
Exit Function
Error_Handler:
srvYSWIACR0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWIALI0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWIALI0 As typeZSWIALI0)
On Error GoTo Error_Handler
srvYSWIALI0_GetBuffer_ODBC = Null
lYSWIALI0.SWIALIETA = rsado("SWIALIETA")
lYSWIALI0.SWIALIAGE = rsado("SWIALIAGE")
lYSWIALI0.SWIALISER = rsado("SWIALISER")
lYSWIALI0.SWIALISSE = rsado("SWIALISSE")
lYSWIALI0.SWIALIMES = rsado("SWIALIMES")
lYSWIALI0.SWIALINUM = rsado("SWIALINUM")
lYSWIALI0.SWIALINEN = rsado("SWIALINEN")
lYSWIALI0.SWIALINLI = rsado("SWIALINLI")
lYSWIALI0.SWIALIDON = rsado("SWIALIDON")
lYSWIALI0.SWIALIOK = rsado("SWIALIOK")
Exit Function
Error_Handler:
srvYSWIALI0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWIBUF0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWIBUF0 As typeZSWIBUF0)
On Error GoTo Error_Handler
srvYSWIBUF0_GetBuffer_ODBC = Null
lYSWIBUF0.SWIBUFETA = rsado("SWIBUFETA")
lYSWIBUF0.SWIBUFAGE = rsado("SWIBUFAGE")
lYSWIBUF0.SWIBUFSER = rsado("SWIBUFSER")
lYSWIBUF0.SWIBUFSSE = rsado("SWIBUFSSE")
lYSWIBUF0.SWIBUFREF = rsado("SWIBUFREF")
lYSWIBUF0.SWIBUFNLI = rsado("SWIBUFNLI")
lYSWIBUF0.SWIBUFDON = rsado("SWIBUFDON")
Exit Function
Error_Handler:
srvYSWIBUF0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWICCI0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWICCI0 As typeZSWICCI0)
On Error GoTo Error_Handler
srvYSWICCI0_GetBuffer_ODBC = Null
lYSWICCI0.SWICCIETA = rsado("SWICCIETA")
lYSWICCI0.SWICCIAGE = rsado("SWICCIAGE")
lYSWICCI0.SWICCISER = rsado("SWICCISER")
lYSWICCI0.SWICCISSE = rsado("SWICCISSE")
lYSWICCI0.SWICCIMES = rsado("SWICCIMES")
lYSWICCI0.SWICCINUM = rsado("SWICCINUM")
lYSWICCI0.SWICCINEN = rsado("SWICCINEN")
lYSWICCI0.SWICCINLI = rsado("SWICCINLI")
lYSWICCI0.SWICCIDON = rsado("SWICCIDON")
lYSWICCI0.SWICCIOK = rsado("SWICCIOK")
Exit Function
Error_Handler:
srvYSWICCI0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWICLA0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWICLA0 As typeZSWICLA0)
On Error GoTo Error_Handler
srvYSWICLA0_GetBuffer_ODBC = Null
lYSWICLA0.SWICLAETA = rsado("SWICLAETA")
lYSWICLA0.SWICLAAGE = rsado("SWICLAAGE")
lYSWICLA0.SWICLASER = rsado("SWICLASER")
lYSWICLA0.SWICLASES = rsado("SWICLASES")
lYSWICLA0.SWICLAOPR = rsado("SWICLAOPR")
lYSWICLA0.SWICLANUM = rsado("SWICLANUM")
lYSWICLA0.SWICLACLA = rsado("SWICLACLA")
lYSWICLA0.SWICLAMES = rsado("SWICLAMES")
lYSWICLA0.SWICLACRI = rsado("SWICLACRI")
lYSWICLA0.SWICLAREF = rsado("SWICLAREF")
lYSWICLA0.SWICLAINT = rsado("SWICLAINT")
lYSWICLA0.SWICLANEN = rsado("SWICLANEN")
Exit Function
Error_Handler:
srvYSWICLA0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWICRI0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWICRI0 As typeZSWICRI0)
On Error GoTo Error_Handler
srvYSWICRI0_GetBuffer_ODBC = Null
lYSWICRI0.SWICRIETA = rsado("SWICRIETA")
lYSWICRI0.SWICRICRI = rsado("SWICRICRI")
lYSWICRI0.SWICRISEQ = rsado("SWICRISEQ")
lYSWICRI0.SWICRIPRI = rsado("SWICRIPRI")
lYSWICRI0.SWICRINPR = rsado("SWICRINPR")
lYSWICRI0.SWICRITYP = rsado("SWICRITYP")
lYSWICRI0.SWICRIDON = rsado("SWICRIDON")
Exit Function
Error_Handler:
srvYSWICRI0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWIECA0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWIECA0 As typeZSWIECA0)
On Error GoTo Error_Handler
srvYSWIECA0_GetBuffer_ODBC = Null
lYSWIECA0.SWIECAETA = rsado("SWIECAETA")
lYSWIECA0.SWIECAREF = rsado("SWIECAREF")
lYSWIECA0.SWIECAMES = rsado("SWIECAMES")
lYSWIECA0.SWIECAPRI = rsado("SWIECAPRI")
lYSWIECA0.SWIECAEME = rsado("SWIECAEME")
lYSWIECA0.SWIECADVA = rsado("SWIECADVA")
lYSWIECA0.SWIECADE1 = rsado("SWIECADE1")
lYSWIECA0.SWIECAMON = rsado("SWIECAMON")
lYSWIECA0.SWIECADRE = rsado("SWIECADRE")
lYSWIECA0.SWIECAHRE = rsado("SWIECAHRE")
lYSWIECA0.SWIECAINT = rsado("SWIECAINT")
lYSWIECA0.SWIECACET = rsado("SWIECACET")
lYSWIECA0.SWIECAAGE = rsado("SWIECAAGE")
lYSWIECA0.SWIECASER = rsado("SWIECASER")
lYSWIECA0.SWIECASSE = rsado("SWIECASSE")
lYSWIECA0.SWIECAUTI = rsado("SWIECAUTI")
Exit Function
Error_Handler:
srvYSWIECA0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWIECB0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWIECB0 As typeZSWIECB0)
On Error GoTo Error_Handler
srvYSWIECB0_GetBuffer_ODBC = Null
lYSWIECB0.SWIECBETA = rsado("SWIECBETA")
lYSWIECB0.SWIECBNUM = rsado("SWIECBNUM")
lYSWIECB0.SWIECBNOR = rsado("SWIECBNOR")
lYSWIECB0.SWIECBCHA = rsado("SWIECBCHA")
lYSWIECB0.SWIECBIND = rsado("SWIECBIND")
lYSWIECB0.SWIECBZON = rsado("SWIECBZON")
lYSWIECB0.SWIECBSZO = rsado("SWIECBSZO")
lYSWIECB0.SWIECBINR = rsado("SWIECBINR")
lYSWIECB0.SWIECBVAL = rsado("SWIECBVAL")
Exit Function
Error_Handler:
srvYSWIECB0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWIEHA0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWIEHA0 As typeZSWIEHA0)
On Error GoTo Error_Handler
srvYSWIEHA0_GetBuffer_ODBC = Null
lYSWIEHA0.SWIEHAETA = rsado("SWIEHAETA")
lYSWIEHA0.SWIEHANUM = rsado("SWIEHANUM")
lYSWIEHA0.SWIEHAREF = rsado("SWIEHAREF")
lYSWIEHA0.SWIEHAMES = rsado("SWIEHAMES")
lYSWIEHA0.SWIEHAEME = rsado("SWIEHAEME")
lYSWIEHA0.SWIEHADRE = rsado("SWIEHADRE")
lYSWIEHA0.SWIEHAHRE = rsado("SWIEHAHRE")
lYSWIEHA0.SWIEHAAGE = rsado("SWIEHAAGE")
lYSWIEHA0.SWIEHASER = rsado("SWIEHASER")
lYSWIEHA0.SWIEHASSE = rsado("SWIEHASSE")
lYSWIEHA0.SWIEHAUTI = rsado("SWIEHAUTI")
lYSWIEHA0.SWIEHADTR = rsado("SWIEHADTR")
lYSWIEHA0.SWIEHAAVI = rsado("SWIEHAAVI")
lYSWIEHA0.SWIEHADVA = rsado("SWIEHADVA")
lYSWIEHA0.SWIEHADEV = rsado("SWIEHADEV")
lYSWIEHA0.SWIEHAMON = rsado("SWIEHAMON")
Exit Function
Error_Handler:
srvYSWIEHA0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWIEHB0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWIEHB0 As typeZSWIEHB0)
On Error GoTo Error_Handler
srvYSWIEHB0_GetBuffer_ODBC = Null
lYSWIEHB0.SWIEHBETA = rsado("SWIEHBETA")
lYSWIEHB0.SWIEHBNUM = rsado("SWIEHBNUM")
lYSWIEHB0.SWIEHBNOR = rsado("SWIEHBNOR")
lYSWIEHB0.SWIEHBCHA = rsado("SWIEHBCHA")
lYSWIEHB0.SWIEHBIND = rsado("SWIEHBIND")
lYSWIEHB0.SWIEHBZON = rsado("SWIEHBZON")
lYSWIEHB0.SWIEHBSZO = rsado("SWIEHBSZO")
lYSWIEHB0.SWIEHBINR = rsado("SWIEHBINR")
lYSWIEHB0.SWIEHBVAL = rsado("SWIEHBVAL")
Exit Function
Error_Handler:
srvYSWIEHB0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWIENA0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWIENA0 As typeZSWIENA0)
On Error GoTo Error_Handler
srvYSWIENA0_GetBuffer_ODBC = Null
lYSWIENA0.SWIENAETA = rsado("SWIENAETA")
lYSWIENA0.SWIENAREF = rsado("SWIENAREF")
lYSWIENA0.SWIENAMES = rsado("SWIENAMES")
lYSWIENA0.SWIENAPRI = rsado("SWIENAPRI")
lYSWIENA0.SWIENAEME = rsado("SWIENAEME")
lYSWIENA0.SWIENADVA = rsado("SWIENADVA")
lYSWIENA0.SWIENADE1 = rsado("SWIENADE1")
lYSWIENA0.SWIENAMON = rsado("SWIENAMON")
lYSWIENA0.SWIENADRE = rsado("SWIENADRE")
lYSWIENA0.SWIENAHRE = rsado("SWIENAHRE")
lYSWIENA0.SWIENAINT = rsado("SWIENAINT")
lYSWIENA0.SWIENACET = rsado("SWIENACET")
lYSWIENA0.SWIENAAGE = rsado("SWIENAAGE")
lYSWIENA0.SWIENASER = rsado("SWIENASER")
lYSWIENA0.SWIENASSE = rsado("SWIENASSE")
lYSWIENA0.SWIENAUTI = rsado("SWIENAUTI")
Exit Function
Error_Handler:
srvYSWIENA0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWIENB0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWIENB0 As typeZSWIENB0)
On Error GoTo Error_Handler
srvYSWIENB0_GetBuffer_ODBC = Null
lYSWIENB0.SWIENBETA = rsado("SWIENBETA")
lYSWIENB0.SWIENBNUM = rsado("SWIENBNUM")
lYSWIENB0.SWIENBNOR = rsado("SWIENBNOR")
lYSWIENB0.SWIENBCHA = rsado("SWIENBCHA")
lYSWIENB0.SWIENBIND = rsado("SWIENBIND")
lYSWIENB0.SWIENBZON = rsado("SWIENBZON")
lYSWIENB0.SWIENBSZO = rsado("SWIENBSZO")
lYSWIENB0.SWIENBINR = rsado("SWIENBINR")
lYSWIENB0.SWIENBVAL = rsado("SWIENBVAL")
Exit Function
Error_Handler:
srvYSWIENB0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWIENI0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWIENI0 As typeZSWIENI0)
On Error GoTo Error_Handler
srvYSWIENI0_GetBuffer_ODBC = Null
lYSWIENI0.SWIENIETA = rsado("SWIENIETA")
lYSWIENI0.SWIENIAGE = rsado("SWIENIAGE")
lYSWIENI0.SWIENISER = rsado("SWIENISER")
lYSWIENI0.SWIENISSE = rsado("SWIENISSE")
lYSWIENI0.SWIENIMES = rsado("SWIENIMES")
lYSWIENI0.SWIENINUM = rsado("SWIENINUM")
lYSWIENI0.SWIENINEN = rsado("SWIENINEN")
lYSWIENI0.SWIENINLI = rsado("SWIENINLI")
lYSWIENI0.SWIENIDON = rsado("SWIENIDON")
lYSWIENI0.SWIENIOK = rsado("SWIENIOK")
Exit Function
Error_Handler:
srvYSWIENI0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWIEVC0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWIEVC0 As typeZSWIEVC0)
On Error GoTo Error_Handler
srvYSWIEVC0_GetBuffer_ODBC = Null
lYSWIEVC0.SWIEVCDON = rsado("SWIEVCDON")
Exit Function
Error_Handler:
srvYSWIEVC0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWIEVI0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWIEVI0 As typeZSWIEVI0)
On Error GoTo Error_Handler
srvYSWIEVI0_GetBuffer_ODBC = Null
lYSWIEVI0.SWIEVIETA = rsado("SWIEVIETA")
lYSWIEVI0.SWIEVIAGE = rsado("SWIEVIAGE")
lYSWIEVI0.SWIEVISER = rsado("SWIEVISER")
lYSWIEVI0.SWIEVISSE = rsado("SWIEVISSE")
lYSWIEVI0.SWIEVIMES = rsado("SWIEVIMES")
lYSWIEVI0.SWIEVINUM = rsado("SWIEVINUM")
lYSWIEVI0.SWIEVINEN = rsado("SWIEVINEN")
lYSWIEVI0.SWIEVINLI = rsado("SWIEVINLI")
lYSWIEVI0.SWIEVIDON = rsado("SWIEVIDON")
lYSWIEVI0.SWIEVIOK = rsado("SWIEVIOK")
Exit Function
Error_Handler:
srvYSWIEVI0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWIFTA0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWIFTA0 As typeZSWIFTA0)
On Error GoTo Error_Handler
srvYSWIFTA0_GetBuffer_ODBC = Null
lYSWIFTA0.SWIFTAETA = rsado("SWIFTAETA")
lYSWIFTA0.SWIFTAREF = rsado("SWIFTAREF")
lYSWIFTA0.SWIFTANEN = rsado("SWIFTANEN")
lYSWIFTA0.SWIFTAPRI = rsado("SWIFTAPRI")
lYSWIFTA0.SWIFTAMES = rsado("SWIFTAMES")
lYSWIFTA0.SWIFTADOR = rsado("SWIFTADOR")
lYSWIFTA0.SWIFTADES = rsado("SWIFTADES")
lYSWIFTA0.SWIFTADVA = rsado("SWIFTADVA")
lYSWIFTA0.SWIFTADE1 = rsado("SWIFTADE1")
lYSWIFTA0.SWIFTAMON = rsado("SWIFTAMON")
lYSWIFTA0.SWIFTADE2 = rsado("SWIFTADE2")
lYSWIFTA0.SWIFTADEN = rsado("SWIFTADEN")
lYSWIFTA0.SWIFTAHEN = rsado("SWIFTAHEN")
lYSWIFTA0.SWIFTACOM = rsado("SWIFTACOM")
lYSWIFTA0.SWIFTATES = rsado("SWIFTATES")
lYSWIFTA0.SWIFTASUP = rsado("SWIFTASUP")
lYSWIFTA0.SWIFTAVAL = rsado("SWIFTAVAL")
lYSWIFTA0.SWIFTAAGE = rsado("SWIFTAAGE")
lYSWIFTA0.SWIFTASER = rsado("SWIFTASER")
lYSWIFTA0.SWIFTASSE = rsado("SWIFTASSE")
lYSWIFTA0.SWIFTAUTI = rsado("SWIFTAUTI")
lYSWIFTA0.SWIFTANUM = rsado("SWIFTANUM")
lYSWIFTA0.SWIFTAUT1 = rsado("SWIFTAUT1")
lYSWIFTA0.SWIFTAPVA = rsado("SWIFTAPVA")
lYSWIFTA0.SWIFTAUT2 = rsado("SWIFTAUT2")
Exit Function
Error_Handler:
srvYSWIFTA0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWIFTB0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWIFTB0 As typeZSWIFTB0)
On Error GoTo Error_Handler
srvYSWIFTB0_GetBuffer_ODBC = Null
lYSWIFTB0.SWIFTBETA = rsado("SWIFTBETA")
lYSWIFTB0.SWIFTBNUM = rsado("SWIFTBNUM")
lYSWIFTB0.SWIFTBNEN = rsado("SWIFTBNEN")
lYSWIFTB0.SWIFTBNLI = rsado("SWIFTBNLI")
lYSWIFTB0.SWIFTBDET = rsado("SWIFTBDET")
Exit Function
Error_Handler:
srvYSWIFTB0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWIFTC0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWIFTC0 As typeZSWIFTC0)
On Error GoTo Error_Handler
srvYSWIFTC0_GetBuffer_ODBC = Null
lYSWIFTC0.SWIFTCETA = rsado("SWIFTCETA")
lYSWIFTC0.SWIFTCNUM = rsado("SWIFTCNUM")
lYSWIFTC0.SWIFTCNEN = rsado("SWIFTCNEN")
lYSWIFTC0.SWIFTCNLI = rsado("SWIFTCNLI")
lYSWIFTC0.SWIFTCNSE = rsado("SWIFTCNSE")
lYSWIFTC0.SWIFTCSOC = rsado("SWIFTCSOC")
lYSWIFTC0.SWIFTCNCH = rsado("SWIFTCNCH")
lYSWIFTC0.SWIFTCCOC = rsado("SWIFTCCOC")
lYSWIFTC0.SWIFTCNLC = rsado("SWIFTCNLC")
lYSWIFTC0.SWIFTCSEQ = rsado("SWIFTCSEQ")
lYSWIFTC0.SWIFTCCHA = rsado("SWIFTCCHA")
lYSWIFTC0.SWIFTCILI = rsado("SWIFTCILI")
lYSWIFTC0.SWIFTCFAC = rsado("SWIFTCFAC")
lYSWIFTC0.SWIFTCSIG = rsado("SWIFTCSIG")
lYSWIFTC0.SWIFTCSMA = rsado("SWIFTCSMA")
lYSWIFTC0.SWIFTCCMA = rsado("SWIFTCCMA")
lYSWIFTC0.SWIFTCSMI = rsado("SWIFTCSMI")
lYSWIFTC0.SWIFTCCMI = rsado("SWIFTCCMI")
Exit Function
Error_Handler:
srvYSWIFTC0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWIGRN0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWIGRN0 As typeZSWIGRN0)
On Error GoTo Error_Handler
srvYSWIGRN0_GetBuffer_ODBC = Null
lYSWIGRN0.SWIGRNETA = rsado("SWIGRNETA")
lYSWIGRN0.SWIGRNGRP = rsado("SWIGRNGRP")
lYSWIGRN0.SWIGRNORD = rsado("SWIGRNORD")
lYSWIGRN0.SWIGRNNAT = rsado("SWIGRNNAT")
Exit Function
Error_Handler:
srvYSWIGRN0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWIHIA0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWIHIA0 As typeZSWIHIA0)
On Error GoTo Error_Handler
srvYSWIHIA0_GetBuffer_ODBC = Null
lYSWIHIA0.SWIHIAETA = rsado("SWIHIAETA")
lYSWIHIA0.SWIHIAREF = rsado("SWIHIAREF")
lYSWIHIA0.SWIHIANEN = rsado("SWIHIANEN")
lYSWIHIA0.SWIHIAPRI = rsado("SWIHIAPRI")
lYSWIHIA0.SWIHIAMES = rsado("SWIHIAMES")
lYSWIHIA0.SWIHIADOR = rsado("SWIHIADOR")
lYSWIHIA0.SWIHIADES = rsado("SWIHIADES")
lYSWIHIA0.SWIHIADVA = rsado("SWIHIADVA")
lYSWIHIA0.SWIHIADE1 = rsado("SWIHIADE1")
lYSWIHIA0.SWIHIAMON = rsado("SWIHIAMON")
lYSWIHIA0.SWIHIADE2 = rsado("SWIHIADE2")
lYSWIHIA0.SWIHIADEN = rsado("SWIHIADEN")
lYSWIHIA0.SWIHIAHEN = rsado("SWIHIAHEN")
lYSWIHIA0.SWIHIACOM = rsado("SWIHIACOM")
lYSWIHIA0.SWIHIATES = rsado("SWIHIATES")
lYSWIHIA0.SWIHIASUP = rsado("SWIHIASUP")
lYSWIHIA0.SWIHIAVAL = rsado("SWIHIAVAL")
lYSWIHIA0.SWIHIAAGE = rsado("SWIHIAAGE")
lYSWIHIA0.SWIHIASER = rsado("SWIHIASER")
lYSWIHIA0.SWIHIASSE = rsado("SWIHIASSE")
lYSWIHIA0.SWIHIAUTI = rsado("SWIHIAUTI")
lYSWIHIA0.SWIHIANUM = rsado("SWIHIANUM")
lYSWIHIA0.SWIHIAUT1 = rsado("SWIHIAUT1")
lYSWIHIA0.SWIHIAPVA = rsado("SWIHIAPVA")
lYSWIHIA0.SWIHIAUT2 = rsado("SWIHIAUT2")
Exit Function
Error_Handler:
srvYSWIHIA0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWIHIB0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWIHIB0 As typeZSWIHIB0)
On Error GoTo Error_Handler
srvYSWIHIB0_GetBuffer_ODBC = Null
lYSWIHIB0.SWIHIBETA = rsado("SWIHIBETA")
lYSWIHIB0.SWIHIBNUM = rsado("SWIHIBNUM")
lYSWIHIB0.SWIHIBNEN = rsado("SWIHIBNEN")
lYSWIHIB0.SWIHIBNLI = rsado("SWIHIBNLI")
lYSWIHIB0.SWIHIBDET = rsado("SWIHIBDET")
Exit Function
Error_Handler:
srvYSWIHIB0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWIFTB0_YSWIHIB0(lYSWIFTB0 As typeZSWIFTB0, lYSWIHIB0 As typeZSWIHIB0)
On Error GoTo Error_Handler
srvYSWIFTB0_YSWIHIB0 = Null
lYSWIHIB0.SWIHIBETA = lYSWIFTB0.SWIFTBETA
lYSWIHIB0.SWIHIBNUM = lYSWIFTB0.SWIFTBNUM
lYSWIHIB0.SWIHIBNEN = lYSWIFTB0.SWIFTBNEN
lYSWIHIB0.SWIHIBNLI = lYSWIFTB0.SWIFTBNLI
lYSWIHIB0.SWIHIBDET = lYSWIFTB0.SWIFTBDET
Exit Function
Error_Handler:
srvYSWIFTB0_YSWIHIB0 = Error
End Function

Public Function srvYSWIHIC0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWIHIC0 As typeZSWIHIC0)
On Error GoTo Error_Handler
srvYSWIHIC0_GetBuffer_ODBC = Null
lYSWIHIC0.SWIHICETA = rsado("SWIHICETA")
lYSWIHIC0.SWIHICNUM = rsado("SWIHICNUM")
lYSWIHIC0.SWIHICNEN = rsado("SWIHICNEN")
lYSWIHIC0.SWIHICNLI = rsado("SWIHICNLI")
lYSWIHIC0.SWIHICNSE = rsado("SWIHICNSE")
lYSWIHIC0.SWIHICSOC = rsado("SWIHICSOC")
lYSWIHIC0.SWIHICNCH = rsado("SWIHICNCH")
lYSWIHIC0.SWIHICCOC = rsado("SWIHICCOC")
lYSWIHIC0.SWIHICNLC = rsado("SWIHICNLC")
lYSWIHIC0.SWIHICSEQ = rsado("SWIHICSEQ")
lYSWIHIC0.SWIHICCHA = rsado("SWIHICCHA")
lYSWIHIC0.SWIHICILI = rsado("SWIHICILI")
lYSWIHIC0.SWIHICFAC = rsado("SWIHICFAC")
lYSWIHIC0.SWIHICSIG = rsado("SWIHICSIG")
lYSWIHIC0.SWIHICSMA = rsado("SWIHICSMA")
lYSWIHIC0.SWIHICCMA = rsado("SWIHICCMA")
lYSWIHIC0.SWIHICSMI = rsado("SWIHICSMI")
lYSWIHIC0.SWIHICCMI = rsado("SWIHICCMI")
Exit Function
Error_Handler:
srvYSWIHIC0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWIHIT0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWIHIT0 As typeZSWIHIT0)
On Error GoTo Error_Handler
srvYSWIHIT0_GetBuffer_ODBC = Null
lYSWIHIT0.SWIHITETA = rsado("SWIHITETA")
lYSWIHIT0.SWIHITNUM = rsado("SWIHITNUM")
lYSWIHIT0.SWIHITNEN = rsado("SWIHITNEN")
lYSWIHIT0.SWIHITNSE = rsado("SWIHITNSE")
lYSWIHIT0.SWIHITSEQ = rsado("SWIHITSEQ")
lYSWIHIT0.SWIHITOSE = rsado("SWIHITOSE")
lYSWIHIT0.SWIHITCHA = rsado("SWIHITCHA")
lYSWIHIT0.SWIHITOCH = rsado("SWIHITOCH")
lYSWIHIT0.SWIHITIND = rsado("SWIHITIND")
lYSWIHIT0.SWIHITZON = rsado("SWIHITZON")
lYSWIHIT0.SWIHITOZO = rsado("SWIHITOZO")
lYSWIHIT0.SWIHITSZO = rsado("SWIHITSZO")
lYSWIHIT0.SWIHITOSZ = rsado("SWIHITOSZ")
lYSWIHIT0.SWIHITCON = rsado("SWIHITCON")
lYSWIHIT0.SWIHITCOM = rsado("SWIHITCOM")
lYSWIHIT0.SWIHITVAL = rsado("SWIHITVAL")
Exit Function
Error_Handler:
srvYSWIHIT0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWIJOB0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWIJOB0 As typeZSWIJOB0)
On Error GoTo Error_Handler
srvYSWIJOB0_GetBuffer_ODBC = Null
lYSWIJOB0.SWIJOBETA = rsado("SWIJOBETA")
lYSWIJOB0.SWIJOBPRO = rsado("SWIJOBPRO")
lYSWIJOB0.SWIJOBJOB = rsado("SWIJOBJOB")
lYSWIJOB0.SWIJOBUSR = rsado("SWIJOBUSR")
lYSWIJOB0.SWIJOBNBR = rsado("SWIJOBNBR")
lYSWIJOB0.SWIJOBDLA = rsado("SWIJOBDLA")
lYSWIJOB0.SWIJOBHLA = rsado("SWIJOBHLA")
lYSWIJOB0.SWIJOBENV = rsado("SWIJOBENV")
lYSWIJOB0.SWIJOBTER = rsado("SWIJOBTER")
lYSWIJOB0.SWIJOBACT = rsado("SWIJOBACT")
Exit Function
Error_Handler:
srvYSWIJOB0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWIMEA0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWIMEA0 As typeZSWIMEA0)
On Error GoTo Error_Handler
srvYSWIMEA0_GetBuffer_ODBC = Null
lYSWIMEA0.SWIMEAETA = rsado("SWIMEAETA")
lYSWIMEA0.SWIMEANUM = rsado("SWIMEANUM")
lYSWIMEA0.SWIMEAREF = rsado("SWIMEAREF")
lYSWIMEA0.SWIMEAMES = rsado("SWIMEAMES")
lYSWIMEA0.SWIMEAEME = rsado("SWIMEAEME")
lYSWIMEA0.SWIMEADRE = rsado("SWIMEADRE")
lYSWIMEA0.SWIMEAHRE = rsado("SWIMEAHRE")
lYSWIMEA0.SWIMEAAGE = rsado("SWIMEAAGE")
lYSWIMEA0.SWIMEASER = rsado("SWIMEASER")
lYSWIMEA0.SWIMEASSE = rsado("SWIMEASSE")
lYSWIMEA0.SWIMEAUTI = rsado("SWIMEAUTI")
lYSWIMEA0.SWIMEADTR = rsado("SWIMEADTR")
lYSWIMEA0.SWIMEAAVI = rsado("SWIMEAAVI")
lYSWIMEA0.SWIMEADVA = rsado("SWIMEADVA")
lYSWIMEA0.SWIMEADEV = rsado("SWIMEADEV")
lYSWIMEA0.SWIMEAMON = rsado("SWIMEAMON")
Exit Function
Error_Handler:
srvYSWIMEA0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWIMEB0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWIMEB0 As typeZSWIMEB0)
On Error GoTo Error_Handler
srvYSWIMEB0_GetBuffer_ODBC = Null
lYSWIMEB0.SWIMEBETA = rsado("SWIMEBETA")
lYSWIMEB0.SWIMEBNUM = rsado("SWIMEBNUM")
lYSWIMEB0.SWIMEBNOR = rsado("SWIMEBNOR")
lYSWIMEB0.SWIMEBCHA = rsado("SWIMEBCHA")
lYSWIMEB0.SWIMEBIND = rsado("SWIMEBIND")
lYSWIMEB0.SWIMEBZON = rsado("SWIMEBZON")
lYSWIMEB0.SWIMEBSZO = rsado("SWIMEBSZO")
lYSWIMEB0.SWIENBINR = rsado("SWIENBINR")
lYSWIMEB0.SWIMEBVAL = rsado("SWIMEBVAL")
Exit Function
Error_Handler:
srvYSWIMEB0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWIMEC0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWIMEC0 As typeZSWIMEC0)
On Error GoTo Error_Handler
srvYSWIMEC0_GetBuffer_ODBC = Null
lYSWIMEC0.SWIMECETA = rsado("SWIMECETA")
lYSWIMEC0.SWIMECNUM = rsado("SWIMECNUM")
lYSWIMEC0.SWIMECNOR = rsado("SWIMECNOR")
lYSWIMEC0.SWIMECCOM = rsado("SWIMECCOM")
Exit Function
Error_Handler:
srvYSWIMEC0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWIMEM0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWIMEM0 As typeZSWIMEM0)
On Error GoTo Error_Handler
srvYSWIMEM0_GetBuffer_ODBC = Null
lYSWIMEM0.SWIMEMETA = rsado("SWIMEMETA")
lYSWIMEM0.SWIMEMNUM = rsado("SWIMEMNUM")
lYSWIMEM0.SWIMEMNOR = rsado("SWIMEMNOR")
lYSWIMEM0.SWIMEMMEG = rsado("SWIMEMMEG")
lYSWIMEM0.SWIMEMREG = rsado("SWIMEMREG")
Exit Function
Error_Handler:
srvYSWIMEM0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWIMEO0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWIMEO0 As typeZSWIMEO0)
On Error GoTo Error_Handler
srvYSWIMEO0_GetBuffer_ODBC = Null
lYSWIMEO0.SWIMEOETA = rsado("SWIMEOETA")
lYSWIMEO0.SWIMEONUM = rsado("SWIMEONUM")
lYSWIMEO0.SWIMEONOR = rsado("SWIMEONOR")
lYSWIMEO0.SWIMEOOPR = rsado("SWIMEOOPR")
lYSWIMEO0.SWIMEONAT = rsado("SWIMEONAT")
lYSWIMEO0.SWIMEONOP = rsado("SWIMEONOP")
Exit Function
Error_Handler:
srvYSWIMEO0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWIRAL0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWIRAL0 As typeZSWIRAL0)
On Error GoTo Error_Handler
srvYSWIRAL0_GetBuffer_ODBC = Null
lYSWIRAL0.SWIRALDON = rsado("SWIRALDON")
lYSWIRAL0.SWIRALETA = rsado("SWIRALETA")
lYSWIRAL0.SWIRALMES = rsado("SWIRALMES")
Exit Function
Error_Handler:
srvYSWIRAL0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWIRDE0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWIRDE0 As typeZSWIRDE0)
On Error GoTo Error_Handler
srvYSWIRDE0_GetBuffer_ODBC = Null
lYSWIRDE0.SWIRDEETA = rsado("SWIRDEETA")
lYSWIRDE0.SWIRDEBIC = rsado("SWIRDEBIC")
lYSWIRDE0.SWIRDENUM = rsado("SWIRDENUM")
lYSWIRDE0.SWIRDECOM = rsado("SWIRDECOM")
lYSWIRDE0.SWIRDEDAT = rsado("SWIRDEDAT")
lYSWIRDE0.SWIRDECOU = rsado("SWIRDECOU")
Exit Function
Error_Handler:
srvYSWIRDE0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWIREC0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWIREC0 As typeZSWIREC0)
On Error GoTo Error_Handler
srvYSWIREC0_GetBuffer_ODBC = Null
lYSWIREC0.SWIRECETA = rsado("SWIRECETA")
lYSWIREC0.SWIRECNUM = rsado("SWIRECNUM")
lYSWIREC0.SWIRECNLI = rsado("SWIRECNLI")
lYSWIREC0.SWIRECMES = rsado("SWIRECMES")
lYSWIREC0.SWIRECDON = rsado("SWIRECDON")
Exit Function
Error_Handler:
srvYSWIREC0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWIRED0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWIRED0 As typeZSWIRED0)
On Error GoTo Error_Handler
srvYSWIRED0_GetBuffer_ODBC = Null
lYSWIRED0.SWIREDETA = rsado("SWIREDETA")
lYSWIRED0.SWIREDAGE = rsado("SWIREDAGE")
lYSWIRED0.SWIREDSER = rsado("SWIREDSER")
lYSWIRED0.SWIREDSSE = rsado("SWIREDSSE")
lYSWIRED0.SWIREDREF = rsado("SWIREDREF")
lYSWIRED0.SWIREDME1 = rsado("SWIREDME1")
lYSWIRED0.SWIREDME2 = rsado("SWIREDME2")
lYSWIRED0.SWIREDEM1 = rsado("SWIREDEM1")
lYSWIRED0.SWIREDEM2 = rsado("SWIREDEM2")
lYSWIRED0.SWIREDNU1 = rsado("SWIREDNU1")
lYSWIRED0.SWIREDNU2 = rsado("SWIREDNU2")
lYSWIRED0.SWIREDDAT = rsado("SWIREDDAT")
lYSWIRED0.SWIREDAVI = rsado("SWIREDAVI")
Exit Function
Error_Handler:
srvYSWIRED0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWIRST0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWIRST0 As typeZSWIRST0)
On Error GoTo Error_Handler
srvYSWIRST0_GetBuffer_ODBC = Null
lYSWIRST0.SWIRSTDON = rsado("SWIRSTDON")
lYSWIRST0.SWIRSTETA = rsado("SWIRSTETA")
lYSWIRST0.SWIRSTMES = rsado("SWIRSTMES")
Exit Function
Error_Handler:
srvYSWIRST0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWISCA0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWISCA0 As typeZSWISCA0)
On Error GoTo Error_Handler
srvYSWISCA0_GetBuffer_ODBC = Null
lYSWISCA0.SWISCAETA = rsado("SWISCAETA")
lYSWISCA0.SWISCAREF = rsado("SWISCAREF")
lYSWISCA0.SWISCANEN = rsado("SWISCANEN")
lYSWISCA0.SWISCAPRI = rsado("SWISCAPRI")
lYSWISCA0.SWISCAMES = rsado("SWISCAMES")
lYSWISCA0.SWISCADOR = rsado("SWISCADOR")
lYSWISCA0.SWISCADES = rsado("SWISCADES")
lYSWISCA0.SWISCADVA = rsado("SWISCADVA")
lYSWISCA0.SWISCADE1 = rsado("SWISCADE1")
lYSWISCA0.SWISCAMON = rsado("SWISCAMON")
lYSWISCA0.SWISCADE2 = rsado("SWISCADE2")
lYSWISCA0.SWISCADEN = rsado("SWISCADEN")
lYSWISCA0.SWISCAHEN = rsado("SWISCAHEN")
lYSWISCA0.SWISCACOM = rsado("SWISCACOM")
lYSWISCA0.SWISCATES = rsado("SWISCATES")
lYSWISCA0.SWISCASUP = rsado("SWISCASUP")
lYSWISCA0.SWISCAVAL = rsado("SWISCAVAL")
lYSWISCA0.SWISCAAGE = rsado("SWISCAAGE")
lYSWISCA0.SWISCASER = rsado("SWISCASER")
lYSWISCA0.SWISCASSE = rsado("SWISCASSE")
lYSWISCA0.SWISCAUTI = rsado("SWISCAUTI")
lYSWISCA0.SWISCANUM = rsado("SWISCANUM")
lYSWISCA0.SWISCAUT1 = rsado("SWISCAUT1")
lYSWISCA0.SWISCAPVA = rsado("SWISCAPVA")
lYSWISCA0.SWISCAUT2 = rsado("SWISCAUT2")
Exit Function
Error_Handler:
srvYSWISCA0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWISCB0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWISCB0 As typeZSWISCB0)
On Error GoTo Error_Handler
srvYSWISCB0_GetBuffer_ODBC = Null
lYSWISCB0.SWISCBETA = rsado("SWISCBETA")
lYSWISCB0.SWISCBNUM = rsado("SWISCBNUM")
lYSWISCB0.SWISCBNEN = rsado("SWISCBNEN")
lYSWISCB0.SWISCBNLI = rsado("SWISCBNLI")
lYSWISCB0.SWISCBDET = rsado("SWISCBDET")
Exit Function
Error_Handler:
srvYSWISCB0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWISCC0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWISCC0 As typeZSWISCC0)
On Error GoTo Error_Handler
srvYSWISCC0_GetBuffer_ODBC = Null
lYSWISCC0.SWISCCETA = rsado("SWISCCETA")
lYSWISCC0.SWISCCNUM = rsado("SWISCCNUM")
lYSWISCC0.SWISCCNEN = rsado("SWISCCNEN")
lYSWISCC0.SWISCCNLI = rsado("SWISCCNLI")
lYSWISCC0.SWISCCNSE = rsado("SWISCCNSE")
lYSWISCC0.SWISCCSOC = rsado("SWISCCSOC")
lYSWISCC0.SWISCCNCH = rsado("SWISCCNCH")
lYSWISCC0.SWISCCCOC = rsado("SWISCCCOC")
lYSWISCC0.SWISCCNLC = rsado("SWISCCNLC")
lYSWISCC0.SWISCCSEQ = rsado("SWISCCSEQ")
lYSWISCC0.SWISCCCHA = rsado("SWISCCCHA")
lYSWISCC0.SWISCCILI = rsado("SWISCCILI")
lYSWISCC0.SWISCCFAC = rsado("SWISCCFAC")
lYSWISCC0.SWISCCSIG = rsado("SWISCCSIG")
lYSWISCC0.SWISCCSMA = rsado("SWISCCSMA")
lYSWISCC0.SWISCCCMA = rsado("SWISCCCMA")
lYSWISCC0.SWISCCSMI = rsado("SWISCCSMI")
lYSWISCC0.SWISCCCMI = rsado("SWISCCCMI")
Exit Function
Error_Handler:
srvYSWISCC0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWISCT0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWISCT0 As typeZSWISCT0)
On Error GoTo Error_Handler
srvYSWISCT0_GetBuffer_ODBC = Null
lYSWISCT0.SWISCTETA = rsado("SWISCTETA")
lYSWISCT0.SWISCTNUM = rsado("SWISCTNUM")
lYSWISCT0.SWISCTNEN = rsado("SWISCTNEN")
lYSWISCT0.SWISCTNSE = rsado("SWISCTNSE")
lYSWISCT0.SWISCTSEQ = rsado("SWISCTSEQ")
lYSWISCT0.SWISCTOSE = rsado("SWISCTOSE")
lYSWISCT0.SWISCTCHA = rsado("SWISCTCHA")
lYSWISCT0.SWISCTOCH = rsado("SWISCTOCH")
lYSWISCT0.SWISCTIND = rsado("SWISCTIND")
lYSWISCT0.SWISCTZON = rsado("SWISCTZON")
lYSWISCT0.SWISCTOZO = rsado("SWISCTOZO")
lYSWISCT0.SWISCTSZO = rsado("SWISCTSZO")
lYSWISCT0.SWISCTOSZ = rsado("SWISCTOSZ")
lYSWISCT0.SWISCTCON = rsado("SWISCTCON")
lYSWISCT0.SWISCTCOM = rsado("SWISCTCOM")
lYSWISCT0.SWISCTVAL = rsado("SWISCTVAL")
Exit Function
Error_Handler:
srvYSWISCT0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWISHA0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWISHA0 As typeZSWISHA0)
On Error GoTo Error_Handler
srvYSWISHA0_GetBuffer_ODBC = Null
lYSWISHA0.SWISHAETA = rsado("SWISHAETA")
lYSWISHA0.SWISHAREF = rsado("SWISHAREF")
lYSWISHA0.SWISHANEN = rsado("SWISHANEN")
lYSWISHA0.SWISHAPRI = rsado("SWISHAPRI")
lYSWISHA0.SWISHAMES = rsado("SWISHAMES")
lYSWISHA0.SWISHADOR = rsado("SWISHADOR")
lYSWISHA0.SWISHADES = rsado("SWISHADES")
lYSWISHA0.SWISHADVA = rsado("SWISHADVA")
lYSWISHA0.SWISHADE1 = rsado("SWISHADE1")
lYSWISHA0.SWISHAMON = rsado("SWISHAMON")
lYSWISHA0.SWISHADE2 = rsado("SWISHADE2")
lYSWISHA0.SWISHADEN = rsado("SWISHADEN")
lYSWISHA0.SWISHAHEN = rsado("SWISHAHEN")
lYSWISHA0.SWISHACOM = rsado("SWISHACOM")
lYSWISHA0.SWISHATES = rsado("SWISHATES")
lYSWISHA0.SWISHASUP = rsado("SWISHASUP")
lYSWISHA0.SWISHAVAL = rsado("SWISHAVAL")
lYSWISHA0.SWISHAAGE = rsado("SWISHAAGE")
lYSWISHA0.SWISHASER = rsado("SWISHASER")
lYSWISHA0.SWISHASSE = rsado("SWISHASSE")
lYSWISHA0.SWISHAUTI = rsado("SWISHAUTI")
lYSWISHA0.SWISHANUM = rsado("SWISHANUM")
lYSWISHA0.SWISHAUT1 = rsado("SWISHAUT1")
lYSWISHA0.SWISHAPVA = rsado("SWISHAPVA")
lYSWISHA0.SWISHAUT2 = rsado("SWISHAUT2")
Exit Function
Error_Handler:
srvYSWISHA0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWISHB0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWISHB0 As typeZSWISHB0)
On Error GoTo Error_Handler
srvYSWISHB0_GetBuffer_ODBC = Null
lYSWISHB0.SWISHBETA = rsado("SWISHBETA")
lYSWISHB0.SWISHBNUM = rsado("SWISHBNUM")
lYSWISHB0.SWISHBNEN = rsado("SWISHBNEN")
lYSWISHB0.SWISHBNLI = rsado("SWISHBNLI")
lYSWISHB0.SWISHBDET = rsado("SWISHBDET")
Exit Function
Error_Handler:
srvYSWISHB0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWISHC0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWISHC0 As typeZSWISHC0)
On Error GoTo Error_Handler
srvYSWISHC0_GetBuffer_ODBC = Null
lYSWISHC0.SWISHCETA = rsado("SWISHCETA")
lYSWISHC0.SWISHCNUM = rsado("SWISHCNUM")
lYSWISHC0.SWISHCNEN = rsado("SWISHCNEN")
lYSWISHC0.SWISHCNLI = rsado("SWISHCNLI")
lYSWISHC0.SWISHCNSE = rsado("SWISHCNSE")
lYSWISHC0.SWISHCSOC = rsado("SWISHCSOC")
lYSWISHC0.SWISHCNCH = rsado("SWISHCNCH")
lYSWISHC0.SWISHCCOC = rsado("SWISHCCOC")
lYSWISHC0.SWISHCNLC = rsado("SWISHCNLC")
lYSWISHC0.SWISHCSEQ = rsado("SWISHCSEQ")
lYSWISHC0.SWISHCCHA = rsado("SWISHCCHA")
lYSWISHC0.SWISHCILI = rsado("SWISHCILI")
lYSWISHC0.SWISHCFAC = rsado("SWISHCFAC")
lYSWISHC0.SWISHCSIG = rsado("SWISHCSIG")
lYSWISHC0.SWISHCSMA = rsado("SWISHCSMA")
lYSWISHC0.SWISHCCMA = rsado("SWISHCCMA")
lYSWISHC0.SWISHCSMI = rsado("SWISHCSMI")
lYSWISHC0.SWISHCCMI = rsado("SWISHCCMI")
Exit Function
Error_Handler:
srvYSWISHC0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWISHT0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWISHT0 As typeZSWISHT0)
On Error GoTo Error_Handler
srvYSWISHT0_GetBuffer_ODBC = Null
lYSWISHT0.SWISHTETA = rsado("SWISHTETA")
lYSWISHT0.SWISHTNUM = rsado("SWISHTNUM")
lYSWISHT0.SWISHTNEN = rsado("SWISHTNEN")
lYSWISHT0.SWISHTNSE = rsado("SWISHTNSE")
lYSWISHT0.SWISHTSEQ = rsado("SWISHTSEQ")
lYSWISHT0.SWISHTOSE = rsado("SWISHTOSE")
lYSWISHT0.SWISHTCHA = rsado("SWISHTCHA")
lYSWISHT0.SWISHTOCH = rsado("SWISHTOCH")
lYSWISHT0.SWISHTIND = rsado("SWISHTIND")
lYSWISHT0.SWISHTZON = rsado("SWISHTZON")
lYSWISHT0.SWISHTOZO = rsado("SWISHTOZO")
lYSWISHT0.SWISHTSZO = rsado("SWISHTSZO")
lYSWISHT0.SWISHTOSZ = rsado("SWISHTOSZ")
lYSWISHT0.SWISHTCON = rsado("SWISHTCON")
lYSWISHT0.SWISHTCOM = rsado("SWISHTCOM")
lYSWISHT0.SWISHTVAL = rsado("SWISHTVAL")
Exit Function
Error_Handler:
srvYSWISHT0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWISRC0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWISRC0 As typeZSWISRC0)
On Error GoTo Error_Handler
srvYSWISRC0_GetBuffer_ODBC = Null
lYSWISRC0.AIDCLICOD = rsado("AIDCLICOD")
lYSWISRC0.AIDCLIPRG = rsado("AIDCLIPRG")
lYSWISRC0.AIDCLIFMT = rsado("AIDCLIFMT")
Exit Function
Error_Handler:
srvYSWISRC0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWITAB0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWITAB0 As typeZSWITAB0)
On Error GoTo Error_Handler
srvYSWITAB0_GetBuffer_ODBC = Null
lYSWITAB0.SWITABETA = rsado("SWITABETA")
lYSWITAB0.SWITABNUM = rsado("SWITABNUM")
lYSWITAB0.SWITABARG = rsado("SWITABARG")
lYSWITAB0.SWITABLO1 = rsado("SWITABLO1")
lYSWITAB0.SWITABLO2 = rsado("SWITABLO2")
lYSWITAB0.SWITABDON = rsado("SWITABDON")
Exit Function
Error_Handler:
srvYSWITAB0_GetBuffer_ODBC = Error
End Function
Public Function srvYSWITEM0_GetBuffer_ODBC(rsado As ADODB.Recordset, lYSWITEM0 As typeZSWITEM0)
On Error GoTo Error_Handler
srvYSWITEM0_GetBuffer_ODBC = Null
lYSWITEM0.SWITEMETA = rsado("SWITEMETA")
lYSWITEM0.SWITEMNUM = rsado("SWITEMNUM")
lYSWITEM0.SWITEMNEN = rsado("SWITEMNEN")
lYSWITEM0.SWITEMNSE = rsado("SWITEMNSE")
lYSWITEM0.SWITEMSEQ = rsado("SWITEMSEQ")
lYSWITEM0.SWITEMOSE = rsado("SWITEMOSE")
lYSWITEM0.SWITEMCHA = rsado("SWITEMCHA")
lYSWITEM0.SWITEMOCH = rsado("SWITEMOCH")
lYSWITEM0.SWITEMIND = rsado("SWITEMIND")
lYSWITEM0.SWITEMZON = rsado("SWITEMZON")
lYSWITEM0.SWITEMOZO = rsado("SWITEMOZO")
lYSWITEM0.SWITEMSZO = rsado("SWITEMSZO")
lYSWITEM0.SWITEMOSZ = rsado("SWITEMOSZ")
lYSWITEM0.SWITEMCON = rsado("SWITEMCON")
lYSWITEM0.SWITEMCOM = rsado("SWITEMCOM")
lYSWITEM0.SWITEMVAL = rsado("SWITEMVAL")
Exit Function
Error_Handler:
srvYSWITEM0_GetBuffer_ODBC = Error
End Function
Public Sub srvYSWIACR0_fgDisplay(lYSWIACR0 As typeZSWIACR0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 12
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIACRBIC   11A"
fgDisplay.Col = 1: fgDisplay = "CODE BIC"
fgDisplay.Col = 2: fgDisplay = lYSWIACR0.SWIACRBIC
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWIACRSNP   11A"
fgDisplay.Col = 1: fgDisplay = "CODE BIC SNP"
fgDisplay.Col = 2: fgDisplay = lYSWIACR0.SWIACRSNP
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWIACRTBF   11A"
fgDisplay.Col = 1: fgDisplay = "CODE BIC TBF"
fgDisplay.Col = 2: fgDisplay = lYSWIACR0.SWIACRTBF
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWIACRSIE    1S"
fgDisplay.Col = 1: fgDisplay = "CODE SIEGE"
fgDisplay.Col = 2: fgDisplay = lYSWIACR0.SWIACRSIE
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWIACRBQE    5A"
fgDisplay.Col = 1: fgDisplay = "CODE BANQUE"
fgDisplay.Col = 2: fgDisplay = lYSWIACR0.SWIACRBQE
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWIACRGUI    5A"
fgDisplay.Col = 1: fgDisplay = "CODE GUICHET"
fgDisplay.Col = 2: fgDisplay = lYSWIACR0.SWIACRGUI
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "SWIACRLI1   35A"
fgDisplay.Col = 1: fgDisplay = "LIBELLE 1"
fgDisplay.Col = 2: fgDisplay = lYSWIACR0.SWIACRLI1
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "SWIACRLI2   35A"
fgDisplay.Col = 1: fgDisplay = "LIBELLE 2"
fgDisplay.Col = 2: fgDisplay = lYSWIACR0.SWIACRLI2
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "SWIACRLI3   35A"
fgDisplay.Col = 1: fgDisplay = "LIBELLE 3"
fgDisplay.Col = 2: fgDisplay = lYSWIACR0.SWIACRLI3
fgDisplay.Row = 10
fgDisplay.Col = 0: fgDisplay = "SWIACRRSN   23A"
fgDisplay.Col = 1: fgDisplay = "RIB SNP"
fgDisplay.Col = 2: fgDisplay = lYSWIACR0.SWIACRRSN
fgDisplay.Row = 11
fgDisplay.Col = 0: fgDisplay = "SWIACRRTB   23A"
fgDisplay.Col = 1: fgDisplay = "RIB TBF"
fgDisplay.Col = 2: fgDisplay = lYSWIACR0.SWIACRRTB
End Sub
Public Sub srvYSWIALI0_fgDisplay(lYSWIALI0 As typeZSWIALI0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 11
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIALIETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWIALI0.SWIALIETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWIALIAGE    4B"
fgDisplay.Col = 1: fgDisplay = "AGENCE"
fgDisplay.Col = 2: fgDisplay = lYSWIALI0.SWIALIAGE
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWIALISER    2A"
fgDisplay.Col = 1: fgDisplay = "SERVICE"
fgDisplay.Col = 2: fgDisplay = lYSWIALI0.SWIALISER
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWIALISSE    2A"
fgDisplay.Col = 1: fgDisplay = "SERVICE"
fgDisplay.Col = 2: fgDisplay = lYSWIALI0.SWIALISSE
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWIALIMES    3A"
fgDisplay.Col = 1: fgDisplay = "TYPE MESSAGE"
fgDisplay.Col = 2: fgDisplay = lYSWIALI0.SWIALIMES
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWIALINUM    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWIALI0.SWIALINUM
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "SWIALINEN    1A"
fgDisplay.Col = 1: fgDisplay = "NUMER ENVOI"
fgDisplay.Col = 2: fgDisplay = lYSWIALI0.SWIALINEN
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "SWIALINLI    2P"
fgDisplay.Col = 1: fgDisplay = "NUMERO DE LIGNE"
fgDisplay.Col = 2: fgDisplay = lYSWIALI0.SWIALINLI
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "SWIALIDON  512A"
fgDisplay.Col = 1: fgDisplay = "DONNE MESSAGE"
fgDisplay.Col = 2: fgDisplay = lYSWIALI0.SWIALIDON
fgDisplay.Row = 10
fgDisplay.Col = 0: fgDisplay = "SWIALIOK    1A"
fgDisplay.Col = 1: fgDisplay = "PASSE SI OK"
fgDisplay.Col = 2: fgDisplay = lYSWIALI0.SWIALIOK
End Sub
Public Sub srvYSWIBUF0_fgDisplay(lYSWIBUF0 As typeZSWIBUF0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 8
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIBUFETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWIBUF0.SWIBUFETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWIBUFAGE    4B"
fgDisplay.Col = 1: fgDisplay = "AGENCE"
fgDisplay.Col = 2: fgDisplay = lYSWIBUF0.SWIBUFAGE
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWIBUFSER    2A"
fgDisplay.Col = 1: fgDisplay = "SERVICE"
fgDisplay.Col = 2: fgDisplay = lYSWIBUF0.SWIBUFSER
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWIBUFSSE    2A"
fgDisplay.Col = 1: fgDisplay = "SOUS-SERVICE"
fgDisplay.Col = 2: fgDisplay = lYSWIBUF0.SWIBUFSSE
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWIBUFREF   19A"
fgDisplay.Col = 1: fgDisplay = "REF.GLO.OPERATION"
fgDisplay.Col = 2: fgDisplay = lYSWIBUF0.SWIBUFREF
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWIBUFNLI    4P"
fgDisplay.Col = 1: fgDisplay = "NUMERO DE LIGNE"
fgDisplay.Col = 2: fgDisplay = lYSWIBUF0.SWIBUFNLI
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "SWIBUFDON  990A"
fgDisplay.Col = 1: fgDisplay = "DONNEE"
fgDisplay.Col = 2: fgDisplay = lYSWIBUF0.SWIBUFDON
End Sub
Public Sub srvYSWICCI0_fgDisplay(lYSWICCI0 As typeZSWICCI0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 11
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWICCIETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWICCI0.SWICCIETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWICCIAGE    4B"
fgDisplay.Col = 1: fgDisplay = "AGENCE"
fgDisplay.Col = 2: fgDisplay = lYSWICCI0.SWICCIAGE
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWICCISER    2A"
fgDisplay.Col = 1: fgDisplay = "SERVICE"
fgDisplay.Col = 2: fgDisplay = lYSWICCI0.SWICCISER
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWICCISSE    2A"
fgDisplay.Col = 1: fgDisplay = "SERVICE"
fgDisplay.Col = 2: fgDisplay = lYSWICCI0.SWICCISSE
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWICCIMES    3A"
fgDisplay.Col = 1: fgDisplay = "TYPE MESSAGE"
fgDisplay.Col = 2: fgDisplay = lYSWICCI0.SWICCIMES
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWICCINUM    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWICCI0.SWICCINUM
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "SWICCINEN    1A"
fgDisplay.Col = 1: fgDisplay = "NUMER ENVOI"
fgDisplay.Col = 2: fgDisplay = lYSWICCI0.SWICCINEN
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "SWICCINLI    2P"
fgDisplay.Col = 1: fgDisplay = "NUMERO DE LIGNE"
fgDisplay.Col = 2: fgDisplay = lYSWICCI0.SWICCINLI
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "SWICCIDON  512A"
fgDisplay.Col = 1: fgDisplay = "DONNE MESSAGE"
fgDisplay.Col = 2: fgDisplay = lYSWICCI0.SWICCIDON
fgDisplay.Row = 10
fgDisplay.Col = 0: fgDisplay = "SWICCIOK    1A"
fgDisplay.Col = 1: fgDisplay = "PASSE SI OK"
fgDisplay.Col = 2: fgDisplay = lYSWICCI0.SWICCIOK
End Sub
Public Sub srvYSWICLA0_fgDisplay(lYSWICLA0 As typeZSWICLA0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 13
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWICLAETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWICLA0.SWICLAETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWICLAAGE    4B"
fgDisplay.Col = 1: fgDisplay = "AGENCE"
fgDisplay.Col = 2: fgDisplay = lYSWICLA0.SWICLAAGE
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWICLASER    2A"
fgDisplay.Col = 1: fgDisplay = "SERVICE"
fgDisplay.Col = 2: fgDisplay = lYSWICLA0.SWICLASER
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWICLASES    2A"
fgDisplay.Col = 1: fgDisplay = "SOUS-SERVICE"
fgDisplay.Col = 2: fgDisplay = lYSWICLA0.SWICLASES
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWICLAOPR    3A"
fgDisplay.Col = 1: fgDisplay = "CODE OPERATION"
fgDisplay.Col = 2: fgDisplay = lYSWICLA0.SWICLAOPR
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWICLANUM    9P"
fgDisplay.Col = 1: fgDisplay = "NUMERO OPERATION"
fgDisplay.Col = 2: fgDisplay = lYSWICLA0.SWICLANUM
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "SWICLACLA    2A"
fgDisplay.Col = 1: fgDisplay = "CLASSE"
fgDisplay.Col = 2: fgDisplay = lYSWICLA0.SWICLACLA
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "SWICLAMES    3A"
fgDisplay.Col = 1: fgDisplay = "TYPE MESSAGE"
fgDisplay.Col = 2: fgDisplay = lYSWICLA0.SWICLAMES
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "SWICLACRI    3P"
fgDisplay.Col = 1: fgDisplay = "NUMRO ACCES"
fgDisplay.Col = 2: fgDisplay = lYSWICLA0.SWICLACRI
fgDisplay.Row = 10
fgDisplay.Col = 0: fgDisplay = "SWICLAREF   16A"
fgDisplay.Col = 1: fgDisplay = "IDENTIFICATEUR"
fgDisplay.Col = 2: fgDisplay = lYSWICLA0.SWICLAREF
fgDisplay.Row = 11
fgDisplay.Col = 0: fgDisplay = "SWICLAINT    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWICLA0.SWICLAINT
fgDisplay.Row = 12
fgDisplay.Col = 0: fgDisplay = "SWICLANEN    1A"
fgDisplay.Col = 1: fgDisplay = "NUMERO REENVOI"
fgDisplay.Col = 2: fgDisplay = lYSWICLA0.SWICLANEN
End Sub
Public Sub srvYSWICRI0_fgDisplay(lYSWICRI0 As typeZSWICRI0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 8
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWICRIETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWICRI0.SWICRIETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWICRICRI   25A"
fgDisplay.Col = 1: fgDisplay = "CRITERE DE SELEC."
fgDisplay.Col = 2: fgDisplay = lYSWICRI0.SWICRICRI
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWICRISEQ    2P"
fgDisplay.Col = 1: fgDisplay = "SEQUE DE MESSAGE"
fgDisplay.Col = 2: fgDisplay = lYSWICRI0.SWICRISEQ
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWICRIPRI    1A"
fgDisplay.Col = 1: fgDisplay = "PRIORITE 1-9"
fgDisplay.Col = 2: fgDisplay = lYSWICRI0.SWICRIPRI
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWICRINPR    2S"
fgDisplay.Col = 1: fgDisplay = "NUM PAR COD.PRIOR"
fgDisplay.Col = 2: fgDisplay = lYSWICRI0.SWICRINPR
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWICRITYP    1A"
fgDisplay.Col = 1: fgDisplay = "TYPE LIG.1-2-3"
fgDisplay.Col = 2: fgDisplay = lYSWICRI0.SWICRITYP
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "SWICRIDON   97A"
fgDisplay.Col = 1: fgDisplay = "DONNE DIFF.TYPES"
fgDisplay.Col = 2: fgDisplay = lYSWICRI0.SWICRIDON
End Sub
Public Sub srvYSWIECA0_fgDisplay(lYSWIECA0 As typeZSWIECA0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 17
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIECAETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWIECA0.SWIECAETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWIECAREF   16A"
fgDisplay.Col = 1: fgDisplay = "REFERNECE"
fgDisplay.Col = 2: fgDisplay = lYSWIECA0.SWIECAREF
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWIECAMES    3A"
fgDisplay.Col = 1: fgDisplay = "TYPE MESSAGE"
fgDisplay.Col = 2: fgDisplay = lYSWIECA0.SWIECAMES
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWIECAPRI    2A"
fgDisplay.Col = 1: fgDisplay = "PRIORITE MESSAGE"
fgDisplay.Col = 2: fgDisplay = lYSWIECA0.SWIECAPRI
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWIECAEME   12A"
fgDisplay.Col = 1: fgDisplay = "EMETTEUR"
fgDisplay.Col = 2: fgDisplay = lYSWIECA0.SWIECAEME
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWIECADVA    7P"
fgDisplay.Col = 1: fgDisplay = "DATE VALEUR"
fgDisplay.Col = 2: fgDisplay = lYSWIECA0.SWIECADVA
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "SWIECADE1    3A"
fgDisplay.Col = 1: fgDisplay = "DEVISE 1"
fgDisplay.Col = 2: fgDisplay = lYSWIECA0.SWIECADE1
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "SWIECAMON 15.2P"
fgDisplay.Col = 1: fgDisplay = "DATE VALEUR"
fgDisplay.Col = 2: fgDisplay = lYSWIECA0.SWIECAMON
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "SWIECADRE    7P"
fgDisplay.Col = 1: fgDisplay = "DATE RECEPTION"
fgDisplay.Col = 2: fgDisplay = lYSWIECA0.SWIECADRE
fgDisplay.Row = 10
fgDisplay.Col = 0: fgDisplay = "SWIECAHRE    7P"
fgDisplay.Col = 1: fgDisplay = "HEURE RECEPTION"
fgDisplay.Col = 2: fgDisplay = lYSWIECA0.SWIECAHRE
fgDisplay.Row = 11
fgDisplay.Col = 0: fgDisplay = "SWIECAINT    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWIECA0.SWIECAINT
fgDisplay.Row = 12
fgDisplay.Col = 0: fgDisplay = "SWIECACET    1A"
fgDisplay.Col = 1: fgDisplay = "CODE ETAT"
fgDisplay.Col = 2: fgDisplay = lYSWIECA0.SWIECACET
fgDisplay.Row = 13
fgDisplay.Col = 0: fgDisplay = "SWIECAAGE    4B"
fgDisplay.Col = 1: fgDisplay = "AGENCE"
fgDisplay.Col = 2: fgDisplay = lYSWIECA0.SWIECAAGE
fgDisplay.Row = 14
fgDisplay.Col = 0: fgDisplay = "SWIECASER    2A"
fgDisplay.Col = 1: fgDisplay = "SERVICE"
fgDisplay.Col = 2: fgDisplay = lYSWIECA0.SWIECASER
fgDisplay.Row = 15
fgDisplay.Col = 0: fgDisplay = "SWIECASSE    2A"
fgDisplay.Col = 1: fgDisplay = "SOUS SERVICE"
fgDisplay.Col = 2: fgDisplay = lYSWIECA0.SWIECASSE
fgDisplay.Row = 16
fgDisplay.Col = 0: fgDisplay = "SWIECAUTI   10A"
fgDisplay.Col = 1: fgDisplay = "UTILISATEUR"
fgDisplay.Col = 2: fgDisplay = lYSWIECA0.SWIECAUTI
End Sub
Public Sub srvYSWIECB0_fgDisplay(lYSWIECB0 As typeZSWIECB0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 10
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIECBETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWIECB0.SWIECBETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWIECBNUM    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWIECB0.SWIECBNUM
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWIECBNOR    3P"
fgDisplay.Col = 1: fgDisplay = "ORDRE"
fgDisplay.Col = 2: fgDisplay = lYSWIECB0.SWIECBNOR
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWIECBCHA    2P"
fgDisplay.Col = 1: fgDisplay = "CHAMP"
fgDisplay.Col = 2: fgDisplay = lYSWIECB0.SWIECBCHA
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWIECBIND    2A"
fgDisplay.Col = 1: fgDisplay = "INDICE"
fgDisplay.Col = 2: fgDisplay = lYSWIECB0.SWIECBIND
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWIECBZON    2P"
fgDisplay.Col = 1: fgDisplay = "ZONE"
fgDisplay.Col = 2: fgDisplay = lYSWIECB0.SWIECBZON
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "SWIECBSZO    2P"
fgDisplay.Col = 1: fgDisplay = "SOUS ZONE"
fgDisplay.Col = 2: fgDisplay = lYSWIECB0.SWIECBSZO
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "SWIECBINR    1A"
fgDisplay.Col = 1: fgDisplay = "INDICE REELL"
fgDisplay.Col = 2: fgDisplay = lYSWIECB0.SWIECBINR
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "SWIECBVAL   65A"
fgDisplay.Col = 1: fgDisplay = "VALEUR ZONE"
fgDisplay.Col = 2: fgDisplay = lYSWIECB0.SWIECBVAL
End Sub
Public Sub srvYSWIEHA0_fgDisplay(lYSWIEHA0 As typeZSWIEHA0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 17
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIEHAETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWIEHA0.SWIEHAETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWIEHANUM    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWIEHA0.SWIEHANUM
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWIEHAREF   16A"
fgDisplay.Col = 1: fgDisplay = "REFERENCE"
fgDisplay.Col = 2: fgDisplay = lYSWIEHA0.SWIEHAREF
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWIEHAMES    3A"
fgDisplay.Col = 1: fgDisplay = "TYPE MESSAGE"
fgDisplay.Col = 2: fgDisplay = lYSWIEHA0.SWIEHAMES
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWIEHAEME   12A"
fgDisplay.Col = 1: fgDisplay = "EMETTEUR"
fgDisplay.Col = 2: fgDisplay = lYSWIEHA0.SWIEHAEME
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWIEHADRE    7P"
fgDisplay.Col = 1: fgDisplay = "DATE RECEPTION"
fgDisplay.Col = 2: fgDisplay = lYSWIEHA0.SWIEHADRE
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "SWIEHAHRE    7P"
fgDisplay.Col = 1: fgDisplay = "HEURE RECEPTION"
fgDisplay.Col = 2: fgDisplay = lYSWIEHA0.SWIEHAHRE
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "SWIEHAAGE    4B"
fgDisplay.Col = 1: fgDisplay = "AGENCE"
fgDisplay.Col = 2: fgDisplay = lYSWIEHA0.SWIEHAAGE
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "SWIEHASER    2A"
fgDisplay.Col = 1: fgDisplay = "SERVICE"
fgDisplay.Col = 2: fgDisplay = lYSWIEHA0.SWIEHASER
fgDisplay.Row = 10
fgDisplay.Col = 0: fgDisplay = "SWIEHASSE    2A"
fgDisplay.Col = 1: fgDisplay = "SOUS SERVICE"
fgDisplay.Col = 2: fgDisplay = lYSWIEHA0.SWIEHASSE
fgDisplay.Row = 11
fgDisplay.Col = 0: fgDisplay = "SWIEHAUTI   10A"
fgDisplay.Col = 1: fgDisplay = "UTILISATEUR"
fgDisplay.Col = 2: fgDisplay = lYSWIEHA0.SWIEHAUTI
fgDisplay.Row = 12
fgDisplay.Col = 0: fgDisplay = "SWIEHADTR    7P"
fgDisplay.Col = 1: fgDisplay = "DATE TRAITEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWIEHA0.SWIEHADTR
fgDisplay.Row = 13
fgDisplay.Col = 0: fgDisplay = "SWIEHAAVI    1A"
fgDisplay.Col = 1: fgDisplay = "AVIS EDITE O/N"
fgDisplay.Col = 2: fgDisplay = lYSWIEHA0.SWIEHAAVI
fgDisplay.Row = 14
fgDisplay.Col = 0: fgDisplay = "SWIEHADVA    7P"
fgDisplay.Col = 1: fgDisplay = "DATE VALEUR"
fgDisplay.Col = 2: fgDisplay = lYSWIEHA0.SWIEHADVA
fgDisplay.Row = 15
fgDisplay.Col = 0: fgDisplay = "SWIEHADEV    3A"
fgDisplay.Col = 1: fgDisplay = "DEVISE 1"
fgDisplay.Col = 2: fgDisplay = lYSWIEHA0.SWIEHADEV
fgDisplay.Row = 16
fgDisplay.Col = 0: fgDisplay = "SWIEHAMON 15.2P"
fgDisplay.Col = 1: fgDisplay = "MONTANT"
fgDisplay.Col = 2: fgDisplay = lYSWIEHA0.SWIEHAMON
End Sub
Public Sub srvYSWIEHB0_fgDisplay(lYSWIEHB0 As typeZSWIEHB0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 10
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIEHBETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWIEHB0.SWIEHBETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWIEHBNUM    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWIEHB0.SWIEHBNUM
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWIEHBNOR    3P"
fgDisplay.Col = 1: fgDisplay = "ORDRE"
fgDisplay.Col = 2: fgDisplay = lYSWIEHB0.SWIEHBNOR
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWIEHBCHA    2P"
fgDisplay.Col = 1: fgDisplay = "CHAMP"
fgDisplay.Col = 2: fgDisplay = lYSWIEHB0.SWIEHBCHA
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWIEHBIND    2A"
fgDisplay.Col = 1: fgDisplay = "INDICE"
fgDisplay.Col = 2: fgDisplay = lYSWIEHB0.SWIEHBIND
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWIEHBZON    2P"
fgDisplay.Col = 1: fgDisplay = "ZONE"
fgDisplay.Col = 2: fgDisplay = lYSWIEHB0.SWIEHBZON
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "SWIEHBSZO    2P"
fgDisplay.Col = 1: fgDisplay = "SOUS ZONE"
fgDisplay.Col = 2: fgDisplay = lYSWIEHB0.SWIEHBSZO
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "SWIEHBINR    1A"
fgDisplay.Col = 1: fgDisplay = "INDICE REELL"
fgDisplay.Col = 2: fgDisplay = lYSWIEHB0.SWIEHBINR
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "SWIEHBVAL   65A"
fgDisplay.Col = 1: fgDisplay = "VALEUR ZONE"
fgDisplay.Col = 2: fgDisplay = lYSWIEHB0.SWIEHBVAL
End Sub
Public Sub srvYSWIENA0_fgDisplay(lYSWIENA0 As typeZSWIENA0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 17
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIENAETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWIENA0.SWIENAETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWIENAREF   16A"
fgDisplay.Col = 1: fgDisplay = "REFERNECE"
fgDisplay.Col = 2: fgDisplay = lYSWIENA0.SWIENAREF
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWIENAMES    3A"
fgDisplay.Col = 1: fgDisplay = "TYPE MESSAGE"
fgDisplay.Col = 2: fgDisplay = lYSWIENA0.SWIENAMES
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWIENAPRI    2A"
fgDisplay.Col = 1: fgDisplay = "PRIORITE MESSAGE"
fgDisplay.Col = 2: fgDisplay = lYSWIENA0.SWIENAPRI
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWIENAEME   12A"
fgDisplay.Col = 1: fgDisplay = "EMETTEUR"
fgDisplay.Col = 2: fgDisplay = lYSWIENA0.SWIENAEME
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWIENADVA    7P"
fgDisplay.Col = 1: fgDisplay = "DATE VALEUR"
fgDisplay.Col = 2: fgDisplay = lYSWIENA0.SWIENADVA
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "SWIENADE1    3A"
fgDisplay.Col = 1: fgDisplay = "DEVISE 1"
fgDisplay.Col = 2: fgDisplay = lYSWIENA0.SWIENADE1
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "SWIENAMON 15.2P"
fgDisplay.Col = 1: fgDisplay = "DATE VALEUR"
fgDisplay.Col = 2: fgDisplay = lYSWIENA0.SWIENAMON
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "SWIENADRE    7P"
fgDisplay.Col = 1: fgDisplay = "DATE RECEPTION"
fgDisplay.Col = 2: fgDisplay = lYSWIENA0.SWIENADRE
fgDisplay.Row = 10
fgDisplay.Col = 0: fgDisplay = "SWIENAHRE    7P"
fgDisplay.Col = 1: fgDisplay = "HEURE RECEPTION"
fgDisplay.Col = 2: fgDisplay = lYSWIENA0.SWIENAHRE
fgDisplay.Row = 11
fgDisplay.Col = 0: fgDisplay = "SWIENAINT    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWIENA0.SWIENAINT
fgDisplay.Row = 12
fgDisplay.Col = 0: fgDisplay = "SWIENACET    1A"
fgDisplay.Col = 1: fgDisplay = "CODE ETAT"
fgDisplay.Col = 2: fgDisplay = lYSWIENA0.SWIENACET
fgDisplay.Row = 13
fgDisplay.Col = 0: fgDisplay = "SWIENAAGE    4B"
fgDisplay.Col = 1: fgDisplay = "AGENCE"
fgDisplay.Col = 2: fgDisplay = lYSWIENA0.SWIENAAGE
fgDisplay.Row = 14
fgDisplay.Col = 0: fgDisplay = "SWIENASER    2A"
fgDisplay.Col = 1: fgDisplay = "SERVICE"
fgDisplay.Col = 2: fgDisplay = lYSWIENA0.SWIENASER
fgDisplay.Row = 15
fgDisplay.Col = 0: fgDisplay = "SWIENASSE    2A"
fgDisplay.Col = 1: fgDisplay = "SOUS SERVICE"
fgDisplay.Col = 2: fgDisplay = lYSWIENA0.SWIENASSE
fgDisplay.Row = 16
fgDisplay.Col = 0: fgDisplay = "SWIENAUTI   10A"
fgDisplay.Col = 1: fgDisplay = "UTILISATEUR"
fgDisplay.Col = 2: fgDisplay = lYSWIENA0.SWIENAUTI
End Sub
Public Sub srvYSWIENB0_fgDisplay(lYSWIENB0 As typeZSWIENB0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 10
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIENBETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWIENB0.SWIENBETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWIENBNUM    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWIENB0.SWIENBNUM
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWIENBNOR    3P"
fgDisplay.Col = 1: fgDisplay = "ORDRE"
fgDisplay.Col = 2: fgDisplay = lYSWIENB0.SWIENBNOR
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWIENBCHA    2P"
fgDisplay.Col = 1: fgDisplay = "CHAMP"
fgDisplay.Col = 2: fgDisplay = lYSWIENB0.SWIENBCHA
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWIENBIND    2A"
fgDisplay.Col = 1: fgDisplay = "INDICE"
fgDisplay.Col = 2: fgDisplay = lYSWIENB0.SWIENBIND
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWIENBZON    2P"
fgDisplay.Col = 1: fgDisplay = "ZONE"
fgDisplay.Col = 2: fgDisplay = lYSWIENB0.SWIENBZON
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "SWIENBSZO    2P"
fgDisplay.Col = 1: fgDisplay = "SOUS ZONE"
fgDisplay.Col = 2: fgDisplay = lYSWIENB0.SWIENBSZO
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "SWIENBINR    1A"
fgDisplay.Col = 1: fgDisplay = "INDICE REELL"
fgDisplay.Col = 2: fgDisplay = lYSWIENB0.SWIENBINR
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "SWIENBVAL   65A"
fgDisplay.Col = 1: fgDisplay = "VALEUR ZONE"
fgDisplay.Col = 2: fgDisplay = lYSWIENB0.SWIENBVAL
End Sub
Public Sub srvYSWIENI0_fgDisplay(lYSWIENI0 As typeZSWIENI0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 11
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIENIETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWIENI0.SWIENIETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWIENIAGE    4B"
fgDisplay.Col = 1: fgDisplay = "AGENCE"
fgDisplay.Col = 2: fgDisplay = lYSWIENI0.SWIENIAGE
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWIENISER    2A"
fgDisplay.Col = 1: fgDisplay = "SERVICE"
fgDisplay.Col = 2: fgDisplay = lYSWIENI0.SWIENISER
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWIENISSE    2A"
fgDisplay.Col = 1: fgDisplay = "SERVICE"
fgDisplay.Col = 2: fgDisplay = lYSWIENI0.SWIENISSE
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWIENIMES    3A"
fgDisplay.Col = 1: fgDisplay = "TYPE MESSAGE"
fgDisplay.Col = 2: fgDisplay = lYSWIENI0.SWIENIMES
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWIENINUM    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWIENI0.SWIENINUM
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "SWIENINEN    1A"
fgDisplay.Col = 1: fgDisplay = "NUMER ENVOI"
fgDisplay.Col = 2: fgDisplay = lYSWIENI0.SWIENINEN
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "SWIENINLI    2P"
fgDisplay.Col = 1: fgDisplay = "NUMERO DE LIGNE"
fgDisplay.Col = 2: fgDisplay = lYSWIENI0.SWIENINLI
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "SWIENIDON  250A"
fgDisplay.Col = 1: fgDisplay = "DONNE MESSAGE"
fgDisplay.Col = 2: fgDisplay = lYSWIENI0.SWIENIDON
fgDisplay.Row = 10
fgDisplay.Col = 0: fgDisplay = "SWIENIOK    1A"
fgDisplay.Col = 1: fgDisplay = "PASSE SI OK"
fgDisplay.Col = 2: fgDisplay = lYSWIENI0.SWIENIOK
End Sub
Public Sub srvYSWIEVC0_fgDisplay(lYSWIEVC0 As typeZSWIEVC0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 2
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIEVCDON  512A"
fgDisplay.Col = 1: fgDisplay = "DONNE MESSAGE"
fgDisplay.Col = 2: fgDisplay = lYSWIEVC0.SWIEVCDON
End Sub
Public Sub srvYSWIEVI0_fgDisplay(lYSWIEVI0 As typeZSWIEVI0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 11
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIEVIETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWIEVI0.SWIEVIETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWIEVIAGE    4B"
fgDisplay.Col = 1: fgDisplay = "AGENCE"
fgDisplay.Col = 2: fgDisplay = lYSWIEVI0.SWIEVIAGE
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWIEVISER    2A"
fgDisplay.Col = 1: fgDisplay = "SERVICE"
fgDisplay.Col = 2: fgDisplay = lYSWIEVI0.SWIEVISER
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWIEVISSE    2A"
fgDisplay.Col = 1: fgDisplay = "SERVICE"
fgDisplay.Col = 2: fgDisplay = lYSWIEVI0.SWIEVISSE
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWIEVIMES    3A"
fgDisplay.Col = 1: fgDisplay = "TYPE MESSAGE"
fgDisplay.Col = 2: fgDisplay = lYSWIEVI0.SWIEVIMES
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWIEVINUM    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWIEVI0.SWIEVINUM
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "SWIEVINEN    1A"
fgDisplay.Col = 1: fgDisplay = "NUMER ENVOI"
fgDisplay.Col = 2: fgDisplay = lYSWIEVI0.SWIEVINEN
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "SWIEVINLI    2P"
fgDisplay.Col = 1: fgDisplay = "NUMERO DE LIGNE"
fgDisplay.Col = 2: fgDisplay = lYSWIEVI0.SWIEVINLI
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "SWIEVIDON  512A"
fgDisplay.Col = 1: fgDisplay = "DONNE MESSAGE"
fgDisplay.Col = 2: fgDisplay = lYSWIEVI0.SWIEVIDON
fgDisplay.Row = 10
fgDisplay.Col = 0: fgDisplay = "SWIEVIOK    1A"
fgDisplay.Col = 1: fgDisplay = "PASSE SI OK"
fgDisplay.Col = 2: fgDisplay = lYSWIEVI0.SWIEVIOK
End Sub
Public Sub srvYSWIFTA0_fgDisplay(lYSWIFTA0 As typeZSWIFTA0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 26
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIFTAETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWIFTA0.SWIFTAETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWIFTAREF   16A"
fgDisplay.Col = 1: fgDisplay = "REFERNECE"
fgDisplay.Col = 2: fgDisplay = lYSWIFTA0.SWIFTAREF
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWIFTANEN    1A"
fgDisplay.Col = 1: fgDisplay = "NUMERO DE RENVOI"
fgDisplay.Col = 2: fgDisplay = lYSWIFTA0.SWIFTANEN
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWIFTAPRI    2A"
fgDisplay.Col = 1: fgDisplay = "CODE PROIRITE"
fgDisplay.Col = 2: fgDisplay = lYSWIFTA0.SWIFTAPRI
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWIFTAMES    3A"
fgDisplay.Col = 1: fgDisplay = "TYPE MESSAGE"
fgDisplay.Col = 2: fgDisplay = lYSWIFTA0.SWIFTAMES
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWIFTADOR   12A"
fgDisplay.Col = 1: fgDisplay = "DONNEUR ORDRE"
fgDisplay.Col = 2: fgDisplay = lYSWIFTA0.SWIFTADOR
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "SWIFTADES   12A"
fgDisplay.Col = 1: fgDisplay = "DESTINATAIRE"
fgDisplay.Col = 2: fgDisplay = lYSWIFTA0.SWIFTADES
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "SWIFTADVA    7P"
fgDisplay.Col = 1: fgDisplay = "DATE VALEUR"
fgDisplay.Col = 2: fgDisplay = lYSWIFTA0.SWIFTADVA
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "SWIFTADE1    3A"
fgDisplay.Col = 1: fgDisplay = "DEVISE 1"
fgDisplay.Col = 2: fgDisplay = lYSWIFTA0.SWIFTADE1
fgDisplay.Row = 10
fgDisplay.Col = 0: fgDisplay = "SWIFTAMON 15.2P"
fgDisplay.Col = 1: fgDisplay = "DATE VALEUR"
fgDisplay.Col = 2: fgDisplay = lYSWIFTA0.SWIFTAMON
fgDisplay.Row = 11
fgDisplay.Col = 0: fgDisplay = "SWIFTADE2    3A"
fgDisplay.Col = 1: fgDisplay = "DEVISE 2"
fgDisplay.Col = 2: fgDisplay = lYSWIFTA0.SWIFTADE2
fgDisplay.Row = 12
fgDisplay.Col = 0: fgDisplay = "SWIFTADEN    7P"
fgDisplay.Col = 1: fgDisplay = "DATE ENVOI"
fgDisplay.Col = 2: fgDisplay = lYSWIFTA0.SWIFTADEN
fgDisplay.Row = 13
fgDisplay.Col = 0: fgDisplay = "SWIFTAHEN    7P"
fgDisplay.Col = 1: fgDisplay = "HEURE ENVOI"
fgDisplay.Col = 2: fgDisplay = lYSWIFTA0.SWIFTAHEN
fgDisplay.Row = 14
fgDisplay.Col = 0: fgDisplay = "SWIFTACOM    1A"
fgDisplay.Col = 1: fgDisplay = "COMPLET"
fgDisplay.Col = 2: fgDisplay = lYSWIFTA0.SWIFTACOM
fgDisplay.Row = 15
fgDisplay.Col = 0: fgDisplay = "SWIFTATES    1A"
fgDisplay.Col = 1: fgDisplay = "TEST OU REEL"
fgDisplay.Col = 2: fgDisplay = lYSWIFTA0.SWIFTATES
fgDisplay.Row = 16
fgDisplay.Col = 0: fgDisplay = "SWIFTASUP    1A"
fgDisplay.Col = 1: fgDisplay = "SUPPRIME"
fgDisplay.Col = 2: fgDisplay = lYSWIFTA0.SWIFTASUP
fgDisplay.Row = 17
fgDisplay.Col = 0: fgDisplay = "SWIFTAVAL    1A"
fgDisplay.Col = 1: fgDisplay = "TOP VALIDATION"
fgDisplay.Col = 2: fgDisplay = lYSWIFTA0.SWIFTAVAL
fgDisplay.Row = 18
fgDisplay.Col = 0: fgDisplay = "SWIFTAAGE    4B"
fgDisplay.Col = 1: fgDisplay = "AGENCE"
fgDisplay.Col = 2: fgDisplay = lYSWIFTA0.SWIFTAAGE
fgDisplay.Row = 19
fgDisplay.Col = 0: fgDisplay = "SWIFTASER    2A"
fgDisplay.Col = 1: fgDisplay = "SERVICE"
fgDisplay.Col = 2: fgDisplay = lYSWIFTA0.SWIFTASER
fgDisplay.Row = 20
fgDisplay.Col = 0: fgDisplay = "SWIFTASSE    2A"
fgDisplay.Col = 1: fgDisplay = "SOUS SERVICE"
fgDisplay.Col = 2: fgDisplay = lYSWIFTA0.SWIFTASSE
fgDisplay.Row = 21
fgDisplay.Col = 0: fgDisplay = "SWIFTAUTI   10A"
fgDisplay.Col = 1: fgDisplay = "UTILISATEUR"
fgDisplay.Col = 2: fgDisplay = lYSWIFTA0.SWIFTAUTI
fgDisplay.Row = 22
fgDisplay.Col = 0: fgDisplay = "SWIFTANUM    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWIFTA0.SWIFTANUM
fgDisplay.Row = 23
fgDisplay.Col = 0: fgDisplay = "SWIFTAUT1   10A"
fgDisplay.Col = 1: fgDisplay = "UTILISA SAISIE"
fgDisplay.Col = 2: fgDisplay = lYSWIFTA0.SWIFTAUT1
fgDisplay.Row = 24
fgDisplay.Col = 0: fgDisplay = "SWIFTAPVA    1A"
fgDisplay.Col = 1: fgDisplay = "1ERE VALIDATION"
fgDisplay.Col = 2: fgDisplay = lYSWIFTA0.SWIFTAPVA
fgDisplay.Row = 25
fgDisplay.Col = 0: fgDisplay = "SWIFTAUT2   10A"
fgDisplay.Col = 1: fgDisplay = "UTILISA 1ER VALID"
fgDisplay.Col = 2: fgDisplay = lYSWIFTA0.SWIFTAUT2
End Sub
Public Sub srvYSWIFTB0_fgDisplay(lYSWIFTB0 As typeZSWIFTB0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 6
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIFTBETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWIFTB0.SWIFTBETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWIFTBNUM    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWIFTB0.SWIFTBNUM
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWIFTBNEN    1S"
fgDisplay.Col = 1: fgDisplay = "NUMERO ENVOI"
fgDisplay.Col = 2: fgDisplay = lYSWIFTB0.SWIFTBNEN
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWIFTBNLI    4P"
fgDisplay.Col = 1: fgDisplay = "NUMERO LIGNE"
fgDisplay.Col = 2: fgDisplay = lYSWIFTB0.SWIFTBNLI
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWIFTBDET   70A"
fgDisplay.Col = 1: fgDisplay = "DETAIL"
fgDisplay.Col = 2: fgDisplay = lYSWIFTB0.SWIFTBDET
End Sub
Public Sub srvYSWIFTC0_fgDisplay(lYSWIFTC0 As typeZSWIFTC0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 19
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIFTCETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWIFTC0.SWIFTCETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWIFTCNUM    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWIFTC0.SWIFTCNUM
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWIFTCNEN    1S"
fgDisplay.Col = 1: fgDisplay = "NUMERO ENVOI"
fgDisplay.Col = 2: fgDisplay = lYSWIFTC0.SWIFTCNEN
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWIFTCNLI    4P"
fgDisplay.Col = 1: fgDisplay = "NUMERO LIGNE"
fgDisplay.Col = 2: fgDisplay = lYSWIFTC0.SWIFTCNLI
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWIFTCNSE   40A"
fgDisplay.Col = 1: fgDisplay = "NUMERO SEQUENCE"
fgDisplay.Col = 2: fgDisplay = lYSWIFTC0.SWIFTCNSE
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWIFTCSOC    4P"
fgDisplay.Col = 1: fgDisplay = "NUM OCC SEQUE"
fgDisplay.Col = 2: fgDisplay = lYSWIFTC0.SWIFTCSOC
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "SWIFTCNCH    2P"
fgDisplay.Col = 1: fgDisplay = "NUMERO CHAMP"
fgDisplay.Col = 2: fgDisplay = lYSWIFTC0.SWIFTCNCH
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "SWIFTCCOC    4P"
fgDisplay.Col = 1: fgDisplay = "NUM OCC CHAMP"
fgDisplay.Col = 2: fgDisplay = lYSWIFTC0.SWIFTCCOC
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "SWIFTCNLC    4P"
fgDisplay.Col = 1: fgDisplay = "NUMERO LIGNE CHAM"
fgDisplay.Col = 2: fgDisplay = lYSWIFTC0.SWIFTCNLC
fgDisplay.Row = 10
fgDisplay.Col = 0: fgDisplay = "SWIFTCSEQ    2A"
fgDisplay.Col = 1: fgDisplay = "DESCRIP SEQUENCE"
fgDisplay.Col = 2: fgDisplay = lYSWIFTC0.SWIFTCSEQ
fgDisplay.Row = 11
fgDisplay.Col = 0: fgDisplay = "SWIFTCCHA    2A"
fgDisplay.Col = 1: fgDisplay = "DESCRIP CHAMP"
fgDisplay.Col = 2: fgDisplay = lYSWIFTC0.SWIFTCCHA
fgDisplay.Row = 12
fgDisplay.Col = 0: fgDisplay = "SWIFTCILI    1A"
fgDisplay.Col = 1: fgDisplay = "INDICATEUR DEB"
fgDisplay.Col = 2: fgDisplay = lYSWIFTC0.SWIFTCILI
fgDisplay.Row = 13
fgDisplay.Col = 0: fgDisplay = "SWIFTCFAC    1A"
fgDisplay.Col = 1: fgDisplay = "FACULTATIF"
fgDisplay.Col = 2: fgDisplay = lYSWIFTC0.SWIFTCFAC
fgDisplay.Row = 14
fgDisplay.Col = 0: fgDisplay = "SWIFTCSIG    1A"
fgDisplay.Col = 1: fgDisplay = "SIGNE COMPLET"
fgDisplay.Col = 2: fgDisplay = lYSWIFTC0.SWIFTCSIG
fgDisplay.Row = 15
fgDisplay.Col = 0: fgDisplay = "SWIFTCSMA    4P"
fgDisplay.Col = 1: fgDisplay = "OCCUR SEQ MAXIMUM"
fgDisplay.Col = 2: fgDisplay = lYSWIFTC0.SWIFTCSMA
fgDisplay.Row = 16
fgDisplay.Col = 0: fgDisplay = "SWIFTCCMA    4P"
fgDisplay.Col = 1: fgDisplay = "OCCUR CHA MAXIMUM"
fgDisplay.Col = 2: fgDisplay = lYSWIFTC0.SWIFTCCMA
fgDisplay.Row = 17
fgDisplay.Col = 0: fgDisplay = "SWIFTCSMI    2P"
fgDisplay.Col = 1: fgDisplay = "OCCUR SEQ MINIMUM"
fgDisplay.Col = 2: fgDisplay = lYSWIFTC0.SWIFTCSMI
fgDisplay.Row = 18
fgDisplay.Col = 0: fgDisplay = "SWIFTCCMI    2P"
fgDisplay.Col = 1: fgDisplay = "OCCUR CHA MINIMUM"
fgDisplay.Col = 2: fgDisplay = lYSWIFTC0.SWIFTCCMI
End Sub
Public Sub srvYSWIGRN0_fgDisplay(lYSWIGRN0 As typeZSWIGRN0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 5
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIGRNETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWIGRN0.SWIGRNETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWIGRNGRP    6A"
fgDisplay.Col = 1: fgDisplay = "GRP NATURE"
fgDisplay.Col = 2: fgDisplay = lYSWIGRN0.SWIGRNGRP
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWIGRNORD    3P"
fgDisplay.Col = 1: fgDisplay = "NUMERO PAR NATURE"
fgDisplay.Col = 2: fgDisplay = lYSWIGRN0.SWIGRNORD
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWIGRNNAT    6A"
fgDisplay.Col = 1: fgDisplay = "NATURE"
fgDisplay.Col = 2: fgDisplay = lYSWIGRN0.SWIGRNNAT
End Sub
Public Sub srvYSWIHIA0_fgDisplay(lYSWIHIA0 As typeZSWIHIA0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 26
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIHIAETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWIHIA0.SWIHIAETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWIHIAREF   16A"
fgDisplay.Col = 1: fgDisplay = "REFERNECE"
fgDisplay.Col = 2: fgDisplay = lYSWIHIA0.SWIHIAREF
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWIHIANEN    1A"
fgDisplay.Col = 1: fgDisplay = "NUMERO DE RENVOI"
fgDisplay.Col = 2: fgDisplay = lYSWIHIA0.SWIHIANEN
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWIHIAPRI    2A"
fgDisplay.Col = 1: fgDisplay = "CODE PROIRITE"
fgDisplay.Col = 2: fgDisplay = lYSWIHIA0.SWIHIAPRI
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWIHIAMES    3A"
fgDisplay.Col = 1: fgDisplay = "TYPE MESSAGE"
fgDisplay.Col = 2: fgDisplay = lYSWIHIA0.SWIHIAMES
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWIHIADOR   12A"
fgDisplay.Col = 1: fgDisplay = "DONNEUR ORDRE"
fgDisplay.Col = 2: fgDisplay = lYSWIHIA0.SWIHIADOR
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "SWIHIADES   12A"
fgDisplay.Col = 1: fgDisplay = "DESTINATAIRE"
fgDisplay.Col = 2: fgDisplay = lYSWIHIA0.SWIHIADES
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "SWIHIADVA    7P"
fgDisplay.Col = 1: fgDisplay = "DATE VALEUR"
fgDisplay.Col = 2: fgDisplay = lYSWIHIA0.SWIHIADVA
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "SWIHIADE1    3A"
fgDisplay.Col = 1: fgDisplay = "DEVISE 1"
fgDisplay.Col = 2: fgDisplay = lYSWIHIA0.SWIHIADE1
fgDisplay.Row = 10
fgDisplay.Col = 0: fgDisplay = "SWIHIAMON 15.2P"
fgDisplay.Col = 1: fgDisplay = "MONTANT"
fgDisplay.Col = 2: fgDisplay = lYSWIHIA0.SWIHIAMON
fgDisplay.Row = 11
fgDisplay.Col = 0: fgDisplay = "SWIHIADE2    3A"
fgDisplay.Col = 1: fgDisplay = "DEVISE 2"
fgDisplay.Col = 2: fgDisplay = lYSWIHIA0.SWIHIADE2
fgDisplay.Row = 12
fgDisplay.Col = 0: fgDisplay = "SWIHIADEN    7P"
fgDisplay.Col = 1: fgDisplay = "DATE ENVOI"
fgDisplay.Col = 2: fgDisplay = lYSWIHIA0.SWIHIADEN
fgDisplay.Row = 13
fgDisplay.Col = 0: fgDisplay = "SWIHIAHEN    7P"
fgDisplay.Col = 1: fgDisplay = "HEURE ENVOI"
fgDisplay.Col = 2: fgDisplay = lYSWIHIA0.SWIHIAHEN
fgDisplay.Row = 14
fgDisplay.Col = 0: fgDisplay = "SWIHIACOM    1A"
fgDisplay.Col = 1: fgDisplay = "COMPLET"
fgDisplay.Col = 2: fgDisplay = lYSWIHIA0.SWIHIACOM
fgDisplay.Row = 15
fgDisplay.Col = 0: fgDisplay = "SWIHIATES    1A"
fgDisplay.Col = 1: fgDisplay = "TEST OU REEL"
fgDisplay.Col = 2: fgDisplay = lYSWIHIA0.SWIHIATES
fgDisplay.Row = 16
fgDisplay.Col = 0: fgDisplay = "SWIHIASUP    1A"
fgDisplay.Col = 1: fgDisplay = "SUPPRIME"
fgDisplay.Col = 2: fgDisplay = lYSWIHIA0.SWIHIASUP
fgDisplay.Row = 17
fgDisplay.Col = 0: fgDisplay = "SWIHIAVAL    1A"
fgDisplay.Col = 1: fgDisplay = "TOP VALIDATION"
fgDisplay.Col = 2: fgDisplay = lYSWIHIA0.SWIHIAVAL
fgDisplay.Row = 18
fgDisplay.Col = 0: fgDisplay = "SWIHIAAGE    4B"
fgDisplay.Col = 1: fgDisplay = "AGENCE"
fgDisplay.Col = 2: fgDisplay = lYSWIHIA0.SWIHIAAGE
fgDisplay.Row = 19
fgDisplay.Col = 0: fgDisplay = "SWIHIASER    2A"
fgDisplay.Col = 1: fgDisplay = "SERVICE"
fgDisplay.Col = 2: fgDisplay = lYSWIHIA0.SWIHIASER
fgDisplay.Row = 20
fgDisplay.Col = 0: fgDisplay = "SWIHIASSE    2A"
fgDisplay.Col = 1: fgDisplay = "SOUS SERVICE"
fgDisplay.Col = 2: fgDisplay = lYSWIHIA0.SWIHIASSE
fgDisplay.Row = 21
fgDisplay.Col = 0: fgDisplay = "SWIHIAUTI   10A"
fgDisplay.Col = 1: fgDisplay = "UTILISATEUR"
fgDisplay.Col = 2: fgDisplay = lYSWIHIA0.SWIHIAUTI
fgDisplay.Row = 22
fgDisplay.Col = 0: fgDisplay = "SWIHIANUM    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWIHIA0.SWIHIANUM
fgDisplay.Row = 23
fgDisplay.Col = 0: fgDisplay = "SWIHIAUT1   10A"
fgDisplay.Col = 1: fgDisplay = "UTILISA SAISIE"
fgDisplay.Col = 2: fgDisplay = lYSWIHIA0.SWIHIAUT1
fgDisplay.Row = 24
fgDisplay.Col = 0: fgDisplay = "SWIHIAPVA    1A"
fgDisplay.Col = 1: fgDisplay = "1ERE VALIDATION"
fgDisplay.Col = 2: fgDisplay = lYSWIHIA0.SWIHIAPVA
fgDisplay.Row = 25
fgDisplay.Col = 0: fgDisplay = "SWIHIAUT2   10A"
fgDisplay.Col = 1: fgDisplay = "UTILISA 1ER VALID"
fgDisplay.Col = 2: fgDisplay = lYSWIHIA0.SWIHIAUT2
End Sub
Public Sub srvYSWIHIB0_fgDisplay(lYSWIHIB0 As typeZSWIHIB0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 6
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIHIBETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWIHIB0.SWIHIBETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWIHIBNUM    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWIHIB0.SWIHIBNUM
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWIHIBNEN    1S"
fgDisplay.Col = 1: fgDisplay = "NUMERO ENVOI"
fgDisplay.Col = 2: fgDisplay = lYSWIHIB0.SWIHIBNEN
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWIHIBNLI    4P"
fgDisplay.Col = 1: fgDisplay = "NUMERO LIGNE"
fgDisplay.Col = 2: fgDisplay = lYSWIHIB0.SWIHIBNLI
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWIHIBDET   70A"
fgDisplay.Col = 1: fgDisplay = "DETAIL"
fgDisplay.Col = 2: fgDisplay = lYSWIHIB0.SWIHIBDET
End Sub
Public Sub srvYSWIHIC0_fgDisplay(lYSWIHIC0 As typeZSWIHIC0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 19
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIHICETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWIHIC0.SWIHICETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWIHICNUM    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWIHIC0.SWIHICNUM
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWIHICNEN    1S"
fgDisplay.Col = 1: fgDisplay = "NUMERO ENVOI"
fgDisplay.Col = 2: fgDisplay = lYSWIHIC0.SWIHICNEN
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWIHICNLI    4P"
fgDisplay.Col = 1: fgDisplay = "NUMERO LIGNE"
fgDisplay.Col = 2: fgDisplay = lYSWIHIC0.SWIHICNLI
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWIHICNSE   40A"
fgDisplay.Col = 1: fgDisplay = "NUMERO SEQUENCE"
fgDisplay.Col = 2: fgDisplay = lYSWIHIC0.SWIHICNSE
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWIHICSOC    4P"
fgDisplay.Col = 1: fgDisplay = "NUM OCC SEQUE"
fgDisplay.Col = 2: fgDisplay = lYSWIHIC0.SWIHICSOC
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "SWIHICNCH    2P"
fgDisplay.Col = 1: fgDisplay = "NUMERO CHAMP"
fgDisplay.Col = 2: fgDisplay = lYSWIHIC0.SWIHICNCH
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "SWIHICCOC    4P"
fgDisplay.Col = 1: fgDisplay = "NUM OCC CHAMP"
fgDisplay.Col = 2: fgDisplay = lYSWIHIC0.SWIHICCOC
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "SWIHICNLC    4P"
fgDisplay.Col = 1: fgDisplay = "NUMERO LIGNE CHAM"
fgDisplay.Col = 2: fgDisplay = lYSWIHIC0.SWIHICNLC
fgDisplay.Row = 10
fgDisplay.Col = 0: fgDisplay = "SWIHICSEQ    2A"
fgDisplay.Col = 1: fgDisplay = "DESCRIP SEQUENCE"
fgDisplay.Col = 2: fgDisplay = lYSWIHIC0.SWIHICSEQ
fgDisplay.Row = 11
fgDisplay.Col = 0: fgDisplay = "SWIHICCHA    2A"
fgDisplay.Col = 1: fgDisplay = "DESCRIP CHAMP"
fgDisplay.Col = 2: fgDisplay = lYSWIHIC0.SWIHICCHA
fgDisplay.Row = 12
fgDisplay.Col = 0: fgDisplay = "SWIHICILI    1A"
fgDisplay.Col = 1: fgDisplay = "INDICATEUR DEB"
fgDisplay.Col = 2: fgDisplay = lYSWIHIC0.SWIHICILI
fgDisplay.Row = 13
fgDisplay.Col = 0: fgDisplay = "SWIHICFAC    1A"
fgDisplay.Col = 1: fgDisplay = "FACULTATIF"
fgDisplay.Col = 2: fgDisplay = lYSWIHIC0.SWIHICFAC
fgDisplay.Row = 14
fgDisplay.Col = 0: fgDisplay = "SWIHICSIG    1A"
fgDisplay.Col = 1: fgDisplay = "SIGNE COMPLET"
fgDisplay.Col = 2: fgDisplay = lYSWIHIC0.SWIHICSIG
fgDisplay.Row = 15
fgDisplay.Col = 0: fgDisplay = "SWIHICSMA    4P"
fgDisplay.Col = 1: fgDisplay = "OCCUR SEQ MAXIMUM"
fgDisplay.Col = 2: fgDisplay = lYSWIHIC0.SWIHICSMA
fgDisplay.Row = 16
fgDisplay.Col = 0: fgDisplay = "SWIHICCMA    4P"
fgDisplay.Col = 1: fgDisplay = "OCCUR CHA MAXIMUM"
fgDisplay.Col = 2: fgDisplay = lYSWIHIC0.SWIHICCMA
fgDisplay.Row = 17
fgDisplay.Col = 0: fgDisplay = "SWIHICSMI    2P"
fgDisplay.Col = 1: fgDisplay = "OCCUR SEQ MINIMUM"
fgDisplay.Col = 2: fgDisplay = lYSWIHIC0.SWIHICSMI
fgDisplay.Row = 18
fgDisplay.Col = 0: fgDisplay = "SWIHICCMI    2P"
fgDisplay.Col = 1: fgDisplay = "OCCUR CHA MINIMUM"
fgDisplay.Col = 2: fgDisplay = lYSWIHIC0.SWIHICCMI
End Sub
Public Sub srvYSWIHIT0_fgDisplay(lYSWIHIT0 As typeZSWIHIT0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 17
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIHITETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWIHIT0.SWIHITETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWIHITNUM    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWIHIT0.SWIHITNUM
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWIHITNEN    1S"
fgDisplay.Col = 1: fgDisplay = "NUMERO ENVOI"
fgDisplay.Col = 2: fgDisplay = lYSWIHIT0.SWIHITNEN
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWIHITNSE   40A"
fgDisplay.Col = 1: fgDisplay = "NUMERO SEQUENCE"
fgDisplay.Col = 2: fgDisplay = lYSWIHIT0.SWIHITNSE
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWIHITSEQ    2A"
fgDisplay.Col = 1: fgDisplay = "SEQUENCE"
fgDisplay.Col = 2: fgDisplay = lYSWIHIT0.SWIHITSEQ
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWIHITOSE    4P"
fgDisplay.Col = 1: fgDisplay = "OCCURENCE SEQUE."
fgDisplay.Col = 2: fgDisplay = lYSWIHIT0.SWIHITOSE
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "SWIHITCHA    2P"
fgDisplay.Col = 1: fgDisplay = "CHAMP"
fgDisplay.Col = 2: fgDisplay = lYSWIHIT0.SWIHITCHA
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "SWIHITOCH    4P"
fgDisplay.Col = 1: fgDisplay = "OCCURENCE CHAMP"
fgDisplay.Col = 2: fgDisplay = lYSWIHIT0.SWIHITOCH
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "SWIHITIND    2A"
fgDisplay.Col = 1: fgDisplay = "INDICE"
fgDisplay.Col = 2: fgDisplay = lYSWIHIT0.SWIHITIND
fgDisplay.Row = 10
fgDisplay.Col = 0: fgDisplay = "SWIHITZON    2P"
fgDisplay.Col = 1: fgDisplay = "ZONE"
fgDisplay.Col = 2: fgDisplay = lYSWIHIT0.SWIHITZON
fgDisplay.Row = 11
fgDisplay.Col = 0: fgDisplay = "SWIHITOZO    4P"
fgDisplay.Col = 1: fgDisplay = "OCCURENCE ZONE"
fgDisplay.Col = 2: fgDisplay = lYSWIHIT0.SWIHITOZO
fgDisplay.Row = 12
fgDisplay.Col = 0: fgDisplay = "SWIHITSZO    2P"
fgDisplay.Col = 1: fgDisplay = "SOUS ZONE"
fgDisplay.Col = 2: fgDisplay = lYSWIHIT0.SWIHITSZO
fgDisplay.Row = 13
fgDisplay.Col = 0: fgDisplay = "SWIHITOSZ    4P"
fgDisplay.Col = 1: fgDisplay = "OCCURENCE S-ZONE"
fgDisplay.Col = 2: fgDisplay = lYSWIHIT0.SWIHITOSZ
fgDisplay.Row = 14
fgDisplay.Col = 0: fgDisplay = "SWIHITCON    3P"
fgDisplay.Col = 1: fgDisplay = "COMPTEUR ENREGIS"
fgDisplay.Col = 2: fgDisplay = lYSWIHIT0.SWIHITCON
fgDisplay.Row = 15
fgDisplay.Col = 0: fgDisplay = "SWIHITCOM    1A"
fgDisplay.Col = 1: fgDisplay = "COMPLET"
fgDisplay.Col = 2: fgDisplay = lYSWIHIT0.SWIHITCOM
fgDisplay.Row = 16
fgDisplay.Col = 0: fgDisplay = "SWIHITVAL   65A"
fgDisplay.Col = 1: fgDisplay = "VALEUR ZONE"
fgDisplay.Col = 2: fgDisplay = lYSWIHIT0.SWIHITVAL
End Sub
Public Sub srvYSWIJOB0_fgDisplay(lYSWIJOB0 As typeZSWIJOB0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 11
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIJOBETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWIJOB0.SWIJOBETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWIJOBPRO   10A"
fgDisplay.Col = 1: fgDisplay = "PROCEDURE"
fgDisplay.Col = 2: fgDisplay = lYSWIJOB0.SWIJOBPRO
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWIJOBJOB   10A"
fgDisplay.Col = 1: fgDisplay = "TRAVAIL"
fgDisplay.Col = 2: fgDisplay = lYSWIJOB0.SWIJOBJOB
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWIJOBUSR   10A"
fgDisplay.Col = 1: fgDisplay = "UTILISATEUR"
fgDisplay.Col = 2: fgDisplay = lYSWIJOB0.SWIJOBUSR
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWIJOBNBR    6A"
fgDisplay.Col = 1: fgDisplay = "N° TRAVAIL"
fgDisplay.Col = 2: fgDisplay = lYSWIJOB0.SWIJOBNBR
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWIJOBDLA    7P"
fgDisplay.Col = 1: fgDisplay = "DATE  LANCEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWIJOB0.SWIJOBDLA
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "SWIJOBHLA    6P"
fgDisplay.Col = 1: fgDisplay = "HEURE LANCEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWIJOB0.SWIJOBHLA
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "SWIJOBENV    4P"
fgDisplay.Col = 1: fgDisplay = "NBR MESS.A EVOYER"
fgDisplay.Col = 2: fgDisplay = lYSWIJOB0.SWIJOBENV
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "SWIJOBTER    6P"
fgDisplay.Col = 1: fgDisplay = "HEURE TERMINAISON"
fgDisplay.Col = 2: fgDisplay = lYSWIJOB0.SWIJOBTER
fgDisplay.Row = 10
fgDisplay.Col = 0: fgDisplay = "SWIJOBACT    1A"
fgDisplay.Col = 1: fgDisplay = "ACTIF O/N"
fgDisplay.Col = 2: fgDisplay = lYSWIJOB0.SWIJOBACT
End Sub
Public Sub srvYSWIMEA0_fgDisplay(lYSWIMEA0 As typeZSWIMEA0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 17
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIMEAETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWIMEA0.SWIMEAETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWIMEANUM    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWIMEA0.SWIMEANUM
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWIMEAREF   16A"
fgDisplay.Col = 1: fgDisplay = "REFERENCE"
fgDisplay.Col = 2: fgDisplay = lYSWIMEA0.SWIMEAREF
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWIMEAMES    3A"
fgDisplay.Col = 1: fgDisplay = "TYPE MESSAGE"
fgDisplay.Col = 2: fgDisplay = lYSWIMEA0.SWIMEAMES
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWIMEAEME   12A"
fgDisplay.Col = 1: fgDisplay = "EMETTEUR"
fgDisplay.Col = 2: fgDisplay = lYSWIMEA0.SWIMEAEME
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWIMEADRE    7P"
fgDisplay.Col = 1: fgDisplay = "DATE RECEPTION"
fgDisplay.Col = 2: fgDisplay = lYSWIMEA0.SWIMEADRE
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "SWIMEAHRE    7P"
fgDisplay.Col = 1: fgDisplay = "HEURE RECEPTION"
fgDisplay.Col = 2: fgDisplay = lYSWIMEA0.SWIMEAHRE
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "SWIMEAAGE    4B"
fgDisplay.Col = 1: fgDisplay = "AGENCE"
fgDisplay.Col = 2: fgDisplay = lYSWIMEA0.SWIMEAAGE
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "SWIMEASER    2A"
fgDisplay.Col = 1: fgDisplay = "SERVICE"
fgDisplay.Col = 2: fgDisplay = lYSWIMEA0.SWIMEASER
fgDisplay.Row = 10
fgDisplay.Col = 0: fgDisplay = "SWIMEASSE    2A"
fgDisplay.Col = 1: fgDisplay = "SOUS SERVICE"
fgDisplay.Col = 2: fgDisplay = lYSWIMEA0.SWIMEASSE
fgDisplay.Row = 11
fgDisplay.Col = 0: fgDisplay = "SWIMEAUTI   10A"
fgDisplay.Col = 1: fgDisplay = "UTILISATEUR"
fgDisplay.Col = 2: fgDisplay = lYSWIMEA0.SWIMEAUTI
fgDisplay.Row = 12
fgDisplay.Col = 0: fgDisplay = "SWIMEADTR    7P"
fgDisplay.Col = 1: fgDisplay = "DATE TRAITEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWIMEA0.SWIMEADTR
fgDisplay.Row = 13
fgDisplay.Col = 0: fgDisplay = "SWIMEAAVI    1A"
fgDisplay.Col = 1: fgDisplay = "AVIS EDITE O/N"
fgDisplay.Col = 2: fgDisplay = lYSWIMEA0.SWIMEAAVI
fgDisplay.Row = 14
fgDisplay.Col = 0: fgDisplay = "SWIMEADVA    7P"
fgDisplay.Col = 1: fgDisplay = "DATE VALEUR"
fgDisplay.Col = 2: fgDisplay = lYSWIMEA0.SWIMEADVA
fgDisplay.Row = 15
fgDisplay.Col = 0: fgDisplay = "SWIMEADEV    3A"
fgDisplay.Col = 1: fgDisplay = "DEVISE 1"
fgDisplay.Col = 2: fgDisplay = lYSWIMEA0.SWIMEADEV
fgDisplay.Row = 16
fgDisplay.Col = 0: fgDisplay = "SWIMEAMON 15.2P"
fgDisplay.Col = 1: fgDisplay = "MONTANT"
fgDisplay.Col = 2: fgDisplay = lYSWIMEA0.SWIMEAMON
End Sub
Public Sub srvYSWIMEB0_fgDisplay(lYSWIMEB0 As typeZSWIMEB0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 10
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIMEBETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWIMEB0.SWIMEBETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWIMEBNUM    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWIMEB0.SWIMEBNUM
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWIMEBNOR    3P"
fgDisplay.Col = 1: fgDisplay = "ORDRE"
fgDisplay.Col = 2: fgDisplay = lYSWIMEB0.SWIMEBNOR
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWIMEBCHA    2P"
fgDisplay.Col = 1: fgDisplay = "CHAMP"
fgDisplay.Col = 2: fgDisplay = lYSWIMEB0.SWIMEBCHA
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWIMEBIND    2A"
fgDisplay.Col = 1: fgDisplay = "INDICE"
fgDisplay.Col = 2: fgDisplay = lYSWIMEB0.SWIMEBIND
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWIMEBZON    2P"
fgDisplay.Col = 1: fgDisplay = "ZONE"
fgDisplay.Col = 2: fgDisplay = lYSWIMEB0.SWIMEBZON
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "SWIMEBSZO    2P"
fgDisplay.Col = 1: fgDisplay = "SOUS ZONE"
fgDisplay.Col = 2: fgDisplay = lYSWIMEB0.SWIMEBSZO
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "SWIENBINR    1A"
fgDisplay.Col = 1: fgDisplay = "INDICE REELL"
fgDisplay.Col = 2: fgDisplay = lYSWIMEB0.SWIENBINR
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "SWIMEBVAL   65A"
fgDisplay.Col = 1: fgDisplay = "VALEUR ZONE"
fgDisplay.Col = 2: fgDisplay = lYSWIMEB0.SWIMEBVAL
End Sub
Public Sub srvYSWIMEC0_fgDisplay(lYSWIMEC0 As typeZSWIMEC0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 5
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIMECETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWIMEC0.SWIMECETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWIMECNUM    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWIMEC0.SWIMECNUM
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWIMECNOR    3P"
fgDisplay.Col = 1: fgDisplay = "N° ORDRE DU CHAMP"
fgDisplay.Col = 2: fgDisplay = lYSWIMEC0.SWIMECNOR
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWIMECCOM   50A"
fgDisplay.Col = 1: fgDisplay = "COMMENTA. AJOUTEE"
fgDisplay.Col = 2: fgDisplay = lYSWIMEC0.SWIMECCOM
End Sub
Public Sub srvYSWIMEM0_fgDisplay(lYSWIMEM0 As typeZSWIMEM0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 6
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIMEMETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWIMEM0.SWIMEMETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWIMEMNUM    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWIMEM0.SWIMEMNUM
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWIMEMNOR    3P"
fgDisplay.Col = 1: fgDisplay = "N° ORDRE DU CHAMP"
fgDisplay.Col = 2: fgDisplay = lYSWIMEM0.SWIMEMNOR
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWIMEMMEG    3A"
fgDisplay.Col = 1: fgDisplay = "MESSAGE GENERE"
fgDisplay.Col = 2: fgDisplay = lYSWIMEM0.SWIMEMMEG
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWIMEMREG   16A"
fgDisplay.Col = 1: fgDisplay = "REFERENCE GENEREE"
fgDisplay.Col = 2: fgDisplay = lYSWIMEM0.SWIMEMREG
End Sub
Public Sub srvYSWIMEO0_fgDisplay(lYSWIMEO0 As typeZSWIMEO0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 7
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIMEOETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWIMEO0.SWIMEOETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWIMEONUM    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWIMEO0.SWIMEONUM
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWIMEONOR    3P"
fgDisplay.Col = 1: fgDisplay = "N° ORDRE DU CHAMP"
fgDisplay.Col = 2: fgDisplay = lYSWIMEO0.SWIMEONOR
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWIMEOOPR    6A"
fgDisplay.Col = 1: fgDisplay = "TYPE OPERATION"
fgDisplay.Col = 2: fgDisplay = lYSWIMEO0.SWIMEOOPR
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWIMEONAT    6A"
fgDisplay.Col = 1: fgDisplay = "NATURE OPERATION"
fgDisplay.Col = 2: fgDisplay = lYSWIMEO0.SWIMEONAT
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWIMEONOP    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO OPERATION"
fgDisplay.Col = 2: fgDisplay = lYSWIMEO0.SWIMEONOP
End Sub
Public Sub srvYSWIRAL0_fgDisplay(lYSWIRAL0 As typeZSWIRAL0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 4
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIRALDON  512A"
fgDisplay.Col = 1: fgDisplay = "DONNE MESSAGE"
fgDisplay.Col = 2: fgDisplay = lYSWIRAL0.SWIRALDON
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWIRALETA    4B"
fgDisplay.Col = 1: fgDisplay = ""
fgDisplay.Col = 2: fgDisplay = lYSWIRAL0.SWIRALETA
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWIRALMES    3A"
fgDisplay.Col = 1: fgDisplay = ""
fgDisplay.Col = 2: fgDisplay = lYSWIRAL0.SWIRALMES
End Sub
Public Sub srvYSWIRDE0_fgDisplay(lYSWIRDE0 As typeZSWIRDE0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 7
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIRDEETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWIRDE0.SWIRDEETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWIRDEBIC   12A"
fgDisplay.Col = 1: fgDisplay = "BIC"
fgDisplay.Col = 2: fgDisplay = lYSWIRDE0.SWIRDEBIC
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWIRDENUM    7P"
fgDisplay.Col = 1: fgDisplay = "REFERENCE CLIENT"
fgDisplay.Col = 2: fgDisplay = lYSWIRDE0.SWIRDENUM
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWIRDECOM   35A"
fgDisplay.Col = 1: fgDisplay = "COMPTE"
fgDisplay.Col = 2: fgDisplay = lYSWIRDE0.SWIRDECOM
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWIRDEDAT    7P"
fgDisplay.Col = 1: fgDisplay = "DATE DERNIERE EMI"
fgDisplay.Col = 2: fgDisplay = lYSWIRDE0.SWIRDEDAT
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWIRDECOU    5S"
fgDisplay.Col = 1: fgDisplay = "COMPTEUR QUOTIDIE"
fgDisplay.Col = 2: fgDisplay = lYSWIRDE0.SWIRDECOU
End Sub
Public Sub srvYSWIREC0_fgDisplay(lYSWIREC0 As typeZSWIREC0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 6
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIRECETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWIREC0.SWIRECETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWIRECNUM    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWIREC0.SWIRECNUM
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWIRECNLI    2P"
fgDisplay.Col = 1: fgDisplay = "NUMERO DE LIGNE"
fgDisplay.Col = 2: fgDisplay = lYSWIREC0.SWIRECNLI
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWIRECMES    3A"
fgDisplay.Col = 1: fgDisplay = "NUMERO DE MESSAGE"
fgDisplay.Col = 2: fgDisplay = lYSWIREC0.SWIRECMES
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWIRECDON  250A"
fgDisplay.Col = 1: fgDisplay = "DONNE MESSAGE"
fgDisplay.Col = 2: fgDisplay = lYSWIREC0.SWIRECDON
End Sub
Public Sub srvYSWIRED0_fgDisplay(lYSWIRED0 As typeZSWIRED0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 14
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIREDETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWIRED0.SWIREDETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWIREDAGE    4B"
fgDisplay.Col = 1: fgDisplay = "AGENCE"
fgDisplay.Col = 2: fgDisplay = lYSWIRED0.SWIREDAGE
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWIREDSER    2A"
fgDisplay.Col = 1: fgDisplay = "SERVICE"
fgDisplay.Col = 2: fgDisplay = lYSWIRED0.SWIREDSER
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWIREDSSE    2A"
fgDisplay.Col = 1: fgDisplay = "SOUS SERVICE"
fgDisplay.Col = 2: fgDisplay = lYSWIRED0.SWIREDSSE
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWIREDREF   16A"
fgDisplay.Col = 1: fgDisplay = "REFERENCE"
fgDisplay.Col = 2: fgDisplay = lYSWIRED0.SWIREDREF
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWIREDME1    3A"
fgDisplay.Col = 1: fgDisplay = "TYPE MESSAGE (1)"
fgDisplay.Col = 2: fgDisplay = lYSWIRED0.SWIREDME1
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "SWIREDME2    3A"
fgDisplay.Col = 1: fgDisplay = "TYPE MESSAGE (2)"
fgDisplay.Col = 2: fgDisplay = lYSWIRED0.SWIREDME2
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "SWIREDEM1   12A"
fgDisplay.Col = 1: fgDisplay = "EMETTEUR (1)"
fgDisplay.Col = 2: fgDisplay = lYSWIRED0.SWIREDEM1
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "SWIREDEM2   12A"
fgDisplay.Col = 1: fgDisplay = "EMETTEUR (2)"
fgDisplay.Col = 2: fgDisplay = lYSWIRED0.SWIREDEM2
fgDisplay.Row = 10
fgDisplay.Col = 0: fgDisplay = "SWIREDNU1    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE(1)"
fgDisplay.Col = 2: fgDisplay = lYSWIRED0.SWIREDNU1
fgDisplay.Row = 11
fgDisplay.Col = 0: fgDisplay = "SWIREDNU2    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE(2)"
fgDisplay.Col = 2: fgDisplay = lYSWIRED0.SWIREDNU2
fgDisplay.Row = 12
fgDisplay.Col = 0: fgDisplay = "SWIREDDAT    7P"
fgDisplay.Col = 1: fgDisplay = "DATE TRAITEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWIRED0.SWIREDDAT
fgDisplay.Row = 13
fgDisplay.Col = 0: fgDisplay = "SWIREDAVI    1A"
fgDisplay.Col = 1: fgDisplay = "EDIT OU NON"
fgDisplay.Col = 2: fgDisplay = lYSWIRED0.SWIREDAVI
End Sub
Public Sub srvYSWIRST0_fgDisplay(lYSWIRST0 As typeZSWIRST0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 4
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWIRSTDON   93A"
fgDisplay.Col = 1: fgDisplay = "DONNE MESSAGE"
fgDisplay.Col = 2: fgDisplay = lYSWIRST0.SWIRSTDON
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWIRSTETA    4B"
fgDisplay.Col = 1: fgDisplay = ""
fgDisplay.Col = 2: fgDisplay = lYSWIRST0.SWIRSTETA
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWIRSTMES    3A"
fgDisplay.Col = 1: fgDisplay = ""
fgDisplay.Col = 2: fgDisplay = lYSWIRST0.SWIRSTMES
End Sub
Public Sub srvYSWISCA0_fgDisplay(lYSWISCA0 As typeZSWISCA0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 26
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWISCAETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWISCA0.SWISCAETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWISCAREF   16A"
fgDisplay.Col = 1: fgDisplay = "REFERNECE"
fgDisplay.Col = 2: fgDisplay = lYSWISCA0.SWISCAREF
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWISCANEN    1A"
fgDisplay.Col = 1: fgDisplay = "NUMERO DE RENVOI"
fgDisplay.Col = 2: fgDisplay = lYSWISCA0.SWISCANEN
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWISCAPRI    2A"
fgDisplay.Col = 1: fgDisplay = "CODE PROIRITE"
fgDisplay.Col = 2: fgDisplay = lYSWISCA0.SWISCAPRI
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWISCAMES    3A"
fgDisplay.Col = 1: fgDisplay = "TYPE MESSAGE"
fgDisplay.Col = 2: fgDisplay = lYSWISCA0.SWISCAMES
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWISCADOR   12A"
fgDisplay.Col = 1: fgDisplay = "DONNEUR ORDRE"
fgDisplay.Col = 2: fgDisplay = lYSWISCA0.SWISCADOR
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "SWISCADES   12A"
fgDisplay.Col = 1: fgDisplay = "DESTINATAIRE"
fgDisplay.Col = 2: fgDisplay = lYSWISCA0.SWISCADES
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "SWISCADVA    7P"
fgDisplay.Col = 1: fgDisplay = "DATE VALEUR"
fgDisplay.Col = 2: fgDisplay = lYSWISCA0.SWISCADVA
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "SWISCADE1    3A"
fgDisplay.Col = 1: fgDisplay = "DEVISE 1"
fgDisplay.Col = 2: fgDisplay = lYSWISCA0.SWISCADE1
fgDisplay.Row = 10
fgDisplay.Col = 0: fgDisplay = "SWISCAMON 15.2P"
fgDisplay.Col = 1: fgDisplay = "MONTANT"
fgDisplay.Col = 2: fgDisplay = lYSWISCA0.SWISCAMON
fgDisplay.Row = 11
fgDisplay.Col = 0: fgDisplay = "SWISCADE2    3A"
fgDisplay.Col = 1: fgDisplay = "DEVISE 2"
fgDisplay.Col = 2: fgDisplay = lYSWISCA0.SWISCADE2
fgDisplay.Row = 12
fgDisplay.Col = 0: fgDisplay = "SWISCADEN    7P"
fgDisplay.Col = 1: fgDisplay = "DATE ENVOI"
fgDisplay.Col = 2: fgDisplay = lYSWISCA0.SWISCADEN
fgDisplay.Row = 13
fgDisplay.Col = 0: fgDisplay = "SWISCAHEN    7P"
fgDisplay.Col = 1: fgDisplay = "HEURE ENVOI"
fgDisplay.Col = 2: fgDisplay = lYSWISCA0.SWISCAHEN
fgDisplay.Row = 14
fgDisplay.Col = 0: fgDisplay = "SWISCACOM    1A"
fgDisplay.Col = 1: fgDisplay = "COMPLET"
fgDisplay.Col = 2: fgDisplay = lYSWISCA0.SWISCACOM
fgDisplay.Row = 15
fgDisplay.Col = 0: fgDisplay = "SWISCATES    1A"
fgDisplay.Col = 1: fgDisplay = "TEST OU REEL"
fgDisplay.Col = 2: fgDisplay = lYSWISCA0.SWISCATES
fgDisplay.Row = 16
fgDisplay.Col = 0: fgDisplay = "SWISCASUP    1A"
fgDisplay.Col = 1: fgDisplay = "SUPPRIME"
fgDisplay.Col = 2: fgDisplay = lYSWISCA0.SWISCASUP
fgDisplay.Row = 17
fgDisplay.Col = 0: fgDisplay = "SWISCAVAL    1A"
fgDisplay.Col = 1: fgDisplay = "TOP VALIDATION"
fgDisplay.Col = 2: fgDisplay = lYSWISCA0.SWISCAVAL
fgDisplay.Row = 18
fgDisplay.Col = 0: fgDisplay = "SWISCAAGE    4B"
fgDisplay.Col = 1: fgDisplay = "AGENCE"
fgDisplay.Col = 2: fgDisplay = lYSWISCA0.SWISCAAGE
fgDisplay.Row = 19
fgDisplay.Col = 0: fgDisplay = "SWISCASER    2A"
fgDisplay.Col = 1: fgDisplay = "SERVICE"
fgDisplay.Col = 2: fgDisplay = lYSWISCA0.SWISCASER
fgDisplay.Row = 20
fgDisplay.Col = 0: fgDisplay = "SWISCASSE    2A"
fgDisplay.Col = 1: fgDisplay = "SOUS SERVICE"
fgDisplay.Col = 2: fgDisplay = lYSWISCA0.SWISCASSE
fgDisplay.Row = 21
fgDisplay.Col = 0: fgDisplay = "SWISCAUTI   10A"
fgDisplay.Col = 1: fgDisplay = "UTILISATEUR"
fgDisplay.Col = 2: fgDisplay = lYSWISCA0.SWISCAUTI
fgDisplay.Row = 22
fgDisplay.Col = 0: fgDisplay = "SWISCANUM    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWISCA0.SWISCANUM
fgDisplay.Row = 23
fgDisplay.Col = 0: fgDisplay = "SWISCAUT1   10A"
fgDisplay.Col = 1: fgDisplay = "UTILISA SAISIE"
fgDisplay.Col = 2: fgDisplay = lYSWISCA0.SWISCAUT1
fgDisplay.Row = 24
fgDisplay.Col = 0: fgDisplay = "SWISCAPVA    1A"
fgDisplay.Col = 1: fgDisplay = "1ERE VALIDATION"
fgDisplay.Col = 2: fgDisplay = lYSWISCA0.SWISCAPVA
fgDisplay.Row = 25
fgDisplay.Col = 0: fgDisplay = "SWISCAUT2   10A"
fgDisplay.Col = 1: fgDisplay = "UTILISA 1ER VALID"
fgDisplay.Col = 2: fgDisplay = lYSWISCA0.SWISCAUT2
End Sub
Public Sub srvYSWISCB0_fgDisplay(lYSWISCB0 As typeZSWISCB0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 6
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWISCBETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWISCB0.SWISCBETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWISCBNUM    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWISCB0.SWISCBNUM
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWISCBNEN    1S"
fgDisplay.Col = 1: fgDisplay = "NUMERO ENVOI"
fgDisplay.Col = 2: fgDisplay = lYSWISCB0.SWISCBNEN
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWISCBNLI    4P"
fgDisplay.Col = 1: fgDisplay = "NUMERO LIGNE"
fgDisplay.Col = 2: fgDisplay = lYSWISCB0.SWISCBNLI
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWISCBDET   70A"
fgDisplay.Col = 1: fgDisplay = "DETAIL"
fgDisplay.Col = 2: fgDisplay = lYSWISCB0.SWISCBDET
End Sub
Public Sub srvYSWISCC0_fgDisplay(lYSWISCC0 As typeZSWISCC0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 19
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWISCCETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWISCC0.SWISCCETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWISCCNUM    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWISCC0.SWISCCNUM
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWISCCNEN    1S"
fgDisplay.Col = 1: fgDisplay = "NUMERO ENVOI"
fgDisplay.Col = 2: fgDisplay = lYSWISCC0.SWISCCNEN
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWISCCNLI    4P"
fgDisplay.Col = 1: fgDisplay = "NUMERO LIGNE"
fgDisplay.Col = 2: fgDisplay = lYSWISCC0.SWISCCNLI
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWISCCNSE   40A"
fgDisplay.Col = 1: fgDisplay = "NUMERO SEQUENCE"
fgDisplay.Col = 2: fgDisplay = lYSWISCC0.SWISCCNSE
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWISCCSOC    4P"
fgDisplay.Col = 1: fgDisplay = "NUM OCC SEQUE"
fgDisplay.Col = 2: fgDisplay = lYSWISCC0.SWISCCSOC
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "SWISCCNCH    2P"
fgDisplay.Col = 1: fgDisplay = "NUMERO CHAMP"
fgDisplay.Col = 2: fgDisplay = lYSWISCC0.SWISCCNCH
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "SWISCCCOC    4P"
fgDisplay.Col = 1: fgDisplay = "NUM OCC CHAMP"
fgDisplay.Col = 2: fgDisplay = lYSWISCC0.SWISCCCOC
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "SWISCCNLC    4P"
fgDisplay.Col = 1: fgDisplay = "NUMERO LIGNE CHAM"
fgDisplay.Col = 2: fgDisplay = lYSWISCC0.SWISCCNLC
fgDisplay.Row = 10
fgDisplay.Col = 0: fgDisplay = "SWISCCSEQ    2A"
fgDisplay.Col = 1: fgDisplay = "DESCRIP SEQUENCE"
fgDisplay.Col = 2: fgDisplay = lYSWISCC0.SWISCCSEQ
fgDisplay.Row = 11
fgDisplay.Col = 0: fgDisplay = "SWISCCCHA    2A"
fgDisplay.Col = 1: fgDisplay = "DESCRIP CHAMP"
fgDisplay.Col = 2: fgDisplay = lYSWISCC0.SWISCCCHA
fgDisplay.Row = 12
fgDisplay.Col = 0: fgDisplay = "SWISCCILI    1A"
fgDisplay.Col = 1: fgDisplay = "INDICATEUR DEB"
fgDisplay.Col = 2: fgDisplay = lYSWISCC0.SWISCCILI
fgDisplay.Row = 13
fgDisplay.Col = 0: fgDisplay = "SWISCCFAC    1A"
fgDisplay.Col = 1: fgDisplay = "FACULTATIF"
fgDisplay.Col = 2: fgDisplay = lYSWISCC0.SWISCCFAC
fgDisplay.Row = 14
fgDisplay.Col = 0: fgDisplay = "SWISCCSIG    1A"
fgDisplay.Col = 1: fgDisplay = "SIGNE COMPLET"
fgDisplay.Col = 2: fgDisplay = lYSWISCC0.SWISCCSIG
fgDisplay.Row = 15
fgDisplay.Col = 0: fgDisplay = "SWISCCSMA    4P"
fgDisplay.Col = 1: fgDisplay = "OCCUR SEQ MAXIMUM"
fgDisplay.Col = 2: fgDisplay = lYSWISCC0.SWISCCSMA
fgDisplay.Row = 16
fgDisplay.Col = 0: fgDisplay = "SWISCCCMA    4P"
fgDisplay.Col = 1: fgDisplay = "OCCUR CHA MAXIMUM"
fgDisplay.Col = 2: fgDisplay = lYSWISCC0.SWISCCCMA
fgDisplay.Row = 17
fgDisplay.Col = 0: fgDisplay = "SWISCCSMI    2P"
fgDisplay.Col = 1: fgDisplay = "OCCUR SEQ MINIMUM"
fgDisplay.Col = 2: fgDisplay = lYSWISCC0.SWISCCSMI
fgDisplay.Row = 18
fgDisplay.Col = 0: fgDisplay = "SWISCCCMI    2P"
fgDisplay.Col = 1: fgDisplay = "OCCUR CHA MINIMUM"
fgDisplay.Col = 2: fgDisplay = lYSWISCC0.SWISCCCMI
End Sub
Public Sub srvYSWISCT0_fgDisplay(lYSWISCT0 As typeZSWISCT0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 17
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWISCTETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWISCT0.SWISCTETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWISCTNUM    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWISCT0.SWISCTNUM
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWISCTNEN    1S"
fgDisplay.Col = 1: fgDisplay = "NUMERO ENVOI"
fgDisplay.Col = 2: fgDisplay = lYSWISCT0.SWISCTNEN
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWISCTNSE   40A"
fgDisplay.Col = 1: fgDisplay = "NUMERO SEQUENCE"
fgDisplay.Col = 2: fgDisplay = lYSWISCT0.SWISCTNSE
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWISCTSEQ    2A"
fgDisplay.Col = 1: fgDisplay = "SEQUENCE"
fgDisplay.Col = 2: fgDisplay = lYSWISCT0.SWISCTSEQ
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWISCTOSE    4P"
fgDisplay.Col = 1: fgDisplay = "OCCURENCE SEQUE."
fgDisplay.Col = 2: fgDisplay = lYSWISCT0.SWISCTOSE
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "SWISCTCHA    2P"
fgDisplay.Col = 1: fgDisplay = "CHAMP"
fgDisplay.Col = 2: fgDisplay = lYSWISCT0.SWISCTCHA
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "SWISCTOCH    4P"
fgDisplay.Col = 1: fgDisplay = "OCCURENCE CHAMP"
fgDisplay.Col = 2: fgDisplay = lYSWISCT0.SWISCTOCH
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "SWISCTIND    2A"
fgDisplay.Col = 1: fgDisplay = "INDICE"
fgDisplay.Col = 2: fgDisplay = lYSWISCT0.SWISCTIND
fgDisplay.Row = 10
fgDisplay.Col = 0: fgDisplay = "SWISCTZON    2P"
fgDisplay.Col = 1: fgDisplay = "ZONE"
fgDisplay.Col = 2: fgDisplay = lYSWISCT0.SWISCTZON
fgDisplay.Row = 11
fgDisplay.Col = 0: fgDisplay = "SWISCTOZO    4P"
fgDisplay.Col = 1: fgDisplay = "OCCURENCE ZONE"
fgDisplay.Col = 2: fgDisplay = lYSWISCT0.SWISCTOZO
fgDisplay.Row = 12
fgDisplay.Col = 0: fgDisplay = "SWISCTSZO    2P"
fgDisplay.Col = 1: fgDisplay = "SOUS ZONE"
fgDisplay.Col = 2: fgDisplay = lYSWISCT0.SWISCTSZO
fgDisplay.Row = 13
fgDisplay.Col = 0: fgDisplay = "SWISCTOSZ    4P"
fgDisplay.Col = 1: fgDisplay = "OCCURENCE S-ZONE"
fgDisplay.Col = 2: fgDisplay = lYSWISCT0.SWISCTOSZ
fgDisplay.Row = 14
fgDisplay.Col = 0: fgDisplay = "SWISCTCON    3P"
fgDisplay.Col = 1: fgDisplay = "COMPTEUR ENREGIS"
fgDisplay.Col = 2: fgDisplay = lYSWISCT0.SWISCTCON
fgDisplay.Row = 15
fgDisplay.Col = 0: fgDisplay = "SWISCTCOM    1A"
fgDisplay.Col = 1: fgDisplay = "COMPLET"
fgDisplay.Col = 2: fgDisplay = lYSWISCT0.SWISCTCOM
fgDisplay.Row = 16
fgDisplay.Col = 0: fgDisplay = "SWISCTVAL   65A"
fgDisplay.Col = 1: fgDisplay = "VALEUR ZONE"
fgDisplay.Col = 2: fgDisplay = lYSWISCT0.SWISCTVAL
End Sub
Public Sub srvYSWISHA0_fgDisplay(lYSWISHA0 As typeZSWISHA0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 26
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWISHAETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWISHA0.SWISHAETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWISHAREF   16A"
fgDisplay.Col = 1: fgDisplay = "REFERNECE"
fgDisplay.Col = 2: fgDisplay = lYSWISHA0.SWISHAREF
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWISHANEN    1A"
fgDisplay.Col = 1: fgDisplay = "NUMERO DE RENVOI"
fgDisplay.Col = 2: fgDisplay = lYSWISHA0.SWISHANEN
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWISHAPRI    2A"
fgDisplay.Col = 1: fgDisplay = "CODE PROIRITE"
fgDisplay.Col = 2: fgDisplay = lYSWISHA0.SWISHAPRI
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWISHAMES    3A"
fgDisplay.Col = 1: fgDisplay = "TYPE MESSAGE"
fgDisplay.Col = 2: fgDisplay = lYSWISHA0.SWISHAMES
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWISHADOR   12A"
fgDisplay.Col = 1: fgDisplay = "DONNEUR ORDRE"
fgDisplay.Col = 2: fgDisplay = lYSWISHA0.SWISHADOR
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "SWISHADES   12A"
fgDisplay.Col = 1: fgDisplay = "DESTINATAIRE"
fgDisplay.Col = 2: fgDisplay = lYSWISHA0.SWISHADES
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "SWISHADVA    7P"
fgDisplay.Col = 1: fgDisplay = "DATE VALEUR"
fgDisplay.Col = 2: fgDisplay = lYSWISHA0.SWISHADVA
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "SWISHADE1    3A"
fgDisplay.Col = 1: fgDisplay = "DEVISE 1"
fgDisplay.Col = 2: fgDisplay = lYSWISHA0.SWISHADE1
fgDisplay.Row = 10
fgDisplay.Col = 0: fgDisplay = "SWISHAMON 15.2P"
fgDisplay.Col = 1: fgDisplay = "MONTANT"
fgDisplay.Col = 2: fgDisplay = lYSWISHA0.SWISHAMON
fgDisplay.Row = 11
fgDisplay.Col = 0: fgDisplay = "SWISHADE2    3A"
fgDisplay.Col = 1: fgDisplay = "DEVISE 2"
fgDisplay.Col = 2: fgDisplay = lYSWISHA0.SWISHADE2
fgDisplay.Row = 12
fgDisplay.Col = 0: fgDisplay = "SWISHADEN    7P"
fgDisplay.Col = 1: fgDisplay = "DATE ENVOI"
fgDisplay.Col = 2: fgDisplay = lYSWISHA0.SWISHADEN
fgDisplay.Row = 13
fgDisplay.Col = 0: fgDisplay = "SWISHAHEN    7P"
fgDisplay.Col = 1: fgDisplay = "HEURE ENVOI"
fgDisplay.Col = 2: fgDisplay = lYSWISHA0.SWISHAHEN
fgDisplay.Row = 14
fgDisplay.Col = 0: fgDisplay = "SWISHACOM    1A"
fgDisplay.Col = 1: fgDisplay = "COMPLET"
fgDisplay.Col = 2: fgDisplay = lYSWISHA0.SWISHACOM
fgDisplay.Row = 15
fgDisplay.Col = 0: fgDisplay = "SWISHATES    1A"
fgDisplay.Col = 1: fgDisplay = "TEST OU REEL"
fgDisplay.Col = 2: fgDisplay = lYSWISHA0.SWISHATES
fgDisplay.Row = 16
fgDisplay.Col = 0: fgDisplay = "SWISHASUP    1A"
fgDisplay.Col = 1: fgDisplay = "SUPPRIME"
fgDisplay.Col = 2: fgDisplay = lYSWISHA0.SWISHASUP
fgDisplay.Row = 17
fgDisplay.Col = 0: fgDisplay = "SWISHAVAL    1A"
fgDisplay.Col = 1: fgDisplay = "TOP VALIDATION"
fgDisplay.Col = 2: fgDisplay = lYSWISHA0.SWISHAVAL
fgDisplay.Row = 18
fgDisplay.Col = 0: fgDisplay = "SWISHAAGE    4B"
fgDisplay.Col = 1: fgDisplay = "AGENCE"
fgDisplay.Col = 2: fgDisplay = lYSWISHA0.SWISHAAGE
fgDisplay.Row = 19
fgDisplay.Col = 0: fgDisplay = "SWISHASER    2A"
fgDisplay.Col = 1: fgDisplay = "SERVICE"
fgDisplay.Col = 2: fgDisplay = lYSWISHA0.SWISHASER
fgDisplay.Row = 20
fgDisplay.Col = 0: fgDisplay = "SWISHASSE    2A"
fgDisplay.Col = 1: fgDisplay = "SOUS SERVICE"
fgDisplay.Col = 2: fgDisplay = lYSWISHA0.SWISHASSE
fgDisplay.Row = 21
fgDisplay.Col = 0: fgDisplay = "SWISHAUTI   10A"
fgDisplay.Col = 1: fgDisplay = "UTILISATEUR"
fgDisplay.Col = 2: fgDisplay = lYSWISHA0.SWISHAUTI
fgDisplay.Row = 22
fgDisplay.Col = 0: fgDisplay = "SWISHANUM    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWISHA0.SWISHANUM
fgDisplay.Row = 23
fgDisplay.Col = 0: fgDisplay = "SWISHAUT1   10A"
fgDisplay.Col = 1: fgDisplay = "UTILISA SAISIE"
fgDisplay.Col = 2: fgDisplay = lYSWISHA0.SWISHAUT1
fgDisplay.Row = 24
fgDisplay.Col = 0: fgDisplay = "SWISHAPVA    1A"
fgDisplay.Col = 1: fgDisplay = "1ERE VALIDATION"
fgDisplay.Col = 2: fgDisplay = lYSWISHA0.SWISHAPVA
fgDisplay.Row = 25
fgDisplay.Col = 0: fgDisplay = "SWISHAUT2   10A"
fgDisplay.Col = 1: fgDisplay = "UTILISA 1ER VALID"
fgDisplay.Col = 2: fgDisplay = lYSWISHA0.SWISHAUT2
End Sub
Public Sub srvYSWISHB0_fgDisplay(lYSWISHB0 As typeZSWISHB0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 6
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWISHBETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWISHB0.SWISHBETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWISHBNUM    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWISHB0.SWISHBNUM
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWISHBNEN    1S"
fgDisplay.Col = 1: fgDisplay = "NUMERO ENVOI"
fgDisplay.Col = 2: fgDisplay = lYSWISHB0.SWISHBNEN
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWISHBNLI    4P"
fgDisplay.Col = 1: fgDisplay = "NUMERO LIGNE"
fgDisplay.Col = 2: fgDisplay = lYSWISHB0.SWISHBNLI
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWISHBDET   70A"
fgDisplay.Col = 1: fgDisplay = "DETAIL"
fgDisplay.Col = 2: fgDisplay = lYSWISHB0.SWISHBDET
End Sub
Public Sub srvYSWISHC0_fgDisplay(lYSWISHC0 As typeZSWISHC0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 19
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWISHCETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWISHC0.SWISHCETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWISHCNUM    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWISHC0.SWISHCNUM
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWISHCNEN    1S"
fgDisplay.Col = 1: fgDisplay = "NUMERO ENVOI"
fgDisplay.Col = 2: fgDisplay = lYSWISHC0.SWISHCNEN
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWISHCNLI    4P"
fgDisplay.Col = 1: fgDisplay = "NUMERO LIGNE"
fgDisplay.Col = 2: fgDisplay = lYSWISHC0.SWISHCNLI
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWISHCNSE   40A"
fgDisplay.Col = 1: fgDisplay = "NUMERO SEQUENCE"
fgDisplay.Col = 2: fgDisplay = lYSWISHC0.SWISHCNSE
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWISHCSOC    4P"
fgDisplay.Col = 1: fgDisplay = "NUM OCC SEQUE"
fgDisplay.Col = 2: fgDisplay = lYSWISHC0.SWISHCSOC
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "SWISHCNCH    2P"
fgDisplay.Col = 1: fgDisplay = "NUMERO CHAMP"
fgDisplay.Col = 2: fgDisplay = lYSWISHC0.SWISHCNCH
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "SWISHCCOC    4P"
fgDisplay.Col = 1: fgDisplay = "NUM OCC CHAMP"
fgDisplay.Col = 2: fgDisplay = lYSWISHC0.SWISHCCOC
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "SWISHCNLC    4P"
fgDisplay.Col = 1: fgDisplay = "NUMERO LIGNE CHAM"
fgDisplay.Col = 2: fgDisplay = lYSWISHC0.SWISHCNLC
fgDisplay.Row = 10
fgDisplay.Col = 0: fgDisplay = "SWISHCSEQ    2A"
fgDisplay.Col = 1: fgDisplay = "DESCRIP SEQUENCE"
fgDisplay.Col = 2: fgDisplay = lYSWISHC0.SWISHCSEQ
fgDisplay.Row = 11
fgDisplay.Col = 0: fgDisplay = "SWISHCCHA    2A"
fgDisplay.Col = 1: fgDisplay = "DESCRIP CHAMP"
fgDisplay.Col = 2: fgDisplay = lYSWISHC0.SWISHCCHA
fgDisplay.Row = 12
fgDisplay.Col = 0: fgDisplay = "SWISHCILI    1A"
fgDisplay.Col = 1: fgDisplay = "INDICATEUR DEB"
fgDisplay.Col = 2: fgDisplay = lYSWISHC0.SWISHCILI
fgDisplay.Row = 13
fgDisplay.Col = 0: fgDisplay = "SWISHCFAC    1A"
fgDisplay.Col = 1: fgDisplay = "FACULTATIF"
fgDisplay.Col = 2: fgDisplay = lYSWISHC0.SWISHCFAC
fgDisplay.Row = 14
fgDisplay.Col = 0: fgDisplay = "SWISHCSIG    1A"
fgDisplay.Col = 1: fgDisplay = "SIGNE COMPLET"
fgDisplay.Col = 2: fgDisplay = lYSWISHC0.SWISHCSIG
fgDisplay.Row = 15
fgDisplay.Col = 0: fgDisplay = "SWISHCSMA    4P"
fgDisplay.Col = 1: fgDisplay = "OCCUR SEQ MAXIMUM"
fgDisplay.Col = 2: fgDisplay = lYSWISHC0.SWISHCSMA
fgDisplay.Row = 16
fgDisplay.Col = 0: fgDisplay = "SWISHCCMA    4P"
fgDisplay.Col = 1: fgDisplay = "OCCUR CHA MAXIMUM"
fgDisplay.Col = 2: fgDisplay = lYSWISHC0.SWISHCCMA
fgDisplay.Row = 17
fgDisplay.Col = 0: fgDisplay = "SWISHCSMI    2P"
fgDisplay.Col = 1: fgDisplay = "OCCUR SEQ MINIMUM"
fgDisplay.Col = 2: fgDisplay = lYSWISHC0.SWISHCSMI
fgDisplay.Row = 18
fgDisplay.Col = 0: fgDisplay = "SWISHCCMI    2P"
fgDisplay.Col = 1: fgDisplay = "OCCUR CHA MINIMUM"
fgDisplay.Col = 2: fgDisplay = lYSWISHC0.SWISHCCMI
End Sub
Public Sub srvYSWISHT0_fgDisplay(lYSWISHT0 As typeZSWISHT0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 17
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWISHTETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWISHT0.SWISHTETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWISHTNUM    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWISHT0.SWISHTNUM
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWISHTNEN    1S"
fgDisplay.Col = 1: fgDisplay = "NUMERO ENVOI"
fgDisplay.Col = 2: fgDisplay = lYSWISHT0.SWISHTNEN
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWISHTNSE   40A"
fgDisplay.Col = 1: fgDisplay = "NUMERO SEQUENCE"
fgDisplay.Col = 2: fgDisplay = lYSWISHT0.SWISHTNSE
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWISHTSEQ    2A"
fgDisplay.Col = 1: fgDisplay = "SEQUENCE"
fgDisplay.Col = 2: fgDisplay = lYSWISHT0.SWISHTSEQ
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWISHTOSE    4P"
fgDisplay.Col = 1: fgDisplay = "OCCURENCE SEQUE."
fgDisplay.Col = 2: fgDisplay = lYSWISHT0.SWISHTOSE
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "SWISHTCHA    2P"
fgDisplay.Col = 1: fgDisplay = "CHAMP"
fgDisplay.Col = 2: fgDisplay = lYSWISHT0.SWISHTCHA
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "SWISHTOCH    4P"
fgDisplay.Col = 1: fgDisplay = "OCCURENCE CHAMP"
fgDisplay.Col = 2: fgDisplay = lYSWISHT0.SWISHTOCH
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "SWISHTIND    2A"
fgDisplay.Col = 1: fgDisplay = "INDICE"
fgDisplay.Col = 2: fgDisplay = lYSWISHT0.SWISHTIND
fgDisplay.Row = 10
fgDisplay.Col = 0: fgDisplay = "SWISHTZON    2P"
fgDisplay.Col = 1: fgDisplay = "ZONE"
fgDisplay.Col = 2: fgDisplay = lYSWISHT0.SWISHTZON
fgDisplay.Row = 11
fgDisplay.Col = 0: fgDisplay = "SWISHTOZO    4P"
fgDisplay.Col = 1: fgDisplay = "OCCURENCE ZONE"
fgDisplay.Col = 2: fgDisplay = lYSWISHT0.SWISHTOZO
fgDisplay.Row = 12
fgDisplay.Col = 0: fgDisplay = "SWISHTSZO    2P"
fgDisplay.Col = 1: fgDisplay = "SOUS ZONE"
fgDisplay.Col = 2: fgDisplay = lYSWISHT0.SWISHTSZO
fgDisplay.Row = 13
fgDisplay.Col = 0: fgDisplay = "SWISHTOSZ    4P"
fgDisplay.Col = 1: fgDisplay = "OCCURENCE S-ZONE"
fgDisplay.Col = 2: fgDisplay = lYSWISHT0.SWISHTOSZ
fgDisplay.Row = 14
fgDisplay.Col = 0: fgDisplay = "SWISHTCON    3P"
fgDisplay.Col = 1: fgDisplay = "COMPTEUR ENREGIS"
fgDisplay.Col = 2: fgDisplay = lYSWISHT0.SWISHTCON
fgDisplay.Row = 15
fgDisplay.Col = 0: fgDisplay = "SWISHTCOM    1A"
fgDisplay.Col = 1: fgDisplay = "COMPLET"
fgDisplay.Col = 2: fgDisplay = lYSWISHT0.SWISHTCOM
fgDisplay.Row = 16
fgDisplay.Col = 0: fgDisplay = "SWISHTVAL   65A"
fgDisplay.Col = 1: fgDisplay = "VALEUR ZONE"
fgDisplay.Col = 2: fgDisplay = lYSWISHT0.SWISHTVAL
End Sub
Public Sub srvYSWISRC0_fgDisplay(lYSWISRC0 As typeZSWISRC0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 4
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "AIDCLICOD  6.2S"
fgDisplay.Col = 1: fgDisplay = "111"
fgDisplay.Col = 2: fgDisplay = lYSWISRC0.AIDCLICOD
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "AIDCLIPRG    6S"
fgDisplay.Col = 1: fgDisplay = "222"
fgDisplay.Col = 2: fgDisplay = lYSWISRC0.AIDCLIPRG
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "AIDCLIFMT   80A"
fgDisplay.Col = 1: fgDisplay = "333"
fgDisplay.Col = 2: fgDisplay = lYSWISRC0.AIDCLIFMT
End Sub
Public Sub srvYSWITAB0_fgDisplay(lYSWITAB0 As typeZSWITAB0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 7
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWITABETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWITAB0.SWITABETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWITABNUM    2P"
fgDisplay.Col = 1: fgDisplay = "NUMERO TABLE"
fgDisplay.Col = 2: fgDisplay = lYSWITAB0.SWITABNUM
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWITABARG   25A"
fgDisplay.Col = 1: fgDisplay = "ARGUMENT"
fgDisplay.Col = 2: fgDisplay = lYSWITAB0.SWITABARG
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWITABLO1   12A"
fgDisplay.Col = 1: fgDisplay = "LOGIQUE 1"
fgDisplay.Col = 2: fgDisplay = lYSWITAB0.SWITABLO1
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWITABLO2   12A"
fgDisplay.Col = 1: fgDisplay = "LOGIQUE 2"
fgDisplay.Col = 2: fgDisplay = lYSWITAB0.SWITABLO2
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWITABDON  300A"
fgDisplay.Col = 1: fgDisplay = "DONNEES"
fgDisplay.Col = 2: fgDisplay = lYSWITAB0.SWITABDON
End Sub
Public Sub srvYSWITEM0_fgDisplay(lYSWITEM0 As typeZSWITEM0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 17
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "SWITEMETA    4B"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = lYSWITEM0.SWITEMETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "SWITEMNUM    8P"
fgDisplay.Col = 1: fgDisplay = "NUMERO INTERNE"
fgDisplay.Col = 2: fgDisplay = lYSWITEM0.SWITEMNUM
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "SWITEMNEN    1S"
fgDisplay.Col = 1: fgDisplay = "NUMERO ENVOI"
fgDisplay.Col = 2: fgDisplay = lYSWITEM0.SWITEMNEN
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "SWITEMNSE   40A"
fgDisplay.Col = 1: fgDisplay = "NUMERO SEQUENCE"
fgDisplay.Col = 2: fgDisplay = lYSWITEM0.SWITEMNSE
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "SWITEMSEQ    2A"
fgDisplay.Col = 1: fgDisplay = "SEQUENCE"
fgDisplay.Col = 2: fgDisplay = lYSWITEM0.SWITEMSEQ
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "SWITEMOSE    4P"
fgDisplay.Col = 1: fgDisplay = "OCCURENCE SEQUE."
fgDisplay.Col = 2: fgDisplay = lYSWITEM0.SWITEMOSE
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "SWITEMCHA    2P"
fgDisplay.Col = 1: fgDisplay = "CHAMP"
fgDisplay.Col = 2: fgDisplay = lYSWITEM0.SWITEMCHA
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "SWITEMOCH    4P"
fgDisplay.Col = 1: fgDisplay = "OCCURENCE CHAMP"
fgDisplay.Col = 2: fgDisplay = lYSWITEM0.SWITEMOCH
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "SWITEMIND    2A"
fgDisplay.Col = 1: fgDisplay = "INDICE"
fgDisplay.Col = 2: fgDisplay = lYSWITEM0.SWITEMIND
fgDisplay.Row = 10
fgDisplay.Col = 0: fgDisplay = "SWITEMZON    2P"
fgDisplay.Col = 1: fgDisplay = "ZONE"
fgDisplay.Col = 2: fgDisplay = lYSWITEM0.SWITEMZON
fgDisplay.Row = 11
fgDisplay.Col = 0: fgDisplay = "SWITEMOZO    4P"
fgDisplay.Col = 1: fgDisplay = "OCCURENCE ZONE"
fgDisplay.Col = 2: fgDisplay = lYSWITEM0.SWITEMOZO
fgDisplay.Row = 12
fgDisplay.Col = 0: fgDisplay = "SWITEMSZO    2P"
fgDisplay.Col = 1: fgDisplay = "SOUS ZONE"
fgDisplay.Col = 2: fgDisplay = lYSWITEM0.SWITEMSZO
fgDisplay.Row = 13
fgDisplay.Col = 0: fgDisplay = "SWITEMOSZ    4P"
fgDisplay.Col = 1: fgDisplay = "OCCURENCE S-ZONE"
fgDisplay.Col = 2: fgDisplay = lYSWITEM0.SWITEMOSZ
fgDisplay.Row = 14
fgDisplay.Col = 0: fgDisplay = "SWITEMCON    3P"
fgDisplay.Col = 1: fgDisplay = "COMPTEUR ENREGIS"
fgDisplay.Col = 2: fgDisplay = lYSWITEM0.SWITEMCON
fgDisplay.Row = 15
fgDisplay.Col = 0: fgDisplay = "SWITEMCOM    1A"
fgDisplay.Col = 1: fgDisplay = "COMPLET"
fgDisplay.Col = 2: fgDisplay = lYSWITEM0.SWITEMCOM
fgDisplay.Row = 16
fgDisplay.Col = 0: fgDisplay = "SWITEMVAL   65A"
fgDisplay.Col = 1: fgDisplay = "VALEUR ZONE"
fgDisplay.Col = 2: fgDisplay = lYSWITEM0.SWITEMVAL
End Sub
