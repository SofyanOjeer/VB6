Attribute VB_Name = "rsZFCIGCO0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeZFCIGCO0
    FCIGCOETA       As Long                           ' ETABLISSEMENT
'Séparateur     FCIGCOS01       As String * 1                     ' SEPERATEUR
    FCIGCOCLI       As String * 7                     '  RESPONSABLE
'Séparateur     FCIGCOS02       As String * 1                     ' SEPERATEUR
    FCIGCOPLA       As Long                           ' NUMERO PLAN
'Séparateur     FCIGCOS03       As String * 1                     ' SEPERATEUR
    FCIGCOCPT       As String * 20                    ' NUMERO COMPTE
'Séparateur     FCIGCOS04       As String * 1                     ' SEPERATEUR
    FCIGCONUC       As Long                           ' NUMERO CHEQUE
'Séparateur     FCIGCOS05       As String * 1                     ' SEPERATEUR
    FCIGCOCAR       As String * 16                    ' NUMERO CARTE
'Séparateur     FCIGCOS06       As String * 1                     ' SEPERATEUR
     FCIGCOSES       As Long                           ' NUM¢ SEQUENCE STATUT
'Séparateur     FCIGCOS07       As String * 1                     ' SEPERATEUR
     FCIGCOSEA       As Long                           ' NUM¢ SEQUENCE ACTION
'Séparateur     FCIGCOS08       As String * 1                     ' SEPERATEUR
    FCIGCODLI       As String * 10                    ' DATE LIMITE RETENT¢
'Séparateur     FCIGCOS09       As String * 1                     ' SEPERATEUR
    FCIGCODAJ       As String * 10                    ' DATE JOUR
'Séparateur     FCIGCOS10       As String * 1                     ' SEPERATEUR
    FCIGCOCOU       As String * 6                     ' CODE COURRIER TRANSM
'Séparateur     FCIGCOS11       As String * 1                     ' SEPERATEUR
    FCIGCOLIB       As String * 30                    ' LIBELLE COURRIER
'Séparateur     FCIGCOS12       As String * 1                     ' SEPERATEUR
    FCIGCOTYC       As String * 1                     ' TYPE COURRIER TRANSM
'Séparateur     FCIGCOS13       As String * 1                     ' SEPERATEUR
    FCIGCOLTY       As String * 30                    ' LIBELLE TYPE COURR.
'Séparateur     FCIGCOS14       As String * 1                     ' SEPERATEUR
    FCIGCOENV       As String * 1                     ' ENVOI RECOMANDE
'Séparateur     FCIGCOS15       As String * 1                     ' SEPERATEUR
    FCIGCOREC       As String * 30                    ' LIBELLE RECOMMANDE
'Séparateur     FCIGCOS16       As String * 1                     ' SEPERATEUR
    FCIGCODCP       As String * 10                    ' DATE COURRIER PRECED
'Séparateur     FCIGCOS17       As String * 1                     ' SEPERATEUR
    FCIGCOEDI       As String * 10                    ' DATE EDITION
'Séparateur     FCIGCOS18       As String * 1                     ' SEPERATEUR
    FCIGCORED       As String * 1                     ' REEDITION (O/N)
'Séparateur     FCIGCOS19       As String * 1                     ' SEPERATEUR
    FCIGCONDE       As String * 7                     ' NUMERO CLIENT DESTIN
'Séparateur     FCIGCOS20       As String * 1                     ' SEPERATEUR
    FCIGCOLTD       As String * 30                    ' LIBELLE ETAT DESTINA
'Séparateur     FCIGCOS21       As String * 1                     ' SEPERATEUR
    FCIGCONRD       As String * 32                    ' NOM/RAISON DESTINATA
'Séparateur     FCIGCOS22       As String * 1                     ' SEPERATEUR
    FCIGCOPRD       As String * 32                    ' PRENOM/RAISON DESTIN
'Séparateur     FCIGCOS23       As String * 1                     ' SEPERATEUR
    FCIGCOA1D       As String * 32                    ' ADRESSE 1 DESTINATAT
'Séparateur     FCIGCOS24       As String * 1                     ' SEPERATEUR
    FCIGCOA2D       As String * 32                    ' ADRESSE 2 DESTINATAT
'Séparateur     FCIGCOS25       As String * 1                     ' SEPERATEUR
    FCIGCOA3D       As String * 32                    ' ADRESSE 3 DESTINATAT
'Séparateur     FCIGCOS26       As String * 1                     ' SEPERATEUR
    FCIGCOCPD       As String * 6                     ' CODE POSTAL DESTINAT
'Séparateur     FCIGCOS27       As String * 1                     ' SEPERATEUR
    FCIGCOVID       As String * 25                    ' VILLE DESTINATAIRE
'Séparateur     FCIGCOS28       As String * 1                     ' SEPERATEUR
    FCIGCOLPD       As String * 25                    ' LIBELL.PAYS DESTINAT
'Séparateur     FCIGCOS29       As String * 1                     ' SEPERATEUR
    FCIGCOCLD       As String * 7                     ' NUMERO CLIENT
'Séparateur     FCIGCOS30       As String * 1                     ' SEPERATEUR
    FCIGCOLTC       As String * 30                    ' LIBELLE ETAT CLIENT
'Séparateur     FCIGCOS31       As String * 1                     ' SEPERATEUR
    FCIGCONRC       As String * 32                    ' NOM/RAISON CLIENT
'Séparateur     FCIGCOS32       As String * 1                     ' SEPERATEUR
    FCIGCOPRC       As String * 32                    ' PRENOM/RAISON CLIENT
'Séparateur     FCIGCOS33       As String * 1                     ' SEPERATEUR
    FCIGCOAD1       As String * 32                    ' ADRESSE 1 CLIENT
'Séparateur     FCIGCOS34       As String * 1                     ' SEPERATEUR
    FCIGCOAD2       As String * 32                    ' ADRESSE 2 CLIENT
'Séparateur     FCIGCOS35       As String * 1                     ' SEPERATEUR
    FCIGCOAD3       As String * 32                    ' ADRESSE 3 CLIENT
'Séparateur     FCIGCOS36       As String * 1                     ' SEPERATEUR
    FCIGCOCPC       As String * 6                     ' CODE POSTAL CLIENT
'Séparateur     FCIGCOS37       As String * 1                     ' SEPERATEUR
    FCIGCOVIC       As String * 25                    ' VILLE CLIENT
'Séparateur     FCIGCOS38       As String * 1                     ' SEPERATEUR
    FCIGCOLPC       As String * 25                    ' LIBELL.PAYS CLIENT
'Séparateur     FCIGCOS39       As String * 1                     ' SEPERATEUR
    FCIGCOCLB       As String * 7                     ' NUMERO CLIENT BENEFI
'Séparateur     FCIGCOS40       As String * 1                     ' SEPERATEUR
    FCIGCOLTB       As String * 30                    ' LIBEL ETAT CLI BENEF
'Séparateur     FCIGCOS41       As String * 1                     ' SEPERATEUR
    FCIGCOBNR       As String * 32                    ' NOM/RAIS CLI BENENEF
'Séparateur     FCIGCOS42       As String * 1                     ' SEPERATEUR
    FCIGCOBPR       As String * 32                    ' PRENOM/RAIS CLI BENE
'Séparateur     FCIGCOS43       As String * 1                     ' SEPERATEUR
    FCIGCOA1B       As String * 32                    ' ADRESSE 1 BENEFICIAC
'Séparateur     FCIGCOS44       As String * 1                     ' SEPERATEUR
    FCIGCOA2B       As String * 32                    ' ADRESSE 2 BENEFICIAC
'Séparateur     FCIGCOS45       As String * 1                     ' SEPERATEUR
    FCIGCOA3B       As String * 32                    ' ADRESSE 3 BENEFICIAC
'Séparateur     FCIGCOS46       As String * 1                     ' SEPERATEUR
    FCIGCOCPB       As String * 6                     ' CODE POSTAL BENEFICI
'Séparateur     FCIGCOS47       As String * 1                     ' SEPERATEUR
    FCIGCOBVI       As String * 25                    ' VILLE BENEFICIAIRE
'Séparateur     FCIGCOS48       As String * 1                     ' SEPERATEUR
    FCIGCOLPB       As String * 25                    ' LIBELL.PAYS BENEFIC
'Séparateur     FCIGCOS49       As String * 1                     ' SEPERATEUR
    FCIGCOCLP       As String * 7                     ' N¢ CLIENT PORTEUR
'Séparateur     FCIGCOS50       As String * 1                     ' SEPERATEUR
    FCIGCOLTP       As String * 30                    ' LIBELLE ETAT PORTEUR
'Séparateur     FCIGCOS51       As String * 1                     ' SEPERATEUR
    FCIGCONRP       As String * 32                    ' NOM/RAISON PORTEUR
'Séparateur     FCIGCOS52       As String * 1                     ' SEPERATEUR
    FCIGCOPRP       As String * 32                    ' PRENOM/RAISON PORTEU
'Séparateur     FCIGCOS53       As String * 1                     ' SEPERATEUR
     FCIGCO1DP       As String * 32                    ' ADRESSE 1 PORTEUR
'Séparateur     FCIGCOS54       As String * 1                     ' SEPERATEUR
    FCIGCO2DP       As String * 32                    ' ADRESSE 2 PORTEUR
'Séparateur     FCIGCOS55       As String * 1                     ' SEPERATEUR
    FCIGCO3DP       As String * 32                    ' ADRESSE 3 PORTEUR
'Séparateur     FCIGCOS56       As String * 1                     ' SEPERATEUR
    FCIGCOCPP       As String * 6                     ' CODE POSTAL PORTEUR
'Séparateur     FCIGCOS57       As String * 1                     ' SEPERATEUR
    FCIGCOVIP       As String * 25                    ' VILLE PORTEUR
'Séparateur     FCIGCOS58       As String * 1                     ' SEPERATEUR
    FCIGCOLPP       As String * 25                    ' LIBELL.PAYS PORTEUR
'Séparateur     FCIGCOS59       As String * 1                     ' SEPERATEUR
    FCIGCOCLT       As String * 7                     ' N¢ CLIENT TITULAIRE
'Séparateur     FCIGCOS60       As String * 1                     ' SEPERATEUR
    FCIGCOLIT       As String * 30                    ' LIBELLE ETAT TITULAI
'Séparateur     FCIGCOS61       As String * 1                     ' SEPERATEUR
    FCIGCONOT       As String * 32                    ' NOM/RAISON TITULALAI
'Séparateur     FCIGCOS62       As String * 1                     ' SEPERATEUR
    FCIGCOPRT       As String * 32                    ' PRENOM/RAISON TITULA
'Séparateur     FCIGCOS63       As String * 1                     ' SEPERATEUR
     FCIGCO1DT       As String * 32                    ' ADRESSE 1 TITULAIRE
'Séparateur     FCIGCOS64       As String * 1                     ' SEPERATEUR
    FCIGCO2DT       As String * 32                    ' ADRESSE 2 TITULAIRE
'Séparateur     FCIGCOS65       As String * 1                     ' SEPERATEUR
    FCIGCO3DT       As String * 32                    ' ADRESSE 3 TITULAIRE
'Séparateur     FCIGCOS66       As String * 1                     ' SEPERATEUR
    FCIGCOPOT       As String * 6                     ' CODE POSTAL TITULAIR
'Séparateur     FCIGCOS67       As String * 1                     ' SEPERATEUR
    FCIGCOVIT       As String * 25                    ' VILLE TITULAIRE
'Séparateur     FCIGCOS68       As String * 1                     ' SEPERATEUR
    FCIGCOLPT       As String * 25                    ' LIBELL.PAYS TITULAIR
'Séparateur     FCIGCOS69       As String * 1                     ' SEPERATEUR
    FCIGCOCLC       As String * 7                     ' N¢ CLIENT COTITULAIR
'Séparateur     FCIGCOS70       As String * 1                     ' SEPERATEUR
    FCIGCOLIC       As String * 30                    ' LIBELLE ETAT COTITUL
'Séparateur     FCIGCOS71       As String * 1                     ' SEPERATEUR
    FCIGCONOC       As String * 32                    ' NOM/RAISON COTITULAL
'Séparateur     FCIGCOS72       As String * 1                     ' SEPERATEUR
    FCIGCOPCO       As String * 32                    ' PRENOM/RAISON COTITU
'Séparateur     FCIGCOS73       As String * 1                     ' SEPERATEUR
     FCIGCO1DC       As String * 32                    ' ADRESSE 1 COTITULAIR
'Séparateur     FCIGCOS74       As String * 1                     ' SEPERATEUR
    FCIGCO2DC       As String * 32                    ' ADRESSE 2 COTITULAIR
'Séparateur     FCIGCOS75       As String * 1                     ' SEPERATEUR
    FCIGCO3DC       As String * 32                    ' ADRESSE 3 COTITULAIR
'Séparateur     FCIGCOS76       As String * 1                     ' SEPERATEUR
    FCIGCOPOC       As String * 6                     ' CODE POSTAL COTITULA
'Séparateur     FCIGCOS77       As String * 1                     ' SEPERATEUR
    FCIGCOVLC       As String * 25                    ' VILLE COTITULAIRE
'Séparateur     FCIGCOS78       As String * 1                     ' SEPERATEUR
    FCIGCOPAC       As String * 25                    ' LIBELL.PAYS COTITULA
'Séparateur     FCIGCOS79       As String * 1                     ' SEPERATEUR
    FCIGCOCLM       As String * 7                     ' N¢ CLIENT MANDATAIRE
'Séparateur     FCIGCOS80       As String * 1                     ' SEPERATEUR
    FCIGCOLIM       As String * 30                    ' LIBELLE ETAT MANDATA
'Séparateur     FCIGCOS81       As String * 1                     ' SEPERATEUR
    FCIGCONOM       As String * 32                    ' NOM/RAISON MANDATAIR
'Séparateur     FCIGCOS82       As String * 1                     ' SEPERATEUR
    FCIGCOPRM       As String * 32                    ' PRENOM/RAISON MANDAT
'Séparateur     FCIGCOS83       As String * 1                     ' SEPERATEUR
     FCIGCO1DM       As String * 32                    ' ADRESSE 1 MANDATAIRE
'Séparateur     FCIGCOS84       As String * 1                     ' SEPERATEUR
    FCIGCO2DM       As String * 32                    ' ADRESSE 2 MANDATAIRE
'Séparateur     FCIGCOS85       As String * 1                     ' SEPERATEUR
    FCIGCO3DM       As String * 32                    ' ADRESSE 3 MANDATAIRE
'Séparateur     FCIGCOS86       As String * 1                     ' SEPERATEUR
    FCIGCOCPM       As String * 6                     ' CODE POSTAL MANDATAI
'Séparateur     FCIGCOS87       As String * 1                     ' SEPERATEUR
    FCIGCOVIM       As String * 25                    ' VILLE MANDATAIRE
'Séparateur     FCIGCOS88       As String * 1                     ' SEPERATEUR
    FCIGCOLPM       As String * 25                    ' LIBELL.PAYS MANDATAI
'Séparateur     FCIGCOS89       As String * 1                     ' SEPERATEUR
    FCIGCOCLG       As String * 7                     ' N¢ CLIENT GREFFE TRI
'Séparateur     FCIGCOS90       As String * 1                     ' SEPERATEUR
    FCIGCOLIG       As String * 30                    ' LIBELLE ETAT GREFFE
'Séparateur     FCIGCOS91       As String * 1                     ' SEPERATEUR
    FCIGCONOG       As String * 32                    ' NOM/RAISON GREFFE
'Séparateur     FCIGCOS92       As String * 1                     ' SEPERATEUR
    FCIGCOPRG       As String * 32                    ' PRENOM/RAISON GREFFE
'Séparateur     FCIGCOS93       As String * 1                     ' SEPERATEUR
     FCIGCO1DG       As String * 32                    ' ADRESSE 1 GREFFE
'Séparateur     FCIGCOS94       As String * 1                     ' SEPERATEUR
    FCIGCO2DG       As String * 32                    ' ADRESSE 2 GREFFE
'Séparateur     FCIGCOS95       As String * 1                     ' SEPERATEUR
    FCIGCO3DG       As String * 32                    ' ADRESSE 3 GREFFE
'Séparateur     FCIGCOS96       As String * 1                     ' SEPERATEUR
    FCIGCOCPG       As String * 6                     ' CODE POSTAL GREFFE
'Séparateur     FCIGCOS97       As String * 1                     ' SEPERATEUR
    FCIGCOVIG       As String * 25                    ' VILLE GREFFE
'Séparateur     FCIGCOS98       As String * 1                     ' SEPERATEUR
    FCIGCOLPG       As String * 25                    ' LIBELL.PAYS GREFFE
'Séparateur     FCIGCOS99       As String * 1                     ' SEPERATEUR
    FCIGCOLED       As String * 30                    ' LIEU EDITION
'Séparateur     FCIGCO100       As String * 1                     ' SEPERATEUR
    FCIGCOGES       As String * 32                    ' NOM GESTIONNAIRE
'Séparateur     FCIGCO101       As String * 1                     ' SEPERATEUR
    FCIGCOREL       As String * 32                    ' REFERENCE LIBRE
'Séparateur     FCIGCO102       As String * 1                     ' SEPERATEUR
    FCIGCOTEL       As String * 20                    ' TELEPHONE GESTIONNAI
'Séparateur     FCIGCO103       As String * 1                     ' SEPERATEUR
    FCIGCOREJ       As String * 6                     ' CODE REJET
'Séparateur     FCIGCO104       As String * 1                     ' SEPERATEUR
    FCIGCOLIR       As String * 30                    ' LIBELLE REJET
'Séparateur     FCIGCO105       As String * 1                     ' SEPERATEUR
    FCIGCOMCH       As String * 20                    ' REJETE
'Séparateur     FCIGCO106       As String * 1                     ' SEPERATEUR
    FCIGCODEV       As String * 3                     ' DEVISE MONTANT
'Séparateur     FCIGCO107       As String * 1                     ' SEPERATEUR
    FCIGCOAT1       As String * 1                     ' CAS N¢1 ATTESTION
'Séparateur     FCIGCO108       As String * 1                     ' SEPERATEUR
    FCIGCOAT2       As String * 1                     ' CAS N¢2 ATTESTION
'Séparateur     FCIGCO109       As String * 1                     ' SEPERATEUR
    FCIGCOAT3       As String * 1                     ' CAS N¢3 ATTESTION
'Séparateur     FCIGCO110       As String * 1                     ' SEPERATEUR
    FCIGCOAT4       As String * 1                     ' NON UTILISE PREVISIONN
'Séparateur     FCIGCO111       As String * 1                     ' SEPERATEUR
    FCIGCOIJ1       As String * 1                     ' CAS N¢1 INJONCTION
'Séparateur     FCIGCO112       As String * 1                     ' SEPERATEUR
    FCIGCOIJ2       As String * 1                     ' CAS N¢2 INJONCTION
'Séparateur     FCIGCO113       As String * 1                     ' SEPERATEUR
    FCIGCOIJ3       As String * 1                     ' CAS N¢3 INJONCTION
'Séparateur     FCIGCO114       As String * 1                     ' SEPERATEUR
    FCIGCOIJ4       As String * 1                     ' NON UTILISE PREVISIONN
'Séparateur     FCIGCO115       As String * 1                     ' SEPERATEUR
    FCIGCOMSD       As String * 20                    ' MONTANT SOLDE DISPON
'Séparateur     FCIGCO116       As String * 1                     ' SEPERATEUR
    FCIGCODES       As String * 3                     ' DEVISE SOLDE DISPONI
'Séparateur     FCIGCO117       As String * 1                     ' SEPERATEUR
    FCIGCODEB       As String * 12                    ' MENTION DEBITEUR
'Séparateur     FCIGCO118       As String * 1                     ' SEPERATEUR
    FCIGCOIC1       As String * 1                     ' CAS N¢1 PAS PAYE
'Séparateur     FCIGCO119       As String * 1                     ' SEPERATEUR
    FCIGCOIC2       As String * 1                     ' CAS N¢2 PAYE PARTIEL
'Séparateur     FCIGCO120       As String * 1                     ' SEPERATEUR
    FCIGCOMPP       As String * 20                    ' MONTANT PAIEMT PARTI
'Séparateur     FCIGCO121       As String * 1                     ' SEPERATEUR
    FCIGCODPP       As String * 3                     ' DEVISE MT PAIT PARTI
'Séparateur     FCIGCO122       As String * 1                     ' SEPERATEUR
    FCIGCOCH1       As String * 7                     ' NUMERO CHEQUE 1
'Séparateur     FCIGCO123       As String * 1                     ' SEPERATEUR
    FCIGCOMC1       As String * 20                    ' MONTANT CHEQUE 1
'Séparateur     FCIGCO124       As String * 1                     ' SEPERATEUR
    FCIGCODE1       As String * 3                     ' DEVISE MONTANT 1
'Séparateur     FCIGCO125       As String * 1                     ' SEPERATEUR
    FCIGCOCH2       As String * 7                     ' NUMERO CHEQUE 2
'Séparateur     FCIGCO126       As String * 1                     ' SEPERATEUR
    FCIGCOMC2       As String * 20                    ' MONTANT CHEQUE 2
'Séparateur     FCIGCO127       As String * 1                     ' SEPERATEUR
    FCIGCODE2       As String * 3                     ' DEVISE MONTANT 2
'Séparateur     FCIGCO128       As String * 1                     ' SEPERATEUR
    FCIGCOCH3       As String * 7                     ' NUMERO CHEQUE 3
'Séparateur     FCIGCO129       As String * 1                     ' SEPERATEUR
    FCIGCOMC3       As String * 20                    ' MONTANT CHEQUE 3
'Séparateur     FCIGCO130       As String * 1                     ' SEPERATEUR
    FCIGCODE3       As String * 3                     ' DEVISE MONTANT 3
'Séparateur     FCIGCO131       As String * 1                     ' SEPERATEUR
    FCIGCOCH4       As String * 7                     ' NUMERO CHEQUE 4
'Séparateur     FCIGCO132       As String * 1                     ' SEPERATEUR
    FCIGCOMC4       As String * 20                    ' MONTANT CHEQUE 4
'Séparateur     FCIGCO133       As String * 1                     ' SEPERATEUR
    FCIGCODE4       As String * 3                     ' DEVISE MONTANT 4
'Séparateur     FCIGCO134       As String * 1                     ' SEPERATEUR
    FCIGCOCH5       As String * 7                     ' NUMERO CHEQUE 5
'Séparateur     FCIGCO135       As String * 1                     ' SEPERATEUR
    FCIGCOMC5       As String * 20                    ' MONTANT CHEQUE 5
'Séparateur     FCIGCO136       As String * 1                     ' SEPERATEUR
    FCIGCODE5       As String * 3                     ' DEVISE MONTANT 5
'Séparateur     FCIGCO137       As String * 1                     ' SEPERATEUR
    FCIGCOCH6       As String * 7                     ' NON UTILISE PREVISIO
'Séparateur     FCIGCO138       As String * 1                     ' SEPERATEUR
    FCIGCOMC6       As String * 20                    ' NON UTILISE PREVISIO
'Séparateur     FCIGCO139       As String * 1                     ' SEPERATEUR
    FCIGCODE6       As String * 3                     ' NON UTILISE PREVISIO
'Séparateur     FCIGCO140       As String * 1                     ' SEPERATEUR
    FCIGCODRJ       As String * 10                    ' DATE REJET DES CHQ
'Séparateur     FCIGCO141       As String * 1                     ' SEPERATEUR
    FCIGCODEI       As String * 10                    ' DATE DEPART INTERDIT
'Séparateur     FCIGCO142       As String * 1                     ' SEPERATEUR
    FCIGCODLR       As String * 10                    ' DATE LIMITE REGULARI
'Séparateur     FCIGCO143       As String * 1                     ' SEPERATEUR
    FCIGCOMPN       As String * 20                    ' MONTANT PENALITE
'Séparateur     FCIGCO144       As String * 1                     ' SEPERATEUR
    FCIGCODPN       As String * 3                     ' DEVISE MT PENALITE
'Séparateur     FCIGCO145       As String * 1                     ' SEPERATEUR
    FCIGCOJ21       As String * 1                     ' CAS N1 INJ 2 EV PREC
'Séparateur     FCIGCO146       As String * 1                     ' SEPERATEUR
    FCIGCOJ22       As String * 1                     ' CAS N2 INJ 2 EV PREC
'Séparateur     FCIGCO147       As String * 1                     ' SEPERATEUR
    FCIGCOINT       As String * 30                    ' INTITULE COMPTE
'Séparateur     FCIGCO148       As String * 1                     ' SEPERATEUR
    FCIGCONAG       As String * 32                    ' NOM AGENCE
'Séparateur     FCIGCO149       As String * 1                     ' SEPERATEUR
    FCIGCODAP       As String * 10                    ' DU CHEQUE
'Séparateur     FCIGCO150       As String * 1                     ' SEPERATEUR
    FCIGCOMIM       As String * 20                    ' MONTANT IMPAYE
'Séparateur     FCIGCO151       As String * 1                     ' SEPERATEUR
    FCIGCODIM       As String * 3                     ' DEVISE MT IMPAYE
'Séparateur     FCIGCO152       As String * 1                     ' SEPERATEUR
    FCIGCOCHB       As String * 7                     ' FRAIS PUBLICITE
'Séparateur     FCIGCO153       As String * 1                     ' SEPERATEUR
    FCIGCOMCB       As String * 20                    ' MONTANT CHQ BANQUE
'Séparateur     FCIGCO154       As String * 1                     ' SEPERATEUR
    FCIGCODCH       As String * 3                     ' DEVISE MT CHQ BANQUE
'Séparateur     FCIGCO155       As String * 1                     ' SEPERATEUR
    FCIGCONBQ       As String * 32                    ' NOM BANQUE
'Séparateur     FCIGCO156       As String * 1                     ' SEPERATEUR
    FCIGCOMTA       As String * 20                    ' MONTANT ABUSIF
'Séparateur     FCIGCO157       As String * 1                     ' SEPERATEUR
    FCIGCODEA       As String * 3                     ' DEVISE MONT. ABUSIF
'Séparateur     FCIGCO158       As String * 1                     ' SEPERATEUR
    FCIGCOLCA       As String * 32                    ' LIBELLE TYPE CARTE
'Séparateur     FCIGCO159       As String * 1                     ' SEPERATEUR
    FCIGCOLNA       As String * 32                    ' LIBELLE NATURE CARTE
'Séparateur     FCIGCO160       As String * 1                     ' SEPERATEUR
    FCIGCOVAL       As String * 10                    ' DATE VALIDITE
'Séparateur     FCIGCO161       As String * 1                     ' SEPERATEUR
    FCIGCODUR       As String * 2                     ' DUREE VALIDITE
'Séparateur     FCIGCO162       As String * 1                     ' SEPERATEUR
    FCIGCOLI1       As String * 32                    ' CHAMP LIBRE 1
'Séparateur     FCIGCO163       As String * 1                     ' SEPERATEUR
    FCIGCOLI2       As String * 32                    ' CHAMP LIBRE 2
'Séparateur     FCIGCO164       As String * 1                     ' SEPERATEUR
    FCIGCOPUP       As String * 10                    ' DATE PURGE POSSIBLE
'Séparateur     FCIGCO165       As String * 1                     ' SEPERATEUR
    FCIGCOTYI       As String * 30                    ' TYPE INTERDIT
'Séparateur     FCIGCO166       As String * 1                     ' SEPERATEUR
    FCIGCODDI       As String * 10                    ' TION INTERNE
'Séparateur     FCIGCO167       As String * 1                     ' SEPERATEUR
    FCIGCODFI       As String * 10                    ' ON INTERNE
'Séparateur     FCIGCO168       As String * 1                     ' SEPERATEUR
    FCIGCODDB       As String * 10                    ' TION BANCAIRE
'Séparateur     FCIGCO169       As String * 1                     ' SEPERATEUR
    FCIGCODFB       As String * 10                    ' DATE FIN INTERDICTI-
'Séparateur     FCIGCO170       As String * 1                     ' SEPERATEUR
    FCIGCOACP       As String * 20                    ' COMPTE AV CONVERSION
'Séparateur     FCIGCO171       As String * 1                     ' SEPERATEUR
    FCIGCOIBA       As String * 20                    ' NUMERO IBAN

End Type

Public Function rsZFCIGCO0_GetBuffer(rsAdo As ADODB.Recordset, rsZFCIGCO0 As typeZFCIGCO0)
On Error GoTo Error_Handler
rsZFCIGCO0_GetBuffer = Null
rsZFCIGCO0.FCIGCOETA = rsAdo("FCIGCOETA")
'Séparateur rsZFCIGCO0.FCIGCOS01 = rsADO("FCIGCOS01")
rsZFCIGCO0.FCIGCOCLI = rsAdo("FCIGCOCLI")
'Séparateur rsZFCIGCO0.FCIGCOS02 = rsADO("FCIGCOS02")
rsZFCIGCO0.FCIGCOPLA = rsAdo("FCIGCOPLA")
'Séparateur rsZFCIGCO0.FCIGCOS03 = rsADO("FCIGCOS03")
rsZFCIGCO0.FCIGCOCPT = rsAdo("FCIGCOCPT")
'Séparateur rsZFCIGCO0.FCIGCOS04 = rsADO("FCIGCOS04")
rsZFCIGCO0.FCIGCONUC = rsAdo("FCIGCONUC")
'Séparateur rsZFCIGCO0.FCIGCOS05 = rsADO("FCIGCOS05")
rsZFCIGCO0.FCIGCOCAR = rsAdo("FCIGCOCAR")
'Séparateur rsZFCIGCO0.FCIGCOS06 = rsADO("FCIGCOS06")
rsZFCIGCO0.FCIGCOSES = rsAdo("FCIGCOSES")
'Séparateur rsZFCIGCO0.FCIGCOS07 = rsADO("FCIGCOS07")
rsZFCIGCO0.FCIGCOSEA = rsAdo("FCIGCOSEA")
'Séparateur rsZFCIGCO0.FCIGCOS08 = rsADO("FCIGCOS08")
rsZFCIGCO0.FCIGCODLI = rsAdo("FCIGCODLI")
'Séparateur rsZFCIGCO0.FCIGCOS09 = rsADO("FCIGCOS09")
rsZFCIGCO0.FCIGCODAJ = rsAdo("FCIGCODAJ")
'Séparateur rsZFCIGCO0.FCIGCOS10 = rsADO("FCIGCOS10")
rsZFCIGCO0.FCIGCOCOU = rsAdo("FCIGCOCOU")
'Séparateur rsZFCIGCO0.FCIGCOS11 = rsADO("FCIGCOS11")
rsZFCIGCO0.FCIGCOLIB = rsAdo("FCIGCOLIB")
'Séparateur rsZFCIGCO0.FCIGCOS12 = rsADO("FCIGCOS12")
rsZFCIGCO0.FCIGCOTYC = rsAdo("FCIGCOTYC")
'Séparateur rsZFCIGCO0.FCIGCOS13 = rsADO("FCIGCOS13")
rsZFCIGCO0.FCIGCOLTY = rsAdo("FCIGCOLTY")
'Séparateur rsZFCIGCO0.FCIGCOS14 = rsADO("FCIGCOS14")
rsZFCIGCO0.FCIGCOENV = rsAdo("FCIGCOENV")
'Séparateur rsZFCIGCO0.FCIGCOS15 = rsADO("FCIGCOS15")
rsZFCIGCO0.FCIGCOREC = rsAdo("FCIGCOREC")
'Séparateur rsZFCIGCO0.FCIGCOS16 = rsADO("FCIGCOS16")
rsZFCIGCO0.FCIGCODCP = rsAdo("FCIGCODCP")
'Séparateur rsZFCIGCO0.FCIGCOS17 = rsADO("FCIGCOS17")
rsZFCIGCO0.FCIGCOEDI = rsAdo("FCIGCOEDI")
'Séparateur rsZFCIGCO0.FCIGCOS18 = rsADO("FCIGCOS18")
rsZFCIGCO0.FCIGCORED = rsAdo("FCIGCORED")
'Séparateur rsZFCIGCO0.FCIGCOS19 = rsADO("FCIGCOS19")
rsZFCIGCO0.FCIGCONDE = rsAdo("FCIGCONDE")
'Séparateur rsZFCIGCO0.FCIGCOS20 = rsADO("FCIGCOS20")
rsZFCIGCO0.FCIGCOLTD = rsAdo("FCIGCOLTD")
'Séparateur rsZFCIGCO0.FCIGCOS21 = rsADO("FCIGCOS21")
rsZFCIGCO0.FCIGCONRD = rsAdo("FCIGCONRD")
'Séparateur rsZFCIGCO0.FCIGCOS22 = rsADO("FCIGCOS22")
rsZFCIGCO0.FCIGCOPRD = rsAdo("FCIGCOPRD")
'Séparateur rsZFCIGCO0.FCIGCOS23 = rsADO("FCIGCOS23")
rsZFCIGCO0.FCIGCOA1D = rsAdo("FCIGCOA1D")
'Séparateur rsZFCIGCO0.FCIGCOS24 = rsADO("FCIGCOS24")
rsZFCIGCO0.FCIGCOA2D = rsAdo("FCIGCOA2D")
'Séparateur rsZFCIGCO0.FCIGCOS25 = rsADO("FCIGCOS25")
rsZFCIGCO0.FCIGCOA3D = rsAdo("FCIGCOA3D")
'Séparateur rsZFCIGCO0.FCIGCOS26 = rsADO("FCIGCOS26")
rsZFCIGCO0.FCIGCOCPD = rsAdo("FCIGCOCPD")
'Séparateur rsZFCIGCO0.FCIGCOS27 = rsADO("FCIGCOS27")
rsZFCIGCO0.FCIGCOVID = rsAdo("FCIGCOVID")
'Séparateur rsZFCIGCO0.FCIGCOS28 = rsADO("FCIGCOS28")
rsZFCIGCO0.FCIGCOLPD = rsAdo("FCIGCOLPD")
'Séparateur rsZFCIGCO0.FCIGCOS29 = rsADO("FCIGCOS29")
rsZFCIGCO0.FCIGCOCLD = rsAdo("FCIGCOCLD")
'Séparateur rsZFCIGCO0.FCIGCOS30 = rsADO("FCIGCOS30")
rsZFCIGCO0.FCIGCOLTC = rsAdo("FCIGCOLTC")
'Séparateur rsZFCIGCO0.FCIGCOS31 = rsADO("FCIGCOS31")
rsZFCIGCO0.FCIGCONRC = rsAdo("FCIGCONRC")
'Séparateur rsZFCIGCO0.FCIGCOS32 = rsADO("FCIGCOS32")
rsZFCIGCO0.FCIGCOPRC = rsAdo("FCIGCOPRC")
'Séparateur rsZFCIGCO0.FCIGCOS33 = rsADO("FCIGCOS33")
rsZFCIGCO0.FCIGCOAD1 = rsAdo("FCIGCOAD1")
'Séparateur rsZFCIGCO0.FCIGCOS34 = rsADO("FCIGCOS34")
rsZFCIGCO0.FCIGCOAD2 = rsAdo("FCIGCOAD2")
'Séparateur rsZFCIGCO0.FCIGCOS35 = rsADO("FCIGCOS35")
rsZFCIGCO0.FCIGCOAD3 = rsAdo("FCIGCOAD3")
'Séparateur rsZFCIGCO0.FCIGCOS36 = rsADO("FCIGCOS36")
rsZFCIGCO0.FCIGCOCPC = rsAdo("FCIGCOCPC")
'Séparateur rsZFCIGCO0.FCIGCOS37 = rsADO("FCIGCOS37")
rsZFCIGCO0.FCIGCOVIC = rsAdo("FCIGCOVIC")
'Séparateur rsZFCIGCO0.FCIGCOS38 = rsADO("FCIGCOS38")
rsZFCIGCO0.FCIGCOLPC = rsAdo("FCIGCOLPC")
'Séparateur rsZFCIGCO0.FCIGCOS39 = rsADO("FCIGCOS39")
rsZFCIGCO0.FCIGCOCLB = rsAdo("FCIGCOCLB")
'Séparateur rsZFCIGCO0.FCIGCOS40 = rsADO("FCIGCOS40")
rsZFCIGCO0.FCIGCOLTB = rsAdo("FCIGCOLTB")
'Séparateur rsZFCIGCO0.FCIGCOS41 = rsADO("FCIGCOS41")
rsZFCIGCO0.FCIGCOBNR = rsAdo("FCIGCOBNR")
'Séparateur rsZFCIGCO0.FCIGCOS42 = rsADO("FCIGCOS42")
rsZFCIGCO0.FCIGCOBPR = rsAdo("FCIGCOBPR")
'Séparateur rsZFCIGCO0.FCIGCOS43 = rsADO("FCIGCOS43")
rsZFCIGCO0.FCIGCOA1B = rsAdo("FCIGCOA1B")
'Séparateur rsZFCIGCO0.FCIGCOS44 = rsADO("FCIGCOS44")
rsZFCIGCO0.FCIGCOA2B = rsAdo("FCIGCOA2B")
'Séparateur rsZFCIGCO0.FCIGCOS45 = rsADO("FCIGCOS45")
rsZFCIGCO0.FCIGCOA3B = rsAdo("FCIGCOA3B")
'Séparateur rsZFCIGCO0.FCIGCOS46 = rsADO("FCIGCOS46")
rsZFCIGCO0.FCIGCOCPB = rsAdo("FCIGCOCPB")
'Séparateur rsZFCIGCO0.FCIGCOS47 = rsADO("FCIGCOS47")
rsZFCIGCO0.FCIGCOBVI = rsAdo("FCIGCOBVI")
'Séparateur rsZFCIGCO0.FCIGCOS48 = rsADO("FCIGCOS48")
rsZFCIGCO0.FCIGCOLPB = rsAdo("FCIGCOLPB")
'Séparateur rsZFCIGCO0.FCIGCOS49 = rsADO("FCIGCOS49")
rsZFCIGCO0.FCIGCOCLP = rsAdo("FCIGCOCLP")
'Séparateur rsZFCIGCO0.FCIGCOS50 = rsADO("FCIGCOS50")
rsZFCIGCO0.FCIGCOLTP = rsAdo("FCIGCOLTP")
'Séparateur rsZFCIGCO0.FCIGCOS51 = rsADO("FCIGCOS51")
rsZFCIGCO0.FCIGCONRP = rsAdo("FCIGCONRP")
'Séparateur rsZFCIGCO0.FCIGCOS52 = rsADO("FCIGCOS52")
rsZFCIGCO0.FCIGCOPRP = rsAdo("FCIGCOPRP")
'Séparateur rsZFCIGCO0.FCIGCOS53 = rsADO("FCIGCOS53")
'Séparateur rsZFCIGCO0.FCIGCO1DP = rsADO("FCIGCO1DP")
'Séparateur rsZFCIGCO0.FCIGCOS54 = rsADO("FCIGCOS54")
rsZFCIGCO0.FCIGCO2DP = rsAdo("FCIGCO2DP")
'Séparateur rsZFCIGCO0.FCIGCOS55 = rsADO("FCIGCOS55")
rsZFCIGCO0.FCIGCO3DP = rsAdo("FCIGCO3DP")
'Séparateur rsZFCIGCO0.FCIGCOS56 = rsADO("FCIGCOS56")
rsZFCIGCO0.FCIGCOCPP = rsAdo("FCIGCOCPP")
'Séparateur rsZFCIGCO0.FCIGCOS57 = rsADO("FCIGCOS57")
rsZFCIGCO0.FCIGCOVIP = rsAdo("FCIGCOVIP")
'Séparateur rsZFCIGCO0.FCIGCOS58 = rsADO("FCIGCOS58")
rsZFCIGCO0.FCIGCOLPP = rsAdo("FCIGCOLPP")
'Séparateur rsZFCIGCO0.FCIGCOS59 = rsADO("FCIGCOS59")
rsZFCIGCO0.FCIGCOCLT = rsAdo("FCIGCOCLT")
'Séparateur rsZFCIGCO0.FCIGCOS60 = rsADO("FCIGCOS60")
rsZFCIGCO0.FCIGCOLIT = rsAdo("FCIGCOLIT")
'Séparateur rsZFCIGCO0.FCIGCOS61 = rsADO("FCIGCOS61")
rsZFCIGCO0.FCIGCONOT = rsAdo("FCIGCONOT")
'Séparateur rsZFCIGCO0.FCIGCOS62 = rsADO("FCIGCOS62")
rsZFCIGCO0.FCIGCOPRT = rsAdo("FCIGCOPRT")
'Séparateur rsZFCIGCO0.FCIGCOS63 = rsADO("FCIGCOS63")
rsZFCIGCO0.FCIGCO1DT = rsAdo("FCIGCO1DT")
'Séparateur rsZFCIGCO0.FCIGCOS64 = rsADO("FCIGCOS64")
rsZFCIGCO0.FCIGCO2DT = rsAdo("FCIGCO2DT")
'Séparateur rsZFCIGCO0.FCIGCOS65 = rsADO("FCIGCOS65")
rsZFCIGCO0.FCIGCO3DT = rsAdo("FCIGCO3DT")
'Séparateur rsZFCIGCO0.FCIGCOS66 = rsADO("FCIGCOS66")
rsZFCIGCO0.FCIGCOPOT = rsAdo("FCIGCOPOT")
'Séparateur rsZFCIGCO0.FCIGCOS67 = rsADO("FCIGCOS67")
rsZFCIGCO0.FCIGCOVIT = rsAdo("FCIGCOVIT")
'Séparateur rsZFCIGCO0.FCIGCOS68 = rsADO("FCIGCOS68")
rsZFCIGCO0.FCIGCOLPT = rsAdo("FCIGCOLPT")
'Séparateur rsZFCIGCO0.FCIGCOS69 = rsADO("FCIGCOS69")
rsZFCIGCO0.FCIGCOCLC = rsAdo("FCIGCOCLC")
'Séparateur rsZFCIGCO0.FCIGCOS70 = rsADO("FCIGCOS70")
rsZFCIGCO0.FCIGCOLIC = rsAdo("FCIGCOLIC")
'Séparateur rsZFCIGCO0.FCIGCOS71 = rsADO("FCIGCOS71")
rsZFCIGCO0.FCIGCONOC = rsAdo("FCIGCONOC")
'Séparateur rsZFCIGCO0.FCIGCOS72 = rsADO("FCIGCOS72")
rsZFCIGCO0.FCIGCOPCO = rsAdo("FCIGCOPCO")
'Séparateur rsZFCIGCO0.FCIGCOS73 = rsADO("FCIGCOS73")
rsZFCIGCO0.FCIGCO1DC = rsAdo("FCIGCO1DC")
'Séparateur rsZFCIGCO0.FCIGCOS74 = rsADO("FCIGCOS74")
rsZFCIGCO0.FCIGCO2DC = rsAdo("FCIGCO2DC")
'Séparateur rsZFCIGCO0.FCIGCOS75 = rsADO("FCIGCOS75")
rsZFCIGCO0.FCIGCO3DC = rsAdo("FCIGCO3DC")
'Séparateur rsZFCIGCO0.FCIGCOS76 = rsADO("FCIGCOS76")
rsZFCIGCO0.FCIGCOPOC = rsAdo("FCIGCOPOC")
'Séparateur rsZFCIGCO0.FCIGCOS77 = rsADO("FCIGCOS77")
rsZFCIGCO0.FCIGCOVLC = rsAdo("FCIGCOVLC")
'Séparateur rsZFCIGCO0.FCIGCOS78 = rsADO("FCIGCOS78")
rsZFCIGCO0.FCIGCOPAC = rsAdo("FCIGCOPAC")
'Séparateur rsZFCIGCO0.FCIGCOS79 = rsADO("FCIGCOS79")
rsZFCIGCO0.FCIGCOCLM = rsAdo("FCIGCOCLM")
'Séparateur rsZFCIGCO0.FCIGCOS80 = rsADO("FCIGCOS80")
rsZFCIGCO0.FCIGCOLIM = rsAdo("FCIGCOLIM")
'Séparateur rsZFCIGCO0.FCIGCOS81 = rsADO("FCIGCOS81")
rsZFCIGCO0.FCIGCONOM = rsAdo("FCIGCONOM")
'Séparateur rsZFCIGCO0.FCIGCOS82 = rsADO("FCIGCOS82")
rsZFCIGCO0.FCIGCOPRM = rsAdo("FCIGCOPRM")
'Séparateur rsZFCIGCO0.FCIGCOS83 = rsADO("FCIGCOS83")
rsZFCIGCO0.FCIGCO1DM = rsAdo("FCIGCO1DM")
'Séparateur rsZFCIGCO0.FCIGCOS84 = rsADO("FCIGCOS84")
rsZFCIGCO0.FCIGCO2DM = rsAdo("FCIGCO2DM")
'Séparateur rsZFCIGCO0.FCIGCOS85 = rsADO("FCIGCOS85")
rsZFCIGCO0.FCIGCO3DM = rsAdo("FCIGCO3DM")
'Séparateur rsZFCIGCO0.FCIGCOS86 = rsADO("FCIGCOS86")
rsZFCIGCO0.FCIGCOCPM = rsAdo("FCIGCOCPM")
'Séparateur rsZFCIGCO0.FCIGCOS87 = rsADO("FCIGCOS87")
rsZFCIGCO0.FCIGCOVIM = rsAdo("FCIGCOVIM")
'Séparateur rsZFCIGCO0.FCIGCOS88 = rsADO("FCIGCOS88")
rsZFCIGCO0.FCIGCOLPM = rsAdo("FCIGCOLPM")
'Séparateur rsZFCIGCO0.FCIGCOS89 = rsADO("FCIGCOS89")
rsZFCIGCO0.FCIGCOCLG = rsAdo("FCIGCOCLG")
'Séparateur rsZFCIGCO0.FCIGCOS90 = rsADO("FCIGCOS90")
rsZFCIGCO0.FCIGCOLIG = rsAdo("FCIGCOLIG")
'Séparateur rsZFCIGCO0.FCIGCOS91 = rsADO("FCIGCOS91")
rsZFCIGCO0.FCIGCONOG = rsAdo("FCIGCONOG")
'Séparateur rsZFCIGCO0.FCIGCOS92 = rsADO("FCIGCOS92")
rsZFCIGCO0.FCIGCOPRG = rsAdo("FCIGCOPRG")
'Séparateur rsZFCIGCO0.FCIGCOS93 = rsADO("FCIGCOS93")
rsZFCIGCO0.FCIGCO1DG = rsAdo("FCIGCO1DG")
'Séparateur rsZFCIGCO0.FCIGCOS94 = rsADO("FCIGCOS94")
rsZFCIGCO0.FCIGCO2DG = rsAdo("FCIGCO2DG")
'Séparateur rsZFCIGCO0.FCIGCOS95 = rsADO("FCIGCOS95")
rsZFCIGCO0.FCIGCO3DG = rsAdo("FCIGCO3DG")
'Séparateur rsZFCIGCO0.FCIGCOS96 = rsADO("FCIGCOS96")
rsZFCIGCO0.FCIGCOCPG = rsAdo("FCIGCOCPG")
'Séparateur rsZFCIGCO0.FCIGCOS97 = rsADO("FCIGCOS97")
rsZFCIGCO0.FCIGCOVIG = rsAdo("FCIGCOVIG")
'Séparateur rsZFCIGCO0.FCIGCOS98 = rsADO("FCIGCOS98")
rsZFCIGCO0.FCIGCOLPG = rsAdo("FCIGCOLPG")
'Séparateur rsZFCIGCO0.FCIGCOS99 = rsADO("FCIGCOS99")
rsZFCIGCO0.FCIGCOLED = rsAdo("FCIGCOLED")
'Séparateur rsZFCIGCO0.FCIGCO100 = rsADO("FCIGCO100")
rsZFCIGCO0.FCIGCOGES = rsAdo("FCIGCOGES")
'Séparateur rsZFCIGCO0.FCIGCO101 = rsADO("FCIGCO101")
rsZFCIGCO0.FCIGCOREL = rsAdo("FCIGCOREL")
'Séparateur rsZFCIGCO0.FCIGCO102 = rsADO("FCIGCO102")
rsZFCIGCO0.FCIGCOTEL = rsAdo("FCIGCOTEL")
'Séparateur rsZFCIGCO0.FCIGCO103 = rsADO("FCIGCO103")
rsZFCIGCO0.FCIGCOREJ = rsAdo("FCIGCOREJ")
'Séparateur rsZFCIGCO0.FCIGCO104 = rsADO("FCIGCO104")
rsZFCIGCO0.FCIGCOLIR = rsAdo("FCIGCOLIR")
'Séparateur rsZFCIGCO0.FCIGCO105 = rsADO("FCIGCO105")
rsZFCIGCO0.FCIGCOMCH = rsAdo("FCIGCOMCH")
'Séparateur rsZFCIGCO0.FCIGCO106 = rsADO("FCIGCO106")
rsZFCIGCO0.FCIGCODEV = rsAdo("FCIGCODEV")
'Séparateur rsZFCIGCO0.FCIGCO107 = rsADO("FCIGCO107")
rsZFCIGCO0.FCIGCOAT1 = rsAdo("FCIGCOAT1")
'Séparateur rsZFCIGCO0.FCIGCO108 = rsADO("FCIGCO108")
rsZFCIGCO0.FCIGCOAT2 = rsAdo("FCIGCOAT2")
'Séparateur rsZFCIGCO0.FCIGCO109 = rsADO("FCIGCO109")
rsZFCIGCO0.FCIGCOAT3 = rsAdo("FCIGCOAT3")
'Séparateur rsZFCIGCO0.FCIGCO110 = rsADO("FCIGCO110")
rsZFCIGCO0.FCIGCOAT4 = rsAdo("FCIGCOAT4")
'Séparateur rsZFCIGCO0.FCIGCO111 = rsADO("FCIGCO111")
rsZFCIGCO0.FCIGCOIJ1 = rsAdo("FCIGCOIJ1")
'Séparateur rsZFCIGCO0.FCIGCO112 = rsADO("FCIGCO112")
rsZFCIGCO0.FCIGCOIJ2 = rsAdo("FCIGCOIJ2")
'Séparateur rsZFCIGCO0.FCIGCO113 = rsADO("FCIGCO113")
rsZFCIGCO0.FCIGCOIJ3 = rsAdo("FCIGCOIJ3")
'Séparateur rsZFCIGCO0.FCIGCO114 = rsADO("FCIGCO114")
rsZFCIGCO0.FCIGCOIJ4 = rsAdo("FCIGCOIJ4")
'Séparateur rsZFCIGCO0.FCIGCO115 = rsADO("FCIGCO115")
rsZFCIGCO0.FCIGCOMSD = rsAdo("FCIGCOMSD")
'Séparateur rsZFCIGCO0.FCIGCO116 = rsADO("FCIGCO116")
rsZFCIGCO0.FCIGCODES = rsAdo("FCIGCODES")
'Séparateur rsZFCIGCO0.FCIGCO117 = rsADO("FCIGCO117")
rsZFCIGCO0.FCIGCODEB = rsAdo("FCIGCODEB")
'Séparateur rsZFCIGCO0.FCIGCO118 = rsADO("FCIGCO118")
rsZFCIGCO0.FCIGCOIC1 = rsAdo("FCIGCOIC1")
'Séparateur rsZFCIGCO0.FCIGCO119 = rsADO("FCIGCO119")
rsZFCIGCO0.FCIGCOIC2 = rsAdo("FCIGCOIC2")
'Séparateur rsZFCIGCO0.FCIGCO120 = rsADO("FCIGCO120")
rsZFCIGCO0.FCIGCOMPP = rsAdo("FCIGCOMPP")
'Séparateur rsZFCIGCO0.FCIGCO121 = rsADO("FCIGCO121")
rsZFCIGCO0.FCIGCODPP = rsAdo("FCIGCODPP")
'Séparateur rsZFCIGCO0.FCIGCO122 = rsADO("FCIGCO122")
rsZFCIGCO0.FCIGCOCH1 = rsAdo("FCIGCOCH1")
'Séparateur rsZFCIGCO0.FCIGCO123 = rsADO("FCIGCO123")
rsZFCIGCO0.FCIGCOMC1 = rsAdo("FCIGCOMC1")
'Séparateur rsZFCIGCO0.FCIGCO124 = rsADO("FCIGCO124")
rsZFCIGCO0.FCIGCODE1 = rsAdo("FCIGCODE1")
'Séparateur rsZFCIGCO0.FCIGCO125 = rsADO("FCIGCO125")
rsZFCIGCO0.FCIGCOCH2 = rsAdo("FCIGCOCH2")
'Séparateur rsZFCIGCO0.FCIGCO126 = rsADO("FCIGCO126")
rsZFCIGCO0.FCIGCOMC2 = rsAdo("FCIGCOMC2")
'Séparateur rsZFCIGCO0.FCIGCO127 = rsADO("FCIGCO127")
rsZFCIGCO0.FCIGCODE2 = rsAdo("FCIGCODE2")
'Séparateur rsZFCIGCO0.FCIGCO128 = rsADO("FCIGCO128")
rsZFCIGCO0.FCIGCOCH3 = rsAdo("FCIGCOCH3")
'Séparateur rsZFCIGCO0.FCIGCO129 = rsADO("FCIGCO129")
rsZFCIGCO0.FCIGCOMC3 = rsAdo("FCIGCOMC3")
'Séparateur rsZFCIGCO0.FCIGCO130 = rsADO("FCIGCO130")
rsZFCIGCO0.FCIGCODE3 = rsAdo("FCIGCODE3")
'Séparateur rsZFCIGCO0.FCIGCO131 = rsADO("FCIGCO131")
rsZFCIGCO0.FCIGCOCH4 = rsAdo("FCIGCOCH4")
'Séparateur rsZFCIGCO0.FCIGCO132 = rsADO("FCIGCO132")
rsZFCIGCO0.FCIGCOMC4 = rsAdo("FCIGCOMC4")
'Séparateur rsZFCIGCO0.FCIGCO133 = rsADO("FCIGCO133")
rsZFCIGCO0.FCIGCODE4 = rsAdo("FCIGCODE4")
'Séparateur rsZFCIGCO0.FCIGCO134 = rsADO("FCIGCO134")
rsZFCIGCO0.FCIGCOCH5 = rsAdo("FCIGCOCH5")
'Séparateur rsZFCIGCO0.FCIGCO135 = rsADO("FCIGCO135")
rsZFCIGCO0.FCIGCOMC5 = rsAdo("FCIGCOMC5")
'Séparateur rsZFCIGCO0.FCIGCO136 = rsADO("FCIGCO136")
rsZFCIGCO0.FCIGCODE5 = rsAdo("FCIGCODE5")
'Séparateur rsZFCIGCO0.FCIGCO137 = rsADO("FCIGCO137")
rsZFCIGCO0.FCIGCOCH6 = rsAdo("FCIGCOCH6")
'Séparateur rsZFCIGCO0.FCIGCO138 = rsADO("FCIGCO138")
rsZFCIGCO0.FCIGCOMC6 = rsAdo("FCIGCOMC6")
'Séparateur rsZFCIGCO0.FCIGCO139 = rsADO("FCIGCO139")
rsZFCIGCO0.FCIGCODE6 = rsAdo("FCIGCODE6")
'Séparateur rsZFCIGCO0.FCIGCO140 = rsADO("FCIGCO140")
rsZFCIGCO0.FCIGCODRJ = rsAdo("FCIGCODRJ")
'Séparateur rsZFCIGCO0.FCIGCO141 = rsADO("FCIGCO141")
rsZFCIGCO0.FCIGCODEI = rsAdo("FCIGCODEI")
'Séparateur rsZFCIGCO0.FCIGCO142 = rsADO("FCIGCO142")
rsZFCIGCO0.FCIGCODLR = rsAdo("FCIGCODLR")
'Séparateur rsZFCIGCO0.FCIGCO143 = rsADO("FCIGCO143")
rsZFCIGCO0.FCIGCOMPN = rsAdo("FCIGCOMPN")
'Séparateur rsZFCIGCO0.FCIGCO144 = rsADO("FCIGCO144")
rsZFCIGCO0.FCIGCODPN = rsAdo("FCIGCODPN")
'Séparateur rsZFCIGCO0.FCIGCO145 = rsADO("FCIGCO145")
rsZFCIGCO0.FCIGCOJ21 = rsAdo("FCIGCOJ21")
'Séparateur rsZFCIGCO0.FCIGCO146 = rsADO("FCIGCO146")
rsZFCIGCO0.FCIGCOJ22 = rsAdo("FCIGCOJ22")
'Séparateur rsZFCIGCO0.FCIGCO147 = rsADO("FCIGCO147")
rsZFCIGCO0.FCIGCOINT = rsAdo("FCIGCOINT")
'Séparateur rsZFCIGCO0.FCIGCO148 = rsADO("FCIGCO148")
rsZFCIGCO0.FCIGCONAG = rsAdo("FCIGCONAG")
'Séparateur rsZFCIGCO0.FCIGCO149 = rsADO("FCIGCO149")
rsZFCIGCO0.FCIGCODAP = rsAdo("FCIGCODAP")
'Séparateur rsZFCIGCO0.FCIGCO150 = rsADO("FCIGCO150")
rsZFCIGCO0.FCIGCOMIM = rsAdo("FCIGCOMIM")
'Séparateur rsZFCIGCO0.FCIGCO151 = rsADO("FCIGCO151")
rsZFCIGCO0.FCIGCODIM = rsAdo("FCIGCODIM")
'Séparateur rsZFCIGCO0.FCIGCO152 = rsADO("FCIGCO152")
rsZFCIGCO0.FCIGCOCHB = rsAdo("FCIGCOCHB")
'Séparateur rsZFCIGCO0.FCIGCO153 = rsADO("FCIGCO153")
rsZFCIGCO0.FCIGCOMCB = rsAdo("FCIGCOMCB")
'Séparateur rsZFCIGCO0.FCIGCO154 = rsADO("FCIGCO154")
rsZFCIGCO0.FCIGCODCH = rsAdo("FCIGCODCH")
'Séparateur rsZFCIGCO0.FCIGCO155 = rsADO("FCIGCO155")
rsZFCIGCO0.FCIGCONBQ = rsAdo("FCIGCONBQ")
'Séparateur rsZFCIGCO0.FCIGCO156 = rsADO("FCIGCO156")
rsZFCIGCO0.FCIGCOMTA = rsAdo("FCIGCOMTA")
'Séparateur rsZFCIGCO0.FCIGCO157 = rsADO("FCIGCO157")
rsZFCIGCO0.FCIGCODEA = rsAdo("FCIGCODEA")
'Séparateur rsZFCIGCO0.FCIGCO158 = rsADO("FCIGCO158")
rsZFCIGCO0.FCIGCOLCA = rsAdo("FCIGCOLCA")
'Séparateur rsZFCIGCO0.FCIGCO159 = rsADO("FCIGCO159")
rsZFCIGCO0.FCIGCOLNA = rsAdo("FCIGCOLNA")
'Séparateur rsZFCIGCO0.FCIGCO160 = rsADO("FCIGCO160")
rsZFCIGCO0.FCIGCOVAL = rsAdo("FCIGCOVAL")
'Séparateur rsZFCIGCO0.FCIGCO161 = rsADO("FCIGCO161")
rsZFCIGCO0.FCIGCODUR = rsAdo("FCIGCODUR")
'Séparateur rsZFCIGCO0.FCIGCO162 = rsADO("FCIGCO162")
rsZFCIGCO0.FCIGCOLI1 = rsAdo("FCIGCOLI1")
'Séparateur rsZFCIGCO0.FCIGCO163 = rsADO("FCIGCO163")
rsZFCIGCO0.FCIGCOLI2 = rsAdo("FCIGCOLI2")
'Séparateur rsZFCIGCO0.FCIGCO164 = rsADO("FCIGCO164")
rsZFCIGCO0.FCIGCOPUP = rsAdo("FCIGCOPUP")
'Séparateur rsZFCIGCO0.FCIGCO165 = rsADO("FCIGCO165")
rsZFCIGCO0.FCIGCOTYI = rsAdo("FCIGCOTYI")
'Séparateur rsZFCIGCO0.FCIGCO166 = rsADO("FCIGCO166")
rsZFCIGCO0.FCIGCODDI = rsAdo("FCIGCODDI")
'Séparateur rsZFCIGCO0.FCIGCO167 = rsADO("FCIGCO167")
rsZFCIGCO0.FCIGCODFI = rsAdo("FCIGCODFI")
'Séparateur rsZFCIGCO0.FCIGCO168 = rsADO("FCIGCO168")
rsZFCIGCO0.FCIGCODDB = rsAdo("FCIGCODDB")
'Séparateur rsZFCIGCO0.FCIGCO169 = rsADO("FCIGCO169")
rsZFCIGCO0.FCIGCODFB = rsAdo("FCIGCODFB")
'Séparateur rsZFCIGCO0.FCIGCO170 = rsADO("FCIGCO170")
rsZFCIGCO0.FCIGCOACP = rsAdo("FCIGCOACP")
'Séparateur rsZFCIGCO0.FCIGCO171 = rsADO("FCIGCO171")
rsZFCIGCO0.FCIGCOIBA = rsAdo("FCIGCOIBA")
Exit Function
Error_Handler:
rsZFCIGCO0_GetBuffer = Error
End Function

Public Sub srvZFCIGCO0_fgDisplay(rsZFCIGCO0 As typeZFCIGCO0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 174
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "FCIGCOETA    4S"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "FCIGCOCLI    7A"
fgDisplay.Col = 1: fgDisplay = "RESPONSABLE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOCLI
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "FCIGCOPLA    3S"
fgDisplay.Col = 1: fgDisplay = "NUMERO PLAN"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOPLA
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "FCIGCOCPT   20A"
fgDisplay.Col = 1: fgDisplay = "NUMERO COMPTE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOCPT
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "FCIGCONUC    7S"
fgDisplay.Col = 1: fgDisplay = "NUMERO CHEQUE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCONUC
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "FCIGCOCAR   16A"
fgDisplay.Col = 1: fgDisplay = "NUMERO CARTE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOCAR
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "FCIGCOSES    5S"
fgDisplay.Col = 1: fgDisplay = "NUM¢ SEQUENCE STATUT"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOSES
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "FCIGCOSEA    3S"
fgDisplay.Col = 1: fgDisplay = "NUM¢ SEQUENCE ACTION"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOSEA
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "FCIGCODLI   10A"
fgDisplay.Col = 1: fgDisplay = "DATE LIMITE RETENT¢"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCODLI
fgDisplay.Row = 10
fgDisplay.Col = 0: fgDisplay = "FCIGCODAJ   10A"
fgDisplay.Col = 1: fgDisplay = "DATE JOUR"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCODAJ
fgDisplay.Row = 11
fgDisplay.Col = 0: fgDisplay = "FCIGCOS10    1A"
fgDisplay.Col = 1: fgDisplay = "SEPERATEUR"
fgDisplay.Col = 2: fgDisplay = ""
fgDisplay.Row = 12
fgDisplay.Col = 0: fgDisplay = "FCIGCOCOU    6A"
fgDisplay.Col = 1: fgDisplay = "CODE COURRIER TRANSM"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOCOU
fgDisplay.Row = 13
fgDisplay.Col = 0: fgDisplay = "FCIGCOLIB   30A"
fgDisplay.Col = 1: fgDisplay = "LIBELLE COURRIER"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOLIB
fgDisplay.Row = 14
fgDisplay.Col = 0: fgDisplay = "FCIGCOTYC    1A"
fgDisplay.Col = 1: fgDisplay = "TYPE COURRIER TRANSM"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOTYC
fgDisplay.Row = 15
fgDisplay.Col = 0: fgDisplay = "FCIGCOLTY   30A"
fgDisplay.Col = 1: fgDisplay = "LIBELLE TYPE COURR."
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOLTY
fgDisplay.Row = 16
fgDisplay.Col = 0: fgDisplay = "FCIGCOENV    1A"
fgDisplay.Col = 1: fgDisplay = "ENVOI RECOMANDE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOENV
fgDisplay.Row = 17
fgDisplay.Col = 0: fgDisplay = "FCIGCOREC   30A"
fgDisplay.Col = 1: fgDisplay = "LIBELLE RECOMMANDE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOREC
fgDisplay.Row = 18
fgDisplay.Col = 0: fgDisplay = "FCIGCODCP   10A"
fgDisplay.Col = 1: fgDisplay = "DATE COURRIER PRECED"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCODCP
fgDisplay.Row = 19
fgDisplay.Col = 0: fgDisplay = "FCIGCOEDI   10A"
fgDisplay.Col = 1: fgDisplay = "DATE EDITION"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOEDI
fgDisplay.Row = 20
fgDisplay.Col = 0: fgDisplay = "FCIGCORED    1A"
fgDisplay.Col = 1: fgDisplay = "REEDITION (O/N)"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCORED
fgDisplay.Row = 21
fgDisplay.Col = 0: fgDisplay = "FCIGCONDE    7A"
fgDisplay.Col = 1: fgDisplay = "NUMERO CLIENT DESTIN"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCONDE
fgDisplay.Row = 22
fgDisplay.Col = 0: fgDisplay = "FCIGCOLTD   30A"
fgDisplay.Col = 1: fgDisplay = "LIBELLE ETAT DESTINA"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOLTD
fgDisplay.Row = 23
fgDisplay.Col = 0: fgDisplay = "FCIGCONRD   32A"
fgDisplay.Col = 1: fgDisplay = "NOM/RAISON DESTINATA"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCONRD
fgDisplay.Row = 24
fgDisplay.Col = 0: fgDisplay = "FCIGCOPRD   32A"
fgDisplay.Col = 1: fgDisplay = "PRENOM/RAISON DESTIN"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOPRD
fgDisplay.Row = 25
fgDisplay.Col = 0: fgDisplay = "FCIGCOA1D   32A"
fgDisplay.Col = 1: fgDisplay = "ADRESSE 1 DESTINATAT"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOA1D
fgDisplay.Row = 26
fgDisplay.Col = 0: fgDisplay = "FCIGCOA2D   32A"
fgDisplay.Col = 1: fgDisplay = "ADRESSE 2 DESTINATAT"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOA2D
fgDisplay.Row = 27
fgDisplay.Col = 0: fgDisplay = "FCIGCOA3D   32A"
fgDisplay.Col = 1: fgDisplay = "ADRESSE 3 DESTINATAT"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOA3D
fgDisplay.Row = 28
fgDisplay.Col = 0: fgDisplay = "FCIGCOCPD    6A"
fgDisplay.Col = 1: fgDisplay = "CODE POSTAL DESTINAT"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOCPD
fgDisplay.Row = 29
fgDisplay.Col = 0: fgDisplay = "FCIGCOVID   25A"
fgDisplay.Col = 1: fgDisplay = "VILLE DESTINATAIRE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOVID
fgDisplay.Row = 30
fgDisplay.Col = 0: fgDisplay = "FCIGCOLPD   25A"
fgDisplay.Col = 1: fgDisplay = "LIBELL.PAYS DESTINAT"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOLPD
fgDisplay.Row = 31
fgDisplay.Col = 0: fgDisplay = "FCIGCOCLD    7A"
fgDisplay.Col = 1: fgDisplay = "NUMERO CLIENT"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOCLD
fgDisplay.Row = 32
fgDisplay.Col = 0: fgDisplay = "FCIGCOLTC   30A"
fgDisplay.Col = 1: fgDisplay = "LIBELLE ETAT CLIENT"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOLTC
fgDisplay.Row = 33
fgDisplay.Col = 0: fgDisplay = "FCIGCONRC   32A"
fgDisplay.Col = 1: fgDisplay = "NOM/RAISON CLIENT"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCONRC
fgDisplay.Row = 34
fgDisplay.Col = 0: fgDisplay = "FCIGCOPRC   32A"
fgDisplay.Col = 1: fgDisplay = "PRENOM/RAISON CLIENT"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOPRC
fgDisplay.Row = 35
fgDisplay.Col = 0: fgDisplay = "FCIGCOAD1   32A"
fgDisplay.Col = 1: fgDisplay = "ADRESSE 1 CLIENT"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOAD1
fgDisplay.Row = 36
fgDisplay.Col = 0: fgDisplay = "FCIGCOAD2   32A"
fgDisplay.Col = 1: fgDisplay = "ADRESSE 2 CLIENT"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOAD2
fgDisplay.Row = 37
fgDisplay.Col = 0: fgDisplay = "FCIGCOAD3   32A"
fgDisplay.Col = 1: fgDisplay = "ADRESSE 3 CLIENT"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOAD3
fgDisplay.Row = 38
fgDisplay.Col = 0: fgDisplay = "FCIGCOCPC    6A"
fgDisplay.Col = 1: fgDisplay = "CODE POSTAL CLIENT"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOCPC
fgDisplay.Row = 39
fgDisplay.Col = 0: fgDisplay = "FCIGCOVIC   25A"
fgDisplay.Col = 1: fgDisplay = "VILLE CLIENT"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOVIC
fgDisplay.Row = 40
fgDisplay.Col = 0: fgDisplay = "FCIGCOLPC   25A"
fgDisplay.Col = 1: fgDisplay = "LIBELL.PAYS CLIENT"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOLPC
fgDisplay.Row = 41
fgDisplay.Col = 0: fgDisplay = "FCIGCOCLB    7A"
fgDisplay.Col = 1: fgDisplay = "NUMERO CLIENT BENEFI"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOCLB
fgDisplay.Row = 42
fgDisplay.Col = 0: fgDisplay = "FCIGCOLTB   30A"
fgDisplay.Col = 1: fgDisplay = "LIBEL ETAT CLI BENEF"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOLTB
fgDisplay.Row = 43
fgDisplay.Col = 0: fgDisplay = "FCIGCOBNR   32A"
fgDisplay.Col = 1: fgDisplay = "NOM/RAIS CLI BENENEF"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOBNR
fgDisplay.Row = 44
fgDisplay.Col = 0: fgDisplay = "FCIGCOBPR   32A"
fgDisplay.Col = 1: fgDisplay = "PRENOM/RAIS CLI BENE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOBPR
fgDisplay.Row = 45
fgDisplay.Col = 0: fgDisplay = "FCIGCOA1B   32A"
fgDisplay.Col = 1: fgDisplay = "ADRESSE 1 BENEFICIAC"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOA1B
fgDisplay.Row = 46
fgDisplay.Col = 0: fgDisplay = "FCIGCOA2B   32A"
fgDisplay.Col = 1: fgDisplay = "ADRESSE 2 BENEFICIAC"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOA2B
fgDisplay.Row = 47
fgDisplay.Col = 0: fgDisplay = "FCIGCOA3B   32A"
fgDisplay.Col = 1: fgDisplay = "ADRESSE 3 BENEFICIAC"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOA3B
fgDisplay.Row = 48
fgDisplay.Col = 0: fgDisplay = "FCIGCOCPB    6A"
fgDisplay.Col = 1: fgDisplay = "CODE POSTAL BENEFICI"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOCPB
fgDisplay.Row = 49
fgDisplay.Col = 0: fgDisplay = "FCIGCOBVI   25A"
fgDisplay.Col = 1: fgDisplay = "VILLE BENEFICIAIRE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOBVI
fgDisplay.Row = 50
fgDisplay.Col = 0: fgDisplay = "FCIGCOLPB   25A"
fgDisplay.Col = 1: fgDisplay = "LIBELL.PAYS BENEFIC"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOLPB
fgDisplay.Row = 51
fgDisplay.Col = 0: fgDisplay = "FCIGCOCLP    7A"
fgDisplay.Col = 1: fgDisplay = "N¢ CLIENT PORTEUR"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOCLP
fgDisplay.Row = 52
fgDisplay.Col = 0: fgDisplay = "FCIGCOLTP   30A"
fgDisplay.Col = 1: fgDisplay = "LIBELLE ETAT PORTEUR"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOLTP
fgDisplay.Row = 53
fgDisplay.Col = 0: fgDisplay = "FCIGCONRP   32A"
fgDisplay.Col = 1: fgDisplay = "NOM/RAISON PORTEUR"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCONRP
fgDisplay.Row = 54
fgDisplay.Col = 0: fgDisplay = "FCIGCOPRP   32A"
fgDisplay.Col = 1: fgDisplay = "PRENOM/RAISON PORTEU"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOPRP
fgDisplay.Row = 55
fgDisplay.Col = 0: fgDisplay = "FCIGCO1DP   32A"
fgDisplay.Col = 1: fgDisplay = "ADRESSE 1 PORTEUR"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCO1DP
fgDisplay.Row = 56
fgDisplay.Col = 0: fgDisplay = "FCIGCO2DP   32A"
fgDisplay.Col = 1: fgDisplay = "ADRESSE 2 PORTEUR"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCO2DP
fgDisplay.Row = 57
fgDisplay.Col = 0: fgDisplay = "FCIGCO3DP   32A"
fgDisplay.Col = 1: fgDisplay = "ADRESSE 3 PORTEUR"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCO3DP
fgDisplay.Row = 58
fgDisplay.Col = 0: fgDisplay = "FCIGCOCPP    6A"
fgDisplay.Col = 1: fgDisplay = "CODE POSTAL PORTEUR"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOCPP
fgDisplay.Row = 59
fgDisplay.Col = 0: fgDisplay = "FCIGCOVIP   25A"
fgDisplay.Col = 1: fgDisplay = "VILLE PORTEUR"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOVIP
fgDisplay.Row = 60
fgDisplay.Col = 0: fgDisplay = "FCIGCOLPP   25A"
fgDisplay.Col = 1: fgDisplay = "LIBELL.PAYS PORTEUR"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOLPP
fgDisplay.Row = 61
fgDisplay.Col = 0: fgDisplay = "FCIGCOCLT    7A"
fgDisplay.Col = 1: fgDisplay = "N¢ CLIENT TITULAIRE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOCLT
fgDisplay.Row = 62
fgDisplay.Col = 0: fgDisplay = "FCIGCOLIT   30A"
fgDisplay.Col = 1: fgDisplay = "LIBELLE ETAT TITULAI"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOLIT
fgDisplay.Row = 63
fgDisplay.Col = 0: fgDisplay = "FCIGCONOT   32A"
fgDisplay.Col = 1: fgDisplay = "NOM/RAISON TITULALAI"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCONOT
fgDisplay.Row = 64
fgDisplay.Col = 0: fgDisplay = "FCIGCOPRT   32A"
fgDisplay.Col = 1: fgDisplay = "PRENOM/RAISON TITULA"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOPRT
fgDisplay.Row = 65
fgDisplay.Col = 0: fgDisplay = "FCIGCO1DT   32A"
fgDisplay.Col = 1: fgDisplay = "ADRESSE 1 TITULAIRE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCO1DT
fgDisplay.Row = 66
fgDisplay.Col = 0: fgDisplay = "FCIGCO2DT   32A"
fgDisplay.Col = 1: fgDisplay = "ADRESSE 2 TITULAIRE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCO2DT
fgDisplay.Row = 67
fgDisplay.Col = 0: fgDisplay = "FCIGCO3DT   32A"
fgDisplay.Col = 1: fgDisplay = "ADRESSE 3 TITULAIRE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCO3DT
fgDisplay.Row = 68
fgDisplay.Col = 0: fgDisplay = "FCIGCOPOT    6A"
fgDisplay.Col = 1: fgDisplay = "CODE POSTAL TITULAIR"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOPOT
fgDisplay.Row = 69
fgDisplay.Col = 0: fgDisplay = "FCIGCOVIT   25A"
fgDisplay.Col = 1: fgDisplay = "VILLE TITULAIRE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOVIT
fgDisplay.Row = 70
fgDisplay.Col = 0: fgDisplay = "FCIGCOLPT   25A"
fgDisplay.Col = 1: fgDisplay = "LIBELL.PAYS TITULAIR"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOLPT
fgDisplay.Row = 71
fgDisplay.Col = 0: fgDisplay = "FCIGCOCLC    7A"
fgDisplay.Col = 1: fgDisplay = "N¢ CLIENT COTITULAIR"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOCLC
fgDisplay.Row = 72
fgDisplay.Col = 0: fgDisplay = "FCIGCOLIC   30A"
fgDisplay.Col = 1: fgDisplay = "LIBELLE ETAT COTITUL"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOLIC
fgDisplay.Row = 73
fgDisplay.Col = 0: fgDisplay = "FCIGCONOC   32A"
fgDisplay.Col = 1: fgDisplay = "NOM/RAISON COTITULAL"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCONOC
fgDisplay.Row = 74
fgDisplay.Col = 0: fgDisplay = "FCIGCOPCO   32A"
fgDisplay.Col = 1: fgDisplay = "PRENOM/RAISON COTITU"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOPCO
fgDisplay.Row = 75
fgDisplay.Col = 0: fgDisplay = "FCIGCO1DC   32A"
fgDisplay.Col = 1: fgDisplay = "ADRESSE 1 COTITULAIR"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCO1DC
fgDisplay.Row = 76
fgDisplay.Col = 0: fgDisplay = "FCIGCO2DC   32A"
fgDisplay.Col = 1: fgDisplay = "ADRESSE 2 COTITULAIR"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCO2DC
fgDisplay.Row = 77
fgDisplay.Col = 0: fgDisplay = "FCIGCO3DC   32A"
fgDisplay.Col = 1: fgDisplay = "ADRESSE 3 COTITULAIR"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCO3DC
fgDisplay.Row = 78
fgDisplay.Col = 0: fgDisplay = "FCIGCOPOC    6A"
fgDisplay.Col = 1: fgDisplay = "CODE POSTAL COTITULA"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOPOC
fgDisplay.Row = 79
fgDisplay.Col = 0: fgDisplay = "FCIGCOVLC   25A"
fgDisplay.Col = 1: fgDisplay = "VILLE COTITULAIRE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOVLC
fgDisplay.Row = 80
fgDisplay.Col = 0: fgDisplay = "FCIGCOPAC   25A"
fgDisplay.Col = 1: fgDisplay = "LIBELL.PAYS COTITULA"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOPAC
fgDisplay.Row = 81
fgDisplay.Col = 0: fgDisplay = "FCIGCOCLM    7A"
fgDisplay.Col = 1: fgDisplay = "N¢ CLIENT MANDATAIRE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOCLM
fgDisplay.Row = 82
fgDisplay.Col = 0: fgDisplay = "FCIGCOLIM   30A"
fgDisplay.Col = 1: fgDisplay = "LIBELLE ETAT MANDATA"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOLIM
fgDisplay.Row = 83
fgDisplay.Col = 0: fgDisplay = "FCIGCONOM   32A"
fgDisplay.Col = 1: fgDisplay = "NOM/RAISON MANDATAIR"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCONOM
fgDisplay.Row = 84
fgDisplay.Col = 0: fgDisplay = "FCIGCOPRM   32A"
fgDisplay.Col = 1: fgDisplay = "PRENOM/RAISON MANDAT"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOPRM
fgDisplay.Row = 85
fgDisplay.Col = 0: fgDisplay = "FCIGCO1DM   32A"
fgDisplay.Col = 1: fgDisplay = "ADRESSE 1 MANDATAIRE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCO1DM
fgDisplay.Row = 86
fgDisplay.Col = 0: fgDisplay = "FCIGCO2DM   32A"
fgDisplay.Col = 1: fgDisplay = "ADRESSE 2 MANDATAIRE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCO2DM
fgDisplay.Row = 87
fgDisplay.Col = 0: fgDisplay = "FCIGCO3DM   32A"
fgDisplay.Col = 1: fgDisplay = "ADRESSE 3 MANDATAIRE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCO3DM
fgDisplay.Row = 88
fgDisplay.Col = 0: fgDisplay = "FCIGCOCPM    6A"
fgDisplay.Col = 1: fgDisplay = "CODE POSTAL MANDATAI"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOCPM
fgDisplay.Row = 89
fgDisplay.Col = 0: fgDisplay = "FCIGCOVIM   25A"
fgDisplay.Col = 1: fgDisplay = "VILLE MANDATAIRE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOVIM
fgDisplay.Row = 90
fgDisplay.Col = 0: fgDisplay = "FCIGCOLPM   25A"
fgDisplay.Col = 1: fgDisplay = "LIBELL.PAYS MANDATAI"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOLPM
fgDisplay.Row = 91
fgDisplay.Col = 0: fgDisplay = "FCIGCOCLG    7A"
fgDisplay.Col = 1: fgDisplay = "N¢ CLIENT GREFFE TRI"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOCLG
fgDisplay.Row = 92
fgDisplay.Col = 0: fgDisplay = "FCIGCOLIG   30A"
fgDisplay.Col = 1: fgDisplay = "LIBELLE ETAT GREFFE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOLIG
fgDisplay.Row = 93
fgDisplay.Col = 0: fgDisplay = "FCIGCONOG   32A"
fgDisplay.Col = 1: fgDisplay = "NOM/RAISON GREFFE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCONOG
fgDisplay.Row = 94
fgDisplay.Col = 0: fgDisplay = "FCIGCOPRG   32A"
fgDisplay.Col = 1: fgDisplay = "PRENOM/RAISON GREFFE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOPRG
fgDisplay.Row = 95
fgDisplay.Col = 0: fgDisplay = "FCIGCO1DG   32A"
fgDisplay.Col = 1: fgDisplay = "ADRESSE 1 GREFFE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCO1DG
fgDisplay.Row = 96
fgDisplay.Col = 0: fgDisplay = "FCIGCO2DG   32A"
fgDisplay.Col = 1: fgDisplay = "ADRESSE 2 GREFFE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCO2DG
fgDisplay.Row = 97
fgDisplay.Col = 0: fgDisplay = "FCIGCO3DG   32A"
fgDisplay.Col = 1: fgDisplay = "ADRESSE 3 GREFFE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCO3DG
fgDisplay.Row = 98
fgDisplay.Col = 0: fgDisplay = "FCIGCOCPG    6A"
fgDisplay.Col = 1: fgDisplay = "CODE POSTAL GREFFE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOCPG
fgDisplay.Row = 99
fgDisplay.Col = 0: fgDisplay = "FCIGCOVIG   25A"
fgDisplay.Col = 1: fgDisplay = "VILLE GREFFE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOVIG
fgDisplay.Row = 100
fgDisplay.Col = 0: fgDisplay = "FCIGCOLPG   25A"
fgDisplay.Col = 1: fgDisplay = "LIBELL.PAYS GREFFE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOLPG
fgDisplay.Row = 101
fgDisplay.Col = 0: fgDisplay = "FCIGCOLED   30A"
fgDisplay.Col = 1: fgDisplay = "LIEU EDITION"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOLED
fgDisplay.Row = 102
fgDisplay.Col = 0: fgDisplay = "FCIGCOGES   32A"
fgDisplay.Col = 1: fgDisplay = "NOM GESTIONNAIRE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOGES
fgDisplay.Row = 103
fgDisplay.Col = 0: fgDisplay = "FCIGCOREL   32A"
fgDisplay.Col = 1: fgDisplay = "REFERENCE LIBRE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOREL
fgDisplay.Row = 104
fgDisplay.Col = 0: fgDisplay = "FCIGCOTEL   20A"
fgDisplay.Col = 1: fgDisplay = "TELEPHONE GESTIONNAI"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOTEL
fgDisplay.Row = 105
fgDisplay.Col = 0: fgDisplay = "FCIGCOREJ    6A"
fgDisplay.Col = 1: fgDisplay = "CODE REJET"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOREJ
fgDisplay.Row = 106
fgDisplay.Col = 0: fgDisplay = "FCIGCOLIR   30A"
fgDisplay.Col = 1: fgDisplay = "LIBELLE REJET"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOLIR
fgDisplay.Row = 107
fgDisplay.Col = 0: fgDisplay = "FCIGCOMCH   20A"
fgDisplay.Col = 1: fgDisplay = "REJETE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOMCH
fgDisplay.Row = 108
fgDisplay.Col = 0: fgDisplay = "FCIGCODEV    3A"
fgDisplay.Col = 1: fgDisplay = "DEVISE MONTANT"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCODEV
fgDisplay.Row = 109
fgDisplay.Col = 0: fgDisplay = "FCIGCOAT1    1A"
fgDisplay.Col = 1: fgDisplay = "CAS N¢1 ATTESTION"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOAT1
fgDisplay.Row = 110
fgDisplay.Col = 0: fgDisplay = "FCIGCOAT2    1A"
fgDisplay.Col = 1: fgDisplay = "CAS N¢2 ATTESTION"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOAT2
fgDisplay.Row = 111
fgDisplay.Col = 0: fgDisplay = "FCIGCOAT3    1A"
fgDisplay.Col = 1: fgDisplay = "CAS N¢3 ATTESTION"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOAT3
fgDisplay.Row = 112
fgDisplay.Col = 0: fgDisplay = "FCIGCOAT4    1A"
fgDisplay.Col = 1: fgDisplay = "NON UTILISE PREVISIONN"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOAT4
fgDisplay.Row = 113
fgDisplay.Col = 0: fgDisplay = "FCIGCOIJ1    1A"
fgDisplay.Col = 1: fgDisplay = "CAS N¢1 INJONCTION"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOIJ1
fgDisplay.Row = 114
fgDisplay.Col = 0: fgDisplay = "FCIGCOIJ2    1A"
fgDisplay.Col = 1: fgDisplay = "CAS N¢2 INJONCTION"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOIJ2
fgDisplay.Row = 115
fgDisplay.Col = 0: fgDisplay = "FCIGCOIJ3    1A"
fgDisplay.Col = 1: fgDisplay = "CAS N¢3 INJONCTION"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOIJ3
fgDisplay.Row = 116
fgDisplay.Col = 0: fgDisplay = "FCIGCOIJ4    1A"
fgDisplay.Col = 1: fgDisplay = "NON UTILISE PREVISIONN"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOIJ4
fgDisplay.Row = 117
fgDisplay.Col = 0: fgDisplay = "FCIGCOMSD   20A"
fgDisplay.Col = 1: fgDisplay = "MONTANT SOLDE DISPON"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOMSD
fgDisplay.Row = 118
fgDisplay.Col = 0: fgDisplay = "FCIGCODES    3A"
fgDisplay.Col = 1: fgDisplay = "DEVISE SOLDE DISPONI"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCODES
fgDisplay.Row = 119
fgDisplay.Col = 0: fgDisplay = "FCIGCODEB   12A"
fgDisplay.Col = 1: fgDisplay = "MENTION DEBITEUR"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCODEB
fgDisplay.Row = 120
fgDisplay.Col = 0: fgDisplay = "FCIGCOIC1    1A"
fgDisplay.Col = 1: fgDisplay = "CAS N¢1 PAS PAYE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOIC1
fgDisplay.Row = 121
fgDisplay.Col = 0: fgDisplay = "FCIGCOIC2    1A"
fgDisplay.Col = 1: fgDisplay = "CAS N¢2 PAYE PARTIEL"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOIC2
fgDisplay.Row = 122
fgDisplay.Col = 0: fgDisplay = "FCIGCOMPP   20A"
fgDisplay.Col = 1: fgDisplay = "MONTANT PAIEMT PARTI"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOMPP
fgDisplay.Row = 123
fgDisplay.Col = 0: fgDisplay = "FCIGCODPP    3A"
fgDisplay.Col = 1: fgDisplay = "DEVISE MT PAIT PARTI"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCODPP
fgDisplay.Row = 124
fgDisplay.Col = 0: fgDisplay = "FCIGCOCH1    7A"
fgDisplay.Col = 1: fgDisplay = "NUMERO CHEQUE 1"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOCH1
fgDisplay.Row = 125
fgDisplay.Col = 0: fgDisplay = "FCIGCOMC1   20A"
fgDisplay.Col = 1: fgDisplay = "MONTANT CHEQUE 1"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOMC1
fgDisplay.Row = 126
fgDisplay.Col = 0: fgDisplay = "FCIGCODE1    3A"
fgDisplay.Col = 1: fgDisplay = "DEVISE MONTANT 1"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCODE1
fgDisplay.Row = 127
fgDisplay.Col = 0: fgDisplay = "FCIGCOCH2    7A"
fgDisplay.Col = 1: fgDisplay = "NUMERO CHEQUE 2"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOCH2
fgDisplay.Row = 128
fgDisplay.Col = 0: fgDisplay = "FCIGCOMC2   20A"
fgDisplay.Col = 1: fgDisplay = "MONTANT CHEQUE 2"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOMC2
fgDisplay.Row = 129
fgDisplay.Col = 0: fgDisplay = "FCIGCODE2    3A"
fgDisplay.Col = 1: fgDisplay = "DEVISE MONTANT 2"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCODE2
fgDisplay.Row = 130
fgDisplay.Col = 0: fgDisplay = "FCIGCOCH3    7A"
fgDisplay.Col = 1: fgDisplay = "NUMERO CHEQUE 3"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOCH3
fgDisplay.Row = 131
fgDisplay.Col = 0: fgDisplay = "FCIGCOMC3   20A"
fgDisplay.Col = 1: fgDisplay = "MONTANT CHEQUE 3"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOMC3
fgDisplay.Row = 132
fgDisplay.Col = 0: fgDisplay = "FCIGCODE3    3A"
fgDisplay.Col = 1: fgDisplay = "DEVISE MONTANT 3"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCODE3
fgDisplay.Row = 133
fgDisplay.Col = 0: fgDisplay = "FCIGCOCH4    7A"
fgDisplay.Col = 1: fgDisplay = "NUMERO CHEQUE 4"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOCH4
fgDisplay.Row = 134
fgDisplay.Col = 0: fgDisplay = "FCIGCOMC4   20A"
fgDisplay.Col = 1: fgDisplay = "MONTANT CHEQUE 4"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOMC4
fgDisplay.Row = 135
fgDisplay.Col = 0: fgDisplay = "FCIGCODE4    3A"
fgDisplay.Col = 1: fgDisplay = "DEVISE MONTANT 4"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCODE4
fgDisplay.Row = 136
fgDisplay.Col = 0: fgDisplay = "FCIGCOCH5    7A"
fgDisplay.Col = 1: fgDisplay = "NUMERO CHEQUE 5"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOCH5
fgDisplay.Row = 137
fgDisplay.Col = 0: fgDisplay = "FCIGCOMC5   20A"
fgDisplay.Col = 1: fgDisplay = "MONTANT CHEQUE 5"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOMC5
fgDisplay.Row = 138
fgDisplay.Col = 0: fgDisplay = "FCIGCODE5    3A"
fgDisplay.Col = 1: fgDisplay = "DEVISE MONTANT 5"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCODE5
fgDisplay.Row = 139
fgDisplay.Col = 0: fgDisplay = "FCIGCOCH6    7A"
fgDisplay.Col = 1: fgDisplay = "NON UTILISE PREVISIO"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOCH6
fgDisplay.Row = 140
fgDisplay.Col = 0: fgDisplay = "FCIGCOMC6   20A"
fgDisplay.Col = 1: fgDisplay = "NON UTILISE PREVISIO"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOMC6
fgDisplay.Row = 141
fgDisplay.Col = 0: fgDisplay = "FCIGCODE6    3A"
fgDisplay.Col = 1: fgDisplay = "NON UTILISE PREVISIO"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCODE6
fgDisplay.Row = 142
fgDisplay.Col = 0: fgDisplay = "FCIGCODRJ   10A"
fgDisplay.Col = 1: fgDisplay = "DATE REJET DES CHQ"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCODRJ
fgDisplay.Row = 143
fgDisplay.Col = 0: fgDisplay = "FCIGCODEI   10A"
fgDisplay.Col = 1: fgDisplay = "DATE DEPART INTERDIT"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCODEI
fgDisplay.Row = 144
fgDisplay.Col = 0: fgDisplay = "FCIGCODLR   10A"
fgDisplay.Col = 1: fgDisplay = "DATE LIMITE REGULARI"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCODLR
fgDisplay.Row = 145
fgDisplay.Col = 0: fgDisplay = "FCIGCOMPN   20A"
fgDisplay.Col = 1: fgDisplay = "MONTANT PENALITE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOMPN
fgDisplay.Row = 146
fgDisplay.Col = 0: fgDisplay = "FCIGCODPN    3A"
fgDisplay.Col = 1: fgDisplay = "DEVISE MT PENALITE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCODPN
fgDisplay.Row = 147
fgDisplay.Col = 0: fgDisplay = "FCIGCOJ21    1A"
fgDisplay.Col = 1: fgDisplay = "CAS N1 INJ 2 EV PREC"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOJ21
fgDisplay.Row = 148
fgDisplay.Col = 0: fgDisplay = "FCIGCOJ22    1A"
fgDisplay.Col = 1: fgDisplay = "CAS N2 INJ 2 EV PREC"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOJ22
fgDisplay.Row = 149
fgDisplay.Col = 0: fgDisplay = "FCIGCOINT   30A"
fgDisplay.Col = 1: fgDisplay = "INTITULE COMPTE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOINT
fgDisplay.Row = 150
fgDisplay.Col = 0: fgDisplay = "FCIGCONAG   32A"
fgDisplay.Col = 1: fgDisplay = "NOM AGENCE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCONAG
fgDisplay.Row = 151
fgDisplay.Col = 0: fgDisplay = "FCIGCODAP   10A"
fgDisplay.Col = 1: fgDisplay = "DU CHEQUE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCODAP
fgDisplay.Row = 152
fgDisplay.Col = 0: fgDisplay = "FCIGCOMIM   20A"
fgDisplay.Col = 1: fgDisplay = "MONTANT IMPAYE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOMIM
fgDisplay.Row = 153
fgDisplay.Col = 0: fgDisplay = "FCIGCODIM    3A"
fgDisplay.Col = 1: fgDisplay = "DEVISE MT IMPAYE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCODIM
fgDisplay.Row = 154
fgDisplay.Col = 0: fgDisplay = "FCIGCOCHB    7A"
fgDisplay.Col = 1: fgDisplay = "FRAIS PUBLICITE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOCHB
fgDisplay.Row = 155
fgDisplay.Col = 0: fgDisplay = "FCIGCOMCB   20A"
fgDisplay.Col = 1: fgDisplay = "MONTANT CHQ BANQUE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOMCB
fgDisplay.Row = 156
fgDisplay.Col = 0: fgDisplay = "FCIGCODCH    3A"
fgDisplay.Col = 1: fgDisplay = "DEVISE MT CHQ BANQUE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCODCH
fgDisplay.Row = 157
fgDisplay.Col = 0: fgDisplay = "FCIGCONBQ   32A"
fgDisplay.Col = 1: fgDisplay = "NOM BANQUE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCONBQ
fgDisplay.Row = 158
fgDisplay.Col = 0: fgDisplay = "FCIGCOMTA   20A"
fgDisplay.Col = 1: fgDisplay = "MONTANT ABUSIF"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOMTA
fgDisplay.Row = 159
fgDisplay.Col = 0: fgDisplay = "FCIGCODEA    3A"
fgDisplay.Col = 1: fgDisplay = "DEVISE MONT. ABUSIF"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCODEA
fgDisplay.Row = 160
fgDisplay.Col = 0: fgDisplay = "FCIGCOLCA   32A"
fgDisplay.Col = 1: fgDisplay = "LIBELLE TYPE CARTE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOLCA
fgDisplay.Row = 161
fgDisplay.Col = 0: fgDisplay = "FCIGCOLNA   32A"
fgDisplay.Col = 1: fgDisplay = "LIBELLE NATURE CARTE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOLNA
fgDisplay.Row = 162
fgDisplay.Col = 0: fgDisplay = "FCIGCOVAL   10A"
fgDisplay.Col = 1: fgDisplay = "DATE VALIDITE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOVAL
fgDisplay.Row = 163
fgDisplay.Col = 0: fgDisplay = "FCIGCODUR    2A"
fgDisplay.Col = 1: fgDisplay = "DUREE VALIDITE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCODUR
fgDisplay.Row = 164
fgDisplay.Col = 0: fgDisplay = "FCIGCOLI1   32A"
fgDisplay.Col = 1: fgDisplay = "CHAMP LIBRE 1"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOLI1
fgDisplay.Row = 165
fgDisplay.Col = 0: fgDisplay = "FCIGCOLI2   32A"
fgDisplay.Col = 1: fgDisplay = "CHAMP LIBRE 2"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOLI2
fgDisplay.Row = 166
fgDisplay.Col = 0: fgDisplay = "FCIGCOPUP   10A"
fgDisplay.Col = 1: fgDisplay = "DATE PURGE POSSIBLE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOPUP
fgDisplay.Row = 167
fgDisplay.Col = 0: fgDisplay = "FCIGCOTYI   30A"
fgDisplay.Col = 1: fgDisplay = "TYPE INTERDIT"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOTYI
fgDisplay.Row = 168
fgDisplay.Col = 0: fgDisplay = "FCIGCODDI   10A"
fgDisplay.Col = 1: fgDisplay = "TION INTERNE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCODDI
fgDisplay.Row = 169
fgDisplay.Col = 0: fgDisplay = "FCIGCODFI   10A"
fgDisplay.Col = 1: fgDisplay = "ON INTERNE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCODFI
fgDisplay.Row = 170
fgDisplay.Col = 0: fgDisplay = "FCIGCODDB   10A"
fgDisplay.Col = 1: fgDisplay = "TION BANCAIRE"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCODDB
fgDisplay.Row = 171
fgDisplay.Col = 0: fgDisplay = "FCIGCODFB   10A"
fgDisplay.Col = 1: fgDisplay = "DATE FIN INTERDICTI-"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCODFB
fgDisplay.Row = 172
fgDisplay.Col = 0: fgDisplay = "FCIGCOACP   20A"
fgDisplay.Col = 1: fgDisplay = "COMPTE AV CONVERSION"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOACP
fgDisplay.Row = 173
fgDisplay.Col = 0: fgDisplay = "FCIGCOIBA   20A"
fgDisplay.Col = 1: fgDisplay = "NUMERO IBAN"
fgDisplay.Col = 2: fgDisplay = rsZFCIGCO0.FCIGCOIBA
fgDisplay.TopRow = 1
End Sub
