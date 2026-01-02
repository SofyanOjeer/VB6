Attribute VB_Name = "srvSAB_Export_CSV"
Option Explicit

Public Sub srvYCGSMM10_Export_CSV(lIdFile_Source As Integer, lIdFile_Destination As Integer, loptSelect_CSV_Header As Boolean, lnb As Long)
Dim xIn As String
If loptSelect_CSV_Header Then
    Print #2, "CGSMM1ETA;CGSMM1AGE;CGSMM1SER;CGSMM1SES;CGSMM1OPE;CGSMM1NAT;CGSMM1NUM;CGSMM1MON;CGSMM1NBR;CGSMM1DEV;CGSMM1CLI;CGSMM1COM;CGSMM1ENG;CGSMM1DEB;CGSMM1FIN;CGSMM1DUR;CGSMM1TYP;CGSMM1AUT;CGSMM1CVL;CGSMM1NLO;"
    Print #2, "ETABLISSEMENT;AGENCE;SERVICE;SOUS SERVICE;OPERATION;NATURE;NUMERO;NOMINAL;NOMBRE OPE.;DEVISE;TYPE CLI/CLIENT;COMPTE;DATE ENGAGEMENT;DATE DEBUT;DATE FIN;DUREE PREAVIS;TYPE DE PREAVIS;CODE AUTORISAT.;NOMINAL CONTREV.;NOMBRE DE LOT;"
    Print #2, ";;;;;;;;;;;;;;;;;;;;"
End If
Do Until EOF(lIdFile_Source)
      Line Input #lIdFile_Source, xIn
      lnb = lnb + 1
      Print #lIdFile_Destination, mId$(xIn, 1, 5) & ";" _
      & mId$(xIn, 6, 5) & ";" _
      & mId$(xIn, 11, 2) & ";" _
      & mId$(xIn, 13, 2) & ";" _
      & mId$(xIn, 15, 6) & ";" _
      & mId$(xIn, 21, 6) & ";" _
      & mId$(xIn, 27, 10) & ";" _
      & mId$(xIn, 37, 18) & ";" _
      & mId$(xIn, 55, 7) & ";" _
      & mId$(xIn, 62, 3) & ";" _
      & mId$(xIn, 65, 8) & ";" _
      & mId$(xIn, 73, 20) & ";" _
      & mId$(xIn, 93, 8) & ";" _
      & mId$(xIn, 101, 8) & ";" _
      & mId$(xIn, 109, 8) & ";" _
      & mId$(xIn, 117, 4) & ";" _
      & mId$(xIn, 121, 1) & ";" _
      & mId$(xIn, 122, 3) & ";" _
      & mId$(xIn, 125, 18) & ";" _
      & mId$(xIn, 143, 7) & ";"
Loop
End Sub
Public Sub srvYCGSCOM0_Export_CSV(lIdFile_Source As Integer, lIdFile_Destination As Integer, loptSelect_CSV_Header As Boolean, lnb As Long)
Dim xIn As String
If loptSelect_CSV_Header Then
    Print #2, "CGSCOMETA;CGSCOMCLI;CGSCOMPLA;CGSCOMCOM;CGSCOMDAD;CGSCOMOPE;CGSCOMANA;CGSCOMDAF;CGSCOMMON;CGSCOMNOP;CGSCOMDEV;CGSCOMBAS;CGSCOMNCD;CGSCOMNCC;"
    Print #2, "ETABLISSEMENT;NUMERO CLIENT;NUMERO DE PLAN;NUMERO DE COMPTE;DATE DE DEBUT;CODE OPERATION;CODE ANALYTIQUE;DATE DE FIN;MONTANT TOTAL;NOMBRE COMMISS.;DEVISE DU COMPTE;MT TOTAL EN BASE;N° LIGNE COMM. DB;N° LIGNE COMM. CR;"
    Print #2, ";;;;;;;;;;;;;;"
End If
Do Until EOF(lIdFile_Source)
      Line Input #lIdFile_Source, xIn
      lnb = lnb + 1
      Print #lIdFile_Destination, mId$(xIn, 1, 5) & ";" _
      & mId$(xIn, 6, 7) & ";" _
      & mId$(xIn, 13, 4) & ";" _
      & mId$(xIn, 17, 20) & ";" _
      & mId$(xIn, 37, 8) & ";" _
      & mId$(xIn, 45, 6) & ";" _
      & mId$(xIn, 51, 6) & ";" _
      & mId$(xIn, 57, 8) & ";" _
      & mId$(xIn, 65, 16) & ";" _
      & mId$(xIn, 81, 6) & ";" _
      & mId$(xIn, 87, 3) & ";" _
      & mId$(xIn, 90, 16) & ";" _
      & mId$(xIn, 106, 4) & ";" _
      & mId$(xIn, 110, 4) & ";"
Loop
End Sub
Public Sub srvYCGSMOY0_Export_CSV(lIdFile_Source As Integer, lIdFile_Destination As Integer, loptSelect_CSV_Header As Boolean, lnb As Long)
Dim xIn As String
If loptSelect_CSV_Header Then
    Print #2, "CGSMOYETA;CGSMOYCOM;CGSMOYAMM;CGSMOYDAD;CGSMOYDAF;CGSMOYSM1;CGSMOYSM2;CGSMOYASS;CGSMOYMT1;CGSMOYMT2;"
    Print #2, "ETABLISSEMENT;NUMERO DE COMPTE;SAA MM;DATE DE DEBUT;DATE DE FIN;SOLDE MOYEN DB;SOLDE MOYEN CR;ASSIETTE  COMM;;;"
    Print #2, ";;;;;;;;;;"
End If
Do Until EOF(lIdFile_Source)
      Line Input #lIdFile_Source, xIn
      lnb = lnb + 1
      Print #lIdFile_Destination, mId$(xIn, 1, 5) & ";" _
      & mId$(xIn, 6, 20) & ";" _
      & mId$(xIn, 26, 6) & ";" _
      & mId$(xIn, 32, 8) & ";" _
      & mId$(xIn, 40, 8) & ";" _
      & mId$(xIn, 48, 16) & ";" _
      & mId$(xIn, 64, 16) & ";" _
      & mId$(xIn, 80, 16) & ";" _
      & mId$(xIn, 96, 16) & ";" _
      & mId$(xIn, 112, 16) & ";"
Loop
End Sub
Public Sub srvYCGSENC0_Export_CSV(lIdFile_Source As Integer, lIdFile_Destination As Integer, loptSelect_CSV_Header As Boolean, lnb As Long)
Dim xIn As String
If loptSelect_CSV_Header Then
    Print #2, "CGSENCETA;CGSENCTYP;CGSENCCLI;CGSENCPLA;CGSENCCOM;CGSENCDAD;CGSENCDAF;CGSENCCPT;CGSENCRUB;CGSENCSOL;CGSENCSM1;CGSENCSM2;CGSENCSDB;CGSENCJDB;CGSENCSCR;CGSENCJCR;CGSENCDEV;CGSENCIDB;CGSENCICR;CGSENCMDB;CGSENCMCR;CGSENCTDB;CGSENCTCR;CGSENCCOP;CGSENCBA1;CGSENCBA2;CGSENCID1;CGSENCIC1;CGSENCRET;CGSENCRES;CGSENCTD1;CGSENCTC1;CGSENCTD2;CGSENCTC2;CGSENCNLE;CGSENCNLR;CGSENCTXD;CGSENCTXC;CGSENCIMP;CSGENCTDE;CGSENCMDE;CGSENCTCD;CGSENCMCD;"
    Print #2, "ETABLISSEMENT;TYPE ENCOURS;N° CLIENT;NUMERO DE PLAN;NUMERO DE COMPTE;DATE DE DEBUT;DATE DE FIN;ECH. COMPTA O/N;RUBRIQUE COMPTAB;SOLDE FIN PERIO;SOLDE MOYEN DB;SOLDE MOYEN CR;SOLDE DB MAX;NBJ DE DEBIT;SOLDE CR MAX;NBJ DE CREDIT;DEVISE DE COMPTE;INT. TRESO DB;INT. TRESO CR;MARGE MT DB;MARGE MT CR;MARGE TAUX DB;MARGE TAUX CR;CODE PRODUIT;SLD MOY  DB BASE;SLD MOY  CR BASE;INT TRE. DB BASE;INT TRE. CR BASE;INTERETS RETRO;COUT DES RESERVE;INTERETS DB;INTERETS CR;INTERETS DB BASE;INTERETS CR BASE;N° LIGNE EMPLOIS;N° LIGNE RESSOUR;TAUX ANALYSE DB;TAUX ANALYSE CR;CPTE IMPUTATION;TAUX  IDE;MARGE IDE;TAUX  ICR;MARGE ICR;"
    Print #2, ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;"
End If
Do Until EOF(lIdFile_Source)
      Line Input #lIdFile_Source, xIn
      lnb = lnb + 1
      Print #lIdFile_Destination, mId$(xIn, 1, 5) & ";" & mId$(xIn, 6, 1) & ";" _
      & mId$(xIn, 7, 7) & ";" & mId$(xIn, 14, 4) & ";" _
      & mId$(xIn, 18, 20) & ";" & mId$(xIn, 38, 8) & ";" _
      & mId$(xIn, 46, 8) & ";" & mId$(xIn, 54, 1) & ";" _
      & mId$(xIn, 55, 10) & ";" & mId$(xIn, 65, 18) & ";" _
      & mId$(xIn, 83, 16) & ";" & mId$(xIn, 99, 16) & ";" _
      & mId$(xIn, 115, 16) & ";" & mId$(xIn, 131, 5) & ";" _
      & mId$(xIn, 136, 16) & ";" & mId$(xIn, 152, 5) & ";" _
      & mId$(xIn, 157, 3) & ";" & mId$(xIn, 160, 16) & ";" _
      & mId$(xIn, 176, 16) & ";" & mId$(xIn, 192, 16) & ";" _
      & mId$(xIn, 208, 16) & ";" & mId$(xIn, 224, 15) & ";" _
      & mId$(xIn, 239, 15) & ";" & mId$(xIn, 254, 3) & ";" _
      & mId$(xIn, 257, 16) & ";" & mId$(xIn, 273, 16) & ";" _
      & mId$(xIn, 289, 16) & ";" & mId$(xIn, 305, 16) & ";" _
      & mId$(xIn, 321, 16) & ";" & mId$(xIn, 337, 16) & ";" _
      & mId$(xIn, 353, 16) & ";" & mId$(xIn, 369, 16) & ";" _
      & mId$(xIn, 385, 16) & ";" & mId$(xIn, 401, 16) & ";" _
      & mId$(xIn, 417, 4) & ";" & mId$(xIn, 421, 4) & ";" _
      & mId$(xIn, 425, 15) & ";" & mId$(xIn, 440, 15) & ";" _
      & mId$(xIn, 455, 20) & ";" & mId$(xIn, 475, 6) & ";" _
      & mId$(xIn, 481, 10) & ";" & mId$(xIn, 491, 6) & ";" _
      & mId$(xIn, 497, 10) & ";"
Loop
End Sub

Public Sub srvYCGSMM30_Export_CSV(lIdFile_Source As Integer, lIdFile_Destination As Integer, loptSelect_CSV_Header As Boolean, lnb As Long)
Dim xIn As String
If loptSelect_CSV_Header Then
    Print #2, "CGSMM3ETA;CGSMM3AGE;CGSMM3SER;CGSMM3SES;CGSMM3OPE;CGSMM3NAT;CGSMM3NUM;CGSMM3SEN;CGSMM3SEQ;CGSMM3DEV;CGSMM3REF;CGSMM3APP;CGSMM3TAU;CGSMM3MAR;CGSMM3MRC;CGSMM3DVA;CGSMM3DTR;CGSMM3DRG;CGSMM3INT;CGSMM3COU;CGSMM3DEB;CGSMM3FIN;CGSMM3ASS;CGSMM3NBJ;CGSMM3NBP;CGSMM3BAS;CGSMM3MAC;CGSMM3MIN;CGSMM3TXA;"
    Print #2, "ETABLISSEMENT;AGENCE;SERVICE;SOUS SERVICE;OPERATION;NATURE;NUMERO;SENS;N° SEQUENCE;DEVISE;CODE TAUX;CODE APPLICAT°;TAUX FIXE;MARGE CLIENT;MARGE COMMERC.;DATE VAL CLIENT;DATE VAL TRESO;DATE REGLEMENT;INTERETS DS MOIS;INTERETS COURUS;DATE DEBUT PERIO;DATE FIN PERIODE;MONTANT ASSIETTE;NB JOUR OPE MOIS;NB JOUR PERIODE;BASE DEVISE;MONT. MARGE COM.;MONT. INTS.TRESO;TAUX D ANALYSE;"
    Print #2, ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;"
End If
Do Until EOF(lIdFile_Source)
      Line Input #lIdFile_Source, xIn
      lnb = lnb + 1
      Print #lIdFile_Destination, mId$(xIn, 1, 5) & ";" _
      & mId$(xIn, 6, 5) & ";" & mId$(xIn, 11, 2) & ";" _
      & mId$(xIn, 13, 2) & ";" & mId$(xIn, 15, 6) & ";" _
      & mId$(xIn, 21, 6) & ";" & mId$(xIn, 27, 10) & ";" _
      & mId$(xIn, 37, 1) & ";" & mId$(xIn, 38, 6) & ";" _
      & mId$(xIn, 44, 3) & ";" & mId$(xIn, 47, 6) & ";" _
      & mId$(xIn, 53, 1) & ";" & mId$(xIn, 54, 15) & ";" _
      & mId$(xIn, 69, 15) & ";" & mId$(xIn, 84, 15) & ";" _
      & mId$(xIn, 99, 8) & ";" & mId$(xIn, 107, 8) & ";" _
      & mId$(xIn, 115, 8) & ";" & mId$(xIn, 123, 18) & ";" _
      & mId$(xIn, 141, 18) & ";" & mId$(xIn, 159, 8) & ";" _
      & mId$(xIn, 167, 8) & ";" & mId$(xIn, 175, 18) & ";" _
      & mId$(xIn, 193, 6) & ";" & mId$(xIn, 199, 6) & ";" _
      & mId$(xIn, 205, 4) & ";" & mId$(xIn, 209, 18) & ";" _
      & mId$(xIn, 227, 18) & ";" & mId$(xIn, 245, 15) & ";"
Loop
End Sub

Public Sub srvYCGSMM40_Export_CSV(lIdFile_Source As Integer, lIdFile_Destination As Integer, loptSelect_CSV_Header As Boolean, lnb As Long)
Dim xIn As String
If loptSelect_CSV_Header Then
    Print #2, "CGSMM4ETA;CGSMM4AGE;CGSMM4SER;CGSMM4SES;CGSMM4OPE;CGSMM4NAT;CGSMM4NUM;CGSMM4SEN;CGSMM4SEQ;CGSMM4DEV;CGSMM4COM;CGSMM4MON;"
    Print #2, "ETABLISSEMENT;AGENCE;SERVICE;SOUS SERVICE;OPERATION;NATURE;NUMERO;SENS;N° SEQUENCE;DEVISE;CODE COMMISS°;MONTANT COMMISS°;"
    Print #2, ";;;;;;;;;;;;"
End If
Do Until EOF(lIdFile_Source)
      Line Input #lIdFile_Source, xIn
      lnb = lnb + 1
      Print #lIdFile_Destination, mId$(xIn, 1, 5) & ";" _
      & mId$(xIn, 6, 5) & ";" _
      & mId$(xIn, 11, 2) & ";" _
      & mId$(xIn, 13, 2) & ";" _
      & mId$(xIn, 15, 6) & ";" _
      & mId$(xIn, 21, 6) & ";" _
      & mId$(xIn, 27, 10) & ";" _
      & mId$(xIn, 37, 1) & ";" _
      & mId$(xIn, 38, 6) & ";" _
      & mId$(xIn, 44, 3) & ";" _
      & mId$(xIn, 47, 6) & ";" _
      & mId$(xIn, 53, 18) & ";"
Loop
End Sub


Public Sub srvYMOUVEA0_Export_CSV()
Dim xIn As String
Open "C:\Temp\YMOUVEA0.txt" For Input As #1
Open "C:\Temp\YMOUVEA0.csv" For Output As #2
'Print #2, "MOUVEMETA;MOUVEMPLA;MOUVEMCOM;MOUVEMMON;MOUVEMDOP;MOUVEMDVA;MOUVEMDCO;MOUVEMDTR;MOUVEMPIE;MOUVEMECR;MOUVEMOPE;MOUVEMNUM;MOUVEMSCH;MOUVEMUTI;MOUVEMAGE;MOUVEMSER;MOUVEMSSE;MOUVEMEXO;MOUVEMANA;MOUVEMBDF;MOUVEMANU;MOUVEMRET;MOUVEMEVE;MOUVEMSAN;MOUVEMSAD;"
'Print #2, "ETABLISSEMENT;NUMERO PLAN;NUMERO COMPTE;MONTANT;DATE D'OPERATION;DATE DE VALEUR;DATE COMPTABLE;DATE DE TRAITEMENT;NUMERO DE PIECE;NUMERO D'ECRITURE;CODE OPERATION;NUMERO OPERATION;CODE SCHEMA;UTILISATEUR;AGENCE OPERATRICE;SERVICE OPERATEUR;S/SERVICE OPERATEUR;CODE EXONERATION;CODE ANALYTIQUE;CODE BANQUE DE FR.;CODE ANNULATION;MOUVEMENT RETRO;EVENEMENT;STRUCT ANALY-CODE;STRUCT ANALY-DONNEES;"
'Print #2, ";;;;;;;;;;;;;;;;;;;;;;;;;"
Do Until EOF(1)
      Line Input #1, xIn
      Print #2, mId$(xIn, 1, 5) & ";" & mId$(xIn, 6, 4) & ";" _
      & mId$(xIn, 10, 20) & ";" & mId$(xIn, 30, 18) & ";" _
      & mId$(xIn, 48, 8) & ";" & mId$(xIn, 56, 8) & ";" _
      & mId$(xIn, 64, 8) & ";" & mId$(xIn, 72, 8) & ";" _
      & mId$(xIn, 80, 10) & ";" & mId$(xIn, 90, 8) & ";" _
      & mId$(xIn, 98, 3) & ";" & mId$(xIn, 101, 10) & ";" _
      & mId$(xIn, 111, 5) & ";" & mId$(xIn, 116, 5) & ";" _
      & mId$(xIn, 121, 5) & ";" & mId$(xIn, 126, 2) & ";" _
      & mId$(xIn, 128, 2) & ";" & mId$(xIn, 130, 1) & ";" _
      & mId$(xIn, 131, 6) & ";" & mId$(xIn, 137, 3) & ";" _
      & mId$(xIn, 140, 1) & ";" & mId$(xIn, 141, 1) & ";" _
      & mId$(xIn, 142, 3) & ";" & mId$(xIn, 145, 6) & ";" _
      & mId$(xIn, 151, 80) & ";"
Loop
Close
End Sub

Public Sub srvYMNURUT0_Export_CSV()
Dim xIn As String
Open "C:\Temp\YMNURUT0.txt" For Input As #1
Open "C:\Temp\YMNURUT0.csv" For Output As #2
If frmSPLFJOB.chkAS400_Export_CSV = "1" Then
    Print #2, "MNURUTUTI;MNURUTNOM;MNURUTETB;MNURUTCUT;MNURUTLOG;"
    Print #2, "UTILISATEUR;NOM;ETAB. PAR DEFAUT;CODE INTERNE;ENTREE LOGICIEL;"
    Print #2, ";;;;;"
End If
Do Until EOF(1)
      Line Input #1, xIn
      Print #2, mId$(xIn, 1, 10) & ";" _
       & mId$(xIn, 11, 30) & ";" _
       & mId$(xIn, 41, 5) & ";" _
       & mId$(xIn, 46, 5) & ";" _
       & mId$(xIn, 51, 1)
Loop
Close
End Sub


Public Sub srvYREPCPT0_Export_CSV()
Dim xIn As String
Open "C:\Temp\YREPCPT0.txt" For Input As #1
Open "C:\Temp\YREPCPT0.csv" For Output As #2
'Print #2, "REPCPTETA;REPCPTPLA;REPCPTCOM;REPCPTRUB;REPCPTINT;REPCPTNBR;REPCPTPPAL;REPCPTAGE;REPCPTDEV;REPCPTOPR;REPCPTVAL;REPCPTCPT;REPCPTDTE;REPCPTRES;REPCPTDUR;REPCPTMG1;REPCPTMG2;REPCPTGAR;REPCPTGA1;REPCPTQUA;REPCPTGA2;REPCPTGA3;REPCPTFIL;REPCPTMO1;REPCPTMO2;REPCPTMO3;REPCPTMO4;REPCPTMO5;REPCPTFI2;REPCPTSV1;REPCPTSV2;REPCPTSV3;REPCPTSV4;REPCPTSV5;REPCPTMD1;REPCPTMD2;REPCPTMD3;REPCPTMD4;REPCPTMD5;REPCPTMC1;REPCPTMC2;REPCPTMC3;REPCPTMC4;REPCPTMC5;REPCPTSD1;REPCPTSD2;REPCPTSD3;REPCPTSD4;REPCPTSD5;REPCPTSC1;REPCPTSC2;REPCPTSC3;REPCPTSC4;REPCPTSC5;REPCPTSMD;REPCPTSMC;REPCPTML1;REPCPTML2;REPCPTOUV;REPCPTCLO;REPCPTMFD;REPCPTCFD;REPCPTMFC;REPCPTCFC;REPCPTMFS;REPCPTCFS;REPCPTSM1;REPCPTBA1;REPCPTSM2;REPCPTBA2;REPCPTSM3;REPCPTBA3;REPCPTMMC;REPCPTMMS;REPCPTTEG;"
'Print #2, "ETABLISSEMENT;NUMERO PLAN;NUMERO COMPTE;RUBRIQUE;INTITULE;NOMBRE DE TITULAIRES;TITULAIRE PRINCIPAL;AGENCE;DEVISE;SOLDE DATE OPERATION;SOLDE DATE VALEUR;SOLDE DATE COMPTABLE;DATE D'EXTRACTION;DUREE REST. A COURIR;DUREE INITIALE;GARANTIE DEVISE BASE;GARANTIE EN DEVISE;CODE GARANTIE;GARANT;CODE QUALITE GARANT;CODE RESIDENCE GARAN;NON UTILISE;VOIR PROG EXTRACT;MONTANT1;MONTANT2;MONTANT3;MONTANT4;MONTANT5;ATTRIBUTS INDUITS;SOLDE VALEUR 1;SOLDE VALEUR 2;SOLDE VALEUR 3;SOLDE VALEUR 4;SOLDE VALEUR 5;MOUVEMENTS  DB 1;MOUVEMENTS  DB 2;MOUVEMENTS  DB 3;MOUVEMENTS  DB 4;MOUVEMENTS  DB 5;MOUVEMENTS  CR 1;MOUVEMENTS  CR 2;MOUVEMENTS  CR 3;MOUVEMENTS  CR 4;MOUVEMENTS  CR 5;SOLDE MOYEN DB 1;SOLDE MOYEN DB 2;SOLDE MOYEN DB 3;SOLDE MOYEN DB 4;SOLDE MOYEN DB 5;SOLDE MOYEN CR 1;SOLDE MOYEN CR 2;SOLDE MOYEN CR 3;SOLDE MOYEN CR 4;SOLDE MOYEN CR 5;SOLDE MOYEN DEBIT;SOLDE MOYEN CREDIT;MONTANT LIBRE 1;MONTANT LIBRE 2;DATE OUVERTURE COM;DATE CLOTURE COMPT;MT FLUX INTERET DB;CTVL FLUX INT.  DB;MT FLUX INTERET CR;_"
'CTVL FLUX INT.  CR;MT FLUX GLOBAL INT;CTVL FLUX GLOB.INT;MT ENCOURS MOYENDB;CTVL ENCOUR MOY DB;MT ENCOURS MOYENCR;CTVL ENCOUR MOY CR;MT ENCOURS MOYENGB;CTVL ENCOUR MOY GB;MT ENCOURS CPM. GB;CTVL ENCOUR CPM.GB;TAUX NOMINAL;"
'Print #2, ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;"
Do Until EOF(1)
      Line Input #1, xIn
      Print #2, mId$(xIn, 1, 5) & ";" & mId$(xIn, 6, 4) & ";" & mId$(xIn, 10, 20) & ";" & mId$(xIn, 30, 10) & ";" _
      & mId$(xIn, 40, 32) & ";" & mId$(xIn, 72, 3) & ";" & mId$(xIn, 75, 7) & ";" & mId$(xIn, 82, 5) & ";" _
      & mId$(xIn, 87, 3) & ";" & mId$(xIn, 90, 18) & ";" & mId$(xIn, 108, 18) & ";" & mId$(xIn, 126, 18) & ";" _
      & mId$(xIn, 144, 8) & ";" & mId$(xIn, 152, 5) & ";" & mId$(xIn, 157, 5) & ";" & mId$(xIn, 162, 16) & ";" _
      & mId$(xIn, 178, 16) & ";" & mId$(xIn, 194, 6) & ";" & mId$(xIn, 200, 7) & ";" & mId$(xIn, 207, 3) & ";" _
      & mId$(xIn, 210, 3) & ";" & mId$(xIn, 213, 3) & ";" & mId$(xIn, 216, 100) & ";" & mId$(xIn, 316, 18) & ";" _
      & mId$(xIn, 334, 18) & ";" & mId$(xIn, 352, 18) & ";" & mId$(xIn, 370, 18) & ";" & mId$(xIn, 388, 18) & ";" _
      & mId$(xIn, 406, 65) & ";" & mId$(xIn, 471, 18) & ";" & mId$(xIn, 489, 18) & ";" & mId$(xIn, 507, 18) & ";" _
      & mId$(xIn, 525, 18) & ";" & mId$(xIn, 543, 18) & ";" & mId$(xIn, 561, 18) & ";" & mId$(xIn, 579, 18) & ";" _
      & mId$(xIn, 597, 18) & ";" & mId$(xIn, 615, 18) & ";" & mId$(xIn, 633, 18) & ";" & mId$(xIn, 651, 18) & ";" _
      & mId$(xIn, 669, 18) & ";" & mId$(xIn, 687, 18) & ";" & mId$(xIn, 705, 18) & ";" & mId$(xIn, 723, 18) & ";" _
      & mId$(xIn, 741, 18) & ";" & mId$(xIn, 759, 18) & ";" & mId$(xIn, 777, 18) & ";" & mId$(xIn, 795, 18) & ";" _
      & mId$(xIn, 813, 18) & ";" & mId$(xIn, 831, 18) & ";" & mId$(xIn, 849, 18) & ";" & mId$(xIn, 867, 18) & ";" _
      & mId$(xIn, 885, 18) & ";" & mId$(xIn, 903, 18) & ";" & mId$(xIn, 921, 18) & ";" & mId$(xIn, 939, 18) & ";" _
      & mId$(xIn, 957, 18) & ";" & mId$(xIn, 975, 18) & ";" & mId$(xIn, 993, 8) & ";" & mId$(xIn, 1001, 8) & ";" _
      & mId$(xIn, 1009, 16) & ";" & mId$(xIn, 1025, 16) & ";" & mId$(xIn, 1041, 16) & ";" & mId$(xIn, 1057, 16) & ";" _
      & mId$(xIn, 1073, 16) & ";" & mId$(xIn, 1089, 16) & ";" & mId$(xIn, 1105, 16) & ";" & mId$(xIn, 1121, 16) & ";" _
      & mId$(xIn, 1137, 16) & ";" & mId$(xIn, 1153, 16) & ";" & mId$(xIn, 1169, 16) & ";" & mId$(xIn, 1185, 16) & ";" _
      & mId$(xIn, 1201, 16) & ";" & mId$(xIn, 1217, 16) & ";" & mId$(xIn, 1233, 16) & ";"
Loop
Close
End Sub

Public Sub srvYBASTAU0_Export_CSV()
Dim xIn As String
Open "C:\Temp\YBASTAU0.txt" For Input As #1
Open "C:\Temp\YBASTAU0.csv" For Output As #2
If frmSPLFJOB.chkAS400_Export_CSV = "1" Then
    Print #2, "BASTAUETA;BASTAUDEV;BASTAUTAU;BASTAUABR;BASTAULIB;BASTAUTYP;BASTAUPER;BASTAUNBJ;BASTAUSIG;BASTAUTYJ;BASTAUCAL;BASTAUDEC;BASTAUTXC;BASTAUCAP;BASTAUNBP;BASTAUARR;BASTAUCOT;BASTAUDEB;BASTAUFIN;BASTAUDSU;BASTAUTSU;BASTAUTYA;BASTAUTPE;BASTAUFIL;"
    Print #2, "Etablissement;Code devise;Code taux;Abrégé;Libellé;Mode d'obtention;Périodicité;Délai d'usance;Sens du délai;Type de jours;Mode de calcul;Taux  pour;calcul;Code capitalisation;Nombre de périodes;Arrondi;Devise de cotation;Début de validité;Fin de validité;Taux  de;substitution;Type arrondi;Type de période;Filler;"
    Print #2, ";;;;;;;;;;;;;;;;;;;;;;;;"
End If
Do Until EOF(1)
      Line Input #1, xIn
      Print #2, mId$(xIn, 1, 5) & ";" _
      & mId$(xIn, 6, 3) & ";" _
      & mId$(xIn, 9, 6) & ";" _
      & mId$(xIn, 15, 12) & ";" _
      & mId$(xIn, 27, 30) & ";" _
      & mId$(xIn, 57, 1) & ";" _
      & mId$(xIn, 58, 1) & ";" _
      & mId$(xIn, 59, 4) & ";" _
      & mId$(xIn, 63, 1) & ";" _
      & mId$(xIn, 64, 1) & ";" _
      & mId$(xIn, 65, 1) & ";" _
      & mId$(xIn, 66, 3) & ";" _
      & mId$(xIn, 69, 6) & ";" _
      & mId$(xIn, 75, 3) & ";" _
      & mId$(xIn, 78, 5) & ";" _
      & mId$(xIn, 83, 2) & ";" _
      & mId$(xIn, 85, 3) & ";" _
      & mId$(xIn, 88, 8) & ";" _
      & mId$(xIn, 96, 8) & ";" _
      & mId$(xIn, 104, 3) & ";" _
      & mId$(xIn, 107, 6) & ";" _
      & mId$(xIn, 113, 1) & ";" _
      & mId$(xIn, 114, 1) & ";" _
      & mId$(xIn, 115, 100) & ";"
Loop
Close
End Sub


Public Sub srvYCHEOPP0_Export_CSV()
Dim xIn As String
Open "C:\Temp\YCHEOPP0.txt" For Input As #1
Open "C:\Temp\YCHEOPP0.csv" For Output As #2
If frmSPLFJOB.chkAS400_Export_CSV = "1" Then
    Print #2, "CHEOPPETA;CHEOPPCOM;CHEOPPTYP;CHEOPPCH2;CHEOPPCH1;CHEOPPDTO;CHEOPPMOT;CHEOPPLIO;CHEOPPAGE;CHEOPPDEV;CHEOPPCTA;CHEOPPDTL;CHEOPPSER;CHEOPPTRA;CHEOPPDT1;CHEOPPUS1;CHEOPPUOP;CHEOPPUML;"
    Print #2, "Etablissement;Compte;1/oppos. 2/passés;Chèque  à;Chèque  de;Date  opposition;Motif opposition;Libel.opposition;Agence compte;Devise compte;Code état oppo.;Date levée oppo.;Code serveur;Code traitement;Date validation;User validation;Ut. saisie oppo;Ut. saisie mlvée;"
    Print #2, ";;;;;;;;;;;;;;;;;;"
End If
Do Until EOF(1)
      Line Input #1, xIn
      Print #2, mId$(xIn, 1, 5) & ";" _
      & mId$(xIn, 6, 20) & ";" _
      & mId$(xIn, 26, 1) & ";" _
      & mId$(xIn, 27, 11) & ";" _
      & mId$(xIn, 38, 11) & ";" _
      & mId$(xIn, 49, 8) & ";" _
      & mId$(xIn, 57, 3) & ";" _
      & mId$(xIn, 60, 30) & ";" _
      & mId$(xIn, 90, 5) & ";" _
      & mId$(xIn, 95, 3) & ";" _
      & mId$(xIn, 98, 4) & ";" _
      & mId$(xIn, 102, 8) & ";" _
      & mId$(xIn, 110, 6) & ";" _
      & mId$(xIn, 116, 3) & ";" _
      & mId$(xIn, 119, 8) & ";" _
      & mId$(xIn, 127, 5) & ";" _
      & mId$(xIn, 132, 5) & ";" _
      & mId$(xIn, 137, 5) & ";"
Loop
Close
End Sub
Public Sub srvYCHQCOM0_Export_CSV()
Dim xIn As String
Open "C:\Temp\YCHQCOM0.txt" For Input As #1
Open "C:\Temp\YCHQCOM0.csv" For Output As #2
If frmSPLFJOB.chkAS400_Export_CSV = "1" Then
    Print #2, "CHQCOMETA;CHQCOMCOM;CHQCOMTYP;CHQCOMNOM;CHQCOMSEU;CHQCOMTAD;CHQCOMCOA;CHQCOMCPT;CHQCOMDAT;CHQCOMREN;CHQCOMLIV;CHQCOMCAL;CHQCOMAGE;CHQCOMGUI;CHQCOMDT1;CHQCOMUT1;CHQCOMVAL;CHQCOMDT2;CHQCOMGES;CHQCOMDEV;CHQCOMAG1;"
    Print #2, "ETABLISSEMENT;N° DE COMPTE;TYPE CHEQUIER;NBRE DE CHEQUIERS;SEUIL DE RENOUV.;ADR. CLIENT/CPTE;CODE ADRESSE;COMPTEUR;DATE DERNIERE MAJ;RENOUV. AUTO.;CODE LIVRAISON;CODE ADR LIVRAIS.;AGENCE DU COMPTE;GUICHET LIVRAISON;DATE SUPPRESSION;UTILISATEUR SUPPR;UTILISATEUR REINT;DATE REINTEGRAT.;VALIDATION OBLIG;DEVISE;AGCE DE LIVRAISON;"
    Print #2, ";;;;;;;;;;;;;;;;;;;;;"
End If
Do Until EOF(1)
      Line Input #1, xIn
      Print #2, mId$(xIn, 1, 5) & ";" _
      & mId$(xIn, 6, 20) & ";" _
      & mId$(xIn, 26, 6) & ";" _
      & mId$(xIn, 32, 4) & ";" _
      & mId$(xIn, 36, 4) & ";" _
      & mId$(xIn, 40, 1) & ";" _
      & mId$(xIn, 41, 2) & ";" _
      & mId$(xIn, 43, 4) & ";" _
      & mId$(xIn, 47, 8) & ";" _
      & mId$(xIn, 55, 1) & ";" _
      & mId$(xIn, 56, 3) & ";" _
      & mId$(xIn, 59, 2) & ";" _
      & mId$(xIn, 61, 5) & ";" _
      & mId$(xIn, 66, 5) & ";" _
      & mId$(xIn, 71, 8) & ";" _
      & mId$(xIn, 79, 5) & ";" _
      & mId$(xIn, 84, 5) & ";" _
      & mId$(xIn, 89, 8) & ";" _
      & mId$(xIn, 97, 1) & ";" _
      & mId$(xIn, 98, 3) & ";" _
      & mId$(xIn, 101, 5) & ";"
Loop
Close
End Sub




Public Sub srvYCHQHIS0_Export_CSV()
Dim xIn As String
Open "C:\Temp\YCHQHIS0.txt" For Input As #1
Open "C:\Temp\YCHQHIS0.csv" For Output As #2
If frmSPLFJOB.chkAS400_Export_CSV = "1" Then
    Print #2, "CHQHISETA;CHQHISAG1;CHQHISAGE;CHQHISCOM;CHQHISTYP;CHQHISDAT;CHQHISSEQ;CHQHISNPF;CHQHISNDF;CHQHISREC;CHQHISREM;CHQHISENV;CHQHISCAL;CHQHISAG2;CHQHISDSP;CHQHISUSP;CHQHISGUI;CHQHISNCH;CHQHISOBS;CHQHISDEV;CHQHISREN;CHQHISNRE;CHQHISDRE;CHQHISORI;CHQHISUSR;CHQHISSER;CHQHISTRA;"
    Print #2, "ETABLISSEMENT;AGENCE STOCK;AGENCE DU COMPTE;NUMERO DE COMPTE;TYPE CHEQUIER;DATE DEMANDE;NUM. DE SEQUENCE;NUM. 1ERE FORMULE;NUM. 2EME FORMULE;DATE RECEPTION;DATE REMISE;DATE ENVOI;CODE LIVRAISON;AGENCE COMMANDE;DATE SUPPRESSION;UTILI SUPPRESSION;GUICHET ACHEMINE.;NUM. CHEQUIER;TEXTE LIBRE;DEVISE;CODE ETAT RENOUV.;N° FORMULE DE REF;DATE RENOUV.;ORIGINE DU CHÉQU.;UTILI SUPP RENOUV;CODE SERVEUR;CODE TRAITEMENT;"
    Print #2, ";;;;;;;;;;;;;;;;;;;;;;;;;;;"
End If
Do Until EOF(1)
      Line Input #1, xIn
      Print #2, mId$(xIn, 1, 5) & ";" & mId$(xIn, 6, 5) & ";" & mId$(xIn, 11, 5) & ";" _
      & mId$(xIn, 16, 20) & ";" & mId$(xIn, 36, 6) & ";" _
      & mId$(xIn, 42, 8) & ";" & mId$(xIn, 50, 4) & ";" _
      & mId$(xIn, 54, 8) & ";" & mId$(xIn, 62, 8) & ";" _
      & mId$(xIn, 70, 8) & ";" & mId$(xIn, 78, 8) & ";" _
      & mId$(xIn, 86, 8) & ";" & mId$(xIn, 94, 3) & ";" _
      & mId$(xIn, 97, 5) & ";" & mId$(xIn, 102, 8) & ";" _
      & mId$(xIn, 110, 5) & ";" & mId$(xIn, 115, 5) & ";" _
      & mId$(xIn, 120, 8) & ";" & mId$(xIn, 128, 30) & ";" _
      & mId$(xIn, 158, 3) & ";" & mId$(xIn, 161, 2) & ";" _
      & mId$(xIn, 163, 8) & ";" & mId$(xIn, 171, 8) & ";" _
      & mId$(xIn, 179, 1) & ";" & mId$(xIn, 180, 5) & ";" _
      & mId$(xIn, 185, 6) & ";" & mId$(xIn, 191, 3) & ";"
Loop
Close
End Sub



Public Sub srvYSITLOT0_Export_CSV()
Dim xIn As String
Open "C:\Temp\YSITLOT0.txt" For Input As #1
Open "C:\Temp\YSITLOT0.csv" For Output As #2
If frmSPLFJOB.chkAS400_Export_CSV = "1" Then
    Print #2, "SITLOTETB;SITLOTAGE;SITLOTSER;SITLOTSSE;SITLOTOPE;SITLOTCPT;SITLOTLOT;SITLOTAVI;SITLOTCRE;SITLOTORD;SITLOTMON;SITLOTDAT;SITLOTETA;SITLOTPRO;SITLOTFUT;"
    Print #2, "Etablissement;Agence;Service;Sous-service;Code opération;Numéro de COMPTE;NUMERO LOT ATTRIBUE;AVIS GLOBAL O/N;CRE GLOBAL O/N;NomBRE ORDRES;MONTANT TOTAL;DATE EXECUTION;CODE ETAT;PROCHAINE ECHEANCE.;ZONE FUTURE;"
    Print #2, ";;;;;;;;;;;;;;;"
End If
Do Until EOF(1)
      Line Input #1, xIn
      Print #2, mId$(xIn, 1, 5) & ";" _
      & mId$(xIn, 6, 5) & ";" _
      & mId$(xIn, 11, 2) & ";" _
      & mId$(xIn, 13, 2) & ";" _
      & mId$(xIn, 15, 3) & ";" _
      & mId$(xIn, 18, 20) & ";" _
      & mId$(xIn, 38, 5) & ";" _
      & mId$(xIn, 43, 1) & ";" _
      & mId$(xIn, 44, 1) & ";" _
      & mId$(xIn, 45, 10) & ";" _
      & mId$(xIn, 55, 16) & ";" _
      & mId$(xIn, 71, 8) & ";" _
      & mId$(xIn, 79, 4) & ";" _
      & mId$(xIn, 83, 8) & ";" _
      & mId$(xIn, 91, 100) & ";"
Loop
Close
End Sub
Public Sub srvYSITNUM0_Export_CSV()
Dim xIn As String
Open "C:\Temp\YSITNUM0.txt" For Input As #1
Open "C:\Temp\YSITNUM0.csv" For Output As #2
If frmSPLFJOB.chkAS400_Export_CSV = "1" Then
    Print #2, "SITNUMETB;SITNUMAGE;SITNUMSER;SITNUMSSE;SITNUMOPE;SITNUMNUM;"
    Print #2, "Etablissement;Agence;Service;Sous-service;Code Opération;Numéro attribué;"
    Print #2, ";;;;;;"
End If
Do Until EOF(1)
      Line Input #1, xIn
      Print #2, mId$(xIn, 1, 5) & ";" _
      & mId$(xIn, 6, 5) & ";" _
      & mId$(xIn, 11, 2) & ";" _
      & mId$(xIn, 13, 2) & ";" _
      & mId$(xIn, 15, 3) & ";" _
      & mId$(xIn, 18, 10) & ";"
Loop
Close
End Sub




Public Sub srvYSITORD0_Export_CSV()
Dim xIn As String
Open "C:\Temp\YSITORD0.txt" For Input As #1
Open "C:\Temp\YSITORD0.csv" For Output As #2
If frmSPLFJOB.chkAS400_Export_CSV = "1" Then
    Print #2, "SITORDETB;SITORDAGE;SITORDSER;SITORDSSE;SITORDOPE;SITORDLOT;SITORDNUM;SITORDNAT;SITORDBQE;SITORDGUI;SITORDCPT;SITORDDES;SITORDCPI;SITORDMON;SITORDDPE;SITORDPRO;SITORDPR2;SITORDVAL;SITORDDPR;SITORDJOU;SITORDREP;SITORDZER;SITORDPER;SITORDDEN;SITORDNDO;SITORDNOM;SITORDLI1;SITORDLI2;SITORDAVI;SITORDCO1;SITORDCO2;SITORDMO1;SITORDMO2;SITORDTVA;SITORDDVA;SITORDRES;SITORDBDF;SITORDTYP;SITORDEME;SITORDAVD;SITORDAVL;SITORDCON;SITORDROU;SITORDETA;SITORDTOP;SITORDREF;SITORDECO;SITORDEXO;SITORDFIL;"
    Print #2, "Etablissement;Agence;Service;Sous-service;Code opération;Numéro de lot;Numéro Opération;Nature opération;Banque destinataire;Guichet destinataire;Compte destinataire;Nom destinataire;Compte destinataire;Montant;1ère échéance;prochaine échéance;proc éch.non reporté;Date valeur;Dernière échéance;Jour échéance;Type de report;Remise à 0 montant;Périodicité;Dernier envoi;Numéro D.O.;Nom D.O.;Libellé destinataire;Libellé destinataire;Avis client;Code commission;Code commission;Mouvement commission;Mouvement commission;Montant TVA;Date valeur commiss°;Viremnt non résident;Code pays BDF;Type virement;N° nat° émetteur;Avis pr destinataire;Avis global lot;Consigne retard;Code routage;Code etat;Top avis;Référence;Code Economique;Code Exonération;ZONE FUTURE;"
    Print #2, ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;"
End If
Do Until EOF(1)
      Line Input #1, xIn
      Print #2, mId$(xIn, 1, 5) & ";" & mId$(xIn, 6, 5) & ";" _
      & mId$(xIn, 11, 2) & ";" & mId$(xIn, 13, 2) & ";" _
      & mId$(xIn, 15, 3) & ";" & mId$(xIn, 18, 5) & ";" _
      & mId$(xIn, 23, 10) & ";" _
      & mId$(xIn, 33, 3) & ";" & mId$(xIn, 36, 6) & ";" _
      & mId$(xIn, 42, 6) & ";" & mId$(xIn, 48, 11) & ";" _
      & mId$(xIn, 59, 24) & ";" & mId$(xIn, 83, 20) & ";" _
      & mId$(xIn, 103, 16) & ";" & mId$(xIn, 119, 8) & ";" _
      & mId$(xIn, 127, 8) & ";" & mId$(xIn, 135, 8) & ";" _
      & mId$(xIn, 143, 8) & ";" & mId$(xIn, 151, 8) & ";" _
      & mId$(xIn, 159, 3) & ";" & mId$(xIn, 162, 1) & ";" _
      & mId$(xIn, 163, 1) & ";" & mId$(xIn, 164, 1) & ";" _
      & mId$(xIn, 165, 8) & ";" & mId$(xIn, 173, 20) & ";" _
      & mId$(xIn, 193, 24) & ";" & mId$(xIn, 217, 32) & ";" _
      & mId$(xIn, 249, 32) & ";" & mId$(xIn, 281, 1) & ";" _
      & mId$(xIn, 282, 6) & ";" & mId$(xIn, 288, 6) & ";" _
      & mId$(xIn, 294, 16) & ";" & mId$(xIn, 310, 16) & ";" _
      & mId$(xIn, 326, 16) & ";" & mId$(xIn, 342, 8) & ";" _
      & mId$(xIn, 350, 2) & ";" & mId$(xIn, 352, 3) & ";" _
      & mId$(xIn, 355, 2) & ";" & mId$(xIn, 357, 7) & ";" _
      & mId$(xIn, 364, 1) & ";" & mId$(xIn, 365, 1) & ";" _
      & mId$(xIn, 366, 2) & ";" & mId$(xIn, 368, 3) & ";" _
      & mId$(xIn, 371, 4) & ";" & mId$(xIn, 375, 2) & ";" _
      & mId$(xIn, 377, 12) & ";" & mId$(xIn, 389, 3) & ";" _
      & mId$(xIn, 392, 1) & ";" & mId$(xIn, 393, 30) & ";"
Loop
Close
End Sub

Public Sub srvYSITPR10_Export_CSV()
Dim xIn As String
Open "C:\Temp\YSITPR10.txt" For Input As #1
Open "C:\Temp\YSITPR10.csv" For Output As #2
If frmSPLFJOB.chkAS400_Export_CSV = "1" Then
    Print #2, "SITPREETB;SITPREEME;SITPRECOM;SITPREMON;SITPREDAB;SITPREDAF;SITPREDEV;SITPREREF;SITPRETXT;SITPRETOP;SITPREDTF;SITPRENUM;SITPREGLO;SITPREPRO;SITPRECOC;SITPREMTC;SITPREDEC;SITPRETVA;SITPREDVA;SITPREAGE;SITPRESER;SITPRESSE;SITPREOPE;SITPRENAT;"
    Print #2, "Etablissement;Numéro émetteur;Compte client;Montant;Date début;Date fin;Devise;Référence Emetteur;Commentaire;Top compta;Date Facturation;Numéro Opération;Facturation Globale;Proch Factur;Code commission;Mt commission;Devise commission;Montant TVA;Date valeur comm;Agence;Service;Sous-service;Code Opération;Nature Opération;"
    Print #2, ";;;;;;;;;;;;;;;;;;;;;;;;"
End If
Do Until EOF(1)
      Line Input #1, xIn
      Print #2, mId$(xIn, 1, 5) & ";" _
      & mId$(xIn, 6, 7) & ";" _
      & mId$(xIn, 13, 20) & ";" _
      & mId$(xIn, 33, 16) & ";" _
      & mId$(xIn, 49, 8) & ";" _
      & mId$(xIn, 57, 8) & ";" _
      & mId$(xIn, 65, 3) & ";" _
      & mId$(xIn, 68, 18) & ";" _
      & mId$(xIn, 86, 50) & ";" _
      & mId$(xIn, 136, 1) & ";" _
      & mId$(xIn, 137, 8) & ";" _
      & mId$(xIn, 145, 10) & ";" _
      & mId$(xIn, 155, 1) & ";" _
      & mId$(xIn, 156, 8) & ";" _
      & mId$(xIn, 164, 6) & ";" _
      & mId$(xIn, 170, 16) & ";" _
      & mId$(xIn, 186, 3) & ";" _
      & mId$(xIn, 189, 16) & ";" _
      & mId$(xIn, 205, 8) & ";" _
      & mId$(xIn, 213, 5) & ";" _
      & mId$(xIn, 218, 2) & ";" _
      & mId$(xIn, 220, 2) & ";" _
      & mId$(xIn, 222, 3) & ";" _
      & mId$(xIn, 225, 3) & ";"
Loop
Close
End Sub

Public Sub srvYSITPR20_Export_CSV()
Dim xIn As String
Open "C:\Temp\YSITPR20.txt" For Input As #1
Open "C:\Temp\YSITPR20.csv" For Output As #2
If frmSPLFJOB.chkAS400_Export_CSV = "1" Then
    Print #2, "SITPREETB;SITPRECOM;SITPREDAB;SITPREEME;SITPREMON;SITPRERE1;SITPRERE2;SITPRETXT;SITPRETOP;SITPREDEV;SITPREREF;SITPREREJ;SITPREDAF;SITPREDTF;SITPRENUM;SITPREGLO;SITPREPRO;SITPRECOC;SITPREMTC;SITPREDEC;SITPRETVA;SITPREDVA;SITPREDT2;SITPRENU2;SITPREGL2;SITPREPR2;SITPRECO2;SITPREMT2;SITPREDE2;SITPRETV2;SITPREDV2;SITPRETO2;SITPREAGE;SITPRESER;SITPRESSE;SITPREOPE;SITPRENAT;SITPREMO2;"
    Print #2, "Etablissement;Compte client;Date début;Numéro émetteur;Montant minimum;Référence 1;Référence 2;Commentaires;Top commission;Devise;Référence Emetteur;Code Rejet;Date Fin;Date Facturation;Numéro Opération;Facturation Globale;Proch Factur;Code commission;Mt commission;Devise commission;Montant TVA;Date valeur comm;Date Fact.non levée;Numéro Opé.non levée;Fact.Glob.non levée;Proch Fact.non levée;Code comm.non levée;Mt comm.non levée;Devis.comm.non levée;Mt TVA non levée;Val.comm.non levée;Top comm.non levée;Agence;Service;Sous-service;Code Opération;Nature Opération;Montant maximum;"
    Print #2, ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;"
End If
Do Until EOF(1)
      Line Input #1, xIn
      Print #2, mId$(xIn, 1, 5) & ";" & mId$(xIn, 6, 20) & ";" & mId$(xIn, 26, 8) & ";" & mId$(xIn, 34, 7) & ";" _
      & mId$(xIn, 41, 16) & ";" & mId$(xIn, 57, 32) & ";" & mId$(xIn, 89, 32) & ";" _
      & mId$(xIn, 121, 80) & ";" & mId$(xIn, 201, 1) & ";" & mId$(xIn, 202, 3) & ";" _
      & mId$(xIn, 205, 18) & ";" & mId$(xIn, 223, 2) & ";" & mId$(xIn, 225, 8) & ";" _
      & mId$(xIn, 233, 8) & ";" & mId$(xIn, 241, 10) & ";" & mId$(xIn, 251, 1) & ";" _
      & mId$(xIn, 252, 8) & ";" & mId$(xIn, 260, 6) & ";" & mId$(xIn, 266, 16) & ";" _
      & mId$(xIn, 282, 3) & ";" & mId$(xIn, 285, 16) & ";" & mId$(xIn, 301, 8) & ";" _
      & mId$(xIn, 309, 8) & ";" & mId$(xIn, 317, 10) & ";" & mId$(xIn, 327, 1) & ";" _
      & mId$(xIn, 328, 8) & ";" & mId$(xIn, 336, 6) & ";" & mId$(xIn, 342, 16) & ";" _
      & mId$(xIn, 358, 3) & ";" & mId$(xIn, 361, 16) & ";" & mId$(xIn, 377, 8) & ";" _
      & mId$(xIn, 385, 1) & ";" & mId$(xIn, 386, 5) & ";" & mId$(xIn, 391, 2) & ";" _
      & mId$(xIn, 393, 2) & ";" & mId$(xIn, 395, 3) & ";" & mId$(xIn, 398, 3) & ";" _
      & mId$(xIn, 401, 16) & ";"
Loop
Close
End Sub


