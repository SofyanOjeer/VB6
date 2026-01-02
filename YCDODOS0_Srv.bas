Attribute VB_Name = "srvYCDODOS0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYCDODOS0Len = 1331 ' 34 + 1297
Public Const recYCDODOS0_Block = 10
Public Const constYCDODOS0 = "YCDODOS0"
Dim meYbase As typeYBase
Dim paramYCDODOS0_Import As String

Type typeYCDODOS0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    CDODOSETB       As Integer                        ' CODE ETABLISSEMENT
    CDODOSAGE       As Integer                        ' AGENCE
    CDODOSSER       As String * 2                     ' SERVICE
    CDODOSSSE       As String * 2                     ' SOUS-SERVICE
    CDODOSCOP       As String * 3                     ' CODE OPERATION
    CDODOSDOS       As Long                           ' NUMERO DOSSIER
    CDODOSNUR       As Long                           ' N° RENOUVELLEMENT
    CDODOSNAT       As String * 3                     ' NATURE
    CDODOSEXT       As String * 16                    ' REFERENCE EXTERNE
    CDODOSMON       As Currency                       ' MONTANT DOSSIER
    CDODOSDEV       As String * 3                     ' DEVISE
    CDODOSMOA       As Currency                       ' MONTANT ADDITIONNEL
    CDODOSMOT       As Currency                       ' MONTANT TOTAL
    CDODOSMOC       As Currency                       ' MONTANT CONFIRME
    CDODOSMOD       As Currency                       ' MONTANT DUCROIRE
    CDODOSCON       As String * 1                     ' CONFIRM NOTIFI PARTI
    CDODOSIRR       As String * 1                     ' IRREVOCABLE (O/N)
    CDODOSFRA       As String * 1                     ' FRACTIONNABLE (O/N)
    CDODOSREN       As String * 1                     ' RENOUVELABLE (O/N)
    CDODOSCUM       As String * 1                     ' CUMULATIF (O/N)
    CDODOSTRS       As String * 1                     ' TRANSFERABLE
    CDODOSTOL       As Currency                       ' TOLERANCE +
    CDODOSTO2       As Currency                       ' TOLERANCE -
    CDODOSDOR       As String * 1                     ' DONN. ORDRE CLI/TIE
    CDODOSDON       As String * 7                     ' DONNEUR ORDRE IMPORT
    CDODOSDOE       As String * 64                    ' DONNEUR ORDRE EXPORT
    CDODOSBER       As String * 1                     ' BENEFICIAIR CLI/TIE
    CDODOSBEN       As String * 7                     ' BENEFICIAIRE EXPORT
    CDODOSBEI       As String * 64                    ' BENEFICIAIRE IMPORT
    CDODOSBAR       As String * 1                     ' BANQU.BENEF.CLI/TIE
    CDODOSBAB       As String * 7                     ' BANQUE BENEF
    CDODOSNOR       As String * 1                     ' NOTIF/CONFI OU EMETT
    CDODOSNOT       As String * 7                     ' NOTIF/CONFI OU EMETT
    CDODOSBIC       As String * 12                    ' BIC SUPPLEMEN. EMETT
    CDODOSCOT       As String * 1                     ' CORRESPOND. CLI/TIE
    CDODOSCOR       As String * 7                     ' CORRESPONDANT
    CDODOSPRT       As String * 1                     ' LIEU PRES CLI/TIE
    CDODOSPRR       As String * 7                     ' LIEU PRESENTATION
    CDODOSUTV       As String * 32                    ' LIEU PRESENTATION
    CDODOSPAT       As String * 1                     ' LIEU PAIE CLI/TIE
    CDODOSPAR       As String * 7                     ' LIEU PAIEMENT
    CDODOSPAV       As String * 32                    ' LIEU PAIEMENT
    CDODOSOUV       As Long                           ' DATE OUVERTURE
    CDODOSEMI       As Long                           ' DATE EMISSION
    CDODOSVAL       As Long                           ' DATE VALIDITE
    CDODOSDEP       As Long                           ' DATE EXTREME PAYMT
    CDODOSDTR       As Long                           ' DATE DE TRANSFERT
    CDODOSVCP       As Long                           ' DATE VALID. COMPTA
    CDODOSCLO       As Long                           ' DATE CLOTURE
    CDODOSREJ       As String * 3                     ' MOTIF REJET (CLOTUR)
    CDODOSOBJ       As String * 6                     ' OBJET CREDIT
    CDODOSAVU       As Long                           ' % PAIEM. A VUE
    CDODOSMOV       As Currency                       ' MONTANT A VUE
    CDODOSCAC       As Long                           ' % PAIEM. CTR ACCEPT.
    CDODOSMCA       As Currency                       ' MONTANT CTR ACCEPT.
    CDODOSDIF       As Long                           ' % PAIEM. DIFFERE
    CDODOSMDI       As Currency                       ' MONTANT. DIFFERE
    CDODOSPMO       As Currency                       ' MONTANT PROVISIONNE
    CDODOSPCD       As String * 20                    ' PROV. DEBIT  COMPTE
    CDODOSPCC       As String * 20                    ' PROV. CREDIT COMPTE
    CDODOSPDE       As Currency                       ' PROVISION DEVISE DOS
    CDODOSPPO       As Long                           ' PROVISION POURCEN
    CDODOSAUT       As String * 12                    ' CODE AUTORISATION
    CDODOSREG       As Currency                       ' MONTANT PAYE
    CDODOSENC       As Currency                       ' MONTANT ENCAISSE
    CDODOSDAN       As Long                           ' DATE ANNULATION
    CDODOSANN       As Currency                       ' MONTANT ANNULE
    CDODOSPCO       As Double                         ' COURS DEVPRO/DEVDOS
    CDODOSLEM       As String * 30                    ' LIEU EMBARQUEMENT
    CDODOSLDE       As String * 30                    ' LIEU DESTINATION
    CDODOSDLE       As Long                           ' DATE LIMITE EMBARQU.
    CDODOSEPA       As String * 1                     ' EXPED.PARTIE.AUTORI
    CDODOSTRA       As String * 1                     ' TRANBORDEMENT AUTORI
    CDODOSFCD       As String * 1                     ' FRAI CHARGE D.O. BEN
    CDODOSCUS       As Integer                        ' UTILI. DE SAISIE
    CDODOSCUV       As Integer                        ' 1ER VALIDEUR
    CDODOSCU2       As Integer                        ' 2EME VALIDEUR
    CDODOSOPE       As String * 1                     ' OPERATIVITE DU CRED.
    CDODOSPOO       As String * 1                     ' EXISTENCE POOL
    CDODOSPBE       As Currency                       ' PART.BANQUE EXPORT
    CDODOSGAG       As String * 1                     ' GAGE MARCHANDISE
    CDODOSSTB       As String * 1                     ' STAND BY
    CDODOSMRE       As String * 3                     ' MODE DE REALISAT°
    CDODOSNPD       As Long                           ' NBJ PRES. DOCUMENT
    CDODOSTJD       As String * 1                     ' TY JOUR DOCS
    CDODOSPDO       As String * 60                    ' PER.PRE.DOCS.
    CDODOSGAR       As String * 64                    ' LIBELLE GARANTIE
    CDODOSOBM       As String * 64                    ' OBJET DE MODIF.
    CDODOSTBR       As String * 1                     ' TIERS BQ REMBOURS
    CDODOSBRE       As String * 7                     ' BQ REMBOURSEMENT
    CDODOSBEC       As String * 1                     ' BENEF PAY.COMMIS°
    CDODOSRNO       As String * 16                    ' REF.NOTIFICATEUR
    CDODOSDPA       As String * 3                     ' DESTINATION PAYS
    CDODOSDVI       As String * 32                    ' DESTINATION VILLE
    CDODOSEPY       As String * 3                     ' EMBARQUEMENT PAYS
    CDODOSEVI       As String * 32                    ' EMBARQUEMENT VILLE
    CDODOSVPA       As String * 3                     ' VALIDITE PAYS
    CDODOSVVI       As String * 32                    ' VALIDIT VILLE
    CDODOSNDE       As Long                           ' DOSSIER EXPORT
    CDODOSNAE       As String * 3                     ' NATURE EXPORT
    CDODOSEVE       As String * 2                     ' EVENEMENT
    CDODOSETA       As String * 2                     ' ETAT DOSSIER
    CDODOSDP2       As String * 32                    ' DESTIN.PAYS LIBELLE
    CDODOSEP2       As String * 32                    ' EMBARQ.PAYS LIBELLE
    CDODOSPD2       As String * 80                    ' PER.PRES.DOC.SUITE
    CDODOSAUN       As String * 12                    ' CODE AUT. NOTIFIE
    CDODOSCER       As String * 1                     ' COTAT°(O=CERTAIN/N)
End Type
    
Public Function srvYCDODOS0_Import(lX As String)
Dim xIn As String, X As String, Nb As Long
On Error GoTo Error_Handle

recYBase_Init meYbase
meYbase.Method = "Seek="

meYbase.ID = constYBase
meYbase.K1 = YBIATAB0_DATE_CPT_J & "_" & constYCDODOS0
If tableYBase_Read(meYbase) = 0 Then                            'Fichier du jour déjà lu ?
    srvYCDODOS0_Import = Null
    lX = CStr(meYbase.Text)
    Exit Function
End If


srvYCDODOS0_Import = "?"

paramYCDODOS0_Import = paramYBase_DataF & Trim(constYCDODOS0) & paramYBase_Data_ExtensionP

Open Trim(paramYCDODOS0_Import) For Input As #1

Nb = 0
X = "delete * from YBase where Id = " & Chr$(34) & Trim(constYCDODOS0) & Chr$(34)
MDB.Execute X

meYbase.Method = constAddNew


Do Until EOF(1)
    Nb = Nb + 1
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then

            meYbase.ID = constYCDODOS0
            meYbase.K1 = mId$(xIn, 15, 13) 'recYCDODOS0.CDODOSCOP & recYCDODOS0.CDODOSDOS
            meYbase.Text = xIn
            dbYBase_Update meYbase
            
    End If
Loop


Close
srvYCDODOS0_Import = Null
meYbase.ID = constYBase
meYbase.K1 = YBIATAB0_DATE_CPT_J & "_" & constYCDODOS0
lX = DSys & "_" & time_Hms & "_" & Nb
meYbase.Text = lX
dbYBase_Update meYbase

Exit Function

Error_Handle:
 MsgBox "erreur : srvYCDODOS0_Import" & xIn, vbCritical, Error
Close

srvYCDODOS0_Import = Error
End Function

Public Function srvYCDODOS0_Import_Read(lId As String, lYCDODOS0 As typeYCDODOS0)

Dim xIn As String, X As String

On Error GoTo Error_Handle

srvYCDODOS0_Import_Read = "?"

meYbase.Method = "Seek="
meYbase.ID = constYCDODOS0
meYbase.K1 = lId

If tableYBase_Read(meYbase) = 0 Then
    MsgTxt = Space$(34) & meYbase.Text
    MsgTxtIndex = 0
    srvYCDODOS0_GetBuffer lYCDODOS0
    srvYCDODOS0_Import_Read = Null
End If
        
Exit Function

Error_Handle:
 MsgBox "erreur : srvYCDODOS0_Import_Read" & xIn, vbCritical, Error
srvYCDODOS0_Import_Read = Error
End Function





Public Sub srvYCDODOS0_ElpDisplay(recYCDODOS0 As typeYCDODOS0)
frmElpDisplay.fgData.Rows = 108
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSETB    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSETB
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSAGE    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AGENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSAGE
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSSER    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSSER
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSSSE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOUS-SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSSSE
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSCOP    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE OPERATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSCOP
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSDOS    9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO DOSSIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSDOS
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSNUR    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° RENOUVELLEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSNUR
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSNAT    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NATURE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSNAT
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSEXT   16A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "REFERENCE EXTERNE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSEXT
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSMON 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT DOSSIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSMON
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSDEV    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DEVISE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSDEV
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSMOA 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT ADDITIONNEL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSMOA
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSMOT 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT TOTAL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSMOT
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSMOC 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT CONFIRME"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSMOC
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSMOD 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT DUCROIRE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSMOD
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSCON    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CONFIRM NOTIFI PARTI"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSCON
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSIRR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "IRREVOCABLE (O/N)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSIRR
frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSFRA    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "FRACTIONNABLE (O/N)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSFRA
frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSREN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "RENOUVELABLE (O/N)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSREN
frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSCUM    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CUMULATIF (O/N)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSCUM
frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSTRS    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TRANSFERABLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSTRS
frmElpDisplay.fgData.Row = 22
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSTOL  3.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TOLERANCE +"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSTOL
frmElpDisplay.fgData.Row = 23
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSTO2  3.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TOLERANCE -"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSTO2
frmElpDisplay.fgData.Row = 24
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSDOR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DONN. ORDRE CLI/TIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSDOR
frmElpDisplay.fgData.Row = 25
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSDON    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DONNEUR ORDRE IMPORT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSDON
frmElpDisplay.fgData.Row = 26
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSDOE   64A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DONNEUR ORDRE EXPORT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSDOE
frmElpDisplay.fgData.Row = 27
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSBER    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BENEFICIAIR CLI/TIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSBER
frmElpDisplay.fgData.Row = 28
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSBEN    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BENEFICIAIRE EXPORT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSBEN
frmElpDisplay.fgData.Row = 29
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSBEI   64A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BENEFICIAIRE IMPORT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSBEI
frmElpDisplay.fgData.Row = 30
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSBAR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BANQU.BENEF.CLI/TIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSBAR
frmElpDisplay.fgData.Row = 31
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSBAB    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BANQUE BENEF"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSBAB
frmElpDisplay.fgData.Row = 32
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSNOR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NOTIF/CONFI OU EMETT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSNOR
frmElpDisplay.fgData.Row = 33
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSNOT    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NOTIF/CONFI OU EMETT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSNOT
frmElpDisplay.fgData.Row = 34
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSBIC   12A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BIC SUPPLEMEN. EMETT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSBIC
frmElpDisplay.fgData.Row = 35
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSCOT    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CORRESPOND. CLI/TIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSCOT
frmElpDisplay.fgData.Row = 36
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSCOR    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CORRESPONDANT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSCOR
frmElpDisplay.fgData.Row = 37
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSPRT    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIEU PRES CLI/TIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSPRT
frmElpDisplay.fgData.Row = 38
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSPRR    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIEU PRESENTATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSPRR
frmElpDisplay.fgData.Row = 39
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSUTV   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIEU PRESENTATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSUTV
frmElpDisplay.fgData.Row = 40
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSPAT    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIEU PAIE CLI/TIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSPAT
frmElpDisplay.fgData.Row = 41
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSPAR    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIEU PAIEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSPAR
frmElpDisplay.fgData.Row = 42
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSPAV   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIEU PAIEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSPAV
frmElpDisplay.fgData.Row = 43
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSOUV    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE OUVERTURE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSOUV
frmElpDisplay.fgData.Row = 44
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSEMI    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE EMISSION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSEMI
frmElpDisplay.fgData.Row = 45
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSVAL    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE VALIDITE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSVAL
frmElpDisplay.fgData.Row = 46
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSDEP    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE EXTREME PAYMT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSDEP
frmElpDisplay.fgData.Row = 47
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSDTR    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE DE TRANSFERT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSDTR
frmElpDisplay.fgData.Row = 48
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSVCP    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE VALID. COMPTA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSVCP
frmElpDisplay.fgData.Row = 49
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSCLO    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE CLOTURE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSCLO
frmElpDisplay.fgData.Row = 50
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSREJ    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MOTIF REJET (CLOTUR)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSREJ
frmElpDisplay.fgData.Row = 51
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSOBJ    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "OBJET CREDIT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSOBJ
frmElpDisplay.fgData.Row = 52
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSAVU    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "% PAIEM. A VUE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSAVU
frmElpDisplay.fgData.Row = 53
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSMOV 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT A VUE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSMOV
frmElpDisplay.fgData.Row = 54
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSCAC    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "% PAIEM. CTR ACCEPT."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSCAC
frmElpDisplay.fgData.Row = 55
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSMCA 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT CTR ACCEPT."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSMCA
frmElpDisplay.fgData.Row = 56
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSDIF    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "% PAIEM. DIFFERE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSDIF
frmElpDisplay.fgData.Row = 57
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSMDI 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT. DIFFERE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSMDI
frmElpDisplay.fgData.Row = 58
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSPMO 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT PROVISIONNE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSPMO
frmElpDisplay.fgData.Row = 59
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSPCD   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PROV. DEBIT  COMPTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSPCD
frmElpDisplay.fgData.Row = 60
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSPCC   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PROV. CREDIT COMPTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSPCC
frmElpDisplay.fgData.Row = 61
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSPDE 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PROVISION DEVISE DOS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSPDE
frmElpDisplay.fgData.Row = 62
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSPPO    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PROVISION POURCEN"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSPPO
frmElpDisplay.fgData.Row = 63
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSAUT   12A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE AUTORISATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSAUT
frmElpDisplay.fgData.Row = 64
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSREG 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT PAYE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSREG
frmElpDisplay.fgData.Row = 65
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSENC 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT ENCAISSE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSENC
frmElpDisplay.fgData.Row = 66
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSDAN    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE ANNULATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSDAN
frmElpDisplay.fgData.Row = 67
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSANN 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT ANNULE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSANN
frmElpDisplay.fgData.Row = 68
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSPCO 14.9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COURS DEVPRO/DEVDOS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSPCO
frmElpDisplay.fgData.Row = 69
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSLEM   30A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIEU EMBARQUEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSLEM
frmElpDisplay.fgData.Row = 70
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSLDE   30A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIEU DESTINATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSLDE
frmElpDisplay.fgData.Row = 71
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSDLE    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE LIMITE EMBARQU."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSDLE
frmElpDisplay.fgData.Row = 72
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSEPA    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "EXPED.PARTIE.AUTORI"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSEPA
frmElpDisplay.fgData.Row = 73
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSTRA    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TRANBORDEMENT AUTORI"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSTRA
frmElpDisplay.fgData.Row = 74
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSFCD    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "FRAI CHARGE D.O. BEN"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSFCD
frmElpDisplay.fgData.Row = 75
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSCUS    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "UTILI. DE SAISIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSCUS
frmElpDisplay.fgData.Row = 76
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSCUV    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "1ER VALIDEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSCUV
frmElpDisplay.fgData.Row = 77
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSCU2    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "2EME VALIDEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSCU2
frmElpDisplay.fgData.Row = 78
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSOPE    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "OPERATIVITE DU CRED."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSOPE
frmElpDisplay.fgData.Row = 79
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSPOO    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "EXISTENCE POOL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSPOO
frmElpDisplay.fgData.Row = 80
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSPBE 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PART.BANQUE EXPORT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSPBE
frmElpDisplay.fgData.Row = 81
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSGAG    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "GAGE MARCHANDISE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSGAG
frmElpDisplay.fgData.Row = 82
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSSTB    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "STAND BY"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSSTB
frmElpDisplay.fgData.Row = 83
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSMRE    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MODE DE REALISAT°"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSMRE
frmElpDisplay.fgData.Row = 84
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSNPD    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NBJ PRES. DOCUMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSNPD
frmElpDisplay.fgData.Row = 85
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSTJD    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TY JOUR DOCS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSTJD
frmElpDisplay.fgData.Row = 86
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSPDO   60A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PER.PRE.DOCS."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSPDO
frmElpDisplay.fgData.Row = 87
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSGAR   64A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIBELLE GARANTIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSGAR
frmElpDisplay.fgData.Row = 88
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSOBM   64A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "OBJET DE MODIF."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSOBM
frmElpDisplay.fgData.Row = 89
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSTBR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TIERS BQ REMBOURS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSTBR
frmElpDisplay.fgData.Row = 90
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSBRE    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BQ REMBOURSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSBRE
frmElpDisplay.fgData.Row = 91
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSBEC    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BENEF PAY.COMMIS°"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSBEC
frmElpDisplay.fgData.Row = 92
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSRNO   16A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "REF.NOTIFICATEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSRNO
frmElpDisplay.fgData.Row = 93
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSDPA    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DESTINATION PAYS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSDPA
frmElpDisplay.fgData.Row = 94
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSDVI   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DESTINATION VILLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSDVI
frmElpDisplay.fgData.Row = 95
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSEPY    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "EMBARQUEMENT PAYS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSEPY
frmElpDisplay.fgData.Row = 96
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSEVI   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "EMBARQUEMENT VILLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSEVI
frmElpDisplay.fgData.Row = 97
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSVPA    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "VALIDITE PAYS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSVPA
frmElpDisplay.fgData.Row = 98
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSVVI   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "VALIDIT VILLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSVVI
frmElpDisplay.fgData.Row = 99
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSNDE    9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DOSSIER EXPORT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSNDE
frmElpDisplay.fgData.Row = 100
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSNAE    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NATURE EXPORT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSNAE
frmElpDisplay.fgData.Row = 101
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSEVE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "EVENEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSEVE
frmElpDisplay.fgData.Row = 102
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSETA    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETAT DOSSIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSETA
frmElpDisplay.fgData.Row = 103
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSDP2   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DESTIN.PAYS LIBELLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSDP2
frmElpDisplay.fgData.Row = 104
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSEP2   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "EMBARQ.PAYS LIBELLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSEP2
frmElpDisplay.fgData.Row = 105
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSPD2   80A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PER.PRES.DOC.SUITE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSPD2
frmElpDisplay.fgData.Row = 106
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSAUN   12A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE AUT. NOTIFIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSAUN
frmElpDisplay.fgData.Row = 107
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDODOSCER    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COTAT°(O=CERTAIN/N)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDODOS0.CDODOSCER
frmElpDisplay.Show vbModal
End Sub


Public Sub srvYCDODOS0_Export_CSV(lIdFile_Source As Integer, lIdFile_Destination As Integer, loptSelect_CSV_Header As Boolean, lnb As Long)
Dim xIn As String
Dim X1 As String, X2 As String

If loptSelect_CSV_Header Then
    Print #lIdFile_Destination, "CDODOSETB;CDODOSAGE;CDODOSSER;CDODOSSSE;CDODOSCOP;CDODOSDOS;CDODOSNUR;CDODOSNAT;CDODOSEXT;CDODOSMON;CDODOSDEV;CDODOSMOA;CDODOSMOT;CDODOSMOC;CDODOSMOD;CDODOSCON;CDODOSIRR;CDODOSFRA;CDODOSREN" _
    & ";CDODOSCUM;CDODOSTRS;CDODOSTOL;CDODOSTO2;CDODOSDOR;CDODOSDON;CDODOSDOE;CDODOSBER;CDODOSBEN;CDODOSBEI;CDODOSBAR;CDODOSBAB;CDODOSNOR;CDODOSNOT;CDODOSBIC;CDODOSCOT;CDODOSCOR;CDODOSPRT;CDODOSPRR;CDODOSUTV" _
    & ";CDODOSPAT;CDODOSPAR;CDODOSPAV;CDODOSOUV;CDODOSEMI;CDODOSVAL;CDODOSDEP;CDODOSDTR;CDODOSVCP;CDODOSCLO;CDODOSREJ;CDODOSOBJ;CDODOSAVU;CDODOSMOV;CDODOSCAC;CDODOSMCA;CDODOSDIF;CDODOSMDI;CDODOSPMO;CDODOSPCD" _
    & ";CDODOSPCC;CDODOSPDE;CDODOSPPO;CDODOSAUT;CDODOSREG;CDODOSENC;CDODOSDAN;CDODOSANN;CDODOSPCO;CDODOSLEM;CDODOSLDE;CDODOSDLE;CDODOSEPA;CDODOSTRA;CDODOSFCD;CDODOSCUS;CDODOSCUV;CDODOSCU2;CDODOSOPE;CDODOSPOO" _
    & ";CDODOSPBE;CDODOSGAG;CDODOSSTB;CDODOSMRE;CDODOSNPD;CDODOSTJD;CDODOSPDO;CDODOSGAR;CDODOSOBM;CDODOSTBR;CDODOSBRE;CDODOSBEC;CDODOSRNO;CDODOSDPA;CDODOSDVI;CDODOSEPY;CDODOSEVI;CDODOSVPA;CDODOSVVI;CDODOSNDE" _
    & ";CDODOSNAE;CDODOSEVE;CDODOSETA;CDODOSDP2;CDODOSEP2;CDODOSPD2;CDODOSAUN;CDODOSCER;"
    Print #lIdFile_Destination, "CODE ETABLISSEMENT;AGENCE;SERVICE;SOUS-SERVICE;CODE OPERATION;NUMERO DOSSIER;N° RENOUVELLEMENT;NATURE;REFERENCE EXTERNE;MONTANT DOSSIER;DEVISE;MONTANT ADDITIONNEL;MONTANT TOTAL;MONTANT CONFIRME" _
    & ";MONTANT DUCROIRE;CONFIRM NOTIFI PARTI;IRREVOCABLE (O/N);FRACTIONNABLE (O/N);RENOUVELABLE (O/N);CUMULATIF (O/N);TRANSFERABLE;TOLERANCE +;TOLERANCE -;DONN. ORDRE CLI/TIE;DONNEUR ORDRE IMPORT;DONNEUR ORDRE EXPORT" _
    & ";BENEFICIAIR CLI/TIE;BENEFICIAIRE EXPORT;BENEFICIAIRE IMPORT;BANQU.BENEF.CLI/TIE;BANQUE BENEF;NOTIF/CONFI OU EMETT;NOTIF/CONFI OU EMETT;BIC SUPPLEMEN. EMETT;CORRESPOND. CLI/TIE;CORRESPONDANT;LIEU PRES CLI/TIE" _
    & ";LIEU PRESENTATION;LIEU PRESENTATION;LIEU PAIE CLI/TIE;LIEU PAIEMENT;LIEU PAIEMENT;DATE OUVERTURE;DATE EMISSION;DATE VALIDITE;DATE EXTREME PAYMT;DATE DE TRANSFERT;DATE VALID. COMPTA;DATE CLOTURE;MOTIF REJET (CLOTUR)" _
    & ";OBJET CREDIT;% PAIEM. A VUE;MONTANT A VUE;% PAIEM. CTR ACCEPT.;MONTANT CTR ACCEPT.;% PAIEM. DIFFERE;MONTANT. DIFFERE;MONTANT PROVISIONNE;PROV. DEBIT  COMPTE;PROV. CREDIT COMPTE;PROVISION DEVISE DOS;PROVISION POURCEN" _
    & ";CODE AUTORISATION;MONTANT PAYE;MONTANT ENCAISSE;DATE ANNULATION;MONTANT ANNULE;COURS DEVPRO/DEVDOS;LIEU EMBARQUEMENT;LIEU DESTINATION;DATE LIMITE EMBARQU.;EXPED.PARTIE.AUTORI;TRANBORDEMENT AUTORI;FRAI CHARGE D.O. BEN" _
    & ";UTILI. DE SAISIE;1ER VALIDEUR;2EME VALIDEUR;OPERATIVITE DU CRED.;EXISTENCE POOL;PART.BANQUE EXPORT;GAGE MARCHANDISE;STAND BY;MODE DE REALISAT°;NBJ PRES. DOCUMENT;TY JOUR DOCS;PER.PRE.DOCS.;LIBELLE GARANTIE;OBJET DE MODIF.;TIERS BQ REMBOURS;BQ REMBOURSEMENT;BENEF PAY.COMMIS°;REF.NOTIFICATEUR;DESTINATION PAYS;DESTINATION VILLE;EMBARQUEMENT PAYS;EMBARQUEMENT VILLE;VALIDITE PAYS;VALIDIT VILLE;DOSSIER EXPORT;NATURE EXPORT;EVENEMENT;ETAT DOSSIER;DESTIN.PAYS LIBELLE;EMBARQ.PAYS LIBELLE;PER.PRES.DOC.SUITE;CODE AUT. NOTIFIE;COTAT°(O=CERTAIN/N);"
    Print #lIdFile_Destination, ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;"
End If
Do Until EOF(lIdFile_Source)
      Line Input #lIdFile_Source, xIn
      lnb = lnb + 1
      
      X1 = mId$(xIn, 6, 5) & ";" & mId$(xIn, 11, 2) & ";" & mId$(xIn, 13, 2) & ";" & mId$(xIn, 15, 3) & ";" _
      & mId$(xIn, 18, 10) & ";" & mId$(xIn, 28, 4) & ";" & mId$(xIn, 32, 3) & ";" & mId$(xIn, 35, 16) & ";" _
      & cur_19V(CCur(mId$(xIn, 51, 16)) / 100) & ";" & mId$(xIn, 67, 3) & ";" _
      & cur_19V(CCur(mId$(xIn, 70, 16)) / 100) & ";" _
      & cur_19V(CCur(mId$(xIn, 86, 16)) / 100) & ";" _
      & cur_19V(CCur(mId$(xIn, 102, 16)) / 100) & ";" _
      & cur_19V(CCur(mId$(xIn, 118, 16)) / 100) & ";" _
      & mId$(xIn, 134, 1) & ";" & mId$(xIn, 135, 1) & ";" & mId$(xIn, 136, 1) & ";" & mId$(xIn, 137, 1) & ";" & mId$(xIn, 138, 1) & ";" & mId$(xIn, 139, 1) & ";" & mId$(xIn, 140, 4) & ";" _
      & mId$(xIn, 144, 4) & ";" & mId$(xIn, 148, 1) & ";" & mId$(xIn, 149, 7) & ";" & mId$(xIn, 156, 64) & ";" & mId$(xIn, 220, 1) & ";" _
      & mId$(xIn, 221, 7) & ";" & mId$(xIn, 228, 64) & ";" & mId$(xIn, 292, 1) & ";" & mId$(xIn, 293, 7) & ";" & mId$(xIn, 300, 1) & ";" _
      & mId$(xIn, 301, 7) & ";" & mId$(xIn, 308, 12) & ";" & mId$(xIn, 320, 1) & ";" & mId$(xIn, 321, 7) & ";" & mId$(xIn, 328, 1) & ";" _
      & mId$(xIn, 329, 7) & ";" & mId$(xIn, 336, 32) & ";" & mId$(xIn, 368, 1) & ";" & mId$(xIn, 369, 7) & ";" & mId$(xIn, 376, 32) & ";" _
      & mId$(xIn, 408, 8) & ";" & mId$(xIn, 416, 8) & ";" & mId$(xIn, 424, 8) & ";" & mId$(xIn, 432, 8) & ";" & mId$(xIn, 440, 8) & ";" _
      & mId$(xIn, 448, 8) & ";" & mId$(xIn, 456, 8) & ";" & mId$(xIn, 464, 3) & ";" & mId$(xIn, 467, 6) & ";" & mId$(xIn, 473, 4) & ";"

      X2 = cur_19V(CCur(mId$(xIn, 477, 16)) / 100) & ";" & mId$(xIn, 493, 4) & ";" _
      & cur_19V(CCur(mId$(xIn, 497, 16)) / 100) & ";" & mId$(xIn, 513, 4) & ";" _
      & cur_19V(CCur(mId$(xIn, 517, 16)) / 100) & ";" _
      & cur_19V(CCur(mId$(xIn, 533, 16)) / 100) & ";" _
      & mId$(xIn, 549, 20) & ";" _
      & mId$(xIn, 569, 20) & ";" _
      & cur_19V(CCur(mId$(xIn, 589, 16)) / 100) & ";" _
      & mId$(xIn, 605, 4) & ";" & mId$(xIn, 609, 12) & ";" _
      & cur_19V(CCur(mId$(xIn, 621, 16)) / 100) & ";" _
      & cur_19V(CCur(mId$(xIn, 637, 16)) / 100) & ";" & mId$(xIn, 653, 8) & ";" _
      & cur_19V(CCur(mId$(xIn, 661, 16)) / 100) & ";" & mId$(xIn, 677, 15) & ";" _
      & mId$(xIn, 692, 30) & ";" & mId$(xIn, 722, 30) & ";" & mId$(xIn, 752, 8) & ";" & mId$(xIn, 760, 1) & ";" & mId$(xIn, 761, 1) & ";" _
      & mId$(xIn, 762, 1) & ";" & mId$(xIn, 763, 5) & ";" & mId$(xIn, 768, 5) & ";" & mId$(xIn, 773, 5) & ";" & mId$(xIn, 778, 1) & ";" _
      & mId$(xIn, 779, 1) & ";" _
      & cur_19V(CCur(mId$(xIn, 780, 16)) / 100) & ";" _
      & mId$(xIn, 796, 1) & ";" & mId$(xIn, 797, 1) & ";" & mId$(xIn, 798, 3) & ";" _
      & mId$(xIn, 801, 4) & ";" & mId$(xIn, 805, 1) & ";" & mId$(xIn, 806, 60) & ";" & mId$(xIn, 866, 64) & ";" & mId$(xIn, 930, 64) & ";" _
      & mId$(xIn, 994, 1) & ";" & mId$(xIn, 995, 7) & ";" & mId$(xIn, 1002, 1) & ";" & mId$(xIn, 1003, 16) & ";" & mId$(xIn, 1019, 3) & ";" _
      & mId$(xIn, 1022, 32) & ";" & mId$(xIn, 1054, 3) & ";" & mId$(xIn, 1057, 32) & ";" & mId$(xIn, 1089, 3) & ";" & mId$(xIn, 1092, 32) & ";" _
      & mId$(xIn, 1124, 10) & ";" & mId$(xIn, 1134, 3) & ";" & mId$(xIn, 1137, 2) & ";" & mId$(xIn, 1139, 2) & ";" & mId$(xIn, 1141, 32) & ";" _
      & mId$(xIn, 1173, 32) & ";" & mId$(xIn, 1205, 80) & ";" & mId$(xIn, 1285, 12) & ";" & mId$(xIn, 1297, 1) & ";"
  
      Print #lIdFile_Destination, mId$(xIn, 1, 5) & ";" & X1 & X2
Loop
End Sub

Public Sub srvYCDODOS0_Load(lYCDODOS0() As typeYCDODOS0, lYCDODOS0_Nb As Integer)
Dim mMethod As String, blnYCDODOS0_Suite
Dim wNbMax As Integer
Dim wYCDODOS0 As typeYCDODOS0

mMethod = Trim(lYCDODOS0(0).Method) & "+"
blnYCDODOS0_Suite = True: lYCDODOS0_Nb = 0
wNbMax = recYCDODOS0_Block + 2: ReDim Preserve lYCDODOS0(wNbMax)

wYCDODOS0 = lYCDODOS0(1)
Do Until Not blnYCDODOS0_Suite
    MsgTxtLen = 0
    Call srvYCDODOS0_PutBuffer(wYCDODOS0)
    Call srvYCDODOS0_PutBuffer(lYCDODOS0(0))
    If IsNull(SndRcv()) Then
        MsgTxtIndex = 0
        Do While MsgTxtIndex < MsgTxtLen
            If IsNull(srvYCDODOS0_GetBuffer(wYCDODOS0)) Then
            
                lYCDODOS0_Nb = lYCDODOS0_Nb + 1
                If lYCDODOS0_Nb > wNbMax Then
                    wNbMax = wNbMax + recYCDODOS0_Block
                    ReDim Preserve lYCDODOS0(wNbMax)
                End If
            
                lYCDODOS0(lYCDODOS0_Nb) = wYCDODOS0
                blnYCDODOS0_Suite = True
            Else
                blnYCDODOS0_Suite = False
                Exit Do
            End If
        Loop
    End If

    lYCDODOS0(0).Method = mMethod
Loop

End Sub

'-----------------------------------------------------
Public Function srvYCDODOS0_Monitor(recYCDODOS0 As typeYCDODOS0)
'-----------------------------------------------------

Select Case mId$(Trim(recYCDODOS0.Method), 1, 4)
    Case "Seek"
                srvYCDODOS0_Monitor = srvYCDODOS0_Seek(recYCDODOS0)
    Case Else
                recYCDODOS0.Err = recYCDODOS0.Method
                Call srvYCDODOS0_Error(recYCDODOS0)
                srvYCDODOS0_Monitor = recYCDODOS0.Err
End Select

End Function

'-----------------------------------------------------
Sub srvYCDODOS0_Error(recYCDODOS0 As typeYCDODOS0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YCDODOS0" & Chr$(10) & Chr$(13)

Select Case mId$(recYCDODOS0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYCDODOS0.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : YCDODOS0s.bas  ( " _
                & Trim(recYCDODOS0.obj) & " : " & Trim(recYCDODOS0.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvYCDODOS0_GetBuffer(recYCDODOS0 As typeYCDODOS0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYCDODOS0_GetBuffer = Null
recYCDODOS0.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYCDODOS0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYCDODOS0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYCDODOS0.Err = Space$(10) Then
    recYCDODOS0.CDODOSETB = CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYCDODOS0.CDODOSAGE = CInt(Val(mId$(MsgTxt, K + 6, 5)))
    recYCDODOS0.CDODOSSER = mId$(MsgTxt, K + 11, 2)
    recYCDODOS0.CDODOSSSE = mId$(MsgTxt, K + 13, 2)
    recYCDODOS0.CDODOSCOP = mId$(MsgTxt, K + 15, 3)
    recYCDODOS0.CDODOSDOS = CLng(Val(mId$(MsgTxt, K + 18, 10)))
    recYCDODOS0.CDODOSNUR = CLng(Val(mId$(MsgTxt, K + 28, 4)))
    recYCDODOS0.CDODOSNAT = mId$(MsgTxt, K + 32, 3)
    recYCDODOS0.CDODOSEXT = mId$(MsgTxt, K + 35, 16)
    recYCDODOS0.CDODOSMON = CCur(Val(mId$(MsgTxt, K + 51, 16))) / 100
    recYCDODOS0.CDODOSDEV = mId$(MsgTxt, K + 67, 3)
    recYCDODOS0.CDODOSMOA = CCur(Val(mId$(MsgTxt, K + 70, 16))) / 100
    recYCDODOS0.CDODOSMOT = CCur(Val(mId$(MsgTxt, K + 86, 16))) / 100
    recYCDODOS0.CDODOSMOC = CCur(Val(mId$(MsgTxt, K + 102, 16))) / 100
    recYCDODOS0.CDODOSMOD = CCur(Val(mId$(MsgTxt, K + 118, 16))) / 100
    recYCDODOS0.CDODOSCON = mId$(MsgTxt, K + 134, 1)
    recYCDODOS0.CDODOSIRR = mId$(MsgTxt, K + 135, 1)
    recYCDODOS0.CDODOSFRA = mId$(MsgTxt, K + 136, 1)
    recYCDODOS0.CDODOSREN = mId$(MsgTxt, K + 137, 1)
    recYCDODOS0.CDODOSCUM = mId$(MsgTxt, K + 138, 1)
    recYCDODOS0.CDODOSTRS = mId$(MsgTxt, K + 139, 1)
    recYCDODOS0.CDODOSTOL = CCur(Val(mId$(MsgTxt, K + 140, 4))) / 100
    recYCDODOS0.CDODOSTO2 = CCur(Val(mId$(MsgTxt, K + 144, 4))) / 100
    recYCDODOS0.CDODOSDOR = mId$(MsgTxt, K + 148, 1)
    recYCDODOS0.CDODOSDON = mId$(MsgTxt, K + 149, 7)
    recYCDODOS0.CDODOSDOE = mId$(MsgTxt, K + 156, 64)
    recYCDODOS0.CDODOSBER = mId$(MsgTxt, K + 220, 1)
    recYCDODOS0.CDODOSBEN = mId$(MsgTxt, K + 221, 7)
    recYCDODOS0.CDODOSBEI = mId$(MsgTxt, K + 228, 64)
    recYCDODOS0.CDODOSBAR = mId$(MsgTxt, K + 292, 1)
    recYCDODOS0.CDODOSBAB = mId$(MsgTxt, K + 293, 7)
    recYCDODOS0.CDODOSNOR = mId$(MsgTxt, K + 300, 1)
    recYCDODOS0.CDODOSNOT = mId$(MsgTxt, K + 301, 7)
    recYCDODOS0.CDODOSBIC = mId$(MsgTxt, K + 308, 12)
    recYCDODOS0.CDODOSCOT = mId$(MsgTxt, K + 320, 1)
    recYCDODOS0.CDODOSCOR = mId$(MsgTxt, K + 321, 7)
    recYCDODOS0.CDODOSPRT = mId$(MsgTxt, K + 328, 1)
    recYCDODOS0.CDODOSPRR = mId$(MsgTxt, K + 329, 7)
    recYCDODOS0.CDODOSUTV = mId$(MsgTxt, K + 336, 32)
    recYCDODOS0.CDODOSPAT = mId$(MsgTxt, K + 368, 1)
    recYCDODOS0.CDODOSPAR = mId$(MsgTxt, K + 369, 7)
    recYCDODOS0.CDODOSPAV = mId$(MsgTxt, K + 376, 32)
    recYCDODOS0.CDODOSOUV = CLng(Val(mId$(MsgTxt, K + 408, 8)))
    recYCDODOS0.CDODOSEMI = CLng(Val(mId$(MsgTxt, K + 416, 8)))
    recYCDODOS0.CDODOSVAL = CLng(Val(mId$(MsgTxt, K + 424, 8)))
    recYCDODOS0.CDODOSDEP = CLng(Val(mId$(MsgTxt, K + 432, 8)))
    recYCDODOS0.CDODOSDTR = CLng(Val(mId$(MsgTxt, K + 440, 8)))
    recYCDODOS0.CDODOSVCP = CLng(Val(mId$(MsgTxt, K + 448, 8)))
    recYCDODOS0.CDODOSCLO = CLng(Val(mId$(MsgTxt, K + 456, 8)))
    recYCDODOS0.CDODOSREJ = mId$(MsgTxt, K + 464, 3)
    recYCDODOS0.CDODOSOBJ = mId$(MsgTxt, K + 467, 6)
    recYCDODOS0.CDODOSAVU = CLng(Val(mId$(MsgTxt, K + 473, 4)))
    recYCDODOS0.CDODOSMOV = CCur(Val(mId$(MsgTxt, K + 477, 16))) / 100
    recYCDODOS0.CDODOSCAC = CLng(Val(mId$(MsgTxt, K + 493, 4)))
    recYCDODOS0.CDODOSMCA = CCur(Val(mId$(MsgTxt, K + 497, 16))) / 100
    recYCDODOS0.CDODOSDIF = CLng(Val(mId$(MsgTxt, K + 513, 4)))
    recYCDODOS0.CDODOSMDI = CCur(Val(mId$(MsgTxt, K + 517, 16))) / 100
    recYCDODOS0.CDODOSPMO = CCur(Val(mId$(MsgTxt, K + 533, 16))) / 100
    recYCDODOS0.CDODOSPCD = mId$(MsgTxt, K + 549, 20)
    recYCDODOS0.CDODOSPCC = mId$(MsgTxt, K + 569, 20)
    recYCDODOS0.CDODOSPDE = CCur(Val(mId$(MsgTxt, K + 589, 16))) / 100
    recYCDODOS0.CDODOSPPO = CLng(Val(mId$(MsgTxt, K + 605, 4)))
    recYCDODOS0.CDODOSAUT = mId$(MsgTxt, K + 609, 12)
    recYCDODOS0.CDODOSREG = CCur(Val(mId$(MsgTxt, K + 621, 16))) / 100
    recYCDODOS0.CDODOSENC = CCur(Val(mId$(MsgTxt, K + 637, 16))) / 100
    recYCDODOS0.CDODOSDAN = CLng(Val(mId$(MsgTxt, K + 653, 8)))
    recYCDODOS0.CDODOSANN = CCur(Val(mId$(MsgTxt, K + 661, 16))) / 100
    recYCDODOS0.CDODOSPCO = CDbl(Val(mId$(MsgTxt, K + 677, 15))) / 1000000000
    recYCDODOS0.CDODOSLEM = mId$(MsgTxt, K + 692, 30)
    recYCDODOS0.CDODOSLDE = mId$(MsgTxt, K + 722, 30)
    recYCDODOS0.CDODOSDLE = CLng(Val(mId$(MsgTxt, K + 752, 8)))
    recYCDODOS0.CDODOSEPA = mId$(MsgTxt, K + 760, 1)
    recYCDODOS0.CDODOSTRA = mId$(MsgTxt, K + 761, 1)
    recYCDODOS0.CDODOSFCD = mId$(MsgTxt, K + 762, 1)
    recYCDODOS0.CDODOSCUS = CInt(Val(mId$(MsgTxt, K + 763, 5)))
    recYCDODOS0.CDODOSCUV = CInt(Val(mId$(MsgTxt, K + 768, 5)))
    recYCDODOS0.CDODOSCU2 = CInt(Val(mId$(MsgTxt, K + 773, 5)))
    recYCDODOS0.CDODOSOPE = mId$(MsgTxt, K + 778, 1)
    recYCDODOS0.CDODOSPOO = mId$(MsgTxt, K + 779, 1)
    recYCDODOS0.CDODOSPBE = CCur(Val(mId$(MsgTxt, K + 780, 16))) / 100
    recYCDODOS0.CDODOSGAG = mId$(MsgTxt, K + 796, 1)
    recYCDODOS0.CDODOSSTB = mId$(MsgTxt, K + 797, 1)
    recYCDODOS0.CDODOSMRE = mId$(MsgTxt, K + 798, 3)
    recYCDODOS0.CDODOSNPD = CLng(Val(mId$(MsgTxt, K + 801, 4)))
    recYCDODOS0.CDODOSTJD = mId$(MsgTxt, K + 805, 1)
    recYCDODOS0.CDODOSPDO = mId$(MsgTxt, K + 806, 60)
    recYCDODOS0.CDODOSGAR = mId$(MsgTxt, K + 866, 64)
    recYCDODOS0.CDODOSOBM = mId$(MsgTxt, K + 930, 64)
    recYCDODOS0.CDODOSTBR = mId$(MsgTxt, K + 994, 1)
    recYCDODOS0.CDODOSBRE = mId$(MsgTxt, K + 995, 7)
    recYCDODOS0.CDODOSBEC = mId$(MsgTxt, K + 1002, 1)
    recYCDODOS0.CDODOSRNO = mId$(MsgTxt, K + 1003, 16)
    recYCDODOS0.CDODOSDPA = mId$(MsgTxt, K + 1019, 3)
    recYCDODOS0.CDODOSDVI = mId$(MsgTxt, K + 1022, 32)
    recYCDODOS0.CDODOSEPY = mId$(MsgTxt, K + 1054, 3)
    recYCDODOS0.CDODOSEVI = mId$(MsgTxt, K + 1057, 32)
    recYCDODOS0.CDODOSVPA = mId$(MsgTxt, K + 1089, 3)
    recYCDODOS0.CDODOSVVI = mId$(MsgTxt, K + 1092, 32)
    recYCDODOS0.CDODOSNDE = CLng(Val(mId$(MsgTxt, K + 1124, 10)))
    recYCDODOS0.CDODOSNAE = mId$(MsgTxt, K + 1134, 3)
    recYCDODOS0.CDODOSEVE = mId$(MsgTxt, K + 1137, 2)
    recYCDODOS0.CDODOSETA = mId$(MsgTxt, K + 1139, 2)
    recYCDODOS0.CDODOSDP2 = mId$(MsgTxt, K + 1141, 32)
    recYCDODOS0.CDODOSEP2 = mId$(MsgTxt, K + 1173, 32)
    recYCDODOS0.CDODOSPD2 = mId$(MsgTxt, K + 1205, 80)
    recYCDODOS0.CDODOSAUN = mId$(MsgTxt, K + 1285, 12)
    recYCDODOS0.CDODOSCER = mId$(MsgTxt, K + 1297, 1)
Else
    srvYCDODOS0_GetBuffer = recYCDODOS0.Err
End If

MsgTxtIndex = MsgTxtIndex + recYCDODOS0Len

End Function

'---------------------------------------------------------
Public Function srvYCDODOS0_GetBuffer_ODBC(rsADO As ADODB.Recordset, recYCDODOS0 As typeYCDODOS0)
'---------------------------------------------------------
On Error GoTo Error_Handler
srvYCDODOS0_GetBuffer_ODBC = Null
    recYCDODOS0.CDODOSETB = rsADO("CDODOSETB")   ' CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYCDODOS0.CDODOSAGE = rsADO("CDODOSAGE")   ' CInt(Val(mId$(MsgTxt, K + 6, 5)))
    recYCDODOS0.CDODOSSER = rsADO("CDODOSSER")   ' mId$(MsgTxt, K + 11, 2)
    recYCDODOS0.CDODOSSSE = rsADO("CDODOSSSE")   ' mId$(MsgTxt, K + 13, 2)
    recYCDODOS0.CDODOSCOP = rsADO("CDODOSCOP")   ' mId$(MsgTxt, K + 15, 3)
    recYCDODOS0.CDODOSDOS = rsADO("CDODOSDOS")   ' CLng(Val(mId$(MsgTxt, K + 18, 10)))
    recYCDODOS0.CDODOSNUR = rsADO("CDODOSNUR")   ' CLng(Val(mId$(MsgTxt, K + 28, 4)))
    recYCDODOS0.CDODOSNAT = rsADO("CDODOSNAT")   ' mId$(MsgTxt, K + 32, 3)
    recYCDODOS0.CDODOSEXT = rsADO("CDODOSEXT")   ' mId$(MsgTxt, K + 35, 16)
    recYCDODOS0.CDODOSMON = rsADO("CDODOSMON")   ' CCur(Val(mId$(MsgTxt, K + 51, 16))) / 100
    recYCDODOS0.CDODOSDEV = rsADO("CDODOSDEV")   ' mId$(MsgTxt, K + 67, 3)
    recYCDODOS0.CDODOSMOA = rsADO("CDODOSMOA")   ' CCur(Val(mId$(MsgTxt, K + 70, 16))) / 100
    recYCDODOS0.CDODOSMOT = rsADO("CDODOSMOT")   ' CCur(Val(mId$(MsgTxt, K + 86, 16))) / 100
    recYCDODOS0.CDODOSMOC = rsADO("CDODOSMOC")   ' CCur(Val(mId$(MsgTxt, K + 102, 16))) / 100
    recYCDODOS0.CDODOSMOD = rsADO("CDODOSMOD")   ' CCur(Val(mId$(MsgTxt, K + 118, 16))) / 100
    recYCDODOS0.CDODOSCON = rsADO("CDODOSCON")   ' mId$(MsgTxt, K + 134, 1)
    recYCDODOS0.CDODOSIRR = rsADO("CDODOSIRR")   ' mId$(MsgTxt, K + 135, 1)
    recYCDODOS0.CDODOSFRA = rsADO("CDODOSFRA")   ' mId$(MsgTxt, K + 136, 1)
    recYCDODOS0.CDODOSREN = rsADO("CDODOSREN")   ' mId$(MsgTxt, K + 137, 1)
    recYCDODOS0.CDODOSCUM = rsADO("CDODOSCUM")   ' mId$(MsgTxt, K + 138, 1)
    recYCDODOS0.CDODOSTRS = rsADO("CDODOSTRS")   ' mId$(MsgTxt, K + 139, 1)
    recYCDODOS0.CDODOSTOL = rsADO("CDODOSTOL")   ' CCur(Val(mId$(MsgTxt, K + 140, 4))) / 100
    recYCDODOS0.CDODOSTO2 = rsADO("CDODOSTO2")   ' CCur(Val(mId$(MsgTxt, K + 144, 4))) / 100
    recYCDODOS0.CDODOSDOR = rsADO("CDODOSDOR")   ' mId$(MsgTxt, K + 148, 1)
    recYCDODOS0.CDODOSDON = rsADO("CDODOSDON")   ' mId$(MsgTxt, K + 149, 7)
    recYCDODOS0.CDODOSDOE = rsADO("CDODOSDOE")   ' mId$(MsgTxt, K + 156, 64)
    recYCDODOS0.CDODOSBER = rsADO("CDODOSBER")   ' mId$(MsgTxt, K + 220, 1)
    recYCDODOS0.CDODOSBEN = rsADO("CDODOSBEN")  ' mId$(MsgTxt, K + 221, 7)
    recYCDODOS0.CDODOSBEI = rsADO("CDODOSBEI")   ' mId$(MsgTxt, K + 228, 64)
    recYCDODOS0.CDODOSBAR = rsADO("CDODOSBAR")   ' mId$(MsgTxt, K + 292, 1)
    recYCDODOS0.CDODOSBAB = rsADO("CDODOSBAB")   ' mId$(MsgTxt, K + 293, 7)
    recYCDODOS0.CDODOSNOR = rsADO("CDODOSNOR")   ' mId$(MsgTxt, K + 300, 1)
    recYCDODOS0.CDODOSNOT = rsADO("CDODOSNOT")   ' mId$(MsgTxt, K + 301, 7)
    recYCDODOS0.CDODOSBIC = rsADO("CDODOSBIC")   ' mId$(MsgTxt, K + 308, 12)
    recYCDODOS0.CDODOSCOT = rsADO("CDODOSCOT")   ' mId$(MsgTxt, K + 320, 1)
    recYCDODOS0.CDODOSCOR = rsADO("CDODOSCOR")   ' mId$(MsgTxt, K + 321, 7)
    recYCDODOS0.CDODOSPRT = rsADO("CDODOSPRT")   ' mId$(MsgTxt, K + 328, 1)
    recYCDODOS0.CDODOSPRR = rsADO("CDODOSPRR")   ' mId$(MsgTxt, K + 329, 7)
    recYCDODOS0.CDODOSUTV = rsADO("CDODOSUTV")   ' mId$(MsgTxt, K + 336, 32)
    recYCDODOS0.CDODOSPAT = rsADO("CDODOSPAT")   ' mId$(MsgTxt, K + 368, 1)
    recYCDODOS0.CDODOSPAR = rsADO("CDODOSPAR")   ' mId$(MsgTxt, K + 369, 7)
    recYCDODOS0.CDODOSPAV = rsADO("CDODOSPAV")   ' mId$(MsgTxt, K + 376, 32)
    recYCDODOS0.CDODOSOUV = rsADO("CDODOSOUV")   ' CLng(Val(mId$(MsgTxt, K + 408, 8)))
    recYCDODOS0.CDODOSEMI = rsADO("CDODOSEMI")   ' CLng(Val(mId$(MsgTxt, K + 416, 8)))
    recYCDODOS0.CDODOSVAL = rsADO("CDODOSVAL")   ' CLng(Val(mId$(MsgTxt, K + 424, 8)))
    recYCDODOS0.CDODOSDEP = rsADO("CDODOSDEP")   ' CLng(Val(mId$(MsgTxt, K + 432, 8)))
    recYCDODOS0.CDODOSDTR = rsADO("CDODOSDTR")   ' CLng(Val(mId$(MsgTxt, K + 440, 8)))
    recYCDODOS0.CDODOSVCP = rsADO("CDODOSVCP")   ' CLng(Val(mId$(MsgTxt, K + 448, 8)))
    recYCDODOS0.CDODOSCLO = rsADO("CDODOSCLO")   ' CLng(Val(mId$(MsgTxt, K + 456, 8)))
    recYCDODOS0.CDODOSREJ = rsADO("CDODOSREJ")   ' mId$(MsgTxt, K + 464, 3)
    recYCDODOS0.CDODOSOBJ = rsADO("CDODOSOBJ")   ' mId$(MsgTxt, K + 467, 6)
    recYCDODOS0.CDODOSAVU = rsADO("CDODOSAVU")   ' CLng(Val(mId$(MsgTxt, K + 473, 4)))
    recYCDODOS0.CDODOSMOV = rsADO("CDODOSMOV")   ' CCur(Val(mId$(MsgTxt, K + 477, 16))) / 100
    recYCDODOS0.CDODOSCAC = rsADO("CDODOSCAC")   ' CLng(Val(mId$(MsgTxt, K + 493, 4)))
    recYCDODOS0.CDODOSMCA = rsADO("CDODOSMCA")   ' CCur(Val(mId$(MsgTxt, K + 497, 16))) / 100
    recYCDODOS0.CDODOSDIF = rsADO("CDODOSDIF")   ' CLng(Val(mId$(MsgTxt, K + 513, 4)))
    recYCDODOS0.CDODOSMDI = rsADO("CDODOSMDI")   ' CCur(Val(mId$(MsgTxt, K + 517, 16))) / 100
    recYCDODOS0.CDODOSPMO = rsADO("CDODOSPMO")   ' CCur(Val(mId$(MsgTxt, K + 533, 16))) / 100
    recYCDODOS0.CDODOSPCD = rsADO("CDODOSPCD")   ' mId$(MsgTxt, K + 549, 20)
    recYCDODOS0.CDODOSPCC = rsADO("CDODOSPCC")   ' mId$(MsgTxt, K + 569, 20)
    recYCDODOS0.CDODOSPDE = rsADO("CDODOSPDE")   ' CCur(Val(mId$(MsgTxt, K + 589, 16))) / 100
    recYCDODOS0.CDODOSPPO = rsADO("CDODOSPPO")   ' CLng(Val(mId$(MsgTxt, K + 605, 4)))
    recYCDODOS0.CDODOSAUT = rsADO("CDODOSAUT")   ' mId$(MsgTxt, K + 609, 12)
    recYCDODOS0.CDODOSREG = rsADO("CDODOSREG")   ' CCur(Val(mId$(MsgTxt, K + 621, 16))) / 100
    recYCDODOS0.CDODOSENC = rsADO("CDODOSENC")   ' CCur(Val(mId$(MsgTxt, K + 637, 16))) / 100
    recYCDODOS0.CDODOSDAN = rsADO("CDODOSDAN")   ' CLng(Val(mId$(MsgTxt, K + 653, 8)))
    recYCDODOS0.CDODOSANN = rsADO("CDODOSANN")   ' CCur(Val(mId$(MsgTxt, K + 661, 16))) / 100
    recYCDODOS0.CDODOSPCO = rsADO("CDODOSPCO")   ' CDbl(Val(mId$(MsgTxt, K + 677, 15))) / 1000000000
    recYCDODOS0.CDODOSLEM = rsADO("CDODOSLEM") ' mId$(MsgTxt, K + 692, 30)
    recYCDODOS0.CDODOSLDE = rsADO("CDODOSLDE")   ' mId$(MsgTxt, K + 722, 30)
    recYCDODOS0.CDODOSDLE = rsADO("CDODOSDLE")   ' CLng(Val(mId$(MsgTxt, K + 752, 8)))
    recYCDODOS0.CDODOSEPA = rsADO("CDODOSEPA")   ' mId$(MsgTxt, K + 760, 1)
    recYCDODOS0.CDODOSTRA = rsADO("CDODOSTRA")   ' mId$(MsgTxt, K + 761, 1)
    recYCDODOS0.CDODOSFCD = rsADO("CDODOSFCD")   ' mId$(MsgTxt, K + 762, 1)
    recYCDODOS0.CDODOSCUS = rsADO("CDODOSCUS")   ' CInt(Val(mId$(MsgTxt, K + 763, 5)))
    recYCDODOS0.CDODOSCUV = rsADO("CDODOSCUV")   ' CInt(Val(mId$(MsgTxt, K + 768, 5)))
    recYCDODOS0.CDODOSCU2 = rsADO("CDODOSCU2")   ' CInt(Val(mId$(MsgTxt, K + 773, 5)))
    recYCDODOS0.CDODOSOPE = rsADO("CDODOSOPE")   ' mId$(MsgTxt, K + 778, 1)
    recYCDODOS0.CDODOSPOO = rsADO("CDODOSPOO")   ' mId$(MsgTxt, K + 779, 1)
    recYCDODOS0.CDODOSPBE = rsADO("CDODOSPBE")   ' CCur(Val(mId$(MsgTxt, K + 780, 16))) / 100
    recYCDODOS0.CDODOSGAG = rsADO("CDODOSGAG")   ' mId$(MsgTxt, K + 796, 1)
    recYCDODOS0.CDODOSSTB = rsADO("CDODOSSTB")   ' mId$(MsgTxt, K + 797, 1)
    recYCDODOS0.CDODOSMRE = rsADO("CDODOSMRE")   ' mId$(MsgTxt, K + 798, 3)
    recYCDODOS0.CDODOSNPD = rsADO("CDODOSNPD")   ' CLng(Val(mId$(MsgTxt, K + 801, 4)))
    recYCDODOS0.CDODOSTJD = rsADO("CDODOSTJD")   ' mId$(MsgTxt, K + 805, 1)
    recYCDODOS0.CDODOSPDO = rsADO("CDODOSPDO")   ' mId$(MsgTxt, K + 806, 60)
    recYCDODOS0.CDODOSGAR = rsADO("CDODOSGAR")   ' mId$(MsgTxt, K + 866, 64)
    recYCDODOS0.CDODOSOBM = rsADO("CDODOSOBM")   ' mId$(MsgTxt, K + 930, 64)
    recYCDODOS0.CDODOSTBR = rsADO("CDODOSTBR")   ' mId$(MsgTxt, K + 994, 1)
    recYCDODOS0.CDODOSBRE = rsADO("CDODOSBRE")   ' mId$(MsgTxt, K + 995, 7)
    recYCDODOS0.CDODOSBEC = rsADO("CDODOSBEC")   ' mId$(MsgTxt, K + 1002, 1)
    recYCDODOS0.CDODOSRNO = rsADO("CDODOSRNO")   ' mId$(MsgTxt, K + 1003, 16)
    recYCDODOS0.CDODOSDPA = rsADO("CDODOSDPA")   ' mId$(MsgTxt, K + 1019, 3)
    recYCDODOS0.CDODOSDVI = rsADO("CDODOSDVI")   ' mId$(MsgTxt, K + 1022, 32)
    recYCDODOS0.CDODOSEPY = rsADO("CDODOSEPY") ' mId$(MsgTxt, K + 1054, 3)
    recYCDODOS0.CDODOSEVI = rsADO("CDODOSEVI")   ' mId$(MsgTxt, K + 1057, 32)
    recYCDODOS0.CDODOSVPA = rsADO("CDODOSVPA")   ' mId$(MsgTxt, K + 1089, 3)
    recYCDODOS0.CDODOSVVI = rsADO("CDODOSVVI")   ' mId$(MsgTxt, K + 1092, 32)
    recYCDODOS0.CDODOSNDE = rsADO("CDODOSNDE")   ' CLng(Val(mId$(MsgTxt, K + 1124, 10)))
    recYCDODOS0.CDODOSNAE = rsADO("CDODOSNAE")   ' mId$(MsgTxt, K + 1134, 3)
    recYCDODOS0.CDODOSEVE = rsADO("CDODOSEVE")   ' mId$(MsgTxt, K + 1137, 2)
    recYCDODOS0.CDODOSETA = rsADO("CDODOSETA")   ' mId$(MsgTxt, K + 1139, 2)
    recYCDODOS0.CDODOSDP2 = rsADO("CDODOSDP2")   ' mId$(MsgTxt, K + 1141, 32)
    recYCDODOS0.CDODOSEP2 = rsADO("CDODOSEP2")   ' mId$(MsgTxt, K + 1173, 32)
    recYCDODOS0.CDODOSPD2 = rsADO("CDODOSPD2")   ' mId$(MsgTxt, K + 1205, 80)
    recYCDODOS0.CDODOSAUN = rsADO("CDODOSAUN")   ' mId$(MsgTxt, K + 1285, 12)
    recYCDODOS0.CDODOSCER = rsADO("CDODOSCER")   ' mId$(MsgTxt, K + 1297, 1)
Exit Function

Error_Handler:
srvYCDODOS0_GetBuffer_ODBC = Error

End Function


'---------------------------------------------------------
Private Sub srvYCDODOS0_PutBuffer(recYCDODOS0 As typeYCDODOS0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, recYCDODOS0Len) = Space$(recYCDODOS0Len)

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYCDODOS0.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYCDODOS0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 5) = Format$(recYCDODOS0.CDODOSETB, "0000 ")
    Mid$(MsgTxt, K + 6, 5) = Format$(recYCDODOS0.CDODOSAGE, "0000 ")
    Mid$(MsgTxt, K + 11, 2) = recYCDODOS0.CDODOSSER
    Mid$(MsgTxt, K + 13, 2) = recYCDODOS0.CDODOSSSE
    Mid$(MsgTxt, K + 15, 3) = recYCDODOS0.CDODOSCOP
    Mid$(MsgTxt, K + 18, 10) = Format$(recYCDODOS0.CDODOSDOS, "000000000 ")
    Mid$(MsgTxt, K + 28, 4) = Format$(recYCDODOS0.CDODOSNUR, "000 ")
    Mid$(MsgTxt, K + 32, 3) = recYCDODOS0.CDODOSNAT
    Mid$(MsgTxt, K + 35, 16) = recYCDODOS0.CDODOSEXT
    Mid$(MsgTxt, K + 51, 16) = Format$(recYCDODOS0.CDODOSMON * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 67, 3) = recYCDODOS0.CDODOSDEV
    Mid$(MsgTxt, K + 70, 16) = Format$(recYCDODOS0.CDODOSMOA * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 86, 16) = Format$(recYCDODOS0.CDODOSMOT * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 102, 16) = Format$(recYCDODOS0.CDODOSMOC * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 118, 16) = Format$(recYCDODOS0.CDODOSMOD * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 134, 1) = recYCDODOS0.CDODOSCON
    Mid$(MsgTxt, K + 135, 1) = recYCDODOS0.CDODOSIRR
    Mid$(MsgTxt, K + 136, 1) = recYCDODOS0.CDODOSFRA
    Mid$(MsgTxt, K + 137, 1) = recYCDODOS0.CDODOSREN
    Mid$(MsgTxt, K + 138, 1) = recYCDODOS0.CDODOSCUM
    Mid$(MsgTxt, K + 139, 1) = recYCDODOS0.CDODOSTRS
    Mid$(MsgTxt, K + 140, 4) = Format$(recYCDODOS0.CDODOSTOL * 100, "000 ")
    Mid$(MsgTxt, K + 144, 4) = Format$(recYCDODOS0.CDODOSTO2 * 100, "000 ")
    Mid$(MsgTxt, K + 148, 1) = recYCDODOS0.CDODOSDOR
    Mid$(MsgTxt, K + 149, 7) = recYCDODOS0.CDODOSDON
    Mid$(MsgTxt, K + 156, 64) = recYCDODOS0.CDODOSDOE
    Mid$(MsgTxt, K + 220, 1) = recYCDODOS0.CDODOSBER
    Mid$(MsgTxt, K + 221, 7) = recYCDODOS0.CDODOSBEN
    Mid$(MsgTxt, K + 228, 64) = recYCDODOS0.CDODOSBEI
    Mid$(MsgTxt, K + 292, 1) = recYCDODOS0.CDODOSBAR
    Mid$(MsgTxt, K + 293, 7) = recYCDODOS0.CDODOSBAB
    Mid$(MsgTxt, K + 300, 1) = recYCDODOS0.CDODOSNOR
    Mid$(MsgTxt, K + 301, 7) = recYCDODOS0.CDODOSNOT
    Mid$(MsgTxt, K + 308, 12) = recYCDODOS0.CDODOSBIC
    Mid$(MsgTxt, K + 320, 1) = recYCDODOS0.CDODOSCOT
    Mid$(MsgTxt, K + 321, 7) = recYCDODOS0.CDODOSCOR
    Mid$(MsgTxt, K + 328, 1) = recYCDODOS0.CDODOSPRT
    Mid$(MsgTxt, K + 329, 7) = recYCDODOS0.CDODOSPRR
    Mid$(MsgTxt, K + 336, 32) = recYCDODOS0.CDODOSUTV
    Mid$(MsgTxt, K + 368, 1) = recYCDODOS0.CDODOSPAT
    Mid$(MsgTxt, K + 369, 7) = recYCDODOS0.CDODOSPAR
    Mid$(MsgTxt, K + 376, 32) = recYCDODOS0.CDODOSPAV
    Mid$(MsgTxt, K + 408, 8) = Format$(recYCDODOS0.CDODOSOUV, "0000000 ")
    Mid$(MsgTxt, K + 416, 8) = Format$(recYCDODOS0.CDODOSEMI, "0000000 ")
    Mid$(MsgTxt, K + 424, 8) = Format$(recYCDODOS0.CDODOSVAL, "0000000 ")
    Mid$(MsgTxt, K + 432, 8) = Format$(recYCDODOS0.CDODOSDEP, "0000000 ")
    Mid$(MsgTxt, K + 440, 8) = Format$(recYCDODOS0.CDODOSDTR, "0000000 ")
    Mid$(MsgTxt, K + 448, 8) = Format$(recYCDODOS0.CDODOSVCP, "0000000 ")
    Mid$(MsgTxt, K + 456, 8) = Format$(recYCDODOS0.CDODOSCLO, "0000000 ")
    Mid$(MsgTxt, K + 464, 3) = recYCDODOS0.CDODOSREJ
    Mid$(MsgTxt, K + 467, 6) = recYCDODOS0.CDODOSOBJ
    Mid$(MsgTxt, K + 473, 4) = Format$(recYCDODOS0.CDODOSAVU, "000 ")
    Mid$(MsgTxt, K + 477, 16) = Format$(recYCDODOS0.CDODOSMOV * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 493, 4) = Format$(recYCDODOS0.CDODOSCAC, "000 ")
    Mid$(MsgTxt, K + 497, 16) = Format$(recYCDODOS0.CDODOSMCA * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 513, 4) = Format$(recYCDODOS0.CDODOSDIF, "000 ")
    Mid$(MsgTxt, K + 517, 16) = Format$(recYCDODOS0.CDODOSMDI * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 533, 16) = Format$(recYCDODOS0.CDODOSPMO * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 549, 20) = recYCDODOS0.CDODOSPCD
    Mid$(MsgTxt, K + 569, 20) = recYCDODOS0.CDODOSPCC
    Mid$(MsgTxt, K + 589, 16) = Format$(recYCDODOS0.CDODOSPDE * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 605, 4) = Format$(recYCDODOS0.CDODOSPPO, "000 ")
    Mid$(MsgTxt, K + 609, 12) = recYCDODOS0.CDODOSAUT
    Mid$(MsgTxt, K + 621, 16) = Format$(recYCDODOS0.CDODOSREG * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 637, 16) = Format$(recYCDODOS0.CDODOSENC * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 653, 8) = Format$(recYCDODOS0.CDODOSDAN, "0000000 ")
    Mid$(MsgTxt, K + 661, 16) = Format$(recYCDODOS0.CDODOSANN * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 677, 15) = Format$(recYCDODOS0.CDODOSPCO * 1000000000, "00000000000000 ")
    Mid$(MsgTxt, K + 692, 30) = recYCDODOS0.CDODOSLEM
    Mid$(MsgTxt, K + 722, 30) = recYCDODOS0.CDODOSLDE
    Mid$(MsgTxt, K + 752, 8) = Format$(recYCDODOS0.CDODOSDLE, "0000000 ")
    Mid$(MsgTxt, K + 760, 1) = recYCDODOS0.CDODOSEPA
    Mid$(MsgTxt, K + 761, 1) = recYCDODOS0.CDODOSTRA
    Mid$(MsgTxt, K + 762, 1) = recYCDODOS0.CDODOSFCD
    Mid$(MsgTxt, K + 763, 5) = Format$(recYCDODOS0.CDODOSCUS, "0000 ")
    Mid$(MsgTxt, K + 768, 5) = Format$(recYCDODOS0.CDODOSCUV, "0000 ")
    Mid$(MsgTxt, K + 773, 5) = Format$(recYCDODOS0.CDODOSCU2, "0000 ")
    Mid$(MsgTxt, K + 778, 1) = recYCDODOS0.CDODOSOPE
    Mid$(MsgTxt, K + 779, 1) = recYCDODOS0.CDODOSPOO
    Mid$(MsgTxt, K + 780, 16) = Format$(recYCDODOS0.CDODOSPBE * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 796, 1) = recYCDODOS0.CDODOSGAG
    Mid$(MsgTxt, K + 797, 1) = recYCDODOS0.CDODOSSTB
    Mid$(MsgTxt, K + 798, 3) = recYCDODOS0.CDODOSMRE
    Mid$(MsgTxt, K + 801, 4) = Format$(recYCDODOS0.CDODOSNPD, "000 ")
    Mid$(MsgTxt, K + 805, 1) = recYCDODOS0.CDODOSTJD
    Mid$(MsgTxt, K + 806, 60) = recYCDODOS0.CDODOSPDO
    Mid$(MsgTxt, K + 866, 64) = recYCDODOS0.CDODOSGAR
    Mid$(MsgTxt, K + 930, 64) = recYCDODOS0.CDODOSOBM
    Mid$(MsgTxt, K + 994, 1) = recYCDODOS0.CDODOSTBR
    Mid$(MsgTxt, K + 995, 7) = recYCDODOS0.CDODOSBRE
    Mid$(MsgTxt, K + 1002, 1) = recYCDODOS0.CDODOSBEC
    Mid$(MsgTxt, K + 1003, 16) = recYCDODOS0.CDODOSRNO
    Mid$(MsgTxt, K + 1019, 3) = recYCDODOS0.CDODOSDPA
    Mid$(MsgTxt, K + 1022, 32) = recYCDODOS0.CDODOSDVI
    Mid$(MsgTxt, K + 1054, 3) = recYCDODOS0.CDODOSEPY
    Mid$(MsgTxt, K + 1057, 32) = recYCDODOS0.CDODOSEVI
    Mid$(MsgTxt, K + 1089, 3) = recYCDODOS0.CDODOSVPA
    Mid$(MsgTxt, K + 1092, 32) = recYCDODOS0.CDODOSVVI
    Mid$(MsgTxt, K + 1124, 10) = Format$(recYCDODOS0.CDODOSNDE, "000000000 ")
    Mid$(MsgTxt, K + 1134, 3) = recYCDODOS0.CDODOSNAE
    Mid$(MsgTxt, K + 1137, 2) = recYCDODOS0.CDODOSEVE
    Mid$(MsgTxt, K + 1139, 2) = recYCDODOS0.CDODOSETA
    Mid$(MsgTxt, K + 1141, 32) = recYCDODOS0.CDODOSDP2
    Mid$(MsgTxt, K + 1173, 32) = recYCDODOS0.CDODOSEP2
    Mid$(MsgTxt, K + 1205, 80) = recYCDODOS0.CDODOSPD2
    Mid$(MsgTxt, K + 1285, 12) = recYCDODOS0.CDODOSAUN
    Mid$(MsgTxt, K + 1297, 1) = recYCDODOS0.CDODOSCER
MsgTxtLen = MsgTxtLen + recYCDODOS0Len
End Sub



'---------------------------------------------------------
Private Function srvYCDODOS0_Seek(recYCDODOS0 As typeYCDODOS0)
'---------------------------------------------------------

srvYCDODOS0_Seek = "?"
MsgTxtLen = 0
Call srvYCDODOS0_PutBuffer(recYCDODOS0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If MsgTxtLen > 0 Then
        If IsNull(srvYCDODOS0_GetBuffer(recYCDODOS0)) Then
            srvYCDODOS0_Seek = Null
        Else
            Call srvYCDODOS0_Error(recYCDODOS0)
        End If
    End If
End If

End Function

'-----------------------------------------------------
Function srvYCDODOS0_Update(recYCDODOS0 As typeYCDODOS0)
'-----------------------------------------------------

srvYCDODOS0_Update = "?"

MsgTxtLen = 0
Call srvYCDODOS0_PutBuffer(recYCDODOS0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYCDODOS0_GetBuffer(recYCDODOS0)) Then
        Call srvYCDODOS0_Error(recYCDODOS0)
        srvYCDODOS0_Update = recYCDODOS0.Err
        Exit Function
    Else
        srvYCDODOS0_Update = Null
    End If
Else
    recYCDODOS0.Err = "srv"
End If


'=====================================================
End Function



'---------------------------------------------------------
Public Sub recYCDODOS0_Init(recYCDODOS0 As typeYCDODOS0)
'---------------------------------------------------------
MsgTxt = Space$(recYCDODOS0Len)
MsgTxtIndex = 0
Call srvYCDODOS0_GetBuffer(recYCDODOS0)
recYCDODOS0.obj = "ZCDODOS0_S"
recYCDODOS0.CDODOSETB = 1
recYCDODOS0.CDODOSAGE = 1
recYCDODOS0.CDODOSSER = "00"
recYCDODOS0.CDODOSSSE = "00"
recYCDODOS0.CDODOSETB = 0      'As Integer                        ' CODE ETABLISSEMENT
recYCDODOS0.CDODOSAGE = 0      'As Integer                        ' AGENCE
recYCDODOS0.CDODOSSER = ""      'As String * 2                     ' SERVICE
recYCDODOS0.CDODOSSSE = ""      'As String * 2                     ' SOUS-SERVICE
recYCDODOS0.CDODOSCOP = ""      'As String * 3                     ' CODE OPERATION
recYCDODOS0.CDODOSDOS = 0      'As Long                           ' NUMERO DOSSIER
recYCDODOS0.CDODOSNUR = 0      'As Long                           ' N° RENOUVELLEMENT
recYCDODOS0.CDODOSNAT = ""      'As String * 3                     ' NATURE
recYCDODOS0.CDODOSEXT = ""      'As String * 16                    ' REFERENCE EXTERNE
recYCDODOS0.CDODOSMON = 0      'As Currency                       ' MONTANT DOSSIER
recYCDODOS0.CDODOSDEV = ""      'As String * 3                     ' DEVISE
recYCDODOS0.CDODOSMOA = 0      'As Currency                       ' MONTANT ADDITIONNEL
recYCDODOS0.CDODOSMOT = 0      'As Currency                       ' MONTANT TOTAL
recYCDODOS0.CDODOSMOC = 0      'As Currency                       ' MONTANT CONFIRME
recYCDODOS0.CDODOSMOD = 0      'As Currency                       ' MONTANT DUCROIRE
recYCDODOS0.CDODOSCON = ""      'As String * 1                     ' CONFIRM NOTIFI PARTI
recYCDODOS0.CDODOSIRR = ""      'As String * 1                     ' IRREVOCABLE (O/N)
recYCDODOS0.CDODOSFRA = ""      'As String * 1                     ' FRACTIONNABLE (O/N)
recYCDODOS0.CDODOSREN = ""      'As String * 1                     ' RENOUVELABLE (O/N)
recYCDODOS0.CDODOSCUM = ""      'As String * 1                     ' CUMULATIF (O/N)
recYCDODOS0.CDODOSTRS = ""      'As String * 1                     ' TRANSFERABLE
recYCDODOS0.CDODOSTOL = 0      'As Currency                       ' TOLERANCE +
recYCDODOS0.CDODOSTO2 = 0      'As Currency                       ' TOLERANCE -
recYCDODOS0.CDODOSDOR = ""      'As String * 1                     ' DONN. ORDRE CLI/TIE
recYCDODOS0.CDODOSDON = ""      'As String * 7                     ' DONNEUR ORDRE IMPORT
recYCDODOS0.CDODOSDOE = ""      'As String * 64                    ' DONNEUR ORDRE EXPORT
recYCDODOS0.CDODOSBER = ""      'As String * 1                     ' BENEFICIAIR CLI/TIE
recYCDODOS0.CDODOSBEN = ""      'As String * 7                     ' BENEFICIAIRE EXPORT
recYCDODOS0.CDODOSBEI = ""      'As String * 64                    ' BENEFICIAIRE IMPORT
recYCDODOS0.CDODOSBAR = ""      'As String * 1                     ' BANQU.BENEF.CLI/TIE
recYCDODOS0.CDODOSBAB = ""      'As String * 7                     ' BANQUE BENEF
recYCDODOS0.CDODOSNOR = ""      'As String * 1                     ' NOTIF/CONFI OU EMETT
recYCDODOS0.CDODOSNOT = ""      'As String * 7                     ' NOTIF/CONFI OU EMETT
recYCDODOS0.CDODOSBIC = ""      'As String * 12                    ' BIC SUPPLEMEN. EMETT
recYCDODOS0.CDODOSCOT = ""      'As String * 1                     ' CORRESPOND. CLI/TIE
recYCDODOS0.CDODOSCOR = ""      'As String * 7                     ' CORRESPONDANT
recYCDODOS0.CDODOSPRT = ""      'As String * 1                     ' LIEU PRES CLI/TIE
recYCDODOS0.CDODOSPRR = ""      'As String * 7                     ' LIEU PRESENTATION
recYCDODOS0.CDODOSUTV = ""      'As String * 32                    ' LIEU PRESENTATION
recYCDODOS0.CDODOSPAT = ""      'As String * 1                     ' LIEU PAIE CLI/TIE
recYCDODOS0.CDODOSPAR = ""      'As String * 7                     ' LIEU PAIEMENT
recYCDODOS0.CDODOSPAV = ""      'As String * 32                    ' LIEU PAIEMENT
recYCDODOS0.CDODOSOUV = 0      'As Long                           ' DATE OUVERTURE
recYCDODOS0.CDODOSEMI = 0      'As Long                           ' DATE EMISSION
recYCDODOS0.CDODOSVAL = 0      'As Long                           ' DATE VALIDITE
recYCDODOS0.CDODOSDEP = 0      'As Long                           ' DATE EXTREME PAYMT
recYCDODOS0.CDODOSDTR = 0      'As Long                           ' DATE DE TRANSFERT
recYCDODOS0.CDODOSVCP = 0      'As Long                           ' DATE VALID. COMPTA
recYCDODOS0.CDODOSCLO = 0      'As Long                           ' DATE CLOTURE
recYCDODOS0.CDODOSREJ = ""      'As String * 3                     ' MOTIF REJET (CLOTUR)
recYCDODOS0.CDODOSOBJ = ""      'As String * 6                     ' OBJET CREDIT
recYCDODOS0.CDODOSAVU = 0      'As Long                           ' % PAIEM. A VUE
recYCDODOS0.CDODOSMOV = 0      'As Currency                       ' MONTANT A VUE
recYCDODOS0.CDODOSCAC = 0      'As Long                           ' % PAIEM. CTR ACCEPT.
recYCDODOS0.CDODOSMCA = 0      'As Currency                       ' MONTANT CTR ACCEPT.
recYCDODOS0.CDODOSDIF = 0      'As Long                           ' % PAIEM. DIFFERE
recYCDODOS0.CDODOSMDI = 0      'As Currency                       ' MONTANT. DIFFERE
recYCDODOS0.CDODOSPMO = 0      'As Currency                       ' MONTANT PROVISIONNE
recYCDODOS0.CDODOSPCD = ""      'As String * 20                    ' PROV. DEBIT  COMPTE
recYCDODOS0.CDODOSPCC = ""      'As String * 20                    ' PROV. CREDIT COMPTE
recYCDODOS0.CDODOSPDE = 0      'As Currency                       ' PROVISION DEVISE DOS
recYCDODOS0.CDODOSPPO = 0      'As Long                           ' PROVISION POURCEN
recYCDODOS0.CDODOSAUT = ""      'As String * 12                    ' CODE AUTORISATION
recYCDODOS0.CDODOSREG = 0      'As Currency                       ' MONTANT PAYE
recYCDODOS0.CDODOSENC = 0      'As Currency                       ' MONTANT ENCAISSE
recYCDODOS0.CDODOSDAN = 0      'As Long                           ' DATE ANNULATION
recYCDODOS0.CDODOSANN = 0      'As Currency                       ' MONTANT ANNULE
recYCDODOS0.CDODOSPCO = 0       '         As Double                         ' COURS DEVPRO/DEVDOS
recYCDODOS0.CDODOSLEM = ""      'As String * 30                    ' LIEU EMBARQUEMENT
recYCDODOS0.CDODOSLDE = ""      'As String * 30                    ' LIEU DESTINATION
recYCDODOS0.CDODOSDLE = 0      'As Long                           ' DATE LIMITE EMBARQU.
recYCDODOS0.CDODOSEPA = ""      'As String * 1                     ' EXPED.PARTIE.AUTORI
recYCDODOS0.CDODOSTRA = ""      'As String * 1                     ' TRANBORDEMENT AUTORI
recYCDODOS0.CDODOSFCD = ""      'As String * 1                     ' FRAI CHARGE D.O. BEN
recYCDODOS0.CDODOSCUS = 0      'As Integer                        ' UTILI. DE SAISIE
recYCDODOS0.CDODOSCUV = 0      'As Integer                        ' 1ER VALIDEUR
recYCDODOS0.CDODOSCU2 = 0      'As Integer                        ' 2EME VALIDEUR
recYCDODOS0.CDODOSOPE = ""      'As String * 1                     ' OPERATIVITE DU CRED.
recYCDODOS0.CDODOSPOO = ""      'As String * 1                     ' EXISTENCE POOL
recYCDODOS0.CDODOSPBE = 0      'As Currency                       ' PART.BANQUE EXPORT
recYCDODOS0.CDODOSGAG = ""      'As String * 1                     ' GAGE MARCHANDISE
recYCDODOS0.CDODOSSTB = ""      'As String * 1                     ' STAND BY
recYCDODOS0.CDODOSMRE = ""      'As String * 3                     ' MODE DE REALISAT°
recYCDODOS0.CDODOSNPD = 0      'As Long                           ' NBJ PRES. DOCUMENT
recYCDODOS0.CDODOSTJD = ""      'As String * 1                     ' TY JOUR DOCS
recYCDODOS0.CDODOSPDO = ""      'As String * 60                    ' PER.PRE.DOCS.
recYCDODOS0.CDODOSGAR = ""      'As String * 64                    ' LIBELLE GARANTIE
recYCDODOS0.CDODOSOBM = ""      'As String * 64                    ' OBJET DE MODIF.
recYCDODOS0.CDODOSTBR = ""      'As String * 1                     ' TIERS BQ REMBOURS
recYCDODOS0.CDODOSBRE = ""      'As String * 7                     ' BQ REMBOURSEMENT
recYCDODOS0.CDODOSBEC = ""      'As String * 1                     ' BENEF PAY.COMMIS°
recYCDODOS0.CDODOSRNO = ""      'As String * 16                    ' REF.NOTIFICATEUR
recYCDODOS0.CDODOSDPA = ""      'As String * 3                     ' DESTINATION PAYS
recYCDODOS0.CDODOSDVI = ""      'As String * 32                    ' DESTINATION VILLE
recYCDODOS0.CDODOSEPY = ""      'As String * 3                     ' EMBARQUEMENT PAYS
recYCDODOS0.CDODOSEVI = ""      'As String * 32                    ' EMBARQUEMENT VILLE
recYCDODOS0.CDODOSVPA = ""      'As String * 3                     ' VALIDITE PAYS
recYCDODOS0.CDODOSVVI = ""      'As String * 32                    ' VALIDIT VILLE
recYCDODOS0.CDODOSNDE = 0      'As Long                           ' DOSSIER EXPORT
recYCDODOS0.CDODOSNAE = ""      'As String * 3                     ' NATURE EXPORT
recYCDODOS0.CDODOSEVE = ""      'As String * 2                     ' EVENEMENT
recYCDODOS0.CDODOSETA = ""      'As String * 2                     ' ETAT DOSSIER
recYCDODOS0.CDODOSDP2 = ""      'As String * 32                    ' DESTIN.PAYS LIBELLE
recYCDODOS0.CDODOSEP2 = ""      'As String * 32                    ' EMBARQ.PAYS LIBELLE
recYCDODOS0.CDODOSPD2 = ""      'As String * 80                    ' PER.PRES.DOC.SUITE
recYCDODOS0.CDODOSAUN = ""      'As String * 12                    ' CODE AUT. NOTIFIE
recYCDODOS0.CDODOSCER = ""      'As String * 1                     ' COTAT°(O=CERTAIN/N)

End Sub


Public Function fctYCDODOS0_Compare(recYCDODOS0 As typeYCDODOS0, mYCDODOS0 As typeYCDODOS0)
fctYCDODOS0_Compare = Null
'If recYCDODOS0.IdRéférence <> mYCDODOS0.IdRéférence Then fctYCDODOS0_Compare = "IdRéférence": Exit Function
End Function



