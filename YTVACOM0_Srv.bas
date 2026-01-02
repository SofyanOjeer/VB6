Attribute VB_Name = "srvYTVACOM0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYTVACOM0
 
      TVACOMETA   As Long         'établissement
      TVACOMPLA   As Long         'plan
      TVACOMSER   As String * 2   'service
      TVACOMSSE   As String * 2   'sous-service
      TVACOMPIE   As Long         'pièce
      TVACOMECR   As Long         'écriture
      TVACOMCPT   As String * 20  'compte
      TVACOMDTR   As Long         'date de traitement
      TVACOMDVA   As Long         'date de valeur
      TVACOMOPE   As String * 3   'code opération
      TVACOMNAT   As String * 6   'nature opération
      TVACOMEVE   As String * 3   'événement
      TVACOMDOS   As Long         'N° dossier
      TVACOMDEV   As String * 3   'devise
      TVACOMMON   As Currency     'montant en dev
      TVACOMMONE  As Currency     'montant euro
      TVACOMMTVA  As Currency     'montant TVA issu du dossier
      TVACOMMTVE  As Currency     'montant TVA issu du dossier en CV €
      TVACOMCOMB  As String * 1   'code tiers facturé B S D
      TVACOMCOMC  As String * 6   'code commission
      TVACOMCOME  As String * 1   'code édition espace,W,A
      TVACOMCOMT  As String * 1   'taxable Exonéré,Normal,Réduit
      TVACOMCLIC  As String * 1   'table client espace,G, D
      TVACOMCLI   As String * 7   'code client
      TVACOMCLIP  As String * 2   'pays de résidence
      TVACOMTVAC  As String * 1   'code TVA
      TVACOMFACN  As Long         'n° facture
      TVACOMFACL  As Long         'n° facture liée (N° origine si avoir)
      TVACOMSRVR  As String * 4   'service responsable
      TVACOMQTE   As Long         'quantité
      TVACOMSTA   As String * 1   'statut
      TVACOMUPDS  As Long         'SéQUENCE UPD
      TVACOMUSR   As String * 10   'user
      
      TVACOMXNUR   As Long         'n° renouvellement
      TVACOMXUTI   As Long         'n° utilisation
      TVACOMXEVE   As String * 2   'code événement comptable
      TVACOMXSEQ   As Long         'n° séquence
      TVACOMXSPE   As Long         'n° séquence périodique
      TVACOMECRX   As Long         'écriture contrepartie Client
      TVACOMGTYP   As String * 1   'lien ZCHGDET0
      TVACOMGORD   As String * 1   'lien ZCHGDET0
      
'_________________________________________________________________________ pour gestion interne
      TVACOMAVOIR  As String * 1   ' pour gestion avoir
      'xxxxDenis 26/04/2010
      'TVAREFCLI    As String * 10  'Référence client
      TVAREFCLI    As String * 16  'Référence client

End Type
Public xYTVACOM0 As typeYTVACOM0
Public Function sqlYTVACOM0_Update(newY As typeYTVACOM0, oldY As typeYTVACOM0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYTVACOM0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.TVACOMETA <> newY.TVACOMETA _
Or oldY.TVACOMPLA <> newY.TVACOMPLA _
Or oldY.TVACOMPIE <> newY.TVACOMPIE _
Or oldY.TVACOMECR <> newY.TVACOMECR Then
    sqlYTVACOM0_Update = "Erreur TVACOMPIE : " & newY.TVACOMPIE & "." & oldY.TVACOMECR
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where TVACOMETA = " & oldY.TVACOMETA _
       & " and TVACOMPLA = " & oldY.TVACOMPLA _
       & " and TVACOMPIE = " & oldY.TVACOMPIE _
       & " and TVACOMECR = " & oldY.TVACOMECR _
       & " and TVACOMUPDS = " & oldY.TVACOMUPDS

newY.TVACOMUPDS = newY.TVACOMUPDS + 1
xSet = xSet & " set TVACOMUPDS = " & newY.TVACOMUPDS
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.TVACOMDTR <> oldY.TVACOMDTR Then blnUpdate = True: xSet = xSet & " , TVACOMDTR = " & newY.TVACOMDTR
If newY.TVACOMDVA <> oldY.TVACOMDVA Then blnUpdate = True: xSet = xSet & " , TVACOMDVA = " & newY.TVACOMDVA
If newY.TVACOMDOS <> oldY.TVACOMDOS Then blnUpdate = True: xSet = xSet & " , TVACOMDOS = " & newY.TVACOMDOS
If newY.TVACOMMON <> oldY.TVACOMMON Then blnUpdate = True:  xSet = xSet & " , TVACOMMON = '" & cur_P(newY.TVACOMMON) & "'"
If newY.TVACOMMONE <> oldY.TVACOMMONE Then blnUpdate = True:  xSet = xSet & " , TVACOMMONE = '" & cur_P(newY.TVACOMMONE) & "'"
If newY.TVACOMMTVA <> oldY.TVACOMMTVA Then blnUpdate = True: xSet = xSet & " , TVACOMMTVA = " & newY.TVACOMMTVA
If newY.TVACOMMTVE <> oldY.TVACOMMTVE Then blnUpdate = True: xSet = xSet & " , TVACOMMTVE = " & newY.TVACOMMTVE
If newY.TVACOMFACN <> oldY.TVACOMFACN Then blnUpdate = True: xSet = xSet & " , TVACOMFACN = " & newY.TVACOMFACN
If newY.TVACOMFACL <> oldY.TVACOMFACL Then blnUpdate = True: xSet = xSet & " , TVACOMFACL = " & newY.TVACOMFACL
If newY.TVACOMXNUR <> oldY.TVACOMXNUR Then blnUpdate = True: xSet = xSet & " , tvacomxnur = " & newY.TVACOMXNUR
If newY.TVACOMXUTI <> oldY.TVACOMXUTI Then blnUpdate = True: xSet = xSet & " , tvacomxuti = " & newY.TVACOMXUTI
If newY.TVACOMXSEQ <> oldY.TVACOMXSEQ Then blnUpdate = True: xSet = xSet & " , tvacomxseq = " & newY.TVACOMXSEQ
If newY.TVACOMXSPE <> oldY.TVACOMXSPE Then blnUpdate = True: xSet = xSet & " , tvacomxspe = " & newY.TVACOMXSPE
If newY.TVACOMECRX <> oldY.TVACOMECRX Then blnUpdate = True: xSet = xSet & " , tvacomECRX = " & newY.TVACOMECRX
If newY.TVACOMQTE <> oldY.TVACOMQTE Then blnUpdate = True: xSet = xSet & " , TVACOMQTE = " & newY.TVACOMQTE

If newY.TVACOMSER <> oldY.TVACOMSER Then blnUpdate = True:  xSet = xSet & " , TVACOMSER = '" & newY.TVACOMSER & "'"
If newY.TVACOMSSE <> oldY.TVACOMSSE Then blnUpdate = True:  xSet = xSet & " , TVACOMSSE = '" & newY.TVACOMSSE & "'"
If newY.TVACOMCPT <> oldY.TVACOMCPT Then blnUpdate = True:  xSet = xSet & " , tvacomcpt = '" & newY.TVACOMCPT & "'"
If newY.TVACOMOPE <> oldY.TVACOMOPE Then blnUpdate = True:  xSet = xSet & " , TVACOMOPE = '" & newY.TVACOMOPE & "'"
If newY.TVACOMNAT <> oldY.TVACOMNAT Then blnUpdate = True: xSet = xSet & " , TVACOMNAT = " & newY.TVACOMNAT
If newY.TVACOMEVE <> oldY.TVACOMEVE Then blnUpdate = True:  xSet = xSet & " , TVACOMEVE = '" & Trim(newY.TVACOMEVE) & "'"
If newY.TVACOMDEV <> oldY.TVACOMDEV Then blnUpdate = True:  xSet = xSet & " , TVACOMDEV= '" & newY.TVACOMDEV & "'"
If newY.TVACOMCOMB <> oldY.TVACOMCOMB Then blnUpdate = True:  xSet = xSet & " , tvacomcomb = '" & newY.TVACOMCOMB & "'"
If newY.TVACOMCOMC <> oldY.TVACOMCOMC Then blnUpdate = True:  xSet = xSet & " , TVACOMCOMC= '" & newY.TVACOMCOMC & "'"
If newY.TVACOMCOME <> oldY.TVACOMCOME Then blnUpdate = True:  xSet = xSet & " , TVACOMCOME = '" & Replace(Trim(newY.TVACOMCOME), "'", "''") & "'"
If newY.TVACOMCOMT <> oldY.TVACOMCOMT Then blnUpdate = True:  xSet = xSet & " , TVACOMCOMT = '" & Replace(Trim(newY.TVACOMCOMT), "'", "''") & "'"
If newY.TVACOMCLIC <> oldY.TVACOMCLIC Then blnUpdate = True:  xSet = xSet & " , TVACOMCLIC = '" & Replace(Trim(newY.TVACOMCLIC), "'", "''") & "'"
If newY.TVACOMCLI <> oldY.TVACOMCLI Then blnUpdate = True:  xSet = xSet & " , TVACOMCLI = '" & Replace(Trim(newY.TVACOMCLI), "'", "''") & "'"
If newY.TVACOMCLIP <> oldY.TVACOMCLIP Then blnUpdate = True:  xSet = xSet & " , TVACOMCLIP = '" & Replace(Trim(newY.TVACOMCLIP), "'", "''") & "'"
If newY.TVACOMTVAC <> oldY.TVACOMTVAC Then blnUpdate = True:  xSet = xSet & " , TVACOMTVAC = '" & Replace(Trim(newY.TVACOMTVAC), "'", "''") & "'"
If newY.TVACOMSRVR <> oldY.TVACOMSRVR Then blnUpdate = True:  xSet = xSet & " , TVACOMSRVR = '" & newY.TVACOMSRVR & "'"
If newY.TVACOMSTA <> oldY.TVACOMSTA Then blnUpdate = True:  xSet = xSet & " , TVACOMSTA = '" & newY.TVACOMSTA & "'"
If newY.TVACOMXEVE <> oldY.TVACOMXEVE Then blnUpdate = True:  xSet = xSet & " , tvacomxeve = '" & newY.TVACOMXEVE & "'"
If newY.TVACOMGTYP <> oldY.TVACOMGTYP Then blnUpdate = True:  xSet = xSet & " , tvacomGTYP = '" & newY.TVACOMGTYP & "'"
If newY.TVACOMGORD <> oldY.TVACOMGORD Then blnUpdate = True:  xSet = xSet & " , tvacomGORD = '" & newY.TVACOMGORD & "'"
If newY.TVAREFCLI <> oldY.TVAREFCLI Then blnUpdate = True:  xSet = xSet & " , tvarefCLI = '" & newY.TVAREFCLI & "'"

newY.TVACOMUSR = usrName_UCase10
xSet = xSet & " , TVACOMUSR = '" & usrName_UCase10 & "'"
If newY.TVACOMETA < 0 Then blnUpdate = True  ' records techniques

If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE & ".YTVACOM0" & xSet & xWhere
    Call FEU_ROUGE
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYTVACOM0_Update = "Erreur màj : " & newY.TVACOMETA
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYTVACOM0_Update = Error
End Function

Public Function sqlTVACOMCLI(lTVACOMCLIC As String, lTVACOMCLI As String, lZCLIENA0 As typeZCLIENA0, lZADRESS0 As typeZADRESS0)
Dim V, xSQL As String
Dim xZCDOTIE0 As typeZCDOTIE0
Dim xZCHGPAS0 As typeZCHGPAS0
Dim xZENCTIE0 As typeZENCTIE0

On Error GoTo Error_Handler
sqlTVACOMCLI = "? inconnu"
rsZADRESS0_Init lZADRESS0
rsZCLIENA0_Init lZCLIENA0
lZADRESS0.ADRESSRA1 = "? inconnu"

Select Case lTVACOMCLIC
    Case " "
        xSQL = "select * from " & paramIBM_Library_SAB & ".ZCLIENA0 where CLIENACLI = '" & lTVACOMCLI & "'"
        Set rsSab = cnsab.Execute(xSQL)
        If Not rsSab.EOF Then
            V = rsZCLIENA0_GetBuffer(rsSab, lZCLIENA0)
           lZADRESS0.ADRESSNUM = lTVACOMCLI
           lZADRESS0.ADRESSCOA = "CO"
            If IsNull(rsZADRESS0_Client(lZADRESS0)) Then sqlTVACOMCLI = Null
' surchage responsable de compte par code routage
            xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'TVAFACSTA' and BIATABK1 = 'CLIENARES' and BIATABK2 = '" & lTVACOMCLI & "'"
            Set rsSab = cnsab.Execute(xSQL)
            If Not rsSab.EOF Then lZCLIENA0.CLIENARES = rsSab("BIATABTXT")
            
        End If
        
   Case "D"
    xZCDOTIE0.CDOTIETIE = lTVACOMCLI
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZCDOTIE0 where CDOTIETIE = '" & xZCDOTIE0.CDOTIETIE & "'"
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        V = rsZCDOTIE0_GetBuffer(rsSab, xZCDOTIE0)
        lZCLIENA0.CLIENAETB = xZCDOTIE0.CDOTIEETB
        lZCLIENA0.CLIENACLI = xZCDOTIE0.CDOTIETIE
        lZCLIENA0.CLIENARA1 = xZCDOTIE0.CDOTIERA1
        lZCLIENA0.CLIENARA2 = xZCDOTIE0.CDOTIERA2
        lZCLIENA0.CLIENARSD = xZCDOTIE0.CDOTIEPAR
        lZCLIENA0.CLIENAECO = xZCDOTIE0.CDOTIEECO
        lZCLIENA0.CLIENACAT = xZCDOTIE0.CDOTIECAT
        lZCLIENA0.CLIENAMES = xZCDOTIE0.CDOTIEMES
        lZCLIENA0.CLIENARES = ""
        
        
        lZADRESS0.ADRESSETA = xZCDOTIE0.CDOTIEETB                      ' Etablissement
        lZADRESS0.ADRESSTYP = "D"      ' String * 1                     ' 1 client , 2 compte
        lZADRESS0.ADRESSPLA = 0       ' Long                           ' Numéro de plan
        lZADRESS0.ADRESSNUM = xZCDOTIE0.CDOTIETIE      ' String * 20                    ' ou numéro de client
        lZADRESS0.ADRESSCOA = ""      ' String * 2                     ' Code adresse
        lZADRESS0.ADRESSDLI = 0       ' Long                           ' Date limite validité
        lZADRESS0.ADRESSDDE = 0       ' Long                           ' Date début validité
        lZADRESS0.ADRESSRA1 = xZCDOTIE0.CDOTIERA1      ' String * 32                    ' ou raison sociale 1
        lZADRESS0.ADRESSRA2 = xZCDOTIE0.CDOTIERA2      ' String * 32                    ' ou raison sociale 2
        lZADRESS0.ADRESSAD1 = xZCDOTIE0.CDOTIEAD1     ' String * 32                    ' Adresse 1
        lZADRESS0.ADRESSAD2 = xZCDOTIE0.CDOTIEAD2     ' String * 32                    ' Adresse 2
        lZADRESS0.ADRESSAD3 = xZCDOTIE0.CDOTIEAD3      ' String * 32                    ' Adresse 3
        lZADRESS0.ADRESSCOP = xZCDOTIE0.CDOTIECOP    ' String * 6                     ' Code postal
        lZADRESS0.ADRESSVIL = xZCDOTIE0.CDOTIEVIL      ' String * 25                    ' Ville
        lZADRESS0.ADRESSPAY = xZCDOTIE0.CDOTIEPAY      ' String * 25                    ' Pays
        lZADRESS0.ADRESSTEL = xZCDOTIE0.CDOTIETEL     ' String * 20                    ' No Tel.
        lZADRESS0.ADRESSFAX = xZCDOTIE0.CDOTIEFAX       ' String * 20                    ' No Fax.
        lZADRESS0.ADRESSTEX = xZCDOTIE0.CDOTIETEX        ' String * 20                    ' No Télex
        sqlTVACOMCLI = Null
    End If
   Case "G"
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZCHGPAS0 where CHGPASNU = " & lTVACOMCLI
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        V = rsZCHGPAS0_GetBuffer(rsSab, xZCHGPAS0)
        lZCLIENA0.CLIENAETB = xZCHGPAS0.CHGPASET
        lZCLIENA0.CLIENACLI = Format(xZCHGPAS0.CHGPASNU, "0000000")
        lZCLIENA0.CLIENARA1 = xZCHGPAS0.CHGPASN1
        lZCLIENA0.CLIENARA2 = xZCHGPAS0.CHGPASN2
        lZCLIENA0.CLIENARSD = xZCHGPAS0.CHGPASRE
        lZCLIENA0.CLIENAMES = xZCHGPAS0.CHGPASLG
        lZCLIENA0.CLIENARES = ""
        
        
        lZADRESS0.ADRESSETA = xZCHGPAS0.CHGPASET                      ' Etablissement
        lZADRESS0.ADRESSTYP = "G"      ' String * 1                     ' 1 client , 2 compte
        lZADRESS0.ADRESSPLA = 0       ' Long                           ' Numéro de plan
        lZADRESS0.ADRESSNUM = lZCLIENA0.CLIENACLI      ' String * 20                    ' ou numéro de client
        lZADRESS0.ADRESSCOA = ""      ' String * 2                     ' Code adresse
        lZADRESS0.ADRESSDLI = 0       ' Long                           ' Date limite validité
        lZADRESS0.ADRESSDDE = 0       ' Long                           ' Date début validité
        lZADRESS0.ADRESSRA1 = xZCHGPAS0.CHGPASN1      ' String * 32                    ' ou raison sociale 1
        lZADRESS0.ADRESSRA2 = xZCHGPAS0.CHGPASN2      ' String * 32                    ' ou raison sociale 2
        lZADRESS0.ADRESSAD1 = xZCHGPAS0.CHGPASA1     ' String * 32                    ' Adresse 1
        lZADRESS0.ADRESSAD2 = xZCHGPAS0.CHGPASA2     ' String * 32                    ' Adresse 2
        lZADRESS0.ADRESSCOP = xZCHGPAS0.CHGPASC1    ' String * 6                     ' Code postal
        lZADRESS0.ADRESSVIL = xZCHGPAS0.CHGPASVI      ' String * 25                    ' Ville
        lZADRESS0.ADRESSPAY = xZCHGPAS0.CHGPASPA      ' String * 25                    ' Pays
        sqlTVACOMCLI = Null
    End If
   Case "R"
    xZENCTIE0.ENCTIETIE = lTVACOMCLI
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZENCTIE0 where ENCTIETIE = '" & xZENCTIE0.ENCTIETIE & "'"
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        V = rsZENCTIE0_GetBuffer(rsSab, xZENCTIE0)
        lZCLIENA0.CLIENAETB = xZENCTIE0.ENCTIEETA
        lZCLIENA0.CLIENACLI = xZENCTIE0.ENCTIETIE
        lZCLIENA0.CLIENARA1 = xZENCTIE0.ENCTIERA1
        lZCLIENA0.CLIENARA2 = xZENCTIE0.ENCTIERA2
        lZCLIENA0.CLIENARSD = xZENCTIE0.ENCTIEPAR
        lZCLIENA0.CLIENAECO = xZENCTIE0.ENCTIEECO
        lZCLIENA0.CLIENACAT = ""
        lZCLIENA0.CLIENAMES = xZENCTIE0.ENCTIEMES
        lZCLIENA0.CLIENARES = ""
        
        
        lZADRESS0.ADRESSETA = xZENCTIE0.ENCTIEETA                     ' Etablissement
        lZADRESS0.ADRESSTYP = "R"      ' String * 1                     ' 1 client , 2 compte
        lZADRESS0.ADRESSPLA = 0       ' Long                           ' Numéro de plan
        lZADRESS0.ADRESSNUM = xZENCTIE0.ENCTIETIE      ' String * 20                    ' ou numéro de client
        lZADRESS0.ADRESSCOA = ""      ' String * 2                     ' Code adresse
        lZADRESS0.ADRESSDLI = 0       ' Long                           ' Date limite validité
        lZADRESS0.ADRESSDDE = 0       ' Long                           ' Date début validité
        lZADRESS0.ADRESSRA1 = xZENCTIE0.ENCTIERA1      ' String * 32                    ' ou raison sociale 1
        lZADRESS0.ADRESSRA2 = xZENCTIE0.ENCTIERA2      ' String * 32                    ' ou raison sociale 2
        lZADRESS0.ADRESSAD1 = xZENCTIE0.ENCTIEAD1     ' String * 32                    ' Adresse 1
        lZADRESS0.ADRESSAD2 = xZENCTIE0.ENCTIEAD2     ' String * 32                    ' Adresse 2
        lZADRESS0.ADRESSAD3 = xZENCTIE0.ENCTIEAD3      ' String * 32                    ' Adresse 3
        lZADRESS0.ADRESSCOP = xZENCTIE0.ENCTIECOP    ' String * 6                     ' Code postal
        lZADRESS0.ADRESSVIL = xZENCTIE0.ENCTIEVIL      ' String * 25                    ' Ville
        lZADRESS0.ADRESSPAY = xZENCTIE0.ENCTIEPAY      ' String * 25                    ' Pays
        lZADRESS0.ADRESSTEL = xZENCTIE0.ENCTIETEL     ' String * 20                    ' No Tel.
        lZADRESS0.ADRESSFAX = xZENCTIE0.ENCTIEFAX       ' String * 20                    ' No Fax.
        lZADRESS0.ADRESSTEX = xZENCTIE0.ENCTIETEX        ' String * 20                    ' No Télex
        sqlTVACOMCLI = Null
    End If
   Case "E"
    xZENCTIE0.ENCTIETIE = lTVACOMCLI
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZENCTIE0 where ENCTIETIE = '" & xZENCTIE0.ENCTIETIE & "'"
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        V = rsZENCTIE0_GetBuffer(rsSab, xZENCTIE0)
        lZCLIENA0.CLIENAETB = xZENCTIE0.ENCTIEETA
        lZCLIENA0.CLIENACLI = xZENCTIE0.ENCTIETIE
        lZCLIENA0.CLIENARA1 = xZENCTIE0.ENCTIERA1
        lZCLIENA0.CLIENARA2 = xZENCTIE0.ENCTIERA2
        lZCLIENA0.CLIENARSD = xZENCTIE0.ENCTIEPAR
        lZCLIENA0.CLIENAECO = xZENCTIE0.ENCTIEECO
        lZCLIENA0.CLIENACAT = ""
        lZCLIENA0.CLIENAMES = xZENCTIE0.ENCTIEMES
        lZCLIENA0.CLIENARES = ""
        
        
        lZADRESS0.ADRESSETA = xZENCTIE0.ENCTIEETA                     ' Etablissement
        lZADRESS0.ADRESSTYP = "E"      ' String * 1                     ' 1 client , 2 compte
        lZADRESS0.ADRESSPLA = 0       ' Long                           ' Numéro de plan
        lZADRESS0.ADRESSNUM = xZENCTIE0.ENCTIETIE      ' String * 20                    ' ou numéro de client
        lZADRESS0.ADRESSCOA = ""      ' String * 2                     ' Code adresse
        lZADRESS0.ADRESSDLI = 0       ' Long                           ' Date limite validité
        lZADRESS0.ADRESSDDE = 0       ' Long                           ' Date début validité
        lZADRESS0.ADRESSRA1 = xZENCTIE0.ENCTIERA1      ' String * 32                    ' ou raison sociale 1
        lZADRESS0.ADRESSRA2 = xZENCTIE0.ENCTIERA2      ' String * 32                    ' ou raison sociale 2
        lZADRESS0.ADRESSAD1 = xZENCTIE0.ENCTIEAD1     ' String * 32                    ' Adresse 1
        lZADRESS0.ADRESSAD2 = xZENCTIE0.ENCTIEAD2     ' String * 32                    ' Adresse 2
        lZADRESS0.ADRESSAD3 = xZENCTIE0.ENCTIEAD3      ' String * 32                    ' Adresse 3
        lZADRESS0.ADRESSCOP = xZENCTIE0.ENCTIECOP    ' String * 6                     ' Code postal
        lZADRESS0.ADRESSVIL = xZENCTIE0.ENCTIEVIL      ' String * 25                    ' Ville
        lZADRESS0.ADRESSPAY = xZENCTIE0.ENCTIEPAY      ' String * 25                    ' Pays
        lZADRESS0.ADRESSTEL = xZENCTIE0.ENCTIETEL     ' String * 20                    ' No Tel.
        lZADRESS0.ADRESSFAX = xZENCTIE0.ENCTIEFAX       ' String * 20                    ' No Fax.
        lZADRESS0.ADRESSTEX = xZENCTIE0.ENCTIETEX        ' String * 20                    ' No Télex
        sqlTVACOMCLI = Null
    End If
End Select
Exit Function

Error_Handler:
    V = Error
Error_MsgBox:
    'MsgBox V, vbCritical, Me.Name & " : " & currentAction
    sqlTVACOMCLI = V
End Function

Public Function rsYTVACOM0_GetBuffer(rsAdo As ADODB.Recordset, lYTVACOM0 As typeYTVACOM0)
On Error GoTo Error_Handler
rsYTVACOM0_GetBuffer = Null

lYTVACOM0.TVACOMETA = rsAdo("TVACOMETA")
lYTVACOM0.TVACOMPLA = rsAdo("TVACOMPLA")
lYTVACOM0.TVACOMSER = rsAdo("TVACOMSER")
lYTVACOM0.TVACOMSSE = rsAdo("TVACOMSSE")
lYTVACOM0.TVACOMPIE = rsAdo("TVACOMPIE")
lYTVACOM0.TVACOMECR = rsAdo("TVACOMECR")
lYTVACOM0.TVACOMCPT = rsAdo("TVACOMCPT")
lYTVACOM0.TVACOMDTR = rsAdo("TVACOMDTR")
lYTVACOM0.TVACOMDVA = rsAdo("TVACOMDVA")
lYTVACOM0.TVACOMOPE = rsAdo("TVACOMOPE")
lYTVACOM0.TVACOMNAT = rsAdo("TVACOMNAT")
lYTVACOM0.TVACOMEVE = rsAdo("TVACOMEVE")
lYTVACOM0.TVACOMDOS = rsAdo("TVACOMDOS")
lYTVACOM0.TVACOMDEV = rsAdo("TVACOMDEV")
lYTVACOM0.TVACOMMON = rsAdo("TVACOMMON")
lYTVACOM0.TVACOMMONE = rsAdo("TVACOMMONE")
lYTVACOM0.TVACOMMTVA = rsAdo("TVACOMMTVA")
lYTVACOM0.TVACOMMTVE = rsAdo("TVACOMMTVE")
lYTVACOM0.TVACOMCOMB = rsAdo("TVACOMCOMB")
lYTVACOM0.TVACOMCOMC = rsAdo("TVACOMCOMC")
lYTVACOM0.TVACOMCOME = rsAdo("TVACOMCOME")
lYTVACOM0.TVACOMCOMT = rsAdo("TVACOMCOMT")
lYTVACOM0.TVACOMCLIC = rsAdo("TVACOMCLIC")
lYTVACOM0.TVACOMCLI = rsAdo("TVACOMCLI")
lYTVACOM0.TVACOMCLIP = rsAdo("TVACOMCLIP")
lYTVACOM0.TVACOMTVAC = rsAdo("TVACOMTVAC")
lYTVACOM0.TVACOMFACN = rsAdo("TVACOMFACN")
lYTVACOM0.TVACOMFACL = rsAdo("TVACOMFACL")
lYTVACOM0.TVACOMSRVR = rsAdo("TVACOMSRVR")
lYTVACOM0.TVACOMQTE = rsAdo("TVACOMQTE")
lYTVACOM0.TVACOMSTA = rsAdo("TVACOMSTA")
lYTVACOM0.TVACOMUPDS = rsAdo("TVACOMUPDS")
lYTVACOM0.TVACOMUSR = rsAdo("TVACOMUSR")

lYTVACOM0.TVACOMXNUR = rsAdo("TVACOMXNUR")
lYTVACOM0.TVACOMXUTI = rsAdo("TVACOMXUTI")
lYTVACOM0.TVACOMXEVE = rsAdo("TVACOMXEVE")
lYTVACOM0.TVACOMXSEQ = rsAdo("TVACOMXSEQ")
lYTVACOM0.TVACOMXSPE = rsAdo("TVACOMXSPE")
lYTVACOM0.TVACOMECRX = rsAdo("TVACOMECRX")
lYTVACOM0.TVACOMGTYP = rsAdo("TVACOMGTYP")
lYTVACOM0.TVACOMGORD = rsAdo("TVACOMGORD")
lYTVACOM0.TVAREFCLI = rsAdo("TVAREFCLI")

Exit Function
Error_Handler:
rsYTVACOM0_GetBuffer = Error


End Function

Public Function rsYTVACOM0_Init(lYTVACOM0 As typeYTVACOM0)

lYTVACOM0.TVACOMETA = 0      'établissement
lYTVACOM0.TVACOMPLA = 0      'plan
lYTVACOM0.TVACOMSER = ""
lYTVACOM0.TVACOMSSE = ""
lYTVACOM0.TVACOMPIE = 0      'pièce
lYTVACOM0.TVACOMECR = 0      'écriture
lYTVACOM0.TVACOMCPT = ""
lYTVACOM0.TVACOMDTR = 0      'date de traitement
lYTVACOM0.TVACOMDVA = 0
lYTVACOM0.TVACOMOPE = ""      ' 3   'code opération
lYTVACOM0.TVACOMNAT = ""      ' 6   'nature opération
lYTVACOM0.TVACOMEVE = ""      ' 3   'événement
lYTVACOM0.TVACOMDOS = 0      'N° dossier
lYTVACOM0.TVACOMDEV = ""      ' 3   'devise
lYTVACOM0.TVACOMMON = 0      'montant en dev
lYTVACOM0.TVACOMMONE = 0     'montant euro
lYTVACOM0.TVACOMMTVA = 0     'montant TVA issu du dossier
lYTVACOM0.TVACOMMTVE = 0     'montant TVA issu du dossier
lYTVACOM0.TVACOMCOMB = ""
lYTVACOM0.TVACOMCOMC = ""     ' 6   'code commission
lYTVACOM0.TVACOMCOME = ""     ' 1   'code édition espace,W,A
lYTVACOM0.TVACOMCOMT = ""     ' 1   'taxable Exonéré,Normal,Réduit
lYTVACOM0.TVACOMCLIC = ""     ' 1   'table client espace,G, D
lYTVACOM0.TVACOMCLI = ""      ' 7   'code client
lYTVACOM0.TVACOMCLIP = ""     ' 2   'pays de résidence
lYTVACOM0.TVACOMTVAC = ""     ' 1   'code TVA
lYTVACOM0.TVACOMFACN = 0         'n° facture
lYTVACOM0.TVACOMFACL = 0         'n° facture
lYTVACOM0.TVACOMSRVR = ""      ' 1   'statut
lYTVACOM0.TVACOMQTE = 0
lYTVACOM0.TVACOMSTA = ""      ' 1   'statut
lYTVACOM0.TVACOMUPDS = 0

lYTVACOM0.TVACOMXNUR = 0
lYTVACOM0.TVACOMXUTI = 0
lYTVACOM0.TVACOMXEVE = 0
lYTVACOM0.TVACOMXSEQ = 0
lYTVACOM0.TVACOMXSPE = 0
lYTVACOM0.TVACOMECRX = 0      'écriture
lYTVACOM0.TVACOMGTYP = 0      'écriture
lYTVACOM0.TVACOMGORD = 0      'écriture


End Function



