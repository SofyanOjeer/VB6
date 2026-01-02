Attribute VB_Name = "srvYCRTMVT0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYCRTMVT0
 
      CRTMVTETA   As Long         'établissement
      CRTMVTPLA   As Long         'plan
      CRTMVTPIE   As Long         'pièce
      CRTMVTECR   As Long         'écriture
      CRTMVTCPT   As String       'compte
      CRTMVTDTR   As Long         'date de traitement
      CRTMVTDEV   As String       'devise
      CRTMVTCOMK   As String       'statut
      CRTMVTCLIC  As String       'code client
      CRTMVTCLIN   As String       'client
      CRTMVTCLIP  As String       'pays de résidence
      CRTMVTRUB  As String        'rubriqueCRT
      CRTMVTMTE   As Currency     'montant euro
      CRTMVTSTA   As String       'statut
      CRTMVTORIG  As String       'origine
      CRTMVTSER   As String
      CRTMVTSSE   As String
      CRTMVTOPE   As String
      CRTMVTNAT   As String
      CRTMVTEVE   As String
      CRTMVTDOS   As Long
End Type
Public xYCRTMVT0 As typeYCRTMVT0
Public zYCRTMVT0_OD As typeYCRTMVT0

Public Function sqlCRTMVTCLI(lCRTMVTCLIC As String, lCRTMVTCLIN As String, lZCLIENA0 As typeZCLIENA0, lZADRESS0 As typeZADRESS0)
Dim V, xSQL As String, X As String, xCLIENACLI As String
Dim rsSabX As New ADODB.Recordset
'currentSAB_ETA As Long, currentSAB_AGE

On Error GoTo Error_Handler
sqlCRTMVTCLI = "? inconnu"
rsZADRESS0_Init lZADRESS0
rsZCLIENA0_Init lZCLIENA0
lZADRESS0.ADRESSRA1 = "? inconnu"
xCLIENACLI = Format(Val(lCRTMVTCLIN), "0000000")

Select Case lCRTMVTCLIC
    Case " "
            xSQL = "select * from " & paramIBM_Library_SAB & ".ZCLIENA0" _
                 & " where CLIENAETB = " & currentSAB_ETA & " and CLIENACLI = '" & xCLIENACLI & "'"
            Set rsSabX = cnsab.Execute(xSQL)
            
            If Not rsSabX.EOF Then
               V = rsZCLIENA0_GetBuffer(rsSabX, lZCLIENA0)
               lZADRESS0.ADRESSNUM = lCRTMVTCLIN
               lZADRESS0.ADRESSCOA = "CO"
                If IsNull(rsZADRESS0_Client(lZADRESS0)) Then sqlCRTMVTCLI = Null
            ' surchage responsable de compte par code routage
                xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'TVAFACSTA' and BIATABK1 = 'CLIENARES' and BIATABK2 = '" & lCRTMVTCLIN & "'"
                Set rsSabX = cnsab.Execute(xSQL)
                If Not rsSabX.EOF Then lZCLIENA0.CLIENARES = rsSabX("BIATABTXT")
                
            End If
        
   Case "D"
            xSQL = "select * from " & paramIBM_Library_SAB & ".ZCDOTIE0" _
                 & " where CDOTIEETB = " & currentSAB_ETA & " and CDOTIETIE = '" & xCLIENACLI & "'"
            Set rsSabX = cnsab.Execute(xSQL)
            
            If Not rsSabX.EOF Then
                'V = rsZCDOTIE0_GetBuffer(rssabx, xZCDOTIE0)
                lZCLIENA0.CLIENAETB = rsSabX("CDOTIEETB")
                lZCLIENA0.CLIENACLI = rsSabX("CDOTIETIE")
                lZCLIENA0.CLIENARA1 = rsSabX("CDOTIERA1")
                lZCLIENA0.CLIENARA2 = rsSabX("CDOTIERA2")
                lZCLIENA0.CLIENARSD = rsSabX("CDOTIEPAR")
                lZCLIENA0.CLIENAECO = rsSabX("CDOTIEECO")
                lZCLIENA0.CLIENACAT = rsSabX("CDOTIECAT")
                lZCLIENA0.CLIENAMES = rsSabX("CDOTIEMES")
                lZCLIENA0.CLIENARES = ""
                
                
                lZADRESS0.ADRESSETA = rsSabX("CDOTIEETB")                      ' Etablissement
                lZADRESS0.ADRESSTYP = "D"      ' String * 1                     ' 1 client , 2 compte
                lZADRESS0.ADRESSPLA = 0       ' Long                           ' Numéro de plan
                lZADRESS0.ADRESSNUM = rsSabX("CDOTIETIE")      ' String * 20                    ' ou numéro de client
                lZADRESS0.ADRESSCOA = ""      ' String * 2                     ' Code adresse
                lZADRESS0.ADRESSDLI = 0       ' Long                           ' Date limite validité
                lZADRESS0.ADRESSDDE = 0       ' Long                           ' Date début validité
                lZADRESS0.ADRESSRA1 = rsSabX("CDOTIERA1")      ' String * 32                    ' ou raison sociale 1
                lZADRESS0.ADRESSRA2 = rsSabX("CDOTIERA2")      ' String * 32                    ' ou raison sociale 2
                lZADRESS0.ADRESSAD1 = rsSabX("CDOTIEAD1")     ' String * 32                    ' Adresse 1
                lZADRESS0.ADRESSAD2 = rsSabX("CDOTIEAD2")     ' String * 32                    ' Adresse 2
                lZADRESS0.ADRESSAD3 = rsSabX("CDOTIEAD3")      ' String * 32                    ' Adresse 3
                lZADRESS0.ADRESSCOP = rsSabX("CDOTIECOP")    ' String * 6                     ' Code postal
                lZADRESS0.ADRESSVIL = rsSabX("CDOTIEVIL")      ' String * 25                    ' Ville
                lZADRESS0.ADRESSPAY = rsSabX("CDOTIEPAY")      ' String * 25                    ' Pays
                lZADRESS0.ADRESSTEL = rsSabX("CDOTIETEL")    ' String * 20                    ' No Tel.
                lZADRESS0.ADRESSFAX = rsSabX("CDOTIEFAX")       ' String * 20                    ' No Fax.
                lZADRESS0.ADRESSTEX = rsSabX("CDOTIETEX")        ' String * 20                    ' No Télex
                sqlCRTMVTCLI = Null
            End If
   Case "G"
            xSQL = "select * from " & paramIBM_Library_SAB & ".ZCHGPAS0" _
                & " where  CHGPASET = " & currentSAB_ETA & " and CHGPASAG = " & currentSAB_AGE & " and CHGPASNU = " & Val(lCRTMVTCLIN)
            Set rsSabX = cnsab.Execute(xSQL)
            
            If Not rsSabX.EOF Then
                'V = rsZCHGPAS0_GetBuffer(rsSabX, rssabx(")
                lZCLIENA0.CLIENAETB = rsSabX("CHGPASET")
                lZCLIENA0.CLIENACLI = Format(rsSabX("CHGPASNU"), "0000000")
                lZCLIENA0.CLIENARA1 = rsSabX("CHGPASN1")
                lZCLIENA0.CLIENARA2 = rsSabX("CHGPASN2")
                lZCLIENA0.CLIENARSD = rsSabX("CHGPASRE")
                lZCLIENA0.CLIENAMES = rsSabX("CHGPASLG")
                lZCLIENA0.CLIENARES = ""
                
                
                lZADRESS0.ADRESSETA = rsSabX("CHGPASET")                      ' Etablissement
                lZADRESS0.ADRESSTYP = "G"      ' String * 1                     ' 1 client , 2 compte
                lZADRESS0.ADRESSPLA = 0       ' Long                           ' Numéro de plan
                lZADRESS0.ADRESSNUM = lZCLIENA0.CLIENACLI      ' String * 20                    ' ou numéro de client
                lZADRESS0.ADRESSCOA = ""      ' String * 2                     ' Code adresse
                lZADRESS0.ADRESSDLI = 0       ' Long                           ' Date limite validité
                lZADRESS0.ADRESSDDE = 0       ' Long                           ' Date début validité
                lZADRESS0.ADRESSRA1 = rsSabX("CHGPASN1")      ' String * 32                    ' ou raison sociale 1
                lZADRESS0.ADRESSRA2 = rsSabX("CHGPASN2")      ' String * 32                    ' ou raison sociale 2
                lZADRESS0.ADRESSAD1 = rsSabX("CHGPASA1")     ' String * 32                    ' Adresse 1
                lZADRESS0.ADRESSAD2 = rsSabX("CHGPASA2")     ' String * 32                    ' Adresse 2
                lZADRESS0.ADRESSCOP = rsSabX("CHGPASC1")    ' String * 6                     ' Code postal
                lZADRESS0.ADRESSVIL = rsSabX("CHGPASVI")     ' String * 25                    ' Ville
                lZADRESS0.ADRESSPAY = rsSabX("CHGPASPA")      ' String * 25                    ' Pays
                sqlCRTMVTCLI = Null
            End If
   Case "R"
            xSQL = "select * from " & paramIBM_Library_SAB & ".ZENCTIE0" _
                & " where  ENCTIEETA = " & currentSAB_ETA & " and ENCTIETIE = '" & xCLIENACLI & "'"
            Set rsSabX = cnsab.Execute(xSQL)
            If Not rsSabX.EOF Then
                'V = rsZENCTIE0_GetBuffer(rsSabX, xZENCTIE0)
                lZCLIENA0.CLIENAETB = rsSabX("ENCTIEETA")
                lZCLIENA0.CLIENACLI = rsSabX("ENCTIETIE")
                lZCLIENA0.CLIENARA1 = rsSabX("ENCTIERA1")
                lZCLIENA0.CLIENARA2 = rsSabX("ENCTIERA2")
                lZCLIENA0.CLIENARSD = rsSabX("ENCTIEPAR")
                lZCLIENA0.CLIENAECO = rsSabX("ENCTIEECO")
                lZCLIENA0.CLIENACAT = ""
                lZCLIENA0.CLIENAMES = rsSabX("ENCTIEMES")
                lZCLIENA0.CLIENARES = ""
                
                
                lZADRESS0.ADRESSETA = rsSabX("ENCTIEETA")                     ' Etablissement
                lZADRESS0.ADRESSTYP = "R"      ' String * 1                     ' 1 client , 2 compte
                lZADRESS0.ADRESSPLA = 0       ' Long                           ' Numéro de plan
                lZADRESS0.ADRESSNUM = rsSabX("ENCTIETIE")      ' String * 20                    ' ou numéro de client
                lZADRESS0.ADRESSCOA = ""      ' String * 2                     ' Code adresse
                lZADRESS0.ADRESSDLI = 0       ' Long                           ' Date limite validité
                lZADRESS0.ADRESSDDE = 0       ' Long                           ' Date début validité
                lZADRESS0.ADRESSRA1 = rsSabX("ENCTIERA1")      ' String * 32                    ' ou raison sociale 1
                lZADRESS0.ADRESSRA2 = rsSabX("ENCTIERA2")      ' String * 32                    ' ou raison sociale 2
                lZADRESS0.ADRESSAD1 = rsSabX("ENCTIEAD1")     ' String * 32                    ' Adresse 1
                lZADRESS0.ADRESSAD2 = rsSabX("ENCTIEAD2")     ' String * 32                    ' Adresse 2
                lZADRESS0.ADRESSAD3 = rsSabX("ENCTIEAD3")      ' String * 32                    ' Adresse 3
                lZADRESS0.ADRESSCOP = rsSabX("ENCTIECOP")    ' String * 6                     ' Code postal
                lZADRESS0.ADRESSVIL = rsSabX("ENCTIEVIL")      ' String * 25                    ' Ville
                lZADRESS0.ADRESSPAY = rsSabX("ENCTIEPAY")      ' String * 25                    ' Pays
                lZADRESS0.ADRESSTEL = rsSabX("ENCTIETEL")     ' String * 20                    ' No Tel.
                lZADRESS0.ADRESSFAX = rsSabX("ENCTIEFAX")       ' String * 20                    ' No Fax.
                lZADRESS0.ADRESSTEX = rsSabX("ENCTIETEX")        ' String * 20                    ' No Télex
                sqlCRTMVTCLI = Null
            End If
End Select

Exit Function

Error_Handler:
    V = Error
Error_MsgBox:
    'MsgBox V, vbCritical, Me.Name & " : " & currentAction
    sqlCRTMVTCLI = V
End Function

Public Function sqlCRTMVTCLI_Pays(lCRTMVTCLIC As String, lCRTMVTCLIN As String)
Dim V, xSQL As String, xCLIENACLI As String
Dim rsSabX As New ADODB.Recordset

On Error GoTo Error_Handler
sqlCRTMVTCLI_Pays = ""
xCLIENACLI = Format(Val(lCRTMVTCLIN), "0000000")
Select Case lCRTMVTCLIC
    Case " ", ""
            xSQL = "select CLIENARSD from " & paramIBM_Library_SAB & ".ZCLIENA0 " _
                 & " where CLIENAETB = " & currentSAB_ETA & " and CLIENACLI = '" & xCLIENACLI & "'"
            Set rsSabX = cnsab.Execute(xSQL)
            If Not rsSabX.EOF Then sqlCRTMVTCLI_Pays = rsSabX("CLIENARSD")
        
   Case "D"
            xSQL = "select CDOTIEPAR from " & paramIBM_Library_SAB & ".ZCDOTIE0 " _
                  & " where CDOTIEETB = " & currentSAB_ETA & " and CDOTIETIE = '" & xCLIENACLI & "'"
            Set rsSabX = cnsab.Execute(xSQL)
            If Not rsSabX.EOF Then sqlCRTMVTCLI_Pays = rsSabX("CDOTIEPAR")
            
   Case "G"
        xSQL = "select CHGPASRE from " & paramIBM_Library_SAB & ".ZCHGPAS0 " _
             & " where  CHGPASET = " & currentSAB_ETA & " and CHGPASAG = " & currentSAB_AGE & " and andCHGPASNU = " & Val(lCRTMVTCLIN)
        Set rsSabX = cnsab.Execute(xSQL)
        If Not rsSabX.EOF Then sqlCRTMVTCLI_Pays = rsSabX("CDOTIEPAR")
        
   Case "R"
        xSQL = "select ENCTIEPAR from " & paramIBM_Library_SAB & ".ZENCTIE0 " _
             & " where  ENCTIEETA = " & currentSAB_ETA & " and ENCTIETIE = '" & xCLIENACLI & "'"
        Set rsSabX = cnsab.Execute(xSQL)
        If Not rsSabX.EOF Then sqlCRTMVTCLI_Pays = rsSabX("ENCTIEPAR")
        
End Select

Exit Function

Error_Handler:
    V = Error
Error_MsgBox:
    'MsgBox V, vbCritical, Me.Name & " : " & currentAction
    'sqlCRTMVTCLI_Pays = V
End Function

Public Function sqlYCRTMVT0_Update(newY As typeYCRTMVT0, oldY As typeYCRTMVT0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYCRTMVT0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.CRTMVTETA <> newY.CRTMVTETA _
Or oldY.CRTMVTPLA <> newY.CRTMVTPLA _
Or oldY.CRTMVTPIE <> newY.CRTMVTPIE _
Or oldY.CRTMVTECR <> newY.CRTMVTECR Then
    sqlYCRTMVT0_Update = "Erreur CRTMVTPIE : " & newY.CRTMVTPIE & "." & oldY.CRTMVTECR
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where CRTMVTETA = " & oldY.CRTMVTETA _
       & " and CRTMVTPLA = " & oldY.CRTMVTPLA _
       & " and CRTMVTPIE = " & oldY.CRTMVTPIE _
       & " and CRTMVTECR = " & oldY.CRTMVTECR

xSet = xSet & " set CRTMVTSTA = '" & newY.CRTMVTSTA & "'"
blnUpdate = False
If newY.CRTMVTSTA <> oldY.CRTMVTSTA Then blnUpdate = True
' Détecter les modifications
'===================================================================================
If newY.CRTMVTMTE <> oldY.CRTMVTMTE Then blnUpdate = True:  xSet = xSet & " , CRTMVTMTE = '" & cur_P(newY.CRTMVTMTE) & "'"
If newY.CRTMVTDTR <> oldY.CRTMVTDTR Then blnUpdate = True: xSet = xSet & " , CRTMVTDTR = " & newY.CRTMVTDTR
If newY.CRTMVTDOS <> oldY.CRTMVTDOS Then blnUpdate = True: xSet = xSet & " , CRTMVTDOS = " & newY.CRTMVTDOS

If newY.CRTMVTCPT <> oldY.CRTMVTCPT Then blnUpdate = True:  xSet = xSet & " , CRTMVTCPT= '" & newY.CRTMVTCPT & "'"
If newY.CRTMVTDEV <> oldY.CRTMVTDEV Then blnUpdate = True:  xSet = xSet & " , CRTMVTDEV= '" & newY.CRTMVTDEV & "'"
If newY.CRTMVTCOMK <> oldY.CRTMVTCOMK Then blnUpdate = True:  xSet = xSet & " , CRTMVTCOMK = '" & newY.CRTMVTCOMK & "'"
If newY.CRTMVTCLIN <> oldY.CRTMVTCLIN Then blnUpdate = True:  xSet = xSet & " , CRTMVTCLIN = '" & newY.CRTMVTCLIN & "'"
If newY.CRTMVTCLIC <> oldY.CRTMVTCLIC Then blnUpdate = True:  xSet = xSet & " , CRTMVTCLIC = '" & newY.CRTMVTCLIC & "'"
If newY.CRTMVTCLIP <> oldY.CRTMVTCLIP Then blnUpdate = True:  xSet = xSet & " , CRTMVTCLIP = '" & newY.CRTMVTCLIP & "'"
If newY.CRTMVTRUB <> oldY.CRTMVTRUB Then blnUpdate = True:  xSet = xSet & " , CRTMVTRUB = '" & newY.CRTMVTRUB & "'"
If newY.CRTMVTORIG <> oldY.CRTMVTORIG Then blnUpdate = True:  xSet = xSet & " , CRTMVTORIG = '" & newY.CRTMVTORIG & "'"
If newY.CRTMVTSSE <> oldY.CRTMVTSSE Then blnUpdate = True:  xSet = xSet & " , CRTMVTSSE = '" & newY.CRTMVTSSE & "'"
If newY.CRTMVTSER <> oldY.CRTMVTSER Then blnUpdate = True:  xSet = xSet & " , CRTMVTSER = '" & newY.CRTMVTSER & "'"
If newY.CRTMVTOPE <> oldY.CRTMVTOPE Then blnUpdate = True:  xSet = xSet & " , CRTMVTOPE = '" & newY.CRTMVTOPE & "'"
If newY.CRTMVTNAT <> oldY.CRTMVTNAT Then blnUpdate = True:  xSet = xSet & " , CRTMVTNAT = '" & newY.CRTMVTNAT & "'"
If newY.CRTMVTEVE <> oldY.CRTMVTEVE Then blnUpdate = True:  xSet = xSet & " , CRTMVTEVE = '" & newY.CRTMVTEVE & "'"


If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE & ".YCRTMVT0" & xSet & xWhere
    Call FEU_ROUGE
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYCRTMVT0_Update = "Erreur màj : " & newY.CRTMVTETA
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYCRTMVT0_Update = Error
End Function

Public Function sqlYCRTMVT0_Insert(newY As typeYCRTMVT0)
Dim V
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYCRTMVT0_Insert = Null
xSet = " (CRTMVTCPT"
xValues = " values('" & newY.CRTMVTCPT & "'"

' Détecter les modifications
'===================================================================================
If newY.CRTMVTETA <> 0 Then xSet = xSet & ",CRTMVTETA": xValues = xValues & " ," & newY.CRTMVTETA
If newY.CRTMVTPLA <> 0 Then xSet = xSet & ",CRTMVTPLA": xValues = xValues & " ," & newY.CRTMVTPLA
If newY.CRTMVTPIE <> 0 Then xSet = xSet & ",CRTMVTPIE": xValues = xValues & " ," & newY.CRTMVTPIE
If newY.CRTMVTECR <> 0 Then xSet = xSet & ",CRTMVTECR": xValues = xValues & " ," & newY.CRTMVTECR
If newY.CRTMVTDTR <> 0 Then xSet = xSet & ",CRTMVTDTR": xValues = xValues & " ," & newY.CRTMVTDTR
If newY.CRTMVTMTE <> 0 Then xSet = xSet & ",CRTMVTMTE": xValues = xValues & ", " & cur_P(newY.CRTMVTMTE)
If newY.CRTMVTDOS <> 0 Then xSet = xSet & ",CRTMVTDOS": xValues = xValues & " ," & newY.CRTMVTDOS

If Trim(newY.CRTMVTDEV) <> "" Then xSet = xSet & ",CRTMVTDEV": xValues = xValues & " ,'" & newY.CRTMVTDEV & "'"
If Trim(newY.CRTMVTCOMK) <> "" Then xSet = xSet & ",CRTMVTCOMK": xValues = xValues & " ,'" & newY.CRTMVTCOMK & "'"
If Trim(newY.CRTMVTCLIC) <> "" Then xSet = xSet & ",CRTMVTCLIC": xValues = xValues & " ,'" & newY.CRTMVTCLIC & "'"
If Trim(newY.CRTMVTCLIN) <> "" Then xSet = xSet & ",CRTMVTCLIN": xValues = xValues & " ,'" & newY.CRTMVTCLIN & "'"
If Trim(newY.CRTMVTCLIP) <> "" Then xSet = xSet & ",CRTMVTCLIP": xValues = xValues & " ,'" & newY.CRTMVTCLIP & "'"
If Trim(newY.CRTMVTRUB) <> "" Then xSet = xSet & ",CRTMVTRUB": xValues = xValues & " ,'" & newY.CRTMVTRUB & "'"
If Trim(newY.CRTMVTSTA) <> "" Then xSet = xSet & ",CRTMVTSTA": xValues = xValues & " ,'" & newY.CRTMVTSTA & "'"
If Trim(newY.CRTMVTORIG) <> "" Then xSet = xSet & ",CRTMVTORIG": xValues = xValues & " ,'" & newY.CRTMVTORIG & "'"
If Trim(newY.CRTMVTSER) <> "" Then xSet = xSet & ",CRTMVTSER": xValues = xValues & " ,'" & newY.CRTMVTSER & "'"
If Trim(newY.CRTMVTSSE) <> "" Then xSet = xSet & ",CRTMVTSSE": xValues = xValues & " ,'" & newY.CRTMVTSSE & "'"
If Trim(newY.CRTMVTOPE) <> "" Then xSet = xSet & ",CRTMVTOPE": xValues = xValues & " ,'" & newY.CRTMVTOPE & "'"
If Trim(newY.CRTMVTNAT) <> "" Then xSet = xSet & ",CRTMVTNAT": xValues = xValues & " ,'" & newY.CRTMVTNAT & "'"
If Trim(newY.CRTMVTEVE) <> "" Then xSet = xSet & ",CRTMVTEVE": xValues = xValues & " ,'" & newY.CRTMVTEVE & "'"


xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YCRTMVT0" & xSet & ")" & xValues & ")"
Call FEU_ROUGE
Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYCRTMVT0_Insert = "Erreur màj : " & newY.CRTMVTPIE & newY.CRTMVTECR
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYCRTMVT0_Insert = Error
End Function

Public Function sqlYCRTMVT0_Insert_OD(newY As typeYCRTMVT0)
Dim V
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler

CRTMVTID_Svt:
'=================
sqlYCRTMVT0_Insert_OD = Null
If newY.CRTMVTECR = 0 Then
    xSQL = "select count(*) from " & paramIBM_Library_SABSPE & ".YCRTMVT0 where CRTMVTPIE = 0"
    Set rsAdo = cnSab_Update.Execute(xSQL)
    If Not rsAdo.EOF Then
        zYCRTMVT0_OD.CRTMVTECR = rsAdo(0)
        xSQL = "select CRTMVTECR from " & paramIBM_Library_SABSPE & ".YCRTMVT0 where CRTMVTPIE = 0 and CRTMVTECR >= " & zYCRTMVT0_OD.CRTMVTECR & " order by CRTMVTECR desc"
        Set rsAdo = cnSab_Update.Execute(xSQL)
        If Not rsAdo.EOF Then zYCRTMVT0_OD.CRTMVTECR = rsAdo(0)
    End If
End If

zYCRTMVT0_OD.CRTMVTECR = zYCRTMVT0_OD.CRTMVTECR + 1
newY.CRTMVTECR = zYCRTMVT0_OD.CRTMVTECR
V = sqlYCRTMVT0_Insert(newY)
If Not IsNull(V) Then
    X = Error
    If InStr(X, "SQL0803") > 0 Then GoTo CRTMVTID_Svt
End If
Exit Function

Error_Handler:
    '[IBM][Pilote ODBC iSeries Access][DB2 UDB]SQL0803 - La valeur indiquée est incorrecte car elle produirait une clé en double.
    
    sqlYCRTMVT0_Insert_OD = Error
End Function



Public Function rsYCRTMVT0_GetBuffer(rsAdo As ADODB.Recordset, lYCRTMVT0 As typeYCRTMVT0)
On Error GoTo Error_Handler
rsYCRTMVT0_GetBuffer = Null

lYCRTMVT0.CRTMVTETA = rsAdo("CRTMVTETA")
lYCRTMVT0.CRTMVTPLA = rsAdo("CRTMVTPLA")
lYCRTMVT0.CRTMVTPIE = rsAdo("CRTMVTPIE")
lYCRTMVT0.CRTMVTECR = rsAdo("CRTMVTECR")
lYCRTMVT0.CRTMVTCPT = rsAdo("CRTMVTCPT")
lYCRTMVT0.CRTMVTDTR = rsAdo("CRTMVTDTR")
lYCRTMVT0.CRTMVTDEV = rsAdo("CRTMVTDEV")
lYCRTMVT0.CRTMVTCOMK = rsAdo("CRTMVTCOMK")
lYCRTMVT0.CRTMVTCLIC = rsAdo("CRTMVTCLIC")
lYCRTMVT0.CRTMVTCLIN = rsAdo("CRTMVTCLIN")
lYCRTMVT0.CRTMVTCLIP = rsAdo("CRTMVTCLIP")
lYCRTMVT0.CRTMVTRUB = rsAdo("CRTMVTRUB")
lYCRTMVT0.CRTMVTMTE = rsAdo("CRTMVTMTE")
lYCRTMVT0.CRTMVTSTA = rsAdo("CRTMVTSTA")
lYCRTMVT0.CRTMVTORIG = rsAdo("CRTMVTORIG")

lYCRTMVT0.CRTMVTSER = rsAdo("CRTMVTSER")
lYCRTMVT0.CRTMVTSSE = rsAdo("CRTMVTSSE")
lYCRTMVT0.CRTMVTOPE = rsAdo("CRTMVTOPE")
lYCRTMVT0.CRTMVTNAT = rsAdo("CRTMVTNAT")
lYCRTMVT0.CRTMVTEVE = rsAdo("CRTMVTEVE")
lYCRTMVT0.CRTMVTDOS = rsAdo("CRTMVTDOS")


Exit Function
Error_Handler:
rsYCRTMVT0_GetBuffer = Error


End Function

Public Function rsYCRTMVT0_Init(lYCRTMVT0 As typeYCRTMVT0)

lYCRTMVT0.CRTMVTETA = 0      'établissement
lYCRTMVT0.CRTMVTPLA = 0      'plan
lYCRTMVT0.CRTMVTPIE = 0      'pièce
lYCRTMVT0.CRTMVTECR = 0      'écriture
lYCRTMVT0.CRTMVTCPT = ""      ' 3   'devise
lYCRTMVT0.CRTMVTDTR = 0      'date de traitement
lYCRTMVT0.CRTMVTDEV = ""      ' 3   'devise
lYCRTMVT0.CRTMVTCOMK = ""      ' 1   'statut
lYCRTMVT0.CRTMVTCLIC = ""     '
lYCRTMVT0.CRTMVTCLIN = ""     '
lYCRTMVT0.CRTMVTCLIP = ""     ' 2   'pays de résidence
lYCRTMVT0.CRTMVTRUB = ""
lYCRTMVT0.CRTMVTMTE = 0     'montant euro
lYCRTMVT0.CRTMVTSTA = ""      ' 1   'statut
lYCRTMVT0.CRTMVTORIG = ""      ' 1   'statut

lYCRTMVT0.CRTMVTSER = ""
lYCRTMVT0.CRTMVTSSE = ""
lYCRTMVT0.CRTMVTOPE = ""
lYCRTMVT0.CRTMVTNAT = ""
lYCRTMVT0.CRTMVTEVE = ""
lYCRTMVT0.CRTMVTDOS = 0


End Function



