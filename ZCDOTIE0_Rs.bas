Attribute VB_Name = "rsZCDOTIE0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeZCDOTIE0
    CDOTIEETB       As Integer                        ' CODE ETABLISSEMENT
    CDOTIETIE       As String * 7                     ' NUMERO TIERS
    CDOTIECLI       As String * 7                     ' CLIENT ASSOCIE
    CDOTIERA1       As String * 32                    ' NOM OU DESIGNATION
    CDOTIERA2       As String * 32                    ' PRENOM/DESIGNATION
    CDOTIESIG       As String * 12                    ' SIGLE USUEL
    CDOTIEPAR       As String * 3                     ' CDE PAYS DE RESIDENC
    CDOTIEECO       As String * 3                     ' QUALITE/AG ECONOMIQU
    CDOTIECAT       As String * 3                     ' CATEGORIE CLIENT
    CDOTIEMES       As String * 1                     ' LANGUE MESSAGERIE
    CDOTIEBIC       As String * 16                    ' BIC (SWIFT)
    CDOTIEBAN       As String * 5                     ' CODE BANQUE
    CDOTIEGUI       As String * 5                     ' CODE GUICHET
    CDOTIECOM       As String * 20                    ' COMPTE
    CDOTIEAD1       As String * 32                    ' ADRESSE 1
    CDOTIEAD2       As String * 32                    ' ADRESSE 2
    CDOTIEAD3       As String * 32                    ' COMMUNE
    CDOTIECOP       As String * 6                     ' CODE POSTAL
    CDOTIEVIL       As String * 25                    ' BUREAU DISTRIBUTEUR
    CDOTIEPAY       As String * 32                    ' PAYS
    CDOTIETEL       As String * 20                    ' TELEPHONE
    CDOTIEFAX       As String * 20                    ' No FAX
    CDOTIETEX       As String * 20                    ' No TELEX
    CDOTIESRN       As String * 9                     ' NUMERO SIREN
    CDOTIECOT       As String * 1                     ' CORRESPOND. CLI/TIE
    CDOTIECOR       As String * 7                     ' CORRESPONDANT

End Type
Public Sub rsZCDOTIE_Adresse(lCDODOSxxT As String, lCDODOSxxR As String, lCDODOSxxX As String, lYCDOTIE0 As typeZCDOTIE0, lZADRESS0 As typeZADRESS0, lConcat As String, lCodeAdresse As String)
Dim wId As String
Dim X As String, X1 As String
Dim I As Integer, K As Integer
Dim V, xSQL As String
Dim blnCDODOSxxX As Boolean

blnCDODOSxxX = False
rsZADRESS0_Init lZADRESS0
If lCDODOSxxT = "T" Then
    lYCDOTIE0.CDOTIETIE = lCDODOSxxR
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZCDOTIE0 where CDOTIETIE = '" & lYCDOTIE0.CDOTIETIE & "'"
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        V = rsZCDOTIE0_GetBuffer(rsSab, lYCDOTIE0)
        lZADRESS0.ADRESSETA = lYCDOTIE0.CDOTIEETB                      ' Etablissement
        lZADRESS0.ADRESSTYP = "T"      ' String * 1                     ' 1 client , 2 compte
        lZADRESS0.ADRESSPLA = 0       ' Long                           ' Numéro de plan
        lZADRESS0.ADRESSNUM = lYCDOTIE0.CDOTIETIE      ' String * 20                    ' ou numéro de client
        lZADRESS0.ADRESSCOA = ""      ' String * 2                     ' Code adresse
        lZADRESS0.ADRESSDLI = 0       ' Long                           ' Date limite validité
        lZADRESS0.ADRESSDDE = 0       ' Long                           ' Date début validité
        lZADRESS0.ADRESSRA1 = lYCDOTIE0.CDOTIERA1      ' String * 32                    ' ou raison sociale 1
        lZADRESS0.ADRESSRA2 = lYCDOTIE0.CDOTIERA2      ' String * 32                    ' ou raison sociale 2
        lZADRESS0.ADRESSAD1 = lYCDOTIE0.CDOTIEAD1     ' String * 32                    ' Adresse 1
        lZADRESS0.ADRESSAD2 = lYCDOTIE0.CDOTIEAD2     ' String * 32                    ' Adresse 2
        lZADRESS0.ADRESSAD3 = lYCDOTIE0.CDOTIEAD3      ' String * 32                    ' Adresse 3
        lZADRESS0.ADRESSCOP = lYCDOTIE0.CDOTIECOP    ' String * 6                     ' Code postal
        lZADRESS0.ADRESSVIL = lYCDOTIE0.CDOTIEVIL      ' String * 25                    ' Ville
        lZADRESS0.ADRESSPAY = lYCDOTIE0.CDOTIEPAY      ' String * 25                    ' Pays
        lZADRESS0.ADRESSTEL = lYCDOTIE0.CDOTIETEL     ' String * 20                    ' No Tel.
        lZADRESS0.ADRESSFAX = lYCDOTIE0.CDOTIEFAX       ' String * 20                    ' No Fax.
        lZADRESS0.ADRESSTEX = lYCDOTIE0.CDOTIETEX        ' String * 20                    ' No Télex
    End If
Else
    If Trim(lCDODOSxxR) <> "" Then
'Recherche adresse spécifique CREDOC dans le fichier ZADRESS0
        lZADRESS0.ADRESSTYP = "1"
        lZADRESS0.ADRESSNUM = " " & lCDODOSxxR
        lZADRESS0.ADRESSCOA = lCodeAdresse
        xSQL = "select * from " & paramIBM_Library_SAB & ".ZADRESS0 where ADRESSNUM = '" _
        & lZADRESS0.ADRESSNUM & "' AND ADRESSCOA = '" & lZADRESS0.ADRESSCOA & "'" _
        & " AND ADRESSTYP = '1'"
        Set rsSab = cnsab.Execute(xSQL)
        If Not rsSab.EOF Then
            V = rsZADRESS0_GetBuffer(rsSab, lZADRESS0)
        Else
            V = "?"
        End If
        If Not IsNull(V) And lCodeAdresse <> "  " Then
            wId = "1 " & lCDODOSxxR
            lZADRESS0.ADRESSTYP = "1 "
            lZADRESS0.ADRESSNUM = " " & lCDODOSxxR
            lZADRESS0.ADRESSCOA = ""
        xSQL = "select * from " & paramIBM_Library_SAB & ".ZADRESS0 where ADRESSNUM = '" _
        & lZADRESS0.ADRESSNUM & "' AND ADRESSCOA = '" & lZADRESS0.ADRESSCOA & "'" _
        & " AND ADRESSTYP = '1'"
        Set rsSab = cnsab.Execute(xSQL)
        If Not rsSab.EOF Then
            V = rsZADRESS0_GetBuffer(rsSab, lZADRESS0)
        Else
            V = "?"
        End If
        End If
        If Trim(lZADRESS0.ADRESSRA1) = "" Then
            'meZCLIENA0.CLIENACLI = lCDODOSxxR
            xSQL = "select * from " & paramIBM_Library_SAB & ".ZCLIENA0" _
                 & " where CLIENACLI = '" & lCDODOSxxR & "'"
            Set rsSab = cnsab.Execute(xSQL)
            If Not rsSab.EOF Then
            'V = rsZCLIENA0_GetBuffer(rsSab, lZCLIENA0)
                lZADRESS0.ADRESSRA1 = rsSab("CLIENARA1")
                lZADRESS0.ADRESSRA2 = rsSab("CLIENARA2")
            End If
        End If
    Else
        blnCDODOSxxX = True
        lZADRESS0.ADRESSRA1 = Mid$(lCDODOSxxX, 1, 32)
        lZADRESS0.ADRESSVIL = Mid$(lCDODOSxxX, 33, 32)
    End If
End If
If blnCDODOSxxX Then
    X = lCDODOSxxX
Else
    X = Trim(lZADRESS0.ADRESSRA1) & " - " & Trim(lZADRESS0.ADRESSCOP) & " " & Trim(lZADRESS0.ADRESSVIL) & " - " & Trim(lZADRESS0.ADRESSPAY)
End If
K = 1
lConcat = Mid$(X, 1, 1)
For I = 2 To Len(X)
    X1 = Mid$(X, I, 1)
    If X1 <> " " Or Mid$(lConcat, I, 1) <> " " Then lConcat = lConcat & X1
Next I
End Sub

Public Sub rsZCDOTIE0_Init(rsYCDOTIE0 As typeZCDOTIE0)
rsYCDOTIE0.CDOTIEETB = 0
rsYCDOTIE0.CDOTIETIE = ""
rsYCDOTIE0.CDOTIECLI = ""
rsYCDOTIE0.CDOTIERA1 = ""
rsYCDOTIE0.CDOTIERA2 = ""
rsYCDOTIE0.CDOTIESIG = ""
rsYCDOTIE0.CDOTIEPAR = ""
rsYCDOTIE0.CDOTIEECO = ""
rsYCDOTIE0.CDOTIECAT = ""
rsYCDOTIE0.CDOTIEMES = ""
rsYCDOTIE0.CDOTIEBIC = ""
rsYCDOTIE0.CDOTIEBAN = ""
rsYCDOTIE0.CDOTIEGUI = ""
rsYCDOTIE0.CDOTIECOM = ""
rsYCDOTIE0.CDOTIEAD1 = ""
rsYCDOTIE0.CDOTIEAD2 = ""
rsYCDOTIE0.CDOTIEAD3 = ""
rsYCDOTIE0.CDOTIECOP = ""
rsYCDOTIE0.CDOTIEVIL = ""
rsYCDOTIE0.CDOTIEPAY = ""
rsYCDOTIE0.CDOTIETEL = ""
rsYCDOTIE0.CDOTIEFAX = ""
rsYCDOTIE0.CDOTIETEX = ""
rsYCDOTIE0.CDOTIESRN = ""
rsYCDOTIE0.CDOTIECOT = ""
rsYCDOTIE0.CDOTIECOR = ""
End Sub
Public Function rsZCDOTIE0_GetBuffer(rsAdo As ADODB.Recordset, rsZCDOTIE0 As typeZCDOTIE0)
On Error GoTo Error_Handler
rsZCDOTIE0_GetBuffer = Null
rsZCDOTIE0.CDOTIEETB = rsAdo("CDOTIEETB")
rsZCDOTIE0.CDOTIETIE = rsAdo("CDOTIETIE")
rsZCDOTIE0.CDOTIECLI = rsAdo("CDOTIECLI")
rsZCDOTIE0.CDOTIERA1 = rsAdo("CDOTIERA1")
rsZCDOTIE0.CDOTIERA2 = rsAdo("CDOTIERA2")
rsZCDOTIE0.CDOTIESIG = rsAdo("CDOTIESIG")
rsZCDOTIE0.CDOTIEPAR = rsAdo("CDOTIEPAR")
rsZCDOTIE0.CDOTIEECO = rsAdo("CDOTIEECO")
rsZCDOTIE0.CDOTIECAT = rsAdo("CDOTIECAT")
rsZCDOTIE0.CDOTIEMES = rsAdo("CDOTIEMES")
rsZCDOTIE0.CDOTIEBIC = rsAdo("CDOTIEBIC")
rsZCDOTIE0.CDOTIEBAN = rsAdo("CDOTIEBAN")
rsZCDOTIE0.CDOTIEGUI = rsAdo("CDOTIEGUI")
rsZCDOTIE0.CDOTIECOM = rsAdo("CDOTIECOM")
rsZCDOTIE0.CDOTIEAD1 = rsAdo("CDOTIEAD1")
rsZCDOTIE0.CDOTIEAD2 = rsAdo("CDOTIEAD2")
rsZCDOTIE0.CDOTIEAD3 = rsAdo("CDOTIEAD3")
rsZCDOTIE0.CDOTIECOP = rsAdo("CDOTIECOP")
rsZCDOTIE0.CDOTIEVIL = rsAdo("CDOTIEVIL")
rsZCDOTIE0.CDOTIEPAY = rsAdo("CDOTIEPAY")
rsZCDOTIE0.CDOTIETEL = rsAdo("CDOTIETEL")
rsZCDOTIE0.CDOTIEFAX = rsAdo("CDOTIEFAX")
rsZCDOTIE0.CDOTIETEX = rsAdo("CDOTIETEX")
rsZCDOTIE0.CDOTIESRN = rsAdo("CDOTIESRN")
rsZCDOTIE0.CDOTIECOT = rsAdo("CDOTIECOT")
rsZCDOTIE0.CDOTIECOR = rsAdo("CDOTIECOR")
Exit Function
Error_Handler:
rsZCDOTIE0_GetBuffer = Error
End Function

