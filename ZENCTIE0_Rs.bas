Attribute VB_Name = "rsZENCTIE0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeZENCTIE0
    ENCTIEETA       As Integer                        ' CODE ETABLISSEMENT
    ENCTIETIE       As String * 7                     ' NUMERO TIERS
    ENCTIERA1       As String * 32                    ' NOM OU DESIGNATION
    ENCTIERA2       As String * 32                    ' PRENOM/DESIGNATION
    ENCTIESIG       As String * 12                    ' SIGLE USUEL
    ENCTIEPAR       As String * 3                     ' CDE PAYS DE RESIDENC
    ENCTIEECO       As String * 3                     ' QUALITE/AG ECONOMIQU
    ENCTIEMES       As String * 1                     ' LANGUE MESSAGERIE
    ENCTIEBIC       As String * 16                    ' BIC (SWIFT)
    ENCTIEBAN       As String * 5                     ' CODE BANQUE
    ENCTIEGUI       As String * 5                     ' CODE GUICHET
    ENCTIECOM       As String * 20                    ' COMPTE
    ENCTIEAD1       As String * 32                    ' ADRESSE 1
    ENCTIEAD2       As String * 32                    ' ADRESSE 2
    ENCTIEAD3       As String * 32                    ' COMMUNE
    ENCTIECOP       As String * 6                     ' CODE POSTAL
    ENCTIEVIL       As String * 25                    ' BUREAU DISTRIBUTEUR
    ENCTIEPAY       As String * 32                    ' PAYS
    ENCTIETEL       As String * 20                    ' TELEPHONE
    ENCTIEFAX       As String * 20                    ' No FAX
    ENCTIETEX       As String * 20                    ' No TELEX
    ENCTIEINT       As String * 8                     ' NUMERO SIREN
    ENCTIEOCA       As String * 1                     ' CORRESPOND. CLI/TIE

End Type
Public Sub rsZENCTIE_Adresse(lCDODOSxxT As String, lCDODOSxxR As String, lCDODOSxxX As String, lYENCTIE0 As typeZENCTIE0, lZADRESS0 As typeZADRESS0, lConcat As String, lCodeAdresse As String)
Dim wId As String
Dim X As String, X1 As String
Dim I As Integer, K As Integer
Dim V, xSQL As String
Dim blnCDODOSxxX As Boolean

blnCDODOSxxX = False
rsZADRESS0_Init lZADRESS0
If lCDODOSxxT = "T" Then
    lYENCTIE0.ENCTIETIE = lCDODOSxxR
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZENCTIE0 where ENCTIETIE = '" & lYENCTIE0.ENCTIETIE & "'"
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        V = rsZENCTIE0_GetBuffer(rsSab, lYENCTIE0)
        lZADRESS0.ADRESSETA = lYENCTIE0.ENCTIEETA                      ' Etablissement
        lZADRESS0.ADRESSTYP = "T"      ' String * 1                     ' 1 client , 2 compte
        lZADRESS0.ADRESSPLA = 0       ' Long                           ' Numéro de plan
        lZADRESS0.ADRESSNUM = lYENCTIE0.ENCTIETIE      ' String * 20                    ' ou numéro de client
        lZADRESS0.ADRESSCOA = ""      ' String * 2                     ' Code adresse
        lZADRESS0.ADRESSDLI = 0       ' Long                           ' Date limite validité
        lZADRESS0.ADRESSDDE = 0       ' Long                           ' Date début validité
        lZADRESS0.ADRESSRA1 = lYENCTIE0.ENCTIERA1      ' String * 32                    ' ou raison sociale 1
        lZADRESS0.ADRESSRA2 = lYENCTIE0.ENCTIERA2      ' String * 32                    ' ou raison sociale 2
        lZADRESS0.ADRESSAD1 = lYENCTIE0.ENCTIEAD1     ' String * 32                    ' Adresse 1
        lZADRESS0.ADRESSAD2 = lYENCTIE0.ENCTIEAD2     ' String * 32                    ' Adresse 2
        lZADRESS0.ADRESSAD3 = lYENCTIE0.ENCTIEAD3      ' String * 32                    ' Adresse 3
        lZADRESS0.ADRESSCOP = lYENCTIE0.ENCTIECOP    ' String * 6                     ' Code postal
        lZADRESS0.ADRESSVIL = lYENCTIE0.ENCTIEVIL      ' String * 25                    ' Ville
        lZADRESS0.ADRESSPAY = lYENCTIE0.ENCTIEPAY      ' String * 25                    ' Pays
        lZADRESS0.ADRESSTEL = lYENCTIE0.ENCTIETEL     ' String * 20                    ' No Tel.
        lZADRESS0.ADRESSFAX = lYENCTIE0.ENCTIEFAX       ' String * 20                    ' No Fax.
        lZADRESS0.ADRESSTEX = lYENCTIE0.ENCTIETEX        ' String * 20                    ' No Télex
    End If
Else
    If Trim(lCDODOSxxR) <> "" Then
'Recherche adresse spécifique CREDOC dans le fichier ZADRESS0
        lZADRESS0.ADRESSTYP = "1"
        lZADRESS0.ADRESSNUM = " " & lCDODOSxxR
        lZADRESS0.ADRESSCOA = lCodeAdresse
        xSQL = "select * from " & paramIBM_Library_SAB & ".ZADRESS0 where ADRESSNUM = '" _
        & lZADRESS0.ADRESSNUM & "' AND ADRESSCOA = '" & lZADRESS0.ADRESSCOA & "'"
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
        & lZADRESS0.ADRESSNUM & "' AND ADRESSCOA = '" & lZADRESS0.ADRESSCOA & "'"
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

Public Sub rsZENCTIE0_Init(rsYENCTIE0 As typeZENCTIE0)
rsYENCTIE0.ENCTIEETA = 0
rsYENCTIE0.ENCTIETIE = ""
rsYENCTIE0.ENCTIERA1 = ""
rsYENCTIE0.ENCTIERA2 = ""
rsYENCTIE0.ENCTIESIG = ""
rsYENCTIE0.ENCTIEPAR = ""
rsYENCTIE0.ENCTIEECO = ""
rsYENCTIE0.ENCTIEMES = ""
rsYENCTIE0.ENCTIEBIC = ""
rsYENCTIE0.ENCTIEBAN = ""
rsYENCTIE0.ENCTIEGUI = ""
rsYENCTIE0.ENCTIECOM = ""
rsYENCTIE0.ENCTIEAD1 = ""
rsYENCTIE0.ENCTIEAD2 = ""
rsYENCTIE0.ENCTIEAD3 = ""
rsYENCTIE0.ENCTIECOP = ""
rsYENCTIE0.ENCTIEVIL = ""
rsYENCTIE0.ENCTIEPAY = ""
rsYENCTIE0.ENCTIETEL = ""
rsYENCTIE0.ENCTIEFAX = ""
rsYENCTIE0.ENCTIETEX = ""
rsYENCTIE0.ENCTIEINT = ""
rsYENCTIE0.ENCTIEOCA = ""
End Sub
Public Function rsZENCTIE0_GetBuffer(rsAdo As ADODB.Recordset, rsZENCTIE0 As typeZENCTIE0)
On Error GoTo Error_Handler
rsZENCTIE0_GetBuffer = Null
rsZENCTIE0.ENCTIEETA = rsAdo("ENCTIEETA")
rsZENCTIE0.ENCTIETIE = rsAdo("ENCTIETIE")
rsZENCTIE0.ENCTIERA1 = rsAdo("ENCTIERA1")
rsZENCTIE0.ENCTIERA2 = rsAdo("ENCTIERA2")
rsZENCTIE0.ENCTIESIG = rsAdo("ENCTIESIG")
rsZENCTIE0.ENCTIEPAR = rsAdo("ENCTIEPAR")
rsZENCTIE0.ENCTIEECO = rsAdo("ENCTIEECO")
rsZENCTIE0.ENCTIEMES = rsAdo("ENCTIEMES")
rsZENCTIE0.ENCTIEBIC = rsAdo("ENCTIEBIC")
rsZENCTIE0.ENCTIEBAN = rsAdo("ENCTIEBAN")
rsZENCTIE0.ENCTIEGUI = rsAdo("ENCTIEGUI")
rsZENCTIE0.ENCTIECOM = rsAdo("ENCTIECOM")
rsZENCTIE0.ENCTIEAD1 = rsAdo("ENCTIEAD1")
rsZENCTIE0.ENCTIEAD2 = rsAdo("ENCTIEAD2")
rsZENCTIE0.ENCTIEAD3 = rsAdo("ENCTIEAD3")
rsZENCTIE0.ENCTIECOP = rsAdo("ENCTIECOP")
rsZENCTIE0.ENCTIEVIL = rsAdo("ENCTIEVIL")
rsZENCTIE0.ENCTIEPAY = rsAdo("ENCTIEPAY")
rsZENCTIE0.ENCTIETEL = rsAdo("ENCTIETEL")
rsZENCTIE0.ENCTIEFAX = rsAdo("ENCTIEFAX")
rsZENCTIE0.ENCTIETEX = rsAdo("ENCTIETEX")
rsZENCTIE0.ENCTIEINT = rsAdo("ENCTIEINT")
rsZENCTIE0.ENCTIEOCA = rsAdo("ENCTIEOCA")
Exit Function
Error_Handler:
rsZENCTIE0_GetBuffer = Error
End Function



