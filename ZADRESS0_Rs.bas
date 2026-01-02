Attribute VB_Name = "rsZADRESS0"
'---------------------------------------------------------
Option Explicit
Type typeZADRESS0

    ADRESSETA       As Integer                        ' Etablissement
    ADRESSTYP       As String * 1                     ' 1 client , 2 compte
    ADRESSPLA       As Long                           ' Numéro de plan
    ADRESSNUM       As String * 20                    ' ou numéro de client
    ADRESSCOA       As String * 2                     ' Code adresse
    ADRESSDLI       As Long                           ' Date limite validité
    ADRESSDDE       As Long                           ' Date début validité
    ADRESSRA1       As String * 32                    ' ou raison sociale 1
    ADRESSRA2       As String * 32                    ' ou raison sociale 2
    ADRESSAD1       As String * 32                    ' Adresse 1
    ADRESSAD2       As String * 32                    ' Adresse 2
    ADRESSAD3       As String * 32                    ' Adresse 3
    ADRESSCOP       As String * 6                     ' Code postal
    ADRESSVIL       As String * 25                    ' Ville
    ADRESSPAY       As String * 25                    ' Pays
    ADRESSTEL       As String * 20                    ' No Tel.
    ADRESSFAX       As String * 20                    ' No Fax.
    ADRESSTEX       As String * 20                    ' No Télex


End Type

'---------------------------------------------------------
Public Function rsZADRESS0_Client(lZADRESS0 As typeZADRESS0)
'----------------------------------------------------------------------------------------
'Recherche de l'enregistrement Adresse à partir de la racine client et du code courrier
'!!! ADRESSNUM = 1 espace + racine (7 caractères)
'----------------------------------------------------------------------------------------

Dim wADRESSNUM As String, wADRESSCOA As String
Dim blnOk As Boolean
Dim xSQL As String
Dim V
Dim X77 As String
On Error GoTo Error_Handler
rsZADRESS0_Client = Null
blnOk = False

wADRESSNUM = " " & Trim(lZADRESS0.ADRESSNUM)
wADRESSCOA = lZADRESS0.ADRESSCOA
rsZADRESS0_Init lZADRESS0

'Lecture Adresse avec Code Adresse
'=================================
Set rsSab = Nothing
xSQL = "select * from " & paramIBM_Library_SAB & ".ZADRESS0" _
     & " where ADRESSNUM = '" & wADRESSNUM & "'" _
     & " and  ADRESSCOA = '" & wADRESSCOA & "'" _
     & " and ADRESSTYP = '1'" _
     & " and ADRESSETA = " & currentZMNURUT0.MNURUTETB

Set rsSab = cnsab.Execute(xSQL)

If Not rsSab.EOF Then
    V = rsZADRESS0_GetBuffer(rsSab, lZADRESS0)
    If Not IsNull(V) Then
        rsZADRESS0_Client = "rsZADRESS0_Client_1 : Lecture ZADRESS0 : " & V
        Exit Function
    Else
        blnOk = True
    End If
End If

If Not blnOk Then
    'SINON Lecture Adresse avec type= '  '
    '=================================
    If wADRESSCOA <> "  " Then
        xSQL = "select * from " & paramIBM_Library_SAB & ".ZADRESS0" _
             & " where ADRESSNUM = '" & wADRESSNUM & "'" _
             & " and  ADRESSCOA = '  '" _
             & " and ADRESSTYP = '1'" _
             & " and ADRESSETA = " & currentZMNURUT0.MNURUTETB

        Set rsSab = cnsab.Execute(xSQL)
        If Not rsSab.EOF Then
            V = rsZADRESS0_GetBuffer(rsSab, lZADRESS0)
            If Not IsNull(V) Then
                rsZADRESS0_Client = "rsZADRESS0_Client_2 : Lecture ZADRESS0 : " & V
                Exit Function
            Else
                blnOk = True
            End If
        End If
    End If
End If

If Not blnOk Then
    rsZADRESS0_Client = "Adresse non trouvée  : " & wADRESSNUM
Else
    If Trim(lZADRESS0.ADRESSRA1) = "" Then
        'Lecture CLIENT principal ==> ADRESSRA1
        '=================================
        xSQL = "select CLIENARA1,CLIENARA2,CLIENAETA,CLIENACAT from " & paramIBM_Library_SAB & ".ZCLIENA0" _
             & " where CLIENACLI = '" & Trim(wADRESSNUM) & "'" _
             & " and CLIENAETB = " & currentZMNURUT0.MNURUTETB

        Set rsSab = cnsab.Execute(xSQL)
        If Not rsSab.EOF Then
            lZADRESS0.ADRESSRA1 = rsSab("CLIENARA1")
            lZADRESS0.ADRESSRA2 = rsSab("CLIENARA2")
            If Mid$(rsSab("CLIENACAT"), 1, 1) = "P" Then
               xSQL = "select BIATABTXT from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
                    & " where BIATABID = 'SAB' and BIATABK1 = 'CLIENAETA'" _
                    & " and BIATABK2 = 'CLI" & Trim(rsSab("CLIENAETA")) & "'"

                Set rsSab = cnsab.Execute(xSQL)
                If Not rsSab.EOF Then
                    X77 = Trim(Mid$(rsSab("BIATABTXT"), 1, 12)) & "  " & Trim(lZADRESS0.ADRESSRA1) & " " & Trim(lZADRESS0.ADRESSRA2)
                    If Len(X77) < 32 Then
                        lZADRESS0.ADRESSRA1 = Mid$(X77, 1, 32)
                        lZADRESS0.ADRESSRA2 = Mid$(X77, 33, 32)
                    Else
                        lZADRESS0.ADRESSRA1 = Trim(Mid$(rsSab("BIATABTXT"), 1, 12)) & "  " & Trim(lZADRESS0.ADRESSRA1)
                        lZADRESS0.ADRESSRA2 = Trim(lZADRESS0.ADRESSRA2)
                    End If
                    
                End If
            End If
        End If
    End If
End If

Exit Function

Error_Handler:
rsZADRESS0_Client = Error

End Function


Public Function rsZADRESS0_BIC_Client(lCLIENACLI As String, lBIC As String)
'------------------------------------------------
'Recherche code BIC à partir de la racine client
'!!! ADRESSNUM = 1 espace + racine (7 caractères)
'------------------------------------------------
Dim xSQL As String, wADRESSNUM As String
On Error GoTo Error_Handler

rsZADRESS0_BIC_Client = Null
wADRESSNUM = " " & lCLIENACLI

If paramIBM_AS400_ID = "I5A7" Then
    xSQL = "select ADRESSRA1 from " & paramIBM_Library_SAB & ".ZADRESS0" _
         & " where ADRESSTYP = '4'" _
         & " and ADRESSNUM ='" & wADRESSNUM & " '" _
         & " and ADRESSETA = " & currentZMNURUT0.MNURUTETB
         
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        lBIC = Mid$(rsSab("ADRESSRA1"), 11, 11)
    Else
        lBIC = ""
        rsZADRESS0_BIC_Client = "? BIC " & lCLIENACLI
    
    End If

Else
    xSQL = "select ADRESSRA12 from " & paramIBM_Library_SAB & ".ZADRESS0" _
         & " where ADRESSTYP = '4'" _
         & " and ADRESSNUM ='" & wADRESSNUM & " '" _
         & " and ADRESSETA = " & currentZMNURUT0.MNURUTETB
         
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        lBIC = Mid$(rsSab("ADRESSRA12"), 1, 11)
    Else
        lBIC = ""
        rsZADRESS0_BIC_Client = "? BIC " & lCLIENACLI
    
    End If
End If
Exit Function

Error_Handler:
rsZADRESS0_BIC_Client = Error

End Function
'---------------------------------------------------------
Public Function rsZADRESS0_Compte(lZADRESS0 As typeZADRESS0)
'---------------------------------------------------------
'Initialiser .ADRESSNUM= 'numéro de compte
'            .ADRESSCOA = '  ','CO','CH' ......
'=================================
Dim wADRESSNUM As String, wADRESSCOA As String
Dim blnOk As Boolean
Dim xSQL As String
Dim V

On Error GoTo Error_Handler
rsZADRESS0_Compte = Null
blnOk = False

wADRESSNUM = lZADRESS0.ADRESSNUM
wADRESSCOA = lZADRESS0.ADRESSCOA
rsZADRESS0_Init lZADRESS0
lZADRESS0.ADRESSNUM = wADRESSNUM
lZADRESS0.ADRESSCOA = wADRESSCOA

'Lecture Adresse avec Code Adresse
'=================================
Set rsSab = Nothing
xSQL = "select * from " & paramIBM_Library_SAB & ".ZADRESS0" _
     & " where ADRESSNUM = '" & wADRESSNUM & "'" _
     & " and  ADRESSCOA = '" & wADRESSCOA _
     & "' and ADRESSTYP = '2'" _
     & " and ADRESSETA = " & currentZMNURUT0.MNURUTETB

Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    V = rsZADRESS0_GetBuffer(rsSab, lZADRESS0)
    If Not IsNull(V) Then
        rsZADRESS0_Compte = "rsZADRESS0_Compte_1 : Lecture ZADRESS0 : " & V
        Exit Function
    Else
        blnOk = True
    End If
End If
If blnOk Then Exit Function

'SINON Lecture Adresse avec type= '  '
'=================================
If wADRESSCOA <> "  " Then
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZADRESS0" _
         & " where ADRESSNUM = '" & wADRESSNUM & "'" _
         & " and  ADRESSCOA = '  '" _
         & " and ADRESSTYP = '2'" _
         & " and ADRESSETA = " & currentZMNURUT0.MNURUTETB
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        V = rsZADRESS0_GetBuffer(rsSab, lZADRESS0)
        If Not IsNull(V) Then
            rsZADRESS0_Compte = "rsZADRESS0_Compte_2 : Lecture ZADRESS0 : " & V
            Exit Function
        Else
            blnOk = True
        End If
    End If
End If

If blnOk Then
    'Lecture CLIENT principal ==> ADRESSRA1
    '=================================
    If Trim(lZADRESS0.ADRESSRA1) = "" Then rsZADRESS0_CLIENARA1 lZADRESS0
Else
    'SINON Lecture TITULAIRE principal ==> ADRESSE CLIENT
    '=================================
    rsZADRESS0_Compte = rsZADRESS0_Titulaire(lZADRESS0)
End If


Exit Function

Error_Handler:
rsZADRESS0_Compte = Error

End Function
'---------------------------------------------------------
Public Function rsZADRESS0_Titulaire(lZADRESS0 As typeZADRESS0)
'---------------------------------------------------------
'Initialiser .ADRESSNUM= 'numéro de compte
'            .ADRESSCOA = '  ','CO','CH' ......

' Lecture Titulaire => Client
'=================================
Dim xSQL As String
Dim V

On Error GoTo Error_Handler
rsZADRESS0_Titulaire = Null

'=================================
xSQL = "select TITULACLI from " & paramIBM_Library_SAB & ".ZTITULA0" _
     & " where TITULACOM = '" & lZADRESS0.ADRESSNUM _
     & "' and  TITULATPR = '0'" _
     & " and TITULAETA = " & currentZMNURUT0.MNURUTETB
     
Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    lZADRESS0.ADRESSNUM = rsSab("TITULACLI")
    Call rsZADRESS0_Client(lZADRESS0)
Else
    rsZADRESS0_Titulaire = "rsZADRESS0_Titulaire_3 : Lecture ZTITULA0 : " & lZADRESS0.ADRESSNUM
End If



Exit Function

Error_Handler:
rsZADRESS0_Titulaire = Error

End Function
'---------------------------------------------------------
Public Function rsZADRESS0_CLIENARA1(lZADRESS0 As typeZADRESS0)
'---------------------------------------------------------

' Lecture Titulaire => Client => Raison Sociale 1 & 2
'========================================================
Dim xSQL As String
Dim V

On Error GoTo Error_Handler
rsZADRESS0_CLIENARA1 = Null

'=================================
xSQL = "select TITULACLI from " & paramIBM_Library_SAB & ".ZTITULA0" _
     & " where TITULACOM = '" & lZADRESS0.ADRESSNUM _
     & "' and  TITULATPR = '0'" _
     & " and TITULAETA = " & currentZMNURUT0.MNURUTETB
     
Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
        xSQL = "select CLIENARA1,CLIENARA2 from " & paramIBM_Library_SAB & ".ZCLIENA0" _
             & " where CLIENACLI = '" & rsSab("TITULACLI") & "'" _
             & " and CLIENAETB = " & currentZMNURUT0.MNURUTETB

        Set rsSab = cnsab.Execute(xSQL)
        If Not rsSab.EOF Then
            lZADRESS0.ADRESSRA1 = rsSab("CLIENARA1")
            lZADRESS0.ADRESSRA2 = rsSab("CLIENARA2")
        End If
End If



Exit Function

Error_Handler:
rsZADRESS0_CLIENARA1 = Error

End Function


'---------------------------------------------------------
Public Function rsZADRESS0_BIC_Compte(lCOMPTECOM As String, lBIC As String)
'---------------------------------------------------------
Dim xSQL As String
Dim wCLIENACLI As String
Dim blnOk As Boolean

On Error GoTo Error_Handler

'Lecture TITULAIRE principal ==> RACINE CLIENT
'=================================

rsZADRESS0_BIC_Compte = Null
xSQL = "select TITULACLI from " & paramIBM_Library_SAB & ".ZTITULA0" _
     & " where TITULACOM = '" & lCOMPTECOM _
     & "' and  TITULATPR = '0'" _
     & " and TITULAETA = " & currentZMNURUT0.MNURUTETB
     
Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    wCLIENACLI = rsSab("TITULACLI")
    rsZADRESS0_BIC_Compte = rsZADRESS0_BIC_Client(wCLIENACLI, lBIC)
End If


Exit Function

Error_Handler:
rsZADRESS0_BIC_Compte = Error

End Function


'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZADRESS0_GetBuffer(rsSab As ADODB.Recordset, rsZADRESS0 As typeZADRESS0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZADRESS0_GetBuffer = Null

rsZADRESS0.ADRESSETA = rsSab("ADRESSETA")
rsZADRESS0.ADRESSTYP = rsSab("ADRESSTYP")
rsZADRESS0.ADRESSPLA = rsSab("ADRESSPLA")
rsZADRESS0.ADRESSNUM = rsSab("ADRESSNUM")
rsZADRESS0.ADRESSCOA = rsSab("ADRESSCOA")
rsZADRESS0.ADRESSDLI = rsSab("ADRESSDLI")
rsZADRESS0.ADRESSDDE = rsSab("ADRESSDDE")
If paramIBM_AS400_ID = "I5A7" Then
    rsZADRESS0.ADRESSRA1 = rsSab("ADRESSRA1")
Else
    rsZADRESS0.ADRESSRA1 = rsSab("ADRESSRA11") & rsSab("ADRESSRA12") & rsSab("ADRESSRA13")
End If
rsZADRESS0.ADRESSRA2 = rsSab("ADRESSRA2")
rsZADRESS0.ADRESSAD1 = rsSab("ADRESSAD1")
rsZADRESS0.ADRESSAD2 = rsSab("ADRESSAD2")
rsZADRESS0.ADRESSAD3 = rsSab("ADRESSAD3")
rsZADRESS0.ADRESSCOP = rsSab("ADRESSCOP")
rsZADRESS0.ADRESSVIL = rsSab("ADRESSVIL")
rsZADRESS0.ADRESSPAY = rsSab("ADRESSPAY")
rsZADRESS0.ADRESSTEL = rsSab("ADRESSTEL")
rsZADRESS0.ADRESSFAX = rsSab("ADRESSFAX")
rsZADRESS0.ADRESSTEX = rsSab("ADRESSTEX")
Exit Function

Error_Handler:

rsZADRESS0_GetBuffer = Error

End Function

'---------------------------------------------------------
Public Sub rsZADRESS0_Init(rsZADRESS0 As typeZADRESS0)
'---------------------------------------------------------
rsZADRESS0.ADRESSETA = 0       ' Integer                        ' Etablissement
rsZADRESS0.ADRESSTYP = ""      ' String * 1                     ' 1 client , 2 compte
rsZADRESS0.ADRESSPLA = 0       ' Long                           ' Numéro de plan
rsZADRESS0.ADRESSNUM = ""      ' String * 20                    ' ou numéro de client
rsZADRESS0.ADRESSCOA = ""      ' String * 2                     ' Code adresse
rsZADRESS0.ADRESSDLI = 0       ' Long                           ' Date limite validité
rsZADRESS0.ADRESSDDE = 0       ' Long                           ' Date début validité
rsZADRESS0.ADRESSRA1 = ""      ' String * 32                    ' ou raison sociale 1
rsZADRESS0.ADRESSRA2 = ""      ' String * 32                    ' ou raison sociale 2
rsZADRESS0.ADRESSAD1 = ""      ' String * 32                    ' Adresse 1
rsZADRESS0.ADRESSAD2 = ""      ' String * 32                    ' Adresse 2
rsZADRESS0.ADRESSAD3 = ""      ' String * 32                    ' Adresse 3
rsZADRESS0.ADRESSCOP = ""      ' String * 6                     ' Code postal
rsZADRESS0.ADRESSVIL = ""      ' String * 25                    ' Ville
rsZADRESS0.ADRESSPAY = ""      ' String * 25                    ' Pays
rsZADRESS0.ADRESSTEL = ""      ' String * 20                    ' No Tel.
rsZADRESS0.ADRESSFAX = ""      ' String * 20                    ' No Fax.
rsZADRESS0.ADRESSTEX = ""      ' String * 20                    ' No Télex

End Sub


'








