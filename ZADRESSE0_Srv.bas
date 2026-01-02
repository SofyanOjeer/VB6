Attribute VB_Name = "srvZADRESSE0"
Option Explicit

'---------------------------------------------------------
Public Function srvYADRESS0_Compte_BIC(lCOMPTECOM As String, lBIC As String)
'---------------------------------------------------------
Dim xSql As String
Dim rsADO As ADODB.Recordset
Dim wCLIENACLI As String
Dim blnOk As Boolean

On Error GoTo Error_Handler

lBIC = ""
'Lecture Compte => Racine
'==========================
srvYADRESS0_Compte_BIC = "? YBIACPT0 : srvYADRESS0_Compte_BIC"
blnOk = False
xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where COMPTECOM = '" & lCOMPTECOM & "'"
Set rsADO = cnSAB.Execute(xSql)
If Not rsADO.EOF Then
    wCLIENACLI = rsADO("CLIENACLI")
    srvYADRESS0_Compte_BIC = "? ZADRESS0 : srvYADRESS0_Compte_BIC"

'Lecture Compte => Racine
    '==========================
    Set rsADO = Nothing
    xSql = "select * from " & paramIBM_Library_SAB & ".ZADRESS0 where ADRESSTYP = '4' and ADRESSNUM = ' " & wCLIENACLI & "'"
    Set rsADO = cnSAB.Execute(xSql)
    If Not rsADO.EOF Then
        lBIC = mId$(rsADO("ADRESSRA1"), 11, 11)
        srvYADRESS0_Compte_BIC = Null
    End If
End If

Exit Function

Error_Handler:
srvYADRESS0_Compte_BIC = Error

End Function

'---------------------------------------------------------
Public Function srvYADRESS0_Client(lYADRESS0 As typeZADRESS0)
'---------------------------------------------------------
Dim wADRESSNUM As String, wADRESSCOA As String
Dim blnOk As Boolean
Dim xSql As String
Dim rsADO As ADODB.Recordset
Dim V
Dim wYCLIENA0 As typeZCLIENA0

On Error GoTo Error_Handler
srvYADRESS0_Client = Null
blnOk = False

wADRESSNUM = " " & Trim(lYADRESS0.ADRESSNUM)
wADRESSCOA = lYADRESS0.ADRESSCOA
rsZADRESS0_Init lYADRESS0

'Lecture Adresse avec Code Adresse
'=================================
Set rsADO = Nothing
xSql = "select * from ZADRESS0 where ADRESSNUM = '" & wADRESSNUM & "' AND  ADRESSCOA = '" & wADRESSCOA _
        & "' AND ADRESSTYP = '1'"
Set rsADO = cnSAB.Execute(xSql)

If Not rsADO.EOF Then
    V = rsZADRESS0_GetBuffer(rsADO, lYADRESS0)
    If Not IsNull(V) Then
        srvYADRESS0_Client = "srvYADRESS0_Client_1 : Lecture ZADRESS0 : " & V
        Exit Function
    Else
        blnOk = True
    End If
End If

If Not blnOk Then
    'SINON Lecture Adresse avec type= '  '
    '=================================
    If wADRESSCOA <> "  " Then
        xSql = "select * from ZADRESS0 where ADRESSNUM = '" & wADRESSNUM & "' AND  ADRESSCOA = '  '" _
                & " AND ADRESSTYP = '1'"
        Set rsADO = cnSAB.Execute(xSql)
        If Not rsADO.EOF Then
            V = rsZADRESS0_GetBuffer(rsADO, lYADRESS0)
            If Not IsNull(V) Then
                srvYADRESS0_Client = "srvYADRESS0_Client_2 : Lecture ZADRESS0 : " & V
                Exit Function
            Else
                blnOk = True
            End If
        End If
    End If
End If

If Not blnOk Then
    srvYADRESS0_Client = "Adresse non trouvée  : " & wADRESSNUM
Else
    If Trim(lYADRESS0.ADRESSRA1) = "" Then
        'Lecture CLIENT principal ==> ADRESSRA1
        '=================================
        xSql = "select * from ZCLIENA0 where CLIENACLI = '" & Trim(wADRESSNUM) & "'"
        Set rsADO = cnSAB.Execute(xSql)
        If Not rsADO.EOF Then
            V = rsZCLIENA0_GetBuffer(rsADO, wYCLIENA0)
            If Not IsNull(V) Then
                srvYADRESS0_Client = "srvYADRESS0_Client_3 : Lecture ZCLIENT0 : " & V
                Exit Function
            Else
                lYADRESS0.ADRESSRA1 = wYCLIENA0.CLIENARA1
                lYADRESS0.ADRESSRA2 = wYCLIENA0.CLIENARA2
            End If
        End If
    End If
End If

Exit Function

Error_Handler:
srvYADRESS0_Client = Error

End Function


'---------------------------------------------------------
Public Function srvYADRESS0_Compte(lYADRESS0 As typeZADRESS0)
'---------------------------------------------------------
'Initialiser .ADRESSNUM= 'numéro de compte
'            .ADRESSCOA = '  ','CO','CH' ......
'=================================
Dim wADRESSNUM As String, wADRESSCOA As String
Dim blnOk As Boolean
Dim xSql As String
Dim rsADO As ADODB.Recordset
Dim V
Dim wYTITULA0 As typeZTITULA0

On Error GoTo Error_Handler
srvYADRESS0_Compte = Null
blnOk = False

wADRESSNUM = lYADRESS0.ADRESSNUM
wADRESSCOA = lYADRESS0.ADRESSCOA
rsZADRESS0_Init lYADRESS0

'Lecture Adresse avec Code Adresse
'=================================
Set rsADO = Nothing
xSql = "select * from ZADRESS0 where ADRESSNUM = '" & wADRESSNUM & "' AND  ADRESSCOA = '" & wADRESSCOA _
        & "' AND ADRESSTYP = '2'"
Set rsADO = cnSAB.Execute(xSql)
If Not rsADO.EOF Then
    V = rsZADRESS0_GetBuffer(rsADO, lYADRESS0)
    If Not IsNull(V) Then
        srvYADRESS0_Compte = "srvYADRESS0_Compte_1 : Lecture ZADRESS0 : " & V
        Exit Function
    Else
        blnOk = True
    End If
End If
If blnOk Then Exit Function

'SINON Lecture Adresse avec type= '  '
'=================================
If wADRESSCOA <> "  " Then
    xSql = "select * from ZADRESS0 where ADRESSNUM = '" & wADRESSNUM & "' AND  ADRESSCOA = '  '" _
            & " AND ADRESSTYP = '2'"
    Set rsADO = cnSAB.Execute(xSql)
    If Not rsADO.EOF Then
        V = rsZADRESS0_GetBuffer(rsADO, lYADRESS0)
        If Not IsNull(V) Then
            srvYADRESS0_Compte = "srvYADRESS0_Compte_2 : Lecture ZADRESS0 : " & V
            Exit Function
        Else
            blnOk = True
        End If
    End If
End If

If blnOk Then Exit Function

'SINON Lecture TITULAIRE principal ==> ADRESSE CLIENT
'=================================
xSql = "select * from ZTITULA0 where TITULACOM = '" & wADRESSNUM & "' AND  TITULATPR = '0'"
Set rsADO = cnSAB.Execute(xSql)
If Not rsADO.EOF Then
    V = rsZTITULA0_GetBuffer(rsADO, wYTITULA0)
    If Not IsNull(V) Then
        srvYADRESS0_Compte = "srvYADRESS0_Compte_3 : Lecture ZTITULA0 : " & V
        Exit Function
    Else
        blnOk = True
    End If
End If

lYADRESS0.ADRESSNUM = wYTITULA0.TITULACLI
lYADRESS0.ADRESSCOA = wADRESSCOA

Call srvYADRESS0_Client(lYADRESS0)

If blnOk Then Exit Function
srvYADRESS0_Compte = "Adresse non trouvée  : " & wADRESSNUM & " & " & wYTITULA0.TITULACLI

Exit Function

Error_Handler:
srvYADRESS0_Compte = Error

End Function

