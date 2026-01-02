Attribute VB_Name = "srvYTVANIF0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYTVANIF0
 
      TVANIFCLIC  As String * 1   'table client espace,G, D Bic
      TVANIFCLI   As String * 11   'code client
      TVANIFCLIF  As String * 1   'alias ' ' ou '='
      TVANIFCLIT  As String * 18  'code TVA intracommunautaire
      TVANIFSTA   As String * 1   'statut
      TVANIFUPDS  As Long         'SéQUENCE UPD
      TVANIFUSR   As String * 10   'user
'________________________________________________________ pour compatibilité fgNIF
      TVANIFRS    As String
      TVANIFCLIP  As String
      
End Type
Public xYTVANIF0 As typeYTVANIF0

Public Function sqlYTVANIF0_Insert(newY As typeYTVANIF0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYTVANIF0_Insert = Null

xSet = " (TVANIFCLIC,TVANIFCLI"
xValues = " values('" & newY.TVANIFCLIC & "','" & Format(newY.TVANIFCLI, "0000000") & "'"

' Détecter les modifications
'===================================================================================

If Trim(newY.TVANIFCLIF) <> "" Then xSet = xSet & ",TVANIFCLIF": xValues = xValues & " ,'" & newY.TVANIFCLIF & "'"
If Trim(newY.TVANIFCLIT) <> "" Then xSet = xSet & ",TVANIFCLIT": xValues = xValues & " ,'" & Replace(Trim(newY.TVANIFCLIT), "'", "''") & "'"
If Trim(newY.TVANIFSTA) <> "" Then xSet = xSet & ",TVANIFSTA": xValues = xValues & " ,'" & newY.TVANIFSTA & "'"

newY.TVANIFUSR = usrName_UCase10
xSet = xSet & ",TVANIFUSR": xValues = xValues & " ,'" & usrName_UCase10 & "'"
Call FEU_ROUGE
xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YTVANIF0" & xSet & ")" & xValues & ")"

Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYTVANIF0_Insert = "Erreur màj : " & newY.TVANIFCLI
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYTVANIF0_Insert = Error
End Function

Public Function sqlYTVANIF0_Update(newY As typeYTVANIF0, oldY As typeYTVANIF0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYTVANIF0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.TVANIFCLIC <> newY.TVANIFCLIC _
Or oldY.TVANIFCLI <> newY.TVANIFCLI Then
    sqlYTVANIF0_Update = "Erreur TVANIFCLI: " & newY.TVANIFCLI & "." & oldY.TVANIFCLI
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where TVANIFCLIC = '" & oldY.TVANIFCLIC & "'" _
       & " and TVANIFCLI = '" & oldY.TVANIFCLI & "'" _
       & " and TVANIFUPDS = " & oldY.TVANIFUPDS

newY.TVANIFUPDS = newY.TVANIFUPDS + 1
xSet = xSet & " set TVANIFUPDS = " & newY.TVANIFUPDS
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.TVANIFCLIF <> oldY.TVANIFCLIF Then blnUpdate = True:  xSet = xSet & " , TVANIFCLIF = '" & Replace(Trim(newY.TVANIFCLIF), "'", "''") & "'"
If newY.TVANIFCLIT <> oldY.TVANIFCLIT Then blnUpdate = True:  xSet = xSet & " , TVANIFCLIT = '" & Replace(Trim(newY.TVANIFCLIT), "'", "''") & "'"
If newY.TVANIFSTA <> oldY.TVANIFSTA Then blnUpdate = True:  xSet = xSet & " , TVANIFSTA = '" & newY.TVANIFSTA & "'"

newY.TVANIFUSR = usrName_UCase10
xSet = xSet & " , TVANIFUSR = '" & usrName_UCase10 & "'"

If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE & ".YTVANIF0" & xSet & xWhere
    Call FEU_ROUGE
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYTVANIF0_Update = "Erreur màj : " & newY.TVANIFCLI
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYTVANIF0_Update = Error
End Function

Public Function sqlYTVANIF0_Delete(oldY As typeYTVANIF0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYTVANIF0_Delete = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.TVANIFCLIC <> oldY.TVANIFCLIC _
Or oldY.TVANIFCLI <> oldY.TVANIFCLI Then
    sqlYTVANIF0_Delete = "Erreur TVANIFCLI: " & oldY.TVANIFCLI & "." & oldY.TVANIFCLI
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where TVANIFCLIC = '" & oldY.TVANIFCLIC & "'" _
       & " and TVANIFCLI = '" & Trim(oldY.TVANIFCLI) & "'" _
       & " and TVANIFUPDS = " & oldY.TVANIFUPDS

    
    xSQL = "delete from " & paramIBM_Library_SABSPE & ".YTVANIF0" & xWhere
    Call FEU_ROUGE
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYTVANIF0_Delete = "Erreur màj : " & oldY.TVANIFCLI
        Exit Function
    End If
    


Exit Function
Error_Handler:
    sqlYTVANIF0_Delete = Error
End Function


Public Function rsYTVANIF0_GetBuffer(rsAdo As ADODB.Recordset, lYTVANIF0 As typeYTVANIF0)
On Error GoTo Error_Handler
rsYTVANIF0_GetBuffer = Null

lYTVANIF0.TVANIFCLIC = rsAdo("TVANIFCLIC")
lYTVANIF0.TVANIFCLI = rsAdo("TVANIFCLI")
lYTVANIF0.TVANIFCLIF = rsAdo("TVANIFCLIF")
lYTVANIF0.TVANIFCLIT = rsAdo("TVANIFCLIT")
lYTVANIF0.TVANIFSTA = rsAdo("TVANIFSTA")
lYTVANIF0.TVANIFUPDS = rsAdo("TVANIFUPDS")
lYTVANIF0.TVANIFUSR = rsAdo("TVANIFUSR")

lYTVANIF0.TVANIFRS = ""
lYTVANIF0.TVANIFCLIP = ""

Exit Function
Error_Handler:
rsYTVANIF0_GetBuffer = Error


End Function

Public Function rsYTVANIF0_Init(lYTVANIF0 As typeYTVANIF0)

lYTVANIF0.TVANIFCLIC = ""
lYTVANIF0.TVANIFCLI = ""      '
lYTVANIF0.TVANIFCLIF = ""
lYTVANIF0.TVANIFCLIT = ""
lYTVANIF0.TVANIFSTA = ""      ' 1   'statut
lYTVANIF0.TVANIFUPDS = 0
lYTVANIF0.TVANIFUSR = ""

lYTVANIF0.TVANIFRS = ""
lYTVANIF0.TVANIFCLIP = ""

End Function





Public Function TVANIFCLIT_Format(lX As String)
Select Case Mid$(lX, 1, 2)
    Case "FR": TVANIFCLIT_Format = Format(lX, "@@ @@ @@@ @@@ @@@")
    Case Else: TVANIFCLIT_Format = Format(lX, "@@ @@@ @@@ @@@ @@@")
End Select
End Function

Public Function TVANIFCLIT_Control(lX As String)
Dim lenX As Integer, X As String, xPays As String
Dim blnPaysIsNumeric As Boolean
X = Trim(lX)
lenX = Len(X)
If lenX < 4 Then TVANIFCLIT_Control = "? NIF : préciser le numéro": Exit Function
TVANIFCLIT_Control = Null
xPays = Mid$(X, 1, 2)
If xPays = "AT" Or xPays = "CH" Then
    xPays = Mid$(X, 1, 3)
    X = Mid$(X, 4, lenX - 3)
Else
    X = Mid$(X, 3, lenX - 2)
End If

If Not IsNumeric(X) Then
    If xPays = "FR" Or xPays = "CY" Or xPays = "ES" Or xPays = "IE" Or xPays = "NL" Or xPays = "GB" Then
    Else
        TVANIFCLIT_Control = "? NIF : le numéro doit être numérique": Exit Function
    End If
End If

lenX = Len(X)

Select Case xPays
    Case "FR": If lenX <> 11 Then TVANIFCLIT_Control = "? NIF : 'FR**' suivi de 9 chiffres"
    Case "DE": If lenX <> 9 Then TVANIFCLIT_Control = "? NIF : 'DE' suivi de 9 chiffres"
    Case "ATU": If lenX <> 8 Then TVANIFCLIT_Control = "? NIF : 'ATU' suivi de 8 chiffres"
    Case "BE": If lenX <> 10 Then TVANIFCLIT_Control = "? NIF : 'BE' suivi de 10 chiffres"
    Case "BG": If lenX <> 9 Then TVANIFCLIT_Control = "? NIF : 'BG' suivi de 9 chiffres"
    Case "CY": If lenX <> 9 Then TVANIFCLIT_Control = "? NIF : 'CY' suivi de 9 caractères"
    Case "DK": If lenX <> 8 Then TVANIFCLIT_Control = "? NIF : 'DK' suivi de 8 chiffres"
    Case "ES": If lenX <> 9 Then TVANIFCLIT_Control = "? NIF : 'ES' suivi de 9 caractères"
    Case "EE": If lenX <> 9 Then TVANIFCLIT_Control = "? NIF : 'EE' suivi de 9 chiffres"
    Case "FI": If lenX <> 8 Then TVANIFCLIT_Control = "? NIF : 'FI' suivi de 8 chiffres"
    Case "EL": If lenX <> 9 Then TVANIFCLIT_Control = "? NIF : 'EL' suivi de 9 chiffres"
    Case "HU": If lenX <> 8 Then TVANIFCLIT_Control = "? NIF : 'HU' suivi de 8 chiffres"
    Case "IE": If lenX <> 8 Then TVANIFCLIT_Control = "? NIF : 'IE' suivi de 8 caractères"
    Case "IT": If lenX <> 11 Then TVANIFCLIT_Control = "? NIF : 'IT' suivi de 11 chiffres"
    Case "LV": If lenX <> 11 Then TVANIFCLIT_Control = "? NIF : 'LV' suivi de 11 chiffres"
    Case "LT": If lenX <> 9 And lenX <> 12 Then TVANIFCLIT_Control = "? NIF : 'LT' suivi de 9/12 chiffres"
    Case "LU": If lenX <> 8 Then TVANIFCLIT_Control = "? NIF : 'LU' suivi de 8 chiffres"
    Case "MT": If lenX <> 8 Then TVANIFCLIT_Control = "? NIF : 'MT' suivi de 8 chiffres"
    Case "NL": If lenX <> 12 Then TVANIFCLIT_Control = "? NIF : 'NL' suivi de 12 caractères"
    Case "PL": If lenX <> 10 Then TVANIFCLIT_Control = "? NIF : 'PL' suivi de 10 chiffres"
    Case "PT": If lenX <> 9 Then TVANIFCLIT_Control = "? NIF : 'PT' suivi de 9 chiffres"
    Case "SK": If lenX <> 10 Then TVANIFCLIT_Control = "? NIF : 'SK' suivi de 10 chiffres"
    Case "CZ": If lenX <> 8 And lenX <> 9 And lenX <> 10 Then TVANIFCLIT_Control = "? NIF : 'CZ' suivi de 8/9/10 chiffres"
    Case "RO": If lenX > 10 Then TVANIFCLIT_Control = "? NIF : 'RO' suivi de 2-10 chiffres"
    Case "GB": If lenX <> 5 And lenX <> 9 And lenX <> 12 Then TVANIFCLIT_Control = "? NIF : 'GB' suivi de 5/9/12 caractères"
    Case "SI": If lenX <> 8 Then TVANIFCLIT_Control = "? NIF : 'SI' suivi de 8 chiffres"
    Case "SE": If lenX <> 12 Then TVANIFCLIT_Control = "? NIF : 'SE' suivi de 12 chiffres"

End Select

End Function
