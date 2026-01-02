Attribute VB_Name = "srvYPDCOPE0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYPDCOPE0
 
      PDCOPEDTR     As String * 8   'date comptable
      PDCOPEID      As Long         'identification
      
      PDCOPEREF     As Long         'réf de l'opération
      PDCOPEOPEC    As String * 3   'code opération
      PDCOPEOPEN    As Long         'n° opération
      PDCOPEOPET    As String * 6   'type / nature
      PDCOPESENS    As String * 1   'sens A|V
      PDCOPESENX    As String * 1   'devise principale 1|2
      PDCOPEDEV1    As String * 3   'devise 1
      PDCOPEMTD1    As Currency     'montant en devise 2
      PDCOPEDEV2    As String * 3   'devise 2
      PDCOPEMTD2    As Currency     'montant en devise 1
      PDCOPETAUX    As Double       'taux
      PDCOPEDVA     As String * 8   'date valeur
      PDCOPECLI     As String * 7   'client
      PDCOPESTA     As String * 1   'statut
      PDCOPESTA2    As String * 1   'statut
      PDCOPESTA3    As String * 1   'statut
      PDCOPESER     As String * 2   'service
      PDCOPESSE     As String * 2   'sous-service
      PDCOPEIAMJ    As String * 8   'date saisie
      PDCOPEIHMS    As String * 6   'heure saisie
      PDCOPEIUSR    As String * 12  'utilisateur saisie
      PDCOPEITXT    As String * 64  'commentaire saisie
      PDCOPEVAMJ    As String * 8   'date validation
      PDCOPEVHMS    As String * 6   'heure validation
      PDCOPEVUSR    As String * 12  'utilisateur validation
      PDCOPEVTXT    As String * 64  'commentaire validation
      PDCOPEUPDS    As Long         'séq màj
End Type

Public Function sqlYPDCOPE0_DeleteW(lWhere As String, Nb As Long)
Dim X As String, xSQL As String

On Error GoTo Error_Handler
sqlYPDCOPE0_DeleteW = Null
    
xSQL = "delete from " & paramIBM_Library_SABSPE_XXX & ".YPDCOPE0 " & lWhere
Call FEU_ROUGE
Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
Exit Function
Error_Handler:
    sqlYPDCOPE0_DeleteW = Error
End Function

Public Function sqlYPDCOPE0_Insert(newY As typeYPDCOPE0)
Dim V
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYPDCOPE0_Insert = Null
xSet = " (PDCOPEDTR,PDCOPEID"
xValues = " values('" & newY.PDCOPEDTR & "'," & newY.PDCOPEID


' Détecter les modifications
'===================================================================================
If newY.PDCOPEREF <> 0 Then xSet = xSet & ",PDCOPEREF": xValues = xValues & " ," & newY.PDCOPEREF
If newY.PDCOPEOPEN <> 0 Then xSet = xSet & ",PDCOPEOPEN": xValues = xValues & " ," & newY.PDCOPEOPEN
If newY.PDCOPEMTD1 <> 0 Then xSet = xSet & ",PDCOPEMTD1": xValues = xValues & " ," & Replace(newY.PDCOPEMTD1, ",", ".")
If newY.PDCOPEMTD2 <> 0 Then xSet = xSet & ",PDCOPEMTD2": xValues = xValues & " ," & Replace(newY.PDCOPEMTD2, ",", ".")
If newY.PDCOPETAUX <> 0 Then xSet = xSet & ",PDCOPETAUX": xValues = xValues & " ," & Replace(newY.PDCOPETAUX, ",", ".")

If Trim(newY.PDCOPEOPEC) <> "" Then xSet = xSet & ",PDCOPEOPEC": xValues = xValues & " ,'" & newY.PDCOPEOPEC & "'"
If Trim(newY.PDCOPEOPET) <> "" Then xSet = xSet & ",PDCOPEOPET": xValues = xValues & " ,'" & newY.PDCOPEOPET & "'"
If Trim(newY.PDCOPESENS) <> "" Then xSet = xSet & ",PDCOPESENS": xValues = xValues & " ,'" & newY.PDCOPESENS & "'"
If Trim(newY.PDCOPESENX) <> "" Then xSet = xSet & ",PDCOPESENX": xValues = xValues & " ,'" & newY.PDCOPESENX & "'"
If Trim(newY.PDCOPEDEV1) <> "" Then xSet = xSet & ",PDCOPEDEV1": xValues = xValues & " ,'" & newY.PDCOPEDEV1 & "'"
If Trim(newY.PDCOPEDEV2) <> "" Then xSet = xSet & ",PDCOPEDEV2": xValues = xValues & " ,'" & newY.PDCOPEDEV2 & "'"
If Trim(newY.PDCOPEDVA) <> "" Then xSet = xSet & ",PDCOPEDVA": xValues = xValues & " ,'" & newY.PDCOPEDVA & "'"
If Trim(newY.PDCOPECLI) <> "" Then xSet = xSet & ",PDCOPECLI": xValues = xValues & " ,'" & newY.PDCOPECLI & "'"
If Trim(newY.PDCOPESTA) <> "" Then xSet = xSet & ",PDCOPESTA": xValues = xValues & " ,'" & newY.PDCOPESTA & "'"
If Trim(newY.PDCOPESTA2) <> "" Then xSet = xSet & ",PDCOPESTA2": xValues = xValues & " ,'" & newY.PDCOPESTA2 & "'"
If Trim(newY.PDCOPESER) <> "" Then xSet = xSet & ",PDCOPESER": xValues = xValues & " ,'" & newY.PDCOPESER & "'"
If Trim(newY.PDCOPESSE) <> "" Then xSet = xSet & ",PDCOPESSE": xValues = xValues & " ,'" & newY.PDCOPESSE & "'"
If Trim(newY.PDCOPEIAMJ) <> "" Then xSet = xSet & ",PDCOPEIAMJ": xValues = xValues & " ,'" & newY.PDCOPEIAMJ & "'"
If Trim(newY.PDCOPEIHMS) <> "" Then xSet = xSet & ",PDCOPEIHMS": xValues = xValues & " ,'" & newY.PDCOPEIHMS & "'"
If Trim(newY.PDCOPEIUSR) <> "" Then xSet = xSet & ",PDCOPEIUSR": xValues = xValues & " ,'" & newY.PDCOPEIUSR & "'"
If Trim(newY.PDCOPEITXT) <> "" Then xSet = xSet & ",PDCOPEITXT": xValues = xValues & " ,'" & Replace(Trim(newY.PDCOPEITXT), "'", "''") & "'"
If Trim(newY.PDCOPEVAMJ) <> "" Then xSet = xSet & ",PDCOPEVAMJ": xValues = xValues & " ,'" & newY.PDCOPEVAMJ & "'"
If Trim(newY.PDCOPEVHMS) <> "" Then xSet = xSet & ",PDCOPEVHMS": xValues = xValues & " ,'" & newY.PDCOPEVHMS & "'"
If Trim(newY.PDCOPEVUSR) <> "" Then xSet = xSet & ",PDCOPEVUSR": xValues = xValues & " ,'" & newY.PDCOPEVUSR & "'"
If Trim(newY.PDCOPEVTXT) <> "" Then xSet = xSet & ",PDCOPEVTXT": xValues = xValues & " ,'" & Replace(Trim(newY.PDCOPEVTXT), "'", "''") & "'"

Call FEU_ROUGE
xSQL = "Insert into " & paramIBM_Library_SABSPE_XXX & ".YPDCOPE0" & xSet & ")" & xValues & ")"
Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYPDCOPE0_Insert = "Erreur màj : " & newY.PDCOPEDTR & " " & newY.PDCOPEID & " " & newY.PDCOPEREF
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYPDCOPE0_Insert = Error
End Function


Public Function sqlYPDCOPE0_Update(newY As typeYPDCOPE0, oldY As typeYPDCOPE0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYPDCOPE0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.PDCOPEID <> newY.PDCOPEID _
Or oldY.PDCOPEDTR <> newY.PDCOPEDTR _
Or oldY.PDCOPEUPDS <> newY.PDCOPEUPDS Then
    sqlYPDCOPE0_Update = "Erreur PDCOPEID : " & newY.PDCOPEID & "." & oldY.PDCOPEUPDS
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where PDCOPEDTR = '" & oldY.PDCOPEDTR & "' And PDCOPEID = " & oldY.PDCOPEID _
       & " and PDCOPEUPDS = " & oldY.PDCOPEUPDS

newY.PDCOPEUPDS = newY.PDCOPEUPDS + 1
xSet = xSet & " set PDCOPEUPDS = " & newY.PDCOPEUPDS
blnUpdate = False


' Détecter les modifications
'===================================================================================
If newY.PDCOPEREF <> oldY.PDCOPEREF Then blnUpdate = True:  xSet = xSet & " , PDCOPEREF = " & newY.PDCOPEREF
If newY.PDCOPEOPEN <> oldY.PDCOPEOPEN Then blnUpdate = True:  xSet = xSet & " , PDCOPEOPEN = " & newY.PDCOPEOPEN
If newY.PDCOPEMTD1 <> oldY.PDCOPEMTD1 Then blnUpdate = True:  xSet = xSet & " , PDCOPEMTD1 = " & Replace(newY.PDCOPEMTD1, ",", ".")
If newY.PDCOPEMTD2 <> oldY.PDCOPEMTD2 Then blnUpdate = True:  xSet = xSet & " , PDCOPEMTD2 = " & Replace(newY.PDCOPEMTD2, ",", ".")
If newY.PDCOPETAUX <> oldY.PDCOPETAUX Then blnUpdate = True:  xSet = xSet & " , PDCOPETAUX = " & Replace(newY.PDCOPETAUX, ",", ".")


If newY.PDCOPEOPEC <> oldY.PDCOPEOPEC Then blnUpdate = True:  xSet = xSet & " , PDCOPEOPEC = '" & Replace(Trim(newY.PDCOPEOPEC), "'", "''") & "'"
If newY.PDCOPEOPET <> oldY.PDCOPEOPET Then blnUpdate = True:  xSet = xSet & " , PDCOPEOPET = '" & Replace(Trim(newY.PDCOPEOPET), "'", "''") & "'"
If newY.PDCOPESENS <> oldY.PDCOPESENS Then blnUpdate = True:  xSet = xSet & " , PDCOPESENS = '" & Replace(Trim(newY.PDCOPESENS), "'", "''") & "'"
If newY.PDCOPESENX <> oldY.PDCOPESENX Then blnUpdate = True:  xSet = xSet & " , PDCOPESENX = '" & Replace(Trim(newY.PDCOPESENX), "'", "''") & "'"
If newY.PDCOPEDEV1 <> oldY.PDCOPEDEV1 Then blnUpdate = True:  xSet = xSet & " , PDCOPEDEV1 = '" & Replace(Trim(newY.PDCOPEDEV1), "'", "''") & "'"
If newY.PDCOPEDEV2 <> oldY.PDCOPEDEV2 Then blnUpdate = True:  xSet = xSet & " , PDCOPEDEV2 = '" & Replace(Trim(newY.PDCOPEDEV2), "'", "''") & "'"
If newY.PDCOPEDVA <> oldY.PDCOPEDVA Then blnUpdate = True:  xSet = xSet & " , PDCOPEDVA = '" & Replace(Trim(newY.PDCOPEDVA), "'", "''") & "'"
If newY.PDCOPECLI <> oldY.PDCOPECLI Then blnUpdate = True:  xSet = xSet & " , PDCOPECLI = '" & Replace(Trim(newY.PDCOPECLI), "'", "''") & "'"
If newY.PDCOPESTA <> oldY.PDCOPESTA Then blnUpdate = True:  xSet = xSet & " , PDCOPESTA = '" & Replace(Trim(newY.PDCOPESTA), "'", "''") & "'"
If newY.PDCOPESTA2 <> oldY.PDCOPESTA2 Then blnUpdate = True:  xSet = xSet & " , PDCOPESTA2 = '" & Replace(Trim(newY.PDCOPESTA2), "'", "''") & "'"
If newY.PDCOPESTA3 <> oldY.PDCOPESTA3 Then blnUpdate = True:  xSet = xSet & " , PDCOPESTA3 = '" & Replace(Trim(newY.PDCOPESTA3), "'", "''") & "'"
If newY.PDCOPESER <> oldY.PDCOPESER Then blnUpdate = True:  xSet = xSet & " , PDCOPESER = '" & Replace(Trim(newY.PDCOPESER), "'", "''") & "'"
If newY.PDCOPESSE <> oldY.PDCOPESSE Then blnUpdate = True:  xSet = xSet & " , PDCOPESSE = '" & Replace(Trim(newY.PDCOPESSE), "'", "''") & "'"
If newY.PDCOPEIAMJ <> oldY.PDCOPEIAMJ Then blnUpdate = True:  xSet = xSet & " , PDCOPEIAMJ = '" & Replace(Trim(newY.PDCOPEIAMJ), "'", "''") & "'"
If newY.PDCOPEIHMS <> oldY.PDCOPEIHMS Then blnUpdate = True:  xSet = xSet & " , PDCOPEIHMS = '" & Replace(Trim(newY.PDCOPEIHMS), "'", "''") & "'"
If newY.PDCOPEIUSR <> oldY.PDCOPEIUSR Then blnUpdate = True:  xSet = xSet & " , PDCOPEIUSR = '" & Replace(Trim(newY.PDCOPEIUSR), "'", "''") & "'"
If newY.PDCOPEITXT <> oldY.PDCOPEITXT Then blnUpdate = True:  xSet = xSet & " , PDCOPEITXT = '" & Replace(Trim(newY.PDCOPEITXT), "'", "''") & "'"
If newY.PDCOPEVAMJ <> oldY.PDCOPEVAMJ Then blnUpdate = True:  xSet = xSet & " , PDCOPEVAMJ = '" & Replace(Trim(newY.PDCOPEVAMJ), "'", "''") & "'"
If newY.PDCOPEVHMS <> oldY.PDCOPEVHMS Then blnUpdate = True:  xSet = xSet & " , PDCOPEVHMS = '" & Replace(Trim(newY.PDCOPEVHMS), "'", "''") & "'"
If newY.PDCOPEVUSR <> oldY.PDCOPEVUSR Then blnUpdate = True:  xSet = xSet & " , PDCOPEVUSR = '" & Replace(Trim(newY.PDCOPEVUSR), "'", "''") & "'"
If newY.PDCOPEVTXT <> oldY.PDCOPEVTXT Then blnUpdate = True:  xSet = xSet & " , PDCOPEVTXT = '" & Replace(Trim(newY.PDCOPEVTXT), "'", "''") & "'"


If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE_XXX & ".YPDCOPE0" & xSet & xWhere
    Call FEU_ROUGE
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYPDCOPE0_Update = "Erreur màj : " & newY.PDCOPEID
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYPDCOPE0_Update = Error
End Function


Public Function rsYPDCOPE0_GetBuffer(rsAdo As ADODB.Recordset, lYPDCOPE0 As typeYPDCOPE0)
On Error GoTo Error_Handler
rsYPDCOPE0_GetBuffer = Null

lYPDCOPE0.PDCOPEDTR = rsAdo("PDCOPEDTR")
lYPDCOPE0.PDCOPEID = rsAdo("PDCOPEID")
lYPDCOPE0.PDCOPEREF = rsAdo("PDCOPEREF")
lYPDCOPE0.PDCOPEOPEC = rsAdo("PDCOPEOPEC")
lYPDCOPE0.PDCOPEOPEN = rsAdo("PDCOPEOPEN")
lYPDCOPE0.PDCOPEOPET = rsAdo("PDCOPEOPET")
lYPDCOPE0.PDCOPESENS = rsAdo("PDCOPESENS")
lYPDCOPE0.PDCOPESENX = rsAdo("PDCOPESENX")
lYPDCOPE0.PDCOPEDEV1 = rsAdo("PDCOPEDEV1")
lYPDCOPE0.PDCOPEMTD1 = rsAdo("PDCOPEMTD1")
lYPDCOPE0.PDCOPEDEV2 = rsAdo("PDCOPEDEV2")
lYPDCOPE0.PDCOPEMTD2 = rsAdo("PDCOPEMTD2")
lYPDCOPE0.PDCOPETAUX = rsAdo("PDCOPETAUX")
lYPDCOPE0.PDCOPEDVA = rsAdo("PDCOPEDVA")
lYPDCOPE0.PDCOPECLI = rsAdo("PDCOPECLI")
lYPDCOPE0.PDCOPESTA = rsAdo("PDCOPESTA")
lYPDCOPE0.PDCOPESTA2 = rsAdo("PDCOPESTA2")
lYPDCOPE0.PDCOPESTA3 = rsAdo("PDCOPESTA3")
lYPDCOPE0.PDCOPESER = rsAdo("PDCOPESER")
lYPDCOPE0.PDCOPESSE = rsAdo("PDCOPESSE")
lYPDCOPE0.PDCOPEIAMJ = rsAdo("PDCOPEIAMJ")
lYPDCOPE0.PDCOPEIHMS = rsAdo("PDCOPEIHMS")
lYPDCOPE0.PDCOPEIUSR = rsAdo("PDCOPEIUSR")
lYPDCOPE0.PDCOPEITXT = rsAdo("PDCOPEITXT")
lYPDCOPE0.PDCOPEVAMJ = rsAdo("PDCOPEVAMJ")
lYPDCOPE0.PDCOPEVHMS = rsAdo("PDCOPEVHMS")
lYPDCOPE0.PDCOPEVUSR = rsAdo("PDCOPEVUSR")
lYPDCOPE0.PDCOPEVTXT = rsAdo("PDCOPEVTXT")
lYPDCOPE0.PDCOPEUPDS = rsAdo("PDCOPEUPDS")


Exit Function
Error_Handler:
rsYPDCOPE0_GetBuffer = Error


End Function

Public Function rsYPDCOPE0_Init(lYPDCOPE0 As typeYPDCOPE0)

lYPDCOPE0.PDCOPEDTR = ""
lYPDCOPE0.PDCOPEID = 0
lYPDCOPE0.PDCOPEREF = 0
lYPDCOPE0.PDCOPEOPEC = ""
lYPDCOPE0.PDCOPEOPEN = 0
lYPDCOPE0.PDCOPEOPET = ""
lYPDCOPE0.PDCOPESENS = ""
lYPDCOPE0.PDCOPESENX = ""
lYPDCOPE0.PDCOPEDEV1 = ""
lYPDCOPE0.PDCOPEMTD1 = 0
lYPDCOPE0.PDCOPEDEV2 = ""
lYPDCOPE0.PDCOPEMTD2 = 0
lYPDCOPE0.PDCOPETAUX = 0
lYPDCOPE0.PDCOPEDVA = ""
lYPDCOPE0.PDCOPECLI = ""
lYPDCOPE0.PDCOPESTA = ""
lYPDCOPE0.PDCOPESTA2 = ""
lYPDCOPE0.PDCOPESTA3 = ""
lYPDCOPE0.PDCOPESER = ""
lYPDCOPE0.PDCOPESSE = ""
lYPDCOPE0.PDCOPEIAMJ = ""
lYPDCOPE0.PDCOPEIHMS = ""
lYPDCOPE0.PDCOPEIUSR = ""
lYPDCOPE0.PDCOPEITXT = ""
lYPDCOPE0.PDCOPEVAMJ = ""
lYPDCOPE0.PDCOPEVHMS = ""
lYPDCOPE0.PDCOPEVUSR = ""
lYPDCOPE0.PDCOPEVTXT = ""
lYPDCOPE0.PDCOPEUPDS = 0

End Function





