Attribute VB_Name = "srvCHQ_SCAN"
 '---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeCHQ_SCAN
 
      Id            As String
      Cmc7          As String
      Zone4         As String
      Zone3         As String
      Zone2         As String
      Zone1         As String
      PATH          As String
      IMAGE         As String
      Date          As String
      COMPTE        As String
      NumLot        As String
      CRem          As String
      Zone24        As String
      DateHourScan  As String
      DateHourSaisie  As String
      StatutRem     As String
      MotifNonAJ    As String
      Saisie        As String
      RefClient     As String
      RefInterne    As String
      Nature        As String
      Devise        As String
      
      Adresse0      As String
      Adresse1      As String
      Adresse2      As String
      Adresse3      As String
      Adresse4      As String
      Adresse5      As String
      RLMC          As String
      CodeDevise    As String

End Type
Dim xCHQ_SCAN As typeCHQ_SCAN

Public paramCHQ_SCAN_Image_Local As String
Public paramCHQ_SCAN_Image_Archive As String
Public paramCHQ_SCAN_Image_Folder As String

Public paramCHQ_SCAN_Appli_Local As String
Public paramCHQ_SCAN_Appli_Archive As String
Public paramCHQ_SCAN_Save As String

Public paramCHQ_SCAN_Local_Folder As String
Public paramCHQ_SCAN_Archive_Folder As String

Public paramATHIC_MDB As String
Public paramATHIC_Images As String


'__________________________________________________________
Type typeCHQ_SCAN_Stat
 
      Devise        As String
      Date          As String
      Nature        As String
      REM_nb        As Long
      CHQ_nb        As Long
      CHQ_mt        As Currency
End Type

Public Sub srvCHQ_SCAN_SendMail(lCHQ_SCAN As typeCHQ_SCAN, lRemise_text As String, lIntitulé As String)
Dim wSendMail As typeSendMail
Dim bgColor As String
Dim wPath As String
Dim xMontant As String

wSendMail.FromDisplayName = "CHQ_SCAN"
wSendMail.RecipientDisplayName = "DEON"

bgColor = "YELLOW"
xMontant = Format$(CCur(lCHQ_SCAN.Zone1) / 100, "### ### ### ###.00")
If lCHQ_SCAN.Id = "R" Then
    wSendMail.Subject = "Remise totale : " & xMontant & lCHQ_SCAN.Devise
    wSendMail.Attachment = ""
Else
    wSendMail.Subject = "Détection d'un chèque de " & xMontant
    wPath = paramCHQ_SCAN_Image_Archive & "\" & lCHQ_SCAN.Date & "\Archive\"
    wSendMail.Attachment = wPath & lCHQ_SCAN.IMAGE & ".jpg" & ";" & wPath & "ba" & lCHQ_SCAN.IMAGE & ".jpg"
End If
wSendMail.Message = "<body bgcolor=" & Asc34 & bgColor & Asc34 & ">" _
                    & "<FONT face=" & Asc34 & prtFontName_Comic & Asc34 & ">" _
                    & htmlFontColor("MAGENTA") & "<CENTER>" & wSendMail.Subject _
                    & htmlFontColor("BLUE") & "<BR><BR>" & lIntitulé _
                    & "<BR> <BR>" & lRemise_text
 
                        

wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail

End Sub

Public Sub srvCHQ_SCAN_SendMail_Stat(lRemise_text As String, lIntitulé As String)
Dim wSendMail As typeSendMail
Dim bgColor As String
Dim wPath As String
Dim xMontant As String

wSendMail.FromDisplayName = "CHQ_SCAN"
wSendMail.RecipientDisplayName = "STAT"

bgColor = "CYAN"
wSendMail.Subject = "Activité numérisation des chèques du " & dateImp10(YBIATAB0_DATE_CPT_J)
wSendMail.Attachment = ""
wSendMail.Message = "<body bgcolor=" & Asc34 & bgColor & Asc34 & ">" _
                    & "<FONT face=" & Asc34 & prtFontName_Comic & Asc34 & ">" _
                    & htmlFontColor("MAGENTA") & "<CENTER>" & lIntitulé _
                    & htmlFontColor("BLUE") & "<BR><BR>" & lRemise_text

wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail

End Sub


Public Function srvCHQ_SCAN_param()

srvCHQ_SCAN_param = Null

paramCHQ_SCAN_Image_Folder = "MyVision"
paramCHQ_SCAN_Local_Folder = "DreamCheques"
paramCHQ_SCAN_Archive_Folder = "DreamSearch"

If paramEnvironnement = constProduction Then
    paramCHQ_SCAN_Image_Local = "C:\"
    paramCHQ_SCAN_Image_Archive = "\\BIADOCSRV\.dreamcheques$\"
    
    paramCHQ_SCAN_Appli_Local = "C:\"
    paramCHQ_SCAN_Appli_Archive = "\\BIADOCSRV\.dreamcheques$\"
    
    paramCHQ_SCAN_Save = "\\CADSRV\.BIA_SAVE$\CHQ_SCAN\"
Else
    paramCHQ_SCAN_Image_Local = "C:\Temp\CHQ_SCAN\Local\"
    paramCHQ_SCAN_Image_Archive = "C:\Temp\CHQ_SCAN\Archive\"

    paramCHQ_SCAN_Appli_Local = "C:\Temp\CHQ_SCAN\Local\"
    paramCHQ_SCAN_Appli_Archive = "C:\Temp\CHQ_SCAN\Archive\"
   
    paramCHQ_SCAN_Save = "C:\Temp\CHQ_SCAN\Save\"
End If

paramCHQ_SCAN_Image_Local = paramCHQ_SCAN_Image_Local & paramCHQ_SCAN_Image_Folder
paramCHQ_SCAN_Image_Archive = paramCHQ_SCAN_Image_Archive & paramCHQ_SCAN_Image_Folder

paramCHQ_SCAN_Appli_Local = paramCHQ_SCAN_Appli_Local & paramCHQ_SCAN_Local_Folder
paramCHQ_SCAN_Appli_Archive = paramCHQ_SCAN_Appli_Archive & paramCHQ_SCAN_Archive_Folder

End Function


Public Function srvCHQ_SCAN_GetBuffer_ODBC(rsAdo As ADODB.Recordset, lCHQ_SCAN As typeCHQ_SCAN)
On Error GoTo Error_Handler
srvCHQ_SCAN_GetBuffer_ODBC = Null
'Dim l As Long
'X = rsADO("s_GUID")
'Debug.Print X, rsADO("s_Generation"), rsADO("s_Lineage")
lCHQ_SCAN.Id = rsAdo("ID")
lCHQ_SCAN.Cmc7 = rsAdo("Cmc7")
lCHQ_SCAN.Zone4 = rsAdo("Zone4")
lCHQ_SCAN.Zone3 = rsAdo("Zone3")
lCHQ_SCAN.Zone2 = rsAdo("Zone2")
lCHQ_SCAN.Zone1 = rsAdo("Zone1")
lCHQ_SCAN.PATH = rsAdo("PATH")
lCHQ_SCAN.IMAGE = rsAdo("IMAGE")
lCHQ_SCAN.Date = rsAdo("Date")
lCHQ_SCAN.COMPTE = rsAdo("COMPTE")
lCHQ_SCAN.NumLot = rsAdo("NumLot")
lCHQ_SCAN.CRem = rsAdo("CRem")
lCHQ_SCAN.Zone24 = rsAdo("Zone24")
lCHQ_SCAN.DateHourScan = rsAdo("DateHourScan")
lCHQ_SCAN.DateHourSaisie = rsAdo("DateHourSaisie")
lCHQ_SCAN.StatutRem = rsAdo("StatutRem")
lCHQ_SCAN.MotifNonAJ = rsAdo("MotifNonAJ")
lCHQ_SCAN.Saisie = rsAdo("Saisie")
lCHQ_SCAN.RefClient = rsAdo("RefClient")
lCHQ_SCAN.RefInterne = rsAdo("RefInterne")
lCHQ_SCAN.Nature = rsAdo("Nature")
lCHQ_SCAN.Devise = rsAdo("Devise")
 lCHQ_SCAN.Adresse0 = rsAdo("Adresse0")
 lCHQ_SCAN.Adresse1 = rsAdo("Adresse1")
 lCHQ_SCAN.Adresse2 = rsAdo("Adresse2")
 lCHQ_SCAN.Adresse3 = rsAdo("Adresse3")
 lCHQ_SCAN.Adresse4 = rsAdo("Adresse4")
 lCHQ_SCAN.Adresse5 = rsAdo("Adresse5")
 lCHQ_SCAN.RLMC = rsAdo("RLMC")
lCHQ_SCAN.CodeDevise = rsAdo("CodeDevise")


Exit Function
Error_Handler:
srvCHQ_SCAN_GetBuffer_ODBC = Error


End Function

Public Function srvCHQ_SCAN_Init(lCHQ_SCAN As typeCHQ_SCAN)
lCHQ_SCAN.Id = ""
lCHQ_SCAN.Cmc7 = ""
lCHQ_SCAN.Zone4 = ""
lCHQ_SCAN.Zone3 = ""
lCHQ_SCAN.Zone2 = ""
lCHQ_SCAN.Zone1 = ""
lCHQ_SCAN.PATH = ""
lCHQ_SCAN.IMAGE = ""
lCHQ_SCAN.Date = ""
lCHQ_SCAN.COMPTE = ""
lCHQ_SCAN.NumLot = ""
lCHQ_SCAN.CRem = ""
lCHQ_SCAN.Zone24 = ""
lCHQ_SCAN.DateHourScan = ""
lCHQ_SCAN.StatutRem = ""
lCHQ_SCAN.MotifNonAJ = ""
lCHQ_SCAN.Saisie = ""
lCHQ_SCAN.RefClient = ""
 lCHQ_SCAN.RefClient = ""
lCHQ_SCAN.RefInterne = ""
lCHQ_SCAN.Nature = ""
lCHQ_SCAN.Devise = ""
 lCHQ_SCAN.Adresse0 = ""
 lCHQ_SCAN.Adresse1 = ""
 lCHQ_SCAN.Adresse2 = ""
 lCHQ_SCAN.Adresse3 = ""
 lCHQ_SCAN.Adresse4 = ""
 lCHQ_SCAN.Adresse5 = ""
 lCHQ_SCAN.RLMC = ""

End Function

Public Function sqlCHQ_SCAN_Update(newCHQ_SCAN As typeCHQ_SCAN, oldCHQ_SCAN As typeCHQ_SCAN, rsAdo As ADODB.Recordset, cnAdo As ADODB.Connection)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean
Dim K As String

On Error GoTo Error_Handler
sqlCHQ_SCAN_Update = Null
' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldCHQ_SCAN.IMAGE <> newCHQ_SCAN.IMAGE Then
    sqlCHQ_SCAN_Update = "Erreur Clé (new/old) de l'IMAGE : " & newCHQ_SCAN.IMAGE & " / " & oldCHQ_SCAN.IMAGE
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
'===================================================================================

xWhere = " where IMAGE = '" & oldCHQ_SCAN.IMAGE & "' and Date = '" & oldCHQ_SCAN.Date & "'"

xSet = " set"   '''''''''''RefInterne = '" & newCHQ_SCAN.RefInterne & "'"
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newCHQ_SCAN.COMPTE <> oldCHQ_SCAN.COMPTE Then blnUpdate = True: xSet = xSet & " , COMPTE = '" & newCHQ_SCAN.COMPTE & "'"
If newCHQ_SCAN.RefClient <> oldCHQ_SCAN.RefClient Then blnUpdate = True: xSet = xSet & " , RefClient = '" & newCHQ_SCAN.RefClient & "'"
If newCHQ_SCAN.RefInterne <> oldCHQ_SCAN.RefInterne Then blnUpdate = True: xSet = xSet & " , RefInterne = '" & newCHQ_SCAN.RefInterne & "'"
If newCHQ_SCAN.Nature <> oldCHQ_SCAN.Nature Then blnUpdate = True: xSet = xSet & " , Nature = '" & newCHQ_SCAN.Nature & "'"
If newCHQ_SCAN.Devise <> oldCHQ_SCAN.Devise Then blnUpdate = True: xSet = xSet & " , Devise = '" & newCHQ_SCAN.Devise & "'"
If newCHQ_SCAN.StatutRem <> oldCHQ_SCAN.StatutRem Then blnUpdate = True: xSet = xSet & " , StatutRem  = '" & newCHQ_SCAN.StatutRem & "'"
If newCHQ_SCAN.Zone1 <> oldCHQ_SCAN.Zone1 Then blnUpdate = True: xSet = xSet & " , Zone1  = '" & newCHQ_SCAN.Zone1 & "'"
If newCHQ_SCAN.CRem <> oldCHQ_SCAN.CRem Then blnUpdate = True: xSet = xSet & " , CRem  = '" & newCHQ_SCAN.CRem & "'"
    
'Si modification , supprimer la première virgule
If Len(xSet) > 5 Then
    K = InStr(5, xSet, ","): If K > 0 Then Mid$(xSet, K, 1) = " "
    xSQL = "update CHEQUE" & xSet & xWhere
    Call FEU_ROUGE
    Set rsAdo = cnAdo.Execute(xSQL, Nb)
    Call FEU_VERT
   
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
       sqlCHQ_SCAN_Update = "Erreur màj, IMAGE = " & newCHQ_SCAN.IMAGE
        Exit Function
    End If
End If
Exit Function
'===================================================================================
Error_Handler:
    sqlCHQ_SCAN_Update = Error
End Function

Public Function sqlCHQ_SCAN_Insert(newY As typeCHQ_SCAN, cnAdo As ADODB.Connection)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlCHQ_SCAN_Insert = Null

xSet = " (ID"
xValues = " values('" & newY.Id & "'"

' Détecter les modifications
'===================================================================================
If Trim(newY.Cmc7) <> "" Then xSet = xSet & ",Cmc7": xValues = xValues & " ,'" & newY.Cmc7 & "'"
If Trim(newY.Zone4) <> "" Then xSet = xSet & ",Zone4": xValues = xValues & " ,'" & newY.Zone4 & "'"
If Trim(newY.Zone3) <> "" Then xSet = xSet & ",Zone3": xValues = xValues & " ,'" & newY.Zone3 & "'"
If Trim(newY.Zone2) <> "" Then xSet = xSet & ",Zone2": xValues = xValues & " ,'" & newY.Zone2 & "'"
If Trim(newY.Zone1) <> "" Then xSet = xSet & ",Zone1": xValues = xValues & " ,'" & newY.Zone1 & "'"
If Trim(newY.PATH) <> "" Then xSet = xSet & ",PATH": xValues = xValues & " ,'" & newY.PATH & "'"
If Trim(newY.IMAGE) <> "" Then xSet = xSet & ",IMAGE": xValues = xValues & " ,'" & newY.IMAGE & "'"
If Trim(newY.Date) <> "" Then xSet = xSet & ", [Date]": xValues = xValues & " ,'" & newY.Date & "'"
If Trim(newY.COMPTE) <> "" Then xSet = xSet & ",COMPTE": xValues = xValues & " ,'" & newY.COMPTE & "'"
If Trim(newY.NumLot) <> "" Then xSet = xSet & ",NumLot": xValues = xValues & " ,'" & newY.NumLot & "'"
If Trim(newY.CRem) <> "" Then xSet = xSet & ",CRem": xValues = xValues & " ,'" & newY.CRem & "'"
If Trim(newY.Zone24) <> "" Then xSet = xSet & ",Zone24": xValues = xValues & " ,'" & newY.Zone24 & "'"
If Trim(newY.DateHourScan) <> "" Then xSet = xSet & ",DateHourScan": xValues = xValues & " ,'" & newY.DateHourScan & "'"
If Trim(newY.DateHourSaisie) <> "" Then xSet = xSet & ",DateHourSaisie": xValues = xValues & " ,'" & newY.DateHourSaisie & "'"
If Trim(newY.StatutRem) <> "" Then xSet = xSet & ",StatutRem": xValues = xValues & " ,'" & newY.StatutRem & "'"
If Trim(newY.MotifNonAJ) <> "" Then xSet = xSet & ",MotifNonAJ": xValues = xValues & " ,'" & newY.MotifNonAJ & "'"
If Trim(newY.Saisie) <> "" Then xSet = xSet & ",Saisie": xValues = xValues & " ,'" & newY.Saisie & "'"
If Trim(newY.RefClient) <> "" Then xSet = xSet & ",RefClient": xValues = xValues & " ,'" & newY.RefClient & "'"
If Trim(newY.RefInterne) <> "" Then xSet = xSet & ",RefInterne": xValues = xValues & " ,'" & newY.RefInterne & "'"
If Trim(newY.Nature) <> "" Then xSet = xSet & ",Nature": xValues = xValues & " ,'" & newY.Nature & "'"
If Trim(newY.Devise) <> "" Then xSet = xSet & ",Devise": xValues = xValues & " ,'" & newY.Devise & "'"

 If Trim(newY.Adresse0) <> "" Then xSet = xSet & ",Adresse0": xValues = xValues & " ,'" & newY.Adresse0 & "'"
 If Trim(newY.Adresse1) <> "" Then xSet = xSet & ",Adresse1": xValues = xValues & " ,'" & newY.Adresse1 & "'"
 If Trim(newY.Adresse2) <> "" Then xSet = xSet & ",Adresse2": xValues = xValues & " ,'" & newY.Adresse2 & "'"
 If Trim(newY.Adresse3) <> "" Then xSet = xSet & ",Adresse3": xValues = xValues & " ,'" & newY.Adresse3 & "'"
 If Trim(newY.Adresse4) <> "" Then xSet = xSet & ",Adresse4": xValues = xValues & " ,'" & newY.Adresse4 & "'"
 If Trim(newY.Adresse5) <> "" Then xSet = xSet & ",Adresse5": xValues = xValues & " ,'" & newY.Adresse5 & "'"
 If Trim(newY.RLMC) <> "" Then xSet = xSet & ",RLMC": xValues = xValues & " ,'" & newY.RLMC & "'"
Call FEU_ROUGE
xSQL = "Insert into CHEQUE " & xSet & ")" & xValues & ")"

Set rsAdo = cnAdo.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlCHQ_SCAN_Insert = "Erreur màj : " & newY.IMAGE
    Exit Function
End If
 
Exit Function
Error_Handler:
MsgBox Error, vbCritical, sqlCHQ_SCAN_Insert
    sqlCHQ_SCAN_Insert = Error
End Function

Public Function sqlCHQ_SCAN_Delete(oldCHQ_SCAN As typeCHQ_SCAN, rsAdo As ADODB.Recordset, cnAdo As ADODB.Connection)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String

On Error GoTo Error_Handler
sqlCHQ_SCAN_Delete = Null

xWhere = " where CRem = '" & oldCHQ_SCAN.CRem & "' and Date = '" & oldCHQ_SCAN.Date & "'"

xSQL = "delete * from CHEQUE" & xWhere
    Call FEU_ROUGE
    Set rsAdo = cnAdo.Execute(xSQL, Nb)
    Call FEU_VERT
    
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
       sqlCHQ_SCAN_Delete = "Erreur màj, CRem  = " & oldCHQ_SCAN.CRem
        Exit Function
    End If
    
Exit Function
'===================================================================================
Error_Handler:
    sqlCHQ_SCAN_Delete = Error
End Function


Public Sub srvATHIC_Param()
Dim X As String

Call sqlYBIATAB0_Read("BIA_ATHIC", "MDB", "", X)  ' "C:\Temp\ATHIC\"
paramATHIC_MDB = Trim(Mid$(X, 1, 99))

Call sqlYBIATAB0_Read("BIA_ATHIC", "Server", "", X)  ' "\\appsrv2011\"
paramATHIC_Images = Trim(Mid$(X, 1, 99))
End Sub
