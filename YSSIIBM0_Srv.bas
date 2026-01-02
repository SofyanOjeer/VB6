Attribute VB_Name = "srvYSSIIBM0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYSSIIBM0
    SSIIBMNAT    As String
    SSIIBMUIDD   As Long
    SSIIBMPRFK   As String
    SSIIBMTLNK  As Long
    SSIIBMYFCT  As String
    SSIIBMYUSR  As String
    SSIIBMYAMJ  As Long
    SSIIBMYHMS  As Long
    SSIIBMYVER   As Long
    
    UPUPRF       As String
    UPUSCL       As String
    UPPWCD       As Long
    UPPWEI       As Long
    UPPWEX       As String
    UPPWON       As String
    UPSPAU       As String
    UPINPG       As String
    UPINPL       As String
    UPJBDS       As String
    UPJBDL       As String
    UPGRPF       As String
    UPGRAU       As String
    UPTEXT       As String
    UPSPEN       As String
    UPCRLB       As String
    UPINMN       As String
    UPINML       As String
    UPLTCP       As String
    UPATPG       As String
    UPATPL       As String
    UPSTAT       As String
    UPUID        As String
    UPCRTD       As Long
    UPCHGD       As Long
    UPPSOD       As Long
End Type
Public xYSSIIBM0 As typeYSSIIBM0
Public Function sqlYSSIIBM0_Update(newY As typeYSSIIBM0, oldY As typeYSSIIBM0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYSSIIBM0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.SSIIBMUIDD <> newY.SSIIBMUIDD Then
    sqlYSSIIBM0_Update = "Erreur SSIIBMUIDD : " & newY.SSIIBMUIDD & " / " & oldY.SSIIBMUIDD
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SSIIBMNAT = '" & oldY.SSIIBMNAT & "'" _
       & " and SSIIBMUIDD = " & oldY.SSIIBMUIDD _
       & " and SSIIBMYVER = " & oldY.SSIIBMYVER
       
newY.SSIIBMYVER = newY.SSIIBMYVER + 1
xSet = xSet & " set SSIIBMYVER = " & newY.SSIIBMYVER
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.SSIIBMUIDD <> oldY.SSIIBMUIDD Then blnUpdate = True: xSet = xSet & " , SSIIBMUIDD = " & newY.SSIIBMUIDD
If newY.SSIIBMTLNK <> oldY.SSIIBMTLNK Then blnUpdate = True: xSet = xSet & " , SSIIBMTLNK = " & newY.SSIIBMTLNK
If newY.SSIIBMYAMJ <> oldY.SSIIBMYAMJ Then blnUpdate = True: xSet = xSet & " , SSIIBMYAMJ = " & newY.SSIIBMYAMJ
If newY.SSIIBMYHMS <> oldY.SSIIBMYHMS Then blnUpdate = True: xSet = xSet & " , SSIIBMYHMS = " & newY.SSIIBMYHMS

If newY.SSIIBMPRFK <> oldY.SSIIBMPRFK Then blnUpdate = True:  xSet = xSet & " , SSIIBMPRFK= '" & newY.SSIIBMPRFK & "'"
If newY.SSIIBMYUSR <> oldY.SSIIBMYUSR Then blnUpdate = True:  xSet = xSet & " , SSIIBMYUSR = '" & Replace(Trim(newY.SSIIBMYUSR), "'", "''") & "'"
If newY.SSIIBMYFCT <> oldY.SSIIBMYFCT Then blnUpdate = True:  xSet = xSet & " , SSIIBMYFCT = '" & Replace(Trim(newY.SSIIBMYFCT), "'", "''") & "'"

If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE & ".YSSIIBM0" & xSet & xWhere
    
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYSSIIBM0_Update = "Erreur màj : " & newY.SSIIBMUIDD
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYSSIIBM0_Update = "sqlYSSIIBM0_Update  " & vbCrLf & Error
End Function

Public Function sqlYSSIIBMH_Update(newY As typeYSSIIBM0, oldY As typeYSSIIBM0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYSSIIBMH_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.SSIIBMUIDD <> newY.SSIIBMUIDD Then
    sqlYSSIIBMH_Update = "Erreur SSIIBMUIDD : " & newY.SSIIBMUIDD & " / " & oldY.SSIIBMUIDD
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SSIIBMNAT = '" & oldY.SSIIBMNAT & "'" _
       & " and SSIIBMUIDD = " & oldY.SSIIBMUIDD _
       & " and SSIIBMYVER = " & oldY.SSIIBMYVER
       
newY.SSIIBMYVER = newY.SSIIBMYVER ''''''''''''''''''''+ 1
xSet = xSet & " set SSIIBMYVER = " & newY.SSIIBMYVER
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.SSIIBMUIDD <> oldY.SSIIBMUIDD Then blnUpdate = True: xSet = xSet & " , SSIIBMUIDD = " & newY.SSIIBMUIDD
If newY.SSIIBMTLNK <> oldY.SSIIBMTLNK Then blnUpdate = True: xSet = xSet & " , SSIIBMTLNK = " & newY.SSIIBMTLNK
If newY.SSIIBMYAMJ <> oldY.SSIIBMYAMJ Then blnUpdate = True: xSet = xSet & " , SSIIBMYAMJ = " & newY.SSIIBMYAMJ
If newY.SSIIBMYHMS <> oldY.SSIIBMYHMS Then blnUpdate = True: xSet = xSet & " , SSIIBMYHMS = " & newY.SSIIBMYHMS

If newY.SSIIBMPRFK <> oldY.SSIIBMPRFK Then blnUpdate = True:  xSet = xSet & " , SSIIBMPRFK= '" & newY.SSIIBMPRFK & "'"
If newY.SSIIBMYUSR <> oldY.SSIIBMYUSR Then blnUpdate = True:  xSet = xSet & " , SSIIBMYUSR = '" & Replace(Trim(newY.SSIIBMYUSR), "'", "''") & "'"
If newY.SSIIBMYFCT <> oldY.SSIIBMYFCT Then blnUpdate = True:  xSet = xSet & " , SSIIBMYFCT = '" & Replace(Trim(newY.SSIIBMYFCT), "'", "''") & "'"

If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE & ".YSSIIBMH" & xSet & xWhere
    
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYSSIIBMH_Update = "Erreur màj : " & newY.SSIIBMUIDD
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYSSIIBMH_Update = "sqlYSSIIBMH_Update " & vbCrLf & Error
End Function


Public Function sqlYSSIIBM0_Profil_Insert(newY As typeYSSIIBM0, oldY As typeYSSIIBM0)
Dim X As String, xSQL As String, Nb As Long

On Error GoTo Error_Handler
sqlYSSIIBM0_Profil_Insert = Null

'===================================================================================
xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YSSIIBM0_W" _
     & " select * from " & paramIBM_Library_SABSPE & ".YSSIIBM0" _
     & " where SSIIBMNAT = ' ' and SSIIBMUIDD = " & oldY.SSIIBMUIDD
Call FEU_ROUGE
Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
If Nb = 0 Then
    sqlYSSIIBM0_Profil_Insert = xSQL
    Exit Function
End If
'===================================================================================
xSQL = "update " & paramIBM_Library_SABSPE & ".YSSIIBM0_W" _
     & " set SSIIBMNAT= '$' , SSIIBMUIDD = " & newY.SSIIBMUIDD & " , UPUPRF = '" & Trim(newY.UPUPRF) & "'" _
     & " , SSIIBMTLNK = " & newY.SSIIBMTLNK & " ,SSIIBMPRFK = '" & Trim(newY.SSIIBMPRFK) & "'" _
     & " , SSIIBMYUSR = '" & Trim(newY.SSIIBMYUSR) & "', SSIIBMYAMJ = " & newY.SSIIBMYAMJ _
     & ", SSIIBMYHMS = " & newY.SSIIBMYHMS & ", SSIIBMYVER= " & newY.SSIIBMYVER _
     & " , UPMGQU = '' , UPTEXT = '" & Replace(Trim(newY.UPTEXT), "'", "''") & "'" _
     & " , UPPSOD = 0 , UPCRTD = 0 , UPCHGD = 0"
Call FEU_ROUGE
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
If Nb = 0 Then
        sqlYSSIIBM0_Profil_Insert = xSQL
        Exit Function
    End If
'===================================================================================
xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YSSIIBM0" _
     & " select * from " & paramIBM_Library_SABSPE & ".YSSIIBM0_W" _
     & " where SSIIBMNAT = '$' and SSIIBMUIDD = " & newY.SSIIBMUIDD
Call FEU_ROUGE
Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
If Nb = 0 Then
    sqlYSSIIBM0_Profil_Insert = xSQL
    Exit Function
End If
 '===================================================================================
xSQL = "delete from " & paramIBM_Library_SABSPE & ".YSSIIBM0_W" _
     & " where SSIIBMNAT = '$' and SSIIBMUIDD = " & newY.SSIIBMUIDD
Call FEU_ROUGE
Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
If Nb = 0 Then
    sqlYSSIIBM0_Profil_Insert = xSQL
    Exit Function
End If
   

Exit Function
Error_Handler:
    sqlYSSIIBM0_Profil_Insert = "sqlYSSIIBM0_Profil_Insert " & vbCrLf & Error
End Function


Public Function rsYSSIIBM0_GetBuffer(rsAdo As ADODB.Recordset, lYSSIIBM0 As typeYSSIIBM0)
Dim wAMJ As Long
On Error GoTo Error_Handler
rsYSSIIBM0_GetBuffer = Null
lYSSIIBM0.SSIIBMNAT = rsAdo("SSIIBMNAT")
lYSSIIBM0.SSIIBMUIDD = rsAdo("SSIIBMUIDD")
lYSSIIBM0.SSIIBMPRFK = rsAdo("SSIIBMPRFK")
lYSSIIBM0.SSIIBMTLNK = rsAdo("SSIIBMTLNK")
lYSSIIBM0.SSIIBMYFCT = Trim(rsAdo("SSIIBMYFCT"))
lYSSIIBM0.SSIIBMYUSR = Trim(rsAdo("SSIIBMYUSR"))
lYSSIIBM0.SSIIBMYAMJ = rsAdo("SSIIBMYAMJ")

lYSSIIBM0.SSIIBMYHMS = rsAdo("SSIIBMYHMS")
lYSSIIBM0.SSIIBMYVER = rsAdo("SSIIBMYVER")
lYSSIIBM0.UPUPRF = Trim(rsAdo("UPUPRF"))
lYSSIIBM0.UPUSCL = Trim(rsAdo("UPUSCL"))
lYSSIIBM0.UPPWEI = rsAdo("UPPWEI")
lYSSIIBM0.UPPWEX = Trim(rsAdo("UPPWEX"))
lYSSIIBM0.UPPWON = Trim(rsAdo("UPPWON"))
lYSSIIBM0.UPSPAU = Trim(rsAdo("UPSPAU"))
lYSSIIBM0.UPINPG = Trim(rsAdo("UPINPG"))
lYSSIIBM0.UPINPL = Trim(rsAdo("UPINPL"))
lYSSIIBM0.UPJBDS = Trim(rsAdo("UPJBDS"))
lYSSIIBM0.UPJBDL = Trim(rsAdo("UPJBDL"))
lYSSIIBM0.UPGRPF = Trim(rsAdo("UPGRPF"))
lYSSIIBM0.UPGRAU = Trim(rsAdo("UPGRAU"))
lYSSIIBM0.UPTEXT = Trim(rsAdo("UPTEXT"))
lYSSIIBM0.UPSPEN = Trim(rsAdo("UPSPEN"))
lYSSIIBM0.UPCRLB = Trim(rsAdo("UPCRLB"))
lYSSIIBM0.UPINMN = Trim(rsAdo("UPINMN"))
lYSSIIBM0.UPINML = Trim(rsAdo("UPINML"))
lYSSIIBM0.UPLTCP = Trim(rsAdo("UPLTCP"))
lYSSIIBM0.UPATPG = Trim(rsAdo("UPATPG"))
lYSSIIBM0.UPATPL = Trim(rsAdo("UPATPL"))
lYSSIIBM0.UPSTAT = Trim(rsAdo("UPSTAT"))
lYSSIIBM0.UPUID = Trim(rsAdo("UPUID"))
If lYSSIIBM0.SSIIBMNAT = "$" Then
    lYSSIIBM0.UPCRTD = 0
    lYSSIIBM0.UPCHGD = 0
    lYSSIIBM0.UPPSOD = 0
    lYSSIIBM0.UPPWCD = 0
Else
    lYSSIIBM0.UPCRTD = Val(rsAdo("UPCRTC") & rsAdo("UPCRTD")) + 19000000
    lYSSIIBM0.UPCHGD = Val(rsAdo("UPCHGC") & rsAdo("UPCHGD")) + 19000000
    lYSSIIBM0.UPPSOD = Val(rsAdo("UPPSOC") & rsAdo("UPPSOD")) + 19000000
    lYSSIIBM0.UPPWCD = Val(rsAdo("UPPWCC") & rsAdo("UPPWCD")) + 19000000
End If
Exit Function
Error_Handler:
rsYSSIIBM0_GetBuffer = Error


End Function

Public Function rsYSSIIBM0_Init(lYSSIIBM0 As typeYSSIIBM0)
lYSSIIBM0.SSIIBMYVER = 0
lYSSIIBM0.SSIIBMNAT = ""
lYSSIIBM0.SSIIBMUIDD = 0
lYSSIIBM0.SSIIBMTLNK = 0
lYSSIIBM0.SSIIBMYAMJ = 0
lYSSIIBM0.SSIIBMYFCT = ""
lYSSIIBM0.SSIIBMYUSR = ""
lYSSIIBM0.SSIIBMYHMS = 0
lYSSIIBM0.SSIIBMPRFK = ""

lYSSIIBM0.UPUPRF = ""
lYSSIIBM0.UPUSCL = ""
lYSSIIBM0.UPPWEI = 0
lYSSIIBM0.UPPWEX = ""
lYSSIIBM0.UPPWON = ""
lYSSIIBM0.UPSPAU = ""
lYSSIIBM0.UPINPG = ""
lYSSIIBM0.UPINPL = ""
lYSSIIBM0.UPJBDS = ""
lYSSIIBM0.UPJBDL = ""
lYSSIIBM0.UPGRPF = ""
lYSSIIBM0.UPGRAU = ""
lYSSIIBM0.UPTEXT = ""
lYSSIIBM0.UPSPEN = ""
lYSSIIBM0.UPCRLB = ""
lYSSIIBM0.UPINMN = ""
lYSSIIBM0.UPINML = ""
lYSSIIBM0.UPLTCP = ""
lYSSIIBM0.UPATPG = ""
lYSSIIBM0.UPATPL = ""
lYSSIIBM0.UPSTAT = ""
lYSSIIBM0.UPUID = ""

lYSSIIBM0.UPCRTD = 0
lYSSIIBM0.UPCHGD = 0
lYSSIIBM0.UPPSOD = 0
lYSSIIBM0.UPPWCD = 0

End Function









