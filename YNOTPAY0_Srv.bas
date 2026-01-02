Attribute VB_Name = "srvYNOTPAY0"
'---------------------------------------------------------
Option Explicit
Type typeYNOTPAY0

    NOTPAYISO   As String * 2  ' code ISO pays
    NOTPAYSEQ   As Long        ' N° séquence (info)
    
    NOTPAYHAMJ  As Long        ' DATE d'arrêté
    NOTPAYPROV  As String * 1  ' Provisionable = 'P'
    NOTPAYCOFA  As String * 2  ' notation coface
    NOTPAYCOFK  As String * 1  ' notation coface Auto / Manuel
    NOTPAYCOFD  As Long        ' DATE maj
    NOTPAYCOF2  As String * 2  ' notation coface
    NOTPAYOCDE  As String * 1  ' notation OCDE
    NOTPAYOCDK  As String * 1  ' notation OCDE Auto / Manuel
    NOTPAYOCDD  As Long        ' DATE maj
    NOTPAYSP    As String * 4  ' notation S & P
    NOTPAYSPK   As String * 1  ' notation S & P Auto / Manuel
    NOTPAYSPD   As Long        ' DATE maj
    NOTPAYCEG   As Long        ' critère événement grave
    NOTPAYBIAN  As String * 3  ' notation BIA
    NOTPAYBIAK  As String * 1  ' notation BIA Auto / Manuel
    NOTPAYBIAD  As Long        ' DATE maj
    NOTPAYTAUX  As Double      ' taux BIA
    NOTPAYFISC  As String * 2  ' taux fisc
    NOTPAYTXT   As String * 32 ' commentaire
    NOTPAYXAMJ  As Long        ' DATE maj
    NOTPAYXHMS  As Long        ' heure maj
    NOTPAYXUSR  As String * 10 ' utilisateur maj
    
'____________________________________________________ Journalisation
    JORCV                   As Long
    JOSEQN                  As Long
    JRNBIATRN               As Long
'____________________________________________________ Journalisation
    NOTPAYLIB   As String
    
End Type

Public ddsYNOTPAY0(24) As String * 50
Public NOTPAYBIAN_Row As Integer, NOTPAYTAUX_Row As Integer
Public Function sqlYNOTPAY0_Delete(oldY As typeYNOTPAY0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlYNOTPAY0_Delete = Null


xWhere = " where NOTPAYISO = '" & oldY.NOTPAYISO & "'" _
       & " and   NOTPAYSEQ = " & oldY.NOTPAYSEQ _


' Suppression physique
'===================================================================================

xSQL = "Delete from " & paramIBM_Library_SABSPE & ".YNOTPAY0" & xWhere
Call FEU_ROUGE
Set rsSab = cnsab.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYNOTPAY0_Delete = "Erreur SUP : " & xWhere
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYNOTPAY0_Delete = Error
End Function


Public Function sqlYNOTPAY0_Update_NOTPAYSEQ(oldY As typeYNOTPAY0, lNOTPAYSEQ As Long)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlYNOTPAY0_Update_NOTPAYSEQ = Null


xWhere = " set NOTPAYSEQ = " & lNOTPAYSEQ _
       & " where NOTPAYISO = '" & oldY.NOTPAYISO & "'" _
       & " and   NOTPAYSEQ = " & oldY.NOTPAYSEQ _


' Suppression physique
'===================================================================================

xSQL = "Update " & paramIBM_Library_SABSPE & ".YNOTPAY0" & xWhere
Call FEU_ROUGE
Set rsSab = cnsab.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYNOTPAY0_Update_NOTPAYSEQ = "Erreur Update_NOTPAYSEQ : " & xWhere
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYNOTPAY0_Update_NOTPAYSEQ = Error
End Function

Public Function sqlYNOTPAY0_Delete_Where(lWhere As String)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String

On Error GoTo Error_Handler
sqlYNOTPAY0_Delete_Where = Null

' Suppression physique
'===================================================================================

xSQL = "Delete from " & paramIBM_Library_SABSPE & ".YNOTPAY0" & lWhere
Call FEU_ROUGE
Set rsSab = cnsab.Execute(xSQL, Nb)
Call FEU_VERT
 
Exit Function
Error_Handler:
    sqlYNOTPAY0_Delete_Where = Error
End Function

Public Function sqlYNOTPAY0_Insert(newY As typeYNOTPAY0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYNOTPAY0_Insert = Null
newY.NOTPAYXAMJ = DSys                ' DATE maj
newY.NOTPAYXHMS = time_Hms            ' heure maj
newY.NOTPAYXUSR = usrName_UCase       ' utilisateur maj

xSet = " (NOTPAYISO,NOTPAYSEQ"
xValues = " values('" & Trim(newY.NOTPAYISO) & "', " & newY.NOTPAYSEQ

' Insertion :
'===================================================================================
If newY.NOTPAYHAMJ <> 0 Then xSet = xSet & ",NOTPAYHAMJ": xValues = xValues & ", " & newY.NOTPAYHAMJ
If newY.NOTPAYCOFD <> 0 Then xSet = xSet & ",NOTPAYCOFD": xValues = xValues & ", " & newY.NOTPAYCOFD
If newY.NOTPAYOCDD <> 0 Then xSet = xSet & ",NOTPAYOCDD": xValues = xValues & ", " & newY.NOTPAYOCDD
If newY.NOTPAYSPD <> 0 Then xSet = xSet & ",NOTPAYSPD": xValues = xValues & ", " & newY.NOTPAYSPD
If newY.NOTPAYBIAD <> 0 Then xSet = xSet & ",NOTPAYBIAD": xValues = xValues & ", " & newY.NOTPAYBIAD
If newY.NOTPAYCEG <> 0 Then xSet = xSet & ",NOTPAYCEG": xValues = xValues & ", " & newY.NOTPAYCEG
If newY.NOTPAYTAUX <> 0 Then xSet = xSet & ",NOTPAYTAUX": xValues = xValues & ", " & Comma_Point(newY.NOTPAYTAUX)
If newY.NOTPAYXAMJ <> 0 Then xSet = xSet & ",NOTPAYXAMJ": xValues = xValues & ", " & newY.NOTPAYXAMJ
If newY.NOTPAYXHMS <> 0 Then xSet = xSet & ",NOTPAYXHMS": xValues = xValues & ", " & newY.NOTPAYXHMS

'===================================================================================

If newY.NOTPAYPROV <> "" Then xSet = xSet & ",NOTPAYPROV": xValues = xValues & ", '" & Replace(newY.NOTPAYPROV, "'", "''") & "'"
If newY.NOTPAYCOFA <> "" Then xSet = xSet & ",NOTPAYCOFA": xValues = xValues & ", '" & Replace(newY.NOTPAYCOFA, "'", "''") & "'"
If newY.NOTPAYCOFK <> "" Then xSet = xSet & ",NOTPAYCOFK": xValues = xValues & ", '" & Replace(newY.NOTPAYCOFK, "'", "''") & "'"
If newY.NOTPAYCOF2 <> "" Then xSet = xSet & ",NOTPAYCOF2": xValues = xValues & ", '" & Replace(newY.NOTPAYCOF2, "'", "''") & "'"
If newY.NOTPAYOCDE <> "" Then xSet = xSet & ",NOTPAYOCDE": xValues = xValues & ", '" & Replace(newY.NOTPAYOCDE, "'", "''") & "'"
If newY.NOTPAYOCDK <> "" Then xSet = xSet & ",NOTPAYOCDK": xValues = xValues & ", '" & Replace(newY.NOTPAYOCDK, "'", "''") & "'"

If newY.NOTPAYSP <> "" Then xSet = xSet & ",NOTPAYSP": xValues = xValues & ", '" & Replace(newY.NOTPAYSP, "'", "''") & "'"
If newY.NOTPAYSPK <> "" Then xSet = xSet & ",NOTPAYSPK": xValues = xValues & ", '" & Replace(newY.NOTPAYSPK, "'", "''") & "'"
If newY.NOTPAYBIAN <> "" Then xSet = xSet & ",NOTPAYBIAN": xValues = xValues & ", '" & Replace(newY.NOTPAYBIAN, "'", "''") & "'"
If newY.NOTPAYBIAK <> "" Then xSet = xSet & ",NOTPAYBIAK": xValues = xValues & ", '" & Replace(newY.NOTPAYBIAK, "'", "''") & "'"
If newY.NOTPAYFISC <> "" Then xSet = xSet & ",NOTPAYFISC": xValues = xValues & ", '" & Replace(newY.NOTPAYFISC, "'", "''") & "'"
If Trim(newY.NOTPAYTXT) <> "" Then xSet = xSet & ",NOTPAYTXT": xValues = xValues & ", '" & Replace(Trim(newY.NOTPAYTXT), "'", "''") & "'"
If newY.NOTPAYXUSR <> "" Then xSet = xSet & ",NOTPAYXUSR": xValues = xValues & ", '" & Replace(newY.NOTPAYXUSR, "'", "''") & "'"
Call FEU_ROUGE
xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YNOTPAY0" & xSet & ")" & xValues & ")"

Set rsSab = cnsab.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYNOTPAY0_Insert = "Erreur màj : " & newY.NOTPAYISO & " " & newY.NOTPAYSEQ
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYNOTPAY0_Insert = Error
End Function
Public Function sqlYNOTPAY0_Read(oldY As typeYNOTPAY0)
Dim X As String, xSQL As String, Nb As Long
Dim V

On Error GoTo Error_Handler
sqlYNOTPAY0_Read = Null

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YNOTPAY0 " _
       & " where NOTPAYISO = '" & oldY.NOTPAYISO & "'" _
       & " and   NOTPAYSEQ = " & oldY.NOTPAYSEQ

Set rsSab = cnsab.Execute(xSQL)

If rsSab.EOF Then
    sqlYNOTPAY0_Read = "? inconnu"
Else
    V = rsYNOTPAY0_GetBuffer(rsSab, oldY)
    If Not IsNull(V) Then sqlYNOTPAY0_Read = "? srvYNOTPAY0_GetBuffer"
End If
 
Exit Function
Error_Handler:
    sqlYNOTPAY0_Read = Error
End Function


Public Function sqlYNOTPAY0_Update(newY As typeYNOTPAY0, oldY As typeYNOTPAY0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean
On Error GoTo Error_Handler
sqlYNOTPAY0_Update = Null
blnUpdate = False


' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.NOTPAYSEQ <> newY.NOTPAYSEQ Then
    sqlYNOTPAY0_Update = "Clé erronnée lors mise à jour !"
    Exit Function
End If

newY.NOTPAYXAMJ = DSys                ' DATE maj
newY.NOTPAYXHMS = time_Hms            ' heure maj
newY.NOTPAYXUSR = usrName_UCase       ' utilisateur maj


'===================================================================================
xWhere = " where NOTPAYISO = '" & oldY.NOTPAYISO & "'" _
       & " and   NOTPAYSEQ = " & oldY.NOTPAYSEQ


xSet = " set"

' Détecter les modifications
'===================================================================================

If newY.NOTPAYHAMJ <> oldY.NOTPAYHAMJ Then blnUpdate = True:  xSet = xSet & " , NOTPAYHAMJ = " & newY.NOTPAYHAMJ
If newY.NOTPAYCOFD <> oldY.NOTPAYCOFD Then blnUpdate = True:  xSet = xSet & " , NOTPAYCOFD = " & newY.NOTPAYCOFD
If newY.NOTPAYOCDD <> oldY.NOTPAYOCDD Then blnUpdate = True:  xSet = xSet & " , NOTPAYOCDD = " & newY.NOTPAYOCDD
If newY.NOTPAYSPD <> oldY.NOTPAYSPD Then blnUpdate = True:  xSet = xSet & " , NOTPAYSPD = " & newY.NOTPAYSPD
If newY.NOTPAYBIAD <> oldY.NOTPAYBIAD Then blnUpdate = True:  xSet = xSet & " , NOTPAYBIAD = " & newY.NOTPAYBIAD
If newY.NOTPAYCEG <> oldY.NOTPAYCEG Then blnUpdate = True:  xSet = xSet & " , NOTPAYCEG = " & newY.NOTPAYCEG
If newY.NOTPAYTAUX <> oldY.NOTPAYTAUX Then blnUpdate = True:  xSet = xSet & " , NOTPAYTAUX= " & Comma_Point(newY.NOTPAYTAUX)
If newY.NOTPAYXAMJ <> oldY.NOTPAYXAMJ Then blnUpdate = True:  xSet = xSet & " , NOTPAYXAMJ = " & newY.NOTPAYXAMJ
If newY.NOTPAYXHMS <> oldY.NOTPAYXHMS Then blnUpdate = True:  xSet = xSet & " , NOTPAYXHMS = " & newY.NOTPAYXHMS

If newY.NOTPAYPROV <> oldY.NOTPAYPROV Then blnUpdate = True: xSet = xSet & ",NOTPAYPROV = '" & Replace(newY.NOTPAYPROV, "'", "''") & "'"
If newY.NOTPAYCOFA <> oldY.NOTPAYCOFA Then blnUpdate = True: xSet = xSet & ",NOTPAYCOFA = '" & Replace(newY.NOTPAYCOFA, "'", "''") & "'"
If newY.NOTPAYCOFK <> oldY.NOTPAYCOFK Then blnUpdate = True: xSet = xSet & ",NOTPAYCOFK = '" & Replace(newY.NOTPAYCOFK, "'", "''") & "'"
If newY.NOTPAYCOF2 <> oldY.NOTPAYCOF2 Then blnUpdate = True: xSet = xSet & ",NOTPAYCOF2 = '" & Replace(newY.NOTPAYCOF2, "'", "''") & "'"
If newY.NOTPAYOCDE <> oldY.NOTPAYOCDE Then blnUpdate = True: xSet = xSet & ",NOTPAYOCDE = '" & Replace(newY.NOTPAYOCDE, "'", "''") & "'"
If newY.NOTPAYOCDK <> oldY.NOTPAYOCDK Then blnUpdate = True: xSet = xSet & ",NOTPAYOCDK = '" & Replace(newY.NOTPAYOCDK, "'", "''") & "'"
If newY.NOTPAYSP <> oldY.NOTPAYSP Then blnUpdate = True: xSet = xSet & ",NOTPAYSP = '" & Replace(newY.NOTPAYSP, "'", "''") & "'"
If newY.NOTPAYSPK <> oldY.NOTPAYSPK Then blnUpdate = True: xSet = xSet & ",NOTPAYSPK = '" & Replace(newY.NOTPAYSPK, "'", "''") & "'"
If newY.NOTPAYBIAN <> oldY.NOTPAYBIAN Then blnUpdate = True: xSet = xSet & ",NOTPAYBIAN = '" & Replace(newY.NOTPAYBIAN, "'", "''") & "'"
If newY.NOTPAYBIAK <> oldY.NOTPAYBIAK Then blnUpdate = True: xSet = xSet & ",NOTPAYBIAK = '" & Replace(newY.NOTPAYBIAK, "'", "''") & "'"
If newY.NOTPAYFISC <> oldY.NOTPAYFISC Then blnUpdate = True: xSet = xSet & ",NOTPAYFISC = '" & Replace(newY.NOTPAYFISC, "'", "''") & "'"
If Trim(newY.NOTPAYTXT) <> Trim(oldY.NOTPAYTXT) Then blnUpdate = True: xSet = xSet & ",NOTPAYTXT = '" & Replace(Trim(newY.NOTPAYTXT), "'", "''") & "'"
If newY.NOTPAYXUSR <> oldY.NOTPAYXUSR Then blnUpdate = True: xSet = xSet & ",NOTPAYXUSR = '" & Replace(newY.NOTPAYXUSR, "'", "''") & "'"

If blnUpdate Then

    Mid$(xSet, InStr(xSet, ","), 1) = " "
    xSQL = "update " & paramIBM_Library_SABSPE & ".YNOTPAY0" & xSet & xWhere
    Call FEU_ROUGE
    Set rsSab = cnsab.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYNOTPAY0_Update = "Erreur màj : " & newY.NOTPAYSEQ
    
        Exit Function
    End If
End If
Exit Function
Error_Handler:
    sqlYNOTPAY0_Update = Error
End Function



Public Function sqlYNOTPAY0_Compare(newY As typeYNOTPAY0, oldY As typeYNOTPAY0) As Boolean
Dim X As String
Dim blnCompare As Boolean

sqlYNOTPAY0_Compare = True
blnCompare = True


' Détecter les modifications
'===================================================================================
If newY.NOTPAYISO <> oldY.NOTPAYISO Then blnCompare = False
If newY.NOTPAYPROV <> oldY.NOTPAYPROV Then blnCompare = False
If newY.NOTPAYHAMJ <> oldY.NOTPAYHAMJ Then blnCompare = False

If newY.NOTPAYCOFD <> oldY.NOTPAYCOFD Then blnCompare = False
If newY.NOTPAYCOFA <> oldY.NOTPAYCOFA Then blnCompare = False
If newY.NOTPAYCOFK <> oldY.NOTPAYCOFK Then blnCompare = False
If newY.NOTPAYCOF2 <> oldY.NOTPAYCOF2 Then blnCompare = False

If newY.NOTPAYOCDE <> oldY.NOTPAYOCDE Then blnCompare = False
If newY.NOTPAYOCDK <> oldY.NOTPAYOCDK Then blnCompare = False
If newY.NOTPAYOCDD <> oldY.NOTPAYOCDD Then blnCompare = False

If newY.NOTPAYSP <> oldY.NOTPAYSP Then blnCompare = False
If newY.NOTPAYSPK <> oldY.NOTPAYSPK Then blnCompare = False
If newY.NOTPAYSPD <> oldY.NOTPAYSPD Then blnCompare = False

If newY.NOTPAYCEG <> oldY.NOTPAYCEG Then blnCompare = False
If newY.NOTPAYBIAN <> oldY.NOTPAYBIAN Then blnCompare = False
If newY.NOTPAYBIAK <> oldY.NOTPAYBIAK Then blnCompare = False
If newY.NOTPAYBIAD <> oldY.NOTPAYBIAD Then blnCompare = False
If newY.NOTPAYTAUX <> oldY.NOTPAYTAUX Then blnCompare = False

If newY.NOTPAYFISC <> oldY.NOTPAYFISC Then blnCompare = False

If Trim(newY.NOTPAYTXT) <> Trim(oldY.NOTPAYTXT) Then blnCompare = False

sqlYNOTPAY0_Compare = blnCompare

End Function


'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsYNOTPAY0_GetBuffer(rsAdo As ADODB.Recordset, rsYNOTPAY0 As typeYNOTPAY0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYNOTPAY0_GetBuffer = Null

rsYNOTPAY0.JORCV = 0
rsYNOTPAY0.JOSEQN = 0
rsYNOTPAY0.JRNBIATRN = 0
rsYNOTPAY0.NOTPAYLIB = "?"

rsYNOTPAY0.NOTPAYISO = rsAdo("NOTPAYISO")
rsYNOTPAY0.NOTPAYSEQ = rsAdo("NOTPAYSEQ")
rsYNOTPAY0.NOTPAYHAMJ = rsAdo("NOTPAYHAMJ")
rsYNOTPAY0.NOTPAYPROV = rsAdo("NOTPAYPROV")
rsYNOTPAY0.NOTPAYCOFA = rsAdo("NOTPAYCOFA")
rsYNOTPAY0.NOTPAYCOFK = rsAdo("NOTPAYCOFK")
rsYNOTPAY0.NOTPAYCOFD = rsAdo("NOTPAYCOFD")
rsYNOTPAY0.NOTPAYCOF2 = rsAdo("NOTPAYCOF2")
rsYNOTPAY0.NOTPAYOCDE = rsAdo("NOTPAYOCDE")
rsYNOTPAY0.NOTPAYOCDK = rsAdo("NOTPAYOCDK")
rsYNOTPAY0.NOTPAYOCDD = rsAdo("NOTPAYOCDD")
rsYNOTPAY0.NOTPAYSP = rsAdo("NOTPAYSP")
rsYNOTPAY0.NOTPAYSPK = rsAdo("NOTPAYSPK")
rsYNOTPAY0.NOTPAYSPD = rsAdo("NOTPAYSPD")
rsYNOTPAY0.NOTPAYCEG = rsAdo("NOTPAYCEG")
rsYNOTPAY0.NOTPAYBIAN = rsAdo("NOTPAYBIAN")
rsYNOTPAY0.NOTPAYBIAK = rsAdo("NOTPAYBIAK")
rsYNOTPAY0.NOTPAYBIAD = rsAdo("NOTPAYBIAD")
rsYNOTPAY0.NOTPAYTAUX = rsAdo("NOTPAYTAUX")
rsYNOTPAY0.NOTPAYFISC = rsAdo("NOTPAYFISC")
rsYNOTPAY0.NOTPAYTXT = rsAdo("NOTPAYTXT")
rsYNOTPAY0.NOTPAYXAMJ = rsAdo("NOTPAYXAMJ")
rsYNOTPAY0.NOTPAYXHMS = rsAdo("NOTPAYXHMS")
rsYNOTPAY0.NOTPAYXUSR = rsAdo("NOTPAYXUSR")

Exit Function

Error_Handler:

rsYNOTPAY0_GetBuffer = Error

End Function

'---------------------------------------------------------
Public Function rsJNOTPAY0_GetBuffer(rsAdo As ADODB.Recordset, rsYNOTPAY0 As typeYNOTPAY0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsJNOTPAY0_GetBuffer = Null

rsJNOTPAY0_GetBuffer = rsYNOTPAY0_GetBuffer(rsAdo, rsYNOTPAY0)
rsYNOTPAY0.JORCV = rsAdo("JORCV")
rsYNOTPAY0.JOSEQN = rsAdo("JOSEQN")
rsYNOTPAY0.JRNBIATRN = rsAdo("JRNBIATRN")

Exit Function

Error_Handler:

rsJNOTPAY0_GetBuffer = Error

End Function



Public Sub ddsYNOTPAY0_Init()

ddsYNOTPAY0(1) = "A* NOTPAYISO  code ISO"
ddsYNOTPAY0(2) = "N* NOTPAYSEQ  N° séquence"
ddsYNOTPAY0(3) = "N* NOTPAYHAMJ Date d'arrêté"
ddsYNOTPAY0(4) = "A  NOTPAYPROV Provision"
ddsYNOTPAY0(5) = "A  NOTPAYCOFA notation Coface"
ddsYNOTPAY0(6) = "A  NOTPAYCOFD notation Coface màj"
ddsYNOTPAY0(7) = "A* NOTPAYCOFK notation Coface A/M"
ddsYNOTPAY0(8) = "A  NOTPAYCOF2 notation Coface env affaires"
ddsYNOTPAY0(9) = "A  NOTPAYOCDE notation OCDE"
ddsYNOTPAY0(10) = "A  NOTPAYOCDD notation OCDE màj"
ddsYNOTPAY0(11) = "A* NOTPAYOCDK notation OCDE A/M"
ddsYNOTPAY0(12) = "A  NOTPAYSP   notation S & P"
ddsYNOTPAY0(13) = "A  NOTPAYSPD  notation S & P màj"
ddsYNOTPAY0(14) = "A* NOTPAYSPK  notation S & P A/M"
ddsYNOTPAY0(15) = "S  NOTPAYCEG  Critère événement grave"
ddsYNOTPAY0(16) = "A  NOTPAYBIAN notation BIA"
NOTPAYBIAN_Row = 16
ddsYNOTPAY0(17) = "A  NOTPAYBIAD notation BIA màj"
ddsYNOTPAY0(18) = "A* NOTPAYBIAK notation BIA A/M"
ddsYNOTPAY0(19) = "N* NOTPAYTAUX Taux BIA"
NOTPAYTAUX_Row = 19
ddsYNOTPAY0(20) = "A  NOTPAYFISC Taux Fisc"
ddsYNOTPAY0(21) = "A  NOTPAYTXT  commentaire"
ddsYNOTPAY0(22) = "N* NOTPAYXAMJ Date mise à jour"
ddsYNOTPAY0(23) = "N* NOTPAYXHMS Heure  mise à jour"
ddsYNOTPAY0(24) = "A* NOTPAYXUSR Utilisateur  mise à jour"

End Sub


