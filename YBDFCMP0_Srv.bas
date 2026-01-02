Attribute VB_Name = "srvYBDFCMP0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------


Dim rsAdo As ADODB.Recordset
 
Type typeYBDFCMP0
 
      BDFCMPSER   As String * 2     ' service
      BDFCMPSSE   As String * 2
      BDFCMPOPE   As String * 3   'code opération
      BDFCMPNAT   As String * 3   'nature opération
      BDFCMPDOS   As Long         'N° opération
      BDFCMPMON   As Currency     'montant
      BDFCMPDEV   As String * 3   'devise
      BDFCMPMONE  As Currency     'montant CV EUR
      BDFCMPDCRE  As Long         ' date création
      BDFCMPDOPE  As Long         ' date opérartion
      BDFCMPCREG  As String * 3   'code réglement
      BDFCMPXDB   As String * 20   'compte DB
      BDFCMPXDBN  As String * 3   'nature compte DB
      BDFCMPXCR   As String * 20   'compte CR
      BDFCMPXCRN  As String * 3   'nature compte CR
      BDFCMPBBIC  As String * 8   'BIC bénéficiaire
      BDFCMPPAYS  As String * 2   'pays bénéficiaire
      BDFCMPSTAT  As String * 3   'grille déclaration
      BDFCMPSTA   As String * 1   'statut
      BDFCMPUPDS   As Long         'séquence
      BDFCMPUSR   As String * 10   'utilisateur
      BDFCMPSEQ   As Long         'séquence (clé en double)
      BDFCMP2008  As String * 4   'version 2008
      BDFCMPSABK  As Long  '
      BDFCMPMTK  As String * 3   '
      BDFCMPROUT  As String * 1   '
      BDFCMP50PI  As String * 2   '
      BDFCMP59PI  As String * 2   '
      End Type
Public xYBDFCMP0 As typeYBDFCMP0
Public Function sqlYBDFCMP0_Insert(newY As typeYBDFCMP0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String
Static mBDFCMPSEQ As Long

On Error GoTo Error_Handler
sqlYBDFCMP0_Insert = Null

mBDFCMPSEQ = mBDFCMPSEQ + 1
newY.BDFCMPSEQ = mBDFCMPSEQ
xSet = " (BDFCMPDOS"
xValues = " values(" & newY.BDFCMPDOS

' Détecter les modifications
'===================================================================================
'''If newY.BDFCMPDOS <> 0 Then xSet = xSet & ",BDFCMPDOS": xValues = xValues & " ," & newY.BDFCMPDOS
If newY.BDFCMPMON <> 0 Then xSet = xSet & ",BDFCMPMON": xValues = xValues & " ," & cur_P(newY.BDFCMPMON)
If newY.BDFCMPDCRE <> 0 Then xSet = xSet & ",BDFCMPDCRE": xValues = xValues & " ," & newY.BDFCMPDCRE
If newY.BDFCMPDOPE <> 0 Then xSet = xSet & ",BDFCMPDOPE": xValues = xValues & " ," & newY.BDFCMPDOPE
If newY.BDFCMPSTAT <> 0 Then xSet = xSet & ",BDFCMPSTAT": xValues = xValues & " ," & newY.BDFCMPSTAT
If newY.BDFCMPMONE <> 0 Then xSet = xSet & ",BDFCMPMONE": xValues = xValues & " ," & cur_P(newY.BDFCMPMONE)
If newY.BDFCMPSEQ <> 0 Then xSet = xSet & ",BDFCMPSEQ": xValues = xValues & " ," & newY.BDFCMPSEQ

If Trim(newY.BDFCMPSER) <> "" Then xSet = xSet & ",BDFCMPSER": xValues = xValues & " ,'" & newY.BDFCMPSER & "'"
If Trim(newY.BDFCMPSSE) <> "" Then xSet = xSet & ",BDFCMPSSE": xValues = xValues & " ,'" & newY.BDFCMPSSE & "'"
If Trim(newY.BDFCMPOPE) <> "" Then xSet = xSet & ",BDFCMPOPE": xValues = xValues & " ,'" & newY.BDFCMPOPE & "'"
If Trim(newY.BDFCMPNAT) <> "" Then xSet = xSet & ",BDFCMPNAT": xValues = xValues & " ,'" & newY.BDFCMPNAT & "'"
If Trim(newY.BDFCMPDEV) <> "" Then xSet = xSet & ",BDFCMPDEV": xValues = xValues & " ,'" & newY.BDFCMPDEV & "'"
If Trim(newY.BDFCMPCREG) <> "" Then xSet = xSet & ",BDFCMPCREG": xValues = xValues & " ,'" & Trim(newY.BDFCMPCREG) & "'"
If Trim(newY.BDFCMPXDB) <> "" Then xSet = xSet & ",BDFCMPXDB": xValues = xValues & " ,'" & Replace(Trim(newY.BDFCMPXDB), "'", "''") & "'"
If Trim(newY.BDFCMPXDBN) <> "" Then xSet = xSet & ",BDFCMPXDBN": xValues = xValues & " ,'" & Replace(Trim(newY.BDFCMPXDBN), "'", "''") & "'"
If Trim(newY.BDFCMPXCR) <> "" Then xSet = xSet & ",BDFCMPXCR": xValues = xValues & " ,'" & Replace(Trim(newY.BDFCMPXCR), "'", "''") & "'"
If Trim(newY.BDFCMPXCRN) <> "" Then xSet = xSet & ",BDFCMPXCRN": xValues = xValues & " ,'" & Replace(Trim(newY.BDFCMPXCRN), "'", "''") & "'"
If Trim(newY.BDFCMPBBIC) <> "" Then xSet = xSet & ",BDFCMPBBIC": xValues = xValues & " ,'" & newY.BDFCMPBBIC & "'"
If Trim(newY.BDFCMPPAYS) <> "" Then xSet = xSet & ",BDFCMPPAYS": xValues = xValues & " ,'" & newY.BDFCMPPAYS & "'"
If Trim(newY.BDFCMPSTA) <> "" Then xSet = xSet & ",BDFCMPSTA": xValues = xValues & " ,'" & newY.BDFCMPSTA & "'"
If Trim(newY.BDFCMP2008) <> "" Then xSet = xSet & ",BDFCMP2008": xValues = xValues & " ,'" & newY.BDFCMP2008 & "'"

newY.BDFCMPUSR = usrName_UCase
xSet = xSet & ",BDFCMPUSR": xValues = xValues & " ,'" & usrName_UCase & "'"

xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YBDFCMP0" & xSet & ")" & xValues & ")"
Call FEU_ROUGE
Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYBDFCMP0_Insert = "Erreur màj : " & newY.BDFCMPSER
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYBDFCMP0_Insert = Error
End Function

Public Function sqlYBDFCMP0_Update(newY As typeYBDFCMP0, oldY As typeYBDFCMP0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYBDFCMP0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.BDFCMPDOS <> newY.BDFCMPDOS Then
    sqlYBDFCMP0_Update = "Erreur BDFCMPDOS : " & newY.BDFCMPDOS & " / " & oldY.BDFCMPDOS
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where BDFCMPDOS = " & oldY.BDFCMPDOS _
       & " and BDFCMPUPDS = " & oldY.BDFCMPUPDS _
       & " and BDFCMPOPE = '" & oldY.BDFCMPOPE & "'" _
       & " and BDFCMPDCRE = '" & oldY.BDFCMPDCRE & "'" _
       & " and BDFCMPSEQ = '" & oldY.BDFCMPSEQ & "'" _
       & " and BDFCMPNAT = '" & oldY.BDFCMPNAT & "'" _
       & " and BDFCMPSER = '" & oldY.BDFCMPSER & "'" _
       & " and BDFCMPSSE = '" & oldY.BDFCMPSSE & "'"

newY.BDFCMPUPDS = newY.BDFCMPUPDS + 1
xSet = xSet & " set BDFCMPUPDS = " & newY.BDFCMPUPDS
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.BDFCMPMON <> oldY.BDFCMPMON Then blnUpdate = True:  xSet = xSet & " , BDFCMPMON = '" & cur_P(newY.BDFCMPMON) & "'"
If newY.BDFCMPDCRE <> oldY.BDFCMPDCRE Then blnUpdate = True: xSet = xSet & " , BDFCMPDCRE = " & newY.BDFCMPDCRE
If newY.BDFCMPDOPE <> oldY.BDFCMPDOPE Then blnUpdate = True: xSet = xSet & " , BDFCMPDOPE = " & newY.BDFCMPDOPE
If newY.BDFCMPMONE <> oldY.BDFCMPMONE Then blnUpdate = True:  xSet = xSet & " , BDFCMPMONE = '" & cur_P(newY.BDFCMPMONE) & "'"
If newY.BDFCMPSTAT <> oldY.BDFCMPSTAT Then blnUpdate = True: xSet = xSet & " , BDFCMPSTAT = " & newY.BDFCMPSTAT
If newY.BDFCMPSEQ <> oldY.BDFCMPSEQ Then blnUpdate = True: xSet = xSet & " , BDFCMPSeq= " & newY.BDFCMPSEQ


If newY.BDFCMPSER <> oldY.BDFCMPSER Then blnUpdate = True:  xSet = xSet & " , BDFCMPSER = '" & Trim(newY.BDFCMPSER) & "'"
If newY.BDFCMPSSE <> oldY.BDFCMPSSE Then blnUpdate = True:  xSet = xSet & " , BDFCMPSSE= '" & newY.BDFCMPSSE & "'"
If newY.BDFCMPOPE <> oldY.BDFCMPOPE Then blnUpdate = True:  xSet = xSet & " , BDFCMPOPE = '" & newY.BDFCMPOPE & "'"
If newY.BDFCMPNAT <> oldY.BDFCMPNAT Then blnUpdate = True: xSet = xSet & " , BDFCMPNAT = " & newY.BDFCMPNAT
If newY.BDFCMPDEV <> oldY.BDFCMPDEV Then blnUpdate = True:  xSet = xSet & " , BDFCMPDEV= '" & newY.BDFCMPDEV & "'"
If newY.BDFCMPCREG <> oldY.BDFCMPCREG Then blnUpdate = True:  xSet = xSet & " , BDFCMPCREG = '" & Replace(Trim(newY.BDFCMPCREG), "'", "''") & "'"
If newY.BDFCMPXDB <> oldY.BDFCMPXDB Then blnUpdate = True:  xSet = xSet & " , BDFCMPXDB = '" & Replace(Trim(newY.BDFCMPXDB), "'", "''") & "'"
If newY.BDFCMPXDBN <> oldY.BDFCMPXDBN Then blnUpdate = True:  xSet = xSet & " , BDFCMPXDBN = '" & Replace(Trim(newY.BDFCMPXDBN), "'", "''") & "'"
If newY.BDFCMPXCR <> oldY.BDFCMPXCR Then blnUpdate = True:  xSet = xSet & " , BDFCMPXCR = '" & Replace(Trim(newY.BDFCMPXCR), "'", "''") & "'"
If newY.BDFCMPXCRN <> oldY.BDFCMPXCRN Then blnUpdate = True:  xSet = xSet & " , BDFCMPXCRN = '" & Replace(Trim(newY.BDFCMPXCRN), "'", "''") & "'"
If newY.BDFCMPBBIC <> oldY.BDFCMPBBIC Then blnUpdate = True:  xSet = xSet & " , BDFCMPBBIC = '" & Replace(Trim(newY.BDFCMPBBIC), "'", "''") & "'"
If newY.BDFCMPPAYS <> oldY.BDFCMPPAYS Then blnUpdate = True:  xSet = xSet & " , BDFCMPPAYS = '" & Replace(Trim(newY.BDFCMPPAYS), "'", "''") & "'"
If newY.BDFCMPSTA <> oldY.BDFCMPSTA Then blnUpdate = True:  xSet = xSet & " , BDFCMPSTA = '" & newY.BDFCMPSTA & "'"
If newY.BDFCMP2008 <> oldY.BDFCMP2008 Then blnUpdate = True:  xSet = xSet & " , BDFCMP2008 = '" & newY.BDFCMP2008 & "'"

newY.BDFCMPUSR = usrName_UCase
xSet = xSet & " , BDFCMPUSR = '" & usrName_UCase & "'"
If newY.BDFCMPSER < 0 Then blnUpdate = True  ' records techniques

If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE & ".YBDFCMP0" & xSet & xWhere
    Call FEU_ROUGE
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYBDFCMP0_Update = "Erreur màj : " & newY.BDFCMPSER
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYBDFCMP0_Update = Error
End Function

Public Function rsYBDFCMP0_GetBuffer(rsAdo As ADODB.Recordset, lYBDFCMP0 As typeYBDFCMP0)
On Error GoTo Error_Handler
rsYBDFCMP0_GetBuffer = Null
lYBDFCMP0.BDFCMPSER = rsAdo("BDFCMPSER")
lYBDFCMP0.BDFCMPSSE = rsAdo("BDFCMPSSE")
lYBDFCMP0.BDFCMPOPE = rsAdo("BDFCMPOPE")
lYBDFCMP0.BDFCMPDOS = rsAdo("BDFCMPDOS")
lYBDFCMP0.BDFCMPNAT = rsAdo("BDFCMPNAT")
lYBDFCMP0.BDFCMPMON = rsAdo("BDFCMPMON")
lYBDFCMP0.BDFCMPDEV = rsAdo("BDFCMPDEV")
lYBDFCMP0.BDFCMPMONE = rsAdo("BDFCMPMONE")
lYBDFCMP0.BDFCMPDCRE = rsAdo("BDFCMPDCRE")
lYBDFCMP0.BDFCMPDOPE = rsAdo("BDFCMPDOPE")
lYBDFCMP0.BDFCMPCREG = rsAdo("BDFCMPCREG")
lYBDFCMP0.BDFCMPXDB = rsAdo("BDFCMPXDB")
lYBDFCMP0.BDFCMPXDBN = rsAdo("BDFCMPXDBN")
lYBDFCMP0.BDFCMPXCR = rsAdo("BDFCMPXCR")
lYBDFCMP0.BDFCMPXCRN = rsAdo("BDFCMPXCRN")
lYBDFCMP0.BDFCMPBBIC = rsAdo("BDFCMPBBIC")
lYBDFCMP0.BDFCMPPAYS = rsAdo("BDFCMPPAYS")
lYBDFCMP0.BDFCMPSTAT = rsAdo("BDFCMPSTAT")

lYBDFCMP0.BDFCMPSTA = rsAdo("BDFCMPSTA")
lYBDFCMP0.BDFCMPUPDS = rsAdo("BDFCMPUPDS")
lYBDFCMP0.BDFCMPUSR = rsAdo("BDFCMPUSR")
lYBDFCMP0.BDFCMPSEQ = rsAdo("BDFCMPSEQ")
lYBDFCMP0.BDFCMP2008 = rsAdo("BDFCMP2008")
lYBDFCMP0.BDFCMPSABK = rsAdo("BDFCMPSABK")
lYBDFCMP0.BDFCMPMTK = rsAdo("BDFCMPMTK")
lYBDFCMP0.BDFCMPROUT = rsAdo("BDFCMPROUT")
lYBDFCMP0.BDFCMP50PI = rsAdo("BDFCMP50PI")
lYBDFCMP0.BDFCMP59PI = rsAdo("BDFCMP59PI")




Exit Function
Error_Handler:
rsYBDFCMP0_GetBuffer = Error


End Function

Public Function rsYBDFCMP0_Init(lYBDFCMP0 As typeYBDFCMP0)
lYBDFCMP0.BDFCMPSER = ""
lYBDFCMP0.BDFCMPSSE = ""
lYBDFCMP0.BDFCMPDOS = 0
lYBDFCMP0.BDFCMPNAT = ""
lYBDFCMP0.BDFCMPOPE = ""
lYBDFCMP0.BDFCMPMON = 0
lYBDFCMP0.BDFCMPDEV = ""
lYBDFCMP0.BDFCMPMONE = 0
lYBDFCMP0.BDFCMPDCRE = 0
lYBDFCMP0.BDFCMPDOPE = 0
lYBDFCMP0.BDFCMPCREG = ""
lYBDFCMP0.BDFCMPXDB = ""
lYBDFCMP0.BDFCMPXDBN = ""
lYBDFCMP0.BDFCMPXCR = ""
lYBDFCMP0.BDFCMPXCRN = ""
lYBDFCMP0.BDFCMPBBIC = ""
lYBDFCMP0.BDFCMPPAYS = ""
lYBDFCMP0.BDFCMPSTAT = ""
lYBDFCMP0.BDFCMPSTA = ""
lYBDFCMP0.BDFCMPSTAT = 0
lYBDFCMP0.BDFCMPUPDS = 0
lYBDFCMP0.BDFCMPUSR = ""
lYBDFCMP0.BDFCMPSEQ = 0
lYBDFCMP0.BDFCMP2008 = ""
lYBDFCMP0.BDFCMPSABK = 0
lYBDFCMP0.BDFCMPMTK = ""
lYBDFCMP0.BDFCMPROUT = ""
lYBDFCMP0.BDFCMP50PI = ""
lYBDFCMP0.BDFCMP59PI = ""

End Function



