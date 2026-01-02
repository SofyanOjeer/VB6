Attribute VB_Name = "srvYKYCSTA0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
 
 '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
 ' mise à jour sans journalisation
 '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
 
 
Type typeYKYCSTA0
 
      KYCSTACLI     As String
      KYCSTADSIT    As Long
      
      KYCSTASTAK    As String
      KYCSTASTAX    As String
      KYCSTASTAY    As String
      
      KYCSTACAVC    As Long
      KYCSTACAVT    As Long
      KYCSTACAVX    As Long
      
      KYCSTATECC    As Long
      KYCSTATECT    As Long
      KYCSTATECX    As Long
      
      KYCSTADCLO    As Long
      
      KYCSTAZCOL    As String
      KYCSTAZETA    As String
      KYCSTAZCAT    As String
      KYCSTAZRES    As String
      KYCSTAZRA1    As String
      KYCSTAZPCI    As String
      KYCSTAYKYC    As String
      KYCSTAZNAT    As String
      KYCSTAZRSD    As String
      
      KYCSTAYFCT    As String
      KYCSTAYUSR    As String
      KYCSTAYAMJ    As Long
      KYCSTAYHMS    As Long
      KYCSTAYVER    As Long
End Type
Public Function sqlYKYCSTA0_Delete(oldY As typeYKYCSTA0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYKYCSTA0_Delete = Null
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================
    xWhere = " where KYCSTACLI = '" & oldY.KYCSTACLI & "'" _
           & " and KYCSTADSIT = " & oldY.KYCSTADSIT

'===================================================================================

    
    xSQL = "delete from " & paramIBM_Library_SABSPE & ".YKYCSTA0" & xWhere
    
    'Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_ROUGE
    Set rsAdo = cnsab.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYKYCSTA0_Delete = "Erreur màj : " & xWhere
        Exit Function
    End If
    
Exit Function
Error_Handler:
    sqlYKYCSTA0_Delete = Error
End Function

Public Function sqlYKYCSTA0_Delete_Where(lWhere As String)
Dim X As String, xSQL As String, Nb As Long

On Error GoTo Error_Handler
sqlYKYCSTA0_Delete_Where = Null

    
    xSQL = "delete from " & paramIBM_Library_SABSPE & ".YKYCSTA0" & lWhere
    
    'Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_ROUGE
    Set rsAdo = cnsab.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYKYCSTA0_Delete_Where = "Erreur màj : " & lWhere
        Exit Function
    End If
    

Exit Function
Error_Handler:
    sqlYKYCSTA0_Delete_Where = Error
End Function

Public Function sqlYKYCSTA0_Insert(newY As typeYKYCSTA0)
Dim V
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYKYCSTA0_Insert = Null
xSet = " ( KYCSTACLI , KYCSTADSIT "
xValues = " values('" & newY.KYCSTACLI & "' ," & newY.KYCSTADSIT

newY.KYCSTAYUSR = usrName_UCase
newY.KYCSTAYAMJ = DSys
newY.KYCSTAYHMS = time_Hms

' Détecter les modifications
'===================================================================================

If Trim(newY.KYCSTACAVC) <> "" Then xSet = xSet & ",KYCSTACAVC": xValues = xValues & " ," & newY.KYCSTACAVC
If Trim(newY.KYCSTACAVT) <> "" Then xSet = xSet & ",KYCSTACAVT": xValues = xValues & " ," & newY.KYCSTACAVT
If Trim(newY.KYCSTACAVX) <> "" Then xSet = xSet & ",KYCSTACAVX": xValues = xValues & " ," & newY.KYCSTACAVX

If Trim(newY.KYCSTATECC) <> "" Then xSet = xSet & ",KYCSTATECC": xValues = xValues & " ," & newY.KYCSTATECC
If Trim(newY.KYCSTATECT) <> "" Then xSet = xSet & ",KYCSTATECT": xValues = xValues & " ," & newY.KYCSTATECT
If Trim(newY.KYCSTATECX) <> "" Then xSet = xSet & ",KYCSTATECX": xValues = xValues & " ," & newY.KYCSTATECX

If Trim(newY.KYCSTADCLO) <> "" Then xSet = xSet & ",KYCSTADCLO": xValues = xValues & " ," & newY.KYCSTADCLO

If Trim(newY.KYCSTASTAK) <> "" Then xSet = xSet & ",KYCSTASTAK": xValues = xValues & " ,'" & newY.KYCSTASTAK & "'"
If Trim(newY.KYCSTASTAX) <> "" Then xSet = xSet & ",KYCSTASTAX": xValues = xValues & " ,'" & newY.KYCSTASTAX & "'"
If Trim(newY.KYCSTASTAY) <> "" Then xSet = xSet & ",KYCSTASTAY": xValues = xValues & " ,'" & newY.KYCSTASTAY & "'"
If Trim(newY.KYCSTAZCOL) <> "" Then xSet = xSet & ",KYCSTAZCOL": xValues = xValues & " ,'" & Replace(Trim(newY.KYCSTAZCOL), "'", "''") & "'"
If Trim(newY.KYCSTAZETA) <> "" Then xSet = xSet & ",KYCSTAZETA": xValues = xValues & " ,'" & Replace(Trim(newY.KYCSTAZETA), "'", "''") & "'"
If Trim(newY.KYCSTAZCAT) <> "" Then xSet = xSet & ",KYCSTAZCAT": xValues = xValues & " ,'" & Replace(Trim(newY.KYCSTAZCAT), "'", "''") & "'"
If Trim(newY.KYCSTAZRES) <> "" Then xSet = xSet & ",KYCSTAZRES": xValues = xValues & " ,'" & Replace(Trim(newY.KYCSTAZRES), "'", "''") & "'"
If Trim(newY.KYCSTAZRA1) <> "" Then xSet = xSet & ",KYCSTAZRA1": xValues = xValues & " ,'" & Replace(Trim(newY.KYCSTAZRA1), "'", "''") & "'"
If Trim(newY.KYCSTAZPCI) <> "" Then xSet = xSet & ",KYCSTAZPCI": xValues = xValues & " ,'" & Replace(Trim(newY.KYCSTAZPCI), "'", "''") & "'"
If Trim(newY.KYCSTAZNAT) <> "" Then xSet = xSet & ",KYCSTAZNAT": xValues = xValues & " ,'" & Replace(Trim(newY.KYCSTAZNAT), "'", "''") & "'"
If Trim(newY.KYCSTAZRSD) <> "" Then xSet = xSet & ",KYCSTAZRSD": xValues = xValues & " ,'" & Replace(Trim(newY.KYCSTAZRSD), "'", "''") & "'"
If Trim(newY.KYCSTAYKYC) <> "" Then xSet = xSet & ",KYCSTAYKYC": xValues = xValues & " ,'" & Replace(Trim(newY.KYCSTAYKYC), "'", "''") & "'"


If Trim(newY.KYCSTAYFCT) <> "" Then xSet = xSet & ",KYCSTAYFCT": xValues = xValues & " ,'" & newY.KYCSTAYFCT & "'"
If Trim(newY.KYCSTAYUSR) <> "" Then xSet = xSet & ",KYCSTAYUSR": xValues = xValues & " ,'" & newY.KYCSTAYUSR & "'"
If newY.KYCSTAYVER <> 0 Then xSet = xSet & ",KYCSTAYVER": xValues = xValues & " ," & newY.KYCSTAYVER
If Trim(newY.KYCSTAYAMJ) <> "" Then xSet = xSet & ",KYCSTAYAMJ": xValues = xValues & " ," & newY.KYCSTAYAMJ
If Trim(newY.KYCSTAYHMS) <> "" Then xSet = xSet & ",KYCSTAYHMS": xValues = xValues & " ," & newY.KYCSTAYHMS

Call FEU_ROUGE
xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YKYCSTA0" & xSet & ")" & xValues & ")"

'Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Set rsAdo = cnsab.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYKYCSTA0_Insert = "Erreur màj : " & newY.KYCSTACLI & newY.KYCSTADSIT
    Exit Function
End If


Exit Function
Error_Handler:
    sqlYKYCSTA0_Insert = Error
End Function

Public Function rsYKYCSTA0_GetBuffer(rsAdo As ADODB.Recordset, lYKYCSTA0 As typeYKYCSTA0)
On Error GoTo Error_Handler
rsYKYCSTA0_GetBuffer = Null

lYKYCSTA0.KYCSTACLI = Trim(rsAdo("KYCSTACLI"))
lYKYCSTA0.KYCSTADSIT = rsAdo("KYCSTADSIT")

lYKYCSTA0.KYCSTASTAK = rsAdo("KYCSTASTAK")
lYKYCSTA0.KYCSTASTAX = rsAdo("KYCSTASTAX")
lYKYCSTA0.KYCSTASTAY = rsAdo("KYCSTASTAY")

lYKYCSTA0.KYCSTACAVC = rsAdo("KYCSTACAVC")
lYKYCSTA0.KYCSTACAVT = rsAdo("KYCSTACAVT")
lYKYCSTA0.KYCSTACAVX = rsAdo("KYCSTACAVX")

lYKYCSTA0.KYCSTATECC = rsAdo("KYCSTATECC")
lYKYCSTA0.KYCSTATECT = rsAdo("KYCSTATECT")
lYKYCSTA0.KYCSTATECX = rsAdo("KYCSTATECX")

lYKYCSTA0.KYCSTADCLO = Trim(rsAdo("KYCSTADCLO"))
lYKYCSTA0.KYCSTAZCOL = Trim(rsAdo("KYCSTAZCOL"))
lYKYCSTA0.KYCSTAZETA = Trim(rsAdo("KYCSTAZETA"))
lYKYCSTA0.KYCSTAZCAT = Trim(rsAdo("KYCSTAZCAT"))
lYKYCSTA0.KYCSTAZRES = Trim(rsAdo("KYCSTAZRES"))
lYKYCSTA0.KYCSTAZRA1 = Trim(rsAdo("KYCSTAZRA1"))
lYKYCSTA0.KYCSTAZPCI = Trim(rsAdo("KYCSTAZPCI"))
lYKYCSTA0.KYCSTAZNAT = Trim(rsAdo("KYCSTAZNAT"))
lYKYCSTA0.KYCSTAZRSD = Trim(rsAdo("KYCSTAZRSD"))

lYKYCSTA0.KYCSTAYKYC = Trim(rsAdo("KYCSTAYKYC"))

lYKYCSTA0.KYCSTAYUSR = Trim(rsAdo("KYCSTAYUSR"))
lYKYCSTA0.KYCSTAYAMJ = rsAdo("KYCSTAYAMJ")
lYKYCSTA0.KYCSTAYHMS = rsAdo("KYCSTAYHMS")
lYKYCSTA0.KYCSTAYVER = rsAdo("KYCSTAYVER")
lYKYCSTA0.KYCSTAYFCT = rsAdo("KYCSTAYFCT")

Exit Function
Error_Handler:
rsYKYCSTA0_GetBuffer = Error


End Function

Public Function rsYKYCSTA0_Init(lYKYCSTA0 As typeYKYCSTA0)

lYKYCSTA0.KYCSTACLI = ""
lYKYCSTA0.KYCSTADSIT = 0
      
lYKYCSTA0.KYCSTASTAK = ""
lYKYCSTA0.KYCSTASTAX = ""
lYKYCSTA0.KYCSTASTAY = ""
      
lYKYCSTA0.KYCSTACAVC = 0
lYKYCSTA0.KYCSTACAVT = 0
lYKYCSTA0.KYCSTACAVX = 0
      
lYKYCSTA0.KYCSTATECC = 0
lYKYCSTA0.KYCSTATECT = 0
lYKYCSTA0.KYCSTATECX = 0
      
lYKYCSTA0.KYCSTADCLO = 0
      
lYKYCSTA0.KYCSTAZCOL = ""
lYKYCSTA0.KYCSTAZETA = ""
lYKYCSTA0.KYCSTAZCAT = ""
lYKYCSTA0.KYCSTAZRES = ""
lYKYCSTA0.KYCSTAZRA1 = ""
lYKYCSTA0.KYCSTAZPCI = ""
lYKYCSTA0.KYCSTAYKYC = ""
lYKYCSTA0.KYCSTAZNAT = ""
lYKYCSTA0.KYCSTAZRSD = ""
      
lYKYCSTA0.KYCSTAYFCT = ""
lYKYCSTA0.KYCSTAYUSR = ""
lYKYCSTA0.KYCSTAYAMJ = 0
lYKYCSTA0.KYCSTAYHMS = 0
lYKYCSTA0.KYCSTAYVER = 0

lYKYCSTA0.KYCSTAYUSR = usrName_UCase
lYKYCSTA0.KYCSTAYAMJ = DSys
lYKYCSTA0.KYCSTAYHMS = time_Hms
lYKYCSTA0.KYCSTAYVER = 0
lYKYCSTA0.KYCSTAYFCT = " "

End Function






