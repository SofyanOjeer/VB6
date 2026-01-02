Attribute VB_Name = "srvYCREANO0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeYCREANO0
    CREANOETA       As Integer                        ' ETABLISSEMENT
    CREANOAGE       As Integer                        ' AGENCE OPERATRICE
    CREANOSER       As String                     ' SERVICE OPERATEUR
    CREANOSSE       As String                     ' S/SERVICE OPERATEUR
    CREANOOPE       As String                    ' CODE OPERATION
    CREANONUM       As Long                           ' NUMERO OPERATION
    CREANOEVE       As String                   ' EVENEMENT
    CREANOCRE       As Long                           ' NUMERO CRE
    CREANODCRE      As Long                           ' DATE CRE
    CREANOSTAK      As String                     '
    CREANOSPLF       As Integer
    CREANOLTXT       As Integer
    CREANONB       As Long
    CREANOCREC       As Long
    CREANODTRT      As Long                           ' DATE DE TRAITEMENT
    CREANOPIE       As Long                           ' NUMERO DE PIECE
    
End Type
Public Sub rsYCREANO0_Init(rsYCREANO0 As typeYCREANO0)
rsYCREANO0.CREANOETA = 0
rsYCREANO0.CREANOAGE = 0
rsYCREANO0.CREANOSER = ""
rsYCREANO0.CREANOSSE = ""
rsYCREANO0.CREANOOPE = ""
rsYCREANO0.CREANONUM = 0
rsYCREANO0.CREANOEVE = ""
rsYCREANO0.CREANOCRE = 0
rsYCREANO0.CREANODCRE = 0
rsYCREANO0.CREANOSTAK = ""
rsYCREANO0.CREANOSPLF = 0
rsYCREANO0.CREANOLTXT = 0
rsYCREANO0.CREANONB = 0
rsYCREANO0.CREANOCREC = 0
rsYCREANO0.CREANODTRT = 0
rsYCREANO0.CREANOPIE = 0

End Sub
Public Function sqlYCREANO0_Insert(newY As typeYCREANO0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYCREANO0_Insert = Null

xSet = " (CREANOETA,CREANOAGE,CREANOSER,CREANOSSE,CREANOOPE,CREANONUM,CREANOEVE,CREANOCRE"
xValues = " values(" & newY.CREANOETA & "," & newY.CREANOAGE & ",'" & newY.CREANOSER & "','" & newY.CREANOSSE & "'" _
        & ",'" & newY.CREANOOPE & "'," & newY.CREANONUM & ",'" & newY.CREANOEVE & "'," & newY.CREANOCRE

' Détecter les modifications
'===================================================================================
If Trim(newY.CREANOSTAK) <> "" Then xSet = xSet & ",CREANOSTAK": xValues = xValues & " ,'" & Replace(Trim(newY.CREANOSTAK), "'", "''") & "'"

If newY.CREANODCRE <> 0 Then xSet = xSet & ",CREANODCRE": xValues = xValues & " ," & newY.CREANODCRE
If newY.CREANOSPLF <> 0 Then xSet = xSet & ",CREANOSPLF": xValues = xValues & " ," & newY.CREANOSPLF
If newY.CREANOLTXT <> 0 Then xSet = xSet & ",CREANOLTXT": xValues = xValues & " ," & newY.CREANOLTXT
If newY.CREANONB <> 0 Then xSet = xSet & ",CREANONB": xValues = xValues & " ," & newY.CREANONB
If newY.CREANOCREC <> 0 Then xSet = xSet & ",CREANOCREC": xValues = xValues & " ," & newY.CREANOCREC
If newY.CREANODTRT <> 0 Then xSet = xSet & ",CREANODTRT": xValues = xValues & " ," & newY.CREANODTRT
If newY.CREANOPIE <> 0 Then xSet = xSet & ",CREANOPIE": xValues = xValues & " ," & newY.CREANOPIE


xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YCREANO0" & xSet & ")" & xValues & ")"

Set rsSab = cnsab.Execute(xSQL, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYCREANO0_Insert = "Erreur màj : " & newY.CREANOCRE
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYCREANO0_Insert = "sqlYCREANO0_Insert " & vbCrLf & Error
End Function


Public Function rsYCREANO0_GetBuffer(rsAdo As ADODB.Recordset, rsYCREANO0 As typeYCREANO0)
On Error GoTo Error_Handler
rsYCREANO0_GetBuffer = Null
rsYCREANO0.CREANOETA = rsAdo("CREANOETA")
rsYCREANO0.CREANOAGE = rsAdo("CREANOAGE")
rsYCREANO0.CREANOSER = rsAdo("CREANOSER")
rsYCREANO0.CREANOSSE = rsAdo("CREANOSSE")
rsYCREANO0.CREANOOPE = rsAdo("CREANOOPE")
rsYCREANO0.CREANONUM = rsAdo("CREANONUM")
rsYCREANO0.CREANOEVE = rsAdo("CREANOEVE")
rsYCREANO0.CREANOCRE = rsAdo("CREANOCRE")
rsYCREANO0.CREANODCRE = rsAdo("CREANODCRE")
rsYCREANO0.CREANOSTAK = rsAdo("CREANOSTAK")
rsYCREANO0.CREANOSPLF = rsAdo("CREANOSPLF")
rsYCREANO0.CREANOLTXT = rsAdo("CREANOLTXT")
rsYCREANO0.CREANONB = rsAdo("CREANONB")
rsYCREANO0.CREANOCREC = rsAdo("CREANOCREC")
rsYCREANO0.CREANODTRT = rsAdo("CREANODTRT")
rsYCREANO0.CREANOPIE = rsAdo("CREANOPIE")

Exit Function
Error_Handler:
rsYCREANO0_GetBuffer = Error
End Function

