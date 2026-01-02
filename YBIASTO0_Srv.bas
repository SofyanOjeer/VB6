Attribute VB_Name = "srvYBIASTO0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeYBIASTO0
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    YSTOETA          As String * 2
    YSTOAGE          As String * 2
    YSTOSER          As String * 2
    YSTOSSE          As String * 2
    YSTOOPE          As String * 3
    YSTONUM          As Long
    YSTOSEQ          As Long
    YSTOPCI          As String * 10
    YSTOCCL          As String * 1
    YSTOCLI          As Long
    YSTODEV          As String * 3
    YSTOMON          As Currency
    YSTODEB          As Long
    YSTOFIN          As Long
    YSTOAPP          As String * 3
    YSTONAT          As String * 6
    YSTOCC1          As String * 1
    YSTOCL1          As Long
    YSTOCC2          As String * 1
    YSTOCL2          As Long
    YSTOCTX          As String * 6
    YSTOTAU          As Double
    
End Type

'---------------------------------------------------------
Public Function srvYBIASTO0_GetBuffer_ODBC(rsADO As ADODB.Recordset, recYBIASTO0 As typeYBIASTO0)
'---------------------------------------------------------
On Error Resume Next 'GoTo Error_Handler
srvYBIASTO0_GetBuffer_ODBC = Null

recYBIASTO0.YSTOETA = rsADO("YSTOETA")
recYBIASTO0.YSTOAGE = rsADO("YSTOAGE")
recYBIASTO0.YSTOSER = rsADO("YSTOSER")
recYBIASTO0.YSTOSSE = rsADO("YSTOSSE")
recYBIASTO0.YSTOOPE = rsADO("YSTOOPE")
recYBIASTO0.YSTONUM = rsADO("YSTONUM")
recYBIASTO0.YSTOSEQ = rsADO("YSTOSEQ")
recYBIASTO0.YSTOPCI = rsADO("YSTOPCI")
recYBIASTO0.YSTOCCL = rsADO("YSTOCCL")
recYBIASTO0.YSTOCLI = rsADO("YSTOCLI")
recYBIASTO0.YSTODEV = rsADO("YSTODEV")
recYBIASTO0.YSTOMON = rsADO("YSTOMON")
recYBIASTO0.YSTODEB = rsADO("YSTODEB")
recYBIASTO0.YSTOFIN = rsADO("YSTOFIN")
recYBIASTO0.YSTOAPP = rsADO("YSTOAPP")
recYBIASTO0.YSTONAT = rsADO("YSTONAT")
recYBIASTO0.YSTOCC1 = rsADO("YSTOCC1")
recYBIASTO0.YSTOCL1 = rsADO("YSTOCL1")
recYBIASTO0.YSTOCC2 = rsADO("YSTOCC2")
recYBIASTO0.YSTOCL2 = rsADO("YSTOCL2")
recYBIASTO0.YSTOCTX = rsADO("YSTOCTX")
recYBIASTO0.YSTOTAU = rsADO("YSTOTAU")

Exit Function

Error_Handler:
srvYBIASTO0_GetBuffer_ODBC = Error

End Function


Public Sub srvYBIASTO0_ElpDisplay(lYBIASTO0 As typeYBIASTO0)
frmElpDisplay.fgData.Rows = 23
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "YSTOETA      2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = lYBIASTO0.YSTOETA
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "YSTOAGE      2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AGENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = lYBIASTO0.YSTOAGE
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "YSTOSER      2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = lYBIASTO0.YSTOSER
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "YSTOSSE      2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOUS SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = lYBIASTO0.YSTOSSE
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "YSTOOPE      3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE OPERATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = lYBIASTO0.YSTOOPE
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "YSTONUM      9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° OPERATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = lYBIASTO0.YSTONUM
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "YSTOSEQ      5N"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° SEQUENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = lYBIASTO0.YSTOSEQ
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "YSTOPCI     10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "RUBRIQUE COMPTABLE "
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = lYBIASTO0.YSTOPCI
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "YSTOCCL      1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CLIENT/TIERS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = lYBIASTO0.YSTOCCL
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "YSTOCLI      7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CLIENT PRINCIPAL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = lYBIASTO0.YSTOCLI

frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "YSTODEV      3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DEVISE CONTRAT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = lYBIASTO0.YSTODEV

frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "YSTOMON   18P2 "
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT DISPONIBLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = lYBIASTO0.YSTOMON

frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "YSTODEB      8P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DEBUT CONTRAT   "
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = lYBIASTO0.YSTODEB

frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "YSTOFIN      8P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "FIN CONTRAT     "
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = lYBIASTO0.YSTOFIN

frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "YSTOAPP      3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "APPLICATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = lYBIASTO0.YSTOAPP

frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "YSTONAT      6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE NATURE     "
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = lYBIASTO0.YSTONAT

frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "YSTOCC1      1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CLIENT/TIERS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = lYBIASTO0.YSTOCC1

frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "YSTOCL1      7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CLIENT N° 1     "
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = lYBIASTO0.YSTOCL1

frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "YSTOCC2      1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CLIENT/TIERS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = lYBIASTO0.YSTOCC2

frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "YSTOCL2      7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CLIENT N°2"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = lYBIASTO0.YSTOCL2

frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "YSTOCTX      6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE TAUX       "
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = lYBIASTO0.YSTOCTX

frmElpDisplay.fgData.Row = 22
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "YSTOTAU    14P9"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TAUX / MARGE    "
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = lYBIASTO0.YSTOTAU

frmElpDisplay.Show vbModal
End Sub


Public Sub recYBIASTO0_Init(lYBIASTO0 As typeYBIASTO0)
lYBIASTO0.YSTOETA = ""
lYBIASTO0.YSTOAGE = ""
lYBIASTO0.YSTOSER = ""
lYBIASTO0.YSTOSSE = ""
lYBIASTO0.YSTOOPE = ""
lYBIASTO0.YSTONUM = 0
lYBIASTO0.YSTOSEQ = 0
lYBIASTO0.YSTOPCI = ""
lYBIASTO0.YSTOCCL = ""
lYBIASTO0.YSTOCLI = 0
lYBIASTO0.YSTODEV = ""
lYBIASTO0.YSTOMON = 0
lYBIASTO0.YSTODEB = 0
lYBIASTO0.YSTOFIN = 0
lYBIASTO0.YSTOAPP = ""
lYBIASTO0.YSTONAT = ""
lYBIASTO0.YSTOCC1 = ""
lYBIASTO0.YSTOCL1 = 0
lYBIASTO0.YSTOCC2 = ""
lYBIASTO0.YSTOCL2 = 0
lYBIASTO0.YSTOCTX = ""
lYBIASTO0.YSTOTAU = 0

End Sub


Public Sub YBIASTO0_Sql_Douteux(lYBIACPT0 As typeYBIACPT0, cnADO As ADODB.Connection, rsADO As ADODB.Recordset)
Dim xYBIACPT0 As typeYBIACPT0
Dim xSql As String, X As String
Dim V

' Union Bank / Khalifa Bank
Set rsADO = Nothing

    X = "COMPTEOBL like '98150" _
        & "%' AND  COMPTEDEV =   '" & lYBIACPT0.COMPTEDEV _
        & "' AND CLIENACLI = '" & lYBIACPT0.CLIENACLI & "'"
        
    
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where " & X
    
        Set rsADO = cnADO.Execute(xSql)
        Do While Not rsADO.EOF
             V = srvYBIACPT0_GetBuffer_ODBC(rsADO, xYBIACPT0)
            If Not IsNull(V) Then
                MsgBox V, vbCritical, "frmSAB_Stock.fgSelect_Display"
            Else
               lYBIACPT0.SOLDECEN = lYBIACPT0.SOLDECEN + xYBIACPT0.SOLDECEN
            End If
           rsADO.MoveNext
        Loop


End Sub
Public Sub YBIASTO0_Sql(lWhere As String, lnbDossier As Long, larrYBIASTO0() As typeYBIASTO0, larrYBIACPT0() As typeYBIACPT0, larrCompte_Nb As Long, cnADO As ADODB.Connection, rsADO As ADODB.Recordset)
Dim xYBIASTO0 As typeYBIASTO0, xYBIACPT0 As typeYBIACPT0
Dim xSql As String
Dim blnOk As Boolean, blnCumul As Boolean
Dim V
Dim arrCompte_Max As Long
Dim I As Integer, X As String

ReDim larrYBIASTO0(101)
arrCompte_Max = 100: larrCompte_Nb = 0

blnOk = False
lnbDossier = 0

xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIASTO0" & lWhere & " order by YSTOPCI,YSTODEV,YSTOCLI,YSTONAT"
If blnJPL Then Exit Sub
Set rsADO = cnADO.Execute(xSql)
Do While Not rsADO.EOF
    lnbDossier = lnbDossier + 1
    V = srvYBIASTO0_GetBuffer_ODBC(rsADO, xYBIASTO0)
    If Not IsNull(V) Then
        MsgBox V, vbCritical, "frmSAB_Stock.cmdSelect_Ok_Click"
        Exit Sub
    Else
        If Not blnOk Then
            blnOk = True
            larrCompte_Nb = 1
            larrYBIASTO0(1) = xYBIASTO0
        Else
            blnCumul = False
            If xYBIASTO0.YSTOPCI = larrYBIASTO0(larrCompte_Nb).YSTOPCI _
            And xYBIASTO0.YSTODEV = larrYBIASTO0(larrCompte_Nb).YSTODEV _
            And xYBIASTO0.YSTOCLI = larrYBIASTO0(larrCompte_Nb).YSTOCLI Then
                blnCumul = True
                If xYBIASTO0.YSTOAPP = "DAT" Then
                    If xYBIASTO0.YSTONAT <> larrYBIASTO0(larrCompte_Nb).YSTONAT Then blnCumul = False
                End If
            End If
            If blnCumul Then
                larrYBIASTO0(larrCompte_Nb).YSTOMON = larrYBIASTO0(larrCompte_Nb).YSTOMON + xYBIASTO0.YSTOMON
        
            Else
                larrCompte_Nb = larrCompte_Nb + 1
                If larrCompte_Nb > arrCompte_Max Then
                    arrCompte_Max = arrCompte_Max + 100
                    ReDim Preserve larrYBIASTO0(arrCompte_Max)
                End If
                
                larrYBIASTO0(larrCompte_Nb) = xYBIASTO0
            End If
        End If
    End If
    rsADO.MoveNext
Loop

'=====================================================================================
ReDim larrYBIACPT0(arrCompte_Max)
Set rsADO = Nothing

For I = 1 To larrCompte_Nb
    xYBIASTO0 = larrYBIASTO0(I)
    X = "COMPTEOBL like '" & mId$(xYBIASTO0.YSTOPCI, 1, 5) _
        & "%' AND  COMPTEDEV =   '" & xYBIASTO0.YSTODEV _
        & "' AND CLIENACLI = '" & Format$(xYBIASTO0.YSTOCLI, "0000000") & "'"
        
'DAT : si NANTI le compte se termine par N, sinon par S (6 ème position de la nature)
    If xYBIASTO0.YSTOAPP = "DAT" Then
        X = X & " AND  COMPTECOM like  '%" & mId$(xYBIASTO0.YSTONAT, 6, 1) & " %'"
    End If
    
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where " & X
    
    blnOk = False
    If xYBIASTO0.YSTOCCL = " " Then
        Set rsADO = cnADO.Execute(xSql)
        Do While Not rsADO.EOF
             V = srvYBIACPT0_GetBuffer_ODBC(rsADO, xYBIACPT0)
            If Not IsNull(V) Then
                MsgBox V, vbCritical, "frmSAB_Stock.fgSelect_Display"
            Else
               If xYBIACPT0.COMPTEFON <> "4" Then blnOk = True: Exit Do
            End If
           rsADO.MoveNext
        Loop

    End If
    
    If Not blnOk Then
        recYBIACPT0_Init xYBIACPT0
        xYBIACPT0.COMPTEINT = "??" & xYBIASTO0.YSTOAPP & " " & xYBIASTO0.YSTOOPE & " " & xYBIASTO0.YSTOPCI & " " & xYBIASTO0.YSTOCCL & " " & xYBIASTO0.YSTOCLI
    End If
    
'Cas particulier clients douteux : 2 PCI :99901* et 98150* ( Union Bank , Khalifa bank)
    If mId$(xYBIACPT0.COMPTEOBL, 1, 5) = "99901" Then Call YBIASTO0_Sql_Douteux(xYBIACPT0, cnADO, rsADO)
    
    
    larrYBIACPT0(I) = xYBIACPT0
    
Next I


End Sub


