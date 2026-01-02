Attribute VB_Name = "rsYBIASTO0"
'---------------------------------------------------------
Option Explicit
Type typeYBIASTO0

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

Public Sub YBIASTO0_Sql(lWhere As String, lnbDossier As Long, larrYBIASTO0() As typeYBIASTO0, larrYBIACPT0() As typeYBIACPT0, larrCompte_Nb As Long)
Dim xYBIASTO0 As typeYBIASTO0, xYBIACPT0 As typeYBIACPT0, mYBIACPT0 As typeYBIACPT0
Dim xSQL As String
Dim blnOk As Boolean, blnCumul As Boolean
Dim V
Dim arrCompte_Max As Long
Dim I As Integer, X As String

ReDim larrYBIASTO0(101)
arrCompte_Max = 100: larrCompte_Nb = 0

blnOk = False
lnbDossier = 0

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIASTO0" & lWhere & " order by YSTOPCI,YSTODEV,YSTOCLI,YSTONAT"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    lnbDossier = lnbDossier + 1
    V = rsYBIASTO0_GetBuffer(rsSab, xYBIASTO0)
'$JPL 20130212
'--------------
    If Mid$(xYBIASTO0.YSTONAT, 1, 3) = "BDF" Then Mid$(xYBIASTO0.YSTONAT, 1, 3) = "GEN"
'--------------
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
    rsSab.MoveNext
Loop

'=====================================================================================
ReDim larrYBIACPT0(arrCompte_Max)
Set rsSab = Nothing

For I = 1 To larrCompte_Nb
    xYBIASTO0 = larrYBIASTO0(I)

    If xYBIASTO0.YSTOCCL = " " Then
        X = "COMPTEOBL like '" & Mid$(xYBIASTO0.YSTOPCI, 1, 5) _
            & "%' AND  COMPTEDEV =   '" & xYBIASTO0.YSTODEV _
            & "' AND CLIENACLI = '" & Format$(xYBIASTO0.YSTOCLI, "0000000") & "'"
     Else
        X = "COMPTEOBL like '" & Mid$(xYBIASTO0.YSTOPCI, 1, 5) _
            & "%' AND  COMPTEDEV =   '" & xYBIASTO0.YSTODEV _
            & "' AND COMPTECOM like '%" & Format$(xYBIASTO0.YSTOCLI, "00000") & "%'"
    End If
    
'DAT : si NANTI le compte se termine par N, sinon par S (6 ème position de la nature)
    Select Case xYBIASTO0.YSTOAPP
        Case "DAT": X = X & " AND  COMPTECOM like  '%" & Mid$(xYBIASTO0.YSTONAT, 6, 1) & " %'"
    End Select
    Select Case xYBIASTO0.YSTOOPE
        Case "RDE": X = "COMPTEOBL like '" & Mid$(xYBIASTO0.YSTOPCI, 1, 5) _
                        & "%' AND  COMPTEDEV =   '" & xYBIASTO0.YSTODEV _
                        & "' AND COMPTECOM like '%RDE%'"
        Case "RDI": X = "COMPTEOBL like '" & Mid$(xYBIASTO0.YSTOPCI, 1, 5) _
                        & "%' AND  COMPTEDEV =   '" & xYBIASTO0.YSTODEV _
                        & "' AND COMPTECOM like '%RDI%'"
        
    End Select
    
    
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where " & X
    
    blnOk = False
    'If xYBIASTO0.YSTOCCL = " " Then
        Set rsSab = cnsab.Execute(xSQL)
        Do While Not rsSab.EOF
             V = rsYBIACPT0_GetBuffer(rsSab, xYBIACPT0)
            If Not IsNull(V) Then
                MsgBox V, vbCritical, "frmSAB_Stock.fgSelect_Display"
            Else
               If xYBIACPT0.COMPTEFON <> "4" Then
                    If Not blnOk Then
                        blnOk = True
                        mYBIACPT0 = xYBIACPT0
                    Else
                        mYBIACPT0.SOLDECEN = mYBIACPT0.SOLDECEN + xYBIACPT0.SOLDECEN
                    End If
                End If
                
            End If
           rsSab.MoveNext
        Loop

    'End If
    
    If Not blnOk Then
        rsYBIACPT0_Init mYBIACPT0
        xSQL = "select CLIENARA1 from " & paramIBM_Library_SAB & ".ZCLIENA0 where CLIENACLI = '" & Format$(xYBIASTO0.YSTOCLI, "0000000") & "'"
        Set rsSab = cnsab.Execute(xSQL)
        If Not rsSab.EOF Then
            mYBIACPT0.CLIENARA1 = rsSab("CLIENARA1")
        Else
            mYBIACPT0.CLIENARA1 = "??? " & xYBIASTO0.YSTOCLI
        End If
        
        mYBIACPT0.COMPTEINT = "??" & xYBIASTO0.YSTOAPP & " " & xYBIASTO0.YSTOOPE & " " & xYBIASTO0.YSTOPCI & " " & xYBIASTO0.YSTOCCL & " " & xYBIASTO0.YSTOCLI
    End If
    
'Cas particulier clients douteux : 2 PCI :99901* et 98150* ( Union Bank , Khalifa bank)
    'If Mid$(xYBIACPT0.COMPTEOBL, 1, 5) = "99901" Then Call YBIASTO0_Sql_Douteux(xYBIACPT0)
    
    
    larrYBIACPT0(I) = mYBIACPT0
    
Next I


End Sub
Public Sub YBIASTO0_Sql_Douteux(lYBIACPT0 As typeYBIACPT0)
Dim xYBIACPT0 As typeYBIACPT0
Dim xSQL As String, X As String
Dim V

' Union Bank / Khalifa Bank
Set rsSab = Nothing

    X = "COMPTEOBL like '98150" _
        & "%' AND  COMPTEDEV =   '" & lYBIACPT0.COMPTEDEV _
        & "' AND CLIENACLI = '" & lYBIACPT0.CLIENACLI & "'"
        
    
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where " & X
    
        Set rsSab = cnsab.Execute(xSQL)
        Do While Not rsSab.EOF
             V = rsYBIACPT0_GetBuffer(rsSab, xYBIACPT0)
            If Not IsNull(V) Then
                MsgBox V, vbCritical, "frmSAB_Stock.fgSelect_Display"
            Else
               lYBIACPT0.SOLDECEN = lYBIACPT0.SOLDECEN + xYBIACPT0.SOLDECEN
            End If
           rsSab.MoveNext
        Loop


End Sub

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsYBIASTO0_GetBuffer(rsAdo As ADODB.Recordset, rsYBIASTO0 As typeYBIASTO0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYBIASTO0_GetBuffer = Null

rsYBIASTO0.YSTOETA = rsAdo("YSTOETA")
rsYBIASTO0.YSTOAGE = rsAdo("YSTOAGE")
rsYBIASTO0.YSTOSER = rsAdo("YSTOSER")
rsYBIASTO0.YSTOSSE = rsAdo("YSTOSSE")
rsYBIASTO0.YSTOOPE = rsAdo("YSTOOPE")
rsYBIASTO0.YSTONUM = rsAdo("YSTONUM")
rsYBIASTO0.YSTOSEQ = rsAdo("YSTOSEQ")
rsYBIASTO0.YSTOPCI = rsAdo("YSTOPCI")
rsYBIASTO0.YSTOCCL = rsAdo("YSTOCCL")
rsYBIASTO0.YSTOCLI = rsAdo("YSTOCLI")
rsYBIASTO0.YSTODEV = rsAdo("YSTODEV")
rsYBIASTO0.YSTOMON = rsAdo("YSTOMON")
rsYBIASTO0.YSTODEB = rsAdo("YSTODEB")
rsYBIASTO0.YSTOFIN = rsAdo("YSTOFIN")
rsYBIASTO0.YSTOAPP = rsAdo("YSTOAPP")
rsYBIASTO0.YSTONAT = rsAdo("YSTONAT")
rsYBIASTO0.YSTOCC1 = rsAdo("YSTOCC1")
rsYBIASTO0.YSTOCL1 = rsAdo("YSTOCL1")
rsYBIASTO0.YSTOCC2 = rsAdo("YSTOCC2")
rsYBIASTO0.YSTOCL2 = rsAdo("YSTOCL2")
rsYBIASTO0.YSTOCTX = rsAdo("YSTOCTX")
rsYBIASTO0.YSTOTAU = rsAdo("YSTOTAU")

Exit Function

Error_Handler:

rsYBIASTO0_GetBuffer = Error

End Function

'---------------------------------------------------------
Public Sub rsYBIASTO0_Init(rsYBIASTO0 As typeYBIASTO0)
'---------------------------------------------------------

End Sub


'









