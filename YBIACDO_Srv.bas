Attribute VB_Name = "srvYBIACDO"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsSab As New ADODB.Recordset


Type typeYBIACDOCOM0
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    CDODOSCOP       As String * 3                     ' CODE OPERATION
    CDODOSDOS       As Long                           ' NUMERO DOSSIER
    
    CDOCOMCOM       As String * 6                     ' CODE COMMISSION
    CDOCOMMON       As Currency                       ' MONTANT COMMISSION
    CDOCOMDEV       As String * 3                     ' DEVISE COMMISSION
    CDOCOMMTV       As Currency                       ' MONTANT TVA

    CDOCO2TX1       As Double                         ' Taux tranche 1
    CDOCO2PER       As String * 1                     ' Périodicité

    CDOTC2DEV       As String * 3                     ' Devise
    CDOTC2MTF       As Currency                       ' Montant fixe


End Type
    
Type typeYBIACDO
    YCDODOS0()        As typeZCDODOS0
    YCDODOS0_Nb   As Integer
    YCDOMOD0()        As typeZCDOMOD0
    YCDOMOD0_Nb   As Integer
    YCDOTIE0()        As typeZCDOTIE0
    YCDOTIE0_Nb   As Integer
    YCDOCOM0()        As typeZCDOCOM0
    YCDOCOM0_Nb   As Integer
    YCDOCO20()        As typeZCDOCO20
    YCDOCO20_Nb   As Integer
    YCDOTC20()        As typeZCDOTC20
    YCDOTC20_Nb   As Integer
    YCDODES0()        As typeZCDODES0
    YCDODES0_Nb   As Integer
    YCDOSWI0()        As typeZCDOSWI0
    YCDOSWI0_Nb   As Integer
    YCDOIRR0()        As typeZCDOIRR0
    YCDOIRR0_Nb   As Integer
    YCDOUTI0()        As typeZCDOUTI0
    YCDOUTI0_Nb   As Integer
    YCDOREG0()      As typeZCDOREG0
    YCDOREG0_Nb   As Integer
End Type

Public Sub recYBIACDOCOM0_Init(lYBIACDOCOM0 As typeYBIACDOCOM0)
lYBIACDOCOM0.CDODOSCOP = ""    'As String * 3                     ' CODE OPERATION
lYBIACDOCOM0.CDODOSDOS = 0     'As Long                           ' NUMERO DOSSIER
    
lYBIACDOCOM0.CDOCOMCOM = ""    'As String * 6                     ' CODE COMMISSION
lYBIACDOCOM0.CDOCOMMON = 0     'As Currency                       ' MONTANT COMMISSION
lYBIACDOCOM0.CDOCOMDEV = ""    'As String * 3                     ' DEVISE COMMISSION
lYBIACDOCOM0.CDOCOMMTV = 0     'As Currency                       ' MONTANT TVA

lYBIACDOCOM0.CDOCO2TX1 = 0      'As Double                         ' Taux tranche 1
lYBIACDOCOM0.CDOCO2PER = ""     'As String                         ' Périodicité

lYBIACDOCOM0.CDOTC2DEV = ""     'As String * 3                     ' Devise
lYBIACDOCOM0.CDOTC2MTF = 0      'As Currency                       ' Montant fixe

End Sub

Public Sub srvYBIACDOCOM0_Load(lYBIACDO As typeYBIACDO, lcnf As typeYBIACDOCOM0, lnot As typeYBIACDOCOM0)
Dim mK1 As String, X As String
Dim blnCOM As Boolean, blnCNF As Boolean, blnNOT As Boolean
Dim xYBIACDOCOM0 As typeYBIACDOCOM0
Dim K As Integer, K2 As Integer

blnCNF = False
blnNOT = False

recYBIACDOCOM0_Init lcnf: lnot = lcnf

For K = 1 To lYBIACDO.YCDOCOM0_Nb
    blnCOM = False

    If lYBIACDO.YCDOCOM0(K).CDOCOMSPE = 1 And lYBIACDO.YCDOCOM0(K).CDOCOMCOM = "ECNF  " Then blnCOM = True
    If lYBIACDO.YCDOCOM0(K).CDOCOMSPE = 999 And lYBIACDO.YCDOCOM0(K).CDOCOMCOM = "ENOTIF" Then blnCOM = True
    
    If blnCOM Then
        recYBIACDOCOM0_Init xYBIACDOCOM0
        xYBIACDOCOM0.CDODOSCOP = lYBIACDO.YCDODOS0(1).CDODOSCOP
        xYBIACDOCOM0.CDODOSDOS = lYBIACDO.YCDODOS0(1).CDODOSDOS
        
        xYBIACDOCOM0.CDOCOMCOM = lYBIACDO.YCDOCOM0(K).CDOCOMCOM
        xYBIACDOCOM0.CDOCOMMON = lYBIACDO.YCDOCOM0(K).CDOCOMMON
        xYBIACDOCOM0.CDOCOMMTV = lYBIACDO.YCDOCOM0(K).CDOCOMMTV
        xYBIACDOCOM0.CDOCOMDEV = lYBIACDO.YCDOCOM0(K).CDOCOMDEV
        
        For K2 = 1 To lYBIACDO.YCDOCO20_Nb
            If lYBIACDO.YCDOCOM0(K).CDOCOMCOP = lYBIACDO.YCDOCO20(K2).CDOCO2COP _
            And lYBIACDO.YCDOCOM0(K).CDOCOMDOS = lYBIACDO.YCDOCO20(K2).CDOCO2DOS _
            And lYBIACDO.YCDOCOM0(K).CDOCOMNUR = lYBIACDO.YCDOCO20(K2).CDOCO2NUR _
            And lYBIACDO.YCDOCOM0(K).CDOCOMUTI = lYBIACDO.YCDOCO20(K2).CDOCO2UTI _
            And lYBIACDO.YCDOCOM0(K).CDOCOMEVE = lYBIACDO.YCDOCO20(K2).CDOCO2EVE _
            And lYBIACDO.YCDOCOM0(K).CDOCOMSEQ = lYBIACDO.YCDOCO20(K2).CDOCO2SEQ _
            And lYBIACDO.YCDOCOM0(K).CDOCOMSPE = lYBIACDO.YCDOCO20(K2).CDOCO2SPE _
           Then
                xYBIACDOCOM0.CDOCO2TX1 = lYBIACDO.YCDOCO20(K2).CDOCO2TX1
                xYBIACDOCOM0.CDOCO2PER = lYBIACDO.YCDOCO20(K2).CDOCO2PER
            End If
        Next K2
        
        For K2 = 1 To lYBIACDO.YCDOTC20_Nb
            If lYBIACDO.YCDOCOM0(K).CDOCOMCOP = lYBIACDO.YCDOTC20(K2).CDOTC2COP _
            And lYBIACDO.YCDOCOM0(K).CDOCOMDOS = lYBIACDO.YCDOTC20(K2).CDOTC2DOS _
            And lYBIACDO.YCDOCOM0(K).CDOCOMNUR = lYBIACDO.YCDOTC20(K2).CDOTC2NUR _
            And lYBIACDO.YCDOCOM0(K).CDOCOMUTI = lYBIACDO.YCDOTC20(K2).CDOTC2UTI _
            And lYBIACDO.YCDOCOM0(K).CDOCOMEVE = lYBIACDO.YCDOTC20(K2).CDOTC2EVE _
            And lYBIACDO.YCDOCOM0(K).CDOCOMSEQ = lYBIACDO.YCDOTC20(K2).CDOTC2SEQ _
           Then
                xYBIACDOCOM0.CDOTC2DEV = lYBIACDO.YCDOTC20(K2).CDOTC2DEV
                xYBIACDOCOM0.CDOTC2MTF = lYBIACDO.YCDOTC20(K2).CDOTC2MTF
            End If
        Next K2
        
        
        Select Case lYBIACDO.YCDOCOM0(K).CDOCOMCOM
            Case "ECNF  ":
                            If Not blnCNF Then lcnf = xYBIACDOCOM0: blnCNF = True
             Case "ENOTIF":
                            If Not blnNOT Then lnot = xYBIACDOCOM0: blnNOT = True
       End Select
        
    End If
            
Next K

End Sub



Public Function srvYBIACDO_ODBC(lYBIACDO As typeYBIACDO)
Dim xSQL As String
Dim V
On Error GoTo Error_Handler


srvYBIACDO_ODBC = Null

'Lecture Dossier
'===============
Set rsSab = Nothing
xSQL = "select * from " & paramIBM_Library_SAB & ".ZCDODOS0 where CDODOSCOP = 'CDE' and CDODOSDOS = " & lYBIACDO.YCDODOS0(1).CDODOSDOS
'$2003.11.04 rsSab.Open xSQL, paramODBC_DSN_SAB
    Set rsSab = cnsab.Execute(xSQL)  '$2003.11.04
If rsSab.EOF Then
    V = "Dossier inconnu"
Else
    V = rsZCDODOS0_GetBuffer(rsSab, lYBIACDO.YCDODOS0(1))
End If

If Not IsNull(V) Then
    srvYBIACDO_ODBC = "Lecture ZCDODOS0 : " & V
    Exit Function
Else
    lYBIACDO.YCDODOS0_Nb = 1
End If

'Lecture Modifications / dossier
'=================================
Set rsSab = Nothing
xSQL = "select * from " & paramIBM_Library_SAB & ".ZCDOMOD0 where  CDOMODCOP = 'CDE' and CDOMODDOS = " & lYBIACDO.YCDODOS0(1).CDODOSDOS
'$2003.11.04   rsSab.Open xSQL, paramODBC_DSN_SAB
    Set rsSab = cnsab.Execute(xSQL)  '$2003.11.04
Do While Not rsSab.EOF
    lYBIACDO.YCDOMOD0_Nb = lYBIACDO.YCDOMOD0_Nb + 1
    If lYBIACDO.YCDOMOD0_Nb > 5 Then ReDim Preserve lYBIACDO.YCDOMOD0(lYBIACDO.YCDOMOD0_Nb)
    V = rsZCDOMOD0_GetBuffer(rsSab, lYBIACDO.YCDOMOD0(lYBIACDO.YCDOMOD0_Nb))
    If Not IsNull(V) Then
        srvYBIACDO_ODBC = "Lecture ZCDOMOD0 : " & V
        Exit Function
    End If
    rsSab.MoveNext
Loop


'Lecture Bénéficiaire
'===============
If lYBIACDO.YCDODOS0(1).CDODOSBER = "T" Then
    Set rsSab = Nothing
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZCDOTIE0 where CDOTIETIE = '" & lYBIACDO.YCDODOS0(1).CDODOSBEN & "'"
    '$2003.11.04   rsSab.Open xSQL, paramODBC_DSN_SAB
        Set rsSab = cnsab.Execute(xSQL)  '$2003.11.04
    If rsSab.EOF Then
        V = "Tiers inconnu  : " & lYBIACDO.YCDODOS0(1).CDODOSBEN
    Else
        V = rsZCDOTIE0_GetBuffer(rsSab, lYBIACDO.YCDOTIE0(1))
    End If
    
    If Not IsNull(V) Then
        srvYBIACDO_ODBC = "Lecture ZCDOTIE0 : " & V
        Exit Function
    Else
        lYBIACDO.YCDOTIE0_Nb = 1
    End If
End If

'Lecture Réglements / dossier
'=================================

Set rsSab = Nothing
xSQL = "select * from " & paramIBM_Library_SAB & ".ZCDOREG0 where  CDOREGCOP = 'CDE' and CDOREGDOS = " & lYBIACDO.YCDODOS0(1).CDODOSDOS
'$2003.11.04   rsSab.Open xSQL, paramODBC_DSN_SAB
    Set rsSab = cnsab.Execute(xSQL)  '$2003.11.04
Do While Not rsSab.EOF
    lYBIACDO.YCDOREG0_Nb = lYBIACDO.YCDOREG0_Nb + 1
    If lYBIACDO.YCDOREG0_Nb > 5 Then ReDim Preserve lYBIACDO.YCDOREG0(lYBIACDO.YCDOREG0_Nb)
    V = rsZCDOREG0_GetBuffer(rsSab, lYBIACDO.YCDOREG0(lYBIACDO.YCDOREG0_Nb))
    If Not IsNull(V) Then
        srvYBIACDO_ODBC = "Lecture ZCDOREG0 : " & V
        Exit Function
    End If
    rsSab.MoveNext
Loop


'Lecture Utilisations / dossier
'=================================
Set rsSab = Nothing
xSQL = "select * from " & paramIBM_Library_SAB & ".ZCDOUTI0 where  CDOUTICOP = 'CDE' and CDOUTIDOS = " & lYBIACDO.YCDODOS0(1).CDODOSDOS
'$2003.11.04   rsSab.Open xSQL, paramODBC_DSN_SAB
    Set rsSab = cnsab.Execute(xSQL)  '$2003.11.04
Do While Not rsSab.EOF
    lYBIACDO.YCDOUTI0_Nb = lYBIACDO.YCDOUTI0_Nb + 1
    If lYBIACDO.YCDOUTI0_Nb > 5 Then ReDim Preserve lYBIACDO.YCDOUTI0(lYBIACDO.YCDOUTI0_Nb)
    V = rsZCDOUTI0_GetBuffer(rsSab, lYBIACDO.YCDOUTI0(lYBIACDO.YCDOUTI0_Nb))
    If Not IsNull(V) Then
        srvYBIACDO_ODBC = "Lecture ZCDOUTI0 : " & V
        Exit Function
    End If
    rsSab.MoveNext
Loop



'Lecture Description / dossier
'=================================
Set rsSab = Nothing
xSQL = "select * from " & paramIBM_Library_SAB & ".ZCDODES0 where  CDODESCOP = 'CDE' and CDODESDOS = " & lYBIACDO.YCDODOS0(1).CDODOSDOS
'$2003.11.04   rsSab.Open xSQL, paramODBC_DSN_SAB
    Set rsSab = cnsab.Execute(xSQL)  '$2003.11.04
Do While Not rsSab.EOF
    lYBIACDO.YCDODES0_Nb = lYBIACDO.YCDODES0_Nb + 1
    If lYBIACDO.YCDODES0_Nb > 5 Then ReDim Preserve lYBIACDO.YCDODES0(lYBIACDO.YCDODES0_Nb)
    V = rsZCDODES0_GetBuffer(rsSab, lYBIACDO.YCDODES0(lYBIACDO.YCDODES0_Nb))
    If Not IsNull(V) Then
        srvYBIACDO_ODBC = "Lecture ZCDODES0 : " & V
        Exit Function
    End If
    rsSab.MoveNext
Loop


'Lecture SWIFT / dossier
'=================================
Set rsSab = Nothing
xSQL = "select * from " & paramIBM_Library_SAB & ".ZCDOSWI0 where  CDOSWICOP = 'CDE' and CDOSWIDOS = " & lYBIACDO.YCDODOS0(1).CDODOSDOS
'$2003.11.04   rsSab.Open xSQL, paramODBC_DSN_SAB
    Set rsSab = cnsab.Execute(xSQL)  '$2003.11.04
Do While Not rsSab.EOF
    lYBIACDO.YCDOSWI0_Nb = lYBIACDO.YCDOSWI0_Nb + 1
    If lYBIACDO.YCDOSWI0_Nb > 5 Then ReDim Preserve lYBIACDO.YCDOSWI0(lYBIACDO.YCDOSWI0_Nb)
    V = rsZCDOSWI0_GetBuffer(rsSab, lYBIACDO.YCDOSWI0(lYBIACDO.YCDOSWI0_Nb))
    If Not IsNull(V) Then
        srvYBIACDO_ODBC = "Lecture ZCDOSWI0 : " & V
        Exit Function
    End If
    rsSab.MoveNext
Loop


'Lecture Irrégularités / dossier
'=================================
Set rsSab = Nothing
xSQL = "select * from " & paramIBM_Library_SAB & ".ZCDOIRR0 where  CDOIRRCOP = 'CDE' and CDOIRRDOS = " & lYBIACDO.YCDODOS0(1).CDODOSDOS
'$2003.11.04   rsSab.Open xSQL, paramODBC_DSN_SAB
    Set rsSab = cnsab.Execute(xSQL)  '$2003.11.04
Do While Not rsSab.EOF
    lYBIACDO.YCDOIRR0_Nb = lYBIACDO.YCDOIRR0_Nb + 1
    If lYBIACDO.YCDOIRR0_Nb > 5 Then ReDim Preserve lYBIACDO.YCDOIRR0(lYBIACDO.YCDOIRR0_Nb)
    V = rsZCDOIRR0_GetBuffer(rsSab, lYBIACDO.YCDOIRR0(lYBIACDO.YCDOIRR0_Nb))
    If Not IsNull(V) Then
        srvYBIACDO_ODBC = "Lecture ZCDOIRR0 : " & V
        Exit Function
    End If
    rsSab.MoveNext
Loop


'Lecture Commisssions / dossier
'=================================
Set rsSab = Nothing
xSQL = "select * from " & paramIBM_Library_SAB & ".ZCDOCOM0 where  CDOCOMCOP = 'CDE' and CDOCOMDOS = " & lYBIACDO.YCDODOS0(1).CDODOSDOS
'$2003.11.04   rsSab.Open xSQL, paramODBC_DSN_SAB
    Set rsSab = cnsab.Execute(xSQL)  '$2003.11.04
Do While Not rsSab.EOF
    lYBIACDO.YCDOCOM0_Nb = lYBIACDO.YCDOCOM0_Nb + 1
    If lYBIACDO.YCDOCOM0_Nb > 5 Then ReDim Preserve lYBIACDO.YCDOCOM0(lYBIACDO.YCDOCOM0_Nb)
    V = rsZCDOCOM0_GetBuffer(rsSab, lYBIACDO.YCDOCOM0(lYBIACDO.YCDOCOM0_Nb))
    If Not IsNull(V) Then
        srvYBIACDO_ODBC = "Lecture ZCDOCOM0 : " & V
        Exit Function
    End If
    rsSab.MoveNext
Loop

'Lecture ZCDOCO2 / dossier
'=================================
Set rsSab = Nothing
xSQL = "select * from " & paramIBM_Library_SAB & ".ZCDOCO20 where  CDOCO2COP = 'CDE' and CDOCO2DOS = " & lYBIACDO.YCDODOS0(1).CDODOSDOS
'$2003.11.04   rsSab.Open xSQL, paramODBC_DSN_SAB
    Set rsSab = cnsab.Execute(xSQL)  '$2003.11.04
Do While Not rsSab.EOF
    lYBIACDO.YCDOCO20_Nb = lYBIACDO.YCDOCO20_Nb + 1
    If lYBIACDO.YCDOCO20_Nb > 5 Then ReDim Preserve lYBIACDO.YCDOCO20(lYBIACDO.YCDOCO20_Nb)
    V = rsZCDOCO20_GetBuffer(rsSab, lYBIACDO.YCDOCO20(lYBIACDO.YCDOCO20_Nb))
    If Not IsNull(V) Then
        srvYBIACDO_ODBC = "Lecture ZCDOCO20 : " & V
        Exit Function
    End If
    rsSab.MoveNext
Loop


'Lecture ZCDOTC2 / dossier
'=================================
Set rsSab = Nothing
xSQL = "select * from " & paramIBM_Library_SAB & ".ZCDOTC20 where  CDOTC2COP = 'CDE' and  CDOTC2DOS = " & lYBIACDO.YCDODOS0(1).CDODOSDOS
'$2003.11.04   rsSab.Open xSQL, paramODBC_DSN_SAB
    Set rsSab = cnsab.Execute(xSQL)  '$2003.11.04
Do While Not rsSab.EOF
    lYBIACDO.YCDOTC20_Nb = lYBIACDO.YCDOTC20_Nb + 1
    If lYBIACDO.YCDOTC20_Nb > 5 Then ReDim Preserve lYBIACDO.YCDOTC20(lYBIACDO.YCDOTC20_Nb)
    V = rsZCDOTC20_GetBuffer(rsSab, lYBIACDO.YCDOTC20(lYBIACDO.YCDOTC20_Nb))
    If Not IsNull(V) Then
        srvYBIACDO_ODBC = "Lecture ZCDOTC20 : " & V
        Exit Function
    End If
    rsSab.MoveNext
Loop


rsSab.Close

Exit Function

Error_Handler:
srvYBIACDO_ODBC = Error
End Function
