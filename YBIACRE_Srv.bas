Attribute VB_Name = "srvYBIACRE"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------

Type typeYBIACRE

    mCREDOS         As Long
    mCREPRE         As Long
    prtNb           As Integer
    mCREEMPNCL      As String
    
    Contact         As String
    Annexe_Nb       As Integer
    
    CRE_ZADRESS0 As typeZADRESS0
    CRE_Adresse_Concat As String
    
    ZCREDOS0()      As typeZCREDOS0
    ZCREDOS0_Nb     As Integer
    ZCREDOS0_Index  As Integer
    
    ZCREPRE0()      As typeZCREPRE0
    ZCREPRE0_Nb     As Integer
    ZCREPRE0_Index  As Integer
    
    ZCREPLA0()      As typeZCREPLA0
    ZCREPLA0_Nb     As Integer
    ZCREPLA0_Index  As Integer

    ZCREEMP0()      As typeZCREEMP0
    ZCREEMP0_Nb     As Integer
    ZCREEMP0_Index  As Integer

    ZCREEVE0()      As typeZCREEVE0
    ZCREEVE0_Nb     As Integer
    ZCREEVE0_Index  As Integer


    ZCREAVI0()      As typeZCREAVI0
    ZCREAVI0_Nb     As Integer
    ZCREAVI0_Index  As Integer

    ZCREBIS0()      As typeZCREBIS0
    ZCREBIS0_Nb     As Integer
    ZCREBIS0_Index  As Integer

End Type

Public Function srvYBIACRE_GetBuffer(lYBIACRE As typeYBIACRE)
Dim xSQL As String
Dim V
On Error GoTo Error_Handler

srvYBIACRE_GetBuffer = Null
lYBIACRE.mCREEMPNCL = ""

ReDim lYBIACRE.ZCREDOS0(1): lYBIACRE.ZCREDOS0_Nb = 0: lYBIACRE.ZCREDOS0_Index = 0
ReDim lYBIACRE.ZCREPRE0(1): lYBIACRE.ZCREPRE0_Nb = 0: lYBIACRE.ZCREPRE0_Index = 0
ReDim lYBIACRE.ZCREPLA0(1): lYBIACRE.ZCREPLA0_Nb = 0: lYBIACRE.ZCREPLA0_Index = 0
ReDim lYBIACRE.ZCREEVE0(50): lYBIACRE.ZCREEVE0_Nb = 0: lYBIACRE.ZCREEVE0_Index = 0
ReDim lYBIACRE.ZCREAVI0(50): lYBIACRE.ZCREAVI0_Nb = 0: lYBIACRE.ZCREAVI0_Index = 0
ReDim lYBIACRE.ZCREBIS0(50): lYBIACRE.ZCREBIS0_Nb = 0: lYBIACRE.ZCREBIS0_Index = 0
ReDim lYBIACRE.ZCREEMP0(50): lYBIACRE.ZCREEMP0_Nb = 0: lYBIACRE.ZCREEMP0_Index = 0

rsZADRESS0_Init lYBIACRE.CRE_ZADRESS0
'Lecture Dossier
'===============
Set rsSab = Nothing
xSQL = "select * from " & paramIBM_Library_SAB & ".ZCREDOS0 where CREDOSDOS = " & lYBIACRE.mCREDOS
    Set rsSab = cnsab.Execute(xSQL)
If rsSab.EOF Then
    V = "Dossier inconnu"
Else
    V = rsZCREDOS0_GetBuffer(rsSab, lYBIACRE.ZCREDOS0(1))
End If

If Not IsNull(V) Then
    srvYBIACRE_GetBuffer = "Lecture ZCREDOS0 : " & V
    Exit Function
Else
    lYBIACRE.ZCREDOS0_Nb = 1
End If



'Lecture Prêt
'=================================
Set rsSab = Nothing
xSQL = "select * from " & paramIBM_Library_SAB & ".ZCREPRE0 where CREPREDOS = " & lYBIACRE.mCREDOS & " AND  CREPREPRE = " & lYBIACRE.mCREPRE
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    lYBIACRE.ZCREPRE0_Nb = lYBIACRE.ZCREPRE0_Nb + 1
    If lYBIACRE.ZCREPRE0_Nb > 1 Then ReDim Preserve lYBIACRE.ZCREPRE0(lYBIACRE.ZCREPRE0_Nb)
    V = rsZCREPRE0_GetBuffer(rsSab, lYBIACRE.ZCREPRE0(lYBIACRE.ZCREPRE0_Nb))
    If Not IsNull(V) Then
        srvYBIACRE_GetBuffer = "Lecture ZCREPRE0 : " & V
        Exit Function
    End If
    rsSab.MoveNext
Loop



'Lecture Plan / Prêt
'=================================
Set rsSab = Nothing
xSQL = "select * from " & paramIBM_Library_SAB & ".ZCREPLA0 where CREPLADOS = " & lYBIACRE.mCREDOS & " AND  CREPLAPRE = " & lYBIACRE.mCREPRE
Set rsSab = cnsab.Execute(xSQL)  '$2003.11.04
Do While Not rsSab.EOF
    lYBIACRE.ZCREPLA0_Nb = lYBIACRE.ZCREPLA0_Nb + 1
    If lYBIACRE.ZCREPLA0_Nb > 1 Then ReDim Preserve lYBIACRE.ZCREPLA0(lYBIACRE.ZCREPLA0_Nb)
    V = rsZCREPLA0_GetBuffer(rsSab, lYBIACRE.ZCREPLA0(lYBIACRE.ZCREPLA0_Nb))
    If Not IsNull(V) Then
        srvYBIACRE_GetBuffer = "Lecture ZCREPLA0 : " & V
        Exit Function
    End If
    rsSab.MoveNext
Loop


'Lecture EMP0
'=================================
Set rsSab = Nothing
xSQL = "select * from " & paramIBM_Library_SAB & ".ZCREEMP0 where CREEMPDOS = " & lYBIACRE.mCREDOS
Set rsSab = cnsab.Execute(xSQL)  '$2003.11.04
Do While Not rsSab.EOF
    lYBIACRE.ZCREEMP0_Nb = lYBIACRE.ZCREEMP0_Nb + 1
    If lYBIACRE.ZCREEMP0_Nb > 50 Then ReDim Preserve lYBIACRE.ZCREEMP0(lYBIACRE.ZCREEMP0_Nb)
    V = rsZCREEMP0_GetBuffer(rsSab, lYBIACRE.ZCREEMP0(lYBIACRE.ZCREEMP0_Nb))
    If Not IsNull(V) Then
        srvYBIACRE_GetBuffer = "Lecture ZCREEMP0 : " & V
        Exit Function
    Else
        If lYBIACRE.ZCREEMP0(lYBIACRE.ZCREEMP0_Nb).CREEMPSEQ = 1 Then lYBIACRE.mCREEMPNCL = lYBIACRE.ZCREEMP0(lYBIACRE.ZCREEMP0_Nb).CREEMPNCL
    End If
    rsSab.MoveNext
Loop



'Lecture Evénement / Prêt
'=================================
'$jpl 24.06.2004
'Set rsSab = Nothing
'xSQL = "select * from " & paramIBM_Library_SAB & ".ZCREEVE0 where CREEVEDOS = " & lYBIACRE.mCREDOS & " AND  CREEVEPRE = " & lYBIACRE.mCREPRE
'Set rsSab = cnsab.Execute(xSQL)  '$2003.11.04
'Do While Not rsSab.EOF
'    lYBIACRE.ZCREEVE0_Nb = lYBIACRE.ZCREEVE0_Nb + 1
'    If lYBIACRE.ZCREEVE0_Nb > 50 Then ReDim Preserve lYBIACRE.ZCREEVE0(lYBIACRE.ZCREEVE0_Nb)
'    v=rsZCREEVE0_getbuffer(rsSab, lYBIACRE.ZCREEVE0(lYBIACRE.ZCREEVE0_Nb))
'    If Not IsNull(V) Then
'        srvYBIACRE_GetBuffer = "Lecture ZCREEVE0 : " & V
'        Exit Function
'    End If
'    rsSab.MoveNext
'Loop


'Lecture AVI0
'=================================
Set rsSab = Nothing
xSQL = "select * from " & paramIBM_Library_SAB & ".ZCREAVI0 where CREAVIDOS = " & lYBIACRE.mCREDOS & " AND  CREAVIPRE = " & lYBIACRE.mCREPRE
Set rsSab = cnsab.Execute(xSQL)  '$2003.11.04
Do While Not rsSab.EOF
    lYBIACRE.ZCREAVI0_Nb = lYBIACRE.ZCREAVI0_Nb + 1
    If lYBIACRE.ZCREAVI0_Nb > 50 Then ReDim Preserve lYBIACRE.ZCREAVI0(lYBIACRE.ZCREAVI0_Nb)
    V = rsZCREAVI0_GetBuffer(rsSab, lYBIACRE.ZCREAVI0(lYBIACRE.ZCREAVI0_Nb))
    If Not IsNull(V) Then
        srvYBIACRE_GetBuffer = "Lecture ZCREAVI0 : " & V
        Exit Function
    End If
    rsSab.MoveNext
Loop


'Lecture BIS0  !!!!!!!!!!!!!!!! FICHIER AS400 MEMBRE : BIS0001 : requête via ALIAS
'$ CREATE ALIAS SAB073T/ZCREBIS0_BIS0001 FOR SAB073T/ZCREBIS0 (BIS0001)
'$ create alias sab073/zcrebis0_bis0001 for sab073/zcrebis0(bis0001)
'============================================================================
'A revoir JPL 2004.04.16
        Set rsSab = Nothing
        xSQL = "select * from " & paramIBM_Library_SAB & ".ZCREBIS0_BIS0001 where CREBISDOS = " & lYBIACRE.mCREDOS & " AND  CREBISPRE = " & lYBIACRE.mCREPRE
        Set rsSab = cnsab.Execute(xSQL)  '$2003.11.04
        Do While Not rsSab.EOF
            lYBIACRE.ZCREBIS0_Nb = lYBIACRE.ZCREBIS0_Nb + 1
            If lYBIACRE.ZCREBIS0_Nb > 50 Then ReDim Preserve lYBIACRE.ZCREBIS0(lYBIACRE.ZCREBIS0_Nb)
            V = rsZCREBIS0_GetBuffer(rsSab, lYBIACRE.ZCREBIS0(lYBIACRE.ZCREBIS0_Nb))
            If Not IsNull(V) Then
                srvYBIACRE_GetBuffer = "Lecture ZCREBIS0 : " & V
                Exit Function
            End If
            rsSab.MoveNext
        Loop

'Lecture Evénement / Prêt
'=================================
lYBIACRE.CRE_ZADRESS0.ADRESSRA1 = "srvYBIACRE_GetBuffer : ADRESSE à faire"
lYBIACRE.ZCREDOS0_Index = 1
lYBIACRE.ZCREPRE0_Index = 1

lYBIACRE.CRE_ZADRESS0.ADRESSNUM = lYBIACRE.mCREEMPNCL
lYBIACRE.CRE_ZADRESS0.ADRESSCOA = "CO"

Call rsZADRESS0_Client(lYBIACRE.CRE_ZADRESS0)

Exit Function

Error_Handler:
srvYBIACRE_GetBuffer = Error
End Function

Public Function srvCREEVETYP_Lib(lCREEVETYP As String) As String
Select Case lCREEVETYP
    Case "00": srvCREEVETYP_Lib = "Mise à disposition"
    Case "01": srvCREEVETYP_Lib = "Int intercalaires"
    Case "02": srvCREEVETYP_Lib = "Echéance (Capital + intérêts)"
    Case "03": srvCREEVETYP_Lib = "Echéance d'intérêts"
    Case "04": srvCREEVETYP_Lib = "Echéance de capital"
    Case "05": srvCREEVETYP_Lib = "Appel de fonds au co-part"
    Case "06": srvCREEVETYP_Lib = "reversements de fons au co-part"
    Case "07": srvCREEVETYP_Lib = "Commission non cumulables"
    Case "08": srvCREEVETYP_Lib = "Commission sur coparticipant"
    Case "09": srvCREEVETYP_Lib = "Assurance non cumilable"
    Case "10": srvCREEVETYP_Lib = "Commission cumulable"
    Case "11": srvCREEVETYP_Lib = "Assurance cumulable"
    Case "12": srvCREEVETYP_Lib = "Int courus"
    Case "RP": srvCREEVETYP_Lib = "Rbt anticipé partiel"
    Case "RT": srvCREEVETYP_Lib = "Rbt anticipé total"
    Case Else: srvCREEVETYP_Lib = ""
End Select

End Function

