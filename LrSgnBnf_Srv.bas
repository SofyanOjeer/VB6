Attribute VB_Name = "srvLrsGNbNF"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recLrSgnBnfLen = 534 ' 34 + 500
Type typeLrSgnBnf
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10

    CDBANQ                  As String * 5
    CDDECL                  As String * 5
    RFBENF                  As String * 16
    NSIREN                  As String * 9
    NBDF1                   As String * 13
    AMJ1                    As String * 8
    NBDF2                   As String * 13
    AMJ2                    As String * 8
    NOMBNF                  As String * 60
    PRENOM                  As String * 60
    CDSEXE                  As String * 1
    JMA3                    As String * 6
    CDPAYS1                 As String * 3
    CDDEPT1                 As String * 2
    CDCOMM1                 As String * 3
    LBCOMM1                 As String * 32
    NOMCJT                  As String * 60
    CDACCO                  As String * 5
    CTJURI                  As String * 5
    CDRESI                  As String * 1
    NOVOIE                  As String * 32
    CDPOST                  As String * 5
    LBCOMM2                 As String * 32
    CDDEPT2                 As String * 2
    CDPAYS2                 As String * 3
    CDTRI1                  As String * 16
    CDTRI2                  As String * 16
    CDHABI                  As String * 10
    AMJ4                    As String * 8
    HMSC                    As String * 8
    CDAGCO                  As String * 5
    CDPHMO                  As String * 1
    CDCRMD                  As String * 1
    INDSIR                  As String * 1
    FILL01                  As String * 11
    CDRESI1                 As String * 1
    CDACEN1                 As String * 4
    CTJURN1                 As String * 4
    CDPAYN1                 As String * 2
    CDSEXE1                 As String * 1
    CDPAYN2                 As String * 2
    CDRESI2                 As String * 1
    CDACCO2                 As String * 5
    FILL02                  As String * 14
 
   
  End Type
    
Public arrLrSgnBnf() As typeLrSgnBnf
Public arrLrSgnBnfNb As Integer
Public arrLrSgnBnfNbMax As Integer
Public arrLrSgnBnfIndex As Integer
Public arrLrSgnBnfSuite As Boolean
Public Function Scan(recLrSgnBnf As typeLrSgnBnf) As Integer
Scan = -1
For arrLrSgnBnfIndex = 1 To arrLrSgnBnfNb
    If arrLrSgnBnf(arrLrSgnBnfIndex).Method <> constDelete _
    And arrLrSgnBnf(arrLrSgnBnfIndex).Method <> constIgnore Then
        If arrLrSgnBnf(arrLrSgnBnfIndex).RFBENF = recLrSgnBnf.RFBENF Then
            Scan = arrLrSgnBnfIndex
            Exit For
        End If
    End If
Next arrLrSgnBnfIndex

End Function

'-----------------------------------------------------
Public Function Monitor(recLrSgnBnf As typeLrSgnBnf)
'-----------------------------------------------------

arrLrSgnBnfSuite = False
Select Case Mid$(Trim(recLrSgnBnf.Method), 1, 4)
    Case "Seek":            Monitor = SeekX(recLrSgnBnf)
    Case "Snap", "Prev":    Monitor = Snap(recLrSgnBnf)
    Case Else:              recLrSgnBnf.Err = recLrSgnBnf.Method
                            Call ErrorX(recLrSgnBnf)
                            Monitor = recLrSgnBnf.Err
End Select

End Function

'-----------------------------------------------------
Sub ErrorX(recLrSgnBnf As typeLrSgnBnf)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "LrSgnBnf: "

Select Case Mid$(recLrSgnBnf.Err, 9, 2)
    Case "22"
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recLrSgnBnf.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : .bas  ( " _
                & Trim(recLrSgnBnf.obj) & " : " & Trim(recLrSgnBnf.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function GetBuffer(recLrSgnBnf As typeLrSgnBnf)
'---------------------------------------------------------
Dim K As Integer, I As Integer
GetBuffer = Null
recLrSgnBnf.obj = Mid$(MsgTxt, MsgTxtIndex + 1, 12)
recLrSgnBnf.Method = Mid$(MsgTxt, MsgTxtIndex + 13, 12)
recLrSgnBnf.Err = Mid$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recLrSgnBnf.Err = Space$(10) Then
    recLrSgnBnf.CDBANQ = Mid$(MsgTxt, K + 1, 5)
    recLrSgnBnf.CDDECL = Mid$(MsgTxt, K + 6, 5)
    recLrSgnBnf.RFBENF = Mid$(MsgTxt, K + 11, 16)
    recLrSgnBnf.NSIREN = Mid$(MsgTxt, K + 27, 9)
    recLrSgnBnf.NBDF1 = Mid$(MsgTxt, K + 36, 13)
    recLrSgnBnf.AMJ1 = Mid$(MsgTxt, K + 49, 8)
    recLrSgnBnf.NBDF2 = Mid$(MsgTxt, K + 57, 13)
    recLrSgnBnf.AMJ2 = Mid$(MsgTxt, K + 70, 8)
    recLrSgnBnf.NOMBNF = Mid$(MsgTxt, K + 78, 60)
    recLrSgnBnf.PRENOM = Mid$(MsgTxt, K + 138, 60)
    recLrSgnBnf.CDSEXE = Mid$(MsgTxt, K + 198, 1)
    recLrSgnBnf.JMA3 = Mid$(MsgTxt, K + 199, 6)
    recLrSgnBnf.CDPAYS1 = Mid$(MsgTxt, K + 205, 3)
    recLrSgnBnf.CDDEPT1 = Mid$(MsgTxt, K + 208, 2)
    recLrSgnBnf.CDCOMM1 = Mid$(MsgTxt, K + 210, 3)
    recLrSgnBnf.LBCOMM1 = Mid$(MsgTxt, K + 213, 32)
    recLrSgnBnf.NOMCJT = Mid$(MsgTxt, K + 245, 60)
    recLrSgnBnf.CDACCO = Mid$(MsgTxt, K + 305, 5)
    recLrSgnBnf.CTJURI = Mid$(MsgTxt, K + 310, 5)
    recLrSgnBnf.CDRESI = Mid$(MsgTxt, K + 315, 1)
    recLrSgnBnf.NOVOIE = Mid$(MsgTxt, K + 316, 32)
    recLrSgnBnf.CDPOST = Mid$(MsgTxt, K + 348, 5)
    recLrSgnBnf.LBCOMM2 = Mid$(MsgTxt, K + 353, 32)
    recLrSgnBnf.CDDEPT2 = Mid$(MsgTxt, K + 385, 2)
    recLrSgnBnf.CDPAYS2 = Mid$(MsgTxt, K + 387, 3)
    recLrSgnBnf.CDTRI1 = Mid$(MsgTxt, K + 390, 16)
    recLrSgnBnf.CDTRI2 = Mid$(MsgTxt, K + 406, 16)
    recLrSgnBnf.CDHABI = Mid$(MsgTxt, K + 422, 10)
    recLrSgnBnf.AMJ4 = Mid$(MsgTxt, K + 432, 8)
    recLrSgnBnf.HMSC = Mid$(MsgTxt, K + 440, 8)
    recLrSgnBnf.CDAGCO = Mid$(MsgTxt, K + 448, 5)
    recLrSgnBnf.CDPHMO = Mid$(MsgTxt, K + 453, 1)
    recLrSgnBnf.CDCRMD = Mid$(MsgTxt, K + 454, 1)
    recLrSgnBnf.INDSIR = Mid$(MsgTxt, K + 455, 1)
    recLrSgnBnf.FILL01 = Mid$(MsgTxt, K + 456, 11)
    recLrSgnBnf.CDRESI1 = Mid$(MsgTxt, K + 467, 1)
    recLrSgnBnf.CDACEN1 = Mid$(MsgTxt, K + 468, 4)
    recLrSgnBnf.CTJURN1 = Mid$(MsgTxt, K + 472, 4)
    recLrSgnBnf.CDPAYN1 = Mid$(MsgTxt, K + 476, 2)
    recLrSgnBnf.CDSEXE1 = Mid$(MsgTxt, K + 478, 1)
    recLrSgnBnf.CDPAYN2 = Mid$(MsgTxt, K + 479, 2)
    recLrSgnBnf.CDRESI2 = Mid$(MsgTxt, K + 481, 1)
    recLrSgnBnf.CDACCO2 = Mid$(MsgTxt, K + 482, 5)
    recLrSgnBnf.FILL02 = Mid$(MsgTxt, K + 487, 14)

Else
    GetBuffer = recLrSgnBnf.Err
End If

MsgTxtIndex = MsgTxtIndex + recLrSgnBnfLen

End Function

'---------------------------------------------------------
Private Sub PutBuffer(recLrSgnBnf As typeLrSgnBnf)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recLrSgnBnf.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recLrSgnBnf.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 5) = recLrSgnBnf.CDBANQ
Mid$(MsgTxt, K + 6, 5) = recLrSgnBnf.CDDECL
Mid$(MsgTxt, K + 11, 16) = recLrSgnBnf.RFBENF

MsgTxtLen = MsgTxtLen + recLrSgnBnfLen
End Sub



'---------------------------------------------------------
Private Function SeekX(recLrSgnBnf As typeLrSgnBnf)
'---------------------------------------------------------

SeekX = "?"
MsgTxtLen = 0
Call PutBuffer(recLrSgnBnf)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(GetBuffer(recLrSgnBnf)) Then
        SeekX = Null
    Else
        Call ErrorX(recLrSgnBnf)
    End If
End If

End Function

'---------------------------------------------------------
Private Function Snap(recLrSgnBnf As typeLrSgnBnf)
'---------------------------------------------------------
Dim I As Integer
Snap = "?"
MsgTxtLen = 0
Call PutBuffer(recLrSgnBnf)
Call PutBuffer(arrLrSgnBnf(0))
If IsNull(SndRcv()) Then
    Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(GetBuffer(recLrSgnBnf)) Then
            Call srvLrsGNbNF.AddItem(recLrSgnBnf)
            arrLrSgnBnfSuite = True
        Else
            arrLrSgnBnfSuite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub Init(recLrSgnBnf As typeLrSgnBnf)
'---------------------------------------------------------
MsgTxt = Space$(recLrSgnBnfLen)
MsgTxtIndex = 0
Call GetBuffer(recLrSgnBnf)
recLrSgnBnf.obj = "SRVLRSGNBN"
End Sub

'---------------------------------------------------------
Public Sub AddItem(recLrSgnBnf As typeLrSgnBnf)
'---------------------------------------------------------
          
arrLrSgnBnfNb = arrLrSgnBnfNb + 1
    
If arrLrSgnBnfNb > arrLrSgnBnfNbMax Then
    arrLrSgnBnfNbMax = arrLrSgnBnfNbMax + 50
    ReDim Preserve arrLrSgnBnf(arrLrSgnBnfNbMax)
End If
recLrSgnBnf.Method = ""
arrLrSgnBnfIndex = arrLrSgnBnfNb
arrLrSgnBnf(arrLrSgnBnfIndex) = recLrSgnBnf
End Sub

