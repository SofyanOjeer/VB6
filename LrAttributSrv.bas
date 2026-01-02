Attribute VB_Name = "srvLrAttribut"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recLrAttributLen = 230 ' 34 + 196

Type typeLrAttribut
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    Nature                  As String * 1
    Référence               As String * 11
'attributs Luca Report
    AFFPU                   As String * 1
    AGEMT                   As String * 3
    AGENT                   As String * 3
    APPAR                   As String * 1
    AREFR                   As String * 1
    ATTCF                   As String * 1
    AUTDV                   As String * 1
    BONIF                   As String * 1
    CAROB                   As String * 1
    CATET                   As String * 2
    CDRES                   As String * 1
    CDZON                   As String * 1
    CLCRC                   As String * 1
    COTIT                   As String * 1
    CPEMS                   As String * 1
    CRDIV                   As String * 1
    CREIM                   As String * 1
    CREOR                   As String * 5
    CRETC                   As String * 1
    CRHYP                   As String * 1
    DCTOM                   As String * 1
    DRAC                    As String * 1
    DURIN                   As String * 1
    DUROM                   As String * 1
    DVOPR                   As String * 1
    ECART                   As String * 1
    ECFIN                   As String * 1
    ELIGB                   As String * 1
    FAMDV                   As String * 2
    FOPIF                   As String * 2
    FPRBG                   As String * 1
    GARCF                   As String * 1
    MLFCE                   As String * 1
    MONDV                   As String * 1
    MUTFG                   As String * 1
    NACGA                   As String * 1
    NACGR                   As String * 1
    NACPS                   As String * 1
    NAEGA                   As String * 1
    NAIMO                   As String * 5
    NAOCB                   As String * 4
    NAPRO                   As String * 1
    NARCP                   As String * 1
    NATCP                   As String * 1
    NATCR                   As String * 2
    NATCS                   As String * 3
    NATDD                   As String * 1
    NATER                   As String * 3
    NATIF                   As String * 2
    NATIT                   As String * 3
    NATMA                   As String * 1
    NATOF                   As String * 2
    NATRS                   As String * 2
    NRAST                   As String * 1
    NREHB                   As String * 1
    OPCVM                   As String * 1
    OPEFC                   As String * 1
    OPFDH                   As String * 1
    OPREC                   As String * 1
    PAACT                   As String * 2
    PERIO                   As String * 1
    PRIMP                   As String * 1
    PROCB                   As String * 1
    REDES                   As String * 1
    REDHB                   As String * 1
    RESET                   As String * 1
    REZON                   As String * 1
    RISPA                   As String * 1
    SENOP                   As String * 1
    TCFPE                   As String * 1
    TOPIF                   As String * 1
    TYCGR                   As String * 1
    TYCOM                   As String * 1
    TYDSU                   As String * 1
    TYETS                   As String * 1
    TYPOR                   As String * 3
    TYPSU                   As String * 1
    TYRES                   As String * 1
    ZACTI                   As String * 1
    ZAGDT                   As String * 1
'attributs Luca Risques
    CDCPCO                  As String * 1
    CDCPJO                  As String * 1
    CDCPFU                  As String * 15
    CDAGCO                  As String * 5
    CDREME                  As String * 1
    TYMTDV                  As String * 2
    TYVENT                  As String * 1
    CRVENT                  As String * 15
    CDDURE                  As String * 1
    DUINIT                  As String * 3
    CDCRTI                  As String * 1
    CDCRAC                  As String * 1
    CDBIOR                  As String * 1
    CDDEIN                  As String * 1
    CDCRIM                  As String * 1
    CDCRCO                  As String * 1
    CDCREF                  As String * 1
    CDLODA                  As String * 1
    CDCRET                  As String * 1
    CDOMPO                  As String * 1
    CDOPIM                  As String * 1
    CDSWAP                  As String * 1
 'attributs Réescompte
    REESC1                  As String * 8
    REESC6                  As String * 8
   
  End Type
    
Public arrLrAttribut() As typeLrAttribut
Public arrLrAttributNb As Integer
Public arrLrAttributNbMax As Integer
Public arrLrAttributIndex As Integer
Public arrLrAttributSuite As Boolean
'-----------------------------------------------------
Function Update(recLrAttribut As typeLrAttribut)
'-----------------------------------------------------

Update = "?"

MsgTxtLen = 0
Call PutBuffer(recLrAttribut)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    recLrAttribut.Err = Mid$(MsgTxt, MsgTxtIndex + 25, 10)
    If Trim(recLrAttribut.Err) <> "" Then
        Call ErrorX(recLrAttribut)
        Update = recLrAttribut.Err
        Exit Function
    Else
        Update = Null
    End If
Else
    recLrAttribut.Err = "srv"
End If


'=====================================================
End Function

'-----------------------------------------------------
Public Function Monitor(recLrAttribut As typeLrAttribut)
'-----------------------------------------------------

arrLrAttributSuite = False
Select Case Mid$(Trim(recLrAttribut.Method), 1, 4)
    Case "Seek"
                Monitor = SeekX(recLrAttribut)
    Case "Snap", "Prev"
              Monitor = Snap(recLrAttribut)
    Case Else
                recLrAttribut.Err = recLrAttribut.Method
                Call ErrorX(recLrAttribut)
                Monitor = recLrAttribut.Err
End Select

End Function

'-----------------------------------------------------
Sub ErrorX(recLrAttribut As typeLrAttribut)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Lr Attributs: "

Select Case Mid$(recLrAttribut.Err, 9, 2)
    Case "22"
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recLrAttribut.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : .bas  ( " _
                & Trim(recLrAttribut.obj) & " : " & Trim(recLrAttribut.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function GetBuffer(recLrAttribut As typeLrAttribut)
'---------------------------------------------------------
Dim K As Integer, I As Integer
GetBuffer = Null
recLrAttribut.obj = Mid$(MsgTxt, MsgTxtIndex + 1, 12)
recLrAttribut.Method = Mid$(MsgTxt, MsgTxtIndex + 13, 12)
recLrAttribut.Err = Mid$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recLrAttribut.Err = Space$(10) Then
    recLrAttribut.Nature = Mid$(MsgTxt, K + 1, 1)
     recLrAttribut.Référence = Mid$(MsgTxt, K + 2, 11)
'attributs Luca Report
    recLrAttribut.AFFPU = Mid$(MsgTxt, K + 13, 1)
    recLrAttribut.AGEMT = Mid$(MsgTxt, K + 14, 3)
    recLrAttribut.AGENT = Mid$(MsgTxt, K + 17, 3)
    recLrAttribut.APPAR = Mid$(MsgTxt, K + 20, 1)
    recLrAttribut.AREFR = Mid$(MsgTxt, K + 21, 1)
    recLrAttribut.ATTCF = Mid$(MsgTxt, K + 22, 1)
    recLrAttribut.AUTDV = Mid$(MsgTxt, K + 23, 1)
    recLrAttribut.BONIF = Mid$(MsgTxt, K + 24, 1)
    recLrAttribut.CAROB = Mid$(MsgTxt, K + 25, 1)
    recLrAttribut.CATET = Mid$(MsgTxt, K + 26, 2)
    recLrAttribut.CDRES = Mid$(MsgTxt, K + 28, 1)
    recLrAttribut.CDZON = Mid$(MsgTxt, K + 29, 1)
    recLrAttribut.CLCRC = Mid$(MsgTxt, K + 30, 1)
    recLrAttribut.COTIT = Mid$(MsgTxt, K + 31, 1)
    recLrAttribut.CPEMS = Mid$(MsgTxt, K + 32, 1)
    recLrAttribut.CRDIV = Mid$(MsgTxt, K + 33, 1)
    recLrAttribut.CREIM = Mid$(MsgTxt, K + 34, 1)
    recLrAttribut.CREOR = Mid$(MsgTxt, K + 35, 5)
    recLrAttribut.CRETC = Mid$(MsgTxt, K + 40, 1)
    recLrAttribut.CRHYP = Mid$(MsgTxt, K + 41, 1)
    recLrAttribut.DCTOM = Mid$(MsgTxt, K + 42, 1)
    recLrAttribut.DRAC = Mid$(MsgTxt, K + 43, 1)
    recLrAttribut.DURIN = Mid$(MsgTxt, K + 44, 1)
    recLrAttribut.DUROM = Mid$(MsgTxt, K + 45, 1)
    recLrAttribut.DVOPR = Mid$(MsgTxt, K + 46, 1)
    recLrAttribut.ECART = Mid$(MsgTxt, K + 47, 1)
    recLrAttribut.ECFIN = Mid$(MsgTxt, K + 48, 1)
    recLrAttribut.ELIGB = Mid$(MsgTxt, K + 49, 1)
    recLrAttribut.FAMDV = Mid$(MsgTxt, K + 50, 2)
    recLrAttribut.FOPIF = Mid$(MsgTxt, K + 52, 2)
    recLrAttribut.FPRBG = Mid$(MsgTxt, K + 54, 1)
    recLrAttribut.GARCF = Mid$(MsgTxt, K + 55, 1)
    recLrAttribut.MLFCE = Mid$(MsgTxt, K + 56, 1)
    recLrAttribut.MONDV = Mid$(MsgTxt, K + 57, 1)
    recLrAttribut.MUTFG = Mid$(MsgTxt, K + 58, 1)
    recLrAttribut.NACGA = Mid$(MsgTxt, K + 59, 1)
    recLrAttribut.NACGR = Mid$(MsgTxt, K + 60, 1)
    recLrAttribut.NACPS = Mid$(MsgTxt, K + 61, 1)
    recLrAttribut.NAEGA = Mid$(MsgTxt, K + 62, 1)
    recLrAttribut.NAIMO = Mid$(MsgTxt, K + 63, 5)
    recLrAttribut.NAOCB = Mid$(MsgTxt, K + 68, 4)
    recLrAttribut.NAPRO = Mid$(MsgTxt, K + 72, 1)
    recLrAttribut.NARCP = Mid$(MsgTxt, K + 73, 1)
    recLrAttribut.NATCP = Mid$(MsgTxt, K + 74, 1)
    recLrAttribut.NATCR = Mid$(MsgTxt, K + 75, 2)
    recLrAttribut.NATCS = Mid$(MsgTxt, K + 77, 3)
    recLrAttribut.NATDD = Mid$(MsgTxt, K + 80, 1)
    recLrAttribut.NATER = Mid$(MsgTxt, K + 81, 3)
    recLrAttribut.NATIF = Mid$(MsgTxt, K + 84, 2)
    recLrAttribut.NATIT = Mid$(MsgTxt, K + 86, 3)
    recLrAttribut.NATMA = Mid$(MsgTxt, K + 89, 1)
    recLrAttribut.NATOF = Mid$(MsgTxt, K + 90, 2)
    recLrAttribut.NATRS = Mid$(MsgTxt, K + 92, 2)
    recLrAttribut.NRAST = Mid$(MsgTxt, K + 94, 1)
    recLrAttribut.NREHB = Mid$(MsgTxt, K + 95, 1)
    recLrAttribut.OPCVM = Mid$(MsgTxt, K + 96, 1)
    recLrAttribut.OPEFC = Mid$(MsgTxt, K + 97, 1)
    recLrAttribut.OPFDH = Mid$(MsgTxt, K + 98, 1)
    recLrAttribut.OPREC = Mid$(MsgTxt, K + 99, 1)
    recLrAttribut.PAACT = Mid$(MsgTxt, K + 100, 2)
    recLrAttribut.PERIO = Mid$(MsgTxt, K + 102, 1)
    recLrAttribut.PRIMP = Mid$(MsgTxt, K + 103, 1)
    recLrAttribut.PROCB = Mid$(MsgTxt, K + 104, 1)
    recLrAttribut.REDES = Mid$(MsgTxt, K + 105, 1)
    recLrAttribut.REDHB = Mid$(MsgTxt, K + 106, 1)
    recLrAttribut.RESET = Mid$(MsgTxt, K + 107, 1)
    recLrAttribut.REZON = Mid$(MsgTxt, K + 108, 1)
    recLrAttribut.RISPA = Mid$(MsgTxt, K + 109, 1)
    recLrAttribut.SENOP = Mid$(MsgTxt, K + 110, 1)
    recLrAttribut.TCFPE = Mid$(MsgTxt, K + 111, 1)
    recLrAttribut.TOPIF = Mid$(MsgTxt, K + 112, 1)
    recLrAttribut.TYCGR = Mid$(MsgTxt, K + 113, 1)
    recLrAttribut.TYCOM = Mid$(MsgTxt, K + 114, 1)
    recLrAttribut.TYDSU = Mid$(MsgTxt, K + 115, 1)
    recLrAttribut.TYETS = Mid$(MsgTxt, K + 116, 1)
    recLrAttribut.TYPOR = Mid$(MsgTxt, K + 117, 3)
    recLrAttribut.TYPSU = Mid$(MsgTxt, K + 120, 1)
    recLrAttribut.TYRES = Mid$(MsgTxt, K + 121, 1)
    recLrAttribut.ZACTI = Mid$(MsgTxt, K + 122, 1)
    recLrAttribut.ZAGDT = Mid$(MsgTxt, K + 123, 1)
'attributs Luca Risques
    recLrAttribut.CDCPCO = Mid$(MsgTxt, K + 124, 1)
    recLrAttribut.CDCPJO = Mid$(MsgTxt, K + 125, 1)
    recLrAttribut.CDCPFU = Mid$(MsgTxt, K + 126, 15)
    recLrAttribut.CDAGCO = Mid$(MsgTxt, K + 141, 5)
    recLrAttribut.CDREME = Mid$(MsgTxt, K + 146, 1)
    recLrAttribut.TYMTDV = Mid$(MsgTxt, K + 147, 2)
    recLrAttribut.TYVENT = Mid$(MsgTxt, K + 149, 1)
    recLrAttribut.CRVENT = Mid$(MsgTxt, K + 150, 15)
    recLrAttribut.CDDURE = Mid$(MsgTxt, K + 165, 1)
    recLrAttribut.DUINIT = Mid$(MsgTxt, K + 166, 3)
    recLrAttribut.CDCRTI = Mid$(MsgTxt, K + 169, 1)
    recLrAttribut.CDCRAC = Mid$(MsgTxt, K + 170, 1)
    recLrAttribut.CDBIOR = Mid$(MsgTxt, K + 171, 1)
    recLrAttribut.CDDEIN = Mid$(MsgTxt, K + 172, 1)
    recLrAttribut.CDCRIM = Mid$(MsgTxt, K + 173, 1)
    recLrAttribut.CDCRCO = Mid$(MsgTxt, K + 174, 1)
    recLrAttribut.CDCREF = Mid$(MsgTxt, K + 175, 1)
    recLrAttribut.CDLODA = Mid$(MsgTxt, K + 176, 1)
    recLrAttribut.CDCRET = Mid$(MsgTxt, K + 177, 1)
    recLrAttribut.CDOMPO = Mid$(MsgTxt, K + 178, 1)
    recLrAttribut.CDOPIM = Mid$(MsgTxt, K + 179, 1)
    recLrAttribut.CDSWAP = Mid$(MsgTxt, K + 180, 1)
'attributs Réescompte
    recLrAttribut.REESC1 = Mid$(MsgTxt, K + 181, 8)
    recLrAttribut.REESC6 = Mid$(MsgTxt, K + 189, 8)

Else
    GetBuffer = recLrAttribut.Err
End If

MsgTxtIndex = MsgTxtIndex + recLrAttributLen

End Function

'---------------------------------------------------------
Private Sub PutBuffer(recLrAttribut As typeLrAttribut)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recLrAttribut.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recLrAttribut.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 1) = recLrAttribut.Nature
Mid$(MsgTxt, K + 2, 11) = recLrAttribut.Référence
Mid$(MsgTxt, K + 13, 1) = recLrAttribut.AFFPU
Mid$(MsgTxt, K + 14, 3) = recLrAttribut.AGEMT
Mid$(MsgTxt, K + 17, 3) = recLrAttribut.AGENT
Mid$(MsgTxt, K + 20, 1) = recLrAttribut.APPAR
Mid$(MsgTxt, K + 21, 1) = recLrAttribut.AREFR
Mid$(MsgTxt, K + 22, 1) = recLrAttribut.ATTCF
Mid$(MsgTxt, K + 23, 1) = recLrAttribut.AUTDV
Mid$(MsgTxt, K + 24, 1) = recLrAttribut.BONIF
Mid$(MsgTxt, K + 25, 1) = recLrAttribut.CAROB
Mid$(MsgTxt, K + 26, 2) = recLrAttribut.CATET
Mid$(MsgTxt, K + 28, 1) = recLrAttribut.CDRES
Mid$(MsgTxt, K + 29, 1) = recLrAttribut.CDZON
Mid$(MsgTxt, K + 30, 1) = recLrAttribut.CLCRC
Mid$(MsgTxt, K + 31, 1) = recLrAttribut.COTIT
Mid$(MsgTxt, K + 32, 1) = recLrAttribut.CPEMS
Mid$(MsgTxt, K + 33, 1) = recLrAttribut.CRDIV
Mid$(MsgTxt, K + 34, 1) = recLrAttribut.CREIM
Mid$(MsgTxt, K + 35, 5) = recLrAttribut.CREOR
Mid$(MsgTxt, K + 40, 1) = recLrAttribut.CRETC
Mid$(MsgTxt, K + 41, 1) = recLrAttribut.CRHYP
Mid$(MsgTxt, K + 42, 1) = recLrAttribut.DCTOM
Mid$(MsgTxt, K + 43, 1) = recLrAttribut.DRAC
Mid$(MsgTxt, K + 44, 1) = recLrAttribut.DURIN
Mid$(MsgTxt, K + 45, 1) = recLrAttribut.DUROM
Mid$(MsgTxt, K + 46, 1) = recLrAttribut.DVOPR
Mid$(MsgTxt, K + 47, 1) = recLrAttribut.ECART
Mid$(MsgTxt, K + 48, 1) = recLrAttribut.ECFIN
Mid$(MsgTxt, K + 49, 1) = recLrAttribut.ELIGB
Mid$(MsgTxt, K + 50, 2) = recLrAttribut.FAMDV
Mid$(MsgTxt, K + 52, 2) = recLrAttribut.FOPIF
Mid$(MsgTxt, K + 54, 1) = recLrAttribut.FPRBG
Mid$(MsgTxt, K + 55, 1) = recLrAttribut.GARCF
Mid$(MsgTxt, K + 56, 1) = recLrAttribut.MLFCE
Mid$(MsgTxt, K + 57, 1) = recLrAttribut.MONDV
Mid$(MsgTxt, K + 58, 1) = recLrAttribut.MUTFG
Mid$(MsgTxt, K + 59, 1) = recLrAttribut.NACGA
Mid$(MsgTxt, K + 60, 1) = recLrAttribut.NACGR
Mid$(MsgTxt, K + 61, 1) = recLrAttribut.NACPS
Mid$(MsgTxt, K + 62, 1) = recLrAttribut.NAEGA
Mid$(MsgTxt, K + 63, 5) = recLrAttribut.NAIMO
Mid$(MsgTxt, K + 68, 4) = recLrAttribut.NAOCB
Mid$(MsgTxt, K + 72, 1) = recLrAttribut.NAPRO
Mid$(MsgTxt, K + 73, 1) = recLrAttribut.NARCP
Mid$(MsgTxt, K + 74, 1) = recLrAttribut.NATCP
Mid$(MsgTxt, K + 75, 2) = recLrAttribut.NATCR
Mid$(MsgTxt, K + 77, 3) = recLrAttribut.NATCS
Mid$(MsgTxt, K + 80, 1) = recLrAttribut.NATDD
Mid$(MsgTxt, K + 81, 3) = recLrAttribut.NATER
Mid$(MsgTxt, K + 84, 2) = recLrAttribut.NATIF
Mid$(MsgTxt, K + 86, 3) = recLrAttribut.NATIT
Mid$(MsgTxt, K + 89, 1) = recLrAttribut.NATMA
Mid$(MsgTxt, K + 90, 2) = recLrAttribut.NATOF
Mid$(MsgTxt, K + 92, 2) = recLrAttribut.NATRS
Mid$(MsgTxt, K + 94, 1) = recLrAttribut.NRAST
Mid$(MsgTxt, K + 95, 1) = recLrAttribut.NREHB
Mid$(MsgTxt, K + 96, 1) = recLrAttribut.OPCVM
Mid$(MsgTxt, K + 97, 1) = recLrAttribut.OPEFC
Mid$(MsgTxt, K + 98, 1) = recLrAttribut.OPFDH
Mid$(MsgTxt, K + 99, 1) = recLrAttribut.OPREC
Mid$(MsgTxt, K + 100, 2) = recLrAttribut.PAACT
Mid$(MsgTxt, K + 102, 1) = recLrAttribut.PERIO
Mid$(MsgTxt, K + 103, 1) = recLrAttribut.PRIMP
Mid$(MsgTxt, K + 104, 1) = recLrAttribut.PROCB
Mid$(MsgTxt, K + 105, 1) = recLrAttribut.REDES
Mid$(MsgTxt, K + 106, 1) = recLrAttribut.REDHB
Mid$(MsgTxt, K + 107, 1) = recLrAttribut.RESET
Mid$(MsgTxt, K + 108, 1) = recLrAttribut.REZON
Mid$(MsgTxt, K + 109, 1) = recLrAttribut.RISPA
Mid$(MsgTxt, K + 110, 1) = recLrAttribut.SENOP
Mid$(MsgTxt, K + 111, 1) = recLrAttribut.TCFPE
Mid$(MsgTxt, K + 112, 1) = recLrAttribut.TOPIF
Mid$(MsgTxt, K + 113, 1) = recLrAttribut.TYCGR
Mid$(MsgTxt, K + 114, 1) = recLrAttribut.TYCOM
Mid$(MsgTxt, K + 115, 1) = recLrAttribut.TYDSU
Mid$(MsgTxt, K + 116, 1) = recLrAttribut.TYETS
Mid$(MsgTxt, K + 117, 3) = recLrAttribut.TYPOR
Mid$(MsgTxt, K + 120, 1) = recLrAttribut.TYPSU
Mid$(MsgTxt, K + 121, 1) = recLrAttribut.TYRES
Mid$(MsgTxt, K + 122, 1) = recLrAttribut.ZACTI
Mid$(MsgTxt, K + 123, 1) = recLrAttribut.ZAGDT

'attributs Luca Risques
Mid$(MsgTxt, K + 124, 1) = recLrAttribut.CDCPCO
Mid$(MsgTxt, K + 125, 1) = recLrAttribut.CDCPJO
Mid$(MsgTxt, K + 126, 15) = recLrAttribut.CDCPFU
Mid$(MsgTxt, K + 141, 5) = recLrAttribut.CDAGCO
Mid$(MsgTxt, K + 146, 1) = recLrAttribut.CDREME
Mid$(MsgTxt, K + 147, 2) = recLrAttribut.TYMTDV
Mid$(MsgTxt, K + 149, 1) = recLrAttribut.TYVENT
Mid$(MsgTxt, K + 150, 15) = recLrAttribut.CRVENT
Mid$(MsgTxt, K + 165, 1) = recLrAttribut.CDDURE
Mid$(MsgTxt, K + 166, 3) = recLrAttribut.DUINIT
Mid$(MsgTxt, K + 169, 1) = recLrAttribut.CDCRTI
Mid$(MsgTxt, K + 170, 1) = recLrAttribut.CDCRAC
Mid$(MsgTxt, K + 171, 1) = recLrAttribut.CDBIOR
Mid$(MsgTxt, K + 172, 1) = recLrAttribut.CDDEIN
Mid$(MsgTxt, K + 173, 1) = recLrAttribut.CDCRIM
Mid$(MsgTxt, K + 174, 1) = recLrAttribut.CDCRCO
Mid$(MsgTxt, K + 175, 1) = recLrAttribut.CDCREF
Mid$(MsgTxt, K + 176, 1) = recLrAttribut.CDLODA
Mid$(MsgTxt, K + 177, 1) = recLrAttribut.CDCRET
Mid$(MsgTxt, K + 178, 1) = recLrAttribut.CDOMPO
Mid$(MsgTxt, K + 179, 1) = recLrAttribut.CDOPIM
Mid$(MsgTxt, K + 180, 1) = recLrAttribut.CDSWAP
'attributs Réescompte
Mid$(MsgTxt, K + 181, 8) = recLrAttribut.REESC1
Mid$(MsgTxt, K + 189, 8) = recLrAttribut.REESC6

MsgTxtLen = MsgTxtLen + recLrAttributLen
End Sub



'---------------------------------------------------------
Private Function SeekX(recLrAttribut As typeLrAttribut)
'---------------------------------------------------------

SeekX = "?"
MsgTxtLen = 0
Call PutBuffer(recLrAttribut)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(GetBuffer(recLrAttribut)) Then
        SeekX = Null
    Else
        Call ErrorX(recLrAttribut)
    End If
End If

End Function

'---------------------------------------------------------
Private Function Snap(recLrAttribut As typeLrAttribut)
'---------------------------------------------------------
Dim I As Integer
Snap = "?"
MsgTxtLen = 0
Call PutBuffer(recLrAttribut)
Call PutBuffer(arrLrAttribut(0))
If IsNull(SndRcv()) Then
    Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(GetBuffer(recLrAttribut)) Then
            Call srvLrAttribut.AddItem(recLrAttribut)
            arrLrAttributSuite = True
        Else
            arrLrAttributSuite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub Init(recLrAttribut As typeLrAttribut)
'---------------------------------------------------------
MsgTxt = Space$(recLrAttributLen)
MsgTxtIndex = 0
Call GetBuffer(recLrAttribut)
recLrAttribut.obj = "SRVLRATTR"
End Sub

'---------------------------------------------------------
Public Sub AddItem(recLrAttribut As typeLrAttribut)
'---------------------------------------------------------
          
arrLrAttributNb = arrLrAttributNb + 1
    
If arrLrAttributNb > arrLrAttributNbMax Then
    arrLrAttributNbMax = arrLrAttributNbMax + 50
    ReDim Preserve arrLrAttribut(arrLrAttributNbMax)
End If
recLrAttribut.Method = ""
arrLrAttributIndex = arrLrAttributNb
arrLrAttribut(arrLrAttributIndex) = recLrAttribut
End Sub

Public Function Scan(recLrAttribut As typeLrAttribut) As Integer
Scan = -1
For arrLrAttributIndex = 1 To arrLrAttributNb
    If arrLrAttribut(arrLrAttributIndex).Method <> constDelete _
    And arrLrAttribut(arrLrAttributIndex).Method <> constIgnore Then
        If arrLrAttribut(arrLrAttributIndex).Nature = recLrAttribut.Nature _
        And arrLrAttribut(arrLrAttributIndex).Référence = recLrAttribut.Référence Then
            Scan = arrLrAttributIndex
            Exit For
        End If
    End If
Next arrLrAttributIndex

End Function


