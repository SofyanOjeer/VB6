Attribute VB_Name = "srvDSPFFDY0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recDSPFFDY0Len = 848 ' 34 + 814
Public Const recDSPFFDY0_Block = 30

Type typeDSPFFDY0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
     
    WHFILE          As String * 10                    ' File
    WHLIB           As String * 10                    ' Library
    WHCRTD          As String * 7                     ' File creation date: century/dat
    WHFTYP          As String * 1                     ' Type of file: P=Physical, L=Log
    WHCNT           As Long                           ' Number of record formats
    WHDTTM          As String * 13                    ' Retrieval date: century/date/ti
    WHNAME          As String * 10                    ' Record format
    WHSEQ           As String * 13                    ' Format level identifier
    WHTEXT          As String * 50                    ' Format text description
    WHFLDN          As Long                           ' Number of fields and indicators
    WHRLEN          As Long                           ' Record format length
    WHFLDI          As String * 10                    ' Internal field name
    WHFLDE          As String * 10                    ' External field name
    WHFOBO          As Long                           ' Output buffer position
    WHIBO           As Long                           ' Input buffer position
    WHFLDB          As Long                           ' Field length in bytes
    WHFLDD          As Long                           ' Number of digits
    WHFLDP          As Long                           ' Decimal positions to right of d
    WHFTXT          As String * 50                    ' Field text description
    WHRCDE          As Long                           ' 32=Data Type,64=Name, 128=None
    WHRFIL          As String * 10                    ' Reference file
    WHRLIB          As String * 10                    ' Reference library
    WHRFMT          As String * 10                    ' Reference record format
    WHRFLD          As String * 10                    ' Reference field
    WHCHD1          As String * 20                    ' Column heading 1
    WHCHD2          As String * 20                    ' Column heading 2
    WHCHD3          As String * 20                    ' Column heading 3
    WHFLDT          As String * 1                     ' Field type: B,A,S,P,F,O,J,E,H,L
    WHFIOB          As String * 1                     ' I/O attribute: I=Input,O=Output
    WHECDE          As String * 2                     ' Edit code
    WHEWRD          As String * 32                    ' Edit word: Truncated after 30 c
    WHVCNE          As Long                           ' Number of validity checks
    WHNFLD          As Long                           ' Number of fields
    WHNIND          As Long                           ' Number of indicators
    WHSHFT          As String * 1                     ' Keyboard shift
    WHALTY          As String * 1                     ' Character field may be DBSC act
    WHALIS          As String * 30                    ' Alternative field name
    WHJREF          As Long                           ' Join reference to JFILE
    WHDFTL          As Long                           ' DFT value length: -1=Greater th
    WHDFT           As String * 30                    ' DFT value: Truncated after 30 c
    WHCHRI          As String * 1                     ' Character Id changes allowed: N
    WHCTNT          As String * 1                     ' Translation table used:  N=No,
    WHFONT          As String * 10                    ' Font identifier
    WHCSWD          As Double                         ' Character size width
    WHCSHI          As Double                         ' Character size height
    WHBCNM          As String * 10                    ' Barcode name
    WHBCHI          As Double                         ' Barcode height:0=No height defi
    WHMAP           As String * 1                     ' Substring specified: N=No, Y=Ye
    WHMAPS          As Long                           ' Substring starting position
    WHMAPL          As Long                           ' Substring number of bytes (char
    WHSYSN          As String * 8                     ' System Name (Source System, if
    WHRES1          As String * 2                     ' Reserved
    WHSQLT          As String * 1                     ' SQL file type: 0=None, T=TABLE,
    WHHEX           As String * 1                     ' Hexadecimal:  Y=Yes
    WHPNTS          As Double                         ' Point size: 0 = none
    WHCSID          As Long                           ' Coded Character set Identifier
    WHFMT           As String * 4                     ' Date and time format parameters
    WHSEP           As String * 1                     ' '/', '-', '.', ',', ':', ' ', o
    WHVARL          As String * 1                     ' Variable length field: N=No, Y=
    WHALLC          As Long                           ' Allocated length
    WHNULL          As String * 1                     ' Allow Null Value: N=No,Y=Yes
    WHFCSN          As String * 10                    ' Font character set, blank = non
    WHFCSL          As String * 10                    ' Char set lib, blank = none
    WHFCPN          As String * 10                    ' Font code page, blank = none
    WHFCPL          As String * 10                    ' Code page lib, blank = none
    WHCDFN          As String * 10                    ' Coded Font, blank = none
    WHCDFL          As String * 10                    ' Coded font lib, blank = none
    WHDCDF          As String * 10                    ' DBCS Coded Font, blank = none
    WHDCDL          As String * 10                    ' DBCS font lib, blank = none
    WHTXRT          As Long                           ' Degree of text rotation, -1 = n
    WHFLDG          As Long                           ' Field length in characters, 0 =
    WHFDSL          As Long                           ' Alternate field length
    WHFSPS          As Double                         ' Font character set point size.
    WHCFPS          As Double                         ' Coded font point size. 0=*NONE
    WHIFPS          As Double                         ' DBCS Coded font point size. 0=*
    WHDBLL          As Long                           ' LOB field length
    WHDBUN          As String * 128                   ' User-defined type name
    WHDBUL          As String * 10                    ' User-defined type library
    WHDBFC          As String * 1                     ' 0=No control, 1=File control
    WHDBFI          As String * 1                     ' 0=Selective, 1=All, blank=not v
    WHDBRP          As String * 1                     ' 0=Database, 1=File system, blan
    WHDBWP          As String * 1                     ' 0=Blocked, 1=File system, blank
    WHDBRC          As String * 1                     ' 0=No, 1=Yes, blank=not valid
    WHDBOU          As String * 1                     ' 0=Delete, 1=Restore, blank=not
    WHPSUD          As String * 1                     ' 0=*NOCONVERT, 1=*CONVERT
    WHBCUH          As Currency                       ' Barcode height using UOM
    WHFPSW          As Double                         ' Font point size width. 0=Not us
    WHFSPW          As Double                         ' Font character set point size w
    WHCFPW          As Double                         ' Coded font point size width. 0=
    WHIFPW          As Double                         ' DBCS Coded font point size widt
End Type
    
    
Public arrDSPFFDY0() As typeDSPFFDY0
Public arrDSPFFDY0_NB As Integer
Public arrDSPFFDY0_NBMax As Integer
Public arrDSPFFDY0_Index As Integer
Public arrDSPFFDY0_Suite As Boolean

Dim arrRTF_Line(1000) As String, arrRTF_Balise(1000) As Long, arrRTF_Nb As Long
Dim arrRTF_Bold_SelStart(1000) As Long, arrRTF_Bold_SelLength(1000) As Long, arrRTF_Bold_Nb As Long

Public Sub srvDSPFFDY0_frmRTF(lLst As ListBox, lElpKMSrc_Id As Long, lWHLIB As String, lWHFILE As String, blnPrint As Boolean)
Dim I As Integer, K As Integer, intReturn As Integer, x As String, wPFile As String
Dim xRTF As String
Dim wElpKMInfo As typeElpKMInfo
Dim wDSPFFDY0 As typeDSPFFDY0
Dim w14000 As typeElpKMInfo
Dim kSelStart As Long, kSelLength As Long
Dim m_arrRTF_Nb As Integer

arrRTF_Nb = 1
arrRTF_Line(1) = lWHLIB & " / " & lWHFILE
arrRTF_Balise(1) = 1

wPFile = lWHFILE

wElpKMInfo.Method = "Seek="
wElpKMInfo.ElpKMSrc_Id = lElpKMSrc_Id

For I = 0 To lLst.ListCount - 1
    lLst.ListIndex = I
    x = lLst
    wElpKMInfo.Id = mId$(x, 6, Len(x) - 5)
    intReturn = tableElpKMInfo_Read(wElpKMInfo)
    If intReturn = 0 Then
        MsgTxt = Space$(34) & wElpKMInfo.Memo
        MsgTxtIndex = 0
        srvDSPFFDY0_GetBuffer wDSPFFDY0
        
        arrRTF_Nb = arrRTF_Nb + 1
        arrRTF_Line(arrRTF_Nb) = mId$(wElpKMInfo.Description, 20, 21) & " " & wDSPFFDY0.WHFTXT
       arrRTF_Balise(arrRTF_Nb) = 0

        ''''X = mId$(wElpKMInfo.description, 20, 21) & " " & wDSPFFDY0.WHFTXT
        ''''xRTF = xRTF & X & Asc13 & Chr$(10)

    End If
Next I

m_arrRTF_Nb = arrRTF_Nb + 1
Call srvDSPFFDY0_frmRTF_Description(lElpKMSrc_Id, lWHLIB, wPFile, lWHFILE)
arrRTF_Line(1) = arrRTF_Line(m_arrRTF_Nb)

wElpKMInfo.Method = "Seek>="
wElpKMInfo.ElpKMSrc_Id = lElpKMSrc_Id + 2000 '14***
wElpKMInfo.Id = lWHFILE
intReturn = tableElpKMInfo_Read(wElpKMInfo)
wElpKMInfo.Method = "Seek>"

Do
    If intReturn = 0 Then
        If wPFile <> mId$(wElpKMInfo.Id, 1, 10) Then
            intReturn = -1
        Else
            x = mId$(wElpKMInfo.Id, 11, 10)
            Call srvDSPFFDY0_frmRTF_Description(lElpKMSrc_Id, lWHLIB, wPFile, x)
            intReturn = tableElpKMInfo_Read(wElpKMInfo)
        End If
    End If
    
  
Loop While intReturn = 0

If blnPrint Then srvDSPFFDY0_frmRTF_Print: Exit Sub


arrRTF_Bold_Nb = 0
kSelStart = 0
For K = 1 To arrRTF_Nb
    x = arrRTF_Line(K)
    kSelLength = Len(x)
    If arrRTF_Balise(K) <> 2 Then
            xRTF = xRTF & x & Asc13 & Chr$(10)
    Else
            xRTF = xRTF & Asc13 & Chr$(10) & x & Asc13 & Chr$(10)
            kSelStart = kSelStart + 2
    End If
    
    If arrRTF_Balise(K) <> 0 Then
        arrRTF_Bold_Nb = arrRTF_Bold_Nb + 1
        arrRTF_Bold_SelStart(arrRTF_Bold_Nb) = kSelStart
        arrRTF_Bold_SelLength(arrRTF_Bold_Nb) = kSelLength
    End If
    kSelStart = kSelStart + kSelLength + 2
Next K

frmEdition.txtModèle_RTF.Font.Name = prtFontName_CourierNew '
frmEdition.txtModèle_RTF.Font.Size = 8
frmEdition.txtModèle_RTF = xRTF

For K = 1 To arrRTF_Bold_Nb
        frmEdition.txtModèle_RTF.SelStart = arrRTF_Bold_SelStart(K)
        frmEdition.txtModèle_RTF.SelLength = arrRTF_Bold_SelLength(K)
        frmEdition.txtModèle_RTF.SelBold = True
        frmEdition.txtModèle_RTF.SelUnderline = True
        frmEdition.txtModèle_RTF.SelColor = vbRed
        frmEdition.txtModèle_RTF.SelFontSize = 9
Next K


frmRTF_prtOrientation = vbPRORPortrait
frmRTF_blnCourrier = False

frmRTF_Caller = "srvDSPFFDY0_frmRTF"    ' "frmEdition  Display"
frmRTF_Buffer_Name = ""
frmRTF_blnOK = False
frmRTF.cboRTF_Police = frmRTF.txtRTF.Font.Name
frmRTF.txtRTF_Size = frmRTF.txtRTF.Font.Size
frmRTF.txtRTF.TextRTF = frmEdition.txtModèle_RTF.TextRTF



frmRTF.WindowState = vbMaximized
frmRTF.Show vbModal

End Sub

Public Sub srvDSPFFDY0_frmRTF_Print()
Dim K As Integer, x As String, xR As String
Dim lenX As Integer, Height7_6 As Integer

prtTitleText = arrRTF_Line(1)

prtFontName = prtFontName_CourierNew  'Arial
prtOrientation = vbPRORLandscape
prtPgmName = "prtElpKM"
prtTitleUsr = usrName

prtElpKM.prtStd_Open
Height7_6 = frmElpPrt.prtHeightDelta(7, 6)

prtElpKMPgm_Form
XPrt.FontSize = 7
blnMinX12 = False
prtMinMarge = prtMinX1
prtMaxMarge = prtMedX - 100
''Call frmElpPrt.prtTrame(prtMinMarge, XPrt.CurrentY, prtMaxMarge, XPrt.CurrentY + prtlineHeight, " ", 240)
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight

For K = 2 To arrRTF_Nb
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    If XPrt.CurrentY + 200 > prtMaxY Then srvDSPFFDY0_frmRTF_Print_Page

    

    x = arrRTF_Line(K)
    XPrt.CurrentX = prtMinMarge

    If arrRTF_Balise(K) = 2 Then
        If XPrt.CurrentY + 500 > prtMaxY Then srvDSPFFDY0_frmRTF_Print_Page
        XPrt.FontSize = 9
        XPrt.FontBold = True
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight / 2
        Call frmElpPrt.prtTrame(prtMinMarge, XPrt.CurrentY, prtMaxMarge, XPrt.CurrentY + prtlineHeight, " ", 240)
        XPrt.CurrentX = prtMinMarge
        XPrt.Print x;
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight / 2
    Else
        XPrt.FontSize = 7
        XPrt.FontBold = False
        'XPrt.Print mId$(X, 1, 10);
        'xR = Trim(mId$(X, 11, 5))
        'XPrt.CurrentX = prtMinMarge + 1800 - XPrt.TextWidth(xR): XPrt.Print xR;
        'XPrt.CurrentX = prtMinMarge + 1900: XPrt.Print mId$(X, 16, 1);
        'xR = Trim(mId$(X, 17, 5))
        'XPrt.CurrentX = prtMinMarge + 2500 - XPrt.TextWidth(xR): XPrt.Print xR;
        
        xR = Trim(mId$(x, 11, 5))
        XPrt.CurrentX = prtMinMarge + 1000 - XPrt.TextWidth(xR): XPrt.Print xR;
        XPrt.CurrentX = prtMinMarge + 1100: XPrt.Print mId$(x, 16, 1);
       
        XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print mId$(x, 1, 10);
        lenX = Len(x)
        XPrt.FontSize = 6
        XPrt.CurrentY = XPrt.CurrentY + Height7_6
        xR = Trim(mId$(x, 17, 5))
        XPrt.CurrentX = prtMinMarge + 500 - XPrt.TextWidth(xR): XPrt.Print xR;

        XPrt.CurrentX = prtMinMarge + 2800: XPrt.Print mId$(x, 22, lenX - 22);
        XPrt.CurrentY = XPrt.CurrentY - Height7_6
    End If
Next K


prtElpKM.prtStd_Close

End Sub

Public Sub srvDSPFFDY0_lstAddItem(lLst As ListBox, lElpKMSrc_Id As Long, lWHFILE As String, blnAlpha As Boolean)
Dim X5 As String * 5
Dim wPFile As String * 10
Dim wElpKMInfo As typeElpKMInfo
Dim wDSPFFDY0 As typeDSPFFDY0
Dim wDSPFDY1 As typeDSPFDY1
On Error Resume Next

lLst.Clear
lLst.Visible = False

wPFile = lWHFILE
wElpKMInfo.Method = "Seek="
wElpKMInfo.ElpKMSrc_Id = lElpKMSrc_Id + 3000 '15***
wElpKMInfo.Id = lWHFILE
If tableElpKMInfo_Read(wElpKMInfo) = 0 Then
    MsgTxt = Space$(34) & wElpKMInfo.Memo
    MsgTxtIndex = 0
    srvDSPFDY1_GetBuffer wDSPFDY1
    wPFile = wDSPFDY1.APBOF
End If


wElpKMInfo.Method = "Seek>="
wElpKMInfo.ElpKMSrc_Id = lElpKMSrc_Id
wElpKMInfo.Id = wPFile
intReturn = tableElpKMInfo_Read(wElpKMInfo)
wElpKMInfo.Method = "MoveNext"

Do
    If intReturn = 0 Then
        MsgTxt = Space$(34) & wElpKMInfo.Memo
        MsgTxtIndex = 0
        srvDSPFFDY0_GetBuffer wDSPFFDY0
        If wElpKMInfo.ElpKMSrc_Id <> lElpKMSrc_Id Or wDSPFFDY0.WHFILE <> wPFile Then
            intReturn = -1
        Else

            If blnAlpha Then
                X5 = "00000"
            Else
                X5 = Format$(wDSPFFDY0.WHFOBO, "00000")
            End If
            lLst.AddItem X5 & wDSPFFDY0.WHFILE & wDSPFFDY0.WHFLDE
            intReturn = tableElpKMInfo_Read(wElpKMInfo)
    
        End If
    End If
    
  
Loop While intReturn = 0
End Sub

Public Sub srvDSPFFDY0_frmRTF_Key(lElpKMSrc_Id As Long, lPFile As String, lWHFILE As String)
Dim mElpKMSrc_Id As Long
Dim wElpKMInfo As typeElpKMInfo, x As String
Dim wDSPFDY1 As typeDSPFDY1
Dim wField As typeElpKMInfo
Dim wDSPFFDY0 As typeDSPFFDY0
Dim xDes As String

On Error Resume Next

wField.Method = "Seek="
wField.ElpKMSrc_Id = lElpKMSrc_Id


wElpKMInfo.Method = "Seek>="
mElpKMSrc_Id = lElpKMSrc_Id + 1000
wElpKMInfo.ElpKMSrc_Id = mElpKMSrc_Id '13***
wElpKMInfo.Id = lWHFILE
intReturn = tableElpKMInfo_Read(wElpKMInfo)
wElpKMInfo.Method = "Seek>"

Do
    If intReturn = 0 Then
        MsgTxt = Space$(34) & wElpKMInfo.Memo
        MsgTxtIndex = 0
        srvDSPFDY1_GetBuffer wDSPFDY1
        If wElpKMInfo.ElpKMSrc_Id <> mElpKMSrc_Id Or wDSPFDY1.APFILE <> lWHFILE Then
            intReturn = -1
        Else
            xDes = "   "
            If wDSPFDY1.APUNIQ = "Y" Then Mid$(xDes, 1, 1) = "U"
            If wDSPFDY1.APSELO = "Y" Then Mid$(xDes, 2, 1) = "S"
            If wDSPFDY1.APKSEQ = "D" Then Mid$(xDes, 3, 1) = "D"
           '''X = wDSPFDY1.APKEYF & " " & wDSPFDY1.APKEYN
           ''' lRTF = lRTF & X & Asc13 & Chr$(10)
            wField.Id = lPFile & wDSPFDY1.APKEYF
            intReturn = tableElpKMInfo_Read(wField)
            If intReturn = 0 Then
                MsgTxt = Space$(34) & wField.Memo
                MsgTxtIndex = 0
                srvDSPFFDY0_GetBuffer wDSPFFDY0
                
        
                arrRTF_Nb = arrRTF_Nb + 1
                arrRTF_Line(arrRTF_Nb) = mId$(wField.Description, 20, 21) & " " & wDSPFFDY0.WHFTXT & xDes
               arrRTF_Balise(arrRTF_Nb) = 0
                ''''X = mId$(wField.description, 20, 21) & " " & wDSPFFDY0.WHFTXT
                '''''lRTF = lRTF & X & Asc13 & Chr$(10)
            
            End If

            intReturn = tableElpKMInfo_Read(wElpKMInfo)
    
        End If
    End If
      
Loop While intReturn = 0




End Sub


'-----------------------------------------------------
Function srvDSPFFDY0_Update(recDSPFFDY0 As typeDSPFFDY0)
'-----------------------------------------------------

srvDSPFFDY0_Update = "?"

MsgTxtLen = 0
Call srvDSPFFDY0_PutBuffer(recDSPFFDY0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvDSPFFDY0_GetBuffer(recDSPFFDY0)) Then
        Call srvDSPFFDY0_Error(recDSPFFDY0)
        srvDSPFFDY0_Update = recDSPFFDY0.Err
        Exit Function
    Else
        srvDSPFFDY0_Update = Null
    End If
Else
    recDSPFFDY0.Err = "srv"
End If
End Function

'-----------------------------------------------------
Sub srvDSPFFDY0_Error(recDSPFFDY0 As typeDSPFFDY0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "DSPFFDY0" & Chr$(10) & Chr$(13)

Select Case mId$(recDSPFFDY0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recDSPFFDY0.Err
        I = vbCritical
End Select

MsgBox Msg & " : " & recDSPFFDY0.WHFILE _
        , I, "module : DSPFFDY0s.bas  ( " & Trim(recDSPFFDY0.obj) & " : " & Trim(recDSPFFDY0.Method) & " )"

End Sub


'=====================================================



'-----------------------------------------------------
Public Function srvDSPFFDY0_Monitor(recDSPFFDY0 As typeDSPFFDY0)
'-----------------------------------------------------

arrDSPFFDY0_Suite = False
Select Case mId$(Trim(recDSPFFDY0.Method), 1, 4)
    Case "Snap"
              srvDSPFFDY0_Monitor = srvDSPFFDY0_Snap(recDSPFFDY0)
    Case Else
            srvDSPFFDY0_Monitor = srvDSPFFDY0_Seek(recDSPFFDY0)
End Select

End Function

'---------------------------------------------------------
Public Function srvDSPFFDY0_GetBuffer(recDSPFFDY0 As typeDSPFFDY0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvDSPFFDY0_GetBuffer = Null
recDSPFFDY0.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recDSPFFDY0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recDSPFFDY0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recDSPFFDY0.Err = Space$(10) Then
    recDSPFFDY0.WHFILE = mId$(MsgTxt, K + 1, 10)
    recDSPFFDY0.WHLIB = mId$(MsgTxt, K + 11, 10)
    recDSPFFDY0.WHCRTD = mId$(MsgTxt, K + 21, 7)
    recDSPFFDY0.WHFTYP = mId$(MsgTxt, K + 28, 1)
    recDSPFFDY0.WHCNT = CLng(Val(mId$(MsgTxt, K + 29, 6)))
    recDSPFFDY0.WHDTTM = mId$(MsgTxt, K + 35, 13)
    recDSPFFDY0.WHNAME = mId$(MsgTxt, K + 48, 10)
    recDSPFFDY0.WHSEQ = mId$(MsgTxt, K + 58, 13)
    recDSPFFDY0.WHTEXT = mId$(MsgTxt, K + 71, 50)
    recDSPFFDY0.WHFLDN = CLng(Val(mId$(MsgTxt, K + 121, 6)))
    recDSPFFDY0.WHRLEN = CLng(Val(mId$(MsgTxt, K + 127, 6)))
    recDSPFFDY0.WHFLDI = mId$(MsgTxt, K + 133, 10)
    recDSPFFDY0.WHFLDE = mId$(MsgTxt, K + 143, 10)
    recDSPFFDY0.WHFOBO = CLng(Val(mId$(MsgTxt, K + 153, 6)))
    recDSPFFDY0.WHIBO = CLng(Val(mId$(MsgTxt, K + 159, 6)))
    recDSPFFDY0.WHFLDB = CLng(Val(mId$(MsgTxt, K + 165, 6)))
    recDSPFFDY0.WHFLDD = CLng(Val(mId$(MsgTxt, K + 171, 3)))
    recDSPFFDY0.WHFLDP = CLng(Val(mId$(MsgTxt, K + 174, 3)))
    recDSPFFDY0.WHFTXT = mId$(MsgTxt, K + 177, 50)
    recDSPFFDY0.WHRCDE = CLng(Val(mId$(MsgTxt, K + 227, 4)))
    recDSPFFDY0.WHRFIL = mId$(MsgTxt, K + 231, 10)
    recDSPFFDY0.WHRLIB = mId$(MsgTxt, K + 241, 10)
    recDSPFFDY0.WHRFMT = mId$(MsgTxt, K + 251, 10)
    recDSPFFDY0.WHRFLD = mId$(MsgTxt, K + 261, 10)
    recDSPFFDY0.WHCHD1 = mId$(MsgTxt, K + 271, 20)
    recDSPFFDY0.WHCHD2 = mId$(MsgTxt, K + 291, 20)
    recDSPFFDY0.WHCHD3 = mId$(MsgTxt, K + 311, 20)
    recDSPFFDY0.WHFLDT = mId$(MsgTxt, K + 331, 1)
    recDSPFFDY0.WHFIOB = mId$(MsgTxt, K + 332, 1)
    recDSPFFDY0.WHECDE = mId$(MsgTxt, K + 333, 2)
    recDSPFFDY0.WHEWRD = mId$(MsgTxt, K + 335, 32)
    recDSPFFDY0.WHVCNE = CLng(Val(mId$(MsgTxt, K + 367, 5)))
    recDSPFFDY0.WHNFLD = CLng(Val(mId$(MsgTxt, K + 372, 6)))
    recDSPFFDY0.WHNIND = CLng(Val(mId$(MsgTxt, K + 378, 3)))
    recDSPFFDY0.WHSHFT = mId$(MsgTxt, K + 381, 1)
    recDSPFFDY0.WHALTY = mId$(MsgTxt, K + 382, 1)
    recDSPFFDY0.WHALIS = mId$(MsgTxt, K + 383, 30)
    recDSPFFDY0.WHJREF = CLng(Val(mId$(MsgTxt, K + 413, 3)))
    recDSPFFDY0.WHDFTL = CLng(Val(mId$(MsgTxt, K + 416, 3)))
    recDSPFFDY0.WHDFT = mId$(MsgTxt, K + 419, 30)
    recDSPFFDY0.WHCHRI = mId$(MsgTxt, K + 449, 1)
    recDSPFFDY0.WHCTNT = mId$(MsgTxt, K + 450, 1)
    recDSPFFDY0.WHFONT = mId$(MsgTxt, K + 451, 10)
    recDSPFFDY0.WHCSWD = CDbl(Val(mId$(MsgTxt, K + 461, 4))) / 10
    recDSPFFDY0.WHCSHI = CDbl(Val(mId$(MsgTxt, K + 465, 4))) / 10
    recDSPFFDY0.WHBCNM = mId$(MsgTxt, K + 469, 10)
    recDSPFFDY0.WHBCHI = CDbl(Val(mId$(MsgTxt, K + 479, 4))) / 10
    recDSPFFDY0.WHMAP = mId$(MsgTxt, K + 483, 1)
    recDSPFFDY0.WHMAPS = CLng(Val(mId$(MsgTxt, K + 484, 6)))
    recDSPFFDY0.WHMAPL = CLng(Val(mId$(MsgTxt, K + 490, 6)))
    recDSPFFDY0.WHSYSN = mId$(MsgTxt, K + 496, 8)
    recDSPFFDY0.WHRES1 = mId$(MsgTxt, K + 504, 2)
    recDSPFFDY0.WHSQLT = mId$(MsgTxt, K + 506, 1)
    recDSPFFDY0.WHHEX = mId$(MsgTxt, K + 507, 1)
    recDSPFFDY0.WHPNTS = CDbl(Val(mId$(MsgTxt, K + 508, 5))) / 10
    recDSPFFDY0.WHCSID = CLng(Val(mId$(MsgTxt, K + 513, 6)))
    recDSPFFDY0.WHFMT = mId$(MsgTxt, K + 519, 4)
    recDSPFFDY0.WHSEP = mId$(MsgTxt, K + 523, 1)
    recDSPFFDY0.WHVARL = mId$(MsgTxt, K + 524, 1)
    recDSPFFDY0.WHALLC = CLng(Val(mId$(MsgTxt, K + 525, 6)))
    recDSPFFDY0.WHNULL = mId$(MsgTxt, K + 531, 1)
    recDSPFFDY0.WHFCSN = mId$(MsgTxt, K + 532, 10)
    recDSPFFDY0.WHFCSL = mId$(MsgTxt, K + 542, 10)
    recDSPFFDY0.WHFCPN = mId$(MsgTxt, K + 552, 10)
    recDSPFFDY0.WHFCPL = mId$(MsgTxt, K + 562, 10)
    recDSPFFDY0.WHCDFN = mId$(MsgTxt, K + 572, 10)
    recDSPFFDY0.WHCDFL = mId$(MsgTxt, K + 582, 10)
    recDSPFFDY0.WHDCDF = mId$(MsgTxt, K + 592, 10)
    recDSPFFDY0.WHDCDL = mId$(MsgTxt, K + 602, 10)
    recDSPFFDY0.WHTXRT = CLng(Val(mId$(MsgTxt, K + 612, 4)))
    recDSPFFDY0.WHFLDG = CLng(Val(mId$(MsgTxt, K + 616, 6)))
    recDSPFFDY0.WHFDSL = CLng(Val(mId$(MsgTxt, K + 622, 6)))
    recDSPFFDY0.WHFSPS = CDbl(Val(mId$(MsgTxt, K + 628, 5))) / 10
    recDSPFFDY0.WHCFPS = CDbl(Val(mId$(MsgTxt, K + 633, 5))) / 10
    recDSPFFDY0.WHIFPS = CDbl(Val(mId$(MsgTxt, K + 638, 5))) / 10
    recDSPFFDY0.WHDBLL = CLng(Val(mId$(MsgTxt, K + 643, 11)))
    recDSPFFDY0.WHDBUN = mId$(MsgTxt, K + 654, 128)
    recDSPFFDY0.WHDBUL = mId$(MsgTxt, K + 782, 10)
    recDSPFFDY0.WHDBFC = mId$(MsgTxt, K + 792, 1)
    recDSPFFDY0.WHDBFI = mId$(MsgTxt, K + 793, 1)
    recDSPFFDY0.WHDBRP = mId$(MsgTxt, K + 794, 1)
    recDSPFFDY0.WHDBWP = mId$(MsgTxt, K + 795, 1)
    recDSPFFDY0.WHDBRC = mId$(MsgTxt, K + 796, 1)
    recDSPFFDY0.WHDBOU = mId$(MsgTxt, K + 797, 1)
    recDSPFFDY0.WHPSUD = mId$(MsgTxt, K + 798, 1)
    recDSPFFDY0.WHBCUH = CCur(mId$(MsgTxt, K + 799, 6)) / 100
    recDSPFFDY0.WHFPSW = CDbl(Val(mId$(MsgTxt, K + 805, 5))) / 10
    recDSPFFDY0.WHFSPW = CDbl(Val(mId$(MsgTxt, K + 810, 5))) / 10
    recDSPFFDY0.WHCFPW = CDbl(Val(mId$(MsgTxt, K + 815, 5))) / 10
    recDSPFFDY0.WHIFPW = CDbl(Val(mId$(MsgTxt, K + 820, 5))) / 10

Else
    srvDSPFFDY0_GetBuffer = recDSPFFDY0.Err
End If

MsgTxtIndex = MsgTxtIndex + recDSPFFDY0Len

End Function

'---------------------------------------------------------
Private Sub srvDSPFFDY0_PutBuffer(recDSPFFDY0 As typeDSPFFDY0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recDSPFFDY0.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recDSPFFDY0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 10) = recDSPFFDY0.WHFILE
    Mid$(MsgTxt, K + 11, 10) = recDSPFFDY0.WHLIB
    Mid$(MsgTxt, K + 21, 7) = recDSPFFDY0.WHCRTD
    Mid$(MsgTxt, K + 28, 1) = recDSPFFDY0.WHFTYP
    Mid$(MsgTxt, K + 29, 6) = Format$(recDSPFFDY0.WHCNT, "00000 ")
    Mid$(MsgTxt, K + 35, 13) = recDSPFFDY0.WHDTTM
    Mid$(MsgTxt, K + 48, 10) = recDSPFFDY0.WHNAME
    Mid$(MsgTxt, K + 58, 13) = recDSPFFDY0.WHSEQ
    Mid$(MsgTxt, K + 71, 50) = recDSPFFDY0.WHTEXT
    Mid$(MsgTxt, K + 121, 6) = Format$(recDSPFFDY0.WHFLDN, "00000 ")
    Mid$(MsgTxt, K + 127, 6) = Format$(recDSPFFDY0.WHRLEN, "00000 ")
    Mid$(MsgTxt, K + 133, 10) = recDSPFFDY0.WHFLDI
    Mid$(MsgTxt, K + 143, 10) = recDSPFFDY0.WHFLDE
    Mid$(MsgTxt, K + 153, 6) = Format$(recDSPFFDY0.WHFOBO, "00000 ")
    Mid$(MsgTxt, K + 159, 6) = Format$(recDSPFFDY0.WHIBO, "00000 ")
    Mid$(MsgTxt, K + 165, 6) = Format$(recDSPFFDY0.WHFLDB, "00000 ")
    Mid$(MsgTxt, K + 171, 3) = Format$(recDSPFFDY0.WHFLDD, "00 ")
    Mid$(MsgTxt, K + 174, 3) = Format$(recDSPFFDY0.WHFLDP, "00 ")
    Mid$(MsgTxt, K + 177, 50) = recDSPFFDY0.WHFTXT
    Mid$(MsgTxt, K + 227, 4) = Format$(recDSPFFDY0.WHRCDE, "000 ")
    Mid$(MsgTxt, K + 231, 10) = recDSPFFDY0.WHRFIL
    Mid$(MsgTxt, K + 241, 10) = recDSPFFDY0.WHRLIB
    Mid$(MsgTxt, K + 251, 10) = recDSPFFDY0.WHRFMT
    Mid$(MsgTxt, K + 261, 10) = recDSPFFDY0.WHRFLD
    Mid$(MsgTxt, K + 271, 20) = recDSPFFDY0.WHCHD1
    Mid$(MsgTxt, K + 291, 20) = recDSPFFDY0.WHCHD2
    Mid$(MsgTxt, K + 311, 20) = recDSPFFDY0.WHCHD3
    Mid$(MsgTxt, K + 331, 1) = recDSPFFDY0.WHFLDT
    Mid$(MsgTxt, K + 332, 1) = recDSPFFDY0.WHFIOB
    Mid$(MsgTxt, K + 333, 2) = recDSPFFDY0.WHECDE
    Mid$(MsgTxt, K + 335, 32) = recDSPFFDY0.WHEWRD
    Mid$(MsgTxt, K + 367, 5) = Format$(recDSPFFDY0.WHVCNE, "0000 ")
    Mid$(MsgTxt, K + 372, 6) = Format$(recDSPFFDY0.WHNFLD, "00000 ")
    Mid$(MsgTxt, K + 378, 3) = Format$(recDSPFFDY0.WHNIND, "00 ")
    Mid$(MsgTxt, K + 381, 1) = recDSPFFDY0.WHSHFT
    Mid$(MsgTxt, K + 382, 1) = recDSPFFDY0.WHALTY
    Mid$(MsgTxt, K + 383, 30) = recDSPFFDY0.WHALIS
    Mid$(MsgTxt, K + 413, 3) = Format$(recDSPFFDY0.WHJREF, "00 ")
    Mid$(MsgTxt, K + 416, 3) = Format$(recDSPFFDY0.WHDFTL, "00 ")
    Mid$(MsgTxt, K + 419, 30) = recDSPFFDY0.WHDFT
    Mid$(MsgTxt, K + 449, 1) = recDSPFFDY0.WHCHRI
    Mid$(MsgTxt, K + 450, 1) = recDSPFFDY0.WHCTNT
    Mid$(MsgTxt, K + 451, 10) = recDSPFFDY0.WHFONT
    Mid$(MsgTxt, K + 461, 4) = Format$(recDSPFFDY0.WHCSWD * 10, "000 ")
    Mid$(MsgTxt, K + 465, 4) = Format$(recDSPFFDY0.WHCSHI * 10, "000 ")
    Mid$(MsgTxt, K + 469, 10) = recDSPFFDY0.WHBCNM
    Mid$(MsgTxt, K + 479, 4) = Format$(recDSPFFDY0.WHBCHI * 10, "000 ")
    Mid$(MsgTxt, K + 483, 1) = recDSPFFDY0.WHMAP
    Mid$(MsgTxt, K + 484, 6) = Format$(recDSPFFDY0.WHMAPS, "00000 ")
    Mid$(MsgTxt, K + 490, 6) = Format$(recDSPFFDY0.WHMAPL, "00000 ")
    Mid$(MsgTxt, K + 496, 8) = recDSPFFDY0.WHSYSN
    Mid$(MsgTxt, K + 504, 2) = recDSPFFDY0.WHRES1
    Mid$(MsgTxt, K + 506, 1) = recDSPFFDY0.WHSQLT
    Mid$(MsgTxt, K + 507, 1) = recDSPFFDY0.WHHEX
    Mid$(MsgTxt, K + 508, 5) = Format$(recDSPFFDY0.WHPNTS * 10, "0000 ")
    Mid$(MsgTxt, K + 513, 6) = Format$(recDSPFFDY0.WHCSID, "00000 ")
    Mid$(MsgTxt, K + 519, 4) = recDSPFFDY0.WHFMT
    Mid$(MsgTxt, K + 523, 1) = recDSPFFDY0.WHSEP
    Mid$(MsgTxt, K + 524, 1) = recDSPFFDY0.WHVARL
    Mid$(MsgTxt, K + 525, 6) = Format$(recDSPFFDY0.WHALLC, "00000 ")
    Mid$(MsgTxt, K + 531, 1) = recDSPFFDY0.WHNULL
    Mid$(MsgTxt, K + 532, 10) = recDSPFFDY0.WHFCSN
    Mid$(MsgTxt, K + 542, 10) = recDSPFFDY0.WHFCSL
    Mid$(MsgTxt, K + 552, 10) = recDSPFFDY0.WHFCPN
    Mid$(MsgTxt, K + 562, 10) = recDSPFFDY0.WHFCPL
    Mid$(MsgTxt, K + 572, 10) = recDSPFFDY0.WHCDFN
    Mid$(MsgTxt, K + 582, 10) = recDSPFFDY0.WHCDFL
    Mid$(MsgTxt, K + 592, 10) = recDSPFFDY0.WHDCDF
    Mid$(MsgTxt, K + 602, 10) = recDSPFFDY0.WHDCDL
    Mid$(MsgTxt, K + 612, 4) = Format$(recDSPFFDY0.WHTXRT, "000 ")
    Mid$(MsgTxt, K + 616, 6) = Format$(recDSPFFDY0.WHFLDG, "00000 ")
    Mid$(MsgTxt, K + 622, 6) = Format$(recDSPFFDY0.WHFDSL, "00000 ")
    Mid$(MsgTxt, K + 628, 5) = Format$(recDSPFFDY0.WHFSPS * 10, "0000 ")
    Mid$(MsgTxt, K + 633, 5) = Format$(recDSPFFDY0.WHCFPS * 10, "0000 ")
    Mid$(MsgTxt, K + 638, 5) = Format$(recDSPFFDY0.WHIFPS * 10, "0000 ")
    Mid$(MsgTxt, K + 643, 11) = Format$(recDSPFFDY0.WHDBLL, "0000000000 ")
    Mid$(MsgTxt, K + 654, 128) = recDSPFFDY0.WHDBUN
    Mid$(MsgTxt, K + 782, 10) = recDSPFFDY0.WHDBUL
    Mid$(MsgTxt, K + 792, 1) = recDSPFFDY0.WHDBFC
    Mid$(MsgTxt, K + 793, 1) = recDSPFFDY0.WHDBFI
    Mid$(MsgTxt, K + 794, 1) = recDSPFFDY0.WHDBRP
    Mid$(MsgTxt, K + 795, 1) = recDSPFFDY0.WHDBWP
    Mid$(MsgTxt, K + 796, 1) = recDSPFFDY0.WHDBRC
    Mid$(MsgTxt, K + 797, 1) = recDSPFFDY0.WHDBOU
    Mid$(MsgTxt, K + 798, 1) = recDSPFFDY0.WHPSUD
    Mid$(MsgTxt, K + 799, 6) = Format$(recDSPFFDY0.WHBCUH * 100, "00000 ")
    Mid$(MsgTxt, K + 805, 5) = Format$(recDSPFFDY0.WHFPSW * 10, "0000 ")
    Mid$(MsgTxt, K + 810, 5) = Format$(recDSPFFDY0.WHFSPW * 10, "0000 ")
    Mid$(MsgTxt, K + 815, 5) = Format$(recDSPFFDY0.WHCFPW * 10, "0000 ")
    Mid$(MsgTxt, K + 820, 5) = Format$(recDSPFFDY0.WHIFPW * 10, "0000 ")
    

MsgTxtLen = MsgTxtLen + recDSPFFDY0Len
End Sub


Public Sub srvDSPFFDY0_ElpDisplay(recDSPFFDY0 As typeDSPFFDY0)
frmElpDisplay.fgData.Rows = 91
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHFILE   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "File"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHFILE
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHLIB   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Library"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHLIB
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHCRTD    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "File creation date: century/date"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHCRTD
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHFTYP    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Type of file: P=Physical, L=Logical, D=Device"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHFTYP
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHCNT    5S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Number of record formats"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHCNT
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHDTTM   13A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Retrieval date: century/date/time"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHDTTM
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHNAME   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Record format"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHNAME
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHSEQ   13A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Format level identifier"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHSEQ
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHTEXT   50A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Format text description"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHTEXT
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHFLDN    5S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Number of fields and indicators"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHFLDN
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHRLEN    5S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Record format length"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHRLEN
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHFLDI   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Internal field name"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHFLDI
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHFLDE   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "External field name"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHFLDE
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHFOBO    5S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Output buffer position"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHFOBO
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHIBO    5S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Input buffer position"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHIBO
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHFLDB    5S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Field length in bytes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHFLDB
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHFLDD    2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Number of digits"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHFLDD
frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHFLDP    2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Decimal positions to right of decimal"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHFLDP
frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHFTXT   50A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Field text description"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHFTXT
frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHRCDE    3S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "32=Data Type,64=Name, 128=None"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHRCDE
frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHRFIL   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Reference file"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHRFIL
frmElpDisplay.fgData.Row = 22
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHRLIB   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Reference library"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHRLIB
frmElpDisplay.fgData.Row = 23
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHRFMT   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Reference record format"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHRFMT
frmElpDisplay.fgData.Row = 24
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHRFLD   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Reference field"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHRFLD
frmElpDisplay.fgData.Row = 25
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHCHD1   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Column heading 1"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHCHD1
frmElpDisplay.fgData.Row = 26
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHCHD2   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Column heading 2"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHCHD2
frmElpDisplay.fgData.Row = 27
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHCHD3   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Column heading 3"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHCHD3
frmElpDisplay.fgData.Row = 28
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHFLDT    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Field type: B,A,S,P,F,O,J,E,H,L,T,Z,G,1,2,3,4"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHFLDT
frmElpDisplay.fgData.Row = 29
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHFIOB    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "I/O attribute: I=Input,O=Output,B=Both,N=Neither"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHFIOB
frmElpDisplay.fgData.Row = 30
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHECDE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Edit code"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHECDE
frmElpDisplay.fgData.Row = 31
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHEWRD   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Edit word: Truncated after 30 characters"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHEWRD
frmElpDisplay.fgData.Row = 32
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHVCNE    4S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Number of validity checks"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHVCNE
frmElpDisplay.fgData.Row = 33
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHNFLD    5S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Number of fields"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHNFLD
frmElpDisplay.fgData.Row = 34
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHNIND    2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Number of indicators"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHNIND
frmElpDisplay.fgData.Row = 35
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHSHFT    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Keyboard shift"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHSHFT
frmElpDisplay.fgData.Row = 36
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHALTY    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Character field may be DBSC activated: N=No, Y=Yes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHALTY
frmElpDisplay.fgData.Row = 37
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHALIS   30A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Alternative field name"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHALIS
frmElpDisplay.fgData.Row = 38
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHJREF    2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Join reference to JFILE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHJREF
frmElpDisplay.fgData.Row = 39
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHDFTL    2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DFT value length: -1=Greater than 30 characters"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHDFTL
frmElpDisplay.fgData.Row = 40
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHDFT   30A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DFT value: Truncated after 30 characters"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHDFT
frmElpDisplay.fgData.Row = 41
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHCHRI    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Character Id changes allowed: N=No, Y=Yes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHCHRI
frmElpDisplay.fgData.Row = 42
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHCTNT    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Translation table used:  N=No, Y=Yes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHCTNT
frmElpDisplay.fgData.Row = 43
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHFONT   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Font identifier"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHFONT
frmElpDisplay.fgData.Row = 44
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHCSWD  3.1S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Character size width"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHCSWD
frmElpDisplay.fgData.Row = 45
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHCSHI  3.1S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Character size height"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHCSHI
frmElpDisplay.fgData.Row = 46
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHBCNM   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Barcode name"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHBCNM
frmElpDisplay.fgData.Row = 47
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHBCHI  3.1S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Barcode height:0=No height defined,-2=use WHBCUH"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHBCHI
frmElpDisplay.fgData.Row = 48
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHMAP    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Substring specified: N=No, Y=Yes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHMAP
frmElpDisplay.fgData.Row = 49
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHMAPS    5S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Substring starting position"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHMAPS
frmElpDisplay.fgData.Row = 50
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHMAPL    5S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Substring number of bytes (characters, if graphic)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHMAPL
frmElpDisplay.fgData.Row = 51
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHSYSN    8A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "System Name (Source System, if file is DDM)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHSYSN
frmElpDisplay.fgData.Row = 52
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHRES1    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Reserved"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHRES1
frmElpDisplay.fgData.Row = 53
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHSQLT    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SQL file type: 0=None, T=TABLE, I=INDEX, V=VIEW"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHSQLT
frmElpDisplay.fgData.Row = 54
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHHEX    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Hexadecimal:  Y=Yes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHHEX
frmElpDisplay.fgData.Row = 55
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHPNTS  4.1P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Point size: 0 = none"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHPNTS
frmElpDisplay.fgData.Row = 56
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHCSID    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Coded Character set Identifier"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHCSID
frmElpDisplay.fgData.Row = 57
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHFMT    4A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Date and time format parameters"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHFMT
frmElpDisplay.fgData.Row = 58
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHSEP    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "'/', '-', '.', ',', ':', ' ', or 'J'=*JOB"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHSEP
frmElpDisplay.fgData.Row = 59
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHVARL    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Variable length field: N=No, Y=Yes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHVARL
frmElpDisplay.fgData.Row = 60
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHALLC    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Allocated length"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHALLC
frmElpDisplay.fgData.Row = 61
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHNULL    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Allow Null Value: N=No,Y=Yes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHNULL
frmElpDisplay.fgData.Row = 62
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHFCSN   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Font character set, blank = none"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHFCSN
frmElpDisplay.fgData.Row = 63
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHFCSL   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Char set lib, blank = none"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHFCSL
frmElpDisplay.fgData.Row = 64
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHFCPN   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Font code page, blank = none"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHFCPN
frmElpDisplay.fgData.Row = 65
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHFCPL   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Code page lib, blank = none"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHFCPL
frmElpDisplay.fgData.Row = 66
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHCDFN   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Coded Font, blank = none"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHCDFN
frmElpDisplay.fgData.Row = 67
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHCDFL   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Coded font lib, blank = none"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHCDFL
frmElpDisplay.fgData.Row = 68
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHDCDF   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DBCS Coded Font, blank = none"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHDCDF
frmElpDisplay.fgData.Row = 69
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHDCDL   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DBCS font lib, blank = none"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHDCDL
frmElpDisplay.fgData.Row = 70
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHTXRT    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Degree of text rotation, -1 = not specified"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHTXRT
frmElpDisplay.fgData.Row = 71
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHFLDG    5S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Field length in characters, 0 = not graphic"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHFLDG
frmElpDisplay.fgData.Row = 72
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHFDSL    5S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Alternate field length"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHFDSL
frmElpDisplay.fgData.Row = 73
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHFSPS  4.1P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Font character set point size. 0=*NONE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHFSPS
frmElpDisplay.fgData.Row = 74
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHCFPS  4.1P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Coded font point size. 0=*NONE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHCFPS
frmElpDisplay.fgData.Row = 75
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHIFPS  4.1P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DBCS Coded font point size. 0=*NONE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHIFPS
frmElpDisplay.fgData.Row = 76
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHDBLL   10P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LOB field length"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHDBLL
frmElpDisplay.fgData.Row = 77
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHDBUN  128A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "User-defined type name"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHDBUN
frmElpDisplay.fgData.Row = 78
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHDBUL   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "User-defined type library"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHDBUL
frmElpDisplay.fgData.Row = 79
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHDBFC    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "0=No control, 1=File control"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHDBFC
frmElpDisplay.fgData.Row = 80
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHDBFI    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "0=Selective, 1=All, blank=not valid"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHDBFI
frmElpDisplay.fgData.Row = 81
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHDBRP    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "0=Database, 1=File system, blank=not valid"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHDBRP
frmElpDisplay.fgData.Row = 82
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHDBWP    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "0=Blocked, 1=File system, blank=not valid"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHDBWP
frmElpDisplay.fgData.Row = 83
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHDBRC    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "0=No, 1=Yes, blank=not valid"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHDBRC
frmElpDisplay.fgData.Row = 84
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHDBOU    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "0=Delete, 1=Restore, blank=not valid"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHDBOU
frmElpDisplay.fgData.Row = 85
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHPSUD    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "0=*NOCONVERT, 1=*CONVERT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHPSUD
frmElpDisplay.fgData.Row = 86
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHBCUH  5.2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Barcode height using UOM"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHBCUH
frmElpDisplay.fgData.Row = 87
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHFPSW  4.1P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Font point size width. 0=Not used"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHFPSW
frmElpDisplay.fgData.Row = 88
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHFSPW  4.1P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Font character set point size width. 0=Not used"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHFSPW
frmElpDisplay.fgData.Row = 89
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHCFPW  4.1P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Coded font point size width. 0=Not used"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHCFPW
frmElpDisplay.fgData.Row = 90
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "WHIFPW  4.1P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DBCS Coded font point size width. 0=Not used"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFFDY0.WHIFPW
frmElpDisplay.Show vbModal
End Sub

'---------------------------------------------------------
Private Function srvDSPFFDY0_Seek(recDSPFFDY0 As typeDSPFFDY0)
'---------------------------------------------------------

srvDSPFFDY0_Seek = "?"
MsgTxtLen = 0
Call srvDSPFFDY0_PutBuffer(recDSPFFDY0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvDSPFFDY0_GetBuffer(recDSPFFDY0)) Then
        srvDSPFFDY0_Seek = Null
    Else
        Call srvDSPFFDY0_Error(recDSPFFDY0)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvDSPFFDY0_Snap(recDSPFFDY0 As typeDSPFFDY0)
'---------------------------------------------------------
srvDSPFFDY0_Snap = "?"
MsgTxtLen = 0
Call srvDSPFFDY0_PutBuffer(recDSPFFDY0)
Call srvDSPFFDY0_PutBuffer(arrDSPFFDY0(0))
If IsNull(SndRcv()) Then
    srvDSPFFDY0_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvDSPFFDY0_GetBuffer(recDSPFFDY0)) Then
            Call arrDSPFFDY0_AddItem(recDSPFFDY0)
            arrDSPFFDY0_Suite = True
        Else
            arrDSPFFDY0_Suite = False
            Exit Do
        End If
    Loop
End If

End Function
'---------------------------------------------------------
Public Sub arrDSPFFDY0_AddItem(recDSPFFDY0 As typeDSPFFDY0)
'---------------------------------------------------------
          
arrDSPFFDY0_NB = arrDSPFFDY0_NB + 1
    
If arrDSPFFDY0_NB > arrDSPFFDY0_NBMax Then
    arrDSPFFDY0_NBMax = arrDSPFFDY0_NBMax + recDSPFFDY0_Block
    ReDim Preserve arrDSPFFDY0(arrDSPFFDY0_NBMax)
End If
            
arrDSPFFDY0(arrDSPFFDY0_NB) = recDSPFFDY0
End Sub



'---------------------------------------------------------
Public Sub recDSPFFDY0_Init(recDSPFFDY0 As typeDSPFFDY0)
'---------------------------------------------------------
recDSPFFDY0.obj = "DSPFFDY0"
recDSPFFDY0.Method = ""
recDSPFFDY0.Err = ""
recDSPFFDY0.WHFILE = ""
recDSPFFDY0.WHFLDE = ""
recDSPFFDY0.WHFOBO = 0
recDSPFFDY0.WHFLDB = 0
recDSPFFDY0.WHFLDD = 0
recDSPFFDY0.WHFLDP = 0
recDSPFFDY0.WHCHD1 = ""
recDSPFFDY0.WHFLDT = ""

End Sub




Public Static Function srvDSPFFDY0_Library(lLibrary As String) As Integer
Dim Id As Integer
Select Case mId$(lLibrary, 1, 3)
    Case "SAB": Id = 100
    Case "BIA": Id = 200
    Case Else: Id = 0
End Select
srvDSPFFDY0_Library = Id
End Function

Public Sub srvDSPFFDY0_frmRTF_Description(lElpKMSrc_Id As Long, lWHLIB As String, lPFile As String, lWHFILE As String)
Dim wDSPFDY0 As typeDSPFDY0, w11000 As typeElpKMInfo
Dim x As String

w11000.ElpKMSrc_Id = 11000
w11000.Method = "Seek="
w11000.Id = lWHLIB
Mid$(w11000.Id, 11, 10) = lWHFILE
intReturn = tableElpKMInfo_Read(w11000)
If intReturn = 0 Then
        MsgTxt = Space$(34) & w11000.Memo
        MsgTxtIndex = 0
        srvDSPFDY0_GetBuffer wDSPFDY0
        
        arrRTF_Nb = arrRTF_Nb + 1
        arrRTF_Line(arrRTF_Nb) = Trim(wDSPFDY0.ATLIB) & " / " & Trim(wDSPFDY0.ATFILE) & " : " & wDSPFDY0.ATTXT
        arrRTF_Balise(arrRTF_Nb) = 2

       '''' lRTF = lRTF & Asc13 & Chr$(10)
        ''''X = wDSPFDY0.ATLIB & " / " & wDSPFDY0.ATFILE & " / " & wDSPFDY0.ATTXT
        ''''arrBold_Nb = arrBold_Nb + 1
        ''''arrBold_SelStart(arrBold_Nb) = Len(lRTF)
        ''''arrBold_SelLength(arrBold_Nb) = Len(X)

        ''''lRTF = lRTF & X & Asc13 & Chr$(10)

    Call srvDSPFFDY0_frmRTF_Key(lElpKMSrc_Id, lPFile, lWHFILE)
End If

End Sub

Public Sub srvDSPFFDY0_frmRTF_Print_Page()
If blnMinX12 Then
    frmElpPrt.prtNewPage
    prtElpKMPgm_Form
    XPrt.CurrentY = prtMinY + prtHeaderHeight + 50 + prtlineHeight
    blnMinX12 = False
    prtMinMarge = prtMinX1
    prtMaxMarge = prtMedX - 100
Else
    blnMinX12 = True
    XPrt.CurrentY = prtMinY + prtHeaderHeight + 50
    prtMinMarge = prtMedX + 100
    prtMaxMarge = prtMaxX
End If

End Sub
