Attribute VB_Name = "prtElpKM"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim X As String, I As Integer, Height8_6 As Integer

Dim blnPage As Boolean

Dim col1 As Integer, col2 As Integer, col3 As Integer, Col4 As Integer, Col5 As Integer
Dim Line1 As Integer, Line2 As Integer, Line3 As Integer

Dim meZMNURUT0 As typeZMNURUT0, meZMNUUTI0 As typeZMNUUTI0
Dim grpZMNURUT0 As typeZMNURUT0, grpZMNURUT0_Rupture As typeZMNURUT0
'---------------------------------------------------------
Public Sub prtElpKM_fgSelect_Form(lcolX() As Integer)
'---------------------------------------------------------
Dim X As String

XPrt.FontBold = False
XPrt.DrawWidth = 1

XPrt.Line (lcolX(3) - 50, prtMinY)-(lcolX(3) - 50, prtMaxY), prtLineColor
'---------------------------------------------------------
XPrt.CurrentY = prtMinY + prtHeaderHeight ''- prtlineHeight


End Sub

'---------------------------------------------------------
Public Sub prtElpKMPgm_Form()
'---------------------------------------------------------
Dim X As String

XPrt.FontBold = True
XPrt.DrawWidth = 1

XPrt.Line (prtMedX, prtMinY)-(prtMedX, prtMaxY), prtLineColor
'---------------------------------------------------------
XPrt.CurrentY = prtMinY + prtHeaderHeight + 50


End Sub

'---------------------------------------------------------
Public Sub prtElpKM_UserGroup_Form()
'---------------------------------------------------------
Dim X As String

XPrt.DrawWidth = 1
XPrt.FontSize = 7: XPrt.FontBold = True

Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B", 240)
'---------------------------------------------------------
XPrt.CurrentY = prtMinY + 50

XPrt.CurrentX = col1: XPrt.Print "Utilisateur";
XPrt.CurrentX = col2: XPrt.Print "Groupe Menu";

XPrt.CurrentX = col3: XPrt.Print "   Code";
XPrt.CurrentX = col3 + 1000 - 100: XPrt.Print "Log";
XPrt.CurrentX = col3 + 1200: XPrt.Print "Grp Droits";
XPrt.CurrentX = col3 + 2100: XPrt.Print "Grp Métier";
XPrt.CurrentX = col3 + 3500: XPrt.Print "Outq";
XPrt.CurrentX = col3 + 4300: XPrt.Print "Attn";
XPrt.CurrentX = col3 + 4700: XPrt.Print "Agence/Service..";
XPrt.CurrentY = prtMinY + 50 + prtHeaderHeight
XPrt.FontBold = False

End Sub

Public Sub prtElpKM_Open()
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

'''prtFontName = prtFontNameZ '"Comic Sans MS" '"Century Gothic"
prtFontSize = 7
prtOrientation = vbPRORLandscape 'Portrait '
prtPgmName = "prtElpKM"
prtTitleUsr = usrName '

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 50


prtFormType = ""
frmElpPrt.prtStdInit
col1 = prtMinX
col3 = (prtMaxX - prtMinX) / 2 + prtMinX
col2 = col3 - 3000

Line1 = prtMinY + prtHeaderHeight + prtlineHeight * 4
Line2 = prtMaxY

blnPage = False

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub
Public Sub prtStd_Open()
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 50 ' 100


prtFormType = ""
frmElpPrt.prtStdInit

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

Public Sub prtStd_Close()
On Error GoTo prtError


Call frmElpPrt.prtEndDoc(1000)
frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub


Public Sub prtElpKm_UserGroup(lstW As ListBox, lFct As Integer)
Dim blnUserGroup As Boolean
Dim V, xSQL As String
Dim kUTI As Integer, kGR2 As Integer

Dim mMNUUTIGR3 As String

Select Case lFct
   Case 1:  prtTitleText = "SAB : liste des utilisateurs / groupes"
   Case 2:  prtTitleText = "SAB : liste des groupes / utilisateurs"
End Select

prtFontName = prtFontName_CourierNew
prtOrientation = vbPRORLandscape 'Portrait '
prtPgmName = "prtElpKMPgm"
prtTitleUsr = usrName

prtStd_Open
prtHeaderHeight = 300
grpZMNURUT0_Rupture.MNURUTUTI = ""

Select Case lFct
    Case 1:  blnUserGroup = True
                kUTI = 1: kGR2 = 11
                col1 = prtMinX + 50
                col2 = prtMinX + 5000
    
    Case 2:     blnUserGroup = False
                kUTI = 11: kGR2 = 1
                col1 = prtMinX + 5000
                col2 = prtMinX + 50
End Select

col3 = 10000

prtElpKM_UserGroup_Form
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight

XPrt.FontSize = 7
For I = 0 To lstW.ListCount - 1
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    If XPrt.CurrentY + 200 > prtMaxY Then
        frmElpPrt.prtNewPage
        prtElpKM_UserGroup_Form
    End If
    
    lstW.ListIndex = I
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZMNURUT0," & paramIBM_Library_SAB & ". ZMNUUTI0" _
     & " where MNURUTUTI = '" & Mid$(lstW.Text, kUTI, 10) & "'" _
     & " and MNURUTCUT = MNUUTICUT "
     
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        V = rsZMNURUT0_GetBuffer(rsSab, meZMNURUT0)
        V = rsZMNUUTI0_GetBuffer(rsSab, meZMNUUTI0)
        
       ' xSql = "select * from " & paramIBM_Library_SAB & ".ZMNURUT0" _
       '  & " where MNURUTUTI = '" & Mid$(lstW.Text, kGR2, 10) & "'"
         
        'Set rsSab = cnsab.Execute(xSql)
        'If Not rsSab.EOF Then
        '    V = rsZMNURUT0_GetBuffer(rsSab, grpZMNURUT0)
        'Else
        '    grpZMNURUT0.MNURUTUTI = "??????"
        'End If
    End If
    
    If XPrt.CurrentY + 500 > prtMaxY Then
        frmElpPrt.prtNewPage
        prtElpKM_UserGroup_Form
    End If
    If blnUserGroup Then
                
        XPrt.FontSize = 7: XPrt.FontBold = True
        XPrt.CurrentX = col2: XPrt.Print grpZMNURUT0.MNURUTUTI;
        XPrt.FontSize = 6: XPrt.FontBold = False
        XPrt.CurrentX = col2 + 1000: XPrt.Print grpZMNURUT0.MNURUTNOM;
    Else
        If mMNUUTIGR3 <> meZMNUUTI0.MNUUTIGR3 Then
            XPrt.CurrentY = XPrt.CurrentY + 100
            Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight + 50, " ", 240)
            XPrt.CurrentY = XPrt.CurrentY + 50
            XPrt.FontSize = 7: XPrt.FontBold = True
            XPrt.CurrentX = col2: XPrt.Print meZMNUUTI0.MNUUTIGR3;
            XPrt.FontSize = 6: XPrt.FontBold = False
            'XPrt.CurrentX = col2 + 1000: XPrt.Print grpZMNURUT0.MNURUTNOM;
            mMNUUTIGR3 = meZMNUUTI0.MNUUTIGR3
        End If
    End If
        
    
    XPrt.FontSize = 7: XPrt.FontBold = True
    XPrt.CurrentX = col1: XPrt.Print meZMNURUT0.MNURUTUTI;
    XPrt.FontSize = 6: XPrt.FontBold = False
    XPrt.CurrentX = col1 + 1500: XPrt.Print meZMNURUT0.MNURUTNOM;
    
    XPrt.FontSize = 7: XPrt.FontBold = False
    XPrt.CurrentX = col3 + 500 - XPrt.TextWidth(meZMNURUT0.MNURUTCUT): XPrt.Print meZMNURUT0.MNURUTCUT;
    XPrt.CurrentX = col3 + 1000: XPrt.Print meZMNURUT0.MNURUTLOG;
    XPrt.CurrentX = col3 + 1200: XPrt.Print meZMNUUTI0.MNUUTIGR3;
    XPrt.CurrentX = col3 + 2400: XPrt.Print meZMNUUTI0.MNUUTIGR4;
    XPrt.CurrentX = col3 + 3600: XPrt.Print meZMNUUTI0.MNUUTIOUT;
    XPrt.CurrentX = col3 + 4600: XPrt.Print meZMNUUTI0.MNUUTIMSE;
    XPrt.CurrentX = col3 + 5100: XPrt.Print meZMNUUTI0.MNUUTIAGE & "-" & meZMNUUTI0.MNUUTISER & "-" & meZMNUUTI0.MNUUTISRV;
Next I

prtStd_Close


End Sub
