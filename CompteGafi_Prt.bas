Attribute VB_Name = "prtCompteGafi"


'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim X As String, I As Integer, Height8_6 As Integer
Dim V
Dim Nb1 As Integer, Nb2 As Integer

Dim meCompte As typeCompte
Dim meCptMvt As typeCptMvt
Dim meMvtp0 As typeMvtP0

Dim meCV1 As typeCV, meCV2 As typeCV, meCV3  As typeCV
Dim X1 As String, X2 As String
Dim mLog_Compte      As String * 11
Dim Col4 As Integer, Col5 As Integer, Col6 As Integer, Col7 As Integer, Col8 As Integer

Dim Conversion As String
Dim curEur As Currency, mDevIso As String, curT As Currency
Dim nbDB As Long, nbCR As Long, curDB As Currency, curCR As Currency
Dim paramCompteGafi_Seuil As Currency, paramCompteGafi_curMin As Currency
Dim paramCompteGafi_Etat As String
'---------------------------------------------------------
 Public Sub prtCompteGafi_Open(lMsg As String)
'---------------------------------------------------------

On Error GoTo prtError

Set XPrt = Printer
Height8_6 = frmElpPrt.prtHeightDelta(8, 6)


prtTitleText = "Dispositif de lutte contre le blanchiment : " & lMsg

prtLineNb = 1

frmElpPrt.Show vbModeless


prtOrientation = vbPRORLandscape
prtPgmName = "prtCompteGafi"
prtTitleUsr = usrName

prtlineHeight = 250
prtHeaderHeight = 300

frmElpPrt.prtStdInit

recCompteInit meCompte
meCV1 = CV_Euro
meCV1.CoursCompta = "C"
meCV1.OpéAmj = DSys
meCV1.Normal = "P"
meCV1.AchatVente = " "
meCV2 = meCV1: meCV3 = meCV1
Col4 = 7000: Col5 = 8700: Col6 = 10300
prtCompteGafi_Form

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "prtCompteGafi_Open")
frmElpPrt.Hide
End Sub
'---------------------------------------------------------
 Public Sub prtCompteGafi_Close()
'---------------------------------------------------------
                        
On Error GoTo prtError
        
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.CurrentY = XPrt.CurrentY + 50
XPrt.FontBold = True
XPrt.CurrentX = prtMinX: XPrt.Print Nb1 & " mouvements";
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)

DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub

frmElpPrt.prtEndDoc
frmElpPrt.Hide

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "prtCompteGafi_Close")
frmElpPrt.Hide
End Sub

'---------------------------------------------------------
Public Sub prtCompteGafi_Form()
'---------------------------------------------------------
Dim X As String
XPrt.DrawWidth = 3
XPrt.FontSize = 8
Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B", 235)

'---------------------------------------------------------


XPrt.DrawWidth = 1
XPrt.Line (Col4, prtMinY)-(Col4, prtMaxY)
XPrt.Line (Col6, prtMinY)-(Col6, prtMaxY)
XPrt.CurrentY = prtMinY + 50

XPrt.FontBold = True
XPrt.FontBold = True
XPrt.CurrentX = 400
XPrt.Print "Compte";

XPrt.CurrentX = 1600
XPrt.Print "Intitulé";

XPrt.CurrentX = 8200
XPrt.Print "Débit";

XPrt.CurrentX = 9600
XPrt.Print "Crédit";

XPrt.FontSize = 6

XPrt.CurrentX = prtMaxX - 2800: XPrt.Print "Date Opé";

XPrt.CurrentX = prtMaxX - 3800: XPrt.Print "Date Valeur";

XPrt.CurrentX = 11000
XPrt.FontItalic = True
XPrt.Print "cv/EUR";
XPrt.FontItalic = False

XPrt.CurrentX = prtMaxX - 2000: XPrt.Print "Service";
XPrt.CurrentX = prtMaxX - 1300: XPrt.Print "Référence";
XPrt.CurrentX = prtMaxX - 500: XPrt.Print "Pièce";

'---------------------------------------------------------

XPrt.CurrentY = prtMinY + prtHeaderHeight - XPrt.TextHeight("X")

XPrt.FontSize = 8

End Sub

Public Sub prtCompteGafi_01_MVTP0()
Dim iReturn As Integer, blnNéant As Boolean

mdbMvtP0.tableMvtP0_Open


meMvtp0.Method = "MoveFirst" '

Do
    iReturn = tableMvtP0_Read(meMvtp0)
    If iReturn = 0 Then
        curEur = CCur(Val(mId$(meMvtp0.Text, 1, 19)))
        mDevIso = mId$(meMvtp0.Text, 20, 3)
        MsgTxt = Space$(recCptMvtLen)
        Mid$(MsgTxt, 35, memoCptMvtLen) = mId$(meMvtp0.Text, 23, memoCptMvtLen)
        MsgTxtIndex = 0
        srvCptMvtGetBuffer meCptMvt
        
        prtCompteGafi_Line meCptMvt
        
    End If
    meMvtp0.Method = "MoveNext"
Loop Until iReturn <> 0


mdbMvtP0.tableMvtP0_Close

End Sub

Public Sub prtCompteGafi_02_MVTP0()
Dim iReturn As Integer, blnNéant As Boolean
Dim mIdMin As String, mIdMax As String, mCompte As String
mdbMvtP0.tableMvtP0_Open


meMvtp0.Method = "MoveFirst" '
prtCompteGafi_02_Z
mIdMin = ""
Do
    iReturn = tableMvtP0_Read(meMvtp0)
    If iReturn = 0 Then
        If mCompte <> mId$(meMvtp0.Id, 1, 11) Then
            If curT > paramCompteGafi_Seuil Then prtCompteGafi_02_Compte mIdMin
            mCompte = mId$(meMvtp0.Id, 1, 11)
            mIdMin = meMvtp0.Id
            prtCompteGafi_02_Z
        End If

        curEur = CCur(Val(mId$(meMvtp0.Text, 1, 19)))
        curT = curT + Abs(curEur)
        If Abs(curEur) < paramCompteGafi_curMin Then
            If curEur < 0 Then
                nbDB = nbDB + 1: curDB = curDB + curEur
            Else
                 nbCR = nbCR + 1: curCR = curCR + curEur
           End If
        End If
    End If
    meMvtp0.Method = "MoveNext"
Loop Until iReturn <> 0

If curT > paramCompteGafi_Seuil Then prtCompteGafi_02_Compte mIdMin

mdbMvtP0.tableMvtP0_Close

End Sub

Public Sub prtCompteGafi_02_Compte(lId As String)
Dim iReturn As Integer
Dim mCompte As String

meMvtp0.Method = "Seek=" '
meMvtp0.Id = lId
mCompte = mId$(lId, 1, 11)
Do
    iReturn = tableMvtP0_Read(meMvtp0)
    If iReturn = 0 Then
        If mCompte <> mId$(meMvtp0.Id, 1, 11) Then Exit Do
        
        curEur = CCur(Val(mId$(meMvtp0.Text, 1, 19)))
        mDevIso = mId$(meMvtp0.Text, 20, 3)
        MsgTxt = Space$(recCptMvtLen)
        Mid$(MsgTxt, 35, memoCptMvtLen) = mId$(meMvtp0.Text, 23, memoCptMvtLen)
        MsgTxtIndex = 0
        srvCptMvtGetBuffer meCptMvt
                
        If Abs(curEur) >= paramCompteGafi_curMin Then prtCompteGafi_Line meCptMvt

        
    End If
    meMvtp0.Method = "MoveNext"
Loop Until iReturn <> 0


End Sub

'---------------------------------------------------------
Public Sub prtCompteGafi_Line(lCptMvt As typeCptMvt)
'---------------------------------------------------------
Dim lX As Long, lMax As Long, blnPrintCompte As Boolean

If XPrt.CurrentY + prtlineHeight * 2.5 > prtMaxY Then
    frmElpPrt.prtNewPage
    prtCompteGafi_Form
End If

Nb1 = Nb1 + 1

XPrt.FontSize = 8
XPrt.FontBold = False
'_______________________________________________________________ligne 1-
Dim X As String

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 50
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6
blnPrintCompte = False

If meCompte.Numéro <> lCptMvt.Compte Or meCompte.Devise <> lCptMvt.Devise Then
    meCompte.Devise = lCptMvt.Devise
    meCompte.Numéro = lCptMvt.Compte
    
    mdbCptP0_Find meCompte
    blnPrintCompte = True
End If

If blnPrintCompte Then
    Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, Col4 - 20, XPrt.CurrentY + prtlineHeight, " ", 235)
    XPrt.CurrentY = XPrt.CurrentY + 100
    XPrt.FontBold = True
    XPrt.CurrentX = prtMinX + 50
    XPrt.Print Format$(meCptMvt.Devise, "000") & "  ";
    XPrt.Print Compte_Imp(meCptMvt.Compte);
    XPrt.CurrentX = prtMinX + 1400
    XPrt.Print meCompte.Intitulé;
    XPrt.FontBold = False
    
    If paramCompteGafi_Etat = "02" Then
        XPrt.FontItalic = True
        X = Format$(curT, "## ### ### ### ### ##0.00")
        XPrt.CurrentX = Col4 - 50 - XPrt.TextWidth(X)
        XPrt.Print X;
        If curDB <> 0 Then
            X = Format$(curDB, "## ### ### ### ### ##0.00")
            XPrt.CurrentX = Col5 - 50 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        If curCR <> 0 Then
            X = Format$(curCR, "## ### ### ### ### ##0.00")
            XPrt.CurrentX = Col6 - 50 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        XPrt.CurrentX = Col6 + 100
        If nbDB > 0 Then XPrt.Print nbDB & " mvts au débit   ";
        If nbCR > 0 Then XPrt.Print nbCR & " mvts au crédit ";
       
        XPrt.FontItalic = False
    End If

    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    
End If



'XPrt.FontSize = 8
'XPrt.CurrentY = XPrt.CurrentY - Height8_6

XPrt.FontBold = True

X = Format$(Abs(meCptMvt.MT), "## ### ### ### ### ##0.00")
XPrt.CurrentX = IIf(meCptMvt.MT < 0, Col5, Col6) - 50 - XPrt.TextWidth(X)
XPrt.Print X;

XPrt.FontBold = False

If meCptMvt.Devise <> "978" Then
'    meCV1.DeviseIso = ""
'    meCV1.DeviseN = meCptMvt.Devise
'    meCV1.Montant = meCptMvt.Mt
'    meCV1.OpéAmj = meCptMvt.AmjOpération
'    meCV2.OpéAmj = meCptMvt.AmjOpération
'    Call CV_Transitoire(meCV1, meCV2, meCV3, Conversion)
'    X = Format$(meCV2.Montant, "## ### ### ### ### ##0.00")
    XPrt.FontItalic = True
    X = Format$(curEur, "## ### ### ### ### ##0.00")
    XPrt.CurrentX = Col6 + 1200 - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.FontItalic = False
    XPrt.CurrentX = Col4 - 400
    XPrt.Print mDevIso;  'meCV1.DeviseIso;

End If

XPrt.CurrentX = prtMaxX - 3000: XPrt.Print dateImp(meCptMvt.AmjOpération);

XPrt.CurrentX = prtMaxX - 4000
If meCptMvt.AmjValeur <> meCptMvt.AmjOpération Then XPrt.Print dateImp(meCptMvt.AmjValeur);


XPrt.CurrentX = prtMinX + 1400
XPrt.Print meCptMvt.Libellé;

    XPrt.CurrentX = prtMaxX - 1800: XPrt.Print meCptMvt.Service;
    XPrt.CurrentX = prtMaxX - 1300: XPrt.Print meCptMvt.OpérateurSaisie;
    XPrt.CurrentX = prtMaxX - 500: XPrt.Print Format$(meCptMvt.Pièce, "0000-") & Format$(meCptMvt.Ligne, "000");

XPrt.CurrentY = XPrt.CurrentY - Height8_6
End Sub



Public Sub prtCompteGafi_Monitor(lEtat As String, lcurDB As Currency, lcurCR As Currency, lAmjMin As String, lAmjMax As String)
Dim X As String
paramCompteGafi_Seuil = lcurCR
paramCompteGafi_curMin = Abs(lcurDB)
paramCompteGafi_Etat = lEtat
X = Format$(paramCompteGafi_Seuil, "### ### ### ##0.00") & " Eur -  (du : " & dateImp10(lAmjMin) & "  au : " & dateImp10(lAmjMax) & ")"

Select Case paramCompteGafi_Etat
    Case "01":
        prtCompteGafi_Open "Mvt > " & X
        prtCompteGafi_01_MVTP0
    Case "02":
        prtCompteGafi_Open "cumul (SOBF & ORPA) > " & X
       prtCompteGafi_02_MVTP0
End Select
prtCompteGafi_Close
End Sub

Public Sub prtCompteGafi_02_Z()
curT = 0
nbDB = 0: curDB = 0
nbCR = 0: curCR = 0

End Sub
