Attribute VB_Name = "prtSwift"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim X As String, I As Integer, Height8_6 As Integer
Dim I1 As Integer, I2 As Integer, I3 As Integer

Dim wL As Long, wX As String
'---------------------------------------------------------
 Public Sub prtSwift_Close()
'---------------------------------------------------------

On Error GoTo prtError

frmElpPrt.prtEndDoc 1000
frmElpPrt.Hide

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide
End Sub
'---------------------------------------------------------
 Public Sub prtSwift_Open(Msg As String)
'---------------------------------------------------------
Dim X As String

On Error GoTo prtError

Set XPrt = Printer
'Height8_6 = frmElpPrt.prtHeightDelta(8, 6)
prtTitleText = Msg '"Swift: messages reçus " 'lEnTête

prtLineNb = 1

frmElpPrt.Show vbModeless
prtlineHeight = 250 '200
prtHeaderHeight = 300 '250

prtOrientation = vbPRORPortrait

prtPgmName = "prtSwift"
prtTitleUsr = usrName
frmElpPrt.prtStdInit


prtSwift_Form
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide
End Sub
'---------------------------------------------------------
Public Sub prtSwiftMsgFile_Line(lX1 As String, lK As Integer)
'---------------------------------------------------------
Dim K As Integer

If XPrt.CurrentY + prtlineHeight * 1.9 > prtMaxY Then
    frmElpPrt.prtNewPage
    prtSwift_Form
End If

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If lK = 1 Then
    prtCurrentY = XPrt.CurrentY - prtlineHeight + 20
    Call frmElpPrt.prtTrame(prtMinX, prtCurrentY, prtMaxX, prtCurrentY + prtlineHeight, " ", 230)
    XPrt.CurrentY = prtCurrentY + 50
End If

Select Case lK
    Case 1: XPrt.FontBold = True: XPrt.FontSize = 8: XPrt.ForeColor = vbRed
    Case 2: XPrt.FontBold = True: XPrt.FontSize = 8: XPrt.ForeColor = vbBlue
    Case 20: XPrt.FontBold = True: XPrt.FontSize = 7
    Case Else:    XPrt.FontBold = False: XPrt.FontSize = 7 '6
End Select


Select Case lK
    Case 3:
        Select Case Mid$(lX1, 1, 5)
            Case ":32A:": lX1 = prtSwiftMsgFile_Line_32A(lX1)
            Case ":61:": lX1 = prtSwiftMsgFile_Line_61F(lX1)
            Case ":60F:": lX1 = prtSwiftMsgFile_Line_60F(lX1)
            Case ":62F:": lX1 = prtSwiftMsgFile_Line_60F(lX1)
        End Select
        
            
        I2 = Len(lX1)
        I1 = InStr(2, lX1, ":")
        If I1 <= 0 Then I1 = 1
        XPrt.CurrentX = prtMinX + 300: XPrt.Print Mid$(lX1, 1, I1);
        I3 = I2 - I1
        If I3 < 150 Then
            XPrt.CurrentX = prtMinX + 800: XPrt.Print Mid$(lX1, I1 + 1, I2 - I1);
        Else
            For K = I1 + 1 To I2 Step 150
                XPrt.CurrentX = prtMinX + 800: XPrt.Print Mid$(lX1, K, 100);
                XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
            Next K
                XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
        End If
        
    Case 4
        XPrt.CurrentX = prtMinX: XPrt.Print lX1;
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    Case Else: XPrt.CurrentX = prtMinX: XPrt.Print lX1;
End Select
XPrt.ForeColor = vbBlack

End Sub

'---------------------------------------------------------
Public Sub prtSwift_Form()
'---------------------------------------------------------
Dim X As String, K As Integer
XPrt.FontSize = 8

XPrt.FontBold = True
XPrt.DrawWidth = 3

XPrt.Line (prtMinX, prtMinY)-(prtMaxX, prtMinY), prtLineColor

XPrt.CurrentY = prtMinY + 50
End Sub


Public Function prtSwiftMsgFile_Line_32A(lX1 As String) As String
X = num_Display(Mid$(lX1, 15, Len(lX1) - 14), 17, 2, lX, wX, "#")
prtSwiftMsgFile_Line_32A = Mid$(lX1, 1, 5) & Mid$(lX1, 6, 2) & "." & Mid$(lX1, 8, 2) & "." & Mid$(lX1, 10, 2) & "  " & Mid$(lX1, 12, 3) & "  " & X
End Function
Public Function prtSwiftMsgFile_Line_60F(lX1 As String) As String
X = num_Display(Mid$(lX1, 16, Len(lX1) - 15), 17, 2, lX, wX, "#")
prtSwiftMsgFile_Line_60F = Mid$(lX1, 1, 5) & Mid$(lX1, 6, 2) & "    " & Mid$(lX1, 7, 2) & "." & Mid$(lX1, 9, 2) & "." & Mid$(lX1, 11, 2) & "  " & Mid$(lX1, 13, 3) & "  " & X
End Function

Public Function prtSwiftMsgFile_Line_61F(lX1 As String) As String
X = num_Display(Mid$(lX1, 15, Len(lX1) - 14), 17, 2, lX, wX, "#")
prtSwiftMsgFile_Line_61F = Mid$(lX1, 1, 5) & Mid$(lX1, 6, 2) & "." & Mid$(lX1, 7, 2) & "." & Mid$(lX1, 9, 2) & "  " & Mid$(lX1, 11, 2) & "." & Mid$(lX1, 13, 2) & "  " & X
End Function

