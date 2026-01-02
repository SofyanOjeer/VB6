VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmYCDODOS0 
   AutoRedraw      =   -1  'True
   Caption         =   "SPLFJOB : Gestion des spoules"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9090
   Icon            =   "YCDODOS0.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6210
   ScaleWidth      =   9090
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   5100
      TabIndex        =   4
      Top             =   -15
      Width           =   3495
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   30
      TabIndex        =   2
      Top             =   525
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Test"
      TabPicture(0)   =   "YCDODOS0.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraX"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "YCDODOS0.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.Frame fraX 
         Height          =   4875
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   8850
         Begin VB.Frame fraYCDODOS0 
            Height          =   1125
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   8340
            Begin VB.CommandButton cmdYCDODOSDOS 
               Caption         =   "lecture XCDODOSDOS"
               Height          =   735
               Left            =   6000
               TabIndex        =   7
               Top             =   240
               Width           =   1410
            End
            Begin VB.TextBox txtYCDODOSDOS 
               Height          =   285
               Left            =   1920
               TabIndex        =   6
               Top             =   480
               Width           =   2625
            End
            Begin VB.Label lblYCDODOSDOS 
               Caption         =   "no dossier"
               Height          =   255
               Left            =   255
               TabIndex        =   8
               Top             =   435
               Width           =   1305
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   1140
            Left            =   0
            TabIndex        =   9
            Top             =   3480
            Width           =   8730
            _ExtentX        =   15399
            _ExtentY        =   2011
            _Version        =   393216
            Rows            =   1
            Cols            =   7
            FixedCols       =   0
            RowHeightMin    =   200
            BackColor       =   14737632
            ForeColor       =   12582912
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   14737632
            AllowBigSelection=   0   'False
            TextStyle       =   4
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"YCDODOS0.frx":0342
         End
         Begin VB.Label libYCDODOS0 
            Caption         =   "*"
            Height          =   1575
            Left            =   240
            TabIndex        =   10
            Top             =   1800
            Width           =   8175
            WordWrap        =   -1  'True
         End
      End
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   8600
      Picture         =   "YCDODOS0.frx":0445
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   500
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuContextAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuContextQuitter 
         Caption         =   "Quitter"
      End
   End
End
Attribute VB_Name = "frmYCDODOS0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim mHeight_0 As Integer, mWidth_0 As Integer, mHeight_2 As Integer, mWidth_2 As Integer, mWindowState As Integer
Dim arrTag() As Boolean, arrTagNb As Integer
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String, currentAction As String
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim blnMsgBox_Quit As Boolean, blnAddNew As Boolean, blnGlobalControl As Boolean, blnControl As Boolean
Dim X As String, I As Integer, Msg As String, valX As String, X1 As String, V As Variant, curX As Currency, dblX As Double
Dim intReturn As Integer
Dim YCDODOS0Aut As typeAuthorization


Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim meYCDODOS0 As typeYCDODOS0, xYCDODOS0 As typeYCDODOS0
Dim blnError As Boolean
Private Sub fgSelect_Display()

SSTab1.Tab = 1

fgSelect.Visible = True
fgSelect.Clear: fgSelect.Rows = 1: fgSelect_RowDisplay = 0

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Enabled = True


'For xRtvEdition_Index = 1 To xRtvEdition_Nb
'    If blnEdition_Modèle Or mId$(xRtvEdition(xRtvEdition_Index).Memo2, 18, 1) <> "$" Then
'            fgSelect.Rows = fgSelect.Rows + 1
'            fgSelect.Row = fgSelect.Rows - 1
'            fgSelect_DisplayLine
'        End If
'Next xRtvEdition_Index

fgSelect_SortAD = 6
If fgSelect.Rows > 1 Then fgSelect_Sort
fgSelect.LeftCol = 0

End Sub

Public Sub fgSelect_DisplayLine()
On Error Resume Next

End Sub

Public Sub fgSelect_Sort()
If fgSelect.Rows > 1 Then
    fgSelect.Row = 1
    fgSelect.RowSel = fgSelect.Rows - 1
    
    If fgSelect_Sort1_Old = fgSelect_Sort1 Then
        If fgSelect_SortAD = 5 Then
            fgSelect_SortAD = 6
        Else
            fgSelect_SortAD = 5
        End If
    Else
        fgSelect_SortAD = 5
    End If
    fgSelect_Sort1_Old = fgSelect_Sort1
    
    fgSelect.Col = fgSelect_Sort1
    fgSelect.ColSel = fgSelect_Sort2
    fgSelect.Sort = fgSelect_SortAD
End If

End Sub
Public Sub fgSelect_SortX(lK As Integer)
Dim I As Integer, X As String
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = lK
    X = Format$(Val(fgSelect.Text), "0000000")
    fgSelect.Col = fgSelect_arrIndex - 1
    Select Case lK
        Case 1, 2: fgSelect.Text = X
    End Select
Next I


fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
End Sub



'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init(mId$(Msg, 1, 12), YCDODOS0Aut)

'blnSetfocus = True
Form_Init


End Sub


Public Sub Form_Init()
Me.Enabled = False
Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents

If Not IsNull(param_Init) Then
    MsgBox "paramétrage inconsistent", vbCritical, "frmYCDODOS0.param_init"
    Unload Me
End If

blnControl = False
fgSelect_FormatString = fgSelect.FormatString
fgSelect.Enabled = True

cmdReset
Me.Enabled = True

End Sub


'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
blnControl = False
usrColor_Set
cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
currentAction = ""
recYCDODOS0_Init meYCDODOS0
xYCDODOS0 = meYCDODOS0

blnControl = True

End Sub


Public Function param_Init()
Dim K As Integer, K1 As Integer, X As String

Dim V
param_Init = Null


End Function






Public Sub fgSelect_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgSelect.Row

If lRow > 0 And lRow < fgSelect.Rows Then
    fgSelect.Row = lRow
    For I = 0 To fgSelect_arrIndex
        fgSelect.Col = I: fgSelect.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgSelect.Row = mRow
    If fgSelect.Row > 0 Then
        lRow = fgSelect.Row
        lColor_Old = fgSelect.CellBackColor
        For I = 0 To fgSelect_arrIndex
          fgSelect.Col = I: fgSelect.CellBackColor = lColor
        Next I
        fgSelect.Col = 0
    End If
End If

End Sub

'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
currentActiveControl_Name = C.Name
C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor
End Sub


'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------
arrTag(Val(C.Tag)) = True
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor
End Sub

Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdYCDODOSDOS_Click()

meYCDODOS0.Method = "SeekP0"
meYCDODOS0.CDODOSCOP = "CDE"
meYCDODOS0.CDODOSDOS = Val(txtYCDODOSDOS)
srvYCDODOS0_Monitor meYCDODOS0
If Trim(meYCDODOS0.Err) = "" Then srvYCDODOS0_ElpDisplay meYCDODOS0
libYCDODOS0 = MsgTxt
End Sub

Private Sub fgSelect_Click()
fgSelect.LeftCol = 0

End Sub

Private Sub fgSelect_LeaveCell()
On Error Resume Next
fgSelect.CellBackColor = &HE0E0E0
End Sub

Private Sub mnuContextAbandonner_Click()
cmdContext_Quit
End Sub


Private Sub mnuContextQuitter_Click()
Unload Me
End Sub

'---------------------------------------------------------
Private Sub Form_Activate()
'---------------------------------------------------------
Set XForm = Me
End Sub

'---------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

'---------------------------------------------------------
Select Case KeyCode
    Case Is = 13: KeyCode = 0: cmdContext_Return
    Case Is = 27: cmdContext_Quit
'   Case Is = 34: cmdPageNext_Click
'   Case Is = 33: cmdPagePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select


End Sub

Public Sub cmdContext_Quit()
blnControl = False
lstErr.Clear: lstErr.Height = 200

If currentAction = "" Then
   
Else
    X = MsgBox("Voulez-vous réellement abandonner la mise à jour?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption)
    If X = vbYes Then
        currentAction = ""
    Else
        Exit Sub
    End If
End If

End Sub

Public Sub cmdContext_Return()
SendKeys "{TAB}"
End Sub

'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
blnControl = False
fgSelect.Clear: fgSelect.Row = 0
End Sub





Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim wRow As Long
On Error Resume Next
If Y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 1:  fgSelect_SortX 1
        Case 2: fgSelect_SortX 2
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
        Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        If fgSelect.Col = 2 Or fgSelect.Col = 1 Then
            fgSelect_RowClick = fgSelect.Row
            fgSelect_ColClick = fgSelect.Col
            fgSelect.CellBackColor = vbCyan
       End If
    '   fgSelect_MouseDown_Ok
   ' If Button = vbRightButton Then
     '   Me.PopupMenu mnuSelect, vbPopupMenuLeftButton
   End If
End If
End Sub

Public Sub fgSelect_Reset()
fgSelect.Clear
fgSelect_Sort1 = 0: fgSelect_Sort2 = 0
fgSelect_Sort1_Old = -1
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0
fgSelect_arrIndex = 6
blnfgSelect_DisplayLine = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub

Private Sub Form_Resize()
If mWindowState <> Me.WindowState Then
    If Me.WindowState = 0 Or Me.WindowState = 2 Then
        Elp_Form_Resize Me, mWindowState, mHeight_0, mWidth_0, mHeight_2, mWidth_2
    End If
End If

End Sub

Public Sub MouseMoveActiveControl_Set(C As Control)
If MouseMoveActiveControl_Name <> C.Name Then
    MouseMoveActiveControl_Reset
    If Not C.Enabled Then
        MouseMoveActiveControl_Name = ""
    Else
        MouseMoveActiveControl_Name = C.Name
        If TypeOf C Is CommandButton Then
            MouseMoveActiveControl.BackColor = C.BackColor
            C.BackColor = MouseMoveUsr.BackColor
        Else
            If TypeOf C Is ListBox Then
                Elp_ResizeControl C
            Else
                MouseMoveActiveControl.ForeColor = C.ForeColor
                C.ForeColor = MouseMoveUsr.ForeColor
            End If
        End If
    End If
End If

End Sub


Public Sub MouseMoveActiveControl_Reset()
For Each xobj In Me.Controls
    If MouseMoveActiveControl_Name = xobj.Name Then
        MouseMoveActiveControl_Name = ""
        If TypeOf xobj Is CommandButton Then
            xobj.BackColor = MouseMoveActiveControl.BackColor
        Else
            If TypeOf xobj Is ListBox Then
                xobj.Height = 200
            Else
                xobj.ForeColor = MouseMoveActiveControl.ForeColor
            End If
        End If
        Exit For
    End If
Next xobj
End Sub


Public Sub txt_X()
'Call txt_GotFocus(txt)
'KeyAscii = convUCase(KeyAscii)
'Call txt_LostFocus(txt)

'Call txt_GotFocus(txt)
'If XopDevise(2).maxD = 0 Then
'    Call num_KeyAscii(KeyAscii)
'Else
'    Call num_KeyAsciiD(KeyAscii, txt)
'End If
'Call txt_LostFocus(txt)

End Sub



