VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmDAFI 
   AutoRedraw      =   -1  'True
   Caption         =   "Dafi : Impression"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6375
   ScaleWidth      =   9420
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   6120
      TabIndex        =   2
      Top             =   0
      Width           =   2745
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   8880
      Picture         =   "DAFI.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   500
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Echéancier"
      TabPicture(0)   =   "DAFI.frx":0102
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fgEchéancier"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraEchéancier"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Options"
      TabPicture(1)   =   "DAFI.frx":011E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstService"
      Tab(1).Control(1)=   "lstDevise"
      Tab(1).Control(2)=   "chkDevise"
      Tab(1).Control(3)=   "chkService"
      Tab(1).ControlCount=   4
      Begin VB.CheckBox chkService 
         Caption         =   "sélectionner un service"
         Height          =   375
         Left            =   -69720
         TabIndex        =   14
         Top             =   600
         Width           =   3375
      End
      Begin VB.CheckBox chkDevise 
         Caption         =   "sélectionner une devise"
         Height          =   375
         Left            =   -74520
         TabIndex        =   13
         Top             =   600
         Width           =   3375
      End
      Begin VB.ListBox lstDevise 
         Height          =   4545
         Left            =   -74520
         TabIndex        =   12
         Top             =   1080
         Width           =   3400
      End
      Begin VB.ListBox lstService 
         Height          =   4545
         Left            =   -69720
         TabIndex        =   11
         Top             =   1080
         Width           =   3400
      End
      Begin VB.Frame fraEchéancier 
         Caption         =   "Sélection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   4
         Top             =   4560
         Width           =   9135
         Begin VB.CommandButton cmdRechercher 
            BackColor       =   &H00C0FFC0&
            Caption         =   "&Rechercher"
            Height          =   855
            Left            =   6720
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   240
            Width           =   2175
         End
         Begin MSComCtl2.DTPicker txtAmjMin 
            Height          =   300
            Left            =   1320
            TabIndex        =   6
            Top             =   600
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            CalendarBackColor=   16777215
            CalendarForeColor=   0
            CalendarTitleBackColor=   8421504
            CalendarTitleForeColor=   16777215
            CalendarTrailingForeColor=   12632256
            CustomFormat    =   "dd  MM yyy"
            Format          =   69271555
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin MSComCtl2.DTPicker txtAmjMax 
            Height          =   300
            Left            =   3480
            TabIndex        =   7
            Top             =   600
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            CalendarBackColor=   16777215
            CalendarForeColor=   0
            CalendarTitleBackColor=   8421504
            CalendarTitleForeColor=   16777215
            CalendarTrailingForeColor=   12632256
            CustomFormat    =   "dd  MM yyy"
            Format          =   69271555
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin VB.Label lblAmjMAx 
            Caption         =   "au"
            Height          =   255
            Left            =   2880
            TabIndex        =   9
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lblAmjMin 
            Caption         =   "du"
            Height          =   255
            Left            =   600
            TabIndex        =   8
            Top             =   720
            Width           =   495
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgEchéancier 
         Height          =   4050
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   7144
         _Version        =   393216
         Rows            =   1
         Cols            =   10
         FixedCols       =   0
         RowHeightMin    =   350
         BackColor       =   14737632
         ForeColor       =   12582912
         ForeColorFixed  =   -2147483641
         BackColorSel    =   12648384
         BackColorBkg    =   14737632
         AllowBigSelection=   0   'False
         TextStyleFixed  =   4
         FocusRect       =   2
         HighLight       =   0
         GridLines       =   2
         FormatString    =   $"DAFI.frx":013A
      End
   End
End
Attribute VB_Name = "frmDAFI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim arrTag() As Boolean, arrTagNb As Integer
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String, currentAction As String
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim blnMsgBox_Quit As Boolean, blnAddNew As Boolean, blnGlobalControl As Boolean, blnControl As Boolean
Dim X As String, I As Integer, Msg As String, valX As String, X1 As String

Dim XEmploi As typeEmploi
Dim wAmjMin As String, wAmjMax As String
Dim fgEchéancier_FormatString As String, fgEchéancier_K As Integer

Dim CV As typeCV
Dim mService As String, mDevise As String
Private Sub frmCompte_Show(XDevise As String, xNuméro As String)
X = Space$(100)
Mid$(X, 1, 12) = "frmCompte   "
Mid$(X, 13, 12) = "frmDAFI     "
Mid$(X, 25, 10) = Space$(10)
Mid$(X, 35, 3) = XDevise
Mid$(X, 38, 11) = xNuméro
Mid$(X, 49, 1) = "V"
'blnControl = False
Msg_Monitor X

End Sub

Private Sub fgEchéancier_Load()
Dim K2 As Integer, I As Integer
Dim curDB As Currency, curCR As Currency, curX As Currency

SSTab1.Tab = 0

fgEchéancier_Init
fgEchéancier.Visible = True
fgEchéancier.Clear

fgEchéancier.Rows = 1
fgEchéancier.FormatString = fgEchéancier_FormatString
fgEchéancier.Enabled = True
For I = 1 To arrEmploiNb
    fgEchéancier.Rows = fgEchéancier.Rows + 1
    fgEchéancier.Row = fgEchéancier.Rows - 1

    fgEchéancier_K = (fgEchéancier.Row) * fgEchéancier.Cols
    
    CV.DeviseN = arrEmploi(I).Devise: CV_AttributN CV
 
    fgEchéancier.TextArray(0 + fgEchéancier_K) = dateImp(arrEmploi(I).AmjEchéance) & arrEmploi(I).TagEchéance
    fgEchéancier.TextArray(1 + fgEchéancier_K) = CV.DeviseIso
    fgEchéancier.TextArray(2 + fgEchéancier_K) = Compte_Imp(arrEmploi(I).Compte)
    fgEchéancier.TextArray(3 + fgEchéancier_K) = Trim(arrEmploi(I).Intitulé) & " " & Trim(arrEmploi(I).Intitulé2)
    fgEchéancier.TextArray(4 + fgEchéancier_K) = Format(arrEmploi(I).Capital, "#### ### ###.00")
    fgEchéancier.TextArray(5 + fgEchéancier_K) = Format(arrEmploi(I).Taux, "##.000000")
    fgEchéancier.TextArray(6 + fgEchéancier_K) = Format(arrEmploi(I).Intérêts, "#### ### ###.00")
    fgEchéancier.TextArray(7 + fgEchéancier_K) = Format(arrEmploi(I).NbjBase, "0")
    fgEchéancier.TextArray(8 + fgEchéancier_K) = dateImp(arrEmploi(I).AmjEchéance)
    fgEchéancier.TextArray(9 + fgEchéancier_K) = Format(arrEmploi(I).NbjCouru, "#### ### ###")

Next I
If fgEchéancier.Rows = 1 Then Exit Sub
'fgEchéancier_Sort

End Sub
Public Sub fgEchéancier_Sort()
fgEchéancier.Row = 1
fgEchéancier.RowSel = fgEchéancier.Rows - 1

fgEchéancier.Col = 0
fgEchéancier.ColSel = 0
fgEchéancier.Sort = 1

End Sub

Public Sub fgEchéancier_Init()

ReDim arrEmploi(1): arrEmploiNbMax = 0
arrEmploiSuite = True: arrEmploiNb = 0

recEmploi_Init XEmploi
XEmploi.Method = "SnapLE"
XEmploi.Société = SocId$
XEmploi.Agence = SocAgence$
XEmploi.Devise = "000"
XEmploi.Compte = "00000000000"
XEmploi.AmjEchéance = wAmjMin

If chkService.Enabled Then
    If chkService = "1" Then
        mService = mId$(lstService.Text, 1, 3)
    Else
        mService = "   "
    End If
End If

If chkDevise = "1" Then
    mDevise = mId$(lstDevise.Text, 1, 3)
Else
    mDevise = "   "
End If

XEmploi.Intitulé = mService & mDevise

arrEmploi(0) = XEmploi
arrEmploi(0).Devise = "999"
arrEmploi(0).Compte = "99999999999"
arrEmploi(0).AmjEchéance = wAmjMax


Do Until Not arrEmploiSuite
    srvEmploi_Monitor XEmploi
    XEmploi = arrEmploi(arrEmploiNb)
    XEmploi.Method = "SnapLE+"
Loop

End Sub

Public Sub MouseMoveActiveControl_Reset()
For Each xobj In Me.Controls
    If MouseMoveActiveControl_Name = xobj.Name Then
        MouseMoveActiveControl_Name = ""
         If TypeOf xobj Is CommandButton Or TypeOf xobj Is ListBox Then
           xobj.BackColor = MouseMoveActiveControl.BackColor
        Else
            xobj.ForeColor = MouseMoveActiveControl.ForeColor
        End If
        Exit For
    End If
Next xobj

End Sub

Public Sub MouseMoveActiveControl_Set(C As Control)
If MouseMoveActiveControl_Name <> C.Name Then
    MouseMoveActiveControl_Reset
    If Not C.Enabled Then
        MouseMoveActiveControl_Name = ""
    Else
        MouseMoveActiveControl_Name = C.Name
        If TypeOf C Is CommandButton Or TypeOf C Is ListBox Then
            
            MouseMoveActiveControl.BackColor = C.BackColor
            C.BackColor = MouseMoveUsr.BackColor
        Else
            MouseMoveActiveControl.ForeColor = C.ForeColor
            C.ForeColor = MouseMoveUsr.ForeColor
        End If
    End If
End If

End Sub

Public Sub cmdContext_Quit()
blnControl = False
If blnMsgBox_Quit Then
    X = MsgBox("Voulez-vous réellement abandonner?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
 Else
    X = vbYes
 End If
 If X = vbYes Then Unload Me

End Sub


Public Sub cmdContext_Return()

    SendKeys "{TAB}"

End Sub






'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
'lstErr.Clear
currentActiveControl_Name = C.Name
C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor
End Sub
'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------
'lstErr.Clear
arrTag(Val(C.Tag)) = True
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor
End Sub

Private Sub chkDevise_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkDevise
End Sub


Private Sub chkService_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkService
End Sub


Private Sub cmdContext_Click()
Select Case cmdContext.Caption
    Case Is = constcmdRechercher
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

'---------------------------------------------------------
Private Sub cmdPrint_Click()
'---------------------------------------------------------

If arrEmploiNb > 0 Then prtEmploi_Monitor "Echéancier DAFI "
End Sub


'---------------------------------------------------------
Private Sub cmdQuit_Click()
'---------------------------------------------------------
Unload Me

End Sub




Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdContext
End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdPrint
End Sub


Private Sub cmdRechercher_Click()
txtAmjMin_Control
txtAmjMax_Control
If wAmjMax < wAmjMin Then
    Call lstErr_AddItem(lstErr, cmdContext, "? erreur date au > du")
Else
    fgEchéancier_Load
End If
End Sub

Private Sub cmdRechercher_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdRechercher
End Sub


Private Sub fgEchéancier_Click()
If fgEchéancier.Rows > 1 Then
    fgEchéancier_K = (fgEchéancier.Row) * fgEchéancier.Cols
    Call frmCompte_Show(arrEmploi(fgEchéancier.Row).Devise _
                        , arrEmploi(fgEchéancier.Row).Compte)
End If

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


'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
fgEchéancier_FormatString = fgEchéancier.FormatString

Form_Init " "

End Sub
'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
usrColor_Set
arrEmploiNb = 0
cmdContext.Caption = constcmdAbandonner: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
CV = CV_Euro
End Sub



'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub



Public Sub Msg_Rcv(X As String)
'---------------------------------------------------------
End Sub


Public Sub Msg_Snd(ByVal X As String)
End Sub




Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub

Private Sub fraEchéancier_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Public Sub Form_Init(Msg As String)
cmdReset

wAmjMin = DSys
Call DTPicker_Set(txtAmjMin, wAmjMin)
wAmjMax = dateElp("Jour", 15, DSys)
Call DTPicker_Set(txtAmjMax, wAmjMax)

mDevise = "   "
Call LstDictio(888, lstDevise)
lstDevise.ListIndex = 0

mService = usrService
lstService.Enabled = False
chkService.Enabled = False
chkService.Value = "1"
recElpTable_Init recElpTable
recElpTable.Method = "Seek="
recElpTable.Id = "Param"
recElpTable.K1 = "DAFI"
recElpTable.K2 = usrId
recElpTable.Err = tableElpTable_Read(recElpTable)
If recElpTable.Err = 0 Then
    Call LstDictio(4, lstService)
    chkService.Value = "0"
    chkService.Enabled = True
    lstService.Enabled = True
    lstService.ListIndex = 0
Else
    lstService.AddItem usrService
End If

End Sub
'-------------------------------------------------'
Private Sub txtAmjMin_Control()
'-------------------------------------------------'

Dim X As String
X = Format$(txtAmjMin.Year, "0000") & Format$(txtAmjMin.Month, "00") & Format$(txtAmjMin.Day, "00")
If Not IsNumeric(X) Then
    Call lstErr_AddItem(lstErr, cmdContext, "? erreur date")
    DTPicker_Now txtAmjMin
Else
    wAmjMin = mId$(X, 1, 8)
End If

End Sub

'-------------------------------------------------'
Private Sub txtAmjMax_Control()
'-------------------------------------------------'

Dim X As String
X = Format$(txtAmjMax.Year, "0000") & Format$(txtAmjMax.Month, "00") & Format$(txtAmjMax.Day, "00")
If Not IsNumeric(X) Then
    Call lstErr_AddItem(lstErr, cmdContext, "? erreur date")
    DTPicker_Now txtAmjMax
Else
    wAmjMax = mId$(X, 1, 8)
End If

End Sub


Private Sub lstService_Click()
mService = mId$(lstService.Text, 1, 3)

End Sub

Private Sub txtAmjMax_Change()
txtAmjMax_Control

End Sub


Private Sub txtAmjMax_GotFocus()
DTPicker_GotFocus txtAmjMax
End Sub


Private Sub txtAmjMax_LostFocus()
DTPicker_LostFocus txtAmjMax
txtAmjMax_Control

End Sub

Private Sub txtAmjMax_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DTPicker_GotFocus txtAmjMax
End Sub

Private Sub txtAmjMin_Change()
txtAmjMin_Control
End Sub


Private Sub txtAmjMin_GotFocus()
DTPicker_GotFocus txtAmjMin

End Sub


Private Sub txtAmjMin_LostFocus()
DTPicker_LostFocus txtAmjMin
txtAmjMin_Control

End Sub


Private Sub txtAmjMin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DTPicker_GotFocus txtAmjMin

End Sub


