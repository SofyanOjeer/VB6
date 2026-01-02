VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmBOTC 
   AutoRedraw      =   -1  'True
   Caption         =   "BOTC : Impression"
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
      Picture         =   "BOTC.frx":0000
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
      TabCaption(0)   =   "Position d'arbitrage"
      TabPicture(0)   =   "BOTC.frx":0102
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fgCptBalance"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraCptBalance"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "..."
      TabPicture(1)   =   "BOTC.frx":011E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.Frame fraCptBalance 
         Caption         =   "Impression"
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
         Begin VB.CommandButton cmdPrintX 
            Caption         =   "&Imprimer"
            Height          =   855
            Left            =   6720
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Width           =   2175
         End
         Begin VB.CheckBox chkArbitragePositionPrint 
            Caption         =   "position d'arbitrage"
            Height          =   195
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Value           =   1  'Checked
            Width           =   2535
         End
         Begin VB.CheckBox chkArbitrageMvtPrint 
            Caption         =   "mouvements position d'arbitrage"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   840
            Value           =   1  'Checked
            Width           =   2655
         End
         Begin MSComCtl2.DTPicker txtAmjMin 
            Height          =   300
            Left            =   3600
            TabIndex        =   8
            Top             =   720
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
            Left            =   5280
            TabIndex        =   9
            Top             =   720
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
         Begin VB.Label lblArbitrageAmjMAx 
            Caption         =   "au"
            Height          =   255
            Left            =   4800
            TabIndex        =   11
            Top             =   840
            Width           =   495
         End
         Begin VB.Label lblArbitrageAmjMin 
            Caption         =   "du"
            Height          =   255
            Left            =   3120
            TabIndex        =   10
            Top             =   840
            Width           =   495
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgCptBalance 
         Height          =   4050
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   7144
         _Version        =   393216
         Rows            =   1
         Cols            =   7
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
         FormatString    =   $"BOTC.frx":013A
      End
   End
End
Attribute VB_Name = "frmBOTC"
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

Dim Xcompte As typeCompte, xCompteMax As typeCompte
Dim CV As typeCV
Dim wCV1 As typeCV, wCV2 As typeCV, wCV3 As typeCV
Dim xConversion As String * 1

Dim recTable As typeElpTable
Dim wAmjMin As String, wAmjMax As String
Dim fgCptBalance_FormatString As String, fgCptBalance_K As Integer

Private Sub fgCptBalance_Load()
Dim K2 As Integer, I As Integer
Dim curDB As Currency, curCR As Currency, curX As Currency

SSTab1.Tab = 0

fgCptBalance_Init
fgCptBalance.Visible = True
fgCptBalance.Clear

fgCptBalance.Rows = 1
fgCptBalance.FormatString = fgCptBalance_FormatString
fgCptBalance.Enabled = True
For I = 1 To selCompte_Nb
    fgCptBalance.Rows = fgCptBalance.Rows + 1
    fgCptBalance.Row = fgCptBalance.Rows - 1

    fgCptBalance_K = (fgCptBalance.Row) * fgCptBalance.Cols
    
    CV.DeviseN = selCompte(I).Devise: CV_AttributN CV
 
    fgCptBalance.TextArray(0 + fgCptBalance_K) = CV.DeviseIso
    fgCptBalance.TextArray(1 + fgCptBalance_K) = Compte_Imp(selCompte(I).Numéro)
    fgCptBalance.TextArray(2 + fgCptBalance_K) = Trim(selCompte(I).Intitulé) & "     " & G_CV1.DeviseIso
    curX = selCompte(I).SoldeInstantané
    
    K2 = IIf(curX < 0, 3, 4)
    fgCptBalance.TextArray(K2 + fgCptBalance_K) = Format(curX, "#### ### ###.00")
    fgCptBalance.TextArray(5 + fgCptBalance_K) = dateImp(selCompte(I).MvtAmj)
    fgCptBalance.TextArray(6 + fgCptBalance_K) = I

Next I
If fgCptBalance.Rows = 1 Then fraCptBalance.Visible = False: Exit Sub
fraCptBalance.Visible = True
fgCptBalance_Sort
ReDim arrCptBalance(selCompte_Nb + 1)

For I = 1 To fgCptBalance.Rows - 1
    fgCptBalance.Row = I
    fgCptBalance_K = (fgCptBalance.Row) * fgCptBalance.Cols
    K2 = fgCptBalance.TextArray(6 + fgCptBalance_K)
    arrCptBalance(I) = selCompte(K2)
Next I

Call selCompte_Load(Xcompte, Xcompte, "End")
'$$$$$$$$$$$$$$$$$$$$$$$

End Sub
Public Sub fgCptBalance_Sort()
fgCptBalance.Row = 1
fgCptBalance.RowSel = fgCptBalance.Rows - 1

fgCptBalance.Col = 0
fgCptBalance.ColSel = 0
fgCptBalance.Sort = 1

End Sub

Public Sub fgCptBalance_Init()
Dim mK2 As String, intReturn As Integer, selFct As String

mK2 = "Arbitrage"
selFct = "Init"

recCompteInit Xcompte
Xcompte.Method = "SnapL5"
Xcompte.Société = SocId$
Xcompte.Agence = SocAgence$

recElpTable_Init recElpTable
recElpTable.Id = "Param"
recElpTable.K1 = "BOTC"
recElpTable.Method = "Seek>="
recElpTable.K2 = mK2

'Do
'    intReturn = tableElpTable_Read(recLrRisque)
'    If intReturn = 0 Then
'        If mId$(recElpTable.K2, 1, 9) <> mK2 Then
'            intReturn = 1
'        Else
        Xcompte.Numéro = "00038210003"
        Xcompte.Devise = "000"
        Xcompte.MvtceJour = "V"
        Xcompte.chkAnnul = "0"

        xCompteMax = Xcompte
        xCompteMax.Devise = "999"
        If Not IsNull(selCompte_Load(Xcompte, xCompteMax, selFct)) Then Exit Sub

            recElpTable.Method = "Seek>="
'        End If
'     End If
'Loop While intReturn = 0

'V = dbElpTable_ReadE(recElpTable)
'If Not IsNull(V) Then GoTo Table_Error
'If IsNull(recElpTable.Memo) Then GoTo Memo_Error
'paramGuichetBillets_In = mId$(recElpTable.Memo, 1, 11)
'If Not IsNumeric(paramGuichetBillets_In) Then GoTo Num_Error


Xcompte.Method = "SnapL5"
Xcompte.Numéro = "00093100001"
Xcompte.Devise = "000"

xCompteMax = Xcompte
xCompteMax.Devise = "999"
If Not IsNull(selCompte_Load(Xcompte, xCompteMax, "Add")) Then Exit Sub
Exit Sub

Table_Error:
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : Erreur table", vbCritical, "frmBOTC_fgCptBalance_Init"
Exit Sub

Memo_Error:
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : Mémo absent", vbCritical, "frmBOTC_fgCptBalance_Init"
Exit Sub

Num_Error:
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : " & recElpTable.Memo & " :Mémo non numérique", vbCritical, "frmBOTC_fgCptBalance_Init"

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

Private Sub chkArbitrageMvtPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkArbitrageMvtPrint
End Sub

Private Sub chkarbitragepositionprint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkArbitragePositionPrint

End Sub

Private Sub chkArbitrageMvtPrint_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If chkArbitrageMvtPrint = "1" Then
    txtAmjMax.Enabled = True
    txtAmjMin.Enabled = True
Else
    txtAmjMax.Enabled = False
    txtAmjMin.Enabled = False
End If

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
Dim I As Integer

If chkArbitragePositionPrint = "1" Then prtCptBalance_Monitor "Position d'arbitrage"
If chkArbitrageMvtPrint = "1" Then
    txtAmjMin_Control
    txtAmjMax_Control
    prtCptBalance_CptMvt wAmjMin, wAmjMax
End If
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

Private Sub cmdPrintX_Click()
cmdPrint_Click
End Sub

Private Sub frmCompte_Show(XDevise As String, xNuméro As String)
X = Space$(100)
Mid$(X, 1, 12) = "frmCompte   "
Mid$(X, 13, 12) = "frmBOTC     "
Mid$(X, 25, 10) = Space$(10)
Mid$(X, 35, 3) = XDevise
Mid$(X, 38, 11) = xNuméro
Mid$(X, 49, 1) = "V"
'blnControl = False
Msg_Monitor X

End Sub



Private Sub cmdPrintX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdPrintX
End Sub


Private Sub fgCptBalance_Click()

If fgCptBalance.Row > 0 Then
    fgCptBalance_K = (fgCptBalance.Row) * fgCptBalance.Cols
    Xcompte = arrCptBalance(fgCptBalance.TextArray(6 + fgCptBalance_K))
    Call frmCompte_Show(Xcompte.Devise, Xcompte.Numéro)
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
fgCptBalance_FormatString = fgCptBalance.FormatString

Form_Init " "

End Sub
'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
usrColor_Set

cmdContext.Caption = constcmdAbandonner: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False

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

Private Sub fraCptBalance_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Public Sub Form_Init(Msg As String)


tableElpTable_Open

wCV1.Normal = "C"
wCV2.Normal = "C"

cmdReset

wAmjMin = dateElp("Ouvré", -1, DSys)
Call DTPicker_Set(txtAmjMin, wAmjMin)
wAmjMax = wAmjMin
Call DTPicker_Set(txtAmjMax, wAmjMax)
CV_Init CV

fgCptBalance_Load
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


