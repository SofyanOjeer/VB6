VERSION 5.00
Begin VB.Form frmCommission 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Commissions"
   ClientHeight    =   3480
   ClientLeft      =   555
   ClientTop       =   1950
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3480
   ScaleWidth      =   8370
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Ok"
      Height          =   420
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   0
      Width           =   600
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4920
      TabIndex        =   25
      Top             =   0
      Width           =   2595
   End
   Begin VB.Frame fraCommission 
      Height          =   3135
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   8295
      Begin VB.ListBox lstX 
         Height          =   2790
         Left            =   4200
         Sorted          =   -1  'True
         TabIndex        =   29
         Top             =   120
         Visible         =   0   'False
         Width           =   4005
      End
      Begin VB.TextBox txtComPays 
         Height          =   285
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   0
         Top             =   1900
         Width           =   450
      End
      Begin VB.TextBox txtMontantDevise 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   6720
         TabIndex        =   19
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox txtMontantFRF 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   6720
         TabIndex        =   18
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Frame fraShare 
         Caption         =   "Coefficient"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1300
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   1455
         Begin VB.OptionButton optShare1 
            Caption         =   "100 %"
            Height          =   255
            Left            =   120
            TabIndex        =   1
            Top             =   300
            Value           =   -1  'True
            Width           =   1245
         End
         Begin VB.OptionButton optShare2 
            Caption         =   " 50 %"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   600
            Width           =   1005
         End
         Begin VB.OptionButton optShare3 
            Caption         =   "autre"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   900
            Width           =   885
         End
      End
      Begin VB.Frame fraBarème 
         Caption         =   "Barème"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   1920
         TabIndex        =   21
         Top             =   240
         Width           =   1935
         Begin VB.OptionButton optBarème1 
            Caption         =   "correspondant"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   300
            Value           =   -1  'True
            Width           =   1400
         End
         Begin VB.OptionButton optBarème2 
            Caption         =   "entreprise"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   600
            Width           =   1400
         End
         Begin VB.OptionButton optBarème3 
            Caption         =   "particulier"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   900
            Width           =   1400
         End
         Begin VB.OptionButton optBarème4 
            Caption         =   "autre"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   1200
            Width           =   1400
         End
      End
      Begin VB.Frame fraCommissons 
         Caption         =   "Commissions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   4200
         TabIndex        =   20
         Top             =   240
         Width           =   4000
         Begin VB.CheckBox chk1 
            Caption         =   "frais de transfert "
            Height          =   300
            Left            =   120
            TabIndex        =   8
            Top             =   300
            Width           =   1500
         End
         Begin VB.CheckBox chk2 
            Caption         =   "frais de télex"
            Height          =   300
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   1500
         End
         Begin VB.TextBox txtMontant2 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   2500
            TabIndex        =   14
            Top             =   550
            Width           =   1215
         End
         Begin VB.TextBox txtMontant1 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   2500
            TabIndex        =   13
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtMontant5 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   2500
            TabIndex        =   17
            Top             =   1500
            Width           =   1215
         End
         Begin VB.TextBox txtMontant4 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   2500
            TabIndex        =   16
            Top             =   1150
            Width           =   1215
         End
         Begin VB.TextBox txtMontant3 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   2505
            TabIndex        =   15
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chk3 
            Caption         =   "frais Swift"
            Height          =   300
            Left            =   120
            TabIndex        =   10
            Top             =   900
            Width           =   1500
         End
         Begin VB.CheckBox chk4 
            Caption         =   "frais de change"
            Height          =   300
            Left            =   120
            TabIndex        =   11
            Top             =   1200
            Width           =   1500
         End
         Begin VB.CheckBox chk5 
            Caption         =   "TVA"
            Height          =   300
            Left            =   120
            TabIndex        =   12
            Top             =   1500
            Width           =   1500
         End
      End
      Begin VB.Label libComPays 
         Caption         =   "-"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1800
         TabIndex        =   30
         Top             =   2000
         Width           =   1815
      End
      Begin VB.Label lblComPays 
         Caption         =   "Code Pays"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   2000
         Width           =   855
      End
      Begin VB.Label lblBrut 
         Alignment       =   1  'Right Justify
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   2700
         Width           =   6255
      End
      Begin VB.Label lblMontantFRF 
         Alignment       =   1  'Right Justify
         Caption         =   "total commission"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   23
         Top             =   2400
         Width           =   3735
      End
   End
End
Attribute VB_Name = "frmCommission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim X As String, I As Integer
Dim recOpCpt As typeOpCpt, oldOpCpt As typeOpCpt
Dim recDevise As typeDevise
Dim recBarème As typeBarème
Dim curBrutFRF As Currency, curMontant As Currency
Dim blnControl As Boolean
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor


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
            MouseMoveActiveControl.ForeColor = C.ForeColor
            C.ForeColor = MouseMoveUsr.ForeColor
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
            xobj.ForeColor = MouseMoveActiveControl.ForeColor
        End If
        Exit For
    End If
Next xobj

End Sub

'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor

End Sub


'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------

arrTag(Val(C.Tag)) = True
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor
frmCommission_Control

End Sub





Private Sub chk1_Click()
If blnControl Then recOpCpt.chkCom1 = chk1: frmCommission_Control
End Sub

Private Sub chk1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chk1

End Sub


Private Sub chk2_Click()
If blnControl Then recOpCpt.chkCom2 = chk2: frmCommission_Control
End Sub


Private Sub chk2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chk2

End Sub


Private Sub chk3_Click()
If blnControl Then recOpCpt.chkCom3 = chk3.Value: frmCommission_Control
End Sub


Private Sub chk3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chk3

End Sub


Private Sub chk4_Click()
If blnControl Then recOpCpt.chkCom4 = chk4.Value: frmCommission_Control
End Sub


Private Sub chk4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chk4

End Sub


Private Sub chk5_Click()
If blnControl Then recOpCpt.chkCom5 = chk5.Value: frmCommission_Control
End Sub


Private Sub chk5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chk5

End Sub


Private Sub cmdOk_Click()
frmCommission_Control
If lstErr.ListCount > 0 Then Exit Sub
arrOpCpt(0) = recOpCpt
Unload Me
End Sub

'---------------------------------------------------------
Private Sub cmdQuit_Click()
'---------------------------------------------------------
arrOpCpt(0) = oldOpCpt
Unload Me

End Sub


Private Sub cmdOk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdOk

End Sub

'---------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------
Select Case KeyCode
    Case Is = 27: cmdQuit_Click
'    Case Is = 34: cmdPageNext_Click
 '   Case Is = 33: cmdPagePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
    Case Is = 13: SendKeys "{TAB}"
End Select

End Sub



'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------

Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)

End Sub


'---------------------------------------------------------
Private Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------

For I = 0 To arrTagNb: arrTag(I) = B: Next I

End Sub


Public Sub frmCommission_Clear()
lblBrut = ""
txtMontantDevise = ""
txtMontantFRF = ""
txtMontant1 = ""
txtMontant2 = ""
txtMontant3 = ""
txtMontant4 = ""
txtMontant5 = ""
txtMontantFRF.Enabled = False
txtMontant1.Enabled = False
txtMontant2.Enabled = False
txtMontant3.Enabled = False
txtMontant4.Enabled = False
txtMontant5.Enabled = False

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub

Private Sub fraBarème_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraCommission_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraCommissons_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraShare_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub lblComPays_Click()
lstX_Display
End Sub

Private Sub lblComPays_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set lblComPays
End Sub


Private Sub lstX_Click()
Dim I As Integer
I = InStr(1, lstX.Text, ":= 0")
If I > 0 Then txtComPays = mId$(lstX.Text, I + 4, 3): txtComPays_Control: txtComPays.SetFocus
lstX.Visible = False
End Sub

Private Sub lstX_LostFocus()
lstX.Visible = False
End Sub


Private Sub optBarème1_Click()
If blnControl Then recOpCpt.optComBarème = 1: frmCommission_Change
End Sub


Private Sub optBarème1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optBarème1

End Sub


Private Sub optBarème2_Click()
If blnControl Then recOpCpt.optComBarème = 2: frmCommission_Change

End Sub


Private Sub optBarème2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optBarème2

End Sub


Private Sub optBarème3_Click()
If blnControl Then recOpCpt.optComBarème = 3: frmCommission_Change

End Sub


Private Sub optBarème3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optBarème3

End Sub


Private Sub optBarème4_Click()
If blnControl Then
    recOpCpt.optComBarème = 4: recOpCpt.optComShare = 3
    optShare3 = True
    frmCommission_Change
End If
End Sub


Private Sub optBarème4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optBarème4

End Sub


Private Sub optShare1_Click()
If blnControl Then recOpCpt.optComShare = 1: frmCommission_Change
End Sub


Private Sub optShare1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optShare1

End Sub


Private Sub optShare2_Click()
If blnControl Then recOpCpt.optComShare = 2: frmCommission_Change
End Sub


Private Sub optShare2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optShare2

End Sub


Private Sub optShare3_Click()
If blnControl Then
    recOpCpt.optComShare = 3
    optBarème4 = True
    frmCommission_Change
End If
End Sub



Public Sub frmCommission_Barème()
Select Case recOpCpt.optComBarème
    Case 1: recBarème.Id = "OPTRFCORR"
    Case 2: recBarème.Id = "OPTRFPM"
    Case 3: recBarème.Id = "OPTRFPP"
    Case 4: recBarème.Id = "OPTRF"
End Select
If recOpCpt.optComBarème = 4 Then
    fraShare.Enabled = False
Else
    fraShare.Enabled = True
    If recOpCpt.optComShare = 3 Then recOpCpt.optComShare = 1: optShare1 = True
End If
srvBarème_Load recBarème
curMontant = Barème_Calc(curBrutFRF, recBarème)
If optShare2 Then
    curMontant = curMontant / 2
    txtMontant1 = num_Display(curMontant, 9, 2, Lx, X, " ")
    txtMontant2 = num_Display(recBarème.Télex / 2, 9, 2, Lx, X, " ")
    txtMontant3 = num_Display(recBarème.Swift / 2, 9, 2, Lx, X, " ")
Else
    txtMontant1 = num_Display(curMontant, 9, 2, Lx, X, " ")
    txtMontant2 = num_Display(recBarème.Télex, 9, 2, Lx, X, " ")
    txtMontant3 = num_Display(recBarème.Swift, 9, 2, Lx, X, " ")
End If
txtMontant5 = ""
txtMontant4 = txtMontant1
If recBarème.Id = "OPTRFPM   " And Val(txtMontant4) < 80 Then txtMontant4 = "80.00"
End Sub

Public Sub frmCommission_Control()
Dim curComFRF As Currency, curComTVA As Currency
lstErr.Clear: lstErr.Height = 0
curComFRF = 0: curComTVA = 0
chk1.ForeColor = IIf(chk1 = "1", libUsr.ForeColor, lblUsr.ForeColor)
chk2.ForeColor = IIf(chk2 = "1", libUsr.ForeColor, lblUsr.ForeColor)
chk3.ForeColor = IIf(chk3 = "1", libUsr.ForeColor, lblUsr.ForeColor)
chk4.ForeColor = IIf(chk4 = "1", libUsr.ForeColor, lblUsr.ForeColor)
chk5.ForeColor = IIf(chk5 = "1", libUsr.ForeColor, lblUsr.ForeColor)
If recOpCpt.optComBarème = 4 Then
    txtMontant1.Enabled = IIf(chk1 = "1", True, False)
    txtMontant2.Enabled = IIf(chk2 = "1", True, False)
    txtMontant3.Enabled = IIf(chk3 = "1", True, False)
    txtMontant4.Enabled = IIf(chk4 = "1", True, False)
Else
    txtMontant1.Enabled = False
    txtMontant2.Enabled = False
    txtMontant3.Enabled = False
    txtMontant4.Enabled = False
    txtMontant5.Enabled = False
End If

If recOpCpt.chkCom1 = "0" Then
    recOpCpt.ComMontant1 = 0
Else
    Call num_Control(txtMontant1, recOpCpt.ComMontant1, 9, 2)
End If
If recOpCpt.chkCom2 = "0" Then
    recOpCpt.ComMontant2 = 0
Else
    Call num_Control(txtMontant2, recOpCpt.ComMontant2, 9, 2)
End If
If recOpCpt.chkCom3 = "0" Then
    recOpCpt.ComMontant3 = 0
Else
    Call num_Control(txtMontant3, recOpCpt.ComMontant3, 9, 2)
End If
If recOpCpt.chkCom4 = "0" Then
    recOpCpt.ComMontant4 = 0
Else
    Call num_Control(txtMontant4, recOpCpt.ComMontant4, 9, 2)
End If

curComFRF = recOpCpt.ComMontant1 + recOpCpt.ComMontant2 + recOpCpt.ComMontant3 + recOpCpt.ComMontant4
curComTVA = curComFRF * tauxTVA
txtMontant5 = num_Display(curComTVA, 9, 2, Lx, X, " ")
If recOpCpt.chkCom5 = "0" Then
    recOpCpt.ComMontant5 = 0
Else
    Call num_Control(txtMontant5, recOpCpt.ComMontant5, 9, 2)
    curComFRF = curComFRF + recOpCpt.ComMontant5
End If
recOpCpt.ComMontantFRF = curComFRF
txtMontantFRF = num_Display(curComFRF, 9, 2, Lx, X, " ")

recOpCpt.ComMontant1 = curCV(recOpCpt.ComMontant1, recOpCpt.CoursDevise, recDevise.maxD)
recOpCpt.ComMontant2 = curCV(recOpCpt.ComMontant2, recOpCpt.CoursDevise, recDevise.maxD)
recOpCpt.ComMontant3 = curCV(recOpCpt.ComMontant3, recOpCpt.CoursDevise, recDevise.maxD)
recOpCpt.ComMontant4 = curCV(recOpCpt.ComMontant4, recOpCpt.CoursDevise, recDevise.maxD)
recOpCpt.ComMontant5 = curCV(recOpCpt.ComMontant5, recOpCpt.CoursDevise, recDevise.maxD)
recOpCpt.ComMontantDevise = recOpCpt.ComMontant1 + recOpCpt.ComMontant2 + recOpCpt.ComMontant3 + recOpCpt.ComMontant4 + recOpCpt.ComMontant5
txtMontantDevise = num_Display(recOpCpt.ComMontantDevise, 9, recDevise.maxD, txtMontantDevise.ForeColor, X, "0")
txtComPays_Control
End Sub

Public Sub recOpCpt_Display()
Dim X3 As String
frmCommission_Clear
lblMontantFRF = IIf(recOpCpt.optComImputation = 1, "Commission BENEFICIAIRE", "Commission DONNEUR D'ORDRE")
X3 = DevX(recOpCpt.Devise)
recDevise = XDevise
If recOpCpt.CoursDevise = 0 Then recOpCpt.CoursDevise = 1
lblBrut.ForeColor = warnUsrColor
txtMontantDevise.ForeColor = warnUsrColor
lblBrut = X3 & num_Display(recOpCpt.Brut, 13, recDevise.maxD, Lx, X, "0") _
         & "     cours :" & num_Display(recOpCpt.CoursDevise, 3, 8, Lx, X, "0")
curBrutFRF = recOpCpt.Brut * recOpCpt.CoursDevise
Select Case recOpCpt.optComShare
    Case 1: optShare1 = True
    Case 2: optShare2 = True
    Case 3: optShare3 = True
End Select
Select Case recOpCpt.optComBarème
    Case 1: optBarème1 = True
    Case 2: optBarème2 = True
    Case 3: optBarème3 = True
    Case 4: optBarème4 = True
End Select
If recOpCpt.chkCom1 = 1 Then chk1 = 1
If recOpCpt.chkCom2 = 1 Then chk2 = 1
If recOpCpt.chkCom3 = 1 Then chk3 = 1
If recOpCpt.chkCom4 = 1 Then chk4 = 1
If recOpCpt.chkCom5 = 1 Then chk5 = 1
If (recOpCpt.ComMontantFRF) <> 0 Then
    curMontant = curCVInv(recOpCpt.ComMontant1, recOpCpt.CoursDevise, recDevise.maxD)
    txtMontant1 = num_Display(curMontant, 9, 2, Lx, X, " ")
    curMontant = curCVInv(recOpCpt.ComMontant2, recOpCpt.CoursDevise, recDevise.maxD)
    txtMontant2 = num_Display(curMontant, 9, 2, Lx, X, " ")
    curMontant = curCVInv(recOpCpt.ComMontant3, recOpCpt.CoursDevise, recDevise.maxD)
    txtMontant3 = num_Display(curMontant, 9, 2, Lx, X, " ")
    curMontant = curCVInv(recOpCpt.ComMontant4, recOpCpt.CoursDevise, recDevise.maxD)
    txtMontant4 = num_Display(curMontant, 9, 2, Lx, X, " ")
    curMontant = curCVInv(recOpCpt.ComMontant5, recOpCpt.CoursDevise, recDevise.maxD)
    txtMontant5 = num_Display(curMontant, 9, 2, Lx, X, " ")
    txtMontantFRF = num_Display(recOpCpt.ComMontantFRF, 9, 2, Lx, X, " ")
    txtMontantDevise = num_Display(recOpCpt.ComMontantDevise, 9, recDevise.maxD, Lx, X, " ")
Else
    frmCommission_Change
End If
If Trim(recOpCpt.Method) = constSaisie Then
    cmdOk.Visible = True
    fraShare.Enabled = True
    fraBarème.Enabled = True
    fraCommission.Enabled = True
Else
    cmdOk.Visible = False
    fraShare.Enabled = False
    fraBarème.Enabled = False
    fraCommission.Enabled = False
End If
txtComPays = IIf(recOpCpt.ComPays = "000", "", recOpCpt.ComPays)
End Sub

Public Sub frmCommission_Change()
If Not optShare3 Then frmCommission_Barème
frmCommission_Control
End Sub

'---------------------------------------------------------
Public Sub lstX_Display()
'---------------------------------------------------------

lstX.Visible = True
Call LstDictio(19, lstX, "DicLib")
lstX.SetFocus

End Sub

Private Sub optShare3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optShare3

End Sub

Private Sub txtComPays_GotFocus()
txt_GotFocus txtComPays
End Sub


Private Sub txtComPays_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii
End Sub

Private Sub txtComPays_LostFocus()
txt_LostFocus txtComPays
End Sub

Private Sub txtmontant1_GotFocus()
txt_GotFocus txtMontant1
End Sub


Private Sub txtmontant1_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, txtMontant1)
End Sub

Private Sub txtmontant1_LostFocus()
txt_LostFocus txtMontant1
End Sub



Private Sub txtMontant2_GotFocus()
txt_GotFocus txtMontant2
End Sub


Private Sub txtMontant2_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, txtMontant2)
End Sub


Private Sub txtMontant2_LostFocus()
txt_LostFocus txtMontant2
End Sub


Private Sub txtMontant3_GotFocus()
txt_GotFocus txtMontant3
End Sub


Private Sub txtMontant3_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, txtMontant3)
End Sub


Private Sub txtMontant3_LostFocus()
txt_LostFocus txtMontant3
End Sub


Private Sub txtMontant4_GotFocus()
txt_GotFocus txtMontant4
End Sub


Private Sub txtMontant4_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, txtMontant4)
End Sub


Private Sub txtMontant4_LostFocus()
txt_LostFocus txtMontant4
End Sub


Private Sub txtMontant5_GotFocus()
txt_GotFocus txtMontant5
End Sub


Private Sub txtMontant5_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, txtMontant5)
End Sub


Private Sub txtMontant5_LostFocus()
txt_LostFocus txtMontant5
End Sub




Public Sub txtComPays_Control()
Dim X As String
If Val(txtComPays) = 0 Then Call lstErr_AddItem(lstErr, txtComPays, "?Pays de la commission "): Exit Sub
X = DicLib("019", "0" & Trim(txtComPays))
If X = "-" Then Call lstErr_AddItem(lstErr, txtComPays, "?Pays de la commission "): Exit Sub
libComPays = X
recOpCpt.ComPays = Format$(Val(txtComPays), "000")

End Sub

Public Sub frmCommission_Init()
recOpCpt = arrOpCpt(0)
oldOpCpt = recOpCpt
blnControl = False
recOpCpt_Display
blnControl = True
frmCommission_Change

End Sub
