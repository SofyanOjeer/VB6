VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDWH_Statut 
   AutoRedraw      =   -1  'True
   Caption         =   "DWH_Statut : accessibilté des données"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13875
   Icon            =   "DWH_Statut.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   13875
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   8280
      TabIndex        =   4
      Top             =   45
      Width           =   5055
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8895
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   15690
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Màj du statut des fichiers BODWH"
      TabPicture(0)   =   "DWH_Statut.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraSource"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "?"
      TabPicture(1)   =   "DWH_Statut.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.Frame fraSource 
         Height          =   8445
         Left            =   0
         TabIndex        =   5
         Top             =   360
         Width           =   13560
         Begin VB.Frame fraSource_Update 
            Height          =   5535
            Left            =   8160
            TabIndex        =   7
            Top             =   1560
            Width           =   4335
            Begin VB.OptionButton optSource_W 
               Alignment       =   1  'Right Justify
               BackColor       =   &H000000FF&
               Caption         =   "Invalider        ('V' => 'W')"
               Height          =   375
               Left            =   360
               TabIndex        =   14
               Top             =   2640
               Width           =   2535
            End
            Begin VB.CommandButton cmdSource_Ok 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Lancer le traitement interactif"
               Height          =   885
               Left            =   1080
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   4320
               Width           =   2235
            End
            Begin VB.OptionButton optSource_V 
               Alignment       =   1  'Right Justify
               BackColor       =   &H0080FF80&
               Caption         =   "Valider         ('W' => 'V')"
               Height          =   375
               Left            =   360
               TabIndex        =   10
               Top             =   1920
               Value           =   -1  'True
               Width           =   2535
            End
            Begin VB.TextBox txtSource_VER 
               Height          =   285
               Left            =   2640
               TabIndex        =   9
               Text            =   "1"
               Top             =   600
               Width           =   495
            End
            Begin MSComCtl2.DTPicker txtSelect_DRCHPER 
               Height          =   300
               Left            =   2640
               TabIndex        =   13
               Top             =   1320
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CustomFormat    =   "dd  MM yyy"
               Format          =   44826627
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label lblSource_PER 
               Caption         =   "Période"
               Height          =   255
               Left            =   360
               TabIndex        =   12
               Top             =   1320
               Width           =   1455
            End
            Begin VB.Label lblSource_VER 
               Caption         =   "Version"
               Height          =   255
               Left            =   360
               TabIndex        =   8
               Top             =   600
               Width           =   1455
            End
         End
         Begin VB.ListBox lstSource 
            Height          =   8160
            ItemData        =   "DWH_Statut.frx":0044
            Left            =   120
            List            =   "DWH_Statut.frx":004B
            Style           =   1  'Checkbox
            TabIndex        =   6
            Top             =   240
            Width           =   6240
         End
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   13320
      Picture         =   "DWH_Statut.frx":005A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   500
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1200
   End
   Begin VB.Label libRéférenceInterne 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   0
      Width           =   5535
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuContextOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuContext_x1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuContextQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu mnuselect 
      Caption         =   "mnuSelect"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect_Quit 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuselecté 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "frmDWH_Statut"
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
Dim SAB_MNUAut As typeAuthorization
Dim blnTransaction As Boolean


Dim blncmdOk_Visible As Boolean, blnErr As Boolean, blncmdSave_Visible As Boolean
Dim blnSetfocus As Boolean

Dim lstSource_Lib As String, lstSource_CGR As String
Dim lstSource_TopIndex As Long


Dim blnAuto As Boolean, blnAuto_Ok As Boolean
Dim wAMJ As String, wSTA_Old As String, wSTA_New As String

'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub cmdContext_Quit()
blnControl = False
If currentAction <> "" Then
    X = MsgBox("Voulez-vous réellement abandonner la mise à jour?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
    If X = vbYes Then
        currentAction = ""
    Else
        Exit Sub
    End If
End If


lstErr.Clear
If SSTab1.Tab > 0 Then
    SSTab1.Tab = SSTab1.Tab - 1
Else
End If
End Sub




Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdContext

End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdPrint

End Sub

Private Sub cmdSource_Ok_Click()
Me.Enabled = False
Me.MousePointer = vbHourglass
If optSource_V Then
    Call lstErr_Clear(lstErr, cmdContext, "> MàJ VALIDATION")
    wSTA_Old = "W": wSTA_New = "V"
    Call cmdSource_Update
Else
    If optSource_W Then
        Call lstErr_Clear(lstErr, cmdContext, "> MàJ invalidation")
        wSTA_Old = "V": wSTA_New = "W"
        Call cmdSource_Update
    Else
         Call lstErr_Clear(lstErr, cmdContext, "? Préciser le type de MàJ")
   End If

End If

Call lstErr_AddItem(lstErr, cmdContext, "- Fin du traitement")


Me.MousePointer = 0
Me.Enabled = True

End Sub

Private Sub lstSource_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
lstSource_TopIndex = lstSource.TopIndex
'lstSource_Select lstSource.Selected(lstSource.ListIndex)
'txtlstSourceScan.SetFocus
End Sub



'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
Dim I As Integer

blnControl = False
usrColor_Set
lstSource.BackColor = &HC0C0C0
lstUsr.BackColor = &HE0E0E0

cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
blncmdOk_Visible = False: blncmdSave_Visible = False
currentAction = ""

blnAuto = False
blnAuto_Ok = False

libRéférenceInterne = ""

lstSource.Clear
'----------------------------------------------------------------------
Call lst_LoadK2("DWH", "DWH_Statut", lstSource, True)

For I = 0 To lstSource.ListCount - 1
    lstSource.Selected(I) = True
Next I
blnControl = True
End Sub
Public Sub Form_Init()
Call lstErr_Clear(lstErr, cmdContext, "Initialisation ")
DoEvents

SSTab1.Tab = 0

blnControl = False

cmdReset


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

Private Sub cmdPrint_Click()
Dim Msg As String
Dim I As Integer

Me.Enabled = False: Me.MousePointer = vbHourglass


Me.Enabled = True: Me.MousePointer = 0

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
mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
blnControl = False
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

Private Sub mnuContextAbandonner_Click()
cmdContext_Quit
End Sub

Private Sub mnuContextQuitter_Click()
Unload Me
End Sub

Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init(Mid$(Msg, 1, 12), SAB_MNUAut)

blnSetfocus = True
Form_Init


End Sub


Public Sub cmdContext_Return()
If SSTab1.Tab = 0 Then
'    cmdlstSourceScan_Click
Else
    If currentAction = "" Then
        If SSTab1.Tab > 0 Then
            SSTab1.Tab = 0
        Else
           'SendKeys "{TAB}"
           ' cmdSelect_Click
        End If
    End If
End If
End Sub









Public Sub cmdSource_Update()
Dim K As Long, K1 As Integer
Dim wFileName  As String, wName  As String, wMemo As String
Dim wFileName_Type  As String, wFileName_Champ  As String
Dim V
On Error GoTo Error_Handler

'-------------------------------------------------------
App_Debug = "cmdSource_Update"
'-------------------------------------------------------
Call DTPicker_Control(txtSelect_DRCHPER, wAMJ)
For K = 0 To lstSource.ListCount - 1
    If lstSource.Selected(K) Then
        lstSource.ListIndex = K
        K1 = 0
        wFileName = Space_Scan(lstSource.Text, K1)
        Call rsElpTable_Read("DWH", "DWH_Statut", wFileName, wName, wMemo)
        If wMemo <> "" Then
            K1 = 0
            wFileName_Type = Space_Scan(wMemo, K1)
            wFileName_Champ = Space_Scan(wMemo, K1)
            If wFileName_Champ <> "" Then
                Call cmdSource_Update_Table(wFileName, wFileName_Type, wFileName_Champ)
            End If
        End If
    End If
Next K
GoTo Exit_Sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_Sub:

End Sub
Public Sub cmdSource_Update_Table(lFileName As String, lFileName_Type As String, lFileName_Champ As String)
Dim V, xSql As String
Dim Nb As Long
On Error GoTo Error_Handler

'-------------------------------------------------------
App_Debug = "cmdSource_Update_Table " & lFileName
'-------------------------------------------------------
Call lstErr_AddItem(lstErr, cmdContext, lFileName)


xSql = "update " & paramIBM_Library_BODWH & "." & lFileName _
     & " set " & lFileName_Champ & "STA = '" & wSTA_New & " ' where " & lFileName_Champ & "STA = '" & wSTA_Old & "'" _
     & " and " & lFileName_Champ & "VER = '1' and " & lFileName_Champ & "PER = '" & wAMJ & "'"

Set rsSab = cnsab.Execute(xSql, Nb)


If Nb = 0 Then
    Call lstErr_ChangeLastItem(lstErr, cmdContext, lFileName & " : Pas de Màj")
Else
    Call lstErr_ChangeLastItem(lstErr, cmdContext, lFileName & " : " & Nb)
End If


GoTo Exit_Sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_Sub:

End Sub

