VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTI 
   Caption         =   "TI : interface"
   ClientHeight    =   6345
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   9330
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1200
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4920
      TabIndex        =   0
      Top             =   0
      Width           =   4305
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5700
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   10054
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Interface DB2"
      TabPicture(0)   =   "TI.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraFolder"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "TI.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.Frame fraFolder 
         Height          =   5175
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   8895
         Begin VB.CommandButton cmdFTP 
            BackColor       =   &H00C0FFC0&
            Caption         =   "FTP : TICD_GET"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   6120
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   1200
            Width           =   2160
         End
         Begin VB.TextBox txtTINTPath 
            Height          =   285
            Left            =   2160
            TabIndex        =   9
            Text            =   "\\FR11024427\AS400_IN\"
            Top             =   1440
            Width           =   3015
         End
         Begin VB.TextBox txtTIDB2File 
            Height          =   285
            Left            =   2160
            TabIndex        =   6
            Text            =   "*all"
            Top             =   840
            Width           =   3015
         End
         Begin VB.TextBox txtTIDB2Path 
            Height          =   285
            Left            =   2160
            TabIndex        =   5
            Text            =   "C:\Temp\Ti\"
            Top             =   360
            Width           =   3015
         End
         Begin VB.CommandButton cmdOk_TIDB2_Load 
            BackColor       =   &H00C0FFC0&
            Caption         =   "DB2 => NT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   6120
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   240
            Width           =   2160
         End
         Begin VB.Label lblTINTPath 
            Caption         =   "OutputFile"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label lblTIDB2Input 
            Caption         =   "Input File"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "frmTI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String
Dim blnMsgBox_Quit As Boolean
Dim X As String, X1 As String, I As Integer, Nb As Integer
Dim Msg As String, valX As String
Dim currentMethod As String, lastMethod As String

Dim IdShell

Public Sub Msg_Rcv(X As String)
'---------------------------------------------------------
End Sub



Public Sub cmdContext_Quit()
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
C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor
currentActiveControl_Name = C.Name
End Sub

'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------
lstErr.Clear
arrTag(Val(C.Tag)) = True
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor
End Sub

Private Sub cmdContext_Click()
cmdContext_Quit

End Sub

Private Sub cmdTiDB2_Load()

paramTIDB2_Input = Trim(txtTIDB2Path) & Trim(txtTIDB2File)
 Select Case UCase$(paramTIDB2_Table)
        Case "MASTER":  paramTIDB2_Output = Trim(txtTINTPath) & "CDDOSW0"
        Case "LCMASTER":  paramTIDB2_Output = Trim(txtTINTPath) & "CDDOSW0"
'        Case "CALCTE": TIDB2_CalcText
        Case "POSTING": paramTIDB2_Output = Trim(txtTINTPath) & "CDPOSW0"
        Case "PARTYDTLS": paramTIDB2_Output = Trim(txtTINTPath) & "CDPTYW0"
        Case Else: paramTIDB2_Output = Trim(txtTINTPath) & "X"
End Select

Me.MousePointer = vbHourglass
Me.Enabled = False
srvTI.TIDB2_Load
Me.MousePointer = 0
Me.Enabled = True
End Sub

Private Sub cmdFTP_Click()
srvAs400Cmd.SBMJOB "TICD_GET"

End Sub

Private Sub cmdOk_TIDB2_Load_Click()
Call lstErr_Clear(frmTI.lstErr, frmTI.cmdContext, "TIDB2 : début ...")

If UCase$(Trim(txtTIDB2File)) = "*ALL" Then
    txtTIDB2File = "LcMaster": paramTIDB2_Table = txtTIDB2File: cmdTiDB2_Load
    txtTIDB2File = "Master": paramTIDB2_Table = txtTIDB2File: cmdTiDB2_Load
    txtTIDB2File = "Posting": paramTIDB2_Table = txtTIDB2File: cmdTiDB2_Load
    txtTIDB2File = "PartyDtls": paramTIDB2_Table = txtTIDB2File: cmdTiDB2_Load
Else
    paramTIDB2_Table = Trim(txtTIDB2File): cmdTiDB2_Load
End If
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case Is = 13: KeyCode = 0:  cmdContext_Return
    Case Is = 27:  cmdContext_Quit
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select

End Sub

Private Sub Form_Load()
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
If blnJPL Then txtTINTPath = paramTemp_Folder & "\TI\": txtTIDB2Path = paramTemp_Folder & "\TI\"
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub

Private Sub fraFolder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set SSTab1
End Sub


Public Sub MouseMoveActiveControl_Reset()
For Each xobj In Me.Controls
    If MouseMoveActiveControl_Name = xobj.Name Then
        MouseMoveActiveControl_Name = ""
         If TypeOf xobj Is CommandButton Or TypeOf xobj Is ListBox Or TypeOf xobj Is MSFlexGrid Then
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
        If TypeOf C Is CommandButton Or TypeOf C Is ListBox Or TypeOf C Is MSFlexGrid Then
            MouseMoveActiveControl.BackColor = C.BackColor
            C.BackColor = MouseMoveUsr.BackColor
        Else
            MouseMoveActiveControl.ForeColor = C.ForeColor
             C.ForeColor = MouseMoveUsr.ForeColor
        End If
    End If
End If

End Sub



Private Sub txtTIDB2File_GotFocus()
txt_GotFocus txtTIDB2File
End Sub


Private Sub txtTIDB2File_LostFocus()
txt_LostFocus txtTIDB2File

End Sub


Private Sub txtTIDB2Path_GotFocus()
txt_GotFocus txtTIDB2Path

End Sub


Private Sub txtTIDB2Path_LostFocus()
txt_LostFocus txtTIDB2Path

End Sub


Private Sub txtTINTPath_GotFocus()
txt_GotFocus txtTINTPath

End Sub


Private Sub txtTINTPath_LostFocus()
txt_LostFocus txtTINTPath

End Sub


