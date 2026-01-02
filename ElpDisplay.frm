VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmElpDisplay 
   Caption         =   "frmElpDisplay"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   9135
   Begin MSFlexGridLib.MSFlexGrid fgData 
      Height          =   5490
      Left            =   15
      TabIndex        =   0
      Top             =   -15
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   9684
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      RowHeightMin    =   250
      BackColor       =   15794175
      ForeColor       =   8388608
      BackColorFixed  =   8454143
      ForeColorFixed  =   -2147483641
      BackColorSel    =   12648384
      BackColorBkg    =   15794175
      AllowBigSelection=   0   'False
      TextStyleFixed  =   4
      FocusRect       =   2
      HighLight       =   0
      GridLines       =   2
      AllowUserResizing=   3
      FormatString    =   $"ElpDisplay.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmElpDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mHeight_0 As Integer, mWidth_0 As Integer, mHeight_2 As Integer, mWidth_2 As Integer, mWindowState As Integer
Dim arrTag() As Boolean, arrTagNb As Integer

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

Public Sub cmdContext_Return()
SendKeys "{TAB}"

End Sub
Public Sub cmdContext_Quit()
Unload Me
End Sub


Private Sub Form_Load()
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
fgData.Clear: fgData.Row = 0

mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

End Sub


Private Sub Form_Resize()
If mWindowState <> Me.WindowState Then
    If Me.WindowState = 0 Or Me.WindowState = 2 Then
        Elp_Form_Resize Me, mWindowState, mHeight_0, mWidth_0, mHeight_2, mWidth_2
    End If
End If
End Sub


