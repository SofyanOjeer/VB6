VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSAB_DWH 
   Caption         =   "SAB_DWH"
   ClientHeight    =   6210
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   8775
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   8295
      Picture         =   "SAB_DWH.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   -30
      Width           =   500
   End
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
      Height          =   450
      Left            =   2880
      TabIndex        =   0
      Top             =   0
      Width           =   5265
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   60
      TabIndex        =   1
      Top             =   525
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   9975
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Exporter"
      TabPicture(0)   =   "SAB_DWH.frx":0102
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraFolder"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "..."
      TabPicture(1)   =   "SAB_DWH.frx":011E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "..."
      TabPicture(2)   =   "SAB_DWH.frx":013A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "libExe_Dir"
      Tab(2).Control(1)=   "libExe_File"
      Tab(2).Control(2)=   "libExe_Action"
      Tab(2).ControlCount=   3
      Begin VB.Frame fraFolder 
         Height          =   5280
         Left            =   60
         TabIndex        =   2
         Top             =   315
         Width           =   8520
         Begin VB.Frame fraSelect 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5115
            Left            =   60
            TabIndex        =   4
            Top             =   120
            Width           =   8415
            Begin VB.TextBox txtSelect_Source_Folder 
               Height          =   285
               Left            =   4800
               TabIndex        =   19
               Top             =   2520
               Width           =   2895
            End
            Begin VB.TextBox txtSelect_FileName 
               Height          =   285
               Left            =   4800
               TabIndex        =   15
               Top             =   3720
               Width           =   2895
            End
            Begin VB.TextBox txtSelect_Folder 
               Height          =   285
               Left            =   4800
               TabIndex        =   14
               Top             =   3240
               Width           =   2895
            End
            Begin VB.CommandButton cmdSelect_Ok 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Exporter"
               Height          =   600
               Left            =   5640
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   4200
               Width           =   1215
            End
            Begin VB.Frame fraSelect_Format 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Format"
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
               Left            =   3960
               TabIndex        =   10
               Top             =   360
               Width           =   3975
               Begin VB.OptionButton optSelect_CSV_Header 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "CSV avec en-tête"
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
                  Left            =   240
                  TabIndex        =   16
                  Top             =   1200
                  Width           =   2000
               End
               Begin VB.OptionButton optSelect_TXT 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "TXT"
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
                  Left            =   240
                  TabIndex        =   12
                  Top             =   360
                  Width           =   2000
               End
               Begin VB.OptionButton optSelect_CSV 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "CSV sans en-tête"
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
                  Left            =   240
                  TabIndex        =   11
                  Top             =   720
                  Width           =   2000
               End
            End
            Begin VB.FileListBox filDoc 
               Height          =   4380
               Left            =   240
               TabIndex        =   5
               Top             =   480
               Visible         =   0   'False
               Width           =   2535
            End
            Begin VB.Label lblSelect_Source_Folder 
               Caption         =   "Source : Répertoire"
               Height          =   375
               Left            =   3360
               TabIndex        =   20
               Top             =   2520
               Width           =   1095
            End
            Begin VB.Label lblSelect_Filename 
               Caption         =   "Fichier"
               Height          =   255
               Left            =   3360
               TabIndex        =   18
               Top             =   3720
               Width           =   975
            End
            Begin VB.Label lblSelect_Folder 
               Caption         =   "Destination : Répertoire"
               Height          =   375
               Left            =   3360
               TabIndex        =   17
               Top             =   3240
               Width           =   1335
            End
         End
      End
      Begin VB.Label libExe_Dir 
         Caption         =   "-"
         Height          =   300
         Left            =   -71310
         TabIndex        =   9
         Top             =   5200
         Width           =   4770
      End
      Begin VB.Label libExe_File 
         Caption         =   "-"
         Height          =   300
         Left            =   -71370
         TabIndex        =   8
         Top             =   4800
         Width           =   4770
      End
      Begin VB.Label libExe_Action 
         Caption         =   "-"
         Height          =   315
         Left            =   -72255
         TabIndex        =   7
         Top             =   4800
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmSAB_DWH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mHeight_0 As Integer, mWidth_0 As Integer, mHeight_2 As Integer, mWidth_2 As Integer, mWindowState As Integer
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String
Dim blnMsgBox_Quit As Boolean
Dim X As String, X1 As String, I As Integer
Dim Msg As String, valX As String
Dim currentMethod As String, lastMethod As String
Dim SAB_DWH_Aut As typeAuthorization
Dim currentAction As String
Dim IdShell

Dim blnControl As Boolean, blnError As Boolean

Dim wFileName_Source  As String, wFileName_Destination As String
Dim wIdFile_Destination As Integer, wIdFile_Source As Integer
Dim wNb As Long

Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
Dim X As String
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate
Call BiaPgmAut_Init(mId$(Msg, 1, 12), SAB_DWH_Aut)    '
SSTab1.Tab = 0
cmdReset
End Sub


Private Sub Form_Resize()
If mWindowState <> Me.WindowState Then
    If Me.WindowState = 0 Or Me.WindowState = 2 Then
        Elp_Form_Resize Me, mWindowState, mHeight_0, mWidth_0, mHeight_2, mWidth_2
    End If
End If

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

Private Sub cmdSelect_Ok_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass


cmdSelect_Exe

Me.Enabled = True: Me.MousePointer = 0


End Sub

Private Sub filDoc_Click()
txtSelect_FileName = filDoc.FileName
optSelect_Control

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case Is = 13: KeyCode = 0:  cmdContext_Return
    Case Is = 27:  cmdContext_Quit
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select

End Sub

Private Sub Form_Load()
mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)

''fraMT950.Enabled = False

End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub

Private Sub fraFolder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub optSelect_CSV_Click()
If Me.Enabled Then optSelect_Control
End Sub

Private Sub optSelect_CSV_Header_Click()
If Me.Enabled Then optSelect_Control

End Sub


Private Sub optSelect_TXT_Click()
If Me.Enabled Then optSelect_Control

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




Public Sub cmdReset()
Dim X As String
On Error Resume Next
Me.Enabled = False: Me.MousePointer = vbHourglass

blnControl = False
blnError = False
usrColor_Set
cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
currentAction = ""

txtSelect_Source_Folder = paramEnvironement

filDoc.Path = paramYBase_Data
filDoc.Pattern = "*.XXX"
filDoc.Pattern = "*Y*" & paramYBase_Data_ExtensionP
filDoc.Visible = True
blnControl = True
txtSelect_Folder = "U:\SAB_YBASE"
If blnJPL Then txtSelect_Folder = "C:\TEMP\U\SAB_YBASE"
txtSelect_FileName = ""
optSelect_CSV_Header = True
Me.Enabled = True: Me.MousePointer = 0

End Sub
'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub optSelect_Control()
Dim X As String
Dim K As Integer

X = Trim(txtSelect_FileName)
K = InStr(1, X, ".")
If K > 1 Then X = mId$(X, 1, K - 1)
txtSelect_FileName = X
If optSelect_TXT Then
    txtSelect_FileName = X & ".txt"
Else
    txtSelect_FileName = X & ".csv"
End If

End Sub

Public Sub cmdSelect_Exe()
Dim X As String
Dim V
Dim K As Integer

On Error GoTo Error_Handle

wNb = 0
X = Trim(txtSelect_Source_Folder)
If X = paramEnvironement Then
    wFileName_Source = paramYBase_DataF & filDoc.FileName
Else
    If mId$(X, Len(X), 1) <> "\" Then X = X & "\"
    wFileName_Source = X & filDoc.FileName
End If

If mId$(txtSelect_Folder, Len(txtSelect_Folder), 1) <> "\" Then txtSelect_Folder = txtSelect_Folder & "\"

wFileName_Destination = txtSelect_Folder & txtSelect_FileName
Call lstErr_Clear(lstErr, cmdContext, "Source : " & filDoc.FileName)
Call lstErr_AddItem(lstErr, cmdContext, "Destination : " & wFileName_Destination)

V = File_Export_Monitor("Input", wIdFile_Source, wFileName_Source)

If IsNull(V) Then
    V = File_Export_Monitor("Output", wIdFile_Destination, wFileName_Destination)

    If IsNull(V) Then
        If optSelect_TXT Then
            Close wIdFile_Source
            Close wIdFile_Destination
            msFileSystem.CopyFile wFileName_Source, wFileName_Destination
            Call lstErr_AddItem(lstErr, cmdContext, "Copie terminée")

        Else
            X = UCase$(Trim(filDoc.FileName))
            Select Case X
                Case "YBIAMVT0.TXT": Call srvYBIAMVT0_Export_CSV(wIdFile_Source, wIdFile_Destination, optSelect_CSV_Header, wNb)
                Case "YBIACPT0.TXT": Call srvYBIACPT0_Export_CSV(wIdFile_Source, wIdFile_Destination, optSelect_CSV_Header, wNb)
                Case "YCDODOS0.TXT": Call srvYCDODOS0_Export_CSV(wIdFile_Source, wIdFile_Destination, optSelect_CSV_Header, wNb)
                Case "YCDOUTI0.TXT": Call srvYCDOUTI0_Export_CSV(wIdFile_Source, wIdFile_Destination, optSelect_CSV_Header, wNb)
                Case "YCLIENA0.TXT": Call srvYCLIENA0_Export_CSV(wIdFile_Source, wIdFile_Destination, optSelect_CSV_Header, wNb)
                Case "YCOMPTE0.TXT": Call srvYCOMPTE0_Export_CSV(wIdFile_Source, wIdFile_Destination, optSelect_CSV_Header, wNb)
                Case "YCGSCOM0.TXT": Call srvYCGSCOM0_Export_CSV(wIdFile_Source, wIdFile_Destination, optSelect_CSV_Header, wNb)
                Case "YCGSENC0.TXT": Call srvYCGSENC0_Export_CSV(wIdFile_Source, wIdFile_Destination, optSelect_CSV_Header, wNb)
                Case "YCGSMM10.TXT": Call srvYCGSMM10_Export_CSV(wIdFile_Source, wIdFile_Destination, optSelect_CSV_Header, wNb)
                Case "YCGSMM30.TXT": Call srvYCGSMM30_Export_CSV(wIdFile_Source, wIdFile_Destination, optSelect_CSV_Header, wNb)
                Case "YCGSMM40.TXT": Call srvYCGSMM40_Export_CSV(wIdFile_Source, wIdFile_Destination, optSelect_CSV_Header, wNb)
                Case "YCGSMOY0.TXT": Call srvYCGSMOY0_Export_CSV(wIdFile_Source, wIdFile_Destination, optSelect_CSV_Header, wNb)
                Case "YPLAN0.TXT": Call srvYPLAN0_Export_CSV(wIdFile_Source, wIdFile_Destination, optSelect_CSV_Header, wNb)
                Case "YSOLDE0.TXT": Call srvYSOLDE0_Export_CSV(wIdFile_Source, wIdFile_Destination, optSelect_CSV_Header, wNb)
                Case "YTITULA0.TXT": Call srvYTITULA0_Export_CSV(wIdFile_Source, wIdFile_Destination, optSelect_CSV_Header, wNb)
                Case Else:
                    K = InStr(1, X, "YBIAMVT0.TXT")
                    If K > 0 Then
                        Call srvYBIAMVT0_Export_CSV(wIdFile_Source, wIdFile_Destination, optSelect_CSV_Header, wNb)
                    Else
                        MsgBox "NON PROGRAMME"
                    End If
            End Select
            Call lstErr_AddItem(lstErr, cmdContext, "Enregistrements copiés : " & wNb)

            Close wIdFile_Source
            Close wIdFile_Destination
        End If
    End If
End If

Exit Sub

Error_Handle:
Call lstErr_AddItem(lstErr, cmdContext, "Erreur : " & Error)

MsgBox wFileName_Source & " / " & wFileName_Destination & " : " & Error, vbCritical, Me.Caption
Close wIdFile_Source
Close wIdFile_Destination

End Sub

Private Sub txtSelect_Source_Folder_LostFocus()
On Error Resume Next
filDoc.Path = txtSelect_Source_Folder

End Sub


