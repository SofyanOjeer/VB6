VERSION 5.00
Begin VB.Form ElpShell 
   Caption         =   "ElpShell : lancement de procèdures"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraShell 
      Height          =   3375
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   6015
      Begin VB.CommandButton cmdOk 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Ok"
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
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2280
         Width           =   1200
      End
      Begin VB.TextBox txtShellCmd 
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Text            =   "c:\Temp\ElpShell.bat"
         Top             =   1080
         Width           =   3135
      End
      Begin VB.TextBox txtShellHMS 
         Height          =   375
         Left            =   2280
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label libHMS 
         Caption         =   "HH : MM : SS"
         Height          =   375
         Left            =   4080
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblShellCmd 
         Caption         =   "fichier .bat"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lblShellHMSD 
         Caption         =   "HHMMSS lancement"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "ElpShell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim wShellHMS   As Long, blnShellHMS As Boolean, vShellId
Private Sub cmdOk_Click()
Dim wL As Long
wL = CLng(txtShellHMS)
If wL > time_Hms Then wShellHMS = wL: blnShellHMS = True: fraShell.Enabled = False
End Sub

Public Static Function time_Hms() As String
Dim X As String
X = Time
time_Hms = Mid$(X, 1, 2) & Mid$(X, 4, 2) & Mid$(X, 7, 2)

End Function

Private Sub Form_Load()
blnShellHMS = False
End Sub

Private Sub Timer1_Timer()
libHMS = Time
If blnShellHMS Then
    If time_Hms >= wShellHMS Then
        blnShellHMS = False
        vShellId = Shell(Trim(txtShellCmd), 1)
        AppActivate vShellId
        DoEvents
        SendKeys "%{F4}", True
        End
    End If
End If

End Sub

