VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "JPL"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
Dim X As String, X240 As String * 240
Dim I As Integer
I = 0
Open "s:\pelint\data\send\cdraller" For Input As #1
'Open "s:\pelint\data\send\cdraller.240" For Random As #2 Len = 240
Open "s:\pelint\data\send\cdraller.240" For Output As #2

Do Until EOF(1)
    Line Input #1, X
    I = I + 1
    X240 = X
 '   Put #2, I, X240
    Print #2, X240

Loop

Close
Unload Me
End Sub


