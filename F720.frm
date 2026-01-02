VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "JPL"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFileBat 
      Height          =   405
      Left            =   1800
      TabIndex        =   6
      Text            =   "D:\Ficoba\Ficoba.bat"
      Top             =   1800
      Width           =   3255
   End
   Begin VB.TextBox txtFileOutput 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Text            =   "D:\Ficoba\Ficoba.txt"
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox txtFileInput 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Text            =   "S:\FTP\Ficoba.txt"
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   1215
      Left            =   3000
      TabIndex        =   0
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label lblFileBat 
      Caption         =   "Traitement"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblNb 
      Caption         =   "-"
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lblFileOutput 
      Caption         =   "Ficoba fixe"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblFileInput 
      Caption         =   "Ficoba Variable"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
Dim X As String, X720 As String * 720
Dim I As Integer
I = 0
Open Trim(txtFileInput) For Input As #1
Open Trim(txtFileOutput) For Output As #2

Do Until EOF(1)
    Line Input #1, X
    I = I + 1
    X720 = Space$(720)
    X720 = X
    Select Case Mid$(X720, 88, 2)
        Case "01": Print #2, Mid$(X720, 1, 104)
        Case "02": Print #2, Mid$(X720, 1, 91)
        Case "03": Print #2, Mid$(X720, 1, 215)
        Case "04": Print #2, Mid$(X720, 1, 720)
        Case "08": Print #2, Mid$(X720, 1, 189)
        Case "09": Print #2, Mid$(X720, 1, 99)
    End Select
Loop

Close
lblNb = "Nb : " & I

If Trim(txtFileBat) <> "" Then Call Shell(Trim(txtFileBat), 1)
Unload Me
End Sub


