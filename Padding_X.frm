VERSION 5.00
Begin VB.Form Padding_X 
   Caption         =   "Padding_X"
   ClientHeight    =   3084
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   3084
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCopy 
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      Text            =   "1000"
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Top             =   360
      Width           =   2295
   End
   Begin VB.TextBox txtIter 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Text            =   "1000000"
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Générer un fichier, puis le copier"
      Height          =   2055
      Left            =   5280
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblIter 
      Caption         =   "X"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3720
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblCopy 
      Caption         =   "X"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Nb Copies"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label label2 
      Caption         =   "Fichier"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "NB itération"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "Padding_X"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mFilename As String
Public msFileSystem, msFile

Private Sub Command1_Click()
Dim I As Long, K As Long, X As String
Dim Nb1 As Long

Randomize
mFilename = Trim(txtFileName)
Nb1 = Val(txtIter)
If Nb1 > 0 Then
    Open mFilename For Output As #1
    For I = 1 To Nb1
        If I Mod 100 = 0 Then lblIter = I: DoEvents
        X = ""
        For K = 1 To Int((Rnd * 250 + 1))
            X = X & Chr$(Int((Rnd * 250 + 1)))
        Next K
        Print #1, X
    Next I
    Close 1
End If

Nb1 = Val(txtCopy)
If Nb1 > 0 Then

    For I = 1 To Nb1
        'If I Mod 100 = 0 Then
        lblCopy = I: DoEvents
        
        X = "C:\Temp\" & Int((Rnd * 32000 + 1)) & ".txt"
        If Trim(Dir(X)) = "" Then
            msFileSystem.CopyFile mFilename, X
        End If
    Next I
End If

End Sub

Private Sub Form_Load()
txtFileName = "C:\Temp\" & Int((Rnd * 32000 + 1)) & ".txt"
Set msFileSystem = CreateObject("Scripting.FileSystemObject")

End Sub


