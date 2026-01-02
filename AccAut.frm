VERSION 5.00
Begin VB.Form frmAccAut 
   Caption         =   "Autorisations"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   8115
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5160
      TabIndex        =   3
      Top             =   0
      Width           =   2500
   End
   Begin VB.TextBox txtRecherche 
      Height          =   300
      Left            =   3240
      TabIndex        =   2
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H000000FF&
      Caption         =   "&Recherche"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   7680
      Picture         =   "AccAut.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   400
   End
End
Attribute VB_Name = "frmAccAut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim recAccAut As typeAccAut

Private Sub cmdContext_Click()
arrAccAut_Sql
End Sub


Private Sub cmdPrint_Click()
Dim X As String
X = Format$(1, "000000") & Format$(arrAccAutNb, "000000")
prtAccAutX X
End Sub


Public Sub arrAccAut_Sql()
srvAccAut.Init recAccAut
recAccAut.Method = "SnapP0"
recAccAut.AccAutId = Trim(txtRecherche)
recAccAut.AccAutK1 = ""
arrAccAut(0) = recAccAut
arrAccAut(0).AccAutId = Trim(txtRecherche) & "9z"
arrAccAutNb = 0: arrAccAutIndex = 0
arrAccAutsuite = True
arrAccAutNb = 0
Do Until Not arrAccAutsuite
    srvAccAut.Monitor recAccAut
    If arrAccAutNb > 0 Then
        recAccAut = arrAccAut(arrAccAutNb)
        recAccAut.Method = "SnapP0+"
    End If
Loop

End Sub


Public Sub Msg_Rcv(txtMsg As String)
'---------------------------------------------------------
End Sub


