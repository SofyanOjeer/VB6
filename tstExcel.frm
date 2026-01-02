VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   645
      Left            =   6840
      TabIndex        =   2
      Top             =   600
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   8415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lblEXcel 
      Height          =   615
      Left            =   2640
      TabIndex        =   3
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
List1.Clear

srvTIExcel.appExcel_Monitor lblEXcel
Unload Me

End Sub



