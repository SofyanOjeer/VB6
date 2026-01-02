VERSION 5.00
Begin VB.Form dlgRep 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choix du répertoire..."
   ClientHeight    =   2700
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4905
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   3195
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   180
      Width           =   3195
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   660
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   180
      Width           =   1215
   End
End
Attribute VB_Name = "dlgRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()

    Me.Tag = ""
    Me.Hide
    
End Sub


Private Sub Drive1_Change()

    Dir1.Path = Drive1.Drive

End Sub


Private Sub OKButton_Click()
'=====================================
'   Gestion coté appelant            '
'=====================================
'    Load dlgRep
'    dlgRep.Tag = ""
'    Call dlgRep.Show(vbModal)
'    List1.Clear
'    List1.AddItem dlgRep.Tag
'=====================================

    Me.Tag = Dir1.List(Dir1.ListIndex)
    Me.Hide
    
End Sub


