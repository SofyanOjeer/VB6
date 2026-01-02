VERSION 5.00
Begin VB.Form frmPieces 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CPT_SCHEMA - Pièces jointes"
   ClientHeight    =   4005
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   5220
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3180
      Left            =   120
      TabIndex        =   2
      Top             =   540
      Width           =   3135
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   180
      Width           =   4935
   End
End
Attribute VB_Name = "frmPieces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub affiche_piece()
Dim nomRep As String

    nomRep = frmCPT_SCHEMA.labNomRep.Caption
    ShellExecute Me.hwnd, "open", paramCPT_SCHEMA_Dossier_Path & nomRep & "\" & List1.List(List1.ListIndex), "", App.Path, 1

End Sub

Private Sub CancelButton_Click()

    Unload Me
    
End Sub

Private Sub List1_DblClick()

    Call affiche_piece
    
End Sub


Private Sub OKButton_Click()

    If List1.ListIndex > -1 Then
        Call affiche_piece
    End If
    
End Sub


