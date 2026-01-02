VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmPaieSage 
   Caption         =   "Paie : transfert des fichiers SAGE vers AS400"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13875
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   13875
   Begin VB.Frame fraSave 
      Caption         =   "Sauvegarde"
      Height          =   4455
      Left            =   120
      TabIndex        =   9
      Top             =   4920
      Width           =   13695
      Begin VB.CommandButton cmdArchiveDelete 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Supprimer"
         Height          =   1605
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2640
         Width           =   4260
      End
      Begin VB.DirListBox dirSave 
         Height          =   4140
         Left            =   6600
         TabIndex        =   11
         Top             =   240
         Width           =   6735
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H80000000&
         Caption         =   "Sauvegarde"
         Height          =   1545
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   480
         Width           =   4305
      End
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   7920
      TabIndex        =   8
      Top             =   0
      Width           =   5865
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6960
      Top             =   0
   End
   Begin VB.Frame fraImp 
      Caption         =   "Impression (des fichiers AS400)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   7440
      TabIndex        =   4
      Top             =   960
      Width           =   6060
      Begin VB.CommandButton cmdPrintVirement 
         Caption         =   "Impression des virements"
         Height          =   1185
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   480
         Width           =   3120
      End
      Begin VB.CommandButton cmdPrintComptabilité 
         Caption         =   "Impression des mouvements comptables"
         Height          =   1185
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2160
         Width           =   3120
      End
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Abandonner"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1200
   End
   Begin VB.Frame fraTrf 
      Caption         =   "Transfert des fichiers SAGE vers AS400"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   6060
      Begin VB.CommandButton cmdExportComptabilité 
         Caption         =   "Export des mouvements comptables"
         Height          =   1305
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1920
         Width           =   3600
      End
      Begin VB.CommandButton cmdExportVirement 
         Caption         =   "Export des virements"
         Height          =   1185
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   3480
      End
   End
   Begin ComctlLib.ProgressBar prgBar 
      Height          =   420
      Left            =   1200
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   741
      _Version        =   327682
      Appearance      =   1
      Max             =   15000
   End
End
Attribute VB_Name = "frmPaieSage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String
Dim blnMsgBox_Quit As Boolean
Dim X As String, X1 As String, I As Integer
Dim Msg As String, valX As String
Dim currentMethod As String, lastMethod As String

Dim autPaieSage As typeAuthorization

Dim IdShell

Dim paramPaieSage_Virement As String
Dim paramPaieSage_Comptabilité As String
Dim paramPaieSage_Path As String
Dim paramPaieSage_Dossier As String
Dim paramPaieSage_Archive As String
Dim paramPaieSage_SAB_CPT As String
Dim paramPaieSage_SAB_VIR As String

Dim mArchiveDelete_Select As String
Public Sub Msg_Rcv(X As String)
'---------------------------------------------------------

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

Public Sub cmdContext_Quit()
X = vbYes
If blnMsgBox_Quit Then
   X = MsgBox("Voulez-vous réellement abandonner?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
End If
If X = vbYes Then Unload Me

End Sub


Public Sub cmdContext_Return()

SendKeys "{TAB}"

End Sub


Private Sub cmdArchiveDelete_Click()
On Error Resume Next
dirSave.PATH = mArchiveDelete_Select
msFileSystem.DeleteFolder (mArchiveDelete_Select), True
dirSave.PATH = paramPaieSage_Archive
'dirSave.Pattern = "*.*"
cmdArchiveDelete_False

End Sub

Private Sub cmdSave_Click()
Dim xDest As String, xSrc As String

xSrc = paramPaieSage_Path & paramPaieSage_Dossier
xDest = paramPaieSage_Archive & paramPaieSage_Dossier & "_" & DSys & "_" & time_Hms
Call lstErr_Clear(lstErr, cmdSave, "Sauvegarde " & xSrc)
dirSave.PATH = xSrc

Set msFile = msFileSystem.GetFolder(xSrc)
Screen.MousePointer = vbHourglass
prgBar.Max = 3000
prgBar.Value = 100
Timer1.Enabled = True
prgBar.Visible = True
DoEvents
msFileSystem.CopyFolder xSrc, xDest
Screen.MousePointer = vbDefault
Timer1.Enabled = False
prgBar.Visible = False

Call lstErr_AddItem(lstErr, cmdSave, "Sauvegarde terminée" & xDest)
'dirSave.Visible = True
dirSave.PATH = paramPaieSage_Archive

End Sub

Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Set cmdSave
End Sub


Private Sub cmdContext_Click()
Select Case cmdContext.Caption
'    Case Is = constcmdRechercher
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdExportComptabilité_Click()
Dim xFileName As String
'''    wFile_Import = paramSAA_DataF_from_SAB & paramSAA_Data_from_SAB_YFile & "_" & Format$(wseq, "00000") & paramSAA_Data_from_SAB_ExtensionP_sab
'''    Call Shell_FTP(wFile_Import, paramIBM_Library_SABSPE, paramSAA_Data_from_SAB_YFile, True, False)

' export fichier Comptable vers SAB
Call Shell_FTP(paramPaieSage_Comptabilité, paramIBM_Library_SABSPE, paramPaieSage_SAB_CPT, False, False)

' import fichier Comptable de SAB & impression de contrôle & suppression du fichier
xFileName = paramPaieSage_Comptabilité & "_" & DSys & "_" & time_Hms

Call Shell_FTP(xFileName, paramIBM_Library_SABSPE, paramPaieSage_SAB_CPT, True, False)
cmdPrint_Kill xFileName   'constFTP_Dir & "Salaires_Put"

Call lstErr_AddItem(lstErr, cmdContext, "Suppression :" & xFileName)
X = Dir(xFileName)
If X <> "" Then Kill xFileName

End Sub

Private Sub cmdPrintComptabilité_Click()
Dim Nb As Long, Iter As Integer
Dim xFileName As String

On Error GoTo cmdTransfert_Error

' import fichier Comptable de SAB & impression de contrôle & suppression du fichier
xFileName = paramPaieSage_Comptabilité & "_" & DSys & "_" & time_Hms

Call Shell_FTP(xFileName, paramIBM_Library_SABSPE, paramPaieSage_SAB_CPT, True, False)

cmdPrint_Kill xFileName


Exit Sub

'---------------------------------------------------------
cmdTransfert_Error:
'---------------------------------------------------------

MsgBox "erreur : " & Err & " " & Error$(Err), vbCritical, "PaieSage : cmdPrintComptabilité : " & xFileName
Resume cmdTransfert_End

cmdTransfert_End:


End Sub

Private Sub cmdPrintVirement_Click()
Dim xFileName As String

On Error GoTo cmdTransfert_Error

xFileName = paramPaieSage_Virement & "_" & DSys & "_" & time_Hms

Call Shell_FTP(xFileName, paramIBM_Library_SABSPE, paramPaieSage_SAB_VIR, True, False)


cmdPrint_Kill xFileName


Exit Sub

'---------------------------------------------------------
cmdTransfert_Error:
'---------------------------------------------------------

MsgBox "erreur : " & Err & " " & Error$(Err), vbCritical, "PaiSage : cmdPrintVirement : " & xFileName
Resume cmdTransfert_End

cmdTransfert_End:


End Sub

Private Sub dirSave_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim iLen As Integer, xFile As String

mArchiveDelete_Select = dirSave.List(dirSave.ListIndex)

xFile = paramPaieSage_Archive & paramPaieSage_Dossier & "_"
iLen = Len(xFile)
If Mid$(mArchiveDelete_Select, 1, iLen) = xFile Then
    cmdArchiveDelete.Enabled = True
    cmdArchiveDelete.Caption = "Supprimer : " & Mid$(mArchiveDelete_Select, iLen, Len(mArchiveDelete_Select) - iLen + 1)
Else
    cmdArchiveDelete_False
End If

End Sub

'---------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------
Select Case KeyCode
    Case Is = 13: KeyCode = 0: cmdContext_Return
    Case Is = 27: cmdContext_Quit
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select

End Sub

'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)

Call BiaPgmAut_Init("PaieSage", autPaieSage)
Form_Clear
param_Init
dirSave.PATH = paramPaieSage_Archive

End Sub

'---------------------------------------------------------
Public Sub Form_Clear()
'---------------------------------------------------------
lstErrClear = True
blnMsgBox_Quit = False
usrColor_Set

cmdContext.Enabled = True: cmdContext.BackColor = vbWindowBackground
cmdContext.Caption = constcmdAbandonner: cmdContext.BackColor = errUsr.BackColor
cmdArchiveDelete_False

cmdSave.Enabled = autPaieSage.Valider
cmdExportComptabilité.Enabled = autPaieSage.Valider
cmdExportVirement.Enabled = autPaieSage.Valider
cmdPrintComptabilité.Enabled = autPaieSage.Valider
cmdPrintVirement.Enabled = autPaieSage.Valider
End Sub



Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Set cmdContext

End Sub


Private Sub cmdExportComptabilité_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Set cmdExportComptabilité

End Sub


Private Sub cmdExportVirement_Click()
Dim xFileName As String
'''    wFile_Import = paramSAA_DataF_from_SAB & paramSAA_Data_from_SAB_YFile & "_" & Format$(wseq, "00000") & paramSAA_Data_from_SAB_ExtensionP_sab
'''    Call Shell_FTP(wFile_Import, paramIBM_Library_SABSPE, paramSAA_Data_from_SAB_YFile, True, False)

' export fichier Comptable vers SAB
Call Shell_FTP(paramPaieSage_Virement, paramIBM_Library_SABSPE, paramPaieSage_SAB_VIR, False, False)

' import fichier Comptable de SAB & impression de contrôle & suppression du fichier
xFileName = paramPaieSage_Virement & "_" & DSys & "_" & time_Hms

Call Shell_FTP(xFileName, paramIBM_Library_SABSPE, paramPaieSage_SAB_VIR, True, False)
cmdPrint_Kill xFileName   'constFTP_Dir & "Salaires_Put"

Call lstErr_AddItem(lstErr, cmdContext, "Suppression :" & xFileName)
X = Dir(xFileName)
If X <> "" Then Kill xFileName
End Sub

Public Function param_Init()
Dim V
Dim xName As String, xMemo As String
Dim X As String

param_Init = Null
V = rsElpTable_Read("DRH", "PaieSage", "Virement", xName, paramPaieSage_Virement)
Call lstErr_Clear(lstErr, cmdContext, "Virement :" & paramPaieSage_Virement)

V = rsElpTable_Read("DRH", "PaieSage", "Comptabilité", xName, paramPaieSage_Comptabilité)
Call lstErr_AddItem(lstErr, cmdContext, "Comptabilité :" & paramPaieSage_Comptabilité)

V = rsElpTable_Read("DRH", "PaieSage", "Path", xName, paramPaieSage_Path)
Call lstErr_AddItem(lstErr, cmdContext, "Path :" & paramPaieSage_Path)

V = rsElpTable_Read("DRH", "PaieSage", "Dossier", xName, paramPaieSage_Dossier)
Call lstErr_AddItem(lstErr, cmdContext, "Dossier :" & paramPaieSage_Dossier)

V = rsElpTable_Read("DRH", "PaieSage", "Archive", xName, paramPaieSage_Archive)
Call lstErr_AddItem(lstErr, cmdContext, "Archive :" & paramPaieSage_Archive)

V = rsElpTable_Read("DRH", "PaieSage", "SAB_VIR", xName, paramPaieSage_SAB_VIR)
Call lstErr_AddItem(lstErr, cmdContext, "SAB_VIR :" & paramPaieSage_SAB_VIR)

V = rsElpTable_Read("DRH", "PaieSage", "SAB_CPT", xName, paramPaieSage_SAB_CPT)

If paramEnvironnement = constTest Then
    X = Mid$(paramPaieSage_Virement, 3, Len(paramPaieSage_Virement) - 2)
    paramPaieSage_Virement = "D:\TEMP" & X
    X = Mid$(paramPaieSage_Comptabilité, 3, Len(paramPaieSage_Comptabilité) - 2)
    paramPaieSage_Comptabilité = "D:\TEMP" & X
    paramPaieSage_Path = "D:\"
    paramPaieSage_Dossier = "TEMP\PMSSYBEL"
End If

lstErr.Height = 500
Exit Function

Table_Error:
param_Init = V
Exit Function

End Function



Private Sub cmdExportVirement_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Set cmdExportVirement

End Sub


Private Sub cmdPrintComptabilité_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Set cmdPrintComptabilité

End Sub


Private Sub cmdPrintVirement_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Set cmdPrintVirement

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraImp_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraTrf_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub Timer1_Timer()
If prgBar.Value > prgBar.Max - 300 Then prgBar.Value = 0

prgBar.Value = prgBar.Value + 100

End Sub



Public Sub cmdPrint_Kill(xFileName As String)
Dim Nb As Long, Iter As Integer

Iter = 0
If Nb > 0 Then prgBar.Max = 100
Do
    DoEvents
    X = Dir(xFileName)
    Iter = Iter + 1
'    prgBar.Value = Iter
    If Iter > 30000 Then
        X = MsgBox("Voulez-vous réessayer ?", vbQuestion, "LrBafi : cmdTransfert : FTP en cours ")
        If X = vbYes Then
            Iter = 0
        Else
            Err = 9999: Exit Sub
        End If
    End If
Loop While X = ""
Call lstErr_AddItem(lstErr, cmdPrintVirement, "Impression " & xFileName)

prtPaieSageX xFileName

Call lstErr_AddItem(lstErr, cmdPrintVirement, "Suppression " & xFileName)
X = Dir(xFileName)
If X <> "" Then Kill xFileName

End Sub

Public Sub cmdArchiveDelete_False()
cmdArchiveDelete.Enabled = False

cmdArchiveDelete.Caption = "Sélectionner une sauvegarde à supprimer"

End Sub
