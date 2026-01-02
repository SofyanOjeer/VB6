VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLrSgnBnf 
   AutoRedraw      =   -1  'True
   Caption         =   "LR Risques : bénéficiaires"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6375
   ScaleWidth      =   9420
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   400
      Left            =   8880
      Picture         =   "LrSgnBnf.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   500
   End
   Begin MSFlexGridLib.MSFlexGrid fgLrSgnBnf 
      Height          =   5490
      Left            =   480
      TabIndex        =   2
      Top             =   600
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   9684
      _Version        =   393216
      Rows            =   1
      Cols            =   8
      FixedCols       =   0
      RowHeightMin    =   300
      BackColor       =   14737632
      ForeColor       =   12582912
      ForeColorFixed  =   -2147483641
      BackColorSel    =   12648384
      BackColorBkg    =   14737632
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   0
      GridLines       =   0
      GridLinesFixed  =   1
      FormatString    =   $"LrSgnBnf.frx":0102
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00C0C0C0&
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
      Height          =   400
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1200
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6240
      TabIndex        =   1
      Top             =   0
      Width           =   2500
   End
   Begin VB.Menu cmdPrintmnu 
      Caption         =   "cmdPrint"
      Visible         =   0   'False
      Begin VB.Menu cmdPrintmnuDétail 
         Caption         =   "Détail d'un bénéficiaire"
      End
      Begin VB.Menu cmdPrintmnuList 
         Caption         =   "Tous les bénéficiaires (tableau)"
      End
   End
End
Attribute VB_Name = "frmLrSgnBnf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String
Dim blnMsgBox_Quit As Boolean, autLrSgnBnf As typeAuthorization
Dim X As String, I As Integer
Dim Msg As String, valX As String
Dim currentMethod As String, lastMethod As String
Dim blnAddNew As Boolean

Dim recLrSgnBnf As typeLrSgnBnf
Dim fgLrSgnBnf_FormatString As String, fgLrSgnBnf_K As Integer
Dim fgLrSgnBnf_BackColorFixed As Long, fgLrSgnBnf_BackColor As Long

Dim LrSgnBnf_Name As String, LrSgnBnf_Value As String
Dim fgLrSgnBnf_Col As Integer, fgLrSgnBnf_Colsel As Integer

Dim recAccAut As typeAccAut

Private Sub cmdContext_Click()
Select Case cmdContext.Caption
    Case Is = constcmdRechercher
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdPrint_Click()
cmdPrintmnuDétail.Enabled = False
cmdPrintmnuList.Enabled = True
Me.PopupMenu cmdPrintmnu, vbPopupMenuRightButton
End Sub

Private Sub cmdPrintmnuDétail_Click()
Dim X As String
fgLrSgnBnf_Scan
X = Format$(arrLrSgnBnfIndex, "000000") & Format$(arrLrSgnBnfIndex, "000000") & "L" ' & "D"
prtLrSgnBnfX X
End Sub

Private Sub cmdPrintmnuList_Click()
X = Format$(1, "000000") & Format$(arrLrSgnBnfNb, "000000") & "L"

prtLrSgnBnfX X

End Sub


Private Sub fgLrSgnBnf_Click()
lstErr.Clear
fgLrSgnBnf_K = fgLrSgnBnf.Row * fgLrSgnBnf.Cols
If fgLrSgnBnf.Row > 0 Then
    cmdPrintmnuDétail.Enabled = True
    cmdPrintmnuList.Enabled = True
    Me.PopupMenu cmdPrintmnu, vbPopupMenuRightButton
End If
End Sub

'---------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------
Select Case KeyCode
    Case Is = 13: KeyCode = 0: cmdContext_Return
    Case Is = 27: cmdContext_Quit
    Case Is = 44: frmElpPrt.prtScreen
End Select

End Sub


'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)

Call srvUsrAut("LrSgnBnf", autLrSgnBnf)
Form_Clear
ReDim arrLrSgnBnf(1): arrLrSgnBnfNbMax = 1: arrLrSgnBnfNb = 0
srvLrsGNbNF.Init recLrSgnBnf
fgLrSgnBnf_FormatString = fgLrSgnBnf.FormatString
fgLrSgnBnf_BackColorFixed = fgLrSgnBnf.BackColorFixed
fgLrSgnBnf_BackColor = fgLrSgnBnf.BackColor
'If DeviseCoursAut.Détail Then
    blnAddNew = True
    fgLrSgnBnf_Load
'End If
fgLrSgnBnf_Col = 0: fgLrSgnBnf_Colsel = 1
fgLrSgnBnf_Load
'LrSgnBnf_Name = cboLrSgnBnf.List(1)

End Sub



'---------------------------------------------------------
Public Sub Form_Clear()
'---------------------------------------------------------
lstErrClear = True
blnMsgBox_Quit = False
usrColor_Set

cmdContext.Enabled = True: cmdContext.BackColor = vbWindowBackground
cmdContext.Caption = constcmdAbandonner: cmdContext.BackColor = errUsr.BackColor
fgLrSgnBnf.Enabled = True: fgLrSgnBnf.Clear: fgLrSgnBnf.Rows = 1
Call lstErr_Clear(lstErr, cmdContext, " 'click' Recherche")
End Sub




Public Sub Msg_Rcv(X As String)
'---------------------------------------------------------

End Sub

Private Sub fgLrSgnBnf_Load()

Dim blnValidation As Boolean, blnSaisie As Boolean, X As String
srvLrsGNbNF.Init recLrSgnBnf
currentMethod = "SnapP0"
recLrSgnBnf.Method = currentMethod
recLrSgnBnf.CDBANQ = strSocBdfE
recLrSgnBnf.CDDECL = strSocBdfE
arrLrSgnBnf(0) = recLrSgnBnf
arrLrSgnBnf(0).RFBENF = "999999"
arrLrSgnBnfNb = 0: arrLrSgnBnfIndex = 0
arrLrSgnBnfSuite = True

Do Until Not arrLrSgnBnfSuite
    srvLrsGNbNF.Monitor recLrSgnBnf
    recLrSgnBnf = arrLrSgnBnf(arrLrSgnBnfNb)
    recLrSgnBnf.Method = currentMethod & "+"
'' arrLrSgnBnfSuite = False ''$$$$$$$$$$$$$$$$$$
Loop
fgLrSgnBnf_Display
Call lstErr_Clear(lstErr, fgLrSgnBnf, "ok")

End Sub

Public Sub fgLrSgnBnf_Display()
fgLrSgnBnf.Redraw = False
fgLrSgnBnf.Clear
fgLrSgnBnf.Rows = 1
fgLrSgnBnf.FormatString = fgLrSgnBnf_FormatString & "<" & LrSgnBnf_Name
fgLrSgnBnf.Enabled = True
For arrLrSgnBnfIndex = 1 To arrLrSgnBnfNb
    If arrLrSgnBnf(arrLrSgnBnfIndex).Method <> constDelete _
    And arrLrSgnBnf(arrLrSgnBnfIndex).Method <> constIgnore Then
        fgLrSgnBnf.Rows = fgLrSgnBnf.Rows + 1
        fgLrSgnBnf.Row = fgLrSgnBnf.Rows - 1
        fgLrSgnBnf_DisplayItem
    End If
Next arrLrSgnBnfIndex
If fgLrSgnBnf.Rows > 1 Then fgLrSgnBnf_Sort
fgLrSgnBnf.Redraw = True

End Sub

Public Sub fgLrSgnBnf_DisplayItem()
fgLrSgnBnf_K = (fgLrSgnBnf.Row) * fgLrSgnBnf.Cols
fgLrSgnBnf.TextArray(0 + fgLrSgnBnf_K) = arrLrSgnBnf(arrLrSgnBnfIndex).RFBENF
fgLrSgnBnf.TextArray(1 + fgLrSgnBnf_K) = arrLrSgnBnf(arrLrSgnBnfIndex).NBDF2
fgLrSgnBnf.TextArray(2 + fgLrSgnBnf_K) = arrLrSgnBnf(arrLrSgnBnfIndex).NOMBNF
End Sub

Public Sub cmdContext_Quit()

Unload Me

End Sub

Public Sub cmdContext_Return()
SendKeys "{TAB}"

End Sub

Public Sub fgLrSgnBnf_Scan()
fgLrSgnBnf_K = fgLrSgnBnf.Row * fgLrSgnBnf.Cols
recLrSgnBnf.RFBENF = Trim(fgLrSgnBnf.TextArray(0 + fgLrSgnBnf_K))
If srvLrsGNbNF.Scan(recLrSgnBnf) > 0 Then
    recLrSgnBnf = arrLrSgnBnf(arrLrSgnBnfIndex)
Else
    Call lstErr_AddItem(lstErr, fgLrSgnBnf, "Erreur fgLrSgnBnf_Scan")
End If
End Sub


Public Sub fgLrSgnBnf_Sort()
fgLrSgnBnf.Row = 1
fgLrSgnBnf.Col = fgLrSgnBnf_Col

fgLrSgnBnf.RowSel = 1 'fgLrSgnBnf.Rows - 1
fgLrSgnBnf.ColSel = fgLrSgnBnf_Colsel

fgLrSgnBnf.Sort = flexSortStringAscending
End Sub





