VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmTI2000 
   Caption         =   "TI : interface"
   ClientHeight    =   6345
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   9330
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1200
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4920
      TabIndex        =   0
      Top             =   0
      Width           =   4305
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5700
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   10054
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Interface DB2"
      TabPicture(0)   =   "TI2000.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraFolder"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraPrint"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "TI2000.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdfgCDComd"
      Tab(1).Control(1)=   "txtDossier"
      Tab(1).Control(2)=   "fgCDComD"
      Tab(1).ControlCount=   3
      Begin VB.Frame fraPrint 
         Caption         =   "Impression"
         Height          =   1815
         Left            =   360
         TabIndex        =   16
         Top             =   3720
         Width           =   8535
         Begin VB.OptionButton optPrintTIMt226 
            Caption         =   "Diff Util/ Paie < 1 %"
            Height          =   375
            Left            =   6000
            TabIndex        =   33
            Top             =   1320
            Width           =   2655
         End
         Begin VB.CheckBox chkSelValidité 
            Caption         =   "< 01-12-00"
            Height          =   255
            Left            =   2760
            TabIndex        =   31
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtprtStatut 
            Height          =   285
            Left            =   2160
            TabIndex        =   29
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox txtDevise 
            Height          =   285
            Left            =   2160
            TabIndex        =   27
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox txtCDComD_Type 
            Height          =   285
            Left            =   2160
            TabIndex        =   26
            Text            =   "RI"
            Top             =   360
            Width           =   495
         End
         Begin VB.OptionButton optPrintSituationI 
            Caption         =   "non soldé (TI=36)"
            Height          =   375
            Left            =   4080
            TabIndex        =   25
            Top             =   600
            Width           =   1815
         End
         Begin VB.OptionButton optPrintSituationA 
            Caption         =   "Soldé (TI=36)"
            Height          =   375
            Left            =   4080
            TabIndex        =   24
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optPrintSituationDif 
            Caption         =   "TI <> 36"
            Height          =   375
            Left            =   4080
            TabIndex        =   21
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton cmdPrint 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Print Test"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   6000
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   720
            Width           =   2160
         End
         Begin VB.OptionButton optPrintCommission 
            Caption         =   "Commission"
            Height          =   255
            Left            =   360
            TabIndex        =   18
            Top             =   360
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.OptionButton optPrintSituation 
            Caption         =   "Situation"
            Height          =   375
            Left            =   4080
            TabIndex        =   17
            Top             =   1320
            Width           =   975
         End
         Begin MSComCtl2.DTPicker txtAmjValidité 
            Height          =   300
            Left            =   6840
            TabIndex        =   34
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            CalendarBackColor=   16777215
            CalendarForeColor=   0
            CalendarTitleBackColor=   8421504
            CalendarTitleForeColor=   16777215
            CalendarTrailingForeColor=   12632256
            CustomFormat    =   "dd  MM yyy"
            Format          =   28246019
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin VB.Label lblAmjValidité 
            Caption         =   "Validité"
            Height          =   255
            Left            =   6000
            TabIndex        =   35
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblprtStatut 
            Caption         =   "Statut: , $I,$A"
            Height          =   255
            Left            =   600
            TabIndex        =   30
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label lblDevise 
            Caption         =   "Devise"
            Height          =   255
            Left            =   600
            TabIndex        =   28
            Top             =   840
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdfgCDComd 
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
         Height          =   615
         Left            =   -67320
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   960
      End
      Begin VB.TextBox txtDossier 
         Height          =   375
         Left            =   -73200
         TabIndex        =   14
         Top             =   480
         Width           =   1935
      End
      Begin VB.Frame fraFolder 
         Height          =   5175
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   8895
         Begin VB.CommandButton cmdCDComD_Reprise 
            BackColor       =   &H00C0FFFF&
            Caption         =   "CDComD_Reprise"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   6600
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   1800
            Width           =   2160
         End
         Begin VB.TextBox txtDossierAExclure 
            Height          =   285
            Left            =   2280
            TabIndex        =   22
            Text            =   "D:\Temp\TI\DossierAExclure.txt"
            Top             =   240
            Width           =   4095
         End
         Begin VB.CommandButton cmdImportS36 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Import S36"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   6600
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   960
            Width           =   2160
         End
         Begin VB.CheckBox chkTIDB2Mdb 
            Caption         =   "table.mdb"
            Height          =   255
            Left            =   2280
            TabIndex        =   10
            Top             =   2280
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox chkTIDB2Xls 
            Caption         =   "Output File"
            Height          =   255
            Left            =   2280
            TabIndex        =   9
            Top             =   1800
            Width           =   1455
         End
         Begin VB.TextBox txtTIDB2File 
            Height          =   285
            Left            =   2280
            TabIndex        =   7
            Text            =   "Master"
            Top             =   1320
            Width           =   3015
         End
         Begin VB.TextBox txtTIDB2Path 
            Height          =   285
            Left            =   2280
            TabIndex        =   6
            Text            =   "V:\TI_20010713\"
            Top             =   840
            Width           =   3015
         End
         Begin VB.CommandButton cmdOk_TIDB2_Load 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Ok : Master + CalcText + Posting+S36"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   3960
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   1800
            Width           =   2160
         End
         Begin VB.CommandButton cmdOK_TIDB2 
            BackColor       =   &H00C0FFFF&
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
            Height          =   615
            Left            =   6600
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   240
            Width           =   2160
         End
         Begin MSComCtl2.DTPicker txtAMJSituation 
            Height          =   300
            Left            =   2280
            TabIndex        =   11
            Top             =   2760
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            CalendarBackColor=   16777215
            CalendarForeColor=   0
            CalendarTitleBackColor=   8421504
            CalendarTitleForeColor=   16777215
            CalendarTrailingForeColor=   12632256
            CustomFormat    =   "dd  MM yyy"
            Format          =   28246019
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin VB.Label lblDossierAExclure 
            Caption         =   "File des dossiers à exclure"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label lblAMJSituation 
            Caption         =   "Date situation"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label lblTIDB2Input 
            Caption         =   "Input File"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   1320
            Width           =   1575
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgCDComD 
         Height          =   3570
         Left            =   -74640
         TabIndex        =   13
         Top             =   1800
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   6297
         _Version        =   393216
         Rows            =   1
         Cols            =   13
         FixedCols       =   0
         RowHeightMin    =   350
         BackColor       =   14737632
         ForeColor       =   12582912
         ForeColorFixed  =   -2147483641
         BackColorSel    =   12648384
         BackColorBkg    =   14737632
         AllowBigSelection=   0   'False
         TextStyleFixed  =   4
         FocusRect       =   2
         HighLight       =   0
         GridLines       =   2
         AllowUserResizing=   3
         FormatString    =   $"TI2000.frx":0038
      End
   End
End
Attribute VB_Name = "frmTI2000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String
Dim blnMsgBox_Quit As Boolean
Dim x As String, X1 As String, I As Integer, Nb As Integer
Dim Msg As String, valX As String
Dim currentMethod As String, lastMethod As String

Dim IdShell


Dim fgCDComD_FormatString As String, fgCDComD_K As Integer
Dim fgCDComD_RowDisplay As Integer, fgCDComD_RowClick As Integer
Dim fgCDComD_ColorClick As Long, fgCDComD_ColorDisplay As Long
Dim fgCDComD_Sort1 As Integer, fgCDComD_Sort2 As Integer

Dim wCDComD As typeCDComD, mCDComD As typeCDComD
Dim recMvtp0 As typeMvtP0
Dim zCDDossier As typeCDDossier

Public Sub fgCDComD_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer
mRow = fgCDComD.Row

If lRow > 0 Then
    fgCDComD.Row = lRow
    fgCDComD.Col = 0: fgCDComD.CellBackColor = lColor_Old 'fgCDComD.BackColorBkg
    fgCDComD.Col = 1: fgCDComD.CellBackColor = lColor_Old
    fgCDComD.Col = 2: fgCDComD.CellBackColor = lColor_Old
End If
lRow = 0
If mRow > 0 Then
    fgCDComD.Row = mRow
    If fgCDComD.Row > 0 Then
        lRow = fgCDComD.Row
        lColor_Old = fgCDComD.CellBackColor
        fgCDComD.Col = 0: fgCDComD.CellBackColor = lColor
        fgCDComD.Col = 1: fgCDComD.CellBackColor = lColor
        fgCDComD.Col = 2: fgCDComD.CellBackColor = lColor
    End If
End If

End Sub
Public Sub fgCDComD_DisplayLine()
fgCDComD.Col = 0: fgCDComD.Text = wCDComD.Dossier
fgCDComD.Col = 1: fgCDComD.Text = dateImp(wCDComD.AmjD)
fgCDComD.Col = 2: fgCDComD.Text = dateImp(wCDComD.AmjF)
fgCDComD.Col = 3: fgCDComD.Text = wCDComD.Devise
fgCDComD.Col = 4: fgCDComD.Text = Format$(wCDComD.MvtEngagement, "### ### ### ###.00")
fgCDComD.Col = 5: fgCDComD.Text = Format$(wCDComD.MvtUtilisé, "### ### ### ###.00")
fgCDComD.Col = 6: fgCDComD.Text = Format$(wCDComD.MontantBase, "### ### ### ###.00")
fgCDComD.Col = 7: fgCDComD.Text = Format$(wCDComD.CommissionTaux, "###.00")
fgCDComD.Col = 8: fgCDComD.Text = Format$(wCDComD.CommissionD, "### ### ### ###.00")
fgCDComD.Col = 9: fgCDComD.Text = Format$(wCDComD.CommissionP, "### ### ### ###.00")
fgCDComD.Col = 10: fgCDComD.Text = dateImp(wCDComD.CommissionPAmj)
fgCDComD.Col = 11: fgCDComD.Text = wCDComD.TIChargeKey
fgCDComD.Col = 12: fgCDComD.Text = Format$(wCDComD.CoursEur, "### ###.00000")

End Sub


Public Sub fgCDComD_Load()
SSTab1.Tab = 1

fgCDComD.Visible = True
fgCDComD.Clear: fgCDComD.Rows = 1

fgCDComD.Rows = 1
fgCDComD.FormatString = fgCDComD_FormatString
fgCDComD.Enabled = True

wCDComD = mCDComD
wCDComD.Method = "Seek>="
dbCDComD_ReadE wCDComD

Do
    If wCDComD.Dossier = mCDComD.Dossier Then
        fgCDComD.Rows = fgCDComD.Rows + 1
        fgCDComD.Row = fgCDComD.Rows - 1
        fgCDComD_DisplayLine
        
    End If
    
    wCDComD.Method = "MoveNext"
    intReturn = tableCDComD_Read(wCDComD)
  
Loop While intReturn = 0 And wCDComD.Dossier = mCDComD.Dossier

If fgCDComD.Rows > 1 Then fgCDComD_Sort

End Sub

Public Sub fgCDComD_Sort()
If fgCDComD.Rows > 1 Then
    fgCDComD.Row = 1
    fgCDComD.RowSel = fgCDComD.Rows - 1
    
    fgCDComD.Col = fgCDComD_Sort1
    fgCDComD.ColSel = fgCDComD_Sort2
    fgCDComD.Sort = 1
End If
End Sub

Public Sub Msg_Rcv(x As String)
'---------------------------------------------------------
End Sub


Public Sub cmdContext_Quit()
    If blnMsgBox_Quit Then
       x = MsgBox("Voulez-vous réellement abandonner?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
    Else
       x = vbYes
    End If
    If x = vbYes Then Unload Me

End Sub


Public Sub cmdContext_Return()

SendKeys "{TAB}"

End Sub

'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor
currentActiveControl_Name = C.Name
End Sub

'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------
lstErr.Clear
arrTag(Val(C.Tag)) = True
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor
End Sub

Private Sub cmdCDComD_Reprise_Click()
Dim x As String
Me.Enabled = False
Screen.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdPrint, "cmdCDComD_Reprise : début ...")
DoEvents
TIDB2_CDComD_Reprise x
Call lstErr_AddItem(lstErr, cmdPrint, "cmdCDComD_Reprise : fin _" & x)
Me.Enabled = True
AppActivate Me.Caption
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdContext_Click()
cmdContext_Quit

End Sub

Private Sub cmdfgCDComd_Click()
mdbCDComD.tableCDComD_Open
recCDComD_Init mCDComD
mCDComD.Method = "Seek>="
mCDComD.Dossier = txtDossier
fgCDComD_Load
End Sub

Private Sub cmdImportS36_Click()
cmdImportS36_Load
End Sub

Private Sub cmdOK_TIDB2_Click()
paramTI2000DB2_DossierAExclure = Trim(txtDossierAExclure)

x = Trim(txtTIDB2Path) & Trim(txtTIDB2File)
paramTI2000DB2_Table = UCase$(Trim(txtTIDB2File))
paramTI2000DB2_Input = x & ".txt"
paramTI2000DB2_Output = x & "_Xls.txt"
If chkTIDB2Xls = "1" Then paramTI2000DB2_Xls = True
If chkTIDB2Mdb = "1" Then paramTI2000DB2_mdb = True

Me.MousePointer = vbHourglass
Me.Enabled = False
srvTI2000.TIDB2_Load
Me.MousePointer = 0
Me.Enabled = True
End Sub

Private Sub cmdImportS36_Load()
Dim xInput As String, blnOk As Boolean, xFile As String
Dim curX As Currency
Dim wDossier As Long
Dim wCodif As String

On Error Resume Next

Dim I As Integer
xFile = Trim(txtTIDB2Path) & "SrvCredocP"
x = Dir(xFile)
If x = "" Then Call lstErr_Clear(lstErr, cmdPrint, "? Le fichier SRvCredocP n'existe pas"): Exit Sub

Call lstErr_Clear(lstErr, cmdPrint, " cmdImportS36_Load  ..."): DoEvents
Me.MousePointer = vbHourglass
Me.Enabled = False

mdbCDDossier.tableCDDossier_Open

Open xFile For Input As #1
recCDDossier_Init zCDDossier
recCDDossier = zCDDossier
recCDDossier.Method = "Seek="

blnOk = False
Call lstErr_AddItem(lstErr, cmdPrint, "Chargement  ..."): DoEvents

Do Until EOF(1)
    Line Input #1, xInput
        If mId$(xInput, 1, 3) = "CDE" Then
            wDossier = CLng(Val(mId$(xInput, 7, 5)))
        Else
            wDossier = CLng(Val(mId$(xInput, 5, 5)))
            If mId$(xInput, 4, 1) = "6" Then wDossier = wDossier + 100000
        End If
    'If wDossier = "61299" Then
    '    X = ""
    'End If
    
    If recCDDossier.Dossier <> wDossier Then
        If blnOk Then dbCDDossier_Update recCDDossier
        recCDDossier.Dossier = Format$(wDossier, "00000")
        recCDDossier.Method = "Seek="
        If tableCDDossier_Read(recCDDossier) = 0 Then
            recCDDossier.Method = "Update": blnOk = True
            recCDDossier.S36RC = 0
            recCDDossier.S36RE = 0
            recCDDossier.S36RI = 0
            recCDDossier.S36RA = 0
            recCDDossier.S36Engagement = 0
            recCDDossier.S36Utilisé = 0
        Else
            recCDDossier = zCDDossier
            recCDDossier.Dossier = Format$(wDossier, "00000")
            recCDDossier.AMJSituation = "$I     "
            recCDDossier.Method = "AddNew"
            blnOk = True
            Call lstErr_ChangeLastItem(lstErr, cmdPrint, "? inconnu : " & wDossier)
        End If
    End If
    curX = CCur(Val(mId$(xInput, 24, 19)))
    If mId$(xInput, 43, 1) = "D" Then
        recCDDossier.S36Engagement = recCDDossier.S36Engagement + curX
    Else
        recCDDossier.S36Utilisé = recCDDossier.S36Utilisé + curX
        curX = -curX
    End If
    
    Select Case mId$(xInput, 44, 2)
        Case "38": recCDDossier.S36RC = recCDDossier.S36RC + curX
        Case "39": recCDDossier.S36RE = recCDDossier.S36RE + curX
        Case "50": recCDDossier.S36RI = recCDDossier.S36RI + curX
        Case "45": recCDDossier.S36RA = recCDDossier.S36RA + curX
        Case Else: MsgBox xInput, vbExclamation, "codif ?"
    End Select
    
Loop

Close
If blnOk Then recCDDossier.Method = "Update": dbCDDossier_Update recCDDossier
mdbCDDossier.tableCDDossier_Close
Call lstErr_AddItem(lstErr, cmdPrint, "Fin"): DoEvents
Me.MousePointer = 0
Me.Enabled = True
End Sub


Private Sub cmdOk_TIDB2_Load_Click()

txtTIDB2File = "Master": cmdOK_TIDB2_Click
txtTIDB2File = "CalcText": cmdOK_TIDB2_Click
txtTIDB2File = "Posting": cmdOK_TIDB2_Click
cmdImportS36_Load
End Sub


Private Sub cmdPrint_Click()
Dim Msg As String
Msg = Space$(100)
Call DTPicker_Control(txtAMJSituation, paramTI2000DB2_AMJSituation)
Call DTPicker_Control(txtAmjValidité, paramTI2000DB2_AMJValidité)
Msg = "000001059999"
Msg = "000000999999"
selCDComD_Devise = UCase(Trim(txtDevise))

If optPrintCommission Then
    cmdPrint_Select Msg
   If Nb > 0 Then prtTI2000Commission_Monitor Msg
End If

If optPrintSituation Then
    Msg = "000000999999SG"
    prtTI2000_Monitor Msg
End If

If optPrintSituationDif Then
    'Msg = "63000 63199 SD"
    Msg = "000000999999SD"
    prtTI2000_Monitor Msg
End If
If optPrintSituationA Then
    Msg = "000000999999SA"
    prtTI2000_Monitor Msg
End If
If optPrintSituationI Then
    Msg = "000000999999SI"
    prtTI2000_Monitor Msg
End If
If optPrintTIMt226 Then
    Msg = "000000999999UP"
    prtTI2000_Monitor Msg
End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case Is = 13: KeyCode = 0:  cmdContext_Return
    Case Is = 27:  cmdContext_Quit
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select

End Sub

Private Sub Form_Load()
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
Call DTPicker_Set(txtAMJSituation, DSys) '"20001231") 'DSys)
Call DTPicker_Set(txtAmjValidité, DSys) '"20001231") 'DSys)
fgCDComD_Sort1 = 0: fgCDComD_Sort2 = 0
fgCDComD_FormatString = fgCDComD.FormatString

End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub

Private Sub fraFolder_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set SSTab1
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



Private Sub txtAMJSituation_GotFocus()
DTPicker_GotFocus txtAMJSituation

End Sub


Private Sub txtAMJSituation_LostFocus()
DTPicker_LostFocus txtAMJSituation

End Sub



Public Sub cmdPrint_Select(lMsg As String)
Dim mMax As Long, blnDevise As Boolean, blnType As Boolean
selprtStatut = UCase(Trim(txtprtStatut))
selCDComD_Devise = UCase(Trim(txtDevise))
If selCDComD_Devise = "" Then
    blnDevise = False
Else
    blnDevise = True
End If

selCDComD_Type = UCase(Trim(txtCDComD_Type))
If selCDComD_Type = "" Then
    blnType = False
Else
    blnType = True
End If
Call lstErr_AddItem(lstErr, cmdPrint, "Chargement des dossiers, tri ...")

MDB.Execute "delete * from MvtP0"
mdbMvtP0.tableMvtP0_Open

recMvtP0_Init recMvtp0
recMvtp0.Method = "AddNew"

mdbCDComD.tableCDComD_Open
recCDComD_Init wCDComD
wCDComD.Method = "Seek>="
wCDComD.Dossier = CLng(mId$(lMsg, 1, 6))
mMax = CLng(mId$(lMsg, 7, 6))

dbCDComD_ReadE wCDComD
I = 0: Nb = 0
Do
    If Not blnDevise Or wCDComD.Devise = selCDComD_Devise Then
        If Not blnType Or wCDComD.Type = selCDComD_Type Then
            If wCDComD.AmjD < paramTI2000DB2_AMJSituation Then
                
                Nb = Nb + 1
                recMvtp0.Id = wCDComD.Type & wCDComD.Devise & Format$(wCDComD.Dossier, "00000000000") & wCDComD.AmjD
                recMvtp0.Text = " " ' wCDComD
                dbMvtP0_Update recMvtp0
                
            End If
        End If
    End If
    wCDComD.Method = "MoveNext"
    intReturn = tableCDComD_Read(wCDComD)

Loop While intReturn = 0 And wCDComD.Dossier <= mMax

Call lstErr_AddItem(lstErr, cmdPrint, "Sélection : " & Nb)

End Sub

Private Sub txtCDComD_Type_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub


Private Sub txtDevise_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtprtStatut_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub


