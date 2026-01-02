VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmElpTable 
   AutoRedraw      =   -1  'True
   Caption         =   "Table : mise à jour"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13875
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9495
   ScaleWidth      =   13875
   Begin TabDlg.SSTab SSTab1 
      Height          =   9015
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   15901
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Mise à jour"
      TabPicture(0)   =   "ElpTable.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fgElpTable"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lstId"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraElpTable"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Maintenance"
      TabPicture(1)   =   "ElpTable.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtImportExport"
      Tab(1).Control(1)=   "cmdExport"
      Tab(1).Control(2)=   "cmdImport"
      Tab(1).Control(3)=   "cmdDataBase_Replication"
      Tab(1).Control(4)=   "cmdDataBase_Compact"
      Tab(1).Control(5)=   "cmdDataBase_ReplaceMaster"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "ElpTable.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdTest"
      Tab(2).ControlCount=   1
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test "
         Height          =   735
         Left            =   -72360
         TabIndex        =   23
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CommandButton cmdDataBase_ReplaceMaster 
         BackColor       =   &H000000FF&
         Caption         =   "Mdb REPLACE Master"
         Height          =   780
         Left            =   -74640
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   3600
         Width           =   3495
      End
      Begin VB.CommandButton cmdDataBase_Compact 
         BackColor       =   &H008080FF&
         Caption         =   "Mdb CompactDataBase locale"
         Height          =   780
         Left            =   -74640
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2160
         Width           =   3495
      End
      Begin VB.CommandButton cmdDataBase_Replication 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Mdb Replication Master=> Local"
         Height          =   780
         Left            =   -74640
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   960
         Width           =   3495
      End
      Begin VB.CommandButton cmdImport 
         BackColor       =   &H000000FF&
         Caption         =   "Delete * et Import  ELPTABLE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -65880
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2760
         Width           =   3240
      End
      Begin VB.CommandButton cmdExport 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Export ElpTable"
         Height          =   735
         Left            =   -66000
         MaskColor       =   &H80000000&
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox txtImportExport 
         Height          =   375
         Left            =   -65880
         TabIndex        =   17
         Text            =   "D:\BiaSrc\Dta\"
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Frame fraElpTable 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   2040
         TabIndex        =   6
         Top             =   6480
         Width           =   11655
         Begin VB.TextBox txtName 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            MaxLength       =   40
            TabIndex        =   9
            Top             =   1440
            Width           =   10455
         End
         Begin VB.CommandButton fraElpTable_cmdOk 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Ok"
            Height          =   1125
            Left            =   7800
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   240
            Width           =   1500
         End
         Begin VB.CommandButton fraElpTable_cmdQuit 
            BackColor       =   &H00C0C0FF&
            Caption         =   "X"
            Height          =   1125
            Left            =   9960
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   240
            Width           =   1500
         End
         Begin VB.TextBox txtMemo 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1095
            TabIndex        =   10
            Top             =   1920
            Width           =   10320
         End
         Begin VB.TextBox txtK2 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            MaxLength       =   12
            TabIndex        =   8
            Top             =   840
            Width           =   1935
         End
         Begin VB.TextBox txtK1 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            MaxLength       =   12
            TabIndex        =   7
            Top             =   345
            Width           =   1935
         End
         Begin VB.Label lblName 
            Caption         =   "Intitulé"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label lblMemo 
            Caption         =   "Mémo"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   2040
            Width           =   855
         End
         Begin VB.Label lblK1 
            Caption         =   "Clé 1"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lblK2 
            Caption         =   "Clé 2"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   960
            Width           =   735
         End
      End
      Begin VB.ListBox lstId 
         Height          =   8250
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   480
         Width           =   1935
      End
      Begin MSFlexGridLib.MSFlexGrid fgElpTable 
         Height          =   5685
         Left            =   2040
         TabIndex        =   4
         Top             =   600
         Width           =   11745
         _ExtentX        =   20717
         _ExtentY        =   10028
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
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
         FormatString    =   $"ElpTable.frx":0054
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   400
      Left            =   13320
      Picture         =   "ElpTable.frx":011A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   500
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9960
      TabIndex        =   1
      Top             =   0
      Width           =   3345
   End
   Begin VB.Label lblDataBaseName 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   405
      Left            =   2280
      TabIndex        =   2
      Top             =   0
      Width           =   7095
   End
   Begin VB.Menu mnuElpTable 
      Caption         =   "Table"
      Visible         =   0   'False
      Begin VB.Menu mnuElpTableAddNew 
         Caption         =   "Ajouter un enregistrement"
      End
      Begin VB.Menu mnuElpTableCopy 
         Caption         =   "Copier un enregistrement"
      End
      Begin VB.Menu mnuElpTableX1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuElpTableModify 
         Caption         =   "Modifier un enregistrement"
      End
      Begin VB.Menu mnuElpTableDelete 
         Caption         =   "Supprimer un enregistrement"
      End
      Begin VB.Menu mnuElpTableX2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuElpTablePrint 
         Caption         =   "Impression"
      End
   End
End
Attribute VB_Name = "frmElpTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'à faire : gérer SN SNP SNN Chrono
'          gérer DMin Dmax
'          zones spécifiques pour mémo
'          autorisation màj par table
'           modification K1 K2 (delete , addnew)
' création d'un service ==> créer le répertoire
' personnaliser les saisies
' imprimer toutes les tables


'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim mHeight_0 As Integer, mWidth_0 As Integer, mHeight_2 As Integer, mWidth_2 As Integer, mWindowState As Integer
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String
Dim blnMsgBox_Quit As Boolean
Dim TableAut As typeAuthorization
Dim X As String, X1 As String, I As Integer
Dim Msg As String, valX As String, V As Variant

Dim recElpTable As typeElpTable, xElpTable As typeElpTable
Dim currentMethod As String, currentAMJ As String, currentId As String
Dim fgElpTable_FormatString As String, fgElpTable_K As Integer
Dim fgElpTable_BackColorFixed As Long, fgElpTable_BackColor As Long
Dim fgElpTable_TopRow As Long
Dim blnAddNew As Boolean
Dim AMJDéfaut As String
Dim arrElpTable() As typeElpTable, arrElpTable_Nb As Integer, arrElpTable_index As Integer

Private Sub cmdDataBase_Compact_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdPrint, " > cmdDataBase_Compact :" & DataBase_Local)

MDB_CompactDataBase

Call lstErr_AddItem(lstErr, cmdPrint, " Database_Open : " & DataBase_Open)
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdDataBase_ReplaceMaster_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdPrint, " > cmdDataBase_Replace :" & DataBase_Master & " => " & DataBase_Local)

MDB_ReplaceMaster

Call lstErr_AddItem(lstErr, cmdPrint, " Database_Open : " & DataBase_Open)
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdDataBase_Replication_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdPrint, " > cmdDataBase_Replication :" & DataBase_Local & " => " & DataBase_Master)

MDB_Replication

Call lstErr_AddItem(lstErr, cmdPrint, " Database_Open : " & DataBase_Open)
Me.Enabled = True: Me.MousePointer = 0

End Sub

'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
''cmdControl
C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor
currentActiveControl_Name = C.Name
End Sub
'---------------------------------------------------------
Public Sub lstId_Display()
'---------------------------------------------------------
arrElpTable_Load "Table"

lstId.Clear
For I = 1 To arrElpTable_Nb
    lstId.AddItem arrElpTable(I).K1
Next I

End Sub

'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------
arrTag(Val(C.Tag)) = True
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor
End Sub



Private Sub cmdExport_Click()
Dim X As String, K As Integer
Dim intFile As Integer
On Error GoTo Error_Handler


Me.Enabled = False
Me.MousePointer = vbHourglass
intFile = FreeFile(0)
Open Trim(txtImportExport) For Output As #intFile

Call lstErr_Clear(lstErr, cmdPrint, "Export"): DoEvents

X = "select * from ElpTable "
    
Set rsMDB = cnMDB.Execute(X)
Do While Not rsMDB.EOF
    Call rsElpTable_GetBuffer(rsMDB, recElpTable)

        If IsNull(recElpTable.Memo) Then
            K = 0
        Else
            K = Len(recElpTable.Memo)
        End If
        X = Space$(137 + K)
        Mid$(X, 1, 12) = recElpTable.Id
        Mid$(X, 13, 12) = recElpTable.K1
        Mid$(X, 25, 12) = recElpTable.K2
        Mid$(X, 37, 12) = Format(recElpTable.SNN, "###########0")
        Mid$(X, 49, 12) = Format(recElpTable.SNP, "###########0")
        Mid$(X, 61, 12) = Format(recElpTable.SN, "###########0")
        Mid$(X, 73, 12) = Format(recElpTable.Chrono, "###########0")
        Mid$(X, 85, 36) = recElpTable.Name
        Mid$(X, 121, 8) = recElpTable.Dmin
        Mid$(X, 129, 8) = recElpTable.Dmax
        If K > 0 Then Mid$(X, 137, K) = recElpTable.Memo
        Print #intFile, Trim(X)
    rsMDB.MoveNext
Loop


Close #intFile
Call lstErr_AddItem(lstErr, cmdPrint, "Export : fin"): DoEvents

Me.MousePointer = 0
Me.Enabled = True

Exit Sub
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------


Call MsgBox(Err & " : " & Error(Err), vbCritical, "Export")
Me.MousePointer = 0
Me.Enabled = True

End Sub

Private Sub cmdImport_Click()
Dim xInput As String, K As Integer
Dim V
On Error GoTo Error_Handler


Me.Enabled = False
Me.MousePointer = vbHourglass
Dim intFile As Integer
intFile = FreeFile(0)
Open Trim(txtImportExport) For Input As #intFile

Call lstErr_Clear(lstErr, cmdPrint, "Import"): DoEvents

rsMDB.Close
Call FEU_ROUGE
Set rsMDB = cnMDB.Execute("delete * from ElpTable")
Call FEU_VERT
rsMDB.Open "select * from ElpTable", cnMDB, , adLockOptimistic

Do Until EOF(1)
    Line Input #intFile, xInput
        recElpTable.Id = mId$(xInput, 1, 12)
        recElpTable.K1 = mId$(xInput, 13, 12)
        recElpTable.K2 = mId$(xInput, 25, 12)
        recElpTable.SNN = Val(mId$(xInput, 37, 12))
        recElpTable.SNP = Val(mId$(xInput, 49, 12))
        recElpTable.SN = Val(mId$(xInput, 61, 12))
        recElpTable.Chrono = Val(mId$(xInput, 73, 12))
        recElpTable.Name = mId$(xInput, 85, 36)
        recElpTable.Dmin = mId$(xInput, 121, 8)
        recElpTable.Dmax = mId$(xInput, 129, 8)
        K = Len(xInput)
        If K = 136 Then
            recElpTable.Memo = "" 'Null
        Else
            recElpTable.Memo = mId$(xInput, 137, K - 136)
        End If
        V = adoElpTable_AddNew(rsMDB, recElpTable)
        If Not IsNull(V) Then GoTo Error_MsgBox
Loop

Close #intFile
Call lstErr_AddItem(lstErr, cmdPrint, "Import : fin"): DoEvents

Me.MousePointer = 0
Me.Enabled = True

Exit Sub
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------


Call MsgBox(Err & " : " & Error(Err), vbCritical, "Import")
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & "fraElpTable_cmdOk_Click :"
Me.MousePointer = 0
Me.Enabled = True
Close

End Sub


'---------------------------------------------------------
Private Sub cmdPrint_Click()
'---------------------------------------------------------
cmdPrintX ""
End Sub

Private Sub cmdTest_Click()
Dim xSql As String, K As Integer
xSql = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0 where BASTABNUM = 37 and BASTABARG like 'DZD%' "
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    Debug.Print rsSab("BASTABARG"), rsSab("BASTABDON")
    For K = 1 To 16
        Debug.Print Asc(mId$(rsSab("BASTABARG"), K, 1));
    Next K
        Debug.Print
     For K = 1 To 16
        Debug.Print Asc(mId$(rsSab("BASTABDON"), K, 1));
    Next K
        Debug.Print "----------------------------------------------------------"
   
    rsSab.MoveNext
Loop

End Sub

Private Sub fgElpTable_Click()
lstErr.Clear
'mnuElpTableModify = False
'mnuElpTableDelete = False
fgElpTable_K = fgElpTable.Row * fgElpTable.Cols
If fgElpTable.Row > 0 Then
    arrElpTable_index = arrElpTable_Scan(currentId _
                        , (Trim(fgElpTable.TextArray(0 + fgElpTable_K))) _
                        , (Trim(fgElpTable.TextArray(1 + fgElpTable_K))))
    If arrElpTable_index > -1 Then
        'mnuElpTableModify = True
        'mnuElpTableDelete = True
        recElpTable = arrElpTable(arrElpTable_index)
    End If
End If

'fgElpTable.Col = 1: fgElpTable.CellBackColor = focusUsr.BackColor
Me.PopupMenu mnuElpTable, vbPopupMenuRightButton
'fgElpTable.Col = 1: fgElpTable.CellBackColor = fgElpTable_BackColor
End Sub

Private Sub fgElpTable_GotFocus()
fgElpTable.BackColorFixed = focusUsr.BackColor
fgElpTable.BackColor = fgElpTable_BackColor
End Sub


Private Sub fgElpTable_LostFocus()
fgElpTable.BackColorFixed = fgElpTable_BackColorFixed
'fgElpTable.BackColor = vbWindowBackground
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
mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
Call BiaPgmAut_Init("TABLE", TableAut)

'If TableAut.Saisir Then MDB_Master

mnuElpTableAddNew.Enabled = TableAut.Saisir
mnuElpTableDelete.Enabled = TableAut.Saisir
mnuElpTableModify.Enabled = TableAut.Saisir
cmdDataBase_Compact.Visible = TableAut.Saisir
cmdDataBase_ReplaceMaster.Visible = TableAut.Saisir
cmdDataBase_Replication.Visible = TableAut.Saisir

currentAMJ = DSys
cmdClear
fgElpTable_FormatString = fgElpTable.FormatString
fgElpTable_BackColorFixed = fgElpTable.BackColorFixed
fgElpTable_BackColor = fgElpTable.BackColor
If DataBase_Open = DataBase_Master Then
    lblDataBaseName.BackColor = vbRed
Else
    lblDataBaseName.BackColor = vbGreen
End If
'================================================================================
lblDataBaseName = DataBase_Open
'================================================================================

'tableElpTable_Open
currentId = "Table"
lstId_Display
txtImportExport = DataBase_Local & ".txt"

End Sub


Public Sub MDB_CompactDataBase()
Dim jro As jro.JetEngine
Dim xSource As String, xDestination As String
On Error GoTo Error_Handler
Dim X As String, xNew As String, xOld As String
Dim wFolder As String, wName As String, wExtension As String
    X = MsgBox("CompactDataBase : " & DataBase_Local, vbInformation + vbYesNo + vbDefaultButton2, "Elp : MDB_Local")
    If X = vbYes Then
        X = DataBase_Open
        If X <> "" Then MDB_Close
        DataBase_Open = ""
        
        Call fileName_Split(DataBase_Local, wFolder, wName, wExtension)
        xNew = wFolder & wName & "_New." & wExtension
        xOld = wFolder & wName & "_Old." & wExtension
        If Dir(xNew) <> "" Then Kill xNew
        
        'DBEngine.CompactDatabase DataBase_Local, xNew, , , paramDataBase_Password
        
        xSource = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DataBase_Local & ";Jet OLEDB:DataBase Password=" & paramDataBase_Password
        xDestination = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & xNew & ";Jet OLEDB:DataBase Password=" & paramDataBase_Password
        Set jro = New jro.JetEngine
        
        jro.CompactDatabase xSource, xDestination
        

        If Dir(xOld) <> "" Then Kill xOld
        Name DataBase_Local As xOld
        Name xNew As DataBase_Local
        Kill xOld
        
         If X <> "" Then MDB_Open X, paramDataBase_Password
    End If
    
Exit Sub
Error_Handler:
Shell_MsgBox "ELPVB_MDB_CompactDataBase : " & Error, vbCritical, frmElp_Caption, False
End Sub


'---------------------------------------------------------
Public Sub cmdClear()
'---------------------------------------------------------
arrTag_Set False
lstErrClear = True
blnMsgBox_Quit = False
usrColor_Set

fraElpTable_Clear: fraElpTable.Visible = False
fgElpTable.Enabled = False: fgElpTable.Clear: fgElpTable.Rows = 1
Call lstErr_Clear(lstErr, fgElpTable, "choisir une table 'click'")
lstId.Enabled = True  ': lstId.BackColor = vbWindowBackground

End Sub




'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub Msg_Rcv(X As String)
'---------------------------------------------------------

mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

MsgBox "Mise à jour de la table : " & DataBase_Open, vbInformation, Me.Caption

End Sub

Public Sub Msg_Snd(ByVal X As String)
End Sub

Private Sub Form_Resize()
If mWindowState <> Me.WindowState Then
    If Me.WindowState = 0 Or Me.WindowState = 2 Then
        Elp_Form_Resize Me, mWindowState, mHeight_0, mWidth_0, mHeight_2, mWidth_2
    End If
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
'2003.02.03 MDB_Local
End Sub

Private Sub fraElpTable_cmdOk_Click()
Dim V
On Error GoTo Error_Handler

fraElpTable_Control
If lstErr.ListCount > 0 Then Exit Sub
Select Case currentMethod
    Case constAddNew: V = adoElpTable_AddNew(rsMDB, recElpTable)
    Case constDelete: V = adoElpTable_Delete(rsMDB, recElpTable)
    Case constUpdate: V = adoElpTable_Update(rsMDB, recElpTable)
End Select
'V = adoElpTable_AddNew(rsMDB, recElpTable)
If Not IsNull(V) Then GoTo Error_MsgBox
fraElpTable_cmdQuit_Click
lstId_Click
Exit Sub
'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & "fraElpTable_cmdOk_Click :"

End Sub

Private Sub fraElpTable_cmdOk_GotFocus()
fraElpTable_Control
End Sub

Private Sub fraElpTable_cmdQuit_Click()
blnAddNew = False
fraElpTable.Visible = False: lstId.Enabled = True
Call lstErr_Clear(lstErr, fgElpTable, "choisir un enregistrement 'click'")
'fgElpTable.SetFocus
End Sub



Private Sub lstId_Click()
currentId = Trim(lstId.Text)
arrElpTable_Load currentId
fgElpTable_Display
'If fgElpTable.Rows > 1 Then fgElpTable.TopRow = fgElpTable.Rows - 1: fgElpTable.LeftCol = 0

End Sub

Private Sub mnuElpTableCopy_Click()
fraElpTable_DisplayItem
fraElpTable.Visible = True: lstId.Enabled = False
currentMethod = constAddNew
Call lstErr_Clear(lstErr, txtName, "Copier enregistrement")
fraElpTable_cmdOk.Visible = TableAut.Saisir
txtK1.Enabled = True
txtK2.Enabled = True
txtName.SetFocus

End Sub

Private Sub mnuElpTableModify_Click()
fraElpTable_DisplayItem
fraElpTable.Visible = True: lstId.Enabled = False
currentMethod = constUpdate
Call lstErr_Clear(lstErr, txtName, "Modification enregistrement")
fraElpTable_cmdOk.Visible = TableAut.Saisir
txtK1.Enabled = False
txtK2.Enabled = False
txtName.SetFocus

End Sub

Private Sub mnuElpTablePrint_Click()
cmdPrintX ""

End Sub


Private Sub txtK1_GotFocus()
Call txt_GotFocus(txtK1)

End Sub


Private Sub txtK1_LostFocus()
Call txt_LostFocus(txtK1)

End Sub


Private Sub txtK2_GotFocus()
Call txt_GotFocus(txtK2)

End Sub

Private Sub txtK2_LostFocus()
Call txt_LostFocus(txtK2)

End Sub


Private Sub txtName_GotFocus()
Call txt_GotFocus(txtName)
End Sub


Private Sub txtName_LostFocus()
Call txt_LostFocus(txtName)
End Sub

Private Sub txtMemo_GotFocus()
Call txt_GotFocus(txtMemo)
End Sub

Private Sub txtMemo_LostFocus()
Call txt_LostFocus(txtMemo)
End Sub


Private Sub mnuElpTableAddNew_Click()
blnAddNew = True
fraElpTable_Clear
fraElpTable.Visible = True: lstId.Enabled = False
currentMethod = constAddNew
Call lstErr_Clear(lstErr, txtK1, "Nouvel enregistrement")
rsElpTable_Init recElpTable
fraElpTable_cmdOk.Visible = TableAut.Saisir
txtK1.Enabled = True
txtK2.Enabled = True
txtK1.SetFocus
End Sub

Private Sub mnuElpTableDelete_Click()
fraElpTable_DisplayItem
fraElpTable.Visible = True
If Trim(txtName) = "" Then txtName = "?"
currentMethod = constDelete
Call lstErr_Clear(lstErr, txtName, "Suppression ligne")
X = MsgBox("Voulez-vous réellement supprimer cette ligne ?", vbYesNo + vbQuestion + vbDefaultButton2, "ancienne ligne")
If X = vbYes Then fraElpTable_cmdOk_Click  'fgElpTable_Delete
fraElpTable.Visible = False
fgElpTable.SetFocus
End Sub

Public Sub arrElpTable_Load(lId As String)
Dim X As String

arrElpTable_Nb = 0
ReDim arrElpTable(100)
X = "select * from ElpTable where SNN = 0" _
    & " and id = '" & lId & "'"
    
Set rsMDB = cnMDB.Execute(X)
Do While Not rsMDB.EOF
    If arrElpTable_Nb = UBound(arrElpTable) Then ReDim Preserve arrElpTable(arrElpTable_Nb + 1)
    arrElpTable_Nb = arrElpTable_Nb + 1
    Call rsElpTable_GetBuffer(rsMDB, arrElpTable(arrElpTable_Nb))
    rsMDB.MoveNext
Loop

End Sub

Public Sub fgElpTable_Display()
fgElpTable.Visible = False
fgElpTable_TopRow = fgElpTable.TopRow
fgElpTable.Rows = 1
fgElpTable.Clear
fgElpTable.FormatString = fgElpTable_FormatString
fgElpTable.Enabled = True
For arrElpTable_index = 1 To arrElpTable_Nb
    If arrElpTable(arrElpTable_index).Id <> constDelete Then
   ' And arrElpTable(arrElpTable_index).Method <> constIgnore Then
        fgElpTable.Rows = fgElpTable.Rows + 1
        fgElpTable.Row = fgElpTable.Rows - 1
        fgElpTable_DisplayItem
    End If
Next arrElpTable_index
If fgElpTable.Rows > 1 Then fgElpTable_Sort
If fgElpTable_TopRow < fgElpTable.Rows Then fgElpTable.TopRow = fgElpTable_TopRow
fgElpTable.Visible = True
End Sub

Public Sub fgElpTable_DisplayItem()
fgElpTable_K = (fgElpTable.Row) * fgElpTable.Cols
fgElpTable.TextArray(0 + fgElpTable_K) = Format$(arrElpTable(arrElpTable_index).K1, "@@@@@@@@@@@@")
fgElpTable.TextArray(1 + fgElpTable_K) = Format$(arrElpTable(arrElpTable_index).K2, "@@@@@@@@@@@@")
fgElpTable.TextArray(2 + fgElpTable_K) = arrElpTable(arrElpTable_index).Name
If Not IsNull(arrElpTable(arrElpTable_index).Memo) Then
    fgElpTable.TextArray(3 + fgElpTable_K) = arrElpTable(arrElpTable_index).Memo
Else
    fgElpTable.TextArray(3 + fgElpTable_K) = ""
End If

End Sub

Public Function arrElpTable_Scan(xId As String, xK1 As String, xK2 As String) As Integer
Dim mId As String * 12, mK1 As String * 12, mK2 As String * 12

mId = xId
mK1 = xK1
mK2 = xK2

For arrElpTable_Scan = 0 To arrElpTable_Nb
    If arrElpTable(arrElpTable_Scan).Id = mId Then
        If arrElpTable(arrElpTable_Scan).Id <> constIgnore Then
            If arrElpTable(arrElpTable_Scan).K1 = mK1 _
            And arrElpTable(arrElpTable_Scan).K2 = mK2 Then Exit Function
        End If
    End If
Next arrElpTable_Scan
arrElpTable_Scan = -1
End Function


Public Sub fraElpTable_Clear()
lstErr.Clear
usrColor_Set
txtK1 = ""
txtK2 = ""
txtName = ""
txtMemo = ""
lastActiveControl_Name = "txtMemo"
End Sub

Public Sub fraElpTable_Enabled(ByVal bln As Boolean)
fraElpTable.Enabled = bln
fraElpTable_cmdOk.Visible = bln
End Sub

Public Sub cmdContext_Quit()
If fraElpTable.Visible Then
    fraElpTable_cmdQuit_Click
Else
    If fgElpTable.Rows > 1 Then
        cmdClear
    Else
        If blnMsgBox_Quit Then
            X = MsgBox("Voulez-vous réellement abandonner?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
         Else
            X = vbYes
         End If
         If X = vbYes Then Unload Me
    End If
End If

End Sub

Public Sub cmdContext_Return()
If fraElpTable.Enabled Then
    If ActiveControl.Name = lastActiveControl_Name Then
        fraElpTable_cmdOk_Click
    Else
        SendKeys "{TAB}"
    End If
Else
    SendKeys "{TAB}"
End If

End Sub

Public Sub fraElpTable_DisplayItem()
usrColor_Set
txtK1 = Trim(recElpTable.K1)
txtK2 = Trim(recElpTable.K2)
txtName = Trim(recElpTable.Name)
If Not IsNull(recElpTable.Memo) Then
    txtMemo = Trim(recElpTable.Memo)
    If txtK1 = "PasswordX" Then txtMemo = ElpCipher_D(Trim(recElpTable.Memo), paramElpCypher)
Else
    txtMemo = ""
End If
fraElpTable_cmdOk.Visible = False 'TableAut.Saisir

End Sub

Public Sub fgElpTable_Sort()
fgElpTable.Row = 1
fgElpTable.RowSel = 1 'fgLRAttribut.Rows - 1

fgElpTable.Col = 0
fgElpTable.ColSel = 1
fgElpTable.Sort = flexSortStringAscending

End Sub

Public Sub txtMemo_Control()
recElpTable.Memo = RTrim(txtMemo)
End Sub

Public Sub fraElpTable_Control()
lstErr.Clear: arrTag_Set False
'recElpTable.Method = currentMethod
recElpTable.Id = currentId
txtK1_Control
txtK2_Control
txtName_Control
txtMemo_Control
If Trim(recElpTable.K1) = "PasswordX" Then recElpTable.Memo = ElpCipher_C(Trim(txtMemo), paramElpCypher)

If currentMethod = constAddNew Then
    If arrElpTable_Scan(recElpTable.Id, recElpTable.K1, recElpTable.K2) > 0 Then
        Call lstErr_AddItem(lstErr, txtName, "? Existe déjà")
    End If
End If
End Sub

Public Sub lstId_Scan(strdev As String)
Dim K As Integer
For K = 0 To lstId.ListCount - 1
    lstId.ListIndex = K
    If mId$(lstId.Text, 1, 3) = strdev Then Exit For
Next K
End Sub

Public Sub txtName_Control()
X = Trim(txtName)
If X = "" Then
    Call lstErr_AddItem(lstErr, txtName, "? Préciser l'intitulé")
    txtName.SetFocus
Else
    recElpTable.Name = X
    txtName = recElpTable.Name
End If

End Sub

Public Sub txtK1_Control()
recElpTable.K1 = Trim(txtK1)
txtK1 = recElpTable.K1
End Sub

Public Sub txtK2_Control()
recElpTable.K2 = Trim(txtK2)
txtK2 = recElpTable.K2
End Sub

Public Sub cmdPrintX(Msg As String)
prtElpTableX currentId
End Sub

