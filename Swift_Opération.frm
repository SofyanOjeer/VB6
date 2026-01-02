VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSwift_Opération 
   AutoRedraw      =   -1  'True
   Caption         =   "BIA_Swift"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13560
   Icon            =   "Swift_Opération.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9150
   ScaleWidth      =   13560
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7800
      TabIndex        =   3
      Top             =   0
      Width           =   5175
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8655
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   13530
      _ExtentX        =   23865
      _ExtentY        =   15266
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Suivi des Opérations"
      TabPicture(0)   =   "Swift_Opération.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "....."
      TabPicture(1)   =   "Swift_Opération.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdJPL"
      Tab(1).Control(1)=   "fgSAA"
      Tab(1).ControlCount=   2
      Begin VB.CommandButton cmdJPL 
         Caption         =   "JPL TEST Ne pas utiliser"
         Height          =   1095
         Left            =   -66600
         TabIndex        =   21
         Top             =   1440
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Frame fraTab0 
         Height          =   8205
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   13290
         Begin VB.Frame fraSelect_Options 
            Height          =   1005
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   11355
            Begin VB.ComboBox cboSelect_SQL 
               Height          =   315
               Left            =   120
               TabIndex        =   20
               Text            =   "cboSelect_SQL"
               Top             =   240
               Width           =   3615
            End
            Begin VB.ComboBox cboSelect_SWIOPESTA 
               Height          =   315
               Left            =   4680
               Sorted          =   -1  'True
               TabIndex        =   19
               Text            =   "STA"
               Top             =   600
               Width           =   1300
            End
            Begin VB.ComboBox cboSelect_SWISABCOP 
               Height          =   315
               Left            =   4680
               Sorted          =   -1  'True
               TabIndex        =   8
               Text            =   "OPE"
               Top             =   240
               Width           =   1300
            End
            Begin MSComCtl2.DTPicker txtSelect_SWIOPEFLUD 
               Height          =   300
               Left            =   8520
               TabIndex        =   9
               Top             =   240
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CustomFormat    =   "dd  MM yyy"
               Format          =   237436931
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin MSComCtl2.DTPicker txtSelect_SWIOPEFLUD_Max 
               Height          =   300
               Left            =   9840
               TabIndex        =   10
               Top             =   240
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CustomFormat    =   "dd  MM yyy"
               Format          =   237436931
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin MSComCtl2.DTPicker txtSelect_SWISABCPTD 
               Height          =   300
               Left            =   8520
               TabIndex        =   16
               Top             =   600
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CustomFormat    =   "dd  MM yyy"
               Format          =   237436931
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin MSComCtl2.DTPicker txtSelect_SWISABCPTD_Max 
               Height          =   300
               Left            =   9840
               TabIndex        =   17
               Top             =   600
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CustomFormat    =   "dd  MM yyy"
               Format          =   237436931
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label lblSelect_SWIOPESTA 
               Caption         =   "Statut"
               Height          =   375
               Left            =   3840
               TabIndex        =   18
               Top             =   600
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "Période de comptabilisation"
               Height          =   255
               Left            =   6240
               TabIndex        =   15
               Top             =   600
               Width           =   2055
            End
            Begin VB.Label txtSelect_YSWIOPEFLUD 
               Caption         =   "Période de réception"
               Height          =   255
               Left            =   6240
               TabIndex        =   14
               Top             =   240
               Width           =   1695
            End
            Begin VB.Label lblSelect_SWISABCOP 
               Caption         =   "Opération"
               Height          =   255
               Left            =   3840
               TabIndex        =   11
               Top             =   240
               Width           =   720
            End
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Exécuter la requête"
            Height          =   645
            Left            =   11880
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   240
            Width           =   1095
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   6825
            Left            =   120
            TabIndex        =   12
            Top             =   1200
            Width           =   12840
            _ExtentX        =   22648
            _ExtentY        =   12039
            _Version        =   393216
            Rows            =   1
            Cols            =   13
            FixedCols       =   0
            RowHeightMin    =   300
            BackColor       =   16777210
            ForeColor       =   8388608
            BackColorFixed  =   16776921
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   16777210
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLines       =   3
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"Swift_Opération.frx":047A
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
      Begin MSFlexGridLib.MSFlexGrid fgSAA 
         Height          =   6525
         Left            =   -75000
         TabIndex        =   13
         Top             =   480
         Width           =   6960
         _ExtentX        =   12277
         _ExtentY        =   11509
         _Version        =   393216
         Rows            =   1
         Cols            =   19
         FixedCols       =   0
         RowHeightMin    =   300
         BackColor       =   15269886
         ForeColor       =   8388608
         BackColorFixed  =   12648447
         ForeColorFixed  =   -2147483641
         BackColorSel    =   12648384
         BackColorBkg    =   15269886
         AllowBigSelection=   0   'False
         TextStyleFixed  =   4
         FocusRect       =   2
         HighLight       =   0
         GridLines       =   3
         GridLinesFixed  =   1
         AllowUserResizing=   3
         FormatString    =   $"Swift_Opération.frx":0520
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
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   13080
      Picture         =   "Swift_Opération.frx":0656
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   500
   End
   Begin VB.Label libSelect 
      BackColor       =   &H00FFFED9&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4905
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuContextAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuContextQuitter 
         Caption         =   "Quitter"
      End
      Begin VB.Menu mnux1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuYSWIOPE0_Import 
         Caption         =   "Importer les nouveaux messages SAA"
      End
      Begin VB.Menu mnuYSWIOPE0_Rapprochement 
         Caption         =   "Rapprocher Messages SAA / Opérations SAB"
      End
      Begin VB.Menu mnuYSWIOPE0_Comptabilisé 
         Caption         =   "Rechercher les opérations comptabilisées"
      End
   End
   Begin VB.Menu mnuPrint0 
      Caption         =   "mnuPrint0"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect_Print_Liste 
         Caption         =   "Imprimer liste"
      End
      Begin VB.Menu mnuSelect_Print_Détail 
         Caption         =   "Imprimer liste détaillée"
      End
   End
   Begin VB.Menu mnuSelect 
      Caption         =   "mnuSelect"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect_S999 
         Caption         =   "Annulation : S999"
      End
   End
End
Attribute VB_Name = "frmSwift_Opération"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit

Dim mHeight_0 As Integer, mWidth_0 As Integer, mHeight_2 As Integer, mWidth_2 As Integer, mWindowState As Integer
Dim arrTag() As Boolean, arrTagNb As Integer
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String, currentAction As String
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim blnMsgBox_Quit As Boolean, blnAddNew As Boolean, blnGlobalControl As Boolean, blnControl As Boolean
Dim X As String, I As Integer, Msg As String, valX As String, X1 As String, V As Variant, curX As Currency, dblX As Double
Dim intReturn As Integer
Dim Swift_Opération_Aut As typeAuthorization
Dim curX1 As Currency, curX2 As Currency
Dim blnAuto As Boolean

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim blnTransaction As Boolean

'______________________________________________________________________

Dim wAmjMin As String, wAmjMax As String, wHmsMin As Long, wHmsMax As Long
Dim xYSWIOPE0 As typeYSWIOPE0, newYSWIOPE0 As typeYSWIOPE0, oldYSWIOPE0 As typeYSWIOPE0
Dim arrYSWIOPE0() As typeYSWIOPE0, arrYSWIOPE0_Nb As Long, arrYSWIOPE0_Max As Long, arrYSWIOPE0_Index As Long

'______________________________________________________________________

Dim cnSIDE_DB As New ADODB.Connection, rsSIDE_DB As New ADODB.Recordset

Dim merAppe As typerAppe, xrAppe As typerAppe, xrAppe_E As typerAppe, xrAppe_R As typerAppe
Dim merIntv As typerIntv, xrIntv As typerIntv
Dim merInst As typerInst, xrInst As typerInst
Dim merJrnl As typerJrnl, xrJrnl As typerJrnl
Dim merMesg As typerMesg, xrMesg As typerMesg
Dim merTextField As typerTextField, xrTextField As typerTextField

Dim fgSAA_FormatString As String, fgSAA_K As Integer
Dim fgSAA_RowDisplay As Integer, fgSAA_RowClick As Integer, fgSAA_ColClick As Integer
Dim fgSAA_ColorClick As Long, fgSAA_ColorDisplay As Long
Dim fgSAA_Sort1 As Integer, fgSAA_Sort2 As Integer
Dim fgSAA_SortAD As Integer, fgSAA_Sort1_Old As Integer
Dim fgSAA_arrIndex As Integer
Dim blnfgSAA_DisplayLine As Boolean

Dim arrrAppe() As typerAppe, arrrAppe_Nb As Long, arrrAppe_Max As Long
Dim arrrAppe_E() As typerAppe
Dim arrrAppe_R() As typerAppe
Dim arrrInst() As typerInst, arrrInst_Nb As Long, arrrInst_Max As Long
Dim arrrIntv() As typerIntv, arrrIntv_Nb As Long, arrrIntv_Max As Long
Dim arrrJrnl() As typerJrnl, arrrJrnl_Nb As Long, arrrJrnl_Max As Long
Dim arrrMesg() As typerMesg, arrrMesg_Nb As Long, arrrMesg_Max As Long
Dim arrrTextField() As typerTextField, arrrTextField_Nb As Long, arrrTextField_Max As Long

Dim fgrTextField_FormatString As String, fgrTextField_K As Integer
Dim fgrTextField_RowDisplay As Integer, fgrTextField_RowClick As Integer, fgrTextField_ColClick As Integer
Dim fgrTextField_ColorClick As Long, fgrTextField_ColorDisplay As Long
Dim fgrTextField_Sort1 As Integer, fgrTextField_Sort2 As Integer
Dim fgrTextField_SortAD As Integer, fgrTextField_Sort1_Old As Integer
Dim fgrTextField_arrIndex As Integer
Dim blnfgrTextField_DisplayLine As Boolean



Dim meYSWIOPE0_Status As typeYSWIOPE0, oldYSWIOPE0_Status As typeYSWIOPE0

Dim mTitleText As String
Public Sub fgSAA_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgSAA.Row

If lRow > 0 And lRow < fgSAA.Rows Then
    fgSAA.Row = lRow
    For I = 0 To fgSAA_arrIndex
        fgSAA.Col = I: fgSAA.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgSAA.Row = mRow
    If fgSAA.Row > 0 Then
        lRow = fgSAA.Row
        lColor_Old = fgSAA.CellBackColor
        For I = 0 To fgSAA_arrIndex
          fgSAA.Col = I: fgSAA.CellBackColor = lColor
        Next I
        fgSAA.Col = 0
    End If
End If

End Sub

Private Sub fgSAA_Display()
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
fgSAA.Visible = False
fgSAA_Reset

fgSAA.Rows = 1
fgSAA.FormatString = fgSAA_FormatString
currentAction = "fgSAA_Display"
    
For I = 1 To arrrIntv_Nb
         
    xrIntv = arrrIntv(I)
    xrMesg = arrrMesg(I)
    xrInst = arrrInst(I)
    xrAppe = arrrAppe(I)
    xrAppe_E = arrrAppe_E(I)
    xrAppe_R = arrrAppe_R(I)
    
    fgSAA_DisplayLine (I)
Next I

fgSAA.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Opérations : " & fgSAA.Rows - 1): DoEvents
If fgSAA.Rows > 1 Then
    fgSAA_Sort1 = 0: fgSAA_Sort2 = 1: fgSAA_Sort
End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
fgSAA.Visible = True


End Sub

Public Sub fgSAA_DisplayLine(lIndex As Long)
Dim K As Integer, xDev As String, X As String, xCur As Currency

On Error Resume Next

End Sub



Public Sub fgSAA_Reset()
fgSAA.Clear
fgSAA_Sort1 = 0: fgSAA_Sort2 = 0
fgSAA_Sort1_Old = -1
fgSAA_RowDisplay = 0: fgSAA_RowClick = 0
fgSAA_arrIndex = fgSAA.Cols - 1
blnfgSAA_DisplayLine = False
End Sub

Public Sub fgSAA_Sort()
If fgSAA.Rows > 1 Then
    fgSAA.Row = 1
    fgSAA.RowSel = fgSAA.Rows - 1
    
    If fgSAA_Sort1_Old = fgSAA_Sort1 Then
        If fgSAA_SortAD = 5 Then
            fgSAA_SortAD = 6
        Else
            fgSAA_SortAD = 5
        End If
    Else
        fgSAA_SortAD = 5
    End If
    fgSAA_Sort1_Old = fgSAA_Sort1
    
    fgSAA.Col = fgSAA_Sort1
    fgSAA.ColSel = fgSAA_Sort2
    fgSAA.Sort = fgSAA_SortAD
End If
'cboDevise_Reset
End Sub


Public Sub fgSAA_SortX(lK As Integer)
Dim I As Integer, X As String
Dim xCur As Currency

For I = 1 To fgSAA.Rows - 1
    fgSAA.Row = I
    fgSAA.Col = lK
    Select Case lK
   '     Case 3: X = Format$(Val(fgSAA.Text), "000000000000000.00")
        Case 5:
            xCur = Val(fgSAA.Text)
            X = Format$(xCur, "000000000000000.00")
   '     Case 9: X = Format$(Val(fgSAA.Text), "000000000000000")
    End Select
    fgSAA.Col = fgSAA_arrIndex - 1
    fgSAA.Text = X
Next I


fgSAA_Sort1 = fgSAA_arrIndex - 1: fgSAA_Sort2 = fgSAA_arrIndex - 1
fgSAA_Sort
End Sub



Public Sub fgSelect_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgSelect.Row

If lRow > 0 And lRow < fgSelect.Rows Then
    fgSelect.Row = lRow
    For I = 0 To fgSelect_arrIndex
        fgSelect.Col = I: fgSelect.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgSelect.Row = mRow
    If fgSelect.Row > 0 Then
        lRow = fgSelect.Row
        lColor_Old = fgSelect.CellBackColor
        For I = 0 To fgSelect_arrIndex
          fgSelect.Col = I: fgSelect.CellBackColor = lColor
        Next I
        fgSelect.Col = 0
    End If
End If

End Sub
Private Sub fgSelect_Display()
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
currentAction = "fgselect_Display"
    
For I = 1 To arrYSWIOPE0_Nb
         
    xYSWIOPE0 = arrYSWIOPE0(I)
    If xYSWIOPE0.SWIOPEID > 0 Then
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine I
    End If
Next I

fgSelect.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Opérations : " & arrYSWIOPE0_Nb): DoEvents
'If fgSelect.Rows > 1 Then
'    fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
'End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Private Sub arrYSWIOPE0_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler
ReDim arrYSWIOPE0(101)
arrYSWIOPE0_Max = 100: arrYSWIOPE0_Nb = 0

Set rsSab = Nothing

xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWIOPE0 " & xWhere & " order by SWIOPEXBIC, SWIOPEFLUD"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = srvYSWIOPE0_GetBuffer_ODBC(rsSab, xYSWIOPE0)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmBIA_Swift.fgselect_Display"
        '' Exit Sub
     Else
         arrYSWIOPE0_Nb = arrYSWIOPE0_Nb + 1
         If arrYSWIOPE0_Nb > arrYSWIOPE0_Max Then
             arrYSWIOPE0_Max = arrYSWIOPE0_Max + 50
             ReDim Preserve arrYSWIOPE0(arrYSWIOPE0_Max)
         End If
         
         arrYSWIOPE0(arrYSWIOPE0_Nb) = xYSWIOPE0
    End If
    rsSab.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Private Sub arrrMesg_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler
ReDim arrrMesg(101)
arrrMesg_Max = 100: arrrMesg_Nb = 0

Set rsSIDE_DB = Nothing

xSql = "select * from rMesg " & xWhere & " order by mesg_crea_date_time"
Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)

Do While Not rsSIDE_DB.EOF
    V = srvrMesg_GetBuffer_ODBC(rsSIDE_DB, xrMesg)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "arrrMesg_SQL"
        '' Exit Sub
     Else
         arrrMesg_Nb = arrrMesg_Nb + 1
         If arrrMesg_Nb > arrrMesg_Max Then
             arrrMesg_Max = arrrMesg_Max + 50
             ReDim Preserve arrrMesg(arrrMesg_Max)
         End If
         
         arrrMesg(arrrMesg_Nb) = xrMesg
    End If
    rsSIDE_DB.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub fgSelect_DisplayLine(lIndex As Long)
On Error Resume Next
fgSelect.Col = 0: fgSelect.Text = xYSWIOPE0.SWIOPEXBIC
fgSelect.Col = 1: fgSelect.Text = dateImp10(xYSWIOPE0.SWIOPEFLUD) & " " & Format$(xYSWIOPE0.SWIOPEFLUH, "@@:@@:@@")
fgSelect.Col = 2: fgSelect.Text = dateImp10(xYSWIOPE0.SWISABCPTD)
fgSelect.Col = 3: fgSelect.Text = Format$(xYSWIOPE0.SWIOPEX32A, "### ### ### ###.00")
fgSelect.Col = 4: fgSelect.Text = xYSWIOPE0.SWIOPEX32D
fgSelect.Col = 5: fgSelect.Text = dateImp10(xYSWIOPE0.SWIOPEX32V)

fgSelect.Col = 6: fgSelect.Text = xYSWIOPE0.SWISABCOP & " " & xYSWIOPE0.SWISABDOS
fgSelect.Col = 7: fgSelect.Text = xYSWIOPE0.SWIOPEID
fgSelect.Col = 8: fgSelect.Text = xYSWIOPE0.SWIOPESTA
fgSelect.Col = 9: fgSelect.Text = dateImp10(xYSWIOPE0.SWIOPESTAD) & " " & Format$(xYSWIOPE0.SWIOPESTAH, "@@:@@:@@")
fgSelect.Col = 10: fgSelect.Text = xYSWIOPE0.SWIOPEXTRN
fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex

End Sub


Public Sub fgSelect_Reset()
fgSelect.Clear
fgSelect_Sort1 = 0: fgSelect_Sort2 = 0
fgSelect_Sort1_Old = -1
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0
fgSelect_arrIndex = fgSelect.Cols - 1
blnfgSelect_DisplayLine = False
fgSelect_SortAD = 6
fgSelect.LeftCol = 0

End Sub

Public Sub fgSelect_Sort()
If fgSelect.Rows > 1 Then
    fgSelect.Row = 1
    fgSelect.RowSel = fgSelect.Rows - 1
    
    If fgSelect_Sort1_Old = fgSelect_Sort1 Then
        If fgSelect_SortAD = 5 Then
            fgSelect_SortAD = 6
        Else
            fgSelect_SortAD = 5
        End If
    Else
        fgSelect_SortAD = 5
    End If
    fgSelect_Sort1_Old = fgSelect_Sort1
    
    fgSelect.Col = fgSelect_Sort1
    fgSelect.ColSel = fgSelect_Sort2
    fgSelect.Sort = fgSelect_SortAD
End If

End Sub
Public Sub fgSelect_SortX(lK As Integer)
Dim I As Integer, X As String
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    If lK = 2 Then
        fgSelect.Col = 2
        X = fgSelect.Text
    Else
        X = ""
    End If
    
    fgSelect.Col = 3
    X = X & Format$(Val(fgSelect.Text), "000000000000000.00")
    fgSelect.Col = fgSelect_arrIndex - 1
    fgSelect.Text = X
Next I


fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
End Sub

'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init(Mid$(Msg, 1, 12), Swift_Opération_Aut)
Form_Init


Select Case UCase$(Trim(Mid$(Msg, 1, 12)))
    Case "@AUTO_SWIOPE":
        Call lstErr_Clear(lstErr, cmdContext, "> @AUTO_SWIOPE........"): DoEvents
        blnAuto = True: mnuYSWIOPE0_Import_Click: mnuYSWIOPE0_Rapprochement_Click: mnuYSWIOPE0_Comptabilisé_Click
    Case Else: blnAuto = False
End Select


End Sub


Public Sub Form_Init()
Me.Enabled = False
Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents

If Not IsNull(param_Init) Then
    If Not blnAuto Then MsgBox "paramétrage inconsistant", vbCritical, "frmBIA_Swift.param_init"
    Unload Me
Else
    lstErr.Clear
End If
If Not IsNull(paramSAA_Init) Then
    If Not blnAuto Then MsgBox "paramétrage inconsistant", vbCritical, "frmBIA_Swift.paramSAA_Init"
    Unload Me
Else
    lstErr.Clear
End If


blnControl = False
fgSelect_FormatString = fgSelect.FormatString
fgSAA_FormatString = fgSAA.FormatString

fgSelect.Enabled = True
cmdReset

Me.Enabled = True
Me.MousePointer = 0
End Sub


Private Sub cboSelect_SQL_Click()
Dim X As String

txtSelect_SWISABCPTD.Visible = False
txtSelect_SWISABCPTD_Max.Visible = False
Dim wMonday As String, wFriday As String
DSYS_Init
X = dateElp("Ouvré", -3, DSys)
wMonday = dateElp("Weekday", -2, X)
wFriday = dateElp("Weekday", 6, wMonday)
If wFriday > YBIATAB0_DATE_CPT_J Then wFriday = YBIATAB0_DATE_CPT_J
Select Case Mid$(cboSelect_SQL, 1, 1)
    Case "1":   cboSelect_SWISABCOP = "CDE"
                cboSelect_SWIOPESTA = ""
                Call DTPicker_Set(txtSelect_SWIOPEFLUD, wMonday)
                Call DTPicker_Set(txtSelect_SWIOPEFLUD_Max, wFriday)
                txtSelect_SWISABCPTD = txtSelect_SWIOPEFLUD
                txtSelect_SWISABCPTD_Max = txtSelect_SWIOPEFLUD_Max
    Case "2":   cboSelect_SWISABCOP = "CDE"
                cboSelect_SWIOPESTA = "= 'E900'"
                Call DTPicker_Set(txtSelect_SWIOPEFLUD, "20040101")
                Call DTPicker_Set(txtSelect_SWIOPEFLUD_Max, dateElp("Weekday", -6, wMonday))
                Call DTPicker_Set(txtSelect_SWISABCPTD, wMonday)
                Call DTPicker_Set(txtSelect_SWISABCPTD_Max, wFriday)
                txtSelect_SWISABCPTD.Visible = True
                txtSelect_SWISABCPTD_Max.Visible = True

    Case "3":   cboSelect_SWISABCOP = "CDE"
                cboSelect_SWIOPESTA = "< 'E900'"
                Call DTPicker_Set(txtSelect_SWIOPEFLUD, "20040101")
                Call DTPicker_Set(txtSelect_SWIOPEFLUD_Max, dateElp("Weekday", -6, wMonday))
                Call DTPicker_Set(txtSelect_SWISABCPTD, "20040101")
                Call DTPicker_Set(txtSelect_SWISABCPTD_Max, wFriday)
End Select

End Sub


Private Sub cmdJPL_Click()
Dim paramODBC_DSN_SQL_Server_BIA As String
Dim cnSIDE_DB As New ADODB.Connection, rsSIDE_DB As New ADODB.Recordset
Dim xSql As String

'Migration SQL2010_BIA
'paramODBC_DSN_SQL_Server_BIA = "DSN=SQL_Server_BIA" & ";UID=SIDE_UPDATE" & "; PWD=Emi19lie"
paramODBC_DSN_SQL_Server_BIA = "DSN=SQL2010_BIA" & ";UID=SIDE_UPDATE" & "; PWD=Emi19lie"

cnSIDE_DB.Open paramODBC_DSN_SQL_Server_BIA

Set rsSIDE_DB = Nothing

xSql = "select * from SEPA"

Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)

Do While Not rsSIDE_DB.EOF
    Debug.Print rsSIDE_DB("EUPLABID")
    rsSIDE_DB.MoveNext

Loop
End Sub

Private Sub fgSAA_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Long
On Error Resume Next
If y <= fgSAA.RowHeightMin Then
    Select Case fgSAA.Col
        Case 0: fgSAA_Sort1 = 0: fgSAA_Sort2 = 1: fgSAA_Sort
        Case 1:  fgSAA_Sort1 = 1: fgSAA_Sort2 = 1: fgSAA_Sort
        Case 2: fgSAA_Sort1 = 2: fgSAA_Sort2 = 2: fgSAA_Sort
        Case 3: fgSAA_Sort1 = 3: fgSAA_Sort2 = 3: fgSAA_Sort
        Case 4: fgSAA_Sort1 = 4: fgSAA_Sort2 = 4: fgSAA_Sort
        Case 5: fgSAA_Sort1 = 5: fgSAA_Sort2 = 5: fgSAA_SortX 5
        Case 6: fgSAA_Sort1 = 6: fgSAA_Sort2 = 6: fgSAA_Sort
        Case 7: fgSAA_Sort1 = 7: fgSAA_Sort2 = 7: fgSAA_Sort
        Case 8: fgSAA_Sort1 = 8: fgSAA_Sort2 = 8: fgSAA_Sort
        Case 9: fgSAA_Sort1 = 9: fgSAA_Sort2 = 9: fgSAA_Sort
        Case 10: fgSAA_Sort1 = 10: fgSAA_Sort2 = 10: fgSAA_Sort
        Case 11:  fgSAA_Sort1 = 11: fgSAA_Sort2 = 11: fgSAA_Sort
        Case 12: fgSAA_Sort1 = 12: fgSAA_Sort2 = 12: fgSAA_Sort
        Case 13: fgSAA_Sort1 = 13: fgSAA_Sort2 = 13: fgSAA_Sort
        Case 14: fgSAA_Sort1 = 14: fgSAA_Sort2 = 14: fgSAA_Sort
        Case 15: fgSAA_Sort1 = 15: fgSAA_Sort2 = 15: fgSAA_Sort
   End Select
Else
    If fgSAA.Rows > 1 Then
        Call fgSAA_Color(fgSAA_RowClick, MouseMoveUsr.BackColor, fgSAA_ColorClick)
        fgSAA.Col = fgSAA_arrIndex:  K = CLng(fgSAA.Text)
        xrMesg = arrrMesg(K)
  
        'xrIntv = arrrIntv(K)
        'srvrIntv_ElpDisplay xrIntv

   End If
End If

End Sub



'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
blnControl = False
usrColor_Set
cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
currentAction = ""
fraSelect_Options.Enabled = True
'cmdSelect_Ok_Click

mnuSelect_S999.Enabled = Swift_Opération_Aut.Xspécial

txtSelect_SWIOPEFLUD.Enabled = Swift_Opération_Aut.Xspécial
mnuYSWIOPE0_Comptabilisé.Enabled = Swift_Opération_Aut.Xspécial
mnuYSWIOPE0_Import.Enabled = Swift_Opération_Aut.Xspécial
mnuYSWIOPE0_Rapprochement.Enabled = Swift_Opération_Aut.Xspécial

blnControl = True



End Sub


Public Function param_Init()

param_Init = Null
Call lstErr_Clear(lstErr, cmdContext, ". BIA_Swift_Import cbo"): DoEvents

fgSelect.Visible = False

Call DTPicker_Set(txtSelect_SWIOPEFLUD, DSys)
Call DTPicker_Set(txtSelect_SWIOPEFLUD_Max, DSys)


cboSelect_SQL.Clear
cboSelect_SQL.AddItem ""
cboSelect_SQL.AddItem "1 - CDE reçus (période)"
cboSelect_SQL.AddItem "2 - CDE antérieurs, comptabilisés(période)"
cboSelect_SQL.AddItem "3 - CDE antérieurs non comptabilisés"
cboSelect_SQL.ListIndex = 0

cboSelect_SWIOPESTA.Clear
cboSelect_SWIOPESTA.AddItem ""
cboSelect_SWIOPESTA.AddItem "= 'E200'"
cboSelect_SWIOPESTA.AddItem "= 'E300'"
cboSelect_SWIOPESTA.AddItem "= 'E900'"
cboSelect_SWIOPESTA.AddItem "< 'E900'"
cboSelect_SWIOPESTA.AddItem "= 'E999'"
cboSelect_SWIOPESTA.ListIndex = 0


cboSelect_SWISABCOP.Clear
cboSelect_SWISABCOP.AddItem ""
cboSelect_SWISABCOP.AddItem "CDE"
cboSelect_SWISABCOP.ListIndex = 0

Me.Enabled = True: Me.MousePointer = 0

End Function





'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
currentActiveControl_Name = C.Name
C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor
End Sub


'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------
arrTag(Val(C.Tag)) = True
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor
End Sub



Private Sub mnuSelect_S999_Click()
Dim xSql As String
Me.Enabled = False: Me.MousePointer = vbHourglass
meYSWIOPE0_Status = xYSWIOPE0
meYSWIOPE0_Status.SWIOPESTA = "E999"
meYSWIOPE0_Status.SWIOPESTAD = DSys
meYSWIOPE0_Status.SWIOPESTAH = time_Hms

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
blnTransaction_Set

V = sqlYSWIOPE0_Update(meYSWIOPE0_Status, xYSWIOPE0, cnsab)

If Not IsNull(V) Then
    xSql = "Rollback"
Else
    xSql = "Commit"
    xYSWIOPE0 = meYSWIOPE0_Status
End If

Set rsSab_Update = cnsab.Execute(xSql)
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
fgSelect_DisplayLine arrYSWIOPE0_Index
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdPrint_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Select Case SSTab1.Tab
    Case 0:
            If fgSelect.Rows > 1 Then cmdPrint_Ok
                
    Case 1:
End Select
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_SQL()
Dim V
Dim X As String
Dim xWhere As String, xAnd As String
Dim wAmj7 As Long
On Error GoTo Error_Handler

currentAction = "cmdYSWIOPE0_SQL"
Call DTPicker_Control(txtSelect_SWIOPEFLUD, wAmjMin)
Call DTPicker_Control(txtSelect_SWIOPEFLUD_Max, wAmjMax)
xWhere = " where SWIOPEFLUD >= " & wAmjMin & " and SWIOPEFLUD <= " & wAmjMax

X = Trim(cboSelect_SWISABCOP)
If X <> "" Then xWhere = xWhere & " and SWISABCOP = '" & X & "'"


X = Trim(cboSelect_SWIOPESTA)
If X <> "" Then xWhere = xWhere & " and SWIOPESTA " & X

    
If txtSelect_SWISABCPTD.Visible Then
    Call DTPicker_Control(txtSelect_SWISABCPTD, wAmjMin)
    Call DTPicker_Control(txtSelect_SWISABCPTD_Max, wAmjMax)
    xWhere = xWhere & " and SWISABCPTD >= " & wAmjMin & " and SWISABCPTD <= " & wAmjMax
End If

arrYSWIOPE0_SQL xWhere

fgSelect_Display

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_Ok_Click()
Dim blnOk As Boolean, Nb As Long

blnOk = fraSelect_Options.Enabled
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> BIA_Swift_cmdSelect_Ok ........"): DoEvents

fgSelect.Clear
If blnOk Then
    cmdSelect_Ok.Caption = "Options"
    cmdSelect_Ok.BackColor = &HFFFFFA   '&HC0FFFF
    fraSelect_Options.BackColor = &H8000000F
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Enabled = False
    cmdSelect_SQL
Else
    cmdSelect_Ok.Caption = constcmdRechercher
    cmdSelect_Ok.BackColor = &HC0FFC0
    fraSelect_Options.BackColor = &HFFFFFA    '&HC0FFFF
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Enabled = True
End If
Call lstErr_AddItem(lstErr, cmdContext, "< BIA_Swift_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0


End Sub


Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Long
On Error Resume Next
If y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 6: fgSelect_Sort1 = 6: fgSelect_Sort2 = 6: fgSelect_Sort
        Case 7: fgSelect_Sort1 = 7: fgSelect_Sort2 = 7: fgSelect_Sort
        Case 8: fgSelect_Sort1 = 8: fgSelect_Sort2 = 8: fgSelect_Sort
       Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        fgSelect.Col = fgSelect_arrIndex:  arrYSWIOPE0_Index = CLng(fgSelect.Text)
        fgSelect.LeftCol = 0
        xYSWIOPE0 = arrYSWIOPE0(arrYSWIOPE0_Index)
        
       If xYSWIOPE0.SAAAID = 0 Then Me.PopupMenu mnuSelect, vbPopupMenuLeftButton

   End If
End If
fgSelect.LeftCol = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'If blnControl Then
    cnSIDE_DB.Close
    Set cnSIDE_DB = Nothing

'End If
End Sub

Private Sub mnuContextAbandonner_Click()
cmdContext_Quit
End Sub


Private Sub mnuContextQuitter_Click()
Unload Me
End Sub

'---------------------------------------------------------
Private Sub Form_Activate()
'---------------------------------------------------------
Set XForm = Me
End Sub

'---------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

'---------------------------------------------------------
Select Case KeyCode
    Case Is = 13: KeyCode = 0: cmdContext_Return
    Case Is = 27: cmdContext_Quit
'   Case Is = 34: cmdPageNext_Click
'   Case Is = 33: cmdPagePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select


End Sub

Public Sub cmdContext_Quit()
blnControl = False
lstErr.Clear: lstErr.Height = 200
If SSTab1.Tab = 0 Then
        Unload Me
    Exit Sub
Else
    SSTab1.Tab = SSTab1.Tab - 1
End If

End Sub

Public Sub cmdContext_Return()
If SSTab1.Tab = 0 Then
    fgSelect.Row = fgSelect.TopRow
    fgSelect.Col = fgSelect_arrIndex: ' wK1 = fgSelect.Text
    'cmdSelect txtSelect ''fgSelect.Text

'    cmdSelect_Click
Else
    SendKeys "{TAB}"
End If
End Sub

'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
On Error GoTo Error_Handler

mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
blnControl = False


cnSIDE_DB.Open paramODBC_DSN_SIDE_DB
Exit Sub

Error_Handler:

blnControl = False
If Not blnAuto Then MsgBox Error
End Sub





Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Reset

End Sub

Private Sub Form_Resize()
If mWindowState <> Me.WindowState Then
    If Me.WindowState = 0 Or Me.WindowState = 2 Then
        Elp_Form_Resize Me, mWindowState, mHeight_0, mWidth_0, mHeight_2, mWidth_2
    End If
End If

End Sub

Public Sub MouseMoveActiveControl_Set(C As Control)
If MouseMoveActiveControl_Name <> C.Name Then
    MouseMoveActiveControl_Reset
    If Not C.Enabled Then
        MouseMoveActiveControl_Name = ""
    Else
        MouseMoveActiveControl_Name = C.Name
        If TypeOf C Is CommandButton Then
            MouseMoveActiveControl.BackColor = C.BackColor
            C.BackColor = MouseMoveUsr.BackColor
        Else
            If TypeOf C Is ListBox Then
                Elp_ResizeControl C
            Else
                MouseMoveActiveControl.ForeColor = C.ForeColor
                C.ForeColor = MouseMoveUsr.ForeColor
            End If
        End If
    End If
End If

End Sub


Public Sub MouseMoveActiveControl_Reset()
For Each xobj In Me.Controls
    If MouseMoveActiveControl_Name = xobj.Name Then
        MouseMoveActiveControl_Name = ""
        If TypeOf xobj Is CommandButton Then
            xobj.BackColor = MouseMoveActiveControl.BackColor
        Else
            If TypeOf xobj Is ListBox Then
                xobj.Height = 200
            Else
                xobj.ForeColor = MouseMoveActiveControl.ForeColor
            End If
        End If
        Exit For
    End If
Next xobj
End Sub


Public Sub txt_X()
'Call txt_GotFocus(txt)
'KeyAscii = convUCase(KeyAscii)
'Call txt_LostFocus(txt)

'Call txt_GotFocus(txt)
'If XopDevise(2).maxD = 0 Then
'    Call num_KeyAscii(KeyAscii)
'Else
'    Call num_KeyAsciiD(KeyAscii, txt)
'End If
'Call txt_LostFocus(txt)
'Call txt_GotFocus(txt)
'Call txt_LostFocus(txt)

End Sub


Private Sub mnuSelect_Print_Détail_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdPrint_Ok '"D "
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuSelect_Print_Liste_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdPrint_Ok '"L "
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuYSWIOPE0_Comptabilisé_Click()
Dim X As String
Dim K1 As Integer, K2 As Integer, lenText As Integer
Dim V, xSql As String

Me.Enabled = False: Me.MousePointer = vbHourglass

On Error GoTo Error_Handler
V = Null
Set rsSab = Nothing

Call lstErr_AddItem(lstErr, cmdContext, "> mnuYSWIOPE0_Comptabilisé_Click........"): DoEvents

    
    xSql = " where SWIOPESTA = 'E300'"
    
    arrYSWIOPE0_SQL xSql
    
    Call lstErr_AddItem(lstErr, cmdContext, "- mnuYSWIOPE0_Comptabilisé_Click : " & arrYSWIOPE0_Nb): DoEvents
    If arrYSWIOPE0_Nb > 0 Then

        mnuYSWIOPE0_Comptabilisé_Transaction
    
    End If

Error_Handler:
Exit_sub:
Call lstErr_AddItem(lstErr, cmdContext, "< mnuYSWIOPE0_Comptabilisé_Click"): DoEvents
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuYSWIOPE0_Import_Click()
Dim X As String
Dim K1 As Integer, K2 As Integer, lenText As Integer
Dim V, xSql As String

Me.Enabled = False: Me.MousePointer = vbHourglass

On Error GoTo Error_Handler
V = Null
Set rsSab = Nothing

Call lstErr_AddItem(lstErr, cmdContext, "> mnuYSWIOPE0_Import_Click........"): DoEvents

If IsNull(mnuStatus_Actualiser_YSWIOPE0_Fct("Select -2", wAmjMin, wHmsMin)) Then
    'Call DTPicker_Set(txtStatus_Amj, wAmjMin)
    'txtStatus_Hms = wHmsMin
        '2004.10.04  retrancher 1 minute ?
        
    wHmsMin = Time_Sss_Hms(Time_Hms_Sss(Format$(wHmsMin, "000000")) - 60)   'wHmsMin

    xSql = " where mesg_crea_date_time >= " & SQL_Date_Time(wAmjMin, wHmsMin) _
        & " and mesg_type = '700' and mesg_crea_rp_name = '_SI_from_SWIFT' and  x_own_lt = 'BIARFRPP'"
    arrrMesg_SQL xSql
        Call lstErr_AddItem(lstErr, cmdContext, "- mnuYSWIOPE0_Import_Click : " & arrrMesg_Nb): DoEvents

    If arrrMesg_Nb > 0 Then

      mnuYSWIOPE0_Import_Transaction
    
      'MAJ : tester si Erreur pendant MAJ ==> ? actualisation SWIOPEID = -2
      '----------------------------------------------------------------------
    '$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
      blnTransaction_Set
      X = arrrMesg(arrrMesg_Nb).mesg_crea_date_time & " "  ' il faut ajouter 1 espace pour timeHMS_Scan
      K1 = 0
      wAmjMax = DateJMA_Scan(X, K1)
      wHmsMax = CLng(TimeHMS_Scan(X, K1))
      Call mnuStatus_Actualiser_YSWIOPE0_Fct("Update -2", wAmjMax, wHmsMax)
    If Not IsNull(V) Then
    xSql = "Rollback"
Else
    xSql = "Commit"
End If

Set rsSab_Update = cnsab.Execute(xSql)
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
End If
End If

Error_Handler:
Exit_sub:
Call lstErr_AddItem(lstErr, cmdContext, "< mnuYSWIOPE0_Import_Click"): DoEvents
Me.Enabled = True: Me.MousePointer = 0

End Sub

Public Function mnuYSWIOPE0_Import_Transaction_Init()
Dim xSql As String, K1 As Integer, X As String
Dim V
Dim wSWIOPESTA As String
V = Null

xSql = "Select * from " & paramIBM_Library_SABSPE & ".YSWIOPE0" _
     & " where SAAAID = " & merMesg.Aid _
     & " and SAAUMIDL = " & merMesg.mesg_s_umidl _
     & " and SAAUMIDH = " & merMesg.mesg_s_umidh
Set rsSab = cnsab.Execute(xSql)
If Not rsSab.EOF Then
    V = "Message déjà enregistré dans YSWIOPE0.SAAUMIDL : " & xYSWIOPE0.SAAUMIDL
    GoTo Exit_Function
End If

xSql = "Select * from " & paramIBM_Library_SABSPE & ".YSWIOPE0" _
     & " where SWIOPEXTRN = '" & merMesg.mesg_trn_ref & "'"
Set rsSab = cnsab.Execute(xSql)
If Not rsSab.EOF Then
' il existe déjà un MT700 avec la même identification
    wSWIOPESTA = "E299"
  '  Mid$(merMesg.mesg_trn_ref, 16, 1) = "?"
    merMesg.mesg_trn_ref = "?" & merMesg.mesg_trn_ref
Else
    wSWIOPESTA = "E200"
End If

 V = sqlYSWIOPE0_Init(newYSWIOPE0, cnsab, rsSab)
 
 If IsNull(V) Then
        newYSWIOPE0.SWISABCPTD = 0
        newYSWIOPE0.SWISABCOP = "CDE"
        newYSWIOPE0.SWISABDOS = 0
        
        newYSWIOPE0.SAAAID = merMesg.Aid
        newYSWIOPE0.SAAUMIDL = merMesg.mesg_s_umidl
        newYSWIOPE0.SAAUMIDH = merMesg.mesg_s_umidh
        
        newYSWIOPE0.SWIOPESTA = wSWIOPESTA
        X = merMesg.mesg_crea_date_time
        K1 = 0
        wAmjMax = DateJMA_Scan(X, K1)
        wHmsMax = CLng(TimeHMS_Scan(X, K1))
        newYSWIOPE0.SWIOPEFLUD = wAmjMax
        newYSWIOPE0.SWIOPEFLUH = wHmsMax
        newYSWIOPE0.SWIOPEXMT = merMesg.mesg_type
        newYSWIOPE0.SWIOPEXBIC = merMesg.mesg_sender_X1
        newYSWIOPE0.SWIOPEXTRN = merMesg.mesg_trn_ref
        newYSWIOPE0.SWIOPEX32A = merMesg.x_fin_amount
        newYSWIOPE0.SWIOPEX32D = merMesg.x_fin_ccy
        newYSWIOPE0.SWIOPEX32V = merMesg.x_fin_value_date
    End If
Exit_Function:
    mnuYSWIOPE0_Import_Transaction_Init = V
End Function


Private Sub mnuYSWIOPE0_Rapprochement_Click()
Dim X As String
Dim K1 As Integer, K2 As Integer, lenText As Integer
Dim V, xSql As String

Me.Enabled = False: Me.MousePointer = vbHourglass

On Error GoTo Error_Handler
V = Null
Set rsSab = Nothing

Call lstErr_AddItem(lstErr, cmdContext, "> mnuYSWIOPE0_Rapprochement_Click........"): DoEvents

    
    xSql = " where SWIOPESTA = 'E200'"
    
    arrYSWIOPE0_SQL xSql
    
    Call lstErr_AddItem(lstErr, cmdContext, "- mnuYSWIOPE0_Rapprochement_Click : " & arrYSWIOPE0_Nb): DoEvents
    If arrYSWIOPE0_Nb > 0 Then

        mnuYSWIOPE0_Rapprochement_Transaction
    
    End If

Error_Handler:
Exit_sub:
Call lstErr_AddItem(lstErr, cmdContext, "< mnuYSWIOPE0_Rapprochement_Click"): DoEvents
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
If SSTab1.Tab = 0 Then cmdSelect_Ok.SetFocus

End Sub














Private Sub SSTab1_GotFocus()
Select Case SSTab1.Tab
    Case 0: fgSelect.LeftCol = 0
    Case 1: fgSAA.LeftCol = 0
End Select
End Sub


Public Sub cmdPrint_Ok()
Dim iRow As Integer, K As Integer, I As Integer
Dim blnOk As Boolean

fgSelect.Visible = False
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Impression Etat : " & fgSelect.Rows - 1)

mTitleText = Mid$(Trim(cboSelect_SQL), 1, 4) & " messages reçus [" & txtSelect_SWIOPEFLUD & " - " & txtSelect_SWIOPEFLUD_Max & " ]"
If txtSelect_SWISABCPTD.Visible Then mTitleText = mTitleText & ", comptabilisés [" & txtSelect_SWISABCPTD & " - " & txtSelect_SWISABCPTD_Max & " ]"

prtSWI_Opération_Open "Crédits Documentaires : " & mTitleText
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight

For iRow = 1 To fgSelect.Rows - 1
    
    fgSelect.Row = iRow
    fgSelect.Col = fgSelect_arrIndex:  K = CLng(fgSelect.Text)
    xYSWIOPE0 = arrYSWIOPE0(K)
    
    prtSWI_Opération_Line xYSWIOPE0, cnsab

Next iRow
prtSWI_Opération_Close
fgSelect.Visible = True
Me.Show
End Sub





Public Function mnuStatus_Actualiser_YSWIOPE0_Fct(lFct As String, lAMJ As String, lHMS As Long)
Static sYSWIOPE0_2 As typeYSWIOPE0
Dim V, xSql As String

V = Null
Select Case lFct
    Case "Select -2"
        xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWIOPE0" & " where  SWIOPEID =  -2"
        Set rsSab = Nothing
        Set rsSab = cnsab.Execute(xSql)
        V = srvYSWIOPE0_GetBuffer_ODBC(rsSab, sYSWIOPE0_2)
        
        If Not IsNull(V) Then GoTo Error_MsgBox
        lAMJ = sYSWIOPE0_2.SWIOPESTAD
        lHMS = sYSWIOPE0_2.SWIOPESTAH
        
     Case "Update -2"
        xYSWIOPE0 = sYSWIOPE0_2
        xYSWIOPE0.SWIOPESTAD = lAMJ
        xYSWIOPE0.SWIOPESTAH = lHMS
        V = sqlYSWIOPE0_Update(xYSWIOPE0, sYSWIOPE0_2, cnsab)
        If Not IsNull(V) Then GoTo Exit_Function

        
    Case Else: V = "non programmé " & lFct
End Select
GoTo Exit_Function

Error_Handler:
    V = Error
Error_MsgBox:
    ''If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : mnuStatus_Actualiser_YSWIOPE0_Fct"
        Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents

Exit_Function:
   mnuStatus_Actualiser_YSWIOPE0_Fct = V

End Function

Public Sub blnTransaction_Set()
If Not blnTransaction Then
    blnTransaction = True
    Set rsSab_Update = cnsab.Execute("SET TRANSACTION ISOLATION LEVEL READ COMMITTED")

End If

End Sub

Public Sub mnuYSWIOPE0_Import_Transaction()
Dim xSql As String, I As Integer
blnTransaction_Set
Call lstErr_AddItem(lstErr, cmdContext, "mnuYSWIOPE0_Import_Transaction" & arrrMesg_Nb): DoEvents

For I = 1 To arrrMesg_Nb
    Call lstErr_ChangeLastItem(lstErr, cmdContext, "mnuYSWIOPE0_Import_Transaction" & I & " / " & arrrMesg_Nb): DoEvents

    merMesg = arrrMesg(I)
    '$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    V = mnuYSWIOPE0_Import_Transaction_Init
    If IsNull(V) Then
        V = sqlYSWIOPE0_Insert(newYSWIOPE0, cnsab)

        If Not IsNull(V) Then
            xSql = "Rollback"
        Else
            xSql = "Commit"
        End If
        
        Set rsSab_Update = cnsab.Execute(xSql)
        '$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    End If
Next I
'=============================================================

End Sub
Public Sub mnuYSWIOPE0_Rapprochement_Transaction()
Dim xSql As String, I As Integer
Dim Nb As Integer

blnTransaction_Set
currentAction = "mnuYSWIOPE0_Rapprochement_Transaction"
Call lstErr_AddItem(lstErr, cmdContext, "mnuYSWIOPE0_Rapprochement_Transaction" & arrYSWIOPE0_Nb): DoEvents

For I = 1 To arrYSWIOPE0_Nb
    Call lstErr_ChangeLastItem(lstErr, cmdContext, "mnuYSWIOPE0_Rapprochement_Transaction" & I & " / " & arrYSWIOPE0_Nb): DoEvents

    newYSWIOPE0 = arrYSWIOPE0(I)
    
    xSql = "Select CDODOSCOP, CDODOSDOS from " & paramIBM_Library_SAB & ".ZCDODOS0 where CDODOSEXT = '" & newYSWIOPE0.SWIOPEXTRN & "'"
    Set rsSab = cnsab.Execute(xSql)
    Nb = 0
    Do While Not rsSab.EOF
        newYSWIOPE0.SWISABCOP = rsSab("CDODOSCOP")
        newYSWIOPE0.SWISABDOS = rsSab("CDODOSDOS")
        
        Nb = Nb + 1
        rsSab.MoveNext
    Loop
    
    'réponse unique : Maj dans YSWIMON0 du statut du message
    If Nb = 1 Then
    '$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
         newYSWIOPE0.SWIOPESTA = "E300"
         V = sqlYSWIOPE0_Update(newYSWIOPE0, arrYSWIOPE0(I), cnsab)

        If Not IsNull(V) Then
            xSql = "Rollback"
        Else
            xSql = "Commit"
        End If

        Set rsSab_Update = cnsab.Execute(xSql)
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    'Tester : Doublon, pas de réponse (=> rapprocher) ou  réponse unique ?
    '--------------------------------------------------------------------
    Else
        If Nb > 1 Then
            V = " ! doublon : " & newYSWIOPE0.SWIOPEXTRN
            If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
        End If
    End If
Next I
'=============================================================

End Sub

Public Sub mnuYSWIOPE0_Comptabilisé_Transaction()
Dim xSql As String, I As Integer
Dim Nb As Integer

blnTransaction_Set
currentAction = "mnuYSWIOPE0_Comptabilisé_Transaction"
Call lstErr_AddItem(lstErr, cmdContext, "mnuYSWIOPE0_Comptabilisé_Transaction" & arrYSWIOPE0_Nb): DoEvents

For I = 1 To arrYSWIOPE0_Nb
    Call lstErr_ChangeLastItem(lstErr, cmdContext, "mnuYSWIOPE0_Comptabilisé_Transaction" & I & " / " & arrYSWIOPE0_Nb): DoEvents

    newYSWIOPE0 = arrYSWIOPE0(I)
    
    xSql = "Select MOUREFDCO from " & paramIBM_Library_SAB _
        & ".ZMOUREF0 where MOUREFNUM = " & newYSWIOPE0.SWISABDOS _
        & " and MOUREFOPE = '" & newYSWIOPE0.SWISABCOP & "'" _
        & " and MOUREFEVE = 'OUV'"
        
    Set rsSab = cnsab.Execute(xSql)
    Nb = 0
    Do While Not rsSab.EOF
        If Nb = 0 Then newYSWIOPE0.SWISABCPTD = rsSab("MOUREFDCO") + 19000000
        
        Nb = Nb + 1
        rsSab.MoveNext
    Loop
    
    'réponse unique : Maj dans YSWIMON0 du statut du message
    '$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    If Nb > 0 Then
         newYSWIOPE0.SWIOPESTA = "E900"
         V = sqlYSWIOPE0_Update(newYSWIOPE0, arrYSWIOPE0(I), cnsab)

        If Not IsNull(V) Then
            xSql = "Rollback"
        Else
            xSql = "Commit"
        End If

        Set rsSab_Update = cnsab.Execute(xSql)
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    End If
Next I
'=============================================================

End Sub


