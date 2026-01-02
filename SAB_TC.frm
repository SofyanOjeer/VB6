VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSAB_TC 
   Caption         =   "SAB_TC"
   ClientHeight    =   9195
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   13875
   LinkTopic       =   "Form1"
   ScaleHeight     =   9195
   ScaleWidth      =   13875
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   13320
      Picture         =   "SAB_TC.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   500
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1200
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   8040
      TabIndex        =   0
      Top             =   0
      Width           =   5265
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8595
      Left            =   60
      TabIndex        =   1
      Top             =   525
      Width           =   13740
      _ExtentX        =   24236
      _ExtentY        =   15161
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Sélection"
      TabPicture(0)   =   "SAB_TC.frx":0102
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "ZBASFUT0"
      TabPicture(1)   =   "SAB_TC.frx":011E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraZBASFUT0"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "..."
      TabPicture(2)   =   "SAB_TC.frx":013A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame fraZBASFUT0 
         Height          =   8085
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   13170
         Begin VB.Frame fraYBASFUT0_Options 
            BackColor       =   &H00F0FFFF&
            BorderStyle     =   0  'None
            Height          =   450
            Left            =   120
            TabIndex        =   6
            Top             =   200
            Width           =   12855
            Begin VB.ComboBox cboYBASFUT0_Compte 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   2520
               Sorted          =   -1  'True
               TabIndex        =   14
               Text            =   "Compte"
               Top             =   60
               Width           =   2505
            End
            Begin VB.ComboBox cboYBASFUT0_Devise 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   840
               Sorted          =   -1  'True
               TabIndex        =   11
               Text            =   "Devise"
               Top             =   60
               Width           =   945
            End
            Begin VB.CommandButton cmdYBASFUT0_Options_Ok 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Rechercher"
               Height          =   400
               Left            =   11760
               Style           =   1  'Graphical
               TabIndex        =   9
               Top             =   0
               Width           =   1095
            End
            Begin MSComCtl2.DTPicker txtYBASFUT0_Amj_Max 
               Height          =   300
               Left            =   10080
               TabIndex        =   7
               Top             =   60
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
               Format          =   22872067
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin MSComCtl2.DTPicker txtYBASFUT0_AMJ_Min 
               Height          =   300
               Left            =   8520
               TabIndex        =   8
               Top             =   60
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
               Format          =   22872067
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label lblYBASFUT0_Compte 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Compte"
               Height          =   255
               Left            =   1920
               TabIndex        =   13
               Top             =   120
               Width           =   600
            End
            Begin VB.Label lblYBASFUT0_Devise 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Devise"
               Height          =   255
               Left            =   120
               TabIndex        =   10
               Top             =   120
               Width           =   600
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgYBASFUT0 
            Height          =   3945
            Left            =   120
            TabIndex        =   5
            Top             =   4080
            Width           =   12960
            _ExtentX        =   22860
            _ExtentY        =   6959
            _Version        =   393216
            Rows            =   1
            Cols            =   11
            FixedCols       =   0
            RowHeightMin    =   200
            BackColor       =   14737632
            ForeColor       =   4210688
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   14737632
            AllowBigSelection=   0   'False
            TextStyle       =   4
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"SAB_TC.frx":0156
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid fgYBASFUT0_Total 
            Height          =   3225
            Left            =   120
            TabIndex        =   12
            Top             =   720
            Width           =   12960
            _ExtentX        =   22860
            _ExtentY        =   5689
            _Version        =   393216
            Rows            =   1
            Cols            =   15
            FixedCols       =   0
            RowHeightMin    =   200
            BackColor       =   14737632
            ForeColor       =   4210688
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   14737632
            AllowBigSelection=   0   'False
            TextStyle       =   4
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"SAB_TC.frx":01FD
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
End
Attribute VB_Name = "frmSAB_TC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mHeight_0 As Integer, mWidth_0 As Integer, mHeight_2 As Integer, mWidth_2 As Integer, mWindowState As Integer
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String
Dim blnMsgBox_Quit As Boolean
Dim X As String, X1 As String, I As Integer
Dim Msg As String, valX As String
Dim currentMethod As String, lastMethod As String
Dim SAB_TC_Aut As typeAuthorization
Dim currentAction As String
Dim IdShell

Dim blnControl As Boolean, blnError As Boolean

Dim cnADO As New ADODB.Connection
Dim rsADO As New ADODB.Recordset
Dim mvtADO As New ADODB.Recordset

Dim blnODBC As Boolean
Dim xAmj_Min As String * 8, xAmj_Min_IBM As String * 7
Dim xAmj_Max As String * 8, xAmj_Max_IBM As String * 7

Dim meYBASFUT0 As typeYBASFUT0
Dim arrYBASFUT0() As typeYBASFUT0, arrYBASFUT0_Nb As Long, arrYBASFUT0_NbMax As Long

Dim fgYBASFUT0_FormatString As String, fgYBASFUT0_K As Integer
Dim fgYBASFUT0_RowDisplay As Integer, fgYBASFUT0_RowClick As Integer, fgYBASFUT0_ColClick As Integer
Dim fgYBASFUT0_ColorClick As Long, fgYBASFUT0_ColorDisplay As Long
Dim fgYBASFUT0_Sort1 As Integer, fgYBASFUT0_Sort2 As Integer
Dim fgYBASFUT0_SortAD As Integer, fgYBASFUT0_Sort1_Old As Integer
Dim fgYBASFUT0_arrIndex As Integer
Dim blnfgYBASFUT0_DisplayLine As Boolean


Dim arrYBASFUT0_Ech(5), arrYBASFUT0_Ech_IBM(5)
Dim arrYBASFUT0_Total() As typeYBASFUT0_Total, arrYBASFUT0_Total_Nb As Long, arrYBASFUT0_Total_NbMax As Long

Dim fgYBASFUT0_total_FormatString As String, fgYBASFUT0_total_K As Integer
Dim fgYBASFUT0_total_RowDisplay As Integer, fgYBASFUT0_total_RowClick As Integer, fgYBASFUT0_total_ColClick As Integer
Dim fgYBASFUT0_total_ColorClick As Long, fgYBASFUT0_total_ColorDisplay As Long
Dim fgYBASFUT0_total_Sort1 As Integer, fgYBASFUT0_total_Sort2 As Integer
Dim fgYBASFUT0_total_SortAD As Integer, fgYBASFUT0_total_Sort1_Old As Integer
Dim fgYBASFUT0_total_arrIndex As Integer
Dim blnfgYBASFUT0_total_DisplayLine As Boolean

Public Sub Form_Init()
Dim K As Integer
Me.Enabled = False
Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents

blnControl = False
SSTab1.Tab = 1

''MsgBox "date_cpt_j= 20030930"
''YBIATAB0_DATE_CPT_J = "20030930"

Call DTPicker_Set(txtYBASFUT0_AMJ_Min, YBIATAB0_DATE_CPT_AP1)
xAmj_Max = dateElp("Ouvré", 5, YBIATAB0_DATE_CPT_J)
Call DTPicker_Set(txtYBASFUT0_Amj_Max, xAmj_Max)

arrYBASFUT0_Ech(0) = YBIATAB0_DATE_CPT_J
arrYBASFUT0_Ech_IBM(0) = dateIBM(arrYBASFUT0_Ech(0))
fgYBASFUT0_total_FormatString = " |>SoldeV " & dateImp(arrYBASFUT0_Ech(0))

For K = 1 To 5
    arrYBASFUT0_Ech(K) = dateElp("Ouvré", 1, arrYBASFUT0_Ech(K - 1))
    arrYBASFUT0_Ech_IBM(K) = dateIBM(arrYBASFUT0_Ech(K))
    fgYBASFUT0_total_FormatString = fgYBASFUT0_total_FormatString & " |>     " & dateImp(arrYBASFUT0_Ech(K))
Next K

srvYBIATAB0_Import_cboDevise cboYBASFUT0_Devise

cboYBASFUT0_Compte.Clear
cboYBASFUT0_Compte.AddItem "N"

fgYBASFUT0_Nostro_Load
For K = 1 To arrYBASFUT0_Total_Nb
    cboYBASFUT0_Compte.AddItem arrYBASFUT0_Total(K).COMPTECOM
Next K

fgYBASFUT0_FormatString = fgYBASFUT0.FormatString
fgYBASFUT0.Enabled = True
fgYBASFUT0_total_FormatString = "<Devise|<Compte        " & fgYBASFUT0_total_FormatString  'fgYBASFUT0_Total.FormatString
fgYBASFUT0_Total.Enabled = True
cmdReset
Me.Enabled = True
Me.MousePointer = 0
End Sub



Public Sub fgYBASFUT0_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgYBASFUT0.Row

If lRow > 0 And lRow < fgYBASFUT0.Rows Then
    fgYBASFUT0.Row = lRow
    For I = 0 To fgYBASFUT0_arrIndex
        fgYBASFUT0.Col = I: fgYBASFUT0.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgYBASFUT0.Row = mRow
    If fgYBASFUT0.Row > 0 Then
        lRow = fgYBASFUT0.Row
        lColor_Old = fgYBASFUT0.CellBackColor
        For I = 0 To fgYBASFUT0_arrIndex
          fgYBASFUT0.Col = I: fgYBASFUT0.CellBackColor = lColor
        Next I
        fgYBASFUT0.Col = 0
    End If
End If

End Sub

Public Sub fgYBASFUT0_total_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgYBASFUT0_Total.Row

If lRow > 0 And lRow < fgYBASFUT0_Total.Rows Then
    fgYBASFUT0_Total.Row = lRow
    For I = 0 To fgYBASFUT0_total_arrIndex
        fgYBASFUT0_Total.Col = I: fgYBASFUT0_Total.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgYBASFUT0_Total.Row = mRow
    If fgYBASFUT0_Total.Row > 0 Then
        lRow = fgYBASFUT0_Total.Row
        lColor_Old = fgYBASFUT0_Total.CellBackColor
        For I = 0 To fgYBASFUT0_total_arrIndex
          fgYBASFUT0_Total.Col = I: fgYBASFUT0_Total.CellBackColor = lColor
        Next I
        fgYBASFUT0_Total.Col = 0
    End If
End If

End Sub


Private Sub fgYBASFUT0_Display()
Dim V
Dim Nb As Long
Dim xSQL As String
Dim xW As String

On Error GoTo Error_Handler

SSTab1.Tab = 1
fgYBASFUT0.Visible = False
fgYBASFUT0_Reset
fgYBASFUT0.Rows = 1
fgYBASFUT0.FormatString = fgYBASFUT0_FormatString

fgYBASFUT0_Total.Visible = False
fgYBASFUT0_total_Reset
fgYBASFUT0_Total.Rows = 1
fgYBASFUT0_Total.FormatString = fgYBASFUT0_total_FormatString

ReDim arrYBASFUT0(5000): arrYBASFUT0_Nb = 0: arrYBASFUT0_NbMax = 5000
ReDim arrYBASFUT0_Total(500): arrYBASFUT0_Total_Nb = 0: arrYBASFUT0_Total_NbMax = 500
recYBASFUT0_Init arrYBASFUT0(0) 'remise à blanc
recYBASFUT0_Total_Init arrYBASFUT0_Total(0)

fgYBASFUT0_Nostro_Load

Call DTPicker_Control(txtYBASFUT0_AMJ_Min, xAmj_Min): xAmj_Min_IBM = dateIBM(xAmj_Min)
Call DTPicker_Control(txtYBASFUT0_Amj_Max, xAmj_Max): xAmj_Max_IBM = dateIBM(xAmj_Max)


Set rsADO = Nothing

currentAction = "fgYBASFUT0_Display"

xSQL = "select * from ZBASFUT0 where BASFUTCPT LIKE " & Asc39 & "N%" & Asc39 & " AND BASFUTDVA >= " & xAmj_Min_IBM & " AND  BASFUTDVA <= " & xAmj_Max_IBM
X = Trim(cboYBASFUT0_Devise.Text)
If X <> "" Then xSQL = xSQL & " AND BASFUTDEV = " & Asc39 & X & Asc39
X = Trim(cboYBASFUT0_Compte)
If X <> "" Then xSQL = xSQL & " AND BASFUTCPT LIKE " & Asc39 & X & "%" & Asc39
Set rsADO = cnADO.Execute(xSQL)


Do While Not rsADO.EOF
        Call srvYBASFUT0_GetBuffer_ODBC(rsADO, meYBASFUT0)
        arrYBASFUT0_Nb = arrYBASFUT0_Nb + 1
        If arrYBASFUT0_Nb > arrYBASFUT0_NbMax Then
            arrYBASFUT0_NbMax = arrYBASFUT0_NbMax + 500
            ReDim Preserve arrYBASFUT0(arrYBASFUT0_NbMax)
        End If
        arrYBASFUT0(arrYBASFUT0_Nb) = meYBASFUT0
        
        fgYBASFUT0.Rows = fgYBASFUT0.Rows + 1
        
        fgYBASFUT0.Row = fgYBASFUT0.Rows - 1
        
        fgYBASFUT0.Col = 0: fgYBASFUT0.Text = meYBASFUT0.BASFUTDEV
        fgYBASFUT0.CellForeColor = vbBlue
        fgYBASFUT0.Col = 1: fgYBASFUT0.Text = meYBASFUT0.BASFUTCPT
        fgYBASFUT0.Col = 2: fgYBASFUT0.Text = meYBASFUT0.BASFUTOPE & " " & meYBASFUT0.BASFUTNAT & " " & Format$(meYBASFUT0.BASFUTDOS, "000000000")

        fgYBASFUT0.Col = 3: fgYBASFUT0.Text = Format$(Abs(meYBASFUT0.BASFUTMON), "### ### ### ##0.00")
        Select Case meYBASFUT0.BASFUTSEN
            Case "R": fgYBASFUT0.CellForeColor = vbRed: meYBASFUT0.BASFUTMON = meYBASFUT0.BASFUTMON * -1
            Case "L": fgYBASFUT0.CellForeColor = vbBlue
            Case "else: fgYBASFUT0.CellForeColor = vbmagenta"
        End Select
        fgYBASFUT0.Col = 5: fgYBASFUT0.Text = dateIBM10(meYBASFUT0.BASFUTDVA, False)
        fgYBASFUT0.Col = 4: fgYBASFUT0.Text = meYBASFUT0.BASFUTCLI
        
        fgYBASFUT0.Col = fgYBASFUT0_arrIndex: fgYBASFUT0.Text = arrYBASFUT0_Nb
        
        
        Call fgYBASFUT0_Total_BASFUTCPT
        
    rsADO.MoveNext
Loop

GoTo Exit_Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 1
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

Exit_Sub:


ReDim Preserve arrYBASFUT0(arrYBASFUT0_Nb + 1)
ReDim Preserve arrYBASFUT0_Total(arrYBASFUT0_Total_Nb + 1)

If fgYBASFUT0.Rows > 1 Then
    fgYBASFUT0_Sort1 = 0: fgYBASFUT0_Sort2 = 1: fgYBASFUT0_Sort
    fgYBASFUT0.TopRow = 1
End If
fgYBASFUT0.Visible = True

fgYBASFUT0_Total_Display

End Sub

Private Sub fgYBASFUT0_Nostro_Load()
Dim X As String
Dim V
Dim Nb As Long
Dim xSQL As String
Dim xW As String

On Error GoTo Error_Handler


currentAction = "fgYBASFUT0_Nostro_Load"

xSQL = "select COMPTECOM,COMPTEINT,COMPTEDEV from ZCOMPTE0 where COMPTELOR = " & Asc39 & "N" & Asc39 & " AND COMPTEFON <> " & Asc39 & "4" & Asc39
X = Trim(cboYBASFUT0_Devise.Text)
If X <> "" Then xSQL = xSQL & " AND COMPTEDEV = " & Asc39 & X & Asc39
X = Trim(cboYBASFUT0_Compte)
If X <> "" Then xSQL = xSQL & " AND COMPTECOM LIKE " & Asc39 & X & "%" & Asc39
Set rsADO = cnADO.Execute(xSQL)


ReDim arrYBASFUT0_Total(500): arrYBASFUT0_Total_Nb = 0: arrYBASFUT0_Total_NbMax = 500

Do While Not rsADO.EOF
        arrYBASFUT0_Total_Nb = arrYBASFUT0_Total_Nb + 1
        If arrYBASFUT0_Total_Nb > arrYBASFUT0_Total_NbMax Then
            arrYBASFUT0_Total_NbMax = arrYBASFUT0_Total_NbMax + 500
            ReDim Preserve arrYBASFUT0_Total(arrYBASFUT0_Total_NbMax)
        End If
        arrYBASFUT0_Total(arrYBASFUT0_Total_Nb) = arrYBASFUT0_Total(0)
        arrYBASFUT0_Total(arrYBASFUT0_Total_Nb).COMPTECOM = rsADO("COMPTECOM")
        arrYBASFUT0_Total(arrYBASFUT0_Total_Nb).COMPTEDEV = rsADO("COMPTEDEV")
        arrYBASFUT0_Total(arrYBASFUT0_Total_Nb).COMPTEINT = rsADO("COMPTEINT")
    rsADO.MoveNext
Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 1
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


Private Sub fgYBASFUT0_Total_Display()
Dim V
Dim Nb As Long, K As Long, K2 As Long
Dim xSQL As String
Dim xW As String

On Error GoTo Error_Handler

Set rsADO = Nothing
fgYBASFUT0_Total_Join

currentAction = "fgYBASFUT0_Total_Display"

For K = 1 To arrYBASFUT0_Total_Nb

        fgYBASFUT0_Total.Rows = fgYBASFUT0_Total.Rows + 1
        
        fgYBASFUT0_Total.Row = fgYBASFUT0_Total.Rows - 1
        
        fgYBASFUT0_Total.Col = 0: fgYBASFUT0_Total.Text = arrYBASFUT0_Total(K).COMPTEDEV
        fgYBASFUT0_Total.CellForeColor = vbBlue
        fgYBASFUT0_Total.Col = 1: fgYBASFUT0_Total.Text = arrYBASFUT0_Total(K).COMPTECOM
        
        For K2 = 0 To 6
            fgYBASFUT0_Total.Col = 2 + K2: fgYBASFUT0_Total.Text = Format$(Abs(arrYBASFUT0_Total(K).BASFUTMON(K2)), "### ### ### ##0.00")
            If arrYBASFUT0_Total(K).BASFUTMON(K2) > 0 Then
                fgYBASFUT0_Total.CellForeColor = vbRed
            Else
                fgYBASFUT0_Total.CellForeColor = vbBlue
            End If
        Next K2
          
        fgYBASFUT0_Total.Col = fgYBASFUT0_total_arrIndex: fgYBASFUT0_Total.Text = arrYBASFUT0_Total_Nb
        
        
Next K

GoTo Exit_Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 1
    MsgBox V, vbCritical, Me.Name & " : " & currentAction
    
Exit_Sub:
If fgYBASFUT0_Total.Rows = 1 Then
    fgYBASFUT0_total_Sort1 = 0: fgYBASFUT0_total_Sort2 = 1: fgYBASFUT0_total_Sort
    If fgYBASFUT0_Total.Rows = 1 Then fgYBASFUT0_Total.TopRow = 1
End If
fgYBASFUT0_Total.Visible = True

End Sub



Public Sub fgYBASFUT0_Reset()
fgYBASFUT0.Clear
fgYBASFUT0_Sort1 = 0: fgYBASFUT0_Sort2 = 0
fgYBASFUT0_Sort1_Old = -1
fgYBASFUT0_RowDisplay = 0: fgYBASFUT0_RowClick = 0
fgYBASFUT0_arrIndex = fgYBASFUT0.Cols - 1
blnfgYBASFUT0_DisplayLine = False
fgYBASFUT0_SortAD = 6
fgYBASFUT0.LeftCol = 0

End Sub

Public Sub fgYBASFUT0_total_Reset()
fgYBASFUT0_Total.Clear
fgYBASFUT0_total_Sort1 = 0: fgYBASFUT0_total_Sort2 = 0
fgYBASFUT0_total_Sort1_Old = -1
fgYBASFUT0_total_RowDisplay = 0: fgYBASFUT0_total_RowClick = 0
fgYBASFUT0_total_arrIndex = fgYBASFUT0_Total.Cols - 1
blnfgYBASFUT0_total_DisplayLine = False
fgYBASFUT0_total_SortAD = 6
fgYBASFUT0_Total.LeftCol = 0

End Sub


Public Sub fgYBASFUT0_Sort()
If fgYBASFUT0.Rows > 1 Then
    fgYBASFUT0.Row = 1
    fgYBASFUT0.RowSel = fgYBASFUT0.Rows - 1
    
    If fgYBASFUT0_Sort1_Old = fgYBASFUT0_Sort1 Then
        If fgYBASFUT0_SortAD = 5 Then
            fgYBASFUT0_SortAD = 6
        Else
            fgYBASFUT0_SortAD = 5
        End If
    Else
        fgYBASFUT0_SortAD = 5
    End If
    fgYBASFUT0_Sort1_Old = fgYBASFUT0_Sort1
    
    fgYBASFUT0.Col = fgYBASFUT0_Sort1
    fgYBASFUT0.ColSel = fgYBASFUT0_Sort2
    fgYBASFUT0.Sort = fgYBASFUT0_SortAD
End If

End Sub

Public Sub fgYBASFUT0_total_Sort()
If fgYBASFUT0_Total.Rows > 1 Then
    fgYBASFUT0_Total.Row = 1
    fgYBASFUT0_Total.RowSel = fgYBASFUT0_Total.Rows - 1
    
    If fgYBASFUT0_total_Sort1_Old = fgYBASFUT0_total_Sort1 Then
        If fgYBASFUT0_total_SortAD = 5 Then
            fgYBASFUT0_total_SortAD = 6
        Else
            fgYBASFUT0_total_SortAD = 5
        End If
    Else
        fgYBASFUT0_total_SortAD = 5
    End If
    fgYBASFUT0_total_Sort1_Old = fgYBASFUT0_total_Sort1
    
    fgYBASFUT0_Total.Col = fgYBASFUT0_total_Sort1
    fgYBASFUT0_Total.ColSel = fgYBASFUT0_total_Sort2
    fgYBASFUT0_Total.Sort = fgYBASFUT0_total_SortAD
End If

End Sub


Public Sub fgYBASFUT0_SortX(lK As Integer)
Dim I As Integer, X As String
For I = 1 To fgYBASFUT0.Rows - 1
    fgYBASFUT0.Row = I
    fgYBASFUT0.Col = lK
    X = Format$(Val(fgYBASFUT0.Text), "000000000000000.00")
    fgYBASFUT0.Col = fgYBASFUT0_arrIndex - 1
    fgYBASFUT0.Text = X
Next I


fgYBASFUT0_Sort1 = fgYBASFUT0_arrIndex - 1: fgYBASFUT0_Sort2 = fgYBASFUT0_arrIndex - 1
fgYBASFUT0_Sort
End Sub
Public Sub fgYBASFUT0_total_SortX(lK As Integer)
Dim I As Integer, X As String
For I = 1 To fgYBASFUT0_Total.Rows - 1
    fgYBASFUT0_Total.Row = I
    fgYBASFUT0_Total.Col = lK
    X = Format$(Val(fgYBASFUT0_Total.Text), "000000000000000.00")
    fgYBASFUT0_Total.Col = fgYBASFUT0_total_arrIndex - 1
    fgYBASFUT0_Total.Text = X
Next I


fgYBASFUT0_total_Sort1 = fgYBASFUT0_total_arrIndex - 1: fgYBASFUT0_total_Sort2 = fgYBASFUT0_total_arrIndex - 1
fgYBASFUT0_total_Sort
End Sub

Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
Dim X As String
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate
Call BiaPgmAut_Init(mId$(Msg, 1, 12), SAB_TC_Aut)    '
SSTab1.Tab = 0
Form_Init
End Sub


Private Sub cboYBASFUT0_Devise_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub cmdYBASFUT0_Options_Ok_Click()
Me.Enabled = False: MousePointer = vbHourglass
fgYBASFUT0_Display
Me.Enabled = True: MousePointer = 0
End Sub

Private Sub Form_Resize()
If mWindowState <> Me.WindowState Then
    If Me.WindowState = 0 Or Me.WindowState = 2 Then
        Elp_Form_Resize Me, mWindowState, mHeight_0, mWidth_0, mHeight_2, mWidth_2
    End If
End If

End Sub

Public Sub cmdContext_Quit()
    If blnMsgBox_Quit Then
       X = MsgBox("Voulez-vous réellement abandonner?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
    Else
       X = vbYes
    End If
    If X = vbYes Then Unload Me

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


Private Sub cmdContext_Click()
cmdContext_Quit

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case Is = 13: KeyCode = 0:  cmdContext_Return
    Case Is = 27:  cmdContext_Quit
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select

End Sub

Private Sub Form_Load()
mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)

''fraMT950.Enabled = False
cnADO.Open paramODBC_DSN_SAB

End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub

Private Sub Form_Unload(Cancel As Integer)
cnADO.Close
Set cnADO = Nothing

End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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




Public Sub cmdReset()
Dim X As String
On Error Resume Next
Me.Enabled = False: Me.MousePointer = vbHourglass

blnControl = False
blnError = False
usrColor_Set
cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
currentAction = ""

'fgYBASFUT0_Display

blnControl = True
Me.Enabled = True: Me.MousePointer = 0

End Sub
Private Sub fgYBASFUT0_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim wK1 As Long
On Error Resume Next
If Y <= fgYBASFUT0.RowHeightMin Then
    Select Case fgYBASFUT0.Col
        Case 0: fgYBASFUT0_Sort1 = 0: fgYBASFUT0_Sort2 = 2: fgYBASFUT0_Sort
        Case 1:  fgYBASFUT0_Sort1 = 1: fgYBASFUT0_Sort2 = 2: fgYBASFUT0_Sort
        Case 2: fgYBASFUT0_Sort1 = 2: fgYBASFUT0_Sort2 = 2: fgYBASFUT0_Sort
        Case 3: fgYBASFUT0_Sort1 = 3: fgYBASFUT0_Sort2 = 3: fgYBASFUT0_SortX 3
        Case 4: fgYBASFUT0_Sort1 = 4: fgYBASFUT0_Sort2 = 4: fgYBASFUT0_Sort
        Case 5: fgYBASFUT0_Sort1 = 5: fgYBASFUT0_Sort2 = 5: fgYBASFUT0_Sort
       Case fgYBASFUT0_arrIndex:  fgYBASFUT0_SortX fgYBASFUT0_arrIndex
    End Select
Else
    If fgYBASFUT0.Rows > 1 Then
        Call fgYBASFUT0_Color(fgYBASFUT0_RowClick, MouseMoveUsr.BackColor, fgYBASFUT0_ColorClick)
        fgYBASFUT0.Col = fgYBASFUT0_arrIndex: wK1 = Val(fgYBASFUT0.Text)
        meYBASFUT0 = arrYBASFUT0(wK1): srvYBASFUT0_ElpDisplay meYBASFUT0
   End If
End If
End Sub

'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub



Public Sub fgYBASFUT0_Total_BASFUTCPT()

Dim K As Long, K2 As Long
Dim blnOk As Boolean

'Total par 1 - dev,compte

blnOk = False
For K = 1 To arrYBASFUT0_Total_Nb
    If arrYBASFUT0_Total(K).COMPTEDEV = meYBASFUT0.BASFUTDEV _
    And arrYBASFUT0_Total(K).COMPTECOM = meYBASFUT0.BASFUTCPT Then
        blnOk = True
        Exit For
    End If
Next K

If Not blnOk Then
    arrYBASFUT0_Total_Nb = arrYBASFUT0_Total_Nb + 1
    If arrYBASFUT0_Total_Nb > arrYBASFUT0_Total_NbMax Then
        arrYBASFUT0_Total_NbMax = arrYBASFUT0_Total_NbMax + 500
        ReDim Preserve arrYBASFUT0_Total(arrYBASFUT0_Total_NbMax)
    End If
    
    arrYBASFUT0_Total(arrYBASFUT0_Nb) = arrYBASFUT0_Total(0)
    K = arrYBASFUT0_Total_Nb
    arrYBASFUT0_Total(arrYBASFUT0_Total_Nb).COMPTEDEV = meYBASFUT0.BASFUTDEV
    arrYBASFUT0_Total(arrYBASFUT0_Total_Nb).COMPTECOM = meYBASFUT0.BASFUTCPT
    arrYBASFUT0_Total(arrYBASFUT0_Total_Nb).BASFUTMON(0) = 0
End If
    
    blnOk = False
    For K2 = 1 To 5
        If meYBASFUT0.BASFUTDVA <= arrYBASFUT0_Ech_IBM(K2) Then blnOk = True: Exit For
    Next K2
    If Not blnOk Then
        K2 = 6
        If meYBASFUT0.BASFUTDVA < arrYBASFUT0_Ech(0) Then arrYBASFUT0_Total(K).BASFUTDVA_Err = True
    End If
    arrYBASFUT0_Total(K).BASFUTMON(K2) = arrYBASFUT0_Total(K).BASFUTMON(K2) + meYBASFUT0.BASFUTMON


End Sub

Public Sub fgYBASFUT0_Total_Join()
Dim V
Dim K As Long, K2 As Long
Dim xSQL As String
Dim curX As Currency, xAmj_IBM As String
Dim KM As Long

On Error GoTo Error_Handler

Set rsADO = Nothing

currentAction = "fgYBASFUT0_Total_Join"
'
' Le solde en valeur inclut les mouvements dont la date valeur est jour
For K = 1 To arrYBASFUT0_Total_Nb
    xSQL = "select * from ZSOLDE0 where SOLDECOM = " & Asc39 & arrYBASFUT0_Total(K).COMPTECOM & Asc39
     Set rsADO = cnADO.Execute(xSQL)

     If Not rsADO.EOF Then
         arrYBASFUT0_Total(K).SOLDEVEN = rsADO("SOLDEVEN")
         For KM = 0 To 6
            arrYBASFUT0_Total(K).BASFUTMON(KM) = arrYBASFUT0_Total(K).BASFUTMON(KM) + arrYBASFUT0_Total(K).SOLDEVEN
        Next KM
    End If
    
       xSQL = "select MOUVEMDVA,MOUVEMMON from ZMOUVEA0 where MOUVEMCOM = " & Asc39 & arrYBASFUT0_Total(K).COMPTECOM & Asc39 _
               & " AND MOUVEMDVA > " & arrYBASFUT0_Ech_IBM(0)
       Set mvtADO = Nothing
       Set mvtADO = cnADO.Execute(xSQL)
       Do While Not mvtADO.EOF
            xAmj_IBM = mvtADO("MOUVEMDVA")
            curX = mvtADO("MOUVEMMON")
           KM = 6
           For K2 = 1 To 5
               If xAmj_IBM <= arrYBASFUT0_Ech_IBM(K2) Then KM = K2: Exit For
           Next K2
           If KM = 1 Then
                arrYBASFUT0_Total(K).BASFUTMON(0) = arrYBASFUT0_Total(K).BASFUTMON(0) - curX
           Else
                arrYBASFUT0_Total(K).BASFUTMON(KM) = arrYBASFUT0_Total(K).BASFUTMON(KM) + curX
           End If
          mvtADO.MoveNext
       Loop
Next K

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 1
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
