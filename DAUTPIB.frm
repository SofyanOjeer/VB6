VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmDAUTPIB 
   AutoRedraw      =   -1  'True
   Caption         =   "BIA_DWH"
   ClientHeight    =   9144
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   13560
   Icon            =   "DAUTPIB.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9144
   ScaleWidth      =   13560
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   432
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
      _ExtentX        =   23855
      _ExtentY        =   15261
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Gestion sur les mois traités"
      TabPicture(0)   =   "DAUTPIB.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "....."
      TabPicture(1)   =   "DAUTPIB.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.Frame fraTab0 
         Height          =   8205
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   13290
         Begin VB.CommandButton cmdSelect_Ajouter 
            BackColor       =   &H00FFC0FF&
            Caption         =   "Ajouter"
            Enabled         =   0   'False
            Height          =   285
            Left            =   11640
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   720
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Frame fraDAUTPIB 
            BackColor       =   &H80000013&
            Caption         =   "Gestion du fichier des autorisations Prêts/Emprunts BANQUES  - DAUTPIB -"
            Height          =   5775
            Left            =   6840
            TabIndex        =   11
            Top             =   1680
            Visible         =   0   'False
            Width           =   6015
            Begin VB.CheckBox chkfg_DAUTECH 
               BackColor       =   &H80000013&
               Height          =   255
               Left            =   2040
               TabIndex        =   31
               Top             =   3480
               Width           =   255
            End
            Begin VB.ComboBox cbo_CDEV 
               Height          =   315
               Left            =   2040
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   18
               Top             =   2520
               Width           =   2055
            End
            Begin VB.CommandButton cmdfra_Creer 
               BackColor       =   &H00FFC0FF&
               Caption         =   "Créer"
               Height          =   495
               Left            =   5040
               Style           =   1  'Graphical
               TabIndex        =   29
               Top             =   4920
               Width           =   855
            End
            Begin VB.CommandButton cmdfra_Modifier 
               BackColor       =   &H00FF80FF&
               Caption         =   "Modifier"
               Height          =   495
               Left            =   3960
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   4920
               Width           =   855
            End
            Begin VB.CommandButton cmdfra_Supprimer 
               BackColor       =   &H000000FF&
               Caption         =   "Supprimer"
               Height          =   495
               Left            =   2760
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   4920
               Width           =   855
            End
            Begin VB.TextBox txtfg_DAUTMON 
               Height          =   300
               Left            =   2040
               TabIndex        =   19
               Text            =   "Montant autorisation"
               Top             =   3000
               Width           =   2055
            End
            Begin VB.TextBox txtfg_DAUTAUT 
               Height          =   300
               Left            =   2040
               TabIndex        =   17
               Text            =   "Code autorisation"
               Top             =   2040
               Width           =   975
            End
            Begin VB.TextBox txtfg_DAUTCLI 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2040
               TabIndex        =   16
               Text            =   "No matricule"
               Top             =   1560
               Width           =   975
            End
            Begin VB.TextBox txtfg_DAUTETB 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2040
               TabIndex        =   15
               Text            =   "Etablissement"
               Top             =   1080
               Width           =   375
            End
            Begin VB.TextBox txtfg_DAUTVER 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2640
               TabIndex        =   14
               Text            =   "Version"
               Top             =   600
               Width           =   375
            End
            Begin VB.TextBox txtfg_DAUTSTA 
               Height          =   300
               Left            =   2040
               TabIndex        =   13
               Text            =   "Statut"
               Top             =   600
               Width           =   375
            End
            Begin MSComCtl2.DTPicker txtfg_DAUTECH 
               Height          =   300
               Left            =   2760
               TabIndex        =   20
               Top             =   3480
               Width           =   1335
               _ExtentX        =   2350
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CustomFormat    =   "dd  MM yyy"
               Format          =   10223619
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label lblfg_DAUTECH 
               BackColor       =   &H80000013&
               Caption         =   "Date échéance"
               Height          =   300
               Left            =   240
               TabIndex        =   26
               Top             =   3480
               Width           =   1455
            End
            Begin VB.Label lblfg_DAUTMON 
               BackColor       =   &H80000013&
               Caption         =   "Montant autorisation"
               Height          =   375
               Left            =   240
               TabIndex        =   25
               Top             =   3000
               Width           =   1575
            End
            Begin VB.Label lblfg_DAUTAUT 
               BackColor       =   &H80000013&
               Caption         =   "Code autorisation"
               Height          =   300
               Left            =   240
               TabIndex        =   24
               Top             =   2040
               Width           =   1455
            End
            Begin VB.Label lblfg_DAUTDEV 
               BackColor       =   &H80000013&
               Caption         =   "Code devise"
               Height          =   300
               Left            =   240
               TabIndex        =   23
               Top             =   2520
               Width           =   1455
            End
            Begin VB.Label lblfg_DAUTCLI 
               BackColor       =   &H80000013&
               Caption         =   "Matricule client"
               Height          =   300
               Left            =   240
               TabIndex        =   22
               Top             =   1560
               Width           =   1455
            End
            Begin VB.Label lblfg_DAUTETB 
               BackColor       =   &H80000013&
               Caption         =   "Code établissement"
               Height          =   300
               Left            =   240
               TabIndex        =   21
               Top             =   1080
               Width           =   1575
            End
            Begin VB.Label lblfg_DAUTSTA_VER 
               BackColor       =   &H80000013&
               Caption         =   "Statut / Version"
               Height          =   300
               Left            =   240
               TabIndex        =   12
               Top             =   600
               Width           =   1575
            End
         End
         Begin VB.Frame fraSelect_Options 
            Height          =   1005
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   11355
            Begin MSComCtl2.DTPicker txtSelect_DAUTPER 
               Height          =   300
               Left            =   1920
               TabIndex        =   8
               Top             =   240
               Width           =   1455
               _ExtentX        =   2561
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CustomFormat    =   "dd  MM yyy"
               Format          =   10223619
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label lblSelect_DRCHPER 
               Caption         =   "Période de traitement"
               Height          =   255
               Left            =   120
               TabIndex        =   10
               Top             =   240
               Width           =   1695
            End
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Rechercher"
            Enabled         =   0   'False
            Height          =   315
            Left            =   11640
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   240
            Width           =   1335
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   6825
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   12840
            _ExtentX        =   22648
            _ExtentY        =   12044
            _Version        =   393216
            Rows            =   1
            Cols            =   11
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
            FormatString    =   "Statut|<Version|>Période   |Etab |<Matricule |Code autorisation|Devise|> Montant autorisation |> Echéance ||"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
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
      Picture         =   "DAUTPIB.frx":047A
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
         Size            =   9.6
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
   End
End
Attribute VB_Name = "frmDAUTPIB"
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
Dim BIA_DWH_Aut As typeAuthorization
Dim curX1 As Currency, curX2 As Currency
Dim blnAuto As Boolean

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim cnADO As New ADODB.Connection, rsado As New ADODB.Recordset, errADO As ADODB.Error
Dim blnTransaction As Boolean

'______________________________________________________________________

Dim wAmjMin As String, wAmjMax As String, wAmjEch As String, wHmsMin As Long, wHmsMax As Long
Dim xDAUTPIB As typeDAUTPIB, newDAUTPIB As typeDAUTPIB, oldDAUTPIB As typeDAUTPIB
Dim arrDAUTPIB() As typeDAUTPIB, arrDAUTPIB_Nb As Long, arrDAUTPIB_Max As Long, arrDAUTPIB_Index As Long
Dim xDDEVISE As typeDDEVISE
Dim xWhere_CDEV As String

'______________________________________________________________________

Dim fraDAUTPIB_FormatString As String, fraDAUTPIB_K As Integer
Dim fraDAUTPIB_RowDisplay As Integer, fraDAUTPIB_RowClick As Integer, fraDAUTPIB_ColClick As Integer
Dim fraDAUTPIB_ColorClick As Long, fraDAUTPIB_ColorDisplay As Long
Dim fraDAUTPIB_Sort1 As Integer, fraDAUTPIB_Sort2 As Integer
Dim fraDAUTPIB_SortAD As Integer, fraDAUTPIB_Sort1_Old As Integer
Dim fraDAUTPIB_arrIndex As Integer
Dim blnfraDAUTPIB_DisplayLine As Boolean

Dim meDAUTPIB_Status As typeDAUTPIB, oldDAUTPIB_Status As typeDAUTPIB

Dim rsADO_Update As New ADODB.Recordset
Dim mTitleText As String

Dim fgSelect_TopRow_Memo As Integer







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
currentAction = "fgSelect_Display"
    
For I = 1 To arrDAUTPIB_Nb
         
    xDAUTPIB = arrDAUTPIB(I)
    If xDAUTPIB.DAUTSTA <> "" Then
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine I
    End If
Next I

fgSelect.Visible = True
fraDAUTPIB.Visible = False
cmdSelect_Ok.Enabled = True
cmdSelect_Ajouter.Enabled = True

Call lstErr_AddItem(lstErr, cmdContext, "Opérations : " & arrDAUTPIB_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
    
End Sub

Private Sub arrDAUTPIB_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler
ReDim arrDAUTPIB(101)
arrDAUTPIB_Max = 100: arrDAUTPIB_Nb = 0

Set rsado = Nothing

xSql = "select * from " & paramIBM_Library_BODWH & ".DAUTPIB " & xWhere & " order by DAUTCLI, DAUTAUT, DAUTDEV"
Set rsado = cnADO.Execute(xSql)

Do While Not rsado.EOF
    V = srvDAUTPIB_GetBuffer_ODBC(rsado, xDAUTPIB)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmDAUTPIB.fgSelect_Display"
        '' Exit Sub
     Else
         arrDAUTPIB_Nb = arrDAUTPIB_Nb + 1
         If arrDAUTPIB_Nb > arrDAUTPIB_Max Then
             arrDAUTPIB_Max = arrDAUTPIB_Max + 50
             ReDim Preserve arrDAUTPIB(arrDAUTPIB_Max)
         End If
         
         arrDAUTPIB(arrDAUTPIB_Nb) = xDAUTPIB
    End If
    rsado.MoveNext
Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub DDEVISE_CDEV_SQL(xWhere_CDEV As String)

Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler

cbo_CDEV.Clear

Set rsado = Nothing

xSql = "select * from " & paramIBM_Library_BODWH & ".DDEVISE " & xWhere_CDEV & " order by DDEVDEV, DDEVLIB"
Set rsado = cnADO.Execute(xSql)

Do While Not rsado.EOF
    V = srvDDEVISE_GetBuffer_ODBC(rsado, xDDEVISE)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmDAUTPIB.DDEVISE_CDEV_SQL"
        '' Exit Sub
     Else
         cbo_CDEV.AddItem xDDEVISE.DDEVDEV & "  " & xDDEVISE.DDEVLIB
    End If
    rsado.MoveNext
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

fgSelect.Col = 0: fgSelect.Text = xDAUTPIB.DAUTSTA
fgSelect.Col = 1: fgSelect.Text = xDAUTPIB.DAUTVER
fgSelect.Col = 2: fgSelect.Text = dateImp10(xDAUTPIB.DAUTPER)
fgSelect.Col = 3: fgSelect.Text = xDAUTPIB.DAUTETB
fgSelect.Col = 4: fgSelect.Text = xDAUTPIB.DAUTCLI

fgSelect.Col = 5: fgSelect.Text = xDAUTPIB.DAUTAUT
fgSelect.Col = 6: fgSelect.Text = xDAUTPIB.DAUTDEV
fgSelect.Col = 7: fgSelect.Text = Format$(xDAUTPIB.DAUTMON, "### ### ### ##0.00")

If xDAUTPIB.DAUTMON = 0 Then
    fgSelect.Text = ""
Else
    ' Mettre en rouge les montants débiteurs
    If xDAUTPIB.DAUTMON >= 0 Then
        fgSelect.CellForeColor = vbBlue
    Else
        fgSelect.CellForeColor = vbRed
    End If
End If

fgSelect.Col = 8
If xDAUTPIB.DAUTECH = 0 Then
    fgSelect.Text = ""
Else
    fgSelect.Text = dateImp10(xDAUTPIB.DAUTECH)
End If

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

Public Sub cmdDAUTPIB_Charger()

srvDAUTPIB_Init newDAUTPIB

newDAUTPIB.DAUTSTA = Trim(txtfg_DAUTSTA)
newDAUTPIB.DAUTVER = Val(txtfg_DAUTVER)
newDAUTPIB.DAUTPER = wAmjMax   '''xDAUTPIB.DAUTPER
newDAUTPIB.DAUTETB = Trim(txtfg_DAUTETB)
newDAUTPIB.DAUTCLI = Val(Mid$(txtfg_DAUTCLI, 1, 7))

newDAUTPIB.DAUTAUT = Trim(txtfg_DAUTAUT)
newDAUTPIB.DAUTDEV = Trim(Mid$(cbo_CDEV.Text, 1, 3))

newDAUTPIB.DAUTMON = CCur(Val(txtfg_DAUTMON))

' Case coché = indication d'une date / Case vide = sans date
If chkfg_DAUTECH = "1" Then
    Call DTPicker_Control(txtfg_DAUTECH, wAmjEch)
    newDAUTPIB.DAUTECH = wAmjEch
Else
    newDAUTPIB.DAUTECH = 0
End If

End Sub


Public Sub cmdDAUTPIB_Update()

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'blnTransaction_Set

V = sqlDAUTPIB_Update(newDAUTPIB, xDAUTPIB, cnADO)

'If Not IsNull(V) Then
'    xSql = "Rollback"
'Else
'    xSql = "Commit"
'End If

'Set rsADO_Update = cnAdo.Execute(xSql, Nb)
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

If Not IsNull(V) Then MsgBox V, vbCritical

End Sub

Public Sub cmdDAUTPIB_Delete()

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'blnTransaction_Set

V = sqlDAUTPIB_Delete(newDAUTPIB, xDAUTPIB, cnADO)

'If Not IsNull(V) Then
'    xSql = "Rollback"
'Else
'    xSql = "Commit"
'End If

'Set rsADO_Update = cnAdo.Execute(xSql, Nb)
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

If Not IsNull(V) Then MsgBox V, vbCritical

End Sub

Public Sub cmdDAUTPIB_Insert()

Dim xSql As String, Nb As Integer
Dim wSWICLAOPR As String, wSWICLANUM As Long

xSql = "select * from " & paramIBM_Library_BODWH & ".DAUTPIB " _
    & "where DAUTVER = " & newDAUTPIB.DAUTVER _
    & " and DAUTPER = " & newDAUTPIB.DAUTPER _
    & " and DAUTETB = '" & newDAUTPIB.DAUTETB & "'" _
    & " and DAUTCLI = " & newDAUTPIB.DAUTCLI _
    & " and DAUTAUT = '" & newDAUTPIB.DAUTAUT & "'" _
    & " and DAUTDEV = '" & newDAUTPIB.DAUTDEV & "'"

Set rsado = cnADO.Execute(xSql)
If Not rsado.EOF Then
    MsgBox "Enregistrement existant dans DAUTPIB !", vbCritical, "fgDAUTPIB.cmd_Creer_Insert"
    Exit Sub
End If

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'blnTransaction_Set

' wAmjMax = Période de saisie
newDAUTPIB.DAUTPER = wAmjMax
V = sqlDAUTPIB_Insert(newDAUTPIB, cnADO)

'If Not IsNull(V) Then
'    xSql = "Rollback"
'Else
'    xSql = "Commit"
'End If

'Set rsADO_Update = cnAdo.Execute(xSql, Nb)
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

        
If Not IsNull(V) Then

    MsgBox V, vbCritical
    cmdfra_Creer.Visible = True
    cmdfra_Modifier.Visible = False
    cmdfra_Supprimer.Visible = False
    txtfg_DAUTSTA.Enabled = True
    txtfg_DAUTAUT.Enabled = True
    txtfg_DAUTETB.Enabled = True
    txtfg_DAUTCLI.Enabled = True
    txtfg_DAUTAUT.Enabled = True
    cbo_CDEV.Enabled = True
    
    fraDAUTPIB.Visible = True
    cmdSelect_Ok.Enabled = False
    cmdSelect_Ajouter.Enabled = False

Else
   
    ' Charger newDAUTPIB dans arrDAUTPIB puis
    ' Rafraichir l'affichage de la liste dans -fgSelect-
    ' Maintenir l'affichage de la ligne ajoutée sur la page en cours d'affichage
    
    arrDAUTPIB_Nb = arrDAUTPIB_Nb + 1
    arrDAUTPIB(arrDAUTPIB_Nb) = newDAUTPIB
    xDAUTPIB = newDAUTPIB
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_DisplayLine arrDAUTPIB_Nb

    Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
    fgSelect.TopRow = fgSelect.Row
    
    ' Maintenir l'affichage de la frame de gestion pour le mode -CREATION-
    xDAUTPIB.DAUTSTA = ""
    Call fraDAUTPIB_Display

End If

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

Call BiaPgmAut_Init(Mid$(Msg, 1, 12), BIA_DWH_Aut)
Form_Init

blnAuto = False

End Sub


Public Sub Form_Init()
Me.Enabled = False
Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents

If Not IsNull(param_Init) Then
    If Not blnAuto Then MsgBox "paramétrage inconsistant", vbCritical, "frmBIA_DWH.param_init"
    Unload Me
Else
    lstErr.Clear
End If
If Not IsNull(paramSAA_Init) Then
    If Not blnAuto Then MsgBox "paramétrage inconsistant", vbCritical, "frmBIA_DWH.paramSAA_Init"
    Unload Me
Else
    lstErr.Clear
End If


blnControl = False
fgSelect_FormatString = fgSelect.FormatString

fgSelect.Enabled = True
cmdReset

Me.Enabled = True
Me.MousePointer = 0
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
cmdSelect_Ok.Enabled = True
cmdSelect_Ajouter.Enabled = False
'cmdSelect_Ok_Click



blnControl = True



End Sub


Public Function param_Init()

param_Init = Null
Call lstErr_Clear(lstErr, cmdContext, "Param_Init"): DoEvents

fgSelect.Visible = False

Call DTPicker_Set(txtSelect_DAUTPER, DSys)

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




Private Sub cbo_CDEV_GotFocus()
txt_GotFocus cbo_CDEV
End Sub


Private Sub cbo_CDEV_LostFocus()
txt_LostFocus cbo_CDEV

End Sub



Private Sub cbo_CDEV_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub



Private Sub chkfg_DAUTECH_Click()
If chkfg_DAUTECH = "1" Then
    txtfg_DAUTECH.Visible = True
Else
    txtfg_DAUTECH.Visible = False
End If
    
End Sub

Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdfra_Creer_Click()

Me.Enabled = False: Me.MousePointer = vbHourglass
cmdDAUTPIB_Charger
cmdDAUTPIB_Insert
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdfra_Modifier_Click()

Me.Enabled = False: Me.MousePointer = vbHourglass
cmdDAUTPIB_Charger
cmdDAUTPIB_Update
Me.Enabled = True: Me.MousePointer = 0
fraDAUTPIB.Visible = False
cmdSelect_Ok.Enabled = True
cmdSelect_Ajouter.Enabled = True

' Rafraichir la liste d'affichage -fgSelect-
arrDAUTPIB(arrDAUTPIB_Index) = newDAUTPIB
xDAUTPIB = newDAUTPIB
fgSelect_DisplayLine arrDAUTPIB_Index

End Sub


Private Sub cmdfra_Supprimer_Click()
Dim X As String

X = MsgBox("Confirmer la suppression ?", vbQuestion + vbOKCancel, "BIA_DWH : DAUTPIB")

If X = vbOK Then
    Me.Enabled = False: Me.MousePointer = vbHourglass
    cmdDAUTPIB_Charger
    cmdDAUTPIB_Delete
    fraDAUTPIB.Visible = False
    cmdSelect_Ok.Enabled = True
    cmdSelect_Ajouter.Enabled = True
    Me.Enabled = True: Me.MousePointer = 0
    
    fgSelect_TopRow_Memo = fgSelect.TopRow
    ' Rafraichir la liste sur -fgSelect- et positionner à la page en cours d'affichage
    cmdSelect_SQL
    If fgSelect.Rows > 1 Then fgSelect.TopRow = fgSelect_TopRow_Memo
Else
    fraDAUTPIB.Visible = True
    cmdSelect_Ok.Enabled = False
    cmdSelect_Ajouter.Enabled = False
End If

End Sub


Private Sub cmdPrint_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Select Case SSTab1.Tab
    Case 0:
           ' If fgSelect.Rows > 1 Then cmdPrint_Ok
                
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

currentAction = "cmdDAUTPIB_SQL"
Call DTPicker_Control(txtSelect_DAUTPER, wAmjMax)
xWhere_CDEV = ""
xWhere = " where DAUTPER = " & wAmjMax
  
DDEVISE_CDEV_SQL xWhere_CDEV

arrDAUTPIB_SQL xWhere
fgSelect_Display

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_Ajouter_Click()

Dim blnOk As Boolean, Nb As Long

blnOk = fraSelect_Options.Enabled
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> BIA_DWH_cmdSelect_Ajouter........"): DoEvents

xDAUTPIB.DAUTSTA = ""
Call fraDAUTPIB_Display

Me.Enabled = True: Me.MousePointer = 0

End Sub



Private Sub cmdSelect_Ok_Click()
Dim blnOk As Boolean, Nb As Long

blnOk = fraSelect_Options.Enabled
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> BIA_DWH_cmdSelect_Ok ........"): DoEvents

fgSelect.Clear
fraDAUTPIB.Visible = False
cmdSelect_Ok.Enabled = True
cmdSelect_Ajouter.Enabled = True

If blnOk Then
    cmdSelect_Ok.Caption = "Nouveau mois"
    cmdSelect_Ok.BackColor = &HFFFFFA   '&HC0FFFF
    fraSelect_Options.BackColor = &H8000000F
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Enabled = False
    cmdSelect_Ajouter.Visible = True
    cmdSelect_SQL
Else
    cmdSelect_Ok.Caption = constcmdRechercher
    cmdSelect_Ok.BackColor = &HC0FFC0
    fraSelect_Options.BackColor = &HFFFFFA    '&HC0FFFF
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Enabled = True
    cmdSelect_Ajouter.Visible = False
End If
Call lstErr_AddItem(lstErr, cmdContext, "< BIA_DWH_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim K As Long
On Error Resume Next
If Y <= fgSelect.RowHeightMin Then
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
        fgSelect.Col = fgSelect_arrIndex:  arrDAUTPIB_Index = CLng(fgSelect.Text)
        fgSelect.LeftCol = 0
        
        oldDAUTPIB = arrDAUTPIB(arrDAUTPIB_Index)
        xDAUTPIB = oldDAUTPIB
        Call fraDAUTPIB_Display
   End If
End If
fgSelect.LeftCol = 0
End Sub

Public Sub fraDAUTPIB_DisplayLine()

On Error Resume Next

 txtfg_DAUTSTA = xDAUTPIB.DAUTSTA
 txtfg_DAUTVER = xDAUTPIB.DAUTVER
 txtfg_DAUTETB = xDAUTPIB.DAUTETB
 txtfg_DAUTCLI = xDAUTPIB.DAUTCLI
 txtfg_DAUTAUT = xDAUTPIB.DAUTAUT
 Call cbo_Scan(xDAUTPIB.DAUTDEV, cbo_CDEV)
 txtfg_DAUTMON = cur_P(xDAUTPIB.DAUTMON)
 If xDAUTPIB.DAUTECH = 0 Then
    chkfg_DAUTECH = "0"
    Call DTPicker_Set(txtfg_DAUTECH, DSys)
 Else
    chkfg_DAUTECH = "1"
    wAmjEch = xDAUTPIB.DAUTECH
    Call DTPicker_Set(txtfg_DAUTECH, wAmjEch)
 End If

End Sub

Private Sub fraDAUTPIB_Display()
Dim V
Dim X As String, xSql As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim I As Long

On Error GoTo Error_Handler

fraDAUTPIB.Visible = False
cmdSelect_Ok.Enabled = True
cmdSelect_Ajouter.Enabled = True
cmdfra_Creer.Visible = False
cmdfra_Modifier.Visible = False
cmdfra_Supprimer.Visible = False

currentAction = "fraDAUTPIB_Display"

If Trim(xDAUTPIB.DAUTSTA) <> "" Then   ' Autres que Création
   
    fraDAUTPIB_DisplayLine
    
    cmdfra_Creer.Visible = False
    cmdfra_Modifier.Visible = True
    cmdfra_Supprimer.Visible = True
    txtfg_DAUTVER.Enabled = False
    txtfg_DAUTETB.Enabled = False
    txtfg_DAUTCLI.Enabled = False
    txtfg_DAUTAUT.Enabled = False
    cbo_CDEV.Enabled = False

Else  ' Création
   
    txtfg_DAUTSTA = "W"
    txtfg_DAUTVER = 1
    txtfg_DAUTETB = "01"
    txtfg_DAUTCLI = ""
    txtfg_DAUTAUT = ""
    txtfg_DAUTMON = ""
    Call DTPicker_Set(txtfg_DAUTECH, DSys)
    chkfg_DAUTECH = "1"
   
    cmdfra_Creer.Visible = True
    cmdfra_Modifier.Visible = False
    cmdfra_Supprimer.Visible = False
    txtfg_DAUTVER.Enabled = True
    txtfg_DAUTETB.Enabled = True
    txtfg_DAUTCLI.Enabled = True
    txtfg_DAUTAUT.Enabled = True
    cbo_CDEV.Enabled = True

End If

fraDAUTPIB.Visible = True
cmdSelect_Ok.Enabled = False
cmdSelect_Ajouter.Enabled = False

Call lstErr_AddItem(lstErr, cmdContext, "Lignes : "): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 1
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'If blnControl Then
    cnADO.Close
    Set cnADO = Nothing

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
If fraDAUTPIB.Visible Then
    fraDAUTPIB.Visible = False
    cmdSelect_Ok.Enabled = True
    cmdSelect_Ajouter.Enabled = True
    Exit Sub
End If
    
If SSTab1.Tab = 0 Then
        Unload Me
    Exit Sub
Else
    SSTab1.Tab = SSTab1.Tab - 1
End If

End Sub

Public Sub cmdContext_Return()
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

Set cnADO = New ADODB.Connection
cnADO.CursorLocation = adUseClient
cnADO.Open paramODBC_DSN_SAB

Exit Sub

Error_Handler:

blnControl = False
If Not blnAuto Then MsgBox Error
End Sub





Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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

End Sub














Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
If SSTab1.Tab = 0 Then cmdSelect_Ok.SetFocus

End Sub














Private Sub SSTab1_GotFocus()
Select Case SSTab1.Tab
    Case 0: fgSelect.LeftCol = 0
End Select
End Sub






Public Sub blnTransaction_Set()
If Not blnTransaction Then
    blnTransaction = True
    Set rsADO_Update = cnADO.Execute("SET TRANSACTION ISOLATION LEVEL READ COMMITTED")

End If

End Sub











Private Sub txtfg_DAUTCLI_LostFocus()
txt_LostFocus txtfg_DAUTCLI

End Sub

Private Sub txtfg_DAUTETB_GotFocus()
txt_GotFocus txtfg_DAUTETB
End Sub






Private Sub txtfg_DAUTAUT_GotFocus()
txt_GotFocus txtfg_DAUTAUT
End Sub


Private Sub txtfg_DAUTAUT_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtfg_DAUTAUT_LostFocus()
txt_LostFocus txtfg_DAUTAUT
End Sub


Private Sub txtfg_DAUTCLI_GotFocus()
txt_GotFocus txtfg_DAUTCLI
End Sub




Private Sub txtfg_DAUTETB_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtfg_DAUTMON_KeyPress(KeyAscii As Integer)
num_KeyAsciiD KeyAscii, txtfg_DAUTMON
End Sub


Private Sub txtfg_DAUTMON_LostFocus()
txtfg_DAUTMON = cur_P(CCur(Fix(Val(txtfg_DAUTMON) * 100) / 100))
End Sub


Private Sub txtfg_DAUTSTA_GotFocus()
txt_GotFocus txtfg_DAUTSTA
End Sub

Private Sub txtfg_DAUTSTA_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtfg_DAUTSTA_LostFocus()
txt_LostFocus txtfg_DAUTSTA

End Sub


Private Sub txtfg_DAUTVER_GotFocus()
txt_GotFocus txtfg_DAUTVER
End Sub


