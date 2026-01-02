VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmDCOMM 
   AutoRedraw      =   -1  'True
   Caption         =   "BIA_DWH"
   ClientHeight    =   9144
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   13560
   Icon            =   "DCOMM.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9144
   ScaleWidth      =   13560
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   432
      Left            =   7800
      TabIndex        =   13
      Top             =   0
      Width           =   5175
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8655
      Left            =   0
      TabIndex        =   12
      Top             =   480
      Width           =   13530
      _ExtentX        =   23855
      _ExtentY        =   15261
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Gestion sur les mois traités"
      TabPicture(0)   =   "DCOMM.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "....."
      TabPicture(1)   =   "DCOMM.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.Frame fraTab0 
         Height          =   8205
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   13290
         Begin VB.CommandButton cmdSelect_Ajouter 
            BackColor       =   &H00FFC0FF&
            Caption         =   "Ajouter"
            Enabled         =   0   'False
            Height          =   285
            Left            =   11640
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   720
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Frame fraDCOMM 
            BackColor       =   &H80000013&
            Caption         =   "Gestion du fichier - DCOMM -"
            Height          =   5775
            Left            =   6600
            TabIndex        =   25
            Top             =   2040
            Visible         =   0   'False
            Width           =   6015
            Begin VB.TextBox txtfg_DCOMCTL2 
               Enabled         =   0   'False
               Height          =   300
               Left            =   3960
               TabIndex        =   44
               Text            =   "Comm"
               Top             =   3480
               Width           =   1455
            End
            Begin VB.TextBox txtfg_DCOMCTL1 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2160
               TabIndex        =   43
               Text            =   "Comm"
               Top             =   3480
               Width           =   1455
            End
            Begin VB.TextBox txtfg_DCOMCOM 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2160
               TabIndex        =   41
               Text            =   "Comm"
               Top             =   2760
               Width           =   855
            End
            Begin VB.TextBox txtfg_DCOMCLIA 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2160
               TabIndex        =   39
               Text            =   " /T"
               Top             =   2400
               Width           =   375
            End
            Begin VB.TextBox txtfg_DCOMSEN 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2160
               TabIndex        =   33
               Text            =   "Sens"
               Top             =   1560
               Width           =   375
            End
            Begin VB.TextBox txtfg_DCOMCLE 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2160
               TabIndex        =   38
               Text            =   "Référence"
               Top             =   2040
               Width           =   2535
            End
            Begin VB.TextBox txtfg_DCOMNAT 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2640
               TabIndex        =   31
               Text            =   "Nature"
               Top             =   1200
               Width           =   855
            End
            Begin VB.TextBox txtfg_DCOMSEQ 
               Height          =   300
               Left            =   2640
               TabIndex        =   35
               Text            =   "Seq"
               Top             =   1560
               Width           =   375
            End
            Begin VB.TextBox txtfg_DCOMNUM 
               Height          =   300
               Left            =   3600
               TabIndex        =   32
               Text            =   "No opération"
               Top             =   1200
               Width           =   1095
            End
            Begin VB.TextBox txtfg_DCOMOPE 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2160
               TabIndex        =   30
               Text            =   "Opération"
               Top             =   1200
               Width           =   495
            End
            Begin VB.TextBox txtfg_DCOMSER 
               Enabled         =   0   'False
               Height          =   300
               Left            =   3120
               TabIndex        =   21
               Text            =   "Service"
               Top             =   840
               Width           =   375
            End
            Begin VB.TextBox txtfg_DCOMAGE 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2640
               TabIndex        =   20
               Text            =   "Agence"
               Top             =   840
               Width           =   375
            End
            Begin VB.TextBox txtfg_DCOMSES 
               Enabled         =   0   'False
               Height          =   300
               Left            =   3600
               TabIndex        =   22
               Text            =   "Sous-service"
               Top             =   840
               Width           =   375
            End
            Begin VB.ComboBox cbo_CDEV 
               Height          =   315
               Left            =   2160
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   47
               Top             =   4200
               Width           =   2295
            End
            Begin VB.TextBox txtfg_DCOMMONB 
               Height          =   300
               Left            =   2160
               TabIndex        =   45
               Text            =   "Montant comm Base"
               Top             =   3840
               Width           =   1695
            End
            Begin VB.ComboBox cbo_CRTA 
               Height          =   315
               Left            =   2160
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   42
               Top             =   3120
               Width           =   2775
            End
            Begin VB.CommandButton cmdfra_Creer 
               BackColor       =   &H00FFC0FF&
               Caption         =   "Créer"
               Height          =   495
               Left            =   5040
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   5040
               Width           =   855
            End
            Begin VB.CommandButton cmdfra_Modifier 
               BackColor       =   &H00FF80FF&
               Caption         =   "Modifier"
               Height          =   495
               Left            =   3960
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   5040
               Width           =   855
            End
            Begin VB.CommandButton cmdfra_Supprimer 
               BackColor       =   &H000000FF&
               Caption         =   "Supprimer"
               Height          =   495
               Left            =   2760
               Style           =   1  'Graphical
               TabIndex        =   26
               Top             =   5040
               Width           =   855
            End
            Begin VB.TextBox txtfg_DCOMMOND 
               Height          =   300
               Left            =   3960
               TabIndex        =   46
               Text            =   "Montant comm Cevise"
               Top             =   3840
               Width           =   1695
            End
            Begin VB.TextBox txtfg_DCOMCLIB 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2640
               TabIndex        =   40
               Text            =   "No client"
               Top             =   2400
               Width           =   855
            End
            Begin VB.TextBox txtfg_DCOMETA 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2160
               TabIndex        =   19
               Text            =   "Etablissement"
               Top             =   840
               Width           =   375
            End
            Begin VB.TextBox txtfg_DCOMVER 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2640
               TabIndex        =   10
               Text            =   "Version"
               Top             =   480
               Width           =   375
            End
            Begin VB.TextBox txtfg_DCOMSTA 
               Height          =   300
               Left            =   2160
               TabIndex        =   9
               Text            =   "Statut"
               Top             =   480
               Width           =   375
            End
            Begin VB.Label lblfg_DCOMCTL 
               BackColor       =   &H80000013&
               Caption         =   "Code contrôle 1 / 2"
               Height          =   300
               Left            =   240
               TabIndex        =   37
               Top             =   3480
               Width           =   1815
            End
            Begin VB.Label lblfg_DCOMCOM 
               BackColor       =   &H80000013&
               Caption         =   "Code commission"
               Height          =   300
               Left            =   240
               TabIndex        =   36
               Top             =   2760
               Width           =   1695
            End
            Begin VB.Label lblfg_DCOMDEV 
               BackColor       =   &H80000013&
               Caption         =   "Code devise"
               Height          =   300
               Left            =   240
               TabIndex        =   7
               Top             =   4320
               Width           =   1815
            End
            Begin VB.Label lblfg_DCOMCLE 
               BackColor       =   &H80000013&
               Caption         =   "CLE"
               Height          =   300
               Left            =   240
               TabIndex        =   34
               Top             =   2040
               Width           =   1695
            End
            Begin VB.Label lblfg_DCOMSEN_SEQ 
               BackColor       =   &H80000013&
               Caption         =   "Sens / Seq"
               Height          =   300
               Left            =   240
               TabIndex        =   4
               Top             =   1560
               Width           =   1815
            End
            Begin VB.Label lblfg_DCOMOPE_NUM 
               BackColor       =   &H80000013&
               Caption         =   "Opération / Nature / no"
               Height          =   300
               Left            =   240
               TabIndex        =   3
               Top             =   1200
               Width           =   1815
            End
            Begin VB.Label lblfg_DCOMMON 
               BackColor       =   &H80000013&
               Caption         =   "Mnts Comm Base/Devise"
               Height          =   375
               Left            =   240
               TabIndex        =   6
               Top             =   3840
               Width           =   1815
            End
            Begin VB.Label lblfg_DCOMCRTA 
               BackColor       =   &H80000013&
               Caption         =   "Code renta"
               Height          =   300
               Left            =   240
               TabIndex        =   8
               Top             =   3120
               Width           =   1815
            End
            Begin VB.Label lblfg_DCOMCLI 
               BackColor       =   &H80000013&
               Caption         =   "Tiers / Client "
               Height          =   300
               Left            =   240
               TabIndex        =   5
               Top             =   2400
               Width           =   1815
            End
            Begin VB.Label lblfg_DCOMETA 
               BackColor       =   &H80000013&
               Caption         =   "Etab - Agce - SER - SSE"
               Height          =   300
               Left            =   240
               TabIndex        =   2
               Top             =   840
               Width           =   1815
            End
            Begin VB.Label lblfg_DCOMSTA_VER 
               BackColor       =   &H80000013&
               Caption         =   "Statut / Version"
               Height          =   300
               Left            =   240
               TabIndex        =   1
               Top             =   480
               Width           =   1815
            End
         End
         Begin VB.Frame fraSelect_Options 
            Height          =   1005
            Left            =   240
            TabIndex        =   17
            Top             =   120
            Width           =   11355
            Begin MSComCtl2.DTPicker txtSelect_DCOMPER 
               Height          =   300
               Left            =   1920
               TabIndex        =   18
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
            Begin VB.Label lblSelect_DRTAPER 
               Caption         =   "Période de traitement"
               Height          =   255
               Left            =   120
               TabIndex        =   24
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
            TabIndex        =   16
            Top             =   240
            Width           =   1335
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   6825
            Left            =   120
            TabIndex        =   23
            Top             =   1440
            Width           =   12840
            _ExtentX        =   22648
            _ExtentY        =   12044
            _Version        =   393216
            Rows            =   1
            Cols            =   24
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
            FormatString    =   $"DCOMM.frx":047A
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
      TabIndex        =   11
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   13080
      Picture         =   "DCOMM.frx":055F
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
      TabIndex        =   14
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
Attribute VB_Name = "frmDCOMM"
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

Dim wAmjMin As String, wAmjMax As String, wHmsMin As Long, wHmsMax As Long
Dim xDCOMM As typeDCOMM, newDCOMM As typeDCOMM, oldDCOMM As typeDCOMM
Dim arrDCOMM() As typeDCOMM, arrDCOMM_Nb As Long, arrDCOMM_Max As Long, arrDCOMM_Index As Long
Dim xDRTAGRP As typeDRTAGRP
Dim xWhere_CRTA As String
Dim xDDEVISE As typeDDEVISE
Dim xWhere_CDEV As String

'______________________________________________________________________

Dim fraDCOMM_FormatString As String, fraDCOMM_K As Integer
Dim fraDCOMM_RowDisplay As Integer, fraDCOMM_RowClick As Integer, fraDCOMM_ColClick As Integer
Dim fraDCOMM_ColorClick As Long, fraDCOMM_ColorDisplay As Long
Dim fraDCOMM_Sort1 As Integer, fraDCOMM_Sort2 As Integer
Dim fraDCOMM_SortAD As Integer, fraDCOMM_Sort1_Old As Integer
Dim fraDCOMM_arrIndex As Integer
Dim blnfraDCOMM_DisplayLine As Boolean

Dim meDCOMM_Status As typeDCOMM, oldDCOMM_Status As typeDCOMM

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
currentAction = "fgselect_Display"
    
For I = 1 To arrDCOMM_Nb
         
    xDCOMM = arrDCOMM(I)
    If xDCOMM.DCOMSTA <> "" Then
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine I
    End If
Next I

fgSelect.Visible = True
fraDCOMM.Visible = False
cmdSelect_Ok.Enabled = True
cmdSelect_Ajouter.Enabled = True

Call lstErr_AddItem(lstErr, cmdContext, "Opérations : " & arrDCOMM_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
    
End Sub

Private Sub arrDCOMM_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler
ReDim arrDCOMM(101)
arrDCOMM_Max = 100: arrDCOMM_Nb = 0

Set rsado = Nothing

xSql = "select * from " & paramIBM_Library_BODWH & ".DCOMM " & xWhere & " order by DCOMVER, DCOMETA, DCOMAGE, DCOMSER, DCOMSES, DCOMOPE, DCOMNAT, DCOMNUM, DCOMSEN, DCOMSEQ"
Set rsado = cnADO.Execute(xSql)

Do While Not rsado.EOF
    V = srvDCOMM_GetBuffer_ODBC(rsado, xDCOMM)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmDCOMM.fgselect_Display"
        '' Exit Sub
     Else
         arrDCOMM_Nb = arrDCOMM_Nb + 1
         If arrDCOMM_Nb > arrDCOMM_Max Then
             arrDCOMM_Max = arrDCOMM_Max + 50
             ReDim Preserve arrDCOMM(arrDCOMM_Max)
         End If
         
         arrDCOMM(arrDCOMM_Nb) = xDCOMM
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

Private Sub DRTAGRP_CRTA_SQL(xWhere_CRTA As String)

Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler

cbo_CRTA.Clear

Set rsado = Nothing

xSql = "select * from " & paramIBM_Library_BODWH & ".DRTAGRP " & xWhere_CRTA & " order by DRGRCRTA"
Set rsado = cnADO.Execute(xSql)

Do While Not rsado.EOF
    V = srvDRTAGRP_GetBuffer_ODBC(rsado, xDRTAGRP)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmDCOMM.DRTAGRP_CRTA_SQL"
        '' Exit Sub
     Else
         cbo_CRTA.AddItem xDRTAGRP.DRGRCRTA & "  " & xDRTAGRP.DRGRLIB
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
         If Not blnAuto Then MsgBox V, vbCritical, "frmDCOMM.DDEVISE_CDEV_SQL"
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

fgSelect.Col = 0: fgSelect.Text = xDCOMM.DCOMSTA
fgSelect.Col = 1: fgSelect.Text = xDCOMM.DCOMVER
fgSelect.Col = 2: fgSelect.Text = dateImp10(xDCOMM.DCOMPER)
fgSelect.Col = 3: fgSelect.Text = xDCOMM.DCOMETA
fgSelect.Col = 4: fgSelect.Text = xDCOMM.DCOMAGE
fgSelect.Col = 5: fgSelect.Text = xDCOMM.DCOMSER
fgSelect.Col = 6: fgSelect.Text = xDCOMM.DCOMSES
fgSelect.Col = 7: fgSelect.Text = xDCOMM.DCOMOPE
fgSelect.Col = 8: fgSelect.Text = xDCOMM.DCOMNAT
fgSelect.Col = 9: fgSelect.Text = xDCOMM.DCOMNUM
fgSelect.Col = 10: fgSelect.Text = xDCOMM.DCOMSEN
fgSelect.Col = 11: fgSelect.Text = xDCOMM.DCOMSEQ

fgSelect.Col = 12: fgSelect.Text = xDCOMM.DCOMCLE
fgSelect.Col = 13: fgSelect.Text = xDCOMM.DCOMCLIA
fgSelect.Col = 14: fgSelect.Text = xDCOMM.DCOMCLIB
fgSelect.Col = 15: fgSelect.Text = xDCOMM.DCOMCOM
fgSelect.Col = 16: fgSelect.Text = xDCOMM.DCOMCRTA
fgSelect.Col = 17: fgSelect.Text = xDCOMM.DCOMCTL1
fgSelect.Col = 18: fgSelect.Text = xDCOMM.DCOMCTL2

fgSelect.Col = 19
If xDCOMM.DCOMMONB = 0 Then
    fgSelect.Text = ""
Else
    fgSelect.Text = Format$(xDCOMM.DCOMMONB, "### ### ### ##0.00")
    ' Mettre en rouge les montants débiteurs
    'If xDCOMM.DCOMMNT1 >= 0 Then
    '    fgSelect.CellForeColor = vbBlue
    'Else
    '    fgSelect.CellForeColor = vbRed
    'End If
End If
fgSelect.Col = 20
If xDCOMM.DCOMMOND = 0 Then
    fgSelect.Text = ""
Else
    fgSelect.Text = Format$(xDCOMM.DCOMMOND, "### ### ### ##0.00")
    ' Mettre en rouge les montants débiteurs
    'If xDCOMM.DCOMMNT2 >= 0 Then
    '    fgSelect.CellForeColor = vbBlue
    'Else
    '    fgSelect.CellForeColor = vbRed
    'End If
End If

fgSelect.Col = 21: fgSelect.Text = xDCOMM.DCOMDEV

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

Public Sub cmdDCOMM_Charger()

srvDCOMM_Init newDCOMM

newDCOMM.DCOMSTA = Trim(txtfg_DCOMSTA)
newDCOMM.DCOMVER = Val(txtfg_DCOMVER)
newDCOMM.DCOMPER = wAmjMax

newDCOMM.DCOMETA = Trim(txtfg_DCOMETA)
newDCOMM.DCOMAGE = Trim(txtfg_DCOMAGE)
newDCOMM.DCOMSER = Trim(txtfg_DCOMSER)
newDCOMM.DCOMSES = Trim(txtfg_DCOMSES)

newDCOMM.DCOMOPE = Trim(txtfg_DCOMOPE)
newDCOMM.DCOMNAT = Trim(txtfg_DCOMNAT)
newDCOMM.DCOMNUM = Val(Mid$(txtfg_DCOMNUM, 1, 9))
newDCOMM.DCOMSEN = Trim(txtfg_DCOMSEN)
newDCOMM.DCOMSEQ = Val(Mid$(txtfg_DCOMSEQ, 1, 3))

newDCOMM.DCOMCLE = Trim(txtfg_DCOMCLE)
newDCOMM.DCOMCLIA = Trim(txtfg_DCOMCLIA)
newDCOMM.DCOMCLIB = Val(Mid$(txtfg_DCOMCLIB, 1, 7))
newDCOMM.DCOMCOM = Trim(txtfg_DCOMCOM)
newDCOMM.DCOMCRTA = Val(Mid$(cbo_CRTA.Text, 1, 5))
newDCOMM.DCOMCTL1 = Trim(txtfg_DCOMCTL1)
newDCOMM.DCOMCTL2 = Trim(txtfg_DCOMCTL2)

' Montant tel qu'il est saisi : négatif ou positif
newDCOMM.DCOMMONB = CCur(Val(txtfg_DCOMMONB))
newDCOMM.DCOMMOND = CCur(Val(txtfg_DCOMMOND))
newDCOMM.DCOMDEV = Trim(Mid$(cbo_CDEV.Text, 1, 3))

End Sub


Public Sub cmdDCOMM_Update()

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'blnTransaction_Set

V = sqlDCOMM_Update(newDCOMM, xDCOMM, cnADO)

'If Not IsNull(V) Then
'    xSql = "Rollback"
'Else
'    xSql = "Commit"
'End If

'Set rsADO_Update = cnAdo.Execute(xSql, Nb)
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

If Not IsNull(V) Then MsgBox V, vbCritical

End Sub

Public Sub cmdDCOMM_Delete()

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'blnTransaction_Set

V = sqlDCOMM_Delete(newDCOMM, xDCOMM, cnADO)

'If Not IsNull(V) Then
'    xSql = "Rollback"
'Else
'    xSql = "Commit"
'End If

'Set rsADO_Update = cnAdo.Execute(xSql, Nb)
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

If Not IsNull(V) Then MsgBox V, vbCritical

End Sub

Public Sub cmdDCOMM_Insert()

Dim xSql As String, Nb As Integer
Dim wSWICLAOPR As String, wSWICLANUM As Long

xSql = "select * from " & paramIBM_Library_BODWH & ".DCOMM " _
    & "where DCOMVER = " & newDCOMM.DCOMVER _
    & " and DCOMPER = " & newDCOMM.DCOMPER _
    & " and DCOMETA = '" & newDCOMM.DCOMETA & "'" _
    & " and DCOMAGE = '" & newDCOMM.DCOMAGE & "'" _
    & " and DCOMSER = '" & newDCOMM.DCOMSER & "'" _
    & " and DCOMSES = '" & newDCOMM.DCOMSES & "'" _
    & " and DCOMOPE = '" & newDCOMM.DCOMOPE & "'" _
    & " and DCOMNAT = '" & newDCOMM.DCOMNAT & "'" _
    & " and DCOMNUM = " & newDCOMM.DCOMNUM _
    & " and DCOMSEN = '" & newDCOMM.DCOMSEN & "'" _
    & " and DCOMSEQ = " & newDCOMM.DCOMSEQ

Set rsado = cnADO.Execute(xSql)
If Not rsado.EOF Then
    MsgBox "Enregistrement existant dans DCOMM !", vbCritical, "fgDCOMM.cmd_Creer_Insert"
    Exit Sub
End If

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'blnTransaction_Set

' wAmjMax = Période de saisie
newDCOMM.DCOMPER = wAmjMax
V = sqlDCOMM_Insert(newDCOMM, cnADO)

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
    txtfg_DCOMVER.Enabled = False
    txtfg_DCOMETA.Enabled = False
    txtfg_DCOMAGE.Enabled = False
    txtfg_DCOMSER.Enabled = False
    txtfg_DCOMSES.Enabled = False
    txtfg_DCOMOPE.Enabled = False
    txtfg_DCOMNAT.Enabled = False
    txtfg_DCOMNUM.Enabled = False
    txtfg_DCOMSEN.Enabled = False
    txtfg_DCOMSEQ.Enabled = False
    
    fraDCOMM.Visible = True
    cmdSelect_Ok.Enabled = False
    cmdSelect_Ajouter.Enabled = False

Else
   
    ' Charger newDCOMM dans arrDCOMM puis
    ' Rafraichir l'affichage de la liste dans -fgSelect-
    ' Maintenir l'affichage de la ligne ajoutée sur la page en cours d'affichage
    
    arrDCOMM_Nb = arrDCOMM_Nb + 1
    arrDCOMM(arrDCOMM_Nb) = newDCOMM
    xDCOMM = newDCOMM
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_DisplayLine arrDCOMM_Nb

    Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
    fgSelect.TopRow = fgSelect.Row
     
    ' Maintenir l'affichage de la frame de gestion pour le mode -CREATION-
    xDCOMM.DCOMSTA = ""
    Call fraDCOMM_Display

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

blnControl = True

End Sub


Public Function param_Init()

param_Init = Null
Call lstErr_Clear(lstErr, cmdContext, "Param_Init"): DoEvents

fgSelect.Visible = False

Call DTPicker_Set(txtSelect_DCOMPER, DSys)

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


Private Sub cbo_CDEV_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub cbo_CDEV_LostFocus()
txt_LostFocus cbo_CDEV
End Sub


Private Sub cbo_CRTA_GotFocus()
txt_GotFocus cbo_CRTA
End Sub




Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdfra_Creer_Click()

Me.Enabled = False: Me.MousePointer = vbHourglass
cmdDCOMM_Charger
cmdDCOMM_Insert
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdfra_Modifier_Click()

Me.Enabled = False: Me.MousePointer = vbHourglass
cmdDCOMM_Charger
cmdDCOMM_Update
Me.Enabled = True: Me.MousePointer = 0
fraDCOMM.Visible = False
cmdSelect_Ok.Enabled = True
cmdSelect_Ajouter.Enabled = True

' Rafraichir la liste d'affichage -fgSelect-
arrDCOMM(arrDCOMM_Index) = newDCOMM
xDCOMM = newDCOMM
fgSelect_DisplayLine arrDCOMM_Index

End Sub


Private Sub cmdfra_Supprimer_Click()
Dim X As String
X = MsgBox("Confirmer la suppression ?", vbQuestion + vbOKCancel, "BIA_DWH : DCOMM")

If X = vbOK Then
    Me.Enabled = False: Me.MousePointer = vbHourglass
    cmdDCOMM_Charger
    cmdDCOMM_Delete
    fraDCOMM.Visible = False
    cmdSelect_Ok.Enabled = True
    cmdSelect_Ajouter.Enabled = True
    Me.Enabled = True: Me.MousePointer = 0
    
    fgSelect_TopRow_Memo = fgSelect.TopRow
    ' Rafraichir la liste sur -fgSelect- et positionner à la page en cours d'affichage
    cmdSelect_SQL
    If fgSelect.Rows > 1 Then fgSelect.TopRow = fgSelect_TopRow_Memo
Else
    fraDCOMM.Visible = True
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

currentAction = "cmdDCOMM_SQL"
Call DTPicker_Control(txtSelect_DCOMPER, wAmjMax)
xWhere_CRTA = " where DRGRPER = " & wAmjMax
xWhere = " where DCOMPER = " & wAmjMax
  
DRTAGRP_CRTA_SQL xWhere_CRTA

xWhere_CDEV = ""
DDEVISE_CDEV_SQL xWhere_CDEV

arrDCOMM_SQL xWhere
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

xDCOMM.DCOMSTA = ""
Call fraDCOMM_Display

Me.Enabled = True: Me.MousePointer = 0

End Sub



Private Sub cmdSelect_Ok_Click()
Dim blnOk As Boolean, Nb As Long

blnOk = fraSelect_Options.Enabled
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> BIA_DWH_cmdSelect_Ok ........"): DoEvents

fgSelect.Clear
fraDCOMM.Visible = False
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
        fgSelect.Col = fgSelect_arrIndex:  arrDCOMM_Index = CLng(fgSelect.Text)
        fgSelect.LeftCol = 0
        
        oldDCOMM = arrDCOMM(arrDCOMM_Index)
        xDCOMM = oldDCOMM
        Call fraDCOMM_Display
   End If
End If
fgSelect.LeftCol = 0
End Sub

Public Sub fraDCOMM_DisplayLine()
Dim X As String

On Error Resume Next

 txtfg_DCOMSTA = xDCOMM.DCOMSTA
 txtfg_DCOMVER = xDCOMM.DCOMVER
 txtfg_DCOMETA = xDCOMM.DCOMETA
 txtfg_DCOMAGE = xDCOMM.DCOMAGE
 txtfg_DCOMSER = xDCOMM.DCOMSER
 txtfg_DCOMSES = xDCOMM.DCOMSES
 txtfg_DCOMOPE = xDCOMM.DCOMOPE
 txtfg_DCOMNAT = xDCOMM.DCOMNAT
 txtfg_DCOMNUM = xDCOMM.DCOMNUM
 txtfg_DCOMSEN = xDCOMM.DCOMSEN
 txtfg_DCOMSEQ = xDCOMM.DCOMSEQ

 txtfg_DCOMCLE = xDCOMM.DCOMCLE
 txtfg_DCOMCLIA = xDCOMM.DCOMCLIA
 txtfg_DCOMCLIB = xDCOMM.DCOMCLIB
 txtfg_DCOMCOM = xDCOMM.DCOMCOM
 txtfg_DCOMCTL1 = xDCOMM.DCOMCTL1
 txtfg_DCOMCTL2 = xDCOMM.DCOMCTL2
 txtfg_DCOMMONB = cur_P(xDCOMM.DCOMMONB)
 txtfg_DCOMMOND = cur_P(xDCOMM.DCOMMOND)
 Call cbo_Scan(xDCOMM.DCOMDEV, cbo_CDEV)

 ' Le code renta du DCOMM peut être en 5 ou 4 ou 3 numériques rempli
 X = Format$(xDCOMM.DCOMCRTA, "00000")
 If Mid$(X, 1, 1) = "0" Then
    X = Format$(X, "0000")
    If Mid$(X, 1, 1) = "0" Then
        X = Format$(X, "000")
    End If
 End If
 Call cbo_Scan(X, cbo_CRTA)
 
End Sub

Private Sub fraDCOMM_Display()
Dim V
Dim X As String, xSql As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim I As Long

On Error GoTo Error_Handler

fraDCOMM.Visible = False
cmdSelect_Ok.Enabled = True
cmdSelect_Ajouter.Enabled = True
cmdfra_Creer.Visible = False
cmdfra_Modifier.Visible = False
cmdfra_Supprimer.Visible = False

currentAction = "fraDCOMM_Display"

If Trim(xDCOMM.DCOMSTA) <> "" Then   ' Autres que Création
   
    fraDCOMM_DisplayLine
    
    cmdfra_Creer.Visible = False
    cmdfra_Modifier.Visible = True
    cmdfra_Supprimer.Visible = True
    txtfg_DCOMVER.Enabled = False
    txtfg_DCOMETA.Enabled = False
    txtfg_DCOMAGE.Enabled = False
    txtfg_DCOMSER.Enabled = False
    txtfg_DCOMSES.Enabled = False
    txtfg_DCOMOPE.Enabled = False
    txtfg_DCOMNAT.Enabled = False
    txtfg_DCOMNUM.Enabled = False
    txtfg_DCOMSEN.Enabled = False
    txtfg_DCOMSEQ.Enabled = False

Else  ' Création
   
    txtfg_DCOMSTA = "W"
    txtfg_DCOMVER = 1
    txtfg_DCOMETA = "01"
    txtfg_DCOMAGE = "01"
    txtfg_DCOMSER = "00"
    txtfg_DCOMSES = "00"
    txtfg_DCOMOPE = ""
    txtfg_DCOMNAT = ""
    txtfg_DCOMNUM = ""
    txtfg_DCOMSEN = ""
    txtfg_DCOMSEQ = ""
    txtfg_DCOMCLE = ""
    txtfg_DCOMCLIA = ""
    txtfg_DCOMCLIB = ""
    txtfg_DCOMCOM = ""
    txtfg_DCOMCTL1 = ""
    txtfg_DCOMCTL2 = ""
    txtfg_DCOMMONB = ""
    txtfg_DCOMMOND = ""

    cmdfra_Creer.Visible = True
    cmdfra_Modifier.Visible = False
    cmdfra_Supprimer.Visible = False
    txtfg_DCOMVER.Enabled = False
    txtfg_DCOMETA.Enabled = False
    txtfg_DCOMAGE.Enabled = False
    
    txtfg_DCOMSER.Enabled = True
    txtfg_DCOMSES.Enabled = True
    txtfg_DCOMOPE.Enabled = True
    txtfg_DCOMNAT.Enabled = True
    txtfg_DCOMNUM.Enabled = True
    txtfg_DCOMSEN.Enabled = True
    txtfg_DCOMSEQ.Enabled = True

End If

fraDCOMM.Visible = True
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
If fraDCOMM.Visible Then
    fraDCOMM.Visible = False
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























Private Sub txtfg_DCOMAGE_GotFocus()
txt_GotFocus txtfg_DCOMAGE
End Sub











Private Sub txtfg_DCOMCLE_GotFocus()
txt_GotFocus txtfg_DCOMCLE

End Sub


Private Sub txtfg_DCOMCLE_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtfg_DCOMCLIA_GotFocus()
txt_GotFocus txtfg_DCOMCLIA
End Sub

Private Sub txtfg_DCOMCLIA_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtfg_DCOMCLIB_GotFocus()
txt_GotFocus txtfg_DCOMCLIB

End Sub


Private Sub txtfg_DCOMCLIB_LostFocus()
txt_LostFocus txtfg_DCOMCLIB

End Sub


Private Sub txtfg_DCOMCOM_GotFocus()
txt_GotFocus txtfg_DCOMCOM

End Sub


Private Sub txtfg_DCOMCOM_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtfg_DCOMCTL1_GotFocus()
txt_GotFocus txtfg_DCOMCTL1

End Sub


Private Sub txtfg_DCOMCTL1_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtfg_DCOMCTL2_GotFocus()
txt_GotFocus txtfg_DCOMCTL2

End Sub


Private Sub txtfg_DCOMCTL2_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtfg_DCOMMONB_KeyPress(KeyAscii As Integer)
num_KeyAsciiD KeyAscii, txtfg_DCOMMONB

End Sub


Private Sub txtfg_DCOMMONB_LostFocus()
txtfg_DCOMMONB = cur_P(CCur(Fix(Val(txtfg_DCOMMONB) * 100) / 100))

End Sub


Private Sub txtfg_DCOMMOND_KeyPress(KeyAscii As Integer)
num_KeyAsciiD KeyAscii, txtfg_DCOMMOND

End Sub

Private Sub txtfg_DCOMMOND_LostFocus()
txtfg_DCOMMOND = cur_P(CCur(Fix(Val(txtfg_DCOMMOND) * 100) / 100))

End Sub


Private Sub txtfg_DCOMNAT_GotFocus()
txt_GotFocus txtfg_DCOMNAT

End Sub


Private Sub txtfg_DCOMNAT_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtfg_DCOMETA_GotFocus()
txt_GotFocus txtfg_DCOMETA

End Sub


Private Sub txtfg_DCOMETA_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtfg_DCOMNUM_GotFocus()
txt_GotFocus txtfg_DCOMNUM

End Sub


Private Sub txtfg_DCOMNUM_LostFocus()
txt_LostFocus txtfg_DCOMNUM

End Sub


Private Sub txtfg_DCOMOPE_GotFocus()
txt_GotFocus txtfg_DCOMOPE

End Sub


Private Sub txtfg_DCOMOPE_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub



Private Sub txtfg_DCOMSEN_GotFocus()
txt_GotFocus txtfg_DCOMSEN

End Sub


Private Sub txtfg_DCOMSEN_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtfg_DCOMSEQ_GotFocus()
txt_GotFocus txtfg_DCOMSEQ

End Sub


Private Sub txtfg_DCOMSEQ_LostFocus()
txt_LostFocus txtfg_DCOMSEQ

End Sub


Private Sub txtfg_DCOMSER_GotFocus()
txt_GotFocus txtfg_DCOMSER

End Sub

Private Sub txtfg_DCOMSER_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtfg_DCOMSES_GotFocus()
txt_GotFocus txtfg_DCOMSES

End Sub


Private Sub txtfg_DCOMSES_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtfg_DCOMSTA_GotFocus()
txt_GotFocus txtfg_DCOMSTA
End Sub

Private Sub txtfg_DCOMSTA_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtfg_DCOMSTA_LostFocus()
txt_LostFocus txtfg_DCOMSTA

End Sub





Private Sub txtfg_DCOMVER_GotFocus()
txt_GotFocus txtfg_DCOMVER
End Sub





