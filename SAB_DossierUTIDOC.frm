VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmSAB_DossierUTIDOC 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tableau des documents (1er et 2ème jeu)."
   ClientHeight    =   10065
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7305
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10065
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Édition"
      Height          =   1905
      Left            =   30
      TabIndex        =   3
      Top             =   7590
      Width           =   7305
      Begin VB.CommandButton btnEnlever 
         Caption         =   "Enlever la sélection"
         Height          =   315
         Left            =   3750
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1260
         Width           =   1635
      End
      Begin VB.CommandButton btnModifier 
         Caption         =   "Modifier"
         Height          =   315
         Left            =   2700
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1260
         Width           =   765
      End
      Begin VB.TextBox txtDocument 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1260
         MaxLength       =   2000
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   600
         Width           =   5655
      End
      Begin VB.TextBox txtJeu2 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1260
         MaxLength       =   2000
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   1260
         Width           =   1155
      End
      Begin VB.TextBox txtJeu1 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1260
         MaxLength       =   2000
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   930
         Width           =   1155
      End
      Begin VB.Label labCle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N°"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   3180
         TabIndex        =   14
         Top             =   270
         Width           =   1155
      End
      Begin VB.Label labNum 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N°"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   255
         Left            =   1260
         TabIndex        =   11
         Top             =   330
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "N°"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   180
         TabIndex        =   7
         Top             =   330
         Width           =   180
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "2ème jeu"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   180
         TabIndex        =   6
         Top             =   1350
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "1er jeu"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   180
         TabIndex        =   5
         Top             =   1050
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Document"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   180
         TabIndex        =   4
         Top             =   720
         Width           =   825
      End
   End
   Begin VB.CommandButton btnValider 
      Caption         =   "Valider la page"
      Height          =   345
      Left            =   4620
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9600
      Width           =   1365
   End
   Begin VB.CommandButton btnFermer 
      Caption         =   "Fermer"
      Height          =   345
      Left            =   6090
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9600
      Width           =   1125
   End
   Begin MSFlexGridLib.MSFlexGrid fgUTI_DOC 
      Height          =   7575
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   13361
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      RowHeightMin    =   350
      BackColor       =   16448250
      ForeColor       =   16384
      BackColorFixed  =   33023
      ForeColorFixed  =   -2147483633
      BackColorBkg    =   16448250
      AllowUserResizing=   3
      FormatString    =   "<Document                                                                 |<1er Jeu            |<2ème Jeu       |>N°         "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmSAB_DossierUTIDOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub deselecte_tout()
Dim xCol As Long
Dim yRow As Long

    With fgUTI_DOC
        For yRow = 1 To .Rows - 1
            .Row = yRow
            .Col = 0
            If .CellBackColor = fgUTI_DOC.BackColorSel Then
                frmSAB_Dossier_RDE.fgUTI_DOC.Row = .Row
                frmSAB_Dossier_RDE.fgUTI_DOC.Col = 0
                If frmSAB_Dossier_RDE.fgUTI_DOC.CellBackColor = mColor_G1 Then
                    For xCol = 0 To .Cols - 1
                        .Col = xCol
                        .CellBackColor = mColor_G1
                    Next xCol
                    Exit For
                Else
                    For xCol = 0 To .Cols - 1
                        .Col = xCol
                        .CellBackColor = .BackColor
                    Next xCol
                    Exit For
                End If
            End If
        Next yRow
    End With
    
End Sub
Private Sub etat_normal()

    txtDocument.Text = ""
    txtJeu1.Text = ""
    txtJeu2.Text = ""
    labNum.Caption = ""
    labCle.Caption = ""
    btnModifier.Enabled = False
    btnEnlever.Enabled = False
    btnValider.Enabled = True
    fgUTI_DOC.SetFocus

End Sub

Private Sub select_ligne()
Dim xCol As Long

    With fgUTI_DOC
        For xCol = 0 To .Cols - 2
            .Col = xCol
            .CellBackColor = .BackColorSel
        Next xCol
    End With
    
End Sub


Private Sub btnEnlever_Click()
Dim xSql As String
Dim rs As ADODB.Recordset

    'restaurer le text d'origine
    With fgUTI_DOC
        xSql = "select * from " & paramIBM_Library_SAB & ".ZCDOTAB0" _
           & " where CDOTABETA = 1" _
           & " and CDOTABNUM = 19 and CDOTABARG ='" & labCle.Caption & "'"
        Set rs = cnsab.Execute(xSql)
        If Not rs.EOF Then
            .TextMatrix(.Row, 0) = Trim(rs("CDOTABDON"))
            .TextMatrix(.Row, 1) = ""
            .TextMatrix(.Row, 2) = ""
            .Col = 0: .CellBackColor = .BackColor
            .Col = 1: .CellBackColor = .BackColor
            .Col = 2: .CellBackColor = .BackColor
            .Col = 3: .CellBackColor = .BackColor
        Else
            .RemoveItem .Row
        End If
    End With
    rs.Close
    Set rs = Nothing
    Call etat_normal
    
End Sub

Private Sub btnFermer_Click()

    Unload Me
    
End Sub


Private Sub btnModifier_Click()

    With fgUTI_DOC
        .TextMatrix(.Row, 0) = txtDocument.Text
        .TextMatrix(.Row, 1) = txtJeu1.Text
        .TextMatrix(.Row, 2) = txtJeu2.Text
        .Col = 0: .CellBackColor = mColor_G1
        .Col = 1: .CellBackColor = mColor_G1
        .Col = 2: .CellBackColor = mColor_G1
        .Col = 3: .CellBackColor = mColor_G1
        .Col = 4: .CellBackColor = mColor_G1
    End With
    Call etat_normal
    
End Sub

Private Sub btnValider_Click()
Dim yRow As Long
Dim xCol As Long
Dim modif As Boolean
Dim colEnCours As Long
Dim ii As Long

    modif = False
    With frmSAB_Dossier_RDE.fgUTI_DOC
        For yRow = 0 To fgUTI_DOC.Rows - 1
            For xCol = 0 To fgUTI_DOC.Cols - 1
                .Row = yRow
                fgUTI_DOC.Row = yRow
                .Col = xCol
                fgUTI_DOC.Col = xCol
                If .Text <> fgUTI_DOC.Text Then modif = True
                .Text = fgUTI_DOC.Text
                'If .CellBackColor <> fgUTI_DOC.CellBackColor Then modif = True
                .Col = 4
                fgUTI_DOC.Col = 4
                colEnCours = fgUTI_DOC.CellBackColor
                For ii = 0 To .Cols - 2
                    .Col = ii
                    .CellBackColor = colEnCours
                Next ii
            Next xCol
        Next yRow
    End With
    If modif Then
        frmSAB_Dossier_RDE.fgInfo_M.Col = 2
        frmSAB_Dossier_RDE.fgInfo_M.CellBackColor = mColor_G0
        frmSAB_Dossier_RDE.fgInfo_M.Col = 3
        frmSAB_Dossier_RDE.fgInfo_M.CellBackColor = mColor_G0
    End If
    Call frmSAB_Dossier_RDE.cmdInfo_M_Ok_Visible
    Unload Me
    
End Sub

Public Sub fgUTI_DOC_Click()
Dim aligne As Long
Dim aEnlever As Boolean

    With fgUTI_DOC
        aligne = .Row
        aEnlever = False
        .Col = 0
        If .CellBackColor = mColor_G1 Then
            aEnlever = True
        End If
        Call deselecte_tout
        .Row = aligne
        Call select_ligne
        txtDocument.Text = .TextMatrix(.Row, 0)
        txtJeu1.Text = .TextMatrix(.Row, 1)
        txtJeu2.Text = .TextMatrix(.Row, 2)
        labNum.Caption = .TextMatrix(.Row, 3)
        labCle.Caption = .TextMatrix(.Row, 4)
        btnModifier.Enabled = False
        btnValider.Enabled = True
        btnEnlever.Enabled = False
        If aEnlever Then
            btnEnlever.Enabled = True
        End If
    End With
    
    txtJeu1.SetFocus
    
End Sub


Private Sub Form_Activate()
Dim ouSelection As Long
Dim ou As Long
Dim s() As String
Dim lDebug As String

    If fgUTI_DOC.Rows <= 1 Then
        s = Split(Me.Tag, "|")
        If UBound(s) > -1 Then
            lDebug = "mUTI_DOC_Index--> " & s(0) & vbCrLf
            lDebug = lDebug & "frmSAB_Dossier_RDE.fgUTI_DOC.Rows--> " & frmSAB_Dossier_RDE.fgUTI_DOC.Rows & vbCrLf
            lDebug = lDebug & "frmSAB_Dossier_UTIDOC.fgUTI_DOC.Rows--> " & fgUTI_DOC.Rows & vbCrLf
            lDebug = lDebug & "blnUTI_DOC_Loaded--> " & s(1) & vbCrLf
            lDebug = lDebug & "blnUTI_DOC_Ok--> " & s(2)
            MsgBox lDebug
        End If
    End If
    ouSelection = 1
    fgUTI_DOC.Col = 4
    For ou = 1 To fgUTI_DOC.Rows - 1
        fgUTI_DOC.Row = ou
        If fgUTI_DOC.CellBackColor = mColor_G1 Then
            ouSelection = ou
            Exit For
        End If
    Next ou
    fgUTI_DOC.Row = ouSelection
    fgUTI_DOC.SetFocus
    Call fgUTI_DOC_Click
    
End Sub

Private Sub Form_Load()
Dim yRow As Long
Dim xCol As Long
Dim newColor As Long

    txtDocument.Text = ""
    txtJeu1.Text = ""
    txtJeu2.Text = ""
    labNum.Caption = ""
    labCle.Caption = ""
    labCle.Visible = False
    btnModifier.Enabled = False
    btnEnlever.Enabled = False
    fgUTI_DOC.Clear
    fgUTI_DOC.ColWidth(4) = 0
    fgUTI_DOC.SelectionMode = flexSelectionFree
    fgUTI_DOC.AllowBigSelection = False
    With frmSAB_Dossier_RDE.fgUTI_DOC
        fgUTI_DOC.Rows = .Rows
        For yRow = 0 To .Rows - 1
            For xCol = 0 To .Cols - 1
                fgUTI_DOC.Row = yRow
                .Row = yRow
                fgUTI_DOC.Col = xCol
                .Col = xCol
                If .Col = 0 Then newColor = .CellBackColor
                fgUTI_DOC.Text = .Text
                fgUTI_DOC.CellBackColor = newColor
            Next xCol
        Next yRow
    End With
    

End Sub


Private Sub txtDocument_Change()

    btnModifier.Enabled = True
    btnValider.Enabled = False
    
End Sub


Private Sub txtDocument_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Call etat_normal
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        If btnModifier.Enabled Then
            Call btnModifier_Click
            KeyAscii = 0
        Else
            Call etat_normal
            KeyAscii = 0
        End If
    End If
    
End Sub

Private Sub txtJeu1_Change()

    btnModifier.Enabled = True
    btnValider.Enabled = False

End Sub


Private Sub txtJeu1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Call etat_normal
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        If btnModifier.Enabled Then
            Call btnModifier_Click
            KeyAscii = 0
        Else
            Call etat_normal
            KeyAscii = 0
        End If
    End If

End Sub


Private Sub txtJeu2_Change()

    btnModifier.Enabled = True
    btnValider.Enabled = False

End Sub


Private Sub txtJeu2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Call etat_normal
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        If btnModifier.Enabled Then
            Call btnModifier_Click
            KeyAscii = 0
        Else
            Call etat_normal
            KeyAscii = 0
        End If
    End If

End Sub


