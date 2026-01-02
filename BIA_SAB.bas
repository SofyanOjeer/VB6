Attribute VB_Name = "BIA_SAB"

Option Explicit

'20040830 jpl $$$$$$$$$$$$$$$$$ BIAS820I   ==> BIA_MAIN.bas et BIA_SAB.bas
'---------------------------------------------------------
Public Sub prtSocInit()
'---------------------------------------------------------
prtFormType = "SOC"

frmElpPrt.prtInit

If frmRTF_blnA5 Then prtMaxY = 7700

prtSoc

End Sub

Public Sub fctPCEC_Atribut(lPCEC As String, lDEV As String, blnCptOrdinaire As Boolean, blnRIB As Boolean, blnMédiateur As Boolean)
Dim X5 As String
blnCptOrdinaire = False: blnRIB = False: blnMédiateur = False

X5 = mId$(Trim(lPCEC), 1, 5)
If X5 = "11120" _
Or X5 = "12120" _
Or X5 = "12121" _
Or X5 = "12122" _
Or X5 = "25112" _
Or X5 = "25113" _
Or X5 = "25114" _
Or X5 = "25115" _
Or X5 = "25116" _
Or X5 = "25117" _
                    Then
    blnCptOrdinaire = True:
End If
If X5 = "25111" Then
    blnCptOrdinaire = True: blnMédiateur = True
End If
If blnCptOrdinaire And lDEV = "EUR" Then blnRIB = True

End Sub
Public Function Table_Edition_Form(lEdition_Form As typeEdition_Form) As String
Static meElpTable As typeElpTable, blnInit As Boolean
Static meEdition_Form As typeEdition_Form
If Not blnInit Then
    blnInit = True
    recElpTable_Init meElpTable
    meElpTable.Method = "Seek="
    meElpTable.ID = constEdition_Form
End If
If Trim(meElpTable.K1) <> lEdition_Form.K1 Or Trim(meElpTable.K2) <> lEdition_Form.K2 Then
    meElpTable.K1 = lEdition_Form.K1
    meElpTable.K2 = lEdition_Form.K2
    If tableElpTable_Read(meElpTable) <> 0 Then
        meElpTable.K2 = "?"
        meElpTable.Name = "?Edition_Form"
        Table_Edition_Form_Init meEdition_Form
    Else
        MsgTxt = Space$(34) & meElpTable.Memo
        MsgTxtIndex = 0
        srvEdition_Form_GetBuffer meEdition_Form
    End If
End If
Table_Edition_Form = meElpTable.Name
meEdition_Form.Name = meElpTable.Name
meEdition_Form.K1 = meElpTable.K1
meEdition_Form.K2 = meElpTable.K2
lEdition_Form = meEdition_Form

End Function


Public Function fctUser_Classe_Aut(lClasse As Long) As Boolean

fctUser_Classe_Aut = True
If lClasse > 0 And lClasse < 100 Then
    Select Case mId$(currentYBIAUSR0.MNUUTPCLA, lClasse, 1)
        Case "1", "2"
        Case Else: fctUser_Classe_Aut = False
    End Select
End If

End Function

 



Public Sub mainSoc_YBase_Load()
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
' dans l'ordre !!!! YBIATAB0 , YPLAN0, YSOLDE0
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Dim Nb As Long

Dim wText As String, X As String
mdbYBase.tableYBase_Open
frmElp.lstErr.Clear

If usrId = "BIA_INFO" Then
    frmElp.lstErr.Height = 200
    frmElp.lblMain = "Menu, BIA_INFO : pas d'initialisation"
Else
    frmElp.lstErr.Height = 2000
    frmElp.lstErr.AddItem "Initilialisation....."
    
    srvYBIATAB0_Import wText: frmElp.lstErr.AddItem "YBIATAB0 : " & wText
    srvYBIACPT0_Import wText: frmElp.lstErr.AddItem "YBIACPT0 : " & wText
    srvYPLAN0_Import wText: frmElp.lstErr.AddItem "YPLAN0 : " & wText
    srvYSOLDE0_Import wText: frmElp.lstErr.AddItem "YSOLDE0 : " & wText
    srvYBIAMVT0_Import wText: frmElp.lstErr.AddItem "YBIAMVT0 : " & wText
    
    srvYBIARELH_Import wText: frmElp.lstErr.AddItem "YBIARELH : " & wText
    
    srvYCOMPTE0_Import wText: frmElp.lstErr.AddItem "YCOMPTE0 : " & wText
    
    srvYADRESS0_Import wText: frmElp.lstErr.AddItem "YADRESS0 : " & wText
    srvYTITULA0_Import wText: frmElp.lstErr.AddItem "YTITULA0 : " & wText
    
    srvYAUTE1I0_Import wText: frmElp.lstErr.AddItem "YAUTE1I0 : " & wText
    
    '$$$ à revoir JPL
    srvYCLIENA0_Import_Ybase wText: frmElp.lstErr.AddItem "YCLIENA0 : " & wText

    frmElp.lstErr.Clear
    frmElp.lstErr.AddItem Time & " : Initilialisation Terminée :" & dateImp10(YBIATAB0_DATE_CPT_J)
    frmElp.lstErr.Height = 200
    frmElp.lblMain = "Menu, Soldes au : " & dateImp10(YBIATAB0_DATE_CPT_J)
End If

End Sub
Public Sub prtAdresse(lYADRESS0 As typeYADRESS0, blnPostal As Boolean)
Dim wADRESSPAY As String, blnADRESSPAY As Boolean
Dim wADRESSRA2 As String, blnADRESSRA2 As Boolean
'-----------------------encadrement petit tirets---------
Dim wCurrentX As Integer
wCurrentX = XPrt.CurrentX
XPrt.FontBold = True
XPrt.Print lYADRESS0.ADRESSRA1;

'-----------------------------------------------------
wADRESSRA2 = Trim(lYADRESS0.ADRESSRA2)
If wADRESSRA2 = "" Then
    blnADRESSRA2 = blnPostal
Else
    blnADRESSRA2 = True
End If

XPrt.FontBold = False
If blnADRESSRA2 Then
   XPrt.CurrentY = XPrt.CurrentY + 270
    XPrt.CurrentX = wCurrentX
    XPrt.Print wADRESSRA2;
End If
'-----------------------------------3---------------
If Trim(lYADRESS0.ADRESSAD1) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + 270
    XPrt.CurrentX = wCurrentX
    XPrt.Print lYADRESS0.ADRESSAD1;
End If
'----------------------------------4-------------------
If Trim(lYADRESS0.ADRESSAD2) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + 270
    XPrt.CurrentX = wCurrentX
    XPrt.Print lYADRESS0.ADRESSAD2;
End If

'-----------------------------------5------------------
If Trim(lYADRESS0.ADRESSAD3) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + 270
    XPrt.CurrentX = wCurrentX
    XPrt.Print lYADRESS0.ADRESSAD3;
End If
'------------------------------------6------------------
blnADRESSPAY = False
wADRESSPAY = Trim(lYADRESS0.ADRESSPAY)
If blnPostal Then
    If wADRESSPAY = "" Or wADRESSPAY = "FRANCE" Then
        XPrt.CurrentY = XPrt.CurrentY + 270
    Else
        blnADRESSPAY = True
    End If
End If
If Trim(lYADRESS0.ADRESSCOP) <> "" _
Or Trim(lYADRESS0.ADRESSVIL) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + 270
    XPrt.CurrentX = wCurrentX
    If Trim(lYADRESS0.ADRESSCOP) <> "" Then XPrt.Print lYADRESS0.ADRESSCOP & "  ";
    XPrt.Print lYADRESS0.ADRESSVIL;
End If
'------------------------------------8------------------
If blnADRESSPAY Then
    XPrt.FontBold = True
    XPrt.CurrentY = XPrt.CurrentY + 270
    XPrt.CurrentX = wCurrentX
    XPrt.Print wADRESSPAY;
    XPrt.FontBold = False
End If
'------------------------------------------

End Sub

Public Sub prtAdresse_Enveloppe(lYADRESS0 As typeYADRESS0)
'-----------------------encadrement petits tirets---------

XPrt.Line (5600, 2300)-(5700, 2300)
XPrt.Line (5600, 2300)-(5600, 2400)

XPrt.Line (10900, 2300)-(11000, 2300)
XPrt.Line (11000, 2300)-(11000, 2400)

XPrt.Line (5600, 4300)-(5700, 4300)
XPrt.Line (5600, 4200)-(5600, 4300)

XPrt.Line (10900, 4300)-(11000, 4300)
XPrt.Line (11000, 4200)-(11000, 4300)
XPrt.CurrentY = 2400
XPrt.CurrentX = 5700
prtAdresse lYADRESS0, True
End Sub



'---------------------------------------------------------
Public Sub prtSAB_Compta_Mt(lMonTant As Currency, lcolDb As Integer, lcolCr As Integer)
'---------------------------------------------------------
Dim X As String

XPrt.FontBold = True
X = Format$(Abs(lMonTant), "## ### ### ### ### ##0.00")
XPrt.CurrentX = IIf(lMonTant < 0, lcolCr, lcolDb) - 100 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.FontBold = False

End Sub


