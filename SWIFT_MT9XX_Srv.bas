Attribute VB_Name = "srvSWIFT_MT9XX"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeMT900
 
      MOUVEMPIE         As Long                           ' NUMERO DE PIECE
      MOUVEMCOM         As String * 20                    ' NUMERO COMPTE
      MOUVEMOPE         As String * 3                     ' CODE OPERATION
      MOUVEMNUM         As Long                           ' NUMERO OPERATION
      COMPTEDEV       As String * 3                     ' TABLES BASE 013
    MOUVEMMON       As Currency                       ' MONTANT
    MOUVEMDVA       As Long                           ' DATE DE VALEUR
    LIBELLIB1       As String * 30                    ' Libellé 1
    LIBELLIB2       As String * 30                    ' Libellé 2
    LIBELLIB3       As String * 30                    ' Libellé 3
    LIBELLIB4       As String * 30                    ' Libellé 4
End Type

Public xMT900 As typeMT900
Dim selMT900() As typeMT900, selMT900_Nb As Long, selMT900_Index As Long, selMT900_Max As Long

Dim xYBIAMVT0 As typeYBIAMVT0

Dim wAMJHMS As String
Dim paramMT900_Loro As String, paramMT900F_Loro As String
Dim wBIC As String, wMOUVEMCOM As String
Dim IbmAmjMin As String, IbmAmjMax As String

Public Sub Swift_MT900_Monitor(lAMJMin As String, lAMJMax As String)
Dim V
Dim xFileName As String, X As String, I As Integer
Dim blnFirst As Boolean
Dim wCompte As String
Dim meYBIAMON0 As typeYBIAMON0

On Error GoTo Error_Handler

meYBIAMON0.MONAPP = "COMPTA"
meYBIAMON0.MONFLUX = "MT900"
meYBIAMON0.MONSTATUS = ""

V = fctExploitation_Auto_Control(meYBIAMON0)
If Not IsNull(V) Then Exit Sub

'If Not blnAuto_Exploitation_Ok("DATE_CPT_J", "@MT900") Then
'    Call MsgBox("MT900 : Traitement déjà effectué au " & YBIATAB0_DATE_CPT_J, vbCritical, paramYBase_DataF & "\@MT900_Exploitation_ok.txt")
'    Exit Sub
'End If
'==============================================================================
ReDim selMT900(501): selMT900_Max = 500: selMT900_Nb = 0

wMOUVEMCOM = "": wBIC = ""


IbmAmjMin = dateIBM(lAMJMin)
IbmAmjMax = dateIBM(lAMJMax)

'$JPL 20111017 Call Swift_MT900_Extract("11009978001")
'$JPL 20111017 Call Swift_MT900_Extract("11009400001")
'$JPL 20111017 Call Swift_MT900_Extract("50222400001")
'$JPL 20111017 Call Swift_MT900_Extract("50222978001")
'$JPL 20111017 Call Swift_MT900_Extract("50529978001")
'$JPL 20111017 Call Swift_MT900_Extract("50529400001")

If selMT900_Nb > 0 Then

    wAMJHMS = DSys & "_" & time_Hms & "_"
       
    paramSAA_Init
    
    paramMT900_Loro = "MT900_Loro_" & lAMJMax & "_SW"
    
    paramMT900F_Loro = paramYBase_DataF & paramMT900_Loro & paramYBase_Data_ExtensionP
    If Dir(paramMT900F_Loro) <> "" Then Kill paramMT900F_Loro
    Call FEU_ROUGE
    Open paramMT900F_Loro For Output As #1
    blnFirst = True
    
    For I = 1 To selMT900_Nb
        xMT900 = selMT900(I)
        If wMOUVEMCOM <> xMT900.MOUVEMCOM Then
            wMOUVEMCOM = xMT900.MOUVEMCOM
            V = rsZADRESS0_BIC_Compte(xMT900.MOUVEMCOM, wBIC)
            If Not IsNull(V) Then MsgBox V, vbCritical, wMOUVEMCOM
        End If
        
        If wBIC <> "" Then
            Swift_MT900_RJE blnFirst
            blnFirst = False
        End If
        
    Next I
    
    
    Close #1
    Call FEU_VERT
'====================================================================================================
'    X = MsgBox("Emettre le fichier des MT900 ?", vbCritical + vbYesNo, "A7-Recette")
'    If X = vbYes Then
'        Call blnAuto_Exploitation_Ok("Update", "@MT900")
        X = paramSAA_DataF_Archive & "\SAA_from_MT900" & wAMJHMS & paramMT900_Loro & ".sav"
        msFileSystem.CopyFile paramMT900F_Loro, X
        
        X = paramSAA_DataF_from_SAB & wAMJHMS & paramMT900_Loro & paramSAA_Data_from_SAB_ExtensionP_rje
        msFileSystem.MoveFile paramMT900F_Loro, X
'    End If
End If

V = fctExploitation_Auto_End(meYBIAMON0)
   
'====================================================================================================


Exit Sub

Error_Handler:

Close
Shell_MsgBox "Swift_MT900_Monitor " & Error, vbCritical, "srvSwift_MT900", True

End Sub




Public Sub Swift_MT900_MOUVEMPIE_Cumul()
Dim I As Integer, blnOk As Boolean

blnOk = False
For I = 1 To selMT900_Nb
    If selMT900(I).MOUVEMPIE = xYBIAMVT0.MOUVEMPIE Then blnOk = True: Exit For
Next I
If Not blnOk Then
    selMT900_Nb = selMT900_Nb + 1
    I = selMT900_Nb
    selMT900(I).MOUVEMPIE = xYBIAMVT0.MOUVEMPIE
    selMT900(I).MOUVEMCOM = xYBIAMVT0.MOUVEMCOM
    selMT900(I).MOUVEMOPE = xYBIAMVT0.MOUVEMOPE
    selMT900(I).MOUVEMNUM = xYBIAMVT0.MOUVEMNUM
    selMT900(I).COMPTEDEV = xYBIAMVT0.COMPTEDEV
    selMT900(I).MOUVEMMON = 0
    selMT900(I).MOUVEMDVA = xYBIAMVT0.MOUVEMDVA
    selMT900(I).LIBELLIB1 = xYBIAMVT0.LIBELLIB1
    selMT900(I).LIBELLIB2 = xYBIAMVT0.LIBELLIB2
    selMT900(I).LIBELLIB3 = xYBIAMVT0.LIBELLIB3
    selMT900(I).LIBELLIB4 = xYBIAMVT0.LIBELLIB4
End If

selMT900(I).MOUVEMMON = selMT900(I).MOUVEMMON + xYBIAMVT0.MOUVEMMON
End Sub
Public Function Swift_MT900_RJE(blnFirst As Boolean)
'=========================================================================
Dim V
Dim X As String, X20 As String
V = Null

Dim I As Integer

Dim wBlock1 As String, wBlock2 As String
On Error GoTo Error_Handler
Set rsAdo = Nothing

wBlock1 = "{1:F01" & paramBic8 & "AXXX0000000000}"
wBlock2 = "{2:I900" & wBIC & "XN}"

If blnFirst Then
    Print #1, wBlock1 & wBlock2 & "{4:"
Else
    Print #1, "$" & wBlock1 & wBlock2 & "{4:"
End If

X = "DCOM" & xMT900.MOUVEMOPE & xMT900.MOUVEMNUM
X20 = SAA_Text_Control(X, 16)
Print #1, ":20:" & X20

X = SAA_Text_Control(Trim(xMT900.LIBELLIB1), 16)
If X = "" Then X = X20
Print #1, ":21:" & X

Print #1, ":25:" & Trim(xMT900.MOUVEMCOM)
X = Format$(xMT900.MOUVEMDVA, "0000000")

Print #1, ":32A:" & Mid$(X, 2, 6) & xMT900.COMPTEDEV & cur_AbsV(xMT900.MOUVEMMON)

X = SAA_Text_Control(Trim(xMT900.LIBELLIB1), 35)
If X <> "" Then Print #1, ":72:" & X
X = SAA_Text_Control(Trim(xMT900.LIBELLIB2), 35)
If X <> "" Then Print #1, X
X = SAA_Text_Control(Trim(xMT900.LIBELLIB3), 35)
If X <> "" Then Print #1, X
X = SAA_Text_Control(Trim(xMT900.LIBELLIB4), 35)
If X <> "" Then Print #1, X




Print #1, "-}"


GoTo Exit_Function
'=============================================================
Error_Handler:
    V = Error
Error_MsgBox:
    'MsgBox V, vbCritical, Me.Name & " : cmdYSWIALI0_Update_Transaction_Historique"
    
Exit_Function:
    On Error Resume Next
    Swift_MT900_RJE = V

End Function



Public Sub Swift_MT900_Extract(lMOUVEMCOM As String)
Dim xSQL As String
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH" _
     & " where MOUVEMCOM = '" & lMOUVEMCOM & "'" _
     & " and MOUVEMDTR >= " & IbmAmjMin _
     & " and MOUVEMDTR <= " & IbmAmjMax _
     & " order by MOUVEMDTR, MOUVEMPIE, MOUVEMECR"
     
Set rsSab = Nothing
Set rsSab = cnsab.Execute(xSQL)
Do Until rsSab.EOF

    If rsSab("MOUVEMMON") > 0 Then
        Call rsYBIAMVT0_GetBuffer(rsSab, xYBIAMVT0)
        Swift_MT900_MOUVEMPIE_Cumul
    End If
    rsSab.MoveNext
Loop
                     

End Sub
