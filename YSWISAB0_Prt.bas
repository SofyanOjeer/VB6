Attribute VB_Name = "prtYSWISAB0"
Option Explicit

Dim cnSIDE_DB As New ADODB.Connection, rsSIDE_DB As New ADODB.Recordset

Dim wFontColor As Long
Dim wFiligrane As String
Dim xYSWISAB0 As typeYSWISAB0
Dim xrMesg As typerMesg, xrIntv As typerIntv, xrAppe As typerAppe
Dim xrText As typerText


Dim Mesg_aid As Long, mesg_s_umidl As Long, mesg_s_umidh As Long

Dim blnrAppe As Boolean, blnACK As Boolean, blnNAK As Boolean, blnSwift_Ctl As Boolean
Dim wUnit As String, wStatut As String

Dim wAppe_large_data As String
Public Sub prtYSWISAB0_Open()
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

'Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
prtOrientation = vbPRORPortrait '
prtPgmName = "prtYSWISAB0"
prtTitleUsr = usrName
prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 300
prtFontName = "Calibri"

prtFormType = ""
frmElpPrt.prtStdInit
XPrt.FontSize = 10

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

Public Sub prtYSWISAB0_Monitor(lSWISABSWID As Long, lMesg_aid As Long, lmesg_s_umidl As Long, lmesg_s_umidh As Long)
Dim blnOk As Boolean, xSql As String, X As String, K As Integer
Dim V
On Error GoTo prtError
cnSIDE_DB.Open paramODBC_DSN_SIDE_DB

blnOk = False
blnACK = False: blnNAK = False
blnSwift_Ctl = False
blnrAppe = False
wAppe_large_data = ""
xrAppe.appe_remote_input_reference = ""
xrAppe.appe_checksum_value = ""


prtTitleText = " Message Swift"
Call rsYSWISAB0_Init(xYSWISAB0)

Mesg_aid = lMesg_aid
mesg_s_umidl = lmesg_s_umidl
mesg_s_umidh = lmesg_s_umidh

If lSWISABSWID > 0 Then
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 where SWISABSWID = " & lSWISABSWID
    Set rsSab = cnsab.Execute(xSql)
    
    If Not rsSab.EOF Then
        blnOk = True
        Call rsYSWISAB0_GetBuffer(rsSab, xYSWISAB0)
        Mesg_aid = xYSWISAB0.SWISABWID1
        mesg_s_umidl = xYSWISAB0.SWISABWIDL
        mesg_s_umidh = xYSWISAB0.SWISABWIDH
        If xYSWISAB0.SWISABOPEN > 0 Then
            prtTitleText = "Dossier : " & xYSWISAB0.SWISABSER & " " & xYSWISAB0.SWISABSSE _
                         & "  " & xYSWISAB0.SWISABOPEC & " " & xYSWISAB0.SWISABOPEN
        End If

    End If
End If
'----------------------------------------------------------------
 xSql = "select * from rMesg " _
     & "where Aid = " & Mesg_aid _
     & " and Mesg_s_umidl = " & mesg_s_umidl _
     & " and Mesg_s_umidh  =  " & mesg_s_umidh
Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
   
If Not rsSIDE_DB.EOF Then
     Call srvrMesg_GetBuffer_ODBC(rsSIDE_DB, xrMesg)
     

     xSql = "select * from rAppe " _
         & "where Aid = " & Mesg_aid _
         & " and Appe_s_umidl = " & mesg_s_umidl _
         & " and Appe_s_umidh  =  " & mesg_s_umidh _
         & " and Appe_inst_num = 0" _
         & " order by appe_date_time , appe_seq_nbr"
         
    Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
    
    Do While Not rsSIDE_DB.EOF
        V = rsSIDE_DB("Appe_large_data")
        If Not blnrAppe Then
                
            If Not IsNull(rsSIDE_DB("appe_remote_input_reference")) Then
                blnrAppe = True
                xrAppe.appe_remote_input_reference = rsSIDE_DB("appe_remote_input_reference")
                xrAppe.appe_checksum_value = rsSIDE_DB("appe_checksum_value")
            End If
        Else
            xrAppe.appe_remote_input_reference = ""
            xrAppe.appe_checksum_value = ""
        End If
        If rsSIDE_DB("appe_crea_mpfn_name") = "_SI_to_SWIFT" Then
            Select Case Trim(rsSIDE_DB("appe_network_delivery_status"))
                Case "DLV_NACKED": blnACK = False: blnNAK = True: wStatut = "NAK"
                                   xrAppe.appe_date_time = rsSIDE_DB("appe_date_time")
                                   If Not IsNull(V) Then wAppe_large_data = V
                Case "DLV_ACKED": blnACK = True: blnNAK = False: wStatut = "ACK"
                                   xrAppe.appe_date_time = rsSIDE_DB("appe_date_time")
                                   If Not IsNull(V) Then wAppe_large_data = V
            End Select
        End If
        
        rsSIDE_DB.MoveNext
     Loop
     
        Select Case Trim(xrMesg.x_inst0_unit_name)
            Case "SOBF", "ORPA": wUnit = "GDMP_"
            Case "SOBI": wUnit = "SOBI_"
            Case "DAFI": wUnit = "DAFI_"
            Case "BOTC": wUnit = "BOTC_"
            Case "DCOM": wUnit = "DCOM_"
            Case Else:
                If InStr(xrMesg.mesg_rel_trn_ref, "DAFI") > 0 Then
                    wUnit = "DAFI_"
                Else
                
                 Select Case Mid$(xrMesg.mesg_type, 1, 1)
                     Case "1", "2": wUnit = "GDMP_"
                     Case "7": wUnit = "SOBI_"
                     Case "3": wUnit = "BOTC_"
                     Case Else: wUnit = "XXXX_"
                End Select
            End If
        End Select

    prtYSWISAB0_Open
    If Mid$(xrMesg.mesg_uumid, 1, 1) = "I" Then
         'prtForeColor_Header = RGB(0, 96, 0)
         wFontColor = RGB(0, 96, 0)
         'wFiligrane = "W:\Loulergue W\Filigrane_Swift\" & wUnit & wStatut & ".jpg"
         wFiligrane = paramEditionFiligrane_Folder & "Filigrane_Swift\" & wUnit & wStatut & ".jpg"
        Call frmElpPrt.prtFiligrane(wFiligrane)
        
        prtYSWISAB0_Header_Sortant
        prtYSWISAB0_Message_Text
        prtYSWISAB0_Trailer_Sortant
   
     Else
         wStatut = "SWIFT"
         xSql = "select * from rIntv " _
             & "where Aid = " & Mesg_aid _
             & " and Intv_s_umidl = " & mesg_s_umidl _
             & " and Intv_s_umidh  =  " & mesg_s_umidh _
             & " and Intv_inst_num = 0" _
             & " and intv_mpfn_name = 'OFCS_Detect'" _
             & " order by Intv_date_time , Intv_seq_nbr"
             
        Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
    
        Do While Not rsSIDE_DB.EOF
            V = rsSIDE_DB("intv_merged_text")
            If Not IsNull(V) Then
                If InStr(V, "AutoRcvPbOFAC") > 0 Then
                    wStatut = "SWIFT_CTL"
                    Exit Do
                End If
            End If
           rsSIDE_DB.MoveNext
        Loop

         'prtForeColor_Header = RGB(0, 0, 255)
         wFontColor = RGB(0, 0, 255)
         'wFiligrane = "W:\Loulergue W\Filigrane_Swift\" & wUnit & wStatut & ".jpg"
        wFiligrane = paramEditionFiligrane_Folder & "Filigrane_Swift\" & wUnit & wStatut & ".jpg"
        Call frmElpPrt.prtFiligrane(wFiligrane)
        prtYSWISAB0_Header_Entrant
        prtYSWISAB0_Message_Text
        prtYSWISAB0_Trailer_Entrant
     End If
     
     prtYSWISAB0_Close
    
End If


cnSIDE_DB.Close
Set cnSIDE_DB = Nothing
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------
cnSIDE_DB.Close
Set cnSIDE_DB = Nothing

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub


Public Sub prtYSWISAB0_NewLine()
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + 300 > prtMaxY Then
    frmElpPrt.prtNewPage
    'prtYSWISAB0_Form
End If

End Sub

Public Sub prtYSWISAB0_Close()
Dim X As String
On Error GoTo prtError

frmElpPrt.prtEndDoc 1000
frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub





Public Sub prtYSWISAB0_Message_Text()
Dim xValue As String, iLen As Integer
Dim xText_Data_Block As String, xField_Code As String
Dim K As Integer, K2 As Integer, iAsc13 As Integer
Dim xSql As String, X As String, iVal As Integer
Dim xField As String
Dim V

prtYSWISAB0_NewLine
XPrt.FontBold = True
XPrt.ForeColor = vbBlack
Call frmElpPrt.prtCentré(prtMedX, "------------------- Message Text -------------------")
XPrt.FontBold = False

xSql = "select * from rtextField " _
    & "where Aid = " & Mesg_aid _
    & " and text_s_umidl = " & mesg_s_umidl _
    & " and text_s_umidh  =  " & mesg_s_umidh _
    & " order by field_cnt"

Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
If Not rsSIDE_DB.EOF Then
    Do While Not rsSIDE_DB.EOF
        
        Select Case rsSIDE_DB("field_code")
            Case "45", "46", "47", "77":
                V = rsSIDE_DB("value_memo")
                If IsNull(V) Then V = rsSIDE_DB("value")
            Case Else:
                    V = rsSIDE_DB("value")
        End Select
        If IsNull(V) Then
            xValue = ""
        Else
            xValue = V
        End If

        xField = rsSIDE_DB("field_code") & rsSIDE_DB("field_option")
        prtYSWISAB0_NewLine
        XPrt.CurrentX = prtMinX + 100
        XPrt.ForeColor = vbBlack
        XPrt.Print xField & " " & arrMT_Fields_Scan(xField);
        If xField = "32A" Then
            XPrt.ForeColor = vbBlack
            XPrt.CurrentX = prtMinX + 500: XPrt.Print "Date";
            XPrt.CurrentX = prtMinX + 1300: XPrt.Print ": ";
            XPrt.ForeColor = wFontColor
            XPrt.Print xrMesg.x_fin_value_date;
            prtYSWISAB0_NewLine
            XPrt.ForeColor = vbBlack
            XPrt.CurrentX = prtMinX + 500: XPrt.Print "Currency";
            XPrt.CurrentX = prtMinX + 1300: XPrt.Print ": ";
            XPrt.ForeColor = wFontColor
            XPrt.Print xrMesg.x_fin_ccy;
            prtYSWISAB0_NewLine
            XPrt.ForeColor = vbBlack
            XPrt.CurrentX = prtMinX + 500: XPrt.Print "Amount";
            XPrt.CurrentX = prtMinX + 1300: XPrt.Print ": ";
            XPrt.ForeColor = wFontColor
            XPrt.Print Format(xrMesg.x_fin_amount, "### ### ### ##0.00")
        Else
       
            XPrt.ForeColor = wFontColor
    
            iLen = Len(xValue)
            K = 1
            Do
               iAsc13 = InStr(K, xValue, Asc13)
               If iAsc13 > 0 Then
                   prtYSWISAB0_NewLine
                   XPrt.CurrentX = prtMinX + 500
                   XPrt.ForeColor = wFontColor
                   XPrt.CurrentX = prtMinX + 500
                   XPrt.Print Trim(Mid$(xValue, K, iAsc13 - K));
                   K = iAsc13 + 2
               End If
            Loop Until iAsc13 = 0
            
            prtYSWISAB0_NewLine
            XPrt.CurrentX = prtMinX + 500
            XPrt.ForeColor = wFontColor
            XPrt.Print Trim(Mid$(xValue, K, iLen - K + 1))
        End If
        
        rsSIDE_DB.MoveNext
    Loop
Else

    xSql = "select * from rtext " _
        & "where Aid = " & Mesg_aid _
        & " and text_s_umidl = " & mesg_s_umidl _
        & " and text_s_umidh  =  " & mesg_s_umidh
    Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
    If Not rsSIDE_DB.EOF Then
        Call srvrText_GetBuffer_ODBC(rsSIDE_DB, xrText)
        
        xValue = xrText.text_data_block & Asc13
        iLen = Len(xValue)
        If Mid$(xValue, 1, 3) = Asc13 & Asc10 & ": " Then
            K = 3
        Else
            K = 1
        End If
        Do
            iAsc13 = InStr(K, xValue, Asc13)
            If iAsc13 > 0 Then
                X = Trim(Mid$(xValue, K, iAsc13 - K))
                If Mid$(X, 1, 1) <> ":" Then
                    prtYSWISAB0_NewLine
                    XPrt.ForeColor = wFontColor
                    XPrt.CurrentX = prtMinX + 500
                    XPrt.Print Trim(Mid$(xValue, K, iAsc13 - K));
                Else
                    K2 = InStr(2, X, ":")
                    If K2 > 0 Then
                        prtYSWISAB0_NewLine
                        XPrt.ForeColor = vbBlack
                        XPrt.CurrentX = prtMinX + 100
                        xField = Mid$(X, 2, K2 - 2)
                        'iVal = Val(Mid$(x, 2, 2))
                        XPrt.Print Trim(Mid$(X, 2, K2 - 1)) & " " & arrMT_Fields_Scan(xField);
                        prtYSWISAB0_NewLine
                        XPrt.ForeColor = wFontColor
                        If xField = "32A" Then
                            XPrt.ForeColor = vbBlack
                            XPrt.CurrentX = prtMinX + 500: XPrt.Print "Date";
                            XPrt.CurrentX = prtMinX + 1300: XPrt.Print ": ";
                            XPrt.ForeColor = wFontColor
                            XPrt.FontBold = True
                            XPrt.Print xrMesg.x_fin_value_date;
                            XPrt.FontBold = False
                            prtYSWISAB0_NewLine
                            XPrt.ForeColor = vbBlack
                            XPrt.CurrentX = prtMinX + 500: XPrt.Print "Currency";
                            XPrt.CurrentX = prtMinX + 1300: XPrt.Print ": ";
                            XPrt.ForeColor = wFontColor
                            XPrt.FontBold = True
                            XPrt.Print xrMesg.x_fin_ccy;
                            XPrt.FontBold = False
                            prtYSWISAB0_NewLine
                            XPrt.ForeColor = vbBlack
                            XPrt.CurrentX = prtMinX + 500: XPrt.Print "Amount";
                            XPrt.CurrentX = prtMinX + 1300: XPrt.Print ": ";
                            XPrt.ForeColor = wFontColor
                            XPrt.FontBold = True
                            XPrt.Print Format(xrMesg.x_fin_amount, "### ### ### ##0.00")
                            XPrt.FontBold = False
                      Else

                            XPrt.ForeColor = wFontColor
                            XPrt.CurrentX = prtMinX + 500
                            X = Trim(Mid$(X, K2 + 1, Len(X) - K2))
                            XPrt.Print X;
                            Select Case xField
                                Case "52A", "57A", "58A", "59A", "59F": Call prtYSWISAB0_ZSWIBIC0(X, 500)
                            End Select
                        End If
                    Else
                        prtYSWISAB0_NewLine
                        
                        XPrt.ForeColor = wFontColor
                        XPrt.CurrentX = prtMinX + 500
                        X = Trim(Mid$(xValue, K, iAsc13 - K))
                        XPrt.Print X;
                    End If
                End If
                
                K = iAsc13 + 2
            End If
         Loop Until iAsc13 = 0
    End If
        
End If


End Sub
Public Sub prtYSWISAB0_Header_Entrant()

prtYSWISAB0_NewLine
XPrt.FontBold = True
XPrt.ForeColor = vbBlack
Call frmElpPrt.prtCentré(prtMedX, "------------------- Instance TYPE and Transmission -------------------")
XPrt.FontBold = False
prtYSWISAB0_NewLine

XPrt.CurrentX = prtMinX + 100
XPrt.ForeColor = vbBlack
XPrt.Print "Original received from SWIFT";
XPrt.CurrentX = prtMinX + 2800: XPrt.Print ": ";
XPrt.ForeColor = wFontColor
XPrt.Print xrMesg.mesg_crea_date_time;

prtYSWISAB0_NewLine
XPrt.CurrentX = prtMinX + 100
XPrt.ForeColor = vbBlack
XPrt.Print "Priority";
XPrt.CurrentX = prtMinX + 1300: XPrt.Print ": ";
XPrt.ForeColor = wFontColor
XPrt.Print xrMesg.mesg_network_priority;


prtYSWISAB0_NewLine
XPrt.CurrentX = prtMinX + 100
XPrt.ForeColor = vbBlack
XPrt.Print "Message Output Reference";
XPrt.CurrentX = prtMinX + 2800: XPrt.Print ": ";
XPrt.ForeColor = wFontColor
XPrt.Print "";

prtYSWISAB0_NewLine
XPrt.CurrentX = prtMinX + 100
XPrt.ForeColor = vbBlack
XPrt.Print "Correspondent Input Reference";
XPrt.CurrentX = prtMinX + 2800: XPrt.Print ": ";
XPrt.ForeColor = wFontColor
XPrt.Print xrAppe.appe_remote_input_reference;

prtYSWISAB0_NewLine
XPrt.CurrentX = prtMinX + 100
XPrt.FontBold = True
XPrt.ForeColor = vbBlack
Call frmElpPrt.prtCentré(prtMedX, "------------------- Message Header -------------------")
XPrt.FontBold = False

prtYSWISAB0_NewLine
XPrt.CurrentX = prtMinX + 100
XPrt.ForeColor = vbBlack
XPrt.Print "Swift Output";
XPrt.CurrentX = prtMinX + 2800: XPrt.Print ": ";
XPrt.ForeColor = wFontColor
XPrt.FontBold = True
XPrt.Print xrMesg.mesg_identifier & "  " & arrMT_Type_Scan(xrMesg.mesg_type);
XPrt.FontBold = False
Call arrMT_Fields_Load(xrMesg.mesg_type)
prtYSWISAB0_NewLine
XPrt.CurrentX = prtMinX + 100
XPrt.ForeColor = vbBlack
XPrt.Print "Sender";
XPrt.CurrentX = prtMinX + 1000: XPrt.Print ": ";
XPrt.ForeColor = wFontColor
XPrt.FontBold = True
XPrt.Print xrMesg.mesg_sender_X1;
XPrt.FontBold = False

Call prtYSWISAB0_ZSWIBIC0(xrMesg.mesg_sender_X1, 1000)

prtYSWISAB0_NewLine
XPrt.CurrentX = prtMinX + 100
XPrt.ForeColor = vbBlack
XPrt.Print "Receiver";
XPrt.CurrentX = prtMinX + 1000: XPrt.Print ": ";
XPrt.ForeColor = wFontColor
XPrt.FontBold = True
XPrt.Print xrMesg.x_receiver_X1;
XPrt.FontBold = False

Call prtYSWISAB0_ZSWIBIC0(xrMesg.x_receiver_X1, 1000)

prtYSWISAB0_NewLine
XPrt.CurrentX = prtMinX + 100
XPrt.ForeColor = vbBlack
XPrt.Print "MUR";
XPrt.CurrentX = prtMinX + 1000: XPrt.Print ": ";
XPrt.ForeColor = wFontColor
XPrt.Print xrMesg.mesg_user_reference_text;


End Sub

Public Sub prtYSWISAB0_Header_Sortant()

prtYSWISAB0_NewLine
XPrt.FontBold = True
XPrt.ForeColor = vbBlack
Call frmElpPrt.prtCentré(prtMedX, "------------------- Instance TYPE and Transmission -------------------")
XPrt.FontBold = False
prtYSWISAB0_NewLine

XPrt.CurrentX = prtMinX + 100
XPrt.ForeColor = vbBlack
XPrt.Print "Notification (Transmission) of Original sent to SWIFT (" & wStatut & ")";

prtYSWISAB0_NewLine
XPrt.CurrentX = prtMinX + 100
XPrt.ForeColor = vbBlack
XPrt.Print "Network Delivery Status";
XPrt.CurrentX = prtMinX + 2800: XPrt.Print ": ";
XPrt.ForeColor = wFontColor
XPrt.Print xrAppe.appe_date_time;

prtYSWISAB0_NewLine
XPrt.CurrentX = prtMinX + 100
XPrt.ForeColor = vbBlack
XPrt.Print "Priority/Delivery";
XPrt.CurrentX = prtMinX + 2800: XPrt.Print ": ";
XPrt.ForeColor = wFontColor
XPrt.Print xrMesg.mesg_network_priority;


prtYSWISAB0_NewLine
XPrt.CurrentX = prtMinX + 100
XPrt.ForeColor = vbBlack
XPrt.Print "Message Input Reference";
XPrt.CurrentX = prtMinX + 2800: XPrt.Print ": ";
XPrt.ForeColor = wFontColor
XPrt.Print "";

prtYSWISAB0_NewLine
XPrt.CurrentX = prtMinX + 100
XPrt.FontBold = True
XPrt.ForeColor = vbBlack
Call frmElpPrt.prtCentré(prtMedX, "------------------- Message Header -------------------")
XPrt.FontBold = False

prtYSWISAB0_NewLine
XPrt.CurrentX = prtMinX + 100
XPrt.ForeColor = vbBlack
XPrt.Print "Swift Input";
XPrt.CurrentX = prtMinX + 2800: XPrt.Print ": ";
XPrt.ForeColor = wFontColor
XPrt.FontBold = True
XPrt.Print xrMesg.mesg_identifier & "  " & arrMT_Type_Scan(xrMesg.mesg_type);
XPrt.FontBold = False

prtYSWISAB0_NewLine
XPrt.CurrentX = prtMinX + 100
XPrt.ForeColor = vbBlack
XPrt.Print "Sender";
XPrt.CurrentX = prtMinX + 1000: XPrt.Print ": ";
XPrt.ForeColor = wFontColor
XPrt.FontBold = True
XPrt.Print xrMesg.mesg_sender_X1;
XPrt.FontBold = False
Call prtYSWISAB0_ZSWIBIC0(xrMesg.mesg_sender_X1, 1000)


prtYSWISAB0_NewLine
XPrt.CurrentX = prtMinX + 100
XPrt.ForeColor = vbBlack
XPrt.Print "Receiver";
XPrt.CurrentX = prtMinX + 1000: XPrt.Print ": ";
XPrt.ForeColor = wFontColor
XPrt.FontBold = True
XPrt.Print xrMesg.x_receiver_X1;
XPrt.FontBold = False
Call prtYSWISAB0_ZSWIBIC0(xrMesg.x_receiver_X1, 1000)

End Sub




Public Sub prtYSWISAB0_Trailer_Entrant()

prtYSWISAB0_NewLine
XPrt.FontBold = True
XPrt.ForeColor = vbBlack
Call frmElpPrt.prtCentré(prtMedX, "------------------- Message Trailer -------------------")
XPrt.FontBold = False

prtYSWISAB0_NewLine
XPrt.CurrentX = prtMinX + 100
XPrt.ForeColor = vbBlack
XPrt.Print "{CHK:";
XPrt.ForeColor = wFontColor
XPrt.Print xrAppe.appe_checksum_value;
XPrt.Print "}";

prtYSWISAB0_NewLine
XPrt.CurrentX = prtMinX + 100
XPrt.ForeColor = vbBlack
XPrt.Print "PKI Signature: MAC-Equivalent";

prtYSWISAB0_NewLine
XPrt.CurrentX = prtMinX + 100
XPrt.ForeColor = vbBlack
XPrt.FontBold = True
XPrt.Print "End of Message";

End Sub

Public Sub prtYSWISAB0_Trailer_Sortant()
Dim K As Integer
prtYSWISAB0_NewLine
XPrt.FontBold = True
XPrt.ForeColor = vbBlack
Call frmElpPrt.prtCentré(prtMedX, "------------------- Message Trailer -------------------")
XPrt.FontBold = False

prtYSWISAB0_NewLine
XPrt.CurrentX = prtMinX + 100
XPrt.ForeColor = vbBlack
XPrt.Print "{CHK:";
XPrt.ForeColor = wFontColor
XPrt.Print xrAppe.appe_checksum_value;
XPrt.Print "}";

prtYSWISAB0_NewLine
XPrt.CurrentX = prtMinX + 100
XPrt.ForeColor = vbBlack
XPrt.Print "PKI Signature: MAC-Equivalent";


prtYSWISAB0_NewLine
XPrt.FontBold = True
XPrt.ForeColor = vbBlack
Call frmElpPrt.prtCentré(prtMedX, "------------------- Interventions -------------------")
XPrt.FontBold = False

prtYSWISAB0_NewLine
XPrt.CurrentX = prtMinX + 100
XPrt.ForeColor = vbBlack
XPrt.Print "Text";
prtYSWISAB0_NewLine
XPrt.ForeColor = wFontColor
K = InStr(wAppe_large_data, "}}")
If K > 0 Then XPrt.Print Mid$(wAppe_large_data, 1, K + 1);

prtYSWISAB0_NewLine
XPrt.CurrentX = prtMinX + 100
XPrt.ForeColor = vbBlack
XPrt.FontBold = True
XPrt.Print "End of Message";

End Sub

Public Sub prtYSWISAB0_ZSWIBIC0(lBIC As String, lMinx As Integer)
Dim xSql As String
xSql = "select * from " & paramIBM_Library_SAB & ".ZSWIBIC0 where SWIBICBIC = '" & lBIC & "'"
Set rsSab = cnsab.Execute(xSql)
    
If Not rsSab.EOF Then
    XPrt.ForeColor = wFontColor
    prtYSWISAB0_NewLine
    XPrt.CurrentX = prtMinX + lMinx
    XPrt.Print rsSab("SWIBICIN1");
    prtYSWISAB0_NewLine
    XPrt.CurrentX = prtMinX + lMinx
    XPrt.Print rsSab("SWIBICVIL");
End If


End Sub


