Attribute VB_Name = "rsZBALEX10"
'---------------------------------------------------------
Option Explicit
Type typeZBALEX10

  
     BALEX1DTA       As Long
     BALEX1ETA       As Long
     BALEX1AGE       As Long
     BALEX1SER       As String * 2
     BALEX1SSE       As String * 2
     BALEX1PLA       As Long
     BALEX1OPE       As String * 6
     BALEX1NAT       As String * 10
     BALEX1NUM       As String * 20
     BALEX1TYP       As String * 2
     BALEX1CLI       As String * 7
     BALEX1DEV       As String * 3
     BALEX1DBA       As String * 1
     BALEX1FIN       As Long
     BALEX1DUI       As Long
     BALEX1DUR       As Long
     BALEX1CAA       As String * 6
     BALEX1CAT       As String * 6
     BALEX1REG       As String * 1
     BALEX1TXG       As Long
     BALEX1AGG       As Long
     BALEX1SRG       As String * 2
     BALEX1SSG       As String * 2
     BALEX1OPG       As String * 6
     BALEX1NTG       As String * 6
     BALEX1NDG       As Long
     BALEX1ENC       As Currency
     BALEX1PRO       As Currency
     BALEX1EHB       As Currency
     BALEX1EXB       As Currency
     BALEX1EXN       As Currency
     BALEX1MOT       As String * 100
     BALEX1DAJ       As Long
     BALEX1UAJ       As Long
     BALEX1FIL       As String * 50


End Type

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZBALEX10_GetBuffer(rsSab As ADODB.Recordset, rsZBALEX10 As typeZBALEX10)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZBALEX10_GetBuffer = Null


rsZBALEX10.BALEX1DTA = rsSab("BALEX1DTA")
rsZBALEX10.BALEX1ETA = rsSab("BALEX1ETA")
rsZBALEX10.BALEX1AGE = rsSab("BALEX1AGE")
rsZBALEX10.BALEX1SER = rsSab("BALEX1SER")
rsZBALEX10.BALEX1SSE = rsSab("BALEX1SSE")
rsZBALEX10.BALEX1PLA = rsSab("BALEX1PLA")
rsZBALEX10.BALEX1OPE = rsSab("BALEX1OPE")
rsZBALEX10.BALEX1NAT = rsSab("BALEX1NAT")
rsZBALEX10.BALEX1NUM = rsSab("BALEX1NUM")
rsZBALEX10.BALEX1TYP = rsSab("BALEX1TYP")
rsZBALEX10.BALEX1CLI = rsSab("BALEX1CLI")
rsZBALEX10.BALEX1DEV = rsSab("BALEX1DEV")
rsZBALEX10.BALEX1DBA = rsSab("BALEX1DBA")
rsZBALEX10.BALEX1FIN = rsSab("BALEX1FIN")
rsZBALEX10.BALEX1DUI = rsSab("BALEX1DUI")
rsZBALEX10.BALEX1DUR = rsSab("BALEX1DUR")
rsZBALEX10.BALEX1CAA = rsSab("BALEX1CAA")
rsZBALEX10.BALEX1CAT = rsSab("BALEX1CAT")
rsZBALEX10.BALEX1REG = rsSab("BALEX1REG")
rsZBALEX10.BALEX1TXG = rsSab("BALEX1TXG")
rsZBALEX10.BALEX1AGG = rsSab("BALEX1AGG")
rsZBALEX10.BALEX1SRG = rsSab("BALEX1SRG")
rsZBALEX10.BALEX1SSG = rsSab("BALEX1SSG")
rsZBALEX10.BALEX1OPG = rsSab("BALEX1OPG")
rsZBALEX10.BALEX1NTG = rsSab("BALEX1NTG")
rsZBALEX10.BALEX1NDG = rsSab("BALEX1NDG")
rsZBALEX10.BALEX1ENC = rsSab("BALEX1ENC")
rsZBALEX10.BALEX1PRO = rsSab("BALEX1PRO")
rsZBALEX10.BALEX1EHB = rsSab("BALEX1EHB")
rsZBALEX10.BALEX1EXB = rsSab("BALEX1EXB")
rsZBALEX10.BALEX1EXN = rsSab("BALEX1EXN")
rsZBALEX10.BALEX1MOT = rsSab("BALEX1MOT")
rsZBALEX10.BALEX1DAJ = rsSab("BALEX1DAJ")
rsZBALEX10.BALEX1UAJ = rsSab("BALEX1UAJ")
rsZBALEX10.BALEX1FIL = rsSab("BALEX1FIL")
    

Exit Function

Error_Handler:

rsZBALEX10_GetBuffer = Error

End Function

'---------------------------------------------------------
Public Sub rsZBALEX10_Init(rsZBALEX10 As typeZBALEX10)
'---------------------------------------------------------
rsZBALEX10.BALEX1DTA = 0
rsZBALEX10.BALEX1ETA = 0
rsZBALEX10.BALEX1AGE = 0
rsZBALEX10.BALEX1SER = ""
rsZBALEX10.BALEX1SSE = ""
rsZBALEX10.BALEX1PLA = 0
rsZBALEX10.BALEX1OPE = ""
rsZBALEX10.BALEX1NAT = ""
rsZBALEX10.BALEX1NUM = ""
rsZBALEX10.BALEX1TYP = ""
rsZBALEX10.BALEX1CLI = ""
rsZBALEX10.BALEX1DEV = ""
rsZBALEX10.BALEX1DBA = ""
rsZBALEX10.BALEX1FIN = 0
rsZBALEX10.BALEX1DUI = 0
rsZBALEX10.BALEX1DUR = 0
rsZBALEX10.BALEX1CAA = ""
rsZBALEX10.BALEX1CAT = ""
rsZBALEX10.BALEX1REG = ""
rsZBALEX10.BALEX1TXG = 0
rsZBALEX10.BALEX1AGG = 0
rsZBALEX10.BALEX1SRG = ""
rsZBALEX10.BALEX1SSG = ""
rsZBALEX10.BALEX1OPG = ""
rsZBALEX10.BALEX1NTG = ""
rsZBALEX10.BALEX1NDG = 0
rsZBALEX10.BALEX1ENC = 0
rsZBALEX10.BALEX1PRO = 0
rsZBALEX10.BALEX1EHB = 0
rsZBALEX10.BALEX1EXB = 0
rsZBALEX10.BALEX1EXN = 0
rsZBALEX10.BALEX1MOT = ""
rsZBALEX10.BALEX1DAJ = 0
rsZBALEX10.BALEX1UAJ = 0
rsZBALEX10.BALEX1FIL = ""



End Sub


'










