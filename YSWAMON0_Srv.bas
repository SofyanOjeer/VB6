Attribute VB_Name = "srvYSWAMON0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsADO As ADODB.Recordset
 
Type typeYSWAMON0
    SWAMONETB   As Integer
    SWAMONAGE   As Integer
    SWAMONSER   As String
    SWAMONSES   As String
    SWAMONOPR   As String
    SWAMONNUM   As Long
    SWAMONNAT   As String
    SWAMONHISV  As Long
    SWAMONSTAK  As String
    SWAMONMTK   As String
    SWAMON22A   As String
    SWAMONZSWI  As Long
    
    SWAMONTXT   As String
    SWAMONYUSR  As String
    SWAMONYAMJ  As Long
    SWAMONYHMS  As Long
    SWAMONYVER  As Long

End Type
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsYSWAMON0_GetBuffer(rsADO As ADODB.Recordset, rsYSWAMON0 As typeYSWAMON0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYSWAMON0_GetBuffer = Null

rsYSWAMON0.SWAMONETB = rsADO("SWAMONETB")
rsYSWAMON0.SWAMONAGE = rsADO("SWAMONAGE")
rsYSWAMON0.SWAMONSER = rsADO("SWAMONSER")
rsYSWAMON0.SWAMONSES = rsADO("SWAMONSES")
rsYSWAMON0.SWAMONOPR = rsADO("SWAMONOPR")
rsYSWAMON0.SWAMONNUM = rsADO("SWAMONNUM")
rsYSWAMON0.SWAMONNAT = rsADO("SWAMONNAT")
rsYSWAMON0.SWAMONHISV = rsADO("SWAMONHISV")
rsYSWAMON0.SWAMONSTAK = rsADO("SWAMONSTAK")
rsYSWAMON0.SWAMONMTK = rsADO("SWAMONMTK")
rsYSWAMON0.SWAMON22A = rsADO("SWAMON22A")
rsYSWAMON0.SWAMONZSWI = rsADO("SWAMONZSWI")
rsYSWAMON0.SWAMONTXT = rsADO("SWAMONTXT")
rsYSWAMON0.SWAMONYUSR = rsADO("SWAMONYUSR")
rsYSWAMON0.SWAMONYAMJ = rsADO("SWAMON52A")
rsYSWAMON0.SWAMONYHMS = rsADO("SWAMONYHMS")
rsYSWAMON0.SWAMONYVER = rsADO("SWAMONYVER")

Exit Function

Error_Handler:

rsYSWAMON0_GetBuffer = Error

End Function









Public Sub rsYSWAMON0_Init(lYSWAMON0 As typeYSWAMON0)
lYSWAMON0.SWAMONETB = 0
lYSWAMON0.SWAMONAGE = ""
lYSWAMON0.SWAMONSER = ""
lYSWAMON0.SWAMONSES = ""
lYSWAMON0.SWAMONOPR = ""
lYSWAMON0.SWAMONNUM = 0
lYSWAMON0.SWAMONHISV = 0
lYSWAMON0.SWAMONNAT = ""
lYSWAMON0.SWAMONZSWI = 0
lYSWAMON0.SWAMONSTAK = ""
lYSWAMON0.SWAMONMTK = ""
lYSWAMON0.SWAMON22A = ""
lYSWAMON0.SWAMONTXT = ""
lYSWAMON0.SWAMONYUSR = ""
lYSWAMON0.SWAMONYAMJ = 0
lYSWAMON0.SWAMONYHMS = 0
lYSWAMON0.SWAMONYVER = 0
End Sub



