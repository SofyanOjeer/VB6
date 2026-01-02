Attribute VB_Name = "rsZDWHEXP0"
'---------------------------------------------------------
Option Explicit
Type typeZDWHEXP0

  
     DWHEXPDTA       As Long
     DWHEXPETA       As Long
     DWHEXPAGE       As Long
     DWHEXPSER       As String * 2
     DWHEXPSSE       As String * 2
     DWHEXPPLA       As Long
     DWHEXPOPE       As String * 6
     DWHEXPNAT       As String * 10
     DWHEXPNUM       As String * 20
     DWHEXPTYP       As String * 1
     DWHEXPCOM       As String * 1
     DWHEXPDEV       As String * 3
     DWHEXPFIN       As Long
     DWHEXPDUI       As Long
     DWHEXPDUR       As Long
     DWHEXPTYO       As String * 1
     DWHEXPCLI       As String * 7
     DWHEXPTAU       As Long
     DWHEXPENC       As Currency
     DWHEXPINT       As Currency
     DWHEXPIMP       As Currency
     DWHEXPEXB       As Currency
     DWHEXPPRO       As Currency
     DWHEXPEXN       As Currency
     DWHEXPCAT       As String * 6
     DWHEXPREG       As String * 1
     DWHEXPTXP       As Long
     DWHEXPEXA       As Currency
     DWHEXPEAP       As Currency
     DWHEXPEXS       As Currency
     DWHEXPESP       As Currency
     DWHEXPEXR       As Currency
     DWHEXPFIL       As String * 100


End Type

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZDWHEXP0_GetBuffer(rsSab As ADODB.Recordset, rsZDWHEXP0 As typeZDWHEXP0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZDWHEXP0_GetBuffer = Null


rsZDWHEXP0.DWHEXPDTA = rsSab("DWHEXPDTA")
rsZDWHEXP0.DWHEXPETA = rsSab("DWHEXPETA")
rsZDWHEXP0.DWHEXPAGE = rsSab("DWHEXPAGE")
rsZDWHEXP0.DWHEXPSER = rsSab("DWHEXPSER")
rsZDWHEXP0.DWHEXPSSE = rsSab("DWHEXPSSE")
rsZDWHEXP0.DWHEXPPLA = rsSab("DWHEXPPLA")
rsZDWHEXP0.DWHEXPOPE = rsSab("DWHEXPOPE")
rsZDWHEXP0.DWHEXPNAT = rsSab("DWHEXPNAT")
rsZDWHEXP0.DWHEXPNUM = rsSab("DWHEXPNUM")
rsZDWHEXP0.DWHEXPTYP = rsSab("DWHEXPTYP")
rsZDWHEXP0.DWHEXPCOM = rsSab("DWHEXPCOM")
rsZDWHEXP0.DWHEXPDEV = rsSab("DWHEXPDEV")
rsZDWHEXP0.DWHEXPFIN = rsSab("DWHEXPFIN")
rsZDWHEXP0.DWHEXPDUI = rsSab("DWHEXPDUI")
rsZDWHEXP0.DWHEXPDUR = rsSab("DWHEXPDUR")
rsZDWHEXP0.DWHEXPTYO = rsSab("DWHEXPTYO")
rsZDWHEXP0.DWHEXPCLI = rsSab("DWHEXPCLI")
rsZDWHEXP0.DWHEXPTAU = rsSab("DWHEXPTAU")
rsZDWHEXP0.DWHEXPENC = rsSab("DWHEXPENC")
rsZDWHEXP0.DWHEXPINT = rsSab("DWHEXPINT")
rsZDWHEXP0.DWHEXPIMP = rsSab("DWHEXPIMP")
rsZDWHEXP0.DWHEXPEXB = rsSab("DWHEXPEXB")
rsZDWHEXP0.DWHEXPPRO = rsSab("DWHEXPPRO")
rsZDWHEXP0.DWHEXPEXN = rsSab("DWHEXPEXN")
rsZDWHEXP0.DWHEXPCAT = rsSab("DWHEXPCAT")
rsZDWHEXP0.DWHEXPREG = rsSab("DWHEXPREG")
rsZDWHEXP0.DWHEXPTXP = rsSab("DWHEXPTXP")
rsZDWHEXP0.DWHEXPEXA = rsSab("DWHEXPEXA")
rsZDWHEXP0.DWHEXPEAP = rsSab("DWHEXPEAP")
rsZDWHEXP0.DWHEXPEXS = rsSab("DWHEXPEXS")
rsZDWHEXP0.DWHEXPESP = rsSab("DWHEXPESP")
rsZDWHEXP0.DWHEXPEXR = rsSab("DWHEXPEXR")
rsZDWHEXP0.DWHEXPFIL = rsSab("DWHEXPFIL")


Exit Function

Error_Handler:

rsZDWHEXP0_GetBuffer = Error

End Function

'---------------------------------------------------------
Public Sub rsZDWHEXP0_Init(rsZDWHEXP0 As typeZDWHEXP0)
'---------------------------------------------------------

rsZDWHEXP0.DWHEXPDTA = 0
rsZDWHEXP0.DWHEXPETA = 0
rsZDWHEXP0.DWHEXPAGE = 0
rsZDWHEXP0.DWHEXPSER = ""
rsZDWHEXP0.DWHEXPSSE = ""
rsZDWHEXP0.DWHEXPPLA = 0
rsZDWHEXP0.DWHEXPOPE = ""
rsZDWHEXP0.DWHEXPNAT = ""
rsZDWHEXP0.DWHEXPNUM = ""
rsZDWHEXP0.DWHEXPTYP = ""
rsZDWHEXP0.DWHEXPCOM = ""
rsZDWHEXP0.DWHEXPDEV = ""
rsZDWHEXP0.DWHEXPFIN = 0
rsZDWHEXP0.DWHEXPDUI = 0
rsZDWHEXP0.DWHEXPDUR = 0
rsZDWHEXP0.DWHEXPTYO = ""
rsZDWHEXP0.DWHEXPCLI = ""
rsZDWHEXP0.DWHEXPTAU = 0
rsZDWHEXP0.DWHEXPENC = 0
rsZDWHEXP0.DWHEXPINT = 0
rsZDWHEXP0.DWHEXPIMP = 0
rsZDWHEXP0.DWHEXPEXB = 0
rsZDWHEXP0.DWHEXPPRO = 0
rsZDWHEXP0.DWHEXPEXN = 0
rsZDWHEXP0.DWHEXPCAT = ""
rsZDWHEXP0.DWHEXPREG = ""
rsZDWHEXP0.DWHEXPTXP = 0
rsZDWHEXP0.DWHEXPEXA = 0
rsZDWHEXP0.DWHEXPEAP = 0
rsZDWHEXP0.DWHEXPEXS = 0
rsZDWHEXP0.DWHEXPESP = 0
rsZDWHEXP0.DWHEXPEXR = 0
rsZDWHEXP0.DWHEXPFIL = ""



End Sub


'











