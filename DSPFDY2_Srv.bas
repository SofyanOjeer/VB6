Attribute VB_Name = "srvDSPFDY2"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recDSPFDY2Len = 246 ' 34 + 212
Public Const recDSPFDY2_Block = 30

Type typeDSPFDY2
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    MBRCEN          As String * 1                     ' Retrieval century:  0=19xx, 1=2
    MBRDAT          As String * 6                     ' Retrieval date:  year/month/day
    MBRTIM          As String * 6                     ' Retrieval time:  hour/minute/se
    MBFILE          As String * 10                    ' File
    MBLIB           As String * 10                    ' Library
    MBFTYP          As String * 1                     ' P=PF, L=LF, R=DDM PF, S=DDM LF
    MBFILA          As String * 4                     ' File attribute: *PHY or *LGL
    MBMXD           As String * 3                     ' Reserved
    MBFATR          As String * 6                     ' File attribute:  PF, LF, PF38,
    MBSYSN          As String * 8                     ' System Name (Source System, if
    MBASP           As Long                           ' Auxiliary storage pool ID:  1=S
    MBRES           As String * 4                     ' Reserved
    MBDTAT          As String * 1                     ' File type:  D=*DATA, S=*SRC
    MBWAIT          As Long                           ' Maximum file wait time: -1=*IMM
    MBWATR          As Long                           ' Maximum record wait time: -1=*I
    MBSHAR          As String * 1                     ' Reserved
    MBLVLC          As String * 1                     ' Record format level check: N=*N
    MBTXT           As String * 50                    ' Text 'description'
    MBNOFM          As Long                           ' Number of record formats
    MBFCCN          As String * 1                     ' Century created:  0=19xx, 1=20x
    MBFCDT          As String * 6                     ' Date created: year/month/day
    MBFCTM          As String * 6                     ' Time created: hour/minute/secon
    MBFLS           As String * 1                     ' Externally described file:  N=N
    MBICAP          As String * 1                     ' DBCS capable:  N=No, Y=Yes
    MBRES2          As String * 9                     ' Reserved
    MBACCP          As String * 1                     ' Access path: A=Arrival K=Keyed
    MBSELO          As String * 1                     ' Select/omit file: N=No, Y=Yes
    MBCSEQ          As String * 1                     ' Alternative collating sequence:
    MBNOMB          As Long                           ' Number of members
    MBJOIN          As String * 1                     ' Join logical file:  Y=Yes, N=No
    MBRES4          As String * 9                     ' Reserved
    MBNAME          As String * 10                    ' Member
    MBCCEN          As String * 1                     ' Member creation century: 0=19xx
    MBCDAT          As String * 6                     ' Member creation date: year/mont
    MBCTIM          As String * 6                     ' Member creation time: hour/minu
    MBECEN          As String * 1                     ' Expiration century:   *=*NONE,
    MBEDAT          As String * 6                     ' Expiration date for mbr: year/m
    MBMTXT          As String * 50                    ' Member text description
    MBAPFI          As String * 10                    ' Reserved
    MBAPLB          As String * 10                    ' Reserved
    MBAPMB          As String * 10                    ' Reserved
    MBMAXM          As Long                           ' Maximum members:  0=*NOMAX
    MBMANT          As String * 1                     ' Maintenance: I=*IMMED, R=*REBLD
    MBRECV          As String * 1                     ' Access path recovery: N=*NO,S=*
    MBFKAP          As String * 1                     ' Force keyed access path: N=*NO,
    MBMXKL          As Long                           ' Maximum key length, -1 = See MB
    MBMXRL          As Long                           ' Maximum record length
    MBJRNL          As String * 1                     ' File is currently journaled: N=
    MBJRNM          As String * 10                    ' Current or last journal
    MBJRLB          As String * 10                    ' Current or last journal library
    MBJRIM          As String * 1                     ' Journal images: A=*AFTER, B=*BO
    MBJRSC          As String * 1                     ' Century of last journal start:
    MBJRSD          As String * 6                     ' Date of last journal start: yea
    MBJRST          As String * 6                     ' Time of last journal start: hou
    MBSIZ           As Long                           ' Initial number of records: 0=*N
    MBSIZI          As Long                           ' Increment number of records
    MBSIZM          As Long                           ' Maximum number of increments
    MBCURI          As Long                           ' Current number of increments
    MBRCDC          As Long                           ' Record capacity
    MBNRCD          As Long                           ' Current number of records
    MBNDTR          As Long                           ' Number of deleted records
    MBALLO          As String * 1                     ' Allocate storage: N=*NO, Y=*YES
    MBCONT          As String * 1                     ' Contiguous storage: N=*NO, Y=*Y
    MBUNIT          As Long                           ' Preferred storage unit: 0=*ANY
    MBFMTS          As String * 10                    ' Record format selector program
    MBFMSL          As String * 10                    ' Record format selector program
    MBFRCR          As Long                           ' Records to force a write:  0=*N
    MBRSHR          As String * 1                     ' Share open data path:  N=*NO, Y
    MBDLTP          As Long                           ' Max % deleted records allowed:
    MBDSSZ          As Long                           ' Data space size in bytes, -1 =
    MBISIZ          As Long                           ' Index size in bytes -1 = See MB
    MBIXNT          As Long                           ' Number of index entries for a p
    MBNACC          As Long                           ' Number of member accesses for a
    MBSEU           As String * 4                     ' Source type for S/38 View as it
    MBCHGC          As String * 1                     ' Last change century: 0=19xx, 1=
    MBCHGD          As String * 6                     ' Last change date: year/month/da
    MBCHGT          As String * 6                     ' Last change time: hour/minute/s
    MBUPDC          As String * 1                     ' Last source update century: 0=1
    MBUPDD          As String * 6                     ' Last source update date: year/m
    MBUPDT          As String * 6                     ' Last source update time: hour/m
    MBEXDC          As String * 1                     ' Extract century: 0=19xx, 1=20xx
    MBEXDD          As String * 6                     ' Extract date: year/month/day
    MBEXDT          As String * 6                     ' Extract date time: hour/minute/
    MBLEC           As String * 1                     ' Last extract date century: 0=19
    MBLED           As String * 6                     ' Last extract date: year/month/d
    MBLET           As String * 6                     ' Last extract time: hour/minute/
    MBSAVC          As String * 1                     ' Last save century: 0=19xx, 1=20
    MBSAVD          As String * 6                     ' Last save date: year/month/day
    MBSAVT          As String * 6                     ' Last save time: hour/minute/sec
    MBRSTC          As String * 1                     ' Last restore century: 0=19xx, 1
    MBRSTD          As String * 6                     ' Last restore date: year/month/d
    MBRSTT          As String * 6                     ' Last restore time: hour/minute/
    MBNSCM          As Long                           ' Members accessed by logical fil
    MBBOF           As String * 10                    ' Physical file
    MBBOL           As String * 10                    ' Library
    MBBOM           As String * 10                    ' Member
    MBBOLF          As String * 10                    ' Logical file format
    MBBOR           As Long                           ' Number of index entries
    MBBOMA          As Long                           ' Number of member accesses
    MBJROM          As String * 1                     ' Journal entries to be omitted:
    MBIST           As String * 1                     ' Implicit share type: J=Join sec
    MBISF           As String * 10                    ' File owning access path
    MBISL           As String * 10                    ' Library owning access path
    MBISM           As String * 10                    ' Member owning access path
    MBISMT          As String * 1                     ' Maintenance: I=*IMMED, R=*REBLD
    MBISRV          As String * 1                     ' Access path recovery: N=*NO,S=*
    MBISFK          As String * 1                     ' Force keyed access path: N=*NO,
    MBISUN          As String * 1                     ' Keys must be unique: N=No, Y=Ye
    MBACPJ          As String * 1                     ' Access path journaled:   N=No,
    MBALRD          As String * 1                     ' Allow read operation:  Y=Yes, N
    MBALWT          As String * 1                     ' Allow Write operation:  Y=Yes,
    MBALUP          As String * 1                     ' Allow Update operation:  Y=Yes,
    MBALDT          As String * 1                     ' Allow Delete operation:  Y=Yes,
    MBSEU2          As String * 10                    ' Source type
    MBUCEN          As String * 1                     ' Last Used Century: 0=19xx, 1=20
    MBUDAT          As String * 6                     ' Last Used Date: year/month/day
    MBUCNT          As Long                           ' Days Used Count
    MBTCEN          As String * 1                     ' Usage Data Reset Century: 0=19x
    MBTDAT          As String * 6                     ' Usage Data Reset Date: year/mon
    MBINDX          As String * 1                     ' Access Path Valid: Y=Yes, N=No,
    MBDSZ2          As Long                           ' Data space size in bytes
    MBMXK2          As Long                           ' Maximum key length
    MBOPOP          As Long                           ' Open operations
    MBCLOP          As Long                           ' Close operations
    MBWROP          As Long                           ' Write operations
    MBUPOP          As Long                           ' Update operations
    MBDLOP          As Long                           ' Delete operations
    MBLRDS          As Long                           ' Logical Reads
    MBPRDS          As Long                           ' Physical reads
    MBCROP          As Long                           ' Clear operations
    MBDSCP          As Long                           ' Data space copy operations
    MBRGOP          As Long                           ' Reorganize operations
    MBAPBL          As Long                           ' Access paths builds/rebuilds
    MBRJKY          As Long                           ' Records rejected by key selecti
    MBRJNK          As Long                           ' Records rejected by non-key sel
    MBRJGR          As Long                           ' Records rejected by group-by se
    MBACLR          As Long                           ' Access path logical reads
    MBACPR          As Long                           ' Access path physical reads
    MBISZ2          As Long                           ' Index size in bytes
    MBSTFR          As String * 1                     ' Member storage freed Y=Yes
    MBUEV1          As Long                           ' Number of Encoded  Vector Index
    MBUKV1          As Long                           ' Number of unique key values for
    MBUKV2          As Long                           ' Number of unique key values for
    MBUKV3          As Long                           ' Number of unique key values for
    MBUKV4          As Long                           ' Number of unique key values for
    MBIOVF          As Long                           ' Encoded Vector Index number of End Type
    
 End Type
Public arrDSPFDY2() As typeDSPFDY2
Public arrDSPFDY2_NB As Integer
Public arrDSPFDY2_NBMax As Integer
Public arrDSPFDY2_Index As Integer
Public arrDSPFDY2_Suite As Boolean


Public Sub srvDSPFDY2_ElpDisplay(recDSPFDY2 As typeDSPFDY2)
frmElpDisplay.fgData.Rows = 147
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBRCEN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Retrieval century:  0=19xx, 1=20xx"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBRCEN
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBRDAT    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Retrieval date:  year/month/day"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBRDAT
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBRTIM    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Retrieval time:  hour/minute/second"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBRTIM
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBFILE   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "File"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBFILE
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBLIB   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Library"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBLIB
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBFTYP    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "P=PF, L=LF, R=DDM PF, S=DDM LF"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBFTYP
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBFILA    4A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "File attribute: *PHY or *LGL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBFILA
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBMXD    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Reserved"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBMXD
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBFATR    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "File attribute:  PF, LF, PF38, or LF38"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBFATR
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBSYSN    8A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "System Name (Source System, if file is DDM)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBSYSN
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBASP    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Auxiliary storage pool ID:  1=System ASP"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBASP
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBRES    4A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Reserved"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBRES
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBDTAT    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "File type:  D=*DATA, S=*SRC"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBDTAT
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBWAIT    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Maximum file wait time: -1=*IMMED, 0=*CLS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBWAIT
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBWATR    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Maximum record wait time: -1=*IMMED, -2=*NOMAX"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBWATR
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBSHAR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Reserved"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBSHAR
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBLVLC    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Record format level check: N=*NO, Y=*YES"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBLVLC
frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBTXT   50A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Text 'description'"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBTXT
frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBNOFM    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Number of record formats"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBNOFM
frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBFCCN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Century created:  0=19xx, 1=20xx"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBFCCN
frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBFCDT    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Date created: year/month/day"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBFCDT
frmElpDisplay.fgData.Row = 22
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBFCTM    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Time created: hour/minute/second"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBFCTM
frmElpDisplay.fgData.Row = 23
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBFLS    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Externally described file:  N=No, Y=Yes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBFLS
frmElpDisplay.fgData.Row = 24
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBICAP    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DBCS capable:  N=No, Y=Yes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBICAP
frmElpDisplay.fgData.Row = 25
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBRES2    9A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Reserved"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBRES2
frmElpDisplay.fgData.Row = 26
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBACCP    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Access path: A=Arrival K=Keyed E=EVI S=Shared"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBACCP
frmElpDisplay.fgData.Row = 27
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBSELO    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Select/omit file: N=No, Y=Yes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBSELO
frmElpDisplay.fgData.Row = 28
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBCSEQ    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Alternative collating sequence:  N=No, Y=Yes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBCSEQ
frmElpDisplay.fgData.Row = 29
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBNOMB    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Number of members"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBNOMB
frmElpDisplay.fgData.Row = 30
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBJOIN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Join logical file:  Y=Yes, N=No"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBJOIN
frmElpDisplay.fgData.Row = 31
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBRES4    9A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Reserved"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBRES4
frmElpDisplay.fgData.Row = 32
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBNAME   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Member"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBNAME
frmElpDisplay.fgData.Row = 33
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBCCEN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Member creation century: 0=19xx, 1=20xx"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBCCEN
frmElpDisplay.fgData.Row = 34
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBCDAT    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Member creation date: year/month/day"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBCDAT
frmElpDisplay.fgData.Row = 35
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBCTIM    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Member creation time: hour/minute/second"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBCTIM
frmElpDisplay.fgData.Row = 36
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBECEN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Expiration century:   *=*NONE, 0=19xx, 1=20xx"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBECEN
frmElpDisplay.fgData.Row = 37
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBEDAT    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Expiration date for mbr: year/month/day or NONE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBEDAT
frmElpDisplay.fgData.Row = 38
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBMTXT   50A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Member text description"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBMTXT
frmElpDisplay.fgData.Row = 39
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBAPFI   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Reserved"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBAPFI
frmElpDisplay.fgData.Row = 40
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBAPLB   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Reserved"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBAPLB
frmElpDisplay.fgData.Row = 41
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBAPMB   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Reserved"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBAPMB
frmElpDisplay.fgData.Row = 42
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBMAXM    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Maximum members:  0=*NOMAX"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBMAXM
frmElpDisplay.fgData.Row = 43
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBMANT    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Maintenance: I=*IMMED, R=*REBLD, D=*DLY"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBMANT
frmElpDisplay.fgData.Row = 44
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBRECV    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Access path recovery: N=*NO,S=*IPL,A=*AFTIPL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBRECV
frmElpDisplay.fgData.Row = 45
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBFKAP    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Force keyed access path: N=*NO, Y=*YES"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBFKAP
frmElpDisplay.fgData.Row = 46
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBMXKL    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Maximum key length, -1 = See MBMXK2"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBMXKL
frmElpDisplay.fgData.Row = 47
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBMXRL    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Maximum record length"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBMXRL
frmElpDisplay.fgData.Row = 48
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBJRNL    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "File is currently journaled: N=No, Y=Yes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBJRNL
frmElpDisplay.fgData.Row = 49
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBJRNM   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Current or last journal"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBJRNM
frmElpDisplay.fgData.Row = 50
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBJRLB   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Current or last journal library"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBJRLB
frmElpDisplay.fgData.Row = 51
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBJRIM    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Journal images: A=*AFTER, B=*BOTH"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBJRIM
frmElpDisplay.fgData.Row = 52
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBJRSC    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Century of last journal start: 0=19xx, 1=20xx"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBJRSC
frmElpDisplay.fgData.Row = 53
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBJRSD    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Date of last journal start: year/month/day"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBJRSD
frmElpDisplay.fgData.Row = 54
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBJRST    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Time of last journal start: hour/minute/second"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBJRST
frmElpDisplay.fgData.Row = 55
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBSIZ   10P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Initial number of records: 0=*NOMAX"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBSIZ
frmElpDisplay.fgData.Row = 56
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBSIZI    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Increment number of records"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBSIZI
frmElpDisplay.fgData.Row = 57
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBSIZM    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Maximum number of increments"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBSIZM
frmElpDisplay.fgData.Row = 58
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBCURI   10P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Current number of increments"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBCURI
frmElpDisplay.fgData.Row = 59
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBRCDC   10P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Record capacity"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBRCDC
frmElpDisplay.fgData.Row = 60
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBNRCD   10P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Current number of records"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBNRCD
frmElpDisplay.fgData.Row = 61
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBNDTR   10P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Number of deleted records"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBNDTR
frmElpDisplay.fgData.Row = 62
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBALLO    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Allocate storage: N=*NO, Y=*YES"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBALLO
frmElpDisplay.fgData.Row = 63
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBCONT    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Contiguous storage: N=*NO, Y=*YES"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBCONT
frmElpDisplay.fgData.Row = 64
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBUNIT    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Preferred storage unit: 0=*ANY"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBUNIT
frmElpDisplay.fgData.Row = 65
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBFMTS   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Record format selector program"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBFMTS
frmElpDisplay.fgData.Row = 66
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBFMSL   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Record format selector program library"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBFMSL
frmElpDisplay.fgData.Row = 67
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBFRCR    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Records to force a write:  0=*NONE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBFRCR
frmElpDisplay.fgData.Row = 68
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBRSHR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Share open data path:  N=*NO, Y=*YES"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBRSHR
frmElpDisplay.fgData.Row = 69
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBDLTP    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Max % deleted records allowed: 0=*NONE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBDLTP
frmElpDisplay.fgData.Row = 70
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBDSSZ   10P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Data space size in bytes, -1 = See MBDSZ2"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBDSSZ
frmElpDisplay.fgData.Row = 71
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBISIZ   10P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Index size in bytes -1 = See MBISZ2"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBISIZ
frmElpDisplay.fgData.Row = 72
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBIXNT   10P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Number of index entries for a physical member"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBIXNT
frmElpDisplay.fgData.Row = 73
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBNACC   10P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Number of member accesses for a physical member"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBNACC
frmElpDisplay.fgData.Row = 74
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBSEU    4A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Source type for S/38 View as it appeared on S/38"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBSEU
frmElpDisplay.fgData.Row = 75
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBCHGC    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Last change century: 0=19xx, 1=20xx"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBCHGC
frmElpDisplay.fgData.Row = 76
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBCHGD    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Last change date: year/month/day"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBCHGD
frmElpDisplay.fgData.Row = 77
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBCHGT    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Last change time: hour/minute/second"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBCHGT
frmElpDisplay.fgData.Row = 78
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBUPDC    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Last source update century: 0=19xx, 1=20xx"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBUPDC
frmElpDisplay.fgData.Row = 79
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBUPDD    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Last source update date: year/month/day"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBUPDD
frmElpDisplay.fgData.Row = 80
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBUPDT    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Last source update time: hour/minute/second"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBUPDT
frmElpDisplay.fgData.Row = 81
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBEXDC    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Extract century: 0=19xx, 1=20xx"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBEXDC
frmElpDisplay.fgData.Row = 82
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBEXDD    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Extract date: year/month/day"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBEXDD
frmElpDisplay.fgData.Row = 83
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBEXDT    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Extract date time: hour/minute/second"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBEXDT
frmElpDisplay.fgData.Row = 84
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBLEC    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Last extract date century: 0=19xx, 1=20xx"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBLEC
frmElpDisplay.fgData.Row = 85
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBLED    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Last extract date: year/month/day"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBLED
frmElpDisplay.fgData.Row = 86
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBLET    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Last extract time: hour/minute/second"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBLET
frmElpDisplay.fgData.Row = 87
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBSAVC    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Last save century: 0=19xx, 1=20xx"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBSAVC
frmElpDisplay.fgData.Row = 88
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBSAVD    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Last save date: year/month/day"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBSAVD
frmElpDisplay.fgData.Row = 89
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBSAVT    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Last save time: hour/minute/second"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBSAVT
frmElpDisplay.fgData.Row = 90
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBRSTC    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Last restore century: 0=19xx, 1=20xx"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBRSTC
frmElpDisplay.fgData.Row = 91
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBRSTD    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Last restore date: year/month/day"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBRSTD
frmElpDisplay.fgData.Row = 92
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBRSTT    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Last restore time: hour/minute/second"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBRSTT
frmElpDisplay.fgData.Row = 93
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBNSCM    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Members accessed by logical file"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBNSCM
frmElpDisplay.fgData.Row = 94
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBBOF   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Physical file"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBBOF
frmElpDisplay.fgData.Row = 95
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBBOL   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Library"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBBOL
frmElpDisplay.fgData.Row = 96
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBBOM   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Member"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBBOM
frmElpDisplay.fgData.Row = 97
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBBOLF   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Logical file format"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBBOLF
frmElpDisplay.fgData.Row = 98
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBBOR   10P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Number of index entries"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBBOR
frmElpDisplay.fgData.Row = 99
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBBOMA   10P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Number of member accesses"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBBOMA
frmElpDisplay.fgData.Row = 100
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBJROM    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Journal entries to be omitted: O=*OPNCLO, N=*NONE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBJROM
frmElpDisplay.fgData.Row = 101
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBIST    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Implicit share type: J=Join secondary, N=Normal"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBIST
frmElpDisplay.fgData.Row = 102
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBISF   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "File owning access path"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBISF
frmElpDisplay.fgData.Row = 103
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBISL   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Library owning access path"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBISL
frmElpDisplay.fgData.Row = 104
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBISM   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Member owning access path"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBISM
frmElpDisplay.fgData.Row = 105
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBISMT    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Maintenance: I=*IMMED, R=*REBLD, D=*DLY"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBISMT
frmElpDisplay.fgData.Row = 106
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBISRV    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Access path recovery: N=*NO,S=*IPL,A=*AFTIPL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBISRV
frmElpDisplay.fgData.Row = 107
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBISFK    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Force keyed access path: N=*NO, Y=*YES"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBISFK
frmElpDisplay.fgData.Row = 108
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBISUN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Keys must be unique: N=No, Y=Yes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBISUN
frmElpDisplay.fgData.Row = 109
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBACPJ    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Access path journaled:   N=No, Y=Yes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBACPJ
frmElpDisplay.fgData.Row = 110
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBALRD    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Allow read operation:  Y=Yes, N=No"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBALRD
frmElpDisplay.fgData.Row = 111
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBALWT    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Allow Write operation:  Y=Yes, N=No"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBALWT
frmElpDisplay.fgData.Row = 112
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBALUP    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Allow Update operation:  Y=Yes, N=No"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBALUP
frmElpDisplay.fgData.Row = 113
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBALDT    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Allow Delete operation:  Y=Yes, N=No"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBALDT
frmElpDisplay.fgData.Row = 114
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBSEU2   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Source type"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBSEU2
frmElpDisplay.fgData.Row = 115
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBUCEN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Last Used Century: 0=19xx, 1=20xx"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBUCEN
frmElpDisplay.fgData.Row = 116
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBUDAT    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Last Used Date: year/month/day"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBUDAT
frmElpDisplay.fgData.Row = 117
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBUCNT    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Days Used Count"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBUCNT
frmElpDisplay.fgData.Row = 118
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBTCEN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Usage Data Reset Century: 0=19xx, 1=20xx"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBTCEN
frmElpDisplay.fgData.Row = 119
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBTDAT    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Usage Data Reset Date: year/month/day"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBTDAT
frmElpDisplay.fgData.Row = 120
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBINDX    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Access Path Valid: Y=Yes, N=No, H=Held"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBINDX
frmElpDisplay.fgData.Row = 121
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBDSZ2   15P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Data space size in bytes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBDSZ2
frmElpDisplay.fgData.Row = 122
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBMXK2    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Maximum key length"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBMXK2
frmElpDisplay.fgData.Row = 123
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBOPOP   20P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Open operations"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBOPOP
frmElpDisplay.fgData.Row = 124
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBCLOP   20P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Close operations"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBCLOP
frmElpDisplay.fgData.Row = 125
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBWROP   20P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Write operations"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBWROP
frmElpDisplay.fgData.Row = 126
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBUPOP   20P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Update operations"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBUPOP
frmElpDisplay.fgData.Row = 127
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBDLOP   20P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Delete operations"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBDLOP
frmElpDisplay.fgData.Row = 128
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBLRDS   20P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Logical Reads"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBLRDS
frmElpDisplay.fgData.Row = 129
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBPRDS   20P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Physical reads"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBPRDS
frmElpDisplay.fgData.Row = 130
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBCROP   20P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Clear operations"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBCROP
frmElpDisplay.fgData.Row = 131
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBDSCP   20P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Data space copy operations"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBDSCP
frmElpDisplay.fgData.Row = 132
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBRGOP   20P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Reorganize operations"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBRGOP
frmElpDisplay.fgData.Row = 133
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBAPBL   20P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Access paths builds/rebuilds"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBAPBL
frmElpDisplay.fgData.Row = 134
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBRJKY   20P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Records rejected by key selection"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBRJKY
frmElpDisplay.fgData.Row = 135
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBRJNK   20P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Records rejected by non-key selection"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBRJNK
frmElpDisplay.fgData.Row = 136
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBRJGR   20P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Records rejected by group-by selection"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBRJGR
frmElpDisplay.fgData.Row = 137
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBACLR   20P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Access path logical reads"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBACLR
frmElpDisplay.fgData.Row = 138
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBACPR   20P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Access path physical reads"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBACPR
frmElpDisplay.fgData.Row = 139
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBISZ2   15P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Index size in bytes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBISZ2
frmElpDisplay.fgData.Row = 140
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBSTFR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Member storage freed Y=Yes"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBSTFR
frmElpDisplay.fgData.Row = 141
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBUEV1   15P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Number of Encoded  Vector Index unique key values"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBUEV1
frmElpDisplay.fgData.Row = 142
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBUKV1   15P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Number of unique key values for key field one"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBUKV1
frmElpDisplay.fgData.Row = 143
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBUKV2   15P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Number of unique key values for key fields 1-2"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBUKV2
frmElpDisplay.fgData.Row = 144
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBUKV3   15P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Number of unique key values for key fields 1-3"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBUKV3
frmElpDisplay.fgData.Row = 145
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBUKV4   15P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Number of unique key values for key fields 1-4"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBUKV4
frmElpDisplay.fgData.Row = 146
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MBIOVF   15P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Encoded Vector Index number of overflow key values"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recDSPFDY2.MBIOVF
frmElpDisplay.Show vbModal
End Sub

'-----------------------------------------------------
Function srvDSPFDY2_Update(recDSPFDY2 As typeDSPFDY2)
'-----------------------------------------------------

srvDSPFDY2_Update = "?"

MsgTxtLen = 0
Call srvDSPFDY2_PutBuffer(recDSPFDY2)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvDSPFDY2_GetBuffer(recDSPFDY2)) Then
        Call srvDSPFDY2_Error(recDSPFDY2)
        srvDSPFDY2_Update = recDSPFDY2.Err
        Exit Function
    Else
        srvDSPFDY2_Update = Null
    End If
Else
    recDSPFDY2.Err = "srv"
End If
End Function

'-----------------------------------------------------
Sub srvDSPFDY2_Error(recDSPFDY2 As typeDSPFDY2)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "DSPFDY2" & Chr$(10) & Chr$(13)

Select Case mId$(recDSPFDY2.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recDSPFDY2.Err
        I = vbCritical
End Select

MsgBox Msg & " : " _
        , I, "module : DSPFDY2s.bas  ( " & Trim(recDSPFDY2.obj) & " : " & Trim(recDSPFDY2.Method) & " )"

End Sub


'=====================================================



'-----------------------------------------------------
Public Function srvDSPFDY2_Monitor(recDSPFDY2 As typeDSPFDY2)
'-----------------------------------------------------

arrDSPFDY2_Suite = False
Select Case mId$(Trim(recDSPFDY2.Method), 1, 4)
    Case "Snap"
              srvDSPFDY2_Monitor = srvDSPFDY2_Snap(recDSPFDY2)
    Case Else
            srvDSPFDY2_Monitor = srvDSPFDY2_Seek(recDSPFDY2)
End Select

End Function

'---------------------------------------------------------
Public Function srvDSPFDY2_GetBuffer(recDSPFDY2 As typeDSPFDY2)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvDSPFDY2_GetBuffer = Null
recDSPFDY2.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recDSPFDY2.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recDSPFDY2.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recDSPFDY2.Err = Space$(10) Then
    recDSPFDY2.MBRCEN = mId$(MsgTxt, K + 1, 1)
    recDSPFDY2.MBRDAT = mId$(MsgTxt, K + 2, 6)
    recDSPFDY2.MBRTIM = mId$(MsgTxt, K + 8, 6)
    recDSPFDY2.MBFILE = mId$(MsgTxt, K + 14, 10)
    recDSPFDY2.MBLIB = mId$(MsgTxt, K + 24, 10)
    recDSPFDY2.MBFTYP = mId$(MsgTxt, K + 34, 1)
    recDSPFDY2.MBFILA = mId$(MsgTxt, K + 35, 4)
    recDSPFDY2.MBMXD = mId$(MsgTxt, K + 39, 3)
    recDSPFDY2.MBFATR = mId$(MsgTxt, K + 42, 6)
    recDSPFDY2.MBSYSN = mId$(MsgTxt, K + 48, 8)
    recDSPFDY2.MBASP = CLng(Val(mId$(MsgTxt, K + 56, 4)))
    recDSPFDY2.MBRES = mId$(MsgTxt, K + 60, 4)
    recDSPFDY2.MBDTAT = mId$(MsgTxt, K + 64, 1)
    recDSPFDY2.MBWAIT = CLng(Val(mId$(MsgTxt, K + 65, 6)))
    recDSPFDY2.MBWATR = CLng(Val(mId$(MsgTxt, K + 71, 6)))
    recDSPFDY2.MBSHAR = mId$(MsgTxt, K + 77, 1)
    recDSPFDY2.MBLVLC = mId$(MsgTxt, K + 78, 1)
    recDSPFDY2.MBTXT = mId$(MsgTxt, K + 79, 50)
    recDSPFDY2.MBNOFM = CLng(Val(mId$(MsgTxt, K + 129, 6)))
    recDSPFDY2.MBFCCN = mId$(MsgTxt, K + 135, 1)
    recDSPFDY2.MBFCDT = mId$(MsgTxt, K + 136, 6)
    recDSPFDY2.MBFCTM = mId$(MsgTxt, K + 142, 6)
    recDSPFDY2.MBFLS = mId$(MsgTxt, K + 148, 1)
    recDSPFDY2.MBICAP = mId$(MsgTxt, K + 149, 1)
    recDSPFDY2.MBRES2 = mId$(MsgTxt, K + 150, 9)
    recDSPFDY2.MBACCP = mId$(MsgTxt, K + 159, 1)
    recDSPFDY2.MBSELO = mId$(MsgTxt, K + 160, 1)
    recDSPFDY2.MBCSEQ = mId$(MsgTxt, K + 161, 1)
    recDSPFDY2.MBNOMB = CLng(Val(mId$(MsgTxt, K + 162, 6)))
    recDSPFDY2.MBJOIN = mId$(MsgTxt, K + 168, 1)
    recDSPFDY2.MBRES4 = mId$(MsgTxt, K + 169, 9)
    recDSPFDY2.MBNAME = mId$(MsgTxt, K + 178, 10)
    recDSPFDY2.MBCCEN = mId$(MsgTxt, K + 188, 1)
    recDSPFDY2.MBCDAT = mId$(MsgTxt, K + 189, 6)
    recDSPFDY2.MBCTIM = mId$(MsgTxt, K + 195, 6)
    recDSPFDY2.MBECEN = mId$(MsgTxt, K + 201, 1)
    recDSPFDY2.MBEDAT = mId$(MsgTxt, K + 202, 6)
    recDSPFDY2.MBMTXT = mId$(MsgTxt, K + 208, 50)
    recDSPFDY2.MBAPFI = mId$(MsgTxt, K + 258, 10)
    recDSPFDY2.MBAPLB = mId$(MsgTxt, K + 268, 10)
    recDSPFDY2.MBAPMB = mId$(MsgTxt, K + 278, 10)
    recDSPFDY2.MBMAXM = CLng(Val(mId$(MsgTxt, K + 288, 6)))
    recDSPFDY2.MBMANT = mId$(MsgTxt, K + 294, 1)
    recDSPFDY2.MBRECV = mId$(MsgTxt, K + 295, 1)
    recDSPFDY2.MBFKAP = mId$(MsgTxt, K + 296, 1)
    recDSPFDY2.MBMXKL = CLng(Val(mId$(MsgTxt, K + 297, 4)))
    recDSPFDY2.MBMXRL = CLng(Val(mId$(MsgTxt, K + 301, 6)))
    recDSPFDY2.MBJRNL = mId$(MsgTxt, K + 307, 1)
    recDSPFDY2.MBJRNM = mId$(MsgTxt, K + 308, 10)
    recDSPFDY2.MBJRLB = mId$(MsgTxt, K + 318, 10)
    recDSPFDY2.MBJRIM = mId$(MsgTxt, K + 328, 1)
    recDSPFDY2.MBJRSC = mId$(MsgTxt, K + 329, 1)
    recDSPFDY2.MBJRSD = mId$(MsgTxt, K + 330, 6)
    recDSPFDY2.MBJRST = mId$(MsgTxt, K + 336, 6)
    recDSPFDY2.MBSIZ = CLng(Val(mId$(MsgTxt, K + 342, 11)))
    recDSPFDY2.MBSIZI = CLng(Val(mId$(MsgTxt, K + 353, 6)))
    recDSPFDY2.MBSIZM = CLng(Val(mId$(MsgTxt, K + 359, 6)))
    recDSPFDY2.MBCURI = CLng(Val(mId$(MsgTxt, K + 365, 11)))
    recDSPFDY2.MBRCDC = CLng(Val(mId$(MsgTxt, K + 376, 11)))
    recDSPFDY2.MBNRCD = CLng(Val(mId$(MsgTxt, K + 387, 11)))
    recDSPFDY2.MBNDTR = CLng(Val(mId$(MsgTxt, K + 398, 11)))
    recDSPFDY2.MBALLO = mId$(MsgTxt, K + 409, 1)
    recDSPFDY2.MBCONT = mId$(MsgTxt, K + 410, 1)
    recDSPFDY2.MBUNIT = CLng(Val(mId$(MsgTxt, K + 411, 4)))
    recDSPFDY2.MBFMTS = mId$(MsgTxt, K + 415, 10)
    recDSPFDY2.MBFMSL = mId$(MsgTxt, K + 425, 10)
    recDSPFDY2.MBFRCR = CLng(Val(mId$(MsgTxt, K + 435, 6)))
    recDSPFDY2.MBRSHR = mId$(MsgTxt, K + 441, 1)
    recDSPFDY2.MBDLTP = CLng(Val(mId$(MsgTxt, K + 442, 4)))
    recDSPFDY2.MBDSSZ = CLng(Val(mId$(MsgTxt, K + 446, 11)))
    recDSPFDY2.MBISIZ = CLng(Val(mId$(MsgTxt, K + 457, 11)))
    recDSPFDY2.MBIXNT = CLng(Val(mId$(MsgTxt, K + 468, 11)))
    recDSPFDY2.MBNACC = CLng(Val(mId$(MsgTxt, K + 479, 11)))
    recDSPFDY2.MBSEU = mId$(MsgTxt, K + 490, 4)
    recDSPFDY2.MBCHGC = mId$(MsgTxt, K + 494, 1)
    recDSPFDY2.MBCHGD = mId$(MsgTxt, K + 495, 6)
    recDSPFDY2.MBCHGT = mId$(MsgTxt, K + 501, 6)
    recDSPFDY2.MBUPDC = mId$(MsgTxt, K + 507, 1)
    recDSPFDY2.MBUPDD = mId$(MsgTxt, K + 508, 6)
    recDSPFDY2.MBUPDT = mId$(MsgTxt, K + 514, 6)
    recDSPFDY2.MBEXDC = mId$(MsgTxt, K + 520, 1)
    recDSPFDY2.MBEXDD = mId$(MsgTxt, K + 521, 6)
    recDSPFDY2.MBEXDT = mId$(MsgTxt, K + 527, 6)
    recDSPFDY2.MBLEC = mId$(MsgTxt, K + 533, 1)
    recDSPFDY2.MBLED = mId$(MsgTxt, K + 534, 6)
    recDSPFDY2.MBLET = mId$(MsgTxt, K + 540, 6)
    recDSPFDY2.MBSAVC = mId$(MsgTxt, K + 546, 1)
    recDSPFDY2.MBSAVD = mId$(MsgTxt, K + 547, 6)
    recDSPFDY2.MBSAVT = mId$(MsgTxt, K + 553, 6)
    recDSPFDY2.MBRSTC = mId$(MsgTxt, K + 559, 1)
    recDSPFDY2.MBRSTD = mId$(MsgTxt, K + 560, 6)
    recDSPFDY2.MBRSTT = mId$(MsgTxt, K + 566, 6)
    recDSPFDY2.MBNSCM = CLng(Val(mId$(MsgTxt, K + 572, 4)))
    recDSPFDY2.MBBOF = mId$(MsgTxt, K + 576, 10)
    recDSPFDY2.MBBOL = mId$(MsgTxt, K + 586, 10)
    recDSPFDY2.MBBOM = mId$(MsgTxt, K + 596, 10)
    recDSPFDY2.MBBOLF = mId$(MsgTxt, K + 606, 10)
    recDSPFDY2.MBBOR = CLng(Val(mId$(MsgTxt, K + 616, 11)))
    recDSPFDY2.MBBOMA = CLng(Val(mId$(MsgTxt, K + 627, 11)))
    recDSPFDY2.MBJROM = mId$(MsgTxt, K + 638, 1)
    recDSPFDY2.MBIST = mId$(MsgTxt, K + 639, 1)
    recDSPFDY2.MBISF = mId$(MsgTxt, K + 640, 10)
    recDSPFDY2.MBISL = mId$(MsgTxt, K + 650, 10)
    recDSPFDY2.MBISM = mId$(MsgTxt, K + 660, 10)
    recDSPFDY2.MBISMT = mId$(MsgTxt, K + 670, 1)
    recDSPFDY2.MBISRV = mId$(MsgTxt, K + 671, 1)
    recDSPFDY2.MBISFK = mId$(MsgTxt, K + 672, 1)
    recDSPFDY2.MBISUN = mId$(MsgTxt, K + 673, 1)
    recDSPFDY2.MBACPJ = mId$(MsgTxt, K + 674, 1)
    recDSPFDY2.MBALRD = mId$(MsgTxt, K + 675, 1)
    recDSPFDY2.MBALWT = mId$(MsgTxt, K + 676, 1)
    recDSPFDY2.MBALUP = mId$(MsgTxt, K + 677, 1)
    recDSPFDY2.MBALDT = mId$(MsgTxt, K + 678, 1)
    recDSPFDY2.MBSEU2 = mId$(MsgTxt, K + 679, 10)
    recDSPFDY2.MBUCEN = mId$(MsgTxt, K + 689, 1)
    recDSPFDY2.MBUDAT = mId$(MsgTxt, K + 690, 6)
    recDSPFDY2.MBUCNT = CLng(Val(mId$(MsgTxt, K + 696, 6)))
    recDSPFDY2.MBTCEN = mId$(MsgTxt, K + 702, 1)
    recDSPFDY2.MBTDAT = mId$(MsgTxt, K + 703, 6)
    recDSPFDY2.MBINDX = mId$(MsgTxt, K + 709, 1)
    recDSPFDY2.MBDSZ2 = CLng(Val(mId$(MsgTxt, K + 710, 16)))
    recDSPFDY2.MBMXK2 = CLng(Val(mId$(MsgTxt, K + 726, 6)))
    recDSPFDY2.MBOPOP = CLng(Val(mId$(MsgTxt, K + 732, 21)))
    recDSPFDY2.MBCLOP = CLng(Val(mId$(MsgTxt, K + 753, 21)))
    recDSPFDY2.MBWROP = CLng(Val(mId$(MsgTxt, K + 774, 21)))
    recDSPFDY2.MBUPOP = CLng(Val(mId$(MsgTxt, K + 795, 21)))
    recDSPFDY2.MBDLOP = CLng(Val(mId$(MsgTxt, K + 816, 21)))
    recDSPFDY2.MBLRDS = CLng(Val(mId$(MsgTxt, K + 837, 21)))
    recDSPFDY2.MBPRDS = CLng(Val(mId$(MsgTxt, K + 858, 21)))
    recDSPFDY2.MBCROP = CLng(Val(mId$(MsgTxt, K + 879, 21)))
    recDSPFDY2.MBDSCP = CLng(Val(mId$(MsgTxt, K + 900, 21)))
    recDSPFDY2.MBRGOP = CLng(Val(mId$(MsgTxt, K + 921, 21)))
    recDSPFDY2.MBAPBL = CLng(Val(mId$(MsgTxt, K + 942, 21)))
    recDSPFDY2.MBRJKY = CLng(Val(mId$(MsgTxt, K + 963, 21)))
    recDSPFDY2.MBRJNK = CLng(Val(mId$(MsgTxt, K + 984, 21)))
    recDSPFDY2.MBRJGR = CLng(Val(mId$(MsgTxt, K + 1005, 21)))
    recDSPFDY2.MBACLR = CLng(Val(mId$(MsgTxt, K + 1026, 21)))
    recDSPFDY2.MBACPR = CLng(Val(mId$(MsgTxt, K + 1047, 21)))
    recDSPFDY2.MBISZ2 = CLng(Val(mId$(MsgTxt, K + 1068, 16)))
    recDSPFDY2.MBSTFR = mId$(MsgTxt, K + 1084, 1)
    recDSPFDY2.MBUEV1 = CLng(Val(mId$(MsgTxt, K + 1085, 16)))
    recDSPFDY2.MBUKV1 = CLng(Val(mId$(MsgTxt, K + 1101, 16)))
    recDSPFDY2.MBUKV2 = CLng(Val(mId$(MsgTxt, K + 1117, 16)))
    recDSPFDY2.MBUKV3 = CLng(Val(mId$(MsgTxt, K + 1133, 16)))
    recDSPFDY2.MBUKV4 = CLng(Val(mId$(MsgTxt, K + 1149, 16)))
    recDSPFDY2.MBIOVF = CLng(Val(mId$(MsgTxt, K + 1165, 16)))
Else
    srvDSPFDY2_GetBuffer = recDSPFDY2.Err
End If

MsgTxtIndex = MsgTxtIndex + recDSPFDY2Len

End Function

'---------------------------------------------------------
Private Sub srvDSPFDY2_PutBuffer(recDSPFDY2 As typeDSPFDY2)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recDSPFDY2.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recDSPFDY2.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
    Mid$(MsgTxt, K + 1, 1) = recDSPFDY2.MBRCEN
    Mid$(MsgTxt, K + 2, 6) = recDSPFDY2.MBRDAT
    Mid$(MsgTxt, K + 8, 6) = recDSPFDY2.MBRTIM
    Mid$(MsgTxt, K + 14, 10) = recDSPFDY2.MBFILE
    Mid$(MsgTxt, K + 24, 10) = recDSPFDY2.MBLIB
    Mid$(MsgTxt, K + 34, 1) = recDSPFDY2.MBFTYP
    Mid$(MsgTxt, K + 35, 4) = recDSPFDY2.MBFILA
    Mid$(MsgTxt, K + 39, 3) = recDSPFDY2.MBMXD
    Mid$(MsgTxt, K + 42, 6) = recDSPFDY2.MBFATR
    Mid$(MsgTxt, K + 48, 8) = recDSPFDY2.MBSYSN
    Mid$(MsgTxt, K + 56, 4) = Format$(recDSPFDY2.MBASP, "000 ")
    Mid$(MsgTxt, K + 60, 4) = recDSPFDY2.MBRES
    Mid$(MsgTxt, K + 64, 1) = recDSPFDY2.MBDTAT
    Mid$(MsgTxt, K + 65, 6) = Format$(recDSPFDY2.MBWAIT, "00000 ")
    Mid$(MsgTxt, K + 71, 6) = Format$(recDSPFDY2.MBWATR, "00000 ")
    Mid$(MsgTxt, K + 77, 1) = recDSPFDY2.MBSHAR
    Mid$(MsgTxt, K + 78, 1) = recDSPFDY2.MBLVLC
    Mid$(MsgTxt, K + 79, 50) = recDSPFDY2.MBTXT
    Mid$(MsgTxt, K + 129, 6) = Format$(recDSPFDY2.MBNOFM, "00000 ")
    Mid$(MsgTxt, K + 135, 1) = recDSPFDY2.MBFCCN
    Mid$(MsgTxt, K + 136, 6) = recDSPFDY2.MBFCDT
    Mid$(MsgTxt, K + 142, 6) = recDSPFDY2.MBFCTM
    Mid$(MsgTxt, K + 148, 1) = recDSPFDY2.MBFLS
    Mid$(MsgTxt, K + 149, 1) = recDSPFDY2.MBICAP
    Mid$(MsgTxt, K + 150, 9) = recDSPFDY2.MBRES2
    Mid$(MsgTxt, K + 159, 1) = recDSPFDY2.MBACCP
    Mid$(MsgTxt, K + 160, 1) = recDSPFDY2.MBSELO
    Mid$(MsgTxt, K + 161, 1) = recDSPFDY2.MBCSEQ
    Mid$(MsgTxt, K + 162, 6) = Format$(recDSPFDY2.MBNOMB, "00000 ")
    Mid$(MsgTxt, K + 168, 1) = recDSPFDY2.MBJOIN
    Mid$(MsgTxt, K + 169, 9) = recDSPFDY2.MBRES4
    Mid$(MsgTxt, K + 178, 10) = recDSPFDY2.MBNAME
    Mid$(MsgTxt, K + 188, 1) = recDSPFDY2.MBCCEN
    Mid$(MsgTxt, K + 189, 6) = recDSPFDY2.MBCDAT
    Mid$(MsgTxt, K + 195, 6) = recDSPFDY2.MBCTIM
    Mid$(MsgTxt, K + 201, 1) = recDSPFDY2.MBECEN
    Mid$(MsgTxt, K + 202, 6) = recDSPFDY2.MBEDAT
    Mid$(MsgTxt, K + 208, 50) = recDSPFDY2.MBMTXT
    Mid$(MsgTxt, K + 258, 10) = recDSPFDY2.MBAPFI
    Mid$(MsgTxt, K + 268, 10) = recDSPFDY2.MBAPLB
    Mid$(MsgTxt, K + 278, 10) = recDSPFDY2.MBAPMB
    Mid$(MsgTxt, K + 288, 6) = Format$(recDSPFDY2.MBMAXM, "00000 ")
    Mid$(MsgTxt, K + 294, 1) = recDSPFDY2.MBMANT
    Mid$(MsgTxt, K + 295, 1) = recDSPFDY2.MBRECV
    Mid$(MsgTxt, K + 296, 1) = recDSPFDY2.MBFKAP
    Mid$(MsgTxt, K + 297, 4) = Format$(recDSPFDY2.MBMXKL, "000 ")
    Mid$(MsgTxt, K + 301, 6) = Format$(recDSPFDY2.MBMXRL, "00000 ")
    Mid$(MsgTxt, K + 307, 1) = recDSPFDY2.MBJRNL
    Mid$(MsgTxt, K + 308, 10) = recDSPFDY2.MBJRNM
    Mid$(MsgTxt, K + 318, 10) = recDSPFDY2.MBJRLB
    Mid$(MsgTxt, K + 328, 1) = recDSPFDY2.MBJRIM
    Mid$(MsgTxt, K + 329, 1) = recDSPFDY2.MBJRSC
    Mid$(MsgTxt, K + 330, 6) = recDSPFDY2.MBJRSD
    Mid$(MsgTxt, K + 336, 6) = recDSPFDY2.MBJRST
    Mid$(MsgTxt, K + 342, 11) = Format$(recDSPFDY2.MBSIZ, "0000000000 ")
    Mid$(MsgTxt, K + 353, 6) = Format$(recDSPFDY2.MBSIZI, "00000 ")
    Mid$(MsgTxt, K + 359, 6) = Format$(recDSPFDY2.MBSIZM, "00000 ")
    Mid$(MsgTxt, K + 365, 11) = Format$(recDSPFDY2.MBCURI, "0000000000 ")
    Mid$(MsgTxt, K + 376, 11) = Format$(recDSPFDY2.MBRCDC, "0000000000 ")
    Mid$(MsgTxt, K + 387, 11) = Format$(recDSPFDY2.MBNRCD, "0000000000 ")
    Mid$(MsgTxt, K + 398, 11) = Format$(recDSPFDY2.MBNDTR, "0000000000 ")
    Mid$(MsgTxt, K + 409, 1) = recDSPFDY2.MBALLO
    Mid$(MsgTxt, K + 410, 1) = recDSPFDY2.MBCONT
    Mid$(MsgTxt, K + 411, 4) = Format$(recDSPFDY2.MBUNIT, "000 ")
    Mid$(MsgTxt, K + 415, 10) = recDSPFDY2.MBFMTS
    Mid$(MsgTxt, K + 425, 10) = recDSPFDY2.MBFMSL
    Mid$(MsgTxt, K + 435, 6) = Format$(recDSPFDY2.MBFRCR, "00000 ")
    Mid$(MsgTxt, K + 441, 1) = recDSPFDY2.MBRSHR
    Mid$(MsgTxt, K + 442, 4) = Format$(recDSPFDY2.MBDLTP, "000 ")
    Mid$(MsgTxt, K + 446, 11) = Format$(recDSPFDY2.MBDSSZ, "0000000000 ")
    Mid$(MsgTxt, K + 457, 11) = Format$(recDSPFDY2.MBISIZ, "0000000000 ")
    Mid$(MsgTxt, K + 468, 11) = Format$(recDSPFDY2.MBIXNT, "0000000000 ")
    Mid$(MsgTxt, K + 479, 11) = Format$(recDSPFDY2.MBNACC, "0000000000 ")
    Mid$(MsgTxt, K + 490, 4) = recDSPFDY2.MBSEU
    Mid$(MsgTxt, K + 494, 1) = recDSPFDY2.MBCHGC
    Mid$(MsgTxt, K + 495, 6) = recDSPFDY2.MBCHGD
    Mid$(MsgTxt, K + 501, 6) = recDSPFDY2.MBCHGT
    Mid$(MsgTxt, K + 507, 1) = recDSPFDY2.MBUPDC
    Mid$(MsgTxt, K + 508, 6) = recDSPFDY2.MBUPDD
    Mid$(MsgTxt, K + 514, 6) = recDSPFDY2.MBUPDT
    Mid$(MsgTxt, K + 520, 1) = recDSPFDY2.MBEXDC
    Mid$(MsgTxt, K + 521, 6) = recDSPFDY2.MBEXDD
    Mid$(MsgTxt, K + 527, 6) = recDSPFDY2.MBEXDT
    Mid$(MsgTxt, K + 533, 1) = recDSPFDY2.MBLEC
    Mid$(MsgTxt, K + 534, 6) = recDSPFDY2.MBLED
    Mid$(MsgTxt, K + 540, 6) = recDSPFDY2.MBLET
    Mid$(MsgTxt, K + 546, 1) = recDSPFDY2.MBSAVC
    Mid$(MsgTxt, K + 547, 6) = recDSPFDY2.MBSAVD
    Mid$(MsgTxt, K + 553, 6) = recDSPFDY2.MBSAVT
    Mid$(MsgTxt, K + 559, 1) = recDSPFDY2.MBRSTC
    Mid$(MsgTxt, K + 560, 6) = recDSPFDY2.MBRSTD
    Mid$(MsgTxt, K + 566, 6) = recDSPFDY2.MBRSTT
    Mid$(MsgTxt, K + 572, 4) = Format$(recDSPFDY2.MBNSCM, "000 ")
    Mid$(MsgTxt, K + 576, 10) = recDSPFDY2.MBBOF
    Mid$(MsgTxt, K + 586, 10) = recDSPFDY2.MBBOL
    Mid$(MsgTxt, K + 596, 10) = recDSPFDY2.MBBOM
    Mid$(MsgTxt, K + 606, 10) = recDSPFDY2.MBBOLF
    Mid$(MsgTxt, K + 616, 11) = Format$(recDSPFDY2.MBBOR, "0000000000 ")
    Mid$(MsgTxt, K + 627, 11) = Format$(recDSPFDY2.MBBOMA, "0000000000 ")
    Mid$(MsgTxt, K + 638, 1) = recDSPFDY2.MBJROM
    Mid$(MsgTxt, K + 639, 1) = recDSPFDY2.MBIST
    Mid$(MsgTxt, K + 640, 10) = recDSPFDY2.MBISF
    Mid$(MsgTxt, K + 650, 10) = recDSPFDY2.MBISL
    Mid$(MsgTxt, K + 660, 10) = recDSPFDY2.MBISM
    Mid$(MsgTxt, K + 670, 1) = recDSPFDY2.MBISMT
    Mid$(MsgTxt, K + 671, 1) = recDSPFDY2.MBISRV
    Mid$(MsgTxt, K + 672, 1) = recDSPFDY2.MBISFK
    Mid$(MsgTxt, K + 673, 1) = recDSPFDY2.MBISUN
    Mid$(MsgTxt, K + 674, 1) = recDSPFDY2.MBACPJ
    Mid$(MsgTxt, K + 675, 1) = recDSPFDY2.MBALRD
    Mid$(MsgTxt, K + 676, 1) = recDSPFDY2.MBALWT
    Mid$(MsgTxt, K + 677, 1) = recDSPFDY2.MBALUP
    Mid$(MsgTxt, K + 678, 1) = recDSPFDY2.MBALDT
    Mid$(MsgTxt, K + 679, 10) = recDSPFDY2.MBSEU2
    Mid$(MsgTxt, K + 689, 1) = recDSPFDY2.MBUCEN
    Mid$(MsgTxt, K + 690, 6) = recDSPFDY2.MBUDAT
    Mid$(MsgTxt, K + 696, 6) = Format$(recDSPFDY2.MBUCNT, "00000 ")
    Mid$(MsgTxt, K + 702, 1) = recDSPFDY2.MBTCEN
    Mid$(MsgTxt, K + 703, 6) = recDSPFDY2.MBTDAT
    Mid$(MsgTxt, K + 709, 1) = recDSPFDY2.MBINDX
    Mid$(MsgTxt, K + 710, 16) = Format$(recDSPFDY2.MBDSZ2, "000000000000000 ")
    Mid$(MsgTxt, K + 726, 6) = Format$(recDSPFDY2.MBMXK2, "00000 ")
    Mid$(MsgTxt, K + 732, 21) = Format$(recDSPFDY2.MBOPOP, "00000000000000000000 ")
    Mid$(MsgTxt, K + 753, 21) = Format$(recDSPFDY2.MBCLOP, "00000000000000000000 ")
    Mid$(MsgTxt, K + 774, 21) = Format$(recDSPFDY2.MBWROP, "00000000000000000000 ")
    Mid$(MsgTxt, K + 795, 21) = Format$(recDSPFDY2.MBUPOP, "00000000000000000000 ")
    Mid$(MsgTxt, K + 816, 21) = Format$(recDSPFDY2.MBDLOP, "00000000000000000000 ")
    Mid$(MsgTxt, K + 837, 21) = Format$(recDSPFDY2.MBLRDS, "00000000000000000000 ")
    Mid$(MsgTxt, K + 858, 21) = Format$(recDSPFDY2.MBPRDS, "00000000000000000000 ")
    Mid$(MsgTxt, K + 879, 21) = Format$(recDSPFDY2.MBCROP, "00000000000000000000 ")
    Mid$(MsgTxt, K + 900, 21) = Format$(recDSPFDY2.MBDSCP, "00000000000000000000 ")
    Mid$(MsgTxt, K + 921, 21) = Format$(recDSPFDY2.MBRGOP, "00000000000000000000 ")
    Mid$(MsgTxt, K + 942, 21) = Format$(recDSPFDY2.MBAPBL, "00000000000000000000 ")
    Mid$(MsgTxt, K + 963, 21) = Format$(recDSPFDY2.MBRJKY, "00000000000000000000 ")
    Mid$(MsgTxt, K + 984, 21) = Format$(recDSPFDY2.MBRJNK, "00000000000000000000 ")
    Mid$(MsgTxt, K + 1005, 21) = Format$(recDSPFDY2.MBRJGR, "00000000000000000000 ")
    Mid$(MsgTxt, K + 1026, 21) = Format$(recDSPFDY2.MBACLR, "00000000000000000000 ")
    Mid$(MsgTxt, K + 1047, 21) = Format$(recDSPFDY2.MBACPR, "00000000000000000000 ")
    Mid$(MsgTxt, K + 1068, 16) = Format$(recDSPFDY2.MBISZ2, "000000000000000 ")
    Mid$(MsgTxt, K + 1084, 1) = recDSPFDY2.MBSTFR
    Mid$(MsgTxt, K + 1085, 16) = Format$(recDSPFDY2.MBUEV1, "000000000000000 ")
    Mid$(MsgTxt, K + 1101, 16) = Format$(recDSPFDY2.MBUKV1, "000000000000000 ")
    Mid$(MsgTxt, K + 1117, 16) = Format$(recDSPFDY2.MBUKV2, "000000000000000 ")
    Mid$(MsgTxt, K + 1133, 16) = Format$(recDSPFDY2.MBUKV3, "000000000000000 ")
    Mid$(MsgTxt, K + 1149, 16) = Format$(recDSPFDY2.MBUKV4, "000000000000000 ")
    Mid$(MsgTxt, K + 1165, 16) = Format$(recDSPFDY2.MBIOVF, "000000000000000 ")

    

MsgTxtLen = MsgTxtLen + recDSPFDY2Len
End Sub



'---------------------------------------------------------
Private Function srvDSPFDY2_Seek(recDSPFDY2 As typeDSPFDY2)
'---------------------------------------------------------

srvDSPFDY2_Seek = "?"
MsgTxtLen = 0
Call srvDSPFDY2_PutBuffer(recDSPFDY2)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvDSPFDY2_GetBuffer(recDSPFDY2)) Then
        srvDSPFDY2_Seek = Null
    Else
        Call srvDSPFDY2_Error(recDSPFDY2)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvDSPFDY2_Snap(recDSPFDY2 As typeDSPFDY2)
'---------------------------------------------------------
srvDSPFDY2_Snap = "?"
MsgTxtLen = 0
Call srvDSPFDY2_PutBuffer(recDSPFDY2)
Call srvDSPFDY2_PutBuffer(arrDSPFDY2(0))
If IsNull(SndRcv()) Then
    srvDSPFDY2_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvDSPFDY2_GetBuffer(recDSPFDY2)) Then
            Call arrDSPFDY2_AddItem(recDSPFDY2)
            arrDSPFDY2_Suite = True
        Else
            arrDSPFDY2_Suite = False
            Exit Do
        End If
    Loop
End If

End Function
'---------------------------------------------------------
Public Sub arrDSPFDY2_AddItem(recDSPFDY2 As typeDSPFDY2)
'---------------------------------------------------------
          
arrDSPFDY2_NB = arrDSPFDY2_NB + 1
    
If arrDSPFDY2_NB > arrDSPFDY2_NBMax Then
    arrDSPFDY2_NBMax = arrDSPFDY2_NBMax + recDSPFDY2_Block
    ReDim Preserve arrDSPFDY2(arrDSPFDY2_NBMax)
End If
            
arrDSPFDY2(arrDSPFDY2_NB) = recDSPFDY2
End Sub



'---------------------------------------------------------
Public Sub recDSPFDY2_Init(recDSPFDY2 As typeDSPFDY2)
'---------------------------------------------------------
recDSPFDY2.obj = "DSPFDY2"
recDSPFDY2.Method = ""
recDSPFDY2.Err = ""

End Sub






