VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BIFFReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Unique As Long
Dim sstcount As Long

Private Type BIFF_SECTOR
    Data() As Byte
End Type

Private SST() As String

'0  64  Character array of the name of the entry, always 16-bit Unicode characters, with trailing
'       zero character (results in a maximum name length of 31 characters)
'64  2  Size of the used area of the character buffer of the name (not character count), including
'       the trailing zero character (e.g. 12 for a name with 5 characters: (5+1)·2 = 12)
'66  1  Type of the entry:  00H = Empty 03H = LockBytes (unknown)
'                           01H = User storage 04H = Property (unknown)
'                           02H = User stream 05H = Root storage
'67  1  Node colour of the entry: 00H = Red 01H = Black
'68  4  DID of the left child node inside the red-black tree of all direct members of the parent
'       storage (if this entry is a user storage or stream, .7.1), –1 if there is no left child
'72  4  DID of the right child node inside the red-black tree of all direct members of the parent
'       storage (if this entry is a user storage or stream, .7.1), –1 if there is no right child
'76  4  DID of the root node entry of the red-black tree of all storage members (if this entry is a
'       storage, .7.1), –1 otherwise
'80 16  Unique identifier, if this is a storage (not of interest in the following, may be all 0)
'96  4  User flags (not of interest in the following, may be all 0)
'100 8  Time stamp of creation of this entry (.7.2.3). Most implementations do not write a valid
'       time stamp, but fill up this space with zero bytes.
'108 8  Time stamp of last modification of this entry (.7.2.3). Most implementations do not write
'       a valid time stamp, but fill up this space with zero bytes.
'116 4  SID of first sector or short-sector, if this entry refers to a stream (.7.2.2), SID of first
'       sector of the short-stream container stream (.6.1), if this is the root storage entry, 0
'       otherwise
'120 4  Total stream size in bytes, if this entry refers to a stream (.7.2.2), total size of the shortstream
'       container stream (.6.1), if this is the root storage entry, 0 otherwise
'124 4  Not used

Private Type BIFF_DIRECTORY
    DirName(63) As Byte
    NameLength As Integer
    entrytype As Byte
    NodeColor As Byte
    LeftChildDID As Long
    RightChildDID As Long
    RootDID As Long
    UID(15) As Byte
    UserFlags As Long
    CreateTime(7) As Byte
    ModifyTime(7) As Byte
    FirstSectorSID As Long
    StreamSize As Long
    UnUsed As Long
End Type


Dim XF() As Integer
Dim xfcount As Long

'0    8     Compound document file identifier: D0H CFH 11H E0H A1H B1H 1AH E1H
'8   16     Unique identifier (UID) of this file (not of interest in the following, may be all 0)
'24   2     Revision number of the file format (most used is 003EH)
'26   2     Version number of the file format (most used is 0003H)
'28   2     Byte order identifier (.4.2):   FEH FFH = Little-Endian
'                                           FFH FEH = Big - Endian
'30   2     Size of a sector in the compound document file (.3.1) in power-of-two (ssz), real sector
'           size is sec_size = 2ssz bytes (most used value is 9 which means 512 bytes, minimum
'           value is 7 which means 128 bytes)
'32   2     Size of a short-sector in the short-stream container stream (.6.1) in power-of-two (sssz),
'           real short-sector size is short_sec_size = 2sssz bytes (most used value is 6 which
'           means 64 bytes, maximum value is sector size ssz, see above)
'34  10     Not used
'44   4     Total number of sectors used for the sector allocation table (.5.2)
'48   4     SID of first sector of the directory stream (.7)
'52   4     Not used
'56   4     Minimum size of a standard stream (in bytes, most used size is 4096 bytes), streams
'           smaller than this value are stored as short-streams (.6)
'60   4     SID of first sector of the short-sector allocation table (.6.2), or .2 (End Of Chain SID,
'           .3.1) if not extant
'64   4     Total number of sectors used for the short-sector allocation table (.6.2)
'68   4     SID of first sector of the master sector allocation table (.5.1), or .2 (End Of Chain SID,
'           .3.1) if no additional sectors used
'72   4     Total number of sectors used for the master sector allocation table (.5.1)
'76 436     First part of the master sector allocation table (.5.1) containing 109 SIDs

Private Type BIFF_FILE_HEADER
    Header As String * 8
    UID(0 To 15) As Byte
    Revision As Integer
    Version As Integer
    ByteOrder(0 To 1) As Byte
    SectorSize As Integer
    ShortSectorSize As Integer
    UnUsed(0 To 9) As Byte
    SATSize As Long
    FirstDirSID As Long
    UnUsed2(0 To 3) As Byte
    MinStdStream As Long
    FirstShortSATSID As Long
    ShortSATSize As Long
    FirstMSATSID As Long
    MSATSize As Long
    FirstMSAT(0 To 108) As Long
End Type

Dim checkBIFFHeaderString As String

Dim myBIFFHeader As BIFF_FILE_HEADER

Public WorkBook As New cWorkBook
Dim currentworksheet As cWorkSheet
Dim currow As Long
Dim curcol As Long

Private DateMode As Integer

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

Private Sub Class_Initialize()
    checkBIFFHeaderString = Chr(&HD0) & Chr(&HCF) & Chr(&H11) & Chr(&HE0) & Chr(&HA1) & Chr(&HB1) & Chr(&H1A) & Chr(&HE1)
End Sub

Property Get WorkSheet(index) As cWorkSheet
    Set WorkSheet = WorkBook.WorkSheet(index)
End Property

Property Get WorkSheets() As Collection
    Set WorkSheets = WorkBook.WorkSheets
End Property

Public Function OpenBIFF(filename As String) As Boolean
Dim myFile As Long
Dim Sector() As BIFF_SECTOR

    myFile = FreeFile
    If Dir(filename) <> "" Then
        Open filename For Binary As myFile
        With myBIFFHeader
            Get #myFile, , .Header
            Get #myFile, , .UID
            Get #myFile, , .Revision
            Get #myFile, , .Version
            Get #myFile, , .ByteOrder
            Get #myFile, , .SectorSize
            Get #myFile, , .ShortSectorSize
            Get #myFile, , .UnUsed
            Get #myFile, , .SATSize
            Get #myFile, , .FirstDirSID
            Get #myFile, , .UnUsed2
            Get #myFile, , .MinStdStream
            Get #myFile, , .FirstShortSATSID
            Get #myFile, , .ShortSATSize
            Get #myFile, , .FirstMSATSID
            Get #myFile, , .MSATSize
            Get #myFile, , .FirstMSAT
        End With
                
        If myBIFFHeader.Header = checkBIFFHeaderString Then
            With myBIFFHeader
                'Trace "Byte Order: " & Hex(.ByteOrder(0)) & Hex(.ByteOrder(1))
                'Trace "ver: " & .Version & " rev: " & .Revision
                'Trace "Sector size: " & 2 ^ .SectorSize
                'Trace "Short Sector size: " & 2 ^ .ShortSectorSize
                'Trace "SAT Size: " & .SATSize
                'Trace "Short SAT Size: " & .ShortSATSize
                'Trace "Master SAT size: " & .MSATSize
                'Trace "First Master SAT SID: " & .FirstMSATSID
                'Trace "First Short SAT SID: " & .FirstShortSATSID
                'Trace "First Directory SID: " & .FirstDirSID
                
                .SectorSize = 2 ^ .SectorSize
                .ShortSectorSize = 2 ^ .ShortSectorSize
                
                Dim sectorcount As Long
                
                While Not EOF(myFile)
                    ReDim Preserve Sector(sectorcount)
                    ReDim Preserve Sector(sectorcount).Data(.SectorSize - 1)
                    Get #myFile, , Sector(sectorcount).Data
                    sectorcount = sectorcount + 1
                Wend
                
                Dim SAT() As Long
                ReDim SAT((.SATSize * .SectorSize / 4) - 1)
                
                i = 0
                While .FirstMSAT(i) > -1
                    CopyMemory SAT(i * .SectorSize / 4), Sector(.FirstMSAT(i)).Data(0), .SectorSize
                    i = i + 1
                Wend
                ' Need to handle if MSAT extends beyond first 109 bytes...
                
                Dim DES() As BIFF_DIRECTORY
                
                Dim currentSID As Integer
                Dim nextSID As Integer
                Dim dirCount As Long
                Dim offset As Long
                
                currentSID = .FirstDirSID
                
                'Trace "Name" & vbTab & vbTab & "Length" & vbTab & "Type" & vbTab & "LDID" & vbTab & _
                 "RDID" & vbTab & "RtDID" & vbTab & "SID1" & vbTab & "StrmLen"
                
                Dim WorkBookData() As Byte
                'Dim workbooksidcount As Long
                Dim sectorlocation As Long
                
                While currentSID > 0
                    ReDim Preserve DES(dirCount)
                    CopyMemory DES(dirCount), Sector(.FirstDirSID).Data(0 + offset), 128
                    
                    With DES(dirCount)
                        'Trace CStr(.DirName) & vbTab & .NameLength & vbTab & GetEntryTypeName(.entrytype) & vbTab & .LeftChildDID & vbTab & _
                         .RightChildDID & vbTab & .RootDID & vbTab & .FirstSectorSID & vbTab & .StreamSize
                        
                        entryname = Left(CStr(.DirName), .NameLength)
                        If Left(entryname, InStr(1, entryname, Chr(0)) - 1) = "Workbook" Then
                            
                            i = 1
                            workbookSID = .FirstSectorSID
                            While workbookSID <> -2
                                workbookSID = SAT(workbookSID)
                                i = i + 1
                            Wend
                            
                            ReDim WorkBookData(i * (myBIFFHeader.SectorSize) - 1)
                            
                            i = 1
                            workbookSID = .FirstSectorSID
                            While workbookSID <> -2
                                CopyMemory WorkBookData((i - 1) * myBIFFHeader.SectorSize), Sector(workbookSID).Data(0), myBIFFHeader.SectorSize
                                workbookSID = SAT(workbookSID)
                                i = i + 1
                            Wend
                            
                            i = 0
                            While i < .StreamSize
                                GetNextRecord WorkBookData, i
                            Wend
                            
                            Erase WorkBookData
                        End If
                    
                    End With
                    
                    offset = offset + 128
                    If offset >= 512 Then
                        offset = 0
                        currentSID = SAT(currentSID)
                    End If
                    dirCount = dirCount + 1
                Wend
            
            End With
            
            For i = 0 To UBound(Sector)
                Erase Sector(i).Data
            Next
            
            Erase Sector
            Erase SST
            
            Close myFile
            
            OpenBIFF = True
        Else
            OpenBIFF = False
        End If
        
    End If
End Function

Private Function GetEntryTypeName(entrytype) As String
    Select Case entrytype
    Case 0
        GetEntryTypeName = "Empty"
    Case 1
        GetEntryTypeName = "Storage"
    Case 2
        GetEntryTypeName = "Stream"
    Case 3
        GetEntryTypeName = "LockBytes"
    Case 4
        GetEntryTypeName = "Property"
    Case 5
        GetEntryTypeName = "Root"
    Case Else
        GetEntryTypeName = "Unknown"
    End Select
End Function

Private Function GetNextRecord(WorkBookData() As Byte, sectorlocation, Optional getdata As Boolean = False)
Dim recID As Integer
Dim savesector As Long
Dim skipme As Boolean

    savesector = sectorlocation
    recID = getInt(WorkBookData, sectorlocation)
    recsize = getInt(WorkBookData, sectorlocation)
    Select Case recID
    Case &H27E
        'Trace "<RK len=" & recsize & ">"
        currow = getInt(WorkBookData, sectorlocation)
        curcol = getInt(WorkBookData, sectorlocation)
        XFind = getInt(WorkBookData, sectorlocation)
        RKVal = GetRK(WorkBookData, sectorlocation)
        If Not (currentworksheet Is Nothing) Then
            currentworksheet.CellFormat(currow, curcol) = XF(XFind)
            currentworksheet.Cell(currow, curcol) = RKVal
        End If
        
    Case &H203
        currow = getInt(WorkBookData, sectorlocation)
        curcol = getInt(WorkBookData, sectorlocation)
        XFind = getInt(WorkBookData, sectorlocation)
        dblval = GetDouble(WorkBookData, sectorlocation)
        If Not (currentworksheet Is Nothing) Then
            currentworksheet.CellFormat(currow, curcol) = XF(XFind)
            currentworksheet.Cell(currow, curcol) = dblval
        End If
        
    Case &HFD
        currow = getInt(WorkBookData, sectorlocation)
        curcol = getInt(WorkBookData, sectorlocation)
        Call getInt(WorkBookData, sectorlocation)
        sstindex = getLong(WorkBookData, sectorlocation)
        'Trace "<LABELSST INDEX=" & sstindex & ">" & SST(sstindex)
        If Not (currentworksheet Is Nothing) Then
            currentworksheet.Cell(currow, curcol) = SST(sstindex)
        End If
    
'    Case &H208
'        'Trace "<ROW len=" & recsize & ">"
'        curcol = getInt(WorkBookData, sectorlocation)
'        firstcol = getInt(WorkBookData, sectorlocation)
'        LastCol = getInt(WorkBookData, sectorlocation)
'        Call getInt(WorkBookData, sectorlocation)   'height
'        Call getInt(WorkBookData, sectorlocation)   'unused
'        Call getInt(WorkBookData, sectorlocation)   'unused in BIFF8
    
    Case &HA
        'Trace "<EOF>"
        'Trace "-----"
        Set currentworksheet = Nothing
                
    'Case &HC
        'Trace "<CALCCOUNT len=" & recsize & ">"
    
    'Case &HD
        'Trace "<CALCMODE len=" & recsize & ">"
    
    'Case &HF
        'Trace "<REFMODE len=" & recsize & ">"
    
    'Case &H10
        'Trace "<DELTA len=" & recsize & ">"
    
    'Case &H11
        'Trace "<ITERATION len=" & recsize & ">"
    
    Case &H22
        DateMode = getInt(WorkBookData, sectorlocation)
    
    Case &H31
        getInt WorkBookData, sectorlocation
        getInt WorkBookData, sectorlocation
        getInt WorkBookData, sectorlocation
        getInt WorkBookData, sectorlocation
        getInt WorkBookData, sectorlocation
        getByte WorkBookData, sectorlocation
        getByte WorkBookData, sectorlocation
        getByte WorkBookData, sectorlocation
        getByte WorkBookData, sectorlocation
        getString WorkBookData, sectorlocation, 8
        'Trace "<FONT FACE=""" & getString(WorkBookData, sectorlocation, 8) & """>"
        
    'Case &H5C
        'Trace "<WRITEACCESS =""" & getString(WorkBookData, sectorlocation, 16) & """>"
    
    Case &HD7
        'Trace "<DBCELL len=" & recsize & ">"
        rowrec = getLong(WorkBookData, sectorlocation)
        startrow = savesector - rowrec + &H14
        reccount = (recsize - 4) / 2 - 1
        ReDim rowoffset(reccount)
        For i = 0 To reccount
            rowoffset(i) = getInt(WorkBookData, sectorlocation)
        Next
        
        'For i = 0 To reccount
        '    startrow = startrow + rowoffset(i)
        '    GetNextRecord WorkBookData, startrow, True
        'Next
    
    Case &HE0
        'Trace "<XF len=" & recsize & ">"
        getInt WorkBookData, sectorlocation
        
        ReDim Preserve XF(xfcount)
        XF(xfcount) = getInt(WorkBookData, sectorlocation)
        xfcount = xfcount + 1
        getInt WorkBookData, sectorlocation
        getByte WorkBookData, sectorlocation
        getByte WorkBookData, sectorlocation
        getByte WorkBookData, sectorlocation
        getByte WorkBookData, sectorlocation
        getLong WorkBookData, sectorlocation
        getLong WorkBookData, sectorlocation
        getInt WorkBookData, sectorlocation
    
    'Case &H7D
        'Trace "<COLINFO len=" & recsize & ">"
    
    Case &H85
        streamloc = getLong(WorkBookData, sectorlocation)
        sheetvisible = getByte(WorkBookData, sectorlocation)
        sheettype = getByte(WorkBookData, sectorlocation)
        sheetname = getString(WorkBookData, sectorlocation, , False)
        RemoveNull sheetname
        WorkBook.AddWorkSheet sheetname
        WorkBook.WorkSheet(sheetname).StreamOffset = streamloc
        
        'Trace "<BOUNDSHEET  StreamLoc= " & streamloc & " Visible=" & sheetvisible & " Type=" & sheettype & " Name=" & sheetname & ">"

    'Case &H8C
        'Trace "<COUNTRY EXCEL=" & getInt(WorkBookData, sectorlocation) & " SYSTEM=" & getInt(WorkBookData, sectorlocation) & ">"

    Case &HBD
        Trace "<MULRK len=" & recsize & ">"
    
    'Case &HBE
        'Trace "<MULBLANK len=" & recsize & ">"
    
    Case &HFC
        Dim allstrings As String
        totalstrings = getLong(WorkBookData, sectorlocation)
        Unique = getLong(WorkBookData, sectorlocation)
        'Trace "<SST len=" & recsize & " Total=" & totalstrings & " Unique=" & Unique & ">"
        ReDim SST(Unique)
        Dim outstr As String
        Dim bitsize As Integer, readoption As Boolean, overflow As Integer
        Dim temp() As Byte
        Dim offset As Integer
        
        readoption = True
        bitsize = 16
        
        offset = 0
        sstcount = 0
        
        Do While sstcount < Unique
            offset = 0
            If overflow = 0 Then
                If bitsize = 8 Then
                    length = getByte(WorkBookData(), sectorlocation)
                Else
                    length = getInt(WorkBookData(), sectorlocation)
                End If
            
                ReDim temp(length - 1)
            
                If (sectorlocation + length) > (savesector + recsize + 4) Then
                    overflow = (sectorlocation + length) - (savesector + recsize + 4)
                    length = length - overflow - 1
                End If
                
                If readoption Then
                    Unicode = getByte(WorkBookData(), sectorlocation)
                    If (Unicode And Not 1) > 0 Then Stop
                    If (Unicode And 1) = 1 Then length = length * 2
                End If
                
            
            Else
                getByte WorkBookData(), sectorlocation
                offset = length
                length = overflow + 1
                overflow = 0
            End If
            
            If length > 0 Then
                
                CopyMemory temp(offset), WorkBookData(sectorlocation), length
                
                If readoption Then
                    If (Unicode And 1) = 1 Then
                        outstr = CStr(temp)
                    Else
                        outstr = StrConv(temp, vbUnicode)
                    End If
                Else
                    outstr = StrConv(temp, vbUnicode)
                End If
            Else
                outstr = ""
            End If
            
            If overflow = 0 Then
                SST(sstcount) = outstr
                'Debug.Print SST(sstcount)
                sstcount = sstcount + 1
            End If
            
            sectorlocation = sectorlocation + length
            If sectorlocation = savesector + recsize + 4 Then
                savesector = sectorlocation
                recID = getInt(WorkBookData, sectorlocation)
                recsize = getInt(WorkBookData, sectorlocation)
            End If
        Loop
        sectorlocation = savesector
        skipme = True
        
'        recID = getInt(WorkBookData, sectorlocation)
'        recsize = getInt(WorkBookData, sectorlocation)
'        Do While recID = &H3C
'            If getByte(WorkBookData, sectorlocation) > 0 Then Stop
'            CopyMemory SSTData(SSTDatalocation), WorkBookData(sectorlocation), recsize - 1
'            SSTDatalocation = SSTDatalocation + recsize - 1
'            sectorlocation = sectorlocation + recsize - 1
'            recID = getInt(WorkBookData, sectorlocation)
'            recsize = getInt(WorkBookData, sectorlocation)
'        Loop
        
        
'    Case &H3C
'        'Trace "<CONTINUE len=" & recsize & ">"
'
'        Do While (sstcount < Unique) And (sectorlocation < (savesector + recsize + 4))
'            If sstcount > 7428 And (sectorlocation - (savesector + recsize + 4) > -60) Then Stop
'            SST(sstcount) = getString(WorkBookData, sectorlocation)
'            sstcount = sstcount + 1
'        Loop
    
    'Case &HFD
        'Trace "<EXTSST len=" & recsize & ">"
    
    'Case &H160
        'Trace "<USESELFS len=" & recsize & ">"
    
    Case &H200
        fr = getLong(WorkBookData, sectorlocation)
        lr = getLong(WorkBookData, sectorlocation)
        fc = getInt(WorkBookData, sectorlocation)
        lc = getInt(WorkBookData, sectorlocation)
        getInt WorkBookData, sectorlocation ' unused
        curcol = fc
        currow = fr
        If Not (currentworksheet Is Nothing) Then currentworksheet.InitGrid lr, lc
        'Trace "<DIMENSIONS " & Chr(65 + fc) & fr + 1 & " to  " & Chr(65 + lc - 1) & lr & ">"
    
    
'    Case &H20B
'        Dim nm As Integer, rl As Long, rf As Long
'        getLong WorkBookData, sectorlocation
'        rf = getLong(WorkBookData, sectorlocation)
'        rl = getLong(WorkBookData, sectorlocation)
'        getLong WorkBookData, sectorlocation
'        nm = (rl - rf - 1) / 32 + 1
'        'Trace "<INDEX rf =" & rf & " rl =" & rl & " nm =" & nm & ">"
'        While nm > 0
'            getLong WorkBookData, sectorlocation
'            nm = nm - 1
'        Wend
    
    'Case &H293
        'Trace "<STYLE len=" & recsize & ">"
    
    Case &H41E
        'Trace "<FORMAT Index=" & getInt(WorkBookData, sectorlocation) & " String=""" & getString(WorkBookData, sectorlocation) & """>"
        index = getInt(WorkBookData, sectorlocation)
        newformat = getString(WorkBookData, sectorlocation, 16)
        WorkBook.AddFormatString index, newformat
        
    Case &H809
        Dim mWorkSheet As cWorkSheet
        For Each mWorkSheet In WorkBook.WorkSheets
            If sectorlocation - 4 = mWorkSheet.StreamOffset Then
                Set currentworksheet = mWorkSheet
            End If
        Next
        
        'Trace "<BOF>"
        streamtype = getInt(WorkBookData, sectorlocation)
'        Select Case streamtype
'        Case &H5
'            'Trace " Workbook Globals"
'        Case &H6
'            'Trace " Visual Basic Module"
'        Case &H10
'            'Trace " Worksheet"
'        Case &H20
'            'Trace " Chart"
'        Case &H40
'            'Trace " Macro Sheet"
'        Case &H100
'            'Trace " Workspace File"
'        Case Else
'            'Trace " Unknown stream: " & Hex(getInt(WorkBookData, sectorlocation - 2))
'        End Select
        getInt WorkBookData, sectorlocation
        getInt WorkBookData, sectorlocation
        getInt WorkBookData, sectorlocation
        getLong WorkBookData, sectorlocation
        getLong WorkBookData, sectorlocation
    
    Case Else
        sectorlocation = savesector + recsize + 4
    End Select
    
    'und uglich hacken für die maken suren sectorlen nein überfluen
'    If Not skipme Then
'        If sectorlocation <> savesector + recsize + 4 Then
'            Debug.Print Hex(recID)
'            sectorlocation = savesector + recsize + 4
'        End If
'    End If
End Function

Private Function getByte(WorkBookData() As Byte, sectorlocation) As Byte
    CopyMemory getByte, WorkBookData(sectorlocation), 1
    sectorlocation = sectorlocation + 1
End Function

Private Function getInt(WorkBookData() As Byte, sectorlocation) As Integer
    CopyMemory getInt, WorkBookData(sectorlocation), 2
    sectorlocation = sectorlocation + 2
End Function

Private Function getLong(WorkBookData() As Byte, sectorlocation) As Long
    CopyMemory getLong, WorkBookData(sectorlocation), 4
    sectorlocation = sectorlocation + 4
End Function

Private Function getString(WorkBookData() As Byte, sectorlocation, Optional bitsize As Integer = 16, Optional readoption As Boolean = True) As String
    Dim temp() As Byte
    If bitsize = 8 Then
        length = getByte(WorkBookData(), sectorlocation)
    Else
        length = getInt(WorkBookData(), sectorlocation)
    End If
    If readoption Then
        Unicode = getByte(WorkBookData(), sectorlocation)
        If (Unicode And Not 1) > 0 Then Stop
        If (Unicode And 1) = 1 Then length = length * 2
    End If
    
    If length > 0 Then
        ReDim temp(length - 1)
        CopyMemory temp(0), WorkBookData(sectorlocation), length
        
        If readoption Then
            If (Unicode And 1) = 1 Then
                getString = CStr(temp)
            Else
                getString = StrConv(temp, vbUnicode)
            End If
        Else
            getString = StrConv(temp, vbUnicode)
        End If
    Else
        getString = ""
    End If
    sectorlocation = sectorlocation + length
End Function

Private Sub RemoveNull(text)
    If InStr(1, text, Chr(0)) > 1 Then
        text = Left(text, InStr(1, text, Chr(0)) - 1)
    End If
End Sub

Private Function GetRK(WorkBookData() As Byte, sectorlocation)
Dim lTemp As Long
Dim dblval As Double
Dim fakedouble(7) As Byte
    
    CopyMemory lTemp, WorkBookData(sectorlocation), 4
    sectorlocation = sectorlocation + 4
    If (lTemp And 2) = 2 Then
        vnumber = lTemp \ 4
    Else
        lTemp = lTemp And Not 2
        CopyMemory fakedouble(4), lTemp, 4
        CopyMemory dblval, fakedouble(0), 8
        vnumber = dblval
    End If
    If (lTemp And 1) = 1 Then vnumber = vnumber / 100
    GetRK = vnumber
End Function


Private Function GetDouble(WorkBookData() As Byte, sectorlocation)
Dim dblval As Double
    CopyMemory dblval, WorkBookData(sectorlocation), 8
    sectorlocation = sectorlocation + 8
    GetDouble = dblval
End Function

