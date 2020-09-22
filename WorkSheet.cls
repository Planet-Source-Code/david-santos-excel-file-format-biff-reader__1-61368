VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWorkSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_Name As String
Private m_offset As Long
Private m_Cell() As Variant
Private m_Format() As Integer
Private m_Parent As cWorkBook

Public Sub InitGrid(rows, cols)
    ReDim m_Cell(rows, cols)
    ReDim m_Format(rows, cols)
End Sub

Property Get Parent() As cWorkBook
    Parent = m_Parent
End Property

Property Set Parent(oParent As cWorkBook)
   Set m_Parent = oParent
End Property

Property Get LastRow()
    LastRow = UBound(m_Cell, 1)
End Property

Property Get LastCol()
    LastCol = UBound(m_Cell, 2)
End Property

Property Let Cell(row, col, value)
    m_Cell(row, col) = value
End Property

Property Let CellFormat(row, col, value)
    m_Format(row, col) = value
End Property

Property Get CellFormat(row, col) As Variant
    If row < 1 Or col < 1 Then Err.Raise vbObjectError + 1, "Invalid index"
    If Not IsNumeric(col) Then
        If Len(col) = 2 Then
            col = (Asc(Left("AA", 1)) - 64) * 26 + (Asc(Right("AA", 1)) - 64)
        Else
            col = Asc(Right("AA", 1)) - 64
        End If
    End If
    CellFormat = m_Format(row - 1, col - 1)
End Property

Property Get Cell(row, col) As Variant
    If row < 1 Or col < 1 Then Err.Raise vbObjectError + 1, "Invalid index"
    If Not IsNumeric(col) Then
        If Len(col) = 2 Then
            col = (Asc(Left("AA", 1)) - 64) * 26 + (Asc(Right("AA", 1)) - 64)
        Else
            col = Asc(Right("AA", 1)) - 64
        End If
    End If
    Cell = m_Cell(row - 1, col - 1)
End Property

Property Let name(pName As String)
    m_Name = pName
End Property

Property Let StreamOffset(offset As Long)
    m_offset = offset
End Property

Property Get StreamOffset() As Long
    StreamOffset = m_offset
End Property

Property Get name() As String
    name = m_Name
End Property

Public Sub SaveAs(filename As String)
Dim record As String, outstr As String
    Open filename For Output As 1
    For i = 0 To UBound(m_Cell, 1)
        record = ""
        For j = 0 To UBound(m_Cell, 2)
            outstr = m_Cell(i, j)
            If m_Format(i, j) > 0 Then
                checkformat = m_Parent.GetFormatString(m_Format(i, j))
                If Len(checkformat) > 0 Then outstr = Format(outstr, checkformat)
            End If
            If InStr(1, outstr, ",") > 0 Then outstr = """" & outstr & """"
            record = record & outstr & ","
        Next
        Print #1, record
    Next
    Close 1
End Sub

Private Sub Class_Terminate()
    Set m_Parent = Nothing
    Erase m_Cell
    Erase m_Format
End Sub
