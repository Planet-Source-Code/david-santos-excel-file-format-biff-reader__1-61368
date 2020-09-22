VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWorkBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_WorkSheet As New Collection
Private m_UserName As String

Private Type BIFF_FORMAT
    index As Integer
    FormatString As String
End Type

Private FormatList() As BIFF_FORMAT
Private formatcount As Long

Public Function AddFormatString(index, newformat)
    ReDim Preserve FormatList(formatcount)
    FormatList(formatcount).index = index
    FormatList(formatcount).FormatString = newformat
    formatcount = formatcount + 1
End Function

Public Function GetFormatString(index) As String
    For i = 0 To UBound(FormatList)
        If FormatList(i).index = index Then
            GetFormatString = FormatList(i).FormatString
            Exit For
        End If
    Next
End Function

Property Get WorkSheet(index) As cWorkSheet
    Set WorkSheet = m_WorkSheet(index)
End Property

Property Get WorkSheets() As Collection
    Set WorkSheets = m_WorkSheet
End Property

Public Sub AddWorkSheet(name)
Dim NewWorkSheet As New cWorkSheet
    NewWorkSheet.name = name
    m_WorkSheet.Add NewWorkSheet, name
    Set NewWorkSheet.Parent = Me
End Sub

Property Let UserName(name As String)
    m_UserName = name
End Property

Property Get UserName() As String
    UserName = m_UserName
End Property

Private Sub Class_Terminate()
    While m_WorkSheet.Count > 0
        m_WorkSheet.Remove 1
    Wend
End Sub
