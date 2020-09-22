VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   5295
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   8775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m As BIFFReader
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub Form_Load()
Dim filename As String
    
    Me.Show
    Set m = New BIFFReader
    s = GetTickCount
    filename = App.Path & "\test.xls"
        
    If m.OpenBIFF(filename) Then
        Text1.text = Text1.text & "Parsed """ & filename & """ in " & Round((GetTickCount - s) / 1000, 2) & " seconds." & vbCrLf
        DoEvents
        
        s = GetTickCount
        m.WorkSheet(1).SaveAs App.Path & "\test.csv"
        Text1.text = Text1.text & "Wrote CSV in " & Round((GetTickCount - s) / 1000, 2) & " seconds." & vbCrLf
        DoEvents
    Else
        MsgBox "File header not found. Maybe not a BIFF8/8X file?", vbExclamation
    End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m = Nothing
End Sub


