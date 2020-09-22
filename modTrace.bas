Attribute VB_Name = "modTrace"
Public sTrace As String
Dim lTrace As Long
Dim maxtrace As Long

Public Sub InitTrace()
    lTrace = 1
    maxtrace = 5000
    sTrace = String(maxtrace, Chr(0))
End Sub

Public Sub Trace(ByVal text As String)
Dim ltext As Long
    text = Replace$(text, Chr(0), "") & vbCrLf
    ltext = Len(text)
    
    If (lTrace + ltext) > maxtrace Then
        sTrace = sTrace & String$(maxtrace, Chr(0))
        maxtrace = Len(sTrace)
    End If
    
    Mid$(sTrace, lTrace, ltext) = text
    lTrace = lTrace + ltext
    DoEvents
End Sub

Public Sub DumpTextBox(obj As TextBox)
    MsgBox "Done!", vbInformation
    'obj.text = sTrace
End Sub

Public Sub DumpFile(filename As String)
    sTrace = Left$(sTrace, lTrace)
    Open filename For Output As 1
    Print #1, sTrace
    Close 1
End Sub
