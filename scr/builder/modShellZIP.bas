Attribute VB_Name = "modShellZIP"


Option Explicit

'Asynchronously compresses a file or folder. Result differs if folder has a trailing backslash ("\").
Public Function ShellZip(ByRef Source As String, ByRef DestZip As String) As Boolean
    CreateNewZip DestZip

    On Error Resume Next
    With CreateObject("Shell.Application")  'Late-bound
   'With New Shell                          'Referenced
        If Right$(Source, 1&) = "\" Then
            .NameSpace(CVar(DestZip)).CopyHere .NameSpace(CVar(Source)).Items
        Else
            .NameSpace(CVar(DestZip)).CopyHere CVar(Source)
        End If
    End With

    ShellZip = (Err = 0&)
End Function

'Asynchronously decompresses the contents of SrcZip into the folder DestDir.
Public Function ShellUnzip(ByRef SrcZip As String, ByRef DestDir As String) As Boolean
    On Error Resume Next
    With CreateObject("Shell.Application")  'Late-bound
   'With New Shell                          'Referenced
        .NameSpace(CVar(DestDir)).CopyHere .NameSpace(CVar(SrcZip)).Items
    End With

    ShellUnzip = (Err = 0&)

    RemoveTempDir Right$(SrcZip, Len(SrcZip) - InStrRev(SrcZip, "\"))
End Function

'Creates a new empty Zip file only if it doesn't exist.
Private Function CreateNewZip(ByRef sFileName As String) As Boolean
    With CreateObject("Scripting.FileSystemObject")  'Late-bound
   'With New FileSystemObject                        'Referenced
        On Error GoTo 1
        With .CreateTextFile(sFileName, Overwrite:=False)
            .Write "PK" & Chr$(5&) & Chr$(6&) & String$(18&, vbNullChar)
            .Close
1       End With
    End With

    CreateNewZip = (Err = 0&)
End Function

'Schedules a temporary directory tree for deletion upon reboot.
Private Function RemoveTempDir(ByRef sFolderName As String) As Boolean
    Dim sPath As String, sTemp As String

    On Error Resume Next
    sTemp = Environ$("TEMP") & "\"
    sPath = Dir(sTemp & "Temporary Directory * for " & sFolderName, vbDirectory Or vbHidden)

    If LenB(sPath) Then
        With CreateObject("WScript.Shell")  'Late-bound
       'With New WshShell                   'Referenced
            Do: .RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce\*RD_" & _
                           Replace(sPath, " ", "_"), Environ$("ComSpec") & " /C " & _
                          "@TITLE Removing " & sPath & " ...&" & _
                          "@RD /S /Q """ & sTemp & sPath & """"
                 sPath = Dir
            Loop While LenB(sPath)
        End With
    End If

    RemoveTempDir = (Err = 0&)
End Function

