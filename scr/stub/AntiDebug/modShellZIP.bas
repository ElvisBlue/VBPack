Attribute VB_Name = "modShellZIP"


Option Explicit

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

