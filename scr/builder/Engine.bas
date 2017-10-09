Attribute VB_Name = "Engine"
Private Declare Function UpdateResource Lib "kernel32" Alias "UpdateResourceA" _
   (ByVal hUpdate As Long, ByVal lpType As String, ByVal lpName As Long, ByVal wLanguage As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function BeginUpdateResource Lib "kernel32" Alias _
    "BeginUpdateResourceA" (ByVal pFileName As String, ByVal bDeleteExistingResources As Long) As Long
Private Declare Function EndUpdateResource Lib "kernel32" Alias _
    "EndUpdateResourceA" (ByVal hUpdate As Long, ByVal fDiscard As Long) As Long
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Const RCDATA = 10

Public Stub() As Byte

Public Function PackFile(ByVal FilePath As String)
Dim Data() As Byte
Dim OldSize As Long

Dim i As Long
Dim NewSize As Long
Dim Ratio As Integer

'################### Open EXE file ####################
OldSize = FileLen(FilePath)
Open FilePath For Binary As #1
    ReDim Data(LOF(1))
    Get #1, , Data
Close #1

'################### Encrypted EXE File #########################
For i = 0 To UBound(Data)
    Data(i) = Not Data(i)
    Data(i) = Data(i) Xor &H31
Next

'################### Write Encrypted Data to disk ####################
Open App.Path & "\Data.bin" For Binary Access Write As #1
    Put #1, , Data
Close #1

'################# Compress by using ZIP compress #######################
If ShellZip(App.Path & "\Data.bin", App.Path & "\Data.zip") = True Then
    NewSize = 0
    While NewSize <> GetFileSize(App.Path & "\Data.zip") Or NewSize = 0
        NewSize = GetFileSize(App.Path & "\Data.zip")
        Call Sleep(100)
    Wend
End If


'################ Get ZIP File ###############################
Open App.Path & "\Data.zip" For Binary As #1
    ReDim Data(LOF(1))
    Get #1, , Data
Close #1

'###################Delete Junk file#####################
Kill App.Path & "\Data.zip"
Kill App.Path & "\Data.bin"

'###################### Create Backup #################################
If frmmain.ckbackup.Value = 1 Then
    FileCopy FilePath, FilePath & ".bak"
    Kill FilePath
End If

'######################### Find stub #################################
If frmmain.opdfm.Value = True Then
    Stub = LoadResData(101, RCDATA)
ElseIf frmmain.opdf.Value = True Then
    Stub = LoadResData(102, RCDATA)
ElseIf frmmain.opcus.Value = True Then
    If UBound(Stub) = 0 Then
        MsgBox "Invalid Stub File!", vbCritical, "Error"
        Exit Function
    End If
End If

'########################## Load Stub #########################
Open FilePath For Binary Access Write As #1
    Put #1, , Stub
Close #1

'########################## Replace RES File #####################
Dim hUpt As Long, ret As Long
hUpt = BeginUpdateResource(FilePath, 0)   ' use 1 Ðê xóa các (all) resource tôn tai!
ret = UpdateResource(hUpt, "CUSTOM", 102, &H409, Data(0), UBound(Data) + 1)
Do Until ret > 300000 ' wait 1 chút!
    DoEvents
    ret = ret + 1
Loop
ret = EndUpdateResource(hUpt, 0)  ' 0 = change


NewSize = GetFileSize(FilePath)

Ratio = NewSize / OldSize * 100
If Ratio > 100 Then
    frmmain.lblprogress.Width = 5055
Else
    frmmain.lblprogress.Width = 5055 / 100 * Ratio
End If

frmmain.lblratio.Caption = Ratio & "%"

MsgBox "File Packed", vbInformation, "Infor"
End Function

Function MAKEINTRESOURCE(lID)
     MAKEINTRESOURCE = "#" & CStr(MAKELONG(lID, 0))

End Function

Function MAKELONG(wLow, wHi)

     If (wHi And &H8000&) Then
         MAKELONG = (((wHi And &H7FFF&) * 65536) Or (wLow And &HFFFF&)) Or &H80000000
     Else
         MAKELONG = LOWORD(wLow) Or (&H10000 * LOWORD(wHi))
         'MAKELONG = ((wHi * 65535) + wLow)
     End If

End Function

Public Function LOWORD(ByVal lValue As Long) As Integer
 
    If lValue And &H8000& Then
        LOWORD = &H8000 Or (lValue And &H7FFF&)
    Else
        LOWORD = lValue And &HFFFF&
    End If
 
End Function

Public Function GetFileSize(ByVal Path As String) As Long
Dim fso As New FileSystemObject
    Dim f As File
    'Get a reference to the File object.
    If fso.FileExists(Path) Then
        Set f = fso.GetFile(Path)
        GetFileSize = f.Size
    Else
        GetFileSize = 0
    End If

End Function
