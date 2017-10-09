VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custom Stub Generator"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4320
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtinfo 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1080
      Width           =   5295
   End
   Begin VB.CommandButton cmdcreate 
      Caption         =   "Create"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton cmdbrowser 
      Caption         =   "..."
      Height          =   255
      Left            =   4560
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox txtpath 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Stub Infor"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Sub cmdbrowser_Click()
CommonDialog1.Filter = "Execute File (*.exe)|*.exe|All files (*.*)|*.*"
CommonDialog1.DialogTitle = "Select EXE Stub File"
CommonDialog1.ShowOpen
txtpath.Text = CommonDialog1.FileName
End Sub

Private Sub cmdcreate_Click()
Dim Stub() As Byte
Dim DATStub() As Byte
Dim StubSize As Long
Dim StubInfo As String

If txtpath.Text = vbNullString Then Exit Sub

StubSize = FileLen(txtpath.Text)
Open txtpath.Text For Binary Access Read As #1
    ReDim Stub(StubSize) As Byte
    Get #1, , Stub
Close #1

ReDim DATStub(StubSize + Len(txtinfo.Text) + 5) As Byte
StubInfo = txtinfo.Text & "-%E%-"

Call CopyMemory(DATStub(0), ByVal StubInfo, Len(StubInfo))
Call CopyMemory(DATStub(Len(StubInfo)), Stub(0), StubSize)

Open txtpath.Text & ".DAT" For Binary Access Write As #1
    Put #1, , DATStub
Close #1

MsgBox "Stub " & txtpath.Text & ".DAT has been generated", vbInformation, "Sucessful"

End Sub
