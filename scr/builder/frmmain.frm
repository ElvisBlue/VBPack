VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VBPack 1.0 - Yep another RunPE crap"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   10335
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Stub"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   5640
      TabIndex        =   9
      Top             =   120
      Width           =   4575
      Begin VB.CommandButton cmdbrowser2 
         Caption         =   "..."
         Height          =   255
         Left            =   3720
         TabIndex        =   14
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtpath2 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1440
         Width           =   3255
      End
      Begin VB.OptionButton opcus 
         Caption         =   "Custom Stub"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton opdfm 
         Caption         =   "Default Stub + manifest"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   2535
      End
      Begin VB.OptionButton opdf 
         Caption         =   "Default Stub"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.Label lblstubinfo 
         Caption         =   "This stub give you best compression."
         Height          =   975
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   4335
      End
   End
   Begin VB.CommandButton cmdabout 
      Caption         =   "About"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CheckBox ckbackup 
      Caption         =   "Create Backup"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdpack 
      Caption         =   "Go!"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   2640
      Width           =   1575
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
      BackStyle       =   0  'Transparent
      Caption         =   "Compress Ratio:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblratio 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lblprogress 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   5055
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   5055
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdabout_Click()
MsgBox "Author: Elvis" & vbCrLf & _
        "Language: VB6" & vbCrLf & _
        "Program Version 1.0 (Public Version)" & vbCrLf & vbCrLf & _
        "- VBPack isn't a crypter" & vbCrLf & _
        "- VBPack only support exe file" & vbCrLf & _
        "- VBPack does not protect your program as well as Themida, VmProtect, Enigma,..." & vbCrLf & _
        "- VBPack does not compress your program as well as UPX, AsPack, PECompact....." & vbCrLf & vbCrLf & _
        "+ VBPack is only for fun!", vbInformation, "Holly shit!"
End Sub

Private Sub cmdbrowser_Click()
CommonDialog1.Filter = "Execute File (*.exe)|*.exe|All files (*.*)|*.*"
CommonDialog1.DialogTitle = "Select File To Pack"
CommonDialog1.ShowOpen
txtpath.Text = CommonDialog1.FileName
End Sub

Private Sub cmdbrowser2_Click()
CommonDialog1.Filter = "Stub File File (*.DAT)|*.DAT|All files (*.*)|*.*"
CommonDialog1.DialogTitle = "Select File To Pack"
CommonDialog1.ShowOpen
txtpath2.Text = CommonDialog1.FileName
Call opcus_Click
End Sub

Private Sub cmdpack_Click()
If txtpath.Text = vbNullString Then Exit Sub
Call PackFile(txtpath.Text)
End Sub

Private Sub opcus_Click()
Dim Data As String
Dim Data2() As String
Dim StubSize As Long

If txtpath2.Text = vbNullString Then
    lblstubinfo.Caption = "Custom stub not found!"
    ReDim Stub(0)
    Exit Sub
Else
    StubSize = FileLen(txtpath2.Text)
    Data = Space(StubSize)
    Open txtpath2.Text For Binary As #1
        Get #1, , Data
    Close #1
End If
Data2 = Split(Data, "-%E%-")
If UBound(Data2) <> 1 Then
    MsgBox "Invalid Custom Stub", vbCritical, "Error"
    Exit Sub
End If

lblstubinfo.Caption = Data2(0)
ReDim Stub(StubSize - Len(Data2(0)) - 5)
Stub = StrConv(Data2(1), vbFromUnicode)
End Sub

Private Sub opdfm_Click()
lblstubinfo.Caption = "This stub give you best compression and also add manifest to your file. Be careful because manifest may cause error in some cases"
End Sub

Private Sub opdf_Click()
lblstubinfo.Caption = "This stub give you best compression."
End Sub

