VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Begin VB.Form FRM_IMPORT_LOG 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "USB IMPORT LOG"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      _Version        =   851970
      _ExtentX        =   15478
      _ExtentY        =   8493
      _StockProps     =   79
      ForeColor       =   16711680
      BackColor       =   12648384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin VB.FileListBox File1 
         Height          =   2820
         Left            =   4680
         TabIndex        =   3
         Top             =   840
         Width           =   3735
      End
      Begin VB.DirListBox Dir1 
         Height          =   3690
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   4095
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   4095
      End
      Begin XtremeSuiteControls.PushButton cmdProcess 
         Height          =   735
         Left            =   4680
         TabIndex        =   4
         Top             =   3720
         Width           =   3735
         _Version        =   851970
         _ExtentX        =   6588
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "UPLOAD "
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.ComboBox CMB_Lokasi 
         Height          =   360
         Left            =   4680
         TabIndex        =   5
         Top             =   360
         Width           =   3735
         _Version        =   851970
         _ExtentX        =   6588
         _ExtentY        =   635
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "ComboBox2"
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "* Path"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   7
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "* Lokasi Machine Absen"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   4680
         TabIndex        =   6
         Top             =   120
         Width           =   2535
      End
   End
End
Attribute VB_Name = "FRM_IMPORT_LOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strEmpFileName  As String
Dim strBackSlash  As String
Dim intEmpFileNbr As Integer
Dim strEmpName As String, strEmpName1 As String
Dim pathNm As String

Private Sub Command1_Click()
usbimport
End Sub

Private Sub cmdProcess_Click()
usbimport
End Sub

Private Sub Dir1_Change()
    File1.Pattern = "*.dat"
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = left$(Drive1.Drive, 1) & ":"
End Sub

Private Sub File1_Click()
On Error Resume Next
     pathNm = Dir1.Path & "\" & File1.Filename
End Sub

Private Sub Form_Load()
File1.Pattern = "*.dat"
'lR = SetTopMostWindow(Me.hwnd, True)
Main
prjSysID.FingerMachineNoAll CMB_Lokasi
End Sub

Private Sub usbimport()
    Dim CLM2 As Integer, clm3 As String, clm4 As Integer, clm5 As Integer, clm6 As Integer, clm7 As Integer
    Dim Tgl As Date
    Dim fTempat, Key, FngerId As Integer
    Dim AryDataUSB  As String
    Dim Qry As String
    'strBackSlash = IIf(right$(App.Path, 1) = "\", "", "\")
    'strEmpFileName = App.Path & strBackSlash & "1_attlog.dat"
    intEmpFileNbr = FreeFile
     strEmpFileName = pathNm
     'MsgBox strEmpFileName
    Open strEmpFileName For Input As #intEmpFileNbr
    
    Do Until EOF(intEmpFileNbr)
        Input #intEmpFileNbr, strEmpName
         'strEmpName1 = Left(strEmpName, Len(strEmpName) - 28)
         clm7 = right(left(strEmpName, Len(strEmpName)), 1)
         clm6 = right(left(strEmpName, Len(strEmpName) - 2), 1)
         clm5 = right(left(strEmpName, Len(strEmpName) - 4), 1)
         clm4 = right(left(strEmpName, Len(strEmpName) - 6), 1)
         clm3 = right(left(strEmpName, Len(strEmpName) - 8), 20)
         CLM2 = left(strEmpName, Len(strEmpName) - 28)
         Tgl = clm3
         Key = clm5
         FngerId = CLM2
         fTempat = CMB_Lokasi.Text
        'Print clm2, clm3, clm4, clm5, clm6
        'Print fTempat, FngerId, Key, Format(Tgl, "yyyy-mm-dd HH:mm:ss")
        AryDataUSB = fTempat & _
                    "=" & FngerId & _
                    "=" & Key & _
                    "=" & Format(Tgl, "yyyy-mm-dd HH:mm:ss")
        'Print strEmpName1
        Qry = "pro_get_personal_log_USB('" & AryDataUSB & "')"
        OpRecStt2 Qry, True
       
        
    Loop
    Close #intEmpFileNbr
     Set rsStt2 = Nothing
End Sub

