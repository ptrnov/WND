VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Begin VB.Form FRM_ABSENSI_PBAR 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employe Log IN/OUT"
   ClientHeight    =   1020
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4680
      Top             =   480
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBar1 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4455
      _Version        =   851970
      _ExtentX        =   7858
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   -2147483633
      Value           =   1
      Appearance      =   1
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   2295
      _Version        =   851970
      _ExtentX        =   4048
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "COMPLETED"
      ForeColor       =   255
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
      Alignment       =   2
   End
End
Attribute VB_Name = "FRM_ABSENSI_PBAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
init_Frm
ProgressBar1(0).Visible = True
Screen.MousePointer = vbHourglass
Me.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
'lR = SetTopMostWindow(FRM_ABSENSI.hwnd, True)
FRM_ABSENSI.Enabled = True
End Sub

Private Sub init_Frm()
'Picture1.Height = 1520
'Picture1.Width = 4750
lR = SetTopMostWindow(FRM_ABSENSI.hwnd, False)
FRM_ABSENSI.Enabled = False
lR = SetTopMostWindow(Me.hwnd, True)
End Sub
