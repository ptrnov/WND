VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Begin VB.Form FRM_PAYROLL_REPORT 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "View Report Payroll"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   495
      Left            =   -120
      TabIndex        =   20
      Top             =   2470
      Width           =   7095
      _Version        =   851970
      _ExtentX        =   12515
      _ExtentY        =   873
      _StockProps     =   79
      BackColor       =   16761024
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   45
         Width           =   6975
         _Version        =   851970
         _ExtentX        =   12303
         _ExtentY        =   661
         _StockProps     =   93
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   975
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6975
      _Version        =   851970
      _ExtentX        =   12303
      _ExtentY        =   1720
      _StockProps     =   79
      BackColor       =   16744576
      Appearance      =   1
      Begin XtremeSuiteControls.RadioButton RadioButton1 
         Height          =   255
         Index           =   0
         Left            =   5520
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
         _Version        =   851970
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Weekly"
         ForeColor       =   0
         BackColor       =   16744576
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
      End
      Begin XtremeSuiteControls.ComboBox CmbRpt 
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   2775
         _Version        =   851970
         _ExtentX        =   4895
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Sorted          =   -1  'True
         AutoComplete    =   -1  'True
         DropDownWidth   =   5000
      End
      Begin XtremeSuiteControls.ComboBox CmbRpt 
         Height          =   315
         Index           =   3
         Left            =   2880
         TabIndex        =   15
         Top             =   480
         Width           =   2535
         _Version        =   851970
         _ExtentX        =   4471
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Sorted          =   -1  'True
         AutoComplete    =   -1  'True
         DropDownWidth   =   5000
      End
      Begin XtremeSuiteControls.RadioButton RadioButton1 
         Height          =   255
         Index           =   1
         Left            =   5520
         TabIndex        =   19
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
         _Version        =   851970
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Monthly"
         ForeColor       =   0
         BackColor       =   16744576
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
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF8080&
         Caption         =   "Cabang."
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   2880
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF8080&
         Caption         =   "List Report."
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox3 
      Height          =   1515
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   6975
      _Version        =   851970
      _ExtentX        =   12303
      _ExtentY        =   2672
      _StockProps     =   79
      ForeColor       =   255
      BackColor       =   12640511
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      Begin XtremeSuiteControls.DateTimePicker RptTgl 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   400
         Width           =   1455
         _Version        =   851970
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   68
      End
      Begin XtremeSuiteControls.DateTimePicker RptTgl 
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   2
         Top             =   400
         Width           =   1455
         _Version        =   851970
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   68
      End
      Begin XtremeSuiteControls.ComboBox CmbRpt 
         Height          =   360
         Index           =   1
         Left            =   3120
         TabIndex        =   5
         Top             =   1080
         Width           =   2895
         _Version        =   851970
         _ExtentX        =   5106
         _ExtentY        =   635
         _StockProps     =   77
         BackColor       =   -2147483643
         Sorted          =   -1  'True
         AutoComplete    =   -1  'True
         DropDownWidth   =   5000
      End
      Begin XtremeSuiteControls.ComboBox CmbRpt 
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   2895
         _Version        =   851970
         _ExtentX        =   5106
         _ExtentY        =   635
         _StockProps     =   77
         BackColor       =   -2147483643
         Sorted          =   -1  'True
         AutoComplete    =   -1  'True
         DropDownWidth   =   5000
      End
      Begin XtremeSuiteControls.PushButton cmdRptProcess 
         Height          =   735
         Index           =   0
         Left            =   4920
         TabIndex        =   9
         Top             =   240
         Width           =   975
         _Version        =   851970
         _ExtentX        =   1720
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Print View"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.ComboBox CmbRpt 
         Height          =   315
         Index           =   2
         Left            =   3120
         TabIndex        =   10
         Top             =   400
         Width           =   1695
         _Version        =   851970
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Sorted          =   -1  'True
         AutoComplete    =   -1  'True
         DropDownWidth   =   5000
      End
      Begin XtremeSuiteControls.PushButton cmdRptProcess 
         Height          =   735
         Index           =   1
         Left            =   5900
         TabIndex        =   17
         Top             =   240
         Width           =   975
         _Version        =   851970
         _ExtentX        =   1720
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Excel"
         Appearance      =   6
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Golongan."
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   11
         Top             =   160
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Dept."
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Name."
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   3120
         TabIndex        =   7
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Start date"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   160
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   3
         Top             =   160
         Width           =   975
      End
   End
End
Attribute VB_Name = "FRM_PAYROLL_REPORT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRptProcess_Click(Index As Integer)
If Index = 0 Then
    '-----------------------------
    '----- PRINT VIEW REPORT -----
    '-----------------------------
    Unload FRM_PRINVIEW
    FRM_PRINVIEW.RptNm = CmbRpt(4).Text
    FRM_PRINVIEW.Show
End If
If Index = 1 Then
    '-----------------------------
    '----- EXPORT REPORT EXCEL ---
    '-----------------------------
    If (CmbRpt(4).Text) = "PAYROLL_WEEKLY" Then
        If DateDiff("d", RptTgl(0).Value, RptTgl(1).Value) <= 8 Then
            PrjExportExcel.PAYROLL_WEEKLY Format(RptTgl(0).Value, "yyyy-mm-dd"), Format(RptTgl(1).Value, "yyyy-mm-dd"), CmbRpt(0).Text, CmbRpt(1).Text, CmbRpt(2).Text, CmbRpt(3).Text, "Payroll_WeeklyExcel", 2
        Else
            lR = SetTopMostWindow(Me.hwnd, False)
            MsgBox "Pastikan jumlah hari dalam seminggu ! Silakan di ulang kembali maksimal 8 Hari"
        End If
        lR = SetTopMostWindow(Me.hwnd, True)
    End If
    If (CmbRpt(4).Text) = "PAYROLL_MONTHLY" Then
         If DateDiff("d", RptTgl(0).Value, RptTgl(1).Value) <= 31 Then
            PrjExportExcel.PAYROLL_MONTHLY Format(RptTgl(0).Value, "yyyy-mm-dd"), Format(RptTgl(1).Value, "yyyy-mm-dd"), CmbRpt(0).Text, CmbRpt(1).Text, CmbRpt(2).Text, CmbRpt(3).Text, "Payroll_WeeklyExcel", 2
        Else
            lR = SetTopMostWindow(Me.hwnd, False)
            MsgBox "Pastikan jumlah hari dalam sebulan ! Silakan di ulang kembali maksimal 31 Hari"
        End If
        lR = SetTopMostWindow(Me.hwnd, True)
              
    End If
End If
End Sub

Private Sub Form_Load()
Main
lR = SetTopMostWindow(Me.hwnd, True)
init_DT_CMB_filter
End Sub
'=====================================
'============= ptr.nov ===============
'======= DATE INITIALIZE Filter ======
'=====================================
Private Sub init_DT_CMB_filter()
Dim i As Integer
For i = 0 To 1
    RptTgl(i).Value = Now
    RptTgl(i).Format = xtpPickerShortDate
Next i
prjSysID.Dept CmbRpt(0)
prjSysID.Employe CmbRpt(1), 0, 0, 0
prjSysID.PAYROLL_RPT_LIST CmbRpt(4)
prjSysID.TTGROUP CmbRpt(2)
prjSysID.CABANG CmbRpt(3)
End Sub
'============================================
'================= ptr.nov ==================
'===== PROCESS FILTER DEP-EMPLOYE CHANGE ====
'============================================
Private Sub CmbRpt_Click(Index As Integer)
If Index = 0 Then
    prjSysID.Employe CmbRpt(1), 0, 0, CmbRpt(0).Text
End If
End Sub
