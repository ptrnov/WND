VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Begin VB.Form FRM_PAYROLL 
   BackColor       =   &H00FFFFFF&
   Caption         =   "EMPLOYE ABSENSI PROCESS"
   ClientHeight    =   10635
   ClientLeft      =   90
   ClientTop       =   435
   ClientWidth     =   20160
   ControlBox      =   0   'False
   DrawMode        =   15  'Merge Pen Not
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10635
   ScaleWidth      =   20160
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.PushButton cmdProcess 
      Height          =   375
      Index           =   0
      Left            =   2520
      TabIndex        =   27
      Top             =   3240
      Width           =   1335
      _Version        =   851970
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Tampil >>"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   5775
      Left            =   0
      OleObjectBlob   =   "FRM_PAYROLL.frx":0000
      TabIndex        =   26
      Top             =   4080
      Width           =   3855
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   10575
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   3615
      _Version        =   851970
      _ExtentX        =   6376
      _ExtentY        =   18653
      _StockProps     =   79
      BackColor       =   12648384
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   1815
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   3615
         _Version        =   851970
         _ExtentX        =   6376
         _ExtentY        =   3201
         _StockProps     =   79
         BackColor       =   12648384
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.ComboBox CmbPayroll 
            Height          =   315
            Index           =   4
            Left            =   120
            TabIndex        =   22
            Top             =   1200
            Width           =   3255
            _Version        =   851970
            _ExtentX        =   5741
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Sorted          =   -1  'True
            AutoComplete    =   -1  'True
            DropDownWidth   =   5000
         End
         Begin XtremeSuiteControls.DateTimePicker DFilter 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   17
            ToolTipText     =   "Pilih Tanggal Mulai"
            Top             =   480
            Width           =   1695
            _Version        =   851970
            _ExtentX        =   2990
            _ExtentY        =   661
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   1
         End
         Begin XtremeSuiteControls.DateTimePicker DFilter 
            Height          =   375
            Index           =   1
            Left            =   1800
            TabIndex        =   18
            ToolTipText     =   "Pilih Tanggal Mulai"
            Top             =   480
            Width           =   1575
            _Version        =   851970
            _ExtentX        =   2778
            _ExtentY        =   661
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   1
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "* MENU"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   21
            Top             =   960
            Width           =   1095
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
            Index           =   1
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "End date"
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
            Left            =   1800
            TabIndex        =   19
            Top             =   240
            Width           =   1095
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox8 
         Height          =   2535
         Index           =   1
         Left            =   0
         TabIndex        =   1
         Top             =   1800
         Width           =   3615
         _Version        =   851970
         _ExtentX        =   6376
         _ExtentY        =   4471
         _StockProps     =   79
         BackColor       =   8438015
         UseVisualStyle  =   -1  'True
         BorderStyle     =   2
         Begin XtremeSuiteControls.ProgressBar ProgressBar2 
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   2040
            Width           =   3495
            _Version        =   851970
            _ExtentX        =   6165
            _ExtentY        =   450
            _StockProps     =   93
            Appearance      =   1
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox CmbPayroll 
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   2
            Top             =   1080
            Width           =   1815
            _Version        =   851970
            _ExtentX        =   3201
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Sorted          =   -1  'True
            AutoComplete    =   -1  'True
            DropDownWidth   =   5000
         End
         Begin XtremeSuiteControls.ComboBox CmbPayroll 
            Height          =   315
            Index           =   1
            Left            =   1920
            TabIndex        =   11
            Top             =   480
            Width           =   1575
            _Version        =   851970
            _ExtentX        =   2778
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Sorted          =   -1  'True
            AutoComplete    =   -1  'True
            DropDownWidth   =   5000
         End
         Begin XtremeSuiteControls.ComboBox CmbPayroll 
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   480
            Width           =   1815
            _Version        =   851970
            _ExtentX        =   3201
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Sorted          =   -1  'True
            AutoComplete    =   -1  'True
            DropDownWidth   =   5000
         End
         Begin XtremeSuiteControls.ComboBox CmbPayroll 
            Height          =   315
            Index           =   3
            Left            =   1920
            TabIndex        =   14
            Top             =   1080
            Width           =   1575
            _Version        =   851970
            _ExtentX        =   2778
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Sorted          =   -1  'True
            AutoComplete    =   -1  'True
            DropDownWidth   =   5000
         End
         Begin XtremeSuiteControls.PushButton cmdProcess 
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   32
            Top             =   1560
            Width           =   1215
            _Version        =   851970
            _ExtentX        =   2143
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "<< Calculate >>"
            ForeColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdProcess 
            Height          =   375
            Index           =   2
            Left            =   1320
            TabIndex        =   34
            Top             =   1560
            Width           =   975
            _Version        =   851970
            _ExtentX        =   1720
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Report"
            ForeColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Label Label2 
            BackColor       =   &H0080C0FF&
            Caption         =   "Group"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   1920
            TabIndex        =   15
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H0080C0FF&
            Caption         =   "Dept."
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   13
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label2 
            BackColor       =   &H0080C0FF&
            Caption         =   "Cabang."
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   1920
            TabIndex        =   4
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H0080C0FF&
            Caption         =   "Corp."
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   735
         End
      End
      Begin XtremeSuiteControls.PushButton cmdProcess 
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   5
         Top             =   10080
         Width           =   975
         _Version        =   851970
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Exit"
         Appearance      =   6
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   10575
      Left            =   3600
      TabIndex        =   6
      Top             =   0
      Width           =   17175
      _Version        =   851970
      _ExtentX        =   30295
      _ExtentY        =   18653
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   3
      Color           =   32
      PaintManager.BoldSelected=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.FixedTabWidth=   180
      PaintManager.MaxTabWidth=   180
      PaintManager.MinTabWidth=   180
      ItemCount       =   6
      Item(0).Caption =   "Employe Data"
      Item(0).ControlCount=   4
      Item(0).Control(0)=   "ProgressBar1"
      Item(0).Control(1)=   "dxDBGrid_data(0)"
      Item(0).Control(2)=   "dxDBGrid_data(1)"
      Item(0).Control(3)=   "CmbPayroll(5)"
      Item(1).Caption =   "Employe Presensi"
      Item(1).ControlCount=   3
      Item(1).Control(0)=   "dxDBGrid4"
      Item(1).Control(1)=   "cmdRptProcess(2)"
      Item(1).Control(2)=   "dxDBGrid2"
      Item(2).Caption =   "Employe OverTime"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "dxDBGrid3"
      Item(3).Caption =   "Week Payment Payroll "
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "dxDBGrid_Payroll(0)"
      Item(4).Caption =   "Potongan Bulanan"
      Item(4).ControlCount=   1
      Item(4).Control(0)=   "dxDBGrid_Payroll(1)"
      Item(5).Caption =   "Month Payment Payroll "
      Item(5).ControlCount=   1
      Item(5).Control(0)=   "dxDBGrid_Payroll(2)"
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid4 
         Height          =   735
         Left            =   -1.39880e5
         OleObjectBlob   =   "FRM_PAYROLL.frx":0CA8
         TabIndex        =   7
         Top             =   840
         Visible         =   0   'False
         Width           =   1815
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid_data 
         Height          =   3135
         Index           =   0
         Left            =   360
         OleObjectBlob   =   "FRM_PAYROLL.frx":1950
         TabIndex        =   8
         Top             =   480
         Width           =   16215
      End
      Begin XtremeSuiteControls.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   12120
         TabIndex        =   9
         Top             =   -600
         Width           =   3135
         _Version        =   851970
         _ExtentX        =   5530
         _ExtentY        =   450
         _StockProps     =   93
      End
      Begin XtremeSuiteControls.PushButton cmdRptProcess 
         Height          =   375
         Index           =   2
         Left            =   -1.39880e5
         TabIndex        =   10
         ToolTipText     =   "Tekan Untuk Menampilkan Flter"
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
         _Version        =   851970
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Export Report"
         Appearance      =   6
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid2 
         Height          =   9135
         Left            =   -69640
         OleObjectBlob   =   "FRM_PAYROLL.frx":25F8
         TabIndex        =   23
         Top             =   1080
         Visible         =   0   'False
         Width           =   16095
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid3 
         Height          =   9135
         Left            =   -69640
         OleObjectBlob   =   "FRM_PAYROLL.frx":32A0
         TabIndex        =   24
         Top             =   1080
         Visible         =   0   'False
         Width           =   16095
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid_Payroll 
         Height          =   9135
         Index           =   0
         Left            =   -69640
         OleObjectBlob   =   "FRM_PAYROLL.frx":3F48
         TabIndex        =   25
         Top             =   1080
         Visible         =   0   'False
         Width           =   16095
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid_data 
         Height          =   6015
         Index           =   1
         Left            =   360
         OleObjectBlob   =   "FRM_PAYROLL.frx":4BF0
         TabIndex        =   28
         Top             =   4200
         Width           =   16215
      End
      Begin XtremeSuiteControls.ComboBox CmbPayroll 
         Height          =   315
         Index           =   5
         Left            =   360
         TabIndex        =   29
         Top             =   3840
         Width           =   2535
         _Version        =   851970
         _ExtentX        =   4471
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Sorted          =   -1  'True
         AutoComplete    =   -1  'True
         DropDownWidth   =   5000
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid_Payroll 
         Height          =   9135
         Index           =   1
         Left            =   -69640
         OleObjectBlob   =   "FRM_PAYROLL.frx":5898
         TabIndex        =   30
         Top             =   1080
         Visible         =   0   'False
         Width           =   16095
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid_Payroll 
         Height          =   9135
         Index           =   2
         Left            =   -69640
         OleObjectBlob   =   "FRM_PAYROLL.frx":6540
         TabIndex        =   31
         Top             =   1080
         Visible         =   0   'False
         Width           =   16095
      End
   End
End
Attribute VB_Name = "FRM_PAYROLL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'===================================
'============= ptr.nov =============
'============= VARIABLE ============
'===================================
Dim Item As TaskPanelGroupItem
Dim CLM1, CLM2, clm3, clm4, clm5, clm6, BND1, BND2, BND3, BND4, BND5, BND6 As Variant
Dim Qry As String, Qry0 As String, Qry1 As String, Qry2 As String, Qry3 As String, Qry4 As String, Qry5 As String, Qry6 As String
Dim LoadWas As Boolean
Dim a, i As Integer
Dim b As String
Dim x As VbMsgBoxResult
Dim trecs As Double
Dim STgl As Date, ETgl As Date, SFiter As Date
Dim PERSENSI_STGL As Integer, FilterDateTgl As Integer, F_TGL As String, F_WAKTU As String, F_TGL_WAKTU As String


Private Sub cmdProcess_Click(Index As Integer)
'========================
'---- PROCESS BUTTON ----
'========================
If Index = 0 Then
   
    If (CmbPayroll(4).Text = 1) Then
        '--- TAB PRESENSI ----
        TabControl1.Item(1).Selected = True
        GridFill 2, 0
    End If
    
    If (CmbPayroll(4).Text = 2) Then
        '---- TAB OVERTIME ---
       TabControl1.Item(2).Selected = True
       GridFill 3, 0
    End If
    
    If (CmbPayroll(4).Text = 3) Then
        'TAB PEYROLL PAYMENT MINGGUAN----
        GridFill 4, 0
       TabControl1.Item(3).Selected = True
    End If
    
    If (CmbPayroll(4).Text = 4) Then
        '---- TAB POTONGAN BULANAN ----
       TabControl1.Item(4).Selected = True
       GridFill 5, 0
    End If
       
    If (CmbPayroll(4).Text = 5) Then
        'TAB PEYROLL BULANAN ----
        GridFill 6, 0
       TabControl1.Item(5).Selected = True
    End If
    
    If CmbPayroll(4).Text = 6 Then
        '--- TAB EMPLYOYE DATA ---
        TabControl1.Item(0).Selected = True
    End If
   
    
End If
If Index = 1 Then
       Qry = ""
        b = ""
    a = DateDiff("d", Format(DFilter(0).Value, "yyyy-mm-dd"), Format(DFilter(1).Value, "yyyy-mm-dd"))
    If a <= 8 Then
        Qry = "payroll_process('0=minggu" & _
                    "=" & Format(DFilter(0).Value, "yyyy-mm-dd") & _
                    "=" & Format(DFilter(1).Value, "yyyy-mm-dd") & _
                    "=" & dxDBGrid1.Dataset.FieldValues("KAR_ID") & _
                    "=" & CmbPayroll(1).Text & _
                    "=" & CmbPayroll(3).Text & _
                    "=" & CmbPayroll(2).Text & _
        "')"
        b = "CLOSING MINGGUAN"
    ElseIf a > 8 Then
        Qry = "payroll_process('0=bulan" & _
                    "=" & Format(DFilter(0).Value, "yyyy-mm-dd") & _
                    "=" & Format(DFilter(1).Value, "yyyy-mm-dd") & _
                    "=" & dxDBGrid1.Dataset.FieldValues("KAR_ID") & _
                    "=" & CmbPayroll(1).Text & _
                    "=" & CmbPayroll(3).Text & _
                    "=" & CmbPayroll(2).Text & _
        "')"
        b = "CLOSING BULANAN"
    End If
    x = MsgBox("Anda Akan Memprosess " & b & " ? ", vbYesNo, "PAYROLL PROCESS CALCULATE")
    If x = vbYes Then
        ProgressBar2.Max = 0
        i = 0
        OpRec5 Qry, True
        With rs5
            If Not .EOF Then
                .MoveFirst:
                Do Until .EOF
                    trecs = trecs + 1
                .MoveNext
                Loop
                
                With ProgressBar2
                    .Visible = True
                    .Min = 0
                    .Max = trecs
                End With

                .MoveFirst
                Do Until .EOF
                    ProgressBar2.Value = i
                .MoveNext
                i = i + 1
                Loop
            End If
        End With
        ProgressBar2.Value = trecs
        If ProgressBar2.Value = ProgressBar2.Max Then
            MsgBox "COMPLETE", , "PAYROLL PROCESS CALCULATE"
        End If
    End If
    'MsgBox Qry
    '0=minggu=2015-03-01=2015-03-6=WAN.HO.000087,0,0,0'
    '0=bulan=2015-03-01=2015-03-16=WAN.HO.000087,0,0,0'
    'Menu,periode,sDateRun,eDateRun,KarId,CabId,GrpId,DeptId
End If
If Index = 2 Then
    FRM_PAYROLL_REPORT.Show
End If
If Index = 7 Then
    Unload Me
End If
End Sub

Private Sub CmbPayroll_Click(Index As Integer)
If Index = 1 Or Index = 2 Or Index = 3 Then
    GridFill 1, 0
End If
If Index = 4 Then
     
    If (CmbPayroll(4).Text = 1) Then
        '--- TAB PRESENSI ----
        TabControl1.Item(1).Selected = True
    End If
    
    If (CmbPayroll(4).Text = 2) Then
        '---- TAB OVERTIME ---
       TabControl1.Item(2).Selected = True
    End If
    
    If (CmbPayroll(4).Text = 3) Then
        'TAB PEYROLL PAYMENT MINGGUAN----
       TabControl1.Item(3).Selected = True
    End If
    
    If (CmbPayroll(4).Text = 4) Then
        '---- TAB POTONGAN BULANAN ----
       TabControl1.Item(4).Selected = True
    End If
    
    If CmbPayroll(4).Text = 5 Then
        '--- TAB PEYROLL PAYMENT BULANAN ---
        TabControl1.Item(5).Selected = True
    End If
    
    If CmbPayroll(4).Text = 6 Then
        '--- TAB EMPLYOYE DATA ---
        TabControl1.Item(0).Selected = True
    End If

End If
End Sub

Private Sub dxDBGrid1_OnDblClick()
'=============================
'--- SHOW GRID PER EMPLOYE ---
'=============================
With dxDBGrid1.Dataset
    If dxDBGrid1.Dataset.RecNo <> 0 Then
    
        If (CmbPayroll(4).Text = 1) Then
            '--- TAB PRESENSI ----
            TabControl1.Item(1).Selected = True
            GridFill 2, .FieldValues("KAR_ID")
        End If
        
        If (CmbPayroll(4).Text = 2) Then
            '---- TAB OVERTIME ---
           TabControl1.Item(2).Selected = True
           GridFill 3, .FieldValues("KAR_ID")
        End If
        
        If (CmbPayroll(4).Text = 3) Then
            '---- TAB PEYROLL PAYMENT MINGGUAN ----
            GridFill 4, .FieldValues("KAR_ID")
           TabControl1.Item(3).Selected = True
        End If
        
        If (CmbPayroll(4).Text = 4) Then
            '-----TAB POTONGAN BULANAN ----
            GridFill 5, .FieldValues("KAR_ID")
           TabControl1.Item(4).Selected = True
        End If
        
        If CmbPayroll(4).Text = 5 Then
            '--- TAB EMPLYOYE DATA ---
            GridFill 6, .FieldValues("KAR_ID")
            TabControl1.Item(5).Selected = True
        End If
        
        If CmbPayroll(4).Text = 6 Then
            '--- TAB EMPLYOYE DATA ---
            TabControl1.Item(0).Selected = True
        End If

    End If
End With
End Sub

Private Sub Form_Load()
'==== LOAD ==
Main
init
GridFill 1, 0
End Sub

Private Sub init()
DFilter(0).Value = Now
DFilter(1).Value = Now
prjSysID.CORP CmbPayroll(0)
prjSysID.CABANG CmbPayroll(1)
prjSysID.Dept CmbPayroll(2)
prjSysID.TTGROUP CmbPayroll(3)
prjSysID.MENU_PAYROLL CmbPayroll(4)
End Sub


Private Sub MenuGrid(TabSel As Integer)
'=====================================
'------------ ptr.nov ----------------
'----- MENU CULUMN BND GRID LODING ---
'=====================================
Select Case TabSel
Case Is = 1
     '===========================
     '--- MENU PRESENSI_INOUT ---
     '===========================
    BND1 = Array("EMPLOYE")
    CLM1 = Array(Array("Clm0", "EmployeID", "KAR_ID", gedTextEdit, 0, 0, 100, 0, 0, 0, 0), _
                Array("Clm1", "Employe Name", "KAR_NM", gedTextEdit, 0, 0, 130, 0, 0, 0, 0))
                '--------0---------1--------2-----------3--------4--5--6---7--8--9--10
 Case Is = 2
    '========================================
    '--- MENU PAYROLL_SESSION -> Presensi ---
    '========================================
    BND2 = Array(" EMPLOYE PROPERTIES ", " PAYROLL PRESENSI VALUES")
    CLM2 = Array(Array("Clm0", "EmployeID", "KAR_ID", gedTextEdit, 0, 0, 120, 1, 1, 0, 0), _
                  Array("Clm1", "Employe Name", "KAR_NM", gedTextEdit, 0, 0, 105, 1, 1, 0, 0), _
                  Array("Clm2", "Group", "GRP_ID", gedLookupEdit, 0, 0, 55, 0, 1, 0, 0), _
                  Array("Clm3", "Departmen ", "DEP_ID", gedLookupEdit, 0, 0, 120, 1, 1, 0, 0), _
                  Array("Clm4", "Cabang ", "CAB_ID", gedLookupEdit, 0, 0, 80, 1, 1, 0, 0), _
                  Array("Clm5", "LOCK", "KONCI", gedCheckEdit, 0, 1, 40, 1, 2, 0, 0), _
                  Array("Clm6", "Days", "Hari", gedTextEdit, 0, 1, 50, 1, 1, 0, 3), _
                  Array("Clm7", "Date", "TGL_RUN", gedTextEdit, 0, 1, 70, 1, 2, 0, 3), _
                  Array("Clm8", "Date IN", "DT_IN", gedTextEdit, 0, 1, 120, 1, 2, 0, 2), _
                  Array("Clm9", "Date Out", "DT_OUT", gedTextEdit, 0, 1, 120, 1, 2, 0, 2), _
                  Array("Clm10", "Pagi", "ACTIVE_LATE", gedCurrencyEdit, 0, 1, 60, 4, 1, 0, 0), _
                  Array("Clm11", "Work Time", "WORK_TIME", gedTextEdit, 0, 1, 60, 4, 2, 0, 1), _
                  Array("Clm12", "Hari Aktif", "DAY_ACTIVE", gedTextEdit, 0, 1, 60, 4, 2, 0, 0), _
                  Array("Clm13", "Hari Libur", "HOLIDAY", gedTextEdit, 0, 1, 60, 4, 2, 0, 0), _
                  Array("Clm14", "Absen", "DAY_ABSENT", gedTextEdit, 0, 1, 60, 4, 2, 0, 0), _
                  Array("Clm15", "Ijin", "DAY_EXCP_ID", gedTextEdit, 0, 1, 50, 4, 2, 0, 0), _
                  Array("Clm16", "Late", "LATE", gedTextEdit, 0, 1, 60, 4, 2, 4, 1), _
                  Array("Clm17", "Early", "EARLY", gedTextEdit, 0, 1, 60, 4, 2, 0, 1))
                '--------0---------1--------2-----------3--------4--5--6---7--8--9--10
Case Is = 3
    '========================================
    '--- MENU PAYROLL_SESSION -> OVERTIME ---
    '========================================
    BND3 = Array(" EMPLOYE PROPERTIES ", " PAYROLL OVERTIME VALUES")
    clm3 = Array(Array("Clm0", "EmployeID", "KAR_ID", gedTextEdit, 0, 0, 120, 1, 1, 0, 0), _
                  Array("Clm1", "Employe Name", "KAR_NM", gedTextEdit, 0, 0, 105, 1, 0, 0, 0), _
                  Array("Clm2", "Group", "GRP_ID", gedLookupEdit, 0, 0, 55, 1, 1, 0, 0), _
                  Array("Clm3", "Departmen ", "DEP_ID", gedLookupEdit, 0, 0, 120, 1, 1, 0, 0), _
                  Array("Clm4", "Cabang ", "CAB_ID", gedLookupEdit, 0, 0, 80, 1, 1, 0, 0), _
                  Array("Clm5", "LOCK", "KONCI", gedCheckEdit, 0, 1, 40, 1, 2, 0, 0), _
                  Array("Clm6", "Days", "Hari", gedTextEdit, 0, 1, 50, 1, 1, 2, 3), _
                  Array("Clm7", "Date", "TGL_RUN", gedTextEdit, 0, 1, 70, 1, 2, 0, 3), _
                  Array("Clm8", "Date IN", "DT_IN", gedTextEdit, 0, 1, 120, 1, 2, 0, 2), _
                  Array("Clm9", "Date Out", "DT_OUT", gedTextEdit, 0, 1, 120, 1, 2, 0, 2), _
                  Array("Clm10", "OT Depan", "OT_DPN", gedTextEdit, 0, 1, 60, 10, 2, 0, 1), _
                  Array("Clm11", "OT Hari", "OT_DPN_HARI", gedTextEdit, 0, 1, 50, 10, 1, 0, 6), _
                  Array("Clm12", "Pagi", "ACTIVE_LATE", gedCurrencyEdit, 0, 1, 60, 4, 1, 0, 0), _
                  Array("Clm13", "OT Lev1", "OT1_BLK", gedTextEdit, 0, 1, 60, 4, 2, 0, 1), _
                  Array("Clm14", "OT Hari", "OT1_BLK_HARI", gedTextEdit, 0, 1, 50, 4, 1, 0, 6), _
                  Array("Clm15", "OT Lev2", "OT2_BLK", gedTextEdit, 0, 1, 60, 10, 2, 0, 1), _
                  Array("Clm16", "OT Hari", "OT2_BLK_HARI", gedTextEdit, 0, 1, 50, 10, 1, 0, 6), _
                  Array("Clm17", "OT Lev3", "OT3_BLK", gedTextEdit, 0, 1, 60, 4, 2, 0, 1), _
                  Array("Clm18", "OT Hari", "OT3_BLK_HARI", gedTextEdit, 0, 1, 50, 4, 1, 0, 6))
                '--------0---------1--------2-----------3--------4--5--6---7--8--9--10


Case Is = 4
    '========================================
    '--- MENU PAYROLL_SESSION -> HARIAN ---
    '========================================
    BND4 = Array(" SESI MINGGUAN ", " NILAI PRESENSI PAYROLL")
    clm4 = Array(Array("Clm0", "EmployeID", "KAR_ID", gedTextEdit, 0, 0, 100, 1, 0, 1, 0), _
                  Array("Clm1", "Employe Name", "KAR_NM", gedTextEdit, 0, 0, 90, 1, 1, 0, 0), _
                  Array("Clm2", "Departmen ", "DEP_ID", gedLookupEdit, 0, 0, 80, 1, 1, 0, 0), _
                  Array("Clm3", "sDate", "sTGL", gedTextEdit, 0, 0, 70, 1, 2, 0, 3), _
                  Array("Clm4", "eDate", "eTGL", gedTextEdit, 0, 0, 70, 1, 2, 0, 3), _
                  Array("Clm5", "LOCK", "KONCI", gedCheckEdit, 0, 0, 40, 1, 2, 0, 0), _
                  Array("Clm6", "OT DPN", "TTL_OT_DPN_HARI", gedCurrencyEdit, 0, 1, 70, 10, 1, 0, 0), _
                  Array("Clm7", "OT PAGI", "TTL_PAGI", gedCurrencyEdit, 0, 1, 80, 10, 1, 0, 0), _
                  Array("Clm8", "OT BLK", "TTL_OT_BLK_HARI", gedCurrencyEdit, 0, 1, 80, 10, 1, 0, 0), _
                  Array("Clm9", "UPAH HARIAN", "PAY_DAY", gedCurrencyEdit, 0, 1, 90, 1, 3, 0, 6), _
                  Array("Clm10", "UPAH OT DPN", "TTL_PAY_OT_DPN", gedCurrencyEdit, 0, 1, 90, 10, 3, 0, 6), _
                  Array("Clm11", "UPAH PAGI", "TTL_PAY_HARIAN", gedCurrencyEdit, 0, 1, 90, 10, 3, 0, 6), _
                  Array("Clm12", "UPAH OT BLK", "TTL_PAY_OT_BLK", gedCurrencyEdit, 0, 1, 100, 10, 3, 0, 6), _
                  Array("Clm13", "UPAH DITERIMA", "UPAH_DITERIMA", gedCurrencyEdit, 0, 1, 110, 1, 3, 0, 6))
                '--------0---------1--------2-----------3--------4--5--6---7--8--9--10
Case Is = 5
    '================================================
    '--- MENU PAYROLL_SESSION -> POTONGAN BULANAN ---
    '================================================
    BND5 = Array(" SESI POTONGAN BULANAN", " PAYROLL OVERTIME VALUES")
    clm5 = Array(Array("Clm0", "EmployeID", "KAR_ID", gedTextEdit, 0, 0, 100, 1, 0, 1, 0), _
                  Array("Clm1", "Employe Name", "KAR_NM", gedTextEdit, 0, 0, 90, 1, 1, 0, 0), _
                  Array("Clm2", "Departmen ", "DEP_ID", gedLookupEdit, 0, 0, 80, 1, 1, 0, 0), _
                  Array("Clm3", "sDate", "sTGL", gedTextEdit, 0, 0, 70, 1, 2, 0, 3), _
                  Array("Clm4", "eDate", "eTGL", gedTextEdit, 0, 0, 70, 1, 2, 0, 3), _
                  Array("Clm5", "LOCK", "KONCI", gedCheckEdit, 0, 0, 40, 1, 2, 0, 0), _
                  Array("Clm6", "PPH21", "POT_PPH21", gedCurrencyEdit, 0, 1, 90, 4, 3, 0, 6), _
                  Array("Clm7", "JAMSOS JHT TK", "POT_JAMSOS_JHT_TK", gedCurrencyEdit, 0, 1, 110, 4, 3, 0, 6), _
                  Array("Clm8", "ASURANSI", "POT_ASURANSI", gedCurrencyEdit, 0, 1, 100, 4, 3, 0, 6), _
                  Array("Clm9", "PINJAMAN", "POT_PINJAMAN", gedCurrencyEdit, 0, 1, 100, 4, 3, 0, 6), _
                  Array("Clm10", "TOTAL POTONGAN", "TOTAL_POT_EMP", gedCurrencyEdit, 0, 1, 120, 10, 3, 0, 6), _
                  Array("Clm11", "CORP JAMSOS JHT", "POT_JAMSOS_JHT", gedCurrencyEdit, 0, 1, 120, 1, 3, 0, 6), _
                  Array("Clm12", "CORP PAY JAMSOS", "TOTAL_PAY_CORP", gedCurrencyEdit, 0, 1, 130, 1, 3, 0, 6))
         '--------0---------1--------2-----------3--------4--5--6---7--8--9--10

Case Is = 6
    '========================================
    '--- MENU PAYROLL_SESSION -> BULANAN ---
    '========================================
    BND6 = Array(" SESI BULANAN ", " NILAI PRESENSI PAYROLL")
    clm6 = Array(Array("Clm0", "EmployeID", "KAR_ID", gedTextEdit, 0, 0, 100, 1, 1, 0, 0), Array("Clm1", "Employe Name", "KAR_NM", gedTextEdit, 0, 0, 90, 1, 1, 0, 0), Array("Clm2", "Departmen ", "DEP_ID", gedLookupEdit, 0, 0, 80, 1, 1, 0, 0), _
                  Array("Clm3", "sDate", "sTGL", gedTextEdit, 0, 0, 70, 1, 2, 0, 3), Array("Clm4", "eDate", "eTGL", gedTextEdit, 0, 0, 70, 1, 2, 0, 3), Array("Clm5", "LOCK", "KONCI", gedCheckEdit, 0, 0, 40, 1, 2, 0, 0), _
                  Array("Clm6", "OT DPN", "TTL_OT_DPN_HARI", gedCurrencyEdit, 0, 1, 70, 10, 1, 0, 0), _
                  Array("Clm7", "OT PAGI", "TTL_PAGI", gedCurrencyEdit, 0, 1, 80, 10, 1, 0, 0), _
                  Array("Clm8", "OT BLK", "TTL_OT_BLK_HARI", gedCurrencyEdit, 0, 1, 80, 10, 1, 0, 0), _
                  Array("Clm9", "UPAH HARIAN", "PAY_DAY", gedCurrencyEdit, 0, 1, 90, 1, 3, 0, 6), _
                  Array("Clm10", "UPAH OT DPN", "TTL_PAY_OT_DPN", gedCurrencyEdit, 0, 1, 90, 10, 3, 0, 6), _
                  Array("Clm11", "UPAH PAGI", "TTL_PAY_HARIAN", gedCurrencyEdit, 0, 1, 90, 10, 3, 0, 6), _
                  Array("Clm12", "UPAH OT BLK", "TTL_PAY_OT_BLK", gedCurrencyEdit, 0, 1, 90, 10, 3, 0, 6), _
                  Array("Clm13", "TOTAL UPAH", "TTL_UPAH_PERHARI", gedCurrencyEdit, 0, 1, 90, 1, 3, 0, 6), _
                  Array("Clm14", "POT_PPH21", "POT_PPH21", gedCurrencyEdit, 0, 1, 90, 10, 3, 0, 6), _
                  Array("Clm15", "POT_JAMSOS", "POT_JAMSOS_JHT_TK", gedCurrencyEdit, 0, 1, 90, 10, 3, 0, 6), _
                  Array("Clm16", "POT_ASURANSI", "POT_ASURANSI", gedCurrencyEdit, 0, 1, 90, 10, 3, 0, 6), _
                  Array("Clm17", "POT_PINJAMAN", "POT_PINJAMAN", gedCurrencyEdit, 0, 1, 90, 10, 3, 0, 6), _
                  Array("Clm19", "TOTAL POT", "TOTAL_POT", gedCurrencyEdit, 0, 1, 90, 1, 3, 0, 6), _
                  Array("Clm19", "GAJI POKOK", "PAY_SALARY", gedCurrencyEdit, 0, 1, 90, 10, 3, 0, 6), _
                  Array("Clm20", "TUNJANGAN", "PAY_TUNJANGAN", gedCurrencyEdit, 0, 1, 90, 10, 3, 0, 6), _
                  Array("Clm21", "TRANSPORT", "PAY_TRANSPORT", gedCurrencyEdit, 0, 1, 90, 10, 3, 0, 6), _
                  Array("Clm22", "MAKAN", "PAY_EAT", gedCurrencyEdit, 0, 1, 90, 10, 23, 0, 6), _
                  Array("Clm23", "BONUS", "PAY_BONUS", gedCurrencyEdit, 0, 1, 90, 10, 3, 0, 6), _
                  Array("Clm24", "PAY_ENTERTAIN", "PAY_ENTERTAIN", gedCurrencyEdit, 0, 1, 90, 10, 3, 0, 6), _
                  Array("Clm25", "TOTAl GAJI", "TOTAL_GAJI", gedCurrencyEdit, 0, 1, 90, 1, 3, 0, 6), _
                  Array("Clm26", "GAJI DITERIMA", "GAJI_DITERIMA", gedCurrencyEdit, 0, 1, 120, 8, 3, 0, 6))
                '--------0---------1--------2-----------3--------4--5--6---7--8--9--10
End Select
End Sub

Public Sub GridFill(GrdIndx As Integer, KarId As String)
'=====================================
'------------- ptr.nov ---------------
'----- PROCESS FIRST GRID LODING -----
'=====================================
Select Case GrdIndx
    Case Is = 1
            '===================
            '-- LIST KARYAWAN --
            '===================
            MenuGrid 1
            Qry1 = "payroll_grid('karyawan_list" & _
                                  "=" & Format(DFilter(0).Value, "yyyy-mm-dd") & _
                                  "=" & Format(DFilter(1).Value, "yyyy-mm-dd") & _
                                  "=" & CmbPayroll(2).Text & _
                                  "=0" & _
                                  "=" & CmbPayroll(3).Text & _
                                  "=" & CmbPayroll(1).Text & _
                                  "=0=0" & _
                    "')"

            'Qry1 = "SELECT KAR_ID,KAR_NM FROM karyawan"
            dxDBGrid1.Dataset.Refresh
            PrjSysGrid.GetGrid_Persensi dxDBGrid1, CLM1, False, False, BND1, True, Qry1, "KAR_ID", False
            dxDBGrid1.Dataset.Open
            
    Case Is = 2
            '=================================
            '-- PAYROLL_SESSION -> Presensi --
            '=================================
            TabControl1.Item(1).Selected = True
            MenuGrid 2
            Qry2 = "payroll_grid('payroll_presensi_session" & _
                                  "=" & Format(DFilter(0).Value, "yyyy-mm-dd") & _
                                  "=" & Format(DFilter(1).Value, "yyyy-mm-dd") & _
                                  "=" & CmbPayroll(2).Text & _
                                  "=" & KarId & _
                                  "=" & CmbPayroll(3).Text & _
                                  "=" & CmbPayroll(1).Text & _
                                  "=0=0" & _
                    "')"

            dxDBGrid2.Dataset.Refresh
            PrjSysGrid.GetGrid_Persensi dxDBGrid2, CLM2, True, True, BND2, True, Qry2, "KAR_ID", False
            'dxDBGrid3.Dataset.Open
            Lookup_GridFill2
            With dxDBGrid2
                .Columns.ColumnByName("Clm0").GroupIndex = 0
                .GroupNodeColor = &HC0FFFF
                .M.FullExpand
            End With
            
    Case Is = 3
            '=================================
            '-- PAYROLL_SESSION -> OVERTIME --
            '=================================
            TabControl1.Item(2).Selected = True
            MenuGrid 3
            Qry3 = "payroll_grid('payroll_overtime_session" & _
                                  "=" & Format(DFilter(0).Value, "yyyy-mm-dd") & _
                                  "=" & Format(DFilter(1).Value, "yyyy-mm-dd") & _
                                  "=" & CmbPayroll(2).Text & _
                                  "=" & KarId & _
                                  "=" & CmbPayroll(3).Text & _
                                  "=" & CmbPayroll(1).Text & _
                                  "=0=0" & _
                    "')"

            dxDBGrid3.Dataset.Refresh
            PrjSysGrid.GetGrid_Persensi dxDBGrid3, clm3, True, True, BND3, True, Qry3, "KAR_ID", False
            'dxDBGrid5.Dataset.Open
            Lookup_GridFill3
            With dxDBGrid3
                .Columns.ColumnByName("Clm0").GroupIndex = 0
                .GroupNodeColor = &HC0FFFF
                .M.FullExpand
            End With
            
    Case Is = 4
            '=================================
            '-- PAYROLL_SESSION -> HARIAN --
            '=================================
            TabControl1.Item(3).Selected = True
            MenuGrid 4
            Qry4 = "payroll_grid('payroll_session_week" & _
                                  "=" & Format(DFilter(0).Value, "yyyy-mm-dd") & _
                                  "=" & Format(DFilter(1).Value, "yyyy-mm-dd") & _
                                  "=" & CmbPayroll(2).Text & _
                                  "=" & KarId & _
                                  "=" & CmbPayroll(3).Text & _
                                  "=" & CmbPayroll(1).Text & _
                                  "=0=minggu" & _
                    "')"

            dxDBGrid_Payroll(0).Dataset.Refresh
            PrjSysGrid.GetGrid_Persensi dxDBGrid_Payroll(0), clm4, True, True, BND4, True, Qry4, "KAR_ID", False
            'dxDBGrid_Payroll(1).Dataset.Open
            Lookup_GridFill4
            With dxDBGrid_Payroll(0)
                .Columns.ColumnByName("Clm0").GroupIndex = 0
                .GroupNodeColor = &HC0FFFF
                .M.FullExpand
            End With
    
    Case Is = 5
            '=========================================
            '-- PAYROLL_SESSION -> POTONGAN BULANAN --
            '=========================================
            TabControl1.Item(4).Selected = True
            MenuGrid 5
            Qry5 = "payroll_grid('payroll_session_potongan_month" & _
                                  "=" & Format(DFilter(0).Value, "yyyy-mm-dd") & _
                                  "=" & Format(DFilter(1).Value, "yyyy-mm-dd") & _
                                  "=" & CmbPayroll(2).Text & _
                                  "=" & KarId & _
                                  "=" & CmbPayroll(3).Text & _
                                  "=" & CmbPayroll(1).Text & _
                                  "=0=bulan" & _
                    "')"

            dxDBGrid_Payroll(1).Dataset.Refresh
            PrjSysGrid.GetGrid_Persensi dxDBGrid_Payroll(1), clm5, True, True, BND5, True, Qry5, "KAR_ID", False
            'dxDBGrid_Payroll(1).Dataset.Open
            Lookup_GridFill5
            With dxDBGrid_Payroll(1)
                .Columns.ColumnByName("Clm0").GroupIndex = 0
                .GroupNodeColor = &HC0FFFF
                .M.FullExpand
            End With

    Case Is = 6
            '=================================
            '-- PAYROLL_SESSION -> BULANAN --
            '=================================
            TabControl1.Item(5).Selected = True
            MenuGrid 6
            Qry6 = "payroll_grid('payroll_session_month" & _
                                  "=" & Format(DFilter(0).Value, "yyyy-mm-dd") & _
                                  "=" & Format(DFilter(1).Value, "yyyy-mm-dd") & _
                                  "=" & CmbPayroll(2).Text & _
                                  "=" & KarId & _
                                  "=" & CmbPayroll(3).Text & _
                                  "=" & CmbPayroll(1).Text & _
                                  "=0=bulan" & _
                    "')"

            dxDBGrid_Payroll(2).Dataset.Refresh
            PrjSysGrid.GetGrid_Persensi dxDBGrid_Payroll(2), clm6, True, True, BND6, True, Qry6, "KAR_ID", False
            'dxDBGrid_Payroll(1).Dataset.Open
            Lookup_GridFill6
            With dxDBGrid_Payroll(2)
                .Columns.ColumnByName("Clm0").GroupIndex = 0
                .GroupNodeColor = &HC0FFFF
                .M.FullExpand
            End With

 
    End Select
End Sub

Private Sub Lookup_GridFill5()
'==========================================
'-- PAYROLL_SESSION -> POTONGAN BULANAN  --
'==========================================
With dxDBGrid_Payroll(1).Columns.ColumnByName("Clm2").LookupColumn
'== DEPT NAME ---
            .LookupDataset.EnableControls
            '.LookupDataset.Close
            '.Dataset.Refresh
            .LookupDatasetType = dtADODataset
            .LookupDataset.ADODataset.ConnectionString = StrCon
            .LookupDataset.ADODataset.CursorLocation = clUseClient
            .LookupDataset.ADODataset.LockType = ltReadOnly         ' Check Permmission
            .LookupDataset.ADODataset.CommandType = cmdText
            .LookupDataset.ADODataset.CommandText = "SELECT DEP_ID,DEP_NM FROM departemen " ' Like Join
            '.LookupDataset.Open
            .LookupKeyField = "DEP_ID"
            .LookupResultField = "DEP_NM"
            .ListColumns = "NAMA DEPT"
            .ListFieldName = "DEP_NM"
            .ListWidth = 800
            .DisplaySize = 400
End With
dxDBGrid_Payroll(1).Dataset.Open
End Sub

Private Sub Lookup_GridFill6()
'=================================
'-- PAYROLL_SESSION -> HARIAN --
'=================================
With dxDBGrid_Payroll(2).Columns.ColumnByName("Clm2").LookupColumn
'== DEPT NAME ---
            .LookupDataset.EnableControls
            '.LookupDataset.Close
            '.Dataset.Refresh
            .LookupDatasetType = dtADODataset
            .LookupDataset.ADODataset.ConnectionString = StrCon
            .LookupDataset.ADODataset.CursorLocation = clUseClient
            .LookupDataset.ADODataset.LockType = ltReadOnly         ' Check Permmission
            .LookupDataset.ADODataset.CommandType = cmdText
            .LookupDataset.ADODataset.CommandText = "SELECT DEP_ID,DEP_NM FROM departemen " ' Like Join
            '.LookupDataset.Open
            .LookupKeyField = "DEP_ID"
            .LookupResultField = "DEP_NM"
            .ListColumns = "NAMA DEPT"
            .ListFieldName = "DEP_NM"
            .ListWidth = 800
            .DisplaySize = 400
End With
dxDBGrid_Payroll(2).Dataset.Open
End Sub

Private Sub Lookup_GridFill4()
'=================================
'-- PAYROLL_SESSION -> HARIAN --
'=================================
With dxDBGrid_Payroll(0).Columns.ColumnByName("Clm2").LookupColumn
'== DEPT NAME ---
            .LookupDataset.EnableControls
            '.LookupDataset.Close
            '.Dataset.Refresh
            .LookupDatasetType = dtADODataset
            .LookupDataset.ADODataset.ConnectionString = StrCon
            .LookupDataset.ADODataset.CursorLocation = clUseClient
            .LookupDataset.ADODataset.LockType = ltReadOnly         ' Check Permmission
            .LookupDataset.ADODataset.CommandType = cmdText
            .LookupDataset.ADODataset.CommandText = "SELECT DEP_ID,DEP_NM FROM departemen " ' Like Join
            '.LookupDataset.Open
            .LookupKeyField = "DEP_ID"
            .LookupResultField = "DEP_NM"
            .ListColumns = "NAMA DEPT"
            .ListFieldName = "DEP_NM"
            .ListWidth = 800
            .DisplaySize = 400
End With
dxDBGrid_Payroll(0).Dataset.Open
End Sub

Private Sub Lookup_GridFill2()
'=================================
'-- PAYROLL_SESSION -> Presensi --
'=================================
    With dxDBGrid2.Columns.ColumnByName("Clm2").LookupColumn
    '== GROUP NAME ---
                .LookupDataset.EnableControls
                '.LookupDataset.Close
                '.Dataset.Refresh
                .LookupDatasetType = dtADODataset
                .LookupDataset.ADODataset.ConnectionString = StrCon
                .LookupDataset.ADODataset.CursorLocation = clUseClient
                .LookupDataset.ADODataset.LockType = ltReadOnly         ' Check Permmission
                .LookupDataset.ADODataset.CommandType = cmdText
                .LookupDataset.ADODataset.CommandText = "SELECT TT_GRP_ID,TT_GRP_NM FROM timetable_grp " ' Like Join
                '.LookupDataset.Open
                .LookupKeyField = "TT_GRP_ID"
                .LookupResultField = "TT_GRP_NM"
                .ListColumns = "GROUP TIME TABLE"
                .ListFieldName = "TT_GRP_NM"
                .ListWidth = 800
                .DisplaySize = 400
    End With
    With dxDBGrid2.Columns.ColumnByName("Clm3").LookupColumn
    '== DEPT NAME ---
                .LookupDataset.EnableControls
                '.LookupDataset.Close
                '.Dataset.Refresh
                .LookupDatasetType = dtADODataset
                .LookupDataset.ADODataset.ConnectionString = StrCon
                .LookupDataset.ADODataset.CursorLocation = clUseClient
                .LookupDataset.ADODataset.LockType = ltReadOnly         ' Check Permmission
                .LookupDataset.ADODataset.CommandType = cmdText
                .LookupDataset.ADODataset.CommandText = "SELECT DEP_ID,DEP_NM FROM departemen " ' Like Join
                '.LookupDataset.Open
                .LookupKeyField = "DEP_ID"
                .LookupResultField = "DEP_NM"
                .ListColumns = "NAMA DEPT"
                .ListFieldName = "DEP_NM"
                .ListWidth = 800
                .DisplaySize = 400
    End With
    With dxDBGrid2.Columns.ColumnByName("Clm4").LookupColumn
    '== CABANG NAME ---
                .LookupDataset.EnableControls
                '.LookupDataset.Close
                '.Dataset.Refresh
                .LookupDatasetType = dtADODataset
                .LookupDataset.ADODataset.ConnectionString = StrCon
                .LookupDataset.ADODataset.CursorLocation = clUseClient
                .LookupDataset.ADODataset.LockType = ltReadOnly         ' Check Permmission
                .LookupDataset.ADODataset.CommandType = cmdText
                .LookupDataset.ADODataset.CommandText = "SELECT CAB_ID,CAB_NM FROM cabang " ' Like Join
                '.LookupDataset.Open
                .LookupKeyField = "CAB_ID"
                .LookupResultField = "CAB_NM"
                .ListColumns = "NAMA CABANG"
                .ListFieldName = "CAB_NM"
                .ListWidth = 800
                .DisplaySize = 400
    End With
    dxDBGrid2.Dataset.Open
End Sub


Private Sub Lookup_GridFill3()
'=================================
'-- PAYROLL_SESSION -> Overtime --
'=================================
    With dxDBGrid3.Columns.ColumnByName("Clm2").LookupColumn
    '== GROUP NAME ---
                .LookupDataset.EnableControls
                '.LookupDataset.Close
                '.Dataset.Refresh
                .LookupDatasetType = dtADODataset
                .LookupDataset.ADODataset.ConnectionString = StrCon
                .LookupDataset.ADODataset.CursorLocation = clUseClient
                .LookupDataset.ADODataset.LockType = ltReadOnly         ' Check Permmission
                .LookupDataset.ADODataset.CommandType = cmdText
                .LookupDataset.ADODataset.CommandText = "SELECT TT_GRP_ID,TT_GRP_NM FROM timetable_grp " ' Like Join
                '.LookupDataset.Open
                .LookupKeyField = "TT_GRP_ID"
                .LookupResultField = "TT_GRP_NM"
                .ListColumns = "GROUP TIME TABLE"
                .ListFieldName = "TT_GRP_NM"
                .ListWidth = 800
                .DisplaySize = 400
    End With
    With dxDBGrid3.Columns.ColumnByName("Clm3").LookupColumn
    '== DEPT NAME ---
                .LookupDataset.EnableControls
                '.LookupDataset.Close
                '.Dataset.Refresh
                .LookupDatasetType = dtADODataset
                .LookupDataset.ADODataset.ConnectionString = StrCon
                .LookupDataset.ADODataset.CursorLocation = clUseClient
                .LookupDataset.ADODataset.LockType = ltReadOnly         ' Check Permmission
                .LookupDataset.ADODataset.CommandType = cmdText
                .LookupDataset.ADODataset.CommandText = "SELECT DEP_ID,DEP_NM FROM departemen " ' Like Join
                '.LookupDataset.Open
                .LookupKeyField = "DEP_ID"
                .LookupResultField = "DEP_NM"
                .ListColumns = "NAMA DEPT"
                .ListFieldName = "DEP_NM"
                .ListWidth = 800
                .DisplaySize = 400
    End With
    With dxDBGrid3.Columns.ColumnByName("Clm4").LookupColumn
    '== CABANG NAME ---
                .LookupDataset.EnableControls
                '.LookupDataset.Close
                '.Dataset.Refresh
                .LookupDatasetType = dtADODataset
                .LookupDataset.ADODataset.ConnectionString = StrCon
                .LookupDataset.ADODataset.CursorLocation = clUseClient
                .LookupDataset.ADODataset.LockType = ltReadOnly         ' Check Permmission
                .LookupDataset.ADODataset.CommandType = cmdText
                .LookupDataset.ADODataset.CommandText = "SELECT CAB_ID,CAB_NM FROM cabang " ' Like Join
                '.LookupDataset.Open
                .LookupKeyField = "CAB_ID"
                .LookupResultField = "CAB_NM"
                .ListColumns = "NAMA CABANG"
                .ListFieldName = "CAB_NM"
                .ListWidth = 800
                .DisplaySize = 400
    End With
    dxDBGrid3.Dataset.Open
End Sub


