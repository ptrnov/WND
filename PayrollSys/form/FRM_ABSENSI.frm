VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Begin VB.Form FRM_ABSENSI 
   BackColor       =   &H00FFFFFF&
   Caption         =   "EMPLOYE ABSENSI PROCESS Ver1.1"
   ClientHeight    =   10275
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
   LockControls    =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10275
   ScaleWidth      =   20160
   Tag             =   "Terminal Machine Finger NameId"
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.GroupBox GroupBox5 
      Height          =   1155
      Left            =   10440
      TabIndex        =   28
      Top             =   50
      Width           =   3495
      _Version        =   851970
      _ExtentX        =   6165
      _ExtentY        =   2028
      _StockProps     =   79
      Caption         =   "Global Filter"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.ComboBox CmbAbsensi 
         Height          =   360
         Index           =   2
         Left            =   960
         TabIndex        =   66
         Top             =   240
         Width           =   2415
         _Version        =   851970
         _ExtentX        =   4260
         _ExtentY        =   635
         _StockProps     =   77
         BackColor       =   -2147483643
         Sorted          =   -1  'True
         AutoComplete    =   -1  'True
         DropDownWidth   =   5000
      End
      Begin XtremeSuiteControls.ComboBox CmbAbsensi 
         Height          =   360
         Index           =   3
         Left            =   960
         TabIndex        =   67
         Top             =   720
         Width           =   2415
         _Version        =   851970
         _ExtentX        =   4260
         _ExtentY        =   635
         _StockProps     =   77
         BackColor       =   -2147483643
         Sorted          =   -1  'True
         AutoComplete    =   -1  'True
         DropDownWidth   =   5000
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Group."
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   65
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   64
         Top             =   360
         Width           =   975
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox4 
      Height          =   135
      Left            =   15720
      TabIndex        =   27
      Top             =   120
      Width           =   30
      _Version        =   851970
      _ExtentX        =   53
      _ExtentY        =   238
      _StockProps     =   79
      Caption         =   "GroupBox4"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1215
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   10275
      _Version        =   851970
      _ExtentX        =   18124
      _ExtentY        =   2143
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   915
         Left            =   3960
         TabIndex        =   23
         Top             =   160
         Width           =   4815
         _Version        =   851970
         _ExtentX        =   8493
         _ExtentY        =   1614
         _StockProps     =   79
         Caption         =   "Time Range"
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.DateTimePicker Tgl 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   390
            Width           =   2055
            _Version        =   851970
            _ExtentX        =   3625
            _ExtentY        =   661
            _StockProps     =   68
         End
         Begin XtremeSuiteControls.DateTimePicker Tgl 
            Height          =   375
            Index           =   1
            Left            =   2880
            TabIndex        =   3
            Top             =   390
            Width           =   1815
            _Version        =   851970
            _ExtentX        =   3201
            _ExtentY        =   661
            _StockProps     =   68
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "TO"
            BeginProperty Font 
               Name            =   "Bauhaus 93"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   2280
            TabIndex        =   26
            Top             =   480
            Width           =   375
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
            Left            =   3720
            TabIndex        =   25
            Top             =   190
            Width           =   975
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
            Left            =   1200
            TabIndex        =   24
            Top             =   190
            Width           =   1095
         End
      End
      Begin XtremeSuiteControls.PushButton cmdProcess 
         Height          =   855
         Index           =   0
         Left            =   8880
         TabIndex        =   4
         Top             =   240
         Width           =   1215
         _Version        =   851970
         _ExtentX        =   2143
         _ExtentY        =   1508
         _StockProps     =   79
         Caption         =   "Calculate"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.ComboBox CmbAbsensi 
         Height          =   360
         Index           =   1
         Left            =   960
         TabIndex        =   1
         Top             =   720
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
      Begin XtremeSuiteControls.ComboBox CmbAbsensi 
         Height          =   360
         Index           =   0
         Left            =   960
         TabIndex        =   0
         Top             =   240
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
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
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
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
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
         Left            =   360
         TabIndex        =   18
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   9135
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   19935
      _Version        =   851970
      _ExtentX        =   35163
      _ExtentY        =   16113
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
      Appearance      =   3
      Color           =   32
      PaintManager.BoldSelected=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.FixedTabWidth=   200
      PaintManager.MaxTabWidth=   250
      PaintManager.MinTabWidth=   250
      PaintManager.ControlMargin=   "1,0,0,0"
      ItemCount       =   4
      SelectedItem    =   1
      Item(0).Caption =   "Calculate Persensi In/Out"
      Item(0).ControlCount=   8
      Item(0).Control(0)=   "dxDBGrid1"
      Item(0).Control(1)=   "DFilter(0)"
      Item(0).Control(2)=   "DFilter(1)"
      Item(0).Control(3)=   "CmbFilter(0)"
      Item(0).Control(4)=   "CmbFilter(1)"
      Item(0).Control(5)=   "ProgressBar1"
      Item(0).Control(6)=   "cmdFilterProcess(0)"
      Item(0).Control(7)=   "cmdRptProcess(0)"
      Item(1).Caption =   "Presensi Maintain"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "dxDBGrid2"
      Item(1).Control(1)=   "GroupBox2"
      Item(2).Caption =   "OverTime Leveling"
      Item(2).ControlCount=   8
      Item(2).Control(0)=   "dxDBGrid3"
      Item(2).Control(1)=   "DFilter(2)"
      Item(2).Control(2)=   "CmbFilter(2)"
      Item(2).Control(3)=   "DFilter(3)"
      Item(2).Control(4)=   "CmbFilter(3)"
      Item(2).Control(5)=   "cmdFilterProcess(1)"
      Item(2).Control(6)=   "cmdRptProcess(1)"
      Item(2).Control(7)=   "cmdFilterProcess(3)"
      Item(3).Caption =   "Calculate Closing Session"
      Item(3).ControlCount=   8
      Item(3).Control(0)=   "dxDBGrid4"
      Item(3).Control(1)=   "cmdRptProcess(2)"
      Item(3).Control(2)=   "DFilter(4)"
      Item(3).Control(3)=   "CmbFilter(4)"
      Item(3).Control(4)=   "DFilter(5)"
      Item(3).Control(5)=   "cmdFilterProcess(2)"
      Item(3).Control(6)=   "CmbFilter(5)"
      Item(3).Control(7)=   "cmdProcess(5)"
      Begin XtremeSuiteControls.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   -57880
         TabIndex        =   30
         Top             =   -600
         Visible         =   0   'False
         Width           =   3135
         _Version        =   851970
         _ExtentX        =   5530
         _ExtentY        =   450
         _StockProps     =   93
      End
      Begin XtremeSuiteControls.DateTimePicker DFilter 
         Height          =   375
         Index           =   0
         Left            =   -60160
         TabIndex        =   9
         ToolTipText     =   "Pilih Tanggal Mulai"
         Top             =   480
         Visible         =   0   'False
         Width           =   2055
         _Version        =   851970
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   68
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   8415
         Left            =   15720
         TabIndex        =   21
         Top             =   480
         Width           =   3975
         _Version        =   851970
         _ExtentX        =   7011
         _ExtentY        =   14843
         _StockProps     =   79
         Caption         =   "Manual Repair and Check LOG."
         ForeColor       =   32768
         BackColor       =   12640511
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   2
         Begin XtremeSuiteControls.ComboBox Cmb_LogFrm 
            Height          =   360
            Index           =   0
            Left            =   240
            TabIndex        =   35
            Top             =   1440
            Width           =   3495
            _Version        =   851970
            _ExtentX        =   6165
            _ExtentY        =   635
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.DateTimePicker DT_LogFrm 
            Height          =   375
            Left            =   240
            TabIndex        =   34
            Top             =   720
            Width           =   1935
            _Version        =   851970
            _ExtentX        =   3413
            _ExtentY        =   661
            _StockProps     =   68
         End
         Begin XtremeSuiteControls.PushButton cmdManual 
            Height          =   615
            Index           =   1
            Left            =   240
            TabIndex        =   22
            Top             =   4920
            Width           =   2055
            _Version        =   851970
            _ExtentX        =   3625
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Register Employe Finger"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Rounded MT Bold"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
         End
         Begin XtremeSuiteControls.PushButton cmdManual 
            Height          =   1095
            Index           =   0
            Left            =   2640
            TabIndex        =   31
            Top             =   2880
            Width           =   1095
            _Version        =   851970
            _ExtentX        =   1931
            _ExtentY        =   1931
            _StockProps     =   79
            Caption         =   "+/-"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Wide Latin"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
         End
         Begin XtremeSuiteControls.PushButton cmdManual 
            Height          =   615
            Index           =   2
            Left            =   240
            TabIndex        =   32
            Top             =   5640
            Width           =   2055
            _Version        =   851970
            _ExtentX        =   3625
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Terminal Finger Machine"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Rounded MT Bold"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
         End
         Begin XtremeSuiteControls.PushButton cmdManual 
            Height          =   615
            Index           =   3
            Left            =   2400
            TabIndex        =   33
            Top             =   600
            Width           =   1335
            _Version        =   851970
            _ExtentX        =   2355
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "FILTER"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Rounded MT Bold"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
         End
         Begin XtremeSuiteControls.ComboBox Cmb_LogFrm 
            Height          =   360
            Index           =   1
            Left            =   240
            TabIndex        =   36
            Top             =   2160
            Width           =   3495
            _Version        =   851970
            _ExtentX        =   6165
            _ExtentY        =   635
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.ComboBox Cmb_LogFrm 
            Height          =   360
            Index           =   2
            Left            =   240
            TabIndex        =   40
            Top             =   2880
            Width           =   2295
            _Version        =   851970
            _ExtentX        =   4048
            _ExtentY        =   635
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.ComboBox Cmb_LogFrm 
            Height          =   360
            Index           =   3
            Left            =   240
            TabIndex        =   58
            Top             =   3600
            Width           =   2295
            _Version        =   851970
            _ExtentX        =   4048
            _ExtentY        =   635
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.ComboBox Cmb_LogFrm 
            Height          =   360
            Index           =   4
            Left            =   240
            TabIndex        =   60
            Top             =   4320
            Width           =   3495
            _Version        =   851970
            _ExtentX        =   6165
            _ExtentY        =   635
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Cabang"
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
            Index           =   12
            Left            =   240
            TabIndex        =   62
            Top             =   3360
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Finger Id"
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
            Index           =   10
            Left            =   240
            TabIndex        =   59
            Top             =   4080
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Golongan"
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
            Index           =   9
            Left            =   240
            TabIndex        =   41
            Top             =   2640
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Employe"
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
            Index           =   8
            Left            =   240
            TabIndex        =   39
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
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
            Index           =   7
            Left            =   240
            TabIndex        =   38
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "* Date"
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
            Left            =   240
            TabIndex        =   37
            Top             =   480
            Width           =   1095
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            X1              =   240
            X2              =   3720
            Y1              =   4800
            Y2              =   4800
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid2 
         Height          =   8175
         Left            =   120
         OleObjectBlob   =   "FRM_ABSENSI.frx":0000
         TabIndex        =   20
         Top             =   720
         Width           =   15435
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
         Height          =   8175
         Left            =   -69880
         OleObjectBlob   =   "FRM_ABSENSI.frx":0CA8
         TabIndex        =   14
         Top             =   840
         Visible         =   0   'False
         Width           =   19575
      End
      Begin XtremeSuiteControls.ComboBox CmbFilter 
         Height          =   360
         Index           =   1
         Left            =   -53200
         TabIndex        =   12
         ToolTipText     =   "Pilih Nama Karyawan"
         Top             =   480
         Visible         =   0   'False
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
      Begin XtremeSuiteControls.DateTimePicker DFilter 
         Height          =   375
         Index           =   1
         Left            =   -58120
         TabIndex        =   10
         ToolTipText     =   "Pilih Tanggal Akhir"
         Top             =   480
         Visible         =   0   'False
         Width           =   2055
         _Version        =   851970
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   68
      End
      Begin XtremeSuiteControls.PushButton cmdFilterProcess 
         Height          =   375
         Index           =   0
         Left            =   -61000
         TabIndex        =   15
         ToolTipText     =   "Tekan Untuk Menampilkan Flter"
         Top             =   480
         Visible         =   0   'False
         Width           =   855
         _Version        =   851970
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Filter"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.ComboBox CmbFilter 
         Height          =   360
         Index           =   0
         Left            =   -56080
         TabIndex        =   11
         ToolTipText     =   "Pilih Departement"
         Top             =   480
         Visible         =   0   'False
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
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid3 
         Height          =   8055
         Left            =   -69880
         OleObjectBlob   =   "FRM_ABSENSI.frx":1950
         TabIndex        =   42
         Top             =   840
         Visible         =   0   'False
         Width           =   19515
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid4 
         Height          =   8055
         Left            =   -69880
         OleObjectBlob   =   "FRM_ABSENSI.frx":25F8
         TabIndex        =   43
         Top             =   840
         Visible         =   0   'False
         Width           =   19515
      End
      Begin XtremeSuiteControls.PushButton cmdRptProcess 
         Height          =   375
         Index           =   0
         Left            =   -69880
         TabIndex        =   44
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
      Begin XtremeSuiteControls.PushButton cmdRptProcess 
         Height          =   375
         Index           =   1
         Left            =   -69880
         TabIndex        =   45
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
      Begin XtremeSuiteControls.PushButton cmdRptProcess 
         Height          =   375
         Index           =   2
         Left            =   -53800
         TabIndex        =   46
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
      Begin XtremeSuiteControls.DateTimePicker DFilter 
         Height          =   375
         Index           =   2
         Left            =   -60160
         TabIndex        =   47
         ToolTipText     =   "Pilih Tanggal Mulai"
         Top             =   480
         Visible         =   0   'False
         Width           =   2055
         _Version        =   851970
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   68
      End
      Begin XtremeSuiteControls.ComboBox CmbFilter 
         Height          =   360
         Index           =   3
         Left            =   -53200
         TabIndex        =   48
         ToolTipText     =   "Pilih Nama Karyawan"
         Top             =   480
         Visible         =   0   'False
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
      Begin XtremeSuiteControls.DateTimePicker DFilter 
         Height          =   375
         Index           =   3
         Left            =   -58120
         TabIndex        =   49
         ToolTipText     =   "Pilih Tanggal Akhir"
         Top             =   480
         Visible         =   0   'False
         Width           =   2055
         _Version        =   851970
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   68
      End
      Begin XtremeSuiteControls.PushButton cmdFilterProcess 
         Height          =   375
         Index           =   1
         Left            =   -61000
         TabIndex        =   50
         ToolTipText     =   "Tekan Untuk Menampilkan Flter"
         Top             =   480
         Visible         =   0   'False
         Width           =   855
         _Version        =   851970
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Filter"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.ComboBox CmbFilter 
         Height          =   360
         Index           =   2
         Left            =   -56080
         TabIndex        =   51
         ToolTipText     =   "Pilih Departement"
         Top             =   480
         Visible         =   0   'False
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
      Begin XtremeSuiteControls.DateTimePicker DFilter 
         Height          =   375
         Index           =   4
         Left            =   -69880
         TabIndex        =   52
         ToolTipText     =   "Pilih Tanggal Mulai"
         Top             =   480
         Visible         =   0   'False
         Width           =   2055
         _Version        =   851970
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   68
      End
      Begin XtremeSuiteControls.ComboBox CmbFilter 
         Height          =   360
         Index           =   4
         Left            =   -65800
         TabIndex        =   53
         ToolTipText     =   "Pilih Nama Karyawan"
         Top             =   480
         Visible         =   0   'False
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
      Begin XtremeSuiteControls.DateTimePicker DFilter 
         Height          =   375
         Index           =   5
         Left            =   -67840
         TabIndex        =   54
         ToolTipText     =   "Pilih Tanggal Akhir"
         Top             =   480
         Visible         =   0   'False
         Width           =   2055
         _Version        =   851970
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   68
      End
      Begin XtremeSuiteControls.PushButton cmdFilterProcess 
         Height          =   375
         Index           =   2
         Left            =   -60040
         TabIndex        =   55
         ToolTipText     =   "Tekan Untuk Menampilkan Flter"
         Top             =   480
         Visible         =   0   'False
         Width           =   855
         _Version        =   851970
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Filter"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.ComboBox CmbFilter 
         Height          =   360
         Index           =   5
         Left            =   -62920
         TabIndex        =   56
         ToolTipText     =   "Pilih Departement"
         Top             =   480
         Visible         =   0   'False
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
      Begin XtremeSuiteControls.PushButton cmdFilterProcess 
         Height          =   375
         Index           =   3
         Left            =   -62560
         TabIndex        =   57
         ToolTipText     =   "Tekan Untuk Menampilkan Flter"
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
         _Version        =   851970
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "OT Process"
         Enabled         =   0   'False
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton cmdProcess 
         Height          =   375
         Index           =   5
         Left            =   -52300
         TabIndex        =   63
         Top             =   480
         Visible         =   0   'False
         Width           =   1935
         _Version        =   851970
         _ExtentX        =   3413
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Send To payroll >>"
         Appearance      =   6
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox6 
      Height          =   855
      Left            =   14040
      TabIndex        =   29
      Top             =   120
      Width           =   4815
      _Version        =   851970
      _ExtentX        =   8493
      _ExtentY        =   1508
      _StockProps     =   79
      Caption         =   "Check"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton cmdProcess 
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   975
         _Version        =   851970
         _ExtentX        =   1720
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Off Days"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton cmdProcess 
         Height          =   495
         Index           =   2
         Left            =   1120
         TabIndex        =   6
         Top             =   240
         Width           =   1095
         _Version        =   851970
         _ExtentX        =   1931
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Exception"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton cmdProcess 
         Height          =   495
         Index           =   3
         Left            =   2250
         TabIndex        =   7
         Top             =   240
         Width           =   1215
         _Version        =   851970
         _ExtentX        =   2143
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "TimeTable"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton cmdProcess 
         Height          =   495
         Index           =   4
         Left            =   3500
         TabIndex        =   8
         Top             =   240
         Width           =   1215
         _Version        =   851970
         _ExtentX        =   2143
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Group Emp."
         Appearance      =   6
      End
   End
   Begin XtremeSuiteControls.PushButton cmdProcess 
      Height          =   855
      Index           =   6
      Left            =   18960
      TabIndex        =   17
      Top             =   240
      Width           =   1095
      _Version        =   851970
      _ExtentX        =   1931
      _ExtentY        =   1508
      _StockProps     =   79
      Caption         =   "Exit"
      Appearance      =   6
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Terminal Machine Finger NameId"
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
      Left            =   0
      TabIndex        =   61
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "FRM_ABSENSI"
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
Dim CLM0, CLM1, CLM2, clm3, clm4, BND0, BND1, BND2, BND3, BND4 As Variant
Dim Qry0 As String, Qry1 As String, Qry2 As String, Qry3 As String, Qry4 As String, QryStt As String
Dim AryLog As String
Dim LoadWas As Boolean
Dim STgl As Date, ETgl As Date, SFiter As Date
Dim PERSENSI_STGL As Integer, FilterDateTgl As Integer, F_TGL As String, F_WAKTU As String, F_TGL_WAKTU As String
Dim x As Integer
Dim trecs As Double
Dim GrpTT As Integer, DepId As Integer, KarId As String, sDate As String, eDate As String
Dim i As Integer
Dim Ulang As Boolean
                      
Private Sub CmbAbsensi_Click(Index As Integer)
'=====================================
'------------- ptr.nov ---------------
'===== PROCESS DEP-EMPLOYE CHANGE ====
'=====================================
On Error Resume Next
If Index = 0 Then
    If CmbAbsensi(0).Text <> "" Then
        prjSysID.Employe CmbAbsensi(1), CmbAbsensi(2).Text, CmbAbsensi(3).Text, CmbAbsensi(0).Text
    End If
End If
If Index = 2 Or Index = 3 Then
    If CmbAbsensi(0).Text <> "" Then
        prjSysID.Employe CmbAbsensi(1), CmbAbsensi(2).Text, CmbAbsensi(3).Text, CmbAbsensi(0).Text
    End If
    If cmbFilter(0).Text <> "" Then
        prjSysID.Employe cmbFilter(1), CmbAbsensi(2).Text, CmbAbsensi(3).Text, cmbFilter(0).Text
    End If
    If cmbFilter(2).Text <> "" Then
        prjSysID.Employe cmbFilter(3), CmbAbsensi(2).Text, CmbAbsensi(3).Text, cmbFilter(2).Text
    End If
    If cmbFilter(4).Text <> "" Then
        prjSysID.Employe cmbFilter(5), CmbAbsensi(2).Text, CmbAbsensi(3).Text, cmbFilter(4).Text
    End If
    If Cmb_LogFrm(0).Text <> "" Then
        prjSysID.Employe Cmb_LogFrm(1), CmbAbsensi(2).Text, CmbAbsensi(3).Text, Cmb_LogFrm(0).Text
    End If
End If
End Sub

'====================================
'============= ptr.nov ==============
'========= COMMAND BUTTON MENU ======
'====================================
Private Sub cmdProcess_Click(Index As Integer)
 Select Case Index
    Case Is = 0
            '============================
            '----- PROCESS PERSENSI -----
            '============================
            '-- FRM_ABSENSI_PBAR.Show
             i = 0
             GrpTT = 0
             DepId = 0
             KarId = 0
             sDate = 0
             eDate = 0
             
             If CmbAbsensi(0).Text = 0 And CmbAbsensi(1).Text = "0" Then
                Ulang = True
                ElseIf CmbAbsensi(0).Text = 0 And CmbAbsensi(1).Text <> "0" Then Ulang = False
                ElseIf CmbAbsensi(0).Text <> 0 And CmbAbsensi(1).Text = "0" Then Ulang = True
                ElseIf CmbAbsensi(0).Text <> 0 And CmbAbsensi(1).Text <> "0" Then Ulang = False
             End If
           
           FRM_ABSENSI_PBAR.Show
            If (Ulang = True) Then
                With FRM_ABSENSI_PBAR.ProgressBar1(0)
                    
                    .Visible = True
                    .Min = 0
                    .Max = (CmbAbsensi(1).ListCount - 1)
                    
                    For i = 1 To CmbAbsensi(1).ListCount - 1
                        KarId = CmbAbsensi(1).List(i)
                        DepId = CmbAbsensi(0).Text
                        sDate = Tgl(0).Value
                        eDate = Tgl(1).Value
                        
                        Qry1 = "presensi_inout('0=" & Format(sDate, "yyyy-mm-dd") & "=" & Format(eDate, "yyyy-mm-dd") & "=" & KarId & "')"
                        OpRec1 Qry1, True
                        Qry2 = "presensi_inout('1=" & Format(sDate, "yyyy-mm-dd") & "=" & Format(eDate, "yyyy-mm-dd") & "=" & KarId & "')"
                        OpRec2 Qry2, True
                        Qry3 = "presensi_inout('2=" & Format(sDate, "yyyy-mm-dd") & "=" & Format(eDate, "yyyy-mm-dd") & "=" & KarId & "')"
                        OpRec3 Qry3, True
                        Qry1 = "presensi_inout('3=" & Format(sDate, "yyyy-mm-dd") & "=" & Format(eDate, "yyyy-mm-dd") & "=" & KarId & "')"
                        OpRec1 Qry1, True
                       FRM_ABSENSI_PBAR.Enabled = True
                       FRM_ABSENSI_PBAR.ProgressBar1(0).Value = i
                       FRM_ABSENSI_PBAR.Enabled = False
                     Next i
                 End With
             Else
                FRM_ABSENSI_PBAR.Enabled = False
                KarId = CmbAbsensi(1).Text
                DepId = CmbAbsensi(0).Text
                sDate = Tgl(0).Value
                eDate = Tgl(1).Value
                Qry1 = "presensi_inout('0=" & Format(sDate, "yyyy-mm-dd") & "=" & Format(eDate, "yyyy-mm-dd") & "=" & KarId & "')"
                OpRec1 Qry1, True
                'Qry2 = "presensi_inout('1=" & Format(sDate, "yyyy-mm-dd") & "=" & Format(eDate, "yyyy-mm-dd") & "=" & KarId & "')"
                'OpRec2 Qry2, True
                Qry3 = "presensi_inout('2=" & Format(sDate, "yyyy-mm-dd") & "=" & Format(eDate, "yyyy-mm-dd") & "=" & KarId & "')"
                OpRec3 Qry3, True
                Qry1 = "presensi_inout('3=" & Format(sDate, "yyyy-mm-dd") & "=" & Format(eDate, "yyyy-mm-dd") & "=" & KarId & "')"
                OpRec1 Qry1, True
                FRM_ABSENSI_PBAR.ProgressBar1(0).Value = 1
                FRM_ABSENSI_PBAR.ProgressBar1(0).Value = FRM_ABSENSI_PBAR.ProgressBar1(0).Max
             End If
            
            GridFill 0
            '-- pointer currsor loading
            Screen.MousePointer = vbIconPointer
            FRM_ABSENSI_PBAR.Enabled = True
            
    Case Is = 1
            '========================
            '----- FORM OFF DAY -----
            '========================
            FRM_ABSENSI_OFFDAY.Show
            
    Case Is = 2
            '=============================
            '----- FORM IJIN EMPLOYE -----
            '=============================
            FRM_ABSENSI_IJIN.Show
            
    Case Is = 3
            '===========================
            '----- FORM TIME TABLE -----
            '===========================
            FRM_ABSENSI_TT.Show
            
    Case Is = 4
            '=========================================
            '----- FORM EMPLOYE GROUP TIME TABLE -----
            '=========================================
            FRM_ABSENSI_EMPGRP.Show
       
    Case Is = 5
            '=======================================
            '----- PRESENSI SESSION TO PAYROLL -----
            '=======================================
             i = 0
             GrpTT = 0
             DepId = 0
             KarId = 0
             sDate = 0
             eDate = 0
             
             If cmbFilter(4).Text = 0 And cmbFilter(5).Text = "0" Then
                Ulang = True
                ElseIf cmbFilter(4).Text = 0 And cmbFilter(5).Text <> "0" Then Ulang = False
                ElseIf cmbFilter(4).Text <> 0 And cmbFilter(5).Text = "0" Then Ulang = True
                ElseIf cmbFilter(4).Text <> 0 And cmbFilter(5).Text <> "0" Then Ulang = False
             End If
           
           FRM_ABSENSI_PBAR.Show
           FRM_ABSENSI_PBAR.Caption = " Presensi Session Process "
            If (Ulang = True) Then
                With FRM_ABSENSI_PBAR.ProgressBar1(0)
                    
                    .Visible = True
                    .Min = 0
                    .Max = (CmbAbsensi(1).ListCount - 1)
                    
                    For i = 1 To CmbAbsensi(1).ListCount - 1
                        KarId = cmbFilter(5).List(i)
                        sDate = DFilter(4).Value
                        eDate = DFilter(5).Value
                        '=============
                        '--- insert --
                        '=============
                        Qry1 = "presensi_inout('4=" & Format(sDate, "yyyy-mm-dd") & _
                                            "=" & Format(eDate, "yyyy-mm-dd") & _
                                            "=" & KarId & "')"
                        OpRec1 Qry1, True
                        
                       FRM_ABSENSI_PBAR.Enabled = True
                       FRM_ABSENSI_PBAR.ProgressBar1(0).Value = i
                       FRM_ABSENSI_PBAR.Enabled = False
                     Next i
                 End With
             Else
                FRM_ABSENSI_PBAR.Enabled = False
                KarId = cmbFilter(5).Text
                sDate = DFilter(4).Value
                eDate = DFilter(5).Value
                    '=============
                    '--- insert --
                    '=============
                    Qry1 = "presensi_inout('4=" & Format(sDate, "yyyy-mm-dd") & _
                                        "=" & Format(eDate, "yyyy-mm-dd") & _
                                        "=" & KarId & "')"
                    OpRec1 Qry1, True
                FRM_ABSENSI_PBAR.ProgressBar1(0).Value = 1
                FRM_ABSENSI_PBAR.ProgressBar1(0).Value = FRM_ABSENSI_PBAR.ProgressBar1(0).Max
             End If
            '-- pointer currsor loading
            Screen.MousePointer = vbIconPointer
            FRM_ABSENSI_PBAR.Enabled = True
            
   Case Is = 6
            '==============================
            '----- EXIT COMMAN BUTTON -----
            '==============================
            Unload FRM_ABSENSI_OFFDAY
            Unload FRM_ABSENSI_IJIN
            Unload FRM_ABSENSI_TT
            Unload FRM_ABSENSI_EMPGRP
            Unload FRM_ABSENSI_FINGERGRP
            Unload FRM_ABSENSI_LOG
            Unload FRM_ABSENSI_MACHINE
            Unload FRM_ABSENSI_REPORT
            Unload Me
End Select
End Sub


Private Sub cmdFilterProcess_Click(Index As Integer)
'=================================
'------------ ptr.nov ------------
'------ COMMAND BUTTON FILTER ----
'=================================
Select Case Index
    Case Is = 0
        '==========================================
        '----- PROCESS PERSENSI IN/OUT FILTER -----
        '==========================================
        GridFill 0

    Case Is = 1
        '===================================
        '----- PROCESS OVERTIME FILTER -----
        '===================================
        GridFill 2
    Case Is = 2
        '==========================================
        '----- PROCESS PERSENSI SESSION FILTER ----
        '==========================================
         GridFill 3
End Select
End Sub

Public Sub cmdManual_Click(Index As Integer)
'=================================
'------------ ptr.nov ------------
'------ COMMAND BUTTON LOG -------
'=================================
Select Case Index
    Case Is = 0
            '====================================================
            '----- CHANGE LOG PERSONAL INSERT/UPDATE/DELETE -----
            '====================================================
            FRM_ABSENSI_LOG.Show
            FRM_ABSENSI_LOG.TglLogEdit(0).Value = Now
            FRM_ABSENSI_LOG.TglLogEdit(1).Value = Now
            FRM_ABSENSI_LOG.LogProfile Cmb_LogFrm(1).Text, DT_LogFrm.Value, DT_LogFrm.Value
    Case Is = 1
            '================================
            '----- GROUP FINGER EMPLOYE -----
            '================================
            FRM_ABSENSI_FINGERGRP.Show
    Case Is = 2
            '==========================
            '----- FINGER MACHINE -----
            '==========================
            FRM_ABSENSI_MACHINE.Show
    Case Is = 3
            '===============================
            '----- FILTER LOG PERSINSI -----
            '===============================
            GridFill 1
End Select
End Sub

Private Sub cmdRptProcess_Click(Index As Integer)
'=============================
'------------ ptr.nov --------
'------ COMMAND REPORT -------
'=============================
Select Case Index
    Case Is = 0
            '==================================
            '----- PROCESS ABSENSI FILTER -----
            '==================================
            FRM_ABSENSI_REPORT.Show
            'prjSysReport.Prensensi_IN_OUT Qry0
    Case Is = 1
            '===================================
            '----- PROCESS LEMBURAN FILTER -----
            '===================================
            FRM_ABSENSI_REPORT.Show
            'prjSysReport.Prensensi_OVER_TIME Qry1
    Case Is = 2
            '=========================================
            '----- PROCESS PERSENSI MONTH FILTER -----
            '=========================================
            FRM_ABSENSI_REPORT.Show
End Select
End Sub

Private Sub dxDBGrid2_OnDblClick()
'=========================
'--- CRUD LOG MAINTAIN ---
'=========================
With dxDBGrid2.Dataset
    If dxDBGrid2.Dataset.RecNo <> 0 Then
        FRM_ABSENSI_LOG.Show
        FRM_ABSENSI_LOG.TglLogEdit(0).Value = DateValue(.FieldValues("DateTime"))
        FRM_ABSENSI_LOG.TglLogEdit(1).Value = DateValue(.FieldValues("DateTime"))
        FRM_ABSENSI_LOG.LogProfile .FieldValues("KAR_ID"), DateValue(.FieldValues("DateTime")), DateValue(.FieldValues("DateTime"))
    End If
End With
End Sub

Private Sub Form_Load()
'===========================
'-------- ptr.nov ----------
'----- LOAD FIRST FORM -----
'===========================
'On Error GoTo ErrorLabel
    Main
    prjSysID.TTGROUP CmbAbsensi(3)
     prjSysID.CABANG CmbAbsensi(2)
   
     prjSysID.Dept CmbAbsensi(0)
  
    'MsgBox CmbAbsensi(2).Text
    init_Setting
    init_DT_CMB_Absensi
    init_DT_CMB_filter
    GridFill 0
    GridFill 1
    GridFill 2
    GridFill 3
    GridFill 4
    Me.Icon = LoadPicture(ImagePath("FRM_FRM_ABSENSI"))
'LOADING.SetParm Me, 100
'Exit Sub
'ErrorLabel:
'    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error", vbCritical, "LG Error"
'    LOADING.SetParm Me, 100

End Sub

Private Sub MenuGrid(TabSel As Integer)
'=====================================
'------------ ptr.nov ----------------
'----- MENU CULUMN BND GRID LODING ---
'=====================================
Select Case TabSel
Case Is = 0
     '===========================
     '--- MENU PRESENSI_INOUT ---
     '===========================
    BND0 = Array("TIME TABLE ", "PERSONAL ABSENSI VALUES")
    CLM0 = Array(Array("Clm0", "EmployeID", "KAR_ID", gedTextEdit, 0, 0, 100, 1, 0, 0, 0), _
                Array("Clm1", "Employe Name", "KAR_NM", gedTextEdit, 0, 0, 100, 1, 0, 0, 0), _
                Array("Clm2", "Date", "TGL_RUN", gedTextEdit, 0, 0, 70, 1, 0, 0, 0), _
                Array("Clm3", "Group", "GRP_NM", gedTextEdit, 0, 0, 50, 1, 0, 0, 0), _
                Array("Clm4", "Rule In.", "RULE_IN", gedTimeEdit, 0, 0, 50, 1, 0, 0, 1), _
                Array("Clm5", "Rule Out.", "RULE_OUT", gedTextEdit, 0, 0, 60, 1, 0, 0, 1), _
                Array("Clm6", "Day", "HARI", gedTextEdit, 0, 0, 60, 0, 1, 1, 0), _
                Array("Clm7", "Time In.", "DT_IN", gedTimeEdit, 0, 1, 110, 4, 0, 0, 2), _
                Array("Clm8", "Time out.", "DT_OUT", gedTimeEdit, 0, 1, 110, 4, 0, 0, 2), _
                Array("Clm0", "Work Time", "WORK_TIME", gedTextEdit, 0, 1, 60, 4, 2, 0, 0), _
                Array("Clm10", "Day Active", "DAY_ACTIVE", gedCheckEdit, 0, 1, 60, 4, 2, 0, 0), _
                Array("Clm11", "Holiday", "HOLIDAY", gedCheckEdit, 0, 1, 50, 4, 2, 0, 0), _
                Array("Clm12", "Exception", "DAY_EXCP_ID", gedCheckEdit, 0, 1, 60, 4, 2, 0, 0), _
                Array("Clm13", "Absen", "DAY_ABSENT", gedCheckEdit, 0, 1, 40, 4, 2, 0, 0), _
                Array("Clm14", "Late", "LATE", gedTextEdit, 0, 1, 60, 4, 2, 0, 1), _
                Array("Clm15", "Early", "EARLY", gedTextEdit, 0, 1, 60, 4, 2, 0, 1), _
                Array("Clm16", "OT Depan", "OT_DPN", gedTextEdit, 0, 1, 60, 10, 2, 0, 1), _
                Array("Clm17", "OT Hari", "OT_DPN_HARI", gedTextEdit, 0, 1, 60, 10, 1, 0, 0), _
                Array("Clm18", "Pagi", "ACTIVE_LATE", gedTextEdit, 0, 1, 60, 7, 1, 0, 0), _
                Array("Clm19", "OT Lev1", "OT1_BLK", gedTextEdit, 0, 1, 60, 7, 2, 0, 1), _
                Array("Clm20", "OT Hari", "OT1_BLK_HARI", gedTextEdit, 0, 1, 60, 7, 1, 0, 0), _
                Array("Clm21", "OT Lev2", "OT2_BLK", gedTextEdit, 0, 1, 60, 10, 2, 0, 1), _
                Array("Clm22", "OT Hari", "OT2_BLK_HARI", gedTextEdit, 0, 1, 60, 10, 1, 0, 0), _
                Array("Clm23", "OT Lev3", "OT3_BLK", gedTextEdit, 0, 1, 60, 7, 2, 0, 1), _
                Array("Clm24", "OT Hari", "OT3_BLK_HARI", gedTextEdit, 0, 1, 60, 7, 1, 0, 0))
                '--------0---------1--------2-----------3--------4--5--6---7--8--9--10
 Case Is = 1
     '=========================
     '--- MENU LOG_MAINTAIN ---
     '=========================
     BND1 = Array("EMPLOYE", "PERSONAL LOG VALUES")
     CLM1 = Array(Array("Clm0", "EmployeID", "KAR_ID", gedTextEdit, 0, 0, 120, 1, 0, 0, 0), _
                Array("Clm1", "Employe Name", "KAR_NM", gedTextEdit, 0, 0, 140, 1, 0, 0, 0), _
                Array("Clm2", "Terminal", "TerminalID", gedTextEdit, 0, 1, 180, 1, 0, 0, 0), _
                Array("Clm3", "FingerID", "FingerPrintID", gedTextEdit, 0, 1, 80, 0, 1, 0, 0), _
                Array("Clm4", "User Name", "UserName", gedTextEdit, 0, 1, 100, 0, 0, 0, 0), _
                Array("Clm5", "Key", "FunctionKey", gedLookupEdit, 0, 1, 80, 0, 1, 0, 0), _
                Array("Clm6", "DateTime", "DateTime", gedTextEdit, 0, 1, 120, 0, 2, 0, 2), _
                Array("Clm7", "Editing", "Edited", gedTextEdit, 0, 1, 120, 0, 2, 0, 2), _
                Array("Clm8", "Status", "FlagAbsence", gedTextEdit, 0, 1, 70, 1, 2, 0, 0))
                '--------0---------1--------2-----------3--------4--5--6---7--8--9--10
Case Is = 2
     '============================
     '--- MENU OVERTIME_LEVEL ----
     '============================
    BND2 = Array("PERSONAL ABSENSI ", "OVERTIME DETAIL")
    CLM2 = Array(Array("Clm0", "Date", "TGL_RUN", gedTextEdit, 0, 0, 65, 1, 0, 0, 0), _
                 Array("Clm1", "EmployeID ", "KAR_ID", gedTextEdit, 0, 0, 100, 1, 0, 0, 0), _
                 Array("Clm2", "Employe Name", "KAR_NM", gedTextEdit, 0, 0, 135, 1, 0, 0, 0), _
                 Array("Clm3", "Cabang", "CAB_ID", gedLookupEdit, 0, 0, 100, 1, 2, 0, 0), _
                 Array("Clm4", "Time In.", "DT_IN", gedTimeEdit, 0, 0, 110, 1, 2, 0, 2), _
                 Array("Clm5", "Time Out.", "DT_OUT", gedTextEdit, 0, 0, 110, 1, 2, 0, 2), _
                 Array("Clm6", "Group", "GRP_ID", gedLookupEdit, 0, 0, 70, 1, 1, 0, 0), _
                 Array("Clm7", "OT2 Note.", "TT_NOTE", gedMemoEdit, 0, 1, 60, 4, 0, 0, 0), _
                 Array("Clm8", "Lev2 OT.In", "OT2_IN", gedTimeEdit, 0, 1, 110, 4, 2, 0, 2), _
                 Array("Clm8", "Lev2 OT.Out", "OT2_OUT", gedTimeEdit, 0, 1, 110, 4, 2, 0, 2), _
                 Array("Clm10", "Lev2 OT ", "OT2", gedTextEdit, 0, 1, 55, 4, 2, 0, 1), _
                 Array("Clm11", "Hari ", "OT2_HARI", gedTextEdit, 0, 1, 40, 4, 1, 0, 7), _
                 Array("Clm12", "OT3 Note.", "TT3_NOTE", gedMemoEdit, 0, 1, 60, 4, 0, 0, 0), _
                 Array("Clm13", "Lev3 OT.In", "OT3_IN", gedTimeEdit, 0, 1, 110, 4, 2, 0, 2), _
                 Array("Clm14", "Lev3 OT.Out", "OT3_OUT", gedTimeEdit, 0, 1, 110, 4, 2, 0, 2), _
                 Array("Clm15", "Lev3 OT ", "OT3", gedTextEdit, 0, 1, 55, 4, 2, 0, 7), _
                 Array("Clm16", "Hari ", "OT3_HARI", gedTextEdit, 0, 1, 40, 4, 1, 0, 0))
                '--------0---------1--------2-----------3--------4--5--6---7--8--9--10
Case Is = 3
     '=============================
     '--- MENU PRESENSI_SESSION ---
     '=============================
    BND3 = Array("PRESENSI SESSION ", " ABSENSI VALUES")
    clm3 = Array(Array("Clm0", "Days", "Hari", gedTextEdit, 0, 0, 50, 1, 0, 0, 2), _
                  Array("Clm1", "Date", "TGL_RUN", gedTextEdit, 0, 0, 70, 1, 0, 0, 0), _
                  Array("Clm2", "Work Time", "WORK_TIME", gedTextEdit, 0, 0, 60, 4, 2, 0, 1), _
                  Array("Clm3", "Hari Aktif", "DAY_ACTIVE", gedTextEdit, 0, 0, 60, 4, 2, 0, 0), _
                  Array("Clm4", "Hari Libur", "HOLIDAY", gedTextEdit, 0, 0, 60, 4, 2, 0, 0), _
                  Array("Clm5", "Absen", "DAY_ABSENT", gedTextEdit, 0, 0, 50, 4, 2, 0, 0), _
                  Array("Clm6", "Ijin", "DAY_EXCP_ID", gedTextEdit, 0, 0, 50, 4, 2, 0, 0), _
                  Array("Clm7", "Late", "LATE", gedTextEdit, 0, 0, 60, 4, 2, 4, 1), _
                  Array("Clm8", "Early", "EARLY", gedTextEdit, 0, 0, 60, 4, 2, 0, 1), _
                  Array("Clm9", "OT Depan", "OT_DPN", gedTextEdit, 0, 0, 60, 10, 2, 0, 1), _
                  Array("Clm10", "OT Hari", "OT_DPN_HARI", gedTextEdit, 0, 0, 50, 10, 1, 0, 0), _
                  Array("Clm11", "Pagi", "ACTIVE_LATE", gedTextEdit, 0, 0, 60, 7, 1, 0, 0), _
                  Array("Clm12", "OT Lev1", "OT1_BLK", gedTextEdit, 0, 0, 60, 4, 2, 0, 1), _
                  Array("Clm13", "OT Hari", "OT1_BLK_HARI", gedTextEdit, 0, 0, 50, 4, 1, 0, 0), _
                  Array("Clm14", "OT Lev2", "OT2_BLK", gedTextEdit, 0, 0, 60, 10, 2, 0, 1), _
                  Array("Clm15", "OT Hari", "OT2_BLK_HARI", gedTextEdit, 0, 0, 50, 10, 1, 0, 0), _
                  Array("Clm16", "OT Lev3", "OT3_BLK", gedTextEdit, 0, 0, 60, 4, 2, 0, 1), _
                  Array("Clm17", "OT Hari", "OT3_BLK_HARI", gedTextEdit, 0, 0, 50, 4, 1, 0, 0), _
                  Array("Clm18", "Employe Name", "KAR_NM", gedTextEdit, 0, 1, 130, 1, 0, 0, 0), _
                  Array("Clm19", "Group", "GRP_ID", gedLookupEdit, 0, 1, 55, 0, 0, 0, 0), _
                  Array("Clm20", "EmployeID", "KAR_ID", gedTextEdit, 0, 1, 120, 1, 0, 0, 0), _
                  Array("Clm21", "Departmen", "DEP_ID", gedLookupEdit, 0, 1, 120, 1, 0, 0, 0), _
                  Array("Clm22", "Cabang ", "CAB_ID", gedLookupEdit, 0, 1, 120, 1, 0, 0, 0))
                  '--------0---------1--------2-----------3--------4--5--6---7--8--9--10

End Select
End Sub

Private Sub LookupClm()
'====================================
'-------------- ptr.nov -------------
'---- COLUMN LOOKUP LOG_MAINTAIN ----
'====================================
With dxDBGrid2.Columns.ColumnByName("Clm5").LookupColumn
            .LookupDataset.EnableControls
            .LookupDatasetType = dtADODataset
            .LookupDataset.ADODataset.ConnectionString = StrCon
            .LookupDataset.ADODataset.CursorLocation = clUseClient
            .LookupDataset.ADODataset.LockType = ltReadOnly         ' Check Permmission
            .LookupDataset.ADODataset.CommandType = cmdText
            .LookupDataset.ADODataset.CommandText = "SELECT FunctionKey,FunctionKeyNM FROM key_list" ' Like Join
            .LookupKeyField = "FunctionKey"
            .LookupResultField = "FunctionKeyNM"
            .ListColumns = "KEY NAME"
            .ListFieldName = "FunctionKeyNM"
            .ListWidth = 800
            .DisplaySize = 400
End With
dxDBGrid2.Dataset.Open
End Sub


Private Sub dxDBGrid3_LookupClm()
'======================================
'-------------- ptr.nov ---------------
'---- COLUMN LOOKUP OVERTIME_LEVEL ----
'======================================
With dxDBGrid3.Columns.ColumnByName("Clm6").LookupColumn
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

Private Sub dxDBGrid4_LookupClm()
'========================================
'-------------- ptr.nov -----------------
'---- COLUMN LOOKUP PRESENSI SESSION ----
'========================================
With dxDBGrid4.Columns.ColumnByName("Clm19").LookupColumn
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
With dxDBGrid4.Columns.ColumnByName("Clm21").LookupColumn
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
With dxDBGrid4.Columns.ColumnByName("Clm22").LookupColumn
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
dxDBGrid4.Dataset.Open
End Sub

Public Sub GridFill(GrdIndx As Integer)
'=====================================
'------------- ptr.nov ---------------
'----- PROCESS FIRST GRID LODING -----
'=====================================
Select Case GrdIndx
    Case Is = 0
            '===============================
            '-- Calculate PRESENSI In/Out --
            '===============================
            MenuGrid 0
            Qry0 = "absensi_grid('Absensi_inout" & _
                                  "=" & Format(DFilter(0).Value, "yyyy-mm-dd") & _
                                  "=" & Format(DFilter(1).Value, "yyyy-mm-dd") & _
                                  "=" & cmbFilter(0).Text & _
                                  "=" & cmbFilter(1).Text & _
                                  "=" & CmbAbsensi(1).Text & _
                                  "=" & CmbAbsensi(0).Text & _
                                  "=0" & _
                      "')"
            'MsgBox Qry0
            'Absensi_inout=2014-12-21=2015-01-23=255=DEMO.00002=2=HO'
            dxDBGrid1.Dataset.Refresh
            PrjSysGrid.GetGrid_Persensi dxDBGrid1, CLM0, True, True, BND0, True, Qry0, "KAR_ID", False
            dxDBGrid1.Dataset.Open
             With dxDBGrid1
                 .Columns.ColumnByName("Clm0").GroupIndex = 0
                 .GroupNodeColor = &HC0FFFF
                .M.FullExpand
            End With
   
    Case Is = 1
            '===========================
            '-- PRESENSI MAINTAIN_LOG --
            '===========================
            MenuGrid 1
            Qry1 = "absensi_grid('inout_maintain" & _
                                  "=" & Format(DT_LogFrm.Value, "yyyy-mm-dd") & _
                                  "=" & Format(DT_LogFrm.Value, "yyyy-mm-dd") & _
                                  "=" & Cmb_LogFrm(0).Text & _
                                  "=" & Cmb_LogFrm(1).Text & _
                                  "=" & Cmb_LogFrm(2).Text & _
                                  "=" & Cmb_LogFrm(3).Text & _
                                  "=" & Cmb_LogFrm(4).Text & _
                    "')"
            dxDBGrid2.Dataset.Refresh
            PrjSysGrid.GetGrid_Persensi dxDBGrid2, CLM1, True, True, BND1, True, Qry1, "KAR_ID", False
            'dxDBGrid2.Dataset.Open
            LookupClm
            
    Case Is = 2
            '=======================
            '-- OVERTIME LAVELING --
            '=======================
            MenuGrid 2
            Qry2 = "absensi_grid('ot_detail" & _
                                  "=" & Format(DFilter(2).Value, "yyyy-mm-dd") & _
                                  "=" & Format(DFilter(3).Value, "yyyy-mm-dd") & _
                                  "=" & cmbFilter(2).Text & _
                                  "=" & cmbFilter(3).Text & _
                                  "=" & CmbAbsensi(1).Text & _
                                  "=" & CmbAbsensi(0).Text & _
                                  "=0" & _
                    "')"
                   ' Qry2 = "'ot_detail=2015-02-01=2015-02-28=0=WAN.HO.000087=0=0'"
                    'MsgBox Qry2
            PrjSysGrid.GetGrid_Persensi dxDBGrid3, CLM2, True, True, BND2, True, Qry2, "KAR_ID", False
            dxDBGrid3_LookupClm
            'dxDBGrid3.Dataset.Open
             With dxDBGrid3
                 .Columns.ColumnByName("Clm1").GroupIndex = 0
                 .M.FullExpand
                 .GroupNodeColor = &HC0FFFF
                 .Columns.ColumnByName("Clm0").GroupIndex = 1
            End With
            
    Case Is = 3
            '=======================
            '-- PRESENSI SESSION  --
            '=======================
            MenuGrid 3
            Qry3 = "absensi_grid('presensi_session" & _
                                  "=" & Format(DFilter(4).Value, "yyyy-mm-dd") & _
                                  "=" & Format(DFilter(5).Value, "yyyy-mm-dd") & _
                                  "=" & cmbFilter(4).Text & _
                                  "=" & cmbFilter(5).Text & _
                                  "=" & CmbAbsensi(1).Text & _
                                  "=" & CmbAbsensi(0).Text & _
                                  "=0" & _
                    "')"
            dxDBGrid4.Dataset.Refresh
            PrjSysGrid.GetGrid_Persensi dxDBGrid4, clm3, True, True, BND3, True, Qry3, "KAR_ID", False
            'dxDBGrid4.Dataset.Open
            dxDBGrid4_LookupClm
            With dxDBGrid4
            .Columns.ColumnByName("Clm20").GroupIndex = 0
            .GroupNodeColor = &HC0FFFF
            .Columns.ColumnByName("Clm19").GroupIndex = 1
            
            
            '.Columns. ColumnByName("Clm9").SummaryFooterFormat = "#,###"
            '.Columns.ColumnByName("Clm9").SummaryFooterFormat = "#.###"
            '.Columns.ColumnByName("Clm9").SummaryField = "DAY_ACTIVE"
            '.Columns.ColumnByName("Clm9").SummaryFooterType = cstSum  ' Aktif
            '.Columns.ColumnByName("Clm10").SummaryFooterType = cstCount ' Libur
            '.Columns.ColumnByName("Clm11").SummaryFooterType = cstCount ' Libur
            '.Columns.ColumnByName("Clm12").SummaryFooterType = cstCount ' Ijin
            '.Columns.ColumnByName("Clm13").SummaryFooterType = cstSum ' Late
            '.Columns.ColumnByName("Clm14").SummaryFooterType = cstSum ' Early
            '.Columns.ColumnByName("Clm15").SummaryFooterType = cstSum ' OT Depan
            '.Columns.ColumnByName("Clm16").SummaryFooterType = cstSum ' OTDepan_Hari
            '.Columns.ColumnByName("Clm17").SummaryFooterFormat = "hh:mm:ss"
            '.Columns.ColumnByName("Clm17").SummaryFooterType = cstSum ' OT1
            '.Columns.ColumnByName("Clm17").su
            '.Columns.ColumnByName("Clm18").SummaryFooterFormat = "#.###"
            '.Columns.ColumnByName("Clm18").SummaryFooterType = cstSum ' OT1_Hari
            '.Columns.ColumnByName("Clm19").SummaryFooterType = cstSum ' OT2
            '.Columns.ColumnByName("Clm20").SummaryFooterFormat = "#.###"
            '.Columns.ColumnByName("Clm20").SummaryFooterType = cstSum ' OT2_Hari
            '.Columns.ColumnByName("Clm21").SummaryFooterType = cstSum ' OT3
            '.Columns.ColumnByName("Clm22").SummaryFooterFormat = "#.###"
            '.Columns.ColumnByName("Clm22").SummaryFooterType = cstSum ' OT3_Hari
             End With
    End Select
End Sub

'=====================================
'============= ptr.nov ===============
'=========== GET SETTING DB ==========
'=====================================
Private Sub init_Setting()
    PrjAbsensi.Get_Setting PERSENSI_STGL, F_TGL, F_WAKTU, F_TGL_WAKTU
    'MsgBox "PERSENSI_STGL= " & PERSENSI_STGL & _
           "; " & F_TGL & _
           "; " & F_WAKTU & _
           "; " & F_TGL_WAKTU
End Sub

Private Sub init_DT_CMB_filter()
'=====================================
'------------- ptr.nov ---------------
'------- DATE INITIALIZE Filter ------
'=====================================
Dim i As Integer
For i = 0 To 5
    DFilter(i).Value = Now
    DFilter(i).Format = xtpPickerShortDate
Next i
    prjSysID.Dept cmbFilter(0)
    prjSysID.Employe cmbFilter(1), CmbAbsensi(2).Text, CmbAbsensi(3).Text, 0
    prjSysID.Dept cmbFilter(2)
    prjSysID.Employe cmbFilter(3), CmbAbsensi(2).Text, CmbAbsensi(3).Text, 0
    prjSysID.Dept cmbFilter(4)
    prjSysID.Employe cmbFilter(5), CmbAbsensi(2).Text, CmbAbsensi(3).Text, 0
    '--- LOG FILTer ---
    DT_LogFrm.Value = Now
    DT_LogFrm.Format = xtpPickerShortDate
    prjSysID.Dept Cmb_LogFrm(0)
    prjSysID.TTGROUP Cmb_LogFrm(2)
    prjSysID.CABANG Cmb_LogFrm(3)
    prjSysID.FingerEmploye Cmb_LogFrm(4), 0
End Sub

Private Sub init_DT_CMB_Absensi()
'=====================================
'------------- ptr.nov ---------------
'------- DATE INITIALIZE Absensi -----
'=====================================
'prjSysID.Dept CmbAbsensi(0)
'prjSysID.Employe CmbAbsensi(1), "'" & CmbAbsensi(2).Text & "'", CmbAbsensi(3).Text, CmbAbsensi(0).Text
init_cmbDate
End Sub

Private Sub init_cmbDate()
'=====================================
'------------- ptr.nov ---------------
'----- DATE INITIALIZE Calculate -----
'=====================================
Tgl(1).Value = Now
Tgl(1).Format = xtpPickerShortDate
Tgl(0).Value = Now
Tgl(0).Format = xtpPickerShortDate
End Sub



Private Sub cmbFilter_Click(Index As Integer)
'============================================
'----------------- ptr.nov ------------------
'----- PROCESS FILTER DEP-EMPLOYE CHANGE ----
'============================================
Select Case Index
    Case 0: prjSysID.Employe cmbFilter(1), CmbAbsensi(2).Text, CmbAbsensi(3).Text, cmbFilter(0).Text
    Case 2: prjSysID.Employe cmbFilter(3), CmbAbsensi(2).Text, CmbAbsensi(3).Text, cmbFilter(2).Text
    Case 4: prjSysID.Employe cmbFilter(5), CmbAbsensi(2).Text, CmbAbsensi(3).Text, cmbFilter(4).Text
End Select
End Sub

Private Sub Cmb_LogFrm_Click(Index As Integer)
'========================================
'---------------- ptr.nov ---------------
'----- PROCESS FILTER MACHINE CHANGE ----
'========================================
If Index = 0 Then
    prjSysID.Employe Cmb_LogFrm(1), CmbAbsensi(2).Text, CmbAbsensi(3).Text, Cmb_LogFrm(0).Text
End If
If Index = 3 Then
    prjSysID.FingerEmploye Cmb_LogFrm(4), Cmb_LogFrm(3).Text
End If
End Sub



