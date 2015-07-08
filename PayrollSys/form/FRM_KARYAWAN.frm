VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form FRM_KARYAWAN 
   BackColor       =   &H00FFFFFF&
   Caption         =   "KARYAWAN"
   ClientHeight    =   10830
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
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10830
   ScaleWidth      =   20160
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   11175
      Left            =   0
      TabIndex        =   0
      Top             =   -360
      Width           =   20175
      _Version        =   851970
      _ExtentX        =   35586
      _ExtentY        =   19711
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
      Appearance      =   11
      Color           =   128
      PaintManager.BoldSelected=   -1  'True
      PaintManager.ShowTabs=   0   'False
      PaintManager.FixedTabWidth=   150
      PaintManager.MaxTabWidth=   150
      PaintManager.MinTabWidth=   150
      PaintManager.ControlMargin=   "1,0,0,0"
      ItemCount       =   2
      Item(0).Caption =   "DATA KARYAWAN"
      Item(0).ControlCount=   7
      Item(0).Control(0)=   "cmdList(0)"
      Item(0).Control(1)=   "dxDBGrid"
      Item(0).Control(2)=   "cmdList(1)"
      Item(0).Control(3)=   "cmdList(2)"
      Item(0).Control(4)=   "grpDetail"
      Item(0).Control(5)=   "GroupBox1"
      Item(0).Control(6)=   "Check1"
      Item(1).Caption =   "Departmen"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "PushButton1"
      Item(1).Control(1)=   "wBr"
      Begin VB.CheckBox Check1 
         BackColor       =   &H0080C0FF&
         Caption         =   "New Code"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   42
         Top             =   8040
         Width           =   1335
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   615
         Left            =   360
         TabIndex        =   31
         Top             =   600
         Width           =   10455
         _Version        =   851970
         _ExtentX        =   18441
         _ExtentY        =   1085
         _StockProps     =   79
         BackColor       =   8438015
         UseVisualStyle  =   -1  'True
         BorderStyle     =   2
         Begin XtremeSuiteControls.ComboBox cmbFilter 
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   32
            TabStop         =   0   'False
            ToolTipText     =   "Barang Type "
            Top             =   240
            Width           =   2055
            _Version        =   851970
            _ExtentX        =   3625
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Sorted          =   -1  'True
            Appearance      =   1
            UseVisualStyle  =   0   'False
            AutoComplete    =   -1  'True
            EnableMarkup    =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox cmbFilter 
            Height          =   315
            Index           =   2
            Left            =   3840
            TabIndex        =   33
            TabStop         =   0   'False
            ToolTipText     =   "Barang Satuan"
            Top             =   240
            Width           =   1695
            _Version        =   851970
            _ExtentX        =   2990
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Sorted          =   -1  'True
            Appearance      =   1
            UseVisualStyle  =   0   'False
            AutoComplete    =   -1  'True
            EnableMarkup    =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox cmbFilter 
            Height          =   315
            Index           =   1
            Left            =   2160
            TabIndex        =   34
            TabStop         =   0   'False
            ToolTipText     =   "Barang Kategori"
            Top             =   240
            Width           =   1695
            _Version        =   851970
            _ExtentX        =   2990
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Sorted          =   -1  'True
            Appearance      =   1
            UseVisualStyle  =   0   'False
            AutoComplete    =   -1  'True
            EnableMarkup    =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox cmbFilter 
            Height          =   315
            Index           =   3
            Left            =   6600
            TabIndex        =   35
            TabStop         =   0   'False
            ToolTipText     =   "Barang Type "
            Top             =   240
            Width           =   2055
            _Version        =   851970
            _ExtentX        =   3625
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Sorted          =   -1  'True
            Appearance      =   1
            UseVisualStyle  =   0   'False
            AutoComplete    =   -1  'True
            EnableMarkup    =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox cmbFilter 
            Height          =   315
            Index           =   4
            Left            =   8640
            TabIndex        =   36
            TabStop         =   0   'False
            ToolTipText     =   "Barang Type "
            Top             =   240
            Width           =   1695
            _Version        =   851970
            _ExtentX        =   2990
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Sorted          =   -1  'True
            Appearance      =   1
            UseVisualStyle  =   0   'False
            AutoComplete    =   -1  'True
            EnableMarkup    =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblVar 
            Height          =   255
            Index           =   0
            Left            =   3840
            TabIndex        =   41
            Top             =   0
            Width           =   1215
            _Version        =   851970
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Departemen"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Rounded MT Bold"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblVar 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   40
            Top             =   0
            Width           =   855
            _Version        =   851970
            _ExtentX        =   1508
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Cabang"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Rounded MT Bold"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblVar 
            Height          =   255
            Index           =   2
            Left            =   2160
            TabIndex        =   39
            Top             =   0
            Width           =   1095
            _Version        =   851970
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Golongan"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Rounded MT Bold"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblVar 
            Height          =   255
            Index           =   15
            Left            =   6600
            TabIndex        =   38
            Top             =   0
            Width           =   855
            _Version        =   851970
            _ExtentX        =   1508
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Jabatan"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Rounded MT Bold"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblVar 
            Height          =   255
            Index           =   16
            Left            =   8640
            TabIndex        =   37
            Top             =   0
            Width           =   855
            _Version        =   851970
            _ExtentX        =   1508
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Status"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Rounded MT Bold"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox grpDetail 
         Height          =   6735
         Left            =   14760
         TabIndex        =   7
         Top             =   1200
         Width           =   4935
         _Version        =   851970
         _ExtentX        =   8705
         _ExtentY        =   11880
         _StockProps     =   79
         Caption         =   "Info"
         ForeColor       =   0
         BackColor       =   6929919
         Appearance      =   2
         Begin RichTextLib.RichTextBox rchInfo 
            Height          =   735
            Left            =   1680
            TabIndex        =   26
            Top             =   3600
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   1296
            _Version        =   393217
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            Appearance      =   0
            TextRTF         =   $"FRM_KARYAWAN.frx":0000
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XtremeSuiteControls.FlatEdit txtInfo 
            Height          =   375
            Index           =   0
            Left            =   1680
            TabIndex        =   19
            Top             =   600
            Width           =   3135
            _Version        =   851970
            _ExtentX        =   5530
            _ExtentY        =   661
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Locked          =   -1  'True
            Appearance      =   1
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtInfo 
            Height          =   375
            Index           =   2
            Left            =   1680
            TabIndex        =   20
            Top             =   1320
            Width           =   3135
            _Version        =   851970
            _ExtentX        =   5530
            _ExtentY        =   661
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Locked          =   -1  'True
            Appearance      =   1
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtInfo 
            Height          =   375
            Index           =   3
            Left            =   1680
            TabIndex        =   21
            Top             =   1680
            Width           =   3135
            _Version        =   851970
            _ExtentX        =   5530
            _ExtentY        =   661
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Locked          =   -1  'True
            Appearance      =   1
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtInfo 
            Height          =   375
            Index           =   4
            Left            =   1680
            TabIndex        =   22
            Top             =   2280
            Width           =   3135
            _Version        =   851970
            _ExtentX        =   5530
            _ExtentY        =   661
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Locked          =   -1  'True
            Appearance      =   1
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtInfo 
            Height          =   375
            Index           =   1
            Left            =   1680
            TabIndex        =   23
            Top             =   960
            Width           =   3135
            _Version        =   851970
            _ExtentX        =   5530
            _ExtentY        =   661
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Locked          =   -1  'True
            Appearance      =   1
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtInfo 
            Height          =   375
            Index           =   5
            Left            =   1680
            TabIndex        =   24
            Top             =   2640
            Width           =   3135
            _Version        =   851970
            _ExtentX        =   5530
            _ExtentY        =   661
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Locked          =   -1  'True
            Appearance      =   1
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtInfo 
            Height          =   375
            Index           =   6
            Left            =   1680
            TabIndex        =   25
            Top             =   3000
            Width           =   3135
            _Version        =   851970
            _ExtentX        =   5530
            _ExtentY        =   661
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Locked          =   -1  'True
            Appearance      =   1
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtInfo 
            Height          =   375
            Index           =   7
            Left            =   1680
            TabIndex        =   27
            Top             =   4320
            Width           =   3135
            _Version        =   851970
            _ExtentX        =   5530
            _ExtentY        =   661
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Locked          =   -1  'True
            Appearance      =   1
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtInfo 
            Height          =   375
            Index           =   8
            Left            =   1680
            TabIndex        =   28
            Top             =   4680
            Width           =   3135
            _Version        =   851970
            _ExtentX        =   5530
            _ExtentY        =   661
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Locked          =   -1  'True
            Appearance      =   1
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtInfo 
            Height          =   375
            Index           =   9
            Left            =   1680
            TabIndex        =   29
            Top             =   5040
            Width           =   3135
            _Version        =   851970
            _ExtentX        =   5530
            _ExtentY        =   661
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Locked          =   -1  'True
            Appearance      =   1
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label lblVar 
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   4695
            _Version        =   851970
            _ExtentX        =   8281
            _ExtentY        =   450
            _StockProps     =   79
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblVar 
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   18
            Top             =   4680
            Width           =   1575
            _Version        =   851970
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "BANK / NOREK"
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblVar 
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   17
            Top             =   4320
            Width           =   1575
            _Version        =   851970
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Tgl. Lahir"
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblVar 
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   16
            Top             =   3600
            Width           =   1575
            _Version        =   851970
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Alamat"
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblVar 
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   15
            Top             =   3120
            Width           =   1575
            _Version        =   851970
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "NO.JAMSOS"
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblVar 
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   14
            Top             =   2760
            Width           =   1575
            _Version        =   851970
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "NO.NPWP"
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblVar 
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   13
            Top             =   2400
            Width           =   1575
            _Version        =   851970
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "NO.KTP"
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblVar 
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   12
            Top             =   1800
            Width           =   1575
            _Version        =   851970
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Status "
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblVar 
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   11
            Top             =   1440
            Width           =   1575
            _Version        =   851970
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Tgl. Masuk"
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblVar 
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   10
            Top             =   5040
            Width           =   1575
            _Version        =   851970
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Email Pribadi"
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblVar 
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   9
            Top             =   1080
            Width           =   1575
            _Version        =   851970
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Jabatan"
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblVar 
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   1575
            _Version        =   851970
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Departemen"
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   375
         Left            =   -68680
         TabIndex        =   6
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
         _Version        =   851970
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "PushButton1"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.WebBrowser wBr 
         Height          =   5775
         Left            =   -68680
         TabIndex        =   5
         Top             =   840
         Visible         =   0   'False
         Width           =   11535
         _Version        =   851970
         _ExtentX        =   20346
         _ExtentY        =   10186
         _StockProps     =   173
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton cmdList 
         Height          =   735
         Index           =   0
         Left            =   13560
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   480
         Width           =   735
         _Version        =   851970
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Hapus"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         TextImageRelation=   1
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid 
         Height          =   6615
         Left            =   480
         OleObjectBlob   =   "FRM_KARYAWAN.frx":007C
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1320
         Width           =   14175
      End
      Begin XtremeSuiteControls.PushButton cmdList 
         Height          =   735
         Index           =   1
         Left            =   12120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   480
         Width           =   735
         _Version        =   851970
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Add"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.PushButton cmdList 
         Height          =   735
         Index           =   2
         Left            =   12840
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   480
         Width           =   735
         _Version        =   851970
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Edit"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         ImageAlignment  =   6
         TextImageRelation=   1
         ImageGap        =   0
      End
   End
End
Attribute VB_Name = "FRM_KARYAWAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GroupSubject As TaskPanelGroup
Dim GroupTab As TaskPanelGroup
Dim GroupMsg As TaskPanelGroup
Dim Item As TaskPanelGroupItem
Dim CLM1, CLM2 As Variant
Dim BND1, BND2, BND3 As Variant
Dim Qry As String
Dim DateSort, NewRoId, RoIdDef, RoTglDef As String
Dim LoadWas As Boolean
Dim lR As Long
Public EditSts, isReady As Boolean
Dim defValId, defValNM As String
Dim defForm As String
Dim defRec As Integer
Dim strqry As String
Dim i As Integer
Dim x As VbMsgBoxResult

Private Sub Check1_Click()
'=============================
'--- FILTER Check New Code ---
'=============================
If Check1.Value = 1 Then
    GridFill 1, 1
End If
End Sub

Private Sub cmbFilter_Click(Index As Integer)
'=========================
'---COMBO FILTER CLICK ---
'=========================
If Index = 0 Or Index = 1 Or Index = 2 Or Index = 3 Or Index = 4 Then
    MenuGrid 1
    GridFill 1, 0
End If
End Sub

Private Sub cmdList_Click(Index As Integer)
Select Case Index
Case 0
    '============
    '---Delete---
    '============
    x = MsgBox("Anda yakin akan menghapus [EmpID=" & prjSysID.GetGridDefValue(dxDBGrid, "KAR_ID") & _
                ", EMP Name=" & prjSysID.GetGridDefValue(dxDBGrid, "KAR_NM") & _
                ", Dept Name=" & prjSysID.GetGridDefValue(dxDBGrid, "DEP_ID") & "]", vbYesNo, "EMPLOYE CONFIRM")
    
    If x = vbYes Then
        MsgBox "Contact Administrator for Deleted, Please Prepare for Edit or Stt Resign !", , "EMPLOYE CONFIRM"
    End If
    
Case 1
    '=========
    '---Add---
    '=========
    FRM_KARYAWAN_ADD.Show
    
Case 2
    '==========
    '---Edit---
    '==========
    If cmdList(2).Caption = "Edit" Then
        For i = 1 To dxDBGrid.Columns.Count - 1
            dxDBGrid.Columns(i).DisableEditor = False
            dxDBGrid.Columns(i).Color = &HFFFFFF
            cmdList(2).Caption = "Accept"
        Next i
    Else
        dxDBGrid.Dataset.Refresh
        For i = 1 To dxDBGrid.Columns.Count - 1
            dxDBGrid.Columns(i).DisableEditor = True
            dxDBGrid.Columns(i).Color = &HCBFAFE
            cmdList(2).Caption = "Edit"
        Next i

    End If
    'EDIT_KARYAWAN.Show
    'EDIT_KARYAWAN.EditParm Me, prjSysID.GetGridDefValue(dxDBGrid, "KAR_ID"), False
End Select

End Sub

Private Sub dxDBGrid_OnClick()
On Error Resume Next
'=======================
'--- DETAIL KARYAWAN ---
'=======================
With dxDBGrid.Dataset
    If .RecNo <> 0 Then
        '===departemen===
        txtInfo(0).Text = prjSysID.GetGridDefValue(dxDBGrid, "DEP_ID")
        '===jabatan===
        txtInfo(1).Text = prjSysID.GetGridDefValue(dxDBGrid, "JAB_ID")
        '===tgl MASUK===
        txtInfo(2).Text = Format(prjSysID.GetGridDefValue(dxDBGrid, "KAR_TGLM"), "DD, MMMM YYYY")
        '===STATUS KERJA===
         txtInfo(3).Text = prjSysID.GetGridDefValue(dxDBGrid, "KAR_STS")
        '===ktp===
        txtInfo(4).Text = prjSysID.GetGridDefValue(dxDBGrid, "KAR_KTP")
        '===NPWP===
         txtInfo(5).Text = prjSysID.GetGridDefValue(dxDBGrid, "NPWP")
        '===KPJ===
        txtInfo(6).Text = prjSysID.GetGridDefValue(dxDBGrid, "NO_JAMSOS")
        '===alamat===
        rchInfo.Text = prjSysID.GetGridDefValue(dxDBGrid, "KAR_ALMT") & " " & prjSysID.GetGridDefValue(dxDBGrid, "KAR_KOTA")
        '===tgl lahir===
        txtInfo(7).Text = Format(prjSysID.GetGridDefValue(dxDBGrid, "KAR_TGL_LAHIR"), "DD, MMMM YYYY")
        '===BANK NOREK===
        txtInfo(8).Text = prjSysID.GetGridDefValue(dxDBGrid, "BANK") & " / " & prjSysID.GetGridDefValue(dxDBGrid, "NO_REK")
        '===email===
         txtInfo(9).Text = prjSysID.GetGridDefValue(dxDBGrid, "KAR_MAILK")
        '===Label Name===
        lblVar(8).Caption = prjSysID.GetGridDefValue(dxDBGrid, "KAR_NM")
    End If
End With
End Sub

Private Sub Form_Load()
On Error GoTo ErrorLabel

LOADING.Show
LOADING.SetParm Me, 25
'InitExample
'CreateTaskPanel

'LOADING.SetParm Me, 50
LOADING.SetParm Me, 75
prjSysID.CABANG cmbFilter(0)
prjSysID.TTGROUP cmbFilter(1)
prjSysID.Dept cmbFilter(2)
prjSysID.JABATAN cmbFilter(3)
prjSysID.KAR_STT cmbFilter(4)
GridFill 1, 0
Me.Icon = LoadPicture(ImagePath("FRM_KARYAWAN"))

LOADING.SetParm Me, 100
Exit Sub
ErrorLabel:
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error", vbCritical, "LG Error"
    LOADING.SetParm Me, 100

End Sub


Private Sub MenuGrid(TabSel As Integer)
Select Case TabSel
Case Is = 1 'array untuk TABEL KARYAWAN
    BND1 = Array("Employe Header ", "Employe Detail")
    CLM1 = Array(Array("Clm0", "ID Karyawan", "KAR_ID", gedTextEdit, 0, 0, 110, 1, 1, 0, 0), Array("Clm1", "Nama", "KAR_NM", gedTextEdit, 0, 0, 130, 1, 0, 0, 0), _
                    Array("Clm2", "Departemen", "DEP_ID", gedLookupEdit, 0, 1, 110, 1, 1, 0, 0), _
                    Array("Clm3", "Golongan", "GRP_ID", gedLookupEdit, 0, 1, 110, 1, 1, 0, 0), Array("Clm4", "Jabatan", "JAB_ID", gedLookupEdit, 0, 1, 80, 1, 1, 0, 0), _
                    Array("Clm5", "STATUS", "KAR_STS", gedLookupEdit, 0, 1, 80, 1, 1, 0, 0), _
                    Array("Clm6", "TELP", "KAR_TLP", gedTextEdit, 0, 1, 110, 11, 1, 1, 0), Array("Clm7", "MOBILE", "KAR_HP", gedTextEdit, 0, 1, 110, 1, 1, 0, 0), _
                    Array("Clm8", "ALAMAT", "KAR_ALMT", gedMemoEdit, 0, 1, 150, 1, 1, 0, 0), _
                    Array("Clm9", "Kota", "KAR_KOTA", gedTextEdit, 0, 1, 80, 1, 1, 0, 0), _
                    Array("Clm10", "TGL LAHIR", "KAR_TGL_LAHIR", gedDateEdit, 0, 1, 80, 1, 2, 0, 0), _
                    Array("Clm11", "EMAIL", "KAR_MAILK", gedTextEdit, 0, 1, 100, 1, 1, 0, 0), _
                    Array("Clm12", "BANK", "BANK", gedTextEdit, 0, 1, 80, 1, 0, 1, 0), Array("Clm13", "NOREK", "NO_REK", gedTextEdit, 0, 1, 80, 1, 1, 0, 0), _
                    Array("Clm14", "TGL MASUK", "KAR_TGLM", gedDateEdit, 0, 1, 80, 1, 2, 0, 0), _
                    Array("Clm15", "KTP", "KAR_KTP", gedTextEdit, 0, 1, 80, 1, 1, 0, 0), _
                    Array("Clm16", "NO_NPWP", "NPWP", gedTextEdit, 0, 1, 80, 1, 1, 0, 0), _
                    Array("Clm17", "NO_JAMSOS", "NO_JAMSOS", gedTextEdit, 0, 1, 80, 1, 1, 0, 0), _
                    Array("Clm18", "JAMSOS", "JAMSOS", gedLookupEdit, 0, 1, 80, 1, 1, 0, 0), _
                    Array("Clm19", "PTKP_ID", "PTKP_NM", gedLookupEdit, 0, 1, 80, 1, 2, 0, 0), _
                    Array("Clm20", "STATUS", "STT_ID", gedLookupEdit, 0, 1, 80, 1, 1, 0, 0), _
                    Array("Clm21", "JML ANAK", "JML_ANAK", gedSpinEdit, 0, 1, 80, 1, 2, 0, 0), _
                    Array("Clm22", "NO ASURANSI", "NO_ASR", gedTextEdit, 0, 1, 85, 1, 1, 0, 0), _
                    Array("Clm23", "ASURANSI", "ASR_ID", gedLookupEdit, 0, 1, 80, 1, 2, 0, 0), _
                    Array("Clm24", "STT OT DPN", "STT_OT_DPN", gedLookupEdit, 0, 1, 80, 1, 2, 0, 0), _
                    Array("Clm25", "PENDIDIKAN", "PEN_ID", gedLookupEdit, 0, 1, 100, 1, 2, 0, 0), _
                    Array("Clm26", "AGAMA", "AGAMA_ID", gedLookupEdit, 0, 1, 100, 1, 2, 0, 0), _
                    Array("Clm27", "GENDER", "KAR_JK", gedLookupEdit, 0, 1, 80, 1, 2, 0, 0), _
                    Array("Clm28", "NOTE", "KAR_KET", gedMemoEdit, 0, 1, 100, 1, 1, 0, 0))
                    '--------0---------1--------2-----------3--------4--5--6---7--8--9--10--
End Select
End Sub

Public Sub GridFill(GrdIndx As Integer, KarStt As Integer)
Select Case GrdIndx
    Case Is = 1 'ISI grid item dengan tabel karyawan
                MenuGrid 1
                If KarStt = 1 Then
                    For i = 1 To 4
                        cmbFilter(i).ListIndex = 0
                    Next i
                End If
                strqry = "crud_karyawan(1,'" & cmbFilter(0).Text & _
                                        "','" & cmbFilter(1).Text & _
                                        "','" & cmbFilter(2).Text & _
                                        "','" & cmbFilter(3).Text & _
                                        "','" & cmbFilter(4).Text & "','" & KarStt & "')"
                                        'MsgBox cmbFilter(1).Text
                PrjSysGrid.GetGrid_Persensi dxDBGrid, CLM1, False, False, BND1, True, strqry, "KAR_ID", False
                'dxDBGrid.Dataset.Open
                LookupClm
    End Select
End Sub
Private Sub LookupClm()
With dxDBGrid.Columns.ColumnByName("Clm2").LookupColumn
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
            .ListWidth = 200
            .DisplaySize = 400
End With

With dxDBGrid.Columns.ColumnByName("Clm3").LookupColumn
'== GROUP ---
            .LookupDataset.EnableControls
            '.LookupDataset.Close
            '.Dataset.Refresh
            .LookupDatasetType = dtADODataset
            .LookupDataset.ADODataset.ConnectionString = StrCon
            .LookupDataset.ADODataset.CursorLocation = clUseClient
            .LookupDataset.ADODataset.LockType = ltReadOnly         ' Check Permmission
            .LookupDataset.ADODataset.CommandType = cmdText
            .LookupDataset.ADODataset.CommandText = "SELECT TT_GRP_ID ,TT_GRP_NM FROM timetable_grp " ' Like Join
            '.LookupDataset.Open
            .LookupKeyField = "TT_GRP_ID"
            .LookupResultField = "TT_GRP_NM"
            .ListColumns = "GROUP TIME TABLE"
            .ListFieldName = "TT_GRP_NM"
            .ListWidth = 100
            .DisplaySize = 400
            
End With
With dxDBGrid.Columns.ColumnByName("Clm4").LookupColumn
'== JAbatan ---
            .LookupDataset.EnableControls
            '.LookupDataset.Close
            '.Dataset.Refresh
            .LookupDatasetType = dtADODataset
            .LookupDataset.ADODataset.ConnectionString = StrCon
            .LookupDataset.ADODataset.CursorLocation = clUseClient
            .LookupDataset.ADODataset.LockType = ltReadOnly         ' Check Permmission
            .LookupDataset.ADODataset.CommandType = cmdText
            .LookupDataset.ADODataset.CommandText = "SELECT JAB_ID ,JAB_NM FROM jabatan " ' Like Join
            '.LookupDataset.Open
            .LookupKeyField = "JAB_ID"
            .LookupResultField = "JAB_NM"
            .ListColumns = "JAB_NM TABLE"
            .ListFieldName = "JAB_NM"
            .ListWidth = 100
            .DisplaySize = 400
            
End With
With dxDBGrid.Columns.ColumnByName("Clm5").LookupColumn
'== Status ---
            .LookupDataset.EnableControls
            '.LookupDataset.Close
            '.Dataset.Refresh
            .LookupDatasetType = dtADODataset
            .LookupDataset.ADODataset.ConnectionString = StrCon
            .LookupDataset.ADODataset.CursorLocation = clUseClient
            .LookupDataset.ADODataset.LockType = ltReadOnly         ' Check Permmission
            .LookupDataset.ADODataset.CommandType = cmdText
            .LookupDataset.ADODataset.CommandText = "SELECT KAR_STS_ID ,KAR_STS_NM FROM kar_stt " ' Like Join
            '.LookupDataset.Open
            .LookupKeyField = "KAR_STS_ID"
            .LookupResultField = "KAR_STS_NM"
            .ListColumns = "STATUS KARYAWAN"
            .ListFieldName = "KAR_STS_NM"
            .ListWidth = 100
            .DisplaySize = 400
            
End With
With dxDBGrid.Columns.ColumnByName("Clm18").LookupColumn
'== JAMSOS ---
            .LookupDataset.EnableControls
            '.LookupDataset.Close
            '.Dataset.Refresh
            .LookupDatasetType = dtADODataset
            .LookupDataset.ADODataset.ConnectionString = StrCon
            .LookupDataset.ADODataset.CursorLocation = clUseClient
            .LookupDataset.ADODataset.LockType = ltReadOnly         ' Check Permmission
            .LookupDataset.ADODataset.CommandType = cmdText
            .LookupDataset.ADODataset.CommandText = "SELECT SOS_ID FROM payroll_jamsos_formula "
            '.LookupDataset.Open
            .LookupKeyField = "SOS_ID"
            .LookupResultField = "SOS_ID"
            .ListColumns = "JAMSOS ID"
            .ListFieldName = "SOS_ID"
            .ListWidth = 100
            .DisplaySize = 400
            
End With
With dxDBGrid.Columns.ColumnByName("Clm19").LookupColumn
'== PTKP ---
            .LookupDataset.EnableControls
            '.LookupDataset.Close
            '.Dataset.Refresh
            .LookupDatasetType = dtADODataset
            .LookupDataset.ADODataset.ConnectionString = StrCon
            .LookupDataset.ADODataset.CursorLocation = clUseClient
            .LookupDataset.ADODataset.LockType = ltReadOnly         ' Check Permmission
            .LookupDataset.ADODataset.CommandType = cmdText
            .LookupDataset.ADODataset.CommandText = "SELECT PTKP_NM FROM payroll_ptkp_formula " ' Like Join
            '.LookupDataset.Open
            .LookupKeyField = "PTKP_NM"
            .LookupResultField = "PTKP_NM"
            .ListColumns = "PTKP NAME"
            .ListFieldName = "PTKP_NM"
            .ListWidth = 100
            .DisplaySize = 400
            
End With
With dxDBGrid.Columns.ColumnByName("Clm20").LookupColumn
'== STATUS NIKAH ---
            .LookupDataset.EnableControls
            '.LookupDataset.Close
            '.Dataset.Refresh
            .LookupDatasetType = dtADODataset
            .LookupDataset.ADODataset.ConnectionString = StrCon
            .LookupDataset.ADODataset.CursorLocation = clUseClient
            .LookupDataset.ADODataset.LockType = ltReadOnly         ' Check Permmission
            .LookupDataset.ADODataset.CommandType = cmdText
            .LookupDataset.ADODataset.CommandText = "SELECT STT_ID ,STT_NM FROM payroll_ptkp_stt " ' Like Join
            '.LookupDataset.Open
            .LookupKeyField = "STT_ID"
            .LookupResultField = "STT_NM"
            .ListColumns = "STATUS PERNIKAHAN"
            .ListFieldName = "STT_NM"
            .ListWidth = 200
            .DisplaySize = 400
            
End With
With dxDBGrid.Columns.ColumnByName("Clm23").LookupColumn
'== ASURANSI ---
            .LookupDataset.EnableControls
            '.LookupDataset.Close
            '.Dataset.Refresh
            .LookupDatasetType = dtADODataset
            .LookupDataset.ADODataset.ConnectionString = StrCon
            .LookupDataset.ADODataset.CursorLocation = clUseClient
            .LookupDataset.ADODataset.LockType = ltReadOnly         ' Check Permmission
            .LookupDataset.ADODataset.CommandType = cmdText
            .LookupDataset.ADODataset.CommandText = "SELECT ASR_ID ,ASR_NM FROM payroll_asuransi_formula " ' Like Join
            '.LookupDataset.Open
            .LookupKeyField = "ASR_ID"
            .LookupResultField = "ASR_NM"
            .ListColumns = "ASURANSI"
            .ListFieldName = "ASR_NM"
            .ListWidth = 200
            .DisplaySize = 400
            
End With
With dxDBGrid.Columns.ColumnByName("Clm24").LookupColumn
'== STT OT DPN ---
            .LookupDataset.EnableControls
            '.LookupDataset.Close
            '.Dataset.Refresh
            .LookupDatasetType = dtADODataset
            .LookupDataset.ADODataset.ConnectionString = StrCon
            .LookupDataset.ADODataset.CursorLocation = clUseClient
            .LookupDataset.ADODataset.LockType = ltReadOnly         ' Check Permmission
            .LookupDataset.ADODataset.CommandType = cmdText
            .LookupDataset.ADODataset.CommandText = "SELECT STT_OT_DPN ,STT_OT_DPN_NM FROM ot_dpn_stt " ' Like Join
            '.LookupDataset.Open
            .LookupKeyField = "STT_OT_DPN"
            .LookupResultField = "STT_OT_DPN_NM"
            .ListColumns = "Status Overtime Depan"
            .ListFieldName = "STT_OT_DPN_NM"
            .ListWidth = 200
            .DisplaySize = 400
            
End With
With dxDBGrid.Columns.ColumnByName("Clm25").LookupColumn
'== PENDIDIKAN ---
            .LookupDataset.EnableControls
            '.LookupDataset.Close
            '.Dataset.Refresh
            .LookupDatasetType = dtADODataset
            .LookupDataset.ADODataset.ConnectionString = StrCon
            .LookupDataset.ADODataset.CursorLocation = clUseClient
            .LookupDataset.ADODataset.LockType = ltReadOnly         ' Check Permmission
            .LookupDataset.ADODataset.CommandType = cmdText
            .LookupDataset.ADODataset.CommandText = "SELECT PEN_ID ,PEN_NM FROM pendidikan " ' Like Join
            '.LookupDataset.Open
            .LookupKeyField = "PEN_ID"
            .LookupResultField = "PEN_NM"
            .ListColumns = "PENDIDIKAN"
            .ListFieldName = "PEN_NM"
            .ListWidth = 200
            .DisplaySize = 400
            
End With
With dxDBGrid.Columns.ColumnByName("Clm26").LookupColumn
'== AGAMA ---
            .LookupDataset.EnableControls
            '.LookupDataset.Close
            '.Dataset.Refresh
            .LookupDatasetType = dtADODataset
            .LookupDataset.ADODataset.ConnectionString = StrCon
            .LookupDataset.ADODataset.CursorLocation = clUseClient
            .LookupDataset.ADODataset.LockType = ltReadOnly         ' Check Permmission
            .LookupDataset.ADODataset.CommandType = cmdText
            .LookupDataset.ADODataset.CommandText = "SELECT AGAMA_ID ,AGAMA_NM FROM Agama " ' Like Join
            '.LookupDataset.Open
            .LookupKeyField = "AGAMA_ID"
            .LookupResultField = "AGAMA_NM"
            .ListColumns = "AGAMA"
            .ListFieldName = "AGAMA_NM"
            .ListWidth = 150
            .DisplaySize = 400
            
End With
With dxDBGrid.Columns.ColumnByName("Clm27").LookupColumn
'== GENDER ---
            .LookupDataset.EnableControls
            '.LookupDataset.Close
            '.Dataset.Refresh
            .LookupDatasetType = dtADODataset
            .LookupDataset.ADODataset.ConnectionString = StrCon
            .LookupDataset.ADODataset.CursorLocation = clUseClient
            .LookupDataset.ADODataset.LockType = ltReadOnly         ' Check Permmission
            .LookupDataset.ADODataset.CommandType = cmdText
            .LookupDataset.ADODataset.CommandText = "SELECT JK_ID,JK_NM FROM kar_jk " ' Like Join
            '.LookupDataset.Open
            .LookupKeyField = "JK_ID"
            .LookupResultField = "JK_NM"
            .ListColumns = "GENDER"
            .ListFieldName = "JK_NM"
            .ListWidth = 150
            .DisplaySize = 400
            
End With

dxDBGrid.Dataset.Open
End Sub




Private Sub TabControl1_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
 'refGrid
End Sub


Private Sub dxDBGrid_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
Dim s As Integer
Dim a As Integer
Dim prgVal As Integer
'On Error Resume Next

If LoadWas Then
    'MsgBox dxDBGridKar.Dataset.RecordCount
   ' prgVal = 100 / dxDBGridKar.Dataset.RecordCount
    LOADING.SetParm Me, (Node.RecNo + 1) * prgVal
      '  If Node.RecNo + 1 = dxDBGridKar.Dataset.RecordCount Then
      '      LoadWas = False
      '  End If
End If

'properti warna baris
s = Node.RecNo 'Node.Values(dxDBGridKar.Columns.ColumnByFieldName("NO").Index)
        If s Mod 2 = 0 Then
            Color = RGB(255, 255, 255)
        Else
            Color = &HE0E0E0
        End If
        
       
'properti warna barang akrif/tidak aktif
a = Node.Values(dxDBGrid.Columns.ColumnByFieldName("KAR_STS").Index)
    If a = 0 Then
           FontColor = RGB(255, 0, 0)
        Else
           FontColor = RGB(0, 0, 0)
        End If

'set label jumlah item
'txtPropBrg(0) = dxDBGridKar.Dataset.RecordCount
End Sub
Public Sub refGrid()
Dim i%

'MenuGrid 1
'GridFill 1

For i = 0 To 9
    txtInfo(i).Text = ""
Next i

rchInfo.Text = ""
lblVar(8).Caption = ""

dxDBGrid.Ex.FocusedNumber = defRec
End Sub


Sub InitExample()
 dxDBGrid.Event = 1 'EGOnCustomDrawCell
 dxDBGrid.EventEnabled = True
 'dxDBGridROList.Event = 1 'EGOnCustomDrawCell
 'dxDBGridROList.EventEnabled = True
End Sub

Private Sub PushButton1_Click()
Dim html As String
html = "<HTML><head><style> body {font-family:MS Sans Serif; font-size:13;}table {font-family:MS Sans Serif; font-size:13;}</style></head><body topmargin=0 leftmargin=0 bottommargin=0 rightmargin=0 bgColor=#ffffff></body><div id='debate_1_2243882'></div>" & _
"<script>" & _
  "(function () {" & _
    "var opst = document.createElement('script');" & _
    "var os_host = document.location.protocol == 'https:' ? 'https:' : 'http:';" & _
    "opst.type = 'text/javascript';" & _
    "opst.async = true;" & _
    "opst.src = os_host + '//' + 'www.opinionstage.com/polls/2243882/embed.js';" & _
    "(document.getElementsByTagName('head')[0] ||" & _
      "document.getElementsByTagName('body')[0]).appendChild(opst);" & _
  "}());" & _
"</script></html>"
'MsgBox html
'wBr.InnerHTML = "<HTML>Bego<html>" 'html
wBr.Navigate ("https://apps.facebook.com/my-surveys/wykdwq")
'Set wBr.LocationURL  "http://farindra.com/pool.html"
End Sub

Private Sub txtNm_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
'    refGrid
End If
End Sub
