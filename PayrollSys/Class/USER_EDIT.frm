VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Begin VB.Form USER_EDIT 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User - "
   ClientHeight    =   11145
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "USER_EDIT.frx":0000
   ScaleHeight     =   11145
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   10440
      TabIndex        =   56
      Text            =   "Combo1"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fraNew 
      BackColor       =   &H00FFFFFF&
      Height          =   5295
      Index           =   1
      Left            =   120
      TabIndex        =   21
      Top             =   5640
      Visible         =   0   'False
      Width           =   11535
      Begin VB.CommandButton Command1 
         Caption         =   "Browse..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   7680
         TabIndex        =   72
         Top             =   4440
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   9240
         TabIndex        =   54
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   10320
         TabIndex        =   53
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox txtKarPrp 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Index           =   7
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   52
         Text            =   "000000"
         Top             =   480
         Width           =   3615
      End
      Begin VB.TextBox txtKarPrp 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   1920
         TabIndex        =   51
         Top             =   4800
         Width           =   3615
      End
      Begin VB.TextBox txtKarPrp 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   6
         Left            =   7680
         TabIndex        =   50
         Top             =   2400
         Width           =   3615
      End
      Begin VB.TextBox txtKarPrp 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   7680
         TabIndex        =   45
         Top             =   1920
         Width           =   3615
      End
      Begin VB.TextBox txtKarPrp 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   5
         Left            =   1920
         TabIndex        =   36
         Top             =   4320
         Width           =   3615
      End
      Begin VB.TextBox txtKarPrp 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   1920
         TabIndex        =   32
         Top             =   2880
         Width           =   3615
      End
      Begin VB.TextBox txtKarPrp 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   1920
         TabIndex        =   30
         Top             =   1800
         Width           =   3615
      End
      Begin VB.TextBox txtKarPrp 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   1920
         TabIndex        =   22
         Top             =   1320
         Width           =   3615
      End
      Begin XtremeSuiteControls.ComboBox cmbKar 
         Height          =   315
         Index           =   0
         Left            =   1920
         TabIndex        =   27
         Top             =   3360
         Width           =   3615
         _Version        =   851970
         _ExtentX        =   6376
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.ComboBox cmbKar 
         Height          =   315
         Index           =   1
         Left            =   1920
         TabIndex        =   28
         Top             =   3840
         Width           =   3615
         _Version        =   851970
         _ExtentX        =   6376
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.ComboBox cmbKar 
         Height          =   315
         Index           =   2
         Left            =   7680
         TabIndex        =   39
         Top             =   960
         Width           =   3615
         _Version        =   851970
         _ExtentX        =   6376
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.ComboBox cmbKar 
         Height          =   315
         Index           =   3
         Left            =   7680
         TabIndex        =   40
         Top             =   1440
         Width           =   3615
         _Version        =   851970
         _ExtentX        =   6376
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.DateTimePicker datePickKar 
         Height          =   375
         Index           =   2
         Left            =   7680
         TabIndex        =   47
         Top             =   3960
         Width           =   3615
         _Version        =   851970
         _ExtentX        =   6376
         _ExtentY        =   661
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         MinDate         =   36890
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   3
         CurrentDate     =   36890
      End
      Begin XtremeSuiteControls.DateTimePicker datePickKar 
         Height          =   375
         Index           =   1
         Left            =   7680
         TabIndex        =   38
         Top             =   3360
         Width           =   3615
         _Version        =   851970
         _ExtentX        =   6376
         _ExtentY        =   661
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MinDate         =   36890
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   3
         CurrentDate     =   36890
      End
      Begin XtremeSuiteControls.ComboBox cmbKar 
         Height          =   315
         Index           =   4
         Left            =   7680
         TabIndex        =   55
         Top             =   2880
         Width           =   3615
         _Version        =   851970
         _ExtentX        =   6376
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.DateTimePicker datePickKar 
         Height          =   375
         Index           =   0
         Left            =   1920
         TabIndex        =   37
         Top             =   2280
         Width           =   3615
         _Version        =   851970
         _ExtentX        =   6376
         _ExtentY        =   661
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   1
      End
      Begin XtremeSuiteControls.ComboBox cmbKar 
         Height          =   315
         Index           =   5
         Left            =   1920
         TabIndex        =   67
         Top             =   360
         Width           =   3615
         _Version        =   851970
         _ExtentX        =   6376
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         BackColor       =   12648447
      End
      Begin XtremeSuiteControls.ComboBox cmbKar 
         Height          =   315
         Index           =   6
         Left            =   1920
         TabIndex        =   69
         Top             =   840
         Width           =   3615
         _Version        =   851970
         _ExtentX        =   6376
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         BackColor       =   12648447
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Foto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   29
         Left            =   5880
         TabIndex        =   71
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cabang"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   28
         Left            =   120
         TabIndex        =   70
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Perusahaan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   68
         Top             =   360
         Width           =   1815
      End
      Begin VB.Line Line1 
         X1              =   5760
         X2              =   5760
         Y1              =   360
         Y2              =   5040
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail Karyawan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   24
         Left            =   5880
         TabIndex        =   49
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Keluar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   23
         Left            =   5880
         TabIndex        =   48
         Top             =   3960
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Masuk"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   22
         Left            =   5880
         TabIndex        =   46
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail Pribadi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   5880
         TabIndex        =   44
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Status Karyawan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   5880
         TabIndex        =   43
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "HP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   42
         Top             =   4800
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tlp. Rumah"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   41
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Status Nikah"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   35
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Agama"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   34
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   33
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Lahir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   31
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Identitas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   29
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Departemen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   5880
         TabIndex        =   26
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Jabatan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   5880
         TabIndex        =   25
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ID Karyawan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   5880
         TabIndex        =   24
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Karyawan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   1695
      End
   End
   Begin VB.Frame fraUsr 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   2535
      Index           =   1
      Left            =   5400
      TabIndex        =   57
      Top             =   240
      Width           =   4935
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   5
         Left            =   2400
         MaxLength       =   12
         PasswordChar    =   "*"
         TabIndex        =   63
         Text            =   "xxxxx"
         ToolTipText     =   "Maksimal nama user 12 Karakte"
         Top             =   120
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   2400
         MaxLength       =   12
         PasswordChar    =   "*"
         TabIndex        =   62
         Text            =   "xxxxx"
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   2400
         MaxLength       =   12
         PasswordChar    =   "*"
         TabIndex        =   61
         Text            =   "xxxxx"
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CheckBox ChkStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Show Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   60
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   2520
         TabIndex        =   59
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   3600
         TabIndex        =   58
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Passwor Lama"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   27
         Left            =   240
         TabIndex        =   66
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password Baru"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   26
         Left            =   240
         TabIndex        =   65
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password Baru(ulang)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   25
         Left            =   240
         TabIndex        =   64
         Top             =   1080
         Width           =   2055
      End
   End
   Begin VB.Frame fraNew 
      BackColor       =   &H00FFFFFF&
      Height          =   5175
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton Command1 
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1920
         TabIndex        =   0
         Top             =   2160
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox txtKar 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   2
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1800
         Width           =   2775
      End
      Begin VB.TextBox txtKar 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   1
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox txtKar 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   0
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   840
         Width           =   2775
      End
      Begin VB.Frame fraUsr 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   2415
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   2640
         Width           =   4575
         Begin VB.CommandButton Command1 
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   3360
            TabIndex        =   7
            Top             =   1920
            Width           =   975
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Save"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   2040
            TabIndex        =   6
            Top             =   1920
            Width           =   975
         End
         Begin VB.CheckBox ChkStatus 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            Caption         =   "Aktif"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   2040
            TabIndex        =   5
            Top             =   1560
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   2040
            MaxLength       =   12
            PasswordChar    =   "*"
            TabIndex        =   4
            Text            =   "USR_PASS"
            Top             =   1080
            Width           =   2175
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   2040
            MaxLength       =   12
            PasswordChar    =   "*"
            TabIndex        =   3
            Text            =   "USR_PASS"
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   2040
            MaxLength       =   12
            TabIndex        =   1
            Text            =   "USR_NM"
            ToolTipText     =   "Maksimal nama user 12 Karakte"
            Top             =   120
            Width           =   2175
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   16
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Password (Ulang)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   15
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   14
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama User"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   13
            Top             =   120
            Width           =   1695
         End
      End
      Begin XtremeSuiteControls.ComboBox CmbNama 
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Top             =   360
         Width           =   2775
         _Version        =   851970
         _ExtentX        =   4895
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Sorted          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
         AutoComplete    =   -1  'True
         EnableMarkup    =   -1  'True
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Karyawan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ID Karyawan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Jabatan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Departemen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   1695
      End
   End
End
Attribute VB_Name = "USER_EDIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EditSts As String
Private Sub ChkStatus_Click(Index As Integer)
Select Case Index
    Case Is = 0
    Case Is = 1
        If ChkStatus(1).Value = False Then
            Text1(3).PasswordChar = "*"
            Text1(4).PasswordChar = "*"
            Text1(5).PasswordChar = "*"
        Else
            Text1(3).PasswordChar = ""
            Text1(4).PasswordChar = ""
            Text1(5).PasswordChar = ""
        End If
        
End Select
End Sub

Private Sub cmbKar_Validate(Index As Integer, Cancel As Boolean)
    If EditSts <> "EDIT" Then txtKarPrp(7).Text = PrjSysTrig.RunSqlFunction("fun_new_karyawan('" & GetDtaTbl(conMain, "SELECT CORP_ID FROM tab_corp WHERE CORP_NM='" & cmbKar(5).Text & "'", "CORP_ID") & "')")
End Sub

Private Sub CmbNama_Validate(Cancel As Boolean)
sndData = ""
PrjSysTrig.GetDataField conMain, "karyawan", "KAR_NM", "KAR_NM", CmbNama.Text
'MsgBox sndData
If CmbNama.Text <> sndData Then
    asd = MsgBox("Maaf Nama Karyawan Tersebut Belum Terdaftar, Apakah Anda ingin mendaftarakan karyawan baru?", vbYesNo)
    If asd = vbYes Then KaryawanEdit "add"
    If asd = vbNo Then UserEdit "show"
Else
    sndData = ""
    PrjSysTrig.DataFillText txtKar(0), conMain, "karyawan", "KAR_ID", "WHERE KAR_NM='" & CmbNama.Text & "'"
    PrjSysTrig.GetDataField conMain, "tab_user", "USR_NM", "USR_NM", CmbNama.Text
    If CmbNama.Text <> sndData Then
        sndData = CmbNama.Text
        UserEdit "add"
    ElseIf CmbNama.Text = sndData Then
        asd = MsgBox("Maaf Nama Karyawan '" & CmbNama.Text & "' Sudah memiliki USER ID, Apakah Anda ingin Meng-Edit User tersebut?", vbYesNo)
        If asd = vbYes Then UserEdit "edit"
        If asd = vbNo Then UserEdit "show"
    End If
End If
Fin:
End Sub

Private Sub Command1_Click(Index As Integer) 'tombol SAVE & CANCEL
Dim Val(5) As String
Dim strSql, IdKar As String
Dim ProStr As String
Dim oRS As ADODB.Recordset
Dim Finish As Label
Dim Cnt As String

sndData = ""
Select Case Index
 Case Is = 0
        If Text1(0).Text = "" Then
            MsgBox "Maaf Nama User Tidak Boleh Kosong !!!"
            Text1(0).SetFocus
            GoTo Finish
        ElseIf Text1(1).Text <> Text1(2).Text Then 'cek kecocokan password
            MsgBox "Maaf Password tidak sama !!!"
            Text1(1).SetFocus
            GoTo Finish
        End If
        
        PrjSysTrig.GetDataField conMain, "departemen", "DEP_ID", "DEP_NM", txtKar(1).Text 'cari DEP ID
        Val(1) = sndData
        
        'If ChkStatus(0).Value = 0 Then Val(2) = "0" Else Val(2) = "1" 'STATUS user off/aktif
            

        If sndChoice = "Add" Then 'kondisi menambah user
            'cek apakah user id sudah ada
                        PrjSysTrig.GetDataField conMain, "tab_user", "USER_ID", "USER_ID", Text1(0).Text
            'MsgBox sndData & Text1(0).Text
            If sndData = LTrim(Text1(0).Text) Then
                MsgBox "Maaf  USER ID '" & Text1(0).Text & "' sudah Terdaftar, Gunakan User ID lain"
                Text1(0).SetFocus
                GoTo Finish
            End If
            Val(2) = Str(ChkStatus(0).Value)
            Val(3) = "'" & Text1(0).Text & _
                "','" & Text1(1).Text & _
                "','" & CmbNama.Text & _
                "','" & txtKar(0).Text & _
                "','" & Val(2) & _
                "','" & Val(1) & "'"
            'MsgBox Val(3) & "  Msntsb"
            'insert data user
            PrjSysTrig.DataIns conMain, "tab_user(USER_ID,USR_PASS,USR_NM,KAR_ID,USR_OFF,DEP_ID)", Val(3)
            
            'insert menu standar untuk user
            ExecuteRecordSetMySql "INSERT INTO tab_user_per(MN_ID,PRMS_ID,MN_SHW) SELECT MN_ID,'" & Text1(0).Text & "','1' FROM tab_user_menu", conMain
        Else
         '--kondisi edit user
             Val(2) = Str(ChkStatus(0).Value)
             Val(3) = "USR_PASS='" & Text1(1).Text & _
                    "',USR_NM='" & CmbNama.Text & _
                    "',DEP_ID='" & Val(1) & _
                    "',USR_OFF='" & Val(2) & _
                    "' WHERE USER_ID='" & Text1(0).Text & "' "
            PrjSysTrig.DataUpdate conMain, "tab_user", Val(3)
        End If
    USER.refGrid
    Unload Me
 Case Is = 1, 5
    Unload Me
 Case Is = 2
    Find.Show
 Case Is = 3
 Case Is = 4 'save karyawan
    asd = MsgBox("Apakah data karyawan sudah benar?", vbYesNo)
   ' MsgBox asd
    If asd = vbYes Then
        For i = 0 To 6
            If txtKarPrp(i).Text = "" Or Len(txtKarPrp(i).Text) < 2 Then txtKarPrp(i).Text = "0"
        Next i
    
        ProStr = "('" & txtKarPrp(7) & "','" & _
                                txtKarPrp(0) & "','" & _
                                GetDtaTbl(conMain, "SELECT JAB_ID FROM jabatan WHERE JAB_NM='" & cmbKar(3).Text & "'", "JAB_ID") & "','" & _
                                txtKarPrp(1) & "','" & _
                                txtKarPrp(2) & "','" & _
                                txtKarPrp(5) & "','" & _
                                txtKarPrp(4) & "','" & _
                                datePickKar(0) & "','" & _
                                cmbKar(0) & "','" & _
                                cmbKar(1) & "','" & _
                                datePickKar(1) & "','" & _
                                datePickKar(2) & "','" & _
                                cmbKar(4) & "','" & _
                                txtKarPrp(3) & "','" & _
                                txtKarPrp(6) & "','" & _
                                GetDtaTbl(conMain, "SELECT CORP_ID FROM tab_corp WHERE CORP_NM='" & cmbKar(5).Text & "'", "CORP_ID") & "','" & _
                                GetDtaTbl(conMain, "SELECT CAB_ID FROM tab_cabang WHERE CAB_NM='" & cmbKar(6).Text & "'", "CAB_ID") & "','" & _
                                GetDtaTbl(conMain, "SELECT DEP_ID FROM departemen WHERE DEP_NM='" & cmbKar(2).Text & "'", "DEP_ID") & "')"
            If EditSts = "EDIT" Then
                PrjSysTrig.RunSqlStr conMain, "pro_karyawan_updt" & ProStr, True
             Else
                PrjSysTrig.RunSqlStr conMain, "pro_karyawan_ins" & ProStr, True
            End If
            Unload Me
        Else
            txtKarPrp(0).SetFocus
        End If
  
 Case Is = 6
    Cnt = GetDtaTbl(conMain, "SELECT COUNT(USR_PASS) AS ADA FROM tab_user WHERE USER_ID='" & UsrID & "' AND USR_PASS='" & Text1(5).Text & "'", "ADA")
    'MsgBox Text1(5).Text & "," & UsrID & "," & Cnt
    If Cnt = "1" Then
        If Text1(3).Text = Text1(4).Text Then
            PrjSysTrig.DataUpdate conMain, "tab_user", "USR_PASS='" & Text1(3).Text & "'", " WHERE USER_ID='" & UsrID & "'"
        Else
            MsgBox "Maaf Password Baru Tidak Sama...!", vbInformation
            Text1(4).SetFocus
        End If
    Else
        MsgBox "Password Lama Salah...!", vbExclamation
        Text1(5).SetFocus
    End If
End Select
Finish:
    'MsgBox sndData
End Sub

Private Sub ClearFrm(J As Integer)
Dim i As Integer
On Error Resume Next
Select Case J
Case Is = 1
    For i = 0 To 5
        CmbNama.Text = ""
        Text1(i).Text = ""
        txtKar(i).Text = ""
    Next i
Case Is = 2
End Select
End Sub



Private Sub Form_Load()
Main
   UserEdit "add"         ' ChkStatus(0).Value = Int(GetDtaTbl(conMain, "SELECT USR_OFF FROM tab_user WHERE USER_ID='" & Text1(0).Text & "'", "USR_OFF"))
End Sub

Private Sub Form_Unload(Cancel As Integer)
'  fraNew(0).Visible = False
End Sub
Public Sub UserEdit(Val As String)
'On Error GoTo ErrorLabel
LOADING.Show
LOADING.SetParm Me, 25

    Me.Width = 5295
    Me.Height = 3255 '5880
    fraNew(0).Height = 2535
    CenterForm
    ClearFrm (1)
    fraNew(0).Visible = True
    fraNew(0).left = 120
    fraNew(0).top = 120
    fraUsr(0).Enabled = False
    Text1(0).Enabled = False
    Text1(0).BackColor = &HC0FFFF

    Combo1.Clear
    CmbNama.Clear
    CmbNama.Enabled = True
    
    FillCmbBox 1
    'isi combo box field nama karyawan
    'PrjSysTrig.DataFillCombo Combo1, conMain, "karyawan", "KAR_NM", "ORDER BY KAR_NM ASC", "KAR_NM", "KAR_NM"
    'For I = 0 To Combo1.ListCount - 1
    '    CmbNama.AddItem (Combo1.List(I))
        'CmbNama.ListIndex = 0
    'Next I
    'CmbNama.ListIndex = 0
LOADING.SetParm Me, 35

Select Case Val
Case Is = "show"
    Me.Caption = "User Add/Edit"
    Exit Sub
Case Is = "add" 'tambah data baru user
    Me.Caption = "User Add"
    'MsgBox sndData & CmbNama.Text
    CmbNama.Text = sndData
     'CmbNama.Enabled = True
     LOADING.SetParm Me, 45
    PrjSysTrig.DataFillText txtKar(0), conMain, "karyawan", "KAR_ID", "WHERE KAR_NM='" & CmbNama.Text & "'"
    PrjSysTrig.DataFillText txtKar(1), conMain, "karyawan,departemen", "DEP_NM", "WHERE KAR_ID='" & txtKar(0).Text & "' AND karyawan.DEP_ID=departemen.DEP_ID"
    PrjSysTrig.DataFillText txtKar(2), conMain, "karyawan,jabatan", "JAB_NM", "WHERE KAR_ID='" & txtKar(0).Text & "' AND karyawan.JAB_ID=jabatan.JAB_ID"
    
    LOADING.SetParm Me, 65
    fraUsr(0).Enabled = True
    Text1(0).Enabled = True
    Text1(0).BackColor = &HFFFFFF
    Text1(0).Text = ""
    Text1(1).Text = ""
    Text1(2).Text = ""
'    Text1(0).SetFocus
    sndChoice = "Add"
    Me.Height = 5880
    fraNew(0).Height = 5175
     fraNew(0).Enabled = True


Case Is = "edit" 'edit data user
    Me.Caption = "User Edit"
    CmbNama.Text = sndData
    CmbNama.Enabled = False
    'CmbNama.ListIndex = CmbNama.FindItem(0, sndData, True)
    LOADING.SetParm Me, 45
    PrjSysTrig.DataFillText txtKar(0), conMain, "tab_user", "KAR_ID", "WHERE USR_NM='" & sndData & "'"
    If USER.dxDBGridUser.Columns.Item(2).Value = 0 Then ChkStatus(0).Value = 1 Else ChkStatus(0).Value = 0
    
    LOADING.SetParm Me, 65
        PrjSysTrig.DataFillText txtKar(1), conMain, "karyawan,departemen", "DEP_NM", "WHERE KAR_ID='" & txtKar(0).Text & "' AND karyawan.DEP_ID=departemen.DEP_ID"
        PrjSysTrig.DataFillText txtKar(2), conMain, "karyawan,jabatan", "JAB_NM", "WHERE KAR_ID='" & txtKar(0).Text & "' AND karyawan.JAB_ID=jabatan.JAB_ID"
        PrjSysTrig.DataFillText Text1(0), conMain, "tab_user", "USER_ID", "WHERE USR_NM='" & sndData & "'"
        PrjSysTrig.DataFillText Text1(1), conMain, "tab_user", "USR_PASS", "WHERE USER_ID='" & Text1(0).Text & "'"
        Text1(2).Text = Text1(1).Text
        fraUsr(0).Enabled = True
        Text1(0).Enabled = False
        Text1(1).SetFocus
        sndChoice = "Edit"
        Me.Height = 5880
        fraNew(0).Height = 5175
Case Is = "pass" 'edit user password
    Me.Caption = "User Edit Password"
    fraUsr(1).Enabled = True
    fraUsr(1).top = 120
    fraUsr(1).left = 120
    Me.Height = 3225
    fraNew(0).left = 5280
    Text1(5).SetFocus
    
End Select

LOADING.SetParm Me, 100
'Exit Sub
'ErrorLabel:
 '   If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"
 '   LOADING.SetParm Me, 100
End Sub

Public Sub KaryawanEdit(Val As String)
On Error GoTo ErrorLabel
LOADING.Show
LOADING.SetParm Me, 25

    Me.Width = 11820
    Me.Height = 5880
    ClearFrm (1)
    CenterForm
    datePickKar(1) = Format(Now)
    datePickKar(2) = Format(Now)
    
LOADING.SetParm Me, 35
   
    fraNew(1).Visible = True
    fraNew(1).left = 120
    fraNew(1).top = 120
    txtKarPrp(0).SetFocus
    If LCase(UsrID) = "root" Then
        cmbKar(5).Enabled = True
        cmbKar(6).Enabled = True
    End If

LOADING.SetParm Me, 45
     FillCmbBox 1
     FillCmbBox 2
     FillCmbBox 3
     FillCmbBox 4
     FillCmbBox 5
LOADING.SetParm Me, 65
     FillCmbBox 6
     FillCmbBox 7
     FillCmbBox 8

LOADING.SetParm Me, 75
Select Case Val
Case Is = "add" 'tambah data baru user
    
    LOADING.SetParm Me, 80
    cmbKar(5).Text = GetDtaTbl(conMain, "SELECT CORP_NM FROM tab_corp WHERE CORP_ID='" & IdCorp & "'", "CORP_NM")
    cmbKar(6).Text = GetDtaTbl(conMain, "SELECT CAB_NM FROM tab_cabang WHERE CAB_ID='" & IdCab & "'", "CAB_NM")
    
    txtKarPrp(7).Text = PrjSysTrig.RunSqlFunction("fun_new_karyawan('" & GetDtaTbl(conMain, "SELECT CORP_ID FROM tab_corp WHERE CORP_NM='" & cmbKar(5).Text & "'", "CORP_ID") & "')")
Case Is = "edit" '
    LOADING.SetParm Me, 80
End Select

LOADING.SetParm Me, 100
Exit Sub
ErrorLabel:
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"
    LOADING.SetParm Me, 100
End Sub
Public Sub CenterForm()
    ' Me.Move (Screen.Width - Me.Width) / 2, _
    ' (Screen.Height - Me.Height) / 2
End Sub
Private Sub FillCmbBox(N As Integer) ' isi list item combo box
Dim Itm() As String
Dim K As Integer

Select Case N
Case Is = 1, 0
    CmbNama.Clear
    Itm = GetListTabel(conMain, "SELECT KAR_NM FROM karyawan  WHERE KAR_NM<>'' ORDER BY KAR_NM ASC ", "KAR_NM")
    K = 1
    Do Until Itm(K) = ""
     CmbNama.AddItem Itm(K)
     K = K + 1
    Loop
    CmbNama.ListIndex = 0
    Erase Itm
Case Is = 2, 0
    cmbKar(5).Clear
    Itm = GetListTabel(conMain, "SELECT CORP_NM FROM tab_corp", "CORP_NM")
    K = 1
    Do Until Itm(K) = ""
     cmbKar(5).AddItem Itm(K)
     K = K + 1
    Loop
    cmbKar(5).ListIndex = 0
    Erase Itm
Case Is = 3, 0
    cmbKar(2).Clear
    Itm = GetListTabel(conMain, "SELECT DEP_NM FROM departemen WHERE DEP_NM<>'SYSTEM'", "DEP_NM")
    K = 1
    Do Until Itm(K) = ""
     cmbKar(2).AddItem Itm(K)
     K = K + 1
    Loop
    cmbKar(2).ListIndex = 0
    Erase Itm
Case Is = 4, 0
    cmbKar(3).Clear
    Itm = GetListTabel(conMain, "SELECT JAB_NM FROM jabatan WHERE JAB_NM<>'SYSTEM'", "JAB_NM")
    K = 1
    Do Until Itm(K) = ""
     cmbKar(3).AddItem Itm(K)
     K = K + 1
    Loop
    cmbKar(3).ListIndex = 0
    Erase Itm
Case Is = 5, 0
    cmbKar(0).Clear
    cmbKar(0).AddItem "Islam"
    cmbKar(0).AddItem "Kristen Protestan"
    cmbKar(0).AddItem "Kristen Katolik"
    cmbKar(0).AddItem "Hindu"
    cmbKar(0).AddItem "Buddha"
    cmbKar(0).AddItem "Konghucu"
    cmbKar(0).ListIndex = 0
Case Is = 6, 0
    cmbKar(1).Clear
    cmbKar(1).AddItem "Lajang"
    cmbKar(1).AddItem "Menikah"
    cmbKar(1).AddItem "Cerai"
    cmbKar(1).ListIndex = 0
Case Is = 7, 0
    cmbKar(4).Clear
    cmbKar(4).AddItem "Kontrak"
    cmbKar(4).AddItem "Tetap"
    cmbKar(4).AddItem "Freelance"
    cmbKar(4).ListIndex = 0
Case Is = 8, 0
    cmbKar(6).Clear
    Itm = GetListTabel(conMain, "SELECT CAB_NM FROM tab_cabang", "CAB_NM")
    K = 1
    Do Until Itm(K) = ""
     cmbKar(6).AddItem Itm(K)
     K = K + 1
    Loop
    cmbKar(6).ListIndex = 0
    Erase Itm

End Select
End Sub


Private Sub txtKarPrp_Click(Index As Integer)
 txtKarPrp(Index).SelStart = 0
      txtKarPrp(Index).SelLength = (Len(txtKarPrp(Index).Text))

  ''Clipboard.Clear
    'Clipboard.SetText txtKarPrp(Index).Text
End Sub

Private Sub txtKarPrp_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
    Case Is = 1, 5, 4
        CekText_Number KeyAscii
End Select
End Sub

Private Sub CekText_Number(sndKey As Integer)
'Accepts only numeric input
Select Case sndKey
  Case vbKey0 To vbKey9
  Case vbKeyBack, vbKeyClear, vbKeyDelete
  Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
  Case Else
    sndKey = 0
    Beep
End Select
End Sub
Private Sub ResetForm()
Dim i%

For i = 0 To 6
cmbKar(i).Text = ""
If i <= 2 Then datePickKar(i) = ""
Next i

End Sub

Public Sub EditParm(prmId As String, prmCorp As String, prmCab As String, NewSts As Boolean)
Dim strSql As String
On Error GoTo ErrorLabel
LOADING.Show
LOADING.SetParm Me, 25

        ResetForm 'bersihkan form
        
        'SET COMBO corp & cabang
        cmbKar(5).Text = GetDtaTbl(conMain, "SELECT CORP_NM FROM tab_corp WHERE CORP_ID='" & prmCorp & "'", "CORP_NM")
        cmbKar(6).Text = GetDtaTbl(conMain, "SELECT CAB_NM FROM tab_cabang WHERE CAB_ID='" & prmCab & "'", "CAB_NM")
LOADING.SetParm Me, 35
    If NewSts Then 'baca jika status Edit & New
    LOADING.SetParm Me, 45
        'EditSts = "NEW"
        
    Else
    LOADING.SetParm Me, 45
        EditSts = "EDIT"
        Me.Caption = "KARYAWAN - "
        Me.Caption = Me.Caption & EditSts
                
        'fill data edit
        LOADING.SetParm Me, 55
        txtKarPrp(7).Text = prmId
        txtKarPrp(0).Text = GetDtaTbl(conMain, "SELECT KAR_NM FROM karyawan WHERE KAR_ID='" & prmId & "'", "KAR_NM")
        txtKarPrp(1).Text = GetDtaTbl(conMain, "SELECT KAR_KTP FROM karyawan WHERE KAR_ID='" & prmId & "'", "KAR_KTP")
        datePickKar(0) = GetDtaTbl(conMain, "SELECT KAR_TGL FROM karyawan WHERE KAR_ID='" & prmId & "'", "KAR_TGL")
        txtKarPrp(2).Text = GetDtaTbl(conMain, "SELECT KAR_ALMT FROM karyawan WHERE KAR_ID='" & prmId & "'", "KAR_ALMT")
        cmbKar(0).Text = GetDtaTbl(conMain, "SELECT KAR_AGM FROM karyawan WHERE KAR_ID='" & prmId & "'", "KAR_AGM")
        cmbKar(1).Text = GetDtaTbl(conMain, "SELECT KAR_STSK FROM karyawan WHERE KAR_ID='" & prmId & "'", "KAR_STSK")
        txtKarPrp(5).Text = GetDtaTbl(conMain, "SELECT KAR_TLP FROM karyawan WHERE KAR_ID='" & prmId & "'", "KAR_TLP")
        txtKarPrp(4).Text = GetDtaTbl(conMain, "SELECT KAR_HP FROM karyawan WHERE KAR_ID='" & prmId & "'", "KAR_HP")
        cmbKar(2).Text = GetDtaTbl(conMain, "SELECT DEP_NM FROM departemen WHERE DEP_ID='" & _
                         GetDtaTbl(conMain, "SELECT DEP_ID FROM karyawan WHERE KAR_ID='" & prmId & "'", "DEP_ID") & "'", "DEP_NM")
        cmbKar(3).Text = GetDtaTbl(conMain, "SELECT JAB_NM FROM jabatan WHERE JAB_ID='" & _
                            GetDtaTbl(conMain, "SELECT JAB_ID FROM karyawan WHERE KAR_ID='" & prmId & "'", "JAB_ID"), "JAB_NM")
        txtKarPrp(3).Text = GetDtaTbl(conMain, "SELECT KAR_MAILP FROM karyawan WHERE KAR_ID='" & prmId & "'", "KAR_MAILP")
        txtKarPrp(6).Text = GetDtaTbl(conMain, "SELECT KAR_MAILK FROM karyawan WHERE KAR_ID='" & prmId & "'", "KAR_MAILK")
        cmbKar(4).Text = GetDtaTbl(conMain, "SELECT KAR_STSP FROM karyawan WHERE KAR_ID='" & prmId & "'", "KAR_STSP")
        datePickKar(1) = GetDtaTbl(conMain, "SELECT KAR_TGLM FROM karyawan WHERE KAR_ID='" & prmId & "'", "KAR_TGLM")
        datePickKar(2) = GetDtaTbl(conMain, "SELECT KAR_TGLK FROM karyawan WHERE KAR_ID='" & prmId & "'", "KAR_TGLK")
       ' txtKarPrp(0).Text = GetDtaTbl(conMain, "SELECT KAR_NM FROM karyawan WHERE KAR_ID='" & prmId & "'", "KAR_NM")
        
    End If
    
LOADING.SetParm Me, 100
Exit Sub
ErrorLabel:
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"
    LOADING.SetParm Me, 100
    
End Sub
