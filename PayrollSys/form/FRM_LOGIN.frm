VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Begin VB.Form FRM_LOGIN 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   4410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10845
   Icon            =   "FRM_LOGIN.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox grpLogin 
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   4575
      _Version        =   851970
      _ExtentX        =   8070
      _ExtentY        =   5741
      _StockProps     =   79
      Caption         =   "Basic Setting"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton cmdLogin 
         Height          =   495
         Index           =   0
         Left            =   1080
         TabIndex        =   5
         Top             =   4320
         Width           =   1215
         _Version        =   851970
         _ExtentX        =   2143
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "&Save"
         BackColor       =   16744576
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton cmdLogin 
         Height          =   495
         Index           =   1
         Left            =   2400
         TabIndex        =   6
         Top             =   4320
         Width           =   1215
         _Version        =   851970
         _ExtentX        =   2143
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "&Cancel"
         BackColor       =   16744576
         Appearance      =   6
      End
      Begin XtremeSuiteControls.FlatEdit txtLgn 
         Height          =   375
         Index           =   0
         Left            =   1680
         TabIndex        =   7
         Top             =   480
         Width           =   2655
         _Version        =   851970
         _ExtentX        =   4683
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtLgn 
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   9
         Top             =   960
         Width           =   2655
         _Version        =   851970
         _ExtentX        =   4683
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Text            =   "eka"
      End
      Begin XtremeSuiteControls.FlatEdit txtLgn 
         Height          =   375
         Index           =   2
         Left            =   1680
         TabIndex        =   11
         Top             =   1440
         Width           =   2655
         _Version        =   851970
         _ExtentX        =   4683
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Text            =   "asd123"
         PasswordChar    =   "*"
      End
      Begin XtremeSuiteControls.PushButton cmdLogin 
         Height          =   375
         Index           =   2
         Left            =   1680
         TabIndex        =   13
         Top             =   1920
         Width           =   1335
         _Version        =   851970
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "&Test Conection"
         BackColor       =   16744576
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton cmdLogin 
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   14
         Top             =   2400
         Width           =   1335
         _Version        =   851970
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "&Backup DB"
         BackColor       =   16744576
         Enabled         =   0   'False
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton cmdLogin 
         Height          =   375
         Index           =   4
         Left            =   1680
         TabIndex        =   15
         Top             =   2400
         Width           =   1335
         _Version        =   851970
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "&Restore DB"
         BackColor       =   16744576
         Enabled         =   0   'False
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton cmdLogin 
         Height          =   375
         Index           =   5
         Left            =   3120
         TabIndex        =   16
         Top             =   2400
         Width           =   1335
         _Version        =   851970
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "&Reset DB"
         BackColor       =   16744576
         Enabled         =   0   'False
         Appearance      =   6
      End
      Begin XtremeSuiteControls.Label lblLogin 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   1575
         _Version        =   851970
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Pass Database"
      End
      Begin XtremeSuiteControls.Label lblLogin 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1575
         _Version        =   851970
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "User Database"
      End
      Begin XtremeSuiteControls.Label lblLogin 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1575
         _Version        =   851970
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Server"
      End
   End
   Begin XtremeSuiteControls.TabControlPage TabControlPage1 
      Height          =   2655
      Left            =   720
      TabIndex        =   26
      Top             =   600
      Width           =   3135
      _Version        =   851970
      _ExtentX        =   5530
      _ExtentY        =   4683
      _StockProps     =   1
      Picture         =   "FRM_LOGIN.frx":1272
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   4095
      Left            =   5400
      TabIndex        =   17
      Top             =   120
      Width           =   5295
      _Version        =   851970
      _ExtentX        =   9340
      _ExtentY        =   7223
      _StockProps     =   79
      Caption         =   "Login User"
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtLogin 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   2400
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   2760
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Login"
         Height          =   615
         Index           =   0
         Left            =   600
         TabIndex        =   3
         Top             =   3240
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         Height          =   615
         Index           =   1
         Left            =   2760
         TabIndex        =   4
         Top             =   3240
         Width           =   2175
      End
      Begin VB.TextBox txtLogin 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   2400
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox txtLogin 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   2400
         TabIndex        =   1
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtLogin 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   2400
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox txtLogin 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   2400
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   495
         Index           =   4
         Left            =   960
         TabIndex        =   25
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ID User"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   495
         Index           =   0
         Left            =   840
         TabIndex        =   23
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   495
         Index           =   1
         Left            =   840
         TabIndex        =   22
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   495
         Index           =   3
         Left            =   840
         TabIndex        =   21
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   495
         Index           =   5
         Left            =   840
         TabIndex        =   20
         Top             =   2160
         Width           =   1455
      End
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   375
      Left            =   600
      TabIndex        =   27
      Top             =   3240
      Width           =   3135
      _Version        =   851970
      _ExtentX        =   5530
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Payroll System Ver. 1.0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
End
Attribute VB_Name = "FRM_LOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dt, tm, rsl As Date
Dim L As New LOADING
Private Sub cmdLogin_Click(Index As Integer)
Select Case Index
    Case Is = 0
        If txtLgn(0) = "" Then
            MsgBox "Isi Informasi Server Anda dengan Lengkap!", vbInformation, "Info"
        Else
            lpFilename = App.Path & "\config.ini"

            lpSectionName = "Koneksi"
            lpKeyName = "Server"
            lpValueUser = txtLgn(0).Text

            Call ProfileSaveItem(lpSectionName, _
                     lpKeyName, _
                     lpValueUser, _
                     lpFilename)
          
          Main
          grpLogin.Visible = False

        End If
    Case Is = 1
        grpLogin.Visible = False
    Case Is = 2
        strServer = txtLgn(0).Text
        Main
          If conMain.State = 1 Then
            MsgBox "Koneksi Database Berhasil"
          Else
            MsgBox "Maaf Aplikasi tidak dapat terhubung dengan Server...", vbExclamation
          End If
End Select
End Sub

Private Sub Form_Load()
Dim defIP As String
InfoServ
On Error GoTo ErrorLabel
Main
    
'Image1.left = Me.Width / 6.5
'Image1.top = Me.Height / 12
'Image1.left = Me.left
'Image1.top = Me.top
'Image1.Height = Screen.Height
'Image1.Width = Screen.Width

'Frame1.left = Screen.Width - 5100
'Frame1.top = 500
dt = DateValue(Now)
tm = TimeValue(Now)
rsl = dt + tm
txtLogin(4).Text = Format(rsl, "dd/mm/yyyy h:m:s")
txtLogin(4).Enabled = False
Me.Icon = LoadPicture(ImagePath("FRM_MAIN"))
'MDIForm1.Show

If CekStatusFile("config.ini", App.Path) Then
    
Else
    
End If

InfoServ

defIP = strServer
  'If Ping(defIP) = False Then
  '      MsgBox "Koneksi Database Terputus...!!!", vbCritical
 ' Exit Sub
 ' End If





Exit Sub
ErrorLabel:
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"
End Sub

Private Sub PushButton1_Click()
'MsgBox prjSysID.PaidSaid(txtNum.Text)
End Sub


'========================
'==== VALIDASI LOGIN ====
'========================
Private Sub txtLogin_Change(Index As Integer)
On Error GoTo PastiErorr
If Index = 0 Then
    If txtLogin(0).Text * 0 = 0 Then txtLogin(0).PasswordChar = "*"
    Exit Sub
PastiErorr:
    txtLogin(0).PasswordChar = ""
End If

'-- RUN SCHECULE -
If txtLogin(0).Text = "monitor" Then
    OpRecStt1 "run_scdl_payroll()", True
End If

End Sub

Private Sub txtLogin_Click(Index As Integer)
  txtLogin(Index).SelStart = 0
     txtLogin(Index).SelLength = (Len(txtLogin(Index).Text))
End Sub

Private Sub txtLogin_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case Is = 121
        'InfoServ
        'txtLgn(0).Text = strServer
        'grpLogin.top = Frame1.top + 4440
        'grpLogin.left = Frame1.left
        'grpLogin.Visible = True
        
    Case Is = 13
        Command1_Click (0)
End Select

End Sub

Private Sub txtLogin_LostFocus(Index As Integer)
If Index = 0 Then
     GetUserInfo (txtLogin(0).Text)
    txtLogin(1).Text = UsrNM
    txtLogin(2).Text = DepNM
    txtLogin(1).Text = PrjSysUsr.ValidasiUserID(txtLogin(0).Text)
End If
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0:
 
        LOADING.Show
        LOADING.SetParm Me, 50
        If PrjSysUsr.ValidasiUserID(txtLogin(0).Text) <> "" _
        And PrjSysUsr.ValidasiPass(txtLogin(3).Text) = 1 Then
             'Dim ControlMenu As CommandBarPopup
             'If MDIForm1.CFrame.ActiveMenuBar.Controls.Count <> 0 Then
               ' Set ControlMenu = MDIForm1.CFrame.
                '.Delete
            ' End If
                    LOADING.SetParm Me, 75
            'Sleep 1000
            UsrID = txtLogin(0).Text
            LOADING.SetParm Me, 100
            MDIForm1.Show
            PrjSysMn.LoginCaption txtLogin(0).Text
             MDIForm1.Caption = CorpNM & " - " & CabNM
            'PrjSystem.SysMenu.CreateRiboonBar MDIForm1.CBar1, "ROOT"
           ' PrjSystem.SysMenu.GET_MenuEditor MDIForm1.CFrame, txtLogin(0).Text
            'MDIForm1.LblLogin = txtLogin(0).Text
            Unload Me
            Exit Sub
        Else
            'If Len(PrjSysUsr.ValidasiRFID(txtLogin(0).Text)) > 1 Then
            '    UsrID = PrjSysUsr.ValidasiRFID(txtLogin(0).Text)
            '    GetUserInfo PrjSysUsr.ValidasiRFID(txtLogin(0).Text)
            '    LOADING.SetParm Me, 100
            '    MDIForm1.Show
            '    PrjSysMn.LoginCaption PrjSysUsr.ValidasiRFID(txtLogin(0).Text)
            '    MDIForm1.Caption = "LUKISON GROUP - " & CorpNM & " - " & CabNM
            'PrjSystem.SysMenu.CreateRiboonBar MDIForm1.CBar1, "ROOT"
            ' PrjSystem.SysMenu.GET_MenuEditor MDIForm1.CFrame, txtLogin(0).Text
            'MDIForm1.LblLogin = txtLogin(0).Text
            '    Unload Me
            'Exit Sub
            'End If
            LOADING.SetParm Me, 100
            MsgBox "User atau Password Salah ! ", , "USER INFO"
        End If
  
Case 1: UnloadAll
End Select
End Sub

