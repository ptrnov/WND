VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Begin VB.Form HOME 
   BackColor       =   &H00FFFFFF&
   Caption         =   "HOME"
   ClientHeight    =   10950
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
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10950
   ScaleWidth      =   20160
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.TabControlPage TabControlPage1 
      Height          =   12975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   21375
      _Version        =   851970
      _ExtentX        =   37703
      _ExtentY        =   22886
      _StockProps     =   1
      Picture         =   "HOME.frx":0000
      Begin VB.Timer Timer1 
         Left            =   240
         Top             =   7560
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Auto Refresh"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   13
         Top             =   6960
         Width           =   2415
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   375
         Index           =   0
         Left            =   9720
         TabIndex        =   11
         Top             =   1020
         Width           =   1455
         _Version        =   851970
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "<< Previous"
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
         Height          =   5415
         Left            =   240
         OleObjectBlob   =   "HOME.frx":DE7B1
         TabIndex        =   9
         Top             =   1440
         Width           =   12375
      End
      Begin VB.Timer tmrUpdt 
         Interval        =   1000
         Left            =   14280
         Top             =   240
      End
      Begin XtremeSuiteControls.GroupBox grpSts 
         Height          =   9135
         Left            =   15000
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   4335
         _Version        =   851970
         _ExtentX        =   7646
         _ExtentY        =   16113
         _StockProps     =   79
         Caption         =   "GroupBox1"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.MonthCalendar CalHome 
            Height          =   2745
            Left            =   360
            TabIndex        =   8
            Top             =   360
            Width           =   3735
            _Version        =   851970
            _ExtentX        =   6588
            _ExtentY        =   4842
            _StockProps     =   68
            MinDate         =   20341
            CurrentDate     =   41844
         End
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   375
         Index           =   1
         Left            =   11160
         TabIndex        =   12
         Top             =   1020
         Width           =   1455
         _Version        =   851970
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Today"
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Daily Log Absensi"
         BeginProperty Font 
            Name            =   "Electrofied"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1180
         Width           =   3855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         Index           =   1
         X1              =   120
         X2              =   14520
         Y1              =   960
         Y2              =   960
      End
      Begin XtremeSuiteControls.Label lblHome 
         Height          =   375
         Index           =   0
         Left            =   1320
         TabIndex        =   6
         Top             =   120
         Width           =   3015
         _Version        =   851970
         _ExtentX        =   5318
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "User ID .........................................................................."
         ForeColor       =   8454143
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblHome 
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   5
         Top             =   480
         Width           =   3135
         _Version        =   851970
         _ExtentX        =   5530
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "User Kar_id ....................................................................................."
         ForeColor       =   8454143
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         Index           =   0
         X1              =   7560
         X2              =   7560
         Y1              =   120
         Y2              =   960
      End
      Begin XtremeSuiteControls.Label lblHome 
         Height          =   375
         Index           =   2
         Left            =   4680
         TabIndex        =   4
         Top             =   120
         Width           =   2655
         _Version        =   851970
         _ExtentX        =   4683
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "User Dep  ........................................................................."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblHome 
         Height          =   375
         Index           =   3
         Left            =   7680
         TabIndex        =   3
         Top             =   120
         Width           =   6855
         _Version        =   851970
         _ExtentX        =   12091
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "User Corp ............................................................................"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblHome 
         Height          =   375
         Index           =   4
         Left            =   7680
         TabIndex        =   2
         Top             =   480
         Width           =   6855
         _Version        =   851970
         _ExtentX        =   12091
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "User Cab  ................................................................................................"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin VB.Image Image1 
         Height          =   735
         Left            =   120
         Picture         =   "HOME.frx":DF44B
         Stretch         =   -1  'True
         Top             =   120
         Width           =   855
      End
      Begin XtremeSuiteControls.Label lblHome 
         Height          =   375
         Index           =   5
         Left            =   4680
         TabIndex        =   1
         Top             =   480
         Width           =   2655
         _Version        =   851970
         _ExtentX        =   4683
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "User Jab  ........................................................................."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
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
End
Attribute VB_Name = "HOME"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GroupSubject As TaskPanelGroup
Dim GroupTab As TaskPanelGroup
Dim GroupMsg As TaskPanelGroup
Dim Item As TaskPanelGroupItem
Dim CLM0, CLM1 As Variant
Dim BND0 As Variant
Dim Qry0 As String
Dim DateSort, NewRoId, RoIdDef, RoTglDef As String
Dim NewRo As Boolean
Dim NoClm As Integer
Dim kurang As Integer



Private Sub TaskPanelStupGl_ItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
Dim StrMessage1 As String
Dim StrMessage2 As String
Dim StrMessage3 As String
Dim StrSubject1 As String
Dim StrSubject2 As String
Dim StrSubject3 As String
StrMessage1 = "Form RO (View, Tambah, Hapus dan Edit)  "
StrMessage2 = "Form RO ..."
StrMessage3 = "Form RO ..."
StrSubject1 = "Form Request Order"
StrSubject2 = "Form RO ..."
StrSubject3 = "Form RO ..."

If Item.Id = 1 Then
    GroupSubject.Caption = StrSubject1
    GroupMsg.Items.Find(1).Remove
    GroupMsg.Items.Add 1, StrMessage1, xtpTaskItemTypeText, 4
    'TabControl1.SelectedItem = 0
End If
If Item.Id = 2 Then
    GroupSubject.Caption = StrSubject2
    GroupMsg.Items.Find(1).Remove
    GroupMsg.Items.Add 1, StrMessage2, xtpTaskItemTypeText, 4
    'TabControl1.SelectedItem = 1
End If
If Item.Id = 3 Then
    GroupSubject.Caption = StrSubject3
    GroupMsg.Items.Find(1).Remove
    GroupMsg.Items.Add 1, StrMessage3, xtpTaskItemTypeText, 4
    'TabControl1.SelectedItem = 2
End If
End Sub

Sub InitExample()
 'dxDBGrid1.Event = 1 'EGOnCustomDrawCell
 'dxDBGrid1.EventEnabled = True
End Sub

Private Sub dxDBGrid1_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
Dim i As Integer
'i = i + 1
'        If Color = RGB(255, 255, 255) Then Color = RGB(255, 0, 255) Else Color = RGB(255, 255, 255)
        ' If Node.IsLast = True Then NoClm = 0
         
'Column.CurrencyVal

'Node.SummaryValues(dxDBGrid1.Columns.ColumnByFieldName("NO.").Index) = I
'MsgBox Node.RecNo
End Sub

Private Sub Form_Activate()

TabControlPage1.Width = Me.Width
TabControlPage1.Height = Me.Height

'tab home properti form
'tabHome.Width = Me.Width
'tabHome.Height = Me.Height

'group status properti form
'grpSts.left = Me.Width - 4575
'grpSts.Height = Me.Height - 1500

'status properti forms
'stsHome.Width = Me.Width
'stsHome.Height = Me.Height
'stsHome.top = Me.Height - 1200
'stsHome.left = Me.left

End Sub

Private Sub Form_Load()
On Error GoTo ErrorLabel
'InfoServ
Main
'GetUserInfo "alam_fod"


 LOADING.Show
 LOADING.SetParm Me, 25
InitExample

 LOADING.SetParm Me, 50
'CreateTaskPanel
Me.Icon = LoadPicture(ImagePath("FRM_HOME"))
    XtremeSuiteControls.Icons.LoadBitmap App.Path & "\Img\Menu\user.png", 13, xtpImageNormal

 LOADING.SetParm Me, 85
lblHome(0).Caption = UsrID
lblHome(1).Caption = KarId
lblHome(2).Caption = DepNM
lblHome(5).Caption = JabNM
lblHome(3).Caption = CorpNM
lblHome(4).Caption = CabNM
CalHome.Value = Format(Now)

LOADING.SetParm Me, 100
GridFill 0
kurang = 0
Check1.Value = 1
Timer1_Timer
Exit Sub
ErrorLabel:
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"
    LOADING.SetParm Me, 100
    
End Sub


'=====================================
'============= ptr.nov ===============
'===== PROCESS FIRST GRID LODING =====
'=====================================
Private Sub GridFill(GrdIndx As Integer)
Select Case GrdIndx
    Case Is = 0
        MenuGrid 0
            Qry0 = "personallog_monitoring(0)"
            PrjSysGrid.GetGrid_Persensi dxDBGrid1, CLM0, False, False, BND0, True, Qry0, "TerminalID", False
            dxDBGrid1.Dataset.Open
    End Select
End Sub

'=====================================
'============= ptr.nov ===============
'===== MENU CULUMN BND GRID LODING ===
'=====================================
Private Sub MenuGrid(TabSel As Integer)
Select Case TabSel
Case Is = 0
    BND0 = Array("TIME TABLE ", "PERSONAL ABSENSI VALUES")
    CLM0 = Array(Array("Clm0", "Time", "DateTime", gedDateEdit, 0, 0, 135, 1, 2, 0, 2), _
                Array("Clm1", "Device", "TerminalNm", gedTextEdit, 0, 0, 200, 6, 0, 0, 0), _
                Array("Clm2", "Event", "FunctionKey", gedTextEdit, 0, 0, 70, 6, 0, 0, 0), _
                Array("Clm3", "FingerID", "FingerPrintID", gedTextEdit, 0, 0, 70, 6, 0, 0, 0), _
                Array("Clm4", "User Machine", "username", gedTextEdit, 0, 1, 150, 6, 0, 0, 0), _
                Array("Clm5", "Employe name", "KarNm", gedTextEdit, 0, 1, 180, 6, 0, 0, 0))
                '--------0---------1--------2-----------3--------4--5--6---7--8--9--10
                'TT_ID, TT_GRP_ID,TT_TYP,TT_TYP_KTG,TT_THN,TT_SDAYS,TT_EDAYS,TT_SDATE,TT_EDATE,TT_NOTE,TT_UPDT,TT_ACTIVE,RULE_IN,
                'RULE_OUT , RULE_TOL_IN, RULE_TOL_OUT, RULE_BRK_OUT, RULE_BRK_IN, RULE_DRT_OT_DPN, RULE_DRT_OT_BLK, RULE_DURATION
End Select
End Sub

Private Sub Normilize()
'dxDBGridROitem.Columns.DestroyColumns
'dxDBGridROList.Columns.DestroyColumns
'dxDBGridROList.Dataset.Close
'dxDBGridROitem.Dataset.Close

DateSort = "RO_TGL='" & Format(Now, "YYYY-MM-DD") & "'"
'DtePick(0).Value = Now
'DtePick(1).Value = Now
NewRo = False
'CmdRo(1).Visible = False
'CmdRo(2).Visible = False
MenuGrid (1)
GridFill (1)
End Sub

Private Sub ClearFrm()
    'txtDep.Text = ""
    'txtIDRO.Text = ""
    'txtTgl.Text = ""
    'txtUser.Text = ""
End Sub


Private Sub Pop_StateChanged(Index As Integer)
    
End Sub




Private Sub tabHome_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

End Sub

Private Sub PushButton1_Click(Index As Integer)
If Index = 0 Then
    kurang = kurang + 1
        MenuGrid 0
                Qry0 = "personallog_monitoring(" & kurang & ")"
                PrjSysGrid.GetGrid_Persensi dxDBGrid1, CLM0, False, False, BND0, True, Qry0, "TerminalID", False
                dxDBGrid1.Dataset.Open
        Check1.Value = 0
End If
If Index = 1 Then
    kurang = 0
    MenuGrid 0
            Qry0 = "personallog_monitoring(0)"
            PrjSysGrid.GetGrid_Persensi dxDBGrid1, CLM0, False, False, BND0, True, Qry0, "TerminalID", False
            dxDBGrid1.Dataset.Open
        Check1.Value = 1
End If
End Sub

Private Sub Timer1_Timer()
Timer1.Interval = 10000
If Check1.Value = 1 Then
                Qry0 = "personallog_monitoring(0)"
            PrjSysGrid.GetGrid_Persensi dxDBGrid1, CLM0, False, False, BND0, True, Qry0, "TerminalID", False
            dxDBGrid1.Dataset.Open
            'MsgBox "ok"
End If
End Sub

