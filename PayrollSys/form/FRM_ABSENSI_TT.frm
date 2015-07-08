VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Begin VB.Form FRM_ABSENSI_TT 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "TIME TABLE"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   14400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14175
      _Version        =   851970
      _ExtentX        =   25003
      _ExtentY        =   11668
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.ComboBox ops 
         Height          =   315
         Left            =   5880
         TabIndex        =   9
         Top             =   480
         Width           =   2055
         _Version        =   851970
         _ExtentX        =   3625
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.ComboBox CmbTTGrp 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   2775
         _Version        =   851970
         _ExtentX        =   4895
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
         Height          =   5655
         Left            =   120
         OleObjectBlob   =   "FRM_ABSENSI_TT.frx":0000
         TabIndex        =   1
         Top             =   840
         Width           =   14160
      End
      Begin XtremeSuiteControls.DateTimePicker Tgl 
         Height          =   375
         Left            =   3360
         TabIndex        =   2
         Top             =   480
         Width           =   2055
         _Version        =   851970
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   68
      End
      Begin XtremeSuiteControls.PushButton cmdProcess 
         Height          =   375
         Index           =   1
         Left            =   2880
         TabIndex        =   6
         Top             =   480
         Width           =   495
         _Version        =   851970
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "<<"
         ForeColor       =   255
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton cmdProcess 
         Height          =   375
         Index           =   2
         Left            =   7920
         TabIndex        =   7
         Top             =   480
         Width           =   735
         _Version        =   851970
         _ExtentX        =   1296
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Enter"
         BackColor       =   16777215
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton cmdProcess 
         Height          =   375
         Index           =   0
         Left            =   5400
         TabIndex        =   8
         Top             =   480
         Width           =   495
         _Version        =   851970
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   ">>"
         ForeColor       =   255
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton cmdProcess 
         Height          =   375
         Index           =   3
         Left            =   12960
         TabIndex        =   11
         Top             =   360
         Width           =   1095
         _Version        =   851970
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Formula"
         BackColor       =   16777215
         Appearance      =   6
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Opreational"
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
         Left            =   5880
         TabIndex        =   10
         Top             =   240
         Width           =   1335
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
         Index           =   5
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   3360
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FRM_ABSENSI_TT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CLM0, CLM1, BND0, BND1 As Variant
Dim Qry0 As String, Qry1 As String


Private Sub cmdProcess_Click(Index As Integer)
If Index = 2 Then
    GridFill 0
End If
If Index = 3 Then
    FRM_PAYROLL_FORMULAOT.Show
End If
End Sub

Private Sub Form_Load()
ops.AddItem "NORMAL"
ops.AddItem "OVERTIME"
ops.ListIndex = 0
lR = SetTopMostWindow(Me.hwnd, True)
Main
Tgl.Value = Now
Tgl.Format = xtpPickerShortDate
prjSysID.TTGROUP CmbTTGrp
GridFill 0
End Sub

'=====================================
'============= ptr.nov ===============
'===== PROCESS FIRST GRID LODING =====
'=====================================
Private Sub GridFill(GrdIndx As Integer)
Select Case GrdIndx
    Case Is = 0
        MenuGrid 0
            'Qry0 = "Select * from timetable_normal where TT_GRP_ID=" & CmbTTGrp.Text & " AND ('" & Format(Tgl.Value, "yyyy-mm-dd") & "' BETWEEN TT_SDATE AND TT_EDATE) AND (DAYOFWEEK('" & Format(Tgl.Value, "yyyy-mm-dd") & "') BETWEEN TT_SDAYS AND TT_EDAYS)"
            If ops.Text = "NORMAL" Then
                Qry0 = "crud_timetable(0," & CmbTTGrp.Text & ")"
            ElseIf ops.Text = "OVERTIME" Then
                Qry0 = "crud_timetable(1," & CmbTTGrp.Text & ")"
            End If
            PrjSysGrid.GetGrid_Persensi dxDBGrid1, CLM0, True, True, BND0, True, Qry0, "TerminalID", False
            LookupClm
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
    CLM0 = Array(Array("Clm0", "Group Time Table", "TT_GRP_ID", gedLookupEdit, 0, 0, 110, 1, 0, 0, 0), _
                Array("Clm1", "Type", "TT_TYP", gedTextEdit, 0, 0, 80, 6, 0, 0, 0), _
                Array("Clm2", "Start Date", "TT_SDATE", gedDateEdit, 0, 0, 70, 6, 0, 0, 3), _
                Array("Clm3", "End Date", "TT_EDATE", gedDateEdit, 0, 0, 70, 6, 0, 0, 3), _
                Array("Clm4", "Time In.", "RULE_IN", gedTimeEdit, 0, 1, 70, 6, 0, 0, 1), _
                Array("Clm5", "Time Out.", "RULE_OUT", gedTimeEdit, 0, 1, 70, 6, 0, 0, 1), _
                Array("Clm6", "Teme late", "RULE_TOL_IN", gedTimeEdit, 0, 1, 70, 60, 0, 1, 1), _
                Array("Clm7", "Time Early", "RULE_TOL_OUT", gedTimeEdit, 0, 1, 70, 6, 0, 0, 1), _
                Array("Clm8", "Break Out.", "RULE_BRK_OUT", gedTimeEdit, 0, 1, 70, 6, 0, 0, 1), _
                Array("Clm9", "Break In.", "RULE_TOL_IN", gedTimeEdit, 0, 1, 70, 6, 0, 0, 1), _
                Array("Clm10", "Duration OT", "RULE_DRT_OT_DPN", gedTimeEdit, 0, 6, 70, 6, 0, 0, 1), _
                Array("Clm11", "Duration OT", "RULE_DRT_OT_BLK", gedTimeEdit, 0, 6, 70, 0, 0, 0, 1), _
                Array("Clm12", "Hari Aktif", "TT_NOTE", gedTextEdit, 0, 1, 110, 6, 0, 0, 0))
                '--------0---------1--------2-----------3--------4--5--6---7--8--9--10
                'TT_ID, TT_GRP_ID,TT_TYP,TT_TYP_KTG,TT_THN,TT_SDAYS,TT_EDAYS,TT_SDATE,TT_EDATE,TT_NOTE,TT_UPDT,TT_ACTIVE,RULE_IN,
                'RULE_OUT , RULE_TOL_IN, RULE_TOL_OUT, RULE_BRK_OUT, RULE_BRK_IN, RULE_DRT_OT_DPN, RULE_DRT_OT_BLK, RULE_DURATION
End Select
End Sub

Private Sub LookupClm()
With dxDBGrid1.Columns.ColumnByName("Clm0").LookupColumn
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

dxDBGrid1.Dataset.Open
End Sub
