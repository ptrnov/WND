VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{C8CF160E-7278-4354-8071-850013B36892}#1.0#0"; "vsrpt8.ocx"
Begin VB.Form FRM_PRINVIEW 
   Caption         =   "Form1"
   ClientHeight    =   9195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11055
   Icon            =   "FRM_PRINVIEW.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9195
   ScaleWidth      =   11055
   WindowState     =   2  'Maximized
   Begin VSPrinter8LibCtl.VSPrinter vp 
      Height          =   9015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      _cx             =   18653
      _cy             =   15901
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoRTF         =   -1  'True
      Preview         =   -1  'True
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   1
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   1440
      MarginTop       =   1440
      MarginRight     =   1440
      MarginBottom    =   1440
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   51.9886363636364
      ZoomMode        =   3
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   14737632
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   3
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
      AutoLinkNavigate=   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
   End
   Begin VSReport8LibCtl.VSReport vsr 
      Left            =   120
      Top             =   240
      _rv             =   800
      ReportName      =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OnOpen          =   ""
      OnClose         =   ""
      OnNoData        =   ""
      OnPage          =   ""
      OnError         =   ""
      MaxPages        =   0
      DoEvents        =   -1  'True
      BeginProperty Layout {D853A4F1-D032-4508-909F-18F074BD547A} 
         Width           =   0
         MarginLeft      =   1440
         MarginTop       =   1440
         MarginRight     =   1440
         MarginBottom    =   1440
         Columns         =   1
         ColumnLayout    =   0
         Orientation     =   0
         PageHeader      =   0
         PageFooter      =   0
         PictureAlign    =   7
         PictureShow     =   1
         PaperSize       =   0
      EndProperty
      BeginProperty DataSource {D1359088-0913-44EA-AE50-6A7CD77D4C50} 
         ConnectionString=   ""
         RecordSource    =   ""
         Filter          =   ""
         MaxRecords      =   0
      EndProperty
      GroupCount      =   0
      SectionCount    =   5
      BeginProperty Section0 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Detail"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section1 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section2 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section3 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Page Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section4 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Page Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      FieldCount      =   0
   End
End
Attribute VB_Name = "FRM_PRINVIEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Public RptNm As String

Const DataFileName = "\Report\RPT\printView.xml"

Private Sub cmbReport_Click()
    ' make sure we have a report selected
    'If cmbReport.ListIndex < 0 Then
    '    MsgBox "Please select a report first."
    '    Exit Sub
    'End If
    
    ' load report from XML file
    'vsr.Load App.Path & DataFileName, cmbReport
     '
    ' pass focus to vsprinter control
    'vp.SetFocus
    
    ' render it
    'Command1_Click
End Sub

Private Sub Command1_Click()
'MsgBox RptNm
  ' make sure we have a report selected
    'If cmbReport.ListIndex < 0 Then
    '    MsgBox "Please select a report first."
    '    Exit Sub
    'End If
    vsr.Clear
    vsr.LOAD App.Path & DataFileName, RptNm
    'vp.SetFocus
    ' no reentrancy
    If vsr.IsBusy Then Exit Sub
    
    ' prepare to benchmark
    Dim t
    t = Timer
    MousePointer = 11
    
    ' render report
    On Error Resume Next
    vsr.Render vp
    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & vbCrLf & Err.Description
    End If
    
    ' done, tell user how long it took
    MousePointer = 0
    vp.NavBarText = "Done in " & Format(Timer - t, "#.##") & " seconds"
End Sub

Private Sub Form_Load()
 lR = SetTopMostWindow(Me.hwnd, True)
 Me.Caption = RptNm
    Dim i%, iCnt%, s$
    'Dim sConn$, sRS$
   ' Dim rs As New ADODB.Recordset
Dim StrDSN As String
    'sConn = "DRIVER={MySQL ODBC 5.1 Driver};Persist Security Info=False;SERVER=127.0.0.1;UID=root;PWD=;DATABASE=payroll;"
    'sRS = "RptBukuBesar('2012/2013',1101.001,'2012-03-01','2012-03-31')"
    'rs.Open "RptBukuBesar('2012/2013',1101.001,'2012-03-01','2012-03-31')", sConn, adOpenKeyset
    'If rs.State <> 0 Then
    '    MsgBox "connect"
    'End If
    'Set DsnConn = New ADODB.Connection
    'MsgBox GetStrDSN
    'StrDSN = "UID=Admin;PWD=emmys;DSN=" & GetStrDSN
     'vsr.DataSource.ConnectionString = "DSN=payrolsys;SERVER=localhost;UID=root;PWD=Sp1d3rm4n4;DATABASE=payroll;PORT=3306"
    ' MousePointer = 13
    '  vsr.LOAD App.Path & "\Report\RPT\printView.xml", "Persensi_Session_Closing"
     'vsr.ReportName = "Persensi_Session_Closing"
    ' vsr.Render vp
     
    
    'MousePointer = 0
    'rs.Open "SELECT TT_GRP_ID,TT_TYP,TT_SDATE,TT_EDATE,RULE_IN,RULE_OUT FROM tab_timetable_dtl", StrDSN, adOpenDynamic
    'StrDSN = "UID=Admin;PWD=emmys;DSN=BaGL-Maret-Des"
    'DsnConn.Open StrDSN
    'If DsnConn.State = 1 Then
    'MsgBox "connect"
    'Else
    'MsgBox "not connect"
    'End If
    'vsr.DataSource.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=LocalPayroll"
    'vsr.DataSource.RecordSource = "SELECT TT_GRP_ID,TT_TYP,TT_SDATE,TT_EDATE,RULE_IN,RULE_OUT FROM tab_timetable_dtl"
   ' vsr.DataSource
    'vsr.DataSource.RecordSource
   
    
    'vsr.Render vp
    ' count how many reports are in our definition file
    's = App.Path & DataFileName
    'iCnt = vsr.GetReportInfo(s, vsrRICount)
    
    ' populate list box
    'cmbReport.Clear
    'For i = 0 To iCnt - 1
    '    cmbReport.AddItem vsr.GetReportInfo(s, vsrRIName, i)
    'Next
    Command1_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    ' pressing any key cancels the report
    If vsr.IsBusy Then
        vsr.Cancel = True
    End If
End Sub

Private Sub Form_Resize()
    With vp
        On Error Resume Next
        .Move .left, .top, ScaleWidth - 2 * .left, ScaleHeight - .top - .left
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
 lR = SetTopMostWindow(Me.hwnd, True)
End Sub

Private Sub vsr_OnClose()
    Debug.Print "vsr_OnClose "; vsr.Cancel
End Sub

Private Sub vsr_OnError(ByVal Number As Long, ByVal Description As String, Handled As Boolean)
    Debug.Print "vsr_OnError "; Number; Description
'    Handled = True
End Sub

Private Sub vsr_OnNoData()
    Debug.Print "vsr_OnNoData"
End Sub

Private Sub vsr_OnOpen()
    Dim sDateRun, eDateRun, sDateRunP, eDateRunP, StrqryRslt, StrqryWhr, Strqry0, Strqry1, Strqry2, Strqry3, Strqry4, Strqry5 As String
    Dim KarId, DeptID, CabId, GrpId, MenuRpt, KarIdP, DeptIDP, CabIdP, GrpIdP, MenuRptP As String
    Dim AryRpt As String
    Debug.Print "vsr_OnOpen"
    
    '---------------------------
    '-- DEFAULT VARIABLE NULL --
    '---------------------------
    StrqryRslt = ""
    StrqryWhr = ""
    Strqry0 = ""
    Strqry1 = ""
    Strqry2 = ""
    Strqry3 = ""
    Strqry4 = ""
    Strqry5 = ""
    KarId = ""
    DeptID = ""
    CabId = ""
    GrpId = ""
    MenuRpt = ""
    KarIdP = ""
    DeptIDP = ""
    CabIdP = ""
    GrpIdP = ""
    MenuRptP = ""
    '=======================
    '------ PRESENSI -------
    '--FRM_ABSENSI_REPORT --
    '=======================
    sDateRun = Format(FRM_ABSENSI_REPORT.RptTgl(0).Value, "yyyy-mm-dd")
    eDateRun = Format(FRM_ABSENSI_REPORT.RptTgl(1).Value, "yyyy-mm-dd")
    DeptID = FRM_ABSENSI_REPORT.CmbRpt(0).Text
    KarId = FRM_ABSENSI_REPORT.CmbRpt(1).Text
    GrpId = FRM_ABSENSI_REPORT.CmbRpt(2).Text
    CabId = FRM_ABSENSI_REPORT.CmbRpt(3).Text
    MenuRpt = FRM_ABSENSI_REPORT.CmbRpt(4).Text
    
    '=======================
    '------ PAYROLL -------
    '--FRM_PAYROLL_REPORT --
    '=======================
    sDateRunP = Format(FRM_PAYROLL_REPORT.RptTgl(0).Value, "yyyy-mm-dd")
    eDateRunP = Format(FRM_PAYROLL_REPORT.RptTgl(1).Value, "yyyy-mm-dd")
    DeptIDP = FRM_PAYROLL_REPORT.CmbRpt(0).Text
    KarIdP = FRM_PAYROLL_REPORT.CmbRpt(1).Text
    GrpIdP = FRM_PAYROLL_REPORT.CmbRpt(2).Text
    CabIdP = FRM_PAYROLL_REPORT.CmbRpt(3).Text
    MenuRptP = FRM_PAYROLL_REPORT.CmbRpt(4).Text
      
    '--- END VARIABLE INPUT -----
      
         
    '======================
    '------ PAYROLL -------
    '-- PAYROLL_PRESENSI --
    '======================
    If vsr.ReportName = "PAYROLL_PRESENSI" Then
        AryRpt = ""
        AryRpt = "'Payroll_Presensi" & _
                 "=" & Format(sDateRunP, "yyyy-mm-dd") & _
                 "=" & Format(eDateRunP, "yyyy-mm-dd") & _
                 "=" & DeptIDP & _
                 "=" & KarIdP & _
                 "=" & GrpIdP & _
                 "=" & CabIdP & "'"
        '-- RENDER ----
        vsr.DataSource.ConnectionString = StrCon
        vsr.DataSource.RecordSource = "Vsr_Report(" & AryRpt & ")"
   End If
    
    '====================
    '------ PAYROLL -----
    '-- PAYROLL_WEEKLY --
    '====================
    If vsr.ReportName = "PAYROLL_WEEKLY" Then
        AryRpt = ""
        AryRpt = "'Payroll_Weekly" & _
                 "=" & Format(sDateRunP, "yyyy-mm-dd") & _
                 "=" & Format(eDateRunP, "yyyy-mm-dd") & _
                 "=" & DeptIDP & _
                 "=" & KarIdP & _
                 "=" & GrpIdP & _
                 "=" & CabIdP & "'"
        '-- RENDER ----
        vsr.DataSource.ConnectionString = StrCon
        vsr.DataSource.RecordSource = "Vsr_Report(" & AryRpt & ")"
   End If
    
    '=====================
    '------ PAYROLL ------
    '-- PAYROLL_MONTHLY --
    '=====================
    If vsr.ReportName = "PAYROLL_MONTHLY" Then
        AryRpt = ""
        AryRpt = "'Payroll_Monthly" & _
                 "=" & Format(sDateRunP, "yyyy-mm-dd") & _
                 "=" & Format(eDateRunP, "yyyy-mm-dd") & _
                 "=" & DeptIDP & _
                 "=" & KarIdP & _
                 "=" & GrpIdP & _
                 "=" & CabIdP & "'"
        '-- RENDER ----
        vsr.DataSource.ConnectionString = StrCon
        vsr.DataSource.RecordSource = "Vsr_Report(" & AryRpt & ")"
   End If
   
    '=========================
    '------ PAYROLL ----------
    '-- PAYROLL_MONTHLY_POT --
    '=========================
    If vsr.ReportName = "PAYROLL_MONTHLY_POT" Then
        AryRpt = ""
        AryRpt = "'Payroll_Monthly_Pot" & _
                 "=" & Format(sDateRunP, "yyyy-mm-dd") & _
                 "=" & Format(eDateRunP, "yyyy-mm-dd") & _
                 "=" & DeptIDP & _
                 "=" & KarIdP & _
                 "=" & GrpIdP & _
                 "=" & CabIdP & "'"
        '-- RENDER ----
        vsr.DataSource.ConnectionString = StrCon
        vsr.DataSource.RecordSource = "Vsr_Report(" & AryRpt & ")"
   End If
    '=========================
    '------ PAYROLL ----------
    '-- PAYROLL_SALARY LIST --
    '=========================
    If vsr.ReportName = "PAYROLL_SALARY" Then
        AryRpt = ""
        AryRpt = "'Payroll_Salary" & _
                 "=" & Format(sDateRunP, "yyyy-mm-dd") & _
                 "=" & Format(eDateRunP, "yyyy-mm-dd") & _
                 "=" & DeptIDP & _
                 "=" & KarIdP & _
                 "=" & GrpIdP & _
                 "=" & CabIdP & "'"
        '-- RENDER ----
        vsr.DataSource.ConnectionString = StrCon
        vsr.DataSource.RecordSource = "Vsr_Report(" & AryRpt & ")"
   End If
    '===========================
    '------ PAYROLL ------------
    '-- PAYROLL_SLIPGAJI_WEEK --
    '===========================
    If vsr.ReportName = "PAYROLL_SLIPGAJI_WEEK" Then
        AryRpt = ""
        AryRpt = "'Payroll_SlipGaji_Week" & _
                 "=" & Format(sDateRunP, "yyyy-mm-dd") & _
                 "=" & Format(eDateRunP, "yyyy-mm-dd") & _
                 "=" & DeptIDP & _
                 "=" & KarIdP & _
                 "=" & GrpIdP & _
                 "=" & CabIdP & "'"
        '-- RENDER ----
        vsr.DataSource.ConnectionString = StrCon
        vsr.DataSource.RecordSource = "Vsr_Report(" & AryRpt & ")"
   End If
   
   
    '====================
    '----- PRESENSI -----
    '-- PRESENSI_DAILY --
    '====================
    If vsr.ReportName = "PRESENSI_DAILY" Then
        AryRpt = ""
        AryRpt = "'DailyPresensi" & _
                 "=" & Format(sDateRun, "yyyy-mm-dd") & _
                 "=" & Format(eDateRun, "yyyy-mm-dd") & _
                 "=" & DeptID & _
                 "=" & KarId & _
                 "=" & GrpId & _
                 "=" & CabId & "'"
        '-- RENDER ----
        vsr.DataSource.ConnectionString = StrCon
        vsr.DataSource.RecordSource = "Vsr_Report(" & AryRpt & ")"
   End If
      
    '==================
    '-- DATA_EMPLOYE --
    '==================
    If vsr.ReportName = "DATA_EMPLOYES" Then
        AryRpt = ""
        AryRpt = "'DataEmploye" & _
                 "=" & Format(sDateRun, "yyyy-mm-dd") & _
                 "=" & Format(eDateRun, "yyyy-mm-dd") & _
                 "=" & DeptID & _
                 "=" & KarId & _
                 "=" & GrpId & _
                 "=" & CabId & "'"
        '-- RENDER ----
        vsr.DataSource.ConnectionString = StrCon
        vsr.DataSource.RecordSource = "Vsr_Report(" & AryRpt & ")"
   End If
   
    '====================
    '-- DATA_EXCEPTION --
    '====================
    If vsr.ReportName = "DATA_EXCEPTION" Then
        AryRpt = ""
        AryRpt = "'DataException" & _
                 "=" & Format(sDateRun, "yyyy-mm-dd") & _
                 "=" & Format(eDateRun, "yyyy-mm-dd") & _
                 "=" & DeptID & _
                 "=" & KarId & _
                 "=" & GrpId & _
                 "=" & CabId & "'"
        '-- RENDER ----
        vsr.DataSource.ConnectionString = StrCon
        vsr.DataSource.RecordSource = "Vsr_Report(" & AryRpt & ")"
   End If
    
    '=================
    '-- DATA_FINGER --
    '=================
    If vsr.ReportName = "DATA_FINGER" Then
        AryRpt = ""
        AryRpt = "'DataFinger" & _
                 "=" & Format(sDateRun, "yyyy-mm-dd") & _
                 "=" & Format(eDateRun, "yyyy-mm-dd") & _
                 "=" & DeptID & _
                 "=" & KarId & _
                 "=" & GrpId & _
                 "=" & CabId & "'"
        '-- RENDER ----
        vsr.DataSource.ConnectionString = StrCon
        vsr.DataSource.RecordSource = "Vsr_Report(" & AryRpt & ")"
   End If

    
End Sub

Private Sub vsr_OnPage()
    Debug.Print "vsr_OnPage"
End Sub


