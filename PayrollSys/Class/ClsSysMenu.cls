VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSysMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public ControlAbout As CommandBarControl
Dim RibbonBar As RibbonBar
Dim HeaderTab As RibbonTab
Dim FirstTab As SelectedControls
Dim GroupFile As RibbonGroup
Dim ControlBulatan As CommandBarPopup
Dim ControlBulatan1 As CommandBarsContextMenus
Dim rsRec1 As ADODB.Recordset
Dim rsRec2 As ADODB.Recordset
Dim qry As String

Public Sub CreateRiboonBar(CmnBar As CommandBars, IdUser As String)
On Error Resume Next
    LoadIcons CmnBar ' Load Icon
    
    '==================================
    '====== FIRST SET AddRibbonBar ====
    '===========  ptr.nov   ===========
    '==================================
    Set RibbonBar = CmnBar.AddRibbonBar("The Ribbon")
    RibbonBar.ShowQuickAccess = False ' Menghilangkan button
   
    '========================================
    '========= BUTTON BULATAN START =========
    '============  ptr.nov   ================
    '========================================
    Set ControlBulatan = RibbonBar.AddSystemButton() ' Bunderan
    ControlBulatan.IconId = 1
     CreateRiboonBarOption
   
    
    '==================================
    '======= POSISI TAB MANUALY =======
    '===========  ptr.nov   ===========
    '==================================
    'Set ControlAbout = RibbonBar.Controls.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 2, "&About", False, False)
    'ControlAbout.Flags = XtremeCommandBars.XTPControlFlags.xtpFlagRightAlign

    '==================================
    '========== TAB SETTING ===========
    '===========  ptr.nov   ===========
    '==================================
     DbMnTabRibbon IdUser
    
    '==================================
    '======= SELECT TAB MANUALY =======
    '===========  ptr.nov   ===========
    '==================================
    Set HeaderTab = RibbonBar.FindTab(1)
    HeaderTab.Selected = True
    
End Sub

Private Sub DbMnTabRibbon(UsrID As String)
qry = " SELECT DISTINCT B.MN_PRN as MnPrn, B.MN_ID as MnID, B.MN_NM as MnNm, C.MN_SHW as MnSh, C.MN_RD as MnRd, " & _
      " C.MN_WR as MnRw,C.MN_DEL as MnDl FROM tab_user_per C LEFT OUTER JOIN tab_user_menu B on B.MN_ID=C.MN_ID " & _
      " where  B.MN_PRN=0 and C.MN_SHW <> 0 and C.PRMS_ID='" & UsrID & "' ORDER BY B.MN_ID ASC"
Set rsRec1 = ExecuteRecordSetMySql(qry, conMain)
With rsRec1
On Error Resume Next
    If .RecordCount <> 0 Then
     .MoveFirst
     Do Until .EOF
           '===============================
           '======= SET HEADER TABs =======
           '===========  ptr.nov   ========
           '===============================
           Set HeaderTab = RibbonBar.InsertTab(.Fields("MnID").Value, .Fields("MnNm").Value)
           HeaderTab.Id = .Fields("MnID").Value
           'HeaderTab.RibbonBar.Position = xtpBarPopup
    
                '===========================
                '======= CHAILD TABs =======
                '========  ptr.nov   =======
                '===========================
                qry = " SELECT B.MN_PRN as MnPrn,B.MN_ID as MnID,B.MN_NM as MnNm, C.MN_SHW as MnSh, C.MN_RD as MnRd, " & _
                      " C.MN_WR as MnRw,C.MN_DEL as MnDl FROM tab_user_per C LEFT OUTER JOIN tab_user_menu B on B.MN_ID=C.MN_ID " & _
                      " where C.MN_SHW <> 0 and C.PRMS_ID='" & UsrID & "' AND B.MN_PRN='" & .Fields("MnId").Value & "'" & _
                      " ORDER BY B.MN_ID ASC"
                Set rsRec2 = ExecuteRecordSetMySql(qry, conMain)
                With rsRec2
                    If Not .EOF Then ' = da lebih aman dari error .RecordCount <> 0 Then
                         .MoveFirst
                        Do Until .EOF
                             '---------- SHOW CHAILD ---------
                             '.Controls.Add xtpControlButton, sysprof_rec2.Fields("MnId").Value, sysprof_rec2.Fields("MnNm").Value ' 1= id menu
                            ChldTabRibbon rsRec1.Fields("MnID").Value, .Fields("MnId").Value, .Fields("MnNm").Value
                            
                        .MoveNext
                        Loop
                    End If
                End With
                Set rsRec2 = Nothing
    .MoveNext
    Loop
 End If
End With
Set rsRec1 = Nothing
End Sub

Private Sub ChldTabRibbon(IndexHeaderTab As Integer, IdChld As Integer, NmChld As String)
'===========================
'======= CHAILD TABs =======
'========  ptr.nov   =======
'===========================
    Set GroupFile = HeaderTab.Groups.AddGroup("", 0)
    GroupFile.Add XtremeCommandBars.XTPControlType.xtpControlButton, IdChld, NmChld, False, False
    'GroupFile.Add XtremeCommandBars.XTPControlType.xtpControlButton, IdChld, NmChld, False, False
End Sub

Public Sub CreateRiboonBarOption()
'============================
'======= Buletan FILE =======
'========  ptr.nov   ========
'============================
Set rsRec1 = ExecuteRecordSetMySql("SELECT * FROM tab_user_menu where MN_PRN=30", conMain)
With rsRec1
        If .RecordCount <> 0 Then
         .MoveFirst
         Do Until .EOF
            ControlBulatan.CommandBar.Controls.Add XtremeCommandBars.XTPControlType.xtpControlButton, .Fields("MN_ID").Value, .Fields("MN_NM").Value, False, False
         .MoveNext
         Loop
        End If
End With
Set rsRec1 = Nothing
End Sub

Private Sub LoadIcons(CmnBar As CommandBars)
On Error GoTo ErrorLabel
    
    CmnBar.Options.UseSharedImageList = False
    '===Home====
    'status
     CmnBar.Icons.LoadBitmap ImagePath("MNU_HOME_STATUS"), Array(100), XtremeCommandBars.XTPImageState.xtpImageNormal
    'news
     'CmnBar.Icons.LoadBitmap ImagePath("MNU_HOME_NEWS"), Array(101), XtremeCommandBars.XTPImageState.xtpImageNormal
    'berita acara
     'CmnBar.Icons.LoadBitmap ImagePath("MNU_HOME_BERITA_ACARA"), Array(102), XtremeCommandBars.XTPImageState.xtpImageNormal
    'proses presensi
     CmnBar.Icons.LoadBitmap ImagePath("MNU_PRESENSI_PROSES"), Array(101), XtremeCommandBars.XTPImageState.xtpImageNormal
    'IMPORT
     CmnBar.Icons.LoadBitmap ImagePath("MNU_PRESENSI_IMPORT"), Array(102), XtremeCommandBars.XTPImageState.xtpImageNormal
    'presensi
     CmnBar.Icons.LoadBitmap ImagePath("MNU_PRESENSI_KARYAWAN"), Array(103), XtremeCommandBars.XTPImageState.xtpImageNormal
    'pAYROLL
     CmnBar.Icons.LoadBitmap ImagePath("MNU_HOME_PAYROLL"), Array(104), XtremeCommandBars.XTPImageState.xtpImageNormal
  '===IT====
    CmnBar.Icons.LoadBitmap ImagePath("MNU_IT_FINGERMACHINE"), Array(700), XtremeCommandBars.XTPImageState.xtpImageNormal
   CmnBar.Icons.LoadBitmap ImagePath("MNU_IT_EMPFINGER"), Array(701), XtremeCommandBars.XTPImageState.xtpImageNormal
   
 '===KARYAWAN===
    CmnBar.Icons.LoadBitmap ImagePath("MNU_HRD_KARYAWAN"), Array(600), XtremeCommandBars.XTPImageState.xtpImageNormal
    CmnBar.Icons.LoadBitmap ImagePath("MNU_KARYAWAN_AGAMA"), Array(200), XtremeCommandBars.XTPImageState.xtpImageNormal
    CmnBar.Icons.LoadBitmap ImagePath("MNU_KARYAWAN_BAGIAN"), Array(201), XtremeCommandBars.XTPImageState.xtpImageNormal
    CmnBar.Icons.LoadBitmap ImagePath("MNU_KARYAWAN_DEPARTEMEN"), Array(202), XtremeCommandBars.XTPImageState.xtpImageNormal
    CmnBar.Icons.LoadBitmap ImagePath("MNU_KARYAWAN_JABATAN"), Array(203), XtremeCommandBars.XTPImageState.xtpImageNormal
    CmnBar.Icons.LoadBitmap ImagePath("MNU_KARYAWAN_HARI_KERJA"), Array(204), XtremeCommandBars.XTPImageState.xtpImageNormal
    CmnBar.Icons.LoadBitmap ImagePath("MNU_KARYAWAN_JENIS_JAM_KERJA"), Array(205), XtremeCommandBars.XTPImageState.xtpImageNormal
    CmnBar.Icons.LoadBitmap ImagePath("MNU_KARYAWAN_JAM_KERJA"), Array(206), XtremeCommandBars.XTPImageState.xtpImageNormal
    CmnBar.Icons.LoadBitmap ImagePath("MNU_KARYAWAN_PENDIDIKAN"), Array(207), XtremeCommandBars.XTPImageState.xtpImageNormal
    CmnBar.Icons.LoadBitmap ImagePath("MNU_KARYAWAN_DATA_KARYAWAN"), Array(208), XtremeCommandBars.XTPImageState.xtpImageNormal
    CmnBar.Icons.LoadBitmap ImagePath("MNU_KARYAWAN_KONTRAK_KERJA"), Array(209), XtremeCommandBars.XTPImageState.xtpImageNormal
    
    'CmnBar.Icons.LoadBitmap ImagePath("MNU_PRESENSI_ABSEN"), Array(300), XtremeCommandBars.XTPImageState.xtpImageNormal
    'CmnBar.Icons.LoadBitmap ImagePath("MNU_PRESENSI_IMPORT"), Array(301), XtremeCommandBars.XTPImageState.xtpImageNormal
    'CmnBar.Icons.LoadBitmap ImagePath("MNU_PRESENSI_KARYAWAN"), Array(302), XtremeCommandBars.XTPImageState.xtpImageNormal
    'CmnBar.Icons.LoadBitmap ImagePath("MNU_PRESENSI_MANUAL"), Array(303), XtremeCommandBars.XTPImageState.xtpImageNormal
    'CmnBar.Icons.LoadBitmap ImagePath("MNU_PRESENSI_PROSES"), Array(304), XtremeCommandBars.XTPImageState.xtpImageNormal
    '===PAYROLL===
    CmnBar.Icons.LoadBitmap ImagePath("MNU_GAJI_KOMPONEN"), Array(300), XtremeCommandBars.XTPImageState.xtpImageNormal
    CmnBar.Icons.LoadBitmap ImagePath("MNU_GAJI_PPH"), Array(301), XtremeCommandBars.XTPImageState.xtpImageNormal
    CmnBar.Icons.LoadBitmap ImagePath("MNU_GAJI_PTKP"), Array(302), XtremeCommandBars.XTPImageState.xtpImageNormal
    CmnBar.Icons.LoadBitmap ImagePath("MNU_GAJI_JAMSOSTEK"), Array(303), XtremeCommandBars.XTPImageState.xtpImageNormal
    CmnBar.Icons.LoadBitmap ImagePath("MNU_GAJI_SETTING"), Array(304), XtremeCommandBars.XTPImageState.xtpImageNormal
    CmnBar.Icons.LoadBitmap ImagePath("MNU_GAJI_PROSES"), Array(305), XtremeCommandBars.XTPImageState.xtpImageNormal
    '===LAPORAN===
    CmnBar.Icons.LoadBitmap ImagePath("MNU_LAPORAN_KARYAWAN"), Array(400), XtremeCommandBars.XTPImageState.xtpImageNormal
    CmnBar.Icons.LoadBitmap ImagePath("MNU_LAPORAN_GAJI"), Array(401), XtremeCommandBars.XTPImageState.xtpImageNormal
    CmnBar.Icons.LoadBitmap ImagePath("MNU_LAPORAN_GAJI_MINGGUAN"), Array(402), XtremeCommandBars.XTPImageState.xtpImageNormal
    CmnBar.Icons.LoadBitmap ImagePath("MNU_LAPORAN_KONTRAK"), Array(403), XtremeCommandBars.XTPImageState.xtpImageNormal
    CmnBar.Icons.LoadBitmap ImagePath("MNU_LAPORAN_KONTRAK_TEMPO"), Array(404), XtremeCommandBars.XTPImageState.xtpImageNormal
    CmnBar.Icons.LoadBitmap ImagePath("MNU_LAPORAN_PRESENSI_KARYAWAN"), Array(405), XtremeCommandBars.XTPImageState.xtpImageNormal
    CmnBar.Icons.LoadBitmap ImagePath("MNU_LAPORAN_FORM_1721"), Array(406), XtremeCommandBars.XTPImageState.xtpImageNormal
    CmnBar.Icons.LoadBitmap ImagePath("MNU_LAPORAN_JAMSOSTEK"), Array(407), XtremeCommandBars.XTPImageState.xtpImageNormal
    '===TOOLs===
    CmnBar.Icons.LoadBitmap ImagePath("MNU_TOOLS_CALC"), Array(500), XtremeCommandBars.XTPImageState.xtpImageNormal
    '===SETUP USER===
    CmnBar.Icons.LoadBitmap ImagePath("MNU_SETUP_OPTIONS"), Array(602), XtremeCommandBars.XTPImageState.xtpImageNormal
    CmnBar.Icons.LoadBitmap ImagePath("MNU_SETUP_PASS"), Array(603), XtremeCommandBars.XTPImageState.xtpImageNormal
    CmnBar.Icons.LoadBitmap ImagePath("MNU_SETUP_PROFILE"), Array(604), XtremeCommandBars.XTPImageState.xtpImageNormal
    
    '===User Menu====e
    CmnBar.Icons.LoadBitmap ImagePath("MNU_USER"), Array(600, 601), XtremeCommandBars.XTPImageState.xtpImageNormal
     
    Dim ToolTipContext As ToolTipContext
    Set ToolTipContext = CmnBar.ToolTipContext
    ToolTipContext.Style = xtpToolTipOffice2007
    ToolTipContext.ShowTitleAndDescription True, xtpToolTipIconNone
    'ToolTipContext.ShowImage True, IMAGEBASE
    ToolTipContext.SetMargin 2, 2, 2, 2
    ToolTipContext.MaxTipWidth = 180
    ToolTipContext.ShowShadow = True
Exit Sub

Exit Sub
ErrorLabel:
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"
End Sub

Public Sub LoginCaption(IdUser As String)
'==================================
    '======= POSISI TAB MANUALY =======
    '===========  ptr.nov   ===========
    '==================================
    Set ControlAbout = RibbonBar.Controls.Add(XtremeCommandBars.XTPControlType.xtpControlLabel, 1000, "Wellcome:   ", False, False)
    ControlAbout.Flags = XtremeCommandBars.XTPControlFlags.xtpFlagRightAlign
    Set ControlAbout = RibbonBar.Controls.Add(XtremeCommandBars.XTPControlType.xtpControlLabel, 1001, IdUser, False, False)
    ControlAbout.Flags = XtremeCommandBars.XTPControlFlags.xtpFlagRightAlign
    'Set ControlAbout = RibbonBar.Controls.Add(XtremeCommandBars.XTPControlType.xtpControlLabel, 200, "Corp:   ", False, False)
    'ControlAbout.Flags = XtremeCommandBars.XTPControlFlags.xtpFlagRightAlign
    'Set ControlAbout = RibbonBar.Controls.Add(XtremeCommandBars.XTPControlType.xtpControlLabel, 201, IdCorp, False, False)
    'ControlAbout.Flags = XtremeCommandBars.XTPControlFlags.xtpFlagRightAlign
End Sub

Public Sub UserAktifLogin(ByRef CIdUser As String)
  '==================================
    '======= SELECT TAB MANUALY =======
    '===========  ptr.nov   ===========
    '==================================
    Set ControlAbout = RibbonBar.FindControl("", 201, "", "")
    CIdUser = ControlAbout.Caption
End Sub


















