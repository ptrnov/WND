VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "COC28D~1.OCX"
Begin VB.Form PERMISSION 
   Caption         =   "PERMISSION"
   ClientHeight    =   10230
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17115
   ControlBox      =   0   'False
   DrawMode        =   15  'Merge Pen Not
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10230
   ScaleWidth      =   17115
   WindowState     =   2  'Maximized
   Begin XtremeTaskPanel.TaskPanel TaskPanelStupGl 
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   495
      _Version        =   851970
      _ExtentX        =   873
      _ExtentY        =   1508
      _StockProps     =   64
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   6855
      Left            =   840
      TabIndex        =   0
      Top             =   1800
      Width           =   14175
      _Version        =   851970
      _ExtentX        =   25003
      _ExtentY        =   12091
      _StockProps     =   68
      Appearance      =   2
      Color           =   64
      PaintManager.BoldSelected=   -1  'True
      PaintManager.ShowTabs=   0   'False
      PaintManager.FixedTabWidth=   150
      PaintManager.MaxTabWidth=   150
      PaintManager.MinTabWidth=   150
      PaintManager.ControlMargin=   "1,0,0,0"
      ItemCount       =   3
      SelectedItem    =   1
      Item(0).Caption =   "Cost Center"
      Item(0).ControlCount=   0
      Item(1).Caption =   "Budget"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "Label2"
      Item(2).Caption =   "Standing Jurnal"
      Item(2).ControlCount=   0
      Begin VB.Label Label2 
         Caption         =   "Label1"
         Height          =   1455
         Left            =   360
         TabIndex        =   2
         Top             =   1080
         Width           =   2295
      End
   End
End
Attribute VB_Name = "PERMISSION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GroupSubject As TaskPanelGroup
Dim GroupTab As TaskPanelGroup
Dim GroupMsg As TaskPanelGroup
Dim Item As TaskPanelGroupItem
Dim CLM1 As Variant
Dim BND As Variant
Dim qry As String



Private Sub Form_Resize()
On Error Resume Next
TaskPanelStupGl.Move 0, 0, 2535, ScaleHeight
TabControl1.Move TaskPanelStupGl.Width, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub Form_Load()
    CreateTaskPanel
    MenuGrid
    qry = "SELECT DISTINCT MN_ID,MN_PRN,USR_PASS FROM st01c ORDER BY USR_NM ASC"
    'PrjSysGrid.GetBaseGrid dxDBGridPERMISSION, CLM1, True, True, BND, False, qry, "PERMISSION_ID"
    Label2.Caption = "No"
End Sub
'=========================================== END CONFIG FORM =======================================================

Sub CreateTaskPanel()
    Set GroupSubject = TaskPanelStupGl.Groups.Add(0, "MENU AKTIF")
    GroupSubject.Expandable = False
    GroupSubject.Expanded = False
    GroupSubject.Special = True
  
    Set GroupTab = TaskPanelStupGl.Groups.Add(1, "Menu Pilihan")
        GroupTab.ToolTip = "Pilih Menu yang di inginkan"
        GroupTab.Items.Add 1, "PERMISSION", xtpTaskItemTypeLink, 1
        GroupTab.Items.Add 2, "NONE", xtpTaskItemTypeLink, 2
        GroupTab.Items.Add 3, "NONE", xtpTaskItemTypeLink, 2
        GroupTab.Special = True
    
    Set GroupMsg = TaskPanelStupGl.Groups.Add(2, "Pesan")
        GroupMsg.ToolTip = "Pesan Form Pengunaan"
        GroupMsg.Items.Add 1, "", xtpTaskItemTypeText, 4
        GroupMsg.Special = True
        
    TabControl1.SelectedItem = 0
    GroupSubject.Caption = "PERMISSION"
End Sub

Private Sub TaskPanelStupGl_ItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
Dim StrMessage1 As String
Dim StrMessage2 As String
Dim StrMessage3 As String
Dim StrSubject1 As String
Dim StrSubject2 As String
Dim StrSubject3 As String
StrMessage1 = "Form Plan !  "
StrMessage2 = "Form None ! "
StrMessage3 = "Form None ! "
StrSubject1 = "PERMISSION"
StrSubject2 = "NONE"
StrSubject3 = "NONE"

If Item.Id = 1 Then
    GroupSubject.Caption = StrSubject1
    GroupMsg.Items.Find(1).Remove
    GroupMsg.Items.Add 1, StrMessage1, xtpTaskItemTypeText, 4
    TabControl1.SelectedItem = 0
End If
If Item.Id = 2 Then
    GroupSubject.Caption = StrSubject2
    GroupMsg.Items.Find(1).Remove
    GroupMsg.Items.Add 1, StrMessage2, xtpTaskItemTypeText, 4
    TabControl1.SelectedItem = 1
End If
If Item.Id = 3 Then
    GroupSubject.Caption = StrSubject3
    GroupMsg.Items.Find(1).Remove
    GroupMsg.Items.Add 1, StrMessage3, xtpTaskItemTypeText, 4
    TabControl1.SelectedItem = 2
End If
End Sub

Private Sub MenuGrid()
    '----------
    '== BIND ==
    '----------
    BND = Array("PERMISSION", "Profile")
    '------------
    '== COLUMN ==
    '------------
    CLM1 = Array(Array("Column1", "PERMISSIONID", "PERMISSION_ID", gedTextEdit, 0, 0, 100, 1, 1, 0), _
                    Array("Column2", "Password", "USR_PASS", gedTextEdit, 0, 0, 100, 0, 0, 0), _
                    Array("Column3", "PERMISSIONName", "USR_NM", gedTextEdit, 0, 0, 150, 0, 0, 0), _
                    Array("Column4", "Disable", "USR_OFF", gedCheckEdit, 0, 1, 100, 0, 0, 0), _
                    Array("Column5", "Department", "DEP_ID", gedTextEdit, 0, 1, 100, 0, 0, 0))
    '-----------------------
    '== PROPERTIES COLUMN ==
    '-----------------------
    '(ObjectName0,Caption1,FieldName2,TypeFiled3,RowIndex4,BandIndex5,Width6,RowDisableEdit7,Alignment8,DecimalPlaces9)
    
End Sub


