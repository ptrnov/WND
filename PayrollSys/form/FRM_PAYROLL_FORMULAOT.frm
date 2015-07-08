VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Begin VB.Form FRM_PAYROLL_FORMULAOT 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FORMULA PRSENSI -  PAYROLL"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8325
      _Version        =   851970
      _ExtentX        =   14684
      _ExtentY        =   13150
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
         Height          =   6615
         Left            =   0
         OleObjectBlob   =   "FRM_PAYROLL_FORMULAOT.frx":0000
         TabIndex        =   1
         Top             =   720
         Width           =   8295
      End
   End
End
Attribute VB_Name = "FRM_PAYROLL_FORMULAOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CLM0, CLM1, BND0, BND1 As Variant
Dim Qry0 As String, Qry1 As String

Private Sub Form_Load()
lR = SetTopMostWindow(Me.hwnd, True)
Main
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
            Qry0 = "SELECT FOT_ID,TT_GRP_ID,FOT_NM,FOT_JAM1,FOT_JAM2,FOT_PERSEN FROM presensi_formula"
            PrjSysGrid.GetGrid_Persensi dxDBGrid1, CLM0, True, True, BND0, False, Qry0, "LBR_ID", False
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
    BND0 = Array("OVERTIME PROPERTIES", "OVERTIME VALUE")
    CLM0 = Array(Array("Clm0", "GROUP", "TT_GRP_ID", gedLookupEdit, 0, 0, 100, 0, 2, 0, 0), _
                Array("Clm1", "OT HOUR", "FOT_NM", gedTextEdit, 0, 0, 80, 0, 2, 0, 0), _
                Array("Clm2", "START OT TIME", "FOT_JAM1", gedDateEdit, 0, 0, 103, 0, 2, 0, 1), _
                Array("Clm3", "END OT TIME", "FOT_JAM2", gedDateEdit, 0, 0, 100, 0, 2, 0, 1), _
                Array("Clm4", "PERSENTASE OF DAY", "FOT_PERSEN", gedTextEdit, 0, 1, 150, 0, 1, 0, 0))
                '--------0---------1--------2-----------3--------4--5--6---7--8--9--10
                
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



