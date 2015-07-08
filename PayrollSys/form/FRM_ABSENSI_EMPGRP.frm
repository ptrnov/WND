VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Begin VB.Form FRM_ABSENSI_EMPGRP 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Employe Group Time Table"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      _Version        =   851970
      _ExtentX        =   14208
      _ExtentY        =   13785
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton Btn 
         Height          =   375
         Left            =   6960
         TabIndex        =   6
         Top             =   240
         Width           =   975
         _Version        =   851970
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Edit"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox CmbEmpGrp 
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1935
         _Version        =   851970
         _ExtentX        =   3413
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "ComboBox1"
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
         Height          =   6975
         Left            =   0
         OleObjectBlob   =   "FRM_ABSENSI_EMPGRP.frx":0000
         TabIndex        =   1
         Top             =   720
         Width           =   8055
      End
      Begin XtremeSuiteControls.ComboBox CmbEmpGrp 
         Height          =   315
         Index           =   1
         Left            =   2040
         TabIndex        =   5
         Top             =   360
         Width           =   2175
         _Version        =   851970
         _ExtentX        =   3836
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "ComboBox1"
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
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Dept"
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
         Left            =   2040
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FRM_ABSENSI_EMPGRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CLM0, CLM1, BND0, BND1 As Variant
Dim Qry0 As String, Qry1 As String

Private Sub Btn_Click()
Dim i As Integer
'==========
'---Edit---
'==========
    If Btn.Caption = "Edit" Then
        For i = 2 To 3
            dxDBGrid1.Columns(i).DisableEditor = False
            dxDBGrid1.Columns(i).Color = &HFFFFFF
        Next i
        Btn.Caption = "Accept"
    Else
        For i = 2 To 3
            dxDBGrid1.Dataset.Refresh
            dxDBGrid1.Columns(i).DisableEditor = True
            dxDBGrid1.Columns(i).Color = &HCBFAFE
        Next i
        Btn.Caption = "Edit"
    End If
End Sub

Private Sub CmbEmpGrp_Click(Index As Integer)
If Index = 0 Or Index = 1 Then
    GridFill 0
End If
End Sub

Private Sub cmdFilter_Click()
GridFill 0
End Sub

Private Sub dxDBGrid1_OnAddGroupColumn(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, Allow As Boolean)
If Column.Index = 2 Then
    dxDBGrid1.Columns.ColumnByName("Clm2").DisableEditor = False
End If
End Sub

Private Sub Form_Load()
lR = SetTopMostWindow(Me.hwnd, True)
Main
GridFill 0
 prjSysID.CABANG CmbEmpGrp(0)
 prjSysID.Dept CmbEmpGrp(1)
End Sub

'=====================================
'============= ptr.nov ===============
'===== PROCESS FIRST GRID LODING =====
'=====================================
Private Sub GridFill(GrdIndx As Integer)
Select Case GrdIndx
    Case Is = 0
        MenuGrid 0
           ' Qry0 = "crud_EmpGroup(1,'" & CmbEmpGrp(0).Text & "'," & CmbEmpGrp(1).Tex & ")"
            Qry0 = "crud_EmpGroup(1,'" & CmbEmpGrp(0).Text & "','" & CmbEmpGrp(1).Text & "')"
            PrjSysGrid.GetGrid_Persensi dxDBGrid1, CLM0, False, False, BND0, True, Qry0, "LBR_ID", False
            LookupClm
           ' dxDBGrid1.Dataset.Open
    End Select
End Sub

'=====================================
'============= ptr.nov ===============
'===== MENU CULUMN BND GRID LODING ===
'=====================================
Private Sub MenuGrid(TabSel As Integer)
Select Case TabSel
Case Is = 0
    BND0 = Array("EMPLOYE ", "TIME TABLE SET")
    CLM0 = Array(Array("Clm0", "Employe ID", "KAR_ID", gedTextEdit, 0, 0, 120, 1, 1, 0, 0), _
                Array("Clm1", "Employe Name", "KAR_NM", gedTextEdit, 0, 0, 198, 1, 1, 0, 0), _
                Array("Clm2", "Group Time Table", "GRP_ID", gedLookupEdit, 0, 1, 130, 1, 0, 0, 0), _
                Array("Clm3", "Level ", "LVL_ID", gedLookupEdit, 0, 1, 60, 1, 0, 0, 0))
                '--------0---------1--------2-----------3--------4--5--6---7--8--9--10
                
End Select
End Sub

Private Sub LookupClm()
With dxDBGrid1.Columns.ColumnByName("Clm3").LookupColumn
            .LookupDataset.EnableControls
            '.LookupDataset.Close
            '.Dataset.Refresh
            .LookupDatasetType = dtADODataset
            .LookupDataset.ADODataset.ConnectionString = StrCon
            .LookupDataset.ADODataset.CursorLocation = clUseClient
            .LookupDataset.ADODataset.LockType = ltReadOnly         ' Check Permmission
            .LookupDataset.ADODataset.CommandType = cmdText
            .LookupDataset.ADODataset.CommandText = " SELECT LVL_ID, LVL_NM FROM level " ' Like Join
            '.LookupDataset.Open
            .LookupKeyField = "LVL_ID"
            .LookupResultField = "LVL_NM"
            .ListColumns = "Level Name"
            .ListFieldName = "LVL_NM"
            .ListWidth = 800
            .DisplaySize = 400
End With

With dxDBGrid1.Columns.ColumnByName("Clm2").LookupColumn
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
            .ListWidth = 800
            .DisplaySize = 400
            
End With
dxDBGrid1.Dataset.Open
End Sub




