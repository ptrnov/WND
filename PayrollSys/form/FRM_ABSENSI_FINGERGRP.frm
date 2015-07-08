VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Begin VB.Form FRM_ABSENSI_FINGERGRP 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FINGER GROUP EMPLOYE"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7680
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      _Version        =   851970
      _ExtentX        =   13573
      _ExtentY        =   14631
      _StockProps     =   79
      BackColor       =   8454016
      Appearance      =   1
      Begin XtremeSuiteControls.ComboBox CmbSorted 
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2055
         _Version        =   851970
         _ExtentX        =   3625
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin VB.CommandButton Command1 
         Caption         =   "test"
         Height          =   375
         Left            =   3960
         TabIndex        =   5
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin XtremeSuiteControls.PushButton Btn 
         Height          =   375
         Index           =   0
         Left            =   5400
         TabIndex        =   2
         Top             =   840
         Width           =   735
         _Version        =   851970
         _ExtentX        =   1296
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Add"
         UseVisualStyle  =   -1  'True
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
         Height          =   6975
         Left            =   0
         OleObjectBlob   =   "FRM_ABSENSI_FINGERGRP.frx":0000
         TabIndex        =   1
         Top             =   1320
         Width           =   7650
      End
      Begin XtremeSuiteControls.PushButton Btn 
         Height          =   375
         Index           =   1
         Left            =   6120
         TabIndex        =   3
         Top             =   840
         Width           =   735
         _Version        =   851970
         _ExtentX        =   1296
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Edit"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton Btn 
         Height          =   375
         Index           =   2
         Left            =   6840
         TabIndex        =   4
         Top             =   840
         Width           =   735
         _Version        =   851970
         _ExtentX        =   1296
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Delete"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton Btn 
         Height          =   375
         Index           =   3
         Left            =   3000
         TabIndex        =   6
         Top             =   840
         Width           =   735
         _Version        =   851970
         _ExtentX        =   1296
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Sorted"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox CmbSorted 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   8
         Top             =   360
         Width           =   2775
         _Version        =   851970
         _ExtentX        =   4895
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.ComboBox CmbKar 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   2775
         _Version        =   851970
         _ExtentX        =   4895
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Employe"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Dept."
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   11
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Cabang."
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   975
      End
   End
End
Attribute VB_Name = "FRM_ABSENSI_FINGERGRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CLM0, CLM1, BND0, BND1 As Variant
Dim Qry0 As String, Qry1 As String

Private Sub Btn_Click(Index As Integer)
Select Case Index

    Case 0: '-----------
            '-- SAVED --
            '-----------
            'lR = SetTopMostWindow(Me.hwnd, False)
            FRM_ABSENSI_FINGERGRP_EDITOR.Show
            FRM_ABSENSI_FINGERGRP_EDITOR.BtnProcess.Caption = "Save"
            FRM_ABSENSI_FINGERGRP_EDITOR.Cmb(0).Locked = False
            FRM_ABSENSI_FINGERGRP_EDITOR.Cmb(1).Locked = False
            'FRM_ABSENSI_FINGERGRP_EDITOR.Cmb(2).Locked = False
            'FRM_ABSENSI_FINGERGRP_EDITOR.Cmb(3).Locked = False
    Case 1:
            '------------
            '-- UPDATE --
            '------------
            dxDBGrid1_OnDblClick
            FRM_ABSENSI_FINGERGRP_EDITOR.BtnProcess.Visible = True
            FRM_ABSENSI_FINGERGRP_EDITOR.BtnProcess.Caption = "Update"
            FRM_ABSENSI_FINGERGRP_EDITOR.Cmb(1).Locked = False
            FRM_ABSENSI_FINGERGRP_EDITOR.TxtFinger.Enabled = True
    Case 2:
            '------------
            '-- DELETE --
            '------------
            dxDBGrid1_OnDblClick
            FRM_ABSENSI_FINGERGRP_EDITOR.BtnProcess.Caption = "Delete"
            FRM_ABSENSI_FINGERGRP_EDITOR.BtnProcess.Visible = True
            FRM_ABSENSI_FINGERGRP_EDITOR.TxtFinger.Enabled = False
            FRM_ABSENSI_FINGERGRP_EDITOR.Cmb(1).Locked = True
    Case 3:
            '------------
            '-- FILTER --
            '------------
            FilterGrid
End Select
End Sub

Public Sub FilterGrid()
 MenuGrid 0
            If cmbKar.Text = "0" Then
                'Qry0 = "Select NO_URUT,TerminalID,KAR_ID,FingerPrintID from tab_kar_finger order by TerminalID,KAR_ID ASC"
                Qry0 = "crud_finger(0,0,0)"
            Else
                'Qry0 = "Select NO_URUT,TerminalID,KAR_ID,FingerPrintID from tab_kar_finger where KAR_ID=" & CmbSorted.Text & " order by TerminalID,KAR_ID ASC"
                Qry0 = "crud_finger(1,'" & cmbKar.Text & "',0)"
            End If
            PrjSysGrid.GetGrid_Persensi dxDBGrid1, CLM0, True, True, BND0, True, Qry0, "KAR_ID", False
            LookupClm
End Sub

Private Sub CmbSorted_Click(Index As Integer)
If Index = 0 Or Index = 1 Then
    If CmbSorted(1).Text <> "" Then
        prjSysID.Employe cmbKar, CmbSorted(0).Text, 0, CmbSorted(1).Text
    End If
End If
End Sub

Private Sub Command1_Click()
'dxDBGrid1.Dataset.Filter = dxDBGrid1.Dataset.Locate("Kar_ID", 143, True, True)
'dxDBGrid1.Dataset.Filter = dxDBGrid1.Dataset.

End Sub

Private Sub dxDBGrid1_OnDblClick()
'lR = SetTopMostWindow(Me.hwnd, False)
With dxDBGrid1.Dataset
FRM_ABSENSI_FINGERGRP_EDITOR.Show
   'MsgBox .FieldValues("NO_URUT") & _
            "=" & .FieldValues("TerminalID") & _
            "=" & .FieldValues("KAR_ID") & _
            "=" & .FieldValues("FingerPrintID")
    FRM_ABSENSI_FINGERGRP_EDITOR.Cmb(0).Text = .FieldValues("KAR_ID")
    FRM_ABSENSI_FINGERGRP_EDITOR.Cmb(1).Text = .FieldValues("TerminalID")
    FRM_ABSENSI_FINGERGRP_EDITOR.TxtFinger = .FieldValues("FingerPrintID")
    FRM_ABSENSI_FINGERGRP_EDITOR.Label2 = .FieldValues("NO_URUT")
    FRM_ABSENSI_FINGERGRP_EDITOR.Cmb(0).Locked = True
    FRM_ABSENSI_FINGERGRP_EDITOR.Cmb(1).Locked = True
    FRM_ABSENSI_FINGERGRP_EDITOR.BtnProcess.Visible = False
    FRM_ABSENSI_FINGERGRP_EDITOR.TxtFinger.Enabled = False
End With
End Sub

Private Sub Form_Load()
lR = SetTopMostWindow(Me.hwnd, True)
Main
GridFill 0
prjSysID.CABANG CmbSorted(0)
prjSysID.Dept CmbSorted(1)
prjSysID.Employe cmbKar, CmbSorted(0).Text, 0, CmbSorted(1).Text
End Sub

'=====================================
'============= ptr.nov ===============
'===== PROCESS FIRST GRID LODING =====
'=====================================
Public Sub GridFill(GrdIndx As Integer)
Select Case GrdIndx
    Case Is = 0
        MenuGrid 0
            'Qry0 = "Select NO_URUT,TerminalID,KAR_ID,FingerPrintID from tab_kar_finger order by TerminalID,KAR_ID ASC"
            Qry0 = "crud_finger(0,0,0)"
            PrjSysGrid.GetGrid_Persensi dxDBGrid1, CLM0, True, True, BND0, True, Qry0, "KAR_ID", False
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
    BND0 = Array("PAYROLL REGISTER EMPLOYE", "FINGER MACHINE")
    CLM0 = Array(Array("Clm0", "No.", "NO_URUT", gedTextEdit, 0, 0, 50, 1, 1, 1, 0), _
                 Array("Clm1", "Employe ID", "KAR_ID", gedTextEdit, 0, 0, 120, 0, 1, 0, 0), _
                 Array("Clm2", "Employe Name", "KAR_ID", gedLookupEdit, 0, 0, 150, 0, 1, 0, 0), _
                 Array("Clm3", "Finger ID.", "FingerPrintID", gedTextEdit, 0, 1, 100, 0, 1, 0, 0), _
                 Array("Clm4", "Terminal NM", "TerminalID", gedLookupEdit, 0, 1, 120, 0, 1, 0, 0))
                '--------0---------1--------2-----------3--------4--5--6---7--8--9--10
End Select
End Sub

Private Sub LookupClm()
With dxDBGrid1.Columns.ColumnByName("Clm2").LookupColumn
            .LookupDataset.EnableControls
            '.LookupDataset.Close
            '.Dataset.Refresh
            .LookupDatasetType = dtADODataset
            .LookupDataset.ADODataset.ConnectionString = StrCon
            .LookupDataset.ADODataset.CursorLocation = clUseClient
            .LookupDataset.ADODataset.LockType = ltReadOnly         ' Check Permmission
            .LookupDataset.ADODataset.CommandType = cmdText
            .LookupDataset.ADODataset.CommandText = "SELECT KAR_ID,KAR_NM FROM karyawan " ' Like Join
            '.LookupDataset.Open
            .LookupKeyField = "KAR_ID"
            .LookupResultField = "KAR_NM"
            .ListColumns = "Employe Name"
            .ListFieldName = "KAR_NM"
            .ListWidth = 800
            .DisplaySize = 400
End With
With dxDBGrid1.Columns.ColumnByName("Clm4").LookupColumn
            .LookupDataset.EnableControls
            '.LookupDataset.Close
            '.Dataset.Refresh
            .LookupDatasetType = dtADODataset
            .LookupDataset.ADODataset.ConnectionString = StrCon
            .LookupDataset.ADODataset.CursorLocation = clUseClient
            .LookupDataset.ADODataset.LockType = ltReadOnly         ' Check Permmission
            .LookupDataset.ADODataset.CommandType = cmdText
            .LookupDataset.ADODataset.CommandText = "SELECT TerminalID,MESIN_NM FROM machine " ' Like Join
            '.LookupDataset.Open
            .LookupKeyField = "TerminalID"
            .LookupResultField = "MESIN_NM"
            .ListColumns = "Terminal Name"
            .ListFieldName = "MESIN_NM"
            .ListWidth = 800
            .DisplaySize = 400
End With

dxDBGrid1.Dataset.Open
End Sub




