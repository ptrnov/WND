VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Begin VB.Form FRM_ABSENSI_IJIN 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Employe Exception"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   16020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   16020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15975
      _Version        =   851970
      _ExtentX        =   28178
      _ExtentY        =   11668
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.ComboBox CmbIjin 
         Height          =   315
         Index           =   0
         Left            =   7680
         TabIndex        =   3
         Top             =   600
         Width           =   2415
         _Version        =   851970
         _ExtentX        =   4260
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
         Height          =   5535
         Left            =   120
         OleObjectBlob   =   "FRM_ABSENSI_IJIN.frx":0000
         TabIndex        =   1
         Top             =   920
         Width           =   3135
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid2 
         Height          =   5535
         Left            =   3360
         OleObjectBlob   =   "FRM_ABSENSI_IJIN.frx":0C9A
         TabIndex        =   2
         Top             =   920
         Width           =   12615
      End
      Begin XtremeSuiteControls.ComboBox CmbIjin 
         Height          =   315
         Index           =   1
         Left            =   10080
         TabIndex        =   4
         Top             =   600
         Width           =   2415
         _Version        =   851970
         _ExtentX        =   4260
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton cmdIjinProcess 
         Height          =   495
         Index           =   0
         Left            =   14400
         TabIndex        =   7
         Top             =   360
         Width           =   615
         _Version        =   851970
         _ExtentX        =   1085
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "+"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton cmdIjinProcess 
         Height          =   495
         Index           =   1
         Left            =   15000
         TabIndex        =   8
         Top             =   360
         Width           =   615
         _Version        =   851970
         _ExtentX        =   1085
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "-"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.ComboBox CmbIjin 
         Height          =   315
         Index           =   2
         Left            =   3360
         TabIndex        =   9
         Top             =   600
         Width           =   2175
         _Version        =   851970
         _ExtentX        =   3836
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.ComboBox CmbIjin 
         Height          =   315
         Index           =   3
         Left            =   5520
         TabIndex        =   10
         Top             =   600
         Width           =   2175
         _Version        =   851970
         _ExtentX        =   3836
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
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
         Index           =   3
         Left            =   5520
         TabIndex        =   12
         Top             =   360
         Width           =   1095
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
         Index           =   2
         Left            =   3360
         TabIndex        =   11
         Top             =   360
         Width           =   1335
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
         Height          =   255
         Index           =   1
         Left            =   10080
         TabIndex        =   6
         Top             =   360
         Width           =   1095
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
         Height          =   255
         Index           =   0
         Left            =   7680
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FRM_ABSENSI_IJIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CLM0, CLM1, BND0, BND1 As Variant
Dim Qry0 As String, Qry1 As String

Private Sub CmbIjin_Click(Index As Integer)
If Index = 2 Or Index = 3 Then
    If CmbIjin(2).Text <> "" And CmbIjin(3).Text <> "" Then
        prjSysID.Employe CmbIjin(1), CmbIjin(2), CmbIjin(3), CmbIjin(0)
    End If
End If
If Index = 0 Then
    prjSysID.Employe CmbIjin(1), 0, 0, CmbIjin(0)
End If
End Sub

Private Sub cmdIjinProcess_Click(Index As Integer)
On Error Resume Next
    If Index = 0 And CmbIjin(0).Text <> "0" And CmbIjin(1).Text <> "0" Then
        With dxDBGrid2.Dataset
            .Insert
            .FieldValues("KAR_ID") = CmbIjin(1).Text
            .FieldValues("IJN_ID") = dxDBGrid1.Columns(0).Value
            .EnableControls
            '.Refresh
        End With
    End If
    If Index = 1 Then
       ' MsgBox dxDBGrid1.Columns(0).Value
       lR = SetTopMostWindow(Me.hwnd, False)
       a = MsgBox("Apakah Anada Akan Menghapus Employe ID=" & dxDBGrid2.Columns(0).Value, vbYesNo, "IJIN FORM CONFIRM")
       If a = vbYes Then
         dxDBGrid2.Dataset.Delete
       End If
       lR = SetTopMostWindow(Me.hwnd, True)
    End If
End Sub

Private Sub Form_Load()
lR = SetTopMostWindow(Me.hwnd, True)
Main
GridFill 0
GridFill 1
prjSysID.Dept CmbIjin(0)
prjSysID.CABANG CmbIjin(2)
prjSysID.TTGROUP CmbIjin(3)
End Sub

'=====================================
'============= ptr.nov ===============
'===== PROCESS FIRST GRID LODING =====
'=====================================
Private Sub GridFill(GrdIndx As Integer)
Select Case GrdIndx
    Case Is = 0
        MenuGrid 0
            'Qry0 = "Select IJN_ID,IIJN_NM from  ijin_header"
            Qry0 = "crud_ijin(1)"
            PrjSysGrid.GetGrid_Persensi dxDBGrid1, CLM0, False, False, BND0, True, Qry0, "LBR_ID", False
            dxDBGrid1.Dataset.Open
   Case Is = 1
        MenuGrid 1
            'Qry1 = "Select KAR_ID,IJN_SDATE,IJN_EDATE,IJN_ID,IJN_NOTE from ijin_detail"
            Qry1 = "crud_ijin(2)"
            PrjSysGrid.GetGrid_Persensi dxDBGrid2, CLM1, False, False, BND1, True, Qry1, "IJN_NOTE", False
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
    BND0 = Array("EXCEPTION LIST")
    CLM0 = Array(Array("Clm0", "Izin ID", "IJN_ID", gedTextEdit, 0, 0, 60, 6, 2, 0, 0), _
                Array("Clm1", "IIJN NAME", "IIJN_NM", gedTextEdit, 0, 0, 130, 1, 1, 0, 3))
                '--------0---------1--------2-----------3--------4--5--6---7--8--9--10
                
Case Is = 1
    BND1 = Array("EXCEPTION DATA")
    CLM1 = Array(Array("Clm0", "EMP ID", "KAR_ID", gedTextEdit, 0, 0, 130, 1, 0, 0, 0), _
                 Array("Clm1", "EMP NAME", "KAR_ID", gedLookupEdit, 0, 0, 120, 1, 0, 0, 0), _
                 Array("Clm2", "EXCP NAME", "IJN_ID", gedLookupEdit, 0, 0, 130, 1, 0, 0, 0), _
                 Array("Clm3", "START DATE", "IJN_SDATE", gedDateEdit, 0, 0, 80, 11, 2, 0, 0), _
                 Array("Clm4", "END DATE", "IJN_EDATE", gedDateEdit, 0, 0, 80, 11, 2, 0, 3), _
                 Array("Clm5", "DESCRIPTION", "IJN_NOTE", gedMemoEdit, 0, 0, 300, 11, 0, 0, 0))
                '--------0---------1--------2-----------3--------4--5--6---7--8--9--10
                'KAR_ID,IJN_SDATE,IJN_EDATE,IJN_ID,IJN_NOTE
End Select
End Sub

Private Sub LookupClm()
On Error Resume Next
    With dxDBGrid2.Columns.ColumnByName("Clm1").LookupColumn
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
                .ListColumns = "KAR_NM"
                .ListFieldName = "KAR_NM"
                .ListWidth = 800
                .DisplaySize = 400
    End With
    With dxDBGrid2.Columns.ColumnByName("Clm2").LookupColumn
                .LookupDataset.EnableControls
                '.LookupDataset.Close
                '.Dataset.Refresh
                .LookupDatasetType = dtADODataset
                .LookupDataset.ADODataset.ConnectionString = StrCon
                .LookupDataset.ADODataset.CursorLocation = clUseClient
                .LookupDataset.ADODataset.LockType = ltReadOnly         ' Check Permmission
                .LookupDataset.ADODataset.CommandType = cmdText
                .LookupDataset.ADODataset.CommandText = "SELECT IJN_ID,IIJN_NM FROM ijin_header " ' Like Join
                '.LookupDataset.Open
                .LookupKeyField = "IJN_ID"
                .LookupResultField = "IIJN_NM"
                .ListColumns = "IIJN NAME"
                .ListFieldName = "IIJN_NM"
                .ListWidth = 800
                .DisplaySize = 400
    End With
    dxDBGrid2.Dataset.Open
End Sub





