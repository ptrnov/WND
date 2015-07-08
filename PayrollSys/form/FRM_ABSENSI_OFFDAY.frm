VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Begin VB.Form FRM_ABSENSI_OFFDAY 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "HOLIDAY"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7725
      _Version        =   851970
      _ExtentX        =   13626
      _ExtentY        =   8705
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton cmdLiburtProcess 
         Height          =   495
         Index           =   0
         Left            =   6480
         TabIndex        =   1
         Top             =   240
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
      Begin XtremeSuiteControls.PushButton cmdLiburtProcess 
         Height          =   495
         Index           =   1
         Left            =   7080
         TabIndex        =   2
         Top             =   240
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
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
         Height          =   4095
         Left            =   0
         OleObjectBlob   =   "FRM_ABSENSI_OFFDAY.frx":0000
         TabIndex        =   3
         Top             =   720
         Width           =   7695
      End
   End
End
Attribute VB_Name = "FRM_ABSENSI_OFFDAY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CLM0, CLM1, BND0, BND1 As Variant
Dim Qry0 As String, Qry1 As String
Dim a As Date

Private Sub cmdLiburtProcess_Click(Index As Integer)
On Error Resume Next
    If Index = 0 Then
        With dxDBGrid1.Dataset
            '.Next
            .Insert
            .FieldValues("TAHUN") = Year(a)
            .EnableControls
            .Refresh
            '.DisableControls
        End With
    End If
    If Index = 1 Then
        lR = SetTopMostWindow(Me.hwnd, False)
        x = MsgBox("Anda akan menghapus, YEAR[" & dxDBGrid1.Columns(0).Value & "], NAME HOLIDAY[" & dxDBGrid1.Columns(3).Value & "]", vbYesNo)
        If x = vbYes Then
         dxDBGrid1.Dataset.Delete
        End If
       lR = SetTopMostWindow(Me.hwnd, True)
    End If
End Sub

Private Sub Form_Load()
lR = SetTopMostWindow(Me.hwnd, True)
Main
GridFill 0
  a = Now
End Sub

'=====================================
'============= ptr.nov ===============
'===== PROCESS FIRST GRID LODING =====
'=====================================
Private Sub GridFill(GrdIndx As Integer)
Select Case GrdIndx
    Case Is = 0
        MenuGrid 0
            'Qry0 = "Select LBR_ID,year(LBR_SDATE) as THN,LBR_SDATE,LBR_EDATE,LBR_NM from holiday"
            Qry0 = "crud_libur(1)"
            PrjSysGrid.GetGrid_Persensi dxDBGrid1, CLM0, False, False, BND0, True, Qry0, "LBR_ID", False
            dxDBGrid1.Dataset.Open
            With dxDBGrid1.Columns(0)
                .GroupIndex = 1
                .Sorted = csDown
            End With
            dxDBGrid1.M.FullExpand
    End Select
End Sub

'=====================================
'============= ptr.nov ===============
'===== MENU CULUMN BND GRID LODING ===
'=====================================
Private Sub MenuGrid(TabSel As Integer)
Select Case TabSel
Case Is = 0
    BND0 = Array("HOLIDAY DATE", "HOLIDAY DISCRIPTION")
    CLM0 = Array(Array("Clm0", "YEAR", "TAHUN", gedTextEdit, 0, 0, 80, 1, 2, 0, 0), _
                Array("Clm1", "Start Date", "LBR_SDATE", gedDateEdit, 0, 0, 100, 11, 0, 0, 3), _
                Array("Clm2", "End Date", "LBR_EDATE", gedDateEdit, 0, 0, 100, 11, 0, 0, 3), _
                Array("Clm3", "Holiday Name.", "LBR_NM", gedTextEdit, 0, 1, 300, 11, 0, 0, 0))
                '--------0---------1--------2-----------3--------4--5--6---7--8--9--10
                
End Select
End Sub




