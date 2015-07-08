VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Begin VB.Form FRM_ABSENSI_MACHINE 
   BackColor       =   &H0080FF80&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FINGER MACHINE"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      _Version        =   851970
      _ExtentX        =   8916
      _ExtentY        =   9551
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton BtnMsnProcess 
         Height          =   375
         Index           =   0
         Left            =   2520
         TabIndex        =   2
         Top             =   600
         Width           =   855
         _Version        =   851970
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Add"
         UseVisualStyle  =   -1  'True
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
         Height          =   4335
         Left            =   0
         OleObjectBlob   =   "FRM_ABSENSI_MACHINE.frx":0000
         TabIndex        =   1
         Top             =   1080
         Width           =   5055
      End
      Begin XtremeSuiteControls.PushButton BtnMsnProcess 
         Height          =   375
         Index           =   1
         Left            =   3240
         TabIndex        =   3
         Top             =   600
         Width           =   855
         _Version        =   851970
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Edit"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton BtnMsnProcess 
         Height          =   375
         Index           =   2
         Left            =   3960
         TabIndex        =   4
         Top             =   600
         Width           =   855
         _Version        =   851970
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Delete"
         UseVisualStyle  =   -1  'True
      End
   End
End
Attribute VB_Name = "FRM_ABSENSI_MACHINE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CLM0, CLM1, BND0, BND1 As Variant
Dim Qry0 As String, Qry1 As String

Private Sub BtnMsnProcess_Click(Index As Integer)
Select Case Index
    Case 0:
            '-----------
            '-- SAVED --
            '-----------
            FRM_ABSENSI_MACHINE_EDITOR.Show
            FRM_ABSENSI_MACHINE_EDITOR.Caption = "Finger Machine ADD"
            FRM_ABSENSI_MACHINE_EDITOR.BtnInput.Caption = "Save"
    Case 1:
            '------------
            '-- UPDATE --
            '------------
            dxDBGrid1_OnDblClick
            FRM_ABSENSI_MACHINE_EDITOR.BtnInput.Visible = True
            FRM_ABSENSI_MACHINE_EDITOR.Caption = "Finger Machine EDIT"
            FRM_ABSENSI_MACHINE_EDITOR.BtnInput.Caption = "Update"
            
    Case 2:
            '------------
            '-- DELETE --
            '------------
            dxDBGrid1_OnDblClick
            FRM_ABSENSI_MACHINE_EDITOR.BtnInput.Visible = True
            FRM_ABSENSI_MACHINE_EDITOR.Caption = "Finger Machine DELETE"
            FRM_ABSENSI_MACHINE_EDITOR.BtnInput.Caption = "Delete"
  

End Select
End Sub

Private Sub dxDBGrid1_OnDblClick()
With dxDBGrid1.Dataset
    FRM_ABSENSI_MACHINE_EDITOR.Show
    FRM_ABSENSI_MACHINE_EDITOR.TxtMachine(0).Text = .FieldValues("TerminalID")
    FRM_ABSENSI_MACHINE_EDITOR.TxtMachine(1).Text = .FieldValues("MESIN_NM")
    FRM_ABSENSI_MACHINE_EDITOR.TxtMachine(2).Text = .FieldValues("MESIN_SN")
    FRM_ABSENSI_MACHINE_EDITOR.BtnInput.Visible = False
End With
End Sub

Private Sub Form_Load()
lR = SetTopMostWindow(Me.hwnd, True)
Main
GridFill 0
  
End Sub

'=====================================
'============= ptr.nov ===============
'===== PROCESS FIRST GRID LODING =====
'=====================================
Public Sub GridFill(GrdIndx As Integer)
Select Case GrdIndx
    Case Is = 0
        MenuGrid 0
            Qry0 = "crud_machine(0,'0')"
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
    BND0 = Array("HOLIDAY DATE", "HOLIDAY DISCRIPTION")
    CLM0 = Array(Array("Clm1", "TerminalID", "TerminalID", gedTextEdit, 0, 0, 80, 6, 2, 0, 0), _
                Array("Clm2", "MESIN NAME", "MESIN_NM", gedTextEdit, 0, 0, 127, 6, 0, 0, 0), _
                Array("Clm3", "MESIN SN", "MESIN_SN", gedTextEdit, 0, 0, 110, 6, 0, 0, 0))
                '--------0---------1--------2-----------3--------4--5--6---7--8--9--10
                
End Select
End Sub






