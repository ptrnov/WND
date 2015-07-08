VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Begin VB.Form FRM_ABSENSI_MACHINE_EDITOR 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Finger Machine Editing"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   3135
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3135
      _Version        =   851970
      _ExtentX        =   5530
      _ExtentY        =   5530
      _StockProps     =   79
      BackColor       =   33023
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton BtnInput 
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   2520
         Width           =   1095
         _Version        =   851970
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtMachine 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   0
         Top             =   1200
         Width           =   2775
         _Version        =   851970
         _ExtentX        =   4895
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit TxtMachine 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   1
         Top             =   2040
         Width           =   2775
         _Version        =   851970
         _ExtentX        =   4895
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit TxtMachine 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1095
         _Version        =   851970
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Alignment       =   2
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Employe Name"
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
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Employe Name"
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
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Terminal ID"
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
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FRM_ABSENSI_MACHINE_EDITOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim QryStt As String


Private Sub BtnInput_Click()
    '-----------
    '-- SAVED --
    '-----------
    If BtnInput.Caption = "Save" Then
        AryData = TxtMachine(1).Text & "=" & TxtMachine(2).Text
        QryStt = "crud_machine(1,'" & AryData & "')"
        OpRecStt2 QryStt, True
        Set rsStt2 = Nothing
        FRM_ABSENSI_MACHINE.GridFill 0
        FRM_ABSENSI_MACHINE.dxDBGrid1.Dataset.FindLast
    End If
    
    '------------
    '-- UPDATE --
    '------------
    If BtnInput.Caption = "Update" Then
        AryData = TxtMachine(0).Text & "=" & TxtMachine(1).Text & "=" & TxtMachine(2).Text
        QryStt = "crud_machine(2,'" & AryData & "')"
        OpRecStt2 QryStt, True
        Set rsStt2 = Nothing
        FRM_ABSENSI_MACHINE.dxDBGrid1.Dataset.Refresh
        'FRM_ABSENSI_MACHINE.dxDBGrid1.Dataset.RecNo = TxtMachine(0).Text
    End If
    
    '------------
    '-- DELETE --
    '------------
    If BtnInput.Caption = "Delete" Then
        AryData = TxtMachine(0).Text & "=" & TxtMachine(1).Text & "=" & TxtMachine(2).Text
        QryStt = "crud_machine(3,'" & AryData & "')"
        OpRecStt2 QryStt, True
        Set rsStt2 = Nothing
        FRM_ABSENSI_MACHINE.dxDBGrid1.Dataset.Refresh
    End If
End Sub

Private Sub Form_Load()
init_Frm
End Sub
Private Sub init_Frm()
lR = SetTopMostWindow(FRM_ABSENSI_MACHINE.hwnd, False)
FRM_ABSENSI_MACHINE.Enabled = False
lR = SetTopMostWindow(Me.hwnd, True)
TxtMachine(1).Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FRM_ABSENSI_MACHINE.Enabled = True
    lR = SetTopMostWindow(FRM_ABSENSI_MACHINE.hwnd, True)
End Sub




