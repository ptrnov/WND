VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Begin VB.Form FRM_ABSENSI_FINGERGRP_EDITOR 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Employe FingerID Editing "
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3360
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      _Version        =   851970
      _ExtentX        =   5953
      _ExtentY        =   4895
      _StockProps     =   79
      BackColor       =   8438015
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton BtnProcess 
         Height          =   495
         Left            =   2160
         TabIndex        =   9
         Top             =   2160
         Width           =   1095
         _Version        =   851970
         _ExtentX        =   1931
         _ExtentY        =   873
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtFinger 
         CausesValidation=   0   'False
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   2160
         Width           =   975
         _Version        =   851970
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.ComboBox Cmb 
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   3135
         _Version        =   851970
         _ExtentX        =   5530
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox Cmb 
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   3135
         _Version        =   851970
         _ExtentX        =   5530
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
         Text            =   "ComboBox1"
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   600
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   ": Reg.No"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   7
         Top             =   240
         Width           =   975
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
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Terminal"
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
         TabIndex        =   5
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Finger ID"
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
         Left            =   120
         TabIndex        =   4
         Top             =   1920
         Width           =   975
      End
   End
End
Attribute VB_Name = "FRM_ABSENSI_FINGERGRP_EDITOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim QryStt As String
Private Sub BtnProcess_Click()
    '-----------
    '-- SAVED --
    '-----------
    If BtnProcess.Caption = "Save" Then
        'MsgBox "save"
        AryData = 0 & _
                   "=" & Cmb(0).Text & _
                   "=" & Cmb(1).Text & _
                   "=" & TxtFinger.Text
                    'MsgBox AryData
        QryStt = "kar_finger(2,'0','" & AryData & "')"
        OpRecStt2 QryStt, True
        Set rsStt2 = Nothing
        FRM_ABSENSI_FINGERGRP.FilterGrid
        FRM_ABSENSI_FINGERGRP.dxDBGrid1.Dataset.FindLast
    End If
    '------------
    '-- UPDATE --
    '------------
    If BtnProcess.Caption = "Update" Then
        AryData = Label2.Caption & _
            "=" & Cmb(0).Text & _
            "=" & Cmb(1).Text & _
            "=" & TxtFinger.Text
             'MsgBox AryData
        QryStt = "kar_finger(3,'0','" & AryData & "')"
        OpRecStt2 QryStt, True
        Set rsStt2 = Nothing
        FRM_ABSENSI_FINGERGRP.dxDBGrid1.Dataset.Refresh
        FRM_ABSENSI_FINGERGRP.dxDBGrid1.Dataset.RecNo = Label2.Caption
    End If
    '------------
    '-- DELETE --
    '------------
    If BtnProcess.Caption = "Delete" Then
        AryData = Label2.Caption & _
                  "=" & Cmb(0).Text & _
                  "=" & Cmb(1).Text & _
                  "=" & TxtFinger.Text
                'MsgBox AryData
        QryStt = "kar_finger(4,'0','" & AryData & "')"
        OpRecStt2 QryStt, True
        Set rsStt2 = Nothing
        FRM_ABSENSI_FINGERGRP.dxDBGrid1.Dataset.Refresh

    End If
End Sub

Private Sub Form_Load()
init_Frm

End Sub

Private Sub Form_Unload(Cancel As Integer)
lR = SetTopMostWindow(FRM_ABSENSI_FINGERGRP.hwnd, True)
FRM_ABSENSI_FINGERGRP.Enabled = True
End Sub


Private Sub init_Frm()
lR = SetTopMostWindow(FRM_ABSENSI_FINGERGRP.hwnd, False)
FRM_ABSENSI_FINGERGRP.Enabled = False
lR = SetTopMostWindow(Me.hwnd, True)
prjSysID.EmployeNoAll Cmb(0), 0, 0, 0
prjSysID.FingerMachineNoAll Cmb(1)
End Sub


'-- Text Number input
Private Sub TxtFinger_KeyPress(KeyAscii As Integer)
 If Not IsNumeric(TxtFinger.Text & Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
End Sub
