VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Begin VB.Form FRM_KARYAWAN_ADD 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generate  Employe ID"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   4350
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      _Version        =   851970
      _ExtentX        =   7646
      _ExtentY        =   3201
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton BtnPrcs 
         Height          =   1095
         Index           =   0
         Left            =   3000
         TabIndex        =   3
         Top             =   480
         Width           =   1095
         _Version        =   851970
         _ExtentX        =   1931
         _ExtentY        =   1931
         _StockProps     =   79
         Caption         =   "Get ID"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit Lbl_Kar 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   2535
         _Version        =   851970
         _ExtentX        =   4471
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox Cmb_Kar 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   2535
         _Version        =   851970
         _ExtentX        =   4471
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "ComboBox1"
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "New EmployeID"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Cabang"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FRM_KARYAWAN_ADD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim qry As String

Private Sub BtnPrcs_Click(Index As Integer)
If Index = 0 Then
    qry = "crud_karyawan(0,'" & Cmb_Kar.Text & "',0,0,0,0,0)"
    OpRec2 qry, True
    With rs2
        Lbl_Kar(0).Text = .Fields("KAR_ID").Value
    End With
End If
End Sub

Private Sub Form_Load()
lR = SetTopMostWindow(Me.hwnd, True)
Main
prjSysID.CABANG_NOALL Cmb_Kar
End Sub

Private Sub Form_Unload(Cancel As Integer)
FRM_KARYAWAN.GridFill 1, 1
End Sub
