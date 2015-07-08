VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Begin VB.Form LOADING 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Loading"
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5370
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   16  'Merge Pen
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton cmdBatal 
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   120
      Width           =   735
      _Version        =   851970
      _ExtentX        =   1296
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Batal"
      Appearance      =   6
   End
   Begin XtremeSuiteControls.ProgressBar progBar 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      _Version        =   851970
      _ExtentX        =   7646
      _ExtentY        =   661
      _StockProps     =   93
      Text            =   "Loading..."
      BackColor       =   16777215
      Scrolling       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
      FlatStyle       =   -1  'True
      TextAlignment   =   2
   End
End
Attribute VB_Name = "LOADING"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lR As Long
Dim defFrm As Form

Public Sub SetParm(FrmDef As Form, progVal As Integer)
Set defFrm = FrmDef
progBar.Value = progVal

If FrmDef.Name = "FRM_LOGIN" Then FrmDef.Enabled = False

'MsgBox progVal
If progVal >= 99 Then
    If FrmDef.Name = "FRM_LOGIN" Then FrmDef.Enabled = True 'Else MDIForm1.Enabled = True
        lR = SetTopMostWindow(LOADING.hwnd, False)
        Unload Me
    End If
End Sub

Private Sub cmdBatal_Click()
    defFrm.Enabled = True
    MDIForm1.Enabled = True
    lR = SetTopMostWindow(LOADING.hwnd, False)
    Unload Me
End Sub

Private Sub Form_Activate()
lR = SetTopMostWindow(LOADING.hwnd, True)
End Sub

Private Sub Form_Load()
lR = SetTopMostWindow(LOADING.hwnd, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
lR = SetTopMostWindow(LOADING.hwnd, False)
End Sub

