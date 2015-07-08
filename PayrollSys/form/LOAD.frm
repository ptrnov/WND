VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Begin VB.Form LOAD 
   Caption         =   "Loading..."
   ClientHeight    =   1530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   ScaleHeight     =   1530
   ScaleWidth      =   4425
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.PushButton cmdBatal 
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   480
      Width           =   1095
      _Version        =   851970
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "PushButton1"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ProgressBar progBar 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2895
      _Version        =   851970
      _ExtentX        =   5106
      _ExtentY        =   661
      _StockProps     =   93
   End
   Begin VB.Timer tmrLoad 
      Interval        =   1000
      Left            =   120
      Top             =   960
   End
End
Attribute VB_Name = "LOAD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lR As Long
Dim frmDef As Form
Dim progVal As Integer
Public Sub SetParm(XfrmDef As Form, XprogVal As Integer)
tmrLoad.Enabled = True
Set frmDef = XfrmDef
progVal = XprogVal
End Sub

Private Sub cmdBatal_Click()
    tmrLoad.Enabled = False
    Unload Me
End Sub

Private Sub Form_Load()
lR = SetTopMostWindow(LOADING.hwnd, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
lR = SetTopMostWindow(LOADING.hwnd, False)
End Sub

Private Sub tmrLoad_Timer()
progBar.Value = progVal

If frmDef.Name = "FRM_LOGIN" Then frmDef.Enabled = False

'MsgBox progVal
If progVal >= 99 Then
    If frmDef.Name = "FRM_LOGIN" Then frmDef.Enabled = True 'Else MDIForm1.Enabled = True
        
        tmrLoad.Enabled = False
        lR = SetTopMostWindow(LOADING.hwnd, False)
        Unload Me
End If
    
End Sub
