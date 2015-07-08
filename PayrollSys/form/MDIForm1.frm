VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#13.2#0"; "CODEJO~2.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   7710
   ClientLeft      =   225
   ClientTop       =   255
   ClientWidth     =   15390
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin XtremeCommandBars.CommandBars CBar1 
      Left            =   240
      Top             =   360
      _Version        =   851970
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'API tambahan utk klik kanan menu
Private Declare Sub mouse_event Lib "user32.dll" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public WithEvents StatusBar As XtremeCommandBars.StatusBar
Attribute StatusBar.VB_VarHelpID = -1
Private Sub MDIForm_Load()
Dim Ake As String

InfoServ
Main
'GetUserInfo "erick"
GetUserInfo (UsrID)


Ake = UsrID


If Len(Ake) <= 1 Then Exit Sub
'On Error GoTo ErrorLabel
LOADING.Show
LOADING.SetParm Me, 25
'Main
LOADING.SetParm Me, 50

PrjSysMn.CreateRiboonBar CBar1, Ake

'WorkspaceVisible = False
'CBar1.StatusBar.AddPane(1).Visible = True

LOADING.SetParm Me, 75

CBar1.ShowTabWorkspace True
'CommandBarsFrame1.ShowTabWorkspace True
'setStatusBar
'Me.BackColor = &H80000005
'PrjSystem.SysMenu.GET_MenuEditor CFrame, 1
'Me.Picture = LoadPicture(ImagePath("FRM_MAIN_BACK"))
'UsrID = FRM_LOGIN.txtLogin(0).Text
LOADING.SetParm Me, 85
Me.Icon = LoadPicture(ImagePath("FRM_MAIN"))
LOADING.SetParm Me, 100
CBar1.ActiveMenuBar.EnableDocking xtpFlagAlignLeft
HOME.Show

Exit Sub
ErrorLabel:
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"
    'LOADING.SetParm Me, 100
End Sub

'=======================================
'====== MENU SHOW=======================
'=======================================
Private Sub CBar1_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
Dim defCorp As String
Dim defCab As String
Dim defDep As String

defCorp = IdCorp
defCab = IdCab
defDep = IdDep

On Error Resume Next
    Select Case Control.Id
        Case 30
                    USER_EDIT.Show
                    USER_EDIT.UserEdit "pass"
        '--Login--
        Case 31
                    Unload Me
                    FRM_LOGIN.Show
        Case 32:    Unload Me
        '--Home--
        Case 100:   HOME.Show
        Case 101:   FRM_IMPORT_LOG.Show
        Case 102:   FRM_ABSENSI.Show 'KARYAWAN.Show
        Case 103:   FRM_PAYROLL.Show
        '--Karyawan--
        Case 200:   FRM_AGAMA.Show
        Case 201:   FRM_ABSENSI_IJIN.Show
        Case 202:   FRM_Departemen.Show
        Case 203:   FRM_JABATAN.Show
        Case 204:   FRM_ABSENSI_OFFDAY.Show
        Case 205:   FRM_ABSENSI_TT_GRP.Show
        Case 206:   FRM_ABSENSI_TT.Show
        Case 207:   FRM_PENDIDIKAN.Show
        Case 208:   FRM_KARYAWAN.Show
        
        '--Pengajian --
        Case 300:   FRM_KOMPONEN_GAJI.Show
        Case 301:   FRM_PPH21.Show
        Case 302:   FRM_PTKP.Show
        Case 303:   FRM_JAMSOS.Show
        Case 304:   FRM_SALARY.Show
        'USER
        Case 600:   USER.Show
        Case 700: FRM_ABSENSI_MACHINE.Show
        Case 701: FRM_ABSENSI_FINGERGRP.Show
    End Select
End Sub
'=========================================== END CONFIG FORM =======================================================
Private Sub MDIForm_Unload(Cancel As Integer)
'UnloadAll
End Sub

Sub setStatusBar()
Set StatusBar = CBar1.StatusBar
    
    StatusBar.Visible = True
    Dim Pane As StatusBarPane
    Set Pane = StatusBar.AddPane(ID_INDICATOR_LOGO)
    Pane.Visible = False
    Pane.Text = "Codejock Software"
    Pane.IconIndex = 100
    'Pane.TextColor = vbGrayText
    Pane.TextColor = RGB(64, 100, 176)
    Pane.BackgroundColor = RGB(245, 245, 245)
    Pane.Font.Bold = True
    Pane.Width = 0 'Auto size

    Set Pane = StatusBar.AddPane(0)
    Pane.Style = SBPS_STRETCH Or SBPS_NOBORDERS
    Pane.Text = "Ready"
    Pane.Width = 0 ' Autro Size
    
    Set Pane = StatusBar.AddPane(ID_INDICATOR_ICON)
    Pane.IconIndex = ID_VIEW_ADDITEM
    Pane.Alignment = 1
    Pane.Width = 0 ' Autro Size
    Pane.Visible = False
    Pane.ToolTip = "Click to show properties"
    
    StatusBar.AddPane ID_INDICATOR_CAPS
    StatusBar.AddPane ID_INDICATOR_NUM
    StatusBar.AddPane ID_INDICATOR_SCRL
    StatusBar.IdleText = "Ready"
    
    
    Set Pane = StatusBar.AddPane(ID_INDICATOR_ANIMATION)
    Animation.Open (App.Path & "\heartbeat.avi")
    Pane.Handle = Animation.hwnd
    Pane.Width = 16
    Pane.Visible = False
    
    Set Pane = StatusBar.AddPane(ID_INDICATOR_PROGRESS)
    Pane.Handle = ProgressBar.hwnd
    ProgressBar.Value = 50
    Pane.Width = 100
    Pane.Visible = False
    
    Set Pane = StatusBar.AddPane(ID_INDICATOR_TEXT)
    Pane.Handle = PaneText.hwnd
    Pane.Width = 50
    Pane.Visible = False
End Sub

