VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "COB03C~1.OCX"
Begin VB.Form USER 
   BackColor       =   &H00FFC0C0&
   Caption         =   "User"
   ClientHeight    =   10950
   ClientLeft      =   90
   ClientTop       =   435
   ClientWidth     =   20160
   ControlBox      =   0   'False
   DrawMode        =   15  'Merge Pen Not
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20160
   WindowState     =   2  'Maximized
   Begin XtremeTaskPanel.TaskPanel TaskPanelStupGl 
      Height          =   11535
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2295
      _Version        =   851970
      _ExtentX        =   4048
      _ExtentY        =   20346
      _StockProps     =   64
      VisualTheme     =   6
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   11535
      Left            =   2280
      TabIndex        =   0
      Top             =   0
      Width           =   18135
      _Version        =   851970
      _ExtentX        =   31988
      _ExtentY        =   20346
      _StockProps     =   68
      Appearance      =   3
      Color           =   8
      PaintManager.BoldSelected=   -1  'True
      PaintManager.ShowTabs=   0   'False
      PaintManager.FixedTabWidth=   150
      PaintManager.MaxTabWidth=   150
      PaintManager.MinTabWidth=   150
      PaintManager.ControlMargin=   "1,0,0,0"
      ItemCount       =   4
      SelectedItem    =   1
      Item(0).Caption =   "User"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "GrpUsrDetail"
      Item(0).Control(1)=   "fraUsr"
      Item(1).Caption =   "Permission"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "fraPerm"
      Item(2).Caption =   "TabControlPage1"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "TabControlPage1"
      Item(3).Caption =   "TabControlPage2"
      Item(3).ControlCount=   0
      Begin XtremeSuiteControls.GroupBox GrpUsrDetail 
         Height          =   6855
         Left            =   -63280
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   5415
         _Version        =   851970
         _ExtentX        =   9551
         _ExtentY        =   12091
         _StockProps     =   79
         Caption         =   "User Detail"
         ForeColor       =   16777215
         BackColor       =   12632064
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Begin VB.Image Image1 
            Height          =   3360
            Left            =   1080
            Picture         =   "USER.frx":0000
            Stretch         =   -1  'True
            Top             =   2640
            Width           =   3105
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "ID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   20
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   19
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Perusahaan"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   18
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cabang"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   17
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Departemen"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   16
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Jabatan"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   15
            Top             =   2160
            Width           =   1815
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   1800
            TabIndex        =   14
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   1800
            TabIndex        =   13
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   1800
            TabIndex        =   12
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   1800
            TabIndex        =   11
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   1800
            TabIndex        =   10
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   12
            Left            =   1800
            TabIndex        =   9
            Top             =   2160
            Width           =   615
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "lblUsrDetail"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   2400
            TabIndex        =   8
            Top             =   360
            Width           =   3495
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "lblUsrDetail"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   7
            Top             =   720
            Width           =   3495
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "lblUsrDetail"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   6
            Top             =   1080
            Width           =   3495
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "lblUsrDetail"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   2400
            TabIndex        =   5
            Top             =   1440
            Width           =   3495
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "lblUsrDetail"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   4
            Left            =   2400
            TabIndex        =   4
            Top             =   1800
            Width           =   3495
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "lblUsrDetail"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   5
            Left            =   2400
            TabIndex        =   3
            Top             =   2160
            Width           =   3495
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage1 
         Height          =   11475
         Left            =   -69955
         TabIndex        =   21
         Top             =   30
         Visible         =   0   'False
         Width           =   16740
         _Version        =   851970
         _ExtentX        =   29527
         _ExtentY        =   20241
         _StockProps     =   1
         Page            =   3
      End
      Begin XtremeSuiteControls.TabControlPage fraPerm 
         Height          =   11475
         Left            =   45
         TabIndex        =   22
         Top             =   30
         Width           =   18060
         _Version        =   851970
         _ExtentX        =   31856
         _ExtentY        =   20241
         _StockProps     =   1
         BackColor       =   16777215
         Page            =   2
         PictureAlignment=   32
         Begin XtremeSuiteControls.GroupBox grpUser 
            Height          =   6615
            Index           =   0
            Left            =   240
            TabIndex        =   23
            Top             =   240
            Width           =   5175
            _Version        =   851970
            _ExtentX        =   9128
            _ExtentY        =   11668
            _StockProps     =   79
            Caption         =   "   USER"
            BackColor       =   12648384
            Appearance      =   4
            BorderStyle     =   1
            Begin DXDBGRIDLibCtl.dxDBGrid dxDBGridPermUsr 
               Height          =   6015
               Left            =   120
               OleObjectBlob   =   "USER.frx":6208
               TabIndex        =   26
               Top             =   360
               Width           =   4935
            End
         End
         Begin XtremeSuiteControls.GroupBox grpUser 
            Height          =   6615
            Index           =   2
            Left            =   5760
            TabIndex        =   25
            Top             =   240
            Width           =   4815
            _Version        =   851970
            _ExtentX        =   8493
            _ExtentY        =   11668
            _StockProps     =   79
            Caption         =   "   PERMISSION"
            BackColor       =   12648447
            Appearance      =   4
            BorderStyle     =   1
            Begin XtremeSuiteControls.TreeView trePerm 
               Height          =   6015
               Left            =   120
               TabIndex        =   27
               Top             =   360
               Width           =   4575
               _Version        =   851970
               _ExtentX        =   8070
               _ExtentY        =   10610
               _StockProps     =   77
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
            End
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   6495
            Index           =   1
            Left            =   5880
            Top             =   480
            Width           =   4815
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   6495
            Index           =   0
            Left            =   360
            Top             =   480
            Width           =   5175
         End
      End
      Begin XtremeSuiteControls.TabControlPage fraUsr 
         Height          =   11475
         Left            =   -69955
         TabIndex        =   28
         Top             =   30
         Visible         =   0   'False
         Width           =   18060
         _Version        =   851970
         _ExtentX        =   31856
         _ExtentY        =   20241
         _StockProps     =   1
         BackColor       =   16777215
         Page            =   1
         PictureAlignment=   32
         Begin XtremeSuiteControls.PushButton cmdEditUser 
            Height          =   495
            Index           =   0
            Left            =   240
            TabIndex        =   29
            Top             =   240
            Width           =   2055
            _Version        =   851970
            _ExtentX        =   3625
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "&Tambah"
            Appearance      =   6
         End
         Begin DXDBGRIDLibCtl.dxDBGrid dxDBGridUser 
            Height          =   6375
            Left            =   240
            OleObjectBlob   =   "USER.frx":6EB0
            TabIndex        =   30
            Top             =   720
            Width           =   6135
         End
         Begin XtremeSuiteControls.PushButton cmdEditUser 
            Height          =   495
            Index           =   1
            Left            =   2280
            TabIndex        =   31
            Top             =   240
            Width           =   2055
            _Version        =   851970
            _ExtentX        =   3625
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "&Edit"
            Appearance      =   6
         End
         Begin XtremeSuiteControls.PushButton cmdEditUser 
            Height          =   495
            Index           =   2
            Left            =   4320
            TabIndex        =   32
            Top             =   240
            Width           =   2055
            _Version        =   851970
            _ExtentX        =   3625
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "&Hapus"
            Appearance      =   6
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   6735
            Index           =   3
            Left            =   360
            Top             =   480
            Width           =   6135
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   6735
            Index           =   2
            Left            =   6840
            Top             =   480
            Width           =   5415
         End
      End
   End
   Begin XtremeSuiteControls.GroupBox grpUser 
      Height          =   6975
      Index           =   1
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   5175
      _Version        =   851970
      _ExtentX        =   9128
      _ExtentY        =   12303
      _StockProps     =   79
      Caption         =   "USER"
      BackColor       =   12648384
      Appearance      =   4
      BorderStyle     =   1
   End
End
Attribute VB_Name = "USER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GroupSubject As TaskPanelGroup
Dim GroupTab As TaskPanelGroup
Dim GroupMsg As TaskPanelGroup
Dim Item As TaskPanelGroupItem
Dim CLM1, CLM2 As Variant
Dim BND As Variant
Dim qry As String
Dim isDraw As Boolean
'-----USER TAB COMMAND BUTTON -----
Private Sub cmdEditUser_Click(Index As Integer)
Dim UsrSlct As String 'KAR_ID = id karyawan
Dim UsrNM As String 'Nama User
Dim Rspn As Integer

Select Case Index
Case 0: 'memanggil USER_EDIT dan memberikan status "Add"
    sndChoice = "Add"
    USER_EDIT.Show
    USER_EDIT.UserEdit "show"
    'User_Edit.Caption = User_Edit.Caption & "Tambah Baru"
Case 1: 'memanggil USER_EDIT dan memberikan status "Edit"
    USER_EDIT.Show
    sndData = dxDBGridUser.Columns.Item(1).Value
    USER_EDIT.UserEdit "edit"
Case 2: 'menghapus user yang dipilih pada grid
    UsrSlct = dxDBGridUser.Columns.Item(0).Value
    UsrNM = dxDBGridUser.Columns.Item(1).Value
    
    Rspn = MsgBox("Anda Yakin Akan Menghapus USER = '" & UsrNM & "'", _
                     vbYesNo + vbQuestion + vbDefaultButton2, _
                     "Perhatian!!!")
    If Rspn = vbYes Then
        PrjSysTrig.DataDel conMain, "tab_user", "USER_ID", "'" & UsrSlct & "'" & " AND USR_NM='" & UsrNM & "'"
        PrjSysTrig.DataDel conMain, "tab_user_PER", "PRMS_ID", "'" & UsrSlct & "'"
        
        refGrid
    End If
End Select

End Sub

Private Sub dxDBGridPermUsr_OnClick()
'SetGridPerms
FillUserPerm
End Sub

Private Sub dxDBGridPermUsr_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
Dim s As Integer
Dim stsRo As Boolean

On Error Resume Next
s = Node.RecNo 'Node.Values(dxDBGridPOitem.Columns.ColumnByFieldName("NO").Index)
        If s Mod 2 = 0 Then
           Color = RGB(255, 255, 255)
        Else
           Color = &HE0E0E0
        End If
        
End Sub

Private Sub dxDBGridPermUsrHak_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
'PrjSysTrig.DataUpdate conMain, _
                      "tab_user_per", "MN_SHW='" _
                      & dxDBGridPermUsrHak.Dataset.FieldValues("MN_SHW") & _
                      "' WHERE PRMS_ID='" _
                      & dxDBGridPermUsr.Dataset.FieldValues("USER_ID") & _
                      "' AND MN_ID='" _
                      & dxDBGridPermUsrHak.Dataset.FieldValues("MN_ID") & "'"
'dxDBGridPermUsrHak.Dataset.Close
'dxDBGridPermUsrHak.Dataset.Open


                      

End Sub

Private Sub dxDBGridUser_OnClick()
Dim UsrDefSel As String
UsrDefSel = dxDBGridUser.Dataset.FieldValues("USER_ID")
Label1(0).Caption = UsrDefSel
Label1(1).Caption = GetDtaTbl(conMain, "SELECT USR_NM FROM tab_user WHERE USER_ID='" & UsrDefSel & "'", "USR_NM")
Label1(2).Caption = GetDtaTbl(conMain, "SELECT C.CORP_NM FROM tab_user A,karyawan B,tab_corp C WHERE A.USER_ID='" & UsrDefSel & "' AND A.KAR_ID=B.KAR_ID AND B.CORP_ID=C.CORP_ID", "CORP_NM")
Label1(3).Caption = GetDtaTbl(conMain, "SELECT C.CAB_NM FROM tab_user A,karyawan B,tab_cabang C WHERE A.USER_ID='" & UsrDefSel & "' AND A.KAR_ID=B.KAR_ID AND B.CAB_ID=C.CAB_ID", "CAB_NM")
Label1(4).Caption = GetDtaTbl(conMain, "SELECT C.DEP_NM FROM tab_user A,karyawan B,departemen C WHERE A.USER_ID='" & UsrDefSel & "' AND A.KAR_ID=B.KAR_ID AND B.DEP_ID=C.DEP_ID", "DEP_NM")
Label1(5).Caption = GetDtaTbl(conMain, "SELECT C.JAB_NM FROM tab_user A,karyawan B,jabatan C WHERE A.USER_ID='" & UsrDefSel & "' AND A.KAR_ID=B.KAR_ID AND B.JAB_ID=C.JAB_ID", "JAB_NM")
'Label1(2).Caption = ""


End Sub

Private Sub dxDBGridUser_OnDblClick()
cmdEditUser_Click (1)
End Sub

Private Sub dxDBGridUser_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
 Select Case KeyCode
    Case Is = 46
        cmdEditUser_Click (2)
    Case Is = 13
        cmdEditUser_Click (1)
 End Select
End Sub

Private Sub Form_Resize()
On Error Resume Next
TaskPanelStupGl.Move 0, 0, 2535, ScaleHeight
TabControl1.Move TaskPanelStupGl.Width, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub Form_Load()
On Error GoTo ErrorLabel
'InfoServ
'Main
'GetUserInfo "alam_fod"

LOADING.Show
LOADING.SetParm Me, 25

    CreateTaskPanel 'memanggil fungsi menampilkan properti panel skin
LOADING.SetParm Me, 35
    MenuGrid (0) 'memilih menu pertama dari tab index 0
    refGrid 'menampilkan record user view
    Me.Icon = LoadPicture(ImagePath("FRM_USER"))
    fraUsr.Picture = LoadPicture(ImagePath(UCase(IdCorp)))
    fraPerm.Picture = LoadPicture(ImagePath(UCase(IdCorp)))
    isDraw = False
    InitExample
LOADING.SetParm Me, 100
Exit Sub
ErrorLabel:
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"
    LOADING.SetParm Me, 100
End Sub
'=========================================== END CONFIG FORM =======================================================

Sub CreateTaskPanel()
    Set GroupSubject = TaskPanelStupGl.Groups.Add(0, "MENU AKTIF")
    GroupSubject.Expandable = False
    GroupSubject.Expanded = False
    GroupSubject.Special = True
  
    Set GroupTab = TaskPanelStupGl.Groups.Add(1, "Menu Pilihan")
        GroupTab.ToolTip = "Pilih Menu yang di inginkan"
        GroupTab.Items.Add 1, "User", xtpTaskItemTypeLink, 1
        GroupTab.Items.Add 2, "Permission", xtpTaskItemTypeLink, 2
        'GroupTab.Items.Add 3, "NONE", xtpTaskItemTypeLink, 3
        GroupTab.Special = True
    
    Set GroupMsg = TaskPanelStupGl.Groups.Add(2, "Pesan")
        GroupMsg.ToolTip = "Pesan Form Pengunaan"
        GroupMsg.Items.Add 1, "", xtpTaskItemTypeText, 4
        GroupMsg.Special = True
        
    TabControl1.SelectedItem = 0
    GroupSubject.Caption = "User"
End Sub

'-----------Menu Klik------------
Private Sub TabControl1_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim oRS As New ADODB.Recordset
Dim DefMn As String

MenuGrid (Item.Index)
Select Case Item.Index
Case Is = 0
    refGrid
Case Is = 1
            
    qry = "SELECT DISTINCT USER_ID,USR_NM FROM tab_user ORDER BY USR_NM ASC"
    PrjSysGrid.GetBaseGrid dxDBGridPermUsr, CLM1, False, False, BND, False, qry, "USER_ID", False
    
        qry = "SELECT DISTINCT B.MN_ID,A.MN_NM,B.MN_SHW " & _
          "FROM tab_user_menu A,tab_user_per B  " & _
          "WHERE B.MN_ID=A.MN_ID AND B.PRMS_ID='" & _
          dxDBGridPermUsr.Dataset.FieldValues("USER_ID") & _
          "' ORDER BY B.MN_ID ASC"
    'PrjSysGrid.GetBaseGrid dxDBGridPermUsrHak, CLM2, False, False, BND, False, Qry, "MN_ID", True
End Select
End Sub

Private Sub TaskPanelStupGl_ItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
Dim StrMessage1 As String
Dim StrMessage2 As String
Dim StrMessage3 As String
Dim StrSubject1 As String
Dim StrSubject2 As String
Dim StrSubject3 As String
StrMessage1 = "Form User (View, Tambah, Hapus dan Edit)  "
StrMessage2 = "Form Permission (Memberikan hak akses menu untuk User) "
'StrMessage3 = "Form None ! "
StrSubject1 = "User"
StrSubject2 = "Permission"
'StrSubject3 = "NONE"

If Item.Id = 1 Then
    GroupSubject.Caption = StrSubject1
    GroupMsg.Items.Find(1).Remove
    GroupMsg.Items.Add 1, StrMessage1, xtpTaskItemTypeText, 4
    TabControl1.SelectedItem = 0
End If
If Item.Id = 2 Then
    GroupSubject.Caption = StrSubject2
    GroupMsg.Items.Find(1).Remove
    GroupMsg.Items.Add 1, StrMessage2, xtpTaskItemTypeText, 4
    TabControl1.SelectedItem = 1
End If
If Item.Id = 3 Then
    GroupSubject.Caption = StrSubject3
    GroupMsg.Items.Find(1).Remove
    GroupMsg.Items.Add 1, StrMessage3, xtpTaskItemTypeText, 4
    TabControl1.SelectedItem = 2
End If
End Sub

Private Sub MenuGrid(TabSel As Integer)
Select Case TabSel
Case Is = 0
    '----------
    '== BIND ==
    '----------
    BND = Array("User", "Profile")
    '------------
    '== COLUMN ==
    '------------
   
    CLM1 = Array(Array("Column1", "ID", "USER_ID", gedTextEdit, 0, 0, 100, 1, 1, 0), _
                    Array("Column2", "Nama", "USR_NM", gedTextEdit, 0, 0, 200, 1, 1, 0), _
                    Array("Column3", "Status Aktif", "USR_OFF", gedCheckEdit, 0, 1, 100, 1, 1, 0))
Case Is = 1
    BND = Array("User")
        CLM1 = Array(Array("Column1", "User ID", "USER_ID", gedTextEdit, 0, 0, 100, 1, 1, 0), _
                    Array("Column2", "User Name", "USR_NM", gedTextEdit, 0, 0, 150, 0, 0, 0))
        CLM2 = Array(Array("Column1", "Menu ID", "MN_ID", gedTextEdit, 0, 0, 100, 1, 1, 0), _
                    Array("Column2", "Menu Name", "MN_NM", gedTextEdit, 0, 0, 150, 1, 0, 0), _
                    Array("Column3", "Hak Akses", "MN_SHW", gedCheckEdit, 0, 0, 100, 0, 0, 0))
                    
End Select
End Sub

Public Sub refGrid()
    qry = "SELECT tab_user.USER_ID, tab_user.USR_NM, tab_user.USR_OFF From tab_user ORDER BY tab_user.USR_NM"
    PrjSysGrid.GetBaseGrid dxDBGridUser, CLM1, True, True, BND, False, qry, "USER_ID", False
End Sub

Public Sub SetGridPerms()
    GetMenu dxDBGridPermUsr.Dataset.FieldValues("USER_ID")

    'Qry = "SELECT DISTINCT B.MN_ID,A.MN_NM,B.MN_SHW " & _
          "FROM tab_user_menu A,tab_user_per B  " & _
          "WHERE B.MN_ID=A.MN_ID AND B.PRMS_ID='" & _
          dxDBGridPermUsr.Dataset.FieldValues("USER_ID") & _
          "' ORDER BY B.MN_ID ASC"
    'PrjSysGrid.GetBaseGrid dxDBGridPermUsrHak, CLM2, False, False, BND, False, Qry, "MN_ID", True
End Sub
   
Private Sub GetMenu(UserDef As String)
Dim oRS As New ADODB.Recordset
Dim DefMn As String

      Set oRS = ExecuteRecordSetMySql("SELECT MN_ID FROM tab_user_menu", conMain)
               With oRS
                .MoveFirst
                Do While Not .EOF
                DefMn = GetDtaTbl(conMain, "SELECT MN_ID FROM tab_user_per WHERE MN_ID='" & .Fields("MN_ID").Value & "' AND PRMS_ID='" & UserDef & "'", "MN_ID")
                  If DefMn <> .Fields("MN_ID").Value Then
                    PrjSysTrig.DataIns conMain, "tab_user_per(PRMS_ID,MN_ID,MN_SHW)", "'" & UserDef & "','" & .Fields("MN_ID").Value & "','0'"
                  End If
                  
                'oSheet.Cells(15 + curRow, 23) = .Fields("BRG_STN").Value
                .MoveNext
                Loop
              End With
End Sub

Private Sub FillUserPerm()
 Dim oRS As New ADODB.Recordset
 Dim oRSS As New ADODB.Recordset
 Dim DefMn, defUsr As String
 Dim maxPar, maxChl, K, J As Integer
 K = 1
 J = 1
 isDraw = True
On Error Resume Next
trePerm.SingleSel = True
    trePerm.ShowLines = xtpTreeViewShowLines
    trePerm.CheckBoxes = True
    trePerm.Nodes.Clear
 
      LOADING.Show
      LOADING.SetParm Me, 0
      maxPar = Val(GetDtaTbl(conMain, "SELECT COUNT(MN_ID) AS MNO FROM tab_user_menu WHERE MN_PRN='0'", "MNO"))
      maxPar = 100 / maxPar
      
      Set oRS = ExecuteRecordSetMySql("SELECT MN_ID,MN_NM FROM tab_user_menu WHERE MN_PRN='0'", conMain)
        With oRS
          .MoveFirst
             Do While Not .EOF
              LOADING.SetParm Me, maxPar * J
               trePerm.Nodes.Add , , .Fields("MN_ID"), .Fields("MN_NM")
                   trePerm.Nodes("" & .Fields("MN_ID") & "").Checked = prjSysID.GetUserpPerm(prjSysID.GetGridDefValue(dxDBGridPermUsr, "USER_ID"), "" & .Fields("MN_ID") & "")
                  ' MsgBox "_" & .Fields("MN_ID") & "_ " & .Fields("MN_NM")
                  maxChl = Val(GetDtaTbl(conMain, "SELECT COUNT(MN_ID) AS MNO FROM tab_user_menu MN_ID<>'0' AND MN_PRN='" & .Fields("MN_ID"), "MNO"))
                  maxChl = maxPar / maxChl
                  Set oRSS = ExecuteRecordSetMySql("SELECT MN_ID,MN_NM FROM tab_user_menu WHERE MN_ID<>'0' AND MN_ID<>'999' AND MN_PRN='" & .Fields("MN_ID") & "'", conMain)
                        oRSS.MoveFirst
                            Do While Not oRSS.EOF
                            LOADING.SetParm Me, maxChl * K
                                trePerm.Nodes.Add "" & .Fields("MN_ID") & "", xtpTreeViewChild, "" & oRSS.Fields("MN_ID") & "", "" & oRSS.Fields("MN_NM") & ""
                                trePerm.Nodes("" & oRSS.Fields("MN_ID") & "").Checked = prjSysID.GetUserpPerm(prjSysID.GetGridDefValue(dxDBGridPermUsr, "USER_ID"), "" & oRSS.Fields("MN_ID") & "")
                            oRSS.MoveNext
                            K = K + 1
                            Loop
                trePerm.Nodes("" & .Fields("MN_ID") & "").Expanded = True
                .MoveNext
                J = J + 1
                Loop
              End With
    
    trePerm.Nodes.Item(1).Selected = True
    isDraw = False
    
    LOADING.SetParm Me, 100
End Sub


Sub InitExample()
 dxDBGridPermUsr.Event = 1 'EGOnCustomDrawCell
dxDBGridPermUsr.EventEnabled = True
End Sub

Private Sub trePerm_NodeCheck(ByVal Node As XtremeSuiteControls.TreeViewNode)
If isDraw = False Then
    LOADING.Show
    LOADING.SetParm Me, 50
        PrjSysTrig.UpdateUserPerm prjSysID.GetGridDefValue(dxDBGridPermUsr, "USER_ID"), Str(Node.Key), Node.Checked
    LOADING.SetParm Me, 100
    
End If
End Sub

Private Sub trePerm_NodeClick(ByVal Node As XtremeSuiteControls.TreeViewNode)
'MsgBox trePerm.SelectedItem.Text

End Sub
