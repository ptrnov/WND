VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl32.ocx"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "COC28D~1.OCX"
Begin VB.Form USER_PER 
   Caption         =   "User Permission"
   ClientHeight    =   8280
   ClientLeft      =   90
   ClientTop       =   465
   ClientWidth     =   12390
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   12390
   StartUpPosition =   3  'Windows Default
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   4455
      Left            =   1560
      OleObjectBlob   =   "USER_PER.frx":0000
      TabIndex        =   2
      Top             =   2040
      Width           =   7455
   End
   Begin XtremeTaskPanel.TaskPanel TaskPanelStupGl 
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   495
      _Version        =   851970
      _ExtentX        =   873
      _ExtentY        =   1508
      _StockProps     =   64
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6615
      Left            =   960
      TabIndex        =   0
      Top             =   1080
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   11668
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "USER_PER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub TaskPanelStupGl_ItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
Dim StrMessage1 As String
Dim StrMessage2 As String
Dim StrMessage3 As String
Dim StrSubject1 As String
Dim StrSubject2 As String
Dim StrSubject3 As String
StrMessage1 = "Form Plan !  "
StrMessage2 = "Form None ! "
StrMessage3 = "Form None ! "
StrSubject1 = "User"
StrSubject2 = "NONE"
StrSubject3 = "NONE"

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
