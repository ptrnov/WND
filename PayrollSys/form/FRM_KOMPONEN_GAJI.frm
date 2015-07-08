VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Begin VB.Form FRM_KOMPONEN_GAJI 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "List Komponen Pengajian"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   6795
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControlPage TabBack 
      Height          =   6975
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11415
      _Version        =   851970
      _ExtentX        =   20135
      _ExtentY        =   12303
      _StockProps     =   1
      BackColor       =   4772349
      PictureAlignment=   32
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   495
         Left            =   9600
         TabIndex        =   6
         Top             =   480
         Width           =   1215
         _Version        =   851970
         _ExtentX        =   2143
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Simulasikan"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtSml 
         Height          =   375
         Index           =   0
         Left            =   7320
         TabIndex        =   5
         Top             =   600
         Width           =   1575
         _Version        =   851970
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton cmdList 
         Height          =   735
         Index           =   0
         Left            =   5760
         TabIndex        =   1
         Top             =   6120
         Width           =   735
         _Version        =   851970
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Hapus"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.PushButton cmdList 
         Height          =   735
         Index           =   1
         Left            =   4080
         TabIndex        =   2
         Top             =   6120
         Width           =   735
         _Version        =   851970
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Add"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.PushButton cmdList 
         Height          =   735
         Index           =   2
         Left            =   4920
         TabIndex        =   3
         Top             =   6120
         Width           =   735
         _Version        =   851970
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Edit"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         ImageAlignment  =   6
         TextImageRelation=   1
         ImageGap        =   0
      End
      Begin XtremeSuiteControls.FlatEdit TxtSml 
         Height          =   375
         Index           =   1
         Left            =   7320
         TabIndex        =   7
         Top             =   1440
         Width           =   1575
         _Version        =   851970
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid 
         Height          =   5775
         Left            =   240
         OleObjectBlob   =   "FRM_KOMPONEN_GAJI.frx":0000
         TabIndex        =   0
         Top             =   240
         Width           =   6255
      End
   End
End
Attribute VB_Name = "FRM_KOMPONEN_GAJI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CLM1 As Variant
Dim lR As Long
Dim i As Integer

Private Sub cmdList_Click(Index As Integer)
Dim defId As String
On Error GoTo ErrorLabel
LOADING.Show
LOADING.SetParm Me, 10


Select Case Index
    Case 0 'tombol delete
        LOADING.SetParm Me, 100
        
        Dim Aa As String
        Aa = MsgBox("Anda yakin ingin menghapus Data  '" & prjSysID.GetGridDefValue(dxDBGrid, "NO") & "' ?", vbYesNo, "Perhatian")
        
        If Aa = vbYes Then
            PrjSysTrig.DataDel conMain, "NO", "NO", "'" & prjSysID.GetGridDefValue(dxDBGrid, "NO") & "'"
            MenuGrid
            GridFill
        Else
            Exit Sub
        End If
        Exit Sub
    Case 1
        '=========
        '---Add---
        '=========
        If cmdList(1).Caption = "Add" Then
           With dxDBGrid.Dataset
               .Insert
                For i = 1 To 3
                    dxDBGrid.Columns(i).DisableEditor = False
                    dxDBGrid.Columns(i).Color = &HFFFFFF
                    cmdList(1).Caption = "Accept"
                Next i
            End With
        Else
            cmdList(1).Caption = "Add"
            dxDBGrid.Dataset.Refresh
        End If
    Case 2
        '==========
        '---Edit---
        '==========
        If cmdList(2).Caption = "Edit" Then
            For i = 1 To 3
                dxDBGrid.Columns(i).DisableEditor = False
                dxDBGrid.Columns(i).Color = &HFFFFFF
                cmdList(2).Caption = "Accept"
            Next i
        Else
            For i = 1 To 3
                dxDBGrid.Dataset.Refresh
                dxDBGrid.Columns(i).DisableEditor = True
                dxDBGrid.Columns(i).Color = &HCBFAFE
                cmdList(2).Caption = "Edit"
            Next i
        End If
End Select
LOADING.SetParm Me, 100
Exit Sub
ErrorLabel:
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"
    LOADING.SetParm Me, 100
End Sub

Private Sub dxDBGrid_OnChangeColumn(ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal OldColumn As DXDBGRIDLibCtl.IdxGridColumn, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn)
dxDBGrid.Dataset.Edit
dxDBGrid.Dataset.EnableControls
End Sub

Private Sub dxDBGrid_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
Dim s As Integer
Dim stsRo As Boolean

On Error Resume Next

s = Node.RecNo 'Node.Values(dxDBGridROitem.Columns.ColumnByFieldName("NO").Index)
        If s Mod 2 = 0 Then
           Color = RGB(255, 255, 255)
        Else
           Color = &HE0E0E0                'RGB(255, 255, 222)
        End If
End Sub

Private Sub Form_Load()
lR = SetTopMostWindow(Me.hwnd, True)

On Error GoTo ErrorLabel
LOADING.Show
LOADING.SetParm Me, 25

'InitExample

'SrcBy = "True"
'fill combo reff
'cmbReff.Clear
'cmbReff.AddItem "Cari Nama"
'cmbReff.ListIndex = 0

LOADING.SetParm Me, 45

'fill image properti
'cmdList(0).Picture = LoadPicture(ImagePath("BTN_EDIT")) 'tombol Pilih
'cmdList(1).Picture = LoadPicture(ImagePath("BTN_ADD")) 'tombol TAMBAH
'cmdList(2).Picture = LoadPicture(ImagePath("BTN_WRITE")) 'tombol EDIT
'TabBack.PictureAlignment = xtpPictureTile
'TabBack.Picture = LoadPicture(ImagePath("PATERN1"))
Main
LOADING.SetParm Me, 45
GridFill



LOADING.SetParm Me, 100

Exit Sub
ErrorLabel:
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"
    LOADING.SetParm Me, 100

End Sub

Private Sub MenuGrid()
    CLM1 = Array(Array("Clm0", "NO.", "NOMOR", gedTextEdit, 0, 0, 40, 1, 2, 0, 0), _
            Array("Clm1", "NAMA TABEL", "TBL_NM", gedTextEdit, 0, 0, 160, 1, 1, 1, 0), _
            Array("Clm2", "KETERANGAN", "TBL_KET", gedTextEdit, 0, 0, 255, 1, 1, 0, 0), _
            Array("Clm3", "ACTIVE", "TBL_STT", gedCheckEdit, 0, 0, 123, 1, 3, 0, 0))
            '--------0---------1--------2-----------3--------4--5--6---7--8--9
End Sub

Private Sub GridFill()
Dim strqry As String
    MenuGrid
    strqry = "crud_componen_payment(1)"
    PrjSysGrid.GetGrid_Persensi dxDBGrid, CLM1, False, False, BND1, True, strqry, "NOMOR", False
    dxDBGrid.Dataset.Open
End Sub
Sub InitExample()
 dxDBGrid.Event = 1 'EGOnCustomDrawCell
 dxDBGrid.EventEnabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
lR = SetTopMostWindow(Me.hwnd, False)
End Sub

Private Sub txtNm_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    'MenuGrid
    GridFill
End If
End Sub



