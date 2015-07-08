VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Begin VB.Form FRM_ABSENSI_LOG 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Employe Log IN/OUT"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12255
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   12255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGridEditLog 
      Height          =   2775
      Left            =   0
      OleObjectBlob   =   "FRM_ABSENSI_LOG.frx":0000
      TabIndex        =   0
      Top             =   2280
      Width           =   12255
   End
   Begin XtremeSuiteControls.GroupBox GroupBox3 
      Height          =   2235
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12255
      _Version        =   851970
      _ExtentX        =   21616
      _ExtentY        =   3942
      _StockProps     =   79
      ForeColor       =   255
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.ListBox ListFinger 
         Height          =   1695
         Left            =   5280
         TabIndex        =   12
         Top             =   360
         Width           =   4575
         _Version        =   851970
         _ExtentX        =   8070
         _ExtentY        =   2990
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton cmdRptProcess 
         Height          =   495
         Index           =   0
         Left            =   9960
         TabIndex        =   3
         Top             =   1680
         Width           =   735
         _Version        =   851970
         _ExtentX        =   1296
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Add"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton cmdRptProcess 
         Height          =   495
         Index           =   1
         Left            =   10680
         TabIndex        =   10
         Top             =   1680
         Width           =   735
         _Version        =   851970
         _ExtentX        =   1296
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Edit"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton cmdRptProcess 
         Height          =   375
         Index           =   2
         Left            =   4440
         TabIndex        =   20
         Top             =   360
         Width           =   615
         _Version        =   851970
         _ExtentX        =   1085
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "FIND"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.DateTimePicker TglLogEdit 
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   2
         Top             =   360
         Width           =   1815
         _Version        =   851970
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   68
         Format          =   1
      End
      Begin XtremeSuiteControls.DateTimePicker TglLogEdit 
         Height          =   375
         Index           =   1
         Left            =   2640
         TabIndex        =   11
         Top             =   360
         Width           =   1815
         _Version        =   851970
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   68
         Format          =   1
      End
      Begin XtremeSuiteControls.PushButton cmdRptProcess 
         Height          =   495
         Index           =   3
         Left            =   11400
         TabIndex        =   24
         Top             =   1680
         Width           =   735
         _Version        =   851970
         _ExtentX        =   1296
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "delete"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton cmdRptProcess 
         Height          =   495
         Index           =   4
         Left            =   10440
         TabIndex        =   25
         Top             =   720
         Width           =   1095
         _Version        =   851970
         _ExtentX        =   1931
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Refresh"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin VB.Label LblLog 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Golongan"
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
         Index           =   4
         Left            =   1440
         TabIndex        =   19
         Top             =   1800
         Width           =   3375
      End
      Begin VB.Label LblLog 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Cabang"
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
         Index           =   3
         Left            =   1440
         TabIndex        =   18
         Top             =   1560
         Width           =   3375
      End
      Begin VB.Label LblLog 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Dept"
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
         Index           =   2
         Left            =   1440
         TabIndex        =   17
         Top             =   1320
         Width           =   3375
      End
      Begin VB.Label LblLog 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Index           =   1
         Left            =   1440
         TabIndex        =   16
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label LblLog 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "EmployeID"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   15
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Name :"
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
         Index           =   7
         Left            =   600
         TabIndex        =   14
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Golongan :"
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
         Left            =   240
         TabIndex        =   13
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cabang :"
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
         Index           =   6
         Left            =   480
         TabIndex        =   9
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Registed FingerID to :"
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
         Left            =   5280
         TabIndex        =   8
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Date 2"
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
         Left            =   2880
         TabIndex        =   7
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Date 1"
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
         Left            =   720
         TabIndex        =   6
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "EmployeID :"
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
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Dept :"
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
         Index           =   5
         Left            =   600
         TabIndex        =   4
         Top             =   1320
         Width           =   735
      End
   End
   Begin VB.Label LblLogFingerMachine 
      Caption         =   "UserName"
      Height          =   255
      Index           =   2
      Left            =   12360
      TabIndex        =   23
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label LblLogFingerMachine 
      Caption         =   "TerminalID"
      Height          =   255
      Index           =   1
      Left            =   12360
      TabIndex        =   22
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label LblLogFingerMachine 
      Caption         =   "FingerPrintID"
      Height          =   255
      Index           =   0
      Left            =   12360
      TabIndex        =   21
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "FRM_ABSENSI_LOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CLM, BND As Variant
Dim Qry0 As String, Qry1 As String, Qry2 As String, Qry3 As String, Qry4 As String
Dim i As Integer

Private Sub cmdRptProcess_Click(Index As Integer)
On Error GoTo ErrorLabel
    
    If Index = 0 Then
     '==============
     '---- INSERT --
     '==============
     On Error Resume Next
    
     '---- COPY --
      Dim TerminalID, FingerPrintID, UserName, FlagAbsence, DateTime, FunctionKey As String
      With dxDBGridEditLog.Dataset
        TerminalID = LblLogFingerMachine(1).Caption
        FingerPrintID = LblLogFingerMachine(0).Caption
        If .FieldValues("UserName") <> Null Then
            UserName = .FieldValues("UserName")
        Else
            UserName = LblLogFingerMachine(2).Caption
        End If
         FlagAbsence = "Manual"
      End With
      
      With dxDBGridEditLog.Dataset
        If cmdRptProcess(0).Caption = "Add" Then
            .Insert
            .FieldValues("TerminalID") = TerminalID
            .FieldValues("FingerPrintID") = FingerPrintID
            .FieldValues("UserName") = UserName
            .FieldValues("FlagAbsence") = FlagAbsence
            dxDBGridEditLog.Dataset.Refresh
            GridFill
        End If
      End With
    End If
    
    If Index = 1 Then
    On Error Resume Next
         '==============
         '---- EDIT --
         '==============
        If cmdRptProcess(1).Caption = "Edit" Then
            GridFill
            For i = 3 To 5
                dxDBGridEditLog.Columns(i).DisableEditor = False
                dxDBGridEditLog.Columns(i).Color = &HFFFFFF
                cmdRptProcess(1).Caption = "Accept"
            Next i
        Else
            dxDBGridEditLog.Dataset.Refresh
            For i = 3 To 5
                dxDBGridEditLog.Columns(i).DisableEditor = True
                dxDBGridEditLog.Columns(i).Color = &HCBFAFE
                cmdRptProcess(1).Caption = "Edit"
            Next i
            DateTimeUpdate
               
         
            'GridFill
        End If
    End If
    
    If Index = 2 Then
     '=====================
     '---- FIND DATE LOG --
     '=====================
      GridFill
    End If
    
    If Index = 3 Then
    '=======================
    '---- DELETE DATE LOG --
    '=======================
        dxDBGridEditLog.Dataset.Delete
        GridFill
    End If
    
    If Index = 4 Then
    '=======================
    '---- REFRESH DATE LOG --
    '=======================
        dxDBGridEditLog.Dataset.Cancel
        dxDBGridEditLog.Dataset.Refresh
        GridFill
    End If
 Exit Sub
ErrorLabel:
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"
End Sub

Private Sub Command1_Click()
End Sub

Private Sub dxDBGridEditLog_OnChangeColumn(ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal OldColumn As DXDBGRIDLibCtl.IdxGridColumn, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn)
On Error GoTo ErrorLabel
    If OldColumn.Index = 4 Then
        'With dxDBGridEditLog.Dataset
        '    On Error Resume Next
        '        dxDBGridEditLog.M.CloseEditor
                '.FieldValues("DateTime") = Format(DateValue(.FieldValues("tgl")) & " " & TimeValue(.FieldValues("waktu")), "dd/mm/yyyy hh:mm:ss")
               '.FieldValues("tgl") = Format(DateValue(.FieldValues("tgl")), "dd/mm/yyyy hh:mm:ss")
               '.FieldValues("DateTime") = Format(DateValue(.FieldValues("tgl")), "dd/mm/yyyy hh:mm:ss")
         '      .Refresh
         '       dxDBGridEditLog.M.CloseEditor
               
       '  End With
       DateTimeUpdate
    End If
    If OldColumn.Index = 5 Then
    '    With dxDBGridEditLog.Dataset
    '        On Error Resume Next
    '         dxDBGridEditLog.M.CloseEditor
    '         .FieldValues("waktu") = TimeValue(.FieldValues("waktu"))
    '        .FieldValues("DateTime") = Format(DateValue(.FieldValues("tgl")) & " " & TimeValue(.FieldValues("waktu")), "dd/mm/yyyy hh:mm:ss")
    '        .Refresh
    '         dxDBGridEditLog.M.CloseEditor
    '               '.FieldValues("DateTime") = .FieldValues("waktu")
    '    'dxDBGridEditLog.M.RefreshNodeValues
    '    End With
    DateTimeUpdate
    End If
Exit Sub
ErrorLabel:
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"
End Sub

Private Sub dxDBGridEditLog_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
'If Node.Id = 0 Then
'    dxDBGridEditLog.Dataset.Refresh
'    GridFill
'                For i = 3 To 5
'                dxDBGridEditLog.Columns(i).DisableEditor = False
'                dxDBGridEditLog.Columns(i).Color = &HFFFFFF
'                'cmdRptProcess(1).Caption = "Accept"
'            Next i'
'
'End If
End Sub

Private Sub Form_Load()
lR = SetTopMostWindow(Me.hwnd, True)
Main
'==============
'---- TEMBAK --
'LogProfile "WAN.MRY.000022","2015-02-01","2015-02-03"
TglLogEdit(0).Value = DateValue("2015-02-01")
TglLogEdit(1).Value = DateValue("2015-02-01")
'LogProfile "WAN.MRY.000022", TglLogEdit(0).Value, TglLogEdit(1).Value
'--------------
End Sub

'=====================================
'============= ptr.nov ===============
'===== MENU CULUMN BND GRID LODING ===
'=====================================
Private Sub MenuGrid()
    BND = Array("PERSONAL LOG VALUES")
    CLM = Array(Array("Clm0", "Terminal", "TerminalID", gedTextEdit, 0, 0, 100, 1, 0, 0, 0), _
                Array("Clm1", "FingerID", "FingerPrintID", gedTextEdit, 0, 0, 70, 1, 1, 0, 0), _
                Array("Clm2", "User Name", "UserName", gedTextEdit, 0, 0, 100, 1, 0, 0, 0), _
                Array("Clm3", "Key", "FunctionKey", gedLookupEdit, 0, 0, 80, 1, 1, 0, 0), _
                Array("Clm4", "Date", "tgl", gedDateEdit, 0, 0, 80, 1, 2, 0, 3), _
                Array("Clm5", "Time", "waktu", gedTimeEdit, 0, 0, 80, 1, 2, 0, 1), _
                Array("Clm6", "DateTime", "DateTime", gedTextEdit, 0, 0, 113, 1, 2, 0, 2), _
                Array("Clm7", "Editing", "Edited", gedTextEdit, 0, 0, 112, 1, 2, 0, 2), _
                Array("Clm8", "Status", "FlagAbsence", gedTextEdit, 0, 0, 54, 1, 2, 0, 0))
                '--------0---------1--------2-----------3--------4--5--6---7--8--9--10
                'TerminalID,FingerPrintID,UserName,FlagAbsence,DateTime,FunctionKey
End Sub
'==============================
'============= ptr.nov ========
'===== GRID COLUMN LOOKUP =====
'==============================
Private Sub LookupClm()
On Error GoTo ErrorLabel
    With dxDBGridEditLog.Columns.ColumnByName("Clm3").LookupColumn
                .LookupDataset.EnableControls
                '.LookupDataset.Close
                '.Dataset.Refresh
                .LookupDatasetType = dtADODataset
                .LookupDataset.ADODataset.ConnectionString = StrCon
                .LookupDataset.ADODataset.CursorLocation = clUseClient
                .LookupDataset.ADODataset.LockType = ltReadOnly         ' Check Permmission
                .LookupDataset.ADODataset.CommandType = cmdText
                .LookupDataset.ADODataset.CommandText = "SELECT FunctionKey,FunctionKeyNM FROM key_list" ' Like Join
                '.LookupDataset.Open
                .LookupKeyField = "FunctionKey"
                .LookupResultField = "FunctionKeyNM"
                .ListColumns = "KEY NAME"
                .ListFieldName = "FunctionKeyNM"
                .ListWidth = 800
                .DisplaySize = 400
    End With
    dxDBGridEditLog.Dataset.Open

Exit Sub
ErrorLabel:
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"
End Sub

Private Sub GridFill()
'=====================
'-- CRUD QUERY FILL --
'=====================
On Error GoTo ErrorLabel
    MenuGrid
    Qry0 = "crud_manual_datetimelog(1,'" & LblLog(0).Caption & _
                                    "','" & Format(TglLogEdit(0).Value, "yyyy-mm-dd") & _
                                    "',' " & Format(TglLogEdit(1).Value, "yyyy-mm-dd") & "')"
    'dxDBGridEditLog.Dataset.Refresh
    PrjSysGrid.GetGrid_Persensi dxDBGridEditLog, CLM, False, False, BND, True, Qry0, "idno", False
    LookupClm
    'dxDBGridEditLog.Dataset.Open
    
Exit Sub
ErrorLabel:
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"
End Sub

Private Sub DateTimeUpdate()
Dim qryDT As String
  qryDT = "crud_manual_datetimelog(2,'" & LblLog(0).Caption & _
                                    "','" & Format(TglLogEdit(0).Value, "yyyy-mm-dd") & _
                                    "',' " & Format(TglLogEdit(1).Value, "yyyy-mm-dd") & "')"
  OpRecStt2 qryDT, True
End Sub
 

'=========================
'-------- ptr.nov --------
'-------- PROFILE --------
'=========================
' LogProfile "WAN.MRY.000022","2015-02-01","2015-02-03"

Public Sub LogProfile(EmpID As String, TglLog1 As Date, TglLog2 As Date)
ListFinger.Clear
On Error Resume Next
    QryStt1 = "crud_manual_datetimelog(0,'" & EmpID & "','" & Format(TglLog1, "yyyy-mm-dd") & "','" & Format(TglLog2, "yyyy-mm-dd") & "')"
    OpRecStt1 QryStt1, True
    With rsStt1
        If Not .EOF Then
            LblLog(0).Caption = .Fields("KAR_ID").Value
            LblLog(1).Caption = .Fields("KAR_NM").Value
            LblLog(2).Caption = .Fields("DEP_NM").Value
            LblLog(3).Caption = .Fields("CAB_NM").Value
            LblLog(4).Caption = .Fields("TT_GRP_NM").Value
            LblLogFingerMachine(0).Caption = .Fields("FingerPrintID").Value
            LblLogFingerMachine(1).Caption = .Fields("TerminalID").Value
            LblLogFingerMachine(2).Caption = .Fields("KAR_NM").Value
            .MoveFirst
            Do Until .EOF
                    ListFinger.AddItem .Fields("FinngerReg").Value
                .MoveNext
            Loop
        End If
    End With
    Set rsStt1 = Nothing
GridFill
End Sub

Private Sub Form_Unload(Cancel As Integer)
FRM_ABSENSI.GridFill 1
End Sub

