VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Begin VB.Form PRESENSI_REPORT 
   Caption         =   "PRESENSI"
   ClientHeight    =   7980
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16440
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   16440
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.PushButton cmbProses 
      Height          =   375
      Left            =   9960
      TabIndex        =   6
      Top             =   120
      Width           =   1575
      _Version        =   851970
      _ExtentX        =   2778
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Proses"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cmbKar 
      Height          =   315
      Left            =   7440
      TabIndex        =   4
      Top             =   120
      Width           =   2175
      _Version        =   851970
      _ExtentX        =   3836
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "ALL"
   End
   Begin XtremeSuiteControls.DateTimePicker datePick 
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      _Version        =   851970
      _ExtentX        =   4048
      _ExtentY        =   661
      _StockProps     =   68
      CustomFormat    =   "dd, MMM yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker datePick 
      Height          =   375
      Index           =   1
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   2295
      _Version        =   851970
      _ExtentX        =   4048
      _ExtentY        =   661
      _StockProps     =   68
      CustomFormat    =   "dd, MMM yyyy"
      Format          =   3
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid 
      Height          =   3375
      Left            =   1080
      OleObjectBlob   =   "PRESENSI_REPORT.frx":0000
      TabIndex        =   7
      Top             =   840
      Width           =   15015
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGridReport 
      Height          =   3135
      Left            =   1080
      OleObjectBlob   =   "PRESENSI_REPORT.frx":0CB2
      TabIndex        =   8
      Top             =   4320
      Width           =   15015
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Index           =   2
      Left            =   6600
      TabIndex        =   5
      Top             =   120
      Width           =   975
      _Version        =   851970
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Karyawan"
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Index           =   1
      Left            =   3480
      TabIndex        =   3
      Top             =   120
      Width           =   495
      _Version        =   851970
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "To:"
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   615
      _Version        =   851970
      _ExtentX        =   1085
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Periode"
   End
End
Attribute VB_Name = "PRESENSI_REPORT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim defForm As String
Dim defCorp As String
Dim defCab As String
Dim defRow As Integer
Dim idValue As String
Dim SrcBy, defVal As String
Dim CLM1 As Variant

Private Sub cmbProses_Click()
MenuGrid 0
GridFill 0

MenuGrid 1
GridFill 1

cekPresensi "", True
End Sub

Private Sub Form_Load()
InfoServ
Main
GetUserInfo "erick"
datePick(0).Value = Now()
datePick(1).Value = Now()
FillCmbBox 1

End Sub

Private Sub MenuGrid(xMN As Integer)
Select Case xMN
    Case 0
    'ISI aRRAY Data LOG PRESENSI
    CLM1 = Array(Array("ClmPRESENSIID0", "NO.", "NO", gedTextEdit, 0, 0, 25, 1, 0, 0), _
                    Array("ClmPRESENSIID1", "FINGER ID", "FingerPrintID", gedTextEdit, 0, 0, 30, 1, 0, 1), _
                    Array("ClmPRESENSIID2", "TANGGAL", "DateLog", gedTextEdit, 0, 0, 30, 1, 0, 1), _
                    Array("ClmPRESENSIID3", "WAKTU", "TimeLog", gedTextEdit, 0, 0, 30, 1, 0, 1), _
                    Array("ClmPRESENSIID4", "KEY", "FunctionKey", gedTextEdit, 0, 0, 30, 1, 0, 1), _
                    Array("ClmPRESENSIID5", "DATETIME", "DateTime", gedTextEdit, 0, 0, 80, 1, 0, 1))
                    '--------0---------1--------2-----------3--------4--5--6---7--8--9
    Case 1
    'ISI aRRAY Data LOG PRESENSI
    CLM1 = Array(Array("ClmPRESENSIID0", "NO.", "NO", gedTextEdit, 0, 0, 25, 1, 0, 0), _
                    Array("ClmPRESENSIID1", "FINGER ID", "FingerPrintID", gedTextEdit, 0, 0, 30, 1, 0, 1), _
                    Array("ClmPRESENSIID3", "IN", "LogIN", gedTextEdit, 0, 0, 30, 1, 0, 1), _
                    Array("ClmPRESENSIID2", "CHECK", "LogCHECK", gedTextEdit, 0, 0, 30, 1, 0, 1), _
                    Array("ClmPRESENSIID4", "OUT", "LogOUT", gedTextEdit, 0, 0, 30, 1, 0, 1), _
                    Array("ClmPRESENSIID5", "DURATION", "DURATION", gedTextEdit, 0, 0, 30, 1, 0, 1), _
                    Array("ClmPRESENSIID6", "DURATION", "TYPE", gedTextEdit, 0, 0, 30, 1, 0, 1), _
                    Array("ClmPRESENSIID7", "COST", "SUB_TOTAL", gedTextEdit, 0, 0, 80, 1, 0, 1))
                    '--------0---------1--------2-----------3--------4--5--6---7--8--9
End Select
End Sub

Private Sub GridFill(xGR As Integer)
Dim strqry As String
Dim StrChk(2) As String
Dim SelProp(10) As String
Dim Oprtn(10) As String

 'ISI grid Data Barang - tabel Barang
    'StrQRY = "pro_brg_item('" & txtNm.Text & "','" & defCorp & "','" & SrcBy & "')"
Select Case xGR
Case 0
    strqry = "pro_presensi_item('" & prjSysID.GetFingerID_Karyawan(cmbKar.Text) & "','" & _
              Format(datePick(0).Value, "YYYY-MM-DD 00:00:00") & "','" & _
              Format(datePick(1).Value, "YYYY-MM-DD 23:59:59") & "')"
              
        Clipboard.SetText strqry
        MsgBox strqry
        PrjSysGrid.GetBaseGridRoPoID dxDBGrid, CLM1, strqry, "FingerPrintID", False
        dxDBGrid.Columns.ColumnByFieldName("DateLog").DisplayFormat = "YYYY-MM-DD"
        dxDBGrid.Columns.ColumnByFieldName("TimeLog").DisplayFormat = "hh:mm:ss"
        dxDBGrid.Columns.ColumnByFieldName("DateTime").DisplayFormat = "YYYY-MM-DD hh:mm:ss"

Case 1
    strqry = "pro_presensi_report_item('" & prjSysID.GetFingerID_Karyawan(cmbKar.Text) & "')"
        Clipboard.SetText strqry
        MsgBox strqry
        PrjSysGrid.GetBaseGridRoPoID dxDBGridReport, CLM1, strqry, "FingerPrintID", False
        dxDBGridReport.Columns.ColumnByFieldName("LogIN").DisplayFormat = "YYYY-MM-DD hh:mm:ss"
        dxDBGridReport.Columns.ColumnByFieldName("LogOUT").DisplayFormat = "YYYY-MM-DD hh:mm:ss"
        dxDBGridReport.Columns.ColumnByFieldName("LogCHECK").DisplayFormat = "YYYY-MM-DD hh:mm:ss"
        
End Select
End Sub
Sub InitExample()
 dxDBGrid.Event = 1 'EGOnCustomDrawCell
 dxDBGrid.EventEnabled = True
End Sub
Private Sub FillCmbBox(N As Integer)
Dim Itm() As String
Dim dStr As String
Dim K As Integer

Select Case N
Case Is = 1
    dStr = "SELECT KAR_NM FROM karyawan"
        prjSysID.GetListComboByQRY cmbKar, dStr, False
        cmbKar.AddItem "All"
        cmbKar.ListIndex = 0
    'dStr = "SELECT CAT_NM FROM tab_barang_cat WHERE CORP_ID='" & defCorp & "' AND CAT_NM<>'_NONE'"
     '   prjSysID.GetListComboByQRY cmbFilter(1), dStr, False
     '   cmbFilter(1).AddItem ""
     '   cmbFilter(1).ListIndex = 0
    'dStr = "SELECT SAT_NM FROM tab_barang_satuan WHERE CORP_ID='" & defCorp & "' AND SAT_NM<>'_NONE'"
     '   prjSysID.GetListComboByQRY cmbFilter(2), dStr, False
     '   cmbFilter(2).AddItem ""
     '   cmbFilter(2).ListIndex = 0
    'dStr = "SELECT SPL_NM FROM tab_suplier WHERE SPL_NM<>'_NONE'" ' WHERE CORP_ID='" & prjSysID.GetKodeCorp(CmbBrgEdit(1).Text) & "'"
     '   prjSysID.GetListComboByQRY cmbFilter(3), dStr, False
      '  cmbFilter(3).AddItem ""
       ' cmbFilter(3).ListIndex = 0
End Select
End Sub
Sub cekPresensi(xStrQry As String, Optional isProcedure As Boolean)
Dim TimeIn As String 'Date
Dim TimeOut As String 'Date
Dim TimeDurH As String 'Date
Dim TimeDurM As String 'Date
Dim TimeDurS As String 'Date

xStrQry = "pro_presensi_item('" & prjSysID.GetFingerID_Karyawan(cmbKar.Text) & "','" & _
              Format(datePick(0).Value, "YYYY-MM-DD 00:00:00") & "','" & _
              Format(datePick(1).Value, "YYYY-MM-DD 23:59:59") & "')"
isProcedure = True

On Error Resume Next
Dim oRS As ADODB.Recordset
  Set oRS = New ADODB.Recordset
  
    If isProcedure Then oRS.Open xStrQry, conMain, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc Else oRS.Open xStrQry, conMain, adOpenForwardOnly, adLockReadOnly, adCmdText
    
      PrjSysTrig.RunSqlStr conMain, "truncate  tab_report_presensi;", False
      
      Do While Not oRS.EOF
       With oRS
           If .Fields("FunctionKey").Value = "1" Then
                 PrjSysTrig.RunSqlStr conMain, "INSERT INTO tab_report_presensi(FingerPrintID,LogIN) VALUES('" & .Fields("FingerPrintID").Value & "','" & Format(.Fields("DateTime").Value, "YYYY-MM-DD hh:mm:ss") & "')", False
                 TimeIn = Format(.Fields("DateTime").Value, "YYYY-MM-DD hH:Mm:sS")
           End If
           
           If .Fields("FunctionKey").Value = "2" Then
                 PrjSysTrig.RunSqlStr conMain, "UPDATE tab_report_presensi SET LogOUT='" & Format(.Fields("DateTime").Value, "YYYY-MM-DD hh:mm:ss") & "' WHERE LogIN='" & TimeIn & "'", False
                 TimeOut = Format(.Fields("DateTime").Value, "YYYY-MM-DD hh:mm:ss")
                 'TimeDurH = DateDiff("h", TimeIn, TimeOut)
                 TimeDurn = DateDiff("n", TimeIn, TimeOut)
                 'TimeDurS = DateDiff("s", TimeIn, TimeOut)
                 
                 'MsgBox TimeIn & "#" & TimeOut & "#" & TimeDur
                 PrjSysTrig.RunSqlStr conMain, "UPDATE tab_report_presensi SET DURATION='" & TimeDurn / 60 & "' WHERE LogIN='" & TimeIn & "'", False
           End If
        .MoveNext
       End With
      Loop
  
  oRS.Close
  Set oRS = Nothing
  
xCombo.ListIndex = 0

End Sub
