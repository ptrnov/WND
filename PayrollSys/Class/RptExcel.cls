VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim trecs As Double
Dim p As Integer

Private Function NmHari(Tgl As Date)
If (Day(Tgl)) = 1 Then
    NmHari = "Minggu"
ElseIf (Day(Tgl)) = 2 Then
    NmHari = "Senin"
ElseIf (Day(Tgl)) = 3 Then
    NmHari = "Selasa"
ElseIf (Day(Tgl)) = 4 Then
    NmHari = "Rabu"
ElseIf (Day(Tgl)) = 5 Then
    NmHari = "Kamis"
ElseIf (Day(Tgl)) = 6 Then
    NmHari = "Jumat"
ElseIf (Day(Tgl)) = 7 Then
    NmHari = "Sabtu"
End If

End Function

 
'--------------------------------------
'------ Persensi_OverTime_Daily  ------
'--------------------------------------
Public Sub Persensi_SesionClose()
Dim ApExcel  As Excel.Application, oWkbk As Excel.Workbook, oSheet As Excel.Worksheet
Dim i, J, a, x, h As Long
Dim Qry1 As String
Dim Qry2 As String
Dim STotalStr As String
Set ApExcel = CreateObject("Excel.Application")
    ApExcel.Visible = True
    Set oWkbk = ApExcel.Workbooks.Add
        Set oSheet = oWkbk.Sheets("sheet1")
            oSheet.Name = "Presensi_Session_Close"
            oSheet.Cells.Font.Name = "Utsaah"
            oSheet.Cells.Justify
            oSheet.Cells.Font.Color = &H8000000F
            oSheet.Columns.AutoFit
            oSheet.Cells(1, 1) = "PT. WANINDO PRIMA"
            oSheet.Cells(2, 1) = "PRESENSI SESSION CLOSE"
            Set oSheet = oWkbk.Sheets(1)
             For i = 1 To 19
                For a = 4 To 5
                    oSheet.Cells(a, i).Interior.ColorIndex = 40
                    oSheet.Cells(a, i).Borders.LineStyle = 1
                    oSheet.Cells(a, i).Borders.Color = &H8000000F
                    oSheet.Cells(a, i).Font.Color = &H8000000F
                    oSheet.Cells(a, i).BorderAround , , &HC0C0C0
                    oSheet.Cells(a, i).Font.Bold = True
                    oSheet.Cells(a, i).Font.Size = 9
                    oSheet.Cells(a, i).HorizontalAlignment = 3
                    'Left:    -4131  =2
                    'Center:  -4108  =3
                    'Right:   -4152 =1
                Next a
            Next i
             oSheet.Range("A4", "D4").Merge
             oSheet.Cells(4, 1) = "EMPLOYE DATE TIME"
             '[A]------Marge -----
             oSheet.Range("E4", "M4").Merge
             oSheet.Cells(4, 5) = "EMPLOYE PRESENSI"
             '[B]------Marge -----
             oSheet.Range("N4", "O4").Merge
             oSheet.Cells(4, 14) = "OT LEVEL 1"
             '[C]------Marge -----
             oSheet.Range("P4", "Q4").Merge
             oSheet.Cells(4, 16) = "OT LEVEL 2"
             '[D]------Marge -----
             oSheet.Range("R4", "S4").Merge
             oSheet.Cells(4, 18) = "OT LEVEL 3"
                        'Culumn Alignment
            
            oSheet.Cells(5, 1) = "Date"
            oSheet.Cells(5, 1).ColumnWidth = 8
            oSheet.Cells(5, 2) = "Employe Name"
            oSheet.Cells(5, 2).ColumnWidth = 12
            oSheet.Cells(5, 3) = "DTime In."
            oSheet.Cells(5, 3).ColumnWidth = 9
            oSheet.Cells(5, 4) = "DTime Out:"
            oSheet.Cells(5, 4).ColumnWidth = 9
            '[A]------Marge -----
            
            oSheet.Cells(5, 5) = "Work Time"
            oSheet.Cells(5, 6) = "Day Active"
            oSheet.Cells(5, 7) = "Holiday"
            oSheet.Cells(5, 8) = "Absent"
            oSheet.Cells(5, 9) = "Excpt"
            oSheet.Cells(5, 10) = "Late"
            oSheet.Cells(5, 11) = "Early"
            oSheet.Cells(5, 12) = "OT Depan"
            oSheet.Cells(5, 13) = "OT level"
            '[B]------Marge -----
            oSheet.Cells(5, 14) = "Over Time"
            oSheet.Cells(5, 15) = "Value"
            '[C]------Marge -----
            oSheet.Cells(5, 16) = "Over Time"
            oSheet.Cells(5, 17) = "Value"
            '[D]------Marge -----
            oSheet.Cells(5, 18) = "Over Time"
            oSheet.Cells(5, 19) = "Value"
            
            'Culumn Size
            For x = 5 To 19
                 oSheet.Cells(5, x).ColumnWidth = 6
            Next x
            


            oSheet.Cells.Rows(5).AutoFilter
            Dim Qry_OTDaily As String
            Dim Qry_OTDaily_Sub As String
            
            Qry_OTDaily = "Vsr1('Persensi_OverTime_Daily=2015-01-17=2015-02-18=0=0=0=0')"
            OpnRecRpt1 Qry_OTDaily, True
            With rsRpt1
                i = 6
                If .RecordCount <> 0 Then
                    .MoveFirst
                    Do Until .EOF
                        For x = 1 To 19
                            oSheet.Cells(i, x).Font.Size = 9
                        Next x
                         For x = 1 To 19
                             oSheet.Cells(i, x).HorizontalAlignment = 3
                         Next x
                        oSheet.Cells(i, 2).HorizontalAlignment = 2
                        oSheet.Cells(i, 15).HorizontalAlignment = 2
                        oSheet.Cells(i, 17).HorizontalAlignment = 2
                        oSheet.Cells(i, 19).HorizontalAlignment = 2

                        oSheet.Cells(i, 1) = "2015-01-30"
                        
                        oSheet.Cells(i, 2) = .Fields("Kar_ID").Value
                        oSheet.Cells(i, 3) = "2015-01-30 08:00"
                        oSheet.Cells(i, 4) = "2015-01-30 08:00"
                        oSheet.Cells(i, 5) = "08:00"
                        oSheet.Cells(i, 6) = "100"
                        oSheet.Cells(i, 7) = "100"
                        oSheet.Cells(i, 8) = "100"
                        oSheet.Cells(i, 9) = "100"
                        oSheet.Cells(i, 10) = "10:00"
                        oSheet.Cells(i, 11) = "10:00"
                        oSheet.Cells(i, 12) = "10:00"
                        oSheet.Cells(i, 13) = "HARD"
                        oSheet.Cells(i, 14) = "08:00"
                        oSheet.Cells(i, 15) = "7.6"
                        oSheet.Cells(i, 16) = "08:00"
                        oSheet.Cells(i, 17) = "6.6"
                        oSheet.Cells(i, 18) = "08:00"
                        oSheet.Cells(i, 19) = "7.6"
                      i = i + 1
                    .MoveNext
                    Loop
                End If
            End With
            Set rsRpt1 = Nothing
'oSheet.Cells.Columns.AutoFit
Set ApExcel = Nothing
End Sub

'=============================
'---------- PRESENSI  --------
'------ Persensi Daily  ------
'=============================
Public Sub Persensi_Daily(STgl As String, ETgl As String, DepId As Integer, KarId As String, GrpId As Integer, CabId As String, MenuRpt As String, frm_actv As Integer)
Dim strqry As String
Dim ApExcel  As Excel.Application, oWkbk As Excel.Workbook, oSheet As Excel.Worksheet
Dim i, J, a, x, y As Long
Dim Qry1 As String
Dim Qry2 As String
Dim STotalStr As String
Set ApExcel = CreateObject("Excel.Application")
    ApExcel.Visible = True
    Set oWkbk = ApExcel.Workbooks.Add
        Set oSheet = oWkbk.Sheets("sheet1")
            oSheet.Name = "Daily Persensi"
            oSheet.Cells.Font.Name = "Utsaah"
            oSheet.Cells.Justify
            oSheet.Cells.Font.Color = &H8000000F
            oSheet.Columns.AutoFit
            oSheet.Cells(1, 1) = "PT. WANINDO PRIMA"
            oSheet.Cells(2, 1) = " DAILY PRESENSI "
            Set oSheet = oWkbk.Sheets(1)
         


            'oSheet.Cells.Rows(5).AutoFilter
            Dim Qry_OTDaily As String
            Dim Qry_OTDaily_Sub As String
            Dim menuExcel As String
            Dim Kondisi_DepID As Integer
            DepId = 0
            menuExcel = MenuRpt & "_Excel"
            
            'Qry_OTDaily = "Vsr1('Persensi_Daily=2015-01-17=2015-02-18=0=0=0=0')"
           Qry_OTDaily = "Vsr_Report('DailyPresensi=" & STgl & "=" & ETgl & "=" & DepId & "=" & KarId & "=" & GrpId & "=" & CabId & "')"
           'Qry_OTDaily = "Vsr1('Persensi_Daily=" & stgl & "=" & etgl & " = " & DepId & " = " & KarId & " = " & GrpId & " = " & CabId & "')"
            OpnRecRpt1 Qry_OTDaily, True
            With rsRpt1
                i = 3
                If .RecordCount <> 0 Then
                        
                    .MoveFirst
                    Do Until .EOF
                        trecs = trecs + 1
                    .MoveNext
                    Loop
                    
                    If frm_actv = 1 Then '=== PRESENSI ===
                        With FRM_ABSENSI_REPORT
                            .ProgressBar1.Visible = True
                            .ProgressBar1.Min = 0
                            .ProgressBar1.Max = trecs
                        End With
                    ElseIf frm_actv = 2 Then '=== PAYROLL ===
                        With FRM_PAYROLL_REPORT
                            .ProgressBar1.Visible = True
                            .ProgressBar1.Min = 0
                            .ProgressBar1.Max = trecs
                        End With
                    ElseIf frm_actv = 3 Then '=== ACCOUNTING ===
                        With FRM_ABSENSI_REPORT
                            .ProgressBar1.Visible = True
                            .ProgressBar1.Min = 0
                            .ProgressBar1.Max = trecs
                        End With
                    End If
                    
                    .MoveFirst
                    Do Until .EOF
                        If Kondisi_DepID <> (.Fields("DEP_ID").Value) Then
                            i = i + 1
                            oSheet.Cells(i, 18) = .Fields("DEP_NM").Value
                            oSheet.Cells(i, 1) = "Time Table : " & .Fields("GRP_NM").Value
                            i = i + 1
                                For x = 1 To 19
                                    For y = i To i + 1
                                        oSheet.Cells(y, x).Interior.ColorIndex = 40
                                        oSheet.Cells(y, x).Borders.LineStyle = 1
                                        oSheet.Cells(y, x).Borders.Color = &H8000000F
                                        oSheet.Cells(y, x).Font.Color = &H8000000F
                                        oSheet.Cells(y, x).BorderAround , , &HC0C0C0
                                        oSheet.Cells(y, x).Font.Bold = True
                                        oSheet.Cells(y, x).Font.Size = 9
                                        oSheet.Cells(y, x).HorizontalAlignment = 3
                                        'Left:    -4131  =2
                                        'Center:  -4108  =3
                                        'Right:   -4152 =1
                                        
                                    Next y
                                Next x
                                Dim mrg1, mrg2, mrg3, mrg4, mrg5, mrg6, mrg7, mrg8, mrg9, mrg10, mrg11, mrg12 As String
                                mrg1 = "A" & i
                                mrg2 = "D" & i
                                mrg3 = "E" & i
                                mrg4 = "K" & i
                                mrg5 = "L" & i
                                mrg6 = "M" & i
                                mrg7 = "N" & i
                                mrg8 = "O" & i
                                mrg9 = "P" & i
                                mrg10 = "Q" & i
                                mrg11 = "R" & i
                                mrg12 = "S" & i
                                 oSheet.Range(mrg1, mrg2).Merge
                                 oSheet.Cells(i, 1) = "EMPLOYE DATE TIME"
                                 '[A]------Marge -----
                                 oSheet.Range(mrg3, mrg4).Merge
                                 oSheet.Cells(i, 5) = "EMPLOYE PRESENSI"
                                 '[B]------Marge -----
                                 oSheet.Range(mrg5, mrg6).Merge
                                 oSheet.Cells(i, 12) = "OT DEPAN"
                                 '[C]------Marge -----
                                 oSheet.Range(mrg7, mrg8).Merge
                                 oSheet.Cells(i, 14) = "OT LEVEL 1"
                                 '[D]------Marge -----
                                 oSheet.Range(mrg9, mrg10).Merge
                                 oSheet.Cells(i, 16) = "OT LEVEL 2"
                                 '[E]------Marge -----
                                 oSheet.Range(mrg11, mrg12).Merge
                                 oSheet.Cells(i, 18) = "OT LEVEL 3"
                                            'Culumn Alignment
                                
                                oSheet.Cells(i + 1, 1) = "Date"
                                oSheet.Cells(i + 1, 1).ColumnWidth = 8
                                oSheet.Cells(i + 1, 2) = "Employe Name"
                                oSheet.Cells(i + 1, 2).ColumnWidth = 12
                                oSheet.Cells(i + 1, 3) = "DTime In."
                                oSheet.Cells(i + 1, 3).ColumnWidth = 9
                                oSheet.Cells(i + 1, 4) = "DTime Out:"
                                oSheet.Cells(i + 1, 4).ColumnWidth = 9
                                '[A]------Marge -----
                                
                                oSheet.Cells(i + 1, 5) = "Work Time"
                                oSheet.Cells(i + 1, 6) = "Day Active"
                                oSheet.Cells(i + 1, 7) = "Holiday"
                                oSheet.Cells(i + 1, 8) = "Absent"
                                oSheet.Cells(i + 1, 9) = "Excpt"
                                oSheet.Cells(i + 1, 10) = "Late"
                                oSheet.Cells(i + 1, 11) = "Early"
                                oSheet.Cells(i + 1, 12) = "Over Time"
                                oSheet.Cells(i + 1, 13) = "OT Hari"
                                '[B]------Marge -----
                                oSheet.Cells(i + 1, 14) = "Over Time"
                                oSheet.Cells(i + 1, 15) = "OT Hari"
                                '[C]------Marge -----
                                oSheet.Cells(i + 1, 16) = "Over Time"
                                oSheet.Cells(i + 1, 17) = "OT Hari"
                                '[D]------Marge -----
                                oSheet.Cells(i + 1, 18) = "Over Time"
                                oSheet.Cells(i + 1, 19) = "OT Hari"
                                
                                'Culumn Size
                                For x = 5 To 19
                                     oSheet.Cells(i + 1, x).ColumnWidth = 6
                                Next x
                                i = i + 2
                        End If
                         
                        
                        
                        For x = 1 To 19
                           ' For y = i To i + 1
                                'oSheet.Cells(i, x).Interior.ColorIndex = 40
                                oSheet.Cells(i, x).Borders.LineStyle = 1
                            'Next y
                        Next x
                                
                        For x = 1 To 19
                            oSheet.Cells(i, x).Font.Size = 9
                        Next x
                         For x = 1 To 19
                             oSheet.Cells(i, x).HorizontalAlignment = 3
                         Next x
                        oSheet.Cells(i, 2).HorizontalAlignment = 2
                        oSheet.Cells(i, 15).HorizontalAlignment = 2
                        oSheet.Cells(i, 17).HorizontalAlignment = 2
                        oSheet.Cells(i, 19).HorizontalAlignment = 2
                        
                        
                        'If (.Fields("TGL_RUN").Value) <> Null Then
                            oSheet.Cells(i, 1) = Format(.Fields("TGL_RUN").Value, "dd/mm/yyyy") '"2015-01-30"
                       ' Else
                            'oSheet.Cells(i, 1) = ""
                        'End If
                        oSheet.Cells(i, 2) = .Fields("Kar_NM").Value
                        
                        'If .Fields("DT_IN").Value <> Null Then
                            oSheet.Cells(i, 3) = .Fields("DT_IN").Value '"2015-01-30 08:00"
                        'Else
                         '   oSheet.Cells(i, 3) = ""
                        'End If
                        
                        'If .Fields("DT_OUT").Value <> Null Then
                            oSheet.Cells(i, 4) = .Fields("DT_OUT").Value '"2015-01-30 08:00"
                        'Else
                        '    oSheet.Cells(i, 4) = ""
                        'End If
                        
                        'If .Fields("WORK_TIME").Value <> Null Then
                            oSheet.Cells(i, 5) = Format(.Fields("WORK_TIME").Value, "hh:mm:ss") '"08:00"
                        'Else
                       '     oSheet.Cells(i, 5) = ""
                        'End If
                        
                        'If .Fields("DAY_ACTIVE").Value <> Null Then
                            oSheet.Cells(i, 6) = .Fields("DAY_ACTIVE").Value '"100"
                        'Else
                        '         oSheet.Cells(i, 6) = ""
                        'End If
                                                
                        'If .Fields("HOLIDAY").Value <> Null Then
                            oSheet.Cells(i, 7) = .Fields("HOLIDAY").Value '"100"
                        'Else
                        '        oSheet.Cells(i, 7) = ""
                        'End If
                        
                        'If .Fields("DAY_ABSENT").Value <> Null Then
                            oSheet.Cells(i, 8) = .Fields("DAY_ABSENT").Value '"100"
                        'Else
                        '        oSheet.Cells(i, 8) = ""
                        'End If
                        
                        'If .Fields("DAY_EXCP_ID").Value <> Null Then
                            oSheet.Cells(i, 9) = .Fields("DAY_EXCP_ID").Value '"100"
                        'Else
                        '        oSheet.Cells(i, 9) = ""
                        'End If
                                                
                        'If .Fields("LATE").Value <> Null Then
                            oSheet.Cells(i, 10) = Format(.Fields("LATE").Value, "hh:mm:ss") '"10:00"
                        'Else
                        '    oSheet.Cells(i, 10) = ""
                        'End If
                                                
                        'If .Fields("EARLY").Value <> Null Then
                            oSheet.Cells(i, 11) = Format(.Fields("EARLY").Value, "hh:mm:ss") '"10:00"
                        'Else
                        '    oSheet.Cells(i, 11) = ""
                        'End If
                                                
                        'If .Fields("OT_DPN").Value <> Null Then
                            oSheet.Cells(i, 12) = Format(.Fields("OT_DPN").Value, "hh:mm:ss") '"10:00"
                        'Else
                        '    oSheet.Cells(i, 12) = ""
                        'End If
                                                
                        'If .Fields("OT_DPN_HARI").Value <> Null Then
                            oSheet.Cells(i, 13) = .Fields("OT_DPN_HARI").Value '"HARD"
                        'Else
                        '    oSheet.Cells(i, 13) = ""
                        'End If
                                                
                        'If .Fields("OT1_BLK").Value <> Null Then
                            oSheet.Cells(i, 14) = Format(.Fields("OT1_BLK").Value, "hh:mm:ss") '"08:00"
                        'Else
                        '    oSheet.Cells(i, 14) = ""
                        'End If
                                                
                        'If .Fields("OT1_BLK_HARI").Value <> Null Then
                                oSheet.Cells(i, 15) = .Fields("OT1_BLK_HARI").Value '"7.6"
                       ' Else
                        '        oSheet.Cells(i, 15) = ""
                        'End If
                                                
                        'If .Fields("OT2_BLK").Value <> Null Then
                                oSheet.Cells(i, 16) = Format(.Fields("OT2_BLK").Value, "hh:mm:ss") '"08:00"
                        'Else
                        '        oSheet.Cells(i, 16) = ""
                        'End If
                                                
                        'If .Fields("OT2_BLK_HARI").Value <> Null Then
                                oSheet.Cells(i, 17) = .Fields("OT2_BLK_HARI").Value '"6.6"
                        'Else
                        '        oSheet.Cells(i, 17) = ""
                        'End If
                
                        'If .Fields("OT3_BLK").Value <> Null Then
                                oSheet.Cells(i, 18) = Format(.Fields("OT3_BLK").Value, "hh:mm:ss") '"08:00"
                        'Else
                        '        oSheet.Cells(i, 18) = ""
                        'End If
                                                
                        'If .Fields("OT3_BLK_HARI").Value <> Null Then
                                oSheet.Cells(i, 19) = .Fields("OT3_BLK_HARI").Value '"7.6"
                        'Else
                        '        oSheet.Cells(i, 19) = ""
                        'End If
                        
                        'If .Fields("DEP_ID").Value <> Null Then
                            Kondisi_DepID = .Fields("DEP_ID").Value
                        'Else
                             'Kondisi_DepID = 0
                        'End If
                     If frm_actv = 1 Then '=== PRESENSI ===
                        FRM_ABSENSI_REPORT.ProgressBar1.Value = i
                     ElseIf frm_actv = 2 Then '=== PAYROLL ===
                        FRM_PAYROLL_REPORT.ProgressBar1.Value = i
                     ElseIf frm_actv = 3 Then '=== ACCOUNTING ===
                        FRM_ABSENSI_REPORT.ProgressBar1.Value = i
                     End If

                        
                      i = i + 1
                    .MoveNext
                    Loop
                End If
            End With
            Set rsRpt1 = Nothing
'oSheet.Cells.Columns.AutoFit
Set ApExcel = Nothing
End Sub




'=============================
'------ PAYROLL WEEKLY  ------
'=============================
Public Sub PAYROLL_WEEKLY(STgl As String, ETgl As String, DepId As Integer, KarId As String, GrpId As Integer, CabId As String, MenuRpt As String, frm_actv As Integer)
Dim ApExcel  As Excel.Application, oWkbk As Excel.Workbook, oSheet As Excel.Worksheet
Dim i, J, a, x, h As Long
Dim Qry1 As String
Dim Qry2 As String
Dim STotalStr As String
Set ApExcel = CreateObject("Excel.Application"): ApExcel.Visible = True
Set oWkbk = ApExcel.Workbooks.Add

If True Then '===== SHEET Pertma ==============
On Error Resume Next
       Set oSheet = oWkbk.Sheets.Add
            oSheet.Name = "PAYROLL_WEEKLY"
            oSheet.Cells.Font.Name = "Utsaah"
            oSheet.Cells.Justify
            oSheet.Cells.Font.Color = &H8000000F
            'oSheet.Columns.AutoFit
            oSheet.Cells(1, 1) = "PT. WANINDO PRIMA"
            oSheet.Cells(2, 1) = "PAYROLL WEEKLY SESSION CLOSE"
             oSheet.Cells.Font.Size = 9
            'Set oSheet = oWkbk.Sheets(1)
            Dim Qry_OTDaily As String
            Dim Qry_OTDaily_Sub As String
            Dim Stt_DepID As Integer
            trecs = 0
            Qry_OTDaily = "Vsr_Report('Payroll_WeeklyExcel=" & STgl & "=" & ETgl & "=" & DepId & "=" & KarId & "=" & GrpId & "=" & CabId & "')"
            'Payroll_WeeklyExcel=2015-03-01=2015-03-06=0=0=0=0'
            'Menu,sDateRun,eDateRun,DeptID,KarId,GrpId,CabID
            'STgl,ETgl,DepId, KarId, GrpId,CabId,MenuRpt
            OpnRecRpt1 Qry_OTDaily, True
            Stt_DepID = 0
            i = 4
            With rsRpt1
               ' i = 3
                If .RecordCount <> 0 Then
                    .MoveFirst
                    Do Until .EOF
                        trecs = trecs + 1
                    .MoveNext
                    Loop
                    
                    '=====================
                    '--- PROGRESS BAR ----
                    '=====================
                    If frm_actv = 1 Then '=== PRESENSI ===
                        With FRM_ABSENSI_REPORT
                            .ProgressBar1.Visible = True
                            .ProgressBar1.Min = 0
                            .ProgressBar1.Max = trecs
                        End With
                    ElseIf frm_actv = 2 Then '=== PAYROLL ===
                        With FRM_PAYROLL_REPORT
                            .ProgressBar1.Visible = True
                            .ProgressBar1.Min = 0
                            .ProgressBar1.Max = trecs
                        End With
                    ElseIf frm_actv = 3 Then '=== ACCOUNTING ===
                        With FRM_ABSENSI_REPORT
                            .ProgressBar1.Visible = True
                            .ProgressBar1.Min = 0
                            .ProgressBar1.Max = trecs
                        End With
                    End If

                    .MoveFirst
                    Do Until .EOF
                    '======================
                    '--- Group DEP --------
                    '======================
                    If Stt_DepID <> (.Fields("DEP_ID").Value) Then
                        Stt_DepID = .Fields("DEP_ID").Value
                        oSheet.Cells(3, 1) = "PERIODE : " & .Fields("sTgl").Value & "  TO  " & .Fields("eTgl").Value
                        Dim Tgl As Date
                        Dim TglInt, clmInt, HdrRow As Integer
                        TglInt = 4  'COLOMN HEADER DATE LOOP
                        i = i + 1   'ROW

                        
                        For Tgl = .Fields("sTgl").Value To .Fields("eTgl").Value
                             For clmInt = 1 To TglInt + 7 'TglInt + 2=tgl loop,
                                For HdrRow = i To (i + 2)
                                    oSheet.Cells(HdrRow, clmInt).Interior.ColorIndex = 15
                                    oSheet.Cells(HdrRow, clmInt).Borders.LineStyle = 1
                                    oSheet.Cells(HdrRow, clmInt).Borders.Color = &H8000000F
                                    oSheet.Cells(HdrRow, clmInt).Font.Color = &H8000000F
                                    oSheet.Cells(HdrRow, clmInt).BorderAround , , &HC0C0C0
                                    oSheet.Cells(HdrRow, clmInt).Font.Bold = True
                                    oSheet.Cells(HdrRow, clmInt).Font.Size = 9
                                    oSheet.Cells(HdrRow, clmInt).HorizontalAlignment = 3
                                    'Left:    -4131  =2
                                    'Center:  -4108  =3
                                    'Right:   -4152 =1
                                Next HdrRow
                            Next clmInt

                            '=== TANGGAL ===
                            oSheet.Range(oSheet.Cells(i, TglInt), oSheet.Cells(i, TglInt + 2)).Merge
                            oSheet.Cells(i, TglInt) = "'" & Tgl & "'" 'RowHeight=10.5
                            oSheet.Cells(i, TglInt).RowHeight = 10.5
                            oSheet.Cells(i, TglInt).VerticalAlignment = 2
                            '=== HARI ====
                            oSheet.Range(oSheet.Cells(i + 1, TglInt), oSheet.Cells(i + 1, TglInt + 2)).Merge
                            oSheet.Cells(i + 1, TglInt) = NmHari(Tgl)
                            oSheet.Cells(i + 1, TglInt).RowHeight = 10.5
                            oSheet.Cells(i + 1, TglInt).VerticalAlignment = 2
                            '==== FIELD ===
                            oSheet.Cells(i + 2, TglInt) = "DPN"
                            oSheet.Cells(i + 2, TglInt).ColumnWidth = 3.29
                            oSheet.Cells(i + 2, TglInt).RowHeight = 10.5
                            oSheet.Cells(i + 2, TglInt).VerticalAlignment = 2
                            oSheet.Cells(i + 2, TglInt + 1) = "PGI"
                            oSheet.Cells(i + 2, TglInt + 1).ColumnWidth = 3.29
                            oSheet.Cells(i + 2, TglInt + 1).VerticalAlignment = 2
                            oSheet.Cells(i + 2, TglInt + 2) = "BLK"
                            oSheet.Cells(i + 2, TglInt + 2).ColumnWidth = 3.29
                            oSheet.Cells(i + 2, TglInt + 2).VerticalAlignment = 2

                            TglInt = TglInt + 3 '--- JARAK MERGE
                        Next Tgl
                        '==================================
                        '--HIDER FIELD/COLUMN/ROW Normal --
                        '==================================
                        oSheet.Range(oSheet.Cells(i, 1), oSheet.Cells(i + 2, 1)).Merge
                        oSheet.Cells(i, 1) = "EMPLOYE ID"
                        oSheet.Range(oSheet.Cells(i, 2), oSheet.Cells(i + 2, 2)).Merge
                        oSheet.Cells(i, 2) = "EMPLOYE NAME"
                        oSheet.Range(oSheet.Cells(i, 3), oSheet.Cells(i + 2, 3)).Merge
                        oSheet.Cells(i, 3) = "GOLONGAN"
                        
                        oSheet.Range(oSheet.Cells(i, TglInt), oSheet.Cells(i + 2, TglInt)).Merge ' 27
                        oSheet.Cells(i, TglInt) = "TTL DPN" '6
                        oSheet.Range(oSheet.Cells(i, TglInt + 1), oSheet.Cells(i + 2, TglInt + 1)).Merge
                        oSheet.Cells(i, TglInt + 1) = "TTL PGI" '6
                        oSheet.Range(oSheet.Cells(i, TglInt + 2), oSheet.Cells(i + 2, TglInt + 2)).Merge
                        oSheet.Cells(i, TglInt + 2) = "TTL BLK"  '6
                        oSheet.Range(oSheet.Cells(i, TglInt + 3), oSheet.Cells(i + 2, TglInt + 3)).Merge
                        oSheet.Cells(i, TglInt + 3) = "UPAH HARI"  '6
                        oSheet.Range(oSheet.Cells(i, TglInt + 4), oSheet.Cells(i + 2, TglInt + 4)).Merge
                        oSheet.Cells(i, TglInt + 4) = "TTL UPAH"  '6
                        For h = 1 To 3
                            oSheet.Cells(i, h).ColumnWidth = 6
                            oSheet.Cells(i, h).VerticalAlignment = 2
                            oSheet.Cells(i, h).HorizontalAlignment = 3
                        Next h

                        For h = TglInt To TglInt + 4
                            oSheet.Cells(i, h).ColumnWidth = 6
                            If h = (TglInt + 3) Then
                                oSheet.Cells(i, h).ColumnWidth = 10
                            ElseIf h = (TglInt + 4) Then
                                oSheet.Cells(i, h).ColumnWidth = 10
                            End If
                            oSheet.Cells(i, h).VerticalAlignment = 2
                            oSheet.Cells(i, h).HorizontalAlignment = 3
                        Next h
                        i = i + 3
                    End If
                    
                        '===================
                        '--- value field ===
                        '===================
                        oSheet.Cells(i, 1) = .Fields("KAR_ID").Value '12.14
                        oSheet.Cells(i, 1).ColumnWidth = 12.14
                        oSheet.Cells(i, 2) = .Fields("KAR_NM").Value '17.14 DEP_NM
                        oSheet.Cells(i, 2).ColumnWidth = 17.14
                        oSheet.Cells(i, 3) = .Fields("GRP_NM").Value '17.14
                        oSheet.Cells(i, 3).ColumnWidth = 17.14
                        Dim TglRun As Date
                        TglInt = 4  'COLOMN LOOP
                        For TglRun = .Fields("sTgl").Value To .Fields("eTgl").Value
                            oSheet.Cells(i, TglInt) = .Fields("DPN_" & Day(TglRun) & "_" & Month(TglRun)).Value
                            oSheet.Cells(i, TglInt).HorizontalAlignment = 2
                            oSheet.Cells(i, TglInt + 1) = .Fields("PGI_" & Day(TglRun) & "_" & Month(TglRun)).Value '12.14
                            oSheet.Cells(i, TglInt + 1).HorizontalAlignment = 2
                            oSheet.Cells(i, TglInt + 2) = .Fields("BLK_" & Day(TglRun) & "_" & Month(TglRun)).Value '12.14
                            oSheet.Cells(i, TglInt + 2).HorizontalAlignment = 2
                            TglInt = TglInt + 3
                        Next TglRun
                        oSheet.Cells(i, TglInt) = .Fields("DPN_KAR").Value
                        oSheet.Cells(i, TglInt).HorizontalAlignment = 2
                        oSheet.Cells(i, TglInt + 1) = .Fields("PAGI_KAR").Value
                        oSheet.Cells(i, TglInt + 1).HorizontalAlignment = 2
                        oSheet.Cells(i, TglInt + 2) = .Fields("BLK_KAR").Value
                        oSheet.Cells(i, TglInt + 2).HorizontalAlignment = 2
                        oSheet.Cells(i, TglInt + 3) = .Fields("PAY_DAY").Value
                        oSheet.Cells(i, TglInt + 3).NumberFormat = "#,##0"
                        oSheet.Cells(i, TglInt + 4) = .Fields("TTL_UPAH_KAR").Value
                        oSheet.Cells(i, TglInt + 4).NumberFormat = "#,##0"
                        
                        
                        '==================
                        '--- FOOTER DEP ---
                        '==================
                        oSheet.Cells(i + 1, TglInt) = .Fields("DPN_DEP").Value
                        oSheet.Cells(i + 1, TglInt).HorizontalAlignment = 2
                        oSheet.Cells(i + 1, TglInt + 1) = .Fields("PAGI_DEP").Value
                        oSheet.Cells(i + 1, TglInt + 1).HorizontalAlignment = 2
                        oSheet.Cells(i + 1, TglInt + 2) = .Fields("BLK_DEP").Value
                        oSheet.Cells(i + 1, TglInt + 2).HorizontalAlignment = 2
                        oSheet.Cells(i + 1, TglInt + 4) = .Fields("TTL_UPAH_DEP").Value
                        oSheet.Cells(i + 1, TglInt + 4).NumberFormat = "#,##0"
                        oSheet.Cells(i + 1, 1) = "TOTAL UPAH DEPARTMEN : " & .Fields("DEP_NM").Value
                        oSheet.Cells(i + 1, 1).HorizontalAlignment = 1
                     
                    '=====================
                    '--- PROGRESS BAR ----
                    '=====================
                    If frm_actv = 1 Then '=== PRESENSI ===
                        FRM_ABSENSI_REPORT.ProgressBar1.Value = i
                    ElseIf frm_actv = 2 Then '=== PAYROLL ===
                        FRM_PAYROLL_REPORT.ProgressBar1.Value = i
                    ElseIf frm_actv = 3 Then '=== ACCOUNTING ===
                        FRM_ABSENSI_REPORT.ProgressBar1.Value = i
                    End If
                    
                    
                      i = i + 1
                    .MoveNext
                    Loop
                    
                End If
            End With
            Set rsRpt1 = Nothing
End If
If True Then '===== SHEET 1 ==============
On Error Resume Next
        Set oSheet = oWkbk.Sheets("sheet1")
            oSheet.Name = "WEEKLY DEPARTMEN"
            oSheet.Cells.Font.Name = "Utsaah"
            oSheet.Cells.Justify
            oSheet.Cells.Font.Color = &H8000000F
            oSheet.Columns.AutoFit
            oSheet.Cells(1, 1) = "PT. WANINDO PRIMA"
            oSheet.Cells(2, 1) = "PAYROLL WEEKLY SESSION CLOSE"
             oSheet.Cells.Font.Size = 9
            'Set oSheet = oWkbk.Sheets(1)
            For clmInt = 1 To 6 'TglInt + 2=tgl loop,
                oSheet.Cells(5, clmInt).Interior.ColorIndex = 15
                oSheet.Cells(5, clmInt).Borders.LineStyle = 1
                oSheet.Cells(5, clmInt).Borders.Color = &H8000000F
                oSheet.Cells(5, clmInt).Font.Color = &H8000000F
                oSheet.Cells(5, clmInt).BorderAround , , &HC0C0C0
                oSheet.Cells(5, clmInt).Font.Bold = True
                oSheet.Cells(5, clmInt).Font.Size = 9
                oSheet.Cells(5, clmInt).HorizontalAlignment = 3
            Next clmInt
            oSheet.Cells(5, 1) = "DEPARTEMENT"
            oSheet.Cells(5, 2) = "CABANG"
            oSheet.Cells(5, 3) = "TOTAL OT DEPAN"
            oSheet.Cells(5, 4) = "TOTAL SIANG"
            oSheet.Cells(5, 5) = "TOTAL OT BELAKANG"
            oSheet.Cells(5, 6) = "TOTAL UPAH PER DEPARTEMEN"
                
             i = 5
             Dim Qry_Dep As String
             Dim SttDep As Integer
             Qry_Dep = "Vsr_Report('Payroll_WeeklyExcel=" & STgl & "=" & ETgl & "=" & DepId & "=" & KarId & "=" & GrpId & "=" & CabId & "')"
             OpnRecRpt2 Qry_Dep, True
              With rsRpt2
                 If .RecordCount <> 0 Then
                     .MoveFirst
                     Do Until .EOF
                        If SttDep <> .Fields("DEP_ID").Value Then
                            SttDep = .Fields("DEP_ID").Value
                            oSheet.Cells(3, 1) = "PERIODE : " & .Fields("sTgl").Value & "  TO  " & .Fields("eTgl").Value
                            For clmInt = 1 To 6 'TglInt + 2=tgl loop,
                                    'oSheet.Cells(i + 1, clmInt).Interior.ColorIndex = 15
                                    oSheet.Cells(i + 1, clmInt).Borders.LineStyle = 1
                                    oSheet.Cells(i + 1, clmInt).Borders.Color = &H8000000F
                                    oSheet.Cells(i + 1, clmInt).Font.Color = &H8000000F
                                    oSheet.Cells(i + 1, clmInt).BorderAround , , &HC0C0C0
                                    oSheet.Cells(i + 1, clmInt).Font.Bold = True
                                    oSheet.Cells(i + 1, clmInt).Font.Size = 9
                                    oSheet.Cells(i + 1, clmInt).HorizontalAlignment = 3
                            Next clmInt
                            oSheet.Cells(i + 1, 1) = .Fields("DEP_NM").Value
                            oSheet.Cells(i + 1, 1).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 2) = .Fields("CAB_NM").Value
                            oSheet.Cells(i + 1, 2).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 3) = .Fields("DPN_DEP").Value
                            oSheet.Cells(i + 1, 3).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 4) = .Fields("PAGI_DEP").Value
                            oSheet.Cells(i + 1, 4).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 5) = .Fields("BLK_DEP").Value
                            oSheet.Cells(i + 1, 5).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 6) = .Fields("TTL_UPAH_DEP").Value
                            oSheet.Cells(i + 1, 6).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 6).NumberFormat = "#,##0"
                            oSheet.Columns.AutoFit
                            i = i + 1
                         End If
                       .MoveNext
                     Loop
                 End If
             End With
            
             Set rsRpt2 = Nothing
End If
If True Then '===== SHEET 2 ==============
On Error Resume Next
        Set oSheet = oWkbk.Sheets("sheet2")
            oSheet.Name = "WEEKLY CABANG"
            oSheet.Cells.Font.Name = "Utsaah"
            oSheet.Cells.Justify
            oSheet.Cells.Font.Color = &H8000000F
            oSheet.Columns.AutoFit
            oSheet.Cells(1, 1) = "PT. WANINDO PRIMA"
            oSheet.Cells(2, 1) = "PAYROLL WEEKLY SESSION CLOSE"
            oSheet.Cells.Font.Size = 9
            'Set oSheet = oWkbk.Sheets(2)
            For clmInt = 1 To 5 'TglInt + 2=tgl loop,
                oSheet.Cells(5, clmInt).Interior.ColorIndex = 15
                oSheet.Cells(5, clmInt).Borders.LineStyle = 1
                oSheet.Cells(5, clmInt).Borders.Color = &H8000000F
                oSheet.Cells(5, clmInt).Font.Color = &H8000000F
                oSheet.Cells(5, clmInt).BorderAround , , &HC0C0C0
                oSheet.Cells(5, clmInt).Font.Bold = True
                oSheet.Cells(5, clmInt).Font.Size = 9
                oSheet.Cells(5, clmInt).HorizontalAlignment = 3
            Next clmInt
            oSheet.Cells(5, 1) = "CABANG"
            oSheet.Cells(5, 2) = "TOTAL OT DEPAN"
            oSheet.Cells(5, 3) = "TOTAL SIANG"
            oSheet.Cells(5, 4) = "TOTAL OT BELAKANG"
            oSheet.Cells(5, 5) = "TOTAL UPAH PER CABANG"
                
             i = 5
             Dim Qry_Cab As String
             Dim SttCab As String
             Qry_Cab = "Vsr_Report('Payroll_WeeklyExcel=" & STgl & "=" & ETgl & "=" & DepId & "=" & KarId & "=" & GrpId & "=" & CabId & "')"
             OpnRecRpt2 Qry_Cab, True
              With rsRpt2
                 If .RecordCount <> 0 Then
                     .MoveFirst
                     Do Until .EOF
                        If SttCab <> .Fields("CAB_ID").Value Then
                            SttCab = .Fields("CAB_ID").Value
                            oSheet.Cells(3, 1) = "PERIODE : " & .Fields("sTgl").Value & "  TO  " & .Fields("eTgl").Value
           
                            For clmInt = 1 To 5 'TglInt + 2=tgl loop,
                                    'oSheet.Cells(i + 1, clmInt).Interior.ColorIndex = 15
                                    oSheet.Cells(i + 1, clmInt).Borders.LineStyle = 1
                                    oSheet.Cells(i + 1, clmInt).Borders.Color = &H8000000F
                                    oSheet.Cells(i + 1, clmInt).Font.Color = &H8000000F
                                    oSheet.Cells(i + 1, clmInt).BorderAround , , &HC0C0C0
                                    oSheet.Cells(i + 1, clmInt).Font.Bold = True
                                    oSheet.Cells(i + 1, clmInt).Font.Size = 9
                                    oSheet.Cells(i + 1, clmInt).HorizontalAlignment = 3
                            Next clmInt
                            oSheet.Cells(i + 1, 1) = .Fields("CAB_NM").Value
                            oSheet.Cells(i + 1, 1).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 2) = .Fields("DPN_CAB").Value
                            oSheet.Cells(i + 1, 2).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 3) = .Fields("PAGI_CAB").Value
                            oSheet.Cells(i + 1, 3).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 4) = .Fields("BLK_CAB").Value
                            oSheet.Cells(i + 1, 4).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 5) = .Fields("TTL_UPAH_CAB").Value
                            oSheet.Cells(i + 1, 5).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 5).NumberFormat = "#,##0"
                            oSheet.Columns.AutoFit
                            i = i + 1
                         End If
                       .MoveNext
                     Loop
                 End If
             End With
             Set rsRpt2 = Nothing
End If

If True Then '===== SHEET 3 ==============
On Error Resume Next
        Set oSheet = oWkbk.Sheets("sheet3")
            oSheet.Name = "SLIP GAJI"
            oSheet.Cells.Font.Name = "Utsaah"
            oSheet.Cells.Justify
            oSheet.Cells.Font.Color = &H8000000F
            oSheet.Columns.AutoFit
           ' oSheet.Range(oSheet.Cells(1, 1), oSheet.Cells(3, 2)).Merge
            
    Dim Qry_Slip As String
    Dim Lurus, bawah As Double
    Qry_Slip = "Vsr_Report('Payroll_WeeklyExcel=" & STgl & "=" & ETgl & "=" & DepId & "=" & KarId & "=" & GrpId & "=" & CabId & "')"
    'Qry_Slip = "Vsr_Report('Payroll_WeeklyExcel=2015-03-01=2015-03-06=0=WAN.HO.000087=0=0')"
    OpnRecRpt2 Qry_Slip, True
    Lurus = 7
    bawah = 17
    i = 1
    trecs = 0
     With rsRpt2
        If .RecordCount <> 0 Then
            
                    .MoveFirst
                    Do Until .EOF
                        trecs = trecs + 1
                    .MoveNext
                    
                    Loop
                    
                    '=====================
                    '--- PROGRESS BAR ----
                    '=====================
                    If frm_actv = 1 Then '=== PRESENSI ===
                        With FRM_ABSENSI_REPORT
                            .ProgressBar1.Visible = True
                            .ProgressBar1.Min = 0
                            .ProgressBar1.Max = trecs
                        End With
                    ElseIf frm_actv = 2 Then '=== PAYROLL ===
                        With FRM_PAYROLL_REPORT
                            .ProgressBar1.Visible = True
                            .ProgressBar1.Min = 0
                            .ProgressBar1.Max = trecs
                        End With
                    ElseIf frm_actv = 3 Then '=== ACCOUNTING ===
                        With FRM_ABSENSI_REPORT
                            .ProgressBar1.Visible = True
                            .ProgressBar1.Min = 0
                            .ProgressBar1.Max = trecs
                        End With
                    End If
            
            .MoveFirst
            Do Until .EOF
              If Lurus = 35 Then
                Lurus = 7
                bawah = bawah + 17
              End If
               ' oSheet.Range(oSheet.Cells(bawah - 16, Lurus - 6), oSheet.Cells(bawah, Lurus)).BorderAround LineStyle:=Excel.XlLineStyle.xlContinuous, Weight:=Excel.XlBorderWeight.xlThick
                With oSheet.Range(oSheet.Cells(bawah - 16, Lurus - 6), oSheet.Cells(bawah, Lurus))
                    .Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Borders(xlEdgeRight).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                End With

                With oSheet.Pictures.Insert(App.Path & "\Img\Logo\wanindo.jpg")
                     With .ShapeRange
                         .LockAspectRatio = True
                         .Width = 78
                         .Height = 50
                     End With
                        '.left = 2
                        '.top = 3
                      .left = (oSheet.Cells(bawah - 16, Lurus - 6).left) + 2
                      .top = (oSheet.Cells(bawah - 16, Lurus - 6).top) + 2
                     .Placement = 2
                     .PrintObject = True
                 End With

                 oSheet.Cells(bawah - 15, Lurus - 1) = "PT. WANINDO PRIMA"
                 oSheet.Cells(bawah - 15, Lurus - 1).HorizontalAlignment = -4152
                 oSheet.Range(oSheet.Cells(bawah - 14, Lurus - 4), oSheet.Cells(bawah - 14, Lurus - 2)).Merge
                 oSheet.Cells(bawah - 14, Lurus - 4) = "PAYROLL WEEKLY" ' & WeekdayName(Weekday(.Fields("eTgl").Value)) & " " & Year(.Fields("eTgl").Value)
                 oSheet.Cells(bawah - 14, Lurus - 4).HorizontalAlignment = -4108
                 oSheet.Cells.Font.Size = 9
                 
                 oSheet.Cells(bawah - 12, Lurus - 5) = "EmployeID     "
                 oSheet.Cells(bawah - 11, Lurus - 5) = "Employe Name  "
                 oSheet.Cells(bawah - 10, Lurus - 5) = "Position      "
                 oSheet.Cells(bawah - 12, Lurus - 3) = " : " & .Fields("KAR_ID").Value
                 oSheet.Cells(bawah - 11, Lurus - 3) = " : " & .Fields("KAR_NM").Value
                 oSheet.Cells(bawah - 10, Lurus - 3) = " : " & .Fields("DEP_NM").Value & " / " & .Fields("GRP_NM").Value
                 'PENDAPATAN
                 oSheet.Cells(bawah - 8, Lurus - 5) = "PENDAPATAN :"
                 oSheet.Cells(bawah - 7, Lurus - 5) = "OT Depan :"
                 oSheet.Cells(bawah - 6, Lurus - 5) = "Pagi :"
                 oSheet.Cells(bawah - 5, Lurus - 5) = "OT Belakang :"
                 oSheet.Cells(bawah - 4, Lurus - 5) = "Upah Hari :"
                 oSheet.Cells(bawah - 3, Lurus - 5) = "Total Pendapatan"
                 
                 '-Value Field ------
                 oSheet.Cells(bawah - 7, Lurus - 4) = Round(.Fields("DPN_KAR").Value, 2)
                 oSheet.Cells(bawah - 6, Lurus - 4) = Round(.Fields("PAGI_KAR").Value, 2)
                 oSheet.Cells(bawah - 5, Lurus - 4) = Round(.Fields("BLK_KAR").Value, 2)
                 oSheet.Cells(bawah - 4, Lurus - 4) = .Fields("PAY_DAY").Value
                 oSheet.Cells(bawah - 4, Lurus - 4).NumberFormat = "#,##0"
                 oSheet.Cells(bawah - 3, Lurus - 4) = .Fields("TTL_UPAH_KAR").Value
                 oSheet.Cells(bawah - 3, Lurus - 4).NumberFormat = "#,##0"
                 For HdrRow = (bawah - 7) To (bawah - 3)
                     oSheet.Cells(HdrRow, Lurus - 5).HorizontalAlignment = -4152
                     oSheet.Cells(HdrRow, Lurus - 6).ColumnWidth = 2
                     oSheet.Cells(HdrRow, Lurus - 5).ColumnWidth = 7.29
                     oSheet.Cells(HdrRow, Lurus - 4).ColumnWidth = 7.29
                     'oSheet.Cells(HdrRow, Lurus - 4).NumberFormat = "#,##0"
                 Next HdrRow
                 
                 'POTONGAN
                 oSheet.Cells(bawah - 8, Lurus - 2) = "POTONGAN"
                 oSheet.Cells(bawah - 7, Lurus - 2) = "PPH21 :"
                 oSheet.Cells(bawah - 6, Lurus - 2) = "Jamsos :"
                 oSheet.Cells(bawah - 5, Lurus - 2) = "Asuransi :"
                 oSheet.Cells(bawah - 4, Lurus - 2) = "Pinjaman :"
                 oSheet.Cells(bawah - 3, Lurus - 2) = "Total Potongan"
                 '---- value field -- ,A.,A.,A.POT_PINJAMAN
                 oSheet.Cells(bawah - 7, Lurus - 1) = 0
                 oSheet.Cells(bawah - 6, Lurus - 1) = 0
                 oSheet.Cells(bawah - 5, Lurus - 1) = 0
                 oSheet.Cells(bawah - 4, Lurus - 1) = 0
                 oSheet.Cells(bawah - 3, Lurus - 1) = 0
    
                 For HdrRow = (bawah - 7) To (bawah - 3)
                     oSheet.Cells(HdrRow, Lurus - 2).HorizontalAlignment = -4152
                     oSheet.Cells(HdrRow, Lurus - 3).ColumnWidth = 2
                     oSheet.Cells(HdrRow, Lurus - 2).ColumnWidth = 7.29
                     oSheet.Cells(HdrRow, Lurus - 1).ColumnWidth = 7.29
                     oSheet.Cells(HdrRow, Lurus).ColumnWidth = 2
                     oSheet.Cells(HdrRow, Lurus - 1).NumberFormat = "#,##0"
                 Next HdrRow
                 oSheet.Cells(bawah - 1, Lurus - 5) = "TAKE WEKKLY HOME PAY"
                 oSheet.Cells(bawah - 1, Lurus - 2) = .Fields("TTL_UPAH_KAR").Value
                 oSheet.Cells(bawah - 1, Lurus - 2).HorizontalAlignment = -4152
                 oSheet.Cells(bawah - 1, Lurus - 2).NumberFormat = "#,##0"
                 'Lurus
                ' oSheet.Range(oSheet.Cells(bawah - 6, Lurus - 5), oSheet.Cells(bawah - 1, Lurus - 1)).BorderAround LineStyle:=Excel.XlLineStyle.xlContinuous, Weight:=Excel.XlBorderWeight.xlThick
                 Lurus = Lurus + 7
                    '=====================
                    '--- PROGRESS BAR ----
                    '=====================
                    If frm_actv = 1 Then '=== PRESENSI ===
                        FRM_ABSENSI_REPORT.ProgressBar1.Value = i
                    ElseIf frm_actv = 2 Then '=== PAYROLL ===
                        FRM_PAYROLL_REPORT.ProgressBar1.Value = i
                    ElseIf frm_actv = 3 Then '=== ACCOUNTING ===
                        FRM_ABSENSI_REPORT.ProgressBar1.Value = i
                    End If
                    i = i + 1
                 .MoveNext
                 Loop
        End If
    End With
End If
End Sub


'=============================
'------ PAYROLL monthly ------
'=============================
Public Sub PAYROLL_MONTHLY(STgl As String, ETgl As String, DepId As Integer, KarId As String, GrpId As Integer, CabId As String, MenuRpt As String, frm_actv As Integer)
Dim ApExcel  As Excel.Application, oWkbk As Excel.Workbook, oSheet As Excel.Worksheet
Dim i, J, a, x, h As Long
Dim Qry1 As String
Dim Qry2 As String
Dim STotalStr As String
Set ApExcel = CreateObject("Excel.Application"): ApExcel.Visible = True
Set oWkbk = ApExcel.Workbooks.Add

If True Then '===== SHEET Pertma ==============
On Error Resume Next
        Set oSheet = oWkbk.Sheets.Add
            oSheet.Name = "PAYROLL_MONTHLY"
            oSheet.Cells.Font.Name = "Utsaah"
            oSheet.Cells.Justify
            oSheet.Cells.Font.Color = &H8000000F
            'oSheet.Columns.AutoFit
            oSheet.Cells(1, 1) = "PT. WANINDO PRIMA"
            oSheet.Cells(2, 1) = "PAYROLL MONTHLY SESSION CLOSE"
             oSheet.Cells.Font.Size = 9
            Dim Qry_OTDaily As String
            Dim Qry_OTDaily_Sub As String
            Dim Stt_DepID As Integer
            trecs = 0
            Qry_OTDaily = "Vsr_Report('Payroll_MonthlyExcel=" & STgl & "=" & ETgl & "=" & DepId & "=" & KarId & "=" & GrpId & "=" & CabId & "')"
            'Payroll_WeeklyExcel=2015-03-01=2015-03-06=0=0=0=0'
            'Menu,sDateRun,eDateRun,DeptID,KarId,GrpId,CabID
            'STgl,ETgl,DepId, KarId, GrpId,CabId,MenuRpt
            OpnRecRpt1 Qry_OTDaily, True
            Stt_DepID = 0
            i = 4
            With rsRpt1
               ' i = 3
                If .RecordCount <> 0 Then
                    .MoveFirst
                    Do Until .EOF
                        trecs = trecs + 1
                    .MoveNext
                    
                    Loop
                    
                    '=====================
                    '--- PROGRESS BAR ----
                    '=====================
                    If frm_actv = 1 Then '=== PRESENSI ===
                        With FRM_ABSENSI_REPORT
                            .ProgressBar1.Visible = True
                            .ProgressBar1.Min = 0
                            .ProgressBar1.Max = trecs
                        End With
                    ElseIf frm_actv = 2 Then '=== PAYROLL ===
                        With FRM_PAYROLL_REPORT
                            .ProgressBar1.Visible = True
                            .ProgressBar1.Min = 0
                            .ProgressBar1.Max = trecs
                        End With
                    ElseIf frm_actv = 3 Then '=== ACCOUNTING ===
                        With FRM_ABSENSI_REPORT
                            .ProgressBar1.Visible = True
                            .ProgressBar1.Min = 0
                            .ProgressBar1.Max = trecs
                        End With
                    End If

                    .MoveFirst
                    Do Until .EOF
                    '======================
                    '--- Group DEP --------
                    '======================
                    If Stt_DepID <> (.Fields("DEP_ID").Value) Then
                        Stt_DepID = .Fields("DEP_ID").Value
                        oSheet.Cells(3, 1) = "PERIODE : " & .Fields("sTgl").Value & "  TO  " & .Fields("eTgl").Value
                        Dim clmInt, HdrRow As Integer
                        clmInt = 1  'COLOMN HEADER DATE LOOP
                        i = i + 1   'ROW
                        '==========
                        '--BORDER--
                        '==========
                        For clmInt = 1 To 22
                            For HdrRow = i To i + 1
                                oSheet.Cells(HdrRow, clmInt).Interior.ColorIndex = 15
                                oSheet.Cells(HdrRow, clmInt).Borders.LineStyle = 1
                                oSheet.Cells(HdrRow, clmInt).Borders.Color = &H8000000F
                                oSheet.Cells(HdrRow, clmInt).Font.Color = &H8000000F
                                oSheet.Cells(HdrRow, clmInt).BorderAround , , &HC0C0C0
                                oSheet.Cells(HdrRow, clmInt).Font.Bold = True
                                oSheet.Cells(HdrRow, clmInt).Font.Size = 9
                                oSheet.Cells(HdrRow, clmInt).HorizontalAlignment = 3
                            Next HdrRow
                                'Left:    -4131  =2
                                'Center:  -4108  =3
                                'Right:   -4152 =1
                        Next clmInt
                         oSheet.Range(oSheet.Cells(i, 1), oSheet.Cells(i + 1, 1)).Merge
                         oSheet.Range(oSheet.Cells(i, 2), oSheet.Cells(i + 1, 2)).Merge
                         oSheet.Range(oSheet.Cells(i, 3), oSheet.Cells(i + 1, 3)).Merge
                         oSheet.Range(oSheet.Cells(i, 4), oSheet.Cells(i, 8)).Merge
                         oSheet.Range(oSheet.Cells(i, 9), oSheet.Cells(i, 14)).Merge
                         oSheet.Range(oSheet.Cells(i, 15), oSheet.Cells(i, 21)).Merge
                         oSheet.Range(oSheet.Cells(i, 22), oSheet.Cells(i + 1, 22)).Merge
                       
                        i = i + 1
                        '==================================
                        '--HIDER FIELD/COLUMN/ROW Normal --
                        '==================================
                        oSheet.Cells(i - 1, 1) = "EMPLOYE ID"
                        oSheet.Cells(i - 1, 2) = "EMPLOYE NAME"
                        oSheet.Cells(i - 1, 3) = "GOLONGAN"
                        oSheet.Cells(i, 4) = "TTL DPN"
                        oSheet.Cells(i, 5) = "TTL PGI"
                        oSheet.Cells(i, 6) = "TTL BLK"
                        oSheet.Cells(i, 7) = "UPAH HARI"
                        oSheet.Cells(i, 8) = "TTL UPAH"
                        oSheet.Cells(i, 9) = "TELAT" 'hanya untuk staff
                        oSheet.Cells(i, 10) = "PPH21"
                        oSheet.Cells(i, 11) = "JAMSOS"
                        oSheet.Cells(i, 12) = "ASURANSI"
                        oSheet.Cells(i, 13) = "PINJAMAN"
                        oSheet.Cells(i, 14) = "TOTAL_POTONGAN"
                        oSheet.Cells(i, 15) = "GAJI POKOK"
                        oSheet.Cells(i, 16) = "TUNJANGAN"
                        oSheet.Cells(i, 17) = "TRANSPORT"
                        oSheet.Cells(i, 18) = "MAKAN"
                        oSheet.Cells(i, 19) = "BONUS"
                        oSheet.Cells(i, 20) = "ENTERTAIN"
                        oSheet.Cells(i, 21) = "TOTAL GAJI"
                        oSheet.Cells(i - 1, 22) = "GAJI DITERIMA"
                        
                      
                        For clmInt = 1 To 3
                            oSheet.Cells(i - 1, clmInt).ColumnWidth = 15
                            oSheet.Cells(i - 1, clmInt).VerticalAlignment = 2
                            oSheet.Cells(i - 1, clmInt).HorizontalAlignment = 3
                        Next clmInt
                            oSheet.Cells(i, 14).ColumnWidth = 12
                            oSheet.Cells(i - 1, 22).ColumnWidth = 12
                            oSheet.Cells(i - 1, 22).VerticalAlignment = 2
                            oSheet.Cells(i - 1, 22).HorizontalAlignment = 3
                     

                        i = i + 1
                    End If
                    
                        '===================
                        '--- value field ===
                        '===================
                        oSheet.Cells(i, 1) = .Fields("KAR_ID").Value '12.14
                        oSheet.Cells(i, 1).ColumnWidth = 12.14
                        oSheet.Cells(i, 2) = .Fields("KAR_NM").Value '17.14 DEP_NM
                        oSheet.Cells(i, 2).ColumnWidth = 17.14
                        oSheet.Cells(i, 3) = .Fields("GRP_NM").Value '17.14
                        oSheet.Cells(i, 3).ColumnWidth = 17.14
                        oSheet.Cells(i, 4) = .Fields("DPN_KAR").Value
                        oSheet.Cells(i, 4).HorizontalAlignment = 2
                        oSheet.Cells(i, 5) = .Fields("PGI_KAR").Value
                        oSheet.Cells(i, 5).HorizontalAlignment = 2
                        oSheet.Cells(i, 6) = .Fields("BLK_KAR").Value
                        oSheet.Cells(i, 6).HorizontalAlignment = 2
                        oSheet.Cells(i, 7) = .Fields("PAY_DAY").Value
                        oSheet.Cells(i, 7).NumberFormat = "#,##0"
                        oSheet.Cells(i, 8) = .Fields("TTL_UPAH_BLN").Value
                        oSheet.Cells(i, 8).NumberFormat = "#,##0"
                        oSheet.Cells(i, 9) = "0" ' hanya untu potongan telat harian
                        oSheet.Cells(i, 9).NumberFormat = "#,##0"
                        oSheet.Cells(i, 10) = .Fields("POT_PPH21").Value
                        oSheet.Cells(i, 10).NumberFormat = "#,##0"
                        oSheet.Cells(i, 11) = .Fields("POT_JAMSOS_JHT_TK").Value
                        oSheet.Cells(i, 11).NumberFormat = "#,##0"
                        oSheet.Cells(i, 12) = .Fields("POT_ASURANSI").Value
                        oSheet.Cells(i, 12).NumberFormat = "#,##0"
                        oSheet.Cells(i, 13) = .Fields("POT_PINJAMAN").Value
                        oSheet.Cells(i, 13).NumberFormat = "#,##0"
                        oSheet.Cells(i, 14) = .Fields("TOTAL_POT").Value
                        oSheet.Cells(i, 14).NumberFormat = "#,##0"
                        oSheet.Cells(i, 15) = .Fields("PAY_SALARY").Value
                        oSheet.Cells(i, 15).NumberFormat = "#,##0"
                        oSheet.Cells(i, 16) = .Fields("PAY_TUNJANGAN").Value
                        oSheet.Cells(i, 16).NumberFormat = "#,##0"
                        oSheet.Cells(i, 17) = .Fields("PAY_TRANSPORT").Value
                        oSheet.Cells(i, 17).NumberFormat = "#,##0"
                        oSheet.Cells(i, 18) = .Fields("PAY_EAT").Value
                        oSheet.Cells(i, 18).NumberFormat = "#,##0"
                        oSheet.Cells(i, 19) = .Fields("PAY_BONUS").Value
                        oSheet.Cells(i, 19).NumberFormat = "#,##0"
                        oSheet.Cells(i, 20) = .Fields("PAY_ENTERTAIN").Value
                        oSheet.Cells(i, 20).NumberFormat = "#,##0"
                        oSheet.Cells(i, 21) = .Fields("TOTAL_GAJI").Value
                        oSheet.Cells(i, 21).NumberFormat = "#,##0"
                        oSheet.Cells(i, 22) = .Fields("GAJI_DITERIMA").Value
                        oSheet.Cells(i, 23).NumberFormat = "#,##0"
                        
                        '==================
                        '--- FOOTER DEP ---
                        '==================
                        oSheet.Cells(i + 1, 4) = .Fields("DPN_DEP").Value
                        oSheet.Cells(i + 1, 4).HorizontalAlignment = 2
                        oSheet.Cells(i + 1, 5) = .Fields("PAGI_DEP").Value
                        oSheet.Cells(i + 1, 5).HorizontalAlignment = 2
                        oSheet.Cells(i + 1, 6) = .Fields("BLK_DEP").Value
                        oSheet.Cells(i + 1, 6).HorizontalAlignment = 2
                        oSheet.Cells(i + 1, 8) = .Fields("UPAHBLN_DEP").Value
                        oSheet.Cells(i + 1, 8).NumberFormat = "#,##0"
                        oSheet.Cells(i + 1, 9) = "0"  ' potongan harian hanya untuk staff
                        oSheet.Cells(i + 1, 9).NumberFormat = "#,##0"
                        oSheet.Cells(i + 1, 10) = .Fields("PPH21_DEP").Value
                        oSheet.Cells(i + 1, 10).NumberFormat = "#,##0"
                        oSheet.Cells(i + 1, 11) = .Fields("JSOS_DEP").Value
                        oSheet.Cells(i + 1, 11).NumberFormat = "#,##0"
                        oSheet.Cells(i + 1, 12) = .Fields("ASR_DEP").Value
                        oSheet.Cells(i + 1, 12).NumberFormat = "#,##0"
                        oSheet.Cells(i + 1, 13) = .Fields("PNJ_DEP").Value
                        oSheet.Cells(i + 1, 13).NumberFormat = "#,##0"
                        oSheet.Cells(i + 1, 14) = .Fields("TTLPOT_DEP").Value
                        oSheet.Cells(i + 1, 14).NumberFormat = "#,##0"
                        oSheet.Cells(i + 1, 15) = .Fields("SAL_DEP").Value
                        oSheet.Cells(i + 1, 15).NumberFormat = "#,##0"
                        oSheet.Cells(i + 1, 16) = .Fields("TNJ_DEP").Value
                        oSheet.Cells(i + 1, 16).NumberFormat = "#,##0"
                        oSheet.Cells(i + 1, 17) = .Fields("TRN_DEP").Value
                        oSheet.Cells(i + 1, 17).NumberFormat = "#,##0"
                        oSheet.Cells(i + 1, 18) = .Fields("MKN_DEP").Value
                        oSheet.Cells(i + 1, 18).NumberFormat = "#,##0"
                        oSheet.Cells(i + 1, 19) = .Fields("BNS_DEP").Value
                        oSheet.Cells(i + 1, 19).NumberFormat = "#,##0"
                        oSheet.Cells(i + 1, 20) = .Fields("ENT_DEP").Value
                        oSheet.Cells(i + 1, 20).NumberFormat = "#,##0"
                        oSheet.Cells(i + 1, 21) = .Fields("TTLGAJI_DEP").Value
                        oSheet.Cells(i + 1, 21).NumberFormat = "#,##0"
                        oSheet.Cells(i + 1, 22) = .Fields("GJITRMA_DEP").Value
                        oSheet.Cells(i + 1, 2).NumberFormat = "#,##0"
                       
                        oSheet.Cells(i + 1, 1) = "TOTAL UPAH DEPARTMEN : " & .Fields("DEP_NM").Value
                        oSheet.Cells(i + 1, 1).HorizontalAlignment = 1
                     
                    '=====================
                    '--- PROGRESS BAR ----
                    '=====================
                    If frm_actv = 1 Then '=== PRESENSI ===
                        FRM_ABSENSI_REPORT.ProgressBar1.Value = i
                    ElseIf frm_actv = 2 Then '=== PAYROLL ===
                        FRM_PAYROLL_REPORT.ProgressBar1.Value = i
                    ElseIf frm_actv = 3 Then '=== ACCOUNTING ===
                        FRM_ABSENSI_REPORT.ProgressBar1.Value = i
                    End If
                    
                    
                      i = i + 1
                    .MoveNext
                    Loop
                    
                End If
            End With
            Set rsRpt1 = Nothing
End If

If True Then '===== SHEET 1 ==============
On Error Resume Next
        Set oSheet = oWkbk.Sheets("sheet1")
            oSheet.Name = "MONTHY DEPARTMEN"
            oSheet.Cells.Font.Name = "Utsaah"
            oSheet.Cells.Justify
            oSheet.Cells.Font.Color = &H8000000F
            oSheet.Columns.AutoFit
            oSheet.Cells(1, 1) = "PT. WANINDO PRIMA"
            oSheet.Cells(2, 1) = "PAYROLL MONTHLY SESSION CLOSE"
            oSheet.Cells.Font.Size = 9
            HdrRow = 5
           For clmInt = 1 To 20
                For HdrRow = 5 To 6
                    oSheet.Cells(HdrRow, clmInt).Interior.ColorIndex = 15
                    oSheet.Cells(HdrRow, clmInt).Borders.LineStyle = 1
                    oSheet.Cells(HdrRow, clmInt).Borders.Color = &H8000000F
                    oSheet.Cells(HdrRow, clmInt).Font.Color = &H8000000F
                    oSheet.Cells(HdrRow, clmInt).BorderAround , , &HC0C0C0
                    oSheet.Cells(HdrRow, clmInt).Font.Bold = True
                    oSheet.Cells(HdrRow, clmInt).Font.Size = 9
                    oSheet.Cells(HdrRow, clmInt).HorizontalAlignment = 3
                Next HdrRow
            Next clmInt
            
            oSheet.Range(oSheet.Cells(5, 1), oSheet.Cells(6, 1)).Merge
            oSheet.Cells(5, 1).ColumnWidth = 6
            oSheet.Cells(5, 1).VerticalAlignment = 2
            oSheet.Cells(5, 1).HorizontalAlignment = 3
            oSheet.Range(oSheet.Cells(5, 2), oSheet.Cells(6, 2)).Merge
            oSheet.Cells(5, 2).ColumnWidth = 6
            oSheet.Cells(5, 2).VerticalAlignment = 2
            oSheet.Cells(5, 2).HorizontalAlignment = 3
            
            oSheet.Range(oSheet.Cells(5, 3), oSheet.Cells(5, 6)).Merge
            oSheet.Range(oSheet.Cells(5, 7), oSheet.Cells(5, 12)).Merge
            oSheet.Range(oSheet.Cells(5, 13), oSheet.Cells(5, 19)).Merge
            oSheet.Range(oSheet.Cells(5, 20), oSheet.Cells(6, 20)).Merge
            oSheet.Cells(5, 20).ColumnWidth = 6
            oSheet.Cells(5, 20).VerticalAlignment = 2
            oSheet.Cells(5, 20).HorizontalAlignment = 3
            
            HdrRow = 6
            oSheet.Cells(5, 1) = "DEPARTEMENT"
            oSheet.Cells(5, 2) = "CABANG"
            oSheet.Cells(5, 3) = "JUMLAH HARI DALAM SEBULAN"
            oSheet.Cells(HdrRow, 3) = "TOTAL OT DEPAN"
            oSheet.Cells(HdrRow, 4) = "TOTAL SIANG"
            oSheet.Cells(HdrRow, 5) = "TOTAL OT BELAKANG"
            oSheet.Cells(HdrRow, 6) = "TOTAL UPAH"
            oSheet.Cells(5, 7) = "JUMLAH KOMPONEN POTONGAN"
            oSheet.Cells(HdrRow, 7) = "TELAT" 'hanya untuk staff
            oSheet.Cells(HdrRow, 8) = "PPH21"
            oSheet.Cells(HdrRow, 9) = "JAMSOS"
            oSheet.Cells(HdrRow, 10) = "ASURANSI"
            oSheet.Cells(HdrRow, 11) = "PINJAMAN"
            oSheet.Cells(HdrRow, 12) = "TOTAL_POTONGAN"
            oSheet.Cells(5, 13) = "JUMLAH KOMPONEN GAJI"
            oSheet.Cells(HdrRow, 13) = "GAJI POKOK"
            oSheet.Cells(HdrRow, 14) = "TUNJANGAN"
            oSheet.Cells(HdrRow, 15) = "TRANSPORT"
            oSheet.Cells(HdrRow, 16) = "MAKAN"
            oSheet.Cells(HdrRow, 17) = "BONUS"
            oSheet.Cells(HdrRow, 18) = "ENTERTAIN"
            oSheet.Cells(HdrRow, 19) = "TOTAL GAJI"
            oSheet.Cells(5, 20) = "GAJI DITERIMA"
                
             i = HdrRow
             Dim Qry_Dep As String
             Dim SttDep As Integer
             Qry_Dep = "Vsr_Report('Payroll_MonthlyExcel=" & STgl & "=" & ETgl & "=" & DepId & "=" & KarId & "=" & GrpId & "=" & CabId & "')"
             OpnRecRpt2 Qry_Dep, True
              With rsRpt2
                 If .RecordCount <> 0 Then
                     .MoveFirst
                     Do Until .EOF
                        If SttDep <> .Fields("DEP_ID").Value Then
                            SttDep = .Fields("DEP_ID").Value
                            oSheet.Cells(3, 1) = "PERIODE : " & .Fields("sTgl").Value & "  TO  " & .Fields("eTgl").Value
                            For clmInt = 1 To 20 'TglInt + 2=tgl loop,
                                    'oSheet.Cells(i + 1, clmInt).Interior.ColorIndex = 15
                                    oSheet.Cells(i + 1, clmInt).Borders.LineStyle = 1
                                    oSheet.Cells(i + 1, clmInt).Borders.Color = &H8000000F
                                    oSheet.Cells(i + 1, clmInt).Font.Color = &H8000000F
                                    oSheet.Cells(i + 1, clmInt).BorderAround , , &HC0C0C0
                                    oSheet.Cells(i + 1, clmInt).Font.Bold = True
                                    oSheet.Cells(i + 1, clmInt).Font.Size = 9
                                    oSheet.Cells(i + 1, clmInt).HorizontalAlignment = 3
                            Next clmInt
                            oSheet.Cells(i + 1, 1) = .Fields("DEP_NM").Value
                            oSheet.Cells(i + 1, 1).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 2) = .Fields("CAB_NM").Value
                            oSheet.Cells(i + 1, 2).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 3) = .Fields("DPN_DEP").Value
                            oSheet.Cells(i + 1, 3).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 4) = .Fields("PAGI_DEP").Value
                            oSheet.Cells(i + 1, 4).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 5) = .Fields("BLK_DEP").Value
                            oSheet.Cells(i + 1, 5).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 6) = .Fields("UPAHBLN_DEP").Value
                            oSheet.Cells(i + 1, 6).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 6).NumberFormat = "#,##0"
                            oSheet.Cells(i + 1, 7) = "0" 'telat untuk staff
                            oSheet.Cells(i + 1, 7).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 7).NumberFormat = "#,##0"
                            oSheet.Cells(i + 1, 8) = .Fields("PPH21_DEP").Value
                            oSheet.Cells(i + 1, 8).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 8).NumberFormat = "#,##0"
                            oSheet.Cells(i + 1, 9) = .Fields("JSOS_DEP").Value
                            oSheet.Cells(i + 1, 9).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 9).NumberFormat = "#,##0"
                            oSheet.Cells(i + 1, 10) = .Fields("ASR_DEP").Value
                            oSheet.Cells(i + 1, 10).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 10).NumberFormat = "#,##0"
                            oSheet.Cells(i + 1, 11) = .Fields("PNJ_DEP").Value
                            oSheet.Cells(i + 1, 11).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 11).NumberFormat = "#,##0"
                            oSheet.Cells(i + 1, 12) = .Fields("TTLPOT_DEP").Value
                            oSheet.Cells(i + 1, 12).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 12).NumberFormat = "#,##0"
                            oSheet.Cells(i + 1, 13) = .Fields("SAL_DEP").Value
                            oSheet.Cells(i + 1, 13).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 13).NumberFormat = "#,##0"
                            oSheet.Cells(i + 1, 14) = .Fields("TNJ_DEP").Value
                            oSheet.Cells(i + 1, 14).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 14).NumberFormat = "#,##0"
                            oSheet.Cells(i + 1, 15) = .Fields("TRN_DEP").Value
                            oSheet.Cells(i + 1, 15).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 15).NumberFormat = "#,##0"
                            oSheet.Cells(i + 1, 16) = .Fields("MKN_DEP").Value
                            oSheet.Cells(i + 1, 16).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 16).NumberFormat = "#,##0"
                            oSheet.Cells(i + 1, 17) = .Fields("BNS_DEP").Value
                            oSheet.Cells(i + 1, 17).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 17).NumberFormat = "#,##0"
                            oSheet.Cells(i + 1, 18) = .Fields("ENT_DEP").Value
                            oSheet.Cells(i + 1, 18).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 18).NumberFormat = "#,##0"
                            oSheet.Cells(i + 1, 19) = .Fields("TTLGAJI_DEP").Value
                            oSheet.Cells(i + 1, 19).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 19).NumberFormat = "#,##0"
                            oSheet.Cells(i + 1, 20) = .Fields("GJITRMA_DEP").Value
                            oSheet.Cells(i + 1, 20).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 20).NumberFormat = "#,##0"
                            oSheet.Columns.AutoFit
                            i = i + 1
                         End If
                       .MoveNext
                     Loop
                 End If
             End With
            
             Set rsRpt2 = Nothing
End If
If True Then '===== SHEET 2 ==============
On Error Resume Next
        Set oSheet = oWkbk.Sheets("sheet2")
            oSheet.Name = "MONTHY CABANG"
            oSheet.Cells.Font.Name = "Utsaah"
            oSheet.Cells.Justify
            oSheet.Cells.Font.Color = &H8000000F
            oSheet.Columns.AutoFit
            oSheet.Cells(1, 1) = "PT. WANINDO PRIMA"
            oSheet.Cells(2, 1) = "PAYROLL MONTHLY SESSION CLOSE"
            oSheet.Cells.Font.Size = 9
            HdrRow = 5
           For clmInt = 1 To 19
                For HdrRow = 5 To 6
                    oSheet.Cells(HdrRow, clmInt).Interior.ColorIndex = 15
                    oSheet.Cells(HdrRow, clmInt).Borders.LineStyle = 1
                    oSheet.Cells(HdrRow, clmInt).Borders.Color = &H8000000F
                    oSheet.Cells(HdrRow, clmInt).Font.Color = &H8000000F
                    oSheet.Cells(HdrRow, clmInt).BorderAround , , &HC0C0C0
                    oSheet.Cells(HdrRow, clmInt).Font.Bold = True
                    oSheet.Cells(HdrRow, clmInt).Font.Size = 9
                    oSheet.Cells(HdrRow, clmInt).HorizontalAlignment = 3
                Next HdrRow
            Next clmInt
            oSheet.Range(oSheet.Cells(5, 1), oSheet.Cells(6, 1)).Merge
            oSheet.Cells(5, 1).ColumnWidth = 6
            oSheet.Cells(5, 1).VerticalAlignment = 2
            oSheet.Cells(5, 1).HorizontalAlignment = 3
            oSheet.Range(oSheet.Cells(5, 2), oSheet.Cells(5, 5)).Merge
            oSheet.Range(oSheet.Cells(5, 6), oSheet.Cells(5, 11)).Merge
            oSheet.Range(oSheet.Cells(5, 12), oSheet.Cells(5, 18)).Merge
            oSheet.Range(oSheet.Cells(5, 19), oSheet.Cells(6, 19)).Merge
            oSheet.Cells(5, 19).ColumnWidth = 6
            oSheet.Cells(5, 19).VerticalAlignment = 2
            oSheet.Cells(5, 19).HorizontalAlignment = 3
            
            HdrRow = 6
            oSheet.Cells(5, 1) = "CABANG"
            oSheet.Cells(5, 2) = "JUMLAH HARI DALAM SEBULAN"
            oSheet.Cells(HdrRow, 2) = "TOTAL OT DEPAN"
            oSheet.Cells(HdrRow, 3) = "TOTAL SIANG"
            oSheet.Cells(HdrRow, 4) = "TOTAL OT BELAKANG"
            oSheet.Cells(HdrRow, 5) = "TOTAL UPAH PER CABANG"
            oSheet.Cells(5, 6) = "JUMLAH KOMPONEN POTONGAN"
            oSheet.Cells(HdrRow, 6) = "TELAT" 'hanya untuk staff
            oSheet.Cells(HdrRow, 7) = "PPH21"
            oSheet.Cells(HdrRow, 8) = "JAMSOS"
            oSheet.Cells(HdrRow, 9) = "ASURANSI"
            oSheet.Cells(HdrRow, 10) = "PINJAMAN"
            oSheet.Cells(HdrRow, 11) = "TOTAL_POTONGAN"
            oSheet.Cells(5, 12) = "JUMLAH KOMPONEN GAJI"
            oSheet.Cells(HdrRow, 12) = "GAJI POKOK"
            oSheet.Cells(HdrRow, 13) = "TUNJANGAN"
            oSheet.Cells(HdrRow, 14) = "TRANSPORT"
            oSheet.Cells(HdrRow, 15) = "MAKAN"
            oSheet.Cells(HdrRow, 16) = "BONUS"
            oSheet.Cells(HdrRow, 17) = "ENTERTAIN"
            oSheet.Cells(HdrRow, 18) = "TOTAL GAJI"
            oSheet.Cells(5, 19) = "GAJI DITERIMA"
             i = HdrRow
             Dim Qry_Cab As String
             Dim SttCab As String
             Qry_Cab = "Vsr_Report('Payroll_MonthlyExcel=" & STgl & "=" & ETgl & "=" & DepId & "=" & KarId & "=" & GrpId & "=" & CabId & "')"
             OpnRecRpt2 Qry_Cab, True
              With rsRpt2
                 If .RecordCount <> 0 Then
                     .MoveFirst
                     Do Until .EOF
                        If SttCab <> .Fields("CAB_ID").Value Then
                            SttCab = .Fields("CAB_ID").Value
                            oSheet.Cells(3, 1) = "PERIODE : " & .Fields("sTgl").Value & "  TO  " & .Fields("eTgl").Value
           
                            For clmInt = 1 To 19 'TglInt + 2=tgl loop,
                                    'oSheet.Cells(i + 1, clmInt).Interior.ColorIndex = 15
                                    oSheet.Cells(i + 1, clmInt).Borders.LineStyle = 1
                                    oSheet.Cells(i + 1, clmInt).Borders.Color = &H8000000F
                                    oSheet.Cells(i + 1, clmInt).Font.Color = &H8000000F
                                    oSheet.Cells(i + 1, clmInt).BorderAround , , &HC0C0C0
                                    oSheet.Cells(i + 1, clmInt).Font.Bold = True
                                    oSheet.Cells(i + 1, clmInt).Font.Size = 9
                                    oSheet.Cells(i + 1, clmInt).HorizontalAlignment = 3
                            Next clmInt
                            
                            oSheet.Cells(i + 1, 1) = .Fields("CAB_NM").Value
                            oSheet.Cells(i + 1, 1).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 2) = .Fields("DPN_CAB").Value
                            oSheet.Cells(i + 1, 2).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 3) = .Fields("PAGI_CAB").Value
                            oSheet.Cells(i + 1, 3).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 4) = .Fields("BLK_CAB").Value
                            oSheet.Cells(i + 1, 4).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 5) = .Fields("UPAHBLN_CAB").Value
                            oSheet.Cells(i + 1, 5).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 5).NumberFormat = "#,##0"
                            oSheet.Cells(i + 1, 6) = "0" 'telat untuk staff
                            oSheet.Cells(i + 1, 6).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 6).NumberFormat = "#,##0"
                            oSheet.Cells(i + 1, 7) = .Fields("PPH21_CAB").Value
                            oSheet.Cells(i + 1, 7).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 7).NumberFormat = "#,##0"
                            oSheet.Cells(i + 1, 8) = .Fields("JSOS_CAB").Value
                            oSheet.Cells(i + 1, 8).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 8).NumberFormat = "#,##0"
                            oSheet.Cells(i + 1, 9) = .Fields("ASR_CAB").Value
                            oSheet.Cells(i + 1, 9).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 9).NumberFormat = "#,##0"
                            oSheet.Cells(i + 1, 10) = .Fields("PNJ_CAB").Value
                            oSheet.Cells(i + 1, 10).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 10).NumberFormat = "#,##0"
                            oSheet.Cells(i + 1, 11) = .Fields("TTLPOT_CAB").Value
                            oSheet.Cells(i + 1, 11).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 11).NumberFormat = "#,##0"
                            oSheet.Cells(i + 1, 12) = .Fields("SAL_CAB").Value
                            oSheet.Cells(i + 1, 12).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 12).NumberFormat = "#,##0"
                            oSheet.Cells(i + 1, 13) = .Fields("TNJ_CAB").Value
                            oSheet.Cells(i + 1, 13).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 13).NumberFormat = "#,##0"
                            oSheet.Cells(i + 1, 14) = .Fields("TRN_CAB").Value
                            oSheet.Cells(i + 1, 14).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 14).NumberFormat = "#,##0"
                            oSheet.Cells(i + 1, 15) = .Fields("MKN_CAB").Value
                            oSheet.Cells(i + 1, 15).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 15).NumberFormat = "#,##0"
                            oSheet.Cells(i + 1, 16) = .Fields("BNS_CAB").Value
                            oSheet.Cells(i + 1, 16).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 16).NumberFormat = "#,##0"
                            oSheet.Cells(i + 1, 17) = .Fields("ENT_CAB").Value
                            oSheet.Cells(i + 1, 17).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 17).NumberFormat = "#,##0"
                            oSheet.Cells(i + 1, 18) = .Fields("TTLGAJI_CAB").Value
                            oSheet.Cells(i + 1, 18).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 18).NumberFormat = "#,##0"
                            oSheet.Cells(i + 1, 19) = .Fields("GJITRMA_CAB").Value
                            oSheet.Cells(i + 1, 19).HorizontalAlignment = 1
                            oSheet.Cells(i + 1, 19).NumberFormat = "#,##0"
                            oSheet.Columns.AutoFit
                            i = i + 1
                         End If
                       .MoveNext
                     Loop
                 End If
             End With
             Set rsRpt2 = Nothing
End If
If True Then '===== SHEET 3 ==============
'On Error Resume Next
        Set oSheet = oWkbk.Sheets("sheet3")
            oSheet.Name = "SLIP GAJI"
            oSheet.Cells.Font.Name = "Utsaah"
            oSheet.Cells.Justify
            oSheet.Cells.Font.Color = &H8000000F
            oSheet.Columns.AutoFit
           ' oSheet.Range(oSheet.Cells(1, 1), oSheet.Cells(3, 2)).Merge
            
    Dim Qry_Slip As String
    Dim Lurus, bawah As Double
    Qry_Slip = "Vsr_Report('Payroll_MonthlyExcel=" & STgl & "=" & ETgl & "=" & DepId & "=" & KarId & "=" & GrpId & "=" & CabId & "')"
    'Qry_Slip = "Vsr_Report('Payroll_MonthlyExcel=2015-03-01=2015-03-31=0=WAN.HO.000087=0=0')"
    OpnRecRpt2 Qry_Slip, True
    Lurus = 7
    bawah = 22
    i = 1
    trecs = 0
     With rsRpt2
        If .RecordCount <> 0 Then
            
                    .MoveFirst
                    Do Until .EOF
                        trecs = trecs + 1
                    .MoveNext
                    
                    Loop
                    
                    '=====================
                    '--- PROGRESS BAR ----
                    '=====================
                    If frm_actv = 1 Then '=== PRESENSI ===
                        With FRM_ABSENSI_REPORT
                            .ProgressBar1.Visible = True
                            .ProgressBar1.Min = 0
                            .ProgressBar1.Max = trecs
                        End With
                    ElseIf frm_actv = 2 Then '=== PAYROLL ===
                        With FRM_PAYROLL_REPORT
                            .ProgressBar1.Visible = True
                            .ProgressBar1.Min = 0
                            .ProgressBar1.Max = trecs
                        End With
                    ElseIf frm_actv = 3 Then '=== ACCOUNTING ===
                        With FRM_ABSENSI_REPORT
                            .ProgressBar1.Visible = True
                            .ProgressBar1.Min = 0
                            .ProgressBar1.Max = trecs
                        End With
                    End If
            
            .MoveFirst
            Do Until .EOF
              If Lurus = 35 Then
                Lurus = 7
                bawah = bawah + 22
              End If
                'oSheet.Range(oSheet.Cells(bawah - 21, Lurus - 6), oSheet.Cells(bawah, Lurus)).BorderAround LineStyle:=Excel.XlLineStyle.xlContinuous, Weight:=Excel.XlBorderWeight.xlThick
                With oSheet.Range(oSheet.Cells(bawah - 21, Lurus - 6), oSheet.Cells(bawah, Lurus))
                    .Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Borders(xlEdgeRight).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                End With
  
                With oSheet.Pictures.Insert(App.Path & "\Img\Logo\wanindo.jpg")
                     With .ShapeRange
                         .LockAspectRatio = True
                         .Width = 78
                         .Height = 50
                     End With
                        '.left = 2
                        '.top = 3
                      .left = (oSheet.Cells(bawah - 21, Lurus - 6).left) + 2
                      .top = (oSheet.Cells(bawah - 21, Lurus - 6).top) + 2
                     .Placement = 2
                     .PrintObject = True
                 End With

                 oSheet.Cells(bawah - 20, Lurus - 1) = "PT. WANINDO PRIMA"
                 oSheet.Cells(bawah - 20, Lurus - 1).HorizontalAlignment = -4152
                 oSheet.Range(oSheet.Cells(bawah - 19, Lurus - 4), oSheet.Cells(bawah - 19, Lurus - 2)).Merge
                 oSheet.Cells(bawah - 19, Lurus - 4) = "PAYROLL " & MonthName(Month(.Fields("eTgl").Value)) & " " & Year(.Fields("eTgl").Value)
                 oSheet.Cells(bawah - 19, Lurus - 4).HorizontalAlignment = -4108
                 oSheet.Cells.Font.Size = 9
                 
                 oSheet.Cells(bawah - 17, Lurus - 5) = "EmployeID     "
                 oSheet.Cells(bawah - 16, Lurus - 5) = "Employe Name  "
                 oSheet.Cells(bawah - 15, Lurus - 5) = "Position      "
                 oSheet.Cells(bawah - 17, Lurus - 3) = " : " & .Fields("KAR_ID").Value
                 oSheet.Cells(bawah - 16, Lurus - 3) = " : " & .Fields("KAR_NM").Value
                 oSheet.Cells(bawah - 15, Lurus - 3) = " : " & .Fields("JAB_NM").Value & "/" & .Fields("DEP_NM").Value
                 'PENDAPATAN
                 oSheet.Cells(bawah - 13, Lurus - 5) = "PENDAPATAN :"
                 oSheet.Cells(bawah - 12, Lurus - 5) = "Jumlah Hari :"
                 oSheet.Cells(bawah - 11, Lurus - 5) = "Upah Hari :"
                 oSheet.Cells(bawah - 10, Lurus - 5) = "Upah Bulan :"
                 oSheet.Cells(bawah - 9, Lurus - 5) = "Gaji Pokok :"
                 oSheet.Cells(bawah - 8, Lurus - 5) = "Tunjangan :"
                 oSheet.Cells(bawah - 7, Lurus - 5) = "Tranport :"
                 oSheet.Cells(bawah - 6, Lurus - 5) = "Uang makan :"
                 oSheet.Cells(bawah - 5, Lurus - 5) = "Entertain :"
                 oSheet.Cells(bawah - 4, Lurus - 5) = "Bonus :"
                 oSheet.Cells(bawah - 3, Lurus - 5) = "Total Pendapatan"
                 
                 '-Value Field ------
                 oSheet.Cells(bawah - 12, Lurus - 4) = Round(.Fields("TTL_HARI").Value, 2)
                 oSheet.Cells(bawah - 11, Lurus - 4) = .Fields("PAY_DAY").Value
                 oSheet.Cells(bawah - 10, Lurus - 4) = .Fields("TTL_UPAH_BLN").Value
                 oSheet.Cells(bawah - 9, Lurus - 4) = .Fields("PAY_SALARY").Value
                 oSheet.Cells(bawah - 8, Lurus - 4) = .Fields("PAY_TUNJANGAN").Value
                 oSheet.Cells(bawah - 7, Lurus - 4) = .Fields("PAY_TRANSPORT").Value
                 oSheet.Cells(bawah - 6, Lurus - 4) = .Fields("PAY_EAT").Value
                 oSheet.Cells(bawah - 5, Lurus - 4) = .Fields("PAY_BONUS").Value
                 oSheet.Cells(bawah - 4, Lurus - 4) = .Fields("PAY_ENTERTAIN").Value
                 oSheet.Cells(bawah - 3, Lurus - 4) = .Fields("TOTAL_GAJI").Value + .Fields("TTL_UPAH_BLN").Value
                
                 For HdrRow = (bawah - 11) To (bawah - 3)
                     oSheet.Cells(HdrRow, Lurus - 5).HorizontalAlignment = -4152
                     oSheet.Cells(HdrRow, Lurus - 6).ColumnWidth = 2
                     oSheet.Cells(HdrRow, Lurus - 5).ColumnWidth = 7.29
                     oSheet.Cells(HdrRow, Lurus - 4).ColumnWidth = 7.29
                     oSheet.Cells(HdrRow, Lurus - 4).NumberFormat = "#,##0"
                 Next HdrRow
                 'POTONGAN
                 oSheet.Cells(bawah - 13, Lurus - 2) = "POTONGAN"
                 oSheet.Cells(bawah - 12, Lurus - 2) = "Telat :"
                 oSheet.Cells(bawah - 11, Lurus - 2) = "PPH21 :"
                 oSheet.Cells(bawah - 10, Lurus - 2) = "Jamsos :"
                 oSheet.Cells(bawah - 9, Lurus - 2) = "Asuransi :"
                 oSheet.Cells(bawah - 8, Lurus - 2) = "Pinjaman :"
                 oSheet.Cells(bawah - 3, Lurus - 2) = "Total Potongan"
                 '---- value field -- ,A.,A.,A.POT_PINJAMAN
                 oSheet.Cells(bawah - 12, Lurus - 1) = .Fields("Telat").Value
                 oSheet.Cells(bawah - 11, Lurus - 1) = .Fields("POT_PPH21").Value
                 oSheet.Cells(bawah - 10, Lurus - 1) = .Fields("POT_JAMSOS_JHT_TK").Value
                 oSheet.Cells(bawah - 9, Lurus - 1) = .Fields("POT_ASURANSI").Value
                 oSheet.Cells(bawah - 8, Lurus - 1) = .Fields("POT_PINJAMAN").Value
                 oSheet.Cells(bawah - 3, Lurus - 1) = .Fields("TOTAL_POT").Value
    
                 For HdrRow = (bawah - 12) To (bawah - 3)
                     oSheet.Cells(HdrRow, Lurus - 2).HorizontalAlignment = -4152
                     oSheet.Cells(HdrRow, Lurus - 3).ColumnWidth = 2
                     oSheet.Cells(HdrRow, Lurus - 2).ColumnWidth = 7.29
                     oSheet.Cells(HdrRow, Lurus - 1).ColumnWidth = 7.29
                     oSheet.Cells(HdrRow, Lurus).ColumnWidth = 2
                     oSheet.Cells(HdrRow, Lurus - 1).NumberFormat = "#,##0"
                 Next HdrRow
                 oSheet.Cells(bawah - 1, Lurus - 4) = "TAKE HOME PAY"
                 oSheet.Cells(bawah - 1, Lurus - 2) = .Fields("GAJI_DITERIMA").Value
                 oSheet.Cells(bawah - 1, Lurus - 2).HorizontalAlignment = -4152
                 oSheet.Cells(bawah - 1, Lurus - 2).NumberFormat = "#,##0"
                 'Lurus
                 'oSheet.Range(oSheet.Cells(bawah - 1, Lurus - 5):oSheet.Cells(bawah - 1, Lurus - 1)).BorderAround LineStyle:=3, Weight:=xlThick
                 'oSheet.Range(oSheet.Cells(bawah - 1, Lurus - 5), oSheet.Cells(bawah - 1, Lurus - 1)).BorderAround ColorIndex:=3, Weight:=xlThick
                    'oSheet.Range(oSheet.Cells(bawah - 1, Lurus - 5), oSheet.Cells(bawah - 1, Lurus - 1)).Borders.Item
                With oSheet.Range(oSheet.Cells(bawah - 1, Lurus - 5), oSheet.Cells(bawah - 1, Lurus - 1))
                    .Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Borders(xlEdgeRight).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                End With
                 
                 
                  Lurus = Lurus + 7
                    '=====================
                    '--- PROGRESS BAR ----
                    '=====================
                    If frm_actv = 1 Then '=== PRESENSI ===
                        FRM_ABSENSI_REPORT.ProgressBar1.Value = i
                    ElseIf frm_actv = 2 Then '=== PAYROLL ===
                        FRM_PAYROLL_REPORT.ProgressBar1.Value = i
                    ElseIf frm_actv = 3 Then '=== ACCOUNTING ===
                        FRM_ABSENSI_REPORT.ProgressBar1.Value = i
                    End If
                    i = i + 1
                 .MoveNext
                 Loop
        End If

    End With
End If
End Sub





