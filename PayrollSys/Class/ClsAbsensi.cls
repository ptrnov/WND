VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsAbsensi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim Qry1 As String, Qry2 As String, Qry3 As String
Dim QryStt1 As String, QryStt2 As String, QryStt3 As String
'=====================================
'================ ptr.nov ============
'============ CONFIGURATION ==========
'=====================================
Public Sub Get_Setting(ByRef PERSENSI_STGL As Integer, _
                       ByRef F_TGL As String, ByRef F_WAKTU As String, ByRef F_TGL_WAKTU As String)
    QryStt1 = ""
    QryStt1 = "pro_get_setting()"
    OpRecStt1 QryStt1, True
    With rsStt1
        If Not .EOF Then
            .MoveFirst
             Do Until .EOF
                'PERSENSI_STGL,FORMAT_TGL,FORMAT_WAKTU,FORMAT_TGL_WAKTU
                PERSENSI_STGL = .Fields("PERSENSI_STGL").Value
                F_TGL = .Fields("FORMAT_TGL").Value
                F_WAKTU = .Fields("FORMAT_WAKTU").Value
                F_TGL_WAKTU = .Fields("FORMAT_TGL_WAKTU").Value
           .MoveNext
         Loop
        End If
    End With
    Set rsStt1 = Nothing
 End Sub

'=====================================
'============= ptr.nov ===============
'=========== GROUP Time Table ========
'=====================================
Function GroupTT(KarId As String) As Integer
QryStt1 = "SELECT TT_GRP_ID FROM tab_timetabel_kar WHERE KAR_ID='" & KarId & "';"
OpRecStt1 QryStt1, False
'MsgBox datetime1 & "---" & datetime2
With rsStt1
    If Not .EOF Then
        GroupTT = .Fields("TT_GRP_ID").Value
        'MsgBox Get_SelisihWaktu
    End If
End With
Set rsStt1 = Nothing
End Function

'=====================================
'============= ptr.nov ===============
'=========== SELISIH WAKTU ===========
'=====================================
Function Get_SelisihWaktu(datetime1 As String, datetime2 As String) As String
QryStt1 = "pro_get_selisihwaktu('" & Format(datetime1, "yyyy-mm-dd hh:mm:ss") & "','" & Format(datetime2, "yyyy-mm-dd hh:mm:ss") & "')"
OpRecStt1 QryStt1, True
'MsgBox datetime1 & "---" & datetime2
With rsStt1
    If Not .EOF Then
        Get_SelisihWaktu = .Fields("Rslt").Value
    Else
       Get_SelisihWaktu = ""
       'MsgBox Get_SelisihWaktu
    End If
End With
Set rsStt1 = Nothing
End Function

  
'=====================================
'============= ptr.nov ===============
'======== JOIN DATE & WAKTU ==========
'=====================================
Function Join_DateTime(Tgl As String, waktu As String) As String
Dim tglwaktu As Date
tglwaktu = Tgl & " " & waktu
'MsgBox tglwaktu
Join_DateTime = Format(tglwaktu, "yyyy-mm-dd hh:mm:ss")
End Function


'=====================================
'============= ptr.nov ===============
'======== RUN_ABSENSI CALCULATE ======
'=====================================
Public Sub RUN_ABSENSI(GrpId As Integer, KarId As String, DepId As Integer, STgl As Date, ETgl As Date)
Dim TglRun As Date
Dim ArryLog As String, ArryTT As String, ValueCalculate As String
Dim Day_Active As Integer, Day_Absen As Integer ', Day_Libur As Integer, Day_ExcpID As Integer
Dim Day_WK As String, Day_Late As String, Day_Early As String, OT_DPN As String, OT_BLK As String
Dim trecs, cpt As Double
Dim x As Integer
For TglRun = STgl To ETgl
    '--------------------------------------------------------------------------------------
    '------------------------|DATA PERSONAL LOG|-------------------------------------------
    '-- pro_get_personal_log_run('2015-01-01','2015-01-03','2015-01-01','WAN.HOW.00003') --
    '--------------------------------------------------------------------------------------
    Qry1 = "pro_get_personal_log_run('" & Format(STgl, "yyyy-mm-dd") & _
            "','" & Format(ETgl, "yyyy-mm-dd") & _
            "','" & Format(TglRun, "yyyy-mm-dd") & _
            "','" & KarId & "')"
            'MsgBox Qry1
    OpRec1 Qry1, True
     With rs1
        If Not .EOF Then
             .MoveFirst:
            Do Until .EOF
                trecs = trecs + 1
             .MoveNext
            Loop

            With FRM_ABSENSI_PBAR.ProgressBar1(1)
                FRM_ABSENSI_PBAR.ProgressBar1(0).Value = FRM_ABSENSI_PBAR.ProgressBar1(0).Value
                '.Visible = True
                .Min = 0
                .Max = trecs
            End With
         'x = 0
        .MoveFirst
         Do Until .EOF
                '--|ARRAY VARIABLE PERSONAL LOG|----
                ArryTT = ""
                ArryLog = ""
                ArryLog = Format(.Fields("DTLogIn").Value, "yyyy-mm-dd hh:mm:ss") & _
                          "=" & Format(.Fields("DTLogOUT").Value, "yyyy-mm-dd hh:mm:ss") & _
                          "=" & .Fields("TT_GRP_ID").Value & _
                          "=" & .Fields("LVL_ID").Value
                         ' MsgBox ArryLog
            .MoveNext
            Loop
         End If
    End With
    Set rs1 = Nothing
    Qry1 = ""
        '--------------------------------------------------------
        '------------|INSERT PERSENSI LEVEL2|--------------------
        '-- pro_get_personal_log_ins(TglRun,ArryLog,ArryTT) -----
        '--------------------------------------------------------
        'Qry3 = "pro_get_personal_log_ins('" & Format(TglRun, "yyyy-mm-dd") & "','" & KarId & "','" & ArryLog & "','" & ArryTT & "','" & ValueCalculate & "')"
        Qry3 = "pro_get_personal_log_ins('" & Format(TglRun, "yyyy-mm-dd") & "','" & KarId & "','" & ArryLog & "')"
        OpRec3 Qry3, True
        'With rs3
        '    If Not .EOF Then
        '        .MoveFirst
        '        Do Until .EOF
                    FRM_ABSENSI_PBAR.ProgressBar1(1).Value = x
                    FRM_ABSENSI_PBAR.ProgressBar1(0).Value = FRM_ABSENSI_PBAR.ProgressBar1(0).Value
                    x = x + 1
         '       .MoveNext
         '       Loop
         '   End If
        'End With
        Set rs3 = Nothing
        Qry3 = ""
        ArryLog = ""
Next TglRun
'Unload FRM_ABSENSI_PBAR
End Sub



'=====================================
'============= ptr.nov ===============
'======== RUN_overtime CALCULATE =====
'=====================================
Public Sub RUN_OT(KarId As String, STgl As Date, ETgl As Date)
Dim TglRun, TglRun1 As Date
Dim trecs1, cpt As Double
Dim y As Integer
trecs1 = 0
y = 0
For TglRun = STgl To ETgl
    '--------------------------------------------------------------------------------------
    '------------------------|DATA PERSONAL LOG|-------------------------------------------
    '-- pro_get_personal_log_run('2015-01-01','2015-01-03','2015-01-01','WAN.HOW.00003') --
    '--------------------------------------------------------------------------------------
    Qry2 = "pro_get_personal_log_ins_ot('" & KarId & _
           "','" & Format(TglRun, "yyyy-mm-dd") & _
           "','" & Format(STgl, "yyyy-mm-dd") & _
           "','" & Format(ETgl, "yyyy-mm-dd") & "')"
            'MsgBox Qry1
    OpRec2 Qry2, True
    'For TglRun1 = STgl To ETgl
    '    trecs1 = trecs1 + 1
    'Next TglRun1
    'MsgBox trecs1
    'With FRM_ABSENSI_PBAR.ProgressBar1(1)
        'FRM_ABSENSI_PBAR.ProgressBar1(0).Value = FRM_ABSENSI_PBAR.ProgressBar1(0).Value
    '    .Min = 0
    '    .Max = trecs1
    'End With
    
    'FRM_ABSENSI_PBAR.ProgressBar1(1).Value = y
    'FRM_ABSENSI_PBAR.ProgressBar1(0).Value = FRM_ABSENSI_PBAR.ProgressBar1(0).Value
    'y = y + 1
    Set rs2 = Nothing
Next TglRun
End Sub










