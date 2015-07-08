VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsIdCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Dim Qry As String
Dim RoNew(3), RoLast, BrgLast, RoDep, ROtgl, RoCab, RoPT As String
Dim rsRec As ADODB.Recordset
Dim i, L(2) As Integer

'===========================================
'============ ptr.nov Payroll ==============
'=========== Employe Per dept===============
'===========================================
Public Sub Employe(Cmb As XtremeSuiteControls.ComboBox, CabId As String, GrpId As Integer, Dep_id As Integer)
Dim strqryx As String
Dim i As Integer
strqry = ""
    'If Dep_id = 0 Then
    '    strqry = "SELECT KAR_ID,KAR_NM FROM karyawan order by KAR_NM ASC"
    'Else
    '    strqry = "SELECT KAR_ID,KAR_NM FROM karyawan WHERE DEP_ID=" & Dep_id & " order by KAR_NM ASC"
    'End If
     strqryx = "find_employe('" & CabId & "'," & GrpId & "," & Dep_id & ")"
     'MsgBox strqry
OpRecStt1 strqryx, True
Cmb.Clear
Cmb.Style = xtpComboDropDownList
Cmb.EnableMarkup = True
Cmb.AddItem 0
Cmb.MarkupList(Cmb.NewIndex) = "<StackPanel Orientation='Horizontal'>" & _
                                "<Rectangle  Fill='#FC031C' Stroke='#1E0DF2' Width='14' Height='14'/>" & _
                                "<TextBlock Padding = '2' VerticalAlignment='Center'>" & "ALL" & "</TextBlock></StackPanel>"  '#FC031C

i = 0
With rsStt1
    If Not .EOF Then
        .MoveFirst
        Do Until .EOF
            'cmb.AddItem .Fields("KAR_NM").Value
            Cmb.AddItem .Fields("KAR_ID").Value
            'Cmb.MarkupList(Cmb.NewIndex) = "<Grid><Grid.ColumnDefinitions><ColumnDefinition Width='10'/><ColumnDefinition Width='200'/></Grid.ColumnDefinitions><Grid.RowDefinitions><RowDefinition/></Grid.RowDefinitions>" & _
                                  "<TextBlock Grid.Column='2'> " & .Fields("KAR_NM").Value & "</TextBlock></Grid>"
            Cmb.MarkupList(Cmb.NewIndex) = "<StackPanel Orientation='Horizontal'>" & _
                                            "<Rectangle  Fill='#FC031C' Stroke='#1E0DF2' Width='14' Height='14'/>" & _
                                            "<TextBlock Padding = '2' VerticalAlignment='Center'>" & .Fields("KAR_NM").Value & "</TextBlock></StackPanel>" '#FC031C
       .MoveNext
        i = i + 1
        Loop
    End If
        If Cmb.ListCount <> 0 Then Cmb.ListIndex = 0
End With
Set rsStt1 = Nothing
End Sub

'===========================================
'============ ptr.nov Payroll ==============
'=========== Employe Per dept===============
'=========================================== IN CabId varchar(5), IN Grpid Int, In DepId int
Public Sub EmployeNoAll(Cmb As XtremeSuiteControls.ComboBox, CabId As String, GrpId As Integer, Dep_id As Integer)
Dim strqry As String
Dim i As Integer
strqry = ""
    'If Dep_id = 0 Then
    '    strqry = "SELECT KAR_ID,KAR_NM FROM karyawan Order By KAR_NM ASC "
    'Else
    '    strqry = "SELECT KAR_ID,KAR_NM FROM karyawan WHERE DEP_ID=" & Dep_id
    'End If
    strqry = "find_employe('" & CabId & "','" & GrpId & "','" & Dep_id & "')"
OpRecStt1 strqry, True
Cmb.Clear
Cmb.Style = xtpComboDropDownList
Cmb.EnableMarkup = True
i = 0
With rsStt1
    If Not .EOF Then
        .MoveFirst
        Do Until .EOF
            'cmb.AddItem .Fields("KAR_NM").Value
            Cmb.AddItem .Fields("KAR_ID").Value
           ' Cmb.MarkupList(Cmb.NewIndex) = "<Grid><Grid.ColumnDefinitions><ColumnDefinition Width='10'/><ColumnDefinition Width='200'/></Grid.ColumnDefinitions><Grid.RowDefinitions><RowDefinition/></Grid.RowDefinitions>" & _
                                  "<TextBlock Grid.Column='2'> " & .Fields("KAR_NM").Value & "</TextBlock></Grid>"
            Cmb.MarkupList(Cmb.NewIndex) = "<StackPanel Orientation='Horizontal'>" & _
                                        "<Rectangle  Fill='#FC031C' Stroke='#1E0DF2' Width='14' Height='14'/>" & _
                                        "<TextBlock Padding = '2' VerticalAlignment='Center'>" & .Fields("KAR_NM").Value & "</TextBlock></StackPanel>" '#FC031C

       .MoveNext
        i = i + 1
        Loop
    End If
        If Cmb.ListCount <> 0 Then Cmb.ListIndex = 0
End With
Set rsStt1 = Nothing
End Sub

'===========================================
'============== ptr.nov Payroll ============
'================= DEPT ====================
'===========================================
Public Sub Dept(Cmb As XtremeSuiteControls.ComboBox)
Dim strqry As String
Dim sColor As String
Dim lRed As Long, lGreen As Long, lBlue As Long



strqry = ""
    strqry = "SELECT DEP_ID,DEP_NM FROM departemen ORDER BY DEP_NM ASC"
OpRecStt1 strqry, False
Cmb.Clear
Cmb.Style = xtpComboDropDownList
Cmb.EnableMarkup = True
Cmb.AddItem 0
Cmb.MarkupList(Cmb.NewIndex) = "<StackPanel Orientation='Horizontal'>" & _
            "<Rectangle  Fill='#1E0DF2' Stroke='#1E0DF2' Width='14' Height='14'/>" & _
            "<TextBlock Padding = '2' VerticalAlignment='Center'>" & "ALL" & "</TextBlock></StackPanel>"  '#FC031C

With rsStt1
    If Not .EOF Then
        .MoveFirst
        
        Do Until .EOF
        lRed = lColor And &HFF
lGreen = (lColor And &HFF00&) \ &H100&
lBlue = (lColor And &HFF0000) \ &H10000

        sColor = right$("0" & Hex$(lRed), 2) & right$("0" & Hex$(lGreen), 2) & right$("0" & Hex$(lBlue), 2)

             Cmb.AddItem .Fields("DEP_ID").Value
            ' Cmb.MarkupList(Cmb.NewIndex) = .Fields("DEP_NM").Value
            Cmb.MarkupList(Cmb.NewIndex) = "<StackPanel Orientation='Horizontal'>" & _
            "<Rectangle  Fill='#1E0DF2' Stroke='#1E0DF2' Width='14' Height='14'/>" & _
            "<TextBlock Padding = '2' VerticalAlignment='Center'>" & .Fields("DEP_NM").Value & "</TextBlock></StackPanel>" '#FC031C
            
           ' Cmb.MarkupList(Cmb.NewIndex) = "<Grid><Grid.ColumnDefinitions><ColumnDefinition Width='10'/><ColumnDefinition Width='250'/></Grid.ColumnDefinitions><Grid.RowDefinitions><RowDefinition/></Grid.RowDefinitions>" & _
                                   "<TextBlock Grid.Column='2'>" & .Fields("DEP_NM").Value & "</TextBlock></Grid>"
           ' Cmb.MarkupList(Cmb.NewIndex) = "<Grid><Grid.ColumnDefinitions><ColumnDefinition Width='10'/><ColumnDefinition Width='250'/></Grid.ColumnDefinitions><Grid.RowDefinitions><RowDefinition/></Grid.RowDefinitions>" & _
                                   "<TextBlock Grid.Column='2'>" & .Fields("DEP_NM").Value & "</TextBlock></Grid>"
            'Cmb.MarkupList(Cmb.NewIndex) = "<Grid><Grid.ColumnDefinitions><ColumnDefinition Width='10'/><ColumnDefinition Width='250'/></Grid.ColumnDefinitions><Grid.RowDefinitions><RowDefinition/></Grid.RowDefinitions>" & _
                                  "<TextBlock>" & .Fields("DEP_ID").Value & "</TextBlock>" & _
                                  "<TextBlock Grid.Column='2'> | " & .Fields("DEP_NM").Value & "</TextBlock></Grid>"
       .MoveNext
        Loop
    End If
        If Cmb.ListCount <> 0 Then Cmb.ListIndex = 0
End With
Set rsStt1 = Nothing
End Sub

'===========================================
'============== ptr.nov Payroll ============
'================= DEPT ====================
'===========================================
Public Sub DeptNoAll(Cmb As XtremeSuiteControls.ComboBox)
Dim strqry As String
strqry = ""
    strqry = "SELECT DEP_ID,DEP_NM FROM departemen ORDER BY DEP_NM ASC"
OpRecStt1 strqry, False
Cmb.Clear
Cmb.Style = xtpComboDropDownList
Cmb.EnableMarkup = True
                     '
With rsStt1
    If Not .EOF Then
        .MoveFirst
        
        Do Until .EOF
            Cmb.AddItem .Fields("DEP_ID").Value
            Cmb.MarkupList(Cmb.NewIndex) = "<StackPanel Orientation='Horizontal'>" & _
                                            "<Rectangle  Fill='#1E0DF2' Stroke='#1E0DF2' Width='14' Height='14'/>" & _
                                            "<TextBlock Padding = '2' VerticalAlignment='Center'>" & .Fields("DEP_NM").Value & "</TextBlock></StackPanel>" '#FC031C

            'Cmb.MarkupList(Cmb.NewIndex) = "<Grid><Grid.ColumnDefinitions><ColumnDefinition Width='10'/><ColumnDefinition Width='250'/></Grid.ColumnDefinitions><Grid.RowDefinitions><RowDefinition/></Grid.RowDefinitions>" & _
                                   "<TextBlock Grid.Column='2'>" & .Fields("DEP_NM").Value & "</TextBlock></Grid>"
            'cmb.MarkupList(cmb.NewIndex) = "<Grid><Grid.ColumnDefinitions><ColumnDefinition Width='10'/><ColumnDefinition Width='100'/></Grid.ColumnDefinitions><Grid.RowDefinitions><RowDefinition/></Grid.RowDefinitions>" & _
                                  "<TextBlock>" & .Fields("DEP_ID").Value & "</TextBlock>" & _
                                  "<TextBlock Grid.Column='2'> | " & .Fields("DEP_NM").Value & "</TextBlock></Grid>"
       .MoveNext
        Loop
    End If
        If Cmb.ListCount <> 0 Then Cmb.ListIndex = 0
End With
Set rsStt1 = Nothing
End Sub

'===========================================
'============== ptr.nov Payroll ============
'==============FINGER MACHINE  =============
'===========================================
Public Sub FingerMachine(Cmb As XtremeSuiteControls.ComboBox)
Dim strqry As String
strqry = ""
    strqry = "SELECT TerminalID,MESIN_NM,MESIN_SN FROM machine"
OpRecStt1 strqry, False
Cmb.Clear
Cmb.Style = xtpComboDropDownList
Cmb.EnableMarkup = True
Cmb.AddItem 0
Cmb.MarkupList(Cmb.NewIndex) = "<Grid><Grid.ColumnDefinitions><ColumnDefinition Width='10'/><ColumnDefinition Width='100'/></Grid.ColumnDefinitions><Grid.RowDefinitions><RowDefinition/></Grid.RowDefinitions>" & _
                       "<TextBlock Grid.Column='2'>" & "ALL" & "</TextBlock></Grid>"
                      '
With rsStt1
    If Not .EOF Then
        .MoveFirst
        
        Do Until .EOF
            Cmb.AddItem .Fields("MESIN_SN").Value
            'Cmb.MarkupList(Cmb.NewIndex) = "<Grid><Grid.ColumnDefinitions><ColumnDefinition Width='10'/><ColumnDefinition Width='200'/></Grid.ColumnDefinitions><Grid.RowDefinitions><RowDefinition/></Grid.RowDefinitions>" & _
                                  "<TextBlock>" & .Fields("TerminalID").Value & "</TextBlock>" & _
                                  "<TextBlock Grid.Column='2'> | " & .Fields("MESIN_NM").Value & " | " & .Fields("MESIN_SN").Value & "</TextBlock></Grid>"
            Cmb.MarkupList(Cmb.NewIndex) = "<StackPanel Orientation='Horizontal'>" & _
                                  "<Rectangle  Fill='#FF8000' Stroke='#1E0DF2' Width='14' Height='14'/>" & _
                                  "<TextBlock Padding = '2' VerticalAlignment='Center'>" & .Fields("MESIN_NM").Value & "</TextBlock></StackPanel>"  '#FC031C

       .MoveNext
        Loop
    End If
        If Cmb.ListCount <> 0 Then Cmb.ListIndex = 0
End With
Set rsStt1 = Nothing
End Sub
'===========================================

'===========================================
'============== ptr.nov Payroll ============
'==============FINGER MACHINE  =============
'===========================================
Public Sub FingerMachineNoAll(Cmb As XtremeSuiteControls.ComboBox)
Dim strqry As String
strqry = ""
    strqry = "SELECT TerminalID,MESIN_NM,MESIN_SN FROM machine"
OpRecStt1 strqry, False
Cmb.Clear
Cmb.Style = xtpComboDropDownList
Cmb.EnableMarkup = True
                     '
With rsStt1
    If Not .EOF Then
        .MoveFirst
        
        Do Until .EOF
            Cmb.AddItem .Fields("TerminalID").Value
            'Cmb.MarkupList(Cmb.NewIndex) = "<Grid><Grid.ColumnDefinitions><ColumnDefinition Width='10'/><ColumnDefinition Width='200'/></Grid.ColumnDefinitions><Grid.RowDefinitions><RowDefinition/></Grid.RowDefinitions>" & _
                                  "<TextBlock>" & .Fields("TerminalID").Value & "</TextBlock>" & _
                                  "<TextBlock Grid.Column='2'> | " & .Fields("MESIN_NM").Value & " | " & .Fields("MESIN_SN").Value & "</TextBlock></Grid>"
            Cmb.MarkupList(Cmb.NewIndex) = "<StackPanel Orientation='Horizontal'>" & _
                                  "<Rectangle  Fill='#FF8000' Stroke='#1E0DF2' Width='14' Height='14'/>" & _
                                  "<TextBlock Padding = '2' VerticalAlignment='Center'>" & .Fields("MESIN_NM").Value & "</TextBlock></StackPanel>"  '#FC031C

          
       .MoveNext
        Loop
    End If
        If Cmb.ListCount <> 0 Then Cmb.ListIndex = 0
End With
Set rsStt1 = Nothing
End Sub

'===========================================
'============ ptr.nov Payroll ==============
'=========== Group Finger Employe===========
'===========================================
Public Sub FingerEmploye(Cmb As XtremeSuiteControls.ComboBox, CabId As String)
Dim strqry As String
Dim i As Integer
strqry = ""
    If Terminal = "0" Then
        strqry = " SELECT A.FingerPrintID, (CASE WHEN A.FingerPrintID<>0 THEN 'Used' ELSE 'Not Used' END) As stt A.TerminalID " & _
                 " FROM kar_finger A INNER JOIN  machine B on B.TerminalID=A.TerminalID " & _
                        " LEFT JOIN cabang C on C.CAB_ID=B.CAB_ID "
    Else
        strqry = " SELECT A.FingerPrintID, (CASE WHEN A.FingerPrintID<>0 THEN 'Used' ELSE 'Not Used' END) As stt,A.TerminalID " & _
                 " FROM kar_finger A INNER JOIN  machine B on B.TerminalID=A.TerminalID " & _
                        " LEFT JOIN cabang C on C.CAB_ID=B.CAB_ID " & _
                 " WHERE C.CAB_ID='" & CabId & "'"
    End If
OpRecStt1 strqry, False
Cmb.Clear
Cmb.Style = xtpComboDropDownList
Cmb.EnableMarkup = True
i = 0
Cmb.AddItem 0
Cmb.MarkupList(Cmb.NewIndex) = "<StackPanel Orientation='Horizontal'>" & _
          "<Rectangle  Fill='#00FF00' Stroke='#1E0DF2' Width='14' Height='14'/>" & _
          "<TextBlock Padding = '2' VerticalAlignment='Center'> ALL </TextBlock></StackPanel>"  '#FC031C

With rsStt1
    If Not .EOF Then
        .MoveFirst
        Do Until .EOF
            Cmb.AddItem .Fields("FingerPrintID").Value
            'Cmb.MarkupList(Cmb.NewIndex) = "<Grid><Grid.ColumnDefinitions><ColumnDefinition Width='10'/><ColumnDefinition Width='100'/></Grid.ColumnDefinitions><Grid.RowDefinitions><RowDefinition/></Grid.RowDefinitions>" & _
                                  "<TextBlock Grid.Column='2'> " & .Fields("FingerPrintID").Value & "</TextBlock></Grid>"
            Cmb.MarkupList(Cmb.NewIndex) = "<StackPanel Orientation='Horizontal'>" & _
                      "<Rectangle  Fill='#00FF00' Stroke='#1E0DF2' Width='14' Height='14'/>" & _
                      "<TextBlock Padding = '2' VerticalAlignment='Center'>" & .Fields("FingerPrintID").Value & " | " & .Fields("Stt").Value & " | " & .Fields("TerminalID").Value & "</TextBlock></StackPanel>"  '#FC031C

       .MoveNext
        i = i + 1
        Loop
    End If
        If Cmb.ListCount <> 0 Then Cmb.ListIndex = 0
End With
Set rsStt1 = Nothing
End Sub


'===========================================
'============== ptr.nov Payroll ============
'=============== FUNCTION KEY ==============
'===========================================
Public Sub KeyInOut(Cmb As XtremeSuiteControls.ComboBox)
Dim strqry As String
strqry = ""
    strqry = "SELECT FunctionKey,FunctionKeyNM FROM key"
OpRecStt1 strqry, False
Cmb.Clear
Cmb.Style = xtpComboDropDownList
Cmb.EnableMarkup = True
With rsStt1
    If Not .EOF Then
        .MoveFirst
        
        Do Until .EOF
            Cmb.AddItem .Fields("FunctionKey").Value
            'Cmb.MarkupList(Cmb.NewIndex) = "<Grid><Grid.ColumnDefinitions><ColumnDefinition Width='10'/><ColumnDefinition Width='100'/></Grid.ColumnDefinitions><Grid.RowDefinitions><RowDefinition/></Grid.RowDefinitions>" & _
                                   "<TextBlock Grid.Column='2'>" & .Fields("FunctionKeyNM").Value & "</TextBlock></Grid>"
            'cmb.MarkupList(cmb.NewIndex) = "<Grid><Grid.ColumnDefinitions><ColumnDefinition Width='10'/><ColumnDefinition Width='100'/></Grid.ColumnDefinitions><Grid.RowDefinitions><RowDefinition/></Grid.RowDefinitions>" & _
                                  "<TextBlock>" & .Fields("DEP_ID").Value & "</TextBlock>" & _
                                  "<TextBlock Grid.Column='2'> | " & .Fields("DEP_NM").Value & "</TextBlock></Grid>"
            Cmb.MarkupList(Cmb.NewIndex) = "<StackPanel Orientation='Horizontal'>" & _
                      "<Rectangle  Fill='#FFFF00' Stroke='#1E0DF2' Width='14' Height='14'/>" & _
                      "<TextBlock Padding = '2' VerticalAlignment='Center'>" & .Fields("FunctionKeyNM").Value & "</TextBlock></StackPanel>"  '#FC031C
       
       .MoveNext
        Loop
    End If
        If Cmb.ListCount <> 0 Then Cmb.ListIndex = 0
End With
Set rsStt1 = Nothing
End Sub

'===========================================
'============== ptr.nov Payroll ============
'=============== GROUP TIME TABLE ==========
'===========================================
Public Sub TTGROUP(Cmb As XtremeSuiteControls.ComboBox)
Dim strqry As String
strqry = ""
    strqry = "SELECT TT_GRP_ID,TT_GRP_NM FROM timetable_grp"
OpRecStt1 strqry, False
Cmb.Clear
Cmb.Style = xtpComboDropDownList
Cmb.EnableMarkup = True
Cmb.AddItem 0
Cmb.MarkupList(Cmb.NewIndex) = "<StackPanel Orientation='Horizontal'>" & _
           "<Rectangle  Fill='#1E0DF2' Stroke='#1E0DF2' Width='14' Height='14'/>" & _
           "<TextBlock Padding = '2' VerticalAlignment='Center'>" & "ALL" & "</TextBlock></StackPanel>"

With rsStt1
    If Not .EOF Then
        .MoveFirst
        
        Do Until .EOF
            Cmb.AddItem .Fields("TT_GRP_ID").Value
            'Cmb.MarkupList(Cmb.NewIndex) = "<Grid><Grid.ColumnDefinitions><ColumnDefinition Width='10'/><ColumnDefinition Width='100'/></Grid.ColumnDefinitions><Grid.RowDefinitions><RowDefinition/></Grid.RowDefinitions>" & _
                                   "<TextBlock Grid.Column='2'>" & .Fields("TT_GRP_NM").Value & "</TextBlock></Grid>"
            'cmb.MarkupList(cmb.NewIndex) = "<Grid><Grid.ColumnDefinitions><ColumnDefinition Width='10'/><ColumnDefinition Width='100'/></Grid.ColumnDefinitions><Grid.RowDefinitions><RowDefinition/></Grid.RowDefinitions>" & _
                                  "<TextBlock>" & .Fields("DEP_ID").Value & "</TextBlock>" & _
                                  "<TextBlock Grid.Column='2'> | " & .Fields("DEP_NM").Value & "</TextBlock></Grid>"
            Cmb.MarkupList(Cmb.NewIndex) = "<StackPanel Orientation='Horizontal'>" & _
                      "<Rectangle  Fill='#800000' Stroke='#1E0DF2' Width='14' Height='14'/>" & _
                      "<TextBlock Padding = '2' VerticalAlignment='Center'>" & .Fields("TT_GRP_NM").Value & "</TextBlock></StackPanel>"  '#FC031C

       .MoveNext
        Loop
    End If
        If Cmb.ListCount <> 0 Then Cmb.ListIndex = 0
End With
Set rsStt1 = Nothing
End Sub


'===========================================
'============== ptr.nov Payroll ============
'================= CABANG ==================
'===========================================
Public Sub CABANG(Cmb As XtremeSuiteControls.ComboBox)
Dim strqry As String
strqry = ""
    strqry = "SELECT CAB_ID,CAB_NM FROM cabang " 'ORDER BY CAB_NM ASC"
OpRecStt2 strqry, False
Cmb.Clear
Cmb.Style = xtpComboDropDownList
Cmb.EnableMarkup = True
Cmb.AddItem 0
'CmbCab.MarkupList(CmbCab.NewIndex) = "<Grid><Grid.ColumnDefinitions><ColumnDefinition Width='10'/><ColumnDefinition Width='100'/></Grid.ColumnDefinitions><Grid.RowDefinitions><RowDefinition/></Grid.RowDefinitions>" & _
                       "<TextBlock Grid.Column='2'>" & "ALL" & "</TextBlock></Grid>"
Cmb.MarkupList(Cmb.NewIndex) = "<StackPanel Orientation='Horizontal'>" & _
           "<Rectangle  Fill='#80FFFF' Stroke='#1E0DF2' Width='14' Height='14'/>" & _
           "<TextBlock Padding = '2' VerticalAlignment='Center'>" & "ALL" & "</TextBlock></StackPanel>"

                      
With rsStt2
    If Not .EOF Then
        .MoveFirst
        
        Do Until .EOF
             Cmb.AddItem .Fields("CAB_ID").Value
           'CmbCab.MarkupList(CmbCab.NewIndex) = "<Grid><Grid.ColumnDefinitions><ColumnDefinition Width='10'/><ColumnDefinition Width='250'/></Grid.ColumnDefinitions><Grid.RowDefinitions><RowDefinition/></Grid.RowDefinitions>" & _
                                   "<TextBlock Grid.Column='2'>" & .Fields("CAB_NM").Value & "</TextBlock></Grid>"
            Cmb.MarkupList(Cmb.NewIndex) = "<StackPanel Orientation='Horizontal'>" & _
                      "<Rectangle  Fill='#80FFFF' Stroke='#1E0DF2' Width='14' Height='14'/>" & _
                      "<TextBlock Padding = '2' VerticalAlignment='Center'>" & .Fields("CAB_NM").Value & "</TextBlock></StackPanel>"
            'MsgBox .Fields("CAB_ID").Value
     .MoveNext
        Loop
    End If
        If Cmb.ListCount <> 0 Then Cmb.ListIndex = 0
End With
Set rsStt2 = Nothing
End Sub

'============== ptr.nov Payroll ============
'---------------- CABANG -------------------
'===========================================
Public Sub CABANG_NOALL(CmbCab As XtremeSuiteControls.ComboBox)
Dim strqry As String
strqry = ""
    strqry = "SELECT CAB_ID,CAB_NM FROM cabang ORDER BY CAB_NM ASC"
OpRecStt1 strqry, False
CmbCab.Clear
CmbCab.Style = xtpComboDropDownList
CmbCab.EnableMarkup = True
With rsStt1
    If Not .EOF Then
        .MoveFirst
        
        Do Until .EOF
             CmbCab.AddItem .Fields("CAB_ID").Value
           'CmbCab.MarkupList(CmbCab.NewIndex) = "<Grid><Grid.ColumnDefinitions><ColumnDefinition Width='10'/><ColumnDefinition Width='250'/></Grid.ColumnDefinitions><Grid.RowDefinitions><RowDefinition/></Grid.RowDefinitions>" & _
                                   "<TextBlock Grid.Column='2'>" & .Fields("CAB_NM").Value & "</TextBlock></Grid>"
            CmbCab.MarkupList(CmbCab.NewIndex) = "<StackPanel Orientation='Horizontal'>" & _
                      "<Rectangle  Fill='#80FFFF' Stroke='#1E0DF2' Width='14' Height='14'/>" & _
                      "<TextBlock Padding = '2' VerticalAlignment='Center'>" & .Fields("CAB_NM").Value & "</TextBlock></StackPanel>"

     .MoveNext
        Loop
    End If
        If CmbCab.ListCount <> 0 Then CmbCab.ListIndex = 0
End With
Set rsStt1 = Nothing
End Sub

'===========================================
'============== ptr.nov Payroll ============
'================= Jabatan =================
'===========================================
Public Sub JABATAN(CmbCab As XtremeSuiteControls.ComboBox)
Dim strqry As String
strqry = ""
    strqry = "SELECT JAB_ID,JAB_NM FROM jabatan ORDER BY JAB_NM ASC"
OpRecStt1 strqry, False
CmbCab.Clear
CmbCab.Style = xtpComboDropDownList
CmbCab.EnableMarkup = True
CmbCab.AddItem 0
CmbCab.MarkupList(CmbCab.NewIndex) = "<StackPanel Orientation='Horizontal'>" & _
           "<Rectangle  Fill='#80FFFF' Stroke='#1E0DF2' Width='14' Height='14'/>" & _
           "<TextBlock Padding = '2' VerticalAlignment='Center'>" & "ALL" & "</TextBlock></StackPanel>"
                      
With rsStt1
    If Not .EOF Then
        .MoveFirst
        
        Do Until .EOF
             CmbCab.AddItem .Fields("JAB_ID").Value
             CmbCab.MarkupList(CmbCab.NewIndex) = "<StackPanel Orientation='Horizontal'>" & _
                      "<Rectangle  Fill='#80FFFF' Stroke='#1E0DF2' Width='14' Height='14'/>" & _
                      "<TextBlock Padding = '2' VerticalAlignment='Center'>" & .Fields("JAB_NM").Value & "</TextBlock></StackPanel>"

     .MoveNext
        Loop
    End If
        If CmbCab.ListCount <> 0 Then CmbCab.ListIndex = 0
End With
Set rsStt1 = Nothing
End Sub

'===========================================
'-------------- ptr.nov Payroll ------------
'-------------- STATUS KARYAWAN ------------
'===========================================
Public Sub KAR_STT(CmbCab As XtremeSuiteControls.ComboBox)
Dim strqry As String
strqry = ""
    strqry = "SELECT KAR_STS_ID,KAR_STS_NM FROM kar_stt ORDER BY KAR_STS_NM ASC"
OpRecStt1 strqry, False
CmbCab.Clear
CmbCab.Style = xtpComboDropDownList
CmbCab.EnableMarkup = True
CmbCab.AddItem 0
'CmbCab.MarkupList(CmbCab.NewIndex) = "<Grid><Grid.ColumnDefinitions><ColumnDefinition Width='10'/><ColumnDefinition Width='100'/></Grid.ColumnDefinitions><Grid.RowDefinitions><RowDefinition/></Grid.RowDefinitions>" & _
                       "<TextBlock Grid.Column='2'>" & "ALL" & "</TextBlock></Grid>"
CmbCab.MarkupList(CmbCab.NewIndex) = "<StackPanel Orientation='Horizontal'>" & _
           "<Rectangle  Fill='#80FFFF' Stroke='#1E0DF2' Width='14' Height='14'/>" & _
           "<TextBlock Padding = '2' VerticalAlignment='Center'>" & "ALL" & "</TextBlock></StackPanel>"
                      
With rsStt1
    If Not .EOF Then
        .MoveFirst
        
        Do Until .EOF
             CmbCab.AddItem .Fields("KAR_STS_ID").Value
           'CmbCab.MarkupList(CmbCab.NewIndex) = "<Grid><Grid.ColumnDefinitions><ColumnDefinition Width='10'/><ColumnDefinition Width='250'/></Grid.ColumnDefinitions><Grid.RowDefinitions><RowDefinition/></Grid.RowDefinitions>" & _
                                   "<TextBlock Grid.Column='2'>" & .Fields("CAB_NM").Value & "</TextBlock></Grid>"
            CmbCab.MarkupList(CmbCab.NewIndex) = "<StackPanel Orientation='Horizontal'>" & _
                      "<Rectangle  Fill='#80FFFF' Stroke='#1E0DF2' Width='14' Height='14'/>" & _
                      "<TextBlock Padding = '2' VerticalAlignment='Center'>" & .Fields("KAR_STS_NM").Value & "</TextBlock></StackPanel>"

     .MoveNext
        Loop
    End If
        If CmbCab.ListCount <> 0 Then CmbCab.ListIndex = 0
End With
Set rsStt1 = Nothing
End Sub

'===========================================
'============== ptr.nov Payroll ============
'================= CORP ==================
'===========================================
Public Sub CORP(CmbCab As XtremeSuiteControls.ComboBox)
Dim strqry As String
strqry = ""
    strqry = "SELECT CORP_ID,CORP_NM FROM corp ORDER BY CORP_NM ASC"
OpRecStt1 strqry, False
CmbCab.Clear
CmbCab.Style = xtpComboDropDownList
CmbCab.EnableMarkup = True
                     
With rsStt1
    If Not .EOF Then
        .MoveFirst
        
        Do Until .EOF
            CmbCab.AddItem .Fields("CORP_ID").Value
            CmbCab.MarkupList(CmbCab.NewIndex) = "<StackPanel Orientation='Horizontal'>" & _
            "<Rectangle  Fill='#80FFFF' Stroke='#1E0DF2' Width='14' Height='14'/>" & _
            "<TextBlock Padding = '2' VerticalAlignment='Center'>" & .Fields("CORP_NM").Value & "</TextBlock></StackPanel>"
     .MoveNext
        Loop
    End If
        If CmbCab.ListCount <> 0 Then CmbCab.ListIndex = 0
End With
Set rsStt1 = Nothing
End Sub

'========================================
'---------- ptr.nov Payroll -------------
'------- PRESENSI REPORTING LIST --------
'========================================
Public Sub PRESENSI_RPT_LIST(Cmb As XtremeSuiteControls.ComboBox)
Dim strqry As String
strqry = ""
    strqry = "SELECT RPT_NM FROM presensi_list_rpt"
OpRecStt1 strqry, False
Cmb.Clear
Cmb.Style = xtpComboDropDownList
Cmb.EnableMarkup = True
                     '
With rsStt1
    If Not .EOF Then
        .MoveFirst
        
        Do Until .EOF
            Cmb.AddItem .Fields("RPT_NM").Value
            'Cmb.MarkupList(Cmb.NewIndex) = "<Grid><Grid.ColumnDefinitions><ColumnDefinition Width='10'/><ColumnDefinition Width='200'/></Grid.ColumnDefinitions><Grid.RowDefinitions><RowDefinition/></Grid.RowDefinitions>" & _
                                  "<TextBlock>" & .Fields("TerminalID").Value & "</TextBlock>" & _
                                  "<TextBlock Grid.Column='2'> | " & .Fields("MESIN_NM").Value & " | " & .Fields("MESIN_SN").Value & "</TextBlock></Grid>"
            Cmb.MarkupList(Cmb.NewIndex) = "<StackPanel Orientation='Horizontal'>" & _
                                  "<Rectangle  Fill='#00FF00' Stroke='#1E0DF2' Width='14' Height='14'/>" & _
                                  "<TextBlock Padding = '2' VerticalAlignment='Center'>" & .Fields("RPT_NM").Value & "</TextBlock></StackPanel>"  '#FC031C

          
       .MoveNext
        Loop
    End If
        If Cmb.ListCount <> 0 Then Cmb.ListIndex = 0
End With
Set rsStt1 = Nothing
End Sub

'========================================
'---------- ptr.nov Payroll -------------
'------- PRESENSI REPORTING LIST --------
'========================================
Public Sub PAYROLL_RPT_LIST(Cmb As XtremeSuiteControls.ComboBox)
Dim strqry As String
strqry = ""
    strqry = "SELECT RPT_NM FROM payroll_list_rpt"
OpRecStt1 strqry, False
Cmb.Clear
Cmb.Style = xtpComboDropDownList
Cmb.EnableMarkup = True
                     '
With rsStt1
    If Not .EOF Then
        .MoveFirst
        
        Do Until .EOF
            Cmb.AddItem .Fields("RPT_NM").Value
            'Cmb.MarkupList(Cmb.NewIndex) = "<Grid><Grid.ColumnDefinitions><ColumnDefinition Width='10'/><ColumnDefinition Width='200'/></Grid.ColumnDefinitions><Grid.RowDefinitions><RowDefinition/></Grid.RowDefinitions>" & _
                                  "<TextBlock>" & .Fields("TerminalID").Value & "</TextBlock>" & _
                                  "<TextBlock Grid.Column='2'> | " & .Fields("MESIN_NM").Value & " | " & .Fields("MESIN_SN").Value & "</TextBlock></Grid>"
            Cmb.MarkupList(Cmb.NewIndex) = "<StackPanel Orientation='Horizontal'>" & _
                                  "<Rectangle  Fill='#00FF00' Stroke='#1E0DF2' Width='14' Height='14'/>" & _
                                  "<TextBlock Padding = '2' VerticalAlignment='Center'>" & .Fields("RPT_NM").Value & "</TextBlock></StackPanel>"  '#FC031C

          
       .MoveNext
        Loop
    End If
        If Cmb.ListCount <> 0 Then Cmb.ListIndex = 0
End With
Set rsStt1 = Nothing
End Sub

'===========================================
'============== ptr.nov Payroll ============
'==============MENU PAYROLL    =============
'===========================================
Public Sub MENU_PAYROLL(Cmb As XtremeSuiteControls.ComboBox)
Dim strqry As String
strqry = ""
    strqry = "SELECT PAYMENU_ID,PAYMENU_NM FROM payroll_menu ORDER BY PAYMENU_NM"
OpRecStt1 strqry, False
Cmb.Clear
Cmb.Style = xtpComboDropDownList
Cmb.EnableMarkup = True
                     '
With rsStt1
    If Not .EOF Then
        .MoveFirst
        
        Do Until .EOF
            Cmb.AddItem .Fields("PAYMENU_ID").Value
           Cmb.MarkupList(Cmb.NewIndex) = "<StackPanel Orientation='Horizontal'>" & _
                                  "<Rectangle  Fill='#00FF00' Stroke='#1E0DF2' Width='14' Height='14'/>" & _
                                  "<TextBlock Padding = '2' VerticalAlignment='Center'>" & .Fields("PAYMENU_NM").Value & "</TextBlock></StackPanel>"  '#FC031C

          
       .MoveNext
        Loop
    End If
        If Cmb.ListCount <> 0 Then Cmb.ListIndex = 0
End With
Set rsStt1 = Nothing
End Sub

'fungsi untuk membaca dev dxgrid field valuse
Public Function GetGridDefValue(xGrid As dxDBGrid, xField As String) As String
On Error Resume Next
GetGridDefValue = xGrid.Items(xGrid.Ex.FocusedNode.AbsoluteIndex).Values(xGrid.Columns.ColumnByFieldName(xField).Index)
End Function

Public Function GetNama_Departemen(xDepId As String) As String
On Error Resume Next
 GetNama_Departemen = GetDtaTbl(conMain, "SELECT DEP_NM FROM departemen WHERE DEP_ID='" & xDepId & "'", "DEP_NM")
End Function
Public Function GetID_Departemen(xDepNM As String) As String
On Error Resume Next
 GetID_Departemen = GetDtaTbl(conMain, "SELECT DEP_ID FROM departemen WHERE DEP_NM='" & xDepNM & "'", "DEP_ID")
End Function
Public Function GetNama_Jabatan(xId As String) As String
On Error Resume Next
 GetNama_Jabatan = GetDtaTbl(conMain, "SELECT JAB_NM FROM jabatan WHERE JAB_ID='" & xId & "'", "JAB_NM")
End Function
Public Function GetID_Jabatan(xNM As String) As String
On Error Resume Next
 GetID_Jabatan = GetDtaTbl(conMain, "SELECT JAB_ID FROM jabatan WHERE JAB_NM='" & xNM & "'", "JAB_ID")
End Function

Public Function GetUserpPerm(xUsrId As String, xMN_ID As String) As Boolean
Dim defSts As Boolean
    If GetDtaTbl(conMain, "SELECT MN_SHW FROM tab_user_per WHERE PRMS_ID='" & xUsrId & "' AND MN_ID='" & xMN_ID & "'", "MN_SHW") = "1" Then defSts = True Else defSts = False
    GetUserpPerm = defSts
End Function
Public Function GetID_User_From_IDkaryawan(KaryawanID As String) As String
On Error Resume Next
Dim oRS As ADODB.Recordset
Dim sqlQry As String

 GetID_User_From_IDkaryawan = GetDtaTbl(conMain, "SELECT USER_ID FROM tab_user WHERE KAR_ID='" & KaryawanID & "'", "USER_ID")
End Function
Public Function GetID_karyawan_From_IDUser(userID As String) As String
On Error Resume Next
Dim sqlQry As String

 GetID_karyawan_From_IDUser = GetDtaTbl(conMain, "SELECT KAR_ID FROM tab_user WHERE USER_ID='" & userID & "'", "KAR_ID")
End Function

Public Function GetName_Karyawan(KaryawanID As String) As String
On Error Resume Next
Dim oRS As ADODB.Recordset
Dim sqlQry As String

 GetName_Karyawan = GetDtaTbl(conMain, "SELECT KAR_NM FROM karyawan WHERE KAR_ID='" & KaryawanID & "'", "KAR_NM")
End Function
Public Function GetID_KaryawanJabatan(KaryawanID As String) As String
On Error Resume Next
Dim oRS As ADODB.Recordset
Dim sqlQry As String

 GetID_KaryawanJabatan = GetDtaTbl(conMain, "SELECT JAB_ID FROM karyawan WHERE KAR_ID='" & KaryawanID & "'", "JAB_ID")
End Function

Public Function PaidSaid(Number As String) As String
'======================================================
'|              LG PROJECT               [23 May 2014]|
'======================================================
'|ClsReportExcel - PaidSaid                           |
'|(digunakan mengconvert nominal uang/int ke string   |
'|Progreass = 80%                                     |
'------------------------------------------------------
'|By. FARINDRA                            @01-07-2014 |
'======================================================
Dim NumStr(10) As String
Dim SetVal(8), Dp(5), Satuan(10), Belasan(10), Puluhan(10), Ratusan(10), Ribuan(10), Jutaan(10), Said(5) As String
Dim Sat, Pul, Rat, Rib, Jut As Integer
On Error Resume Next
 
 s = "Se"
 Dp(1) = "Belas"
 Dp(2) = "Puluh"
 Dp(3) = "Ratus"
 Dp(4) = "Ribu"
 Dp(5) = "Juta"
 Satuan(0) = ""
 Satuan(1) = "Satu"
 Satuan(2) = "Dua"
 Satuan(3) = "Tiga"
 Satuan(4) = "Empat"
 Satuan(5) = "Lima"
 Satuan(6) = "Enam"
 Satuan(7) = "Tujuh"
 Satuan(8) = "Delapan"
 Satuan(9) = "Sembilan"
 
 Belasan(1) = s & LCase(Dp(1))
 Puluhan(1) = s & LCase(Dp(2))
 Ratusan(1) = s & LCase(Dp(3))
 Ribuan(1) = s & LCase(Dp(4))
 Jutaan(1) = Satuan(1) & " " & Dp(5)
 
 NumStr(0) = LTrim(Str(Number))

 For J = 2 To 9
    Belasan(J) = Satuan(J) & " " & Dp(1)
    Puluhan(J) = Satuan(J) & " " & Dp(2)
    Ratusan(J) = Satuan(J) & " " & Dp(3)
    Ribuan(J) = Satuan(J) & " " & Dp(4)
    Jutaan(J) = Satuan(J) & " " & Dp(5)
 Next J
   
 For i = 1 To Len(NumStr(0))
    NumStr(i) = Mid(LTrim(NumStr(0)), i, 1)
 Next i
 
 'jumlah digit number 1 (satuan)
 If Len(NumStr(0)) >= 1 Then
    Sat = Int(Len(NumStr(0)))
    PaidSaid = Trim(Satuan(NumStr(Sat)))
 End If
 
 'jumlah digit number 2 (belasan dan puluhan)
 If Len(NumStr(0)) >= 2 Then
    Pul = Int(Len(NumStr(0)) - 1)
    If NumStr(Pul) = "1" Then
        PaidSaid = Belasan(Int(NumStr(Sat)))
        If NumStr(Sat) = "0" Then PaidSaid = Puluhan(NumStr(Pul))
    End If
    
    If NumStr(Pul) <> "1" Then PaidSaid = Puluhan(NumStr(Pul)) & " " & PaidSaid
 
    PaidSaid = Trim(PaidSaid)
 End If
      

 'jumlah digit number 3 (ratusan)
 If Len(NumStr(0)) >= 3 Then
     Rat = Int(Len(NumStr(0)) - 2)
     PaidSaid = Ratusan(NumStr(Rat)) & " " & PaidSaid
     PaidSaid = Trim(PaidSaid)
 End If
 
 'jumlah digit number 4 (ribuan)
 If Len(NumStr(0)) = 4 Then
     Rib = Int(Len(NumStr(0)) - 3)
     PaidSaid = Ribuan(NumStr(Rib)) & " " & PaidSaid
     PaidSaid = Trim(PaidSaid)
 End If
 
 'jumlah digit number 5 (belasan & puluan ribu)
 If Len(NumStr(0)) >= 5 Then
    Sat = Int(Len(NumStr(0))) - 3
    
    Pul = Int(Len(NumStr(0))) - 4
    If NumStr(Pul) = "1" Then
        PaidSaid = Belasan(Int(NumStr(Sat))) & " " & Dp(4) & " " & Replace(PaidSaid, Dp(4), "")
        If NumStr(Sat) = "0" Then PaidSaid = Puluhan(NumStr(Pul)) & " " & Dp(4) & " " & Replace(PaidSaid, Dp(4), "")
    End If
    
    If NumStr(Pul) <> "1" Then
        PaidSaid = Satuan(NumStr(Sat)) & " " & Dp(4) & " " & Replace(PaidSaid, Dp(4), "")
        PaidSaid = Puluhan(NumStr(Pul)) & " " & PaidSaid
     End If
     PaidSaid = Trim(PaidSaid)
 End If

 'jumlah digit number 6 (ratusan ribu)
 If Len(NumStr(0)) >= 6 Then
 Rat = Int(Len(NumStr(0)) - 5)
     PaidSaid = Ratusan(NumStr(Rat)) & " " & PaidSaid
     PaidSaid = Trim(PaidSaid)
     SetVal(1) = PaidSaid
     SetVal(1) = Trim(SetVal(1))
 End If
 
 'jumlah digit number 7 (jutaan)
 If Len(NumStr(0)) >= 7 Then
    Jut = Int(Len(NumStr(0)) - 6)
    If PaidSaid = Dp(4) Then PaidSaid = Jutaan(NumStr(Jut)) Else PaidSaid = Jutaan(NumStr(Jut)) & " " & PaidSaid
    Trim (PaidSaid)
 End If

'jumlah digit number 8 (belasan & puluhan jutan)
 If Len(NumStr(0)) >= 8 Then
    Sat = Int(Len(NumStr(0))) - 6
    Pul = Int(Len(NumStr(0))) - 7
    If NumStr(Pul) = "1" Then
        If SetVal(1) = Dp(4) Then
         PaidSaid = Belasan(Int(NumStr(Sat))) & " " & Dp(5) & " " & Replace(SetVal(1), Dp(4), "")
         If NumStr(Sat) = "0" Then PaidSaid = Puluhan(NumStr(Pul)) & " " & Dp(5) & " " & Replace(SetVal(1), Dp(4), "")
        Else
         PaidSaid = Belasan(Int(NumStr(Sat))) & " " & Dp(5) & " " & SetVal(1)
         If NumStr(Sat) = "0" Then PaidSaid = Puluhan(NumStr(Pul)) & " " & Dp(5) & " " & SetVal(1)
        End If
    End If
    Trim (PaidSaid)
    
    
        If NumStr(Pul) <> "1" And PaidSaid <> "" Then PaidSaid = Puluhan(NumStr(Pul)) & " " & PaidSaid

    If NumStr(Pul) <> "1" And PaidSaid = "" Then PaidSaid = Puluhan(NumStr(Pul)) & " " & Dp(5)
    
    Trim (PaidSaid)
   ' Jut = Int(Len(NumStr(0)) - 6)
End If

'jumlah digit number 9 (ratusan juta)
 If Len(NumStr(0)) >= 9 Then
 Rat = Int(Len(NumStr(0)) - 8)
 'If PaidSaid = Dp(5) Then
 PaidSaid = Ratusan(NumStr(Rat)) & " " & PaidSaid
     PaidSaid = Trim(PaidSaid)
     SetVal(2) = PaidSaid
     SetVal(2) = Trim(SetVal(2))
 End If
 
End Function
Public Function GetFingerID_Karyawan(KaryawanName As String) As String
On Error Resume Next
Dim oRS As ADODB.Recordset
Dim sqlQry As String

 GetFingerID_Karyawan = GetDtaTbl(conMain, "SELECT FingerPrintId FROM karyawan WHERE KAR_NM='" & KaryawanName & "'", "FingerPrintId")
End Function


Public Sub GetListComboByQRY(xCombo As XtremeSuiteControls.ComboBox, xStrQry, Optional isProcedure As Boolean)
On Error GoTo ErrorLabel

Dim oRS As ADODB.Recordset
  Set oRS = New ADODB.Recordset
  
  xCombo.Clear
  
    If isProcedure Then oRS.Open xStrQry, conMain, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc Else oRS.Open xStrQry, conMain, adOpenForwardOnly, adLockReadOnly, adCmdText
    
      Do While Not oRS.EOF
            xCombo.AddItem oRS.Fields.Item(0) '.Fields("BRG_ID").Value
        oRS.MoveNext
      Loop
  
  oRS.Close
  Set oRS = Nothing
  
xCombo.ListIndex = 0

Exit Sub
ErrorLabel:
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"

End Sub
