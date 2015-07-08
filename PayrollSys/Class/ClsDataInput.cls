VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsDataInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function Barang_Satuan(oConn As ADODB.Connection, defId As String, defCorp As String, defCab As String)
Dim defVal As String
LabSatuan:
     defVal = InputBox("Masukan Nama Satuan Baru untuk ID '" & defId & "'", "Barang - Satuan Baru", "")
    If defVal = "" Then
        'MsgBox "Data tidak boleh kosong...!!!"
    Else
        If Val(GetDtaTbl(oConn, "SELECT COUNT(SAT_NM) AS DEFMAX FROM tab_barang_satuan WHERE SAT_NM='" & defVal & "' AND CORP_ID='" & defCorp & "' AND CAB_ID='" & defCab & "'", "DEFMAX")) < 1 Then
            PrjSysTrig.DataIns conMain, "tab_barang_satuan", "'" & defId & "','" & defVal & "','','" & defCorp & "','" & defCab & "'"
        Else
            MsgBox "Nama Satuan '" & defVal & "' Sudah ADA !!!", vbCritical
            GoTo LabSatuan
        End If
    End If
End Function

Public Function Barang_Type(oConn As ADODB.Connection, defId As String, defCorp As String, defCab As String)
Dim defVal As String
LabType:
     defVal = InputBox("Masukan Nama Type Baru untuk ID '" & defId & "'", "Barang - Type Baru", "")
    If defVal = "" Then
        'MsgBox "Data tidak boleh kosong...!!!"
    Else
        If Val(GetDtaTbl(conMain, "SELECT COUNT(TYP_NM) AS DEFMAX FROM tab_barang_type WHERE TYP_NM='" & defVal & "' AND CORP_ID='" & defCorp & "' AND CAB_ID='" & defCab & "'", "DEFMAX")) < 1 Then
            PrjSysTrig.DataIns conMain, "tab_barang_type", "'" & defId & "','" & defVal & "','','" & defCorp & "','" & defCab & "'"
        Else
            MsgBox "Nama Type '" & defVal & "' Sudah ADA !!!", vbCritical
            GoTo LabType
        End If
    End If
End Function
Public Function Barang_Kategori(oConn As ADODB.Connection, defId As String, defCorp As String, defCab As String)
Dim defVal As String
LabKategori:
     defVal = InputBox("Masukan Nama Kategori Baru untuk ID '" & defId & "'", "Barang - Kategori Baru", "")
    If defVal = "" Then
        'MsgBox "Data tidak boleh kosong...!!!"
    Else
        If Val(GetDtaTbl(conMain, "SELECT COUNT(CAT_NM) AS DEFMAX FROM tab_barang_cat WHERE CAT_NM='" & defVal & "' AND CORP_ID='" & defCorp & "' AND CAB_ID='" & defCab & "'", "DEFMAX")) < 1 Then
            PrjSysTrig.DataIns conMain, "tab_barang_cat", "'" & defId & "','" & defVal & "','','" & defCorp & "','" & defCab & "'"
        Else
            MsgBox "Nama Kategori '" & defVal & "' Sudah ADA !!!", vbCritical
            GoTo LabKategori
        End If
    End If
End Function
Public Function NewAgama(oConn As ADODB.Connection)
Dim defVal As String
Dim defId As String
defId = Format(Now, "YYYYMMDD-hms")

Ulangi:
     defVal = InputBox("Masukkan Nama Agama Baru", "Input Agama", "")
    If defVal = "" Then
        'MsgBox "Data tidak boleh kosong...!!!"
    Else
        If Val(GetDtaTbl(oConn, "SELECT COUNT(AGAMA_NM) AS DEFMAX FROM tab_agama WHERE AGAMA_NM='" & defVal & "'", "DEFMAX")) < 1 Then
            PrjSysTrig.DataIns conMain, "tab_agama", "'" & defId & "','" & defVal & "','" & UsrID & "','" & Format(Now, "YYYY-MM-DD h:m:s") & "'"
        Else
            MsgBox "Nama Agama '" & defVal & "' Sudah ADA !!!", vbCritical
            GoTo Ulangi
        End If
    End If
End Function

Public Function EditAgama(oConn As ADODB.Connection, xdefVal As String)
Dim defVal As String
Dim defId As String
defId = Format(Now, "YYYYMMDD-hms")

Ulangi:
     defVal = InputBox("Edit Nama Agama", "Edit Agama", xdefVal)
    If defVal = "" Then
        'MsgBox "Data tidak boleh kosong...!!!"
        'GoTo Ulangi
    Else
       If defVal = xdefVal Then Exit Function
        If Val(GetDtaTbl(oConn, "SELECT COUNT(AGAMA_NM) AS DEFMAX FROM tab_agama WHERE AGAMA_NM='" & defVal & "'", "DEFMAX")) < 1 Then
            PrjSysTrig.DataUpdate oConn, "tab_agama", "AGAMA_NM='" & defVal & "',USER='" & UsrID & "',UPDT='" & Format(Now, "YYYY-MM-DD h:m:s") & "'", "WHERE AGAMA_NM='" & xdefVal & "'"
        Else
            MsgBox "Nama Agama '" & defVal & "' Sudah ADA !!!", vbCritical
            GoTo Ulangi
        End If
    End If
End Function
Public Function NewJamKerja(oConn As ADODB.Connection)
Dim defVal As String
Dim defId As String
'defId = Format(Now, "YYYYMMDD-hms")

Ulangi:
     defVal = InputBox("Masukkan Nama Jam Kerja Baru", "Input Jam Kerja", "")
    If defVal = "" Then
        'MsgBox "Data tidak boleh kosong...!!!"
    Else
        If Val(GetDtaTbl(oConn, "SELECT COUNT(TT_GRP_NM) AS DEFMAX FROM tab_timetabel_grp WHERE TT_GRP_NM='" & defVal & "'", "DEFMAX")) < 1 Then
            defId = Val(GetDtaTbl(conMain, "SELECT MAX(TT_GRP_ID) AS DEFMAX FROM tab_timetabel_grp", "DEFMAX")) + 1
            PrjSysTrig.DataIns conMain, "tab_timetabel_grp", "'" & defId & "','" & defVal & "','1'"
        Else
            MsgBox "Nama Jam Kerja '" & defVal & "' Sudah ADA !!!", vbCritical
            GoTo Ulangi
        End If
    End If
End Function

Public Function EditJamKerja(oConn As ADODB.Connection, xdefVal As String)
Dim defVal As String
Dim defId As String
defId = Format(Now, "YYYYMMDD-hms")

Ulangi:
     defVal = InputBox("Edit Nama Jam Kerja", "Edit Jam Kerja", xdefVal)
    If defVal = "" Then
        'MsgBox "Data tidak boleh kosong...!!!"
        'GoTo Ulangi
    Else
       If defVal = xdefVal Then Exit Function
        If Val(GetDtaTbl(oConn, "SELECT COUNT(TT_GRP_NM) AS DEFMAX FROM tab_timetabel_grp WHERE TT_GRP_NM='" & defVal & "'", "DEFMAX")) < 1 Then
            PrjSysTrig.DataUpdate oConn, "tab_timetabel_grp", "TT_GRP_NM='" & defVal & "'", "WHERE TT_GRP_NM='" & xdefVal & "'"
        Else
            MsgBox "Nama Jam Kerja '" & defVal & "' Sudah ADA !!!", vbCritical
            GoTo Ulangi
        End If
    End If
End Function
Public Function NewIjin(oConn As ADODB.Connection)
Dim defVal As String
Dim defId As String
'defId = Format(Now, "YYYYMMDD-hms")

Ulangi:
     defVal = InputBox("Masukkan Nama Ijin Baru", "Input Ijin", "")
    If defVal = "" Then
        'MsgBox "Data tidak boleh kosong...!!!"
    Else
        If Val(GetDtaTbl(oConn, "SELECT COUNT(IJIN_NM) AS DEFMAX FROM tab_timetabel_ijin WHERE IJIN_NM='" & defVal & "'", "DEFMAX")) < 1 Then
            defId = Val(GetDtaTbl(conMain, "SELECT MAX(IJIN_NM) AS DEFMAX FROM tab_timetabel_ijin", "DEFMAX")) + 1
            PrjSysTrig.DataIns conMain, "tab_timetabel_ijin", "'" & defId & "','" & defVal & "','1'"
        Else
            MsgBox "Nama Ijin '" & defVal & "' Sudah ADA !!!", vbCritical
            GoTo Ulangi
        End If
    End If
End Function

Public Function EditIjin(oConn As ADODB.Connection, xdefVal As String)
Dim defVal As String
Dim defId As String
defId = Format(Now, "YYYYMMDD-hms")

Ulangi:
     defVal = InputBox("Edit Nama Ijin", "Edit Ijin", xdefVal)
    If defVal = "" Then
        'MsgBox "Data tidak boleh kosong...!!!"
        'GoTo Ulangi
    Else
       If defVal = xdefVal Then Exit Function
        If Val(GetDtaTbl(oConn, "SELECT COUNT(IJIN_NM) AS DEFMAX FROM tab_timetabel_ijin WHERE IJIN_NM='" & defVal & "'", "DEFMAX")) < 1 Then
            PrjSysTrig.DataUpdate oConn, "tab_timetabel_ijin", "IJIN_NM='" & defVal & "'", "WHERE IJIN_NM='" & xdefVal & "'"
        Else
            MsgBox "Nama Ijin '" & defVal & "' Sudah ADA !!!", vbCritical
            GoTo Ulangi
        End If
    End If
End Function
Public Function NewBagian(oConn As ADODB.Connection)
Dim defVal As String
Dim defId As String
defId = Format(Now, "YYYYMMDD-hms")

Ulangi:
     defVal = InputBox("Masukkan Nama Bagian Baru", "Input Bagian", "")
    If defVal = "" Then
        'MsgBox "Data tidak boleh kosong...!!!"
    Else
        If Val(GetDtaTbl(oConn, "SELECT COUNT(BAGIAN_NM) AS DEFMAX FROM tab_bagian WHERE BAGIAN_NM='" & defVal & "'", "DEFMAX")) < 1 Then
            PrjSysTrig.DataIns conMain, "tab_bagian", "'" & defId & "','" & defVal & "','" & UsrID & "','" & Format(Now, "YYYY-MM-DD h:m:s") & "'"
        Else
            MsgBox "Nama Bagian '" & defVal & "' Sudah ADA !!!", vbCritical
            GoTo Ulangi
        End If
    End If
End Function

Public Function EditBagian(oConn As ADODB.Connection, xdefVal As String)
Dim defVal As String
Dim defId As String
defId = Format(Now, "YYYYMMDD-hms")

Ulangi:
     defVal = InputBox("Edit Nama Bagian", "Edit Bagian", xdefVal)
    If defVal = "" Then
        'MsgBox "Data tidak boleh kosong...!!!"
        'GoTo Ulangi
    Else
       If defVal = xdefVal Then Exit Function
        If Val(GetDtaTbl(oConn, "SELECT COUNT(BAGIAN_NM) AS DEFMAX FROM tab_bagian WHERE BAGIAN_NM='" & defVal & "'", "DEFMAX")) < 1 Then
            PrjSysTrig.DataUpdate oConn, "tab_bagian", "BAGIAN_NM='" & defVal & "',USER='" & UsrID & "',UPDT='" & Format(Now, "YYYY-MM-DD h:m:s") & "'", "WHERE BAGIAN_NM='" & xdefVal & "'"
        Else
            MsgBox "Nama Bagian '" & defVal & "' Sudah ADA !!!", vbCritical
            GoTo Ulangi
        End If
    End If
End Function
Public Function NewPendidikan(oConn As ADODB.Connection)
Dim defVal As String
Dim defId As String
defId = Format(Now, "YYYYMMDD-hms")

Ulangi:
     defVal = InputBox("Masukkan Nama Pendidikan Baru", "Input Pendidikan ", "")
    If defVal = "" Then
        'MsgBox "Data tidak boleh kosong...!!!"
    Else
        If Val(GetDtaTbl(oConn, "SELECT COUNT(PEN_NM) AS DEFMAX FROM tab_pendidikan WHERE PEN_NM='" & defVal & "'", "DEFMAX")) < 1 Then
            PrjSysTrig.DataIns conMain, "tab_pendidikan", "'" & defId & "','" & defVal & "','" & UsrID & "','" & Format(Now, "YYYY-MM-DD h:m:s") & "'"
        Else
            MsgBox "Nama Pendidikan  '" & defVal & "' Sudah ADA !!!", vbCritical
            GoTo Ulangi
        End If
    End If
End Function

Public Function EditPendidikan(oConn As ADODB.Connection, xdefVal As String)
Dim defVal As String
Dim defId As String
defId = Format(Now, "YYYYMMDD-hms")

Ulangi:
     defVal = InputBox("Edit Nama Pendidikan ", "Edit Pendidikan ", xdefVal)
    If defVal = "" Then
        'MsgBox "Data tidak boleh kosong...!!!"
        'GoTo Ulangi
    Else
       If defVal = xdefVal Then Exit Function
        If Val(GetDtaTbl(oConn, "SELECT COUNT(PEN_NM) AS DEFMAX FROM tab_pendidikan WHERE PEN_NM='" & defVal & "'", "DEFMAX")) < 1 Then
            PrjSysTrig.DataUpdate oConn, "tab_pendidikan", "PEN_NM='" & defVal & "',USER='" & UsrID & "',UPDT='" & Format(Now, "YYYY-MM-DD h:m:s") & "'", "WHERE PEN_NM='" & xdefVal & "'"
        Else
            MsgBox "Nama Pendidikan  '" & defVal & "' Sudah ADA !!!", vbCritical
            GoTo Ulangi
        End If
    End If
End Function
Public Function NewDepartemen(oConn As ADODB.Connection)
Dim defVal As String
Dim defId As String
defId = GetDtaTbl(oConn, "SELECT MAX(DEP_ID) AS DEFMAX FROM departemen", "DEFMAX") + 1

Ulangi:
     defVal = InputBox("Masukkan Nama Departemen Baru", "Input Departemen", "")
    If defVal = "" Then
        'MsgBox "Data tidak boleh kosong...!!!"
    Else
        If Val(GetDtaTbl(oConn, "SELECT COUNT(DEP_NM) AS DEFMAX FROM departemen WHERE DEP_NM='" & defVal & "'", "DEFMAX")) < 1 Then
            PrjSysTrig.DataIns conMain, "departemen", "'" & defId & "','" & defVal & "','" & UsrID & "','" & Format(Now, "YYYY-MM-DD h:m:s") & "'"
        Else
            MsgBox "Nama Departemen '" & defVal & "' Sudah ADA !!!", vbCritical
            GoTo Ulangi
        End If
    End If
End Function

Public Function EditDepartemen(oConn As ADODB.Connection, xdefVal As String)
Dim defVal As String
Dim defId As String

Ulangi:
     defVal = InputBox("Edit Nama Departemen", "Edit Departemen", xdefVal)
    If defVal = "" Then
        'MsgBox "Data tidak boleh kosong...!!!"
        'GoTo Ulangi
    Else
       If defVal = xdefVal Then Exit Function
        If Val(GetDtaTbl(oConn, "SELECT COUNT(DEP_NM) AS DEFMAX FROM departemen WHERE DEP_NM='" & defVal & "'", "DEFMAX")) < 1 Then
            PrjSysTrig.DataUpdate oConn, "departemen", "DEP_NM='" & defVal & "',USER='" & UsrID & "',UPDT='" & Format(Now, "YYYY-MM-DD h:m:s") & "'", "WHERE DEP_NM='" & xdefVal & "'"
        Else
            MsgBox "Nama Departemen '" & defVal & "' Sudah ADA !!!", vbCritical
            GoTo Ulangi
        End If
    End If
End Function
Public Function NewJabatan(oConn As ADODB.Connection)
Dim defVal As String
Dim defId As String
defId = GetDtaTbl(oConn, "SELECT MAX(JAB_ID) AS DEFMAX FROM jabatan", "DEFMAX") + 1

Ulangi:
     defVal = InputBox("Masukkan Nama Jabatan Baru", "Input Jabatan", "")
    If defVal = "" Then
        'MsgBox "Data tidak boleh kosong...!!!"
    Else
        If Val(GetDtaTbl(oConn, "SELECT COUNT(JAB_NM) AS DEFMAX FROM jabatan WHERE JAB_NM='" & defVal & "'", "DEFMAX")) < 1 Then
            PrjSysTrig.DataIns conMain, "jabatan", "'" & defId & "','" & defVal & "','" & UsrID & "','" & Format(Now, "YYYY-MM-DD h:m:s") & "'"
        Else
            MsgBox "Nama Jabatan '" & defVal & "' Sudah ADA !!!", vbCritical
            GoTo Ulangi
        End If
    End If
End Function

Public Function EditJabatan(oConn As ADODB.Connection, xdefVal As String)
Dim defVal As String
Dim defId As String

Ulangi:
     defVal = InputBox("Edit Nama Jabatan", "Edit Jabatan", xdefVal)
    If defVal = "" Then
        'MsgBox "Data tidak boleh kosong...!!!"
        'GoTo Ulangi
    Else
       If defVal = xdefVal Then Exit Function
        If Val(GetDtaTbl(oConn, "SELECT COUNT(JAB_NM) AS DEFMAX FROM jabatan WHERE JAB_NM='" & defVal & "'", "DEFMAX")) < 1 Then
            PrjSysTrig.DataUpdate oConn, "jabatan", "JAB_NM='" & defVal & "',USER='" & UsrID & "',UPDT='" & Format(Now, "YYYY-MM-DD h:m:s") & "'", "WHERE JAB_NM='" & xdefVal & "'"
        Else
            MsgBox "Nama Jabatan '" & defVal & "' Sudah ADA !!!", vbCritical
            GoTo Ulangi
        End If
    End If
End Function
Public Function SetDbDef_Suplier()
Dim oRS As ADODB.Recordset
Dim oPs As ADODB.Recordset
Dim strSql As String
Dim stpSql As String

strSql = "Select CORP_ID from tab_corp"

  Set oRS = New ADODB.Recordset
  oRS.Open strSql, conMain, adOpenForwardOnly, adLockReadOnly, adCmdText
      
  Set oPs = New ADODB.Recordset
      
      Do While Not oRS.EOF
         Nm_DxDbGrid.Columns.ColumnByFieldName("BRG_NM").PickColumn.Items.Add (oRS.Fields("BRG_NM").Value)
         stpSql = "SELECT CAB_ID FROM tab_cabang WHERE CORP_ID='" & oRS.Fields("CORP_ID").Value & "'"
            oPs.Open stpSql, conMain, adOpenForwardOnly, adLockReadOnly, adCmdText
              
        oRS.MoveNext
      Loop
  
  oRS.Close
  Set oRS = Nothing
  oPs.Close
  Set oPs = Nothing
End Function
Public Function GetTextMargin(xText As String, maxMargin As Integer) As String
Dim defLenght As Integer
Dim yText As String
Dim i%, J%
defLenght = Len(xText)
'iya pak ngerinya itu cuma modus biar pada dapet
'max = 50
J = 1
For i = 1 To defLenght
    
    yText = Mid(xText, i, 1)
    
   ' If yText = " " And i > (j * () Then
        
    'End If

Next i

GetTextMargin = yText
MsgBox yText
End Function
Public Sub SelectText(xText As XtremeSuiteControls.FlatEdit)
    xText.SelStart = 0
    xText.SelLength = Len(xText.Text)
End Sub
