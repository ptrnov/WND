Attribute VB_Name = "KONEKSI"
Option Explicit

'Ambil string dari file
Public Declare Function GetPrivateProfileString _
Lib "kernel32" Alias "GetPrivateProfileStringA" _
(ByVal lpSectionName As String, ByVal lpKeyName As String, _
ByVal lpDefault As String, ByVal lpReturnedString As String, _
ByVal nSize As Long, ByVal lpFilename As String) As Long

'tulis string dari file
Public Declare Function WritePrivateProfileString _
Lib "kernel32" Alias "WritePrivateProfileStringA" _
(ByVal lpSectionName As String, ByVal lpKeyName As String, _
ByVal lpValueUser As String, ByVal lpFilename As String) As Long
 
Public lpSectionName As String
Public lpKeyName As String
Public lpValueUser As String
Public lpFilename As String
Public lpReturnedString As String
Public nSize As Long

Public RecStt As New ADODB.Recordset
'=== PTR.NOV PAYROLL ====
Public rs1 As New ADODB.Recordset
Public rs2 As New ADODB.Recordset
Public rs3 As New ADODB.Recordset
Public rs4 As New ADODB.Recordset
Public rs5 As New ADODB.Recordset
Public rsStt1 As New ADODB.Recordset
Public rsStt2 As New ADODB.Recordset
Public rsRpt1 As New ADODB.Recordset
Public rsRpt2 As New ADODB.Recordset
Public PrjAbsensi As New ClsAbsensi
'========================


Public conMain As ADODB.Connection
Public conn_db  As ADODB.Connection
Public PrjSysUsr As New ClsSysUser
Public PrjSysMn As New ClsSysMenu
Public PrjSysGrid As New ClsGrid
Public PrjSysTrig As New ClsTrigger
Public prjSysID As New ClsIdCode
Public prjSysReport As New ClsReportExcel
Public PrjExportExcel As New RptExcel
Public prjSysToolTips As New ClsToolTips
Public PrjSysDataInput As New ClsDataInput
Public PrjSysPing As New ClsPing
Public strServer, strUser, strPassword As String

Public StrCon, dbHost, DBName, _
        dbPass, dbType, sndChoice, _
        sndData, NmDep, IdCorp, IdCab, _
        KarNM, CorpNM, CabNM, DepNM, JabNM, UsrNM, IdDep, UsrID, KarId, IdJab As String

Public Sub Main()
'On Error GoTo err_Handler
  '==== Connection by ptr.nov ===============
  Set conMain = New ADODB.Connection
  '==== SQLSERVER 2000 server 10.10.99.2 ===================
  'StrCon = "Provider=SQLNCLI10;" _
         & "SERVER=" & strServer & ";" _
         & "Database=sss;" _
         & "DataTypeCompatibility=80;" _
         & "User Id=sss;" _
         & "Password=asd123;" 'ConectionTimeout=10" ''SQLNCLI10;"

  '==== SQLSERVER 2000 server 10.10.99.2 ===================
  'StrCon = "Provider=SQLNCLI10;" _
         & "SERVER=" & strServer & ";" _
         & "Database=xxx;" _
         & "DataTypeCompatibility=80;" _
         & "User Id=xxx;" _
         & "Password=asd123;" 'ConectionTimeout=10" ''SQLNCLI10;"
         
  '==== MySQLSERVER Eka Laptop ===================
 ' StrCon = "Driver={MySQL ODBC 5.2a Driver};Server=localhost;Database=payroll;" _
            & "User=root;Password=Sp1d3rm4n4;Option=3;"
  
  '==== MySQLSERVER Eka Laptop ===================
  'StrCon = "DRIVER={MySQL ODBC 5.1 Driver};" _
        & "Persist Security Info=False;" _
        & "SERVER=localhost;UID=root;port=3306;" _
        & "PWD=Sp1d3rm4n4;DATABASE=attpayroll;" _
        & "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 163841

  '==== WANINDO CLIENT ===================
  'StrCon = "DRIVER={MySQL ODBC 5.1 Driver};" _
        & "Persist Security Info=False;" _
        & "SERVER=server_absen;UID=root;port=3306;" _
        & "PWD=Sp1d3rm4n4;DATABASE=attpayroll;" _
        & "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 163841
  
  '==== WANINDO CLIENT ===================
  StrCon = "DRIVER={MySQL ODBC 5.1 Driver};" _
        & "Persist Security Info=False;" _
        & "SERVER=" & strServer & ";UID=root;port=3306;" _
        & "PWD=Sp1d3rm4n4;DATABASE=attpayroll06082015;" _
        & "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 163841
        
        
        
    '==== MySQLSERVER Eka Laptop ===================
 'StrCon = "DRIVER={MySQL ODBC 5.2 Driver};" _
        & "Persist Security Info=False;" _
        & "SERVER=192.168.100.103;UID=root;" _
        & "PWD=werchmp;DATABASE=payroll;" ' _
        & "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 163841
         

  'conMain.Properties("Connect Timeout").Value = 2
  'conMain.Properties("General Timeout").Value = 2000
  'conMain.ConnectionTimeout = 60
    'conMain.CommandTimeout = 1000
    conMain.ConnectionString = StrCon
    conMain.Open
  If conMain.State = 1 Then
    'MsgBox "Connection Database Berhasil"
    Else
    MsgBox "Maaf Aplikasi tidak dapat terhubung dengan Server...", vbExclamation
  End If
Exit Sub
'err_Handler:
  'MsgBox Err.Description, vbOKOnly, "Information"
'      If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"
'MsgBox "Error"
End Sub


'---------------------------------------------
'---------- BUKA RECORD KONDISI --------------
'--------------- BY PTR.NOV-------------------
Sub OpRec1(Qry, isProcedure As Boolean)
Set rs1 = New ADODB.Recordset
rs1.CursorLocation = adUseClient
    If isProcedure Then
        rs1.Open Qry, conMain, adOpenStatic, adLockOptimistic, adCmdStoredProc
      Else
        rs1.Open Qry, conMain, adOpenStatic, adLockOptimistic, adCmdText
    End If
End Sub
'---------------------------------------------

'---------------------------------------------
'---------- BUKA RECORD KONDISI --------------
'--------------- BY PTR.NOV-------------------
Sub OpRec2(Qry, isProcedure As Boolean)
Set rs2 = New ADODB.Recordset
rs2.CursorLocation = adUseClient
    If isProcedure Then
        rs2.Open Qry, conMain, adOpenStatic, adLockOptimistic, adCmdStoredProc
        
      Else
        rs2.Open Qry, conMain, adOpenStatic, adLockOptimistic, adCmdText
        
    End If
End Sub
'---------------------------------------------
'---------- BUKA RECORD KONDISI --------------
'--------------- BY PTR.NOV-------------------
Sub OpRec3(Qry, isProcedure As Boolean)
Set rs3 = New ADODB.Recordset
rs3.CursorLocation = adUseClient
    If isProcedure Then
        rs3.Open Qry, conMain, adOpenStatic, adLockOptimistic, adCmdStoredProc
      Else
        rs3.Open Qry, conMain, adOpenStatic, adLockOptimistic, adCmdText
    End If
End Sub
'---------------------------------------------
'---------- BUKA RECORD KONDISI --------------
'--------------- BY PTR.NOV-------------------
Sub OpRec4(Qry, isProcedure As Boolean)
Set rs4 = New ADODB.Recordset
rs4.CursorLocation = adUseClient
    If isProcedure Then
        rs4.Open Qry, conMain, adOpenStatic, adLockOptimistic, adCmdStoredProc
      Else
        rs4.Open Qry, conMain, adOpenStatic, adLockOptimistic, adCmdText
    End If
End Sub
'---------------------------------------------
'---------- BUKA RECORD KONDISI --------------
'--------------- BY PTR.NOV-------------------
Sub OpRec5(Qry, isProcedure As Boolean)
Set rs5 = New ADODB.Recordset
rs5.CursorLocation = adUseClient
    If isProcedure Then
        rs5.Open Qry, conMain, adOpenStatic, adLockOptimistic, adCmdStoredProc
      Else
        rs5.Open Qry, conMain, adOpenStatic, adLockOptimistic, adCmdText
    End If
End Sub



'---------------------------------------------
'---------- BUKA RECORD KONDISI --------------
'--------------- BY PTR.NOV-------------------
Sub OpRecStt1(Qry, isProcedure As Boolean)
Set rsStt1 = New ADODB.Recordset
rsStt1.CursorLocation = adUseClient
    If isProcedure Then
        rsStt1.Open Qry, conMain, adOpenStatic, adLockOptimistic, adCmdStoredProc
      Else
        rsStt1.Open Qry, conMain, adOpenStatic, adLockOptimistic, adCmdText
    End If
End Sub
'---------------------------------------------
'---------- BUKA RECORD KONDISI --------------
'--------------- BY PTR.NOV-------------------
Sub OpRecStt2(Qry, isProcedure As Boolean)
Set rsStt2 = New ADODB.Recordset
rsStt2.CursorLocation = adUseClient
    If isProcedure Then
        rsStt2.Open Qry, conMain, adOpenStatic, adLockOptimistic, adCmdStoredProc
      Else
        rsStt2.Open Qry, conMain, adOpenStatic, adLockOptimistic, adCmdText
    End If
End Sub

'---------------------------------------------
'---------- BUKA RECORD REPORT EXCEL ---------
'--------------- BY PTR.NOV-------------------
Sub OpnRecRpt1(Qry, isProcedure As Boolean)
Set rsRpt1 = New ADODB.Recordset
rsRpt1.CursorLocation = adUseClient
    If isProcedure Then
        rsRpt1.Open Qry, conMain, adOpenStatic, adLockOptimistic, adCmdStoredProc
      Else
        rsRpt1.Open Qry, conMain, adOpenStatic, adLockOptimistic, adCmdText
    End If
End Sub

'---------------------------------------------
'---------- BUKA RECORD REPORT EXCEL ---------
'--------------- BY PTR.NOV-------------------
Sub OpnRecRpt2(Qry, isProcedure As Boolean)
Set rsRpt2 = New ADODB.Recordset
rsRpt2.CursorLocation = adUseClient
    If isProcedure Then
        rsRpt2.Open Qry, conMain, adOpenStatic, adLockOptimistic, adCmdStoredProc
      Else
        rsRpt2.Open Qry, conMain, adOpenStatic, adLockOptimistic, adCmdText
    End If
End Sub



Public Function ExecuteRecordSetMySqlProc(sql As String, con As ADODB.Connection) As Recordset
  'to execute sql statements which results in a recordset
  Dim rs As ADODB.Recordset
  Dim cmd As ADODB.Command
    
On Error GoTo ErrorLabel
    
  Set cmd = New ADODB.Command
  cmd.ActiveConnection = con
  cmd.CommandType = adCmdStoredProc
  cmd.CommandText = sql
    
  Set ExecuteRecordSetMySqlProc = cmd.Execute()

Exit Function
ErrorLabel:
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"
End Function
Public Function ExecuteRecordSetMySql(sql As String, con As ADODB.Connection) As Recordset
  'to execute sql statements which results in a recordset
  Dim rs As ADODB.Recordset
  Dim cmd As ADODB.Command
    
On Error GoTo ErrorLabel
    
  Set cmd = New ADODB.Command
  cmd.ActiveConnection = con
  cmd.CommandText = sql
    
  Set ExecuteRecordSetMySql = cmd.Execute()

Exit Function
ErrorLabel:
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"
End Function

Public Function ExecuteNonQueryMySql(sql As String, con As ADODB.Connection) As Long
  'to execute non query sql statements such as insert, delete, update
  Dim result As Integer
  Dim cmd As ADODB.Command
    
On Error GoTo ErrorLabel
  Set cmd = New ADODB.Command
  cmd.ActiveConnection = con
  cmd.CommandText = sql
   
  cmd.Execute result
  'result = records affected
    
  ExecuteNonQueryMySql = result

Exit Function
ErrorLabel:
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"
End Function

' NOTED
'---------------------------------------------
'---------- BUKA RECORD KONDISI --------------
'--------------- BY PTR.NOV-------------------
Sub OpRcKondisiA(Querry)
'Set RcKondisiA = New ADODB.Recordset
'RcKondisiA.CursorLocation = adUseClient
'RcKondisiA.Open Querry, sysprof_condb, adOpenStatic, adLockOptimistic
End Sub

'fungsi pencarian data pada database,tabel hasilnya string
Public Function GetDtaTbl(conDB As ADODB.Connection, sqlQry As String, dbField As String) As String
On Error Resume Next
Dim oRS As ADODB.Recordset  'baca data
  Set oRS = New ADODB.Recordset
  oRS.Open sqlQry, conDB, adOpenForwardOnly, adLockReadOnly, adCmdText
        'MsgBox "" & oRS.Fields(dbField).Value
         GetDtaTbl = oRS.Fields(dbField).Value
  If Len(GetDtaTbl) < 1 Then GetDtaTbl = ""
  oRS.Close                 'Tidy up
  Set oRS = Nothing
End Function

'fungsi pencarian LIST data pada database,tabel hasilnya string ARRAY
Public Function GetListTabel(oConn As ADODB.Connection, strSql As String, tblField As String) As String()
Dim J(1 To 999999) As String
Dim K As Integer
Dim oRS As ADODB.Recordset  'baca data
  Set oRS = New ADODB.Recordset
Erase J
K = 1
  oRS.Open strSql, oConn, adOpenForwardOnly, adLockReadOnly, adCmdText
      Do While Not oRS.EOF      'tanpa ItemData
        J(K) = oRS.Fields(tblField).Value
        oRS.MoveNext
        K = K + 1
      Loop
  GetListTabel = J
  oRS.Close                 'Tidy up
  Set oRS = Nothing
End Function
'fungsi pencarian LIST data pada database,tabel hasilnya string ARRAY
Public Function GetListTabelProc(oConn As ADODB.Connection, strSql As String, tblField As String) As String()
Dim J(1 To 999999) As String
Dim K As Integer
Dim oRS As ADODB.Recordset  'baca data
  Set oRS = New ADODB.Recordset
Erase J
K = 1
  oRS.Open strSql, oConn, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
      Do While Not oRS.EOF      'tanpa ItemData
        J(K) = oRS.Fields(0).Value
        oRS.MoveNext
        K = K + 1
      Loop
  GetListTabelProc = J
  oRS.Close                 'Tidy up
  Set oRS = Nothing
End Function

'fungsi untuk mengecek file ada/tidak
Public Function CekStatusFile(NmFile As String, LokasiFile As String) As Integer
Dim lSize As Long
   On Error Resume Next
   lSize = -1
   lSize = FileLen(LokasiFile & NmFile)
   If lSize > -1 Then
      CekStatusFile = 1
   Else
      CekStatusFile = 0
   End If
End Function

Public Sub InfoServ()
On Error GoTo ErrorLabel
If CekStatusFile("\config.ini", App.Path) Then
    lpFilename = App.Path & "\config.ini"

    'Server
    lpReturnedString = Space$(255)
    nSize = Len(lpReturnedString)
    nSize = GetPrivateProfileString("Koneksi", _
    "Server", " ", lpReturnedString, 50, lpFilename)
    lpReturnedString = Mid(lpReturnedString, 1, nSize)
    strServer = lpReturnedString

    'User
    lpReturnedString = ""
    lpReturnedString = Space$(255)
    nSize = Len(lpReturnedString)
    nSize = GetPrivateProfileString("Koneksi", _
    "User", " ", lpReturnedString, 50, lpFilename)
    lpReturnedString = Mid(lpReturnedString, 1, nSize)
    strUser = lpReturnedString

    'Password
    lpReturnedString = ""
    lpReturnedString = Space$(255)
    nSize = Len(lpReturnedString)
    nSize = GetPrivateProfileString("Koneksi", _
    "Password", " ", lpReturnedString, 50, lpFilename)
    lpReturnedString = Mid(lpReturnedString, 1, nSize)
    strPassword = lpReturnedString

Else
    MsgBox "File Konfigurasi Tidak Ditemukan !", vbCritical, "Error"
    New_Config
End If

Exit Sub
ErrorLabel:
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"
End Sub
Public Sub ProfileSaveItem(lpSectionName As String, _
lpKeyName As String, lpValueUser As String, lpFilename As String)
   Call WritePrivateProfileString(lpSectionName, _
lpKeyName, lpValueUser, lpFilename)
End Sub

'buat file config.ini jika tidak ditemukan
Public Sub New_Config()
Dim i As Integer
On Error GoTo ErrorLabel

Open App.Path & "\config.ini" For Output As #1
Print #1, "// --------------------------------------------"
Print #1, "// | Lukison Group System " & " Ver. 1.0 Beta |"
Print #1, "// --------------------------------------------"
Print #1, " "
Print #1, "// Hati-hati dalam merubah file ini"
Print #1, " "
Print #1, "[Koneksi]"
Print #1, "Server=0.0.0.0"
Close #1

ope:

Exit Sub
ErrorLabel:
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"
End Sub


Sub OpnRecStt(Qry)
On Error GoTo ErrorLabel

Set RecStt = New ADODB.Recordset
With RecStt
.CursorLocation = adUseClient
.Open Qry, conMain, adOpenStatic, adLockOptimistic

ErrorLabel:
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"
End With
End Sub
'mengambil informasi user, [Karyawan ID, Perusahaan, Cabang, etc]

Public Sub GetUserInfo(userID As String)
On Error GoTo ErrorLabel
    KarId = GetDtaTbl(conMain, _
            "SELECT KAR_ID FROM tab_user WHERE USER_ID='" & userID & "'", "KAR_ID")
    IdJab = GetDtaTbl(conMain, _
            "SELECT JAB_ID FROM karyawan WHERE KAR_ID='" & KarId & "'", "JAB_ID")
    IdDep = GetDtaTbl(conMain, _
            "SELECT DEP_ID FROM karyawan WHERE KAR_ID='" & KarId & "'", "DEP_ID")
    IdCorp = GetDtaTbl(conMain, _
            "SELECT CORP_ID FROM karyawan WHERE KAR_ID='" & KarId & "'", "CORP_ID")
    IdCab = GetDtaTbl(conMain, _
            "SELECT CAB_ID FROM karyawan WHERE KAR_ID='" & KarId & "'", "CAB_ID")
    CorpNM = GetDtaTbl(conMain, _
            "SELECT CORP_NM FROM corp WHERE CORP_ID='" & IdCorp & "'", "CORP_NM")
    CabNM = GetDtaTbl(conMain, _
            "SELECT CAB_NM FROM cabang WHERE CAB_ID='" & IdCab & "' AND CORP_ID='" & IdCorp & "'", "CAB_NM")
    DepNM = GetDtaTbl(conMain, _
            "SELECT DEP_NM FROM departemen WHERE DEP_ID='" & IdDep & "'", "DEP_NM")
    JabNM = GetDtaTbl(conMain, _
            "SELECT JAB_NM FROM jabatan WHERE JAB_ID='" & IdJab & "'", "JAB_NM")
    UsrNM = GetDtaTbl(conMain, _
            "SELECT USR_NM FROM tab_user WHERE USER_ID='" & userID & "'", "USR_NM")
    UsrID = userID
Exit Sub
ErrorLabel:
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"
End Sub


