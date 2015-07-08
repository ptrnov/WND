VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsTrigger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'----DELETE database record
Public Sub DataDel(oConn As ADODB.Connection, _
                     dbTable As String, _
                     dbField As String, _
                     dbValue As String, _
                     Optional dbOptStr As String)

 On Error Resume Next

Dim strQry As String
strQry = "DELETE FROM " & dbTable & " WHERE " & dbField & "=" & dbValue & dbOptStr

'oConn.Open
oConn.Execute strQry
'oConn.Close

End Sub

'----INSERT database record
Public Sub DataIns(oConn As ADODB.Connection, _
                     dbTableField As String, _
                     dbValue As String, _
                     Optional dbOptStr As String)
                    

Dim strQry As String
On Error GoTo ErrorLabel
strQry = "INSERT INTO " & dbTableField & " VALUES(" & dbValue & ") " & dbOptStr
'oConn.Open
oConn.Execute strQry
'oConn.Close
ErrorLabel:
    LOADING.SetParm MDIForm1, 100
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"

End Sub
Public Sub DataInsInto(oConn As ADODB.Connection, _
                     dbTableTarget As String, _
                     dbValueTarget As String, _
                     dbTableSource As String, _
                     dbValueSource As String, _
                     Optional dbOptStr As String)
                    
      On Error Resume Next
Dim strQry As String

strQry = "INSERT INTO " & dbTableTarget & "(" & dbValueTarget & ") SELECT " & dbValueSource & " FROM " & dbTableSource & " " & dbOptStr
'MsgBox strQry

'oConn.Open
oConn.Execute strQry
MsgBox "Proses Berhasil"

'oConn.Close

End Sub
'----UPDATEdatabase record
Public Sub DataUpdate(oConn As ADODB.Connection, _
                     dbTable As String, _
                     dbValue As String, _
                     Optional dbOptStr As String)
                    
 On Error Resume Next
Dim strQry As String

strQry = "UPDATE " & dbTable & " SET " & dbValue & " " & dbOptStr

'MsgBox strQry

'oConn.Open
oConn.Execute strQry
'oConn.Close

End Sub

'----------------------TIGGER TANPA GRID-----------

'---isi data combo box
Public Sub DataFillCombo(objComboBox As ComboBox, _
                     oConn As ADODB.Connection, _
                     dbTable As String, _
                     dbField As String, _
                     dbOptStr As String, _
                     strFieldToShow As String, _
                     Optional strFieldForItemData As String)
                     'objComboBox = nama combo box yang digunakan
                     'oConn = koneksi ke database
                     'dbTable = koneksi ke tabel
                     'dbField = field yang diinginkan
                     'dbOptStr = penambahan SQL statement
                     'strFieldToShow = field yang digunakan pada list
                     
Dim strSql As String
On Error Resume Next
strSql = "SELECT " & dbField & " FROM " & dbTable & " " & dbOptStr

Dim oRS As ADODB.Recordset  'baca data
  Set oRS = New ADODB.Recordset
  oRS.Open strSql, oConn, adOpenForwardOnly, adLockReadOnly, adCmdText
                            
  With objComboBox          'isi combo box
    .Clear
    If strFieldForItemData = "" Then
      Do While Not oRS.EOF      'tanpa ItemData
        .AddItem oRS.Fields(strFieldToShow).Value
        oRS.MoveNext
      Loop
    Else
      Do While Not oRS.EOF      'dengan ItemData
        .AddItem oRS.Fields(strFieldToShow).Value
        '.ItemData(.NewIndex) = oRS.Fields(strFieldForItemData).Value
        oRS.MoveNext
      Loop
    End If
  End With

    objComboBox.ListIndex = 0
  oRS.Close                 'Tidy up
  Set oRS = Nothing

End Sub
Public Sub DataFillText(objTextBox As TextBox, _
                     oConn As ADODB.Connection, _
                     dbTable As String, _
                     dbField As String, _
                     Optional dbOptStr As String)
                     'objComboBox = nama combo box yang digunakan
                     'oConn = koneksi ke database
                     'dbTable = koneksi ke tabel
                     'dbField = field yang diinginkan
                     'dbOptStr = penambahan SQL statement
                     'strFieldToShow = field yang digunakan pada list
                     
Dim strSql As String

strSql = "SELECT " & dbField & " FROM " & dbTable & " " & dbOptStr

Dim oRS As ADODB.Recordset  'baca data
  Set oRS = New ADODB.Recordset
  oRS.Open strSql, oConn, adOpenForwardOnly, adLockReadOnly, adCmdText
                            
  With objTextBox          'isi text box
      Do While Not oRS.EOF
        .Text = oRS.Fields(dbField).Value
        oRS.MoveNext
      Loop
  End With

 ' oRS.Close                 'Tidy up
  Set oRS = Nothing

End Sub
'--- cari data input sesuai nilai yang diInginkan pada database record
Public Sub GetDataField(oConn As ADODB.Connection, _
                     dbTable As String, _
                     dbField As String, _
                     dbCondField As String, _
                     dbValue As String, _
                     Optional dbOptStr As String)

                     'oConn = koneksi ke database
                     'dbTable = koneksi ke tabel
                     'dbField = field yang diinginkan
                     'dbCondField = dengan kondisi field tertentu
                     'dbValue = dengan value tertentu
                    'dbOptStr = penambahan SQL statement
On Error Resume Next
Dim strSql As String
strSql = "SELECT " & dbField & " FROM " & dbTable & " WHERE " & dbCondField & "='" & dbValue & "' " & dbOptStr
'MsgBox strSQL

Dim oRS As ADODB.Recordset  'baca data
  Set oRS = New ADODB.Recordset
  oRS.Open strSql, oConn, adOpenForwardOnly, adLockReadOnly, adCmdText

        sndData = oRS.Fields(dbField).Value
  'oRS.Close                 'Tidy up
  Set oRS = Nothing

End Sub

'running sqlQry dengan kondisi tipe prosedur/text
Public Sub RunSqlStr(oConn As ADODB.Connection, strQry As String, isProcedure As Boolean)
On Error GoTo ErrorLabel
  Dim oRS As ADODB.Recordset
  Set oRS = New ADODB.Recordset
  
    If isProcedure Then
         oRS.Open strQry, oConn, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    Else
        oRS.Open strQry, oConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If
  
  Set oRS = Nothing


ErrorLabel:
    LOADING.SetParm MDIForm1, 100
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"
End Sub
Public Function RunSqlFunction(FuncName As String) As String
On Error GoTo ErrorLabel
        Dim cmdTemp As New ADODB.Command
        Set cmdTemp.ActiveConnection = conMain
        cmdTemp.Parameters.Append cmdTemp.CreateParameter("Ret", adInteger, adParamReturnValue)
        cmdTemp.CommandType = adCmdStoredProc
        cmdTemp.CommandText = FuncName
        cmdTemp.Execute
        RunSqlFunction = cmdTemp.Parameters("Ret")
        
ErrorLabel:
    LOADING.SetParm MDIForm1, 100
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"
End Function
Public Function UpdateUserPerm(xUsrId As String, xMN_ID As String, xSts As Boolean)
On Error Resume Next
Dim qryStr As String
Dim ySts As String
'MsgBox xSts
If xSts Then ySts = "1" Else ySts = "0"

qryStr = "pro_user_per_update('" & Trim(xUsrId) & "','" & Trim(xMN_ID) & "','" & Trim(ySts) & "')"
'MsgBox qryStr
PrjSysTrig.RunSqlStr conMain, qryStr, True

End Function
