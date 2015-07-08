VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSysUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim rsRec As ADODB.Recordset
Dim Qry As String

Public Function ValidasiUserID(USER_ID As String) As String
'===========================================
'====== GET VALIDASI USER ==================
'==========================================
On Error GoTo ErrorLabel
Qry = "Select USR_NM,USER_ID from tab_user where USER_ID='" & USER_ID & "'"
Set rsRec = ExecuteRecordSetMySql(Qry, conMain)
With rsRec
    If .RecordCount <> 0 Then
        ValidasiUserID = .Fields("USR_NM").Value
        UsrID = .Fields("USER_ID").Value
    End If
End With
Set rsRec = Nothing


Exit Function

ErrorLabel:
    If Err.Number = 3021 Then Exit Function
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error" & ". Error Number : " & Err.Number
End Function
Public Function ValidasiPass(USR_PASS As String) As Integer
'===========================================
'====== GET VALIDASI PASSWORD ==============
'===========================================
On Error GoTo ErrorLabel

Qry = "Select USR_NM from tab_user where USR_PASS='" & USR_PASS & "'"
Set rsRec = ExecuteRecordSetMySql(Qry, conMain)
With rsRec
    If Not .EOF Then
        ValidasiPass = 1
    End If
End With
Set rsRec = Nothing
'MsgBox (ValidasiPass)

Exit Function
ErrorLabel:
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"

End Function
Public Function ValidasiRFID(USR_RFID As String) As String
On Error GoTo ErrorLabel

Qry = "Select USER_ID from tab_user where USR_RFID='" & USR_RFID & "'"
Set rsRec = ExecuteRecordSetMySql(Qry, conMain)
With rsRec
    If Not .EOF Then
        ValidasiRFID = .Fields("USER_ID").Value
    End If
End With
Set rsRec = Nothing

Exit Function
ErrorLabel:
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"

End Function


