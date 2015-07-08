VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------
'------------ MENU USR PROFILE ---------------
'----------------- GetSysUSER ----------------
'--------------- BY PTR.NOV ------------------
'---------------------------------------------
Public Sub GetBaseGrid(Nm_DxDbGrid As dxDBGrid, AryColomn As Variant, SttBND As Boolean, fixBND As Boolean, StrBnd As Variant, SttQry As Boolean, StrQuerry As String, KeyField As String, editAble As Boolean)

Dim a As Long
Dim NewCol As dxGridColumn
Dim i As Integer

On Error GoTo ErrorLabel

Nm_DxDbGrid.Dataset.Active = False
If SttBND Then
'------------------
'== BIND SETTING ==
'------------------
    Nm_DxDbGrid.Options.Set (egoShowBands)
    Nm_DxDbGrid.Options.Set (egoSmartRefresh)
    Nm_DxDbGrid.Options.Set (egoSmartReload)
    Nm_DxDbGrid.Options.Unset (egoShowIndicator)
    Nm_DxDbGrid.Options.Set (egoAutoWidth)
    Nm_DxDbGrid.Options.Set (egoRowSelect)
    With Nm_DxDbGrid.Bands   'bands adjusting
      .Add
      .Item(0).Caption = StrBnd(0)
      .Item(1).Caption = StrBnd(1)
      .BandFont.Bold = True
      .BandColor = &HC0C0C0
      If fixBND Then .Item(0).Fixed = bfLeft
    End With

    With Nm_DxDbGrid.Bands
        If .Count <> 2 Then
           .Remove (.Count - 1)
        End If
    End With
End If


With Nm_DxDbGrid.Columns
'-------------------
'== COLUMNS ARRAY ==
'-------------------

.DestroyColumns
'On Error Resume Next
    For a = 0 To UBound(AryColomn)
        Set NewCol = .Add(AryColomn(a)(3))
        NewCol.ObjectName = AryColomn(a)(0)
        NewCol.Caption = AryColomn(a)(1) '
        NewCol.FieldName = AryColomn(a)(2) 'isi nama tabel
        NewCol.RowIndex = AryColomn(a)(4) 'isi tabel record
        NewCol.BandIndex = AryColomn(a)(5) 'index tabel
        NewCol.Width = AryColomn(a)(6)
        '-header-
        NewCol.HeaderAlignment = taCenter
        '-Disable Editor-
         
        'MsgBox AryColomn(3)(4)
         If NewCol.ColumnType = gedCheckEdit Then
            NewCol.CheckColumn.ValueChecked = 1
            NewCol.CheckColumn.ValueUnchecked = 0
        End If
        

        
        If AryColomn(a)(7) = 1 Then
            NewCol.DisableEditor = True
            NewCol.Color = &HC0E0FF
        End If
        '-Alignment-
        If AryColomn(a)(8) = 1 Then
                NewCol.Alignment = taLeftJustify
            ElseIf AryColomn(a)(8) = 2 Then
                NewCol.Alignment = taCenter
            ElseIf AryColomn(a)(8) = 3 Then
                NewCol.Alignment = taRightJustify
        End If
        '-Desimal-
        If AryColomn(a)(9) = 1 Then
            NewCol.Visible = False
            
        End If
        
        
    Next

    'MsgBox UBound(AryColomn)
End With

'--------------------------------------
'--------- TABEL/RECORD FIRST ---------
'--------------------------------------
With Nm_DxDbGrid
    .KeyField = KeyField
    .Dataset.ADODataset.ConnectionString = StrCon
    .Dataset.ADODataset.CursorLocation = clUseClient
    If editAble Then .Dataset.ADODataset.LockType = ltBatchOptimistic Else .Dataset.ADODataset.LockType = ltOptimistic
    
    If SttQry = False Then
    '------------------
    '== Normal Query ==
    '------------------
    .Dataset.ADODataset.CommandType = cmdText
    Else
    '-------------------------
    '== PLSQL cmdStoredProc ==
    '-------------------------
    .Dataset.ADODataset.CommandType = cmdStoredProc
    End If
    .Dataset.ADODataset.CommandText = StrQuerry
    .Dataset.Active = True
End With

Exit Sub
ErrorLabel:
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"
End Sub
'------------GRID BASE---------------
Public Sub GetBaseGridAll(Nm_DxDbGrid As dxDBGrid, AryColomn As Variant, StrQuerry As String, KeyField As String, StrType As Boolean)
Dim a As Long
Dim NewCol As dxGridColumn
Dim i As Integer

On Error GoTo ErrorLabel

    'set grid options
    Nm_DxDbGrid.Dataset.Active = False
    'Nm_DxDbGrid.Options.Unset (egoShowIndicator)
    'Nm_DxDbGrid.Options.Set (egoAutoWidth)

With Nm_DxDbGrid.Columns
.DestroyColumns
'On Error Resume Next
    For a = 0 To UBound(AryColomn)
        Set NewCol = .Add(AryColomn(a)(3))
        NewCol.ObjectName = AryColomn(a)(0)
        NewCol.Caption = AryColomn(a)(1) '
        NewCol.FieldName = AryColomn(a)(2) 'isi nama tabel
        NewCol.RowIndex = AryColomn(a)(4) 'isi tabel record
        NewCol.BandIndex = AryColomn(a)(5) 'index tabel
        NewCol.Width = AryColomn(a)(6)
        NewCol.Visible = AryColomn(a)(9)

        'PROPERTI KOLOM
        'cheK
         If NewCol.ColumnType = gedCheckEdit Then
            NewCol.CheckColumn.ValueChecked = 1
            NewCol.CheckColumn.ValueUnchecked = 0
        End If
        
        'spin
        If NewCol.ColumnType = gedSpinEdit Then
            NewCol.SpinColumn.MaxValue = 1000000
            NewCol.SpinColumn.MinValue = 1
        End If
        
        '-Alignment-
        If AryColomn(a)(8) = 1 Then
                NewCol.Alignment = taLeftJustify
            ElseIf AryColomn(a)(8) = 2 Then
                NewCol.Alignment = taCenter
            ElseIf AryColomn(a)(8) = 3 Then
                NewCol.Alignment = taRightJustify
        End If
        NewCol.HeaderAlignment = taCenter
    Next
End With
'--------------------------------------
'--------- TABEL/RECORD FIRST ---------
'--------------------------------------
With Nm_DxDbGrid
    .KeyField = KeyField
    If StrType = True Then
        .Dataset.ADODataset.CommandType = cmdText
    Else
        .Dataset.ADODataset.CommandType = cmdStoredProc
    End If
    .Dataset.ADODataset.ConnectionString = StrCon
    .Dataset.ADODataset.CursorLocation = clUseClient
    .Dataset.ADODataset.LockType = ltOptimistic
    .Dataset.ADODataset.CommandText = StrQuerry
    .Dataset.Active = True
End With

Exit Sub
ErrorLabel:
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"

End Sub

'------------GRID RO---------------
Public Sub GetBaseGridRoPoID(Nm_DxDbGrid As dxDBGrid, AryColomn As Variant, StrQuerry As String, KeyField As String, StrType As Boolean)
Dim a As Long
Dim NewCol As dxGridColumn
Dim i As Integer
'On Error Resume Next
    'set grid options
    Nm_DxDbGrid.Dataset.Active = False
    'Nm_DxDbGrid.Options.Unset (egoShowIndicator)
    'Nm_DxDbGrid.Options.Set (egoAutoWidth)
'On Error GoTo ErrorLabel

With Nm_DxDbGrid.Columns
.DestroyColumns
'On Error Resume Next
    For a = 0 To UBound(AryColomn)
        Set NewCol = .Add(AryColomn(a)(3))
        NewCol.ObjectName = AryColomn(a)(0)
        NewCol.Caption = AryColomn(a)(1) '
        NewCol.FieldName = AryColomn(a)(2) 'isi nama tabel
        NewCol.RowIndex = AryColomn(a)(4) 'isi tabel record
        NewCol.BandIndex = AryColomn(a)(5) 'index tabel
        NewCol.Width = AryColomn(a)(6)
        NewCol.Visible = AryColomn(a)(9)

        'PROPERTI KOLOM
        'cheK
         If NewCol.ColumnType = gedCheckEdit Then
            NewCol.CheckColumn.ValueChecked = 1
            NewCol.CheckColumn.ValueUnchecked = 0
        End If
        
        'spin
        If NewCol.ColumnType = gedSpinEdit Then
            NewCol.SpinColumn.MaxValue = 1000000
            NewCol.SpinColumn.MinValue = 1
        End If
        'read only
        If AryColomn(a)(7) = 1 Then
            NewCol.DisableEditor = True
            NewCol.Color = &HC0E0FF
            
        End If

        '-Alignment-
        If AryColomn(a)(8) = 1 Then
                NewCol.Alignment = taLeftJustify
            ElseIf AryColomn(a)(8) = 2 Then
                NewCol.Alignment = taCenter
            ElseIf AryColomn(a)(8) = 3 Then
                NewCol.Alignment = taRightJustify
        End If
        NewCol.HeaderAlignment = taCenter
    Next
End With
'--------------------------------------
'--------- TABEL/RECORD FIRST ---------
'--------------------------------------
With Nm_DxDbGrid
    .KeyField = KeyField
    If StrType = True Then
        .Dataset.ADODataset.CommandType = cmdText
    Else
        .Dataset.ADODataset.CommandType = cmdStoredProc
    End If
    
    .Dataset.ADODataset.ConnectionString = StrCon
    .Dataset.ADODataset.CursorLocation = clUseClient
    .Dataset.ADODataset.LockType = ltBatchOptimistic ' ltOptimistic
    .Dataset.ADODataset.CommandText = StrQuerry
    .Dataset.Active = True
End With

Exit Sub
ErrorLabel:
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"
End Sub

Public Sub FillGridFast(Nm_DxDbGrid As dxDBGrid, SttBND As Boolean, fixBND As Boolean, SttQry As Boolean, StrQuerry As String, KeyField As String)
 Dim a As Long
Dim NewCol As dxGridColumn
Dim i As Integer
On Error GoTo ErrorLabel

Nm_DxDbGrid.Dataset.Active = False
 Nm_DxDbGrid.Columns.DestroyColumns

'--------------------------------------
'--------- TABEL/RECORD FIRST ---------
'--------------------------------------
With Nm_DxDbGrid
    .KeyField = KeyField
    .Dataset.ADODataset.ConnectionString = StrCon
    .Dataset.ADODataset.CursorLocation = clUseClient
    .Dataset.ADODataset.LockType = ltOptimistic
    If SttQry = False Then
    '------------------
    '== Normal Query ==
    '------------------
    .Dataset.ADODataset.CommandType = cmdText
    Else
    '-------------------------
    '== PLSQL cmdStoredProc ==
    '-------------------------
    .Dataset.ADODataset.CommandType = cmdStoredProc
    End If
    .Dataset.ADODataset.CommandText = StrQuerry
    .Dataset.Active = True
    
    .Columns.RetrieveFields
    'MsgBox .Columns.Count
    For i = 1 To .Columns.Count - 1
     .Columns(i).Visible = 1
     '.Columns(i).FieldName=
    Next
    'Refresh_RS
    '.M.EndUpdate
   ' If .Dataset.RecordCount > 1 Then mnuAdd.Enabled = False
End With

Exit Sub
ErrorLabel:
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"
End Sub
Public Sub GetBaseGridRoPoITEM(Nm_DxDbGrid As dxDBGrid, AryColomn As Variant, StrQuerry As String, IsCmdTxt As Boolean, RowIndex As Boolean, KeyField As String)
Dim i, a As Integer
Dim NewCol As dxGridColumn

On Error GoTo ErrorLabel

Nm_DxDbGrid.Dataset.Active = False
Nm_DxDbGrid.Dataset.Close

With Nm_DxDbGrid.Columns
.DestroyColumns

 For a = 0 To UBound(AryColomn)
        Set NewCol = .Add(AryColomn(a)(3))
        NewCol.ObjectName = AryColomn(a)(0)
        NewCol.Caption = AryColomn(a)(1)
        NewCol.FieldName = AryColomn(a)(2)
        NewCol.RowIndex = AryColomn(a)(4)
        NewCol.BandIndex = AryColomn(a)(5)
        NewCol.Width = AryColomn(a)(6)
        NewCol.Visible = AryColomn(a)(9)
        
     'PROPERTI KOLOM
        'cheK
         If NewCol.ColumnType = gedCheckEdit Then
            NewCol.CheckColumn.ValueChecked = 1
            NewCol.CheckColumn.ValueUnchecked = 0
        End If
        
        'spin
        If NewCol.ColumnType = gedSpinEdit Then
            NewCol.SpinColumn.MaxValue = 1000000
            NewCol.SpinColumn.MinValue = 1
        End If
        'read only
        If AryColomn(a)(7) = 1 Then
            NewCol.DisableEditor = True
            NewCol.Color = &HC0E0FF
            
        End If
        
        '-Alignment-
        If AryColomn(a)(8) = 1 Then
                NewCol.Alignment = taLeftJustify
            ElseIf AryColomn(a)(8) = 2 Then
                NewCol.Alignment = taCenter
            ElseIf AryColomn(a)(8) = 3 Then
                NewCol.Alignment = taRightJustify
        End If
         NewCol.HeaderAlignment = taCenter
 Next
 
 End With
 
 With Nm_DxDbGrid
    .KeyField = KeyField
    If IsCmdTxt Then .Dataset.ADODataset.CommandType = cmdText Else .Dataset.ADODataset.CommandType = cmdStoredProc
    .Dataset.ADODataset.ConnectionString = StrCon
    .Dataset.ADODataset.CursorLocation = clUseClient
    .Dataset.ADODataset.LockType = ltBatchOptimistic
    .Dataset.ADODataset.CommandText = StrQuerry
    .Dataset.Active = True
 End With
Last:

Exit Sub
ErrorLabel:
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"
End Sub
Public Sub FillGridCombo(Nm_DxDbGrid As dxDBGrid, strSql As String, dbField As String, clmField As String)
On Error Resume Next
Dim oRS As ADODB.Recordset
  Set oRS = New ADODB.Recordset
  oRS.Open strSql, conMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  Nm_DxDbGrid.Columns.ColumnByFieldName("BRG_NM").PickColumn.Items.Clear
      Do While Not oRS.EOF
       'MsgBox oRS.Fields("BRG_NM").Value
         Nm_DxDbGrid.Columns.ColumnByFieldName("BRG_NM").PickColumn.Items.Add (oRS.Fields("BRG_NM").Value)
        oRS.MoveNext
      Loop
  'oRS.Close                 'Tidy up
  Set oRS = Nothing
End Sub

Public Sub SetSpin(Nm_DxDbGrid As dxDBGrid, dbField As String, MinVal As Integer, MaxVal As Integer)
With Nm_DxDbGrid.Columns
    .ColumnByFieldName(dbField).SpinColumn.MinValue = MinVal
    .ColumnByFieldName(dbField).SpinColumn.MaxValue = MaxVal
End With
End Sub

Public Sub KeyPress(Nm_DxDbGrid As dxDBGrid, KeySnd As Integer)
Dim N As Integer

With Nm_DxDbGrid.Dataset
 Select Case KeySnd
 Case Is = 9
  ' MsgBox "colom " & Nm_DxDbGrid.Columns.FocusedIndex & "," & Nm_DxDbGrid.Columns.Count - 1 & " | " & "Row " & .RecNo & "," & .RecordCount
  
  If .RecordCount = .RecNo And Nm_DxDbGrid.Columns.FocusedIndex = Nm_DxDbGrid.Columns.VisibleCount - 1 Then
   'MsgBox "colom " & Nm_DxDbGrid.Columns.FocusedIndex & "," & Nm_DxDbGrid.Columns.Count - 1 & " | " & "Row " & .Dataset.RecNo & "," & .Dataset.RecordCount
    .Append
    .FieldValues("NO") = .RecordCount + 1
    .Refresh
    .Close
    .Open
    .FindLast
  End If
  
 Case Is = 46
    If .RecordCount = 0 Then
        MsgBox "Data Kosong"
    Else
        .Delete
        .Refresh
            'MsgBox "KOsong"
           For N = 1 To .RecordCount
            .Edit
            .RecNo = N
            .FieldValues("NO") = N
            .Refresh
            .Close
            .Open
            '.FindLast
            'dxDBGrid1.Dataset.Post
            Next
 
        If .RecordCount = 0 Then
            .Close
            Nm_DxDbGrid.Columns.DestroyColumns
        End If
    End If
  End Select
  End With
End Sub

Public Sub LookUpField(Nm_DxDbGrid As dxDBGrid, strqry As String, clmField As String, dbField As String)
Dim i As Integer
Nm_DxDbGrid.Dataset.Active = False
For i = 0 To Nm_DxDbGrid.Columns.Count - 1
   If Nm_DxDbGrid.Columns(i).ColumnType = gedLookupEdit Then
     With Nm_DxDbGrid.Columns(i).LookupColumn
       .LookupDataset.ADODataset.ConnectionString = conMain
       Select Case Nm_DxDbGrid.Columns(i).FieldName
        Case clmField
         .ListFieldName = dbField
         .ListWidth = 300
         .ListAutoWidth = True
        'Case "COLUMNPART"
        ' .ListFieldName = "Description;OnHand;OnOrder;VendorName"
        ' .ListWidth = 450
        ' .ListAutoWidth = True
        ' Nm_DxDbGrid.Columns(i).DropDownColumn.DropDownRows = 15
         'Case "COLUMNVENDOR"
         '.ListFieldName = "VendorName;PartNo"
         '.ListWidth = 250
         '.ListAutoWidth = True
         'Nm_DxDbGrid.Columns(i).DropDownColumn.DropDownRows = 10
       End Select
     End With
   End If
Next
Nm_DxDbGrid.Dataset.Active = False

End Sub



'---------------------------------------------
'------------ PERSENSI PERSONAL --------------
'--------------- BY PTR.NOV ------------------
'---------------------------------------------
Public Sub GetGrid_Persensi(Nm_DxDbGrid As dxDBGrid, AryColomn As Variant, SttBND As Boolean, fixBND As Boolean, StrBnd As Variant, SttQry As Boolean, StrQuerry As String, KeyField As String, editAble As Boolean)

Dim a As Long
Dim NewCol As dxGridColumn
Dim i As Integer

'On Error GoTo ErrorLabel

Nm_DxDbGrid.Dataset.Active = False
If SttBND Then
'------------------
'== BIND SETTING ==
'------------------
    Nm_DxDbGrid.Options.Set (egoShowBands)
    Nm_DxDbGrid.Options.Set (egoSmartRefresh)
    Nm_DxDbGrid.Options.Set (egoSmartReload)
    Nm_DxDbGrid.Options.Unset (egoShowIndicator)
    'Nm_DxDbGrid.Options.Set (egoAutoWidth)
    Nm_DxDbGrid.Options.Set (egoRowSelect)
    Nm_DxDbGrid.Options.Set (egoKeepColumnWidth)
    With Nm_DxDbGrid.Bands   'bands adjusting
      .Add
      .Item(0).Caption = StrBnd(0)
      .Item(1).Caption = StrBnd(1)
      .BandFont.Bold = True
      .BandColor = &HC0C0C0
      If fixBND Then .Item(0).Fixed = bfLeft
    End With

    With Nm_DxDbGrid.Bands
        If .Count <> 2 Then
           .Remove (.Count - 1)
        End If
    End With
End If


With Nm_DxDbGrid.Columns
'-------------------
'== COLUMNS ARRAY ==
'-------------------

.DestroyColumns
'On Error Resume Next
    For a = 0 To UBound(AryColomn)
        Set NewCol = .Add(AryColomn(a)(3))
        NewCol.ObjectName = AryColomn(a)(0)
        NewCol.Caption = AryColomn(a)(1) '
        NewCol.FieldName = AryColomn(a)(2) 'isi nama tabel
        NewCol.RowIndex = AryColomn(a)(4) 'isi tabel record
        NewCol.BandIndex = AryColomn(a)(5) 'index tabel
        NewCol.Width = AryColomn(a)(6)
        
        '-header Alignment-
        NewCol.HeaderAlignment = taCenter
         
        '--- CHECKLIST AUTO ---
         If NewCol.ColumnType = gedCheckEdit Then
            NewCol.CheckColumn.ValueChecked = 1
            NewCol.CheckColumn.ValueUnchecked = 0
        End If
        
        '--- COLOR ---
        If AryColomn(a)(7) = 1 Then
            NewCol.DisableEditor = True
            'NewCol.Color = &HFF7D87
            ' NewCol.Color = &HF0FFF1
             NewCol.Color = &HCBFAFE
        ElseIf AryColomn(a)(7) = 2 Then
            'NewCol.Color = &HFF7D87
        ElseIf AryColomn(a)(7) = 3 Then
            NewCol.Color = &HFECAFF
        ElseIf AryColomn(a)(7) = 4 Then
            'NewCol.Color = &HCBFAFE
            NewCol.Color = &HF0FFF1
        ElseIf AryColomn(a)(7) = 5 Then
            NewCol.Color = &HFFC0C0
        ElseIf AryColomn(a)(7) = 6 Then
            NewCol.DisableEditor = True
        ElseIf AryColomn(a)(7) = 7 Then
            NewCol.Color = &HBBF8FB
        ElseIf AryColomn(a)(7) = 8 Then
            NewCol.Color = &HFFC0C0
        ElseIf AryColomn(a)(7) = 9 Then
            NewCol.Color = &HFFC0C0
            a = &HFEFCE2
        ElseIf AryColomn(a)(7) = 10 Then
            NewCol.Color = &HFEFCE2
        ElseIf AryColomn(a)(7) = 11 Then
            NewCol.Color = &HFFFFFF
            NewCol.DisableEditor = False
        End If

        '-Alignment-
        If AryColomn(a)(8) = 1 Then
                NewCol.Alignment = taLeftJustify
            ElseIf AryColomn(a)(8) = 2 Then
                NewCol.Alignment = taCenter
            ElseIf AryColomn(a)(8) = 3 Then
                NewCol.Alignment = taRightJustify
        End If
        
        '-HIDDEN SHOE-
        If AryColomn(a)(9) = 1 Then
            NewCol.Visible = False
            
        End If
        
        '-FORMAT TIME OR TANGGAL -
        If AryColomn(a)(10) = 1 Then
            NewCol.DisplayFormat = "HH:mm:ss"
        ElseIf AryColomn(a)(10) = 2 Then
            NewCol.DisplayFormat = "yyyy-mm-dd HH:mm:ss"
        ElseIf AryColomn(a)(10) = 3 Then
            NewCol.DisplayFormat = "yyyy-mm-dd"
        ElseIf AryColomn(a)(10) = 4 Then
            NewCol.DisplayFormat = "yyyy"
        ElseIf AryColomn(a)(10) = 5 Then
            NewCol.DisplayFormat = "DD"
        ElseIf AryColomn(a)(10) = 6 Then
            NewCol.DisplayFormat = "#,###"
        ElseIf AryColomn(a)(10) = 7 Then
            NewCol.DisplayFormat = "#.###"
        End If
         
        '-Disable Editor-
        'If AryColomn(A)(11) = 1 Then
        '    NewCol.DisableEditor = True
        'End If
    Next

    'MsgBox UBound(AryColomn)
End With

'--------------------------------------
'--------- TABEL/RECORD FIRST ---------
'--------------------------------------
With Nm_DxDbGrid
    .KeyField = KeyField
    .Dataset.ADODataset.ConnectionString = StrCon
    .Dataset.ADODataset.CursorLocation = clUseClient
    If editAble Then .Dataset.ADODataset.LockType = ltBatchOptimistic Else .Dataset.ADODataset.LockType = ltOptimistic
    
    If SttQry = False Then
    '------------------
    '== Normal Query ==
    '------------------
    .Dataset.ADODataset.CommandType = cmdText
    Else
    '-------------------------
    '== PLSQL cmdStoredProc ==
    '-------------------------
    .Dataset.ADODataset.CommandType = cmdStoredProc
    End If
    .Dataset.ADODataset.CommandText = StrQuerry
    '.Dataset.Active = True
    '.Dataset.Open
End With

Exit Sub
'ErrorLabel:
'    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"
End Sub


'---------------------------------------------
'--------- MANUAL LOG IN/OUT MAINTAIN -------
'--------------- BY PTR.NOV ------------------
'---------------------------------------------
Public Sub Grid_LogMaintain(Nm_DxDbGrid As dxDBGrid, AryColomn As Variant, SttBND As Boolean, fixBND As Boolean, StrBnd As Variant, SttQry As Boolean, StrQuerry As String, KeyField As String, editAble As Boolean)

Dim a As Long
Dim NewCol As dxGridColumn
Dim i As Integer

'On Error GoTo ErrorLabel

Nm_DxDbGrid.Dataset.Active = False
If SttBND Then
'------------------
'== BIND SETTING ==
'------------------
    Nm_DxDbGrid.Options.Set (egoShowBands)
    Nm_DxDbGrid.Options.Set (egoSmartRefresh)
    Nm_DxDbGrid.Options.Set (egoSmartReload)
    Nm_DxDbGrid.Options.Unset (egoShowIndicator)
    Nm_DxDbGrid.Options.Set (egoAutoWidth)
    Nm_DxDbGrid.Options.Set (egoRowSelect)
    With Nm_DxDbGrid.Bands   'bands adjusting
      .Add
      .Item(0).Caption = StrBnd(0)
      '.Item(1).Caption = StrBnd(1)
      .BandFont.Bold = True
      .BandColor = &HC0C0C0
      If fixBND Then .Item(0).Fixed = bfLeft
    End With

    With Nm_DxDbGrid.Bands
        If .Count <> 2 Then
           .Remove (.Count - 1)
        End If
    End With
End If


With Nm_DxDbGrid.Columns
'-------------------
'== COLUMNS ARRAY ==
'-------------------

.DestroyColumns
'On Error Resume Next
    For a = 0 To UBound(AryColomn)
        Set NewCol = .Add(AryColomn(a)(3))
        NewCol.ObjectName = AryColomn(a)(0)
        NewCol.Caption = AryColomn(a)(1) '
        NewCol.FieldName = AryColomn(a)(2) 'isi nama tabel
        NewCol.RowIndex = AryColomn(a)(4) 'isi tabel record
        NewCol.BandIndex = AryColomn(a)(5) 'index tabel
        NewCol.Width = AryColomn(a)(6)
        '-header-
        NewCol.HeaderAlignment = taCenter
        '-Disable Editor-
        'NewCol.DisableEditor = False
        'MsgBox AryColomn(3)(4)
         If NewCol.ColumnType = gedCheckEdit Then
            NewCol.CheckColumn.ValueChecked = 1
            NewCol.CheckColumn.ValueUnchecked = 0
        End If
        

        
        If AryColomn(a)(7) = 1 Then
            NewCol.DisableEditor = True
            NewCol.Color = &HC0E0FF
        End If
        '-Alignment-
        If AryColomn(a)(8) = 1 Then
                NewCol.Alignment = taLeftJustify
            ElseIf AryColomn(a)(8) = 2 Then
                NewCol.Alignment = taCenter
            ElseIf AryColomn(a)(8) = 3 Then
                NewCol.Alignment = taRightJustify
        End If
        '-Desimal-
        If AryColomn(a)(9) = 1 Then
            NewCol.Visible = False
            
        End If
                '-FORMAT TIME OR TANGGAL -
        If AryColomn(a)(10) = 1 Then
            NewCol.DisplayFormat = "HH:mm:ss"
        ElseIf AryColomn(a)(10) = 2 Then
            NewCol.DisplayFormat = "yyyy-mm-dd hh:mm:ss"
        ElseIf AryColomn(a)(10) = 3 Then
            NewCol.DisplayFormat = "yyyy-mm-dd"
        ElseIf AryColomn(a)(10) = 4 Then
            NewCol.DisplayFormat = "yyyy"
        ElseIf AryColomn(a)(10) = 5 Then
            NewCol.DisplayFormat = "DD"
        ElseIf AryColomn(a)(10) = 6 Then
            NewCol.DisplayFormat = "#.###"
        End If
   
        
    Next

    'MsgBox UBound(AryColomn)
End With

'--------------------------------------
'--------- TABEL/RECORD FIRST ---------
'--------------------------------------
With Nm_DxDbGrid
    .KeyField = KeyField
    .Dataset.ADODataset.ConnectionString = StrCon
    .Dataset.ADODataset.CursorLocation = clUseClient
    If editAble Then .Dataset.ADODataset.LockType = ltBatchOptimistic Else .Dataset.ADODataset.LockType = ltOptimistic
    
    If SttQry = False Then
    '------------------
    '== Normal Query ==
    '------------------
    .Dataset.ADODataset.CommandType = cmdText
    Else
    '-------------------------
    '== PLSQL cmdStoredProc ==
    '-------------------------
    .Dataset.ADODataset.CommandType = cmdStoredProc
    End If
    .Dataset.ADODataset.CommandText = StrQuerry
    '.Dataset.Active = True
End With

'Exit Sub
'ErrorLabel:
'    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"
End Sub
