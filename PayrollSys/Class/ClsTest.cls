VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Sub FillDxGrid(Nm_DxDbGrid As dxDBGrid, AryColomn As Variant, SttBND As Boolean, fixBND As Boolean, StrBnd As Variant, SttQry As Boolean, StrQuerry As String, keyField As String)
Dim a As Long
Dim NewCol As dxGridColumn
Dim i As Integer
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
        If NewCol.Index = 2 Then
            NewCol.CheckColumn.ValueChecked = 0
            NewCol.CheckColumn.ValueUnchecked = 1
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
            NewCol.DecimalPlaces = 2
        End If
        
        
    Next

    'MsgBox UBound(AryColomn)
End With

'--------------------------------------
'--------- TABEL/RECORD FIRST ---------
'--------------------------------------
With Nm_DxDbGrid
    .keyField = keyField
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
End With

End Sub

'------------GRID RO---------------
Public Sub GetBaseGridRO(Nm_DxDbGrid As dxDBGrid, AryColomn As Variant, StrQuerry As String, keyField As String)
Dim a As Long
Dim NewCol As dxGridColumn
Dim i As Integer
    
    'set grid options
    Nm_DxDbGrid.Dataset.Active = False
    Nm_DxDbGrid.Options.Unset (egoShowIndicator)
    Nm_DxDbGrid.Options.Set (egoAutoWidth)
    'Nm_DxDbGrid.Options.Set (egoed)


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
        
  
        If NewCol.Index = 3 Then
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

        NewCol.HeaderAlignment = taCenter
    Next
    
    
    
'On Error Resume Next
'Dim strSql As String
'strSql = "SELECT BRG_NM FROM barang"
'MsgBox strSQL

'Dim oRS As ADODB.Recordset  'baca data
'  Set oRS = New ADODB.Recordset
'  oRS.Open strSql, conMain, adOpenForwardOnly, adLockReadOnly, adCmdText
'      Do While Not oRS.EOF      'dengan ItemData
'         .ColumnByFieldName("BRG_NM").PickColumn.Items.Add (oRS.Fields("BRG_NM").Value)
'        oRS.MoveNext
'      Loop'

  'oRS.Close                 'Tidy up
'  Set oRS = Nothing

 .ColumnByFieldName("RO_QTY").SpinColumn.MinValue = 1
 .ColumnByFieldName("RO_QTY").SpinColumn.MaxValue = 1000000
' .ColumnByFieldName("NO").Visible = 0

End With

'--------------------------------------
'--------- TABEL/RECORD FIRST ---------
'--------------------------------------
With Nm_DxDbGrid
    .keyField = keyField
    .Dataset.ADODataset.ConnectionString = StrCon
    .Dataset.ADODataset.CursorLocation = clUseClient
    .Dataset.ADODataset.LockType = ltOptimistic
    .Dataset.ADODataset.CommandText = StrQuerry
    .Dataset.Active = True

For i = 1 To .Count
 '.Dataset.SetFields ' .FieldValues("NO") = "PP"
Next i

End With
End Sub

Public Sub SetNewGrid(Nm_DxDbGrid As dxDBGrid, AryColomn As Variant, StrQuerry As String, IsCmdTxt As Boolean, RowIndex As Boolean, keyField As String)
Dim i, a As Integer
Dim NewCol As dxGridColumn
'Dim StrQuerry As String

'Nm_DxDbGrid.DatasetType = dtADODataset
  Nm_DxDbGrid.Dataset.Active = False
With Nm_DxDbGrid.Columns
.DestroyColumns

 For a = 0 To UBound(AryColomn)
        Set NewCol = .Add(AryColomn(a)(3))
        NewCol.ObjectName = AryColomn(a)(0)
        NewCol.Caption = AryColomn(a)(1) '
        NewCol.FieldName = AryColomn(a)(2) 'isi nama tabel
        NewCol.RowIndex = AryColomn(a)(4) 'isi tabel record
        NewCol.BandIndex = AryColomn(a)(5) 'index tabel
        NewCol.Width = AryColomn(a)(6)
        NewCol.Visible = AryColomn(a)(9)
        
'Nm_DxDbGrid.Dataset.Lookup "BRG_NM", 10, "ProductID;FirstName"
    'MsgBox NewCol.ColumnType
   
    
     If NewCol.ColumnType = 11 Then
      With Nm_DxDbGrid.Columns.ColumnByFieldName("BRG_NM").LookupColumn
        .LookupKeyField = "BRG_NM"
        'Set the field name which values will be displayed within column cells
        .LookupResultField = "BRG_NM"
        'Set the fields that will be displayed within the dropdown
        .ListFieldName = "BRG_NM"
        'Set column captions and widths
        .ListColumns = "*:50;Description"
        End With
    End If
    
        'chek
        If NewCol.Index = 5 Then
            NewCol.CheckColumn.ValueChecked = 1
            NewCol.CheckColumn.ValueUnchecked = 0
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
 Next

 .ColumnByFieldName("NO").Visible = RowIndex
 .ColumnByFieldName("RO_QTY").SpinColumn.MinValue = 1
 .ColumnByFieldName("RO_QTY").SpinColumn.MaxValue = 1000000
 
 
 End With
      
 
'GoTo Last
 With Nm_DxDbGrid
    .keyField = "BRG_ID"
        If IsCmdTxt Then .Dataset.ADODataset.CommandType = cmdText Else .Dataset.ADODataset.CommandType = cmdStoredProc
    .Dataset.ADODataset.ConnectionString = StrCon
    .Dataset.ADODataset.CursorLocation = clUseClient
    .Dataset.ADODataset.LockType = ltOptimistic
    .Dataset.ADODataset.CommandText = StrQuerry
    .Dataset.Active = True
 End With
Last:
End Sub

Public Sub FilGridCombo(Nm_DxDbGrid As dxDBGrid, strSql As String, dbField As String, clmField As String)
On Error Resume Next
Dim oRS As ADODB.Recordset  'baca data
  Set oRS = New ADODB.Recordset
  oRS.Open strSql, conMain, adOpenForwardOnly, adLockReadOnly, adCmdText
      Do While Not oRS.EOF      'dengan ItemData
       ' MsgBox oRS.Fields("BRG_NM").Value
         Nm_DxDbGrid.Columns.ColumnByFieldName("BRG_NM").PickColumn.Items.Add (oRS.Fields("BRG_NM").Value)
        oRS.MoveNext
      Loop

  'oRS.Close                 'Tidy up
  Set oRS = Nothing
End Sub
 Public Sub LookupGrid(Nm_DxDbGrid As dxDBGrid)
Dim CString As String, i As Byte
Nm_DxDbGrid.Dataset.Active = False
  CString = conMain
      Nm_DxDbGrid.Columns.ColumnByFieldName("BRG_NM").ColumnType = gedLookupEdit
'For i = 0 To Nm_DxDbGrid.Columns.Count - 1
   'If Nm_DxDbGrid.Columns(i).ColumnType = gedLookupEdit Then
     With Nm_DxDbGrid.Columns.ColumnByFieldName("BRG_NM").LookupColumn
       .LookupDataset.ADODataset.ConnectionString = CString
       
         .ListFieldName = "BRG_NM;BRG_ID"
         .ListWidth = 300
         .ListAutoWidth = True
         MsgBox "ee"
        'Case "COLUMNPART"
        ' Nm_DxDbGrid.Columns(i).ColumnType = 11
        ' .ListFieldName = "Description;OnHand;OnOrder;VendorName"
        ' .ListWidth = 450
        ' .ListAutoWidth = True
         Nm_DxDbGrid.Columns.ColumnByFieldName("BRG_NM").DropDownColumn.DropDownRows = 15
        'Case "COLUMNVENDOR"
        ' .ListFieldName = "VendorName;PartNo"
       '  .ListWidth = 250
       '  .ListAutoWidth = True
       '  Nm_DxDbGrid.Columns(i).DropDownColumn.DropDownRows = 10
       
     End With
   'End If
 ' Next
  Nm_DxDbGrid.Dataset.ADODataset.ConnectionString = CString
 Nm_DxDbGrid.Dataset.Active = True
 
  'MsgBox "File " & dbName & " is not found", vbCritical

 End Sub
 
 Public Sub LkUpGrid(Nm_DxDbGrid As dxDBGrid, clmField As String, keyField As String, ResField As String, LsField As String)
Nm_DxDbGrid.Columns.ColumnByFieldName(clmField).ColumnType = gedLookupEdit

With Nm_DxDbGrid.Columns.ColumnByFieldName(clmField).LookupColumn
  'Set the dataset properties
  .LookupDatasetType = dtADODataset
  With .LookupDataset.ADODataset
    .ConnectionString = conMain
    .CommandType = cmdText
    .CommandText = "select BRG_NM from barang"
  End With
  .LookupDataset.Open
    'Nm_DxDbGrid.Dataset.Lookup "BRG_NM", "10", "ProductID;FirstName"

  'Set the field name which will be used for search
 .LookupKeyField = keyField
  'Set the field name which values will be displayed within column cells
  '.LookupResultField = ResField
  'Set the fields that will be displayed within the dropdown
  '.ListFieldName = keyField & ";" & ResField
  'Set column captions and widths
  .ListColumns = LsField

  Nm_DxDbGrid.Dataset.Open

End With
 End Sub
