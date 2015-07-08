VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsReportExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim oExcel As Excel.Application
Dim oBook As Excel.Workbook
Dim oSheet As Excel.Worksheet
Dim oRS, oRSitem As ADODB.Recordset 'baca data
Dim strSql As String
Dim strSQLitem As String
Dim i%, Curr_Row%
Public Sub Po2Excel(xPOID As String) 'RepCorp As String, RepCab As String, TglPo As String, CheckOnly As Boolean, PicPath As String)
'============================================
'|              LG PROJECT     [23 May 2014]|
'============================================
'|ClsReportExcel - Po2Excel                 |
'|(digunakan untuk menampilkan report PO)   |
'|Progress = 100%                           |
'--------------------------------------------
'|By. FARINDRA                  @01-07-2014 |
'============================================
Dim RoId As String
Dim SetEks, SetCorp, Setcab, SetTgl As String
Dim Col(10), Row(10) As Integer
Dim curRow, DevSheet, i, K, PoCount, PoPPN As Integer
Dim oRS As New ADODB.Recordset
Dim oRSitem As New ADODB.Recordset

On Error GoTo ErrHandReport
'cek jika file template ada "template_report_po.xlsx"
If CekStatusFile("template_report_po.xlsx", App.Path & "\Report\") Then
    File_Copy App.Path & "\Report\", "template_report_po.xlsx", App.Path & "\Report\PO\", xPOID & ".xlsx"  ' "PO_" & Format(Now, "YYYYMMDD") & ".xlsx"

    'buka file excel tersebut
    Set oExcel = CreateObject("Excel.Application"):  oExcel.Visible = False '= ONOFF ==
    Set oBook = oExcel.Workbooks.Open(App.Path & "\Report\PO\" & xPOID & ".xlsx") 'PO_" & Format(Now, "YYYYMMDD") & ".xlsx")
    DevSheet = 1
    Set oSheet = oBook.Worksheets(DevSheet)
    'Goto Dataprose
Else
    'jika tidak ada

'buat file excel baru jika tidak ada file template "template_report_po.xlsx"
Set oExcel = CreateObject("Excel.Application"): oExcel.Visible = False '= ONOFF ==
Set oBook = oExcel.Workbooks.Add:


'<FILL PO DESIGN>

'DEFAULT PROPERTI CELL
    'oSheet.Columns.AutoFit
DevSheet = 1
Set oSheet = oBook.Worksheets(DevSheet)
With oSheet
    .Cells.Font.Name = "Arial"
    .Cells.Font.Size = 11
    .Cells.ColumnWidth = 4.15
    .PageSetup.Zoom = 65
    .PageSetup.PaperSize = xlPaperA4
    
 '[DESIGN PO - JUDUL]
    'TEXT
    .Cells(1, 1) = "PURCHASE ORDER"
    
    Row(1) = 1 'baris 1
    Row(2) = 1 'baris 1
    Col(1) = 1 'kolom A
    Col(2) = 28 'kolom AB
    'MARGE
        .Range(.Cells(Row(1), Col(1)), .Cells(Row(2), Col(2))).MergeCells = True
    'FONT & ALIGMENT
        .Range(.Cells(Row(1), Col(1)), .Cells(Row(2), Col(2))).Font.Name = "Arial"
        .Range(.Cells(Row(1), Col(1)), .Cells(Row(2), Col(2))).Font.Size = 18
        .Range(.Cells(Row(1), Col(1)), .Cells(Row(2), Col(2))).Font.Bold = True
        .Range(.Cells(Row(1), Col(1)), .Cells(Row(2), Col(2))).HorizontalAlignment = 3
 
 '[DESIGN PO - BODY]
    'TEXT HEADER
    .Cells(3, 1) = "Messrs." 'A3
    .Cells(8, 1) = "Attn"   'A7
    .Cells(9, 1) = "Tlp"    'A8
    .Cells(10, 1) = "Mail"   'A9
    .Cells(8, 2) = ":"
    .Cells(9, 2) = ":"
    .Cells(10, 2) = ":"

    
    .Cells(3, 20 - 1) = "Date"               'S3
    .Cells(4, 20 - 1) = "Order No"           'S4
    .Cells(5, 20 - 1) = "Seller's Ref. No."  'S5
    .Cells(6, 20 - 1) = "Terms of Sales"      'S6
    .Cells(7, 20 - 1) = "Latest Delivery"     'S7
    .Cells(3, 24 - 1) = ":"
    .Cells(4, 24 - 1) = ":"
    .Cells(5, 24 - 1) = ":"
    .Cells(6, 24 - 1) = ":"
    .Cells(7, 24 - 1) = ":"
    
    '.Cells(12, 1) = "Acceptance of this order and the attachment is subject to the terms and conditions on the back here of" 'A12
    '.Cells(13, 1) = "which are incorporated and made a part here of :"    'A13
        
    .Cells(15, 1) = "No."
    .Cells(15, 2) = "Description"
    .Cells(15, 17) = "Unit Price"
    .Cells(15, 21) = "Quantity"
    .Cells(15, 25) = "Total"
    

    'TEXT FOOTER
    .Cells(53, 2) = "Said Total:"
    .Cells(52, 17) = "Include PPN"
    .Cells(54, 17) = "GRAND TOTAL"
    .Cells(56, 1) = "Delivery  Address :"
    .Cells(61, 1) = "Tlp"
    .Cells(62, 1) = "Fax"
    .Cells(63, 1) = "CP"
    .Cells(65, 1) = "General Notes:"
    .Cells(56, 15) = "Note:"
    .Cells(62, 15) = "Cara pembayaran:"
    
    For K = 1 To 4
        If K <= 2 Then .Cells(60 + K, 7) = " / "
        If K <= 3 Then .Cells(60 + K, 2) = ":"
        .Cells(65 + K, 1) = "=" & Chr(34) & Str(K) & "." & Chr(34)
        .Cells(56 + K, 15) = "=" & Chr(34) & Str(K) & "." & Chr(34)
        .Cells(62 + K, 15) = "=" & Chr(34) & Str(K) & "." & Chr(34)
    Next K

    'MARGE & ALIGN & FONT
    For i = 1 To 54
            'MsgBox i
        If i <= 2 Then
            .Range(.Cells(11 + i, 1), .Cells(11 + i, 28)).MergeCells = True 'A12-A13:AB12-AB13
            .Range(.Cells(76 + i, 9), .Cells(76 + i, 13)).MergeCells = True 'I77-I78:M77-M78
            .Range(.Cells(76 + i, 14), .Cells(76 + i, 18)).MergeCells = True 'N77-N78:R77-R78
            .Range(.Cells(76 + i, 19), .Cells(76 + i, 23)).MergeCells = True 'O77-O78:W77-W78
            '.Range(.Cells(52 + i, 2), .Cells(52 + i, 16)).MergeCells = True 'B53-B54:P53-P54
        End If
        If i <= 4 Then
            .Range(.Cells(2 + i, 1), .Cells(2 + i, 18)).MergeCells = True   'A3-A6:R3-R6
            .Range(.Cells(2 + i, 24), .Cells(2 + i, 24)).HorizontalAlignment = xlRight
            .Range(.Cells(7 + i, 3), .Cells(7 + i, 18)).MergeCells = True   'C8-C11:R8-R11
            .Range(.Cells(2 + i, 19), .Cells(2 + i, 22)).MergeCells = True   'T3-T6:W3-W6
            .Range(.Cells(2 + i, 24), .Cells(2 + i, 28)).MergeCells = True   'Y3-Y6:AB3-AB6
            .Range(.Cells(55 + i, 1), .Cells(55 + i, 13)).MergeCells = True   'A56-A59:M56-M59
            .Range(.Cells(65 + i, 2), .Cells(65 + i, 13)).MergeCells = True   'B66-B69:M66-M69
            .Range(.Cells(60 + i, 3), .Cells(60 + i, 6)).MergeCells = True   'C61-C64:F61-F64
            .Range(.Cells(60 + i, 8), .Cells(60 + i, 11)).MergeCells = True   'H61-H64:K61-K64
            
            .Range(.Cells(56 + i, 16), .Cells(56 + i, 28)).MergeCells = True  'H61-H64:K61-K64
            .Range(.Cells(62 + i, 16), .Cells(62 + i, 28)).MergeCells = True  'H61-H64:K61-K64

        End If
        If i = 7 Then
            .Range(.Cells(i, 19), .Cells(i, 22)).MergeCells = True  'T7:W7
            .Range(.Cells(i, 24), .Cells(i, 28)).MergeCells = True  'Y7:AB7
            .Range(.Cells(i, 1), .Cells(i, 18)).MergeCells = True  'A7:R7
        End If
        If i >= 16 Then
            .Range(.Cells(i, 2), .Cells(i, 16)).MergeCells = True  'B6-B36:P6-P39 --> Description
            .Range(.Cells(i, 17), .Cells(i, 20)).MergeCells = True  'Q6-Q36:T6-T39 --> Unit Price
            .Range(.Cells(i, 21), .Cells(i, 22)).MergeCells = True  'U6-U36:V6-V39 --> Quantity
            .Range(.Cells(i, 23), .Cells(i, 24)).MergeCells = True  'W6-W36:X6-X39 --> Quantity
            .Range(.Cells(i, 25), .Cells(i, 28)).MergeCells = True  'Y6-Y36:AB6-AB39 --> Total
            .Cells(i, 25) = "=IF(A" & i & ">=1,Q" & i & "*U" & i & "," & Chr(34) & Chr(34) & ")" '=IF(V16>=1,Q16*U16,"") --> rumus total
            
        End If
    
        
    Next
        'KOLOM HEADER
        .Range(.Cells(15, 2), .Cells(15, 16)).MergeCells = True ' DESCRIPTION
        .Range(.Cells(15, 17), .Cells(15, 20)).MergeCells = True ' UNIT PRICE
        .Range(.Cells(15, 21), .Cells(15, 24)).MergeCells = True 'QTY
        .Range(.Cells(15, 25), .Cells(15, 28)).MergeCells = True 'TOTAL
        .Range(.Cells(15, 2), .Cells(15, 28)).HorizontalAlignment = 3
        
        'FONT BOLD / ITALIC / ALIGN
        .Range(.Cells(3, 1), .Cells(11, 2)).Font.Bold = True
        .Range(.Cells(4, 1), .Cells(7, 1)).Font.Bold = False
        .Range(.Cells(3, 1), .Cells(11, 2)).HorizontalAlignment = 1
        .Range(.Cells(3, 19), .Cells(7, 23)).Font.Bold = True
        
        
        .Range(.Cells(12, 1), .Cells(13, 2)).Font.Italic = True
        .Range(.Cells(15, 1), .Cells(15, 28)).Font.Bold = True
        .Range(.Cells(52, 1), .Cells(56, 17)).Font.Bold = True
        .Range(.Cells(54, 2), .Cells(54, 2)).Font.Italic = True 'SAID PAID
        .Range(.Cells(54, 2), .Cells(54, 2)).Font.Bold = False
        
        .Range(.Cells(62, 15), .Cells(62, 15)).Font.Bold = True
        .Range(.Cells(61, 1), .Cells(65, 1)).Font.Bold = True
        .Range(.Cells(77, 9), .Cells(77, 19)).Font.Bold = True
        .Range(.Cells(77, 9), .Cells(78, 19)).HorizontalAlignment = 3
        .Range(.Cells(77, 9), .Cells(78, 19)).Font.Underline = True
        
        'FORMAT CELL
       .Range(.Cells(16, 17), .Cells(54, 25)).NumberFormat = "#,##0" '=Y16:AB54 TOTAL FORMAT
       .Range(.Cells(53, 21), .Cells(53, 21)).NumberFormat = "00%"  '=U53 PPN FORMAT
       
       .Cells(52, 21) = "=SUM(U16:V51)" 'TOTAL QTY
       .Cells(52, 25) = "=SUM(Y16:AB51)" 'TOTAL HARGA
       .Cells(53, 25) = "=U53*Y52" 'TOTAL PPN
        .Range(.Cells(53, 25), .Cells(53, 25)).ColumnWidth = "4.14"
       .Cells(54, 25) = "=Y52+Y53" 'GRAND TOTAL

    'BORDER
    .Range(.Cells(15, 1), .Cells(15, 28)).BorderAround ColorIndex:=1, Weight:=xlMedium 'NO
    .Range(.Cells(15, 1), .Cells(51, 1)).BorderAround ColorIndex:=1, Weight:=xlMedium 'DESCRIPTION
    .Range(.Cells(15, 2), .Cells(51, 16)).BorderAround ColorIndex:=1, Weight:=xlMedium 'UNIT PRICE
    .Range(.Cells(15, 17), .Cells(51, 20)).BorderAround ColorIndex:=1, Weight:=xlMedium
    .Range(.Cells(15, 21), .Cells(51, 24)).BorderAround ColorIndex:=1, Weight:=xlMedium 'QUANTITY
        .Range(.Cells(52, 21), .Cells(52, 24)).BorderAround ColorIndex:=1, Weight:=xlMedium
        .Range(.Cells(16, 23), .Cells(51, 23)).Borders(xlEdgeLeft).LineStyle = xlDash
    .Range(.Cells(15, 25), .Cells(51, 28)).BorderAround ColorIndex:=1, Weight:=xlMedium 'TOTAL
        .Range(.Cells(52, 25), .Cells(54, 28)).BorderAround ColorIndex:=1, Weight:=xlMedium
    .Range(.Cells(56, 15), .Cells(67, 28)).BorderAround ColorIndex:=1, Weight:=xlThin 'HEADER
    End With
'</FILL PO DESIGN>
End If

'mengambil attribut
'---If CheckOnly Then SetEks = "'1'" Else SetEks = "'0'"
'---If RepCorp = "" Then SetCorp = "" Else SetCorp = " AND B.CORP_ID='" & RepCorp & "'"
'---If RepCab = "" Then Setcab = "" Else Setcab = " AND B.CAB_ID='" & RepCab & "'"
'---SetTgl = " AND B." & TglPo



'menghitung jumlah List PO
PoCount = 1
'---PoCount = GetDtaTbl(conMain, "SELECT COUNT(DISTINCT B.PO_ID) AS PoCount FROM tab_po_header B WHERE B.PO_STS='0'" & SetCorp & Setcab & SetTgl & " AND B.EXPORT=" & SetEks, "PoCount")
'MsgBox PoCount
'<FILL DATA PO>
On Error Resume Next
    Curr_Row = 1
    strSql = "SELECT PO_ID FROM tab_po_header WHERE PO_ID='" & xPOID & "'" '"SELECT B.PO_ID FROM tab_po_header B WHERE B.PO_STS='0'" & SetCorp & Setcab & SetTgl & " AND B.EXPORT=" & SetEks
    'MsgBox strSql
    'Clipboard.SetText strSql
    Set oRS = ExecuteRecordSetMySql(strSql, conMain)
    With oRS
         .MoveFirst
        Do While Not .EOF
            Set oSheet = oBook.Worksheets(Curr_Row)
            RoId = .Fields("PO_ID").Value
             'MsgBox RoId
            oSheet.Name = RoId 'NAMA SHEET
            
            'Copy SHEET TEMPLATE
            If Curr_Row <> PoCount Then
                oSheet.Copy After:=oSheet
                Curr_Row = Curr_Row + 1
            End If
            
            '[HEADER-DATA]

            'Messrs. [A4-A10]
            oSheet.Cells(4, 1) = GetDtaTbl(conMain, "SELECT SPL_NM FROM tab_po_detail_export  WHERE PO_ID='" & RoId & "'", "SPL_NM")
            oSheet.Cells(5, 1) = GetDtaTbl(conMain, "SELECT SPL_ALM_NM FROM tab_po_detail_export  WHERE PO_ID='" & RoId & "'", "SPL_ALM_NM")
            oSheet.Cells(6, 1) = GetDtaTbl(conMain, "SELECT SPL_ALM_JLN  FROM tab_po_detail_export  WHERE PO_ID='" & RoId & "'", "SPL_ALM_JLN")
            oSheet.Cells(7, 1) = GetDtaTbl(conMain, "SELECT SPL_ALM_CTY  FROM tab_po_detail_export  WHERE PO_ID='" & RoId & "'", "SPL_ALM_CTY")
            oSheet.Cells(8, 3) = GetDtaTbl(conMain, "SELECT SPL_CP FROM  tab_po_detail_export  WHERE PO_ID='" & RoId & "'", "SPL_CP")
            oSheet.Cells(9, 3) = GetDtaTbl(conMain, "SELECT SPL_TLP1   FROM tab_po_detail_export  WHERE PO_ID='" & RoId & "'", "SPL_TLP1") & " / " & _
                                 GetDtaTbl(conMain, "SELECT SPL_TLP2   FROM tab_po_detail_export  WHERE PO_ID='" & RoId & "'", "SPL_TLP2")
            oSheet.Cells(10, 3) = GetDtaTbl(conMain, "SELECT SPL_MAIL1   FROM tab_po_detail_export  WHERE PO_ID='" & RoId & "'", "SPL_MAIL1") & " / " & _
                                  GetDtaTbl(conMain, "SELECT SPL_MAIL2   FROM tab_po_detail_export  WHERE PO_ID='" & RoId & "'", "SPL_MAIL2")
            'HEADER [A12 & A13]
            'oSheet.Cells(12, 1) = GetDtaTbl(conMain, "SELECT A.REP_HDR1 FROM tab_po_detail_report A,tab_po_header B WHERE B.PO_ID='" & RoId & "' AND A.SPL_ID=B.SPL_ID", "REP_HDR1")
            'oSheet.Cells(13, 1) = GetDtaTbl(conMain, "SELECT A.REP_HDR2 FROM tab_po_detail_report A,tab_po_header B WHERE B.PO_ID='" & RoId & "' AND A.SPL_ID=B.SPL_ID", "REP_HDR2")
            
            'LOGO
            oSheet.Shapes.AddPicture ImagePath("EXC_" & GetDtaTbl(conMain, "SELECT CORP_ID FROM tab_po_header WHERE PO_ID='" & RoId & "'", "CORP_ID")), False, True, 600, 30, 100, 50
            
            'DETAIL PO ID
            oSheet.Cells(6, 24) = Format(GetDtaTbl(conMain, "SELECT PO_TGL FROM tab_po_header WHERE PO_ID='" & RoId & "'", "PO_TGL"), "DD, MMM YYYY")
            oSheet.Cells(7, 24) = RoId
            oSheet.Cells(8, 24) = prjSysID.GetName_Karyawan(prjSysID.GetID_karyawan_From_IDUser(GetDtaTbl(conMain, "SELECT USER_ID FROM tab_po_header  WHERE PO_ID='" & RoId & "'", "USER_ID")))
            oSheet.Cells(10, 24) = Format(GetDtaTbl(conMain, "SELECT PO_ETD FROM tab_po_header WHERE PO_ID='" & RoId & "'", "PO_ETD"), "DD, MMM YYYY")
            oSheet.Cells(11, 24) = Format(GetDtaTbl(conMain, "SELECT PO_ETA FROM tab_po_header WHERE PO_ID='" & RoId & "'", "PO_ETA"), "DD, MMM YYYY")
            
            
            '[/HEADER-DATA]
            
            '[BODY-DATA]
            'LIST ITEM
            'strSQLitem = "SELECT B.BRG_NM,B.BRG_HRG,B.PO_QTY,D.BRG_SAT " & _
                         "FROM tab_po_header A,tab_po_detail_suplier B,tab_barang D " & _
                         "WHERE A.PO_ID='" & RoId & "' " & _
                         "AND A.PO_DS_ID=B.PO_DS_ID AND B.BRG_ID=D.BRG_ID " 'AND c.CORP_ID='" & RepCorp & "' AND C.CAB_ID='" & RepCab & "'"
            strSQLitem = "SELECT * " & _
                         "FROM tab_po_detail " & _
                         "WHERE PO_ID='" & RoId & "'"
            'MsgBox strSQLitem
            '''Clipboard.Clear
            ''Clipboard.SetText strSQLitem

            Set oRSitem = ExecuteRecordSetMySql(strSQLitem, conMain)
            curRow = 1
            With oRSitem
                .MoveFirst
                Do While Not .EOF
                    oSheet.Cells(15 + curRow, 1) = curRow
                    oSheet.Cells(15 + curRow, 2) = .Fields("BRG_NM").Value
                    oSheet.Cells(15 + curRow, 17) = .Fields("BRG_HRG").Value
                    oSheet.Cells(15 + curRow, 21) = .Fields("BRG_QTY").Value
                    oSheet.Cells(15 + curRow, 23) = .Fields("BRG_SAT").Value
                .MoveNext
                curRow = curRow + 1
                Loop
            End With
          .MoveNext
          
          '[/BODY-DATA]
          
          '[FOOTER-DATA]
            'dim Ppn as decimal
            'PPN/discount/ongkos
            ' oSheet.Cells(53, 21) = GetDtaTbl(conMain, "SELECT A.REP_PPN FROM tab_po_detail_report A,tab_po_header B WHERE B.PO_ID='" & RoId & "' AND A.SPL_ID=B.SPL_ID", "REP_PPN") / 100 'PPN
            oSheet.Cells(51, 25) = GetDtaTbl(conMain, "SELECT PO_DISCOUNT FROM tab_po_header WHERE PO_ID='" & RoId & "'", "PO_DISCOUNT")
            oSheet.Cells(52, 21) = GetDtaTbl(conMain, "SELECT PO_PPN FROM tab_po_header WHERE PO_ID='" & RoId & "'", "PO_PPN") / 100 'PPN
            oSheet.Cells(53, 25) = GetDtaTbl(conMain, "SELECT PO_ONGKOS FROM tab_po_header WHERE PO_ID='" & RoId & "'", "PO_ONGKOS")
            
            'SAID TOTAL
             oSheet.Cells(54, 2) = prjSysID.PaidSaid(Replace(oSheet.Cells(54, 25), ",", "")) & " (" & GetDtaTbl(conMain, "SELECT PO_CUR FROM tab_po_header WHERE PO_ID='" & RoId & "'", "PO_CUR") & ")"
             oSheet.Cells(54, 21) = "TOTAL " & "(" & GetDtaTbl(conMain, "SELECT PO_CUR FROM tab_po_header WHERE PO_ID='" & RoId & "'", "PO_CUR") & ")"
            
            'SHIPPING TLP & FAX
            oSheet.Cells(61, 3) = GetDtaTbl(conMain, "SELECT SHIP_TLP1 FROM tab_po_detail_export  WHERE PO_ID='" & RoId & "'", "SHIP_TLP1") 'TELP1
                oSheet.Cells(61, 8) = GetDtaTbl(conMain, "SELECT SHIP_TLP2 FROM tab_po_detail_export  WHERE PO_ID='" & RoId & "'", "SHIP_TLP2")  'TELP2
            oSheet.Cells(62, 3) = GetDtaTbl(conMain, "SELECT SHIP_FAX FROM tab_po_detail_export  WHERE PO_ID='" & RoId & "'", "SHIP_FAX")   'FAX1
            
            'SHIPPING Address
            Clipboard.Clear
            Clipboard.SetText GetDtaTbl(conMain, "SELECT SHIP_ADR FROM tab_po_detail_export  WHERE PO_ID='" & RoId & "'", "SHIP_ADR")  'Jam
            oSheet.Range(oSheet.Cells(57, 1), oSheet.Cells(57, 1)).PasteSpecial Paste:=xlPasteAll ', Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            ''Clipboard.GetText
            'oSheet.Cells(57, 1) = GetDtaTbl(conMain, "SELECT SHIP_ADR FROM tab_po_detail_export  WHERE PO_ID='" & RoId & "'", "SHIP_ADR") 'DELIVER ADDRESS NAMA TEMPAT
            'oSheet.Cells(58, 1) = GetDtaTbl(conMain, "SELECT A.CAB_ALM_JLN FROM tab_cabang A, tab_po_header B  WHERE B.PO_ID='" & RoId & "' AND B.PO_SHIP=A.CAB_ID", "CAB_ALM_JLN") 'DELIVER ADDRESS NAMA JALAN
            'oSheet.Cells(59, 1) = GetDtaTbl(conMain, "SELECT A.CAB_ALM_CTY FROM tab_cabang A, tab_po_header B  WHERE B.PO_ID='" & RoId & "' AND B.PO_SHIP=A.CAB_ID", "CAB_ALM_CTY")  'DELIVER ADDRESS NAMA KOTA

            'BILLING TLP & FAX
            oSheet.Cells(61, 17) = GetDtaTbl(conMain, "SELECT BILL_TLP1 FROM tab_po_detail_export  WHERE PO_ID='" & RoId & "'", "BILL_TLP1")  'TELP1
                oSheet.Cells(61, 22) = GetDtaTbl(conMain, "SELECT BILL_TLP2 FROM tab_po_detail_export  WHERE PO_ID='" & RoId & "'", "BILL_TLP2")  'TELP2
            oSheet.Cells(62, 17) = GetDtaTbl(conMain, "SELECT BILL_FAX FROM tab_po_detail_export  WHERE PO_ID='" & RoId & "'", "BILL_FAX")   'FAX1
            
            'BILLING Address
            Clipboard.Clear
            Clipboard.SetText GetDtaTbl(conMain, "SELECT BILL_ADR FROM tab_po_detail_export  WHERE PO_ID='" & RoId & "'", "BILL_ADR") 'DELIVER ADDRESS NAMA TEMPAT
            oSheet.Range(oSheet.Cells(57, 15), oSheet.Cells(57, 15)).PasteSpecial Paste:=xlPasteAll
            'oSheet.Cells(58, 15) = GetDtaTbl(conMain, "SELECT A.CAB_ALM_JLN FROM tab_cabang A, tab_po_header B  WHERE B.PO_ID='" & RoId & "' AND B.PO_BILL=A.CAB_ID", "CAB_ALM_JLN") 'DELIVER ADDRESS NAMA JALAN
            'oSheet.Cells(59, 15) = GetDtaTbl(conMain, "SELECT A.CAB_ALM_CTY FROM tab_cabang A, tab_po_header B  WHERE B.PO_ID='" & RoId & "' AND B.PO_BILL=A.CAB_ID", "CAB_ALM_CTY")  'DELIVER ADDRESS NAMA KOTA

            'NOTE
            oSheet.Cells(66, 1) = GetDtaTbl(conMain, "SELECT PO_NOT FROM tab_po_header  WHERE PO_ID='" & RoId & "'", "PO_NOT") 'NOTE 1
            'oSheet.Cells(58, 16) = GetDtaTbl(conMain, "SELECT A.REP_NOT2 FROM tab_po_detail_report A,tab_po_header B WHERE B.PO_ID='" & RoId & "' AND A.SPL_ID=B.SPL_ID", "REP_NOT2") 'NOTE 2
            'oSheet.Cells(59, 16) = GetDtaTbl(conMain, "SELECT A.REP_NOT3 FROM tab_po_detail_report A,tab_po_header B WHERE B.PO_ID='" & RoId & "' AND A.SPL_ID=B.SPL_ID", "REP_NOT3") 'NOTE 3
            'oSheet.Cells(60, 16) = GetDtaTbl(conMain, "SELECT A.REP_NOT4 FROM tab_po_detail_report A,tab_po_header B WHERE B.PO_ID='" & RoId & "' AND A.SPL_ID=B.SPL_ID", "REP_NOT4") 'NOTE 4

            'CARA PEMBAYARAN
            'oSheet.Cells(63, 16) = GetDtaTbl(conMain, "SELECT A.REP_PAY1 FROM tab_po_detail_report A,tab_po_header B WHERE B.PO_ID='" & RoId & "' AND A.SPL_ID=B.SPL_ID", "REP_PAY1") 'CARA PEMBAYARAN 1
            'oSheet.Cells(64, 16) = GetDtaTbl(conMain, "SELECT A.REP_PAY2 FROM tab_po_detail_report A,tab_po_header B WHERE B.PO_ID='" & RoId & "' AND A.SPL_ID=B.SPL_ID", "REP_PAY2") 'CARA PEMBAYARAN 2
            'oSheet.Cells(65, 16) = GetDtaTbl(conMain, "SELECT A.REP_PAY3 FROM tab_po_detail_report A,tab_po_header B WHERE B.PO_ID='" & RoId & "' AND A.SPL_ID=B.SPL_ID", "REP_PAY3") 'CARA PEMBAYARAN 3
            'oSheet.Cells(66, 16) = GetDtaTbl(conMain, "SELECT A.REP_PAY4 FROM tab_po_detail_report A,tab_po_header B WHERE B.PO_ID='" & RoId & "' AND A.SPL_ID=B.SPL_ID", "REP_PAY4") 'CARA PEMBAYARAN 4

          
            'GENERAL NOTE
            'oSheet.Cells(66, 2) = GetDtaTbl(conMain, "SELECT A.REP_NOT_GEN_1A FROM tab_po_detail_report A,tab_po_header B WHERE B.PO_ID='" & RoId & "' AND A.SPL_ID=B.SPL_ID", "REP_NOT_GEN_1A") 'GENERAL NOTE 1A
            'oSheet.Cells(67, 2) = GetDtaTbl(conMain, "SELECT A.REP_NOT_GEN_1B FROM tab_po_detail_report A,tab_po_header B WHERE B.PO_ID='" & RoId & "' AND A.SPL_ID=B.SPL_ID", "REP_NOT_GEN_1B") 'GENERAL NOTE 1B
            'oSheet.Cells(68, 2) = GetDtaTbl(conMain, "SELECT A.REP_NOT_GEN_2A FROM tab_po_detail_report A,tab_po_header B WHERE B.PO_ID='" & RoId & "' AND A.SPL_ID=B.SPL_ID", "REP_NOT_GEN_2A") 'GENERAL NOTE 2A
            'oSheet.Cells(69, 2) = GetDtaTbl(conMain, "SELECT A.REP_NOT_GEN_2B FROM tab_po_detail_report A,tab_po_header B WHERE B.PO_ID='" & RoId & "' AND A.SPL_ID=B.SPL_ID", "REP_NOT_GEN_2B") 'GENERAL NOTE 2B

            'SIGN NAMA
            oSheet.Cells(77, 9) = prjSysID.GetName_Karyawan(prjSysID.GetID_karyawan_From_IDUser(GetDtaTbl(conMain, "SELECT USER_ID FROM tab_po_header  WHERE PO_ID='" & RoId & "'", "USER_ID")))     'SIGN NM 1
            'oSheet.Cells(77, 14) = GetDtaTbl(conMain, "SELECT A.REP_SIGN2_NM FROM tab_po_detail_report A,tab_po_header B WHERE B.PO_ID='" & RoId & "' AND A.SPL_ID=B.SPL_ID", "REP_SIGN2_NM") 'SIGN NM 2
            oSheet.Cells(77, 19) = GetDtaTbl(conMain, "SELECT PO_APROVE FROM tab_po_header  WHERE PO_ID='" & RoId & "'", "PO_APROVE") 'SIGN NM 3


            'SIGN JABATAN
            'oSheet.Cells(78, 9) = GetDtaTbl(conMain, FROM tab_po_header  WHERE PO_ID='" & RoId & "'", "PO_NOT") 'SIGN JAB 1
            'oSheet.Cells(78, 14) = GetDtaTbl(conMain, "SELECT A.REP_SIGN2_JAB FROM tab_po_detail_report A,tab_po_header B WHERE B.PO_ID='" & RoId & "' AND A.SPL_ID=B.SPL_ID", "REP_SIGN2_JAB") 'SIGN JAB 2
            'oSheet.Cells(77, 19) = GetDtaTbl(conMain, "SELECT PO_APROVE FROM tab_po_header  WHERE PO_ID='" & RoId & "'", "PO_NOT") 'SIGN JAB 3
        
          '[/FOOTER-DATA]
        Loop
        oSheet.Cells.Select
    End With
    
    
    Set oRS = Nothing
'</FILL DATA PO>

ter:
  ' oBook.SaveAs "D:\PO.xls"
  ' oExcel.Quit
   
   'oRS.MoveFirst
   'oRS.Close
   'MsgBox "Report File format  File Saved", 64, "Info"
  
  Clipboard.Clear
  oExcel.Visible = True '= ONOFF ==

Exit Sub
ErrHandReport:
Select Case Err.Number
    Case Is = 1
        MsgBox "Harap Tutup File PO Excel yang sedang aktif", vbCritical, "Perhatian!"
End Select

 
End Sub
Public Sub RO2Excel(RepCorp As String, RepCab As String, TglRo As String, CheckOnly As Boolean, PicPath As String)
'============================================
'|              LG PROJECT     [23 May 2014]|
'============================================
'|ClsReportExcel - RO2Excel                 |
'|(digunakan untuk menampilkan report RO)   |
'|Progress = 100%                           |
'--------------------------------------------
'|By. FARINDRA                  @01-10-2014 |
'============================================
Dim RoId As String
Dim SetEks, SetCorp, Setcab, SetTgl As String
Dim Col(10), Row(10) As Integer
Dim curRow, DevSheet, i, K, PoCount, PoPPN As Integer
Dim oRS As New ADODB.Recordset
Dim oRSitem As New ADODB.Recordset

On Error GoTo ErrHandReport
'cek jika file template ada "template_report_po.xlsx"
If CekStatusFile("template_report_ro_A4a.xlsx", App.Path & "\Report\") Then
    File_Copy App.Path & "\Report\", "template_report_ro_A4a.xlsx", App.Path & "\Report\RO\", "RO_" & Format(Now, "YYYYMMDD") & ".xlsx"

    'buka file excel tersebut
    Set oExcel = CreateObject("Excel.Application"):  oExcel.Visible = True '= ONOFF ==
    Set oBook = oExcel.Workbooks.Open(App.Path & "\Report\RO\RO_" & Format(Now, "YYYYMMDD") & ".xlsx")
    DevSheet = 1
    Set oSheet = oBook.Worksheets(DevSheet)
    'Goto Dataprose
Else
    'jika tidak ada

'buat file excel baru jika tidak ada file template "template_report_po.xlsx"
Set oExcel = CreateObject("Excel.Application"): oExcel.Visible = True '= ONOFF ==
Set oBook = oExcel.Workbooks.Add


'<FILL RO DESIGN>

'DEFAULT PROPERTI CELL
    'oSheet.Columns.AutoFit
DevSheet = 1
Set oSheet = oBook.Worksheets(DevSheet)
With oSheet
    .Cells.Font.Name = "Arial"
    .Cells.Font.Size = 11
    .Cells.ColumnWidth = 4.15
    .PageSetup.Zoom = 65
    .PageSetup.PaperSize = xlPaperA4
    
 '[DESIGN RO - JUDUL]
    'TEXT
    .Cells(1, 1) = "PURCHASE ORDER"
    
    Row(1) = 1 'baris 1
    Row(2) = 1 'baris 1
    Col(1) = 1 'kolom A
    Col(2) = 28 'kolom AB
    'MARGE
        .Range(.Cells(Row(1), Col(1)), .Cells(Row(2), Col(2))).MergeCells = True
    'FONT & ALIGMENT
        .Range(.Cells(Row(1), Col(1)), .Cells(Row(2), Col(2))).Font.Name = "Arial"
        .Range(.Cells(Row(1), Col(1)), .Cells(Row(2), Col(2))).Font.Size = 18
        .Range(.Cells(Row(1), Col(1)), .Cells(Row(2), Col(2))).Font.Bold = True
        .Range(.Cells(Row(1), Col(1)), .Cells(Row(2), Col(2))).HorizontalAlignment = 3
 
 '[DESIGN PO - BODY]
    'TEXT HEADER
    .Cells(3, 1) = "Messrs." 'A3
    .Cells(8, 1) = "Attn"   'A7
    .Cells(9, 1) = "Tlp"    'A8
    .Cells(10, 1) = "Mail"   'A9
    .Cells(8, 2) = ":"
    .Cells(9, 2) = ":"
    .Cells(10, 2) = ":"

    
    .Cells(3, 20 - 1) = "Date"               'S3
    .Cells(4, 20 - 1) = "Order No"           'S4
    .Cells(5, 20 - 1) = "Seller's Ref. No."  'S5
    .Cells(6, 20 - 1) = "Terms of Sales"      'S6
    .Cells(7, 20 - 1) = "Latest Delivery"     'S7
    .Cells(3, 24 - 1) = ":"
    .Cells(4, 24 - 1) = ":"
    .Cells(5, 24 - 1) = ":"
    .Cells(6, 24 - 1) = ":"
    .Cells(7, 24 - 1) = ":"
    
    '.Cells(12, 1) = "Acceptance of this order and the attachment is subject to the terms and conditions on the back here of" 'A12
    '.Cells(13, 1) = "which are incorporated and made a part here of :"    'A13
        
    .Cells(15, 1) = "No."
    .Cells(15, 2) = "Description"
    .Cells(15, 17) = "Unit Price"
    .Cells(15, 21) = "Quantity"
    .Cells(15, 25) = "Total"
    

    'TEXT FOOTER
    .Cells(53, 2) = "Said Total:"
    .Cells(52, 17) = "TOTAL"
    .Cells(53, 17) = "Include PPN"
    .Cells(54, 17) = "GRAND TOTAL"
    .Cells(56, 1) = "Delivery  Address :"
    .Cells(61, 1) = "Tlp"
    .Cells(62, 1) = "Fax"
    .Cells(63, 1) = "CP"
    .Cells(65, 1) = "General Notes:"
    .Cells(56, 15) = "Note:"
    .Cells(62, 15) = "Cara pembayaran:"
    
    For K = 1 To 4
        If K <= 2 Then .Cells(60 + K, 7) = " / "
        If K <= 3 Then .Cells(60 + K, 2) = ":"
        .Cells(65 + K, 1) = "=" & Chr(34) & Str(K) & "." & Chr(34)
        .Cells(56 + K, 15) = "=" & Chr(34) & Str(K) & "." & Chr(34)
        .Cells(62 + K, 15) = "=" & Chr(34) & Str(K) & "." & Chr(34)
    Next K

    'MARGE & ALIGN & FONT
    For i = 1 To 54
            'MsgBox i
        If i <= 2 Then
            .Range(.Cells(11 + i, 1), .Cells(11 + i, 28)).MergeCells = True 'A12-A13:AB12-AB13
            .Range(.Cells(76 + i, 9), .Cells(76 + i, 13)).MergeCells = True 'I77-I78:M77-M78
            .Range(.Cells(76 + i, 14), .Cells(76 + i, 18)).MergeCells = True 'N77-N78:R77-R78
            .Range(.Cells(76 + i, 19), .Cells(76 + i, 23)).MergeCells = True 'O77-O78:W77-W78
            '.Range(.Cells(52 + i, 2), .Cells(52 + i, 16)).MergeCells = True 'B53-B54:P53-P54
        End If
        If i <= 4 Then
            .Range(.Cells(2 + i, 1), .Cells(2 + i, 18)).MergeCells = True   'A3-A6:R3-R6
            .Range(.Cells(2 + i, 24), .Cells(2 + i, 24)).HorizontalAlignment = xlRight
            .Range(.Cells(7 + i, 3), .Cells(7 + i, 18)).MergeCells = True   'C8-C11:R8-R11
            .Range(.Cells(2 + i, 19), .Cells(2 + i, 22)).MergeCells = True   'T3-T6:W3-W6
            .Range(.Cells(2 + i, 24), .Cells(2 + i, 28)).MergeCells = True   'Y3-Y6:AB3-AB6
            .Range(.Cells(55 + i, 1), .Cells(55 + i, 13)).MergeCells = True   'A56-A59:M56-M59
            .Range(.Cells(65 + i, 2), .Cells(65 + i, 13)).MergeCells = True   'B66-B69:M66-M69
            .Range(.Cells(60 + i, 3), .Cells(60 + i, 6)).MergeCells = True   'C61-C64:F61-F64
            .Range(.Cells(60 + i, 8), .Cells(60 + i, 11)).MergeCells = True   'H61-H64:K61-K64
            
            .Range(.Cells(56 + i, 16), .Cells(56 + i, 28)).MergeCells = True  'H61-H64:K61-K64
            .Range(.Cells(62 + i, 16), .Cells(62 + i, 28)).MergeCells = True  'H61-H64:K61-K64

        End If
        If i = 7 Then
            .Range(.Cells(i, 19), .Cells(i, 22)).MergeCells = True  'T7:W7
            .Range(.Cells(i, 24), .Cells(i, 28)).MergeCells = True  'Y7:AB7
            .Range(.Cells(i, 1), .Cells(i, 18)).MergeCells = True  'A7:R7
        End If
        If i >= 16 Then
            .Range(.Cells(i, 2), .Cells(i, 16)).MergeCells = True  'B6-B36:P6-P39 --> Description
            .Range(.Cells(i, 17), .Cells(i, 20)).MergeCells = True  'Q6-Q36:T6-T39 --> Unit Price
            .Range(.Cells(i, 21), .Cells(i, 22)).MergeCells = True  'U6-U36:V6-V39 --> Quantity
            .Range(.Cells(i, 23), .Cells(i, 24)).MergeCells = True  'W6-W36:X6-X39 --> Quantity
            .Range(.Cells(i, 25), .Cells(i, 28)).MergeCells = True  'Y6-Y36:AB6-AB39 --> Total
            .Cells(i, 25) = "=IF(A" & i & ">=1,Q" & i & "*U" & i & "," & Chr(34) & Chr(34) & ")" '=IF(V16>=1,Q16*U16,"") --> rumus total
            
        End If
    
        
    Next
        'KOLOM HEADER
        .Range(.Cells(15, 2), .Cells(15, 16)).MergeCells = True ' DESCRIPTION
        .Range(.Cells(15, 17), .Cells(15, 20)).MergeCells = True ' UNIT PRICE
        .Range(.Cells(15, 21), .Cells(15, 24)).MergeCells = True 'QTY
        .Range(.Cells(15, 25), .Cells(15, 28)).MergeCells = True 'TOTAL
        .Range(.Cells(15, 2), .Cells(15, 28)).HorizontalAlignment = 3
        
        'FONT BOLD / ITALIC / ALIGN
        .Range(.Cells(3, 1), .Cells(11, 2)).Font.Bold = True
        .Range(.Cells(4, 1), .Cells(7, 1)).Font.Bold = False
        .Range(.Cells(3, 1), .Cells(11, 2)).HorizontalAlignment = 1
        .Range(.Cells(3, 19), .Cells(7, 23)).Font.Bold = True
        
        
        .Range(.Cells(12, 1), .Cells(13, 2)).Font.Italic = True
        .Range(.Cells(15, 1), .Cells(15, 28)).Font.Bold = True
        .Range(.Cells(52, 1), .Cells(56, 17)).Font.Bold = True
        .Range(.Cells(54, 2), .Cells(54, 2)).Font.Italic = True 'SAID PAID
        .Range(.Cells(54, 2), .Cells(54, 2)).Font.Bold = False
        
        .Range(.Cells(62, 15), .Cells(62, 15)).Font.Bold = True
        .Range(.Cells(61, 1), .Cells(65, 1)).Font.Bold = True
        .Range(.Cells(77, 9), .Cells(77, 19)).Font.Bold = True
        .Range(.Cells(77, 9), .Cells(78, 19)).HorizontalAlignment = 3
        .Range(.Cells(77, 9), .Cells(78, 19)).Font.Underline = True
        
        'FORMAT CELL
       .Range(.Cells(16, 17), .Cells(54, 25)).NumberFormat = "#,##0" '=Y16:AB54 TOTAL FORMAT
       .Range(.Cells(53, 21), .Cells(53, 21)).NumberFormat = "00%"  '=U53 PPN FORMAT
       
       .Cells(52, 21) = "=SUM(U16:V51)" 'TOTAL QTY
       .Cells(52, 25) = "=SUM(Y16:AB51)" 'TOTAL HARGA
       .Cells(53, 25) = "=U53*Y52" 'TOTAL PPN
        .Range(.Cells(53, 25), .Cells(53, 25)).ColumnWidth = "4.14"
       .Cells(54, 25) = "=Y52+Y53" 'GRAND TOTAL

    'BORDER
    .Range(.Cells(15, 1), .Cells(15, 28)).BorderAround ColorIndex:=1, Weight:=xlMedium 'NO
    .Range(.Cells(15, 1), .Cells(51, 1)).BorderAround ColorIndex:=1, Weight:=xlMedium 'DESCRIPTION
    .Range(.Cells(15, 2), .Cells(51, 16)).BorderAround ColorIndex:=1, Weight:=xlMedium 'UNIT PRICE
    .Range(.Cells(15, 17), .Cells(51, 20)).BorderAround ColorIndex:=1, Weight:=xlMedium
    .Range(.Cells(15, 21), .Cells(51, 24)).BorderAround ColorIndex:=1, Weight:=xlMedium 'QUANTITY
        .Range(.Cells(52, 21), .Cells(52, 24)).BorderAround ColorIndex:=1, Weight:=xlMedium
        .Range(.Cells(16, 23), .Cells(51, 23)).Borders(xlEdgeLeft).LineStyle = xlDash
    .Range(.Cells(15, 25), .Cells(51, 28)).BorderAround ColorIndex:=1, Weight:=xlMedium 'TOTAL
        .Range(.Cells(52, 25), .Cells(54, 28)).BorderAround ColorIndex:=1, Weight:=xlMedium
    .Range(.Cells(56, 15), .Cells(67, 28)).BorderAround ColorIndex:=1, Weight:=xlThin 'HEADER
    End With
'</FILL PO DESIGN>
End If

'mengambil attribut
If CheckOnly Then SetEks = "'1'" Else SetEks = "'0'"
If RepCorp = "" Then SetCorp = "" Else SetCorp = " AND B.CORP_ID='" & RepCorp & "'"
If RepCab = "" Then Setcab = "" Else Setcab = " AND B.CAB_ID='" & RepCab & "'"
SetTgl = " AND B." & TglRo

'menghitung jumlah List PO
PoCount = GetDtaTbl(conMain, "SELECT COUNT(DISTINCT B.RO_ID) AS RoCount FROM tab_ro_header B WHERE B.RO_STS='0'" & SetCorp & Setcab & SetTgl & " AND B.EXPORT=" & SetEks, "RoCount")
'MsgBox PoCount
'<FILL DATA PO>
On Error Resume Next
   
    Curr_Row = 1
    strSql = "SELECT B.RO_ID FROM tab_Ro_header B WHERE B.RO_STS='0'" & SetCorp & Setcab & SetTgl & " AND B.EXPORT=" & SetEks
    'MsgBox strSql
    '''Clipboard.Clear
   ' 'Clipboard.SetText strSql
    Set oRS = ExecuteRecordSetMySql(strSql, conMain)
    With oRS
         .MoveFirst
        Do While Not .EOF
            Set oSheet = oBook.Worksheets(Curr_Row)
            RoId = .Fields("RO_ID").Value
             'MsgBox RoId
            oSheet.Name = RoId 'NAMA SHEET
            
            'Copy SHEET TEMPLATE
            If Curr_Row <> PoCount Then
                oSheet.Copy After:=oSheet
                Curr_Row = Curr_Row + 1
            End If
            
            '[HEADER-DATA]
            'RO ID. [R2-R4]
            oSheet.Cells(2, 18) = RoId
            oSheet.Cells(3, 18) = GetDtaTbl(conMain, "SELECT RO_TGL FROM tab_ro_header WHERE RO_ID='" & RoId & "'", "RO_TGL")
            oSheet.Cells(4, 18) = GetDtaTbl(conMain, "SELECT DISTINCT DEP_NM FROM tab_ro_detail WHERE RO_ID='" & RoId & "'", "DEP_NM")
            
            'Set Logo Perusahaan
            oSheet.Shapes.AddPicture PicPath, False, True, 10, 20, 100, 50
                 
            
            '[BODY-DATA]
            'LIST ITEM
            strSQLitem = "SELECT B.BRG_NM,B.BRG_HRG,B.RO_QTY,D.BRG_SAT " & _
                         "FROM tab_ro_header A,tab_ro_detail B,tab_barang D " & _
                         "WHERE A.RO_ID='" & RoId & "' " & _
                         "AND A.RO_ID=B.RO_ID AND B.BRG_ID=D.BRG_ID" ' AND  C.CORP_ID='" & RepCorp & "' AND C.CAB_ID='" & RepCab & "'"
            
            Set oRSitem = ExecuteRecordSetMySql(strSQLitem, conMain)
            curRow = 1
            With oRSitem
                .MoveFirst
                Do While Not .EOF
                    oSheet.Cells(6 + curRow, 1) = curRow
                    oSheet.Cells(6 + curRow, 2) = .Fields("BRG_NM").Value
                    oSheet.Cells(6 + curRow, 18) = .Fields("BRG_SAT").Value
                    oSheet.Cells(6 + curRow, 20) = .Fields("RO_QTY").Value
                .MoveNext
                curRow = curRow + 1
                Loop
            End With
          .MoveNext
          
          '[/BODY-DATA]
          
          '[FOOTER-DATA]

            'SIGN JABATAN
            'oSheet.Cells(78, 9) = GetDtaTbl(conMain, "SELECT A.REP_SIGN1_JAB FROM tab_po_detail_report A,tab_po_header B WHERE B.PO_ID='" & RoId & "' AND A.SPL_ID=B.SPL_ID", "REP_SIGN1_JAB") 'SIGN JAB 1
            'oSheet.Cells(78, 14) = GetDtaTbl(conMain, "SELECT A.REP_SIGN2_JAB FROM tab_po_detail_report A,tab_po_header B WHERE B.PO_ID='" & RoId & "' AND A.SPL_ID=B.SPL_ID", "REP_SIGN2_JAB") 'SIGN JAB 2
            'oSheet.Cells(78, 19) = GetDtaTbl(conMain, "SELECT A.REP_SIGN3_JAB FROM tab_po_detail_report A,tab_po_header B WHERE B.PO_ID='" & RoId & "' AND A.SPL_ID=B.SPL_ID", "REP_SIGN3_JAB") 'SIGN JAB 3
        
          '[/FOOTER-DATA]
          

        Loop
    End With
    
    Set oRS = Nothing
'</FILL DATA PO>

ter:
  ' oBook.SaveAs "D:\PO.xls"
  ' oExcel.Quit
   
   'oRS.MoveFirst
   'oRS.Close
   'MsgBox "Report File format  File Saved", 64, "Info"
Exit Sub
ErrHandReport:
Select Case Err.Number
    Case Is = 1
        MsgBox "Harap Tutup File PO Excel yang sedang aktif", vbCritical, "Perhatian!"
End Select

 
End Sub
Public Sub Berita2Excel(BeritaId As String)
'============================================
'|              LG PROJECT     [23 May 2014]|
'============================================
'|ClsReportExcel - Berita2Excel             |
'|(digunakan untuk menampilkan report BA)   |
'|Progreass = 10%                           |
'--------------------------------------------
'|By. FARINDRA                  @18-07-2014 |
'============================================
Dim curRow, DevSheet, i, K As Integer

Set oExcel = CreateObject("Excel.Application"): oExcel.Visible = True '= ONOFF ==
Set oBook = oExcel.Workbooks.Add

'<FILL BERITA DESIGN>

'DEFAULT PROPERTI CELL
    'oSheet.Columns.AutoFit
DevSheet = 1
Set oSheet = oBook.Worksheets(DevSheet)
With oSheet
    .Cells.Font.Name = "Arial"
    .Cells.Font.Size = 12
    .Cells.ColumnWidth = 4.5
    .PageSetup.Zoom = 92
    .PageSetup.PaperSize = xlPaperA4
    .PageSetup.TopMargin = 0.5
    .PageSetup.LeftMargin = 0.5
    .PageSetup.RightMargin = 0.5
    
    
 '[DESIGN BERITA ACARA - JUDUL]
    'TEXT
    .Cells(2, 2) = "BERITA ACARA" 'B2
    
    'MARGE
        .Range(.Cells(2, 2), .Cells(2, 20)).MergeCells = True
    'FONT & ALIGMENT
        .Range(.Cells(2, 2), .Cells(2, 20)).Font.Size = 18
        .Range(.Cells(2, 2), .Cells(2, 20)).Font.Bold = True
        .Range(.Cells(2, 2), .Cells(2, 20)).HorizontalAlignment = 3
 
 '[DESIGN BERITA ACARA - BODY]
    'TEXT HEADER
    .Cells(5, 2) = "NO."
    .Cells(6, 2) = "Tanggal"
    .Cells(5, 17) = "Jam"
    .Cells(8, 2) = "Peristiwa / Kejadian:"
    
    .Cells(5, 4) = ":"
    .Cells(6, 4) = ":"
    .Cells(5, 18) = ":"

    'TEXT FOOTER
    .Cells(20, 2) = "Dibuat Oleh,"
    .Cells(20, 12) = "Mengetahui,"

    'MARGE & ALIGN & FONT
        .Range(.Cells(5, 5), .Cells(5, 9)).MergeCells = True   'No
        .Range(.Cells(6, 5), .Cells(6, 9)).MergeCells = True   'Tanggal
        .Range(.Cells(5, 19), .Cells(5, 20)).MergeCells = True  'Jam
        .Range(.Cells(8, 2), .Cells(8, 5)).MergeCells = True  'Peristiwa <-----
        
        .Range(.Cells(20, 2), .Cells(20, 5)).MergeCells = True  'Dibuat Oleh
        .Range(.Cells(20, 2), .Cells(20, 5)).HorizontalAlignment = 3
        .Range(.Cells(20, 12), .Cells(20, 15)).MergeCells = True  'Mengetahui
        .Range(.Cells(20, 12), .Cells(20, 5)).HorizontalAlignment = 3
        
        .Range(.Cells(24, 2), .Cells(24, 5)).MergeCells = True  'Nama Dibuat Oleh
        .Range(.Cells(24, 2), .Cells(24, 5)).HorizontalAlignment = 3
            .Range(.Cells(25, 2), .Cells(25, 5)).MergeCells = True  'jabatan Dibuat Oleh
            .Range(.Cells(25, 2), .Cells(25, 5)).HorizontalAlignment = 3
        .Range(.Cells(24, 7), .Cells(24, 10)).MergeCells = True  'Nama Mengetahui 1
        .Range(.Cells(24, 7), .Cells(24, 10)).HorizontalAlignment = 3
            .Range(.Cells(25, 7), .Cells(25, 10)).MergeCells = True  'jabatan Mengetahui 1
            .Range(.Cells(25, 7), .Cells(25, 10)).HorizontalAlignment = 3
        .Range(.Cells(24, 12), .Cells(24, 15)).MergeCells = True  'Nama Mengetahui 2
        .Range(.Cells(24, 12), .Cells(24, 15)).HorizontalAlignment = 3
            .Range(.Cells(25, 12), .Cells(25, 15)).MergeCells = True  'jabatan Mengetahui 2
            .Range(.Cells(25, 12), .Cells(25, 15)).HorizontalAlignment = 3
        .Range(.Cells(24, 17), .Cells(24, 20)).MergeCells = True  'Nama Mengetahui 3
        .Range(.Cells(24, 17), .Cells(24, 20)).HorizontalAlignment = 3
            .Range(.Cells(25, 17), .Cells(25, 20)).MergeCells = True  'jabatan Mengetahui 3
            Range(.Cells(25, 17), .Cells(25, 20)).HorizontalAlignment = 3
        
        'FONT BOLD / ITALIC / ALIGN
        .Range(.Cells(2, 2), .Cells(3, 20)).Font.Bold = True
        .Range(.Cells(24, 2), .Cells(25, 20)).Font.Size = 9
        .Range(.Cells(24, 2), .Cells(24, 20)).Font.Bold = True
        
        'FORMAT CELL
        .Range(.Cells(6, 5), .Cells(6, 5)).NumberFormat = "dd MMMM yyyy"
       '.Range(.Cells(53, 21), .Cells(53, 21)).NumberFormat = "00%"  '=U53 PPN FORMAT
       
    'BORDER
    .Range(.Cells(1, 1), .Cells(26, 21)).BorderAround ColorIndex:=1, Weight:=xlMedium 'BORDER
    
    .Range(.Cells(5, 5), .Cells(5, 9)).Borders(xlEdgeBottom).Weight = xlThin 'No Underline
    .Range(.Cells(6, 5), .Cells(6, 9)).Borders(xlEdgeBottom).Weight = xlThin 'Tanggal Underline
    .Range(.Cells(5, 19), .Cells(5, 20)).Borders(xlEdgeBottom).Weight = xlThin 'Jam Underline
    
    For i = 10 To 17
        .Range(.Cells(i, 2), .Cells(i, 20)).Borders(xlEdgeBottom).LineStyle = xlDot   'isi
    Next i
    
    .Range(.Cells(24, 2), .Cells(24, 5)).Borders(xlEdgeBottom).Weight = xlThin 'dibuat oleh
    .Range(.Cells(24, 7), .Cells(24, 10)).Borders(xlEdgeBottom).Weight = xlThin 'mengetahui 1
    .Range(.Cells(24, 12), .Cells(24, 15)).Borders(xlEdgeBottom).Weight = xlThin 'mengetahui 2
    .Range(.Cells(24, 17), .Cells(24, 20)).Borders(xlEdgeBottom).Weight = xlThin 'mengetahui 3
    
    
'</FILL BERITA DESIGN>

'<FILL DATA>
    'HEADER
    .Cells(5, 5) = BeritaId 'No BERITA
    .Cells(6, 5) = GetDtaTbl(conMain, "SELECT TGL FROM tab_bt_header WHERE BT_ID='" & BeritaId & "'", "TGL") 'tanggal
    .Cells(5, 19) = GetDtaTbl(conMain, "SELECT WAKTU FROM tab_bt_header WHERE BT_ID='" & BeritaId & "'", "WAKTU") 'Jam
    
    'BODY
   ' ''Clipboard.Clear
    ''Clipboard.SetText GetDtaTbl(conMain, "SELECT BT_ISI FROM tab_bt_header WHERE BT_ID='" & BeritaId & "'", "BT_ISI")  'Jam
    '.Cells(10, 2) = 'Clipboard.GetText 'GetDtaTbl(conMain, "SELECT BT_ISI FROM tab_bt_header WHERE BT_ID='" & BeritaId & "'", "BT_ISI") 'Jam
    .Range(.Cells(10, 2), .Cells(10, 2)).PasteSpecial Paste:=xlPasteAll ', Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ''Clipboard.GetText
    
    'FOOTER
    .Cells(24, 2) = GetDtaTbl(conMain, "SELECT DISTINCT USER_NM FROM tab_bt_detail WHERE BT_ID='" & BeritaId & "'", "USER_NM") 'Nama Pembuat
    .Cells(25, 2) = GetDtaTbl(conMain, "SELECT DISTINCT DEP_NM FROM tab_bt_detail WHERE BT_ID='" & BeritaId & "'", "DEP_NM") & " " & _
                    GetDtaTbl(conMain, "SELECT DISTINCT JAB_NM FROM tab_bt_detail WHERE BT_ID='" & BeritaId & "'", "JAB_NM") 'departemen & jabatan Pembuat
    
    .Cells(24, 7) = GetDtaTbl(conMain, "SELECT DISTINCT TARGET_KAR_NM FROM tab_bt_detail WHERE BT_ID='" & BeritaId & "' AND BT_KON_TYPE='M0'", "TARGET_KAR_NM")
    .Cells(25, 7) = GetDtaTbl(conMain, "SELECT DISTINCT TARGET_KAR_DEP_NM FROM tab_bt_detail WHERE BT_ID='" & BeritaId & "' AND BT_KON_TYPE='M0'", "TARGET_KAR_DEP_NM") & " " & _
                    GetDtaTbl(conMain, "SELECT DISTINCT TARGET_KAR_JAB_NM FROM tab_bt_detail WHERE BT_ID='" & BeritaId & "' AND BT_KON_TYPE='M0'", "TARGET_KAR_JAB_NM") ' "Mengetahui 1"
    
    .Cells(24, 12) = GetDtaTbl(conMain, "SELECT DISTINCT TARGET_KAR_NM FROM tab_bt_detail WHERE BT_ID='" & BeritaId & "' AND BT_KON_TYPE='M1'", "TARGET_KAR_NM")
    .Cells(25, 12) = GetDtaTbl(conMain, "SELECT DISTINCT TARGET_KAR_DEP_NM FROM tab_bt_detail WHERE BT_ID='" & BeritaId & "' AND BT_KON_TYPE='M1'", "TARGET_KAR_DEP_NM") & " " & _
                     GetDtaTbl(conMain, "SELECT DISTINCT TARGET_KAR_JAB_NM FROM tab_bt_detail WHERE BT_ID='" & BeritaId & "' AND BT_KON_TYPE='M1'", "TARGET_KAR_JAB_NM") ' "Mengetahui 1"
    
    .Cells(24, 17) = GetDtaTbl(conMain, "SELECT DISTINCT TARGET_KAR_NM FROM tab_bt_detail WHERE BT_ID='" & BeritaId & "' AND BT_KON_TYPE='M2'", "TARGET_KAR_NM")
    .Cells(25, 17) = GetDtaTbl(conMain, "SELECT DISTINCT TARGET_KAR_DEP_NM FROM tab_bt_detail WHERE BT_ID='" & BeritaId & "' AND BT_KON_TYPE='M2'", "TARGET_KAR_DEP_NM") & " " & _
                     GetDtaTbl(conMain, "SELECT DISTINCT TARGET_KAR_JAB_NM FROM tab_bt_detail WHERE BT_ID='" & BeritaId & "' AND BT_KON_TYPE='M2'", "TARGET_KAR_JAB_NM") ' "Mengetahui 1"
'</FILL DATA>
End With
End Sub
Public Sub MargeCell(Cell As String, Align As Integer)
Dim MaxCol As Integer
Dim Col(28) As String
Dim Row(2) As Integer
Dim J As Integer

Dim MidCell As Integer
MaxCol = 28

'set karakter A-Z sebagai colom
For i = 1 To MaxCol
    Col(i) = Chr(64 + i)
Next

'membaca poperti cell
UCase (Cell)
MidCell = InStr(1, Cell, ":")

For J = 1 To Len(Cell)
    
Next J



End Sub
Public Sub Prensensi_IN_OUT(xQryFilter As String) 'RepCorp As String, RepCab As String, TglPo As String, CheckOnly As Boolean, PicPath As String)
Dim RoId As String
Dim SetEks, SetCorp, Setcab, SetTgl As String
Dim Col(10), Row(10) As Integer
Dim curRow, DevSheet, i, K, PoCount, PoPPN As Integer
Dim oRS As New ADODB.Recordset
Dim oRSitem As New ADODB.Recordset

'On Error GoTo ErrHandReport
'cek jika file template ada
If CekStatusFile("Template_Presensi_IN_OUT.xlsx", App.Path & "\Report\") Then
    File_Copy App.Path & "\Report\", "Template_Presensi_IN_OUT.xlsx", App.Path & "\Report\PRESENSI\", "PRESENSI_IN_OUT_" & Format(Now, "YYYYMMDD") & ".xlsx"

    'buka file excel tersebut
    Set oExcel = CreateObject("Excel.Application"):  oExcel.Visible = False
    Set oBook = oExcel.Workbooks.Open(App.Path & "\Report\PRESENSI\PRESENSI_IN_OUT_" & Format(Now, "YYYYMMDD") & ".xlsx")
    DevSheet = 1
    Set oSheet = oBook.Worksheets(DevSheet)
    'Goto Dataprose
Else
    'jika tidak ada
    Exit Sub
    'buat file excel baru jika tidak ada file template "template_report_po.xlsx"
    Set oExcel = CreateObject("Excel.Application"): 'oExcel.Visible = False '= ONOFF ==
    Set oBook = oExcel.Workbooks.Add:

    '<FILL DESIGN>

    '</FILL DESIGN>
    End If

'<FILL>
'On Error Resume Next
    'HEADER PT
LOADING.Show
LOADING.SetParm FRM_ABSENSI, 10
    oSheet.Cells(1, 1) = GetDtaTbl(conMain, "SELECT CORP_NM FROM tab_corp WHERE CORP_ID='" & IdCorp & "'", "CORP_NM")
    oSheet.Cells(2, 1) = GetDtaTbl(conMain, "SELECT CORP_ALM_JLN FROM tab_corp WHERE CORP_ID='" & IdCorp & "'", "CORP_ALM_JLN") & " " & GetDtaTbl(conMain, "SELECT CORP_ALM_CTY FROM tab_corp WHERE CORP_ID='" & IdCorp & "'", "CORP_ALM_CTY")
    oSheet.Cells(3, 1) = "LAPORAN PRESENSI - IN_OUT"
    
    Set oRSitem = ExecuteRecordSetMySqlProc(xQryFilter, conMain)
    curRow = 1
    
    With oRSitem
        .MoveFirst
        Do While Not .EOF
        If (curRow / 10) <= 90 Then LOADING.SetParm FRM_ABSENSI, 10 + (curRow / 10) Else LOADING.SetParm FRM_ABSENSI, 95
    
            oSheet.Cells(7 + curRow, 1) = curRow
            
            oSheet.Range(oSheet.Cells(7 + curRow, 2), oSheet.Cells(7 + curRow, 4)).MergeCells = True
                oSheet.Cells(7 + curRow, 2) = Format(.Fields("TGL_RUN").Value, "DD/MM/YYYY")
            oSheet.Range(oSheet.Cells(7 + curRow, 5), oSheet.Cells(7 + curRow, 7)).MergeCells = True
                oSheet.Cells(7 + curRow, 5) = .Fields("TT_GRP_NM").Value
            oSheet.Range(oSheet.Cells(7 + curRow, 8), oSheet.Cells(7 + curRow, 10)).MergeCells = True
                oSheet.Cells(7 + curRow, 8) = .Fields("TT_TYP").Value
            oSheet.Range(oSheet.Cells(7 + curRow, 14), oSheet.Cells(7 + curRow, 15)).MergeCells = True
                oSheet.Cells(7 + curRow, 14) = GetDtaTbl(conMain, "SELECT Hari_Nm FROM tab_timetabel_hari WHERE Hari_id='" & .Fields("hari").Value & "'", "Hari_Nm")
            oSheet.Range(oSheet.Cells(7 + curRow, 16), oSheet.Cells(7 + curRow, 17)).MergeCells = True
                oSheet.Cells(7 + curRow, 16) = Format(.Fields("RULE_IN").Value, "hh:mm")
            oSheet.Range(oSheet.Cells(7 + curRow, 18), oSheet.Cells(7 + curRow, 19)).MergeCells = True
                oSheet.Cells(7 + curRow, 18) = Format(.Fields("RULE_OUT").Value, "hh:mm")
            oSheet.Range(oSheet.Cells(7 + curRow, 11), oSheet.Cells(7 + curRow, 13)).MergeCells = True
                oSheet.Cells(7 + curRow, 11) = .Fields("TT_NOTE").Value
            oSheet.Range(oSheet.Cells(7 + curRow, 20), oSheet.Cells(7 + curRow, 24)).MergeCells = True
                oSheet.Cells(7 + curRow, 20) = .Fields("KAR_NM").Value
            oSheet.Range(oSheet.Cells(7 + curRow, 25), oSheet.Cells(7 + curRow, 26)).MergeCells = True
                oSheet.Cells(7 + curRow, 25) = .Fields("DateTimeLogIN").Value
            oSheet.Range(oSheet.Cells(7 + curRow, 27), oSheet.Cells(7 + curRow, 28)).MergeCells = True
                oSheet.Cells(7 + curRow, 27) = .Fields("DateTimeLogOUT").Value
            oSheet.Range(oSheet.Cells(7 + curRow, 29), oSheet.Cells(7 + curRow, 30)).MergeCells = True
                oSheet.Cells(7 + curRow, 29) = GetDtaTbl(conMain, "SELECT LOCK_NM FROM tab_presensi_lev2_lock WHERE LOCK_STT='" & .Fields("LOCK_STT").Value & "'", "LOCK_NM")
            oSheet.Range(oSheet.Cells(7 + curRow, 31), oSheet.Cells(7 + curRow, 32)).MergeCells = True
                oSheet.Cells(7 + curRow, 31) = .Fields("DAY_ACTIVE").Value
            oSheet.Range(oSheet.Cells(7 + curRow, 33), oSheet.Cells(7 + curRow, 34)).MergeCells = True
                oSheet.Cells(7 + curRow, 33) = .Fields("DAY_LIBUR").Value
            oSheet.Range(oSheet.Cells(7 + curRow, 35), oSheet.Cells(7 + curRow, 36)).MergeCells = True
                oSheet.Cells(7 + curRow, 35) = .Fields("DAY_ABSENT").Value
            oSheet.Range(oSheet.Cells(7 + curRow, 37), oSheet.Cells(7 + curRow, 38)).MergeCells = True
                oSheet.Cells(7 + curRow, 37) = .Fields("DAY_EXCP_ID").Value
            oSheet.Range(oSheet.Cells(7 + curRow, 39), oSheet.Cells(7 + curRow, 40)).MergeCells = True
               ' oSheet.Cells(7 + curRow, 39) = .Fields("WORK_TIME").Value
            oSheet.Range(oSheet.Cells(7 + curRow, 41), oSheet.Cells(7 + curRow, 42)).MergeCells = True
                oSheet.Cells(7 + curRow, 41) = .Fields("LATE").Value
            oSheet.Range(oSheet.Cells(7 + curRow, 43), oSheet.Cells(7 + curRow, 44)).MergeCells = True
                oSheet.Cells(7 + curRow, 43) = .Fields("EARLY").Value
            oSheet.Range(oSheet.Cells(7 + curRow, 45), oSheet.Cells(7 + curRow, 47)).MergeCells = True
                oSheet.Cells(7 + curRow, 45) = .Fields("OT_DPN").Value
            'oSheet.Cells(7 + curRow, 48) = .Fields("OT_BLK").Value
            'oSheet.Range(oSheet.Cells(7 + curRow, 51), oSheet.Cells(7 + curRow, 4)).MergeCells = True
            'oSheet.Cells(7 + curRow, 51) = .Fields("KAR_ID").Value
            
            oSheet.Range(oSheet.Cells(7 + curRow, 1), oSheet.Cells(7 + curRow, 50)).Borders.LineStyle = xlContinuous
            .MoveNext
            curRow = curRow + 1
        Loop
        End With
    Set oRSitem = Nothing
'</FILL>

  'Clipboard.Clear
  oExcel.Visible = True '= ONOFF ==
  LOADING.SetParm FRM_ABSENSI, 100
Exit Sub

ErrHandReport:
Select Case Err.Number
    Case Is = 1
        MsgBox "Harap Tutup File Excel yang sedang aktif", vbCritical, "Perhatian!"
End Select
End Sub

Public Sub Prensensi_OVER_TIME(xQryFilter As String) 'RepCorp As String, RepCab As String, TglPo As String, CheckOnly As Boolean, PicPath As String)
Dim RoId As String
Dim SetEks, SetCorp, Setcab, SetTgl As String
Dim Col(10), Row(10) As Integer
Dim curRow, DevSheet, i, K, PoCount, PoPPN As Integer
Dim oRS As New ADODB.Recordset
Dim oRSitem As New ADODB.Recordset

'On Error GoTo ErrHandReport
'cek jika file template ada
If CekStatusFile("Template_Presensi_IN_OUT.xlsx", App.Path & "\Report\") Then
    File_Copy App.Path & "\Report\", "Template_Presensi_IN_OUT.xlsx", App.Path & "\Report\PRESENSI\", "PRESENSI_IN_OUT_" & Format(Now, "YYYYMMDD") & ".xlsx"

    'buka file excel tersebut
    Set oExcel = CreateObject("Excel.Application"):  oExcel.Visible = False
    Set oBook = oExcel.Workbooks.Open(App.Path & "\Report\PRESENSI\PRESENSI_IN_OUT_" & Format(Now, "YYYYMMDD") & ".xlsx")
    DevSheet = 1
    Set oSheet = oBook.Worksheets(DevSheet)
    'Goto Dataprose
Else
    'jika tidak ada
    Exit Sub
    'buat file excel baru jika tidak ada file template "template_report_po.xlsx"
    Set oExcel = CreateObject("Excel.Application"): 'oExcel.Visible = False '= ONOFF ==
    Set oBook = oExcel.Workbooks.Add:

    '<FILL DESIGN>

    '</FILL DESIGN>
    End If

'<FILL>
'On Error Resume Next
    'HEADER PT
LOADING.Show
LOADING.SetParm FRM_ABSENSI, 10
    oSheet.Cells(1, 1) = GetDtaTbl(conMain, "SELECT CORP_NM FROM tab_corp WHERE CORP_ID='" & IdCorp & "'", "CORP_NM")
    oSheet.Cells(2, 1) = GetDtaTbl(conMain, "SELECT CORP_ALM_JLN FROM tab_corp WHERE CORP_ID='" & IdCorp & "'", "CORP_ALM_JLN") & " " & GetDtaTbl(conMain, "SELECT CORP_ALM_CTY FROM tab_corp WHERE CORP_ID='" & IdCorp & "'", "CORP_ALM_CTY")
    oSheet.Cells(3, 1) = "LAPORAN PRESENSI - IN_OUT"
    
    Set oRSitem = ExecuteRecordSetMySqlProc(xQryFilter, conMain)
    curRow = 1
    
    With oRSitem
        .MoveFirst
        Do While Not .EOF
        If (curRow / 10) <= 90 Then LOADING.SetParm FRM_ABSENSI, 10 + (curRow / 10) Else LOADING.SetParm FRM_ABSENSI, 95
    
            oSheet.Cells(7 + curRow, 1) = curRow
            
            oSheet.Range(oSheet.Cells(7 + curRow, 2), oSheet.Cells(7 + curRow, 4)).MergeCells = True
                oSheet.Cells(7 + curRow, 2) = Format(.Fields("TGL_RUN").Value, "DD/MM/YYYY")
            oSheet.Range(oSheet.Cells(7 + curRow, 5), oSheet.Cells(7 + curRow, 7)).MergeCells = True
                oSheet.Cells(7 + curRow, 5) = .Fields("TT_GRP_NM").Value
            oSheet.Range(oSheet.Cells(7 + curRow, 8), oSheet.Cells(7 + curRow, 10)).MergeCells = True
                oSheet.Cells(7 + curRow, 8) = .Fields("TT_TYP").Value
            oSheet.Range(oSheet.Cells(7 + curRow, 14), oSheet.Cells(7 + curRow, 15)).MergeCells = True
                oSheet.Cells(7 + curRow, 14) = GetDtaTbl(conMain, "SELECT Hari_Nm FROM tab_timetabel_hari WHERE Hari_id='" & .Fields("hari").Value & "'", "Hari_Nm")
            oSheet.Range(oSheet.Cells(7 + curRow, 16), oSheet.Cells(7 + curRow, 17)).MergeCells = True
                oSheet.Cells(7 + curRow, 16) = Format(.Fields("RULE_IN").Value, "hh:mm")
            oSheet.Range(oSheet.Cells(7 + curRow, 18), oSheet.Cells(7 + curRow, 19)).MergeCells = True
                oSheet.Cells(7 + curRow, 18) = Format(.Fields("RULE_OUT").Value, "hh:mm")
            oSheet.Range(oSheet.Cells(7 + curRow, 11), oSheet.Cells(7 + curRow, 13)).MergeCells = True
                oSheet.Cells(7 + curRow, 11) = .Fields("TT_NOTE").Value
            oSheet.Range(oSheet.Cells(7 + curRow, 20), oSheet.Cells(7 + curRow, 24)).MergeCells = True
                oSheet.Cells(7 + curRow, 20) = .Fields("KAR_NM").Value
            oSheet.Range(oSheet.Cells(7 + curRow, 25), oSheet.Cells(7 + curRow, 26)).MergeCells = True
                oSheet.Cells(7 + curRow, 25) = .Fields("DateTimeLogIN").Value
            oSheet.Range(oSheet.Cells(7 + curRow, 27), oSheet.Cells(7 + curRow, 28)).MergeCells = True
                oSheet.Cells(7 + curRow, 27) = .Fields("DateTimeLogOUT").Value
            oSheet.Range(oSheet.Cells(7 + curRow, 29), oSheet.Cells(7 + curRow, 30)).MergeCells = True
                oSheet.Cells(7 + curRow, 29) = GetDtaTbl(conMain, "SELECT LOCK_NM FROM tab_presensi_lev2_lock WHERE LOCK_STT='" & .Fields("LOCK_STT").Value & "'", "LOCK_NM")
            oSheet.Range(oSheet.Cells(7 + curRow, 31), oSheet.Cells(7 + curRow, 32)).MergeCells = True
                oSheet.Cells(7 + curRow, 31) = .Fields("DAY_ACTIVE").Value
            oSheet.Range(oSheet.Cells(7 + curRow, 33), oSheet.Cells(7 + curRow, 34)).MergeCells = True
                oSheet.Cells(7 + curRow, 33) = .Fields("DAY_LIBUR").Value
            oSheet.Range(oSheet.Cells(7 + curRow, 35), oSheet.Cells(7 + curRow, 36)).MergeCells = True
                oSheet.Cells(7 + curRow, 35) = .Fields("DAY_ABSENT").Value
            oSheet.Range(oSheet.Cells(7 + curRow, 37), oSheet.Cells(7 + curRow, 38)).MergeCells = True
                oSheet.Cells(7 + curRow, 37) = .Fields("DAY_EXCP_ID").Value
            oSheet.Range(oSheet.Cells(7 + curRow, 39), oSheet.Cells(7 + curRow, 40)).MergeCells = True
               ' oSheet.Cells(7 + curRow, 39) = .Fields("WORK_TIME").Value
            oSheet.Range(oSheet.Cells(7 + curRow, 41), oSheet.Cells(7 + curRow, 42)).MergeCells = True
                oSheet.Cells(7 + curRow, 41) = .Fields("LATE").Value
            oSheet.Range(oSheet.Cells(7 + curRow, 43), oSheet.Cells(7 + curRow, 44)).MergeCells = True
                oSheet.Cells(7 + curRow, 43) = .Fields("EARLY").Value
            oSheet.Range(oSheet.Cells(7 + curRow, 45), oSheet.Cells(7 + curRow, 47)).MergeCells = True
                oSheet.Cells(7 + curRow, 45) = .Fields("OT_DPN").Value
            'oSheet.Cells(7 + curRow, 48) = .Fields("OT_BLK").Value
            'oSheet.Range(oSheet.Cells(7 + curRow, 51), oSheet.Cells(7 + curRow, 4)).MergeCells = True
            'oSheet.Cells(7 + curRow, 51) = .Fields("KAR_ID").Value
            
            oSheet.Range(oSheet.Cells(7 + curRow, 1), oSheet.Cells(7 + curRow, 50)).Borders.LineStyle = xlContinuous
            .MoveNext
            curRow = curRow + 1
        Loop
        End With
    Set oRSitem = Nothing
'</FILL>

  'Clipboard.Clear
  oExcel.Visible = True '= ONOFF ==
  LOADING.SetParm FRM_ABSENSI, 100
Exit Sub

ErrHandReport:
Select Case Err.Number
    Case Is = 1
        MsgBox "Harap Tutup File Excel yang sedang aktif", vbCritical, "Perhatian!"
End Select
End Sub

