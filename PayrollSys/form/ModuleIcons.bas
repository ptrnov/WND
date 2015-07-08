Attribute VB_Name = "ModuleIcons"
Public Declare Function ShellExecute Lib "shell32.dll" _
            Alias "ShellExecuteA" _
            (ByVal hwnd As Long, _
            ByVal lpOperation As String, _
            ByVal lpFile As String, _
            ByVal lpParameters As String, _
            ByVal lpDirectory As String, _
            ByVal nShowCmd As Long) As Long
            

Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst _
    As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 _
    As Long, ByVal un2 As Long) As Long
    

Public Const LR_LOADMAP3DCOLORS = &H1000
Public Const LR_CREATEDIBSECTION = &H2000
Public Const LR_LOADFROMFILE = &H10
Public Const LR_LOADTRANSPARENT = &H20
Public Const LR_COPYRETURNORG = &H4
Public Const IMAGE_BITMAP = 0
Public Const IMAGE_ICON = 1

Private Const ILC_COLOR = &H0
Private Const ILC_MASK = &H1
Public Const ILC_COLOR4 = &H4
Public Const ILC_COLOR8 = &H8
Public Const ILC_COLOR16 = &H10
Public Const ILC_COLOR24 = &H18
Public Const ILC_COLOR32 = &H20
Public Const ILD_NORMAL = 0

Public Function ImagePath(ObjectTarget As String) As String
Dim MnuFolder, IconFolder, BackFolder, ButtonFolder, KarFoto As String

MnuFolder = App.Path & "\Img\Menu\"
IconFolder = App.Path & "\Img\Icons\"
BackFolder = App.Path & "\Img\Background\"
LogoFolder = App.Path & "\Img\Logo\"
ButtonFolder = App.Path & "\Img\Button\"
KarFoto = App.Path & "\Img\Karyawan\"

Select Case ObjectTarget

'-------------------------LOGO-----------------------
Case Is = "LGO_SSS": ImagePath = LogoFolder & "SSS.jpg"
Case Is = "LGO_LIPAT": ImagePath = LogoFolder & "Lipat.jpg"
Case Is = "LGO_FOODTOWN": ImagePath = LogoFolder & "Foodtown.jpg"
Case Is = "LGO_LUKISON": ImagePath = LogoFolder & "Lukison.jpg"

Case Is = "EXC_SSS": ImagePath = LogoFolder & "SSS.jpg"
Case Is = "EXC_ALG": ImagePath = LogoFolder & "Lipat.jpg"
Case Is = "EXC_FOD": ImagePath = LogoFolder & "Foodtown.jpg"
Case Is = "EXC_LG": ImagePath = LogoFolder & "Lukison.jpg"

'-------------------------MENU-----------------------
'--HOME --
Case Is = "MNU_HOME_STATUS": ImagePath = MnuFolder & "Home\status.png"
Case Is = "MNU_HOME_PROSES_PRESENSI": ImagePath = MnuFolder & "Home\proses_presensi.png"
Case Is = "MNU_HOME_PRESENSI": ImagePath = MnuFolder & "Home\presensi.png"
Case Is = "MNU_HOME_IMPORT": ImagePath = MnuFolder & "Home\import.png"
Case Is = "MNU_HOME_PAYROLL": ImagePath = MnuFolder & "Home\payroll.png"
'--IT--
Case Is = "MNU_IT_FINGERMACHINE": ImagePath = MnuFolder & "IT\FingerMachine.png"
Case Is = "MNU_IT_EMPFINGER": ImagePath = MnuFolder & "IT\EmpFinger.png"


Case Is = "MNU_PURCHASE_RO": ImagePath = MnuFolder & "Purchase\ro.png"
Case Is = "MNU_PURCHASE_PO": ImagePath = MnuFolder & "Purchase\po.png"


Case Is = "MNU_WAREHOUSE_RECEIVE": ImagePath = MnuFolder & "Warehouse\receive.png"
Case Is = "MNU_WAREHOUSE_TRANSFER": ImagePath = MnuFolder & "Warehouse\transfer.png"
Case Is = "MNU_WAREHOUSE_CASE": ImagePath = MnuFolder & "Warehouse\case.png"
Case Is = "MNU_WAREHOUSE_REFISI": ImagePath = MnuFolder & "Warehouse\refisi.png"
Case Is = "MNU_WAREHOUSE_TRANSAKSI": ImagePath = MnuFolder & "Warehouse\transaksi.png"
Case Is = "MNU_WAREHOUSE_DAFTAR_BARANG": ImagePath = MnuFolder & "Warehouse\daftar_barang.png"
Case Is = "MNU_WAREHOUSE_STOCK_CARD": ImagePath = MnuFolder & "Warehouse\stock_card.png"
Case Is = "MNU_WAREHOUSE_MONTHLY": ImagePath = MnuFolder & "Warehouse\monthly.png"
Case Is = "MNU_WAREHOUSE_TRANSFER": ImagePath = MnuFolder & "Warehouse\transfer.png"
Case Is = "MNU_WAREHOUSE_DATA_GUDANG": ImagePath = MnuFolder & "Warehouse\data_gudang.png"
Case Is = "MNU_WAREHOUSE_DATA_BARANG": ImagePath = MnuFolder & "Warehouse\data_barang.png"
Case Is = "MNU_WAREHOUSE_KATEGORI": ImagePath = MnuFolder & "Warehouse\kategori.png"
Case Is = "MNU_WAREHOUSE_SATUAN": ImagePath = MnuFolder & "Warehouse\satuan.png"
Case Is = "MNU_WAREHOUSE_KONVERSI_SATUAN": ImagePath = MnuFolder & "Warehouse\konversi_satuan.png"
Case Is = "MNU_WAREHOUSE_TYPE": ImagePath = MnuFolder & "Warehouse\type.png"
Case Is = "MNU_WAREHOUSE_BARCODE": ImagePath = MnuFolder & "Warehouse\barcode.png"
Case Is = "MNU_WAREHOUSE_HARGA_JUAL": ImagePath = MnuFolder & "Warehouse\harga_jual.png"
Case Is = "MNU_WAREHOUSE_PERM_AKUN": ImagePath = MnuFolder & "Warehouse\perm_akun.png"

Case Is = "MNU_INVENTORY": ImagePath = MnuFolder & "inventory.png"
Case Is = "MNU_USER": ImagePath = MnuFolder & "user.png"
Case Is = "MNU_BERITA_MASUK": ImagePath = MnuFolder & "beritamasuk.png"
Case Is = "MNU_BERITA_KELUAR": ImagePath = MnuFolder & "beritakeluar.png"
Case Is = "MNU_WAREHOUSE_CARD": ImagePath = MnuFolder & "warehouse_card.png"
Case Is = "MNU_WAREHOUSE_TRANSAKSI": ImagePath = MnuFolder & "warehouse_transaksi.png"
Case Is = "MNU_WAREHOUSE_MONTHLY": ImagePath = MnuFolder & "warehouse_monthly.png"
Case Is = "MNU_WAREHOUSE_PERMISSION": ImagePath = MnuFolder & "warehouse_permission.png"
Case Is = "MNU_WAREHOUSE_OPNAME": ImagePath = MnuFolder & "warehouse_opname.png"

Case Is = "MNU_KARYAWAN_AGAMA": ImagePath = MnuFolder & "Personalia\AGAMA.png"
Case Is = "MNU_KARYAWAN_BAGIAN": ImagePath = MnuFolder & "Personalia\BAGIAN.png"
Case Is = "MNU_KARYAWAN_DEPARTEMEN": ImagePath = MnuFolder & "Personalia\DEPARTEMEN.png"
Case Is = "MNU_KARYAWAN_JABATAN": ImagePath = MnuFolder & "Personalia\JABATAN.png"
Case Is = "MNU_KARYAWAN_PENDIDIKAN": ImagePath = MnuFolder & "Personalia\PENDIDIKAN.png"
Case Is = "MNU_KARYAWAN_HARI_KERJA": ImagePath = MnuFolder & "Personalia\HARI_KERJA.png"
Case Is = "MNU_KARYAWAN_JENIS_JAM_KERJA": ImagePath = MnuFolder & "Personalia\JENIS_JAM_KERJA.png"
Case Is = "MNU_KARYAWAN_JAM_KERJA": ImagePath = MnuFolder & "Personalia\JAM_KERJA.png"
Case Is = "MNU_KARYAWAN_DATA_KARYAWAN": ImagePath = MnuFolder & "Personalia\DATA_KARYAWAN.png"
Case Is = "MNU_KARYAWAN_KONTRAK_KERJA": ImagePath = MnuFolder & "Personalia\KONTRAK_KERJA.png"

Case Is = "MNU_PRESENSI_ABSEN": ImagePath = MnuFolder & "Personalia\ABSEN.png"
Case Is = "MNU_PRESENSI_IMPORT": ImagePath = MnuFolder & "Personalia\IMPORT.png"
Case Is = "MNU_PRESENSI_KARYAWAN": ImagePath = MnuFolder & "Personalia\PRESENSI.png"
Case Is = "MNU_PRESENSI_MANUAL": ImagePath = MnuFolder & "Personalia\MANUAL.png"
Case Is = "MNU_PRESENSI_PROSES": ImagePath = MnuFolder & "Personalia\PRESENSI_PROSES.png"

Case Is = "MNU_GAJI_KOMPONEN": ImagePath = MnuFolder & "Personalia\KOMPONEN.png"
Case Is = "MNU_GAJI_PPH": ImagePath = MnuFolder & "Personalia\PPH.png"
Case Is = "MNU_GAJI_PTKP": ImagePath = MnuFolder & "Personalia\PTKP.png"
Case Is = "MNU_GAJI_SETTING": ImagePath = MnuFolder & "Personalia\GAJI_SETTING.png"
Case Is = "MNU_GAJI_PROSES": ImagePath = MnuFolder & "Personalia\PROSES.png"
Case Is = "MNU_GAJI_JAMSOSTEK": ImagePath = MnuFolder & "Personalia\JAMSOSTEK.png"

Case Is = "MNU_LAPORAN_KARYAWAN": ImagePath = MnuFolder & "Personalia\KARYAWAN.png"
Case Is = "MNU_LAPORAN_GAJI": ImagePath = MnuFolder & "Personalia\GAJI.png"
Case Is = "MNU_LAPORAN_GAJI_MINGGUAN": ImagePath = MnuFolder & "Personalia\GAJI7.png"
Case Is = "MNU_LAPORAN_KONTRAK": ImagePath = MnuFolder & "Personalia\KONTRAK.png"
Case Is = "MNU_LAPORAN_KONTRAK_TEMPO": ImagePath = MnuFolder & "Personalia\TEMPO.png"
Case Is = "MNU_LAPORAN_PRESENSI_KARYAWAN": ImagePath = MnuFolder & "Personalia\PRESENSI_KARYAWAN.png"
Case Is = "MNU_LAPORAN_FORM_1721": ImagePath = MnuFolder & "Personalia\FORM21.png"

Case Is = "MNU_TOOLS_CALC": ImagePath = MnuFolder & "Personalia\CALC.png"

Case Is = "MNU_SETUP_OPTIONS": ImagePath = MnuFolder & "Personalia\OPTIONS.png"
Case Is = "MNU_SETUP_PASS": ImagePath = MnuFolder & "Personalia\PASS.png"
Case Is = "MNU_SETUP_PROFILE": ImagePath = MnuFolder & "Personalia\PROFILE.png"

'-------------------------FORM-----------------------
Case Is = "FRM_MAIN": ImagePath = IconFolder & "corp.ico"
Case Is = "FRM_HRD_KARYAWAN": ImagePath = IconFolder & "hrd_karyawan.ico"
Case Is = "FRM_HOME": ImagePath = IconFolder & "home.ico"
Case Is = "FRM_RO": ImagePath = IconFolder & "ro.ico"
Case Is = "FRM_PO": ImagePath = IconFolder & "po.ico"
Case Is = "FRM_USER": ImagePath = IconFolder & "user.ico"
Case Is = "FRM_InvEditSpl": ImagePath = IconFolder & "invedit.ico"
Case Is = "FRM_BERITA": ImagePath = IconFolder & "berita.ico"
Case Is = "FRM_WAREHOUSE_CARD": ImagePath = IconFolder & "warehouse_card.ico"
Case Is = "FRM_WAREHOUSE_RECEIVE": ImagePath = IconFolder & "warehouse_receive.ico"
Case Is = "FRM_WAREHOUSE_MONTHLY": ImagePath = IconFolder & "warehouse_monthly.ico"
Case Is = "FRM_WAREHOUSE_PERMISSION": ImagePath = IconFolder & "warehouse_permission.ico"
Case Is = "FRM_WAREHOUSE_OPNAME": ImagePath = IconFolder & "warehouse_opname.ico"
Case Is = "FRM_EDIT_BARANG": ImagePath = BackFolder & "frmBrg.jpg"
Case Is = "FRM_EDIT_SUPPLIER": ImagePath = BackFolder & "frmSpl.jpg"

'------------------------BACKGROUND------------------
Case Is = "FRM_MAIN_BACK": ImagePath = BackFolder & "back.gif"
Case Is = "SSS": ImagePath = BackFolder & "SSS.jpg"
Case Is = "ALG": ImagePath = BackFolder & "ALG.jpg"
Case Is = "POS": ImagePath = BackFolder & "POS.jpg"


'------------------------BUTTON------------------
Case Is = "BTN_SAVE": ImagePath = ButtonFolder & "save.gif"
Case Is = "BTN_ADD": ImagePath = ButtonFolder & "add.gif"
Case Is = "BTN_DEL": ImagePath = ButtonFolder & "del.gif"
Case Is = "BTN_EXP": ImagePath = ButtonFolder & "export.gif"
Case Is = "BTN_LS_USR": ImagePath = ButtonFolder & "listuser.gif"
Case Is = "BTN_CANCEL": ImagePath = ButtonFolder & "cancel.gif"
Case Is = "BTN_REFRESH": ImagePath = ButtonFolder & "refresh.gif"
Case Is = "BTN_WRITE": ImagePath = ButtonFolder & "write.gif"
Case Is = "BTN_NEW": ImagePath = ButtonFolder & "new.gif"
Case Is = "BTN_EDIT": ImagePath = ButtonFolder & "edit.gif"
Case Is = "BTN_SEARCH": ImagePath = ButtonFolder & "search.gif"
Case Is = "BTN_FORWARD": ImagePath = ButtonFolder & "forward.gif"
Case Is = "BTN_SETUP_OVER": ImagePath = ButtonFolder & "setup_over.jpg"
Case Is = "BTN_SETUP": ImagePath = ButtonFolder & "setup.jpg"
Case Is = "BTN_KONTAK_OVER": ImagePath = ButtonFolder & "kontak_over.jpg"
Case Is = "BTN_KONTAK": ImagePath = ButtonFolder & "kontak.jpg"
Case Is = "BTN_PRODUK_OVER": ImagePath = ButtonFolder & "produk_over.jpg"
Case Is = "BTN_PRODUK": ImagePath = ButtonFolder & "produk.jpg"
Case Is = "BTN_PENJUALAN_OVER": ImagePath = ButtonFolder & "penjualan_over.jpg"
Case Is = "BTN_PENJUALAN": ImagePath = ButtonFolder & "penjualan.jpg"
Case Is = "BTN_DAFTAR_TRANSAKSI_OVER": ImagePath = ButtonFolder & "daftar_transaksi_over.jpg"
Case Is = "BTN_DAFTAR_TRANSAKSI": ImagePath = ButtonFolder & "daftar_transaksi.jpg"
Case Is = "BTN_LAPORAN_OVER": ImagePath = ButtonFolder & "laporan_over.jpg"
Case Is = "BTN_LAPORAN": ImagePath = ButtonFolder & "laporan.jpg"

'------------------------PATH-----------
Case Is = "pathKar": ImagePath = KarFoto
End Select

End Function

