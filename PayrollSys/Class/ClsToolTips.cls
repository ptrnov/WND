VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsToolTips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const IDOK = 1
Const IDCLOSE = 2
Const IDSITE = 3

Const ID_POPUP0 = 0
Const ID_POPUP1 = 1
Const ID_POPUP2 = 2
Public Sub ShowToolTips(ToolTip As XtremeSuiteControls.PopupControl, Header As String, Tgl As String, Detail As String, Optional FrmDef As Form)
    On Error Resume Next
       
    Dim x As Integer
    Dim lastPane As Integer
    
    lastPane = ID_POPUP0 'IIf(chkMultiplePopup, ID_POPUP2, ID_POPUP0)
    
    For x = ID_POPUP0 To lastPane
        'Dim Popup As XtremeSuiteControls.PopupControl
        'Set Popup = PopupControl(x)
        
        ToolTip.Animation = xtpPopupAnimationSlide
        ToolTip.AnimateDelay = 256
        ToolTip.ShowDelay = 10000
        ToolTip.Transparency = 200
        ToolTip.DefaultLocation = xtpPopupLocationNearTaskBar 'IIf(chkMultiplePopup, xtpPopupLocationBottomRight, xtpPopupLocationNearTaskBar)
        
        SetOffice2007Theme ToolTip, Header, Tgl, Detail
        Next
    
    Dim Control
    'For 'Each Control In FrmDef.Controls
     '   Control.Enabled = False
    'Next
    
   ToolTip.Show  'PopupControl(ID_POPUP0).Show
    
   ' If chkMultiplePopup Then
   '     PopupControl(ID_POPUP1).right = PopupControl(ID_POPUP0).right
   '     PopupControl(ID_POPUP1).bottom = (PopupControl(ID_POPUP0).bottom - PopupControl(ID_POPUP0).Height)
   '     PopupControl(ID_POPUP1).AnimateDelay = PopupControl(ID_POPUP0).AnimateDelay + 256
   '     PopupControl(ID_POPUP1).ShowDelay = PopupControl(ID_POPUP0).ShowDelay + 1000
    '    PopupControl(ID_POPUP1).Show
    '
    '    PopupControl(ID_POPUP2).right = PopupControl(ID_POPUP1).right
    '    PopupControl(ID_POPUP2).bottom = (PopupControl(ID_POPUP1).bottom - PopupControl(ID_POPUP1).Height)
    '    PopupControl(ID_POPUP2).AnimateDelay = PopupControl(ID_POPUP1).AnimateDelay + 256
    '    PopupControl(ID_POPUP2).ShowDelay = PopupControl(ID_POPUP1).ShowDelay + 1000
    ''    PopupControl(ID_POPUP2).Show
   ' End If

End Sub
Function LoadBitmap(Path As String) As Long
    LoadBitmap = LoadImage(App.hInstance, App.Path + "\Img\" + Path, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)
End Function

Function LoadIcon(Path As String, cx As Long, cy As Long) As Long
    LoadIcon = LoadImage(App.hInstance, App.Path + "\Img\" + Path, IMAGE_ICON, cx, cy, LR_LOADFROMFILE)
End Function
Sub SetOffice2007Theme(Popup As XtremeSuiteControls.PopupControl, Header As String, Tgl As String, Detail As String)
    SetOffice2003Theme Popup, Header, Tgl, Detail
    Popup.VisualTheme = xtpPopupThemeOffice2007
End Sub

Sub SetOffice2003Theme(Popup As XtremeSuiteControls.PopupControl, Header As String, Tgl As String, Detail As String)
    Dim Item As PopupControlItem
    
    Popup.RemoveAllItems
    Popup.Icons.RemoveAll
    
    Set Item = Popup.AddItem(50, 10, 200, 45, Header)
    Item.Bold = True
    
    Set Item = Popup.AddItem(50, 27, 270, 65, "Tanggal : " & Tgl)
    
    Set Item = Popup.AddItem(50, 45, 270, 65, Detail)

    Set Item = Popup.AddItem(50, 65, 170, 95, "More..." & vbCrLf & "Thanks.")
    Item.TextColor = RGB(0, 61, 178)
    Item.TextAlignment = DT_LEFT Or DT_WORDBREAK
    
    Set Item = Popup.AddItem(12, 20, 12, 27, "")
    Item.SetIcon LoadIcon("Icons\berita.ico", 32, 32), xtpPopupItemIconNormal
    
    Set Item = Popup.AddItem(250, 10, 266, 26, "")
    Item.SetIcon LoadIcon("Icons\ico00002.ico", 16, 16), xtpPopupItemIconNormal
    Item.Id = IDCLOSE
    Item.Button = True
    
    Set Item = Popup.AddItem(12, 65, 28, 81, "")
    Item.SetIcon LoadIcon("Icons\flag.ico", 16, 16), xtpPopupItemIconNormal
    Item.Button = True
    
    Set Item = Popup.AddItem(30, 65, 46, 81, "")
    Item.SetIcon LoadIcon("Icons\cross.ico", 16, 16), xtpPopupItemIconNormal
    Item.Button = True
    
   
    Popup.VisualTheme = xtpPopupThemeOffice2003
    Popup.SetSize 270, 100

End Sub
Sub SetToolTipTheme(Popup As XtremeSuiteControls.PopupControl, Header As String, Tgl As String, Detail As String)
    Dim Item As PopupControlItem
    
    Popup.RemoveAllItems
    Popup.Icons.RemoveAll
    
    Set Item = Popup.AddItem(0, 0, 220, 90, "", RGB(255, 255, 225), 0)
    
    Set Item = Popup.AddItem(20, 30, 200, 100, Detail)
    Item.TextAlignment = DT_CENTER Or DT_WORDBREAK
    Item.Id = IDSITE
    
    Set Item = Popup.AddItem(5, 0, 170, 25, Header)
    Item.TextAlignment = DT_SINGLELINE Or DT_VCENTER
    Item.Bold = True
    Item.Hyperlink = False
    
    Set Item = Popup.AddItem(220 - 20, 2, 220 - 2, 2 + 18, "")
    Item.SetIcons LoadBitmap("Icons\CloseTooltip.bmp"), 0, xtpPopupItemIconNormal Or xtpPopupItemIconSelected Or xtpPopupItemIconPressed
    Item.Id = IDCLOSE
    
    Popup.VisualTheme = xtpPopupThemeCustom
    Popup.SetSize 220, 90

End Sub
