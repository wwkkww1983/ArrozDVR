VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*************************************************************************
'**模 块 名：cTray
'**说    明：设置托盘图标
'**创 建 人：马大哈 http://www.m5home.com/
'**日    期：2007年3月5日
'**版    本：V1.0
'**备    注：改自老外一个标准模块,封装成为一个类.
'**修 改 人：马大哈
'**日    期：2007年3月27日
'**备    注：添加了气泡提示窗口
'*************************************************************************
'使用方法
'Dim WithEvents mcTray As cTray
'Private Sub Form_Load()
'    Set mcTray = New cTray
'    With mcTray
'        .AddTrayIcon Pic1           '传送一个图片框
'        .SetTrayIcon Me.Icon        '传送一个图标
'        .SetTrayTip Me.Caption      '托盘提示文字
'    End With
'End Sub
'Private Sub Form_Unload(Cancel As Integer)
'    mcTray.DelTrayIcon           '删除托盘图标
'    Set mcTray = Nothing
'End Sub
'Private Sub mcTray_MouseClick(ByVal Button As Long, ByVal DBClick As Boolean)
'    '托盘事件返回
'    'Button: 鼠标的按键
'    'DBClick: T为双击,F为单击
'
'    Select Case Button
'        Case vbLeftButton           '左键单/双击
'            If DBClick = True Then Call ShowHideMe
'        Case vbRightButton          '右键单/双击
'            If DBClick = False Then PopupMenu mPop
'        Case vbMiddleButton         '中键单/双击
'            Debug.Print "vbMiddleButton " & IIf(DBClick = False, "Click", "DBClick")
'    End Select
'End Sub
'Private Sub mPops_Click(Index As Integer)
'    '菜单单击事件
'    Select Case Index
'        Case 0      'Show/Hide
'            Call ShowHideMe
'        Case 2      'About
'            MsgBox "托盘图标例子" & vbCrLf & vbCrLf & "   By 马大哈"
'        Case 3      'Exit
'            Unload Me
'    End Select
'End Sub
'Private Sub mcTray_TrayMsgBoxClick(ByVal CloseClick As Boolean)
'    '托盘气泡单击事件
'    If CloseClick = True Then
'        Debug.Print "                           点击关闭按钮"
'    Else
'        Debug.Print "                           点击气泡窗体"
'    End If
'End Sub
'Private Sub Cmd1_Click()
'    Dim I As Long, J As String
'
'    For I = 1 To 20
'        J = J & vbCrLf
'    Next I
'    mcTray.SetTrayMsgbox Space(35) & "测试气泡提示框" & J & Space(14) & "嘿嘿:)" & J & "     By 马大哈", NIIF_INFO, "提示"
'End Sub
'
'Private Function ShowHideMe()
'    If Me.WindowState = vbNormal Then
'        Me.WindowState = vbMinimized
'        Me.Hide
'    Else
'        Me.WindowState = vbNormal
'        Me.Show
'    End If
'End Function


Option Explicit

Private Const WM_MOUSEFIRST = &H200
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MOUSELAST = &H209
Private Const WM_USER = &H400
Private Const TRAY_CALLBACK = (WM_USER + 1001&)
Private Const PM_REMOVE = &H1
Private Const PM_NOREMOVE = &H0

Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_MESSAGE = &H1
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4

Private Const NOTIFYICON_VERSION = 3       'V5 style taskbar
Private Const NOTIFYICON_OLDVERSION = 0    'Win95 style taskbar

Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2

'Private Const NIIF_NONE = &H0
'Private Const NIIF_WARNING = &H2
'Private Const NIIF_ERROR = &H3
'Private Const NIIF_INFO = &H1
'Private Const NIIF_GUID = &H4

Public Enum NIIFENUM
    NIIF_NONE = &H0
    NIIF_WARNING = &H2
    NIIF_ERROR = &H3
    NIIF_INFO = &H1
    NIIF_GUID = &H4
End Enum

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 128
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeout As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type Msg
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Event MouseClick(ByVal Button As Long, ByVal DBClick As Boolean)
Public Event TrayMsgBoxClick(ByVal CloseClick As Boolean)

Dim TheData As NOTIFYICONDATA
Dim WithEvents cPic As PictureBox
Attribute cPic.VB_VarHelpID = -1
Dim TrayIco As StdPicture
Dim TrayTip As String

Public Sub DelTrayIcon()
    '删除托盘图标
    With TheData
        .uFlags = 0
    End With
    Shell_NotifyIcon NIM_DELETE, TheData
    DoEvents
End Sub

Public Sub AddTrayIcon(EventPic As PictureBox)
    '添加托盘图标
    On Error Resume Next
    
    Set cPic = EventPic
    With TheData
        .uID = vbNull
        .hWnd = cPic.hWnd
        .cbSize = Len(TheData)
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .dwState = 0
        .dwStateMask = 0
    End With
    Shell_NotifyIcon NIM_ADD, TheData
End Sub

Public Sub SetTrayTip(Tip As String)
    '设置提示信息
    With TheData
        .szTip = Tip & vbNullChar
        .uFlags = NIF_TIP
    End With
    Shell_NotifyIcon NIM_MODIFY, TheData
End Sub

Public Sub SetTrayIcon(Pic As StdPicture)
    '设置托盘图标
    If Pic.Type <> vbPicTypeIcon Then Exit Sub
    Set TrayIco = Pic
    With TheData
        .hIcon = TrayIco.Handle
        .uFlags = NIF_ICON
    End With
    Shell_NotifyIcon NIM_MODIFY, TheData
End Sub

Public Function SetTrayMsgbox(ByVal MsgInfo As String, Optional ByVal MsgFlags As NIIFENUM, Optional MsgTitle As String = "", Optional MsgTimeout As Long = 2000)
    '设置托盘气泡提示框
    With TheData
        .cbSize = Len(TheData)
        .hWnd = cPic.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = TrayIco.Handle
        .szTip = TrayTip & vbNullChar
        .dwState = 0
        .dwStateMask = 0
        .szInfo = MsgInfo & Chr(0)
        .szInfoTitle = MsgTitle & Chr(0)
        .dwInfoFlags = IIf(MsgTitle = "", NIIF_NONE, MsgFlags)
        .uTimeout = MsgTimeout
    End With
  
    Shell_NotifyIcon NIM_MODIFY, TheData
End Function

Private Sub cPic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    x = x / Screen.TwipsPerPixelX
    Select Case x
        Case WM_LBUTTONUP                   '左键抬起
            RaiseEvent MouseClick(vbLeftButton, False)
        Case WM_LBUTTONDBLCLK               '左键双击
            RaiseEvent MouseClick(vbLeftButton, True)
        Case WM_RBUTTONUP                   '右键抬起
            RaiseEvent MouseClick(vbRightButton, False)
        Case WM_RBUTTONDBLCLK               '右键双击
            RaiseEvent MouseClick(vbRightButton, True)
        Case WM_MBUTTONUP                   '中键抬起
            RaiseEvent MouseClick(vbMiddleButton, False)
        Case WM_MBUTTONDBLCLK               '中键双击
            RaiseEvent MouseClick(vbMiddleButton, True)
        Case 1028                           '点击托盘气泡关闭按钮
            RaiseEvent TrayMsgBoxClick(True)
        Case 1029                           '点击托盘气泡窗体本身
            RaiseEvent TrayMsgBoxClick(False)
    End Select
    'Debug.Print x
End Sub

