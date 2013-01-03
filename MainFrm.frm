VERSION 5.00
Begin VB.Form MainFrm 
   Caption         =   "Form1"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10425
   Icon            =   "MainFrm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   545
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   695
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox pic2 
      Height          =   375
      Left            =   5880
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Left            =   6120
      Top             =   1200
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   10515
      TabIndex        =   2
      Top             =   0
      Width           =   10575
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Command1"
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Image ImgRecording 
      Height          =   240
      Left            =   7560
      Picture         =   "MainFrm.frx":3BFA
      Top             =   1680
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'摄像头视频监控工具，提供以下功能：
'1.

Private m_hCapWin As Long, m_Recording As Boolean, m_SaveDir As String, m_SaveFile As String
Private m_AutoSize As Boolean, m_AutoHide As Boolean, m_IsFullScreen As Boolean, m_TopMost As Boolean
Private m_Refresh As Boolean, m_Connected As Boolean, m_KeepAwake As Boolean
Private m_FrameRate As Long
Private m_MaxRecordMinutes As Long '最长记录时间，以分钟为单位，小于等于零则不检查记录时间
Private m_MinFreeDiskSpace As Long ' 最小保留磁盘空间，以兆字节为单位，小于等于零则不检查磁盘空间
Private m_HoursPerFile As Long   '多长时间分隔一个录像文件，单位为小时，小于等于0则不自动分隔文件
Private m_CheckRecordTimer As Long, m_FlashTrayIconCnt As Long, m_CheckDiskSpaceTimer As Long
Private m_KeepAwakeTimer As Long
Private m_CompressionClicked As Boolean

Private WithEvents m_Tray As cTray
Attribute m_Tray.VB_VarHelpID = -1

Private Const DEF_FRAME_RATE = 30 'FPS
Private Const DEF_MAX_RECORD_MINUTES = 120 '默认最长记录时间，2小时
Private Const DEF_MIN_FREE_DISK_SPACE = 0 '默认不检查剩余空间
Private Const DEF_HOURS_PER_FILE = 1 '默认每小时一个文件，整点分隔
Private Declare Sub InitCommonControls Lib "comctl32" ()

Private Sub StopRecord()
    m_Recording = False
    m_CheckRecordTimer = -1
    capCaptureStop m_hCapWin
    SetCaption " "
    m_Tray.SetTrayIcon Me.Icon
    m_Tray.SetTrayTip App.Title & vbCrLf & "摄像头监控录像工具"
End Sub

Private Sub StartRecord()
    Dim f As String, nDir As String, nF As String
    Dim nParms As CAPTUREPARMS
    
    '避免忘记设置压缩率
    If Not m_CompressionClicked Then
        Cmd "VideoCompression"
    End If
    
    nDir = GetSavePath()
    
    If Not MakePath(nDir) Then
        MsgBox "在指定的位置无法建立目录：" & vbCrLf & nDir, vbInformation, "保存视频文件"
        Exit Sub
    End If
    
    '生成文件名
    nF = GetSaveFile()
    
    f = JoinPathFile(nDir, nF)
    If CheckDirFile(f) = 1 Then
        If vbNo = MsgBox("文件已存在，覆盖此文件吗？" & vbCrLf & f, vbInformation + vbYesNo, "开始录像") Then Exit Sub
        On Error GoTo Cuo
        SetAttr f, 0
        Kill f
        On Error GoTo 0
    End If
    
    m_Recording = False
    SetWin m_hCapWin, es_Size, , , , 1
    m_Recording = True
    SetCaption "正在录像：" & nF
    ControlEnabled True
    DoEvents
    
    capCaptureGetSetup m_hCapWin, VarPtr(nParms), Len(nParms) '获取参数的设置
    
    If m_FrameRate <= 0 Then m_FrameRate = DEF_FRAME_RATE
    nParms.dwRequestMicroSecPerFrame = 1000000 / m_FrameRate  ' 捕捉帧率
    nParms.fYield = 1                                                        '用一个后台线程来进行视频捕捉
    nParms.fAbortLeftMouse = False                                              '关闭：单击鼠标左键停止录像的功能。
    nParms.fAbortRightMouse = False                                             '关闭：单击鼠标右键停止录像的功能
    nParms.fCaptureAudio = False        '不捕获音频
    
    capCaptureSetSetup m_hCapWin, VarPtr(nParms), Len(nParms)
    
    capFileSetCaptureFile m_hCapWin, f  '设置录像保存的文件
    
    capCaptureSequence m_hCapWin '开始捕捉
    
    m_CheckRecordTimer = 2
    m_Tray.SetTrayTip App.Title & vbCrLf & "正在录像中..."
    Exit Sub
Cuo:
    MsgBox "无法写文件：" & vbCrLf & vbCrLf & f, vbInformation, "录像 - 错误"
End Sub

'打开文本框，提示用户输入要保存的目录
Private Sub AskForDir()
    Dim nStr As String
    m_SaveDir = GetSavePath()
    nStr = "设置录像保存的文件夹。" & vbCrLf & "输入“<>”表示使用默认文件夹：" & vbCrLf & App.Path & "\videos"
    nStr = Trim(InputBox(nStr, "录像保存的文件夹", m_SaveDir))
    If Len(nStr) = 0 Then Exit Sub
    m_SaveDir = nStr
End Sub

'打开文本框，提示用户输入要保存的文件名
Private Sub AskForFile()
    Dim nStr As String, nF As String
    
    nF = String(255, " ")
    capFileGetCaptureFile m_hCapWin, VarPtr(nF), Len(nF)
    
    nF = GetStrLeft(nF, vbNullChar)
    
    If Trim(m_SaveFile) = "" Then m_SaveFile = "<>"
    nStr = "设置录像保存的文件名(不带路径)。" & vbCrLf & "输入“<>”表示使用默认文件名：日期-时间.扩展名"
    nStr = Trim(InputBox(nStr, "录像保存的文件名", m_SaveFile))
    If Len(nStr) = 0 Then Exit Sub
    m_SaveFile = nStr
End Sub

Private Sub AskForFrameRate()
    Dim nStr As String, nRate As Long
    nStr = Trim(InputBox("设置录像和预览的帧率FPS：", "帧率", m_FrameRate))
    If Len(nStr) = 0 Then Exit Sub
    If IsNumeric(nStr) Then
        nRate = CLng(nStr)
        If nRate > 0 Then
            m_FrameRate = nRate
        End If
    End If
End Sub

Private Sub AskForMaxRecordTime()
    Dim nStr As String, nTime As Long
    nStr = Trim(InputBox("设置录像的最长保留时间（分钟），小于等于零则不限时间。", "录像保留时间", m_MaxRecordMinutes))
    If Len(nStr) = 0 Then Exit Sub
    If IsNumeric(nStr) Then
        nTime = CLng(nStr)
        m_MaxRecordMinutes = nTime
    End If
End Sub

Private Sub AskForMinFreeDiskSpace()
    Dim nStr As String, nSpace As Long
    nStr = Trim(InputBox("设置磁盘的最小保留剩余空间（兆字节），小于等于零则不检查剩余空间。" & vbCrLf & Left$(GetSavePath(), 1) & _
        " 盘当前剩余空间：" & CLng(GetDiskFreeSpace(Left$(GetSavePath(), 3)) / 100) & "MB", "磁盘保留空间", m_MinFreeDiskSpace))
    If Len(nStr) = 0 Then Exit Sub
    If IsNumeric(nStr) Then
        nSpace = CLng(nStr)
        m_MinFreeDiskSpace = nSpace
    End If
End Sub

Private Sub AskForHoursPerFile()
    Dim nStr As String, nHours As Long
    nStr = Trim(InputBox("设置分隔录像文件的时间间隔（小时），小于等于零则不自动分隔文件。" & vbCrLf & "注意：软件仅在整点时分隔文件，所以第一个文件的长度小于等于设定的时间间隔。" _
        , "自动分隔文件", m_HoursPerFile))
    If Len(nStr) = 0 Then Exit Sub
    If IsNumeric(nStr) Then
        nHours = CLng(nStr)
        m_HoursPerFile = nHours
    End If
End Sub

Private Function GetStrLeft(nStr As String, Fu As String) As String
    '去掉 Fu 及后面的字符
    Dim s As Long
    s = InStr(nStr, Fu)
    If s > 0 Then GetStrLeft = Left(nStr, s - 1) Else GetStrLeft = nStr
End Function

Private Sub form_Initialize()
InitCommonControls

End Sub

Private Sub Form_Load()
    Dim W As Long, H As Long
    
    SetCaption ""
    
    Me.ScaleMode = 3
    Picture1.ScaleMode = 3
    Picture1.BorderStyle = 0
    Set Command1(0).Container = Picture1
    Set Check1(0).Container = Picture1
    
    m_MaxRecordMinutes = DEF_MAX_RECORD_MINUTES
    m_FrameRate = DEF_FRAME_RATE
    m_MinFreeDiskSpace = DEF_MIN_FREE_DISK_SPACE
    m_HoursPerFile = DEF_HOURS_PER_FILE
    m_TopMost = False
    m_IsFullScreen = False
    m_CheckRecordTimer = -1
    m_CompressionClicked = False
    
    ReadSaveSetting False                                                            '读取用户设置
    
    '装载数组控件
    AddControl Command1, "连", "Connect", "连接摄像头"
    AddControl Command1, "断", "DisConnect", "断开与摄像头的连接"
    AddControl Command1, "-"
    AddControl Command1, "源", "VideoSource", "选择：视频源"
    AddControl Command1, "格", "VideoFormat", "设置：视频格式，分辨率"
    AddControl Command1, "显", "VideoDisplay", "视频显示对话框。某些显卡不支持此功能。"
    AddControl Command1, "-"
    AddControl Command1, "夹", "AskForDir", "设置录像文件保存的文件夹。默认为主程序所在目录下的“videos”文件夹"
    AddControl Command1, "文", "AskForFile", "录像保存的文件名，默认为：时间-编号.扩展名"
    AddControl Command1, "压", "VideoCompression", "设置视频录像文件的压缩方式"
    AddControl Command1, "帧", "FrameRate", "设置录像帧率"
    AddControl Command1, "时", "MaxRecordTime", "设置最长保留录像时间"
    AddControl Command1, "剩", "MinFreeDiskSpace", "设置最小磁盘空间，磁盘空间小于此值则自动删除老记录"
    AddControl Command1, "割", "HoursPerFile", "设置录像文件的分隔时间"
    AddControl Command1, "-"
    AddControl Command1, "录", "Record", "开始录像"
    AddControl Command1, "停", "StopRecord", "停止录像"
    AddControl Command1, "图", "CopyImg", "将当前图像复制到剪贴板"
    AddControl Command1, "-"
    AddControl Command1, "全", "ToggleFullScreen", "切换：全屏/窗口"
    AddControl Command1, "关", "Exit", "关闭：退出程序"
    
    AddControl(Check1, "自", "AutoSize", "视频窗口是否随主窗口自动改变大小").Value = IIf(m_AutoSize, 1, 0)
    AddControl(Check1, "隐", "AutoHide", "最小化时自动隐藏主窗口").Value = IIf(m_AutoHide, 1, 0)
    AddControl(Check1, "醒", "KeepAwake", "防止系统进入待机或休眠").Value = IIf(m_KeepAwake, 1, 0)
    AddControl(Check1, "顶", "TopMost", "设置窗口置顶").Value = IIf(m_TopMost, 1, 0)
    
    ListControl Command1, Command1(0).Height * 0.1                                   '排列数组控件
    W = Command1(Command1.UBound).Left + Command1(Command1.UBound).Width * 2
    ListControl Check1, W                                                            '排列数组控件
    Picture1.Height = Command1(0).Height * 1.2
    
    m_Refresh = True
    
    CreateCapWin         '创建视频窗口
    
    ControlEnabled True
    
    Timer1.Interval = 600
    Timer1.Enabled = True
    
    Set m_Tray = New cTray
    m_Tray.AddTrayIcon pic2
    m_Tray.SetTrayIcon Me.Icon
    m_Tray.SetTrayTip App.Title & vbCrLf & "摄像头监控录像工具"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If m_Recording Then StopRecord
    Cmd "DisConnect"                                                            '断开摄像头连接
    SetWin m_hCapWin, es_Close
    ReadSaveSetting True                                                      '保存用户设置
    m_Tray.DelTrayIcon
End Sub

Private Sub Form_Resize()
    Picture1.Move 0, 0, Me.ScaleWidth, Command1(0).Height * 1.2
    If m_AutoSize Then SetWin m_hCapWin, es_Size                '视频子窗口随主窗口自动改变大小
    If m_AutoHide And Me.WindowState = vbMinimized Then Me.Hide
End Sub

Private Sub m_Tray_MouseClick(ByVal Button As Long, ByVal DBClick As Boolean)
    If DBClick = True Or Me.WindowState = vbMinimized Then
        ToggleWindowState
    End If
End Sub

'在全屏状态下，如果鼠标移动到屏幕顶端，则弹出工具栏
Private Sub Timer1_Timer()
    Dim nP As POINTAPI, x As Long, y As Long, H As Long
    Dim dNow As Date, nHour As Long, nMinute As Long, nSecond As Long
    Dim nStatus As CAPSTATUS, crFreeSpace As Currency
    
    '判断录像是否已经异常中止
    If m_CheckRecordTimer = 0 Then
        If capGetStatus(m_hCapWin, VarPtr(nStatus), Len(nStatus)) Then
            If nStatus.fCapturingNow = False And m_Recording Then
                Cmd "StopRecord"
                m_Tray.SetTrayMsgbox "录像过程已经中止，请检查中止原因！", NIIF_WARNING, "中止"
            End If
        End If
        m_CheckRecordTimer = 2
    ElseIf m_CheckRecordTimer > 0 Then
        m_CheckRecordTimer = m_CheckRecordTimer - 1
    End If
    
    If m_Recording Then
        '闪烁托盘图标
        If m_FlashTrayIconCnt > 0 Then
            m_Tray.SetTrayIcon Me.Icon
            m_FlashTrayIconCnt = 0
        Else
            m_Tray.SetTrayIcon ImgRecording.Picture
            m_FlashTrayIconCnt = 1
        End If
        
        '整点时才分隔文件，如果需要的话
        If m_HoursPerFile > 0 Then
            dNow = Now
            nMinute = Minute(dNow)
            nSecond = Second(dNow)
            If nMinute = 0 And nSecond = 0 Then  '整点
                nHour = Hour(dNow)
                If nHour Mod m_HoursPerFile = 0 Then
                    StopRecord
                    StartRecord
                    If m_MaxRecordMinutes > 0 Then '分隔文件的同时删除过期文件
                        DeleteExpiredFiles
                    End If
                End If
            End If
        End If
        
        '检查剩余磁盘空间
        If m_MinFreeDiskSpace > 0 Then
            m_CheckDiskSpaceTimer = m_CheckDiskSpaceTimer + 1
            If m_CheckDiskSpaceTimer > 101 Then
                m_CheckDiskSpaceTimer = 0
                crFreeSpace = GetDiskFreeSpace(Left$(GetSavePath(), 3))
                If CLng(crFreeSpace / 100) < m_MinFreeDiskSpace Then
                    DeleteOldestFile
                End If
            End If
        End If
    End If
    
    '阻止休眠和待机
    If m_KeepAwake Then
        m_KeepAwakeTimer = m_KeepAwakeTimer + 1
        If m_KeepAwakeTimer > 210 Then
            m_KeepAwakeTimer = 0
            ResetIdleTime
        End If
    End If
    
    '以下是处理全屏状态下的工具栏显示和隐藏
    If Not m_IsFullScreen Then Exit Sub
    
    GetCursorPos nP
    x = nP.x - Me.Left / Screen.TwipsPerPixelX
    y = nP.y - Me.Top / Screen.TwipsPerPixelY
    
    H = Me.Height / Screen.TwipsPerPixelY - Me.ScaleHeight                      '窗口标题栏高度
    If y > -1 And y < H + Picture1.Height Then
        If Picture1.Visible Then Exit Sub
        Picture1.Visible = True
    Else
        If Not Picture1.Visible Then Exit Sub
        Picture1.Visible = False
    End If
    SetWin m_hCapWin, es_Size
    
End Sub

Private Sub SetCaption(Optional nCap As String)
    If nCap <> "" Then Me.Tag = Trim(nCap)
    If m_IsFullScreen Then                                                        '全屏方式
        Me.Caption = ""
    Else                                                                        '窗口方式
        If Me.Tag = "" Then Me.Caption = "ArrozDVR" Else Me.Caption = "ArrozDVR - " & Me.Tag
    End If
End Sub

Private Sub Check1_Click(Index As Integer)
    Dim nTag As String, TF As Boolean
    
    If Not m_Refresh Then Exit Sub
    nTag = Check1(Index).Tag
    TF = Check1(Index).Value = 1
    Select Case LCase(nTag)
        Case LCase("AutoSize")
            m_AutoSize = TF
            SendMessage m_hCapWin, WM_CAP_SET_SCALE, m_AutoSize, 0                   '预览图像随窗口自动缩放
            Call SetWin(m_hCapWin, es_Size)
        Case LCase("AutoHide")
            m_AutoHide = TF
        Case LCase("KeepAwake")
            m_KeepAwake = TF
        Case LCase("TopMost")
            ToggleTopMost
    End Select
End Sub

Private Sub Command1_Click(Index As Integer)
    SendMessage Command1(Index).hWnd, WM_KILLFOCUS, 0, 0
    Cmd Command1(Index).Tag
End Sub

Private Sub Cmd(nCmd As String)
    Select Case LCase(nCmd)
        Case LCase("Connect"):
            CapConnect                             ' 连接摄像头
        Case LCase("DisConnect"):
            m_Connected = False
            capDriverDisconnect m_hCapWin '断开摄像头连接
        Case LCase("VideoSource"):
            capDlgVideoSource m_hCapWin '对话框：视频源
        Case LCase("VideoFormat"):
            capDlgVideoFormat m_hCapWin
            SetWin m_hCapWin, es_Size '显示对话框：视频格式,分辨率
        Case LCase("VideoDisplay"):
            capDlgVideoDisplay m_hCapWin '对话框：视频显示。某些显卡不支持？
        Case LCase("AskForDir"):
            AskForDir
        Case LCase("AskForFile"):
            AskForFile
        Case LCase("VideoCompression"):
            m_CompressionClicked = True
            capDlgVideoCompression m_hCapWin '对话框：视频压缩
        Case LCase("FrameRate"):
            AskForFrameRate
        Case LCase("MaxRecordTime"):
            AskForMaxRecordTime
        Case LCase("MinFreeDiskSpace"):
            AskForMinFreeDiskSpace
        Case LCase("HoursPerFile"):
            AskForHoursPerFile
        Case LCase("Record"):
            StartRecord
        Case LCase("StopRecord"):
            StopRecord
        
        Case LCase("CopyImg"):
            CaptureImg
        Case LCase("ToggleFullScreen"):
            ToggleFullScreen
        Case LCase("Exit"):
            Unload Me
            Exit Sub
    End Select
    
    ControlEnabled True
End Sub

'全屏切换
Public Sub ToggleFullScreen()
    m_IsFullScreen = Not m_IsFullScreen
    Picture1.Visible = Not m_IsFullScreen
    If m_IsFullScreen Then Me.BorderStyle = 0 Else Me.BorderStyle = 2
    Call SetCaption("")
    
    If m_IsFullScreen Then                                                        '全屏方式
        Me.WindowState = 2
        Check1(KjIndex(Check1, "AutoSize")).Value = 1                           '切换到：视频窗口随主窗口自动改变大小
    Else                                                                        '窗口方式
        Me.WindowState = 0
    End If
    Check1(KjIndex(Check1, "AutoSize")).Enabled = Not m_IsFullScreen
End Sub

Public Sub ToggleTopMost()
    If m_TopMost Then
        SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Else
        SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If
    m_TopMost = Not m_TopMost
End Sub

'获取视频的大小尺寸
Private Sub VideoSize(W As Long, H As Long)
    Dim nInf As BitMapInfo
    capGetVideoFormat m_hCapWin, VarPtr(nInf), Len(nInf)
    W = nInf.bmiHeader.biWidth
    H = nInf.bmiHeader.biHeight
End Sub

Private Function AddControl(Kj As Object, nCap As String, Optional nTag As String, Optional nNote As String) As Control
    '装载一个数组控件
    Dim i As Long
    
    i = Kj.UBound
    If Kj(i).Tag <> "" Then i = i + 1: Load Kj(i)
    On Error Resume Next
    Kj(i).Caption = nCap
    If nTag = "" Then Kj(i).Tag = Kj(i).Name & "-" & i Else Kj(i).Tag = nTag
    If Len(nNote) > 0 Then Kj(i).ToolTipText = nNote
    Set AddControl = Kj(i)
End Function

Private Sub ListControl(Kj As Object, L As Long)
    '排列数组控件
    Dim i As Long, H1 As Long, T As Long, W As Long
    
    H1 = Picture1.TextHeight("A"): T = H1 * 0.25: W = H1 * 2
    For i = Kj.lBound To Kj.UBound
        If Kj(i).Caption = "-" Then
            L = L + H1: Kj(i).Visible = False
        Else
            Kj(i).Move L, T, W, W: Kj(i).Visible = True
            L = L + W
        End If
    Next
End Sub

Private Function KjIndex(Kj As Object, nTag As String) As Long
    Dim i As Long
    For i = Kj.lBound To Kj.UBound
        If LCase(Kj(i).Tag) = LCase(nTag) Then KjIndex = i: Exit Function
    Next
    KjIndex = -1
End Function

Private Sub ControlEnabled(Optional nEnabled As Boolean)
    Dim Kj, TF As Boolean, nType As String
    On Error Resume Next
    For Each Kj In Me.Controls
        nType = LCase(TypeName(Kj))
        If nType = "commandbutton" Or nType = "checkbox" Then
            Kj.Enabled = nEnabled
        End If
    Next
    
    Command1(KjIndex(Command1, "ToggleFullScreen")).Enabled = True
    Command1(KjIndex(Command1, "Exit")).Enabled = True
    Check1(KjIndex(Check1, "AutoSize")).Enabled = Not m_IsFullScreen
    If Not nEnabled Then Exit Sub
    
    TF = m_Connected
    If m_Recording Then TF = False
    
    Command1(KjIndex(Command1, "Connect")).Enabled = Not TF
    Command1(KjIndex(Command1, "DisConnect")).Enabled = TF                      '按钮在摄像头连接状态才可用
    
    Command1(KjIndex(Command1, "VideoSource")).Enabled = TF
    Command1(KjIndex(Command1, "VideoFormat")).Enabled = TF
    Command1(KjIndex(Command1, "VideoDisplay")).Enabled = TF
    
    Command1(KjIndex(Command1, "VideoCompression")).Enabled = TF
    Command1(KjIndex(Command1, "Record")).Enabled = TF
    Command1(KjIndex(Command1, "StopRecord")).Enabled = TF
    Command1(KjIndex(Command1, "CopyImg")).Enabled = TF
    
    Command1(KjIndex(Command1, "FrameRate")).Enabled = Not m_Recording
    
    If Not m_Recording Then Exit Sub
    Command1(KjIndex(Command1, "Record")).Enabled = False
    Command1(KjIndex(Command1, "StopRecord")).Enabled = True
    Command1(KjIndex(Command1, "AskForFile")).Enabled = False
    Command1(KjIndex(Command1, "AskForDir")).Enabled = False
End Sub

Private Sub CreateCapWin()
    '创建视频窗口
    Dim nStyle As Long, s As Long
    Dim lpszName As String * 128
    Dim lpszVer As String * 128
    
    Do
        If capGetDriverDescriptionA(s, lpszName, 128, lpszVer, 128) = 0 Then Exit Do '获得驱动程序名称和版本信息
        s = s + 1
    Loop
    nStyle = WS_CHILD + WS_VISIBLE + WS_THICKFRAME  ' + WS_CAPTION  '子窗口+可见+标题栏+边框
    If m_hCapWin <> 0 Then Exit Sub
    m_hCapWin = capCreateCaptureWindow("myDVR", nStyle, 0, 0, 640, 480, Me.hWnd, 0)
    If m_hCapWin = 0 Then Exit Sub
    SetWin m_hCapWin, es_Move, 0, Command1(0).Top + Command1(0).Height + 3, 640, 480
    capSetCallbackOnError m_hCapWin, AddressOf MyErrorCallback
End Sub

'打开摄像头
Private Sub CapConnect()
    Dim d As Long
    d = capDriverConnect(m_hCapWin, 0)                      '连接一个视频驱动，成功返回真(1)
    capPreviewScale m_hCapWin, m_AutoSize                       '预览图像随窗口自动缩放
    capPreviewRate m_hCapWin, m_FrameRate                         '设置预览显示频率
    capPreview m_hCapWin, 1                              '第三个参数：1-预览模式有效,0-预览模式无效
    
    m_Connected = True
    SetWin m_hCapWin, es_Size                                               '调整视频窗口为正确的大小
End Sub

'设置窗口的状态
Private Sub SetWin(hWnd As Long, nSet As enWinSet, Optional ByVal nLeft As Long, Optional ByVal nTop As Long, Optional ByVal nWidth As Long, Optional ByVal nHeight As Long)
    Dim hWndZOrder As Long, wFlags As Long
    
    If hWnd = 0 Then Exit Sub
    Select Case nSet
    Case es_Close: SendMessage hWnd, WM_CLOSE, 0, 0: Exit Sub
    Case es_Hide: wFlags = SWP_NOMOVE + SWP_NOSIZE + SWP_NOZORDER + SWP_HIDEWINDOW '隐藏
    Case es_Show: hWndZOrder = HWND_TOP: wFlags = SWP_NOSIZE + SWP_SHOWWINDOW   '显示
    Case es_Move
        hWndZOrder = HWND_TOP: wFlags = SWP_NOACTIVATE + SWP_NOSIZE
    Case es_Size
        hWndZOrder = HWND_TOP: wFlags = SWP_NOACTIVATE
        If m_Recording Then wFlags = wFlags + SWP_NOSIZE '录像状态下改变视频窗口大小，有时会出现莫名其妙的错误
        
        nLeft = 0
        If Picture1.Visible Then nTop = Picture1.Height + 3
        If m_AutoSize Then
            nWidth = Me.ScaleWidth - nLeft
            nHeight = IIf(nHeight = 1, Me.ScaleHeight, Me.ScaleHeight - nTop)
        Else
            VideoSize nWidth, nHeight                                     '获取视频的实际大小
        End If
        If nWidth < 20 Or nHeight < 20 Then Exit Sub
    End Select
    
    SetWindowPos hWnd, hWndZOrder, nLeft, nTop, nWidth, nHeight, wFlags
End Sub

Private Sub CaptureImg()
    Clipboard.Clear
    capEditCopy m_hCapWin '将当前图像复制到剪贴板
    
    m_Tray.SetTrayMsgbox "图像已经复制到了剪贴板。", NIIF_NONE, "复制成功"
End Sub

'设置或读取用户配置信息
Private Sub ReadSaveSetting(IsSave As Boolean)
    Dim sTitle As String
    sTitle = App.Title
    If IsSave Then
        SaveSetting sTitle, "Setting", "AutoSize", m_AutoSize
        SaveSetting sTitle, "Setting", "AutoHide", m_AutoHide
        SaveSetting sTitle, "Setting", "KeepAwake", m_KeepAwake
        SaveSetting sTitle, "Setting", "SavePath", m_SaveDir
        SaveSetting sTitle, "Setting", "SaveFile", m_SaveFile
        SaveSetting sTitle, "Setting", "FrameRate", m_FrameRate
        SaveSetting sTitle, "Setting", "MaxRecordMinutes", m_MaxRecordMinutes
        SaveSetting sTitle, "Setting", "MinFreeDiskSpace", m_MinFreeDiskSpace
        SaveSetting sTitle, "Setting", "HoursPerFile", m_HoursPerFile
        SaveSetting sTitle, "Setting", "TopMost", m_TopMost
    Else
        m_AutoSize = GetSetting(sTitle, "Setting", "AutoSize", True)
        m_AutoHide = GetSetting(sTitle, "Setting", "AutoHide", False)
        m_KeepAwake = GetSetting(sTitle, "Setting", "KeepAwake", False)
        m_SaveDir = GetSetting(sTitle, "Setting", "SavePath", "")
        m_SaveFile = GetSetting(sTitle, "Setting", "SaveFile", "")
        m_FrameRate = GetSetting(sTitle, "Setting", "FrameRate", DEF_FRAME_RATE)
        m_MaxRecordMinutes = GetSetting(sTitle, "Setting", "MaxRecordMinutes", DEF_MAX_RECORD_MINUTES)
        m_MinFreeDiskSpace = GetSetting(sTitle, "Setting", "MinFreeDiskSpace", DEF_MIN_FREE_DISK_SPACE)
        m_HoursPerFile = GetSetting(sTitle, "Setting", "HoursPerFile", DEF_HOURS_PER_FILE)
        m_TopMost = GetSetting(sTitle, "Setting", "TopMost", False)
    End If
End Sub

'获取当前要保存的目录
Private Function GetSavePath() As String
    '如果路径不存在，用程序目录下子目录videos中使用默认文件名：年份-事件.avi
    GetSavePath = Trim(m_SaveDir)
    If Len(GetSavePath) = 0 Or GetSavePath = "<>" Or GetSavePath = "<默认>" Or GetSavePath = "<Default>" Then
        GetSavePath = JoinPathFile(App.Path, "videos\")
    End If
End Function

'获取当前要保存的文件名
Private Function GetSaveFile() As String
    GetSaveFile = Trim(m_SaveFile)
    If Len(GetSaveFile) = 0 Or GetSaveFile = "<>" Or GetSaveFile = "<默认>" Or GetSaveFile = "<Default>" Then
        GetSaveFile = Format(Now, "yyyy-mm-dd-hh_mm_ss") & ".avi"
    End If
    If InStr(GetSaveFile, ".") <= 0 Then GetSaveFile = GetSaveFile & ".avi"
End Function

'删除一段时间前的视频录制文件
Private Sub DeleteExpiredFiles()
    Dim sFiles() As String, i As Long, s As String, sTime As String, numFileDeleted As Long
    
    sFiles = SearchFiles(GetSavePath(), "????-??-??-??_??_??.avi") 'yyyy-mm-dd-hh_mm_ss
    
    For i = 0 To UBound(sFiles)
        s = sFiles(i)
        If LCase(Right$(s, 4)) = ".avi" Then
            s = Left$(Right$(s, 23), 19)
            sTime = Left$(s, 10) & " " & Replace(Right$(s, 8), "_", ":")
            
            If DateDiff("n", CDate(sTime), Now) > m_MaxRecordMinutes Then
                On Error Resume Next
                Kill sFiles(i)
                numFileDeleted = numFileDeleted + 1
                On Error GoTo 0
            End If
        End If
    Next
    
    If numFileDeleted > 0 Then
        m_Tray.SetTrayMsgbox "已经删除了 " & numFileDeleted & "个过期的录像文件。", NIIF_INFO
    End If
End Sub

'删除一个最老的文件
Private Sub DeleteOldestFile()
    Dim sFiles() As String, i As Long, s As String
    Dim sTime As String, nDiff As Long, nMaxDiff As Long, sOldestFile As String
    
    sFiles = SearchFiles(GetSavePath(), "????-??-??-??_??_??.avi") 'yyyy-mm-dd-hh_mm_ss
    If UBound(sFiles) <= 0 Then Exit Sub '至少有两个文件才可删除最老的
    
    For i = 0 To UBound(sFiles)
        s = sFiles(i)
        If LCase(Right$(s, 4)) = ".avi" Then
            s = Left$(Right$(s, 23), 19)
            sTime = Left$(s, 10) & " " & Replace(Right$(s, 8), "_", ":")
            
            nDiff = DateDiff("s", CDate(sTime), Now)
            If nDiff > nMaxDiff Then
                nMaxDiff = nDiff
                sOldestFile = sFiles(i)
            End If
        End If
    Next
    
    If Len(sOldestFile) Then
        On Error Resume Next
        Kill sOldestFile
        On Error GoTo 0
        'm_Tray.SetTrayMsgbox "已经成功删除了一个最老的文件。", NIIF_INFO, "删除文件", 1000
    End If
End Sub

Private Function ToggleWindowState()
    If Me.WindowState <> vbMinimized Then
        Me.WindowState = vbMinimized
        'Me.Hide
    Else
        Me.WindowState = vbNormal
        Me.Show
    End If
End Function

