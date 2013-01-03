Attribute VB_Name = "Common"
Option Explicit

Public Enum enWinSet
'  en_Copy = -1
    es_Show = 0
    es_Hide
    es_Close
    es_Move
    es_Size
End Enum


Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
'Public Declare Function GetDiskFreeSpaceEx Lib "kernel32.dll" Alias "GetDiskFreeSpaceExA" (ByVal lpDirectoryName As String, lpFreeBytesAvailableToCaller As ULARGE_INTEGER, lpTotalNumberOfBytes As ULARGE_INTEGER, lpTotalNumberOfFreeBytes As ULARGE_INTEGER) As Long
Public Declare Function SHGetDiskFreeSpace Lib "shell32" Alias "SHGetDiskFreeSpaceA" (ByVal pszVolume As String, pqwFreeCaller As Currency, pqwTot As Currency, pqwFree As Currency) As Long
Private Declare Function SetThreadExecutionState Lib "kernel32" (ByVal esFlags As Long) As Long

Public Const ES_CONTINUOUS = &H80000000
Public Const ES_DISPLAY_REQUIRED = &H2
Public Const ES_SYSTEM_REQUIRED = &H1
Public Const ES_AWAYMODE_REQUIRED = &H40

'Public Type ULARGE_INTEGER
'    LowPart  As Long
'    HighPart  As Long
'End Type

Public Type BitMapInfoHeader                                                   'tagBitMapInfoHeader Structure
    biSize As Long                                                              '
    biWidth As Long
    biHeight As Long                                                            'LONG DWORD
    biPlanes As Integer                                                         'WORD
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Public Type BitMapInfo
    bmiHeader As BitMapInfoHeader                                               '
    bmiColors As Byte                                                           'RGBQUAD
End Type

Public Const WS_BORDER = &H800000
Public Const WS_CAPTION = &HC00000
Public Const WS_SYSMENU = &H80000
Public Const WS_CHILD = &H40000000
Public Const WS_VISIBLE = &H10000000
Public Const WS_OVERLAPPED = &H0&
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_THICKFRAME = &H40000
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = 1
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80
Public Const HWND_TOP = 0                                                              'hWndInsertAfter 参数：Z序列的顶部
Public Const HWND_TOPMOST = -1                                                         '最前
Public Const HWND_NOTOPMOST = -2                                                       '不在最前
Public Const HWND_BOTTOM = 1                                                           '位于底层

Public Const WM_CLOSE = &H10
Public Const WM_KILLFOCUS = &H8

'获取硬盘剩余空间，注意返回值为Currency，要乘于10000才得到一个定点数，单位为byte
Public Function GetDiskFreeSpace(sDisk As String) As Currency
    Dim FreeCaller As Currency, Tot As Currency
    SHGetDiskFreeSpace sDisk, FreeCaller, Tot, GetDiskFreeSpace
    GetDiskFreeSpace = GetDiskFreeSpace
End Function

'重置系统IDLETIME，防止休眠或待机
Public Sub ResetIdleTime()
    SetThreadExecutionState (ES_SYSTEM_REQUIRED Or ES_DISPLAY_REQUIRED Or ES_CONTINUOUS)
End Sub

Public Sub PowerSaveOn()
    SetThreadExecutionState (ES_CONTINUOUS)
End Sub
