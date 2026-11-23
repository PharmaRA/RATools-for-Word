Attribute VB_Name = "Core_Window"
Option Explicit

' =============================================
' 公共层：窗口所有权与层级（Win32 封装）
'
' Word 2013+ 是 SDI：每个文档都是一个独立顶层窗口。modeless UserForm 默认由
' Show 那一刻的活动文档窗口 own，owner 一旦销毁（关掉该文档），窗体会被系统
' 一并销毁——这是悬浮工具栏"跟随文档而不是跟随 Word 进程"的根因。
'
' 对策：Show 之后把 owner 改为 0，让窗体成为不依附任何文档的顶层窗口；同时加
' WS_EX_TOOLWINDOW 使其不进 Alt+Tab 与任务栏。owner 置 0 后窗体不再自动压在
' Word 之上：Word 每次被激活都会把文档窗口抬到窗体前面，而工具窗既没有任务栏
' 按钮也不进 Alt+Tab，被盖住就等于消失、再也点不到。因此脱离 owner 的同时必须
' 把窗体钉进置顶层（见 DetachUserFormWindowFromOwner 与 KeepUserFormWindowOnTop），
' 文档切换时由 clsAppEvents 再确认一次，Word 退出时卸载窗体保存位置。
'
' 假定 VBA7（Office 2010+），与 frmMacroList 的 Win32 用法一致。
' =============================================

#If Win64 Then
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" ( _
        ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" ( _
        ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#End If

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function IsWindowVisible Lib "user32" ( _
    ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function SetWindowPos Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, _
    ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, _
    ByVal uFlags As Long) As Long
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
    ByVal hWndParent As LongPtr, ByVal hWndChildAfter As LongPtr, _
    ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function GetWindow Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal uCmd As Long) As LongPtr
Private Declare PtrSafe Function GetAncestor Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal gaFlags As Long) As LongPtr
Private Declare PtrSafe Function GetParent Lib "user32" ( _
    ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetWindowRect Lib "user32" ( _
    ByVal hWnd As LongPtr, ByRef windowRect As WINDOWRECT) As Long
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" ( _
    ByVal hWnd As LongPtr, ByRef processId As Long) As Long
Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" ( _
    ByVal hWnd As LongPtr, ByVal className As String, ByVal maxCount As Long) As Long
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare PtrSafe Function DwmSetWindowAttribute Lib "dwmapi.dll" ( _
    ByVal hWnd As LongPtr, ByVal dwAttribute As Long, _
    ByRef pvAttribute As Long, ByVal cbAttribute As Long) As Long
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" ( _
    ByVal hInstance As LongPtr, ByVal lpCursorName As LongPtr) As LongPtr
Private Declare PtrSafe Function SetCursor Lib "user32" ( _
    ByVal hCursor As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseCapture Lib "user32" () As Long
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hWnd As LongPtr, ByVal wMsg As Long, _
    ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Const MAX_TOOLTIP_WIDTH As Long = 400
Private Const MAX_TOOLTIP_HEIGHT As Long = 60

' MSForms UserForm 的窗口类名
Public Const USERFORM_WINDOW_CLASS As String = "ThunderDFrame"

Private Type WINDOWRECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const GWL_EXSTYLE As Long = -20
Private Const GW_OWNER As Long = 4
Private Const GA_ROOT As Long = 2
Private Const GA_ROOTOWNER As Long = 3

Private Const GWLP_HWNDPARENT As Long = -8
Private Const WS_EX_TOOLWINDOW As Long = &H80
Private Const WS_EX_APPWINDOW As Long = &H40000
Private Const WS_EX_TOPMOST As Long = &H8
Private Const SW_HIDE As Long = 0
Private Const SW_SHOWNA As Long = 8
Private Const HWND_TOPMOST As LongPtr = -1
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOACTIVATE As Long = &H10

' Win11 DWM 窗口外观属性（dwmapi.dll）与手型光标
Private Const DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1 As Long = 19
Private Const DWMWA_USE_IMMERSIVE_DARK_MODE As Long = 20
Private Const DWMWA_WINDOW_CORNER_PREFERENCE As Long = 33
Private Const DWMWA_BORDER_COLOR As Long = 34
Private Const DWMWA_CAPTION_COLOR As Long = 35
Private Const DWMWA_TEXT_COLOR As Long = 36
Private Const DWMWCP_ROUND As Long = 2
Private Const IDC_HAND As LongPtr = 32649
Private Const GWL_STYLE As Long = -16
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_BORDER As Long = &H800000
Private Const WS_SYSMENU As Long = &H80000
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_FRAMECHANGED As Long = &H20
Private Const WM_NCLBUTTONDOWN As Long = &HA1
Private Const HTCAPTION As LongPtr = 2

' 按类名+标题定位已显示的 UserForm 窗口；未显示或未找到返回 0
Public Function FindUserFormWindow(ByVal windowCaption As String) As LongPtr
    On Error GoTo ErrH
    FindUserFormWindow = FindWindow(USERFORM_WINDOW_CLASS, windowCaption)
    Exit Function

ErrH:
    FindUserFormWindow = 0
End Function

' 窗体窗口当前是否存在且可见
Public Function IsUserFormWindowVisible(ByVal windowCaption As String) As Boolean
    Dim windowHandle As LongPtr

    On Error GoTo ErrH
    windowHandle = FindUserFormWindow(windowCaption)
    If windowHandle = 0 Then Exit Function

    IsUserFormWindowVisible = (IsWindowVisible(windowHandle) <> 0)
    Exit Function

ErrH:
    IsUserFormWindowVisible = False
End Function

' 当前 owner 句柄；0 表示已脱离文档窗口（测试据此断言）
Public Function GetUserFormWindowOwner(ByVal windowCaption As String) As LongPtr
    Dim windowHandle As LongPtr

    On Error GoTo ErrH
    windowHandle = FindUserFormWindow(windowCaption)
    If windowHandle = 0 Then Exit Function

    GetUserFormWindowOwner = GetWindowLongCompat(windowHandle, GWLP_HWNDPARENT)
    Exit Function

ErrH:
    GetUserFormWindowOwner = 0
End Function

' 让窗体脱离文档窗口：owner 置 0，并从 Alt+Tab / 任务栏隐去。
' 必须在窗体已 Show（hwnd 已存在）之后调用；返回是否改写成功。
Public Function DetachUserFormWindowFromOwner(ByVal windowCaption As String) As Boolean
    Dim windowHandle As LongPtr
    Dim exStyle As LongPtr

    On Error GoTo ErrH

    windowHandle = FindUserFormWindow(windowCaption)
    If windowHandle = 0 Then Exit Function

    ' WS_EX_TOOLWINDOW 对任务栏按钮的影响由 shell 在窗口"隐藏->显示"时结算，
    ' 所以先隐藏再改样式；SW_SHOWNA 重新显示且不抢焦点。
    ShowWindow windowHandle, SW_HIDE

    SetWindowLongCompat windowHandle, GWLP_HWNDPARENT, 0

    exStyle = GetWindowLongCompat(windowHandle, GWL_EXSTYLE)
    exStyle = (exStyle Or WS_EX_TOOLWINDOW) And Not WS_EX_APPWINDOW
    SetWindowLongCompat windowHandle, GWL_EXSTYLE, exStyle

    ShowWindow windowHandle, SW_SHOWNA

    ' owner 置 0 的窗体不再被 Word 自动压在文档窗口之上，必须同时钉进置顶层，
    ' 否则 Word 一激活就把它盖住，而工具窗又没有任务栏按钮可以点回来。
    KeepUserFormWindowOnTop windowCaption

    DetachUserFormWindowFromOwner = (GetWindowLongCompat(windowHandle, GWLP_HWNDPARENT) = 0)
    Exit Function

ErrH:
    DetachUserFormWindowFromOwner = False
End Function

' 把窗体钉进置顶层并移到该层最前，且不抢焦点；owner 置 0 后靠它维持"悬浮"观感。
' Windows 只按 owner 关系或置顶层决定层级，owner 已经置 0，置顶就是唯一能让窗体
' 稳定停在 Word 之上的手段；代价是它同样浮在其他程序之上，不需要时从 Ribbon 隐藏。
' 本过程幂等，可在事件里反复调用以确认状态。
Public Function KeepUserFormWindowOnTop(ByVal windowCaption As String) As Boolean
    Dim windowHandle As LongPtr

    On Error GoTo ErrH

    windowHandle = FindUserFormWindow(windowCaption)
    If windowHandle = 0 Then Exit Function

    SetWindowPos windowHandle, HWND_TOPMOST, 0, 0, 0, 0, _
                 SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
    KeepUserFormWindowOnTop = IsUserFormWindowOnTop(windowCaption)
    Exit Function

ErrH:
    KeepUserFormWindowOnTop = False
End Function

' 窗体当前是否在置顶层（测试与诊断据此断言）
Public Function IsUserFormWindowOnTop(ByVal windowCaption As String) As Boolean
    Dim windowHandle As LongPtr

    On Error GoTo ErrH

    windowHandle = FindUserFormWindow(windowCaption)
    If windowHandle = 0 Then Exit Function

    IsUserFormWindowOnTop = _
        ((GetWindowLongCompat(windowHandle, GWL_EXSTYLE) And WS_EX_TOPMOST) <> 0)
    Exit Function

ErrH:
    IsUserFormWindowOnTop = False
End Function

' 悬浮窗钉入置顶层后，ControlTipText 的提示气泡是独立的 MSForms 窗口
' （类名 "F3 Tooltip ..."），默认不在置顶带，显示时会落在窗体层级下面。
' 把本窗体可见的气泡钉回置顶带；气泡隐藏时不动它，反复抬高由调用方的周期轮询完成。
Public Function EnsureTooltipOnTop(ByVal windowCaption As String) As Boolean
    Dim formHandle As LongPtr
    Dim tipHandle As LongPtr

    On Error GoTo ErrH

    formHandle = FindUserFormWindow(windowCaption)
    If formHandle = 0 Then Exit Function

    ' 类名带进程相关后缀且每次运行可能变化，只能枚举全部顶层窗口做前缀匹配。
    Dim scanCount As Long
    tipHandle = FindWindowEx(0, 0, vbNullString, vbNullString)
    Do While tipHandle <> 0
        ' 保险：限制扫描窗口数，防止任何异常导致的死循环
        scanCount = scanCount + 1
        If scanCount > 5000 Then Exit Do

        If IsTooltipOfForm(tipHandle, formHandle) Then
            If IsWindowVisible(tipHandle) <> 0 Then
                ' 注意：SetWindowPos(HWND_TOPMOST) 会改变窗口 Z 序，
                ' 继续 FindWindowEx 枚举会被打乱导致死循环，所以抬升后立即退出。
                SetWindowPos tipHandle, HWND_TOPMOST, 0, 0, 0, 0, _
                             SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
                Exit Do
            End If
        End If
        tipHandle = FindWindowEx(0, tipHandle, vbNullString, vbNullString)
    Loop
    EnsureTooltipOnTop = True
    Exit Function

ErrH:
    EnsureTooltipOnTop = False
End Function

' 仅匹配当前 Word 进程的 MSForms 提示窗口，再按窗体矩形筛选。
Private Function IsTooltipOfForm(ByVal tooltipHandle As LongPtr, _
                                 ByVal formHandle As LongPtr) As Boolean
    Dim tipRect As WINDOWRECT
    Dim formRect As WINDOWRECT
    Dim tooltipProcess As Long
    Dim formProcess As Long
    Dim className As String
    Dim classLength As Long

    On Error GoTo ErrH
    If GetWindowThreadProcessId(tooltipHandle, tooltipProcess) = 0 Then Exit Function
    If GetWindowThreadProcessId(formHandle, formProcess) = 0 Then Exit Function
    If tooltipProcess <> GetCurrentProcessId() Or formProcess <> tooltipProcess Then Exit Function
    className = String$(256, vbNullChar)
    classLength = GetClassName(tooltipHandle, className, Len(className))
    If classLength = 0 Then Exit Function
    className = Left$(className, classLength)
    If StrComp(Left$(className, Len("F3 Tooltip")), "F3 Tooltip", vbTextCompare) <> 0 Then Exit Function

    ' 气泡是小尺寸窗口；先按尺寸过滤掉窗体本身、容器与普通窗口。
    If GetWindowRect(tooltipHandle, tipRect) = 0 Then Exit Function
    If tipRect.Right - tipRect.Left > MAX_TOOLTIP_WIDTH Then Exit Function
    If tipRect.Bottom - tipRect.Top > MAX_TOOLTIP_HEIGHT Then Exit Function

    ' 气泡的 owner 是悬浮窗之外的窗口（实测），但显示时必与窗体矩形重叠；
    ' 已确认进程及提示类名后，用矩形相交进一步限定候选。
    If GetWindowRect(formHandle, formRect) = 0 Then Exit Function
    If tipRect.Right <= formRect.Left Then Exit Function
    If tipRect.Left >= formRect.Right Then Exit Function
    If tipRect.Bottom <= formRect.Top Then Exit Function
    If tipRect.Top >= formRect.Bottom Then Exit Function

    IsTooltipOfForm = True
    Exit Function

ErrH:
    IsTooltipOfForm = False
End Function

' 应用 Win11 原生圆角与主题同步标题栏/边框色（Win10 及更早版本静默忽略）
Public Function StyleUserFormWindowChrome(ByVal windowCaption As String, _
                                          ByVal isDark As Boolean, _
                                          ByVal captionColor As Long, _
                                          ByVal borderColor As Long, _
                                          ByVal textColor As Long, _
                                          Optional ByVal applyCaptionColor As Boolean = True) As Boolean
    Dim windowHandle As LongPtr
    Dim darkFlag As Long
    Dim cornerPref As Long
    Dim cCaption As Long
    Dim cBorder As Long
    Dim cText As Long

    On Error GoTo ErrH

    windowHandle = FindUserFormWindow(windowCaption)
    If windowHandle = 0 Then Exit Function

    darkFlag = IIf(isDark, 1&, 0&)
    cornerPref = DWMWCP_ROUND
    cCaption = captionColor
    cBorder = borderColor
    cText = textColor

    On Error Resume Next
    DwmSetWindowAttribute windowHandle, DWMWA_USE_IMMERSIVE_DARK_MODE, darkFlag, 4
    DwmSetWindowAttribute windowHandle, DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1, darkFlag, 4
    DwmSetWindowAttribute windowHandle, DWMWA_WINDOW_CORNER_PREFERENCE, cornerPref, 4
    DwmSetWindowAttribute windowHandle, DWMWA_BORDER_COLOR, cBorder, 4
    If applyCaptionColor Then
        DwmSetWindowAttribute windowHandle, DWMWA_CAPTION_COLOR, cCaption, 4
        DwmSetWindowAttribute windowHandle, DWMWA_TEXT_COLOR, cText, 4
    End If
    On Error GoTo ErrH

    StyleUserFormWindowChrome = True
    Exit Function

ErrH:
    StyleUserFormWindowChrome = False
End Function

' 在悬停扁平标签按钮时将鼠标指针切换为系统手型
Public Sub SetHandCursor()
    Dim handCursor As LongPtr

    On Error Resume Next
    handCursor = LoadCursor(0, IDC_HAND)
    If handCursor <> 0 Then SetCursor handCursor
    On Error GoTo 0
End Sub

' 去掉 UserForm 原生标题栏（保留 1px 边框以维持 Win11 DWM 圆角与边框色）
Public Function MakeUserFormWindowFrameless(ByVal windowCaption As String) As Boolean
    Dim windowHandle As LongPtr
    Dim styleVal As LongPtr

    On Error GoTo ErrH

    windowHandle = FindUserFormWindow(windowCaption)
    If windowHandle = 0 Then Exit Function

    styleVal = GetWindowLongCompat(windowHandle, GWL_STYLE)
    styleVal = (styleVal And Not CLngPtr(WS_CAPTION) And Not CLngPtr(WS_SYSMENU)) Or CLngPtr(WS_BORDER)
    SetWindowLongCompat windowHandle, GWL_STYLE, styleVal

    SetWindowPos windowHandle, 0, 0, 0, 0, 0, _
                 SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_FRAMECHANGED
    MakeUserFormWindowFrameless = True
    Exit Function

ErrH:
    MakeUserFormWindowFrameless = False
End Function

' 通过模拟标题栏命中实现无边框窗体平滑拖拽
Public Sub DragUserFormWindow(ByVal windowCaption As String)
    Dim windowHandle As LongPtr

    On Error Resume Next
    windowHandle = FindUserFormWindow(windowCaption)
    If windowHandle = 0 Then Exit Sub

    ReleaseCapture
    SendMessage windowHandle, WM_NCLBUTTONDOWN, HTCAPTION, 0
    On Error GoTo 0
End Sub

' ====== 32/64 位 GetWindowLong 差异收敛在此两个包装内 ======

Private Function GetWindowLongCompat(ByVal windowHandle As LongPtr, _
                                     ByVal fieldIndex As Long) As LongPtr
    #If Win64 Then
        GetWindowLongCompat = GetWindowLongPtr(windowHandle, fieldIndex)
    #Else
        GetWindowLongCompat = GetWindowLong(windowHandle, fieldIndex)
    #End If
End Function

Private Function SetWindowLongCompat(ByVal windowHandle As LongPtr, _
                                     ByVal fieldIndex As Long, _
                                     ByVal newValue As LongPtr) As LongPtr
    #If Win64 Then
        SetWindowLongCompat = SetWindowLongPtr(windowHandle, fieldIndex, newValue)
    #Else
        SetWindowLongCompat = SetWindowLong(windowHandle, fieldIndex, CLng(newValue))
    #End If
End Function
