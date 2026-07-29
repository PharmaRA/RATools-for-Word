Attribute VB_Name = "Mod_Core_Window"
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
' Word 之上，因此切换文档时由 clsAppEvents 调 RaiseUserFormWindow 抬升，
' Word 退出时卸载窗体保存位置。
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

' MSForms UserForm 的窗口类名
Public Const USERFORM_WINDOW_CLASS As String = "ThunderDFrame"

Private Const GWL_EXSTYLE As Long = -20
Private Const GWLP_HWNDPARENT As Long = -8
Private Const WS_EX_TOOLWINDOW As Long = &H80
Private Const WS_EX_APPWINDOW As Long = &H40000
Private Const SW_HIDE As Long = 0
Private Const SW_SHOWNA As Long = 8
Private Const HWND_TOP As LongPtr = 0
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOACTIVATE As Long = &H10

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

    DetachUserFormWindowFromOwner = (GetWindowLongCompat(windowHandle, GWLP_HWNDPARENT) = 0)
    Exit Function

ErrH:
    DetachUserFormWindowFromOwner = False
End Function

' 把窗体抬到 z 序最前但不抢焦点；owner 置 0 后靠它维持"悬浮"观感
Public Function RaiseUserFormWindow(ByVal windowCaption As String) As Boolean
    Dim windowHandle As LongPtr

    On Error GoTo ErrH

    windowHandle = FindUserFormWindow(windowCaption)
    If windowHandle = 0 Then Exit Function

    SetWindowPos windowHandle, HWND_TOP, 0, 0, 0, 0, _
                 SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
    RaiseUserFormWindow = True
    Exit Function

ErrH:
    RaiseUserFormWindow = False
End Function

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
