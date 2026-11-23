Attribute VB_Name = "UI_Theme"
Option Explicit

' =============================================
' 公共层：双主题配色与现代扁平控件样式（UI_Theme）
'
' 支持两套随 Word 自动切换的界面主题：
'   1) Word Fluent 浅色（默认）
'   2) 石墨深色（随 Word 深灰/黑色主题或 Windows 深色模式）
'
' 主题探测顺序：
'   HKCU\Software\Microsoft\Office\<Version>\Common\UI Theme
'     3 / 4 -> 深色（深灰 / 黑色）
'     0 / 5 / 7 -> 浅色（彩色 / 白色）
'     6 或未设置 -> 回退读取 Windows AppsUseLightTheme
' =============================================

Private Declare PtrSafe Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" ( _
    ByVal hKey As LongPtr, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, ByRef phkResult As LongPtr) As Long
Private Declare PtrSafe Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
    ByVal hKey As LongPtr, ByVal lpValueName As String, ByVal lpReserved As LongPtr, _
    ByRef lpType As Long, ByRef lpData As Long, ByRef lpcbData As Long) As Long
Private Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" ( _
    ByVal hKey As LongPtr) As Long

Private Const HKEY_CURRENT_USER As LongPtr = &H80000001
Private Const KEY_READ As Long = &H20019
Private Const REG_DWORD As Long = 4

Private Const DEFAULT_FONT_NAME As String = "Microsoft YaHei"
Private Const CODE_FONT_NAME As String = "Consolas"
Private Const ICON_FONT_NAME As String = "Segoe MDL2 Assets"

' 测试/预览覆盖值："" 表示自动跟随 Word，"light" / "dark" 表示强制主题
Private m_ThemeOverride As String

Public Sub SetThemeOverrideForTest(ByVal mode As String)
    m_ThemeOverride = LCase$(Trim$(mode))
End Sub

' 读取 HKCU 下的 REG_DWORD 注册表值
Private Function TryReadHkcuDword(ByVal subKey As String, _
                                  ByVal valueName As String, _
                                  ByRef outValue As Long) As Boolean
    Dim regHandle As LongPtr
    Dim valType As Long
    Dim valData As Long
    Dim dataSize As Long
    Dim rc As Long

    On Error GoTo ErrH
    rc = RegOpenKeyEx(HKEY_CURRENT_USER, subKey, 0&, KEY_READ, regHandle)
    If rc <> 0 Or regHandle = 0 Then Exit Function

    dataSize = 4
    rc = RegQueryValueEx(regHandle, valueName, 0, valType, valData, dataSize)
    RegCloseKey regHandle
    regHandle = 0

    If rc = 0 And valType = REG_DWORD Then
        outValue = valData
        TryReadHkcuDword = True
    End If
    Exit Function

ErrH:
    On Error Resume Next
    If regHandle <> 0 Then RegCloseKey regHandle
End Function

' 判断当前 Word 是否运行于深色外观
Public Function IsOfficeDarkTheme() As Boolean
    Dim officeTheme As Long
    Dim winLight As Long
    Dim foundOfficeTheme As Boolean
    Dim appVer As String

    If m_ThemeOverride = "dark" Then
        IsOfficeDarkTheme = True
        Exit Function
    ElseIf m_ThemeOverride = "light" Then
        IsOfficeDarkTheme = False
        Exit Function
    End If

    On Error Resume Next
    appVer = CStr(Application.Version)
    On Error GoTo 0

    If Len(appVer) > 0 Then
        foundOfficeTheme = TryReadHkcuDword("Software\Microsoft\Office\" & appVer & "\Common", _
                                            "UI Theme", officeTheme)
    End If
    If Not foundOfficeTheme Then
        foundOfficeTheme = TryReadHkcuDword("Software\Microsoft\Office\16.0\Common", _
                                            "UI Theme", officeTheme)
    End If
    If Not foundOfficeTheme Then
        foundOfficeTheme = TryReadHkcuDword("Software\Microsoft\Office\15.0\Common", _
                                            "UI Theme", officeTheme)
    End If

    If foundOfficeTheme Then
        Select Case officeTheme
            Case 3, 4
                IsOfficeDarkTheme = True
                Exit Function
            Case 0, 5, 7
                IsOfficeDarkTheme = False
                Exit Function
            Case 6
                ' 跟随 Windows 系统设置，继续向下查询 AppsUseLightTheme
            Case Else
                IsOfficeDarkTheme = False
                Exit Function
        End Select
    End If

    If TryReadHkcuDword("Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", _
                        "AppsUseLightTheme", winLight) Then
        IsOfficeDarkTheme = (winLight = 0)
    Else
        IsOfficeDarkTheme = False
    End If
End Function

' =============================================
' 双主题色板（Word Fluent 浅色 / 石墨深色）
' =============================================

Public Function ThemeCanvasColor() As Long
    If IsOfficeDarkTheme() Then
        ThemeCanvasColor = RGB(32, 34, 37)
    Else
        ThemeCanvasColor = RGB(245, 247, 250)
    End If
End Function

Public Function ThemeCardColor() As Long
    If IsOfficeDarkTheme() Then
        ThemeCardColor = RGB(43, 46, 51)
    Else
        ThemeCardColor = RGB(255, 255, 255)
    End If
End Function

Public Function ThemeBorderColor() As Long
    If IsOfficeDarkTheme() Then
        ThemeBorderColor = RGB(62, 66, 74)
    Else
        ThemeBorderColor = RGB(220, 225, 232)
    End If
End Function

Public Function ThemeAccentColor() As Long
    If IsOfficeDarkTheme() Then
        ThemeAccentColor = RGB(79, 156, 249)
    Else
        ThemeAccentColor = RGB(24, 90, 189)
    End If
End Function

Public Function ThemeAccentHoverColor() As Long
    If IsOfficeDarkTheme() Then
        ThemeAccentHoverColor = RGB(59, 130, 246)
    Else
        ThemeAccentHoverColor = RGB(18, 72, 154)
    End If
End Function

Public Function ThemeAccentActiveColor() As Long
    If IsOfficeDarkTheme() Then
        ThemeAccentActiveColor = RGB(37, 99, 235)
    Else
        ThemeAccentActiveColor = RGB(14, 58, 126)
    End If
End Function

Public Function ThemeAccentTextColor() As Long
    If IsOfficeDarkTheme() Then
        ThemeAccentTextColor = RGB(15, 23, 42)
    Else
        ThemeAccentTextColor = RGB(255, 255, 255)
    End If
End Function

Public Function ThemeHoverFillColor() As Long
    If IsOfficeDarkTheme() Then
        ThemeHoverFillColor = RGB(53, 59, 69)
    Else
        ThemeHoverFillColor = RGB(235, 243, 252)
    End If
End Function

Public Function ThemeActiveFillColor() As Long
    If IsOfficeDarkTheme() Then
        ThemeActiveFillColor = RGB(38, 59, 89)
    Else
        ThemeActiveFillColor = RGB(215, 231, 250)
    End If
End Function

Public Function ThemeTextPrimaryColor() As Long
    If IsOfficeDarkTheme() Then
        ThemeTextPrimaryColor = RGB(241, 245, 249)
    Else
        ThemeTextPrimaryColor = RGB(31, 41, 55)
    End If
End Function

Public Function ThemeTextMutedColor() As Long
    If IsOfficeDarkTheme() Then
        ThemeTextMutedColor = RGB(148, 163, 184)
    Else
        ThemeTextMutedColor = RGB(100, 116, 139)
    End If
End Function

Public Function ThemeDividerColor() As Long
    If IsOfficeDarkTheme() Then
        ThemeDividerColor = RGB(56, 60, 68)
    Else
        ThemeDividerColor = RGB(229, 233, 240)
    End If
End Function

Public Function ThemeDisabledBgColor() As Long
    If IsOfficeDarkTheme() Then
        ThemeDisabledBgColor = RGB(38, 41, 46)
    Else
        ThemeDisabledBgColor = RGB(238, 241, 245)
    End If
End Function

Public Function ThemeDisabledTextColor() As Long
    If IsOfficeDarkTheme() Then
        ThemeDisabledTextColor = RGB(100, 116, 139)
    Else
        ThemeDisabledTextColor = RGB(156, 163, 175)
    End If
End Function

' 悬浮工具栏图块底色：深色下维持浅色柔和图块，避免白底 BMP 图标与暗色描边冲突
Public Function ThemeToolbarTileColor() As Long
    If IsOfficeDarkTheme() Then
        ThemeToolbarTileColor = RGB(244, 246, 249)
    Else
        ThemeToolbarTileColor = RGB(255, 255, 255)
    End If
End Function

Public Function ThemeToolbarTileHoverColor() As Long
    If IsOfficeDarkTheme() Then
        ThemeToolbarTileHoverColor = RGB(215, 231, 250)
    Else
        ThemeToolbarTileHoverColor = RGB(235, 243, 252)
    End If
End Function

Public Function ThemeToolbarTileActiveColor() As Long
    If IsOfficeDarkTheme() Then
        ThemeToolbarTileActiveColor = RGB(191, 219, 254)
    Else
        ThemeToolbarTileActiveColor = RGB(215, 231, 250)
    End If
End Function

Public Function ThemeToolbarTileBorderColor() As Long
    If IsOfficeDarkTheme() Then
        ThemeToolbarTileBorderColor = RGB(62, 66, 74)
    Else
        ThemeToolbarTileBorderColor = RGB(220, 225, 232)
    End If
End Function

Public Function ThemeDefaultFontName() As String
    ThemeDefaultFontName = DEFAULT_FONT_NAME
End Function

Public Function ThemeCodeFontName() As String
    ThemeCodeFontName = CODE_FONT_NAME
End Function

Public Function ThemeIconFontName() As String
    ThemeIconFontName = ICON_FONT_NAME
End Function

' =============================================
' 窗体与控件统一扁平化样式助手
' =============================================

Public Sub ApplyFormBaseTheme(ByVal frm As Object)
    If frm Is Nothing Then Exit Sub
    On Error Resume Next
    frm.BackColor = ThemeCanvasColor()
    frm.ForeColor = ThemeTextPrimaryColor()
    frm.Font.Name = DEFAULT_FONT_NAME
    On Error GoTo 0
End Sub

Public Sub ApplyFormWindowChrome(ByVal windowCaption As String, _
                                 Optional ByVal applyCaptionColor As Boolean = True)
    On Error Resume Next
    StyleUserFormWindowChrome windowCaption, _
                              IsOfficeDarkTheme(), _
                              ThemeCanvasColor(), _
                              ThemeBorderColor(), _
                              ThemeTextPrimaryColor(), _
                              applyCaptionColor
    On Error GoTo 0
End Sub

Public Sub StyleFlatTextBox(ByVal txt As MSForms.TextBox, _
                            Optional ByVal fontSize As Single = 9.5)
    If txt Is Nothing Then Exit Sub
    On Error Resume Next
    With txt
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleNone
        .BackStyle = fmBackStyleOpaque
        .BackColor = ThemeCardColor()
        .ForeColor = ThemeTextPrimaryColor()
        .Font.Name = DEFAULT_FONT_NAME
        .Font.Size = fontSize
    End With
    On Error GoTo 0
End Sub

Public Sub StyleFlatInputBackdrop(ByVal backdrop As MSForms.Label, _
                                  Optional ByVal isFocused As Boolean = False)
    If backdrop Is Nothing Then Exit Sub
    On Error Resume Next
    With backdrop
        .Caption = vbNullString
        .BackStyle = fmBackStyleOpaque
        .BackColor = ThemeCardColor()
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        If isFocused Then
            .BorderColor = ThemeAccentColor()
        Else
            .BorderColor = ThemeBorderColor()
        End If
    End With
    On Error GoTo 0
End Sub

Public Sub StyleFlatListBox(ByVal lst As MSForms.ListBox, _
                            Optional ByVal fontSize As Single = 9.5)
    If lst Is Nothing Then Exit Sub
    On Error Resume Next
    With lst
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = ThemeBorderColor()
        .BackColor = ThemeCardColor()
        .ForeColor = ThemeTextPrimaryColor()
        .Font.Name = DEFAULT_FONT_NAME
        .Font.Size = fontSize
    End With
    On Error GoTo 0
End Sub

Public Sub StyleFlatComboBox(ByVal cbo As MSForms.ComboBox, _
                             Optional ByVal fontSize As Single = 9)
    If cbo Is Nothing Then Exit Sub
    On Error Resume Next
    With cbo
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = ThemeBorderColor()
        .BackColor = ThemeCardColor()
        .ForeColor = ThemeTextPrimaryColor()
        .Font.Name = DEFAULT_FONT_NAME
        .Font.Size = fontSize
    End With
    On Error GoTo 0
End Sub

Public Sub StyleFlatCard(ByVal ctrl As Object)
    If ctrl Is Nothing Then Exit Sub
    On Error Resume Next
    With ctrl
        .BackStyle = fmBackStyleOpaque
        .BackColor = ThemeCardColor()
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = ThemeBorderColor()
    End With
    On Error GoTo 0
End Sub

Public Sub StyleDividerLine(ByVal lineCtrl As MSForms.Label)
    If lineCtrl Is Nothing Then Exit Sub
    On Error Resume Next
    With lineCtrl
        .Caption = vbNullString
        .Height = 0.75
        .BackStyle = fmBackStyleOpaque
        .BackColor = ThemeDividerColor()
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleNone
    End With
    On Error GoTo 0
End Sub

Public Sub StylePrimaryCommandButton(ByVal btn As MSForms.CommandButton, _
                                     Optional ByVal fontSize As Single = 9)
    If btn Is Nothing Then Exit Sub
    On Error Resume Next
    With btn
        .BackStyle = fmBackStyleOpaque
        If .Enabled Then
            .BackColor = ThemeAccentColor()
            .ForeColor = ThemeAccentTextColor()
        Else
            .BackColor = ThemeDisabledBgColor()
            .ForeColor = ThemeDisabledTextColor()
        End If
        .Font.Name = DEFAULT_FONT_NAME
        .Font.Size = fontSize
        .Font.Bold = True
    End With
    On Error GoTo 0
End Sub

Public Sub StyleSecondaryCommandButton(ByVal btn As MSForms.CommandButton, _
                                       Optional ByVal fontSize As Single = 9)
    If btn Is Nothing Then Exit Sub
    On Error Resume Next
    With btn
        .BackStyle = fmBackStyleOpaque
        If .Enabled Then
            .BackColor = ThemeCardColor()
            .ForeColor = ThemeTextPrimaryColor()
        Else
            .BackColor = ThemeDisabledBgColor()
            .ForeColor = ThemeDisabledTextColor()
        End If
        .Font.Name = DEFAULT_FONT_NAME
        .Font.Size = fontSize
        .Font.Bold = False
    End With
    On Error GoTo 0
End Sub

Public Sub StyleLabelButton(ByVal lbl As MSForms.Label, _
                            ByVal isPrimary As Boolean, _
                            Optional ByVal fontSize As Single = 9)
    If lbl Is Nothing Then Exit Sub
    On Error Resume Next
    With lbl
        .BackStyle = fmBackStyleOpaque
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .TextAlign = fmTextAlignCenter
        .Font.Name = DEFAULT_FONT_NAME
        .Font.Size = fontSize
        .Font.Bold = isPrimary
        If Not .Enabled Then
            .BackColor = ThemeDisabledBgColor()
            .ForeColor = ThemeDisabledTextColor()
            .BorderColor = ThemeBorderColor()
        ElseIf isPrimary Then
            .BackColor = ThemeAccentColor()
            .ForeColor = ThemeAccentTextColor()
            .BorderColor = ThemeAccentColor()
        Else
            .BackColor = ThemeCardColor()
            .ForeColor = ThemeTextPrimaryColor()
            .BorderColor = ThemeBorderColor()
        End If
    End With
    On Error GoTo 0
End Sub

' 注册支持悬停/按压反馈的按钮并返回包装实例
Public Function RegisterHoverButton(ByVal store As Collection, _
                                    ByVal ctrl As Object, _
                                    ByVal role As String, _
                                    Optional ByVal hostForm As Object = Nothing) As clsHoverButton
    Dim item As clsHoverButton

    If ctrl Is Nothing Then Exit Function

    Set item = New clsHoverButton
    Set item.Owner = hostForm
    item.Role = LCase$(Trim$(role))

    If TypeOf ctrl Is MSForms.Label Then
        item.AttachLabel ctrl
    ElseIf TypeOf ctrl Is MSForms.CommandButton Then
        item.AttachCommandButton ctrl
    Else
        Exit Function
    End If

    item.ApplyVisualState
    If Not store Is Nothing Then store.Add item
    Set RegisterHoverButton = item
End Function

Public Function CreateFlatButton(ByVal store As Collection, _
                                 ByVal container As Object, _
                                 ByVal baseName As String, _
                                 ByVal captionText As String, _
                                 ByVal buttonRole As String, _
                                 ByVal actionKey As String, _
                                 ByVal hostForm As Object, _
                                 Optional ByVal fontSize As Single = 9) As clsFlatButton
    Dim btn As clsFlatButton

    Set btn = New clsFlatButton
    btn.Init container, baseName, captionText, buttonRole, actionKey, hostForm, fontSize
    If Not store Is Nothing Then store.Add btn
    Set CreateFlatButton = btn
End Function

Public Sub RefreshHoverButtons(ByVal store As Collection)
    Dim item As Object

    If store Is Nothing Then Exit Sub
    On Error Resume Next
    For Each item In store
        item.ApplyVisualState
    Next item
    On Error GoTo 0
End Sub

Public Sub ResetHoverButtons(ByVal store As Collection, _
                             Optional ByVal keepItem As Object = Nothing)
    Dim item As Object

    If store Is Nothing Then Exit Sub
    On Error Resume Next
    For Each item In store
        If Not (item Is keepItem) Then item.ResetHover
    Next item
    On Error GoTo 0
End Sub

Public Sub ClearHoverButtons(ByRef store As Collection)
    Dim item As Object

    If store Is Nothing Then Exit Sub
    On Error Resume Next
    For Each item In store
        If TypeName(item) = "clsFlatButton" Then
            item.Dispose
        Else
            Set item.Owner = Nothing
            Set item.Lbl = Nothing
            Set item.Btn = Nothing
        End If
    Next item
    Set store = Nothing
    On Error GoTo 0
End Sub
