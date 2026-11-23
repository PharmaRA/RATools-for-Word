VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFlatButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================
' 纯扁平无 3D 边框按钮（基于多层 Label 组合）
'
' 解决 MSForms 原生 CommandButton 无法消除 Win98 立体边框与黑边焦点框，
' 以及单层 Label 文字贴顶偏上、无法垂直居中的问题：
'   - m_Bg:   不透明填充底板 + 1px 细边框
'   - m_Icon: 可选矢量字体图标层（Segoe MDL2 Assets 等），高 DPI 锐利
'   - m_Text: 透明前景文字层，按字号精确垂直居中
'   - m_Hit:  顶层透明点击热区，捕获悬停 / 按压 / 点击事件
' =============================================

Private m_Bg As MSForms.Label
Private m_Icon As MSForms.Label
Private m_Text As MSForms.Label
Public WithEvents m_Hit As MSForms.Label
Attribute m_Hit.VB_VarHelpID = -1

Public Owner As Object
Public ActionKey As String
Public Role As String
Public Enabled As Boolean

Private m_Container As Object
Private m_BaseName As String
Private m_Hovered As Boolean
Private m_Pressed As Boolean
Private m_Visible As Boolean
Private m_FontSize As Single
Private m_CaptionText As String
Private m_HasIcon As Boolean
Private m_IconGlyph As String
Private m_IconFontName As String
Private m_IconSize As Single
Private m_Left As Single
Private m_Top As Single
Private m_Width As Single
Private m_Height As Single

Public Sub Init(ByVal container As Object, _
                ByVal baseName As String, _
                ByVal captionText As String, _
                ByVal buttonRole As String, _
                ByVal actionName As String, _
                ByVal hostForm As Object, _
                Optional ByVal fontSize As Single = 9)
    Set m_Container = container
    m_BaseName = baseName
    Set Owner = hostForm
    Role = LCase$(Trim$(buttonRole))
    ActionKey = actionName
    Enabled = True
    m_Visible = True
    m_Hovered = False
    m_Pressed = False
    m_FontSize = fontSize
    m_CaptionText = captionText
    m_HasIcon = False

    Set m_Bg = container.Controls.Add("Forms.Label.1", baseName & "_bg", True)
    With m_Bg
        .Caption = vbNullString
        .BackStyle = fmBackStyleOpaque
        .BorderStyle = fmBorderStyleSingle
        .SpecialEffect = fmSpecialEffectFlat
    End With

    Set m_Text = container.Controls.Add("Forms.Label.1", baseName & "_txt", True)
    With m_Text
        .Caption = captionText
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .TextAlign = fmTextAlignCenter
        .WordWrap = False
        .Font.Name = ThemeDefaultFontName()
        .Font.Size = m_FontSize
        .Font.Bold = (Role = "primary")
    End With

    Set m_Hit = container.Controls.Add("Forms.Label.1", baseName & "_hit", True)
    With m_Hit
        .Caption = vbNullString
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
    End With

    ApplyVisualState
End Sub

Public Sub SetIcon(ByVal glyph As String, _
                   Optional ByVal iconFont As String = "", _
                   Optional ByVal iconSize As Single = 8.5)
    Dim resolvedFont As String

    If m_Container Is Nothing Then Exit Sub
    resolvedFont = Trim$(iconFont)
    If Len(resolvedFont) = 0 Then resolvedFont = ThemeIconFontName()

    m_IconGlyph = glyph
    m_IconFontName = resolvedFont
    m_IconSize = iconSize
    m_HasIcon = (Len(glyph) > 0)

    If m_HasIcon And (m_Icon Is Nothing) Then
        Set m_Icon = m_Container.Controls.Add("Forms.Label.1", m_BaseName & "_ico", True)
        With m_Icon
            .BackStyle = fmBackStyleTransparent
            .BorderStyle = fmBorderStyleNone
            .TextAlign = fmTextAlignCenter
            .WordWrap = False
        End With
    End If

    If Not m_Icon Is Nothing Then
        On Error Resume Next
        m_Icon.Font.Name = m_IconFontName
        m_Icon.Font.Size = m_IconSize
        m_Icon.Caption = m_IconGlyph
        m_Icon.Visible = (m_Visible And m_HasIcon)
        On Error GoTo 0
    End If

    ApplyVisualState
    If m_Width > 0 And m_Height > 0 Then
        MoveTo m_Left, m_Top, m_Width, m_Height
    End If
End Sub

Public Sub MoveTo(ByVal leftPos As Single, _
                  ByVal topPos As Single, _
                  ByVal widthVal As Single, _
                  ByVal heightVal As Single)
    Dim textH As Single
    Dim textW As Single
    Dim iconH As Single
    Dim iconW As Single
    Dim gap As Single
    Dim estTextW As Single
    Dim groupW As Single
    Dim startX As Single

    m_Left = leftPos
    m_Top = topPos
    m_Width = widthVal
    m_Height = heightVal

    If m_Bg Is Nothing Or m_Text Is Nothing Or m_Hit Is Nothing Then Exit Sub

    On Error Resume Next
    textH = m_FontSize * 1.45
    If textH > heightVal - 2 Then textH = heightVal - 2
    If textH < 10 Then textH = 10

    m_Bg.Move leftPos, topPos, widthVal, heightVal
    m_Hit.Move leftPos, topPos, widthVal, heightVal

    If m_HasIcon And Not (m_Icon Is Nothing) Then
        iconH = m_IconSize * 1.35
        If iconH > heightVal - 2 Then iconH = heightVal - 2
        If iconH < 9 Then iconH = 9

        If Len(Trim$(m_CaptionText)) = 0 Then
            m_Icon.TextAlign = fmTextAlignCenter
            m_Icon.Move leftPos, topPos + (heightVal - iconH) / 2, widthVal, iconH
            m_Icon.Visible = m_Visible
            m_Text.Visible = False
        Else
            iconW = m_IconSize * 1.25
            If iconW < 9 Then iconW = 9
            gap = 3
            estTextW = EstimateCaptionWidth(m_CaptionText, m_FontSize, (Role = "primary"))
            groupW = iconW + gap + estTextW
            If groupW > widthVal - 6 Then
                startX = leftPos + 3
            Else
                startX = leftPos + (widthVal - groupW) / 2
            End If

            m_Icon.TextAlign = fmTextAlignCenter
            m_Icon.Move startX, topPos + (heightVal - iconH) / 2, iconW, iconH
            m_Icon.Visible = m_Visible

            textW = widthVal - (startX - leftPos) - iconW - gap - 2
            If textW < 8 Then textW = 8
            m_Text.TextAlign = fmTextAlignLeft
            m_Text.Move startX + iconW + gap, topPos + (heightVal - textH) / 2, textW, textH
            m_Text.Visible = m_Visible
        End If
    Else
        textW = widthVal - 4
        If textW < 4 Then textW = widthVal
        m_Text.TextAlign = fmTextAlignCenter
        m_Text.Move leftPos + 2, topPos + (heightVal - textH) / 2, textW, textH
        m_Text.Visible = m_Visible
        If Not m_Icon Is Nothing Then m_Icon.Visible = False
    End If

    m_Bg.ZOrder 0
    If Not m_Icon Is Nothing Then m_Icon.ZOrder 0
    m_Text.ZOrder 0
    m_Hit.ZOrder 0
    On Error GoTo 0
End Sub

Private Function EstimateCaptionWidth(ByVal textStr As String, _
                                      ByVal fontSize As Single, _
                                      ByVal isBold As Boolean) As Single
    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim total As Single

    For i = 1 To Len(textStr)
        ch = Mid$(textStr, i, 1)
        code = AscW(ch)
        If code < 0 Or code > 127 Then
            total = total + fontSize * 1.05
        ElseIf ch = " " Then
            total = total + fontSize * 0.3
        Else
            total = total + fontSize * 0.55
        End If
    Next i
    If isBold Then total = total * 1.05
    EstimateCaptionWidth = total
End Function

Public Sub SetCaption(ByVal captionText As String)
    m_CaptionText = captionText
    If m_Text Is Nothing Then Exit Sub
    On Error Resume Next
    m_Text.Caption = captionText
    On Error GoTo 0
    If m_Width > 0 And m_Height > 0 Then
        MoveTo m_Left, m_Top, m_Width, m_Height
    End If
End Sub

Public Sub SetEnabled(ByVal value As Boolean)
    Enabled = value
    If Not Enabled Then
        m_Hovered = False
        m_Pressed = False
    End If
    ApplyVisualState
End Sub

Public Sub SetVisible(ByVal value As Boolean)
    m_Visible = value
    On Error Resume Next
    If Not m_Bg Is Nothing Then m_Bg.Visible = value
    If Not m_Icon Is Nothing Then m_Icon.Visible = (value And m_HasIcon)
    If Not m_Text Is Nothing Then m_Text.Visible = (value And (Not m_HasIcon Or Len(Trim$(m_CaptionText)) > 0))
    If Not m_Hit Is Nothing Then m_Hit.Visible = value
    On Error GoTo 0
End Sub

Public Sub SetToolTip(ByVal tipText As String)
    If m_Hit Is Nothing Then Exit Sub
    On Error Resume Next
    m_Hit.ControlTipText = tipText
    On Error GoTo 0
End Sub

Public Sub ResetHover()
    If Not m_Hovered And Not m_Pressed Then Exit Sub
    m_Hovered = False
    m_Pressed = False
    ApplyVisualState
End Sub

Public Sub ApplyVisualState()
    Dim bg As Long
    Dim fg As Long
    Dim bd As Long

    If m_Bg Is Nothing Or m_Text Is Nothing Then Exit Sub

    ResolveColors bg, fg, bd

    On Error Resume Next
    If m_Bg.BackColor <> bg Then m_Bg.BackColor = bg
    If m_Bg.BorderColor <> bd Then m_Bg.BorderColor = bd
    If m_Text.ForeColor <> fg Then m_Text.ForeColor = fg
    If Not m_Icon Is Nothing Then
        If m_Icon.ForeColor <> fg Then m_Icon.ForeColor = fg
    End If
    On Error GoTo 0
End Sub

Private Sub ResolveColors(ByRef bg As Long, _
                          ByRef fg As Long, _
                          ByRef bd As Long)
    If Not Enabled Then
        bg = ThemeDisabledBgColor()
        fg = ThemeDisabledTextColor()
        bd = ThemeBorderColor()
        Exit Sub
    End If

    Select Case Role
        Case "primary"
            fg = ThemeAccentTextColor()
            If m_Pressed Then
                bg = ThemeAccentActiveColor()
            ElseIf m_Hovered Then
                bg = ThemeAccentHoverColor()
            Else
                bg = ThemeAccentColor()
            End If
            bd = bg

        Case "ghost"
            If m_Pressed Then
                bg = ThemeActiveFillColor()
                fg = ThemeTextPrimaryColor()
                bd = ThemeBorderColor()
            ElseIf m_Hovered Then
                bg = ThemeHoverFillColor()
                fg = ThemeTextPrimaryColor()
                bd = ThemeBorderColor()
            Else
                bg = ThemeCardColor()
                fg = ThemeTextMutedColor()
                bd = ThemeCardColor()
            End If

        Case "header"
            If m_Pressed Then
                bg = ThemeActiveFillColor()
                fg = ThemeTextPrimaryColor()
                bd = ThemeAccentColor()
            ElseIf m_Hovered Then
                bg = ThemeHoverFillColor()
                fg = ThemeTextPrimaryColor()
                bd = ThemeBorderColor()
            Else
                bg = ThemeCanvasColor()
                fg = ThemeTextMutedColor()
                bd = ThemeCanvasColor()
            End If

        Case "close"
            If m_Pressed Then
                bg = RGB(196, 15, 30)
                fg = RGB(255, 255, 255)
                bd = bg
            ElseIf m_Hovered Then
                bg = RGB(232, 17, 35)
                fg = RGB(255, 255, 255)
                bd = bg
            Else
                bg = ThemeCanvasColor()
                fg = ThemeTextMutedColor()
                bd = ThemeCanvasColor()
            End If

        Case Else
            fg = ThemeTextPrimaryColor()
            If m_Pressed Then
                bg = ThemeActiveFillColor()
                bd = ThemeAccentColor()
            ElseIf m_Hovered Then
                bg = ThemeHoverFillColor()
                bd = ThemeAccentColor()
            Else
                bg = ThemeCardColor()
                bd = ThemeBorderColor()
            End If
    End Select
End Sub

Public Sub Dispose()
    On Error Resume Next
    Set Owner = Nothing
    Set m_Container = Nothing
    Set m_Hit = Nothing
    Set m_Icon = Nothing
    Set m_Text = Nothing
    Set m_Bg = Nothing
    On Error GoTo 0
End Sub

Private Sub m_Hit_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                            ByVal X As Single, ByVal Y As Single)
    If Not Enabled Then Exit Sub
    SetHandCursor
    If Not m_Hovered Then
        m_Hovered = True
        On Error Resume Next
        If Not Owner Is Nothing Then Owner.OnHoverControlChanged Me
        On Error GoTo 0
        ApplyVisualState
    End If
End Sub

Private Sub m_Hit_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                            ByVal X As Single, ByVal Y As Single)
    If Button <> 1 Or Not Enabled Then Exit Sub
    m_Pressed = True
    ApplyVisualState
End Sub

Private Sub m_Hit_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, _
                          ByVal X As Single, ByVal Y As Single)
    If Not m_Pressed Then Exit Sub
    m_Pressed = False
    ApplyVisualState
End Sub

Private Sub m_Hit_Click()
    If Not Enabled Then Exit Sub
    If Owner Is Nothing Then Exit Sub
    On Error Resume Next
    Owner.OnFlatButtonClick ActionKey
    On Error GoTo 0
End Sub
