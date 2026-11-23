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
' 纯扁平无 3D 边框按钮（基于三层 Label 组合）
'
' 解决 MSForms 原生 CommandButton 无法消除 Win98 立体边框与黑边焦点框，
' 以及单层 Label 文字贴顶偏上、无法垂直居中的问题：
'   - m_Bg:   不透明填充底板 + 1px 细边框
'   - m_Text: 透明前景文字层，按字号精确垂直居中
'   - m_Hit:  顶层透明点击热区，捕获悬停 / 按压 / 点击事件
' =============================================

Private m_Bg As MSForms.Label
Private m_Text As MSForms.Label
Public WithEvents m_Hit As MSForms.Label
Attribute m_Hit.VB_VarHelpID = -1

Public Owner As Object
Public ActionKey As String
Public Role As String
Public Enabled As Boolean

Private m_Hovered As Boolean
Private m_Pressed As Boolean
Private m_FontSize As Single

Public Sub Init(ByVal container As Object, _
                ByVal baseName As String, _
                ByVal captionText As String, _
                ByVal buttonRole As String, _
                ByVal actionName As String, _
                ByVal hostForm As Object, _
                Optional ByVal fontSize As Single = 9)
    Set Owner = hostForm
    Role = LCase$(Trim$(buttonRole))
    ActionKey = actionName
    Enabled = True
    m_Hovered = False
    m_Pressed = False
    m_FontSize = fontSize

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

Public Sub MoveTo(ByVal leftPos As Single, _
                  ByVal topPos As Single, _
                  ByVal widthVal As Single, _
                  ByVal heightVal As Single)
    Dim textH As Single
    Dim textW As Single

    If m_Bg Is Nothing Or m_Text Is Nothing Or m_Hit Is Nothing Then Exit Sub

    On Error Resume Next
    textH = m_FontSize * 1.45
    If textH > heightVal - 2 Then textH = heightVal - 2
    If textH < 10 Then textH = 10

    textW = widthVal - 4
    If textW < 4 Then textW = widthVal

    m_Bg.Move leftPos, topPos, widthVal, heightVal
    m_Text.Move leftPos + 2, topPos + (heightVal - textH) / 2, textW, textH
    m_Hit.Move leftPos, topPos, widthVal, heightVal

    m_Bg.ZOrder 0
    m_Text.ZOrder 0
    m_Hit.ZOrder 0
    On Error GoTo 0
End Sub

Public Sub SetCaption(ByVal captionText As String)
    If m_Text Is Nothing Then Exit Sub
    On Error Resume Next
    m_Text.Caption = captionText
    On Error GoTo 0
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
    On Error Resume Next
    If Not m_Bg Is Nothing Then m_Bg.Visible = value
    If Not m_Text Is Nothing Then m_Text.Visible = value
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
    Set m_Hit = Nothing
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
