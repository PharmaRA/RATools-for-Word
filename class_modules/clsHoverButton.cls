VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHoverButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================
' 公共界面交互类：扁平按钮悬停与按压状态管理
'
' 为 MSForms Label / CommandButton 统一提供：
'   - 悬停底色与边框高亮（MouseMove）
'   - 按下瞬间深度反馈（MouseDown / MouseUp）
'   - 手型鼠标光标（Win32 SetHandCursor）
'   - 主题切换时通过 ApplyVisualState 即时重绘
' =============================================

Public WithEvents Lbl As MSForms.Label
Attribute Lbl.VB_VarHelpID = -1
Public WithEvents Btn As MSForms.CommandButton
Attribute Btn.VB_VarHelpID = -1

Public Owner As Object
Public Role As String

Private m_Hovered As Boolean
Private m_Pressed As Boolean

Public Sub AttachLabel(ByVal target As MSForms.Label)
    Set Lbl = target
    Set Btn = Nothing
End Sub

Public Sub AttachCommandButton(ByVal target As MSForms.CommandButton)
    Set Btn = target
    Set Lbl = Nothing
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
    Dim isEnabled As Boolean

    On Error Resume Next

    If Not Lbl Is Nothing Then
        isEnabled = Lbl.Enabled
    ElseIf Not Btn Is Nothing Then
        isEnabled = Btn.Enabled
    Else
        Exit Sub
    End If

    ResolveColors isEnabled, bg, fg, bd

    If Not Lbl Is Nothing Then
        If Lbl.BackColor <> bg Then Lbl.BackColor = bg
        If Lbl.ForeColor <> fg Then Lbl.ForeColor = fg
        If Lbl.BorderStyle = fmBorderStyleSingle Then
            If Lbl.BorderColor <> bd Then Lbl.BorderColor = bd
        End If
    ElseIf Not Btn Is Nothing Then
        If Btn.BackColor <> bg Then Btn.BackColor = bg
        If Btn.ForeColor <> fg Then Btn.ForeColor = fg
    End If

    On Error GoTo 0
End Sub

Private Sub ResolveColors(ByVal isEnabled As Boolean, _
                          ByRef bg As Long, _
                          ByRef fg As Long, _
                          ByRef bd As Long)
    If Not isEnabled Then
        bg = ThemeDisabledBgColor()
        fg = ThemeDisabledTextColor()
        bd = ThemeBorderColor()
        Exit Sub
    End If

    Select Case LCase$(Role)
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

        Case "toolbar"
            fg = ThemeTextPrimaryColor()
            If m_Pressed Then
                bg = ThemeToolbarTileActiveColor()
                bd = ThemeAccentColor()
            ElseIf m_Hovered Then
                bg = ThemeToolbarTileHoverColor()
                bd = ThemeAccentColor()
            Else
                bg = ThemeToolbarTileColor()
                bd = ThemeToolbarTileBorderColor()
            End If

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

Private Sub NotifyHoverChanged()
    If Owner Is Nothing Then Exit Sub
    On Error Resume Next
    Owner.OnHoverControlChanged Me
    On Error GoTo 0
End Sub

Private Sub Lbl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                          ByVal X As Single, ByVal Y As Single)
    If Lbl.Enabled Then SetHandCursor
    If Not m_Hovered Then
        m_Hovered = True
        NotifyHoverChanged
        ApplyVisualState
    End If
End Sub

Private Sub Lbl_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                          ByVal X As Single, ByVal Y As Single)
    If Button <> 1 Or Not Lbl.Enabled Then Exit Sub
    m_Pressed = True
    ApplyVisualState
End Sub

Private Sub Lbl_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, _
                        ByVal X As Single, ByVal Y As Single)
    If Not m_Pressed Then Exit Sub
    m_Pressed = False
    ApplyVisualState
End Sub

Private Sub Btn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                          ByVal X As Single, ByVal Y As Single)
    If Btn.Enabled Then SetHandCursor
    If Not m_Hovered Then
        m_Hovered = True
        NotifyHoverChanged
        ApplyVisualState
    End If
End Sub

Private Sub Btn_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                          ByVal X As Single, ByVal Y As Single)
    If Button <> 1 Or Not Btn.Enabled Then Exit Sub
    m_Pressed = True
    ApplyVisualState
End Sub

Private Sub Btn_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, _
                        ByVal X As Single, ByVal Y As Single)
    If Not m_Pressed Then Exit Sub
    m_Pressed = False
    ApplyVisualState
End Sub
