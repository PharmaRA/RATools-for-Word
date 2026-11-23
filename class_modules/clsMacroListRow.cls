VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMacroListRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================
' 宏管理器自定义列表行（内嵌于 Frame 容器）
'
' 替代视觉僵硬且存在底部 IntegralHeight 留白的原生活动 ListBox：
'   - 24pt 宽松行高 + 左侧 3pt 主题蓝选中指示条
'   - 悬停淡蓝底高亮 + 手型指针
'   - 支持单击选中与双击执行
' =============================================

Private Const ROW_HEIGHT As Single = 24
Private Const TEXT_HEIGHT As Single = 13.5

Private m_Bg As MSForms.Label
Private m_Bar As MSForms.Label
Private m_Text As MSForms.Label
Public WithEvents m_Hit As MSForms.Label
Attribute m_Hit.VB_VarHelpID = -1

Public RowIndex As Long
Public Owner As Object
Public IsSelected As Boolean

Private m_Hovered As Boolean

Public Sub Init(ByVal container As MSForms.Frame, _
                ByVal idx As Long, _
                ByVal titleText As String, _
                ByVal rowTop As Single, _
                ByVal rowWidth As Single, _
                ByVal hostForm As Object)
    Dim textW As Single

    RowIndex = idx
    Set Owner = hostForm
    IsSelected = False
    m_Hovered = False

    textW = rowWidth - 14
    If textW < 20 Then textW = rowWidth

    Set m_Bg = container.Controls.Add("Forms.Label.1", "rowBg_" & idx, True)
    With m_Bg
        .Caption = vbNullString
        .Move 4, rowTop, rowWidth, ROW_HEIGHT
        .BackStyle = fmBackStyleOpaque
        .BorderStyle = fmBorderStyleSingle
        .SpecialEffect = fmSpecialEffectFlat
    End With

    Set m_Bar = container.Controls.Add("Forms.Label.1", "rowBar_" & idx, True)
    With m_Bar
        .Caption = vbNullString
        .Move 4, rowTop + 2, 3, ROW_HEIGHT - 4
        .BackStyle = fmBackStyleOpaque
        .BorderStyle = fmBorderStyleNone
        .Visible = False
    End With

    Set m_Text = container.Controls.Add("Forms.Label.1", "rowTxt_" & idx, True)
    With m_Text
        .Caption = titleText
        .Move 12, rowTop + (ROW_HEIGHT - TEXT_HEIGHT) / 2, textW, TEXT_HEIGHT
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .WordWrap = False
        .Font.Name = ThemeDefaultFontName()
        .Font.Size = 9.5
    End With

    Set m_Hit = container.Controls.Add("Forms.Label.1", "rowHit_" & idx, True)
    With m_Hit
        .Caption = vbNullString
        .Move 4, rowTop, rowWidth, ROW_HEIGHT
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .ControlTipText = titleText
    End With

    On Error Resume Next
    m_Bg.ZOrder 0
    m_Bar.ZOrder 0
    m_Text.ZOrder 0
    m_Hit.ZOrder 0
    On Error GoTo 0

    ApplyVisualState
End Sub

Public Sub ResizeWidth(ByVal rowWidth As Single)
    Dim textW As Single

    If m_Bg Is Nothing Then Exit Sub
    textW = rowWidth - 14
    If textW < 20 Then textW = rowWidth

    On Error Resume Next
    m_Bg.Width = rowWidth
    m_Text.Width = textW
    m_Hit.Width = rowWidth
    On Error GoTo 0
End Sub

Public Sub SetSelected(ByVal selected As Boolean)
    If IsSelected = selected Then Exit Sub
    IsSelected = selected
    ApplyVisualState
End Sub

Public Sub ResetHover()
    If Not m_Hovered Then Exit Sub
    m_Hovered = False
    ApplyVisualState
End Sub

Public Sub ApplyVisualState()
    If m_Bg Is Nothing Or m_Bar Is Nothing Or m_Text Is Nothing Then Exit Sub

    On Error Resume Next
    If IsSelected Then
        m_Bg.BackColor = ThemeActiveFillColor()
        m_Bg.BorderColor = ThemeBorderColor()
        m_Bar.BackColor = ThemeAccentColor()
        m_Bar.Visible = True
        m_Text.ForeColor = ThemeTextPrimaryColor()
        m_Text.Font.Bold = True
    ElseIf m_Hovered Then
        m_Bg.BackColor = ThemeHoverFillColor()
        m_Bg.BorderColor = ThemeHoverFillColor()
        m_Bar.Visible = False
        m_Text.ForeColor = ThemeTextPrimaryColor()
        m_Text.Font.Bold = False
    Else
        m_Bg.BackColor = ThemeCardColor()
        m_Bg.BorderColor = ThemeCardColor()
        m_Bar.Visible = False
        m_Text.ForeColor = ThemeTextPrimaryColor()
        m_Text.Font.Bold = False
    End If
    On Error GoTo 0
End Sub

Public Sub Dispose()
    On Error Resume Next
    Set Owner = Nothing
    Set m_Hit = Nothing
    Set m_Text = Nothing
    Set m_Bar = Nothing
    Set m_Bg = Nothing
    On Error GoTo 0
End Sub

Private Sub m_Hit_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                            ByVal X As Single, ByVal Y As Single)
    SetHandCursor
    If Not m_Hovered Then
        m_Hovered = True
        On Error Resume Next
        If Not Owner Is Nothing Then Owner.OnMacroRowHover Me
        On Error GoTo 0
        ApplyVisualState
    End If
End Sub

Private Sub m_Hit_Click()
    If Owner Is Nothing Then Exit Sub
    On Error Resume Next
    Owner.OnMacroRowClick RowIndex
    On Error GoTo 0
End Sub

Private Sub m_Hit_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Owner Is Nothing Then Exit Sub
    On Error Resume Next
    Owner.OnMacroRowDblClick RowIndex
    On Error GoTo 0
End Sub
