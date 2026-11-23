VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBatchMode
   Caption         =   "选择处理范围"
   ClientHeight    =   4800
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6840
   OleObjectBlob   =   "frmBatchMode.frx":0000
   StartUpPosition =   1
End
Attribute VB_Name = "frmBatchMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =============================================
' 公共批处理三模式选择弹窗（替代原生 InputBox）
' 返回值："1" / "2" / "3"，取消或关闭返回空串。
' =============================================

Private m_CurrentMode As String
Private m_Result As String
Private m_HoveredIndex As Long
Private m_HoverButtons As Collection

Private m_LblPrompt As MSForms.Label
Private m_Cards(1 To 3) As MSForms.Label
Private m_AccentBars(1 To 3) As MSForms.Label
Private m_BadgeBgs(1 To 3) As MSForms.Label
Private m_BadgeTexts(1 To 3) As MSForms.Label
Private m_Titles(1 To 3) As MSForms.Label
Private m_Descs(1 To 3) As MSForms.Label

Private WithEvents m_Hit1 As MSForms.Label
Private WithEvents m_Hit2 As MSForms.Label
Private WithEvents m_Hit3 As MSForms.Label
Private WithEvents m_BtnOk As MSForms.CommandButton
Private WithEvents m_BtnCancel As MSForms.CommandButton
Private m_FlatOk As clsFlatButton
Private m_FlatCancel As clsFlatButton

Private Sub UserForm_Initialize()
    Me.Caption = "选择处理范围"
    Me.Width = 368
    Me.Height = 264
    m_CurrentMode = "1"
    m_Result = vbNullString
    m_HoveredIndex = 0
    Set m_HoverButtons = New Collection

    BuildControls
    ApplyTheme
End Sub

Private Sub UserForm_Activate()
    ApplyTheme
    On Error Resume Next
    m_BtnOk.SetFocus
    On Error GoTo 0
End Sub

Public Sub InitDialog(ByVal dialogTitle As String, _
                      Optional ByVal defaultMode As String = "1")
    If Len(Trim$(dialogTitle)) > 0 Then Me.Caption = dialogTitle
    Select Case Trim$(defaultMode)
        Case "1", "2", "3"
            m_CurrentMode = Trim$(defaultMode)
        Case Else
            m_CurrentMode = "1"
    End Select
    m_Result = vbNullString
    ApplyTheme
End Sub

Public Property Get SelectedMode() As String
    SelectedMode = m_Result
End Property

Private Sub BuildControls()
    Dim titles As Variant
    Dim descs As Variant
    Dim i As Long
    Dim cardTop As Single
    Dim cardW As Single
    Dim cardH As Single
    Dim pad As Single
    Dim hit As MSForms.Label

    pad = 14
    cardW = Me.InsideWidth - pad * 2
    cardH = 44

    Set m_LblPrompt = Me.Controls.Add("Forms.Label.1", "lblPrompt", True)
    With m_LblPrompt
        .Caption = "请选择本次批处理要作用的文档范围："
        .Move pad, 12, cardW, 15
        .BackStyle = fmBackStyleTransparent
        .Font.Name = ThemeDefaultFontName()
        .Font.Size = 9
    End With

    titles = Array("仅当前文档", "多文件模式", "文件夹模式")
    descs = Array( _
        "处理当前处于活动状态的 Word 文档", _
        "弹出文件选择框，选取一个或多个 Word 文档", _
        "选定目录后递归处理内部全部 Word 文档")

    For i = 1 To 3
        cardTop = 34 + (i - 1) * (cardH + 8)

        Set m_Cards(i) = Me.Controls.Add("Forms.Label.1", "lblCard" & i, True)
        With m_Cards(i)
            .Caption = vbNullString
            .Move pad, cardTop, cardW, cardH
            .BackStyle = fmBackStyleOpaque
            .BorderStyle = fmBorderStyleSingle
            .SpecialEffect = fmSpecialEffectFlat
        End With

        Set m_AccentBars(i) = Me.Controls.Add("Forms.Label.1", "lblAccentBar" & i, True)
        With m_AccentBars(i)
            .Caption = vbNullString
            .Move pad, cardTop + 1, 3, cardH - 2
            .BackStyle = fmBackStyleOpaque
            .BorderStyle = fmBorderStyleNone
            .Visible = False
        End With

        Set m_BadgeBgs(i) = Me.Controls.Add("Forms.Label.1", "lblBadgeBg" & i, True)
        With m_BadgeBgs(i)
            .Caption = vbNullString
            .Move pad + 12, cardTop + 12, 20, 20
            .BackStyle = fmBackStyleOpaque
            .BorderStyle = fmBorderStyleSingle
            .SpecialEffect = fmSpecialEffectFlat
        End With

        Set m_BadgeTexts(i) = Me.Controls.Add("Forms.Label.1", "lblBadgeTxt" & i, True)
        With m_BadgeTexts(i)
            .Caption = CStr(i)
            .Move pad + 12, cardTop + 15.5, 20, 13
            .BackStyle = fmBackStyleTransparent
            .BorderStyle = fmBorderStyleNone
            .TextAlign = fmTextAlignCenter
            .Font.Name = ThemeDefaultFontName()
            .Font.Size = 9.5
            .Font.Bold = True
        End With

        Set m_Titles(i) = Me.Controls.Add("Forms.Label.1", "lblTitle" & i, True)
        With m_Titles(i)
            .Caption = titles(i - 1)
            .Move pad + 42, cardTop + 7.5, cardW - 52, 14
            .BackStyle = fmBackStyleTransparent
            .Font.Name = ThemeDefaultFontName()
            .Font.Size = 9.5
            .Font.Bold = True
        End With

        Set m_Descs(i) = Me.Controls.Add("Forms.Label.1", "lblDesc" & i, True)
        With m_Descs(i)
            .Caption = descs(i - 1)
            .Move pad + 42, cardTop + 23.5, cardW - 52, 13
            .BackStyle = fmBackStyleTransparent
            .Font.Name = ThemeDefaultFontName()
            .Font.Size = 8.5
        End With

        Set hit = Me.Controls.Add("Forms.Label.1", "lblHit" & i, True)
        With hit
            .Caption = vbNullString
            .Move pad, cardTop, cardW, cardH
            .BackStyle = fmBackStyleTransparent
            .BorderStyle = fmBorderStyleNone
            .ZOrder 0
        End With
        Select Case i
            Case 1: Set m_Hit1 = hit
            Case 2: Set m_Hit2 = hit
            Case 3: Set m_Hit3 = hit
        End Select
    Next i

    ' 保留离屏 Default / Cancel 原生按钮承接 Enter / Esc 快捷键
    Set m_BtnOk = Me.Controls.Add("Forms.CommandButton.1", "btnOk", True)
    With m_BtnOk
        .Caption = "继续"
        .Default = True
        .TabStop = False
        .Move -500, -500, 10, 10
    End With

    Set m_BtnCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel", True)
    With m_BtnCancel
        .Caption = "取消"
        .Cancel = True
        .TabStop = False
        .Move -500, -500, 10, 10
    End With

    Set m_FlatOk = CreateFlatButton(m_HoverButtons, Me, "flatOk", "继续", "primary", "ok", Me, 9)
    m_FlatOk.MoveTo Me.InsideWidth - pad - 74, Me.InsideHeight - 36, 74, 24

    Set m_FlatCancel = CreateFlatButton(m_HoverButtons, Me, "flatCancel", "取消", "secondary", "cancel", Me, 9)
    m_FlatCancel.MoveTo Me.InsideWidth - pad - 74 - 8 - 66, Me.InsideHeight - 36, 66, 24
End Sub

Public Sub ApplyTheme()
    On Error Resume Next
    ApplyFormBaseTheme Me
    m_LblPrompt.ForeColor = ThemeTextMutedColor()
    RefreshCardStyles
    RefreshHoverButtons m_HoverButtons
    ApplyFormWindowChrome Me.Caption
    On Error GoTo 0
End Sub

Private Sub RefreshCardStyles()
    Dim i As Long
    Dim isSelected As Boolean
    Dim isHovered As Boolean

    On Error Resume Next
    For i = 1 To 3
        isSelected = (CStr(i) = m_CurrentMode)
        isHovered = (i = m_HoveredIndex)

        With m_Cards(i)
            If isSelected Then
                .BackColor = ThemeActiveFillColor()
                .BorderColor = ThemeBorderColor()
            ElseIf isHovered Then
                .BackColor = ThemeHoverFillColor()
                .BorderColor = ThemeAccentColor()
            Else
                .BackColor = ThemeCardColor()
                .BorderColor = ThemeBorderColor()
            End If
        End With

        m_AccentBars(i).BackColor = ThemeAccentColor()
        m_AccentBars(i).Visible = isSelected

        With m_BadgeBgs(i)
            If isSelected Then
                .BackColor = ThemeAccentColor()
                .BorderColor = ThemeAccentColor()
            Else
                .BackColor = ThemeCanvasColor()
                .BorderColor = ThemeBorderColor()
            End If
        End With

        If isSelected Then
            m_BadgeTexts(i).ForeColor = ThemeAccentTextColor()
        Else
            m_BadgeTexts(i).ForeColor = ThemeTextMutedColor()
        End If

        m_Titles(i).ForeColor = ThemeTextPrimaryColor()
        m_Descs(i).ForeColor = ThemeTextMutedColor()
    Next i
    On Error GoTo 0
End Sub

Private Sub SelectMode(ByVal idx As Long)
    m_CurrentMode = CStr(idx)
    RefreshCardStyles
End Sub

Private Sub ConfirmMode(ByVal idx As Long)
    m_CurrentMode = CStr(idx)
    m_Result = m_CurrentMode
    Me.Hide
End Sub

Public Sub OnHoverControlChanged(ByVal activeItem As Object)
    ResetHoverButtons m_HoverButtons, activeItem
    If m_HoveredIndex <> 0 Then
        m_HoveredIndex = 0
        RefreshCardStyles
    End If
End Sub

Public Sub OnFlatButtonClick(ByVal actionKey As String)
    If actionKey = "ok" Then
        m_BtnOk_Click
    ElseIf actionKey = "cancel" Then
        m_BtnCancel_Click
    End If
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetHoverButtons m_HoverButtons
    If m_HoveredIndex <> 0 Then
        m_HoveredIndex = 0
        RefreshCardStyles
    End If
End Sub

Private Sub HoverCard(ByVal idx As Long)
    SetHandCursor
    ResetHoverButtons m_HoverButtons
    If m_HoveredIndex <> idx Then
        m_HoveredIndex = idx
        RefreshCardStyles
    End If
End Sub

Private Sub m_Hit1_Click(): SelectMode 1: End Sub
Private Sub m_Hit2_Click(): SelectMode 2: End Sub
Private Sub m_Hit3_Click(): SelectMode 3: End Sub

Private Sub m_Hit1_DblClick(ByVal Cancel As MSForms.ReturnBoolean): ConfirmMode 1: End Sub
Private Sub m_Hit2_DblClick(ByVal Cancel As MSForms.ReturnBoolean): ConfirmMode 2: End Sub
Private Sub m_Hit3_DblClick(ByVal Cancel As MSForms.ReturnBoolean): ConfirmMode 3: End Sub

Private Sub m_Hit1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverCard 1: End Sub
Private Sub m_Hit2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverCard 2: End Sub
Private Sub m_Hit3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverCard 3: End Sub

Private Sub m_BtnOk_Click()
    m_Result = m_CurrentMode
    Me.Hide
End Sub

Private Sub m_BtnCancel_Click()
    m_Result = vbNullString
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        m_Result = vbNullString
        Me.Hide
    Else
        ClearHoverButtons m_HoverButtons
    End If
End Sub
