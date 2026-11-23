VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAbout 
   Caption         =   "RATools"
   ClientHeight    =   3444
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4824
   OleObjectBlob   =   "frmAbout.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_BarAccent As MSForms.Label
Private m_LblRepoSection As MSForms.Label
Private m_LblUpdateSection As MSForms.Label
Private m_DivTop As MSForms.Label
Private m_DivBottom As MSForms.Label
Private m_HoverButtons As Collection
Private m_FlatGitHub As clsFlatButton
Private m_FlatGitee As clsFlatButton
Private m_FlatCheckUpdate As clsFlatButton
Private m_FlatClose As clsFlatButton

Private Sub UserForm_Initialize()
    Me.Caption = "RATools"
    Me.Width = 276
    Me.Height = 232
    Set m_HoverButtons = New Collection

    EnsureDecorativeControls

    SetCaptionSafe "lblTitle", "RATools for Word"
    SetCaptionSafe "lblVersion", BuildVersionLabelPrefix() & GetAppVersion()
    SetCaptionSafe "lblRepoSection", BuildRepositorySectionTitle()
    SetCaptionSafe "lblUpdateSection", BuildUpdateSectionTitle()
    SetCaptionSafe "btnGitHub", "GitHub"
    SetCaptionSafe "btnGitee", "Gitee"
    SetCaptionSafe "btnCheckUpdate", BuildCheckUpdateButtonText()
    SetCaptionSafe "btnClose", BuildCloseButtonText()

    ConfigureControls
    CreateFlatControls
    LayoutControls
    ApplyTheme
End Sub

Private Sub UserForm_Activate()
    ApplyTheme
End Sub

Private Sub EnsureDecorativeControls()
    Set m_BarAccent = EnsureLabel("barAccent")
    Set m_DivTop = EnsureLabel("divAboutTop")
    Set m_LblRepoSection = EnsureLabel("lblRepoSection")
    Set m_LblUpdateSection = EnsureLabel("lblUpdateSection")
    Set m_DivBottom = EnsureLabel("divAboutBottom")
End Sub

Private Function EnsureLabel(ByVal ctrlName As String) As MSForms.Label
    On Error Resume Next
    Set EnsureLabel = Me.Controls(ctrlName)
    On Error GoTo 0
    If EnsureLabel Is Nothing Then
        Set EnsureLabel = Me.Controls.Add("Forms.Label.1", ctrlName, True)
    End If
End Function

Private Function FindControl(ByVal ctrlName As String) As Object
    On Error Resume Next
    Set FindControl = Me.Controls(ctrlName)
    On Error GoTo 0
End Function

Private Sub ConfigureControls()
    Dim ctrl As Object
    Dim btnName As Variant

    On Error Resume Next
    Set ctrl = FindControl("lblTitle")
    If Not ctrl Is Nothing Then
        ctrl.BackStyle = fmBackStyleTransparent
        ctrl.Font.Name = ThemeDefaultFontName()
        ctrl.Font.Size = 13
        ctrl.Font.Bold = True
    End If

    Set ctrl = FindControl("lblVersion")
    If Not ctrl Is Nothing Then
        ctrl.BackStyle = fmBackStyleTransparent
        ctrl.Font.Name = ThemeDefaultFontName()
        ctrl.Font.Size = 9
    End If

    With m_LblRepoSection
        .BackStyle = fmBackStyleTransparent
        .Font.Name = ThemeDefaultFontName()
        .Font.Size = 8.5
        .Font.Bold = True
    End With

    With m_LblUpdateSection
        .BackStyle = fmBackStyleTransparent
        .Font.Name = ThemeDefaultFontName()
        .Font.Size = 8.5
        .Font.Bold = True
    End With

    ' 原生 CommandButton 移至离屏保留 Esc 语义与 FRX 完整
    For Each btnName In Array("btnGitHub", "btnGitee", "btnCheckUpdate", "btnClose")
        Set ctrl = FindControl(CStr(btnName))
        If Not ctrl Is Nothing Then
            ctrl.TabStop = False
            ctrl.Move -500, -500, 10, 10
        End If
    Next btnName

    Set ctrl = FindControl("btnClose")
    If Not ctrl Is Nothing Then ctrl.Cancel = True
    On Error GoTo 0
End Sub

Private Sub CreateFlatControls()
    Set m_FlatGitHub = CreateFlatButton(m_HoverButtons, Me, "flatGitHub", "GitHub", "secondary", "github", Me, 9)
    Set m_FlatGitee = CreateFlatButton(m_HoverButtons, Me, "flatGitee", "Gitee", "secondary", "gitee", Me, 9)
    Set m_FlatCheckUpdate = CreateFlatButton(m_HoverButtons, Me, "flatCheckUpdate", BuildCheckUpdateButtonText(), "secondary", "update", Me, 9)
    Set m_FlatClose = CreateFlatButton(m_HoverButtons, Me, "flatClose", BuildCloseButtonText(), "primary", "close", Me, 9)
End Sub

Private Sub LayoutControls()
    Dim pad As Single
    Dim innerW As Single
    Dim halfW As Single
    Dim ctrl As Object

    On Error Resume Next
    pad = 18
    innerW = Me.InsideWidth - pad * 2
    If innerW < 180 Then innerW = 180
    halfW = (innerW - 8) / 2

    m_BarAccent.Move 0, 0, Me.InsideWidth, 3

    Set ctrl = FindControl("lblTitle")
    If Not ctrl Is Nothing Then ctrl.Move pad, 14, innerW, 20

    Set ctrl = FindControl("lblVersion")
    If Not ctrl Is Nothing Then ctrl.Move pad, 36, innerW, 15

    m_DivTop.Move pad, 56, innerW, 0.75
    m_LblRepoSection.Move pad, 64, innerW, 14

    m_FlatGitHub.MoveTo pad, 82, halfW, 24
    m_FlatGitee.MoveTo pad + halfW + 8, 82, halfW, 24

    m_LblUpdateSection.Move pad, 115, innerW, 14
    m_FlatCheckUpdate.MoveTo pad, 133, innerW, 24

    m_DivBottom.Move pad, 166, innerW, 0.75
    m_FlatClose.MoveTo Me.InsideWidth - pad - 74, 174, 74, 24
    On Error GoTo 0
End Sub

Public Sub ApplyTheme()
    Dim ctrl As Object

    On Error Resume Next
    ApplyFormBaseTheme Me

    With m_BarAccent
        .Caption = vbNullString
        .BackStyle = fmBackStyleOpaque
        .BorderStyle = fmBorderStyleNone
        .BackColor = ThemeAccentColor()
    End With

    Set ctrl = FindControl("lblTitle")
    If Not ctrl Is Nothing Then ctrl.ForeColor = ThemeTextPrimaryColor()

    Set ctrl = FindControl("lblVersion")
    If Not ctrl Is Nothing Then ctrl.ForeColor = ThemeTextMutedColor()

    m_LblRepoSection.ForeColor = ThemeTextMutedColor()
    m_LblUpdateSection.ForeColor = ThemeTextMutedColor()

    StyleDividerLine m_DivTop
    StyleDividerLine m_DivBottom

    RefreshHoverButtons m_HoverButtons
    ApplyFormWindowChrome Me.Caption
    On Error GoTo 0
End Sub

Public Sub OnHoverControlChanged(ByVal activeItem As Object)
    ResetHoverButtons m_HoverButtons, activeItem
End Sub

Public Sub OnFlatButtonClick(ByVal actionKey As String)
    Select Case actionKey
        Case "github": btnGitHub_Click
        Case "gitee": btnGitee_Click
        Case "update": btnCheckUpdate_Click
        Case "close": btnClose_Click
    End Select
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetHoverButtons m_HoverButtons
End Sub

Private Sub btnGitHub_Click()
    OpenGitHubRepository
End Sub

Private Sub btnGitee_Click()
    OpenGiteeRepository
End Sub

Private Sub btnCheckUpdate_Click()
    CheckForUpdatesManually
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Function BuildVersionLabelPrefix() As String
    BuildVersionLabelPrefix = FromCodePoints(Array(24403, 21069, 29256, 26412, 65306))
End Function

Private Function BuildRepositorySectionTitle() As String
    BuildRepositorySectionTitle = FromCodePoints(Array(39033, 30446, 20179, 24211))
End Function

Private Function BuildUpdateSectionTitle() As String
    BuildUpdateSectionTitle = FromCodePoints(Array(26356, 26032))
End Function

Private Function BuildCheckUpdateButtonText() As String
    BuildCheckUpdateButtonText = FromCodePoints(Array(26816, 26597, 26356, 26032))
End Function

Private Function BuildCloseButtonText() As String
    BuildCloseButtonText = FromCodePoints(Array(20851, 38381))
End Function

Private Sub SetCaptionSafe(ByVal controlName As String, ByVal captionText As String)
    On Error Resume Next
    Me.Controls(controlName).Caption = captionText
    On Error GoTo 0
End Sub

Private Function FromCodePoints(ByVal values As Variant) As String
    Dim i As Long
    For i = LBound(values) To UBound(values)
        FromCodePoints = FromCodePoints & ChrW$(CLng(values(i)))
    Next i
End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ClearHoverButtons m_HoverButtons
End Sub
