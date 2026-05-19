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

Private Sub UserForm_Initialize()
    Me.Caption = "RATools"
    SetCaptionSafe "lblTitle", "RATools for Word"
    SetCaptionSafe "lblVersion", BuildVersionLabelPrefix() & GetAppVersion()
    SetCaptionSafe "lblRepoSection", BuildRepositorySectionTitle()
    SetCaptionSafe "lblUpdateSection", BuildUpdateSectionTitle()
    SetCaptionSafe "btnGitHub", "GitHub"
    SetCaptionSafe "btnGitee", "Gitee"
    SetCaptionSafe "btnCheckUpdate", BuildCheckUpdateButtonText()
    SetCaptionSafe "btnClose", BuildCloseButtonText()
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
