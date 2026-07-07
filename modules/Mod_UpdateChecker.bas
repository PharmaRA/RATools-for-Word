Attribute VB_Name = "Mod_UpdateChecker"
Option Explicit

Private Const APP_VERSION As String = "v0.6.4"
Private Const RELEASES_API_URL As String = "https://api.github.com/repos/PharmaRA/RATools-for-Word/releases/latest"
Private Const GITHUB_RELEASE_URL_PREFIX As String = "https://github.com/PharmaRA/RATools-for-Word/releases/tag/"
Private Const GITEE_RELEASE_URL_PREFIX As String = "https://gitee.com/PharmaRA/RATools-for-Word/releases/tag/"
Private Const GITHUB_REPOSITORY_URL As String = "https://github.com/PharmaRA/RATools-for-Word"
Private Const GITEE_REPOSITORY_URL As String = "https://gitee.com/PharmaRA/RATools-for-Word"

Public Function GetAppVersion() As String
    GetAppVersion = APP_VERSION
End Function

Public Sub OpenGitHubRepository()
    OpenExternalPage GITHUB_REPOSITORY_URL
End Sub

Public Sub OpenGiteeRepository()
    OpenExternalPage GITEE_REPOSITORY_URL
End Sub

Public Sub CheckForUpdatesManually()
    CheckForUpdatesCore
End Sub

Private Sub CheckForUpdatesCore()
    Dim responseText As String
    Dim latestVersion As String
    Dim githubReleaseUrl As String
    Dim giteeReleaseUrl As String
    Dim compareResult As Long
    Dim userChoice As VbMsgBoxResult

    On Error GoTo ErrHandler

    If Not GetLatestReleaseJson(responseText) Then GoTo FetchFailed

    latestVersion = ExtractJsonStringValue(responseText, "tag_name")
    githubReleaseUrl = ExtractJsonStringValue(responseText, "html_url")

    If latestVersion = "" Or githubReleaseUrl = "" Then GoTo FetchFailed

    giteeReleaseUrl = BuildReleaseUrl(GITEE_RELEASE_URL_PREFIX, latestVersion)

    compareResult = CompareVersions(APP_VERSION, latestVersion)

    If compareResult <= 0 Then
        MsgBox BuildCurrentLatestMessage(), vbInformation, BuildUpdateTitle()
        Exit Sub
    End If

    userChoice = MsgBox(BuildUpdatePrompt(latestVersion), _
                        vbYesNoCancel + vbInformation, BuildNewVersionTitle())

    Select Case userChoice
        Case vbYes
            OpenReleasePage githubReleaseUrl
        Case vbNo
            OpenReleasePage giteeReleaseUrl
    End Select
    Exit Sub

FetchFailed:
    MsgBox BuildUpdateFailedMessage(), vbExclamation, BuildUpdateTitle()
    Exit Sub

ErrHandler:
    MsgBox BuildUpdateFailedMessage(), vbExclamation, BuildUpdateTitle()
End Sub

Private Function GetLatestReleaseJson(ByRef responseText As String) As Boolean
    Dim http As Object

    On Error GoTo ErrHandler

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", RELEASES_API_URL, False
    http.setRequestHeader "User-Agent", "RATools-for-Word"
    http.send

    If http.Status = 200 Then
        responseText = CStr(http.responseText)
        GetLatestReleaseJson = True
    End If

    Exit Function

ErrHandler:
    GetLatestReleaseJson = False
End Function

Private Function ExtractJsonStringValue(ByVal jsonText As String, ByVal keyName As String) As String
    Dim keyToken As String
    Dim keyPos As Long
    Dim colonPos As Long
    Dim startQuote As Long
    Dim endQuote As Long

    keyToken = Chr$(34) & keyName & Chr$(34)
    keyPos = InStr(1, jsonText, keyToken, vbTextCompare)
    If keyPos = 0 Then Exit Function

    colonPos = InStr(keyPos + Len(keyToken), jsonText, ":")
    If colonPos = 0 Then Exit Function

    startQuote = InStr(colonPos + 1, jsonText, Chr$(34))
    If startQuote = 0 Then Exit Function

    endQuote = FindJsonStringEnd(jsonText, startQuote + 1)
    If endQuote = 0 Then Exit Function

    ExtractJsonStringValue = Mid$(jsonText, startQuote + 1, endQuote - startQuote - 1)
End Function

Private Function FindJsonStringEnd(ByVal jsonText As String, ByVal startPos As Long) As Long
    Dim i As Long

    For i = startPos To Len(jsonText)
        If Mid$(jsonText, i, 1) = Chr$(34) Then
            If i = 1 Or Mid$(jsonText, i - 1, 1) <> "\" Then
                FindJsonStringEnd = i
                Exit Function
            End If
        End If
    Next i
End Function

Private Function NormalizeVersion(ByVal versionText As String) As String
    versionText = Trim$(versionText)
    If Len(versionText) > 0 Then
        If Left$(versionText, 1) = "v" Or Left$(versionText, 1) = "V" Then
            versionText = Mid$(versionText, 2)
        End If
    End If
    NormalizeVersion = versionText
End Function

Private Function GetVersionPart(ByVal versionText As String, ByVal index As Long) As Long
    Dim normalizedVersion As String
    Dim parts() As String

    normalizedVersion = NormalizeVersion(versionText)
    If Len(normalizedVersion) = 0 Then Exit Function

    parts = Split(normalizedVersion, ".")
    If index <= UBound(parts) Then
        If IsNumeric(parts(index)) Then
            GetVersionPart = CLng(parts(index))
        End If
    End If
End Function

Private Function CompareVersions(ByVal currentVersion As String, ByVal latestVersion As String) As Long
    Dim i As Long
    Dim currentPart As Long
    Dim latestPart As Long

    For i = 0 To 2
        currentPart = GetVersionPart(currentVersion, i)
        latestPart = GetVersionPart(latestVersion, i)

        If latestPart > currentPart Then
            CompareVersions = 1
            Exit Function
        ElseIf latestPart < currentPart Then
            CompareVersions = -1
            Exit Function
        End If
    Next i
End Function

Private Sub OpenReleasePage(ByVal releaseUrl As String)
    OpenExternalPage releaseUrl
End Sub

Private Function BuildReleaseUrl(ByVal releaseUrlPrefix As String, ByVal versionText As String) As String
    BuildReleaseUrl = releaseUrlPrefix & versionText
End Function

Private Sub OpenExternalPage(ByVal targetUrl As String)
    If Len(Trim$(targetUrl)) = 0 Then Exit Sub
    ActiveDocument.FollowHyperlink Address:=targetUrl, NewWindow:=True
End Sub

Private Function BuildUpdateTitle() As String
    BuildUpdateTitle = FromCodePoints(Array(26816, 26597, 26356, 26032))
End Function

Private Function BuildNewVersionTitle() As String
    BuildNewVersionTitle = FromCodePoints(Array(21457, 29616, 26032, 29256, 26412))
End Function

Private Function BuildCurrentLatestMessage() As String
    BuildCurrentLatestMessage = FromCodePoints(Array(24403, 21069, 24050, 26159, 26368, 26032, 29256, 26412, 12290))
End Function

Private Function BuildUpdateFailedMessage() As String
    BuildUpdateFailedMessage = FromCodePoints(Array(26816, 26597, 26356, 26032, 22833, 36133, 65292, 35831, 31245, 21518, 37325, 35797, 12290))
End Function

Private Function BuildUpdatePrompt(ByVal latestVersion As String) As String
    BuildUpdatePrompt = FromCodePoints(Array(21457, 29616, 26032, 29256, 26412, 65306)) & latestVersion & vbCrLf & _
                        FromCodePoints(Array(24403, 21069, 29256, 26412, 65306)) & APP_VERSION & vbCrLf & vbCrLf & _
                        FromCodePoints(Array(26159, 65306, 25171, 24320)) & " GitHub Release" & vbCrLf & _
                        FromCodePoints(Array(21542, 65306, 25171, 24320)) & " Gitee Release" & vbCrLf & _
                        FromCodePoints(Array(21462, 28040, 65306, 20851, 38381, 25552, 31034))
End Function

Private Function FromCodePoints(ByVal values As Variant) As String
    Dim i As Long
    For i = LBound(values) To UBound(values)
        FromCodePoints = FromCodePoints & ChrW$(CLng(values(i)))
    Next i
End Function


