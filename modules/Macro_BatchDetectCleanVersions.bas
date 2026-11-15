Attribute VB_Name = "Macro_BatchDetectCleanVersions"
Option Explicit

Private Const DIALOG_TITLE As String = "批量检测 clean 版本"
Private Const STATUS_CLEAN As String = "Clean"
Private Const STATUS_NOT_CLEAN As String = "非 Clean"
Private Const STATUS_UNKNOWN As String = "无法检测"

' 只读入口：检测当前内容，不接受修订、不删除批注、不修改修订开关。
Public Sub BatchDetectCleanVersions()
    Dim mode As String
    Dim folderPath As String
    Dim targetFiles As Collection
    Dim result As Variant
    Dim reportDoc As Document
    Dim messageStyle As VbMsgBoxStyle

    On Error GoTo Failed
    mode = ChooseBatchMode(DIALOG_TITLE)
    If Len(mode) = 0 Then Exit Sub

    Select Case mode
        Case BATCH_MODE_CURRENT
            If Application.Documents.Count = 0 Then
                MsgBox "没有打开的文档。", vbExclamation, DIALOG_TITLE
                Exit Sub
            End If
            result = InspectCleanDocument(ActiveDocument)
            messageStyle = vbInformation
            If result(0) <> STATUS_CLEAN Then messageStyle = vbExclamation
            MsgBox "文档：" & ActiveDocument.Name & vbCrLf & _
                   "结论：" & result(0) & vbCrLf & result(1), messageStyle, DIALOG_TITLE
            Exit Sub

        Case BATCH_MODE_FILES
            Set targetFiles = PickWordFiles("选择待检测的 Word 文档")
            If targetFiles Is Nothing Then Exit Sub

        Case BATCH_MODE_FOLDER
            folderPath = PickFolder("选择待检测的文件夹")
            If Len(folderPath) = 0 Then Exit Sub
            Set targetFiles = New Collection
            CollectWordFiles folderPath, targetFiles

        Case Else
            MsgBox "无效输入，请输入 1、2 或 3。", vbExclamation, DIALOG_TITLE
            Exit Sub
    End Select

    If targetFiles.Count = 0 Then
        MsgBox "未找到 Word 文档。", vbExclamation, DIALOG_TITLE
        Exit Sub
    End If

    Set reportDoc = DetectCleanFiles(targetFiles)
    reportDoc.Activate
    Exit Sub

Failed:
    MsgBox "无法完成检测：" & Err.Description, vbExclamation, DIALOG_TITLE
End Sub

' 每份结果独立存储；报告创建在检测之后，避免把报告作为源文档处理。
Private Function DetectCleanFiles(ByVal targetFiles As Collection) As Document
    Dim results As New Collection
    Dim fileIndex As Long
    Dim previousUpdating As Boolean
    Dim previousAlerts As WdAlertLevel
    Dim settingsCaptured As Boolean
    Dim failureNumber As Long
    Dim failureDescription As String

    On Error GoTo Failed
    previousUpdating = Application.ScreenUpdating
    previousAlerts = Application.DisplayAlerts
    settingsCaptured = True
    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone

    For fileIndex = 1 To targetFiles.Count
        ShowBatchProgress fileIndex, targetFiles.Count, CStr(targetFiles(fileIndex))
        results.Add InspectCleanFile(CStr(targetFiles(fileIndex)))
    Next fileIndex

    Set DetectCleanFiles = CreateCleanDetectionReport(results)
    GoTo Finish

Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    Resume Finish

Finish:
    On Error Resume Next
    If settingsCaptured Then
        Application.ScreenUpdating = previousUpdating
        Application.DisplayAlerts = previousAlerts
        Application.StatusBar = False
    End If
    On Error GoTo 0
    If failureNumber <> 0 Then Err.Raise failureNumber, DIALOG_TITLE, failureDescription
End Function

' 返回 [路径, 结论, 原因]。仅关闭本次打开的文档，始终不保存。
Private Function InspectCleanFile(ByVal filePath As String) As Variant
    Dim doc As Document
    Dim wasAlreadyOpen As Boolean
    Dim openError As String
    Dim closeError As String
    Dim result As Variant
    Dim inspectionApp As Word.Application
    Dim previousSecurity As MsoAutomationSecurity
    Dim securityCaptured As Boolean
    Dim failureDescription As String

    result = Array(STATUS_UNKNOWN, "Word 未能打开文件。")
    On Error GoTo Failed
    ' 先复用当前实例的打开对象，保留用户未保存内容。
    On Error Resume Next
    Set doc = Application.Documents(filePath)
    On Error GoTo Failed
    wasAlreadyOpen = Not doc Is Nothing
    If Not wasAlreadyOpen And (LCase$(Right$(filePath, 4)) = ".doc" Or _
                               LCase$(Right$(filePath, 5)) = ".docm") Then
        ' 在运行 VBA 的实例中禁宏打开实际 VBA 文件会中断调用链。
        ' 独立实例只负责读取，避免源文件工程影响当前插件的执行。
        Set inspectionApp = New Word.Application
        inspectionApp.Visible = False
        inspectionApp.DisplayAlerts = wdAlertsNone
        inspectionApp.AutomationSecurity = msoAutomationSecurityForceDisable
        Set doc = inspectionApp.Documents.Open(FileName:=filePath, ReadOnly:=True, _
                                               Visible:=False, AddToRecentFiles:=False)
    ElseIf Not wasAlreadyOpen Then
        previousSecurity = Application.AutomationSecurity
        securityCaptured = True
        Application.AutomationSecurity = msoAutomationSecurityForceDisable
        Set doc = SafeOpenDocument(filePath, True, wasAlreadyOpen, openError)
    End If

    If doc Is Nothing Or Len(openError) > 0 Then
        If Len(openError) = 0 Then openError = "Word 未能打开文件。"
        result = Array(STATUS_UNKNOWN, "无法打开文档：" & openError)
    Else
        result = InspectCleanDocument(doc)
    End If
    GoTo Finish

Failed:
    failureDescription = Err.Description
    result = Array(STATUS_UNKNOWN, "检测失败：" & failureDescription)
    Resume Finish

Finish:
    On Error Resume Next
    If Not doc Is Nothing Then
        If Not wasAlreadyOpen Then
            If Not inspectionApp Is Nothing Then
                CloseCleanInspectionDocument doc, inspectionApp, closeError
            ElseIf Not CloseDocumentQuietly(doc, False, closeError) Then
                If Len(closeError) = 0 Then closeError = Err.Description
                If Len(closeError) = 0 Then closeError = "文档未能关闭，仍保持打开。"
            End If
            If Len(closeError) > 0 Then
                result(0) = STATUS_UNKNOWN
                result(1) = result(1) & "；关闭失败：" & closeError
            End If
        End If
    End If
    If Not inspectionApp Is Nothing Then
        Err.Clear
        inspectionApp.Quit SaveChanges:=wdDoNotSaveChanges
        If Err.Number <> 0 Then
            result(0) = STATUS_UNKNOWN
            result(1) = result(1) & "；无法退出检测实例：" & Err.Description
        End If
        Set inspectionApp = Nothing
    End If
    Err.Clear
    If securityCaptured Then
        Application.AutomationSecurity = previousSecurity
        If Err.Number <> 0 Then
            result(0) = STATUS_UNKNOWN
            result(1) = result(1) & "；无法恢复宏打开设置：" & Err.Description
        End If
    End If
    On Error GoTo 0
    InspectCleanFile = Array(filePath, result(0), result(1))
End Function

' 专用实例中的关闭校验必须检查该实例，不能枚举插件所在实例。
Private Sub CloseCleanInspectionDocument(ByVal doc As Document, _
                                         ByVal inspectionApp As Word.Application, _
                                         ByRef closeError As String)
    Dim openDoc As Document
    On Error GoTo Failed
    doc.Close SaveChanges:=wdDoNotSaveChanges
    For Each openDoc In inspectionApp.Documents
        If openDoc Is doc Then
            closeError = "关闭被取消，文档保持打开。"
            Exit Sub
        End If
    Next openDoc
    Exit Sub
Failed:
    closeError = Err.Description
End Sub

' 返回 [结论, 原因]；任何检查异常均不能返回 Clean。
Private Function InspectCleanDocument(ByVal doc As Document) As Variant
    Dim hasRevisions As Boolean
    Dim hasComments As Boolean
    Dim trackingEnabled As Boolean
    Dim rootStory As Range
    Dim storyRange As Range
    Dim reasons As String
    Dim failureDescription As String
    Dim stageName As String

    On Error GoTo Failed
    stageName = "修订跟踪读取"
    trackingEnabled = doc.TrackRevisions
    stageName = "文档批注读取"
    hasComments = (doc.Comments.Count > 0)

    stageName = "故事区域枚举"
    For Each rootStory In doc.StoryRanges
        Set storyRange = rootStory
        Do While Not storyRange Is Nothing
            stageName = "区域修订读取"
            If storyRange.Revisions.Count > 0 Then hasRevisions = True
            stageName = "下一故事区域读取"
            Set storyRange = storyRange.NextStoryRange
        Loop
    Next rootStory

    reasons = GetCleanFailureReasons(hasRevisions, hasComments, trackingEnabled)
    If Len(reasons) = 0 Then
        InspectCleanDocument = Array(STATUS_CLEAN, "无修订；无批注；修订跟踪已关闭。")
    Else
        InspectCleanDocument = Array(STATUS_NOT_CLEAN, reasons)
    End If
    Exit Function

Failed:
    failureDescription = stageName & "：" & Err.Description
    reasons = GetCleanFailureReasons(hasRevisions, hasComments, trackingEnabled)
    AppendCleanReason reasons, "检查失败：" & failureDescription
    InspectCleanDocument = Array(STATUS_UNKNOWN, reasons)
End Function

Private Function GetCleanFailureReasons(ByVal hasRevisions As Boolean, _
                                       ByVal hasComments As Boolean, _
                                       ByVal trackingEnabled As Boolean) As String
    Dim reasons As String
    If hasRevisions Then AppendCleanReason reasons, "存在修订"
    If hasComments Then AppendCleanReason reasons, "存在批注"
    If trackingEnabled Then AppendCleanReason reasons, "修订跟踪仍开启"
    GetCleanFailureReasons = reasons
End Function

Private Sub AppendCleanReason(ByRef reasons As String, ByVal reason As String)
    If Len(reasons) > 0 Then reasons = reasons & "；"
    reasons = reasons & reason
End Sub

Private Function CreateCleanDetectionReport(ByVal results As Collection) As Document
    Dim reportDoc As Document
    Dim tbl As Table
    Dim entry As Variant
    Dim rowIndex As Long
    Dim cleanCount As Long
    Dim notCleanCount As Long
    Dim unknownCount As Long

    For Each entry In results
        Select Case entry(1)
            Case STATUS_CLEAN: cleanCount = cleanCount + 1
            Case STATUS_NOT_CLEAN: notCleanCount = notCleanCount + 1
            Case Else: unknownCount = unknownCount + 1
        End Select
    Next entry

    Set reportDoc = Application.Documents.Add
    reportDoc.TrackRevisions = False
    reportDoc.Content.Text = "Clean 版本批量检测报告" & vbCr & _
        "判定条件：无修订、无批注且修订跟踪关闭。高亮不纳入判定。" & vbCr & _
        "共 " & results.Count & " 个文件；Clean " & cleanCount & _
        "；非 Clean " & notCleanCount & "；无法检测 " & unknownCount & "。" & vbCr & _
        "已打开的文档检测当前内容；其余文件只读打开。源文档不修改、不保存。" & vbCr

    Set tbl = reportDoc.Tables.Add(reportDoc.Range(reportDoc.Content.End - 1, _
        reportDoc.Content.End - 1), results.Count + 1, 3)
    tbl.Cell(1, 1).Range.Text = "文件路径"
    tbl.Cell(1, 2).Range.Text = "结论"
    tbl.Cell(1, 3).Range.Text = "原因"
    rowIndex = 1
    For Each entry In results
        rowIndex = rowIndex + 1
        tbl.Cell(rowIndex, 1).Range.Text = CStr(entry(0))
        tbl.Cell(rowIndex, 2).Range.Text = CStr(entry(1))
        tbl.Cell(rowIndex, 3).Range.Text = CStr(entry(2))
    Next entry
    tbl.AutoFitBehavior wdAutoFitWindow
    Set CreateCleanDetectionReport = reportDoc
End Function
