Attribute VB_Name = "Mod_BatchConvertWordToPDF"
Option Explicit

' ==================== 常量定义 ====================
Private Const MODE_CURRENT_DOC As String = "1"   ' 当前文档
Private Const MODE_FILE_SELECT As String = "2"   ' 文件模式
Private Const MODE_FOLDER As String = "3"        ' 文件夹模式
Private Const LOG_SEPARATOR As String = "--------------------------------------------------"

' ==================== 模块级变量 ====================
Private m_fso As Object
Private processLog As String
Private successCount As Integer
Private failCount As Integer
Private totalFileCount As Integer
Private processedFileCount As Integer
Private updateType As Integer  ' 刷新方式（1=整个目录，2=仅页码）

' ==================== 主入口 ====================
Sub BatchConvertWordToPDF()
    Dim modeInput As String
    Dim folderPath As String
    Dim i As Integer
    Dim reportDoc As Document
    Dim viewReport As Integer
    Dim refreshAnswer As Integer

    ' 检查Word版本
    If Val(Application.Version) < 12 Then
        MsgBox "当前Word版本不支持PDF导出，请使用Word 2007及以上版本。", vbCritical
        Exit Sub
    End If

    ' 强制重置全局变量
    processLog = ""
    successCount = 0
    failCount = 0
    totalFileCount = 0
    processedFileCount = 0
    updateType = 2  ' 默认仅刷新页码

    processLog = "【批量转PDF处理报告】" & vbCrLf & "时间：" & Now & vbCrLf & LOG_SEPARATOR & vbCrLf

    Set m_fso = CreateObject("Scripting.FileSystemObject")

    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone
    On Error GoTo ErrorHandler

    ' 第一步：选择模式
    modeInput = InputBox("请输入模式编号：" & vbCrLf & vbCrLf & _
        "1 - 【当前文档】" & vbCrLf & _
        "2 - 【文件模式】单个或多个文件" & vbCrLf & _
        "3 - 【文件夹模式】文件夹中所有文档", _
        "Word转PDF - 模式选择", "1")

    If modeInput = "" Then GoTo Cleanup

    ' 第二步：选择刷新方式（是=刷新整个目录，否=仅刷新页码）
    refreshAnswer = MsgBox("是否刷新整个目录（标题+页码）？" & vbCrLf & vbCrLf & _
        "【是】= 刷新整个目录（标题+页码）" & vbCrLf & _
        "【否】= 仅刷新页码（保留目录格式）", _
        vbYesNo + vbDefaultButton2 + vbQuestion, "目录刷新方式")

    If refreshAnswer = vbYes Then
        updateType = 1
    Else
        updateType = 2
    End If

    ' 记录到日志
    processLog = processLog & "刷新方式：" & IIf(updateType = 1, "刷新整个目录", "仅刷新页码") & vbCrLf & LOG_SEPARATOR & vbCrLf

    Select Case modeInput
    Case MODE_CURRENT_DOC
        If Documents.count > 0 Then
            ConvertActiveDocument
        Else
            MsgBox "当前没有打开的文档！", vbExclamation
            GoTo Cleanup
        End If

    Case MODE_FILE_SELECT
        With Application.FileDialog(msoFileDialogFilePicker)
            .Title = "请选择一个或多个Word文档"
            .Filters.Clear
            .Filters.Add "Word文档", "*.doc;*.docx;*.docm"
            .AllowMultiSelect = True
            If .Show <> -1 Then GoTo Cleanup
            totalFileCount = .SelectedItems.count
            processedFileCount = 0
            For i = 1 To .SelectedItems.count
                ConvertOneFile .SelectedItems(i)
            Next i
        End With

    Case MODE_FOLDER
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "请选择包含Word文档的文件夹"
            If .Show <> -1 Then GoTo Cleanup
            folderPath = .SelectedItems(1)
        End With
        If folderPath <> "" Then
            totalFileCount = CountAllWordFiles(folderPath)
            processedFileCount = 0
            ProcessFolderWithSubfolders folderPath
        End If

    Case Else
        MsgBox "输入无效，请输入 1、2 或 3。", vbExclamation
        GoTo Cleanup
    End Select

    Application.ScreenUpdating = True

    ' 结果反馈
    If modeInput = MODE_CURRENT_DOC Then
        If successCount > 0 Or failCount > 0 Then
            MsgBox "当前文档处理完成！" & vbCrLf & _
                IIf(failCount > 0, "注意：转换失败。", "转换成功，PDF已保存在同级目录下。"), vbInformation
        End If
    Else
        viewReport = MsgBox("处理完成！" & vbCrLf & _
            "成功: " & successCount & " 个" & vbCrLf & _
            "失败: " & failCount & " 个" & vbCrLf & vbCrLf & _
            "是否生成并查看详细处理报告？", vbYesNo + vbQuestion, "批量转换完成")
        If viewReport = vbYes Then
            Set reportDoc = Documents.Add
            With reportDoc.Content
                .Text = processLog & vbCrLf & String(50, "=") & vbCrLf & _
                    "处理完成！" & vbCrLf & _
                    "成功：" & successCount & " 个" & vbCrLf & _
                    "失败：" & failCount & " 个"
                .Font.Name = "微软雅黑"
                .Font.Size = 10
            End With
        End If
    End If

Cleanup:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = wdAlertsAll
    Application.StatusBar = False
    Set m_fso = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "发生意外错误: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

' ==================== 处理当前活动文档 ====================
Sub ConvertActiveDocument()
    Dim doc As Document
    Dim pdfFileName As String

    Set doc = ActiveDocument
    On Error GoTo ActiveDocError

    If doc.Path = "" Then
        MsgBox "请先保存当前文档，以便确定PDF输出位置。", vbExclamation
        Exit Sub
    End If

    Application.StatusBar = "正在处理: " & doc.Name
    DoEvents

    Call RefreshTableOfContents(doc, updateType)

    pdfFileName = m_fso.BuildPath(doc.Path, m_fso.GetBaseName(doc.Name) & ".pdf")

    Call SafeExportAsPDF(doc, pdfFileName)

    successCount = successCount + 1
    processLog = processLog & "[成功] " & doc.Name & " (当前文档)" & vbCrLf
    Exit Sub

ActiveDocError:
    Application.ScreenUpdating = False
    failCount = failCount + 1
    processLog = processLog & "[失败] " & doc.Name & " - 原因: " & GetFriendlyErrorMessage(Err.Number, Err.Description) & vbCrLf
    MsgBox "转换失败：" & GetFriendlyErrorMessage(Err.Number, Err.Description), vbCritical
End Sub

' ==================== 递归处理文件夹 ====================
Sub ProcessFolderWithSubfolders(folderPath As String)
    Dim mainFolder As Object
    Dim subFolder As Object
    Dim file As Object

    Set mainFolder = m_fso.GetFolder(folderPath)

    For Each file In mainFolder.Files
        If IsWordDocument(file.Path) Then
            ConvertOneFile file.Path
        End If
    Next

    For Each subFolder In mainFolder.SubFolders
        ProcessFolderWithSubfolders subFolder.Path
    Next

    Set mainFolder = Nothing
End Sub

' ==================== 处理单个文件 ====================
Sub ConvertOneFile(filePath As String)
    Dim doc As Document
    Dim pdfFileName As String
    Dim fileName As String

    fileName = m_fso.GetFileName(filePath)

    On Error GoTo FileError

    processedFileCount = processedFileCount + 1
    UpdateProgress processedFileCount, totalFileCount, fileName

    Set doc = Documents.Open(fileName:=filePath, Visible:=True, ReadOnly:=True, AddToRecentFiles:=False)

    doc.ActiveWindow.Visible = True
    If doc.ActiveWindow.View.Type <> wdPrintView Then
        doc.ActiveWindow.View.Type = wdPrintView
    End If

    Call RefreshTableOfContents(doc, updateType)

    pdfFileName = m_fso.BuildPath(m_fso.GetParentFolderName(filePath), _
        m_fso.GetBaseName(filePath) & ".pdf")

    Call SafeExportAsPDF(doc, pdfFileName)

    doc.Close SaveChanges:=wdDoNotSaveChanges

    successCount = successCount + 1
    processLog = processLog & "[成功] " & fileName & vbCrLf
    GoTo Finally

FileError:
    Application.ScreenUpdating = False
    failCount = failCount + 1
    processLog = processLog & "[失败] " & fileName & " - 原因: " & GetFriendlyErrorMessage(Err.Number, Err.Description) & vbCrLf
    If Not doc Is Nothing Then
        doc.Close SaveChanges:=wdDoNotSaveChanges
    End If

Finally:
    Set doc = Nothing
    DoEvents
End Sub

' ==================== 公共函数：刷新目录 ====================
Sub RefreshTableOfContents(doc As Document, uType As Integer)
    Dim toc As TableOfContents
    Dim tof As TableOfFigures

    If uType = 1 Then
        If doc.TablesOfContents.count > 0 Then
            For Each toc In doc.TablesOfContents
                toc.Update
            Next toc
        End If
        If doc.TablesOfFigures.count > 0 Then
            For Each tof In doc.TablesOfFigures
                tof.Update
            Next tof
        End If
    ElseIf uType = 2 Then
        If doc.TablesOfContents.count > 0 Then
            For Each toc In doc.TablesOfContents
                toc.UpdatePageNumbers
            Next toc
        End If
        If doc.TablesOfFigures.count > 0 Then
            For Each tof In doc.TablesOfFigures
                tof.UpdatePageNumbers
            Next tof
        End If
    End If
End Sub

' ==================== 公共函数：安全导出PDF ====================
Sub SafeExportAsPDF(doc As Document, pdfFileName As String)
    If doc.ActiveWindow.View.Type <> wdPrintView Then
        doc.ActiveWindow.View.Type = wdPrintView
    End If

    doc.Repaginate

    Application.ScreenUpdating = True
    DoEvents

    doc.ExportAsFixedFormat _
        OutputFileName:=pdfFileName, _
        ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=False, _
        OptimizeFor:=wdExportOptimizeForPrint, _
        CreateBookmarks:=wdExportCreateHeadingBookmarks, _
        DocStructureTags:=True

    Application.ScreenUpdating = False
End Sub

' ==================== 公共函数：更新进度显示 ====================
Sub UpdateProgress(currentIndex As Integer, totalCount As Integer, fileName As String)
    If totalCount > 0 Then
        Application.StatusBar = "正在处理 (" & currentIndex & "/" & totalCount & "): " & fileName
    Else
        Application.StatusBar = "正在处理: " & fileName
    End If
    DoEvents
End Sub

' ==================== 公共函数：递归统计Word文件数 ====================
Function CountAllWordFiles(folderPath As String) As Integer
    Dim mainFolder As Object
    Dim subFolder As Object
    Dim file As Object
    Dim cnt As Integer

    cnt = 0
    Set mainFolder = m_fso.GetFolder(folderPath)

    For Each file In mainFolder.Files
        If IsWordDocument(file.Path) Then
            cnt = cnt + 1
        End If
    Next

    For Each subFolder In mainFolder.SubFolders
        cnt = cnt + CountAllWordFiles(subFolder.Path)
    Next

    CountAllWordFiles = cnt
    Set mainFolder = Nothing
End Function

' ==================== 公共函数：检查是否为Word文档 ====================
Function IsWordDocument(filePath As String) As Boolean
    Dim ext As String
    Dim fileName As String

    ext = LCase(m_fso.GetExtensionName(filePath))
    fileName = m_fso.GetFileName(filePath)

    If (ext = "doc" Or ext = "docx" Or ext = "docm") And Left(fileName, 2) <> "~$" Then
        IsWordDocument = True
    Else
        IsWordDocument = False
    End If
End Function

' ==================== 公共函数：友好的错误提示 ====================
Function GetFriendlyErrorMessage(errNumber As Long, errDesc As String) As String
    Select Case errNumber
    Case 5124
        GetFriendlyErrorMessage = "文件被其他程序占用，无法打开"
    Case 5174
        GetFriendlyErrorMessage = "文件不存在或路径无效"
    Case 5152
        GetFriendlyErrorMessage = "文件已损坏，无法打开"
    Case 6148
        GetFriendlyErrorMessage = "文档结构异常，无法正常处理"
    Case Else
        Dim msg As String
        msg = errDesc
        If InStr(msg, "being used") > 0 Or InStr(msg, "locked") > 0 Then
            GetFriendlyErrorMessage = "文件被其他程序占用，无法打开"
        ElseIf InStr(msg, "could not be opened") > 0 Then
            GetFriendlyErrorMessage = "文件不存在或无法打开"
        ElseIf InStr(msg, "password") > 0 Then
            GetFriendlyErrorMessage = "文档已加密，需要密码才能打开"
        ElseIf InStr(msg, "read-only") > 0 Then
            GetFriendlyErrorMessage = "文档为只读状态，可能被其他程序占用"
        Else
            GetFriendlyErrorMessage = msg
        End If
    End Select
End Function


