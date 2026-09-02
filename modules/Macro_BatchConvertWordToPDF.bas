Attribute VB_Name = "Macro_BatchConvertWordToPDF"
Option Explicit

' ==========================================
' Word 批量转 PDF（含锁定域所见即所得、目录刷新与转换报告）
' 模式选择/文件选择/遍历改用 Mod_Core_* 公共层；
' 原 7 个模块级状态收敛为过程内会话上下文（ConvertSession）传参。
' 注意：本宏刻意以可见窗口打开文档导出（修复页眉页脚边框线丢失），
' 不使用公共层的后台打开函数。
' ==========================================

Private Const DIALOG_TITLE As String = "Word批量转PDF"
Private Const LOG_SEPARATOR As String = "--------------------------------------------------"

' 转换选项序号常量
Private Const CONVERT_OPT_WYSIWYG As String = "1"     ' 所见即所得（锁定全部域导出，推荐）
Private Const CONVERT_OPT_PAGE_ONLY As String = "2"   ' 仅刷新页码（保留目录格式）
Private Const CONVERT_OPT_FULL_TOC As String = "3"    ' 刷新整个目录（更新标题+页码）

' 会话上下文：一次批量转换的全部状态
Private Type ConvertSession
    fso As Object
    processLog As String
    successCount As Long
    failCount As Long
    totalFileCount As Long
    processedFileCount As Long
    updateType As Long   ' 转换选项（1=所见即所得/锁定全部域，2=仅刷新页码，3=刷新整个目录）
End Type

' ==================== 主入口 ====================
Sub BatchConvertWordToPDF()
    Dim session As ConvertSession
    Dim modeInput As String
    Dim optionInput As String
    Dim folderPath As String
    Dim fileList As Collection
    Dim i As Long
    Dim reportDoc As Document
    Dim viewReport As VbMsgBoxResult
    Dim optionDesc As String

    ' 检查Word版本
    If Val(Application.Version) < 12 Then
        MsgBox "当前Word版本不支持PDF导出，请使用Word 2007及以上版本。", vbCritical, DIALOG_TITLE
        Exit Sub
    End If

    session.updateType = 1  ' 默认所见即所得（锁定全部域）
    session.processLog = "【批量转PDF处理报告】" & vbCrLf & "时间：" & Now & vbCrLf & LOG_SEPARATOR & vbCrLf
    Set session.fso = CreateObject("Scripting.FileSystemObject")

    ' 第一步：选择模式
    modeInput = ChooseBatchMode(DIALOG_TITLE)
    If modeInput = "" Then Exit Sub

    ' 第二步：选择转换选项（1=所见即所得，2=仅刷新页码，3=刷新整个目录）
    optionInput = ChooseConvertOption(DIALOG_TITLE)
    If optionInput = "" Then Exit Sub

    Select Case optionInput
    Case CONVERT_OPT_WYSIWYG
        session.updateType = 1
        optionDesc = "所见即所得（锁定全部域）"
    Case CONVERT_OPT_PAGE_ONLY
        session.updateType = 2
        optionDesc = "仅刷新页码"
    Case CONVERT_OPT_FULL_TOC
        session.updateType = 3
        optionDesc = "刷新整个目录"
    Case Else
        MsgBox "输入无效，请输入 1、2 或 3。", vbExclamation, DIALOG_TITLE
        Exit Sub
    End Select

    session.processLog = session.processLog & "转换选项：" & optionDesc & vbCrLf & LOG_SEPARATOR & vbCrLf

    ' 第三步：按模式收集文件
    Select Case modeInput
    Case BATCH_MODE_CURRENT
        If Documents.Count = 0 Then
            MsgBox "当前没有打开的文档！", vbExclamation, DIALOG_TITLE
            Exit Sub
        End If

    Case BATCH_MODE_FILES
        Set fileList = PickWordFiles("请选择一个或多个Word文档")
        If fileList Is Nothing Then Exit Sub

    Case BATCH_MODE_FOLDER
        folderPath = PickFolder("请选择包含Word文档的文件夹")
        If folderPath = "" Then Exit Sub

        Set fileList = New Collection
        Application.StatusBar = "正在扫描文件..."
        CollectWordFiles folderPath, fileList

        If fileList.Count = 0 Then
            Application.StatusBar = False
            MsgBox "所选文件夹（含子文件夹）中未找到 Word 文档。", vbExclamation, DIALOG_TITLE
            Exit Sub
        End If

    Case Else
        MsgBox "输入无效，请输入 1、2 或 3。", vbExclamation, DIALOG_TITLE
        Exit Sub
    End Select

    BeginBatchUI suppressAlerts:=True
    On Error GoTo ErrorHandler

    ' 第四步：执行转换
    If modeInput = BATCH_MODE_CURRENT Then
        ConvertActiveDocument session
    Else
        session.totalFileCount = fileList.Count
        For i = 1 To fileList.Count
            ConvertOneFile session, CStr(fileList(i))
        Next i
    End If

    Application.ScreenUpdating = True

    ' 完成提示
    If modeInput = BATCH_MODE_CURRENT Then
        If session.successCount > 0 Or session.failCount > 0 Then
            MsgBox "当前文档处理完成：" & vbCrLf & _
                IIf(session.failCount > 0, "注意：转换失败。", "转换成功！PDF已保存在同级目录下。"), _
                vbInformation, DIALOG_TITLE
        End If
    Else
        viewReport = MsgBox("处理完成！" & vbCrLf & _
            "成功: " & session.successCount & " 个" & vbCrLf & _
            "失败: " & session.failCount & " 个" & vbCrLf & vbCrLf & _
            "是否生成并查看详细处理报告？", vbYesNo + vbQuestion, DIALOG_TITLE)
        If viewReport = vbYes Then
            Set reportDoc = Documents.Add
            With reportDoc.Content
                .Text = session.processLog & vbCrLf & String(50, "=") & vbCrLf & _
                    "处理完成！" & vbCrLf & _
                    "成功：" & session.successCount & " 个" & vbCrLf & _
                    "失败：" & session.failCount & " 个"
                .Font.Name = "微软雅黑"
                .Font.Size = 10
            End With
        End If
    End If

Cleanup:
    EndBatchUI
    Set session.fso = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "处理过程出错: " & Err.Description, vbCritical, DIALOG_TITLE
    Resume Cleanup
End Sub

' ==================== 转换选项选择对话框 ====================
Private Function ChooseConvertOption(ByVal dialogTitle As String) As String
    Dim answer As String

    answer = InputBox("请输入 PDF 转换选项序号：" & vbCrLf & vbCrLf & _
                      "1 - 所见即所得（推荐）：锁定全部域导出，绝对不刷新任何目录/编号/交叉引用，导出后自动恢复" & vbCrLf & _
                      "2 - 仅刷新页码：仅更新目录页码，保留目录排版与文本" & vbCrLf & _
                      "3 - 刷新整个目录：更新目录全部标题与页码", _
                      dialogTitle, CONVERT_OPT_WYSIWYG)

    If StrPtr(answer) = 0 Then
        ChooseConvertOption = ""
    Else
        ChooseConvertOption = Trim$(answer)
    End If
End Function

' ==================== 处理当前活动文档 ====================
Private Sub ConvertActiveDocument(ByRef session As ConvertSession)
    Dim doc As Document
    Dim pdfFileName As String
    Dim wasLocked As Boolean

    Set doc = ActiveDocument
    wasLocked = False
    On Error GoTo ActiveDocError

    If doc.Path = "" Then
        MsgBox "请先保存当前文档，以便确定PDF输出位置。", vbExclamation, DIALOG_TITLE
        Exit Sub
    End If

    Application.StatusBar = "正在处理: " & doc.Name
    DoEvents

    ' 根据选项执行转换前处理
    If session.updateType = 1 Then
        ' 所见即所得：锁定全部域
        SetAllFieldsLocked doc, True
        wasLocked = True
    ElseIf session.updateType = 2 Then
        ' 仅刷新页码
        RefreshTableOfContents doc, 2
    ElseIf session.updateType = 3 Then
        ' 刷新整个目录
        RefreshTableOfContents doc, 1
    End If

    pdfFileName = session.fso.BuildPath(doc.Path, session.fso.GetBaseName(doc.Name) & ".pdf")

    SafeExportAsPDF doc, pdfFileName

    ' 所见即所得：导出完成后恢复解除域锁定
    If wasLocked Then
        SetAllFieldsLocked doc, False
        wasLocked = False
    End If

    session.successCount = session.successCount + 1
    session.processLog = session.processLog & "[成功] " & doc.Name & " (当前文档)" & vbCrLf
    Exit Sub

ActiveDocError:
    If wasLocked Then
        SetAllFieldsLocked doc, False
        wasLocked = False
    End If
    Application.ScreenUpdating = False
    session.failCount = session.failCount + 1
    session.processLog = session.processLog & "[失败] " & doc.Name & " - 原因: " & _
        GetFriendlyErrorMessage(Err.Number, Err.Description) & vbCrLf
    MsgBox "转换失败：" & GetFriendlyErrorMessage(Err.Number, Err.Description), vbCritical, DIALOG_TITLE
End Sub

' ==================== 处理单个文件 ====================
Private Sub ConvertOneFile(ByRef session As ConvertSession, ByVal filePath As String)
    Dim doc As Document
    Dim pdfFileName As String
    Dim fileName As String

    fileName = session.fso.GetFileName(filePath)

    On Error GoTo FileError

    session.processedFileCount = session.processedFileCount + 1
    ShowBatchProgress session.processedFileCount, session.totalFileCount, fileName
    DoEvents

    ' 刻意以可见窗口打开：不可见窗口导出会丢失页眉/页脚中的边框线
    Set doc = Documents.Open(FileName:=filePath, Visible:=True, ReadOnly:=True, AddToRecentFiles:=False)

    doc.ActiveWindow.Visible = True
    If doc.ActiveWindow.View.Type <> wdPrintView Then
        doc.ActiveWindow.View.Type = wdPrintView
    End If

    ' 根据选项执行转换前处理
    If session.updateType = 1 Then
        ' 所见即所得：锁定全部域
        SetAllFieldsLocked doc, True
    ElseIf session.updateType = 2 Then
        ' 仅刷新页码
        RefreshTableOfContents doc, 2
    ElseIf session.updateType = 3 Then
        ' 刷新整个目录
        RefreshTableOfContents doc, 1
    End If

    pdfFileName = session.fso.BuildPath(session.fso.GetParentFolderName(filePath), _
        session.fso.GetBaseName(filePath) & ".pdf")

    SafeExportAsPDF doc, pdfFileName

    doc.Close SaveChanges:=wdDoNotSaveChanges

    session.successCount = session.successCount + 1
    session.processLog = session.processLog & "[成功] " & fileName & vbCrLf
    GoTo Finally

FileError:
    Application.ScreenUpdating = False
    session.failCount = session.failCount + 1
    session.processLog = session.processLog & "[失败] " & fileName & " - 原因: " & _
        GetFriendlyErrorMessage(Err.Number, Err.Description) & vbCrLf
    If Not doc Is Nothing Then
        doc.Close SaveChanges:=wdDoNotSaveChanges
    End If

Finally:
    Set doc = Nothing
    DoEvents
End Sub

' ==================== 辅助函数：批量锁定/解锁全部域 ====================
Public Sub SetAllFieldsLocked(doc As Document, ByVal lockState As Boolean)
    Dim storyRange As Range
    Dim fld As Field
    Dim shp As Shape

    On Error Resume Next
    ' 1. 遍历全部故事区（正文、页眉、页脚、脚注、尾注、文本框等）
    For Each storyRange In doc.StoryRanges
        Do
            storyRange.Fields.Locked = lockState
            For Each fld In storyRange.Fields
                fld.Locked = lockState
            Next fld
            Set storyRange = storyRange.NextStoryRange
        Loop While Not (storyRange Is Nothing)
    Next storyRange

    ' 2. 补充遍历主文档形状内的文本框域（防遗漏）
    For Each shp In doc.Shapes
        If shp.TextFrame.HasText <> 0 Then
            shp.TextFrame.TextRange.Fields.Locked = lockState
            For Each fld In shp.TextFrame.TextRange.Fields
                fld.Locked = lockState
            Next fld
        End If
    Next shp
    On Error GoTo 0
End Sub

' ==================== 辅助函数：刷新目录 ====================
Private Sub RefreshTableOfContents(doc As Document, ByVal uType As Long)
    Dim toc As TableOfContents
    Dim tof As TableOfFigures

    If uType = 1 Then
        ' 刷新整个目录（标题+页码）
        If doc.TablesOfContents.Count > 0 Then
            For Each toc In doc.TablesOfContents
                toc.Update
            Next toc
        End If
        If doc.TablesOfFigures.Count > 0 Then
            For Each tof In doc.TablesOfFigures
                tof.Update
            Next tof
        End If
    ElseIf uType = 2 Then
        ' 仅刷新页码
        If doc.TablesOfContents.Count > 0 Then
            For Each toc In doc.TablesOfContents
                toc.UpdatePageNumbers
            Next toc
        End If
        If doc.TablesOfFigures.Count > 0 Then
            For Each tof In doc.TablesOfFigures
                tof.UpdatePageNumbers
            Next tof
        End If
    End If
End Sub

' ==================== 辅助函数：安全导出PDF ====================
Private Sub SafeExportAsPDF(doc As Document, ByVal pdfFileName As String)
    If doc.ActiveWindow.View.Type <> wdPrintView Then
        doc.ActiveWindow.View.Type = wdPrintView
    End If

    doc.Repaginate

    ' 导出期间必须开启屏幕刷新，否则部分版式元素（页眉页脚边框）会丢失
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

' ==================== 辅助函数：友好的错误提示 ====================
Private Function GetFriendlyErrorMessage(ByVal errNumber As Long, ByVal errDesc As String) As String
    Select Case errNumber
    Case 5124
        GetFriendlyErrorMessage = "文件被锁定或被占用，无法打开"
    Case 5174
        GetFriendlyErrorMessage = "文件不存在或路径无效"
    Case 5152
        GetFriendlyErrorMessage = "文件名损坏，无法打开"
    Case 6148
        GetFriendlyErrorMessage = "文档结构异常，无法完成操作"
    Case Else
        Dim msg As String
        msg = errDesc
        If InStr(msg, "being used") > 0 Or InStr(msg, "locked") > 0 Then
            GetFriendlyErrorMessage = "文件被锁定或被占用，无法打开"
        ElseIf InStr(msg, "could not be opened") > 0 Then
            GetFriendlyErrorMessage = "文件不存在或无法打开"
        ElseIf InStr(msg, "password") > 0 Then
            GetFriendlyErrorMessage = "文档已加密，需要密码才能打开"
        ElseIf InStr(msg, "read-only") > 0 Then
            GetFriendlyErrorMessage = "文档为只读状态，可能被保护或被占用"
        Else
            GetFriendlyErrorMessage = msg
        End If
    End Select
End Function
