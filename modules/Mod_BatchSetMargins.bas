Attribute VB_Name = "Mod_BatchSetMargins"
Option Explicit

' ==========================================
' 一键设置页边距（上下左右 2.54 厘米）
' 交互与遍历统一走 Mod_Core_* 公共层
' ==========================================

Private Const DIALOG_TITLE As String = "一键设置页边距"

Sub BatchSetMargins()
    Dim choice As String
    Dim fileList As Collection
    Dim folderPath As String
    Dim processedCount As Long
    Dim failedCount As Long
    Dim startTime As Single

    choice = ChooseBatchMode(DIALOG_TITLE)
    If choice = "" Then Exit Sub

    startTime = Timer

    Select Case choice
        Case BATCH_MODE_CURRENT
            ' 模式 1 保持屏幕刷新，让用户能直接看到边距变化
            If Documents.count > 0 Then
                ProcessSingleDoc ActiveDocument, True
            Else
                MsgBox "没有打开的文档！", vbExclamation, DIALOG_TITLE
            End If
            Exit Sub

        Case BATCH_MODE_FILES
            Set fileList = PickWordFiles("请选择要处理的Word文件（按住Ctrl或Shift可多选）")
            If fileList Is Nothing Then Exit Sub

        Case BATCH_MODE_FOLDER
            folderPath = PickFolder("请选择包含Word文件的文件夹")
            If folderPath = "" Then Exit Sub

            Set fileList = New Collection
            Application.StatusBar = "正在扫描文件..."
            CollectWordFiles folderPath, fileList

            If fileList.count = 0 Then
                Application.StatusBar = False
                MsgBox "所选文件夹（含子文件夹）中未找到 Word 文档。", vbExclamation, DIALOG_TITLE
                Exit Sub
            End If

        Case Else
            MsgBox "无效输入，请输入 1、2 或 3。", vbExclamation, DIALOG_TITLE
            Exit Sub
    End Select

    BeginBatchUI
    On Error GoTo ErrH

    Dim i As Long
    For i = 1 To fileList.count
        ShowBatchProgress i, fileList.count, CStr(fileList(i))
        If ProcessFile(CStr(fileList(i))) Then
            processedCount = processedCount + 1
        Else
            failedCount = failedCount + 1
        End If
    Next i

    EndBatchUI

    If failedCount = 0 Then
        MsgBox "处理完成：成功 " & processedCount & " 个文件，耗时 " & _
               Format(Timer - startTime, "0.00") & " 秒。", vbInformation, DIALOG_TITLE
    Else
        MsgBox "处理完成：成功 " & processedCount & " 个，失败 " & failedCount & _
               " 个（无法打开或被保护，详见后述提示）。", vbExclamation, DIALOG_TITLE
    End If
    Exit Sub

ErrH:
    EndBatchUI
    MsgBox "处理出错：" & Err.Description, vbCritical, DIALOG_TITLE
End Sub

' 处理单个文件路径；返回是否成功
Private Function ProcessFile(ByVal filePath As String) As Boolean
    Dim doc As Document
    Dim wasAlreadyOpen As Boolean

    Set doc = SafeOpenDocument(filePath, False, wasAlreadyOpen)
    If doc Is Nothing Then
        ProcessFile = False
        Exit Function
    End If

    If doc.ProtectionType <> wdNoProtection Then
        If Not wasAlreadyOpen Then CloseDocumentQuietly doc, False
        ProcessFile = False
        Exit Function
    End If

    ProcessFile = SetMarginsSafely(doc)

    ' 若文档是本过程打开的，保存并关闭；用户已打开的保持原状
    If Not wasAlreadyOpen Then CloseDocumentQuietly doc, True
End Function

' 处理当前文档对象（检查保护 + 结果提示）
Private Sub ProcessSingleDoc(doc As Document, showMsg As Boolean)
    If doc.ProtectionType <> wdNoProtection Then
        If showMsg Then MsgBox "当前文档受保护（只读），无法修改边距。", vbCritical, DIALOG_TITLE
        Exit Sub
    End If

    Dim success As Boolean
    success = SetMarginsSafely(doc)

    If showMsg Then
        If success Then
            doc.Repaginate ' 强制重排以显示效果
            MsgBox "当前文档处理完成，所有节的边距已调整。", vbInformation, DIALOG_TITLE
        Else
            MsgBox "处理过程中遇到问题，请检查文档页面设置。", vbExclamation, DIALOG_TITLE
        End If
    End If
End Sub

' ==========================================
' 核心逻辑：安全设置边距（防 4608 错误码 + 遍历 Sections）
' ==========================================
Private Function SetMarginsSafely(doc As Document) As Boolean
    Dim targetPoints As Single
    Dim sec As Section
    Dim hasError As Boolean

    ' 2.54 厘米 = 72 磅
    targetPoints = 72
    hasError = False

    On Error Resume Next

    ' Word 文档可能包含多个节，直接设置 doc.PageSetup 有时只对第一节生效
    For Each sec In doc.Sections
        Err.Clear

        ' --- 尝试 1: 标准转换 ---
        With sec.PageSetup
            .TopMargin = Application.CentimetersToPoints(2.54)
            .BottomMargin = Application.CentimetersToPoints(2.54)
            .LeftMargin = Application.CentimetersToPoints(2.54)
            .RightMargin = Application.CentimetersToPoints(2.54)
        End With

        If Err.Number <> 0 Then
            Err.Clear

            ' --- 尝试 2: 直接按磅赋值 (绕过转换函数问题) ---
            With sec.PageSetup
                .TopMargin = targetPoints
                .BottomMargin = targetPoints
                .LeftMargin = targetPoints
                .RightMargin = targetPoints
            End With

            If Err.Number <> 0 Then
                Debug.Print "节处理失败: " & doc.Name & " Section:" & sec.Index
                hasError = True
            End If
        End If
    Next sec

    On Error GoTo 0

    SetMarginsSafely = Not hasError
End Function