Attribute VB_Name = "Macro_AcceptFormatRevisions"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function GetLongPathNameW Lib "kernel32" ( _
        ByVal lpszShortPath As LongPtr, _
        ByVal lpszLongPath As LongPtr, _
        ByVal cchBuffer As Long) As Long
#Else
    Private Declare Function GetLongPathNameW Lib "kernel32" ( _
        ByVal lpszShortPath As Long, _
        ByVal lpszLongPath As Long, _
        ByVal cchBuffer As Long) As Long
#End If
' =============================================
' 模块名：Macro_AcceptFormatRevisions
' 功能：只接受"格式类修订"（字符格式、段落格式），
'       插入、删除、替换、移动等内容类修订保持未接受状态。
' 实现：
'   .docx / .docm 走 OOXML 快速路径：保存并备份原文件后关闭文档，
'   用 Shell.Application 解压、MSXML 删除 word/*.xml 中的
'   w:rPrChange 与 w:pPrChange 元素，再重新压缩并打开文档。
'   全程为进程内 COM 调用，不启动 PowerShell 等外部进程，
'   可兼容"阻止 Office 创建子进程"等安全策略。
'   其他格式（.doc 等）回退到 Word 对象模型逐个接受。
' =============================================

Private Const DIALOG_TITLE As String = "接受格式修订"
Private Const W_NAMESPACE As String = "xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'"

' Ribbon 回调：接受全文格式类修订
Public Sub AcceptFormatRevisions(control As IRibbonControl)
    Dim doc As Document
    Dim docPath As String

    If Application.Documents.Count = 0 Then
        MsgBox "当前没有打开的文档。", vbExclamation, DIALOG_TITLE
        Exit Sub
    End If

    Set doc = ActiveDocument
    docPath = doc.FullName

    ' OneDrive/SharePoint 云端文档的 FullName 可能返回 https 地址：
    ' 先换算为本机同步副本路径；无法换算时回退到 Word 对象模型
    If IsWebPath(docPath) Then
        Dim syncLocalPath As String
        syncLocalPath = ResolveWebPathToLocalFolder(docPath)
        If Len(syncLocalPath) = 0 Then
            AcceptFormatRevisionsByWordApi doc
            Exit Sub
        End If
        docPath = syncLocalPath
    End If

    ' 未保存的新文档没有磁盘路径，无法使用 OOXML 快速路径
    If Len(docPath) = 0 Then
        MsgBox "请先将文档保存为 .docx 或 .docm 格式后再使用。", vbExclamation, DIALOG_TITLE
        Exit Sub
    End If

    If IsOoxmlDocument(docPath) Then
        AcceptFormatRevisionsByOoxml docPath
    Else
        AcceptFormatRevisionsByWordApi doc
    End If
End Sub

' 宏列表入口：Application.Run 不支持带参数的过程，包装为无参入口
Public Sub Wrapper_AcceptFormatRevisions()
    AcceptFormatRevisions Nothing
End Sub

' 是否为 OOXML 文档（.docx / .docm）
Private Function IsOoxmlDocument(ByVal docPath As String) As Boolean
    Dim ext As String

    If Len(docPath) < 5 Then Exit Function
    ext = LCase$(Right$(docPath, 5))
    IsOoxmlDocument = (ext = ".docx" Or ext = ".docm")
End Function

' =============================================
' 快速路径：直接处理 OOXML（进程内 COM，约 2-6 秒）
' =============================================
Private Sub AcceptFormatRevisionsByOoxml(ByVal docPath As String)
    Dim doc As Document
    Dim fso As Object
    Dim backupPath As String
    Dim workDir As String
    Dim zipCopyPath As String
    Dim outZipPath As String
    Dim removed As Long
    Dim ok As Boolean
    Dim remaining As Long

    Set doc = ActiveDocument

    ' 1. 自动保存（OOXML 需要文件已落盘）
    If Not doc.Saved Then
        On Error Resume Next
        doc.Save
        If Err.Number <> 0 Then
            MsgBox "保存文档失败：" & Err.Description, vbExclamation, DIALOG_TITLE
            Err.Clear
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    backupPath = docPath & ".bak"

    ' 2. 备份原文件
    On Error Resume Next
    If fso.FileExists(backupPath) Then fso.DeleteFile backupPath, True
    fso.CopyFile docPath, backupPath, True
    If Err.Number <> 0 Then
        MsgBox "无法创建备份文件：" & Err.Description, vbExclamation, DIALOG_TITLE
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    ' 3. 创建临时工作目录与 zip 副本
    Randomize
    workDir = fso.GetSpecialFolder(2) & "\RATools_Ooxml_" & Format$(Now, "yyyymmddhhnnss") & "_" & CStr(Int(Rnd() * 100000))
    On Error Resume Next
    fso.CreateFolder workDir
    If Err.Number <> 0 Then
        MsgBox "无法创建临时目录：" & Err.Description, vbExclamation, DIALOG_TITLE
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    zipCopyPath = workDir & "\_src.zip"
    outZipPath = workDir & "\_out.zip"
    fso.CopyFile docPath, zipCopyPath, True
    Err.Clear
    On Error GoTo 0

    ' 4. 关闭文档（已保存，无提示）
    doc.Close wdSaveChanges
    Set doc = Nothing

    ' 5. 解压 -> 删除格式修订标记 -> 重新压缩
    ok = ProcessOoxml(zipCopyPath, workDir, outZipPath, removed)

    ' 6. 成功则替换原文件，失败则从备份恢复
    If ok Then
        On Error Resume Next
        fso.CopyFile outZipPath, docPath, True
        If Err.Number <> 0 Then
            ok = False
            Err.Clear
        End If
        On Error GoTo 0
    End If
    If Not ok Then
        On Error Resume Next
        fso.CopyFile backupPath, docPath, True
        On Error GoTo 0
    End If

    ' 7. 重新打开文档
    On Error Resume Next
    Set doc = Application.Documents.Open(docPath)
    If Err.Number <> 0 Then
        MsgBox "文档已处理，但重新打开失败，请手动打开：" & vbCrLf & docPath & vbCrLf & Err.Description, vbExclamation, DIALOG_TITLE
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    ' 剩余修订数量（Count 很快，不做逐条枚举）
    remaining = doc.Revisions.Count

    ' 8. 清理临时目录
    On Error Resume Next
    fso.DeleteFolder workDir, True
    On Error GoTo 0

    ' 9. 结果提示
    If ok Then
        MsgBox "已完成：格式类修订已接受（处理 " & removed & " 处格式标记）。" & vbCrLf & vbCrLf & _
               "文档中剩余 " & remaining & " 处修订（插入、删除等内容类），均已保持未接受状态。" & vbCrLf & vbCrLf & _
               "原文件备份：" & backupPath, vbInformation, DIALOG_TITLE
    Else
        MsgBox "处理失败，已从备份恢复原文件并重新打开。", vbExclamation, DIALOG_TITLE
    End If
End Sub

' 解压 -> 删除格式修订标记 -> 重新压缩
' 8.3 short path -> long path (Shell.Application.NameSpace does not accept short paths)
Private Function ToLongPath(ByVal path As String) As String
    Dim buf As String
    Dim n As Long

    On Error Resume Next
    If Len(path) = 0 Then Exit Function
    buf = String$(1024, vbNullChar)
    n = GetLongPathNameW(StrPtr(path), StrPtr(buf), 1024)
    If n > 0 And n <= 1024 Then
        ToLongPath = Left$(buf, n)
    Else
        ToLongPath = path
    End If
End Function
Private Function ProcessOoxml(ByVal zipCopyPath As String, ByVal workDir As String, ByVal outZipPath As String, ByRef removed As Long) As Boolean
    Dim fso As Object
    Dim shell As Object
    Dim zipNs As Object
    Dim destNs As Object
    Dim srcNs As Object
    Dim outNs As Object
    Dim payloadDir As String

    On Error GoTo Failed
    Set fso = CreateObject("Scripting.FileSystemObject")
    payloadDir = workDir & "\payload"
    fso.CreateFolder payloadDir

    Set shell = CreateObject("Shell.Application")
    If shell Is Nothing Then GoTo Failed

    ' 解压到 payload 子目录（zip 副本与输出文件不进入 payload）
    Set zipNs = shell.NameSpace(ToLongPath(zipCopyPath))
    If zipNs Is Nothing Then GoTo Failed
    Set destNs = shell.NameSpace(ToLongPath(payloadDir))
    If destNs Is Nothing Then GoTo Failed
    destNs.CopyHere zipNs.Items(), 20
    If Not WaitForExtract(payloadDir, 60) Then GoTo Failed

    ' 删除字符格式与段落格式修订标记
    If Not StripFormatRevisions(payloadDir, removed) Then GoTo Failed

    ' 重新压缩
    If Not CreateEmptyZip(outZipPath) Then GoTo Failed
    Set outNs = shell.NameSpace(ToLongPath(outZipPath))
    If outNs Is Nothing Then GoTo Failed
    Set srcNs = shell.NameSpace(ToLongPath(payloadDir))
    If srcNs Is Nothing Then GoTo Failed
    outNs.CopyHere srcNs.Items(), 20
    If Not WaitForZipFile(outZipPath, 120) Then GoTo Failed

    ProcessOoxml = True
    Exit Function

Failed:
    ProcessOoxml = False
End Function

' 等待解压完成：payload\word 下 xml 文件数量稳定
Private Function WaitForExtract(ByVal payloadDir As String, ByVal timeoutSec As Long) As Boolean
    Dim fso As Object
    Dim wordDir As String
    Dim startTime As Single
    Dim lastCount As Long
    Dim count As Long

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    wordDir = payloadDir & "\word"
    startTime = Timer
    lastCount = -1

    Do
        If fso.FolderExists(wordDir) Then
            count = fso.GetFolder(wordDir).Files.count
            If count > 0 And count = lastCount Then
                WaitForExtract = True
                Exit Function
            End If
            lastCount = count
        End If
        If Timer - startTime > timeoutSec Then Exit Function
        DoEvents
    Loop
End Function

' 等待压缩完成：zip 文件大小稳定
Private Function WaitForZipFile(ByVal zipPath As String, ByVal timeoutSec As Long) As Boolean
    Dim fso As Object
    Dim startTime As Single
    Dim s1 As Double
    Dim s2 As Double

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    startTime = Timer

    Do
        If fso.FileExists(zipPath) Then
            s1 = fso.GetFile(zipPath).Size
            DoEvents
            s2 = fso.GetFile(zipPath).Size
            If s1 > 22 And s1 = s2 Then
                WaitForZipFile = True
                Exit Function
            End If
        End If
        If Timer - startTime > timeoutSec Then Exit Function
        DoEvents
    Loop
End Function

' 创建仅含 EOCD 的空 zip，供 Shell.Application 写入
Private Function CreateEmptyZip(ByVal zipPath As String) As Boolean
    Dim f As Integer
    Dim bytes(0 To 21) As Byte

    On Error GoTo Failed
    bytes(0) = 80
    bytes(1) = 75
    bytes(2) = 5
    bytes(3) = 6
    f = FreeFile
    Open zipPath For Binary As #f
    Put #f, , bytes
    Close #f
    CreateEmptyZip = True
    Exit Function

Failed:
    On Error Resume Next
    Close #f
    On Error GoTo 0
    CreateEmptyZip = False
End Function

' 删除 word/*.xml 中的 w:rPrChange 与 w:pPrChange
Private Function StripFormatRevisions(ByVal payloadDir As String, ByRef removed As Long) As Boolean
    Dim fso As Object
    Dim wordDir As String
    Dim file As Object
    Dim xml As Object
    Dim nodes As Object
    Dim i As Long
    Dim changed As Long

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    wordDir = payloadDir & "\word"
    If Not fso.FolderExists(wordDir) Then Exit Function

    For Each file In fso.GetFolder(wordDir).Files
        If LCase$(Right$(file.Name, 4)) = ".xml" Then
            changed = 0
            Set xml = CreateObject("MSXML2.DOMDocument.6.0")
            xml.async = False
            xml.validateOnParse = False
            xml.preserveWhiteSpace = True
            If xml.load(file.Path) Then
                xml.setProperty "SelectionNamespaces", W_NAMESPACE
                Set nodes = xml.selectNodes("//w:rPrChange | //w:pPrChange")
                If Not nodes Is Nothing Then
                    For i = nodes.Length - 1 To 0 Step -1
                        nodes.Item(i).parentNode.removeChild nodes.Item(i)
                        changed = changed + 1
                    Next i
                End If
                If changed > 0 Then
                    xml.save file.Path
                    removed = removed + changed
                End If
            End If
            Set xml = Nothing
        End If
    Next file

    StripFormatRevisions = True
End Function

' =============================================
' 回退路径：非 OOXML 文档，使用 Word 对象模型逐个接受
' =============================================
Private Sub AcceptFormatRevisionsByWordApi(ByVal doc As Document)
    Dim accepted As Long
    Dim remaining As Long
    Dim formatLeft As Long
    Dim contentLeft As Long
    Dim hasError As Boolean
    Dim errDesc As String
    Dim msgText As String

    BeginBatchUI
    Application.StatusBar = "正在接受格式类修订..."
    On Error Resume Next
    accepted = AcceptFormatRevisionsInDocument(doc)
    If Err.Number <> 0 Then
        hasError = True
        errDesc = Err.Description
        Err.Clear
    End If
    remaining = doc.Revisions.Count
    formatLeft = CountFormatRevisions(doc)
    Err.Clear
    On Error GoTo 0
    EndBatchUI

    contentLeft = remaining - formatLeft
    If contentLeft < 0 Then contentLeft = 0

    If hasError Then
        msgText = "接受格式修订时出现问题：" & errDesc & vbCrLf & vbCrLf & _
                  "本次已接受 " & accepted & " 处格式类修订。"
        If remaining > 0 Then
            msgText = msgText & vbCrLf & "文档中仍有 " & remaining & " 处修订未处理。"
        End If
    ElseIf accepted = 0 And remaining = 0 Then
        msgText = "当前文档没有修订。"
    ElseIf formatLeft > 0 Then
        msgText = "已接受 " & accepted & " 处格式类修订，但仍有 " & formatLeft & " 处格式类修订无法接受。" & vbCrLf & vbCrLf & _
                  "通常是因为这些格式修订与插入、删除等内容类修订重叠，Word 无法单独接受；" & vbCrLf & _
                  "请先处理相关的内容修订，或在「审阅」窗格中手动处理。"
    ElseIf accepted = 0 Then
        msgText = "未发现格式类修订（字符格式、段落格式）。" & vbCrLf & vbCrLf & _
                  "文档中仍有 " & contentLeft & " 处内容类修订（插入、删除、替换、移动等），已保持未接受状态。"
    Else
        msgText = "已接受 " & accepted & " 处格式类修订。"
        If contentLeft > 0 Then
            msgText = msgText & vbCrLf & vbCrLf & _
                      "文档中仍有 " & contentLeft & " 处内容类修订（插入、删除、替换、移动等），已保持未接受状态。"
        End If
    End If

    MsgBox msgText, vbInformation, DIALOG_TITLE
End Sub

' 接受文档中的格式类修订，返回接受数量
Private Function AcceptFormatRevisionsInDocument(ByVal doc As Document) As Long
    Dim rev As Revision
    Dim targets As Collection
    Dim accepted As Long
    Dim scanRound As Long
    Dim acceptedInRound As Long
    Dim i As Long
    Dim hasContentRev As Boolean

    On Error Resume Next
    For scanRound = 1 To 3
        Set targets = New Collection
        hasContentRev = False
        For Each rev In doc.Revisions
            If rev.Type = wdRevisionProperty Or rev.Type = wdRevisionParagraphProperty Then
                targets.Add rev
            Else
                hasContentRev = True
            End If
        Next rev

        If targets.count = 0 Then Exit For

        If Not hasContentRev Then
            Err.Clear
            doc.Revisions.AcceptAll
            If Err.Number = 0 Then
                accepted = accepted + targets.count
            End If
            Err.Clear
            Exit For
        End If

        acceptedInRound = 0
        For i = 1 To targets.count
            Err.Clear
            targets(i).Accept
            If Err.Number = 0 Then
                accepted = accepted + 1
                acceptedInRound = acceptedInRound + 1
            End If
            Err.Clear
            If (i Mod 100) = 0 Then
                Application.StatusBar = "正在接受格式类修订... " & i & "/" & targets.count
            End If
        Next i

        If acceptedInRound = 0 Then Exit For
    Next scanRound
    On Error GoTo 0

    AcceptFormatRevisionsInDocument = accepted
End Function

' 统计文档中剩余的格式类修订数量
Private Function CountFormatRevisions(ByVal doc As Document) As Long
    Dim rev As Revision
    Dim count As Long

    On Error Resume Next
    For Each rev In doc.Revisions
        If rev.Type = wdRevisionProperty Or rev.Type = wdRevisionParagraphProperty Then
            count = count + 1
        End If
    Next rev
    On Error GoTo 0

    CountFormatRevisions = count
End Function
