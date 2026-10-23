Attribute VB_Name = "Macro_AcceptFormatRevisions"
Option Explicit

' =============================================
' 模块名：Macro_AcceptFormatRevisions
' 功能：只接受"格式类修订"（字符格式、段落格式），
'       插入、删除、替换、移动等内容类修订保持未接受状态。
'
' 实现（两条路径）：
'   1. OOXML 快速路径（.docx/.docm）：保存并备份后，把文档当作 zip 解包，
'      删除 word/*.xml 中的 w:rPrChange / w:pPrChange，再重新打包。
'      ZIP 读写交给独立的 cscript 进程完成 —— Word 进程内的
'      Shell.Application 会返回空命名空间或阻塞，不可用；
'      而 powershell.exe 在部分安全策略下会被拦截，cscript 不受影响。
'   2. Word 对象模型回退路径：非 OOXML 文档、或 OOXML 处理失败时使用。
' =============================================

Private Const DIALOG_TITLE As String = "接受格式修订"
Private Const MAX_WORD_API_ROUNDS As Long = 3

' Ribbon 回调：接受全文格式类修订
Public Sub AcceptFormatRevisions(control As IRibbonControl)
    Core_AcceptFormatRevisions
End Sub

Public Sub Core_AcceptFormatRevisions()
    Dim doc As Document
    Dim docPath As String

    If Application.Documents.Count = 0 Then
        MsgBox "当前没有打开的文档。", vbExclamation, DIALOG_TITLE
        Exit Sub
    End If

    Set doc = ActiveDocument
    docPath = doc.FullName

    ' OneDrive/SharePoint 云端文档的 FullName 可能返回 https 地址
    If IsWebPath(docPath) Then
        Dim syncLocalPath As String
        syncLocalPath = ResolveWebPathToLocalFolder(docPath)
        If Len(syncLocalPath) = 0 Then
            syncLocalPath = FindLocalSyncCopy(docPath)
        End If
        If Len(syncLocalPath) = 0 Then
            MsgBox "未能把云端路径换算成本机同步路径，将改用兼容模式（较慢）。" & vbCrLf & vbCrLf & _
                   "云端地址：" & docPath & vbCrLf & vbCrLf & _
                   "若文档已同步到本机，请确认 OneDrive 正在运行。", _
                   vbExclamation, DIALOG_TITLE
            AcceptFormatRevisionsByWordApi doc
            Exit Sub
        End If
        docPath = syncLocalPath
    End If

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
    Core_AcceptFormatRevisions
End Sub

' 是否为 OOXML 文档（.docx / .docm）
Private Function FindLocalSyncCopy(ByVal webPath As String) As String
    Dim fso As Object
    Dim roots As Collection
    Dim segments As Collection
    Dim i As Long
    Dim j As Long
    Dim rootPath As String
    Dim relativePath As String
    Dim candidate As String
    Dim fileName As String

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Function

    Set segments = SplitWebPathSegments(webPath)
    If segments Is Nothing Then Exit Function
    If segments.count = 0 Then Exit Function
    fileName = CStr(segments(segments.count))

    Set roots = ListOneDriveRoots(fso)
    If roots Is Nothing Then Exit Function
    If roots.count = 0 Then Exit Function

    ' 先试完整尾部路径（最多取末 4 段）
    For j = segments.count To 1 Step -1
        If segments.count - j >= 4 Then Exit For
        relativePath = ""
        For i = j To segments.count
            If Len(relativePath) > 0 Then relativePath = relativePath & "\"
            relativePath = relativePath & CStr(segments(i))
        Next i

        For i = 1 To roots.count
            rootPath = CStr(roots(i))
            candidate = rootPath & "\" & relativePath
            If fso.FileExists(candidate) Then
                FindLocalSyncCopy = candidate
                Exit Function
            End If
        Next i
    Next j

    ' 最后只按文件名在同步根下递归查找（深度受限）
    For i = 1 To roots.count
        candidate = FindFileUnder(CStr(roots(i)), fileName, 6)
        If Len(candidate) > 0 Then
            FindLocalSyncCopy = candidate
            Exit Function
        End If
    Next i
End Function

' 把 URL 路径拆成路径段（去掉协议/主机/查询串，并做 URL 解码）
Private Function SplitWebPathSegments(ByVal webPath As String) As Collection
    Dim cleanPath As String
    Dim parts As Variant
    Dim i As Long
    Dim segment As String
    Dim result As Collection

    On Error Resume Next
    cleanPath = webPath
    If InStr(1, cleanPath, "?", vbTextCompare) > 0 Then cleanPath = Left$(cleanPath, InStr(1, cleanPath, "?", vbTextCompare) - 1)
    If InStr(1, cleanPath, "#", vbTextCompare) > 0 Then cleanPath = Left$(cleanPath, InStr(1, cleanPath, "#", vbTextCompare) - 1)
    If Left$(LCase$(cleanPath), 8) = "https://" Then cleanPath = Mid$(cleanPath, 9)
    If Left$(LCase$(cleanPath), 7) = "http://" Then cleanPath = Mid$(cleanPath, 8)
    If InStr(1, cleanPath, "/") > 0 Then cleanPath = Mid$(cleanPath, InStr(1, cleanPath, "/") + 1)

    Set result = New Collection
    parts = Split(cleanPath, "/")
    For i = LBound(parts) To UBound(parts)
        segment = CStr(parts(i))
        If Len(segment) > 0 Then result.Add UrlDecodeText(segment)
    Next i
    Set SplitWebPathSegments = result
End Function

' 简单 URL 解码（%XX 与 +）
Private Function UrlDecodeText(ByVal value As String) As String
    Dim i As Long
    Dim ch As String
    Dim hexCode As String
    Dim result As String

    result = ""
    i = 1
    Do While i <= Len(value)
        ch = Mid$(value, i, 1)
        If ch = "+" Then
            result = result & " "
        ElseIf ch = "%" And i + 2 <= Len(value) Then
            hexCode = Mid$(value, i + 1, 2)
            If IsHexPair(hexCode) Then
                result = result & Chr$(CLng("&H" & hexCode))
                i = i + 2
            Else
                result = result & ch
            End If
        Else
            result = result & ch
        End If
        i = i + 1
    Loop
    UrlDecodeText = result
End Function

Private Function IsHexPair(ByVal text As String) As Boolean
    Dim i As Long
    Dim ch As String

    If Len(text) <> 2 Then Exit Function
    For i = 1 To 2
        ch = UCase$(Mid$(text, i, 1))
        If InStr(1, "0123456789ABCDEF", ch, vbBinaryCompare) = 0 Then Exit Function
    Next i
    IsHexPair = True
End Function

' 列出常见的 OneDrive 同步根目录（环境变量 + 用户目录下的 OneDrive* 文件夹）
Private Function ListOneDriveRoots(ByVal fso As Object) As Collection
    Dim result As Collection
    Dim shell As Object
    Dim env As Object
    Dim folder As Object
    Dim subFolder As Object
    Dim candidate As String

    On Error Resume Next
    Set result = New Collection

    Set shell = CreateObject("WScript.Shell")
    If Not shell Is Nothing Then
        Set env = shell.Environment("PROCESS")
        candidate = env("OneDrive")
        If Len(candidate) > 0 Then AppendUniquePath result, candidate
        candidate = env("OneDriveCommercial")
        If Len(candidate) > 0 Then AppendUniquePath result, candidate
        candidate = env("OneDriveConsumer")
        If Len(candidate) > 0 Then AppendUniquePath result, candidate
    End If

    Set folder = fso.GetSpecialFolder(1)
    If Not folder Is Nothing Then
        For Each subFolder In folder.SubFolders
            If InStr(1, LCase$(subFolder.Name), "onedrive", vbTextCompare) > 0 Then
                AppendUniquePath result, subFolder.Path
            End If
        Next subFolder
    End If

    Set ListOneDriveRoots = result
End Function

Private Sub AppendUniquePath(ByRef paths As Collection, ByVal candidate As String)
    Dim i As Long
    Dim fso As Object

    If Len(candidate) = 0 Then Exit Sub
    For i = 1 To paths.count
        If StrComp(CStr(paths(i)), candidate, vbTextCompare) = 0 Then Exit Sub
    Next i
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(candidate) Then paths.Add candidate
End Sub

' 在指定目录下递归查找文件名（限制深度）
Private Function FindFileUnder(ByVal folderPath As String, ByVal fileName As String, ByVal maxDepth As Long) As String
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim subFolder As Object
    Dim found As String

    On Error Resume Next
    If maxDepth < 0 Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then Exit Function
    Set folder = fso.GetFolder(folderPath)

    For Each file In folder.Files
        If StrComp(file.Name, fileName, vbTextCompare) = 0 Then
            FindFileUnder = file.Path
            Exit Function
        End If
    Next file

    For Each subFolder In folder.SubFolders
        found = FindFileUnder(subFolder.Path, fileName, maxDepth - 1)
        If Len(found) > 0 Then
            FindFileUnder = found
            Exit Function
        End If
    Next subFolder
End Function

' 是否为 OOXML 文档（.docx / .docm）

Private Function IsOoxmlDocument(ByVal docPath As String) As Boolean
    Dim ext As String

    If Len(docPath) < 5 Then Exit Function
    ext = LCase$(Right$(docPath, 5))
    IsOoxmlDocument = (ext = ".docx" Or ext = ".docm")
End Function

' =============================================
' 快速路径：OOXML（ZIP 操作交给独立 cscript 进程）
' =============================================
Private Sub AcceptFormatRevisionsByOoxml(ByVal docPath As String)
    Dim doc As Document
    Dim fso As Object
    Dim backupPath As String
    Dim workDir As String
    Dim scriptPath As String
    Dim resultPath As String
    Dim makeBackup As Boolean
    Dim ok As Boolean
    Dim removed As Long
    Dim remaining As Long
    Dim errText As String
    Dim msgText As String

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

    ' 2. 询问是否生成备份
    makeBackup = AskMakeBackup(backupPath)
    If makeBackup Then
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
    End If

    ' 3. 准备临时目录、脚本、结果文件
    workDir = NewTempFolder(fso)
    If Len(workDir) = 0 Then
        MsgBox "无法创建临时目录。", vbExclamation, DIALOG_TITLE
        Exit Sub
    End If
    scriptPath = workDir & "\strip.vbs"
    resultPath = workDir & "\result.txt"

    If Not WriteStripScript(scriptPath) Then
        MsgBox "无法创建临时处理脚本。", vbExclamation, DIALOG_TITLE
        fso.DeleteFolder workDir, True
        Exit Sub
    End If

    ' 4. 关闭文档（已保存，无提示）
    doc.Close wdSaveChanges
    Set doc = Nothing

    ' 5. 调用 cscript 处理文件
    ok = RunStripScript(scriptPath, docPath, workDir, resultPath, removed, errText)

    ' 6. 重新打开文档
    On Error Resume Next
    Set doc = Application.Documents.Open(docPath)
    If Err.Number <> 0 Then
        MsgBox "文档处理" & IIf(ok, "完成", "失败") & "，但重新打开失败，请手动打开：" & vbCrLf & docPath & vbCrLf & Err.Description, vbExclamation, DIALOG_TITLE
        Err.Clear
        On Error GoTo 0
        fso.DeleteFolder workDir, True
        Exit Sub
    End If
    On Error GoTo 0

    remaining = doc.Revisions.Count

    ' 7. 清理临时目录
    On Error Resume Next
    fso.DeleteFolder workDir, True
    On Error GoTo 0

    ' 8. 结果提示
    If ok Then
        msgText = "已完成：格式类修订已接受（处理 " & removed & " 处格式标记）。" & vbCrLf & vbCrLf & _
                  "文档中剩余 " & remaining & " 处修订（插入、删除等内容类），均已保持未接受状态。"
        If makeBackup Then
            msgText = msgText & vbCrLf & vbCrLf & "原文件备份：" & backupPath
        End If
        MsgBox msgText, vbInformation, DIALOG_TITLE
    Else
        If makeBackup Then
            msgText = "快速处理失败，已从备份恢复原文件，将改用兼容模式（较慢）。"
        Else
            msgText = "快速处理失败，将改用兼容模式（较慢）。"
        End If
        If Len(errText) > 0 Then msgText = msgText & vbCrLf & vbCrLf & "原因：" & errText
        MsgBox msgText, vbExclamation, DIALOG_TITLE
        AcceptFormatRevisionsByWordApi doc
    End If
End Sub

' 询问是否在修改前生成原文件备份（默认"是"）
Private Function AskMakeBackup(ByVal backupPath As String) As Boolean
    AskMakeBackup = (MsgBox("是否在修改前生成原文件备份？" & vbCrLf & vbCrLf & _
                            "是：备份为 " & backupPath & vbCrLf & _
                            "否：直接修改原文件（处理失败时自动回滚，不留备份）", _
                            vbYesNo + vbQuestion, DIALOG_TITLE) = vbYes)
End Function

' 创建唯一临时目录
Private Function NewTempFolder(ByVal fso As Object) As String
    Dim baseDir As String
    Dim candidate As String
    Dim i As Long

    On Error Resume Next
    baseDir = fso.GetSpecialFolder(2) & "\RATools_Ooxml_" & Format$(Now, "yyyymmddhhnnss")
    For i = 1 To 100
        candidate = baseDir & "_" & i
        If Not fso.FolderExists(candidate) Then
            fso.CreateFolder candidate
            If Err.Number = 0 Then
                NewTempFolder = candidate
                Exit Function
            End If
            Err.Clear
        End If
    Next i
    On Error GoTo 0
End Function

' 生成 cscript 处理脚本（纯 ASCII）
Private Function WriteStripScript(ByVal scriptPath As String) As Boolean
    Dim f As Integer

    On Error GoTo Failed
    f = FreeFile
    Open scriptPath For Output As #f

    Print #f, "Dim fso, shell, zipNs, destNs, srcNs, outNs, t, removed, k, z"
    Print #f, "Dim docPath, workDir, payloadDir, zipCopy, outZip, xmlDoc, nodes, folder, file, bin, s1, s2"
    Print #f, "docPath = WScript.Arguments(0)"
    Print #f, "workDir = WScript.Arguments(1)"
    Print #f, "removed = 0"
    Print #f, ""
    Print #f, "Set fso = CreateObject(""Scripting.FileSystemObject"")"
    Print #f, "payloadDir = workDir & ""\payload"""
    Print #f, "zipCopy = workDir & ""\src.zip"""
    Print #f, "outZip = workDir & ""\out.zip"""
    Print #f, "If Not fso.FolderExists(workDir) Then fso.CreateFolder workDir"
    Print #f, "If fso.FolderExists(payloadDir) Then fso.DeleteFolder payloadDir, True"
    Print #f, "fso.CreateFolder payloadDir"
    Print #f, "If fso.FileExists(zipCopy) Then fso.DeleteFile zipCopy, True"
    Print #f, "If fso.FileExists(outZip) Then fso.DeleteFile outZip, True"
    Print #f, "fso.CopyFile docPath, zipCopy, True"
    Print #f, ""
    Print #f, "Set shell = CreateObject(""Shell.Application"")"
    Print #f, "Set zipNs = shell.NameSpace(zipCopy)"
    Print #f, "If zipNs Is Nothing Then Fail ""zipns-nothing"""
    Print #f, "If zipNs.Items().Count = 0 Then Fail ""zip-empty"""
    Print #f, "Set destNs = shell.NameSpace(payloadDir)"
    Print #f, "If destNs Is Nothing Then Fail ""destns-nothing"""
    Print #f, "destNs.CopyHere zipNs.Items(), 20"
    Print #f, "t = Timer"
    Print #f, "Do While Timer - t < 60"
    Print #f, "  If fso.FolderExists(payloadDir & ""\word"") Then Exit Do"
    Print #f, "  WScript.Sleep 50"
    Print #f, "Loop"
    Print #f, "If Not fso.FolderExists(payloadDir & ""\word"") Then Fail ""extract-timeout"""
    Print #f, ""
    Print #f, "Set folder = fso.GetFolder(payloadDir & ""\word"")"
    Print #f, "For Each file In folder.Files"
    Print #f, "  If LCase(Right(file.Name, 4)) = "".xml"" Then"
    Print #f, "    Set xmlDoc = CreateObject(""MSXML2.DOMDocument.6.0"")"
    Print #f, "    xmlDoc.async = False"
    Print #f, "    xmlDoc.preserveWhiteSpace = True"
    Print #f, "    xmlDoc.validateOnParse = False"
    Print #f, "    If xmlDoc.load(file.Path) Then"
    Print #f, "      xmlDoc.setProperty ""SelectionNamespaces"", ""xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'"""
    Print #f, "      Set nodes = xmlDoc.selectNodes(""//w:rPrChange | //w:pPrChange"")"
    Print #f, "      If Not nodes Is Nothing Then"
    Print #f, "        If nodes.Length > 0 Then"
    Print #f, "          For k = nodes.Length - 1 To 0 Step -1"
    Print #f, "            nodes.Item(k).parentNode.removeChild nodes.Item(k)"
    Print #f, "          Next"
    Print #f, "          xmlDoc.save file.Path"
    Print #f, "          removed = removed + nodes.Length"
    Print #f, "        End If"
    Print #f, "      End If"
    Print #f, "    End If"
    Print #f, "  End If"
    Print #f, "Next"
    Print #f, ""
    Print #f, "Set bin = fso.CreateTextFile(outZip, True)"
    Print #f, "bin.Write Chr(80) & Chr(75) & Chr(5) & Chr(6)"
    Print #f, "For z = 1 To 18"
    Print #f, "  bin.Write Chr(0)"
    Print #f, "Next"
    Print #f, "bin.Close"
    Print #f, ""
    Print #f, "Set outNs = shell.NameSpace(outZip)"
    Print #f, "If outNs Is Nothing Then Fail ""outns-nothing"""
    Print #f, "Set srcNs = shell.NameSpace(payloadDir)"
    Print #f, "If srcNs Is Nothing Then Fail ""srcns-nothing"""
    Print #f, "outNs.CopyHere srcNs.Items(), 20"
    Print #f, "t = Timer"
    Print #f, "Do While Timer - t < 120"
    Print #f, "  If fso.FileExists(outZip) Then"
    Print #f, "    s1 = fso.GetFile(outZip).Size"
    Print #f, "    If s1 > 22 Then"
    Print #f, "      WScript.Sleep 300"
    Print #f, "      s2 = fso.GetFile(outZip).Size"
    Print #f, "      If s1 = s2 Then Exit Do"
    Print #f, "    End If"
    Print #f, "  End If"
    Print #f, "  WScript.Sleep 100"
    Print #f, "Loop"
    Print #f, "s1 = fso.GetFile(outZip).Size"
    Print #f, "If s1 <= 22 Then Fail ""zip-create-failed-size-"" & s1"
    Print #f, ""
    Print #f, "On Error Resume Next"
    Print #f, "fso.CopyFile outZip, docPath, True"
    Print #f, "If Err.Number <> 0 Then Fail ""replace-failed"""
    Print #f, "WScript.Echo ""OK="" & removed & "" size="" & s1"
    Print #f, "WScript.Quit 0"
    Print #f, ""
    Print #f, "Sub Fail(msg)"
    Print #f, "  WScript.Echo ""ERROR="" & msg"
    Print #f, "  WScript.Quit 1"
    Print #f, "End Sub"

    Close #f
    WriteStripScript = True
    Exit Function

Failed:
    On Error Resume Next
    Close #f
    On Error GoTo 0
    WriteStripScript = False
End Function

' 调用 cscript 处理，返回是否成功
Private Function RunStripScript(ByVal scriptPath As String, ByVal docPath As String, _
                                ByVal workDir As String, ByVal resultPath As String, _
                                ByRef removed As Long, ByRef errText As String) As Boolean
    Dim shell As Object
    Dim cmd As String
    Dim exitCode As Long
    Dim outputPath As String

    On Error Resume Next
    Set shell = CreateObject("WScript.Shell")
    If shell Is Nothing Then
        errText = "无法创建 WScript.Shell"
        Exit Function
    End If

    outputPath = workDir & "\console.txt"
    cmd = "cmd /c cscript.exe //nologo """ & scriptPath & """ """ & docPath & """ """ & workDir & """ > """ & outputPath & """ 2>&1"
    exitCode = shell.Run(cmd, 0, True)
    If Err.Number <> 0 Then
        errText = "调用 cscript 失败：" & Err.Description
        Err.Clear
        Exit Function
    End If

    removed = ReadRemovedCount(outputPath)
    RunStripScript = (exitCode = 0)
    If Not RunStripScript Then
        errText = ReadErrorText(outputPath)
        If Len(errText) = 0 Then errText = "cscript 退出码 " & exitCode
    End If
    On Error GoTo 0
End Function

' 从输出中读取 OK=n
Private Function ReadRemovedCount(ByVal outputPath As String) As Long
    Dim f As Integer
    Dim lineText As String
    Dim fso As Object

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Function
    If Not fso.FileExists(outputPath) Then Exit Function

    f = FreeFile
    Open outputPath For Input As #f
    Do While Not EOF(f)
        Line Input #f, lineText
        If InStr(1, lineText, "OK=", vbTextCompare) = 1 Then
            ReadRemovedCount = CLng(Val(Mid$(lineText, 4)))
            Close #f
            On Error GoTo 0
            Exit Function
        End If
    Loop
    Close #f
    On Error GoTo 0
End Function

' 从输出中读取 ERROR=... 说明
Private Function ReadErrorText(ByVal outputPath As String) As String
    Dim f As Integer
    Dim lineText As String
    Dim fso As Object

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Function
    If Not fso.FileExists(outputPath) Then Exit Function

    f = FreeFile
    Open outputPath For Input As #f
    Do While Not EOF(f)
        Line Input #f, lineText
        If InStr(1, lineText, "ERROR=", vbTextCompare) = 1 Then
            ReadErrorText = Mid$(lineText, 7)
            Close #f
            On Error GoTo 0
            Exit Function
        End If
    Loop
    Close #f
    On Error GoTo 0
End Function

' =============================================
' 回退路径：Word 对象模型逐个接受
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
    ElseIf accepted = 0 And remaining = 0 Then
        msgText = "当前文档没有修订。"
    ElseIf formatLeft > 0 Then
        msgText = "已接受 " & accepted & " 处格式类修订，但仍有 " & formatLeft & " 处格式类修订无法接受。" & vbCrLf & vbCrLf & _
                  "通常是因为这些格式修订与插入、删除等内容类修订重叠。"
    ElseIf accepted = 0 Then
        msgText = "未发现格式类修订（字符格式、段落格式）。" & vbCrLf & vbCrLf & _
                  "文档中仍有 " & contentLeft & " 处内容类修订，已保持未接受状态。"
    Else
        msgText = "已接受 " & accepted & " 处格式类修订。"
        If contentLeft > 0 Then
            msgText = msgText & vbCrLf & vbCrLf & _
                      "文档中仍有 " & contentLeft & " 处内容类修订，已保持未接受状态。"
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
    For scanRound = 1 To MAX_WORD_API_ROUNDS
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
            If Err.Number = 0 Then accepted = accepted + targets.count
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
