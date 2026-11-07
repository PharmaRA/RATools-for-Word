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
'      ZIP 由独立 PowerShell/.NET 进程同步读写，完整校验后替换。
'      策略禁止 PowerShell 或处理失败时回退到 Word 对象模型。
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
' 是否为 OOXML 文档（.docx / .docm）

Private Function IsOoxmlDocument(ByVal docPath As String) As Boolean
    Dim ext As String

    If Len(docPath) < 5 Then Exit Function
    ext = LCase$(Right$(docPath, 5))
    IsOoxmlDocument = (ext = ".docx" Or ext = ".docm")
End Function

' =============================================
' 快速路径：OOXML（同步 ZIP 处理、校验及原子替换）
' =============================================
Private Sub AcceptFormatRevisionsByOoxml(ByVal docPath As String)
    Dim doc As Document
    Dim fso As Object
    Dim backupPath As String
    Dim originalPath As String
    Dim recoveryPath As String
    Dim workDir As String
    Dim scriptPath As String
    Dim resultPath As String
    Dim ok As Boolean
    Dim removed As Long
    Dim errText As String
    Dim failureText As String
    Dim suffix As Long

    On Error GoTo Failed
    Set doc = ActiveDocument
    If Not doc.Saved Then doc.Save
    If Not doc.Saved Then Err.Raise vbObjectError + 2103, , "保存未完成，文档保持打开。"
    Set fso = CreateObject("Scripting.FileSystemObject")
    backupPath = docPath & ".bak"
    If AskMakeBackup(backupPath) Then
        Do While fso.FileExists(backupPath)
            suffix = suffix + 1
            backupPath = docPath & "." & CStr(suffix) & ".bak"
        Loop
        fso.CopyFile docPath, backupPath, False
    End If

    workDir = NewTempFolder(fso)
    If Len(workDir) = 0 Then Err.Raise vbObjectError + 2101, , "无法创建临时目录。"
    originalPath = workDir & "\original.docx"
    fso.CopyFile docPath, originalPath, False
    scriptPath = workDir & "\strip.ps1"
    resultPath = workDir & "\result.txt"
    recoveryPath = docPath & ".ratools-" & fso.GetFileName(workDir) & ".tmp.bak"
    If Not WriteStripScript(scriptPath) Then Err.Raise vbObjectError + 2102, , "无法创建处理脚本。"

    doc.Close wdDoNotSaveChanges
    Set doc = Nothing
    ok = RunStripScript(scriptPath, docPath, workDir, resultPath, removed, errText)
    On Error GoTo ReopenFailed
    Set doc = Application.Documents.Open(docPath)
    On Error GoTo Failed
    If ok Then
        MsgBox "已接受 " & removed & " 处格式修订，内容类修订保持未接受。", vbInformation, DIALOG_TITLE
    Else
        If fso.FileExists(recoveryPath) Then
            MsgBox "事务恢复未确认，已停止后续修改。原件保留在：" & vbCrLf & _
                   recoveryPath & vbCrLf & originalPath & vbCrLf & errText, vbCritical, DIALOG_TITLE
            Exit Sub
        End If
        MsgBox "快速处理失败，将对当前文档使用兼容模式。" & vbCrLf & errText, vbExclamation, DIALOG_TITLE
        AcceptFormatRevisionsByWordApi doc
    End If
    RecycleFormatWork recoveryPath
    RecycleFormatWork workDir
    Exit Sub

ReopenFailed:
    failureText = Err.Description
    ' Word rejects the completed package: restore the saved original before reopening.
    On Error Resume Next
    Err.Clear
    fso.CopyFile originalPath, docPath, True
    If Err.Number = 0 Then
        Set doc = Application.Documents.Open(docPath)
        If Err.Number = 0 Then
            MsgBox "处理后的文档无法打开，已恢复原文件。" & vbCrLf & failureText, vbExclamation, DIALOG_TITLE
            RecycleFormatWork recoveryPath
            RecycleFormatWork workDir
            Exit Sub
        End If
    End If
    MsgBox "无法自动恢复并打开文档。原件副本保留在：" & vbCrLf & originalPath, vbCritical, DIALOG_TITLE
    Exit Sub
Failed:
    failureText = Err.Description
    MsgBox "格式修订处理失败：" & failureText & vbCrLf & _
           "如已生成临时原件，其位置为：" & originalPath, vbExclamation, DIALOG_TITLE
End Sub

' Only recycle the operation's own temporary files; keep originals on recovery failures.
Private Sub RecycleFormatWork(ByVal path As String)
    Dim fso As Object
    Dim shell As Object
    Dim quotedPath As String
    On Error Resume Next
    If Len(path) = 0 Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) And Not fso.FolderExists(path) Then Exit Sub
    Set shell = CreateObject("Shell.Application")
    shell.Namespace(10).MoveHere path, 4 Or 16 Or 1024
End Sub

' 询问是否在修改前生成原文件备份（默认"是"）
Private Function AskMakeBackup(ByVal backupPath As String) As Boolean
    AskMakeBackup = (MsgBox("是否在修改前生成原文件备份？" & vbCrLf & vbCrLf & _
                            "是：备份为 " & backupPath & vbCrLf & _
                            "否：不保留额外备份；处理期间仍保留恢复副本", _
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

' 生成同步 ZIP 处理脚本（纯 ASCII；文档路径通过参数传递）
Private Function WriteStripScript(ByVal scriptPath As String) As Boolean
    Dim f As Integer
    On Error GoTo Failed
    f = FreeFile
    Open scriptPath For Output As #f
    Print #f, "param([string]$DocumentPath, [string]$WorkDirectory, [string]$ResultPath)"
    Print #f, "$ErrorActionPreference = 'Stop'"
    Print #f, "Add-Type -AssemblyName System.IO.Compression.FileSystem"
    Print #f, "Add-Type -AssemblyName Microsoft.VisualBasic"
    Print #f, "$ready = $DocumentPath + '.ratools-' + [IO.Path]::GetFileName($WorkDirectory) + '.tmp'"
    Print #f, "$recovery = $ready + '.bak'"
    Print #f, "$committed = $false"
    Print #f, "$readyOwned = $false"
    Print #f, "$removed = 0"
    Print #f, "function Read-Part($entry) {"
    Print #f, "    $inputStream = $entry.Open()"
    Print #f, "    $buffer = [IO.MemoryStream]::new()"
    Print #f, "    try { $inputStream.CopyTo($buffer); return ,$buffer.ToArray() }"
    Print #f, "    finally { $inputStream.Dispose(); $buffer.Dispose() }"
    Print #f, "}"
    Print #f, "function Parse-Xml([byte[]]$bytes) {"
    Print #f, "    $stream = [IO.MemoryStream]::new($bytes, $false)"
    Print #f, "    $settings = [Xml.XmlReaderSettings]::new()"
    Print #f, "    $settings.DtdProcessing = [Xml.DtdProcessing]::Prohibit"
    Print #f, "    $settings.XmlResolver = $null"
    Print #f, "    $reader = [Xml.XmlReader]::Create($stream, $settings)"
    Print #f, "    try {"
    Print #f, "        $xml = [Xml.XmlDocument]::new()"
    Print #f, "        $xml.PreserveWhitespace = $true"
    Print #f, "        $xml.XmlResolver = $null"
    Print #f, "        $xml.Load($reader)"
    Print #f, "        return ,$xml"
    Print #f, "    } finally { $reader.Dispose(); $stream.Dispose() }"
    Print #f, "}"
    Print #f, "try {"
    Print #f, "    if ((Test-Path -LiteralPath $ready) -or (Test-Path -LiteralPath $recovery)) { throw 'Transaction paths already exist' }"
    Print #f, "    $before = (Get-FileHash -LiteralPath $DocumentPath -Algorithm SHA256).Hash"
    Print #f, "    [IO.File]::Copy($DocumentPath, $ready, $false)"
    Print #f, "    $readyOwned = $true"
    Print #f, "    $archive = [IO.Compression.ZipFile]::Open($ready, 'Update')"
    Print #f, "    $manifest = @{}"
    Print #f, "    try {"
    Print #f, "        foreach ($required in @('[Content_Types].xml', '_rels/.rels', 'word/document.xml')) {"
    Print #f, "            if ($null -eq $archive.GetEntry($required)) { throw ""Missing package part: $required"" }"
    Print #f, "        }"
    Print #f, "        foreach ($entry in $archive.Entries) {"
    Print #f, "            if ($manifest.ContainsKey($entry.FullName)) { throw 'Duplicate package part' }"
    Print #f, "            $bytes = Read-Part $entry"
    Print #f, "            if ($entry.FullName -match '\.(xml|rels)$') {"
    Print #f, "                $xml = Parse-Xml $bytes"
    Print #f, "                if ($entry.FullName -match '^word/[^/]+\.xml$') {"
    Print #f, "                    $ns = [Xml.XmlNamespaceManager]::new($xml.NameTable)"
    Print #f, "                    $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')"
    Print #f, "                    $nodes = @($xml.SelectNodes('//w:rPrChange | //w:pPrChange', $ns))"
    Print #f, "                    foreach ($node in $nodes) { [void]$node.ParentNode.RemoveChild($node); $removed++ }"
    Print #f, "                    if ($nodes.Count -gt 0) {"
    Print #f, "                        $output = [IO.MemoryStream]::new()"
    Print #f, "                        try { $xml.Save($output); $bytes = $output.ToArray() } finally { $output.Dispose() }"
    Print #f, "                        $entryStream = $entry.Open()"
    Print #f, "                        try { $entryStream.SetLength(0); $entryStream.Write($bytes, 0, $bytes.Length) } finally { $entryStream.Dispose() }"
    Print #f, "                    }"
    Print #f, "                }"
    Print #f, "            }"
    Print #f, "            $sha = [Security.Cryptography.SHA256]::Create()"
    Print #f, "            try { $manifest[$entry.FullName] = [Convert]::ToBase64String($sha.ComputeHash($bytes)) } finally { $sha.Dispose() }"
    Print #f, "        }"
    Print #f, "    } finally { $archive.Dispose() }"
    Print #f, "    # Reopen the fully closed package and verify every byte of every expected part."
    Print #f, "    $archive = [IO.Compression.ZipFile]::OpenRead($ready)"
    Print #f, "    try {"
    Print #f, "        if ($archive.Entries.Count -ne $manifest.Count) { throw 'Package part count changed' }"
    Print #f, "        foreach ($entry in $archive.Entries) {"
    Print #f, "            $bytes = Read-Part $entry"
    Print #f, "            $sha = [Security.Cryptography.SHA256]::Create()"
    Print #f, "            try { $hash = [Convert]::ToBase64String($sha.ComputeHash($bytes)) } finally { $sha.Dispose() }"
    Print #f, "            if ($manifest[$entry.FullName] -ne $hash) { throw ""Package verification failed: $($entry.FullName)"" }"
    Print #f, "        }"
    Print #f, "    } finally { $archive.Dispose() }"
    Print #f, "    if ((Get-FileHash -LiteralPath $DocumentPath -Algorithm SHA256).Hash -ne $before) { throw 'Original document changed during processing' }"
    Print #f, "    # Check the result channel before committing; replace and backup are on the same volume."
    Print #f, "    [IO.File]::WriteAllText($ResultPath, 'PREPARED', [Text.Encoding]::ASCII)"
    Print #f, "    [IO.File]::Replace($ready, $DocumentPath, $recovery)"
    Print #f, "    $committed = $true"
    Print #f, "    [IO.File]::WriteAllText($ResultPath, ('OK=' + $removed), [Text.Encoding]::ASCII)"
    Print #f, "    exit 0"
    Print #f, "} catch {"
    Print #f, "    $failure = $_.Exception.Message"
    Print #f, "    if ($committed) {"
    Print #f, "        try { [IO.File]::Replace($recovery, $DocumentPath, $ready) }"
    Print #f, "        catch { $failure += '; rollback failed, recovery copy: ' + $recovery }"
    Print #f, "    }"
    Print #f, "    [IO.File]::WriteAllText($ResultPath, ('ERROR=' + $failure), [Text.Encoding]::Unicode)"
    Print #f, "    exit 1"
    Print #f, "} finally {"
    Print #f, "    if ($readyOwned -and (Test-Path -LiteralPath $ready)) {"
    Print #f, "        try { [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($ready, 'OnlyErrorDialogs', 'SendToRecycleBin') } catch { }"
    Print #f, "    }"
    Print #f, "}"
    Close #f
    WriteStripScript = True
    Exit Function
Failed:
    On Error Resume Next
    Close #f
End Function

' 同步调用隔离的处理进程；策略禁止 PowerShell 时回退 Word API
Private Function RunStripScript(ByVal scriptPath As String, ByVal docPath As String, _
                                ByVal workDir As String, ByVal resultPath As String, _
                                ByRef removed As Long, ByRef errText As String) As Boolean
    Dim shell As Object
    Dim cmd As String
    Dim exitCode As Long
    Dim fso As Object
    Dim resultFile As Object
    On Error GoTo Failed
    Set shell = CreateObject("WScript.Shell")
    cmd = "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File """ & scriptPath & _
          """ -DocumentPath """ & docPath & """ -WorkDirectory """ & workDir & """ -ResultPath """ & resultPath & """"
    exitCode = shell.Run(cmd, 0, True)
    If exitCode = 0 Then
        removed = ReadRemovedCount(resultPath)
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set resultFile = fso.OpenTextFile(resultPath, 1)
        RunStripScript = (Left$(resultFile.ReadAll, 3) = "OK=")
        resultFile.Close
    Else
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(resultPath) Then
            Set resultFile = fso.OpenTextFile(resultPath, 1, False, -1)
            errText = resultFile.ReadAll
            resultFile.Close
        Else
            errText = "处理程序未完成，退出码 " & exitCode
        End If
    End If
    Exit Function
Failed:
    errText = Err.Description
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
