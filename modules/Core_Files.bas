Attribute VB_Name = "Core_Files"
Option Explicit

' =============================================
' 云端路径解析读取的 OneDrive 同步登记位置（模块级声明必须在过程之前）
' =============================================
Private Const ONEDRIVE_SYNC_REG_KEY As String = "HKCU\Software\SyncEngines\Providers\OneDrive"
Private Const ONEDRIVE_ACCOUNTS_REG_KEY As String = "HKCU\Software\Microsoft\OneDrive\Accounts"

' =============================================
' 公共层：文件系统工具
' =============================================

' 判断扩展名是否为 Word 文档（doc/docx/docm）
Public Function IsWordDocumentExt(ByVal extName As String) As Boolean
    Select Case LCase$(extName)
        Case "doc", "docx", "docm"
            IsWordDocumentExt = True
        Case Else
            IsWordDocumentExt = False
    End Select
End Function

' 递归收集文件夹内全部 Word 文档路径（排除 ~$ 临时文件）
Public Sub CollectWordFiles(ByVal folderPath As String, _
                            ByRef fileCollection As Collection, _
                            Optional ByVal recursive As Boolean = True)
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    CollectWordFilesCore fso, folderPath, fileCollection, recursive
    Set fso = Nothing
End Sub

Private Sub CollectWordFilesCore(ByVal fso As Object, _
                                 ByVal folderPath As String, _
                                 ByRef fileCollection As Collection, _
                                 ByVal recursive As Boolean)
    Dim targetFolder As Object
    Dim subFolder As Object
    Dim fileItem As Object

    On Error Resume Next
    Set targetFolder = fso.GetFolder(folderPath)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    For Each fileItem In targetFolder.Files
        If IsWordDocumentExt(fso.GetExtensionName(fileItem.Name)) Then
            If Left$(fileItem.Name, 2) <> "~$" Then fileCollection.Add fileItem.Path
        End If
    Next fileItem

    If recursive Then
        For Each subFolder In targetFolder.SubFolders
            CollectWordFilesCore fso, subFolder.Path, fileCollection, True
        Next subFolder
    End If
End Sub

' =============================================
' 云端路径解析：OneDrive/SharePoint 的 Web 路径 -> 本机同步文件夹
'
' Word 对存放在 OneDrive/SharePoint 的文档，Document.Path 返回 https 地址；
' 直接把该地址交给 explorer.exe 只会打开浏览器。这里读取 OneDrive 同步引擎
' 在 HKCU\Software\SyncEngines\Providers\OneDrive 下登记的
' UrlNamespace -> MountPoint 映射，把 Web 路径换算成本地同步文件夹。
'
' 注册表内容用 Windows 自带的 reg.exe 读取（/reg:64 与 /reg:32 两个视图都查，
' 32 位 Office 也能读到 64 位视图的登记），不依赖 WMI。
' =============================================

' 判断路径是否为 Web 地址（http/https）
Public Function IsWebPath(ByVal pathValue As String) As Boolean
    IsWebPath = (InStr(1, pathValue, "http://", vbTextCompare) = 1) Or _
                (InStr(1, pathValue, "https://", vbTextCompare) = 1)
End Function

' 把 Web 路径换算为本机同步文件夹；无法解析或文件夹不存在时返回空串
Public Function ResolveWebPathToLocalFolder(ByVal webPath As String) As String
    Dim urlNamespaces As Collection
    Dim localRoots As Collection
    Dim resolvedPath As String

    If Not IsWebPath(webPath) Then Exit Function

    Set urlNamespaces = New Collection
    Set localRoots = New Collection
    LoadOneDriveSyncMappings urlNamespaces, localRoots
    If urlNamespaces.Count = 0 Then Exit Function

    resolvedPath = MapWebPathToLocalFolder(webPath, urlNamespaces, localRoots)
    If Len(resolvedPath) = 0 Then Exit Function
    If Not FolderExistsSync(resolvedPath) Then
        If Not FileExistsSync(resolvedPath) Then Exit Function
    End If

    ResolveWebPathToLocalFolder = resolvedPath
End Function

' 按同步映射把 Web 路径换算为本地路径；多条映射命中时取最长前缀；无命中返回空串
Public Function MapWebPathToLocalFolder(ByVal webPath As String, _
                                        ByVal urlNamespaces As Collection, _
                                        ByVal localRoots As Collection) As String
    Dim i As Long
    Dim mappedPath As String
    Dim bestNamespaceLength As Long
    Dim bestMappedPath As String

    For i = 1 To urlNamespaces.Count
        mappedPath = MapSingleWebPath(webPath, CStr(urlNamespaces(i)), CStr(localRoots(i)))
        If Len(mappedPath) > 0 Then
            If Len(CStr(urlNamespaces(i))) > bestNamespaceLength Then
                bestNamespaceLength = Len(CStr(urlNamespaces(i)))
                bestMappedPath = mappedPath
            End If
        End If
    Next i

    MapWebPathToLocalFolder = bestMappedPath
End Function

' 单条映射：UrlNamespace 前缀匹配后，把剩余相对路径转成反斜杠形式。
' 两侧都先做 URL 解码，避免登记值用空格、Word 返回 %20 时匹配失败。
Private Function MapSingleWebPath(ByVal webPath As String, _
                                  ByVal urlNamespace As String, _
                                  ByVal localRoot As String) As String
    Dim cleanWebPath As String
    Dim cleanNamespace As String
    Dim cleanRoot As String
    Dim remainder As String

    cleanWebPath = UrlDecodePath(StripUrlQueryAndFragment(webPath))
    cleanNamespace = TrimTrailingSeparators(Trim$(UrlDecodePath(StripUrlQueryAndFragment(urlNamespace))), "/")
    cleanRoot = TrimTrailingSeparators(Trim$(localRoot), "\")

    If Len(cleanNamespace) = 0 Or Len(cleanRoot) = 0 Then Exit Function
    If Len(cleanWebPath) < Len(cleanNamespace) Then Exit Function
    If StrComp(Left$(cleanWebPath, Len(cleanNamespace)), cleanNamespace, vbTextCompare) <> 0 Then Exit Function

    remainder = Mid$(cleanWebPath, Len(cleanNamespace) + 1)
    ' 必须完整匹配到路径分隔处，避免 /Documents2 误匹配 /Documents
    If Len(remainder) > 0 Then
        If Left$(remainder, 1) <> "/" Then Exit Function
    End If

    remainder = Replace(remainder, "/", "\")
    remainder = TrimTrailingSeparators(TrimLeadingSeparators(remainder, "\"), "\")

    If Len(remainder) = 0 Then
        MapSingleWebPath = cleanRoot
    Else
        MapSingleWebPath = cleanRoot & "\" & remainder
    End If
End Function

' 截掉 URL 的查询串与片段
Private Function StripUrlQueryAndFragment(ByVal value As String) As String
    Dim cutPosition As Long

    cutPosition = InStr(1, value, "?")
    If cutPosition > 0 Then value = Left$(value, cutPosition - 1)

    cutPosition = InStr(1, value, "#")
    If cutPosition > 0 Then value = Left$(value, cutPosition - 1)

    StripUrlQueryAndFragment = value
End Function

' 去掉字符串末尾的指定分隔符
Private Function TrimTrailingSeparators(ByVal value As String, ByVal separator As String) As String
    Do While Len(value) > 0
        If Right$(value, 1) = separator Then
            value = Left$(value, Len(value) - 1)
        Else
            Exit Do
        End If
    Loop
    TrimTrailingSeparators = value
End Function

' 去掉字符串开头的指定分隔符
Private Function TrimLeadingSeparators(ByVal value As String, ByVal separator As String) As String
    Do While Len(value) > 0
        If Left$(value, 1) = separator Then
            value = Mid$(value, 2)
        Else
            Exit Do
        End If
    Loop
    TrimLeadingSeparators = value
End Function

' 解码 URL 中的 %XX 转义（按 UTF-8 字节解释）；非法转义保持原样
Private Function UrlDecodePath(ByVal value As String) As String
    Dim i As Long
    Dim decoded As String
    Dim pendingBytes() As Byte
    Dim pendingCount As Long
    Dim byteValue As Long

    If InStr(1, value, "%") = 0 Then
        UrlDecodePath = value
        Exit Function
    End If

    ReDim pendingBytes(0 To Len(value))
    i = 1

    Do While i <= Len(value)
        byteValue = -1
        If Mid$(value, i, 1) = "%" And i + 2 <= Len(value) Then
            byteValue = HexPairToByte(Mid$(value, i + 1, 2))
        End If

        If byteValue >= 0 Then
            pendingBytes(pendingCount) = CByte(byteValue)
            pendingCount = pendingCount + 1
            i = i + 3
        Else
            If pendingCount > 0 Then
                decoded = decoded & DecodeUtf8Bytes(pendingBytes, pendingCount)
                pendingCount = 0
            End If
            decoded = decoded & Mid$(value, i, 1)
            i = i + 1
        End If
    Loop

    If pendingCount > 0 Then decoded = decoded & DecodeUtf8Bytes(pendingBytes, pendingCount)
    UrlDecodePath = decoded
End Function

' 两个十六进制字符转字节；非法返回 -1
Private Function HexPairToByte(ByVal hexText As String) As Long
    Dim highValue As Long
    Dim lowValue As Long

    HexPairToByte = -1
    If Len(hexText) <> 2 Then Exit Function

    highValue = HexDigitValue(Left$(hexText, 1))
    lowValue = HexDigitValue(Mid$(hexText, 2, 1))
    If highValue < 0 Or lowValue < 0 Then Exit Function

    HexPairToByte = highValue * 16 + lowValue
End Function

' 单个十六进制字符转数值；非法返回 -1
Private Function HexDigitValue(ByVal digitText As String) As Long
    Dim codeValue As Long

    HexDigitValue = -1
    If Len(digitText) <> 1 Then Exit Function

    codeValue = AscW(digitText)
    Select Case codeValue
        Case 48 To 57      ' 0-9
            HexDigitValue = codeValue - 48
        Case 65 To 70      ' A-F
            HexDigitValue = codeValue - 55
        Case 97 To 102     ' a-f
            HexDigitValue = codeValue - 87
    End Select
End Function

' UTF-8 字节序列转 Unicode 字符串；非法序列用 U+FFFD 代替
Private Function DecodeUtf8Bytes(ByRef bytes() As Byte, ByVal byteCount As Long) As String
    Dim i As Long
    Dim firstByte As Long
    Dim secondByte As Long
    Dim thirdByte As Long
    Dim fourthByte As Long
    Dim codePoint As Long
    Dim decoded As String

    i = 0
    Do While i < byteCount
        firstByte = bytes(i)

        If firstByte < &H80 Then
            decoded = decoded & ChrW$(firstByte)
            i = i + 1
        ElseIf (firstByte And &HE0) = &HC0 And i + 1 < byteCount Then
            secondByte = bytes(i + 1)
            codePoint = ((firstByte And &H1F) * &H40) + (secondByte And &H3F)
            If codePoint < &H80 Then
                decoded = decoded & ChrW$(&HFFFD)
            Else
                decoded = decoded & ChrW$(codePoint)
            End If
            i = i + 2
        ElseIf (firstByte And &HF0) = &HE0 And i + 2 < byteCount Then
            secondByte = bytes(i + 1)
            thirdByte = bytes(i + 2)
            codePoint = ((firstByte And &HF) * &H1000) + ((secondByte And &H3F) * &H40) + (thirdByte And &H3F)
            If codePoint < &H800 Then
                decoded = decoded & ChrW$(&HFFFD)
            Else
                decoded = decoded & ChrW$(codePoint)
            End If
            i = i + 3
        ElseIf (firstByte And &HF8) = &HF0 And i + 3 < byteCount Then
            secondByte = bytes(i + 1)
            thirdByte = bytes(i + 2)
            fourthByte = bytes(i + 3)
            codePoint = ((firstByte And &H7) * &H40000) + ((secondByte And &H3F) * &H1000) + _
                        ((thirdByte And &H3F) * &H40) + (fourthByte And &H3F)
            If codePoint < &H10000 Or codePoint > &H10FFFF Then
                decoded = decoded & ChrW$(&HFFFD)
            Else
                codePoint = codePoint - &H10000
                decoded = decoded & ChrW$(&HD800 + (codePoint \ &H400)) & ChrW$(&HDC00 + (codePoint And &H3FF))
            End If
            i = i + 4
        Else
            decoded = decoded & ChrW$(&HFFFD)
            i = i + 1
        End If
    Loop

    DecodeUtf8Bytes = decoded
End Function

' 读取 OneDrive 同步引擎登记的全部 UrlNamespace -> MountPoint 映射。
' 64 位视图与 32 位视图都读：64 位 Office 读前者即可，32 位 Office 在
' 64 位 Windows 上会被重定向，两个视图各读一次最稳妥。
Private Sub LoadOneDriveSyncMappings(ByRef urlNamespaces As Collection, _
                                     ByRef localRoots As Collection)
    Dim accountKeys As Collection
    Dim i As Long

    ' 鏂扮増 OneDrive 鐧昏鍦?SyncEngines\Providers\OneDrive 涓?    CollectSyncMappingsFromRegistry ONEDRIVE_SYNC_REG_KEY, "/reg:64", urlNamespaces, localRoots
    CollectSyncMappingsFromRegistry ONEDRIVE_SYNC_REG_KEY, "/reg:32", urlNamespaces, localRoots

    ' 閮ㄥ垎鐗堟湰/绉熸埛鍙櫥璁板湪 Accounts\<Personal|BusinessN> 涓?    Set accountKeys = ListAccountSubKeys()
    If Not accountKeys Is Nothing Then
        For i = 1 To accountKeys.Count
            CollectSyncMappingsFromRegistry CStr(accountKeys(i)), "/reg:64", urlNamespaces, localRoots
            CollectSyncMappingsFromRegistry CStr(accountKeys(i)), "/reg:32", urlNamespaces, localRoots
        Next i
    End If
End Sub

' 鍒楀嚭 Accounts 涓嬬殑璐﹀彿瀛愰敭锛圥ersonal銆丅usiness1銆丅usiness2...锛?Private Function ListAccountSubKeys() As Collection
    Dim lines As Collection
    Dim keys As Collection
    Dim i As Long
    Dim lineText As String
    Dim prefixLength As Long

    Set keys = New Collection
    Set ListAccountSubKeys = keys

    Set lines = ReadRegistryQueryLines(ONEDRIVE_ACCOUNTS_REG_KEY, "/reg:64")
    If lines Is Nothing Then Exit Function

    prefixLength = Len("HKEY_CURRENT_USER\" & Replace(ONEDRIVE_ACCOUNTS_REG_KEY, "HKCU\", ""))
    For i = 1 To lines.Count
        lineText = CStr(lines(i))
        If IsRegistryKeyLine(lineText) Then
            lineText = RTrim$(lineText)
            ' 鍙敹 Accounts 鐨勭洿鎺ュ瓙閿紝璺宠繃 Accounts 鑷韩
            If Len(lineText) > prefixLength + 1 Then
                If StrComp(Left$(lineText, prefixLength), "HKEY_CURRENT_USER\" & Replace(ONEDRIVE_ACCOUNTS_REG_KEY, "HKCU\", ""), vbTextCompare) = 0 Then
                    keys.Add "HKCU\" & Mid$(lineText, prefixLength + 2)
                End If
            End If
        End If
    Next i
End Function

' 执行一次 reg query /s 并把每个子键的 UrlNamespace、MountPoint 配对收集
Private Sub CollectSyncMappingsFromRegistry(ByVal rootKey As String, _
                                            ByVal regView As String, _
                                            ByRef urlNamespaces As Collection, _
                                            ByRef localRoots As Collection)
    Dim lines As Collection
    Dim i As Long
    Dim lineText As String
    Dim namespaceUrl As String
    Dim mountPoint As String
    Dim valueName As String
    Dim valueData As String

    Set lines = ReadRegistryQueryLines(rootKey, regView)
    If lines Is Nothing Then Exit Sub

    For i = 1 To lines.Count
        lineText = CStr(lines(i))

        If IsRegistryKeyLine(lineText) Then
            ' 进入下一个子键：先把上一个子键的配对收尾
            If Len(namespaceUrl) > 0 And Len(mountPoint) > 0 Then
                AppendSyncMapping urlNamespaces, localRoots, namespaceUrl, mountPoint
            End If
            namespaceUrl = ""
            mountPoint = ""
        ElseIf ParseRegistryValueLine(lineText, valueName, valueData) Then
            If StrComp(valueName, "UrlNamespace", vbTextCompare) = 0 Then
                namespaceUrl = valueData
            ElseIf StrComp(valueName, "MountPoint", vbTextCompare) = 0 Then
                mountPoint = valueData
            End If
        End If
    Next i

    If Len(namespaceUrl) > 0 And Len(mountPoint) > 0 Then
        AppendSyncMapping urlNamespaces, localRoots, namespaceUrl, mountPoint
    End If
End Sub

' 用 reg.exe 执行 reg query /s，返回去 CR 的文本行集合；读取失败返回空集合。
' 输出先落到临时文件再读回：WScript.Shell.Exec 在 Word 的隐藏自动化实例里
' 读 StdOut 会永久阻塞（实测），Run + 重定向没有这个问题。
Private Function ReadRegistryQueryLines(ByVal rootKey As String, _
                                        ByVal regView As String) As Collection
    Dim shell As Object
    Dim fso As Object
    Dim outputFile As Object
    Dim outputText As String
    Dim tempPath As String
    Dim lines As Collection

    Set lines = New Collection
    Set ReadRegistryQueryLines = lines

    On Error GoTo CleanUp

    Set fso = CreateObject("Scripting.FileSystemObject")
    tempPath = fso.GetSpecialFolder(2).Path & "\" & fso.GetTempName()

    Set shell = CreateObject("WScript.Shell")
    shell.Run "cmd /c reg query """ & rootKey & """ /s " & regView & " > """ & tempPath & """ 2>nul", 0, True

    ' 键不存在时 reg 只写 stderr，重定向得到 0 字节文件；对空文件 ReadAll
    ' 会报"输入超出文件尾"，所以先看大小。
    If fso.FileExists(tempPath) Then
        If fso.GetFile(tempPath).Size > 0 Then
            Set outputFile = fso.OpenTextFile(tempPath, 1)
            outputText = outputFile.ReadAll
            outputFile.Close
            Set outputFile = Nothing
        End If
    End If

    If Len(outputText) > 0 Then
        CopyLinesNoCarriageReturn lines, outputText
    End If

CleanUp:
    If Err.Number <> 0 Then Err.Clear
    On Error Resume Next
    If Len(tempPath) > 0 Then
        If fso.FileExists(tempPath) Then fso.DeleteFile tempPath, True
    End If
    On Error GoTo 0
    Set shell = Nothing
    Set fso = Nothing
End Function

' 把文本按行拆开，去掉行尾 CR（后续比较不再受 CRLF/LF 影响）
Private Sub CopyLinesNoCarriageReturn(ByRef lines As Collection, ByVal outputText As String)
    Dim rawLines As Variant
    Dim i As Long
    Dim lineText As String

    rawLines = Split(outputText, vbLf)
    For i = LBound(rawLines) To UBound(rawLines)
        lineText = CStr(rawLines(i))
        If Len(lineText) > 0 Then
            If Right$(lineText, 1) = vbCr Then lineText = Left$(lineText, Len(lineText) - 1)
        End If
        lines.Add lineText
    Next i
End Sub

' reg query 输出的键行以 HKEY_ 开头（根键回显与各子键都是）
Private Function IsRegistryKeyLine(ByVal lineText As String) As Boolean
    IsRegistryKeyLine = (StrComp(Left$(LTrim$(lineText), 5), "HKEY_", vbTextCompare) = 0)
End Function

' 解析 reg query 的值行："    UrlNamespace    REG_SZ    https://..."
Private Function ParseRegistryValueLine(ByVal lineText As String, _
                                        ByRef valueName As String, _
                                        ByRef valueData As String) As Boolean
    Dim trimmed As String
    Dim typePosition As Long

    ParseRegistryValueLine = False
    valueName = ""
    valueData = ""

    trimmed = LTrim$(lineText)
    typePosition = InStr(1, trimmed, "REG_SZ", vbTextCompare)
    If typePosition <= 1 Then Exit Function

    valueName = RTrim$(Left$(trimmed, typePosition - 1))
    If Len(valueName) = 0 Then Exit Function

    valueData = LTrim$(Mid$(trimmed, typePosition + Len("REG_SZ")))
    valueData = TrimTrailingSeparators(valueData, " ")
    ParseRegistryValueLine = True
End Function

' 追加一条映射（同一 UrlNamespace 只保留首次出现）
Private Sub AppendSyncMapping(ByRef urlNamespaces As Collection, _
                              ByRef localRoots As Collection, _
                              ByVal namespaceUrl As String, _
                              ByVal mountPoint As String)
    Dim i As Long

    For i = 1 To urlNamespaces.Count
        If StrComp(CStr(urlNamespaces(i)), namespaceUrl, vbTextCompare) = 0 Then Exit Sub
    Next i

    urlNamespaces.Add namespaceUrl
    localRoots.Add mountPoint
End Sub

' 判断文件夹是否存在
Private Function FolderExistsSync(ByVal folderPath As String) As Boolean
    Dim fso As Object

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then FolderExistsSync = fso.FolderExists(folderPath)
    Set fso = Nothing
    On Error GoTo 0
End Function

' 鍒ゆ柇鏂囦欢鏄惁瀛樺湪
Private Function FileExistsSync(ByVal filePath As String) As Boolean
    Dim fso As Object

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then FileExistsSync = fso.FileExists(filePath)
    Set fso = Nothing
    On Error GoTo 0
End Function
