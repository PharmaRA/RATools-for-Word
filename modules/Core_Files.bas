Attribute VB_Name = "Core_Files"
Option Explicit

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