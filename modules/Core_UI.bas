Attribute VB_Name = "Core_UI"
Option Explicit

' =============================================
' 公共层：批处理交互（模式选择/文件选择/进度）
' =============================================

' 三模式取值约定
Public Const BATCH_MODE_CURRENT As String = "1"
Public Const BATCH_MODE_FILES As String = "2"
Public Const BATCH_MODE_FOLDER As String = "3"

' 统一的三模式选择对话框；返回 "1"/"2"/"3"，用户取消返回空串。
' 使用 StrPtr 区分「点了取消」与「清空后确定」，两者都视为取消。
Public Function ChooseBatchMode(ByVal dialogTitle As String) As String
    Dim answer As String

    answer = InputBox("请输入处理模式序号：" & vbCrLf & vbCrLf & _
                      "1 - 仅当前文档：处理当前打开的文档" & vbCrLf & _
                      "2 - 多文件模式：选择单个或多个文件" & vbCrLf & _
                      "3 - 文件夹模式：递归处理所选文件夹内全部 Word 文档", _
                      dialogTitle, BATCH_MODE_CURRENT)

    If StrPtr(answer) = 0 Then
        ChooseBatchMode = ""
    Else
        ChooseBatchMode = Trim$(answer)
    End If
End Function

' Word 文档多选对话框；返回所选路径集合，用户取消返回 Nothing
Public Function PickWordFiles(ByVal dialogTitle As String) As Collection
    Dim result As Collection
    Dim i As Long

    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = dialogTitle
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Word 文档", "*.doc; *.docx; *.docm", 1
        If .Show = -1 Then
            Set result = New Collection
            For i = 1 To .SelectedItems.count
                result.Add .SelectedItems(i)
            Next i
        End If
    End With

    Set PickWordFiles = result
End Function

' 文件夹选择对话框；用户取消返回空串
Public Function PickFolder(ByVal dialogTitle As String) As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = dialogTitle
        If .Show = -1 Then PickFolder = .SelectedItems(1)
    End With
End Function

' 状态栏进度显示
Public Sub ShowBatchProgress(ByVal currentIndex As Long, _
                             ByVal totalCount As Long, _
                             ByVal itemLabel As String)
    Application.StatusBar = "正在处理 [" & currentIndex & "/" & totalCount & "] " & itemLabel
End Sub