Attribute VB_Name = "UI_RibbonCommands"
Option Explicit

' =============================================
' 文档快捷命令（Ribbon onAction 回调集合）
' 从 mRibbon 拆出（v0.8 重构阶段5）。所有 Public Sub 名称不可改。
' =============================================

'=====================  一 键 大 写  =====================
Public Sub btnCap_Click(ByVal control As IRibbonControl)
    On Error Resume Next
    Selection.Range.Case = wdUpperCase
End Sub


'================  设置文字为蓝色  ================
Public Sub SetTextBlue(control As IRibbonControl)
    On Error Resume Next
    Selection.Font.Color = wdColorBlue
End Sub


'================  域格式保护  ================
'为选区内 REF/PAGEREF 加 \* MERGEFORMAT,保护域格式
'包含智能判断(全文/选区) + 结果弹窗
Public Sub ProtectFieldFormat(control As IRibbonControl)
    Dim fld As field, rng As Range
    Dim targetFields As Fields ' 目标域集合
    Dim msgTip As String       ' 提示信息

    ' 判断：如果是光标插入点(wdSelectionIP)则处理全文，否则处理选区
    If Selection.Type = wdSelectionIP Then
        Set targetFields = ActiveDocument.Fields
        msgTip = "未选中文字，已对【全文】域代码进行格式保护。"
    Else
        Set targetFields = Selection.Fields
        msgTip = "已对【选中区域】域代码进行格式保护。"
    End If

    ' 遍历处理
    For Each fld In targetFields
        If fld.Type = wdFieldRef Or fld.Type = wdFieldPageRef Then
            Set rng = fld.Code
            If InStr(1, rng.Text, "mergeformat", vbTextCompare) = 0 Then
                rng.Text = rng.Text & " \* MERGEFORMAT "
                fld.Update
            End If
        End If
    Next fld

    ' 操作完成后弹出提示
    MsgBox msgTip, vbInformation, "操作完成"
End Sub


'================  打开文件所在文件夹  ================
Public Sub OpenDocumentFolder(control As IRibbonControl)
        
    ' 1. 声明一个变量，用来临时存储文档的路径
    Dim docPath As String
    
    ' 2. 获取当前活动文档的路径
    docPath = ActiveDocument.Path
    
    ' 3. 检查文档是否已经保存过（如果没有保存过，路径会是空的）
    If docPath = "" Then
        ' 弹出一个消息框提醒用户
        MsgBox "这个文档还没有保存过哦！请先保存文档，然后再尝试打开所在文件夹。", vbExclamation, "温馨提示"
        ' 停止运行后面的代码
        Exit Sub
    End If
    
    ' 4. 如果路径存在，使用 Windows 资源管理器打开该路径
    ' explorer.exe 是 Windows 自带的文件夹管理程序
    Shell "explorer.exe """ & docPath & """", vbNormalFocus
    
End Sub


'================  段前分页切换  ================
' 说明：切换选中段落的“段前分页”属性 (PageBreakBefore)
' 逻辑：如果是混合状态或关闭状态 -> 设为开启；如果是纯开启状态 -> 设为关闭
Public Sub TogglePageBreakBefore(control As IRibbonControl)
    On Error Resume Next
    Dim currentStatus As Long
    
    ' 获取当前选中段落的段前分页状态
    ' 0 = False (关), -1 = True (开), 9999999 = wdUndefined (混合)
    currentStatus = Selection.ParagraphFormat.PageBreakBefore
    
    ' 如果全关(0)，则开启(-1)
    ' 如果混合(wdUndefined)，也统一开启(-1)
    ' 如果全开(-1)，则关闭(0)
    If currentStatus = -1 Then
        Selection.ParagraphFormat.PageBreakBefore = 0 ' 关闭
    Else
        Selection.ParagraphFormat.PageBreakBefore = -1 ' 开启
    End If
End Sub


'================  表格功能：根据窗口自动调整  ================
Public Sub AutoFitTableWindow(control As IRibbonControl)
    On Error Resume Next
    
    ' 检查光标是否在表格内
    If Selection.Information(wdWithInTable) Then
        ' 将当前所在的表格设置为：根据窗口自动调整
        Selection.Tables(1).AutoFitBehavior wdAutoFitWindow
    Else
        MsgBox "请先将光标定位在表格内部。", vbExclamation, "提示"
    End If
End Sub


'================  下拉选择对齐方式  ================
'================  下拉菜单：左对齐  ================
Public Sub AlignLeft_Click(control As IRibbonControl)
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
End Sub


'================  顶部大按钮：直接设为居中  ================
Public Sub AlignCenter_Click(control As IRibbonControl)
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
End Sub


'================  下拉菜单：右对齐  ================
Public Sub AlignRight_Click(control As IRibbonControl)
    Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
End Sub


'================  下拉菜单：两端对齐  ================
Public Sub AlignJustify_Click(control As IRibbonControl)
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
End Sub



' 显示/隐藏样式管理窗格
Public Sub ShowStylePane(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    ' 尝试使用内置命令打开样式窗格
    Application.CommandBars.ExecuteMso "StylesPane"
    
    Exit Sub
    
ErrorHandler:
    ' 如果内置命令失败，使用快捷键
    SendKeys "%^{+}s", True
End Sub


'=====================  Wrapper 包装器  =====================
' 解释：因为很多宏是 Ribbon 回调 (带 control 参数)，通常是通过按钮直接调用功能
' Application.Run 无法自动提供 control 参数，直接运行会报错。
' 所以我们需要一些不带参数的“外壳”过程。

Public Sub Wrapper_ProtectFieldFormat()
    ' 调用原有的逻辑
    ' 注意：因为原 Sub 需要 control 参数，我们传 Nothing 进去
    ' 只要原 Sub 内部没用到 control.ID 或 control.Tag，这样写就是安全的
    ProtectFieldFormat Nothing
End Sub

Public Sub Wrapper_AddMissingCaptionLabels(control As IRibbonControl)
    AddMissingCaptionLabels
End Sub



Public Sub Wrapper_RenameCurrentDocument(control As IRibbonControl)
    RenameCurrentDocument
End Sub
