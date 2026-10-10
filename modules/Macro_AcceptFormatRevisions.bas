Attribute VB_Name = "Macro_AcceptFormatRevisions"
Option Explicit

' =============================================
' 模块名：Macro_AcceptFormatRevisions
' 功能：只接受"格式类修订"（字符格式、段落格式），
'       插入、删除、替换、移动等内容类修订保持未接受状态。
' 说明：Document.Revisions 已覆盖正文、页眉页脚、脚注尾注、
'       文本框等全部故事区域，无需逐 StoryRange 遍历。
' =============================================

Private Const DIALOG_TITLE As String = "接受格式修订"

' 最大轮次：每轮只枚举一次修订集合，总量有上界；
' 多轮仅用于兜底个别修订无法一次消化的情况
Private Const MAX_SCAN_ROUNDS As Long = 3

' Ribbon 回调：接受全文格式类修订
Public Sub AcceptFormatRevisions(control As IRibbonControl)
    Dim doc As Document
    Dim accepted As Long
    Dim remaining As Long
    Dim formatLeft As Long
    Dim contentLeft As Long
    Dim hasError As Boolean
    Dim errDesc As String
    Dim msgText As String

    If Application.Documents.Count = 0 Then
        MsgBox "当前没有打开的文档。", vbExclamation, DIALOG_TITLE
        Exit Sub
    End If

    Set doc = ActiveDocument

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

    ' 剩余修订分类：格式类（未能接受）与内容类
    contentLeft = remaining - formatLeft
    If contentLeft < 0 Then contentLeft = 0

    ' 组装结果提示：明确区分"已接受的格式类修订"与"未接受的内容类修订"
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
                  "文档可能处于保护状态或修订被锁定，请解除限制后重试。"
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

' 宏列表入口：Application.Run 不支持带参数的过程，包装为无参入口
Public Sub Wrapper_AcceptFormatRevisions()
    AcceptFormatRevisions Nothing
End Sub

' 接受文档中的格式类修订，返回接受数量
' 策略：每轮先枚举收集格式类修订对象，再逐个接受，最多 MAX_SCAN_ROUNDS 轮。
' 每轮仅枚举一次修订集合，总操作量有明确上界，
' 不会因个别修订无法消化而在集合上反复空转。
Private Function AcceptFormatRevisionsInDocument(ByVal doc As Document) As Long
    Dim rev As Revision
    Dim targets As Collection
    Dim accepted As Long
    Dim scanRound As Long
    Dim acceptedInRound As Long
    Dim i As Long

    On Error Resume Next
    For scanRound = 1 To MAX_SCAN_ROUNDS
        ' 收集本轮检测到的格式类修订；
        ' 白名单：wdRevisionProperty=字符格式，wdRevisionParagraphProperty=段落格式，
        ' 插入、删除、替换、移动等内容类修订一律不收集
        Set targets = New Collection
        For Each rev In doc.Revisions
            If rev.Type = wdRevisionProperty Or rev.Type = wdRevisionParagraphProperty Then
                targets.Add rev
            End If
        Next rev

        ' 逐个接受；对象引用在集合变化后失效的情况由 Err 保护，跳过即可
        acceptedInRound = 0
        For i = 1 To targets.count
            Err.Clear
            targets(i).Accept
            If Err.Number = 0 Then
                accepted = accepted + 1
                acceptedInRound = acceptedInRound + 1
            End If
            Err.Clear
        Next i
        Set targets = Nothing

        ' 本轮没有接受任何修订（无格式类修订，或全部无法接受）时结束
        If acceptedInRound = 0 Then Exit For

        ' 已无格式类修订残留时结束
        If CountFormatRevisions(doc) = 0 Then Exit For
    Next scanRound
    On Error GoTo 0

    AcceptFormatRevisionsInDocument = accepted
End Function

' 统计文档中剩余的格式类修订数量（用于识别受保护/被锁定导致的接受失败）
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
