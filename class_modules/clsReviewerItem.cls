VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsReviewerItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================
' 审阅者筛选面板里的一个动态复选框。
' MSForms 的动态控件事件只能由带 WithEvents 的类接收，因此每个复选框
' 配一个本类实例，把 Click 转回窗体。
'
' 关键点：代码给 CheckBox.Value 赋值同样会触发 Click，而且事件可能延迟到
' 下一次消息循环才派发（实测）。若不加区分，回填勾选状态会反过来把筛选
' 结果改掉。这里记住"最后一次由代码写入的值"，值与它相同的 Click 视为回显
' 直接丢弃；用户点击会先翻转控件值，必然与写入值不同，能正常转发。
' =============================================

Public WithEvents chk As MSForms.CheckBox
Attribute chk.VB_VarHelpID = -1

' 窗体（FRM 的默认实例），只调用它公开的过程
Public Owner As Object

' 作者下标；-1 表示「全部审阅者」这一项
Public AuthorIndex As Long

Private m_WrittenValue As Boolean
Private m_HasWritten As Boolean

' 由窗体调用：写入勾选状态，并记录为"代码写入值"
Public Sub SyncValue(ByVal newValue As Boolean)
    m_WrittenValue = newValue
    m_HasWritten = True

    On Error Resume Next
    If chk.Value <> newValue Then chk.Value = newValue
    On Error GoTo 0
End Sub

Private Sub chk_Click()
    Dim currentValue As Boolean

    If Owner Is Nothing Then Exit Sub

    currentValue = (chk.Value = True)

    ' 与最后一次代码写入一致 = 程序化回填引起的回显（可能延迟派发），忽略
    If m_HasWritten Then
        If currentValue = m_WrittenValue Then Exit Sub
    End If

    m_WrittenValue = currentValue
    m_HasWritten = True
    Owner.OnItemClicked AuthorIndex, currentValue
End Sub
