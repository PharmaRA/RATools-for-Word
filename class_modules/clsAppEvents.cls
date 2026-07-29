VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 定义一个带事件的 Application 对象
Public WithEvents App As Word.Application
Attribute App.VB_VarHelpID = -1

' 初始化类：绑定 Word 应用程序对象
Private Sub Class_Initialize()
    Set App = Word.Application
End Sub

' ==========================================
' 场景1：新建文档时 -> 激活选项卡
' ==========================================
Private Sub App_NewDocument(ByVal Doc As Document)
    ' 新建文档时，调用 mRibbon 中的方法激活选项卡
    mRibbon.ActivateRAToolsTab
End Sub

' ==========================================
' 场景2：打开文档时 -> 激活选项卡
' ==========================================
Private Sub App_DocumentOpen(ByVal Doc As Document)
    ' 打开现有文档时，同样激活选项卡
    mRibbon.ActivateRAToolsTab
End Sub

' ==========================================
' 案例3：切换/新建/打开/关闭文档后 -> 抬升悬浮窗
' 悬浮窗的 owner 被置为 0（见 Mod_Core_Window），因此不再由 Word
' 自动置于文档窗口之上，需要在文档窗口变化后显式抬回前面。
' ==========================================
Private Sub App_DocumentChange()
    Mod_QuickToolbar.RaiseQuickToolbarIfVisible
End Sub

' ==========================================
' 案例4：激活某个文档窗口后 -> 抬升悬浮窗
' ==========================================
Private Sub App_WindowActivate(ByVal Doc As Document, ByVal Wn As Window)
    Mod_QuickToolbar.RaiseQuickToolbarIfVisible
End Sub

' ==========================================
' 案例5：Word 退出前 -> 卸载悬浮窗并存盘位置
' owner 为 0 后窗体不会被文档窗口销毁带走，必须显式卸载。
' ==========================================
Private Sub App_Quit()
    Mod_QuickToolbar.UnloadQuickToolbarOnQuit
End Sub
