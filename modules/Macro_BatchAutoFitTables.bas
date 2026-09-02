Attribute VB_Name = "Macro_BatchAutoFitTables"
Option Explicit

' ==========================================
' 一键表格自动调整：全文档表格设为“根据窗口自动调整”
' ==========================================

Private Const DIALOG_TITLE As String = "一键表格自动调整"

Sub BatchAutoFitTablesToWindow()
    Dim objTable As Table
    Dim tableCount As Long

    If Documents.count = 0 Then
        MsgBox "没有打开的文档！", vbExclamation, DIALOG_TITLE
        Exit Sub
    End If

    BeginBatchUI
    On Error GoTo ErrH

    tableCount = 0

    If ActiveDocument.Tables.count > 0 Then
        For Each objTable In ActiveDocument.Tables
            objTable.AutoFitBehavior wdAutoFitWindow
            tableCount = tableCount + 1
        Next objTable
    Else
        EndBatchUI
        MsgBox "当前文档中没有发现表格。", vbExclamation, DIALOG_TITLE
        Exit Sub
    End If

    EndBatchUI
    MsgBox "处理完成，已成功调整 " & tableCount & " 个表格。", vbInformation, DIALOG_TITLE
    Exit Sub

ErrH:
    EndBatchUI
    MsgBox "表格调整失败：" & Err.Description, vbCritical, DIALOG_TITLE
End Sub