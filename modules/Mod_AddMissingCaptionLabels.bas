Attribute VB_Name = "Mod_AddMissingCaptionLabels"
Option Explicit

Public Sub AddMissingCaptionLabels()
    Dim fld As field
    Dim fieldCode As String
    Dim labelName As String
    Dim existingLabel As CaptionLabel
    Dim labelExists As Boolean
    Dim parts() As String

    On Error GoTo CleanFail

    Application.ScreenUpdating = False

    For Each fld In ActiveDocument.Fields
        If fld.Type = wdFieldSequence Then
            fieldCode = Trim$(fld.Code.Text)

            If InStr(1, fieldCode, "SEQ ", vbTextCompare) > 0 Then
                parts = Split(fieldCode, " ")

                If UBound(parts) >= 1 Then
                    labelName = parts(1)
                    labelName = Replace(labelName, "\", "")
                    labelName = Replace(labelName, "*", "")
                    labelName = Trim$(labelName)

                    If labelName <> "" Then
                        labelExists = False

                        For Each existingLabel In CaptionLabels
                            If StrComp(existingLabel.Name, labelName, vbTextCompare) = 0 Then
                                labelExists = True
                                Exit For
                            End If
                        Next existingLabel

                        If Not labelExists Then
                            CaptionLabels.Add Name:=labelName
                        End If
                    End If
                End If
            End If
        End If
    Next fld

    Application.ScreenUpdating = True
    MsgBox "已检查并恢复所有缺失的题注标签。现在您可以去“交叉引用”中查看了。", vbInformation, "完成"
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "恢复题注标签时出错：" & Err.Description, vbExclamation, "失败"
End Sub
