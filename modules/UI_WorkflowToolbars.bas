Attribute VB_Name = "UI_WorkflowToolbars"
Option Explicit

Private mActionRunning As Boolean

Private Function LoadedWorkflowForm(ByVal formName As String) As Object
    Dim item As Object
    For Each item In VBA.UserForms
        If TypeName(item) = formName Then
            Set LoadedWorkflowForm = item
            Exit Function
        End If
    Next item
End Function

Public Sub ToggleWritingToolbar(ByVal control As IRibbonControl)
    ToggleWorkflowToolbar "Writing"
End Sub

Public Sub ToggleWorkflowToolbar(ByVal kind As String)
    Dim toolbar As Object
    On Error GoTo Failed
    Set toolbar = LoadedWorkflowForm("frm" & kind & "Toolbar")
    If Not toolbar Is Nothing Then
        If toolbar.Visible Then
            toolbar.SaveToolbarPosition
            toolbar.Hide
            Exit Sub
        End If
    End If
    ShowWorkflowToolbar kind
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "RATools"
End Sub

Public Sub ShowWorkflowToolbar(ByVal kind As String)
    Dim toolbar As Object
    On Error GoTo Failed
    UI_Ribbon.InitEvents
    Set toolbar = LoadedWorkflowForm("frm" & kind & "Toolbar")
    If toolbar Is Nothing Then
        Set toolbar = VBA.UserForms.Add("frm" & kind & "Toolbar")
    End If
    toolbar.Show vbModeless
    DetachUserFormWindowFromOwner toolbar.Caption
    RefreshWorkflowToolbarsIfVisible
    ReturnFocusToDocumentWindow
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "RATools"
End Sub

Public Sub RefreshWorkflowToolbarsIfVisible()
    Dim toolbar As Object
    On Error Resume Next
    For Each toolbar In VBA.UserForms
        If TypeName(toolbar) = "frmWritingToolbar" Then
            If toolbar.Visible Then
                toolbar.RefreshDocumentState mActionRunning
                KeepUserFormWindowOnTop toolbar.Caption
            End If
        End If
    Next toolbar
End Sub

Public Sub UnloadWorkflowToolbarsOnQuit()
    Dim index As Long
    Dim toolbar As Object
    On Error Resume Next
    For index = VBA.UserForms.Count - 1 To 0 Step -1
        Set toolbar = VBA.UserForms(index)
        If TypeName(toolbar) = "frmWritingToolbar" Then
            toolbar.SaveToolbarPosition
            Unload toolbar
        End If
    Next index
End Sub

Public Sub RunWorkflowAction(ByVal actionKey As String)
    Dim target As Document
    Dim keepFocus As Boolean
    Dim failureText As String
    If mActionRunning Then Exit Sub
    On Error GoTo Failed
    If Documents.Count = 0 Then
        RefreshWorkflowToolbarsIfVisible
        Exit Sub
    End If
    Set target = ActiveDocument
    mActionRunning = True
    RefreshWorkflowToolbarsIfVisible
    If Left$(actionKey, 6) = "style:" Then
        ApplyRAToolsStyle Mid$(actionKey, 7)
    ElseIf Left$(actionKey, 4) = "mso:" Then
        ExecuteQuickToolbarMso Mid$(actionKey, 5)
        Select Case Mid$(actionKey, 5)
            Case "ParagraphDialog", "PageSetupDialog", "SectionBreakInsert": keepFocus = True
        End Select
    Else
        Select Case actionKey
            Case "page_break_before": TogglePageBreakBefore Nothing
            Case "autofit_table": AutoFitTableWindow Nothing
            Case "text_blue": SetTextBlue Nothing
            Case "style_pane": ShowStylePane Nothing: keepFocus = True
            Case "normalize_terms": NormalizeScientificTerms
            Case "extract_abbreviations": ExtractAbbreviations
            Case "autofit_all_tables": BatchAutoFitTablesToWindow
            Case "link_sections": LinkToThePreviousSection
            Case "batch_margins": BatchSetMargins
            Case Else: Err.Raise vbObjectError + 2100, , "Unknown workflow action: " & actionKey
        End Select
    End If
Finished:
    mActionRunning = False
    Set target = Nothing
    RefreshWorkflowToolbarsIfVisible
    If Not keepFocus Then ReturnFocusToDocumentWindow
    If Len(failureText) > 0 Then MsgBox failureText, vbExclamation, "RATools"
    Exit Sub
Failed:
    failureText = Err.Description
    Resume Finished
End Sub
