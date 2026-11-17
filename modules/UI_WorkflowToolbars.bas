Attribute VB_Name = "UI_WorkflowToolbars"
Option Explicit

Private mActionRunning As Boolean
Private mRefreshing As Boolean

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

Public Sub ToggleFinalToolbar(ByVal control As IRibbonControl)
    ToggleWorkflowToolbar "Final"
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
    If kind = "Final" Then toolbar.ReadDocumentSnapshot
    DetachUserFormWindowFromOwner toolbar.Caption
    RefreshWorkflowToolbarsIfVisible
    ReturnFocusToDocumentWindow
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "RATools"
End Sub

Public Sub RefreshWorkflowToolbarsIfVisible(Optional ByVal mayHaveEdited As Boolean = False)
    Dim toolbar As Object
    If mRefreshing Then Exit Sub
    mRefreshing = True
    On Error Resume Next
    For Each toolbar In VBA.UserForms
        If TypeName(toolbar) = "frmWritingToolbar" Or TypeName(toolbar) = "frmFinalToolbar" Then
            If toolbar.Visible Then
                If TypeName(toolbar) = "frmFinalToolbar" Then
                    toolbar.RefreshDocumentState mActionRunning, mayHaveEdited
                Else
                    toolbar.RefreshDocumentState mActionRunning
                End If
                KeepUserFormWindowOnTop toolbar.Caption
            End If
        End If
    Next toolbar
    mRefreshing = False
End Sub

Public Sub UnloadWorkflowToolbarsOnQuit()
    Dim index As Long
    Dim toolbar As Object
    On Error Resume Next
    For index = VBA.UserForms.Count - 1 To 0 Step -1
        Set toolbar = VBA.UserForms(index)
        If TypeName(toolbar) = "frmWritingToolbar" Or TypeName(toolbar) = "frmFinalToolbar" Then
            toolbar.SaveToolbarPosition
            Unload toolbar
        End If
    Next index
End Sub

' Snapshot counts follow the existing clean detector's story traversal.
Public Function ReadWorkflowReviewSnapshot(ByVal doc As Document) As Variant
    Dim rootStory As Range, story As Range
    Dim revisionCount As Long, commentCount As Long
    commentCount = doc.Comments.Count
    For Each rootStory In doc.StoryRanges
        Set story = rootStory
        Do While Not story Is Nothing
            revisionCount = revisionCount + story.Revisions.Count
            Set story = story.NextStoryRange
        Loop
    Next rootStory
    ReadWorkflowReviewSnapshot = Array(revisionCount, commentCount, CBool(doc.TrackRevisions))
End Function

Public Sub RunWorkflowAction(ByVal actionKey As String)
    Dim target As Document
    Dim keepFocus As Boolean
    Dim failureText As String
    Dim invalidateSnapshot As Boolean
    Dim toolbar As Object
    If mActionRunning Then Exit Sub
    On Error GoTo Failed
    If Documents.Count = 0 Then
        RefreshWorkflowToolbarsIfVisible
        Exit Sub
    End If
    Set target = ActiveDocument
    invalidateSnapshot = True
    mActionRunning = True
    RefreshWorkflowToolbarsIfVisible
    Set target = ActiveDocument
    If Left$(actionKey, 6) = "style:" Then
        ApplyRAToolsStyle Mid$(actionKey, 7)
    ElseIf Left$(actionKey, 4) = "mso:" Then
        ExecuteQuickToolbarMso Mid$(actionKey, 5)
        Select Case Mid$(actionKey, 5)
            Case "ParagraphDialog", "PageSetupDialog", "SectionBreakInsert", "TableOfContentsUpdate", "TableOfFiguresUpdate": keepFocus = True
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
            Case "refresh_final_state"
                invalidateSnapshot = False
                Set toolbar = LoadedWorkflowForm("frmFinalToolbar")
                If Not toolbar Is Nothing Then toolbar.ReadDocumentSnapshot
            Case "detect_clean_versions": BatchDetectCleanVersions: invalidateSnapshot = False
            Case "detect_highlights": BatchDetectHighlights: invalidateSnapshot = False
            Case "protect_field_format": ProtectFieldFormat Nothing
            Case "accept_revisions_comments": BatchAcceptAndClean
            Case "heading_numbers_text": ConvertHeadingNumbers
            Case "save_current": target.Save
            Case "export_pdf": BatchConvertWordToPDF
            Case "rename_current": Wrapper_RenameCurrentDocument Nothing
            Case "open_folder": OpenDocumentFolder Nothing: keepFocus = True: invalidateSnapshot = False
            Case Else: Err.Raise vbObjectError + 2100, , "Unknown workflow action: " & actionKey
        End Select
    End If
Finished:
    mActionRunning = False
    Set target = Nothing
    RefreshWorkflowToolbarsIfVisible invalidateSnapshot
    If Not keepFocus Then ReturnFocusToDocumentWindow
    If Len(failureText) > 0 Then MsgBox failureText, vbExclamation, "RATools"
    Exit Sub
Failed:
    failureText = Err.Description
    Resume Finished
End Sub
