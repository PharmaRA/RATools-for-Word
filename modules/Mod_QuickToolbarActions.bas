Attribute VB_Name = "Mod_QuickToolbarActions"
Option Explicit

Private mQuickToolbarMsoTestMode As Boolean
Private mLastQuickToolbarMso As String

Public Function TryRunExpandedQuickToolbarAction(ByVal actionKey As String) As Boolean
    Select Case LCase$(Trim$(actionKey))
        Case "style_heading4"
            ApplyRAToolsStyle NumberedHeadingStyle(4)
        Case "style_unnumbered_heading1"
            ApplyRAToolsStyle UnnumberedHeadingStyle(1)
        Case "style_unnumbered_heading2"
            ApplyRAToolsStyle UnnumberedHeadingStyle(2)
        Case "style_unnumbered_heading3"
            ApplyRAToolsStyle UnnumberedHeadingStyle(3)
        Case "style_unnumbered_heading4"
            ApplyRAToolsStyle UnnumberedHeadingStyle(4)
        Case "format_painter"
            ExecuteQuickToolbarMso "FormatPainter"
        Case "paragraph_settings"
            ExecuteQuickToolbarMso "ParagraphDialog"
        Case "style_table_title"
            ApplyRAToolsStyle TableTitleStyle()
        Case "style_figure_title"
            ApplyRAToolsStyle FigureTitleStyle()
        Case "insert_cross_reference"
            ExecuteQuickToolbarMso "CrossReferenceInsert"
        Case "update_fields"
            ExecuteQuickToolbarMso "FieldsUpdate"
        Case "hyperlinks_fields_blue"
            SetHyperlinksAndFieldsToBlue
        Case "accept_revisions_comments"
            BatchAcceptAndClean
        Case "detect_highlights"
            BatchDetectHighlights
        Case Else
            Exit Function
    End Select

    TryRunExpandedQuickToolbarAction = True
End Function

Public Sub SetQuickToolbarMsoTestMode(ByVal enabled As Boolean)
    mQuickToolbarMsoTestMode = enabled
    mLastQuickToolbarMso = vbNullString
End Sub

Public Function GetLastQuickToolbarMsoForTest() As String
    GetLastQuickToolbarMsoForTest = mLastQuickToolbarMso
End Function

Private Sub ExecuteQuickToolbarMso(ByVal commandId As String)
    If mQuickToolbarMsoTestMode Then
        mLastQuickToolbarMso = commandId
    Else
        Application.CommandBars.ExecuteMso commandId
    End If
End Sub

