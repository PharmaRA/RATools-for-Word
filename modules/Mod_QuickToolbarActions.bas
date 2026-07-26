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

Private Function NumberedHeadingStyle(ByVal level As Long) As String
    NumberedHeadingStyle = FromCodePoints(Array(26631, 39064)) & CStr(level) & "-F"
End Function

Private Function UnnumberedHeadingStyle(ByVal level As Long) As String
    UnnumberedHeadingStyle = FromCodePoints(Array(26080, 32534, 21495, 26631, 39064)) & _
                             CStr(level) & "-F"
End Function

Private Function TableTitleStyle() As String
    TableTitleStyle = FromCodePoints(Array(34920, 26631, 39064)) & "-F"
End Function

Private Function FigureTitleStyle() As String
    FigureTitleStyle = FromCodePoints(Array(22270, 26631, 39064)) & "-F"
End Function

Private Function FromCodePoints(ByVal values As Variant) As String
    Dim i As Long

    For i = LBound(values) To UBound(values)
        FromCodePoints = FromCodePoints & ChrW$(CLng(values(i)))
    Next i
End Function
