Attribute VB_Name = "Mod_NormalizeScientificTerms"
Option Explicit

Public Sub NormalizeScientificTerms()
    NormalizeScientificTermsForCurrentTarget True
End Sub

Public Function NormalizeScientificTermsForCurrentTarget(Optional ByVal showMessage As Boolean = True) As Long
    On Error GoTo ErrHandler

    Dim targetRange As Range
    Dim formattedCount As Long

    If Selection.Type = wdSelectionIP Then
        Set targetRange = ActiveDocument.Content
    Else
        Set targetRange = Selection.Range
    End If

    formattedCount = NormalizeScientificTermsInRange(targetRange)
    NormalizeScientificTermsForCurrentTarget = formattedCount

    If showMessage Then
        MsgBox "Formatted " & formattedCount & " scientific term(s).", _
               vbInformation, "Normalize Scientific Terms"
    End If

    Exit Function

ErrHandler:
    NormalizeScientificTermsForCurrentTarget = 0
    If showMessage Then
        MsgBox "Scientific term formatting failed: " & Err.Description, _
               vbExclamation, "Normalize Scientific Terms"
    End If
End Function

Public Function NormalizeScientificTermsInRange(ByVal targetRange As Range) As Long
    If targetRange Is Nothing Then Exit Function

    Dim sourceText As String
    Dim formattedCount As Long

    sourceText = targetRange.Text

    formattedCount = formattedCount + ApplyTermPattern(targetRange, sourceText, _
        "(^|[^A-Za-z0-9_])(STD[0-9]+)(?=[^A-Za-z0-9_]|$)", 3)
    formattedCount = formattedCount + ApplyTermPattern(targetRange, sourceText, _
        "(^|[^A-Za-z0-9_])([CT](max|min))(?=[^A-Za-z0-9_]|$)", 1)

    formattedCount = formattedCount + ApplyTermPattern(targetRange, sourceText, _
        "(^|[^A-Za-z0-9_])(AUC0-t)(?=[^A-Za-z0-9_]|$)", 3)
    formattedCount = formattedCount + ApplyTermPattern(targetRange, sourceText, _
        "(^|[^A-Za-z0-9_])(AUC0-inf)(?=[^A-Za-z0-9_]|$)", 3)
    formattedCount = formattedCount + ApplyTermPattern(targetRange, sourceText, _
        "(^|[^A-Za-z0-9_])(AUC0-last)(?=[^A-Za-z0-9_]|$)", 3)
    formattedCount = formattedCount + ApplyTermPattern(targetRange, sourceText, _
        "(^|[^A-Za-z0-9_])(AUC0-24h)(?=[^A-Za-z0-9_]|$)", 3)
    formattedCount = formattedCount + ApplyTermPattern(targetRange, sourceText, _
        "(^|[^A-Za-z0-9_])(AUC0-" & ChrW$(&H221E) & ")(?=[^A-Za-z0-9_]|$)", 3)
    formattedCount = formattedCount + ApplyTermPattern(targetRange, sourceText, _
        "(^|[^A-Za-z0-9_])(AUC0-" & ChrW$(&H3C4) & ")(?=[^A-Za-z0-9_]|$)", 3)
    formattedCount = formattedCount + ApplyTermPattern(targetRange, sourceText, _
        "(^|[^A-Za-z0-9_])(AUC0-tau)(?=[^A-Za-z0-9_]|$)", 3)
    formattedCount = formattedCount + ApplyTermPattern(targetRange, sourceText, _
        "(^|[^A-Za-z0-9_])(AUCtau)(?=[^A-Za-z0-9_]|$)", 3)
    formattedCount = formattedCount + ApplyTermPattern(targetRange, sourceText, _
        "(^|[^A-Za-z0-9_])(AUClast)(?=[^A-Za-z0-9_]|$)", 3)
    formattedCount = formattedCount + ApplyTermPattern(targetRange, sourceText, _
        "(^|[^A-Za-z0-9_])(t1/2)(?=[^A-Za-z0-9_]|$)", 1)
    formattedCount = formattedCount + ApplyTermPattern(targetRange, sourceText, _
        "(^|[^A-Za-z0-9_])(C(0|t))(?=[^A-Za-z0-9_]|$)", 1)
    formattedCount = formattedCount + ApplyTermPattern(targetRange, sourceText, _
        "(^|[^A-Za-z0-9_])(C(ss|avg|ave|trough))(?=[^A-Za-z0-9_]|$)", 1)
    formattedCount = formattedCount + ApplyTermPattern(targetRange, sourceText, _
        "(^|[^A-Za-z0-9_])(V[dz])(?=[^A-Za-z0-9_]|$)", 1)
    formattedCount = formattedCount + ApplyTermPattern(targetRange, sourceText, _
        "(^|[^A-Za-z0-9_])(CLz)(?=[^A-Za-z0-9_]|$)", 2)

    NormalizeScientificTermsInRange = formattedCount
End Function

Private Function ApplyTermPattern(ByVal targetRange As Range, _
                                  ByVal sourceText As String, _
                                  ByVal pattern As String, _
                                  ByVal baseLength As Long) As Long
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim prefixText As String
    Dim termText As String
    Dim termStart As Long
    Dim formattedCount As Long

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = False
    regex.Pattern = pattern

    Set matches = regex.Execute(sourceText)

    For Each match In matches
        prefixText = CStr(match.SubMatches(0))
        termText = CStr(match.SubMatches(1))
        termStart = targetRange.Start + match.FirstIndex + Len(prefixText)

        ApplyTermFormatting targetRange, termStart, Len(termText), baseLength
        formattedCount = formattedCount + 1
    Next match

    ApplyTermPattern = formattedCount
End Function

Private Sub ApplyTermFormatting(ByVal targetRange As Range, _
                                ByVal termStart As Long, _
                                ByVal termLength As Long, _
                                ByVal baseLength As Long)
    Dim termRange As Range
    Dim suffixRange As Range

    Set termRange = targetRange.Duplicate
    termRange.Start = termStart
    termRange.End = termStart + termLength
    termRange.Font.Subscript = False

    Set suffixRange = targetRange.Duplicate
    suffixRange.Start = termStart + baseLength
    suffixRange.End = termStart + termLength
    suffixRange.Font.Subscript = True
End Sub
