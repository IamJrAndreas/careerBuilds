Attribute VB_Name = "NewMacros"
'  Not for permanent macros. This is where macros go when recorded.

Sub HighlightDuplicateParagraphs()
Dim pFindTxt As String
Dim pReplaceTxt As String
' Based on the amended version by Emilia B in the comments at https://www.extendoffice.com/documents/word/5450-word-find-duplicate-sentences.html.
' ANDREAS: Please add some find/replace checks

' Add question asking whether to convert sentences to paragraphs. If so, F/R ". " with "^p".


' New section to replace things that might prevent detection of duplicates
Application.ScreenUpdating = False
pFindTxt = "  "
pReplaceTxt = " "
Call FindReplaceAnywhere(pFindTxt, pReplaceTxt)
Call FindReplaceAnywhere(pFindTxt, pReplaceTxt)
Call FindReplaceAnywhere(pFindTxt, pReplaceTxt)
Call FindReplaceAnywhere(pFindTxt, pReplaceTxt)

'pReplaceTxt = "[country]"
'pFindTxt = "Kenya"
'Call FindReplaceAnywhere(pFindTxt, pReplaceTxt)
'pFindTxt = "Rwanda"
'Call FindReplaceAnywhere(pFindTxt, pReplaceTxt)
'pFindTxt = "Seychelles"
'Call FindReplaceAnywhere(pFindTxt, pReplaceTxt)
'pFindTxt = "South Sudan"
'Call FindReplaceAnywhere(pFindTxt, pReplaceTxt)
'pFindTxt = "Djibouti"
'Call FindReplaceAnywhere(pFindTxt, pReplaceTxt)

'The above is no longer used. It is better to replace the country name in each file before moving them, otherwise putting them back in again at the end is complicated.


' Consider adding a section to manually find/replace other phrases

Dim StartTime, SecondsElapsed As Date
Dim secondsPerComparison As Double
Dim I, J, PC, totalComparisons, comparisonsDone, C, secondsToFinish As Long
Dim xRngFind, xRng As Range
Dim xStrg, minutesToFinish As String
Dim currentParag, nextParag As Paragraph
'Options.DefaultHighlightColorIndex = wdYellow
Application.ScreenUpdating = False
With ActiveDocument
StartTime = Now()
C = 0
PC = .Paragraphs.Count
totalComparisons = CLng((PC * (PC + 1)) / 2)
Set currentParag = .Paragraphs(1)
For I = 1 To PC - 1
'Debug.Print "processing paragraph " & I & " of a total of " & PC & " " & currentParag.Range.Text
'Debug.Print Len(currentParag) & currentParag
If currentParag.Range.HighlightColorIndex <> wdTeal Then ' Consider using hidden text instead of Teal, so that PerfectIt ignores it, though this might cause issues with checks such as acronyms.
If currentParag.Range.HighlightColorIndex <> wdGray50 Then
Set nextParag = currentParag
For J = I + 1 To PC
Set nextParag = nextParag.Next
If currentParag.Range.text = nextParag.Range.text Then
currentParag.Range.HighlightColorIndex = wdGray50
nextParag.Range.HighlightColorIndex = wdTeal
Debug.Print "found one!! " & amp; " I = " & amp; I & amp; " J = " & amp; J & amp; nextParag.Range.text
End If
Next
End If
End If
DoEvents
comparisonsDone = PC * (I - 1) + (J - I)
SecondsElapsed = DateDiff("s", StartTime, Now())
secondsPerComparison = CLng(SecondsElapsed) / comparisonsDone
secondsToFinish = CLng(secondsPerComparison * (totalComparisons - comparisonsDone))
minutesToFinish = Format(secondsToFinish / 86400, "hh:mm:ss")
elapsedTime = Format(SecondsElapsed / 86400, "hh:mm:ss")
Debug.Print "Finished procesing paragraph " & amp; I & amp; " of " & amp; PC & amp; ". Elapsed time = " & amp; elapsedTime & amp; ". Time to finish = " & amp; minutesToFinish
Set currentParag = currentParag.Next
Next
End With
Application.ScreenUpdating = True
MsgBox ("Macro finished")
End Sub

' Method 1 at https://www.datanumen.com/blogs/2-ways-quickly-merge-multiple-word-documents-one-via-vba/
Sub MergeMultiDocsIntoOne()
  Dim dlgFile As FileDialog
  Dim nTotalFiles As Integer
  Dim nEachSelectedFile As Integer

  Set dlgFile = Application.FileDialog(msoFileDialogFilePicker)
 
  With dlgFile
    .AllowMultiSelect = True
    If .Show <> -1 Then
      Exit Sub
    Else
      nTotalFiles = .SelectedItems.Count
    End If
  End With
 
  For nEachSelectedFile = 1 To nTotalFiles
    Selection.InsertFile dlgFile.SelectedItems.Item(nEachSelectedFile)
    If nEachSelectedFile < nTotalFiles Then
      Selection.InsertBreak Type:=wdPageBreak
    Else
      If nEachSelectedFile = nTotalFiles Then
        Exit Sub
      End If
    End If
  Next nEachSelectedFile
End Sub


Sub ExportParagraphEmptyDoc()
Attribute ExportParagraphEmptyDoc.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.ExportParagraphEmptyDoc"
'
' ExportParagraphEmptyDoc Macro
'
'
' Quick and dirty macro. Change the document names, and switch off clipboard history in Windows settings.
' Activated with ctrl+alt+w

    Selection.EscapeKey ' Exit the comment (if currently in one)
    Selection.MoveUp Unit:=wdParagraph, Count:=1 ' Move to start of paragraph
    Selection.MoveDown Unit:=wdParagraph, Count:=1, Extend:=wdExtend ' Move to end of paragraph
    Selection.Copy
    ActiveDocument.TrackRevisions = False
    Selection.Copy
'    Documents("temp-queries.docx").Paste
'    'AndFormat (wdPasteDefault)
    Documents("temp-queries.docx").Activate ' Change document name if necessary
'    Documents("temp-queries.docx").Select ' Change document name if necessary
    Selection.Paste
    Selection.MoveDown Unit:=wdParagraph, Count:=1
    Windows("Quantifying the impacts of tropospheric_JN 221318_en.docx"). _
        Activate
    ActiveDocument.TrackRevisions = True
End Sub


Sub KKGChar5()
Attribute KKGChar5.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.KKGSpecialChars"
'
' KKGSpecialChars Macro
'
'
    Selection.EscapeKey
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "¸"
        .Replacement.text = ChrW(450)
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    With Selection
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
        .Find.Execute Replace:=wdReplaceOne
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseEnd
        Else
            .Collapse Direction:=wdCollapseStart
        End If
        .Find.Execute
    End With
    With Selection
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
        .Find.Execute Replace:=wdReplaceOne
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseEnd
        Else
            .Collapse Direction:=wdCollapseStart
        End If
        .Find.Execute
    End With
End Sub
Sub KKGChar1()
Attribute KKGChar1.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.KKGChar1"
'
' KKGChar1 Macro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "Æ"
        .Replacement.text = ChrW(449)
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    With Selection
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
        .Find.Execute Replace:=wdReplaceOne
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseEnd
        Else
            .Collapse Direction:=wdCollapseStart
        End If
        .Find.Execute
    End With
    With Selection
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
        .Find.Execute Replace:=wdReplaceOne
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseEnd
        Else
            .Collapse Direction:=wdCollapseStart
        End If
        .Find.Execute
    End With
    With Selection
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
        .Find.Execute Replace:=wdReplaceOne
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseEnd
        Else
            .Collapse Direction:=wdCollapseStart
        End If
        .Find.Execute
    End With
End Sub
Sub KKGChar2()
Attribute KKGChar2.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.KKGChar2"
'
' KKGChar2 Macro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "!"
        .Replacement.text = ChrW(451)
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    With Selection
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
        .Find.Execute Replace:=wdReplaceOne
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseEnd
        Else
            .Collapse Direction:=wdCollapseStart
        End If
        .Find.Execute
    End With
    With Selection
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
        .Find.Execute Replace:=wdReplaceOne
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseEnd
        Else
            .Collapse Direction:=wdCollapseStart
        End If
        .Find.Execute
    End With
End Sub
Sub KKGChar3()
Attribute KKGChar3.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.KKGChar3"
'
' KKGChar3 Macro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "æ"
        .Replacement.text = ChrW(448)
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.Find.Execute
    Selection.Find.Execute
    Selection.Find.Execute
    Selection.Find.Execute
    Selection.Find.Execute
    Selection.Find.Execute
    Selection.Find.Execute
    With Selection
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
        .Find.Execute Replace:=wdReplaceOne
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseEnd
        Else
            .Collapse Direction:=wdCollapseStart
        End If
        .Find.Execute
    End With
    Selection.WholeStory
    Selection.Font.Name = "Calibri"
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "æ"
        .Replacement.text = ChrW(448)
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.Find.Execute
    Selection.Find.Execute
    Selection.Find.Execute
    With Selection
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
        .Find.Execute Replace:=wdReplaceOne
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseEnd
        Else
            .Collapse Direction:=wdCollapseStart
        End If
        .Find.Execute
    End With
    With Selection
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
        .Find.Execute Replace:=wdReplaceOne
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseEnd
        Else
            .Collapse Direction:=wdCollapseStart
        End If
        .Find.Execute
    End With
End Sub
Sub KKGChar4()
Attribute KKGChar4.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.KKGChar4"
'
' KKGChar4 Macro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "Å"
        .Replacement.text = ChrW(448)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    With Selection
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
        .Find.Execute Replace:=wdReplaceOne
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseEnd
        Else
            .Collapse Direction:=wdCollapseStart
        End If
        .Find.Execute
    End With
End Sub
Sub KKGCharNarrowExclamationMarkMatchCase()
Attribute KKGCharNarrowExclamationMarkMatchCase.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.KKGCharNarrowExclamationMarkMatchCase"
'
' KKGCharNarrowExclamationMarkMatchCase Macro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "æ"
        .Replacement.text = ChrW(451)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchByte = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.Find.Execute
    Selection.Find.Execute
End Sub
Sub RemoveContentControl()
' Source: https://eileenslounge.com/viewtopic.php?t=26995

Do While ActiveDocument.ContentControls.Count > 0
    For Each oCC In ActiveDocument.ContentControls
    If oCC.LockContentControl = True Then oCC.LockContentControl = False
    oCC.Delete False
Next
Loop
End Sub
Sub FindReplaceKKG()
' Problems:
' double pipe (Æ) takes more precedence than the narrow exclamation mark(æ).

With Selection.Find
.MatchCase = True
End With

'Call FindReplaceAnywhere("æ", ChrW(449)) 'narrow exclamation mark
Call FindReplaceAnywhere("Æ", ChrW(449)) 'double pipe
Call FindReplaceAnywhere("!", ChrW(451)) 'broad exclamation mark
Call FindReplaceAnywhere("Å", ChrW(448)) 'single pipe
Call FindReplaceAnywhere("¸", ChrW(450)) 'pipe with equal sign

End Sub

Sub UNNoAbbreviation()
Attribute UNNoAbbreviation.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.UNNoAbbreviation"
'
' Before you run this macro, select the phrase as well as the abbreviation.
'
' There is a way to use regular expression in a macro: https://stackoverflow.com/questions/25102372/how-to-use-enable-regexp-object-regular-expression-using-vba-macro-in-word
'
    Dim clipData As DataObject
    Dim txtPhrase As String
    Dim txtAbbr As String
    Dim doc As Document
    Dim regEx As RegExp
    Set regEx = New RegExp
    Set doc = ActiveDocument
    Set clipData = New DataObject
    

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "[!\(]{1,}" 'Searches for the phrase.
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend ' Move left one character of the space before the opening parenthesis (of the abbr.).
    Selection.Copy ' Copies the entire phrase without abbreviation.

    
    clipData.GetFromClipboard ' Gets the immediate copied clipboard content.
    txtPhrase = clipData.GetText ' Assigns the clipboard content to string variable.
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = " \([A-Za-z]@*\)" ' Searches for the abbreviation including the parenthesis.
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    
    Selection.Find.Execute
    Selection.Copy ' Copies the abbreviation to the clipboard.
    
    
    txtAbbr = clipData.GetText ' Gets the abbreviation from the clipboard.
    
    txtAbbr = Replace(txtAbbr, " ", "") ' Replaces the space before the opening parentheses with nothing.
    txtAbbr = Replace(txtAbbr, "(", "") ' Replaces the opening parenthesis with nothing.
    txtAbbr = Replace(txtAbbr, ")", "") ' Replaces the closing parenthesis with nothing.
    Debug.Print ("Copied clipboard content is: " & txtAbbr)
    
    
    With Selection
        If .Find.Forward = True Then ' Search in the forward direction of the selection.
            .Collapse Direction:=wdCollapseStart ' After the selection, move down.
        Else
            .Collapse Direction:=wdCollapseEnd ' After the selection, move up.
        End If
        .Find.Execute Replace:=wdReplaceOne '  Search for the next occurrence of the specified find criteria within the selected content and replace it (once at a time) with the specified replacement text.
        If .Find.Forward = True Then ' Find direction after previous replacement.
            .Collapse Direction:=wdCollapseEnd ' From the selection to the end of the document.
        Else
            .Collapse Direction:=wdCollapseStart ' Or from the selection to the start of the document.
        End If
    End With
    
    
    With doc.Content.Find
        .Forward = True ' Search direction: Downwards of the abbreviation.
        .ClearFormatting
        .MatchWholeWord = True
        .Wrap = wdFindContinue ' Allows the search to continue from the beginning if not found.
        .Execute txtAbbr, True, , , , , , , , txtPhrase, wdReplaceAll
    End With
    
End Sub
