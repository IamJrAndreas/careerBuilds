Attribute VB_Name = "NewMacros1"
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
