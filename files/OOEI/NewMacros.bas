Attribute VB_Name = "NewMacros9"
Sub Macros()
    
    'Установки для курсача
    Selection.Font.Name = "Times New Roman"
    Selection.Font.Size = 14
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    ActiveDocument.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpace1pt5
    With Selection.PageSetup
        .TopMargin = CentimetersToPoints(2)
        .BottomMargin = CentimetersToPoints(2)
        .LeftMargin = CentimetersToPoints(2)
        .RightMargin = CentimetersToPoints(2)
    End With
   
    'Ставиться пробел после знака пунктуации
    With Selection.Find
        .Text = "([.,:;\!\?])"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdStore
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=1

End Sub
