Sub Body()
'Author Stu
'Macro for IEEE 2 column that will format a word document to have proper margins, line spacing, font.
'*WARNING:zRUN THIS MACRO BEFORE RUNNING ANY OTHER MACRO, THIS MACRO AFFECTS ALL TEXT IN DOCUMENT*
'
    Selection.WholeStory
'
'Changes page margins
    With Selection.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientPortrait
        .TopMargin = InchesToPoints(0.75)
        .BottomMargin = InchesToPoints(1)
        .LeftMargin = InchesToPoints(0.63)
        .RightMargin = InchesToPoints(0.63)
        .Gutter = InchesToPoints(0)
        .HeaderDistance = InchesToPoints(0.5)
        .FooterDistance = InchesToPoints(0.5)
        .PageWidth = InchesToPoints(8.5)
        .PageHeight = InchesToPoints(11)
        .FirstPageTray = wdPrinterDefaultBin
        .OtherPagesTray = wdPrinterDefaultBin
        .SectionStart = wdSectionNewPage
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = wdAlignVerticalTop
        .SuppressEndnotes = False
        .MirrorMargins = False
        .TwoPagesOnOne = False
        .BookFoldPrinting = False
        .BookFoldRevPrinting = False
        .BookFoldPrintingSheets = 1
        .GutterPos = wdGutterPosLeft
    End With
    
'Changes line spacing
    With Selection.ParagraphFormat
        .LeftIndent = InchesToPoints(0)
        .RightIndent = InchesToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 6
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = InchesToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
    End With
    
'Justifies text
    With Selection.ParagraphFormat
        .LeftIndent = InchesToPoints(0)
        .RightIndent = InchesToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 6
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphJustify
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = InchesToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
    End With
    
'Puts text into two columns
    Selection.WholeStory
    If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
    End If
    If ActiveWindow.ActivePane.View.Type <> wdPrintView Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    With ActiveDocument.PageSetup.TextColumns
        .SetCount NumColumns:=2
        .EvenlySpaced = True
        .LineBetween = False
        .Width = InchesToPoints(3.5)
        .Spacing = InchesToPoints(0.25)
        
'Changes font into times new roman
    End With
        Selection.WholeStory
        Selection.font.Name = "Times New Roman"
'need font size change here
End Sub
