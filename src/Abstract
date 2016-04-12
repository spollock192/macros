Sub Abstract()
'
' Abstract Macro
'
'
    Selection.font.Bold = wdToggle
    Selection.font.Size = 9
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="Abstact—"
    Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
    Selection.font.Italic = wdToggle
    With Selection.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .FirstLineIndent = InchesToPoints(0.19)
    End With
End Sub
