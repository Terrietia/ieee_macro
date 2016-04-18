Sub body_macro()
'
' AUTHOR: Matthew Valuet
' DESCRIPTION: Macro for formatting body text for IEEE two column document
' DATE: 11 April 2016
'
' Setting page margins.
With ActiveDocument.PageSetup
    .TopMargin = InchesToPoints(0.75)
    .BottomMargin = InchesToPoints(1)
    .LeftMargin = InchesToPoints(0.63)
    .RightMargin = InchesToPoints(0.63)
End With
'
' Setting paragraph formatting.
With ActiveDocument.Paragraphs
    .Alignment = wdAlignParagraphJustify
    .LineSpacingRule = wdLineSpaceMultiple
    .LineSpacing = LinesToPoints(0.95)
    .LeftIndent = InchesToPoints(0)
    .RightIndent = InchesToPoints(0)
    .FirstLineIndent = InchesToPoints(0.2)
    .SpaceBefore = 0
    .SpaceAfter = 6
End With
'
' Setting column width and spacing.
With ActiveDocument.PageSetup.TextColumns
    .SetCount NumColumns:=2
    .EvenlySpaced = True
    .LineBetween = False
    .Width = InchesToPoints(3.5)
    .Spacing = InchesToPoints(0.24)
End With
'
' Setting list levels such that the lists do not change with the running
' of the macro. Specifically, this block is for ListLevel(1).
With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(1)
    .NumberFormat = "%1."
    .TrailingCharacter = wdTrailingTab
    .NumberStyle = wdListNumberStyleUppercaseRoman
    .NumberPosition = InchesToPoints(0.15)
    .Alignment = wdListLevelAlignCenter
    .TabPosition = InchesToPoints(0.4)
    .ResetOnHigher = 0
    .StartAt = 1
    With .Font
        .Bold = False
        .Italic = False
        .StrikeThrough = False
        .Subscript = False
        .Superscript = False
        .Shadow = False
        .Outline = False
        .Emboss = False
        .Engrave = False
        .AllCaps = False
        .Hidden = False
        .Underline = False
        .Color = wdColorAutomatic
        .Size = 10
        .DoubleStrikeThrough = False
        .Name = "Times New Roman"
    End With
    .LinkedStyle = "Heading 1"
End With
'
' Setting list levels such that the lists do not change with the running
' of the macro. Specifically, this block is for ListLevel(2).
With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(2)
    .NumberFormat = "%2."
    .TrailingCharacter = wdTrailingTab
    .NumberStyle = wdListNumberStyleUppercaseLetter
    .NumberPosition = InchesToPoints(0)
    .Alignment = wdListLevelAlignLeft
    .TextPosition = InchesToPoints(0.2)
    .TabPosition = InchesToPoints(0.25)
    .ResetOnHigher = 1
    .StartAt = 1
    With .Font
        .Bold = False
        .Italic = True
        .StrikeThrough = False
        .Subscript = False
        .Superscript = False
        .Shadow = False
        .Outline = False
        .Emboss = False
        .Engrave = False
        .AllCaps = False
        .Hidden = False
        .Underline = False
        .Color = wdColorAutomatic
        .Size = 10
        .DoubleStrikeThrough = False
        .Name = "Times New Roman"
    End With
    .LinkedStyle = "Heading 2"
End With
'
' Setting list levels such that the lists do not change with the running
' of the macro. Specifically, this block is for ListLevel(3).
With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(3)
    .NumberFormat = "%3)"
    .TrailingCharacter = wdTrailingTab
    .NumberStyle = wdListNumberStyleArabic
    .Alignment = wdListLevelAlignLeft
    .TabPosition = InchesToPoints(0.38)
    .ResetOnHigher = 2
    .StartAt = 1
    With .Font
        .Bold = False
        .Italic = True
        .StrikeThrough = False
        .Subscript = False
        .Superscript = False
        .Shadow = False
        .Outline = False
        .Emboss = False
        .Engrave = False
        .AllCaps = False
        .Hidden = False
        .Underline = False
        .Color = wdColorAutomatic
        .Size = 10
        .DoubleStrikeThrough = False
        .Name = "Times New Roman"
    End With
    .LinkedStyle = "Heading 3"
End With
'
' Miscellaneous list formatting and editing.
ListGalleries(wdOutlineNumberGallery).ListTemplates(1).Name = ""
Selection.Range.ListFormat.ApplyListTemplate ListTemplate:=ListGalleries( _
    wdOutlineNumberGallery).ListTemplates(1), ContinuePreviousList:=True, _
    ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
    wdWord10ListBehavior
End Sub
