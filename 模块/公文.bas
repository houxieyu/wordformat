Attribute VB_Name = "公文"
Sub 公文()
    Application.ScreenUpdating = False '关闭屏幕更新
    公文页面
    'A00_网页格式
    Selection.WholeStory
    删空行
    Selection.WholeStory
    'A00_公文正文格式
    公文正文
    公文一级标题
    公文二级标题
    ActiveDocument.Paragraphs(1).Range.Select
    公文标题
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.HomeKey Unit:=wdStory
    'A01_图片版式由浮动型转换为嵌入型
    Application.ScreenUpdating = True '恢复屏幕更新
End Sub

Sub 公文页面()
'
' Macro17 Macro
' 宏在 2013-4-2 由 戴宏国: 录制
'
    With ActiveDocument.Styles(wdStyleNormal).Font
        If .NameFarEast = .NameAscii Then
            .NameAscii = ""
        End If
        .NameFarEast = ""
    End With
    With ActiveDocument.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientPortrait
        .TopMargin = CentimetersToPoints(2.54)
        .BottomMargin = CentimetersToPoints(2.54)
        .LeftMargin = CentimetersToPoints(2.7)
        .RightMargin = CentimetersToPoints(2.7)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(1.5)
        .FooterDistance = CentimetersToPoints(1.75)
        .PageWidth = CentimetersToPoints(21)
        .PageHeight = CentimetersToPoints(29.7)
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
        .LayoutMode = wdLayoutModeLineGrid
    End With
End Sub


Sub 公文二级标题()
    
    Application.ScreenUpdating = False '关闭屏幕更新
    Selection.WholeStory
    Dim A As Variant
        A = Array("一", "二", "三", "四", "五", "六", "七", "八", "九", "十", "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十")
        C = Chr(-24157) '句号
    For j = 0 To UBound(A)
        For Each para In ActiveDocument.Paragraphs
            para.Range.Select
            TT = para.Range.Text
            If Left(TT, Len(A(j)) + 4) = "　　（" & A(j) & "）" Then
                L = Len(TT)
                P = InStr(1, TT, C, 1)
                If P > 0 Then
                    S1 = Left(TT, P - 1)
                    S2 = Right(TT, L - P)
                    Selection.Paragraphs(1).Range.Select
                    Selection.MoveLeft Unit:=wdCharacter, Count:=1
                    Selection.MoveRight Unit:=wdCharacter, Count:=Len(S1) + 1, Extend:=wdExtend
                        'Selection.Font.Bold = True
                        With Selection.Font
                            .NameFarEast = "楷体_GB2312"
                            .NameAscii = "楷体_GB2312"
                            .name = "楷体_GB2312"
                            .Size = 15
                            .Bold = True
                        End With
                    Selection.Paragraphs(1).Range.Select
                    Selection.MoveRight Unit:=wdCharacter, Count:=1
                Else
                    para.Range.Select
                    'Selection.Range.Font.Bold = True
                        With Selection.Font
                            .NameFarEast = "楷体_GB2312"
                            .NameAscii = "楷体_GB2312"
                            .name = "楷体_GB2312"
                            .Size = 15
                            .Bold = True
                        End With

                    Selection.MoveRight Unit:=wdCharacter, Count:=1
                End If
            End If
        Next para
        
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Next j
    
    Selection.WholeStory
    Selection.HomeKey Unit:=wdStory
    Application.ScreenUpdating = True '恢复屏幕更新
End Sub


Sub 公文正文小三仿行30()
    With Selection.Font
        .NameFarEast = "仿宋_GB2312"
        .NameAscii = "仿宋_GB2312"
        .name = "仿宋_GB2312"
        .Size = 15
        .Bold = False
    End With
    With Selection.ParagraphFormat
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 30
        .Alignment = wdAlignParagraphJustify
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
End Sub


Sub 公文标题()
Set MyRange = Selection.Range
    MyRange.Text = Trim(MyRange.Text)
    
    With MyRange.Font
        .NameFarEast = "方正小标宋_GBK"
        .NameAscii = "方正小标宋_GBK"
        .NameOther = "Times New Roman"
        .name = "方正小标宋_GBK"
        .Size = 22
        .Bold = True
    End With
    With MyRange.ParagraphFormat
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 30
        .Alignment = wdAlignParagraphCenter
        .LineUnitBefore = 0.5
        .LineUnitAfter = 0.5
    End With
End Sub


Sub 公文一级标题()
    Selection.WholeStory
    Dim A As Variant
        A = Array("一", "二", "三", "四", "五", "六", "七", "八", "九", "十", "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十")
    For j = 0 To UBound(A)
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="　　" & A(j) & "、"
        If .Found = True Then
            .Parent.Expand Unit:=wdParagraph
        Else
        Exit For
        End If
    End With
'    Selection.Paragraphs(1).Range.Font.Bold = True
    With Selection.Font
        .NameFarEast = "黑体"
        .NameAscii = "黑体"
        .NameOther = "Times New Roman"
        .name = "黑体"
        .Size = 16
        .Bold = False
    End With
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Next j
    
    Selection.WholeStory
    Selection.HomeKey Unit:=wdStory

End Sub


Sub 公文正文()
    With Selection.Font
        .NameFarEast = "仿宋_GB2312"
        .NameAscii = "仿宋_GB2312"
        .NameOther = "Times New Roman"
        .name = "仿宋_GB2312"
        .Size = 16
    End With
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 30
        .Alignment = wdAlignParagraphJustify
    End With
End Sub

Sub A00_公文页面设置()
    '将当前文档设置为公文页面规格
    Selection.WholeStory
    With ActiveDocument.Styles(wdStyleNormal).Font
        If .NameFarEast = .NameAscii Then
            .NameAscii = ""
        End If
        .NameFarEast = ""
    End With
    With ActiveDocument.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientPortrait
        .TopMargin = CentimetersToPoints(3.3)
        .BottomMargin = CentimetersToPoints(2.5)
        .LeftMargin = CentimetersToPoints(2.8)
        .RightMargin = CentimetersToPoints(2.8)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(1.5)
        .FooterDistance = CentimetersToPoints(2.1)
        .PageWidth = CentimetersToPoints(21)
        .PageHeight = CentimetersToPoints(29.7)
        .FirstPageTray = wdPrinterDefaultBin
        .OtherPagesTray = wdPrinterDefaultBin
        .SectionStart = wdSectionNewPage
        .OddAndEvenPagesHeaderFooter = True
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = wdAlignVerticalTop
        .SuppressEndnotes = False
        .MirrorMargins = False
        .TwoPagesOnOne = False
        .BookFoldPrinting = False
        .BookFoldRevPrinting = False
        .BookFoldPrintingSheets = 1
        .GutterPos = wdGutterPosLeft
        .LayoutMode = wdLayoutModeLineGrid
    End With
End Sub

Sub A00_公文标题()
'
' A00_公文标题 宏
'
'
    With Selection.Font
        .NameFarEast = "方正小标宋_GBK"
        .NameAscii = "方正小标宋_GBK"
        .NameOther = "黑体"
        .name = "方正小标宋_GBK"
        .Size = 22
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Spacing = 0
        .Scaling = 100
        .Position = 0
        .Kerning = 1
        .Animation = wdAnimationNone
        .DisableCharacterSpaceGrid = False
        .EmphasisMark = wdEmphasisMarkNone
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 30
        .Alignment = wdAlignParagraphCenter
        .WidowControl = False
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
        .AutoAdjustRightIndent = False
        .DisableLineHeightGrid = True
        .FarEastLineBreakControl = True
        .WordWrap = True
        .HangingPunctuation = True
        .HalfWidthPunctuationOnTopOfLine = False
        .AddSpaceBetweenFarEastAndAlpha = True
        .AddSpaceBetweenFarEastAndDigit = True
        .BaseLineAlignment = wdBaselineAlignAuto
    End With
End Sub
Sub A00_公文正文格式()
    '对选定区域按公文正文格式要求进行设置，字体为3号仿宋_GB2312，首行缩进2字符
    With Selection.Font
        .NameFarEast = "仿宋_GB2312"
        .NameAscii = "仿宋_GB2312"
        .NameOther = ""
        .name = "仿宋_GB2312"
        .Size = 16
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Spacing = -0.2
        .Scaling = 100
        .Position = 0
        .Kerning = 1
        .Animation = wdAnimationNone
        .DisableCharacterSpaceGrid = False
        .EmphasisMark = wdEmphasisMarkNone
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 28
        .Alignment = wdAlignParagraphJustify
        .WidowControl = False
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0.35)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 2
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
        .AutoAdjustRightIndent = False
        .DisableLineHeightGrid = True
        .FarEastLineBreakControl = True
        .WordWrap = True
        .HangingPunctuation = True
        .HalfWidthPunctuationOnTopOfLine = False
        .AddSpaceBetweenFarEastAndAlpha = True
        .AddSpaceBetweenFarEastAndDigit = True
        .BaseLineAlignment = wdBaselineAlignAuto
    End With
End Sub

Sub A00_网页格式()

    Dim MyRange As Range
    Dim O As Variant
    Dim R As Variant
    Dim A As Variant
    Dim C As Variant
    Dim D As Variant
    Dim CC As Variant
    Dim DD As Variant
    
    On Error Resume Next
    Application.ScreenUpdating = False '关闭屏幕更新
    
    O = Array("^l", Chr(32) & Chr(32), "^p^p", "^p", "　　^p", Chr(32) & Chr(13), Chr(58) & Chr(13))
    R = Array("^p", Chr(-24159), "^p", "^p　　", "", Chr(13), Chr(-23622) & Chr(13) & Chr(13))
    A = Array("一", "二", "三", "四", "五", "六", "七", "八", "九", "十", _
          , "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十", _
          "二十一", "二十二", "二十三", "二十四", "二十五", "二十六", "二十七", "二十八", "二十九", "三十")
    C = Array("０", "１", "２", "３", "４", "５", "６", "７", "８", "９", ",", ";", "％", "?", "(", ")")
    D = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "，", "；", "%", ".", "（", "）")
    CC = Array("^l", ",", ";", "０", "１", "２", "３", "４", "５", "６", "７", "８", "９", "．", "Ａ", "Ｂ", "Ｃ", "Ｄ", "Ｅ", "Ｆ", "Ｇ", _
        "Ｈ", "Ｉ", "Ｊ", "Ｋ", "Ｌ", "Ｍ", "Ｎ", "Ｏ", "Ｐ", "Ｑ", "Ｒ", "Ｓ", "Ｔ", "Ｕ", "Ｖ", "Ｗ", "Ｘ", "Ｙ", "Ｚ", _
        "ａ", "ｂ", "ｃ", "ｄ", "ｅ", "ｆ", "ｇ", "ｈ", "ｉ", "ｊ", "ｋ", "ｌ", "ｍ", "ｎ", "ｏ", "ｐ", "ｑ", "ｒ", "ｓ", _
        "ｔ", "ｕ", "ｖ", "ｗ", "ｘ", "ｙ", "ｚ")
    DD = Array("^p", "，", "；", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", _
        "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "a", "b", "c", "d", "e", "f", "g", _
        "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z")
    
    A00_删除网页空格
    
    '删除段落前面空格
    If Len(Selection.Range.Text) = 0 Then Selection.WholeStory
        Set MyRange = Selection.Range
        Selection.ClearFormatting
        With MyRange
            pn = MyRange.Paragraphs.Count
            For i = 1 To pn
                T = Trim(MyRange.Paragraphs(i).Range.Text)
                MyRange.Paragraphs(i).Range.Text = Trim(Left(T, Len(T) - 1)) & Right(T, 1)
            Next i
        End With
        Selection.WholeStory
        Selection.EndKey Unit:=wdStory
        Selection.Delete Unit:=wdCharacter, Count:=1
    
    Selection.WholeStory
    Selection.ClearFormatting
    Selection.HomeKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.WholeStory
    With Selection.Font
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Size = 12
    End With
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
    End With
 
    For i = 0 To UBound(O)
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = O(i)
        .Replacement.Text = R(i)
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i
    
    Selection.WholeStory
 
    For j = 0 To UBound(A)
    
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="　　" & A(j) & "、"
        If .Found = True Then
            .Parent.Expand Unit:=wdParagraph
        Else
            Exit For
        End If
    End With
    Selection.Paragraphs(1).Range.Font.Bold = True
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Next j
    
    Selection.WholeStory
    'Selection.Copy
    Selection.HomeKey Unit:=wdStory
    Selection.Delete Unit:=wdCharacter, Count:=1

    Selection.HomeKey Unit:=wdStory
    For i = 0 To UBound(C)
    With Selection.Find
        .Text = C(i)
        .Replacement.Text = D(i)
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i
    
    Selection.HomeKey Unit:=wdStory
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = "^p^p"
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.WholeStory
    
    Selection.HomeKey Unit:=wdStory
    For i = 0 To UBound(CC)
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = CC(i)
            .Replacement.Text = DD(i)
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
    Next i

    
    Selection.EndKey Unit:=wdStory
    Selection.TypeBackspace
    Selection.TypeBackspace
    Selection.TypeBackspace
    'Selection.TypeBackspace
    Selection.HomeKey Unit:=wdStory
    
    Selection.HomeKey Unit:=wdStory '定位到文档开头
    
    '第一段如果为标题，则加粗居中
    Selection.Paragraphs(1).Range.Select
        s = Selection.Paragraphs(1).Range.Text
        If Len(s) < 30 Then
            Selection.Font.Bold = True
            Selection.Paragraphs(1).Range.Text = Trim(s)
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End If
        
        '替换间隔号
        Selection.HomeKey Unit:=wdStory '定位到文档开头
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = ChrW(8226)
            .Replacement.Text = "·"
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        
        Selection.HomeKey Unit:=wdStory '定位到文档开头
    Application.ScreenUpdating = True '恢复屏幕更新
    
End Sub

