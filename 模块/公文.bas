Attribute VB_Name = "����"
Sub ����()
    Application.ScreenUpdating = False '�ر���Ļ����
    ����ҳ��
    'A00_��ҳ��ʽ
    Selection.WholeStory
    ɾ����
    Selection.WholeStory
    'A00_�������ĸ�ʽ
    ��������
    ����һ������
    ���Ķ�������
    ActiveDocument.Paragraphs(1).Range.Select
    ���ı���
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.HomeKey Unit:=wdStory
    'A01_ͼƬ��ʽ�ɸ�����ת��ΪǶ����
    Application.ScreenUpdating = True '�ָ���Ļ����
End Sub

Sub ����ҳ��()
'
' Macro17 Macro
' ���� 2013-4-2 �� �����: ¼��
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


Sub ���Ķ�������()
    
    Application.ScreenUpdating = False '�ر���Ļ����
    Selection.WholeStory
    Dim A As Variant
        A = Array("һ", "��", "��", "��", "��", "��", "��", "��", "��", "ʮ", "ʮһ", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "��ʮ")
        C = Chr(-24157) '���
    For j = 0 To UBound(A)
        For Each para In ActiveDocument.Paragraphs
            para.Range.Select
            TT = para.Range.Text
            If Left(TT, Len(A(j)) + 4) = "������" & A(j) & "��" Then
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
                            .NameFarEast = "����_GB2312"
                            .NameAscii = "����_GB2312"
                            .name = "����_GB2312"
                            .Size = 15
                            .Bold = True
                        End With
                    Selection.Paragraphs(1).Range.Select
                    Selection.MoveRight Unit:=wdCharacter, Count:=1
                Else
                    para.Range.Select
                    'Selection.Range.Font.Bold = True
                        With Selection.Font
                            .NameFarEast = "����_GB2312"
                            .NameAscii = "����_GB2312"
                            .name = "����_GB2312"
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
    Application.ScreenUpdating = True '�ָ���Ļ����
End Sub


Sub ��������С������30()
    With Selection.Font
        .NameFarEast = "����_GB2312"
        .NameAscii = "����_GB2312"
        .name = "����_GB2312"
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


Sub ���ı���()
Set MyRange = Selection.Range
    MyRange.Text = Trim(MyRange.Text)
    
    With MyRange.Font
        .NameFarEast = "����С����_GBK"
        .NameAscii = "����С����_GBK"
        .NameOther = "Times New Roman"
        .name = "����С����_GBK"
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


Sub ����һ������()
    Selection.WholeStory
    Dim A As Variant
        A = Array("һ", "��", "��", "��", "��", "��", "��", "��", "��", "ʮ", "ʮһ", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "��ʮ")
    For j = 0 To UBound(A)
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="����" & A(j) & "��"
        If .Found = True Then
            .Parent.Expand Unit:=wdParagraph
        Else
        Exit For
        End If
    End With
'    Selection.Paragraphs(1).Range.Font.Bold = True
    With Selection.Font
        .NameFarEast = "����"
        .NameAscii = "����"
        .NameOther = "Times New Roman"
        .name = "����"
        .Size = 16
        .Bold = False
    End With
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Next j
    
    Selection.WholeStory
    Selection.HomeKey Unit:=wdStory

End Sub


Sub ��������()
    With Selection.Font
        .NameFarEast = "����_GB2312"
        .NameAscii = "����_GB2312"
        .NameOther = "Times New Roman"
        .name = "����_GB2312"
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

Sub A00_����ҳ������()
    '����ǰ�ĵ�����Ϊ����ҳ����
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

Sub A00_���ı���()
'
' A00_���ı��� ��
'
'
    With Selection.Font
        .NameFarEast = "����С����_GBK"
        .NameAscii = "����С����_GBK"
        .NameOther = "����"
        .name = "����С����_GBK"
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
Sub A00_�������ĸ�ʽ()
    '��ѡ�����򰴹������ĸ�ʽҪ��������ã�����Ϊ3�ŷ���_GB2312����������2�ַ�
    With Selection.Font
        .NameFarEast = "����_GB2312"
        .NameAscii = "����_GB2312"
        .NameOther = ""
        .name = "����_GB2312"
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

Sub A00_��ҳ��ʽ()

    Dim MyRange As Range
    Dim O As Variant
    Dim R As Variant
    Dim A As Variant
    Dim C As Variant
    Dim D As Variant
    Dim CC As Variant
    Dim DD As Variant
    
    On Error Resume Next
    Application.ScreenUpdating = False '�ر���Ļ����
    
    O = Array("^l", Chr(32) & Chr(32), "^p^p", "^p", "����^p", Chr(32) & Chr(13), Chr(58) & Chr(13))
    R = Array("^p", Chr(-24159), "^p", "^p����", "", Chr(13), Chr(-23622) & Chr(13) & Chr(13))
    A = Array("һ", "��", "��", "��", "��", "��", "��", "��", "��", "ʮ", _
          , "ʮһ", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "��ʮ", _
          "��ʮһ", "��ʮ��", "��ʮ��", "��ʮ��", "��ʮ��", "��ʮ��", "��ʮ��", "��ʮ��", "��ʮ��", "��ʮ")
    C = Array("��", "��", "��", "��", "��", "��", "��", "��", "��", "��", ",", ";", "��", "?", "(", ")")
    D = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "��", "��", "%", ".", "��", "��")
    CC = Array("^l", ",", ";", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", _
        "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", _
        "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", _
        "��", "��", "��", "��", "��", "��", "��")
    DD = Array("^p", "��", "��", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", _
        "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "a", "b", "c", "d", "e", "f", "g", _
        "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z")
    
    A00_ɾ����ҳ�ո�
    
    'ɾ������ǰ��ո�
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
        .NameFarEast = "����"
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
        .Execute FindText:="����" & A(j) & "��"
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
    
    Selection.HomeKey Unit:=wdStory '��λ���ĵ���ͷ
    
    '��һ�����Ϊ���⣬��Ӵ־���
    Selection.Paragraphs(1).Range.Select
        s = Selection.Paragraphs(1).Range.Text
        If Len(s) < 30 Then
            Selection.Font.Bold = True
            Selection.Paragraphs(1).Range.Text = Trim(s)
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End If
        
        '�滻�����
        Selection.HomeKey Unit:=wdStory '��λ���ĵ���ͷ
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = ChrW(8226)
            .Replacement.Text = "��"
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        
        Selection.HomeKey Unit:=wdStory '��λ���ĵ���ͷ
    Application.ScreenUpdating = True '�ָ���Ļ����
    
End Sub

