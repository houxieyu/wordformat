
Public Sub �����Ű�()
    �Զ��Ű� 2
End Sub

Public Sub �����Ű�()
    �Զ��Ű� 1
End Sub

Public Sub ֪ͨ�Ű�()
    �Զ��Ű� 3
End Sub
Private Sub �Զ��Ű�(pbtype As Integer)
    If Application.Documents.Count = 0 Then
        Exit Sub
    End If
    ActiveDocument.TrackRevisions = False '�ر��޶�
    Application.ScreenUpdating = False '�ر���Ļ����
    ����ҳ��
    ����ո�
    �������
    ��������
    ���ı��� '�����ʽ������λ�ڵ�һ�У���û�зֶ�
    ��������
    һ������ 'һ��������"һ��"Ϊ������λ������
    �������� pbtype '����������"��һ��"Ϊ��������Ϊ�������ţ���λ������
    ͼƬ����
    ͼƬ����
    ��ע
    If pbtype = 2 Then
        ����ʽ
    End If
    ��ע
    ��Ҫ
    ����˵����λ����
    ���Ļ��ظ�ʽ����
    ����ҳ��
    Application.ScreenUpdating = True '�ָ���Ļ����
    Selection.HomeKey unit:=wdStory
End Sub

Private Sub ����ҳ��2()
    '
    '
    With ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary)
        Set rng = .Range
        rng.Font.Size = 16
        rng.Font.name = "Times New Roman"
        rng.Text = "- "
        rng.Collapse wdCollapseEnd
        ActiveDocument.Fields.Add rng, wdFieldEmpty, "Page"
        Set rng = .Range
        rng.Collapse wdCollapseEnd
        rng.Text = " -"
        .Range.Fields.update
        .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End With
    If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
    End If
    If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
        ActivePane.View.Type = wdOutlineView Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    With Selection.ParagraphFormat
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth075pt
        .DefaultBorderColor = wdColorAutomatic
    End With
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
End Sub
'�������ĩβ
Private Function clearParagraphEnd(str As String) As String
    clearParagraphEnd = Replace(Replace(str, Chr(10), ""), Chr(13), "")
End Function
'�����ʽ����������ǰ�и����򸽼�+���֣���ð�ţ��Ҷ����ɶΡ�������һ��Ϊ������
Private Sub ��������()
    'ǰ�����ҳ��
    For i = 1 To ActiveDocument.Paragraphs.Count
        With ActiveDocument.Paragraphs(i).Range
            rgntxt = Replace(Replace(Replace(.Text, Chr(10), ""), Chr(13), ""), Chr(12), "")
            Dim isfj As Boolean
            isfj = False
            For j = 1 To 20
                If rgntxt = "����" & j Then
                    Debug.Print "����" & j
                    isfj = True
                    Exit For
                End If
            Next j
            If rgntxt = "����" Or isfj = True Then
                .Select
                With Selection.Font
                    .NameFarEast = "����"
                    .NameAscii = "����"
                    .NameOther = "Times New Roman"
                    .name = "����"
                    .Size = 16
                    .Bold = False
                End With
                With Selection.ParagraphFormat
                    .Alignment = wdAlignParagraphLeft
                    .CharacterUnitLeftIndent = 0
                    .CharacterUnitFirstLineIndent = 0
                    .FirstLineIndent = CentimetersToPoints(0)
                End With
                '�����ҳ��
                If Left(Selection.Text, 1) <> Chr(12) Then
                    Selection.MoveLeft
                    Selection.InsertBreak Type:=wdPageBreak
                    Selection.Expand unit:=wdParagraph
                End If
                '������������м���ӿ���
                Selection.InsertAfter vbCrLf
                '�������ʽ
                Selection.MoveDown
                Selection.Expand unit:=wdParagraph
                With Selection.Font
                    .NameFarEast = "����"
                    .NameAscii = "����"
                    .NameOther = "Times New Roman"
                    .name = "����"
                    .Size = 22
                    .Bold = False
                End With
                With Selection.ParagraphFormat
                    .Alignment = wdAlignParagraphCenter
                    .CharacterUnitLeftIndent = 0
                    .CharacterUnitFirstLineIndent = 0
                    .FirstLineIndent = CentimetersToPoints(0)
                    .LineUnitBefore = 1
                    .LineUnitAfter = 1
                End With
                '����������
                Selection.MoveDown
                Selection.Expand unit:=wdParagraph
                If Left(Right(Selection.Text, 2), 1) <> "��" And Left(Right(Selection.Text, 2), 1) <> "��" And Left(Selection.Text, 2) <> "һ��" Then
                    '������
                    With Selection.Font
                        .NameFarEast = "����_GB2312"
                        .NameAscii = "����_GB2312"
                        .NameOther = "Times New Roman"
                        .name = "����_GB2312"
                        .Size = 16
                        .Bold = False
                    End With
                    With Selection.ParagraphFormat
                        .LineSpacingRule = wdLineSpaceExactly
                        .LineSpacing = 33
                        .Alignment = wdAlignParagraphCenter
                        .LineUnitBefore = 0
                        .SpaceBefore = 0
                        .LineUnitAfter = 1
                        .CharacterUnitFirstLineIndent = 0
                        .FirstLineIndent = CentimetersToPoints(0)
                    End With
                    '����������κ���Ϊ0
                    Selection.MoveLeft
                    Selection.MoveUp unit:=wdParagraph
                    Selection.Expand unit:=wdParagraph
                    Selection.ParagraphFormat.LineUnitAfter = 0
                    Selection.ParagraphFormat.SpaceAfter = 0
                    If Left(Right(Selection.Text, 2), 1) = "��" Then
                        '��β��ð�ţ�����̧ͷ
                        With Selection.ParagraphFormat
                            .Alignment = wdAlignParagraphLeft
                            .CharacterUnitFirstLineIndent = 0
                            .FirstLineIndent = CentimetersToPoints(0)
                        End With
                    End If
                End If
                Exit For
            End If
        End With
    Next i
End Sub
'�����ʽ������+����ð�Ż��߸���+���ֱ��+����ð�ţ��Ҷ����ɶ�
'�����ʽ���¿�һ�У�������ַ���*������ƺ�ı�����
Private Sub ����˵����λ����()
    Selection.HomeKey unit:=wdStory
    With Selection.Find
        .Text = "������"
        .Forward = True
        .Wrap = wdFindStop
        If .Execute Then
            Selection.InsertBefore (vbCrLf)
        End If
    End With
    
    Selection.HomeKey unit:=wdStory
    With Selection.Find
        .Text = "����^#��"
        .Forward = True
        .Wrap = wdFindStop
        If .Execute Then
            Selection.InsertBefore (vbCrLf)
        End If
    End With
    Selection.EndKey unit:=wdStory
End Sub

'�����ʽ�����Ļ��ض�������ڸ���˵�������·�
'�����ʽ���¿����У��ҿ����ַ�
Private Sub ���Ļ��ظ�ʽ����()
    '�Ӻ���ǰ��"����+���ֱ�ţ�"��ͷ�ĵ�һ������
    hasAttachment = False
    Selection.EndKey unit:=wdStory
    With Selection.Find
        .Text = "����^#��"
        .Forward = False
        .Wrap = wdFindStop
        If .Execute Then
            hasAttachment = True
        End If
    End With
    'ǰ�涨λ���ˣ����治�ٶ�λ
    If hasAttachment = False Then
        '��ǰ�����"������"��ͷ�ĵ�һ������
        Selection.HomeKey unit:=wdStory
        With Selection.Find
            .Text = "������"
            .Forward = True
            .Wrap = wdFindStop
            If .Execute Then
                hasAttachment = True
            End If
        End With
    End If
    '���û�и���˵������λ���ĵ�����趨Ϊ����û�п���
    If hasAttachment = False Then
        Selection.EndKey unit:=wdStory
        Selection.MoveUp unit:=wdLine, Count:=2
    End If
    '�����ʽ
    Selection.Expand unit:=wdParagraph
    Selection.InsertAfter (vbCrLf)
    Selection.InsertAfter (vbCrLf)
    ���Ļ��ظ�ʽ
End Sub

Private Sub ���Ļ��ظ�ʽ()
    Selection.MoveDown
    Selection.HomeKey
    Selection.EndKey Extend:=wdExtend
    Selection.MoveDown unit:=wdLine, Extend:=wdExtend
    Selection.EndKey Extend:=wdExtend
    With Selection.ParagraphFormat
        .Alignment = wdAlignParagraphRight
        .CharacterUnitRightIndent = 5.5
        '.CharacterUnitFirstLineIndent = 0
        '.FirstLineIndent = CentimetersToPoints(0)
    End With
End Sub

'����ֶ��ֶη��������ʽ�ַ���Trim�ո�
Private Sub ����ո�()
    
    pn = ActiveDocument.Paragraphs.Count
    For i = 1 To pn
        If ActiveDocument.Paragraphs(i).Range.Information(wdWithInTable) = False And ActiveDocument.Paragraphs(i).Range.InlineShapes.Count = 0 And ActiveDocument.Paragraphs(i).Range.Find.Execute(FindText:="��^#��") = False Then
            ActiveDocument.Paragraphs(i).Range.Text = Trim(ActiveDocument.Paragraphs(i).Range.Text)
        End If
    Next i
    With ActiveDocument.Range.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Wrap = wdFindStop
        '�ֶ�����
        .Text = "^l"
        .Replacement.Text = "^p"
        .Execute Replace:=wdReplaceAll
        '�����ַ�������Ͽո�
        .Text = "^w^p"
        .Replacement.Text = "^p"
        .Execute Replace:=wdReplaceAll
        '�����ַ�������Ͽո�
        .Text = "^p^w"
        .Replacement.Text = "^p"
        .Execute Replace:=wdReplaceAll
    End With
    '    For i = 1 To pn
    '    With ActiveDocument.Paragraphs(i).Range
    '        If .Information(wdWithInTable) = False And .InlineShapes.Count = 0 And .Find.Execute(FindText:="��^#��") = False And Left(.Text, 1) <> "ͼ" Then
    '        With .Find
    '        '�����ַ�������Ͽո�
    '        .Text = "^w"
    '        .Replacement.Text = ""
    '        .Execute Replace:=wdReplaceAll
    '        End With
    '        End If
    '    End With
    '    Next i
    
End Sub

Private Sub ����ʽ()
    ' �����ĵ��еı��Ӧ�ú꣺���B
    
    For j = 1 To ActiveDocument.Tables.Count
        ActiveDocument.Tables(j).Select
        ���B
    Next j
End Sub

'�����ʽ��ͼƬΪǶ��ʽ
Private Sub ͼƬ����()
    Dim oShape As InlineShape
    For Each oShape In ActiveDocument.InlineShapes
        oShape.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0
        oShape.Range.ParagraphFormat.FirstLineIndent = CentimetersToPoints(0)
        oShape.Range.Paragraphs.Alignment = wdAlignParagraphCenter
    Next
End Sub

Private Sub ��ע()
    Selection.WholeStory
    With Selection.Find
        .ClearFormatting
        .MatchWholeWord = False
        .Execute FindText:="��ע��"
        If .Found = True Then
            Selection.Expand unit:=wdParagraph
            Debug.Print Selection.Range.Text
            With Selection.Font
                .NameFarEast = "����"
                .NameAscii = "����"
                .NameOther = "Times New Roman"
                .name = "����"
                .Size = 12
                .Bold = False
            End With
        End If
    End With
End Sub

Private Sub ��Ҫ()
    Selection.WholeStory
    With Selection.Find
        .ClearFormatting
        .MatchWholeWord = False
        .Execute FindText:="������Ҫ��"
        If .Found = True Then
            Selection.Expand unit:=wdParagraph
            Debug.Print Selection.Range.Text
            With Selection.Font
                .NameFarEast = "����"
                .NameAscii = "����"
                .NameOther = "Times New Roman"
                .name = "����"
                .Size = 14
                .Bold = False
            End With
        End If
    End With
End Sub

Private Sub �������()
    
    For Each i In ActiveDocument.Paragraphs
        If Len(Trim(i.Range.Text)) = 1 Then
            i.Range.Delete
        End If
        
    Next
End Sub

Private Sub ����ҳ��()
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
Private Sub ����ҳ��()
    '
    ' �淶����ҳ�룬��ż�ֿ�
    '
    '
    Application.ScreenUpdating = False
    With ActiveDocument.Sections(1)
        .PageSetup.OddAndEvenPagesHeaderFooter = True
        With .Footers(wdHeaderFooterPrimary)
            With .PageNumbers
                .Add PageNumberAlignment:=wdAlignPageNumberRight
                .NumberStyle = wdPageNumberStyleNumberInDash
            End With
            With .Range.Frames(1)
                .HorizontalPosition = wdFrameRight
                With .Range.ParagraphFormat
                    .Alignment = wdAlignParagraphRight
                    .CharacterUnitRightIndent = 1
                End With
            End With
        End With
        With .Footers(wdHeaderFooterEvenPages).Range.Frames(1)
            .HorizontalPosition = wdFrameLeft
            With .Range.ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .CharacterUnitLeftIndent = 1
            End With
        End With
    End With
    Application.ScreenUpdating = True
    Selection.HomeKey unit:=wdStory
    Selection.GoTo wdGoToPage, wdGoToNext, , "15 "
    If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
    End If
    If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow.ActivePane.View.Type = wdOutlineView Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    If Selection.HeaderFooter.IsHeader = True Then
        ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
    Else
        ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    End If
    Selection.WholeStory
    Selection.Font.name = "����"
    Selection.Font.Size = 14
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    Selection.HomeKey unit:=wdStory
    If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
    End If
    If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow.ActivePane.View.Type = wdOutlineView Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    If Selection.HeaderFooter.IsHeader = True Then
        ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
    Else
        ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    End If
    Selection.WholeStory
    Selection.Font.name = "����"
    Selection.Font.Size = 14
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
End Sub
Public Sub ���������Ű�()
    '
    ' Macro1 Macro
    ' ���� 2013 �� 11�� 14 �� �� user ¼��
    '
    If Application.Documents.Count = 0 Then
        Exit Sub
    End If
    
    Application.Run "�����Ű�"
    ActiveDocument.Save
    ActiveWindow.Close
    Application.Run MacroName:="���������Ű�"
End Sub
Public Sub ���������Ű�()
    '
    ' Macro1 Macro
    ' ���� 2013 �� 11�� 14 �� �� user ¼��
    '
    If Application.Documents.Count = 0 Then
        Exit Sub
    End If
    
    Application.Run "�����Ű�"
    ActiveDocument.Save
    ActiveWindow.Close
    Application.Run MacroName:="���������Ű�"
End Sub

'�����ʽ���ѽ��й�����ո񡢿��У���"��һ��"��ͷ�������ɶΣ���β�޾��
Private Sub ��������(pbtype As Integer)
    Application.ScreenUpdating = False '�ر���Ļ����
    nums = Array("һ", "��", "��", "��", "��", "��", "��", "��", "��", "ʮ", "ʮһ", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "��ʮ")
    
    pn = ActiveDocument.Paragraphs.Count
    Dim prg As Paragraph
    For i = 1 To ActiveDocument.Paragraphs.Count
        If i > ActiveDocument.Paragraphs.Count Then
            Exit For
        End If
        '��ȡ��һ�εĿ�ͷ
        Dim nextPreStr As String
        If i + 1 <= ActiveDocument.Paragraphs.Count Then
            nextPreStr = Left(ActiveDocument.Paragraphs(i + 1), 2)
        End If
        
        Set prg = ActiveDocument.Paragraphs(i)
        For j = 0 To UBound(nums)
            numtxt = nums(j)
            numlen = Len(numtxt) + 2
            Debug.Print Left(prg.Range.Text, numlen)
            '��������ƥ��ɹ�
            If Left(prg.Range.Text, numlen) = "��" & numtxt & "��" Then
                Debug.Print numtxt
                '�����βû�о�ţ���ʽ���������Ӿ�ţ��ϲ��������
                If Left(Right(prg.Range.Text, 2), 1) <> "��" Then
                    '��ʽ������
                    formatRng prg.Range, pbtype
                    addJuHao prg.Range
                    combineNext prg.Range, nextPreStr
                Else
                    '�����β�Ǿ�ţ��Ҳ�ֹһ������ȡ����һ��Ȼ���ʽ��
                    If countStr(prg.Range.Text, "��") > 1 Then
                        Dim trng As Range
                        Set trng = prg.Range
                        trng.Find.Execute ("��")
                        trng.SetRange prg.Range.start, trng.End
                        '��ʽ����һ��
                        formatRng trng, pbtype
                        '�����β�Ǿ��������ֻ��һ������ʽ�����κ󣬺ϲ��������
                    Else
                        If countStr(prg.Range.Text, "��") = 1 Then
                            '��ʽ������
                            formatRng prg.Range, pbtype
                            combineNext prg.Range, nextPreStr
                        End If
                    End If
                    
                End If
            End If
        Next j
    Next i
    
    Application.ScreenUpdating = True '�ָ���Ļ����
End Sub

Private Function countStr(srcStr As String, findStr As String) As Integer
    countStr = Len(srcStr) - Len(Replace(srcStr, findStr, ""))
End Function

Private Sub formatRng(rng As Range, pbtype As Integer)
    With rng.Font
    If pbtype <> 3 Then
        .NameFarEast = "����_GB2312"
        .NameAscii = "����_GB2312"
        .name = "����_GB2312"
    End If
        .Size = 16
        .Bold = False
    End With
    
End Sub
'��β�Ӿ��
Private Sub addJuHao(rng As Range)
    rng.MoveEnd wdWord, -1
    rng.InsertAfter ("��")
End Sub

'�ϲ���Σ����۶�β��û�о�ţ�������������������1.�����ϲ�����
Private Sub combineNext(rng As Range, nextPreStr As String)
    If nextPreStr = "1." Or nextPreStr = "1��" Or nextPreStr = "1��" Then
        Exit Sub
    End If
    
    With rng.Find
        .Text = "^p"
        .Replacement.Text = ""
        .Execute Replace:=wdReplaceAll
    End With
    
End Sub


Private Sub ���ı���()
    For i = 1 To ActiveDocument.Paragraphs.Count
        With ActiveDocument.Paragraphs(i).Range
            EndChar = Left(Right(.Text, 2), 1)
            If EndChar <> "��" And EndChar <> "��" Then
                Debug.Print EndChar
                '�������"��""��"�ȣ�Ϊ�ĺŴ���
                If EndChar = "��" And InStr(.Text, "��") > 0 Then
                    With .ParagraphFormat
                        .Alignment = wdAlignParagraphCenter
                        .FirstLineIndent = CentimetersToPoints(0)
                        .CharacterUnitFirstLineIndent = 0
                        .LineUnitAfter = 1
                    End With
                    ActiveDocument.Paragraphs(i - 1).Range.ParagraphFormat.LineUnitAfter = 0
                    ActiveDocument.Paragraphs(i - 1).Range.ParagraphFormat.SpaceAfter = 0
                Else
                    '��βû�б����ţ���Ϊ���⴦��
                    Dim isZBT
                    If i = 1 Then
                        '������
                        With .Font
                            .NameFarEast = "����С���μ���"
                            .NameAscii = "����С���μ���"
                            .NameOther = "Times New Roman"
                            .name = "����С���μ���"
                            .Size = 22
                            .Bold = False
                        End With
                        With .ParagraphFormat
                            .LineSpacingRule = wdLineSpaceExactly
                            .LineSpacing = 33
                            .Alignment = wdAlignParagraphCenter
                            .LineUnitBefore = 1
                            .LineUnitAfter = 1
                            .CharacterUnitFirstLineIndent = 0
                            .FirstLineIndent = CentimetersToPoints(0)
                        End With
                    Else
                        If Left(.Text, 2) = "һ��" Then
                            Exit Sub
                        End If
                        '������
                        With .Font
                            .NameFarEast = "����_GB2312"
                            .NameAscii = "����_GB2312"
                            .NameOther = "Times New Roman"
                            .name = "����_GB2312"
                            .Size = 16
                            .Bold = False
                        End With
                        With .ParagraphFormat
                            .LineSpacingRule = wdLineSpaceExactly
                            .LineSpacing = 33
                            .Alignment = wdAlignParagraphCenter
                            .LineUnitBefore = 0
                            .LineUnitAfter = 1
                            .CharacterUnitFirstLineIndent = 0
                            .FirstLineIndent = CentimetersToPoints(0)
                        End With
                        ActiveDocument.Paragraphs(i - 1).Range.ParagraphFormat.LineUnitAfter = 0
                        ActiveDocument.Paragraphs(i - 1).Range.ParagraphFormat.SpaceAfter = 0
                    End If
                End If
            Else
                If Left(Right(.Text, 2), 1) = "��" Then
                    '��β��ð�ţ�����̧ͷ
                    With .ParagraphFormat
                        .Alignment = wdAlignParagraphLeft
                        .CharacterUnitFirstLineIndent = 0
                        .FirstLineIndent = CentimetersToPoints(0)
                    End With
                End If
                '�Ǳ��⣬��̧ͷ������ϣ�����ɨ��ѭ��
                Exit For
            End If
        End With
    Next i
    
End Sub

Private Sub ͼע0()
    Selection.Expand unit:=wdParagraph
    Debug.Print Selection.Range.Text
    With Selection.Font
        .NameFarEast = "����"
        .NameAscii = "����"
        .NameOther = "Times New Roman"
        .name = "����"
        .Size = 14
        .Bold = False
    End With
    With Selection.ParagraphFormat
        .FirstLineIndent = CentimetersToPoints(0)
        .CharacterUnitFirstLineIndent = 0
    End With
    Selection.Range.Text = LTrim(Selection.Range.Text)
    Selection.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
End Sub

Private Sub ��ע()
    Selection.HomeKey unit:=wdStory
    With Selection.Find
        .Wrap = wdFindStop
        .Forward = True
        .ClearFormatting
        .MatchWholeWord = False
        Do While .Execute(FindText:="��^#��")
            If .Found = True Then
                Selection.Expand unit:=wdParagraph
                Debug.Print Selection.Range.Text
                With Selection.Font
                    .NameFarEast = "����"
                    .NameAscii = "����"
                    .NameOther = "Times New Roman"
                    .name = "����"
                    .Size = 14
                    .Bold = False
                End With
                With Selection.ParagraphFormat
                    .FirstLineIndent = CentimetersToPoints(0)
                    .CharacterUnitFirstLineIndent = 0
                End With
                Selection.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                Selection.MoveRight
            End If
        Loop
    End With
End Sub

'�����ʽ��ͼһ��ͼ1�������пո�
Private Sub ͼƬ����()
    Dim A As Variant
    A = Array("һ", "��", "��", "��", "��", "��", "��", "��", "��", "ʮ", "1", "2", "3", "4", "5", "6", "7", "8", "9")
    For j = 0 To UBound(A)
        Selection.WholeStory
        
        With Selection.Find
            .ClearFormatting
            .MatchWholeWord = False
            .Execute FindText:="ͼ" & A(j) & " "
            If .Found = True Then
                ͼע0
            End If
        End With
        
        With Selection.Find
            .ClearFormatting
            .MatchWholeWord = False
            .Execute FindText:="ͼ" & A(j) & "��"
            If .Found = True Then
                ͼע0
            End If
        End With
    Next j
    
    Selection.HomeKey unit:=wdStory
    
End Sub

Private Sub һ������()
    nums = Array("һ", "��", "��", "��", "��", "��", "��", "��", "��", "ʮ", "ʮһ", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "��ʮ")
    
    pn = ActiveDocument.Paragraphs.Count
    Dim prg As Paragraph
    For i = 1 To ActiveDocument.Paragraphs.Count
        If i > ActiveDocument.Paragraphs.Count Then
            Exit For
        End If
        
        Set prg = ActiveDocument.Paragraphs(i)
        For j = 0 To UBound(nums)
            numtxt = nums(j) & "��"
            numlen = Len(numtxt)
            Debug.Print Left(prg.Range.Text, numlen)
            If Left(prg.Range.Text, numlen) = numtxt Then
                If Left(Right(prg.Range.Text, 2), 1) = "��" And countStr(prg.Range.Text, "��") > 1 Then
                    Dim trng As Range
                    Set trng = prg.Range
                    trng.Find.Execute ("��")
                    trng.SetRange prg.Range.start, trng.End
                    trng.InsertAfter vbCrLf
                    oneTitleFormat trng
                Else
                    oneTitleFormat prg.Range
                End If
                Debug.Print numtxt
            End If
        Next j
    Next i
End Sub

Sub oneTitleFormat(rng As Range)
    With rng.Font
        .NameFarEast = "����"
        .NameAscii = "����"
        .NameOther = "Times New Roman"
        .name = "����"
        .Size = 16
        .Bold = False
    End With
    With rng.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 28
        .Alignment = wdAlignParagraphJustify
        .CharacterUnitFirstLineIndent = 2
    End With
    
End Sub

Private Sub ��������()
    For i = 1 To ActiveDocument.Paragraphs.Count
        If ActiveDocument.Paragraphs(i).Range.InlineShapes.Count = 0 Then
            With ActiveDocument.Paragraphs(i).Range.Font
                .NameFarEast = "����_GB2312"
                .NameAscii = "����_GB2312"
                .NameOther = "Times New Roman"
                .name = "����_GB2312"
                .Size = 16
                .ColorIndex = wdBlack
                .Bold = False
            End With
            With ActiveDocument.Paragraphs(i).Range.ParagraphFormat
                .LeftIndent = CentimetersToPoints(0)
                .RightIndent = CentimetersToPoints(0)
                .SpaceBefore = 0
                .SpaceBeforeAuto = False
                .SpaceAfter = 0
                .SpaceAfterAuto = False
                .LineSpacingRule = wdLineSpaceExactly
                .LineSpacing = 28
                .Alignment = wdAlignParagraphJustify
                If ActiveDocument.Paragraphs(i).Range.Information(wdWithInTable) = False Then
                    .CharacterUnitFirstLineIndent = 2
                End If
            End With
        End If
    Next
    
End Sub


Private Sub ���B()
    
    On Error Resume Next
    Application.ScreenUpdating = False '�ر���Ļ����
    If Selection.Information(wdWithInTable) = True Then
        Selection.Tables(1).Select
        Selection.Tables(1).PreferredWidthType = wdPreferredWidthPoints
        Selection.Tables(1).PreferredWidth = CentimetersToPoints(14.7)
        Selection.Tables(1).Rows.LeftIndent = CentimetersToPoints(0)
        Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
        Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed)
        
        
        Selection.Tables(1).Rows.Alignment = wdAlignRowLeft
        With Selection.Tables(1)
            .TopPadding = CentimetersToPoints(0)
            .BottomPadding = CentimetersToPoints(0)
            .LeftPadding = CentimetersToPoints(0)
            .RightPadding = CentimetersToPoints(0)
            .Spacing = 0
            .AllowPageBreaks = True
            .AllowAutoFit = False
        End With
        Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
        Selection.Columns.PreferredWidthType = wdPreferredWidthAuto
        Selection.Columns.PreferredWidth = 0
        With Selection.Cells(1)
            .WordWrap = True
            .FitText = False
        End With
        With Selection.Font
            .NameFarEast = "����"
            .NameAscii = "Times New Roman"
            .NameOther = "Times New Roman"
            .name = ""
            .Size = 10.5
            .Bold = False
            .Italic = False
        End With
        With Selection.ParagraphFormat
            .LeftIndent = CentimetersToPoints(0.1)
            .RightIndent = CentimetersToPoints(0.1)
            .SpaceBefore = 0
            .SpaceBeforeAuto = False
            .SpaceAfter = 0
            .SpaceAfterAuto = False
            .LineSpacingRule = wdLineSpaceAtLeast
            .LineSpacing = 12
            .CharacterUnitLeftIndent = 0
            .CharacterUnitRightIndent = 0
            .LineUnitBefore = 0
            .LineUnitAfter = 0
            .WordWrap = True
        End With
        
        Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
        Selection.MoveLeft unit:=wdCharacter, Count:=1
        Selection.SelectColumn
        Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
        Selection.MoveLeft unit:=wdCharacter, Count:=1
        Selection.SelectRow
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Selection.MoveLeft unit:=wdCharacter, Count:=1
        Selection.Tables(1).Select
        Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
        Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone
        Selection.Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
        Selection.Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        Selection.Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        Selection.Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        With Selection.Borders(wdBorderVertical)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
        End With
        With Selection.Borders(wdBorderTop)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
        End With
        With Selection.Borders(wdBorderBottom)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
        End With
        Selection.MoveLeft unit:=wdCharacter, Count:=1
        Selection.SelectRow
        With Selection.Borders(wdBorderBottom)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
        End With
        Selection.MoveLeft unit:=wdCharacter, Count:=1
        Selection.Tables(1).Select
        Selection.Rows.HeightRule = wdRowHeightAtLeast
        Selection.Rows.Height = CentimetersToPoints(0.5)
        Selection.Tables(1).Rows.LeftIndent = CentimetersToPoints(0)
        Selection.MoveLeft unit:=wdCharacter, Count:=1
        
        'Application.Run MacroName:="Normal.NewMacros.tabletest"
        
        '�����ͷֻ��һ�У��򽫵�һ�еĸ߶�����Ϊ1����
        Selection.Tables(1).Cell(1, 1).Select
        Selection.MoveLeft unit:=wdCharacter, Count:=1
        Selection.MoveDown unit:=wdLine, Count:=1
        H1 = Selection.Information(wdStartOfRangeRowNumber)
        If H1 = 2 Then
            Selection.Tables(1).Cell(1, 1).Select
            Selection.SelectRow
            Selection.Rows.Height = CentimetersToPoints(1#)
        End If
        
        '���������Ϊ����
        Selection.MoveLeft unit:=wdCharacter, Count:=1
        Selection.Tables(1).Rows.Alignment = wdAlignRowCenter
        
        '�����ڵ������
        Selection.Tables(1).Select
        Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
        Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
        
        '�̶������п�
        Selection.Tables(1).Select
        Selection.MoveLeft unit:=wdCharacter, Count:=1
        Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed)
        Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed)
        
        tabletest
        A01_�����Ӵֱ���е��ض���
        A00_���ÿ�ж��뷽ʽ
        A00_����������Ҷ���
        
        Selection.Tables(1).Cell(1, 1).Select
        Selection.SelectRow
        With Selection.Borders(wdBorderBottom)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
        End With
        
        With Selection.Borders(wdBorderHorizontal)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
        End With
        
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Selection.MoveLeft unit:=wdCharacter, Count:=1
        Selection.Tables(1).Cell(1, 1).Select
        Selection.MoveLeft unit:=wdCharacter, Count:=1
        
    Else
        MsgBox "��ע�⡿����㲻�ڱ���У�" & Chr(13) & _
            "���������뽫�����ŵ��������ⵥԪ" & Chr(13) & _
            "�����������У� Ȼ����ִ�б��꣬лл��"
    End If
    Application.ScreenUpdating = True '�ָ���Ļ����
    
End Sub

Private Sub ����ҳ��0()
    'ͨ��¼�ƺ��޸�'
    Selection.Sections(1).Footers(1).PageNumbers.Add PageNumberAlignment:= _
        wdAlignPageNumberOutside, FirstPage:=True
    If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
    End If
    If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
        ActivePane.View.Type = wdOutlineView Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    If Selection.HeaderFooter.IsHeader = True Then
        ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
    Else
        ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    End If
    Selection.MoveLeft unit:=wdCharacter, Count:=2
    Selection.TypeText Text:="��"
    Selection.MoveLeft unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Copy
    Selection.MoveRight unit:=wdCharacter, Count:=2
    Selection.PasteAndFormat (wdPasteDefault)
    Selection.MoveLeft unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.Font.Size = 12
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
End Sub


Private Sub tabletest()
    Dim myrng As Range
    Dim MyTab As Table
    
    If Selection.Information(wdWithInTable) = True Then
        Set myrng = Selection.Tables(1).Range
        Set MyTab = Selection.Tables(1)
        MyTab.Cell(1, 1).Select
        Selection.SelectRow
        '   Selection.MoveDown Unit:=wdLine, Count:=2, Extend:=wdExtend
        '   Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        
        '   �ı�Ĭ�ϱ���ߵ����ã�����Ϊ150pt
        With Options
            .DefaultBorderLineStyle = wdLineStyleSingle
            .DefaultBorderLineWidth = wdLineWidth150pt
            .DefaultBorderColor = wdColorBlack
        End With
        
        MyTab.Select
        
        '�ı�����
        With Selection.Borders(wdBorderTop)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
        End With
        
        '�ı������
        With Selection.Borders(wdBorderBottom)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
        End With
        
        '   �ı�Ĭ�ϱ���ߵ����ã��Ļ�ԭ����Ĭ��ֵ��025pt��
        With Options
            .DefaultBorderLineStyle = wdLineStyleSingle
            .DefaultBorderLineWidth = wdLineWidth025pt
            .DefaultBorderColor = wdColorBlack
        End With
        Selection.MoveLeft unit:=wdCharacter, Count:=1
        
    Else
        MsgBox "��ע�⡿����㲻�ڱ���У�" & Chr(13) & _
            "�뽫�����ŵ��������ⵥԪ���У� " & Chr(13) & _
            "Ȼ����ִ�б��꣬лл��"
    End If
    
End Sub



Private Sub A01_�����Ӵֱ���е��ض���()
    '����д��С�һ���������������������������мӴ֣����Ӵ�������ʮ��
    Dim A As Variant
    A = Array("һ", "��", "��", "��", "��", "��", "��", "��", "��", "ʮ" _
        , "ʮһ", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "��ʮ")
    Selection.Tables(1).Select
    Selection.MoveLeft unit:=wdCharacter, Count:=1
    Selection.SelectColumn
    Set MyRange = Selection.Range
    
    For j = 0 To UBound(A)
        With Selection.Find
            .ClearFormatting
            .Execute FindText:=A(j) & "��"
            If .Found = True Then
                .Parent.Expand unit:=wdParagraph
            Else
                Exit For
            End If
        End With
        Selection.SelectRow
        Selection.Range.Font.Bold = True
        Selection.MoveLeft unit:=wdCharacter, Count:=1
        MyRange.Select
    Next j
    
    Selection.Tables(1).Select
    Selection.MoveLeft unit:=wdCharacter, Count:=1
    Selection.Tables(1).Cell(1, 2).Select
    Selection.SelectColumn
    Set MyRange = Selection.Range
    
    For j = 0 To UBound(A)
        With Selection.Find
            .ClearFormatting
            .Execute FindText:=A(j) & "��"
            If .Found = True Then
                .Parent.Expand unit:=wdParagraph
            Else
                Exit For
            End If
        End With
        Selection.SelectRow
        Selection.Range.Font.Bold = True
        Selection.MoveLeft unit:=wdCharacter, Count:=1
        MyRange.Select
    Next j
End Sub

Private Sub A00_���ÿ�ж��뷽ʽ()
    
    '������������������У����ϴ��������������
    
    Dim MyTab As Table
    Dim LN() As Variant
    Dim H, L, M, N, H1, L1 As Integer
    Dim TT As String
    
    On Error Resume Next
    Application.ScreenUpdating = False '�ر���Ļ����
    If Selection.Information(wdWithInTable) = False Then
        MsgBox "��ѡ����"
        Exit Sub
    End If
    
    Set MyTab = Selection.Tables(1)
    MyTab.Cell(1, 1).Select
    Selection.MoveLeft unit:=wdCharacter, Count:=1
    Selection.MoveDown unit:=wdLine, Count:=1
    H1 = Selection.Information(wdStartOfRangeRowNumber)
    
    For C = 1 To MyTab.Columns.Count
        
        For j = 0 To MyTab.Rows.Count
            TT = MyTab.Cell(j + H1, C).Range.Text
            TT = Left(TT, Len(TT) - 2)
            L = Len(TT)
            ReDim Preserve LN(j - 1)
            LN(j - 2) = L
        Next j
        
        M = 1
        For i = 0 To UBound(LN)
            T = LN(i)
            If T > M Then M = T
        Next i
        
        N = LN(0)
        For i = 1 To UBound(LN) - 1
            T = LN(i)
            If T < N Then N = T
        Next i
        'MsgBox "��󳤶�Ϊ��" & M & "  ��С����Ϊ��" & N & "  ���" & M - N
        
        MyTab.Cell(H1, C).Select
        Selection.SelectColumn
        If M - N < 3 Then
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Else
            Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
        End If
    Next C
    
    MyTab.Cell(1, 1).Select
    Selection.MoveLeft unit:=wdCharacter, Count:=1
    Application.ScreenUpdating = True '�ָ���Ļ����
    
End Sub

Private Sub A00_����������Ҷ���()
    
    Dim MyTab As Table
    Dim L, H1 As Integer
    Dim TT As String
    
    On Error Resume Next
    Application.ScreenUpdating = False '�ر���Ļ����
    If Selection.Information(wdWithInTable) = False Then
        MsgBox "��ѡ����"
        Exit Sub
    End If
    
    Set MyTab = Selection.Tables(1)
    
    '��λ���Ŀ�ʼ��
    MyTab.Cell(1, 1).Select
    Selection.MoveLeft unit:=wdCharacter, Count:=1
    Selection.MoveDown unit:=wdLine, Count:=1
    H1 = Selection.Information(wdStartOfRangeRowNumber)
    
    For L = 2 To MyTab.Columns.Count
        '��λÿ�еķǿյ�Ԫ��
        Do While True
            TT = MyTab.Cell(H1, L).Range.Text
            TT = Left(TT, Len(TT) - 2)
            TT = Trim(TT)
            If Len(TT) > 0 Or TT = "" Then
                Exit Do
            Else
                H1 = H1 + 1
            End If
        Loop
        '���Ϊ�����У��������Ҷ���
        If Abs(Val(TT)) > 0 Then
            MyTab.Cell(H1, L).Select
            Selection.SelectColumn
            Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
        End If
    Next L
    MyTab.Cell(1, 1).Select
    Selection.MoveLeft unit:=wdCharacter, Count:=1
    Application.ScreenUpdating = True '�ָ���Ļ����
End Sub

