
Public Sub 分析排版()
    自动排版 2
End Sub

Public Sub 公文排版()
    自动排版 1
End Sub

Public Sub 通知排版()
    自动排版 3
End Sub
Private Sub 自动排版(pbtype As Integer)
    If Application.Documents.Count = 0 Then
        Exit Sub
    End If
    ActiveDocument.TrackRevisions = False '关闭修订
    Application.ScreenUpdating = False '关闭屏幕更新
    公文页面
    清除空格
    清除空行
    公文正文
    公文标题 '输入格式：标题位于第一行，且没有分段
    附件标题
    一级标题 '一级标题以"一、"为例，且位于行首
    二级标题 pbtype '二级标题以"（一）"为例，括号为中文括号，且位于行首
    图片标题
    图片居中
    表注
    If pbtype = 2 Then
        表格格式
    End If
    附注
    提要
    附件说明行位调整
    发文机关格式调整
    插入页码
    Application.ScreenUpdating = True '恢复屏幕更新
    Selection.HomeKey unit:=wdStory
End Sub

Private Sub 插入页码2()
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
'清除段落末尾
Private Function clearParagraphEnd(str As String) As String
    clearParagraphEnd = Replace(Replace(str, Chr(10), ""), Chr(13), "")
End Function
'输入格式：附件正文前有附件或附件+数字，无冒号，且独立成段。紧邻下一行为标题行
Private Sub 附件标题()
    '前插入分页符
    For i = 1 To ActiveDocument.Paragraphs.Count
        With ActiveDocument.Paragraphs(i).Range
            rgntxt = Replace(Replace(Replace(.Text, Chr(10), ""), Chr(13), ""), Chr(12), "")
            Dim isfj As Boolean
            isfj = False
            For j = 1 To 20
                If rgntxt = "附件" & j Then
                    Debug.Print "附件" & j
                    isfj = True
                    Exit For
                End If
            Next j
            If rgntxt = "附件" Or isfj = True Then
                .Select
                With Selection.Font
                    .NameFarEast = "黑体"
                    .NameAscii = "黑体"
                    .NameOther = "Times New Roman"
                    .name = "黑体"
                    .Size = 16
                    .Bold = False
                End With
                With Selection.ParagraphFormat
                    .Alignment = wdAlignParagraphLeft
                    .CharacterUnitLeftIndent = 0
                    .CharacterUnitFirstLineIndent = 0
                    .FirstLineIndent = CentimetersToPoints(0)
                End With
                '插入分页符
                If Left(Selection.Text, 1) <> Chr(12) Then
                    Selection.MoveLeft
                    Selection.InsertBreak Type:=wdPageBreak
                    Selection.Expand unit:=wdParagraph
                End If
                '附件行与标题行间添加空行
                Selection.InsertAfter vbCrLf
                '主标题格式
                Selection.MoveDown
                Selection.Expand unit:=wdParagraph
                With Selection.Font
                    .NameFarEast = "黑体"
                    .NameAscii = "黑体"
                    .NameOther = "Times New Roman"
                    .name = "黑体"
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
                '附件副标题
                Selection.MoveDown
                Selection.Expand unit:=wdParagraph
                If Left(Right(Selection.Text, 2), 1) <> "。" And Left(Right(Selection.Text, 2), 1) <> "：" And Left(Selection.Text, 2) <> "一、" Then
                    '副标题
                    With Selection.Font
                        .NameFarEast = "楷体_GB2312"
                        .NameAscii = "楷体_GB2312"
                        .NameOther = "Times New Roman"
                        .name = "楷体_GB2312"
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
                    '处理主标题段后间距为0
                    Selection.MoveLeft
                    Selection.MoveUp unit:=wdParagraph
                    Selection.Expand unit:=wdParagraph
                    Selection.ParagraphFormat.LineUnitAfter = 0
                    Selection.ParagraphFormat.SpaceAfter = 0
                    If Left(Right(Selection.Text, 2), 1) = "：" Then
                        '段尾有冒号，则是抬头
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
'输入格式：附件+中文冒号或者附件+数字编号+中文冒号，且独立成段
'输出格式：下空一行，左空两字符。*清除名称后的标点符号
Private Sub 附件说明行位调整()
    Selection.HomeKey unit:=wdStory
    With Selection.Find
        .Text = "附件："
        .Forward = True
        .Wrap = wdFindStop
        If .Execute Then
            Selection.InsertBefore (vbCrLf)
        End If
    End With
    
    Selection.HomeKey unit:=wdStory
    With Selection.Find
        .Text = "附件^#："
        .Forward = True
        .Wrap = wdFindStop
        If .Execute Then
            Selection.InsertBefore (vbCrLf)
        End If
    End With
    Selection.EndKey unit:=wdStory
End Sub

'输入格式：发文机关段落紧邻于附件说明段落下方
'输出格式：下空两行，右空四字符
Private Sub 发文机关格式调整()
    '从后往前查"附件+数字编号："开头的第一个段落
    hasAttachment = False
    Selection.EndKey unit:=wdStory
    With Selection.Find
        .Text = "附件^#："
        .Forward = False
        .Wrap = wdFindStop
        If .Execute Then
            hasAttachment = True
        End If
    End With
    '前面定位到了，后面不再定位
    If hasAttachment = False Then
        '从前往后查"附件："开头的第一个段落
        Selection.HomeKey unit:=wdStory
        With Selection.Find
            .Text = "附件："
            .Forward = True
            .Wrap = wdFindStop
            If .Execute Then
                hasAttachment = True
            End If
        End With
    End If
    '如果没有附件说明，定位到文档最后，设定为后面没有空行
    If hasAttachment = False Then
        Selection.EndKey unit:=wdStory
        Selection.MoveUp unit:=wdLine, Count:=2
    End If
    '处理格式
    Selection.Expand unit:=wdParagraph
    Selection.InsertAfter (vbCrLf)
    Selection.InsertAfter (vbCrLf)
    发文机关格式
End Sub

Private Sub 发文机关格式()
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

'清除手动分段符、特殊格式字符、Trim空格
Private Sub 清除空格()
    
    pn = ActiveDocument.Paragraphs.Count
    For i = 1 To pn
        If ActiveDocument.Paragraphs(i).Range.Information(wdWithInTable) = False And ActiveDocument.Paragraphs(i).Range.InlineShapes.Count = 0 And ActiveDocument.Paragraphs(i).Range.Find.Execute(FindText:="表^#：") = False Then
            ActiveDocument.Paragraphs(i).Range.Text = Trim(ActiveDocument.Paragraphs(i).Range.Text)
        End If
    Next i
    With ActiveDocument.Range.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Wrap = wdFindStop
        '手动换行
        .Text = "^l"
        .Replacement.Text = "^p"
        .Execute Replace:=wdReplaceAll
        '不明字符，不间断空格
        .Text = "^w^p"
        .Replacement.Text = "^p"
        .Execute Replace:=wdReplaceAll
        '不明字符，不间断空格
        .Text = "^p^w"
        .Replacement.Text = "^p"
        .Execute Replace:=wdReplaceAll
    End With
    '    For i = 1 To pn
    '    With ActiveDocument.Paragraphs(i).Range
    '        If .Information(wdWithInTable) = False And .InlineShapes.Count = 0 And .Find.Execute(FindText:="表^#：") = False And Left(.Text, 1) <> "图" Then
    '        With .Find
    '        '不明字符，不间断空格
    '        .Text = "^w"
    '        .Replacement.Text = ""
    '        .Execute Replace:=wdReplaceAll
    '        End With
    '        End If
    '    End With
    '    Next i
    
End Sub

Private Sub 表格格式()
    ' 将新文档中的表格应用宏：表格B
    
    For j = 1 To ActiveDocument.Tables.Count
        ActiveDocument.Tables(j).Select
        表格B
    Next j
End Sub

'输入格式：图片为嵌入式
Private Sub 图片居中()
    Dim oShape As InlineShape
    For Each oShape In ActiveDocument.InlineShapes
        oShape.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0
        oShape.Range.ParagraphFormat.FirstLineIndent = CentimetersToPoints(0)
        oShape.Range.Paragraphs.Alignment = wdAlignParagraphCenter
    Next
End Sub

Private Sub 附注()
    Selection.WholeStory
    With Selection.Find
        .ClearFormatting
        .MatchWholeWord = False
        .Execute FindText:="附注："
        If .Found = True Then
            Selection.Expand unit:=wdParagraph
            Debug.Print Selection.Range.Text
            With Selection.Font
                .NameFarEast = "仿宋"
                .NameAscii = "仿宋"
                .NameOther = "Times New Roman"
                .name = "仿宋"
                .Size = 12
                .Bold = False
            End With
        End If
    End With
End Sub

Private Sub 提要()
    Selection.WholeStory
    With Selection.Find
        .ClearFormatting
        .MatchWholeWord = False
        .Execute FindText:="内容提要："
        If .Found = True Then
            Selection.Expand unit:=wdParagraph
            Debug.Print Selection.Range.Text
            With Selection.Font
                .NameFarEast = "仿宋"
                .NameAscii = "仿宋"
                .NameOther = "Times New Roman"
                .name = "仿宋"
                .Size = 14
                .Bold = False
            End With
        End If
    End With
End Sub

Private Sub 清除空行()
    
    For Each i In ActiveDocument.Paragraphs
        If Len(Trim(i.Range.Text)) = 1 Then
            i.Range.Delete
        End If
        
    Next
End Sub

Private Sub 公文页面()
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
Private Sub 插入页码()
    '
    ' 规范公文页码，奇偶分开
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
    Selection.Font.name = "宋体"
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
    Selection.Font.name = "宋体"
    Selection.Font.Size = 14
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
End Sub
Public Sub 批量公文排版()
    '
    ' Macro1 Macro
    ' 宏在 2013 年 11月 14 日 由 user 录制
    '
    If Application.Documents.Count = 0 Then
        Exit Sub
    End If
    
    Application.Run "公文排版"
    ActiveDocument.Save
    ActiveWindow.Close
    Application.Run MacroName:="批量公文排版"
End Sub
Public Sub 批量分析排版()
    '
    ' Macro1 Macro
    ' 宏在 2013 年 11月 14 日 由 user 录制
    '
    If Application.Documents.Count = 0 Then
        Exit Sub
    End If
    
    Application.Run "分析排版"
    ActiveDocument.Save
    ActiveWindow.Close
    Application.Run MacroName:="批量分析排版"
End Sub

'输入格式：已进行过清除空格、空行，以"（一）"开头，单独成段，结尾无句号
Private Sub 二级标题(pbtype As Integer)
    Application.ScreenUpdating = False '关闭屏幕更新
    nums = Array("一", "二", "三", "四", "五", "六", "七", "八", "九", "十", "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十")
    
    pn = ActiveDocument.Paragraphs.Count
    Dim prg As Paragraph
    For i = 1 To ActiveDocument.Paragraphs.Count
        If i > ActiveDocument.Paragraphs.Count Then
            Exit For
        End If
        '获取下一段的开头
        Dim nextPreStr As String
        If i + 1 <= ActiveDocument.Paragraphs.Count Then
            nextPreStr = Left(ActiveDocument.Paragraphs(i + 1), 2)
        End If
        
        Set prg = ActiveDocument.Paragraphs(i)
        For j = 0 To UBound(nums)
            numtxt = nums(j)
            numlen = Len(numtxt) + 2
            Debug.Print Left(prg.Range.Text, numlen)
            '二级标题匹配成功
            If Left(prg.Range.Text, numlen) = "（" & numtxt & "）" Then
                Debug.Print numtxt
                '如果段尾没有句号，格式化段落后添加句号，合并后面段落
                If Left(Right(prg.Range.Text, 2), 1) <> "。" Then
                    '格式化段落
                    formatRng prg.Range, pbtype
                    addJuHao prg.Range
                    combineNext prg.Range, nextPreStr
                Else
                    '如果段尾是句号，且不止一个，提取出第一句然后格式化
                    If countStr(prg.Range.Text, "。") > 1 Then
                        Dim trng As Range
                        Set trng = prg.Range
                        trng.Find.Execute ("。")
                        trng.SetRange prg.Range.start, trng.End
                        '格式化第一句
                        formatRng trng, pbtype
                        '如果段尾是句号且整段只有一个，格式化本段后，合并后面段落
                    Else
                        If countStr(prg.Range.Text, "。") = 1 Then
                            '格式化段落
                            formatRng prg.Range, pbtype
                            combineNext prg.Range, nextPreStr
                        End If
                    End If
                    
                End If
            End If
        Next j
    Next i
    
    Application.ScreenUpdating = True '恢复屏幕更新
End Sub

Private Function countStr(srcStr As String, findStr As String) As Integer
    countStr = Len(srcStr) - Len(Replace(srcStr, findStr, ""))
End Function

Private Sub formatRng(rng As Range, pbtype As Integer)
    With rng.Font
    If pbtype <> 3 Then
        .NameFarEast = "楷体_GB2312"
        .NameAscii = "楷体_GB2312"
        .name = "楷体_GB2312"
    End If
        .Size = 16
        .Bold = False
    End With
    
End Sub
'段尾加句号
Private Sub addJuHao(rng As Range)
    rng.MoveEnd wdWord, -1
    rng.InsertAfter ("。")
End Sub

'合并后段，不论段尾有没有句号，如果后面是三级标题的1.，不合并段落
Private Sub combineNext(rng As Range, nextPreStr As String)
    If nextPreStr = "1." Or nextPreStr = "1、" Or nextPreStr = "1．" Then
        Exit Sub
    End If
    
    With rng.Find
        .Text = "^p"
        .Replacement.Text = ""
        .Execute Replace:=wdReplaceAll
    End With
    
End Sub


Private Sub 公文标题()
    For i = 1 To ActiveDocument.Paragraphs.Count
        With ActiveDocument.Paragraphs(i).Range
            EndChar = Left(Right(.Text, 2), 1)
            If EndChar <> "。" And EndChar <> "：" Then
                Debug.Print EndChar
                '如果含有"号""字"等，为文号处理
                If EndChar = "号" And InStr(.Text, "字") > 0 Then
                    With .ParagraphFormat
                        .Alignment = wdAlignParagraphCenter
                        .FirstLineIndent = CentimetersToPoints(0)
                        .CharacterUnitFirstLineIndent = 0
                        .LineUnitAfter = 1
                    End With
                    ActiveDocument.Paragraphs(i - 1).Range.ParagraphFormat.LineUnitAfter = 0
                    ActiveDocument.Paragraphs(i - 1).Range.ParagraphFormat.SpaceAfter = 0
                Else
                    '段尾没有标点符号，作为标题处理
                    Dim isZBT
                    If i = 1 Then
                        '主标题
                        With .Font
                            .NameFarEast = "方正小标宋简体"
                            .NameAscii = "方正小标宋简体"
                            .NameOther = "Times New Roman"
                            .name = "方正小标宋简体"
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
                        If Left(.Text, 2) = "一、" Then
                            Exit Sub
                        End If
                        '副标题
                        With .Font
                            .NameFarEast = "楷体_GB2312"
                            .NameAscii = "楷体_GB2312"
                            .NameOther = "Times New Roman"
                            .name = "楷体_GB2312"
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
                If Left(Right(.Text, 2), 1) = "：" Then
                    '段尾有冒号，则是抬头
                    With .ParagraphFormat
                        .Alignment = wdAlignParagraphLeft
                        .CharacterUnitFirstLineIndent = 0
                        .FirstLineIndent = CentimetersToPoints(0)
                    End With
                End If
                '非标题，且抬头处理完毕，跳出扫描循环
                Exit For
            End If
        End With
    Next i
    
End Sub

Private Sub 图注0()
    Selection.Expand unit:=wdParagraph
    Debug.Print Selection.Range.Text
    With Selection.Font
        .NameFarEast = "宋体"
        .NameAscii = "宋体"
        .NameOther = "Times New Roman"
        .name = "宋体"
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

Private Sub 表注()
    Selection.HomeKey unit:=wdStory
    With Selection.Find
        .Wrap = wdFindStop
        .Forward = True
        .ClearFormatting
        .MatchWholeWord = False
        Do While .Execute(FindText:="表^#：")
            If .Found = True Then
                Selection.Expand unit:=wdParagraph
                Debug.Print Selection.Range.Text
                With Selection.Font
                    .NameFarEast = "宋体"
                    .NameAscii = "宋体"
                    .NameOther = "Times New Roman"
                    .name = "宋体"
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

'输入格式：图一或图1并后面有空格
Private Sub 图片标题()
    Dim A As Variant
    A = Array("一", "二", "三", "四", "五", "六", "七", "八", "九", "十", "1", "2", "3", "4", "5", "6", "7", "8", "9")
    For j = 0 To UBound(A)
        Selection.WholeStory
        
        With Selection.Find
            .ClearFormatting
            .MatchWholeWord = False
            .Execute FindText:="图" & A(j) & " "
            If .Found = True Then
                图注0
            End If
        End With
        
        With Selection.Find
            .ClearFormatting
            .MatchWholeWord = False
            .Execute FindText:="图" & A(j) & "　"
            If .Found = True Then
                图注0
            End If
        End With
    Next j
    
    Selection.HomeKey unit:=wdStory
    
End Sub

Private Sub 一级标题()
    nums = Array("一", "二", "三", "四", "五", "六", "七", "八", "九", "十", "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十")
    
    pn = ActiveDocument.Paragraphs.Count
    Dim prg As Paragraph
    For i = 1 To ActiveDocument.Paragraphs.Count
        If i > ActiveDocument.Paragraphs.Count Then
            Exit For
        End If
        
        Set prg = ActiveDocument.Paragraphs(i)
        For j = 0 To UBound(nums)
            numtxt = nums(j) & "、"
            numlen = Len(numtxt)
            Debug.Print Left(prg.Range.Text, numlen)
            If Left(prg.Range.Text, numlen) = numtxt Then
                If Left(Right(prg.Range.Text, 2), 1) = "。" And countStr(prg.Range.Text, "。") > 1 Then
                    Dim trng As Range
                    Set trng = prg.Range
                    trng.Find.Execute ("。")
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
        .NameFarEast = "黑体"
        .NameAscii = "黑体"
        .NameOther = "Times New Roman"
        .name = "黑体"
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

Private Sub 公文正文()
    For i = 1 To ActiveDocument.Paragraphs.Count
        If ActiveDocument.Paragraphs(i).Range.InlineShapes.Count = 0 Then
            With ActiveDocument.Paragraphs(i).Range.Font
                .NameFarEast = "仿宋_GB2312"
                .NameAscii = "仿宋_GB2312"
                .NameOther = "Times New Roman"
                .name = "仿宋_GB2312"
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


Private Sub 表格B()
    
    On Error Resume Next
    Application.ScreenUpdating = False '关闭屏幕更新
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
            .NameFarEast = "宋体"
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
        
        '如果表头只有一行，则将第一行的高度设置为1厘米
        Selection.Tables(1).Cell(1, 1).Select
        Selection.MoveLeft unit:=wdCharacter, Count:=1
        Selection.MoveDown unit:=wdLine, Count:=1
        H1 = Selection.Information(wdStartOfRangeRowNumber)
        If H1 = 2 Then
            Selection.Tables(1).Cell(1, 1).Select
            Selection.SelectRow
            Selection.Rows.Height = CentimetersToPoints(1#)
        End If
        
        '将表格设置为居中
        Selection.MoveLeft unit:=wdCharacter, Count:=1
        Selection.Tables(1).Rows.Alignment = wdAlignRowCenter
        
        '按窗口调整表格
        Selection.Tables(1).Select
        Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
        Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
        
        '固定表格的列宽
        Selection.Tables(1).Select
        Selection.MoveLeft unit:=wdCharacter, Count:=1
        Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed)
        Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed)
        
        tabletest
        A01_批量加粗表格中的特定行
        A00_表格每列对齐方式
        A00_表格数字列右对齐
        
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
        MsgBox "【注意】插入点不在表格中！" & Chr(13) & _
            "　　　　请将插入点放到表格的任意单元" & Chr(13) & _
            "　　　　格中， 然后再执行本宏，谢谢！"
    End If
    Application.ScreenUpdating = True '恢复屏幕更新
    
End Sub

Private Sub 插入页码0()
    '通过录制宏修改'
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
    Selection.TypeText Text:="—"
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
        
        '   改变默认表格线的设置，设置为150pt
        With Options
            .DefaultBorderLineStyle = wdLineStyleSingle
            .DefaultBorderLineWidth = wdLineWidth150pt
            .DefaultBorderColor = wdColorBlack
        End With
        
        MyTab.Select
        
        '改变表格顶线
        With Selection.Borders(wdBorderTop)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
        End With
        
        '改变表格底线
        With Selection.Borders(wdBorderBottom)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
        End With
        
        '   改变默认表格线的设置，改回原来的默认值（025pt）
        With Options
            .DefaultBorderLineStyle = wdLineStyleSingle
            .DefaultBorderLineWidth = wdLineWidth025pt
            .DefaultBorderColor = wdColorBlack
        End With
        Selection.MoveLeft unit:=wdCharacter, Count:=1
        
    Else
        MsgBox "【注意】插入点不在表格中！" & Chr(13) & _
            "请将插入点放到表格的任意单元格中， " & Chr(13) & _
            "然后再执行本宏，谢谢！"
    End If
    
End Sub



Private Sub A01_批量加粗表格中的特定行()
    '表格中带有“一、”、“二、”、“三、”的行加粗，最多加粗至“二十”
    Dim A As Variant
    A = Array("一", "二", "三", "四", "五", "六", "七", "八", "九", "十" _
        , "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十")
    Selection.Tables(1).Select
    Selection.MoveLeft unit:=wdCharacter, Count:=1
    Selection.SelectColumn
    Set MyRange = Selection.Range
    
    For j = 0 To UBound(A)
        With Selection.Find
            .ClearFormatting
            .Execute FindText:=A(j) & "、"
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
            .Execute FindText:=A(j) & "、"
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

Private Sub A00_表格每列对齐方式()
    
    '文字栏：字数相近居中，相差较大居左；数字栏居右
    
    Dim MyTab As Table
    Dim LN() As Variant
    Dim H, L, M, N, H1, L1 As Integer
    Dim TT As String
    
    On Error Resume Next
    Application.ScreenUpdating = False '关闭屏幕更新
    If Selection.Information(wdWithInTable) = False Then
        MsgBox "请选择表格！"
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
        'MsgBox "最大长度为：" & M & "  最小长度为：" & N & "  相差" & M - N
        
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
    Application.ScreenUpdating = True '恢复屏幕更新
    
End Sub

Private Sub A00_表格数字列右对齐()
    
    Dim MyTab As Table
    Dim L, H1 As Integer
    Dim TT As String
    
    On Error Resume Next
    Application.ScreenUpdating = False '关闭屏幕更新
    If Selection.Information(wdWithInTable) = False Then
        MsgBox "请选择表格！"
        Exit Sub
    End If
    
    Set MyTab = Selection.Tables(1)
    
    '定位表文开始行
    MyTab.Cell(1, 1).Select
    Selection.MoveLeft unit:=wdCharacter, Count:=1
    Selection.MoveDown unit:=wdLine, Count:=1
    H1 = Selection.Information(wdStartOfRangeRowNumber)
    
    For L = 2 To MyTab.Columns.Count
        '定位每列的非空单元格
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
        '如果为数字列，则整列右对齐
        If Abs(Val(TT)) > 0 Then
            MyTab.Cell(H1, L).Select
            Selection.SelectColumn
            Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
        End If
    Next L
    MyTab.Cell(1, 1).Select
    Selection.MoveLeft unit:=wdCharacter, Count:=1
    Application.ScreenUpdating = True '恢复屏幕更新
End Sub

