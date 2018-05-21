Attribute VB_Name = "网页"
Sub 标准网页格式W()

    'st = VBA.Timer '程序运行计时器
    Application.ScreenUpdating = False '关闭屏幕更新
    
    On Error Resume Next
    
    Selection.HomeKey Unit:=wdStory
     A01_间隔号替换
    
    '将原文档全部内容转换为纯文本格式
    Selection.WholeStory
    Selection.Cut
    Selection.Collapse Direction:=wdCollapseStart
    
    CommandBars("Office Clipboard").Visible = False
    Selection.PasteSpecial Link:=False, DataType:=wdPasteText, Placement:=wdInLine, DisplayAsIcon:=False
    
    'Selection.Range.PasteSpecial DataType:=wdPasteText
    Selection.HomeKey Unit:=wdStory
    Selection.TypeParagraph
    
    Application.Run MacroName:="删段前空"      '应用宏命令：删段前空
    Application.Run MacroName:="删空行"        '应用宏命令：删空行
    Application.Run MacroName:="全角转换"      '应用宏命令：全角转换
    Application.Run MacroName:="段前加空"      '应用宏命令：段前加空
    
    '将文档格式变为小四宋，去掉缩进设置
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
    
    '将文档中的一级标题变为粗体

    Selection.WholeStory
    Dim A As Variant
        A = Array("一", "二", "三", "四", "五", "六", "七", "八", "九", "十" _
             , "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十")
    For j = 0 To 19
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
    Selection.HomeKey Unit:=wdStory
    ActiveDocument.Paragraphs(1).Range.Select
    Selection.Delete Unit:=wdCharacter, Count:=1
    
    '每个段落间加一空行
    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = "^p^p"
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
        Selection.WholeStory
        Selection.EndKey Unit:=wdStory
        Selection.Delete Unit:=wdCharacter, Count:=1
        Selection.TypeBackspace
        Selection.WholeStory
        Selection.Copy
        Selection.HomeKey Unit:=wdStory
        Selection.EndKey Unit:=wdStory
        Selection.TypeBackspace
        Selection.TypeBackspace
        Selection.HomeKey Unit:=wdStory
        A01_间隔号替换1
        Selection.HomeKey Unit:=wdStory
        
    Application.ScreenUpdating = True '恢复屏幕更新
    'MsgBox "程序运行用时" & VBA.Format(Timer - st, "0.00") & "秒" '显示程序运行的时间

    'CommandBars("Office Clipboard").Visible = True

End Sub
Sub 标准网页格式W1()
    
    'st = VBA.Timer '程序运行计时器
    Application.ScreenUpdating = False '关闭屏幕更新
    
    On Error Resume Next

    Selection.HomeKey Unit:=wdStory
    A01_间隔号替换

    Selection.WholeStory
    Selection.Cut
    Selection.Collapse Direction:=wdCollapseStart
    
    CommandBars("Office Clipboard").Visible = False
    Selection.PasteSpecial Link:=False, DataType:=wdPasteText, Placement:=wdInLine, DisplayAsIcon:=False
    
    'Selection.Range.PasteSpecial DataType:=wdPasteText
    
   '删除段落前和段落后的空格
    Dim MyRange As Range
    Selection.WholeStory
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

    
    Selection.HomeKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.WholeStory
        Selection.EndKey Unit:=wdStory
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

    Dim O As Variant
    Dim R As Variant
    O = Array("^l", "  ", "^p^p", "^p", "　　^p")
    R = Array("^p", " ", "^p", "^p　　", "")
    
    For i = 0 To 4
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = O(i)
        .Replacement.Text = R(i)
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i
    
    Selection.WholeStory

Dim A As Variant
A = Array("一", "二", "三", "四", "五", "六", "七", "八", "九", "十" _
          , "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十")
    
    For j = 0 To 19
    
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
    
    Application.Run MacroName:="全角转换"      '应用宏命令：全角转换

    Selection.WholeStory
    Selection.HomeKey Unit:=wdStory
    ActiveDocument.Paragraphs(1).Range.Select
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.WholeStory
    Selection.EndKey Unit:=wdStory
    Selection.TypeBackspace
    Selection.HomeKey Unit:=wdStory
    A01_间隔号替换1
    Selection.HomeKey Unit:=wdStory
        
    Application.ScreenUpdating = True '关闭屏幕更新
   ' MsgBox "程序运行用时" & VBA.Format(Timer - st, "0.00") & "秒" '显示程序运行的时间
    
   CommandBars("Office Clipboard").Visible = True
     
End Sub

