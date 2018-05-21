Attribute VB_Name = "A00"

Sub A00_图片拷贝至新文档()

    st = VBA.Timer '程序运行计时器

    ChangeFileOpenDirectory "F:\DOC-文图表"

    Application.ScreenUpdating = False '关闭屏幕更新
    On Error Resume Next '忽略错误

    '如果文档中有浮动式图片，则将其转换为嵌入式图片
    If ActiveDocument.Shapes.Count > 0 Then
        For Each oShape In ActiveDocument.Shapes
            oShape.ConvertToInlineShape
        Next
    End If

    'A01_不合并表头表注

    ' 定义三个变量：NM -- 文件名； CN -- 图片数  TN--表格数
    Dim NM As String
    Dim CN As Integer
    Dim TN As Integer
        
    ' 取得当前文件名
    Set MyDOC = Application.ActiveWindow.Document '指定要处理的文档为MyDoc
    NM = Left(MyDOC, Len(MyDOC) - 4)
    CN = ActiveDocument.InlineShapes.Count ' 取得文档中图片数
    TN = ActiveDocument.Tables.Count ' 取得文档中表格数

    ' 给原文档中的图片加上名称(TU1, TU2, ...)，以便拷贝回来
    MyDOC.Activate
    If CN > 0 Then
        For j = 1 To CN
            ActiveDocument.InlineShapes(j).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
            Selection.InsertAfter Text:="TU" & j
        Next j
    End If
    
    '-----------------------------------------------------
    ' 新建一个空白文档, 用于临时存储图片，指定为DOC_CN
    If CN > 0 Then
        Documents.Add DocumentType:=wdNewBlankDocument
        Set DOC_CN = Application.ActiveWindow.Document
        
        页面设置 '执行“页面设置”宏命令
        
        MyDOC.Activate
        For j = 1 To CN
            MyDOC.InlineShapes(j).Select
            Selection.Copy
            DOC_CN.Activate
            Selection.EndKey Unit:=wdStory
            Selection.TypeParagraph
            Selection.Paste
            MyDOC.Activate
        Next j
        
        ' 给新文档中的图片加上名称(TU1, TU2, ...)，以便拷贝回来
        DOC_CN.Activate
        For j = 1 To CN
            DOC_CN.InlineShapes(j).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
            Selection.InsertAfter Text:="TU" & j
        Next j
        
        DOC_CN.Activate
        ActiveDocument.SaveAs FileName:=NM & "_图片.doc", FileFormat:=wdFormatDocument
        
    End If
    
End Sub

Sub A00_标准网页格式OLD()

    ' 标准网页格式 Macro
    
    Dim O As Variant
    Dim R As Variant
    Dim A As Variant
    Dim C As Variant
    Dim D As Variant
    Dim CC As Variant
    Dim DD As Variant
    
    On Error Resume Next
    Application.ScreenUpdating = False '关闭屏幕更新
    
    'O = Array("^l", "  ", "^p^p", "^p", "　　^p", " ^p")
    'R = Array("^p", "　", "^p", "^p　　", "", "^p")
    
    O = Array("^l", Chr(32) & Chr(32), "^p^p", "^p", "　　^p", Chr(32) & Chr(13))
    R = Array("^p", Chr(-24159), "^p", "^p　　", "", Chr(13))
    A = Array("一", "二", "三", "四", "五", "六", "七", "八", "九", "十", _
          , "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十", _
          "二十一", "二十二", "二十三", "二十四", "二十五", "二十六", "二十七", "二十八", "二十九", "三十")
    C = Array("０", "１", "２", "３", "４", "５", "６", "７", "８", "９", ",", ";", "％", "?", ":", "(", ")")
    D = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "，", "；", "%", ".", "：", "（", "）")
    CC = Array("^l", ",", ";", "０", "１", "２", "３", "４", "５", "６", "７", "８", "９", "．", "Ａ", "Ｂ", "Ｃ", "Ｄ", "Ｅ", "Ｆ", "Ｇ", _
        "Ｈ", "Ｉ", "Ｊ", "Ｋ", "Ｌ", "Ｍ", "Ｎ", "Ｏ", "Ｐ", "Ｑ", "Ｒ", "Ｓ", "Ｔ", "Ｕ", "Ｖ", "Ｗ", "Ｘ", "Ｙ", "Ｚ", _
        "ａ", "ｂ", "ｃ", "ｄ", "ｅ", "ｆ", "ｇ", "ｈ", "ｉ", "ｊ", "ｋ", "ｌ", "ｍ", "ｎ", "ｏ", "ｐ", "ｑ", "ｒ", "ｓ", _
        "ｔ", "ｕ", "ｖ", "ｗ", "ｘ", "ｙ", "ｚ")
    DD = Array("^p", "，", "；", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", _
        "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "a", "b", "c", "d", "e", "f", "g", _
        "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z")
    
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
    
    '-------------------------------------------------
    '半角“:”替换为全角“：”
    Selection.HomeKey Unit:=wdStory '定位到文档开头
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = ":^p"
            .Replacement.Text = "：^p"
            .Wrap = wdFindStop
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
    '----------------------------------------------------
    Selection.HomeKey Unit:=wdStory '定位到文档开头
    
    '第一段如果为标题，则加粗居中
    Selection.Paragraphs(1).Range.Select
        s = Selection.Paragraphs(1).Range.Text
        If Len(s) < 30 Then
            Selection.Font.Bold = True
            Selection.Paragraphs(1).Range.Text = Trim(s)
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End If
        
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



Sub A00_文图表()

    '设定文档保存目录
    Dim FD1 As String
    Dim FD2 As String
    
    FD1 = "C:\Users\zlzx-dhg\Desktop\00 OK_DOC"
    FD2 = "D:\00 OK_DOC"
    st = VBA.Timer '程序运行计时器

    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FolderExists(FD1) = True Then
        ChangeFileOpenDirectory FD1
    Else
        If fs.FolderExists(FD2) = False Then
            Set A = fs.CreateFolder(FD2)
            ChangeFileOpenDirectory FD2
        Else
            ChangeFileOpenDirectory FD2
        End If
    End If

    Application.ScreenUpdating = False '关闭屏幕更新
    On Error Resume Next '忽略错误

    '如果文档中有浮动式图片，则将其转换为嵌入式图片
    If ActiveDocument.Shapes.Count > 0 Then
        For Each oShape In ActiveDocument.Shapes
            oShape.ConvertToInlineShape
        Next
    End If

    If ActiveDocument.Tables.Count > 0 Then
        A00_去表头表注
    End If
    Selection.HomeKey Unit:=wdStory
    A00_删除网页空格
    Selection.HomeKey Unit:=wdStory

    ' 定义三个变量：NM -- 文件名； CN -- 图片数  TN--表格数
    Dim NM As String
    Dim CN As Integer
    Dim TN As Integer
        
    ' 取得当前文件名
    Set MyDOC = Application.ActiveWindow.Document '指定要处理的文档为MyDoc
    If InStr(1, MyDOC, "Docx", 1) > 0 Then
        NM = Left(MyDOC, Len(MyDOC) - 5)
    Else
        NM = Left(MyDOC, Len(MyDOC) - 4)
    End If
    
    CN = ActiveDocument.InlineShapes.Count ' 取得文档中图片数
    TN = ActiveDocument.Tables.Count ' 取得文档中表格数

    ' 给原文档中的图片加上名称(TU1, TU2, ...)，以便拷贝回来
    MyDOC.Activate
    If CN > 0 Then
        For j = 1 To CN
            ActiveDocument.InlineShapes(j).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
            Selection.InsertAfter Text:="TU" & j
        Next j
    End If

    ' 给原文档中的表格加上名称(TAB1, TAB2, ...)，以便拷贝回来
    If TN > 0 Then
        For j = 1 To TN
            ActiveDocument.Tables(j).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
            Selection.InsertAfter Text:="TAB" & j
        Next j
    End If

    '-----------------------------------------------------
    ' 新建一个空白文档, 用于临时存储图片，指定为DOC_CN
    If CN > 0 Then
        Documents.Add DocumentType:=wdNewBlankDocument
        Set DOC_CN = Application.ActiveWindow.Document
        
        页面设置 '执行“页面设置”宏命令
        
        MyDOC.Activate
        For j = 1 To CN
            MyDOC.InlineShapes(j).Select
            Selection.Copy
            DOC_CN.Activate
            Selection.EndKey Unit:=wdStory
            Selection.TypeParagraph
            Selection.Paste
            MyDOC.Activate
        Next j
        
        ' 给新文档中的图片加上名称(TU1, TU2, ...)，以便拷贝回来
        DOC_CN.Activate
        For j = 1 To CN
            DOC_CN.InlineShapes(j).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
            Selection.InsertAfter Text:="TU" & j
        Next j
        
        DOC_CN.Activate
        'ActiveDocument.SaveAs FileName:=NM & "_图片.doc", FileFormat:=wdFormatDocument
        
    End If

    '-----------------------------------------------------
    ' 新建一个空白文档, 用于临时存储表格，指定为DOC_TN
    If TN > 0 Then
        Documents.Add DocumentType:=wdNewBlankDocument
        Set DOC_TN = Application.ActiveWindow.Document
        
        页面设置 '执行“页面设置”宏命令
        
        ' 把表格拷贝到新建文档中
        MyDOC.Activate
        For j = 1 To TN
            MyDOC.Tables(j).Range.Copy
            DOC_TN.Activate
            Selection.EndKey Unit:=wdStory
            Selection.TypeParagraph
            Selection.Paste
            MyDOC.Activate
        Next j
        
        ' 给新文档中的表格加上名称(TAB1, TAB2, ...)，以便拷贝回来
        DOC_TN.Activate
        For j = 1 To TN
            ActiveDocument.Tables(j).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
            Selection.InsertAfter Text:="TAB" & j
        Next j
        
        ' 将新文档中的表格应用宏：表格B
        DOC_TN.Activate
        For j = 1 To TN
            DOC_TN.Tables(j).Select
            表格B
        Next j
        
        DOC_TN.Activate
        'ActiveDocument.SaveAs FileName:=NM & "_表格.doc", FileFormat:=wdFormatDocument
        
    End If
    
    '-----------------------------------------------------

    ' 删除原文档中的表格
    MyDOC.Activate
    For j = 1 To TN
        If MyDOC.Tables.Count > 0 Then
            MyDOC.Tables(1).Select
            Selection.Cut
        End If
    Next j
    
    ' 删除原文档中的图片
    MyDOC.Activate
    For j = 1 To CN
        If MyDOC.InlineShapes.Count > 0 Then
            MyDOC.InlineShapes(1).Select
            Selection.Cut
        End If
    Next j
    
    MyDOC.Activate
    'ActiveDocument.SaveAs FileName:=NM & "_WZ.doc", FileFormat:=wdFormatDocument
    
    '-----------------------------------------------------
    
    '新建空白文档，把原文档中的文本拷贝至新文档，并执行“A00_网页格式”宏命令
    
    Documents.Add DocumentType:=wdNewBlankDocument
    Set DOC_OK = Application.ActiveWindow.Document
    
    页面设置 '执行“页面设置”宏命令
    
    MyDOC.Activate
    MyText = MyDOC.Content.Text
    DOC_OK.Activate
    DOC_OK.Content.Text = MyText
    'A00_国际经济信息 '执行“A00_国际经济信息”宏命令
    A00_网页格式  '执行“A00_网页格式”宏命令
    'A00_国际经济信息 '执行“A00_国际经济信息”宏命令

    '-----------------------------------------------------
    
    '把表格拷贝回来
    DOC_OK.Activate
    For j = 1 To TN
        With Selection.Find
            .ClearFormatting
            .Execute FindText:="TAB" & j
            If .Found = True Then
                .Parent.Expand Unit:=wdParagraph
            Else
                Exit For
            End If
        End With
        Selection.Delete Unit:=wdCharacter, Count:=1
        DOC_TN.Activate
            ActiveDocument.Tables(j).Select
            Selection.Copy
        DOC_OK.Activate
        Selection.Paste
    Next j
    
    DOC_OK.Activate
    Selection.HomeKey Unit:=wdStory '定位到文档开头
    
    '把图片拷贝回来
    For j = 1 To CN
        With Selection.Find
            .ClearFormatting
            .Execute FindText:="TU" & j
            If .Found = True Then
                .Parent.Expand Unit:=wdParagraph
            Else
                Exit For
            End If
        End With
        Selection.Delete Unit:=wdCharacter, Count:=1
        DOC_CN.Activate
            DOC_CN.InlineShapes(j).Select
            Selection.Copy
        DOC_OK.Activate

        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Selection.Paste
        Selection.MoveRight Unit:=wdCharacter, Count:=1
        Selection.TypeParagraph
    Next j
    
    DOC_OK.Activate
    Selection.HomeKey Unit:=wdStory
    
    空格替换
    A01_批量加粗表格合计行
    A00_批量加表头表注
    A00_批量加表注
    A00_附注
    A00_PMI

    '-----------------------------------------------------

   Application.ScreenUpdating = True '恢复屏幕更新


    DOC_OK.Activate

    '标准化后的文档另存,加了_OK
    ActiveDocument.SaveAs FileName:=NM & "_OK.doc", FileFormat:=wdFormatDocument
    ' ActiveDocument.SaveAs FileName:=NM & "_TB.doc", FileFormat:=wdFormatDocument

    Documents.Close SaveChanges:=wdDoNotSaveChanges
    Documents.Open (NM & "_OK.doc")

    'MsgBox "文档已经按要求标准化，本文档共有： " & Chr(13) & "    " & TN & " 张表格" & Chr(13) & "    " & CN & " 张图片" & Chr(13) & "程序运行用时" & VBA.Format(Timer - st, "0.00") & "秒"
    
End Sub

Sub A00_文图表N()

    '设定文档保存目录
    Dim FD1, FD2 As String
    
    FD1 = "C:\Users\zlzx-dhg\Desktop\00 OK_DOC"
    FD2 = "D:\00 OK_DOC"
    st = VBA.Timer '程序运行计时器

    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FolderExists(FD1) = True Then
        ChangeFileOpenDirectory FD1
    Else
        If fs.FolderExists(FD2) = False Then
            Set A = fs.CreateFolder(FD2)
            ChangeFileOpenDirectory FD2
        Else
            ChangeFileOpenDirectory FD2
        End If
    End If

    Application.ScreenUpdating = False '关闭屏幕更新
    On Error Resume Next '忽略错误

    '如果文档中有浮动式图片，则将其转换为嵌入式图片
    If ActiveDocument.Shapes.Count > 0 Then
        For Each oShape In ActiveDocument.Shapes
            oShape.ConvertToInlineShape
        Next
    End If

    If ActiveDocument.Tables.Count > 0 Then
        A00_去表头表注
    End If
    Selection.HomeKey Unit:=wdStory
    A00_删除网页空格
    Selection.HomeKey Unit:=wdStory

    ' 定义三个变量：NM -- 文件名； CN -- 图片数  TN--表格数
    Dim NM As String
    Dim CN As Integer
    Dim TN As Integer
        
    ' 取得当前文件名
    Set MyDOC = Application.ActiveWindow.Document '指定要处理的文档为MyDoc
    If InStr(1, MyDOC, "Docx", 1) > 0 Then
        NM = Left(MyDOC, Len(MyDOC) - 5)
    Else
        NM = Left(MyDOC, Len(MyDOC) - 4)
    End If
    
    CN = ActiveDocument.InlineShapes.Count ' 取得文档中图片数
    TN = ActiveDocument.Tables.Count ' 取得文档中表格数

    ' 给原文档中的图片加上名称(TU1, TU2, ...)，以便拷贝回来
    MyDOC.Activate
    If CN > 0 Then
        For j = 1 To CN
            ActiveDocument.InlineShapes(j).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
            Selection.InsertAfter Text:="TU" & j
        Next j
    End If

    ' 给原文档中的表格加上名称(TAB1, TAB2, ...)，以便拷贝回来
    If TN > 0 Then
        For j = 1 To TN
            ActiveDocument.Tables(j).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
            Selection.InsertAfter Text:="TAB" & j
        Next j
    End If

    '-----------------------------------------------------
    ' 新建一个空白文档, 用于临时存储图片，指定为DOC_CN
    If CN > 0 Then
        Documents.Add DocumentType:=wdNewBlankDocument
        Set DOC_CN = Application.ActiveWindow.Document
        
        页面设置 '执行“页面设置”宏命令
        
        MyDOC.Activate
        For j = 1 To CN
            MyDOC.InlineShapes(j).Select
            Selection.Copy
            DOC_CN.Activate
            Selection.EndKey Unit:=wdStory
            Selection.TypeParagraph
            Selection.Paste
            MyDOC.Activate
        Next j
        
        ' 给新文档中的图片加上名称(TU1, TU2, ...)，以便拷贝回来
        DOC_CN.Activate
        For j = 1 To CN
            DOC_CN.InlineShapes(j).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
            Selection.InsertAfter Text:="TU" & j
        Next j
        
        DOC_CN.Activate
        'ActiveDocument.SaveAs FileName:=NM & "_图片.doc", FileFormat:=wdFormatDocument
        
    End If

    '-----------------------------------------------------
    ' 新建一个空白文档, 用于临时存储表格，指定为DOC_TN
    If TN > 0 Then
        Documents.Add DocumentType:=wdNewBlankDocument
        Set DOC_TN = Application.ActiveWindow.Document
        
        页面设置 '执行“页面设置”宏命令
        
        ' 把表格拷贝到新建文档中
        MyDOC.Activate
        For j = 1 To TN
            MyDOC.Tables(j).Range.Copy
            DOC_TN.Activate
            Selection.EndKey Unit:=wdStory
            Selection.TypeParagraph
            Selection.Paste
            MyDOC.Activate
        Next j
        
        ' 给新文档中的表格加上名称(TAB1, TAB2, ...)，以便拷贝回来
        DOC_TN.Activate
        For j = 1 To TN
            ActiveDocument.Tables(j).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
            Selection.InsertAfter Text:="TAB" & j
        Next j
        
        ' 将新文档中的表格应用宏：表格B
        DOC_TN.Activate
        For j = 1 To TN
            DOC_TN.Tables(j).Select
            表格B
        Next j
        
        DOC_TN.Activate
        'ActiveDocument.SaveAs FileName:=NM & "_表格.doc", FileFormat:=wdFormatDocument
        
    End If
    
    '-----------------------------------------------------

    ' 删除原文档中的表格
    MyDOC.Activate
    For j = 1 To TN
        If MyDOC.Tables.Count > 0 Then
            MyDOC.Tables(1).Select
            Selection.Cut
        End If
    Next j
    
    ' 删除原文档中的图片
    MyDOC.Activate
    For j = 1 To CN
        If MyDOC.InlineShapes.Count > 0 Then
            MyDOC.InlineShapes(1).Select
            Selection.Cut
        End If
    Next j
    
    MyDOC.Activate
    'ActiveDocument.SaveAs FileName:=NM & "_WZ.doc", FileFormat:=wdFormatDocument
    
    '-----------------------------------------------------
    
    '新建空白文档，把原文档中的文本拷贝至新文档，并执行“A00_网页格式”宏命令
    
    Documents.Add DocumentType:=wdNewBlankDocument
    Set DOC_OK = Application.ActiveWindow.Document
    
    页面设置 '执行“页面设置”宏命令
    
    MyDOC.Activate
    MyText = MyDOC.Content.Text
    DOC_OK.Activate
    DOC_OK.Content.Text = MyText
    A00_网页格式  '执行“A00_网页格式”宏命令

    '-----------------------------------------------------
    
    '把表格拷贝回来
    DOC_OK.Activate
    For j = 1 To TN
        With Selection.Find
            .ClearFormatting
            .Execute FindText:="TAB" & j
            If .Found = True Then
                .Parent.Expand Unit:=wdParagraph
            Else
                Exit For
            End If
        End With
        Selection.Delete Unit:=wdCharacter, Count:=1
        DOC_TN.Activate
            ActiveDocument.Tables(j).Select
            Selection.Copy
        DOC_OK.Activate
        Selection.Paste
    Next j
    
    DOC_OK.Activate
    Selection.HomeKey Unit:=wdStory '定位到文档开头
    
    '把图片拷贝回来
    For j = 1 To CN
        With Selection.Find
            .ClearFormatting
            .Execute FindText:="TU" & j
            If .Found = True Then
                .Parent.Expand Unit:=wdParagraph
            Else
                Exit For
            End If
        End With
        Selection.Delete Unit:=wdCharacter, Count:=1
        DOC_CN.Activate
            DOC_CN.InlineShapes(j).Select
            Selection.Copy
        DOC_OK.Activate

        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Selection.Paste
        Selection.MoveRight Unit:=wdCharacter, Count:=1
        Selection.TypeParagraph
    Next j
    
    DOC_OK.Activate
    Selection.HomeKey Unit:=wdStory
    
    空格替换
    A01_批量加粗表格合计行
    A00_附注
    A00_PMI
    A00_表标题加粗居中
    A00_单位居右_批量
    A00_表注_批量
    A00_表标题_批量

    '-----------------------------------------------------

   Application.ScreenUpdating = True '恢复屏幕更新


    DOC_OK.Activate

    '标准化后的文档另存,加了_OK
    ActiveDocument.SaveAs FileName:=NM & "_OK.doc", FileFormat:=wdFormatDocument
    ' ActiveDocument.SaveAs FileName:=NM & "_TB.doc", FileFormat:=wdFormatDocument

    Documents.Close SaveChanges:=wdDoNotSaveChanges
    Documents.Open (NM & "_OK.doc")

    'MsgBox "文档已经按要求标准化，本文档共有： " & Chr(13) & "    " & TN & " 张表格" & Chr(13) & "    " & CN & " 张图片" & Chr(13) & "程序运行用时" & VBA.Format(Timer - st, "0.00") & "秒"
    
End Sub

Sub A00_文图表_信息公开()

    '设定文档保存目录
    Dim FD1 As String
    Dim FD2 As String
    
    FD1 = "C:\Users\zlzx-dhg\Desktop\00 OK_DOC"
    FD2 = "D:\00 OK_DOC"
    st = VBA.Timer '程序运行计时器

    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FolderExists(FD1) = True Then
        ChangeFileOpenDirectory FD1
    Else
        If fs.FolderExists(FD2) = False Then
            Set A = fs.CreateFolder(FD2)
            ChangeFileOpenDirectory FD2
        Else
            ChangeFileOpenDirectory FD2
        End If
    End If

    Application.ScreenUpdating = False '关闭屏幕更新
    On Error Resume Next '忽略错误

    '如果文档中有浮动式图片，则将其转换为嵌入式图片
    If ActiveDocument.Shapes.Count > 0 Then
        For Each oShape In ActiveDocument.Shapes
            oShape.ConvertToInlineShape
        Next
    End If

    If ActiveDocument.Tables.Count > 0 Then
        A00_去表头表注
    End If
    Selection.HomeKey Unit:=wdStory
    A00_删除网页空格
    Selection.HomeKey Unit:=wdStory

    ' 定义三个变量：NM -- 文件名； CN -- 图片数  TN--表格数
    Dim NM As String
    Dim CN As Integer
    Dim TN As Integer
        
    ' 取得当前文件名
    Set MyDOC = Application.ActiveWindow.Document '指定要处理的文档为MyDoc
    If InStr(1, MyDOC, "Docx", 1) > 0 Then
        NM = Left(MyDOC, Len(MyDOC) - 5)
    Else
        NM = Left(MyDOC, Len(MyDOC) - 4)
    End If
    
    CN = ActiveDocument.InlineShapes.Count ' 取得文档中图片数
    TN = ActiveDocument.Tables.Count ' 取得文档中表格数

    ' 给原文档中的图片加上名称(TU1, TU2, ...)，以便拷贝回来
    MyDOC.Activate
    If CN > 0 Then
        For j = 1 To CN
            ActiveDocument.InlineShapes(j).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
            Selection.InsertAfter Text:="TU" & j
        Next j
    End If

    ' 给原文档中的表格加上名称(TAB1, TAB2, ...)，以便拷贝回来
    If TN > 0 Then
        For j = 1 To TN
            ActiveDocument.Tables(j).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
            Selection.InsertAfter Text:="TAB" & j
        Next j
    End If

    '-----------------------------------------------------
    ' 新建一个空白文档, 用于临时存储图片，指定为DOC_CN
    If CN > 0 Then
        Documents.Add DocumentType:=wdNewBlankDocument
        Set DOC_CN = Application.ActiveWindow.Document
        
        页面设置 '执行“页面设置”宏命令
        
        MyDOC.Activate
        For j = 1 To CN
            MyDOC.InlineShapes(j).Select
            Selection.Copy
            DOC_CN.Activate
            Selection.EndKey Unit:=wdStory
            Selection.TypeParagraph
            Selection.Paste
            MyDOC.Activate
        Next j
        
        ' 给新文档中的图片加上名称(TU1, TU2, ...)，以便拷贝回来
        DOC_CN.Activate
        For j = 1 To CN
            DOC_CN.InlineShapes(j).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
            Selection.InsertAfter Text:="TU" & j
        Next j
        
        DOC_CN.Activate
        'ActiveDocument.SaveAs FileName:=NM & "_图片.doc", FileFormat:=wdFormatDocument
        
    End If

    '-----------------------------------------------------
    ' 新建一个空白文档, 用于临时存储表格，指定为DOC_TN
    If TN > 0 Then
        Documents.Add DocumentType:=wdNewBlankDocument
        Set DOC_TN = Application.ActiveWindow.Document
        
        页面设置 '执行“页面设置”宏命令
        
        ' 把表格拷贝到新建文档中
        MyDOC.Activate
        For j = 1 To TN
            MyDOC.Tables(j).Range.Copy
            DOC_TN.Activate
            Selection.EndKey Unit:=wdStory
            Selection.TypeParagraph
            Selection.Paste
            MyDOC.Activate
        Next j
        
        ' 给新文档中的表格加上名称(TAB1, TAB2, ...)，以便拷贝回来
        DOC_TN.Activate
        For j = 1 To TN
            ActiveDocument.Tables(j).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
            Selection.InsertAfter Text:="TAB" & j
        Next j
        
        ' 将新文档中的表格应用宏：表格B
        DOC_TN.Activate
        For j = 1 To TN
            DOC_TN.Tables(j).Select
            表格B
        Next j
        
        DOC_TN.Activate
        'ActiveDocument.SaveAs FileName:=NM & "_表格.doc", FileFormat:=wdFormatDocument
        
    End If
    
    '-----------------------------------------------------

    ' 删除原文档中的表格
    MyDOC.Activate
    For j = 1 To TN
        If MyDOC.Tables.Count > 0 Then
            MyDOC.Tables(1).Select
            Selection.Cut
        End If
    Next j
    
    ' 删除原文档中的图片
    MyDOC.Activate
    For j = 1 To CN
        If MyDOC.InlineShapes.Count > 0 Then
            MyDOC.InlineShapes(1).Select
            Selection.Cut
        End If
    Next j
    
    MyDOC.Activate
    'ActiveDocument.SaveAs FileName:=NM & "_WZ.doc", FileFormat:=wdFormatDocument
    
    '-----------------------------------------------------
    
    '新建空白文档，把原文档中的文本拷贝至新文档，并执行“A00_网页格式”宏命令
    
    Documents.Add DocumentType:=wdNewBlankDocument
    Set DOC_OK = Application.ActiveWindow.Document
    
    页面设置 '执行“页面设置”宏命令
    
    MyDOC.Activate
    MyText = MyDOC.Content.Text
    DOC_OK.Activate
    DOC_OK.Content.Text = MyText
    A00_网页格式  '执行“A00_网页格式”宏命令

    '-----------------------------------------------------
    
    '把表格拷贝回来
    DOC_OK.Activate
    For j = 1 To TN
        With Selection.Find
            .ClearFormatting
            .Execute FindText:="TAB" & j
            If .Found = True Then
                .Parent.Expand Unit:=wdParagraph
            Else
                Exit For
            End If
        End With
        Selection.Delete Unit:=wdCharacter, Count:=1
        DOC_TN.Activate
            ActiveDocument.Tables(j).Select
            Selection.Copy
        DOC_OK.Activate
        Selection.Paste
    Next j
    
    DOC_OK.Activate
    Selection.HomeKey Unit:=wdStory '定位到文档开头
    
    '把图片拷贝回来
    For j = 1 To CN
        With Selection.Find
            .ClearFormatting
            .Execute FindText:="TU" & j
            If .Found = True Then
                .Parent.Expand Unit:=wdParagraph
            Else
                Exit For
            End If
        End With
        Selection.Delete Unit:=wdCharacter, Count:=1
        DOC_CN.Activate
            DOC_CN.InlineShapes(j).Select
            Selection.Copy
        DOC_OK.Activate

        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Selection.Paste
        Selection.MoveRight Unit:=wdCharacter, Count:=1
        Selection.TypeParagraph
    Next j
    
    DOC_OK.Activate
    Selection.HomeKey Unit:=wdStory
    
    空格替换
    A01_批量加粗表格合计行
    A00_附注
    A00_PMI
    A00_表标题加粗居中
    A00_单位_批量_信息公开
    A00_表注_批量

    '-----------------------------------------------------

   Application.ScreenUpdating = True '恢复屏幕更新


    DOC_OK.Activate

    '标准化后的文档另存,加了_OK
    ActiveDocument.SaveAs FileName:=NM & "_OK.doc", FileFormat:=wdFormatDocument
    ' ActiveDocument.SaveAs FileName:=NM & "_TB.doc", FileFormat:=wdFormatDocument

    Documents.Close SaveChanges:=wdDoNotSaveChanges
    Documents.Open (NM & "_OK.doc")

    'MsgBox "文档已经按要求标准化，本文档共有： " & Chr(13) & "    " & TN & " 张表格" & Chr(13) & "    " & CN & " 张图片" & Chr(13) & "程序运行用时" & VBA.Format(Timer - st, "0.00") & "秒"
    
End Sub

Sub A00_表格()
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
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.SelectColumn
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.SelectRow
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
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
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.SelectRow
    With Selection.Borders(wdBorderBottom)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.Tables(1).Select
    Selection.Rows.HeightRule = wdRowHeightAtLeast
    Selection.Rows.Height = CentimetersToPoints(0.5)
    Selection.Tables(1).Rows.LeftIndent = CentimetersToPoints(0)
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    
    Application.Run MacroName:="Normal.NewMacros.tabletest"
    
    '将第一行的高度设置为1厘米
    Selection.SelectRow
    Selection.Rows.Height = CentimetersToPoints(1#)
    
    '将表格设置为居中
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.Tables(1).Rows.Alignment = wdAlignRowCenter
    
    '按窗口调整表格
    Selection.Tables(1).Select
    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    
    '固定表格的列宽
    Selection.Tables(1).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed)
    
    Selection.Tables(1).Cell(1, 1).Select
    Selection.SelectRow
        With Selection.Borders(wdBorderBottom)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
        End With
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.MoveLeft Unit:=wdCharacter, Count:=1

   Else
    MsgBox "亲，插入点不在表格中！不符合执行本宏命令的条件！" & Chr(13) & _
           "请将插入点放到表格的任意单元格中，" & Chr(13) & _
           "然后再执行本宏，谢谢！"
   End If
      Application.ScreenUpdating = True '恢复屏幕更新

End Sub

Sub A00_答记者问()
    Dim AA As Variant
    AA = Array("毛盛勇", "宁吉喆", "李克强", "邢志宏", "刘爱华", "有关负责人") ' 定义发言人数组
    
    '判断是否符合执行本宏命令的条件，正文中必须有“记者”
    Selection.HomeKey Unit:=wdStory '定位到文档开头
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="记者："
        If .Found = False Then
            MsgBox "亲，当前文档不符合执行本宏命令的条件！" & Chr(13) _
            & "你是点错键了吧，请点击“确定”退出吧！"
            Selection.HomeKey Unit:=wdStory '定位到文档开头
            Exit Sub
        End If
    End With
    
    On Error Resume Next
    Application.ScreenUpdating = False '关闭屏幕更新
    
   ' A00_网页格式
    
    For i = 0 To UBound(AA)
    
    Selection.HomeKey Unit:=wdStory '定位到文档开头
        For j = 1 To ActiveDocument.Paragraphs.Count
            With Selection.Find
                .ClearFormatting
                .Execute FindText:="　　" & AA(i) & "："
                If .Found = True Then
                   ' .Parent.Expand Unit:=wdParagraph
                    Selection.Range.Font.Bold = True
                End If
            End With
            Selection.MoveRight Unit:=wdCharacter, Count:=1
        Next j
    Next i
        
    Selection.HomeKey Unit:=wdStory '定位到文档开头
        For j = 1 To ActiveDocument.Paragraphs.Count
            With Selection.Find
                .ClearFormatting
                .Execute FindText:="记者："
                If .Found = True Then
                    Selection.MoveRight Unit:=wdCharacter, Count:=1
                    Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
                    Selection.Font.Bold = True
                    Selection.Paragraphs(1).Range.Select
                    Selection.Font.NameFarEast = "楷体"
                End If
            End With
            Selection.MoveRight Unit:=wdCharacter, Count:=1
        Next j
    Selection.HomeKey Unit:=wdStory '定位到文档开头
    Application.ScreenUpdating = True '恢复屏幕更新
End Sub

Sub A00_条文加粗()

    '查找特定的字符串，如："条"，然后将该字符串前的内容加粗
    
    Application.ScreenUpdating = False '关闭屏幕更新
    
    '将两个西文空格替换为一个中文空格
    Selection.HomeKey Unit:=wdStory '定位到文档开头
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = Chr(32) & Chr(32)
        .Replacement.Text = Chr(-24159)
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    '将单个西文空格替换为一个中文空格
    Selection.HomeKey Unit:=wdStory '定位到文档开头
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = Chr(32)
        .Replacement.Text = Chr(-24159)
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll


    '判断是否符合执行本宏命令的条件，正文中必须有“条”
    Selection.HomeKey Unit:=wdStory '定位到文档开头
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="条" & Chr(-24159)
        If .Found = False Then
            MsgBox "亲，当前文档不符合执行本宏命令的条件！" & Chr(13) _
            & "宏宏觉得，你是点错键了吧，请点击“确定”退出吧！"
            Selection.HomeKey Unit:=wdStory '定位到文档开头
            Exit Sub
        End If
    End With

    '对条文前的内容加粗
    Selection.WholeStory
    Selection.HomeKey Unit:=wdStory
    For j = 0 To ActiveDocument.Paragraphs.Count - 1
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="条" & Chr(-24159)
        If .Found = True Then
            Selection.MoveRight Unit:=wdCharacter, Count:=1
            Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
            Selection.Font.Bold = True
        End If
    End With
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Next j
    
    '对章进行处理，黑体+居中
    Selection.WholeStory
    Selection.HomeKey Unit:=wdStory
    For j = 0 To ActiveDocument.Paragraphs.Count - 1
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="章" & Chr(-24159)
        If .Found = True Then
            'Selection.MoveRight Unit:=wdCharacter, Count:=1
            'Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
            Selection.Paragraphs(1).Range.Select
            Selection.Font.NameFarEast = "黑体"
            s = Selection.Paragraphs(1).Range.Text
            Selection.Paragraphs(1).Range.Text = Trim(s)
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            Selection.Paragraphs(1).Range.Select
        End If
    End With
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Next j
    Selection.WholeStory
    Selection.HomeKey Unit:=wdStory
    
    '对节进行处理，宋体加粗+居中
    Selection.WholeStory
    Selection.HomeKey Unit:=wdStory
    For j = 0 To ActiveDocument.Paragraphs.Count - 1
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="节" & Chr(-24159)
        If .Found = True Then
            'Selection.MoveRight Unit:=wdCharacter, Count:=1
            'Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
            Selection.Paragraphs(1).Range.Select
            Selection.Font.Bold = True
            s = Selection.Paragraphs(1).Range.Text
            Selection.Paragraphs(1).Range.Text = Trim(s)
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            Selection.Paragraphs(1).Range.Select
        End If
    End With
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Next j
    Selection.WholeStory
    Selection.HomeKey Unit:=wdStory
    Application.ScreenUpdating = True '恢复屏幕更新
End Sub

Sub A00_通知()
    
    A00_网页格式
    
    '判断是否符合执行本宏命令的条件，要符合通知特征
    Selection.HomeKey Unit:=wdStory '定位到文档开头
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="：" & Chr(13)
        If .Found = True Then
            Selection.Paragraphs(1).Range.Select
            s = Selection.Paragraphs(1).Range.Text
            Selection.Paragraphs(1).Range.Text = Trim(s)
            Selection.HomeKey Unit:=wdStory '定位到文档开头
        End If
    End With
    
    '文号居中
    With Selection.Find
        .ClearFormatting
        .Execute FindText:=Chr(-24142)
        If .Found = True Then
            Selection.Paragraphs(1).Range.Select
            s = Selection.Paragraphs(1).Range.Text
            Selection.Paragraphs(1).Range.Text = Trim(s)
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            'Selection.HomeKey Unit:=wdStory '定位到文档开头
        End If
    End With
    
    Set aRange = ActiveDocument.Range(start:=ActiveDocument.Range.start, End:=Selection.Paragraphs(1).Range.End)
    ParaNum = aRange.Paragraphs.Count

    Select Case ParaNum
        Case 7
            ActiveDocument.Paragraphs(3).Range.Select
            Selection.Font.Bold = True
            s = Selection.Paragraphs(1).Range.Text
            Selection.Paragraphs(1).Range.Text = Trim(s)
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            ActiveDocument.Paragraphs(5).Range.Select
            Selection.Font.Bold = True
            s = Selection.Paragraphs(1).Range.Text
            Selection.Paragraphs(1).Range.Text = Trim(s)
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            ActiveDocument.Paragraphs(4).Range.Select
            Selection.Delete
            ActiveDocument.Paragraphs(2).Range.Select
            Selection.Delete
        Case 5
            ActiveDocument.Paragraphs(3).Range.Select
            Selection.Font.Bold = True
            s = Selection.Paragraphs(1).Range.Text
            Selection.Paragraphs(1).Range.Text = Trim(s)
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            ActiveDocument.Paragraphs(2).Range.Select
            Selection.Delete
    End Select
    
    '判断发文日期
    Selection.HomeKey Unit:=wdStory '定位到文档开头
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="日" & Chr(13)
        If .Found = True Then
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            Selection.Paragraphs(1).Range.Select
            s = Selection.Paragraphs(1).Range.Text
            Selection.Paragraphs(1).Range.Text = Trim(s)
            Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
            Selection.Paragraphs(1).Range.Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            Selection.TypeBackspace
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            Selection.Paragraphs(1).Range.Select
            s = Selection.Paragraphs(1).Range.Text
            Selection.Paragraphs(1).Range.Text = Trim(s)
            Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
            Selection.HomeKey Unit:=wdStory '定位到文档开头
        End If
    End With
    
End Sub

Sub A01_各地联系方式表格标准化()
    页面设置
    ActiveDocument.Tables(1).Select
    A01_各地联系方式表格标准化1
    ActiveDocument.Tables(2).Select
    A01_各地联系方式表格标准化2
    ActiveDocument.Tables(3).Select
    A01_各地联系方式表格标准化3
End Sub
Sub A01_各地联系方式表格标准化1()
    
    表格B
    
    Selection.SelectRow
    Selection.Font.Bold = True '表头字体加粗
    
    Selection.Tables(1).Select '整表行高最小值0.8厘米
    Selection.Rows.HeightRule = wdRowHeightAtLeast
    Selection.Rows.Height = CentimetersToPoints(0.8)
    
    Selection.SelectColumn '第一列居中
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    Selection.Font.Color = wdColorAutomatic
    
    Selection.Tables(1).Columns(2).Select '第二列居左
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Font.Color = wdColorAutomatic
    
    Selection.Tables(1).Columns(3).Select '第二列居中
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Font.Color = wdColorAutomatic
    
    Selection.Tables(1).Columns(4).Select '第二列居左
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    Selection.Font.Color = wdColorAutomatic
    
    Selection.Tables(1).Select '设置表格线为1/4磅线
    With Selection.Borders(wdBorderTop)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderLeft)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderBottom)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderRight)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderHorizontal)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderVertical)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    
    Selection.Tables(1).Select
    '全表底纹设置为10%灰色
    Selection.Shading.Texture = wdTextureNone
    Selection.Shading.ForegroundPatternColor = wdColorAutomatic
    Selection.Shading.BackgroundPatternColor = wdColorGray10
    
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.SelectRow '第一行居中
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.MoveLeft Unit:=wdCharacter, Count:=1

End Sub

Sub A01_各地联系方式表格标准化2()

    表格B
    
    Selection.SelectRow
    Selection.Font.Bold = True '表头字体加粗
    
    Selection.Tables(1).Select '整表行高最小值0.8厘米
    Selection.Rows.HeightRule = wdRowHeightAtLeast
    Selection.Rows.Height = CentimetersToPoints(0.8)
    
    Selection.SelectColumn '第一列居左
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    Selection.Font.Color = wdColorAutomatic
    
    Selection.Tables(1).Columns(2).Select '第二列居中
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Font.Color = wdColorAutomatic
    
    Selection.Tables(1).Columns(3).Select  '第三列居左
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    Selection.Font.Color = wdColorAutomatic
    
    Selection.Tables(1).Select '设置表格线为1/4磅线
    With Selection.Borders(wdBorderTop)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderLeft)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderBottom)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderRight)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderHorizontal)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderVertical)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    
    Selection.Tables(1).Select
    '全表底纹设置为10%灰色
    Selection.Shading.Texture = wdTextureNone
    Selection.Shading.ForegroundPatternColor = wdColorAutomatic
    Selection.Shading.BackgroundPatternColor = wdColorGray10
    
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.SelectRow '第一行居中
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.MoveLeft Unit:=wdCharacter, Count:=1

End Sub

Sub A01_各地联系方式表格标准化3()
    Application.Run MacroName:="Normal.NewMacros.表格B"
    
    Selection.SelectRow
    Selection.Cells.Merge
    Selection.Font.Bold = True '表头字体加粗
    
    Selection.Tables(1).Select '整表行高最小值0.8厘米
    Selection.Rows.HeightRule = wdRowHeightAtLeast
    Selection.Rows.Height = CentimetersToPoints(0.8)
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify

    Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderVertical).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
    
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.SelectRow
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    
    Selection.HomeKey Unit:=wdStory
    

    
End Sub

Sub A00_部委简称替换()
    Dim AA, BB As Variant
    Application.ScreenUpdating = False '关闭屏幕更新
    AA = Array("发改委", "工信部", "卫计委", "住建部", "环保部", "人社部", "国土部", "交通部")
    BB = Array("发展改革委", "工业和信息化部", "卫生计生委", "住房城乡建设部", "环境保护部", "人力资源社会保障部", "国土资源部", "交通运输部")
    For j = 0 To UBound(AA)
        Selection.HomeKey Unit:=wdStory '定位到文档开头
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = AA(j)
            .Replacement.Text = BB(j)
            .Wrap = wdFindStop
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
    Next j
    Application.ScreenUpdating = True '恢复屏幕更新
End Sub

Sub A00_不规范文号替换()
    Application.ScreenUpdating = False '关闭屏幕更新
    Selection.HomeKey Unit:=wdStory '定位到文档开头
    Do While True
    With Selection.Find
        .ClearFormatting
        .Execute FindText:=Chr(91) '查找“[”
        If .Found = True Then
            Selection.MoveRight Unit:=wdCharacter, Count:=5, Extend:=wdExtend
            Set MyRange = Selection.Range
            If Right(MyRange.Text, 1) = Chr(93) Then '判断第5个字符是否为“]”
                Z1 = Right(Left(MyRange.Text, 2), 1)
                Z2 = Right(Left(MyRange.Text, 3), 1)
                Z3 = Right(Left(MyRange.Text, 4), 1)
                Z4 = Right(Left(MyRange.Text, 5), 1)
                Z5 = Right(MyRange.Text, 1)
                TT = Chr(-24142) & Z1 & Z2 & Z3 & Z4 & Chr(-24141)
                '判断中间的四个字符是否都是数字,如果是，则替换为六角括号，否则，忽略
                If Asc(Z1) > 47 And Asc(Z1) < 58 And Asc(Z2) > 47 And Asc(Z2) < 58 And Asc(Z3) > 47 And Asc(Z3) < 58 And Asc(Z4) > 47 And Asc(Z4) < 58 Then
                    MyRange.Text = TT
                End If
            End If
        Else
            Exit Do
        End If
    End With
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Loop
    Application.ScreenUpdating = True '恢复屏幕更新
    Selection.HomeKey Unit:=wdStory '定位到文档开头
    
    
End Sub


Sub A0_OA另存_党组文件()

N = ActiveDocument.Shapes.Count ' 取得文档中图片数

MsgBox N

    For Each s In ActiveDocument.Shapes
        With s.TextFrame
            If .HasText Then MsgBox .TextRange.Text
        End With
    Next

    Selection.MoveDown Unit:=wdLine, Count:=1
End Sub

Sub A00_加注释序号()

    '查找特定的字符串，如："["，然后将该字符串前的内容加粗
    
    Application.ScreenUpdating = False '关闭屏幕更新
    
    '判断是否符合执行本宏命令的条件，正文中必须有“[ ]”
    Selection.HomeKey Unit:=wdStory '定位到文档开头
    i = 0
    Do While True
    With Selection.Find
        .ClearFormatting
        .Execute FindText:=Chr(91) & Chr(32) & Chr(93)
        If .Found = True Then
            i = i + 1
            Selection.Range.Text = Chr(91) & i & Chr(93)
            Selection.MoveRight Unit:=wdCharacter, Count:=1
        Else
            Selection.HomeKey Unit:=wdStory '定位到文档开头
            Exit Do
        End If
    End With
    Loop
    
    MsgBox i
    
    Application.ScreenUpdating = True '恢复屏幕更新
End Sub

Sub AutoNew()
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FolderExists("D:\00 Word_Dot") = True Then
        AddIns("D:\00 Word_Dot\Normal.dot").Installed = True
    Else
        If fs.FolderExists("F:\00 Word_Dot") = True Then
            AddIns("F:\00 Word_Dot\Normal.dot").Installed = True
        End If
    End If
    With ActiveDocument
        .UpdateStylesOnOpen = False
        .AttachedTemplate = "Normal"
    End With
End Sub
Sub AutoOpen()
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FolderExists("D:\00 Word_Dot") = True Then
        AddIns("D:\00 Word_Dot\Normal.dot").Installed = True
    Else
        If fs.FolderExists("F:\00 Word_Dot") = True Then
            AddIns("F:\00 Word_Dot\Normal.dot").Installed = True
        End If
    End If
    With ActiveDocument
        .UpdateStylesOnOpen = False
        .AttachedTemplate = "Normal"
    End With
End Sub


Sub A00_删多余空行()

    Application.ScreenUpdating = False '关闭屏幕更新
    On Error Resume Next
    If ActiveDocument.Tables.Count = 0 Then Exit Sub
    Selection.HomeKey Unit:=wdStory
    For i = 1 To ActiveDocument.Tables.Count
        ActiveDocument.Tables(i).Select
        Selection.MoveLeft Unit:=wdCharacter, Count:=2
        Selection.TypeBackspace
        ActiveDocument.Tables(i).Select
        Selection.MoveRight Unit:=wdCharacter, Count:=1
    Next i
    Selection.HomeKey Unit:=wdStory
    Application.ScreenUpdating = True '恢复屏幕更新
End Sub


Sub A01_文转表()
    
    Application.ScreenUpdating = False '关闭屏幕更新
    Dim TN, RN As Integer
    Dim TT As String
    TT = Selection.Range.Text
    TN = 0
    
    If Len(TT) < 2 Then
        MsgBox "请选择需要转换的内容！"
    Else
        RN = Selection.Paragraphs.Count
        If InStr(1, TT, Chr(9), 1) > 0 Then
            For Each st In Selection.Paragraphs(1).Range.Characters
                If st = Chr(9) Then
                    TN = TN + 1
                 End If
            Next st
            Selection.ConvertToTable Separator:=wdSeparateByTabs, NumColumns:=TN + 1, NumRows:=RN, AutoFitBehavior:=wdAutoFitFixed
            Application.Run MacroName:="表格B"
        Else
            Selection.ConvertToTable Separator:=wdSeparateByParagraphs, NumColumns:=1, NumRows:=RN, AutoFitBehavior:=wdAutoFitFixed
        End If
    End If
    Selection.Tables(1).Select
    Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
    Selection.Tables(1).Select
    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    Selection.Tables(1).Select
    Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed)
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Application.ScreenUpdating = True '恢复屏幕更新
    
End Sub


Function BackTXT(TT) '返回删除指定字符集后的文本
    Dim P1 As Integer
    Dim C As Variant
    C = Array(Chr(13), Chr(32), Chr(-24159), Chr(11), Chr(9), Chr(7))
    For i = 0 To UBound(C)
        P1 = InStr(1, TT, C(i), 1)
        If P1 > 0 Then
            Do While P1 > 0
                TT = Left(TT, P1 - 1) & Right(TT, Len(TT) - P1)
                P1 = InStr(1, TT, C(i), 1)
            Loop
        End If
    Next i
    MsgBox TT
    'MsgBox Asc(Right(TT, 1))
    MsgBox Len(TT)
    
End Function
Sub A00_单位列居中()
    On Error Resume Next
    Set MyTab = Selection.Tables(1)
    L = MyTab.Columns.Count
    For i = 1 To MyTab.Columns.Count
        MyTab.Cell(1, i).Select
        TT = Selection.Range.Text
        If InStr(1, TT, "单", 1) > 0 And InStr(1, TT, "位", 1) > 0 Then
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            Selection.SelectColumn
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End If
    Next i
    Selection.Tables(1).Cell(1, 1).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=1

End Sub
Sub A01_测试()  '功能：调试宏命令

    Set MyTab = Selection.Tables(1)
    L = MyTab.Columns.Count
    For i = 1 To MyTab.Columns.Count
        MyTab.Cell(1, i).Select
        TT = Selection.Range.Text
        If InStr(1, TT, "单", 1) > 0 And InStr(1, TT, "位", 1) > 0 Then
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            Selection.SelectColumn
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End If
    Next i
    Selection.Tables(1).Cell(1, 1).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=1

End Sub
Sub A00_删除特定空行()
    'MsgBox Len(Selection.Paragraphs(1).Range.Text)
    
    
    On Error Resume Next
    Selection.HomeKey Unit:=wdStory '定位到文档开头
    For Each aPara In ActiveDocument.Paragraphs
        LN = Len(aPara.Range.Text)
        If LN = 4 And Left(aPara.Range.Text, 2) = "　　" Then
            aPara.Range.Select
            Selection.Delete
            Selection.Delete
        End If
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Next aPara
    'Selection.Range.PasteSpecial DataType:=wdPasteText
    Selection.HomeKey Unit:=wdStory
End Sub
Sub A00_查找数字列()
    
    On Error Resume Next
    Set MyTab = Selection.Tables(1)
    Selection.Tables(1).Cell(1, 1).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    For H = 2 To MyTab.Rows.Count
    For L = 2 To MyTab.Columns.Count
            TT = MyTab.Cell(H, L).Range.Text
            TT = Left(TT, Len(TT) - 2)
            C1 = Left(TT, 1)
            C2 = Right(TT, 1)
            If (Asc(C1) > 47 And Asc(C1) < 58) And (Asc(C2) > 47 And Asc(C2) < 58) And Val(TT) <> 0 Then
               MyTab.Cell(H, L).Select
               LN = Len(Selection.Range.Text)
               H1 = Selection.Information(wdStartOfRangeRowNumber)
               L1 = Selection.Information(wdStartOfRangeColumnNumber)
                MsgBox "第" & H1 & "行第" & L1 & "列开始为数字"
               Exit For
            End If
    Next L
    If H1 > 1 And Val(TT) <> 0 Then Exit For
    Next H

End Sub

Sub A00_第一列字数相近居中()
    
    Dim MyTab As Table
    Dim LN() As Variant
    Dim H, L, M, N, H1, L1 As Integer
    Dim TT As String
    
    On Error Resume Next
    Set MyTab = Selection.Tables(1)
    Selection.Tables(1).Cell(1, 1).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=1
    H1 = Selection.Information(wdStartOfRangeRowNumber)
    
    For j = 0 To MyTab.Rows.Count
        TT = MyTab.Cell(j + H1, 1).Range.Text
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
    
    If M - N < 4 Then
        Selection.SelectColumn
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End If
    MyTab.Cell(1, 1).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
End Sub
Sub A00_表格顶线和底线加粗()
    
    Dim TB As Table '定义表格变量
    
    Application.ScreenUpdating = False '关闭屏幕更新
    
    '如果没有选择表格，提示用户选择要处理的表格
    If Selection.Information(wdWithInTable) = False Then
        MsgBox "请选择表格！"
        Exit Sub
    End If

    Set TB = Selection.Tables(1)
    TB.Select
        With TB.Borders(wdBorderTop)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth150pt
            .Color = wdColorBlack
        End With
        TB.Cell(TB.Rows.Count, 1).Select
        Selection.SelectRow
        
        With Selection.Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth150pt
            .Color = wdColorBlack
        End With
    TB.Cell(1, 1).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Application.ScreenUpdating = True '恢复屏幕更新
End Sub

Sub A00_表格每列对齐方式()
    
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
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=1
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
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Application.ScreenUpdating = True '恢复屏幕更新
    
End Sub

Sub A00_表格数字列右对齐()
    
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
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=1
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
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Application.ScreenUpdating = True '恢复屏幕更新
End Sub

Sub A00_关闭所有文档并退出()
    Documents.Close SaveChanges:=wdDoNotSaveChanges
    Application.Quit SaveChanges:=wdDoNotSaveChanges
End Sub

Sub A00_附注()
    
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="　　附注"
        If .Found = True Then
            Selection.MoveRight Unit:=wdCharacter, Count:=1
            Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
            Selection.Font.Bold = True
            Set MyRange = ActiveDocument.Range(start:=Selection.Paragraphs(1).Range.start, End:=ActiveDocument.Range.End)
            MyRange.Font.NameFarEast = "楷体"
        Else
            Exit Sub
        End If
    End With

End Sub

Sub A00_PMI()
    
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="　　中国物流与采购联合会"
        If .Found = True Then
            Selection.MoveRight Unit:=wdCharacter, Count:=1
            Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
            Selection.Font.NameFarEast = "楷体"
            Selection.Paragraphs(1).Range.Text = "中国物流与采购联合会" & Chr(13)
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            Selection.MoveUp Unit:=wdLine, Count:=1
            Selection.TypeBackspace
            Selection.Paragraphs(1).Range.Select
            Selection.Font.NameFarEast = "楷体"
            Selection.Paragraphs(1).Range.Text = "国家统计局服务业调查中心" & Chr(13)
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Else
            Exit Sub
        End If
    End With
    Selection.HomeKey Unit:=wdStory

End Sub


Sub A00_当前文档转变为纯文本()
    
    页面设置
    Selection.HomeKey Unit:=wdStory
    Selection.WholeStory
    TT = ActiveDocument.Content.Text
    Documents.Add DocumentType:=wdNewBlankDocument
    页面设置
    ActiveDocument.Content.Text = TT
    
End Sub


Sub A00_国际经济信息()
    
    Application.ScreenUpdating = False '关闭屏幕更新
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .Text = "返回目录"
        .Replacement.Text = ""
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    ActiveDocument.Save
    A00_去表注
    A00_去表头
    A00_去表头
    A00_文图表
    A00_IEI
    
    Application.ScreenUpdating = True '恢复屏幕更新

End Sub

Sub A00_IEI_标题居中()
    TT = Selection.Paragraphs(1).Range.Text
    TT = Left(TT, Len(TT) - 1)
    JH = Right(TT, 1)
    If Len(TT) < 40 And Len(TT) > 2 And JH <> "。" Then
        Selection.Paragraphs(1).Range.Select
        Selection.Font.Bold = True
        Selection.Paragraphs(1).Range.Text = Trim(TT) & Chr(13)
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        'Selection.MoveRight Unit:=wdCharacter, Count:=1
    End If
End Sub

Sub A00_IEI_指定字符串居左()
    Dim A As Variant
    A = Array("国际经济数据", "国际经济政策", "国际市场动态", "国外观点")
    
    TT = Selection.Paragraphs(1).Range.Text
    For k = 0 To UBound(A)
        If InStr(1, TT, A(k) & Chr(13), 1) > 0 Then
            Selection.Paragraphs(1).Range.Text = A(k) & Chr(13)
            Selection.Paragraphs(1).Range.Select
            Selection.Font.Bold = True
            Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
        End If
    Next k
End Sub

Sub A00_IEI()
    
    Dim RNG As Range
    Dim MyDOC As Document
    Dim TN As Integer
    
    Application.ScreenUpdating = False '关闭屏幕更新
    On Error Resume Next
    Set MyDOC = ActiveDocument
    TN = MyDOC.Tables.Count
    
    Select Case TN
    Case 0
        Set RNG = MyDOC.Range
        For j = 1 To RNG.Paragraphs.Count
            RNG.Paragraphs(j).Range.Select
            A00_IEI_标题居中
            A00_IEI_指定字符串居左
        Next j
    Case 1
        Set RNG = MyDOC.Range(start:=MyDOC.Paragraphs(1).Range.start, End:=MyDOC.Tables(1).Range.start - 1)
        For j = 1 To RNG.Paragraphs.Count
            RNG.Paragraphs(j).Range.Select
            A00_IEI_标题居中
            A00_IEI_指定字符串居左
        Next j
        
        Set RNG = MyDOC.Range(start:=MyDOC.Tables(1).Range.End + 1, End:=MyDOC.Range.End)
        For j = 1 To RNG.Paragraphs.Count
            RNG.Paragraphs(j).Range.Select
            A00_IEI_标题居中
            A00_IEI_指定字符串居左
        Next j
    
    Case 2
        Set RNG = MyDOC.Range(start:=MyDOC.Paragraphs(1).Range.start, End:=MyDOC.Tables(1).Range.start - 1)
        For j = 1 To RNG.Paragraphs.Count
            RNG.Paragraphs(j).Range.Select
            A00_IEI_标题居中
            A00_IEI_指定字符串居左
        Next j
        
        Set RNG = MyDOC.Range(start:=MyDOC.Tables(1).Range.End + 1, End:=MyDOC.Tables(2).Range.start - 1)
        For j = 1 To RNG.Paragraphs.Count
            RNG.Paragraphs(j).Range.Select
            A00_IEI_标题居中
            A00_IEI_指定字符串居左
        Next j
        
        Set RNG = MyDOC.Range(start:=MyDOC.Tables(2).Range.End + 1, End:=MyDOC.Range.End)
        For j = 1 To RNG.Paragraphs.Count
            RNG.Paragraphs(j).Range.Select
            A00_IEI_标题居中
            A00_IEI_指定字符串居左
        Next j
        
    Case 2
        Set RNG = MyDOC.Range(start:=MyDOC.Paragraphs(1).Range.start, End:=MyDOC.Tables(1).Range.start - 1)
        For j = 1 To RNG.Paragraphs.Count
            RNG.Paragraphs(j).Range.Select
            A00_IEI_标题居中
            A00_IEI_指定字符串居左
        Next j
        
        Set RNG = MyDOC.Range(start:=MyDOC.Tables(1).Range.End + 1, End:=MyDOC.Range.End)
        For j = 1 To RNG.Paragraphs.Count
            RNG.Paragraphs(j).Range.Select
            A00_IEI_标题居中
            A00_IEI_指定字符串居左
        Next j
    
    Case 3
        Set RNG = MyDOC.Range(start:=MyDOC.Paragraphs(1).Range.start, End:=MyDOC.Tables(1).Range.start - 1)
        For j = 1 To RNG.Paragraphs.Count
            RNG.Paragraphs(j).Range.Select
            A00_IEI_标题居中
            A00_IEI_指定字符串居左
        Next j
        
        Set RNG = MyDOC.Range(start:=MyDOC.Tables(1).Range.End + 1, End:=MyDOC.Tables(2).Range.start - 1)
        For j = 1 To RNG.Paragraphs.Count
            RNG.Paragraphs(j).Range.Select
            A00_IEI_标题居中
            A00_IEI_指定字符串居左
        Next j
        
        Set RNG = MyDOC.Range(start:=MyDOC.Tables(2).Range.End + 1, End:=MyDOC.Tables(3).Range.start - 1)
        For j = 1 To RNG.Paragraphs.Count
            RNG.Paragraphs(j).Range.Select
            A00_IEI_标题居中
            A00_IEI_指定字符串居左
        Next j
        
        Set RNG = MyDOC.Range(start:=MyDOC.Tables(3).Range.End + 1, End:=MyDOC.Range.End)
        For j = 1 To RNG.Paragraphs.Count
            RNG.Paragraphs(j).Range.Select
            A00_IEI_标题居中
            A00_IEI_指定字符串居左
        Next j
        
    End Select
    
    Selection.HomeKey Unit:=wdStory
    Application.ScreenUpdating = True '恢复屏幕更新
End Sub
Sub A00_活动文档转变为纯文本()
    
    Dim MyRange As Range  '定义一个范围（Range）变量
    '如果没有选择范围，则指定范围为整个文档
    On Error Resume Next
    If Len(Selection.Range.Text) = 0 Then Selection.WholeStory
    Set MyRange = Selection.Range '设定范围变量为选择的范围
    MyRange.Copy
    Documents.Add DocumentType:=wdNewBlankDocument
    页面设置
    Selection.PasteSpecial Link:=False, DataType:=wdPasteText, Placement:=wdInLine, DisplayAsIcon:=False
    A00_网页格式
        
End Sub

Sub A00_去表头()
    
    Dim RNG As Range
    Dim MyDOC As Document
    Dim TN As Integer
    
    Application.ScreenUpdating = False '关闭屏幕更新
    On Error Resume Next
    Set MyDOC = ActiveDocument
    TN = MyDOC.Tables.Count
    
    Selection.HomeKey Unit:=wdStory
    For i = 1 To TN
        MyDOC.Tables(i).Select
        MyDOC.Tables(i).Cell(1, 1).Select
        Selection.SelectRow
        C = Selection.Cells.Count
        If C < 2 Then
            TT = Selection.Range.Text
            TT = Left(TT, Len(TT) - 2)
            Selection.Cut
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            Selection.TypeParagraph
            Selection.TypeParagraph
            Selection.MoveUp Unit:=wdLine, Count:=1
            Selection.Paragraphs(1).Range.Text = TT
        End If
    Next i
    Selection.HomeKey Unit:=wdStory
    Application.ScreenUpdating = True '恢复屏幕更新

End Sub

Sub A00_去表注()
    
    Dim RNG As Range
    Dim MyDOC As Document
    Dim TN As Integer
    
    Application.ScreenUpdating = False '关闭屏幕更新
    On Error Resume Next
    Set MyDOC = ActiveDocument
    TN = MyDOC.Tables.Count
    
    Selection.HomeKey Unit:=wdStory
    For i = 1 To TN
        MyDOC.Tables(i).Select
        R = MyDOC.Tables(i).Rows.Count
        MyDOC.Tables(i).Cell(R, 1).Select
        Selection.SelectRow
        C = Selection.Cells.Count
        If C < 2 Then
            TT = Selection.Range.Text
            TT = Left(TT, Len(TT) - 2)
            Selection.Cut
            Selection.TypeParagraph
            Selection.TypeParagraph
            Selection.MoveUp Unit:=wdLine, Count:=1
            Selection.Paragraphs(1).Range.Text = TT
        End If
    Next i
    Selection.HomeKey Unit:=wdStory
    Application.ScreenUpdating = True '恢复屏幕更新

End Sub

Sub A00_批量加表注()
    Dim RNG As Range
    Dim MyDOC As Document
    Dim TN As Integer
    
    Application.ScreenUpdating = False '关闭屏幕更新
    On Error Resume Next
    Set MyDOC = ActiveDocument
    TN = MyDOC.Tables.Count
    
    Selection.HomeKey Unit:=wdStory
    For i = 1 To TN
        MyDOC.Tables(i).Select
        Selection.MoveRight Unit:=wdCharacter, Count:=2
        Selection.Paragraphs(1).Range.Select
        TT = Selection.Range.Text
        TT = Left(TT, Len(TT) - 1)
        If InStr(1, TT, "注：", 1) > 0 Or InStr(1, TT, "资料来源：", 1) > 0 Then
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            Selection.TypeBackspace
            Selection.Paragraphs(1).Range.Select
            B01_选定文本变为表注
        End If
    Next i
   
    Application.ScreenUpdating = True '恢复屏幕更新

End Sub

Sub A00_间隔号()
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ChrW(8226)
        .Replacement.Text = "·"
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub A00_网页空格替换()
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = Chr(32) & Chr(13)
        .Replacement.Text = Chr(13)
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub



Sub A00_TabToArray()
    
    Dim A As Variant
    Dim MyDOC As Document
    Dim MyOK As Document
    Set MyDOC = ActiveDocument
    Documents.Open ("C:\Users\zlzx-dhg\Desktop\IEI_Tab1.Docx")
    
    Set MyOK = ActiveDocument
    
    MyDOC.Activate
    If ActiveDocument.Tables.Count > 0 Then
        Set TB = ActiveDocument.Tables(1)
    End If
    
    N = TB.Range.Cells.Count
    
    ReDim A(1)
        TT = TB.Range.Cells(1).Range.Text
        TT = Left(TT, Len(TT) - 2)
        A(0) = TT
    
    For i = 2 To N
        ReDim Preserve A(i)
        TT = TB.Range.Cells(i).Range.Text
        TT = Left(TT, Len(TT) - 2)
        A(i - 1) = TT
    Next i
    
    MyOK.Activate
    Set TB1 = MyOK.Tables(1)
    N1 = TB1.Range.Cells.Count
    If N1 = N Then
    
    For j = 0 To UBound(A) - 1
        TB1.Range.Cells(j + 1).Range.Text = A(j)
    Next j
    End If
    
    TB1.Range.Copy
    MyDOC.Activate
    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.Paste
    MyOK.Activate
    ActiveDocument.Close wdDoNotSaveChanges
    MyDOC.Activate
    
End Sub

Sub A00_IEI_SI()
    Dim A As Variant
    Dim B As Variant
    Dim C As Variant
    
    Set TB1 = ActiveDocument.Tables(1)
    Set TB2 = ActiveDocument.Tables(2)
    ReDim A(1)
        TT = TB1.Cell(1, 2).Range.Text
        TT = Left(TT, Len(TT) - 2)
        A(0) = TT & Chr(13) & "涨跌幅"
    
    For i = 2 To 7
        ReDim Preserve A(i)
        TT = TB1.Cell(1, i + 1).Range.Text
        TT = Left(TT, Len(TT) - 2)
        A(i - 1) = TT & Chr(13) & "涨跌幅"
    Next i
    
    For j = 1 To UBound(A)
        TB2.Cell(2, j + 1).Range.Text = A(j - 1)
    Next j
    
    ReDim B(1)
        TT = TB1.Cell(3, 8).Range.Text
        TT = Left(TT, Len(TT) - 2)
        B(0) = TT
    ReDim C(1)
        TT = TB1.Cell(4, 8).Range.Text
        TT = Left(TT, Len(TT) - 2)
        C(0) = TT
   
    For i = 2 To 6
        ReDim Preserve B(i)
        TT = TB1.Cell(i + i + 2 + 1, 8).Range.Text
        TT = Left(TT, Len(TT) - 2)
        B(i - 1) = TT
    Next i
    
    For i = 2 To 6
        ReDim Preserve C(i)
        TT = TB1.Cell(i + i * 2 + 2, 8).Range.Text
        TT = Left(TT, Len(TT) - 2)
        C(i - 1) = TT
    Next i

    
    For j = 1 To UBound(B) + 1
        TB2.Cell(j + 2, 8).Range.Text = B(j - 1) & Chr(13) & C(j - 1)
    Next j
    
End Sub

Sub A00_C135()
    Dim DC1 As Document
    Dim DC2 As Document
    Documents("专栏").Activate
    Set DC1 = ActiveDocument
    Documents("国家“十三五”时期文化发展改革规划纲要").Activate
    Set DC2 = ActiveDocument
    DC1.Activate
    For i = 1 To DC1.Tables.Count
        NM = "TAB" & i
        DC1.Tables(i).Range.Copy
        DC2.Activate
        With Selection.Find
            .ClearFormatting
            .Execute FindText:="TAB" & i
            If .Found = True Then
                .Parent.Expand Unit:=wdParagraph
            Else
                Exit For
            End If
        End With
        Selection.Delete Unit:=wdCharacter, Count:=1
        Selection.Paste
    Next i
    
End Sub

Sub A00_IEI_TB4()
    Dim A As Variant
    Dim B As Variant
    Dim C As Variant
    Dim D As Variant
    
    Application.ScreenUpdating = False '关闭屏幕更新
    On Error Resume Next
    
    Set TB1 = ActiveDocument.Tables(1)
    Set TB2 = ActiveDocument.Tables(2)
    ReDim A(1)
        TT = TB1.Cell(1, 2).Range.Text
        TT = Left(TT, Len(TT) - 2)
        A(0) = TT & Chr(13) & "涨跌幅"
    
    For i = 2 To 7
        ReDim Preserve A(i)
        TT = TB1.Cell(1, i + 1).Range.Text
        TT = Left(TT, Len(TT) - 2)
        A(i - 1) = TT & Chr(13) & "涨跌幅"
    Next i
    
    For j = 1 To UBound(A)
        TB2.Cell(2, j + 1).Range.Text = A(j - 1)
    Next j
    
        TT = TB1.Cell(3, 8).Range.Text
        TT = Left(TT, Len(TT) - 2)
        B1 = TT
        
        TT = TB1.Cell(5, 8).Range.Text
        TT = Left(TT, Len(TT) - 2)
        B2 = TT
        
        TT = TB1.Cell(7, 8).Range.Text
        TT = Left(TT, Len(TT) - 2)
        B3 = TT
        
        TT = TB1.Cell(9, 8).Range.Text
        TT = Left(TT, Len(TT) - 2)
        B4 = TT
        
        TT = TB1.Cell(11, 8).Range.Text
        TT = Left(TT, Len(TT) - 2)
        B5 = TT
        
        TT = TB1.Cell(13, 8).Range.Text
        TT = Left(TT, Len(TT) - 2)
        B6 = TT
        
        TT = TB1.Cell(15, 8).Range.Text
        TT = Left(TT, Len(TT) - 2)
        B7 = TT
        
        TT = TB1.Cell(17, 8).Range.Text
        TT = Left(TT, Len(TT) - 2)
        B8 = TT
        
        TT = TB1.Cell(19, 8).Range.Text
        TT = Left(TT, Len(TT) - 2)
        B9 = TT
        
        B = Array(B1, B2, B3, B4, B5, B6, B7, B8, B9)
        
        TT = TB1.Cell(4, 8).Range.Text
        TT = Left(TT, Len(TT) - 2)
        C1 = TT
    
        TT = TB1.Cell(6, 8).Range.Text
        TT = Left(TT, Len(TT) - 2)
        C2 = TT
        
        TT = TB1.Cell(8, 8).Range.Text
        TT = Left(TT, Len(TT) - 2)
        C3 = TT
        
        TT = TB1.Cell(10, 8).Range.Text
        TT = Left(TT, Len(TT) - 2)
        C4 = TT
        
        TT = TB1.Cell(12, 8).Range.Text
        TT = Left(TT, Len(TT) - 2)
        C5 = TT
        
        TT = TB1.Cell(14, 8).Range.Text
        TT = Left(TT, Len(TT) - 2)
        C6 = TT
        
        TT = TB1.Cell(16, 8).Range.Text
        TT = Left(TT, Len(TT) - 2)
        C7 = TT
        
        TT = TB1.Cell(18, 8).Range.Text
        TT = Left(TT, Len(TT) - 2)
        C8 = TT
        
        TT = TB1.Cell(20, 8).Range.Text
        TT = Left(TT, Len(TT) - 2)
        C9 = TT
        
        C = Array(C1, C2, C3, C4, C5, C6, C7, C8, C9)
    
    For j = 1 To UBound(B) + 1
        TB2.Cell(j + 2, 8).Range.Text = B(j - 1) & Chr(13) & C(j - 1)
    Next j
    
        For i = 1 To 7
        
        TT = TB1.Cell(3, i).Range.Text
        TT = Left(TT, Len(TT) - 2)
        D1 = TT
    
        TT = TB1.Cell(5, i).Range.Text
        TT = Left(TT, Len(TT) - 2)
        D2 = TT
        
        TT = TB1.Cell(7, i).Range.Text
        TT = Left(TT, Len(TT) - 2)
        D3 = TT
        
        TT = TB1.Cell(9, i).Range.Text
        TT = Left(TT, Len(TT) - 2)
        D4 = TT
        
        TT = TB1.Cell(11, i).Range.Text
        TT = Left(TT, Len(TT) - 2)
        D5 = TT
        
        TT = TB1.Cell(13, i).Range.Text
        TT = Left(TT, Len(TT) - 2)
        D6 = TT
        
        TT = TB1.Cell(15, i).Range.Text
        TT = Left(TT, Len(TT) - 2)
        D7 = TT
        
        TT = TB1.Cell(17, i).Range.Text
        TT = Left(TT, Len(TT) - 2)
        D8 = TT
        
        TT = TB1.Cell(19, i).Range.Text
        TT = Left(TT, Len(TT) - 2)
        D9 = TT
        
        D = Array(D1, D2, D3, D4, D5, D6, D7, D8, D9)
        
    For j = 1 To UBound(D) + 1
        TB2.Cell(j + 2, i).Range.Text = D(j - 1)
    Next j
    Next i
    TB2.Cell(22, 1).Range.Text = Left(TB1.Cell(21, 1).Range.Text, Len(TB1.Cell(21, 1).Range.Text) - 2)
    Application.ScreenUpdating = True '恢复屏幕更新
End Sub

Sub A00_IEI_TB5()
    Dim A As Variant
    Dim B As Variant
    
    Application.ScreenUpdating = False '关闭屏幕更新
    On Error Resume Next
    
    Set TB1 = ActiveDocument.Tables(1)
    Set TB2 = ActiveDocument.Tables(2)
    
    For j = 1 To 8
        ReDim A(1)
            TT = TB1.Cell(4, j).Range.Text
            TT = Left(TT, Len(TT) - 2)
            A(0) = TT
        
        For i = 2 To 8
            ReDim Preserve A(i)
            TT = TB1.Cell(i + 3, j).Range.Text
            TT = Left(TT, Len(TT) - 2)
            A(i - 1) = TT
        Next i
        For k = 1 To UBound(A) + 1
            TB2.Cell(k + 4, j).Range.Text = A(k - 1)
        Next k
    Next j
        
    TB2.Cell(2, 2).Range.Text = Left(TB1.Cell(1, 2).Range.Text, Len(TB1.Cell(1, 2).Range.Text) - 2)
    TB2.Cell(2, 3).Range.Text = Left(TB1.Cell(1, 3).Range.Text, Len(TB1.Cell(1, 3).Range.Text) - 2)
    TB2.Cell(3, 3).Range.Text = Left(TB1.Cell(2, 3).Range.Text, Len(TB1.Cell(2, 3).Range.Text) - 2)
    TB2.Cell(3, 4).Range.Text = Left(TB1.Cell(2, 4).Range.Text, Len(TB1.Cell(2, 4).Range.Text) - 2)
    TB2.Cell(3, 5).Range.Text = Left(TB1.Cell(2, 5).Range.Text, Len(TB1.Cell(2, 5).Range.Text) - 2)
    TB2.Cell(3, 6).Range.Text = Left(TB1.Cell(2, 6).Range.Text, Len(TB1.Cell(2, 6).Range.Text) - 2)
    TB2.Cell(3, 7).Range.Text = Left(TB1.Cell(2, 7).Range.Text, Len(TB1.Cell(2, 7).Range.Text) - 2)
    TB2.Cell(4, 8).Range.Text = Left(TB1.Cell(3, 8).Range.Text, Len(TB1.Cell(3, 8).Range.Text) - 2)
    TB2.Cell(13, 1).Range.Text = Left(TB1.Cell(14, 1).Range.Text, Len(TB1.Cell(14, 1).Range.Text) - 2)
    
    Application.ScreenUpdating = True '恢复屏幕更新
End Sub

Sub A00_IEI_TB1()
    Dim A As Variant
    Dim B As Variant
    
    Application.ScreenUpdating = False '关闭屏幕更新
    On Error Resume Next
    
    Set TB1 = ActiveDocument.Tables(1)
    Set TB2 = ActiveDocument.Tables(2)
    
    For j = 1 To 8
        ReDim A(1)
            TT = TB1.Cell(4, j).Range.Text
            TT = Left(TT, Len(TT) - 2)
            A(0) = TT
        
        For i = 2 To 6
            ReDim Preserve A(i)
            TT = TB1.Cell(i + 3, j).Range.Text
            TT = Left(TT, Len(TT) - 2)
            A(i - 1) = TT
        Next i
        For k = 1 To UBound(A) + 1
            TB2.Cell(k + 5, j).Range.Text = A(k - 1)
        Next k
    Next j
        
    TB2.Cell(3, 2).Range.Text = Left(TB1.Cell(1, 2).Range.Text, Len(TB1.Cell(1, 2).Range.Text) - 2)
    TB2.Cell(3, 3).Range.Text = Left(TB1.Cell(1, 3).Range.Text, Len(TB1.Cell(1, 3).Range.Text) - 2)
    TB2.Cell(4, 3).Range.Text = Left(TB1.Cell(2, 3).Range.Text, Len(TB1.Cell(2, 3).Range.Text) - 2)
    TB2.Cell(4, 4).Range.Text = Left(TB1.Cell(2, 4).Range.Text, Len(TB1.Cell(2, 4).Range.Text) - 2)
    TB2.Cell(4, 5).Range.Text = Left(TB1.Cell(2, 5).Range.Text, Len(TB1.Cell(2, 5).Range.Text) - 2)
    TB2.Cell(4, 6).Range.Text = Left(TB1.Cell(2, 6).Range.Text, Len(TB1.Cell(2, 6).Range.Text) - 2)
    TB2.Cell(4, 7).Range.Text = Left(TB1.Cell(2, 7).Range.Text, Len(TB1.Cell(2, 7).Range.Text) - 2)
    TB2.Cell(5, 8).Range.Text = Left(TB1.Cell(3, 8).Range.Text, Len(TB1.Cell(3, 8).Range.Text) - 2)
    TB2.Cell(12, 1).Range.Text = Left(TB1.Cell(10, 1).Range.Text, Len(TB1.Cell(10, 1).Range.Text) - 2)
    
    Application.ScreenUpdating = True '恢复屏幕更新
End Sub

Sub A00_IEI_TB2()
    Dim A As Variant
    Dim B As Variant
    
    Application.ScreenUpdating = False '关闭屏幕更新
    On Error Resume Next
    
    Set TB1 = ActiveDocument.Tables(1)
    Set TB2 = ActiveDocument.Tables(2)
    
    For j = 1 To 8
        ReDim A(1)
            TT = TB1.Cell(4, j).Range.Text
            TT = Left(TT, Len(TT) - 2)
            A(0) = TT
        
        For i = 2 To 3
            ReDim Preserve A(i)
            TT = TB1.Cell(i + 3, j).Range.Text
            TT = Left(TT, Len(TT) - 2)
            A(i - 1) = TT
        Next i
        For k = 1 To UBound(A) + 1
            TB2.Cell(k + 5, j).Range.Text = A(k - 1)
        Next k
    Next j
        
    TB2.Cell(3, 2).Range.Text = Left(TB1.Cell(1, 2).Range.Text, Len(TB1.Cell(1, 2).Range.Text) - 2)
    TB2.Cell(3, 3).Range.Text = Left(TB1.Cell(1, 3).Range.Text, Len(TB1.Cell(1, 3).Range.Text) - 2)
    TB2.Cell(4, 3).Range.Text = Left(TB1.Cell(2, 3).Range.Text, Len(TB1.Cell(2, 3).Range.Text) - 2)
    TB2.Cell(4, 4).Range.Text = Left(TB1.Cell(2, 4).Range.Text, Len(TB1.Cell(2, 4).Range.Text) - 2)
    TB2.Cell(4, 5).Range.Text = Left(TB1.Cell(2, 5).Range.Text, Len(TB1.Cell(2, 5).Range.Text) - 2)
    TB2.Cell(4, 6).Range.Text = Left(TB1.Cell(2, 6).Range.Text, Len(TB1.Cell(2, 6).Range.Text) - 2)
    TB2.Cell(4, 7).Range.Text = Left(TB1.Cell(2, 7).Range.Text, Len(TB1.Cell(2, 7).Range.Text) - 2)
    TB2.Cell(5, 8).Range.Text = Left(TB1.Cell(3, 8).Range.Text, Len(TB1.Cell(3, 8).Range.Text) - 2)
    TB2.Cell(9, 1).Range.Text = Left(TB1.Cell(7, 1).Range.Text, Len(TB1.Cell(7, 1).Range.Text) - 2)
    
    Application.ScreenUpdating = True '恢复屏幕更新
End Sub

Sub A00_IEI_TB3()
    Dim A As Variant
    Dim B As Variant
    
    Application.ScreenUpdating = False '关闭屏幕更新
    On Error Resume Next
    
    Set TB1 = ActiveDocument.Tables(1)
    Set TB2 = ActiveDocument.Tables(2)
    
        ReDim A(1)
            TT = TB1.Cell(4, j).Range.Text
            TT = Left(TT, Len(TT) - 2)
            A(0) = TT
        
        For j = 2 To 8
            ReDim Preserve A(j)
            TT = TB1.Cell(4, j).Range.Text
            TT = Left(TT, Len(TT) - 2)
            A(j - 1) = TT
        Next j
        For k = 1 To UBound(A) + 1
            TB2.Cell(5, k).Range.Text = A(k - 1)
        Next k
        
    TB2.Cell(2, 2).Range.Text = Left(TB1.Cell(1, 2).Range.Text, Len(TB1.Cell(1, 2).Range.Text) - 2)
    TB2.Cell(2, 3).Range.Text = Left(TB1.Cell(1, 3).Range.Text, Len(TB1.Cell(1, 3).Range.Text) - 2)
    TB2.Cell(3, 3).Range.Text = Left(TB1.Cell(2, 3).Range.Text, Len(TB1.Cell(2, 3).Range.Text) - 2)
    TB2.Cell(3, 4).Range.Text = Left(TB1.Cell(2, 4).Range.Text, Len(TB1.Cell(2, 4).Range.Text) - 2)
    TB2.Cell(3, 5).Range.Text = Left(TB1.Cell(2, 5).Range.Text, Len(TB1.Cell(2, 5).Range.Text) - 2)
    TB2.Cell(3, 6).Range.Text = Left(TB1.Cell(2, 6).Range.Text, Len(TB1.Cell(2, 6).Range.Text) - 2)
    TB2.Cell(3, 7).Range.Text = Left(TB1.Cell(2, 7).Range.Text, Len(TB1.Cell(2, 7).Range.Text) - 2)
    TB2.Cell(4, 8).Range.Text = Left(TB1.Cell(3, 8).Range.Text, Len(TB1.Cell(3, 8).Range.Text) - 2)
    TB2.Cell(6, 1).Range.Text = Left(TB1.Cell(5, 1).Range.Text, Len(TB1.Cell(5, 1).Range.Text) - 2)
    
    Application.ScreenUpdating = True '恢复屏幕更新
End Sub

Sub A00_IEI_TB()
    
    Dim RNG As Range
    Dim A, B, C, D, FF, Z As Variant
    Z = Array("全球主要利率情况", "国际原油价格", "波罗的海干散货运指数", "世界主要经济体股票指数", "全球主要货币汇率情况")
    Dim DC1 As Document
    Dim TB1 As Table
    Dim TB2 As Table
    
    Application.ScreenUpdating = False '关闭屏幕更新
    On Error Resume Next
    
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="国际经济信息"
        If .Found = False Then
            MsgBox "不符合执行本宏命令条件！"
            Exit Sub
        End If
    End With
    
    Set DC1 = ActiveDocument
    Selection.HomeKey Unit:=wdStory
    'A00_去表头
    Application.ScreenUpdating = False '关闭屏幕更新
    Selection.HomeKey Unit:=wdStory
        
        With Selection.Find
            .ClearFormatting
            .Execute FindText:=Z(0) & Chr(13) '全球主要利率情况
            If .Found = True Then
                .Parent.Expand Unit:=wdParagraph
                Do While True
                    Selection.MoveDown Unit:=wdLine, Count:=1
                    If Selection.Information(wdWithInTable) = True Then
                        Exit Do
                    End If
                Loop
                Selection.Tables(1).Select
                Set TB1 = Selection.Tables(1)
                    R = TB1.Rows.Count
                    If R - 10 > 0 Then
                        For i = R - 10 To 1 Step -1
                            TB1.Cell(R - 10, 1).Select
                            Selection.SelectRow
                            Selection.Cut
                        Next i
                    End If
                    Set TB1 = Selection.Tables(1)
                        ReDim A(1)
                            A(0) = Left(TB1.Range.Cells(1).Range.Text, Len(TB1.Range.Cells(1).Range.Text) - 2)
                        For i = 2 To TB1.Range.Cells.Count
                            ReDim Preserve A(i)
                            A(i - 1) = Left(TB1.Range.Cells(i).Range.Text, Len(TB1.Range.Cells(i).Range.Text) - 2)
                        Next i
                        TB1.Range.Select
                        Selection.Cut
                        Selection.Paragraphs(1).Range.Select
                        Selection.MoveLeft Unit:=wdCharacter, Count:=1
                        Selection.TypeParagraph
                        Selection.TypeParagraph
                        Selection.MoveUp Unit:=wdLine, Count:=2
                        Selection.Paragraphs(1).Range.Text = "IEI_TAB1"
                End If
            
        End With
        Selection.HomeKey Unit:=wdStory
        
        With Selection.Find
            .ClearFormatting
            .Execute FindText:=Z(1) & Chr(13) '国际原油价格
            If .Found = True Then
                .Parent.Expand Unit:=wdParagraph
                Do While True
                    Selection.MoveDown Unit:=wdLine, Count:=1
                    If Selection.Information(wdWithInTable) = True Then
                        Exit Do
                    End If
                Loop
                Set TB1 = Selection.Tables(1)
                    R = TB1.Rows.Count
                    If R - 7 > 0 Then
                        For i = R - 10 To 1 Step -1
                            TB1.Cell(R - 7, 1).Select
                            Selection.SelectRow
                            Selection.Cut
                        Next i
                    End If
                    Set TB1 = Selection.Tables(1)
                        ReDim B(1)
                            B(0) = Left(TB1.Range.Cells(1).Range.Text, Len(TB1.Range.Cells(1).Range.Text) - 2)
                        For i = 2 To TB1.Range.Cells.Count
                            ReDim Preserve B(i)
                            B(i - 1) = Left(TB1.Range.Cells(i).Range.Text, Len(TB1.Range.Cells(i).Range.Text) - 2)
                        Next i
                        TB1.Range.Select
                        Selection.Cut
                        Selection.Paragraphs(1).Range.Select
                        Selection.MoveLeft Unit:=wdCharacter, Count:=1
                        Selection.TypeParagraph
                        Selection.TypeParagraph
                        Selection.MoveUp Unit:=wdLine, Count:=2
                        Selection.Paragraphs(1).Range.Text = "IEI_TAB2"
                End If
            
        End With
        Selection.HomeKey Unit:=wdStory
        
        With Selection.Find
            .ClearFormatting
            .Execute FindText:=Z(2) & Chr(13) '波罗的海干散货运指数
            If .Found = True Then
                .Parent.Expand Unit:=wdParagraph
                Do While True
                    Selection.MoveDown Unit:=wdLine, Count:=1
                    If Selection.Information(wdWithInTable) = True Then
                        Exit Do
                    End If
                Loop
                Selection.Tables(1).Select
                    Set TB1 = Selection.Tables(1)
                    R = TB1.Rows.Count
                    If R - 5 > 0 Then
                        For i = R - 10 To 1 Step -1
                            TB1.Cell(R - 5, 1).Select
                            Selection.SelectRow
                            Selection.Cut
                        Next i
                    End If
                    Set TB1 = Selection.Tables(1)
                        ReDim C(1)
                            C(0) = Left(TB1.Range.Cells(1).Range.Text, Len(TB1.Range.Cells(1).Range.Text) - 2)
                        For i = 2 To TB1.Range.Cells.Count
                            ReDim Preserve C(i)
                            C(i - 1) = Left(TB1.Range.Cells(i).Range.Text, Len(TB1.Range.Cells(i).Range.Text) - 2)
                        Next i
                        TB1.Range.Select
                        Selection.Cut
                        Selection.Paragraphs(1).Range.Select
                        Selection.MoveLeft Unit:=wdCharacter, Count:=1
                        Selection.TypeParagraph
                        Selection.TypeParagraph
                        Selection.MoveUp Unit:=wdLine, Count:=2
                        Selection.Paragraphs(1).Range.Text = "IEI_TAB3"
                End If
            
        End With
        Selection.HomeKey Unit:=wdStory
    
    With Selection.Find
        .ClearFormatting
        .Execute FindText:=Z(3) & Chr(13)  '世界主要经济体股票指数
        If .Found = True Then
            .Parent.Expand Unit:=wdParagraph
                Do While True
                    Selection.MoveDown Unit:=wdLine, Count:=1
                    If Selection.Information(wdWithInTable) = True Then
                        Exit Do
                    End If
                Loop
                Set TB1 = Selection.Tables(1)
                
                '如果有表注行，则将表注文本赋值给变量BZ，并删除表注行
                R = TB1.Rows.Count
                TB1.Cell(R, 1).Select
                Selection.SelectRow
                CN = Selection.Cells.Count
                If CN = 1 Then
                    BZ = Selection.Range.Text
                    BZ = Left(BZ, Len(BZ) - 2)
                    'MsgBox BZ
                    Selection.Cut
                End If
                
                '如果表头有两行，则合并
                TB1.Cell(1, 1).Range.Select
                Selection.MoveDown Unit:=wdLine, Count:=1
                H1 = Selection.Information(wdStartOfRangeRowNumber)
                If H1 = 3 Then
                    For i = 2 To 8
                        TB1.Cell(1, i).Range.Select
                        Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
                        Selection.Cells.Merge
                    Next i
                End If
                
                '如果第8列的行数多于10，则合并有关单元格
                RN = TB1.Rows.Count
                If RN > 10 Then
                    For i = 2 To 10
                        TT = Left(TB1.Cell(i, 8).Range.Text, Len(TB1.Cell(i, 8).Range.Text) - 2)
                        If TT <> "休市" Then
                            TB1.Cell(i, 8).Range.Select
                            Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
                            Selection.Cells.Merge
                        End If
                    Next i
                End If
            
            '将表格内容赋值给数组D
            If TB1.Rows.Count = 10 Then
                ReDim D(1)
                    D(0) = Left(TB1.Range.Cells(1).Range.Text, Len(TB1.Range.Cells(1).Range.Text) - 2)
                For i = 2 To TB1.Range.Cells.Count
                    ReDim Preserve D(i)
                    D(i - 1) = Left(TB1.Range.Cells(i).Range.Text, Len(TB1.Range.Cells(i).Range.Text) - 2)
                Next i
                TB1.Range.Select
                Selection.Cut
                Selection.Paragraphs(1).Range.Select
                Selection.MoveLeft Unit:=wdCharacter, Count:=1
                Selection.TypeParagraph
                Selection.TypeParagraph
                Selection.MoveUp Unit:=wdLine, Count:=2
                Selection.Paragraphs(1).Range.Text = "IEI_TAB4"
            End If
            End If
        
    End With
        
        Selection.HomeKey Unit:=wdStory
        With Selection.Find
            .ClearFormatting
            .Execute FindText:=Z(4) & Chr(13) '"全球主要货币汇率情况"
            If .Found = True Then
                .Parent.Expand Unit:=wdParagraph
                Do While True
                    Selection.MoveDown Unit:=wdLine, Count:=1
                    If Selection.Information(wdWithInTable) = True Then
                        Exit Do
                    End If
                Loop
                    Selection.Tables(1).Select
                    Set TB1 = Selection.Tables(1)
                    R = TB1.Rows.Count
                        ReDim FF(1)
                            FF(0) = Left(TB1.Range.Cells(1).Range.Text, Len(TB1.Range.Cells(1).Range.Text) - 2)
                        For i = 2 To TB1.Range.Cells.Count
                            ReDim Preserve FF(i)
                            FF(i - 1) = Left(TB1.Range.Cells(i).Range.Text, Len(TB1.Range.Cells(i).Range.Text) - 2)
                        Next i
                        TB1.Range.Select
                        Selection.Cut
                        Selection.Paragraphs(1).Range.Select
                        Selection.MoveLeft Unit:=wdCharacter, Count:=1
                        Selection.TypeParagraph
                        Selection.TypeParagraph
                        Selection.MoveUp Unit:=wdLine, Count:=2
                        Selection.Paragraphs(1).Range.Text = "IEI_TAB5"
                End If
            
        End With
        Selection.HomeKey Unit:=wdStory
    
    
    A00_网页格式
    A00_IEI
    Application.ScreenUpdating = False '关闭屏幕更新
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="IEI_TAB1"
        If .Found = True Then
            .Parent.Expand Unit:=wdParagraph
    Set RNG = Selection.Range
    ActiveDocument.Tables.Add Range:=RNG, NumRows:=9, NumColumns:=8
    Set TB2 = Selection.Tables(1)
    For i = 10 To UBound(A) - 1
        TB2.Range.Cells(i + 15).Range.Text = Trim(A(i))
    Next i
    
    TB2.Cell(1, 1).Range.Select
    Selection.MoveDown Unit:=wdLine, Count:=2, Extend:=wdExtend
    Selection.Cells.Merge
    
    TB2.Cell(1, 2).Range.Select
    Selection.MoveDown Unit:=wdLine, Count:=2, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.Range.Text = A(1)
    
    TB2.Cell(1, 3).Range.Select
    Selection.MoveRight Unit:=wdCharacter, Count:=6, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.Range.Text = A(2)
    
    TB2.Cell(2, 3).Range.Select
    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.Range.Text = A(3)
    
    TB2.Cell(2, 4).Range.Select
    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.Range.Text = A(4)
    
    TB2.Cell(2, 5).Range.Select
    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.Range.Text = A(5)
    
    TB2.Cell(2, 6).Range.Select
    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.Range.Text = A(6)
    
    TB2.Cell(2, 7).Range.Select
    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.Range.Text = A(7)
    
    TB2.Cell(3, 8).Range.Select
    Selection.Range.Text = A(9)
    
    表格B
    
    TB2.Select
    Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
    
    TB2.Select
    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    
    TB2.Select
    Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed)
    
    TB2.Cell(2, 8).Range.Select
    Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
    
    TB2.Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    
    End If
    End With
    
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="IEI_TAB2"
        If .Found = True Then
            .Parent.Expand Unit:=wdParagraph
    Set RNG = Selection.Range
    ActiveDocument.Tables.Add Range:=RNG, NumRows:=6, NumColumns:=8
    Selection.MoveDown Unit:=wdLine, Count:=1
    Set TB3 = Selection.Tables(1)
    For i = 10 To UBound(B) - 1
        TB3.Range.Cells(i + 15).Range.Text = Trim(B(i))
    Next i
    
    TB3.Cell(1, 1).Range.Select
    Selection.MoveDown Unit:=wdLine, Count:=2, Extend:=wdExtend
    Selection.Cells.Merge
    
    TB3.Cell(1, 2).Range.Select
    Selection.MoveDown Unit:=wdLine, Count:=2, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.Range.Text = B(1)
    
    TB3.Cell(1, 3).Range.Select
    Selection.MoveRight Unit:=wdCharacter, Count:=6, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.Range.Text = B(2)
    
    TB3.Cell(2, 3).Range.Select
    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.Range.Text = B(3)
    
    TB3.Cell(2, 4).Range.Select
    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.Range.Text = B(4)
    
    TB3.Cell(2, 5).Range.Select
    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.Range.Text = B(5)
    
    TB3.Cell(2, 6).Range.Select
    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.Range.Text = B(6)
    
    TB3.Cell(2, 7).Range.Select
    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.Range.Text = B(7)
    
    TB3.Cell(3, 8).Range.Select
    Selection.Range.Text = B(9)
    
    表格B
    
    TB3.Select
    Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
    
    TB3.Select
    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    
    TB3.Select
    Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed)
    
    TB3.Cell(2, 8).Range.Select
    Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
    
    TB3.Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    
    End If
    End With
    
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="IEI_TAB3"
        If .Found = True Then
            .Parent.Expand Unit:=wdParagraph
    Set RNG = Selection.Range
    ActiveDocument.Tables.Add Range:=RNG, NumRows:=4, NumColumns:=8
    Selection.MoveDown Unit:=wdLine, Count:=1
    Set TB3 = Selection.Tables(1)
    For i = 10 To UBound(C) - 1
        TB3.Range.Cells(i + 15).Range.Text = Trim(C(i))
    Next i
    
    TB3.Cell(1, 1).Range.Select
    Selection.MoveDown Unit:=wdLine, Count:=2, Extend:=wdExtend
    Selection.Cells.Merge
    
    TB3.Cell(1, 2).Range.Select
    Selection.MoveDown Unit:=wdLine, Count:=2, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.Range.Text = C(1)
    
    TB3.Cell(1, 3).Range.Select
    Selection.MoveRight Unit:=wdCharacter, Count:=6, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.Range.Text = C(2)
    
    TB3.Cell(2, 3).Range.Select
    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.Range.Text = C(3)
    
    TB3.Cell(2, 4).Range.Select
    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.Range.Text = C(4)
    
    TB3.Cell(2, 5).Range.Select
    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.Range.Text = C(5)
    
    TB3.Cell(2, 6).Range.Select
    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.Range.Text = C(6)
    
    TB3.Cell(2, 7).Range.Select
    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.Range.Text = C(7)
    
    TB3.Cell(3, 8).Range.Select
    Selection.Range.Text = C(9)
    
    表格B
    
    TB3.Select
    Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
    
    TB3.Select
    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    
    TB3.Select
    Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed)
    
    TB3.Cell(2, 8).Range.Select
    Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
    
    TB3.Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    
    End If
    End With
    
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="IEI_TAB4"
        If .Found = True Then
            .Parent.Expand Unit:=wdParagraph
    Set RNG = Selection.Range
    ActiveDocument.Tables.Add Range:=RNG, NumRows:=10, NumColumns:=8
    Set TB4 = Selection.Tables(1)
    For i = 1 To UBound(D) - 1
        TB4.Range.Cells(i + 1).Range.Text = Trim(D(i))
    Next i
    
    表格B
    
    TB4.Select
    Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
    
    TB4.Select
    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    
    TB4.Select
    Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed)
    
    TB4.Select
    With Selection.ParagraphFormat
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 12
    End With
    Selection.Rows.HeightRule = wdRowHeightAtLeast
    Selection.Rows.Height = CentimetersToPoints(1)
    With Selection.Borders(wdBorderHorizontal)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    
    TB4.Select
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    Selection.TypeBackspace
    End If
    End With
    
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="IEI_TAB5"
        If .Found = True Then
            .Parent.Expand Unit:=wdParagraph
    Set RNG = Selection.Range
    ActiveDocument.Tables.Add Range:=RNG, NumRows:=11, NumColumns:=8
    Set TB5 = Selection.Tables(1)
    For i = 10 To UBound(FF) - 1
        TB5.Range.Cells(i + 15).Range.Text = Trim(FF(i))
    Next i
    
    TB5.Cell(1, 1).Range.Select
    Selection.MoveDown Unit:=wdLine, Count:=2, Extend:=wdExtend
    Selection.Cells.Merge
    
    TB5.Cell(1, 2).Range.Select
    Selection.MoveDown Unit:=wdLine, Count:=2, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.Range.Text = FF(1)
    
    TB5.Cell(1, 3).Range.Select
    Selection.MoveRight Unit:=wdCharacter, Count:=6, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.Range.Text = FF(2)
    
    TB5.Cell(2, 3).Range.Select
    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.Range.Text = FF(3)
    
    TB5.Cell(2, 4).Range.Select
    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.Range.Text = FF(4)
    
    TB5.Cell(2, 5).Range.Select
    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.Range.Text = FF(5)
    
    TB5.Cell(2, 6).Range.Select
    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.Range.Text = FF(6)
    
    TB5.Cell(2, 7).Range.Select
    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.Range.Text = FF(7)
    
    TB5.Cell(3, 8).Range.Select
    Selection.Range.Text = FF(9)
    
    表格B
    
    TB5.Select
    Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
    
    TB5.Select
    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    
    TB5.Select
    Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed)
    
    TB5.Cell(2, 8).Range.Select
    Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
    
    TB5.Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    
    End If
    End With
    
    Application.ScreenUpdating = True '恢复屏幕更新

End Sub


Sub A00_表格转置()

    If Selection.Information(wdWithInTable) = True Then
        R = ActiveDocument.Tables(1).Rows.Count
        C = ActiveDocument.Tables(1).Columns.Count
    Else
        MsgBox "插入点不在表格中！"
    End If

    Set MyRange = ActiveDocument.Range(start:=0, End:=0)
    ActiveDocument.Tables.Add Range:=MyRange, NumRows:=C, NumColumns:=R

    Set mytable1 = ActiveDocument.Tables(1)
    Set mytable2 = ActiveDocument.Tables(2)
    i = 1
    j = 1
    For i = 1 To C
        For j = 1 To R
            Selection.Collapse Direction:=wdClolapseEnd
            mytable1.Cell(Row:=i, Column:=j).Select
            TT = mytable2.Cell(Row:=j, Column:=i).Range.Text
            TT = Left(TT, Len(TT) - 2)
            Selection.Range.Text = TT
        Next j
    Next i

End Sub

Sub A00_IEI_TB4_合并有关单元格()
    
    Dim TB1 As Table
    Dim RNG As Range
    
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="世界主要经济体股票指数"
        If .Found = True Then
            .Parent.Expand Unit:=wdParagraph
            Selection.MoveDown Unit:=wdLine, Count:=3
            If Selection.Information(wdWithInTable) = True Then
                Set TB1 = Selection.Tables(1)
                
                '如果有表注行，则将表注文本赋值给变量BZ，并删除表注行
                R = TB1.Rows.Count
                TB1.Cell(R, 1).Select
                Selection.SelectRow
                CN = Selection.Cells.Count
                If CN = 1 Then
                    BZ = Selection.Range.Text
                    BZ = Left(BZ, Len(BZ) - 2)
                    MsgBox BZ
                    Selection.Cut
                End If
                
                '如果表头有两行，则合并
                TB1.Cell(1, 1).Range.Select
                Selection.MoveDown Unit:=wdLine, Count:=1
                H1 = Selection.Information(wdStartOfRangeRowNumber)
                If H1 = 3 Then
                    For i = 2 To 8
                        TB1.Cell(1, i).Range.Select
                        Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
                        Selection.Cells.Merge
                    Next i
                End If
                
                '如果第8列的行数多于10，则合并有关单元格
                RN = TB1.Rows.Count
                If (RN - 1) / 2 = 9 Then
                    For i = RN To 3 Step -2
                        TB1.Cell(i, 8).Range.Select
                        Selection.MoveUp Unit:=wdLine, Count:=1, Extend:=wdExtend
                        Selection.Cells.Merge
                    Next i
                Else
                    For i = 2 To 10
                        TT = Left(TB1.Cell(i, 8).Range.Text, Len(TB1.Cell(i, 8).Range.Text) - 2)
                        If TT <> "休市" Then
                            TB1.Cell(i, 8).Range.Select
                            Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
                            Selection.Cells.Merge
                        End If
                    Next i
                End If
            End If
        End If
    End With
End Sub
Sub A00_测试()
    A00_IEI_TB
    'A00_表格数字列右对齐
    
End Sub
Sub A02_测试()  '功能：调试宏命令

End Sub

Sub A00_去表头表注()
    
    Dim RNG As Range
    Dim MyDOC As Document
    Dim TB As Table
    Dim TN As Integer
    Dim H As Integer
    Dim TT As String
    Dim DW As String
    Dim BZ As String
    
    Application.ScreenUpdating = False '关闭屏幕更新
    On Error Resume Next
    Set MyDOC = ActiveDocument
    TN = MyDOC.Tables.Count
    
    Selection.HomeKey Unit:=wdStory
    For i = 1 To TN
        Set TB = MyDOC.Tables(i)
        '去表头
        TB.Cell(1, 1).Select
        Selection.SelectRow
        N = Selection.Cells.Count
        If N > 1 Then
            H = 0
        Else
            TB.Cell(2, 1).Select
            Selection.SelectRow
            N = Selection.Cells.Count
            If N = 1 Then
                H = 2
            Else
                H = 1
            End If
        End If
            
        If H = 2 Then
            TB.Cell(2, 1).Select
            DW = Left(Selection.Range.Text, Len(Selection.Range.Text) - 2)
            TT = Left(TB.Cell(1, 1).Range.Text, Len(TB.Cell(1, 1).Range.Text) - 2)
            TT = TT & Chr(13) & DW
            Selection.SelectRow
            Selection.Cut
            TB.Cell(1, 1).Select
            Selection.SelectRow
            Selection.Cut
            TB.Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
            Selection.TypeParagraph
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            Selection.Paragraphs(1).Range.Text = TT
        End If
        
        If H = 1 Then
            TB.Cell(1, 1).Select
            Selection.SelectRow
            TT = Left(Selection.Range.Text, Len(Selection.Range.Text) - 2)
            Selection.Cut
            TB.Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
            Selection.TypeParagraph
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            Selection.Paragraphs(1).Range.Text = TT
        End If
        
        If H = 0 Then
            TB.Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
        End If
        
        '去表注
        TB.Cell(TB.Rows.Count, 1).Select
        Selection.SelectRow
        N = Selection.Cells.Count
        If N = 1 Then
            BZ = Left(Selection.Range.Text, Len(Selection.Range.Text) - 2)
            Selection.Cut
            TB.Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
            Selection.TypeParagraph
            'Selection.TypeParagraph
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            Selection.Paragraphs(1).Range.Text = BZ
        End If
    Next i
    Selection.HomeKey Unit:=wdStory
    Application.ScreenUpdating = True '恢复屏幕更新
End Sub

Sub A00_加表头表注()
    Dim RNG As Range
    Dim MyDOC As Document
    Dim TB As Table
    Dim TN As Integer
    Dim H As Integer
    Dim TT As String
    Dim DW As String
    Dim BZ As String
    
    'Application.ScreenUpdating = False '关闭屏幕更新
    On Error Resume Next
    Set MyDOC = ActiveDocument
    Set TB = MyDOC.Tables(1)
    
    TB.Select
    
    '查找离表头最近的非空行
    Do While True
        Selection.MoveUp Unit:=wdLine, Count:=1
        Selection.Paragraphs(1).Range.Select
        TT = Selection.Range.Text
        If Len(TT) < 2 Then
            Selection.Delete
        Else
            Exit Do
        End If
    Loop
    
    '如果内容包含“单位”,则查找下一个非空行，选择变为带单位的两行表头
    If InStr(1, TT, "单位", 1) > 0 Then
        Do While True
            Selection.MoveUp Unit:=wdLine, Count:=1
            Selection.Paragraphs(1).Range.Select
            TT = Selection.Range.Text
            If Len(TT) < 2 Then
                Selection.Delete
            Else
                Exit Do
            End If
        Loop
        Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
        B01_选定文字变为表头
    Else
         TXT = Left(TT, Len(TT) - 1)
         'MsgBox Right(TXT, 1)
         If Right(TXT, 1) = "。" Or Right(TXT, 1) = "：" Or Len(TXT) > 30 Then
            TB.Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
        Else
            Selection.Paragraphs(1).Range.Select
            B01_选定文字变为表头
        End If
    End If
    
    TB.Select
    
    '查找离表底最近的非空行
    Do While True
        TB.Select
        Selection.MoveDown Unit:=wdLine, Count:=1
        Selection.Paragraphs(1).Range.Select
        TT = Selection.Range.Text
        If Len(TT) < 2 Then
            Selection.Delete
        Else
            Exit Do
        End If
    Loop
    
    If InStr(1, TT, "注：", 1) > 0 Or InStr(1, TT, "资料来源：", 1) > 0 Then
        Selection.Paragraphs(1).Range.Select
        Selection.Paragraphs(1).Range.Text = Trim(TT)
        Selection.Paragraphs(1).Range.Select
        B01_选定文本变为表注
    Else
        Selection.Paragraphs(1).Range.Select
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        Selection.TypeParagraph
    End If
    
    End Sub
Sub A00_批量加表头表注()
    Dim RNG As Range
    Dim MyDOC As Document
    Dim TB As Table
    Dim TN As Integer
    Dim H As Integer
    Dim TT As String
    Dim DW As String
    Dim BZ As String
    
    Application.ScreenUpdating = False '关闭屏幕更新
    On Error Resume Next
    Set MyDOC = ActiveDocument
    'Set TB = MyDOC.Tables(1)
    TN = MyDOC.Tables.Count
    If TN > 0 Then
    For i = 1 To TN
    Set TB = MyDOC.Tables(i)
    
    TB.Select
    
    '查找离表头最近的非空行
    Do While True
        Selection.MoveUp Unit:=wdLine, Count:=1
        Selection.Paragraphs(1).Range.Select
        TT = Selection.Range.Text
        If Len(TT) < 2 Then
            Selection.Delete
        Else
            Exit Do
        End If
    Loop
    
    '如果内容包含“单位”,则查找下一个非空行，选择变为带单位的两行表头
    If InStr(1, TT, "单位", 1) > 0 Then
        Do While True
            Selection.MoveUp Unit:=wdLine, Count:=1
            Selection.Paragraphs(1).Range.Select
            TT = Selection.Range.Text
            If Len(TT) < 2 Then
                Selection.Delete
            Else
                Exit Do
            End If
        Loop
        Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
        B01_选定文字变为表头
    Else
         TXT = Left(TT, Len(TT) - 1)
         'MsgBox Right(TXT, 1)
         If Right(TXT, 1) = "。" Or Right(TXT, 1) = "：" Or Len(TXT) > 30 Then
            TB.Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
        Else
            Selection.Paragraphs(1).Range.Select
            B01_选定文字变为表头
        End If
    End If
    
    TB.Select
    
    '查找离表底最近的非空行
    Do While True
        TB.Select
        Selection.MoveDown Unit:=wdLine, Count:=1
        Selection.Paragraphs(1).Range.Select
        TT = Selection.Range.Text
        If Len(TT) < 2 Then
            Selection.Delete
        Else
            Exit Do
        End If
    Loop
    
    If InStr(1, TT, "注：", 1) > 0 Or InStr(1, TT, "资料来源：", 1) > 0 Then
        Selection.Paragraphs(1).Range.Select
        Selection.Paragraphs(1).Range.Text = Trim(TT)
        Selection.Paragraphs(1).Range.Select
        B01_选定文本变为表注
    Else
        Selection.Paragraphs(1).Range.Select
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        Selection.TypeParagraph
    End If
    
    Next i
    End If
    
    Selection.HomeKey Unit:=wdStory
    Application.ScreenUpdating = True '恢复屏幕更新
End Sub
Sub A00_删除空行()
    If ActiveDocument.Tables.Count > 0 Then
        MsgBox "文档中有表格，不能执行本宏命令！"
        Exit Sub
    End If
    TT = ActiveDocument.Content.Text
    Do While True
        If InStr(1, TT, Chr(13) & Chr(13), 1) = 0 Then
            Exit Do
        Else
            P1 = InStr(1, TT, Chr(13) & Chr(13), 1)
            TT = Left(TT, P1 - 1) & Right(TT, Len(TT) - P1)
            P1 = InStr(1, TT, Chr(13) & Chr(13), 1)
        End If
    Loop
    ActiveDocument.Content.Text = TT
     
End Sub
Sub A00_删除选定区域空行()
    
    Dim RNG As Range
    Set RNG = Selection.Range
    If Len(Trim(RNG.Text)) < 2 Then
        MsgBox "亲，你没有选定文本，不能执行本宏命令！"
        Exit Sub
    Else
        If RNG.Tables.Count > 0 Then
            MsgBox "文档中有表格，不能执行本宏命令！"
            Exit Sub
        End If
    End If
        
    TT = RNG.Text
    Do While True
        If InStr(1, TT, Chr(13) & Chr(13), 1) = 0 Then
            Exit Do
        Else
            P1 = InStr(1, TT, Chr(13) & Chr(13), 1)
            TT = Left(TT, P1 - 1) & Right(TT, Len(TT) - P1)
            P1 = InStr(1, TT, Chr(13) & Chr(13), 1)
        End If
    Loop
    RNG.Text = TT
     
End Sub




Sub A00_测试N()
    
    A00_删除空行
    'A00_去表头表注
    'MsgBox "已去表头表注"
    'A00_批量加表头表注
    'MsgBox "已加表头表注"
End Sub


Sub A00_删除网页空格()
    
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " " & Chr(13)
        .Replacement.Text = Chr(13)
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.HomeKey Unit:=wdStory
End Sub

Sub A00_表标题加粗居中()
    
    Dim N As Integer
    Dim TT As String
    
    Application.ScreenUpdating = False '关闭屏幕更新
    On Error Resume Next
    Selection.HomeKey Unit:=wdStory
    N = ActiveDocument.Tables.Count
    If N > 0 Then
        TT = "表" & N
        With Selection.Find
            .ClearFormatting
            .Execute FindText:=TT
            If .Found = True Then
                Selection.HomeKey Unit:=wdStory
                For i = 1 To N
                    With Selection.Find
                        .ClearFormatting
                        .Execute FindText:="表" & i
                        If .Found = True Then
                        .Parent.Expand Unit:=wdParagraph
                        Selection.Paragraphs(1).Range.Text = Trim(Selection.Paragraphs(1).Range.Text)
                        Selection.Paragraphs(1).Range.Select
                        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
                        Selection.Font.Bold = True
                        Selection.MoveRight Unit:=wdCharacter, Count:=1
                        End If
                    End With
                Next i
            End If
        End With
    End If
    Selection.HomeKey Unit:=wdStory
    Application.ScreenUpdating = True '恢复屏幕更新

End Sub

Sub A00_表标题_批量()
    Dim RNG As Range
    Dim MyDOC As Document
    Dim TB As Table
    Dim TN As Integer
    Dim TT As String
    
    Application.ScreenUpdating = False '关闭屏幕更新
    On Error Resume Next
    Set MyDOC = ActiveDocument
    TN = MyDOC.Tables.Count
    If TN > 0 Then
        For i = 1 To TN
            Set TB = MyDOC.Tables(i)
            TB.Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.MoveUp Unit:=wdLine, Count:=1
            Selection.Paragraphs(1).Range.Select
            Set RNG1 = Selection.Range
            TT = Trim(RNG1.Text)
            TT = Left(TT, Len(TT) - 1)
            If Right(TT, 1) <> "。" Or Right(TT, 1) <> "：" And Len(TT) < 40 Then
                RNG1.Text = TT & Chr(13)
                RNG1.Select
                Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
                Selection.Font.Bold = True
            End If
            TB.Select
            Selection.MoveDown Unit:=wdLine, Count:=1
        Next i
    End If
    
    Selection.HomeKey Unit:=wdStory
    Application.ScreenUpdating = True '恢复屏幕更新
End Sub

Sub A00_单位居右()
    
    TT = Selection.Range.Text
    TT = Trim(TT)
    TT = Left(TT, Len(TT) - 1)
    Set MyRange = Selection.Range
    If InStr(1, TT, "单位", 1) > 0 And Selection.Paragraphs.Count = 1 Then
        Selection.MoveDown Unit:=wdLine, Count:=1
            If Selection.Information(wdWithInTable) = True Then
                Set myrng = Selection.Tables(1).Range
                Set MyTab = Selection.Tables(1)
                MyTab.Cell(1, 1).Select
                Selection.SelectRow
                Selection.InsertRowsAbove 1
                MyTab.Cell(1, 1).Select
                Selection.SelectRow
                Selection.Cells.Merge
                Selection.Cells(1).Range.Text = TT
                Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
                Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
                Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
                Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone
                Selection.Shading.Texture = wdTextureNone
                Selection.Shading.ForegroundPatternColor = wdColorAutomatic
                Selection.Shading.BackgroundPatternColor = wdColorAutomatic
                Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
                Selection.Borders(wdBorderBottom).LineWidth = wdLineWidth150pt
                Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
                Selection.SelectRow
                Selection.Font.Size = 10.5
                Selection.Range.Font.Bold = False
                 Selection.Rows.HeightRule = wdRowHeightAtLeast
                Selection.Rows.Height = CentimetersToPoints(0.5)
                Selection.MoveLeft Unit:=wdCharacter, Count:=1
            End If
    Else
        Exit Sub
    End If
    
    MyRange.Select
    Selection.Delete
    Selection.TypeBackspace
    
End Sub

Sub A00_单位居右_批量()
    Dim RNG1 As Range
    Dim RNG2 As Range
    Dim MyDOC As Document
    Dim TB As Table
    Dim TN As Integer
    Dim TT As String
    
    Application.ScreenUpdating = False '关闭屏幕更新
    On Error Resume Next
    Set MyDOC = ActiveDocument
    TN = MyDOC.Tables.Count
    If TN > 0 Then
        For i = 1 To TN
            Set TB = MyDOC.Tables(i)
            TB.Select
            Selection.MoveUp Unit:=wdLine, Count:=1
            Selection.Paragraphs(1).Range.Select
            Set RNG1 = Selection.Range
            Selection.MoveUp Unit:=wdLine, Count:=1
            Selection.Paragraphs(1).Range.Select
            Set RNG2 = Selection.Range
            TT = RNG2.Text
            '如果内容包含“单位”,则执行“A00_单位居右”宏命令
            If InStr(1, TT, "单位", 1) > 0 Then
                RNG1.Select
                Selection.Delete
                RNG2.Select
                A00_单位居右
            End If
            TB.Select
            Selection.MoveDown Unit:=wdLine, Count:=1
        Next i
    End If
    
    Selection.HomeKey Unit:=wdStory
    Application.ScreenUpdating = True '恢复屏幕更新
End Sub
Sub A00_单位_批量_信息公开()
    Dim RNG1 As Range
    Dim RNG2 As Range
    Dim MyDOC As Document
    Dim TB As Table
    Dim TN As Integer
    Dim TT As String
    
    Application.ScreenUpdating = False '关闭屏幕更新
    On Error Resume Next
    Set MyDOC = ActiveDocument
    TN = MyDOC.Tables.Count
    If TN > 0 Then
        For i = 1 To TN
            Set TB = MyDOC.Tables(i)
            TB.Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.MoveUp Unit:=wdLine, Count:=1
            Set RNG2 = Selection.Paragraphs(1).Range
            TT = Trim(RNG2.Text)
            '如果内容包含“单位”,则执行“A00_单位居右”宏命令
            If InStr(1, TT, "单位", 1) > 0 Then
                TT1 = ""
                For j = 1 To 50 - Len(TT)
                    TT1 = TT1 & Chr(-24159)
                Next j
                TT = TT1 & TT
                RNG2.Text = TT
                RNG2.Select
                Selection.MoveRight Unit:=wdCharacter, Count:=1
                Selection.Delete
            End If
            TB.Select
            Selection.MoveDown Unit:=wdLine, Count:=1
        Next i
    End If
    
    Selection.HomeKey Unit:=wdStory
    Application.ScreenUpdating = True '恢复屏幕更新
End Sub

Sub A00_表注_批量()
    Dim RNG1 As Range
    Dim RNG2 As Range
    Dim MyDOC As Document
    Dim TB As Table
    Dim TN As Integer
    Dim TT As String
    
    Application.ScreenUpdating = False '关闭屏幕更新
    On Error Resume Next
    Set MyDOC = ActiveDocument
    TN = MyDOC.Tables.Count
    If TN > 0 Then
        For i = 1 To TN
            Set TB = MyDOC.Tables(i)
            TB.Select
            Selection.MoveDown Unit:=wdLine, Count:=1
            Selection.Paragraphs(1).Range.Select
            Set RNG1 = Selection.Range
            Selection.MoveDown Unit:=wdLine, Count:=1
            Selection.Paragraphs(1).Range.Select
            Set RNG2 = Selection.Range
            TT = RNG2.Text
            If InStr(1, TT, "注：", 1) > 0 Or InStr(1, TT, "资料来源：", 1) > 0 Then
                RNG1.Select
                Selection.Delete
                RNG2.Select
                B01_选定文本变为表注
            End If
            TB.Select
            Selection.MoveDown Unit:=wdLine, Count:=1
        Next i
    End If
    
    Selection.HomeKey Unit:=wdStory
    Application.ScreenUpdating = True '恢复屏幕更新
End Sub

Sub A00_设定文档保存目录()

    '设定文档保存目录
    Dim FD1 As String
    Dim FD2 As String
    
    FD1 = "C:\Users\zlzx-dhg\Desktop\00 OK_DOC"
    FD2 = "D:\00 OK_DOC"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FolderExists(FD1) = True Then
        ChangeFileOpenDirectory FD1
    Else
        If fs.FolderExists(FD2) = Fase Then
            Set A = fs.CreateFolder(FD2)
            ChangeFileOpenDirectory FD2
        Else
            ChangeFileOpenDirectory FD2
        End If
    End If
End Sub

Sub A00_失信企业()

    '声明变量
    Dim DOC1, DOC2 As Document
    Dim FD1, NM1, TT As String
    Dim AA() As Variant
    Dim RNG As Range
    Dim TB, TB2 As Table
    Dim i, RN, CN As Integer
    
    FD1 = "F:\03 SXQY"
    NM1 = "统计上严重失信企业信息公示："
    
    '设定文档保存目录
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FolderExists(FD1) = True Then
        ChangeFileOpenDirectory FD1
    Else
        Set A = fs.CreateFolder(FD1)
        ChangeFileOpenDirectory FD1
    End If
    
    Application.ScreenUpdating = False '关闭屏幕更新
    On Error Resume Next
    
    Documents("失信企业公示信息.Docx").Activate
    Set DOC1 = ActiveDocument
    Documents("失信企业信息公示模板.Docx").Activate
    Set DOC2 = ActiveDocument
    'MsgBox ActiveDocument.Paragraphs(1).Range.Text
    DOC1.Activate
    'MsgBox ActiveDocument.Paragraphs(1).Range.Text
    Set TB = DOC1.Tables(1)
    RN = TB.Rows.Count '行数
    CN = TB.Columns.Count '列数
    'MsgBox "行数：" & RN & "列数：" & CN
    For i = 2 To RN
    TT = Left(TB.Cell(i, 1).Range.Text, Len(TB.Cell(i, 1).Range.Text) - 2)
    ReDim AA(1)
    AA(0) = TT
    For j = 2 To CN
        TT = Left(TB.Cell(i, j).Range.Text, Len(TB.Cell(i, j).Range.Text) - 2)
        ReDim Preserve AA(j)
        AA(j - 1) = TT
    Next j
    DOC2.Activate
    Selection.WholeStory
    Selection.Copy
    Documents.Add
    页面设置
    Selection.Paste
    Set TB2 = ActiveDocument.Tables(1)
    TB2.Select
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.Delete
    Selection.HomeKey Unit:=wdStory
    For H = 0 To UBound(AA)
        TB2.Cell(H + 1, 2).Range.Text = AA(H)
    Next H
    ActiveDocument.Paragraphs(1).Range.Text = NM1 & AA(6) & Chr(13)
    ActiveDocument.Paragraphs(1).Range.Select
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.HomeKey Unit:=wdStory
    ActiveDocument.SaveAs FileName:="A" & IIf(Len(i - 1) = 1, "0" & i - 1, i - 1) & " " & NM1 & AA(6) & ".doc", FileFormat:=wdFormatDocument
    ActiveDocument.Close
    DOC1.Activate
    Next i
    Selection.HomeKey Unit:=wdStory
    Application.ScreenUpdating = True '恢复屏幕更新
    
End Sub

Sub A00_OA另存1()
    
    Dim TT As String
    Dim P1 As Integer
    Dim C As Variant
    C = Array(Chr(13), Chr(32), Chr(-24159), Chr(11))
    
    '如果没有选择文本，提示用户选择
    If Len(Selection.Range.Text) = 0 Then
        MsgBox "【注意】没有选择作为另存文档文件名的文本" & Chr(13) & _
           "请选择或键入文件名后再选中， " & Chr(13) & _
           "然后再执行本宏，谢谢！"
        Exit Sub
    End If
        
    TT = Selection.Range.Text
    
    '去掉文本中指定字符集
    For i = 0 To UBound(C)
        P1 = InStr(1, TT, C(i), 1)
        If P1 > 0 Then
            Do While P1 > 0
                TT = Left(TT, P1 - 1) & Right(TT, Len(TT) - P1)
                P1 = InStr(1, TT, C(i), 1)
            Loop
        End If
    Next i
        
    ChangeFileOpenDirectory "C:\Users\zlzx-dhg\Desktop\03 OA"
    ActiveDocument.SaveAs FileName:=TT & ".doc", FileFormat:=wdFormatDocument

End Sub

Sub A00_OA另存()

    Dim TT, FD1, FD2 As String
    Dim P1 As Integer
    Dim C As Variant
    
    '设定文档保存目录
    FD1 = "F:\03 OA_DOC"
    FD2 = "D:\03 OA_DOC"
    C = Array(Chr(13), Chr(32), Chr(-24159), Chr(11))

    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FolderExists(FD1) = True Then
        ChangeFileOpenDirectory FD1
    Else
        If fs.FolderExists(FD2) = False Then
            Set A = fs.CreateFolder(FD2)
            ChangeFileOpenDirectory FD2
        Else
            ChangeFileOpenDirectory FD2
        End If
    End If

    Application.ScreenUpdating = False '关闭屏幕更新
    On Error Resume Next '忽略错误
    
    '如果没有选择文本，提示用户选择
    If Len(Selection.Range.Text) = 0 Then
        If Left(ActiveDocument.Paragraphs(1).Range.Text, Len(ActiveDocument.Paragraphs(1).Range.Text) - 1) = "工作情况交流" Then NM1 = "工作情况交流"
        If Left(ActiveDocument.Paragraphs(4).Range.Text, Len(ActiveDocument.Paragraphs(4).Range.Text) - 1) = "每日调查" Then NM1 = "每日调查"
        NM2 = ActiveDocument.Shapes(1).TextFrame.TextRange.Text
        NM2 = Left(NM2, Len(NM2) - 1)
        TT = NM1 & NM2
        If NM1 = "" Then
            MsgBox "【注意】没有选择作为另存文档文件名的文本" & Chr(13) & _
               "请选择或键入文件名后再选中， " & Chr(13) & _
               "然后再执行本宏，谢谢！"
            Exit Sub
        End If
    Else
        TT = Selection.Range.Text
    End If
    
    '去掉文本中指定字符集
    For i = 0 To UBound(C)
        P1 = InStr(1, TT, C(i), 1)
        If P1 > 0 Then
            Do While P1 > 0
                TT = Left(TT, P1 - 1) & Right(TT, Len(TT) - P1)
                P1 = InStr(1, TT, C(i), 1)
            Loop
        End If
    Next i
    
    ActiveDocument.SaveAs FileName:=TT & ".doc", FileFormat:=wdFormatDocument
    'ChangeFileOpenDirectory "C:\Users\zlzx-dhg\Desktop"
    
End Sub

